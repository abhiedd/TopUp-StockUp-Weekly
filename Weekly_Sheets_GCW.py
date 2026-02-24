import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from PIL import Image
from rembg import remove
import requests
import zipfile
import io
import re
import concurrent.futures

# --- PAGE CONFIG ---
st.set_page_config(page_title="Milkbasket Campaign Auto-Processor", layout="wide")

AWS_BASE_URL = "https://design-figma.s3.ap-south-1.amazonaws.com/"

# --- CACHED DATA FETCHING ---
@st.cache_data(show_spinner=False)
def fetch_google_sheet_data(url):
    """Extracts the ID from a Google Sheet URL and fetches it as a CSV."""
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if not match: return None
    sheet_id = match.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    try:
        df = pd.read_csv(export_url)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Error reading Google Sheet: {e}")
        return None

# --- HELPER FUNCTIONS ---
def fix_pid(pid):
    if pd.isna(pid) or str(pid).strip().lower() == "nan": return ""
    try: return str(int(float(pid))).strip()
    except: return str(pid).strip()

def make_amz_link(pid):
    if not pid: return ""
    return f"{AWS_BASE_URL}{pid}.png"

def make_img_map(product_df):
    """Creates a dictionary map of Valid PIDs to Image URLs."""
    img_map, src_map = {}, {}
    for _, row in product_df.iterrows():
        pid = fix_pid(row.get('PID', row.get('MB_id', ''))) 
        img_src = str(row.get('image_src', '')).strip()
        if pid and img_src and pid.lower() != 'nan' and img_src.lower() != 'nan':
            img_map[pid] = f"https://file.milkbasket.com/products/{img_src}"
            src_map[pid] = img_src
    return img_map, src_map

def clean_tab_name(campaign, asset):
    return f"{campaign.strip()} | {asset.strip()}"

def clean_sheet_name(name):
    # Excel sheets have a 31 character limit and restrict certain characters
    return re.sub(r'[\[\]\*:/\\?]', '', str(name)).strip()[:31]

# --- SMART MESSY DATA AUTO-CLEANER ---
def auto_clean_messy_tab(df, hub, img_map):
    """Parses messy sheets using Database Validation and Header Context."""
    cleaned_rows = []
    current_campaign = "Unknown Campaign"
    current_asset = "Unknown Asset"
    asset_keywords = ['atc', 'hero', 'cards', 'medium', 'bg', 'banner', '3*2', 'store', 'slider']
    
    # Grab all headers and make them lowercase for context checking
    headers = [str(c).strip().lower() for c in df.columns]
    
    for _, row in df.iterrows():
        pid1, pid2 = "", ""
        pre_pid = []
        found_first_pid = False
        
        # Iterate through the row using the exact column index (i)
        for i, col_name in enumerate(df.columns):
            val = row.iloc[i]
            
            # Skip empty cells
            if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan":
                continue
                
            val_str = str(val).strip()
            if val_str.endswith('.0'): val_str = val_str[:-2] # Clean float parsing
            
            # 1. Database Validation & Header Checking
            # It must look like a number AND exist in our Master PID Database
            if re.match(r'^\d{3,8}$', val_str) and val_str in img_map:
                col_header = headers[i]
                
                # Context Check: Look at the header to place it in the correct slot
                if "2" in col_header and ("pid" in col_header or "mb" in col_header):
                    pid2 = val_str
                elif "1" in col_header and ("pid" in col_header or "mb" in col_header):
                    pid1 = val_str
                else:
                    # Fallback: if header doesn't specify 1 or 2, fill them sequentially
                    if not pid1:
                        pid1 = val_str
                    elif not pid2:
                        pid2 = val_str
                        
                found_first_pid = True
                
            elif not found_first_pid:
                # Keep collecting metadata until we hit the first valid PID
                pre_pid.append(val_str)
                
        # If the row had no valid data at all, skip it entirely
        if not pre_pid and not pid1 and not pid2:
            continue
            
        # 2. Extract Call Out and clean Metadata
        call_out = ""
        new_pre_pid = []
        for val in pre_pid:
            v_low = val.lower()
            if "upto" in v_low or "% off" in v_low or "buy" in v_low or "%" in v_low or "free" in v_low:
                call_out = val
            else:
                new_pre_pid.append(val)
                
        # Safeguard: Remove actual header words if they accidentally got caught in the sweep
        new_pre_pid = [x for x in new_pre_pid if not re.match(r'(?i)^(campaign|asset|grid|call out|remarks|pid|name|mb)', x)]
        
        # 3. Inherit or Update Grid State
        grid_detail = ""
        if len(new_pre_pid) >= 3:
            current_campaign = new_pre_pid[0]
            current_asset = new_pre_pid[1]
            grid_detail = new_pre_pid[2]
        elif len(new_pre_pid) == 2:
            if any(k in new_pre_pid[1].lower() for k in asset_keywords):
                current_campaign = new_pre_pid[0]
                current_asset = new_pre_pid[1]
            else:
                current_campaign = new_pre_pid[0]
                grid_detail = new_pre_pid[1]
        elif len(new_pre_pid) == 1:
            grid_detail = new_pre_pid[0]
            
        # 4. Append Valid Product Rows
        if pid1 or pid2:
            if current_asset.lower() in ["atc", "atc background"]:
                continue
                
            cleaned_rows.append({
                "tab": clean_tab_name(current_campaign, current_asset),
                "Hub": hub,
                "Title": grid_detail,
                "PID1": pid1,
                "PID2": pid2,
                "Img1": img_map.get(pid1, "") if pid1 else "",
                "Img2": img_map.get(pid2, "") if pid2 else "",
                "AmzId1": make_amz_link(pid1),
                "AmzId2": make_amz_link(pid2),
                "Callout": call_out,
                "Framename": f"{hub}-{grid_detail}" if grid_detail else f"{hub}"
            })
            
    return cleaned_rows

def excel_export(tabs, all_pids_tab):
    output = BytesIO()
    wb = Workbook()
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    bold_font = Font(bold=True)

    for tname, rows in tabs.items():
        ws = wb.create_sheet(title=clean_sheet_name(tname))
        
        # Dynamically append Callout column if any row in the tab has one
        has_callout = any(r.get("Callout") for r in rows)
        cols = ["Hub", "Title", "PID1", "PID2", "Img1", "Img2", "AmzId1", "AmzId2"]
        if has_callout: cols.append("Callout")
        cols.append("Framename")
        
        df = pd.DataFrame(rows)[cols]
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            ws.append(row)
            if r_idx == 1:
                for c_idx in range(1, len(row) + 1):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.fill = yellow_fill
                    cell.font = bold_font
                    
    ws2 = wb.create_sheet("All_PIDs")
    pid_df = pd.DataFrame(all_pids_tab)
    for r_idx, row in enumerate(dataframe_to_rows(pid_df, index=False, header=True), 1):
        ws2.append(row)
        
    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    wb.save(output)
    output.seek(0)
    return output

# --- PARALLEL IMAGE PROCESSING ---
def has_transparency(img):
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        alpha = img.getchannel("A") if "A" in img.getbands() else None
        if alpha and alpha.getextrema()[0] < 255: return True
    return False

def process_single_image(item, use_rembg=False):
    pid = item["PID"]
    img_url = item["Img Link"]
    filename = f"{pid}.png"
    try:
        r = requests.get(img_url, timeout=10)
        if r.status_code == 200:
            img = Image.open(io.BytesIO(r.content)).convert("RGBA")
            img_byte_arr = io.BytesIO()
            if use_rembg and not has_transparency(img):
                result = remove(img)
                result.save(img_byte_arr, format='PNG')
            else:
                img.save(img_byte_arr, format='PNG')
            return {"filename": filename, "data": img_byte_arr.getvalue(), "success": True}
    except Exception:
        pass
    return {"filename": filename, "success": False}

def batch_download_images(image_list, use_rembg=False):
    zip_buffer = io.BytesIO()
    progress_bar = st.progress(0)
    status_text = st.empty()
    success_count = 0
    total = len(image_list)
    
    # SMART SCALING for Free Tier environments (e.g. Streamlit Cloud)
    workers = 1 if use_rembg else 10 
    
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(process_single_image, item, use_rembg): item for item in image_list}
            completed = 0
            for future in concurrent.futures.as_completed(futures):
                result = future.result()
                if result["success"]:
                    zipf.writestr(result["filename"], result["data"])
                    success_count += 1
                completed += 1
                progress_bar.progress(completed / total)
                status_text.text(f"Processed {completed}/{total} images... (Using {workers} threads)")
                
    zip_buffer.seek(0)
    status_text.success(f"✅ Packaging complete! Successfully processed {success_count} out of {total} images.")
    return zip_buffer


# ==========================================
# 🎨 STREAMLIT UI LAYOUT
# ==========================================
with st.sidebar:
    st.header("⚙️ Configuration")
    st.markdown("**1. Product Database (Google Sheet)**")
    gsheet_url = st.text_input(
        "Paste Google Sheet Link:", 
        value="https://docs.google.com/spreadsheets/d/1yQJfr9UhSpfXBdUDksj_zDdhevLZ9OXD3pktJOlRtwg/edit",
        help="Make sure the sheet is accessible (Anyone with the link can view)."
    )
    st.markdown("---")
    st.markdown("**2. Campaign Data (Excel)**")
    uploaded_file = st.file_uploader("Upload Messy Campaign Excel", type=["xlsx"])

st.title("🚀 Campaign Asset Auto-Processor")
st.markdown("Automates data cleanup, Excel generation, AWS link creation, and bulk image processing.")

if uploaded_file and gsheet_url:
    with st.spinner("Fetching Product Database from Google Sheets..."):
        product_df = fetch_google_sheet_data(gsheet_url)
        
    if product_df is not None:
        img_map, src_map = make_img_map(product_df)
        all_rows = []
        xls = pd.ExcelFile(uploaded_file)
        
        for tab in xls.sheet_names:
            # We read with header=0 so we capture the true header names for context mapping
            df = pd.read_excel(xls, sheet_name=tab, header=0)
            rows = auto_clean_messy_tab(df, tab, img_map)
            all_rows.extend(rows)

        if all_rows:
            tabs = {}
            for r in all_rows:
                tabs.setdefault(r["tab"], []).append({k: v for k, v in r.items() if k != "tab"})
            
            pid_set = {r[col] for r in all_rows for col in ["PID1", "PID2"] if r[col]}
            all_pids = sorted(pid_set, key=lambda x: (0, int(x)) if str(x).isdigit() else (1, str(x)))
            all_pids_tab = [{"PID": pid, "Img Link": img_map.get(pid, ""), "AmzID": make_amz_link(pid)} for pid in all_pids]
            
            tab1, tab2 = st.tabs(["📊 Excel & Data", "🖼️ Image Download Tools"])
            with tab1:
                st.subheader("Data Preview & Export")
                selected_tab = st.selectbox("Select Hub Tab to Preview:", sorted(tabs.keys()))
                st.dataframe(pd.DataFrame(tabs[selected_tab]), use_container_width=True)
                output_excel = excel_export(tabs, all_pids_tab)
                st.download_button(
                    "📥 Download Master Excel",
                    data=output_excel,
                    file_name="Campaign_Master_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            with tab2:
                st.subheader("Bulk Image Processing")
                st.info(f"Found {len(all_pids_tab)} unique products across all campaigns.")
                all_img_rows = [r for r in all_pids_tab if r.get("Img Link")]
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("### Standard Download")
                    if st.button("⚡ Start Standard Download", use_container_width=True):
                        zip_data = batch_download_images(all_img_rows, use_rembg=False)
                        st.download_button("📥 Save Standard ZIP", data=zip_data, file_name="Standard_Images.zip", mime="application/zip", type="primary")

                with col2:
                    st.markdown("### AI Background Removal")
                    if st.button("🤖 Start Rembg Download", use_container_width=True):
                        zip_data = batch_download_images(all_img_rows, use_rembg=True)
                        st.download_button("📥 Save Rembg ZIP", data=zip_data, file_name="Rembg_Images.zip", mime="application/zip", type="primary")
else:
    st.info("👈 Please paste your Google Sheet Link and upload your Campaign Excel file in the sidebar to begin.")
