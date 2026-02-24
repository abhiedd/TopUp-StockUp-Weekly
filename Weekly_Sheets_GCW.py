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
    img_map, src_map = {}, {}
    for _, row in product_df.iterrows():
        pid = fix_pid(row.get('PID', row.get('MB_id', ''))) 
        img_src = str(row.get('image_src', '')).strip()
        if pid and img_src and pid.lower() != 'nan' and img_src.lower() != 'nan':
            img_map[pid] = f"https://file.milkbasket.com/products/{img_src}"
            src_map[pid] = img_src
    return img_map, src_map

def clean_excel_name(name):
    name = str(name)
    name = name.replace('*', 'x')
    name = re.sub(r'[\[\]\:\/\\\?]', '', name)
    
    # V13 Standard Abbreviations
    if len(name) > 28:
        subs = {"Background": "BG", "Essentials": "Ess", "Category": "Cat", "Discount": "Disc"}
        for full, abbr in subs.items():
            name = re.sub(rf'(?i)\b{full}\b', abbr, name)
            
    name = name[:31].strip()
    name = re.sub(r'[\s\|\&\-]+$', '', name).strip()
    return name

def clean_tab_name(campaign, asset):
    c = str(campaign).strip()
    a = str(asset).strip()
    if c.lower() in ["", "nan", "unknown campaign"]: c = ""
    if a.lower() in ["", "nan", "unknown asset"]: a = ""
    
    if c and a: name = f"{c} | {a}"
    elif a:     name = a
    elif c:     name = c
    else:       name = "Unnamed Asset"
        
    return clean_excel_name(name)

# --- SMART MESSY DATA AUTO-CLEANER ---
def auto_clean_messy_tab(df, hub, img_map):
    cleaned_rows = []
    current_campaign = "Unknown Campaign"
    current_asset = "Unknown Asset"
    
    promo_keywords = ["upto", "% off", "buy", "%", "free", "flat ", "discount", "cashback", "bogo", "save ₹", "save rs"]
    headers = [str(c).strip().lower() for c in df.columns]
    
    for _, row in df.iterrows():
        pid1, pid2 = "", ""
        grid_detail = ""
        call_out = ""
        unmapped_vals = []
        found_first_pid = False
        
        for i, col_name in enumerate(df.columns):
            val = row.iloc[i]
            if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan":
                continue
                
            val_str = str(val).strip()
            if val_str.endswith('.0'): val_str = val_str[:-2] 
            col_header = headers[i]
            
            if "remark" in col_header:
                continue
                
            if re.match(r'^\d{3,8}$', val_str) and val_str in img_map:
                if "2" in col_header and ("pid" in col_header or "mb" in col_header):
                    pid2 = val_str
                elif "1" in col_header and ("pid" in col_header or "mb" in col_header):
                    pid1 = val_str
                else:
                    if not pid1: pid1 = val_str
                    elif not pid2: pid2 = val_str
                found_first_pid = True
                
            elif not found_first_pid:
                if "campaign" in col_header:
                    current_campaign = val_str
                elif "asset" in col_header:
                    current_asset = val_str
                elif "grid" in col_header or "title" in col_header or "item" in col_header:
                    grid_detail = val_str
                elif "call" in col_header or "offer" in col_header:
                    call_out = val_str
                else:
                    unmapped_vals.append(val_str)
                    
        if not pid1 and not pid2:
            continue
            
        for val in unmapped_vals:
            v_low = val.lower()
            if not call_out and any(k in v_low for k in promo_keywords):
                call_out = val
            elif not grid_detail and not re.match(r'(?i)^(campaign|asset|grid|call out|remarks|pid|name|mb)', val):
                grid_detail = val
                
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
        ws = wb.create_sheet(title=clean_excel_name(tname))
        
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
    
    # --- 🐾 APP COMPANION (GINGER NERD CAT - PURE SVG) ---
    st.markdown("---")
    st.markdown("<div style='text-align: center; color: #888; font-size: 14px; margin-bottom: 10px;'>Workspace Buddy</div>", unsafe_allow_html=True)
    st.markdown(
        """
        <div style="display: flex; justify-content: center; margin-bottom: 20px;">
            <svg width="160" height="160" viewBox="0 0 200 180" xmlns="http://www.w3.org/2000/svg">
                <defs>
                    <radialGradient id="glow" cx="50%" cy="50%" r="50%">
                        <stop offset="0%" stop-color="#FFF3B0" stop-opacity="0.4"/>
                        <stop offset="100%" stop-color="#FFF3B0" stop-opacity="0"/>
                    </radialGradient>
                </defs>
                <style>
                    .paw-l { animation: type-l 0.2s infinite alternate; }
                    .paw-r { animation: type-r 0.25s infinite alternate-reverse; }
                    .tail { animation: wag 2s ease-in-out infinite alternate; transform-origin: 145px 130px; }
                    #eye-l { animation: blink 4s infinite; transform-origin: 90px 71px; }
                    #eye-r { animation: blink 4s infinite; transform-origin: 138px 71px; }
                    @keyframes type-l { 0% { transform: translateY(0px); } 100% { transform: translateY(-4px); } }
                    @keyframes type-r { 0% { transform: translateY(0px); } 100% { transform: translateY(-4px); } }
                    @keyframes wag { 0% { transform: rotate(-5deg); } 100% { transform: rotate(15deg); } }
                    @keyframes blink { 0%, 95%, 100% { transform: scaleY(1); } 97% { transform: scaleY(0.1); } }
                </style>

                <g stroke="#34495e" stroke-width="2" stroke-linejoin="round">
                    <polygon points="25,40 85,55 70,130 10,115" fill="#57606f" />
                    <polygon points="28,44 82,58 68,126 14,112" fill="#ced6e0" stroke="none" />
                </g>

                <polygon points="28,44 140,40 160,160 14,112" fill="url(#glow)" style="mix-blend-mode: overlay;" />

                <g stroke="#593618" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                    
                    <g class="tail">
                        <path d="M 145 130 Q 185 140 185 90 Q 185 60 160 60 Q 145 60 150 80" fill="none" stroke="#593618" stroke-width="24"/>
                        <path d="M 145 130 Q 185 140 185 90 Q 185 60 160 60 Q 145 60 150 80" fill="none" stroke="#FFAC4A" stroke-width="18"/>
                        <path d="M 145 130 Q 185 140 185 90 Q 185 60 160 60 Q 145 60 150 80" fill="none" stroke="#D97A29" stroke-width="18" stroke-dasharray="10 15"/>
                    </g>
                    
                    <path d="M 85 160 C 85 100 165 100 165 160 Z" fill="#FFAC4A" />
                    <path d="M 110 160 C 110 125 155 125 160 160 Z" fill="#FFEDC2" stroke="none" />
                    <path d="M 145 110 Q 155 115 163 112" stroke="#D97A29" stroke-width="3" fill="none"/>
                    <path d="M 150 125 Q 158 130 164 125" stroke="#D97A29" stroke-width="3" fill="none"/>

                    <polygon points="90,45 80,15 110,38" fill="#FFAC4A" />
                    <polygon points="138,45 148,15 118,38" fill="#FFAC4A" />
                    <polygon points="90,42 83,23 105,37" fill="#F19A9B" stroke="none" />
                    <polygon points="138,42 145,23 123,37" fill="#F19A9B" stroke="none" />

                    <ellipse cx="114" cy="75" rx="42" ry="34" fill="#FFAC4A" />
                    
                    <path d="M 114 42 L 114 52 M 104 46 L 109 56 M 124 46 L 119 56" stroke="#D97A29" stroke-width="3" fill="none"/>
                    
                    <ellipse cx="114" cy="88" rx="18" ry="12" fill="#FFEDC2" stroke="none" />
                    <path d="M 114 83 L 110 87 A 4 4 0 0 0 118 87 Z" fill="#E88383" stroke-linejoin="round"/>
                    <path d="M 114 87 Q 109 92 105 89 M 114 87 Q 119 92 123 89" stroke-width="1.5" fill="none"/>
                    
                    <ellipse cx="82" cy="85" rx="7" ry="4" fill="#F19A9B" stroke="none" opacity="0.8" />
                    <ellipse cx="146" cy="85" rx="7" ry="4" fill="#F19A9B" stroke="none" opacity="0.8" />
                    
                    <path d="M 78 80 L 60 75 M 78 86 L 58 86 M 78 92 L 62 97" stroke-width="1.5" fill="none"/>
                    <path d="M 150 80 L 168 75 M 150 86 L 170 86 M 150 92 L 166 97" stroke-width="1.5" fill="none"/>
                </g>

                <circle id="eye-l" cx="90" cy="71" r="4.5" fill="#2C3A47" />
                <circle id="eye-r" cx="138" cy="71" r="4.5" fill="#2C3A47" />
                
                <g stroke="#2C3A47" stroke-width="4.5" fill="rgba(255,255,255,0.15)">
                    <rect x="70" y="55" width="40" height="30" rx="8" />
                    <rect x="118" y="55" width="40" height="30" rx="8" />
                    <line x1="110" y1="70" x2="118" y2="70" />
                    <line x1="70" y1="65" x2="55" y2="60" stroke-linecap="round"/>
                    <line x1="158" y1="65" x2="173" y2="60" stroke-linecap="round"/>
                </g>

                <g stroke="#34495e" stroke-width="2" stroke-linejoin="round">
                    <polygon points="10,115 70,130 130,155 10,145" fill="#a4b0be" />
                    <polygon points="10,145 130,155 130,162 10,152" fill="#747d8c" />
                    <path d="M 22 124 L 70 134 M 18 132 L 80 143 M 15 140 L 95 152" stroke="#57606f" stroke-width="1.5" opacity="0.6"/>
                </g>

                <g stroke="#593618" stroke-width="2.5">
                    <ellipse class="paw-l" cx="70" cy="138" rx="16" ry="11" fill="#FFAC4A" />
                    <ellipse class="paw-r" cx="105" cy="146" rx="16" ry="11" fill="#FFAC4A" />
                </g>
            </svg>
        </div>
        """, unsafe_allow_html=True
    )

st.title("🚀 Campaign Asset Auto-Processor")
st.markdown("Automates data cleanup, Excel generation, AWS link creation, and bulk image processing.")

# --- FANCY UI: HEADER REFERENCE GUIDE ---
with st.expander("📋 Quick Guide: Best Practices for Excel Headers", expanded=False):
    st.info("""
    **To guarantee 100% accurate data extraction, use these header names in Row 1 of your Excel file:**
    
    * 🏷️ **`Campaign`** - Extracts the main campaign name.
    * 🎨 **`Asset`** - Groups items into tabs (e.g., *Banner, Medium Cards*).
    * 📝 **`Grid`** or **`Title`** - Maps the product category or grid detail.
    * 🔑 **`PID 1`** and **`PID 2`** - Explicitly tells the engine which product is which.
    * 📢 **`Callout`** or **`Offer`** - Extracts promotional text (e.g., *Upto 50% Off*).
    * 🚫 **`Remarks`** - Add this header to *any* column you want the engine to completely ignore.
    
    *(Note: Even if your sheet is messy or missing these exact headers, the engine will still try to intelligently auto-map the data!)*
    """)

if uploaded_file and gsheet_url:
    with st.spinner("Fetching Product Database from Google Sheets..."):
        product_df = fetch_google_sheet_data(gsheet_url)
        
    if product_df is not None:
        img_map, src_map = make_img_map(product_df)
        all_rows = []
        xls = pd.ExcelFile(uploaded_file)
        
        for tab in xls.sheet_names:
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
