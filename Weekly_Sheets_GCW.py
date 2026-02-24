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
import streamlit.components.v1 as components

# --- PAGE CONFIG ---
st.set_page_config(page_title="Milkbasket Campaign Auto-Processor", layout="wide")

AWS_BASE_URL = "https://design-figma.s3.ap-south-1.amazonaws.com/"

# --- CACHED DATA FETCHING ---
@st.cache_data(show_spinner=False)
def fetch_google_sheet_data(url):
    """Fetches the Master PID database from your Google Sheet."""
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

def clean_excel_name(name):
    """V13 Standard Abbreviation Logic for Excel Tab Names (31 char limit)."""
    name = str(name).replace('*', 'x')
    name = re.sub(r'[\[\]\:\/\\\?]', '', name)
    if len(name) > 28:
        subs = {"Background": "BG", "Essentials": "Ess", "Category": "Cat", "Discount": "Disc"}
        for full, abbr in subs.items():
            name = re.sub(rf'(?i)\b{full}\b', abbr, name)
    name = name[:31].strip()
    return re.sub(r'[\s\|\&\-]+$', '', name).strip()

def clean_tab_name(campaign, asset):
    c, a = str(campaign).strip(), str(asset).strip()
    if c.lower() in ["", "nan", "unknown campaign"]: c = ""
    if a.lower() in ["", "nan", "unknown asset"]: a = ""
    name = f"{c} | {a}" if c and a else (a if a else (c if c else "Unnamed Asset"))
    return clean_excel_name(name)

# --- SMART MESSY DATA AUTO-CLEANER ---
def auto_clean_messy_tab(df, hub, img_map):
    cleaned_rows = []
    current_campaign, current_asset = "Unknown Campaign", "Unknown Asset"
    promo_keywords = ["upto", "% off", "buy", "%", "free", "flat ", "discount", "cashback", "bogo", "save ₹", "save rs"]
    headers = [str(c).strip().lower() for c in df.columns]

    for _, row in df.iterrows():
        pid1, pid2, grid_detail, call_out, unmapped_vals = "", "", "", "", []
        found_first_pid = False

        for i, col_name in enumerate(df.columns):
            val = row.iloc[i]
            if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == "nan": continue
            val_str = str(val).strip()
            if val_str.endswith('.0'): val_str = val_str[:-2]
            col_header = headers[i]
            
            if "remark" in col_header: continue 

            if re.match(r'^\d{3,8}$', val_str) and val_str in img_map:
                if "2" in col_header and ("pid" in col_header or "mb" in col_header): pid2 = val_str
                elif "1" in col_header and ("pid" in col_header or "mb" in col_header): pid1 = val_str
                else:
                    if not pid1: pid1 = val_str
                    elif not pid2: pid2 = val_str
                found_first_pid = True
            elif not found_first_pid:
                if "campaign" in col_header: current_campaign = val_str
                elif "asset" in col_header: current_asset = val_str
                elif "grid" in col_header or "title" in col_header: grid_detail = val_str
                elif "call" in col_header or "offer" in col_header: call_out = val_str
                else: unmapped_vals.append(val_str)

        if not pid1 and not pid2: continue
        for val in unmapped_vals:
            v_low = val.lower()
            if not call_out and any(k in v_low for k in promo_keywords): call_out = val
            elif not grid_detail and not re.match(r'(?i)^(campaign|asset|grid|call out|remarks|pid|name|mb)', val): grid_detail = val

        if current_asset.lower() in ["atc", "atc background"]: continue
        cleaned_rows.append({
            "tab": clean_tab_name(current_campaign, current_asset),
            "Hub": hub, "Title": grid_detail, "PID1": pid1, "PID2": pid2,
            "Img1": img_map.get(pid1, ""), "Img2": img_map.get(pid2, ""),
            "AmzId1": make_amz_link(pid1), "AmzId2": make_amz_link(pid2),
            "Callout": call_out, "Framename": f"{hub}-{grid_detail}" if grid_detail else f"{hub}"
        })
    return cleaned_rows

def excel_export(tabs, all_pids_tab):
    output = BytesIO()
    wb = Workbook()
    yellow_fill, bold_font = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid"), Font(bold=True)
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
                    cell.fill, cell.font = yellow_fill, bold_font
    ws2 = wb.create_sheet("All_PIDs")
    pid_df = pd.DataFrame(all_pids_tab)
    for r_idx, row in enumerate(dataframe_to_rows(pid_df, index=False, header=True), 1): ws2.append(row)
    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    wb.save(output)
    output.seek(0)
    return output

# --- IMAGE PROCESSING ---
def process_single_image(item, use_rembg=False):
    pid, img_url = item["PID"], item["Img Link"]
    try:
        r = requests.get(img_url, timeout=10)
        if r.status_code == 200:
            img = Image.open(io.BytesIO(r.content)).convert("RGBA")
            img_byte_arr = io.BytesIO()
            if use_rembg and not (img.mode in ("RGBA", "LA") and img.getchannel("A").getextrema()[0] < 255):
                remove(img).save(img_byte_arr, format='PNG')
            else: img.save(img_byte_arr, format='PNG')
            return {"filename": f"{pid}.png", "data": img_byte_arr.getvalue(), "success": True}
    except: pass
    return {"filename": f"{pid}.png", "success": False}

def batch_download_images(image_list, use_rembg=False):
    zip_buffer = io.BytesIO()
    progress_bar, status_text = st.progress(0), st.empty()
    success_count, total = 0, len(image_list)
    workers = 1 if use_rembg else 10
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(process_single_image, item, use_rembg): item for item in image_list}
            for i, future in enumerate(concurrent.futures.as_completed(futures)):
                res = future.result()
                if res["success"]:
                    zipf.writestr(res["filename"], res["data"])
                    success_count += 1
                progress_bar.progress((i + 1) / total)
                status_text.text(f"Processed {i+1}/{total} images...")
    zip_buffer.seek(0)
    status_text.success(f"✅ Packaged {success_count}/{total} images.")
    return zip_buffer

# ==========================================
# 🎨 STREAMLIT UI LAYOUT
# ==========================================

with st.sidebar:
    st.header("⚙️ Configuration")
    gsheet_url = st.text_input("Google Sheet Link:", value="https://docs.google.com/spreadsheets/d/1yQJfr9UhSpfXBdUDksj_zDdhevLZ9OXD3pktJOlRtwg/edit")
    uploaded_file = st.file_uploader("Upload Campaign Excel", type=["xlsx"])
    
    st.markdown("---")
    st.markdown("<div style='text-align: center; color: #888; font-size: 14px;'>Buddy is hard at work...</div>", unsafe_allow_html=True)
    
    components.html(
        """
        <div style="display: flex; justify-content: center; align-items: center; height: 100%;">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2257.7 2065.34" style="width: 180px; height: 180px;">
              <style>
                #Cat_Body path { fill: #f0a552; } 
                #Tail path { fill: #f0a552; animation: wag 1.5s ease-in-out infinite alternate; transform-origin: 500px 1900px; }
                #Paw_1 { animation: type 0.2s infinite alternate; }
                #Paw_2 { animation: type 0.2s 0.1s infinite alternate; }
                #Cat_Pupil_1, #Cat_Pupil_2 { animation: blink 4s infinite; transform-origin: center; }
                .cls-1 { fill: #3a3a3a; } 
                .cls-5 { fill: #57606f; } 
                .cls-2, .cls-3 { stroke: #fff; stroke-width: 15px; fill: #f0a552; } 
                @keyframes type { from { transform: translateY(0); } to { transform: translateY(-40px); } }
                @keyframes wag { from { transform: rotate(-5deg); } to { transform: rotate(10deg); } }
                @keyframes blink { 0%, 90%, 100% { transform: scaleY(1); } 95% { transform: scaleY(0); } }
              </style>
              <g id="Tail">
                <path d="M576.97,1910.42c0,44.78-33.52,83.67-77.56,82.7-13.71-.3-27.95-1.66-41.19-4.41-70.86-14.68-136.18-45.08-193.49-95.18-73.3-64.07-121.32-144.08-132.42-242.03-8.6-75.84,1.08-149.7,42.44-216.31,20.89-33.65,49.82-56.47,90.14-63.47,46.06-8,94.46,18.84,110.95,62.62,16.83,44.67,1.47,92.99-42.38,115.15-26.74,13.51-31,34.15-32.35,58.26-2.3,41.03,11.06,78.16,32.95,111.99,35.55,54.93,86.4,88.3,147.81,106.1,3.54,1.03,7.12,2.01,10.73,2.93,7.67,1.96,11.29,1.82,15.95,2.32,47.36,5.07,68.42,37.66,68.42,79.33Z"/>
              </g>
              <g id="Cat_Body">
                  <path d="M1043.75,1993.15c-23.62-.96-47.15-1.83-70.92-.64-52.8,2.66-105.66,4.14-158.53,4.78-102.41,1.26-205.13,4.91-307.22-3.28-47.33-3.8-42.98-169.68-41.97-184.22,7.03-100.6,34.78-196.31,71.26-289.78,12.83-32.88,26.74-65.31,43-96.62,6.3-12.12,3.51-17.06-8.82-21.88-57.94-22.64-113.12-50.9-164.3-86.34-107.95-74.77-186.73-173.39-236.75-294.65-6.97-16.91-15.55-23.21-33.99-22.1-30.4,1.83-60.97.9-91.47.66-25.29-.2-42.93-16.76-43.25-39.79-.33-23.68,17.54-40.52,43.85-40.73,26.22-.2,52.45-.27,78.65.35,10.37.25,14.91-2.08,12.09-13.43-3.46-13.89,1.88-33.06-7.78-41.17-10.34-8.69-28.77-2.19-43.66-2.55-15.5-.37-31.08.87-46.5-.3-21.55-1.64-36.78-18.31-37.43-38.93-.66-21.13,12.97-38,35.9-39.46,26.61-1.7,53.45-.83,80.14.14,12.68.46,16.41-4.46,17.45-16.35,7.6-86.82,32.7-168.66,72.95-245.72,8.81-16.86,9.9-31.41,4.18-49.79-40.58-130.43-55.34-264.42-50.5-400.6,1.98-55.65,39.83-82.13,92.83-65.65,108.17,33.64,206.06,87.94,299.07,151.61,20.28,13.88,39.98,28.63,59.65,43.36,6.63,4.96,12.34,5.48,20.51,2.97,143.25-43.96,286.99-45.77,430.77-2.85,14.48,4.32,24.49,3.9,37.41-5.99,95.55-73.12,198.29-133.99,311.14-176.8,13.46-5.11,27.17-9.7,41.01-13.66,50.25-14.37,89.04,13.69,90.66,66.22,1.95,63.12,1.17,126.17-5.97,189.07-8.39,73.9-22.03,146.68-45.08,217.42-4.7,14.42-3.82,25.88,3.08,39.17,41.79,80.47,68.31,165.69,76.62,256.2.8,8.71,4.58,10.74,12.09,10.63,22.47-.33,44.94-.42,67.41-.58,9.64-.07,19.31-.06,28.33,4.02,15.27,6.9,25.64,24.71,23.68,40.32-2.2,17.51-17.8,33.5-36.92,34.63-23.99,1.41-48.15,1.56-72.16.52-15.78-.68-24.92,2.69-22.2,20.55.16,1.04-.03,2.14-.11,3.2q-2.5,33.68,30.94,33.55c18.73-.07,37.46-.4,56.18-.11,23.62.37,43.93,19.03,44.2,40.05.27,21.3-19.22,39.83-43.7,40.18-34.77.5-69.56.28-104.34-.02-8.35-.07-13.14,1.71-16.63,10.63-71.17,181.83-201.47,305.04-379.13,381.02-7.37,3.15-14.6,6.63-21.99,9.73q-20.15,8.44-10.5,28.78c46.79,98.34,83.3,200.18,102.94,307.56,8.88,48.52,14.29,97.5,5.1,146.65-12.26,65.6-67.12,111.04-133.56,111.52-19.41.14-38.6-.7-57.73-1.49ZM948.93,1148.16c-.11-14.21-5.8-23.16-13.7-30.75-13.49-12.98-27.07-25.89-41.11-38.27-6.41-5.66-10.22-11.23-9.4-20.39.87-9.73-1.7-19.41-8.12-27.06-9.38-11.17-21.31-15.49-35.84-11.14-14.31,4.28-24.53,13.15-23.11,28.52,1.77,19.21-7.8,30.61-21.28,41.5-11.2,9.05-21.65,19.16-31.65,29.55-13.74,14.27-14.35,31.65-2.54,44.64,12.51,13.75,29.79,14.47,45.47.69,12.43-10.92,24.27-22.61,35.53-34.73,6.41-6.9,10.24-6.18,16.63.09,13.33,13.09,25.16,27.78,40.42,38.86,10.67,7.74,21.86,7.77,33.2,1.92,10.49-5.41,15.35-14.41,15.5-23.42Z"/>
                  <path style="fill:#fff; opacity:0.3;" d="M1212.88,777.96c-35.03,0-63.42,46.56-63.42,104s28.39,104,63.42,104,63.42-46.56,63.42-104-28.39-104-63.42-104Z"/>
                  <path style="fill:#fff; opacity:0.3;" d="M534.91,777.96c-35.03,0-63.42,46.56-63.42,104s28.39,104,63.42,104,63.42-46.56,63.42-104-28.39-104-63.42-104Z"/>
              </g>
              <g id="Laptop_Layer_2"><path class="cls-5" d="M2038.82,2065.34H981.86c-11,0-19.15-10.24-16.67-20.96l169.3-731.95c9.26-40.03,44.91-68.37,86-68.37h948.9c56.78,0,98.79,52.84,86,108.16l-153.32,662.84c-6.81,29.44-33.03,50.29-63.25,50.29Z"/></g>
              <g id="Laptop_Layer"><path class="cls-1" d="M2032.45,2065.34H992.35l174.15-752.91c9.26-40.03,44.91-68.37,86-68.37h916.9c56.78,0,98.79,52.84,86,108.16l-151.88,656.63c-7.65,33.08-37.11,56.5-71.06,56.5Z"/></g>
              <g id="Cat_Pupil_2"><circle cx="1216.85" cy="922.1" r="26" style="fill:#000;"/></g>
              <g id="Cat_Pupil_1"><circle cx="544.16" cy="925.55" r="26" style="fill:#000;"/></g>
              <g id="Paw_1"><path class="cls-2" d="M821.61,2027c-9.16-2.72-32.05-20.83-32.05-20.83,0,0-19.41,20.83-41.48,20.83-25.59,0-45.01-18.18-45.01-18.18,0,0-24.17,15.83-33.38,18.18-22.75,5.79-42.69-19.21-42.69-42.69v-64.53c0-37.63,30.79-68.41,68.41-68.41h100.46c37.63,0,68.41,30.79,68.41,68.41v64.53c0,23.48-20.18,49.36-42.69,42.69Z"/></g>
              <g id="Paw_2"><path class="cls-3" d="M1089.2,2023.78c-9.16-2.72-32.05-20.83-32.05-20.83,0,0-19.41,20.83-41.48,20.83-25.59,0-45.01-18.18-45.01-18.18,0,0-24.17,15.83-33.38,18.18-22.75,5.79-42.69-19.21-42.69-42.69v-64.53c0-37.63,30.79-68.41,68.41-68.41h100.46c37.63,0,68.41,30.79,68.41,68.41v64.53c0,23.48-20.18,49.36-42.69,42.69Z"/></g>
            </svg>
        </div>
        """, height=220
    )

# --- MAIN APP ---
st.title("🚀 Campaign Asset Auto-Processor")
st.markdown("Automates data cleanup, Excel generation, and bulk image processing.")

if uploaded_file:
    with st.spinner("Buddy is processing your data..."):
        product_df = fetch_google_sheet_data(gsheet_url)
        if product_df is not None:
            img_map, _ = make_img_map(product_df)
            xls = pd.ExcelFile(uploaded_file)
            all_rows = []
            for tab in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=tab)
                all_rows.extend(auto_clean_messy_tab(df, tab, img_map))
            
            if all_rows:
                tabs = {}
                for r in all_rows:
                    tabs.setdefault(r["tab"], []).append({k: v for k, v in r.items() if k != "tab"})
                
                pid_set = {r[c] for r in all_rows for c in ["PID1", "PID2"] if r[c]}
                all_pids_tab = [{"PID": p, "Img Link": img_map.get(p, ""), "AmzID": make_amz_link(p)} for p in sorted(pid_set)]
                
                t1, t2 = st.tabs(["📊 Data & Export", "🖼️ Images"])
                with t1:
                    st.download_button("📥 Master Excel", excel_export(tabs, all_pids_tab), "Campaign_Output.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.dataframe(pd.DataFrame(all_pids_tab))
                with t2:
                    if st.button("⚡ Start Image Processing"):
                        zip_file = batch_download_images([r for r in all_pids_tab if r["Img Link"]])
                        st.download_button("💾 Save ZIP", zip_file, "Images.zip", "application/zip")
else:
    st.info("👈 Upload your Campaign Excel file to begin.")
