import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from PIL import Image
from rembg import remove, new_session
import requests
import zipfile
import io
import re
import os
import concurrent.futures
import streamlit.components.v1 as components

# --- PAGE CONFIG ---
st.set_page_config(page_title="Milkbasket Campaign Auto-Processor", layout="wide")

AWS_BASE_URL = "https://design-figma.s3.ap-south-1.amazonaws.com/"

# --- CACHED RESOURCES & DATA ---
@st.cache_resource(show_spinner=False)
def get_bria_session():
    """Loads the Bria RMbg 1.4 model once into memory for superior e-commerce cutouts."""
    return new_session("briarmbg1.4")

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
def has_transparency(img):
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        alpha = img.getchannel("A") if "A" in img.getbands() else None
        if alpha and alpha.getextrema()[0] < 255: return True
    return False

def process_single_image(item, use_rembg=False, rembg_session=None):
    pid = item["PID"]
    img_url = item["Img Link"]
    filename = f"{pid}.png"
    try:
        r = requests.get(img_url, timeout=10)
        if r.status_code == 200:
            img = Image.open(io.BytesIO(r.content)).convert("RGBA")
            img_byte_arr = io.BytesIO()
            
            if use_rembg:
                # 1. Remove background if the image doesn't already have one
                if not has_transparency(img):
                    img = remove(img, session=rembg_session)
                
                # 2. Auto-Crop Excess Transparent Pixels
                # Gets the bounding box by looking exclusively at the Alpha channel
                alpha = img.getchannel("A")
                bbox = alpha.getbbox()
                if bbox:
                    img = img.crop(bbox)

            img.save(img_byte_arr, format='PNG')
            return {"filename": filename, "data": img_byte_arr.getvalue(), "success": True}
    except Exception:
        pass
    return {"filename": filename, "success": False}

def batch_download_images(image_list, use_rembg=False, rembg_session=None):
    zip_buffer = io.BytesIO()
    progress_bar = st.progress(0)
    status_text = st.empty()
    success_count = 0
    total = len(image_list)
    
    workers = 1 if use_rembg else 10 
    
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(process_single_image, item, use_rembg, rembg_session): item for item in image_list}
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
    
    # --- 🐾 APP COMPANION (USER CUSTOM SVG + CSS ANIMATION) ---
    st.markdown("---")
    st.markdown("<div style='text-align: center; color: #888; font-size: 14px; margin-bottom: 5px;'>Workspace Buddy</div>", unsafe_allow_html=True)
    
    components.html(
        """
        <div style="display: flex; justify-content: center; align-items: center; width: 100%; height: 100%; overflow: hidden;">
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 2313.69 2165.45" style="max-width: 240px; height: auto;">
              <defs>
                <style>
                  .cls-1 { fill: #3a3a3a; }
                  .cls-2 { fill: #fff; }
                  .cls-3 { fill: #666; }
                  .cls-4 { fill: #515151; }
                  .cls-5 { stroke-width: 29px; }
                  .cls-5, .cls-6 { fill: #3a3a3a; stroke: #fff; }
                  .cls-5, .cls-6, .cls-7 { stroke-miterlimit: 10; }
                  .cls-6 { stroke-width: 21px; }
                  .cls-8 { fill: #2d2d2d; }
                  .cls-7 { fill: #eeedef; stroke: #231f20; }
                  
                  /* Dynamic CSS Animations */
                  #Tail { 
                      animation: wag 1.6s ease-in-out infinite alternate; 
                      transform-origin: 80% 90%; 
                      transform-box: fill-box; 
                  }
                  #Paw_1 { animation: type 0.2s infinite alternate; }
                  #Paw_2 { animation: type 0.2s 0.1s infinite alternate; }
                  #Cat_Pupil_1, #Cat_Pupil_2 { 
                      animation: blink 4s infinite; 
                      transform-origin: center; 
                      transform-box: fill-box; 
                  }
                  
                  @keyframes wag { 
                      from { transform: rotate(-5deg); } 
                      to { transform: rotate(10deg); } 
                  }
                  @keyframes type { 
                      from { transform: translateY(0); } 
                      to { transform: translateY(-80px); } 
                  }
                  @keyframes blink { 
                      0%, 94%, 100% { transform: scaleY(1); } 
                      97% { transform: scaleY(0.1); } 
                  }
                </style>
              </defs>
              <g id="Tail">
                <path class="cls-5" d="M682.44,1953.7c0,44.78-33.52,83.67-77.56,82.7-13.71-.3-27.95-1.66-41.19-4.41-70.86-14.68-136.18-45.08-193.49-95.18-73.3-64.07-121.32-144.08-132.42-242.03-8.6-75.84,1.08-149.7,42.44-216.31,20.89-33.65,49.82-56.47,90.14-63.47,46.06-8,94.46,18.84,110.95,62.62,16.83,44.67,1.47,92.99-42.38,115.15-26.74,13.51-31,34.15-32.35,58.26-2.3,41.03,11.06,78.16,32.95,111.99,35.55,54.93,86.4,88.3,147.81,106.1,3.54,1.03,7.12,2.01,10.73,2.93,7.67,1.96,11.29,1.82,15.95,2.32,47.36,5.07,68.42,37.66,68.42,79.33Z"/>
              </g>
              <g id="Cat_Body" data-name="Cat Body">
                <g>
                  <path class="cls-1" d="M693.56,2053.2c-59.07,0-107.35-1.67-151.92-5.24-18.35-1.47-32.5-14.88-42.04-39.86-6.12-16.01-10.67-37.62-13.53-64.24-4.93-45.91-3.42-92.12-2.91-99.32,6.35-90.94,29.37-184.7,72.44-295.07,14.62-37.47,28.17-68.04,42.53-95.92-60.93-23.96-116.81-53.42-166.13-87.58-109.92-76.14-191.73-177.94-243.14-302.59-4.57-11.09-6.99-11.09-13.43-11.09-.88,0-1.83.03-2.83.09-12.92.78-26.87,1.14-43.89,1.14-10.21,0-20.53-.13-30.52-.25-6.1-.08-12.19-.15-18.28-.2-34.95-.27-60.65-24.47-61.1-57.54-.22-15.66,5.41-30.15,15.85-40.81,11.32-11.57,27.61-18.02,45.86-18.16,12.64-.1,22.85-.14,32.16-.14,14.53,0,27.62.11,39.77.35-.38-4.08-.43-8.06-.48-11.76-.05-3.46-.11-8.41-.62-10.99-.68-.12-1.8-.25-3.51-.25-3.25,0-7.16.43-11.3.89-4.95.55-10.07,1.12-15.35,1.12-.56,0-1.11,0-1.66-.02-1.88-.05-3.87-.07-6.06-.07-3.9,0-7.77.07-11.87.14-4.07.07-8.28.15-12.5.15-6.51,0-11.91-.18-17-.57-30.37-2.31-53.1-25.99-54.05-56.32-.99-31.59,21.19-55.98,52.74-57.99,10.54-.67,21.7-1,34.12-1,15.38,0,30.29.48,46.7,1.08,7.62-86.88,32.83-171.79,74.91-252.38,6.56-12.56,7.33-22.02,2.95-36.11-20.08-64.54-34.54-132.27-42.96-201.31-8.03-65.82-10.84-134.88-8.33-205.28,1.9-53.37,33.79-87.85,81.26-87.85,11.09,0,22.83,1.9,34.9,5.66,98.75,30.71,195.32,79.63,303.89,153.94,20.62,14.11,40.78,29.21,60.27,43.81.72.54,1.15.77,1.33.85.35-.01,1.26-.1,3.12-.68,75.29-23.1,151.41-34.82,226.24-34.82s143.98,10.74,214.96,31.93c4.25,1.27,7.4,1.86,9.92,1.86s5.78-.59,11.41-4.9c103.88-79.49,207.15-138.16,315.69-179.34,14.17-5.37,28.45-10.13,42.45-14.13,10.29-2.94,20.47-4.44,30.26-4.44,48.23,0,81.73,35.12,83.34,87.4,2.29,74.1.36,135-6.07,191.66-9.41,82.87-24.4,155.15-45.85,220.97-3.22,9.9-2.73,16.29,1.95,25.3,42.94,82.68,69.17,169.12,78.01,257.05,15.85-.2,31.83-.31,47.34-.41l14.66-.1c1.13,0,2.26-.02,3.39-.02,9.69,0,20.99.44,32.48,5.63,22.41,10.13,37.08,35.48,34.12,58.97-1.6,12.71-7.81,24.84-17.48,34.17-10.01,9.65-22.88,15.39-36.24,16.18-13.84.81-27.99,1.22-42.04,1.22-10.58,0-21.33-.23-31.96-.69-.98-.04-1.9-.06-2.74-.06-.31,0-.6,0-.87,0,.44,2.98.2,5.47.06,6.81l-.03.35c-.38,5.09-.76,10.28-.88,14.11,3.06.12,6.56.12,8.35.12s3.51,0,5.43-.01c5.49-.02,10.97-.06,16.45-.11,7.6-.06,15.46-.12,23.22-.12,6.17,0,11.69.04,16.88.12,33.18.52,61.53,26.99,61.92,57.82.4,31.71-27.16,57.91-61.44,58.4-14.01.2-28.77.3-45.14.3-20.25,0-40.62-.15-59.61-.32h-.03c-35.72,90.95-87.33,170.22-153.42,235.63-63.79,63.13-142.88,115.14-235.09,154.58-3.23,1.38-6.42,2.82-9.8,4.35-3.98,1.8-8.1,3.66-12.31,5.43,0,0-1.89.79-2.76,1.16.51,1.09,1.04,2.21,1.56,3.29,51.71,108.68,85.86,210.75,104.39,312.05,8.25,45.09,15.3,98.55,5.09,153.2-13.79,73.77-75.94,125.67-151.13,126.21-1.47.01-2.93.02-4.4.02-17.22,0-34.34-.7-50.91-1.38l-3.29-.13c-16.12-.66-30.68-1.2-45.5-1.2-8.53,0-16.31.18-23.78.56-50.42,2.54-102.5,4.11-159.22,4.8-17.19.21-34.7.49-51.64.77-34.5.56-70.18,1.14-105.32,1.14ZM821.73,969.64c.11,7.33,2.26,11.98,8.17,17.66,12.05,11.59,26.12,25.02,40.54,37.75,7.4,6.53,17.04,17.31,15.42,35.49-.34,3.77.08,9.24,3.98,13.88,4.75,5.66,8.55,6.38,11.34,6.38,1.67,0,3.54-.31,5.56-.91,10.94-3.27,10.58-7.09,10.35-9.62-2.62-28.39,12.63-44.83,27.89-57.16,9.24-7.47,18.78-16.38,30-28.03,6.97-7.24,7.7-13.98,2.19-20.05-2.89-3.18-5.86-4.79-8.83-4.79-3.34,0-7.4,1.96-11.44,5.5-11.33,9.95-22.84,21.21-34.23,33.46-2.45,2.63-9.89,10.65-21.27,10.65-10.82,0-18.33-7.37-21.15-10.15-4.3-4.23-8.41-8.56-12.39-12.76-8.61-9.08-16.74-17.66-25.99-24.37-3.27-2.38-5.55-2.65-6.97-2.65-2.13,0-4.62.73-7.4,2.16-3.7,1.91-5.63,4.44-5.74,7.53ZM1224.74,805.44c-62.46,0-109.8,47.89-110.12,111.39-.15,30.33,11.43,58.75,32.61,80.04,20.33,20.44,47.12,31.69,75.43,31.7h.56c30.39-.16,58.68-11.63,79.64-32.29,20.92-20.61,32.61-48.47,32.92-78.44.3-29.71-11.14-57.84-32.22-79.21-21.01-21.3-48.85-33.08-78.4-33.18h-.42ZM549.56,805.43c-60.61,0-110.18,50.16-110.51,111.81-.16,29.46,11.38,57.38,32.49,78.6,20.87,20.98,48.28,32.54,77.18,32.54,29.88-.03,57.85-11.37,78.86-31.95,20.96-20.53,32.7-48.04,33.06-77.45.38-30.64-10.98-59.28-32-80.65-20.73-21.08-48.64-32.76-78.6-32.9h-.49Z"/>
                  <path class="cls-2" d="M1513.86,36c37.44,0,64.01,26.67,65.35,69.96,1.95,63.12,1.17,126.17-5.97,189.07-8.39,73.9-22.03,146.68-45.08,217.42-4.7,14.42-3.82,25.88,3.08,39.17,41.79,80.47,68.31,165.69,76.62,256.2.78,8.5,4.4,10.63,11.54,10.63.18,0,.37,0,.55,0,22.47-.33,44.94-.42,67.41-.58,1.09,0,2.18-.01,3.27-.01,8.55,0,17.06.42,25.06,4.03,15.27,6.9,25.64,24.71,23.68,40.32-2.21,17.51-17.8,33.5-36.92,34.63-13.63.8-27.31,1.19-40.98,1.19-10.4,0-20.8-.23-31.18-.68-1.21-.05-2.38-.08-3.51-.08-13.6,0-21.19,4.14-18.68,20.63.16,1.04-.03,2.14-.11,3.2-2.35,31.71-2.49,33.57,25.41,33.57,1.73,0,3.57,0,5.53-.01,13.2-.05,26.4-.23,39.59-.23,5.53,0,11.06.03,16.59.12,23.62.37,43.93,19.03,44.2,40.05.27,21.3-19.22,39.83-43.7,40.18-14.96.22-29.92.3-44.88.3-19.82,0-39.64-.15-59.46-.32-.16,0-.32,0-.47,0-8.05,0-12.73,1.88-16.16,10.64-71.17,181.83-201.47,305.04-379.13,381.03-7.37,3.15-14.6,6.63-21.99,9.73q-20.15,8.44-10.5,28.78c46.79,98.34,83.3,200.18,102.94,307.56,8.88,48.53,14.29,97.5,5.1,146.65-12.26,65.6-67.11,111.04-133.56,111.52-1.43.01-2.86.02-4.29.02-17.95,0-35.72-.78-53.44-1.5-15.42-.63-30.8-1.22-46.23-1.22-8.21,0-16.43.17-24.69.58-52.8,2.66-105.66,4.14-158.53,4.78-52.2.64-104.49,1.9-156.73,1.9s-100.44-1.17-150.49-5.18c-47.33-3.8-42.98-169.68-41.97-184.22,7.03-100.6,34.78-196.31,71.26-289.78,12.83-32.88,26.74-65.31,43-96.62,6.3-12.12,3.51-17.06-8.83-21.88-57.94-22.64-113.12-50.9-164.3-86.34-107.95-74.77-186.73-173.39-236.75-294.65-6.46-15.66-14.29-22.23-30.07-22.23-1.25,0-2.56.04-3.91.12-14.24.86-28.52,1.11-42.81,1.11-16.22,0-32.45-.32-48.66-.45-25.29-.2-42.93-16.76-43.25-39.79-.33-23.68,17.54-40.52,43.85-40.73,10.68-.08,21.36-.14,32.04-.14,15.54,0,31.08.13,46.61.5.47.01.93.02,1.38.02,9.39,0,13.41-2.61,10.71-13.44-3.46-13.89,1.88-33.06-7.78-41.17-4.19-3.52-9.71-4.55-15.78-4.55-8.5,0-18.09,2.01-26.65,2.01-.41,0-.82,0-1.23-.01-2.16-.05-4.33-.07-6.49-.07-8.12,0-16.25.29-24.37.29-5.23,0-10.44-.12-15.64-.51-21.55-1.64-36.78-18.31-37.43-38.93-.66-21.13,12.97-38,35.9-39.46,10.96-.7,21.97-.96,32.98-.96,15.72,0,31.47.54,47.17,1.11.54.02,1.06.03,1.57.03,11.43,0,14.89-4.99,15.88-16.38,7.6-86.82,32.7-168.66,72.95-245.72,8.81-16.86,9.9-31.41,4.18-49.79-40.58-130.43-55.34-264.42-50.5-400.6,1.59-44.62,26.24-70.49,63.27-70.49,9.15,0,19.05,1.58,29.56,4.85,108.17,33.64,206.06,87.94,299.07,151.61,20.28,13.88,39.98,28.63,59.65,43.36,4.1,3.07,7.86,4.44,12.07,4.44,2.58,0,5.34-.52,8.45-1.47,73.54-22.57,147.23-34.03,220.96-34.03,69.9,0,139.84,10.29,209.81,31.18,5.54,1.65,10.42,2.61,15.07,2.61,7.5,0,14.37-2.5,22.35-8.6,95.55-73.12,198.29-133.99,311.14-176.8,13.46-5.11,27.17-9.7,41.01-13.66,8.84-2.53,17.32-3.74,25.31-3.74M548.72,1046.37h.11c71.37-.07,128.93-56.46,129.8-127.17.9-72.93-56.17-131.45-128.51-131.77-.19,0-.38,0-.56,0-70.49,0-128.13,58.11-128.51,129.71-.38,70.56,57.6,129.23,127.67,129.23M1222.65,1046.57c.22,0,.44,0,.66,0,72.6-.38,129.73-56.67,130.46-128.55.73-71.32-57.38-130.34-128.56-130.58-.15,0-.32,0-.48,0-72.2,0-127.75,56.02-128.12,129.3-.36,71.48,56.37,129.83,126.03,129.83M901.17,1098.81c3.4,0,6.97-.54,10.72-1.67,14.31-4.28,24.53-13.15,23.11-28.52-1.77-19.21,7.8-30.61,21.28-41.5,11.2-9.05,21.65-19.16,31.65-29.55,13.74-14.27,14.35-31.65,2.54-44.64-6.41-7.05-14.08-10.67-22.15-10.67-7.66,0-15.68,3.27-23.32,9.98-12.43,10.92-24.27,22.61-35.53,34.73-3.11,3.34-5.61,4.9-8.08,4.9-2.64,0-5.25-1.76-8.54-4.99-13.33-13.09-25.16-27.78-40.42-38.86-5.72-4.15-11.59-6.08-17.55-6.08-5.16,0-10.39,1.45-15.65,4.16-10.49,5.41-15.35,14.41-15.5,23.42.11,14.21,5.8,23.16,13.7,30.75,13.49,12.98,27.07,25.89,41.1,38.27,6.41,5.66,10.22,11.23,9.4,20.4-.87,9.73,1.7,19.41,8.12,27.06,6.96,8.29,15.33,12.81,25.12,12.81M1513.86,0c-11.46,0-23.31,1.73-35.21,5.13-14.48,4.14-29.25,9.06-43.88,14.61-56.34,21.38-112.6,48.12-167.22,79.48-51.11,29.35-102.6,63.8-153.03,102.39-.61.47-1.1.81-1.47,1.05-.75-.14-1.96-.42-3.77-.96-72.65-21.69-146.71-32.68-220.11-32.68s-150.63,11.39-225.75,33.86c-17.82-13.33-36.17-26.96-55.05-39.87-51.74-35.41-100.02-64.63-147.61-89.31-54.62-28.33-107.32-50.24-161.11-66.97-13.8-4.29-27.34-6.47-40.25-6.47-13.93,0-27.07,2.52-39.05,7.48-12.46,5.16-23.39,12.82-32.47,22.77-17.06,18.7-26.65,44.62-27.73,74.96-2.54,71.33.31,141.34,8.46,208.1,8.56,70.11,23.24,138.9,43.64,204.48,2.93,9.42,2.56,14.24-1.72,22.43-21.32,40.82-38.58,83.06-51.29,125.55-11.39,38.05-19.4,77.2-23.89,116.61-10.23-.31-20.36-.52-30.48-.52-12.81,0-24.35.34-35.27,1.04-20.49,1.31-38.79,10.01-51.54,24.51C5.84,821.59-.57,840.06.04,859.66c.59,18.9,7.95,36.73,20.71,50.21,6.64,7.01,14.53,12.62,23.24,16.62-8.34,3.88-15.84,9.14-22.21,15.64-13.81,14.11-21.26,33.17-20.98,53.65.29,21.13,8.74,40.44,23.8,54.38,14.39,13.32,33.98,20.74,55.16,20.91,6.04.05,12.22.12,18.2.2,10.04.12,20.43.25,30.75.25,16.78,0,30.68-.35,43.59-1.09,26.25,63.59,60.49,121.92,101.77,173.38,41.54,51.78,91.23,97.86,147.68,136.96,45.49,31.51,96.4,59.09,151.58,82.15-11.47,23.7-22.65,49.64-34.5,80.01-43.75,112.11-67.15,207.56-73.63,300.36-.52,7.46-2.1,55.27,2.97,102.49,3.03,28.16,7.94,51.29,14.61,68.74,4.9,12.84,10.8,23.01,18.03,31.09,13.65,15.25,28.95,19.45,39.38,20.29,45.06,3.61,93.79,5.3,153.37,5.3,35.29,0,71.04-.58,105.61-1.14,16.91-.27,34.4-.56,51.56-.77,56.95-.7,109.25-2.28,159.9-4.83,7.17-.36,14.66-.54,22.88-.54,14.49,0,28.85.54,44.77,1.19l3.34.14c16.72.69,34.01,1.39,51.57,1.39,1.52,0,3.03,0,4.55-.02,83.92-.6,153.29-58.54,168.69-140.9,10.76-57.53,3.47-113.04-5.07-159.74-18.05-98.66-50.63-197.93-99.49-303.04,2.68-1.21,5.25-2.36,7.79-3.44,46.58-19.92,90.23-43.13,129.74-69,40.5-26.51,77.82-56.57,110.93-89.33,33.78-33.44,64.12-70.63,90.17-110.56,24.12-36.96,45.18-77.19,62.7-119.77,15.43.12,31.52.22,47.53.22s31.3-.1,45.4-.3c20.69-.3,40.26-8.06,55.11-21.86,15.55-14.44,24.32-34.41,24.06-54.77-.25-19.99-9.4-39.65-25.1-53.94-6.1-5.55-12.92-10.12-20.23-13.6,23.84-10.53,41.76-33.19,45.13-59.89,1.94-15.39-1.63-31.77-10.03-46.12-8.17-13.94-20.43-25.12-34.53-31.5-14.05-6.35-27.48-7.23-39.89-7.23-1.17,0-2.35,0-3.52.02-4.87.03-9.74.07-14.61.1-10.25.07-20.71.13-31.2.23-4.9-39.35-13.19-78.49-24.72-116.65-13.29-43.99-31.25-87.91-53.38-130.52-2.29-4.41-2.61-5.88-.8-11.42,21.82-66.99,37.07-140.43,46.62-224.52,6.53-57.53,8.5-119.26,6.18-194.24-.93-30.02-10.97-55.92-29.04-74.92-18.36-19.3-44.04-29.93-72.3-29.93h0ZM548.72,1010.37c-24.07,0-46.95-9.67-64.42-27.23-17.71-17.8-27.38-41.17-27.25-65.81.28-51.78,41.78-93.9,92.51-93.9h.43c25.09.11,48.47,9.88,65.81,27.52,17.62,17.92,27.15,42,26.83,67.81-.3,24.62-10.13,47.63-27.66,64.81-17.62,17.26-41.12,26.78-66.18,26.81h-.08ZM1222.65,1010.57c-23.48,0-45.73-9.37-62.66-26.39-17.78-17.87-27.5-41.75-27.37-67.26.13-25.95,9.61-49.73,26.69-66.97,16.94-17.1,40.18-26.51,65.43-26.51h.32c24.75.08,48.06,9.97,65.68,27.83,17.69,17.93,27.29,41.51,27.04,66.39-.26,25.18-10.04,48.55-27.55,65.8-17.6,17.35-41.44,26.97-67.11,27.11h-.47ZM901.38,1027.88c2.39,0,4.75-.19,7.05-.58-2.65,4.65-4.75,9.48-6.31,14.49-.99-3.68-2.36-7.24-4.1-10.67-.63-1.24-1.3-2.43-1.98-3.57,1.71.22,3.49.34,5.34.34h0Z"/>
                </g>
                <path d="M1248.87,813.96c-35.03,0-63.42,46.56-63.42,104s28.39,104,63.42,104,63.42-46.56,63.42-104-28.39-104-63.42-104Z"/>
                <path d="M570.91,813.96c-35.03,0-63.42,46.56-63.42,104s28.39,104,63.42,104,63.42-46.56,63.42-104-28.39-104-63.42-104Z"/>
              </g>
              <g id="Paw_1" data-name="Paw 1">
                <path class="cls-6" d="M806.32,2002.66c-9.16-2.72-32.05-20.83-32.05-20.83,0,0-19.41,20.83-41.48,20.83-25.59,0-45.01-18.18-45.01-18.18,0,0-24.17,15.83-33.38,18.18-22.75,5.79-42.69-19.21-42.69-42.69v-64.53c0-37.63,30.79-68.41,68.41-68.41h100.46c37.63,0,68.41,30.79,68.41,68.41v64.53c0,23.48-20.18,49.36-42.69,42.69Z"/>
              </g>
              <g id="Paw_2" data-name="Paw 2">
                <path class="cls-6" d="M1085.97,2037.15c-9.16-2.72-32.05-20.83-32.05-20.83,0,0-19.41,20.83-41.48,20.83-25.59,0-45.01-18.18-45.01-18.18,0,0-24.17,15.83-33.38,18.18-22.75,5.79-42.69-19.21-42.69-42.69v-64.53c0-37.63,30.79-68.41,68.41-68.41h100.46c37.63,0,68.41,30.79,68.41,68.41v64.53c0,23.48-20.18,49.36-42.69,42.69Z"/>
              </g>
              <g id="Laptop_keyboard" data-name="Laptop keyboard">
                <g>
                  <path class="cls-8" d="M486.5,2048.78h1563.14c19.09,0,34.59,15.5,34.59,34.59v37.48c0,19.09-15.5,34.59-34.59,34.59H486.5c-19.09,0-34.59-15.5-34.59-34.59v-37.48c0-19.09,15.5-34.59,34.59-34.59Z"/>
                  <path class="cls-2" d="M2049.64,2058.78c13.58,0,24.59,11.01,24.59,24.59v37.48c0,13.58-11.01,24.59-24.59,24.59H486.5c-13.58,0-24.59-11.01-24.59-24.59v-37.48c0-13.58,11.01-24.59,24.59-24.59h1563.14M2049.64,2038.78H486.5c-24.59,0-44.59,20-44.59,44.59v37.48c0,24.59,20,44.59,44.59,44.59h1563.14c24.59,0,44.59-20,44.59-44.59v-37.48c0-24.59-20-44.59-44.59-44.59h0Z"/>
                </g>
              </g>
              <g id="Laptop_Layer_2" data-name="Laptop Layer 2">
                <g>
                  <path class="cls-3" d="M1017.86,2111.34c-8.31,0-16.04-3.73-21.22-10.24-5.17-6.5-7.06-14.88-5.19-22.98l169.3-731.95c4.98-21.55,17.27-41,34.58-54.77,17.32-13.77,39.04-21.35,61.16-21.35h948.9c30.13,0,58.17,13.53,76.92,37.11,18.75,23.58,25.61,53.95,18.82,83.3l-153.32,662.84c-7.9,34.17-37.92,58.04-72.99,58.04H1017.86Z"/>
                  <path class="cls-2" d="M2205.39,1280.05c56.78,0,98.79,52.84,86,108.16l-153.32,662.84c-6.81,29.44-33.03,50.29-63.25,50.29H1017.86c-11,0-19.15-10.24-16.67-20.96l169.3-731.95c9.26-40.03,44.91-68.38,86-68.38h948.9M2205.39,1260.05h-948.9c-24.37,0-48.3,8.35-67.38,23.52-19.08,15.17-32.61,36.6-38.1,60.35l-169.3,731.95c-2.56,11.08.03,22.55,7.11,31.45,7.08,8.91,17.67,14.01,29.04,14.01h1056.96c19.12,0,37.89-6.55,52.85-18.45,14.96-11.9,25.58-28.71,29.89-47.33l153.32-662.84c7.48-32.34-.08-65.79-20.74-91.78-20.66-25.98-51.55-40.89-84.74-40.89h0Z"/>
                </g>
              </g>
              <g id="Laptop_Layer" data-name="Laptop Layer">
                <path class="cls-4" d="M2068.45,2101.34h-1040.1l174.15-752.91c9.26-40.03,44.91-68.37,86-68.37h916.9c56.78,0,98.79,52.84,86,108.16l-151.88,656.63c-7.65,33.08-37.11,56.5-71.06,56.5Z"/>
              </g>
              <g id="Cat_Pupil_2" data-name="Cat Pupil 2">
                <path class="cls-7" d="M1252.85,984.26c-14.26,0-26.16-11.9-26.16-26.16s11.9-26.16,26.16-26.16,26.16,11.9,26.16,26.16-11.9,26.16-26.16,26.16Z"/>
              </g>
              <g id="Cat_Pupil_1" data-name="Cat Pupil 1">
                <path class="cls-7" d="M580.16,987.71c-14.26,0-26.16-11.9-26.16-26.16s11.9-26.16,26.16-26.16,26.16,11.9,26.16,26.16-11.9,26.16-26.16,26.16Z"/>
              </g>
            </svg>
        </div>
        """, height=260
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
    # Dynamically capture original file name
    base_file_name = os.path.splitext(uploaded_file.name)[0]
    
    with st.spinner("Buddy is processing your data..."):
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
                        file_name=f"Formatted_{base_file_name}.xlsx",
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
                            st.download_button(
                                "📥 Save Standard ZIP", 
                                data=zip_data, 
                                file_name=f"{base_file_name}_Images.zip", 
                                mime="application/zip", 
                                type="primary"
                            )

                    with col2:
                        st.markdown("### AI Background Removal")
                        if st.button("🤖 Start Rembg Download", use_container_width=True):
                            bria_session = get_bria_session()
                            zip_data = batch_download_images(all_img_rows, use_rembg=True, rembg_session=bria_session)
                            st.download_button(
                                "📥 Save Rembg ZIP", 
                                data=zip_data, 
                                file_name=f"{base_file_name}_Rembg_Images.zip", 
                                mime="application/zip", 
                                type="primary"
                            )
else:
    st.info("👈 Please paste your Google Sheet Link and upload your Campaign Excel file in the sidebar to begin.")
