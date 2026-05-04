import streamlit as st
import pandas as pd
import os
import shutil
import re
import base64
import streamlit.components.v1 as components
from openpyxl import load_workbook
from PIL import Image
import io

# --- 1. Page Configuration ---
st.set_page_config(page_title="LED Direct Product System", layout="wide")

# --- 2. Advanced UI Design Styling ---
st.markdown("""
    <style>
    /* 全局背景优化 */
    .main { background-color: #ffffff; }
    .block-container { padding: 2rem 3rem; }
    
    /* 🚀 左侧导航容器：增加垂直分割线 */
    [data-testid="column"]:nth-of-type(1) {
        background-color: #f8f9fa; /* 极简浅灰背景 */
        padding: 25px 30px 25px 25px !important; /* 右侧留出更多间距给分割线 */
        border-radius: 12px 0 0 12px;
        border: 1px solid #eeeeee;
        border-right: 1px solid #dddddd !important; /* 💡 核心改动：显眼的垂直分割线 */
        box-shadow: inset 0 1px 3px rgba(0,0,0,0.02);
    }
    
    /* 🚀 右侧详情容器 */
    [data-testid="column"]:nth-of-type(2) {
        background-color: #ffffff;
        padding-left: 60px !important; /* 💡 增加左侧间距，让内容远离分割线 */
        padding-top: 10px !important;
    }
    
    /* 侧边栏按钮：更清爽的现代感 */
    div.stButton > button { 
        padding: 4px 5px !important; 
        font-size: 10px !important; 
        height: auto !important;
        border: 1px solid #e0e0e0 !important;
        background-color: white !important;
        color: #666 !important;
        transition: all 0.2s ease;
    }
    div.stButton > button:hover {
        border-color: #ff0000 !important;
        color: #ff0000 !important;
        background-color: #fffafa !important;
    }

    /* 价格卡片样式 */
    .price-card { 
        background-color: #ffffff; 
        border-left: 5px solid #ff0000; 
        padding: 18px 22px; 
        margin-bottom: 20px; 
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border-radius: 4px;
    }
    .price-amount { color: #FF0000; font-size: 38px; font-weight: 900; letter-spacing: -1px; }

    /* 横排规格表格样式 */
    .spec-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    .spec-table thead th { 
        background-color: #1a1a1a; color: white; text-align: left; 
        padding: 12px 15px; font-size: 11px; text-transform: uppercase; letter-spacing: 1px;
    }
    .spec-table tbody td { padding: 14px 15px; color: #333; border-bottom: 1px solid #f0f0f0; font-size: 13px; }

    /* 九宫格图片容器：纯白底图 */
    .gallery-container {
        width: 100%;
        aspect-ratio: 1 / 1; 
        background-color: #ffffff; 
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        overflow: hidden;
        border: 1px solid #eeeeee;
        margin-bottom: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    }
    .gallery-container img {
        max-width: 88%;
        max-height: 88%;
        object-fit: contain; 
    }
    </style>
    """, unsafe_allow_html=True)

# --- 🚀 核心辅助函数 ---
def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

def get_image_base64(path):
    if os.path.exists(path):
        try:
            with open(path, "rb") as f: return base64.b64encode(f.read()).decode('utf-8')
        except: return ""
    return ""

# --- 3. Robust Data Processing ---
ORIGINAL_FILE = "Test.xlsx"
IMAGE_DIR = "product_images"

@st.cache_resource
def extract_images_robustly():
    """高清提取引擎"""
    try:
        if os.path.exists(IMAGE_DIR):
            shutil.rmtree(IMAGE_DIR, ignore_errors=True)
        os.makedirs(IMAGE_DIR, exist_ok=True) 

        if not os.path.exists(ORIGINAL_FILE): return
        pxl_doc = load_workbook(ORIGINAL_FILE)
        
        sheet = pxl_doc.worksheets[0]
        for s in pxl_doc.worksheets:
            if s.cell(row=1, column=1).value == "MODEL" or s.cell(row=1, column=2).value == "CODE":
                sheet = s
                break

        code_map = {}
        for row in range(2, sheet.max_row + 1):
            val = sheet.cell(row=row, column=2).value 
            if val: code_map[row] = str(val).strip()

        if hasattr(sheet, '_images'):
            for img in sheet._images:
                try:
                    col_idx = img.anchor._from.col
                    row_idx = img.anchor._from.row + 1 
                    if col_idx == 2:
                        code_name = code_map.get(row_idx)
                        if code_name:
                            pil_img = None
                            if hasattr(img, '_data'): 
                                pil_img = Image.open(io.BytesIO(img._data()))
                            elif hasattr(img, 'image'): 
                                pil_img = img.image
                            
                            if pil_img:
                                if pil_img.mode in ("RGBA", "P"): pil_img = pil_img.convert("RGB")
                                pil_img.save(f"{IMAGE_DIR}/{code_name}.png", "PNG", optimize=False, compress_level=1)
                except: continue
        pxl_doc.close()
    except Exception as e:
        st.error(f"Extraction Error: {e}")

@st.cache_data
def load_data():
    extract_images_robustly()
    xl = pd.ExcelFile(ORIGINAL_FILE)
    target_sheet = xl.sheet_names[0]
    for name in xl.sheet_names:
        temp_df = pd.read_excel(ORIGINAL_FILE, sheet_name=name, nrows=1)
        if "MODEL" in [str(c).upper() for c in temp_df.columns]:
            target_sheet = name
            break
            
    df = pd.read_excel(ORIGINAL_FILE, sheet_name=target_sheet)
    df.columns = [str(c).upper().strip().replace('\n', '') for c in df.columns]
    if "MODEL" in df.columns: df["MODEL"] = df["MODEL"].ffill()
    
    dim_col = next((c for c in df.columns if "DIM" in c), None)
    if dim_col:
        df = df[df[dim_col].astype(str).str.upper().str.contains("DALI")]
    return df

# --- 4. Main UI Layout ---
try:
    df = load_data()
    unique_series = sorted(df["MODEL"].dropna().unique(), key=natural_sort_key)
    if 'selected_series' not in st.session_state:
        st.session_state.selected_series = unique_series[0]

    # --- 布局比例调整 ---
    left_col, right_col = st.columns([3, 5]) 

    # --- 【左侧：带分割线的画廊侧边栏】 ---
    with left_col:
        st.markdown("<h3 style='color:#333; margin-bottom:20px; font-weight:800; letter-spacing:-0.5px;'>Series Gallery</h3>", unsafe_allow_html=True)
        grid_cols = 4 
        for i in range(0, len(unique_series), grid_cols):
            cols = st.columns(grid_cols)
            for j in range(grid_cols):
                idx = i + j
                if idx < len(unique_series):
                    s_name = unique_series[idx]
                    s_data = df[df["MODEL"] == s_name]
                    c_code = str(s_data.iloc[0]["CODE"]).strip()
                    path = f"{IMAGE_DIR}/{c_code}.png"
                    
                    with cols[j]:
                        img_b64_thumb = get_image_base64(path)
                        if img_b64_thumb:
                            st.markdown(f"""
                                <div class="gallery-container">
                                    <img src="data:image/png;base64,{img_b64_thumb}">
                                </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="gallery-container" style="font-size:9px; color:#999;">No Img</div>', unsafe_allow_html=True)
                        
                        is_act = (st.session_state.selected_series == s_name)
                        if st.button(s_name, key=f"n_{s_name}", use_container_width=True, type="primary" if is_act else "secondary"):
                            st.session_state.selected_series = s_name
                            st.rerun()
            st.write("")

    # --- 【右侧：产品详情区】 ---
    with right_col:
        act_s = st.session_state.selected_series
        st.markdown(f"<h2 style='font-weight:900; color:#1a1a1a; margin-top:0;'>📦 {act_s}</h2>", unsafe_allow_html=True)
        
        s_df = df[df["MODEL"] == act_s]
        c_list = sorted(s_df["CODE"].dropna().unique(), key=natural_sort_key)
        sel_c = st.selectbox("Model Variant Select:", c_list)
        res_r = s_df[s_df["CODE"] == sel_c].iloc[0]
        
        # 详情图比例优化
        c1, c2 = st.columns([0.8, 2])
        det_path = f"{IMAGE_DIR}/{str(sel_c).strip()}.png"
        i_b64 = get_image_base64(det_path)

        with c1:
            if os.path.exists(det_path):
                st.image(det_path, use_container_width=True)
            else:
                st.warning("Missing data.")
        
        with c2:
            price_col = next((c for c in df.columns if "PRICE" in c), None)
            if price_col:
                p_v = str(res_r[price_col]).strip()
                st.markdown(f'<div class="price-card"><div style="font-size:12px; font-weight:bold; color:#666; letter-spacing:1px;">PRICE UNIT (DALI)</div><div class="price-amount">{p_v if p_v.startswith("$") else "$"+p_v}</div></div>', unsafe_allow_html=True)

            order = ['MODEL', 'CODE', 'IMAGE', 'COLOR', 'WATTAGE', 'CCT', 'CRI', 'DEGREE', 'DIM METHOD', 'IP', 'DIMENSION', 'CUTOUT', 'INSTALLATION', 'PRICE']
            copy_v = []
            for k in order:
                if k == 'IMAGE': copy_v.append("")
                else:
                    m = next((c for c in res_r.index if c.upper() == k.upper()), None)
                    copy_v.append(str(res_r[m]) if m else "")
            
            components.html(f"""
                <script>
                function copyT() {{ navigator.clipboard.writeText(`{"\\t".join(copy_v)}`); alert('✅ Data Copied to Clipboard!'); }}
                async function copyI() {{
                    const b = "{i_b64}"; if(!b) return;
                    const blob = await (await fetch('data:image/png;base64,'+b)).blob();
                    await navigator.clipboard.write([new ClipboardItem({{'image/png': blob}})]);
                    alert('🖼️ Image Copied to Clipboard!');
                }}
                </script>
                <div style="display:flex; gap:10px;">
                    <button onclick="copyT()" style="background:#007bff;color:white;border:none;padding:10px 18px;border-radius:6px;cursor:pointer;font-weight:bold;font-family:sans-serif;font-size:13px;">📋 Copy Row Data</button>
                    <button onclick="copyI()" style="background:#28a745;color:white;border:none;padding:10px 18px;border-radius:6px;cursor:pointer;font-weight:bold;font-family:sans-serif;font-size:13px;">🖼️ Copy Image</button>
                </div>
            """, height=60)

        st.markdown("<br>", unsafe_allow_html=True)
        # 横排规格展示
        d_list = []
        for o in order:
            m = next((c for c in res_r.index if c.upper() == o.upper()), None)
            d_list.append({"Key": o, "Value": res_r[m] if m else "-"})
        
        h_t = '<div style="overflow-x:auto;"><table class="spec-table"><thead><tr>'
        for col in d_list: h_t += f'<th>{col["Key"]}</th>'
        h_t += '</tr></thead><tbody><tr>'
        for col in d_list: h_t += f'<td>{col["Value"]}</td>'
        h_t += '</tr></tbody></table></div>'
        st.markdown(h_t, unsafe_allow_html=True)

except Exception as e:
    st.error(f"System Error: {e}")
