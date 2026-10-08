import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import io

# 1. ตั้งค่า Page Config
st.set_page_config(
    page_title="Recorder NB1 Debinder",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. บังคับ Dark Mode CSS และตั้งค่าตัวหนังสือปุ่มดาวน์โหลดเป็นสีขาว
st.markdown("""
    <style>
        /* ซ่อนแถบขาว Header ด้านบน */
        header[data-testid="stHeader"] {
            background-color: transparent !important;
            display: none !important;
        }
        [data-testid="stToolbar"] {
            display: none !important;
        }
        
        /* ตั้งค่าพื้นหลัง Dark Mode */
        html, body, .stApp, [data-testid="stAppViewContainer"] {
            background-color: #0e1117 !important;
            color: #ffffff !important;
        }
        [data-testid="stSidebar"] {
            background-color: #161b22 !important;
        }
        .stMarkdown, h1, h2, h3, p, span, label {
            color: #ffffff !important;
        }

        /* ปุ่มเคลียร์ข้อมูลใน Sidebar */
        [data-testid="stSidebar"] div.stButton > button {
            background-color: #21262d !important;
            color: #ffffff !important;
            border: 1px solid #F0B90B !important;
            font-weight: bold !important;
            width: 100% !important;
            padding: 8px 16px !important;
        }
        [data-testid="stSidebar"] div.stButton > button:hover {
            background-color: #F0B90B !important;
            color: #000000 !important;
        }

        /* กล่อง File Uploader */
        [data-testid="stFileUploader"] {
            background-color: #161b22 !important;
            border: 1.5px solid #F0B90B !important;
            border-radius: 8px !important;
            padding: 10px !important;
        }
        [data-testid="stFileUploader"] section {
            background-color: #1c2128 !important;
            border: 1px dashed #F0B90B !important;
            border-radius: 6px !important;
        }
        [data-testid="stFileUploader"] section div, 
        [data-testid="stFileUploader"] section span,
        [data-testid="stFileUploader"] section small {
            color: #e6edf3 !important;
        }

        /* การ์ดไฟล์ที่อัปโหลดแล้ว */
        [data-testid="stFileUploaderFileData"],
        [data-testid="stFileUploaderFileData"] > div,
        [data-testid="stFileUploaderFile"] {
            background-color: #21262d !important;
            border: 1px solid #F0B90B !important;
            border-radius: 6px !important;
        }
        [data-testid="stFileUploaderFileData"] *,
        [data-testid="stFileUploaderFile"] * {
            color: #ffffff !important;
            font-weight: bold !important;
        }

        /* 1. ปรับแถบ Expander */
        [data-testid="stExpander"] {
            background-color: #161b22 !important;
            border: 1px solid #30363d !important;
            border-radius: 8px !important;
        }
        [data-testid="stExpander"] details summary {
            background-color: #21262d !important;
            color: #ffffff !important;
            border-radius: 8px !important;
        }
        [data-testid="stExpander"] details summary * {
            color: #ffffff !important;
        }

        /* 2. ปรับแต่งตาราง Dataframe */
        [data-testid="stDataFrame"] {
            background-color: #161b22 !important;
            border: 1px solid #30363d !important;
            border-radius: 8px !important;
        }
        div[data-testid="stDataFrame"] div[role="grid"] {
            background-color: #161b22 !important;
            color: #ffffff !important;
        }
        div[data-testid="stDataFrame"] div[role="columnheader"] {
            background-color: #21262d !important;
            color: #ffffff !important;
        }

        /* 3. ปรับแต่งกล่องพิมพ์ข้อความ (Text Input) */
        div[data-baseweb="input"] {
            background-color: #21262d !important;
            border: 1px solid #30363d !important;
            color: #ffffff !important;
            border-radius: 6px !important;
        }
        div[data-baseweb="input"] input {
            background-color: #21262d !important;
            color: #ffffff !important;
        }

        /* 4. ปรับแต่งปุ่มดาวน์โหลด Excel (บังคับข้อความให้เป็นสีขาวสด) */
        div.stDownloadButton > button {
            background-color: #21262d !important;
            border: 1.5px solid #F0B90B !important;
            border-radius: 6px !important;
            padding: 8px 16px !important;
            transition: all 0.2s ease-in-out;
        }
        div.stDownloadButton > button, 
        div.stDownloadButton > button *,
        div.stDownloadButton > button p,
        div.stDownloadButton > button span {
            color: #ffffff !important;
            font-weight: bold !important;
            font-size: 15px !important;
        }
        div.stDownloadButton > button:hover {
            background-color: #F0B90B !important;
            border-color: #F0B90B !important;
        }
        div.stDownloadButton > button:hover,
        div.stDownloadButton > button:hover *,
        div.stDownloadButton > button:hover p,
        div.stDownloadButton > button:hover span {
            color: #000000 !important;
        }
    </style>
""", unsafe_allow_html=True)

# 3. แสดงชื่อโปรแกรมหลัก
st.title("🏭 Recorder NB1 Debinder")

# 4. ฟังก์ชันอ่านไฟล์ทีละบรรทัดและดึงคอลัมน์ตรงๆ
def parse_single_file(uploaded_file):
    uploaded_file.seek(0)
    raw_bytes = uploaded_file.read()
    
    text_content = None
    for enc in ['cp932', 'shift_jis', 'utf-8', 'tis-620', 'latin1']:
        try:
            text_content = raw_bytes.decode(enc)
            break
        except Exception:
            continue
            
    if text_content is None:
        text_content = raw_bytes.decode('utf-8', errors='ignore')

    lines = text_content.splitlines()
    parsed_rows = []
    
    for line in lines:
        line_str = line.strip()
        if line_str.startswith("20") and "," in line_str:
            parts = [p.strip() for p in line_str.split(",")]
            
            if len(parts) >= 17:
                try:
                    dt_val = parts[0]
                    z1 = float(parts[4])   # TH_CH1Ave.
                    z2 = float(parts[7])   # TH_CH2Ave.
                    z3 = float(parts[10])  # TH_CH3Ave.
                    z4 = float(parts[13])  # TH_CH4Ave.
                    comb = float(parts[16]) # TH_CH5Ave.
                    
                    parsed_rows.append({
                        "DateTime": dt_val,
                        "Zone #1": z1,
                        "Zone #2": z2,
                        "Zone #3": z3,
                        "Zone #4": z4,
                        "Combustion Air Temp": comb
                    })
                except ValueError:
                    continue

    if not parsed_rows:
        return pd.DataFrame()

    df = pd.DataFrame(parsed_rows)
    df["DateTime"] = pd.to_datetime(df["DateTime"], errors="coerce")
    return df.dropna(subset=["DateTime"]).reset_index(drop=True)

def process_multiple_files(uploaded_files):
    combined_dfs = []
    for file in uploaded_files:
        df_single = parse_single_file(file)
        if not df_single.empty:
            combined_dfs.append(df_single)
            
    if not combined_dfs:
        return pd.DataFrame()

    full_df = pd.concat(combined_dfs, ignore_index=True)
    full_df = full_df.drop_duplicates(subset=["DateTime"]).sort_values("DateTime").reset_index(drop=True)
    return full_df

# ฟังก์ชันแปลง DataFrame เป็น Binary สำหรับดาวน์โหลดเป็นไฟล์ Excel (.xlsx)
def to_excel_bytes(dataframe):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export = dataframe.copy()
        if pd.api.types.is_datetime64_any_dtype(df_export["DateTime"]):
            df_export["DateTime"] = df_export["DateTime"].dt.strftime('%Y-%m-%d %H:%M:%S')
        df_export.to_excel(writer, index=False, sheet_name='Debinder Data')
    output.seek(0)
    return output.getvalue()


# ========================================================
# 5. เมนู Sidebar (ปรับปรุงระบบเคลียร์ไฟล์)
# ========================================================
st.sidebar.header("📁 เมนูอัปโหลดข้อมูล")

# 5.1 สร้าง State สำหรับตัวแปร Key ของ File Uploader
if "file_uploader_key" not in st.session_state:
    st.session_state["file_uploader_key"] = 0

# 5.2 เมื่อกดปุ่ม ให้เพิ่มค่า Key ขึ้น 1 (เป็นการบังคับรีเซ็ต Widget)
if st.sidebar.button("🧹 เคลียร์ข้อมูลไฟล์เก่าทั้งหมด"):
    st.cache_data.clear()
    st.session_state["file_uploader_key"] += 1 
    st.rerun()

# 5.3 ส่งค่า Key เข้าไปในกล่องอัปโหลด
uploaded_files = st.sidebar.file_uploader(
    "อัปโหลดไฟล์ CSV (.csv) ได้มากกว่า 1 ไฟล์", 
    type=["csv"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['file_uploader_key']}" 
)
# ========================================================


# 6. แสดงผลกราฟและปุ่มเลือกดาวน์โหลด Excel
if uploaded_files:
    raw_df = process_multiple_files(uploaded_files)
    
    if raw_df.empty:
        st.error("⚠️ ไม่สามารถอ่านข้อมูลจากไฟล์ที่อัปโหลดได้ กรุณาตรวจสอบว่าเป็นไฟล์ CSV จากเครื่อง Recorder หรือไม่")
    else:
        st.sidebar.success(f"รวมข้อมูลสำเร็จ {len(uploaded_files)} ไฟล์ ({len(raw_df)} แถว)")

        st.sidebar.markdown("---")
        st.sidebar.header("🎛️ Dynamic Controls")
        
        min_time = raw_df["DateTime"].min().to_pydatetime()
        max_time = raw_df["DateTime"].max().to_pydatetime()
        
        selected_time = st.sidebar.slider(
            "⏱️ ช่วงเวลา:",
            min_value=min_time,
            max_value=max_time,
            value=(min_time, max_time),
            format="MM-DD HH:mm"
        )
        
        df = raw_df[(raw_df["DateTime"] >= selected_time[0]) & (raw_df["DateTime"] <= selected_time[1])].copy()

        st.subheader("📊 Debinder Temperature & Combustion Air Monitor")

        # สร้างกราฟ 2 แกน Y
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        # สีตามเครื่องจริง: Z#1 Red, Z#2 Blue, Z#3 Green, Z#4 Orange
        zone_colors = ["#FF3333", "#1E90FF", "#2ED573", "#FF9F1A"]
        
        # สี Combustion Air Temp: สีฟ้าสว่าง (Cyan) มองเห็นชัดบน Dark Mode
        combustion_color = "#00E5FF"

        # 1. Zone #1 - #4 (แกน Y ซ้ายมือ)
        for i in range(1, 5):
            fig.add_trace(
                go.Scatter(
                    x=df["DateTime"],
                    y=df[f"Zone #{i}"],
                    name=f"Zone #{i}",
                    mode="lines",
                    line=dict(color=zone_colors[i-1], width=2)
                ),
                secondary_y=False
            )

        # 2. Combustion Air Temp (แกน Y ขวามือ)
        fig.add_trace(
            go.Scatter(
                x=df["DateTime"],
                y=df["Combustion Air Temp"],
                name="Combustion Air Temp",
                mode="lines",
                line=dict(color=combustion_color, width=2, dash="dash")
            ),
            secondary_y=True
        )

        fig.update_layout(
            template="plotly_dark",
            plot_bgcolor="#161b22",
            paper_bgcolor="#0e1117",
            hovermode="x unified",
            showlegend=True,
            legend=dict(
                font=dict(color="#FFFFFF", size=12, family="Arial Bold"),
                bgcolor="rgba(27, 31, 36, 0.95)",
                bordercolor="#F0B90B",
                borderwidth=1.5,
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left",
                x=1.05
            ),
            xaxis=dict(
                title=dict(text="Date & Time", font=dict(color="#FFFFFF", size=12)),
                tickfont=dict(color="#CCCCCC", size=10),
                showgrid=True,
                gridcolor="rgba(255,255,255,0.08)",
                linecolor="#555555",
                type="date"
            ),
            yaxis=dict(
                title=dict(text="Zone Temperature (°C)", font=dict(color="#FFFFFF", size=12)),
                tickfont=dict(color="#CCCCCC", size=10),
                showgrid=True,
                gridcolor="rgba(255,255,255,0.08)",
                zeroline=False,
                linecolor="#555555",
                range=[0, 400]
            ),
            yaxis2=dict(
                title=dict(text="Combustion Air Temp (°C)", font=dict(color=combustion_color, size=12)),
                tickfont=dict(color=combustion_color, size=10),
                showgrid=False,
                overlaying="y",
                side="right",
                linecolor=combustion_color,
                range=[0, 150]
            ),
            height=500,
            margin=dict(l=60, r=180, t=30, b=40)
        )

        st.plotly_chart(fig, use_container_width=True)

        # ส่วนตรวจสอบและเลือกดาวน์โหลด Excel (.xlsx)
        with st.expander("📋 ตรวจสอบและเลือกดาวน์โหลดตารางข้อมูล Excel (.xlsx)"):
            st.dataframe(df)
            
            st.markdown("---")
            st.markdown("##### 📥 ตัวเลือกการดาวน์โหลดไฟล์ Excel")
            
            col_opt1, col_opt2 = st.columns([2, 1])
            with col_opt1:
                custom_filename = st.text_input(
                    "ตั้งชื่อไฟล์ดาวน์โหลด:", 
                    value="combined_debinder_data.xlsx"
                )
                if not custom_filename.endswith('.xlsx'):
                    custom_filename += '.xlsx'
                    
            with col_opt2:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                excel_bytes = to_excel_bytes(df)
                st.download_button(
                    label="📊 ดาวน์โหลดไฟล์ Excel",
                    data=excel_bytes,
                    file_name=custom_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

else:
    st.info("👈 กรุณาเลือกอัปโหลดไฟล์ (.csv) ที่เมนูด้านซ้าย สามารถเลือกอัปโหลดได้มากกว่า 1 ไฟล์")
