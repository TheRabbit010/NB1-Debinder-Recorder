import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import io

# 1. ตั้งค่า Page Config (แสดง Sidebar ค้างไว้เป็นค่าเริ่มต้น)
st.set_page_config(
    page_title="Recorder NB1 Debinder",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. CSS คลีนๆ เพื่อเปิดเมนูขวาบน (System/Light/Dark) และปุ่มเปิด-ปิด Sidebar
st.markdown("""
    <style>
        /* ซ่อนเฉพาะปุ่ม Deploy และ Footer */
        .stDeployButton { display: none !important; }
        footer { visibility: hidden !important; display: none !important; }
        
        /* ตกแต่งไอคอนลูกศรเปิด-ปิด Sidebar ให้เป็นสีเหลืองทอง มองเห็นเด่นชัดทุกโหมด */
        [data-testid="collapsedControl"] svg,
        [data-testid="stSidebarCollapseButton"] svg {
            fill: #F0B90B !important;
            color: #F0B90B !important;
        }
    </style>
""", unsafe_allow_html=True)

# 3. แสดงชื่อโปรแกรมหลัก
st.title("🏭 Recorder NB1 Debinder")

# 4. ฟังก์ชันอ่านไฟล์พร้อมระบบตรวจเช็คไฟล์ Recorder NB1 Debinder อัตโนมัติ
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
    
    # 🔍 ตรวจสอบเบื้องต้นว่าเป็นไฟล์ของ Recorder NB1 Debinder หรือไม่
    is_debinder_file = False
    for line in lines[:50]:
        # ต้องมีข้อมูลของ Debinder / TH_CH1 ถึง TH_CH5
        if "TH_CH1" in line or "TH_CH" in line or "Debinder" in line or line.startswith("20"):
            is_debinder_file = True
            break
            
    if not is_debinder_file:
        return None  # คืนค่า None เพื่อแจ้งเตือนว่าไม่ใช่ไฟล์ NB1 Debinder

    for line in lines:
        line_str = line.strip()
        if line_str.startswith("20") and "," in line_str:
            parts = [p.strip() for p in line_str.split(",")]
            
            if len(parts) >= 17:
                try:
                    dt_val = parts[0]
                    z1 = float(parts[4])    # TH_CH1Ave.
                    z2 = float(parts[7])    # TH_CH2Ave.
                    z3 = float(parts[10])   # TH_CH3Ave.
                    z4 = float(parts[13])   # TH_CH4Ave.
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
    invalid_files = []

    for file in uploaded_files:
        df_single = parse_single_file(file)
        if df_single is None:
            invalid_files.append(file.name)
        elif not df_single.empty:
            combined_dfs.append(df_single)
            
    # 🚨 แจ้งเตือนและหยุดทำงานทันทีหากพบไฟล์ที่ไม่ใช่ RECORDER NB1 Debinder
    if invalid_files:
        invalid_file_names = ", ".join(f"'{name}'" for name in invalid_files)
        st.error(
            f"❌ **ตรวจพบไฟล์ที่ไม่ใช่ RECORDER NB1 DEBINDER:** {invalid_file_names}\n\n"
            f"📌 *สาเหตุ:* โครงสร้างคอลัมน์ไม่ตรงตามมาตรฐานของ Recorder NB1 Debinder\n\n"
            f"กรุณากดปุ่ม **'🧹 เคลียร์ข้อมูลไฟล์เก่าทั้งหมด'** ด้านซ้าย แล้วเลือกอัปโหลดไฟล์ใหม่อีกครั้ง"
        )
        st.stop()  # หยุดการทำงานทันที

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
if st.sidebar.button("🧹 เคลียร์ข้อมูลไฟล์เก่าทั้งหมด", type="primary"):
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
        st.error("⚠️ ไม่สามารถอ่านข้อมูลจากไฟล์ที่อัปโหลดได้ กรุณาตรวจสอบว่าเป็นไฟล์ CSV จากเครื่อง Recorder NB1 Debinder หรือไม่")
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
            hovermode="x unified",
            showlegend=True,
            legend=dict(
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left",
                x=1.05
            ),
            xaxis=dict(
                title=dict(text="Date & Time"),
                showgrid=True,
                type="date"
            ),
            yaxis=dict(
                title=dict(text="Zone Temperature (°C)"),
                showgrid=True,
                zeroline=False,
                range=[0, 400]
            ),
            yaxis2=dict(
                title=dict(text="Combustion Air Temp (°C)"),
                showgrid=False,
                overlaying="y",
                side="right",
                range=[0, 150]
            ),
            height=500,
            margin=dict(l=60, r=180, t=30, b=40)
        )

        # รองรับการสลับโทนสี Light/Dark ตามธีมหลัก Streamlit
        st.plotly_chart(fig, use_container_width=True, theme="streamlit")

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
                    use_container_width=True,
                    type="primary"
                )

else:
    st.info("👈 กรุณาเลือกอัปโหลดไฟล์ (.csv) ที่เมนูด้านซ้าย สามารถเลือกอัปโหลดได้มากกว่า 1 ไฟล์")
