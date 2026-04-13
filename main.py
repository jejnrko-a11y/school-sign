import streamlit as st
from streamlit_drawable_canvas import st_canvas
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from PIL import Image
import io
import base64
import xlsxwriter

# --- 앱 설정 ---
st.set_page_config(page_title="재구조화 동의서 관리 시스템", layout="centered")

# 구글 시트 연결
conn = st.connection("gsheets", type=GSheetsConnection)

# --- [추가] 데이터 로드 캐싱 함수 ---
@st.cache_data(ttl=10) # 10초간 데이터 유지 (서명 중 끊김 방지)
def get_cached_data(_conn):
    df = _conn.read(worksheet="동의서양식", ttl=0)
    # 필수 컬럼 자동 생성
    for col in ['번호', '성함', '동의여부', '서명']:
        if col not in df.columns:
            df[col] = ""
    return df.fillna("")

# --- [함수] 엑셀 생성 로직 ---
def generate_excel_with_images(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_display = df[['번호', '성함', '동의여부', '서명']]
        df_display.to_excel(writer, index=False, sheet_name='동의서_결과')
        workbook  = writer.book
        worksheet = writer.sheets['동의서_결과']
        format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.set_column('A:A', 8); worksheet.set_column('B:B', 12); worksheet.set_column('C:C', 12); worksheet.set_column('D:D', 25)
        for idx, row in df.iterrows():
            row_num = idx + 1; worksheet.set_row(row_num, 60)
            sig_data = str(row.get('서명', ''))
            if sig_data.startswith("'"): sig_data = sig_data[1:]
            if len(sig_data) > 100:
                try:
                    img_io = io.BytesIO(base64.b64decode(sig_data))
                    worksheet.insert_image(row_num, 3, f'sig_{idx}.png', {'image_data': img_io, 'x_scale': 0.5, 'y_scale': 0.5, 'x_offset': 5, 'y_offset': 5})
                    worksheet.write(row_num, 3, "", format_center) 
                except: pass
    return output.getvalue()

# --- 화면 구성 ---
st.title("2026 경기기공 재구조화 동의서")
menu_tab1, menu_tab2 = st.tabs(["✍️ 동의서 서명", "📊 관리자 메뉴"])

# --- 데이터 로드 (캐시 사용) ---
df_all = get_cached_data(conn)
unsigned_df = df_all[df_all['서명'].apply(lambda x: str(x).strip() == "")]
unsigned_list = unsigned_df['성함'].tolist()
total_count = len(df_all)
signed_count = total_count - len(unsigned_df)

# --- 탭 1: 서명하기 ---
with menu_tab1:
    st.info("성함을 선택하신 후 의견을 표시하고 서명해 주세요.")
    st.progress(signed_count / total_count if total_count > 0 else 0)
    
    if not unsigned_list:
        st.success("🎉 모든 교직원의 참여가 완료되었습니다!")
    else:
        selected_name = st.selectbox(f"성함을 선택하세요 (미완료 {len(unsigned_list)}명)", ["선택하세요"] + unsigned_list)
        if selected_name != "선택하세요":
            st.markdown("---")
            agree_status = st.radio("본 사업 추진에 동의하십니까?", ["동의", "미동의"], horizontal=True)
            canvas_result = st_canvas(fill_color="rgba(255, 255, 255, 0)", stroke_width=4, stroke_color="#000000", background_color="#FFFFFF", height=180, width=320, drawing_mode="freedraw", key="sig_pad")

            if st.button("최종 제출하기", use_container_width=True):
                if canvas_result.image_data is not None:
                    with st.spinner("제출 중..."):
                        rgba_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                        white_bg = Image.new("RGB", rgba_img.size, (255, 255, 255))
                        white_bg.paste(rgba_img, mask=rgba_img.split()[3])
                        img = white_bg; img.thumbnail((250, 120))
                        buf = io.BytesIO(); img.save(buf, format="PNG")
                        sig_b64 = base64.b64encode(buf.getvalue()).decode()

                        row_idx = df_all.index[df_all['성함'] == selected_name].tolist()[0]
                        df_all.at[row_idx, '동의여부'] = agree_status
                        df_all.at[row_idx, '서명'] = f"'{sig_b64}"
                        conn.update(worksheet="동의서양식", data=df_all)
                        
                        st.cache_data.clear() # 제출 후 캐시 삭제하여 목록 갱신
                        st.balloons(); st.success("제출 완료!"); st.rerun()
                else: st.error("서명을 완료해 주세요.")

# --- 탭 2: 관리자 메뉴 ---
with menu_tab2:
    if 'admin_authenticated' not in st.session_state: st.session_state.admin_authenticated = False
    if not st.session_state.admin_authenticated:
        pw = st.text_input("관리자 비밀번호", type="password")
        if st.button("인증"):
            if pw == st.secrets.get("auth", {}).get("admin_password", "1234"):
                st.session_state.admin_authenticated = True; st.rerun()
    else:
        if st.button("🖼️ 엑셀 생성 및 다운로드"):
            excel_data = generate_excel_with_images(df_all)
            st.download_button(label="💾 엑셀 저장", data=excel_data, file_name=f"동의서_결과_{pd.Timestamp.now().strftime('%m%d_%H%M')}.xlsx")
        if st.button("🔄 강제 새로고침"):
            st.cache_data.clear(); st.rerun()
        st.dataframe(df_all[['번호', '성함', '동의여부']], hide_index=True)
