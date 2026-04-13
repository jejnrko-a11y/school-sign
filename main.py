import streamlit as st
from streamlit_drawable_canvas import st_canvas
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from PIL import Image
import io
import base64
import xlsxwriter

st.set_page_config(page_title="재구조화 동의서 관리 시스템", layout="wide")
conn = st.connection("gsheets", type=GSheetsConnection)

# --- [함수] 구글 시트 저장 로직 (열 위치 강제 고정) ---
def save_to_gsheet(df_to_save):
    """데이터가 항상 C, D, K열에 위치하도록 11개 열을 강제로 만들어 저장합니다."""
    # 1. 11개의 빈 열(A~K)을 가진 데이터프레임 생성
    # 구글 시트의 헤더 위치와 이름을 정확히 매칭시킴
    final_df = pd.DataFrame(columns=[f"col_{i}" for i in range(11)]) 
    
    # 2. 지정된 위치에 데이터 매핑 (0:A, 1:B, 2:C, 3:D ... 10:K)
    final_df.iloc[:, 2] = df_to_save['성함']
    final_df.iloc[:, 3] = df_to_save['서명']
    final_df.iloc[:, 10] = df_to_save['동의여부']
    
    # 3. 헤더 이름 수정 (구글 시트에 표시될 이름)
    # A, B열과 중간 열들은 빈 이름으로, C, D, K는 지정된 이름으로
    new_cols = [''] * 11
    new_cols[2], new_cols[3], new_cols[10] = '성함', '서명', '동의여부'
    final_df.columns = new_cols
    
    # 4. 업데이트 (이때 A1부터 덮어씌워지므로 위치가 고정됨)
    conn.update(worksheet="동의서출력", data=final_df)

# --- [함수] 이미지 포함 엑셀 생성 (C, D, K열 고정) ---
def generate_excel_with_layout(df_master):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('최종동의서')

    format_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    format_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2'})

    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 30)
    worksheet.set_column('K:K', 15)

    worksheet.write('C1', '성함', format_header)
    worksheet.write('D1', '서명', format_header)
    worksheet.write('K1', '동의여부', format_header)

    for idx, row in df_master.iterrows():
        row_num = idx + 1
        worksheet.set_row(row_num, 65)

        # 데이터 안전하게 추출
        name_val = str(row.get('성함', ''))
        agree_val = str(row.get('동의여부', ''))
        sig_data = str(row.get('서명', '')).strip()

        worksheet.write(row_num, 2, name_val, format_cell) # C열
        worksheet.write(row_num, 10, agree_val, format_cell) # K열

        if sig_data.startswith("'"): sig_data = sig_data[1:]
        if sig_data and sig_data != "" and sig_data != "nan":
            try:
                img_data = base64.b64decode(sig_data)
                img_io = io.BytesIO(img_data)
                worksheet.insert_image(row_num, 3, 'sig.png', {'image_data': img_io, 'x_scale': 0.55, 'y_scale': 0.55, 'x_offset': 10, 'y_offset': 5})
                worksheet.write(row_num, 3, "", format_cell)
            except: worksheet.write(row_num, 3, "", format_cell)
        else:
            worksheet.write(row_num, 3, "", format_cell)

    workbook.close()
    return output.getvalue()

# --- 데이터 로드 및 전처리 ---
try:
    df_unsigned = conn.read(worksheet="동의서양식", ttl=0).fillna("")
    
    # 동의서출력 로드 시, 이름 기반으로 컬럼을 찾음
    raw_master = conn.read(worksheet="동의서출력", ttl=0).fillna("")
    
    # [핵심] 읽어올 때 열이 밀려있어도 '성함', '서명', '동의여부'라는 이름만 있으면 찾아냄
    # 필요한 컬럼만 추출하여 내부용 df_master 정의
    needed_cols = ['성함', '서명', '동의여부']
    df_master = raw_master[needed_cols].copy()
    
except Exception as e:
    st.error(f"데이터 로드 에러: {e}")
    st.stop()

# --- 메인 화면 ---
st.title("2026 경기기공 재구조화 동의서")
tab1, tab2 = st.tabs(["✍️ 서명하기", "📊 관리자 전용"])

with tab1:
    unsigned_list = [n for n in df_unsigned['성함'].tolist() if str(n).strip() != ""]
    if not unsigned_list:
        st.success("🎉 모든 대상자의 서명이 완료되었습니다!")
    else:
        selected_name = st.selectbox(f"성함 선택 (남은 인원: {len(unsigned_list)}명)", ["선택하세요"] + unsigned_list)
        if selected_name != "선택하세요":
            agree_status = st.radio("본 사업 추진에 동의하십니까?", ["동의", "미동의"], horizontal=True)
            canvas_result = st_canvas(height=180, width=320, stroke_width=4, stroke_color="#000000", background_color="#FFFFFF", key="canvas")

            if st.button("제출하기", use_container_width=True):
                if canvas_result.image_data is not None:
                    # 이미지 처리
                    rgba_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                    white_bg = Image.new("RGB", rgba_img.size, (255, 255, 255))
                    white_bg.paste(rgba_img, mask=rgba_img.split()[3])
                    white_bg.thumbnail((250, 120))
                    buf = io.BytesIO(); white_bg.save(buf, format="PNG")
                    sig_b64 = base64.b64encode(buf.getvalue()).decode()

                    # 데이터 업데이트
                    if selected_name in df_master['성함'].values:
                        m_idx = df_master.index[df_master['성함'] == selected_name].tolist()[0]
                        df_master.at[m_idx, '동의여부'] = agree_status
                        df_master.at[m_idx, '서명'] = f"'{sig_b64}"
                        
                        # [핵심] 위치 고정 저장 함수 호출
                        save_to_gsheet(df_master)
                        
                        # 동의서양식에서 제거
                        df_unsigned = df_unsigned[df_unsigned['성함'] != selected_name]
                        conn.update(worksheet="동의서양식", data=df_unsigned)
                        
                        st.balloons(); st.success("제출 완료!"); st.rerun()

with tab2:
    if 'auth' not in st.session_state: st.session_state.auth = False
    if not st.session_state.auth:
        pw = st.text_input("비밀번호", type="password")
        if st.button("인증"):
            if pw == st.secrets.get("auth", {}).get("admin_password", "1234"):
                st.session_state.auth = True; st.rerun()
    else:
        if st.button("🖼️ 최종 엑셀 파일 생성", use_container_width=True):
            excel_data = generate_excel_with_layout(df_master)
            st.download_button(label="💾 엑셀 다운로드", data=excel_data, file_name=f"최종_동의서_{pd.Timestamp.now().strftime('%m%d')}.xlsx")
