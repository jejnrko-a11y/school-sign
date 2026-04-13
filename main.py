import streamlit as st
from streamlit_drawable_canvas import st_canvas
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from PIL import Image
import io
import base64
import xlsxwriter

# --- 앱 설정 ---
st.set_page_config(page_title="재구조화 동의서 관리 시스템", layout="wide")

# 구글 시트 연결
conn = st.connection("gsheets", type=GSheetsConnection)

# --- [함수] 이미지 포함 엑셀 생성 (C, D, K열 고정) ---
def generate_excel_with_layout(df_master):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('최종동의서')

    format_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    format_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2'})

    # 열 너비 설정
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 30)
    worksheet.set_column('K:K', 15)

    # 헤더 작성 (고정 위치)
    worksheet.write('C1', '성함', format_header)
    worksheet.write('D1', '서명', format_header)
    worksheet.write('K1', '동의여부', format_header)

    # 데이터 작성
    for idx, row in df_master.iterrows():
        row_num = idx + 1
        worksheet.set_row(row_num, 65)

        # 컬럼 이름으로 안전하게 데이터 가져오기 (위치 상관없음)
        name_val = str(row.get('성함', ''))
        agree_val = str(row.get('동의여부', ''))
        sig_data = str(row.get('서명', '')).strip()

        worksheet.write(row_num, 2, name_val, format_cell) # C열
        worksheet.write(row_num, 10, agree_status, format_cell) if 'agree_status' in locals() else worksheet.write(row_num, 10, agree_val, format_cell) # K열

        if sig_data.startswith("'"): sig_data = sig_data[1:]
        
        if sig_data and sig_data != "" and sig_data != "nan":
            try:
                img_data = base64.b64decode(sig_data)
                img_io = io.BytesIO(img_data)
                worksheet.insert_image(row_num, 3, 'sig.png', {'image_data': img_io, 'x_scale': 0.55, 'y_scale': 0.55, 'x_offset': 10, 'y_offset': 5})
                worksheet.write(row_num, 3, "", format_cell) # D열 테두리
            except: worksheet.write(row_num, 3, "", format_cell)
        else:
            worksheet.write(row_num, 3, "", format_cell)

    workbook.close()
    return output.getvalue()

# --- 데이터 로드 ---
try:
    # 동의서양식: A1, B1 기준 (보통 첫 행부터 읽음)
    df_unsigned = conn.read(worksheet="동의서양식", ttl=0).fillna("")
    
    # 동의서출력: C1, D1, K1 헤더를 자동으로 찾기 위해 시트 전체 로드
    df_master = conn.read(worksheet="동의서출력", ttl=0).fillna("")
    
    # [중요] 구글 시트에서 빈 열(A, B)이 있으면 'Unnamed'로 읽힐 수 있으므로 컬럼명 전처리
    # '성함'이라는 글자가 포함된 컬럼을 실제 '성함' 컬럼으로 인식하게 함
    df_master.columns = [c if not 'Unnamed' in str(c) else '' for c in df_master.columns]
except Exception as e:
    st.error(f"데이터 로드 에러: {e}")
    st.stop()

# --- 화면 구성 ---
st.title("2026 경기기공 재구조화 동의서")
tab1, tab2 = st.tabs(["✍️ 서명하기", "📊 관리자 전용"])

with tab1:
    st.info("성함을 선택하신 후 의견을 표시하고 서명해 주세요.")
    # 동의서양식(A1, B1)에서 성함 리스트 추출
    if '성함' in df_unsigned.columns:
        unsigned_list = [n for n in df_unsigned['성함'].tolist() if n.strip() != ""]
    else:
        st.error("'동의서양식' 시트에 '성함' 컬럼(B1)이 없습니다.")
        st.stop()
    
    if not unsigned_list:
        st.success("🎉 모든 대상자의 서명이 완료되었습니다!")
    else:
        selected_name = st.selectbox(f"성함 선택 (남은 인원: {len(unsigned_list)}명)", ["선택하세요"] + unsigned_list)

        if selected_name != "선택하세요":
            st.markdown(f"### 📍 {selected_name} 선생님의 의견")
            agree_status = st.radio("본 사업 추진에 동의하십니까?", ["동의", "미동의"], horizontal=True)

            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)", stroke_width=4,
                stroke_color="#000000", background_color="#FFFFFF",
                height=180, width=320, drawing_mode="freedraw", key="canvas"
            )

            if st.button("제출하기", use_container_width=True):
                if canvas_result.image_data is not None:
                    with st.spinner("기록 중..."):
                        # 이미지 처리
                        rgba_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                        white_bg = Image.new("RGB", rgba_img.size, (255, 255, 255))
                        white_bg.paste(rgba_img, mask=rgba_img.split()[3])
                        white_bg.thumbnail((250, 120))
                        buf = io.BytesIO()
                        white_bg.save(buf, format="PNG")
                        sig_b64 = base64.b64encode(buf.getvalue()).decode()

                        # 1. 동의서출력(마스터) 업데이트
                        if '성함' in df_master.columns:
                            if selected_name in df_master['성함'].values:
                                m_idx = df_master.index[df_master['성함'] == selected_name].tolist()[0]
                                df_master.at[m_idx, '동의여부'] = agree_status
                                df_master.at[m_idx, '서명'] = f"'{sig_b64}"
                                # 업데이트 시 컬럼 순서 유지를 위해 전체 데이터 전송
                                conn.update(worksheet="동의서출력", data=df_master)
                                
                                # 2. 동의서양식에서 제거
                                df_unsigned = df_unsigned[df_unsigned['성함'] != selected_name]
                                conn.update(worksheet="동의서양식", data=df_unsigned)
                                
                                st.balloons(); st.success("제출되었습니다."); st.rerun()
                            else:
                                st.error(f"'동의서출력' 시트에 {selected_name} 선생님이 없습니다.")
                        else:
                            st.error("'동의서출력' 시트 1행에서 '성함' 컬럼을 찾을 수 없습니다.")

with tab2:
    if 'auth' not in st.session_state: st.session_state.auth = False
    if not st.session_state.auth:
        pw = st.text_input("비밀번호", type="password")
        if st.button("인증"):
            if pw == st.secrets.get("auth", {}).get("admin_password", "1234"):
                st.session_state.auth = True; st.rerun()
    else:
        st.subheader("🏁 최종 결과 출력")
        if st.button("🖼️ 최종 엑셀 파일 생성", use_container_width=True):
            excel_data = generate_excel_with_layout(df_master)
            st.download_button(label="💾 엑셀 다운로드", data=excel_data, file_name=f"최종_동의서_{pd.Timestamp.now().strftime('%m%d')}.xlsx")
        
        st.divider()
        if '성함' in df_master.columns and '서명' in df_master.columns:
            view_df = df_master[['성함', '동의여부']].copy()
            view_df['상태'] = df_master['서명'].apply(lambda x: "✅ 완료" if len(str(x)) > 50 else "⏳ 미완료")
            st.dataframe(view_df, hide_index=True, use_container_width=True)
