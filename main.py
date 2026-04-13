import streamlit as st
from streamlit_drawable_canvas import st_canvas
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from PIL import Image
import io
import base64
import xlsxwriter

# --- 앱 설정 ---
st.set_page_config(page_title="재구조화 동의서 통합 관리", layout="centered")

# 구글 시트 연결
conn = st.connection("gsheets", type=GSheetsConnection)

# --- [함수] 이미지 포함 엑셀 생성 (동의서출력 시트 기준) ---
def generate_excel_from_master(df_master):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # '동의서출력' 시트 데이터를 엑셀로 변환
        df_master.to_excel(writer, index=False, sheet_name='최종동의서')
        workbook  = writer.book
        worksheet = writer.sheets['최종동의서']

        format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.set_column('B:B', 12) # 성함
        worksheet.set_column('C:C', 10) # 동의여부
        worksheet.set_column('D:D', 25) # 서명
        
        for idx, row in df_master.iterrows():
            row_num = idx + 1
            worksheet.set_row(row_num, 60) # 행 높이 고정
            
            sig_data = str(row['서명']).strip()
            if sig_data.startswith("'"): sig_data = sig_data[1:]
            
            # 데이터가 있으면 이미지 삽입
            if sig_data and sig_data != "" and sig_data != "nan":
                try:
                    img_data = base64.b64decode(sig_data)
                    img_io = io.BytesIO(img_data)
                    # D열(index 3)에 이미지 삽입
                    worksheet.insert_image(row_num, 3, 'signature.png', 
                                         {'image_data': img_io, 'x_scale': 0.5, 'y_scale': 0.5, 'x_offset': 5, 'y_offset': 5})
                    worksheet.write(row_num, 3, "", format_center) 
                except: pass
    return output.getvalue()

# --- 데이터 로드 ---
try:
    # 1. 미참여자 명단 (동의서양식)
    df_unsigned = conn.read(worksheet="동의서양식", ttl=0).fillna("")
    # 2. 마스터 시트 (동의서출력)
    df_master = conn.read(worksheet="동의서출력", ttl=0).fillna("")
except:
    st.error("시트 로드 실패! '동의서양식'과 '동의서출력' 시트가 있는지 확인해 주세요.")
    st.stop()

# --- 화면 구성 ---
st.title("2026 경기기공 재구조화 동의서")

tab1, tab2 = st.tabs(["✍️ 동의서 서명", "📊 관리자 전용"])

# --- 탭 1: 서명하기 (동의서양식 기반) ---
with tab1:
    st.info("재구조화 사업 안내를 숙지하신 후, 본인의 성함을 선택하여 서명해 주시기 바랍니다.")
    
    unsigned_list = df_unsigned['성함'].tolist()
    
    if not unsigned_list:
        st.success("🎉 모든 대상자의 서명이 완료되었습니다!")
    else:
        selected_name = st.selectbox(f"성함을 선택하세요 (대상: {len(unsigned_list)}명)", ["선택하세요"] + unsigned_list)

        if selected_name != "선택하세요":
            st.markdown(f"### 📍 {selected_name} 선생님의 의견")
            agree_status = st.radio("본 사업 추진에 동의하십니까?", ["동의", "미동의"], horizontal=True)

            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)", stroke_width=4,
                stroke_color="#000000", background_color="#FFFFFF",
                height=180, width=320, drawing_mode="freedraw", key="canvas"
            )

            if st.button("동의서 제출하기", use_container_width=True):
                if canvas_result.image_data is not None:
                    with st.spinner("최종본에 기록 중..."):
                        # 1. 서명 이미지 처리 (배경 흰색)
                        rgba_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                        white_bg = Image.new("RGB", rgba_img.size, (255, 255, 255))
                        white_bg.paste(rgba_img, mask=rgba_img.split()[3])
                        white_bg.thumbnail((250, 120))
                        buf = io.BytesIO()
                        white_bg.save(buf, format="PNG")
                        sig_b64 = base64.b64encode(buf.getvalue()).decode()

                        # 2. 마스터 시트(동의서출력) 업데이트
                        if selected_name in df_master['성함'].values:
                            master_idx = df_master.index[df_master['성함'] == selected_name].tolist()[0]
                            df_master.at[master_idx, '동의여부'] = agree_status
                            df_master.at[master_idx, '서명'] = f"'{sig_b64}"
                            conn.update(worksheet="동의서출력", data=df_master)
                            
                            # 3. 미참여자 시트(동의서양식)에서 제거
                            df_unsigned = df_unsigned[df_unsigned['성함'] != selected_name]
                            conn.update(worksheet="동의서양식", data=df_unsigned)
                            
                            st.balloons()
                            st.success(f"{selected_name} 선생님, 제출되었습니다.")
                            st.rerun()
                        else:
                            st.error("오류: 마스터 시트(동의서출력)에서 이름을 찾을 수 없습니다.")

# --- 탭 2: 관리자 전용 (최종본 출력) ---
with tab2:
    if 'auth' not in st.session_state: st.session_state.auth = False
    
    if not st.session_state.auth:
        pw = st.text_input("관리자 비밀번호", type="password")
        if st.button("인증하기"):
            if pw == st.secrets.get("auth", {}).get("admin_password", "1234"):
                st.session_state.auth = True
                st.rerun()
    else:
        st.subheader("🏁 최종 동의서 출력")
        st.write("'동의서출력' 시트의 모든 데이터를 기반으로 엑셀을 생성합니다.")
        
        if st.button("🖼️ 최종 엑셀 파일 생성", use_container_width=True):
            with st.spinner("이미지 복원 및 엑셀 제작 중..."):
                excel_data = generate_excel_from_master(df_master)
                st.download_button(
                    label="💾 통합 동의서 다운로드 (.xlsx)",
                    data=excel_data,
                    file_name=f"최종_재구조화_동의서_{pd.Timestamp.now().strftime('%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        st.divider()
        st.write("### 📋 전체 마스터 현황")
        # 동의서출력 시트의 참여 현황 요약
        view_df = df_master[['번호', '성함', '동의여부']].copy()
        view_df['진행상태'] = df_master['서명'].apply(lambda x: "✅ 완료" if (str(x).strip() != "" and str(x) != "nan") else "⏳ 미완료")
        st.dataframe(view_df, hide_index=True, use_container_width=True)
        
        if st.button("로그아웃"):
            st.session_state.auth = False
            st.rerun()
