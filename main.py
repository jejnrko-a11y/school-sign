import streamlit as st
from streamlit_drawable_canvas import st_canvas
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from PIL import Image
import io
import base64
import xlsxwriter
from datetime import datetime

# --- 1. 앱 기본 설정 ---
st.set_page_config(page_title="재구조화 동의서 관리 시스템", layout="wide")

# 구글 시트 연결
conn = st.connection("gsheets", type=GSheetsConnection)

# --- 2. 핵심 함수 정의 ---

def save_to_gsheet(df_master):
    """
    데이터가 항상 구글 시트의 C(2), D(3), K(10) 열에 위치하도록 
    11개의 열을 강제로 구성하여 저장합니다.
    """
    try:
        # 11개의 빈 열(0~10)을 가진 데이터프레임 생성
        rows = len(df_master)
        # 컬럼명을 숫자로 임시 설정하여 위치 고정
        final_df = pd.DataFrame(index=range(rows), columns=range(11))
        
        # 지정된 인덱스에 데이터 매핑 (C:2, D:3, K:10)
        final_df.iloc[:, 2] = df_master['성함'].values
        final_df.iloc[:, 3] = df_master['서명'].values
        final_df.iloc[:, 10] = df_master['동의여부'].values
        
        # 구글 시트 헤더 이름 설정 (빈 열은 이름 없음)
        header_names = [''] * 11
        header_names[2], header_names[3], header_names[10] = '성함', '서명', '동의여부'
        final_df.columns = header_names
        
        # 시트에 업데이트 (A1부터 덮어쓰기 되어 위치가 고정됨)
        conn.update(worksheet="동의서출력", data=final_df)
        return True
    except Exception as e:
        st.error(f"시트 저장 중 오류 발생: {e}")
        return False

def generate_excel_with_layout(df_master):
    """
    이미지를 포함하여 교육청 양식(C, D, K열)에 맞는 엑셀 파일을 생성합니다.
    """
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('최종동의서')

    # 서식 설정
    format_cell = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
    format_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2'})

    # 열 너비 설정
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 30)
    worksheet.set_column('K:K', 15)

    # 헤더 작성
    worksheet.write('C1', '성함', format_header)
    worksheet.write('D1', '서명', format_header)
    worksheet.write('K1', '동의여부', format_header)

    # 데이터 작성
    for idx, row in df_master.iterrows():
        row_num = idx + 1
        worksheet.set_row(row_num, 65) # 행 높이

        name_val = str(row.get('성함', ''))
        agree_val = str(row.get('동의여부', ''))
        sig_data = str(row.get('서명', '')).strip()

        worksheet.write(row_num, 2, name_val, format_cell)   # C열
        worksheet.write(row_num, 10, agree_val, format_cell) # K열

        if sig_data.startswith("'"): sig_data = sig_data[1:]
        
        if sig_data and sig_data != "" and sig_data != "nan":
            try:
                img_data = base64.b64decode(sig_data)
                img_io = io.BytesIO(img_data)
                # D열에 이미지 삽입
                worksheet.insert_image(row_num, 3, 'sig.png', 
                                     {'image_data': img_io, 'x_scale': 0.55, 'y_scale': 0.55, 'x_offset': 10, 'y_offset': 5})
                worksheet.write(row_num, 3, "", format_cell)
            except:
                worksheet.write(row_num, 3, "이미지오류", format_cell)
        else:
            worksheet.write(row_num, 3, "", format_cell)

    workbook.close()
    return output.getvalue()

# --- 3. 데이터 로드 및 전처리 ---

try:
    # 시트 1: 미참여자 명단 (A1:번호, B1:성함)
    df_unsigned = conn.read(worksheet="동의서양식", ttl=0).fillna("")
    
    # 시트 2: 마스터 출력 시트 (C1:성함, D1:서명, K1:동의여부)
    raw_master = conn.read(worksheet="동의서출력", ttl=0).fillna("")
    
    # 필요한 컬럼만 추출 (열이 밀려있어도 이름으로 찾아냄)
    if '성함' in raw_master.columns and '서명' in raw_master.columns and '동의여부' in raw_master.columns:
        df_master = raw_master[['성함', '서명', '동의여부']].copy()
    else:
        st.error("보안/설정 오류: '동의서출력' 시트의 헤더(성함, 서명, 동의여부)를 찾을 수 없습니다.")
        st.stop()
        
except Exception as e:
    st.error(f"데이터 로드 에러: {e}")
    st.stop()

# --- 4. 메인 UI 화면 ---

st.title("🏫 2026 경기기계공업고등학교 재구조화 동의서")

tab1, tab2 = st.tabs(["✍️ 동의서 서명하기", "📊 관리자 전용 메뉴"])

# --- 탭 1: 서명 섹션 ---
with tab1:
    st.info("안내 문구를 숙지하신 후 본인의 성함을 선택하여 의견 표시와 서명을 진행해 주세요.")
    
    # 미참여자 명단 추출 (동의서양식 기준)
    if '성함' in df_unsigned.columns:
        unsigned_list = [n for n in df_unsigned['성함'].tolist() if str(n).strip() != ""]
    else:
        unsigned_list = []

    if not unsigned_list:
        st.success("🎉 현재 모든 대상자의 서명이 완료되었습니다!")
    else:
        selected_name = st.selectbox(f"본인의 성함을 선택하세요 (남은 인원: {len(unsigned_list)}명)", ["선택하세요"] + unsigned_list)

        if selected_name != "선택하세요":
            st.markdown("---")
            st.subheader(f"📍 {selected_name} 선생님의 의견 선택")
            agree_status = st.radio("본 사업 추진에 동의하십니까?", ["동의", "미동의"], horizontal=True, key="agree_radio")
            
            st.markdown("### 🖋️ 서명 (흰색 배경에 정자로 서명해 주세요)")
            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)", stroke_width=4,
                stroke_color="#000000", background_color="#FFFFFF",
                height=180, width=320, drawing_mode="freedraw", key="canvas_pad"
            )

            if st.button("동의서 최종 제출", use_container_width=True):
                if canvas_result.image_data is not None:
                    with st.spinner("서명을 처리하고 기록 중입니다..."):
                        # 1. 이미지 처리 (투명 배경 -> 흰색 배경)
                        rgba_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                        white_bg = Image.new("RGB", rgba_img.size, (255, 255, 255))
                        white_bg.paste(rgba_img, mask=rgba_img.split()[3])
                        white_bg.thumbnail((250, 120))
                        buf = io.BytesIO()
                        white_bg.save(buf, format="PNG")
                        sig_b64 = base64.b64encode(buf.getvalue()).decode()

                        # 2. 마스터 시트(동의서출력) 데이터 업데이트
                        if selected_name in df_master['성함'].values:
                            m_idx = df_master.index[df_master['성함'] == selected_name].tolist()[0]
                            df_master.at[m_idx, '동의여부'] = agree_status
                            df_master.at[m_idx, '서명'] = f"'{sig_b64}"
                            
                            # 위치 고정 저장 함수 실행
                            if save_to_gsheet(df_master):
                                # 3. 미참여자 시트(동의서양식)에서 이름 제거
                                df_unsigned = df_unsigned[df_unsigned['성함'] != selected_name]
                                conn.update(worksheet="동의서양식", data=df_unsigned)
                                
                                st.balloons()
                                st.success(f"{selected_name} 선생님, 소중한 의견 감사합니다. 제출되었습니다.")
                                st.rerun()
                        else:
                            st.error(f"'동의서출력' 시트에 {selected_name} 선생님의 정보가 없습니다. 관리자에게 문의하세요.")
                else:
                    st.error("서명란에 서명을 입력해 주세요.")

# --- 탭 2: 관리자 섹션 ---
with tab2:
    st.subheader("🔐 행정 관리자 메뉴")
    
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        with st.form("admin_login"):
            pw = st.text_input("관리자 비밀번호", type="password")
            if st.form_submit_button("관리자 인증하기"):
                if pw == st.secrets.get("auth", {}).get("admin_password", "1234"):
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("비밀번호가 올바르지 않습니다.")
    else:
        st.success("✅ 관리자 권한으로 인증되었습니다.")
        
        # 통계 표시
        total_total = len(df_master)
        done_count = len(df_master[df_master['서명'] != ""])
        st.metric("현재 서명 진행률", f"{done_count} / {total_total} ({(done_count/total_total*100):.1f}%)")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("🖼️ 이미지 포함 엑셀 파일 생성", use_container_width=True):
                with st.spinner("엑셀 파일을 생성 중입니다..."):
                    excel_data = generate_excel_with_layout(df_master)
                    st.download_button(
                        label="💾 생성된 엑셀 다운로드",
                        data=excel_data,
                        file_name=f"최종_재구조화_동의서_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        with col2:
            if st.button("🔄 시트 데이터 새로고침", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
        
        st.divider()
        st.write("### 📋 전체 마스터 현황 (동의서출력 기준)")
        
        # 관리자용 가독성 데이터프레임
        view_df = df_master.copy()
        view_df['상태'] = view_df['서명'].apply(lambda x: "✅ 완료" if len(str(x)) > 50 else "⏳ 미참여")
        st.dataframe(view_df[['성함', '동의여부', '상태']], hide_index=True, use_container_width=True)

        if st.button("🔒 관리자 로그아웃"):
            st.session_state.authenticated = False
            st.rerun()
