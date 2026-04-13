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

# --- [함수] 엑셀 생성 로직 (이미지 포함) ---
def generate_excel_with_images(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_display = df[['번호', '성함', '서명']] # 출력할 컬럼만 선택
        df_display.to_excel(writer, index=False, sheet_name='동의서_결과')
        workbook  = writer.book
        worksheet = writer.sheets['동의서_결과']

        format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.set_column('B:B', 15) 
        worksheet.set_column('C:C', 25) 
        
        for idx, row in df.iterrows():
            row_num = idx + 1
            worksheet.set_row(row_num, 60)
            
            sig_data = str(row['서명'])
            if sig_data.startswith("'"): sig_data = sig_data[1:]
            
            if sig_data and sig_data != "":
                try:
                    img_data = base64.b64decode(sig_data)
                    img_io = io.BytesIO(img_data)
                    worksheet.insert_image(row_num, 2, 'signature.png', 
                                         {'image_data': img_io, 'x_scale': 0.5, 'y_scale': 0.5, 'x_offset': 5, 'y_offset': 5})
                    worksheet.write(row_num, 2, "", format_center) # 글자 데이터는 숨김
                except: pass
    return output.getvalue()

# --- 화면 구성 ---
st.title("2026 경기기공 재구조화 동의서")

menu_tab1, menu_tab2 = st.tabs(["✍️ 동의서 서명", "📊 관리자 메뉴"])

# --- 탭 1: 서명하기 ---
with menu_tab1:
    with st.container(border=True):
        st.subheader("📢 재구조화 사업 안내")
        st.info("본 사업은 학과 개편 및 교육환경 개선을 위해 추진됩니다. 내용을 확인하시고 동의 여부를 서명으로 확인해 주시기 바랍니다.")

    try:
        # 데이터 실시간 로드
        df_all = conn.read(worksheet="동의서양식", ttl=0).fillna("")
        
        # [핵심 로직] 서명이 비어있는 사람만 필터링
        unsigned_df = df_all[df_all['서명'] == ""]
        unsigned_list = unsigned_df['성함'].tolist()
        
        total_count = len(df_all)
        signed_count = total_count - len(unsigned_df)
    except:
        st.error("시트를 불러올 수 없습니다. '동의서양식' 탭을 확인해 주세요.")
        st.stop()

    # 진행률 표시
    st.progress(signed_count / total_count)
    st.caption(f"현재 참여 현황: {signed_count}명 완료 / 총 {total_count}명")

    if not unsigned_list:
        st.success("🎉 모든 교직원의 서명이 완료되었습니다!")
    else:
        selected_name = st.selectbox(f"성함을 선택하세요 (미완료자 {len(unsigned_list)}명)", ["선택하세요"] + unsigned_list)

        if selected_name != "선택하세요":
            st.markdown(f"### 🖋️ {selected_name} 선생님의 서명")
            
            # 서명 패드 (배경색을 흰색으로 설정)
            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)", stroke_width=4,
                stroke_color="#000000", background_color="#FFFFFF",
                height=180, width=320, drawing_mode="freedraw", key="sig_pad"
            )

            if st.button("동의 및 서명 제출", use_container_width=True):
                if canvas_result.image_data is not None:
                    with st.spinner("처리 중..."):
                        # 1. 서명 이미지 처리 (배경 흰색 입히기)
                        rgba_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                        # 투명한 배경을 흰색으로 채우는 과정
                        white_bg = Image.new("RGB", rgba_img.size, (255, 255, 255))
                        white_bg.paste(rgba_img, mask=rgba_img.split()[3]) # 3번 인덱스가 알파(투명도) 채널
                        
                        img = white_bg
                        img.thumbnail((250, 120))
                        buf = io.BytesIO()
                        img.save(buf, format="PNG")
                        sig_b64 = base64.b64encode(buf.getvalue()).decode()

                        # 2. 데이터 업데이트
                        row_idx = df_all.index[df_all['성함'] == selected_name].tolist()[0]
                        df_all.at[row_idx, '서명'] = f"'{sig_b64}"
                        conn.update(worksheet="동의서양식", data=df_all)
                        
                        st.balloons()
                        st.success(f"{selected_name} 선생님, 제출이 완료되었습니다. 감사합니다!")
                        st.rerun() # 목록 갱신을 위해 재실행
                else:
                    st.error("서명을 그려주세요.")

# --- 탭 2: 관리자 메뉴 ---
with menu_tab2:
    st.subheader("📥 행정용 결과물 생성")
    pw = st.text_input("관리자 비밀번호", type="password")
    if pw == st.secrets.get("auth", {}).get("admin_password", "1234"):
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🖼️ 이미지 포함 엑셀 생성", use_container_width=True):
                with st.spinner("엑셀 제작 중..."):
                    excel_data = generate_excel_with_images(df_all)
                    st.download_button(
                        label="💾 엑셀 다운로드", data=excel_data,
                        file_name=f"재구조화_동의서_결과_{pd.Timestamp.now().strftime('%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        with col2:
            if st.button("🔄 전체 현황 새로고침", use_container_width=True):
                st.cache_data.clear()
                st.rerun()
        
        st.divider()
        st.write("### 📋 현재 명단 기록 상황")
        st.dataframe(df_all[['번호', '성함', '서명']].assign(상태=df_all['서명'].apply(lambda x: "✅ 완료" if x else "⏳ 미완료")), hide_index=True)
