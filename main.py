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

# --- [함수] 엑셀 생성 로직 (동의여부 및 이미지 포함) ---
def generate_excel_with_images(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 출력할 컬럼 순서 조정
        df_display = df[['번호', '성함', '동의여부', '서명']]
        df_display.to_excel(writer, index=False, sheet_name='동의서_결과')
        workbook  = writer.book
        worksheet = writer.sheets['동의서_결과']

        format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        worksheet.set_column('B:B', 12) # 성함
        worksheet.set_column('C:C', 12) # 동의여부
        worksheet.set_column('D:D', 25) # 서명
        
        for idx, row in df.iterrows():
            row_num = idx + 1
            worksheet.set_row(row_num, 60)
            
            sig_data = str(row['서명'])
            if sig_data.startswith("'"): sig_data = sig_data[1:]
            
            if sig_data and sig_data != "":
                try:
                    img_data = base64.b64decode(sig_data)
                    img_io = io.BytesIO(img_data)
                    # 서명 이미지를 D열(index 3)에 삽입
                    worksheet.insert_image(row_num, 3, 'signature.png', 
                                         {'image_data': img_io, 'x_scale': 0.5, 'y_scale': 0.5, 'x_offset': 5, 'y_offset': 5})
                    worksheet.write(row_num, 3, "", format_center) 
                except: pass
    return output.getvalue()

# --- 화면 구성 ---
st.title("2026 경기기공 재구조화 동의서")

menu_tab1, menu_tab2 = st.tabs(["✍️ 동의서 서명", "📊 관리자 메뉴"])

# --- 탭 1: 서명하기 ---
with menu_tab1:
    with st.container(border=True):
        st.subheader("📢 재구조화 사업 안내")
        st.info("본 사업에 대한 안내를 충분히 숙지하셨습니까? 성함을 선택하신 후 본인의 의사를 표시하고 서명해 주시기 바랍니다.")

    try:
        df_all = conn.read(worksheet="동의서양식", ttl=0).fillna("")
        unsigned_df = df_all[df_all['서명'] == ""]
        unsigned_list = unsigned_df['성함'].tolist()
        
        total_count = len(df_all)
        signed_count = total_count - len(unsigned_df)
    except:
        st.error("시트를 불러올 수 없습니다. '동의서양식' 탭의 컬럼(번호, 성함, 동의여부, 서명)을 확인해 주세요.")
        st.stop()

    st.progress(signed_count / total_count if total_count > 0 else 0)
    st.caption(f"참여 현황: {signed_count}명 완료 / 총 {total_count}명")

    if not unsigned_list:
        st.success("🎉 모든 교직원의 참여가 완료되었습니다!")
    else:
        selected_name = st.selectbox(f"성함을 선택하세요 (미완료 {len(unsigned_list)}명)", ["선택하세요"] + unsigned_list)

        if selected_name != "선택하세요":
            # [추가] 동의/미동의 선택
            st.markdown("---")
            st.markdown(f"### 📍 {selected_name} 선생님의 의견 선택")
            agree_status = st.radio("본 사업 추진에 동의하십니까?", ["동의", "미동의"], index=0, horizontal=True)
            
            status_color = "green" if agree_status == "동의" else "red"
            st.markdown(f"의견: :{status_color}[**{agree_status}**]")

            st.markdown("### 🖋️ 서명")
            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 0)", stroke_width=4,
                stroke_color="#000000", background_color="#FFFFFF",
                height=180, width=320, drawing_mode="freedraw", key="sig_pad"
            )

            if st.button("최종 제출하기", use_container_width=True):
                if canvas_result.image_data is not None:
                    with st.spinner("제출 중..."):
                        # 서명 이미지 처리
                        rgba_img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                        white_bg = Image.new("RGB", rgba_img.size, (255, 255, 255))
                        white_bg.paste(rgba_img, mask=rgba_img.split()[3])
                        
                        img = white_bg
                        img.thumbnail((250, 120))
                        buf = io.BytesIO()
                        img.save(buf, format="PNG")
                        sig_b64 = base64.b64encode(buf.getvalue()).decode()

                        # 데이터 업데이트 (동의여부 포함)
                        row_idx = df_all.index[df_all['성함'] == selected_name].tolist()[0]
                        df_all.at[row_idx, '동의여부'] = agree_status
                        df_all.at[row_idx, '서명'] = f"'{sig_b64}"
                        conn.update(worksheet="동의서양식", data=df_all)
                        
                        st.balloons()
                        st.success("소중한 의견 감사합니다. 제출되었습니다.")
                        st.rerun()
                else:
                    st.error("서명을 완료해 주세요.")

# --- 탭 2: 관리자 메뉴 ---
with menu_tab2:
    st.subheader("🔐 관리자 전용 영역")
    if 'admin_authenticated' not in st.session_state:
        st.session_state.admin_authenticated = False

    if not st.session_state.admin_authenticated:
        pw = st.text_input("관리자 비밀번호", type="password")
        if st.button("관리자 모드 활성화", use_container_width=True):
            if pw == st.secrets.get("auth", {}).get("admin_password", "1234"):
                st.session_state.admin_authenticated = True
                st.rerun()
            else:
                st.error("❌ 비밀번호가 틀렸습니다.")
    else:
        st.info("✅ 관리자 권한 접속 중")
        
        # [추가] 찬반 통계 보기
        agree_counts = df_all['동의여부'].value_counts()
        c1, c2, c3 = st.columns(3)
        c1.metric("전체 참여", f"{signed_count}명")
        c2.metric("동의", f"{agree_counts.get('동의', 0)}명")
        c3.metric("미동의", f"{agree_counts.get('미동의', 0)}명")

        if st.button("🖼️ 이미지 포함 엑셀 생성", use_container_width=True):
            with st.spinner("엑셀 제작 중..."):
                excel_data = generate_excel_with_images(df_all)
                st.download_button(
                    label="💾 엑셀 다운로드", data=excel_data,
                    file_name=f"재구조화_동의서_결과_{pd.Timestamp.now().strftime('%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        if st.button("🔄 전체 현황 새로고침", use_container_width=True):
            st.cache_data.clear(); st.rerun()

        st.divider()
        st.write("### 📋 세부 현황")
        view_df = df_all[['번호', '성함', '동의여부']].copy()
        view_df['상태'] = df_all['서명'].apply(lambda x: "✅ 완료" if x else "⏳ 미완료")
        st.dataframe(view_df, hide_index=True, use_container_width=True)

        if st.button("🔐 관리자 모드 종료"):
            st.session_state.admin_authenticated = False
            st.rerun()
