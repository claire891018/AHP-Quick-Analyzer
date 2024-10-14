import streamlit as st
import requests

st.set_page_config(page_title="AHP 分析工具", page_icon="https://api.dicebear.com/9.x/thumbs/svg?seed=Liam" , layout="wide")

API_URL = "http://localhost:8000"  # FastAPI 伺服器位址

home_title = 'AHP 分析工具'
st.markdown(f"""# {home_title} <span style=color:#2E9BF5><font size=5>Beta</font></span>""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file is not None:
    if st.button("取得工作表列表"):
        response = requests.post(
            f"{API_URL}/list-sheets/", 
            files={"file": uploaded_file}
        )

        if response.status_code == 200:
            sheets = response.json().get("sheets", [])
            st.session_state["sheets"] = sheets
            # sheet_name = st.selectbox("選擇工作表", sheets)
        else:
            st.error(f"無法取得工作表列表：{response.text}")
            
    if "sheets" in st.session_state:
        sheet_name = st.selectbox("選擇工作表", st.session_state["sheets"])
        st.session_state["sheet_name"] = sheet_name  
        
    start_row = st.number_input("開始行數", min_value=1, step=1)
    end_row = st.number_input("結束行數", min_value=1, step=1)

    if st.button("開始分析"):
        if start_row > end_row:
            st.error("開始行數不可大於結束行數")
        else:
            with st.spinner("正在分析中..."):
                response = requests.post(
                    f"{API_URL}/analyze/",
                    params={
                        "sheet_name": st.session_state.get("sheet_name"),
                        "start_row": start_row,
                        "end_row": end_row,
                    },
                    files={"file": uploaded_file},
                )

            if response.status_code == 200:
                st.success("分析完成！")
                with open("analysis_results.xlsx", "wb") as f:
                    f.write(response.content)
                st.download_button(
                    label="下載分析結果",
                    data=response.content,
                    file_name="analysis_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error(f"分析失敗：{response.text}")