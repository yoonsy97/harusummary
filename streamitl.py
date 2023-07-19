import streamlit as st
import pandas as pd
import os

def read_excel_files(file_path):
    dfs = []
    dfs_name = []
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        for sheet_name in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            dfs.append(df)
            dfs_name.append(sheet_name)
    except Exception as e:
        st.error(f"{file_path} 파일을 읽는 중 오류가 발생했습니다: {e}")
    return dfs_name, dfs

def search_excel_files(folder_path):
    excel_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(root, file)
                excel_files.append(file_path)
    return excel_files

# 파일명만 추출하는 함수
def get_excel_filename(file_path):
    file_name_with_extension = os.path.basename(file_path)
    file_name, extension = os.path.splitext(file_name_with_extension)
    return file_name

# Streamlit 앱 구성
def main():
   
    st.title("네이버 검색 기록")
    # 경로 설정
    folder_path = "./excel"

    excel_files = search_excel_files(folder_path)
    
    # 파일명만 추출하여 사용
    page = st.sidebar.selectbox("Go to", [get_excel_filename(file) for file in excel_files])

    if page:
        file_path = os.path.join(folder_path, f"{page}.xlsx")
        dfs_name, dfs = read_excel_files(file_path)

        if dfs:
            # 페이지 내용 출력
            st.header(f"{page} 일자 ")
            for df_name, df in zip(dfs_name, dfs):
                st.subheader(f"{df_name}")
                st.dataframe(df)
                print("df.columns : ", df.columns)
                # title 컬럼에 하이퍼링크 생성
                # if 'title' in df.columns:
                #     df['title'] = df.apply(lambda x: f"[{x['title']}]({x['link']})", axis=1)
                #     st.dataframe(df)

# Streamlit 앱 실행
if __name__ == "__main__":
    main()