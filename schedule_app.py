import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import json

st.set_page_config(page_title="시공 일정 관리", layout="wide")

# ---------------------------
# 1. 구글시트 연결
# ---------------------------
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]


creds_dict = st.secrets["gcp_service_account"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# 👉 시트 이름 (구글시트 이름과 동일하게!)
SHEET_NAME = "시공일정"

sheet = client.open(SHEET_NAME).sheet1


# ---------------------------
# 2. 데이터 불러오기
# ---------------------------
def load_data():
    data = sheet.get_all_records()
    return pd.DataFrame(data)


def save_data(df):
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.values.tolist())


df = load_data()

if df.empty:
    df = pd.DataFrame(columns=["날짜", "설치현장", "시공담당", "수량", "비고", "상태", "완료일"])


# ---------------------------
# 3. 상단 요약
# ---------------------------
st.title("📅 시공 일정 관리 프로그램")

total = len(df)
today = len(df[df["날짜"] == str(datetime.today().date())])
progress = len(df[df["상태"] == "진행중"])
done = len(df[df["상태"] == "완료"])
qty = df["수량"].sum() if not df.empty else 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("전체 일정", total)
col2.metric("오늘 일정", today)
col3.metric("진행중", progress)
col4.metric("완료", done)
col5.metric("총 수량", qty)

st.divider()


# ---------------------------
# 4. 일정 등록
# ---------------------------
st.subheader("1. 시공 일정 등록")

c1, c2, c3 = st.columns(3)

date = c1.date_input("시공 날짜")
site = c2.text_input("설치현장")
manager = c3.text_input("시공담당")

c4, c5 = st.columns(2)
quantity = c4.number_input("수량", min_value=0)
memo = c5.text_input("비고")

if st.button("등록하기"):
    new_data = pd.DataFrame([{
        "날짜": str(date),
        "설치현장": site,
        "시공담당": manager,
        "수량": quantity,
        "비고": memo,
        "상태": "진행중",
        "완료일": ""
    }])

    df = pd.concat([df, new_data], ignore_index=True)
    save_data(df)
    st.success("등록 완료!")
    st.rerun()


# ---------------------------
# 5. 일정 보기
# ---------------------------
st.subheader("2. 시공 일정 보기")
st.dataframe(df, use_container_width=True)


# ---------------------------
# 6. 완료 처리
# ---------------------------
st.subheader("3. 완료 처리")

if not df.empty:
    options = [
        f"{i} | {row['날짜']} | {row['설치현장']}"
        for i, row in df.iterrows()
        if row["상태"] == "진행중"
    ]

    if options:
        selected = st.selectbox("완료할 일정 선택", options)

        if st.button("완료 처리"):
            idx = int(selected.split("|")[0])
            df.loc[idx, "상태"] = "완료"
            df.loc[idx, "완료일"] = str(datetime.today().date())
            save_data(df)
            st.success("완료 처리됨")
            st.rerun()


# ---------------------------
# 7. 삭제
# ---------------------------
st.subheader("4. 일정 삭제")

if not df.empty:
    options = [
        f"{i} | {row['날짜']} | {row['설치현장']}"
        for i, row in df.iterrows()
    ]

    selected = st.selectbox("삭제할 일정 선택", options)

    if st.button("삭제"):
        idx = int(selected.split("|")[0])
        df = df.drop(index=idx).reset_index(drop=True)
        save_data(df)
        st.success("삭제 완료")
        st.rerun()