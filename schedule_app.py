import streamlit as st
import pandas as pd
import os
from datetime import date, datetime

st.set_page_config(page_title="시공 일정 관리", layout="wide")

DATA_FILE = "construction_schedule.csv"

# -----------------------------
# 데이터 불러오기 / 저장
# -----------------------------
def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE, encoding="utf-8-sig")

        expected_cols = ["날짜", "설치현장", "시공담당", "수량", "비고", "완료", "완료일"]
        for col in expected_cols:
            if col not in df.columns:
                df[col] = ""

        if not df.empty:
            df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce").dt.date
            df["완료일"] = pd.to_datetime(df["완료일"], errors="coerce").dt.date

            # 완료 컬럼 bool 처리
            df["완료"] = df["완료"].astype(str).str.lower().map({
                "true": True,
                "false": False
            })
            df["완료"] = df["완료"].fillna(False)

            # 수량 숫자 처리
            df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)

        return df

    return pd.DataFrame(columns=["날짜", "설치현장", "시공담당", "수량", "비고", "완료", "완료일"])


def save_data(df):
    save_df = df.copy()
    save_df["날짜"] = save_df["날짜"].astype(str)
    save_df["완료일"] = save_df["완료일"].astype(str)
    save_df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")


df = load_data()

# -----------------------------
# 제목
# -----------------------------
st.title("📅 시공 일정 관리 프로그램")
st.write("시공팀 일정 등록, 완료 체크, 진행 현황 확인용 프로그램입니다.")

# -----------------------------
# 상단 요약
# -----------------------------
total_count = len(df)
done_count = len(df[df["완료"] == True])
pending_count = len(df[df["완료"] == False])
total_qty = pd.to_numeric(df["수량"], errors="coerce").fillna(0).sum()

today = date.today()
today_count = 0
if not df.empty:
    today_count = len(df[df["날짜"] == today])

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("전체 일정", total_count)
c2.metric("오늘 일정", today_count)
c3.metric("진행중", pending_count)
c4.metric("완료", done_count)
c5.metric("총 수량", int(total_qty))

st.divider()

# -----------------------------
# 일정 등록
# -----------------------------
st.subheader("1. 시공 일정 등록")

with st.form("schedule_form"):
    col1, col2, col3 = st.columns(3)

    with col1:
        work_date = st.date_input("시공 날짜", value=date.today())
    with col2:
        site_name = st.text_input("설치현장")
    with col3:
        manager = st.text_input("시공담당")

    col4, col5 = st.columns(2)

    with col4:
        quantity = st.number_input("수량", min_value=0, step=1)
    with col5:
        note = st.text_input("비고")

    submit = st.form_submit_button("등록하기")

    if submit:
        if site_name.strip() == "" or manager.strip() == "":
            st.warning("설치현장과 시공담당은 꼭 입력해주세요.")
        else:
            new_row = pd.DataFrame([{
                "날짜": work_date,
                "설치현장": site_name.strip(),
                "시공담당": manager.strip(),
                "수량": int(quantity),
                "비고": note.strip(),
                "완료": False,
                "완료일": pd.NaT
            }])

            df = pd.concat([df, new_row], ignore_index=True)
            save_data(df)
            st.success("시공 일정이 등록되었습니다.")
            st.rerun()

st.divider()

# -----------------------------
# 일정 보기
# -----------------------------
st.subheader("2. 시공 일정 보기")

col_a, col_b = st.columns(2)
with col_a:
    status_filter = st.selectbox("상태 선택", ["전체", "진행중", "완료"])
with col_b:
    date_filter = st.selectbox("날짜 기준", ["전체", "오늘", "이번주", "지난 일정"])

filtered_df = df.copy()

if status_filter == "진행중":
    filtered_df = filtered_df[filtered_df["완료"] == False]
elif status_filter == "완료":
    filtered_df = filtered_df[filtered_df["완료"] == True]

if date_filter == "오늘":
    filtered_df = filtered_df[filtered_df["날짜"] == today]
elif date_filter == "지난 일정":
    filtered_df = filtered_df[filtered_df["날짜"] < today]
elif date_filter == "이번주":
    today_ts = pd.Timestamp(today)
    start_week = (today_ts - pd.Timedelta(days=today_ts.weekday())).date()
    end_week = (pd.Timestamp(start_week) + pd.Timedelta(days=6)).date()
    filtered_df = filtered_df[
        (filtered_df["날짜"] >= start_week) & (filtered_df["날짜"] <= end_week)
    ]

if filtered_df.empty:
    st.info("조건에 맞는 일정이 없습니다.")
else:
    display_df = filtered_df.copy()
    display_df = display_df.sort_values(by=["완료", "날짜", "설치현장"], ascending=[True, True, True])

    display_df["상태"] = display_df["완료"].apply(lambda x: "완료" if x else "진행중")
    display_df = display_df[["날짜", "설치현장", "시공담당", "수량", "비고", "상태", "완료일"]]

    st.dataframe(display_df, use_container_width=True, hide_index=True)

st.divider()

# -----------------------------
# 완료 처리
# -----------------------------
st.subheader("3. 완료 처리")

incomplete_df = df[df["완료"] == False]

if incomplete_df.empty:
    st.success("현재 진행중인 일정이 없습니다.")
else:
    options = [
        f"{idx} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
        for idx, row in incomplete_df.iterrows()
    ]

    selected_option = st.selectbox("완료 처리할 일정을 선택하세요", options)

    if st.button("완료로 변경"):
        selected_idx = int(selected_option.split("|")[0].strip())
        df.loc[selected_idx, "완료"] = True
        df.loc[selected_idx, "완료일"] = today
        save_data(df)
        st.success("완료 처리되었습니다.")
        st.rerun()

# -----------------------------
# 완료 취소
# -----------------------------
st.subheader("4. 완료 취소")

completed_df = df[df["완료"] == True]

if completed_df.empty:
    st.info("완료된 일정이 없습니다.")
else:
    complete_options = [
        f"{idx} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
        for idx, row in completed_df.iterrows()
    ]

    selected_complete = st.selectbox("진행중으로 되돌릴 일정을 선택하세요", complete_options)

    if st.button("진행중으로 변경"):
        selected_idx = int(selected_complete.split("|")[0].strip())
        df.loc[selected_idx, "완료"] = False
        df.loc[selected_idx, "완료일"] = pd.NaT
        save_data(df)
        st.success("진행중 상태로 변경되었습니다.")
        st.rerun()

st.divider()

# -----------------------------
# 일정 삭제
# -----------------------------
st.subheader("5. 일정 삭제")

if df.empty:
    st.info("삭제할 일정이 없습니다.")
else:
    delete_options = [
        f"{idx} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
        for idx, row in df.iterrows()
    ]

    selected_delete = st.selectbox("삭제할 일정을 선택하세요", delete_options)

    if st.button("선택 일정 삭제"):
        selected_idx = int(selected_delete.split("|")[0].strip())
        df = df.drop(index=selected_idx).reset_index(drop=True)
        save_data(df)
        st.success("일정이 삭제되었습니다.")
        st.rerun()