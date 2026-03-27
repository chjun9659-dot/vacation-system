import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date
import json

st.set_page_config(page_title="시공 일정 관리", layout="wide")

# ---------------------------
# 1. 구글시트 연결
# ---------------------------
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

if "gcp_service_account" in st.secrets:
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
else:
    creds = ServiceAccountCredentials.from_json_keyfile_name("key.json", scope)

client = gspread.authorize(creds)

# 구글시트 이름
SHEET_NAME = "시공일정"
sheet = client.open(SHEET_NAME).sheet1

EXPECTED_COLUMNS = ["날짜", "설치현장", "시공담당", "수량", "비고", "상태", "완료일"]


# ---------------------------
# 2. 데이터 불러오기 / 저장
# ---------------------------
def ensure_sheet_header():
    values = sheet.get_all_values()
    if not values:
        sheet.update("A1:G1", [EXPECTED_COLUMNS])
        return

    header = values[0]
    if header != EXPECTED_COLUMNS:
        existing = pd.DataFrame(values[1:], columns=header if header else None)
        for col in EXPECTED_COLUMNS:
            if col not in existing.columns:
                existing[col] = ""
        existing = existing[EXPECTED_COLUMNS]
        save_data(existing)


def load_data():
    ensure_sheet_header()
    data = sheet.get_all_records()
    df = pd.DataFrame(data)

    if df.empty:
        return pd.DataFrame(columns=EXPECTED_COLUMNS)

    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df = df[EXPECTED_COLUMNS]
    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)
    df["날짜"] = df["날짜"].astype(str)
    df["완료일"] = df["완료일"].astype(str)
    df["상태"] = df["상태"].astype(str).replace("", "진행중")

    return df


def save_data(df):
    save_df = df.copy()
    for col in EXPECTED_COLUMNS:
        if col not in save_df.columns:
            save_df[col] = ""

    save_df = save_df[EXPECTED_COLUMNS].fillna("")
    save_df["수량"] = pd.to_numeric(save_df["수량"], errors="coerce").fillna(0).astype(int)

    rows = [save_df.columns.tolist()] + save_df.astype(str).values.tolist()
    sheet.clear()
    sheet.update(rows)


df = load_data()

# 행 번호 고정용
df = df.reset_index(drop=True)
df["row_id"] = df.index


# ---------------------------
# 3. 상단 제목 / 요약
# ---------------------------
st.title("📅 시공 일정 관리 프로그램")
st.write("시공팀 일정 등록, 수정, 완료 체크, 진행 현황 확인용 프로그램입니다.")

today_str = str(date.today())

total_count = len(df)
today_count = len(df[df["날짜"] == today_str])
progress_count = len(df[df["상태"] == "진행중"])
done_count = len(df[df["상태"] == "완료"])
total_qty = int(df["수량"].sum()) if not df.empty else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("전체 일정", total_count)
c2.metric("오늘 일정", today_count)
c3.metric("진행중", progress_count)
c4.metric("완료", done_count)
c5.metric("총 수량", total_qty)

st.divider()


# ---------------------------
# 4. 일정 등록
# ---------------------------
st.subheader("1. 시공 일정 등록")

with st.form("add_schedule_form"):
    a1, a2, a3 = st.columns(3)
    work_date = a1.date_input("시공 날짜", value=date.today())
    site_name = a2.text_input("설치현장")
    manager_name = a3.text_input("시공담당")

    a4, a5 = st.columns(2)
    quantity = a4.number_input("수량", min_value=0, step=1, value=0)
    note = a5.text_input("비고")

    submitted = st.form_submit_button("등록하기")

    if submitted:
        if not site_name.strip():
            st.warning("설치현장을 입력해주세요.")
        elif not manager_name.strip():
            st.warning("시공담당을 입력해주세요.")
        else:
            new_row = pd.DataFrame([{
                "날짜": str(work_date),
                "설치현장": site_name.strip(),
                "시공담당": manager_name.strip(),
                "수량": int(quantity),
                "비고": note.strip(),
                "상태": "진행중",
                "완료일": ""
            }])

            save_df = df[EXPECTED_COLUMNS].copy() if not df.empty else pd.DataFrame(columns=EXPECTED_COLUMNS)
            save_df = pd.concat([save_df, new_row], ignore_index=True)
            save_data(save_df)
            st.success("등록 완료!")
            st.rerun()

st.divider()


# ---------------------------
# 5. 오늘 일정
# ---------------------------
st.subheader("2. 오늘 일정")

today_df = df[df["날짜"] == today_str].copy()

if today_df.empty:
    st.info("오늘 일정이 없습니다.")
else:
    show_today = today_df[["날짜", "설치현장", "시공담당", "수량", "비고", "상태", "완료일"]]
    st.dataframe(show_today, use_container_width=True, hide_index=True)

st.divider()


# ---------------------------
# 6. 전체 일정 보기 / 필터
# ---------------------------
st.subheader("3. 시공 일정 보기")

managers = ["전체"] + sorted([m for m in df["시공담당"].dropna().unique().tolist() if str(m).strip() != ""])

f1, f2, f3, f4 = st.columns(4)
status_filter = f1.selectbox("상태 선택", ["전체", "진행중", "완료"])
manager_filter = f2.selectbox("담당자 선택", managers)
date_filter = f3.selectbox("날짜 기준", ["전체", "오늘", "미래", "지난 일정"])
keyword = f4.text_input("검색", placeholder="설치현장 / 비고 검색")

filtered_df = df.copy()

if status_filter != "전체":
    filtered_df = filtered_df[filtered_df["상태"] == status_filter]

if manager_filter != "전체":
    filtered_df = filtered_df[filtered_df["시공담당"] == manager_filter]

if date_filter == "오늘":
    filtered_df = filtered_df[filtered_df["날짜"] == today_str]
elif date_filter == "미래":
    filtered_df = filtered_df[filtered_df["날짜"] > today_str]
elif date_filter == "지난 일정":
    filtered_df = filtered_df[filtered_df["날짜"] < today_str]

if keyword.strip():
    kw = keyword.strip()
    filtered_df = filtered_df[
        filtered_df["설치현장"].astype(str).str.contains(kw, case=False, na=False) |
        filtered_df["비고"].astype(str).str.contains(kw, case=False, na=False)
    ]

show_df = filtered_df[["날짜", "설치현장", "시공담당", "수량", "비고", "상태", "완료일"]].copy()

if show_df.empty:
    st.info("조건에 맞는 일정이 없습니다.")
else:
    st.dataframe(show_df, use_container_width=True, hide_index=True)

st.divider()


# ---------------------------
# 7. 일정 수정
# ---------------------------
st.subheader("4. 일정 수정")

if df.empty:
    st.info("수정할 일정이 없습니다.")
else:
    edit_options = [
        f"{row['row_id']} | {row['날짜']} | {row['설치현장']} | {row['시공담당']}"
        for _, row in df.iterrows()
    ]

    selected_edit = st.selectbox("수정할 일정 선택", edit_options, key="edit_select")
    edit_idx = int(selected_edit.split("|")[0].strip())
    edit_row = df.loc[df["row_id"] == edit_idx].iloc[0]

    with st.form("edit_schedule_form"):
        e1, e2, e3 = st.columns(3)
        edit_date = e1.date_input(
            "시공 날짜 수정",
            value=pd.to_datetime(edit_row["날짜"]).date() if str(edit_row["날짜"]).strip() else date.today()
        )
        edit_site = e2.text_input("설치현장 수정", value=str(edit_row["설치현장"]))
        edit_manager = e3.text_input("시공담당 수정", value=str(edit_row["시공담당"]))

        e4, e5, e6 = st.columns(3)
        edit_qty = e4.number_input("수량 수정", min_value=0, step=1, value=int(edit_row["수량"]))
        edit_note = e5.text_input("비고 수정", value=str(edit_row["비고"]))
        edit_status = e6.selectbox("상태 수정", ["진행중", "완료"], index=0 if edit_row["상태"] == "진행중" else 1)

        edit_submit = st.form_submit_button("수정 저장")

        if edit_submit:
            save_df = df[EXPECTED_COLUMNS].copy()
            save_df.loc[edit_idx, "날짜"] = str(edit_date)
            save_df.loc[edit_idx, "설치현장"] = edit_site.strip()
            save_df.loc[edit_idx, "시공담당"] = edit_manager.strip()
            save_df.loc[edit_idx, "수량"] = int(edit_qty)
            save_df.loc[edit_idx, "비고"] = edit_note.strip()
            save_df.loc[edit_idx, "상태"] = edit_status

            if edit_status == "완료" and not str(save_df.loc[edit_idx, "완료일"]).strip():
                save_df.loc[edit_idx, "완료일"] = today_str
            elif edit_status == "진행중":
                save_df.loc[edit_idx, "완료일"] = ""

            save_data(save_df)
            st.success("수정 완료!")
            st.rerun()

st.divider()


# ---------------------------
# 8. 완료 처리
# ---------------------------
st.subheader("5. 완료 처리")

progress_df = df[df["상태"] == "진행중"].copy()

if progress_df.empty:
    st.info("완료 처리할 일정이 없습니다.")
else:
    complete_options = [
        f"{row['row_id']} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
        for _, row in progress_df.iterrows()
    ]

    selected_complete = st.selectbox("완료 처리 일정 선택", complete_options, key="complete_select")

    if st.button("완료로 변경"):
        complete_idx = int(selected_complete.split("|")[0].strip())
        save_df = df[EXPECTED_COLUMNS].copy()
        save_df.loc[complete_idx, "상태"] = "완료"
        save_df.loc[complete_idx, "완료일"] = today_str
        save_data(save_df)
        st.success("완료 처리되었습니다.")
        st.rerun()

st.divider()


# ---------------------------
# 9. 완료 취소
# ---------------------------
st.subheader("6. 완료 취소")

done_df = df[df["상태"] == "완료"].copy()

if done_df.empty:
    st.info("완료 취소할 일정이 없습니다.")
else:
    cancel_options = [
        f"{row['row_id']} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
        for _, row in done_df.iterrows()
    ]

    selected_cancel = st.selectbox("완료 취소 일정 선택", cancel_options, key="cancel_select")

    if st.button("진행중으로 변경"):
        cancel_idx = int(selected_cancel.split("|")[0].strip())
        save_df = df[EXPECTED_COLUMNS].copy()
        save_df.loc[cancel_idx, "상태"] = "진행중"
        save_df.loc[cancel_idx, "완료일"] = ""
        save_data(save_df)
        st.success("완료 취소되었습니다.")
        st.rerun()

st.divider()


# ---------------------------
# 10. 일정 삭제
# ---------------------------
st.subheader("7. 일정 삭제")

if df.empty:
    st.info("삭제할 일정이 없습니다.")
else:
    delete_options = [
        f"{row['row_id']} | {row['날짜']} | {row['설치현장']} | {row['시공담당']} | 수량 {row['수량']}"
        for _, row in df.iterrows()
    ]

    selected_delete = st.selectbox("삭제할 일정 선택", delete_options, key="delete_select")

    if st.button("선택 일정 삭제"):
        delete_idx = int(selected_delete.split("|")[0].strip())
        save_df = df[EXPECTED_COLUMNS].copy()
        save_df = save_df.drop(index=delete_idx).reset_index(drop=True)
        save_data(save_df)
        st.success("삭제 완료!")
        st.rerun()