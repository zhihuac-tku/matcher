import streamlit as st
import pandas as pd
import requests
from supabase import create_client
from datetime import datetime

# =======================
# Supabase 初始化
# =======================
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# =======================
# Streamlit 設定
# =======================
st.set_page_config(page_title="SmartSlot | 淡江智慧媒合", layout="wide", page_icon="🏫")

# =======================
# Supabase DB functions
# =======================

def load_data():
    res = supabase.table("case_mapping").select("*").execute()
    data = res.data

    if not data:
        return pd.DataFrame(columns=[
            "timestamp","case_id","teacher_a","teacher_b",
            "candidate_slots","final_day","final_slot","is_recommend"
        ])

    return pd.DataFrame(data)


def save_mapping(case_id, a, b, slots):
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    supabase.table("case_mapping").insert({
        "timestamp": now,
        "case_id": str(case_id),
        "teacher_a": a,
        "teacher_b": b,
        "candidate_slots": ",".join(slots[:6]),
        "final_day": "",
        "final_slot": "",
        "is_recommend": ""
    }).execute()


def save_final(case_id, a, b, slots, day, slot, is_rec):
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    supabase.table("case_mapping").insert({
        "timestamp": now,
        "case_id": str(case_id),
        "teacher_a": a,
        "teacher_b": b,
        "candidate_slots": ",".join(slots),
        "final_day": day,
        "final_slot": slot,
        "is_recommend": is_rec
    }).execute()

# =======================
# 節次時間
# =======================
TIME_MAP = {
    "1": "(08:10 ~ 09:00)", "2": "(09:10 ~ 10:00)",
    "3": "(10:10 ~ 11:00)", "4": "(11:10 ~ 12:00)",
    "5": "(12:10 ~ 13:00)", "6": "(13:10 ~ 14:00)",
    "7": "(14:10 ~ 15:00)", "8": "(15:10 ~ 16:00)",
    "9": "(16:10 ~ 17:00)", "10": "(18:10 ~ 19:00)",
    "11": "(19:10 ~ 20:00)", "12": "(20:10 ~ 21:00)",
    "13": "(21:10 ~ 22:00)", "14": "(22:10 ~ 23:00)"
}

# =======================
# 課表抓取
# =======================
def fetch_and_clean_schedule(url):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=10)
        response.encoding = 'utf-8'
        dfs = pd.read_html(response.text)

        target_df = None
        for df in dfs:
            if len(df) > 10:
                target_df = df
                break

        if target_df is None:
            return None

        df_clean = target_df.iloc[:, :8].copy()
        df_clean.columns = ['節次','一','二','三','四','五','六','日']
        df_clean = df_clean.fillna('').astype(str)

        df_clean = df_clean[df_clean['節次'].str.contains('第|\\d', na=False)].reset_index(drop=True)

        def add_time(x):
            k = x.strip().replace("第","").replace("節","")
            return f"{x} {TIME_MAP.get(k,'')}" if k in TIME_MAP else x

        df_clean['節次'] = df_clean['節次'].apply(add_time)

        return df_clean

    except:
        return None

# =======================
# 媒合邏輯（完全保留）
# =======================
def find_all_slots(df_a, df_b):
    all_slots = []
    days = ['一','二','三','四','五']

    def exclude(t):
        t = t.replace(" ","")
        if any(k in t for k in ["第1節","第5節","第一節","第五節"]):
            return True
        if t.startswith("1(") or t.startswith("5("):
            return True
        if t in ["1","5"]:
            return True
        return False

    for i, row_a in df_a.iterrows():
        if i >= len(df_b):
            break

        row_b = df_b.iloc[i]
        slot = str(row_a['節次']).strip()

        if exclude(slot):
            continue

        for d in days:
            a = str(row_a[d]).strip()
            b = str(row_b[d]).strip()

            def ok(v):
                v = v.replace("nan","").replace("None","").replace(" ","").strip()
                return v == "" or v == "◎" or ("◎在校研究" in v and len(v) < 10)

            if ok(a) and ok(b):
                all_slots.append(f"星期{d} {slot}")

            if len(all_slots) >= 6:
                return all_slots

    return all_slots

# =======================
# UI
# =======================
st.sidebar.title("🧭 系統選單")
mode = st.sidebar.radio("選擇階段", ["1. 智慧媒合比對", "2. 最終結果登記"])

# =======================
# 第一階段
# =======================
if mode == "1. 智慧媒合比對":

    st.title("教師駐校時間媒合系統")

    case_id = st.text_input("書審流水號")
    file = st.file_uploader("上傳老師 Excel", type=["xlsx"])

    if file:
        df = pd.read_excel(file)
        df['科系'] = df['科系'].fillna('未分類')
        df['姓名'] = df['姓名'].fillna('未知')

        col1, col2 = st.columns(2)

        with col1:
            dept_a = st.selectbox("A科系", sorted(df['科系'].unique()))
            name_a = st.selectbox("A姓名", df[df['科系']==dept_a]['姓名'])
            url_a = df[(df['科系']==dept_a)&(df['姓名']==name_a)]['連結'].values[0]

        with col2:
            dept_b = st.selectbox("B科系", sorted(df['科系'].unique()))
            name_b = st.selectbox("B姓名", df[df['科系']==dept_b]['姓名'])
            url_b = df[(df['科系']==dept_b)&(df['姓名']==name_b)]['連結'].values[0]

        if st.button("開始媒合"):

            df_a = fetch_and_clean_schedule(url_a)
            df_b = fetch_and_clean_schedule(url_b)

            if df_a is not None and df_b is not None:
                results = find_all_slots(df_a, df_b)

                if results:
                    save_mapping(case_id, name_a, name_b, results)

                st.write(results)

# =======================
# 第二階段
# =======================
else:

    st.title("最終結果登記")

    df_all = load_data()

    if not df_all.empty:

        df_all = df_all.sort_values("timestamp")

        df_map = df_all[df_all["final_day"] == ""].groupby("case_id").tail(1)
        df_final = df_all[df_all["final_day"] != ""].groupby("case_id").tail(1)

        df_mapping = pd.merge(
            df_map,
            df_final[["case_id","final_day","final_slot","is_recommend"]],
            on="case_id",
            how="left"
        )

        case_list = df_mapping["case_id"].tolist()
    else:
        case_list = []

    search = st.selectbox("選案件", ["請選擇"] + case_list)

    if search != "請選擇":

        row = df_mapping[df_mapping["case_id"] == search].iloc[0]

        slots = row["candidate_slots"].split(",")

        choice = st.selectbox("推薦時段", slots)

        if st.button("提交"):

            parts = choice.split(" ")

            save_final(
                search,
                row["teacher_a"],
                row["teacher_b"],
                slots,
                parts[0].replace("星期",""),
                parts[1],
                "Yes"
            )

            st.success("完成")