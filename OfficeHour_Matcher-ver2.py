import streamlit as st
import pandas as pd
import requests
import gspread
from datetime import datetime, timedelta

# --- 0. 基本配置 ---
st.set_page_config(page_title="SmartSlot | 淡江智慧媒合", layout="wide", page_icon="🏫")

# --- Google Sheets ---
def connect_gsheet():
    try:
        # 直接使用 gspread 內建的功能，它會自動處理驗證
        client = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
        return client.open("TKU_Match_Log").worksheet("CaseMapping")
    except Exception as e:
        st.error(f"❌ 無法連線至 Google Sheets: {e}")
        st.stop()

def load_data():
    try:
        sheet = connect_gsheet()
        data = sheet.get_all_records()
        return pd.DataFrame(data)
    except:
        return pd.DataFrame()

def save_mapping(case_id, a, b, slots):
    sheet = connect_gsheet()
    now_tw = pd.Timestamp.now(tz='Asia/Taipei').strftime('%Y-%m-%d %H:%M:%S')
    
    sheet.append_row([
        now_tw,
        str(case_id),
        a,
        b,
        ",".join(slots[:6]),
        "",
        "",
        ""
    ])

def save_final(case_id, a, b, slots, day, slot, is_rec):
    sheet = connect_gsheet()
    now_tw = pd.Timestamp.now(tz='Asia/Taipei').strftime('%Y-%m-%d %H:%M:%S')
    
    sheet.append_row([
        now_tw,
        str(case_id),
        a,
        b,
        ",".join(slots),
        day,
        slot,
        is_rec
    ])

# --- 節次 ---
TIME_MAP = {
    "1": "(08:10 ~ 09:00)", "2": "(09:10 ~ 10:00)",
    "3": "(10:10 ~ 11:00)", "4": "(11:10 ~ 12:00)",
    "5": "(12:10 ~ 13:00)", "6": "(13:10 ~ 14:00)",
    "7": "(14:10 ~ 15:00)", "8": "(15:10 ~ 16:00)",
    "9": "(16:10 ~ 17:00)", "10": "(18:10 ~ 19:00)",
    "11": "(19:10 ~ 20:00)", "12": "(20:10 ~ 21:00)",
    "13": "(21:10 ~ 22:00)", "14": "(22:10 ~ 23:00)"
}

# --- 課表處理 ---
def fetch_and_clean_schedule(url):
    try:
        response = requests.get(url, timeout=10)
        dfs = pd.read_html(response.text)

        for df in dfs:
            if len(df) > 10:
                df = df.iloc[:, :8].copy()
                df.columns = ['節次','一','二','三','四','五','六','日']
                df = df.fillna('').astype(str)
                df = df[df['節次'].str.contains('第|\d', na=False)]

                def add_time(x):
                    k = x.replace("第","").replace("節","").strip()
                    return f"{x} {TIME_MAP[k]}" if k in TIME_MAP else x

                df['節次'] = df['節次'].apply(add_time)
                return df.reset_index(drop=True)
    except:
        return None

def find_all_slots(df_a, df_b):
    res = []
    days = ['一','二','三','四','五']

    def ok(v):
        v = v.replace('nan','').replace('None','').replace(' ','')
        return v=='' or v=='◎' or ('◎在校研究' in v and len(v)<10)

    for i,row in df_a.iterrows():
        if i>=len(df_b): break
        r2 = df_b.iloc[i]
        slot = row['節次']

        if any(x in slot for x in ["第1節","第5節","第一節","第五節","1","5"]):
            continue

        for d in days:
            if ok(row[d]) and ok(r2[d]):
                res.append(f"星期{d} {slot}")
            if len(res)>=6: return res
    return res

# --- UI ---
st.sidebar.title("🧭 系統選單")
mode = st.sidebar.radio("選擇階段", ["1. 智慧媒合比對", "2. 最終結果登記"])

# =======================
# 第一階段
# =======================
if mode == "1. 智慧媒合比對":
    st.title("教師駐校時間媒合系統")

    case_id = st.text_input("📑 請輸入書審案件流水號", placeholder="例如：II11301")
    file = st.file_uploader("1️⃣ 上傳老師名單 (Excel)", type=["xlsx"])

    if file:
        df = pd.read_excel(file)
        df['科系'] = df['科系'].fillna('未分類')
        df['姓名'] = df['姓名'].fillna('未知')

        col_sel1, col_sel2 = st.columns(2)
        with col_sel1:
            st.info("👤 委員A")
            dept_list_a = sorted(df['科系'].unique())
            dept_a = st.selectbox("選擇科系 (A)", dept_list_a, key="da")
            name_a = st.selectbox("選擇姓名 (A)", sorted(df[df['科系'] == dept_a]['姓名'].tolist()), key="na")
            url_a = df[(df['科系'] == dept_a) & (df['姓名'] == name_a)]['連結'].values[0]
    
        with col_sel2:
            st.info("👤 委員B")
            dept_list_b = sorted(df['科系'].unique())
            dept_b = st.selectbox("選擇科系 (B)", dept_list_b, key="db")
            name_b = st.selectbox("選擇姓名 (B)", sorted(df[df['科系'] == dept_b]['姓名'].tolist()), key="nb")
            url_b = df[(df['科系'] == dept_b) & (df['姓名'] == name_b)]['連結'].values[0]

        if st.button("⚡ 開始媒合"):
            if not case_id:
                st.warning("⚠️ 請先輸入『案件流水號』再開始媒合，以便後續結果登記。")
            else:
                with st.spinner("正在精準分析課表時段..."):
                    df_a = fetch_and_clean_schedule(url_a)
                    df_b = fetch_and_clean_schedule(url_b)

            if df_a is not None and df_b is not None:
                results = find_all_slots(df_a, df_b)

                if results:
                    # 呼叫寫入 CaseMapping 分頁的函數
                    save_mapping(case_id, name_a, name_b, results)
                    st.toast(f"✅ 案件 {case_id} 資料已儲存", icon="☁️")

                st.subheader("💡 系統推薦：最佳媒合時段 Top 3")
                top_3 = results[:3]
                other_3 = results[3:6]
        
                if top_3:
                    cols = st.columns(len(top_3))
                    for i, slot in enumerate(top_3):
                        cols[i].success(f"🏆 推薦順位 {i+1}\n\n**{slot}**")
                        
                    if other_3:
                        st.markdown("---")
                        st.subheader("📋 其他可參考的時段")
                        cols_other = st.columns(len(other_3))
                        for i, slot in enumerate(other_3):
                            cols_other[i].info(f"📍 備選方案 {i+1}\n\n**{slot}**")
                else:
                    st.warning("查無符合條件的共同時段。請檢查下方原始課表是否有共同空白處。")
        
                st.divider()
                v_col1, v_col2 = st.columns(2)
                with v_col1:
                    st.caption(f"📊 {name_a} 老師原始課表")
                    st.dataframe(df_a, use_container_width=True, hide_index=True)
                with v_col2:
                    st.caption(f"📊 {name_b} 老師原始課表")
                    st.dataframe(df_b, use_container_width=True, hide_index=True)
                        
            else:
                st.error("讀取失敗，請確認網址是否有效。")

# =======================
# 第二階段（完整版）
# =======================
elif mode == "2. 最終結果登記":
    st.title("✍️ 會議安排時段回饋")

    # -------------------------------
    # 1. 讀取資料
    # -------------------------------
    df_all = load_data()

    df_mapping = pd.DataFrame()
    case_options = ["請選擇流水號..."]
    t_a, t_b, candidate_slots = "未知", "未知", []

    if not df_all.empty:
        try:
            df_all = df_all.sort_values("timestamp")

            # 👉 分 mapping / final（用 final_day 判斷）
            df_map = df_all[df_all["final_day"] == ""]
            df_final = df_all[df_all["final_day"] != ""]

            # 👉 每個 case 最新 mapping
            df_map = (
                df_map.groupby("case_id")
                .tail(1)
            )

            # 👉 每個 case 最新 final
            df_final = (
                df_final.groupby("case_id")
                .tail(1)
            )

            # 👉 合併（模擬你原本 merge）
            df_mapping = pd.merge(
                df_map,
                df_final[["case_id", "final_day", "final_slot", "is_recommend"]],
                on="case_id",
                how="left"
            )

            if not df_mapping.empty:
                case_list = df_mapping['case_id'].astype(str).tolist()
                case_options = ["請選擇流水號..."] + case_list[::-1]

        except Exception as e:
            st.error(f"讀取資料失敗: {e}")

    # -------------------------------
    # 2. 選擇案件
    # -------------------------------
    search_id = st.selectbox(
        "🔍 選擇書審案件流水號",
        options=case_options,
        key="main_search"
    )

    if search_id != "請選擇流水號...":
        match = df_mapping[df_mapping['case_id'].astype(str) == str(search_id)]

        if not match.empty:
            row = match.iloc[0]

            t_a = str(row['teacher_a'])
            t_b = str(row['teacher_b'])
            raw_slots = str(row['candidate_slots'])
            candidate_slots = [s.strip() for s in raw_slots.split(",") if s.strip()]

            st.success(f"✅ 已帶入案件：{search_id} | 委員：{t_a} & {t_b}")

            # 👉 顯示已選結果
            if pd.notna(row.get("final_slot")) and row.get("final_slot") != "":
                st.info(f"📌 已登記結果：星期{row['final_day']} {row['final_slot']}")

    st.divider()

    # -------------------------------
    # 3. 輸入模式（保留原版）
    # -------------------------------
    input_mode = st.radio(
        "請選擇輸入模式：",
        ["從推薦時段中挑選", "手動輸入其他時段"],
        horizontal=True
    )

    final_day = ""
    final_slot = ""
    is_recommend = "No"

    # -------------------------------
    # 4. 表單（還原原版 UX）
    # -------------------------------
    with st.form("final_form", clear_on_submit=False):

        if input_mode == "從推薦時段中挑選":
            display_options = ["-- 請選擇一個推薦時段 --"] + candidate_slots

            chosen_recommend = st.selectbox(
                "系統推薦名單：",
                options=display_options
            )

            if chosen_recommend and "--" not in chosen_recommend:
                parts = chosen_recommend.split(" ", 1)
                final_day = parts[0].replace("星期", "")
                final_slot = parts[1] if len(parts) > 1 else ""
                is_recommend = "Yes"

        else:
            col_d, col_t = st.columns(2)

            with col_d:
                final_day = st.selectbox("確認星期", ["一", "二", "三", "四", "五"])

            with col_t:
                final_slot = st.text_input(
                    "手動輸入其他時段",
                    placeholder="例：14:10 ~ 15:00"
                )

            is_recommend = "No"

        # -------------------------------
        # 5. 提交（改成寫入 Google Sheets）
        # -------------------------------
        if st.form_submit_button("📤 提交最終成交紀錄"):

            if not final_slot or search_id == "請選擇流水號...":
                st.error("❌ 請確保已選擇流水號並填寫時段")

            else:
                try:
                    save_final(
                        search_id,
                        t_a,
                        t_b,
                        candidate_slots,
                        final_day,
                        final_slot,
                        is_recommend
                    )

                    st.balloons()
                    st.success(f"🎉 案件 {search_id} 已更新成功！")

                except Exception as e:
                    st.error(f"寫入失敗：{e}")
