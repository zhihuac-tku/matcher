import streamlit as st
import pandas as pd
import requests
from supabase import create_client
from datetime import datetime
from io import StringIO
import time
import pytz

# =======================
# Supabase 初始化
# =======================
SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

tz = pytz.timezone("Asia/Taipei")
now = datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')

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
    now_iso = datetime.now(pytz.timezone("Asia/Taipei")).isoformat()

    supabase.table("case_mapping").insert({
        "timestamp": now_iso,
        "case_id": str(case_id),
        "teacher_a": a,
        "teacher_b": b,
        "candidate_slots": ",".join(slots[:6]),
        "final_day": "",
        "final_slot": "",
        "is_recommend": ""
    }).execute()

def save_final(case_id, a, b, slots, day, slot, is_rec):
    now_iso = datetime.now(pytz.timezone("Asia/Taipei")).isoformat()

    supabase.table("case_mapping").upsert({
        "case_id": str(case_id),       
        "timestamp": now_iso,        
        "teacher_a": a,
        "teacher_b": b,
        "candidate_slots": ",".join(slots),
        "final_day": day,
        "final_slot": slot,
        "is_recommend": is_rec
    }, on_conflict="case_id").execute() 

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
college_map = {
    "文學院": ["資圖系", "中文系", "歷史系", "資傳系", "大傳系"],
    "理學院": ["尖端材料科學學程", "化學系", "數學系", "物理系", "應用科學博士班"],
    "工學院": ["建築系", "機械系", "土木系", "化材系", "資工系", "航太系", "電機系", "水環系"],
    "商管學院": ["會計系", "財金系", "企管系", "國企系", "管科系", "資管系", "風保系", "公行系", "運管系", "經濟系", "統計系"],
    "外國語文學院": ["日文系", "英文系", "歐語系"],
    "國際事務學院": ["觀光系", "外交系", "政經系", "戰略所"],
    "教育學院": ["教心所", "教設系", "師培中心", "教科系"],
    "體育事務處": ["學動組"],
    "教務處": ["通核中心"],
    "AI創智學院": ["AI系"],
    "精準健康學院": ["高齡健康所", "智慧照護所"]
}
# =======================
# 課表抓取
# =======================
def fetch_and_clean_schedule(url):
    try:
        # 👉 建立 session（模擬瀏覽器）
        session = requests.Session()

        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
            "Referer": "https://www.google.com/",
            "Connection": "keep-alive"
        }

        # 👉 用 session 取代 requests.get
        response = session.get(url, headers=headers, timeout=10)
        response.encoding = 'utf-8'

        # 🔍 Debug（建議先保留）
        # st.write("Status:", response.status_code)
        # st.write(response.text[:500])

        dfs = pd.read_html(StringIO(response.text))

        target_df = None
        for df in dfs:
            if len(df) > 10:
                target_df = df
                break

        if target_df is not None:
            df_clean = target_df.iloc[:, :8].copy()
            df_clean.columns = ['節次', '一', '二', '三', '四', '五', '六', '日']
            df_clean = df_clean.fillna('').astype(str)

            df_clean = df_clean[df_clean['節次'].str.contains('第|\\d', na=False)].reset_index(drop=True)

            def add_time_info(slot_text):
                st_clean = slot_text.strip().replace("第", "").replace("節", "")
                if st_clean in TIME_MAP:
                    return f"{slot_text} {TIME_MAP[st_clean]}"
                return slot_text

            df_clean['節次'] = df_clean['節次'].apply(add_time_info)

            return df_clean

        return None

    except Exception as e:
        # 👉 很重要：把錯誤印出來
        st.error(f"抓課表錯誤: {e}")
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

    case_id = st.text_input("📑 請輸入書審案件流水號", placeholder="例如：II11301")
    file = st.file_uploader("1️⃣ 上傳老師名單 (Excel)", type=["xlsx"])

    if file:
        df = pd.read_excel(file)
        df['科系'] = df['科系'].fillna('未分類')
        df['姓名'] = df['姓名'].fillna('未知')

        col_sel1, col_sel2 = st.columns(2)

        # 👉 委員 A
        with col_sel1:
            st.info("👤 委員A")
        
            college_a = st.selectbox(
                "選擇學院 (A)",
                list(college_map.keys()),
                key="college_a"
            )
        
            # 🔥 關鍵：科系只顯示該學院
            dept_list_a = college_map[college_a]
        
            dept_a = st.selectbox(
                "選擇科系 (A)",
                dept_list_a,
                key="da"
            )
        
            # 👉 再過濾老師
            name_a = st.selectbox(
                "選擇姓名 (A)",
                sorted(df[df['科系'] == dept_a]['姓名'].tolist()),
                key="na"
            )
        
            url_a = df[
                (df['科系'] == dept_a) &
                (df['姓名'] == name_a)
            ]['連結'].values[0]

        # 👉 委員 B
        with col_sel2:
            st.info("👤 委員B")
        
            college_b = st.selectbox(
                "選擇學院 (B)",
                list(college_map.keys()),
                key="college_b"
            )
        
            dept_list_b = college_map[college_b]
        
            dept_b = st.selectbox(
                "選擇科系 (B)",
                dept_list_b,
                key="db"
            )
        
            name_b = st.selectbox(
                "選擇姓名 (B)",
                sorted(df[df['科系'] == dept_b]['姓名'].tolist()),
                key="nb"
            )
        
            url_b = df[
                (df['科系'] == dept_b) &
                (df['姓名'] == name_b)
            ]['連結'].values[0]

        # =======================
        # 👉 開始媒合
        # =======================
        if st.button("⚡ 開始媒合"):

            if not case_id:
                st.warning("⚠️ 請先輸入『案件流水號』")
                st.stop()

            with st.spinner("正在分析課表..."):
                df_a = fetch_and_clean_schedule(url_a)
                df_b = fetch_and_clean_schedule(url_b)

            if df_a is not None and df_b is not None:

                results = find_all_slots(df_a, df_b)

                # 👉 寫入 Supabase
                if results:
                    try:
                        save_mapping(case_id, name_a, name_b, results)
                        st.toast(f"✅ 案件 {case_id} 已儲存", icon="☁️")
                    except Exception as e:
                        st.error(f"❌ 寫入失敗：{e}")

                # =======================
                # 👉 推薦 UI（回來了🔥）
                # =======================
                st.subheader("💡 系統推薦：最佳媒合時段 Top 3")

                top_3 = results[:3]
                other_3 = results[3:6]

                if top_3:
                    cols = st.columns(len(top_3))
                    for i, slot in enumerate(top_3):
                        cols[i].success(f"🏆 推薦順位 {i+1}\n\n**{slot}**")

                    if other_3:
                        st.markdown("---")
                        st.subheader("📋 其他可參考時段")

                        cols2 = st.columns(len(other_3))
                        for i, slot in enumerate(other_3):
                            cols2[i].info(f"📍 備選 {i+1}\n\n**{slot}**")

                else:
                    st.warning("⚠️ 沒有共同時段")

                # =======================
                # 👉 🔥 這就是你要的 table（重點）
                # =======================
                st.divider()

                col1, col2 = st.columns(2)

                with col1:
                    st.caption(f"📊 {name_a} 老師課表")
                    st.dataframe(df_a, use_container_width=True, hide_index=True)

                with col2:
                    st.caption(f"📊 {name_b} 老師課表")
                    st.dataframe(df_b, use_container_width=True, hide_index=True)

            else:
                st.error("❌ 課表讀取失敗（請檢查網址）")

# =======================
# 第二階段
# =======================
# =======================
# 第二階段（完整版 + Supabase）
# =======================
else:

    st.title("✍️ 會議安排時段回饋")

    # -------------------------------
    # 1. 讀取資料
    # -------------------------------
    df_all = load_data()

    df_mapping = pd.DataFrame()
    case_options = ["請選擇流水號..."]

    if not df_all.empty:
        try:
            df_all = df_all.sort_values("timestamp")

            # 👉 mapping / final 分流
            df_map = df_all[df_all["final_day"] == ""].groupby("case_id").tail(1)
            df_final = df_all[df_all["final_day"] != ""].groupby("case_id").tail(1)

            # 👉 merge
            df_mapping = pd.merge(
                df_map,
                df_final[["case_id","final_day","final_slot","is_recommend"]],
                on="case_id",
                how="left"
            )

            if not df_mapping.empty:
                case_list = df_mapping["case_id"].astype(str).tolist()
                case_options = ["請選擇流水號..."] + case_list[::-1]

        except Exception as e:
            st.error(f"讀取資料失敗: {e}")

    # -------------------------------
    # 2. 選案件
    # -------------------------------
    search_id = st.selectbox(
        "🔍 選擇書審案件流水號",
        options=case_options,
        key="case_select"
    )

    t_a, t_b = "未知", "未知"
    candidate_slots = []

    if search_id != "請選擇流水號..." and not df_mapping.empty:

        match = df_mapping[df_mapping["case_id"].astype(str) == str(search_id)]

        if not match.empty:
            row = match.iloc[0]

            t_a = str(row["teacher_a"])
            t_b = str(row["teacher_b"])
            raw_slots = str(row["candidate_slots"])
            candidate_slots = [s.strip() for s in raw_slots.split(",") if s.strip()]

            st.success(f"✅ 已帶入案件：{search_id} | 👤 {t_a} & {t_b}")

            # 👉 顯示推薦清單（透明化）
            with st.expander("📋 查看系統推薦時段歷史"):
                for i, s in enumerate(candidate_slots, 1):
                    st.write(f"{i}. {s}")

            # 👉 已填結果顯示
            if pd.notna(row.get("final_slot")) and row.get("final_slot") != "":
                st.info(f"📌 已登記結果：星期{row['final_day']} {row['final_slot']}")

    st.divider()

    # -------------------------------
    # 3. 輸入模式
    # -------------------------------
    input_mode = st.radio(
        "請選擇輸入方式",
        ["從推薦時段中挑選", "手動輸入其他時段"],
        horizontal=True
    )

    final_day = ""
    final_slot = ""
    is_recommend = "No"

    # -------------------------------
    # 4. 表單（避免誤觸）
    # -------------------------------
    with st.form("final_form"):

        # 👉 推薦模式
        if input_mode == "從推薦時段中挑選":

            options = ["-- 請選擇推薦時段 --"] + candidate_slots

            chosen = st.selectbox("系統推薦時段", options=options)

            if chosen and "--" not in chosen:
                parts = chosen.split(" ", 1)
                final_day = parts[0].replace("星期", "")
                final_slot = parts[1] if len(parts) > 1 else ""
                is_recommend = "Yes"

        # 👉 手動模式
        else:
            col1, col2 = st.columns(2)

            with col1:
                final_day = st.selectbox("選擇星期", ["一","二","三","四","五"])

            with col2:
                time_options = [f"第{k}節 {v}" for k, v in TIME_MAP.items()]
                chosen_time = st.selectbox(
                    "選擇時段",
                    options=time_options,
                    index=None,
                    placeholder="請選擇節次..."
                )
                
                # 將選中的文字存入 final_slot
                if chosen_time:
                    # 例如選中 "第2節 (09:10 ~ 10:00)"，我們只存後面的時間或整串
                    final_slot = chosen_time
                else:
                    final_slot = ""

            is_recommend = "No"

        # -------------------------------
        # 5. 提交
        # -------------------------------
        submitted = st.form_submit_button("📤 提交最終結果")

        if submitted:

            # 👉 防呆
            if search_id == "請選擇流水號...":
                st.error("❌ 請先選擇案件")
                st.stop()

            if not final_slot:
                st.error("❌ 請填寫或選擇時段")
                st.stop()
                
            original_ts = str(row["timestamp"])
            
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
                st.toast(f"✅ 案件 {search_id} 已成功更新！", icon="🎉")

                # 👉 強制刷新（超重要）
                time.sleep(0.5)
                st.rerun()

            except Exception as e:
                st.error(f"❌ 寫入失敗：{e}")
