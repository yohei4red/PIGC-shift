# ========= フリーズ防止のおまじない =========
import matplotlib
matplotlib.use('Agg')
# ==========================================

import streamlit as st
import pulp
import pandas as pd
import matplotlib.pyplot as plt
import japanize_matplotlib  # クラウド用日本語フォント
from datetime import datetime, timedelta
import random

st.set_page_config(page_title="PIGC シフト自動作成アプリ", layout="wide")

# ==========================================
# ⚙️ サイドバー設定（入力フォーム化）
# ==========================================
st.sidebar.header("⚙️ 基本設定")

# 入力内容を確定させるための「フォーム」を作成
with st.sidebar.form(key='settings_form'):
    st.subheader("📅 期間の設定")
    start_date = st.date_input("シフト開始日", value=datetime(2026, 3, 21))
    num_days = st.number_input("期間（日数）", min_value=7, max_value=60, value=31)

    st.subheader("👤 全スタッフの登録")
    st.caption("1行に1人ずつ入力してください")
    default_staffs = "小畑 明美\n柴田 美幸\n河本 陽平\n植木 健太郎\n深田 尚\n佐藤 大吾郎\n中谷 吉男\n浜崎 圭子"
    master_staff_str = st.text_area("全スタッフ名", default_staffs, height=200)
    all_staffs = [s.strip() for s in master_staff_str.split("\n") if s.strip()]

    st.subheader("🏢 ポジション設定")
    staffs_am = st.multiselect("午前スタッフ", options=all_staffs, default=[s for s in ["小畑 明美", "柴田 美幸", "河本 陽平", "植木 健太郎"] if s in all_staffs])
    staffs_pm = st.multiselect("午後スタッフ", options=all_staffs, default=[s for s in ["深田 尚", "佐藤 大吾郎", "中谷 吉男", "浜崎 圭子"] if s in all_staffs])
    cross_staffs = st.multiselect("通し(午前・午後)スタッフ", options=all_staffs, default=[s for s in ["小畑 明美", "柴田 美幸"] if s in all_staffs])

    st.subheader("🆖 NGペア設定")
    df_ng_init = pd.DataFrame([{"スタッフ1": "河本 陽平", "スタッフ2": "植木 健太郎"}])
    edited_ng = st.data_editor(df_ng_init, num_rows="dynamic", hide_index=True, column_config={
        "スタッフ1": st.column_config.SelectboxColumn("スタッフ1", options=all_staffs),
        "スタッフ2": st.column_config.SelectboxColumn("スタッフ2", options=all_staffs),
    })

    st.subheader("📊 出勤日数")
    default_min_dict = {"小畑 明美": 20, "柴田 美幸": 15, "植木 健太郎": 13, "深田 尚": 10, "中谷 吉男": 10, "佐藤 大吾郎": 8}
    default_exact_dict = {"河本 陽平": 5, "浜崎 圭子": 0}
    df_rules = pd.DataFrame({
        "スタッフ": all_staffs,
        "最低日数": [default_min_dict.get(s, 0) for s in all_staffs],
        "ぴったり": [default_exact_dict.get(s, 0) for s in all_staffs]
    })
    edited_rules = st.data_editor(df_rules, hide_index=True)

    # スマホ利用者のための「更新ボタン」
    submit_button = st.form_submit_button(label='🔄 設定をアプリに反映する', use_container_width=True)

# ==========================================
# データ処理と状態管理
# ==========================================
# NGペアのリスト化
ng_pairs = []
for _, row in edited_ng.iterrows():
    s1, s2 = row.get("スタッフ1"), row.get("スタッフ2")
    if pd.notna(s1) and pd.notna(s2) and s1 != s2:
        ng_pairs.append((s1, s2))

# 出勤ルールの辞書化
min_shifts = {row["スタッフ"]: int(row["最低日数"]) for _, row in edited_rules.iterrows() if int(row["最低日数"]) > 0}
exact_shifts = {row["スタッフ"]: int(row["ぴったり"]) for _, row in edited_rules.iterrows() if int(row["ぴったり"]) > 0}

# 日付情報の生成
days_map = ["月", "火", "水", "木", "金", "土", "日"]
dates_info = []
for i in range(num_days): 
    current_date = start_date + timedelta(days=i)
    date_str = f"{current_date.month}/{current_date.day}"
    day_str = days_map[current_date.weekday()]
    dates_info.append({
        "id": date_str, "display": f"{day_str}\n{date_str}", 
        "is_weekend": current_date.weekday() >= 5,
        "is_sat": current_date.weekday() == 5, "is_sun": current_date.weekday() == 6
    })
date_ids = [d["id"] for d in dates_info]

# 状態の初期化と更新（スタッフ変更時など）
if 'gui_states' not in st.session_state or st.session_state.get('last_staffs') != all_staffs or st.session_state.get('last_dates') != date_ids:
    st.session_state.gui_states = {s: {d: 0 for d in date_ids} for s in all_staffs}
    st.session_state.last_staffs = all_staffs
    st.session_state.last_dates = date_ids

def get_val(var):
    val = pulp.value(var)
    return val if val is not None else 0

# ==========================================
# ファイル出力（Excel/PNG）
# ==========================================
def generate_files(x_am, x_pm, shain_am_vars, shain_pm_vars, status_box):
    schedule_data = []
    header_row = ["曜日\n日付"] + staffs_am + ["日付\n曜日"] + staffs_pm + ["午前", "午後"]
    
    title_row = [""] * len(header_row)
    title_row[1] = "【 午前打席 】"
    if len(staffs_am) + 2 < len(title_row): title_row[len(staffs_am) + 2] = "【 午後打席 】"
    title_row[-2] = "【 社員 】"
    
    schedule_data.append(title_row)
    schedule_data.append(header_row)
    
    for d_info in dates_info:
        d_id = d_info["id"]
        row = [d_info["display"]]
        for s in staffs_am:
            am_val, pm_val = get_val(x_am[s, d_id]), get_val(x_pm[s, d_id])
            if am_val == 1 and pm_val == 1: row.append("●☆")
            elif am_val == 1: row.append("●")
            else: row.append("")
        row.append(d_info["display"])
        for s in staffs_pm:
            row.append("●" if get_val(x_pm[s, d_id]) == 1 else "")
        row.append("●" if get_val(shain_am_vars[d_id]) > 0.5 else "")
        row.append("●" if get_val(shain_pm_vars[d_id]) > 0.5 else "")
        schedule_data.append(row)
        
    # 合計・日数行
    row_total = ["出勤日数"]
    for s in staffs_am: row_total.append(str(int(sum(get_val(x_am[s, d]) for d in date_ids))))
    row_total.append("出勤日数")
    for s in staffs_pm: row_total.append(str(int(sum(get_val(x_pm[s, d]) for d in date_ids))))
    row_total.extend([str(int(sum(get_val(shain_am_vars[d]) for d in date_ids))), str(int(sum(get_val(shain_pm_vars[d]) for d in date_ids)))])
    schedule_data.append(row_total)

    df = pd.DataFrame(schedule_data)
    
    # Excel
    status_box.info("⏳ エクセル作成中...")
    excel_file = "full_shift_schedule.xlsx"
    df.to_excel(excel_file, index=False, header=False)
    
    # PNG
    status_box.info("⏳ 写真(画像)作成中...")
    fig, ax = plt.subplots(figsize=(max(14, len(all_staffs)*1.2), 12))
    ax.axis('off')
    table = ax.table(cellText=df.values, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.5)
    
    for (r, c), cell in table.get_celld().items():
        if r == 1: cell.set_facecolor('#d9edf7')
        elif 2 <= r <= 1 + num_days:
            if dates_info[r-2]["is_weekend"]: cell.set_facecolor('#e6f2ff')
        elif r >= len(schedule_data) - 1: cell.set_facecolor('#fff2cc')
            
    plt.tight_layout()
    png_file = "full_shift_schedule.png"
    plt.savefig(png_file, dpi=200, bbox_inches='tight')
    plt.close()
    return excel_file, png_file

# ==========================================
# メインUI
# ==========================================
st.title("⛳️ PIGC シフト自動作成アプリ")
if submit_button:
    st.toast("設定を反映しました！", icon="✅")

st.write("①左のサイドバーで設定し、**反映ボタン**を押してください。  \n②各スタッフのタブで**休み**を設定し、一番下の**作成ボタン**を押してください。")

def toggle_state(staff, d_id):
    curr = st.session_state.gui_states[staff][d_id]
    if staff in cross_staffs:
        st.session_state.gui_states[staff][d_id] = (curr + 1) % 3
    else:
        st.session_state.gui_states[staff][d_id] = (curr + 1) % 2

tabs = st.tabs(all_staffs)
for i, staff in enumerate(all_staffs):
    with tabs[i]:
        st.caption(f"{staff} さんの休み設定")
        for r_idx in range((num_days // 7) + 1):
            cols = st.columns(7)
            for c_idx in range(7):
                d_idx = r_idx * 7 + c_idx
                if d_idx < num_days:
                    d_id = dates_info[d_idx]["id"]
                    state = st.session_state.gui_states[staff][d_id]
                    if staff in cross_staffs:
                        labels = {0: f"🟢両方\n{d_id}", 1: f"❌休み\n{d_id}", 2: f"🟠午前\n{d_id}"}
                    else:
                        labels = {0: f"🟢出勤\n{d_id}", 1: f"❌休み\n{d_id}"}
                    cols[c_idx].button(labels[state], key=f"b_{staff}_{d_id}", on_click=toggle_state, args=(staff, d_id), use_container_width=True)

st.divider()

if st.button("✨ この条件でシフト表を作成する", type="primary", use_container_width=True):
    status_box = st.empty()
    status_box.info("⏳ AIが最適な組み合わせを計算中...")
    
    prob = pulp.LpProblem("Shift", pulp.LpMinimize)
    x_am = {(s, d): pulp.LpVariable(f"xam_{s}_{d}", cat="Binary") for s in all_staffs for d in date_ids}
    x_pm = {(s, d): pulp.LpVariable(f"xpm_{s}_{d}", cat="Binary") for s in all_staffs for d in date_ids}
    s_am = {d: pulp.LpVariable(f"sam_{d}", cat="Integer", lowBound=0) for d in date_ids}
    s_pm = {d: pulp.LpVariable(f"spm_{d}", cat="Integer", lowBound=0) for d in date_ids}

    # 目的関数
    prob += pulp.lpSum(1000 * (s_am[d] + s_pm[d]) for d in date_ids) + pulp.lpSum(random.random() * (x_am[s,d] + x_pm[s,d]) for s in all_staffs for d in date_ids)

    # 1日の必要人数（打席2名＋社員）
    for d in date_ids:
        prob += pulp.lpSum(x_am[s, d] for s in all_staffs) + s_am[d] == 2
        prob += pulp.lpSum(x_pm[s, d] for s in all_staffs) + s_pm[d] == 2

    for s in all_staffs:
        for d in date_ids:
            state = st.session_state.gui_states[s][d]
            if state == 1: # 休み
                prob += x_am[s, d] == 0
                prob += x_pm[s, d] == 0
            elif state == 2: # 午前のみ
                prob += x_pm[s, d] == 0

            # スタッフのポジション制約
            if s in cross_staffs: prob += x_pm[s, d] <= x_am[s, d]
            elif s in staffs_am: prob += x_pm[s, d] == 0
            elif s in staffs_pm: prob += x_am[s, d] == 0
            else: prob += x_am[s, d] == 0; prob += x_pm[s, d] == 0

    # NGペアの制約
    for p1, p2 in ng_pairs:
        for d in date_ids:
            prob += x_am[p1, d] + x_am[p2, d] <= 1
            prob += x_pm[p1, d] + x_pm[p2, d] <= 1

    # 出勤日数の制約
    for s in all_staffs:
        work = pulp.lpSum(x_am[s, d] if s in staffs_am else x_pm[s, d] for d in date_ids)
        if s in min_shifts: prob += work >= min_shifts[s]
        if s in exact_shifts: prob += work == exact_shifts[s]

    # ★追加部分：最大6連勤の制約（連続する7日間のうち、出勤日数は6日以下にする）
    for s in all_staffs:
        for i in range(len(date_ids) - 6):
            if s in staffs_am:
                prob += pulp.lpSum(x_am[s, date_ids[i+j]] for j in range(7)) <= 6
            else:
                prob += pulp.lpSum(x_pm[s, date_ids[i+j]] for j in range(7)) <= 6

    solver = pulp.PULP_CBC_CMD(timeLimit=10, msg=False)
    
    # 計算結果の判定
    if prob.solve(solver) == 1:
        ex, pg = generate_files(x_am, x_pm, s_am, s_pm, status_box)
        status_box.success("🎉 シフト作成完了！")
        c1, c2 = st.columns(2)
        with c1: st.download_button("📗 Excel保存", open(ex, "rb"), "shift.xlsx", use_container_width=True)
        with c2: st.download_button("🖼 画像保存", open(pg, "rb"), "shift.png", use_container_width=True)
        st.image(pg)
    else:
        # エラーメッセージを分かりやすく修正
        status_box.error("⚠️ 条件が厳しすぎてシフトが作成できませんでした。\n\n**【よくある原因】**\n・誰かの休みが多すぎて、他の人が**「最大6連勤」**のルールを超えてしまう\n・「最低出勤日数」や「NGペア」の条件が重なり合って計算できない\n\n👉 休みや出勤日数の条件を少しゆるめて、再度お試しください。")
