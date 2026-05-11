import streamlit as st
import pandas as pd
import random
from datetime import datetime, date, time
import io
import holidays
import ast
from openpyxl import load_workbook
from openpyxl.styles import Font

# --- [0] 기본 설정 및 세션 초기화 ---
st.set_page_config(page_title="2026 Smart Scheduler", layout="wide", page_icon="📅")

def handle_upload():
    uploader = st.session_state.get('file_uploader_key')
    if uploader is not None:
        try:
            xls = pd.ExcelFile(uploader)
            # 1. 거점 정보 복구
            if 'Locations' in xls.sheet_names:
                df_loc = pd.read_excel(xls, sheet_name='Locations')
                for i, loc in enumerate(df_loc.to_dict('records')):
                    st.session_state[f"ln_{i}"] = str(loc.get('loc_name', f"LOC {i+1}"))
                    st.session_state[f"lm_{i}"] = int(loc.get('loc_min', 1))
                    c_days = loc.get('closed_days', [])
                    if isinstance(c_days, str):
                        try: c_days = ast.literal_eval(c_days)
                        except: c_days = []
                    st.session_state[f"lc_{i}"] = c_days

            # 2. 직원 정보 복구
            staff_sheet = 'Staff' if 'Staff' in xls.sheet_names else ('Sheet1' if 'Sheet1' in xls.sheet_names else None)
            if staff_sheet:
                df_staff = pd.read_excel(xls, sheet_name=staff_sheet)
                staff_list = df_staff.to_dict('records')
                st.session_state['num_staff_val'] = len(staff_list)
                for i, row in enumerate(staff_list):
                    # ✨ 버그 수정: 신버전(pure_name)과 구버전(name) 모두 인식하여 이름 복구
                    s_name = row.get('pure_name', row.get('name', f"Staff{i+1}"))
                    st.session_state[f"sn_{i}"] = str(s_name)
                    
                    st.session_state[f"af_{i}"] = str(row.get('affiliation', "本社"))
                    st.session_state[f"to_{i}"] = int(row.get('target_off', 8))
                    
                    # ✨ 시간 복구 기능 추가
                    shift_str = str(row.get('shift', '09:00-18:00'))
                    try:
                        s_st, s_et = shift_str.split('-')
                        st.session_state[f"st_{i}"] = datetime.strptime(s_st.strip(), "%H:%M").time()
                        st.session_state[f"et_{i}"] = datetime.strptime(s_et.strip(), "%H:%M").time()
                    except:
                        pass

                    # 날짜 리스트 복구
                    for key_prefix, col_name in [("or", "off_list"), ("hr", "hq_list"), ("sl", "possible_locs")]:
                        val = row.get(col_name, [])
                        if isinstance(val, str):
                            try: val = ast.literal_eval(val)
                            except: val = []
                        st.session_state[f"{key_prefix}_{i}"] = val
            st.session_state['upload_msg'] = "success"
        except Exception as e:
            st.session_state['upload_msg'] = f"error: {e}"

try: jp_holidays = holidays.Japan(years=2026)
except: jp_holidays = {}

weekdays_ko = ["월", "화", "수", "목", "금", "토", "일"]
weekdays_jp = ["月", "火", "水", "木", "金", "土", "日"]

lang_dict = {
    "日本語": {
        "co_name": "株式会社NEXTスタッフサービス", "group_name": "Enrich MR Holdings", "author": "制作: HWANG YOUNGSEON",
        "run_btn": "🚀 2026年 勤務表を生成", "download": "📥 勤務表 保存",
        "settings": "⚙️ 1. 設定", "days": "対象月", "num_staff": "全人員", 
        "loc_settings": "📍 2. 拠点設定", "loc_count": "拠点数", "loc_name": "拠点名", "loc_min": "人数", "closed_days": "休業曜日",
        "staff_settings": "👤 勤務者詳細設定", "name": "氏名", "affiliation": "所属", "hq_staff": "本社", "disp_staff": "派遣", "possible_locs": "投入可能拠点", 
        "total_off": "休日数", "off_req": "希望休日 (カレンダー)", "hq_req": "本社出勤日",
        "load_save": "💾 データ管理", "upload": "バックアップ アップロード", "backup_btn": "📥 全設定をバックアップ",
        "template_msg": "📁 様式アップロード", "result_title": "📊 結果", "hq_col": "★本社出勤★", "shortage": "⚠️不足", "loc_off": "X (休み)",
        "time_set": "⏰ 勤務時間", "start": "開始", "end": "終了", "msg_load": "✅ ロード成功", "msg_done": "✅ 生成完了"
    },
    "한국어": {
        "co_name": "株式会社NEXTスタッフ서비스", "group_name": "Enrich MR Holdings", "author": "제작자: HWANG YOUNGSEON",
        "run_btn": "🚀 2026년 통합 근무표 생성", "download": "📥 근무표 다운로드",
        "settings": "⚙️ 1. 기본 설정", "days": "대상 월", "num_staff": "전체 인원", 
        "loc_settings": "📍 2. 거점 설정", "loc_count": "거점 개수", "loc_name": "거점명", "loc_min": "인원", "closed_days": "휴무 요일",
        "staff_settings": "👤 근무자 상세 설정", "name": "성함", "affiliation": "소속", "hq_staff": "본사", "disp_staff": "파견", "possible_locs": "투입 가능 거점", 
        "total_off": "목표 휴무", "off_req": "희망 휴일", "hq_req": "본사 출사일",
        "load_save": "💾 데이터 관리", "upload": "백업 업로드", "backup_btn": "📥 모든 설정 백업하기",
        "template_msg": "📁 양식 업로드", "result_title": "📊 결과", "hq_col": "★본사출사★", "shortage": "⚠️부족", "loc_off": "X (휴무)",
        "time_set": "⏰ 시간", "start": "시작", "end": "종료", "msg_load": "✅ 로드 성공", "msg_done": "✅ 완료"
    }
}

selected_lang = st.sidebar.selectbox("🌐 Language", ["日本語", "한국어"], index=0)
L = lang_dict[selected_lang]
weekdays_active = weekdays_jp if selected_lang == "日本語" else weekdays_ko

st.markdown(f"""
    <style>
    .header-bar {{ position: fixed; top: 0; left: 0; width: 100%; background: #1E3A8A; color: white; padding: 10px 20px; display: flex; justify-content: space-between; align-items: center; z-index: 1000; }}
    .stApp {{ margin-top: 60px; }}
    .stButton>button {{ width: 100%; background-color: #1E3A8A !important; color: white !important; font-weight: bold; border-radius: 8px; height: 3.5em; }}
    </style>
    <div class="header-bar"><span>🏢 {L['co_name']} ({L['group_name']})</span><span>{L['author']}</span></div>
    """, unsafe_allow_html=True)

# --- [3] 사이드바 ---
st.sidebar.header(L["load_save"])
st.sidebar.file_uploader(L["upload"], type=["xlsx"], key="file_uploader_key", on_change=handle_upload)
template_file = st.sidebar.file_uploader(L["template_msg"], type=["xlsx"])

target_month = st.sidebar.number_input(L["days"], 1, 12, 1)
days_in_month = 28 if target_month == 2 else (30 if target_month in [4,6,9,11] else 31)

def format_day(d):
    curr_d = date(2026, target_month, d)
    return f"{d} ({weekdays_active[curr_d.weekday()]})"

num_locations = st.sidebar.number_input(L["loc_count"], 1, 10, value=st.session_state.get('loc_count_val', 4), key='loc_count_val')
location_names = []; location_configs = {}
for i in range(num_locations):
    with st.sidebar.expander(f"📍 {L['loc_name']} {i+1}", expanded=False):
        l_name = st.text_input(L["loc_name"], key=f"ln_{i}", value=st.session_state.get(f"ln_{i}", f"LOC {i+1}"))
        l_min = st.number_input(L["loc_min"], 0, 10, key=f"lm_{i}", value=st.session_state.get(f"lm_{i}", 1))
        
        saved_closed = st.session_state.get(f"lc_{i}", [])
        default_idx = [weekdays_ko.index(d) for d in saved_closed if d in weekdays_ko]
        l_closed_display = st.multiselect(L["closed_days"], weekdays_active, default=[weekdays_active[idx] for idx in default_idx], key=f"lc_disp_{i}")
        
        l_closed = [weekdays_ko[weekdays_active.index(d)] for d in l_closed_display]
        st.session_state[f"lc_{i}"] = l_closed
        if l_name: location_names.append(l_name); location_configs[l_name] = {"min": l_min, "closed": l_closed}

num_staff = st.sidebar.slider(L["num_staff"], 1, 30, value=st.session_state.get('num_staff_val', 10), key='num_staff_val')
st.header(L["staff_settings"])
staff_data = []
c1, c2 = st.columns(2)
affil_options = [L["hq_staff"], L["disp_staff"]]

for i in range(num_staff):
    curr_sn = st.session_state.get(f"sn_{i}", f"Staff{i+1}")
    curr_af = st.session_state.get(f"af_{i}", affil_options[0])
    curr_to = st.session_state.get(f"to_{i}", 8)
    curr_or = st.session_state.get(f"or_{i}", [])
    curr_hr = st.session_state.get(f"hr_{i}", [])
    curr_sl = st.session_state.get(f"sl_{i}", [])
    
    if curr_af not in affil_options:
        curr_af = L["hq_staff"] if curr_af in ["本社", "본사"] else L["disp_staff"]
        
    curr_or = [d for d in curr_or if 1 <= d <= days_in_month]
    curr_hr = [d for d in curr_hr if 1 <= d <= days_in_month]

    with (c1 if i % 2 == 0 else c2).expander(f"👤 {curr_sn}", expanded=False):
        col_n, col_a = st.columns([2, 1])
        s_name = col_n.text_input(L["name"], value=curr_sn, key=f"sn_{i}", label_visibility="collapsed")
        
        af_idx = affil_options.index(curr_af) if curr_af in affil_options else 0
        s_affil = col_a.selectbox(L["affiliation"], affil_options, index=af_idx, key=f"af_{i}", label_visibility="collapsed")
        
        tc1, tc2 = st.columns(2)
        # ✨ 백업된 시간 설정 불러오기
        st_t = tc1.time_input(L["start"], value=st.session_state.get(f"st_{i}", time(9, 0)), key=f"st_{i}")
        et_t = tc2.time_input(L["end"], value=st.session_state.get(f"et_{i}", time(18, 0)), key=f"et_{i}")
        shift_str = f"{st_t.strftime('%H:%M')}-{et_t.strftime('%H:%M')}"
        
        s_locs = st.multiselect(L["possible_locs"], location_names, default=[loc for loc in curr_sl if loc in location_names], key=f"sl_{i}")
        t_off = st.number_input(L["total_off"], 0, 20, value=curr_to, key=f"to_{i}")
        off_l = st.multiselect(L["off_req"], range(1, days_in_month + 1), default=curr_or, key=f"or_{i}", format_func=format_day)
        hq_l = st.multiselect(L["hq_req"], range(1, days_in_month + 1), default=curr_hr, key=f"hr_{i}", format_func=format_day)
        
        staff_data.append({"name": f"{s_name} ({s_affil})", "pure_name": s_name, "affiliation": s_affil, "shift": shift_str, "possible_locs": s_locs, "target_off": t_off, "off_list": off_l, "hq_list": hq_l})

if staff_data:
    out_cfg = io.BytesIO()
    with pd.ExcelWriter(out_cfg, engine='openpyxl') as writer:
        pd.DataFrame(staff_data).drop(columns=['name']).to_excel(writer, index=False, sheet_name='Staff')
        loc_backup = [{"loc_name": k, "loc_min": v['min'], "closed_days": str(v['closed'])} for k, v in location_configs.items()]
        pd.DataFrame(loc_backup).to_excel(writer, index=False, sheet_name='Locations')
    st.sidebar.download_button(L["backup_btn"], out_cfg.getvalue(), f"full_backup.xlsx")

if st.button(L["run_btn"]):
    schedule_results = []; holiday_list = []
    for d in range(1, days_in_month + 1):
        curr_d = date(2026, target_month, d); curr_weekday_ko = weekdays_ko[curr_d.weekday()]
        if curr_d in jp_holidays or curr_d.weekday() >= 5: holiday_list.append(d)
        res = {"Date": f"{d} ({weekdays_active[curr_d.weekday()]})"}; assigned = []
        hq = [f"{s['pure_name']}({s['shift']})" for s in staff_data if d in s['hq_list']]
        res[L["hq_col"]] = ", ".join(hq) if hq else "-"; assigned.extend([s['pure_name'] for s in staff_data if d in s['hq_list']])
        avail = [s for s in staff_data if s['pure_name'] not in assigned and d not in s['off_list']]
        random.shuffle(avail)
        for loc, config in location_configs.items():
            if curr_weekday_ko in config['closed']: res[loc] = L["loc_off"]
            else:
                sel = [s for s in avail if loc in s['possible_locs']][:config['min']]
                res[loc] = ", ".join([f"{s['pure_name']}({s['shift']})" for s in sel]) if sel else L["shortage"]
                for s in sel: assigned.append(s['pure_name']); avail.remove(s)
        schedule_results.append(res)
    df_res = pd.DataFrame(schedule_results)
    st.dataframe(df_res.style.apply(lambda row: ['color: red' if int(row['Date'].split()[0]) in holiday_list else '' for _ in row], axis=1), use_container_width=True)
    out_res = io.BytesIO()
    if template_file:
        wb = load_workbook(template_file); ws = wb.active
        for i, row in enumerate(schedule_results):
            for j, val in enumerate(row.values()): ws.cell(row=2+i, column=1+j, value=val)
        wb.save(out_res)
    else:
        with pd.ExcelWriter(out_res, engine='openpyxl') as writer: df_res.to_excel(writer, index=False)
    st.success(L["msg_done"])
    st.download_button(L["download"], out_res.getvalue(), f"Schedule_{target_month}.xlsx")
