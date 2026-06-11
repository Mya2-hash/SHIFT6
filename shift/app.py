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
    if st.session_state.file_uploader_key is not None:
        try:
            xls = pd.ExcelFile(st.session_state.file_uploader_key)
            
            # 1. 근무지 정보 복구
            if 'Locations' in xls.sheet_names:
                df_loc = pd.read_excel(xls, sheet_name='Locations')
                loc_list = df_loc.to_dict('records')
                st.session_state['loc_count_val'] = len(loc_list)
                for i, loc in enumerate(loc_list):
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
                    st.session_state[f"sn_{i}"] = str(row.get('name', f"Staff{i+1}"))
                    st.session_state[f"af_{i}"] = str(row.get('affiliation', "本社" if selected_lang == "日本語" else "본사"))
                    st.session_state[f"to_{i}"] = int(row.get('target_off', 8))
                    
                    # 리스트 형태의 날짜 데이터 복구 (희망휴일, 본사출사일 등)
                    for key, col in [("or_{i}", 'off_list'), ("hr_{i}", 'hq_list'), ("sl_{i}", 'possible_locs')]:
                        val = row.get(col, [])
                        if isinstance(val, str):
                            try: val = ast.literal_eval(val)
                            except: val = []
                        st.session_state[key.format(i=i)] = val
                        
                    # 🆕 거점별 지정 출근일 복구
                    fl_val = row.get('fixed_locs', "{}")
                    if isinstance(fl_val, str):
                        try: fl_val = ast.literal_eval(fl_val)
                        except: fl_val = {}
                    st.session_state[f"fl_dict_{i}"] = fl_val if isinstance(fl_val, dict) else {}
                        
            st.session_state['upload_msg'] = "success"
        except Exception as e:
            st.session_state['upload_msg'] = f"error: {e}"

if 'staff_configs' not in st.session_state: st.session_state['staff_configs'] = []
try: jp_holidays = holidays.Japan(years=2026)
except: jp_holidays = {}

weekdays_ko = ["월", "화", "수", "목", "금", "토", "일"]
weekdays_jp = ["月", "火", "水", "木", "金", "土", "日"]

# --- [1] 다국어(번역) 설정 ---
lang_dict = {
    "日本語": {
        "co_name": "株式会社NEXTスタッフサービス", "group_name": "Enrich MR Holdings", "author": "制作: HWANG YOUNGSEON",
        "run_btn": "🚀 2026年 勤務表を生成", "download": "📥 勤務表 エクセル保存",
        "settings": "⚙️ 1. 設定", "days": "対象月 (1~12月)", "num_staff": "全人員", 
        "loc_settings": "📍 2. 拠点別運営設定", "loc_count": "拠点数", "loc_name": "拠点名", "loc_min": "必要人数", "closed_days": "休業日を選択",
        "staff_settings": "👤 勤務者詳細設定", "name": "氏名", "affiliation": "所属", "hq_staff": "本社", "disp_staff": "派遣", "possible_locs": "投入可能地域", 
        "fixed_loc": "指定出勤日", "fixed_loc_msg": "📍 拠点別の指定出勤日 (オプション)", # 🆕 추가된 번역
        "total_off": "今月の休日数", "off_req": "希望休日 (曜日確認)", "hq_req": "本社出勤日",
        "load_save": "💾 データ管理", "upload": "バックアップ エクセルアップロード", "backup_btn": "📥 全設定を保存する",
        "template_msg": "📁 テンプレート(様式) アップロード",
        "result_title": "📊 生成結果", "hq_col": "★本社出勤★", "shortage": "⚠️不足", "loc_off": "X (休み)",
        "time_set": "⏰ 勤務時間", "start": "開始", "end": "終了", "name_ph": "名前を入力",
        "msg_load": "✅ データの読み込みに成功しました！", "msg_err": "❌ エラー", "msg_done": "✅ 作成が完了しました！"
    },
    "한국어": {
        "co_name": "株式会社NEXTスタッフ서비스", "group_name": "Enrich MR Holdings", "author": "제작자: HWANG YOUNGSEON",
        "run_btn": "🚀 2026년 통합 근무표 생성", "download": "📥 근무표 엑셀 다운로드",
        "settings": "⚙️ 1. 기본 환경 설정", "days": "대상 월 (1~12월)", "num_staff": "전체 인원", 
        "loc_settings": "📍 2. 거점별 운영 설정", "loc_count": "거점 개수", "loc_name": "거점명", "loc_min": "필요 인원", "closed_days": "휴무 요일 선택",
        "staff_settings": "👤 근무자 상세 설정", "name": "이름 수정", "affiliation": "소속", "hq_staff": "본사", "disp_staff": "파견", "possible_locs": "투입 가능 지역", 
        "fixed_loc": "지정 출사일", "fixed_loc_msg": "📍 거점별 지정 출근일 (선택사항)", # 🆕 추가된 번역
        "total_off": "목표 휴무", "off_req": "희망 휴일 (요일 확인)", "hq_req": "본사 출사일",
        "load_save": "💾 데이터 관리", "upload": "백업 엑셀 업로드", "backup_btn": "📥 모든 설정 백업하기",
        "template_msg": "📁 엑셀 양식(Template) 업로드",
        "result_title": "📊 생성 결과", "hq_col": "★본사출사★", "shortage": "⚠️부족", "loc_off": "X (휴무)",
        "time_set": "⏰ 시간 설정", "start": "시작", "end": "종료", "name_ph": "성함을 입력하세요",
        "msg_load": "✅ 데이터 로드 성공!", "msg_err": "❌ 로드 오류", "msg_done": "✅ 생성이 완료되었습니다!"
    }
}

selected_lang = st.sidebar.selectbox("🌐 言語 / 언어", ["日本語", "한국어"], index=0)
L = lang_dict[selected_lang]
weekdays_active = weekdays_jp if selected_lang == "日本語" else weekdays_ko

# --- [2] 디자인 (CSS) ---
st.markdown(f"""
    <style>
    .header-bar {{ position: fixed; top: 0; left: 0; width: 100%; background: #1E3A8A; color: white; padding: 10px 20px; display: flex; justify-content: space-between; align-items: center; z-index: 1000; }}
    .stApp {{ margin-top: 60px; }}
    .stButton>button {{ width: 100%; background-color: #1E3A8A !important; color: white !important; font-weight: bold; border-radius: 8px; height: 3.5em; }}
    </style>
    <div class="header-bar"><span>🏢 {L['co_name']} ({L['group_name']})</span><span>{L['author']}</span></div>
    """, unsafe_allow_html=True)

# --- [3] 사이드바: 설정 ---
st.sidebar.header(L["load_save"])
st.sidebar.file_uploader(L["upload"], type=["xlsx"], key="file_uploader_key", on_change=handle_upload)

if 'upload_msg' in st.session_state:
    if st.session_state['upload_msg'] == "success": st.sidebar.success(L["msg_load"])
    else: st.sidebar.error(f"{L['msg_err']} {st.session_state['upload_msg']}")
    del st.session_state['upload_msg']

template_file = st.sidebar.file_uploader(L["template_msg"], type=["xlsx"])

st.sidebar.divider()
target_month = st.sidebar.number_input(L["days"], 1, 12, 1)

# 해당 월의 일수 계산
days_in_month = 28 if target_month == 2 else (30 if target_month in [4,6,9,11] else 31)

def format_day(d):
    curr_d = date(2026, target_month, d)
    wd = weekdays_active[curr_d.weekday()]
    return f"{d}日 ({wd})" if selected_lang == "日本語" else f"{d}일 ({wd})"

st.sidebar.subheader(L["loc_settings"])
num_locations = st.sidebar.number_input(L["loc_count"], 1, 10, value=st.session_state.get('loc_count_val', 4), key='loc_count_val')
location_names = []; location_configs = {}
for i in range(num_locations):
    with st.sidebar.expander(f"📍 {L['loc_name']} {i+1}", expanded=False):
        l_name = st.text_input(L["loc_name"], key=f"ln_{i}")
        l_min = st.number_input(L["loc_min"], 0, 10, key=f"lm_{i}")
        
        saved_closed = st.session_state.get(f"lc_{i}", [])
        default_idx = [weekdays_ko.index(d) for d in saved_closed if d in weekdays_ko]
        default_display = [weekdays_active[idx] for idx in default_idx]
        
        l_closed_display = st.multiselect(L["closed_days"], weekdays_active, default=default_display, key=f"lc_disp_{i}")
        l_closed = [weekdays_ko[weekdays_active.index(d)] for d in l_closed_display]
        st.session_state[f"lc_{i}"] = l_closed
        
        if l_name:
            location_names.append(l_name)
            location_configs[l_name] = {"min": l_min, "closed": l_closed}

# --- [4] 근무자 설정 ---
num_staff = st.sidebar.slider(L["num_staff"], 1, 30, value=st.session_state.get('num_staff_val', 10), key='num_staff_val')
st.header(L["staff_settings"])
staff_data = []
c1, c2 = st.columns(2)

affil_options = [L["hq_staff"], L["disp_staff"]]

for i in range(num_staff):
    if f"sn_{i}" not in st.session_state: st.session_state[f"sn_{i}"] = f"Staff{i+1}"
    
    with (c1 if i % 2 == 0 else c2).expander(f"👤 {st.session_state[f'sn_{i}']}", expanded=False):
        col_n, col_a = st.columns([2, 1])
        s_name = col_n.text_input(L["name"], key=f"sn_{i}", label_visibility="collapsed", placeholder=L["name_ph"])
        
        saved_affil = st.session_state.get(f"af_{i}", affil_options[0])
        s_affil = col_a.selectbox(L["affiliation"], affil_options, index=affil_options.index(saved_affil) if saved_affil in affil_options else 0, key=f"af_{i}", label_visibility="collapsed")
        
        st.write(f"**{L['time_set']}**")
        tc1, tc2 = st.columns(2)
        st_t = tc1.time_input(L["start"], value=time(9, 0), key=f"st_{i}")
        et_t = tc2.time_input(L["end"], value=time(18, 0), key=f"et_{i}")
        shift_str = f"{st_t.strftime('%H:%M')}-{et_t.strftime('%H:%M')}"
        
        s_locs = st.multiselect(L["possible_locs"], location_names, key=f"sl_{i}")
        
        # 🆕 거점별 지정 출근일 설정 UI 생성
        fixed_locs = {}
        if s_locs:
            st.caption(L["fixed_loc_msg"])
            for loc in s_locs:
                loc_idx = location_names.index(loc)
                saved_fl = st.session_state.get(f"fl_dict_{i}", {})
                default_dates = [d for d in saved_fl.get(loc, []) if 1 <= d <= days_in_month]
                
                f_dates = st.multiselect(f" - {loc} {L['fixed_loc']}", range(1, days_in_month + 1), default=default_dates, key=f"fl_{i}_{loc_idx}", format_func=format_day)
                if f_dates:
                    fixed_locs[loc] = f_dates

        t_off = st.number_input(L["total_off"], 0, 20, key=f"to_{i}")
        off_l = st.multiselect(L["off_req"], range(1, days_in_month + 1), key=f"or_{i}", format_func=format_day)
        hq_l = st.multiselect(L["hq_req"], range(1, days_in_month + 1), key=f"hr_{i}", format_func=format_day)
        
        staff_data.append({
            "name": f"{s_name} ({s_affil})", 
            "pure_name": s_name,
            "affiliation": s_affil,
            "shift": shift_str, 
            "possible_locs": s_locs, 
            "fixed_locs": fixed_locs, # 🆕 딕셔너리로 저장
            "target_off": t_off, 
            "off_list": off_l, 
            "hq_list": hq_l
        })

# 통합 백업 다운로드
if staff_data:
    out_cfg = io.BytesIO()
    with pd.ExcelWriter(out_cfg, engine='openpyxl') as writer:
        pd.DataFrame(staff_data).drop(columns=['name']).to_excel(writer, index=False, sheet_name='Staff') 
        loc_backup = [{"loc_name": k, "loc_min": v['min'], "closed_days": str(v['closed'])} for k, v in location_configs.items()]
        pd.DataFrame(loc_backup).to_excel(writer, index=False, sheet_name='Locations')
    st.sidebar.download_button(L["backup_btn"], out_cfg.getvalue(), f"full_backup.xlsx")

# --- [5] 근무표 생성 ---
st.divider()
if st.button(L["run_btn"]):
    schedule_results = []; holiday_list = []
    
    for d in range(1, days_in_month + 1):
        curr_d = date(2026, target_month, d)
        curr_weekday_ko = weekdays_ko[curr_d.weekday()]
        curr_weekday_disp = weekdays_active[curr_d.weekday()]
        is_h = curr_d in jp_holidays or curr_d.weekday() >= 5
        if is_h: holiday_list.append(d)
        
        res = {"Date": f"{d} ({curr_weekday_disp})"}; assigned = []
        
        # 1. 본사 최우선 배정
        hq = [f"{s['pure_name']}({s['shift']})" for s in staff_data if d in s['hq_list']]
        res[L["hq_col"]] = ", ".join(hq) if hq else "-"
        assigned.extend([s['pure_name'] for s in staff_data if d in s['hq_list']])
        
        # 2. 🆕 거점 지정 출근 우선 배정 (본사 배정과 동급)
        fixed_assignments = {loc: [] for loc in location_configs.keys()}
        for s in staff_data:
            if s['pure_name'] not in assigned and d not in s['off_list']:
                for loc, dates in s.get('fixed_locs', {}).items():
                    if d in dates and loc in location_configs:
                        fixed_assignments[loc].append(s)
                        assigned.append(s['pure_name'])
                        break # 하루에 한 명은 한 거점에만 배정되도록
        
        # 3. 나머지 거점 필요 인원에 따른 랜덤 배정
        avail = [s for s in staff_data if s['pure_name'] not in assigned and d not in s['off_list']]
        random.shuffle(avail)
        
        for loc, config in location_configs.items():
            if curr_weekday_ko in config['closed']:
                res[loc] = L["loc_off"]
            else:
                needed = config['min']
                
                # 이미 지정 출근으로 채워진 인원 계산
                already_assigned = fixed_assignments.get(loc, [])
                still_needed = max(0, needed - len(already_assigned))
                
                # 부족한 만큼만 랜덤하게 투입
                sel = [s for s in avail if loc in s['possible_locs']][:still_needed]
                
                # 최종 투입 인원 = 지정 인원 + 랜덤 투입 인원
                final_loc_staff = already_assigned + sel
                res[loc] = ", ".join([f"{s['pure_name']}({s['shift']})" for s in final_loc_staff]) if final_loc_staff else L["shortage"]
                
                for s in sel: 
                    assigned.append(s['pure_name'])
                    avail.remove(s)
                    
        schedule_results.append(res)

    df_res = pd.DataFrame(schedule_results)
    st.subheader(L["result_title"])
    st.dataframe(df_res.style.apply(lambda row: ['color: red' if int(row['Date'].split()[0]) in holiday_list else '' for _ in row], axis=1), use_container_width=True)
    
    # 엑셀 다운로드
    out_res = io.BytesIO()
    if template_file:
        wb = load_workbook(template_file); ws = wb.active
        for i, row in enumerate(schedule_results):
            for j, val in enumerate(row.values()): ws.cell(row=2+i, column=1+j, value=val)
        wb.save(out_res)
    else:
        with pd.ExcelWriter(out_res, engine='openpyxl') as writer: df_res.to_excel(writer, index=False)
    
    st.success(L["msg_done"])
    st.download_button(L["download"], out_res.getvalue(), f"Schedule_{target_month}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown(f'<div class="footer">© 2026 {L["co_name"]}. {L["author"]}</div>', unsafe_allow_html=True)
