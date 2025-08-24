import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta

st.set_page_config(page_title="Lab Scheduler", layout="wide")

st.title("Планировщик химико-аналитических исследований\n(по ТЗ)")

# --- Helper: sample templates

def sample_labs():
    return pd.DataFrame([
        {
            "lab_id": 1,
            "name": "Лаборатория А",
            "supported_tests": "Residue, Purity",
            "capacity_per_day": 10,
            "turnaround_days": 7,
            "storage_conditions_accepted": "+4C, -20C",
            "seasons_allowed": "all",
            "price_per_test": 120.0
        },
        {
            "lab_id": 2,
            "name": "Лаборатория Б",
            "supported_tests": "Residue",
            "capacity_per_day": 5,
            "turnaround_days": 14,
            "storage_conditions_accepted": "+4C",
            "seasons_allowed": "summer,autumn",
            "price_per_test": 100.0
        },
    ])


def sample_tests():
    return pd.DataFrame([
        {"test_id": 1, "test_name": "Residue", "duration_days": 3, "required_storage_condition": "+4C", "season_required": ""},
        {"test_id": 2, "test_name": "Purity", "duration_days": 2, "required_storage_condition": "room", "season_required": ""},
    ])


def sample_contracts():
    return pd.DataFrame([
        {
            "contract_id": 1,
            "product_name": "Препарат X",
            "active_substance": "ДВ1",
            "required_tests": "Residue;Purity",
            "sample_collection_date": (datetime.now()).date(),
            "contract_deadline": (datetime.now() + timedelta(days=30)).date(),
            "max_storage_days": 14
        }
    ])

# --- UI: uploads or templates
st.sidebar.header("Исходные Excel файлы")
labs_file = st.sidebar.file_uploader("Загрузить labs.xlsx (шаблон см. ниже)", type=["xlsx"]) 
tests_file = st.sidebar.file_uploader("Загрузить tests.xlsx", type=["xlsx"]) 
contracts_file = st.sidebar.file_uploader("Загрузить contracts.xlsx", type=["xlsx"]) 

use_sample = st.sidebar.checkbox("Использовать примерные данные (шаблоны)", value=True)

@st.cache_data
def load_df_from_upload(uploaded):
    if uploaded is None:
        return None
    try:
        df = pd.read_excel(uploaded)
        return df
    except Exception as e:
        st.error(f"Не удалось прочитать Excel: {e}")
        return None

if use_sample:
    labs_df = sample_labs()
    tests_df = sample_tests()
    contracts_df = sample_contracts()
else:
    labs_df = load_df_from_upload(labs_file) or pd.DataFrame()
    tests_df = load_df_from_upload(tests_file) or pd.DataFrame()
    contracts_df = load_df_from_upload(contracts_file) or pd.DataFrame()

col1, col2 = st.columns(2)
with col1:
    st.subheader("Лаборатории (labs.xlsx)")
    st.dataframe(labs_df)
with col2:
    st.subheader("Методы/тесты (tests.xlsx)")
    st.dataframe(tests_df)

st.subheader("Договоры / заявки (contracts.xlsx)")
st.dataframe(contracts_df)

# --- Scheduler logic

def parse_list_field(x):
    if pd.isna(x):
        return []
    if isinstance(x, (list, tuple)):
        return x
    return [s.strip() for s in str(x).split(",") if s.strip()]


def months_to_season(month):
    # simple mapping by month number to season name
    if month in [12,1,2]:
        return "winter"
    if month in [3,4,5]:
        return "spring"
    if month in [6,7,8]:
        return "summer"
    return "autumn"


def schedule_for_contract(contract_row, labs_df, tests_df):
    # contract_row: series
    required_tests = [t.strip() for t in str(contract_row.get('required_tests','')).replace(';',',').split(',') if t.strip()]
    sample_date = pd.to_datetime(contract_row.get('sample_collection_date')).date()
    deadline = pd.to_datetime(contract_row.get('contract_deadline')).date()
    max_storage = int(contract_row.get('max_storage_days', 30))

    assignments = []
    day_of_sample = sample_date
    season = months_to_season(day_of_sample.month)

    # naive greedy: for each test choose lab minimizing (start+duration) while meeting constraints
    for test_name in required_tests:
        test_row = tests_df[tests_df['test_name'].str.lower() == test_name.lower()]
        if test_row.empty:
            assignments.append({
                'test_name': test_name,
                'status': 'test not found in tests.xlsx',
            })
            continue
        test_row = test_row.iloc[0]
        duration = int(test_row.get('duration_days', 1))
        req_storage = str(test_row.get('required_storage_condition',''))
        season_required = str(test_row.get('season_required',''))

        candidates = []
        for _, lab in labs_df.iterrows():
            # supported tests
            supported = [s.strip().lower() for s in str(lab.get('supported_tests','')).replace(';',',').split(',') if s.strip()]
            if test_name.lower() not in supported:
                continue
            # season
            lab_seasons = [s.strip().lower() for s in str(lab.get('seasons_allowed','all')).replace(';',',').split(',') if s.strip()]
            if 'all' not in lab_seasons and season not in lab_seasons:
                continue
            # storage
            lab_storage = [s.strip().lower() for s in str(lab.get('storage_conditions_accepted','')).replace(';',',').split(',') if s.strip()]
            if req_storage and req_storage.lower() not in [ls.lower() for ls in lab_storage]:
                continue
            # compute earliest finish
            turnaround = int(lab.get('turnaround_days', 0))
            # we'll assume the test itself takes 'duration' days, and lab needs turnaround after receipt
            est_start = day_of_sample
            est_finish = est_start + timedelta(days=duration + turnaround)
            days_until_deadline = (deadline - est_finish).days
            price = float(lab.get('price_per_test', np.nan))
            candidates.append((lab['lab_id'], lab['name'], est_start, est_finish, days_until_deadline, price))

        if not candidates:
            assignments.append({'test_name': test_name, 'status': 'no suitable lab found'})
            continue
        # choose candidate with smallest finish date, tie-breaker lowest price
        candidates_sorted = sorted(candidates, key=lambda x: (x[3], x[5]))
        chosen = candidates_sorted[0]
        assignments.append({
            'test_name': test_name,
            'lab_id': chosen[0],
            'lab_name': chosen[1],
            'start_date': chosen[2],
            'finish_date': chosen[3],
            'days_before_deadline': chosen[4],
            'price': chosen[5],
            'status': 'scheduled' if chosen[4] >= 0 else 'will miss deadline'
        })
        # sequence: next test sample date becomes finish_date (serial execution)
        day_of_sample = chosen[3]

    return pd.DataFrame(assignments)

# Select a contract to schedule
if contracts_df is None or contracts_df.empty:
    st.warning("Нет данных по контрактам. Воспользуйтесь примерами в sidebar или загрузите contracts.xlsx")
else:
    contract_idx = st.selectbox("Выберите договор/заявку для планирования", contracts_df.index.tolist())
    contract_row = contracts_df.loc[contract_idx]
    st.markdown("### Результат планирования для договора:")
    sch = schedule_for_contract(contract_row, labs_df, tests_df)
    st.dataframe(sch)

    # Summary and export
    total_cost = sch[sch['status']=='scheduled']['price'].sum() if 'price' in sch.columns else 0
    missed = sch[sch['status'] == 'will miss deadline'].shape[0] if 'status' in sch.columns else 0
    st.write(f"Итоговая стоимость (прибл.): {total_cost}")
    if missed:
        st.error(f"Кол-во тестов, которые не уложатся в срок: {missed}")

    def to_excel_bytes(df_dict):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for name, df in df_dict.items():
                df.to_excel(writer, sheet_name=name[:31], index=False)
        return output.getvalue()

    export_bytes = to_excel_bytes({'schedule': sch, 'contract': pd.DataFrame([contract_row])})
    st.download_button("Скачать план (Excel)", data=export_bytes, file_name=f"plan_contract_{contract_row.get('contract_id')}.xlsx")

# --- Footer: templates description
st.markdown("---")
st.subheader("Формат шаблонов Excel (рекомендуемый)")
st.markdown(
"""
**labs.xlsx** (таблица):
- lab_id (int)
- name (str)
- supported_tests (str, разделитель "," или ";", например "Residue, Purity")
- capacity_per_day (int)
- turnaround_days (int) — время подготовки результатов после поступления образца
- storage_conditions_accepted (str, например "+4C; -20C")
- seasons_allowed (str, например "all" или "summer,autumn")
- price_per_test (число)

**tests.xlsx** (таблица):
- test_id
- test_name
- duration_days (сколько лабораторная часть занимает рабочих дней)
- required_storage_condition (например "+4C")
- season_required (опционально)

**contracts.xlsx** (таблица):
- contract_id
- product_name
- active_substance
- required_tests (строка, разделитель ";" или ",", например "Residue;Purity")
- sample_collection_date (дата)
- contract_deadline (дата)
- max_storage_days (int)

"""
)