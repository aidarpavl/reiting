import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import requests
import json

# ==================== КОНФИГУРАЦИЯ ====================
st.set_page_config(
    page_title="Мұғалімдер рейтингі",
    page_icon="📊",
    layout="wide"
)

# Apps Script URL (сіздің URL-іңізбен ауыстырыңыз)
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxcg32lZRFEJMb1sgFEEXK9jTww1CUUOXzkwqdyDYzObGjVgVs6LKDiAvfx9zA5JQN3hw/exec"

# ==================== ДЕРЕКТЕР ҚҰРЫЛЫМЫ ====================
columns = [
    "ФИО", "Предмет", "Коэф",
    "Орг.демалыс", "Вед.рейтинг", "Флеш-моб", "Внекл.мероп", "Дежурство",
    "Конк(гор)", "Конк(обл)", "Конк(респ)", "Внекл.мероп2", "СМИ", "Обобщ.опыт", "Публикац",
    "Сдача отчетов", "Посещаемость", "Кл.уголок", "ПДД", "Питание", "Журнал"
]

# Демо деректер
demo_data = [
    ["АХШАЛОВА", "Математика", 1.1, 1, 2, 3, 5, 5, 2, 3, 3, 1, 3, 3, 3, 3, 3, 3, 3, 3, 3],
    ["АЗЫМБАЕВА", "Химия", 1, 4, 5, 8, 2, 5, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2],
    ["АЙТБАЙ", "Математика", 1.1, 4, 7, 9, 3, 5, 1, 3, 3, 1, 3, 3, 3, 3, 3, 3, 3, 3, 3],
    ["АЛЬМЕНОВА", "Русский язык", 0.9, 4, 8, 5, 1, 5, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    ["АЛЬНАЗИРОВА", "Ин яз", 1.1, 4, 7, 3, 4, 5, 2, 4, 4, 2, 4, 4, 4, 4, 4, 4, 4, 4, 4],
    ["АРЫНГАЗИНА", "Каз яз", 0.9, 4, 7, 4, 5, 5, 3, 5, 5, 1, 5, 5, 5, 5, 5, 5, 5, 5, 5],
    ["ИЛЬЯСОВА", "Ин яз", 1, 4, 6, 7, 2, 5, 2, 2, 2, 3, 2, 2, 2, 2, 2, 2, 2, 2, 2],
    ["ИТЕМГЕНОВ", "история", 0.8, 4, 5, 10, 1, 5, 2, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    ["КАБЫЛОВА", "Каз яз", 0.9, 4, 1, 8, 2, 5, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2],
    ["КАЗАНГАПОВ", "Ин яз", 1.1, 4, 9, 11, 1, 5, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    ["КАЙДАРОВА", "География", 0.6, 4, 3, 3, 1, 5, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    ["КОЖАГЕЛЬДИНОВА", "Математика", 1.1, 4, 4, 10, 1, 5, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    ["МОМЫНОВ", "Информатика", 1, 4, 3, 8, 2, 5, 1, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2]
]

# ==================== ЕСЕПТЕУ ФУНКЦИЯЛАРЫ ====================
def calculate_extracurricular(row):
    """Внеклассная работа итогын есептеу"""
    return row['Орг.демалыс'] + row['Вед.рейтинг'] + row['Флеш-моб'] + row['Внекл.мероп'] + row['Дежурство']

def calculate_methodical(row):
    """Методическая деятельность итогын есептеу"""
    return (row['Конк(гор)'] + row['Конк(обл)'] + row['Конк(респ)'] + 
            row['Внекл.мероп2'] + row['СМИ'] + row['Обобщ.опыт'] + row['Публикац'])

def calculate_admin(row):
    """Администрация итогын есептеу"""
    return (row['Сдача отчетов'] + row['Посещаемость'] + row['Кл.уголок'] + 
            row['ПДД'] + row['Питание'] + row['Журнал'])

def calculate_total(row):
    """Жалпы итогты есептеу"""
    return calculate_extracurricular(row) + calculate_methodical(row) + calculate_admin(row)

def process_data(df):
    """Деректерді өңдеу және қосымша бағандар қосу"""
    df = df.copy()
    df['Внекл итог'] = df.apply(calculate_extracurricular, axis=1)
    df['Метод итог'] = df.apply(calculate_methodical, axis=1)
    df['Админ итог'] = df.apply(calculate_admin, axis=1)
    df['Жалпы итог'] = df.apply(calculate_total, axis=1)
    return df

def calculate_percentages(df):
    """Пайыздарды есептеу"""
    min_score = df['Жалпы итог'].min()
    max_score = df['Жалпы итог'].max()
    range_score = max_score - min_score
    
    if range_score == 0:
        df['Пайыз'] = 100.0
    else:
        df['Пайыз'] = ((df['Жалпы итог'] - min_score) / range_score) * 100
    return df

# ==================== GOOGLE SHEETS-ТЕН ОҚУ ====================
@st.cache_data(ttl=300)
def load_from_google_sheets():
    """Google Apps Script арқылы деректерді оқу"""
    try:
        response = requests.get(APPS_SCRIPT_URL, timeout=30)
        if response.status_code == 200:
            data = response.json()
            if data and len(data) > 1:
                # Бірінші жолды атаулар ретінде алып, қалғанын деректер ретінде
                headers = data[0] if isinstance(data[0], list) else columns
                rows = data[1:] if len(data) > 1 else []
                
                df = pd.DataFrame(rows, columns=headers[:len(columns)])
                return df
    except Exception as e:
        st.error(f"Google Sheets-тен оқу қатесі: {str(e)}")
    return None

# ==================== ИНТЕРФЕЙС ====================
st.title("📊 Мұғалімдердің кешенді рейтингі")
st.caption("Өрлеу бағдарламасы бойынша - Класс жетекшілерді бағалау жүйесі")

# Бақылау батырмалары
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    if st.button("🔄 Google Sheets-тен оқу", use_container_width=True):
        with st.spinner("Деректер жүктелуде..."):
            df = load_from_google_sheets()
            if df is not None and not df.empty:
                st.session_state.df = df
                st.success(f"✅ {len(df)} мұғалім сәтті жүктелді!")
            else:
                st.warning("⚠️ Google Sheets-тен оқу мүмкін болмады. Демо-деректер қолданылуда.")
                st.session_state.df = pd.DataFrame(demo_data, columns=columns)

with col2:
    if st.button("📈 Пайыздарды есептеу", use_container_width=True):
        if 'df' in st.session_state:
            st.session_state.df = calculate_percentages(st.session_state.df)
            st.success("✅ Пайыздар есептелді! Кестенің соңғы бағанын қараңыз.")
            st.rerun()

with col3:
    if st.button("📎 Excel сақтау", use_container_width=True):
        if 'df' in st.session_state:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.df.to_excel(writer, index=False, sheet_name='Мұғалімдер рейтингі')
            st.download_button(
                label="📥 Excel жүктеп алу",
                data=output.getvalue(),
                file_name=f"teacher_rating_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Алдымен деректерді жүктеңіз!")

with col4:
    if st.button("📊 Графикті жаңарту", use_container_width=True):
        st.rerun()

with col5:
    if st.button("🔄 Барлығын есептеу", use_container_width=True):
        if 'df' in st.session_state:
            st.session_state.df = process_data(st.session_state.df)
            st.success("✅ Барлық есептеулер жаңартылды!")

# Бастапқы деректерді жүктеу
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame(demo_data, columns=columns)
    st.session_state.df = process_data(st.session_state.df)

# Іздеу және сүзгі
search = st.text_input("🔍 Мұғалім іздеу (аты/пәні бойынша)", placeholder="Атын немесе пәнін жазыңыз...")
filtered_df = st.session_state.df.copy()
if search:
    filtered_df = filtered_df[
        filtered_df['ФИО'].str.contains(search, case=False, na=False) | 
        filtered_df['Предмет'].str.contains(search, case=False, na=False)
    ]
st.info(f"👩‍🏫 Барлығы: **{len(filtered_df)}** мұғалім")

# Негізгі кесте
st.subheader("📋 Мұғалімдер кестесі")

# Бағандарды таңдау (барлық бағандарды көрсету)
display_columns = ['ФИО', 'Предмет', 'Коэф', 'Внекл итог', 'Метод итог', 'Админ итог', 'Жалпы итог']
if 'Пайыз' in filtered_df.columns:
    display_columns.append('Пайыз')

st.dataframe(
    filtered_df[display_columns] if all(c in filtered_df.columns for c in display_columns) else filtered_df,
    use_container_width=True,
    height=400,
    column_config={
        "ФИО": st.column_config.TextColumn("ФИО", width="medium"),
        "Предмет": st.column_config.TextColumn("Предмет", width="small"),
        "Коэф": st.column_config.NumberColumn("Коэф", format="%.1f"),
        "Внекл итог": st.column_config.NumberColumn("Внекл итог", format="%.1f"),
        "Метод итог": st.column_config.NumberColumn("Метод итог", format="%.1f"),
        "Админ итог": st.column_config.NumberColumn("Админ итог", format="%.0f"),
        "Жалпы итог": st.column_config.NumberColumn("Жалпы итог", format="%.1f"),
        "Пайыз": st.column_config.NumberColumn("Пайыз (%)", format="%.1f%%"),
    }
)

# Үздік және төмен көрсеткішті мұғалімдер
col1, col2 = st.columns(2)

with col1:
    st.subheader("🏆 Үздік мұғалімдер")
    top3 = filtered_df.nlargest(3, 'Жалпы итог')[['ФИО', 'Предмет', 'Жалпы итог']]
    for i, (_, row) in enumerate(top3.iterrows(), 1):
        medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉"
        st.metric(f"{medal} {i}-орын", row['ФИО'], f"{row['Жалпы итог']:.1f} балл")

with col2:
    st.subheader("⚠️ Көмек қажет мұғалімдер")
    bottom3 = filtered_df.nsmallest(3, 'Жалпы итог')[['ФИО', 'Предмет', 'Жалпы итог']]
    for i, (_, row) in enumerate(bottom3.iterrows(), 1):
        st.metric(f"{i}-ең төмен", row['ФИО'], f"{row['Жалпы итог']:.1f} балл", delta_color="inverse")

# График
st.subheader("📊 Мұғалімдердің жалпы итогтары")
fig = px.bar(
    filtered_df,
    x='ФИО',
    y='Жалпы итог',
    color='Жалпы итог',
    color_continuous_scale=['red', 'yellow', 'green'],
    title="Мұғалімдердің рейтингі",
    labels={'ФИО': 'Мұғалім', 'Жалпы итог': 'Жалпы итог (балл)'},
    text='Жалпы итог'
)
fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')
fig.update_layout(height=500, showlegend=False)
st.plotly_chart(fig, use_container_width=True)

# Статистика
st.subheader("📈 Статистикалық мәліметтер")
stat_cols = st.columns(4)
with stat_cols[0]:
    st.metric("📊 Орташа итог", f"{filtered_df['Жалпы итог'].mean():.1f}")
with stat_cols[1]:
    st.metric("🏆 Ең жоғары көрсеткіш", f"{filtered_df['Жалпы итог'].max():.1f}")
with stat_cols[2]:
    st.metric("⚠️ Ең төмен көрсеткіш", f"{filtered_df['Жалпы итог'].min():.1f}")
with stat_cols[3]:
    positive = len(filtered_df[filtered_df['Жалпы итог'] > 0])
    st.metric("✅ Оң нәтижелілер", f"{positive}/{len(filtered_df)}")

# Көмек ұсыныстары
st.subheader("💡 Төмен балл алған мұғалімдерге көмек ұсыныстары")
with st.expander("📋 Көмек шаралары", expanded=True):
    help_items = [
        "📚 Әдістемелік бірлестік отырыстарына жүйелі түрде қатысу",
        "👥 Тәжірибелі әріптестермен жұптасып сабақ өткізу",
        "📊 Жеке даму жоспарын құру (3-6 айға)",
        "🏫 Ішкі оқыту семинарларына қатысу",
        "📝 Портфолио жүргізу және жетістіктерді тіркеу",
        "🎯 Әдістемелік көмек көрсету үшін тәлімгер тағайындау"
    ]
    for item in help_items:
        st.write(f"- {item}")

# Мұғалімдер тізімін өңдеу
st.subheader("✏️ Мұғалімдерді басқару")

# Өшіру үшін таңдау
teacher_to_delete = st.selectbox("Өшіретін мұғалімді таңдаңыз:", [''] + list(st.session_state.df['ФИО'].unique()))
if teacher_to_delete and st.button("🗑 Мұғалімді өшіру", type="secondary"):
    st.session_state.df = st.session_state.df[st.session_state.df['ФИО'] != teacher_to_delete]
    st.session_state.df = process_data(st.session_state.df)
    st.success(f"✅ {teacher_to_delete} мұғалімі өшірілді!")
    st.rerun()

# Жаңа мұғалім қосу
with st.expander("➕ Жаңа мұғалім қосу", expanded=False):
    with st.form("add_teacher_form"):
        new_fio = st.text_input("ФИО")
        new_subject = st.selectbox("Предмет", ["Математика", "Қазақ тілі", "Орыс тілі", "Ағылшын тілі", "Физика", "Химия", "Биология", "Тарих", "География", "Информатика"])
        new_coef = st.number_input("Коэффицент", min_value=0.5, max_value=2.0, value=1.0, step=0.1)
        
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            new_org = st.number_input("Орг.демалыс", min_value=0, max_value=5, value=0)
            new_rating = st.number_input("Вед.рейтинг", min_value=0, max_value=5, value=0)
            new_flash = st.number_input("Флеш-моб", min_value=0, max_value=5, value=0)
            new_vnekl = st.number_input("Внекл.мероп", min_value=0, max_value=5, value=0)
            new_desh = st.number_input("Дежурство", min_value=0, max_value=5, value=0)
        
        if st.form_submit_button("➕ Қосу"):
            if new_fio:
                new_row = pd.DataFrame([[
                    new_fio, new_subject, new_coef,
                    new_org, new_rating, new_flash, new_vnekl, new_desh,
                    0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                ]], columns=columns)
                st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
                st.session_state.df = process_data(st.session_state.df)
                st.success(f"✅ {new_fio} мұғалімі қосылды!")
                st.rerun()
            else:
                st.error("ФИО толтырыңыз!")

st.markdown("---")
st.caption("© Мұғалімдер рейтингі - Өрлеу бағдарламасы бойынша")
