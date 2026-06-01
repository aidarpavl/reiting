import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import requests
from urllib.parse import urlparse

# ==================== КОНФИГУРАЦИЯ ====================
st.set_page_config(
    page_title="Мұғалімдер рейтингі",
    page_icon="📊",
    layout="wide"
)

# GitHub-тағы Excel файлының RAW URL-і (ӨЗІҢІЗДІҢ URL-ІҢІЗБЕН АУЫСТЫРЫҢЫЗ)
GITHUB_EXCEL_URL = "https://raw.githubusercontent.com/aidarpavl/reiting/refs/heads/main/reiting.xlsx"

# Демо деректер (кестедегі ақпарат бойынша)
DEMO_DATA = {
    "ФИО": ["АХШАЛОВА", "АЗЫМБАЕВА", "АЙТБАЙ", "АЛЬМЕНОВА", "АЛЬНАЗИРОВА", 
             "АРЫНГАЗИНА", "ИЛЬЯСОВА", "ИТЕМГЕНОВ", "КАБЫЛОВА", "КАЗАНГАПОВ", 
             "КАЙДАРОВА", "КОЖАГЕЛЬДИНОВА", "МОМЫНОВ"],
    "Предмет": ["Математика", "Химия", "Математика", "Русский язык", "Ин яз",
                "Каз яз", "Ин яз", "история", "Каз яз", "Ин яз",
                "География", "Математика", "Информатика"],
    "Коэф": [1.1, 1.0, 1.1, 0.9, 1.1, 0.9, 1.0, 0.8, 0.9, 1.1, 0.6, 1.1, 1.0],
    "Орг.демалыс": [1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4],
    "Вед.рейтинг": [2, 5, 7, 8, 7, 7, 6, 5, 1, 9, 3, 4, 3],
    "Флеш-моб": [3, 8, 9, 5, 3, 4, 7, 10, 8, 11, 3, 10, 8],
    "Внекл.мероп": [5, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Дежурство": [5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5],
    "Конк(гор)": [2, 2, 1, 2, 2, 3, 2, 2, 2, 2, 2, 1, 1],
    "Конк(обл)": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Конк(респ)": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Внекл.мероп2": [1, 2, 1, 1, 2, 1, 3, 2, 1, 1, 1, 2, 1],
    "СМИ": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Обобщ.опыт": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Публикац": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Сдача отчетов": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Посещаемость": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Кл.уголок": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "ПДД": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Питание": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2],
    "Журнал": [3, 2, 3, 1, 4, 5, 2, 1, 2, 1, 1, 1, 2]
}

# ==================== ЕСЕПТЕУ ФУНКЦИЯЛАРЫ ====================
def calculate_extracurricular(row):
    """Внеклассная работа итогын есептеу"""
    return (row.get('Орг.демалыс', 0) + row.get('Вед.рейтинг', 0) + 
            row.get('Флеш-моб', 0) + row.get('Внекл.мероп', 0) + row.get('Дежурство', 0))

def calculate_methodical(row):
    """Методическая деятельность итогын есептеу"""
    return (row.get('Конк(гор)', 0) + row.get('Конк(обл)', 0) + row.get('Конк(респ)', 0) + 
            row.get('Внекл.мероп2', 0) + row.get('СМИ', 0) + row.get('Обобщ.опыт', 0) + row.get('Публикац', 0))

def calculate_admin(row):
    """Администрация итогын есептеу"""
    return (row.get('Сдача отчетов', 0) + row.get('Посещаемость', 0) + row.get('Кл.уголок', 0) + 
            row.get('ПДД', 0) + row.get('Питание', 0) + row.get('Журнал', 0))

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
        df['Пайыз (%)'] = 100.0
    else:
        df['Пайыз (%)'] = ((df['Жалпы итог'] - min_score) / range_score) * 100
    return df

# ==================== GITHUB-ТАН EXCEL ОҚУ ====================
@st.cache_data(ttl=300)
def load_from_github():
    """GitHub-тан Excel файлын оқу"""
    try:
        response = requests.get(GITHUB_EXCEL_URL, timeout=30)
        if response.status_code == 200:
            excel_data = BytesIO(response.content)
            df = pd.read_excel(excel_data, engine='openpyxl')
            return df
        else:
            st.warning(f"GitHub-тан оқу мүмкін болмады (статус: {response.status_code})")
            return None
    except Exception as e:
        st.error(f"Қате: {str(e)}")
        return None

# ==================== ИНТЕРФЕЙС ====================
st.title("📊 Мұғалімдердің кешенді рейтингі")
st.caption("Өрлеу бағдарламасы бойынша - Класс жетекшілерді бағалау жүйесі")

# Бүйірлік панель - басқару
with st.sidebar:
    st.header("⚙️ Басқару панелі")
    
    # GitHub URL-ді өзгерту мүмкіндігі
    github_url = st.text_input("GitHub Excel URL:", value=GITHUB_EXCEL_URL)
    if github_url != GITHUB_EXCEL_URL:
        GITHUB_EXCEL_URL = github_url
        st.cache_data.clear()
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 GitHub-тан оқу", use_container_width=True):
            with st.spinner("Деректер жүктелуде..."):
                df = load_from_github()
                if df is not None and not df.empty:
                    # Баған атауларын стандарттау
                    df.columns = [col.strip() for col in df.columns]
                    st.session_state.df = process_data(df)
                    st.session_state.data_source = "GitHub"
                    st.success(f"✅ {len(df)} мұғалім GitHub-тан жүктелді!")
                    st.rerun()
                else:
                    st.error("GitHub-тан оқу мүмкін болмады. Демо-деректер қолданылуда.")
                    st.session_state.df = pd.DataFrame(DEMO_DATA)
                    st.session_state.df = process_data(st.session_state.df)
                    st.session_state.data_source = "Demo"
    
    with col2:
        if st.button("📋 Демо-деректер", use_container_width=True):
            st.session_state.df = pd.DataFrame(DEMO_DATA)
            st.session_state.df = process_data(st.session_state.df)
            st.session_state.data_source = "Demo"
            st.success("✅ Демо-деректер жүктелді!")
            st.rerun()
    
    st.markdown("---")
    
    if st.button("📈 Пайыздарды есептеу", use_container_width=True):
        if 'df' in st.session_state:
            st.session_state.df = calculate_percentages(st.session_state.df)
            st.success("✅ Пайыздар есептелді!")
            st.rerun()
    
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
    
    st.markdown("---")
    st.caption(f"📌 Деректер көзі: {st.session_state.get('data_source', 'Жүктелмеген')}")

# Негізгі контент
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame(DEMO_DATA)
    st.session_state.df = process_data(st.session_state.df)
    st.session_state.data_source = "Demo"

# Іздеу
search = st.text_input("🔍 Мұғалім іздеу (аты/пәні бойынша)", placeholder="Атын немесе пәнін жазыңыз...")
filtered_df = st.session_state.df.copy()
if search:
    filtered_df = filtered_df[
        filtered_df['ФИО'].str.contains(search, case=False, na=False) | 
        filtered_df['Предмет'].str.contains(search, case=False, na=False)
    ]

st.info(f"👩‍🏫 Барлығы: **{len(filtered_df)}** мұғалім | Деректер көзі: **{st.session_state.get('data_source', 'Demo')}**")

# Негізгі кесте
st.subheader("📋 Мұғалімдер кестесі")

# Көрсетілетін бағандар
display_columns = ['ФИО', 'Предмет', 'Коэф', 'Внекл итог', 'Метод итог', 'Админ итог', 'Жалпы итог']
if 'Пайыз (%)' in filtered_df.columns:
    display_columns.append('Пайыз (%)')

st.dataframe(
    filtered_df[display_columns],
    use_container_width=True,
    height=450,
    column_config={
        "ФИО": st.column_config.TextColumn("ФИО", width="medium"),
        "Предмет": st.column_config.TextColumn("Предмет", width="small"),
        "Коэф": st.column_config.NumberColumn("Коэф", format="%.1f"),
        "Внекл итог": st.column_config.NumberColumn("Внекл итог", format="%.1f"),
        "Метод итог": st.column_config.NumberColumn("Метод итог", format="%.1f"),
        "Админ итог": st.column_config.NumberColumn("Админ итог", format="%.0f"),
        "Жалпы итог": st.column_config.NumberColumn("Жалпы итог", format="%.1f"),
        "Пайыз (%)": st.column_config.NumberColumn("Пайыз (%)", format="%.1f%%"),
    }
)

# Үздік және төмен көрсеткішті мұғалімдер
col1, col2 = st.columns(2)

with col1:
    st.subheader("🏆 Үздік мұғалімдер")
    top3 = filtered_df.nlargest(3, 'Жалпы итог')[['ФИО', 'Предмет', 'Жалпы итог']]
    for i, (_, row) in enumerate(top3.iterrows(), 1):
        medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉"
        with st.container():
            st.markdown(f"""
            <div style='background: linear-gradient(135deg, #f5f7fa, #c3cfe2); 
                        border-radius: 10px; padding: 15px; margin: 5px 0;'>
                <span style='font-size: 24px;'>{medal}</span>
                <span style='font-size: 18px; font-weight: bold; margin-left: 10px;'>{row['ФИО']}</span>
                <span style='float: right; font-size: 20px; color: #667eea;'>{row['Жалпы итог']:.1f} балл</span>
                <div style='font-size: 12px; color: #666; margin-top: 5px;'>{row['Предмет']}</div>
            </div>
            """, unsafe_allow_html=True)

with col2:
    st.subheader("⚠️ Көмек қажет мұғалімдер")
    bottom3 = filtered_df.nsmallest(3, 'Жалпы итог')[['ФИО', 'Предмет', 'Жалпы итог']]
    for i, (_, row) in enumerate(bottom3.iterrows(), 1):
        with st.container():
            st.markdown(f"""
            <div style='background: #fff3e0; border-radius: 10px; padding: 15px; margin: 5px 0; border-left: 5px solid #f39c12;'>
                <span style='font-size: 18px; font-weight: bold;'>{row['ФИО']}</span>
                <span style='float: right; font-size: 18px; color: #e74c3c;'>{row['Жалпы итог']:.1f} балл</span>
                <div style='font-size: 12px; color: #666; margin-top: 5px;'>{row['Предмет']}</div>
            </div>
            """, unsafe_allow_html=True)

# График
st.subheader("📊 Мұғалімдердің жалпы итогтары")

tab1, tab2 = st.tabs(["📊 Бағандық диаграмма", "📈 Сызықтық диаграмма"])

with tab1:
    fig = px.bar(
        filtered_df,
        x='ФИО',
        y='Жалпы итог',
        color='Жалпы итог',
        color_continuous_scale=['#e74c3c', '#f39c12', '#27ae60'],
        title="Мұғалімдердің рейтингі",
        labels={'ФИО': 'Мұғалім', 'Жалпы итог': 'Жалпы итог (балл)'},
        text='Жалпы итог',
        height=500
    )
    fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')
    fig.update_layout(showlegend=False, plot_bgcolor='white')
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    fig_line = px.line(
        filtered_df,
        x='ФИО',
        y='Жалпы итог',
        markers=True,
        title="Мұғалімдер рейтингі (тренд)",
        labels={'ФИО': 'Мұғалім', 'Жалпы итог': 'Жалпы итог (балл)'},
        height=500
    )
    fig_line.update_traces(line_color='#667eea', marker=dict(size=10, color='#764ba2'))
    fig_line.update_layout(plot_bgcolor='white')
    st.plotly_chart(fig_line, use_container_width=True)

# Статистика
st.subheader("📈 Статистикалық мәліметтер")
stat_cols = st.columns(4)
with stat_cols[0]:
    st.metric("📊 Орташа итог", f"{filtered_df['Жалпы итог'].mean():.1f}")
with stat_cols[1]:
    st.metric("🏆 Ең жоғары көрсеткіш", f"{filtered_df['Жалпы итог'].max():.1f}", 
              delta=f"+{filtered_df['Жалпы итог'].max() - filtered_df['Жалпы итог'].mean():.1f}")
with stat_cols[2]:
    st.metric("⚠️ Ең төмен көрсеткіш", f"{filtered_df['Жалпы итог'].min():.1f}",
              delta=f"{filtered_df['Жалпы итог'].min() - filtered_df['Жалпы итог'].mean():.1f}")
with stat_cols[3]:
    positive = len(filtered_df[filtered_df['Жалпы итог'] > 0])
    st.metric("✅ Оң нәтижелілер", f"{positive}/{len(filtered_df)}")

# Көмек ұсыныстары
st.subheader("💡 Төмен балл алған мұғалімдерге көмек ұсыныстары")
with st.expander("📋 Көмек шараларын көру", expanded=True):
    help_cols = st.columns(2)
    help_items_left = [
        "📚 Әдістемелік бірлестік отырыстарына жүйелі түрде қатысу",
        "👥 Тәжірибелі әріптестермен жұптасып сабақ өткізу",
        "📊 Жеке даму жоспарын құру (3-6 айға)",
    ]
    help_items_right = [
        "🏫 Ішкі оқыту семинарларына қатысу",
        "📝 Портфолио жүргізу және жетістіктерді тіркеу",
        "🎯 Әдістемелік көмек көрсету үшін тәлімгер тағайындау",
    ]
    
    with help_cols[0]:
        for item in help_items_left:
            st.write(f"- {item}")
    with help_cols[1]:
        for item in help_items_right:
            st.write(f"- {item}")

# Мұғалімдерді басқару
st.subheader("✏️ Мұғалімдерді басқару")

tab_edit, tab_add, tab_delete = st.tabs(["📝 Өңдеу", "➕ Қосу", "🗑 Өшіру"])

with tab_edit:
    st.info("Кестедегі мәндерді тікелей өңдеуге болады. Жоғарыдағы кестедегі ұяшықтарды екі рет басыңыз.")
    
with tab_add:
    with st.form("add_teacher_form"):
        st.subheader("Жаңа мұғалім мәліметтері")
        col_a, col_b = st.columns(2)
        with col_a:
            new_fio = st.text_input("ФИО *", placeholder="Тегі Аты")
            new_subject = st.selectbox("Предмет *", ["Математика", "Қазақ тілі", "Орыс тілі", "Ағылшын тілі", "Физика", "Химия", "Биология", "Тарих", "География", "Информатика"])
            new_coef = st.number_input("Коэффицент", min_value=0.5, max_value=2.0, value=1.0, step=0.1)
        
        with col_b:
            new_class = st.text_input("Сынып", placeholder="7А, 8Б...")
            new_exp = st.number_input("Жұмыс өтілі", min_value=0, max_value=40, value=5)
        
        st.subheader("Внеклассная работа (1-5 балл)")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            new_org = st.number_input("Орг.демалыс", min_value=0, max_value=5, value=0)
        with col2:
            new_rating = st.number_input("Вед.рейтинг", min_value=0, max_value=5, value=0)
        with col3:
            new_flash = st.number_input("Флеш-моб", min_value=0, max_value=5, value=0)
        with col4:
            new_vnekl = st.number_input("Внекл.мероп", min_value=0, max_value=5, value=0)
        with col5:
            new_desh = st.number_input("Дежурство", min_value=0, max_value=5, value=0)
        
        if st.form_submit_button("➕ Мұғалімді қосу", use_container_width=True):
            if new_fio:
                new_row = pd.DataFrame([{
                    'ФИО': new_fio, 'Предмет': new_subject, 'Коэф': new_coef,
                    'Орг.демалыс': new_org, 'Вед.рейтинг': new_rating, 'Флеш-моб': new_flash,
                    'Внекл.мероп': new_vnekl, 'Дежурство': new_desh,
                    'Конк(гор)': 0, 'Конк(обл)': 0, 'Конк(респ)': 0, 'Внекл.мероп2': 0,
                    'СМИ': 0, 'Обобщ.опыт': 0, 'Публикац': 0, 'Сдача отчетов': 0,
                    'Посещаемость': 0, 'Кл.уголок': 0, 'ПДД': 0, 'Питание': 0, 'Журнал': 0
                }])
                st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
                st.session_state.df = process_data(st.session_state.df)
                st.success(f"✅ {new_fio} мұғалімі қосылды!")
                st.rerun()
            else:
                st.error("ФИО толтырыңыз!")

with tab_delete:
    teacher_to_delete = st.selectbox("Өшіретін мұғалімді таңдаңыз:", [''] + list(st.session_state.df['ФИО'].unique()))
    if teacher_to_delete and st.button("🗑 Мұғалімді өшіру", type="secondary"):
        st.session_state.df = st.session_state.df[st.session_state.df['ФИО'] != teacher_to_delete]
        st.session_state.df = process_data(st.session_state.df)
        st.success(f"✅ {teacher_to_delete} мұғалімі өшірілді!")
        st.rerun()

# Footer
st.markdown("---")
st.caption(f"© Мұғалімдер рейтингі - Өрлеу бағдарламасы бойынша | Соңғы жаңарту: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")