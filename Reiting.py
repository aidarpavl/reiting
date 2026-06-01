import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import requests

# ==================== КОНФИГУРАЦИЯ ====================
st.set_page_config(
    page_title="Мұғалімдер рейтингі",
    page_icon="📊",
    layout="wide"
)

# GitHub-тағы Excel файлының RAW URL-і
GITHUB_EXCEL_URL = "https://raw.githubusercontent.com/aidarpavl/reiting/refs/heads/main/reiting.xlsx"

# ==================== ДЕРЕКТЕРДІ ОҚУ ФУНКЦИЯСЫ ====================
@st.cache_data(ttl=300)
def load_data():
    """GitHub-тан Excel файлын оқып, мұғалімдер деректерін қайтару"""
    try:
        response = requests.get(GITHUB_EXCEL_URL, timeout=30)
        if response.status_code == 200:
            excel_data = BytesIO(response.content)
            
            # Барлық парақтарды оқу
            xl = pd.ExcelFile(excel_data)
            st.sidebar.info(f"📑 Парақтар: {', '.join(xl.sheet_names)}")
            
            # "Kl ruk" парағын оқу (мұғалімдер деректері)
            if "Kl ruk" in xl.sheet_names:
                df_raw = pd.read_excel(excel_data, sheet_name="Kl ruk", header=None)
                return parse_kl_ruk_sheet(df_raw)
            else:
                st.error("'Kl ruk' парағы табылмады")
                return None
        else:
            st.error(f"GitHub-тан оқу мүмкін болмады (статус: {response.status_code})")
            return None
    except Exception as e:
        st.error(f"Қате: {str(e)}")
        return None

def parse_kl_ruk_sheet(df_raw):
    """Kl ruk парағын парсингтеу - мұғалімдер деректерін алу"""
    teachers = []
    
    # Деректер 4-ші жолдан басталады (индекс 3)
    # Бағандар: A=№, B=ФИО, C=Предмет, D=Коэф, E=Орг.демалыс, F=Вед.рейтинг, 
    # G=Флеш-моб, H=Внекл.мероп, I=Дежурство, J=Итог (формула), K-?...
    
    for idx in range(3, len(df_raw)):
        row = df_raw.iloc[idx]
        
        # ФИО бар жолдарды анықтау
        fio = row[1] if pd.notna(row[1]) else None
        if not fio or str(fio).strip() == '':
            continue
        
        # Негізгі деректерді алу
        subject = row[2] if pd.notna(row[2]) else "Пән"
        coef = float(row[3]) if pd.notna(row[3]) else 1.0
        
        # Внеклассная работа (бағандар E, F, G, H, I)
        org = float(row[4]) if pd.notna(row[4]) else 0
        rating = float(row[5]) if pd.notna(row[5]) else 0
        flash = float(row[6]) if pd.notna(row[6]) else 0
        vnekl = float(row[7]) if pd.notna(row[7]) else 0
        desh = float(row[8]) if pd.notna(row[8]) else 0
        
        # Методическая деятельность (K, L, M, N, O, P, Q бағандары)
        gor = float(row[10]) if len(row) > 10 and pd.notna(row[10]) else 0
        obl = float(row[11]) if len(row) > 11 and pd.notna(row[11]) else 0
        resp = float(row[12]) if len(row) > 12 and pd.notna(row[12]) else 0
        vneklMero = float(row[13]) if len(row) > 13 and pd.notna(row[13]) else 0
        smi = float(row[14]) if len(row) > 14 and pd.notna(row[14]) else 0
        obob = float(row[15]) if len(row) > 15 and pd.notna(row[15]) else 0
        publ = float(row[16]) if len(row) > 16 and pd.notna(row[16]) else 0
        
        # Рейтинг у администрации (S, T, U, V, W, X бағандары)
        otchet = float(row[18]) if len(row) > 18 and pd.notna(row[18]) else 0
        posesh = float(row[19]) if len(row) > 19 and pd.notna(row[19]) else 0
        ugalok = float(row[20]) if len(row) > 20 and pd.notna(row[20]) else 0
        pdd = float(row[21]) if len(row) > 21 and pd.notna(row[21]) else 0
        pitanie = float(row[22]) if len(row) > 22 and pd.notna(row[22]) else 0
        zhurnal = float(row[23]) if len(row) > 23 and pd.notna(row[23]) else 0
        
        teachers.append({
            "ФИО": str(fio).strip(),
            "Предмет": str(subject).strip(),
            "Коэф": coef,
            "Орг.демалыс": org,
            "Вед.рейтинг": rating,
            "Флеш-моб": flash,
            "Внекл.мероп": vnekl,
            "Дежурство": desh,
            "Конк(гор)": gor,
            "Конк(обл)": obl,
            "Конк(респ)": resp,
            "Внекл.мероп2": vneklMero,
            "СМИ": smi,
            "Обобщ.опыт": obob,
            "Публикац": publ,
            "Сдача отчетов": otchet,
            "Посещаемость": posesh,
            "Кл.уголок": ugalok,
            "ПДД": pdd,
            "Питание": pitanie,
            "Журнал": zhurnal,
        })
    
    if teachers:
        df = pd.DataFrame(teachers)
        return df
    return None

# ==================== ЕСЕПТЕУ ФУНКЦИЯЛАРЫ ====================
def calculate_extracurricular(row):
    return (row.get('Орг.демалыс', 0) + row.get('Вед.рейтинг', 0) + 
            row.get('Флеш-моб', 0) + row.get('Внекл.мероп', 0) + row.get('Дежурство', 0))

def calculate_methodical(row):
    return (row.get('Конк(гор)', 0) + row.get('Конк(обл)', 0) + row.get('Конк(респ)', 0) + 
            row.get('Внекл.мероп2', 0) + row.get('СМИ', 0) + row.get('Обобщ.опыт', 0) + row.get('Публикац', 0))

def calculate_admin(row):
    return (row.get('Сдача отчетов', 0) + row.get('Посещаемость', 0) + row.get('Кл.уголок', 0) + 
            row.get('ПДД', 0) + row.get('Питание', 0) + row.get('Журнал', 0))

def process_data(df):
    df = df.copy()
    df['Внекл итог'] = df.apply(calculate_extracurricular, axis=1)
    df['Метод итог'] = df.apply(calculate_methodical, axis=1)
    df['Админ итог'] = df.apply(calculate_admin, axis=1)
    df['Жалпы итог'] = df['Внекл итог'] + df['Метод итог'] + df['Админ итог']
    return df

def calculate_percentages(df):
    min_score = df['Жалпы итог'].min()
    max_score = df['Жалпы итог'].max()
    range_score = max_score - min_score
    if range_score == 0:
        df['Пайыз (%)'] = 100.0
    else:
        df['Пайыз (%)'] = ((df['Жалпы итог'] - min_score) / range_score) * 100
    return df

# ==================== ИНТЕРФЕЙС ====================
st.title("📊 Мұғалімдердің кешенді рейтингі")
st.caption("Өрлеу бағдарламасы бойынша - Класс жетекшілерді бағалау жүйесі")

# Бүйірлік панель
with st.sidebar:
    st.header("⚙️ Басқару")
    
    # URL өзгерту
    github_url = st.text_input("GitHub Excel URL:", value=GITHUB_EXCEL_URL)
    if github_url != GITHUB_EXCEL_URL:
        GITHUB_EXCEL_URL = github_url
        st.cache_data.clear()
    
    if st.button("🔄 Деректерді жүктеу", use_container_width=True):
        with st.spinner("Деректер жүктелуде..."):
            df = load_data()
            if df is not None and not df.empty:
                st.session_state.df = process_data(df)
                st.session_state.data_source = "GitHub"
                st.success(f"✅ {len(df)} мұғалім жүктелді!")
                st.rerun()
    
    # Демо-деректер
    if st.button("📋 Демо-деректерді пайдалану", use_container_width=True):
        st.session_state.use_demo = True
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
                label="📥 Жүктеп алу",
                data=output.getvalue(),
                file_name=f"teacher_rating_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# Демо-деректер (кестедегі мәліметтер бойынша)
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

# Бастапқы деректер
if 'df' not in st.session_state:
    st.session_state.df = pd.DataFrame(DEMO_DATA)
    st.session_state.df = process_data(st.session_state.df)
    st.session_state.data_source = "Demo"
    st.session_state.use_demo = True

# Бүйірлік панельде деректер көзін көрсету
st.sidebar.info(f"📌 Деректер көзі: {st.session_state.get('data_source', 'Белгісіз')}")

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
display_columns = ['ФИО', 'Предмет', 'Коэф', 'Внекл итог', 'Метод итог', 'Админ итог', 'Жалпы итог']
if 'Пайыз (%)' in filtered_df.columns:
    display_columns.append('Пайыз (%)')

st.dataframe(
    filtered_df[display_columns],
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

# Footer
st.markdown("---")
st.caption(f"© Мұғалімдер рейтингі - Өрлеу бағдарламасы бойынша | Соңғы жаңарту: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")