import streamlit as st
import pandas as pd
import requests
import base64
from io import BytesIO

# ==================== GITHUB ПАРАМЕТРЛЕРІ ====================
GITHUB_RAW_URL = "https://raw.githubusercontent.com/aidarpavl/reiting/refs/heads/main/reiting1.xlsx"
GITHUB_API_URL = "https://api.github.com/repos/aidarpavl/reiting/contents/reiting1.xlsx"

# ==================== БЕТ КОНФИГУРАЦИЯСЫ ====================
st.set_page_config(
    page_title="Мұғалімдер рейтингі",
    page_icon="📊",
    layout="wide"
)

# ==================== СЕССИЯ КҮЙІН ИНИЦИАЛИЗАЦИЯЛАУ ====================
if 'df' not in st.session_state:
    st.session_state.df = None
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'github_token' not in st.session_state:
    st.session_state.github_token = None

# ==================== ФУНКЦИЯЛАР ====================
def show_token_form():
    """Сол жақ панельде GitHub токенін енгізу формасы"""
    with st.sidebar:
        st.markdown("### 🔐 GitHub авторизациясы")
        st.markdown("GitHub-қа тікелей сақтау үшін токен қажет")
        token = st.text_input("GitHub Personal Access Token:", type="password", key="token_input")
        
        if token:
            st.session_state.github_token = token
            st.success("✅ Токен сақталды!")
        
        st.markdown("""
        **Токен алу үшін:**
        1. [GitHub Settings → Tokens](https://github.com/settings/tokens)
        2. **Generate new token (classic)** басыңыз
        3. `repo` рұқсатын қосыңыз
        4. Токенді көшіріп, жоғарыға енгізіңіз
        """)

def load_from_github():
    """GitHub-тан Excel файлын жүктеу"""
    try:
        with st.spinner("📥 GitHub-тан жүктелуде..."):
            response = requests.get(GITHUB_RAW_URL, timeout=30)
            response.raise_for_status()
            
            # Excel оқу
            df = pd.read_excel(BytesIO(response.content), engine='openpyxl')
            
            # Баған атауларын стандарттау
            if len(df.columns) >= 11:
                df = df.iloc[:, :11]
                df.columns = ['№', 'ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік', 
                             'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
            else:
                st.error(f"Күтілетін баған саны 11, алынған: {len(df.columns)}")
                return False
            
            # Сандық бағандарды конвертациялау
            numeric_cols = ['Коэф', 'Ср.балл', 'Үштік', 'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # Нөмірлерді қайта есептеу
            df['№'] = range(1, len(df) + 1)
            
            # Рейтингті қайта есептеу
            df['Результативность'] = (df['Ср.балл'] * df['Коэф']) - (df['Үштік'] * 0.9)
            df['Жалпы итог'] = df['Результативность'] + df['Внеклас.'] + df['Метод.'] + df['Админ.рейтинг']
            df['Результативность'] = df['Результативность'].round(1)
            df['Жалпы итог'] = df['Жалпы итог'].round(1)
            
            st.session_state.df = df
            st.session_state.data_loaded = True
            st.success(f"✅ {len(df)} мұғалім жүктелді!")
            return True
            
    except requests.exceptions.ConnectionError:
        st.error("❌ Интернет қосылымын тексеріңіз!")
        return False
    except Exception as e:
        st.error(f"❌ Жүктеу қатесі: {str(e)}")
        return False

def save_to_github_direct():
    """GitHub API арқылы тікелей сақтау"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    token = st.session_state.github_token
    if not token:
        st.error("🔐 GitHub токені енгізілмеген! Сол жақ панельге токеніңізді енгізіңіз.")
        return
    
    try:
        with st.spinner("💾 GitHub-қа сақталуда..."):
            # DataFrame-ді байттарға түрлендіру
            output = BytesIO()
            save_df = st.session_state.df.copy()
            save_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            content_bytes = output.getvalue()
            content_b64 = base64.b64encode(content_bytes).decode('utf-8')
            
            # Файлдың ағымдағы SHA мәнін алу
            headers = {
                "Authorization": f"token {token}",
                "Accept": "application/vnd.github.v3+json"
            }
            get_response = requests.get(GITHUB_API_URL, headers=headers)
            sha = get_response.json().get('sha') if get_response.status_code == 200 else None
            
            # GitHub-қа жүктеу
            data = {
                "message": "Update reiting1.xlsx",
                "content": content_b64,
                "branch": "main"
            }
            if sha:
                data["sha"] = sha
            
            put_response = requests.put(GITHUB_API_URL, headers=headers, json=data)
            
            if put_response.status_code in [200, 201]:
                st.success("✅ Файл GitHub-қа сәтті сақталды!")
            else:
                error_msg = put_response.json().get('message', 'Белгісіз қате') if put_response.text else 'Белгісіз қате'
                st.error(f"❌ GitHub API қатесі: {error_msg}")
                
    except Exception as e:
        st.error(f"❌ Сақтау қатесі: {str(e)}")

def save_to_local():
    """Жергілікті файлға сақтау (жүктеу)"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    try:
        df = st.session_state.df.copy()
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        st.download_button(
            label="📥 Excel файлын жүктеу",
            data=output,
            file_name="reiting1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel"
        )
    except Exception as e:
        st.error(f"❌ Қате: {str(e)}")

def calculate_rating():
    """Рейтингті есептеу"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    df = st.session_state.df
    df['Результативность'] = (df['Ср.балл'] * df['Коэф']) - (df['Үштік'] * 0.9)
    df['Жалпы итог'] = df['Результативность'] + df['Внеклас.'] + df['Метод.'] + df['Админ.рейтинг']
    df['Результативность'] = df['Результативность'].round(1)
    df['Жалпы итог'] = df['Жалпы итог'].round(1)
    st.session_state.df = df
    st.success("✅ Рейтинг сәтті есептелді!")
    st.rerun()

def add_teacher():
    """Жаңа мұғалім қосу"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    with st.form("add_teacher_form"):
        st.subheader("➕ Жаңа мұғалім қосу")
        
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("ФИО *")
            subject = st.text_input("Предмет *")
            coeff = st.number_input("Коэф.", value=1.0, step=0.1, format="%.1f")
            avg_score = st.number_input("Ср.балл", value=4.0, step=0.1, format="%.1f")
            triplets = st.number_input("Үштік саны", value=0, step=1)
        with col2:
            extra = st.number_input("Внеклас. работа", value=0.0, step=0.5)
            method = st.number_input("Метод. деятельность", value=0.0, step=0.5)
            admin = st.number_input("Админ. рейтинг", value=0.0, step=0.5)
        
        submitted = st.form_submit_button("Қосу")
        
        if submitted:
            if not name or not subject:
                st.error("ФИО және Предмет толтырылуы керек!")
                return
            
            new_row = pd.DataFrame([{
                '№': len(st.session_state.df) + 1,
                'ФИО': name,
                'Предмет': subject,
                'Коэф': coeff,
                'Ср.балл': avg_score,
                'Үштік': triplets,
                'Внеклас.': extra,
                'Метод.': method,
                'Админ.рейтинг': admin,
                'Результативность': 0,
                'Жалпы итог': 0
            }])
            
            st.session_state.df = pd.concat([st.session_state.df, new_row], ignore_index=True)
            st.session_state.df['№'] = range(1, len(st.session_state.df) + 1)
            calculate_rating()
            st.success(f"✅ {name} қосылды!")
            st.rerun()

def delete_teacher():
    """Мұғалімді өшіру"""
    if st.session_state.df is None or len(st.session_state.df) == 0:
        st.warning("Өшіретін мұғалім жоқ!")
        return
    
    df = st.session_state.df
    teacher_options = {f"{row['№']}. {row['ФИО']} - {row['Предмет']} (Итог: {row['Жалпы итог']})": idx 
                       for idx, row in df.iterrows()}
    
    selected = st.selectbox("Өшіретін мұғалімді таңдаңыз:", list(teacher_options.keys()))
    
    if st.button("🗑 Өшіру", type="secondary"):
        idx = teacher_options[selected]
        teacher_name = df.loc[idx, 'ФИО']
        st.session_state.df = df.drop(idx).reset_index(drop=True)
        st.session_state.df['№'] = range(1, len(st.session_state.df) + 1)
        st.success(f"🗑 {teacher_name} өшірілді!")
        st.rerun()

# ==================== НЕГІЗГІ ИНТЕРФЕЙС ====================

# Сол жақ панель
show_token_form()

# Негізгі бет
st.title("📊 Мұғалімдер рейтингі – GitHub интеграциясы")
st.markdown("---")

# Батырмалар
col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    if st.button("📥 GitHub-тан жүктеу", use_container_width=True):
        load_from_github()

with col2:
    if st.button("💾 GitHub-қа сақтау", use_container_width=True):
        save_to_github_direct()

with col3:
    if st.button("📁 Жергілікті сақтау", use_container_width=True):
        save_to_local()

with col4:
    if st.button("➕ Жаңа мұғалім қосу", use_container_width=True):
        add_teacher()

with col5:
    if st.button("🔄 Рейтингті есептеу", use_container_width=True):
        calculate_rating()

with col6:
    if st.button("🗑 Мұғалім өшіру", use_container_width=True):
        delete_teacher()

st.markdown("---")

# Кестені көрсету
if st.session_state.df is not None and len(st.session_state.df) > 0:
    st.subheader("📋 Мұғалімдер тізімі")
    
    # Кесте
    st.dataframe(st.session_state.df, use_container_width=True, height=500)
    
    # Статистика (NaN мәндерін өңдеу)
    st.subheader("📊 Статистика")
    col1, col2, col3, col4 = st.columns(4)
    
    valid_df = st.session_state.df.dropna(subset=['Жалпы итог'])
    
    with col1:
        st.metric("📚 Мұғалімдер саны", len(valid_df))
    with col2:
        avg_total = valid_df['Жалпы итог'].mean()
        st.metric("📊 Орташа жалпы итог", f"{avg_total:.1f}")
    with col3:
        if len(valid_df) > 0:
            best_idx = valid_df['Жалпы итог'].idxmax()
            best = valid_df.loc[best_idx, 'ФИО']
            best_score = valid_df['Жалпы итог'].max()
            st.metric("🏆 Ең жақсы нәтиже", f"{best} ({best_score:.1f})")
    with col4:
        if len(valid_df) > 0:
            worst_idx = valid_df['Жалпы итог'].idxmin()
            worst = valid_df.loc[worst_idx, 'ФИО']
            worst_score = valid_df['Жалпы итог'].min()
            st.metric("⚠️ Ең төмен нәтиже", f"{worst} ({worst_score:.1f})")
    
    # CSV экспорт
    csv = valid_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 CSV форматында жүктеу",
        data=csv,
        file_name="reiting_export.csv",
        mime="text/csv",
    )

elif st.session_state.df is not None and len(st.session_state.df) == 0:
    st.warning("Деректер жоқ. Жаңа мұғалім қосыңыз!")
else:
    st.info("👆 Бастау үшін **«GitHub-тан жүктеу»** батырмасын басыңыз.")
    
    if st.button("📋 Үлгі кесте жасау"):
        sample_data = {
            '№': [1, 2, 3],
            'ФИО': ['АХШАЛОВА', 'АЗЫМБАЕВА', 'АЙТБАЙ'],
            'Предмет': ['Математика', 'Химия', 'Математика'],
            'Коэф': [1.1, 1.0, 1.1],
            'Ср.балл': [4.3, 4.5, 4.7],
            'Үштік': [11, 8, 9],
            'Результативность': [-7.8, -4.5, -5.4],
            'Внеклас.': [0, 0, 0],
            'Метод.': [0, 0, 0],
            'Админ.рейтинг': [-6, -6, -6],
            'Жалпы итог': [-13.8, -10.5, -11.4]
        }
        st.session_state.df = pd.DataFrame(sample_data)
        st.success("✅ Үлгі кесте жасалды!")
        st.rerun()

st.markdown("---")
st.caption("© Мұғалімдер рейтингі | GitHub-қа сақтау үшін сол жақ панельге токен енгізіңіз")