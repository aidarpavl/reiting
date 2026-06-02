import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import os
import traceback

# GitHub файл сілтемесі
GITHUB_URL = "https://raw.githubusercontent.com/aidarpavl/reiting/refs/heads/main/reiting1.xlsx"
LOCAL_FILE = "reiting1.xlsx"

# Бетті конфигурациялау
st.set_page_config(
    page_title="Мұғалімдер рейтингі",
    page_icon="📊",
    layout="wide"
)

# Сессия күйін инициализациялау
if 'df' not in st.session_state:
    st.session_state.df = None
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'error_message' not in st.session_state:
    st.session_state.error_message = None

def load_from_github():
    """GitHub-тан Excel файлын жүктеу"""
    try:
        with st.spinner("📥 GitHub-тан жүктелуде..."):
            st.session_state.error_message = None
            
            # Файлды жүктеу
            response = requests.get(GITHUB_URL, timeout=30)
            
            if response.status_code == 404:
                st.session_state.error_message = "Файл табылмады! GitHub сілтемесін тексеріңіз."
                st.error(st.session_state.error_message)
                return False
                
            response.raise_for_status()
            
            # Жергілікті сақтау
            with open(LOCAL_FILE, 'wb') as f:
                f.write(response.content)
            
            # Excel оқу
            df = pd.read_excel(LOCAL_FILE, engine='openpyxl')
            
            # Баған атауларын стандарттау
            expected_columns = ['№', 'ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік', 
                               'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
            
            # Егер баған саны сәйкес келсе, атауларын өзгерту
            if len(df.columns) == len(expected_columns):
                df.columns = expected_columns
            elif len(df.columns) >= 11:
                df = df.iloc[:, :11]
                df.columns = expected_columns
            else:
                st.warning(f"Баған саны сәйкес емес. Күтілген: 11, Алынған: {len(df.columns)}")
            
            # Нөмірлерді қайта есептеу
            df['№'] = range(1, len(df) + 1)
            
            # Сандық бағандарды конвертациялау
            numeric_cols = ['Коэф', 'Ср.балл', 'Үштік', 'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            st.session_state.df = df
            st.session_state.data_loaded = True
            st.success(f"✅ Жүктелді! {len(df)} мұғалім")
            return True
            
    except requests.exceptions.ConnectionError:
        st.session_state.error_message = "Интернет қосылымын тексеріңіз!"
        st.error(st.session_state.error_message)
        return False
    except Exception as e:
        st.session_state.error_message = str(e)
        st.error(f"❌ Жүктеу қатесі: {str(e)}")
        return False

def save_to_github():
    """Өзгерістерді жергілікті файлға сақтау"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return False
    
    try:
        # Көрсету үшін көшірме жасау
        save_df = st.session_state.df.copy()
        save_df.to_excel(LOCAL_FILE, index=False, engine='openpyxl')
        st.success("✅ Файл жергілікті компьютерге сақталды!")
        
        # GitHub-қа жүктеу нұсқауы
        with st.expander("📌 GitHub-қа қалай жүктеу керек?"):
            st.markdown("""
            1. [GitHub репозиторийіне](https://github.com/aidarpavl/reiting) өтіңіз
            2. `reiting1.xlsx` файлын тауып, **Update** батырмасын басыңыз
            3. Жаңа файлды таңдаңыз (жергілікті `reiting1.xlsx`)
            4. **Commit changes** батырмасын басыңыз
            """)
        return True
    except Exception as e:
        st.error(f"❌ Сақтау қатесі: {str(e)}")
        return False

def calculate_rating():
    """Рейтингті есептеу"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    try:
        df = st.session_state.df.copy()
        
        # Результативность = (Ср.балл × Коэф) – (Үштік × 0.9)
        df['Результативность'] = (df['Ср.балл'] * df['Коэф']) - (df['Үштік'] * 0.9)
        
        # Жалпы итог = Результативность + Внеклас. + Метод. + Админ.рейтинг
        df['Жалпы итог'] = (df['Результативность'] + 
                             df['Внеклас.'] + 
                             df['Метод.'] + 
                             df['Админ.рейтинг'])
        
        # Сандарды дөңгелектеу
        df['Результативность'] = df['Результативность'].round(1)
        df['Жалпы итог'] = df['Жалпы итог'].round(1)
        
        st.session_state.df = df
        st.success("✅ Рейтинг сәтті есептелді!")
        st.rerun()
    except Exception as e:
        st.error(f"❌ Есептеу қатесі: {str(e)}")

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
    teacher_options = {f"{row['№']}. {row['ФИО']} - {row['Предмет']}": idx 
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

st.title("📊 Мұғалімдер рейтингі – GitHub интеграциясы")
st.markdown("---")

# Батырмалар панелі
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    if st.button("📥 GitHub-тан жүктеу", use_container_width=True):
        load_from_github()

with col2:
    if st.button("💾 Жергілікті сақтау", use_container_width=True):
        save_to_github()

with col3:
    if st.button("➕ Жаңа мұғалім қосу", use_container_width=True):
        add_teacher()

with col4:
    if st.button("🔄 Рейтингті есептеу", use_container_width=True):
        calculate_rating()

with col5:
    if st.button("🗑 Мұғалім өшіру", use_container_width=True):
        delete_teacher()

st.markdown("---")

# Кестені көрсету
if st.session_state.df is not None:
    st.subheader("📋 Мұғалімдер тізімі")
    
    # Түсті кодтау функциясы (жаңа pandas нұсқасына бейімделген)
    def color_negative(val):
        try:
            val = float(val)
            if val < 0:
                return 'color: red'
            elif val > 5:
                return 'color: green'
            return ''
        except:
            return ''
    
    # Кестені стильдеу (applymap орнына map қолданылды)
    try:
        styled_df = st.session_state.df.style.map(color_negative, subset=['Жалпы итог', 'Результативность'])
        st.dataframe(styled_df, use_container_width=True, height=400)
    except:
        # Егер стильдеу қате берсе, жай кестені көрсету
        st.dataframe(st.session_state.df, use_container_width=True, height=400)
    
    # Статистика
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📚 Мұғалімдер саны", len(st.session_state.df))
    with col2:
        avg_total = st.session_state.df['Жалпы итог'].mean()
        st.metric("📊 Орташа жалпы итог", f"{avg_total:.1f}")
    with col3:
        if len(st.session_state.df) > 0:
            best_idx = st.session_state.df['Жалпы итог'].idxmax()
            best = st.session_state.df.loc[best_idx, 'ФИО']
            best_score = st.session_state.df['Жалпы итог'].max()
            st.metric("🏆 Ең жақсы нәтиже", f"{best} ({best_score:.1f})")
    with col4:
        if len(st.session_state.df) > 0:
            worst_idx = st.session_state.df['Жалпы итог'].idxmin()
            worst = st.session_state.df.loc[worst_idx, 'ФИО']
            worst_score = st.session_state.df['Жалпы итог'].min()
            st.metric("⚠️ Ең төмен нәтиже", f"{worst} ({worst_score:.1f})")
    
    # Экспорт батырмасы
    csv = st.session_state.df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 CSV форматында жүктеу",
        data=csv,
        file_name="reiting_export.csv",
        mime="text/csv",
    )

else:
    st.info("👆 Бастау үшін **«GitHub-тан жүктеу»** батырмасын басыңыз.")
    
    # Үлгі кесте жасау батырмасы
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

# Футер
st.markdown("---")
st.caption("© Мұғалімдер рейтингі | Деректер GitHub репозиторийінде сақталады")