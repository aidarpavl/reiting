import streamlit as st
import pandas as pd
import requests
from io import BytesIO
import os

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

def load_from_github():
    """GitHub-тан Excel файлын жүктеу"""
    try:
        with st.spinner("📥 GitHub-тан жүктелуде..."):
            response = requests.get(GITHUB_URL, timeout=30)
            response.raise_for_status()
            
            # Жергілікті сақтау
            with open(LOCAL_FILE, 'wb') as f:
                f.write(response.content)
            
            # Excel оқу
            df = pd.read_excel(LOCAL_FILE, engine='openpyxl')
            
            # Баған атауларын стандарттау
            expected_columns = ['№', 'ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік', 
                               'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
            
            if len(df.columns) == len(expected_columns):
                df.columns = expected_columns
            else:
                df = df.iloc[:, :11]
                df.columns = expected_columns
            
            df['№'] = range(1, len(df) + 1)
            st.session_state.df = df
            st.session_state.data_loaded = True
            return True
    except Exception as e:
        st.error(f"❌ Жүктеу қатесі: {str(e)}")
        return False

def save_to_github():
    """Өзгерістерді жергілікті файлға сақтау"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return False
    
    try:
        st.session_state.df.to_excel(LOCAL_FILE, index=False, engine='openpyxl')
        st.success("✅ Файл жергілікті компьютерге сақталды!")
        
        # GitHub-қа жүктеу нұсқауы
        st.info("""
        📌 **GitHub-қа жүктеу үшін:**
        1. [GitHub репозиторийіне](https://github.com/aidarpavl/reiting) өтіңіз
        2. `reiting1.xlsx` файлын тауып, **Update** батырмасын басыңыз
        3. Жаңа файлды таңдап, **Commit changes** батырмасын басыңыз
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
    
    df = st.session_state.df
    
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

def add_teacher():
    """Жаңа мұғалім қосу"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    with st.form("add_teacher_form"):
        st.subheader("➕ Жаңа мұғалім қосу")
        
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("ФИО")
            subject = st.text_input("Предмет")
            coeff = st.number_input("Коэф.", value=1.0, step=0.1)
            avg_score = st.number_input("Ср.балл", value=4.0, step=0.1)
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
    
    if st.button("🗑 Өшіру", type="primary"):
        if st.button("Растау", key="confirm_delete"):
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
    if st.button("💾 GitHub-қа сақтау", use_container_width=True):
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
    
    # Түсті кодтау үшін стиль
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
    
    # Кестені көрсету
    styled_df = st.session_state.df.style.applymap(color_negative, subset=['Жалпы итог', 'Результативность'])
    st.dataframe(styled_df, use_container_width=True, height=400)
    
    # Статистика
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📚 Мұғалімдер саны", len(st.session_state.df))
    with col2:
        avg_total = st.session_state.df['Жалпы итог'].mean()
        st.metric("📊 Орташа жалпы итог", f"{avg_total:.1f}")
    with col3:
        best = st.session_state.df.loc[st.session_state.df['Жалпы итог'].idxmax(), 'ФИО']
        best_score = st.session_state.df['Жалпы итог'].max()
        st.metric("🏆 Ең жақсы нәтиже", f"{best} ({best_score:.1f})")
    with col4:
        worst = st.session_state.df.loc[st.session_state.df['Жалпы итог'].idxmin(), 'ФИО']
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
    st.info("👆 Бастау үшін «GitHub-тан жүктеу» батырмасын басыңыз.")

# Футер
st.markdown("---")
st.caption("© Мұғалімдер рейтингі | Деректер GitHub репозиторийінде сақталады")