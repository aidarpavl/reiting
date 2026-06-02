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
if 'raw_data' not in st.session_state:
    st.session_state.raw_data = None
if 'github_token' not in st.session_state:
    st.session_state.github_token = None

# ==================== ФУНКЦИЯЛАР ====================
def show_token_form():
    """Сол жақ панельде GitHub токенін енгізу формасы"""
    with st.sidebar:
        st.markdown("### 🔐 GitHub авторизациясы")
        token = st.text_input("GitHub Personal Access Token:", type="password", key="github_token_input")
        if token:
            st.session_state.github_token = token
            st.success("✅ Токен сақталды!")
        
        st.markdown("""
        **Токен алу үшін:**
        1. [GitHub Settings → Tokens](https://github.com/settings/tokens)
        2. **Generate new token (classic)** басыңыз
        3. `repo` рұқсатын қосыңыз
        4. Токенді көшіріңіз
        """)

def load_from_github():
    """GitHub-тан Excel файлын жүктеу - ТОЛЫҚ НҰСҚА"""
    try:
        with st.spinner("📥 GitHub-тан жүктелуде..."):
            # 1. Файлды жүктеу
            response = requests.get(GITHUB_RAW_URL, timeout=30)
            response.raise_for_status()
            
            # 2. Excel оқу (барлық деректерді сақтау)
            excel_data = BytesIO(response.content)
            df_raw = pd.read_excel(excel_data, engine='openpyxl', header=None)
            
            st.write("### 📄 Жүктелген файлдың алғашқы көрінісі:")
            st.dataframe(df_raw.head(20), use_container_width=True)
            
            # 3. Баған атауларын анықтау (бірінші жол - баған атаулары)
            if len(df_raw) > 0:
                # Бірінші жолды баған атауы ретінде алу
                column_names = df_raw.iloc[0].fillna('').astype(str).tolist()
                # Деректер жолдары
                data_rows = df_raw.iloc[1:].reset_index(drop=True)
                
                # Баған атауларын тазалау
                clean_names = []
                for name in column_names:
                    name = str(name).strip()
                    if 'ФИО' in name or 'фио' in name:
                        clean_names.append('ФИО')
                    elif 'Предмет' in name or 'предмет' in name:
                        clean_names.append('Предмет')
                    elif 'Коэф' in name or 'коэф' in name:
                        clean_names.append('Коэф')
                    elif 'Ср.балл' in name or 'ср.балл' in name or 'Средний' in name:
                        clean_names.append('Ср.балл')
                    elif 'Үштік' in name or 'үштік' in name or 'троек' in name:
                        clean_names.append('Үштік')
                    elif 'Результативность' in name or 'результативность' in name:
                        clean_names.append('Результативность')
                    elif 'Внеклас' in name or 'внеклас' in name:
                        clean_names.append('Внеклас.')
                    elif 'Метод' in name or 'метод' in name:
                        clean_names.append('Метод.')
                    elif 'Админ' in name or 'админ' in name or 'рейтинг' in name:
                        clean_names.append('Админ.рейтинг')
                    elif 'Жалпы' in name or 'жалпы' in name or 'Итог' in name:
                        clean_names.append('Жалпы итог')
                    else:
                        clean_names.append(name)
                
                # Бағандардың дұрыстығын тексеру
                required_cols = ['ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік']
                missing_cols = [col for col in required_cols if col not in clean_names]
                
                if missing_cols:
                    st.warning(f"Келесі бағандар табылмады: {missing_cols}")
                    st.info("Баған атауларын қолмен тексеріңіз.")
                    # Үлгі ретінде бағандарды орнату
                    if len(clean_names) >= 11:
                        clean_names = ['№', 'ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік', 
                                      'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
                    else:
                        clean_names = [f'Баған_{i}' for i in range(len(clean_names))]
                
                # DataFrame құру
                df = pd.DataFrame(data_rows.values, columns=clean_names)
                
                # Бағандарды сандық түрге келтіру
                numeric_cols = ['Коэф', 'Ср.балл', 'Үштік', 'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
                for col in numeric_cols:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
                # '№' бағаны жоқ болса, қосу
                if '№' not in df.columns:
                    df.insert(0, '№', range(1, len(df) + 1))
                else:
                    df['№'] = range(1, len(df) + 1)
                
                # ФИО және Предмет бағандарындағы бос мәндерді толтыру
                df['ФИО'] = df['ФИО'].fillna('Белгісіз')
                df['Предмет'] = df['Предмет'].fillna('Белгісіз')
                
                # Рейтингті есептеу
                if 'Результативность' in df.columns and 'Жалпы итог' in df.columns:
                    pass  # есептеу қажет емес
                else:
                    df['Результативность'] = (df['Ср.балл'] * df['Коэф']) - (df['Үштік'] * 0.9)
                    df['Жалпы итог'] = df['Результативность'] + df['Внеклас.'] + df['Метод.'] + df['Админ.рейтинг']
                
                df['Результативность'] = df['Результативность'].round(1)
                df['Жалпы итог'] = df['Жалпы итог'].round(1)
                
                # Бос жолдарды өшіру
                df = df.dropna(subset=['ФИО'], how='all')
                df = df[df['ФИО'] != 'Белгісіз']
                
                st.session_state.df = df
                st.session_state.raw_data = df_raw
                st.success(f"✅ {len(df)} мұғалім жүктелді!")
                
                # Жүктелген кестені көрсету
                st.subheader("📋 Жүктелген кесте")
                st.dataframe(df, use_container_width=True, height=400)
                return True
            
        return False
            
    except Exception as e:
        st.error(f"❌ Жүктеу қатесі: {str(e)}")
        st.info("Файл құрылымын тексеріңіз. Бірінші жолда баған атаулары болуы керек.")
        return False

def save_to_github_direct():
    """GitHub API арқылы тікелей сақтау"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    token = st.session_state.github_token
    if not token:
        st.error("🔐 GitHub токені енгізілмеген!")
        return
    
    try:
        with st.spinner("💾 GitHub-қа сақталуда..."):
            output = BytesIO()
            save_df = st.session_state.df.copy()
            save_df.to_excel(output, index=False, engine='openpyxl')
            output.seek(0)
            content_b64 = base64.b64encode(output.getvalue()).decode('utf-8')
            
            headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
            get_response = requests.get(GITHUB_API_URL, headers=headers)
            sha = get_response.json().get('sha') if get_response.status_code == 200 else None
            
            data = {"message": "Update reiting1.xlsx", "content": content_b64, "branch": "main"}
            if sha:
                data["sha"] = sha
            
            put_response = requests.put(GITHUB_API_URL, headers=headers, json=data)
            
            if put_response.status_code in [200, 201]:
                st.success("✅ Файл GitHub-қа сәтті сақталды!")
            else:
                st.error(f"❌ GitHub API қатесі: {put_response.status_code}")
                
    except Exception as e:
        st.error(f"❌ Сақтау қатесі: {str(e)}")

def save_to_local():
    """Жергілікті файлға сақтау"""
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    output = BytesIO()
    st.session_state.df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    st.download_button(
        label="📥 Excel файлын жүктеу",
        data=output,
        file_name="reiting1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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
            coeff = st.number_input("Коэф.", value=1.0, step=0.1)
            avg_score = st.number_input("Ср.балл", value=4.0, step=0.1)
            triplets = st.number_input("Үштік саны", value=0, step=1)
        with col2:
            extra = st.number_input("Внеклас. работа", value=0.0, step=0.5)
            method = st.number_input("Метод. деятельность", value=0.0, step=0.5)
            admin = st.number_input("Админ. рейтинг", value=0.0, step=0.5)
        
        if st.form_submit_button("Қосу"):
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
    
    teacher_options = {f"{row['№']}. {row['ФИО']} - {row['Предмет']}": idx 
                       for idx, row in st.session_state.df.iterrows()}
    selected = st.selectbox("Өшіретін мұғалімді таңдаңыз:", list(teacher_options.keys()))
    
    if st.button("🗑 Өшіру", type="secondary"):
        idx = teacher_options[selected]
        teacher_name = st.session_state.df.loc[idx, 'ФИО']
        st.session_state.df = st.session_state.df.drop(idx).reset_index(drop=True)
        st.session_state.df['№'] = range(1, len(st.session_state.df) + 1)
        st.success(f"🗑 {teacher_name} өшірілді!")
        st.rerun()

# ==================== НЕГІЗГІ ИНТЕРФЕЙС ====================

show_token_form()

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
    st.dataframe(st.session_state.df, use_container_width=True, height=500)
    
    # Статистика
    st.subheader("📊 Статистика")
    valid_df = st.session_state.df.dropna(subset=['Жалпы итог'])
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📚 Мұғалімдер саны", len(valid_df))
    with col2:
        st.metric("📊 Орташа жалпы итог", f"{valid_df['Жалпы итог'].mean():.1f}")
    with col3:
        if len(valid_df) > 0:
            best = valid_df.loc[valid_df['Жалпы итог'].idxmax(), 'ФИО']
            st.metric("🏆 Ең жақсы нәтиже", f"{best} ({valid_df['Жалпы итог'].max():.1f})")
    with col4:
        if len(valid_df) > 0:
            worst = valid_df.loc[valid_df['Жалпы итог'].idxmin(), 'ФИО']
            st.metric("⚠️ Ең төмен нәтиже", f"{worst} ({valid_df['Жалпы итог'].min():.1f})")
    
    # CSV экспорт
    csv = valid_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(label="📥 CSV форматында жүктеу", data=csv, file_name="reiting_export.csv", mime="text/csv")

elif st.session_state.df is not None and len(st.session_state.df) == 0:
    st.warning("Деректер жоқ. Жаңа мұғалім қосыңыз!")
else:
    st.info("👆 Бастау үшін **«GitHub-тан жүктеу»** батырмасын басыңыз.")
    
    if st.button("📋 Үлгі кесте жасау"):
        sample_data = pd.DataFrame({
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
        })
        st.session_state.df = sample_data
        st.success("✅ Үлгі кесте жасалды!")
        st.rerun()

st.markdown("---")
st.caption("© Мұғалімдер рейтингі | GitHub-қа сақтау үшін сол жақ панельге токен енгізіңіз")