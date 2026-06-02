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
if 'github_token' not in st.session_state:
    st.session_state.github_token = None

# ==================== ФУНКЦИЯЛАР ====================
def show_token_form():
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
    """GitHub-тан Excel файлын жүктеу"""
    try:
        with st.spinner("📥 GitHub-тан жүктелуде..."):
            response = requests.get(GITHUB_RAW_URL, timeout=30)
            response.raise_for_status()
            
            # Excel оқу
            excel_data = BytesIO(response.content)
            df_raw = pd.read_excel(excel_data, engine='openpyxl', header=None)
            
            # Бірінші жолдағы баған атаулары
            first_row = df_raw.iloc[0].fillna('').astype(str).tolist()
            
            # Қажетті бағандардың индекстерін табу
            col_mapping = {}
            for i, name in enumerate(first_row):
                name_lower = str(name).lower().strip()
                if 'фио' in name_lower or 'ф.и.о' in name_lower or 'фамилия' in name_lower:
                    col_mapping['ФИО'] = i
                elif 'предмет' in name_lower:
                    col_mapping['Предмет'] = i
                elif 'коэф' in name_lower or 'коэффициент' in name_lower:
                    col_mapping['Коэф'] = i
                elif 'ср.балл' in name_lower or 'средний балл' in name_lower or 'средний' in name_lower:
                    col_mapping['Ср.балл'] = i
                elif 'үштік' in name_lower or 'троек' in name_lower or '3' in name_lower:
                    col_mapping['Үштік'] = i
                elif 'результативность' in name_lower:
                    col_mapping['Результативность'] = i
                elif 'внеклас' in name_lower:
                    col_mapping['Внеклас.'] = i
                elif 'метод' in name_lower:
                    col_mapping['Метод.'] = i
                elif 'админ' in name_lower or 'рейтинг' in name_lower:
                    col_mapping['Админ.рейтинг'] = i
                elif 'жалпы' in name_lower or 'итог' in name_lower:
                    col_mapping['Жалпы итог'] = i
            
            # Қажетті бағандардың тізімі
            required_cols = ['ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік']
            missing_cols = [col for col in required_cols if col not in col_mapping]
            
            if missing_cols:
                st.warning(f"Келесі бағандар табылмады: {missing_cols}")
                st.info("Баған атаулары: " + ", ".join([f"'{x}'" for x in first_row if x]))
                return False
            
            # Деректерді жинау
            data_rows = []
            for idx in range(1, len(df_raw)):
                row = df_raw.iloc[idx]
                # Егер ФИО бос болса, өткізіп жіберу
                if pd.isna(row[col_mapping['ФИО']]):
                    continue
                    
                new_row = {}
                for key, col_idx in col_mapping.items():
                    val = row[col_idx]
                    if key in ['Коэф', 'Ср.балл', 'Үштік', 'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']:
                        try:
                            new_row[key] = float(val) if pd.notna(val) else 0
                        except:
                            new_row[key] = 0
                    else:
                        new_row[key] = str(val) if pd.notna(val) else ''
                
                data_rows.append(new_row)
            
            if not data_rows:
                st.error("Деректер табылмады!")
                return False
            
            # DataFrame құру
            df = pd.DataFrame(data_rows)
            
            # Сандық бағандарды дұрыстау
            for col in ['Коэф', 'Ср.балл', 'Үштік', 'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # Нөмірлерді қосу
            df.insert(0, '№', range(1, len(df) + 1))
            
            # Егер Результативность және Жалпы итог бағандары болмаса, есептеу
            if 'Результативность' not in df.columns:
                df['Результативность'] = (df['Ср.балл'] * df['Коэф']) - (df['Үштік'] * 0.9)
            if 'Жалпы итог' not in df.columns:
                df['Жалпы итог'] = df['Результативность'] + df['Внеклас.'] + df['Метод.'] + df['Админ.рейтинг']
            
            df['Результативность'] = df['Результативность'].round(1)
            df['Жалпы итог'] = df['Жалпы итог'].round(1)
            
            st.session_state.df = df
            st.success(f"✅ {len(df)} мұғалім жүктелді!")
            
            # Кестені көрсету
            st.subheader("📋 Жүктелген кесте")
            st.dataframe(df, use_container_width=True, height=400)
            return True
            
    except Exception as e:
        st.error(f"❌ Жүктеу қатесі: {str(e)}")
        return False

def save_to_github_direct():
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
            save_df = save_df.drop(columns=['№'], errors='ignore')
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
    if st.session_state.df is None:
        st.warning("Алдымен файлды жүктеңіз!")
        return
    
    output = BytesIO()
    save_df = st.session_state.df.copy()
    save_df = save_df.drop(columns=['№'], errors='ignore')
    save_df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    
    st.download_button(
        label="📥 Excel файлын жүктеу",
        data=output,
        file_name="reiting1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def calculate_rating():
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
    
    if len(valid_df) > 0:
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📚 Мұғалімдер саны", len(valid_df))
        with col2:
            st.metric("📊 Орташа жалпы итог", f"{valid_df['Жалпы итог'].mean():.1f}")
        with col3:
            best_idx = valid_df['Жалпы итог'].idxmax()
            best = valid_df.loc[best_idx, 'ФИО']
            st.metric("🏆 Ең жақсы нәтиже", f"{best} ({valid_df['Жалпы итог'].max():.1f})")
        with col4:
            worst_idx = valid_df['Жалпы итог'].idxmin()
            worst = valid_df.loc[worst_idx, 'ФИО']
            st.metric("⚠️ Ең төмен нәтиже", f"{worst} ({valid_df['Жалпы итог'].min():.1f})")
    
    # CSV экспорт
    csv = st.session_state.df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(label="📥 CSV форматында жүктеу", data=csv, file_name="reiting_export.csv", mime="text/csv")

else:
    st.info("👆 Бастау үшін **«GitHub-тан жүктеу»** батырмасын басыңыз.")

st.markdown("---")
st.caption("© Мұғалімдер рейтингі | GitHub-қа сақтау үшін сол жақ панельге токен енгізіңіз")