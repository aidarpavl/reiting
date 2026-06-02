import streamlit as st
import pandas as pd
import requests
import base64
from io import BytesIO

# ==================== GITHUB ПАРАМЕТРЛЕРІ ====================
GITHUB_RAW_URL = "https://raw.githubusercontent.com/aidarpavl/reiting/refs/heads/main/reiting1.xlsx"
GITHUB_API_URL = "https://api.github.com/repos/aidarpavl/reiting/contents/reiting1.xlsx"

# ==================== БАҒАН АТАУЛАРЫНЫҢ СӘЙКЕСТІГІ ====================
COLUMN_MAPPING = {
    'ФИО': ['фио', 'ф.и.о', 'фамилия', 'ф.и.о.', 'fio'],
    'Предмет': ['предмет', 'пән', 'subject', 'пәні'],
    'Коэф': ['коэф', 'коэффициент', 'коэф.', 'k', 'kоэф', 'коэфициент'],
    'Ср.балл': ['ср балл', 'ср.балл', 'средний балл', 'орташа балл', 'ср балл', 'ср_балл'],
    'Үштік': ['үштік', 'троек', 'кол-во итоговых троек', 'итоговые тройки', '3'],
    'Внеклас.': ['внеклас', 'внекласная работа', 'итог внекл.работа', 'внекл.работа'],
    'Метод.': ['метод', 'методическая деятельность', 'итог метод. деятельности'],
    'Админ.рейтинг': ['админ', 'рейтинг учителя', 'итог у администрации', 'админ.рейтинг'],
    'Результативность': ['результативность', 'результативность по предмету'],
    'Жалпы итог': ['жалпы итог', 'общий итог', 'итог']
}

def find_column_index(headers, target_variants):
    """Баған атауын әртүрлі нұсқалар бойынша іздеу"""
    for i, header in enumerate(headers):
        header_lower = str(header).lower().strip()
        for variant in target_variants:
            if variant.lower() in header_lower or header_lower == variant.lower():
                return i
    return None

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
            headers = df_raw.iloc[0].fillna('').astype(str).tolist()
            
            st.info(f"📋 Файлдан табылған бағандар: {', '.join(headers[:15])}...")
            
            # Бағандарды сәйкестендіру
            col_indices = {}
            for target, variants in COLUMN_MAPPING.items():
                idx = find_column_index(headers, variants)
                if idx is not None:
                    col_indices[target] = idx
                    st.success(f"✅ '{target}' → '{headers[idx]}'")
                else:
                    st.warning(f"⚠️ '{target}' бағаны табылмады. Іздеген нұсқалар: {variants}")
            
            # Міндетті бағандарды тексеру
            required = ['ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік']
            missing = [r for r in required if r not in col_indices]
            if missing:
                st.error(f"Келесі міндетті бағандар табылмады: {missing}")
                st.info("Excel файлының бірінші жолында баған атаулары дұрыс екенін тексеріңіз.")
                return False
            
            # Деректерді жинау
            data_rows = []
            for row_idx in range(1, len(df_raw)):
                row = df_raw.iloc[row_idx]
                
                # ФИО бос болса өткізіп жіберу
                fio_val = row[col_indices['ФИО']]
                if pd.isna(fio_val) or str(fio_val).strip() == '':
                    continue
                
                new_row = {}
                for target, col_idx in col_indices.items():
                    val = row[col_idx]
                    if target in ['Коэф', 'Ср.балл', 'Үштік', 'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']:
                        try:
                            new_row[target] = float(val) if pd.notna(val) else 0
                        except (ValueError, TypeError):
                            new_row[target] = 0
                    else:
                        new_row[target] = str(val) if pd.notna(val) else ''
                
                data_rows.append(new_row)
            
            if not data_rows:
                st.error("Деректер табылмады!")
                return False
            
            # DataFrame құру
            df = pd.DataFrame(data_rows)
            
            # Сандық бағандарды дұрыстау
            numeric_cols = ['Коэф', 'Ср.балл', 'Үштік', 'Внеклас.', 'Метод.', 'Админ.рейтинг']
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            # Результативность есептеу
            df['Результативность'] = (df['Ср.балл'] * df['Коэф']) - (df['Үштік'] * 0.9)
            
            # Жалпы итог есептеу
            df['Жалпы итог'] = df['Результативность'] + df['Внеклас.'] + df['Метод.'] + df['Админ.рейтинг']
            
            # Дөңгелектеу
            df['Результативность'] = df['Результативность'].round(1)
            df['Жалпы итог'] = df['Жалпы итог'].round(1)
            
            # Нөмірлерді қосу
            df.insert(0, '№', range(1, len(df) + 1))
            
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
    
    token = st.session_state.get('github_token')
    if not token:
        st.error("🔐 GitHub токені енгізілмеген! Сол жақ панельге токеніңізді енгізіңіз.")
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
                error_detail = put_response.json() if put_response.text else {"message": "Белгісіз қате"}
                st.error(f"❌ GitHub API қатесі: {error_detail.get('message', put_response.status_code)}")
                
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

# Сол жақ панель
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