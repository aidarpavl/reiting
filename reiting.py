import pandas as pd
import requests
from io import BytesIO
import os
from tabulate import tabulate

# GitHub файл сілтемесі
GITHUB_URL = "https://raw.githubusercontent.com/aidarpavl/reiting/refs/heads/main/reiting1.xlsx"
LOCAL_FILE = "reiting1.xlsx"

def load_from_github():
    """GitHub-тан Excel файлын жүктеп алу"""
    try:
        response = requests.get(GITHUB_URL)
        response.raise_for_status()
        # Жүктелген файлды жергілікті сақтау
        with open(LOCAL_FILE, 'wb') as f:
            f.write(response.content)
        print("✅ Файл GitHub-тан сәтті жүктелді!")
        return True
    except Exception as e:
        print(f"❌ GitHub-тан жүктеу қатесі: {e}")
        return False

def save_to_github(df):
    """Өзгерістерді жергілікті файлға сақтайды (GitHub-қа жүктеу нұсқаулығымен)"""
    try:
        df.to_excel(LOCAL_FILE, index=False, engine='openpyxl')
        print("✅ Файл жергілікті компьютерге сақталды!")
        print("\n📌 GitHub-қа жүктеу үшін келесі қадамдарды орындаңыз:")
        print("   1. https://github.com/aidarpavl/reiting сайтына кіріңіз")
        print("   2. 'reiting1.xlsx' файлын тауып, 'Update' немесе 'Upload' батырмасын басыңыз")
        print(f"   3. Жаңартылған '{LOCAL_FILE}' файлын таңдап, жүктеңіз")
        return True
    except Exception as e:
        print(f"❌ Файлды сақтау қатесі: {e}")
        return False

def calculate_rating(row):
    """Бір мұғалімнің рейтингін есептеу"""
    # Результативность по предмету = (Ср.балл × Коэф) – (Үштік × 0.9)
    subject_result = (row['Ср.балл'] * row['Коэф']) - (row['Үштік'] * 0.9)
    # Жалпы итог
    total = subject_result + row['Внеклас.'] + row['Метод.'] + row['Админ.рейтинг']
    return round(subject_result, 1), round(total, 1)

def display_table(df):
    """Кестені терминалда көрсету"""
    print("\n" + "="*100)
    print("МҰҒАЛІМДЕР РЕЙТИНГІ")
    print("="*100)
    # Көрсету үшін қажетті бағандар
    display_df = df[['№', 'ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік', 
                     'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']].copy()
    print(tabulate(display_df, headers='keys', tablefmt='grid', showindex=False))
    print("="*100)

def add_teacher(df):
    """Жаңа мұғалім қосу"""
    print("\n--- ЖАҢА МҰҒАЛІМ ҚОСУ ---")
    name = input("ФИО: ")
    subject = input("Пәні: ")
    coeff = float(input("Коэффициент (мын. 1.0, 1.1, 0.9): "))
    avg = float(input("Орташа балл (0-5): "))
    triplets = int(input("Үштік саны: "))
    extra = float(input("Внекласная работа (ұпай): "))
    method = float(input("Методическая деятельность (ұпай): "))
    admin = float(input("Администрация рейтингі (ұпай): "))
    
    new_id = df['№'].max() + 1 if len(df) > 0 else 1
    
    new_row = {
        '№': new_id,
        'ФИО': name,
        'Предмет': subject,
        'Коэф': coeff,
        'Ср.балл': avg,
        'Үштік': triplets,
        'Внеклас.': extra,
        'Метод.': method,
        'Админ.рейтинг': admin,
        'Результативность': 0,
        'Жалпы итог': 0
    }
    
    # Рейтинг есептеу
    subject_res, total = calculate_rating(new_row)
    new_row['Результативность'] = subject_res
    new_row['Жалпы итог'] = total
    
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    # Нөмірлерді қайта реттеу
    df['№'] = range(1, len(df) + 1)
    print(f"✅ {name} қосылды!")
    return df

def delete_teacher(df):
    """Мұғалімді өшіру"""
    if len(df) == 0:
        print("Өшіретін мұғалім жоқ!")
        return df
    
    print("\n--- МҰҒАЛІМДІ ӨШІРУ ---")
    display_table(df)
    try:
        idx = int(input("Өшіргіңіз келетін мұғалімнің № нөмірін енгізіңіз: "))
        if idx < 1 or idx > len(df):
            print("Қате нөмір!")
            return df
        
        deleted_name = df.loc[df['№'] == idx, 'ФИО'].values[0]
        df = df[df['№'] != idx].reset_index(drop=True)
        df['№'] = range(1, len(df) + 1)
        print(f"✅ {deleted_name} өшірілді!")
        return df
    except ValueError:
        print("Қате формат! Сан енгізіңіз.")
        return df

def recalculate_all(df):
    """Барлық мұғалімдердің рейтингін қайта есептеу"""
    for idx, row in df.iterrows():
        subject_res, total = calculate_rating(row)
        df.at[idx, 'Результативность'] = subject_res
        df.at[idx, 'Жалпы итог'] = total
    print("✅ Рейтингтер қайта есептелді!")
    return df

def main():
    print("🚀 Мұғалімдер рейтингі – GitHub интеграциясы")
    print("-" * 50)
    
    # GitHub-тан жүктеу
    if not load_from_github():
        print("Файл жүктелмеді. Қолмен тексеріңіз.")
        return
    
    # Excel файлын оқу
    try:
        df = pd.read_excel(LOCAL_FILE, engine='openpyxl')
        # Баған атауларын стандарттау
        df.columns = ['№', 'ФИО', 'Предмет', 'Коэф', 'Ср.балл', 'Үштік', 
                      'Результативность', 'Внеклас.', 'Метод.', 'Админ.рейтинг', 'Жалпы итог']
        print("✅ Файл сәтті оқылды!")
    except Exception as e:
        print(f"❌ Файлды оқу қатесі: {e}")
        return
    
    # Негізгі цикл
    while True:
        print("\n" + "="*50)
        print("📋 МЕНЮ:")
        print("1. Кестені көрсету")
        print("2. Жаңа мұғалім қосу")
        print("3. Мұғалімді өшіру")
        print("4. Рейтингті қайта есептеу")
        print("5. Өзгерістерді сақтау (жергілікті)")
        print("6. Шығу")
        print("="*50)
        
        choice = input("Таңдауыңыз (1-6): ")
        
        if choice == '1':
            display_table(df)
        elif choice == '2':
            df = add_teacher(df)
        elif choice == '3':
            df = delete_teacher(df)
        elif choice == '4':
            df = recalculate_all(df)
            display_table(df)
        elif choice == '5':
            save_to_github(df)
        elif choice == '6':
            save = input("Өзгерістерді сақтап шығу? (y/n): ")
            if save.lower() == 'y':
                save_to_github(df)
            print("👋 Бағдарлама аяқталды!")
            break
        else:
            print("Қате таңдау! 1-6 аралығында сан енгізіңіз.")

if __name__ == "__main__":
    # Қажетті кітапханаларды орнату нұсқауы
    try:
        import pandas
        import requests
        import tabulate
        import openpyxl
    except ImportError as e:
        print("❌ Қажетті кітапханалар орнатылмаған!")
        print("Терминалға келесі команданы жазыңыз:")
        print("pip install pandas requests tabulate openpyxl")
        exit(1)
    
    main()