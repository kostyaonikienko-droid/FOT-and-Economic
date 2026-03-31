import sqlite3
from typing import List, Dict, Optional
from utils.file_utils import get_db_path
import datetime

DB_PATH = get_db_path('spirit.db')

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS works (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date_work TEXT,
            date_invoice TEXT,
            invoice_num TEXT,
            amount REAL,
            work_name TEXT,
            quantity INTEGER,
            tab_number TEXT,
            import_date TEXT,
            group_name TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS act_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            setting_key TEXT UNIQUE,
            setting_value TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS spirit_norms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_name TEXT UNIQUE,
            norm_liters REAL
        )
    ''')

    cursor.execute("SELECT COUNT(*) FROM spirit_norms")
    if cursor.fetchone()[0] == 0:
        default_norms = [
            ("Ареометры", 0.005),
            ("Вискозиметры", 0.015),
            ("Жидкости", 0.1),
            ("Рефрактометры", 0.005)
        ]
        cursor.executemany("INSERT INTO spirit_norms (group_name, norm_liters) VALUES (?, ?)", default_norms)

    cursor.execute("SELECT COUNT(*) FROM act_settings")
    if cursor.fetchone()[0] == 0:
        default_settings = [
            ("act_number", ""),
            ("act_date", ""),
            ("organization", "ФБУ \"Пермский ЦСМ\""),
            ("commission_chairman", "зам. директора по инновац. развитию Карташев А.Л."),
            ("commission_members", "гл. бухгалтер Рожкова Е.Д., бухгалтер Мартемьянова О.А., бухгалтер Кожевникова Н.А., гл. метролог Кудрявцева О.А."),
            ("responsible_person", "И.о. начальника отдела Оникиенко К.С."),
            ("economist", "экономист Гуляева Е.Н."),
            ("antiseptic_solution_ticket", "№ 947 от 03.12.2025 в кол-ве 5 литров"),
            ("antiseptic_inv_number", "321947"),
            ("antiseptic_amount", "5")
        ]
        cursor.executemany("INSERT INTO act_settings (setting_key, setting_value) VALUES (?, ?)", default_settings)

    conn.commit()
    conn.close()

def save_works(works: List[Dict], import_date: str):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    for w in works:
        cursor.execute('''
            SELECT id FROM works
            WHERE date_work = ? AND date_invoice = ? AND invoice_num = ? AND work_name = ?
        ''', (w['date_work'], w['date_invoice'], w['invoice_num'], w['work_name']))
        existing = cursor.fetchone()
        if existing:
            cursor.execute('''
                UPDATE works SET quantity = ?, tab_number = ?, import_date = ?
                WHERE id = ?
            ''', (w['quantity'], w['tab_number'], import_date, existing[0]))
        else:
            cursor.execute('''
                INSERT INTO works (date_work, date_invoice, invoice_num, amount, work_name, quantity, tab_number, import_date, group_name)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (w['date_work'], w['date_invoice'], w['invoice_num'], w['amount'],
                  w['work_name'], w['quantity'], w['tab_number'], import_date,
                  w.get('group_name', '')))
    conn.commit()
    conn.close()

def update_work_quantity(work_id: int, new_quantity: int):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE works SET quantity = ? WHERE id = ?", (new_quantity, work_id))
    conn.commit()
    conn.close()

def get_works_by_period(start_date: str, end_date: str) -> List[Dict]:
    """Возвращает работы за период (дата выполнения). Фильтрация по реальной дате."""
    all_works = get_all_works()
    # Преобразуем границы в объекты date
    try:
        start_parts = start_date.split('.')
        end_parts = end_date.split('.')
        start_dt = datetime.date(int(start_parts[2]), int(start_parts[1]), int(start_parts[0]))
        end_dt = datetime.date(int(end_parts[2]), int(end_parts[1]), int(end_parts[0]))
    except:
        # Если не удалось разобрать, вернуть пустой список
        return []

    filtered = []
    for w in all_works:
        try:
            parts = w['date_work'].split('.')
            w_date = datetime.date(int(parts[2]), int(parts[1]), int(parts[0]))
            if start_dt <= w_date <= end_dt:
                filtered.append(w)
        except:
            continue
    return filtered

def get_all_works() -> List[Dict]:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM works ORDER BY date_work")
    rows = cursor.fetchall()
    works = [dict(row) for row in rows]
    conn.close()
    return works

def delete_works_by_ids(ids: List[int]):
    if not ids:
        return
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    placeholders = ','.join('?' * len(ids))
    cursor.execute(f"DELETE FROM works WHERE id IN ({placeholders})", ids)
    conn.commit()
    conn.close()

def get_setting(key: str) -> Optional[str]:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT setting_value FROM act_settings WHERE setting_key = ?", (key,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else None

def set_setting(key: str, value: str):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO act_settings (setting_key, setting_value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()

def get_all_settings() -> Dict[str, str]:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT setting_key, setting_value FROM act_settings")
    rows = cursor.fetchall()
    settings = {row[0]: row[1] for row in rows}
    conn.close()
    return settings

def get_spirit_norm(group_name: str) -> float:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT norm_liters FROM spirit_norms WHERE group_name = ?", (group_name,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else 0.0