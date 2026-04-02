import pandas as pd
import re
from typing import List, Dict, Optional

def parse_excel(file_path: str) -> List[Dict]:
    """
    Парсит Excel-файл, поддерживая два формата:
    1. Единая таблица с заголовками (старый формат).
    2. Блочный формат с заголовками и строками "Таб № X".
    """
    df = pd.read_excel(file_path, header=None, dtype=str, keep_default_na=False)
    df = df.fillna('')

    # Ищем строку, содержащую заголовки
    header_row_idx = None
    col_mapping = {}
    quantity_col = None
    col_keywords = {
        'date_work': ['дата вып', 'дата выполн'],
        'date_invoice': ['дата сч', 'дата счета'],
        'invoice_num': ['№ сч', 'номер сч', '№ счета'],
        'amount': ['сумма'],
        'work_name': ['работа', 'наименование']
    }
    quantity_keywords = ['к-во', 'кол-во', 'количество']

    for idx, row in df.iterrows():
        for col_idx, cell in enumerate(row):
            cell_str = str(cell).strip().lower()
            if not cell_str:
                continue
            for key, keywords in col_keywords.items():
                if any(kw in cell_str for kw in keywords):
                    col_mapping[key] = col_idx
            if any(kw in cell_str for kw in quantity_keywords):
                quantity_col = col_idx

        # Если найдены все ключевые колонки – это строка заголовков
        if all(k in col_mapping for k in ['date_work', 'date_invoice', 'invoice_num', 'amount', 'work_name']):
            header_row_idx = idx
            break

    if header_row_idx is None:
        raise ValueError("Не удалось найти заголовки таблицы в файле")

    max_cols = df.shape[1]
    works = []
    current_tab = None

    # Начинаем искать данные после строки заголовков
    for idx in range(header_row_idx + 1, len(df)):
        row = df.iloc[idx]
        first_cell = str(row.iloc[0]).strip() if max_cols > 0 else ''

        # Пропускаем итоговые строки
        if first_cell == '' or 'итого' in first_cell.lower() or 'всего' in first_cell.lower():
            continue

        # Обнаружена строка с табельным номером
        if 'таб №' in first_cell.lower():
            match = re.search(r'таб\s*№\s*(\d+)', first_cell, re.IGNORECASE)
            if match:
                current_tab = match.group(1)
                # Если табельный номер 0 – это служебные записи, не будем их использовать
                if current_tab == '0':
                    current_tab = None
            continue

        # Если это пустая строка после блока – пропускаем
        if not first_cell and all(str(c).strip() == '' for c in row):
            continue

        # Извлекаем данные по известным индексам
        date_work = row.iloc[col_mapping['date_work']] if col_mapping['date_work'] < max_cols else ''
        date_invoice = row.iloc[col_mapping['date_invoice']] if col_mapping['date_invoice'] < max_cols else ''
        invoice_num = row.iloc[col_mapping['invoice_num']] if col_mapping['invoice_num'] < max_cols else ''
        amount = row.iloc[col_mapping['amount']] if col_mapping['amount'] < max_cols else ''
        work_name = row.iloc[col_mapping['work_name']] if col_mapping['work_name'] < max_cols else ''

        # Если нет ни даты выполнения, ни названия работы – пропускаем
        if not date_work and not work_name:
            continue

        # Пропускаем служебные строки с нулевыми временными метками
        date_work_str = str(date_work).strip()
        if date_work_str == '0:00:00':
            continue
        # Пропускаем служебные строки с *** MoveW
        if '*** movew' in str(work_name).lower():
            continue

        date_invoice_str = str(date_invoice).strip()
        invoice_num = str(invoice_num).strip()
        amount_str = str(amount).replace(',', '.').strip()
        work_name = str(work_name).strip()

        # Фильтрация мусорных строк (старый фильтр)
        if date_work_str.startswith('1900-01-01'):
            continue
        if invoice_num == '0':
            continue

        # Количество (если есть)
        quantity = 0
        if quantity_col is not None and quantity_col < max_cols:
            qty_val = row.iloc[quantity_col]
            if qty_val and str(qty_val).strip():
                try:
                    quantity = int(float(qty_val))
                except:
                    quantity = 0
        if quantity == 0:
            # Пропускать строки с нулевым количеством только если это не спирт? Пока пропускаем.
            # Но в новых файлах нет колонки количества, поэтому quantity останется 0.
            # В новых файлах количество всегда 1, так как нет колонки к-во. Поэтому нужно по умолчанию ставить 1, если нет колонки.
            # Поскольку quantity_col может быть None, нужно понимать: если колонки количества нет, то quantity = 1.
            # Уточним: в старых файлах количество было в отдельной колонке, в новых – нет.
            # Поэтому, если quantity_col is None, считаем, что количество = 1.
            pass

        # Если колонка количества не найдена, считаем количество = 1
        if quantity_col is None:
            quantity = 1
        else:
            # Если нашли, но значение 0, то может быть строка без количества – тоже ставим 1
            if quantity == 0:
                quantity = 1

        # Обработка дат: если есть время, отрезаем его (оставляем только дату)
        if ' ' in date_work_str:
            date_work_str = date_work_str.split(' ')[0]
        if ' ' in date_invoice_str:
            date_invoice_str = date_invoice_str.split(' ')[0]

        try:
            amount_val = float(amount_str) if amount_str else 0.0
        except:
            amount_val = 0.0

        works.append({
            'date_work': date_work_str,
            'date_invoice': date_invoice_str,
            'invoice_num': invoice_num,
            'amount': amount_val,
            'work_name': work_name,
            'quantity': quantity,
            'tab_number': current_tab
        })

    return works