from typing import List, Dict

GROUPS = {
    "Ареометры": "Ареометры всех типов (п.п. 57)",
    "Вискозиметры": "Вискозиметры условной вязкости (п.п. 53)",
    "Жидкости": "Аттестация жидкостей градуировочных для поверки вискозиметров (п.п. 54)",
    "Рефрактометры": "Рефрактометры (п.п. 58)"
}

def determine_group(work_name: str) -> str:
    name_lower = work_name.lower()
    # Жидкость и градуировочная – самый приоритет
    if "жидкость" in name_lower or "градуировочная" in name_lower:
        return "Жидкости"
    if "ареометр" in name_lower:
        return "Ареометры"
    if "вискозиметр" in name_lower or "реометр" in name_lower:
        return "Вискозиметры"
    if "рефрактометр" in name_lower:
        return "Рефрактометры"
    return None

def group_works(works: List[Dict]) -> Dict[str, List[Dict]]:
    result = {full_name: [] for full_name in GROUPS.values()}
    for w in works:
        group_key = determine_group(w['work_name'])
        if group_key:
            full_name = GROUPS[group_key]
            result[full_name].append(w)
    return result