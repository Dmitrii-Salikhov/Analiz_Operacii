# analyzers/form14_map.py
"""
Сопоставление кодов A16 (номенклатура услуг) со строками ФСН № 14
таблиц 4000 / 4001 «Хирургическая работа организации».

Строки — по приказу Росстата (форма № 14, ред. 2025: пр. 42 / 712).
Маппинг практический (анатомия кода + ключевые слова), не официальный CSV Минздрава.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Sequence, Tuple

# Основные строки 4000/4001 (одинаковый перечень видов; 4001 — срез по возрасту)
FORM14_LINES: Dict[str, str] = {
    "1": "Всего операций",
    "2": "операции на нервной системе",
    "2.1": "удаление травматической внутричерепной гематомы / очага ушиба / вдавленного перелома",
    "2.5": "удаление опухолей головного, спинного мозга",
    "2.7": "декомпрессивные, стабилизирующие при позвоночно-спинальной травме",
    "2.8": "декомпрессивные, стабилизирующие при дегенеративных заболеваниях позвоночника",
    "2.9": "операции на периферических нервах",
    "2.10": "ликворошунтирующие операции",
    "3": "операции на эндокринной системе",
    "3.1": "тиреотомии",
    "4": "операции на органе зрения",
    "5": "операции на органах уха, горла, носа",
    "5.1": "из них — на ухе",
    "5.2": "из них — на миндалинах и аденоидах",
    "6": "операции на органах дыхания",
    "6.1": "из них — на трахее",
    "6.2": "пневмонэктомия",
    "6.3": "эксплоративная торакотомия",
    "7": "операции на сердце",
    "8": "операции на сосудах",
    "8.1": "из них — на артериях",
    "8.2": "из них — на венах",
    "9": "операции на органах брюшной полости",
    "9.1": "на желудке по поводу язвенной болезни",
    "9.2": "аппендэктомии при хроническом аппендиците",
    "9.3": "грыжесечение при неущемленной грыже",
    "9.4": "холецистэктомия при хроническом холецистите",
    "9.5": "лапаротомия диагностическая",
    "9.6": "на кишечнике",
    "9.6.1": "из них на прямой кишке",
    "9.7": "по поводу геморроя",
    "10": "операции на почках и мочеточниках",
    "11": "операции на мужских половых органах",
    "11.1": "из них — на предстательной железе",
    "12": "операции по поводу стерилизации мужчин",
    "13": "операции на женских половых органах",
    "14": "акушерские операции",
    "15": "операции на костно-мышечной системе",
    "15.1": "корригирующие остеотомии",
    "15.2": "на челюстно-лицевой области",
    "15.3": "при травмах костей таза",
    "15.4": "при около- и внутрисуставных переломах",
    "15.5": "на позвоночнике",
    "16": "операции на молочной железе",
    "17": "операции на коже и подкожной клетчатке",
    "18": "операции на средостении",
    "18.1": "из них — на вилочковой железе",
    "19": "операции на пищеводе",
    "20": "операции на лимфатической системе",
    "21": "прочие операции",
}

# Класс A16.XX → базовая строка ФСН (до уточнения по ключевым словам)
A16_CLASS_DEFAULT: Dict[str, str] = {
    "01": "17",  # кожа, ПЖК
    "02": "15",  # мышцы, сухожилия
    "03": "15",  # кости
    "04": "15",  # суставы
    "05": "9",  # селезёнка / кроветворение → брюшная
    "06": "20",  # лимфоузлы / тимус (уточняется)
    "07": "15.2",  # полость рта / ЧЛО
    "08": "5",  # дыхательные пути верхние / ЛОР
    "09": "6",  # лёгкие / плевра
    "10": "7",  # сердце
    "11": "18",  # средостение
    "12": "8",  # сосуды
    "14": "9",  # печень, ЖВП
    "15": "9",  # поджелудочная
    "16": "9",  # пищевод/желудок/ДПК (пищевод может → 19)
    "17": "9.6",  # тонкая кишка
    "18": "9.6",  # толстая кишка
    "19": "9.6.1",  # прямая кишка / анус
    "20": "13",  # женские половые
    "21": "11",  # мужские половые
    "22": "3",  # щитовидная / эндокринная
    "23": "2",  # ЦНС / череп
    "24": "2.9",  # периферические нервы
    "25": "5.1",  # ухо
    "26": "4",  # глаз
    "27": "5",  # придаточные пазухи носа
    "28": "10",  # почки, мочеточники
    "30": "9",  # прочие / грыжи / лапаротомия — уточняется
}

_CODE_RE = re.compile(r"A16\.(\d{2})(?:\.(\d{3})(?:\.(\d{3}))?)?", re.I)

# Калибровка по формулам сводной ЛОР (config form_4001.line_categories / LINE_TOTAL_CATS).
# Локальная практика: нос/глотка/гортань → стр. 6 (не 5). Трахеостомия → 6.1.
# Имена категорий — как в surgery_categories ЛОР.
MANUAL_LINE_BY_CATEGORY: Dict[str, str] = {
    "Аденотомия": "5.2",
    "Антромастоидотомия": "5.1",
    "Биопсия гортани ": "6",
    "Биопсия гортани": "6",
    "Гайморотомия": "6",
    "Заглоточный абсцесс": "6",
    "Миринготомия план": "5.1",
    "Миринготомия экстр": "5.1",
    "Наложение вторичных швов": "17",
    "Остановка кров": "6",
    "ПТА": "5.2",
    "ПХО": "17",
    "Пластика местными тканями": "17",
    "Пластика нёба": "6",
    "Пластика раковин": "6",
    "Полипотомия": "6",
    "Ревизия п/о полости": "6",
    "Репозиция костей носа": "6",
    "Рубцы мочки уха": "5.1",
    "Септопластика": "6",
    "Синехии нос": "6",
    "Субперостальный абсцесс за ухом": "5.1",
    "Тозиллэктомия": "5.2",
    "Тонзиллотомия": "5.2",
    "Трахеостомия": "6.1",
    "Увулопластика": "6",
    "Удаление инородного тела": "6",
    "Удаление новообр глотки": "6",
    "Удаление новообр гортани": "6",
    "Удаление новообр носа": "6",
    "Удаление новообр уха": "5.1",
    "Флегмона шеи": "6",
    "Фронтотомия": "6",
    "Фурункул НСП": "5.1",
    "Фурункул носа": "6",
}

# Коды из рубрикатора ЛОР → та же калибровка (для map_code_to_form14 без имени категории).
MANUAL_LINE_BY_CODE: Dict[str, str] = {
    "A11.08.001": "6",
    "A16.01.008.001": "17",
    "A16.01.010.002": "17",
    "A16.01.011": "5.1",
    "A16.01.031": "5.1",
    "A16.03.034.002": "6",
    "A16.07.087": "6",
    "A16.08.001.001": "5.2",
    "A16.08.002.001": "5.2",
    "A16.08.003": "6.1",
    "A16.08.009.001": "6",
    "A16.08.010.003": "6",
    "A16.08.012": "5.2",
    "A16.08.013.001": "6",
    "A16.08.014": "6",
    "A16.08.017.001": "6",
    "A16.08.018": "6",
    "A16.08.035.001": "6",
    "A16.08.040": "6",
    "A16.08.040.003": "6",
    "A16.08.040.008": "6",
    "A16.08.054": "6",
    "A16.08.054.002": "6",
    "A16.08.055": "6",
    "A16.08.064": "5.2",
    "A16.08.071": "6",
    "A16.12.020.001": "6",
    "A16.25.001": "5.1",
    "A16.25.008": "6",
    "A16.25.008.001": "6",
    "A16.25.011": "5.1",
    "A16.25.035": "5.1",
    "A16.25.040": "5.1",
}


def _manual_line_for_category(name: str) -> Optional[str]:
    raw = str(name or "")
    if raw in MANUAL_LINE_BY_CATEGORY:
        return MANUAL_LINE_BY_CATEGORY[raw]
    stripped = raw.strip()
    if stripped in MANUAL_LINE_BY_CATEGORY:
        return MANUAL_LINE_BY_CATEGORY[stripped]
    return None


@dataclass
class Form14Mapping:
    code: str
    name: str
    a16_class: str
    line: str
    line_name: str
    confidence: str  # high | medium | low
    rule: str
    notes: str = ""


def parse_a16_class(code: str) -> str:
    m = _CODE_RE.match(str(code or "").strip())
    return m.group(1) if m else ""


def _has_any(text: str, words: Sequence[str]) -> bool:
    t = (text or "").lower()
    return any(w.lower() in t for w in words)


def refine_line_by_keywords(code: str, name: str, base_line: str) -> Tuple[str, str, str]:
    """
    Уточняет строку по названию. Возвращает (line, confidence, rule).
    """
    text = f"{code} {name}".lower()
    cls = parse_a16_class(code)

    # --- подстроки ЛОР ---
    if _has_any(
        text,
        ["аденоид", "аденотоми", "тонзил", "тозилл", "паратонзил", "миндалин", "пта"],
    ):
        return "5.2", "high", "keyword: миндалины/аденоиды"
    if cls == "25" or _has_any(
        text, ["миринготоми", "мастоидит", "антромастоид", "наружного уха", "среднего уха", "ушн"]
    ):
        if cls in ("25", "08") or "уха" in text or "ухо" in text:
            return "5.1", "high", "keyword/class: ухо"
    if cls == "08" and _has_any(text, ["трахеостом", "трахеи"]):
        return "6.1", "high", "keyword: трахея (из ЛОР-кодов)"

    # --- дыхание ---
    if _has_any(text, ["трахеостом", "трахеotom", "на трахее", "трахеи"]):
        return "6.1", "high", "keyword: трахея"
    if _has_any(text, ["пневмонэктоми", "пневмонэктомия"]):
        return "6.2", "high", "keyword: пневмонэктомия"
    if _has_any(text, ["эксплоративн", "торакотоми"]):
        return "6.3", "medium", "keyword: торакотомия"
    if _has_any(text, ["торакоцентез", "плевральн", "дренирование плевр"]):
        return "6", "high", "keyword: плевра/лёгкие"

    # --- брюшная полость ---
    if _has_any(text, ["аппендэктоми", "аппендицит"]):
        return "9.2", "high", "keyword: аппендэктомия"
    if _has_any(text, ["грыж", "герниопласт", "герниотом"]):
        return "9.3", "high", "keyword: грыжа"
    if _has_any(text, ["холецистэктоми", "желчного пузыр", "холецист"]):
        return "9.4", "high", "keyword: желчный пузырь"
    if _has_any(text, ["лапаротоми"]) and _has_any(text, ["диагностич", "релапаротом", "эксплоратив"]):
        return "9.5", "medium", "keyword: лапаротомия диагностическая"
    if _has_any(text, ["лапаротоми", "релапаротом"]):
        return "9.5", "medium", "keyword: лапаротомия"
    if _has_any(text, ["геморро", "геморроидаль"]):
        return "9.7", "high", "keyword: геморрой"
    if _has_any(text, ["прям", "анус", "анальн", "парапроктит", "свищ прямой", "трещин"]):
        if _has_any(text, ["кишк", "тонк", "толст", "ободочн"]):
            return "9.6", "medium", "keyword: кишка"
        return "9.6.1", "high", "keyword: прямая кишка/анус"
    if _has_any(text, ["язвенн"]) and _has_any(text, ["желуд", "двенадцатиперст"]):
        return "9.1", "high", "keyword: язва желудка/ДПК"
    if cls == "16" and _has_any(text, ["пищевод"]):
        return "19", "high", "keyword: пищевод"
    if cls == "16" and _has_any(text, ["желуд", "гастр"]):
        return "9", "high", "class+keyword: желудок → брюшная"
    if _has_any(text, ["селезен", "спленэктоми"]):
        return "9", "high", "keyword: селезёнка → брюшная"
    if _has_any(text, ["печен", "желчн", "холедох", "панкреат", "поджелудочн"]):
        return "9", "high", "keyword: гепатопанкреатобилиарная → брюшная"

    # --- мочеполовая ---
    if _has_any(text, ["простат", "предстательн"]):
        return "11.1", "high", "keyword: простата"
    if _has_any(text, ["стерилизац"]) and _has_any(text, ["мужч", "семян", "вазэктоми"]):
        return "12", "high", "keyword: стерилизация мужчин"
    if _has_any(text, ["кесарев", "акушер", "внематочн", "аборт", "вакуум-экстрак"]):
        return "14", "high", "keyword: акушерство"
    if _has_any(text, ["мочев", "уретр"]) and not _has_any(text, ["почк", "мочеточн"]):
        return "21", "medium", "keyword: мочевой пузырь/уретра → прочие (по методичке МИАЦ)"
    if _has_any(text, ["забрюшин"]):
        return "21", "medium", "keyword: забрюшинное → прочие"

    # --- костно-мышечная ---
    if _has_any(text, ["остеотоми"]):
        return "15.1", "high", "keyword: остеотомия"
    if _has_any(text, ["челюст", "скулов", "верхнечелюст", "нижнечелюст"]):
        return "15.2", "high", "keyword: ЧЛО"
    if _has_any(text, ["таз"]) and _has_any(text, ["перелом", "травм", "остеосинтез"]):
        return "15.3", "medium", "keyword: кости таза"
    if _has_any(text, ["внутрисуставн", "околосуставн"]):
        return "15.4", "medium", "keyword: около/внутрисуставной перелом"
    if _has_any(text, ["позвоночн", "дискэктоми", "ламинэктоми", "спондилодез"]):
        # дегенеративное vs травма — по умолчанию 15.5; нервная 2.7/2.8 если нейро
        if _has_any(text, ["спинн", "нервн", "декомпресс"]) and cls == "23":
            return "2.8", "medium", "keyword: позвоночник + нейро"
        return "15.5", "high", "keyword: позвоночник"

    # --- молочная / кожа / средостение / лимфа ---
    if _has_any(text, ["молочн", "маммопласт", "мастэктоми"]):
        return "16", "high", "keyword: молочная железа"
    if _has_any(text, ["вилочков", "тимус"]):
        return "18.1", "high", "keyword: тимус"
    if _has_any(text, ["средостен", "медиастинот"]):
        return "18", "high", "keyword: средостение"
    if _has_any(text, ["лимфатич", "лимфаденэктоми", "лимфатическ"]):
        return "20", "high", "keyword: лимфатическая система"
    if cls == "01" or _has_any(
        text,
        [
            "кож",
            "подкожн",
            "фурункул",
            "флегмон",
            "абсцесс кожи",
            "некрэктоми",
            "панариц",
            "ран",
            "пхо",
            "аутодермопласт",
        ],
    ):
        if cls in ("01",) or base_line == "17":
            return "17", "high" if cls == "01" else "medium", "class/keyword: кожа/ПЖК"

    # --- эндокринная ---
    if _has_any(text, ["тиреоид", "щитовидн", "струмэктоми", "тиреотом"]):
        return "3.1", "high", "keyword: щитовидная"

    # --- сосуды ---
    if cls == "12":
        if _has_any(text, ["вен", "варикоз", "флебэктоми", "варикозн"]):
            return "8.2", "high", "keyword: вены"
        return "8.1", "medium", "class: сосуды → артерии (по умолчанию)"

    # --- нервная ---
    if cls == "23":
        if _has_any(text, ["опух", "новообраз"]):
            return "2.5", "medium", "keyword: опухоль ЦНС"
        return "2", "medium", "class: ЦНС"
    if cls == "24":
        return "2.9", "high", "class: периферические нервы"

    # --- A16.30 прочее ---
    if cls == "30":
        if _has_any(text, ["грыж"]):
            return "9.3", "high", "A16.30 + грыжа"
        if _has_any(text, ["лапаротом", "релапаротом"]):
            return "9.5", "medium", "A16.30 + лапаротомия"
        if _has_any(text, ["лапароскоп"]):
            return "9", "medium", "A16.30 + лапароскопия → брюшная"
        return "21", "low", "A16.30 без явной анатомии → прочие (проверить)"

    # база по классу
    if base_line:
        conf = "high" if cls and cls in A16_CLASS_DEFAULT else "medium"
        return base_line, conf, f"class A16.{cls} → {base_line}"

    return "21", "low", "не удалось определить класс → прочие"


def map_code_to_form14(
    code: str,
    name: str = "",
    *,
    category: str = "",
    summary_key: Optional[str] = None,
    overrides: Optional[dict] = None,
    use_overrides: bool = True,
) -> Form14Mapping:
    """
    Сопоставление кода/категории со строкой ФСН 14.

    Калибровка ЛОР (MANUAL_LINE_*) применяется только для summary_key «lor»
    или если ключ не передан (обратная совместимость юнит-тестов ЛОР).
    Для остальных отделений — класс A16 + ключевые слова по их названиям.
    """
    code = str(code or "").strip()
    name = str(name or "").strip()
    category = str(category or "").strip()
    cls = parse_a16_class(code)
    # ЛОР-калибровка только для ЛОР; явный чужой ключ — без неё
    use_lor_manual = summary_key is None or str(summary_key).strip() in ("", "lor")

    # 0) ручные overrides хирурга (YAML / конструктор)
    if use_overrides:
        from analyzers.form14_overrides import lookup_override

        store = overrides
        if store is None:
            store = {}
        hit = lookup_override(store, code=code, category=category or name) if store else None
        if hit:
            line, entry = hit
            conf = "high"
            rule = "override хирурга"
            notes = str(entry.get("comment") or "")
            return Form14Mapping(
                code=code,
                name=name or category,
                a16_class=cls,
                line=line,
                line_name=FORM14_LINES.get(line, ""),
                confidence=conf,
                rule=rule,
                notes=notes,
            )

    # 1–2) калибровка формул сводной ЛОР — только для отделения ЛОР
    if use_lor_manual:
        manual = _manual_line_for_category(category) or _manual_line_for_category(name)
        if manual:
            line, confidence, rule = manual, "high", "калибровка ЛОР (формулы сводной)"
        elif code in MANUAL_LINE_BY_CODE:
            line, confidence, rule = (
                MANUAL_LINE_BY_CODE[code],
                "high",
                "калибровка ЛОР по коду (формулы сводной)",
            )
        else:
            base = A16_CLASS_DEFAULT.get(cls, "")
            line, confidence, rule = refine_line_by_keywords(code, name or category, base)
    else:
        base = A16_CLASS_DEFAULT.get(cls, "")
        line, confidence, rule = refine_line_by_keywords(code, name or category, base)

    line_name = FORM14_LINES.get(line, "")
    notes = ""
    if line == "21":
        notes = "Строка 21 по методичке МИАЦ: только мочевой пузырь/уретра, забрюшинное, костный мозг; иначе пересмотреть"
    if confidence == "low":
        notes = (notes + "; " if notes else "") + "требует ручной сверки"
    return Form14Mapping(
        code=code,
        name=name or category,
        a16_class=cls,
        line=line,
        line_name=line_name,
        confidence=confidence,
        rule=rule,
        notes=notes,
    )


def map_code_auto(
    code: str,
    name: str = "",
    *,
    category: str = "",
    summary_key: Optional[str] = None,
) -> Form14Mapping:
    """Авто без surgeon overrides (для колонки «Авто» в конструкторе)."""
    return map_code_to_form14(
        code,
        name,
        category=category,
        summary_key=summary_key,
        overrides=None,
        use_overrides=False,
    )


def map_categories(
    categories: List[dict],
    *,
    overrides: Optional[dict] = None,
    summary_key: Optional[str] = None,
) -> List[Form14Mapping]:
    """Категории surgery_categories → маппинги (по кодам)."""
    out: List[Form14Mapping] = []
    for cat in categories or []:
        codes = cat.get("codes") or []
        name = str(cat.get("category") or "")
        if not codes:
            m = map_code_to_form14(
                "", name, category=name, overrides=overrides, summary_key=summary_key
            )
            if m.rule != "override хирурга" and not (
                (summary_key is None or summary_key == "lor") and _manual_line_for_category(name)
            ):
                m.confidence = "low"
                m.rule = "нет кода — только по названию"
            if name:
                m.notes = ((m.notes + "; ") if m.notes else "") + f"категория: {name}"
            out.append(m)
            continue
        for code in codes:
            from analyzers.ksg_catalog import get_catalog

            ksg_name = get_catalog().name_for(str(code)) or name
            m = map_code_to_form14(
                str(code),
                ksg_name,
                category=name,
                overrides=overrides,
                summary_key=summary_key,
            )
            m.name = ksg_name or name
            if name:
                m.notes = ((m.notes + "; ") if m.notes else "") + f"категория: {name}"
            out.append(m)
    return out


def mapping_to_rows(mappings: List[Form14Mapping]) -> List[dict]:
    return [asdict(m) for m in mappings]


def summarize_mappings(mappings: List[Form14Mapping]) -> dict:
    by_conf: Dict[str, int] = {"high": 0, "medium": 0, "low": 0}
    by_line: Dict[str, int] = {}
    disputed = []
    for m in mappings:
        by_conf[m.confidence] = by_conf.get(m.confidence, 0) + 1
        by_line[m.line] = by_line.get(m.line, 0) + 1
        if m.confidence == "low" or m.line == "21":
            disputed.append(m)
    return {"by_confidence": by_conf, "by_line": by_line, "disputed": disputed, "total": len(mappings)}
