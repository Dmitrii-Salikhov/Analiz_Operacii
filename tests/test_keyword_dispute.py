"""Ничьи и переклассификация по name_keywords."""
from analyzers.surgery import classify_by_name, reclassify_ops_by_keywords
import pandas as pd


CATS = [
    {
        "category": "Мирингопластика",
        "name_keywords": ["мирингопластика"],
        "group": "ухо",
        "line": "6",
        "histology": False,
    },
    {
        "category": "Тимпанопластика",
        "name_keywords": ["тимпанопластика"],
        "group": "ухо",
        "line": "6",
        "histology": False,
    },
    {
        "category": "Пластика уха A",
        "name_keywords": ["пластика", "уха"],
        "group": "ухо",
        "line": "6",
        "histology": False,
    },
    {
        "category": "Пластика уха B",
        "name_keywords": ["пластика", "раковин"],
        "group": "ухо",
        "line": "6",
        "histology": False,
    },
]


def test_single_keyword_winner_no_dispute():
    cats = [
        {
            "category": "Мирингопластика",
            "name_keywords": ["мирингопластика"],
            "group": "ухо",
            "line": "6",
            "histology": False,
        },
        {
            "category": "Тимпанопластика",
            "name_keywords": ["тимпанопластика"],
            "group": "ухо",
            "line": "6",
            "histology": False,
        },
    ]
    info, disputed, cands = classify_by_name(cats, "Операция: мирингопластика слева")
    assert info is not None
    assert info[0] == "Мирингопластика"
    assert disputed is False
    assert cands == "Мирингопластика"


def test_tie_marks_dispute():
    # оба набирают score=1 по «пластика»
    info, disputed, cands = classify_by_name(CATS, "Пластика мягких тканей")
    assert info is not None
    assert disputed is True
    assert "Пластика уха A" in cands
    assert "Пластика уха B" in cands
    assert info[0] == "Пластика уха A"  # первая из лидеров


def test_higher_score_wins_no_dispute():
    cats = [
        {
            "category": "Пластика общая",
            "name_keywords": ["пластика"],
            "group": "ухо",
            "line": "6",
            "histology": False,
        },
        {
            "category": "Пластика раковин",
            "name_keywords": ["пластика", "раковин"],
            "group": "ухо",
            "line": "6",
            "histology": False,
        },
    ]
    info, disputed, cands = classify_by_name(cats, "Пластика раковин носа")
    assert info is not None
    assert info[0] == "Пластика раковин"
    assert disputed is False
    assert cands == "Пластика раковин"


def test_reclassify_skips_manual():
    df = pd.DataFrame(
        [
            {
                "Код": "",
                "Услуга": "Пластика мягких тканей",
                "Категория": "Старая",
                "Группа": "x",
                "Строка_4001": "",
                "Гистология": False,
                "Тип_ЭМК": "",
                "Спор_ключей": False,
                "Спорные_категории": "",
                "Ручная_категория": True,
            }
        ]
    )
    out = reclassify_ops_by_keywords(df, CATS)
    assert out.iloc[0]["Категория"] == "Старая"
    assert bool(out.iloc[0]["Ручная_категория"]) is True


def test_reclassify_updates_dispute():
    df = pd.DataFrame(
        [
            {
                "Код": "",
                "Услуга": "Пластика мягких тканей",
                "Категория": "Не классифицировано",
                "Группа": "прочее",
                "Строка_4001": "",
                "Гистология": False,
                "Тип_ЭМК": "",
                "Спор_ключей": False,
                "Спорные_категории": "",
                "Ручная_категория": False,
            }
        ]
    )
    out = reclassify_ops_by_keywords(df, CATS)
    assert bool(out.iloc[0]["Спор_ключей"]) is True
    assert "Пластика уха A" in str(out.iloc[0]["Спорные_категории"])
