# analyzers/stationary.py
import pandas as pd
import numpy as np
from collections import Counter

class StationaryAnalyzer:
    """
    Анализатор стационарных данных по ЛОР-отделению (или любому другому).
    Проверяет качество документации и считает основные показатели.
    """
    def __init__(self, df, department, config):
        # Фильтруем строки, где Отделение содержит указанное название
        mask = df['Отделение'].str.contains(department, na=False)
        self.df = df[mask].copy()
        self.department = department
        self.config = config
        self.thresholds = config.get('thresholds', {})
        # Корректируем койко-дни: 0 -> 1
        bed_col = 'Всего дней проведено в стационаре (от поступления до исхода в днях)'
        if bed_col in self.df.columns:
            self.df[bed_col] = self.df[bed_col].apply(
                lambda x: 1 if pd.notna(x) and int(x) == 0 else int(x) if pd.notna(x) else 1
            )

    def calculate_metrics(self):
        """Считает общие показатели: всего пациентов, средний койко-день, план/экстр, возрастные группы."""
        total = len(self.df)
        bed_col = 'Всего дней проведено в стационаре (от поступления до исхода в днях)'
        avg_bed = self.df[bed_col].mean() if bed_col in self.df.columns else 0
        plan_mask = self.df['Тип госпитализации'] == 'плановая'
        emerg_mask = self.df['Тип госпитализации'] == 'экстренная'
        plan_cnt = plan_mask.sum()
        emerg_cnt = emerg_mask.sum()

        # Возрастные группы
        def age_group(age):
            if pd.isna(age):
                return 'неизвестно'
            age = int(age)
            if age <= 14:
                return '0-14 лет'
            elif age <= 17:
                return '15-17 лет'
            elif age <= 64:
                return '18-64 года'
            else:
                return '65+ лет'
        age_series = self.df['Возраст на момент госпитализации в стационар'].apply(age_group)
        age_dist = age_series.value_counts().reindex(['0-14 лет','15-17 лет','18-64 года','65+ лет'], fill_value=0)

        metrics = {
            'total_patients': total,
            'avg_bed_days': round(avg_bed, 2),
            'plan_count': plan_cnt,
            'emerg_count': emerg_cnt,
            'plan_percent': round(plan_cnt/total*100, 1) if total else 0,
            'emerg_percent': round(emerg_cnt/total*100, 1) if total else 0,
            'age_distribution': age_dist.to_dict()
        }
        return metrics

    def find_violations(self):
        """Основная функция: собирает все нарушения документации."""
        violations = []
        for idx, row in self.df.iterrows():
            kvs = row['Номер КВС']
            age = row['Возраст на момент госпитализации в стационар']
            doctor = row.get('Лечащий врач ', 'Не указан')  # обратите внимание на пробел после "врач"
            bed_days = row.get('Всего дней проведено в стационаре (от поступления до исхода в днях)', 0)
            # 3.1 Первичный осмотр
            prim_osmotr = str(row.get('Наличие заполненного первичного осмотра  в указанном движении', '')).strip().upper()
            if prim_osmotr != 'ДА':
                violations.append((kvs, age, doctor, 'Отсутствует первичный осмотр'))
            # 3.2 Эпикриз
            epicrisis = str(row.get('Наличие оформленного эпикриза в указанном движении', '')).strip().upper()
            if epicrisis != 'ДА':
                violations.append((kvs, age, doctor, 'Отсутствует эпикриз'))
            # 3.3 МКСБ
            mksb_status = str(row.get('Статус МКСБ', '')).strip()
            if mksb_status != 'Подписана':
                violations.append((kvs, age, doctor, 'МКСБ не подписана'))
            # 3.4 Лекарственные назначения
            drug_count = self._safe_int(row.get('Наличие оформленных лекарственных назначений в указанном движении', 0))
            if drug_count == 0:
                violations.append((kvs, age, doctor, 'Нет лекарственных назначений'))
            # 3.5 Дневники
            needed = self._safe_int(row.get('Количество дневниковых записей, которое необходимо было завести в указанном движении', 0))
            done = self._safe_int(row.get('Количество оформленных дневниковых записей  в указанном движении', 0))
            if done < needed:
                violations.append((kvs, age, doctor, f'Дневники: необх. {needed}, оформ. {done}'))
            # 3.6 ИДС
            related = str(row.get('Другие связанные документы', ''))
            if '83 - Информированное добровольное согласие' not in related:
                violations.append((kvs, age, doctor, 'Отсутствует ИДС'))
            # 3.7 Превышение койко-дня
            if bed_days > self.thresholds.get('max_bed_days', 7):
                violations.append((kvs, age, doctor, f'Превышение койко-дня: {bed_days} дн.'))
            # 3.8 Протоколы операций
            ops_count = self._safe_int(row.get('Хир. активность (количество)', 0))
            prot_count = self._safe_int(row.get('Хир. активность (протоколы)', 0))
            if ops_count > 0 and prot_count < ops_count:
                violations.append((kvs, age, doctor, f'Протоколы: {prot_count}, операций: {ops_count}'))
            elif prot_count > ops_count:
                violations.append((kvs, age, doctor, f'Избыток протоколов: прот. {prot_count}, опер. {ops_count}'))
        return violations

    def top_doctors_by_violations(self, violations):
        """Возвращает список врачей, отсортированный по числу нарушений."""
        docs = [v[2] for v in violations]  # индекс врача в кортеже
        counter = Counter(docs)
        return counter.most_common()

    def top_doctors_ids(self, violations):
        """Топ врачей по отсутствию ИДС."""
        ids_docs = [v[2] for v in violations if 'ИДС' in v[3]]
        counter = Counter(ids_docs)
        return counter.most_common()

    def analytical_note(self, violations):
        """Краткая аналитическая записка."""
        total = len(violations)
        if total == 0:
            return "Нарушений не выявлено. Документация ведётся отлично."
        top_docs = self.top_doctors_by_violations(violations)
        note = f"Всего выявлено нарушений: {total}.\n"
        note += "Топ врачей по общему числу нарушений:\n"
        for doc, cnt in top_docs[:5]:
            note += f"  - {doc}: {cnt}\n"
        ids_top = self.top_doctors_ids(violations)
        if ids_top:
            note += "Топ врачей по отсутствию ИДС:\n"
            for doc, cnt in ids_top[:3]:
                note += f"  - {doc}: {cnt}\n"
        return note

    @staticmethod
    def _safe_int(val):
        try:
            return int(float(val))
        except:
            return 0