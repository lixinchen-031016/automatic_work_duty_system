"""国家法定节假日与调休数据导入。

内置数据来源于 ``chinese-days``（MIT）发布的年度 JSON；缺失年份可在界面中
主动在线更新。国家日历只说明“哪些日期放假/上班”，本模块会根据学期起始日
推断补课日代表的逻辑周与星期，并将结果交给用户预览确认。
"""

from __future__ import annotations

import json
import ssl
from dataclasses import dataclass
from datetime import date, timedelta
from itertools import combinations, permutations
from urllib.request import Request, urlopen

from .calendar import MAX_WEEK, CalendarEntry

HOLIDAY_DATA_URL = "https://unpkg.com/chinese-days@latest/dist/years/{year}.json"

# 内置最近两个学年会覆盖到的官方数据；其他年份可按需在线更新。
BUILTIN_HOLIDAY_DATA: dict[int, dict[str, dict[str, str]]] = {
    2025: {
        "holidays": {
            "2025-01-01": "New Year's Day,元旦,1",
            "2025-01-28": "Spring Festival,春节,4",
            "2025-01-29": "Spring Festival,春节,4",
            "2025-01-30": "Spring Festival,春节,4",
            "2025-01-31": "Spring Festival,春节,4",
            "2025-02-01": "Spring Festival,春节,4",
            "2025-02-02": "Spring Festival,春节,4",
            "2025-02-03": "Spring Festival,春节,4",
            "2025-02-04": "Spring Festival,春节,4",
            "2025-04-04": "Tomb-sweeping Day,清明,1",
            "2025-04-05": "Tomb-sweeping Day,清明,1",
            "2025-04-06": "Tomb-sweeping Day,清明,1",
            "2025-05-01": "Labour Day,劳动节,2",
            "2025-05-02": "Labour Day,劳动节,2",
            "2025-05-03": "Labour Day,劳动节,2",
            "2025-05-04": "Labour Day,劳动节,2",
            "2025-05-05": "Labour Day,劳动节,2",
            "2025-05-31": "Dragon Boat Festival,端午,1",
            "2025-06-01": "Dragon Boat Festival,端午,1",
            "2025-06-02": "Dragon Boat Festival,端午,1",
            "2025-10-01": "National Day,国庆节,3",
            "2025-10-02": "National Day,国庆节,3",
            "2025-10-03": "National Day,国庆节,3",
            "2025-10-04": "National Day,国庆节,3",
            "2025-10-05": "National Day,国庆节,3",
            "2025-10-06": "Mid-autumn Festival,中秋,1",
            "2025-10-07": "National Day,国庆节,3",
            "2025-10-08": "National Day,国庆节,3",
        },
        "workdays": {
            "2025-01-26": "Spring Festival,春节,4",
            "2025-02-08": "Spring Festival,春节,4",
            "2025-04-27": "Labour Day,劳动节,2",
            "2025-09-28": "National Day,国庆节,3",
            "2025-10-11": "National Day,国庆节,3",
        },
        "inLieuDays": {
            "2025-02-03": "Spring Festival,春节,4",
            "2025-02-04": "Spring Festival,春节,4",
            "2025-05-05": "Labour Day,劳动节,2",
            "2025-10-07": "National Day,国庆节,3",
            "2025-10-08": "National Day,国庆节,3",
        },
    },
    2026: {
        "holidays": {
            "2026-01-01": "New Year's Day,元旦,1",
            "2026-01-02": "New Year's Day,元旦,1",
            "2026-01-03": "New Year's Day,元旦,1",
            "2026-02-15": "Spring Festival,春节,4",
            "2026-02-16": "Spring Festival,春节,4",
            "2026-02-17": "Spring Festival,春节,4",
            "2026-02-18": "Spring Festival,春节,4",
            "2026-02-19": "Spring Festival,春节,4",
            "2026-02-20": "Spring Festival,春节,4",
            "2026-02-21": "Spring Festival,春节,4",
            "2026-02-22": "Spring Festival,春节,4",
            "2026-02-23": "Spring Festival,春节,4",
            "2026-04-04": "Tomb-sweeping Day,清明,1",
            "2026-04-05": "Tomb-sweeping Day,清明,1",
            "2026-04-06": "Tomb-sweeping Day,清明,1",
            "2026-05-01": "Labour Day,劳动节,2",
            "2026-05-02": "Labour Day,劳动节,2",
            "2026-05-03": "Labour Day,劳动节,2",
            "2026-05-04": "Labour Day,劳动节,2",
            "2026-05-05": "Labour Day,劳动节,2",
            "2026-06-19": "Dragon Boat Festival,端午,1",
            "2026-06-20": "Dragon Boat Festival,端午,1",
            "2026-06-21": "Dragon Boat Festival,端午,1",
            "2026-09-25": "Mid-autumn Festival,中秋,1",
            "2026-09-26": "Mid-autumn Festival,中秋,1",
            "2026-09-27": "Mid-autumn Festival,中秋,1",
            "2026-10-01": "National Day,国庆节,3",
            "2026-10-02": "National Day,国庆节,3",
            "2026-10-03": "National Day,国庆节,3",
            "2026-10-04": "National Day,国庆节,3",
            "2026-10-05": "National Day,国庆节,3",
            "2026-10-06": "National Day,国庆节,3",
            "2026-10-07": "National Day,国庆节,3",
        },
        "workdays": {
            "2026-01-04": "New Year's Day,元旦,1",
            "2026-02-14": "Spring Festival,春节,4",
            "2026-02-28": "Spring Festival,春节,4",
            "2026-05-09": "Labour Day,劳动节,2",
            "2026-09-20": "National Day,国庆节,3",
            "2026-10-10": "National Day,国庆节,3",
        },
        "inLieuDays": {
            "2026-01-02": "New Year's Day,元旦,1",
            "2026-02-20": "Spring Festival,春节,4",
            "2026-02-23": "Spring Festival,春节,4",
            "2026-05-05": "Labour Day,劳动节,2",
            "2026-10-06": "National Day,国庆节,3",
            "2026-10-07": "National Day,国庆节,3",
        },
    },
}


@dataclass(frozen=True)
class NationalHolidayData:
    year: int
    holidays: dict[date, str]
    workdays: dict[date, str]
    in_lieu_days: dict[date, str]
    source: str


@dataclass
class HolidayImportPlan:
    entries: list[CalendarEntry]
    warnings: list[str]
    source: str


def _parse_date_map(values: dict[str, str]) -> dict[date, str]:
    return {date.fromisoformat(key): str(value) for key, value in values.items()}


def parse_holiday_data(
    year: int,
    payload: dict,
    source: str = "在线数据",
) -> NationalHolidayData:
    """解析 chinese-days 年度 JSON。"""
    return NationalHolidayData(
        year=year,
        holidays=_parse_date_map(payload.get("holidays", {})),
        workdays=_parse_date_map(payload.get("workdays", {})),
        in_lieu_days=_parse_date_map(payload.get("inLieuDays", {})),
        source=source,
    )


def get_builtin_holiday_data(year: int) -> NationalHolidayData | None:
    payload = BUILTIN_HOLIDAY_DATA.get(year)
    if payload is None:
        return None
    return parse_holiday_data(year, payload, source="内置数据")


def fetch_holiday_data(year: int, timeout: int = 10) -> NationalHolidayData:
    """在线获取年度节假日数据。"""
    request = Request(
        HOLIDAY_DATA_URL.format(year=year),
        headers={"User-Agent": "DutySystem/1.0"},
    )
    try:
        context = None
        try:
            import certifi
            context = ssl.create_default_context(cafile=certifi.where())
        except ImportError:
            pass
        with urlopen(request, timeout=timeout, context=context) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception as exc:
        raise ValueError(f"无法获取 {year} 年国家调休数据：{exc}") from exc
    return parse_holiday_data(year, payload, source=HOLIDAY_DATA_URL.format(year=year))


def _holiday_name(value: str) -> str:
    parts = [part.strip() for part in value.split(",") if part.strip()]
    return parts[1] if len(parts) > 1 else (parts[0] if parts else "法定假日")


def _minimum_distance_pairs(
    workdays: list[date],
    candidates: list[date],
) -> dict[date, date]:
    """在两组日期间寻找总距离最小的对应关系。"""
    if not workdays or not candidates:
        return {}
    pair_count = min(len(workdays), len(candidates))
    best: tuple[int, list[tuple[date, date]]] | None = None
    if pair_count <= 6:
        for work_subset in combinations(workdays, pair_count):
            for candidate_subset in combinations(candidates, pair_count):
                for ordered in permutations(candidate_subset):
                    cost = sum(
                        abs((workday - candidate).days)
                        for workday, candidate in zip(work_subset, ordered)
                    )
                    if best is None or cost < best[0]:
                        best = (cost, list(zip(work_subset, ordered)))
    else:  # 正常情况下国家调休每组不超过 3 个，保留兜底避免组合爆炸
        remaining = list(candidates)
        pairs = []
        for workday in workdays[:pair_count]:
            candidate = min(remaining, key=lambda item: abs((workday - item).days))
            remaining.remove(candidate)
            pairs.append((workday, candidate))
        best = (0, pairs)
    return dict(best[1]) if best is not None else {}


def infer_workday_mappings(
    term_start: date,
    data: NationalHolidayData,
) -> tuple[dict[date, tuple[int, int]], list[str]]:
    """按同一节日分组，在调休上班日和对应的放假日之间推断 maps_to。"""
    term_end = term_start + timedelta(days=MAX_WEEK * 7 - 1)
    workdays = sorted(
        workday for workday in data.workdays
        if term_start <= workday <= term_end
    )
    holiday_candidates_by_name: dict[str, list[date]] = {}
    for holiday in data.holidays:
        if term_start <= holiday <= term_end:
            name = _holiday_name(data.holidays[holiday])
            holiday_candidates_by_name.setdefault(name, []).append(holiday)
    in_lieu_by_name: dict[str, list[date]] = {}
    for holiday in data.in_lieu_days:
        if term_start <= holiday <= term_end:
            name = _holiday_name(data.in_lieu_days[holiday])
            in_lieu_by_name.setdefault(name, []).append(holiday)
    for values in [*holiday_candidates_by_name.values(), *in_lieu_by_name.values()]:
        values[:] = sorted(set(values))

    mappings: dict[date, tuple[int, int]] = {}
    unmatched: list[str] = []
    workdays_by_name: dict[str, list[date]] = {}
    for workday in workdays:
        name = _holiday_name(data.workdays[workday])
        workdays_by_name.setdefault(name, []).append(workday)

    for name, named_workdays in workdays_by_name.items():
        candidates = (in_lieu_by_name.get(name)
                      or holiday_candidates_by_name.get(name, []))
        pairs = _minimum_distance_pairs(named_workdays, candidates)
        for workday, holiday in pairs.items():
            offset = (holiday - term_start).days
            if 0 <= offset < MAX_WEEK * 7:
                mappings[workday] = (offset // 7 + 1, holiday.isoweekday())
        for workday in named_workdays:
            if workday not in pairs:
                unmatched.append(
                    f"{workday.isoformat()} {name} 调休补课未能自动推断代表日，请手动确认")
    return mappings, unmatched


def build_holiday_import_plan(
    term_start: date,
    data: NationalHolidayData,
) -> HolidayImportPlan:
    """把国家年度数据转换为可预览编辑的学期日历覆盖项。"""
    term_end = term_start + timedelta(days=MAX_WEEK * 7 - 1)
    entries: list[CalendarEntry] = []
    for holiday, info in sorted(data.holidays.items()):
        if term_start <= holiday <= term_end:
            entries.append(CalendarEntry(
                holiday, "off", note=f"{_holiday_name(info)}放假"))

    mappings, warnings = infer_workday_mappings(term_start, data)
    for workday, info in sorted(data.workdays.items()):
        if not term_start <= workday <= term_end:
            continue
        mapping = mappings.get(workday)
        if mapping is None:
            continue
        week, weekday = mapping
        entries.append(CalendarEntry(
            workday, "class", week, weekday,
            note=f"{_holiday_name(info)}调休补课（自动推断，请确认）",
        ))
    return HolidayImportPlan(
        entries=sorted(entries, key=lambda entry: entry.date),
        warnings=warnings,
        source=data.source,
    )
