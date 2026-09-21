"""学期日历覆盖：逻辑教学周与真实日期之间的映射。

排班引擎始终使用逻辑 (week, weekday, block) 作为键。本模块只负责把逻辑
教学日解析为真实日期，支持两类覆盖：

* ``off``：逻辑工作日对应的自然日期放假。
* ``class``：周末补课日，``maps_to`` 指向它代表的逻辑工作日。

调休成对出现时，被放假日期到补课日期形成双向映射；纯法定假日只有 ``off``，
该逻辑日没有真实日期，展示层应整体跳过。
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Literal

from .parser import WEEKDAY_LABELS

OverrideType = Literal["off", "class"]
MAX_WEEK = 25


@dataclass(frozen=True)
class CalendarEntry:
    """一条学期日历覆盖项。"""

    date: date
    override_type: OverrideType
    maps_to_week: int | None = None
    maps_to_weekday: int | None = None
    note: str = ""

    def __post_init__(self) -> None:
        if isinstance(self.date, str):
            object.__setattr__(self, "date", date.fromisoformat(self.date))
        override_type = str(self.override_type).strip().lower()
        object.__setattr__(self, "override_type", override_type)
        object.__setattr__(self, "note", str(self.note or "").strip())
        if override_type not in ("off", "class"):
            raise ValueError("日历类型只能是 off 或 class")
        if override_type == "off":
            if self.maps_to_week is not None or self.maps_to_weekday is not None:
                raise ValueError("off 项不应设置代表周/星期")
        else:
            if self.maps_to_week is None or self.maps_to_weekday is None:
                raise ValueError("class 项必须设置代表周和星期")
            if not 1 <= int(self.maps_to_week) <= MAX_WEEK:
                raise ValueError("代表周应为 1–25")
            if not 1 <= int(self.maps_to_weekday) <= 7:
                raise ValueError("代表星期应为 1–7")
            object.__setattr__(self, "maps_to_week", int(self.maps_to_week))
            object.__setattr__(self, "maps_to_weekday", int(self.maps_to_weekday))


def natural_date(term_start: date, week: int, weekday: int) -> date:
    """纯净校历下的逻辑教学日日期。"""
    return term_start + timedelta(days=(week - 1) * 7 + (weekday - 1))


def validate_calendar_entries(
    entries: Iterable[CalendarEntry],
    term_start: date | None = None,
    *,
    allow_unpaired_off: bool = False,
) -> None:
    """校验覆盖项成对关系；非法时抛 ``ValueError``。

    ``allow_unpaired_off`` 用于纯法定假日：此类 off 没有补课项，也不产生真实
    排班日期。默认严格模式要求每个 off 都存在唯一 class 回指。
    """
    rows = list(entries)
    off_by_date: dict[date, CalendarEntry] = {}
    classes_by_target: dict[tuple[int, int], list[CalendarEntry]] = defaultdict(list)
    seen_dates: set[date] = set()

    for entry in rows:
        if not isinstance(entry, CalendarEntry):
            raise TypeError("日历条目必须是 CalendarEntry")
        if entry.date in seen_dates:
            raise ValueError(f"日历日期重复：{entry.date.isoformat()}")
        seen_dates.add(entry.date)
        if entry.override_type == "off":
            off_by_date[entry.date] = entry
        else:
            classes_by_target[(entry.maps_to_week, entry.maps_to_weekday)].append(entry)

    for target, classes in classes_by_target.items():
        if len(classes) > 1:
            dates = "、".join(x.date.isoformat() for x in sorted(classes, key=lambda x: x.date))
            raise ValueError(
                f"代表第{target[0]}周{WEEKDAY_LABELS[target[1]]}的补课日不唯一：{dates}")

    for entry in rows:
        if entry.override_type != "class":
            continue
        if entry.date.isoweekday() not in (6, 7):
            raise ValueError(f"补课日 {entry.date.isoformat()} 应为周六或周日")
        if term_start is not None:
            class_offset = (entry.date - term_start).days
            if class_offset < 0 or class_offset >= MAX_WEEK * 7:
                raise ValueError(
                    f"补课日 {entry.date.isoformat()} 超出学期 1–{MAX_WEEK} 周范围")
            assert entry.maps_to_week is not None and entry.maps_to_weekday is not None
            target_date = natural_date(
                term_start, entry.maps_to_week, entry.maps_to_weekday)
            target_off = off_by_date.get(target_date)
            if target_off is None:
                raise ValueError(
                    f"补课日 {entry.date.isoformat()} 指向的 "
                    f"第{entry.maps_to_week}周{WEEKDAY_LABELS[entry.maps_to_weekday]} "
                    f"({target_date.isoformat()}) 未设为 off")

    if not allow_unpaired_off:
        if term_start is None:
            # 没有学期起始日时无法反推 week/weekday，只能做结构上的 1:1 校验。
            if len(off_by_date) != len(classes_by_target):
                raise ValueError("调休 off/class 项数量不一致")
        else:
            for entry in rows:
                if entry.override_type != "off":
                    continue
                offset = (entry.date - term_start).days
                target = (offset // 7 + 1, offset % 7 + 1)
                if not classes_by_target.get(target):
                    raise ValueError(
                        f"off 项 {entry.date.isoformat()} 缺少对应的 class 补课项")


class TermCalendar:
    """某个学期起始日下的日历覆盖解析器。"""

    def __init__(self, term_start: date, entries: Iterable[CalendarEntry]):
        rows = list(entries)
        validate_calendar_entries(rows, term_start, allow_unpaired_off=True)
        self.term_start = term_start
        self.entries = tuple(sorted(rows, key=lambda x: x.date))
        self._by_date = {entry.date: entry for entry in self.entries}
        self._classes_by_target: dict[tuple[int, int], list[CalendarEntry]] = defaultdict(list)
        for entry in self.entries:
            if entry.override_type == "class":
                self._classes_by_target[
                    (entry.maps_to_week, entry.maps_to_weekday)
                ].append(entry)

    def logical_to_date(self, week: int, weekday: int) -> date | None:
        """逻辑教学日 -> 真实日期；off 且无补课映射时返回 ``None``。"""
        nat = natural_date(self.term_start, week, weekday)
        override = self._by_date.get(nat)
        if override is None or override.override_type != "off":
            return nat
        matches = self._classes_by_target.get((week, weekday), ())
        return matches[0].date if len(matches) == 1 else None

    def date_to_logical(self, target: date) -> tuple[int, int] | None:
        """真实日期 -> 逻辑教学日；off 日期返回 ``None``。"""
        offset = (target - self.term_start).days
        if offset < 0 or offset >= MAX_WEEK * 7:
            raise ValueError(
                f"日期 {target.isoformat()} 超出学期 1–{MAX_WEEK} 周范围")
        week, weekday = offset // 7 + 1, offset % 7 + 1
        override = self._by_date.get(target)
        if override is not None:
            if override.override_type == "off":
                return None
            return override.maps_to_week, override.maps_to_weekday
        return week, weekday

    def is_off(self, week: int, weekday: int) -> bool:
        """逻辑日是否被纯放假（没有对应真实日期）。"""
        return self.logical_to_date(week, weekday) is None

    def off_natural_date(self, week: int, weekday: int) -> date | None:
        """返回逻辑工作日的自然日期；若该日设为 off，则返回该日期。"""
        nat = natural_date(self.term_start, week, weekday)
        override = self._by_date.get(nat)
        if override is not None and override.override_type == "off":
            return nat
        return None
