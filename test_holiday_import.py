"""国家调休年历导入：内置数据、映射推断与在线数据解析。"""

from __future__ import annotations

import json
from datetime import date

from duty_system.calendar import validate_calendar_entries
from duty_system.holiday import (
    build_holiday_import_plan,
    fetch_holiday_data,
    get_builtin_holiday_data,
)


def test_builtin_2026_import_plan_infers_makeup_mappings() -> None:
    term_start = date(2026, 9, 14)
    data = get_builtin_holiday_data(2026)
    assert data is not None

    plan = build_holiday_import_plan(term_start, data)
    entries = {(entry.date, entry.override_type): entry for entry in plan.entries}
    assert entries[(date(2026, 10, 1), "off")].note == "国庆节放假"
    assert entries[(date(2026, 9, 20), "class")].maps_to_week == 4
    assert entries[(date(2026, 9, 20), "class")].maps_to_weekday == 2
    assert entries[(date(2026, 10, 10), "class")].maps_to_week == 4
    assert entries[(date(2026, 10, 10), "class")].maps_to_weekday == 3
    validate_calendar_entries(plan.entries, term_start, allow_unpaired_off=True)


def test_fetch_holiday_data_parses_online_payload(monkeypatch) -> None:
    payload = {
        "holidays": {"2030-01-01": "New Year's Day,元旦,1"},
        "workdays": {"2030-01-05": "New Year's Day,元旦,1"},
        "inLieuDays": {"2030-01-01": "New Year's Day,元旦,1"},
    }

    class Response:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def read(self) -> bytes:
            return json.dumps(payload).encode()

    monkeypatch.setattr("duty_system.holiday.urlopen", lambda *_a, **_k: Response())
    data = fetch_holiday_data(2030)
    assert data.holidays == {date(2030, 1, 1): "New Year's Day,元旦,1"}
    assert data.workdays == {date(2030, 1, 5): "New Year's Day,元旦,1"}

