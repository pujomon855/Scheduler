# -*- coding: utf-8 -*-

from dataclasses import dataclass
from enum import IntEnum


class EMonitorsColIdx(IntEnum):
    MONITOR_NO = 0
    MONITOR_NAME = 1
    FIX_SPECIALIST = 2
    MONITOR_NOTE = 3

    COMBO_NO = 5
    COMBO_MEMBER1 = 6
    COMBO_MEMBER2 = 7
    COMBO_NOTE = 8


DATA_START_ROW_IDX = 8


@dataclass(frozen=True)
class Monitor:
    name: str
    is_fix_specialist: bool


def load_monitors_info(wb):
    ws = wb['monitors']
    monitor_dict = {}
    tmp_must_work_at_office_groups = []

    for row in ws.iter_rows(min_row=DATA_START_ROW_IDX):
        # Monitors
        if name := row[EMonitorsColIdx.MONITOR_NAME].value:
            is_fix_specialist = (row[EMonitorsColIdx.FIX_SPECIALIST].value == 1)
            monitor_dict[name] = Monitor(name, is_fix_specialist)

        # Groups
        if name1 := row[EMonitorsColIdx.COMBO_MEMBER1].value:
            name2 = row[EMonitorsColIdx.COMBO_MEMBER2].value
            tmp_must_work_at_office_groups.append((name1, name2))

    # convert group's name str to Monitor
    must_work_at_office_groups = []
    for name1, name2 in tmp_must_work_at_office_groups:
        m1 = monitor_dict.get(name1)
        m2 = monitor_dict.get(name2)
        if m1 and m2:
            must_work_at_office_groups.append((m1, m2))

    return monitor_dict, must_work_at_office_groups
