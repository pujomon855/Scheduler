# -*- coding: utf-8 -*-

import copy
from dataclasses import dataclass
from datetime import datetime, timedelta
import openpyxl
import random
import sys
import time

from combo import ERole, MonitorSchedule, assign_role_maxes, gen_monitor_combos
from combo import MONITOR_ROLES_ALL, MONITOR_ROLES_AM
import monitors

HEADER_ROW_IDX = 7
DATA_START_ROW_IDX = HEADER_ROW_IDX + 1


@dataclass
class DayPriority:
    """
    日付と優先度を持つクラス
    priorityが小さいほど優先度が高い。
    """
    day: datetime
    priority: int = 0


def make_schedule(excel_path):
    keep_vba = True if excel_path.endswith('xlsm') else False
    wb = openpyxl.load_workbook(excel_path, keep_vba=keep_vba)
    monitor_dict, must_work_at_office_groups = monitors.load_monitors_info(wb)
    ws = wb['latest']
    monitor_schedule_dict, weekday_dict = load_initial_schedules(ws, monitor_dict)
    sorted_day_priorities = sorted(weekday_dict.values(), key=lambda dp: dp.priority)
    weekdays = [dp.day for dp in sorted_day_priorities]
    assign_role_maxes(monitor_schedule_dict, MONITOR_ROLES_ALL, len(weekday_dict))
    cp_monitor_schedule_dict = copy_monitor_schedule_dict(monitor_schedule_dict)
    res_monitor_schedule_dict = _try_assign_monitors(
        monitor_schedule_dict, cp_monitor_schedule_dict, weekdays, 1000)
    if res_monitor_schedule_dict:
        monitor_schedule_dict = res_monitor_schedule_dict
    else:
        cp_monitor_schedule_dict = _try_assign_monitors(
            monitor_schedule_dict, cp_monitor_schedule_dict, weekdays, 1000, True)
        if cp_monitor_schedule_dict:
            monitor_schedule_dict = cp_monitor_schedule_dict
        else:
            assign_monitors(monitor_schedule_dict, weekdays, True, True)

    debug_schedules(monitor_schedule_dict, weekdays)
    output_schedules(ws, monitor_schedule_dict, weekday_dict)
    wb.save(excel_path)


def _try_assign_monitors(msd, cp_msd, weekdays, try_cnt, ignore_prev_day=False):
    for i in range(try_cnt):
        if assign_monitors(cp_msd, weekdays, ignore_prev_day):
            print(f'{ignore_prev_day=}: {i + 1}: found.')
            return cp_msd
        cp_msd = copy_monitor_schedule_dict(msd)
    print(f'{ignore_prev_day=}: {try_cnt}: not found.')
    return None


def debug_schedules(monitor_schedule_dict, weekdays):
    for day in weekdays:
        md = {}
        for ms in monitor_schedule_dict.values():
            if (role := ms.schedule.get(day)) in MONITOR_ROLES_ALL:
                md[role] = ms.monitor
        if len(md) < 3:
            print(day)
        else:
            print(f'{day}: {md[ERole.AM1].name}, {md[ERole.AM2].name}, {md[ERole.PM].name}')
    print()
    for m, ms in monitor_schedule_dict.items():
        print(f'{m.name},{ms.am1_count},{ms.am2_count},{ms.pm_count},{ms.monitor_count}')


def load_initial_schedules(ws, monitor_dict):
    """
    指定シートからあらかじめ代入されている予定を読み取り、各監視者のスケジュールを初期化する。

    :param ws: ワークシート
    :param monitor_dict: 監視者の辞書(key:=name, item:=Monitor)
    :return: 監視者のスケジュールの辞書(key:=Monitor, item:=MonitorSchedule),
                日付の辞書(key:=row_idx(int), item:=DayPriority)
    """
    monitor_schedule_dict = init_monitor_schedules(ws, monitor_dict)
    num_of_monitors = len(monitor_dict)
    monitor_list = monitor_dict.values()
    weekday_dict = {}
    for row_idx, row in enumerate(ws.iter_rows(min_row=DATA_START_ROW_IDX, max_col=num_of_monitors + 2),
                                  DATA_START_ROW_IDX):
        day = row[0].value
        if not day:
            break
        holiday_cell = row[num_of_monitors + 1]
        if not is_weekday(day, holiday_cell):
            continue
        day_priority = DayPriority(day)
        weekday_dict[row_idx] = day_priority
        for idx, monitor in enumerate(monitor_list, 1):
            if val := row[idx].value:
                schedule = monitor_schedule_dict[monitor]
                role = convert_val_to_role(val)
                schedule.schedule[day] = role
                if role in MONITOR_ROLES_ALL:
                    day_priority.priority -= 2
                else:
                    day_priority.priority -= 1
        if day_priority.priority >= 0:
            day_priority.priority = day.day
    return monitor_schedule_dict, weekday_dict


def init_monitor_schedules(ws, monitor_dict):
    monitor_schedule_dict = {}
    for row in ws.iter_rows(min_row=HEADER_ROW_IDX, max_row=HEADER_ROW_IDX,
                            min_col=2, max_col=len(monitor_dict) + 1):
        for cell in row:
            name = cell.value
            if name not in monitor_dict:
                raise ValueError(f'{name} is not in monitors.')
            monitor = monitor_dict[name]
            monitor_schedule_dict[monitor] = MonitorSchedule(monitor, cell.column)
    return monitor_schedule_dict


def is_weekday(date, holiday_cell):
    if holiday_cell.value:
        return False
    return date.weekday() < 5  # 5 and 6 mean Saturday and Sunday respectively.


def convert_val_to_role(val):
    for role in ERole:
        if role.name == val:
            return role
    return ERole.OTHER


def copy_monitor_schedule_dict(monitor_schedule_dict):
    cp = {}
    for m, ms in monitor_schedule_dict.items():
        cp[m] = copy.copy(ms)
    return cp


def assign_monitors(monitor_schedule_dict, weekdays, ignore_prev_day=False, force_exec=False):
    """
    監視当番の割り振りを行う。

    :param monitor_schedule_dict: 監視スケジュールのdict(key:=Monitor, item:=MonitorSchedule)
    :param weekdays: 営業日のIterable
    :param ignore_prev_day: 前日からの連続条件を除外する場合はTrue
    :param force_exec: 均等な割り振りが不可の場合でも、その日を除いて処理を続行する場合はTrueを設定する
    :return: 割り振りが完了した場合はTrue
    """
    monitor_set = monitor_schedule_dict.keys()
    all_monitor_combos = set(gen_monitor_combos(monitor_set))
    for day in weekdays:
        filters = create_filters(day, monitor_schedule_dict)
        # extract monitor combo that meets all filters.
        monitor_combos = [mc for mc in all_monitor_combos if all([f(mc) for f in filters])]
        if not monitor_combos:
            if not ignore_prev_day:
                return False
            filters = create_filters(day, monitor_schedule_dict, False)
            monitor_combos = [mc for mc in all_monitor_combos if all([f(mc) for f in filters])]
            if not monitor_combos:
                if not force_exec:
                    return False
                continue

        # Choice a monitor combo at random.
        monitor_combo = random.choice(monitor_combos)
        monitor_schedule_dict[monitor_combo.monitor_am1].schedule[day] = ERole.AM1
        monitor_schedule_dict[monitor_combo.monitor_am2].schedule[day] = ERole.AM2
        monitor_schedule_dict[monitor_combo.monitor_pm].schedule[day] = ERole.PM
    return True


def create_filters(day, monitor_schedule_dict, include_pre_day=True):
    filters = []
    for ms in monitor_schedule_dict.values():
        # Manually input day role
        if role := ms.schedule.get(day):
            if role in MONITOR_ROLES_ALL:
                filters.append(create_filter(ms.monitor, roles=[role]))
            elif role == ERole.OTHER:
                filters.append(create_filter(ms.monitor, include=False))
        else:
            # Adjust monitoring days
            filter_roles = []
            if ms.is_role_max(ERole.AM1):
                filter_roles.append(ERole.AM1)
            if ms.is_role_max(ERole.AM2):
                filter_roles.append(ERole.AM2)
            if ms.is_role_max(ERole.PM):
                filter_roles.append(ERole.PM)
            if filter_roles:
                filters.append(create_filter(ms.monitor, include=False, roles=filter_roles))

        # Not to monitor am if monitor pm 1 day before.
        if include_pre_day:
            pre_day = day - timedelta(days=1)
            if role := ms.schedule.get(pre_day):
                if role == ERole.PM:
                    filters.append(create_filter(ms.monitor, include=False, roles=MONITOR_ROLES_ALL))
                if role in MONITOR_ROLES_AM:
                    filters.append(create_filter(ms.monitor, include=False, roles=MONITOR_ROLES_AM))

            next_day = day + timedelta(days=1)
            if role := ms.schedule.get(next_day):
                if role in MONITOR_ROLES_AM:
                    filters.append(create_filter(ms.monitor, include=False, roles=MONITOR_ROLES_ALL))
    return filters


def create_filter(monitor, include=True, roles=None):
    if include:
        return lambda monitor_combo: monitor_combo.contains_monitor(monitor, roles)
    else:
        return lambda monitor_combo: not monitor_combo.contains_monitor(monitor, roles)


def output_schedules(ws, monitor_schedule_dict, weekday_dict):
    monitor_name_st_col = len(monitor_schedule_dict) + 4
    monitor_name_cols = {
        ERole.AM1: monitor_name_st_col,
        ERole.AM2: monitor_name_st_col + 1,
        ERole.PM: monitor_name_st_col + 2
    }
    for row_idx, day_priority in weekday_dict.items():
        for ms in monitor_schedule_dict.values():
            if (role := ms.schedule.get(day_priority.day)) and role in MONITOR_ROLES_ALL:
                ws.cell(row=row_idx, column=ms.col_idx, value=role.name)
                if col_idx := monitor_name_cols.get(role):
                    ws.cell(row=row_idx, column=col_idx, value=ms.monitor.name)


def elapsed_time(f):
    def wrapper(*args, **kwargs):
        st = time.time()
        v = f(*args, **kwargs)
        print(f'{f.__name__}: {time.time() - st}')
        return v

    return wrapper


@elapsed_time
def main(file_path='./schedules/MonitorSchedule2020_test.xlsm'):
    make_schedule(file_path)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        main()
