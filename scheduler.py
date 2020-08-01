# -*- coding: utf-8 -*-

import copy
from dataclasses import dataclass
from datetime import datetime
from itertools import combinations
import openpyxl
import random
import sys
import time

from combo import MonitorSchedule, assign_remote_max, assign_role_maxes, gen_monitor_combos
from combo import ERole, MONITOR_ROLES_ALL, NOT_AT_OFFICE_ROLES, OUTPUT_ROLES
from filters import ERemoteFilters, FILTER_PRIORITY1, FILTER_PRIORITY2, get_filters_for_monitor_combo
import monitors

HEADER_ROW_IDX = 7
DATA_START_ROW_IDX = HEADER_ROW_IDX + 1
REMOTE_MAX_ROW_IDX = HEADER_ROW_IDX - 1


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
    all_monitors = monitor_schedule_dict.keys()
    sorted_day_priorities = sorted(weekday_dict.values(), key=lambda dp: dp.priority)
    weekdays = [dp.day for dp in sorted_day_priorities]
    days = len(weekday_dict)
    assign_role_maxes(monitor_schedule_dict, MONITOR_ROLES_ALL, days)

    monitor_schedule_dict = assign_monitors(all_monitors, monitor_schedule_dict, weekdays)

    load_manual_remote_max(ws, monitor_schedule_dict)
    assign_remote_max(monitor_schedule_dict, days)
    monitor_schedule_dict = assign_remotes(
        monitor_schedule_dict, must_work_at_office_groups, sorted(weekdays))

    fill_in_blanks_to(monitor_schedule_dict, weekdays, ERole.N)
    debug_schedules(monitor_schedule_dict, weekdays)
    print()
    for m, ms in monitor_schedule_dict.items():
        print(f'{m.name}\'s {ms.role_max[ERole.R]=}')
    print()

    # output_schedules(ws, monitor_schedule_dict, weekday_dict)
    # wb.save(excel_path)


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


def assign_monitors(all_monitors, monitor_schedule_dict, weekdays,
                    try_cnt1=1000, try_cnt2=1000):
    """
    監視当番の割り当てを行う。

    :param all_monitors: 監視者のset
    :param monitor_schedule_dict: 監視スケジュールのdict(key:=Monitor, item:=MonitorSchedule)
    :param weekdays: 営業日のIterable
    :param try_cnt1: 全フィルタを使用しての割り当て試行回数
    :param try_cnt2: 条件を緩くしての割り当て試行回数
    :return: 監視当番を割り当てた監視スケジュールのdict
    """
    cp_monitor_schedule_dict = copy_monitor_schedule_dict(monitor_schedule_dict)
    res_monitor_schedule_dict = _try_assign_monitors(
        all_monitors, monitor_schedule_dict, cp_monitor_schedule_dict, weekdays, try_cnt1,
        FILTER_PRIORITY2)
    if res_monitor_schedule_dict:
        return res_monitor_schedule_dict
    else:
        cp_monitor_schedule_dict = _try_assign_monitors(
            all_monitors, monitor_schedule_dict, cp_monitor_schedule_dict, weekdays, try_cnt2,
            FILTER_PRIORITY1)
        if cp_monitor_schedule_dict:
            return cp_monitor_schedule_dict
        else:
            _assign_monitors(all_monitors, monitor_schedule_dict, weekdays, True, True)
            return monitor_schedule_dict


def copy_monitor_schedule_dict(monitor_schedule_dict):
    cp = {}
    for m, ms in monitor_schedule_dict.items():
        cp[m] = copy.copy(ms)
    return cp


def _try_assign_monitors(all_monitors, msd, cp_msd, weekdays, try_cnt, filter_priority):
    for i in range(try_cnt):
        if _assign_monitors(all_monitors, cp_msd, weekdays, filter_priority):
            print(f'{filter_priority=}: {i + 1}: found.')
            return cp_msd
        cp_msd = copy_monitor_schedule_dict(msd)
    print(f'{filter_priority=}: {try_cnt}: not found.')
    return None


def _assign_monitors(monitor_set, monitor_schedule_dict, weekdays,
                     filter_priority, force_exec=False):
    """
    監視当番の割り振りを行う。

    :param monitor_set: 監視者のset
    :param monitor_schedule_dict: 監視スケジュールのdict(key:=Monitor, item:=MonitorSchedule)
    :param weekdays: 営業日のIterable
    :param filter_priority: フィルタ優先度
    :param force_exec: 均等な割り振りが不可の場合でも、その日を除いて処理を続行する場合はTrueを設定する
    :return: 割り振りが完了した場合はTrue
    """
    all_monitor_combos = set(gen_monitor_combos(monitor_set))
    ms_list = monitor_schedule_dict.values()
    for day in weekdays:
        filters = get_filters_for_monitor_combo(ms_list, day, filter_priority)
        # extract monitor combo that meets all filters.
        monitor_combos = [mc for mc in all_monitor_combos if all([f(mc) for f in filters])]
        if not monitor_combos:
            if force_exec:
                continue
            return False

        # Choice a monitor combo at random.
        monitor_combo = random.choice(monitor_combos)
        monitor_schedule_dict[monitor_combo.monitor_am1].schedule[day] = ERole.AM1
        monitor_schedule_dict[monitor_combo.monitor_am2].schedule[day] = ERole.AM2
        monitor_schedule_dict[monitor_combo.monitor_pm].schedule[day] = ERole.PM
    return True


def load_manual_remote_max(ws, monitor_schedule_dict):
    """
    手動で入力された在宅勤務数の上限を読み込み、MonitorScheduleに設定する。

    :param ws: 読み込むシート
    :param monitor_schedule_dict: 監視スケジュールのdict(key:=Monitor, item:=MonitorSchedule)
    :return: None
    """
    for ms in monitor_schedule_dict.values():
        if remote_max := ws.cell(row=REMOTE_MAX_ROW_IDX, column=ms.col_idx).value:
            ms.role_max[ERole.R] = remote_max


def assign_remotes(monitor_schedule_dict, must_work_at_office_groups, weekdays,
                   max_num_of_remotes_per_day=2, try_cnt1=500, try_cnt2=300):
    """
    在宅勤務の割り当てを行う。
    条件によっては割り当てられない日もある。

    :param monitor_schedule_dict: 監視スケジュールのdict(key:=Monitor, item:=MonitorSchedule)
    :param must_work_at_office_groups: 最低１人は出社する必要のある監視者の組み合わせのリスト
    :param weekdays: 営業日のIterable
    :param max_num_of_remotes_per_day: 1日の在宅勤務の割り当て人数
    :param try_cnt1: 全フィルタを使用しての割り当て試行回数
    :param try_cnt2: 条件を緩くしての割り当て試行回数
    :return: 在宅勤務を割り当てた監視スケジュールのdict
    """
    tmp_msd = _try_assign_remotes(monitor_schedule_dict, must_work_at_office_groups, weekdays,
                                  max_num_of_remotes_per_day, try_cnt1, FILTER_PRIORITY2)
    if tmp_msd:
        return tmp_msd
    tmp_msd = _try_assign_remotes(monitor_schedule_dict, must_work_at_office_groups, weekdays,
                                  max_num_of_remotes_per_day, try_cnt2, FILTER_PRIORITY1)
    if tmp_msd:
        return tmp_msd
    _assign_remotes(monitor_schedule_dict, must_work_at_office_groups, weekdays,
                    max_num_of_remotes_per_day, FILTER_PRIORITY2, True)
    return monitor_schedule_dict


def _try_assign_remotes(msd, must_work_at_office_groups, weekdays,
                        max_num_of_remotes_per_day, try_cnt, filter_priority):
    for i in range(try_cnt):
        cp_msd = copy_monitor_schedule_dict(msd)
        if _assign_remotes(cp_msd, must_work_at_office_groups, weekdays,
                           max_num_of_remotes_per_day, filter_priority):
            print(f'{filter_priority=}: {i + 1}: found.')
            return cp_msd
    print(f'{filter_priority=}: {try_cnt}: not found.')
    return None


def _assign_remotes(monitor_schedule_dict, must_work_at_office_groups, weekdays,
                    max_num_of_remotes_per_day, filter_priority, force_exec=False):
    """
    在宅勤務の割り当てを行う。
    条件によっては割り当てられない日もある。

    :param monitor_schedule_dict: 監視スケジュールのdict(key:=Monitor, item:=MonitorSchedule)
    :param must_work_at_office_groups: 最低１人は出社する必要のある監視者の組み合わせのリスト
    :param weekdays: 営業日のIterable
    :param max_num_of_remotes_per_day: 1日の在宅勤務の割り当て人数
    :param filter_priority: フィルタ優先度
    :param force_exec: 均等な割り振りが不可の場合でも、その日を除いて処理を続行する場合はTrueを設定する
    :return: 全日程で割り当てが試みられた場合はTrue
    """
    for day in weekdays:
        not_at_office_monitors = set()
        at_office_but_not_monitors = set()
        for m, ms in monitor_schedule_dict.items():
            role = ms.schedule.get(day)
            if role is None:
                at_office_but_not_monitors.add(m)
            elif role in NOT_AT_OFFICE_ROLES:
                not_at_office_monitors.add(m)
        num_of_remote_monitors = max_num_of_remotes_per_day - len(not_at_office_monitors)
        if num_of_remote_monitors <= 0:
            continue

        remote_groups = []
        remote_filters = get_remote_filters(
            monitor_schedule_dict, day, must_work_at_office_groups, filter_priority)

        for group in combinations(at_office_but_not_monitors, num_of_remote_monitors):
            g = set(group)
            if all([f(g) for f in remote_filters]):
                remote_groups.append(g)
        if not remote_groups:
            if force_exec:
                continue
            else:
                return False

        remote_monitor_set = random.choice(remote_groups)
        for m in remote_monitor_set:
            monitor_schedule_dict[m].schedule[day] = ERole.R
    return True


def get_remote_filters(ms_dict: dict, day: datetime, must_work_at_office_groups: list,
                       filter_priority: int):
    remote_filters = []
    for remote_filter in ERemoteFilters:
        if remote_filter.priority <= filter_priority:
            remote_filters.extend(
                remote_filter.get_filters(ms_dict, day, must_work_at_office_groups))
    return remote_filters


def fill_in_blanks_to(monitor_schedule_dict, weekdays, role):
    for day in weekdays:
        if [ms for ms in monitor_schedule_dict.values()
                if ms.schedule.get(day) == ERole.R]:
            for ms in monitor_schedule_dict.values():
                if day not in ms.schedule:
                    ms.schedule[day] = role


def output_schedules(ws, monitor_schedule_dict, weekday_dict):
    monitor_name_st_col = len(monitor_schedule_dict) + 4
    monitor_name_cols = {
        ERole.AM1: monitor_name_st_col,
        ERole.AM2: monitor_name_st_col + 1,
        ERole.PM: monitor_name_st_col + 2
    }
    for row_idx, day_priority in weekday_dict.items():
        for ms in monitor_schedule_dict.values():
            if (role := ms.schedule.get(day_priority.day)) and role in OUTPUT_ROLES:
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


def debug_schedules(monitor_schedule_dict, weekdays):
    date_str = 'date'
    print(f'{date_str: <19}: {ERole.AM1.name}, {ERole.AM2.name}, {ERole.PM.name}, '
          f'{ERole.N.name}, {ERole.R.name}')
    print_roles = {ERole.AM1, ERole.AM2, ERole.PM, ERole.N, ERole.R, }
    ms_list = monitor_schedule_dict.values()
    for day in sorted(weekdays):
        md = {}
        normals = []
        remotes = []
        for ms in ms_list:
            if (role := ms.schedule.get(day)) in print_roles:
                if role == ERole.R:
                    remotes.append(ms.monitor.name)
                elif role == ERole.N:
                    normals.append(ms.monitor.name)
                else:
                    md[role] = ms.monitor.name
        am1 = md.get(ERole.AM1)
        am2 = md.get(ERole.AM2)
        pm = md.get(ERole.PM)
        n = ' & '.join(normals)
        r = ' & '.join(remotes)
        print(f'{day}: {am1}, {am2}, {pm}, {n}, {r}')
    print()
    print('name, AM1, AM2, PM, SUM, R')
    for m, ms in monitor_schedule_dict.items():
        print(f'{m.name}, {ms.am1_count}, {ms.am2_count}, {ms.pm_count}, {ms.monitor_count}, '
              f'{ms.get_role_count(ERole.R)}')


@elapsed_time
def main(file_path='./schedules/MonitorSchedule2020_test.xlsm'):
    make_schedule(file_path)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        main()
