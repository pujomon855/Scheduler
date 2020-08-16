# -*- coding: utf-8 -*-

from datetime import datetime, timedelta
from enum import Enum
from openpyxl.worksheet.worksheet import Worksheet

from monitors import ERole, MONITOR_ROLES_ALL, MONITOR_ROLES_AM, NOT_AT_OFFICE_ROLES, Monitor


FILTER_PRIORITY1 = 1
FILTER_PRIORITY2 = 2


class FilterManager:
    _FILTER_DATA_ST_ROW_IDX = 7

    def __init__(self, filter_cls, ws: Worksheet, name_col_idx: int, disable_col_idx: int):
        self.filter_cls = filter_cls
        self.filters = set()
        for row in ws.iter_rows(min_row=FilterManager._FILTER_DATA_ST_ROW_IDX,
                                min_col=name_col_idx, max_col=disable_col_idx):
            filter_name = row[0].value
            if filter_name is None:
                break
            try:
                filter_enum = convert_str_to_filter(self.filter_cls, filter_name)
            except ValueError as e:
                print(f'{e}')
                continue
            else:
                if row[disable_col_idx - name_col_idx].value != 'Y':
                    self.filters.add(filter_enum)

    def get_filters(self, monitors, day: datetime, filter_priority=FILTER_PRIORITY2):
        raise NotImplementedError


def convert_str_to_filter(filter_cls, name: str):
    for e in filter_cls:
        if e.name == name:
            return e
    raise ValueError(f'{name} is not a member of {filter_cls}.')


class RemoteFilterManager(FilterManager):
    _NAME_COL_IDX = 9
    _DISABLE_COL_IDX = 11

    def __init__(self, ws, must_work_at_office_groups: list):
        super().__init__(ERemoteFilters, ws,
                         RemoteFilterManager._NAME_COL_IDX, RemoteFilterManager._DISABLE_COL_IDX)
        self.must_work_at_office_groups = must_work_at_office_groups

    def get_filters(self, monitors, day: datetime, filter_priority=FILTER_PRIORITY2):
        filters = []
        for filter_enum in self.filters:
            if filter_enum.priority <= filter_priority:
                filters.extend(
                    filter_enum.get_filters(monitors, day, self.must_work_at_office_groups))
        return filters


class MonitorFilterManager(FilterManager):
    _NAME_COL_IDX = 3
    _DISABLE_COL_IDX = 5

    def __init__(self, ws):
        super().__init__(EMonitorComboFilters, ws,
                         MonitorFilterManager._NAME_COL_IDX, MonitorFilterManager._DISABLE_COL_IDX)

    def get_filters(self, monitors, day, filter_priority=FILTER_PRIORITY2):
        filters = []
        for monitors in monitors:
            for filter_enum in self.filters:
                if filter_enum.priority <= filter_priority:
                    filters.extend(filter_enum.get_filters(monitors, day))
        return filters


# Filters for remotes

# key:=(Monitor, bool), item:=filter function
__MONITOR_FILTERS = {}


def filter_remote_2days_in_a_row(
        monitors: list, day: datetime, must_work_at_office_groups: list):
    pre_day = day - timedelta(days=1)
    next_day = day + timedelta(days=1)
    filters = []
    for monitor in monitors:
        if monitor.schedule.get(pre_day) == ERole.R or monitor.schedule.get(next_day) == ERole.R:
            filters.append(_get_and_set_if_absent_monitor_name_filter(monitor.name, False))
    return filters


def filter_must_work_at_office(
        monitors: list, day: datetime, must_work_at_office_groups: list):
    filters = []
    not_office_monitor_names = {monitor.name for monitor in monitors
                                if monitor.schedule.get(day) in NOT_AT_OFFICE_ROLES}
    for must_work_at_office_group in must_work_at_office_groups:
        must_work_at_office_monitors = must_work_at_office_group - not_office_monitor_names
        if must_work_at_office_monitors:
            filters.append(_create_filter_func(must_work_at_office_monitors))
    return filters


def _create_filter_func(must_work_at_office_monitors):
    def func(monitor_name_set: set):
        return not (monitor_name_set >= must_work_at_office_monitors)
    return func


def filter_remote_max(
        monitors: list, day: datetime, must_work_at_office_groups: list):
    filters = []
    for monitor in monitors:
        if monitor.is_role_max(ERole.R):
            filters.append(_get_and_set_if_absent_monitor_name_filter(monitor.name, False))
    return filters


def _get_and_set_if_absent_monitor_name_filter(monitor_name: str, include: bool):
    key = (monitor_name, include)
    monitor_filter = __MONITOR_FILTERS.get(key)
    if not monitor_filter:
        monitor_filter = _create_monitor_filter(*key)
        __MONITOR_FILTERS[key] = monitor_filter
    return monitor_filter


def _create_monitor_filter(monitor_name: str, include: bool):
    if include:
        return lambda monitor_name_set: monitor_name in monitor_name_set
    else:
        return lambda monitor_name_set: monitor_name not in monitor_name_set


class ERemoteFilters(Enum):
    REMOTE_2DAYS_IN_A_ROW = (FILTER_PRIORITY2, filter_remote_2days_in_a_row)
    MUST_WORK_AT_OFFICE_GROUP = (FILTER_PRIORITY1, filter_must_work_at_office)
    REMOTE_MAX = (FILTER_PRIORITY1, filter_remote_max)

    def __init__(self, priority: int, filter_func):
        self.__priority: int = priority
        self.__filter_func = filter_func

    @property
    def priority(self) -> int:
        return self.__priority

    def get_filters(self, monitors: list, day: datetime, must_work_at_office_groups: list):
        return self.__filter_func(monitors, day, must_work_at_office_groups)

    def __repr__(self):
        return f'({self.__priority}, {self.name})'


# Filters for MonitorCombo

def filter_manual_input(monitor: Monitor, day: datetime):
    filters = []
    # Manually input day role
    role = monitor.schedule.get(day)
    if role:
        if role in MONITOR_ROLES_ALL:
            filters.append(_create_monitor_combo_filter(monitor.name, roles=[role]))
        elif role == ERole.OTHER:
            filters.append(_create_monitor_combo_filter(monitor.name, include=False))
    return filters


def filter_monitoring_max(monitor: Monitor, day: datetime):
    filters = []
    filter_roles = []
    if monitor.is_role_max(ERole.AM1):
        filter_roles.append(ERole.AM1)
    if monitor.is_role_max(ERole.AM2):
        filter_roles.append(ERole.AM2)
    if monitor.is_role_max(ERole.PM):
        filter_roles.append(ERole.PM)
    if filter_roles:
        filters.append(
            _create_monitor_combo_filter(monitor.name, include=False, roles=filter_roles))
    return filters


def filter_am_am_in_a_row(monitor: Monitor, day: datetime):
    filters = []
    pre_day = day - timedelta(days=1)
    next_day = day + timedelta(days=1)
    if (monitor.schedule.get(pre_day) in MONITOR_ROLES_AM or
            monitor.schedule.get(next_day) in MONITOR_ROLES_AM):
        filters.append(_create_monitor_combo_filter(
            monitor.name, include=False, roles=MONITOR_ROLES_AM))
    return filters


def filter_pm_am_in_a_row(monitor: Monitor, day: datetime):
    filters = []
    pre_day = day - timedelta(days=1)
    next_day = day + timedelta(days=1)
    if monitor.schedule.get(pre_day) == ERole.PM:
        filters.append(_create_monitor_combo_filter(
            monitor.name, include=False, roles=MONITOR_ROLES_AM))
    if monitor.schedule.get(next_day) == ERole.PM:
        filters.append(_create_monitor_combo_filter(
            monitor.name, include=False, roles=[ERole.PM]))
    return filters


def filter_pm_pm_in_a_row(monitor: Monitor, day: datetime):
    filters = []
    pre_day = day - timedelta(days=1)
    next_day = day + timedelta(days=1)
    if monitor.schedule.get(pre_day) == ERole.PM or monitor.schedule.get(next_day) == ERole.PM:
        filters.append(_create_monitor_combo_filter(
            monitor.name, include=False, roles=[ERole.PM]))
    return filters


def _create_monitor_combo_filter(monitor_name: str, include: bool = True, roles=None):
    def filter_func(monitor_combo: dict) -> bool:
        is_include = False
        if roles:
            for role, name in monitor_combo.items():
                if role in roles and name == monitor_name:
                    is_include = True
        else:
            is_include = monitor_name in monitor_combo.values()
        return is_include if include else not is_include

    return filter_func


class EMonitorComboFilters(Enum):
    MANUAL_INPUT = (FILTER_PRIORITY1, filter_manual_input)
    MONITORING_MAX = (FILTER_PRIORITY1, filter_monitoring_max)
    AM_AM_IN_A_ROW = (FILTER_PRIORITY2, filter_am_am_in_a_row)
    PM_AM_IN_A_ROW = (FILTER_PRIORITY2, filter_pm_am_in_a_row)
    PM_PM_IN_A_ROW = (FILTER_PRIORITY2, filter_pm_pm_in_a_row)

    def __init__(self, priority, filter_func):
        self.__priority = priority
        self.__filter_func = filter_func

    @property
    def priority(self):
        return self.__priority

    def get_filters(self, monitor: Monitor, day: datetime):
        return self.__filter_func(monitor, day)

    def __repr__(self):
        return f'({self.__priority}, {self.name})'
