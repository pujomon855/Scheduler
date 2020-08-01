# -*- coding: utf-8 -*-

from datetime import datetime, timedelta
from enum import Enum

from combo import ERole, MONITOR_ROLES_ALL, MONITOR_ROLES_AM, NOT_AT_OFFICE_ROLES, MonitorSchedule


# Filters for remotes

# key:=(Monitor, bool), item:=filter function
__MONITOR_FILTERS = {}


def filter_remote_2days_in_a_row(
        ms_dict: dict, day: datetime, must_work_at_office_groups: list):
    pre_day = day - timedelta(days=1)
    next_day = day + timedelta(days=1)
    filters = []
    for ms in ms_dict.values():
        if ms.schedule.get(pre_day) == ERole.R or ms.schedule.get(next_day) == ERole.R:
            filters.append(_get_and_set_if_absent_monitor_filter(ms.monitor, False))
    return filters


def filter_must_work_at_office(
        ms_dict: dict, day: datetime, must_work_at_office_groups: list):
    filters = []
    not_office_monitors = {ms.monitor for ms in ms_dict.values()
                           if ms.schedule.get(day) in NOT_AT_OFFICE_ROLES}
    for must_work_at_office_group in must_work_at_office_groups:
        if must_work_at_office_monitors := must_work_at_office_group - not_office_monitors:
            filters.append(lambda monitor_set: not (monitor_set >= must_work_at_office_monitors))
    return filters


def filter_remote_max(
        ms_dict: dict, day: datetime, must_work_at_office_groups: list):
    filters = []
    for m, ms in ms_dict.items():
        if ms.is_role_max(ERole.R):
            filters.append(_get_and_set_if_absent_monitor_filter(m, False))
    return filters


def _get_and_set_if_absent_monitor_filter(monitor, include):
    key = (monitor, include)
    monitor_filter = __MONITOR_FILTERS.get(key)
    if not monitor_filter:
        monitor_filter = _create_monitor_filter(*key)
        __MONITOR_FILTERS[key] = monitor_filter
    return monitor_filter


def _create_monitor_filter(monitor, include):
    if include:
        return lambda monitor_set: monitor in monitor_set
    else:
        return lambda monitor_set: monitor not in monitor_set


FILTER_PRIORITY1 = 1
FILTER_PRIORITY2 = 2


class ERemoteFilters(Enum):
    REMOTE_2DAYS_IN_A_ROW = (FILTER_PRIORITY2, filter_remote_2days_in_a_row)
    MUST_WORK_AT_OFFICE_GROUP = (FILTER_PRIORITY1, filter_must_work_at_office)
    REMOTE_MAX = (FILTER_PRIORITY1, filter_remote_max)

    def __init__(self, priority, filter_func):
        self.__priority = priority
        self.__filter_func = filter_func

    @property
    def priority(self):
        return self.__priority

    def get_filters(self, ms_dict: dict, day: datetime, must_work_at_office_groups: list):
        return self.__filter_func(ms_dict, day, must_work_at_office_groups)


# Filters for MonitorCombo

def filter_manual_input(ms: MonitorSchedule, day: datetime):
    filters = []
    # Manually input day role
    if role := ms.schedule.get(day):
        if role in MONITOR_ROLES_ALL:
            filters.append(_create_monitor_combo_filter(ms.monitor, roles=[role]))
        elif role == ERole.OTHER:
            filters.append(_create_monitor_combo_filter(ms.monitor, include=False))
    return filters


def filter_monitoring_max(ms: MonitorSchedule, day: datetime):
    filters = []
    filter_roles = []
    if ms.is_role_max(ERole.AM1):
        filter_roles.append(ERole.AM1)
    if ms.is_role_max(ERole.AM2):
        filter_roles.append(ERole.AM2)
    if ms.is_role_max(ERole.PM):
        filter_roles.append(ERole.PM)
    if filter_roles:
        filters.append(
            _create_monitor_combo_filter(ms.monitor, include=False, roles=filter_roles))
    return filters


def filter_am_am_in_a_row(ms: MonitorSchedule, day: datetime):
    filters = []
    pre_day = day - timedelta(days=1)
    next_day = day + timedelta(days=1)
    if (ms.schedule.get(pre_day) in MONITOR_ROLES_AM or
            ms.schedule.get(next_day) in MONITOR_ROLES_AM):
        filters.append(_create_monitor_combo_filter(
            ms.monitor, include=False, roles=MONITOR_ROLES_AM))
    return filters


def filter_pm_am_in_a_row(ms: MonitorSchedule, day: datetime):
    filters = []
    pre_day = day - timedelta(days=1)
    next_day = day + timedelta(days=1)
    if ms.schedule.get(pre_day) == ERole.PM:
        filters.append(_create_monitor_combo_filter(
            ms.monitor, include=False, roles=MONITOR_ROLES_AM))
    if ms.schedule.get(next_day) == ERole.PM:
        filters.append(_create_monitor_combo_filter(
            ms.monitor, include=False, roles=[ERole.PM]))
    return filters


def filter_pm_pm_in_a_row(ms: MonitorSchedule, day: datetime):
    filters = []
    pre_day = day - timedelta(days=1)
    next_day = day + timedelta(days=1)
    if ms.schedule.get(pre_day) == ERole.PM or ms.schedule.get(next_day) == ERole.PM:
        filters.append(_create_monitor_combo_filter(
            ms.monitor, include=False, roles=[ERole.PM]))
    return filters


# def filter_by_pre_day_role(ms_list, day: datetime):
#     filters = []
#     for ms in ms_list:
#         pre_day = day - timedelta(days=1)
#         if role := ms.schedule.get(pre_day):
#             if role == ERole.PM:
#                 filters.append(_create_monitor_combo_filter(
#                     ms.monitor, include=False, roles=MONITOR_ROLES_ALL))
#             if role in MONITOR_ROLES_AM:
#                 filters.append(_create_monitor_combo_filter(
#                     ms.monitor, include=False, roles=MONITOR_ROLES_AM))
#
#         next_day = day + timedelta(days=1)
#         if role := ms.schedule.get(next_day):
#             if role in MONITOR_ROLES_AM:
#                 filters.append(_create_monitor_combo_filter(
#                     ms.monitor, include=False, roles=MONITOR_ROLES_ALL))


def _create_monitor_combo_filter(monitor, include=True, roles=None):
    if include:
        return lambda monitor_combo: monitor_combo.contains_monitor(monitor, roles)
    else:
        return lambda monitor_combo: not monitor_combo.contains_monitor(monitor, roles)


class EMonitorComboFilters(Enum):
    MANUAL_INPUT = (FILTER_PRIORITY1, filter_manual_input)
    MONITORING_MAX = (FILTER_PRIORITY1, filter_monitoring_max)
    AM_AM_IN_A_ROW = (FILTER_PRIORITY2, filter_am_am_in_a_row)
    PM_AM_IN_A_ROW = (FILTER_PRIORITY2, filter_pm_am_in_a_row)
    # PM_PM_IN_A_ROW = (FILTER_PRIORITY2, filter_pm_pm_in_a_row)

    def __init__(self, priority, filter_func):
        self.__priority = priority
        self.__filter_func = filter_func

    @property
    def priority(self):
        return self.__priority

    def get_filters(self, ms: MonitorSchedule, day: datetime):
        return self.__filter_func(ms, day)


def get_filters_for_monitor_combo(ms_list, day, filter_priority=FILTER_PRIORITY2):
    filters = []
    for ms in ms_list:
        for mc_filter in EMonitorComboFilters:
            if mc_filter.priority <= filter_priority:
                filters.extend(mc_filter.get_filters(ms, day))
    return filters