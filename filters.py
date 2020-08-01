# -*- coding: utf-8 -*-

from datetime import datetime, timedelta
from enum import Enum

from combo import ERole, NOT_AT_OFFICE_ROLES


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
            filters.append(lambda monitor_set: not(monitor_set >= must_work_at_office_monitors))
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
