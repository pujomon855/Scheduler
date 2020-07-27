# -*- coding: utf-8 -*-

from dataclasses import dataclass
from enum import Enum, auto
import itertools
import random

from monitors import Monitor


class ERole(Enum):
    AM1 = auto()
    AM2 = auto()
    PM = auto()
    N = auto()
    R = auto()
    OTHER = auto()


@dataclass(frozen=True)
class MonitorCombo:
    monitor_am1: Monitor
    monitor_am2: Monitor
    monitor_pm: Monitor

    def contains_monitor(self, monitor, roles=None):
        """
        monitorがこの組み合わせに含まれているかを判定する。
        rolesの指定があれば、指定のroleで組み合わせに含まれているかを判定する。

        :param monitor: 監視者
        :param roles: 当番の種類のリスト(iterable)
        :return: monitorがこの組み合わせに含まれている場合はTrue、そうでない場合はFalse
        """
        if roles:
            if ERole.AM1 in roles:
                if monitor == self.monitor_am1:
                    return True
            if ERole.AM2 in roles:
                if monitor == self.monitor_am2:
                    return True
            if ERole.PM in roles:
                if monitor == self.monitor_pm:
                    return True
        else:
            return (monitor == self.monitor_am1 or
                    monitor == self.monitor_am2 or
                    monitor == self.monitor_pm)
        return False


class MonitorSchedule:
    def __init__(self, monitor, col_idx):
        self.monitor = monitor
        # key: datetime.datetime, item: ERole
        self.schedule = {}
        # key: ERole, item: max
        self.role_max = {}
        self.col_idx = col_idx

    @property
    def am1_count(self):
        return len([role for role in self.schedule.values()
                    if role == ERole.AM1])

    @property
    def am2_count(self):
        return len([role for role in self.schedule.values()
                    if role == ERole.AM2])

    @property
    def pm_count(self):
        return len([role for role in self.schedule.values()
                    if role == ERole.PM])

    @property
    def monitor_count(self):
        return len([role for role in self.schedule.values()
                    if role in (ERole.AM1, ERole.AM2, ERole.PM)])

    @property
    def max_monitor_count(self):
        return sum(self.role_max.values())

    def is_role_max(self, role):
        """
        この監視者に割り当てられた役割の日数が設定された上限値に達しているかを判定する。
        役割に上限が設定されていなかった場合はFalseを返す。

        :param role: 役割
        :return: この監視者に割り当てられた役割の日数が設定された上限値に達している場合はTrue
        """
        if max_count := self.role_max.get(role):
            return len([r for r in self.schedule.values() if r == role]) >= max_count
        return False

    def __repr__(self):
        return f'{self.monitor.name}\'s schedule: {self.col_idx=}'

    def __copy__(self):
        cp = MonitorSchedule(self.monitor, self.col_idx)
        cp.schedule = self.schedule.copy()
        cp.role_max = self.role_max.copy()
        return cp


def gen_monitor_combos(members):
    """
    監視の組み合わせのgeneratorを返す。

    :param members: 全監視メンバー(None不可)
    :return: 監視の組み合わせ(generator)
    """
    if len(members) < 3:
        return
    for m1, m2, m3 in itertools.permutations(members, 3):
        if m1.is_fix_specialist or m2.is_fix_specialist:
            yield MonitorCombo(m1, m2, m3)


def assign_role_maxes(monitor_schedule_dict, roles, days):
    """
    各監視者に割り当てられた役割の日数の上限値の合計が等しくなるようにランダムに上限を設定する。
    合計値の差は最大1とする。

    :param monitor_schedule_dict: 上限を設定するMonitorScheduleのdict
        (key:=Monitor, item:=MonitorSchedule)
    :param roles: ERoleのIterable
    :param days: 割り当て日数
    :return: None
    """
    num_of_monitors = len(monitor_schedule_dict)
    min_max_cnt = int(days / num_of_monitors)
    num_of_hi_freq_monitors = days % num_of_monitors
    if num_of_hi_freq_monitors == 0:
        for ms in monitor_schedule_dict.values():
            for role in roles:
                ms.role_max[role] = min_max_cnt
        return

    monitors = monitor_schedule_dict.keys()
    hi_freq_monitors = set(random.sample(monitors, num_of_hi_freq_monitors))
    for role in roles:
        for m in hi_freq_monitors:
            monitor_schedule_dict[m].role_max[role] = min_max_cnt + 1

        lo_freq_monitors = monitors - hi_freq_monitors
        for m in lo_freq_monitors:
            monitor_schedule_dict[m].role_max[role] = min_max_cnt
        hi_freq_monitors = _find_lower_frequency(monitor_schedule_dict, num_of_hi_freq_monitors)


def _find_lower_frequency(monitor_schedule_dict, find_num):
    """
    監視当番数の合計が少ない監視者のsetを、サイズがfind_numとなるように返す。
    合計が同値でfind_numを超える監視者は、同値の中でランダムに選出される。
    例：監視当番の合計が[4, 4, 5, 5, 5, 6]で、find_numが3の場合、
        返されるsetは監視当番の合計が4のMonitorと、5のMonitorの中からランダムに1人のsetとなる。

    :param monitor_schedule_dict: MonitorScheduleのdict(key:=Monitor, item:=MonitorSchedule)
    :param find_num: 検索する監視者数(0 < find_num <= len(monitor_schedule_dict))
    :return: 監視当番数の合計が少ない監視者(Monitor)のset
    """
    sorted_ms = sorted(monitor_schedule_dict.values(), key=lambda ms: ms.max_monitor_count)
    monitor_set = set()
    cur_mon_cnt = 0
    for monitor_schedule in sorted_ms:
        if find_num <= len(monitor_set) and cur_mon_cnt < monitor_schedule.max_monitor_count:
            break
        monitor_set.add(monitor_schedule.monitor)
        cur_mon_cnt = monitor_schedule.max_monitor_count

    if find_num >= len(monitor_set):
        return monitor_set

    # extract monitors whose monitor_count is lower than cur_mon_cnt
    tmp_monitor_set = {ms.monitor for ms in sorted_ms if ms.max_monitor_count < cur_mon_cnt}
    # now, monitor_set is only including monitors whose monitor_count is cur_mon_cnt
    monitor_set -= tmp_monitor_set
    # returns tmp_monitor_set and selected monitors from monitor_set at random
    monitor_set = tmp_monitor_set | set(random.sample(monitor_set, find_num - len(tmp_monitor_set)))
    return monitor_set
