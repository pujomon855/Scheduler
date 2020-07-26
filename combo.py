# -*- coding: utf-8 -*-

from dataclasses import dataclass
from enum import Enum, auto
import itertools

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
        self.col_idx = col_idx

    @property
    def am1_count(self):
        return len([am_role for am_role in self.schedule.values()
                    if am_role == ERole.AM1])

    @property
    def am2_count(self):
        return len([am_role for am_role in self.schedule.values()
                    if am_role == ERole.AM2])

    @property
    def pm_count(self):
        return len([am_role for am_role in self.schedule.values()
                    if am_role == ERole.PM])

    @property
    def monitor_count(self):
        return len([am_role for am_role in self.schedule.values()
                    if am_role in (ERole.AM1, ERole.AM2, ERole.PM)])

    def __repr__(self):
        return f'{self.monitor.name}\'s schedule: {self.col_idx=}'


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
