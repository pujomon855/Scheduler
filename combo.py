# -*- coding: utf-8 -*-

from dataclasses import dataclass
import itertools

from monitors import Monitor


@dataclass(frozen=True)
class MonitorCombo:
    AM1: Monitor
    AM2: Monitor
    PM: Monitor


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

