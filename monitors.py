# -*- coding: utf-8 -*-

from dataclasses import dataclass


@dataclass(frozen=True)
class Monitor:
    name: str
    is_fix_specialist: bool

