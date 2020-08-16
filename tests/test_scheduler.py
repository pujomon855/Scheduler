import unittest

from monitors import ERole, Monitor


class GetMonitorCombos(unittest.TestCase):
    def test_less_than_3members(self):
        from monitors import Monitor
        from scheduler import gen_monitor_combos

        members = {
            Monitor('A', True),
            Monitor('B', True),
        }
        expected = []
        actual = list(gen_monitor_combos(members))
        self.assertEqual(expected, actual)

    def test_3members_1fix(self):
        from monitors import Monitor
        from scheduler import gen_monitor_combos

        m1 = Monitor('A', True)
        m2 = Monitor('B', False)
        m3 = Monitor('C', False)
        members = [m1, m2, m3]
        expected = [
            _create_monitor_combo(m1, m2, m3), _create_monitor_combo(m2, m1, m3),
            _create_monitor_combo(m1, m3, m2), _create_monitor_combo(m3, m1, m2),
        ]
        actual = list(gen_monitor_combos(members))
        self.assertCountEqual(expected, actual)

    def test_3members_3fix(self):
        from monitors import Monitor
        from scheduler import gen_monitor_combos

        m1 = Monitor('A', True)
        m2 = Monitor('B', True)
        m3 = Monitor('C', True)
        members = [m1, m2, m3]
        expected = [
            _create_monitor_combo(m1, m2, m3), _create_monitor_combo(m2, m1, m3),
            _create_monitor_combo(m1, m3, m2), _create_monitor_combo(m3, m1, m2),
            _create_monitor_combo(m2, m3, m1), _create_monitor_combo(m3, m2, m1),
        ]
        actual = list(gen_monitor_combos(members))
        self.assertCountEqual(expected, actual)

    def test_4members_2fix(self):
        from monitors import Monitor
        from scheduler import gen_monitor_combos

        m1 = Monitor('A', True)
        m2 = Monitor('B', True)
        m3 = Monitor('C', False)
        m4 = Monitor('D', False)
        members = [m1, m2, m3, m4]
        expected = [
            _create_monitor_combo(m1, m2, m3), _create_monitor_combo(m1, m2, m4),
            _create_monitor_combo(m2, m1, m3), _create_monitor_combo(m2, m1, m4),
            _create_monitor_combo(m1, m3, m2), _create_monitor_combo(m1, m3, m4),
            _create_monitor_combo(m3, m1, m2), _create_monitor_combo(m3, m1, m4),
            _create_monitor_combo(m1, m4, m2), _create_monitor_combo(m1, m4, m3),
            _create_monitor_combo(m4, m1, m2), _create_monitor_combo(m4, m1, m3),
            _create_monitor_combo(m2, m3, m1), _create_monitor_combo(m2, m3, m4),
            _create_monitor_combo(m3, m2, m1), _create_monitor_combo(m3, m2, m4),
            _create_monitor_combo(m2, m4, m1), _create_monitor_combo(m2, m4, m3),
            _create_monitor_combo(m4, m2, m1), _create_monitor_combo(m4, m2, m3),
        ]
        actual = list(gen_monitor_combos(members))
        self.assertCountEqual(expected, actual)


def _create_monitor_combo(m1: Monitor, m2: Monitor, m3: Monitor):
    return {ERole.AM1: m1.name, ERole.AM2: m2.name, ERole.PM: m3.name}


if __name__ == '__main__':
    unittest.main()
