import unittest


class GetMonitorCombos(unittest.TestCase):
    def test_less_than_3members(self):
        from combo import gen_monitor_combos
        from monitors import Monitor

        members = {
            Monitor('A', True),
            Monitor('B', True),
        }
        expected = []
        actual = list(gen_monitor_combos(members))
        self.assertEqual(expected, actual)

    def test_3members_1fix(self):
        from combo import gen_monitor_combos, MonitorCombo
        from monitors import Monitor

        m1 = Monitor('A', True)
        m2 = Monitor('B', False)
        m3 = Monitor('C', False)
        members = {m1, m2, m3}
        expected = [
            MonitorCombo(m1, m2, m3), MonitorCombo(m2, m1, m3),
            MonitorCombo(m1, m3, m2), MonitorCombo(m3, m1, m2),
        ]
        actual = list(gen_monitor_combos(members))
        self.assertCountEqual(expected, actual)

    def test_3members_3fix(self):
        from combo import gen_monitor_combos, MonitorCombo
        from monitors import Monitor

        m1 = Monitor('A', True)
        m2 = Monitor('B', True)
        m3 = Monitor('C', True)
        members = {m1, m2, m3}
        expected = [
            MonitorCombo(m1, m2, m3), MonitorCombo(m2, m1, m3),
            MonitorCombo(m1, m3, m2), MonitorCombo(m3, m1, m2),
            MonitorCombo(m2, m3, m1), MonitorCombo(m3, m2, m1),
        ]
        actual = list(gen_monitor_combos(members))
        self.assertCountEqual(expected, actual)

    def test_4members_2fix(self):
        from combo import gen_monitor_combos, MonitorCombo
        from monitors import Monitor

        m1 = Monitor('A', True)
        m2 = Monitor('B', True)
        m3 = Monitor('C', False)
        m4 = Monitor('D', False)
        members = {m1, m2, m3, m4}
        expected = [
            MonitorCombo(m1, m2, m3), MonitorCombo(m1, m2, m4),
            MonitorCombo(m2, m1, m3), MonitorCombo(m2, m1, m4),
            MonitorCombo(m1, m3, m2), MonitorCombo(m1, m3, m4),
            MonitorCombo(m3, m1, m2), MonitorCombo(m3, m1, m4),
            MonitorCombo(m1, m4, m2), MonitorCombo(m1, m4, m3),
            MonitorCombo(m4, m1, m2), MonitorCombo(m4, m1, m3),
            MonitorCombo(m2, m3, m1), MonitorCombo(m2, m3, m4),
            MonitorCombo(m3, m2, m1), MonitorCombo(m3, m2, m4),
            MonitorCombo(m2, m4, m1), MonitorCombo(m2, m4, m3),
            MonitorCombo(m4, m2, m1), MonitorCombo(m4, m2, m3),
        ]
        actual = list(gen_monitor_combos(members))
        self.assertCountEqual(expected, actual)


if __name__ == '__main__':
    unittest.main()
