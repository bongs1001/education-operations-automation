from datetime import datetime, timedelta
import unittest

from attendance_reference.analyze_samples import class_gap_candidates, disconnected_intervals, overlaps_window, parse_duration, parse_timestamp, subtract_breaks, unique_index


class SessionTests(unittest.TestCase):
    def setUp(self):
        self.base = datetime(2026, 9, 7, 9)

    def interval(self, a, b):
        return self.base + timedelta(seconds=a), self.base + timedelta(seconds=b)

    def test_ten_minute_boundary(self):
        self.assertEqual(disconnected_intervals([self.interval(0, 60), self.interval(659, 800)]), [])
        self.assertEqual(disconnected_intervals([self.interval(0, 60), self.interval(660, 800)]), [self.interval(60, 660)])

    def test_unsorted_overlapping_and_nested_devices(self):
        self.assertEqual(disconnected_intervals([self.interval(1900, 2000), self.interval(20, 40), self.interval(0, 1200), self.interval(1100, 1300)]), [self.interval(1300, 1900)])

    def test_touching_and_single_session(self):
        self.assertEqual(disconnected_intervals([self.interval(0, 100), self.interval(100, 1000)]), [])
        self.assertEqual(disconnected_intervals([self.interval(1000, 2000)]), [])
        self.assertEqual(disconnected_intervals([]), [])

    def test_invalid_interval(self):
        with self.assertRaises(ValueError):
            disconnected_intervals([self.interval(100, 0)])

    def test_timestamps_and_duration(self):
        self.assertEqual(parse_timestamp("9/07/26, 12:29:39 PM").hour, 12)
        self.assertEqual(parse_timestamp("9/07/26, 12:29:39 AM").hour, 0)
        self.assertEqual(parse_duration("8시간 5분 41초"), 29141)
        self.assertEqual(parse_duration("4초"), 4)

    def test_duplicate_name_blocks_mapping(self):
        with self.assertRaises(ValueError):
            unique_index([{"name": "홍길동"}, {"name": " 홍길동 "}], "name")

    def test_break_splits_short_gaps_instead_of_adding_them(self):
        sessions = [self.interval(0, 60), self.interval(1260, 1400)]
        self.assertEqual(class_gap_candidates(sessions, *self.interval(0, 2000), [self.interval(600, 720)]), [])

    def test_lunch_and_class_bounds(self):
        sessions = [self.interval(0, 30), self.interval(1800, 1900)]
        self.assertEqual(class_gap_candidates(sessions, *self.interval(60, 1600), [self.interval(100, 900)]), [self.interval(900, 1600)])

    def test_overlapping_breaks_and_full_coverage(self):
        self.assertEqual(subtract_breaks(*self.interval(100, 1000), [self.interval(0, 500), self.interval(400, 1100)]), [])

    def test_user_sample_lunch_shift(self):
        at = lambda text: datetime.strptime("2026-09-07 " + text, "%Y-%m-%d %H:%M:%S")
        sessions = [(at("08:57:18"), at("12:44:18")), (at("15:01:47"), at("17:50:29"))]
        expected = [(at("14:00:00"), at("15:01:47"))]
        for lunch_start in ("12:41:00", "12:45:00", "12:50:00"):
            self.assertEqual(class_gap_candidates(sessions, at("09:10:00"), at("17:50:00"), [(at(lunch_start), at("14:00:00"))]), expected)

    def test_project_check_window_overlap(self):
        window = self.interval(600, 1200)
        self.assertTrue(overlaps_window([self.interval(0, 600)], *window))
        self.assertTrue(overlaps_window([self.interval(1200, 1300)], *window))
        self.assertFalse(overlaps_window([self.interval(0, 599)], *window))


if __name__ == "__main__":
    unittest.main()
