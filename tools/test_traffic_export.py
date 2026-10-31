"""Offline fixtures for the traffic-history builder and alert."""
import json
import os
import stat
import tempfile
import unittest
from datetime import UTC, datetime
from pathlib import Path

import traffic_export as traffic


def write_api(folder: Path, views: int = 10, clones: int = 4, stars: int = 2, forks: int = 0,
              referrers: list[dict[str, object]] | None = None) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    days = ["2026-10-30", "2026-10-31"]
    (folder / "views.json").write_text(json.dumps({"count": views, "uniques": 3, "views": [
        {"timestamp": f"{day}T00:00:00Z", "count": 5, "uniques": 2} for day in days]}))
    (folder / "clones.json").write_text(json.dumps({"count": clones, "uniques": 1, "clones": [
        {"timestamp": f"{day}T00:00:00Z", "count": 2, "uniques": 1} for day in days]}))
    default_referrers: list[dict[str, object]] = [{"referrer": "github.com", "count": 7, "uniques": 2}]
    (folder / "referrers.json").write_text(json.dumps(referrers or default_referrers))
    (folder / "paths.json").write_text(json.dumps([{"path": "/r", "title": "Repo", "count": 9, "uniques": 4}]))
    (folder / "repo.json").write_text(json.dumps({"stargazers_count": stars, "forks_count": forks,
                                                 "subscribers_count": 1, "open_issues_count": 5}))


class TrafficFixture(unittest.TestCase):
    def setUp(self) -> None:
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.api, self.data = self.root / "api", self.root / "data"

    def build(self, day: str, views: int = 10, stars: int = 2,
              referrers: list[dict[str, object]] | None = None) -> None:
        write_api(self.api, views=views, stars=stars, referrers=referrers)
        traffic.build(self.api, self.data, datetime.fromisoformat(f"{day}T06:00:00+00:00"))


class BuildTests(TrafficFixture):
    def test_first_run_writes_history_and_badges(self) -> None:
        self.build("2026-10-31")
        self.assertEqual((self.data / "traffic.csv").read_bytes(), (
            b"snapshot_date,snapshot_datetime_utc,views_14d,unique_views_14d,clones_14d,unique_clones_14d,"
            b"stars,forks,watchers,open_issues\r\n2026-10-31,2026-10-31T06:00:00+00:00,10,3,4,1,2,0,1,5\r\n"))
        self.assertEqual((self.data / "traffic_daily.csv").read_text(encoding="utf-8").splitlines()[1:],
                         ["2026-10-30,5,2,2,1", "2026-10-31,5,2,2,1"])
        self.assertEqual((self.data / "badges" / "clones.json").read_text(),
                         '{"schemaVersion": 1, "label": "clones", "message": "4", "color": "217346"}\n')
        self.assertEqual(json.loads((self.data / "badges" / "visits.json").read_text())["message"], "4")
        self.assertEqual(json.loads((self.data / "badges" / "visitors.json").read_text())["label"], "visitors (14d)")
        self.assertIn("2026-10-31,1,/r,Repo,9,4", (self.data / "popular_paths.csv").read_text(encoding="utf-8"))

    def test_same_day_rerun_replaces_the_snapshot(self) -> None:
        self.build("2026-10-31")
        self.build("2026-10-31", views=20, referrers=[{"referrer": "other.example", "count": 1, "uniques": 1}])
        rows = traffic.read_csv(self.data / "traffic.csv")
        self.assertEqual([(r["snapshot_date"], r["views_14d"]) for r in rows], [("2026-10-31", "20")])
        referrers = traffic.read_csv(self.data / "referrers.csv")
        self.assertEqual([r["referrer"] for r in referrers], ["other.example"])

    def test_legacy_rows_are_normalized(self) -> None:
        self.data.mkdir()
        (self.data / "traffic.csv").write_text(
            "date,views,unique_views,clones,unique_clones,stars,forks,watchers,open_issues\r\n"
            "2026-05-01T06:00:00+00:00,5,2,7,3,1,0,1,0\r\n", encoding="utf-8")
        self.build("2026-10-31")
        rows = traffic.read_csv(self.data / "traffic.csv")
        self.assertEqual([(r["snapshot_date"], r["views_14d"]) for r in rows],
                         [("2026-05-01", "5"), ("2026-10-31", "10")])

    def test_unexpected_api_response_fails(self) -> None:
        write_api(self.api)
        (self.api / "views.json").write_text("[]")
        with self.assertRaises(RuntimeError):
            traffic.build(self.api, self.data, datetime.now(UTC))

    def test_compact_count_and_to_int(self) -> None:
        self.assertEqual([traffic.compact_count(v) for v in (999, 1000, 4130, 2_000_000, 2_500_000)],
                         ["999", "1k", "4.1k", "2M", "2.5M"])
        self.assertEqual([traffic.to_int(v) for v in ("7", 7.9, "", None, "x", [1])], [7, 7, 0, 0, 0, 0])


class AlertTests(TrafficFixture):
    def test_no_alert_on_a_quiet_day_or_a_single_snapshot(self) -> None:
        self.build("2026-10-30")
        self.assertIsNone(traffic.alert(self.data))
        self.build("2026-10-31", views=12)
        self.assertIsNone(traffic.alert(self.data))

    def test_spike_stars_and_new_referrer_trigger_one_alert(self) -> None:
        self.build("2026-10-30")
        self.build("2026-10-31", views=500, stars=3,
                   referrers=[{"referrer": "github.com", "count": 7, "uniques": 2},
                              {"referrer": "news.example", "count": 40, "uniques": 30}])
        result = traffic.alert(self.data)
        assert result is not None
        title, body = result
        self.assertEqual(title, "Traffic alert: 2026-10-31")
        self.assertIn("Trigger reason:\nViews spike, Stars increased, New referrer detected\n", body)
        self.assertIn("New referrers:\n- news.example: 40 views, 30 unique\n", body)
        self.assertTrue(body.startswith("Automated traffic alert") and body.endswith("using the component"))

    @unittest.skipIf(os.name == "nt", "uses a POSIX shell script as a fake gh")
    def test_issue_is_created_once(self) -> None:
        self.build("2026-10-30")
        self.build("2026-10-31", stars=9)
        log = self.root / "gh.log"
        fake = self.root / "gh"
        fake.write_text(f'#!/bin/sh\necho "$2" >> "{log}"\n'
                        f'[ "$2" = list ] && [ -f "{self.root}/exists" ] && echo "#1 Traffic alert"\nexit 0\n')
        fake.chmod(fake.stat().st_mode | stat.S_IEXEC)
        self.assertEqual(traffic.open_alert_issue(self.data, str(fake)), 0)
        (self.root / "exists").touch()
        self.assertEqual(traffic.open_alert_issue(self.data, str(fake)), 0)
        calls = log.read_text().splitlines()
        self.assertEqual(calls, ["list", "create", "list"])


if __name__ == "__main__":
    unittest.main()
