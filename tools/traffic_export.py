#!/usr/bin/env python3
"""Turn GitHub traffic API snapshots into the traffic-history CSVs and badges,
and open an alert issue on a traffic spike or new referrer.

Used by the Daily traffic export workflow. It never calls the traffic API and
never sees the API token: the workflow fetches the JSON first, then runs
``build`` on the saved files and ``alert`` on the resulting history.
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import subprocess
import sys
from collections.abc import Callable, Sequence
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

Row = dict[str, Any]

TRAFFIC_FIELDS = ["snapshot_date", "snapshot_datetime_utc", "views_14d", "unique_views_14d", "clones_14d",
                  "unique_clones_14d", "stars", "forks", "watchers", "open_issues"]
DAILY_FIELDS = ["date", "views", "unique_views", "clones", "unique_clones"]
REFERRER_FIELDS = ["snapshot_date", "rank", "referrer", "count", "uniques"]
PATH_FIELDS = ["snapshot_date", "rank", "path", "title", "count", "uniques"]


def load_json(path: Path) -> Any:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def read_csv(path: Path) -> list[Row]:
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, rows: Sequence[Row], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    # csv's default dialect ends rows with CRLF; the history files have always used it.
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow({field: row.get(field, "") for field in fieldnames})


def row_key(row: Row, key_fields: list[str]) -> tuple[str, ...]:
    return tuple(str(row.get(key, "")) for key in key_fields)


def upsert_rows(path: Path, new_rows: list[Row], fieldnames: list[str], key_fields: list[str],
                normalizer: Callable[[Row], Row | None] | None = None) -> None:
    """Replace rows whose key matches a new row, keep the rest, drop rows with an
    empty key, and write the result sorted by key."""
    existing_rows = read_csv(path)
    if normalizer is not None:
        existing_rows = [n for n in (normalizer(row) for row in existing_rows) if n is not None]
    new_keys = {row_key(row, key_fields) for row in new_rows}
    combined = [row for row in existing_rows if row_key(row, key_fields) not in new_keys] + new_rows
    combined = [row for row in combined if all(str(row.get(key, "")) for key in key_fields)]
    combined.sort(key=lambda row: row_key(row, key_fields))
    write_csv(path, combined, fieldnames)


def to_int(value: object, default: int = 0) -> int:
    """int(value) for JSON and CSV values; anything else, or a failed conversion, gives ``default``."""
    if isinstance(value, (int, float, str, bytes)):
        try:
            return int(value)
        except (ValueError, OverflowError):
            return default
    return default


def normalize_traffic_row(row: Row) -> Row | None:
    """Accept the current snapshot layout and the original date/views layout."""
    if row.get("snapshot_date"):
        return {field: row.get(field, "" if field == "snapshot_datetime_utc" else "0") for field in TRAFFIC_FIELDS}
    if row.get("date"):
        return {
            "snapshot_date": row.get("date", "")[:10],
            "snapshot_datetime_utc": row.get("date", ""),
            "views_14d": row.get("views", "0"),
            "unique_views_14d": row.get("unique_views", "0"),
            "clones_14d": row.get("clones", "0"),
            "unique_clones_14d": row.get("unique_clones", "0"),
            "stars": row.get("stars", "0"),
            "forks": row.get("forks", "0"),
            "watchers": row.get("watchers", "0"),
            "open_issues": row.get("open_issues", "0"),
        }
    return None


def compact_count(value: int) -> str:
    if value >= 1000000:
        return f"{value / 1000000:.1f}M".replace(".0M", "M")
    if value >= 1000:
        return f"{value / 1000:.1f}k".replace(".0k", "k")
    return str(value)


def write_badge(data: Path, name: str, label: str, value: int) -> None:
    """Write one shields.io endpoint file."""
    (data / "badges").mkdir(parents=True, exist_ok=True)
    with open(data / "badges" / f"{name}.json", "w", encoding="utf-8") as f:
        json.dump({"schemaVersion": 1, "label": label, "message": compact_count(value), "color": "217346"}, f)
        f.write("\n")


def ranked_snapshot(path: Path, snapshot_date: str, items: list[Row], fields: list[str]) -> None:
    """Replace today's ranked rows (referrers or paths) and keep earlier days."""
    existing = [row for row in read_csv(path) if row.get("snapshot_date") != snapshot_date]
    rows: list[Row] = []
    for rank, item in enumerate(items, start=1):
        row: Row = {"snapshot_date": snapshot_date, "rank": rank}
        for field in fields[2:]:
            row[field] = to_int(item.get(field, 0)) if field in ("count", "uniques") else item.get(field, "")
        rows.append(row)
    combined = existing + rows
    combined.sort(key=lambda row: (row.get("snapshot_date", ""), to_int(row.get("rank", 999))))
    write_csv(path, combined, fields)


def build(api_dir: Path, data: Path, now: datetime) -> None:
    """Merge one set of API responses into the history under ``data``."""
    views = load_json(api_dir / "views.json")
    clones = load_json(api_dir / "clones.json")
    referrers = load_json(api_dir / "referrers.json")
    paths = load_json(api_dir / "paths.json")
    repo = load_json(api_dir / "repo.json")

    if not isinstance(views, dict) or "views" not in views or "count" not in views:
        raise RuntimeError("Unexpected views API response")
    if not isinstance(clones, dict) or "clones" not in clones or "count" not in clones:
        raise RuntimeError("Unexpected clones API response")
    if not isinstance(referrers, list):
        raise RuntimeError("Unexpected referrers API response")
    if not isinstance(paths, list):
        raise RuntimeError("Unexpected popular paths API response")
    if not isinstance(repo, dict) or "stargazers_count" not in repo:
        raise RuntimeError("Unexpected repository metadata API response")

    snapshot_date = now.date().isoformat()
    traffic_row: Row = {
        "snapshot_date": snapshot_date,
        "snapshot_datetime_utc": now.isoformat(timespec="seconds"),
        "views_14d": to_int(views.get("count", 0)),
        "unique_views_14d": to_int(views.get("uniques", 0)),
        "clones_14d": to_int(clones.get("count", 0)),
        "unique_clones_14d": to_int(clones.get("uniques", 0)),
        "stars": to_int(repo.get("stargazers_count", 0)),
        "forks": to_int(repo.get("forks_count", 0)),
        "watchers": to_int(repo.get("subscribers_count", 0)),
        "open_issues": to_int(repo.get("open_issues_count", 0)),
    }
    upsert_rows(data / "traffic.csv", [traffic_row], TRAFFIC_FIELDS, ["snapshot_date"], normalize_traffic_row)

    daily: dict[str, Row] = {}
    for source, count_field, unique_field in ((views, "views", "unique_views"), (clones, "clones", "unique_clones")):
        for item in source.get(count_field, []):
            day = str(item.get("timestamp", ""))[:10]
            if not day:
                continue
            entry = daily.setdefault(day, {"date": day, "views": 0, "unique_views": 0, "clones": 0, "unique_clones": 0})
            entry[count_field] = to_int(item.get("count", 0))
            entry[unique_field] = to_int(item.get("uniques", 0))
    upsert_rows(data / "traffic_daily.csv", list(daily.values()), DAILY_FIELDS, ["date"])

    # README badges:
    #   clones   total clones recorded in the daily history since tracking began
    #   visitors unique visitors over GitHub's rolling 14-day window, from
    #            today's snapshot (daily uniques cannot be summed into a unique count)
    #   visits   cumulative daily unique visitors since tracking began; a visitor
    #            counts once per day they visit
    history = read_csv(data / "traffic_daily.csv")
    write_badge(data, "clones", "clones", sum(to_int(row.get("clones", 0)) for row in history))
    write_badge(data, "visits", "visits", sum(to_int(row.get("unique_views", 0)) for row in history))
    write_badge(data, "visitors", "visitors (14d)", to_int(traffic_row["unique_views_14d"]))

    ranked_snapshot(data / "referrers.csv", snapshot_date, referrers, REFERRER_FIELDS)
    ranked_snapshot(data / "popular_paths.csv", snapshot_date, paths, PATH_FIELDS)


def top_by_count(rows: list[Row], limit: int = 10) -> list[Row]:
    return sorted(rows, key=lambda row: to_int(row.get("count", 0)), reverse=True)[:limit]


def alert(data: Path) -> tuple[str, str] | None:
    """Return the (title, body) of today's alert issue, or None when nothing triggered."""
    traffic = [row for row in read_csv(data / "traffic.csv") if row.get("snapshot_date")]
    traffic.sort(key=lambda row: row.get("snapshot_date", ""))
    if len(traffic) < 2:
        return None
    previous, current = traffic[-2], traffic[-1]
    current_date, previous_date = current["snapshot_date"], previous["snapshot_date"]

    views_spike = to_int(current["views_14d"]) >= max(to_int(previous["views_14d"]) * 2,
                                                      to_int(previous["views_14d"]) + 100)
    clones_spike = to_int(current["clones_14d"]) >= max(to_int(previous["clones_14d"]) * 2,
                                                        to_int(previous["clones_14d"]) + 10)
    stars_up = to_int(current["stars"]) > to_int(previous["stars"])
    forks_up = to_int(current["forks"]) > to_int(previous["forks"])

    referrers = read_csv(data / "referrers.csv")
    paths = read_csv(data / "popular_paths.csv")
    current_referrers = [row for row in referrers if row.get("snapshot_date") == current_date]
    current_paths = [row for row in paths if row.get("snapshot_date") == current_date]
    known_before = {row.get("referrer", "") for row in referrers if row.get("snapshot_date", "") < current_date}
    new_referrers = [row for row in current_referrers
                     if row.get("referrer", "") and row.get("referrer", "") not in known_before
                     and to_int(row.get("count", 0)) >= 5]

    reasons = [reason for reason, triggered in (("Views spike", views_spike), ("Clones spike", clones_spike),
                                                ("Stars increased", stars_up), ("Forks increased", forks_up),
                                                ("New referrer detected", bool(new_referrers))) if triggered]
    if not reasons:
        return None

    def lines(rows: list[Row], key: str, empty: str) -> str:
        return "\n".join(f"- {row.get(key, '')}: {row.get('count', '0')} views, {row.get('uniques', '0')} unique"
                         for row in top_by_count(rows)) or empty

    referrer_lines = lines(current_referrers, "referrer", "- No referrers available")
    path_lines = lines(current_paths, "path", "- No popular paths available")
    new_referrer_lines = lines(new_referrers, "referrer", "- No new referrer above threshold")

    body = "\n".join([
        "Automated traffic alert for **VBA-DateTimePicker**.",
        "",
        "Trigger reason:",
        ", ".join(reasons),
        "",
        "Current 14-day snapshot:",
        f"- Views: {current['views_14d']} vs {previous['views_14d']} on {previous_date}",
        f"- Unique views: {current['unique_views_14d']} vs {previous['unique_views_14d']} on {previous_date}",
        f"- Clones: {current['clones_14d']} vs {previous['clones_14d']} on {previous_date}",
        f"- Unique clones: {current['unique_clones_14d']} vs {previous['unique_clones_14d']} on {previous_date}",
        f"- Stars: {current['stars']} vs {previous['stars']} on {previous_date}",
        f"- Forks: {current['forks']} vs {previous['forks']} on {previous_date}",
        "",
        "New referrers:",
        new_referrer_lines,
        "",
        "Current top referrers:",
        referrer_lines,
        "",
        "Current popular paths:",
        path_lines,
        "",
        "Suggested checks:",
        "1. Go to **Insights → Traffic**",
        "2. Check whether one of the referrers is a blog, forum, Reddit, StackOverflow, GitHub project, or newsletter",
        "3. Check which page, release, screenshot, source file, or README section is attracting traffic",
        "4. Consider adding a short README CTA asking users to star the repository if they are using the component",
    ])
    return f"Traffic alert: {current_date}", body


def open_alert_issue(data: Path, gh: str = "gh") -> int:
    """Create today's alert issue with the gh CLI unless one already exists."""
    result = alert(data)
    if result is None:
        return 0
    title, body = result
    existing = subprocess.run([gh, "issue", "list", "--state", "all", "--search", f'"{title}" in:title',
                               "--limit", "10"], capture_output=True, text=True, check=False)
    if existing.returncode != 0:
        print("::error::Could not determine whether an alert issue already exists")
        print(existing.stderr.strip())
        return 1
    if existing.stdout.strip():
        print(f"Alert issue already exists for {title.removeprefix('Traffic alert: ')}")
        return 0
    created = subprocess.run([gh, "issue", "create", "--title", title, "--body", body],
                             capture_output=True, text=True, check=False)
    if created.returncode != 0:
        print("::error::Traffic alert triggered but the issue could not be created")
        print(created.stderr.strip())
        return 1
    print(created.stdout.strip())
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    commands = parser.add_subparsers(dest="command", required=True)
    build_parser = commands.add_parser("build", help="Merge saved API responses into the history")
    build_parser.add_argument("--api-dir", type=Path, required=True,
                              help="Folder with views.json, clones.json, referrers.json, paths.json, repo.json")
    build_parser.add_argument("--data-dir", type=Path, required=True)
    build_parser.add_argument("--now", help="ISO 8601 UTC timestamp to record instead of the current time")
    alert_parser = commands.add_parser("alert", help="Open an alert issue if the latest snapshot warrants one")
    alert_parser.add_argument("--data-dir", type=Path, required=True)
    args = parser.parse_args(argv)
    if args.command == "build":
        now = datetime.fromisoformat(args.now) if args.now else datetime.now(UTC)
        build(args.api_dir, args.data_dir, now)
        return 0
    return open_alert_issue(args.data_dir, os.environ.get("GH_CLI", "gh"))


if __name__ == "__main__":
    sys.exit(main())
