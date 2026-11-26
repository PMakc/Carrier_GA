# -*- coding: utf-8 -*-
"""VK Career Day bot, prepared for Bothost deployment.

Secrets and runtime settings are read from environment variables.
SQLite database is stored in /app/data on Bothost by default.
"""

from __future__ import annotations

import html
import io
import json
import logging
import os
import random
import re
import secrets
import sqlite3
import time
import urllib.error
import urllib.parse
import urllib.request
import http.client
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

def _env_int(name: str, default: int) -> int:
    value = os.getenv(name, str(default)).strip()
    try:
        return int(value)
    except ValueError as exc:
        raise RuntimeError(f"Переменная {name} должна быть целым числом") from exc


def _parse_admin_ids() -> set[int]:
    raw = os.getenv("ADMIN_IDS", "")
    ids: set[int] = set()
    for item in re.split(r"[,;\s]+", raw.strip()):
        if not item:
            continue
        try:
            ids.add(int(item))
        except ValueError as exc:
            raise RuntimeError("ADMIN_IDS должен содержать VK ID через запятую") from exc
    return ids


VK_TOKEN = os.getenv("VK_TOKEN", "").strip()
GROUP_ID = _env_int("VK_GROUP_ID", 241674433)
ADMIN_IDS = _parse_admin_ids()
ORGANIZER_PASSWORD = os.getenv("ORGANIZER_PASSWORD", "")
API_VERSION = os.getenv("VK_API_VERSION", "5.199")
MIN_TIME_BETWEEN_STATIONS = _env_int("MIN_TIME_BETWEEN_STATIONS", 180)
ANTI_SPAM_COOLDOWN = _env_int("ANTI_SPAM_COOLDOWN", 2)

DATA_DIR = Path(os.getenv("DATA_DIR", "/app/data"))
if os.getenv("DATABASE_PATH"):
    DB_PATH = Path(os.environ["DATABASE_PATH"])
elif os.getenv("DB_PATH"):
    DB_PATH = Path(os.environ["DB_PATH"])
else:
    # Bothost injects DATA_DIR=/app/data; use it even if the folder is created at startup.
    DB_PATH = (DATA_DIR / "career_bot_vk.db") if (os.getenv("DATA_DIR") or DATA_DIR.exists()) else Path("career_bot_vk.db")

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("career_day_vk_bot")

# Стартовый маршрут мероприятия. Кодовые слова можно изменить позже
# через удаление станции и добавление её заново в панели администратора.
BASE_STATIONS: dict[int, dict[str, str]] = {
    1: {"name": "АО «Разрез Харанорский»", "code_word": "ХАРАНОР"},
    2: {"name": "Филиал «Харанорская ГРЭС»", "code_word": "ГРЭС"},
    3: {"name": "ПАО «ВымпелКом»", "code_word": "БИЛАЙН"},
    4: {"name": "ГРК «Быстринское»", "code_word": "БЫСТРИНСКОЕ"},
    5: {"name": "ООО «Удоканская медь»", "code_word": "УДОКАН"},
    6: {"name": "ООО «Мангазея Майнинг»", "code_word": "МАНГАЗЕЯ"},
    7: {"name": "ООО «Золото Дельмачик»", "code_word": "ДЕЛЬМАЧИК"},
    8: {"name": "ПАО «Ростелеком»", "code_word": "РОСТЕЛЕКОМ"},
    9: {"name": "ПАО «Россети»", "code_word": "РОССЕТИ"},
    10: {"name": "ПАО «Полюс»", "code_word": "ПОЛЮС"},
    11: {"name": "АО «Хиагда»", "code_word": "ХИАГДА"},
    12: {"name": "МКАО «АРЕАЛ»", "code_word": "АРЕАЛ"},
    13: {"name": "ПАО Сбербанк", "code_word": "СБЕРБАНК"},
    14: {"name": "АО «Азот-Взрыв»", "code_word": "АЗОТВЗРЫВ"},
}

# Предыдущий стартовый набор. При первом запуске обновлённой версии эти
# станции архивируются, а новые добавляются с новыми ID; посещения остаются в БД,
# но не засчитываются в новом маршруте.
LEGACY_DEFAULT_STATIONS: set[tuple[str, str]] = {
    ("ППГХО им. Е.П. Славского", "УРАН"),
    ("Хиагда", "УРАН2"),
    ("ГРК Быстринское", "СКОРОСТЬ"),
    ("Мангазея Майнинг", "СЕРЕБРО"),
    ("Разрез Харанорский", "УГОЛЬ"),
    ("РазрезУголь", "УГОЛЬ2"),
    ("Удоканская медь", "МЕДЬ"),
    ("Связь РЖД", "СВЯЗЬ"),
    ("Ростелеком", "ИНТЕРНЕТ"),
    ("Забайкальский Информационный центр", "ЦИФРА"),
    ("Аэронавигация Восточной Сибири", "НЕБО"),
    ("Золото Дельмачик", "ЗОЛОТО"),
    ("РОССЕТИ СИБИРЬ – ЧИТАЭНЕРГО", "ЭНЕРГИЯ"),
}
STATION_SEED_VERSION = "career-day-2026-14-stations-v1"


# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------

class Database:
    def __init__(self, db_path: Path | str = DB_PATH):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.init_db()

    @contextmanager
    def connect(self):
        conn = sqlite3.connect(str(self.db_path), timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA busy_timeout = 30000")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def init_db(self) -> None:
        with self.connect() as conn:
            # Detect the original BOT.py database layout before creating the
            # stations table. Its visits table already used station IDs 1..N,
            # so the new station IDs must not reuse those values.
            had_stations_table = conn.execute(
                "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'stations'"
            ).fetchone() is not None
            legacy_visit_max_id = 0
            if not had_stations_table:
                has_visits_table = conn.execute(
                    "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'visits'"
                ).fetchone() is not None
                if has_visits_table:
                    columns = {row[1] for row in conn.execute("PRAGMA table_info(visits)").fetchall()}
                    if "station_id" in columns:
                        legacy_visit_max_id = int(
                            conn.execute("SELECT COALESCE(MAX(station_id), 0) FROM visits").fetchone()[0]
                        )

            conn.execute("PRAGMA journal_mode = WAL")
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS users (
                    user_id INTEGER PRIMARY KEY,
                    username TEXT,
                    first_name TEXT,
                    last_name TEXT,
                    user_fio TEXT,
                    user_role TEXT DEFAULT 'participant',
                    registered_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS stations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    code_word TEXT NOT NULL,
                    normalized_code TEXT NOT NULL UNIQUE,
                    is_active INTEGER NOT NULL DEFAULT 1,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    deactivated_at TEXT
                );

                CREATE TABLE IF NOT EXISTS visits (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    station_id INTEGER NOT NULL,
                    station_name TEXT NOT NULL,
                    visited_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(user_id, station_id)
                );

                CREATE TABLE IF NOT EXISTS user_cooldown (
                    user_id INTEGER PRIMARY KEY,
                    last_visit_time TEXT
                );

                CREATE TABLE IF NOT EXISTS spam_protection (
                    user_id INTEGER PRIMARY KEY,
                    last_message_time TEXT,
                    warning_count INTEGER DEFAULT 0,
                    last_warning_time TEXT
                );

                CREATE TABLE IF NOT EXISTS raffle_draws (
                    user_id INTEGER PRIMARY KEY,
                    organizer_id INTEGER NOT NULL,
                    drawn_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS app_settings (
                    setting_key TEXT PRIMARY KEY,
                    setting_value TEXT NOT NULL
                );
                """
            )

            seeded_version = conn.execute(
                "SELECT setting_value FROM app_settings WHERE setting_key = 'station_seed_version'"
            ).fetchone()
            if not seeded_version or seeded_version[0] != STATION_SEED_VERSION:
                # Archive only the previous built-in stations. Any extra
                # stations intentionally created by an admin remain active.
                old_rows = conn.execute(
                    "SELECT id, name, normalized_code FROM stations WHERE is_active = 1"
                ).fetchall()
                for old_row in old_rows:
                    legacy_key = (str(old_row[1]), str(old_row[2]).strip().upper())
                    if legacy_key in LEGACY_DEFAULT_STATIONS:
                        conn.execute(
                            "UPDATE stations SET is_active = 0, deactivated_at = ? WHERE id = ?",
                            (datetime.now().isoformat(timespec="seconds"), int(old_row[0])),
                        )

                for index, station in BASE_STATIONS.items():
                    code = station["code_word"].strip().upper()
                    exists = conn.execute(
                        "SELECT id FROM stations WHERE normalized_code = ?", (code,)
                    ).fetchone()
                    if exists:
                        # Do not silently overwrite an admin-created station
                        # that already owns this code word.
                        if conn.execute(
                            "SELECT 1 FROM stations WHERE id = ? AND name = ?",
                            (int(exists[0]), station["name"]),
                        ).fetchone():
                            conn.execute(
                                "UPDATE stations SET is_active = 1, deactivated_at = NULL WHERE id = ?",
                                (int(exists[0]),),
                            )
                        else:
                            logger.warning(
                                "Не удалось добавить стартовую станцию %s: код %s уже занят другой станцией",
                                station["name"], code,
                            )
                        continue

                    # For a legacy BOT.py database without a stations table,
                    # avoid reusing IDs that may already exist in visits.
                    preferred_id = index + legacy_visit_max_id if not had_stations_table else None
                    if preferred_id is not None:
                        conn.execute(
                            """INSERT INTO stations
                               (id, name, code_word, normalized_code, is_active)
                               VALUES (?, ?, ?, ?, 1)""",
                            (preferred_id, station["name"], code, code),
                        )
                    else:
                        conn.execute(
                            """INSERT INTO stations
                               (name, code_word, normalized_code, is_active)
                               VALUES (?, ?, ?, 1)""",
                            (station["name"], code, code),
                        )

                conn.execute(
                    "INSERT INTO app_settings (setting_key, setting_value) VALUES ('station_seed_version', ?) "
                    "ON CONFLICT(setting_key) DO UPDATE SET setting_value = excluded.setting_value",
                    (STATION_SEED_VERSION,),
                )
            conn.commit()

    @staticmethod
    def _now() -> str:
        return datetime.now().isoformat(timespec="seconds")

    def add_user(self, user_id: int, username: str | None, first_name: str | None,
                 last_name: str | None) -> None:
        with self.connect() as conn:
            conn.execute(
                """INSERT INTO users (user_id, username, first_name, last_name)
                   VALUES (?, ?, ?, ?)
                   ON CONFLICT(user_id) DO UPDATE SET
                       username = COALESCE(excluded.username, users.username),
                       first_name = COALESCE(excluded.first_name, users.first_name),
                       last_name = COALESCE(excluded.last_name, users.last_name)""",
                (user_id, username, first_name, last_name),
            )

    def update_fio(self, user_id: int, fio: str) -> None:
        with self.connect() as conn:
            conn.execute("UPDATE users SET user_fio = ? WHERE user_id = ?", (fio, user_id))

    def update_role(self, user_id: int, role: str) -> None:
        with self.connect() as conn:
            conn.execute("UPDATE users SET user_role = ? WHERE user_id = ?", (role, user_id))

    def get_fio(self, user_id: int) -> str | None:
        with self.connect() as conn:
            row = conn.execute("SELECT user_fio FROM users WHERE user_id = ?", (user_id,)).fetchone()
            return row[0] if row else None

    def get_role(self, user_id: int) -> str:
        with self.connect() as conn:
            row = conn.execute("SELECT user_role FROM users WHERE user_id = ?", (user_id,)).fetchone()
            return row[0] if row and row[0] else "participant"

    # -- Anti-spam and station cooldown ------------------------------------
    def check_spam_protection(self, user_id: int) -> tuple[bool, str | None]:
        now = datetime.now()
        now_s = now.isoformat(timespec="seconds")
        with self.connect() as conn:
            row = conn.execute(
                "SELECT last_message_time, warning_count, last_warning_time "
                "FROM spam_protection WHERE user_id = ?", (user_id,)
            ).fetchone()
            if not row or not row[0]:
                conn.execute(
                    "INSERT OR REPLACE INTO spam_protection "
                    "(user_id, last_message_time, warning_count, last_warning_time) VALUES (?, ?, 0, NULL)",
                    (user_id, now_s),
                )
                return True, None

            try:
                last_message = datetime.fromisoformat(row[0])
            except (TypeError, ValueError):
                last_message = now
            elapsed = (now - last_message).total_seconds()
            warnings = int(row[1] or 0)
            if elapsed < ANTI_SPAM_COOLDOWN:
                last_warning = None
                if row[2]:
                    try:
                        last_warning = datetime.fromisoformat(row[2])
                    except ValueError:
                        pass
                show_warning = last_warning is None or (now - last_warning).total_seconds() >= 5
                message = None
                if show_warning:
                    warnings += 1
                    conn.execute(
                        "UPDATE spam_protection SET warning_count = ?, last_warning_time = ? WHERE user_id = ?",
                        (warnings, now_s, user_id),
                    )
                    left = max(0.1, ANTI_SPAM_COOLDOWN - elapsed)
                    if warnings >= 5:
                        message = f"🚫 Слишком много запросов. Подождите {left:.1f} сек."
                    elif warnings >= 3:
                        message = f"⚠️ Не отправляйте сообщения так часто. Подождите {left:.1f} сек."
                    else:
                        message = f"⏱ Подождите {left:.1f} сек. перед следующим сообщением."
                return False, message

            if elapsed > ANTI_SPAM_COOLDOWN * 2:
                warnings = 0
            conn.execute(
                "UPDATE spam_protection SET last_message_time = ?, warning_count = ? WHERE user_id = ?",
                (now_s, warnings, user_id),
            )
            return True, None

    def can_visit_station(self, user_id: int) -> tuple[bool, float]:
        with self.connect() as conn:
            row = conn.execute(
                "SELECT last_visit_time FROM user_cooldown WHERE user_id = ?", (user_id,)
            ).fetchone()
        if not row or not row[0]:
            return True, 0
        try:
            last_visit = datetime.fromisoformat(row[0])
        except ValueError:
            return True, 0
        elapsed = (datetime.now() - last_visit).total_seconds()
        remaining = MIN_TIME_BETWEEN_STATIONS - elapsed
        return (remaining <= 0, max(0, remaining))

    # -- Stations -----------------------------------------------------------
    def get_active_stations(self) -> list[dict[str, Any]]:
        with self.connect() as conn:
            rows = conn.execute(
                "SELECT id, name, code_word FROM stations WHERE is_active = 1 ORDER BY id"
            ).fetchall()
            return [dict(row) for row in rows]

    def get_station_by_code(self, code_word: str) -> dict[str, Any] | None:
        normalized = code_word.strip().upper()
        with self.connect() as conn:
            row = conn.execute(
                "SELECT id, name, code_word FROM stations "
                "WHERE normalized_code = ? AND is_active = 1", (normalized,)
            ).fetchone()
            return dict(row) if row else None

    def add_station(self, name: str, code_word: str) -> int:
        normalized = code_word.strip().upper()
        with self.connect() as conn:
            cursor = conn.execute(
                "INSERT INTO stations (name, code_word, normalized_code, is_active) VALUES (?, ?, ?, 1)",
                (name.strip(), normalized, normalized),
            )
            return int(cursor.lastrowid)

    def deactivate_station(self, station_id: int) -> bool:
        with self.connect() as conn:
            active_count = conn.execute(
                "SELECT COUNT(*) FROM stations WHERE is_active = 1"
            ).fetchone()[0]
            if active_count <= 1:
                return False
            cursor = conn.execute(
                "UPDATE stations SET is_active = 0, deactivated_at = ? "
                "WHERE id = ? AND is_active = 1", (self._now(), station_id)
            )
            return cursor.rowcount > 0

    def record_visit(self, user_id: int, station_id: int, station_name: str) -> bool:
        now = self._now()
        with self.connect() as conn:
            station = conn.execute(
                "SELECT name FROM stations WHERE id = ? AND is_active = 1", (station_id,)
            ).fetchone()
            if not station:
                return False
            try:
                conn.execute(
                    "INSERT INTO visits (user_id, station_id, station_name, visited_at) VALUES (?, ?, ?, ?)",
                    (user_id, station_id, station[0] or station_name, now),
                )
            except sqlite3.IntegrityError:
                return False
            conn.execute(
                "INSERT INTO user_cooldown (user_id, last_visit_time) VALUES (?, ?) "
                "ON CONFLICT(user_id) DO UPDATE SET last_visit_time = excluded.last_visit_time",
                (user_id, now),
            )
            return True

    def get_user_stats(self, user_id: int) -> dict[str, int | bool]:
        with self.connect() as conn:
            total = int(conn.execute(
                "SELECT COUNT(*) FROM stations WHERE is_active = 1"
            ).fetchone()[0])
            visited = int(conn.execute(
                """SELECT COUNT(*) FROM visits v
                   JOIN stations s ON s.id = v.station_id
                   WHERE v.user_id = ? AND s.is_active = 1""", (user_id,)
            ).fetchone()[0])
        return {
            "visited_count": visited,
            "total_stations": total,
            "remaining": max(0, total - visited),
            "completed": total > 0 and visited >= total,
        }

    def get_visits(self, user_id: int) -> list[tuple[str, str]]:
        with self.connect() as conn:
            rows = conn.execute(
                """SELECT v.station_name, v.visited_at FROM visits v
                   JOIN stations s ON s.id = v.station_id
                   WHERE v.user_id = ? AND s.is_active = 1
                   ORDER BY v.visited_at""", (user_id,)
            ).fetchall()
            return [(row[0], row[1]) for row in rows]

    def get_all_participants(self) -> list[tuple[int, str, str | None]]:
        with self.connect() as conn:
            rows = conn.execute(
                "SELECT user_id, user_fio, username FROM users "
                "WHERE user_role = 'participant' AND user_fio IS NOT NULL ORDER BY user_fio"
            ).fetchall()
            return [(int(r[0]), r[1], r[2]) for r in rows]

    def get_all_users_with_progress(self) -> list[tuple[int, str, str | None, int, int]]:
        with self.connect() as conn:
            rows = conn.execute(
                """SELECT u.user_id, u.user_fio, u.username,
                          COUNT(DISTINCT CASE WHEN s.is_active = 1 THEN v.station_id END) AS visited_count,
                          (SELECT COUNT(*) FROM stations WHERE is_active = 1) AS total_count
                   FROM users u
                   LEFT JOIN visits v ON v.user_id = u.user_id
                   LEFT JOIN stations s ON s.id = v.station_id
                   WHERE u.user_fio IS NOT NULL
                   GROUP BY u.user_id
                   ORDER BY visited_count DESC, u.user_fio"""
            ).fetchall()
            return [(int(r[0]), r[1], r[2], int(r[3]), int(r[4])) for r in rows]

    def get_winners(self) -> list[tuple[int, str, str | None]]:
        with self.connect() as conn:
            total = int(conn.execute(
                "SELECT COUNT(*) FROM stations WHERE is_active = 1"
            ).fetchone()[0])
            if total == 0:
                return []
            rows = conn.execute(
                """SELECT u.user_id, u.user_fio, u.username
                   FROM users u
                   WHERE u.user_fio IS NOT NULL AND
                     (SELECT COUNT(DISTINCT v.station_id)
                      FROM visits v JOIN stations s ON s.id = v.station_id
                      WHERE v.user_id = u.user_id AND s.is_active = 1) >= ?
                   ORDER BY u.user_fio""", (total,)
            ).fetchall()
            return [(int(r[0]), r[1], r[2]) for r in rows]

    def get_drawable_winners(self) -> list[tuple[int, str, str | None]]:
        with self.connect() as conn:
            total = int(conn.execute(
                "SELECT COUNT(*) FROM stations WHERE is_active = 1"
            ).fetchone()[0])
            if total == 0:
                return []
            rows = conn.execute(
                """SELECT u.user_id, u.user_fio, u.username
                   FROM users u
                   WHERE u.user_fio IS NOT NULL
                     AND (SELECT COUNT(DISTINCT v.station_id)
                          FROM visits v JOIN stations s ON s.id = v.station_id
                          WHERE v.user_id = u.user_id AND s.is_active = 1) >= ?
                     AND NOT EXISTS (SELECT 1 FROM raffle_draws r WHERE r.user_id = u.user_id)
                   ORDER BY u.user_fio""", (total,)
            ).fetchall()
            return [(int(r[0]), r[1], r[2]) for r in rows]

    def record_raffle_winner(self, user_id: int, organizer_id: int) -> None:
        with self.connect() as conn:
            conn.execute(
                "INSERT OR IGNORE INTO raffle_draws (user_id, organizer_id, drawn_at) VALUES (?, ?, ?)",
                (user_id, organizer_id, self._now()),
            )

    def reset_raffle(self) -> int:
        with self.connect() as conn:
            count = int(conn.execute("SELECT COUNT(*) FROM raffle_draws").fetchone()[0])
            conn.execute("DELETE FROM raffle_draws")
            return count

    def get_all_stats(self) -> dict[str, int]:
        with self.connect() as conn:
            total_users = int(conn.execute(
                "SELECT COUNT(*) FROM users WHERE user_fio IS NOT NULL"
            ).fetchone()[0])
            total_visits = int(conn.execute(
                """SELECT COUNT(*) FROM visits v JOIN stations s ON s.id = v.station_id
                   WHERE s.is_active = 1"""
            ).fetchone()[0])
        return {
            "total_users": total_users,
            "total_visits": total_visits,
            "completed_users": len(self.get_winners()),
            "total_stations": len(self.get_active_stations()),
        }

    def export_dataset(self) -> dict[str, Any]:
        stations = self.get_active_stations()
        with self.connect() as conn:
            users = conn.execute(
                """SELECT user_id, user_fio, username, registered_at
                   FROM users WHERE user_fio IS NOT NULL ORDER BY user_fio"""
            ).fetchall()
            visit_rows = conn.execute(
                """SELECT v.user_id, v.station_id, v.station_name, v.visited_at
                   FROM visits v JOIN stations s ON s.id = v.station_id
                   WHERE s.is_active = 1 ORDER BY v.visited_at"""
            ).fetchall()
        visits_by_user: dict[int, dict[int, str]] = {}
        visit_details: list[dict[str, Any]] = []
        for row in visit_rows:
            uid = int(row[0])
            sid = int(row[1])
            visits_by_user.setdefault(uid, {})[sid] = str(row[3])
            visit_details.append({
                "user_id": uid,
                "station_id": sid,
                "station_name": row[2],
                "visited_at": str(row[3]),
            })
        participant_rows = []
        for row in users:
            uid = int(row[0])
            user_visits = visits_by_user.get(uid, {})
            participant_rows.append({
                "user_id": uid,
                "fio": row[1],
                "username": row[2],
                "registered_at": str(row[3] or ""),
                "visited_count": len(user_visits),
                "total_stations": len(stations),
                "visits": user_visits,
            })
        return {"stations": stations, "participants": participant_rows, "visits": visit_details}

    def complete_all_stations_for_user(self, user_id: int) -> int:
        stations = self.get_active_stations()
        now = self._now()
        added = 0
        with self.connect() as conn:
            for station in stations:
                cursor = conn.execute(
                    """INSERT OR IGNORE INTO visits (user_id, station_id, station_name, visited_at)
                       VALUES (?, ?, ?, ?)""",
                    (user_id, station["id"], station["name"], now),
                )
                added += cursor.rowcount
            if added:
                conn.execute(
                    "INSERT INTO user_cooldown (user_id, last_visit_time) VALUES (?, ?) "
                    "ON CONFLICT(user_id) DO UPDATE SET last_visit_time = excluded.last_visit_time",
                    (user_id, now),
                )
        return added


# ---------------------------------------------------------------------------
# VK API client (Long Poll)
# ---------------------------------------------------------------------------

class VkBot:
    def __init__(self, token: str, group_id: int):
        self.token = token
        self.group_id = group_id
        self.api_version = API_VERSION
        self.lp_server: str | None = None
        self.lp_key: str | None = None
        self.lp_ts: str | None = None

    @staticmethod
    def _https_request(host: str, path: str, params: dict[str, Any], method: str = "GET") -> Any:
        encoded = urllib.parse.urlencode(params)
        conn = http.client.HTTPSConnection(host, timeout=40)
        try:
            if method.upper() == "GET":
                request_path = f"{path}?{encoded}" if encoded else path
                conn.request("GET", request_path)
            else:
                conn.request(
                    "POST", path, encoded,
                    {"Content-Type": "application/x-www-form-urlencoded; charset=utf-8"},
                )
            response = conn.getresponse()
            raw = response.read().decode("utf-8", errors="replace")
            if response.status >= 400:
                raise RuntimeError(f"HTTP {response.status}: {raw[:300]}")
            return json.loads(raw) if raw else None
        finally:
            conn.close()

    def call(self, method: str, **params: Any) -> Any:
        payload = dict(params)
        payload["access_token"] = self.token
        payload["v"] = self.api_version
        try:
            result = self._https_request("api.vk.com", f"/method/{method}", payload, "POST")
        except Exception:
            logger.exception("Ошибка сетевого запроса VK API: %s", method)
            return None
        if not result:
            logger.error("VK API %s вернул пустой ответ", method)
            return None
        if "error" in result:
            error = result["error"]
            logger.error("VK API %s: код %s, %s", method,
                         error.get("error_code"), error.get("error_msg"))
            return None
        return result.get("response")

    def send_message(self, user_id: int, text: str,
                     reply_markup: dict[str, Any] | None = None,
                     attachment: str | None = None) -> int | None:
        # VK messages do not render HTML; remove any leftover formatting tags.
        text = html.unescape(re.sub(r"<[^>]*>", "", str(text)))
        chunks = [text[i:i + 3500] for i in range(0, len(text), 3500)] or [" "]
        first_message_id: int | None = None
        for index, chunk in enumerate(chunks):
            params: dict[str, Any] = {
                "peer_id": user_id,
                "message": chunk,
                "random_id": secrets.randbelow(2**31 - 1) + 1,
            }
            if attachment and index == 0:
                params["attachment"] = attachment
            if reply_markup and index == len(chunks) - 1:
                params["keyboard"] = json.dumps(reply_markup, ensure_ascii=False)
            result = self.call("messages.send", **params)
            if index == 0 and isinstance(result, int):
                first_message_id = result
        return first_message_id

    def edit_message(self, peer_id: int, message_id: int, text: str) -> bool:
        text = html.unescape(re.sub(r"<[^>]*>", "", str(text)))
        result = self.call("messages.edit", peer_id=peer_id, message_id=message_id, message=text)
        return bool(result)

    def get_user(self, user_id: int) -> tuple[str | None, str | None, str | None]:
        result = self.call("users.get", user_ids=user_id, fields="screen_name")
        if result and isinstance(result, list):
            user = result[0]
            return user.get("screen_name"), user.get("first_name"), user.get("last_name")
        return None, None, None

    def send_document(self, peer_id: int, filename: str, content: bytes,
                      caption: str = "") -> bool:
        """Upload an in-memory file to VK documents and send it to a private chat."""
        upload = self.call("docs.getMessagesUploadServer", peer_id=peer_id, type="doc")
        if not upload or not upload.get("upload_url"):
            logger.error("VK did not return a document upload URL")
            return False
        boundary = "----VKBotBoundary" + secrets.token_hex(12)
        body = (
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
        ).encode("utf-8") + content + f"\r\n--{boundary}--\r\n".encode("utf-8")
        request = urllib.request.Request(
            upload["upload_url"], data=body,
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                upload_result = json.loads(response.read().decode("utf-8"))
        except (urllib.error.URLError, TimeoutError, ValueError) as exc:
            logger.error("Не удалось загрузить Excel в VK: %s", exc)
            return False

        saved = self.call("docs.save", file=upload_result.get("file"), title=filename)
        if not saved:
            return False
        document = saved[0] if isinstance(saved, list) and saved else saved
        if not isinstance(document, dict) or "owner_id" not in document or "id" not in document:
            logger.error("Неожиданный ответ docs.save")
            return False
        attachment = f"doc{document['owner_id']}_{document['id']}"
        if document.get("access_key"):
            attachment += f"_{document['access_key']}"
        self.send_message(peer_id, caption or "📊 Excel-отчёт готов.", attachment=attachment)
        return True

    def _init_longpoll(self, keep_ts: bool = False) -> bool:
        previous_ts = self.lp_ts
        response = self.call("groups.getLongPollServer", group_id=self.group_id)
        if not response:
            logger.error("Не удалось получить параметры VK Long Poll")
            return False
        self.lp_server = response["server"]
        self.lp_key = response["key"]
        self.lp_ts = previous_ts if keep_ts and previous_ts else response["ts"]
        return True

    def get_updates(self) -> list[dict[str, Any]]:
        try:
            if not self.lp_server and not self._init_longpoll():
                time.sleep(4)
                return []
            parsed = urllib.parse.urlparse(self.lp_server or "")
            result = self._https_request(
                parsed.netloc, parsed.path,
                {"act": "a_check", "key": self.lp_key, "ts": self.lp_ts, "wait": 25},
                "GET",
            )
            if not result:
                return []
            if "failed" in result:
                failed = int(result["failed"])
                if failed == 1:
                    self.lp_ts = str(result.get("ts", self.lp_ts))
                elif failed == 2:
                    self._init_longpoll(keep_ts=True)
                else:
                    # failed=3 means events were lost: start from the new timestamp.
                    self._init_longpoll(keep_ts=False)
                return []
            self.lp_ts = str(result.get("ts", self.lp_ts))
            return result.get("updates", []) or []
        except Exception:
            logger.exception("Ошибка получения событий Long Poll")
            self.lp_server = self.lp_key = self.lp_ts = None
            time.sleep(3)
            return []


# ---------------------------------------------------------------------------
# Bot interaction and business logic
# ---------------------------------------------------------------------------

class BotHandler:
    def __init__(self, db: Database | None = None, bot: VkBot | None = None):
        self.db = db or Database()
        self.bot = bot or VkBot(VK_TOKEN, GROUP_ID)
        self.user_states: dict[int, str] = {}

    @staticmethod
    def _button(label: str, color: str = "primary", payload: dict[str, Any] | None = None) -> dict[str, Any]:
        action: dict[str, Any] = {"type": "text", "label": label}
        if payload is not None:
            action["payload"] = json.dumps(payload, ensure_ascii=False)
        return {"action": action, "color": color}

    def create_keyboard(self, rows: Iterable[Iterable[Any]]) -> dict[str, Any]:
        buttons = []
        for row in rows:
            prepared = []
            for item in row:
                if isinstance(item, dict):
                    prepared.append(item)
                elif isinstance(item, tuple):
                    label, color, *rest = item
                    payload = rest[0] if rest else None
                    prepared.append(self._button(label, color, payload))
                else:
                    prepared.append(self._button(str(item)))
            if prepared:
                buttons.append(prepared)
        return {"one_time": False, "buttons": buttons}

    @staticmethod
    def remove_keyboard() -> dict[str, Any]:
        return {"one_time": True, "buttons": []}

    def is_admin(self, user_id: int) -> bool:
        return user_id in ADMIN_IDS

    def get_participant_menu(self) -> dict[str, Any]:
        return self.create_keyboard([
            [("📊 Мой прогресс", "primary")],
            [("❓ Помощь", "secondary")],
        ])

    def get_organizer_menu(self) -> dict[str, Any]:
        return self.create_keyboard([
            [("🔑 Коды станций", "primary"), ("📈 Статистика", "primary")],
            [("🏆 Победители", "positive"), ("🎰 Розыгрыш", "positive")],
            [("📋 Полный отчёт", "secondary"), ("❓ Помощь", "secondary")],
        ])

    def get_admin_menu(self) -> dict[str, Any]:
        return self.create_keyboard([
            [("🔑 Коды станций", "primary"), ("📈 Статистика", "primary")],
            [("🏆 Победители", "positive"), ("🎰 Розыгрыш", "positive")],
            [("🛠 Управление станциями", "primary"), ("📊 Excel-отчёт", "positive")],
            [("♻️ Сбросить розыгрыш", "secondary"), ("❓ Помощь", "secondary")],
        ])

    def menu_for(self, user_id: int) -> dict[str, Any]:
        if self.is_admin(user_id):
            return self.get_admin_menu()
        return self.get_organizer_menu() if self.db.get_role(user_id) == "organizer" else self.get_participant_menu()

    def handle_start(self, user_id: int, username: str | None,
                     first_name: str | None, last_name: str | None) -> None:
        self.db.add_user(user_id, username, first_name, last_name)
        self.user_states.pop(user_id, None)
        if self.is_admin(user_id):
            self.db.update_role(user_id, "admin")
            self.bot.send_message(
                user_id,
                "👑 Панель администратора\n\nУправляйте станциями, смотрите статистику, выгружайте Excel и проводите розыгрыш.",
                self.get_admin_menu(),
            )
            return
        role = self.db.get_role(user_id)
        if role == "organizer":
            self.bot.send_message(user_id, "🏢 С возвращением! Панель организатора:", self.get_organizer_menu())
        elif self.db.get_fio(user_id):
            self.bot.send_message(user_id, "👋 С возвращением! Ваше меню участника:", self.get_participant_menu())
        else:
            self.bot.send_message(
                user_id,
                "👋 Добро пожаловать на День карьеры!\n\nВыберите, как продолжить:",
                self.create_keyboard([[("🎯 Участник", "primary"), ("🏢 Организатор", "secondary")]]),
            )

    def handle_role_selection(self, user_id: int, role: str) -> None:
        if role == "🎯 Участник":
            self.db.update_role(user_id, "participant")
            self.user_states[user_id] = "waiting_fio"
            self.bot.send_message(
                user_id,
                "🎯 Регистрация участника\n\nВведите ФИО полностью: фамилия, имя и отчество.\nПример: Иванов Иван Иванович.\n\nДля отмены напишите «Отмена».",
                self.remove_keyboard(),
            )
        elif role == "🏢 Организатор":
            if not ORGANIZER_PASSWORD:
                self.bot.send_message(user_id, "🔒 Вход организатора временно отключён. Обратитесь к главному администратору.")
                return
            self.user_states[user_id] = "waiting_organizer_password"
            self.bot.send_message(
                user_id,
                "🏢 Вход организатора\n\nВведите пароль доступа. Для отмены напишите «Отмена».",
                self.remove_keyboard(),
            )

    def handle_organizer_password(self, user_id: int, password: str) -> None:
        self.user_states.pop(user_id, None)
        if secrets.compare_digest(password, ORGANIZER_PASSWORD):
            self.db.update_role(user_id, "organizer")
            self.bot.send_message(user_id, "✅ Доступ предоставлен. Панель организатора готова.", self.get_organizer_menu())
        else:
            self.bot.send_message(user_id, "❌ Пароль не подошёл. Повторите /start и попробуйте ещё раз.")

    def handle_fio(self, user_id: int, fio_text: str) -> None:
        fio = " ".join(fio_text.split())
        if len(fio) >= 7 and len(fio.split()) >= 2 and re.fullmatch(r"[A-Za-zА-Яа-яЁё -]+", fio):
            self.db.update_fio(user_id, fio)
            self.user_states.pop(user_id, None)
            stats = self.db.get_user_stats(user_id)
            self.bot.send_message(
                user_id,
                f"✅ Вы зарегистрированы, {fio}!\n\n"
                "Как участвовать:\n"
                "1. Посетите все станции.\n"
                "2. Получите у представителя станции кодовое слово.\n"
                "3. Отправьте код в этот чат.\n\n"
                f"📊 Прогресс: {stats['visited_count']}/{stats['total_stations']}\n"
                f"⏱ Между отметками действует интервал {MIN_TIME_BETWEEN_STATIONS // 60} мин.",
                self.get_participant_menu(),
            )
        else:
            self.bot.send_message(user_id, "⚠️ Укажите ФИО буквами, минимум фамилию и имя. Попробуйте ещё раз или напишите «Отмена».")

    @staticmethod
    def _progress_bar(done: int, total: int, width: int = 10) -> str:
        ratio = done / total if total else 0
        filled = min(width, round(ratio * width))
        return "🟧" * filled + "⬜" * (width - filled)

    def handle_my_progress(self, user_id: int) -> str:
        if self.db.get_role(user_id) != "participant" or not self.db.get_fio(user_id):
            return "Сначала зарегистрируйтесь как участник через /start."
        stats = self.db.get_user_stats(user_id)
        visits = self.db.get_visits(user_id)
        can_visit, remaining = self.db.can_visit_station(user_id)
        response = (
            f"👤 {self.db.get_fio(user_id)}\n\n"
            f"📊 Прогресс: {stats['visited_count']} из {stats['total_stations']} станций\n"
            f"{self._progress_bar(int(stats['visited_count']), int(stats['total_stations']))}\n"
            f"Осталось: {stats['remaining']}"
        )
        if stats["completed"]:
            response += "\n\n🎉 Маршрут завершён! Вы участвуете в розыгрыше."
        elif not can_visit:
            minutes, seconds = divmod(int(remaining), 60)
            response += f"\n\n⏱ Следующая отметка будет доступна через {minutes} мин {seconds} сек."
        if visits:
            response += "\n\n✅ Посещённые станции:\n" + "\n".join(
                f"• {name} — {visited_at[:16].replace('T', ' ')}" for name, visited_at in visits
            )
        return response

    def handle_actual_codes(self, user_id: int) -> str:
        if self.db.get_role(user_id) not in ("organizer", "admin") and not self.is_admin(user_id):
            return "❌ Эта функция доступна только организаторам."
        stations = self.db.get_active_stations()
        if not stations:
            return "Сейчас нет активных станций."
        lines = ["🔑 АКТУАЛЬНЫЕ КОДЫ СТАНЦИЙ\n"]
        for station in stations:
            lines.append(f"📍 {station['id']}. {station['name']}\n   Код: {station['code_word']}\n   Команда: /visit_{station['id']}\n")
        return "\n".join(lines)

    def handle_statistics(self, user_id: int) -> str:
        if self.db.get_role(user_id) not in ("organizer", "admin") and not self.is_admin(user_id):
            return "❌ Эта функция доступна только организаторам."
        stats = self.db.get_all_stats()
        return (
            "📈 СТАТИСТИКА ДНЯ КАРЬЕРЫ\n\n"
            f"👥 Зарегистрировано участников: {stats['total_users']}\n"
            f"✅ Всего отметок: {stats['total_visits']}\n"
            f"🏁 Завершили маршрут: {stats['completed_users']}\n"
            f"📍 Активных станций: {stats['total_stations']}"
        )

    def handle_winners(self, user_id: int) -> str:
        if self.db.get_role(user_id) not in ("organizer", "admin") and not self.is_admin(user_id):
            return "❌ Эта функция доступна только организаторам."
        winners = self.db.get_winners()
        if not winners:
            return "🏆 Пока нет участников, завершивших все активные станции."
        lines = ["🏆 УЧАСТНИКИ, ЗАВЕРШИВШИЕ МАРШРУТ\n"]
        for index, (winner_id, fio, username) in enumerate(winners, 1):
            profile = f"@{username}" if username else f"VK ID {winner_id}"
            lines.append(f"{index}. {fio} — {profile}")
        lines.append(f"\nВсего: {len(winners)}")
        return "\n".join(lines)

    def handle_full_report(self, user_id: int) -> str:
        if self.db.get_role(user_id) not in ("organizer", "admin") and not self.is_admin(user_id):
            return "❌ Эта функция доступна только организаторам."
        stats = self.db.get_all_stats()
        users = self.db.get_all_users_with_progress()
        lines = [
            "📋 ПОЛНЫЙ ОТЧЁТ ДНЯ КАРЬЕРЫ\n",
            f"Участников: {stats['total_users']}",
            f"Посещений: {stats['total_visits']}",
            f"Завершили маршрут: {stats['completed_users']}",
            f"Активных станций: {stats['total_stations']}",
            "\nПРОГРЕСС УЧАСТНИКОВ:",
        ]
        for index, (uid, fio, username, visited, total) in enumerate(users, 1):
            profile = f"@{username}" if username else f"ID {uid}"
            status = "ЗАВЕРШЁН" if total > 0 and visited >= total else f"{visited}/{total}"
            lines.append(f"{index}. {fio} ({profile}) — {status}")
        # VK message size is limited; export is the full detailed report.
        return "\n".join(lines)

    def _run_roulette(self, user_id: int) -> None:
        candidates = self.db.get_drawable_winners()
        if not candidates:
            completed = self.db.get_winners()
            if completed:
                message = (
                    "🎰 Все участники, завершившие маршрут, уже выигрывали.\n\n"
                    "Администратор может сбросить историю розыгрыша и начать новый круг."
                )
            else:
                message = "🎰 РОЗЫГРЫШ\n\nПока нет участников, прошедших все активные станции."
            self.bot.send_message(user_id, message, self.menu_for(user_id))
            return

        winner = random.SystemRandom().choice(candidates)
        winner_id, winner_fio, username = winner
        # Reserve the winner before animation so concurrent clicks cannot select them twice.
        self.db.record_raffle_winner(winner_id, user_id)

        message_id = self.bot.send_message(
            user_id,
            "🎰 РОЗЫГРЫШ ПРИЗОВ\n\nПроверяем список участников…\n\n🔄 Рулетка запускается!",
        )
        if message_id:
            frames = ["◴", "◷", "◶", "◵", "🎯", "🎰"]
            for index in range(10):
                candidate = random.SystemRandom().choice(candidates)
                frame = frames[index % len(frames)]
                names = "   •   ".join(item[1] for item in random.sample(candidates, min(3, len(candidates))))
                animation = (
                    "🎰 РОЗЫГРЫШ ПРИЗОВ\n\n"
                    f"{frame}  {candidate[1]}\n\n"
                    f"Участники: {len(candidates)}\n"
                    f"{names}\n\n"
                    f"{'▰' * (index + 1)}{'▱' * (9 - index)}"
                )
                self.bot.edit_message(user_id, message_id, animation)
                time.sleep(0.45 + index * 0.035)
            final_text = (
                "🎉🎉🎉 ПОБЕДИТЕЛЬ! 🎉🎉🎉\n\n"
                f"🏆 {winner_fio}\n"
                f"👤 {'@' + username if username else 'VK ID ' + str(winner_id)}\n\n"
                "Поздравляем! Победитель зафиксирован в истории розыгрыша."
            )
            if not self.bot.edit_message(user_id, message_id, final_text):
                self.bot.send_message(user_id, final_text, self.menu_for(user_id))
            else:
                self.bot.send_message(user_id, "Выберите следующее действие:", self.menu_for(user_id))
        else:
            self.bot.send_message(
                user_id,
                f"🏆 Победитель: {winner_fio}\n" + (f"@{username}" if username else f"VK ID {winner_id}"),
                self.menu_for(user_id),
            )

    def handle_help(self, user_id: int) -> str:
        role = "admin" if self.is_admin(user_id) else self.db.get_role(user_id)
        if role == "admin":
            return (
                "❓ ПОМОЩЬ АДМИНИСТРАТОРА\n\n"
                "• Управление станциями — добавление и отключение станций.\n"
                "• Excel-отчёт — сводка, участники и подробный журнал посещений.\n"
                "• Розыгрыш — последовательный выбор победителей без повторов.\n"
                "• Сбросить розыгрыш — очистить историю выбранных победителей.\n\n"
                "Команды: /visit_ID [VK_ID], /complete_all VK_ID, /export."
            )
        if role == "organizer":
            return (
                "❓ ПОМОЩЬ ОРГАНИЗАТОРА\n\n"
                "• Коды станций — список кодовых слов.\n"
                "• Статистика и победители — результаты мероприятия.\n"
                "• Розыгрыш — случайный выбор среди прошедших все активные станции.\n\n"
                "Для отметки участника администратор использует /visit_ID [VK_ID]."
            )
        return (
            "❓ ПОМОЩЬ УЧАСТНИКА\n\n"
            "1. Посетите каждую станцию.\n"
            "2. Получите кодовое слово у представителя.\n"
            "3. Отправьте кодовое слово в чат.\n"
            "4. Следите за прогрессом в разделе «Мой прогресс».\n\n"
            f"Между отметками действует интервал {MIN_TIME_BETWEEN_STATIONS // 60} мин."
        )

    # -- Admin station management -----------------------------------------
    def send_station_admin_menu(self, user_id: int) -> None:
        if not self.is_admin(user_id):
            self.bot.send_message(user_id, "❌ Недостаточно прав.")
            return
        stations = self.db.get_active_stations()
        lines = ["🛠 УПРАВЛЕНИЕ СТАНЦИЯМИ\n", f"Активных станций: {len(stations)}", ""]
        for station in stations:
            lines.append(f"{station['id']}. {station['name']} — код: {station['code_word']}")
        keyboard = self.create_keyboard([
            [("➕ Добавить станцию", "positive", {"action": "station_add"})],
            [("➖ Убрать станцию", "negative", {"action": "station_remove_menu"})],
            [("↩️ В меню", "secondary", {"action": "main_menu"})],
        ])
        self.bot.send_message(user_id, "\n".join(lines), keyboard)

    def send_station_remove_menu(self, user_id: int) -> None:
        if not self.is_admin(user_id):
            self.bot.send_message(user_id, "❌ Недостаточно прав.")
            return
        stations = self.db.get_active_stations()
        if len(stations) <= 1:
            self.bot.send_message(user_id, "Нельзя отключить последнюю активную станцию.", self.get_admin_menu())
            return
        rows = []
        for station in stations:
            label = f"🗑 {station['id']} — {station['name']}"
            # VK button labels are length-limited; the ID in payload is authoritative.
            rows.append([(label[:40], "negative", {"action": "station_remove", "station_id": station["id"]})])
        rows.append([("↩️ Назад", "secondary", {"action": "station_admin"})])
        self.bot.send_message(user_id, "Выберите станцию, которую нужно отключить:\nИстория посещений сохранится, но станция исчезнет из текущего маршрута.", self.create_keyboard(rows))

    def start_add_station(self, user_id: int) -> None:
        if not self.is_admin(user_id):
            self.bot.send_message(user_id, "❌ Недостаточно прав.")
            return
        self.user_states[user_id] = "waiting_station_data"
        self.bot.send_message(
            user_id,
            "➕ ДОБАВЛЕНИЕ СТАНЦИИ\n\n"
            "Отправьте данные одним сообщением в формате:\n"
            "Название станции | КОДОВОЕ_СЛОВО\n\n"
            "Пример: Новое предприятие | МЕДЬ2026\n\n"
            "Код должен быть уникальным. Для отмены напишите «Отмена».",
            self.create_keyboard([[ ("Отмена", "secondary", {"action": "cancel_state"}) ]]),
        )

    def add_station_from_text(self, user_id: int, text: str) -> None:
        self.user_states.pop(user_id, None)
        parts = re.split(r"\s*[|;]\s*", text.strip(), maxsplit=1)
        if len(parts) != 2:
            self.bot.send_message(user_id, "⚠️ Формат не распознан. Нужно: Название станции | КОДОВОЕ_СЛОВО", self.get_admin_menu())
            return
        name, code = parts[0].strip(), parts[1].strip().upper()
        if not (2 <= len(name) <= 90):
            self.bot.send_message(user_id, "Название станции должно содержать от 2 до 90 символов.", self.get_admin_menu())
            return
        if not re.fullmatch(r"[A-ZА-ЯЁ0-9_-]{2,32}", code):
            self.bot.send_message(user_id, "Код: от 2 до 32 символов — только буквы, цифры, дефис или подчёркивание.", self.get_admin_menu())
            return
        try:
            station_id = self.db.add_station(name, code)
        except sqlite3.IntegrityError:
            self.bot.send_message(user_id, "❌ Такой код уже используется активной или архивной станцией. Выберите другой.", self.get_admin_menu())
            return
        self.bot.send_message(
            user_id,
            f"✅ Станция добавлена!\n\n📍 {name}\n🔑 Код: {code}\n🆔 ID: {station_id}\nКоманда отметки: /visit_{station_id}",
            self.get_admin_menu(),
        )

    def remove_station(self, user_id: int, station_id: int) -> None:
        if not self.is_admin(user_id):
            self.bot.send_message(user_id, "❌ Недостаточно прав.")
            return
        station = next((s for s in self.db.get_active_stations() if s["id"] == station_id), None)
        if not station:
            self.bot.send_message(user_id, "Эта станция уже отключена или не найдена.", self.get_admin_menu())
            return
        success = self.db.deactivate_station(station_id)
        if success:
            self.bot.send_message(
                user_id,
                f"✅ Станция отключена: {station['name']}\n\nИстория посещений сохранена. Прогресс и условия розыгрыша пересчитаны по активным станциям.",
                self.get_admin_menu(),
            )
        else:
            self.bot.send_message(user_id, "❌ Нельзя отключить последнюю активную станцию.", self.get_admin_menu())

    # -- Excel export ------------------------------------------------------
    def build_excel_report(self) -> bytes:
        dataset = self.db.export_dataset()
        stations = dataset["stations"]
        participants = dataset["participants"]
        visits = dataset["visits"]
        stats = self.db.get_all_stats()

        wb = Workbook()
        summary = wb.active
        summary.title = "Сводка"
        summary.append(["Показатель", "Значение"])
        summary_rows = [
            ("Дата выгрузки", datetime.now().strftime("%d.%m.%Y %H:%M")),
            ("Зарегистрировано участников", stats["total_users"]),
            ("Активных станций", stats["total_stations"]),
            ("Всего действующих отметок", stats["total_visits"]),
            ("Участников завершили маршрут", stats["completed_users"]),
            ("Участников в розыгрыше без предыдущего выигрыша", len(self.db.get_drawable_winners())),
        ]
        for row in summary_rows:
            summary.append(list(row))

        sheet = wb.create_sheet("Участники")
        headers = ["№", "ФИО", "VK ID", "Username", "Дата регистрации", "Посещено", "Всего станций", "Прогресс, %", "Статус"]
        headers.extend(station["name"] for station in stations)
        sheet.append(headers)
        for index, participant in enumerate(participants, 1):
            total = participant["total_stations"]
            done = participant["visited_count"]
            percent = round(done / total * 100, 1) if total else 0
            row = [
                index,
                participant["fio"],
                participant["user_id"],
                participant["username"] or "",
                participant["registered_at"],
                done,
                total,
                percent,
                "Маршрут завершён" if total and done >= total else "В процессе",
            ]
            for station in stations:
                visited_at = participant["visits"].get(station["id"])
                row.append(visited_at.replace("T", " ") if visited_at else "—")
            sheet.append(row)

        visit_sheet = wb.create_sheet("Посещения")
        visit_sheet.append(["№", "ФИО", "VK ID", "Станция", "Дата и время"])
        fio_by_id = {p["user_id"]: p["fio"] for p in participants}
        for index, visit in enumerate(visits, 1):
            visit_sheet.append([
                index,
                fio_by_id.get(visit["user_id"], ""),
                visit["user_id"],
                visit["station_name"],
                visit["visited_at"].replace("T", " "),
            ])

        station_sheet = wb.create_sheet("Станции")
        station_sheet.append(["ID", "Станция", "Посещений", "Уникальных участников"])
        for station in stations:
            matched = [v for v in visits if v["station_id"] == station["id"]]
            station_sheet.append([
                station["id"], station["name"], len(matched), len({v["user_id"] for v in matched})
            ])

        header_fill = PatternFill("solid", fgColor="17365D")
        header_font = Font(color="FFFFFF", bold=True)
        thin_gray = Side(style="thin", color="D9E2F3")
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions if ws.max_row > 1 else "A1"
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(bottom=thin_gray)
            ws.row_dimensions[1].height = 32
            for column_cells in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column_cells[0].column)
                for cell in column_cells[:250]:
                    value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, min(len(value), 48))
                    if cell.row > 1:
                        cell.alignment = Alignment(vertical="top", wrap_text=True)
                ws.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 42)
        summary.column_dimensions["A"].width = 48
        summary.column_dimensions["B"].width = 28

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    def export_excel(self, user_id: int) -> None:
        if not self.is_admin(user_id):
            self.bot.send_message(user_id, "❌ Excel-выгрузка доступна только администратору.")
            return
        try:
            content = self.build_excel_report()
            filename = f"career_day_report_{datetime.now():%Y%m%d_%H%M}.xlsx"
            sent = self.bot.send_document(
                user_id, filename, content,
                "📊 Отчёт готов. В файле 4 листа: сводка, участники, посещения и статистика по станциям.",
            )
            if not sent:
                self.bot.send_message(user_id, "⚠️ Excel сформирован, но VK не принял файл. Проверьте права доступа сообщества к сообщениям и документам.", self.get_admin_menu())
        except Exception:
            logger.exception("Ошибка формирования Excel")
            self.bot.send_message(user_id, "❌ Не удалось сформировать Excel. Подробности доступны в логах хостинга.", self.get_admin_menu())

    def handle_admin_visit(self, user_id: int, command_text: str) -> str:
        if not self.is_admin(user_id):
            return "❌ У вас нет прав для этой команды."
        try:
            parts = command_text.split()
            station_id = int(parts[0])
            target_user_id = int(parts[1]) if len(parts) > 1 else user_id
        except (ValueError, IndexError):
            return "Формат: /visit_ID или /visit_ID VK_ID"
        station = next((s for s in self.db.get_active_stations() if s["id"] == station_id), None)
        if not station:
            return "❌ Активная станция не найдена. Проверьте ID в меню станций."
        if self.db.record_visit(target_user_id, station_id, station["name"]):
            stats = self.db.get_user_stats(target_user_id)
            fio = self.db.get_fio(target_user_id) or f"Пользователь {target_user_id}"
            status = "\n🎉 Маршрут завершён!" if stats["completed"] else ""
            return f"✅ Отметка добавлена: {fio}\n📍 {station['name']}\n📊 {stats['visited_count']}/{stats['total_stations']}{status}"
        return f"ℹ️ Не удалось добавить отметку. Возможно, пользователь уже посещал станцию {station['name']}."

    def handle_complete_all_for_user(self, user_id: int, command_text: str) -> str:
        if not self.is_admin(user_id):
            return "❌ У вас нет прав для этой команды."
        try:
            target_user_id = int(command_text.strip())
        except ValueError:
            return "Формат: /complete_all VK_ID"
        fio = self.db.get_fio(target_user_id)
        if not fio:
            return "Пользователь не найден или ещё не указал ФИО."
        added = self.db.complete_all_stations_for_user(target_user_id)
        stats = self.db.get_user_stats(target_user_id)
        return f"✅ Для {fio} добавлено отметок: {added}.\nПрогресс: {stats['visited_count']}/{stats['total_stations']}"

    def handle_code_word(self, user_id: int, code_word: str) -> str:
        if self.db.get_role(user_id) != "participant":
            return "Эта функция доступна только участникам."
        if not self.db.get_fio(user_id):
            return "Сначала зарегистрируйтесь: напишите /start и укажите ФИО."
        can_visit, remaining = self.db.can_visit_station(user_id)
        if not can_visit:
            minutes, seconds = divmod(int(remaining), 60)
            return f"⏱ Следующую станцию можно отметить через {minutes} мин {seconds} сек."
        station = self.db.get_station_by_code(code_word)
        if not station:
            return "❌ Код не распознан. Проверьте написание или уточните код у представителя станции."
        if self.db.record_visit(user_id, station["id"], station["name"]):
            stats = self.db.get_user_stats(user_id)
            response = (
                f"✅ Отметка принята!\n\n📍 {station['name']}\n"
                f"📊 Прогресс: {stats['visited_count']}/{stats['total_stations']}\n"
                f"Осталось: {stats['remaining']}"
            )
            if stats["completed"]:
                response += "\n\n🎉 Вы завершили маршрут и допущены к розыгрышу призов!"
            return response
        return f"ℹ️ Станция «{station['name']}» уже отмечена в вашем профиле."

    @staticmethod
    def _parse_payload(raw: Any) -> dict[str, Any]:
        if isinstance(raw, dict):
            return raw
        if isinstance(raw, str) and raw:
            try:
                value = json.loads(raw)
                return value if isinstance(value, dict) else {}
            except json.JSONDecodeError:
                return {}
        return {}

    def process_update(self, update: dict[str, Any]) -> None:
        if update.get("type") != "message_new":
            return
        message = (update.get("object") or {}).get("message") or {}
        try:
            user_id = int(message.get("from_id", 0))
            peer_id = int(message.get("peer_id", 0))
        except (TypeError, ValueError):
            return
        if user_id <= 0 or peer_id != user_id:
            return  # only private user-to-community messages

        text = str(message.get("text", "")).strip()
        payload = self._parse_payload(message.get("payload"))

        can_send, spam_message = self.db.check_spam_protection(user_id)
        if not can_send:
            if spam_message:
                self.bot.send_message(user_id, spam_message)
            return

        if text.lower() in ("/start", "начать", "start"):
            username, first_name, last_name = self.bot.get_user(user_id)
            self.handle_start(user_id, username, first_name, last_name)
            return

        action = payload.get("action")
        if action == "main_menu":
            self.bot.send_message(user_id, "Главное меню:", self.menu_for(user_id))
            return
        if action == "station_admin":
            self.send_station_admin_menu(user_id)
            return
        if action == "station_add":
            self.start_add_station(user_id)
            return
        if action == "station_remove_menu":
            self.send_station_remove_menu(user_id)
            return
        if action == "station_remove":
            try:
                self.remove_station(user_id, int(payload.get("station_id")))
            except (TypeError, ValueError):
                self.bot.send_message(user_id, "Не удалось определить станцию.")
            return
        if action == "cancel_state":
            self.user_states.pop(user_id, None)
            self.bot.send_message(user_id, "Действие отменено.", self.get_admin_menu() if self.is_admin(user_id) else self.menu_for(user_id))
            return
        if action == "raffle_reset_confirm":
            if self.is_admin(user_id):
                count = self.db.reset_raffle()
                self.bot.send_message(user_id, f"♻️ История розыгрыша очищена. Сброшено победителей: {count}.", self.get_admin_menu())
            return
        if action == "raffle_reset_cancel":
            self.bot.send_message(user_id, "Сброс отменён.", self.get_admin_menu())
            return

        if self.user_states.get(user_id) == "waiting_station_data":
            if text.lower() == "отмена":
                self.user_states.pop(user_id, None)
                self.bot.send_message(user_id, "Добавление станции отменено.", self.get_admin_menu())
            elif self.is_admin(user_id):
                self.add_station_from_text(user_id, text)
            return
        if self.user_states.get(user_id) == "waiting_fio":
            if text.lower() == "отмена":
                self.user_states.pop(user_id, None)
                self.bot.send_message(user_id, "Регистрация отменена. Чтобы начать снова, напишите /start.")
            else:
                self.handle_fio(user_id, text)
            return
        if self.user_states.get(user_id) == "waiting_organizer_password":
            if text.lower() == "отмена":
                self.user_states.pop(user_id, None)
                self.bot.send_message(user_id, "Вход отменён. Напишите /start для выбора роли.")
            else:
                self.handle_organizer_password(user_id, text)
            return

        if text in ("🎯 Участник", "🏢 Организатор"):
            self.handle_role_selection(user_id, text)
        elif text == "📊 Мой прогресс":
            self.bot.send_message(user_id, self.handle_my_progress(user_id), self.menu_for(user_id))
        elif text in ("🔑 Актуальные коды", "🔑 Коды станций"):
            self.bot.send_message(user_id, self.handle_actual_codes(user_id), self.menu_for(user_id))
        elif text == "📈 Статистика":
            self.bot.send_message(user_id, self.handle_statistics(user_id), self.menu_for(user_id))
        elif text == "🏆 Победители":
            self.bot.send_message(user_id, self.handle_winners(user_id), self.menu_for(user_id))
        elif text == "🎰 Розыгрыш":
            if self.db.get_role(user_id) in ("organizer", "admin") or self.is_admin(user_id):
                self._run_roulette(user_id)
            else:
                self.bot.send_message(user_id, "❌ Розыгрыш доступен только организаторам.")
        elif text in ("📋 Полный отчёт", "📋 Полный отчет"):
            self.bot.send_message(user_id, self.handle_full_report(user_id), self.menu_for(user_id))
        elif text == "❓ Помощь":
            self.bot.send_message(user_id, self.handle_help(user_id), self.menu_for(user_id))
        elif text == "🛠 Управление станциями":
            self.send_station_admin_menu(user_id)
        elif text == "📊 Excel-отчёт":
            self.export_excel(user_id)
        elif text == "♻️ Сбросить розыгрыш":
            if self.is_admin(user_id):
                self.bot.send_message(
                    user_id,
                    "Сбросить историю победителей? После подтверждения ранее выбранные участники снова смогут участвовать.",
                    self.create_keyboard([[
                        ("Да, сбросить", "negative", {"action": "raffle_reset_confirm"}),
                        ("Отмена", "secondary", {"action": "raffle_reset_cancel"}),
                    ]]),
                )
            else:
                self.bot.send_message(user_id, "❌ Недостаточно прав.")
        elif text.startswith("/visit_"):
            self.bot.send_message(user_id, self.handle_admin_visit(user_id, text[len("/visit_"):]), self.menu_for(user_id))
        elif text.startswith("/complete_all"):
            self.bot.send_message(user_id, self.handle_complete_all_for_user(user_id, text[len("/complete_all"):]), self.menu_for(user_id))
        elif text == "/export":
            self.export_excel(user_id)
        elif text and not text.startswith("/"):
            station = self.db.get_station_by_code(text)
            if station:
                self.bot.send_message(user_id, self.handle_code_word(user_id, text), self.get_participant_menu())
            else:
                self.bot.send_message(
                    user_id,
                    "Не удалось распознать команду или код. Воспользуйтесь кнопками меню или напишите /start.",
                    self.menu_for(user_id),
                )


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

def main() -> None:
    if not VK_TOKEN:
        raise RuntimeError("Не задан VK_TOKEN. Добавьте токен сообщества в переменные окружения Bothost.")
    if GROUP_ID <= 0:
        raise RuntimeError("VK_GROUP_ID должен быть положительным ID сообщества без знака минус.")
    if not ADMIN_IDS:
        raise RuntimeError("Не заданы ADMIN_IDS. Укажите VK ID администраторов через запятую.")

    logger.info("Запуск VK-бота Дня карьеры")
    logger.info("ID сообщества: %s", GROUP_ID)
    logger.info("Администраторов в конфигурации: %s", len(ADMIN_IDS))
    logger.info("База данных: %s", DB_PATH)
    logger.info("Минимальный интервал между отметками: %s сек.", MIN_TIME_BETWEEN_STATIONS)

    handler = BotHandler()
    active_count = len(handler.db.get_active_stations())
    logger.info("Активных станций: %s", active_count)

    while True:
        try:
            for update in handler.bot.get_updates():
                try:
                    handler.process_update(update)
                except Exception:
                    logger.exception("Ошибка обработки события VK")
        except KeyboardInterrupt:
            logger.info("Остановка по сигналу")
            break
        except Exception:
            logger.exception("Ошибка основного цикла")
            time.sleep(5)


if __name__ == "__main__":
    main()
