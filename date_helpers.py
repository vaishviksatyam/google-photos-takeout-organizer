"""Date parsing helpers used by the Takeout organizer."""

import json
import re
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

from PIL import Image

MIN_VALID_YEAR = 1950
MAX_VALID_YEAR = 2100


def is_valid_media_datetime(value: datetime | None) -> bool:
    return value is not None and MIN_VALID_YEAR <= value.year <= MAX_VALID_YEAR


def get_exif_datetime(filepath: Path) -> datetime | None:
    try:
        with Image.open(filepath) as image:
            exif = image.getexif()
            if not exif:
                return None
            date_string = exif.get(36867) or exif.get(306)
            if date_string:
                parsed = datetime.strptime(date_string, "%Y:%m:%d %H:%M:%S")
                return parsed if is_valid_media_datetime(parsed) else None
    except Exception:
        pass
    return None


def find_thm_sidecar(video_path: Path) -> Path | None:
    for extension in (".thm", ".THM"):
        sidecar = video_path.with_suffix(extension)
        if sidecar.is_file():
            return sidecar
    return None


def get_video_created_datetime(filepath: Path) -> datetime | None:
    """Read a valid creation_time tag with ffprobe when it is installed."""
    executable = shutil.which("ffprobe")
    if not executable:
        return None
    try:
        result = subprocess.run(
            [
                executable,
                "-v", "error",
                "-show_entries", "format_tags=creation_time:stream_tags=creation_time",
                "-of", "json",
                str(filepath),
            ],
            capture_output=True,
            text=True,
            timeout=30,
            check=True,
        )
        metadata = json.loads(result.stdout)
    except (OSError, subprocess.SubprocessError, json.JSONDecodeError):
        return None

    if not isinstance(metadata, dict):
        return None
    format_data = metadata.get("format")
    format_tags = format_data.get("tags", {}) if isinstance(format_data, dict) else {}
    streams = metadata.get("streams", [])
    if not isinstance(streams, list):
        streams = []
    tag_sets = [format_tags]
    tag_sets.extend(
        stream.get("tags", {}) for stream in streams if isinstance(stream, dict)
    )
    for tags in tag_sets:
        if not isinstance(tags, dict):
            continue
        for name, date_text in tags.items():
            if name.casefold() != "creation_time" or not isinstance(date_text, str):
                continue
            try:
                parsed = datetime.fromisoformat(date_text.replace("Z", "+00:00"))
            except ValueError:
                try:
                    parsed = datetime.strptime(date_text, "%Y:%m:%d %H:%M:%S")
                except ValueError:
                    continue
            if parsed.tzinfo:
                parsed = parsed.astimezone().replace(tzinfo=None)
            if is_valid_media_datetime(parsed):
                return parsed
    return None


def parse_datetime_from_filename(filename: str) -> datetime | None:
    """Parse supported camera, messaging-app, and video date patterns."""
    patterns = (
        # Dotted WhatsApp date/time and VLC-style video snapshot.
        (r"(\d{4})-(\d{2})-(\d{2})(?:[ T]| at )(\d{2})\.(\d{2})\.(\d{2})", "ymdhms"),
        (r"(\d{4})-(\d{2})-(\d{2})-(\d{2})h(\d{2})m(\d{2})s", "ymdhms"),
        # Camera timestamp formats: separated, compact, and screenshot names.
        (r"(\d{4})(\d{2})(\d{2})_(\d{2})_(\d{2})_(\d{2})", "ymdhms"),
        (r"(\d{4})(\d{2})(\d{2})[-_](\d{2})(\d{2})(\d{2})", "ymdhms"),
        (r"(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", "ymdhms"),
        (r"(\d{4})-(\d{2})-(\d{2})[-_](\d{2})-(\d{2})-(\d{2})", "ymdhms"),
        # Takeout/video names and date-only formats, including day-first dates.
        (r"Video(\d{4})(\d)(\d{2})\d{6}", "ymd"),
        (r"(?<!\d)(\d{2})(\d{2})(\d{4})(?:\d{0,6})", "dmy"),
        (r"WP_(\d{4})(\d{2})(\d{2})", "ymd"),
        (r"(\d{4})(\d{2})(\d{2})", "ymd"),
    )
    for pattern, date_kind in patterns:
        match = re.search(pattern, filename)
        if not match:
            continue
        groups = match.groups()
        if date_kind == "ymdhms":
            date_parts = groups
            date_format = "%Y-%m-%d-%H-%M-%S"
        elif date_kind == "dmy":
            date_parts = (groups[2], groups[1], groups[0])
            date_format = "%Y-%m-%d"
        else:
            date_parts = groups
            date_format = "%Y-%m-%d"
        try:
            parsed = datetime.strptime("-".join(date_parts), date_format)
        except ValueError:
            continue
        if is_valid_media_datetime(parsed):
            return parsed
    return None