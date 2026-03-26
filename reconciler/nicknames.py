import json
import re
from pathlib import Path

_NICKNAMES_FILE = Path(__file__).parent.parent / "nicknames.json"


def load_nicknames() -> list[dict]:
    """Return list of {id, pattern, nickname} dicts."""
    try:
        data = json.loads(_NICKNAMES_FILE.read_text(encoding="utf-8"))
        # Ensure every entry has an id
        for i, entry in enumerate(data):
            if "id" not in entry:
                entry["id"] = i
        return data
    except FileNotFoundError:
        return []
    except Exception:
        return []


def save_nicknames(nicknames: list[dict]) -> None:
    _NICKNAMES_FILE.write_text(
        json.dumps(nicknames, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def match_description(description: str, nicknames: list[dict] | None = None) -> list[dict]:
    """Return list of matching {id, pattern, nickname} for the given description."""
    if nicknames is None:
        nicknames = load_nicknames()
    matches = []
    for entry in nicknames:
        try:
            if re.search(entry["pattern"], description, re.IGNORECASE):
                matches.append(entry)
        except re.error:
            pass
    return matches


def best_match(description: str, nicknames: list[dict] | None = None) -> str | None:
    """Return the nickname of the first matching entry, or None."""
    matches = match_description(description, nicknames)
    return matches[0]["nickname"] if matches else None


def next_id(nicknames: list[dict]) -> int:
    return max((e.get("id", 0) for e in nicknames), default=-1) + 1
