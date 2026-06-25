"""Claude API assist — propose, never apply.

The deterministic engine stays in charge of every hard rule (students never
double-booked, A211 protected, group hierarchy, the export gate) and computes the
*legal* options. The AI only does the fuzzy work:

  interpret_comment(...)   read a messy "Other comments" cell -> structured hint
  suggest_resolution(...)  rank/explain the legal move candidates the solver found

Every result carries: the change, a reason, a confidence, and the source it read.
No key -> {ok: False, error: "no_key"}; the whole app still works without AI.

Key is read from the env var ANTHROPIC_API_KEY or a local (untracked) .env file.
Nothing here is ever written to the repo.
"""
from __future__ import annotations

import datetime
import hashlib
import json
import os
import re
import urllib.error
import urllib.request

from .dictionaries import APP_DIR, CONFIG

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENV_PATH = os.path.join(APP_DIR, ".env")          # written/read next to the app
API_URL = "https://api.anthropic.com/v1/messages"
DEFAULT_CHEAP = "claude-haiku-4-5-20251001"
DEFAULT_STRONG = "claude-sonnet-4-6"

# Model options shown in the picker, with a plain-language recommendation. Pricing is
# USD per 1M tokens (input, output) — approximate, for the spend-cap estimate only.
MODELS = [
    {"id": "claude-opus-4-8", "label": "Opus 4.8 — most capable",
     "note": "Best judgment on tricky, multi-way conflicts. Highest cost.", "price": (15, 75)},
    {"id": "claude-sonnet-4-6", "label": "Sonnet 4.6 — balanced (recommended)",
     "note": "Strong reasoning at a much lower cost. A good default.", "price": (3, 15)},
    {"id": "claude-haiku-4-5-20251001", "label": "Haiku 4.5 — fast & cheap",
     "note": "Quick and inexpensive; fine for simple conflicts.", "price": (1, 5)},
]
_PRICE = {m["id"]: m["price"] for m in MODELS}
_CACHE: dict[str, dict] = {}      # prompt hash -> result, avoids repeat billing in a session


def _settings():
    for base in (os.path.join(APP_DIR, "config"), CONFIG):
        p = os.path.join(base, "settings.json")
        if os.path.exists(p):
            try:
                with open(p, encoding="utf-8-sig") as fh:
                    return json.load(fh) or {}
            except (ValueError, OSError):
                return {}
    return {}


def _env_path():
    for p in (ENV_PATH, os.path.join(ROOT, ".env")):
        if os.path.exists(p):
            return p
    return ENV_PATH


def _key():
    k = os.environ.get("ANTHROPIC_API_KEY")
    if k:
        return k.strip()
    p = _env_path()
    if os.path.exists(p):
        with open(p, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line.startswith("ANTHROPIC_API_KEY") and "=" in line:
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    return None


def write_env(api_key):
    """Create/update ANTHROPIC_API_KEY in the .env next to the app (other lines kept)."""
    lines, found = [], False
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, encoding="utf-8") as fh:
            lines = fh.read().splitlines()
    out = []
    for ln in lines:
        if ln.strip().startswith("ANTHROPIC_API_KEY"):
            out.append(f"ANTHROPIC_API_KEY={api_key}")
            found = True
        else:
            out.append(ln)
    if not found:
        out.append(f"ANTHROPIC_API_KEY={api_key}")
    os.makedirs(APP_DIR, exist_ok=True)
    with open(ENV_PATH, "w", encoding="utf-8") as fh:
        fh.write("\n".join(out) + "\n")
    os.environ["ANTHROPIC_API_KEY"] = api_key      # take effect immediately
    return True


def strong_model():
    return os.environ.get("BA_MODEL_STRONG") or _settings().get("ai_model") or DEFAULT_STRONG


def _cap():
    try:
        return float(_settings().get("ai_cap_usd", 0) or 0)
    except (TypeError, ValueError):
        return 0.0


def _usage_path():
    return os.path.join(APP_DIR, "config", "ai_usage.json")


def _usage():
    month = datetime.datetime.now().strftime("%Y-%m")
    try:
        with open(_usage_path(), encoding="utf-8") as fh:
            u = json.load(fh)
        if u.get("month") == month:
            return u
    except (OSError, ValueError):
        pass
    return {"month": month, "spent_usd": 0.0}


def _record(model, usage):
    pin, pout = _PRICE.get(model, (3, 15))
    cost = (usage.get("input_tokens", 0) / 1e6) * pin + (usage.get("output_tokens", 0) / 1e6) * pout
    u = _usage()
    u["spent_usd"] = round(u.get("spent_usd", 0.0) + cost, 4)
    try:
        os.makedirs(os.path.dirname(_usage_path()), exist_ok=True)
        with open(_usage_path(), "w", encoding="utf-8") as fh:
            json.dump(u, fh)
    except OSError:
        pass


def status():
    cap, u = _cap(), _usage()
    return {"available": bool(_key()), "model": strong_model(), "model_cheap": DEFAULT_CHEAP,
            "cap_usd": cap, "spent_usd": u.get("spent_usd", 0.0)}


def _call(model, system, user, max_tokens=900, temperature=0.0):
    """Raw Messages API call. Raises RuntimeError('no_key'|'cap_exceeded'|msg)."""
    key = _key()
    if not key:
        raise RuntimeError("no_key")
    cap = _cap()
    if cap > 0 and _usage().get("spent_usd", 0.0) >= cap:
        raise RuntimeError(f"cap_exceeded: AI spending cap of ${cap:.2f} reached this month.")
    body = json.dumps({
        "model": model, "max_tokens": max_tokens, "temperature": temperature,
        "system": system, "messages": [{"role": "user", "content": user}],
    }).encode("utf-8")
    req = urllib.request.Request(API_URL, data=body, headers={
        "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.load(r)
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300]
        raise RuntimeError(f"api_error {e.code}: {detail}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"network_error: {e.reason}")
    if data.get("usage"):
        _record(model, data["usage"])
    return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")


def _json_block(text):
    """Tolerantly pull the first JSON object out of a model reply."""
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(0))
    except json.JSONDecodeError:
        return None


def _cached(kind, payload, fn):
    h = kind + ":" + hashlib.sha1(json.dumps(payload, sort_keys=True,
                                             ensure_ascii=False).encode("utf-8")).hexdigest()
    if h in _CACHE:
        return {**_CACHE[h], "cached": True}
    if not _key():
        return {"ok": False, "error": "no_key",
                "message": "Add ANTHROPIC_API_KEY to a local .env to enable AI assist."}
    try:
        res = fn()
    except RuntimeError as e:
        msg = str(e)
        return {"ok": False, "error": msg.split(":", 1)[0], "message": msg}
    if res.get("ok"):
        _CACHE[h] = res
    return res


# --- 1. interpret a messy "Other comments" cell ----------------------------
_INTERPRET_SYS = (
    "You help a Finnish film & media school turn messy free-text booking-request "
    "comments (Swedish, Finnish or English) into structured scheduling hints. "
    "You only read; you never decide. Be conservative: if something is not clearly "
    "stated, leave it null. Reply with ONLY a JSON object:\n"
    '{"summary": "<one short plain-English line>",'
    ' "time_range": "<HH:MM-HH:MM or null>",'
    ' "preferred_slot": "<AM|PM|FULL|null>",'
    ' "room": "<requested room or null>",'
    ' "frequency": "<e.g. 2x90 or null>",'
    ' "double_booking_ok": <true|false|null>,'
    ' "needs_computer": <true|false|null>,'
    ' "confidence": "<high|medium|low>",'
    ' "reason": "<why, citing the words you relied on>"}'
)


def interpret_comment(comment, course=""):
    comment = (comment or "").strip()
    if not comment:
        return {"ok": False, "error": "empty", "message": "No comment text to interpret."}

    def run():
        user = f"Course: {course}\nComment: {comment}\n\nExtract the scheduling hints."
        out = _json_block(_call(DEFAULT_CHEAP, _INTERPRET_SYS, user, max_tokens=500))
        if not out:
            return {"ok": False, "error": "parse", "message": "AI reply was not valid JSON."}
        out.update(ok=True, source="Other comments", model=DEFAULT_CHEAP)
        return out

    return _cached("interpret", {"c": comment, "k": course}, run)


# --- 2. rank/explain the legal resolution candidates the solver found -------
_SUGGEST_SYS = (
    "You advise a scheduler at a film & media school. The deterministic solver has "
    "already enforced every hard rule and given you ONLY legal move options. Your job "
    "is to rank and explain them in plain language — you do not invent new options and "
    "you do not apply anything. Prefer keeping a course on its requested day, honouring "
    "stated time constraints in the comment, minimal disruption, and mid-week days "
    "(Tue-Thu) over Mon/Fri. If the scheduler's house rules say this particular clash is "
    "acceptable to keep, set \"action\":\"approve\". Reply with ONLY a JSON object:\n"
    '{"explanation": "<2-3 sentences: what clashes and the trade-off>",'
    ' "recommend_index": <integer index into candidates, or -1 if none is good>,'
    ' "action": "<move|approve>",'
    ' "ranking": [{"index": <int>, "why": "<short>"}],'
    ' "confidence": "<high|medium|low>",'
    ' "caveat": "<risk to double-check, or empty>"}'
)


def conflict_instructions():
    return (_settings().get("ai_instructions") or "").strip()


def suggest_resolution(payload):
    """payload: {booking:{...}, conflicts:[...], candidates:[{index,label,...}]}."""
    cands = payload.get("candidates") or []
    if not cands:
        return {"ok": False, "error": "no_candidates",
                "message": "The solver found no legal alternative slot to move this to."}
    model = strong_model()
    instr = conflict_instructions()
    system = _SUGGEST_SYS + (f"\n\nThe scheduler's house rules (follow these):\n{instr}" if instr else "")

    def run():
        user = json.dumps(payload, ensure_ascii=False, indent=2)
        out = _json_block(_call(model, system, "Rank these options:\n" + user, max_tokens=900))
        if not out:
            return {"ok": False, "error": "parse", "message": "AI reply was not valid JSON."}
        out.update(ok=True, source="conflict + comment + legal candidates", model=model)
        return out

    # cache key includes model + instructions so a settings change re-asks
    return _cached("suggest", {**payload, "_m": model, "_i": instr}, run)
