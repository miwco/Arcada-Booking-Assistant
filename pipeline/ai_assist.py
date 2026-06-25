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

import hashlib
import json
import os
import re
import urllib.error
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
API_URL = "https://api.anthropic.com/v1/messages"
# Cheap, high-volume model for reading comments; strong model for ranking/explaining.
MODEL_CHEAP = os.environ.get("BA_MODEL_CHEAP", "claude-haiku-4-5-20251001")
MODEL_STRONG = os.environ.get("BA_MODEL_STRONG", "claude-opus-4-8")

_CACHE: dict[str, dict] = {}      # prompt hash -> result, avoids repeat billing in a session


def _key():
    k = os.environ.get("ANTHROPIC_API_KEY")
    if k:
        return k.strip()
    path = os.path.join(ROOT, ".env")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if line.startswith("ANTHROPIC_API_KEY"):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    return None


def status():
    return {"available": bool(_key()), "model_cheap": MODEL_CHEAP, "model_strong": MODEL_STRONG}


def _call(model, system, user, max_tokens=900, temperature=0.0):
    """Raw Messages API call. Raises RuntimeError('no_key') / RuntimeError(msg)."""
    key = _key()
    if not key:
        raise RuntimeError("no_key")
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
        out = _json_block(_call(MODEL_CHEAP, _INTERPRET_SYS, user, max_tokens=500))
        if not out:
            return {"ok": False, "error": "parse", "message": "AI reply was not valid JSON."}
        out.update(ok=True, source="Other comments", model=MODEL_CHEAP)
        return out

    return _cached("interpret", {"c": comment, "k": course}, run)


# --- 2. rank/explain the legal resolution candidates the solver found -------
_SUGGEST_SYS = (
    "You advise a scheduler at a film & media school. The deterministic solver has "
    "already enforced every hard rule and given you ONLY legal move options. Your job "
    "is to rank and explain them in plain language — you do not invent new options and "
    "you do not apply anything. Prefer keeping a course on its requested day, honouring "
    "stated time constraints in the comment, minimal disruption, and mid-week days "
    "(Tue-Thu) over Mon/Fri. Reply with ONLY a JSON object:\n"
    '{"explanation": "<2-3 sentences: what clashes and the trade-off>",'
    ' "recommend_index": <integer index into candidates, or -1 if none is good>,'
    ' "ranking": [{"index": <int>, "why": "<short>"}],'
    ' "confidence": "<high|medium|low>",'
    ' "caveat": "<risk to double-check, or empty>"}'
)


def suggest_resolution(payload):
    """payload: {booking:{...}, conflicts:[...], candidates:[{index,label,...}]}."""
    cands = payload.get("candidates") or []
    if not cands:
        return {"ok": False, "error": "no_candidates",
                "message": "The solver found no legal alternative slot to move this to."}

    def run():
        user = json.dumps(payload, ensure_ascii=False, indent=2)
        out = _json_block(_call(MODEL_STRONG, _SUGGEST_SYS,
                                "Rank these options:\n" + user, max_tokens=900))
        if not out:
            return {"ok": False, "error": "parse", "message": "AI reply was not valid JSON."}
        out.update(ok=True, source="conflict + comment + legal candidates", model=MODEL_STRONG)
        return out

    return _cached("suggest", payload, run)
