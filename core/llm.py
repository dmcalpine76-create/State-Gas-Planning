"""
core/llm.py — one Anthropic client, one retry wrapper, one JSON parser.

Replaces three diverged copies of _api_call_with_retry and eight call-sites
of ad-hoc markdown-fence stripping. The robust parser (fence stripping,
outermost-block extraction, truncation repair) previously existed only in
board_briefing.py; every tool now gets it.
"""

import os
import re
import json
import time

import anthropic

# Model used for every call, overridable without touching code:
#   set BOARD_TOOLS_MODEL in .env or the environment.
DEFAULT_MODEL = os.environ.get("BOARD_TOOLS_MODEL", "claude-sonnet-4-6")


def get_client() -> "anthropic.Anthropic":
    return anthropic.Anthropic()


def api_call_with_retry(client, max_retries: int = 3,
                        retry_delay: float = 8.0, **kwargs):
    """
    client.messages.create() with backoff. Retries on connection errors,
    timeouts, rate limits, and 5xx. Honours the API's retry-after header
    when present; otherwise 65s for rate limits, linear backoff for the rest.
    """
    kwargs.setdefault("model", DEFAULT_MODEL)
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            return client.messages.create(**kwargs)
        except Exception as e:
            last_exc = e
            err_str  = str(e).lower()
            retriable = any(k in err_str for k in (
                "connection", "timeout", "rate", "429", "500", "502", "503", "overloaded"
            ))
            if not retriable or attempt == max_retries:
                raise
            is_rate_limit = "429" in err_str or "rate_limit" in err_str
            wait = 65.0 if is_rate_limit else retry_delay * attempt
            # Prefer the server's own retry-after if it gave one
            resp = getattr(e, "response", None)
            if resp is not None:
                try:
                    ra = resp.headers.get("retry-after")
                    if ra:
                        wait = max(float(ra), 1.0)
                except Exception:
                    pass
            print(f" ⟳ retry {attempt}/{max_retries - 1} in {wait:.0f}s", end=" ", flush=True)
            time.sleep(wait)
    raise last_exc


def robust_json_parse(text: str):
    """
    Parse JSON from a Claude response with progressive repair strategies:
    markdown fences, outermost-block extraction, truncated-response repair.
    Raises json.JSONDecodeError if every strategy fails.
    """
    text = text.strip()
    for pat in (r'^```(?:json)?\s*', r'\s*```$'):
        text = re.sub(pat, '', text, flags=re.MULTILINE)
    text = text.strip()

    # Strategy 1: direct parse
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Strategy 2: extract the outermost block of whichever bracket type the
    # response leads with. (Trying both types unconditionally could return an
    # embedded array from a truncated object, silently dropping the wrapper.)
    first_obj = text.find('{')
    first_arr = text.find('[')
    if first_obj == -1 or (first_arr != -1 and first_arr < first_obj):
        pairs = [('[', ']'), ('{', '}')]
    else:
        pairs = [('{', '}'), ('[', ']')]

    sc, ec = pairs[0]
    start, end = text.find(sc), text.rfind(ec)
    if start != -1 and end > start:
        try:
            return json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            pass

    # Strategy 3: truncated response — keep the last complete top-level value
    starts = [i for i in (first_obj, first_arr) if i != -1]
    lead   = text[min(starts):] if starts else text
    if lead.startswith('{'):
        last_safe = -1
        depth = 0
        for i, c in enumerate(lead):
            if c in '{[':
                depth += 1
            elif c in '}]':
                depth -= 1
                if depth == 1:
                    last_safe = i + 1
        if last_safe > 0:
            truncated = lead[:last_safe].rstrip().rstrip(',')
            try:
                return json.loads(truncated + '\n}')
            except json.JSONDecodeError:
                pass

    # Strategy 4: last resort — the other bracket type's outermost block
    sc, ec = pairs[1]
    start, end = text.find(sc), text.rfind(ec)
    if start != -1 and end > start:
        try:
            return json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            pass

    raise json.JSONDecodeError("All JSON repair strategies failed", text, 0)


def call_json(client, prompt: str, max_tokens: int, label: str = "",
              model: str = None):
    """
    One-shot convenience: API call with retry → robust parse.
    Returns the parsed object, or None (with a printed warning) on failure.
    """
    try:
        msg = api_call_with_retry(
            client,
            model      = model or DEFAULT_MODEL,
            max_tokens = max_tokens,
            messages   = [{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
    except Exception as e:
        print(f"  ⚠️  {label or 'API call'} failed: {e}")
        return None
    try:
        return robust_json_parse(raw)
    except json.JSONDecodeError as e:
        print(f"  ⚠️  {label or 'response'} — could not parse JSON: {e}")
        print(f"      Raw snippet: {raw[:200]}")
        return None
