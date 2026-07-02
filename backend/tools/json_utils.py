"""
tools/json_utils.py
===================
JSON repair utilities for cleaning up LLM output before json.loads().

Public API
----------
fix_json(text: str) -> str
split_json_arrays(text: str) -> list[str]
"""

import re


def fix_json(s: str) -> str:
    """Repair common LLM JSON errors so the string can be parsed by json.loads().

    Fixes applied (in order):
      1. Strip markdown fences if still present
      2. Remove JavaScript-style comments  (// ...)
      3. Remove trailing commas before } or ]
      4. Replace Python literals  None → null, True → true, False → false
    """
    s = re.sub(r"^```[a-z]*\s*", "", s.strip())
    s = re.sub(r"\s*```$", "", s)
    s = re.sub(r"//[^\n]*", "", s)
    s = re.sub(r",\s*([}\]])", r"\1", s)
    s = re.sub(r"\bNone\b",  "null",  s)
    s = re.sub(r"\bTrue\b",  "true",  s)
    s = re.sub(r"\bFalse\b", "false", s)
    return s


def split_json_arrays(text: str) -> list:
    """Extract every top-level JSON array from text, respecting nesting and strings.

    Returns a list of raw array strings.  Used when a vision LLM returns one
    array per record on separate lines instead of one combined array.
    """
    arrays = []
    i = 0
    n = len(text)
    while i < n:
        if text[i] == "[":
            depth = 0
            in_str = False
            escape = False
            j = i
            while j < n:
                c = text[j]
                if escape:
                    escape = False
                elif in_str:
                    if c == "\\":
                        escape = True
                    elif c == '"':
                        in_str = False
                else:
                    if c == '"':
                        in_str = True
                    elif c == "[":
                        depth += 1
                    elif c == "]":
                        depth -= 1
                        if depth == 0:
                            arrays.append(text[i:j + 1])
                            i = j + 1
                            break
                j += 1
            else:
                i += 1   # no matching ']' — skip this '['
        else:
            i += 1
    return arrays
