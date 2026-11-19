"""Excel number formats rendered as PowerPoint shows them, with no python-pptx, so
layout can measure value labels before a chart exists."""


def first_section(fmt):
    out, quoted = [], False
    for ch in fmt:
        if ch == '"':
            quoted = not quoted
        if ch == ";" and not quoted:
            break
        out.append(ch)
    return "".join(out)


def _tokens(section):
    """(is_literal, text) pieces of one number-format section."""
    out, i = [], 0
    while i < len(section):
        ch = section[i]
        if ch == '"':
            end = section.find('"', i + 1)
            end = len(section) if end < 0 else end
            out.append((True, section[i + 1:end]))
            i = end + 1
        elif ch == "\\" and i + 1 < len(section):
            out.append((True, section[i + 1]))
            i += 2
        else:
            out.append((False, ch))
            i += 1
    return out


def format_number(value, fmt):
    """Render a value the way PowerPoint will show it, for alt text and labels.

    Covers what decks use: literal prefixes and suffixes, decimals, thousands
    separators, percent and trailing-comma scaling. Dates and conditional
    sections are out of scope; they fall back to plain numbers.
    """
    if not fmt or fmt.strip().lower() == "general":
        return f"{value:,.10g}" if abs(value) >= 10000 else f"{value:.10g}"
    toks = _tokens(first_section(fmt))
    code = [k for k, (lit, t) in enumerate(toks) if not lit and t in "0#?"]
    if not code:
        return f"{value:g}"
    first, last = code[0], code[-1]
    prefix = "".join(t for _, t in toks[:first])
    middle = "".join(t for _, t in toks[first:last + 1])
    tail = toks[last + 1:]
    scale = sum(1 for lit, t in tail if not lit and t == ",")
    suffix = "".join(t for lit, t in tail if lit or t != ",")
    percent = any(not lit and t == "%" for lit, t in toks)
    decimals = len(middle.split(".", 1)[1]) if "." in middle else 0
    v = value * (100 if percent else 1) / (1000 ** scale)
    body = f"{abs(v):,.{decimals}f}" if "," in middle.split(".")[0] else f"{abs(v):.{decimals}f}"
    negative = v < 0 and float(body.replace(",", "")) != 0
    return ("-" if negative else "") + prefix + body + suffix
