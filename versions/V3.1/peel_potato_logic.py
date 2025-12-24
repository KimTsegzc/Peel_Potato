"""
Pure logic helpers for Peel Potato.
No dependencies on Excel or PyQt — suitable for unit tests.
"""
import re


def col_letter_to_index(letter: str):
    """Convert Excel column letters like 'A' or 'AA' to 1-based index.
    Returns None for invalid input.
    """
    if letter is None:
        return None
    letter = letter.strip().upper()
    if not letter.isalpha():
        return None
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def expand_column_range(cols_str: str):
    """Expand a column specification into a list of column letters.
    Accepts forms like 'B,E' or 'B:E' or 'B:C,E'.
    Returns list of uppercase column letters.
    """
    if not cols_str:
        return []
    cols = []
    cols_str = cols_str.strip().upper()
    # continuous range like B:E (and no comma)
    if ':' in cols_str and ',' not in cols_str:
        parts = [p.strip() for p in cols_str.split(':')]
        if len(parts) == 2:
            start_idx = col_letter_to_index(parts[0])
            end_idx = col_letter_to_index(parts[1])
            if start_idx and end_idx and start_idx <= end_idx:
                for i in range(start_idx, end_idx + 1):
                    col = ''
                    idx = i
                    while idx > 0:
                        idx, rem = divmod(idx - 1, 26)
                        col = chr(65 + rem) + col
                    cols.append(col)
        return cols

    # comma-separated parts
    for part in cols_str.split(','):
        part = part.strip()
        if not part:
            continue
        if ':' in part:
            sub = [s.strip() for s in part.split(':')]
            if len(sub) == 2:
                s_idx = col_letter_to_index(sub[0])
                e_idx = col_letter_to_index(sub[1])
                if s_idx and e_idx and s_idx <= e_idx:
                    for i in range(s_idx, e_idx + 1):
                        col = ''
                        idx = i
                        while idx > 0:
                            idx, rem = divmod(idx - 1, 26)
                            col = chr(65 + rem) + col
                        cols.append(col)
        else:
            if part.isalpha():
                cols.append(part)
    return cols


_cartesian_re = re.compile(r'\(([^)]+)\)\s*\*\s*\(([^)]+)\)')

def parse_cartesian_spec(values_text: str):
    """Parse cartesian product specification like '(B,E)*(2,7)'.
    Returns (cols_list, rows_list) or None if not matched.
    """
    if not values_text:
        return None
    m = _cartesian_re.match(values_text.strip())
    if not m:
        return None
    cols_str = m.group(1).strip()
    rows_str = m.group(2).strip()

    cols = expand_column_range(cols_str)

    # parse rows
    rows = []
    if ':' in rows_str and ',' not in rows_str:
        parts = rows_str.split(':')
        try:
            s = int(parts[0].strip())
            e = int(parts[1].strip())
            rows = list(range(s, e + 1))
        except Exception:
            rows = []
    else:
        for part in rows_str.split(','):
            part = part.strip()
            if not part:
                continue
            if ':' in part:
                sub = part.split(':')
                try:
                    s = int(sub[0].strip())
                    e = int(sub[1].strip())
                    rows.extend(range(s, e + 1))
                except Exception:
                    continue
            else:
                try:
                    rows.append(int(part))
                except Exception:
                    continue

    if cols and rows:
        return (cols, rows)
    return None
