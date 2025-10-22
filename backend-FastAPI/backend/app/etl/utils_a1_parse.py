"""
utils_a1_parse.py
Helper functions for parsing Excel A1-style cell references and ranges.
"""
import re

def col_to_idx(col):
    """Convert column letters (A, B, AA, etc.) to numeric index (1-based)."""
    col = col.upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

def idx_to_col(idx):
    """Convert numeric column index (1-based) to letters."""
    col = ''
    while idx > 0:
        idx -= 1
        col = chr(ord('A') + idx % 26) + col
        idx //= 26
    return col

# Regex for A1-style address: $?COL$?ROW[:$?COL$?ROW]
ADDR_RE = re.compile(r'^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$')

def parse_a1_address(addr):
    """
    Parse A1-style address into (row1, col1, row2, col2).
    Returns list of tuples for multi-part ranges (comma-separated).
    Single cell returns [(row, col, row, col)].
    """
    addr = str(addr).strip()
    if not addr or addr.lower() == 'nan':
        return []
    parts = [p.strip() for p in addr.split(',')]
    ranges = []
    for p in parts:
        m = ADDR_RE.match(p)
        if not m:
            continue
        c1 = col_to_idx(m.group(1))
        r1 = int(m.group(2))
        c2 = c1 if m.group(3) is None else col_to_idx(m.group(3))
        r2 = r1 if m.group(4) is None else int(m.group(4))
        ranges.append((min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)))
    return ranges

def expand_range(addr):
    """
    Expand A1 range into list of (row, col) tuples.
    Example: 'A1:B2' -> [(1,1), (1,2), (2,1), (2,2)]
    """
    cells = []
    for r1, c1, r2, c2 in parse_a1_address(addr):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cells.append((r, c))
    return cells

def range_size(addr):
    """Count total cells in an A1 range."""
    total = 0
    for r1, c1, r2, c2 in parse_a1_address(addr):
        total += (r2 - r1 + 1) * (c2 - c1 + 1)
    return total
