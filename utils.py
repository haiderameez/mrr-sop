# utils.py
import pandas as pd
import re
import numpy as np
from datetime import datetime, timedelta
from config import TARGET_YEAR, TARGET_MONTH

def extract_all(inv):
    nums = re.findall(r'\d+', str(inv))
    cleaned = [n[-3:] for n in nums if len(n) >= 3]
    return ['INV-25-26-000' + d for d in cleaned]

def get_month_start_end(month, year):
    start = datetime(year, month, 1)
    if month == 12:
        end = datetime(year, 12, 31)
    else:
        end = datetime(year, month + 1, 1) - timedelta(days=1)
    return start, end

def normalize_name(s):
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = re.sub(r"[^\w\s]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.lower()

def parse_token_month_year(token):
    token = token.strip().lower().replace('.', '')
    m = re.search(r'([a-z]+)', token)
    y = re.search(r'(\d{2,4})', token)
    mm = None
    yy = None
    if m:
        key = m.group(1)[:3]
        mm = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
              'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}.get(key)
    if y:
        val = y.group(1)
        yy = 2000 + int(val) if len(val) == 2 else int(val)
    return mm, yy

def months_list_from_field(v):
    if pd.isna(v):
        return []
    if isinstance(v, (pd.Timestamp, datetime)):
        return [v.month]
    s = str(v).strip().lower()
    if 'to' in s:
        parts = re.split(r'\s+to\s+', s)
        if len(parts) != 2:
            return []
        a, b = parts
        sm, sy = parse_token_month_year(a)
        em, ey = parse_token_month_year(b)
        if not sm or not em:
            return []
        if not sy:
            sy = TARGET_YEAR
        if not ey:
            ey = sy
        start = datetime(sy, sm, 1)
        end = datetime(ey, em, 28) + timedelta(days=4)
        end = end.replace(day=1) - timedelta(days=1)
        cur = start
        out = []
        while cur <= end:
            out.append(cur.month)
            ny = cur.year
            nm = cur.month + 1
            if nm == 13:
                nm = 1
                ny += 1
            cur = cur.replace(year=ny, month=nm, day=1)
        return out
    if '&' in s or ',' in s:
        tokens = re.split(r'[,&]', s)
        out = []
        for t in tokens:
            mm, yy = parse_token_month_year(t)
            if mm:
                out.append(mm)
        return sorted(set(out))
    m = re.findall(r'[A-Za-z]+', s)
    out = []
    for t in m:
        mm = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
              'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}.get(t[:3])
        if mm:
            out.append(mm)
    return sorted(set(out))

def overlap_days(start, end, period_start, period_end):
    s = max(start, period_start)
    e = min(end, period_end)
    if e < s:
        return 0
    return (e - s).days + 1

def safe_round(x):
    if pd.isna(x) or x == '':
        return x
    try:
        return round(float(x), 3)
    except (ValueError, TypeError):
        return x