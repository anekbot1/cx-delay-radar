#!/usr/bin/env python3
"""
Enrich FOP delay data with historical METAR weather observations.

Usage:
    python enrich_weather.py input.xlsx [output.xlsx]

Fetches METARs from Iowa State IEM archive for each departure/arrival
station at the scheduled time, parses into numeric features, and writes
an enriched Excel file ready for modelling.
"""

import sys
import time
import re
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    sys.exit("Missing openpyxl: pip install openpyxl")

try:
    import requests
except ImportError:
    sys.exit("Missing requests: pip install requests")

try:
    import pandas as pd
except ImportError:
    sys.exit("Missing pandas: pip install pandas")


# ── METAR parsing ──

def parse_metar(raw):
    """Parse a raw METAR string into numeric features."""
    feat = {
        'wind_dir': None, 'wind_kt': None, 'gust_kt': None,
        'vis_m': None, 'ceiling_ft': None,
        'temp_c': None, 'dewpt_c': None, 'qnh_hpa': None,
        'has_rain': 0, 'has_ts': 0, 'has_snow': 0, 'has_fog': 0,
        'has_cb': 0, 'has_tcu': 0,
    }
    if not raw:
        return feat

    # Wind: 36010G20KT or VRB05KT
    m = re.search(r'(\d{3}|VRB)(\d{2,3})(?:G(\d{2,3}))?KT', raw)
    if m:
        feat['wind_dir'] = None if m.group(1) == 'VRB' else int(m.group(1))
        feat['wind_kt'] = int(m.group(2))
        feat['gust_kt'] = int(m.group(3)) if m.group(3) else None

    # Visibility in meters (e.g., 9999, 0800) or SM (e.g., 10SM, 3SM, 1/2SM)
    m = re.search(r'\b(\d{4})\b', raw)
    if m and m.group(1) not in raw.split()[0:2]:  # avoid matching time
        vis_candidate = int(m.group(1))
        if vis_candidate <= 9999:
            feat['vis_m'] = vis_candidate
    m_sm = re.search(r'(\d+)?(?:\s+)?(\d/\d)?SM', raw)
    if m_sm:
        vis = 0
        if m_sm.group(1):
            vis += int(m_sm.group(1))
        if m_sm.group(2):
            n, d = m_sm.group(2).split('/')
            vis += int(n) / int(d)
        feat['vis_m'] = int(vis * 1609.34)  # convert SM to meters

    # Ceiling: lowest BKN or OVC layer
    for cm in re.finditer(r'(BKN|OVC)(\d{3})', raw):
        alt = int(cm.group(2)) * 100
        if feat['ceiling_ft'] is None or alt < feat['ceiling_ft']:
            feat['ceiling_ft'] = alt

    # Temp/dewpoint: M05/M08 or 25/20
    m = re.search(r'\b(M?\d{2})/(M?\d{2})\b', raw)
    if m:
        def parse_temp(s):
            return -int(s[1:]) if s.startswith('M') else int(s)
        feat['temp_c'] = parse_temp(m.group(1))
        feat['dewpt_c'] = parse_temp(m.group(2))

    # QNH
    m = re.search(r'Q(\d{4})', raw)
    if m:
        feat['qnh_hpa'] = int(m.group(1))
    m = re.search(r'A(\d{4})', raw)
    if m:
        feat['qnh_hpa'] = round(int(m.group(1)) * 0.338639, 0)

    # Weather phenomena
    wx = raw.upper()
    if 'RA' in wx or 'DZ' in wx or 'SH' in wx:
        feat['has_rain'] = 1
    if 'TS' in wx:
        feat['has_ts'] = 1
    if 'SN' in wx or 'SG' in wx or 'GR' in wx:
        feat['has_snow'] = 1
    if 'FG' in wx or 'BR' in wx or 'HZ' in wx:
        feat['has_fog'] = 1
    if 'CB' in wx:
        feat['has_cb'] = 1
    if 'TCU' in wx:
        feat['has_tcu'] = 1

    return feat


# ── IEM METAR fetch ──

IEM_URL = "https://mesonet.agron.iastate.edu/cgi-bin/request/asos.py"

def fetch_metars_for_station(station, start_date, end_date):
    """Fetch all METARs for a station in a date range from IEM.
    Returns dict: {(date, hour): raw_metar_string}
    """
    params = {
        'station': station,
        'data': 'metar',
        'tz': 'Etc/UTC',
        'format': 'onlycomma',
        'latlon': 'no',
        'elev': 'no',
        'missing': 'empty',
        'trace': 'empty',
        'direct': 'no',
        'report_type': '3',  # METAR + SPECI
        'year1': start_date.year, 'month1': start_date.month, 'day1': start_date.day,
        'year2': end_date.year, 'month2': end_date.month, 'day2': end_date.day,
    }

    for attempt in range(3):
        try:
            resp = requests.get(IEM_URL, params=params, timeout=60)
            if resp.status_code == 200:
                break
        except requests.RequestException:
            pass
        time.sleep(2 ** attempt)
    else:
        print(f"  WARNING: Failed to fetch METARs for {station}")
        return {}

    # Parse CSV: station,valid,metar
    metars = {}
    lines = resp.text.strip().split('\n')
    for line in lines[1:]:  # skip header
        parts = line.split(',', 2)
        if len(parts) < 3:
            continue
        try:
            dt = datetime.strptime(parts[1].strip(), '%Y-%m-%d %H:%M')
            key = (dt.strftime('%Y-%m-%d'), dt.hour)
            # Keep the latest METAR for each hour
            metars[key] = parts[2].strip()
        except ValueError:
            continue

    return metars


def resolve_col(row, candidates, default=''):
    """Try multiple column names, return first match."""
    for c in candidates:
        if c in row and row[c] is not None and str(row[c]).strip() != '':
            return row[c]
    return default


def parse_datetime(val):
    """Parse a datetime value from Excel (could be datetime obj or string)."""
    if isinstance(val, datetime):
        return val
    if isinstance(val, str):
        for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%d/%m/%Y %H:%M', '%Y-%m-%dT%H:%M:%S']:
            try:
                return datetime.strptime(val.strip(), fmt)
            except ValueError:
                continue
    return None


def main():
    if len(sys.argv) < 2:
        print("Usage: python enrich_weather.py input.xlsx [output.xlsx]")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2]) if len(sys.argv) > 2 else input_path.with_stem(input_path.stem + '_wx')

    print(f"Reading {input_path}...")
    df = pd.read_excel(input_path)
    print(f"  {len(df)} rows, {len(df.columns)} columns")

    # Resolve column names
    dep_col = next((c for c in ['DA', 'Departure'] if c in df.columns), None)
    arr_col = next((c for c in ['AA', 'Arrival'] if c in df.columns), None)
    std_col = next((c for c in ['STD (UTC)', 'STD'] if c in df.columns), None)
    sta_col = next((c for c in ['STA (UTC)', 'STA'] if c in df.columns), None)

    if not dep_col or not arr_col:
        sys.exit("Cannot find departure (DA) or arrival (AA) columns")
    if not std_col and not sta_col:
        sys.exit("Cannot find STD or STA columns")

    # Collect unique station + date ranges
    stations = set()
    for _, row in df.iterrows():
        dep = str(row.get(dep_col, '')).strip()
        arr = str(row.get(arr_col, '')).strip()
        if dep:
            stations.add(dep)
        if arr:
            stations.add(arr)

    # Find date range
    dates = []
    for col in [std_col, sta_col]:
        if not col:
            continue
        for val in df[col]:
            dt = parse_datetime(val)
            if dt:
                dates.append(dt)

    if not dates:
        sys.exit("Cannot parse any dates from STD/STA columns")

    min_date = min(dates) - timedelta(days=1)
    max_date = max(dates) + timedelta(days=1)

    print(f"\nStations: {sorted(stations)}")
    print(f"Date range: {min_date.strftime('%Y-%m-%d')} to {max_date.strftime('%Y-%m-%d')}")
    print(f"Fetching METARs for {len(stations)} stations...")

    # Fetch METARs for each station
    station_metars = {}
    for i, stn in enumerate(sorted(stations)):
        print(f"  [{i+1}/{len(stations)}] Fetching {stn}...", end='', flush=True)
        metars = fetch_metars_for_station(stn, min_date, max_date)
        station_metars[stn] = metars
        print(f" {len(metars)} observations")
        if i < len(stations) - 1:
            time.sleep(1)  # be nice to IEM

    # Enrich each row
    print("\nEnriching flight data...")
    wx_features = [
        'wind_dir', 'wind_kt', 'gust_kt', 'vis_m', 'ceiling_ft',
        'temp_c', 'dewpt_c', 'qnh_hpa',
        'has_rain', 'has_ts', 'has_snow', 'has_fog', 'has_cb', 'has_tcu'
    ]

    # Create new columns
    for prefix in ['dep_wx_', 'arr_wx_']:
        for feat in wx_features:
            df[prefix + feat] = None

    matched_dep = 0
    matched_arr = 0

    for idx, row in df.iterrows():
        # Departure weather
        dep = str(row.get(dep_col, '')).strip()
        if std_col and dep:
            std = parse_datetime(row.get(std_col))
            if std and dep in station_metars:
                key = (std.strftime('%Y-%m-%d'), std.hour)
                raw = station_metars[dep].get(key, '')
                if raw:
                    matched_dep += 1
                    feat = parse_metar(raw)
                    for f in wx_features:
                        df.at[idx, 'dep_wx_' + f] = feat[f]

        # Arrival weather
        arr = str(row.get(arr_col, '')).strip()
        if sta_col and arr:
            sta = parse_datetime(row.get(sta_col))
            if sta and arr in station_metars:
                key = (sta.strftime('%Y-%m-%d'), sta.hour)
                raw = station_metars[arr].get(key, '')
                if raw:
                    matched_arr += 1
                    feat = parse_metar(raw)
                    for f in wx_features:
                        df.at[idx, 'arr_wx_' + f] = feat[f]

    print(f"\nMatch rate:")
    print(f"  Departure: {matched_dep}/{len(df)} ({matched_dep/len(df)*100:.1f}%)")
    print(f"  Arrival:   {matched_arr}/{len(df)} ({matched_arr/len(df)*100:.1f}%)")

    # Save
    print(f"\nSaving to {output_path}...")
    df.to_excel(output_path, index=False)
    print("Done!")

    # Summary stats
    print("\n── Weather Feature Summary ──")
    for prefix, label in [('dep_wx_', 'Departure'), ('arr_wx_', 'Arrival')]:
        print(f"\n{label}:")
        for feat in ['wind_kt', 'gust_kt', 'vis_m', 'ceiling_ft', 'has_rain', 'has_ts', 'has_fog']:
            col = prefix + feat
            non_null = df[col].notna().sum()
            if non_null > 0:
                if feat.startswith('has_'):
                    pct = df[col].sum() / non_null * 100
                    print(f"  {feat}: {int(df[col].sum())} occurrences ({pct:.1f}%)")
                else:
                    print(f"  {feat}: mean={df[col].mean():.1f}, max={df[col].max():.0f}")


if __name__ == '__main__':
    main()
