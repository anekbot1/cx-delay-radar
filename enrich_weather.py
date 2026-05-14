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
import json
import sqlite3
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


# ── IATA to ICAO mapping (CX network + common destinations) ──

IATA_TO_ICAO = {
    # Hong Kong hub
    'HKG': 'VHHH',
    # Greater China
    'PEK': 'ZBAA', 'PKX': 'ZBAD', 'PVG': 'ZSPD', 'SHA': 'ZSSS',
    'CAN': 'ZGGG', 'SZX': 'ZGSZ', 'CTU': 'ZUUU', 'TFU': 'ZUTF',
    'CKG': 'ZUCK', 'WUH': 'ZHHH', 'NKG': 'ZSNJ', 'HGH': 'ZSHC',
    'XMN': 'ZSAM', 'FOC': 'ZSFZ', 'TAO': 'ZSQD', 'DLC': 'ZYTL',
    'SHE': 'ZYTX', 'HRB': 'ZYHB', 'CSX': 'ZGHA', 'KMG': 'ZPPP',
    'XIY': 'ZLXY', 'TSN': 'ZBTJ', 'CGO': 'ZHCC', 'HAK': 'ZJHK',
    'NNG': 'ZGNN', 'KWE': 'ZUGY', 'KWL': 'ZGKL', 'URC': 'ZWWW',
    'LHW': 'ZLLL', 'HET': 'ZBHH', 'ZUH': 'ZGSD', 'WNZ': 'ZSWZ',
    'NGB': 'ZSNB', 'TNA': 'ZSJN', 'SJW': 'ZBSJ', 'TYN': 'ZBYN',
    'INC': 'ZLIC', 'XNN': 'ZLXN', 'JJN': 'ZSQZ', 'YNT': 'ZSYT',
    'CZX': 'ZSCG', 'HFE': 'ZSOF', 'KHN': 'ZSCN', 'NTG': 'ZSNT',
    'LYA': 'ZHLY', 'CGQ': 'ZYCC', 'MDG': 'ZYMD', 'LJG': 'ZPLJ',
    'SYX': 'ZJSY', 'YIH': 'ZHYC', 'ZHA': 'ZGZJ', 'LXA': 'ZULS',
    'NBO': 'HKJK',
    # Taiwan
    'TPE': 'RCTP', 'KHH': 'RCKH', 'RMQ': 'RCMQ',
    # Japan
    'NRT': 'RJAA', 'HND': 'RJTT', 'KIX': 'RJBB', 'NGO': 'RJGG',
    'FUK': 'RJFF', 'CTS': 'RJCC', 'OKA': 'ROAH',
    # Korea
    'ICN': 'RKSI', 'GMP': 'RKSS', 'PUS': 'RKPK',
    # Southeast Asia
    'SIN': 'WSSS', 'BKK': 'VTBS', 'KUL': 'WMKK', 'MNL': 'RPLL',
    'HKT': 'VTSP', 'CNX': 'VTCC', 'USM': 'VTSM', 'HDY': 'VTSS',
    'SGN': 'VVTS', 'HAN': 'VVNB', 'DAD': 'VVDN', 'PNH': 'VDPP',
    'RGN': 'VYYY', 'CGK': 'WIII', 'DPS': 'WADD', 'SUB': 'WARR',
    'REP': 'VDSR', 'VTE': 'VLVT', 'LPQ': 'VLLB',
    # South Asia
    'DEL': 'VIDP', 'BOM': 'VABB', 'MAA': 'VOMM', 'BLR': 'VOBL',
    'HYD': 'VOHS', 'CCU': 'VECC', 'CMB': 'VCBI', 'DAC': 'VGHS',
    'KTM': 'VNKT', 'MLE': 'VRMM',
    # Middle East
    'DXB': 'OMDB', 'AUH': 'OMAA', 'DOH': 'OTHH', 'BAH': 'OBBI',
    'RUH': 'OERK', 'JED': 'OEJN', 'TLV': 'LLBG',
    # Australia & NZ
    'SYD': 'YSSY', 'MEL': 'YMML', 'BNE': 'YBBN', 'PER': 'YPPH',
    'ADL': 'YPAD', 'CNS': 'YBCS', 'AKL': 'NZAA', 'CHC': 'NZCH',
    'WLG': 'NZWN',
    # Europe
    'LHR': 'EGLL', 'LGW': 'EGKK', 'MAN': 'EGCC', 'CDG': 'LFPG',
    'AMS': 'EHAM', 'FRA': 'EDDF', 'MUC': 'EDDM', 'FCO': 'LIRF',
    'MXP': 'LIMC', 'MAD': 'LEMD', 'BCN': 'LEBL', 'ZRH': 'LSZH',
    'VIE': 'LOWW', 'CPH': 'EKCH', 'ARN': 'ESSA', 'HEL': 'EFHK',
    'IST': 'LTFM', 'ATH': 'LGAV', 'DUB': 'EIDW', 'BRU': 'EBBR',
    'LIS': 'LPPT',
    # North America
    'JFK': 'KJFK', 'LAX': 'KLAX', 'SFO': 'KSFO', 'ORD': 'KORD',
    'EWR': 'KEWR', 'BOS': 'KBOS', 'IAD': 'KIAD', 'DFW': 'KDFW',
    'YVR': 'CYVR', 'YYZ': 'CYYZ',
    # Africa
    'JNB': 'FAOR', 'CPT': 'FACT',
    # Pacific
    'HNL': 'PHNL',
    # South America
    'GRU': 'SBGR', 'EZE': 'SAEZ', 'SCL': 'SCEL', 'BOG': 'SKBO',
    'MEX': 'MMMX',
    # CX regional / other
    'CEB': 'RPVM', 'CRK': 'RPLC', 'PEN': 'WMKP', 'LGK': 'WMKL',
    'KCH': 'WBGG', 'BKI': 'WBKK', 'BWN': 'WBSB',
    'AFW': 'KAFW',
}

# Cache file for ICAO lookups we resolve at runtime
ICAO_CACHE_FILE = Path(__file__).parent / '.icao_cache.json'


def load_icao_cache():
    if ICAO_CACHE_FILE.exists():
        with open(ICAO_CACHE_FILE) as f:
            return json.load(f)
    return {}


def save_icao_cache(cache):
    with open(ICAO_CACHE_FILE, 'w') as f:
        json.dump(cache, f, indent=2)


# ── SQLite METAR cache ──

METAR_DB_FILE = Path(__file__).parent / 'metar_cache.db'


def init_metar_db():
    """Create the METAR cache database and table if needed."""
    conn = sqlite3.connect(str(METAR_DB_FILE))
    conn.execute('''
        CREATE TABLE IF NOT EXISTS metars (
            station TEXT NOT NULL,
            obs_date TEXT NOT NULL,
            obs_hour INTEGER NOT NULL,
            raw_metar TEXT NOT NULL,
            fetched_at TEXT NOT NULL,
            PRIMARY KEY (station, obs_date, obs_hour)
        )
    ''')
    conn.execute('''
        CREATE INDEX IF NOT EXISTS idx_metars_station_date
        ON metars (station, obs_date)
    ''')
    conn.commit()
    return conn


def get_cached_metars(conn, station, dates):
    """Get cached METARs for a station and set of dates.
    Returns dict: {(date_str, hour): raw_metar}
    """
    if not dates:
        return {}
    metars = {}
    # Query in batches to avoid huge IN clauses
    date_list = sorted(dates)
    for i in range(0, len(date_list), 100):
        batch = date_list[i:i+100]
        placeholders = ','.join('?' * len(batch))
        cursor = conn.execute(
            f'SELECT obs_date, obs_hour, raw_metar FROM metars WHERE station = ? AND obs_date IN ({placeholders})',
            [station] + batch
        )
        for row in cursor:
            metars[(row[0], row[1])] = row[2]
    return metars


def get_cached_date_range(conn, station):
    """Return set of dates we already have cached for a station."""
    cursor = conn.execute(
        'SELECT DISTINCT obs_date FROM metars WHERE station = ?',
        [station]
    )
    return {row[0] for row in cursor}


def store_metars(conn, station, metars):
    """Store fetched METARs into the cache DB."""
    now = datetime.utcnow().isoformat()
    rows = [(station, date_str, hour, raw, now) for (date_str, hour), raw in metars.items()]
    conn.executemany(
        'INSERT OR REPLACE INTO metars (station, obs_date, obs_hour, raw_metar, fetched_at) VALUES (?, ?, ?, ?, ?)',
        rows
    )
    conn.commit()


def iata_to_icao(iata, cache):
    """Convert IATA code to ICAO. Uses built-in map, then cache, then API fallback."""
    iata = iata.upper().strip()
    if iata in IATA_TO_ICAO:
        return IATA_TO_ICAO[iata]
    if iata in cache:
        return cache[iata] if cache[iata] else None

    # Try a heuristic: if it's already 4 chars, might be ICAO
    if len(iata) == 4:
        cache[iata] = iata
        return iata

    # Fallback: query IEM with common prefixes
    # Many ICAO codes are just a prefix + IATA
    for prefix in ['K', 'C', 'E', 'L', 'R', 'Z', 'V', 'W', 'Y', 'N', 'O', 'S', 'F', 'H', 'D']:
        candidate = prefix + iata
        try:
            resp = requests.get(IEM_URL, params={
                'station': candidate, 'data': 'metar', 'tz': 'Etc/UTC',
                'format': 'onlycomma', 'latlon': 'no', 'elev': 'no',
                'missing': 'empty', 'trace': 'empty', 'direct': 'no',
                'report_type': '3',
                'year1': 2025, 'month1': 1, 'day1': 1,
                'year2': 2025, 'month2': 1, 'day2': 2,
            }, timeout=15)
            lines = resp.text.strip().split('\n')
            if len(lines) > 1:
                cache[iata] = candidate
                return candidate
        except requests.RequestException:
            pass
        time.sleep(0.3)

    print(f"  WARNING: No ICAO code found for {iata}")
    cache[iata] = None
    return None


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
    if m and m.group(1) not in raw.split()[0:2]:
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
        feat['vis_m'] = int(vis * 1609.34)

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
    Returns dict: {(date_str, hour): raw_metar_string}
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
        'report_type': '3',
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
        print(f" FAILED")
        return {}

    metars = {}
    lines = resp.text.strip().split('\n')
    for line in lines[1:]:
        parts = line.split(',', 2)
        if len(parts) < 3:
            continue
        try:
            dt = datetime.strptime(parts[1].strip(), '%Y-%m-%d %H:%M')
            key = (dt.strftime('%Y-%m-%d'), dt.hour)
            metars[key] = parts[2].strip()
        except ValueError:
            continue

    return metars


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

    # Collect unique IATA stations
    iata_stations = set()
    for _, row in df.iterrows():
        dep = str(row.get(dep_col, '')).strip()
        arr = str(row.get(arr_col, '')).strip()
        if dep:
            iata_stations.add(dep)
        if arr:
            iata_stations.add(arr)

    # Convert IATA to ICAO
    print(f"\nResolving {len(iata_stations)} station codes (IATA → ICAO)...")
    icao_cache = load_icao_cache()
    iata_icao_map = {}
    for stn in sorted(iata_stations):
        icao = iata_to_icao(stn, icao_cache)
        if icao:
            iata_icao_map[stn] = icao
            print(f"  {stn} → {icao}")
        else:
            print(f"  {stn} → NOT FOUND (will skip)")
    save_icao_cache(icao_cache)

    unmapped = iata_stations - set(iata_icao_map.keys())
    if unmapped:
        print(f"\n  WARNING: {len(unmapped)} stations unmapped: {sorted(unmapped)}")

    # Build per-station date requirements (only fetch dates we need)
    station_dates = defaultdict(set)  # icao -> set of (date_str, hour)
    for _, row in df.iterrows():
        dep = str(row.get(dep_col, '')).strip()
        arr = str(row.get(arr_col, '')).strip()
        if std_col and dep and dep in iata_icao_map:
            std = parse_datetime(row.get(std_col))
            if std:
                station_dates[iata_icao_map[dep]].add(std.strftime('%Y-%m-%d'))
        if sta_col and arr and arr in iata_icao_map:
            sta = parse_datetime(row.get(sta_col))
            if sta:
                station_dates[iata_icao_map[arr]].add(sta.strftime('%Y-%m-%d'))

    # Init SQLite METAR cache
    metar_conn = init_metar_db()

    # Fetch METARs per station, checking cache first
    unique_icao = sorted(station_dates.keys())
    print(f"\nLoading METARs for {len(unique_icao)} stations...")

    station_metars = {}  # icao -> {(date_str, hour): metar}
    total_cached = 0
    total_fetched = 0

    for i, icao in enumerate(unique_icao):
        needed_dates = sorted(station_dates[icao])

        # Check what we already have cached
        cached_dates = get_cached_date_range(metar_conn, icao)
        missing_dates = [d for d in needed_dates if d not in cached_dates]

        # Load cached data
        cached_metars = get_cached_metars(metar_conn, icao, needed_dates)
        total_cached += len(cached_metars)

        if missing_dates:
            # Fetch only the missing date range from IEM
            min_d = datetime.strptime(missing_dates[0], '%Y-%m-%d')
            max_d = datetime.strptime(missing_dates[-1], '%Y-%m-%d') + timedelta(days=1)
            n_days = (max_d - min_d).days
            print(f"  [{i+1}/{len(unique_icao)}] {icao}: {len(cached_metars)} cached, fetching {n_days} new days...", end='', flush=True)
            new_metars = fetch_metars_for_station(icao, min_d, max_d)
            if new_metars:
                store_metars(metar_conn, icao, new_metars)
                total_fetched += len(new_metars)
            print(f" {len(new_metars)} new obs")
            # Merge cached + new
            cached_metars.update(new_metars)
            if i < len(unique_icao) - 1:
                time.sleep(0.5)
        else:
            print(f"  [{i+1}/{len(unique_icao)}] {icao}: {len(cached_metars)} cached (all dates covered)")

        station_metars[icao] = cached_metars

    print(f"\n  Cache: {total_cached} from DB, {total_fetched} newly fetched")

    # Enrich each row
    print("\nEnriching flight data...")
    wx_features = [
        'wind_dir', 'wind_kt', 'gust_kt', 'vis_m', 'ceiling_ft',
        'temp_c', 'dewpt_c', 'qnh_hpa',
        'has_rain', 'has_ts', 'has_snow', 'has_fog', 'has_cb', 'has_tcu'
    ]

    for prefix in ['dep_wx_', 'arr_wx_']:
        for feat in wx_features:
            df[prefix + feat] = None

    matched_dep = 0
    matched_arr = 0

    for idx, row in df.iterrows():
        # Departure weather
        dep = str(row.get(dep_col, '')).strip()
        if std_col and dep and dep in iata_icao_map:
            icao = iata_icao_map[dep]
            std = parse_datetime(row.get(std_col))
            if std and icao in station_metars:
                key = (std.strftime('%Y-%m-%d'), std.hour)
                raw = station_metars[icao].get(key, '')
                if raw:
                    matched_dep += 1
                    feat = parse_metar(raw)
                    for f in wx_features:
                        df.at[idx, 'dep_wx_' + f] = feat[f]

        # Arrival weather
        arr = str(row.get(arr_col, '')).strip()
        if sta_col and arr and arr in iata_icao_map:
            icao = iata_icao_map[arr]
            sta = parse_datetime(row.get(sta_col))
            if sta and icao in station_metars:
                key = (sta.strftime('%Y-%m-%d'), sta.hour)
                raw = station_metars[icao].get(key, '')
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
