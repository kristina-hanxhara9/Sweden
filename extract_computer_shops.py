#!/usr/bin/env python3
"""
Extract Swedish Computer/Electronics Shops from government bulk register files.

Data sources:
- SCB bulk file (SNI codes, addresses, legal form)
- Bolagsverket bulk file (business descriptions)
- SCB Branschnyckeltal API (turnover estimation)

Both bulk files are free EU High Value Datasets from Bolagsverket, CC-BY-4.0.
"""

import pandas as pd
import requests
import re
import datetime
import os
import sys
import io
import zipfile
import json

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

SCB_URL = "https://vardefulla-datamangder.bolagsverket.se/scb/scb_bulkfil.zip"
BV_URL = "https://vardefulla-datamangder.bolagsverket.se/bolagsverket/bolagsverket_bulkfil.zip"
SCB_API_URL = "https://api.scb.se/OV0104/v1/doris/sv/ssd/START/NV/NV0109/NV0109O/BNTT01"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MANUAL_STORES_FILE = os.path.join(SCRIPT_DIR, "manual_stores.csv")

OUTPUT_FILE = os.path.join(SCRIPT_DIR, "swedish_computer_shops.csv")
MANUAL_OUTPUT_FILE = os.path.join(SCRIPT_DIR, "swedish_computer_shops_manual.csv")
COMBINED_OUTPUT_FILE = os.path.join(SCRIPT_DIR, "swedish_computer_shops_combined.csv")

TODAY = datetime.date.today().isoformat()

# ---------------------------------------------------------------------------
# SNI code sets
# ---------------------------------------------------------------------------

PRIMARY_SNI = {47401, 47402, 47403}  # Retail computers, software, consumer electronics
SECONDARY_SNI = set()  # Wholesale removed — retail only
TERTIARY_SNI = {95101, 95102, 47430}  # Repair + audio/video (need name confirmation)
BROAD_RETAIL_SNI = {47112, 47122, 47123}  # General retail (only for known chain matching)
# 47404 is EXCLUDED from standalone filtering
# 46501, 46502 (wholesale) are EXCLUDED — retail shops only

ALL_STANDALONE_SNI = PRIMARY_SNI | SECONDARY_SNI  # Retail only
ALL_SNI_FOR_REFERENCE = ALL_STANDALONE_SNI | TERTIARY_SNI | {47404, 46501, 46502} | BROAD_RETAIL_SNI

SNI_DESCRIPTIONS = {
    47401: "Retail: computers & peripherals",
    47402: "Retail: software",
    47403: "Retail: consumer electronics",
    47404: "Retail: telecom equipment",
    47430: "Retail: audio/video equipment",
    46501: "Wholesale: computers & peripherals",
    46502: "Wholesale: electronic components",
    95101: "Repair: computers",
    95102: "Repair: telecom equipment",
    47112: "General retail: mainly food",
    47122: "General retail: non-food",
    47123: "Internet retail: broad assortment",
}

SHOP_CATEGORY_MAP = {
    47401: "Computer & Peripherals",
    47402: "Software",
    47403: "Consumer Electronics",
    47404: "Telecom Equipment",
    47430: "Audio/Video",
    46501: "IT Distribution (Wholesale)",
    46502: "Electronic Components (Wholesale)",
    95101: "Computer Repair/Service",
    95102: "Telecom Repair/Service",
    47112: "General Retail",
    47122: "General Retail",
    47123: "Internet Retail",
}

LEGAL_FORM_MAP = {
    "5": ("AB", "Aktiebolag"),
    "7": ("EkFor", "Ekonomisk forening"),
    "8": ("IdFor", "Ideell forening / Stiftelse"),
    "9": ("HB/KB", "Handelsbolag / Kommanditbolag"),
    "2": ("Stat", "Statlig / kommunal"),
}

# ---------------------------------------------------------------------------
# Known entities (verified org numbers, HIGH confidence)
# ---------------------------------------------------------------------------

KNOWN_ENTITIES = {
    "556471447-4": {"name": "Elgiganten", "chain": "Elgiganten", "type": "chain", "parent": "Elkjop Nordic AS", "ultimate_owner": "Currys plc (UK)", "listed": "listed"},
    "556558822-4": {"name": "Webhallen", "chain": "Komplett Group", "type": "chain", "parent": "Komplett Group", "ultimate_owner": "Komplett Group", "listed": "private"},
    "556520413-7": {"name": "NetOnNet", "chain": "Komplett Group", "type": "chain", "parent": "Komplett Group", "ultimate_owner": "Komplett Group", "listed": "private"},
    "556581064-4": {"name": "Komplett Services Sweden", "chain": "Komplett Group", "type": "chain", "parent": "Komplett Group", "ultimate_owner": "Komplett Group", "listed": "private"},
    "556599350-7": {"name": "Komplett Distribution", "chain": "Komplett Group", "type": "chain", "parent": "Komplett Group", "ultimate_owner": "Komplett Group", "listed": "private"},
    "559503261-5": {"name": "Komplett Business Nordic", "chain": "Komplett Group", "type": "chain", "parent": "Komplett Group", "ultimate_owner": "Komplett Group", "listed": "private"},
    "556237878-5": {"name": "Dustin", "chain": "Dustin Group", "type": "chain", "parent": "Dustin Group AB", "ultimate_owner": "Dustin Group AB", "listed": "listed"},
    "556666101-2": {"name": "Dustin Sverige", "chain": "Dustin Group", "type": "chain", "parent": "Dustin Group AB", "ultimate_owner": "Dustin Group AB", "listed": "listed"},
    "556703306-2": {"name": "Dustin Group", "chain": "Dustin Group", "type": "chain", "parent": "Dustin Group AB", "ultimate_owner": "Dustin Group AB", "listed": "listed"},
    "556400537-8": {"name": "Kjell & Co Elektronik", "chain": "Kjell Group", "type": "chain", "parent": "Kjell Group AB", "ultimate_owner": "Kjell Group AB", "listed": "listed"},
    "556412064-9": {"name": "Kjell & Company", "chain": "Kjell Group", "type": "chain", "parent": "Kjell Group AB", "ultimate_owner": "Kjell Group AB", "listed": "listed"},
    "559115844-8": {"name": "Kjell Group", "chain": "Kjell Group", "type": "chain", "parent": "Kjell Group AB", "ultimate_owner": "Kjell Group AB", "listed": "listed"},
    "556591888-4": {"name": "Inet", "chain": "Inet", "type": "independent_chain", "parent": "Inet Group AB", "ultimate_owner": "Inet Group AB", "listed": "private"},
    "559008431-4": {"name": "Inet Group", "chain": "Inet", "type": "independent_chain", "parent": "Inet Group AB", "ultimate_owner": "Inet Group AB", "listed": "private"},
    "556687521-6": {"name": "Power Sverige", "chain": "Power International", "type": "chain", "parent": "Expert ASA (Norway)", "ultimate_owner": "Expert ASA (Norway)", "listed": "private"},
    "559071846-5": {"name": "Power Retail Sweden", "chain": "Power International", "type": "chain", "parent": "Expert ASA (Norway)", "ultimate_owner": "Expert ASA (Norway)", "listed": "private"},
    "556112974-2": {"name": "SIBA", "chain": "Komplett Group (legacy SIBA)", "type": "chain_legacy", "parent": "Komplett Group", "ultimate_owner": "Komplett Group", "listed": "private"},
    "556120221-8": {"name": "HiFi Klubben", "chain": "HiFi Klubben", "type": "chain", "parent": "HiFi Klubben A/S (Denmark)", "ultimate_owner": "HiFi Klubben A/S", "listed": "private"},
    "502074529-4": {"name": "Proshop", "chain": "Proshop", "type": "chain", "parent": "Proshop (Denmark)", "ultimate_owner": "Proshop", "listed": "private"},
    "556065405-4": {"name": "Elon AB", "chain": "Elon Group", "type": "buying_group", "parent": "Elon Group AB", "ultimate_owner": "Elon Group AB", "listed": "private"},
    "556470980-5": {"name": "Elon Drift", "chain": "Elon Group", "type": "buying_group", "parent": "Elon Group AB", "ultimate_owner": "Elon Group AB", "listed": "private"},
    "556492607-8": {"name": "Elon Logistics", "chain": "Elon Group", "type": "buying_group", "parent": "Elon Group AB", "ultimate_owner": "Elon Group AB", "listed": "private"},
    "556518176-4": {"name": "ALSO Sweden", "chain": "ALSO Group", "type": "it_distributor", "parent": "ALSO Holding AG", "ultimate_owner": "ALSO Holding AG (Switzerland)", "listed": "listed"},
    "556254845-2": {"name": "Ingram Micro", "chain": "Ingram Micro", "type": "it_distributor", "parent": "Ingram Micro Inc.", "ultimate_owner": "Platinum Equity (USA)", "listed": "private"},
}

# ---------------------------------------------------------------------------
# Chain name patterns (Level 2 — MEDIUM confidence)
# ---------------------------------------------------------------------------

CHAIN_NAME_PATTERNS = [
    (r"^elgiganten\b", "Elgiganten", "chain", "Elkjop Nordic AS", "Currys plc (UK)"),
    (r"^webhallen\b", "Komplett Group", "chain", "Komplett Group", "Komplett Group"),
    (r"^netonnet\b|^net\s*on\s*net\b", "Komplett Group", "chain", "Komplett Group", "Komplett Group"),
    (r"^komplett\s+(?!bygg|bemanning|stad|service\s+i\s+)",
     "Komplett Group", "chain", "Komplett Group", "Komplett Group"),
    (r"^dustin\s+(?!service|städ|clean)",
     "Dustin Group", "chain", "Dustin Group AB", "Dustin Group AB"),
    (r"^kjell\s*[&]\s*co", "Kjell Group", "chain", "Kjell Group AB", "Kjell Group AB"),
    (r"^inet\s*(?:ab|group)?\s*$", "Inet", "independent_chain", "Inet Group AB", "Inet Group AB"),
    (r"^power\s+(?:sverige|retail\s+sweden)", "Power International", "chain",
     "Expert ASA (Norway)", "Expert ASA (Norway)"),
    (r"^atea\s+(?:sverige|logistics|it|ab)", "Atea", "chain", "Atea ASA", "Atea ASA"),
    (r"^advania\s+(?:sverige|ab)", "Advania", "chain", "Advania Group", "Advania Group"),
    (r"^td\s*synnex\b", "TD Synnex", "it_distributor", "TD Synnex Corp", "TD Synnex Corp"),
    (r"^elon\s+(?!musk)", "Elon Group", "buying_group", "Elon Group AB", "Elon Group AB"),
    (r"^hifi\s*klubben\b", "HiFi Klubben", "chain", "HiFi Klubben A/S", "HiFi Klubben A/S"),
    (r"^proshop\b", "Proshop", "chain", "Proshop (Denmark)", "Proshop"),
    (r"^siba\s+(?:ab|aktiebolag)", "Komplett Group (legacy SIBA)", "chain_legacy",
     "Komplett Group", "Komplett Group"),
    (r"\beuronics\b", "Euronics", "buying_group",
     "Euronics International", "Euronics International (Netherlands)"),
]

# ---------------------------------------------------------------------------
# Chain keywords — explicit 4 keywords per chain (separate from SNI search)
# ---------------------------------------------------------------------------
# Each company matching any of these keywords will be added to the output,
# even if it does NOT have a computer-related SNI code. Match is tagged with
# match_method = "chain_keyword" and classification_method = "chain_keyword"

CHAIN_KEYWORDS = {
    "Elgiganten": {
        "keywords": ["elgiganten", "elkjop", "elkjøp", "giganten"],
        "chain_type": "chain",
        "parent": "Elkjop Nordic AS",
        "ultimate_owner": "Currys plc (UK)",
    },
    "Komplett Group": {
        "keywords": ["komplett", "webhallen", "netonnet", "net on net"],
        "chain_type": "chain",
        "parent": "Komplett Group",
        "ultimate_owner": "Komplett Group",
    },
    "Dustin Group": {
        "keywords": ["dustin", "dustin group", "dustin sverige", "dustin ab"],
        "chain_type": "chain",
        "parent": "Dustin Group AB",
        "ultimate_owner": "Dustin Group AB",
    },
    "Kjell Group": {
        "keywords": ["kjell & co", "kjell company", "kjell group", "kjell elektronik"],
        "chain_type": "chain",
        "parent": "Kjell Group AB",
        "ultimate_owner": "Kjell Group AB",
    },
    "Inet": {
        "keywords": ["inet ab", "inet group", "inet.se", "inet sverige"],
        "chain_type": "independent_chain",
        "parent": "Inet Group AB",
        "ultimate_owner": "Inet Group AB",
    },
    "Power International": {
        "keywords": ["power sverige", "power retail", "power international", "power nordic"],
        "chain_type": "chain",
        "parent": "Expert ASA (Norway)",
        "ultimate_owner": "Expert ASA (Norway)",
    },
    "MediaMarkt": {
        "keywords": ["mediamarkt", "media markt", "ceconomy", "mediamarket"],
        "chain_type": "chain_legacy",
        "parent": "Ceconomy AG",
        "ultimate_owner": "Ceconomy AG (Germany)",
    },
    "SIBA": {
        "keywords": ["siba ab", "siba aktiebolag", "siba elektronik", "siba invest"],
        "chain_type": "chain_legacy",
        "parent": "Komplett Group",
        "ultimate_owner": "Komplett Group",
    },
    "HiFi Klubben": {
        "keywords": ["hifi klubben", "hi-fi klubben", "hifi klub", "hifiklubben"],
        "chain_type": "chain",
        "parent": "HiFi Klubben A/S",
        "ultimate_owner": "HiFi Klubben A/S (Denmark)",
    },
    "Proshop": {
        "keywords": ["proshop", "proshop.se", "proshop sweden", "proshop nordic"],
        "chain_type": "chain",
        "parent": "Proshop A/S",
        "ultimate_owner": "Proshop (Denmark)",
    },
    "Euronics": {
        "keywords": ["euronics", "euronics sweden", "euronics nordic", "euronics international"],
        "chain_type": "buying_group",
        "parent": "Euronics International",
        "ultimate_owner": "Euronics International (Netherlands)",
    },
    "Expert": {
        "keywords": ["expert sverige", "expert nordic", "expert norden", "expert butik"],
        "chain_type": "buying_group",
        "parent": "Expert ASA",
        "ultimate_owner": "Expert ASA (Norway)",
    },
    "Elon Group": {
        "keywords": ["elon group", "elon sverige", "elon ljud", "audio video"],
        "chain_type": "buying_group",
        "parent": "Elon Group AB",
        "ultimate_owner": "Elon Group AB",
    },
    "ALSO Group": {
        "keywords": ["also sweden", "also nordic", "also holding", "also ab"],
        "chain_type": "it_distributor",
        "parent": "ALSO Holding AG",
        "ultimate_owner": "ALSO Holding AG (Switzerland)",
    },
    "Ingram Micro": {
        "keywords": ["ingram micro", "ingrammicro", "ingram sweden", "ingram nordic"],
        "chain_type": "it_distributor",
        "parent": "Ingram Micro Inc.",
        "ultimate_owner": "Platinum Equity (USA)",
    },
    "TD Synnex": {
        "keywords": ["td synnex", "techdata", "tech data", "synnex sweden"],
        "chain_type": "it_distributor",
        "parent": "TD Synnex Corp",
        "ultimate_owner": "TD Synnex Corp",
    },
    "Atea": {
        "keywords": ["atea sverige", "atea logistics", "atea it", "atea ab"],
        "chain_type": "chain",
        "parent": "Atea ASA",
        "ultimate_owner": "Atea ASA (Norway)",
    },
    "Advania": {
        "keywords": ["advania sverige", "advania ab", "advania nordic", "advania data"],
        "chain_type": "chain",
        "parent": "Advania Group",
        "ultimate_owner": "Advania Group",
    },
}

# ---------------------------------------------------------------------------
# Exclusion patterns — filter out non-retail entities
# ---------------------------------------------------------------------------

EXCLUDE_NAME_PATTERNS = [
    r"\bholding\b",
    r"\bförvaltning",
    r"\binvest(?:ment)?s?\b",
    r"\bfastighet",
    r"\bhuvudkontor\b",
    r"\bhead\s*office\b",
    r"\bhuawei\b",
    r"\bsamsung\s*(?:electronics|experience)\b",
    r"\bxiaomi\b",
    r"\boppo\b",
    r"\bvivo\s*mobile\b",
    r"\btelia\s",
    r"\btele2\s",
    r"\btelenor\s",
    r"\btre\s+(?:sverige|ab)\b",
    r"\bcomviq\b",
    r"\bphone\s*house\b",
    r"\bonly\s*phones\b",
    r"\bmobiltelefon",
    r"\bmobil\s*(?:butik|shop|center)\b",
    r"\blogistic(?:s|k)\b",
    r"\btransport\b",
]

# Keywords for name-matching computer shops
NAME_KEYWORDS = [
    r"\bdatorbutik", r"\bdatorhandel", r"\bdatorservice\b", r"\bdatorf[öo]rs[äa]ljning",
    r"\bcomputer\s*(?:shop|store|center)\b", r"\bit[-\s]butik\b",
    r"\belektronikhandel\b", r"\bgaming\s*(?:shop|store|butik)\b", r"\bdatorutrustning\b",
]

DESCRIPTION_KEYWORDS = [
    r"f[öo]rs[äa]ljning\s+av\s+dator", r"detaljhandel\s+med\s+dator",
    r"dator.*tillbeh[öo]r", r"datorutrustning", r"partihandel\s+med\s+dator",
    r"butikshandel\s+med\s+dator", r"f[öo]rs[äa]ljning\s+av\s+it",
]

ATTRIBUTION_HEADER = """# Swedish Computer Shops Database
# Generated: {date}
# License: CC-BY-4.0
# Attribution: Contains data from Bolagsverket and SCB, licensed under CC-BY-4.0
#   as EU High Value Datasets.
# Sources:
#   SCB bulk file: {scb_url}
#   Bolagsverket bulk file: {bv_url}
#   SCB Branschnyckeltal API (turnover estimates)
# Note: Sole traders excluded for privacy.
#   Turnover estimates are industry medians by SNI code, not company-specific.
#   Only active companies included. Headquarters and holding companies excluded.
"""


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def normalize_org_nr(pe_org_nr):
    """Convert 12-digit PeOrgNr (e.g. 165564714474) to 10-digit dash format (556471447-4)."""
    s = str(int(pe_org_nr)).zfill(12)
    ten = s[2:]  # strip "16" prefix
    return f"{ten[:9]}-{ten[9]}"


def org_nr_to_legal_form(org_nr_10):
    """Derive legal form from first digit of 10-digit org number."""
    first = org_nr_10[0] if org_nr_10 else ""
    return LEGAL_FORM_MAP.get(first, ("Other", "Other"))


def is_excluded(name):
    """Check if a company name matches any exclusion pattern."""
    if pd.isna(name):
        return False
    name_str = str(name).strip()
    for pat in EXCLUDE_NAME_PATTERNS:
        if re.search(pat, name_str, re.IGNORECASE):
            # Exception: Elgiganten Phonehouse is OK
            if "phonehouse" in pat.lower() and "elgiganten" in name_str.lower():
                continue
            return True
    return False


def classify_company(pe_org_nr, name, kw_chain=None, kw_type=None, kw_parent=None,
                     kw_owner=None, kw_matched_keyword=None):
    """4-level chain classification. Returns dict with chain info.

    Priority (highest first):
    1. Org number lookup (HIGH)
    2. Name pattern regex (MEDIUM)
    3. Chain keyword match (MEDIUM) — passed in from filter_by_chain_keywords
    4. Independent (default)
    """
    result = {
        "chain_group": "Independent", "chain_type": "Independent",
        "parent_company": "", "ultimate_owner": "",
        "classification_confidence": "N/A", "classification_method": "default",
        "matched_keyword": "",
    }
    # Level 1: Org number lookup (HIGH)
    org_10 = normalize_org_nr(pe_org_nr) if pd.notna(pe_org_nr) else ""
    if org_10 in KNOWN_ENTITIES:
        ent = KNOWN_ENTITIES[org_10]
        result.update({
            "chain_group": ent["chain"], "chain_type": ent["type"],
            "parent_company": ent["parent"], "ultimate_owner": ent["ultimate_owner"],
            "classification_confidence": "high", "classification_method": "org_number",
        })
        return result
    # Level 2: Name pattern (MEDIUM)
    if pd.notna(name):
        name_str = str(name).strip()
        for pat, chain, ctype, parent, owner in CHAIN_NAME_PATTERNS:
            if re.search(pat, name_str, re.IGNORECASE):
                result.update({
                    "chain_group": chain, "chain_type": ctype,
                    "parent_company": parent, "ultimate_owner": owner,
                    "classification_confidence": "medium", "classification_method": "name_pattern",
                })
                return result
    # Level 3: Chain keyword match (MEDIUM) — from filter_by_chain_keywords
    if kw_chain and pd.notna(kw_chain):
        result.update({
            "chain_group": kw_chain, "chain_type": kw_type or "chain",
            "parent_company": kw_parent or "", "ultimate_owner": kw_owner or "",
            "classification_confidence": "medium", "classification_method": "chain_keyword",
            "matched_keyword": kw_matched_keyword or "",
        })
        return result
    # Level 4: Independent (default)
    return result


def b2c_b2b_heuristic(sni_code):
    """Classify B2C/B2B based on SNI prefix."""
    try:
        code = int(sni_code)
    except (ValueError, TypeError):
        return ""
    if 47000 <= code < 48000:
        return "B2C"
    elif 46000 <= code < 47000:
        return "B2B"
    elif 95000 <= code < 96000:
        return "Service"
    return ""


# ---------------------------------------------------------------------------
# Data download functions
# ---------------------------------------------------------------------------

def download_scb_data():
    """Download and parse the SCB bulk file."""
    print("Downloading SCB bulk file (~100MB)...")
    df = pd.read_csv(SCB_URL, sep="\t", encoding="ISO-8859-1",
                     low_memory=False, compression="zip")
    print(f"  Loaded {len(df):,} companies from SCB")
    print(f"  SCB columns: {list(df.columns)}")
    for col in df.columns:
        sample = df[col].dropna().head(1).tolist()
        print(f"    {col}: {sample}")
    df = df[df["PeOrgNr"] < 190000000000].copy()
    print(f"  After removing sole traders: {len(df):,}")
    return df


def download_bv_data():
    """Download and parse the Bolagsverket bulk file."""
    print("Downloading Bolagsverket bulk file...")
    df = pd.read_csv(BV_URL, sep=";", encoding="utf-8", low_memory=False,
                     quotechar='"', escapechar="\\", on_bad_lines="warn",
                     compression="zip")
    print(f"  Loaded {len(df):,} entries from Bolagsverket")
    print(f"  BV columns: {list(df.columns)}")
    for col in df.columns:
        sample = df[col].dropna().head(2).tolist()
        print(f"    {col}: {sample}")
    df = df[~df["organisationsidentitet"].str.contains(
        r"\$PERSON-IDORG", regex=True, na=False)].copy()
    df["organisationsidentitet"] = (
        df["organisationsidentitet"].str.replace("$ORGNR-IDORG", "", regex=False))
    df["organisationsnamn"] = df["organisationsnamn"].str.split("$FORETAGSNAMN").str[0]
    if "postadress" in df.columns:
        df["postadress"] = (df["postadress"].fillna("")
            .str.replace("$SE-LAND", "", regex=False)
            .str.replace("$$", ", ", regex=False)
            .str.replace("$", ", ", regex=False))
    print(f"  After cleaning: {len(df):,} entries")
    return df


def fetch_turnover_stats():
    """Fetch median turnover by SNI code from SCB Branschnyckeltal API."""
    print("Fetching turnover estimates from SCB API...")
    try:
        meta = requests.get(SCB_API_URL, timeout=30).json()
        print(f"  API variables: {[v['code'] for v in meta.get('variables', [])]}")
        sni_var = None
        for v in meta.get("variables", []):
            if "SNI" in v.get("code", "").upper():
                sni_var = v["code"]
        if not sni_var:
            print("  WARNING: No SNI variable found in API")
            return {}
        sni_strs = [str(c) for c in sorted(ALL_STANDALONE_SNI | TERTIARY_SNI)]
        query = {
            "query": [
                {"code": sni_var, "selection": {"filter": "item", "values": sni_strs}},
            ],
            "response": {"format": "json"},
        }
        resp = requests.post(SCB_API_URL, json=query, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        turnover = {}
        for entry in data.get("data", []):
            keys = entry.get("key", [])
            values = entry.get("values", [])
            if keys and values:
                try:
                    turnover[int(keys[0])] = int(float(values[0]) * 1000)
                except (ValueError, TypeError):
                    pass
        print(f"  Turnover estimates for {len(turnover)} SNI codes: {turnover}")
        return turnover
    except Exception as e:
        print(f"  WARNING: Turnover API failed: {e}")
        return {}


# ---------------------------------------------------------------------------
# Filtering functions
# ---------------------------------------------------------------------------

def get_sni_columns(df):
    """Find all SNI code columns."""
    return sorted([c for c in df.columns if re.match(r"^[Nn]g\d+$", c)])


def find_name_column(df):
    """Find the company name column."""
    for c in ["Namn", "namn", "ForetagsNamn"]:
        if c in df.columns:
            return c
    for c in df.columns:
        if "namn" in c.lower():
            return c
    return None


def filter_by_sni(df, sni_codes):
    """Filter by SNI codes across all Ng columns."""
    sni_cols = get_sni_columns(df)
    if not sni_cols:
        print("  WARNING: No SNI columns found!")
        return pd.DataFrame()
    print(f"  SNI columns: {sni_cols}")
    mask = pd.Series(False, index=df.index)
    for col in sni_cols:
        mask = mask | pd.to_numeric(df[col], errors="coerce").isin(sni_codes)
    return df[mask].copy()


def filter_tertiary_sni(df, name_col):
    """Filter tertiary SNI — only keep if name confirms computer-related."""
    sni_cols = get_sni_columns(df)
    if not sni_cols or not name_col:
        return pd.DataFrame()
    mask = pd.Series(False, index=df.index)
    for col in sni_cols:
        mask = mask | pd.to_numeric(df[col], errors="coerce").isin(TERTIARY_SNI)
    tertiary = df[mask].copy()
    kw = "|".join(NAME_KEYWORDS + [r"\bdator", r"\bcomputer", r"\bPC\b",
                                    r"\bit[-\s]", r"\belektronik", r"\bgaming"])
    return tertiary[tertiary[name_col].fillna("").str.contains(kw, case=False, regex=True)].copy()


def filter_by_name(df, name_col):
    """Find additional shops by name keywords and chain patterns."""
    kw_pattern = "|".join(NAME_KEYWORDS)
    mask = df[name_col].fillna("").str.contains(kw_pattern, case=False, regex=True)
    for pat, *_ in CHAIN_NAME_PATTERNS:
        mask = mask | df[name_col].fillna("").str.contains(pat, case=False, regex=True)
    return df[mask].copy()


def filter_by_chain_keywords(df, name_col):
    """Filter by explicit chain keywords (separate from SNI filtering).

    Returns a DataFrame with added columns:
      _kw_chain_group, _kw_chain_type, _kw_parent, _kw_ultimate_owner, _kw_matched_keyword
    """
    if not name_col or name_col not in df.columns:
        return pd.DataFrame()
    names = df[name_col].fillna("").str.lower()
    matches = []
    for idx, name in names.items():
        if not name:
            continue
        for chain, info in CHAIN_KEYWORDS.items():
            for kw in info["keywords"]:
                # Use word boundaries for short keywords to avoid false positives
                if len(kw) <= 5:
                    pattern = r"\b" + re.escape(kw) + r"\b"
                    if re.search(pattern, name):
                        matches.append((idx, chain, info, kw))
                        break
                else:
                    if kw in name:
                        matches.append((idx, chain, info, kw))
                        break
            else:
                continue
            break  # first chain wins
    if not matches:
        return pd.DataFrame()
    match_indices = [m[0] for m in matches]
    matched = df.loc[match_indices].copy()
    matched["_kw_chain_group"] = [m[1] for m in matches]
    matched["_kw_chain_type"] = [m[2]["chain_type"] for m in matches]
    matched["_kw_parent"] = [m[2]["parent"] for m in matches]
    matched["_kw_ultimate_owner"] = [m[2]["ultimate_owner"] for m in matches]
    matched["_kw_matched_keyword"] = [m[3] for m in matches]
    return matched


def normalize_for_dedup(name):
    """Normalize store name for dedup: lowercase first word, strip AB/HB."""
    if pd.isna(name):
        return ""
    s = re.sub(r"\s+(ab|hb|kb|aktiebolag)\s*$", "", str(name).strip().lower())
    return s.split()[0] if s else ""


def normalize_postal(pc):
    """Normalize postal code: remove spaces."""
    return str(pc).replace(" ", "").strip() if pd.notna(pc) else ""


def build_output(all_shops, name_col, sni_cols, desc_col, turnover_stats):
    """Build the 34-column output DataFrame."""
    out = pd.DataFrame()
    # Identity
    out["org_number"] = all_shops["PeOrgNr"].apply(
        lambda x: f"{int(x):012d}" if pd.notna(x) else "")
    out["company_name"] = all_shops[name_col] if name_col else ""
    org_10s = out["org_number"].apply(lambda x: x[2:] if len(x) == 12 else x)
    lf = org_10s.apply(org_nr_to_legal_form)
    out["legal_form"] = lf.apply(lambda x: x[0])
    out["legal_form_description"] = lf.apply(lambda x: x[1])
    for c in ["RegDatKtid", "regdatktid"]:
        if c in all_shops.columns:
            out["registration_date"] = all_shops[c]; break
    else:
        out["registration_date"] = ""
    out["status"] = "active"
    # Location
    for out_col, candidates in [("address", ["Gatuadress", "gatuadress"]),
                                 ("co_address", ["COAdress", "coadress"]),
                                 ("postal_code", ["PostNr", "postnr"]),
                                 ("city", ["PostOrt", "postort"]),
                                 ("municipality", ["Kommun", "kommun", "KommunKod"]),
                                 ("county", ["Lan", "lan"])]:
        for c in candidates:
            if c in all_shops.columns:
                out[out_col] = all_shops[c]; break
        else:
            out[out_col] = ""
    # Classification
    if sni_cols:
        out["primary_sni"] = all_shops[sni_cols[0]]
        out["primary_sni_description"] = pd.to_numeric(
            all_shops[sni_cols[0]], errors="coerce").map(SNI_DESCRIPTIONS).fillna("")
        out["all_sni_codes"] = all_shops[sni_cols].apply(
            lambda r: ",".join(str(int(v)) for v in r if pd.notna(v) and v != 0), axis=1)
        out["all_sni_descriptions"] = all_shops[sni_cols].apply(
            lambda r: "; ".join(SNI_DESCRIPTIONS.get(int(v), str(int(v)))
                                for v in r if pd.notna(v) and v != 0), axis=1)
    else:
        for c in ["primary_sni", "primary_sni_description", "all_sni_codes", "all_sni_descriptions"]:
            out[c] = ""
    out["business_description"] = all_shops[desc_col] if desc_col and desc_col in all_shops.columns else ""
    out["match_method"] = all_shops["match_method"]
    out["match_confidence"] = all_shops["match_confidence"]
    # Business type
    out["b2c_b2b"] = out["primary_sni"].apply(b2c_b2b_heuristic)
    out["shop_category"] = pd.to_numeric(out["primary_sni"], errors="coerce").map(SHOP_CATEGORY_MAP).fillna("")
    # Chain/Group
    for c in ["chain_group", "chain_type", "parent_company", "ultimate_owner",
              "classification_confidence", "classification_method", "matched_keyword"]:
        out[c] = all_shops[c] if c in all_shops.columns else ""
    # Financials
    if turnover_stats:
        out["estimated_turnover_sek"] = pd.to_numeric(
            out["primary_sni"], errors="coerce").map(turnover_stats).fillna("")
        out["turnover_source"] = out["estimated_turnover_sek"].apply(
            lambda x: "SCB Branschnyckeltal (industry median)" if x != "" else "")
        out["turnover_year"] = out["estimated_turnover_sek"].apply(
            lambda x: "2024" if x != "" else "")
    else:
        out["estimated_turnover_sek"] = ""
        out["turnover_source"] = ""
        out["turnover_year"] = ""
    # Meta
    out["data_source"] = "SCB + Bolagsverket bulk files"
    out["extraction_date"] = TODAY
    out["data_license"] = "CC-BY-4.0"
    return out


def main():
    print("=" * 60)
    print("Swedish Computer Shop Extractor")
    print(f"Date: {TODAY}")
    print("=" * 60)

    # Step 1: Download SCB
    scb = download_scb_data()
    name_col = find_name_column(scb)
    print(f"Name column: {name_col}")

    # Step 2: Filter primary + secondary SNI
    print("\n--- Primary/secondary SNI filter ---")
    sni_filtered = filter_by_sni(scb, ALL_STANDALONE_SNI)
    sni_filtered["match_method"] = "sni_code"
    sni_filtered["match_confidence"] = "high"
    print(f"  {len(sni_filtered):,} companies")

    # Step 3: Tertiary SNI (name-confirmed only)
    print("\n--- Tertiary SNI (name-confirmed) ---")
    tertiary = filter_tertiary_sni(scb, name_col)
    tertiary = tertiary[~tertiary["PeOrgNr"].isin(sni_filtered["PeOrgNr"])].copy()
    tertiary["match_method"] = "sni_name_confirmed"
    tertiary["match_confidence"] = "medium"
    print(f"  {len(tertiary):,} additional")

    # Step 4: Name matching
    print("\n--- Name matching ---")
    if name_col:
        name_matched = filter_by_name(scb, name_col)
        seen = set(sni_filtered["PeOrgNr"]) | set(tertiary["PeOrgNr"])
        name_matched = name_matched[~name_matched["PeOrgNr"].isin(seen)].copy()
        name_matched["match_method"] = "name_match"
        name_matched["match_confidence"] = "medium"
        print(f"  {len(name_matched):,} additional")
    else:
        name_matched = pd.DataFrame()

    # Step 4b: Chain keyword matching (separate from SNI, explicit 4 keywords per chain)
    print("\n--- Chain keyword matching ---")
    kw_matched = filter_by_chain_keywords(scb, name_col)
    if not kw_matched.empty:
        seen = set(sni_filtered["PeOrgNr"]) | set(tertiary["PeOrgNr"]) | set(name_matched["PeOrgNr"])
        kw_matched = kw_matched[~kw_matched["PeOrgNr"].isin(seen)].copy()
        kw_matched["match_method"] = "chain_keyword"
        kw_matched["match_confidence"] = "medium"
        print(f"  {len(kw_matched):,} additional from chain keywords")
        # Show breakdown by chain
        if len(kw_matched) > 0:
            print(f"  Breakdown by chain:")
            print(kw_matched["_kw_chain_group"].value_counts().to_string())
    else:
        print("  0 additional from chain keywords")

    all_shops = pd.concat([sni_filtered, tertiary, name_matched, kw_matched], ignore_index=True)
    print(f"\nTotal before exclusions: {len(all_shops):,}")

    # Step 5: Exclusion filters
    print("\n--- Exclusion filters ---")
    if name_col and name_col in all_shops.columns:
        excl = all_shops[name_col].apply(is_excluded)
        n_excl = excl.sum()
        if n_excl > 0:
            print(f"  Excluding {n_excl:,} non-retail entities:")
            for _, r in all_shops[excl].head(20).iterrows():
                print(f"    - {r.get(name_col, '?')}")
        all_shops = all_shops[~excl].copy()
    print(f"  After exclusions: {len(all_shops):,}")

    # Step 6: Enrich with BV data
    print("\n--- BV enrichment ---")
    desc_col = None
    try:
        bv = download_bv_data()
        bv["org_nr_numeric"] = pd.to_numeric(bv["organisationsidentitet"].str.strip(), errors="coerce")
        for col in bv.columns:
            if any(kw in col.lower() for kw in ["verksamhet", "beskrivning", "syfte"]):
                desc_col = col
                break
        if not desc_col:
            for col in bv.columns:
                if col not in ["organisationsidentitet", "organisationsnamn", "postadress"]:
                    samples = bv[col].dropna().head(5).tolist()
                    if any(isinstance(s, str) and len(s) > 30 for s in samples):
                        desc_col = col
                        break
        print(f"  Description column: {desc_col or 'NOT FOUND'}")
        bv_cols = ["org_nr_numeric", "organisationsnamn"]
        if desc_col:
            bv_cols.append(desc_col)
        all_shops = all_shops.merge(
            bv[bv_cols].drop_duplicates(subset=["org_nr_numeric"]),
            left_on="PeOrgNr", right_on="org_nr_numeric", how="left", suffixes=("", "_bv"))
    except Exception as e:
        print(f"  WARNING: BV failed: {e}")

    # Step 7: Turnover
    turnover_stats = fetch_turnover_stats()

    # Step 8: Classify
    print("\n--- Chain classification ---")
    cls = [classify_company(
        r["PeOrgNr"], r.get(name_col),
        kw_chain=r.get("_kw_chain_group"),
        kw_type=r.get("_kw_chain_type"),
        kw_parent=r.get("_kw_parent"),
        kw_owner=r.get("_kw_ultimate_owner"),
        kw_matched_keyword=r.get("_kw_matched_keyword"),
    ) for _, r in all_shops.iterrows()]
    cls_df = pd.DataFrame(cls)
    for c in cls_df.columns:
        all_shops[c] = cls_df[c].values

    # Step 9: Build output
    sni_cols = get_sni_columns(all_shops)
    out = build_output(all_shops, name_col, sni_cols, desc_col, turnover_stats)
    out = out.drop_duplicates(subset=["org_number"])
    out = out.sort_values(["chain_type", "chain_group", "company_name"]).reset_index(drop=True)

    # Step 10: Write CSV
    header = ATTRIBUTION_HEADER.format(date=TODAY, scb_url=SCB_URL, bv_url=BV_URL)
    with open(OUTPUT_FILE, "w", encoding="utf-8-sig") as f:
        f.write(header)
        out.to_csv(f, index=False)
    print(f"\nSaved {len(out):,} shops to {OUTPUT_FILE}")

    # Step 11: Manual stores
    print("\n--- Manual stores ---")
    if os.path.exists(MANUAL_STORES_FILE):
        manual = pd.read_csv(MANUAL_STORES_FILE, encoding="utf-8-sig")
        print(f"  Loaded {len(manual):,} entries")
        out["_dp"] = out["postal_code"].apply(normalize_postal)
        out["_dn"] = out["company_name"].apply(normalize_for_dedup)
        manual["_dp"] = manual["Postal Code"].apply(normalize_postal)
        manual["_dn"] = manual["Store Name"].apply(normalize_for_dedup)
        m = manual.merge(out[["_dp", "_dn", "org_number"]], on=["_dp", "_dn"], how="left")
        in_scb = m["org_number"].notna()
        print(f"  {in_scb.sum()} already in SCB (deduplicated)")
        manual_new = manual[~in_scb.values].copy()
        out.drop(columns=["_dp", "_dn"], inplace=True)
        manual_new.drop(columns=["_dp", "_dn"], errors="ignore", inplace=True)
        manual_new.to_csv(MANUAL_OUTPUT_FILE, index=False, encoding="utf-8-sig")
        print(f"  Saved {len(manual_new):,} new manual stores to {MANUAL_OUTPUT_FILE}")
        # Combined
        cm = pd.DataFrame()
        cm["org_number"] = ""
        cm["company_name"] = manual_new["Store Name"]
        cm["address"] = manual_new.get("Address", "")
        cm["postal_code"] = manual_new.get("Postal Code", "")
        cm["city"] = manual_new.get("City", "")
        cm["chain_group"] = manual_new.get("Chain / Group", "")
        cm["chain_type"] = manual_new.get("Type", "")
        cm["parent_company"] = manual_new.get("Parent Company", "")
        cm["data_source"] = manual_new.get("Source", "manual")
        for c in out.columns:
            if c not in cm.columns:
                cm[c] = ""
        out["data_source"] = "SCB + Bolagsverket"
        combined = pd.concat([out, cm[out.columns]], ignore_index=True)
        with open(COMBINED_OUTPUT_FILE, "w", encoding="utf-8-sig") as f:
            f.write(header)
            combined.to_csv(f, index=False)
        print(f"  Saved {len(combined):,} total to {COMBINED_OUTPUT_FILE}")

    # Summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Total shops (SCB): {len(out):,}")
    print(f"\nBy match method:\n{out['match_method'].value_counts().to_string()}")
    print(f"\nBy chain type:\n{out['chain_type'].value_counts().to_string()}")
    print(f"\nBy category:\n{out['shop_category'].value_counts().head(10).to_string()}")
    chains = out[out["chain_type"] != "Independent"]
    if len(chains) > 0:
        print(f"\nChain entities ({len(chains)}):")
        for _, r in chains.iterrows():
            print(f"  {r['company_name']} -> {r['chain_group']} [{r['classification_confidence']}]")
    print(f"\nSample independents (first 20):")
    print(out[out["chain_type"] == "Independent"][
        ["company_name", "city", "primary_sni", "shop_category"]].head(20).to_string())
    print(f"\nColumn fill rates:")
    for c in out.columns:
        filled = (out[c] != "").sum() if out[c].dtype == object else out[c].notna().sum()
        pct = filled / len(out) * 100 if len(out) > 0 else 0
        print(f"  {c}: {pct:.0f}%")


if __name__ == "__main__":
    main()
