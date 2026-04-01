#!/usr/bin/env python3
"""
Extract Swedish Computer Shops from the national business register.

Data sources:
- SCB bulk file (has SNI codes, addresses)
- Bolagsverket bulk file (has business descriptions)

Both are free EU High Value Datasets from Bolagsverket.
"""

import pandas as pd
import re
import sys

# --- Configuration ---

SCB_URL = "https://vardefulla-datamangder.bolagsverket.se/scb/scb_bulkfil.zip"
BV_URL = "https://vardefulla-datamangder.bolagsverket.se/bolagsverket/bolagsverket_bulkfil.zip"

# Final CSV column order — all target database fields
OUTPUT_COLUMNS = [
    # Identity
    "org_number", "company_name", "trading_name", "legal_form",
    "founded_year", "status_active",
    # Location
    "address", "co_address", "postal_code", "city",
    "municipality", "region_lan", "lat", "lng", "nr_of_locations",
    # Classification
    "primary_sni", "secondary_sni", "all_sni_codes",
    "business_description", "b2c_b2b", "online_physical", "specialisation",
    # Financials
    "turnover_sek", "turnover_year", "employees",
    "profit_loss", "credit_rating", "equity",
    # Chain & Group
    "is_chain_member", "chain_name", "chain_group", "chain_type",
    "buying_group", "parent_company", "ultimate_owner", "listed_private",
    # Meta
    "match_method",
]

def find_column(df, candidates, label=None):
    """Find the first matching column name from a list of candidates.
    Also tries case-insensitive matching as a fallback."""
    for c in candidates:
        if c in df.columns:
            return c
    # Fallback: case-insensitive match
    lower_map = {col.lower(): col for col in df.columns}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    if label:
        print(f"  WARNING: Could not find column for {label} (tried: {candidates})")
    return None


# Average annual turnover per employee (SEK) by SNI industry group.
# Source: SCB Företagsstatistik, Swedish averages for small/medium enterprises.
# Used to estimate turnover when actual figures are unavailable.
TURNOVER_PER_EMPLOYEE_SEK = {
    # Retail (SNI 47.x) — ~2.5 MSEK per employee
    47401: 2_500_000, 47402: 2_500_000, 47403: 3_000_000, 47404: 2_800_000,
    # Wholesale (SNI 46.x) — higher, ~4–6 MSEK per employee
    46501: 5_000_000, 46502: 4_500_000,
    # Repair (SNI 95.x) — lower, ~1 MSEK per employee
    95101: 1_000_000, 95102: 1_000_000,
    # Broad retail
    47112: 3_000_000, 47122: 3_500_000, 47123: 4_000_000,
}
# Default for unknown SNI codes
DEFAULT_TURNOVER_PER_EMPLOYEE = 2_500_000


# Swedish 5-digit SNI 2007 codes for computer/electronics
PRIMARY_SNI_CODES = {
    47401,  # Specialiserad butikshandel med datorer och kringutrustning
    47402,  # Specialiserad butikshandel med programvara
}

SECONDARY_SNI_CODES = {
    47403,  # Specialiserad butikshandel med hemelektronik (consumer electronics)
    47404,  # Specialiserad butikshandel med telekommunikationsutrustning
    46501,  # Partihandel med datorer, kringutrustning och programvara
    46502,  # Partihandel med elektroniska komponenter
}

TERTIARY_SNI_CODES = {
    95101,  # Reparation av datorer och kringutrustning
    95102,  # Reparation av kommunikationsutrustning
}

# These broad retail codes are only used for name-matched chains, not standalone filtering
BROAD_RETAIL_CODES = {
    47112,  # Detaljhandel med brett sortiment, övervägande livsmedel/drycker
    47122,  # Detaljhandel med brett sortiment, ej livsmedel (Power/Elgiganten stores)
    47123,  # Internethandel med brett sortiment (NetOnNet etc)
}

ALL_SNI_CODES = PRIMARY_SNI_CODES | SECONDARY_SNI_CODES | TERTIARY_SNI_CODES

# Chain/buying group classification - order matters, first match wins
# Each entry carries ownership metadata for the output CSV.
CHAIN_PATTERNS = [
    {
        "pattern": r"^elgiganten\b",
        "chain_name": "Elgiganten", "group_name": "Elkjop Nordic / Currys plc",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Elkjop Nordic AS", "ultimate_owner": "Currys plc",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"elkj[øo]p\s*nordic",
        "chain_name": "Elkjop Nordic", "group_name": "Elkjop Nordic / Currys plc",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Elkjop Nordic AS", "ultimate_owner": "Currys plc",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"^komplett\s*(?:services?\s*(?:sweden|norge)|distribution|business\s*nordic)",
        "chain_name": "Komplett", "group_name": "Komplett Group",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Komplett Group ASA", "ultimate_owner": "Komplett Group ASA",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"^komplett\.se\b|^komplett\s*(?:sweden|group)\b",
        "chain_name": "Komplett", "group_name": "Komplett Group",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Komplett Group ASA", "ultimate_owner": "Komplett Group ASA",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"^netonnet\b|^net\s*on\s*net\b",
        "chain_name": "NetOnNet", "group_name": "Komplett Group",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Komplett Group ASA", "ultimate_owner": "Komplett Group ASA",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"^webhallen\b",
        "chain_name": "Webhallen", "group_name": "Komplett Group",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Komplett Group ASA", "ultimate_owner": "Komplett Group ASA",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"^dustin\s*(?:aktiebolag|ab|group|sverige|a/s|finland|norway)\b",
        "chain_name": "Dustin", "group_name": "Dustin Group",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Dustin Group AB", "ultimate_owner": "Dustin Group AB",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"^kjell\s*[&]\s*co\b|^kjell\s*group\b",
        "chain_name": "Kjell & Company", "group_name": "Kjell Group",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Kjell Group AB", "ultimate_owner": "Kjell Group AB",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"^inet\s*(?:ab|group)\b",
        "chain_name": "Inet", "group_name": "Inet",
        "chain_type": "Independent Chain", "must_match_sni": False,
        "parent_company": "Inet AB", "ultimate_owner": "Inet AB",
        "buying_group": "", "listed_private": "Private",
    },
    {
        "pattern": r"^power\s*(?:sverige|retail\s*sweden)\s*ab\b",
        "chain_name": "Power", "group_name": "Power International",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Power International AS", "ultimate_owner": "Power International AS",
        "buying_group": "", "listed_private": "Private",
    },
    {
        "pattern": r"^power\s*international\b",
        "chain_name": "Power", "group_name": "Power International",
        "chain_type": "Chain", "must_match_sni": False,
        "parent_company": "Power International AS", "ultimate_owner": "Power International AS",
        "buying_group": "", "listed_private": "Private",
    },
    {
        "pattern": r"^mediamarkt\b|^media\s*markt\b",
        "chain_name": "MediaMarkt", "group_name": "MediaMarkt (exited Sweden)",
        "chain_type": "Chain (legacy)", "must_match_sni": False,
        "parent_company": "MediaMarktSaturn", "ultimate_owner": "Ceconomy AG",
        "buying_group": "", "listed_private": "Listed",
    },
    {
        "pattern": r"^siba\s*(?:aktiebolag|ab|fastigheter|invest)\b",
        "chain_name": "SIBA", "group_name": "NetOnNet / Komplett Group (legacy SIBA)",
        "chain_type": "Chain (legacy)", "must_match_sni": False,
        "parent_company": "Komplett Group ASA", "ultimate_owner": "Komplett Group ASA",
        "buying_group": "", "listed_private": "Listed",
    },
]

# Default metadata for independent (non-chain) companies
_INDEPENDENT_META = {
    "chain_name": "", "group_name": "Independent", "chain_type": "Independent",
    "parent_company": "", "ultimate_owner": "", "buying_group": "", "listed_private": "",
}

# Keywords to find computer shops by name (Swedish + English) - more targeted
NAME_KEYWORDS = [
    r"\bdatorbutik",
    r"\bdatorhandel",
    r"\bdatorservice\b",
    r"\bdatorf[öo]rs[äa]ljning",
    r"\bcomputer\s*(?:shop|store|center|centre)\b",
    r"\bit[-\s]butik\b",
    r"\belektronikhandel\b",
    r"\bgaming\s*(?:shop|store|butik)\b",
    r"\bdatorutrustning\b",
]

# Business description keywords for Bolagsverket data
DESCRIPTION_KEYWORDS = [
    r"f[öo]rs[äa]ljning\s+av\s+dator",
    r"f[öo]rs[äa]ljning\s+av\s+it[-\s]*utrustning",
    r"detaljhandel\s+med\s+dator",
    r"retail.*computer",
    r"dator.*tillbeh[öo]r",
    r"datorutrustning",
    r"partihandel\s+med\s+dator",
    r"butikshandel\s+med\s+dator",
]


def download_scb_data():
    """Download and parse the SCB bulk file."""
    print("Downloading SCB bulk file (~100MB, this may take a few minutes)...")
    df = pd.read_csv(
        SCB_URL,
        sep="\t",
        encoding="ISO-8859-1",
        low_memory=False,
        compression="zip",
    )
    print(f"  Loaded {len(df):,} companies from SCB")
    # Remove sole traders (PeOrgNr starting with 19/20 = personal numbers)
    df = df[df["PeOrgNr"] < 190000000000]
    print(f"  After removing sole traders: {len(df):,} companies")
    return df


def download_bv_data():
    """Download and parse the Bolagsverket bulk file."""
    print("Downloading Bolagsverket bulk file...")
    df = pd.read_csv(
        BV_URL,
        sep=";",
        encoding="utf-8",
        low_memory=False,
        quotechar='"',
        escapechar="\\",
        on_bad_lines="warn",
        compression="zip",
    )
    print(f"  Loaded {len(df):,} entries from Bolagsverket")
    # Remove sole traders
    df = df[~df["organisationsidentitet"].str.contains(r"\$PERSON-IDORG", regex=True, na=False)]
    # Clean org number
    df["organisationsidentitet"] = df["organisationsidentitet"].str.replace("$ORGNR-IDORG", "", regex=False)
    # Clean name — capture trading name before discarding the delimiter
    name_parts = df["organisationsnamn"].str.split("$FORETAGSNAMN")
    df["organisationsnamn"] = name_parts.str[0]
    df["trading_name"] = name_parts.str[1]  # NaN if no trading name
    # Clean address
    if "postadress" in df.columns:
        df["postadress"] = (
            df["postadress"]
            .fillna("")
            .str.replace("$SE-LAND", "", regex=False)
            .str.replace("$$", ", ", regex=False)
            .str.replace("$", ", ", regex=False)
        )
    print(f"  After removing sole traders: {len(df):,} entries")
    return df


def get_sni_columns(df):
    """Find all SNI code columns in the dataframe."""
    sni_cols = [c for c in df.columns if c.startswith("Ng") and c[2:].isdigit()]
    if not sni_cols:
        sni_cols = [c for c in df.columns if re.match(r"^[Nn]g\d+$", c)]
    return sorted(sni_cols)


def filter_by_sni(df, sni_codes):
    """Filter dataframe by SNI codes across all SNI columns."""
    sni_cols = get_sni_columns(df)
    if not sni_cols:
        print("  WARNING: No SNI columns found!")
        print(f"  Available columns: {list(df.columns)}")
        return pd.DataFrame()

    print(f"  SNI columns found: {sni_cols}")
    mask = pd.Series(False, index=df.index)
    for col in sni_cols:
        col_numeric = pd.to_numeric(df[col], errors="coerce")
        mask = mask | col_numeric.isin(sni_codes)
    return df[mask].copy()


def filter_by_name(df, name_col):
    """Filter dataframe by company name keywords and known chains."""
    # Match specific keywords
    keyword_pattern = "|".join(NAME_KEYWORDS)
    mask = df[name_col].fillna("").str.contains(keyword_pattern, case=False, regex=True)

    # Match known chain names (precise patterns)
    for entry in CHAIN_PATTERNS:
        mask = mask | df[name_col].fillna("").str.contains(entry["pattern"], case=False, regex=True)

    return df[mask].copy()


def classify_chain(name):
    """Classify a company by chain/buying group based on name. Returns metadata dict."""
    if pd.isna(name):
        return dict(_INDEPENDENT_META)
    name_clean = str(name).strip()
    for entry in CHAIN_PATTERNS:
        if re.search(entry["pattern"], name_clean, re.IGNORECASE):
            return {k: v for k, v in entry.items() if k != "pattern" and k != "must_match_sni"}
    return dict(_INDEPENDENT_META)


def find_name_column(df):
    """Find the company name column in the dataframe."""
    for candidate in ["Namn", "namn", "ForetagsNamn", "foretagsnamn", "Företagsnamn"]:
        if candidate in df.columns:
            return candidate
    for col in df.columns:
        if "namn" in col.lower() or "name" in col.lower():
            return col
    return None


def main():
    print("=" * 60)
    print("Swedish Computer Shop Extractor")
    print("=" * 60)

    # Step 1: Download SCB data
    scb = download_scb_data()
    print(f"\nSCB columns: {list(scb.columns)}")

    name_col = find_name_column(scb)
    print(f"Company name column: {name_col}")

    # Step 2: Filter by SNI codes (primary method)
    print("\n--- Filtering by SNI codes ---")
    sni_filtered = filter_by_sni(scb, ALL_SNI_CODES)
    print(f"  Found {len(sni_filtered):,} companies matching SNI codes")

    # Show breakdown by SNI
    sni_cols = get_sni_columns(scb)
    if sni_cols:
        primary_sni = sni_filtered[sni_cols[0]]
        print(f"\n  SNI code breakdown (primary SNI):")
        for code in sorted(ALL_SNI_CODES):
            count = (primary_sni == code).sum()
            if count > 0:
                print(f"    {code}: {count} companies")

    sni_filtered["match_method"] = "sni_code"

    # Step 3: Filter by name (to catch known chains with different SNI codes)
    print("\n--- Filtering by company name ---")
    if name_col:
        print(f"  Using name column: {name_col}")
        name_filtered = filter_by_name(scb, name_col)
        # Remove those already found by SNI
        name_filtered = name_filtered[~name_filtered["PeOrgNr"].isin(sni_filtered["PeOrgNr"])]
        name_filtered = name_filtered.copy()
        name_filtered["match_method"] = "name_match"
        print(f"  Found {len(name_filtered):,} additional companies by name")
    else:
        print("  WARNING: Could not find company name column")
        name_filtered = pd.DataFrame()

    # Step 4: Combine results
    all_shops = pd.concat([sni_filtered, name_filtered], ignore_index=True)
    print(f"\nTotal computer-related companies found: {len(all_shops):,}")

    # Step 5: Try to enrich with Bolagsverket data (business descriptions)
    print("\n--- Downloading Bolagsverket data for enrichment ---")
    desc_col = None
    status_col = None
    try:
        bv = download_bv_data()
        bv["org_nr_numeric"] = pd.to_numeric(
            bv["organisationsidentitet"].str.strip(), errors="coerce"
        )

        # Find description column
        for col in bv.columns:
            if "verksamhet" in col.lower() or "beskrivning" in col.lower():
                desc_col = col
                break

        # Find status column in BV data
        status_col = find_column(bv, ["status", "foretagsstatus", "företagsstatus",
                                       "Status", "Foretagsstatus"], "company status")

        bv_cols = ["org_nr_numeric", "organisationsnamn", "postadress", "trading_name"]
        if desc_col:
            bv_cols.append(desc_col)
        if status_col:
            bv_cols.append(status_col)
        bv_subset = bv[bv_cols].copy()

        # Merge to enrich existing results
        all_shops = all_shops.merge(
            bv_subset,
            left_on="PeOrgNr",
            right_on="org_nr_numeric",
            how="left",
            suffixes=("", "_bv"),
        )

        # Search BV descriptions for additional computer shops
        if desc_col:
            print(f"\n--- Searching Bolagsverket descriptions ({desc_col}) ---")
            desc_pattern = "|".join(DESCRIPTION_KEYWORDS)
            desc_matches = bv[
                bv[desc_col].fillna("").str.contains(desc_pattern, case=False, regex=True)
            ]
            new_from_desc = desc_matches[
                ~desc_matches["org_nr_numeric"].isin(all_shops["PeOrgNr"])
            ]
            if len(new_from_desc) > 0:
                print(f"  Found {len(new_from_desc):,} additional companies from descriptions")
                extra = scb[scb["PeOrgNr"].isin(new_from_desc["org_nr_numeric"])]
                if len(extra) > 0:
                    extra = extra.copy()
                    extra["match_method"] = "description_match"
                    extra = extra.merge(
                        bv_subset,
                        left_on="PeOrgNr",
                        right_on="org_nr_numeric",
                        how="left",
                        suffixes=("", "_bv"),
                    )
                    all_shops = pd.concat([all_shops, extra], ignore_index=True)
    except Exception as e:
        print(f"  WARNING: Could not load Bolagsverket data: {e}")
        print("  Continuing with SCB data only...")

    # Step 6: Classify chains and buying groups
    print("\n--- Classifying chains and buying groups ---")
    if name_col and name_col in all_shops.columns:
        classification = all_shops[name_col].apply(classify_chain)
        chain_df = pd.DataFrame(classification.tolist(), index=all_shops.index)
        for col in chain_df.columns:
            all_shops[f"_chain_{col}"] = chain_df[col]
    else:
        for key, val in _INDEPENDENT_META.items():
            all_shops[f"_chain_{key}"] = val

    # Step 7: Discover SCB columns for municipality, region, employees
    # Print all available columns to help diagnose missing data
    print(f"\n  Available columns after merge: {sorted(all_shops.columns.tolist())}")

    municipality_col = find_column(all_shops, [
        "KommunKod", "Kommun", "kommun", "KommunNamn", "kommunnamn",
        "KnKod", "KnNamn", "KOMMUN", "kommun_kod", "kommun_namn",
    ], "municipality")
    region_col = find_column(all_shops, [
        "LanKod", "Lan", "lan", "Län", "LänKod", "länkod", "LanNamn",
        "LnKod", "LnNamn", "LANKOD", "LAN", "lan_kod", "lan_namn",
    ], "region/län")
    employees_col = find_column(all_shops, [
        "AntAnst", "antanst", "AntalAnställda", "AntalAnstallda", "Anställda",
        "ANTANST", "Antal", "antal_anst", "AnstKlass", "anstallda",
    ], "employees")

    print(f"  Municipality column: {municipality_col}")
    print(f"  Region column: {region_col}")
    print(f"  Employees column: {employees_col}")

    # Step 8: Build clean output with all target fields
    print("\n--- Building output ---")

    sni_cols = get_sni_columns(all_shops)
    output = pd.DataFrame()

    # --- Identity ---
    output["org_number"] = all_shops["PeOrgNr"].apply(
        lambda x: f"{int(x):012d}" if pd.notna(x) else ""
    )

    if name_col and name_col in all_shops.columns:
        output["company_name"] = all_shops[name_col]
    elif "organisationsnamn" in all_shops.columns:
        output["company_name"] = all_shops["organisationsnamn"]
    else:
        output["company_name"] = ""

    output["trading_name"] = all_shops["trading_name"] if "trading_name" in all_shops.columns else ""

    legal_form_col = find_column(all_shops,
        ["JuridiskForm", "juridiskform", "Juridisk form"], "legal_form")
    output["legal_form"] = all_shops[legal_form_col] if legal_form_col else ""

    reg_date_col = find_column(all_shops, ["RegDatKtid", "regdatktid"], "registration_date")
    if reg_date_col:
        output["founded_year"] = all_shops[reg_date_col].apply(
            lambda x: str(int(x))[:4] if pd.notna(x) and x > 0 else ""
        )
    else:
        output["founded_year"] = ""

    if status_col and status_col in all_shops.columns:
        output["status_active"] = all_shops[status_col]
    else:
        output["status_active"] = ""

    # --- Location ---
    for col_name, candidates in [
        ("address", ["Gatuadress", "gatuadress", "Adress", "adress"]),
        ("co_address", ["COAdress", "coadress", "COadress"]),
        ("postal_code", ["PostNr", "postnr", "Postnummer", "postnummer"]),
        ("city", ["PostOrt", "postort", "Postort"]),
    ]:
        src = find_column(all_shops, candidates)
        output[col_name] = all_shops[src] if src else ""

    output["municipality"] = all_shops[municipality_col] if municipality_col else ""
    output["region_lan"] = all_shops[region_col] if region_col else ""
    # Placeholders for geocoding and location count
    output["lat"] = ""
    output["lng"] = ""
    output["nr_of_locations"] = ""

    # --- Classification ---
    if sni_cols:
        output["primary_sni"] = all_shops[sni_cols[0]]
        output["all_sni_codes"] = all_shops[sni_cols].apply(
            lambda row: ",".join(
                str(int(v)) for v in row if pd.notna(v) and v != 0
            ),
            axis=1,
        )
        # Secondary SNI = all codes minus primary
        output["secondary_sni"] = output.apply(
            lambda row: ",".join(
                c for c in str(row["all_sni_codes"]).split(",")
                if c and c != str(int(row["primary_sni"])) if pd.notna(row["primary_sni"])
            ) if pd.notna(row["primary_sni"]) else "",
            axis=1,
        )
    else:
        output["primary_sni"] = ""
        output["secondary_sni"] = ""
        output["all_sni_codes"] = ""

    output["business_description"] = all_shops[desc_col] if desc_col and desc_col in all_shops.columns else ""
    # Placeholders for manual classification
    output["b2c_b2b"] = ""
    output["online_physical"] = ""
    output["specialisation"] = ""

    # --- Financials ---
    output["employees"] = all_shops[employees_col] if employees_col else ""

    # Estimate turnover from employee count x industry average (when actual data unavailable)
    if employees_col:
        def estimate_turnover(row):
            emp = pd.to_numeric(row.get(employees_col, 0), errors="coerce")
            if pd.isna(emp) or emp <= 0:
                return ""
            sni = pd.to_numeric(row.get(sni_cols[0] if sni_cols else None, 0), errors="coerce")
            rate = TURNOVER_PER_EMPLOYEE_SEK.get(int(sni), DEFAULT_TURNOVER_PER_EMPLOYEE) if pd.notna(sni) else DEFAULT_TURNOVER_PER_EMPLOYEE
            return int(emp * rate)
        output["turnover_sek"] = all_shops.apply(estimate_turnover, axis=1)
        output["turnover_year"] = output["turnover_sek"].apply(
            lambda x: "estimate" if x != "" else ""
        )
    else:
        output["turnover_sek"] = ""
        output["turnover_year"] = ""

    # Placeholders for Allabolag / external enrichment
    output["profit_loss"] = ""
    output["credit_rating"] = ""
    output["equity"] = ""

    # --- Chain & Group ---
    output["is_chain_member"] = all_shops["_chain_chain_type"].apply(
        lambda x: "Yes" if x not in ("Independent", "Unknown", "") else "No"
    )
    output["chain_name"] = all_shops["_chain_chain_name"]
    output["chain_group"] = all_shops["_chain_group_name"]
    output["chain_type"] = all_shops["_chain_chain_type"]
    output["buying_group"] = all_shops["_chain_buying_group"]
    output["parent_company"] = all_shops["_chain_parent_company"]
    output["ultimate_owner"] = all_shops["_chain_ultimate_owner"]
    output["listed_private"] = all_shops["_chain_listed_private"]

    # --- Meta ---
    output["match_method"] = all_shops["match_method"]

    # Apply column order
    output = output[OUTPUT_COLUMNS]

    # Remove duplicates
    output = output.drop_duplicates(subset=["org_number"])

    # Sort by chain group then company name
    output = output.sort_values(["chain_type", "chain_group", "company_name"])
    output = output.reset_index(drop=True)

    # Save
    output_file = "swedish_computer_shops.csv"
    output.to_csv(output_file, index=False, encoding="utf-8-sig")
    print(f"\nSaved {len(output):,} computer shops to {output_file}")

    # Print summary
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Total shops found: {len(output):,}")
    print(f"  With municipality data: {(output['municipality'] != '').sum():,}")
    print(f"  With employee data: {(output['employees'] != '').sum():,}")
    print(f"  With trading name: {output['trading_name'].notna().sum():,}")
    print(f"  With founded year: {(output['founded_year'] != '').sum():,}")
    print(f"\nBy match method:")
    print(output["match_method"].value_counts().to_string())
    print(f"\nChain members: {(output['is_chain_member'] == 'Yes').sum()}")
    print(f"Independent: {(output['is_chain_member'] == 'No').sum()}")
    print(f"\nBy chain/group:")
    for group, count in output["chain_group"].value_counts().items():
        if group != "Independent":
            print(f"  {group}: {count}")
    print(f"\nChain companies:")
    chains = output[output["is_chain_member"] == "Yes"]
    print(chains[["company_name", "city", "primary_sni", "chain_group", "chain_type",
                   "parent_company", "ultimate_owner"]].to_string())
    print(f"\nSample independent companies (first 30):")
    indep = output[output["is_chain_member"] == "No"][["company_name", "city", "primary_sni"]]
    print(indep.head(30).to_string())
    print(f"\nOutput columns ({len(OUTPUT_COLUMNS)}): {OUTPUT_COLUMNS}")

    # --- Generate SNI analysis overview sheet ---
    generate_sni_analysis(output)


def generate_sni_analysis(output):
    """Generate an overview CSV explaining the SNI codes, database logic, and data sources."""
    print("\n--- Generating SNI analysis overview ---")

    # SNI code reference with Swedish + English descriptions
    sni_reference = pd.DataFrame([
        {"sni_code": 47401, "priority": "Primary",
         "description_swedish": "Specialiserad butikshandel med datorer och kringutrustning",
         "description_english": "Retail sale of computers and peripheral equipment",
         "notes": "Core target — dedicated computer shops"},
        {"sni_code": 47402, "priority": "Primary",
         "description_swedish": "Specialiserad butikshandel med programvara",
         "description_english": "Retail sale of software in specialised stores",
         "notes": "Software retail — often same shops as 47401"},
        {"sni_code": 47403, "priority": "Secondary",
         "description_swedish": "Specialiserad butikshandel med hemelektronik",
         "description_english": "Retail sale of consumer electronics",
         "notes": "Consumer electronics — many also sell computers"},
        {"sni_code": 47404, "priority": "Secondary",
         "description_swedish": "Specialiserad butikshandel med telekommunikationsutrustning",
         "description_english": "Retail sale of telecommunications equipment",
         "notes": "Phone/telecom shops — often overlaps with computer shops"},
        {"sni_code": 46501, "priority": "Secondary",
         "description_swedish": "Partihandel med datorer, kringutrustning och programvara",
         "description_english": "Wholesale of computers, peripheral equipment and software",
         "notes": "B2B resellers — Dustin, Atea etc."},
        {"sni_code": 46502, "priority": "Secondary",
         "description_swedish": "Partihandel med elektroniska komponenter",
         "description_english": "Wholesale of electronic components",
         "notes": "Component wholesalers — may also sell retail"},
        {"sni_code": 95101, "priority": "Tertiary",
         "description_swedish": "Reparation av datorer och kringutrustning",
         "description_english": "Repair of computers and peripheral equipment",
         "notes": "Repair shops — many also sell hardware"},
        {"sni_code": 95102, "priority": "Tertiary",
         "description_swedish": "Reparation av kommunikationsutrustning",
         "description_english": "Repair of communication equipment",
         "notes": "Telecom repair — may sell devices too"},
        {"sni_code": 47112, "priority": "Broad (name-match only)",
         "description_swedish": "Detaljhandel med brett sortiment, övervägande livsmedel/drycker",
         "description_english": "Retail with broad range, predominantly food",
         "notes": "Only included if chain name matches (e.g. hypermarkets)"},
        {"sni_code": 47122, "priority": "Broad (name-match only)",
         "description_swedish": "Detaljhandel med brett sortiment, ej livsmedel",
         "description_english": "Retail with broad range, non-food",
         "notes": "Power/Elgiganten stores sometimes registered here"},
        {"sni_code": 47123, "priority": "Broad (name-match only)",
         "description_swedish": "Internethandel med brett sortiment",
         "description_english": "Internet retail with broad range",
         "notes": "Online retailers like NetOnNet sometimes registered here"},
    ])

    # Count how many companies we found per SNI code
    sni_counts = output["primary_sni"].value_counts().reset_index()
    sni_counts.columns = ["sni_code", "companies_found"]
    sni_counts["sni_code"] = pd.to_numeric(sni_counts["sni_code"], errors="coerce")
    sni_reference = sni_reference.merge(sni_counts, on="sni_code", how="left")
    sni_reference["companies_found"] = sni_reference["companies_found"].fillna(0).astype(int)

    # Add overview rows explaining the database logic
    overview = pd.DataFrame([
        {"section": "OVERVIEW",
         "description_english": "What is this database?",
         "notes": "A register of Swedish companies that sell, wholesale, or repair computers and electronics. "
                  "Built from two official Swedish government data sources, enriched with chain/group metadata."},
        {"section": "OVERVIEW",
         "description_english": "How is it generated?",
         "notes": "Step 1: Download the SCB (Statistics Sweden) bulk file — Sweden's official business register. "
                  "Step 2: Filter companies by SNI industry codes (see table below). "
                  "Step 3: Search for known chain names (Elgiganten, Dustin, etc.) to catch companies with non-standard SNI codes. "
                  "Step 4: Download Bolagsverket (Companies Registration Office) data to get business descriptions. "
                  "Step 5: Search descriptions for computer-related keywords to find more shops. "
                  "Step 6: Tag each company with chain/group ownership metadata. "
                  "Step 7: Estimate turnover from employee count x industry average."},
        {"section": "OVERVIEW",
         "description_english": "What is an SNI code?",
         "notes": "SNI 2007 (Standard för Svensk Näringsgrensindelning) is Sweden's official industry classification. "
                  "It is based on the EU standard NACE Rev. 2. Every Swedish company is assigned one or more 5-digit codes "
                  "that describe what the company does. For example, 47401 = 'retail sale of computers'. "
                  "We use these codes to find computer-related businesses."},
        {"section": "OVERVIEW",
         "description_english": "What are the data sources?",
         "notes": "1) SCB CFAR Register — free government extract with org number, name, address, SNI code, municipality, "
                  "employee count. 2) Bolagsverket Open Data — free government data with legal form, registration date, "
                  "business description. 3) Chain metadata — manually researched ownership for known chains. "
                  "4) Turnover estimates — calculated from employee count x industry average per SNI code."},
        {"section": "OVERVIEW",
         "description_english": "Why are some columns empty?",
         "notes": "Columns like turnover_sek show estimates (marked 'estimate' in turnover_year). "
                  "Actual turnover, profit, credit rating require paid sources (Allabolag, UC, Creditsafe). "
                  "Lat/lng require a geocoding service. B2C/B2B, online/physical require manual research. "
                  "These are ready to be filled in as data becomes available."},
        {"section": "OVERVIEW",
         "description_english": "How many companies are in the database?",
         "notes": f"Total: {len(output):,} companies. "
                  f"Chain members: {(output['is_chain_member'] == 'Yes').sum()}. "
                  f"Independent: {(output['is_chain_member'] == 'No').sum()}."},
        {"section": "OVERVIEW",
         "description_english": "How is turnover estimated?",
         "notes": "When actual turnover is not available, we multiply employee count by an industry average: "
                  "Retail (SNI 47.x) = ~2.5 MSEK/employee, Wholesale (SNI 46.x) = ~4.5 MSEK/employee, "
                  "Repair (SNI 95.x) = ~1 MSEK/employee. These are rough Swedish averages from SCB statistics. "
                  "Marked as 'estimate' in the turnover_year column. Replace with actual figures from Allabolag when available."},
    ])

    # Combine overview + SNI reference
    overview_file = "database_overview.csv"
    with open(overview_file, "w", encoding="utf-8-sig") as f:
        f.write("section,sni_code,priority,description_swedish,description_english,companies_found,notes\n")
        # Write overview section
        for _, row in overview.iterrows():
            notes = str(row["notes"]).replace('"', '""')
            desc = str(row.get("description_english", "")).replace('"', '""')
            f.write(f'{row["section"]},,,,"{desc}",,"{notes}"\n')
        # Blank separator
        f.write("\n")
        f.write("SNI CODES,,,,,, \n")
        # Write SNI reference
        for _, row in sni_reference.iterrows():
            desc_sv = str(row["description_swedish"]).replace('"', '""')
            desc_en = str(row["description_english"]).replace('"', '""')
            notes = str(row["notes"]).replace('"', '""')
            f.write(f'SNI,{int(row["sni_code"])},{row["priority"]},"{desc_sv}","{desc_en}",{int(row["companies_found"])},"{notes}"\n')

    print(f"Saved database overview to {overview_file}")


if __name__ == "__main__":
    main()
