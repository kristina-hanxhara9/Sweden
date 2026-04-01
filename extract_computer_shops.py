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
    """Find the first matching column name from a list of candidates."""
    for c in candidates:
        if c in df.columns:
            return c
    if label:
        print(f"  WARNING: Could not find column for {label} (tried: {candidates})")
    return None


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
    municipality_col = find_column(all_shops,
        ["KommunKod", "Kommun", "kommun", "KommunNamn", "kommunnamn"], "municipality")
    region_col = find_column(all_shops,
        ["LanKod", "Lan", "lan", "Län", "LänKod", "länkod", "LanNamn"], "region/län")
    employees_col = find_column(all_shops,
        ["AntAnst", "antanst", "AntalAnställda", "AntalAnstallda", "Anställda"], "employees")

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
    # Placeholders for Allabolag / external enrichment
    output["turnover_sek"] = ""
    output["turnover_year"] = ""
    output["employees"] = all_shops[employees_col] if employees_col else ""
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


if __name__ == "__main__":
    main()
