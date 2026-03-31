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

# Chain/buying group classification - EXACT patterns to avoid false positives
CHAIN_PATTERNS = [
    # (pattern, group_name, chain_type, must_match_sni) - order matters, first match wins
    (r"^elgiganten\b", "Elkjop Nordic / Currys plc", "Chain", False),
    (r"elkj[øo]p\s*nordic", "Elkjop Nordic / Currys plc", "Chain", False),
    (r"^komplett\s*(?:services?\s*(?:sweden|norge)|distribution|business\s*nordic)", "Komplett Group", "Chain", False),
    (r"^komplett\.se\b|^komplett\s*(?:sweden|group)\b", "Komplett Group", "Chain", False),
    (r"^netonnet\b|^net\s*on\s*net\b", "Komplett Group", "Chain", False),
    (r"^webhallen\b", "Komplett Group", "Chain", False),
    (r"^dustin\s*(?:aktiebolag|ab|group|sverige|a/s|finland|norway)\b", "Dustin Group", "Chain", False),
    (r"^kjell\s*[&]\s*co\b|^kjell\s*group\b", "Kjell Group", "Chain", False),
    (r"^inet\s*(?:ab|group)\b", "Inet", "Independent Chain", False),
    (r"^power\s*(?:sverige|retail\s*sweden)\s*ab\b", "Power International", "Chain", False),
    (r"^power\s*international\b", "Power International", "Chain", False),
    (r"^mediamarkt\b|^media\s*markt\b", "MediaMarkt (exited Sweden)", "Chain (legacy)", False),
    (r"^siba\s*(?:aktiebolag|ab|fastigheter|invest)\b", "NetOnNet / Komplett Group (legacy SIBA)", "Chain (legacy)", False),
]

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
    # Clean name
    df["organisationsnamn"] = df["organisationsnamn"].str.split("$FORETAGSNAMN").str[0]
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
    for pattern, group, chain_type, _ in CHAIN_PATTERNS:
        mask = mask | df[name_col].fillna("").str.contains(pattern, case=False, regex=True)

    return df[mask].copy()


def classify_chain(name):
    """Classify a company by chain/buying group based on name."""
    if pd.isna(name):
        return "Independent", "Independent"
    name_clean = str(name).strip()
    for pattern, group, chain_type, _ in CHAIN_PATTERNS:
        if re.search(pattern, name_clean, re.IGNORECASE):
            return group, chain_type
    return "Independent", "Independent"


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

        bv_cols = ["org_nr_numeric", "organisationsnamn", "postadress"]
        if desc_col:
            bv_cols.append(desc_col)
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
        all_shops["chain_group"] = classification.apply(lambda x: x[0])
        all_shops["chain_type"] = classification.apply(lambda x: x[1])
    else:
        all_shops["chain_group"] = "Unknown"
        all_shops["chain_type"] = "Unknown"

    # Step 7: Build clean output
    print("\n--- Building output ---")

    sni_cols = get_sni_columns(all_shops)
    output = pd.DataFrame()
    output["org_number"] = all_shops["PeOrgNr"].apply(
        lambda x: f"{int(x):012d}" if pd.notna(x) else ""
    )

    if name_col and name_col in all_shops.columns:
        output["company_name"] = all_shops[name_col]
    elif "organisationsnamn" in all_shops.columns:
        output["company_name"] = all_shops["organisationsnamn"]
    else:
        output["company_name"] = ""

    for col_name, candidates in [
        ("address", ["Gatuadress", "gatuadress", "Adress", "adress"]),
        ("co_address", ["COAdress", "coadress", "COadress"]),
        ("postal_code", ["PostNr", "postnr", "Postnummer", "postnummer"]),
        ("city", ["PostOrt", "postort", "Postort"]),
    ]:
        found = False
        for c in candidates:
            if c in all_shops.columns:
                output[col_name] = all_shops[c]
                found = True
                break
        if not found:
            output[col_name] = ""

    if sni_cols:
        output["primary_sni"] = all_shops[sni_cols[0]]
        output["all_sni_codes"] = all_shops[sni_cols].apply(
            lambda row: ",".join(
                str(int(v)) for v in row if pd.notna(v) and v != 0
            ),
            axis=1,
        )
    else:
        output["primary_sni"] = ""
        output["all_sni_codes"] = ""

    output["business_description"] = all_shops[desc_col] if desc_col and desc_col in all_shops.columns else ""

    for c in ["JuridiskForm", "juridiskform", "Juridisk form"]:
        if c in all_shops.columns:
            output["legal_form"] = all_shops[c]
            break
    else:
        output["legal_form"] = ""

    for c in ["RegDatKtid", "regdatktid"]:
        if c in all_shops.columns:
            output["registration_date"] = all_shops[c]
            break
    else:
        output["registration_date"] = ""

    output["match_method"] = all_shops["match_method"]
    output["chain_group"] = all_shops["chain_group"]
    output["chain_type"] = all_shops["chain_type"]

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
    print(f"\nBy match method:")
    print(output["match_method"].value_counts().to_string())
    print(f"\nBy chain type:")
    print(output["chain_type"].value_counts().to_string())
    print(f"\nBy chain/group:")
    for group, count in output["chain_group"].value_counts().items():
        print(f"  {group}: {count}")
    print(f"\nChain companies:")
    chains = output[output["chain_type"] != "Independent"]
    print(chains[["company_name", "city", "primary_sni", "chain_group", "chain_type"]].to_string())
    print(f"\nSample independent companies (first 30):")
    indep = output[output["chain_type"] == "Independent"][["company_name", "city", "primary_sni"]]
    print(indep.head(30).to_string())


if __name__ == "__main__":
    main()
