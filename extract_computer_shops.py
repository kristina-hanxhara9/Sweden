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

# Primary SNI codes for computer/electronics shops
PRIMARY_SNI_CODES = {
    47410,  # Retail sale of computers, peripheral units and software
}

SECONDARY_SNI_CODES = {
    47420,  # Retail sale of telecommunications equipment
    46510,  # Wholesale of computers, peripheral equipment and software
}

TERTIARY_SNI_CODES = {
    47430,  # Retail sale of audio and video equipment
    95110,  # Repair of computers and peripheral equipment
    47190,  # Other retail sale in non-specialised stores (name-match only)
}

ALL_SNI_CODES = PRIMARY_SNI_CODES | SECONDARY_SNI_CODES | TERTIARY_SNI_CODES

# Chain/buying group classification
CHAIN_PATTERNS = {
    r"elgiganten": ("Elkjop Nordic / Currys plc", "Chain"),
    r"elkj[øo]p": ("Elkjop Nordic / Currys plc", "Chain"),
    r"komplett(?!\s*group)": ("Komplett Group", "Chain"),
    r"netonnet|net\s*on\s*net": ("Komplett Group", "Chain"),
    r"webhallen": ("Komplett Group", "Chain"),
    r"dustin": ("Dustin Group", "Chain"),
    r"kjell\s*[&]\s*co|kjell\s*och\s*co": ("Kjell Group", "Chain"),
    r"kjell\s*group": ("Kjell Group", "Chain"),
    r"\binet\b": ("Inet", "Independent Chain"),
    r"power\s*(?:sverige|international|retail)": ("Power International", "Chain"),
    r"mediamarkt|media\s*markt": ("MediaMarkt (exited Sweden)", "Chain (legacy)"),
    r"\bsiba\b": ("NetOnNet / Komplett Group (legacy SIBA)", "Chain (legacy)"),
}

# Keywords to find computer shops by name/description (Swedish + English)
NAME_KEYWORDS = [
    r"datorbutik",
    r"datorhandel",
    r"datorservice",
    r"datorf[öo]rs[äa]ljning",
    r"dator\s*&",
    r"data\s*&\s*(?:it|tele)",
    r"computer\s*(?:shop|store|center|centre)",
    r"it[-\s]butik",
    r"it[-\s]handel",
    r"elektronikhandel",
    r"gaming\s*(?:shop|store|butik)",
]

DESCRIPTION_KEYWORDS = [
    r"f[öo]rs[äa]ljning\s+av\s+dator",
    r"f[öo]rs[äa]ljning\s+av\s+it",
    r"f[öo]rs[äa]ljning\s+av\s+computer",
    r"detaljhandel\s+med\s+dator",
    r"retail.*computer",
    r"sale.*computer",
    r"dator.*tillbeh[öo]r",
    r"computer.*peripheral",
    r"datorutrustning",
    r"it[-\s]*utrustning",
    r"partihandel\s+med\s+dator",
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
        # Try alternative naming
        sni_cols = [c for c in df.columns if "sni" in c.lower() or "ng" in c.lower()]
    return sni_cols


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
        mask = mask | df[col].isin(sni_codes)
    return df[mask].copy()


def filter_by_name(df, name_col):
    """Filter dataframe by company name keywords."""
    combined_pattern = "|".join(NAME_KEYWORDS)
    mask = df[name_col].fillna("").str.contains(combined_pattern, case=False, regex=True)
    # Also match known chain names
    chain_pattern = "|".join(CHAIN_PATTERNS.keys())
    mask = mask | df[name_col].fillna("").str.contains(chain_pattern, case=False, regex=True)
    return df[mask].copy()


def classify_chain(name):
    """Classify a company by chain/buying group based on name."""
    if pd.isna(name):
        return "Independent", "Independent"
    name_lower = str(name).lower()
    for pattern, (group, chain_type) in CHAIN_PATTERNS.items():
        if re.search(pattern, name_lower):
            return group, chain_type
    return "Independent", "Independent"


def main():
    # Step 1: Download data
    print("=" * 60)
    print("Swedish Computer Shop Extractor")
    print("=" * 60)

    scb = download_scb_data()

    print(f"\nSCB columns: {list(scb.columns)}")

    # Step 2: Filter by SNI codes
    print("\n--- Filtering by SNI codes ---")
    sni_filtered = filter_by_sni(scb, ALL_SNI_CODES)
    print(f"  Found {len(sni_filtered):,} companies matching SNI codes")
    sni_filtered = sni_filtered.copy()
    sni_filtered["match_method"] = "sni_code"

    # Step 3: Filter by name (to catch shops with different SNI codes)
    print("\n--- Filtering by company name ---")
    name_col = None
    for candidate in ["Namn", "namn", "ForetagsNamn", "foretagsnamn", "Företagsnamn"]:
        if candidate in scb.columns:
            name_col = candidate
            break
    if not name_col:
        # Try to find a name-like column
        for col in scb.columns:
            if "namn" in col.lower() or "name" in col.lower():
                name_col = col
                break

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
    try:
        bv = download_bv_data()
        # Create org number for joining (SCB uses numeric, BV uses string)
        bv["org_nr_numeric"] = pd.to_numeric(
            bv["organisationsidentitet"].str.strip(), errors="coerce"
        )

        # Join to get business descriptions
        bv_subset = bv[["org_nr_numeric", "organisationsnamn", "postadress"]].copy()
        # Add verksamhetsbeskrivning if it exists
        desc_col = None
        for col in bv.columns:
            if "verksamhet" in col.lower() or "beskrivning" in col.lower():
                desc_col = col
                bv_subset[col] = bv[col]
                break

        all_shops = all_shops.merge(
            bv_subset,
            left_on="PeOrgNr",
            right_on="org_nr_numeric",
            how="left",
            suffixes=("", "_bv"),
        )

        # Also search BV descriptions for computer shops not yet found
        if desc_col:
            print(f"\n--- Searching Bolagsverket descriptions ({desc_col}) ---")
            desc_pattern = "|".join(DESCRIPTION_KEYWORDS)
            desc_matches = bv[
                bv[desc_col].fillna("").str.contains(desc_pattern, case=False, regex=True)
            ]
            # Filter to those not already in our results
            new_from_desc = desc_matches[
                ~desc_matches["org_nr_numeric"].isin(all_shops["PeOrgNr"])
            ]
            if len(new_from_desc) > 0:
                print(f"  Found {len(new_from_desc):,} additional companies from descriptions")
                # Need to get SCB data for these
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

    # Company name - prefer SCB name
    if name_col and name_col in all_shops.columns:
        output["company_name"] = all_shops[name_col]
    elif "organisationsnamn" in all_shops.columns:
        output["company_name"] = all_shops["organisationsnamn"]
    else:
        output["company_name"] = ""

    # Address fields
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

    # SNI codes
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

    # Business description from Bolagsverket
    desc_col_out = None
    for col in all_shops.columns:
        if "verksamhet" in col.lower() or "beskrivning" in col.lower():
            desc_col_out = col
            break
    output["business_description"] = all_shops[desc_col_out] if desc_col_out else ""

    # Legal form
    for c in ["JuridiskForm", "juridiskform", "Juridisk form"]:
        if c in all_shops.columns:
            output["legal_form"] = all_shops[c]
            break
    else:
        output["legal_form"] = ""

    # Registration date
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
    print(output["chain_group"].value_counts().to_string())
    print(f"\nSample entries:")
    print(output[["company_name", "city", "primary_sni", "chain_group", "chain_type"]].head(20).to_string())


if __name__ == "__main__":
    main()
