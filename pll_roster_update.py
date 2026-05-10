# ============================================================
# PLL ROSTER UPDATE SCRIPT
# ------------------------------------------------------------
# Purpose:
#   - Scrape official current PLL rosters
#   - Update existing Google Sheet in-place
#   - Preserve formatting and manual user selections
#   - Update visible successful roster update timestamps
#
# Updates only:
#   - Master Player Database values
#   - Lists values
#   - Team roster tables A10:J44
#   - Dropdown validation source ranges
#   - Dashboard update status cells
#   - Dashboard!B4 last successful roster update timestamp
#   - Master Player Database!B4 last successful roster update timestamp
#
# Does NOT update:
#   - Depth chart selections
#   - Projected lineup selections
#   - Injury tracker entries
#   - Tier board selections
#   - Sheet formatting/layout/merged cells
# ============================================================

import os
import re
import json
import time
import asyncio
import pandas as pd
import gspread

from datetime import datetime
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials
from gspread.utils import a1_range_to_grid_range
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError


# ============================================================
# CONFIG
# ============================================================

SPREADSHEET_ID = os.environ.get("PLL_SPREADSHEET_ID", "").strip()
SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
TRIGGERED_BY = os.environ.get("TRIGGERED_BY", "unknown")
TRIGGER_SOURCE = os.environ.get("TRIGGER_SOURCE", "unknown")

if not SPREADSHEET_ID:
    raise RuntimeError("Missing PLL_SPREADSHEET_ID environment variable.")

if not SERVICE_ACCOUNT_JSON:
    raise RuntimeError("Missing GOOGLE_SERVICE_ACCOUNT_JSON environment variable.")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADLESS = True
PAGE_TIMEOUT_MS = 60000
CARD_WAIT_TIMEOUT_MS = 30000
SCROLL_PASSES = 8
SCROLL_PAUSE_MS = 600

FIXED_ROSTER_ROWS = 35
FIXED_INJURY_ROWS = 20
LISTS_MAX_ROWS = 420

ROSTER_START_ROW = 10
ROSTER_END_ROW = 44

INJURY_START_ROW = 10
INJURY_END_ROW = 29

MASTER_HEADER_ROW = 7
MASTER_DATA_START_ROW = 8
MASTER_CLEAR_END_ROW = 1000

PLL_TEAMS = [
    {
        "Team_Code": "BOS",
        "Team": "Boston Cannons",
        "Tab": "BOS Cannons",
        "Division": "Eastern",
        "URLs": [
            "https://premierlacrosseleague.com/teams/boston-cannons/roster",
            "https://premierlacrosseleague.com/teams/Cannons/roster",
        ],
    },
    {
        "Team_Code": "CAL",
        "Team": "California Redwoods",
        "Tab": "CAL Redwoods",
        "Division": "Western",
        "URLs": [
            "https://premierlacrosseleague.com/teams/california-redwoods/roster",
            "https://premierlacrosseleague.com/teams/Redwoods/roster",
            "https://premierlacrosseleague.com/teams/redwoods/roster",
        ],
    },
    {
        "Team_Code": "CAR",
        "Team": "Carolina Chaos",
        "Tab": "CAR Chaos",
        "Division": "Western",
        "URLs": [
            "https://premierlacrosseleague.com/teams/carolina-chaos/roster",
            "https://premierlacrosseleague.com/teams/Chaos/roster",
        ],
    },
    {
        "Team_Code": "DEN",
        "Team": "Denver Outlaws",
        "Tab": "DEN Outlaws",
        "Division": "Western",
        "URLs": [
            "https://premierlacrosseleague.com/teams/denver-outlaws/roster",
            "https://premierlacrosseleague.com/teams/Outlaws/roster",
        ],
    },
    {
        "Team_Code": "MD",
        "Team": "Maryland Whipsnakes",
        "Tab": "MD Whipsnakes",
        "Division": "Eastern",
        "URLs": [
            "https://premierlacrosseleague.com/teams/maryland-whipsnakes/roster",
            "https://premierlacrosseleague.com/teams/Whipsnakes/roster",
        ],
    },
    {
        "Team_Code": "NY",
        "Team": "New York Atlas",
        "Tab": "NY Atlas",
        "Division": "Eastern",
        "URLs": [
            "https://premierlacrosseleague.com/teams/new-york-atlas/roster",
            "https://premierlacrosseleague.com/teams/Atlas/roster",
        ],
    },
    {
        "Team_Code": "PHI",
        "Team": "Philadelphia Waterdogs",
        "Tab": "PHI Waterdogs",
        "Division": "Eastern",
        "URLs": [
            "https://premierlacrosseleague.com/teams/philadelphia-waterdogs/roster",
            "https://premierlacrosseleague.com/teams/Waterdogs/roster",
        ],
    },
    {
        "Team_Code": "UTA",
        "Team": "Utah Archers",
        "Tab": "UTA Archers",
        "Division": "Western",
        "URLs": [
            "https://premierlacrosseleague.com/teams/utah-archers/roster",
            "https://premierlacrosseleague.com/teams/Archers/roster",
        ],
    },
]

POSITIONS = ["A", "M", "SSDM", "LSM", "D", "FO", "G"]

POSITION_ORDER = {
    "A": 1,
    "M": 2,
    "SSDM": 3,
    "LSM": 4,
    "D": 5,
    "FO": 6,
    "G": 7,
    "UNK": 99,
}

POSITION_GROUP = {
    "A": "Attack",
    "M": "Midfield",
    "SSDM": "Short Stick Defensive Midfield",
    "LSM": "Long Stick Midfield",
    "D": "Defense",
    "FO": "Faceoff",
    "G": "Goalie",
    "UNK": "Unknown",
}

LINEUP_STATUSES = [
    "Active",
    "Starter",
    "Rotation",
    "Depth",
    "Inactive",
    "Out",
    "Practice",
    "Unknown",
]

INJURY_STATUSES = [
    "Healthy",
    "Questionable",
    "Doubtful",
    "Out",
    "Injured Reserve",
    "Suspended",
    "Unknown",
]

MASTER_TIERS = [
    "Franchise",
    "Elite",
    "Starter",
    "Rotation",
    "Depth",
    "Prospect",
    "Free Agent",
    "Watchlist",
    "Scrub",
]

SCRAPE_COLUMNS = [
    "Player",
    "First_Name",
    "Last_Name",
    "Team",
    "Team_Code",
    "Division",
    "Position",
    "Position_Group",
    "Jersey",
    "Handedness",
    "Height",
    "Age",
    "College",
    "Country",
    "Image_Slug",
    "Image_URL",
    "Page_URL",
    "Page_Title",
    "Extracted_At",
]

SHEET_MASTER_COLUMNS = [
    "Player",
    "First Name",
    "Last Name",
    "Team",
    "Team Code",
    "Division",
    "Position",
    "Position Group",
    "Jersey",
    "Handedness",
    "Height",
    "Age",
    "College",
    "Country",
    "Image Slug",
    "Image URL",
    "Source URL",
    "Page Title",
    "Last Updated",
    "Tier",
    "Lineup Status",
    "Injury Status",
    "Manual Notes",
]

TEAM_ROSTER_COLUMNS = [
    "Player",
    "Position",
    "Jersey",
    "Handedness",
    "Height",
    "Age",
    "College",
    "Country",
    "Lineup Status",
    "Notes",
]


# ============================================================
# BASIC HELPERS
# ============================================================

def now_label():
    """
    Returns a readable Eastern Time timestamp for roster update tracking.
    GitHub Actions runners use UTC by default, so this forces ET.
    """
    return datetime.now(ZoneInfo("America/New_York")).strftime("%Y-%m-%d %I:%M:%S %p ET")


def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def normalize_key(value):
    value = clean_text(value).lower()
    value = re.sub(r"[^a-z0-9]+", "", value)
    return value


def player_key_from_row(row):
    image_slug = clean_text(row.get("Image Slug", row.get("Image_Slug", "")))
    player = clean_text(row.get("Player", ""))

    if image_slug:
        return f"slug::{normalize_key(image_slug)}"

    return f"name::{normalize_key(player)}"


def normalize_position(pos):
    pos = clean_text(pos).upper()

    aliases = {
        "ATTACK": "A",
        "MIDFIELD": "M",
        "DEFENSE": "D",
        "FACEOFF": "FO",
        "FACE-OFF": "FO",
        "GOALIE": "G",
        "GOALTENDER": "G",
        "LONG STICK MIDFIELD": "LSM",
        "SHORT STICK DEFENSIVE MIDFIELD": "SSDM",
    }

    return aliases.get(pos, pos if pos else "UNK")


def clean_age(value):
    value = clean_text(value)
    match = re.search(r"\b(\d{1,2})\b", value)
    return match.group(1) if match else ""


def clean_height(value):
    return clean_text(value).replace("`", "").replace('"', "")


def col_to_letter(col_num):
    result = ""

    while col_num:
        col_num, rem = divmod(col_num - 1, 26)
        result = chr(65 + rem) + result

    return result


def grid_range(ws, a1_range):
    return a1_range_to_grid_range(a1_range, sheet_id=ws.id)


def safe_batch_update(sh, requests, chunk_size=80):
    requests = [r for r in requests if r]

    if not requests:
        return

    for i in range(0, len(requests), chunk_size):
        chunk = requests[i:i + chunk_size]
        sh.batch_update({"requests": chunk})
        time.sleep(0.25)


def write_values(ws, start_cell, values, value_input_option="USER_ENTERED"):
    if not values:
        return

    ws.update(
        range_name=start_cell,
        values=values,
        value_input_option=value_input_option,
    )


def validation_one_of_list(ws, a1_range, options, strict=False):
    return {
        "setDataValidation": {
            "range": grid_range(ws, a1_range),
            "rule": {
                "condition": {
                    "type": "ONE_OF_LIST",
                    "values": [{"userEnteredValue": str(x)} for x in options],
                },
                "showCustomUi": True,
                "strict": strict,
            },
        }
    }


def validation_one_of_range(ws, a1_range, range_formula, strict=False):
    return {
        "setDataValidation": {
            "range": grid_range(ws, a1_range),
            "rule": {
                "condition": {
                    "type": "ONE_OF_RANGE",
                    "values": [{"userEnteredValue": f"={range_formula}"}],
                },
                "showCustomUi": True,
                "strict": strict,
            },
        }
    }


# ============================================================
# SCRAPER
# ============================================================

ROSTER_EXTRACTOR_JS = """
(teamInfo) => {
  function cleanText(x) {
    return (x || "").replace(/\\s+/g, " ").trim();
  }

  function normalizePosition(pos) {
    pos = cleanText(pos).toUpperCase();

    const aliases = {
      "ATTACK": "A",
      "MIDFIELD": "M",
      "DEFENSE": "D",
      "FACEOFF": "FO",
      "FACE-OFF": "FO",
      "GOALIE": "G",
      "GOALTENDER": "G",
      "LONG STICK MIDFIELD": "LSM",
      "SHORT STICK DEFENSIVE MIDFIELD": "SSDM"
    };

    return aliases[pos] || pos || "UNK";
  }

  function extractDetails(card) {
    const details = {};

    card.querySelectorAll("div").forEach(div => {
      const spans = Array.from(div.children || []).filter(el => el.tagName === "SPAN");

      if (spans.length >= 2) {
        const label = cleanText(spans[0].innerText);
        const value = cleanText(spans[1].innerText);

        if (label && value) {
          details[label] = value;
        }
      }
    });

    return details;
  }

  const cards = Array.from(document.querySelectorAll("div.css-fps5zs"));
  const rows = [];

  cards.forEach(card => {
    const firstName = cleanText(card.querySelector("p.firstName")?.innerText);
    const lastName = cleanText(card.querySelector("p.lastName")?.innerText);
    const player = cleanText(`${firstName} ${lastName}`);

    const jersey = cleanText(card.querySelector(".points")?.innerText);

    const playerImg = card.querySelector(".playerImg img");
    const imageSlug = cleanText(playerImg?.getAttribute("alt"));
    const imageURL = cleanText(playerImg?.getAttribute("src"));

    let country = "";

    card.querySelectorAll("img").forEach(img => {
      const alt = cleanText(img.getAttribute("alt"));

      if (alt.toLowerCase().startsWith("country")) {
        country = cleanText(alt.replace(/^Country:\\s*/i, ""));
      }
    });

    const details = extractDetails(card);

    const row = {
      Player: player,
      First_Name: firstName,
      Last_Name: lastName,
      Team: teamInfo.Team,
      Team_Code: teamInfo.Team_Code,
      Division: teamInfo.Division,
      Position: normalizePosition(details["Position"]),
      Jersey: jersey,
      Handedness: cleanText(details["Hand"]),
      Height: cleanText(details["Height"]),
      Age: cleanText(details["Age"]),
      College: cleanText(details["College"]),
      Country: country,
      Image_Slug: imageSlug,
      Image_URL: imageURL,
      Page_URL: window.location.href,
      Page_Title: document.title,
      Extracted_At: new Date().toISOString()
    };

    if (row.Player && row.Position && row.Position !== "UNK") {
      rows.push(row);
    }
  });

  return {
    raw_card_count: cards.length,
    raw_valid_rows: rows.length,
    rows: rows
  };
}
"""


def dedupe_team_rows(rows):
    best = {}

    for row in rows:
        player = clean_text(row.get("Player"))
        image_slug = clean_text(row.get("Image_Slug"))
        key = f"{player}|{image_slug}"

        if not player:
            continue

        score = sum(1 for v in row.values() if clean_text(v))

        if key not in best:
            best[key] = (score, row)
        else:
            old_score, _ = best[key]
            if score > old_score:
                best[key] = (score, row)

    return [x[1] for x in best.values()]


async def launch_browser(playwright):
    args = [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--disable-setuid-sandbox",
        "--disable-software-rasterizer",
    ]

    return await playwright.chromium.launch(
        headless=HEADLESS,
        args=args,
    )


async def scrape_team_roster(page, team):
    print(f"\nSCRAPING {team['Team_Code']} — {team['Team']}")

    diagnostics = []
    best_rows = []

    for url in team["URLs"]:
        print(f"Trying URL: {url}")

        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT_MS)

            try:
                await page.wait_for_selector("p.firstName", timeout=CARD_WAIT_TIMEOUT_MS)
            except PlaywrightTimeoutError:
                print("  Warning: p.firstName selector did not appear before timeout.")

            for _ in range(SCROLL_PASSES):
                await page.mouse.wheel(0, 2500)
                await page.wait_for_timeout(SCROLL_PAUSE_MS)

            await page.mouse.wheel(0, -10000)
            await page.wait_for_timeout(1000)

            result = await page.evaluate(
                ROSTER_EXTRACTOR_JS,
                {
                    "Team": team["Team"],
                    "Team_Code": team["Team_Code"],
                    "Division": team["Division"],
                },
            )

            raw_card_count = result.get("raw_card_count", 0)
            raw_valid_rows = result.get("raw_valid_rows", 0)
            rows = result.get("rows", [])
            deduped_rows = dedupe_team_rows(rows)

            print(f"  Raw card containers: {raw_card_count}")
            print(f"  Raw valid rows: {raw_valid_rows}")
            print(f"  Deduped players: {len(deduped_rows)}")

            diagnostics.append({
                "Team_Code": team["Team_Code"],
                "Team": team["Team"],
                "URL_Tried": url,
                "Final_URL": page.url,
                "Raw_Card_Containers": raw_card_count,
                "Raw_Valid_Rows": raw_valid_rows,
                "Deduped_Players": len(deduped_rows),
                "Status": "OK" if len(deduped_rows) else "NO_PLAYERS",
                "Error": "",
            })

            if len(deduped_rows) > len(best_rows):
                best_rows = deduped_rows

            if len(deduped_rows) >= 15:
                break

        except Exception as e:
            print(f"  ERROR: {e}")

            diagnostics.append({
                "Team_Code": team["Team_Code"],
                "Team": team["Team"],
                "URL_Tried": url,
                "Final_URL": "",
                "Raw_Card_Containers": 0,
                "Raw_Valid_Rows": 0,
                "Deduped_Players": 0,
                "Status": "ERROR",
                "Error": str(e),
            })

    print(f"FINAL {team['Team_Code']} PLAYERS: {len(best_rows)}")

    return best_rows, diagnostics


def sort_master_roster(df):
    if df.empty:
        return df

    team_rank = {team["Team_Code"]: i for i, team in enumerate(PLL_TEAMS)}

    out = df.copy()
    out["_team_rank"] = out["Team_Code"].map(team_rank)
    out["_pos_rank"] = out["Position"].map(lambda x: POSITION_ORDER.get(x, 99))
    out["_last_name"] = out["Last_Name"].astype(str).str.lower()

    out = (
        out.sort_values(["_team_rank", "_pos_rank", "_last_name", "Player"])
        .drop(columns=["_team_rank", "_pos_rank", "_last_name"], errors="ignore")
        .reset_index(drop=True)
    )

    return out


async def scrape_all_pll_rosters_async():
    all_rows = []
    all_diagnostics = []

    async with async_playwright() as p:
        browser = await launch_browser(p)

        context = await browser.new_context(
            viewport={"width": 1600, "height": 2400},
            user_agent=(
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
            ),
        )

        page = await context.new_page()

        try:
            for team in PLL_TEAMS:
                rows, diagnostics = await scrape_team_roster(page, team)
                all_rows.extend(rows)
                all_diagnostics.extend(diagnostics)
                await page.wait_for_timeout(1000)

        finally:
            await context.close()
            await browser.close()

    roster_df = pd.DataFrame(all_rows)

    if roster_df.empty:
        roster_df = pd.DataFrame(columns=SCRAPE_COLUMNS)
    else:
        for col in SCRAPE_COLUMNS:
            if col not in roster_df.columns:
                roster_df[col] = ""

        roster_df = roster_df[[
            "Player",
            "First_Name",
            "Last_Name",
            "Team",
            "Team_Code",
            "Division",
            "Position",
            "Jersey",
            "Handedness",
            "Height",
            "Age",
            "College",
            "Country",
            "Image_Slug",
            "Image_URL",
            "Page_URL",
            "Page_Title",
            "Extracted_At",
        ]].copy()

        for col in roster_df.columns:
            roster_df[col] = roster_df[col].map(clean_text)

        roster_df["Position"] = roster_df["Position"].map(normalize_position)
        roster_df["Position_Group"] = roster_df["Position"].map(lambda x: POSITION_GROUP.get(x, "Unknown"))
        roster_df["Height"] = roster_df["Height"].map(clean_height)
        roster_df["Age"] = roster_df["Age"].map(clean_age)

        roster_df = roster_df[SCRAPE_COLUMNS].copy()
        roster_df = roster_df.drop_duplicates(subset=["Team_Code", "Player", "Image_Slug"])
        roster_df = sort_master_roster(roster_df)

    diagnostics_df = pd.DataFrame(all_diagnostics)

    return roster_df, diagnostics_df


# ============================================================
# SHEET DATA BUILDERS
# ============================================================

def standardize_for_sheet(pll_rosters_df):
    df = pll_rosters_df.copy()

    df = df.rename(columns={
        "First_Name": "First Name",
        "Last_Name": "Last Name",
        "Team_Code": "Team Code",
        "Position_Group": "Position Group",
        "Image_Slug": "Image Slug",
        "Image_URL": "Image URL",
        "Page_URL": "Source URL",
        "Page_Title": "Page Title",
        "Extracted_At": "Last Updated",
    })

    for col in SHEET_MASTER_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df["Tier"] = ""
    df["Lineup Status"] = "Active"
    df["Injury Status"] = "Healthy"
    df["Manual Notes"] = ""

    df = df[SHEET_MASTER_COLUMNS].copy()

    return df


def read_existing_master_manual_fields(master_ws):
    values = master_ws.get_values(f"A{MASTER_HEADER_ROW}:W{MASTER_CLEAR_END_ROW}")

    if not values or len(values) < 2:
        return {}

    headers = values[0]
    rows = values[1:]

    manual_map = {}

    for row_values in rows:
        row = {}

        for i, header in enumerate(headers):
            row[header] = row_values[i] if i < len(row_values) else ""

        player = clean_text(row.get("Player"))
        if not player:
            continue

        key = player_key_from_row(row)

        manual_map[key] = {
            "Tier": clean_text(row.get("Tier")),
            "Lineup Status": clean_text(row.get("Lineup Status")) or "Active",
            "Injury Status": clean_text(row.get("Injury Status")) or "Healthy",
            "Manual Notes": clean_text(row.get("Manual Notes")),
        }

    return manual_map


def preserve_master_manual_fields(new_df, manual_map):
    out = new_df.copy()

    for idx, row in out.iterrows():
        key = player_key_from_row(row.to_dict())

        if key in manual_map:
            for col, value in manual_map[key].items():
                out.at[idx, col] = value
        else:
            out.at[idx, "Lineup Status"] = out.at[idx, "Lineup Status"] or "Active"
            out.at[idx, "Injury Status"] = out.at[idx, "Injury Status"] or "Healthy"

    return out


def read_existing_team_roster_manual_fields(ws):
    values = ws.get_values(f"A{ROSTER_START_ROW}:J{ROSTER_END_ROW}")

    manual = {}

    for row in values:
        row = row + [""] * (len(TEAM_ROSTER_COLUMNS) - len(row))

        player = clean_text(row[0])

        if not player:
            continue

        manual[normalize_key(player)] = {
            "Lineup Status": clean_text(row[8]) or "Active",
            "Notes": clean_text(row[9]),
        }

    return manual


def build_team_roster_values(team, sheet_df, existing_manual):
    code = team["Team_Code"]

    team_df = sheet_df[sheet_df["Team Code"] == code].copy()

    team_df["_pos_rank"] = team_df["Position"].map(lambda x: POSITION_ORDER.get(x, 99))
    team_df["_last_name"] = team_df["Last Name"].astype(str).str.lower()
    team_df = team_df.sort_values(["_pos_rank", "_last_name", "Player"])

    roster_df = team_df[[
        "Player",
        "Position",
        "Jersey",
        "Handedness",
        "Height",
        "Age",
        "College",
        "Country",
    ]].copy()

    roster_df["Lineup Status"] = "Active"
    roster_df["Notes"] = ""

    for idx, row in roster_df.iterrows():
        key = normalize_key(row["Player"])

        if key in existing_manual:
            roster_df.at[idx, "Lineup Status"] = existing_manual[key].get("Lineup Status") or "Active"
            roster_df.at[idx, "Notes"] = existing_manual[key].get("Notes") or ""

    roster_df = roster_df[TEAM_ROSTER_COLUMNS].copy()

    if len(roster_df) < FIXED_ROSTER_ROWS:
        blank_rows = pd.DataFrame(
            [[""] * len(TEAM_ROSTER_COLUMNS)] * (FIXED_ROSTER_ROWS - len(roster_df)),
            columns=TEAM_ROSTER_COLUMNS,
        )
        roster_df = pd.concat([roster_df, blank_rows], ignore_index=True)

    if len(roster_df) > FIXED_ROSTER_ROWS:
        print(f"Warning: {code} has {len(roster_df)} players, more than fixed {FIXED_ROSTER_ROWS} rows.")

    return roster_df.head(FIXED_ROSTER_ROWS).fillna("").astype(str)


def build_lists_matrix(sheet_df):
    list_columns = []

    def add_list_col(name, values):
        clean_values = []

        for value in values:
            value = clean_text(value)
            if value:
                clean_values.append(value)

        list_columns.append((name, clean_values))

    add_list_col("TEAM_CODES", [t["Team_Code"] for t in PLL_TEAMS])
    add_list_col("POSITIONS", POSITIONS)
    add_list_col("INJURY_STATUSES", INJURY_STATUSES)
    add_list_col("LINEUP_STATUSES", LINEUP_STATUSES)
    add_list_col("MASTER_TIERS", MASTER_TIERS)

    all_players_display = (
        sheet_df
        .assign(Display=lambda d: d["Player"] + " (" + d["Team Code"] + " - " + d["Position"] + ")")
        ["Display"]
        .tolist()
    )

    add_list_col("ALL_PLAYERS_DISPLAY", all_players_display)

    for pos in POSITIONS:
        pos_display = (
            sheet_df[sheet_df["Position"] == pos]
            .assign(Display=lambda d: d["Player"] + " (" + d["Team Code"] + ")")
            ["Display"]
            .tolist()
        )
        add_list_col(f"ALL_{pos}", pos_display)

    for team in PLL_TEAMS:
        code = team["Team_Code"]
        team_df = sheet_df[sheet_df["Team Code"] == code].copy()

        add_list_col(f"{code}_ALL", team_df["Player"].tolist())

        for pos in POSITIONS:
            add_list_col(f"{code}_{pos}", team_df[team_df["Position"] == pos]["Player"].tolist())

    max_len = LISTS_MAX_ROWS - 1

    headers = [name for name, _ in list_columns]
    matrix = [headers]

    for i in range(max_len):
        row = []

        for _, values in list_columns:
            row.append(values[i] if i < len(values) else "")

        matrix.append(row)

    list_ranges = {}

    for idx, (name, values) in enumerate(list_columns, start=1):
        col_letter = col_to_letter(idx)
        list_ranges[name] = f"'Lists'!${col_letter}$2:${col_letter}${LISTS_MAX_ROWS}"

    return matrix, list_ranges, list_columns


# ============================================================
# DROPDOWN VALIDATION UPDATER
# ============================================================

def projected_card_specs():
    return [
        {"slot": "A1", "pos": "A", "player_range": "G52:I52"},
        {"slot": "A2", "pos": "A", "player_range": "J52:L52"},
        {"slot": "A3", "pos": "A", "player_range": "M52:O52"},

        {"slot": "M1", "pos": "M", "player_range": "G57:I57"},
        {"slot": "M2", "pos": "M", "player_range": "J57:L57"},
        {"slot": "M3", "pos": "M", "player_range": "M57:O57"},

        {"slot": "D1", "pos": "D", "player_range": "G62:I62"},
        {"slot": "D2", "pos": "D", "player_range": "J62:L62"},
        {"slot": "D3", "pos": "D", "player_range": "M62:O62"},

        {"slot": "G", "pos": "G", "player_range": "G67:H67"},
        {"slot": "FO1", "pos": "FO", "player_range": "I67:J67"},
        {"slot": "LSM1", "pos": "LSM", "player_range": "K67:L67"},
        {"slot": "SSDM1", "pos": "SSDM", "player_range": "M67:O67"},
    ]


def build_depth_chart_rows():
    plan = [
        ("Attack", "A", 6),
        ("Midfield", "M", 6),
        ("Defense", "D", 5),
        ("SSDM", "SSDM", 3),
        ("LSM", "LSM", 3),
        ("Faceoff", "FO", 2),
        ("Goalie", "G", 2),
    ]

    rows = []

    for group, pos, count in plan:
        for rank in range(1, count + 1):
            rows.append([group, pos, rank])

    return rows


def reapply_all_dropdown_validations(sh, list_ranges, sheet_row_count):
    requests = []

    # Master DB editable dropdowns
    master_ws = sh.worksheet("Master Player Database")
    master_end = max(MASTER_DATA_START_ROW + sheet_row_count - 1, MASTER_DATA_START_ROW)

    requests += [
        validation_one_of_list(master_ws, f"T{MASTER_DATA_START_ROW}:T{master_end}", MASTER_TIERS, strict=False),
        validation_one_of_list(master_ws, f"U{MASTER_DATA_START_ROW}:U{master_end}", LINEUP_STATUSES, strict=False),
        validation_one_of_list(master_ws, f"V{MASTER_DATA_START_ROW}:V{master_end}", INJURY_STATUSES, strict=False),
    ]

    # Team tabs
    depth_rows = build_depth_chart_rows()

    for team in PLL_TEAMS:
        code = team["Team_Code"]
        ws = sh.worksheet(team["Tab"])

        requests.append(
            validation_one_of_list(ws, "I10:I44", LINEUP_STATUSES, strict=False)
        )

        requests.append(
            validation_one_of_range(ws, "L10:L29", list_ranges[f"{code}_ALL"], strict=False)
        )

        requests.append(
            validation_one_of_list(ws, "M10:M29", INJURY_STATUSES, strict=False)
        )

        # Depth chart dropdowns
        for row_idx, depth_row in enumerate(depth_rows, start=50):
            pos = depth_row[1]
            requests.append(
                validation_one_of_range(ws, f"D{row_idx}:D{row_idx}", list_ranges[f"{code}_{pos}"], strict=False)
            )

        # Projected lineup dropdowns
        for spec in projected_card_specs():
            pos = spec["pos"]
            requests.append(
                validation_one_of_range(ws, spec["player_range"], list_ranges[f"{code}_{pos}"], strict=False)
            )

    # Tier tabs
    tier_tabs = {
        "Attack Tiers": ("ALL_A", 7),
        "Midfield Tiers": ("ALL_M", 7),
        "Defense Tiers": ("ALL_D", 7),
        "Goalie Tiers": ("ALL_G", 7),
    }

    for tab_name, (list_key, tier_count) in tier_tabs.items():
        try:
            ws = sh.worksheet(tab_name)
        except gspread.WorksheetNotFound:
            continue

        for tier_idx in range(tier_count):
            player_col = 2 + tier_idx * 4
            player_col_letter = col_to_letter(player_col)
            requests.append(
                validation_one_of_range(ws, f"{player_col_letter}6:{player_col_letter}35", list_ranges[list_key], strict=False)
            )

    # Specialists Tiers
    try:
        ws = sh.worksheet("Specialists Tiers")

        specialist_sections = [
            ("ALL_LSM", 7, 26),
            ("ALL_SSDM", 37, 56),
            ("ALL_FO", 67, 86),
        ]

        for list_key, start_row, end_row in specialist_sections:
            for tier_idx in range(6):
                player_col = 2 + tier_idx * 4
                player_col_letter = col_to_letter(player_col)
                requests.append(
                    validation_one_of_range(ws, f"{player_col_letter}{start_row}:{player_col_letter}{end_row}", list_ranges[list_key], strict=False)
                )

    except gspread.WorksheetNotFound:
        pass

    safe_batch_update(sh, requests)


# ============================================================
# UPDATE GOOGLE SHEET
# ============================================================

def authenticate_gspread():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def update_dashboard_status(sh, status, detail=""):
    try:
        ws = sh.worksheet("Dashboard")

        ws.update(
            range_name="A37:B42",
            values=[
                ["PLL Roster Update Status", status],
                ["Last Update Attempt", now_label()],
                ["Triggered By", TRIGGERED_BY],
                ["Trigger Source", TRIGGER_SOURCE],
                ["Detail", detail],
                ["Note", "Roster data updated in-place. Formatting and manual lineup/tier selections are preserved."],
            ],
            value_input_option="USER_ENTERED",
        )

    except Exception as e:
        print(f"Warning: could not update dashboard status: {e}")


def update_last_roster_update_cells(sh, timestamp=None):
    """
    Updates the visible last-successful-roster-update timestamp cells.

    Updates:
      - Dashboard!B4
      - Master Player Database!B4

    This function is called only after the roster update succeeds.
    """
    timestamp = timestamp or now_label()

    targets = [
        ("Dashboard", "B4"),
        ("Master Player Database", "B4"),
    ]

    for sheet_name, cell in targets:
        try:
            ws = sh.worksheet(sheet_name)
            ws.update(
                range_name=cell,
                values=[[timestamp]],
                value_input_option="USER_ENTERED",
            )
            print(f"Updated {sheet_name}!{cell} to {timestamp}")

        except Exception as e:
            print(f"Warning: could not update {sheet_name}!{cell}: {e}")


def update_google_sheet(pll_rosters_df, diagnostics_df):
    gc = authenticate_gspread()
    sh = gc.open_by_key(SPREADSHEET_ID)

    update_dashboard_status(sh, "Running", "Scrape complete. Updating sheet ranges.")

    master_ws = sh.worksheet("Master Player Database")
    lists_ws = sh.worksheet("Lists")

    sheet_df = standardize_for_sheet(pll_rosters_df)

    # Preserve manual master fields
    manual_master_map = read_existing_master_manual_fields(master_ws)
    sheet_df = preserve_master_manual_fields(sheet_df, manual_master_map)

    # Update Master Player Database values
    master_values = [SHEET_MASTER_COLUMNS] + sheet_df.fillna("").astype(str).values.tolist()

    master_ws.batch_clear([f"A{MASTER_HEADER_ROW}:W{MASTER_CLEAR_END_ROW}"])
    write_values(master_ws, f"A{MASTER_HEADER_ROW}", master_values)

    # Update Lists values
    lists_matrix, list_ranges, _ = build_lists_matrix(sheet_df)

    lists_ws.batch_clear([f"A1:ZZ{LISTS_MAX_ROWS}"])
    write_values(lists_ws, "A1", lists_matrix)

    # Update team roster tables only
    for team in PLL_TEAMS:
        code = team["Team_Code"]
        tab = team["Tab"]

        ws = sh.worksheet(tab)

        existing_manual = read_existing_team_roster_manual_fields(ws)
        roster_df = build_team_roster_values(team, sheet_df, existing_manual)

        ws.batch_clear([f"A{ROSTER_START_ROW}:J{ROSTER_END_ROW}"])
        write_values(ws, f"A{ROSTER_START_ROW}", roster_df.values.tolist())

        print(f"Updated roster table for {code}: {len(roster_df)} fixed rows.")

    # Reapply validations only. This does not change user selections.
    reapply_all_dropdown_validations(sh, list_ranges, len(sheet_df))

    # Build success summary.
    total_players = len(sheet_df)
    teams = sheet_df["Team Code"].nunique()
    successful_update_time = now_label()

    detail = (
        f"Updated {total_players} players across {teams} teams. "
        f"Last successful roster update: {successful_update_time}"
    )

    # Update visible last-successful-update timestamp cells.
    # These are the two cells requested:
    #   - Dashboard!B4
    #   - Master Player Database!B4
    update_last_roster_update_cells(sh, successful_update_time)

    # Update Dashboard status area.
    update_dashboard_status(sh, "Success", detail)

    print(detail)

    return sh


# ============================================================
# MAIN
# ============================================================

def validate_scrape(roster_df):
    issues = []

    if roster_df.empty:
        issues.append("NO_PLAYERS_SCRAPED")
        return issues

    total_players = len(roster_df)
    teams = roster_df["Team_Code"].nunique()

    if total_players < 120:
        issues.append(f"LOW_TOTAL_PLAYERS_{total_players}")

    if teams < 8:
        issues.append(f"MISSING_TEAMS_{teams}")

    counts = roster_df.groupby("Team_Code").size().to_dict()

    for team in PLL_TEAMS:
        code = team["Team_Code"]
        count = counts.get(code, 0)

        if count < 15:
            issues.append(f"{code}_LOW_ROSTER_COUNT_{count}")

    return issues


async def main():
    print("=" * 100)
    print("PLL ROSTER UPDATE STARTED")
    print("=" * 100)

    roster_df, diagnostics_df = await scrape_all_pll_rosters_async()

    print("\nScrape summary:")
    print("Total players:", len(roster_df))
    print("Teams:", roster_df["Team_Code"].nunique() if not roster_df.empty else 0)

    print(
        roster_df
        .groupby(["Team_Code", "Team"])
        .size()
        .reset_index(name="Roster_Count")
        .to_string(index=False)
    )

    issues = validate_scrape(roster_df)

    if issues:
        print("VALIDATION ISSUES:")
        for issue in issues:
            print("-", issue)
        raise RuntimeError("Scrape validation failed. Sheet was not updated.")

    sh = update_google_sheet(roster_df, diagnostics_df)

    print("=" * 100)
    print("PLL ROSTER UPDATE COMPLETE")
    print("=" * 100)
    print("Spreadsheet ID:", SPREADSHEET_ID)


if __name__ == "__main__":
    asyncio.run(main())
