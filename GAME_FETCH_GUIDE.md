# Games Sheet - Fetch & Write Guide (Simplified)

This project fetches your Chess.com games and writes one row per game into the `Games` sheet. This guide documents the minimal "how it works" and what each column means, without callback, daily stats, sorting/dedupe, or ETag complexity.

## How to run a full fetch (all games)

1) In the spreadsheet, open the Extensions → Apps Script project (this repo).  
2) In the Script Editor, run `setupSheets()` once to create/align the `Games` sheet headers.  
3) Run `fetchAllGamesInitial()` to fetch your entire archive and write all games into `Games`.

Notes:
- This guide assumes no ETag/last-URL logic. It simply describes the end state after a full fetch.  
- Updates (incremental) can be run later with `fetchChesscomGames()`, but that’s beyond this minimal scope.

## What the `Games` sheet contains (columns)

Order is left→right, as written by the code.

- Game ID: The game identifier (string), taken from the game URL last segment.
- Start: Local epoch seconds when the game started. If not present in PGN, this is inferred from End and Duration when available; otherwise blank.
- End: Local epoch seconds when the game ended. Derived from game’s end timestamp.
- Date: Google Sheets serial date for the local day of the game end. Used for simple grouping by day.
- Time: Seconds since local midnight at game end (integer). Date + Time fully specifies the local end.
- Archive (MM/YY): Convenience tag from End local time (e.g., "10/25").
- Is Live: TRUE for live games (Bullet/Blitz/Rapid/live variants), FALSE for Daily/Correspondence.
- Time Class: Display time class from API (Bullet, Blitz, Rapid, Daily). Kept for reference.
- Format: Short code for the rating pool (preferred order: Bul, Blz, Rap, Dly, D960, L960, Bug, Czh, KotH, 3Chk).
  - Mappings: Bullet→Bul, Blitz→Blz, Rapid→Rap, Daily→Dly, Daily960→D960, Live960→L960, Bughouse→Bug, Crazyhouse→Czh, KingoftheHill→KotH, Threecheck→3Chk.
- Base Time (s): Parsed from `time_control`; for Daily (e.g., `1/86400`) this will be 0.
- Increment (s): Parsed from `time_control` (e.g., `60+1` → 1); for Daily this is 0.
- Correspondence Time (s): For Daily `1/86400`, this becomes 86400. For live, 0.
- Is White: TRUE if you played White.
- Opponent: Opponent username.
- My Rating: Your rating after the game, from the games archive API.
- Opp Rating: Opponent rating after the game (from the API).
- Rating Before: Last known rating for the same Format immediately prior to this game.  
  - Filled quickly during processing using an in-memory map of "last seen rating per Format" within the current write batch.  
  - For the very first game encountered in a Format, this may be blank.
- Delta: My Rating − Rating Before, when both are available; otherwise blank.
- Outcome: Win/Draw/Loss (from result mapping).
- Termination: Termination reason, mapped (e.g., Timeout, Resignation).
- ECO: ECO code extracted from PGN if present.
- Opening Name: Looked up from `Openings DB` using the ECO URL/slug; stores a friendly name.
- Opening Family: Looked up from `Openings DB`; broader opening family.
- Ply Count: Number of plies (full moves × 2) parsed from PGN move list.
- tcn: Encoded move list from the archive JSON (`tcn`), if present.
- clocks: Move clocks encoded as a compact base‑36 decisecond string, parsed from PGN `%clk` tags when available. For Daily, typically empty.

## How values are computed (at a glance)

- Start/End (local epoch seconds):
  - End is converted to local time from the game’s end timestamp.  
  - Start is taken from PGN headers if available; otherwise inferred from End − Duration; otherwise blank.

- Date/Time:
  - Date is the Sheets serial date computed from local End (days since 1899‑12‑30).  
  - Time is seconds since local midnight at End.

- Time control parsing (`Base Time (s)`, `Increment (s)`, `Correspondence Time (s)`):
  - `X+Y` → Base=X seconds, Increment=Y seconds, Correspondence=0.
  - `X` → Base=X seconds, Increment=0, Correspondence=0.
  - `A/B` (Daily) → Correspondence=B seconds, Base=0, Increment=0.

- Is Live: TRUE if Time Class ≠ Daily; FALSE if Daily.

- My Rating/Opp Rating: Taken from the archive JSON (`white.rating` / `black.rating`). The side is chosen based on your color.

- Rating Before/Delta:
  - Uses a fast per‑Format cache while writing the current batch (no sheet scans per row).  
  - Rating Before = last My Rating seen for that Format in this batch (or blank if none yet).  
  - Delta = My Rating − Rating Before when both present.

- Outcome/Termination: Mapped from API result strings.

- ECO / Opening Name / Opening Family:
  - ECO from `[ECO "..."]` in PGN, ECO URL from `[ECOUrl "..."]`.  
  - Opening Name and Family looked up against `Openings DB` by slug.

- Ply Count: Count of SAN tokens in PGN’s move section (ignoring annotations/comments), or from parsed representation.

- tcn / clocks:
  - tcn: direct from the archive JSON when present.  
  - clocks: parsed from PGN `%clk` tags, encoded to base‑36 deciseconds and dot‑delimited.

## Running a simple full import (ignore advanced concerns)

- Run `setupSheets()`  
- Run `fetchAllGamesInitial()`  
- When done, the `Games` sheet will be populated with one row per game and all columns above.

That’s it. You can later add callback enrichment, daily stats, or event adjustments, but for now this document covers the basics you asked for.
