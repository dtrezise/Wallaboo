# Mine Trap — Company Portfolio Tracker

## Portfolio record

| Field | Current company-side record |
| --- | --- |
| Game title | **Mine Trap** |
| Project type | Standalone boardgame product/IP |
| Portfolio role | **Flagship development project** (working designation) |
| Development gate | Polished physical prototype preparation and rules validation |
| Source-of-truth repository | [`/Users/dan/Documents/GitHub/MineTrap`](../../MineTrap/) |
| Remote repository | [github.com/dtrezise/MineTrap](https://github.com/dtrezise/MineTrap) |
| Company tracker | This file; it summarizes and links but does not own Mine Trap assets or rules |
| Last company review | 2026-07-17 |
| MineTrap evidence point | `main` at `dbcb854` (2026-05-11, “Raise River card badge”) |

> **Repository boundary:** MineTrap remains the authoritative product repository. Mr Wallaboo Games tracks portfolio status, gates, risks, and production readiness here. Do not copy the MineTrap asset library into this repository.

## Current status summary

Mine Trap is a cooperative survival and exploration board game for 1–4 players. Players explore a collapsing coal mine, gather Coal, uncover a buried Elevator, and escape before the shared air supply runs out.

The project has a substantial digital prototype package: a live rules draft, a prototype roadmap, official game-test #22 component counts, 72 tracked card images, and five pawn/component reference images. The latest committed rules update was made on 2026-05-11. Those rules incorporate badged card variants and contain a buildable component inventory.

The project is **prototype-ready in concept, but not yet print- or production-ready**. The rules still identify themselves as a live draft; only Turn Structure and Inventory Cooldown carry explicit `Locked` badges. The repository does not currently contain a standalone print manifest, current print sheets/PDFs, a committed playtest log, or evidence that the latest component set has been assembled and blind-tested. The older roadmap also contains answers that conflict with the newer rules and stores user-entered status in browser local storage, so its live UI state is not recoverable from Git.

### Company status

- **Flagship designation:** Active working flagship for Mr Wallaboo Games.
- **Company gate:** Prototype and validation; not manufacturing-ready.
- **Confidence:** Medium. The rules and component assets are well advanced, but current playtest evidence and production files are incomplete.
- **Next company review trigger:** Canonical rules reconciliation plus a committed print manifest and current prototype test report in MineTrap.

## Latest rules document

- **Current rules authority:** [`rules.html`](../../MineTrap/rules.html)
- **Document label:** “Physical Prototype Rules Draft” / “Live rules draft”
- **Latest rules commit:** `6325e78` — 2026-05-11, “Update prototype rules with badged cards”
- **Latest count basis:** Official game test #22, incorporated in commit `7cd9d2c` on 2026-05-11
- **Important qualification:** The rules contain newer decisions than [`index.html`](../../MineTrap/index.html). When the two disagree, treat `rules.html` as the later evidence until MineTrap explicitly establishes a versioned canonical rules policy.

## Prototype component status

| Component area | Status | Evidence / gap |
| --- | --- | --- |
| Rules | Advanced draft | Playable flow, setup, card rules, examples, variants, and resolved designer answers are in `rules.html`; final rules lock is not documented. |
| Mine and item cards | Art library present | 72 tracked PNG card assets, including mine-, starter-, and Tool-Chest-badged variants. The files use multiple pixel dimensions and are not yet normalized into a documented print specification. |
| Pawns and air marker | Specified | Four player pawn references plus one white Air Counter marker; rules specify 29 mm × 13 mm wooden pawns. |
| Component counts | Embedded draft manifest | Game-test #22 counts are in `rules.html`, including the mine deck, starting inventory, Tool Chest, rule cards, and air marker. No standalone print manifest exists. |
| Print layout | Not current in repo | No tracked print-ready PDF or production sheet was discovered. The roadmap says older printable PDFs should be replaced with sheets generated from the locked card set. |
| Physical prototype | Unverified | The roadmap includes a build checklist, but no committed record confirms that the current rules/assets were printed, assembled, and tested together. |
| Playtest evidence | Partial reference only | The rules cite “official game test #22,” but no standalone playtest notes, results log, or issue register was discovered. |
| Player aids | Assets partially present | Turn and progress-circle reference images exist; the roadmap still calls for a one-page teach sheet/player-aid outline. |

### Current standard-game inventory recorded in the rules

- Mine cards: 4 Pit, 4 River, 4 Tool Chest, 2 two-Coal, 2 three-Coal, 4 Map, 5 Spider, 1 Elevator, plus 6 High Rubble and 6 standard Coal reduced to 8 mixed cards during setup.
- Starting items: 8 player-color Pickaxes and 4 player-color Coal cards.
- Tool Chest: 3 Pickaxes, 3 Ropes, and 3 Spider Repellents.
- Other: 4 rule/reference cards; four colored player pawns; white Air Counter pawn; pile locators; furnace/Coal area.
- Optional/variant content includes Cavern, easier Pit, Air Pocket, Air Masks, chained flooding, Coal difficulty, and Tool Chest tuning.

## Open design questions

1. **Canonical air scaling:** `rules.html` says to choose an Air Counter matching player count (18/15/13/11), but the Core Rules and game-test #22 baseline say the standard game starts at 11. Decide whether 11 is universal or only the four-player value.
2. **Player-count balance:** Confirm whether the fixed 12-Coal target and fixed 3/3/3 Tool Chest mix work for all 1–4 player counts or require scaling.
3. **Roadmap/rules reconciliation:** Resolve older `index.html` answers that differ from `rules.html`, including Elevator burial, Dig Deep reveal behavior, and Spider reveal/movement behavior.
4. **Rules-to-card copy lock:** Identify every card whose printed wording is stale, then establish one final card-text table before new exports.
5. **Hazard and pacing validation:** Test the 12-Coal target, air economy, Spider burden, River/flooding pressure, and win/loss pacing across player counts.
6. **Core versus optional content:** Lock whether Cavern, alternate Pit, Air Pocket, Air Masks, and the extended Spider set belong in the standard prototype or remain variants.
7. **Production specification:** Decide card size, bleed, safe area, backs, duplex method, stock/sleeve plan, cut guides, color profile, and final table footprint.
8. **Blind-teach quality:** Confirm that players can set up and complete a game from the rules and player aids without designer coaching.

## Recent decisions and changes

- **2026-05-11 — River badge adjusted:** The mine-deck badge on the River card was raised (`dbcb854`).
- **2026-05-11 — Badged prototype set integrated:** Mine, starter, and Tool Chest variants were added and the rules were updated to display/use them (`6325e78`).
- **2026-05-11 — Pawn specification added:** Blue, yellow, green, and red player pawns plus the white Air Counter marker were documented; pawn size is 29 mm × 13 mm (`68efe3b`).
- **2026-05-11 — Official test #22 counts adopted:** Fixed prototype counts, a 3/3/3 Tool Chest, the standard 12-Coal goal, and an 11-air test baseline were folded into the rules (`7cd9d2c`).
- **2026-05-11 — Air Pocket art revised:** The optional Air Pocket card received an art update (`731a5a9`).
- **2026-04-27 to 2026-04-29 — Rules clarifications:** Map movement, Spider reveal/movement, Dig Deep edge cases, Rope crossing, progress-circle restrictions, and buried Elevator placement were incorporated into the current rules draft.

## Next production steps

These actions belong in the MineTrap source repository; this file should only record their completion.

1. Reconcile `index.html` with `rules.html` and declare one versioned canonical rules source.
2. Resolve the air-scaling and player-count questions, then record the chosen standard and variants in the rules.
3. Create a standalone print manifest in MineTrap with exact quantities, fronts, backs, dimensions, source asset, and print-sheet placement for every component.
4. Audit every current image against final card text; replace stale exports and normalize the print specification without discarding source art.
5. Generate current prototype sheets/PDFs in MineTrap with bleed, cut guides, backs, and a documented assembly method.
6. Assemble the current physical prototype and commit a dated build record with the exact rules and manifest versions used.
7. Run internal tests across 1–4 players; log setup time, playtime, Coal/air outcomes, hazard burden, rules questions, and component failures.
8. Run blind tests using only the rules and aids; convert findings into a bounded rules/card-text revision.
9. Lock the one-page teach aid and final prototype component manifest.
10. After validation, obtain prototype/manufacturing quotes and update the Wallaboo financial model with Mine Trap’s actual bill of materials and channel assumptions.

## Key MineTrap links

- [MineTrap README](../../MineTrap/README.md)
- [Prototype roadmap / status canvas](../../MineTrap/index.html)
- [Latest rules draft](../../MineTrap/rules.html)
- [Card asset library](../../MineTrap/assets/cards/)
- [Pawn and component references](../../MineTrap/assets/components/)
- [MineTrap repository on GitHub](https://github.com/dtrezise/MineTrap)

### Files not discovered in the current MineTrap repository

- Standalone print manifest
- Current print-ready sheets or PDFs
- Dedicated playtest notes/results log
- Production bill of materials or supplier quote package

## Tracking protocol

- Update this page from committed MineTrap evidence, not recollection or browser-local notes.
- Record the MineTrap commit hash and review date whenever status changes.
- Link to authoritative MineTrap files; do not duplicate its rules, source art, print files, or playtest archive here.
- Treat company-side status labels as portfolio decisions. Treat game rules, components, and design history as MineTrap decisions.

