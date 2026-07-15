# Release / milestone plan

Small, individually releasable milestones. One milestone = one coherent set of
commits, prefixed `R<n>:`. Tags follow `v0.<n>` per milestone; `v1.0` at R7.

| # | Milestone | Delivers | Acceptance criteria |
|---|---|---|---|
| R0 | Scaffold | Repo layout, README, CLAUDE.md, LICENSE, PLAN, doc skeletons, `legacy/` import | Repo browsable; legacy template usable as-is from `legacy/` |
| R1 | Spec first | Complete `docs/SPEC.md` for core parity, reverse-engineered from the legacy workbook | A competent dev could rebuild the workbook from SPEC.md alone, without opening the legacy xlsx |
| R2 | Generator (core parity) | `generate.py` + `model.py`: full template from code — all 15 sheets, masters, type-ahead dropdowns, 6 charts, Guide, sample data | Generated xlsx is feature-equivalent to `legacy/Family_Portfolio_Tracker.xlsx`; opens clean in Excel & LibreOffice; round-trip identity test passes |
| R3 | Cross-platform updater | `update.py` + `reader.py` + `fetch/` + `compute/xirr.py`: one command replaces all three PS1 scripts; timestamped backups; console log | On a workbook with sample data: prices, NAVs, masters and XIRR refresh correctly on Linux/Mac/Windows with only Python installed |
| R4 | One-click packaging | PyInstaller builds (Windows `.exe`, macOS `.command` + binary), release-zip layout, 1-page end-user README | A non-developer can download a zip, double-click, and get an updated workbook. First public-usable release |
| R5 | Quick wins | Red/green conditional formats (green ≥ 0 / red < 0 / amber stale), bond Maturity Value + coupon-aware XIRR, bank-master dropdown for FDs, "Expected @ FY-end" column | Each visible in the generated workbook; XIRR for a couponed bond matches a hand-checked value |
| R6 | FMV 2018 + delisted | `data/fmv_2018-01-31.csv` bundled; blank buy-price falls back to FMV (amber + comment); Stock_Master Status/Last-Traded-Date; stale-price flagging | Unknown-cost holding values correctly with visible flag; a delisted sample scrip shows last price, amber, excluded from day-change |
| R7 | Corporate actions → **v1.0** | BSE corporate-action fetch for held ISINs, split/bonus/consolidation adjustment engine, Corporate_Actions audit sheet, manual-action rows | Scenario tests pass (split, bonus, combined, future-dated, idempotent re-run); raw user rows never mutated |

Post-v1: see [ROADMAP.md](ROADMAP.md).

## Release artifact layout (from R4)

```
NetWorth-<version>-windows.zip
├── Family_Portfolio_Tracker.xlsx
├── Update Portfolio.exe
└── README.txt                      # 1-page quick start
NetWorth-<version>-macos.zip
├── Family_Portfolio_Tracker.xlsx
├── Update Portfolio.command        # tiny wrapper
├── networth-updater                # PyInstaller binary
└── README.txt
```
