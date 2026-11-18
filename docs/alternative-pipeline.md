# Proposal: produce `sdmt_database_import.xlsx` without `SDMT Report.xlsx`

## 1. Goal and constraints

Same final outcome as today — `sdmt_database_import.xlsx` with the 7 sheets
(POINT, DMTG, DMTT, Seismic_Downhole, Vs_Data, Vs_Parameters_Setup,
Vs_Parameters_Data) next to the `.sdmt` file — but:

- **neater**: small, single-purpose modules; each output sheet built by one pure
  function from typed records; no global mutable state; a schema registry instead
  of positional header offsets;
- **more robust**: parse the native XML/CSV files by element/attribute name, never
  by row offset; tolerate missing files/elements with explicit warnings; validate
  against the example;
- **no `SDMT Report.xlsx` required**: everything derived from the folder structure
  itself (see §3). If a zero-export pipeline is too costly for the DMT derived
  parameters, the `.uni` export (stable DMTUNI92 text format) is a drop-in,
  far more robust substitute for the xlsx (§4).

Keep the tkinter GUI as a thin wrapper; add a CLI (`python -m sdmt_reader job.sdmt`)
so it also runs headless.

## 2. What is actually available natively (verified on `examples/sdmt_pro_example`)

| Output data | Native source | Verified? |
|---|---|---|
| POINT (id, date, crew, remarks, Zw, flags) | `TEST.dat` | yes |
| DMTG (Zw, ITEM, BladeA/BladeB) | `TEST.dat` + latest `Calibrazione.dat` Item | yes |
| DMTT raw columns (TOP, A/B/C pressures, ΔA/ΔB) | `DMT.dat` `DatiDMT` + `Calibrazione.dat` | yes (114 rows match) |
| DMTT derived columns (P0/P1/P2, Id, Ed, Kd, Ud, M, Su, φ, K0, σ'h, OCR, soil, γ, σ'v, u0) | **not in native files** — computed by SDMT Pro; present in `SDMT-01.uni` / report DMT sheet | see §3/§4 |
| Seismic_Downhole (ITEM, Zw, SourceOffset_Vs) | `TEST.dat` + `DistHammer` from `VS/*.sdm` headers (fixes today's NaN) | yes |
| Vs_Data (TOP, VsAvg, VsLB/UB, VarCoeff, Repeatability) | `VS/TI.dat`: per shot `Vs = DeltaS/(DeltaT/1000)`; per depth mean/min/max; CV of rounded per-shot Vs | **yes — matches report exactly** (Z=1: 199.2/243.8 → "199, 244", Vs 222) |
| Vs_Parameters_Data (TOP, BulkDensity, ShearModAvg/LB/UB) | TI.dat Vs + density ρ; `G = ρ·Vs²/10⁶` MPa (verified: 1990·222²/10⁶ = 98.1 = report G₀) | formula yes; **ρ per depth only appears in the report** — see §5 |

## 3. Recommended design

```
sdmt_reader/
├── __main__.py / cli.py     # python -m sdmt_reader <job>.sdmt  → writes xlsx, prints summary
├── gui.py                   # existing tkinter UI, calls the same pipeline
├── parsers.py               # one function per native file → list[dict] / DataFrame
│                            #   parse_site(.sdmt), parse_test(TEST.dat), parse_dmt(DMT.dat),
│                            #   parse_calibrazione, parse_ti(VS/TI.dat), iter_sdm(VS/*.sdm)
├── uni.py                   # optional: parse DMTUNI92 .uni (fallback for derived DMT params)
├── marchetti.py             # optional: derive P0/P1/P2, Id, Ed, Kd, … from raw A/B/C + ΔA/ΔB + Zw + γ
├── model.py                 # dataclasses: Site, Test, DmtReading, Calibration, VsShot, SdmMeta
├── builders.py              # pure functions records → the 7 output DataFrames (schema registry)
├── schema.py                # canonical column names + aliases + units + coercion (from v1.5.1,
│                            #   but as data, not scattered code)
├── export.py                # DataFrames → xlsx via pandas.ExcelWriter (bold header, autosize)
└── tests/test_example.py    # run against examples/sdmt_pro_example; assert sheet shapes
                             #   POINT=1, DMTG=1, DMTT=114, Vs_Data=47, … and spot values
```

Pipeline: `parse_site → for each test folder: parse_test/parse_dmt/parse_calib/parse_ti/iter_sdm
→ build 7 tables → export`. Each parser returns typed records; builders are pure
`records → DataFrame`; a single `run_job(sdmt_path, out_path)` does everything, used
by both CLI and GUI. Errors are collected per file (warn + NaN) instead of only
printing, and logged.

Robustness measures vs v1.5.1:

- No xlsx parsing at all → no `header=2/skiprows=[3]` offsets, no `"X - DMT"` sheet
  naming, no Greek-header aliasing, no `Zabs` special case.
- Discovery by glob (`<test>/*.dat`, `VS/TI.dat`, `VS/*.sdm`) with explicit
  expected-name constants; missing pieces degrade to empty tables with a warning,
  not a crash.
- XML read by tag/attribute name; coordinates read from `<Coordinate><Latitude/>`
  (fixes the always-NaN POINT lat/long).
- `SourceOffset_Vs` taken from the `.sdm` `DistHammer` (fixes the NaN).
- Calibration chosen by parsed `Date`, not string max.
- Sheet/tables built from a declarative schema registry: canonical name, source
  aliases, dtype, units — adding a column is a one-line change.
- Tests pin the example: row counts per sheet, plus spot checks
  (Vs at Z=1 → 222 / LB 199 / UB 244; DMTT rows = 114; G₀ at Z=1 → 98.1).

## 4. The one hard part: DMTT derived parameters without any export

`DMTT`'s derived columns (ACorrected/BCorrected/CCorrected = P0/P1/P2, MaterialIndex,
Dilatometer, HorizontalStress, PoreIndex, Modulus, Su, FrictionAngle, OCR,
Description…) are computed by SDMT Pro's Marchetti formulas and exist **only in the
exports** (`SDMT Report.xlsx` DMT sheet or `SDMT-01.uni`). Two options:

**Option A (recommended): parse `SDMT-01.uni` instead of the xlsx.**
The `.uni` is a fixed DMTUNI92 text format (magic line, quoted header, count, then
`Z, P0, P1, γ, σ'v, u0, Id, Kd, Ed, Ud, M, Su, φ, soil-code` per depth) — trivial and
stable to parse, no row offsets. It already exists in the example job. Raw A/B/C and
ΔA/ΔB still come natively from `DMT.dat` + `Calibrazione.dat`, so DMTT = native raw
columns + `.uni` derived columns, joined on Z. This removes the `SDMT Report.xlsx`
requirement entirely (the `.uni` is a simple text export, not the formatted report).

**Option B (fully export-free): reimplement the Marchetti formulas in `marchetti.py`**
from `DMT.dat` raw readings + calibration + Zw + unit weight. The basics check out
against the example: `u0 = (Z − Zw)·9.81`, `P1 = B − ΔB`, `P2 = C − ΔA`, `σ'v` by
integrating γ above/below the water table, then Id/Ed/Kd/Ud/M/Su/φ/K0/OCR and the soil
classification per Marchetti (1980)/SDMT Pro. Risk: matching the software's exact
rounding/edge cases (e.g. the report's `Var Coefficient` and soil strings); mitigate
with the example-based tests. REFERENCE column would change to cite the formulas
implemented.

Option A is the sweet spot for robustness; Option B is there if exports must be
eliminated completely. Both share the same builders — only the DMTT source changes.

## 5. Open point: bulk density ρ for `Vs_Parameters_Data`

`ρ` (BulkDensity) per depth appears only in the report's S-wave sheet in the
example; the native files carry only a single `Gamma` marker in `TEST.dat` (empty
here). So either (a) take ρ from the `.uni` γ per depth (ρ ≈ γ·1000/9.81, interpolated
to Vs depths), or (b) if a soil-layer profile exists in real job folders, parse it
from there, or (c) keep `Vs_Parameters_Data.BulkDensity/ShearMod*` optional and
compute them only where ρ is known. Recommend (a), with (c) as graceful fallback.

## 6. Migration path

1. Add the `sdmt_reader` package alongside the existing scripts; don't modify
   `sdmt_pro_reader_v1.5.1.py` yet.
2. Implement parsers + builders + CLI; wire the GUI buttons to `run_job`.
3. Validate: run on `examples/sdmt_pro_example`, diff the new
   `sdmt_database_import.xlsx` against the committed example sheet-by-sheet
   (DMTT/Vs_* should match within rounding when using the `.uni`; raw columns
   exact).
4. Keep `SDMT Report.xlsx` support as an optional fallback source for one release,
   then drop it.
