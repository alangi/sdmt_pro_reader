# SDMT Pro Reader — Data Model

How the Marchetti SDMT job folder fits together, and how
`sdmt_pro_reader_v1.5.1.py` (latest; `v1.4` is the earlier variant) turns it into
`sdmt_database_import.xlsx`.

Reference example: `examples/sdmt_pro_example/` (one site `SDMT-01`, one test folder
`TEST_1`).

---

## 1. Repo layout

```
sdmt_pro_reader/
├── sdmt_pro_reader_v1.4.py          # earlier version of the GUI tool
├── sdmt_pro_reader_v1.5.1.py        # current version (tkinter GUI)
├── examples/
│   └── sdmt_pro_example/            # one "job" as produced by SDMT Pro
│       ├── SDMT-01.sdmt             # EXPORT/native site file (XML) — the entry point
│       ├── SDMT-01.uni              # EXPORT — DMTUNI92 processed DMT data (CSV-like)
│       ├── SDMT Report.xlsx         # EXPORT — processed report (DMT / S-wave / Inclination sheets)
│       ├── sdmt_database_import.xlsx# EXPORT — final output of this tool (7 sheets)
│       └── TEST_1/                  # one folder per test (named in the .sdmt <Tests>)
│           ├── TEST.dat             # test metadata (XML)
│           ├── DMT.dat              # DMT setup + raw A/B/C readings (XML)
│           ├── Calibrazione.dat     # blade calibration history (XML)
│           ├── INCL.dat             # inclination readings (XML) — not currently used
│           └── VS/
│               ├── TI.dat           # seismic processing setup per depth (XML)
│               └── 1.sdm … 97.sdm   # raw seismograms, one per shot (XML)
└── docs/
```

The GUI (`v1.5.1`) flow: **Open a File** → pick a `*.sdmt` →
`process_filename()` parses everything into pandas DataFrames →
**Database Export** writes `sdmt_database_import.xlsx` next to the `.sdmt`.
Extra buttons dump debug (`sdmt_import.xlsx`), show DataFrames, and export the
raw `.sdm` waveforms as text (`Individual VS`, `-S1-A.txt` etc.).

---

## 2. The native folder structure (the real data)

### `SDMT-01.sdmt` — site file (XML), the entry point
```xml
<Site><Name>251109</Name><Firm>…</Firm><Customer>…</Customer><Job>…</Job>
  <Tests><Test Folder="TEST_1" /></Tests>
</Site>
```
Only the `<Tests>/<Test Folder="…">` list is used — it enumerates the per-test
folders. The site header (Name/Firm/Customer/Job) is currently ignored.

### `TEST_1/TEST.dat` — per-test metadata (XML)
Key elements used: `Name` (→ `POINT_ID`, e.g. `SDMT-01` — note: this differs from
the folder name `TEST_1`; the code maps between them via `TEST_N`), `Date`,
`Operator`, `Zw` (groundwater depth), `IsDmtActive`, `IsVsActive`, `IsVpActive`,
`HasCalibration`, … plus `U`/`Gamma` unit markers and a `Coordinate` element
(`Latitude`/`Longitude` children — currently **not** parsed; the code looks for a
`GPS` element that doesn't exist in this format, so POINT.Latitude/Longitude are
always NaN in the example).

### `TEST_1/DMT.dat` — DMT setup + raw readings (XML)
- Header/setup: `ZmCAL`, `ZmAB`, `ZmC`, `So`, `GammaTop`, `DeltaA`, `DeltaB`
  (→ `dmt_setup_df`).
- `<DatiDMT><DatoDMT Z="2.40" A="165" B="793" C="-37"/>…` — 114 raw readings:
  depth `Z` and gauge pressures `A`, `B`, `C` (kPa) (→ `dati_dmt_df`).
  **All derived soil parameters are computed by SDMT Pro, not stored here.**

### `TEST_1/Calibrazione.dat` — calibration history (XML)
`<Calibrazioni><Item Date="…" DeltaA="42" DeltaB="79"/></Calibrazioni>` — the code
takes the **latest by Date** → `DeltaA`/`DeltaB` (→ `calib_df`, becomes
`DMTG.BladeA/BladeB` and the per-row `DMTT.DeltaA/DeltaB`).

### `TEST_1/INCL.dat` — inclination readings (XML)
`<Inclinazione>` per depth (`Z`, `IX/IY/IZ`, `DataOra`, …). Parsed by nothing in
v1.5.1 — the report's "Inclination" sheet is ignored too.

### `TEST_1/VS/TI.dat` — seismic processing results (XML)
```xml
<Processings>
  <Processing Z="1.5">
    <Seismogram Selection="g*" FileNumber="3" DeltaS="0.4595…" DeltaT="1.7285…">
      <Filter …/><Window …/><Rephase …/>
    </Seismogram>
    <Seismogram Selection="g" FileNumber="4" …/>
  </Processing>…
</Processings>
```
47 depths × 2 shots (selections `g*` and `g`). **This is the native source of Vs:**
per shot, `Vs = DeltaS / (DeltaT / 1000)` m/s (DeltaS in m, DeltaT in ms).
Verified against the example: at Z=1.0 the two shots give 199.2 and 243.8 m/s →
report shows Repeatability "199, 244" and Vs = 222 = the mean. Per depth:
`VsAvg` = mean, `VsLB/VsUB` = min/max (the "Repeatability" pair), `VarCoefficient`
≈ coefficient of variation of the rounded per-shot Vs values.

### `TEST_1/VS/*.sdm` — raw seismograms (XML), one file per shot (97 files)
Header metadata: `ZRef` (test depth), `WaveType` (S/P), `DistHammer` (source offset,
e.g. 0.5 m), `DateTime`, `SampleTime` (µs), `Samples`, `Sensor ID=1/2 ZSensor=…`,
`ExternalTrigger Delay`, then `<S S1="…" S2="…"/>` sample pairs (8-bit, 0–255,
two geophones). Currently only used to write text waveform exports
(`Individual VS/{n}-S1-A.txt` etc.); `DistHammer` from here is **not** fed into
`Seismic_Downhole` (that sheet's `SourceOffset_Vs` is NaN in the example because
`TEST.dat` has no `DistHammer` element).

---

## 3. The exports (derived data — not part of the folder structure)

### `SDMT Report.xlsx` — processed report (the fragile dependency)
One sheet per test and type, named `"<TEST NAME> - DMT"`, `"… - S-wave"`,
`"… - P-wave"`, `"… - Inclination"`. Each sheet has a title row, a blank row, a
**header row**, a **units row**, then data. The code reads them by **fixed
position**: `header=2, skiprows=[3]` for DMT/S-wave, `header=3, usecols=range(4),
skiprows=[4]` for P-wave, splitting sheet names on `" - "`.
- **DMT sheet** (25 cols): `Z, Δ A, Δ B, A, B, C, γ, σ v, σ' v, U₀, P₀, P₁, P₂,
  Id, Ed, Kd, Ud, M, Su, φ, K₀, σ' h, σ h, OCR, Soil description` — i.e. the raw
  readings **plus all Marchetti-derived parameters**, computed by SDMT Pro.
- **S-wave sheet** (6 cols): `Z, Vs, Repeatability, Var Coefficient, ρ, G₀` —
  `VsAvg`, `VsLB/VsUB` (from Repeatability "199, 244"), `VarCoefficient`, bulk
  density `ρ` (kg/m³ — **only appears here**, not in any native file), and
  `G₀ = ρ·Vs²/10⁶` MPa (verified: 1990·222²/10⁶ = 98.1).
- **P-wave sheet**: parsed but never used in any output table.

### `SDMT-01.uni` — DMTUNI92 export
Plain CSV-ish text: `DMTUNI92` magic, quoted header lines (firm, client, job,
date, test name), a count line, then one line per depth:
`Z, P0, P1, γ, σ'v, u0, Id, Kd, Ed, Ud, M, Su, φ, soil-code` — the same
software-computed parameters as the DMT sheet, in a far more stable format.

### `sdmt_database_import.xlsx` — the final output of this tool
Seven sheets (written by `export_dataframes_to_excel`):

| Sheet | Rows (example) | Source |
|---|---|---|
| `POINT` | 1/test | `TEST.dat`: Name→POINT_ID, Date→StartDate, Operator→Crew, Remark→Remarks, lat/long (currently always NaN) |
| `DMTG` | 1/test (only if `IsDmtActive≠0`) | `TEST.dat` Zw→GroundWaterLevel, IsDmtActive→ITEM + `Calibrazione.dat` latest ΔA/ΔB→BladeA/BladeB |
| `DMTT` | 114 (= DatoDMT count) | **SDMT Report.xlsx DMT sheet**: Z→TOP, A/B/C→pressures, ΔA/ΔB, and all derived params (P0→ACorrected, P1→BCorrected, P2→CCorrected, Id→MaterialIndex, Ed→Dilatometer, Kd→HorizontalStress, Ud→PoreIndex, M→…, Su→UndrainedShearStrength, φ→FrictionAngle, K0→…, σ'h→…, OCR, Soil description→Description); ITEM from IsDmtActive; REFERENCE constant |
| `Seismic_Downhole` | 1/test (only if `IsVsActive≠0`) | `TEST.dat`: IsVsActive→ITEM, Zw→GroundWaterLevel, DistHammer→SourceOffset_Vs (**NaN** — not in TEST.dat; it lives in the `.sdm` files) |
| `Vs_Data` | 47 (= S-wave rows = TI.dat depths) | **SDMT Report.xlsx S-wave sheet**: Z→TOP, Vs→VsAvg, Repeatability→VsLB/VsUB (min/max), Var Coefficient→VarCoefficient, ITEM from IsVsActive |
| `Vs_Parameters_Setup` | 1/test | `TEST.dat`: IsVsActive→ITEM, REFERENCE="Marchetti TI", Methods="Marchetti True Interval" |
| `Vs_Parameters_Data` | 47 | S-wave sheet: Z→TOP, ρ→BulkDensity, G₀→ShearModAvg; ShearModLB/UB = ρ·VsLB²/10⁶, ρ·VsUB²/10⁶ (verified: 1990·199²/10⁶ = 78.8) |

`POINT_ID` vs `TEST_N`: the test **folder** name (`TEST_1`) is `TEST_N`; the test
**name** inside `TEST.dat` (`SDMT-01`) is `POINT_ID`. The report sheets are named by
`POINT_ID`. `transform_raw_data()` joins the two via lookups built from `TEST.dat`.

---

## 4. Pipeline as implemented (v1.5.1)

```
.sdmt ──► read_sdmt_xml ──► test folder names
   ├─ TEST_1/TEST.dat ──────► test_meta_df ─┐
   ├─ TEST_1/DMT.dat ───────► dmt_setup_df, dati_dmt_df (raw Z/A/B/C)
   ├─ TEST_1/Calibrazione.dat ► calib_df (latest ΔA/ΔB)
   ├─ SDMT Report.xlsx ─────► s_wave_df, p_wave_df, dmt_data_df  ← the xlsx dependency
   ├─ TEST_1/VS/TI.dat ─────► vs_setup_df (per-shot processing params)
   └─ TEST_1/VS/*.sdm ──────► vs_sdm_info_df (+ optional txt waveform export)
                │
        transform_raw_data()  (POINT_ID ↔ TEST_N joins, merges calib + first SdmNumber into test_meta)
                │
        create_POINT_df / DMTG / DMTT / Seismic_Downhole / Vs_Data / Vs_Parameters_*
                │
        export_dataframes_to_excel() ──► sdmt_database_import.xlsx (7 sheets)
```

Schema handling: `normalize_name()` strips Greek letters/subscripts from report
column headers (`Δ A`→`deltaa`, `σ' v`→`sigma_v_eff`, …) and
`select_and_rename()` maps them to canonical names, filling missing columns with
NaN.

## 5. Fragilities / quirks worth knowing

1. **The whole derived dataset (DMTT, Vs_Data, Vs_Parameters_Data) comes from
   `SDMT Report.xlsx`**, read by hard-coded row offsets and the `"X - DMT"` /
   `"X - S-wave"` sheet-name convention. Any change in the report layout (extra
   title row, renamed sheet, new column) silently shifts or drops data.
2. `Zabs` column dropped by name if present; Greek headers handled by
   normalization — but only for the aliases currently listed.
3. `POINT.Latitude/Longitude` are always NaN: the parser looks for a `<GPS>`
   element, but `TEST.dat` stores coordinates in `<Coordinate><Latitude/>…`.
4. `Seismic_Downhole.SourceOffset_Vs` is NaN: `DistHammer` is read from
   `TEST.dat` (absent) instead of the `.sdm` headers where it actually lives.
5. `INCL.dat` and the report's Inclination sheet are ignored entirely.
6. P-wave data is parsed but unused. `dati_dmt_df` (raw DMT) is parsed and shown
   in debug but not used in the output — the output DMTT comes from the report.
7. `Calibrazione.dat` picks the latest item by Date **string** comparison;
   missing files only print a warning and yield NaN columns.
8. Multi-test jobs: everything loops over `<Test Folder>`s, but the report sheet
   names must match `TEST.dat/Name` exactly for the POINT_ID↔TEST_N mapping.
9. v1.4 vs v1.5.1: v1.5.1 adds the Calibrazione pipeline, the
   `normalize_name`/`select_and_rename` schema layer, the transformed-debug
   export (`sdmt_import.xlsx`), and the GUI info buttons; the overall shape is
   the same.
