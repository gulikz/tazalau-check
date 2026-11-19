# Soft-collection segmentation: auto loans, DPD 1-30

`soft_collection_segmentation.ipynb` scores every entry into delinquency with a PD model, then assigns a treatment segment. It uses a rule overlay first, then a risk x capacity x engagement matrix. The current DPD 1-30 stock comes out with its segment, treatment, challenger arm and compliance dates.

## Files
| file | what |
|---|---|
| `soft_collection_segmentation.ipynb` | the notebook, saved without outputs |
| `demo_data.py` | synthetic generator with the real schema of the 13 tables (fake people); `make_demo_tables()` |
| `sql/hes_pkb_monitoring_agg.sql` | Vertica aggregate of `dds.hes_pkb_monitoring_event` (long format, as of `:snapshot_date`). **The IIN and event-timestamp column names are placeholders** (`*__PLACEHOLDER`) |
| `out/` | run outputs. Git ignores this folder because the files contain IINs |

## Method in brief
1. **Episodes and target.**
   - `ovrd` gives a monthly maximum DPD per contract. Each row implies an overdue start, `start = act_date - (dpd - 1)`. A new episode begins when the start date jumps by more than 5 days.
   - `obs_date` is the first overdue day.
   - `bad` = the contract reaches 90+ DPD within 6 months. Secondary targets: `self_cure30` and `roll31_60d`.
2. **Features.**
   - There are 13 feature blocks, `fb_*` (contract, DPD history, client profile, internal and external debits, enforcement, calls, app logins, exclusions, bankruptcy, applications, rejections, PKB). Each uses only events **before** `obs_date`.
   - An **as-of self-test** rebuilds the data as it looked on an entry date and checks that every block gives identical features.
   - Sources that start late (geo, vnpt, vneshpt, PKB) are used for rules only (`MODEL_BLOCKS='auto'`).
3. **Model.**
   - LightGBM, with a fallback to sklearn `HistGradientBoosting`, benchmarked against a logistic regression.
   - Time split: OOT = the last 3 labelled months. Validation is grouped by contract.
   - Train PDs are **out-of-fold**, so the grade cutpoints and the backtest are not over-fitted.
4. **Grades** G1-G5 from train PD quantiles 40/65/85/95%. The monotonicity check uses a significance test.
5. **Segments.**
   - Rules R01-R09 first: deceased, bankruptcy, bankruptcy attempt, social exclusion lists, enforcement, abroad/skip, FPD/fraud, gambling, collateral.
   - Then S1 self-cure, S2 low-risk remind, S3 standard, S4 hardship, S5 high-risk contactable and S6 high-risk unreachable.
   - Treatments live in `SEGMENT_TREATMENT` and follow the Kazakhstan guardrails written into the notebook (section 9.0): the 10-day notice, the art. 61 change-of-terms window, call and visit caps, and the debit limit of 50% / 2xПМ.

## How to run
- **Your kernel (way a).** Load the tables under their usual names (`contr`, `ovrd`, `ver_df_all_2024_2026`, `bankrupt`, `ocrm_call`, `vneshpt_ot`, `vnpt`, `excl`, `vneshpt`, `geo`, `dm`, `dm_rej`, optionally `hes_pkb_agg`), then run `%run -i soft_collection_segmentation.ipynb`. Your DataFrames are read and never modified.
- **Parquet (way b).** Set `CONFIG['DATA_DIR']` and the file names in `CONFIG['FILES']` (globs allowed, e.g. `dm_*_rejected.parquet`). Only the needed columns and the population's keys are read, with pyarrow push-down.
- **DEMO (way c).** With no tables found, `MODE='auto'` uses `demo_data.py`. It runs in about 25 s with 8,000 synthetic contracts.
- **Headless.** The kernel inherits `SEG_*` environment variables, which override `CONFIG`: `SEG_MODE`, `SEG_DROP_TABLES` (comma list), `SEG_OUT_DIR`, `SEG_DATA_DIR`, `SEG_DEMO_N_CONTRACTS`, `SEG_USE_LIGHTGBM` and `SEG_STRICT_BLOCKS` (1 = a failing block raises). For example:
  ```bash
  SEG_DROP_TABLES=geo,vnpt python -c "import nbformat; from nbclient import NotebookClient; \
  nb = nbformat.read('soft_collection_segmentation.ipynb', 4); NotebookClient(nb, timeout=1200, kernel_name='python3').execute(); \
  nbformat.write(nb, 'executed.ipynb')"
  ```

Only `contr` and `ovrd` are required. A missing or broken optional table switches its block off with a warning, and the rules that depend on it cannot fire.

## Outputs (`out/`)
- `segments_current_stock.parquet` / `.xlsx`: one row per stock contract. Columns include `pd`, `risk_grade`, `segment`, `segment_reason`, `treatment`, `arm` (champion/challenger), `rules_hit`, `capacity`, `engaged`, `debitable_funds`, `no_credit_offer`, `third_party_contact_flag`, `notify_due_date`, `restructuring_window_end` and `top_drivers` (TreeSHAP).
- `segments_history.parquet`: labelled episodes with PD, grade, segment, cluster and the outcomes.
- `segment_report.xlsx`: sheets `data_quality`, `block_coverage`, `model_metrics`, `grades`, `segments_backtest`, `segments_current`, `clusters`, `feature_importance`, `psi_dropped`, `monthly`, `dealers`, `monitoring`, `treatments`, `load_log`.
- `model/`: `lgbm_model.txt`, `model_bundle.pkl` (model, LR benchmark, features, category levels, cutpoints, config), and `features.json`, `config.json`, `cutpoints.json`.
- `charts/*.png` (5 charts) and `run_summary.json`.

## Main knobs (`CONFIG`)
| knob | default | meaning |
|---|---|---|
| `MODE` | `auto` | `auto` = globals -> parquet -> DEMO; `real`; `demo` |
| `SOFT_DPD_MIN/MAX` | 1 / 30 | soft-collection window of the current stock |
| `BAD_DPD`, `HORIZON_MONTHS` | 90 / 6 | target definition |
| `EPISODE_GAP_DAYS` | 5 | episode split tolerance |
| `OOT_MONTHS`, `VALID_SHARE` | 3 / 0.2 | time split |
| `GRADE_QUANTILES` | 0.40, 0.65, 0.85, 0.95 | PD cutpoints G1..G5 |
| `MODEL_BLOCKS` | `auto` | block feeds the model only if its source starts before the train P10 `obs_date` |
| `MODEL_EXCLUDE` | snapshot age, term / remaining months | time artefacts and `contr.end_date`-based leakage risks |
| `PSI_DROP_THRESHOLD` | 0.25 | unstable features are dropped (PSI train -> OOT) |
| `CAPACITY_RATIO_MIN`, `DEPO_COVER_MIN`, `CONTACT_RATE_ENGAGED` | 1.5 / 1.0 / 0.5 | segment axes |
| `R02_PKB_COURT` | True | the PKB "court" status (possibly at another creditor) also stops collection; Legal to decide |
| `CHALLENGER_SHARE` | 0.15 | champion/challenger routing by contract hash |
| `COMPLIANCE` | memo values | notice / restructuring days, call caps, ПМ, MRP, debit limits |

## Data caveats
- **Survivorship.** `contr` must include closed, written-off and sold contracts, or the history under-represents the worst outcomes. `close_date`, `status_name`, `restructing_fl` and `end_date` are extraction-time values: they are never model features.
- **`ovrd` holds only monthly maxima.** Partial payments or cures within a month are invisible, and the episode split is approximate. For the last, partial month, the stock keeps only rows still overdue on the cut (`STOCK_ONLY_AT_CUT`).
- **Treatment contamination.** Outcomes reflect today's SC 4-30 calling. The backtest is *risk under the current treatment*, not uplift; use the challenger arm to measure uplift.
- **Short histories.** `geo` starts 2025-09, `vnpt` 2026-01, `vneshpt` 2026-03 and the PKB events 2026-09-07, so these are overlay/rules only until they have 6+ months of labelled history.
- **`excl`** is a snapshot list, so its history is approximate (`updated_date`). It is rules-only.
- **Demo numbers are not real.** The synthetic data has a planted signal and inflated delinquency rates, so its Gini and bad rates say nothing about the real portfolio. Some of the 77 `attr_*` filler columns in the synthetic `dm` are pure noise and can still appear among the top features.
- The regulatory constants come from a web-search memo (2026). Items marked **[verify]** must be confirmed by Legal on adilet.zan.kz.
