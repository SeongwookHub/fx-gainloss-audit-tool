# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-pipeline audit-automation tool that verifies 외환차손익 (realized FX gain/loss on
settlement) and 외화환산손익 (unrealized FX translation gain/loss at year-end) from a company's
분개장(journal) / 명세서(외화자산부채명세서, FX asset-liability schedule) / 계정별원장(general
ledger). It reproduces a manual audit procedure (pulling a sample, tracing applied rates against
bank evidence) as an automated two-stage pipeline. Almost the entire implementation lives in one
file: `fx_verification_pipeline.py`.

## Commands

```bash
pip install -r requirements.txt          # runtime deps
pip install -r requirements-dev.txt      # + pytest

# Run the pipeline (defaults to sample_data/, no args needed)
python fx_verification_pipeline.py

# Run against real files with explicit CLI args
python fx_verification_pipeline.py \
  --journal 분개장.xlsx \
  --schedule 명세서_외화자산부채명세서.xlsx \
  --ledger 계정별원장.xlsx \
  --year-end-date 2025-12-31 \
  --output 검증결과.xlsx \
  --ampt 3000000

pytest                                   # full test suite (no network calls; APIs are monkeypatched)
pytest tests/test_fx_verification_pipeline.py::TestCheckAggregateMateriality  # single test class
pytest tests/test_fx_verification_pipeline.py -k reference_date               # by keyword
```

Required env vars for a real (non-mocked) run: `EXIM_AUTH_KEY` (한국수출입은행 Open API — official
매매기준율) and `ANTHROPIC_API_KEY` (Claude Vision, only needed for stage-2 evidence OCR). Optional:
`FX_AMPT` (default 3,000,000, overridden by `--ampt`), `FX_CLAUDE_MODEL` (default
`claude-sonnet-4-6`).

Regenerate sample/demo fixtures: `python generate_samples.py` (writes `sample_data/*.xlsx` with
real historical rates baked into `TestSampleDataOracle`'s `SAMPLE_FAKE_RATES` — if you regenerate
samples with different transactions/dates, update that fixture too). `robustness_test/generate_messy_sample.py`
produces the deliberately messy fixtures used by `TestLoadJournalRobustness`.

## Architecture

### Why two verification strategies, not one

외화환산손익 (unrealized, year-end) has exactly one correct answer — a single official rate on the
balance-sheet date applies to every open position — so it is **fully automated, 100% of the
population**. 외환차손익 (realized, on settlement) has no single "correct" rate — actual bank
transfer terms, preferential rates, etc. vary — so a "correct rate" only exists in the settlement
evidence. That drives a two-stage design: stage 1 screens the full population against the official
rate (cheap, no OCR) and flags outliers; stage 2 sends only the flagged transactions' evidence
images through Claude Vision OCR to extract the actual applied rate. AI is used only where
determinism is impossible (reading unstructured evidence documents) — every rate recalculation,
deviation %, and materiality rollup is plain arithmetic.

### Pipeline stages (in `main()`, reflecting the data flow)

1. **Load & normalize** — `load_journal()` auto-detects the header row (title rows / blank rows
   above the real table are common in real exports), then `extract_settlement_transactions()` /
   `extract_yearend_transactions()` / `get_unsettled_yearend_candidates()` split the journal into
   settled vs. still-open FX positions.
2. **Stage 1 screening** (`screen_fx_settlements`) — backs out an **implied rate** from
   booked KRW ÷ FX amount for each settlement (never trusts a "적용환율" label column directly,
   since it can be wrong even when the booked amounts are internally consistent) and compares it
   to the official rate fetched via `get_official_rate()` → `fetch_rates_for_date()`
   (수출입은행 API, in-memory `_rate_cache` per run). Deviation > `TOLERANCE_PCT` (5%) flags a row.
3. **Two supporting checks that catch what per-row screening misses**:
   - `check_aggregate_materiality()` — sums signed KRW differences across *all* settlements
     (including ones that individually passed) against `AMPT`, since small-% deviations on large
     transactions, or errors netted against other accounts, don't trip a 5% per-row test but do
     add up.
   - `detect_reference_date_mismatch()` — for rows near the tolerance boundary, tests whether the
     implied rate exactly matches a *different* date's official rate (common patterns only: prior
     month-end, prior business day, 1-2 weeks prior — see `_candidate_wrong_dates()`) to catch
     "used the wrong base date" errors that a magnitude-only test wouldn't distinguish from a
     genuine small variance.
4. **Stage 2 evidence OCR** (`verify_with_evidence` → `extract_rate_from_evidence`) — only for
   stage-1-flagged rows. Looks for `evidence/{거래ID}.{png,jpg,pdf}`, sends it to Claude Vision
   (`CLAUDE_MODEL`) to extract the rate actually shown on the bank confirmation, and reconciles
   against the booked rate.
5. **Year-end translation** (`verify_yearend_translation`) — recalculates every open FX position
   at the single official year-end rate; also cross-checks against the FX schedule to catch
   positions missing a revaluation entry entirely (completeness, not just accuracy).
6. **Completeness / reconciliation** — `verify_ledger_reconciliation()` diffs journal-derived FX
   gain/loss account totals against the general ledger (catches manual adjustments booked directly
   to the ledger and missed by the journal-based analysis); `build_rollforward_verification()`
   checks counterparty-level FX exposure rollforward (beginning + activity − ending) as an
   alternative to transaction-ID matching for cases where 1:1 matching individual disbursements to
   a payoff is impractical (e.g., a loan repaid in many installments).
7. **Reporting** (`export_results_to_excel` and helpers below `# ---- Excel export ----`) — writes
   a multi-sheet workbook (안내/요약/A~E, see README "결과 엑셀 시트 구성"), auto-color-coding rows
   via keyword matching (`FLAG_KEYWORDS`/`OK_KEYWORDS` in `_row_status`) and sorting flagged rows
   to the top.

### Robustness pattern used throughout

Every call site that depends on an external API (`get_official_rate`, `fetch_rates_for_date`,
`extract_rate_from_evidence`) is wrapped so a single row's failure (rate lookup miss, missing API
key, OCR error) flags that row (e.g. "오류(환율조회실패)") and lets the batch continue rather than
raising. When adding a new external-dependency call, follow this same per-row try/except pattern
instead of letting an exception propagate out of a batch function.

Functions that can receive an empty input (e.g., a company with zero settlements, only unrealized
exposure) must return a DataFrame with the *same typed columns* the non-empty path produces, not a
bare `pd.DataFrame()` — downstream code merges/concats on those columns and a columnless or
wrong-dtype empty frame breaks the merge. See `SETTLEMENT_COLUMNS`, `YEAREND_RESULT_COLUMNS`, and
the `pd.concat([..., extra], axis=1)` pattern in `screen_fx_settlements` (used specifically because
it preserves dtypes, unlike `pd.DataFrame(columns=[...])`).

### Currency unit handling

한국수출입은행 API quotes some currencies per 100 units (currently only JPY, see
`HUNDRED_UNIT_CURRENCIES`). `resolve_cur_unit()` / `normalize_rate()` handle this; a currency not
present in `CUR_UNIT_MAP` is never silently mismapped — it produces an explicit "통화코드 매핑
미등록" warning on the row instead.

### Account mapping is a manual-in-the-loop step, not automatic

`build_account_mapping()` keyword-classifies a company's chart of accounts against
`FX_ACCOUNT_KEYWORDS`/`STANDARD_ACCOUNT_NAME`, but ambiguous accounts are left as "확인필요" for a
human to resolve; the confirmed mapping is then applied via `apply_account_mapping()`. This exists
because account naming/coding varies per company (by site, country, HQ vs. branch), and is
intentionally not fully automatic.

## Testing conventions

- `tests/test_fx_verification_pipeline.py` never hits the network: `fetch_rates_for_date` is
  monkeypatched via the `_make_fake_fetch(rate_table_by_date)` helper, and Anthropic's client is
  replaced with `_FakeAnthropicClient`/`_FakeMessages`/`_FakeMessage`/`_FakeTextBlock`.
  `tests/conftest.py` only adds the repo root to `sys.path`.
- `TestSampleDataOracle` is the end-to-end regression test: it runs the pipeline against
  `sample_data/` and checks every one of the 8 sample transactions (TXN001-008) against the
  documented verdicts in `sample_data/검증포인트_참고용.xlsx` / README "샘플 데이터" — 3 clean
  cases must produce no false positive, 5 seeded-error cases must all be flagged.
- pandas/numpy booleans from filter expressions are `np.True_`/`np.False_`, not Python `True`/
  `False` — assert on truthiness (`assert x`, `assert not x`), not identity (`is True`).
