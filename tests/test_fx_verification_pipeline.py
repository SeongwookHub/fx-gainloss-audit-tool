# -*- coding: utf-8 -*-
"""
fx_verification_pipeline.py 테스트 스위트.

원칙:
- 네트워크 호출(수출입은행 API, Claude API)은 절대 실제로 하지 않는다.
  fetch_rates_for_date/anthropic.Anthropic을 monkeypatch로 대체한다.
- 순수 로직 함수는 모킹 없이 직접 검증한다.
- 새 픽스처를 만들기보다 sample_data/, robustness_test/의 기존 파일을 재사용한다.
"""
import os
from pathlib import Path

import pandas as pd
import pytest

import fx_verification_pipeline as fxp

ROOT_DIR = Path(__file__).resolve().parent.parent
SAMPLE_DIR = ROOT_DIR / "sample_data"
ROBUST_DIR = ROOT_DIR / "robustness_test"


# ------------------------------------------------------------------
# 순수 로직 함수 - 모킹 불필요
# ------------------------------------------------------------------

class TestNormalizeRate:
    def test_jpy_divides_by_100(self):
        assert fxp.normalize_rate(937.73, "JPY") == pytest.approx(9.3773)

    def test_usd_passthrough(self):
        assert fxp.normalize_rate(1430.20, "USD") == 1430.20


class TestNormalizeAccountCode:
    @pytest.mark.parametrize("raw, expected", [
        ("108", "108"),
        (108, "108"),
        (108.0, "108"),
        ("108.0", "108"),
        (" 108 ", "108"),
        ("ABC", "ABC"),  # 숫자로 변환 불가능한 코드는 원본을 그대로 유지
    ])
    def test_various_formats(self, raw, expected):
        assert fxp._normalize_account_code(raw) == expected


class TestResolveCurUnit:
    def test_known_currency(self):
        cur_unit, warning = fxp.resolve_cur_unit("USD")
        assert cur_unit == "USD"
        assert warning is None

    def test_unknown_currency_returns_warning(self):
        cur_unit, warning = fxp.resolve_cur_unit("XYZ")
        assert cur_unit == "XYZ"
        assert "매핑 미등록" in warning


class TestCandidateWrongDates:
    def test_month_boundary(self):
        # 결제일이 1월 3일이면 '전월말'은 전년도 12월 31일이어야 한다.
        candidates = fxp._candidate_wrong_dates(pd.Timestamp("2025-01-03"))
        candidate_strs = {d.strftime("%Y-%m-%d") for d in candidates}
        assert "2024-12-31" in candidate_strs

    def test_all_candidates_before_settle_date(self):
        settle_date = pd.Timestamp("2025-05-15")
        candidates = fxp._candidate_wrong_dates(settle_date)
        assert all(d < settle_date for d in candidates)
        assert len(candidates) == len(set(candidates))  # 중복 없음


class TestCheckAggregateMateriality:
    def _screened(self, diffs):
        return pd.DataFrame({"회사계상액과의차이(KRW)": diffs})

    def test_within_tolerance(self):
        agg = fxp.check_aggregate_materiality(self._screened([1000000, -900000, 999999]), ampt=3000000)
        assert not agg["중요성초과여부"]
        assert agg["확인불가건수"] == 0

    def test_breach_on_gross_sum(self):
        # 순합계는 작아도(방향이 반대라 상쇄) 절대값 합계 기준으로는 AMPT를 넘는 경우
        agg = fxp.check_aggregate_materiality(self._screened([2000000, -2000000, 2500000]), ampt=3000000)
        assert agg["중요성초과여부"]

    def test_none_values_excluded_but_counted(self):
        # 환율조회 실패로 None이 된 행은 합계에서 빠지되, 확인불가건수로 드러나야 한다.
        agg = fxp.check_aggregate_materiality(self._screened([1000000, None, None]), ampt=3000000)
        assert agg["순차이합계(KRW)"] == 1000000
        assert agg["확인불가건수"] == 2


class TestRowStatus:
    def test_flag_keyword(self):
        assert fxp._row_status(["이상치(정밀검증필요)"]) == "flag"

    def test_ok_keyword(self):
        assert fxp._row_status(["적정(스크리닝통과)"]) == "ok"

    def test_neutral(self):
        assert fxp._row_status(["그 외 아무 값"]) == "neutral"

    def test_outlier_count_header_flags_even_without_keyword(self):
        assert fxp._row_status([3], headers=["이상치건수"]) == "flag"
        assert fxp._row_status([0], headers=["이상치건수"]) == "neutral"


class TestAutoAssignVoucherNumber:
    def test_groups_balanced_two_line_vouchers(self):
        df = pd.DataFrame({
            "차변금액": [100000, None, 50000, None],
            "대변금액": [None, 100000, None, 50000],
        })
        result = fxp.auto_assign_voucher_number(df)
        assert list(result["전표번호(거래ID)"]) == ["AUTO00001", "AUTO00001", "AUTO00002", "AUTO00002"]

    def test_raises_when_last_voucher_unbalanced(self):
        df = pd.DataFrame({"차변금액": [100000, None, 50000], "대변금액": [None, 100000, None]})
        with pytest.raises(ValueError):
            fxp.auto_assign_voucher_number(df)


# ------------------------------------------------------------------
# 지저분한 원본 파일 대응 (robustness_test/messy_*.xlsx 재사용)
# ------------------------------------------------------------------

class TestLoadJournalRobustness:
    def test_messy_journal_loads_with_header_detection_and_normalization(self):
        path = ROBUST_DIR / "messy_분개장.xlsx"
        if not path.exists():
            pytest.skip("robustness_test/messy_분개장.xlsx 없음")
        df = fxp.load_journal(str(path))
        # 제목행/빈행이 섞여 있어도 필수 컬럼이 정상적으로 인식되어야 한다.
        for col in fxp.REQUIRED_JOURNAL_COLUMNS:
            assert col in df.columns
        # 계정코드는 전부 문자열로 정규화되어야 한다 ("108.0" 같은 형태가 남아있으면 안 됨).
        assert df["계정코드"].apply(lambda v: not str(v).endswith(".0")).all()


# ------------------------------------------------------------------
# fetch_rates_for_date monkeypatch로 테스트 - 네트워크 호출 없음
# ------------------------------------------------------------------

def _make_fake_fetch(rate_table_by_date: dict):
    def _fake_fetch(date_str, _retries=2):
        return rate_table_by_date.get(date_str, {})
    return _fake_fetch


def _settlement_row(**overrides):
    row = {
        "거래ID": "T1", "결제일": pd.Timestamp("2025-05-15"), "통화": "USD", "구분": "채권",
        "거래처": "테스트거래처", "데이터무결성경고": None,
        "외화금액": 20000, "결제원화금액": 20000 * 1438.50, "발생원화금액": 20000 * 1460.80,
        "회사적용환율(내재)": 1438.50, "회사계상_외환차익차손": 20000 * (1438.50 - 1460.80),
    }
    row.update(overrides)
    return row


class TestScreenFxSettlements:
    def test_flags_deviation_over_tolerance(self, monkeypatch):
        # 공식환율(1000)과 회사환율(1438.5)이 5%를 훨씬 넘게 벌어진 케이스
        monkeypatch.setattr(fxp, "fetch_rates_for_date",
                             _make_fake_fetch({"20250515": {"USD": 1000.0}}))
        settlements = pd.DataFrame([_settlement_row()])
        result = fxp.screen_fx_settlements(settlements)
        assert result.iloc[0]["1차플래그"] == "이상치(정밀검증필요)"

    def test_txn002_style_case_passes_individual_but_shows_diff(self, monkeypatch):
        # 실제 결제일 환율(1415.80) 대신 전월말 환율(1438.50)을 잘못 쓴 경우
        # 괴리율은 5% 미만이라 개별 기준은 통과해야 한다 (합계 중요성에서 걸리는 게 설계 의도).
        monkeypatch.setattr(fxp, "fetch_rates_for_date",
                             _make_fake_fetch({"20250515": {"USD": 1415.80}}))
        settlements = pd.DataFrame([_settlement_row()])
        result = fxp.screen_fx_settlements(settlements)
        assert result.iloc[0]["1차플래그"] == "적정(스크리닝통과)"
        assert result.iloc[0]["회사계상액과의차이(KRW)"] != 0

    def test_degrades_gracefully_when_rate_lookup_fails(self, monkeypatch):
        # API 키 누락/만료 등으로 환율을 아예 못 찾는 상황(5영업일 폴백까지 전부 실패)에도
        # 배치 전체가 죽지 않고 해당 행만 오류로 표시되어야 한다.
        monkeypatch.setattr(fxp, "fetch_rates_for_date", _make_fake_fetch({}))
        settlements = pd.DataFrame([_settlement_row(), _settlement_row(거래ID="T2")])
        result = fxp.screen_fx_settlements(settlements)
        assert len(result) == 2
        assert (result["1차플래그"] == "오류(환율조회실패)").all()
        assert result["공식매매기준율"].isna().all()


class TestDetectReferenceDateMismatch:
    def test_finds_prior_month_end_misuse(self, monkeypatch):
        monkeypatch.setattr(fxp, "fetch_rates_for_date", _make_fake_fetch({
            "20250515": {"USD": 1415.80},  # 실제 결제일 환율
            "20250430": {"USD": 1438.50},  # 회사가 실수로 사용한 전월말 환율
        }))
        settlements = pd.DataFrame([_settlement_row()])
        result = fxp.detect_reference_date_mismatch(settlements)
        row = result.iloc[0]
        assert row["기준일판정"].startswith("기준일 오류 의심")
        assert row["추정사용일자"] == "2025-04-30"

    def test_normal_case_confirms_settle_date_rate(self, monkeypatch):
        monkeypatch.setattr(fxp, "fetch_rates_for_date", _make_fake_fetch({
            "20250515": {"USD": 1438.50},
        }))
        # 기본 _settlement_row()가 이미 회사적용환율(내재)=1438.50이므로 그대로 사용
        settlements = pd.DataFrame([_settlement_row()])
        result = fxp.detect_reference_date_mismatch(settlements)
        assert result.iloc[0]["기준일판정"] == "정상(결제일 환율 사용 확인)"

    def test_degrades_gracefully_when_rate_lookup_fails(self, monkeypatch):
        monkeypatch.setattr(fxp, "fetch_rates_for_date", _make_fake_fetch({}))
        settlements = pd.DataFrame([_settlement_row()])
        result = fxp.detect_reference_date_mismatch(settlements)
        assert result.iloc[0]["기준일판정"].startswith("확인불가")


class TestVerifyWithEvidence:
    def test_no_evidence_file_marks_request_needed(self, tmp_path, monkeypatch):
        monkeypatch.setattr(fxp, "EVIDENCE_DIR", str(tmp_path))  # 빈 폴더
        flagged = pd.DataFrame([{"거래ID": "TXN999", "통화": "USD",
                                  "외화금액": 1000, "회사적용환율(내재)": 1400.0}])
        result = fxp.verify_with_evidence(flagged)
        assert result.iloc[0]["증빙상태"] == "증빙 요청 필요"

    def test_empty_input_returns_correctly_shaped_dataframe(self):
        # 이상치가 하나도 없는 정상적인 기간에도 다운스트림에서 컬럼 접근 시 에러가 나면 안 된다.
        empty = pd.DataFrame(columns=["거래ID", "통화", "외화금액", "회사적용환율(내재)"])
        result = fxp.verify_with_evidence(empty)
        assert result.empty
        for col in ("거래ID", "증빙상태", "증빙확인환율", "최종판정"):
            assert col in result.columns

    def test_ocr_call_failure_degrades_gracefully(self, tmp_path, monkeypatch):
        monkeypatch.setattr(fxp, "EVIDENCE_DIR", str(tmp_path))
        evidence_file = tmp_path / "TXN999.png"
        evidence_file.write_bytes(b"fake-image-bytes")

        def _raise(*args, **kwargs):
            raise RuntimeError("ANTHROPIC_API_KEY 미설정")

        monkeypatch.setattr(fxp, "extract_rate_from_evidence", _raise)
        flagged = pd.DataFrame([{"거래ID": "TXN999", "통화": "USD",
                                  "외화금액": 1000, "회사적용환율(내재)": 1400.0}])
        result = fxp.verify_with_evidence(flagged)
        assert result.iloc[0]["증빙상태"] == "OCR 호출 실패 - 수기 확인 필요"


class _FakeTextBlock:
    type = "text"

    def __init__(self, text):
        self.text = text


class _FakeMessage:
    def __init__(self, text):
        self.content = [_FakeTextBlock(text)]


class _FakeMessages:
    def __init__(self, response_text):
        self._response_text = response_text

    def create(self, **kwargs):
        return _FakeMessage(self._response_text)


class _FakeAnthropicClient:
    def __init__(self, response_text):
        self.messages = _FakeMessages(response_text)


class TestExtractRateFromEvidenceParsing:
    """Claude Vision API 자체는 호출하지 않고, anthropic 클라이언트를 목킹해
    JSON 파싱 성공/실패 분기만 검증한다."""

    def _run_with_fake_response(self, monkeypatch, tmp_path, response_text):
        import anthropic
        img = tmp_path / "evidence.png"
        img.write_bytes(b"fake-image-bytes")
        monkeypatch.setattr(anthropic, "Anthropic",
                             lambda: _FakeAnthropicClient(response_text))
        return fxp.extract_rate_from_evidence(str(img))

    def test_parses_valid_json_response(self, monkeypatch, tmp_path):
        response = '{"거래일자": "2025-05-15", "통화": "USD", "적용환율": 1415.8, "외화금액": 20000, "원화금액": 28316000}'
        result = self._run_with_fake_response(monkeypatch, tmp_path, response)
        assert result["적용환율"] == 1415.8

    def test_unparseable_response_falls_back_to_error(self, monkeypatch, tmp_path):
        result = self._run_with_fake_response(monkeypatch, tmp_path, "이건 JSON이 아닙니다")
        assert "error" in result


# ------------------------------------------------------------------
# End-to-end 오라클 테스트: sample_data/ 8개 거래를 검증포인트_참고용.xlsx와 대조
# ------------------------------------------------------------------

# generate_samples.py에 기록된 실제 매매기준율(당시 수출입은행 API 조회값)을 그대로 재현한
# 오프라인 환율표. 다른 후보일자(전월말/전영업일 등)는 의도적으로 비워둬 candidate 탐색이
# '매칭 없음'으로 안전하게 종료되는지도 같이 검증한다.
SAMPLE_FAKE_RATES = {
    "20250425": {"USD": 1430.20},   # TXN001 결제일
    "20250515": {"USD": 1415.80},   # TXN002 실제 결제일 환율
    "20250430": {"USD": 1438.50},   # TXN002가 잘못 사용한 전월말 환율
    "20250718": {"JPY(100)": 937.73},  # TXN003 결제일
    "20250930": {"USD": 1402.20},   # TXN004 결제일
    "20250611": {"EUR": 1554.64},   # TXN005 결제일
    "20251231": {"USD": 1434.90},   # 2025 결산일 공식환율 (TXN006/TXN008 판정 기준)
}


@pytest.fixture
def sample_pipeline_inputs():
    journal = fxp.load_journal(str(SAMPLE_DIR / "분개장.xlsx"))
    schedule = pd.read_excel(SAMPLE_DIR / "명세서_외화자산부채명세서.xlsx")
    return journal, schedule


class TestSampleDataOracle:
    def test_answer_key_matches_expected_error_normal_split(self):
        """검증포인트_참고용.xlsx 자체가 이 테스트의 전제(어떤 거래가 오류/정상인지)와
        어긋나지 않는지 먼저 확인 - 샘플 데이터가 나중에 바뀌면 이 테스트가 먼저 깨진다."""
        answer_key = pd.read_excel(SAMPLE_DIR / "검증포인트_참고용.xlsx")
        answer_map = dict(zip(answer_key["거래ID"], answer_key["정상/오류"]))
        for txn in ["TXN002", "TXN003", "TXN004", "TXN007", "TXN008"]:
            assert answer_map[txn] == "오류", f"{txn}는 오류 케이스여야 함"
        for txn in ["TXN001", "TXN005", "TXN006"]:
            assert answer_map[txn] == "정상", f"{txn}는 정상 케이스여야 함"

    def test_settlement_screening_matches_expected_flags(self, sample_pipeline_inputs, monkeypatch):
        journal, _schedule = sample_pipeline_inputs
        monkeypatch.setattr(fxp, "fetch_rates_for_date", _make_fake_fetch(SAMPLE_FAKE_RATES))

        settlements = fxp.extract_settlement_transactions(journal)
        screened = fxp.screen_fx_settlements(settlements).set_index("거래ID")

        # TXN001·TXN005: 실제 환율 그대로 정확히 계상 - 개별 스크리닝 통과(false positive 없음)
        assert screened.loc["TXN001", "1차플래그"] == "적정(스크리닝통과)"
        assert screened.loc["TXN005", "1차플래그"] == "적정(스크리닝통과)"
        # TXN002·TXN004: 문제가 있지만 괴리율 자체는 작아 1단계는 통과 (합계중요성에서 걸리는 설계)
        assert screened.loc["TXN002", "1차플래그"] == "적정(스크리닝통과)"
        assert screened.loc["TXN004", "1차플래그"] == "적정(스크리닝통과)"
        # TXN003: JPY 100단위 나누기 누락 - 괴리가 커서 1단계에서 바로 걸려야 함
        assert screened.loc["TXN003", "1차플래그"] == "이상치(정밀검증필요)"

        # 개별로는 안 걸려도, 합계로 보면 중요성을 초과해야 한다(README가 설명하는 설계 의도).
        agg = fxp.check_aggregate_materiality(screened.reset_index(), ampt=3000000)
        assert agg["중요성초과여부"]
        assert agg["확인불가건수"] == 0

    def test_reference_date_mismatch_matches_expected(self, sample_pipeline_inputs, monkeypatch):
        journal, _schedule = sample_pipeline_inputs
        monkeypatch.setattr(fxp, "fetch_rates_for_date", _make_fake_fetch(SAMPLE_FAKE_RATES))

        settlements = fxp.extract_settlement_transactions(journal)
        ref_check = fxp.detect_reference_date_mismatch(settlements).set_index("거래ID")

        assert ref_check.loc["TXN001", "기준일판정"] == "정상(결제일 환율 사용 확인)"
        assert ref_check.loc["TXN002", "기준일판정"].startswith("기준일 오류 의심")
        assert ref_check.loc["TXN002", "추정사용일자"] == "2025-04-30"
        assert ref_check.loc["TXN005", "기준일판정"] == "정상(결제일 환율 사용 확인)"

    def test_yearend_translation_matches_expected(self, sample_pipeline_inputs, monkeypatch):
        journal, schedule = sample_pipeline_inputs
        monkeypatch.setattr(fxp, "fetch_rates_for_date", _make_fake_fetch(SAMPLE_FAKE_RATES))

        ye_txns = fxp.extract_yearend_transactions(journal)
        missing_ye = fxp.get_unsettled_yearend_candidates(journal, schedule)
        ye_result = fxp.verify_yearend_translation(ye_txns, missing_ye, "2025-12-31").set_index("거래ID")

        assert ye_result.loc["TXN006", "일치여부"] == "적정"
        assert ye_result.loc["TXN008", "일치여부"] == "부적정(환율 오류)"
        assert ye_result.loc["TXN007", "일치여부"] == "부적정(기말평가 누락)"

    def test_ledger_reconciliation_catches_seeded_adjustment_gap(self, sample_pipeline_inputs):
        journal, _schedule = sample_pipeline_inputs
        ledger = pd.read_excel(SAMPLE_DIR / "계정별원장.xlsx")
        result = fxp.verify_ledger_reconciliation(journal, ledger).set_index("계정코드")

        # generate_samples.py가 907(외환차익) 계정원장 기말잔액에 50,000원 수기조정분을
        # 의도적으로 심어뒀으므로, 분개장 합계와 어긋나야 한다(완전성 이슈 탐지).
        assert result.loc["907", "일치여부"] == "부적정(분개장-원장 불일치)"
        assert result.loc["957", "일치여부"] == "적정"
        assert result.loc["908", "일치여부"] == "적정"
        assert result.loc["958", "일치여부"] == "적정"


# ------------------------------------------------------------------
# 엑셀 내보내기 - 셀 단위 검증은 하지 않고, 예외 없이 예상 시트가 생성되는지만 확인
# ------------------------------------------------------------------

class TestExportResultsToExcel:
    def test_creates_workbook_with_expected_sheets(self, tmp_path):
        # 이번 기간에 결제 건이 하나도 없는 극단적인 경우(연중 미결제 건만 있는 회사 등)를
        # 실제 함수 체인으로 재현한다 - extract_settlement_transactions부터 시작해야
        # '결제일' 컬럼의 datetime dtype이 실제 파이프라인과 동일하게 유지된다.
        empty_journal = pd.DataFrame(columns=["전표번호(거래ID)", "일자", "구분", "계정코드",
                                               "계정과목", "차변금액", "대변금액", "통화",
                                               "외화금액", "적용환율", "거래처"])
        empty_settlements = fxp.extract_settlement_transactions(empty_journal)
        screened = fxp.screen_fx_settlements(empty_settlements)
        ref_check = fxp.detect_reference_date_mismatch(empty_settlements)
        verified = fxp.verify_with_evidence(screened)
        ye_result = fxp.verify_yearend_translation(
            pd.DataFrame(columns=["거래ID", "통화", "외화금액", "회사적용환율(기말)"]),
            pd.DataFrame(columns=["전표번호(거래ID)", "통화", "기말미결제외화잔액"]),
            "2025-12-31",
        )
        empty = pd.DataFrame()
        agg = {"AMPT": 3000000, "순차이합계(KRW)": 0, "절대값차이합계(KRW)": 0,
               "중요성초과여부": False, "판정": "전체 적정(합계 중요성 이내)", "확인불가건수": 0}
        output_path = tmp_path / "검증결과_테스트.xlsx"

        fxp.export_results_to_excel(
            str(output_path),
            counterparty_summary=empty, rate_trend=empty,
            screened=screened, agg=agg, ref_check=ref_check, verified=verified,
            ye_result=ye_result, ledger_result=empty, rollforward_result=empty,
            full_detail=empty,
        )

        assert output_path.exists()
        import openpyxl
        wb = openpyxl.load_workbook(output_path)
        assert set(wb.sheetnames) == {
            "안내", "요약", "A.표본전체상세", "B.분석적검토",
            "C.외환차손익_상세", "D.외화환산손익", "E.완전성검증",
        }
