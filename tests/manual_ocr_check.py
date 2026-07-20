# -*- coding: utf-8 -*-
"""2단계 증빙 OCR(extract_rate_from_evidence)이 실제 Claude Vision API로 evidence/OCR_TEST_001.png
를 정확히 읽어내는지 수동으로 확인하는 스크립트.

pytest 스위트(tests/test_fx_verification_pipeline.py)는 Anthropic 클라이언트를 전부
monkeypatch로 대체해 네트워크를 타지 않지만, 이 스크립트는 반대로 실제 API 키로 실제
호출을 해봐야 하므로 이름에 test_/_test 접두/접미를 쓰지 않아 pytest 자동수집 대상이
아니다. CI에도 포함하지 않는다.

사용법:
    export ANTHROPIC_API_KEY="발급받은_키"
    python tests/generate_ocr_test_sample.py   # 샘플 증빙 이미지 생성(최초 1회, 이미 있으면 생략 가능)
    python tests/manual_ocr_check.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))          # tests/ (형제 모듈 import용)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # repo root

from generate_ocr_test_sample import GROUND_TRUTH, IMAGE_PATH  # noqa: E402
from fx_verification_pipeline import extract_rate_from_evidence  # noqa: E402


def _matches(expected, actual) -> bool:
    if actual is None:
        return False
    if isinstance(expected, float):
        try:
            return abs(float(actual) - expected) < 0.01
        except (TypeError, ValueError):
            return False
    return str(actual).strip() == str(expected).strip()


def main() -> int:
    if not os.path.exists(IMAGE_PATH):
        print(f"샘플 증빙 이미지가 없습니다. 먼저 실행하세요: python tests/generate_ocr_test_sample.py")
        return 1

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ANTHROPIC_API_KEY 환경변수가 설정되어 있지 않습니다. 발급받은 키를 먼저 설정하세요.")
        return 1

    print(f"이미지: {IMAGE_PATH}")
    print("Claude Vision으로 증빙 추출 중...\n")
    extracted = extract_rate_from_evidence(IMAGE_PATH)

    print("=== 추출 결과(raw) ===")
    print(extracted)

    if "error" in extracted:
        print("\nOCR 응답을 JSON으로 파싱하지 못했습니다. 위 raw_text를 보고 프롬프트를 조정하세요.")
        return 1

    print("\n=== 정답지 대비 확인 ===")
    all_ok = True
    for field, expected in GROUND_TRUTH.items():
        actual = extracted.get(field)
        ok = _matches(expected, actual)
        all_ok = all_ok and ok
        print(f"  [{'OK' if ok else 'MISMATCH'}] {field}: 기대값={expected} / 추출값={actual}")

    print("\n결과:", "전체 일치 - OCR 파이프라인 정상 동작 확인" if all_ok
          else "일부 불일치 - 위 MISMATCH 항목을 보고 프롬프트/이미지를 조정하세요.")
    return 0 if all_ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
