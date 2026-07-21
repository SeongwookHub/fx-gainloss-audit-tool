# -*- coding: utf-8 -*-
"""계정매핑(build_account_mapping/apply_account_mapping)과 전표번호 자동채번
(auto_assign_voucher_number) 사용 예시.

이 두 함수는 fx_verification_pipeline.main()/CLI(--workpaper 등)에 자동으로 연결돼
있지 않다 - 회사마다 계정명이 다르거나(계정매핑) 전표번호 컬럼 자체가 없는(자동채번)
원본 파일은, 파이프라인을 돌리기 *전에* 이 함수들로 먼저 표준화해야 한다. 이 스크립트는
그 사전 처리 단계를 실제로 어떻게 호출하는지 보여준다.

사용법: python scripts/demo_account_mapping_and_voucher_numbering.py
"""
import os
import sys

import pandas as pd

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

import fx_verification_pipeline as fxp


def demo_account_mapping():
    print("=" * 70)
    print("1. 계정매핑 (회사마다 다른 계정명 대응)")
    print("=" * 70)

    # 회사의 계정과목 목록(시산표 등) - 표준 명칭과 다른 사내 관행 명칭이 섞여 있다.
    company_accounts = pd.DataFrame([
        {"회사계정코드": "1081", "회사계정명": "외상매출금(달러)"},   # "외상매출금" 포함 -> 자동확정(108)
        {"회사계정코드": "2511", "회사계정명": "외상매입금-USD"},      # "외상매입금" 포함 -> 자동확정(251)
        {"회사계정코드": "9071", "회사계정명": "환차익"},              # "환차익" 포함 -> 자동확정(907)
        {"회사계정코드": "9999", "회사계정명": "달러채권"},            # 표준 키워드 매칭 없음 -> 확인필요
    ])

    mapping_result = fxp.build_account_mapping(company_accounts)
    print(mapping_result.to_string(index=False))

    # 확인필요 건("9999")은 감사인이 직접 표준분류를 확정해야 한다 - 여기서는
    # "달러채권"이 실질적으로 108(외상매출금)이라고 사람이 확인했다고 가정.
    confirmed_mapping = {
        row["회사계정코드"]: row["추천표준분류"]
        for _, row in mapping_result.iterrows() if row["추천표준분류"] is not None
    }
    confirmed_mapping["9999"] = "108"
    print("\n확정된 매핑:", confirmed_mapping)

    # 이 매핑을 실제 분개장(발췌)에 적용
    journal_excerpt = pd.DataFrame([
        {"계정코드": "9999", "계정과목": "달러채권", "차변금액": 14000000, "대변금액": None},
        {"계정코드": "9071", "계정과목": "환차익", "차변금액": None, "대변금액": 120000},
    ])
    mapped = fxp.apply_account_mapping(journal_excerpt, confirmed_mapping)
    print("\n적용 전:")
    print(journal_excerpt.to_string(index=False))
    print("\n적용 후 (계정코드/계정과목이 표준값으로 치환됨, 원본은 원본계정코드/원본계정과목에 보존):")
    print(mapped.to_string(index=False))
    return mapping_result, confirmed_mapping, mapped


def demo_auto_voucher_numbering():
    print("\n" + "=" * 70)
    print("2. 전표번호 자동채번 (전표번호(거래ID) 컬럼 자체가 없는 원본)")
    print("=" * 70)

    # 전표번호 컬럼이 아예 없는 원본 - 차변=대변이 맞아떨어지는 지점을 기준으로
    # 전표 단위를 자동으로 나눈다.
    no_voucher_journal = pd.DataFrame({
        "일자": ["2025-01-15", "2025-01-15", "2025-02-20", "2025-02-20"],
        "계정과목": ["외상매출금(외화)", "매출", "외상매입금(외화)", "원재료매입"],
        "차변금액": [14000000, None, None, 9500000],
        "대변금액": [None, 14000000, 9500000, None],
    })
    result = fxp.auto_assign_voucher_number(no_voucher_journal)
    print("자동채번 결과:")
    print(result.to_string(index=False))
    return result


if __name__ == "__main__":
    demo_account_mapping()
    demo_auto_voucher_numbering()
