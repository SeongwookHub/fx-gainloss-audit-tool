# -*- coding: utf-8 -*-
"""2단계 증빙 OCR(extract_rate_from_evidence) 동작 확인용 샘플 증빙 이미지 생성기.

evidence/OCR_TEST_001.png 를 만든다. 이 거래ID는 sample_data/분개장.xlsx 에는 존재하지
않으므로, 파이프라인을 정상 실행해도 이 샘플이 결과에 섞여 들어가지 않는다 — 순수하게
manual_ocr_check.py 에서 OCR 추출 정확도만 확인하기 위한 용도.

사용법:
    python tests/generate_ocr_test_sample.py
"""
import os

from PIL import Image, ImageDraw, ImageFont

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMAGE_PATH = os.path.join(REPO_ROOT, "evidence", "OCR_TEST_001.png")

# extract_rate_from_evidence()가 반환해야 할 정답값
GROUND_TRUTH = {
    "거래일자": "2025-09-23",
    "통화": "EUR",
    "적용환율": 1532.40,
    "외화금액": 12500.0,
    "원화금액": 19155000,
}


def _font(size: int, bold: bool = False):
    # 한글(Hangul) 글리프가 없는 Arial 대신, 맑은 고딕 등 한글 지원 폰트를 우선 사용
    candidates = (
        ["malgunbd.ttf", "malgun.ttf"] if bold else ["malgun.ttf"]
    ) + ["arialbd.ttf" if bold else "arial.ttf"]
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def build_image(path: str = IMAGE_PATH) -> str:
    width, height = 900, 640
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    draw.rectangle([4, 4, width - 5, height - 5], outline="black", width=2)

    draw.text((40, 30), "SHINHAN BANK", font=_font(28, bold=True), fill="black")
    draw.text((40, 68), "외환거래확인서 (Foreign Exchange Deal Confirmation)",
              font=_font(15), fill="black")
    draw.line([(40, 100), (width - 40, 100)], fill="black", width=1)

    rows = [
        ("거래번호 (Deal No.)", "FX-2025-092312345"),
        ("거래일자 (Trade Date)", GROUND_TRUTH["거래일자"]),
        ("고객명 (Customer)", "(주)삼도테크 (당사)"),
        ("상대처 (Counterparty)", "Nordic Fjord Seafood AS"),
        ("통화 (Currency)", GROUND_TRUTH["통화"]),
        ("외화금액 (FC Amount)", f"{GROUND_TRUTH['외화금액']:,.2f}"),
        ("적용환율 (Applied Rate)", f"{GROUND_TRUTH['적용환율']:,.2f}"),
        ("원화금액 (KRW Amount)", f"{GROUND_TRUTH['원화금액']:,}"),
        ("입금계좌 (Account)", "외화보통예금 110-yyy-yyyyyy"),
    ]

    y = 130
    for label, value in rows:
        draw.text((40, y), label, font=_font(15), fill="black")
        draw.text((420, y), value, font=_font(15, bold=True), fill="black")
        y += 45

    draw.line([(40, y + 10), (width - 40, y + 10)], fill="black", width=1)
    draw.text((40, y + 30), "본 확인서는 당행 외환거래 시스템에서 발급된 정식 문서입니다.",
              font=_font(13), fill="black")
    draw.text((40, y + 55), "Shinhan Bank Foreign Exchange Center", font=_font(13), fill="gray")

    img.save(path)
    print(f"생성 완료: {path}")
    return path


if __name__ == "__main__":
    build_image()
