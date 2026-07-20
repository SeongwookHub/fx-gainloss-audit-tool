import os
import sys

# fx_verification_pipeline.py는 패키지가 아니라 저장소 루트의 단일 스크립트이므로,
# 어디서 pytest를 실행하든 import가 되도록 루트 경로를 sys.path에 추가한다.
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)
