"""
Global configuration for Multi-Channel Sensor Calibration Tool (MCAL).
"""
import sys, os

# ── Paths ────────────────────────────────────────────────────────────────────
# PyInstaller onefile 번들 동작:
#   - sys.executable : 실제 exe 파일 경로 (출력 폴더 기준으로 사용)
#   - sys._MEIPASS   : 런타임 임시 압축 해제 폴더 (번들 에셋 기준으로 사용)
# 개발 모드에서는 두 경로 모두 프로젝트 루트를 가리킨다.
if getattr(sys, "frozen", False):
    _ROOT       = os.path.dirname(sys.executable)   # 출력 디렉터리 기준
    _BUNDLE_DIR = sys._MEIPASS                       # 번들 에셋 기준
else:
    _ROOT       = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _BUNDLE_DIR = _ROOT

BASE_DIR      = _ROOT
REFERENCE_DIR = os.path.join(_BUNDLE_DIR, "reference")  # 번들 내 에셋
OUTPUT_DIR    = os.path.join(_ROOT, "outputs")           # 사용자 출력 폴더
OUTPUT_XLSX   = os.path.join(OUTPUT_DIR, "xlsx")
OUTPUT_DOCX   = os.path.join(OUTPUT_DIR, "docx")
OUTPUT_PDF    = os.path.join(OUTPUT_DIR, "pdf")

# ── ADC Specifications ───────────────────────────────────────────────────────
ADC_CENTER        = 32767
ADC_FULL_RANGE    = 65536
ADC_VOLTAGE_RANGE = 20.0   # ±10 V → 20 V total span

# ── Channel Configuration ────────────────────────────────────────────────────
# Channels are assigned by column order in CSV (col 1 = CH01, col 2 = CH02, …)
# Header names in the CSV are intentionally ignored.
CHANNEL_COUNT = 16
CHANNEL_NAMES = [f"CH{i:02d}" for i in range(1, CHANNEL_COUNT + 1)]

# ── Data-Reading Mode ────────────────────────────────────────────────────────
# "auto"   : read every data row in the CSV file (recommended)
# "last_n" : use only the final (SAMPLING_HZ × USE_LAST_SECONDS) rows
DATA_MODE        = "auto"
SAMPLING_HZ      = 100    # used for duration estimation and last_n row count
USE_LAST_SECONDS = 180    # seconds to keep when DATA_MODE == "last_n"

# ── 3-Wire Reference ─────────────────────────────────────────────────────────
THREE_WIRE_EXCEL = os.path.join(REFERENCE_DIR, "3Wire(Reference값).xlsx")
