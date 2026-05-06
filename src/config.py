"""
Global configuration for Multi-Channel Sensor Calibration Tool (MCAL).
"""
import sys, os

# ── Paths ────────────────────────────────────────────────────────────────────
# PyInstaller 번들 안에서는 실행 파일 위치를 기준으로 경로를 설정한다.
if getattr(sys, "frozen", False):
    _ROOT = os.path.dirname(sys.executable)
else:
    _ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BASE_DIR      = _ROOT
REFERENCE_DIR = os.path.join(BASE_DIR, "reference")
OUTPUT_DIR    = os.path.join(BASE_DIR, "outputs")
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
