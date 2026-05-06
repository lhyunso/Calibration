# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Multi-Channel Sensor Calibration Tool (MCAL)
Build command: pyinstaller MCAL.spec
"""

import sys
from pathlib import Path
import customtkinter
import matplotlib

block_cipher = None

# ── 번들에 포함할 데이터 파일 ──────────────────────────────────────────────
ctk_path   = Path(customtkinter.__file__).parent
mpl_path   = Path(matplotlib.__file__).parent

added_datas = [
    # customtkinter 테마/이미지 에셋
    (str(ctk_path), "customtkinter"),
    # matplotlib 폰트·스타일시트
    (str(mpl_path / "mpl-data"), "matplotlib/mpl-data"),
]

# ── 숨겨진 import (동적 로더가 놓치는 것들) ─────────────────────────────────
hidden_imports = [
    "customtkinter",
    "PIL._tkinter_finder",
    "PIL.Image",
    "openpyxl.cell._writer",
    "openpyxl.styles.stylesheet",
    "reportlab.graphics.barcode.code128",
    "reportlab.graphics.barcode.code39",
    "reportlab.lib.pagesizes",
    "reportlab.platypus",
    "matplotlib.backends.backend_tkagg",
    "matplotlib.backends.backend_agg",
    "pkg_resources.py2_warn",
    # 내부 패키지 (src/ 아래)
    "config",
    "sensors",
    "sensors.pt100",
    "sensors.base",
    "processing",
    "processing.csv_reader",
    "processing.calibration",
    "output",
    "output.xlsx_writer",
    "output.docx_writer",
    "output.pdf_writer",
    "reference",
    "reference.three_wire",
]

a = Analysis(
    ["src/gui.py"],
    pathex=["src"],                  # src/ 안의 모듈을 최상위 패키지로 인식
    binaries=[],
    datas=added_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["_tkinter"],           # tkinter는 시스템 제공
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="MCAL",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # 콘솔 창 숨김 (GUI 전용)
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="assets/icon.ico",  # 아이콘 파일이 있으면 주석 해제
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="MCAL",            # dist/MCAL/ 폴더에 출력
)
