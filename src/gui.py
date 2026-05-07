#!/usr/bin/env python3
"""
Multi-Channel Sensor Calibration Tool (MCAL) — GUI
====================================================
customtkinter 기반 3-탭 인터페이스:
  탭 1. 캘리브레이션   — CSV 로드 → 계산 → xlsx / docx / pdf 출력
  탭 2. 3-Wire 레퍼런스 — 단일값 조회 / 범위 테이블 / 역산
  탭 3. 설정           — 기본값 저장
"""
import sys, os, json, threading
from datetime import datetime
from typing import Dict, Optional

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import numpy as np

from config import (
    REFERENCE_DIR, OUTPUT_XLSX, OUTPUT_DOCX, OUTPUT_PDF,
    DATA_MODE, SAMPLING_HZ, USE_LAST_SECONDS,
)
from sensors.pt100 import get_sensor
from processing.csv_reader import load_datasets, auto_detect_csv_files
from processing.calibration import calibrate_all_channels, ChannelCalibration
from output.xlsx_writer import CalibrationXlsxWriter
from output.docx_writer import CalibrationDocxWriter
from output.pdf_writer  import CalibrationPdfWriter
from reference.three_wire import (
    build_reference_table, lookup_single, STANDARD_GAINS,
    find_resistance_from_voltage, _SENSOR_RANGES,
)

# ── 상수 ──────────────────────────────────────────────────────────────────────
# PyInstaller 번들(.exe) 안에서는 실행 파일 옆에 settings.json을 저장
def _exe_dir() -> str:
    """실행 파일(또는 스크립트)이 위치한 디렉터리를 반환."""
    if getattr(sys, "frozen", False):          # PyInstaller 번들 실행 중
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

SETTINGS_PATH = os.path.join(_exe_dir(), "settings.json")
NAVY    = "#1F497D"
PASS_C  = "#00B050"
FAIL_C  = "#FF4444"

def _ts():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# ── 설정 저장/로드 ─────────────────────────────────────────────────────────────

def _load_settings() -> dict:
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_settings(data: dict):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[settings] 저장 오류: {e}")


# ── TTK 스타일 ─────────────────────────────────────────────────────────────────

def _apply_ttk_style():
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview",
        background="#2b2b2b", foreground="white",
        fieldbackground="#2b2b2b", rowheight=22,
        font=("맑은 고딕", 9),
    )
    style.configure("Treeview.Heading",
        background=NAVY, foreground="white",
        font=("맑은 고딕", 9, "bold"),
    )
    style.map("Treeview", background=[("selected", NAVY)])


# ═══════════════════════════════════════════════════════════════════════════════
# 탭 1: 캘리브레이션
# ═══════════════════════════════════════════════════════════════════════════════

class CalibrationTab(ctk.CTkFrame):
    def __init__(self, master, settings: dict):
        super().__init__(master)
        self._settings    = settings
        self._csv_rows    = []
        self._calibrations: Optional[Dict[str, ChannelCalibration]] = None
        self._datasets    = None
        self._sensor      = None
        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ── 좌측 설정 패널 ─────────────────────────────────────────────────
        left = ctk.CTkScrollableFrame(self, width=310, label_text="설정 패널")
        left.grid(row=0, column=0, padx=(8, 4), pady=8, sticky="nsew")
        left.grid_columnconfigure(1, weight=1)

        row = 0

        # 센서 타입
        ctk.CTkLabel(left, text="센서 타입", anchor="w").grid(
            row=row, column=0, sticky="w", padx=6, pady=(10, 2))
        self.sensor_var = ctk.StringVar(value="pt100")
        ctk.CTkOptionMenu(left, variable=self.sensor_var,
                          values=["pt100", "pt1000", "strain350"],
                          command=self._on_sensor_change, width=160
                          ).grid(row=row, column=1, sticky="ew", padx=6, pady=(10, 2))
        row += 1

        self._sep(left, row); row += 1

        # 모듈 정보
        self._section_lbl(left, row, "[ 모듈 정보 ]"); row += 1
        s = self._settings
        mod_fields = [
            ("모듈명",   "module_name",   s.get("module_name", "")),
            ("모델명",   "model",         s.get("model", "")),
            ("S/N",      "serial",        ""),
            ("제조사",   "manufacturer",  s.get("manufacturer", "")),
            ("FW 버전",  "fw_version",    ""),
        ]
        self._meta_entries = {}
        for label, key, default in mod_fields:
            row = self._meta_row(left, row, label, key, default)

        self._sep(left, row); row += 1

        # 교정 조건
        self._section_lbl(left, row, "[ 교정 조건 ]"); row += 1
        cond_fields = [
            ("교정 일자", "date",          datetime.now().strftime("%Y.%m.%d")),
            ("장소",      "location",      s.get("location", "")),
            ("담당자",    "operator",      s.get("operator", "")),
            ("온도/습도", "temp_humidity", ""),
        ]
        for label, key, default in cond_fields:
            row = self._meta_row(left, row, label, key, default)

        self._sep(left, row); row += 1

        # 교정 설정
        self._section_lbl(left, row, "[ 교정 설정 ]"); row += 1
        setup_fields = [
            ("케이블",     "cable",         s.get("cable", "전용케이블")),
            ("Inst. Gain", "inst_amp_gain", "1"),
            ("문서번호",   "doc_number",    f"CAL-{datetime.now().strftime('%Y')}-0001"),
            ("Rev",        "revision",      "00"),
        ]
        for label, key, default in setup_fields:
            row = self._meta_row(left, row, label, key, default)

        # 데이터 모드
        ctk.CTkLabel(left, text="데이터 모드", anchor="w").grid(
            row=row, column=0, sticky="w", padx=6, pady=2)
        self.data_mode_var = ctk.StringVar(value="auto")
        ctk.CTkOptionMenu(left, variable=self.data_mode_var,
                          values=["auto", "last_n"],
                          command=self._on_mode_change, width=160
                          ).grid(row=row, column=1, sticky="ew", padx=6, pady=2)
        row += 1

        self.lastn_frame = ctk.CTkFrame(left, fg_color="transparent")
        self.lastn_frame.grid(row=row, column=0, columnspan=2, sticky="ew"); row += 1
        ctk.CTkLabel(self.lastn_frame, text="Hz", anchor="w", width=60).pack(side="left", padx=(6,2))
        self.hz_entry = ctk.CTkEntry(self.lastn_frame, width=60)
        self.hz_entry.insert(0, str(SAMPLING_HZ))
        self.hz_entry.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(self.lastn_frame, text="초", anchor="w", width=30).pack(side="left", padx=(0,2))
        self.sec_entry = ctk.CTkEntry(self.lastn_frame, width=60)
        self.sec_entry.insert(0, str(USE_LAST_SECONDS))
        self.sec_entry.pack(side="left")
        self.lastn_frame.grid_remove()

        self._sep(left, row); row += 1

        # CSV 파일
        self._section_lbl(left, row, "[ CSV 파일 ]"); row += 1
        self._csv_frame = ctk.CTkFrame(left)
        self._csv_frame.grid(row=row, column=0, columnspan=2, sticky="ew", padx=4)
        self._csv_frame.grid_columnconfigure(1, weight=1)
        row += 1

        btn_f = ctk.CTkFrame(left, fg_color="transparent")
        btn_f.grid(row=row, column=0, columnspan=2, sticky="ew", padx=4, pady=4)
        ctk.CTkButton(btn_f, text="자동 감지", width=100, command=self._auto_detect).pack(side="left", padx=3)
        ctk.CTkButton(btn_f, text="행 추가",   width=80,  command=lambda: self._add_csv_row()).pack(side="left", padx=3)
        row += 1

        self._sep(left, row); row += 1

        # 실행 버튼
        self.run_btn = ctk.CTkButton(
            left, text="▶  캘리브레이션 실행",
            fg_color=NAVY, hover_color="#2d6aad",
            height=42, font=ctk.CTkFont(size=13, weight="bold"),
            command=self._run,
        )
        self.run_btn.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=4); row += 1

        self.progress = ctk.CTkProgressBar(left)
        self.progress.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=2); row += 1
        self.progress.set(0)

        self.status_lbl = ctk.CTkLabel(left, text="대기 중", text_color="gray", wraplength=280)
        self.status_lbl.grid(row=row, column=0, columnspan=2, padx=6, pady=2); row += 1

        self._sep(left, row); row += 1

        # 내보내기
        self._section_lbl(left, row, "[ 내보내기 ]"); row += 1
        out_dir_f = ctk.CTkFrame(left, fg_color="transparent")
        out_dir_f.grid(row=row, column=0, columnspan=2, sticky="ew", padx=4, pady=2)
        ctk.CTkLabel(out_dir_f, text="출력폴더", width=60, anchor="w").pack(side="left", padx=4)
        self.out_dir_var = ctk.StringVar(value=s.get("output_dir", os.path.join(os.path.dirname(__file__), "..", "outputs")))
        ctk.CTkEntry(out_dir_f, textvariable=self.out_dir_var, width=130).pack(side="left")
        ctk.CTkButton(out_dir_f, text="…", width=28,
                      command=self._browse_outdir).pack(side="left", padx=2)
        row += 1

        exp_f = ctk.CTkFrame(left)
        exp_f.grid(row=row, column=0, columnspan=2, sticky="ew", padx=4, pady=4)
        for text, cmd in [("xlsx", self._export_xlsx), ("docx", self._export_docx), ("pdf", self._export_pdf)]:
            btn = ctk.CTkButton(exp_f, text=text, width=68, state="disabled", command=cmd)
            btn.pack(side="left", padx=3, pady=4)
            setattr(self, f"btn_{text}", btn)

        # ── 우측: 결과 탭뷰 ────────────────────────────────────────────────
        right = ctk.CTkTabview(self)
        right.grid(row=0, column=1, padx=(4, 8), pady=8, sticky="nsew")
        right.add("요약 2-1")
        right.add("요약 2-2")
        right.add("채널 상세")
        right.add("차트")

        self._build_summary_tab(right.tab("요약 2-1"), "21")
        self._build_summary_tab(right.tab("요약 2-2"), "22")
        self._build_detail_tab(right.tab("채널 상세"))
        self._build_chart_tab(right.tab("차트"))

        self._on_sensor_change("pt100")

    # ── 헬퍼 위젯 ──────────────────────────────────────────────────────────

    def _sep(self, parent, row):
        ctk.CTkLabel(parent, text="─" * 34, text_color="gray40").grid(
            row=row, column=0, columnspan=2, pady=3)

    def _section_lbl(self, parent, row, text):
        ctk.CTkLabel(parent, text=text, anchor="w",
                     font=ctk.CTkFont(weight="bold"),
                     text_color="#a0c4e8").grid(
            row=row, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 0))

    def _meta_row(self, parent, row, label, key, default):
        ctk.CTkLabel(parent, text=label, anchor="w", width=72).grid(
            row=row, column=0, sticky="w", padx=6, pady=2)
        e = ctk.CTkEntry(parent, width=160)
        e.grid(row=row, column=1, sticky="ew", padx=6, pady=2)
        if default:
            e.insert(0, default)
        self._meta_entries[key] = e
        return row + 1

    # ── 데이터 모드 ─────────────────────────────────────────────────────────

    def _on_mode_change(self, mode: str):
        if mode == "last_n":
            self.lastn_frame.grid()
        else:
            self.lastn_frame.grid_remove()

    # ── CSV 관리 ────────────────────────────────────────────────────────────

    def _on_sensor_change(self, sensor_type: str):
        self._sensor = get_sensor(sensor_type)
        self._clear_csv_rows()
        for r in self._sensor.default_resistances:
            self._add_csv_row(resistance=r)

    def _clear_csv_rows(self):
        for widgets in self._csv_rows:
            for w in widgets[:3]:
                w.destroy()
        self._csv_rows.clear()

    def _add_csv_row(self, resistance: float = None):
        f   = self._csv_frame
        idx = len(self._csv_rows)
        r_var    = ctk.StringVar(value=str(int(resistance)) if resistance else "")
        path_var = ctk.StringVar()

        r_ent = ctk.CTkEntry(f, textvariable=r_var, width=50, placeholder_text="Ω")
        r_ent.grid(row=idx, column=0, padx=(4, 2), pady=2)
        p_ent = ctk.CTkEntry(f, textvariable=path_var, width=158, placeholder_text="파일 경로")
        p_ent.grid(row=idx, column=1, padx=(0, 2), pady=2, sticky="ew")
        btn = ctk.CTkButton(f, text="…", width=28, command=lambda pv=path_var: self._browse(pv))
        btn.grid(row=idx, column=2, padx=(0, 4), pady=2)
        self._csv_rows.append((r_ent, p_ent, btn, r_var, path_var))

    def _browse(self, path_var):
        p = filedialog.askopenfilename(
            initialdir=REFERENCE_DIR,
            filetypes=[("CSV", "*.csv"), ("모든 파일", "*.*")],
        )
        if p:
            path_var.set(p)

    def _browse_outdir(self):
        d = filedialog.askdirectory(initialdir=self.out_dir_var.get())
        if d:
            self.out_dir_var.set(d)

    def _auto_detect(self):
        files = auto_detect_csv_files(REFERENCE_DIR)
        if not files:
            messagebox.showwarning("자동 감지", f"CSV 파일을 찾을 수 없습니다.\n{REFERENCE_DIR}")
            return
        self._clear_csv_rows()
        for r in sorted(files.keys()):
            self._add_csv_row(resistance=r)
            self._csv_rows[-1][4].set(files[r])
        self._set_status(f"{len(files)}개 파일 감지", "gray")

    def _collect_csv(self) -> Dict[float, str]:
        result = {}
        for (re, pe, btn, rv, pv) in self._csv_rows:
            try:
                r = float(rv.get().strip())
            except ValueError:
                continue
            p = pv.get().strip()
            if p and os.path.exists(p):
                result[r] = p
        return result

    def _collect_meta(self) -> dict:
        return {k: e.get() for k, e in self._meta_entries.items()}

    # ── 요약 탭 ────────────────────────────────────────────────────────────

    def _build_summary_tab(self, frame, attr_suffix: str):
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        cols = ("채널", "Gain", "Offset", "판정")
        tv = ttk.Treeview(frame, columns=cols, show="headings", height=18)
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=80, anchor="center")
        sb = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        tv.tag_configure("pass", foreground=PASS_C)
        tv.tag_configure("fail", foreground=FAIL_C)
        setattr(self, f"tv_{attr_suffix}", tv)

    def _populate_summary(self):
        cals   = self._calibrations
        sensor = self._sensor
        rs     = sorted(next(iter(cals.values())).voltages_avg.keys())

        for attr, dev_k, off_k in [
            ("tv_21", "dev_final_100",  "offset_100"),
            ("tv_22", "dev_final_mean", "offset_mean"),
        ]:
            tv = getattr(self, attr)
            tv.delete(*tv.get_children())
            new_cols = ("채널", "Gain", "Offset") + tuple(f"{int(r)}Ω" for r in rs) + ("판정",)
            tv["columns"] = new_cols
            for c in new_cols:
                tv.heading(c, text=c)
                tv.column(c, width=80 if c not in ("채널",) else 70, anchor="center")

            for ch, cal in sorted(cals.items()):
                off  = getattr(cal, off_k)
                devs = getattr(cal, dev_k)
                ok   = all(abs(devs.get(r, 0)) <= sensor.tolerance_ohm for r in rs)
                tv.insert("", "end", tags=("pass" if ok else "fail",), values=(
                    ch, f"{cal.gain:.6f}", f"{off:.6f}",
                    *[f"{devs.get(r, 0):.4f}" for r in rs],
                    "PASS" if ok else "FAIL",
                ))

    # ── 채널 상세 탭 ───────────────────────────────────────────────────────

    def _build_detail_tab(self, frame):
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(frame)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ctk.CTkLabel(top, text="채널 선택:").pack(side="left", padx=8)
        self.detail_ch_var = ctk.StringVar()
        self.detail_menu   = ctk.CTkOptionMenu(
            top, variable=self.detail_ch_var,
            values=["(계산 후 표시)"],
            command=self._show_detail, width=160,
        )
        self.detail_menu.pack(side="left", padx=4)

        cols = ("항목", "구분", "값")
        self.detail_tv = ttk.Treeview(frame, columns=cols, show="headings", height=22)
        for c in cols:
            self.detail_tv.heading(c, text=c)
            self.detail_tv.column(c, width=180 if c in ("항목","구분") else 300, anchor="center")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.detail_tv.yview)
        self.detail_tv.configure(yscrollcommand=sb.set)
        self.detail_tv.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

    def _show_detail(self, ch: str):
        if not self._calibrations or ch not in self._calibrations:
            return
        cal = self._calibrations[ch]
        rs  = sorted(cal.voltages_avg.keys())

        new_cols = ("항목", "구분") + tuple(f"{int(r)}Ω" for r in rs)
        self.detail_tv["columns"] = new_cols
        for c in new_cols:
            self.detail_tv.heading(c, text=c)
            self.detail_tv.column(c, width=150 if c in ("항목","구분") else 90, anchor="center")
        self.detail_tv.delete(*self.detail_tv.get_children())

        def _r(sec, sub, vals):
            fmtd = []
            for v in vals:
                if v is None:    fmtd.append("-")
                elif isinstance(v, float): fmtd.append(f"{v:.4f}")
                else:            fmtd.append(str(int(v)))
            self.detail_tv.insert("", "end", values=(sec, sub, *fmtd))

        _r("Decimal",            "AVG",    [cal.decimals_avg.get(r) for r in rs])
        _r("",                   "MIN",    [cal.decimals_min.get(r) for r in rs])
        _r("",                   "MAX",    [cal.decimals_max.get(r) for r in rs])
        _r("전압 (before gain)", "Ref",    [cal.voltage_ref.get(r) for r in rs])
        _r("",                   "AVG",    [cal.voltages_avg.get(r) for r in rs])
        _r("저항 (before gain)", "AVG",    [cal.r_before_gain.get(r) for r in rs])
        _r("",                   "편차",   [cal.dev_before_gain.get(r) for r in rs])
        self.detail_tv.insert("", "end", values=("Gain", f"{cal.gain:.8f}", *[""] * len(rs)))
        _r("저항 (after gain)",  "AVG",    [cal.r_after_gain_avg.get(r) for r in rs])
        _r("",                   "MIN",    [cal.r_after_gain_min.get(r) for r in rs])
        _r("",                   "MAX",    [cal.r_after_gain_max.get(r) for r in rs])
        _r("",                   "편차",   [cal.dev_after_gain.get(r) for r in rs])
        self.detail_tv.insert("", "end",
            values=("Offset",
                    f"2-1: {cal.offset_100:.6f}  /  2-2: {cal.offset_mean:.6f}",
                    *[""] * len(rs)))
        _r("2-1 최종 (R_nom)",   "편차",   [cal.dev_final_100.get(r) for r in rs])
        _r("",                   "허용(+)",[cal.tolerance_max_100.get(r) for r in rs])
        _r("",                   "허용(-)",[cal.tolerance_min_100.get(r) for r in rs])
        _r("2-2 최종 (Mean)",    "편차",   [cal.dev_final_mean.get(r) for r in rs])
        _r("",                   "허용(+)",[cal.tolerance_max_mean.get(r) for r in rs])
        _r("",                   "허용(-)",[cal.tolerance_min_mean.get(r) for r in rs])

    # ── 차트 탭 ────────────────────────────────────────────────────────────

    def _build_chart_tab(self, frame):
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(frame)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ctk.CTkLabel(top, text="채널:").pack(side="left", padx=8)
        self.chart_ch_var = ctk.StringVar()
        self.chart_menu   = ctk.CTkOptionMenu(
            top, variable=self.chart_ch_var,
            values=["(계산 후 표시)"],
            command=self._draw_chart, width=160,
        )
        self.chart_menu.pack(side="left", padx=4)

        chart_f = ctk.CTkFrame(frame)
        chart_f.grid(row=1, column=0, sticky="nsew")
        chart_f.grid_columnconfigure(0, weight=1)
        chart_f.grid_rowconfigure(0, weight=1)

        self.fig, self.axes = plt.subplots(1, 2, figsize=(11, 4.5))
        self.fig.patch.set_facecolor("#1a1a2e")
        for ax in self.axes:
            ax.set_facecolor("#16213e")
            ax.tick_params(colors="white", labelsize=7)
            for sp in ax.spines.values():
                sp.set_edgecolor("#444")

        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_f)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky="nsew")

    def _draw_chart(self, ch: str):
        if not self._calibrations or ch not in self._calibrations:
            return
        cal = self._calibrations[ch]
        rs  = sorted(cal.voltages_avg.keys())
        tol = self._sensor.tolerance_ohm if self._sensor else 0.385

        ax0, ax1 = self.axes
        ax0.cla(); ax1.cla()
        for ax in self.axes:
            ax.set_facecolor("#16213e")
            ax.tick_params(colors="white", labelsize=7)
            for sp in ax.spines.values(): sp.set_edgecolor("#444")

        vr = [cal.voltage_ref[r]  for r in rs]
        vm = [cal.voltages_avg[r] for r in rs]
        ax0.scatter(vr, vm, color="#4fc3f7", s=60, zorder=5)
        if len(vr) >= 2:
            c  = np.polyfit(vr, vm, 1)
            xl = np.linspace(min(vr)*1.08, max(vr)*1.08, 100)
            ax0.plot(xl, np.polyval(c, xl), color="#ef5350", lw=1.5,
                     label=f"Gain = {cal.gain:.6f}")
        ax0.set_xlabel("Ref Voltage (V)", color="white", fontsize=8)
        ax0.set_ylabel("Meas Voltage (V)", color="white", fontsize=8)
        ax0.set_title(f"Gain  —  {ch}", color="white", fontsize=9, pad=6)
        ax0.legend(fontsize=7, labelcolor="white", facecolor="#1a1a2e", edgecolor="#555")
        ax0.grid(True, ls="--", alpha=0.25, color="#666")

        x  = np.arange(len(rs)); w = 0.32
        d21 = [cal.dev_final_100.get(r,  0) for r in rs]
        d22 = [cal.dev_final_mean.get(r, 0) for r in rs]
        ax1.bar(x - w/2, d21, w, label="2-1 (R_nom)", color="#1565c0", alpha=0.85)
        ax1.bar(x + w/2, d22, w, label="2-2 (Mean)",  color="#ef6c00", alpha=0.85)
        ax1.axhline(y= tol, color="#ef5350", ls="--", lw=1.2, label=f"+Tol ({tol})")
        ax1.axhline(y=-tol, color="#42a5f5", ls="--", lw=1.2, label=f"−Tol ({-tol})")
        ax1.axhline(y=0,    color="#888",    ls="-",  lw=0.5)
        ax1.set_xticks(x)
        ax1.set_xticklabels([f"{int(r)}Ω" for r in rs], color="white", fontsize=7)
        ax1.set_ylabel("Deviation (Ω)", color="white", fontsize=8)
        ax1.set_title(f"Deviation  —  {ch}", color="white", fontsize=9, pad=6)
        ax1.legend(fontsize=6, labelcolor="white", facecolor="#1a1a2e", edgecolor="#555")
        ax1.grid(True, axis="y", ls="--", alpha=0.25, color="#666")

        self.fig.tight_layout(pad=1.2)
        self.canvas.draw()

    # ── 실행 ───────────────────────────────────────────────────────────────

    def _set_status(self, msg: str, color: str = "gray"):
        self.status_lbl.configure(text=msg, text_color=color)

    def _run(self):
        csv_paths = self._collect_csv()
        if not csv_paths:
            messagebox.showwarning("오류", "CSV 파일을 먼저 선택하거나 자동 감지하세요.")
            return
        self.run_btn.configure(state="disabled")
        self.progress.set(0.05)
        self._set_status("CSV 로드 중…", "orange")

        mode = self.data_mode_var.get()
        try:
            hz  = int(self.hz_entry.get())
            sec = int(self.sec_entry.get())
        except ValueError:
            hz, sec = SAMPLING_HZ, USE_LAST_SECONDS

        def _worker():
            try:
                ds = load_datasets(csv_paths, data_mode=mode, sampling_hz=hz, use_last_seconds=sec)
                self.after(0, lambda: self.progress.set(0.4))
                self.after(0, lambda: self._set_status("계산 중…", "orange"))
                sensor = self._sensor or get_sensor("pt100")
                meta   = self._collect_meta()
                gain   = float(meta.get("inst_amp_gain", "1") or "1")
                cals   = calibrate_all_channels(
                    ds,
                    r_nominal=sensor.r_nominal,
                    excitation=sensor.excitation,
                    inst_amp_gain=gain,
                    tolerance=sensor.tolerance_ohm,
                    sensor=sensor,
                )
                self._calibrations = cals
                self._datasets     = ds
                self.after(0, lambda: self._on_done(cals))
            except Exception as e:
                import traceback; traceback.print_exc()
                self.after(0, lambda: self._on_error(str(e)))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_done(self, cals):
        self.progress.set(1.0)
        self._set_status(f"완료 — {len(cals)}채널", PASS_C)
        self.run_btn.configure(state="normal")
        for name in ("xlsx", "docx", "pdf"):
            getattr(self, f"btn_{name}").configure(state="normal")
        self._populate_summary()
        chs = sorted(cals.keys())
        self.detail_menu.configure(values=chs)
        self.detail_ch_var.set(chs[0]); self._show_detail(chs[0])
        self.chart_menu.configure(values=chs)
        self.chart_ch_var.set(chs[0]); self._draw_chart(chs[0])

    def _on_error(self, msg: str):
        self.progress.set(0)
        self._set_status(f"오류: {msg}", FAIL_C)
        self.run_btn.configure(state="normal")
        messagebox.showerror("오류", msg)

    # ── 내보내기 ────────────────────────────────────────────────────────────

    def _build_meta(self) -> dict:
        meta = self._collect_meta()
        if self._datasets:
            durs = [ds.duration_seconds for ds in self._datasets.values() if ds.duration_seconds > 0]
            if durs:
                meta["duration_sec"]  = f"{max(durs):.1f}"
                meta["sampling_hz"]   = self.hz_entry.get()
        return meta

    def _out_path(self, ext: str) -> Optional[str]:
        sensor   = self._sensor or get_sensor("pt100")
        base_dir = self.out_dir_var.get()
        sub      = {"xlsx": "xlsx", "docx": "docx", "pdf": "pdf"}.get(ext, ext)
        out_dir  = os.path.join(base_dir, sub)
        os.makedirs(out_dir, exist_ok=True)
        return filedialog.asksaveasfilename(
            initialdir=out_dir,
            defaultextension=f".{ext}",
            initialfile=f"{sensor.sensor_type}_cal_{_ts()}.{ext}",
            filetypes=[(ext.upper(), f"*.{ext}")],
        )

    def _export_xlsx(self):
        if not self._calibrations: return
        path = self._out_path("xlsx")
        if not path: return
        sensor = self._sensor or get_sensor("pt100")
        meta   = self._build_meta()
        self._set_status("xlsx 생성 중…", "orange")
        def _w():
            try:
                CalibrationXlsxWriter(sensor, self._calibrations, meta).write(path)
                self.after(0, lambda: self._set_status("xlsx 저장 완료", PASS_C))
                self.after(0, lambda: messagebox.showinfo("저장 완료", path))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("xlsx 오류", str(e)))
        threading.Thread(target=_w, daemon=True).start()

    def _export_docx(self):
        if not self._calibrations: return
        path = self._out_path("docx")
        if not path: return
        sensor = self._sensor or get_sensor("pt100")
        meta   = self._build_meta()
        self._set_status("docx 생성 중…", "orange")
        def _w():
            try:
                CalibrationDocxWriter(sensor, self._calibrations, meta).write(path)
                self.after(0, lambda: self._set_status("docx 저장 완료", PASS_C))
                self.after(0, lambda: messagebox.showinfo("저장 완료", path))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("docx 오류", str(e)))
        threading.Thread(target=_w, daemon=True).start()

    def _export_pdf(self):
        if not self._calibrations: return
        path = self._out_path("pdf")
        if not path: return
        sensor = self._sensor or get_sensor("pt100")
        meta   = self._build_meta()
        self._set_status("pdf 생성 중…", "orange")
        def _w():
            try:
                CalibrationPdfWriter(sensor, self._calibrations, meta).write(path)
                self.after(0, lambda: self._set_status("pdf 저장 완료", PASS_C))
                self.after(0, lambda: messagebox.showinfo("저장 완료", path))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("pdf 오류", str(e)))
        threading.Thread(target=_w, daemon=True).start()


# ═══════════════════════════════════════════════════════════════════════════════
# 탭 2: 3-Wire 레퍼런스
# ═══════════════════════════════════════════════════════════════════════════════

class ReferenceTab(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        self._ref_data = None
        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        sub = ctk.CTkTabview(self)
        sub.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        sub.add("단일값 조회")
        sub.add("범위 테이블")
        sub.add("역산  V → R")
        self._build_single(sub.tab("단일값 조회"))
        self._build_table(sub.tab("범위 테이블"))
        self._build_reverse(sub.tab("역산  V → R"))

    def _build_single(self, frame):
        p = {"padx": 20, "pady": 10, "sticky": "w"}
        row = 0
        ctk.CTkLabel(frame, text="센서 타입").grid(row=row, column=0, **p)
        self.s_sensor = ctk.StringVar(value="pt100")
        ctk.CTkOptionMenu(frame, variable=self.s_sensor,
                          values=["pt100","pt1000","strain350"], width=160
                          ).grid(row=row, column=1, **p); row += 1
        ctk.CTkLabel(frame, text="저항  R  (Ω)").grid(row=row, column=0, **p)
        self.s_r = ctk.CTkEntry(frame, width=130, placeholder_text="예) 110")
        self.s_r.grid(row=row, column=1, **p); row += 1
        ctk.CTkLabel(frame, text="Gain").grid(row=row, column=0, **p)
        self.s_gain = ctk.CTkEntry(frame, width=130)
        self.s_gain.insert(0, "1")
        self.s_gain.grid(row=row, column=1, **p); row += 1
        ctk.CTkButton(frame, text="조회", width=130,
                      command=self._lookup_single).grid(
            row=row, column=0, columnspan=2, padx=20, pady=14, sticky="w"); row += 1
        self.s_result = ctk.CTkLabel(
            frame, text="V_ref  =  —",
            font=ctk.CTkFont(size=20, weight="bold"), text_color="#4fc3f7")
        self.s_result.grid(row=row, column=0, columnspan=2, padx=20, pady=6, sticky="w"); row += 1
        self.s_note = ctk.CTkLabel(frame, text="", text_color="gray", wraplength=420)
        self.s_note.grid(row=row, column=0, columnspan=2, padx=20, pady=4, sticky="w")

    def _lookup_single(self):
        try:
            r    = float(self.s_r.get())
            gain = float(self.s_gain.get())
        except ValueError:
            messagebox.showwarning("입력 오류", "숫자를 입력하세요."); return
        stype  = self.s_sensor.get()
        sensor = get_sensor(stype)
        v      = lookup_single(stype, r, gain)
        self.s_result.configure(text=f"V_ref  =  {v:.6f}  V")
        self.s_note.configure(
            text=f"공식: {sensor.ref_formula}  |  R={r}Ω,  ΔR={r-sensor.r_nominal:+.1f}Ω,  Gain={gain}")

    def _build_table(self, frame):
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)
        ctrl = ctk.CTkFrame(frame)
        ctrl.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        self.t_sensor = ctk.StringVar(value="pt100")
        ctk.CTkLabel(ctrl, text="센서:").pack(side="left", padx=(8,2))
        ctk.CTkOptionMenu(ctrl, variable=self.t_sensor,
                          values=["pt100","pt1000","strain350"],
                          width=110).pack(side="left", padx=(0,10))
        def _ent(lbl, default, w=70):
            ctk.CTkLabel(ctrl, text=lbl).pack(side="left", padx=(6,2))
            e = ctk.CTkEntry(ctrl, width=w); e.insert(0, default); e.pack(side="left", padx=(0,6))
            return e
        self.t_start = _ent("R 시작", "80")
        self.t_end   = _ent("R 끝",   "120")
        self.t_step  = _ent("간격",   "1", 50)
        self.t_gains = _ent("Gain", "1,10,100", 110)
        ctk.CTkButton(ctrl, text="생성",     width=70,  command=self._gen_table).pack(side="left", padx=4)
        ctk.CTkButton(ctrl, text="CSV 저장", width=80,  command=self._export_csv).pack(side="left", padx=2)
        self.t_container = ctk.CTkFrame(frame)
        self.t_container.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        self.t_container.grid_columnconfigure(0, weight=1)
        self.t_container.grid_rowconfigure(0, weight=1)

    def _gen_table(self):
        try:
            r0, r1 = float(self.t_start.get()), float(self.t_end.get())
            step   = float(self.t_step.get())
            gains  = [float(g.strip()) for g in self.t_gains.get().split(",")]
        except ValueError:
            messagebox.showwarning("입력 오류", "숫자를 올바르게 입력하세요."); return
        stype  = self.t_sensor.get()
        sensor = get_sensor(stype)
        rs, gs, tbl = build_reference_table(stype, r0, r1, step, gains)
        self._ref_data = (stype, sensor, rs, gs, tbl)
        for w in self.t_container.winfo_children(): w.destroy()
        cols = ("R (Ω)", "ΔR (Ω)") + tuple(f"Gain={int(g)}" for g in gs)
        tv = ttk.Treeview(self.t_container, columns=cols, show="headings")
        for c in cols:
            tv.heading(c, text=c); tv.column(c, width=90, anchor="center")
        sb = ttk.Scrollbar(self.t_container, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.grid(row=0, column=0, sticky="nsew"); sb.grid(row=0, column=1, sticky="ns")
        for r in rs:
            tv.insert("", "end", values=(f"{r:.1f}", f"{r-sensor.r_nominal:+.1f}",
                                          *[f"{v:.5f}" for v in tbl[r]]))

    def _export_csv(self):
        if not self._ref_data:
            messagebox.showwarning("없음", "먼저 테이블을 생성하세요."); return
        import csv
        stype, sensor, rs, gs, tbl = self._ref_data
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=f"3wire_ref_{stype}_{_ts()}.csv",
            filetypes=[("CSV", "*.csv")],
        )
        if not path: return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["R (Ω)", "ΔR (Ω)"] + [f"Gain={int(g)}" for g in gs])
            for r in rs:
                w.writerow([f"{r:.1f}", f"{r-sensor.r_nominal:+.1f}"] +
                            [f"{v:.6f}" for v in tbl[r]])
        messagebox.showinfo("저장 완료", path)

    def _build_reverse(self, frame):
        p = {"padx": 20, "pady": 10, "sticky": "w"}
        row = 0
        ctk.CTkLabel(frame, text="센서 타입").grid(row=row, column=0, **p)
        self.rv_sensor = ctk.StringVar(value="pt100")
        ctk.CTkOptionMenu(frame, variable=self.rv_sensor,
                          values=["pt100","pt1000","strain350"], width=160
                          ).grid(row=row, column=1, **p); row += 1
        ctk.CTkLabel(frame, text="전압  V  (V)").grid(row=row, column=0, **p)
        self.rv_v = ctk.CTkEntry(frame, width=130, placeholder_text="예) 0.005")
        self.rv_v.grid(row=row, column=1, **p); row += 1
        ctk.CTkLabel(frame, text="Gain").grid(row=row, column=0, **p)
        self.rv_gain = ctk.CTkEntry(frame, width=130); self.rv_gain.insert(0, "1")
        self.rv_gain.grid(row=row, column=1, **p); row += 1
        ctk.CTkButton(frame, text="역산", width=130,
                      command=self._reverse).grid(
            row=row, column=0, columnspan=2, padx=20, pady=14, sticky="w"); row += 1
        self.rv_result = ctk.CTkLabel(
            frame, text="R  ≈  —",
            font=ctk.CTkFont(size=20, weight="bold"), text_color="#4fc3f7")
        self.rv_result.grid(row=row, column=0, columnspan=2, padx=20, pady=6, sticky="w")

    def _reverse(self):
        try:
            v    = float(self.rv_v.get())
            gain = float(self.rv_gain.get())
        except ValueError:
            messagebox.showwarning("입력 오류", "숫자를 입력하세요."); return
        r = find_resistance_from_voltage(self.rv_sensor.get(), v, gain)
        self.rv_result.configure(text=f"R  ≈  {r:.4f}  Ω")


# ═══════════════════════════════════════════════════════════════════════════════
# 탭 3: 설정
# ═══════════════════════════════════════════════════════════════════════════════

class SettingsTab(ctk.CTkFrame):
    def __init__(self, master, settings: dict, on_save):
        super().__init__(master)
        self._settings = settings
        self._on_save  = on_save
        self._build()

    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        p = {"padx": 20, "pady": 8, "sticky": "w"}
        row = 0

        ctk.CTkLabel(self, text="기본값 설정",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="#a0c4e8").grid(
            row=row, column=0, columnspan=2, sticky="w", padx=20, pady=(20, 10)); row += 1

        s = self._settings
        fields = [
            ("기본 제조사",   "manufacturer",  s.get("manufacturer", "")),
            ("기본 담당자",   "operator",      s.get("operator", "")),
            ("기본 장소",     "location",      s.get("location", "")),
            ("기본 케이블",   "cable",         s.get("cable", "전용케이블")),
            ("기본 모듈명",   "module_name",   s.get("module_name", "")),
            ("기본 모델명",   "model",         s.get("model", "")),
        ]
        self._entries = {}
        for label, key, default in fields:
            ctk.CTkLabel(self, text=label, anchor="w", width=120).grid(
                row=row, column=0, **p)
            e = ctk.CTkEntry(self, width=280)
            e.grid(row=row, column=1, padx=(0, 20), pady=8, sticky="ew")
            if default:
                e.insert(0, default)
            self._entries[key] = e
            row += 1

        # 출력 폴더
        ctk.CTkLabel(self, text="기본 출력 폴더", anchor="w", width=120).grid(
            row=row, column=0, **p)
        out_f = ctk.CTkFrame(self, fg_color="transparent")
        out_f.grid(row=row, column=1, padx=(0,20), pady=8, sticky="ew")
        self.out_dir_entry = ctk.CTkEntry(out_f, width=230)
        self.out_dir_entry.insert(0, s.get("output_dir",
            os.path.join(os.path.dirname(__file__), "..", "outputs")))
        self.out_dir_entry.pack(side="left")
        ctk.CTkButton(out_f, text="…", width=36,
                      command=self._browse_out).pack(side="left", padx=4)
        row += 1

        # 테마
        ctk.CTkLabel(self, text="테마", anchor="w", width=120).grid(
            row=row, column=0, **p)
        self.theme_var = ctk.StringVar(value=s.get("theme", "dark"))
        ctk.CTkOptionMenu(self, variable=self.theme_var,
                          values=["dark","light","system"],
                          command=self._apply_theme, width=140
                          ).grid(row=row, column=1, padx=(0,20), pady=8, sticky="w")
        row += 1

        ctk.CTkButton(self, text="저장", width=120, fg_color=NAVY,
                      command=self._save).grid(
            row=row, column=0, columnspan=2, padx=20, pady=20, sticky="w")

    def _browse_out(self):
        d = filedialog.askdirectory(initialdir=self.out_dir_entry.get())
        if d:
            self.out_dir_entry.delete(0, "end")
            self.out_dir_entry.insert(0, d)

    def _apply_theme(self, theme: str):
        ctk.set_appearance_mode(theme)

    def _save(self):
        data = {k: e.get() for k, e in self._entries.items()}
        data["output_dir"] = self.out_dir_entry.get()
        data["theme"]      = self.theme_var.get()
        _save_settings(data)
        self._settings.update(data)
        self._on_save(data)
        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")


# ═══════════════════════════════════════════════════════════════════════════════
# 메인 앱
# ═══════════════════════════════════════════════════════════════════════════════

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        _apply_ttk_style()
        self._settings = _load_settings()
        ctk.set_appearance_mode(self._settings.get("theme", "dark"))
        self.title("Multi-Channel Sensor Calibration Tool  (MCAL)")
        self.geometry("1340x860")
        self.minsize(1100, 740)
        self._build()

    def _build(self):
        # 헤더 바
        hdr = ctk.CTkFrame(self, height=52, corner_radius=0, fg_color=NAVY)
        hdr.pack(side="top", fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr, text="  ⚙   Multi-Channel Sensor Calibration Tool",
            font=ctk.CTkFont(size=15, weight="bold"), text_color="white",
        ).pack(side="left", padx=16)
        ctk.CTkLabel(
            hdr, text="MCAL  |  PT100 / PT1000 / Strain 350Ω  |  CH01 ~ CH16",
            font=ctk.CTkFont(size=11), text_color="#a0c4e8",
        ).pack(side="left", padx=4)

        # 메인 탭뷰
        tabs = ctk.CTkTabview(self, anchor="nw")
        tabs.pack(side="top", fill="both", expand=True, padx=6, pady=6)
        tabs.add("캘리브레이션")
        tabs.add("3-Wire 레퍼런스")
        tabs.add("설정")

        self._cal_tab = CalibrationTab(tabs.tab("캘리브레이션"), self._settings)
        self._cal_tab.pack(fill="both", expand=True)

        ReferenceTab(tabs.tab("3-Wire 레퍼런스")).pack(fill="both", expand=True)

        SettingsTab(tabs.tab("설정"), self._settings,
                    on_save=self._on_settings_saved).pack(fill="both", expand=True)

    def _on_settings_saved(self, data: dict):
        self._settings.update(data)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
