"""
PDF calibration report generator using ReportLab.
Produces the same content as the DOCX report but in PDF format.
"""
import os
import io
from datetime import datetime
from typing import Dict, List, Optional

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer,
    PageBreak, Image, HRFlowable, KeepTogether,
)
from reportlab.platypus.flowables import Flowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from processing.calibration import ChannelCalibration
from sensors.base import SensorConfig
from reference.three_wire import build_reference_table, STANDARD_GAINS

# ── Colours ──────────────────────────────────────────────────────────────────
C_NAVY   = colors.HexColor("#1F497D")
C_LBLUE  = colors.HexColor("#D6E4F0")
C_ALT    = colors.HexColor("#F2F7FB")
C_PASS   = colors.HexColor("#0070C0")
C_FAIL   = colors.HexColor("#C00000")
C_GREY   = colors.HexColor("#A0A0A0")
C_WHITE  = colors.white
C_BLACK  = colors.black

# ── Font registration ─────────────────────────────────────────────────────────
_FONTS_REGISTERED = False

def _register_fonts():
    global _FONTS_REGISTERED
    if _FONTS_REGISTERED:
        return
    # Try to register Malgun Gothic for Korean text
    font_paths = [
        "/System/Library/Fonts/Supplemental/Malgun Gothic.ttf",
        "/Library/Fonts/Malgun Gothic.ttf",
        "/usr/share/fonts/truetype/malgun/Malgun Gothic.ttf",
    ]
    registered = False
    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("MalgunGothic", path))
                registered = True
                break
            except Exception:
                pass

    if not registered:
        # Fallback: use Apple SD Gothic Neo or NanumGothic
        fallback_paths = [
            "/System/Library/Fonts/AppleSDGothicNeo.ttc",
            "/Library/Fonts/NanumGothic.ttf",
        ]
        for path in fallback_paths:
            if os.path.exists(path):
                try:
                    pdfmetrics.registerFont(TTFont("MalgunGothic", path))
                    registered = True
                    break
                except Exception:
                    pass

    if not registered:
        # Last resort: AppleGothic TTF fallback
        apple_gothic = "/System/Library/Fonts/Supplemental/AppleGothic.ttf"
        if os.path.exists(apple_gothic):
            try:
                pdfmetrics.registerFont(TTFont("MalgunGothic", apple_gothic))
                registered = True
            except Exception:
                pass

    # If still not registered, "MalgunGothic" stays unregistered → _font() returns Helvetica
    _FONTS_REGISTERED = True


def _font():
    """Return font name to use."""
    _register_fonts()
    return "MalgunGothic" if "MalgunGothic" in pdfmetrics.getRegisteredFontNames() else "Helvetica"


# ── Styles ────────────────────────────────────────────────────────────────────

def _styles():
    fn = _font()
    return {
        "title": ParagraphStyle("title", fontName=fn, fontSize=18, textColor=C_NAVY,
                                 alignment=TA_CENTER, spaceAfter=6, spaceBefore=4),
        "h1":    ParagraphStyle("h1", fontName=fn, fontSize=12, textColor=C_NAVY,
                                 leading=16, spaceAfter=4, spaceBefore=6),
        "h2":    ParagraphStyle("h2", fontName=fn, fontSize=10, textColor=C_NAVY,
                                 leading=14, spaceAfter=3, spaceBefore=4),
        "body":  ParagraphStyle("body", fontName=fn, fontSize=8, leading=11),
        "small": ParagraphStyle("small", fontName=fn, fontSize=7, leading=10),
        "hdr":   ParagraphStyle("hdr", fontName=fn, fontSize=8, textColor=C_WHITE,
                                 alignment=TA_CENTER, leading=10),
        "cell":  ParagraphStyle("cell", fontName=fn, fontSize=8, alignment=TA_CENTER, leading=10),
        "cellL": ParagraphStyle("cellL", fontName=fn, fontSize=8, alignment=TA_LEFT, leading=10),
        "cellR": ParagraphStyle("cellR", fontName=fn, fontSize=8, alignment=TA_RIGHT, leading=10),
        "pass":  ParagraphStyle("pass", fontName=fn, fontSize=8, textColor=C_PASS,
                                 alignment=TA_CENTER, leading=10),
        "fail":  ParagraphStyle("fail", fontName=fn, fontSize=8, textColor=C_FAIL,
                                 alignment=TA_CENTER, leading=10),
    }


def _fmt(val: Optional[float], digits=4) -> str:
    if val is None:
        return "-"
    return f"{val:.{digits}f}"


def _pass_style(dev: float, tol: float, styles):
    return styles["pass"] if abs(dev) <= tol else styles["fail"]


# ── Chart helpers ─────────────────────────────────────────────────────────────

def _gain_chart_image(cal: ChannelCalibration, w_cm=8.0, h_cm=5.0) -> Image:
    rs = sorted(cal.voltages_avg.keys())
    v_meas = [cal.voltages_avg[r] for r in rs]
    v_ref  = [cal.voltage_ref[r]  for r in rs]

    fig, ax = plt.subplots(figsize=(w_cm / 2.54, h_cm / 2.54))
    ax.scatter(v_ref, v_meas, color="#1F497D", zorder=5, s=40, label="Measured")
    if len(v_ref) >= 2:
        coeffs = np.polyfit(v_ref, v_meas, 1)
        x_line = np.linspace(min(v_ref)*1.05, max(v_ref)*1.05, 100)
        y_line = np.polyval(coeffs, x_line)
        ax.plot(x_line, y_line, color="#C00000", linewidth=1.5,
                label=f"Gain={cal.gain:.6f}")
    ax.set_xlabel("Ref Voltage (V)", fontsize=7)
    ax.set_ylabel("Meas Voltage (V)", fontsize=7)
    ax.set_title(f"Gain — {cal.channel}", fontsize=8, fontweight="bold")
    ax.legend(fontsize=6)
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.tick_params(labelsize=6)
    fig.tight_layout(pad=0.5)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return Image(buf, width=w_cm*cm, height=h_cm*cm)


def _dev_chart_image(cal: ChannelCalibration, resistances: List[float],
                     w_cm=8.0, h_cm=5.0) -> Image:
    rs = sorted([r for r in resistances if r in cal.dev_final_100])
    dev_100  = [cal.dev_final_100.get(r, 0) for r in rs]
    dev_mean = [cal.dev_final_mean.get(r, 0) for r in rs]
    tol = cal.tolerance_max_100.get(rs[0], 0.385) if rs else 0.385

    x = np.arange(len(rs))
    width = 0.35
    fig, ax = plt.subplots(figsize=(w_cm / 2.54, h_cm / 2.54))
    ax.bar(x - width/2, dev_100,  width, label="2-1 (100Ω)", color="#1F497D", alpha=0.8)
    ax.bar(x + width/2, dev_mean, width, label="2-2 (Mean)", color="#ED7D31", alpha=0.8)
    ax.axhline(y= tol, color="red",   linestyle="--", linewidth=1.0)
    ax.axhline(y=-tol, color="blue",  linestyle="--", linewidth=1.0)
    ax.axhline(y=0,    color="black", linestyle="-",  linewidth=0.5)
    ax.set_xticks(x)
    ax.set_xticklabels([f"{int(r)}Ω" for r in rs], fontsize=6)
    ax.set_ylabel("Dev (Ω)", fontsize=7)
    ax.set_title(f"Deviation — {cal.channel}", fontsize=8, fontweight="bold")
    ax.legend(fontsize=6)
    ax.grid(True, axis="y", linestyle="--", alpha=0.4)
    ax.tick_params(labelsize=6)
    fig.tight_layout(pad=0.5)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return Image(buf, width=w_cm*cm, height=h_cm*cm)


# ── Table builder helpers ─────────────────────────────────────────────────────

def _hdr(text, styles) -> Paragraph:
    return Paragraph(text, styles["hdr"])


def _cell(text, styles, align="center") -> Paragraph:
    key = {"center": "cell", "left": "cellL", "right": "cellR"}.get(align, "cell")
    return Paragraph(str(text), styles[key])


def _dev_cell(dev: Optional[float], tol: float, styles) -> Paragraph:
    if dev is None:
        return _cell("-", styles)
    txt = _fmt(dev, 4)
    s = _pass_style(dev, tol, styles)
    return Paragraph(txt, s)


HDR_STYLE = TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), C_NAVY),
    ("TEXTCOLOR",  (0, 0), (-1, 0), C_WHITE),
    ("FONTSIZE",   (0, 0), (-1, -1), 8),
    ("GRID",       (0, 0), (-1, -1), 0.5, C_GREY),
    ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
    ("ALIGN",      (0, 0), (-1, -1), "CENTER"),
])


# ── PDF Writer ────────────────────────────────────────────────────────────────

class CalibrationPdfWriter:
    def __init__(
        self,
        sensor: SensorConfig,
        calibrations: Dict[str, ChannelCalibration],
        metadata: Optional[Dict[str, str]] = None,
    ):
        self.sensor = sensor
        self.calibrations = calibrations
        self.channels = sorted(calibrations.keys())
        self.meta = metadata or {}
        self.resistances: List[float] = []
        if calibrations:
            first = next(iter(calibrations.values()))
            self.resistances = sorted(first.voltages_avg.keys())

    def _cover_elements(self, styles) -> list:
        fn  = _font()
        now = datetime.now()
        doc_no = self.meta.get("doc_number", f"CAL-{now.strftime('%Y')}-0001")
        rev    = self.meta.get("revision", "00")

        elems = [
            Paragraph("CALIBRATION REPORT", styles["title"]),
            Paragraph(
                f'<font size="8" color="#1F497D">문서번호: {doc_no}　Rev: {rev}</font>',
                ParagraphStyle("dn", fontName=fn, fontSize=8,
                               textColor=C_NAVY, alignment=TA_RIGHT)
            ),
            Spacer(1, 0.4*cm),
        ]

        def _sec_hdr(text):
            return Paragraph(
                f'<b>{text}</b>',
                ParagraphStyle("sh", fontName=fn, fontSize=9,
                               textColor=C_WHITE, alignment=TA_LEFT,
                               backColor=C_NAVY, leftPadding=6)
            )

        def _row(lbl1, val1, lbl2="", val2=""):
            return [
                _cell(lbl1, styles, "left"),
                _cell(val1, styles, "left"),
                _cell(lbl2, styles, "left"),
                _cell(val2, styles, "left"),
            ]

        cw = [3.8*cm, 5.2*cm, 3.8*cm, 5.2*cm]
        lbl_style = TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), C_LBLUE),
            ("BACKGROUND", (2, 0), (2, -1), C_LBLUE),
            ("GRID",       (0, 0), (-1, -1), 0.5, C_GREY),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (1, 0), (1, -1), [C_WHITE]),
            ("ROWBACKGROUNDS", (3, 0), (3, -1), [C_WHITE]),
        ])

        # Section 1: Module Information
        elems.append(_sec_hdr("1. 모듈 정보 (Module Information)"))
        mod_data = [
            _row("모듈명 (Module Name)",  self.meta.get("module_name",""),
                 "모델명 (Model No.)",    self.meta.get("model","")),
            _row("시리얼 번호 (S/N)",      self.meta.get("serial",""),
                 "제조사 (Manufacturer)", self.meta.get("manufacturer","")),
            _row("FW / SW 버전",          self.meta.get("fw_version",""), "", ""),
        ]
        t1 = Table(mod_data, colWidths=cw)
        t1.setStyle(lbl_style)
        elems += [t1, Spacer(1, 0.25*cm)]

        # Section 2: Calibration Conditions
        elems.append(_sec_hdr("2. 교정 조건 (Calibration Conditions)"))
        cond_data = [
            _row("교정 일자 (Date)",  self.meta.get("date", now.strftime("%Y.%m.%d")),
                 "교정 장소 (Location)", self.meta.get("location","")),
            _row("담당자 (Technician)", self.meta.get("operator",""),
                 "온도 / 습도",          self.meta.get("temp_humidity","")),
        ]
        t2 = Table(cond_data, colWidths=cw)
        t2.setStyle(lbl_style)
        elems += [t2, Spacer(1, 0.25*cm)]

        # Section 3: Calibration Setup
        elems.append(_sec_hdr("3. 교정 설정 (Calibration Setup)"))
        setup_data = [
            _row("센서 타입 (Sensor)",   self.sensor.name,
                 "케이블 (Cable)",        self.meta.get("cable","전용케이블")),
            _row("공칭 저항 (Nominal R)", f"{self.sensor.r_nominal:.0f} Ω",
                 "Inst. Amp. Gain",       str(self.meta.get("inst_amp_gain","1"))),
            _row("모사저항 값",
                 " / ".join(f"{r:.0f}Ω" for r in self.resistances),
                 "허용 편차 (Tolerance)", f"±{self.sensor.tolerance_ohm:.3f} Ω"),
            _row("샘플링 속도",          f"{self.meta.get('sampling_hz',100)} Hz",
                 "측정 시간",             f"{self.meta.get('duration_sec','-')} 초"),
            _row("채널 수 (Channels)",   str(len(self.channels)),
                 "레퍼런스 수식",          self.sensor.ref_formula),
        ]
        t3 = Table(setup_data, colWidths=cw)
        t3.setStyle(lbl_style)
        elems += [t3, Spacer(1, 0.4*cm)]

        # Signature block
        elems.append(_sec_hdr("서명 (Signatures)"))
        sig_data = [[
            _hdr("작성 (Prepared)", styles),
            _hdr("검토 (Reviewed)", styles),
            _hdr("승인 (Approved)", styles),
        ], ["", "", ""]]
        sig_cw = [6.0*cm, 6.0*cm, 6.0*cm]
        ts = Table(sig_data, colWidths=sig_cw, rowHeights=[None, 1.5*cm])
        ts.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), C_NAVY),
            ("TEXTCOLOR",  (0, 0), (-1, 0), C_WHITE),
            ("GRID",       (0, 0), (-1, -1), 0.5, C_GREY),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ]))
        elems.append(ts)
        return elems

    def _summary_elements(self, styles) -> list:
        rs = self.resistances
        n_rs = len(rs)
        elems = []

        for method, offset_key, dev_key, title in [
            ("2-1", "offset_100",  "dev_final_100",  "4. 교정 결과 요약 — Method 2-1  (R_nom 기준 Offset)"),
            ("2-2", "offset_mean", "dev_final_mean", "5. 교정 결과 요약 — Method 2-2  (편차 평균 Offset)"),
        ]:
            elems.append(Spacer(1, 0.3*cm))
            elems.append(Paragraph(title, styles["h1"]))

            header = [_hdr("채널명", styles), _hdr("Excitation", styles),
                      _hdr("Gain", styles), _hdr("Offset", styles)]
            header += [_hdr(f"{r:.0f}Ω", styles) for r in rs]
            data = [header]

            for ri, ch in enumerate(self.channels):
                cal = self.calibrations[ch]
                offset_val = getattr(cal, offset_key)
                dev_dict   = getattr(cal, dev_key)
                tol = self.sensor.tolerance_ohm

                row_bg = C_ALT if ri % 2 == 0 else C_WHITE
                row_data = [
                    _cell(ch, styles),
                    _cell(f"{cal.excitation:.4f}", styles),
                    _cell(f"{cal.gain:.6f}", styles),
                    _cell(f"{offset_val:.6f}", styles),
                ]
                for r in rs:
                    dev = dev_dict.get(r)
                    row_data.append(_dev_cell(dev, tol, styles))
                data.append(row_data)

            cw = [3.0*cm, 2.2*cm, 2.2*cm, 2.4*cm] + [1.8*cm] * n_rs
            t = Table(data, colWidths=cw, repeatRows=1)
            ts = TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), C_NAVY),
                ("TEXTCOLOR",  (0, 0), (-1, 0), C_WHITE),
                ("GRID",       (0, 0), (-1, -1), 0.5, C_GREY),
                ("FONTSIZE",   (0, 0), (-1, -1), 8),
                ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_WHITE, C_ALT]),
            ])
            t.setStyle(ts)
            elems.append(t)

        return elems

    def _channel_elements(self, ch_idx: int, ch: str, styles) -> list:
        cal = self.calibrations[ch]
        rs = self.resistances
        tol = self.sensor.tolerance_ohm
        elems = [PageBreak()]

        # Header
        elems.append(Paragraph(
            f"Calibration Sheet ({ch_idx + 1}/{len(self.channels)})  —  {ch}",
            styles["h1"]
        ))

        # Result banner
        hdr_row = [_hdr("", styles), _hdr("채널명", styles), _hdr("Excitation", styles),
                   _hdr("Gain", styles), _hdr("Offset", styles)]
        hdr_row += [_hdr(f"{r:.0f}Ω", styles) for r in rs]
        banner_data = [hdr_row]
        for method_label, offset_val, dev_dict in [
            ("2-1\n(100Ω)", cal.offset_100,  cal.dev_final_100),
            ("2-2\n(Mean)", cal.offset_mean, cal.dev_final_mean),
        ]:
            row = [_cell(method_label, styles), _cell(ch, styles),
                   _cell(f"{cal.excitation:.4f}", styles),
                   _cell(f"{cal.gain:.6f}", styles),
                   _cell(f"{offset_val:.6f}", styles)]
            for r in rs:
                dev = dev_dict.get(r)
                row.append(_dev_cell(dev, tol, styles))
            banner_data.append(row)

        cw = [1.5*cm, 2.5*cm, 2.0*cm, 2.2*cm, 2.2*cm] + [1.8*cm] * len(rs)
        banner = Table(banner_data, colWidths=cw)
        banner.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), C_NAVY),
            ("TEXTCOLOR",  (0, 0), (-1, 0), C_WHITE),
            ("BACKGROUND", (0, 1), (0, -1), C_LBLUE),
            ("GRID",       (0, 0), (-1, -1), 0.5, C_GREY),
            ("FONTSIZE",   (0, 0), (-1, -1), 8),
            ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ]))
        elems.append(banner)
        elems.append(Spacer(1, 0.3*cm))

        # Charts
        gain_img = _gain_chart_image(cal, w_cm=8.5, h_cm=5.5)
        dev_img  = _dev_chart_image(cal, rs, w_cm=8.5, h_cm=5.5)
        chart_tbl = Table([[gain_img, dev_img]], colWidths=[8.7*cm, 8.7*cm])
        chart_tbl.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "MIDDLE"),
                                        ("ALIGN",  (0,0), (-1,-1), "CENTER")]))
        elems.append(chart_tbl)
        elems.append(Spacer(1, 0.3*cm))

        # Detail table
        elems.append(Paragraph("상세 정보 (Detail)", styles["h2"]))

        def _drow(s1, s2, vals, bg=None):
            row = [_cell(s1, styles, "left"), _cell(s2, styles, "left")]
            row += [_cell(_fmt(v, 4) if isinstance(v, float) else str(v) if v is not None else "-",
                          styles) for v in vals]
            return row

        ddata = []
        # header
        ddata.append([_hdr("구분", styles), _hdr("항목", styles)] +
                     [_hdr(f"{r:.0f}Ω", styles) for r in rs])

        ddata.append(_drow("1) Decimal", "AVG", [int(cal.decimals_avg.get(r,0)) for r in rs]))
        ddata.append(_drow("", "MIN", [int(cal.decimals_min.get(r,0)) for r in rs]))
        ddata.append(_drow("", "MAX", [int(cal.decimals_max.get(r,0)) for r in rs]))
        ddata.append(_drow("2) 전압(before gain)", "Ref", [cal.voltage_ref.get(r) for r in rs]))
        ddata.append(_drow("", "AVG", [cal.voltages_avg.get(r) for r in rs]))
        ddata.append(_drow("3) 저항(before gain)", "Ref", [r for r in rs]))
        ddata.append(_drow("", "AVG",  [cal.r_before_gain.get(r) for r in rs]))
        ddata.append(_drow("", "편차", [cal.dev_before_gain.get(r) for r in rs]))
        ddata.append(_drow("4) Gain", f"{cal.gain:.8f}", [None]*len(rs)))
        ddata.append(_drow("5) 저항(after gain)", "AVG",  [cal.r_after_gain_avg.get(r) for r in rs]))
        ddata.append(_drow("", "MIN",  [cal.r_after_gain_min.get(r) for r in rs]))
        ddata.append(_drow("", "MAX",  [cal.r_after_gain_max.get(r) for r in rs]))
        ddata.append(_drow("", "편차", [cal.dev_after_gain.get(r) for r in rs]))
        ddata.append(_drow("6) Offset", f"100Ω:{cal.offset_100:.6f} / Mean:{cal.offset_mean:.6f}",
                           [None]*len(rs)))
        ddata.append(_drow("7) 2-1 (100Ω offset)", "AVG",    [cal.r_final_100_avg.get(r) for r in rs]))
        ddata.append(_drow("", "편차",    [cal.dev_final_100.get(r) for r in rs]))
        ddata.append(_drow("", "허용(+)", [cal.tolerance_max_100.get(r) for r in rs]))
        ddata.append(_drow("", "허용(-)", [cal.tolerance_min_100.get(r) for r in rs]))
        ddata.append(_drow("8) 2-2 (Mean offset)", "AVG",    [cal.r_final_mean_avg.get(r) for r in rs]))
        ddata.append(_drow("", "편차",    [cal.dev_final_mean.get(r) for r in rs]))
        ddata.append(_drow("", "허용(+)", [cal.tolerance_max_mean.get(r) for r in rs]))
        ddata.append(_drow("", "허용(-)", [cal.tolerance_min_mean.get(r) for r in rs]))

        cw2 = [3.5*cm, 2.5*cm] + [1.9*cm]*len(rs)
        dtbl = Table(ddata, colWidths=cw2, repeatRows=1)
        dtbl.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), C_NAVY),
            ("TEXTCOLOR",     (0, 0), (-1, 0), C_WHITE),
            ("GRID",          (0, 0), (-1, -1), 0.5, C_GREY),
            ("FONTSIZE",      (0, 0), (-1, -1), 7),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [C_WHITE, C_ALT]),
        ]))
        elems.append(dtbl)
        return elems

    def _3w_reference_elements(self, styles) -> list:
        """3-Wire reference voltage table page."""
        gains       = [g for g in STANDARD_GAINS if g in (1, 10, 100)]
        sensor_type = self.sensor.sensor_type
        r_nom       = self.sensor.r_nominal
        try:
            rs, gs, tbl = build_reference_table(
                sensor_type, r_nom * 0.5, r_nom * 1.5, 1.0, gains
            )
        except Exception:
            return [PageBreak(), Paragraph("3-Wire 레퍼런스 생성 실패", styles["h1"])]

        elems = [PageBreak(),
                 Paragraph(f"3-Wire Reference Voltage  [{self.sensor.name}]", styles["h1"]),
                 Spacer(1, 0.2*cm)]

        header = ([_hdr("R (Ω)", styles), _hdr("ΔR (Ω)", styles)] +
                  [_hdr(f"Gain={g:.0f}", styles) for g in gs])
        data = [header]
        for ri, r in enumerate(rs):
            row = [_cell(f"{r:.0f}", styles), _cell(f"{r-r_nom:+.0f}", styles)]
            row += [_cell(f"{v:.6f}", styles) for v in tbl[r]]
            data.append(row)

        n_g  = len(gs)
        cw   = [2.0*cm, 2.0*cm] + [3.2*cm] * n_g
        t    = Table(data, colWidths=cw, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0),  C_NAVY),
            ("TEXTCOLOR",     (0, 0), (-1, 0),  C_WHITE),
            ("GRID",          (0, 0), (-1, -1), 0.5, C_GREY),
            ("FONTSIZE",      (0, 0), (-1, -1), 7),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [C_WHITE, C_ALT]),
        ]))
        elems.append(t)
        return elems

    def write(self, output_path: str):
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        _register_fonts()
        styles = _styles()

        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            leftMargin=2*cm, rightMargin=2*cm,
            topMargin=2*cm,  bottomMargin=2*cm,
        )

        story = []
        story += self._cover_elements(styles)
        story += self._summary_elements(styles)
        for idx, ch in enumerate(self.channels):
            story += self._channel_elements(idx, ch, styles)
        story += self._3w_reference_elements(styles)

        doc.build(story)
        return output_path
