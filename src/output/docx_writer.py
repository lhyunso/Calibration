"""
DOCX calibration report generator.
Uses 'Calibration sheet template.docx' as the visual base:
  - Rows 0-11 of the template table are preserved (header + info fields)
  - Rows 12+ are removed and replaced with calibration-specific content
  - Per-channel detail pages follow the same styling

Template fonts  : 나눔스퀘어 ExtraBold (title), KoPub돋움체 Medium (body)
Template colours: section header #666666 · table header #F2F2F2 · border black
Page            : A4, margins top/bottom ~2 cm, left/right ~1.5 cm (from template)
"""
import os
import io
from datetime import datetime
from typing import Dict, List, Optional

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

from config import REFERENCE_DIR
from processing.calibration import ChannelCalibration
from sensors.base import SensorConfig


# ── Asset paths ───────────────────────────────────────────────────────────────
TEMPLATE_PATH = os.path.join(REFERENCE_DIR, "Calibration sheet template.docx")

# ── Colour palette ────────────────────────────────────────────────────────────
C_GREY_BG  = RGBColor(0x66, 0x66, 0x66)   # dark grey — section header bg
C_THEAD_BG = RGBColor(0xF2, 0xF2, 0xF2)   # light grey — table column header bg
C_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
C_BLACK    = RGBColor(0x00, 0x00, 0x00)
C_ROW_ALT  = RGBColor(0xFA, 0xFA, 0xFA)   # alternating data row tint
C_PASS     = RGBColor(0x00, 0x70, 0xC0)   # blue  — within tolerance
C_FAIL     = RGBColor(0xC0, 0x00, 0x00)   # red   — out of tolerance

# ── Fonts (same as template) ──────────────────────────────────────────────────
F_TITLE = "나눔스퀘어 ExtraBold"
F_BODY  = "KoPub돋움체 Medium"

# ── Content width (template: A4 11906 DXA − margins 850×2 = 10206 DXA ≈ 18 cm) ─
CONTENT_W  = 18.0          # cm
_CM2DXA    = 1440 / 2.54   # 1 cm → DXA (twips)


def _cm_dxa(cm: float) -> int:
    return int(cm * _CM2DXA)


# ── Low-level XML helpers ─────────────────────────────────────────────────────

def _set_font(run, name: str, size: float, bold: bool = False,
              color: Optional[RGBColor] = None) -> None:
    run.bold = bold
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    rPr    = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), name)


def _set_cell_bg(cell, rgb: RGBColor) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")
    tcPr.append(shd)


def _set_cell_borders(cell, color: str = "000000", size: int = 6) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    bdr  = OxmlElement("w:tcBorders")
    for side in ("top", "bottom", "left", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "single")
        b.set(qn("w:sz"),    str(size))
        b.set(qn("w:color"), color)
        b.set(qn("w:space"), "0")
        bdr.append(b)
    tcPr.append(bdr)


def _cell_write(cell, text: str, font: str = F_BODY, size: float = 9,
                bold: bool = False,
                align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.LEFT,
                color: Optional[RGBColor] = None) -> None:
    para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    para.clear()
    para.alignment = align
    run = para.add_run(text)
    _set_font(run, font, size, bold=bold, color=color)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _set_row_height_cm(row, cm: float) -> None:
    trPr     = row._tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"),   str(_cm_dxa(cm)))
    trHeight.set(qn("w:hRule"), "atLeast")
    trPr.append(trHeight)


def _para_spacing(para, before_dxa: int = 0, after_dxa: int = 0) -> None:
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), str(before_dxa))
    sp.set(qn("w:after"),  str(after_dxa))
    para._p.get_or_add_pPr().append(sp)


# ── Template field helpers ────────────────────────────────────────────────────

def _replace_para_text(para, new_text: str) -> None:
    """Replace all runs in a paragraph with new_text in the first run."""
    if not para.runs:
        run = para.add_run(new_text)
        return
    para.runs[0].text = new_text
    for run in para.runs[1:]:
        run.text = ""


def _replace_cell_text(cell, new_text: str) -> None:
    """Replace the first paragraph's text in a cell, preserving font."""
    if cell.paragraphs:
        _replace_para_text(cell.paragraphs[0], new_text)


# ── Composite cell helpers (for new tables added after template) ──────────────

def _col_header_cell(cell, text: str, size: float = 8,
                     align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER) -> None:
    _set_cell_bg(cell, C_THEAD_BG)
    _cell_write(cell, text, F_BODY, size, align=align)
    _set_cell_borders(cell)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _data_cell(cell, text: str, size: float = 8,
               align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER,
               bold: bool = False,
               color: Optional[RGBColor] = None,
               bg: Optional[RGBColor] = None) -> None:
    if bg is not None:
        _set_cell_bg(cell, bg)
    _cell_write(cell, text, F_BODY, size, bold=bold, align=align, color=color)
    _set_cell_borders(cell)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def _section_header_para(doc: Document, text: str) -> None:
    """Full-width paragraph with dark-grey shading — section divider."""
    para = doc.add_paragraph()
    pPr  = para._p.get_or_add_pPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  "666666")
    pPr.append(shd)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "60")
    sp.set(qn("w:after"),  "60")
    pPr.append(sp)
    run = para.add_run("  " + text)
    _set_font(run, F_BODY, 9, bold=True, color=C_WHITE)


# ── Formatting utilities ──────────────────────────────────────────────────────

def _fmt(val: Optional[float], digits: int = 6) -> str:
    return "-" if val is None else f"{val:.{digits}f}"


def _pass_color(dev: float, tol: float) -> RGBColor:
    return C_PASS if abs(dev) <= tol else C_FAIL


def _is_pass(cal: ChannelCalibration, rs: List[float],
             method: str = "100") -> bool:
    dev_d = cal.dev_final_100  if method == "100" else cal.dev_final_mean
    tol_d = cal.tolerance_max_100 if method == "100" else cal.tolerance_max_mean
    return all(abs(dev_d.get(r, 0)) <= tol_d.get(r, 1e9)
               for r in rs if r in dev_d)


# ── Chart generators ──────────────────────────────────────────────────────────

def _make_gain_chart(cal: ChannelCalibration) -> io.BytesIO:
    rs = sorted(r for r in cal.dev_before_gain
                if r in cal.dev_after_gain and r in cal.dev_final_100)
    fig, ax = plt.subplots(figsize=(5.5, 3.2))
    if rs:
        x   = np.arange(len(rs))
        tol = cal.tolerance_max_100.get(rs[0], 0.385)
        ax.plot(x, [cal.dev_before_gain[r] for r in rs],
                "o-", color="#A6A6A6", lw=1.5, ms=5, label="① Before Gain")
        ax.plot(x, [cal.dev_after_gain[r]  for r in rs],
                "s-", color="#1F497D", lw=1.5, ms=5, label="② After Gain")
        ax.plot(x, [cal.dev_final_100[r]   for r in rs],
                "^-", color="#ED7D31", lw=1.5, ms=5, label="③ Offset 2-1")
        ax.plot(x, [cal.dev_final_mean[r]  for r in rs],
                "D-", color="#70AD47", lw=1.5, ms=5, label="④ Offset 2-2")
        ax.axhline( tol, color="red",   ls="--", lw=1.0, label=f"+{tol}Ω")
        ax.axhline(-tol, color="blue",  ls="--", lw=1.0, label=f"-{tol}Ω")
        ax.axhline(0,    color="black", ls="-",  lw=0.6)
        ax.set_xticks(x)
        ax.set_xticklabels([f"{int(r)}" for r in rs], fontsize=7)
        ax.legend(fontsize=6, loc="best", ncol=2)
    ax.set_xlabel("Resistance (Ω)", fontsize=8)
    ax.set_ylabel("Deviation (Ω)",  fontsize=8)
    ax.set_title(f"Calibration Effect — {cal.channel}",
                 fontsize=9, fontweight="bold")
    ax.grid(True, linestyle="--", alpha=0.4)
    ax.tick_params(labelsize=7)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf


def _make_deviation_chart(cal: ChannelCalibration,
                          resistances: List[float]) -> io.BytesIO:
    rs      = sorted(r for r in resistances if r in cal.dev_final_100)
    dev_100 = [cal.dev_final_100.get(r, 0)  for r in rs]
    dev_m   = [cal.dev_final_mean.get(r, 0) for r in rs]
    tol     = cal.tolerance_max_100.get(rs[0], 0.385) if rs else 0.385
    x, w    = np.arange(len(rs)), 0.35
    fig, ax = plt.subplots(figsize=(5.5, 3.2))
    ax.bar(x - w/2, dev_100, w, label="Method 2-1 (100Ω offset)",
           color="#1F497D", alpha=0.8)
    ax.bar(x + w/2, dev_m,   w, label="Method 2-2 (Mean offset)",
           color="#ED7D31", alpha=0.8)
    ax.axhline( tol, color="red",   ls="--", lw=1.0, label=f"+Tol ({tol}Ω)")
    ax.axhline(-tol, color="blue",  ls="--", lw=1.0, label=f"-Tol ({-tol}Ω)")
    ax.axhline(0,    color="black", ls="-",  lw=0.5)
    ax.set_xticks(x)
    ax.set_xticklabels([f"{int(r)}Ω" for r in rs], fontsize=7)
    ax.set_ylabel("Deviation (Ω)", fontsize=8)
    ax.set_title(f"Calibrated Deviation — {cal.channel}",
                 fontsize=9, fontweight="bold")
    ax.legend(fontsize=6, loc="upper right")
    ax.grid(True, axis="y", linestyle="--", alpha=0.4)
    ax.tick_params(labelsize=7)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf


# ── Document writer ───────────────────────────────────────────────────────────

class CalibrationDocxWriter:
    """
    Generates a DOCX calibration report using the DANAM template layout.

    The template's first 12 rows (header + info fields) are preserved verbatim
    (including logo, fonts, and styling).  Rows 12 onwards are removed and
    replaced with calibration-specific sections.
    """

    def __init__(
        self,
        sensor: SensorConfig,
        calibrations: Dict[str, ChannelCalibration],
        metadata: Optional[Dict[str, str]] = None,
    ):
        self.sensor       = sensor
        self.calibrations = calibrations
        self.channels     = sorted(calibrations.keys())
        self.meta         = metadata or {}
        self.resistances: List[float] = []
        if calibrations:
            first = next(iter(calibrations.values()))
            self.resistances = sorted(first.voltages_avg.keys())
        self.tolerance = float(
            self.meta.get("tolerance_ohm") or sensor.tolerance_ohm
        )

    # ── Step 1: Update template info fields ───────────────────────────────────

    def _update_template_fields(self, tbl) -> None:
        """
        Rewrite the template's info fields (rows 0-11) with MCAL metadata.

        Template row layout:
          Row 0  : Title (cols 0-2 merged) | Logo (cols 3-4, image)
          Row 1  : "Document number :" (cols 0-2 merged)
          Row 2  : spacer
          Row 3  : label [col 0] | value [cols 1-2] | empty [cols 3-4]
          Rows 4-7  : same as row 3
          Row 8  : spacer
          Rows 9-11 : same as row 3
        """
        m = self.meta

        # ── Row 0: Title ──────────────────────────────────────────────────────
        # cells[0] = cells[1] = cells[2] (merged) — contains the title text
        _replace_cell_text(tbl.rows[0].cells[0], "캘리브레이션 성적서")

        # ── Row 1: Document number (3-col merged label cell) ──────────────────
        doc_no = m.get("doc_number", "")
        rev    = m.get("revision", "00")
        _replace_cell_text(tbl.rows[1].cells[0],
                           f"문서번호 :  {doc_no}   Rev. {rev}")

        # ── Rows 3-7: label / value pairs ─────────────────────────────────────
        fields = [
            (3,  "장치명",     m.get("module_name", "")),
            (4,  "모델번호",   m.get("model", "")),
            (5,  "시리얼번호", m.get("serial", "")),
            (6,  "제조사",     m.get("manufacturer", "")),
            (7,  "교정 유형",  f"{self.sensor.name} Sensor Calibration"),
        ]
        for row_i, label, value in fields:
            _replace_cell_text(tbl.rows[row_i].cells[0], label)
            # cells[1] == cells[2] (merged value cell)
            _replace_cell_text(tbl.rows[row_i].cells[1], value)

        # ── Rows 9-11: label / value pairs ────────────────────────────────────
        fields2 = [
            (9,  "담당자",     m.get("operator", "")),
            (10, "시험 장소",  m.get("location", "")),
            (11, "시험 일자",  m.get("date",
                                    datetime.now().strftime("%Y.%m.%d"))),
        ]
        for row_i, label, value in fields2:
            _replace_cell_text(tbl.rows[row_i].cells[0], label)
            _replace_cell_text(tbl.rows[row_i].cells[1], value)

    # ── Step 2: Trim template table rows 12+ ─────────────────────────────────

    @staticmethod
    def _trim_table(tbl, keep_rows: int) -> None:
        """Delete all rows from index keep_rows onwards."""
        tbl_el   = tbl._tbl
        all_rows = tbl_el.findall(qn("w:tr"))
        for row_el in all_rows[keep_rows:]:
            tbl_el.remove(row_el)

    # ── Step 3: Calibration Setup section ────────────────────────────────────

    def _add_setup_section(self, doc: Document) -> None:
        _section_header_para(doc, "교정 설정  /  Calibration Setup")

        m     = self.meta
        exc   = float(m.get("excitation_ma",
                            str(self.sensor.excitation * 1000))
                      or str(self.sensor.excitation * 1000))
        gain  = m.get("inst_amp_gain", str(int(self.sensor.inst_amp_gain)))
        cable = m.get("cable", "전용케이블")
        rs_str = "  /  ".join(f"{r:.0f} Ω" for r in self.resistances)

        rows_data = [
            ("센서 타입",   self.sensor.name,
             "공칭 저항",   f"{self.sensor.r_nominal:.0f} Ω"),
            ("여기 전류",   f"{exc:.3g} mA",
             "Inst. Gain",  gain),
            ("허용 오차",   f"±{self.tolerance:.4f} Ω",
             "케이블",      cable),
            ("모사 저항값", rs_str,   "", ""),
        ]
        col_widths = [Cm(3.2), Cm(5.8), Cm(3.2), Cm(5.8)]

        tbl = doc.add_table(rows=len(rows_data), cols=4)
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        tbl.autofit   = False

        for ri, (l1, v1, l2, v2) in enumerate(rows_data):
            row = tbl.rows[ri]
            for ci, cell in enumerate(row.cells):
                cell.width = col_widths[ci]

            _set_cell_bg(row.cells[0], C_THEAD_BG)
            _cell_write(row.cells[0], "  " + l1, F_BODY, 8.5)
            _set_cell_borders(row.cells[0])
            _cell_write(row.cells[1], "  " + v1, F_BODY, 8.5)
            _set_cell_borders(row.cells[1])

            if l2:
                _set_cell_bg(row.cells[2], C_THEAD_BG)
                _cell_write(row.cells[2], "  " + l2, F_BODY, 8.5)
                _set_cell_borders(row.cells[2])
                _cell_write(row.cells[3], "  " + v2, F_BODY, 8.5)
                _set_cell_borders(row.cells[3])
            else:
                # 모사 저항값 row: merge cols 2-3 for wider display
                merged = row.cells[2].merge(row.cells[3])
                _cell_write(merged, "", F_BODY, 8.5)
                _set_cell_borders(merged)

            _set_row_height_cm(row, 0.58)

    # ── Step 4: Channel summary tables ───────────────────────────────────────

    def _add_summary(self, doc: Document) -> None:
        _section_header_para(doc, "교정 결과 요약  /  Calibration Summary")

        rs   = self.resistances
        n_rs = len(rs)

        # column widths: 채널(1.8) + Exc(1.4) + G_cal(2.1) + V_offset(2.8) + 판정(1.5) = 9.6
        fixed_cm = 9.6
        r_cm     = round((CONTENT_W - fixed_cm) / max(n_rs, 1), 3)
        col_w    = ([Cm(1.8), Cm(1.4), Cm(2.1), Cm(2.8)]
                    + [Cm(r_cm)] * n_rs
                    + [Cm(1.5)])
        n_cols   = 5 + n_rs

        for method, label, dev_attr, tol_attr, off_attr in [
            ("100",  "Method 2-1  (100 Ω Offset)",
             "dev_final_100",  "tolerance_max_100",  "offset_100_v"),
            ("mean", "Method 2-2  (Mean Offset)",
             "dev_final_mean", "tolerance_max_mean", "offset_mean_v"),
        ]:
            sp  = doc.add_paragraph()
            rsp = sp.add_run(f"  {label}")
            _set_font(rsp, F_BODY, 8.5, bold=True, color=C_GREY_BG)
            _para_spacing(sp, before_dxa=110, after_dxa=20)

            tbl = doc.add_table(rows=1 + len(self.channels), cols=n_cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
            tbl.autofit   = False

            hdr    = tbl.rows[0]
            h_lbls = (["채널", "Exc.(mA)", "G_cal", "V_offset [V]"]
                      + [f"{r:.0f} Ω" for r in rs] + ["판정"])
            for ci, (cell, lbl) in enumerate(zip(hdr.cells, h_lbls)):
                cell.width = col_w[ci]
                _col_header_cell(cell, lbl, size=8)
            _set_row_height_cm(hdr, 0.52)

            for ri, ch in enumerate(self.channels):
                cal  = self.calibrations[ch]
                row  = tbl.rows[1 + ri]
                bg   = C_ROW_ALT if ri % 2 == 0 else None
                devs = getattr(cal, dev_attr)
                tols = getattr(cal, tol_attr)
                off  = getattr(cal, off_attr)

                for ci, (cell, txt) in enumerate(zip(row.cells, [
                        ch,
                        f"{cal.excitation * 1000:.3g}",
                        f"{cal.gain:.6f}",
                        f"{off:.6f}"])):
                    cell.width = col_w[ci]
                    _data_cell(cell, txt, size=8,
                               align=WD_ALIGN_PARAGRAPH.CENTER, bg=bg)

                for j, r in enumerate(rs):
                    dev = devs.get(r)
                    tol = tols.get(r, self.tolerance)
                    txt = _fmt(dev, 4) if dev is not None else "-"
                    clr = _pass_color(dev, tol) if dev is not None else None
                    cell = row.cells[4 + j]
                    cell.width = col_w[4 + j]
                    _data_cell(cell, txt, size=8,
                               align=WD_ALIGN_PARAGRAPH.CENTER,
                               color=clr, bg=bg)

                ok   = _is_pass(cal, rs, method)
                cell = row.cells[4 + n_rs]
                cell.width = col_w[4 + n_rs]
                _data_cell(cell, "합격" if ok else "불합격", size=8,
                           align=WD_ALIGN_PARAGRAPH.CENTER,
                           color=C_PASS if ok else C_FAIL, bg=bg)
                _set_row_height_cm(row, 0.52)

    # ── Step 5: Per-channel detail page ──────────────────────────────────────

    def _add_channel_page(self, doc: Document,
                          ch_idx: int, ch: str) -> None:
        doc.add_page_break()
        cal  = self.calibrations[ch]
        rs   = self.resistances
        n_rs = len(rs)

        # ── Channel header ───────────────────────────────────────────────────
        header_p = doc.add_paragraph()
        r_ch = header_p.add_run(f"채널 상세  —  {ch}")
        _set_font(r_ch, F_BODY, 11, bold=True, color=C_BLACK)
        header_p.add_run("    ")
        r_info = header_p.add_run(
            f"{self.meta.get('doc_number', '')}  "
            f"Rev.{self.meta.get('revision', '00')}     "
            f"({ch_idx + 1} / {len(self.channels)})"
        )
        _set_font(r_info, F_BODY, 8, color=C_GREY_BG)
        _para_spacing(header_p, before_dxa=0, after_dxa=80)

        # ── Coefficient banner (2-1 / 2-2) ───────────────────────────────────
        _section_header_para(doc, "교정 계수  /  Calibration Coefficients")

        # cols: 구분(1.2) | 채널(1.8) | Exc(1.4) | G_cal(2.1) | V_offset(2.2) | Rs… | 판정(1.5)
        fixed_ch  = 10.2
        r_cm_ch   = round((CONTENT_W - fixed_ch) / max(n_rs, 1), 3)
        col_w_ch  = ([Cm(1.2), Cm(1.8), Cm(1.4), Cm(2.1), Cm(2.2)]
                     + [Cm(r_cm_ch)] * n_rs + [Cm(1.5)])
        n_cols_ch = 6 + n_rs

        tbl = doc.add_table(rows=3, cols=n_cols_ch)
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        tbl.autofit   = False

        hdr    = tbl.rows[0]
        h_lbls = (["구분", "채널", "Exc.(mA)", "G_cal", "V_offset [V]"]
                  + [f"{r:.0f}Ω" for r in rs] + ["판정"])
        for ci, (cell, lbl) in enumerate(zip(hdr.cells, h_lbls)):
            cell.width = col_w_ch[ci]
            _col_header_cell(cell, lbl, size=8)
        _set_row_height_cm(hdr, 0.52)

        for row_i, (m_lbl, dev_d, tol_d, off_val, method) in enumerate([
            ("2-1\n100Ω",
             cal.dev_final_100,  cal.tolerance_max_100,  cal.offset_100_v,  "100"),
            ("2-2\nMean",
             cal.dev_final_mean, cal.tolerance_max_mean, cal.offset_mean_v, "mean"),
        ], start=1):
            row = tbl.rows[row_i]
            for ci, cell in enumerate(row.cells):
                cell.width = col_w_ch[ci]

            _set_cell_bg(row.cells[0], C_THEAD_BG)
            _cell_write(row.cells[0], m_lbl, F_BODY, 7.5, bold=True,
                        align=WD_ALIGN_PARAGRAPH.CENTER)
            _set_cell_borders(row.cells[0])
            row.cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            for ci, txt in enumerate([
                ch,
                f"{cal.excitation * 1000:.3g}",
                f"{cal.gain:.6f}",
                f"{off_val:.6f}",
            ], start=1):
                _data_cell(row.cells[ci], txt, size=8,
                           align=WD_ALIGN_PARAGRAPH.CENTER)

            for j, r in enumerate(rs):
                dev = dev_d.get(r)
                tol = tol_d.get(r, self.tolerance)
                txt = _fmt(dev, 4) if dev is not None else "-"
                _data_cell(row.cells[5 + j], txt, size=8,
                           align=WD_ALIGN_PARAGRAPH.CENTER,
                           color=_pass_color(dev, tol) if dev is not None else None)

            ok = _is_pass(cal, rs, method)
            _data_cell(row.cells[5 + n_rs],
                       "합격" if ok else "불합격", size=8,
                       align=WD_ALIGN_PARAGRAPH.CENTER,
                       color=C_PASS if ok else C_FAIL)
            _set_row_height_cm(row, 0.58)

        # ── Charts (side by side) ─────────────────────────────────────────────
        doc.add_paragraph()
        chart_tbl = doc.add_table(rows=1, cols=2)
        chart_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for ci, buf in enumerate([_make_gain_chart(cal),
                                   _make_deviation_chart(cal, rs)]):
            chart_tbl.rows[0].cells[ci].paragraphs[0].add_run(
            ).add_picture(buf, width=Cm(8.8))

        # ── Detail stats table ────────────────────────────────────────────────
        doc.add_paragraph()
        _section_header_para(doc, "상세 측정 데이터  /  Measurement Detail")

        fixed_dt = 5.0
        r_cm_dt  = round((CONTENT_W - fixed_dt) / max(n_rs, 1), 3)
        col_w_dt = [Cm(2.6), Cm(2.4)] + [Cm(r_cm_dt)] * n_rs
        n_dt     = 2 + n_rs

        dtbl = doc.add_table(rows=0, cols=n_dt)
        dtbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        dtbl.autofit   = False

        def _hrow():
            row = dtbl.add_row()
            _col_header_cell(row.cells[0], "구분", size=8,
                             align=WD_ALIGN_PARAGRAPH.LEFT)
            _col_header_cell(row.cells[1], "항목", size=8,
                             align=WD_ALIGN_PARAGRAPH.LEFT)
            for j, r in enumerate(rs):
                _col_header_cell(row.cells[2 + j], f"{r:.0f} Ω", size=8)
            for ci, cell in enumerate(row.cells):
                cell.width = col_w_dt[ci]
            _set_row_height_cm(row, 0.50)

        def _drow(main: str, sub: str, vals: list,
                  bg: Optional[RGBColor] = None):
            row = dtbl.add_row()
            _data_cell(row.cells[0], main, size=7.5, bold=(bg is not None),
                       align=WD_ALIGN_PARAGRAPH.LEFT, bg=bg)
            _data_cell(row.cells[1], sub,  size=7.5,
                       align=WD_ALIGN_PARAGRAPH.LEFT, bg=bg)
            for j, val in enumerate(vals):
                txt = (_fmt(val, 4) if isinstance(val, float)
                       else str(val) if val is not None else "-")
                _data_cell(row.cells[2 + j], txt, size=7.5,
                           align=WD_ALIGN_PARAGRAPH.RIGHT, bg=bg)
            for ci, cell in enumerate(row.cells):
                cell.width = col_w_dt[ci]
            _set_row_height_cm(row, 0.48)

        def _srow(main: str, detail: str,
                  bg: Optional[RGBColor] = None):
            """Row whose value spans all right columns."""
            row = dtbl.add_row()
            _data_cell(row.cells[0], main, size=7.5, bold=True,
                       align=WD_ALIGN_PARAGRAPH.LEFT, bg=bg)
            row.cells[0].width = col_w_dt[0]
            merged = row.cells[1]
            for k in range(2, n_dt):
                merged = merged.merge(row.cells[k])
            _data_cell(merged, detail, size=7.5,
                       align=WD_ALIGN_PARAGRAPH.LEFT, bg=bg)
            _set_row_height_cm(row, 0.48)

        _hrow()
        _drow("1) Decimal",
              "AVG", [int(cal.decimals_avg.get(r, 0)) for r in rs],
              bg=C_THEAD_BG)
        _drow("", "MIN", [int(cal.decimals_min.get(r, 0)) for r in rs])
        _drow("", "MAX", [int(cal.decimals_max.get(r, 0)) for r in rs])

        _drow("2) 전압 (before gain)",
              "Ref", [cal.voltage_ref.get(r) for r in rs], bg=C_THEAD_BG)
        _drow("", "AVG", [cal.voltages_avg.get(r) for r in rs])

        _drow("3) 저항 (before gain)",
              "Ref", [r for r in rs], bg=C_THEAD_BG)
        _drow("", "AVG",  [cal.r_before_gain.get(r)  for r in rs])
        _drow("", "편차", [cal.dev_before_gain.get(r) for r in rs])

        _srow("4) G_cal", f"{cal.gain:.8f}", bg=C_THEAD_BG)

        _drow("5) 저항 (after gain)",
              "AVG",  [cal.r_after_gain_avg.get(r) for r in rs], bg=C_THEAD_BG)
        _drow("", "MIN",  [cal.r_after_gain_min.get(r) for r in rs])
        _drow("", "MAX",  [cal.r_after_gain_max.get(r) for r in rs])
        _drow("", "편차", [cal.dev_after_gain.get(r)   for r in rs])

        _srow("6) V_offset",
              f"2-1: {cal.offset_100_v:.6f} V    /    "
              f"2-2: {cal.offset_mean_v:.6f} V",
              bg=C_THEAD_BG)

        _drow("7) 2-1 최종 (100Ω offset)",
              "AVG",     [cal.r_final_100_avg.get(r) for r in rs],
              bg=C_THEAD_BG)
        _drow("", "편차",    [cal.dev_final_100.get(r)     for r in rs])
        _drow("", "허용 (+)", [cal.tolerance_max_100.get(r) for r in rs])
        _drow("", "허용 (−)", [cal.tolerance_min_100.get(r) for r in rs])

        _drow("8) 2-2 최종 (Mean offset)",
              "AVG",     [cal.r_final_mean_avg.get(r) for r in rs],
              bg=C_THEAD_BG)
        _drow("", "편차",    [cal.dev_final_mean.get(r)     for r in rs])
        _drow("", "허용 (+)", [cal.tolerance_max_mean.get(r) for r in rs])
        _drow("", "허용 (−)", [cal.tolerance_min_mean.get(r) for r in rs])

    # ── Public API ────────────────────────────────────────────────────────────

    def write(self, output_path: str) -> str:
        """
        Generate the full calibration DOCX report.

        Workflow:
          1. Open template (preserves header, logo, fonts, page setup, footer)
          2. Replace info field values in rows 0-11
          3. Remove rows 12+ (Device information section and below)
          4. Append calibration sections (setup, summary, per-channel pages)
        """
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # ── Load template ──────────────────────────────────────────────────
        doc = Document(TEMPLATE_PATH)
        tbl = doc.tables[0]   # the single master table in the template

        # ── Update info fields ─────────────────────────────────────────────
        self._update_template_fields(tbl)

        # ── Remove Device information and below (rows 12+) ─────────────────
        self._trim_table(tbl, keep_rows=12)

        # ── Append calibration content ─────────────────────────────────────
        self._add_setup_section(doc)
        self._add_summary(doc)

        for idx, ch in enumerate(self.channels):
            self._add_channel_page(doc, idx, ch)

        doc.save(output_path)
        return output_path
