#!/usr/bin/env python3
"""
Multi-Channel Sensor Calibration Tool (MCAL)
=============================================
CLI entry point.

Menus:
  1. PT100 캘리브레이션
  2. PT1000 캘리브레이션
  3. Strain 350Ω 캘리브레이션
  4. 3-Wire 레퍼런스 단일값 조회
  5. 3-Wire 레퍼런스 테이블 출력
  6. 3-Wire 레퍼런스 역산
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from datetime import datetime
from typing import Dict, Optional

from rich.console import Console
from rich.table import Table as RichTable
from rich.panel import Panel
from rich.prompt import Prompt, Confirm
from rich import box

from config import (
    REFERENCE_DIR, OUTPUT_XLSX, OUTPUT_DOCX, OUTPUT_PDF,
    DATA_MODE, SAMPLING_HZ, USE_LAST_SECONDS,
)
from sensors.pt100 import SENSOR_REGISTRY, get_sensor
from processing.csv_reader import load_datasets, auto_detect_csv_files
from processing.calibration import calibrate_all_channels
from output.xlsx_writer import CalibrationXlsxWriter
from output.docx_writer import CalibrationDocxWriter
from output.pdf_writer import CalibrationPdfWriter
from reference.three_wire import (
    build_reference_table, lookup_single, STANDARD_GAINS,
    find_resistance_from_voltage,
)

console = Console()


# ── Helpers ────────────────────────────────────────────────────────────────────

def _clear():
    os.system("clear" if os.name == "posix" else "cls")


def _banner():
    console.print(Panel(
        "[bold cyan]Multi-Channel Sensor Calibration Tool[/bold cyan]  [dim](MCAL)[/dim]\n"
        "[dim]PT100 / PT1000 / Strain 350Ω  |  CH01 ~ CH16[/dim]",
        expand=False,
        border_style="cyan",
    ))


def _pause():
    Prompt.ask("\n[dim]Press Enter to continue[/dim]", default="")


def _input_float(prompt: str, default: Optional[float] = None) -> float:
    while True:
        raw = Prompt.ask(prompt, default=str(default) if default is not None else "")
        try:
            return float(raw)
        except ValueError:
            console.print("[red]숫자를 입력하세요.[/red]")


def _input_int(prompt: str, default: Optional[int] = None) -> int:
    while True:
        raw = Prompt.ask(prompt, default=str(default) if default is not None else "")
        try:
            return int(raw)
        except ValueError:
            console.print("[red]정수를 입력하세요.[/red]")


# ── Metadata collection ────────────────────────────────────────────────────────

def _collect_metadata(sensor_name: str) -> Dict[str, str]:
    console.print("\n[bold]교정 정보 입력[/bold] [dim](Enter → 기본값 사용)[/dim]\n")

    now = datetime.now()

    # Section 1: Module Info
    console.print("[bold cyan][ 모듈 정보 ][/bold cyan]")
    meta: Dict[str, str] = {
        "module_name":   Prompt.ask("  모듈명 (Module Name)",      default=""),
        "model":         Prompt.ask("  모델명 (Model No.)",         default=""),
        "serial":        Prompt.ask("  시리얼 번호 (S/N)",           default=""),
        "manufacturer":  Prompt.ask("  제조사 (Manufacturer)",       default=""),
        "fw_version":    Prompt.ask("  FW / SW 버전",               default=""),
    }

    # Section 2: Calibration Conditions
    console.print("\n[bold cyan][ 교정 조건 ][/bold cyan]")
    meta.update({
        "date":          Prompt.ask("  교정 일자 (Date)",            default=now.strftime("%Y.%m.%d")),
        "location":      Prompt.ask("  교정 장소 (Location)",        default=""),
        "operator":      Prompt.ask("  담당자 (Technician)",         default=""),
        "temp_humidity": Prompt.ask("  온도 / 습도 (예: 23°C / 50%)", default=""),
    })

    # Section 3: Calibration Setup
    console.print("\n[bold cyan][ 교정 설정 ][/bold cyan]")
    meta.update({
        "cable":         Prompt.ask("  케이블 (Cable)",              default="전용케이블"),
        "inst_amp_gain": Prompt.ask("  Inst. Amp. Gain",             default="1"),
        "doc_number":    Prompt.ask("  문서번호 (Doc. No.)",          default=f"CAL-{now.strftime('%Y')}-0001"),
        "revision":      Prompt.ask("  Rev",                         default="00"),
    })

    # Data mode
    console.print("\n[bold cyan][ 데이터 설정 ][/bold cyan]")
    console.print(f"  데이터 읽기 모드: [cyan]1[/cyan]=전체 자동 (기본)  [cyan]2[/cyan]=끝부분 지정")
    mode_choice = Prompt.ask("  선택", default="1")
    if mode_choice == "2":
        hz  = _input_int("  샘플링 속도 (Hz)", default=SAMPLING_HZ)
        sec = _input_int("  사용할 마지막 구간 (초)", default=USE_LAST_SECONDS)
        meta["data_mode"]        = "last_n"
        meta["sampling_hz"]      = str(hz)
        meta["use_last_seconds"] = str(sec)
    else:
        meta["data_mode"]   = "auto"
        meta["sampling_hz"] = str(SAMPLING_HZ)

    return meta


# ── CSV file selection ─────────────────────────────────────────────────────────

def _select_csv_files(sensor) -> Dict[float, str]:
    console.print("\n[bold]CSV 파일 선택[/bold]")
    console.print(f"기본 모사저항: {sensor.default_resistances}")
    console.print(f"레퍼런스 폴더: [cyan]{REFERENCE_DIR}[/cyan]\n")

    auto = auto_detect_csv_files(REFERENCE_DIR)
    if auto:
        console.print("[green]자동 감지된 CSV 파일:[/green]")
        t = RichTable(box=box.SIMPLE)
        t.add_column("저항(Ω)", style="cyan")
        t.add_column("파일명")
        for r_val in sorted(auto.keys()):
            t.add_row(str(r_val), os.path.basename(auto[r_val]))
        console.print(t)
        if Confirm.ask("위 파일을 사용하시겠습니까?", default=True):
            return auto

    # Manual
    resistances_str = Prompt.ask(
        "모사저항 값 입력 (쉼표 구분)",
        default=",".join(str(int(r)) for r in sensor.default_resistances),
    )
    resistances = [float(r.strip()) for r in resistances_str.split(",")]
    csv_files: Dict[float, str] = {}
    for r_val in resistances:
        while True:
            path = Prompt.ask(f"  {r_val:.0f}Ω CSV 파일 경로")
            if os.path.exists(path):
                csv_files[r_val] = path
                break
            console.print(f"  [red]파일 없음: {path}[/red]")
    return csv_files


# ── Calibration summary display ────────────────────────────────────────────────

def _show_summary(calibrations, sensor):
    if not calibrations:
        return
    first    = next(iter(calibrations.values()))
    rs       = sorted(first.voltages_avg.keys())
    tol      = sensor.tolerance_ohm

    t = RichTable(
        title="Calibration Summary — Method 2-1 (R_nom Offset)",
        box=box.ROUNDED,
    )
    t.add_column("CH",     style="cyan", no_wrap=True)
    t.add_column("Gain",   justify="right")
    t.add_column("Offset", justify="right")
    for r_val in rs:
        t.add_column(f"{r_val:.0f}Ω dev", justify="right")
    t.add_column("판정", justify="center")

    for ch, cal in calibrations.items():
        devs     = [cal.dev_final_100.get(r_val) for r_val in rs]
        pass_all = all(d is not None and abs(d) <= tol for d in devs)
        color    = "green" if pass_all else "red"
        dev_strs = []
        for d in devs:
            if d is None:
                dev_strs.append("-")
            else:
                ok = abs(d) <= tol
                dev_strs.append(f"[{'green' if ok else 'red'}]{d:.4f}[/]")
        t.add_row(
            f"[{color}]{ch}[/]",
            f"{cal.gain:.6f}",
            f"{cal.offset_100:.6f}",
            *dev_strs,
            f"[{color}]{'PASS' if pass_all else 'FAIL'}[/]",
        )
    console.print()
    console.print(t)


# ── Export ─────────────────────────────────────────────────────────────────────

def _export_reports(sensor, calibrations, meta, datasets):
    ts        = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"{sensor.sensor_type}_calibration_{ts}"

    # Enrich meta with actual duration info
    if datasets:
        durations = [ds.duration_seconds for ds in datasets.values() if ds.duration_seconds > 0]
        if durations:
            meta["duration_sec"] = f"{max(durations):.1f}"

    # Ask which formats to export
    console.print("\n[bold]출력 형식 선택[/bold]")
    do_xlsx = Confirm.ask("  xlsx 출력?", default=True)
    do_docx = Confirm.ask("  docx 출력?", default=True)
    do_pdf  = Confirm.ask("  pdf  출력?", default=False)

    if do_xlsx:
        path = os.path.join(OUTPUT_XLSX, f"{base_name}.xlsx")
        os.makedirs(OUTPUT_XLSX, exist_ok=True)
        try:
            CalibrationXlsxWriter(sensor, calibrations, meta).write(path)
            console.print(f"[green]xlsx 저장: {path}[/green]")
        except Exception as e:
            console.print(f"[red]xlsx 오류: {e}[/red]")
            import traceback; traceback.print_exc()

    if do_docx:
        path = os.path.join(OUTPUT_DOCX, f"{base_name}.docx")
        os.makedirs(OUTPUT_DOCX, exist_ok=True)
        try:
            CalibrationDocxWriter(sensor, calibrations, meta).write(path)
            console.print(f"[green]docx 저장: {path}[/green]")
        except Exception as e:
            console.print(f"[red]docx 오류: {e}[/red]")
            import traceback; traceback.print_exc()

    if do_pdf:
        path = os.path.join(OUTPUT_PDF, f"{base_name}.pdf")
        os.makedirs(OUTPUT_PDF, exist_ok=True)
        try:
            CalibrationPdfWriter(sensor, calibrations, meta).write(path)
            console.print(f"[green]pdf  저장: {path}[/green]")
        except Exception as e:
            console.print(f"[red]pdf  오류: {e}[/red]")
            import traceback; traceback.print_exc()


# ── Calibration workflow ───────────────────────────────────────────────────────

def run_calibration(sensor_type: str):
    sensor = get_sensor(sensor_type)
    console.print(Panel(f"[bold]{sensor.name} 캘리브레이션[/bold]", border_style="green"))

    # Step 1: CSV
    csv_files = _select_csv_files(sensor)
    console.print(f"\n[green]총 {len(csv_files)}개 저항 데이터 로드 중...[/green]")
    try:
        meta_preview = {"data_mode": DATA_MODE}
        datasets = load_datasets(
            csv_files,
            data_mode=DATA_MODE,
            sampling_hz=SAMPLING_HZ,
            use_last_seconds=USE_LAST_SECONDS,
        )
    except Exception as e:
        console.print(f"[red]CSV 로드 오류: {e}[/red]")
        _pause(); return

    # Show actual sample counts
    for r_val, ds in sorted(datasets.items()):
        console.print(
            f"  {r_val:.0f}Ω : {ds.sample_count:,} 샘플  ({ds.duration_seconds:.1f}초)"
        )

    # Step 2: Metadata
    meta = _collect_metadata(sensor.name)

    # Reload with actual data_mode from metadata
    data_mode        = meta.get("data_mode", DATA_MODE)
    sampling_hz      = int(meta.get("sampling_hz", SAMPLING_HZ))
    use_last_seconds = int(meta.get("use_last_seconds", USE_LAST_SECONDS))

    if data_mode != DATA_MODE or data_mode == "last_n":
        console.print(f"\n[green]데이터 모드 [{data_mode}] 로 재로드 중...[/green]")
        try:
            datasets = load_datasets(csv_files, data_mode, sampling_hz, use_last_seconds)
        except Exception as e:
            console.print(f"[red]재로드 오류: {e}[/red]")
            _pause(); return

    # Step 3: Calibrate
    console.print("\n[bold]캘리브레이션 계산 중...[/bold]")
    try:
        calibrations = calibrate_all_channels(
            datasets,
            r_nominal=sensor.r_nominal,
            excitation=sensor.excitation,
            inst_amp_gain=float(meta.get("inst_amp_gain", "1")),
            tolerance=sensor.tolerance_ohm,
        )
    except Exception as e:
        console.print(f"[red]계산 오류: {e}[/red]")
        import traceback; traceback.print_exc()
        _pause(); return

    console.print(f"[green]계산 완료: {len(calibrations)}개 채널[/green]")

    # Step 4: Summary display
    _show_summary(calibrations, sensor)

    # Step 5: Export
    if Confirm.ask("\n결과를 파일로 저장하시겠습니까?", default=True):
        _export_reports(sensor, calibrations, meta, datasets)


# ── 3-Wire Reference menus ─────────────────────────────────────────────────────

SENSOR_DISPLAY = {
    "pt100":     "PT100   (R_nom=100Ω,  V=(R-100)/200 × Gain)",
    "pt1000":    "PT1000  (R_nom=1000Ω, V=(R-1000)/2000 × Gain)",
    "strain350": "Strain 350Ω (V=(3.5/2)×(ΔR/350) × Gain)",
}


def _select_sensor_type() -> str:
    console.print("\n[bold]센서 타입 선택:[/bold]")
    keys = list(SENSOR_DISPLAY.keys())
    for i, k in enumerate(keys, 1):
        console.print(f"  {i}. {SENSOR_DISPLAY[k]}")
    while True:
        raw = Prompt.ask("번호 입력", default="1")
        try:
            idx = int(raw) - 1
            if 0 <= idx < len(keys):
                return keys[idx]
        except ValueError:
            pass
        console.print("[red]잘못된 입력[/red]")


def run_reference_lookup():
    console.print(Panel("[bold]3-Wire 레퍼런스 전압 조회[/bold]", border_style="blue"))
    sensor_type = _select_sensor_type()
    while True:
        r_val = _input_float("저항값 R (Ω)", default=100.0)
        gain  = _input_float("Gain",          default=1.0)
        v_ref = lookup_single(sensor_type, r_val, gain)
        console.print(
            f"\n  [cyan]{SENSOR_DISPLAY[sensor_type]}[/cyan]\n"
            f"  R = {r_val} Ω,  Gain = {gain}\n"
            f"  [bold green]Reference Voltage = {v_ref:.6f} V[/bold green]\n"
        )
        if not Confirm.ask("다른 값 조회?", default=True):
            break


def run_reference_table():
    console.print(Panel("[bold]3-Wire 레퍼런스 테이블 출력[/bold]", border_style="blue"))
    sensor_type = _select_sensor_type()
    sensor      = get_sensor(sensor_type)

    from reference.three_wire import _SENSOR_RANGES
    default_range = _SENSOR_RANGES.get(sensor_type, (0, 100))

    r_start = _input_float("R 시작값 (Ω)", default=float(default_range[0]))
    r_end   = _input_float("R 종료값 (Ω)", default=float(default_range[1]))
    r_step  = _input_float("R 간격 (Ω)",   default=1.0)

    console.print(f"\n  표준 Gain: {STANDARD_GAINS}")
    gains_raw = Prompt.ask("사용할 Gain 값 (쉼표 구분)", default="1,10,100")
    try:
        gains = [float(g.strip()) for g in gains_raw.split(",")]
    except ValueError:
        gains = [1.0, 10.0, 100.0]

    resistances, gain_list, table = build_reference_table(
        sensor_type, r_start, r_end, r_step, gains
    )

    t = RichTable(
        title=f"3-Wire Reference Voltage  [{sensor.name}]",
        box=box.ROUNDED,
    )
    t.add_column("R (Ω)",  style="cyan", justify="right")
    t.add_column("ΔR (Ω)", justify="right")
    for g in gain_list:
        t.add_column(f"Gain={g:.0f}", justify="right")

    for r_val in resistances:
        delta = r_val - sensor.r_nominal
        vs    = [f"{v:.5f}" for v in table[r_val]]
        t.add_row(f"{r_val:.1f}", f"{delta:+.1f}", *vs)
    console.print(t)

    if Confirm.ask("\n결과를 CSV로 저장하시겠습니까?", default=False):
        import csv as _csv
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(OUTPUT_PDF, f"3wire_ref_{sensor_type}_{ts}.csv")
        os.makedirs(OUTPUT_PDF, exist_ok=True)
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = _csv.writer(f)
            w.writerow(["R (Ω)", "ΔR (Ω)"] + [f"Gain={g:.0f}" for g in gain_list])
            for r_val in resistances:
                delta = r_val - sensor.r_nominal
                w.writerow([f"{r_val:.1f}", f"{delta:+.1f}"] + [f"{v:.6f}" for v in table[r_val]])
        console.print(f"[green]CSV 저장: {path}[/green]")


def run_reverse_lookup():
    console.print(Panel("[bold]전압 → 저항 역산[/bold]", border_style="blue"))
    sensor_type = _select_sensor_type()
    while True:
        v    = _input_float("레퍼런스 전압 V (V)")
        gain = _input_float("Gain", default=1.0)
        r_val = find_resistance_from_voltage(sensor_type, v, gain)
        console.print(
            f"\n  V = {v:.6f} V,  Gain = {gain}\n"
            f"  [bold green]R ≈ {r_val:.4f} Ω[/bold green]\n"
        )
        if not Confirm.ask("다른 값 역산?", default=True):
            break


# ── Main menu ──────────────────────────────────────────────────────────────────

def main():
    while True:
        _clear()
        _banner()
        console.print()
        console.print("[bold]메인 메뉴[/bold]")
        console.print()
        console.print("  [cyan]1.[/cyan]  PT100    캘리브레이션  (CSV → xlsx / docx / pdf)")
        console.print("  [cyan]2.[/cyan]  PT1000   캘리브레이션  (CSV → xlsx / docx / pdf)")
        console.print("  [cyan]3.[/cyan]  Strain 350Ω 캘리브레이션  (CSV → xlsx / docx / pdf)")
        console.print()
        console.print("  [blue]4.[/blue]  3-Wire 레퍼런스 단일값 조회  (R, Gain → V)")
        console.print("  [blue]5.[/blue]  3-Wire 레퍼런스 테이블 출력  (범위 조회 + CSV 저장)")
        console.print("  [blue]6.[/blue]  3-Wire 레퍼런스 역산  (V, Gain → R)")
        console.print()
        console.print("  [dim]0.  종료[/dim]")
        console.print()

        choice = Prompt.ask("선택", default="1")
        if choice == "0":
            console.print("[dim]종료합니다.[/dim]")
            break
        elif choice == "1":
            run_calibration("pt100");    _pause()
        elif choice == "2":
            run_calibration("pt1000");   _pause()
        elif choice == "3":
            run_calibration("strain350"); _pause()
        elif choice == "4":
            run_reference_lookup();       _pause()
        elif choice == "5":
            run_reference_table();        _pause()
        elif choice == "6":
            run_reverse_lookup();         _pause()
        else:
            console.print("[red]잘못된 입력입니다.[/red]")
            _pause()


if __name__ == "__main__":
    main()
