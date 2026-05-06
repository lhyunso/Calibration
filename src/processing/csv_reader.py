"""
CSV data reader for calibration measurement files.

Channels are mapped by column order, not by header name:
  column 0 = timestamp (skipped)
  column 1 = CH01, column 2 = CH02, … column 16 = CH16

Supports two reading modes (configured in config.py or overridden per call):
  "auto"   — read all data rows in the file
  "last_n" — use only the final (sampling_hz × use_last_seconds) rows
"""
import os
import csv
import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional
import statistics

from config import (
    CHANNEL_COUNT, CHANNEL_NAMES,
    DATA_MODE, SAMPLING_HZ, USE_LAST_SECONDS,
    ADC_CENTER, ADC_VOLTAGE_RANGE, ADC_FULL_RANGE,
)


# ── Data containers ──────────────────────────────────────────────────────────

@dataclass
class ChannelStats:
    """Statistics for a single channel at one resistance step."""
    channel: str
    resistance: float
    avg: float
    min_val: float
    max_val: float
    sample_count: int
    duration_seconds: float

    @property
    def voltage_avg(self) -> float:
        return (int(self.avg) - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)

    @property
    def voltage_min(self) -> float:
        return (int(self.min_val) - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)

    @property
    def voltage_max(self) -> float:
        return (int(self.max_val) - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)


@dataclass
class ResistanceDataset:
    """All channel data for one resistance setting."""
    resistance: float
    sample_count: int = 0
    duration_seconds: float = 0.0
    channel_stats: Dict[str, ChannelStats] = field(default_factory=dict)


# ── Core reader ──────────────────────────────────────────────────────────────

def read_csv_file(
    filepath: str,
    resistance: float,
    data_mode: str = DATA_MODE,
    sampling_hz: int = SAMPLING_HZ,
    use_last_seconds: int = USE_LAST_SECONDS,
) -> ResistanceDataset:
    """
    Read a calibration CSV file and return per-channel statistics.

    The first column is treated as a timestamp and skipped.
    Columns 1…CHANNEL_COUNT are assigned to CH01…CH16 in order.
    Header rows are detected automatically and skipped.
    """
    all_rows: List[List[float]] = []
    t_values: List[float] = []

    with open(filepath, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        for raw in reader:
            row = [c.strip() for c in raw]
            if not row or not row[0]:
                continue
            # Skip any row whose first cell is not numeric (header lines)
            try:
                t = float(row[0])
            except ValueError:
                continue

            try:
                values = []
                for cell in row:
                    try:
                        values.append(float(cell))
                    except ValueError:
                        values.append(math.nan)
                all_rows.append(values)
                t_values.append(t)
            except Exception:
                continue

    if not all_rows:
        return ResistanceDataset(resistance=resistance)

    # ── Apply data-mode windowing ────────────────────────────────────────────
    if data_mode == "last_n":
        n = max(1, sampling_hz * use_last_seconds)
        rows = all_rows[-n:]
        ts   = t_values[-n:]
    else:
        rows = all_rows
        ts   = t_values

    sample_count = len(rows)

    # Estimate actual duration from timestamps when available
    if len(ts) >= 2:
        dt = 1.0 / sampling_hz if sampling_hz > 0 else 0.0
        duration = round(ts[-1] - ts[0] + dt, 3)
    else:
        duration = round(sample_count / sampling_hz, 3) if sampling_hz > 0 else 0.0

    dataset = ResistanceDataset(
        resistance=resistance,
        sample_count=sample_count,
        duration_seconds=duration,
    )

    # ── Compute per-channel statistics ───────────────────────────────────────
    for ch_idx, ch_name in enumerate(CHANNEL_NAMES):
        col = ch_idx + 1          # column 0 is the timestamp
        data: List[float] = []
        for row in rows:
            if col < len(row) and not math.isnan(row[col]):
                data.append(row[col])
        if not data:
            continue

        dataset.channel_stats[ch_name] = ChannelStats(
            channel=ch_name,
            resistance=resistance,
            avg=statistics.mean(data),
            min_val=min(data),
            max_val=max(data),
            sample_count=len(data),
            duration_seconds=duration,
        )

    return dataset


# ── Batch loader ─────────────────────────────────────────────────────────────

def load_datasets(
    csv_files: Dict[float, str],
    data_mode: str = DATA_MODE,
    sampling_hz: int = SAMPLING_HZ,
    use_last_seconds: int = USE_LAST_SECONDS,
) -> Dict[float, ResistanceDataset]:
    """Load multiple CSV files keyed by resistance value (Ω)."""
    datasets: Dict[float, ResistanceDataset] = {}
    for resistance, filepath in sorted(csv_files.items()):
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"CSV not found: {filepath}")
        datasets[resistance] = read_csv_file(
            filepath, resistance, data_mode, sampling_hz, use_last_seconds
        )
    return datasets


# ── Auto-detection ───────────────────────────────────────────────────────────

def auto_detect_csv_files(directory: str) -> Dict[float, str]:
    """
    Scan a directory for calibration CSV files whose names contain a resistance
    value (e.g. 'data_80.csv', 'Calibration_110.csv').
    Returns {resistance_ohm: filepath}, sorted by resistance.
    """
    result: Dict[float, str] = {}
    for fname in sorted(os.listdir(directory)):
        if not fname.lower().endswith(".csv"):
            continue
        stem = os.path.splitext(fname)[0]
        parts = stem.replace("-", "_").split("_")
        for part in reversed(parts):
            try:
                val = float(part)
                if 0 < val < 10_000:
                    result[val] = os.path.join(directory, fname)
                    break
            except ValueError:
                continue
    return result
