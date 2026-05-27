"""
Core calibration math for Quarter Bridge measurement modules.

Calibration flow:
1. Extract AVG/MIN/MAX decimals from CSV for each resistance step
2. Convert decimals → voltage: V = (int(D) - 32767) × (20/65536)
3. Reference voltage (normalized, no G_inst):
     V_ref = (R − R_nom) / (2×R_nom)
4. G_cal: slope of measured voltage vs reference voltage
     G_cal = (V_meas_max − V_meas_min) / (V_ref_max − V_ref_min)
5. Resistance from voltage:
     R = 2×R_nom×V / G_cal + R_nom
6. V_offset (voltage domain, for ground software: V_meas/G_cal + V_offset):
     Method 2-1: residual at R_nom
     Method 2-2: mean of all residuals
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from processing.csv_reader import ResistanceDataset, ChannelStats


@dataclass
class ChannelCalibration:
    """Full calibration result for one channel."""
    channel: str
    r_nominal: float               # Nominal resistance (e.g., 100 for PT100)
    excitation: float              # Excitation current [A] — reference/display only
    inst_amp_gain: float           # Instrument amplifier gain — reference/display only

    # Computed
    gain: float = 0.0              # G_cal: 지상SW 적용값

    # Per-resistance stats (keys are resistance values)
    decimals_avg: Dict[float, float] = field(default_factory=dict)
    decimals_min: Dict[float, float] = field(default_factory=dict)
    decimals_max: Dict[float, float] = field(default_factory=dict)

    voltages_avg: Dict[float, float] = field(default_factory=dict)
    voltages_min: Dict[float, float] = field(default_factory=dict)
    voltages_max: Dict[float, float] = field(default_factory=dict)
    voltage_ref: Dict[float, float] = field(default_factory=dict)

    # Pre-gain resistance
    r_before_gain: Dict[float, float] = field(default_factory=dict)
    dev_before_gain: Dict[float, float] = field(default_factory=dict)   # R_before - R_ref

    # Post-gain resistance (before offset)
    r_after_gain_avg: Dict[float, float] = field(default_factory=dict)
    r_after_gain_min: Dict[float, float] = field(default_factory=dict)
    r_after_gain_max: Dict[float, float] = field(default_factory=dict)
    dev_after_gain: Dict[float, float] = field(default_factory=dict)    # R_gain - R_ref

    # Offsets — resistance domain (internal use)
    offset_100: float = 0.0        # Method 2-1: offset at nominal R [Ω]
    offset_mean: float = 0.0       # Method 2-2: mean of deviations [Ω]

    # V_offset — voltage domain (지상SW 적용: V_meas/G_cal + V_offset)
    offset_100_v: float = 0.0      # Method 2-1: V_offset at nominal R [V]
    offset_mean_v: float = 0.0     # Method 2-2: mean V_offset [V]

    # Final results - Method 2-1 (offset by nominal R)
    r_final_100_avg: Dict[float, float] = field(default_factory=dict)
    dev_final_100: Dict[float, float] = field(default_factory=dict)
    tolerance_max_100: Dict[float, float] = field(default_factory=dict)
    tolerance_min_100: Dict[float, float] = field(default_factory=dict)

    # Final results - Method 2-2 (offset by mean)
    r_final_mean_avg: Dict[float, float] = field(default_factory=dict)
    dev_final_mean: Dict[float, float] = field(default_factory=dict)
    tolerance_max_mean: Dict[float, float] = field(default_factory=dict)
    tolerance_min_mean: Dict[float, float] = field(default_factory=dict)


def calibrate_channel(
    channel: str,
    datasets: Dict[float, ResistanceDataset],
    r_nominal: float,
    excitation: float,
    inst_amp_gain: float = 1.0,
    tolerance: float = 0.385,
    sensor=None,
) -> ChannelCalibration:
    """
    Run full calibration for a single channel across all resistance steps.

    datasets      : keyed by resistance value (e.g., 80, 90, 100, 110, 120)
    excitation    : reference only (displayed in report, not used in G_cal calc)
    inst_amp_gain : reference only (displayed in report, not used in G_cal calc)
    sensor        : SensorConfig instance — used for V_ref and R(V) formulas.
    """
    resistances = sorted(datasets.keys())
    cal = ChannelCalibration(
        channel=channel,
        r_nominal=r_nominal,
        excitation=excitation,
        inst_amp_gain=inst_amp_gain,
    )

    from config import ADC_CENTER, ADC_VOLTAGE_RANGE, ADC_FULL_RANGE

    # ── 센서별 공식 선택 ────────────────────────────────────────────────────
    def _ref_v(r: float) -> float:
        """Normalized reference voltage (no G_inst): (R - R_nom) / (2 × R_nom)"""
        if sensor is not None:
            # sensor.ref_voltage는 × G_inst 포함 — 여기선 G_inst=1로 호출
            return sensor.ref_voltage(r, 1.0)
        return (r - r_nominal) / (2 * r_nominal)

    def _r_from_v(v: float, g: float) -> float:
        """Resistance from measured voltage using calibrated gain."""
        if sensor is not None:
            return sensor.resistance_from_voltage(v, g)
        if g == 0:
            return r_nominal
        return (2 * r_nominal * v / g) + r_nominal

    # Step 1: Collect decimals and convert to voltages
    for r in resistances:
        stats: Optional[ChannelStats] = datasets[r].channel_stats.get(channel)
        if stats is None:
            continue

        d_avg = int(stats.avg)
        d_min = int(stats.min_val)
        d_max = int(stats.max_val)

        cal.decimals_avg[r] = d_avg
        cal.decimals_min[r] = d_min
        cal.decimals_max[r] = d_max

        cal.voltages_avg[r] = (d_avg - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)
        cal.voltages_min[r] = (d_min - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)
        cal.voltages_max[r] = (d_max - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)
        cal.voltage_ref[r]  = _ref_v(r)   # normalized, no G_inst

    # Step 2: G_cal — slope of V_meas vs V_ref (endpoints)
    # G_cal = (V_meas_max - V_meas_min) / (V_ref_max - V_ref_min)
    r_list = [r for r in resistances if r in cal.voltages_avg and r in cal.voltage_ref]
    if len(r_list) >= 2:
        r_min_r = min(r_list)
        r_max_r = max(r_list)
        v_meas_range = cal.voltages_avg[r_max_r] - cal.voltages_avg[r_min_r]
        v_ref_range  = cal.voltage_ref[r_max_r]  - cal.voltage_ref[r_min_r]
        cal.gain = (v_meas_range / v_ref_range) if v_ref_range != 0 else 1.0

    # Step 3: Pre-gain resistance
    for r in r_list:
        cal.r_before_gain[r]   = _r_from_v(cal.voltages_avg[r], 1.0)
        cal.dev_before_gain[r] = cal.r_before_gain[r] - r

    # Step 4: Post-gain resistance
    for r in r_list:
        cal.r_after_gain_avg[r] = _r_from_v(cal.voltages_avg[r], cal.gain)
        cal.r_after_gain_min[r] = _r_from_v(cal.voltages_min[r], cal.gain)
        cal.r_after_gain_max[r] = _r_from_v(cal.voltages_max[r], cal.gain)
        cal.dev_after_gain[r]   = cal.r_after_gain_avg[r] - r

    # Step 5: Resistance-domain offsets (내부 계산용)
    cal.offset_100 = cal.dev_after_gain.get(r_nominal, 0.0)
    if cal.dev_after_gain:
        cal.offset_mean = sum(cal.dev_after_gain.values()) / len(cal.dev_after_gain)

    # Step 5b: Voltage-domain offsets (지상SW 적용값: V_meas/G_cal + V_offset)
    # V_offset = V_ref_normalized(R) - V_meas(R)/G_cal
    # = (R - R_nom)/(2×R_nom) - V_meas(R)/G_cal
    # = -offset_resistance / (2 × R_nom)
    denom = 2.0 * r_nominal
    cal.offset_100_v  = -cal.offset_100  / denom if denom else 0.0
    cal.offset_mean_v = -cal.offset_mean / denom if denom else 0.0

    # Step 6: Method 2-1 final (offset by nominal R)
    for r in r_list:
        cal.r_final_100_avg[r] = cal.r_after_gain_avg[r] - cal.offset_100
        cal.dev_final_100[r]   = cal.r_final_100_avg[r] - r
        cal.tolerance_max_100[r] = tolerance
        cal.tolerance_min_100[r] = -tolerance

    # Step 7: Method 2-2 final (offset by mean)
    for r in r_list:
        cal.r_final_mean_avg[r] = cal.r_after_gain_avg[r] - cal.offset_mean
        cal.dev_final_mean[r]   = cal.r_final_mean_avg[r] - r
        cal.tolerance_max_mean[r] = tolerance
        cal.tolerance_min_mean[r] = -tolerance

    return cal


def calibrate_all_channels(
    datasets: Dict[float, ResistanceDataset],
    r_nominal: float,
    excitation: float,
    inst_amp_gain: float = 1.0,
    tolerance: float = 0.385,
    channels: Optional[List[str]] = None,
    sensor=None,
) -> Dict[str, ChannelCalibration]:
    """Calibrate all channels. Returns dict keyed by channel name (CH01…CH16)."""
    from config import CHANNEL_NAMES
    if channels is None:
        channels = CHANNEL_NAMES

    available = set()
    for ds in datasets.values():
        available.update(ds.channel_stats.keys())
    channels = [ch for ch in channels if ch in available]

    results = {}
    for ch in channels:
        results[ch] = calibrate_channel(
            ch, datasets, r_nominal, excitation, inst_amp_gain, tolerance,
            sensor=sensor,
        )
    return results
