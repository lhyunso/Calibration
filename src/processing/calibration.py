"""
Core calibration math for Quarter Bridge measurement modules.

Calibration flow:
1. Extract AVG/MIN/MAX decimals from CSV for each resistance step
2. Convert decimals → voltage: V = (D - 32767) * (20/65536)
3. Reference voltage: V_ref = (R - R_nom) / (2*R_nom)
4. Gain: slope of measured voltage vs reference voltage (linear regression over all steps)
5. R_new with gain applied: R = (2*R_nom * V / gain) + R_nom
6. Offset Method 2-1 (100Ω basis): offset = deviation at nominal R
   Offset Method 2-2 (mean basis):  offset = mean of all deviations
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from processing.csv_reader import ResistanceDataset, ChannelStats


@dataclass
class ChannelCalibration:
    """Full calibration result for one channel."""
    channel: str
    r_nominal: float               # Nominal resistance (e.g., 100 for PT100)
    excitation: float              # Excitation factor (e.g., 0.001 V/Ω for PT100)
    inst_amp_gain: float           # Instrument amplifier gain (default 1)

    # Computed
    gain: float = 0.0

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

    # Offsets
    offset_100: float = 0.0        # Method 2-1: offset at nominal R
    offset_mean: float = 0.0       # Method 2-2: mean of deviations

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


def _voltage_to_resistance(voltage: float, gain: float, r_nominal: float) -> float:
    """Convert measured voltage to resistance using calibrated gain."""
    if gain == 0:
        return r_nominal
    return (2 * r_nominal * voltage / gain) + r_nominal


def _reference_voltage(r: float, r_nominal: float, inst_amp_gain: float = 1.0) -> float:
    """Calculate reference (ideal) voltage for a given resistance."""
    return (r - r_nominal) / (2 * r_nominal) * inst_amp_gain


def calibrate_channel(
    channel: str,
    datasets: Dict[float, ResistanceDataset],
    r_nominal: float,
    excitation: float,
    inst_amp_gain: float = 1.0,
    tolerance: float = 0.385,
) -> ChannelCalibration:
    """
    Run full calibration for a single channel across all resistance steps.

    datasets: keyed by resistance value (e.g., 80, 90, 100, 110, 120)
    """
    resistances = sorted(datasets.keys())
    cal = ChannelCalibration(
        channel=channel,
        r_nominal=r_nominal,
        excitation=excitation,
        inst_amp_gain=inst_amp_gain,
    )

    from config import ADC_CENTER, ADC_VOLTAGE_RANGE, ADC_FULL_RANGE

    # Step 1: Collect decimals and convert to voltages
    for r in resistances:
        stats: Optional[ChannelStats] = datasets[r].channel_stats.get(channel)
        if stats is None:
            continue

        # Excel uses INT(AVERAGE()) — truncate to integer before voltage conversion
        d_avg = int(stats.avg)
        d_min = int(stats.min_val)
        d_max = int(stats.max_val)

        cal.decimals_avg[r] = d_avg
        cal.decimals_min[r] = d_min
        cal.decimals_max[r] = d_max

        cal.voltages_avg[r] = (d_avg - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)
        cal.voltages_min[r] = (d_min - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)
        cal.voltages_max[r] = (d_max - ADC_CENTER) * (ADC_VOLTAGE_RANGE / ADC_FULL_RANGE)
        cal.voltage_ref[r] = _reference_voltage(r, r_nominal, inst_amp_gain)

    # Step 2: Gain from slope of measured voltage vs reference voltage
    r_list = [r for r in resistances if r in cal.voltages_avg and r in cal.voltage_ref]
    if len(r_list) >= 2:
        r_min_r = min(r_list)
        r_max_r = max(r_list)
        v_meas_range = cal.voltages_avg[r_max_r] - cal.voltages_avg[r_min_r]
        v_ref_range = cal.voltage_ref[r_max_r] - cal.voltage_ref[r_min_r]
        cal.gain = (v_meas_range / v_ref_range) * inst_amp_gain if v_ref_range != 0 else inst_amp_gain

    # Step 3: Pre-gain resistance and deviation
    for r in r_list:
        cal.r_before_gain[r] = _voltage_to_resistance(
            cal.voltages_avg[r], inst_amp_gain, r_nominal
        )
        cal.dev_before_gain[r] = cal.r_before_gain[r] - r

    # Step 4: Post-gain resistance
    for r in r_list:
        cal.r_after_gain_avg[r] = _voltage_to_resistance(cal.voltages_avg[r], cal.gain, r_nominal)
        cal.r_after_gain_min[r] = _voltage_to_resistance(cal.voltages_min[r], cal.gain, r_nominal)
        cal.r_after_gain_max[r] = _voltage_to_resistance(cal.voltages_max[r], cal.gain, r_nominal)
        cal.dev_after_gain[r] = cal.r_after_gain_avg[r] - r

    # Step 5: Offsets
    cal.offset_100 = cal.dev_after_gain.get(r_nominal, 0.0)
    if cal.dev_after_gain:
        cal.offset_mean = sum(cal.dev_after_gain.values()) / len(cal.dev_after_gain)

    # Step 6: Method 2-1 final (offset by nominal R)
    for r in r_list:
        cal.r_final_100_avg[r] = cal.r_after_gain_avg[r] - cal.offset_100
        cal.dev_final_100[r] = cal.r_final_100_avg[r] - r
        cal.tolerance_max_100[r] = tolerance
        cal.tolerance_min_100[r] = -tolerance

    # Step 7: Method 2-2 final (offset by mean)
    for r in r_list:
        cal.r_final_mean_avg[r] = cal.r_after_gain_avg[r] - cal.offset_mean
        cal.dev_final_mean[r] = cal.r_final_mean_avg[r] - r
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
) -> Dict[str, ChannelCalibration]:
    """
    Calibrate all 16 channels.
    Returns dict keyed by channel name (CH01 … CH16).
    """
    from config import CHANNEL_NAMES
    if channels is None:
        channels = CHANNEL_NAMES

    # Only use channels present in at least one dataset
    available = set()
    for ds in datasets.values():
        available.update(ds.channel_stats.keys())
    channels = [ch for ch in channels if ch in available]

    results = {}
    for ch in channels:
        results[ch] = calibrate_channel(
            ch, datasets, r_nominal, excitation, inst_amp_gain, tolerance
        )
    return results
