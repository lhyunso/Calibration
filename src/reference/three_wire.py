"""
3-Wire reference voltage calculator.
Reproduces the reference tables from '3Wire(Reference값).xlsx'.

Formulas:
  PT100:    V = (0.1/2) * (ΔR/100) * gain   where ΔR = R - 100
  PT1000:   V = (0.1/2) * (ΔR/1000) * gain  where ΔR = R - 1000
  Strain350:V = (3.5/2) * (ΔR/350) * gain   where ΔR = R - 350
"""
from typing import Dict, List, Optional, Tuple
import math


# Standard gain values matching the reference Excel
STANDARD_GAINS = [1, 2, 4, 8, 10, 20, 40, 80, 100, 1000]


def ref_voltage_pt100(r: float, gain: float = 1.0) -> float:
    """Reference voltage for PT100: V = (0.1/2) * ((R-100)/100) * gain"""
    return (0.1 / 2) * ((r - 100) / 100) * gain


def ref_voltage_pt1000(r: float, gain: float = 1.0) -> float:
    """Reference voltage for PT1000: V = (0.1/2) * ((R-1000)/1000) * gain"""
    return (0.1 / 2) * ((r - 1000) / 1000) * gain


def ref_voltage_strain350(r: float, gain: float = 1.0) -> float:
    """Reference voltage for 350Ω strain gauge: V = (3.5/2) * ((R-350)/350) * gain"""
    return (3.5 / 2) * ((r - 350) / 350) * gain


# Sensor-type dispatch
_SENSOR_FUNCS = {
    "pt100": ref_voltage_pt100,
    "pt1000": ref_voltage_pt1000,
    "strain350": ref_voltage_strain350,
}

_SENSOR_RANGES = {
    "pt100":     (75, 125),
    "pt1000":    (750, 1250),
    "strain350": (320, 380),
}


def get_ref_voltage(sensor_type: str, r: float, gain: float = 1.0) -> float:
    """Calculate single reference voltage for given sensor type, resistance, and gain."""
    func = _SENSOR_FUNCS.get(sensor_type.lower())
    if func is None:
        raise ValueError(f"Unknown sensor type: {sensor_type!r}")
    return func(r, gain)


def build_reference_table(
    sensor_type: str,
    r_start: Optional[float] = None,
    r_end: Optional[float] = None,
    r_step: float = 1.0,
    gains: Optional[List[float]] = None,
) -> Tuple[List[float], List[float], Dict[float, List[float]]]:
    """
    Build a reference table for given sensor type and resistance range.

    Returns:
        (resistance_list, gain_list, table_dict)
        table_dict: {resistance: [v_gain1, v_gain2, ...]}
    """
    if gains is None:
        gains = STANDARD_GAINS

    default_range = _SENSOR_RANGES.get(sensor_type.lower(), (0, 100))
    if r_start is None:
        r_start = default_range[0]
    if r_end is None:
        r_end = default_range[1]

    func = _SENSOR_FUNCS.get(sensor_type.lower())
    if func is None:
        raise ValueError(f"Unknown sensor type: {sensor_type!r}")

    # Generate resistance list
    n_steps = int(round((r_end - r_start) / r_step)) + 1
    resistances = [r_start + i * r_step for i in range(n_steps)]

    table: Dict[float, List[float]] = {}
    for r in resistances:
        table[r] = [func(r, g) for g in gains]

    return resistances, gains, table


def lookup_single(
    sensor_type: str,
    r: float,
    gain: float,
) -> float:
    """Quick lookup: reference voltage for one sensor/resistance/gain combo."""
    return get_ref_voltage(sensor_type, r, gain)


def find_resistance_from_voltage(
    sensor_type: str,
    voltage: float,
    gain: float,
) -> float:
    """Back-calculate resistance from reference voltage."""
    func = _SENSOR_FUNCS.get(sensor_type.lower())
    if func is None:
        raise ValueError(f"Unknown sensor type: {sensor_type!r}")

    ranges = _SENSOR_RANGES.get(sensor_type.lower(), (0, 10000))

    # Simple bisection
    lo, hi = ranges[0] - 100, ranges[1] + 100
    for _ in range(60):
        mid = (lo + hi) / 2
        v = func(mid, gain)
        if abs(v - voltage) < 1e-9:
            break
        if v < voltage:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2
