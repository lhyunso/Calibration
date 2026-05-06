"""
PT100 RTD sensor configuration.
3-wire quarter bridge.
Reference voltage formula: V = (0.1/2) * ((R - 100) / 100) * gain
  = (R - 100) / 200 * gain
"""
from sensors.base import SensorConfig


class PT100Config(SensorConfig):
    def __init__(self):
        super().__init__(
            name="RTD(PT100)",
            sensor_type="pt100",
            r_nominal=100.0,
            excitation=0.001,      # 1mA excitation current (V/Ω)
            inst_amp_gain=1.0,
            tolerance_ohm=0.385,
            default_resistances=[80.0, 90.0, 100.0, 110.0, 120.0],
            description="PT100 RTD, 3-wire, quarter bridge",
            ref_formula="V = (R - 100) / 200 × Gain",
        )

    def ref_voltage(self, r: float, gain: float = 1.0) -> float:
        """Reference voltage for PT100: V = (R - 100) / 200 * gain"""
        return (r - self.r_nominal) / (2 * self.r_nominal) * gain


class PT1000Config(SensorConfig):
    """PT1000 RTD - same formula, different nominal resistance."""
    def __init__(self):
        super().__init__(
            name="RTD(PT1000)",
            sensor_type="pt1000",
            r_nominal=1000.0,
            excitation=0.0001,     # 0.1mA excitation
            inst_amp_gain=1.0,
            tolerance_ohm=3.85,
            default_resistances=[800.0, 900.0, 1000.0, 1100.0, 1200.0],
            description="PT1000 RTD, 3-wire, quarter bridge",
            ref_formula="V = (R - 1000) / 2000 × Gain",
        )

    def ref_voltage(self, r: float, gain: float = 1.0) -> float:
        return (r - self.r_nominal) / (2 * self.r_nominal) * gain


class Strain350Config(SensorConfig):
    """350Ω Strain Gauge quarter bridge configuration."""
    def __init__(self):
        super().__init__(
            name="Strain Gauge 350Ω",
            sensor_type="strain350",
            r_nominal=350.0,
            excitation=3.5,        # 3.5V bridge excitation
            inst_amp_gain=1.0,
            tolerance_ohm=1.35,
            default_resistances=[320.0, 330.0, 340.0, 350.0, 360.0, 370.0, 380.0],
            description="350Ω Strain Gauge, 3-wire, quarter bridge",
            ref_formula="V = (3.5/2) × (R_offset / 350) × Gain",
        )

    def ref_voltage(self, r: float, gain: float = 1.0) -> float:
        """Reference voltage for 350Ω strain gauge: V = (Vex/2) * (ΔR/R_nom) * gain"""
        return (self.excitation / 2) * ((r - self.r_nominal) / self.r_nominal) * gain


# Registry: sensor_type key → config instance
SENSOR_REGISTRY: dict = {
    "pt100": PT100Config(),
    "pt1000": PT1000Config(),
    "strain350": Strain350Config(),
}


def get_sensor(sensor_type: str) -> SensorConfig:
    """Get sensor config by type string."""
    cfg = SENSOR_REGISTRY.get(sensor_type.lower())
    if cfg is None:
        raise ValueError(
            f"Unknown sensor type: {sensor_type!r}. "
            f"Available: {list(SENSOR_REGISTRY.keys())}"
        )
    return cfg
