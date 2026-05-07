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
            excitation=0.001,      # 1 mA 정전류 여기
            inst_amp_gain=10.0,    # 기본 Inst. Amp 게인 (RTD ×10)
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
            excitation=0.0001,     # 0.1 mA 정전류 여기
            inst_amp_gain=10.0,    # 기본 Inst. Amp 게인 (RTD ×10)
            tolerance_ohm=3.85,
            default_resistances=[800.0, 900.0, 1000.0, 1100.0, 1200.0],
            description="PT1000 RTD, 3-wire, quarter bridge",
            ref_formula="V = (R - 1000) / 2000 × Gain",
        )

    def ref_voltage(self, r: float, gain: float = 1.0) -> float:
        return (r - self.r_nominal) / (2 * self.r_nominal) * gain


class Strain350Config(SensorConfig):
    """350Ω Strain Gauge quarter bridge — 정전류 여기 방식.
    하드웨어가 PT100/PT1000과 동일한 정전류 회로이므로 공식도 동일:
      V_ref = (R − R_nom) / (2 × R_nom) × inst_amp_gain
    """
    def __init__(self):
        super().__init__(
            name="Strain Gauge 350Ω",
            sensor_type="strain350",
            r_nominal=350.0,
            excitation=0.001,      # 1 mA 정전류 여기
            inst_amp_gain=100.0,   # 기본 Inst. Amp 게인 (Strain ×100)
            tolerance_ohm=1.35,
            default_resistances=[320.0, 330.0, 340.0, 350.0, 360.0, 370.0, 380.0],
            description="350Ω Strain Gauge, 3-wire, 정전류 여기",
            ref_formula="V = (R − 350) / 700 × Gain",
        )

    def ref_voltage(self, r: float, gain: float = 1.0) -> float:
        """정전류 방식 기준전압 — PT100/PT1000과 동일 공식
        V_ref = (R − R_nom) / (2 × R_nom) × gain
        """
        return (r - self.r_nominal) / (2 * self.r_nominal) * gain


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
