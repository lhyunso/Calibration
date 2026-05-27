"""
Base class for sensor calibration configurations.
Extend this for PT100, PT1000, Strain 350Ω, etc.
"""
from dataclasses import dataclass
from typing import List, Dict


@dataclass
class SensorConfig:
    """Sensor-specific calibration parameters."""
    name: str                    # Display name, e.g., "RTD(PT100)"
    sensor_type: str             # Internal key: "pt100", "pt1000", "strain350"
    r_nominal: float             # Nominal resistance at reference point
    excitation: float            # Excitation factor (V/Ω or bridge excitation voltage)
    inst_amp_gain: float         # Instrumentation amplifier gain
    tolerance_ohm: float         # Acceptable deviation in ohms
    default_resistances: List[float]  # Expected calibration resistance steps
    description: str = ""        # Additional description for reports

    # Reference formula description for 3-wire reference
    ref_formula: str = ""

    def ref_voltage(self, r: float, gain: float = 1.0) -> float:
        """Calculate reference voltage for given resistance.
        Formula (single-element varying bridge, current-source excitation):
            V_ref = (V_B / 2) × (ΔR / R_nom) × gain
                  = I_exc × ΔR × gain / 2
        where V_B = I_exc × R_nom, ΔR = r − R_nom.
        Override per sensor if needed.
        """
        return (r - self.r_nominal) * self.excitation * gain / 2.0

    def resistance_from_voltage(self, voltage: float, gain: float) -> float:
        """Convert measured voltage back to resistance.
        Inverts: V_ref = I_exc × ΔR × G_inst / 2  (with calibration gain g ≈ 1)
            ΔR = V × 2 / (gain × I_exc × G_inst)
             R = R_nom + V × 2 / (gain × excitation × inst_amp_gain)
        """
        if gain == 0:
            return self.r_nominal
        return self.r_nominal + voltage * 2.0 / (gain * self.excitation * self.inst_amp_gain)
