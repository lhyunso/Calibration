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
        """Calculate reference voltage for given resistance. Override per sensor."""
        raise NotImplementedError

    def resistance_from_voltage(self, voltage: float, gain: float) -> float:
        """Convert measured voltage back to resistance.
        RTD (PT100/PT1000): V = (R - R_nom) / (2×R_nom) × gain
          → R = 2×R_nom×V / gain + R_nom
        Override this method for sensors with different bridge excitation scaling.
        """
        if gain == 0:
            return self.r_nominal
        return (2 * self.r_nominal * voltage / gain) + self.r_nominal
