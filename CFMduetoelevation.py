# -*- coding: utf-8 -*-
"""
Calculates the adjusted CFM required at a target elevation to deliver the
equivalent air mass flow as a baseline CFM at a baseline elevation.

This is based on the principle that air density changes with elevation (pressure),
and HVAC calculations often rely on mass flow rate. To maintain constant mass
flow, the volume flow rate (CFM) must be adjusted inversely proportional
to the density change.
"""

import psychrocalc as pc
import psychrolib as pl
import sys

# --- User Inputs ---
# Modify these values as needed

baseline_cfm = 1000     # Airflow rate at the baseline elevation (Cubic Feet per Minute)
baseline_elevation_ft = 0   # Baseline elevation (e.g., sea level) in feet
target_elevation_ft = 5280  # Target elevation where adjusted CFM is needed, in feet (e.g., Denver)

# Assume constant air conditions for density calculation comparison.
# You can change these if specific conditions are known for each elevation.
dry_bulb_temp_f = 70.0  # Dry-bulb temperature in Fahrenheit
relative_humidity_percent = 50.0 # Relative humidity in percent (e.g., 50 for 50%)

# --- Calculations ---

# Set psychrolib to use IP (Inch-Pound) units
try:
    pl.SetUnitSystem(pl.IP)
except AttributeError:
    print("Error: Could not set unit system. Please ensure psychrolib is correctly installed.")
    print("You might need version 2.5.0 or later: pip install psychrolib==2.5.0")
    sys.exit(1)


# Ensure relative humidity is a fraction (0 to 1) for calculations
rh_fraction = relative_humidity_percent / 100.0
if not (0 <= rh_fraction <= 1):
    print(f"Error: Relative humidity ({relative_humidity_percent}%) must be between 0 and 100.")
    sys.exit(1)

# Calculate standard atmospheric pressure at baseline elevation
try:
    baseline_pressure_pa = pl.GetStandardAtmPressure(baseline_elevation_ft)
    # Convert pressure from Pa to psia for psychrocalc (1 psi = 6894.76 Pa)
    baseline_pressure_psia = baseline_pressure_pa / 6894.76
except ValueError as e:
     print(f"Error calculating baseline pressure: {e}")
     sys.exit(1)

# Calculate standard atmospheric pressure at target elevation
try:
    target_pressure_pa = pl.GetStandardAtmPressure(target_elevation_ft)
    # Convert pressure from Pa to psia for psychrocalc
    target_pressure_psia = target_pressure_pa / 6894.76
except ValueError as e:
     print(f"Error calculating target pressure: {e}")
     sys.exit(1)

# Calculate psychrometric properties (including density) at baseline
try:
    baseline_props = pc.calc_psy_props(
        dry_bulb_temp_f,
        rh_fraction, # Use RH for humidity input type
        pressure_psia=baseline_pressure_psia
    )
    baseline_density = baseline_props['density'] # Density is in lb/ft³ in IP units
    if baseline_density is None or baseline_density <= 0:
        raise ValueError("Calculated baseline density is invalid.")
except Exception as e:
    print(f"Error calculating psychrometric properties at baseline: {e}")
    sys.exit(1)

# Calculate psychrometric properties (including density) at target
try:
    target_props = pc.calc_psy_props(
        dry_bulb_temp_f,
        rh_fraction, # Use RH for humidity input type
        pressure_psia=target_pressure_psia
    )
    target_density = target_props['density'] # Density is in lb/ft³ in IP units
    if target_density is None or target_density <= 0:
        raise ValueError("Calculated target density is invalid.")
except Exception as e:
    print(f"Error calculating psychrometric properties at target: {e}")
    sys.exit(1)

# Calculate the adjusted CFM
# Adjusted CFM = Baseline CFM * (Baseline Density / Target Density)
adjusted_cfm = baseline_cfm * (baseline_density / target_density)

# --- Output ---
print("-" * 40)
print("CFM Adjustment for Elevation Calculator")
print("-" * 40)
print("Inputs:")
print(f"  Baseline CFM:          {baseline_cfm:.2f} CFM")
print(f"  Baseline Elevation:    {baseline_elevation_ft:.0f} ft")
print(f"  Target Elevation:      {target_elevation_ft:.0f} ft")
print(f"  Air Temperature:       {dry_bulb_temp_f:.1f} °F")
print(f"  Relative Humidity:     {relative_humidity_percent:.1f} %")
print("-" * 40)
print("Calculated Values:")
print(f"  Baseline Pressure:     {baseline_pressure_psia:.4f} psia")
print(f"  Baseline Air Density:  {baseline_density:.4f} lb/ft³")
print(f"  Target Pressure:       {target_pressure_psia:.4f} psia")
print(f"  Target Air Density:    {target_density:.4f} lb/ft³")
print("-" * 40)
print("Result:")
print(f"  Adjusted CFM required at {target_elevation_ft:.0f} ft elevation: {adjusted_cfm:.2f} CFM")
print("-" * 40)