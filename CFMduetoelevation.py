#!/usr/bin/env python3
import psychrolib as psy
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

# --- Tooltip Class (Adjusted for Standard/Light Mode) ---
class Tooltip:
    """
    Creates a tooltip (pop-up window) for a given Tkinter widget.
    Styled for standard light backgrounds.
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.bg_color = "#FFFFE0" # Light yellow background
        # Default foreground (black) will be used

        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)
        self.widget.bind("<ButtonPress>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return

        x, y, _, _ = self.widget.bbox("insert")
        if x is None:
             x, y = 0, 0
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20

        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.attributes('-topmost', True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background=self.bg_color, # Light background
                         # Default foreground (usually black)
                         relief=tk.SOLID, borderwidth=1,
                         wraplength=250,
                         font=("tahoma", "8", "normal"),
                         padx=5, pady=5)
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            try:
                tw.destroy()
            except tk.TclError:
                pass


# --- Core Calculation Logic (Unchanged) ---
try:
    psy.SetUnitSystem(psy.IP)
except Exception as e:
    messagebox.showerror("Initialization Error", f"Failed to set PsychroLib unit system: {e}", icon='error')
    sys.exit(1)

def calculate_adjusted_cfm(cfm_sl, elevation_ft, temp_f, rel_hum_fraction):
    # (Calculation function remains the same as previous version)
    if cfm_sl <= 0: raise ValueError("Sea Level CFM must be a positive value.")
    pressure_sl_psi = psy.GetStandardAtmPressure(0.0)
    if pressure_sl_psi <= 0: raise RuntimeError("Failed to calc sea level pressure.")
    hum_ratio_sl = psy.GetHumRatioFromRelHum(temp_f, rel_hum_fraction, pressure_sl_psi)
    density_sl_lbm_ft3 = psy.GetMoistAirDensity(temp_f, hum_ratio_sl, pressure_sl_psi)
    if density_sl_lbm_ft3 <= 0: raise RuntimeError(f"Calc sea level density ({density_sl_lbm_ft3}) non-positive.")

    pressure_alt_psi = psy.GetStandardAtmPressure(elevation_ft)
    if pressure_alt_psi <= 0: raise ValueError(f"Calc pressure at {elevation_ft:,.0f} ft ({pressure_alt_psi} psi) non-positive. Elevation too high?")
    hum_ratio_alt = psy.GetHumRatioFromRelHum(temp_f, rel_hum_fraction, pressure_alt_psi)
    density_alt_lbm_ft3 = psy.GetMoistAirDensity(temp_f, hum_ratio_alt, pressure_alt_psi)
    if density_alt_lbm_ft3 <= 0: raise RuntimeError(f"Calc altitude density ({density_alt_lbm_ft3}) non-positive.")

    density_ratio = density_sl_lbm_ft3 / density_alt_lbm_ft3
    adjusted_cfm = cfm_sl * density_ratio
    return adjusted_cfm, density_sl_lbm_ft3, density_alt_lbm_ft3, pressure_sl_psi, pressure_alt_psi, density_ratio


# --- GUI Application Class ---
class CfmCalculatorApp:
    def __init__(self, master):
        self.master = master
        master.title("CFM Elevation Adjustment Calculator") # Removed (Dark Mode)
        master.resizable(False, False)

        # --- Configure Style (Using default ttk theme) ---
        style = ttk.Style()
        # Let ttk use the OS's default theme
        theme_used = style.theme_use()
        print(f"Using default TTK theme: {theme_used}") # Info

        # Basic styling - rely on theme defaults for colors
        style.configure("TLabel", padding=5, font=('Segoe UI', 9))
        style.configure("TEntry", padding=5, font=('Segoe UI', 9))
        style.configure("TButton", padding=5, font=('Segoe UI', 9, 'bold'))
        style.configure("TLabelframe", padding=5)
        style.configure("TLabelframe.Label", padding=(5,0), font=('Segoe UI', 10, 'bold'))

        # Custom styles (using default foreground colors, just bolding result)
        style.configure("Result.TLabel", font=('Segoe UI', 10, "bold")) # e.g., Black on Windows default
        style.configure("Detail.TLabel", font=('Segoe UI', 9)) # Default color


        # --- Main frame ---
        main_frame = ttk.Frame(master, padding="15 15 15 15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # --- Input Fields ---
        # (Widget creation and layout remains the same, tooltips use light style)
        # Row 0: Sea Level CFM
        cfm_label = ttk.Label(main_frame, text="Sea Level CFM:")
        cfm_label.grid(row=0, column=0, sticky=tk.W)
        self.cfm_sl_var = tk.StringVar()
        cfm_entry = ttk.Entry(main_frame, textvariable=self.cfm_sl_var, width=18)
        cfm_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))
        Tooltip(cfm_label, "Baseline Volumetric Flow Rate (Cubic Feet per Minute)\nrequired at standard sea level conditions (approx 14.7 psi).")
        Tooltip(cfm_entry, "Enter the equipment's rated CFM at sea level.")

        # Row 1: Elevation
        elev_label = ttk.Label(main_frame, text="Elevation (ft):")
        elev_label.grid(row=1, column=0, sticky=tk.W)
        self.elevation_var = tk.StringVar()
        elev_entry = ttk.Entry(main_frame, textvariable=self.elevation_var, width=18)
        elev_entry.grid(row=1, column=1, sticky=(tk.W, tk.E))
        Tooltip(elev_label, "The target elevation above sea level where the\nequipment will operate, in feet.")
        Tooltip(elev_entry, "Enter elevation in feet (e.g., 5280 for Denver).")

        # Row 2: Temperature
        temp_label = ttk.Label(main_frame, text="Temperature (°F):")
        temp_label.grid(row=2, column=0, sticky=tk.W)
        self.temp_f_var = tk.StringVar(value="70.0")
        temp_entry = ttk.Entry(main_frame, textvariable=self.temp_f_var, width=18)
        temp_entry.grid(row=2, column=1, sticky=(tk.W, tk.E))
        Tooltip(temp_label, "Reference dry-bulb air temperature in Fahrenheit.\nUsed for calculating air density at both locations.")
        Tooltip(temp_entry, "Enter temperature in °F (e.g., 70). Affects density.")

        # Row 3: Relative Humidity
        rh_label = ttk.Label(main_frame, text="Relative Humidity (%):")
        rh_label.grid(row=3, column=0, sticky=tk.W)
        self.rel_hum_var = tk.StringVar(value="50.0")
        rh_entry = ttk.Entry(main_frame, textvariable=self.rel_hum_var, width=18)
        rh_entry.grid(row=3, column=1, sticky=(tk.W, tk.E))
        Tooltip(rh_label, "Reference relative humidity (0-100%).\nUsed for calculating moist air density.")
        Tooltip(rh_entry, "Enter RH percentage (e.g., 50). Minor effect on density.")

        # --- Calculate Button ---
        calc_button = ttk.Button(main_frame, text="Calculate Adjusted CFM", command=self.perform_calculation)
        calc_button.grid(row=4, column=0, columnspan=2, pady=15)

        # --- Output Area ---
        result_frame = ttk.LabelFrame(main_frame, text="Results", padding="10 10 10 10")
        result_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E))

        # Adjusted CFM (Main Result)
        self.adj_cfm_var = tk.StringVar(value="---")
        adj_cfm_lbl_desc = ttk.Label(result_frame, text="Adjusted CFM at Elevation:", font=('Segoe UI', 10, 'bold')) # Default color
        adj_cfm_lbl_desc.grid(row=0, column=0, sticky=tk.W)
        adj_cfm_lbl_val = ttk.Label(result_frame, textvariable=self.adj_cfm_var, style="Result.TLabel") # Bold style
        adj_cfm_lbl_val.grid(row=0, column=1, sticky=tk.W)
        Tooltip(adj_cfm_lbl_desc, "Required volumetric flow rate at the target elevation\nto deliver the equivalent MASS flow rate as the Sea Level CFM.")
        Tooltip(adj_cfm_lbl_val, "Calculated CFM needed at altitude.")

        # Density Ratio (Key Factor)
        self.density_ratio_var = tk.StringVar(value="---")
        den_ratio_lbl_desc = ttk.Label(result_frame, text="Density Ratio (SL/Alt):", style="Detail.TLabel") # Detail style (default color)
        den_ratio_lbl_desc.grid(row=1, column=0, sticky=tk.W)
        den_ratio_lbl_val = ttk.Label(result_frame, textvariable=self.density_ratio_var, style="Detail.TLabel")
        den_ratio_lbl_val.grid(row=1, column=1, sticky=tk.W)
        Tooltip(den_ratio_lbl_desc, "Ratio of air density at sea level to air density at altitude.\nThis is the multiplier applied to the Sea Level CFM.")
        Tooltip(den_ratio_lbl_val, "Calculated ratio: Density(Sea Level) / Density(Altitude).")

        # --- Detailed Output ---
        detail_frame = ttk.LabelFrame(main_frame, text="Details", padding="10 10 10 10")
        detail_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10,0))

        # Detail labels use the "Detail.TLabel" style (default colors)
        self.pressure_sl_var = tk.StringVar(value="---")
        ttk.Label(detail_frame, text="Sea Level Pressure (psi):", style="Detail.TLabel").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(detail_frame, textvariable=self.pressure_sl_var, style="Detail.TLabel").grid(row=0, column=1, sticky=tk.W)

        self.density_sl_var = tk.StringVar(value="---")
        ttk.Label(detail_frame, text="Sea Level Density (lb/ft³):", style="Detail.TLabel").grid(row=1, column=0, sticky=tk.W)
        ttk.Label(detail_frame, textvariable=self.density_sl_var, style="Detail.TLabel").grid(row=1, column=1, sticky=tk.W)

        self.pressure_alt_var = tk.StringVar(value="---")
        ttk.Label(detail_frame, text="Elevation Pressure (psi):", style="Detail.TLabel").grid(row=2, column=0, sticky=tk.W)
        ttk.Label(detail_frame, textvariable=self.pressure_alt_var, style="Detail.TLabel").grid(row=2, column=1, sticky=tk.W)

        self.density_alt_var = tk.StringVar(value="---")
        ttk.Label(detail_frame, text="Elevation Density (lb/ft³):", style="Detail.TLabel").grid(row=3, column=0, sticky=tk.W)
        ttk.Label(detail_frame, textvariable=self.density_alt_var, style="Detail.TLabel").grid(row=3, column=1, sticky=tk.W)

        # Configure padding for grid cells
        for child in main_frame.winfo_children():
             if isinstance(child, ttk.LabelFrame):
                 for sub_child in child.winfo_children():
                      sub_child.grid_configure(padx=5, pady=2)
             else:
                  child.grid_configure(padx=5, pady=4)

    # perform_calculation and clear_results methods remain the same

    def perform_calculation(self):
        """Reads inputs, validates, calls calculation, and updates GUI."""
        try:
            cfm_sl = float(self.cfm_sl_var.get())
            elevation_ft = float(self.elevation_var.get())
            temp_f = float(self.temp_f_var.get())
            rel_hum_percent = float(self.rel_hum_var.get())

            if not (0 <= rel_hum_percent <= 100):
                raise ValueError("Relative humidity must be between 0% and 100%.")
            rel_hum_fraction = rel_hum_percent / 100.0

            (adj_cfm, den_sl, den_alt, pres_sl, pres_alt, den_ratio) = calculate_adjusted_cfm(
                cfm_sl, elevation_ft, temp_f, rel_hum_fraction
            )

            self.adj_cfm_var.set(f"{adj_cfm:,.2f} CFM")
            self.density_ratio_var.set(f"{den_ratio:.4f}")
            self.density_sl_var.set(f"{den_sl:.5f}")
            self.density_alt_var.set(f"{den_alt:.5f}")
            self.pressure_sl_var.set(f"{pres_sl:.3f}")
            self.pressure_alt_var.set(f"{pres_alt:.3f}")

        except ValueError as ve:
            messagebox.showerror("Input Error", f"Invalid input: {ve}", icon='warning')
            self.clear_results(clear_inputs=False)
        except RuntimeError as re:
             messagebox.showerror("Calculation Error", f"Calculation failed: {re}", icon='error')
             self.clear_results(clear_inputs=False)
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}", icon='error')
            self.clear_results(clear_inputs=False)

    def clear_results(self, clear_inputs=False):
        """Clears the result fields and optionally input fields."""
        self.adj_cfm_var.set("---")
        self.density_ratio_var.set("---")
        self.density_sl_var.set("---")
        self.density_alt_var.set("---")
        self.pressure_sl_var.set("---")
        self.pressure_alt_var.set("---")
        if clear_inputs:
            self.cfm_sl_var.set("")
            self.elevation_var.set("")
            self.temp_f_var.set("70.0")
            self.rel_hum_var.set("50.0")


# --- Main Execution ---
if __name__ == "__main__":
    root = None
    try:
        # Check psychrolib early
        if 'psy' not in globals():
             raise ImportError("PsychroLib could not be initialized.")

        # Use standard Tkinter root window
        root = tk.Tk()

        app = CfmCalculatorApp(root)
        root.mainloop()

    except ImportError as ie:
         # Handle missing psychrolib
         msg = f"Missing required library: {ie}\n"
         if 'psychrolib' in str(ie).lower(): msg += "Please install it using: pip install psychrolib"
         if root: root.destroy()
         messagebox.showerror("Dependency Error", msg, icon='error')
    except Exception as e:
        # Catch any other unexpected startup errors
        if root: root.destroy()
        messagebox.showerror("Startup Error", f"An error occurred on startup: {e}", icon='error')