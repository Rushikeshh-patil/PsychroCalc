import tkinter as tk
from tkinter import ttk # Themed widgets for a slightly nicer look
from tkinter import messagebox # For showing error popups

import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg, NavigationToolbar2Tk)
import numpy as np

# --- Core Calculation Logic (adapted for GUI) ---
# Keep calculation separate from GUI interaction where possible

def calculate_system_data(cfm_design, sp_design, cfm_reference=None):
    """
    Calculates system coefficient and points needed for plotting.

    Args:
        cfm_design (float): Design CFM.
        sp_design (float): Design SP (in. w.g.).
        cfm_reference (float, optional): Reference CFM. Defaults to None.

    Returns:
        tuple: (system_k, cfm_plot_range, sp_plot_values, sp_reference_calculated)
               Returns None for sp_reference_calculated if cfm_reference is None.
               Raises ValueError on invalid input.
    """
    # Validation
    if cfm_design <= 0:
        raise ValueError("Design CFM must be positive.")
    if sp_design < 0:
        raise ValueError("Design Static Pressure cannot be negative.")
    if cfm_reference is not None and cfm_reference < 0:
        raise ValueError("Reference CFM cannot be negative.")

    # Calculate System Resistance Coefficient (k)
    system_k = 0.0
    # Check cfm_design > 0 before division, even though validated above (robustness)
    if sp_design > 0 and cfm_design > 0:
         system_k = sp_design / (cfm_design ** 2)
    elif sp_design == 0:
        # If SP is 0, resistance is 0 regardless of CFM (as long as CFM > 0)
        system_k = 0.0


    # Determine plot range based on design and potential reference point
    max_cfm_for_range = cfm_design
    if cfm_reference is not None:
        max_cfm_for_range = max(cfm_design, cfm_reference)

    # Ensure max_cfm_for_range is positive before multiplying (edge case if cfm_design=0 and no ref)
    if max_cfm_for_range <= 0 : max_cfm_for_range = 1 # Avoid plotting range issues if input somehow allows 0

    cfm_max_plot = max_cfm_for_range * 1.5 # Plot up to 1.5x the max CFM shown
    cfm_plot_range = np.linspace(0, cfm_max_plot, 100) # 100 points for smoothness

    # Calculate curve points
    sp_plot_values = system_k * (cfm_plot_range ** 2)

    # Calculate reference SP if needed
    sp_reference_calculated = None
    if cfm_reference is not None:
        # Handle cfm_reference == 0 explicitly if system_k calculation depends on cfm_design > 0
        if cfm_reference == 0:
             sp_reference_calculated = 0.0
        else:
            sp_reference_calculated = system_k * (cfm_reference ** 2)

    return system_k, cfm_plot_range, sp_plot_values, sp_reference_calculated

# --- GUI Application Class ---

class SystemCurveApp:
    def __init__(self, root):
        self.root = root
        self.root.title("System Curve Generator")
        # Set minimum size to prevent collapsing too small
        self.root.minsize(600, 500)

        # --- Main Frame Container ---
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)

        # Configure grid columns/rows for resizing
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1) # Allow plot area (row 1) to expand

        # --- Input Frame (Top Row) ---
        input_frame = ttk.LabelFrame(main_frame, text="Inputs", padding="10")
        # Use grid layout manager for the main layout
        input_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        input_frame.columnconfigure(1, weight=1) # Allow entry column to expand slightly

        # Design Point Inputs
        ttk.Label(input_frame, text="Design CFM:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.cfm_design_var = tk.StringVar()
        self.cfm_design_entry = ttk.Entry(input_frame, textvariable=self.cfm_design_var, width=15)
        self.cfm_design_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(input_frame, text="Design SP (in. w.g.):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sp_design_var = tk.StringVar()
        self.sp_design_entry = ttk.Entry(input_frame, textvariable=self.sp_design_var, width=15)
        self.sp_design_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        # Reference Point Inputs
        self.add_ref_var = tk.BooleanVar()
        self.ref_check = ttk.Checkbutton(input_frame, text="Add Reference Point?", variable=self.add_ref_var, command=self.toggle_ref_entry)
        self.ref_check.grid(row=2, column=0, padx=5, pady=10, sticky="w")

        self.ref_cfm_label = ttk.Label(input_frame, text="Reference CFM:")
        self.ref_cfm_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.cfm_ref_var = tk.StringVar()
        self.cfm_ref_entry = ttk.Entry(input_frame, textvariable=self.cfm_ref_var, width=15)
        self.cfm_ref_entry.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        # Initially disable reference input
        self.toggle_ref_entry()

        # Plot Button
        self.plot_button = ttk.Button(input_frame, text="Generate Curve", command=self.plot_curve)
        self.plot_button.grid(row=4, column=0, columnspan=2, pady=(10, 5))

        # --- Plot Frame (Bottom Row, Expanding) ---
        plot_frame = ttk.Frame(main_frame, padding="5")
        plot_frame.grid(row=1, column=0, padx=10, pady=(5, 10), sticky="nsew") # North South East West sticky
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(1, weight=1) # Canvas row expands

        # Create Matplotlib Figure and Axes
        self.fig = Figure(figsize=(6, 4), dpi=100) # Adjusted figsize for embedding
        self.ax = self.fig.add_subplot(111)

        # Create Tkinter Canvas to display the figure
        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        # Place canvas in row 1, allowing it to expand
        self.canvas_widget.grid(row=1, column=0, sticky="nsew")

        # Add Matplotlib Navigation Toolbar (Place above canvas in row 0)
        toolbar_frame = ttk.Frame(plot_frame) # Frame to hold toolbar
        toolbar_frame.grid(row=0, column=0, sticky="ew")
        toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        toolbar.update()

        # Initialize plot area appearance
        self.initialize_plot()
        self.fig.tight_layout() # Adjust layout initially


    def toggle_ref_entry(self):
        """Enables or disables the Reference CFM entry based on the checkbox."""
        if self.add_ref_var.get():
            self.cfm_ref_entry.config(state=tk.NORMAL)
            self.ref_cfm_label.config(state=tk.NORMAL)
        else:
            self.cfm_ref_entry.config(state=tk.DISABLED)
            self.ref_cfm_label.config(state=tk.DISABLED)
            self.cfm_ref_var.set("") # Clear the entry when disabled

    def initialize_plot(self):
        """Sets up the initial appearance of the plot axes."""
        self.ax.clear()
        self.ax.set_xlabel("Airflow (CFM)")
        self.ax.set_ylabel("Static Pressure (inches w.g.)")
        self.ax.set_title("Duct System Curve")
        self.ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        # Optionally set initial limits if desired
        self.ax.set_xlim(left=0)
        self.ax.set_ylim(bottom=0)
        self.canvas.draw()

    def plot_curve(self):
        """Gets inputs, calculates data, and plots the curve on the canvas."""
        try:
            # --- Get and Validate Inputs ---
            if not self.cfm_design_var.get() or not self.sp_design_var.get():
                 raise ValueError("Design CFM and SP cannot be empty.")

            cfm_design = float(self.cfm_design_var.get())
            sp_design = float(self.sp_design_var.get())

            cfm_reference = None
            sp_reference_calculated = None # Initialize here

            if self.add_ref_var.get():
                if not self.cfm_ref_var.get():
                    raise ValueError("Reference CFM cannot be empty when checked.")
                cfm_reference = float(self.cfm_ref_var.get())

            # Perform core calculations (includes internal validation)
            system_k, cfm_plot_range, sp_plot_values, sp_reference_calculated = calculate_system_data(
                cfm_design, sp_design, cfm_reference
            )

            # --- Plotting ---
            self.ax.clear() # Clear previous plot content before drawing new

            # Plot the system curve
            self.ax.plot(cfm_plot_range, sp_plot_values, label='System Curve', color='blue', linewidth=2)

            # Mark the design point
            self.ax.plot(cfm_design, sp_design, 'ro', markersize=7, label=f'Design ({cfm_design:.0f}, {sp_design:.2f})')
            # Add simple annotation (adjust offset points as needed)
            self.ax.annotate(f' {cfm_design:.0f} CFM\n {sp_design:.2f}" wg',
                             xy=(cfm_design, sp_design), xytext=(5, 5),
                             textcoords='offset points', ha='left', va='bottom',
                             fontsize=9)

            # Mark the reference point if applicable
            if cfm_reference is not None and sp_reference_calculated is not None:
                self.ax.plot(cfm_reference, sp_reference_calculated, 'gs', markersize=7,
                             label=f'Reference ({cfm_reference:.0f}, {sp_reference_calculated:.2f})')
                # Add annotation for reference point
                self.ax.annotate(f' {cfm_reference:.0f} CFM\n {sp_reference_calculated:.2f}" wg',
                             xy=(cfm_reference, sp_reference_calculated), xytext=(5, -15),
                             textcoords='offset points', ha='left', va='top',
                             fontsize=9, color='green')

            # --- Finalize Plot ---
            self.ax.set_xlabel("Airflow (CFM)")
            self.ax.set_ylabel("Static Pressure (inches w.g.)")
            self.ax.set_title("Duct System Curve")
            self.ax.grid(True, which='both', linestyle='--', linewidth=0.5)
            self.ax.legend(fontsize='small')

            # Set Axis Limits dynamically (ensure visibility)
            max_sp_on_plot = max(sp_plot_values) if len(sp_plot_values) > 0 else sp_design
            if sp_reference_calculated is not None:
                 max_sp_on_plot = max(max_sp_on_plot, sp_reference_calculated)
            # Handle case where max SP is 0 to avoid empty ylim
            if max_sp_on_plot <= 0: max_sp_on_plot = 1.0

            self.ax.set_xlim(left=0, right=max(cfm_plot_range) * 1.05 if len(cfm_plot_range) > 0 else cfm_design * 1.5)
            self.ax.set_ylim(bottom=0, top=max_sp_on_plot * 1.1)

            # Adjust layout AFTER setting limits and drawing
            self.fig.tight_layout()
            self.canvas.draw() # Redraw the canvas with the new plot

        except ValueError as e:
            messagebox.showerror("Input Error", f"Invalid input: {e}")
            self.initialize_plot() # Reset plot on input error
        except Exception as e:
            # Catch any other unexpected errors during plotting/calculation
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            self.initialize_plot() # Reset plot on unexpected error


# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = SystemCurveApp(root)
    root.mainloop()