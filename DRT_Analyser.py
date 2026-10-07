"""
DRT Peak Integration & Analysis Tool
------------------------------------
This script processes Electrochemical Impedance Spectroscopy (EIS) data transformed 
into the Distribution of Relaxation Times (DRT) domain. It automatically identifies 
relaxation peaks based on physical prominence, integrates the area under each peak 
using the trapezoidal rule to calculate the polarization resistance (Rp) of individual 
electrochemical processes, and outputs an Excel report and a plot.

Conceptualization, methodology, and technical validation: Mariana Ilieva
Code generation and optimization assistance: Gemini (Google AI)
License: MIT
"""

import os
import platform
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import matplotlib as mpl
from scipy.signal import find_peaks, peak_widths
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==============================================================================
# CONFIGURATION PARAMETERS (Adjust based on experimental noise and limits)
# ==============================================================================
PROMINENCE_FACTOR = 0.005  # 0.5% of max peak height to ignore background noise
MIN_PEAK_WIDTH = 2         # Minimum width to filter out delta-function artifacts
REL_HEIGHT_VAL = 0.85      # Integration boundaries set at 85% relative peak depth

# ==============================================================================
# BACKEND AUTOMATION
# ==============================================================================
try:
    shell = get_ipython().__class__.__name__
    is_jupyter = (shell == 'ZMQInteractiveShell')
except NameError:
    is_jupyter = False

if not is_jupyter:
    if platform.system() == 'Linux' and os.environ.get('DISPLAY', '') == '':
        mpl.use('Agg')

# ==============================================================================
# 1. TYPOGRAPHY AND GEOMETRY
# ==============================================================================
TARGET_WIDTH_CM = 12.0
TARGET_WIDTH_INCHES = TARGET_WIDTH_CM / 2.54 
ASPECT_RATIO = 0.75 

mpl.rcParams.update({
    'figure.figsize': (TARGET_WIDTH_INCHES, TARGET_WIDTH_INCHES * ASPECT_RATIO),
    'font.family': 'sans-serif',
    'font.sans-serif': ['Arial', 'Helvetica'],
    'font.size': 10,
    'axes.labelsize': 10,
    'xtick.labelsize': 9,
    'ytick.labelsize': 9,
    'legend.fontsize': 8,
    'axes.unicode_minus': True,
    'axes.facecolor': 'white',
    'axes.edgecolor': 'black',
    'axes.linewidth': 1.0,
    'axes.grid': False,
    'savefig.dpi': 600,
})

# Generic color palette for multiple datasets
COLORS = ['#000000', '#D55E00', '#0072B2', '#009E73', '#CC79A7', '#F0E442', '#56B4E9']

# ==============================================================================
# 2. HELPER FUNCTIONS
# ==============================================================================
def hz_to_tau(x):
    return 1 / (2 * np.pi * np.where(x == 0, np.nan, x))

def tau_to_hz(x):
    return 1 / (2 * np.pi * np.where(x == 0, np.nan, x))

# ==============================================================================
# 3. MAIN PROCESSING AND PLOTTING FUNCTION
# ==============================================================================
def process_and_plot_drt():
    data_path = 'DRT_Data.xlsx'
    
    try:
        df = pd.read_excel(data_path)
    except FileNotFoundError:
        print(f"Error: Input file '{data_path}' not found in the current directory.")
        return

    # --- INITIALIZE EXCEL REPORT ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DRT Results Summary"
    ws.views.sheetView[0].showGridLines = True
    
    # Excel Styles
    navy_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Arial", size=11)
    bold_font = Font(name="Arial", size=11, bold=True)
    zebra_fill = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
    total_fill = PatternFill(start_color="E6EDF5", end_color="E6EDF5", fill_type="solid")
    
    thin_border = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'),
                         top=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
    double_bottom = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    # Setup Excel Headers
    headers = ["Dataset", "Peak No.", "Frequency f (Hz)", "Time Constant τ (s)", "Resistance Rp (Ω)", "Capacitance C (F)"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    excel_row = 2

    # --- PLOTTING PREPARATION ---
    fig, ax1 = plt.subplots()
    global_min_f, global_max_f, global_max_gamma = float('inf'), float('-inf'), 0
    
    print(f"{'Dataset':<10} | {'Freq. f (Hz)':<15} | {'Tau (s)':<15} | {'Rp (Ohm)':<15}")
    print("-" * 65)

    # Dynamically detect how many datasets (f/gamma pairs) exist in the Excel file
    col_pairs = [int(col[1:]) for col in df.columns if col.startswith('f') and col[1:].isdigit()]
    
    if not col_pairs:
        print("Error: No valid 'f' and 'gamma' columns found. Format should be f1, gamma1, f2, gamma2, etc.")
        return

    for i in sorted(col_pairs):
        col_f, col_gamma = f'f{i}', f'gamma{i}'
        
        if col_f in df.columns and col_gamma in df.columns:
            f = df[col_f].dropna().values
            gamma = df[col_gamma].dropna().values
            
            if len(f) > 0:
                sort_idx = np.argsort(f)
                f_sorted = f[sort_idx]
                gamma_sorted = gamma[sort_idx]

                global_min_f = min(global_min_f, np.min(f_sorted[f_sorted > 0]))
                global_max_f = max(global_max_f, np.max(f_sorted))
                global_max_gamma = max(global_max_gamma, np.max(gamma_sorted))

                # Identify peaks physically based on prominence
                prominence_threshold = PROMINENCE_FACTOR * np.max(gamma_sorted)
                peaks, _ = find_peaks(gamma_sorted, prominence=prominence_threshold, width=MIN_PEAK_WIDTH)
                
                dataset_label = f"Dataset {i}"
                color = COLORS[(i-1) % len(COLORS)]
                
                if len(peaks) > 0:
                    widths, width_heights, left_ips, right_ips = peak_widths(gamma_sorted, peaks, rel_height=REL_HEIGHT_VAL)
                    start_excel_row = excel_row 
                    
                    for p_idx, (p, left, right) in enumerate(zip(peaks, left_ips, right_ips), 1):
                        f_peak = f_sorted[p]
                        
                        l_idx = max(0, int(np.floor(left)))
                        r_idx = min(len(f_sorted)-1, int(np.ceil(right)))
                        
                        f_window = f_sorted[l_idx:r_idx+1]
                        tau_window = 1 / (2 * np.pi * f_window)
                        gamma_window = gamma_sorted[l_idx:r_idx+1]
                        
                        if len(tau_window) > 1:
                            sort_tau_idx = np.argsort(tau_window)
                            # Logarithmic integration
                            area = np.trapezoid(gamma_window[sort_tau_idx], x=np.log(tau_window[sort_tau_idx]))
                        else:
                            area = 0.0
                            
                        print(f"{dataset_label:<10} | {f_peak:<15.2f} | {1/(2*np.pi*f_peak):<15.5e} | {area:<15.2f}")
                        
                        # Write to Excel
                        ws.cell(row=excel_row, column=1, value=dataset_label).alignment = Alignment(horizontal="center")
                        ws.cell(row=excel_row, column=2, value=f"Peak {p_idx}").alignment = Alignment(horizontal="center")
                        
                        c_f = ws.cell(row=excel_row, column=3, value=f_peak)
                        c_f.number_format = '#,##0.0'
                        
                        c_tau = ws.cell(row=excel_row, column=4, value=f"=1/(2*PI()*C{excel_row})")
                        c_tau.number_format = '0.000E+00'
                        
                        c_r = ws.cell(row=excel_row, column=5, value=area)
                        c_r.number_format = '#,##0.00'
                        
                        c_c = ws.cell(row=excel_row, column=6, value=f"=D{excel_row}/E{excel_row}")
                        c_c.number_format = '0.000E+00'
                        
                        for col in range(1, 7):
                            cell = ws.cell(row=excel_row, column=col)
                            cell.font = regular_font
                            cell.border = thin_border
                            if excel_row % 2 == 1:
                                cell.fill = zebra_fill
                                
                        excel_row += 1
                        
                    # Write Total Rp Row
                    ws.cell(row=excel_row, column=1, value=f"Total {dataset_label} (Rp)").font = bold_font
                    c_sum = ws.cell(row=excel_row, column=5, value=f"=SUM(E{start_excel_row}:E{excel_row-1})")
                    c_sum.number_format = '#,##0.00'
                    c_sum.font = bold_font
                    
                    for col in range(1, 7):
                        cell = ws.cell(row=excel_row, column=col)
                        cell.fill = total_fill
                        cell.border = double_bottom
                        
                    excel_row += 2 
                
                ax1.plot(f_sorted, gamma_sorted, color=color, linewidth=1.5, label=dataset_label)

    # Auto-adjust Excel column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

    wb.save('DRT_Analysis_Report.xlsx')
    print("\n[INFO] Excel report successfully saved as 'DRT_Analysis_Report.xlsx'")

    # --- PLOT FORMATTING ---
    ax1.set_xscale('log')
    ax1.set_xlabel('Frequency $f$ (Hz)', color='black', labelpad=6)
    ax1.set_ylabel(r'Distribution Function $\gamma$ ($\Omega\cdot\mathregular{s^{-1}}$)', color='black', labelpad=6)

    if global_min_f != float('inf'):
        ax1.set_xlim(global_min_f, global_max_f)
        ax1.set_ylim(0, global_max_gamma * 1.05)

    ax2 = ax1.secondary_xaxis('top', functions=(hz_to_tau, tau_to_hz))
    ax2.set_xlabel(r'Relaxation Time $\tau$ (s)', color='black', labelpad=6)

    for axis in [ax1.xaxis, ax2.xaxis]:
        axis.set_major_locator(ticker.LogLocator(base=10.0, numticks=10))
        axis.set_minor_locator(ticker.LogLocator(base=10.0, subs=np.arange(2, 10) * 0.1, numticks=10))
        axis.set_minor_formatter(ticker.NullFormatter())
        
    ax1.tick_params(axis='x', which='major', direction='out', pad=4, length=5, width=1.0)
    ax1.tick_params(axis='y', which='major', direction='out', pad=4, length=5, width=1.0)
    ax2.tick_params(axis='x', which='major', direction='out', pad=4, length=5, width=1.0)
    ax1.tick_params(axis='both', which='minor', direction='out', length=3, width=0.8)
    ax2.tick_params(axis='x', which='minor', direction='out', length=3, width=0.8)
    ax1.legend(loc='best', frameon=True, edgecolor='black', handlelength=1.5)

    plt.tight_layout()
    fig.savefig('drt_output_plot.pdf', format='pdf', bbox_inches='tight', pad_inches=0.1)
    fig.savefig('drt_output_plot.svg', format='svg', bbox_inches='tight', pad_inches=0.1)
    print(f"[INFO] High-resolution plots saved as PDF and SVG.")
    
    if is_jupyter:
        plt.show()
    else:
        plt.close(fig) 

if __name__ == '__main__':
    process_and_plot_drt()