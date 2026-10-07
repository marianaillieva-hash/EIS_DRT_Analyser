# EIS DRT Peak Integration Analyser

A lightweight, automated Python tool for post-processing and analysing Electrochemical Impedance Spectroscopy (EIS) data in the Distribution of Relaxation Times (DRT) domain. 
While standard deconvolution software calculates the DRT function $\gamma$ from impedance data, this script automates the physical interpretation step: it identifies discrete relaxation peaks, isolates them, and computes the exact polarisation resistance $R_p$ of individual electrochemical processes via numerical logarithmic integration.
## Key Features
* **Automated Peak Detection:** Identifies physical relaxation peaks based on prominence, ignoring background noise and delta-function artifacts.
* **Logarithmic Integration:** Uses the trapezoidal rule to numerically integrate the area under each peak with respect to $\ln(\tau)$ or $\ln(f)$, strictly adhering to the mathematical definition of polarisation resistance.
* **Automated Reporting:** Generates an Excel report (`DRT_Analysis_Report.xlsx`) containing frequencies, time constants, resistances, and effective capacitances.
* **Publication-Quality Plotting:** Exports high-resolution (600 DPI) vector (SVG) and PDF plots with automatically scaled dual axes (Frequency and Relaxation Time).
## Requirements
The script requires Python 3.8+ and the following libraries:
```bash
pip install numpy pandas scipy matplotlib openpyxl
````

## Data Input Format (`DRT_Data.xlsx`)
The script reads input data from an Excel file named strictly `DRT_Data.xlsx` located in the same directory.
The data must be organised in pairs of columns for each dataset (e.g., different inhibitor concentrations or temperatures). The headers must follow this exact naming convention:
* `f1`, `gamma1` (Dataset 1)
* `f2`, `gamma2` (Dataset 2)
* `f3`, `gamma3` (Dataset 3)
* ...and so on.
* **f**: Frequency values in Hertz (Hz).
* **gamma**: The DRT distribution function values ($\gamma$) in $\Omega \cdot s^{-1}$.
_Note: The script dynamically detects how many pairs exist. You can provide 1 pair or 20 pairs; the tool will process them all. The input file may contain additional exported columns such as `tau1`, `tau2`, etc. To ensure maximum mathematical precision and avoid rounding errors from third-party exports, this script dynamically recalculates the relaxation times directly from the frequency arrays. Therefore, only the `f` and `gamma` columns are strictly required._
## Usage
1. Place your `DRT_Data.xlsx` file in the same folder as the script.
2. Open `drt_analyzer.py` and adjust the configuration parameters if necessary (e.g., `PROMINENCE_FACTOR`, `MIN_PEAK_WIDTH`) to match your specific noise floor.
3. Run the script:
```bash
python drt_analyzer.py
```

4. The script will output the results to the console and generate `DRT_Analysis_Report.xlsx`, `drt_output_plot.pdf`, and `drt_output_plot.svg`.
## Acknowledgements
This tool was conceptualised and mathematically designed by Mariana Ilieva for the analysis of electrochemical impedance data. 
The code generation, structural optimisation, and debugging were assisted by Gemini (Google AI) under the author's direct technical supervision and strict theoretical validation.

## Citation

If you use this script in your research, please cite it using the Zenodo DOI:
> https://doi.org/10.5281/zenodo.23206366
> _Ilieva, M. (2026). DRT Peak Integration Analyser. GitHub repository._

## License

This project is licensed under the MIT License.
