# Witec-Matlab

*(Based on [SOFTX-D-20-00088](https://github.com/ElsevierSoftwareX/SOFTX-D-20-00088))*

**Version:** v2.0.0
**Status:** Stable
**Language:** MATLAB
**Category:** Scientific Data Processing / Spectroscopy Analysis

Witec-Matlab is a MATLAB-based data processing toolbox designed for **WITec Raman / PL spectroscopy data**, with a particular focus on **batch data extraction, polarization-resolved analysis, and parameter calculation** (e.g., spectral dependence and Landé g-factor–related quantities).
The toolbox is developed as a **research-oriented workflow**, optimized for reproducibility, extensibility, and high-throughput analysis in experimental condensed-matter and optical spectroscopy studies.

This project builds upon and extends the open-source **WITio** framework (Elsevier SoftwareX, SOFTX-D-20-00088), providing higher-level, unified analysis pipelines tailored for practical research use.

------

## Key Features

- Unified data extraction and preprocessing for WITec spectroscopy files
- Polarization-resolved data management and batch processing
- Modular spectral parameter calculation workflows
- Flexible Excel merging and data normalization utilities
- Keyword-based file organization for large experimental datasets
- Clean, extensible MATLAB script structure suitable for long-term research projects

------

## What’s New in v2.0.0

Version **v2.0.0** represents a **major structural refactor** of the toolbox rather than a minor update.

### Highlights

- **Unified pipeline design**: legacy scripts have been consolidated into `_Unified` modules with consistent interfaces
- **Reduced redundancy**: removed duplicated logic across preprocessing and calculation steps
- **Improved maintainability**: clearer module boundaries and naming conventions
- **Better scalability**: designed for batch analysis of large datasets

### Breaking Changes

- This version is **not fully backward compatible** with earlier script layouts
- Script names and calling logic have been updated
- Users are strongly encouraged to migrate directly to v2.0.0

------

## Repository Structure (v2.0.0)

```
Witec-Matlab/
│
├── 1. Data extraction and preprocessing/
│   ├── A1_WITec_Spectrum_Export_and_Classification_Unified.m
│   ├── A2_PolarDataManager_Unified.m
│   └── A3_FlexiExcelMerge_Unified.m
│
├── 2. Parameter calculation/
│   ├── A5_Spectral_Batch_Analysis_Unified.m
│   ├── B1_FileMoverByKeyword_Unified.m
│   ├── B2_Normalization.m
│   └── B3_SplitOddEvenColumns.m
│
├── docs/              # Documentation and notes (optional)
├── examples/          # Example scripts and usage demos (recommended)
└── README.md
```

------

## Typical Workflow

A standard data-analysis workflow using Witec-Matlab is as follows:

1. **Raw data import**
   Import WITec spectral data using WITio-based routines.
2. **Spectrum classification & preprocessing**
   Organize spectra by experimental conditions (e.g., polarization, magnetic field, bias).
3. **Polarization data management**
   Extract and manage polarization-resolved datasets using unified data structures.
4. **Batch spectral analysis**
   Perform peak extraction, spectral fitting, or parameter calculation across datasets.
5. **Post-processing & export**
   Normalize, reorganize, and export results (e.g., Excel-ready tables).

------

## Dependencies

- MATLAB (recommended R2020a or later)
- **WITio toolbox** (Elsevier SoftwareX, SOFTX-D-20-00088)

Reference repository:
https://github.com/ElsevierSoftwareX/SOFTX-D-20-00088

------

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/groele/Witec-Matlab.git
   ```

2. Add the project directory to MATLAB path:

   ```matlab
   addpath(genpath('path_to/Witec-Matlab'));
   ```

3. Ensure WITio is correctly installed and accessible in MATLAB.

------

## Usage Notes

- Scripts are designed to be **called modularly**, but are most effective when used as a workflow
- Parameter names and file-naming conventions should be kept consistent across datasets
- For batch processing, ensure directory structures are well organized before execution

Example usage scripts are recommended to be placed in the `examples/` directory.

------

## Citation & Acknowledgment

If you use this toolbox in academic work, please acknowledge or cite appropriately.

This project is built upon the open-source WITio framework published in *SoftwareX* (Elsevier), and extends it for advanced spectroscopy data analysis workflows.

------

## License

Please refer to the original **WITio license** for data-import components.
Additional scripts in this repository are provided for academic and research use.

------

## Roadmap

- **v2.1**: Improved documentation and example workflows
- **v2.2**: Optional GUI utilities for batch analysis
- **v3.0**: Fully modularized pipeline with configuration-driven execution

