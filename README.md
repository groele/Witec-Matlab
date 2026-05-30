<div align="center">

# Witec-Matlab

**面向 WITec Raman / PL 光谱数据的 MATLAB 批处理与偏振分析工作流**  
*MATLAB toolbox for WITec Raman / PL spectroscopy processing, polarization-resolved analysis, batch metrics, and figure-ready export.*

![Type](https://img.shields.io/badge/type-MATLAB%20Toolbox-blue?style=flat-square)
![Domain](https://img.shields.io/badge/domain-Raman%20%2F%20PL%20spectroscopy-green?style=flat-square)
![Status](https://img.shields.io/badge/status-stable-purple?style=flat-square)
![Version](https://img.shields.io/badge/version-v2.0.0-orange?style=flat-square)
![License](https://img.shields.io/badge/license-academic--use-yellow?style=flat-square)

Part of **ResearchFlow Lab** — a local-first research productivity ecosystem for literature, manuscripts, data, and scientific visualization.

</div>

---

## 01. Overview

**Witec-Matlab** is a MATLAB-based spectroscopy data-processing workflow for WITec Raman / photoluminescence datasets. It focuses on batch extraction, spectrum organization, polarization-resolved analysis, spectral metrics, Excel-ready output, and long-term reproducibility for experimental condensed-matter and optical spectroscopy research.

**Witec-Matlab** 是一个面向 WITec Raman / PL 光谱数据的 MATLAB 分析工具箱，服务于低温光致发光、偏振分辨光谱、磁场依赖光谱、栅压依赖光谱和批量数据导出等实验流程。

This project builds upon and extends the open-source **WITio** framework, adding higher-level scripts and unified analysis pipelines tailored for practical research use.

---

## 02. Why this project exists

Spectroscopy experiments often generate large sets of files across polarization angles, magnetic fields, gate voltages, spatial maps, temperatures, and repeated measurements. Without a consistent pipeline, analysis becomes fragile: file naming is inconsistent, Excel exports are difficult to merge, polarization metrics are recomputed manually, and figure data are hard to reproduce.

Witec-Matlab provides a structured analysis workflow from raw WITec data to figure-ready tables.

核心目标：

- Standardize WITec spectral data import and classification.
- Support batch processing across many experimental conditions.
- Provide reusable modules for polarization-resolved PL/Raman analysis.
- Export clean Excel-ready data for plotting and manuscript figures.
- Preserve a reproducible analysis path for long-term research projects.

---

## 03. Key features

| Module | What it does | 中文说明 |
|---|---|---|
| WITec Import | Reads and organizes WITec spectroscopy outputs through WITio-based routines | 基于 WITio 读取并整理 WITec 光谱输出 |
| Spectrum Classification | Classifies spectra by filename keywords, condition labels, or experiment folders | 按文件名关键词、实验条件和目录结构分类光谱 |
| Polarization Data Manager | Organizes polarization-resolved datasets for angle-dependent analysis | 管理偏振分辨数据，支持角度依赖分析 |
| Batch Spectral Analysis | Computes peak positions, intensities, ratios, and other spectral parameters | 批量计算峰位、峰强、比值和光谱参数 |
| Normalization Tools | Provides reusable normalization and column-splitting utilities | 提供归一化、奇偶列拆分等通用工具 |
| Excel Merge | Merges processed outputs into clean Excel-ready tables | 将批处理结果合并为可直接绘图的 Excel 表格 |
| File Organization | Moves or groups files by keywords and experimental conditions | 根据关键词和实验条件整理文件 |
| Reproducible Workflow | Keeps analysis scripts modular and traceable | 保持脚本模块化和分析过程可追溯 |

---

## 04. Product philosophy

Witec-Matlab follows four design principles:

1. **Data provenance** — raw, intermediate, and exported data should remain traceable.
2. **Batch first** — analysis should scale from a single spectrum to large experimental folders.
3. **Figure-ready output** — exported tables should directly support MATLAB, Origin, Python, or manuscript plotting.
4. **Research-specific metrics** — scripts should encode real spectroscopy workflows, not generic file conversion only.

---

## 05. Architecture

```text
Raw WITec Data
    ↓
WITio-based Import
    ↓
Spectrum Export and Classification
    ↓
Polarization / Condition Data Manager
    ↓
Batch Spectral Analysis
├── peak position
├── peak intensity
├── spectral ratio
├── polarization metrics
└── custom experimental parameters
    ↓
Post-processing
├── normalization
├── odd/even column splitting
├── keyword-based file grouping
└── Excel merge
    ↓
Figure-ready Tables / Manuscript Data
```

---

## 06. Quick start

Requirements:

| Requirement | Recommendation |
|---|---|
| MATLAB | R2020a or later recommended |
| WITio | Required for WITec data import |
| OS | Windows recommended for WITec-origin workflows |

Clone and add to path:

```bash
git clone https://github.com/groele/Witec-Matlab.git
```

In MATLAB:

```matlab
addpath(genpath('path_to/Witec-Matlab'));
```

Make sure the WITio toolbox is installed and available in the MATLAB path.

---

## 07. Recommended workflow

```text
Export WITec raw data → Classify spectra by condition
                      → Build polarization / bias / field data groups
                      → Run batch spectral metrics
                      → Normalize and merge Excel outputs
                      → Plot publication-ready figures
```

Typical use cases:

- Polarization-resolved PL/Raman analysis.
- Gate-voltage-dependent exciton/trion spectral tracking.
- Magnetic-field-dependent Zeeman or g-factor related datasets.
- Batch export of figure source data.
- Reorganization of large experimental folders.

---

## 08. Project structure

```text
Witec-Matlab
├── 1. Data extraction and preprocessing/
│   ├── A1_WITec_Spectrum_Export_and_Classification_Unified.m
│   ├── A2_PolarDataManager_Unified.m
│   └── A3_FlexiExcelMerge_Unified.m
├── 2. Parameter calculation/
│   ├── A5_Spectral_Batch_Analysis_Unified.m
│   ├── B1_FileMoverByKeyword_Unified.m
│   ├── B2_Normalization.m
│   └── B3_SplitOddEvenColumns.m
├── docs/
├── examples/
└── README.md
```

---

## 09. Roadmap

- [ ] Add documented example datasets
- [ ] Add standard templates for DOCP, DOLP, and polarization fitting outputs
- [ ] Add Zeeman / g-factor analysis examples
- [ ] Add batch plotting templates for manuscript figures
- [ ] Improve configuration-driven execution
- [ ] Add optional GUI utilities for batch analysis
- [ ] Add test datasets for regression checking

---

## 10. Privacy and data ownership

Witec-Matlab is a local MATLAB workflow. Input data, processed files, and exported tables remain on the user's machine unless manually uploaded elsewhere. Example datasets should be anonymized before public release.

---

## 11. Related projects

- **ResearchFlow Companion** — research workflow operating system
- **Scientific Color Lab** — scientific color and visualization workspace
- **ManuGuide** — Microsoft Word manuscript formatting and style checker
- **PaperPilot Pro** — academic search and publisher-page enhancement
- **ClipNote** — browser-native quick notes and Markdown capture

---

## 12. License and acknowledgment

This project builds upon the open-source **WITio** framework. Please refer to the original WITio license for data-import components.

Additional scripts are provided for academic and research use.

Developed by **Shikun Hou / groele**.
