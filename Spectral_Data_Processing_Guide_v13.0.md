# WITEC数据处理与计算说明

**Code-Assisted Data Processing (CADP)**  
        **Version:** V13.0  
        **Date:** 2025.10.09  
        **Author:** Shikun Hou  
        **Email:** gro_ele@163.com  

> © 2025 Shikun Hou. All rights reserved.
>

**注意：本项目依赖库：**  
使用前务必安装 `WITio.mltbx`（Matlab toolbox）  

**注意：**使用本脚本之前务必将“xxx.wip”中的数据规范化，即备份后，将每个“xxx.wip”文件中仅保留一种数据类型，即一个位点下磁场依赖、功率依赖、温度依赖的数据，分多个文件处理，以提高数据处理效率！！！

**注意：**本项目仅对所有的单光谱数据做提取，不支持线偏振，图片，Mapping等数据类型！！！

==**注意：**命名规范可参考：Quick Copy Item.html，测试时可将其作为快速剪贴板使用，确保命名的规范性，剪贴板可随意增删，自定义❗❗❗==

---

## **1. 数据转写与文件规范格式部分（Data extraction and preprocessing）**

### **A1_WITec_Spectrum_Export_and_Classification.m**

**选择工作文件📃。**  

**功能：**  
批量导出 WITec 光谱并进行关键词分类。  

**结果文件结构：**

```
Exported_Spectra
 └── xxxxxx_20251009
     ├── 1_PL
     │    ├── a_45_deg
     │    ├── b_-45_deg
     │    ├── c_Power_dependence
     │    ├── d_Voltage_dependence
     │    ├── e_Magnetic_field
     │    ├── f_Temperature
     │    └── g_Others
     ├── 2_Raman
     ├── 3_Absorb
     ├── 4_Series
     ├── 5_Spectrum
     └── 6_Others
```
❗❗❗基于“关键字”匹配，关键字见代码。  

运行结束后会删除空文件夹！！！

---

### **A2_PolarDataManager.m**

**选择工作文件夹📂。**  

**功能：**  
整理圆偏振数据文件，**生成规范格式以便后续 DOCP 与 g 因子计算**。  

**输入文件夹：**

```
Exported_Spectra
 └── xxxxxx_20251009
     └── 1_PL
         ├── a_45_deg (数据文件夹📂！！！！！)
         └── b_-45_deg (数据文件夹📂！！！！！)
```

**结果文件结构：**
```
Exported_Spectra
 └── xxxxxx_20251009
     └── 1_PL
         ├── a_45_deg (结果文件夹📂❗❗❗)
         │   ├── -9
         │   ├── ...
         │   ├── Total.xlsx (结果文件📃❗)
         └── b_-45_deg
```

---

### **A3_FlexiExcelMerge.m**

**选择工作文件夹📂。**  

**功能：**  
功率依赖、温度依赖数据的汇总与格式化，**为后续计算做准备**。  

**输入文件夹：**

```
Exported_Spectra
 └── xxxxxx_20251009
     └── 1_PL
         ├── a_45_deg
         ├── b_-45_deg
         ├── c_Power_dependence (数据文件夹📂！！！！！)
         ├── d_Voltage_dependence 
         ├── e_Magnetic_field 
         ├── f_Temperature
         └── g_Others
```

**结果文件：**
```
Exported_Spectra
 └── xxxxxx_20251009
     └── 1_PL
         └── c_Power_dependence (结果文件夹📂❗❗❗)
             └── MergedData.xlsx (结果文件📃❗)
```



---

## **2. 数据处理与计算部分（Parameter calculation）**

### **A4_CircularPolarization_gFactor_Analysis.m**
选择工作文件📃。  

❗数据由“A2_PolarDataManager.m”预处理，规范化后运行此代码

**功能：**  
圆偏振光谱分析，**计算 DOCP 与 g 因子**。包括：  

- 背景扣除  
- 平滑处理  
- 数据截断  
- 计算 DOCP 与 g 因子  

**输入文件夹：**
```
Exported_Spectra
 └── xxxxxx_20251009
     └── 1_PL
         └── a_45_deg
             ├── -9
             ├── ...
             ├── Total.xlsx (数据文件夹📂！！！！！)
```

**结果文件结构：**
```
Exported_Spectra
 └── xxxxxx_20251009
     └── 1_PL
         └── a_45_deg
             ├── -9
             ├── ...
             ├── Total.xlsx
             └── Total_20251009 (结果文件夹📂❗❗❗)
                 ├── 0_Plots_DOCP_Total
                 ├── ...
                 └── 7_Summary_Combined_Total.xlsx (结果文件📃❗)
```

**主要参数：**

```
峰范围区间限定，复制数据到origin Pro中绘图，鼠标看下峰的范围区间！！！
startRow       = 600 % 数据起始行
endRow         = 1200 % 数据终止行
userBaseline   = 486 % 用户设定的基线值
smoothType     = 'loess' % 平滑算法
smoothParam    = 0.2 % 平滑参数，看下平滑效果确保峰形不失真
```

### **A5_SpectralDependenceAnalyzer.m**

选择工作文件📃。  

❗数据由“A3_FlexiExcelMerge.m”预处理，规范化后运行此代码

**功能：**  
功率依赖与温度依赖的光谱预处理与参数提取，包括：  

- 背景扣除  
- 平滑处理  
- 峰区截断  

**输入文件夹：**
```
Exported_Spectra
 └── xxxxxx_20251009
     └── 1_PL
         ├── a_45_deg
         ├── b_-45_deg
         ├── c_Power_dependence (数据文件夹📂！！！！！)
         ├── d_Voltage_dependence
         ├── e_Magnetic_field
         ├── f_Temperature
         └── g_Others
```
**结果文件结构：**
```
Exported_Spectra
 └── xxxxxx_20251009
     └── 1_PL
         ├── c_Power_dependence 
                 ├── 20251009_MergedData (结果文件夹📂❗❗❗)
                 	├── 0、Parameters_MergedData
                 	├── .....
                    ├── 4、Filtered_MergedData.xlsx (二维绘图数据📃❗)
                 	└── 5、Result_MergedData.xlsx
```
**主要参数：**
```
峰范围区间限定，复制数据到origin Pro中绘图，鼠标看下峰的范围区间！！！
startRow       = 600 % 数据起始行
endRow         = 1200 % 数据终止行
userBaseline   = 486 % 用户设定的基线值
smoothType     = 'loess' % 平滑算法
smoothParam    = 0.2 % 平滑参数，看下平滑效果确保峰形不失真
```

---

