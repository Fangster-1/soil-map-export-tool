# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

县级土壤类型图**多布局批量出图工具**，基于 ArcGIS Pro 的 `arcpy.mp` 自动生成布局工程并批量导出高清地图。整个程序是单文件应用：`soil_map_export_tool20260616.py`（约 1791 行），无第三方框架，仅在出图时延迟加载 `arcpy` 和 `openpyxl`。

## 运行方式

必须使用 ArcGIS Pro 的 Python 环境（含 `arcpy`），即全局配置的 `D:\GD\arcgispro_clone\python.exe`。

```powershell
# 图形界面（无参数时自动启动 Tkinter GUI）
D:\GD\arcgispro_clone\python.exe soil_map_export_tool20260616.py

# 命令行：第一步，生成多布局工程 .aprx
D:\GD\arcgispro_clone\python.exe soil_map_export_tool20260616.py --no-gui --mode generate `
  --county 宾川县 --mxd <图框.mxd> --source-layer <data.gdb\图层名> `
  --color-table <配色表.csv> --stats-table <统计表.xlsx> --output-dir <输出目录>

# 命令行：第二步，按布局批量导出 JPG
D:\GD\arcgispro_clone\python.exe soil_map_export_tool20260616.py --no-gui --mode export `
  --output-dir <输出目录> --dpi 600
```

`一键出图.bat` 是给最终用户的双击入口，但**当前调用的文件名 `soil_map_export_tool.py` 与实际文件 `soil_map_export_tool20260616.py` 不一致**，部署时需重命名其一。

无测试、无构建、无 lint 配置；`arcpy` 仅在 ArcGIS Pro 环境可用，开发机上无法直接运行出图流程。

## 两步工作流（核心设计）

工具刻意将出图拆成两步，中间留出人工微调窗口：

1. **generate**（`export_with_arcgis_pro` 中 `mode != "export"` 分支）：复制 ArcGIS Pro 内置空白工程 `Blank.aprx`，逐土类导入 MXD 图框模板生成布局，应用配色，保存为 `自动化出图工作空间.aprx`，同时导出配色对照 Excel。用户随后在 ArcGIS Pro 中手动调整版面并保存。
2. **export**（`mode == "export"`）：读取输出目录中的 `.aprx`，遍历所有布局导出为 JPG。

## 数据流与三层土壤分类

字段映射默认 `TL`（土类）/`YL`（亚类）/`TS`（土属），可在 GUI「字段映射」或 `--soil-field/--subclass-field/--genus-field` 覆盖。**土属字段留空时降级为按亚类出图**（`render_genus` 在 `export_with_arcgis_pro` 中据此切换）。

出图所用土类是 **统计表与图层的交集**（`soil_classes = [s for s in stats_by_soil if s in layer_soils]`）。每个土类生成一张分布图，外加一张全域土壤类型分布图（按亚类配色）。

`build_type_mismatch_report` + `confirm_type_mismatch` 会在统计表与图层的土类/亚类/土属不一致时报告并要求确认（`--confirm-mismatch` 弹窗 / `--continue-on-mismatch` 直接继续）。

## 关键技术约束

- **单图层唯一值渲染**：每个土类只创建 **1 个要素图层**，通过 `CIMUniqueValueRenderer` 在同一图层内按字段值区分颜色（见 `add_single_layer_with_unique_values`、`apply_cim_symbology`）。这是相对旧版「每个分类值一个图层」的核心重构。
- **CIM 而非 `layer.symbology`**：配色一律走 `layer.getDefinition("V2")` / `setDefinition` 的纯 JSON CIM 路径。原因：`importDocument` 多布局场景下原始 Layer 对象指针会失效，赋值 `layer.symbology` 会报 "attribute 'symbology' is not supported"。`_refresh_layer` 通过 `mf.map.listLayers()` 重新捞取有效图层引用来规避。
- **配色逻辑**：`_palette_pick_order` 按特定规则从配色表选色号（≤4 个值从 4 号倒序取，否则从值数量倒序取），不够时由 `_select_palette_colors` 插值/明暗变化补足；`--color-direction` 控制深浅渐变方向。
- **零依赖 Excel 兜底**：`openpyxl` 缺失时 `read_stats_table` 回退到 `_parse_xlsx_xml`，直接用 `zipfile` + `ElementTree` 解析 .xlsx（Office Open XML），仅读取，写出对照表仍需 openpyxl。
- **面积换算**：统计表「面积」字段原始单位为平方米，读取时即换算为亩（`SQM_TO_MU`）；Excel 对照表再按 `--area-unit`（亩/万亩）折算。

## 路径配置

文件顶部「用户路径配置区」（`CFG_*` 常量，基于 `PROJECT_ROOT = 脚本父目录的父目录`）定义各类输入/输出的默认目录，是 GUI 字段和 CLI 参数的默认值。部署到新环境时按实际目录结构修改这些常量。

## 代码组织（单文件内分九部分）

环境初始化 → 配色表读取与颜色计算 → 统计表读取 → 数据聚合 → ArcGIS Pro 布局操作 → Excel 对照表导出 → 主出图工作流 → Tkinter GUI → CLI 入口。每部分以 `# ===` 注释分隔。

所有用户可见输出（print、报错、GUI、注释）均为中文。
