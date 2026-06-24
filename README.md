# 县级土壤类型图多布局批量出图工具

基于 **ArcGIS Pro / arcpy** 的县级土壤类型图自动化批量出图工具。一次性为每个土类生成一张分布图、外加一张全域土壤类型分布图，自动配色、自动排版、批量导出高清 JPG，并同步输出配色对照 Excel。

适用场景：第三次全国土壤普查（三普）等成果中，县级土壤类型系列图的批量制图与出图。

---

## 目录

- [功能特性](#功能特性)
- [运行环境](#运行环境)
- [快速开始](#快速开始)
- [核心工作流（两步法）](#核心工作流两步法)
- [输入数据准备](#输入数据准备)
- [图形界面（GUI）使用](#图形界面gui使用)
- [输出成果](#输出成果)
- [技术架构与实现细节](#技术架构与实现细节)
- [打包为 EXE](#打包为-exe)
- [路径配置](#路径配置)
- [常见问题与排错](#常见问题与排错)
- [文件清单](#文件清单)

---

## 功能特性

- **逐土类批量成图**：以「统计表 ∩ 图层」的土类为准，每个土类生成一张独立布局，外加一张全域亚类配色分布图。
- **三级土壤分类**：支持「土类 / 亚类 / 土属」三层字段；**土属字段留空时自动降级为按亚类出图**。
- **智能配色**：从配色表按规则取色，颜色不足时自动插值 / 明暗渐变补足，可选深→浅 / 浅→深方向。
- **单图层唯一值渲染**：每个土类仅 1 个要素图层，通过 CIM 唯一值渲染器在同一图层内按字段值着色（性能与稳定性优于「一值一图层」）。
- **配色对照 Excel**：自动导出多表单 Excel，每个土类一张表 + 全域亚类表，单元格按实际出图颜色填充。面积可按「亩 / 万亩」显示：亩为整数、万亩保留两位小数，均不带千分位分隔符。
- **数据一致性校验**：出图前比对统计表与图层的土类 / 亚类 / 土属差异，不一致时弹窗 / 命令行确认。
- **记忆界面配置**：GUI 会自动记住上次填写的所有路径与字段，下次启动自动回填，无需重复设置。
- **两步工作流**：生成布局工程后留出人工微调窗口，再批量导出，兼顾自动化与版面精修。
- **零依赖兜底**：缺少 `openpyxl` 时用内置 `zipfile + ElementTree` 直接解析 `.xlsx` 读取统计表。
- **图形界面**：双击启动 Tkinter GUI，操作直观。
- **可打包为单文件 EXE**：含自定义图标、内嵌主程序、可配置 Python 路径。

---

## 运行环境

| 项目 | 要求 |
|------|------|
| 操作系统 | Windows 10 / 11 |
| **ArcGIS Pro** | **必须安装**（提供 `arcpy`，本工具核心依赖；本项目基于 3.6.3 验证） |
| Python | 使用 ArcGIS Pro 自带 / 克隆环境的 Python（含 `arcpy`），如 `D:\GD\arcgispro_clone\python.exe` |
| 第三方库 | `openpyxl`（缺失时自动尝试安装；仅写出 Excel 对照表必须）。打包 EXE 需 `PyInstaller`、生成图标需 `Pillow` |

> ⚠️ **关键限制**：`arcpy` 绑定整个 ArcGIS Pro 安装与许可证，**无法脱离 ArcGIS Pro 独立运行**。无论是源码、批处理还是 EXE，目标电脑都必须装有 ArcGIS Pro。

---

## 快速开始

### 方式一：双击 EXE（推荐给最终用户）

1. 将 `出图工具_发布版(单文件)\土壤类型出图工具.exe` 拷到目标电脑（已装 ArcGIS Pro）。
2. 双击运行。首次运行会自动查找 ArcGIS Pro 的 Python；找不到时弹窗让你浏览选择 `python.exe`，选定后记忆到同目录 `出图工具配置.txt`。
3. 在界面中设置数据路径，按两步法出图。

### 方式二：批处理启动源码

双击 `一键出图.bat`（内置 ArcGIS Pro Python 自动探测）。

### 方式三：直接用 Python 运行源码

```powershell
D:\GD\arcgispro_clone\python.exe soil_map_export_tool.py
```

无参数时自动启动 GUI。

---

## 核心工作流（两步法）

工具刻意将出图拆成两步，中间留出人工微调版面的窗口：

```
┌─────────────────┐      人工在 ArcGIS Pro       ┌─────────────────┐
│  第一步 generate │  ───  中逐一微调版面  ───▶  │  第二步 export   │
│  生成多布局工程   │      （图例/比例尺/排版）     │  批量导出 JPG    │
│  .aprx + 对照Excel│                              │                 │
└─────────────────┘                              └─────────────────┘
```

1. **第一步 generate**：复制 ArcGIS Pro 内置空白工程 `Blank.aprx`，逐土类导入 MXD 图框模板生成布局、应用配色，保存为 `自动化出图工作空间.aprx`，同时导出配色对照 Excel。
2. **人工微调**：在 ArcGIS Pro 中打开该工程，逐个布局调整图例、比例尺、版面，保存。
3. **第二步 export**：读取输出目录中的 `.aprx`，遍历所有布局批量导出为高清 JPG。

---

## 输入数据准备

### 1. 图框模板（.mxd）

ArcGIS 的地图文档，作为每张图的版面模板（A3 图框、图例、比例尺、指北针、图名文本等）。工具会：
- 自动选取面积最大的地图框作为主底图框（忽略位置示意小图框）；
- 自动定位「县/市名称」文本元素并替换为对应图名（如「xx县红壤分布图」）。

### 2. 土壤要素图层（.gdb 中的要素类）

面要素图层，需包含三层分类字段。默认字段名：

| 含义 | 默认字段 | 说明 |
|------|----------|------|
| 土类 | `TL` | 必需 |
| 亚类 | `YL` | 必需 |
| 土属 | `TS` | 可选，留空则按亚类出图 |

字段名可在 GUI「字段映射」或 CLI 参数中覆盖。

### 3. 配色表（.csv）

定义每个土类的色带。**必需列**：`土类`、`R`、`G`、`B`；可选列：`色标序号`（同一土类内排序用）。编码自动识别 GBK / UTF-8。

示例：

```csv
土类,色标序号,R,G,B
红壤,1,255,235,225
红壤,2,240,180,160
红壤,3,210,120,95
红壤,4,180,70,55
黄壤,1,250,245,210
黄壤,2,230,215,140
```

> 配色规则：≤4 个分类值时从 4 号色倒序取，否则从「值数量」号倒序取；颜色不够时自动插值或同色系明暗补足。`色带方向` 控制深→浅或浅→深。

### 4. 平差面积统计表（.xlsx）

**必需列**：`土类`、`亚类`、`土属`、`面积`。其中 **`面积` 列原始单位为平方米**，工具读取时自动换算为亩（再按需折算万亩）。支持「土类 / 亚类」纵向合并单元格（自动向下填充）。相同「亚类+土属」组合会自动累加去重。

---

## 图形界面（GUI）使用

启动后界面字段：

| 字段 | 说明 |
|------|------|
| A3 图框底图 | 选择 `.mxd` 模板 |
| 土壤要素图层 | 选择 `.gdb` 目录并输入图层名 |
| 县/市名称 | 如「xx县」，留空则从数据路径自动识别 |
| 配色表（固定） | 选择配色 `.csv` |
| 平差面积统计 | 选择统计 `.xlsx` |
| 成果输出目录 | 布局工程、JPG、对照表均输出至此 |
| 剔除土类 | 逗号分隔，默认「建设用地,河流水面」 |
| 面积单位 | 亩 / 万亩（对照表显示用） |
| 色带方向 | 深→浅 / 浅→深 |
| 字段映射 | `土类,亚类,土属`，留空土属则按亚类出图 |

底部两个按钮分别对应第一步（生成布局工程）和第二步（批量出图）。

### 界面配置记忆

GUI 会在**点击出图按钮时**及**关闭窗口时**自动保存当前所有字段（各路径、县名、剔除土类、字段映射、面积单位、色带方向），下次启动自动回填。

配置文件位置（用户级，跨次稳定，单文件 EXE 也适用）：

```
%LOCALAPPDATA%\土壤类型出图工具\界面配置.json
```

如需重置为默认值，删除该文件即可。注意：它与解释器路径配置（exe 同目录的 `出图工具配置.txt`）是两个不同的文件，各管各的。

---

## 输出成果

第一步（generate）在输出目录生成：

- `自动化出图工作空间.aprx` —— 含所有布局的 ArcGIS Pro 工程
- `{县名}土壤类型配色表.xlsx` —— 多表单配色对照表（面积按所选单位显示：亩取整数、万亩保留两位小数，不带千分位分隔符）

第二步（export）生成：

- 每个布局一张 `.jpg`（文件名即布局名，如 `xx县红壤分布图.jpg`）

---

## 技术架构与实现细节

整个程序为**单文件应用** `soil_map_export_tool.py`（约 1800 行），仅在实际出图时延迟加载 `arcpy` / `openpyxl`。代码按九部分组织：环境初始化 → 配色读取与颜色计算 → 统计表读取 → 数据聚合 → ArcGIS Pro 布局操作 → Excel 对照表导出 → 主出图工作流 → Tkinter GUI → CLI 入口。

### 关键技术点

- **单图层唯一值渲染**：每个土类只创建 1 个要素图层，通过 `CIMUniqueValueRenderer`（纯 JSON CIM）在同一图层内按字段值区分颜色（`add_single_layer_with_unique_values`、`apply_cim_symbology`）。未配色的字段值用全透明符号处理。

- **走 CIM 而非 `layer.symbology`**：配色一律通过 `layer.getDefinition("V2")` / `setDefinition` 的纯 CIM 路径完成。原因：`importDocument` 多布局场景下，原始 `Layer` 对象指针会失效，直接赋值 `layer.symbology` 会报 `"attribute 'symbology' is not supported"`。`_refresh_layer` 通过 `mf.map.listLayers()` 重新捞取有效图层引用来规避。

- **配色算法**：`_palette_pick_order` 按规则选色号，`_select_palette_colors` 在颜色不足时插值（多锚点）或明暗变化（单锚点）补足，`--color-direction` 控制渐变方向。

- **零依赖 Excel 兜底**：`openpyxl` 缺失时 `read_stats_table` 回退到 `_parse_xlsx_xml`，用 `zipfile + ElementTree` 直接解析 Office Open XML（仅读取；写对照表仍需 openpyxl）。

- **面积换算与显示**：统计表面积原始单位为平方米，读取即换算为亩（`SQM_TO_MU = 1/666.6667`）；对照表再按 `--area-unit` 折算。Excel 单元格数字格式：亩用 `"0"`（整数），万亩用 `"0.00"`（两位小数），均不使用千分位分隔符。底层仍存完整数值，仅显示按格式四舍五入。

- **出图土类取交集**：`soil_classes = [s for s in stats_by_soil if s in layer_soils]`，仅对统计表与图层都存在的土类出图。

- **GUI 线程安全**：出图在后台线程执行，避免界面卡死；类型不一致的确认弹窗通过 `root.after` 调度回主线程执行（Tkinter 非线程安全，子线程直接弹窗会与主线程 `mainloop` 死锁）。

- **界面配置记忆**：`load_gui_config` / `save_gui_config` 将界面字段以 JSON 持久化到 `%LOCALAPPDATA%\土壤类型出图工具\界面配置.json`。选用用户目录而非脚本目录，是因为单文件 EXE 运行时主程序位于临时解压目录（`_MEIPASS`），脚本目录无法跨次稳定保存。

### EXE 启动器架构

`launcher.py` 是一个**不含 arcpy 的轻量启动器**，被打包为单文件 EXE：

1. 解析 ArcGIS Pro Python 路径：**配置文件 → 自动探测 → 弹窗浏览选择**（结果记忆到 `出图工具配置.txt`）；
2. 把内嵌主程序（`--add-data` 嵌入）**复制到一个纯净临时目录**后再交给解释器运行；
3. 以**净化过的环境**启动子进程（PATH 去除 `_MEIPASS`，清除一切值指向 `_MEIPASS` 的变量如 `TCL_LIBRARY`/`TK_LIBRARY`，统一 `PYTHONIOENCODING=utf-8`）；
4. 子进程输出重定向到同目录 `出图日志.txt`（UTF-8 带 BOM），便于排查。

这样设计的好处：启动器自身无 arcpy 依赖，打包稳定、体积小；真正的出图仍在 ArcGIS Pro 原生环境中运行，规避了「冻结 arcpy」的兼容性问题。

> **为什么必须复制到纯净目录**：PyInstaller 的临时解压目录 `_MEIPASS` 里含有「打包那台机器」的二进制（`pythonXX.dll`、`_tkinter.pyd` 等）。若让目标机的 ArcGIS Python 直接运行 `_MEIPASS` 内的脚本，其 `sys.path[0]` 即 `_MEIPASS`，`import _tkinter` 会误加载这里的扩展并拉起配套的 `pythonXX.dll`，与正在运行的 ArcGIS `pythonXX.dll` 冲突，报 `Module use of pythonXX.dll conflicts with this version of Python`。复制到不含二进制的纯净目录即可彻底规避。

---

## 打包为 EXE

环境需 `PyInstaller`（打包）和 `Pillow`（生成图标），均装在 ArcGIS Pro Python 中。

**一键打包**：双击 `重新打包exe.bat`（自动同步内嵌副本 → PyInstaller 打包 → 拷贝到发布目录）。

**手动打包**：

```powershell
D:\GD\arcgispro_clone\python.exe -m PyInstaller --noconfirm --clean --onefile --noconsole `
  --icon "app.ico" `
  --add-data "soil_map_export_tool.py;." `
  --name "土壤类型出图工具" launcher.py
```

产物在 `dist\土壤类型出图工具.exe`。更换图标：替换 `app.ico` 后重新打包。

---

## 路径配置

源码顶部「用户路径配置区」（`CFG_*` 常量，基于 `PROJECT_ROOT = 脚本父目录的父目录`）定义各类输入/输出默认目录，是 GUI 字段的默认值：

| 常量 | 含义 |
|------|------|
| `CFG_FRAMES_BASE_DIR` | 图框 `.mxd` 父目录 |
| `CFG_SOIL_GDB_BASE_DIR` | 土壤 `.gdb` 父目录 |
| `CFG_COLOR_TABLE` | 配色表 `.csv` 固定路径 |
| `CFG_STATS_BASE_DIR` | 统计表 `.xlsx` 目录 |
| `CFG_OUTPUT_BASE_DIR` | 成果输出目录 |

部署到新环境时按实际目录结构修改这些常量。EXE 用户无需改源码——在界面里直接选路径即可；Python 解释器路径则由 `出图工具配置.txt` 控制。

---

## 常见问题与排错

| 现象 | 原因 / 解决 |
|------|-------------|
| `ModuleNotFoundError: No module named 'arcpy'` | 用了普通 Python。必须用 ArcGIS Pro 的 Python（含 arcpy）。EXE 用户检查 `出图工具配置.txt` 指向是否正确。 |
| `Module use of python3xx.dll conflicts with this version of Python` | 换机器后出现：旧版内嵌脚本从 PyInstaller 临时目录 `_MEIxxxx` 运行，混入了打包机的 `_tkinter.pyd`/`pythonXX.dll`，被目标机 ArcGIS Python 误加载。新版已修复（脚本改到纯净临时目录运行并净化子进程环境）。请使用最新 EXE。 |
| `Can't find a usable init.tcl ... version conflict for package "Tcl"` | 换机器后出现：PyInstaller 注入的 `TCL_LIBRARY`/`TK_LIBRARY` 指向打包机的 Tcl 数据（如 8.6.13），被目标机 ArcGIS Python（如 8.6.15）继承导致版本冲突。新版已修复（子进程环境清除一切指向 `_MEIPASS` 的变量）。请使用最新 EXE。 |
| 设置好数据后一直运行、无结果 | 旧版在后台线程弹 Tkinter 确认窗导致死锁，新版已修复（确认窗回主线程）。 |
| 弹窗「未找到解释器」 | 本机未装 ArcGIS Pro，或在弹窗中手动选择其 `python.exe`。 |
| 双击 EXE 无反应/闪退 | 查看同目录 `出图日志.txt` 末尾报错。 |
| 提示统计表与图层类型不一致 | 正常校验。确认无误后选择「继续」，或用 `--continue-on-mismatch`。 |
| 某土类全灰 / 透明 | 配色表中缺该土类颜色记录，或字段值与统计表对不上。 |
| 杀毒软件拦截 EXE | 单文件 PyInstaller 程序偶被误报，加信任放行即可。 |
| 找不到空白模板 `Blank.aprx` | 确认 ArcGIS Pro 安装完整，或在 `get_blank_aprx()` 中手动指定路径。 |

排错日志：EXE 运行输出在 `出图日志.txt`；命令行运行时直接打印到控制台。

---

## 文件清单

```
soil-map-export-tool/
├─ soil_map_export_tool.py    主程序（源码，单文件应用，约 1800 行）
├─ launcher.py                EXE 启动器（解析解释器 + 运行内嵌主程序）
├─ app.ico                    应用图标
├─ 一键出图.bat                源码启动入口（自动探测 ArcGIS Pro Python）
├─ 重新打包exe.bat             一键重新打包 EXE
├─ CLAUDE.md                  面向 AI 协作者的项目说明
└─ README.md                  本文档
```

> EXE 发布包（`SoilMapExportTool.exe` + `README_Usage.txt`）通过 [GitHub Releases](https://github.com/Fangster-1/soil-map-export-tool/releases) 下载，不纳入仓库。

---

## 作者

方庆坪 · 微信同号：19988312343
