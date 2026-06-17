#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
土壤类型出图工具 —— 轻量启动器（单文件内嵌版）。

本启动器本身不依赖 arcpy，职责：
  1. 解析本机含 arcpy 的 ArcGIS Pro Python 解释器路径，顺序为：
     配置文件 → 自动探测 → 弹窗让用户浏览选择（并记忆到配置文件）；
  2. 用该解释器运行内嵌的主程序 soil_map_export_tool.py（GUI）；
  3. 运行输出写入 exe 同目录的 出图日志.txt，便于排查。

主程序已被打包进 exe，无需随附 .py 文件。
用户可在 exe 同目录的「出图工具配置.txt」中查看/修改解释器路径，
删除该文件即可重新自动探测或重新选择。
"""

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# 内嵌的主出图脚本文件名
SCRIPT_NAME = "soil_map_export_tool.py"

# 解释器路径配置文件名（位于 exe 同目录，用户可编辑）
CONFIG_NAME = "出图工具配置.txt"

# 候选 ArcGIS Pro Python 解释器（自动探测，按优先级）
PYTHON_CANDIDATES = [
    r"D:\GD\arcgispro_clone\python.exe",
    r"C:\Program Files\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe",
    os.path.expandvars(
        r"%LOCALAPPDATA%\Programs\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe"
    ),
    os.path.expandvars(
        r"%PROGRAMFILES%\ArcGIS\Pro\bin\Python\envs\arcgispro-py3\python.exe"
    ),
]


def app_dir() -> Path:
    """exe（或脚本）所在的真实目录，用于放配置、日志、输出。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def script_path() -> Path:
    """内嵌主程序脚本的实际路径（打包后位于解压目录 _MEIPASS）。"""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", app_dir())) / SCRIPT_NAME
    return app_dir() / SCRIPT_NAME


def show_error(msg: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("启动失败", msg)
        root.destroy()
    except Exception:
        try:
            (app_dir() / "启动错误.txt").write_text(msg, encoding="utf-8")
        except Exception:
            pass
        print(msg, file=sys.stderr)


# ── 解释器路径：配置文件读写 ──────────────────────────────────────────────

def config_file() -> Path:
    return app_dir() / CONFIG_NAME


def read_config_python() -> str:
    f = config_file()
    if not f.exists():
        return ""
    try:
        for line in f.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#"):
                return line.strip('"')
    except Exception:
        pass
    return ""


def write_config_python(path: str) -> None:
    try:
        config_file().write_text(
            "# ArcGIS Pro 的 python.exe 路径（必须含 arcpy）。\n"
            "# 修改此路径可指定其他电脑上的解释器；删除本文件可重新自动探测/选择。\n"
            f"{path}\n",
            encoding="utf-8",
        )
    except Exception:
        pass


def autodetect_python() -> str:
    for c in PYTHON_CANDIDATES:
        if c and Path(c).exists():
            return c
    return ""


def ask_python_from_user(reason: str) -> str:
    """弹窗让用户浏览选择 python.exe。"""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "选择 ArcGIS Pro 的 Python",
            reason
            + "\n\n请在接下来的窗口中选择 ArcGIS Pro 的 python.exe，"
            "通常位于：\nC:\\Program Files\\ArcGIS\\Pro\\bin\\Python\\envs\\arcgispro-py3\\python.exe",
        )
        chosen = filedialog.askopenfilename(
            title="选择 ArcGIS Pro 的 python.exe",
            filetypes=[("Python 解释器", "python*.exe"), ("所有文件", "*.*")],
        )
        root.destroy()
        return chosen.strip() if chosen else ""
    except Exception as e:
        show_error(f"无法弹出选择窗口：{e}")
        return ""


def resolve_python() -> str:
    """配置文件 → 自动探测 → 用户选择，依次解析解释器路径。"""
    # 1) 配置文件
    cfg = read_config_python()
    if cfg and Path(cfg).exists():
        return cfg

    # 2) 自动探测
    auto = autodetect_python()
    if auto:
        write_config_python(auto)
        return auto

    # 3) 让用户选择
    reason = (
        "未自动找到含 arcpy 的 ArcGIS Pro Python。"
        if not cfg
        else f"配置文件中的解释器路径已失效：\n{cfg}"
    )
    chosen = ask_python_from_user(reason)
    if chosen and Path(chosen).exists():
        write_config_python(chosen)
        return chosen
    return ""


def stage_script(src: Path) -> Path:
    """把内嵌脚本复制到一个干净的临时目录后再交给子进程运行。

    关键：PyInstaller 的 _MEIPASS 解压目录里塞满了「打包那台机器」的
    二进制（python3xx.dll、_tkinter.pyd 等）。若直接让 ArcGIS 的 python
    运行 _MEIPASS 内的脚本，子进程 sys.path[0] 即 _MEIPASS，import _tkinter
    会优先加载这里的 _tkinter.pyd，进而拉起与之配套的 python3xx.dll，
    与正在运行的 ArcGIS python3xx.dll 冲突，报
    "Module use of python3xx.dll conflicts with this version of Python"。
    复制到不含任何二进制的纯净目录后，子进程便会正确加载 ArcGIS 自身的扩展。
    """
    work = Path(tempfile.mkdtemp(prefix="soilmap_run_"))
    target = work / SCRIPT_NAME
    shutil.copy2(str(src), str(target))
    return target


def child_env() -> dict:
    """为子进程构造净化环境，避免误用打包进来的运行时 / DLL / Tcl-Tk 数据。

    PyInstaller 的运行钩子会在启动器进程里注入一批指向 _MEIPASS 的环境变量
    （TCL_LIBRARY / TK_LIBRARY 等）和把 _MEIPASS 加进 PATH。子进程（目标机的
    ArcGIS Python）若继承这些，会去打包机的目录找 DLL / Tcl 数据，导致：
      - "Module use of pythonXX.dll conflicts..."（python DLL 冲突）
      - "Can't find a usable init.tcl ... version conflict for package Tcl"
        （Tcl/Tk 版本冲突，如打包机 8.6.13 vs 目标机 8.6.15）
    因此：从 PATH 去除 _MEIPASS，并清除一切「值指向 _MEIPASS」的环境变量。
    """
    env = dict(os.environ)
    mei = getattr(sys, "_MEIPASS", "")
    mei_norm = os.path.normcase(str(Path(mei)).rstrip("\\")) if mei else ""

    # 1) PATH 去除 _MEIPASS
    if mei_norm:
        parts = [
            p for p in env.get("PATH", "").split(os.pathsep)
            if p and os.path.normcase(str(Path(p)).rstrip("\\")) != mei_norm
        ]
        env["PATH"] = os.pathsep.join(parts)

        # 2) 清除一切值指向 _MEIPASS 的变量（自动涵盖 TCL_LIBRARY / TK_LIBRARY 等）
        for k in list(env.keys()):
            if k.upper() == "PATH":
                continue
            v = env.get(k, "")
            if v and mei_norm in os.path.normcase(v):
                env.pop(k, None)

    # 3) 兜底显式移除已知会污染子进程的变量
    for k in ("PYTHONPATH", "PYTHONHOME", "TCL_LIBRARY", "TK_LIBRARY", "TIX_LIBRARY"):
        env.pop(k, None)

    # 统一子进程输出为 UTF-8，避免中文 Windows 下日志按 GBK 输出与头部 UTF-8 混编乱码
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def main() -> int:
    here = app_dir()
    script = script_path()
    if not script.exists():
        show_error(f"未找到内嵌主程序脚本：\n{script}\n打包可能不完整，请重新生成 exe。")
        return 1

    pyexe = resolve_python()
    if not pyexe:
        show_error(
            "未指定有效的 ArcGIS Pro Python 解释器，程序无法运行。\n"
            "本工具需要本机已安装 ArcGIS Pro。\n"
            f"你也可以手动编辑配置文件后重试：\n{config_file()}"
        )
        return 1

    # 复制到纯净目录运行（仅打包后必要，开发态也无害）
    work_dir = None
    try:
        run_script = stage_script(script)
        work_dir = run_script.parent
    except Exception as e:
        show_error(f"准备运行脚本失败：{e}")
        return 1

    log_path = here / "出图日志.txt"
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        with log_path.open("w", encoding="utf-8-sig") as log:
            log.write(f"解释器: {pyexe}\n脚本: {run_script}\n{'=' * 60}\n")
            log.flush()
            proc = subprocess.run(
                [pyexe, str(run_script)],
                cwd=str(here),
                stdout=log,
                stderr=subprocess.STDOUT,
                creationflags=creationflags,
                env=child_env(),
            )
    except Exception as e:
        show_error(f"启动主程序失败：{e}\n\n解释器：{pyexe}\n脚本：{run_script}")
        return 1
    finally:
        if work_dir is not None:
            shutil.rmtree(str(work_dir), ignore_errors=True)

    if proc.returncode != 0:
        show_error(
            f"程序异常退出（代码 {proc.returncode}）。\n详情请查看：\n{log_path}"
        )
        return proc.returncode
    return 0


if __name__ == "__main__":
    sys.exit(main())
