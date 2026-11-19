# -*- coding: utf-8 -*-
"""资源路径解析：兼容源码运行与 PyInstaller --onefile 解包后的临时目录。"""
import pathlib
import sys


def is_frozen():
    return getattr(sys, "frozen", False)


def app_dir():
    """exe 所在目录（源码运行时为工程根目录）。"""
    if is_frozen():
        return pathlib.Path(sys.executable).resolve().parent
    return pathlib.Path(__file__).resolve().parent.parent


def resource_dir():
    """内置数据目录。"""
    if is_frozen():
        return pathlib.Path(getattr(sys, "_MEIPASS", app_dir())) / "data"
    return pathlib.Path(__file__).resolve().parent / "data"


def resource(name):
    return resource_dir() / name


def external_path(name):
    """同目录下的外置数据（若存在则覆盖内置），用于后续补库。"""
    return app_dir() / name
