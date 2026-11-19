# -*- coding: utf-8 -*-
"""
剪贴板变化检测（Windows）。

用 GetClipboardSequenceNumber 只取一个递增序号，不打开剪贴板、不占用剪贴板锁，
比反复 clipboard_get 轻得多，也不会干扰考试页面自己的复制操作。
文本读取仍在主线程用 tkinter.clipboard_get 完成，避免跨线程访问 Tk。

非 Windows 平台退化为「总是返回有变化」，由调用方按文本内容去重。
"""
import ctypes
import sys

_MIN_LEN = 6       # 太短的选中文本不当题目
_MAX_LEN = 3000    # 超长文本（整页复制）忽略，避免卡顿


class ClipboardWatcher:
    def __init__(self):
        self._user32 = None
        if sys.platform == "win32":
            try:
                self._user32 = ctypes.windll.user32
                self._user32.GetClipboardSequenceNumber.restype = ctypes.c_uint32
            except Exception:
                self._user32 = None
        self._last_seq = None
        self._last_text = ""
        self.enabled = False
        self.reset()

    def reset(self):
        self._last_seq = self._seq()
        self._last_text = ""

    def _seq(self):
        if self._user32 is None:
            return None
        try:
            return int(self._user32.GetClipboardSequenceNumber())
        except Exception:
            return None

    def enable(self):
        self.enabled = True
        self.reset()

    def disable(self):
        self.enabled = False

    def changed(self):
        """本次轮询剪贴板序号是否变化。"""
        if not self.enabled:
            return False
        seq = self._seq()
        if seq is None:
            return True  # 非 Windows：交给文本去重
        if seq == self._last_seq:
            return False
        self._last_seq = seq
        return True

    def accept(self, text):
        """
        判断这段文本是否作为一次查询。
        返回规范化后的文本，或 None（忽略）。
        """
        if not self.enabled or not text:
            return None
        s = str(text).strip()
        if not (_MIN_LEN <= len(s) <= _MAX_LEN):
            return None
        if s == self._last_text:
            return None
        self._last_text = s
        return s
