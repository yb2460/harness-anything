# -*- coding: utf-8 -*-
"""REPL skin for Illustrator CLI (powered by cli-anything repl_skin)."""
import os, sys
from pathlib import Path

_RESET = "\033[0m"; _BOLD = "\033[1m"
_CYAN = "\033[38;5;80m"; _WHITE = "\033[97m"
_GRAY = "\033[38;5;245m"; _DARK_GRAY = "\033[38;5;240m"
_LIGHT_GRAY = "\033[38;5;250m"
_GREEN = "\033[38;5;78m"; _YELLOW = "\033[38;5;220m"
_RED = "\033[38;5;196m"; _BLUE = "\033[38;5;75m"
_ACCENT = "\033[38;5;214m"  # Orange for Illustrator

class ReplSkin:
    def __init__(self, software="illustrator", version="1.0.0"):
        self.software = software; self.version = version
        self._color = sys.stdout.isatty()

    def _c(self, code, text):
        return f"{code}{text}{_RESET}" if self._color else text

    def print_banner(self):
        print(f"  {self._c(_CYAN+_BOLD, '◆')}  {self._c(_ACCENT, 'cli-anything-illustrator')} {self._c(_GRAY, 'v'+self.version)}")
        print()

    def success(self, msg): print(f"  {self._c(_GREEN, '✓')} {self._c(_GREEN, msg)}")
    def error(self, msg): print(f"  {self._c(_RED, '✗')} {self._c(_RED, msg)}")
    def info(self, msg): print(f"  {self._c(_BLUE, '●')} {self._c(_LIGHT_GRAY, msg)}")
    def prompt(self, doc_name=""):
        ctx = f"[{doc_name}] " if doc_name else ""
        return f"{self._c(_CYAN, '◆')} {self._c(_ACCENT, 'ai')} {self._c(_DARK_GRAY, ctx)}{self._c(_GRAY, '❯ ')}"

    def help(self, cmds: dict):
        print(f"\n  {self._c(_ACCENT+_BOLD, 'Commands:')}")
        print()
        max_c = max(len(c) for c in cmds)
        for cmd, desc in cmds.items():
            print(f"  {self._c(_ACCENT, cmd):<{max_c+10}}  {self._c(_GRAY, desc)}")
        print()
