# -*- coding: utf-8 -*-
"""cli-anything-illustrator - local fix.

The upstream setup.py in this repo is a copy of a PowerPoint harness manifest:
it reads cli_anything/powerpoint/README.md (absent) and points its console
script at cli_anything.powerpoint.powerpoint_cli (absent), so `pip install -e .`
aborts before installing anything. This file matches the package layout the
harness code itself imports (cli_anything.illustrator).
"""
from setuptools import setup, find_namespace_packages

setup(
    name="cli-anything-illustrator",
    version="1.0.0",
    description="CLI harness for Adobe Illustrator via COM automation",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    python_requires=">=3.10",
    install_requires=["click>=8.0", "pywin32>=305"],
    entry_points={
        "console_scripts": [
            "cli-anything-illustrator=cli_anything.illustrator.illustrator_cli:cli",
        ],
    },
)