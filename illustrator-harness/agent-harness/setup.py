# -*- coding: utf-8 -*-
"""Packaging configuration for the CLI-Anything Illustrator harness."""

from pathlib import Path

from setuptools import find_namespace_packages, setup


ROOT = Path(__file__).resolve().parent

setup(
    name="cli-anything-illustrator",
    version="1.0.0",
    description="CLI harness for controlling Adobe Illustrator through COM automation",
    long_description=(ROOT / "README.md").read_text(encoding="utf-8"),
    long_description_content_type="text/markdown",
    author="cli-anything contributors",
    url="https://github.com/yb2460/harness-anything",
    license="MIT",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    python_requires=">=3.10",
    install_requires=[
        "click>=8.0",
        "pywin32>=305; platform_system == 'Windows'",
    ],
    extras_require={
        "dev": ["pytest>=7"],
    },
    entry_points={
        "console_scripts": [
            "cli-anything-illustrator=cli_anything.illustrator.illustrator_cli:cli",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Topic :: Multimedia :: Graphics :: Editors :: Vector-Based",
    ],
)
