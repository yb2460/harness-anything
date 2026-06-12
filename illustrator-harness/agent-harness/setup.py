# -*- coding: utf-8 -*-
"""cli-anything-illustrator PyPI 包配置 — Adobe Illustrator COM 自动化。

PEP 420 命名空间包: cli_anything/ (无 __init__.py)
"""

from setuptools import setup, find_namespace_packages

setup(
    name="cli-anything-illustrator",
    version="1.0.0",
    description="CLI harness for Adobe Illustrator — create and edit vector graphics via COM automation.",
    author="cli-anything",
    url="https://github.com/yb2460/harness-anything",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    package_data={
        "cli_anything.illustrator": ["skills/*.md"],
    },
    python_requires=">=3.10",
    install_requires=[
        "click>=8.0",
        "pywin32>=305",
    ],
    extras_require={
        "repl": ["prompt_toolkit>=3.0"],
        "all": ["prompt_toolkit>=3.0"],
    },
    entry_points={
        "console_scripts": [
            "cli-anything-illustrator=cli_anything.illustrator.illustrator_cli:cli",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Topic :: Multimedia :: Graphics",
        "Topic :: Utilities",
    ],
)
