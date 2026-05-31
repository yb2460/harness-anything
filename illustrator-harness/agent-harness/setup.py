# -*- coding: utf-8 -*-
"""cli-anything-powerpoint PyPI 包配置。

PEP 420 命名空间包: cli_anything/ (无 __init__.py)
"""

from setuptools import setup, find_namespace_packages

with open("cli_anything/powerpoint/README.md", "r", encoding="utf-8") as f:
    long_description = f.read()

setup(
    name="cli-anything-powerpoint",
    version="1.0.0",
    description="CLI harness for controlling PowerPoint — create, edit, and export presentations from the command line.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="cli-anything",
    url="https://github.com/HKUDS/CLI-Anything",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    package_data={
        "cli_anything.powerpoint": ["skills/*.md"],
    },
    python_requires=">=3.10",
    install_requires=[
        "python-pptx>=0.6.21",
        "click>=8.0",
        "Pillow>=9.0",
    ],
    extras_require={
        "repl": ["prompt_toolkit>=3.0"],
        "wps": ["pywin32>=300"],
        "all": ["prompt_toolkit>=3.0", "pywin32>=300"],
    },
    entry_points={
        "console_scripts": [
            "cli-anything-powerpoint=cli_anything.powerpoint.powerpoint_cli:cli",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Topic :: Multimedia :: Graphics :: Presentation",
        "Topic :: Utilities",
    ],
)
