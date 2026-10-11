"""WPS Office CLI harness - setup."""
from setuptools import setup, find_namespace_packages

setup(
    name="cli-anything-wps",
    version="1.0.0",
    description="CLI harness for WPS Office — command-line control via COM automation",
    author="CLI-Anything",
    python_requires=">=3.10",
    packages=find_namespace_packages(include=["cli_anything.*"]),
    install_requires=[
        "click>=8.0.0",
        "prompt-toolkit>=3.0.0",
        "pywin32>=305; platform_system == 'Windows'",
    ],
    extras_require={
        "dev": ["pytest", "pytest-cov"],
        "web": ["mcp>=1.26,<2", "httpx>=0.27,<1"],
    },
    entry_points={
        "console_scripts": [
            "cli-anything-wps=cli_anything.wps.wps_cli:main",
        ],
    },
    package_data={
        "cli_anything.wps": ["skills/*.md"],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Environment :: Console",
        "Intended Audience :: Developers",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
)
