from setuptools import setup, find_packages
from pathlib import Path

setup(
    name="report-generator",
    version="1.0.1",  # bump since packaging behavior changed
    description="GUI-based patient report generator using Word templates and database integration",
    author="Nghia Nguyen",
    packages=find_packages(),
    include_package_data=True,  # respect MANIFEST.in for sdists
    package_data={
        # include templates and data inside the package for wheels
        "report_generator_v1": [
            "templates/*.docx",
            "data/*.xlsx",
            "data/*.xls", 
        ]
    },
    install_requires=Path("requirements.txt").read_text(encoding="utf-8").splitlines(),
    entry_points={
        "console_scripts": [
            "report-generator = report_generator_v1.main:main",
        ],
    },
    python_requires=">=3.9",
)
