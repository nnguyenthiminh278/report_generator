from setuptools import setup, find_packages

setup(
    name="report-generator",
    version="1.0.0",
    description="GUI-based patient report generator using Word templates and database integration",
    author="Nghia Nguyen",
    packages=find_packages(),
    include_package_data=True,
    install_requires=open("requirements.txt").read().splitlines(),
    entry_points={
        "console_scripts": [
            "report-generator = report_generator_v1.main:main",
        ],
    },
    python_requires=">=3.9",
)