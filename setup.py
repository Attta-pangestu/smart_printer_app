"""Setup script for Printer Sharing System."""

from setuptools import setup, find_packages
from pathlib import Path

# Read README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding='utf-8')

# Read requirements
requirements = []
requirements_file = this_directory / "requirements.txt"
if requirements_file.exists():
    with open(requirements_file, 'r', encoding='utf-8') as f:
        requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]

setup(
    name="printer-sharing-system",
    version="1.0.0",
    author="Rebinmas",
    author_email="rebinmas@example.com",
    description="A client-server system for sharing printers across LAN networks",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/rebinmas/printer-sharing-system",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Intended Audience :: System Administrators",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: System :: Networking",
        "Topic :: System :: Systems Administration",
        "Topic :: Printing",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-asyncio>=0.21.0",
            "black>=22.0.0",
            "flake8>=4.0.0",
            "mypy>=0.950",
        ],
        "gui": [
            "tkinter",  # Usually included with Python
        ],
    },
    entry_points={
        "console_scripts": [
            "printer-server=server.main:main",
            "printer-client=client.cli:main",
            "printer-gui=client.gui:main",
        ],
    },
    include_package_data=True,
    package_data={
        "web": ["*.html", "*.css", "*.js"],
        "server": ["*.yaml", "*.yml"],
    },
    zip_safe=False,
    keywords="printer sharing network lan client server windows",
    project_urls={
        "Bug Reports": "https://github.com/rebinmas/printer-sharing-system/issues",
        "Source": "https://github.com/rebinmas/printer-sharing-system",
        "Documentation": "https://github.com/rebinmas/printer-sharing-system/wiki",
    },
)