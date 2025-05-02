# PDF Reader

[![Build Windows EXE](https://github.com/mspitzbarth/pdf-reader/actions/workflows/windows-build.yml/badge.svg)](https://github.com/mspitzbarth/pdf-reader/actions/workflows/windows-build.yml)

A simple PDF reader that extracts and processes tables from PDF files.

## Features

- Extract tables from PDF files.
- Process extracted data using `pandas` and output in an Excel file format with custom formatting.
- Built as a Windows executable using PyInstaller for easy distribution.

## Installation

You can clone the repository and install the required dependencies:

```bash
git clone https://github.com/mspitzbarth/pdf-reader.git
cd pdf-reader
pip install -r requirements.txt
```

## Usage

To run the script as a Python application:

```bash
python main.py
```

Or download the latest Windows EXE from the [releases page](https://github.com/mspitzbarth/pdf-reader/releases).

## Build Windows EXE

This project uses [GitHub Actions](https://github.com/mspitzbarth/pdf-reader/actions) to automatically build the EXE for Windows. The build badge above shows the status of the latest build.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
