# MSG to EML Converter

A command-line tool to convert Microsoft Office MSG files to EML format.

## Features

- Convert MSG files to EML format
- Process multiple MSG files at once
- Preserve email headers, body (text/HTML), and attachments
- Output to the same directory as input files

## Requirements

- Python 3.7 or higher

## Installation

1. Clone or download the repository

2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

To convert a single MSG file:

```bash
python msg2eml.py file.msg
```

To convert multiple MSG files at once:

```bash
python msg2eml.py file1.msg file2.msg file3.msg
```

### Output

Converted EML files are saved in the same directory as the original MSG files. The filename extension is changed to `.eml`.

Example:
- `input.msg` → `input.eml` (saved in the same directory)

### Display Help

```bash
python msg2eml.py --help
```

## Dependencies

- **extract-msg**: Library for reading and parsing MSG files

## Error Handling

The tool handles the following errors appropriately:

- File not found
- Corrupted files
- Permission errors
- Attachment processing errors

When an error occurs, an error message is displayed and processing continues. The number of successful and failed conversions is displayed at the end.

## Notes

- Existing EML files will be overwritten
- Processing may take time for large MSG files or files with many attachments
