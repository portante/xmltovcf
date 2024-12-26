# XML Processing Tools

This repository contains two Python utilities for processing XML files:

## ppxml - XML Pretty Printer

A tool that formats XML files with proper indentation and formatting options.

### Features
- Pretty prints XML with customizable indentation
- Option to include or omit XML declarations
- Proper error handling for malformed XML
- Preserves document structure and encoding

### Usage
```bash
./ppxml <xml_file> [--no-declaration] [--indent=SPACES]
```

#### Options
- `--no-declaration`: Omit the XML declaration from output
- `--indent=SPACES`: Set number of spaces for indentation (default: 2)

#### Examples
```bash
# Basic pretty printing
./ppxml contacts.xml

# Pretty print with 4-space indentation
./ppxml contacts.xml --indent=4

# Pretty print without XML declaration
./ppxml contacts.xml --no-declaration
```

## dumpxmltags - XML Tag Extractor

A utility for extracting and listing XML tags from contacts files, useful for analyzing XML structure and generating tag inventories.

### Features
- Lists all unique XML tags
- Optional attribute extraction
- Recursive tag processing for nested elements
- Sorted output for better readability
- Validates contact count against XML metadata

### Usage
```bash
./dumpxmltags <xml_file> [--with-attrs] [--recursive]
```

#### Options
- `--with-attrs`: Include attributes in the output (e.g., `tag[@attribute]`)
- `--recursive`: Process nested elements recursively

#### Examples
```bash
# List all tags
./dumpxmltags contacts.xml

# List tags with their attributes
./dumpxmltags contacts.xml --with-attrs

# List all tags including nested elements
./dumpxmltags contacts.xml --recursive

# Combine with sort and uniq for unique tag list
./dumpxmltags contacts.xml | sort | uniq
```

## Requirements
- Python 3.6 or higher
- lxml library (`pip install lxml`)

## Error Handling
Both tools include comprehensive error handling for:
- File I/O errors
- XML parsing errors
- Invalid command-line arguments

Error messages are written to stderr with descriptive information to help diagnose issues.