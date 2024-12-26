# Contacts XML Processing Utilities

A collection of Python tools for processing and converting Microsoft Office XML
contact files.

## Tools

### xmltovcf
Convert Microsoft Office XML contact files to vCard (.vcf) format.

This is NOT intended to be a general purpose tool to convert any XML file to
vCard. It is more of an _example_ of how you perform such a conversion.

However, there are nuances in how each of us uses/used Microsoft Office
Contacts, so it's probably not going to work out of the box for you. As a
result, this tool is a little less complex than a "one size fits all"
solution.

The list XML element tags considered was derived from the output of the
`contacttagfreq` tool. From there, we manually added tags as we reviewed the
source XML file.

The `ppxml` tool is provided out of convenience since most of the generated
XML from Microsoft Office is pretty ugly.

The email address handling is a little odd, as there are four lists of email
addresses. It's not clear how the lists are related, so we just created a
dictionary of all the unique email addresses and then processed them when we
have finished with all the contact elements. It _seems_ like the email address
"type" of "1" _seems_ to mean "home", "0" _seems_ to mean "work", and "2"
_seems_ to mean "other". That may not be the same for another data set.

This code also handles the weird way we had used "title" to store our
Christmas Card address label names. Those values are converted using the
`X-RELATEDNAMES` and `X-ABLabel` vCard fields.

#### Features
- Converts complex Microsoft Office XML contact exports to standard vCard
  format
- Preserves contact details including:
  - Name (first, last, middle, suffix)
  - Phone numbers (home, work, cell)
  - Email addresses
  - Physical addresses
  - Web pages
  - Notes and related information

#### Usage
```bash
./xmltovcf contacts.xml > contacts.vcf
```

### ppxml
Pretty print XML files with customizable formatting.

#### Features
- Formats XML with proper indentation
- Optional XML declaration control
- Preserves document structure

#### Usage
```bash
./ppxml contacts.xml
```

### contacttagfreq
Analyze and report tag frequencies in XML contact files.

#### Features
- Counts occurrences of different XML tags
- Helps understand the structure of contact XML exports

#### Usage
```bash
./contacttagfreq contacts.xml
```

## Requirements
- Python 3.6+
- lxml library
- vobject library

## Installation
```bash
pip install lxml vobject
```

## License
GPLv3