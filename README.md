# DOC to PDF Converter

Delphi project for converting DOC/DOCX (Microsoft Word) and other document formats to PDF using TotalDocConverterX COM component.

## Requirements

- Delphi 
- **TotalDocConverterX** installed on the system
  - Download from: https://www.coolutils.com/TotalDocConverterX
- Windows OS

## Description

This project provides a simple interface for converting Microsoft Word documents (DOC, DOCX) and other text document formats to PDF using the TotalDocConverterX COM object from CoolUtils.

## Supported Input Formats

- **DOC** - Microsoft Word 97-2003 documents
- **DOCX** - Microsoft Word 2007+ documents
- **RTF** - Rich Text Format
- **TXT** - Plain text files
- **ODT** - OpenDocument Text
- **WPD** - WordPerfect documents

## Features

- DOC/DOCX to PDF conversion
- Support for multiple document formats
- Error handling and logging
- Preserve document formatting in PDF output
- Configurable conversion parameters
- Simple and clean code structure

## Installation

### Step 1: Install TotalDocConverterX

1. Download TotalDocConverterX from https://www.coolutils.com/TotalDocConverterX
2. Install and register the component on your system
3. Make sure the COM object is properly registered

### Step 2: Clone and Build Project

```bash
git clone https://github.com/yourusername/doc-to-pdf-converter.git
cd doc-to-pdf-converter
```

Open the project in Delphi, build and run.

## Usage

### Basic Example

```pascal
uses 
  Dialogs, 
  Vcl.OleAuto;

var
  Converter: OleVariant;
begin
  Converter := CreateOleObject('DocConverter.DocConverterX');
  Converter.Convert('c:\test\source.docx', 'c:\test\dest.pdf', '-cPDF -log c:\test\Doc.log');
  if Converter.ErrorMessage <> '' then
    ShowMessage(Converter.ErrorMessage);
end;
```

### Using TDocConverter Class

```pascal
var
  Converter: TDocConverter;
begin
  Converter := TDocConverter.Create;
  try
    if Converter.ConvertDocToPDF('document.docx', 'output.pdf', 'conversion.log') then
      ShowMessage('Conversion completed!')
    else
      ShowMessage('Error: ' + Converter.LastError);
  finally
    Converter.Free;
  end;
end;
```

## Configuration

You can modify the following parameters in the Convert method:
- **Source file path** - path to input DOC/DOCX/RTF file
- **Destination PDF file path** - path to output PDF file
- **Conversion options:**
  - `-cPDF` - output format (note: no space after -c)
  - `-log <path>` - log file path
  - Additional parameters supported by TotalDocConverterX

## Error Handling

The application checks for error messages returned by the converter and displays them using a message dialog. All errors are also logged to the specified log file.

## Project Structure

```
doc-to-pdf-converter/
├── README.md
├── LICENSE
├── .gitignore
├── uDocConverter.pas         # Main converter unit
└── DocToPDFConverter.dpr     # Example program
```

## License

MIT License - see LICENSE file for details

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## Author

**Alexander Buzaev**

## Acknowledgments

- CoolUtils for TotalDocConverterX component
- https://www.coolutils.com/TotalDocConverterX

## Support

For issues with TotalDocConverterX component, please refer to:
- Official documentation: https://www.coolutils.com/TotalDocConverterX
- CoolUtils support

For issues with this Delphi wrapper, please open an issue on GitHub.
