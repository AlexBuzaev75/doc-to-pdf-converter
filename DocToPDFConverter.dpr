program DocToPDFConverter;

{*******************************************************************************
  DOC to PDF Converter
  
  Author: Alexander Buzaev
  Description: Example program demonstrating DOC/DOCX to PDF conversion
               using TotalDocConverterX
  Requirements: 
    - Delphi 10.4+
    - TotalDocConverterX from https://www.coolutils.com/TotalDocConverterX
*******************************************************************************}

uses
  Vcl.Forms,
  Dialogs,
  SysUtils,
  uDocConverter in 'uDocConverter.pas';

{$R *.res}

procedure ConvertDocExample;
var
  Converter: TDocConverter;
  SourceFile, DestFile, LogFile: string;
begin
  // Set file paths
  SourceFile := 'c:\test\source.docx';
  DestFile := 'c:\test\dest.pdf';
  LogFile := 'c:\test\conversion.log';
  
  // Create converter instance
  Converter := TDocConverter.Create;
  try
    // Perform conversion
    if Converter.ConvertDocToPDF(SourceFile, DestFile, LogFile) then
    begin
      ShowMessage('Conversion completed successfully!' + sLineBreak +
                  'PDF file: ' + DestFile);
    end
    else
    begin
      ShowMessage('Conversion failed!' + sLineBreak +
                  'Error: ' + Converter.LastError + sLineBreak +
                  'Check log file: ' + LogFile);
    end;
  finally
    Converter.Free;
  end;
end;

procedure ConvertRTFExample;
var
  Converter: TDocConverter;
  SourceFile, DestFile: string;
begin
  // Example: Converting RTF file
  SourceFile := 'c:\test\document.rtf';
  DestFile := 'c:\test\document.pdf';
  
  Converter := TDocConverter.Create;
  try
    if Converter.ConvertToPDF(SourceFile, DestFile) then
      ShowMessage('RTF converted to PDF successfully!');
  finally
    Converter.Free;
  end;
end;

procedure SimpleConvertExample;
var
  c: OleVariant;
begin
  // Simple conversion example (original code style)
  try
    c := CreateOleObject('DocConverter.DocConverterX');
    c.Convert('c:\test\source.docx', 'c:\test\dest.pdf', 
              '-cPDF -log c:\test\Doc.log');
    
    if c.ErrorMessage <> '' then
      ShowMessage('Error: ' + c.ErrorMessage)
    else
      ShowMessage('Conversion completed successfully!');
  except
    on E: Exception do
      ShowMessage('Error: ' + E.Message);
  end;
end;

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  
  try
    // Use one of the examples:
    
    // Method 1: Converting DOCX file (recommended)
    ConvertDocExample;
    
    // Method 2: Converting RTF file
    // ConvertRTFExample;
    
    // Method 3: Simple direct conversion
    // SimpleConvertExample;
    
  except
    on E: Exception do
      ShowMessage('Application error: ' + E.Message);
  end;
end.