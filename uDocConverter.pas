unit uDocConverter;

{*******************************************************************************
  DOC/DOCX to PDF Converter Unit
  
  Author: Alexander Buzaev
  Description: Wrapper for TotalDocConverterX COM component
  Requirements: TotalDocConverterX from https://www.coolutils.com/TotalDocConverterX
  
  Usage:
    var
      Converter: TDocConverter;
    begin
      Converter := TDocConverter.Create;
      try
        if Converter.ConvertDocToPDF('document.docx', 'output.pdf', 'log.txt') then
          ShowMessage('Success!');
      finally
        Converter.Free;
      end;
    end;
*******************************************************************************}

interface

uses
  System.SysUtils, System.Variants, Dialogs, Vcl.OleAuto;

type
  TDocConverter = class
  private
    FConverter: OleVariant;
    FLastError: string;
    FLogFile: string;
  public
    constructor Create;
    destructor Destroy; override;
    
    /// <summary>
    /// Converts DOC/DOCX/RTF file to PDF format
    /// </summary>
    /// <param name="ASourceFile">Full path to source document file</param>
    /// <param name="ADestFile">Full path to destination PDF file</param>
    /// <param name="ALogFile">Full path to log file (optional)</param>
    /// <returns>True if conversion successful, False otherwise</returns>
    function ConvertDocToPDF(const ASourceFile, ADestFile: string; 
      const ALogFile: string = ''): Boolean;
    
    /// <summary>
    /// Converts document file to PDF (alias for ConvertDocToPDF)
    /// </summary>
    function ConvertToPDF(const ASourceFile, ADestFile: string; 
      const ALogFile: string = ''): Boolean;
    
    /// <summary>
    /// Returns last error message if conversion failed
    /// </summary>
    property LastError: string read FLastError;
    
    /// <summary>
    /// Path to log file
    /// </summary>
    property LogFile: string read FLogFile write FLogFile;
  end;

implementation

{ TDocConverter }

constructor TDocConverter.Create;
begin
  inherited;
  FLastError := '';
  FLogFile := '';
  
  try
    FConverter := CreateOleObject('DocConverter.DocConverterX');
  except
    on E: Exception do
    begin
      FLastError := 'Failed to create TotalDocConverterX object. ' +
                    'Please make sure TotalDocConverterX is installed. ' +
                    'Download from: https://www.coolutils.com/TotalDocConverterX. ' +
                    'Error details: ' + E.Message;
      raise Exception.Create(FLastError);
    end;
  end;
end;

destructor TDocConverter.Destroy;
begin
  FConverter := Unassigned;
  inherited;
end;

function TDocConverter.ConvertDocToPDF(const ASourceFile, ADestFile: string;
  const ALogFile: string = ''): Boolean;
var
  ConvertParams: string;
  LogPath: string;
  FileExt: string;
begin
  Result := False;
  FLastError := '';
  
  try
    // Check if source file exists
    if not FileExists(ASourceFile) then
    begin
      FLastError := 'Source file not found: ' + ASourceFile;
      ShowMessage('Error: ' + FLastError);
      Exit;
    end;
    
    // Check file extension
    FileExt := LowerCase(ExtractFileExt(ASourceFile));
    if not ((FileExt = '.doc') or (FileExt = '.docx') or (FileExt = '.rtf') or 
            (FileExt = '.txt') or (FileExt = '.odt') or (FileExt = '.wpd')) then
    begin
      FLastError := 'Unsupported file format: ' + FileExt + '. ' +
                    'Supported formats: .doc, .docx, .rtf, .txt, .odt, .wpd';
      ShowMessage('Error: ' + FLastError);
      Exit;
    end;
    
    // Determine log file path
    if ALogFile <> '' then
      LogPath := ALogFile
    else if FLogFile <> '' then
      LogPath := FLogFile
    else
      LogPath := ChangeFileExt(ADestFile, '.log');
    
    // Prepare conversion parameters
    // -cPDF: output format is PDF (note: no space after -c)
    // -log: path to log file
    ConvertParams := Format('-cPDF -log "%s"', [LogPath]);
    
    // Perform conversion
    FConverter.Convert(ASourceFile, ADestFile, ConvertParams);
    
    // Check for errors
    if VarToStr(FConverter.ErrorMessage) <> '' then
    begin
      FLastError := VarToStr(FConverter.ErrorMessage);
      ShowMessage('Conversion error: ' + FLastError);
      Exit;
    end;
    
    // Verify output file was created
    if not FileExists(ADestFile) then
    begin
      FLastError := 'Conversion completed but output file was not created: ' + ADestFile;
      ShowMessage('Warning: ' + FLastError);
      Exit;
    end;
    
    Result := True;
    
  except
    on E: Exception do
    begin
      FLastError := 'Exception during conversion: ' + E.Message;
      ShowMessage('Error: ' + FLastError);
    end;
  end;
end;

function TDocConverter.ConvertToPDF(const ASourceFile, ADestFile: string;
  const ALogFile: string = ''): Boolean;
begin
  // Alias method for better readability
  Result := ConvertDocToPDF(ASourceFile, ADestFile, ALogFile);
end;

end.