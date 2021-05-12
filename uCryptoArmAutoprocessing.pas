unit uCryptoArmAutoprocessing;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  System.StrUtils, System.Types,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.OleCtrls, MSScriptControl_TLB,
  Vcl.StdCtrls, ActiveX, Vcl.FileCtrl, System.Masks, DateUtils,
  Vcl.Buttons, Vcl.Samples.Spin, Vcl.ExtCtrls, frxClass, frxGradient,
  frxExportPDF;

type
  TFormMain = class(TForm)
    ScriptControlVB: TScriptControl;
    ButtonManualProcessing: TButton;
    MemoLog: TMemo;
    frxReportProtocolConfirmed: TfrxReport;
    frxReportProtocolNotConfirmed: TfrxReport;
    LabelPath: TLabel;
    EditPath: TEdit;
    ButtonPath: TButton;
    SpeedButtonPlay: TSpeedButton;
    SpeedButtonStop: TSpeedButton;
    LabelAutoProcessingInterval: TLabel;
    LabelMin: TLabel;
    LabelSec: TLabel;
    SpinEditMin: TSpinEdit;
    SpinEditSec: TSpinEdit;
    TimerAutoProcessingState: TTimer;
    LabelAutoProcessingState: TLabel;
    TimerAutoProcessing: TTimer;
    frxPDFExportProtocol: TfrxPDFExport;
    frxReportTypeProtocol: TfrxReport;
    procedure FormCreate(Sender: TObject);
    procedure ButtonManualProcessingClick(Sender: TObject);
    procedure ButtonPathClick(Sender: TObject);
    procedure SpeedButtonPlayMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpeedButtonPlayMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpeedButtonStopMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpeedButtonStopMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpinEditMinKeyPress(Sender: TObject; var Key: Char);
    procedure SpinEditSecKeyPress(Sender: TObject; var Key: Char);
    procedure SpeedButtonPlayClick(Sender: TObject);
    procedure SpeedButtonStopClick(Sender: TObject);
    procedure TimerAutoProcessingStateTimer(Sender: TObject);
    procedure TimerAutoProcessingTimer(Sender: TObject);
    procedure SpinEditSecChange(Sender: TObject);
    procedure SpinEditMinChange(Sender: TObject);

  private
    { Private declarations }
  public
    function SignatureVerify(inputFileName, inputFileNameSignature: string; out arrayResultsDescription: TStringDynArray): TSmallIntDynArray;
    function SignatureInformation(InputFileNameSignature: string): TStringDynArray;
    function CertificateInformation(InputFileNameSignature: string): TStringDynArray;

    function CheckErrorsWithinArchive(inputArchiveFileName: string): boolean;
    function CheckFileName(inputFileName: string): boolean;
    function ifFileExistsRename(inputFileName: string): string;
    function CorrectPath(inputDirectory: string): string;

    procedure CreateResponceFileToOutput(inputFileName, DescriptionError: string);

    procedure CreateProtocol(inputFileName: string;
                             inputFileNameSignature: array of string;
                             directoryFiles: string;
                             directoryExport: string;
                             inputOriginalArchiveFileName: string);

    procedure UpdateDirectories(inputDirectoryRoot: string);
    procedure SortErrorFiles;
    procedure MoveFilesToErrors(inputFileName: string);

    procedure Processed(inputArchiveFileName: string);
    procedure MoveFilesToProcessed(inputArchiveFileName, inputNotSigFile: string; inputSigFilesArray: array of string);
  end;

  TSignatureFile = class(TObject)
    Name: string;
    DateCreate: string;
    Size: string;
    SignatureInformation: TStringDynArray;
    CertificateInformation: TStringDynArray;
    VerifyStatus: TSmallIntDynArray;
    VerifyStatusDesctiption: TStringDynArray;
  end;

  TNotSignatureFile = class(TObject)
    Name: string;
    DateCreate: string;
    Size: string;
  end;

var
  FormMain: TFormMain;
var
  DirectoryRoot, DirectoryErrors, DirectoryProcessed, DirectoryOutput, DescriptionErrorArchive: string;
const
  SIGN_CORRECT = 1;

implementation

uses
  FWZipReader;

{$R *.dfm}

procedure TFormMain.FormCreate(Sender: TObject);
var ScriptFile: TextFile;
    Script, LineScript: String;
begin

  Script := '';

  if FileExists(ExtractFilePath(ParamStr(0))+'VerifyScript.vbs') then
    begin
      AssignFile(ScriptFile, ExtractFilePath(ParamStr(0))+'VerifyScript.vbs');
      Reset(ScriptFile);

      while not EOF(ScriptFile) do
        begin
          readln(ScriptFile, LineScript);
          Script := Script + LineScript + #13#10;
        end;

      CloseFile(ScriptFile);

      ScriptControlVB.Language := 'VBScript';
      ScriptControlVB.AddCode(Script);
    end
  else
    begin
      ShowMessage('���� "VerifyScript.vbs" ����������� � ����� � ����������. ��� ���� ��������� �� ����������.');
      Application.Terminate;
      Exit;
    end;

end;

procedure TFormMain.ButtonManualProcessingClick(Sender: TObject);
var SearchResult: TSearchRec;
begin
  ButtonManualProcessing.Enabled := False;
  TimerAutoProcessing.Enabled := False;
  SpeedButtonPlay.Enabled := False;
  SpeedButtonStop.Enabled := False;

  DirectoryRoot := CorrectPath(EditPath.Text);
  if System.SysUtils.DirectoryExists(DirectoryRoot) = False then
    ShowMessage('��������� ���� � ����������. ����� �� ����������.')
  else
    begin

      UpdateDirectories(DirectoryRoot);

      SortErrorFiles;

      if FindFirst(DirectoryRoot + '*.*', faNormal, SearchResult) = 0 then
        begin
          repeat
            if CheckFileName(SearchResult.Name) and CheckErrorsWithinArchive(SearchResult.Name) then
              begin
                MoveFilesToErrors(SearchResult.Name);
                MemoLog.Lines.Add( DateToStr(Now) + ' ' + TimeToStr(Now) + '  ' + DescriptionErrorArchive + #13#10);

                CreateResponceFileToOutput(SearchResult.Name, DescriptionErrorArchive);
              end
            else Processed(SearchResult.Name);
          until FindNext(SearchResult) <> 0;
          FindClose(SearchResult);
        end;

    end;

  ButtonManualProcessing.Enabled := True;
  SpeedButtonPlay.Enabled := True;
  SpeedButtonStop.Enabled := True;
end;

procedure TFormMain.Processed(inputArchiveFileName: string);
var i, arrayIndex: integer;
    Archive: TFWZipReader;
    SigFilesArray: array of string;
    NotSigFile: string;
begin
  Archive := TFWZipReader.Create;
  try
    Archive.LoadFromFile(DirectoryRoot + inputArchiveFileName);
    Archive.ExtractAll(DirectoryRoot);

    arrayIndex := 0;
    for i := 0 to Archive.Count-1 do
      begin
        if LowerCase(ExtractFileExt(Archive.item[i].FileName)) = '.sig' then
          begin
            SetLength(SigFilesArray, arrayIndex + 1);
            SigFilesArray[arrayIndex] := Archive.item[i].FileName;
            arrayIndex := arrayIndex + 1;
          end
        else
          begin
            NotSigFile := Archive.item[i].FileName;
          end;
      end;

  finally
    Archive.Free;
  end;

  MoveFilesToProcessed(inputArchiveFileName, NotSigFile, SigFilesArray);

end;

procedure TFormMain.CreateProtocol(inputFileName: string;
                                   inputFileNameSignature: array of string;
                                   directoryFiles: string;
                                   directoryExport: string;
                                   inputOriginalArchiveFileName: string);
var SignatureFiles: array of TSignatureFile;
    NotSignatureFile: TNotSignatureFile;

    NotSigFileDateTime, SigFileDateTime: TDateTime;
    NotSigFile, SigFile: File of Byte;

    i, j, counterCorrectStatus: integer;

    frxNotSigFileName, frxNotSigFileDateCreate, frxNotSigFileSize: TfrxMemoView;
    frxSigFileName, frxSigFileDateCreate, frxSigFileSize: TfrxMemoView;
    frxSigStatus, frxSigInformation, frxCertInformation: TfrxMemoView;

    protocolName: string;
begin
  NotSignatureFile := TNotSignatureFile.Create;
  NotSignatureFile.Name := directoryFiles + inputFileName;

  FileAge(NotSignatureFile.Name, NotSigFileDateTime, True);
  NotSignatureFile.DateCreate := DateTimeToStr(NotSigFileDateTime);

  AssignFile(NotSigFile, NotSignatureFile.Name);
  reset(NotSigFile);
  NotSignatureFile.Size := IntToStr(FileSize(NotSigFile)) + ' ����';
  CloseFile(NotSigFile);

  SetLength(SignatureFiles, Length(InputFileNameSignature));
  MemoLog.Lines.Add(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ������ �������� �������� �� ������: "' + inputOriginalArchiveFileName + '". ����� �������� ����� ������� ����� 30 ������. ��������...' + #13#10);
  for i := 0 to High(SignatureFiles) do
    begin
      SignatureFiles[i] := TSignatureFile.Create;
      SignatureFiles[i].Name := directoryFiles + inputFileNameSignature[i];

      FileAge(SignatureFiles[i].Name, SigFileDateTime, True);
      SignatureFiles[i].DateCreate := DateTimeToStr(SigFileDateTime);

      AssignFile(SigFile, SignatureFiles[i].Name);
      reset(SigFile);
      SignatureFiles[i].Size := IntToStr(FileSize(SigFile)) + ' ����';
      CloseFile(SigFile);

      SignatureFiles[i].CertificateInformation := CertificateInformation(SignatureFiles[i].Name);
      SignatureFiles[i].SignatureInformation := SignatureInformation(SignatureFiles[i].Name);

      SignatureFiles[i].VerifyStatus := SignatureVerify(NotSignatureFile.Name, SignatureFiles[i].Name, SignatureFiles[i].VerifyStatusDesctiption);
      For j := 0 to High(SignatureFiles[i].VerifyStatusDesctiption) do
        begin
          if SignatureFiles[i].VerifyStatus[j] = SIGN_CORRECT then
            SignatureFiles[i].VerifyStatusDesctiption[j] := '������ �������� ������� ' + '�' + IntToStr(j+1) + ': '
                                                          + SignatureFiles[i].VerifyStatusDesctiption[j] + #13#10
                                                          + '������� ������������' + #13#10 + #13#10
          else
            SignatureFiles[i].VerifyStatusDesctiption[j] := '������ �������� ������� ' + '�' + IntToStr(j+1) + ': '
                                                          + SignatureFiles[i].VerifyStatusDesctiption[j] + #13#10
                                                          + '������� �� ������������' + #13#10 + #13#10
        end;

      //���� ������ ����� *.sig �������� ���� �� ���� ������������ �������,
      //�� ������������ ������ ��������� ��� ��������������� ��������
      counterCorrectStatus := 0;
      For j := 0 to High(SignatureFiles[i].VerifyStatus) do
        begin
          if SignatureFiles[i].VerifyStatus[j] = SIGN_CORRECT then
            begin
              counterCorrectStatus := counterCorrectStatus + 1;
            end
          else
            begin
              frxReportTypeProtocol := frxReportProtocolNotConfirmed;
              protocolName := 'ProtocolNotConfirmed_';
              Break;
            end;
        end;
      if counterCorrectStatus = Length(SignatureFiles[i].VerifyStatus) then
        begin
          frxReportTypeProtocol := frxReportProtocolConfirmed;
          protocolName := 'ProtocolConfirmed_';
        end;

      frxNotSigFileName := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoNotSigFileName'));
      frxNotSigFileName.Memo.Text := ExtractFileName(NotSignatureFile.Name);
      frxSigFileName := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoSigFileName'));
      frxSigFileName.Memo.Text := ExtractFileName(SignatureFiles[i].Name);

      frxNotSigFileDateCreate := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoNotSigFileDateCreate'));
      frxNotSigFileDateCreate.Memo.Text := NotSignatureFile.DateCreate;
      frxSigFileDateCreate := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoSigFileDateCreate'));
      frxSigFileDateCreate.Memo.Text := SignatureFiles[i].DateCreate;

      frxNotSigFileSize := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoNotSigFileSize'));
      frxNotSigFileSize.Memo.Text := NotSignatureFile.Size;
      frxSigFileSize := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoSigFileSize'));
      frxSigFileSize.Memo.Text := SignatureFiles[i].Size;

      frxCertInformation := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoCertificateInformation'));
      frxCertInformation.Memo.Text := '';
      frxSigInformation := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoSignatureInformation'));
      frxSigInformation.Memo.Text := '';
      frxSigStatus := TfrxMemoView(frxReportTypeProtocol.FindObject('MemoSignatureStatus'));
      frxSigStatus.Memo.Text := '';
      For j := 0 to High(SignatureFiles[i].VerifyStatus) do
        begin
          frxCertInformation.Memo.Text := frxCertInformation.Memo.Text + SignatureFiles[i].CertificateInformation[j];
          frxSigInformation.Memo.Text := frxSigInformation.Memo.Text + SignatureFiles[i].SignatureInformation[j];
          if SignatureFiles[i].VerifyStatus[j] = SIGN_CORRECT then
            frxSigStatus.Memo.Text := frxSigStatus.Memo.Text + SignatureFiles[i].VerifyStatusDesctiption[j]
          else
            frxSigStatus.Memo.Text := frxSigStatus.Memo.Text + SignatureFiles[i].VerifyStatusDesctiption[j];
        end;
      frxCertInformation.Memo.Text := TrimRight(frxCertInformation.Memo.Text);
      frxSigInformation.Memo.Text := TrimRight(frxSigInformation.Memo.Text);
      frxSigStatus.Memo.Text := TrimRight(frxSigStatus.Memo.Text);
      MemoLog.Lines.Add(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ��������� ������� "' + ExtractFileName(SignatureFiles[i].Name) + '"' + #13#10);

      //��������� ���������� �� ���������� ����� "Output"
      //����� ��� ��� � �� ����������� ��������
      if System.SysUtils.DirectoryExists(DirectoryOutput) = False then
        System.SysUtils.ForceDirectories(DirectoryOutput);

      frxReportTypeProtocol.PrepareReport(true);
      frxPDFexportProtocol.Compressed := True;
      frxPDFexportProtocol.Background := True;
      frxPDFexportProtocol.PrintOptimized := False;
      frxPDFexportProtocol.OpenAfterExport := False;
      frxPDFexportProtocol.ShowProgress := False;
      frxPDFexportProtocol.ShowDialog := False;

      frxPDFexportProtocol.FileName := directoryExport + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
      frxReportTypeProtocol.Export(frxPDFexportProtocol);
      //�������� ���������� �� � ����� "Output" ���� � ����� �� ���������
      //���� ����������, �������� ��������
      frxPDFexportProtocol.FileName := directoryOutput + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
      frxPDFexportProtocol.FileName := ifFileExistsRename(frxPDFexportProtocol.FileName);
      frxReportTypeProtocol.Export(frxPDFexportProtocol);
    end;

end;

function TFormMain.SignatureVerify (inputFileName, inputFileNameSignature: string; out arrayResultsDescription: TStringDynArray): TSmallIntDynArray;
var VArr, resultFromVBS: Variant;
    functionParameters: PSafeArray;
    arrayResults: TSmallIntDynArray;
    arrayResultsD: TStringDynArray;
    i: integer;
begin
  try
    VArr:=VarArrayCreate([0, 1], varVariant);
    VArr[0] := inputFileName;
    VArr[1] := inputFileNameSignature;

    functionParameters := PSafeArray(TVarData(VArr).VArray);

    resultFromVBS := ScriptControlVB.Run('SignatureVerify', functionParameters);
    arrayResults := ResultFromVBS;
    SetLength(arrayResultsD, Length(arrayResults));
    For i := 0 to High(arrayResults) do
      begin
        case arrayResults[i] of
          1 : arrayResultsD[i] := '�����';
          3 : arrayResultsD[i] := '������� ����������� ��� � ��� ��� �������';
        else arrayResultsD[i] := '������ �� ��������';
        end;
      end;

  except
    on E: Exception do
    MessageDlg(PWideChar(E.Message), mtError, [mbOk], 0);
  end;

  result := arrayResults;
  arrayResultsDescription := arrayResultsD;
end;

function TFormMain.SignatureInformation(InputFileNameSignature: string): TStringDynArray;
var VArr, resultFromVBS: Variant;
    functionParameters: PSafeArray;
    arrayResults: TStringDynArray;
begin
  try
    VArr:=VarArrayCreate([0, 0], varVariant);
    VArr[0] := inputFileNameSignature;

    functionParameters := PSafeArray(TVarData(VArr).VArray);

    resultFromVBS := ScriptControlVB.Run('SignatureInformation', FunctionParameters);
    arrayResults := resultFromVBS;

  except
    on E: Exception do
    MessageDlg(PWideChar(E.Message), mtError, [mbOk], 0);
  end;

  result := arrayResults;
end;

function TFormMain.CertificateInformation(InputFileNameSignature: string): TStringDynArray;
var VArr, resultFromVBS: Variant;
    functionParameters: PSafeArray;
    arrayResults: TStringDynArray;
begin
  try
    VArr:=VarArrayCreate([0, 0], varVariant);
    VArr[0] := inputFileNameSignature;

    functionParameters := PSafeArray(TVarData(VArr).VArray);

    resultFromVBS := ScriptControlVB.Run('CertificateInformation', FunctionParameters);
    arrayResults := resultFromVBS;

  except
    on E: Exception do
    MessageDlg(PWideChar(E.Message), mtError, [mbOk], 0);
  end;

  result := arrayResults;
end;

procedure TFormMain.SortErrorFiles;
var SearchResult: TSearchRec;
begin
  if System.SysUtils.DirectoryExists(DirectoryErrors) = False then
    System.SysUtils.ForceDirectories(DirectoryErrors);

  if FindFirst(DirectoryRoot + '*.*', faNormal, SearchResult) = 0 then
    begin
      repeat
        if (LowerCase(ExtractFileExt(SearchResult.Name)) <> '.zip') or
           (CheckFileName(SearchResult.Name) = false) then
          begin
            MoveFilesToErrors(SearchResult.Name);
            MemoLog.Lines.Add(DateToStr(Now) + ' ' + TimeToStr(Now) + '  �������� ��� ����� "' + SearchResult.Name + '"' + #13#10);
          end;
      until FindNext(SearchResult) <> 0;
      FindClose(SearchResult);
    end;
end;

procedure TFormMain.CreateResponceFileToOutput(inputFileName: string; DescriptionError: string);
var responceTextFile: TextFile;
    responceTextFileName: string;
begin
  if System.SysUtils.DirectoryExists(DirectoryOutput) = False then
    System.SysUtils.ForceDirectories(DirectoryOutput);
  responceTextFileName := DirectoryOutput + 'response_' + StringReplace(inputFileName, ExtractFileExt(inputFileName), '', [rfIgnoreCase]) + '.txt';
  responceTextFileName := ifFileExistsRename(responceTextFileName);
  AssignFile(responceTextFile, responceTextFileName);
  ReWrite(responceTextFile);
  WriteLn(responceTextFile, DescriptionError);
  CloseFile(responceTextFile);
end;

function TFormMain.ifFileExistsRename(inputFileName: string): string;
var counterName: integer;
begin
  result := inputFileName;

  counterName := 0;
  while FileExists(inputFileName) do
    begin
      counterName := counterName + 1;
      if counterName = 1 then
        begin
          Insert(' (' + IntToStr(counterName) + ')', inputFileName, Length(inputFileName)-3);
          result := inputFileName;
        end
      else
        begin
          inputFileName := StringReplace(inputFileName, ' (' + IntToStr(counterName-1) + ')', ' (' + IntToStr(counterName) + ')', []);
          result := inputFileName;
        end;
    end;
end;

function TFormMain.CheckFileName(inputFileName: string): boolean;
begin
  Result := False;

  if MatchesMask(inputFileName, 'SH_*_*_*.zip') or
     MatchesMask(inputFileName, 'SHO_*_*_*.zip') or
     MatchesMask(inputFileName, 'SMP_*_*_*.zip') or
     MatchesMask(inputFileName, 'SHCP_*_*_��������.zip') or
     MatchesMask(inputFileName, 'MSHO_*_MTR_*.zip') or
     MatchesMask(inputFileName, 'MSH_*_MTR_*.zip') or
     MatchesMask(inputFileName, 'MSMP_*_MTR_*.zip') then
    begin
      Result := True;
    end;

end;

function TFormMain.CheckErrorsWithinArchive(inputArchiveFileName: string): boolean;
var i, Counter: integer;
    Archive: TFWZipReader;
begin

  Result := False;

  Archive := TFWZipReader.Create;
  try
    Archive.LoadFromFile(DirectoryRoot + inputArchiveFileName);

    Counter := 0;
    //�������� �� ���������� ������, ����������� � zip-������ ��� ����������. �� ���������� � ������ ������ ���� 1 ���� ��� ����������.
    for i := 0 to Archive.Count-1 do
      begin
        if LowerCase(ExtractFileExt(Archive.item[i].FileName)) <> '.sig' then
          Counter := Counter + 1;
      end;
    if Counter > 1 then
      begin
        Result := True;
        DescriptionErrorArchive := '� zip-������ "' + inputArchiveFileName + '" ����� ������ ����� ��� ����������.';
      end;
    if Counter = 0 then
      begin
        Result := True;
        DescriptionErrorArchive := '� zip-������ "' + inputArchiveFileName + '" ����������� ����� ��� ����������.';
      end;

    //�������� �� ���������� ��������. ���� � zip-������ ������� �����������, �� � �����.
    Counter := 0;
    for i := 0 to Archive.Count-1 do
      begin
        if LowerCase(ExtractFileExt(Archive.item[i].FileName)) = '.sig' then
          counter := Counter + 1;
      end;
    if Counter = 0 then
      begin
        Result := True;
        DescriptionErrorArchive := '� zip-������ "' + inputArchiveFileName + '" ����������� �����-������� � ����������� ".sig"';
      end;

    //�������� �� ������������ ��� ������ ������ zip-������
    for i := 0 to Archive.Count-1 do
      begin
        if Not MatchesMask( Archive.item[i].FileName, Copy(inputArchiveFileName, 1, AnsiPos('_', inputArchiveFileName) + 12) + '*' ) then
          begin
            Result := True;
            DescriptionErrorArchive := '����� ������ zip-������ "' + inputArchiveFileName + '" �� ������������� ��� ��������';
          end;
      end;

  finally
    Archive.Free;
  end;

end;

procedure TFormMain.MoveFilesToProcessed(inputArchiveFileName: string; inputNotSigFile: string; inputSigFilesArray: array of string);
var DirectoryFrom, DirectoryTo, fileDirectoryFrom, fileDirectoryTo: string;
    MO: string;
    pointerFileDirectoryFrom, pointerFileDirectoryTo: PWideChar;
    Year, Month: integer;
    i, indexArray: integer;
begin
  Year := YearOf(Date);
  Month := MonthOf(Date);
  MO := Copy(inputArchiveFileName, AnsiPos('_', inputArchiveFileName) + 1, 6);

  DirectoryFrom := DirectoryRoot;

  DirectoryTo := DirectoryProcessed + IntToStr(Year) + '\' + IntToStr(Month) + '\' + MO + '\' +
                 StringReplace(inputArchiveFileName, ExtractFileExt(inputArchiveFileName), '', [rfIgnoreCase]) + '\';
  //��������� ���������� �� ���������� � ����� �� ���������
  if System.SysUtils.DirectoryExists(DirectoryTo) then
    begin
      i := 0;
      while System.SysUtils.DirectoryExists(DirectoryTo) do
        begin
          i := i+1;
          if i = 1 then
            begin
              Insert(' (' + IntToStr(i) + ')', DirectoryTo, Length(DirectoryTo));
            end
          else
            begin
              DirectoryTo := StringReplace(DirectoryTo, ' (' + IntToStr(i-1) +')', ' (' + IntToStr(i) + ')', []);
            end;
        end;
      System.SysUtils.ForceDirectories(DirectoryTo);
    end
  else
    System.SysUtils.ForceDirectories(DirectoryTo);

  //������ ��������
  CreateProtocol(inputNotSigFile, inputSigFilesArray, DirectoryFrom, DirectoryTo, inputArchiveFileName);

  //��������� �����
  fileDirectoryFrom := DirectoryFrom + inputArchiveFileName;
  pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
  fileDirectoryTo := DirectoryTo + inputArchiveFileName;
  pointerFileDirectoryTo := Addr(fileDirectoryTo[1]);
  MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryTo);

  fileDirectoryFrom := DirectoryFrom + inputNotSigFile;
  pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
  fileDirectoryTo := DirectoryTo + inputNotSigFile;
  pointerFileDirectoryTo := Addr(fileDirectoryTo[1]);
  MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryTo);

  For indexArray := 0 to High(inputSigFilesArray) do
    begin
      fileDirectoryFrom := DirectoryFrom + inputSigFilesArray[indexArray];
      pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
      fileDirectoryTo := DirectoryTo + inputSigFilesArray[indexArray];
      pointerFileDirectoryTo := Addr(fileDirectoryTo[1]);
      MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryTo);
    end;

end;

procedure TFormMain.MoveFilesToErrors(inputFileName: string);
var fileDirectoryFrom, fileDirectoryTo: string;
    pointerFileDirectoryFrom, pointerFileDirectoryTo: PWideChar;
begin
  if System.SysUtils.DirectoryExists(DirectoryErrors) = False then
    System.SysUtils.ForceDirectories(DirectoryErrors);

  fileDirectoryFrom := DirectoryRoot + inputFileName;
  pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);

  fileDirectoryTo := DirectoryErrors + inputFileName;
  fileDirectoryTo := ifFileExistsRename(fileDirectoryTo);
  pointerFileDirectoryTo := Addr(fileDirectoryTo[1]);

  MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryTo);
end;

function TFormMain.CorrectPath(inputDirectory: string): string;
begin
  if Pos('/', inputDirectory) <> 0 then
    begin
      inputDirectory := StringReplace(inputDirectory, '/', '\', [rfReplaceAll]);
    end;

  if inputDirectory[length(inputDirectory)] <> '\' then
    Result := inputDirectory + '\'
  else
    Result := inputDirectory;
end;

procedure TFormMain.UpdateDirectories(inputDirectoryRoot: string);
begin
  DirectoryErrors := DirectoryRoot + 'Errors';
  DirectoryErrors := CorrectPath(DirectoryErrors);

  DirectoryProcessed := DirectoryRoot + 'Processed';
  DirectoryProcessed := CorrectPath(DirectoryProcessed);

  DirectoryOutput := DirectoryRoot + 'Output';
  DirectoryOutput := CorrectPath(DirectoryOutput);
end;

procedure TFormMain.ButtonPathClick(Sender: TObject);
begin
  if SelectDirectory('�������� ����� ��� ������ ���������������:', '', DirectoryRoot, [sdNewFolder, sdShowShares, sdValidateDir]) then
    EditPath.Text := DirectoryRoot;
end;

procedure TFormMain.SpeedButtonPlayClick(Sender: TObject);
begin
  SpeedButtonPlay.Visible := False;
  SpeedButtonStop.Visible := True;

  if (SpinEditSec.Value > 59) or
     (SpinEditSec.Value < 0) or
     (SpinEditMin.Value < 0) or
     ( (SpinEditSec.Value = 0) and (SpinEditMin.Value = 0) ) then
    ShowMessage('������� ������� ������/�������')
  else
    begin
      TimerAutoProcessing.Interval := SpinEditMin.Value * 60000 + SpinEditSec.Value * 1000;
      TimerAutoProcessing.Enabled := True;

      TimerAutoProcessingState.Enabled := True;
      LabelAutoProcessingState.Caption := '�������������� �������';
    end;
end;

procedure TFormMain.SpeedButtonStopClick(Sender: TObject);
begin
  SpeedButtonStop.Visible := False;
  SpeedButtonPlay.Visible := True;

  TimerAutoProcessingState.Enabled := False;
  LabelAutoProcessingState.Caption := '�������������� �� �������';

  TimerAutoProcessing.Enabled := False;
end;

procedure TFormMain.TimerAutoProcessingStateTimer(Sender: TObject);
begin
  LabelAutoProcessingState.Caption := LabelAutoProcessingState.Caption + '.';
  if LabelAutoProcessingState.Caption = '�������������� �������....' then
    LabelAutoProcessingState.Caption := '�������������� �������';
end;

procedure TFormMain.TimerAutoProcessingTimer(Sender: TObject);
begin
  TimerAutoProcessing.Enabled := False;
  ButtonManualProcessingClick(Self);
  TimerAutoProcessing.Enabled := True;
end;

procedure TFormMain.SpeedButtonPlayMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  SpeedButtonPlay.Glyph.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'Icons\PlayPush.bmp');
end;

procedure TFormMain.SpeedButtonPlayMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  SpeedButtonPlay.Glyph.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'Icons\Play.bmp');
end;

procedure TFormMain.SpeedButtonStopMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  SpeedButtonStop.Glyph.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'Icons\StopPush.bmp');
end;

procedure TFormMain.SpeedButtonStopMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  SpeedButtonStop.Glyph.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'Icons\Stop.bmp');
end;

procedure TFormMain.SpinEditMinKeyPress(Sender: TObject; var Key: Char);
begin
  SpinEditMin.SelLength := 1;
end;

procedure TFormMain.SpinEditSecKeyPress(Sender: TObject; var Key: Char);
begin
  SpinEditSec.SelLength := 1;
end;

procedure TFormMain.SpinEditSecChange(Sender: TObject);
begin
  if SpinEditSec.Text = '' then
    SpinEditSec.Text := '0';
end;

procedure TFormMain.SpinEditMinChange(Sender: TObject);
begin
  if SpinEditMin.Text = '' then
    SpinEditMin.Text := '0';
end;

end.
