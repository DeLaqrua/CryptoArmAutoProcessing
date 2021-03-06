unit uCryptoArmAutoprocessing;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  System.StrUtils, System.Types,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.OleCtrls, MSScriptControl_TLB,
  Vcl.StdCtrls, ActiveX, Vcl.FileCtrl, System.Masks, DateUtils,
  Vcl.Buttons, Vcl.Samples.Spin, Vcl.ExtCtrls, frxClass, frxGradient,
  frxExportPDF, Vcl.ComCtrls, SHLOBJ, FWZipReader;

type
  TFormMain = class(TForm)
    ScriptControlVB: TScriptControl;
    ButtonManualProcessing: TButton;
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
    LabelInvoicePath: TLabel;
    LabelInvoiceMTRpath: TLabel;
    EditInvoicePath: TEdit;
    ButtonInvoicePath: TButton;
    EditInvoiceMTRpath: TEdit;
    ButtonInvoiceMTRpath: TButton;
    RichEditLog: TRichEdit;
    LabelOutput: TLabel;
    EditOutput: TEdit;
    ButtonOutput: TButton;
    statusbarProcessing: TStatusBar;
    buttonSaveLog: TButton;
    saveDialogLog: TSaveDialog;
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
    procedure ButtonInvoicePathClick(Sender: TObject);
    procedure ButtonInvoiceMTRpathClick(Sender: TObject);
    procedure ButtonOutputClick(Sender: TObject);
    procedure buttonSaveLogClick(Sender: TObject);

  private
    { Private declarations }
  public
    function SignatureVerify(inputFileName, inputFileNameSignature: string; out arrayResultsDescription: TStringDynArray): TSmallIntDynArray;
    function SignatureInformation(InputFileNameSignature: string): TStringDynArray;
    function CertificateInformation(InputFileNameSignature: string): TStringDynArray;

    function CheckErrorsWithinArchive(inputArchiveFileName: string): boolean;
    function CheckFileName(inputFileName: string): boolean;

    function ifFileExistsRename(inputFileName: string): string;
    function ifFolderExistsRename(inputFolderName: string): string;

    function CorrectPath(inputDirectory: string): string;

    function ExtractArchiveItemFileName(inputArchiveItemFileName: string): string; //?????? ?????? ????? ??????????? ?????.
                                                                                   //? ???????? ????? ?????? ?????? ??????????? ???????? ?????,
                                                                                   //? ??????? ????? ????.

    function CreateProtocol(inputFileName: string;
                             inputFileNameSignature: array of string;
                             directoryFiles: string;
                             directoryExportToProcessed: string;
                             directoryExportToInvoice: string;
                             directoryExportToInvoiceMTR: string;
                             inputOriginalArchiveFileName: string): boolean;

    procedure CreateResponceFileToOutput(inputFileName, descriptionError: string);

    procedure UpdateDirectories(inputDirectoryRoot: string);
    procedure SortErrorFiles;
    procedure MoveFilesToErrors(inputFileName: string);

    procedure Processed(inputArchiveFileName: string);
    procedure MoveFilesToProcessedAndOtherFolders(inputArchiveFileName, inputNotSigFile: string; inputSigFilesArray: array of string);

    procedure AddLog(inputString: string; LogType: integer); //LogType ??????:
                                                             //isError ? ???? ?????? ???????
                                                             //isSuccess ? ???? ?????? ???????
                                                             //isInformation ? ???? ?????? ??????
    procedure updateStatusBar;

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
  DirectoryRoot, DirectoryErrors, DirectoryTFOMSErrors, DirectoryProcessed, DirectoryOutput, DirectoryInvoice, DirectoryInvoiceMTR, DirectoryLog: string;
  descriptionErrorArchive: string;
  InvoiceType: integer;
  protocolVerifyStatusResult: integer; //???????? ??? ????????: CONFIRMED ? NOT_CONFIRMED
  statusTotal, statusSignCorrect, statusSignUncorrect, statusError: integer;

const
  SIGN_CORRECT = 1;
  SIGN_UNCORRECT = 3;
  CERTIFICATE_REVOKED = 22;

  SIGN_STATUS_FOR_PROTOCOL = [SIGN_CORRECT, SIGN_UNCORRECT, CERTIFICATE_REVOKED];

  REGULAR_INVOICE = 1;
  MTR_INVOICE = 2;

  CONFIRMED = 1;
  NOT_CONFIRMED = 0;

  isError = 0;
  isSuccess = 1;
  isInformation = 2;

implementation

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
      ShowMessage('???? "VerifyScript.vbs" ??????????? ? ????? ? ??????????. ??? ???? ????????? ?? ??????????.');
      Application.Terminate;
      Exit;
    end;

  DirectoryRoot := CorrectPath(EditPath.Text);
  DirectoryInvoice := CorrectPath(EditInvoicePath.Text);
  DirectoryInvoiceMTR := CorrectPath(EditInvoiceMTRpath.Text);
  DirectoryOutput := CorrectPath(EditOutput.Text);

  UpdateDirectories(DirectoryRoot);

  AddLog('???? ???????? ?????????: ' + DateToStr(Now) + ' ' + TimeToStr(Now) + #13#10, isSuccess);

  //??? StatusBar'?
  statusTotal := 0;
  statusSignCorrect := 0;
  statusSignUncorrect := 0;
  statusError := 0;
  updateStatusBar;

end;

procedure TFormMain.ButtonManualProcessingClick(Sender: TObject);
var SearchResult: TSearchRec;
begin
  ButtonManualProcessing.Enabled := False;
  TimerAutoProcessing.Enabled := False;
  SpeedButtonPlay.Enabled := False;
  SpeedButtonStop.Enabled := False;

  try
    DirectoryRoot := CorrectPath(EditPath.Text);
    DirectoryInvoice := CorrectPath(EditInvoicePath.Text);
    DirectoryInvoiceMTR := CorrectPath(EditInvoiceMTRpath.Text);
    DirectoryOutput := CorrectPath(EditOutput.Text);
    if System.SysUtils.DirectoryExists(DirectoryRoot) = False then
      ShowMessage('????????? ???? ? ?????????? ??? ?????? ???????????????. ????? ?? ??????????')
    else
    if System.SysUtils.DirectoryExists(DirectoryInvoice) = False then
      ShowMessage('????????? ???? ? ?????????? ? ???????. ????? ?? ??????????')
    else
    if System.SysUtils.DirectoryExists(DirectoryInvoiceMTR) = False then
      ShowMessage('????????? ???? ? ?????????? ? ???????-???. ????? ?? ??????????')
    else
    if System.SysUtils.DirectoryExists(DirectoryOutput) = False then
      ShowMessage('????????? ???? ? ?????????? ??? ???????? ???????. ????? ?? ??????????')

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
                  AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ' + descriptionErrorArchive + #13#10, isError);
                  statusTotal := statusTotal + 1;
                  statusError := statusError + 1;
                  updateStatusBar;

                  CreateResponceFileToOutput(SearchResult.Name, descriptionErrorArchive);
                end
              else Processed(SearchResult.Name);
            until FindNext(SearchResult) <> 0;
            FindClose(SearchResult);
          end;

      end;

  finally
    ButtonManualProcessing.Enabled := True;
    SpeedButtonPlay.Enabled := True;
    SpeedButtonStop.Enabled := True;
  end;
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

    arrayIndex := 0;
    for i := 0 to Archive.Count-1 do
      begin
        if Not Archive.Item[i].IsFolder then
          begin
            if LowerCase(ExtractFileExt(Archive.item[i].FileName)) = '.sig' then
              begin
                SetLength(SigFilesArray, arrayIndex + 1);
                SigFilesArray[arrayIndex] := ExtractArchiveItemFileName(Archive.item[i].FileName);
                arrayIndex := arrayIndex + 1;
                Archive.Item[i].Extract(DirectoryRoot, ExtractArchiveItemFileName(Archive.item[i].FileName), '');
              end
            else
              begin
                NotSigFile := ExtractArchiveItemFileName(Archive.item[i].FileName);
                Archive.Item[i].Extract(DirectoryRoot, ExtractArchiveItemFileName(Archive.item[i].FileName), '');
              end;
          end;
      end;

  finally
    Archive.Free;
  end;

  MoveFilesToProcessedAndOtherFolders(inputArchiveFileName, NotSigFile, SigFilesArray);

end;

function TFormMain.CreateProtocol(inputFileName: string;
                                   inputFileNameSignature: array of string;
                                   directoryFiles: string;
                                   directoryExportToProcessed: string;
                                   directoryExportToInvoice: string;
                                   directoryExportToInvoiceMTR: string;
                                   inputOriginalArchiveFileName: string): boolean;
var SignatureFiles: array of TSignatureFile;
    NotSignatureFile: TNotSignatureFile;

    NotSigFileDateTime, SigFileDateTime: TDateTime;
    NotSigFile, SigFile: File of Byte;

    i, j, k, counterCorrectStatus: integer;

    frxNotSigFileName, frxNotSigFileDateCreate, frxNotSigFileSize: TfrxMemoView;
    frxSigFileName, frxSigFileDateCreate, frxSigFileSize: TfrxMemoView;
    frxSigStatus, frxSigInformation, frxCertInformation: TfrxMemoView;

    protocolName: string;

begin
  Result := False;

  try
    NotSignatureFile := TNotSignatureFile.Create;
    NotSignatureFile.Name := directoryFiles + inputFileName;

    FileAge(NotSignatureFile.Name, NotSigFileDateTime, True);
    NotSignatureFile.DateCreate := DateTimeToStr(NotSigFileDateTime);

    AssignFile(NotSigFile, NotSignatureFile.Name);
    FileMode := fmOpenRead; // ??????? ???????, ?????? ??? ?????? ?????????? ????? "?????? ??? ??????".
                            // ????? reset ????????? ????? ????? ? ???????,
                            // ?????? ??? ?? ????????? ????????? ?????????? FileMode,
                            // ?????????? ?? ????? ???????? ????? ??????? reset, ????? ???????? 2 (fmOpenReadWrite).
                            // ???????? ?????????? ???????? ?? 0 (fmOpenRead).
    reset(NotSigFile);
    NotSignatureFile.Size := IntToStr(FileSize(NotSigFile)) + ' ????';
    CloseFile(NotSigFile);

    SetLength(SignatureFiles, Length(InputFileNameSignature));
    AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ?????? ???????? ???????? ?? ??????: "' + inputOriginalArchiveFileName + '". ????? ???????? ????? ??????? ?????? ?? 30 ??????. ????????...' + #13#10, isInformation);
    for i := 0 to High(SignatureFiles) do
      begin
        SignatureFiles[i] := TSignatureFile.Create;
        SignatureFiles[i].Name := directoryFiles + inputFileNameSignature[i];

        FileAge(SignatureFiles[i].Name, SigFileDateTime, True);
        SignatureFiles[i].DateCreate := DateTimeToStr(SigFileDateTime);

        AssignFile(SigFile, SignatureFiles[i].Name);
        reset(SigFile);
        SignatureFiles[i].Size := IntToStr(FileSize(SigFile)) + ' ????';
        CloseFile(SigFile);

        SignatureFiles[i].CertificateInformation := CertificateInformation(SignatureFiles[i].Name);
        SignatureFiles[i].SignatureInformation := SignatureInformation(SignatureFiles[i].Name);

        SignatureFiles[i].VerifyStatus := SignatureVerify(NotSignatureFile.Name, SignatureFiles[i].Name, SignatureFiles[i].VerifyStatusDesctiption);
        For j := 0 to High(SignatureFiles[i].VerifyStatusDesctiption) do
          begin
            IF (SignatureFiles[i].VerifyStatus[j] in SIGN_STATUS_FOR_PROTOCOL) then
              BEGIN
                Result := True;
                if SignatureFiles[i].VerifyStatus[j] = SIGN_CORRECT then
                  SignatureFiles[i].VerifyStatusDesctiption[j] := '?????? ???????? ??????? ' + '?' + IntToStr(j+1) + ': '
                                                                + SignatureFiles[i].VerifyStatusDesctiption[j] + #13#10
                                                                + '??????? ????????????' + #13#10 + #13#10
                else
                  SignatureFiles[i].VerifyStatusDesctiption[j] := '?????? ???????? ??????? ' + '?' + IntToStr(j+1) + ': '
                                                                + SignatureFiles[i].VerifyStatusDesctiption[j] + #13#10
                                                                + '??????? ?? ????????????' + #13#10 + #13#10
              END
            ELSE
              BEGIN
                Result := False;
                AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ?????? ??? ???????? ??????? ' + SignatureFiles[i].Name + '. ' +
                       SignatureFiles[i].VerifyStatusDesctiption[j] + #13#10, isError);
                statusTotal := statusTotal + 1;
                statusError := statusError + 1;
                updateStatusBar;
                Exit;
              END;
          end;

        //???? ?????? ????? *.sig ????????? ???? ?? ???? ???????????? ???????,
        //?? ???????????? ?????? ????????? ??? ???????????????? ????????
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
                protocolVerifyStatusResult := NOT_CONFIRMED;
                Break;
              end;
          end;
        if counterCorrectStatus = Length(SignatureFiles[i].VerifyStatus) then
          begin
            frxReportTypeProtocol := frxReportProtocolConfirmed;
            protocolName := 'ProtocolConfirmed_';
            protocolVerifyStatusResult := CONFIRMED;
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
              begin
                frxSigStatus.Memo.Text := frxSigStatus.Memo.Text + SignatureFiles[i].VerifyStatusDesctiption[j];
                AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ????????? ??????? "' + ExtractFileName(SignatureFiles[i].Name) + '". ' + TrimRight(SignatureFiles[i].VerifyStatusDesctiption[j]) + #13#10, isSuccess);
                statusTotal := statusTotal + 1;
                statusSignCorrect := statusSignCorrect + 1;
                updateStatusBar;
              end
            else
              begin
                frxSigStatus.Memo.Text := frxSigStatus.Memo.Text + SignatureFiles[i].VerifyStatusDesctiption[j];
                AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ????????? ??????? "' + ExtractFileName(SignatureFiles[i].Name) + '". ' + TrimRight(SignatureFiles[i].VerifyStatusDesctiption[j]) + #13#10, isError);
                statusTotal := statusTotal + 1;
                statusSignUncorrect := statusSignUncorrect + 1;
                updateStatusBar;
              end
          end;
        frxCertInformation.Memo.Text := TrimRight(frxCertInformation.Memo.Text);
        frxSigInformation.Memo.Text := TrimRight(frxSigInformation.Memo.Text);
        frxSigStatus.Memo.Text := TrimRight(frxSigStatus.Memo.Text);

        frxReportTypeProtocol.PrepareReport(true);
        frxPDFexportProtocol.Compressed := True;
        frxPDFexportProtocol.Background := True;
        frxPDFexportProtocol.PrintOptimized := False;
        frxPDFexportProtocol.OpenAfterExport := False;
        frxPDFexportProtocol.ShowProgress := False;
        frxPDFexportProtocol.ShowDialog := False;

        frxPDFexportProtocol.FileName := directoryExportToProcessed + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
        //????????? ???????? ? Processed
        frxReportTypeProtocol.Export(frxPDFexportProtocol);

        //????????? ?????????? ?? ????? "Output"
        //????? ??? ??? ? ??? ??????????? ????????
        if System.SysUtils.DirectoryExists(DirectoryOutput) = False then
          System.SysUtils.ForceDirectories(DirectoryOutput);
        //???????? ?????????? ?? ? ????? "Output" ???? ? ????? ?? ?????????
        //???? ??????????, ???????? ????????
        frxPDFexportProtocol.FileName := DirectoryOutput + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
        frxPDFexportProtocol.FileName := ifFileExistsRename(frxPDFexportProtocol.FileName);
        //????????? ???????? ? Output
        frxReportTypeProtocol.Export(frxPDFexportProtocol);

        if (protocolVerifyStatusResult = CONFIRMED) and (InvoiceType = REGULAR_INVOICE) then
          begin
            //????????? ???????? ? Invoice
            if System.SysUtils.DirectoryExists(DirectoryExportToInvoice) = False then
              System.SysUtils.ForceDirectories(DirectoryExportToInvoice);
            frxPDFexportProtocol.FileName := directoryExportToInvoice + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
            frxReportTypeProtocol.Export(frxPDFexportProtocol);
          end;
        if (protocolVerifyStatusResult = CONFIRMED) and (InvoiceType = MTR_INVOICE) then
          begin
            //????????? ???????? ? InvoiceMTR
            if System.SysUtils.DirectoryExists(DirectoryExportToInvoiceMTR) = False then
              System.SysUtils.ForceDirectories(DirectoryExportToInvoiceMTR);
            frxPDFexportProtocol.FileName := directoryExportToInvoiceMTR + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
            frxReportTypeProtocol.Export(frxPDFexportProtocol);
          end;
      end;
  finally
    NotSignatureFile.Free;
    if Length(SignatureFiles) > 0 then
      for k := 0 to High(SignatureFiles) do
        SignatureFiles[k].Free;
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
          1 : arrayResultsD[i] := '?????';
          3 : arrayResultsD[i] := '??????? ??????????? ??? ? ??? ??? ???????';
          5 : arrayResultsD[i] := '????? ???? ???????? ???????????';
          15 : arrayResultsD[i] := '??????? ????????????? ?????????, ?? ??? ??????? ??????? ? ?????? ??? ?????????? ???????????? ???????. ???????? ?????????? ???????? ???????? ?????????? ??? TSL ????? ????????? ????????? (????????? --> ??????)';
          22 : arrayResultsD[i] := '?????????? ???????';
          23 : arrayResultsD[i] := '???? ??? ????????? ???????? ??????????? ??? ??? ???????. ???????? ?????????? ???????? ???????? ?????????? ? ????????????? ?????????. ???? ?????????? ?????????????? ????????? ?? ????? ????? ??????';
        else arrayResultsD[i] := '?????? ?? ?????????. ????? ??????????? ?????? ' + IntToStr(arrayResults[i]) + '.';
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
    descriptionError: string;
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
            descriptionError := '???????? ??? ?/??? ?????????? ????? "' + SearchResult.Name + '"';
            AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ' + descriptionError + #13#10, isError);
            statusTotal := statusTotal + 1;
            statusError := statusError + 1;
            updateStatusBar;

            CreateResponceFileToOutput(SearchResult.Name, descriptionError + #13#10
                                                          + '?????? ????? ? ?????????? ?????? ?????? ??????????????? ?????? (???????? ?????? ?????????):' + #13#10
                                                          + 'SH_{??? ??}_{??? ???}_{????????/???????}.zip' + #13#10
                                                          + 'SH3_{??? ??}_{??? ???}_{????????/???????}.zip' + #13#10
                                                          + 'SHO_{??? ??}_{??? ???}_{????????/???????}.zip' + #13#10
                                                          + 'SHUD_{??? ??}_{??? ???}_{????????/???????}.zip' + #13#10
                                                          + 'SMP_{??? ??}_{??? ???}_{????????/???????}.zip' + #13#10
                                                          + 'SHCP_{??? ??}_{??? ???}_????????.zip' + #13#10
                                                          + 'MSHO_{??? ??}_???_{????????/???????}.zip' + #13#10
                                                          + 'MSH_{??? ??}_???_{????????/???????}.zip' + #13#10
                                                          + 'MSMP_{??? ??}_???_{????????/???????}.zip');
          end;
      until FindNext(SearchResult) <> 0;
      FindClose(SearchResult);
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

procedure TFormMain.MoveFilesToProcessedAndOtherFolders(inputArchiveFileName: string; inputNotSigFile: string; inputSigFilesArray: array of string);
var DirectoryFrom, DirectoryToProcessed, DirectoryToInvoice, DirectoryToInvoiceMTR, DirectoryToTFOMSErrors: string;
    fileDirectoryFrom, fileDirectoryToProcessed, fileDirectoryToInvoice, fileDirectoryToInvoiceMTR, fileDirectoryToTFOMSErrors: string;
    pointerFileDirectoryFrom, pointerFileDirectoryToProcessed, pointerFileDirectoryToInvoice, pointerFileDirectoryToInvoiceMTR, pointerFileDirectoryToTFOMSErrors: PWideChar;
    MO, SMO: string;
    Year: integer;
    Month: string;
    i: integer;
begin
  Year := YearOf(Date);
  case MonthOf(Date) of
    1 : Month := '??????';
    2 : Month := '???????';
    3 : Month := '????';
    4 : Month := '??????';
    5 : Month := '???';
    6 : Month := '????';
    7 : Month := '????';
    8 : Month := '??????';
    9 : Month := '????????';
    10 : Month := '???????';
    11 : Month := '??????';
    12 : Month := '???????';
  else Month := '??????????? ????? ??';
  end;
  MO := Copy(inputArchiveFileName, AnsiPos('_', inputArchiveFileName) + 1, 6);

  DirectoryFrom := DirectoryRoot;

  DirectoryToProcessed := DirectoryProcessed + IntToStr(Year) + '\' + Month + '\' + MO + '\' +
                          StringReplace(inputArchiveFileName, ExtractFileExt(inputArchiveFileName), '', [rfIgnoreCase]) + '\';
  DirectoryToProcessed := ifFolderExistsRename(DirectoryToProcessed);
  if System.SysUtils.DirectoryExists(DirectoryToProcessed) = False then
    System.SysUtils.ForceDirectories(DirectoryToProcessed);

  if (AnsiPos('MTP', UpperCase(inputArchiveFileName)) = 0) and
     (AnsiPos('???', UpperCase(inputArchiveFileName)) = 0) then //????????? "???" ? ????? ????? ??? ??-??????, ??? ? ??-?????????
    InvoiceType := REGULAR_INVOICE
  else
    InvoiceType := MTR_INVOICE;

  if InvoiceType = REGULAR_INVOICE then
    begin
      SMO := Copy(inputArchiveFileName, AnsiPos(MO, inputArchiveFileName)+7, 5);
      DirectoryToInvoice := DirectoryInvoice + IntToStr(Year) + '\' + Month + '\' + SMO + '\' +
                            StringReplace(inputArchiveFileName, ExtractFileExt(inputArchiveFileName), '', [rfIgnoreCase]) + '\';
      DirectoryToInvoice := ifFolderExistsRename(DirectoryToInvoice);
    end
  else
    begin
      DirectoryToInvoiceMTR := DirectoryInvoiceMTR + IntToStr(Year) + '\' + Month + '\' + MO + '\' +
                               StringReplace(inputArchiveFileName, ExtractFileExt(inputArchiveFileName), '', [rfIgnoreCase]) + '\';
      DirectoryToInvoiceMTR := ifFolderExistsRename(DirectoryToInvoiceMTR);
    end;

  //??????? ????????
  IF CreateProtocol(inputNotSigFile, inputSigFilesArray, DirectoryFrom, DirectoryToProcessed, DirectoryToInvoice, DirectoryToInvoiceMTR, inputArchiveFileName) = True THEN
    BEGIN
      //???????? ? ????????? ????? ? ????? "Processed":

      //? ????????? ???????????? zip-????
      fileDirectoryFrom := DirectoryFrom + inputArchiveFileName;
      pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
      fileDirectoryToProcessed := DirectoryToProcessed + inputArchiveFileName;
      pointerFileDirectoryToProcessed := Addr(fileDirectoryToProcessed[1]);
      MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryToProcessed);

      //? ???????? ????-???? ? ????? ? ??????? / ???-??????? ? ????????? ??? ? ????? Processed
      fileDirectoryFrom := DirectoryFrom + inputNotSigFile;
      pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
      if (InvoiceType = REGULAR_INVOICE) and (protocolVerifyStatusResult = CONFIRMED) then
        begin
          fileDirectoryToInvoice := DirectoryToInvoice + inputNotSigFile;
          pointerFileDirectoryToInvoice := Addr(fileDirectoryToInvoice[1]);
          CopyFile(pointerFileDirectoryFrom, pointerFileDirectoryToInvoice, false);
        end;
      if (InvoiceType = MTR_INVOICE) and (protocolVerifyStatusResult = CONFIRMED) then
        begin
          fileDirectoryToInvoiceMTR := DirectoryToInvoiceMTR + inputNotSigFile;
          pointerFileDirectoryToInvoiceMTR := Addr(fileDirectoryToInvoiceMTR[1]);
          CopyFile(pointerFileDirectoryFrom, pointerFileDirectoryToInvoiceMTR, false);
        end;
      fileDirectoryToProcessed := DirectoryToProcessed + inputNotSigFile;
      pointerFileDirectoryToProcessed := Addr(fileDirectoryToProcessed[1]);
      MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryToProcessed);

      //? ???????? sig-????? ? ????? ? ??????? / ???-??????? ? ????????? ?? ? ????? Processed
      For i := 0 to High(inputSigFilesArray) do
        begin
          fileDirectoryFrom := DirectoryFrom + inputSigFilesArray[i];
          pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
          if (InvoiceType = REGULAR_INVOICE) and (protocolVerifyStatusResult = CONFIRMED) then
            begin
              fileDirectoryToInvoice := DirectoryToInvoice + inputSigFilesArray[i];
              pointerFileDirectoryToInvoice := Addr(fileDirectoryToInvoice[1]);
              CopyFile(pointerFileDirectoryFrom, pointerFileDirectoryToInvoice, false);
            end;
          if (InvoiceType = MTR_INVOICE) and (protocolVerifyStatusResult = CONFIRMED) then
            begin
              fileDirectoryToInvoiceMTR := DirectoryToInvoiceMTR + inputSigFilesArray[i];
              pointerFileDirectoryToInvoiceMTR := Addr(fileDirectoryToInvoiceMTR[1]);
              CopyFile(pointerFileDirectoryFrom, pointerFileDirectoryToInvoiceMTR, false);
            end;
          fileDirectoryToProcessed := DirectoryToProcessed + inputSigFilesArray[i];
          pointerFileDirectoryToProcessed := Addr(fileDirectoryToProcessed[1]);
          MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryToProcessed);
        end;
    END
  ELSE //???? ???????? ?? ???????? ??-?? ????????? ??????, ?????????? ????? ? ????? "TFOMS Errors" ??? ???????????? ????
    BEGIN
      if System.SysUtils.DirectoryExists(DirectoryTFOMSErrors) = False then
          System.SysUtils.ForceDirectories(DirectoryTFOMSErrors);

      //? ????? "TFOMS Errors" ??????? ????? ? ?????? ????????????? zip-?????, ??????????? ??
      DirectoryToTFOMSErrors := DirectoryTFOMSErrors + StringReplace(inputArchiveFileName, ExtractFileExt(inputArchiveFileName), '', [rfIgnoreCase]) + '\';
      DirectoryToTFOMSErrors := ifFolderExistsRename(DirectoryToTFOMSErrors);
      if System.SysUtils.DirectoryExists(DirectoryToTFOMSErrors) = False then
          System.SysUtils.ForceDirectories(DirectoryToTFOMSErrors);

      //????????? ????? ? ????? "TFOMS Errors":

      //- ????????? zip-????
      fileDirectoryFrom := DirectoryFrom + inputArchiveFileName;
      pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
      fileDirectoryToTFOMSErrors := DirectoryToTFOMSErrors + inputArchiveFileName;
      pointerFileDirectoryToTFOMSErrors := Addr(fileDirectoryToTFOMSErrors[1]);
      MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryToTFOMSErrors);
      MoveFile('E:\Proba\AutoProcessingFiles\SH_830082_83001_????????.zip', 'E:\Proba\AutoProcessingFiles\TFOMS Errors\SH_830082_83001_????????\SH_830082_83001_????????.zip');

      //- ????????? ????-????
      fileDirectoryFrom := DirectoryFrom + inputNotSigFile;
      pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
      fileDirectoryToTFOMSErrors := DirectoryToTFOMSErrors + inputNotSigFile;
      pointerFileDirectoryToTFOMSErrors := Addr(fileDirectoryToTFOMSErrors[1]);
      MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryToTFOMSErrors);

      //- ????????? sig-?????
      For i := 0 to High(inputSigFilesArray) do
        begin
          fileDirectoryFrom := DirectoryFrom + inputSigFilesArray[i];
          pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
          fileDirectoryToTFOMSErrors := DirectoryToTFOMSErrors + inputSigFilesArray[i];
          pointerFileDirectoryToTFOMSErrors := Addr(fileDirectoryToTFOMSErrors[1]);
          MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryToTFOMSErrors);
        end;

      AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ????? ?????????? ? ????? "TFOMS Errors" ??? ???????????? ??????' + #13#10 + #13#10, isError);
    END;

end;

function TFormMain.CheckFileName(inputFileName: string): boolean;
begin
  Result := False;

  if MatchesMask(inputFileName, 'SH_*_*_*.zip') or
     MatchesMask(inputFileName, 'SH3_*_*_*.zip') or
     MatchesMask(inputFileName, 'SHO_*_*_*.zip') or
     MatchesMask(inputFileName, 'SHUD_*_*_*.zip') or
     MatchesMask(inputFileName, 'SMP_*_*_*.zip') or
     MatchesMask(inputFileName, 'SHCP_*_*_????????.zip') or
     MatchesMask(inputFileName, 'MSHO_*_MTP_*.zip') or //MTP ? ??-?????????
     MatchesMask(inputFileName, 'MSHO_*_???_*.zip') or //MTP ? ??-??????
     MatchesMask(inputFileName, 'MSH_*_MTP_*.zip') or //MTP ? ??-?????????
     MatchesMask(inputFileName, 'MSH_*_???_*.zip') or //MTP ? ??-??????
     MatchesMask(inputFileName, 'MSMP_*_MTP_*.zip') or //MTP ? ??-?????????
     MatchesMask(inputFileName, 'MSMP_*_???_*.zip') then //MTP ? ??-??????
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
    //???????? ?? ?????????? ??????, ??????????? ? zip-?????? ??? ??????????. ?? ?????????? ? ?????? ?????? ???? 1 ????-???? ??? ??????????.
    for i := 0 to Archive.Count-1 do
      begin
        if Not Archive.Item[i].IsFolder then
          begin
            if LowerCase(ExtractFileExt(Archive.item[i].FileName)) <> '.sig' then
              Counter := Counter + 1;
          end;
      end;
    if Counter > 1 then
      begin
        Result := True;
        descriptionErrorArchive := '? zip-?????? "' + inputArchiveFileName + '" ????? ?????? ?????-????? ??? ??????????.';
      end;
    if Counter = 0 then
      begin
        Result := True;
        descriptionErrorArchive := '? zip-?????? "' + inputArchiveFileName + '" ??????????? ????-???? ??? ??????????.';
      end;

    //???????? ?? ?????????? ????????. ???? ? zip-?????? ??????? ???????????, ?? ? ?????.
    Counter := 0;
    for i := 0 to Archive.Count-1 do
      begin
        if LowerCase(ExtractFileExt(Archive.item[i].FileName)) = '.sig' then
          counter := Counter + 1;
      end;
    if Counter = 0 then
      begin
        Result := True;
        descriptionErrorArchive := '? zip-?????? "' + inputArchiveFileName + '" ??????????? ?????-??????? ? ??????????? ".sig"';
      end;

    //???????? ?? ???????????? ???? ?????? ?????? zip-??????
    for i := 0 to Archive.Count-1 do
      begin
        if Not Archive.Item[i].IsFolder then
          begin
            if ( LowerCase(ExtractFileExt(Archive.Item[i].FileName)) <> '.sig' ) and
               ( Not MatchesMask(ExtractArchiveItemFileName(Archive.item[i].FileName), StringReplace(inputArchiveFileName, ExtractFileExt(inputArchiveFileName), '', [rfIgnoreCase]) + '*') ) then
              begin
                Result := True;
                descriptionErrorArchive := '????-???? "' + ExtractArchiveItemFileName(Archive.Item[i].FileName) + '" ?????? zip-?????? "' + inputArchiveFileName + '" ?? ????????????? ??? ????????';
              end;
          end;
      end;

  finally
    Archive.Free;
  end;

end;

procedure TFormMain.CreateResponceFileToOutput(inputFileName: string; descriptionError: string);
var responceTextFile: TextFile;
    responceTextFileName: string;
begin
  if System.SysUtils.DirectoryExists(DirectoryOutput) = False then
    System.SysUtils.ForceDirectories(DirectoryOutput);
  responceTextFileName := DirectoryOutput + 'response_' + StringReplace(inputFileName, ExtractFileExt(inputFileName), '', [rfIgnoreCase]) + '.txt';
  responceTextFileName := ifFileExistsRename(responceTextFileName);
  AssignFile(responceTextFile, responceTextFileName);
  ReWrite(responceTextFile);
  WriteLn(responceTextFile, descriptionError);
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
          Insert(' (' + IntToStr(counterName) + ')', inputFileName, Length(inputFileName)-Length(ExtractFileExt(inputFileName))+1);
          result := inputFileName;
        end
      else
        begin
          inputFileName := StringReplace(inputFileName, ' (' + IntToStr(counterName-1) + ')', ' (' + IntToStr(counterName) + ')', []);
          result := inputFileName;
        end;
    end;
end;

function TFormMain.ifFolderExistsRename(inputFolderName: string): string;
var counterName: integer;
begin
  result := inputFolderName;

  counterName := 0;
  while System.SysUtils.DirectoryExists(inputFolderName) do
    begin
      counterName := counterName + 1;
      if counterName = 1 then
        begin
          Insert(' (' + IntToStr(counterName) + ')', inputFolderName, Length(inputFolderName));
          result := inputFolderName;
        end
      else
        begin
          inputFolderName := StringReplace(inputFolderName, ' (' + IntToStr(counterName-1) +')', ' (' + IntToStr(counterName) + ')', []);
          result := inputFolderName;
        end;
    end;

end;

function TFormMain.CorrectPath(inputDirectory: string): string;
begin
  if inputDirectory = '' then
    Result := ''
  else
    begin
      inputDirectory := Trim(inputDirectory);

      if Pos('/', inputDirectory) <> 0 then
        begin
          inputDirectory := StringReplace(inputDirectory, '/', '\', [rfReplaceAll]);
        end;

      if inputDirectory[length(inputDirectory)] <> '\' then
        Result := inputDirectory + '\'
      else
        Result := inputDirectory;
    end;
end;

function TFormMain.ExtractArchiveItemFileName(inputArchiveItemFileName: string): string;  //?????? ?????? ????? ??????????? ?????.
                                                                                          //? ???????? ????? ?????? ?????? ??????????? ???????? ?????,
                                                                                          //? ??????? ????? ????.
begin
  result := ExtractFileName(StringReplace(inputArchiveItemFileName, '/', '\', []))
end;

procedure TFormMain.UpdateDirectories(inputDirectoryRoot: string);
begin
  DirectoryErrors := DirectoryRoot + 'Errors';
  DirectoryErrors := CorrectPath(DirectoryErrors);

  DirectoryTFOMSErrors := DirectoryRoot + 'TFOMS Errors';
  DirectoryTFOMSErrors := CorrectPath(DirectoryTFOMSErrors);

  DirectoryProcessed := DirectoryRoot + 'Processed';
  DirectoryProcessed := CorrectPath(DirectoryProcessed);

  DirectoryLog := DirectoryRoot + 'Log';
  DirectoryLog := CorrectPath(DirectoryLog);
end;

procedure TFormMain.ButtonPathClick(Sender: TObject);
begin
  if SelectDirectory('???????? ????? ??? ?????? ???????????????:', '', DirectoryRoot, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditPath.Text := DirectoryRoot;
end;

procedure TFormMain.buttonSaveLogClick(Sender: TObject);
begin
  saveDialogLog.Title := '???????? ????? ???? ????????? ??? ???:';
  saveDialogLog.Filter := 'RTF-????|*.rtf|????????? ????|*.txt';

  if saveDialogLog.FilterIndex = 1 then
    begin
      RichEditLog.PlainText := False; //????? ???????? ? RTF ?????????? ??????? False,
                                      //????? ???????? ?????????? ? ????-??????? (????, ???????? ?????? ? ?.?.)
      saveDialogLog.DefaultExt := 'rtf';
    end
  else
    begin
      RichEditLog.PlainText := True; //????? ???????? ? TXT ?????????? ??????? True,
                                     //????? ? ???????? ?????????? ?????? ????? ??? ????-??????
      saveDialogLog.DefaultExt := 'txt';
    end;

  if not saveDialogLog.Execute then
    exit
  else
    RichEditLog.Lines.SaveToFile(saveDialogLog.FileName);

  RichEditLog.PlainText := True; //?????? ?????????? ???????? ?? False ??? ?????????????? ????? ? TXT-???????
end;

procedure TFormMain.ButtonInvoicePathClick(Sender: TObject);
begin
  if SelectDirectory('???????? ????? ??? ???????? ??????:', '', DirectoryInvoice, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditInvoicePath.Text := DirectoryInvoice;
end;

procedure TFormMain.ButtonInvoiceMTRpathClick(Sender: TObject);
begin
  if SelectDirectory('???????? ????? ??? ???????? ??????-???:', '', DirectoryInvoiceMTR, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditInvoiceMTRpath.Text := DirectoryInvoiceMTR;
end;

procedure TFormMain.ButtonOutputClick(Sender: TObject);
begin
  if SelectDirectory('???????? ????? ??? ???????? ?????????? ? ?????? ? ????????:', '', DirectoryOutput, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditOutput.Text := DirectoryOutput;
end;

procedure TFormMain.SpeedButtonPlayClick(Sender: TObject);
begin
  buttonManualProcessing.Enabled := False;

  SpeedButtonPlay.Visible := False;
  SpeedButtonStop.Visible := True;

  if (SpinEditSec.Value > 59) or
     (SpinEditSec.Value < 0) or
     (SpinEditMin.Value < 0) or
     ( (SpinEditSec.Value = 0) and (SpinEditMin.Value = 0) ) then
    ShowMessage('??????? ??????? ??????/???????')
  else
    begin
      TimerAutoProcessing.Interval := SpinEditMin.Value * 60000 + SpinEditSec.Value * 1000;
      TimerAutoProcessing.Enabled := True;

      TimerAutoProcessingState.Enabled := True;
      LabelAutoProcessingState.Caption := '?????????????? ???????';
    end;
end;

procedure TFormMain.SpeedButtonStopClick(Sender: TObject);
begin
  buttonManualProcessing.Enabled := True;

  SpeedButtonStop.Visible := False;
  SpeedButtonPlay.Visible := True;

  TimerAutoProcessingState.Enabled := False;
  LabelAutoProcessingState.Caption := '?????????????? ?? ???????';

  TimerAutoProcessing.Enabled := False;
end;

procedure TFormMain.TimerAutoProcessingStateTimer(Sender: TObject);
begin
  LabelAutoProcessingState.Caption := LabelAutoProcessingState.Caption + '.';
  if LabelAutoProcessingState.Caption = '?????????????? ???????....' then
    LabelAutoProcessingState.Caption := '?????????????? ???????';
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

procedure TFormMain.AddLog(inputString: string; LogType: integer); //LogType ??????:
                                                                   //isError ? ???? ?????? ???????
                                                                   //isSuccess ? ???? ?????? ??????, ??????
                                                                   //isInformation ? ???? ?????? ??????
var logFile: TextFile;
    logFileName: string;
    Month, Year, Day: string;
begin
  case LogType of
    isError: begin
               RichEditLog.SelAttributes.Color := clRed;
               RichEditLog.Lines.Add(inputString);
               RichEditLog.Refresh;
             end;
    isSuccess: begin
                 RichEditLog.SelAttributes.Color := clBlack;
                 RichEditLog.SelAttributes.Style := [fsItalic];
                 RichEditLog.Lines.Add(inputString);
                 RichEditLog.Refresh;
               end;
    isInformation: begin
                     RichEditLog.SelAttributes.Color := clBlack;
                     RichEditLog.Lines.Add(inputString);
                     RichEditLog.Refresh;
                   end;
  end;

  if System.SysUtils.DirectoryExists(DirectoryLog) = False then
    System.SysUtils.ForceDirectories(DirectoryLog);

  Year := IntToStr(YearOf(Date));
  if MonthOf(Date) < 10 then
    Month := '0' + IntToStr(MonthOf(Date))
  else
    Month := IntToStr(MonthOf(Date));
  if DayOf(Date) < 10 then
    Day := '0' + IntToStr(DayOf(Date))
  else
    Day := IntToStr(DayOf(Date));

  logFileName := DirectoryLog + Day + '-' + Month + '-' + Year + '.txt';

  AssignFile(logFile, logFileName);
  if FileExists(logFileName) = True then
    begin
      Append(logFile);
      WriteLN(logFile, inputString);
      CloseFile(logFile);
    end
  else
    begin
      ReWrite(logFile);
      WriteLN(logFile, inputString);
      CloseFile(logFile);
    end;
end;

procedure TFormMain.updateStatusBar;
begin
  statusbarProcessing.Panels.Items[0].Text := '?????: ' + IntToStr(statusTotal);
  statusbarProcessing.Panels.Items[1].Text := '?? ???????? "?????": ' + IntToStr(statusSignCorrect);
  statusbarProcessing.Panels.Items[2].Text := '?? ???????? "??????? ?? ????????????": ' + IntToStr(statusSignUncorrect);
  statusbarProcessing.Panels.Items[3].Text := '??????: ' + IntToStr(statusError);
end;

end.
