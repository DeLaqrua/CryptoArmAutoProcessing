unit uCryptoArmAutoprocessing;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  System.StrUtils, System.Types,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.OleCtrls, MSScriptControl_TLB,
  Vcl.StdCtrls, ActiveX, Vcl.FileCtrl, System.Masks, DateUtils,
  Vcl.Buttons, Vcl.Samples.Spin, Vcl.ExtCtrls, frxClass, frxGradient,
  frxExportPDF, Vcl.ComCtrls;

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

    procedure CreateResponceFileToOutput(inputFileName, descriptionError: string);

    procedure CreateProtocol(inputFileName: string;
                             inputFileNameSignature: array of string;
                             directoryFiles: string;
                             directoryExportToProcessed: string;
                             directoryExportToInvoice: string;
                             directoryExportToInvoiceMTR: string;
                             inputOriginalArchiveFileName: string);

    procedure UpdateDirectories(inputDirectoryRoot: string);
    procedure SortErrorFiles;
    procedure MoveFilesToErrors(inputFileName: string);

    procedure Processed(inputArchiveFileName: string);
    procedure MoveFilesToProcessedAndOtherFolders(inputArchiveFileName, inputNotSigFile: string; inputSigFilesArray: array of string);

    procedure AddLog(inputString: string; LogType: integer); //LogType бывает:
                                                             //isError – цвет текста Красный
                                                             //isSuccess – цвет текста Зелёный
                                                             //isInformation – цвет текста Чёрный
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
  DirectoryRoot, DirectoryErrors, DirectoryProcessed, DirectoryOutput, DirectoryInvoice, DirectoryInvoiceMTR: string;
  descriptionErrorArchive: string;
  InvoiceType: integer;
  protocolVerifyStatusResult: integer; //содержит два значения: CONFIRMED и NOT_CONFIRMED
const
  SIGN_CORRECT = 1;
  REGULAR_INVOICE = 1;
  MTR_INVOICE = 2;

  CONFIRMED = 1;
  NOT_CONFIRMED = 0;

  isError = 0;
  isSuccess = 1;
  isInformation = 2;

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
      ShowMessage('Файл "VerifyScript.vbs" отсутствует в папке с программой. Без него программа не запустится.');
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
  DirectoryInvoice := CorrectPath(EditInvoicePath.Text);
  DirectoryInvoiceMTR := CorrectPath(EditInvoiceMTRpath.Text);
  DirectoryOutput := CorrectPath(EditOutput.Text);
  if (System.SysUtils.DirectoryExists(DirectoryRoot) = False) or
     (System.SysUtils.DirectoryExists(DirectoryInvoice) = False) or
     (System.SysUtils.DirectoryExists(DirectoryInvoiceMTR) = False) or
     (System.SysUtils.DirectoryExists(DirectoryOutput) = False) then
    ShowMessage('Проверьте путь к директории. Папки не существует.')
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

                CreateResponceFileToOutput(SearchResult.Name, descriptionErrorArchive);
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

  MoveFilesToProcessedAndOtherFolders(inputArchiveFileName, NotSigFile, SigFilesArray);

end;

procedure TFormMain.CreateProtocol(inputFileName: string;
                                   inputFileNameSignature: array of string;
                                   directoryFiles: string;
                                   directoryExportToProcessed: string;
                                   directoryExportToInvoice: string;
                                   directoryExportToInvoiceMTR: string;
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
  NotSignatureFile.Size := IntToStr(FileSize(NotSigFile)) + ' байт';
  CloseFile(NotSigFile);

  SetLength(SignatureFiles, Length(InputFileNameSignature));
  AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  Начало проверки подписей из архива: "' + inputOriginalArchiveFileName + '". Время проверки одной длится подписи до 30 секунд. Ожидайте...' + #13#10, isInformation);
  for i := 0 to High(SignatureFiles) do
    begin
      SignatureFiles[i] := TSignatureFile.Create;
      SignatureFiles[i].Name := directoryFiles + inputFileNameSignature[i];

      FileAge(SignatureFiles[i].Name, SigFileDateTime, True);
      SignatureFiles[i].DateCreate := DateTimeToStr(SigFileDateTime);

      AssignFile(SigFile, SignatureFiles[i].Name);
      reset(SigFile);
      SignatureFiles[i].Size := IntToStr(FileSize(SigFile)) + ' байт';
      CloseFile(SigFile);

      SignatureFiles[i].CertificateInformation := CertificateInformation(SignatureFiles[i].Name);
      SignatureFiles[i].SignatureInformation := SignatureInformation(SignatureFiles[i].Name);

      SignatureFiles[i].VerifyStatus := SignatureVerify(NotSignatureFile.Name, SignatureFiles[i].Name, SignatureFiles[i].VerifyStatusDesctiption);
      For j := 0 to High(SignatureFiles[i].VerifyStatusDesctiption) do
        begin
          if SignatureFiles[i].VerifyStatus[j] = SIGN_CORRECT then
            SignatureFiles[i].VerifyStatusDesctiption[j] := 'Статус проверки подписи ' + '№' + IntToStr(j+1) + ': '
                                                          + SignatureFiles[i].VerifyStatusDesctiption[j] + #13#10
                                                          + 'ПОДПИСЬ ПОДТВЕРЖДЕНА' + #13#10 + #13#10
          else
            SignatureFiles[i].VerifyStatusDesctiption[j] := 'Статус проверки подписи ' + '№' + IntToStr(j+1) + ': '
                                                          + SignatureFiles[i].VerifyStatusDesctiption[j] + #13#10
                                                          + 'ПОДПИСЬ НЕ ПОДТВЕРЖДЕНА' + #13#10 + #13#10
        end;

      //Если внутри файла *.sig попадётся хотя бы одна некорректная подпись,
      //то используется шаблон протокола для неподтверждённых подписей
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
              //В будущем в этой строчке добавится Зелёный цвет текста протокола <<< Не забыть!
              AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  Проверена подпись "' + ExtractFileName(SignatureFiles[i].Name) + '". ' + TrimRight(SignatureFiles[i].VerifyStatusDesctiption[j]) + #13#10, isSuccess);
            end
          else
            begin
              frxSigStatus.Memo.Text := frxSigStatus.Memo.Text + SignatureFiles[i].VerifyStatusDesctiption[j];
              //В будущем в этой строчке добавится Красный цвет текста протокола <<< Не забыть!
              AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  Проверена подпись "' + ExtractFileName(SignatureFiles[i].Name) + '". ' + TrimRight(SignatureFiles[i].VerifyStatusDesctiption[j]) + #13#10, isError);
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
      //Формируем протокол в Processed
      frxReportTypeProtocol.Export(frxPDFexportProtocol);

      //Проверяем существует ли папка "Output"
      //перед тем как в неё переместить протокол
      if System.SysUtils.DirectoryExists(DirectoryOutput) = False then
        System.SysUtils.ForceDirectories(DirectoryOutput);
      //Проверка существует ли в папке "Output" файл с таким же названием
      //Если существует, название меняется
      frxPDFexportProtocol.FileName := DirectoryOutput + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
      frxPDFexportProtocol.FileName := ifFileExistsRename(frxPDFexportProtocol.FileName);
      //Формируем протокол в Output
      frxReportTypeProtocol.Export(frxPDFexportProtocol);

      if (protocolVerifyStatusResult = CONFIRMED) and (InvoiceType = REGULAR_INVOICE) then
        begin
          //Формируем протокол в Invoice
          if System.SysUtils.DirectoryExists(DirectoryExportToInvoice) = False then
            System.SysUtils.ForceDirectories(DirectoryExportToInvoice);
          frxPDFexportProtocol.FileName := directoryExportToInvoice + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
          frxReportTypeProtocol.Export(frxPDFexportProtocol);
        end;
      if (protocolVerifyStatusResult = CONFIRMED) and (InvoiceType = MTR_INVOICE) then
        begin
          //Формируем протокол в InvoiceMTR
          if System.SysUtils.DirectoryExists(DirectoryExportToInvoiceMTR) = False then
            System.SysUtils.ForceDirectories(DirectoryExportToInvoiceMTR);
          frxPDFexportProtocol.FileName := directoryExportToInvoiceMTR + ProtocolName + Copy(ExtractFileName(SignatureFiles[i].Name), 1, Length(ExtractFileName(SignatureFiles[i].Name))-4) + '.pdf';
          frxReportTypeProtocol.Export(frxPDFexportProtocol);
        end;
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
          1 : arrayResultsD[i] := 'Успех';
          3 : arrayResultsD[i] := 'Подпись некорректна или к ней нет доверия';
        else arrayResultsD[i] := 'Статус не определён';
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
            descriptionError := 'Неверное имя файла "' + SearchResult.Name + '"';
            AddLog(DateToStr(Now) + ' ' + TimeToStr(Now) + '  ' + descriptionError + #13#10, isError);

            CreateResponceFileToOutput(SearchResult.Name, descriptionError + #13#10
                                                          + 'Верные имена файлов должны соответствовать маскам(фигурные скобки убираются):' + #13#10
                                                          + 'SH_{Код МО}_{Код СМО}_{основной/доплата}.zip' + #13#10
                                                          + 'SHO_{Код МО}_{Код СМО}_{основной/доплата}.zip' + #13#10
                                                          + 'SMP_{Код МО}_{Код СМО}_{основной/доплата}.zip' + #13#10
                                                          + 'SHCP_{Код МО}_{Код СМО}_основной.zip' + #13#10
                                                          + 'MSHO_{Код МО}_MTR_{основной/доплата}.zip' + #13#10
                                                          + 'MSH_{Код МО}_MTR_{основной/доплата}.zip' + #13#10
                                                          + 'MSMP_{Код МО}_MTR_{основной/доплата}.zip');
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
var DirectoryFrom, DirectoryToProcessed, DirectoryToInvoice, DirectoryToInvoiceMTR: string;
    fileDirectoryFrom, fileDirectoryToProcessed, fileDirectoryToInvoice, fileDirectoryToInvoiceMTR: string;
    pointerFileDirectoryFrom, pointerFileDirectoryToProcessed, pointerFileDirectoryToInvoice, pointerFileDirectoryToInvoiceMTR: PWideChar;
    MO, SMO: string;
    Year: integer;
    Month: string;
    i: integer;
begin
  Year := YearOf(Date);
  case MonthOf(Date) of
    1 : Month := 'Январь';
    2 : Month := 'Февраль';
    3 : Month := 'Март';
    4 : Month := 'Апрель';
    5 : Month := 'Май';
    6 : Month := 'Июнь';
    7 : Month := 'Июль';
    8 : Month := 'Август';
    9 : Month := 'Сентябрь';
    10 : Month := 'Октябрь';
    11 : Month := 'Ноябрь';
    12 : Month := 'Декабрь';
  else Month := 'Неизвестный месяц оО';
  end;
  MO := Copy(inputArchiveFileName, AnsiPos('_', inputArchiveFileName) + 1, 6);

  DirectoryFrom := DirectoryRoot;

  DirectoryToProcessed := DirectoryProcessed + IntToStr(Year) + '\' + Month + '\' + MO + '\' +
                 StringReplace(inputArchiveFileName, ExtractFileExt(inputArchiveFileName), '', [rfIgnoreCase]) + '\';
  DirectoryToProcessed := ifFolderExistsRename(DirectoryToProcessed);
  if System.SysUtils.DirectoryExists(DirectoryToProcessed) = False then
    System.SysUtils.ForceDirectories(DirectoryToProcessed);

  if AnsiPos('MTR', UpperCase(inputArchiveFileName)) = 0 then
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

  //Создаём протокол
  CreateProtocol(inputNotSigFile, inputSigFilesArray, DirectoryFrom, DirectoryToProcessed, DirectoryToInvoice, DirectoryToInvoiceMTR, inputArchiveFileName);

  //Копируем и переносим файлы в папку Processed:
  //– переносим оригинальный zip-файл
  fileDirectoryFrom := DirectoryFrom + inputArchiveFileName;
  pointerFileDirectoryFrom := Addr(fileDirectoryFrom[1]);
  fileDirectoryToProcessed := DirectoryToProcessed + inputArchiveFileName;
  pointerFileDirectoryToProcessed := Addr(fileDirectoryToProcessed[1]);
  MoveFile(pointerFileDirectoryFrom, pointerFileDirectoryToProcessed);

  //– копируем файл-счёт в папку с счетами / МТР-счетами и переносим его в папку Processed
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

  //– копируем sig-файлы в папку с счетами / МТР-счетами и переносим их в папку Processed
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

end;

function TFormMain.CheckFileName(inputFileName: string): boolean;
begin
  Result := False;

  if MatchesMask(inputFileName, 'SH_*_*_*.zip') or
     MatchesMask(inputFileName, 'SHO_*_*_*.zip') or
     MatchesMask(inputFileName, 'SMP_*_*_*.zip') or
     MatchesMask(inputFileName, 'SHCP_*_*_основной.zip') or
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
    //Проверка на количество файлов, прилагаемых в zip-архиве для подписания. По регламенту в архиве должен быть 1 файл-счёт для подписания.
    for i := 0 to Archive.Count-1 do
      begin
        if LowerCase(ExtractFileExt(Archive.item[i].FileName)) <> '.sig' then
          Counter := Counter + 1;
      end;
    if Counter > 1 then
      begin
        Result := True;
        descriptionErrorArchive := 'В zip-архиве "' + inputArchiveFileName + '" более одного файла-счёта для подписания.';
      end;
    if Counter = 0 then
      begin
        Result := True;
        descriptionErrorArchive := 'В zip-архиве "' + inputArchiveFileName + '" отсутствует файл-счёт для подписания.';
      end;

    //Проверка на количество подписей. Если в zip-архиве подписи отсутствуют, то в мусор.
    Counter := 0;
    for i := 0 to Archive.Count-1 do
      begin
        if LowerCase(ExtractFileExt(Archive.item[i].FileName)) = '.sig' then
          counter := Counter + 1;
      end;
    if Counter = 0 then
      begin
        Result := True;
        descriptionErrorArchive := 'В zip-архиве "' + inputArchiveFileName + '" отсутствуют файлы-подписи с расширением ".sig"';
      end;

    //Проверка на правильность имён файлов внутри zip-архива
    for i := 0 to Archive.Count-1 do
      begin
        if ( LowerCase(ExtractFileExt(Archive.Item[i].FileName)) <> '.sig' ) and
           ( Not MatchesMask(Archive.item[i].FileName, StringReplace(inputArchiveFileName, ExtractFileExt(inputArchiveFileName), '', [rfIgnoreCase]) + '*') ) then
          begin
            Result := True;
            descriptionErrorArchive := 'Файл-счёт "' + Archive.Item[i].FileName + '" внутри zip-архива "' + inputArchiveFileName + '" не соответствует его названию';
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

procedure TFormMain.UpdateDirectories(inputDirectoryRoot: string);
begin
  DirectoryErrors := DirectoryRoot + 'Errors';
  DirectoryErrors := CorrectPath(DirectoryErrors);

  DirectoryProcessed := DirectoryRoot + 'Processed';
  DirectoryProcessed := CorrectPath(DirectoryProcessed);
end;

procedure TFormMain.ButtonPathClick(Sender: TObject);
begin
  if SelectDirectory('Выберите папку для работы Автопроцессинга:', '', DirectoryRoot, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditPath.Text := DirectoryRoot;
end;

procedure TFormMain.ButtonInvoicePathClick(Sender: TObject);
begin
  if SelectDirectory('Выберите папку для выгрузки счетов:', '', DirectoryInvoice, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditInvoicePath.Text := DirectoryInvoice;
end;

procedure TFormMain.ButtonInvoiceMTRpathClick(Sender: TObject);
begin
  if SelectDirectory('Выберите папку для выгрузки счетов-МТР:', '', DirectoryInvoiceMTR, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditInvoiceMTRpath.Text := DirectoryInvoiceMTR;
end;

procedure TFormMain.ButtonOutputClick(Sender: TObject);
begin
  if SelectDirectory('Выберите папку для отправки протоколов и файлов с ошибками:', '', DirectoryOutput, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditOutput.Text := DirectoryOutput;
end;

procedure TFormMain.SpeedButtonPlayClick(Sender: TObject);
begin
  SpeedButtonPlay.Visible := False;
  SpeedButtonStop.Visible := True;

  if (SpinEditSec.Value > 59) or
     (SpinEditSec.Value < 0) or
     (SpinEditMin.Value < 0) or
     ( (SpinEditSec.Value = 0) and (SpinEditMin.Value = 0) ) then
    ShowMessage('Неверно указаны Минуты/Секунды')
  else
    begin
      TimerAutoProcessing.Interval := SpinEditMin.Value * 60000 + SpinEditSec.Value * 1000;
      TimerAutoProcessing.Enabled := True;

      TimerAutoProcessingState.Enabled := True;
      LabelAutoProcessingState.Caption := 'Автопроцессинг запущен';
    end;
end;

procedure TFormMain.SpeedButtonStopClick(Sender: TObject);
begin
  SpeedButtonStop.Visible := False;
  SpeedButtonPlay.Visible := True;

  TimerAutoProcessingState.Enabled := False;
  LabelAutoProcessingState.Caption := 'Автопроцессинг не запущен';

  TimerAutoProcessing.Enabled := False;
end;

procedure TFormMain.TimerAutoProcessingStateTimer(Sender: TObject);
begin
  LabelAutoProcessingState.Caption := LabelAutoProcessingState.Caption + '.';
  if LabelAutoProcessingState.Caption = 'Автопроцессинг запущен....' then
    LabelAutoProcessingState.Caption := 'Автопроцессинг запущен';
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

procedure TFormMain.AddLog(inputString: string; LogType: integer); //LogType бывает:
                                                                   //isError – цвет текста Красный
                                                                   //isSuccess – цвет текста Зелёный
                                                                   //isInformation – цвет текста Чёрный
begin
  case LogType of
    isError: begin
               RichEditLog.SelAttributes.Color := clRed;
               RichEditLog.Lines.Add(inputString);
               RichEditLog.Refresh;
             end;
    isSuccess: begin
                 RichEditLog.SelAttributes.Color := clGreen;
                 RichEditLog.Lines.Add(inputString);
                 RichEditLog.Refresh;
               end;
    isInformation: begin
                     RichEditLog.SelAttributes.Color := clBlack;
                     RichEditLog.Lines.Add(inputString);
                     RichEditLog.Refresh;
                   end;
  end;

end;

end.
