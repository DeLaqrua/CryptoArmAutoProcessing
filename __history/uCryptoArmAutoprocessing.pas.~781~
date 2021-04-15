unit uCryptoArmAutoprocessing;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.OleCtrls, MSScriptControl_TLB,
  Vcl.StdCtrls, ActiveX, frxClass, System.Zip, Vcl.FileCtrl, System.Masks;

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
    procedure FormCreate(Sender: TObject);
    procedure ButtonManualProcessingClick(Sender: TObject);
    procedure ButtonPathClick(Sender: TObject);

  private
    { Private declarations }
  public
    function SignatureVerify(InputFileName, InputFileNameSignature: string): string;
    function CheckErrorsWithinArchive(InputArchiveFileName: string): boolean;
    function CorrectPath(InputDirectory: string): string;
    function CheckFileName(InputFileName: string): boolean;

    procedure UpdateDirectories(InputDirectoryRoot: string);
    procedure SortErrorFiles;
    procedure MoveFilesToErrors(InputFileName: string);

    procedure Processed(InputArchiveFileName: string);
  end;

var
  FormMain: TFormMain;
var
  DirectoryRoot, DirectoryErrors, DirectoryProcessed, DirectoryOutput, DescriptionErrorArchive: string;

implementation

{$R *.dfm}
//Hello GitHub
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
    responceTextFile: TextFile;
    responceTextFileName: string;
    i: integer;
begin

  DirectoryRoot := CorrectPath(EditPath.Text);
  if System.SysUtils.DirectoryExists(DirectoryRoot) = False then
    ShowMessage('Проверьте путь к директории. Папки не существует.')
  else
    begin

      UpdateDirectories(DirectoryRoot);

      SortErrorFiles;

      if FindFirst(DirectoryRoot + '*.*', faNormal, SearchResult) = 0 then
        begin
          repeat
            if CheckFileName(SearchResult.Name) and CheckErrorsWithinArchive(SearchResult.Name)then
              begin
                MoveFilesToErrors(SearchResult.Name);

                if System.SysUtils.DirectoryExists(DirectoryOutput) = False then
                  System.SysUtils.ForceDirectories(DirectoryOutput);
                responceTextFileName := DirectoryOutput + 'response_' + Copy(SearchResult.Name, 1, Length(SearchResult.Name)-4) + '.txt';
                i := 0;
                while FileExists(responceTextFileName) do
                  begin
                    i := i+1;
                    if i = 1 then
                      Insert(' (' + IntToStr(i) + ')', responceTextFileName, Length(responceTextFileName)-3)
                    else
                      responceTextFileName := StringReplace(responceTextFileName, ' (' + IntToStr(i-1) + ')', ' (' + IntToStr(i) + ')', []);
                  end;
                AssignFile(responceTextFile, responceTextFileName);
                ReWrite(responceTextFile);
                WriteLn(responceTextFile, DescriptionErrorArchive);
                CloseFile(responceTextFile);
              end
            else Processed(SearchResult.Name);
          until FindNext(SearchResult) <> 0;
          FindClose(SearchResult);
        end;

    end;
end;

procedure TFormMain.Processed(InputArchiveFileName: string);
var i, j: integer;
    Archive: TZipFile;
    SigFilesArray: array of string;
    NotSigFile: string;
begin
  Archive := TZipFile.Create;
  try
    Archive.Open(DirectoryRoot + InputArchiveFileName, zmRead);
    Archive.ExtractAll(DirectoryRoot);

    for i := 0 to Archive.FileCount-1 do
      begin
        if LowerCase(Copy(Archive.FileName[i], length(Archive.FileName[i])-3, 4)) = '.sig' then
          begin
            SigFilesArray[i] := Archive.FileName[i];
          end
        else
          begin
            NotSigFile := Archive.FileName[i];
          end;
      end;

    Archive.Close;
  finally
    Archive.Free;
  end;

  for i := 0 to High(SigFilesArray) do
    begin
      MemoLog.Lines.Add( SignatureVerify(DirectoryRoot + NotSigFile, DirectoryRoot + SigFilesArray[i]) );
    end;

end;

function TFormMain.SignatureVerify (InputFileName, InputFileNameSignature: string): string;
var VArr, ResultFromVB: Variant;
    ResultDescription: string;
    FunctionParameters: PSafeArray;
begin
  try
    VArr:=VarArrayCreate([0, 1], varVariant);
    VArr[0] := InputFileName;
    VArr[1] := InputFileNameSignature;

    FunctionParameters := PSafeArray(TVarData(VArr).VArray);

    ResultFromVB := ScriptControlVB.Run('SignatureVerify', FunctionParameters);
    case ResultFromVB of
      1 : ResultDescription := 'Успех';
      3 : ResultDescription := 'Подпись некорректна или к ней нет доверия';
    else ResultDescription := 'Статус не определён';
    end;

  except
    on E: Exception do
    MessageDlg(PWideChar(E.Message), mtError, [mbOk], 0);
  end;

  Result := ResultDescription;
end;

procedure TFormMain.SortErrorFiles;
var SearchResult: TSearchRec;
begin
  if System.SysUtils.DirectoryExists(DirectoryErrors) = False then
    System.SysUtils.ForceDirectories(DirectoryErrors);

  if FindFirst(DirectoryRoot + '*.*', faNormal, SearchResult) = 0 then
    begin
      repeat
        if (LowerCase(Copy(SearchResult.Name, length(SearchResult.Name)-3, 4)) <> '.zip') or
           (CheckFileName(SearchResult.Name) = false) then
          begin
            MoveFilesToErrors(SearchResult.Name);
          end;
      until FindNext(SearchResult) <> 0;
      FindClose(SearchResult);
    end;
end;

function TFormMain.CheckFileName(InputFileName: string): boolean;
begin
  Result := False;

  if MatchesMask(InputFileName, 'SH_*_*_*.zip') or
     MatchesMask(InputFileName, 'SHO_*_*_*.zip') or
     MatchesMask(InputFileName, 'MSHO_*_*_*.zip') or
     MatchesMask(InputFileName, 'MSH_*_*_*.zip') or
     MatchesMask(InputFileName, 'MSMP_*_*_*.zip') or
     MatchesMask(InputFileName, 'SMP_*_*_*.zip') then
    Result := True;

end;

function TFormMain.CheckErrorsWithinArchive(InputArchiveFileName: string): boolean;
var i, Counter: integer;
    Archive: TZipFile;
begin
  DescriptionErrorArchive := '';

  Result := False;

  Archive := TZipFile.Create;
  try
    Archive.Open(DirectoryRoot + InputArchiveFileName, zmRead);

    Counter := 0;
    //Проверка на количество файлов, прилагаемых в zip-архиве для подписания. По регламенту в архиве должен быть 1 файл для подписания.
    for i := 0 to Archive.FileCount-1 do
      begin
        if LowerCase(Copy(Archive.FileName[i], length(Archive.FileName[i])-3, 4)) <> '.sig' then
          Counter := Counter + 1;
      end;
    if Counter <> 1 then
      begin
        Result := True;
        DescriptionErrorArchive := 'В zip-архиве "' + InputArchiveFileName + '" более одного файла для подписания.';
      end;

    //Проверка на количество подписей. Если в zip-архиве подписи отсутствуют, то в мусор.
    Counter := 0;
    for i := 0 to Archive.FileCount-1 do
      begin
        if LowerCase(Copy(Archive.FileName[i], length(Archive.FileName[i])-3, 4)) = '.sig' then
          counter := Counter + 1;
      end;
    if Counter = 0 then
      begin
        Result := True;
        DescriptionErrorArchive := 'В zip-архиве "' + InputArchiveFileName + '" отсутствуют файлы-подписи с расширением ".sig"';
      end;

    Archive.Close;
  finally
    Archive.Free;
  end;

end;

procedure TFormMain.MoveFilesToErrors(InputFileName: string);
var DirectoryFrom, DirectoryTo: string;
    pointerDirectoryFrom, pointerDirectoryTo: PWideChar;
    i: integer;
begin
  if System.SysUtils.DirectoryExists(DirectoryErrors) = False then
    System.SysUtils.ForceDirectories(DirectoryErrors);

  DirectoryFrom := DirectoryRoot + InputFileName;
  DirectoryTo := DirectoryErrors + InputFileName;
  pointerDirectoryFrom := Addr(DirectoryFrom[1]);
  pointerDirectoryTo := Addr(DirectoryTo[1]);

  i := 0;
  while FileExists(DirectoryTo) do
    begin
      i := i+1;
      if i = 1 then
        begin
          Insert(' (' + IntToStr(i) + ')', DirectoryTo, Length(DirectoryTo)-3);
          pointerDirectoryTo := Addr(DirectoryTo[1]);
        end
      else
        begin
          DirectoryTo := StringReplace(DirectoryTo, ' (' + IntToStr(i-1) + ')', ' (' + IntToStr(i) + ')', []);
          pointerDirectoryTo := Addr(DirectoryTo[1]);
        end;
    end;

  MoveFile(pointerDirectoryFrom, pointerDirectoryTo);
end;

function TFormMain.CorrectPath(InputDirectory: string): string;
begin
  if Pos('/', InputDirectory) <> 0 then
    begin
      InputDirectory := StringReplace(InputDirectory, '/', '\', [rfReplaceAll]);
    end;

  if InputDirectory[length(InputDirectory)] <> '\' then
    Result := InputDirectory + '\'
  else
    Result := InputDirectory;
end;

procedure TFormMain.UpdateDirectories(InputDirectoryRoot: string);
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
  if SelectDirectory('Выберите папку для работы Автопроцессинга:', '', DirectoryRoot, [sdNewFolder, sdShowShares, sdNewUI, sdValidateDir]) then
    EditPath.Text := DirectoryRoot;
end;

end.
