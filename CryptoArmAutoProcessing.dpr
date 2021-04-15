program CryptoArmAutoProcessing;

uses
  Vcl.Forms,
  uCryptoArmAutoprocessing in 'uCryptoArmAutoprocessing.pas' {FormMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
