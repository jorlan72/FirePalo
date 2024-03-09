program FirePalo;

uses
  Vcl.Forms,
  main in 'main.pas' {frmMain},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Windows10 BlackPearl');
  Application.Title := 'FirePalo';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
