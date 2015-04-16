program BandReport;

uses
  Forms,
  MainGUI in 'MainGUI.pas' {FMainGUI},
  Acatera.Report in 'Acatera.Report.pas',
  XLSfile in 'XLSfile.pas',
  Acatera.Utils in 'Acatera.Utils.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFMainGUI, FMainGUI);
  Application.Run;
end.
