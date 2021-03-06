program HeatsFormer;

uses
  Forms,
  uMain in 'Units\uMain\uMain.pas' {frmMain},
  uProgram in 'Units\uProgram\uProgram.pas',
  uSwimmer in 'Units\uSwimmer\uSwimmer.pas',
  uExcel in 'Units\uExcel\uExcel.pas',
  uTechnical in 'Units\uTechnical\uTechnical.pas',
  uError in 'Units\uError\uError.pas',
  uApplicat in 'Units\uApplicat\uApplicat.pas',
  uFinish in 'Units\uFinish\uFinish.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
