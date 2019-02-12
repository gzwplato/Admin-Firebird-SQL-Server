program admin;

uses
  Forms,
  sts in 'sts.pas' {Form1},
  connops in 'connops.pas' {Form2},
  correctsum in 'correctsum.pas' {CorSUM1},
  blob in 'blob.pas' {bb},
  kurs in 'kurs.pas' {KS1};

{$R *.res}

begin
  Application.Initialize;
  Application.Title:='StalkerSTS';
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TCorSUM1, CorSUM1);
  Application.CreateForm(Tbb, bb);
  Application.CreateForm(TKS1, KS1);
  Form2.ShowModal;
  Application.Run;
end.
