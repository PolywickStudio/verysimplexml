program Project1;

uses
  Vcl.Forms,
  Unit1 in 'Unit1.pas' {Form1},
  KwikQat in 'KwikQat.pas',
  Xml.VerySimple in 'Xml.VerySimple.pas',
  Xml.VerySimpleFIXED in 'Xml.VerySimpleFIXED.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
