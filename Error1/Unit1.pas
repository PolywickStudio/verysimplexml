﻿unit Unit1;
interface
uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs;

type
  TForm1 = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses
  KwikQat;
{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  RepairQAT('word.org');
  RepairQAT('word2.org');
  RepairQAT('word3.org');
end;

end.
