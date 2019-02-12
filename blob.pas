unit blob;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  Tbb = class(TForm)
    mmo1: TMemo;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  bb: Tbb;

implementation

{$R *.dfm}

procedure Tbb.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FileExists(ExtractFilePath(ParamStr(0))+'BlobSave\x.blob') then
   bb.mmo1.Lines.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\x.blob');
end;

end.
