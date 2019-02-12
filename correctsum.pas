unit correctsum;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TCorSUM1 = class(TForm)
    stat1: TStatusBar;
    grp1: TGroupBox;
    corsum: TEdit;
    minus: TCheckBox;
    plus: TCheckBox;
    ext: TCheckBox;
    procedure minusClick(Sender: TObject);
    procedure plusClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure corsumChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure extClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    sum: string;
  end;

var
  CorSUM1: TCorSUM1;
  extf: Boolean;

implementation

uses sts;

{$R *.dfm}

procedure TCorSUM1.minusClick(Sender: TObject);
var
  i,y,z: Double;
begin
  if minus.Checked then begin
     stat1.Panels[0].Text:='Введеная сумма будет отнята от текущей!';
     plus.Checked:=False;
  end;
  if CorSUM1.minus.Checked then begin
     z:=0;
     i:=StrToFloat(Form1.tmps);
     y:=Pos('.',CorSUM1.corsum.Text);
     if y >= 0 then
        y:=StrToFloat(Form1.StrToZap(CorSUM1.corsum.Text))
     else
        y:=StrToFloat(CorSUM1.corsum.Text);
     if (i > 0) and (y >= 0) then z:=i-y;
     if (i <= 0) and (y >= 0) then z:=i-y;
        sum:=FloatToStr(z);
        grp1.Caption:='Текущая сумма: '+Form1.tmps+' - '+corsum.Text+' =  '+FloatToStr(z);
  end;
end;

procedure TCorSUM1.plusClick(Sender: TObject);
var
  i,y,z: Double;
begin
  if plus.Checked then begin
     stat1.Panels[0].Text:='Введеная сумма будет прибавлена к текущей!';
     minus.Checked:=False;
  end;
  if CorSUM1.plus.Checked then begin
     z:=0;
     i:=StrToFloat(Form1.tmps);
     y:=Pos('.',CorSUM1.corsum.Text);
     if y >= 0 then
        y:=StrToFloat(Form1.StrToZap(CorSUM1.corsum.Text))
     else
        y:=StrToFloat(CorSUM1.corsum.Text);
     if (i > 0) and (y >= 0) then z:=i+y;
     if (i <= 0) and (y >= 0) then z:=y+i;
        sum:=FloatToStr(z);
        grp1.Caption:='Текущая сумма: '+Form1.sum+' + '+corsum.Text+' =  '+FloatToStr(z);
  end;
end;

procedure TCorSUM1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
  y,z: Double;
begin
     CanClose:=extf;
  if (not minus.Checked) and (not plus.Checked) and (corsum.Text = '') then begin
  if not extf then
     MessageBox(Handle,PChar('Не выбрана операция коррекции суммы!'),PChar('Внимание'),64);
     form1.Edit4.Enabled:=True;
     form1.Edit4.Text:=corsum.Text;
     CanClose:=extf;
     Exit;
  end else begin
  if (not minus.Checked) and (not plus.Checked) then begin
     form1.Edit4.Enabled:=False;
     form1.Edit4.Text:=corsum.Text;
     y:=Pos('.',CorSUM1.corsum.Text);
     if y >= 0 then
        z:=StrToFloat(Form1.StrToZap(CorSUM1.corsum.Text))
     else
        z:=StrToFloat(CorSUM1.corsum.Text);
     sum:=FloatToStr(z);
     CanClose:=extf;
     CorSUM1.Close;
  end else begin
     form1.Edit4.Enabled:=False;
     form1.Edit4.Text:=sum;
     CanClose:=extf;
     CorSUM1.Close;
  end;
  end;
end;

procedure TCorSUM1.corsumChange(Sender: TObject);
var
  i,y,z: Double;
begin
       z:=0;
       i:=StrToFloat(Form1.StrToZap(Form1.tmps));
       y:=Pos('.',CorSUM1.corsum.Text);
       if y >= 0 then
       y:=StrToFloat(Form1.StrToZap(CorSUM1.corsum.Text))
       else
       y:=StrToFloat(CorSUM1.corsum.Text);
       if (i > 0) and (y >= 0) then z:=i-y;
       if (plus.Checked) and (i > 0) and (y >= 0) then z:=i+y;
       if (plus.Checked) and (i < 0) and (y >= 0) then z:=y+i;
       if (minus.Checked) and (i > 0) and (y >= 0) then z:=i-y;
       if (minus.Checked) and (i < 0) and (y >= 0) then z:=i-y;
       sum:=FloatToStr(z);
       grp1.Caption:='Текущая сумма: '+Form1.tmps+' - '+corsum.Text+' =  '+FloatToStr(z);
end;

procedure TCorSUM1.FormCreate(Sender: TObject);
begin
  ext.Checked:=True;
  grp1.Caption:='Текущая сумма: '+Form1.sum;
end;

procedure TCorSUM1.extClick(Sender: TObject);
begin
  if ext.Checked then extf:=True
  else extf:=False;
end;

procedure TCorSUM1.FormActivate(Sender: TObject);
begin
  minus.Checked:=False;
  plus.Checked:=False;
end;

end.
