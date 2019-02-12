unit convert;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls;

type
  TForm3 = class(TForm)
    lbl1: TLabel;
    lbl2: TLabel;
    lbl3: TLabel;
    btn1: TButton;
    lbl4: TLabel;
    lbl5: TLabel;
    lbl6: TLabel;
    lbl7: TLabel;
    lbl8: TLabel;
    lbl9: TLabel;
    lbl10: TLabel;
    lbl11: TLabel;
    lbl12: TLabel;
    lbl13: TLabel;
    edt1: TEdit;
    edt2: TEdit;
    edt3: TEdit;
    edt4: TEdit;
    edt5: TEdit;
    tmr1: TTimer;
    procedure FormActivate(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure edt1KeyPress(Sender: TObject; var Key: Char);
    procedure FormCreate(Sender: TObject);
    procedure edt1Change(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;
  s: string;
  usd, eur, rub, xdr: currency;

implementation

{$R *.dfm}

uses sts, Clipbrd;

function HandlePasteClipboardContentsToTextEdit(
// Дескриптор окна текстового контроля.
wnd: HWND;
// Сообщение операционной системы.
uMsg: UINT;
// Старший параметр.
wParam: WPARAM;
// Младший параметр.
lParam: LPARAM): Integer; stdcall; forward;

function HandlePasteClipboardContentsToTextEdit(wnd: HWND; uMsg: UINT;
wParam: WPARAM; lParam: LPARAM): Integer; stdcall;
var
ClipboardText: String;
I: Integer;
begin
// Обработка вставки текста из буфера обмена.
if (uMsg = WM_PASTE) and Clipboard.HasFormat(CF_TEXT) then
begin
ClipboardText := Clipboard.AsText;
for I := 1 to Length(ClipboardText) do
// Если в тексте буфера обмена есть символы отличные от цифр, отменяем вставку.
if not (ClipboardText[I] in ['0'..'9']) then
begin
uMsg := 0;
Break;
end;
end;
{ $WARN UNSAFE_TYPE OFF }
Result := CallWindowProc(Pointer(GetWindowLong(wnd, GWL_USERDATA)),
wnd, uMsg, wParam, lParam);
{ $WARN UNSAFE_TYPE ON }
end;

//точку преобразовать в запятую
function StrToExt(const AValue: String): String;
const
  FError = -0.000013;
var
  sTemp: string;
  i: Integer;
begin
  Result := AValue;
  if Result = AValue then
  begin
    sTemp := AValue;
    for i := 1 to Length(sTemp) do
      if sTemp[i] = '.' then
      begin
        sTemp[i] := ',';
        Break;
      end;
    Result := sTemp;
  end;
end;

procedure TForm3.FormActivate(Sender: TObject);
var
  k: Extended;
  s,s1,s2,s3: string;
begin
  //США
  lbl1.Caption:= Form1.usd1;
  k:=StrToFloat(StrToExt(Form1.usd1));
  k:=k/100;
  s:=FloatToStr(k);
  edt2.Text:=s;
  //EURO
  lbl2.Caption:= Form1.eur1;
  k:=StrToFloat(StrToExt(Form1.eur1));
  k:=k/100;
  s1:=FloatToStr(k);
  edt3.Text:=s1;
  //RUB
  lbl3.Caption:= Form1.rub1;
  k:=StrToFloat(StrToExt(Form1.rub1));
  k:=k/10;
  s2:=FloatToStr(k);
  edt4.Text:=s2;
  //XDR
  lbl4.Caption:= Form1.xdr1;
  k:=StrToFloat(StrToExt(Form1.xdr1));
  k:=k/100;
  s3:=FloatToStr(k);
  edt5.Text:=s3;
  //////////////////
  usd:=strtofloat(s);
  eur:=strtofloat(s1);
  rub:=strtofloat(s2);
  xdr:=strtofloat(s3);
  /////////////////
  s:=edt1.Text;
  lbl12.Caption:=lbl12.Caption+s;
  lbl13.Caption:=lbl13.Caption+s;
  lbl8.Caption:=lbl8.Caption+s;
  lbl9.Caption:=lbl9.Caption+s;
end;

procedure TForm3.btn1Click(Sender: TObject);
begin
  if edt1.text<>'' then
    begin
      edt2.Text:=floattostr(strtofloat(edt1.Text)*usd)+' грн.';
      edt3.Text:=floattostr(strtofloat(edt1.Text)*eur)+' грн.';
      edt4.Text:=floattostr(strtofloat(edt1.Text)*rub)+' грн.';
      edt5.Text:=floattostr(strtofloat(edt1.Text)*xdr)+' грн.';
    end
  else begin
    edt2.Text:='';
    edt3.Text:='';
    edt4.Text:='';
    edt5.Text:='';
    edt1.Clear;
    edt1.Text:='1';
  end;
end;

procedure TForm3.edt1KeyPress(Sender: TObject; var Key: Char);
begin
 if Not (Key in ['0'..'9', #8])then Key:=#0;
end;

procedure TForm3.FormCreate(Sender: TObject);
begin
SetWindowLong(edt1.Handle, GWL_STYLE,
GetWindowLong(edt1.Handle, GWL_STYLE) or ES_NUMBER);
{ $WARN UNSAFE_CODE OFF }
SetWindowLong(edt1.Handle, GWL_USERDATA,
SetWindowLong(edt1.Handle, GWL_WNDPROC,
LPARAM(@HandlePasteClipboardContentsToTextEdit)));
{ $WARN UNSAFE_CODE ON }
end;

procedure TForm3.edt1Change(Sender: TObject);
begin
  if edt1.Text <> '' then s:=edt1.Text
  else Exit;
  lbl12.Caption:=lbl12.Caption+s;
  lbl13.Caption:=lbl13.Caption+s;
  lbl8.Caption:=lbl8.Caption+s;
  lbl9.Caption:=lbl9.Caption+s;
end;

procedure TForm3.tmr1Timer(Sender: TObject);
begin
  if s <> '' then begin
     lbl12.Caption:='USD/'+s;
     lbl13.Caption:='EUR/'+s;
     lbl8.Caption:='RUB/'+s;
     lbl9.Caption:='XDR/'+s;
     s:='';
  end;
end;

end.
