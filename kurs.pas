unit kurs;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IdBaseComponent, IdComponent, IdRawBase, IdRawClient,
  IdIcmpClient,
  HTTPSend,
  XPMan,
  ClipBrd,
  Winsock,
  WinInet, UrlTools,
  StdCtrls, IdTCPConnection, IdTCPClient, IdHTTP;

type
  TKS1 = class(TForm)
    IdIcmpClient1: TIdIcmpClient;
    grp2: TGroupBox;
    mmo1: TMemo;

    IdHTTP1: TIdHTTP;    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  KS1: TKS1;

implementation

{$R *.dfm}

uses sts;

//Защита от отладчика
function DebuggerPresent:boolean;
type
  TDebugProc = function:boolean; stdcall;
var
   Kernel32:HMODULE;
   DebugProc:TDebugProc;
begin
   Result:=false;
   Kernel32:=GetModuleHandle('kernel32.dll');
   if kernel32 <> 0 then
    begin
      @DebugProc:=GetProcAddress(kernel32, 'IsDebuggerPresent');
      if Assigned(DebugProc) then
         Result:=DebugProc;
    end;                                  
end;

procedure log(s: string);
begin
  KS1.mmo1.Lines.Add('['+DateToStr(Now)+' | '+TimeToStr(Now)+'] '+s);
end;

function OpenInternet(Name: WideString): pointer;
begin
  result := InternetOpenW(@Name[1], INTERNET_OPEN_TYPE_PRECONFIG,
  nil, nil, 0);
end;

function ProxyHttpPostURL(const URL, URLData: string; const Data: TStream): Boolean;
var
  int: int64;
  http:TIdHTTP;
  str:TFileStream;
  UrlSite  : string;
  SrcPathCount: Integer;
  SrcHost, SrcPath, Srcfname, Srcfext:String;
begin
  if DebuggerPresent then Application.Terminate;
  OpenInternet('Mozilla Firefox');
  //Создим класс для закачки
  http:=TIdHTTP.Create(nil);
  http.ProtocolVersion:=pv1_1;
  http.HandleRedirects:=true;
  if Form1.chk8.Checked then begin
  if Form1.proxyIP = '' then Exit;
  if Form1.proxyPort = '' then Exit;
    //Настройка прокси
    http.ProxyParams.BasicAuthentication := true;
    http.ProxyParams.ProxyServer:=Form1.proxyIP;   //10.220.1.7
    http.ProxyParams.ProxyPort:=StrToInt(Form1.proxyPort); //3129
  end;
  //каталог, куда файл положить
  ForceDirectories(ExtractFileDir(ExtractFilePath(ParamStr(0))));
  //Поток для сохранения
  str:=TFileStream.Create(ExtractFilePath(ParamStr(0))+'nbu.txt', fmCreate);
  SrcPathCount:=SplitFullURL(Trim(URL), Srchost,Srcpath,Srcfname,Srcfext);
  if (SrcPathCount=-1) then Exit;
     UrlSite:='http://'+SrcHost+SrcPath+Srcfname+Srcfext;
  try
     Application.ProcessMessages;
     http.Get(URL,str);
     Data.CopyFrom(str,int);
  except on e: Exception do
     ShowMessage(e.Message); // вот здесь получаем 'Operation aborted'
  end;
  http.Free;
  str.Free;
  DeleteFile('nbu.txt');
end;

function ParseStr(str, sub1, sub2: string): string;
var
 st, fin: Integer;
 n: string;
begin
 st := Pos(sub1, str);
if st = 0 then n:='';
if st > 0 then begin
 str := Copy(str, st + length(sub1), length(str) - 1);
 st := 1;
 fin := Pos(sub2, str);
 Result := Copy(str, st, fin - st);
 str := Copy(str, fin + length(sub2), length(str) - 1);
end;
end;

function StrToZap(const AValue: String): String;
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

procedure TKS1.FormActivate(Sender: TObject);
var
  i,y: integer;
  s,s1: string;
  statnbu,statp: Boolean;
  tx: TStringList;
  st: TMemoryStream;
  k1,k2,k3,k4,k5,k6: string;
begin
    mmo1.Lines.Clear;
    tx:=TStringList.Create;
    st:=TMemoryStream.Create;
    if DebuggerPresent then Application.Terminate;
    if Form1.HostToIP('finance.liga.net',s) then
    if s <> '' then s:=s;
       log('Подключение к '+s);
       s:=Form1.IPAddrToName(s);
    if s <> 'no' then begin
       log('Определение хоста '+s);
       log('Проверка подключения ...');
       IdIcmpClient1.Host:='10.202.14.15';
       IdIcmpClient1.ReceiveTimeout:=1000;
       IdIcmpClient1.Ping('32');
       log('Connect: --> '+IntToStr(IdIcmpClient1.ReplyStatus.MsRoundTripTime)+' мс.');
       Application.ProcessMessages;
    if IdIcmpClient1.ReplyStatus.MsRoundTripTime > 0 then begin
    log('Подключение к 10.202.14.15');
    ProxyHTTPpostURL('http://10.202.14.15/valuta/val.html','<table class="ui-tabs-panel ui-widget-content ui-corner-bottom" id="link_nal">', st);
    st.Seek(0,soFromBeginning);
    tx.LoadFromStream(st);
    //tx.LoadFromFile('nbu.txt');   //Для теста
    //tx.SaveToFile('nbu.txt'); //Для теста
    s:= Utf8ToAnsi(tx.Text);
    for i:=0 to tx.Count-1 do begin
        s:= tx.Strings[i];
        s1:=Copy(s,0,Pos('=',s)-1);
        s1:=Trim(s1);
        if s = '[NBU]' then begin
           log('Покупка: ');
           statnbu:=True;
           statp:=False;
        end;
        if s = '[SELL]' then begin
           log('Продажа: ');
           statnbu:=False;
           statp:=True;
        end;
        if s = '[DATE]' then begin
           s:= tx.Strings[i+1];
           s1:=Copy(s,Pos('=',s)+1,Length(tx.Strings[i+1])-Pos('=',s));
           s1:=Trim(s1);
           log('Дата: '+s1);
        end;
        if s <> '[NBU]' then begin //Покупка
        if s <> '[SELL]' then
        if s <> '[DATE]' then
        if s <> '' then
        if not statp then
        if statnbu then
        if s1 = 'USD' then begin
           log(s);
           k4:=s;
           y:=Pos('=',s);
           if y > 0 then
              Form1.nbuUSA:=Trim(Copy(s,y+1,Length(s)));
        end;
        if not statp then
        if statnbu then
        if s1 = 'EUR' then begin
           log(s);
           k1:=s;
           y:=Pos('=',s);
           if y > 0 then
              Form1.nbuEuro:=Trim(Copy(s,y+1,Length(s)));
        end;
        if not statp then
        if statnbu then
        if s1 = 'SPZ' then begin
           log(s);
           k3:=s;
           y:=Pos('=',s);
           if y > 0 then
              Form1.nbuSPZ:=Trim(Copy(s,y+1,Length(s)));
        end;
        end;
        if s <> '[SELL]' then begin  //Продажа
        if s <> '[NBU]' then
        if s <> '[DATE]' then
        if s <> '' then
        if not statnbu then
        if statp then
        if s1 = 'USD' then begin
           log(s);
           statp:=True;
           k5:=s;
           y:=Pos('=',s);
           if y > 0 then
              Form1.USAP:=Trim(Copy(s,y+1,Length(s)));
        end;
        if not statnbu then
        if statp then
        if s1 = 'EUR' then begin
           log(s);
           statp:=True;
           k2:=s;
           y:=Pos('=',s);
           if y > 0 then
              Form1.EUROP:=Trim(Copy(s,y+1,Length(s)));
        end;
        if not statnbu then
        if statp then
        if s1 = 'SPZ' then begin
           log(s);
           statp:=True;
           k6:=s;
        end;
        end;
    end;
    end;
    end else log('Сервер не найден!');
    tx.Free;
    st.Free;
    mmo1.Lines.SaveToFile('kurs.txt');
    KS1.Close;
end;

end.
