unit connops;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, IdBaseComponent, IdComponent, IdRawBase, IdRawClient,
  IdIcmpClient, ExtCtrls,
  XPMan,
  WinSock,
  ComCtrls;

type
  TForm2 = class(TForm)
    ListBox1: TListBox;
    IdIcmpClient1: TIdIcmpClient;
    tmr1: TTimer;
    stat1: TStatusBar;
    procedure IdIcmpClient1Reply(ASender: TComponent;
      const AReplyStatus: TReplyStatus);
    procedure tmr1Timer(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

type
  PNetResourceArray = ^TNetResourceArray;
  TNetResourceArray = array[0..100] of TNetResource;

var
  Form2: TForm2;
  Os,z: Integer;
  L: TStrings;
  s,s1,s2: string;
  ////////////////////
  hNetEnum: THandle;
  i,ResourceBuf,EntriesToGet: DWORD;
  ResourceBuffer: array[1..2000] of TNetResource;
  NetContainerToOpen: NETRESOURCE;

implementation

{$R *.dfm}

uses sts;

procedure CreateFormInRightBottomCorner;
var
 r : TRect;
begin
 SystemParametersInfo(SPI_GETWORKAREA, 0, Addr(r), 0);
 Form2.Left := r.Right-Form2.Width;
 Form2.Top := r.Bottom-Form2.Height;
end;

//Нахождение шары в сети
function CreateNetResourceList(ResourceType: DWord;
                              NetResource: PNetResource;
                              out Entries: DWord;
                              out List: PNetResourceArray): Boolean;
var
  EnumHandle: THandle;
  BufSize: DWord;
  Res: DWord;
begin
  Result := False;
  List := Nil;
  Entries := 0;
  if WNetOpenEnum(RESOURCE_GLOBALNET,
                  ResourceType,
                  0,
                  NetResource,
                  EnumHandle) = NO_ERROR then begin
    try
      BufSize := $4000;  // 16 kByte
      GetMem(List, BufSize);
      try
        repeat
          Entries := DWord(-1);
          FillChar(List^, BufSize, 0);
          Res := WNetEnumResource(EnumHandle, Entries, List, BufSize);
          if Res = ERROR_MORE_DATA then
           begin
            ReAllocMem(List, BufSize);
           end;
        until Res <> ERROR_MORE_DATA;
              Result := Res = NO_ERROR;
        if not Result then
        begin
          FreeMem(List);
          List := Nil;
          Entries := 0;
        end;
        except
          FreeMem(List);
        raise;
    end;
    finally
      WNetCloseEnum(EnumHandle);
    end;
  end;
end;

//IP по имени хоста
function GetIPFromHost(const HostName: string): string;
type
  TaPInAddr = array[0..10] of PInAddr;
  PaPInAddr = ^TaPInAddr;
var
  phe: PHostEnt;
  pptr: PaPInAddr;
  i: Integer;
  GInitData: TWSAData;
begin
  WSAStartup($101, GInitData);
  Result := '';
  phe := GetHostByName(PChar(HostName));
  if phe = nil then Exit;
  pPtr := PaPInAddr(phe^.h_addr_list);
  i := 0;
  while pPtr^[i] <> nil do
  begin
    Result := inet_ntoa(pptr^[i]^);
    Inc(i);
  end;
  WSACleanup;
end;

//Удаление ненужных символов
function DelSubstr(sub,s:string):string;
var
 i,t:integer;
begin
 t:=1;
 i:=pos(sub,s);
 while i>0 do
  begin
   t:=t+i-1;
   delete(s,t,length(sub));
   i:=pos(sub,copy(s,t,length(s)));
  end;
   DelSubstr:=s;
end;

//Сканируем сеть на открытые ресурсы
procedure ScanNetworkResources(ResourceType, DisplayType: DWord; List: TStrings; IPS: string);
procedure ScanLevel(NetResource: PNetResource);
var
  //Entries: DWord;
  Ts:TStringList;
  //NetResourceList: PNetResourceArray;
  //a,b,z,s,s1,
  s2,s3,s4,s5,
  s6,s7,s8,s9,s10,s11,harkov,net,datex,datey:string;
  //y: Integer;
begin
  L:= TStringList.Create;
  if IPS<>'' then begin
         NetContainerToOpen.dwScope:=RESOURCE_GLOBALNET;
         NetContainerToOpen.dwType:=RESOURCETYPE_ANY;
         NetContainerToOpen.lpLocalName:=nil;
         NetContainerToOpen.lpRemoteName:= PChar('\\'+IPS); //IPS
         NetContainerToOpen.lpProvider:= nil;
         WNetOpenEnum(RESOURCE_GLOBALNET, RESOURCETYPE_ANY, RESOURCEUSAGE_CONNECTABLE or RESOURCEUSAGE_CONTAINER,
         @NetContainerToOpen, hNetEnum);
       while TRUE do
        begin
         ResourceBuf := sizeof(ResourceBuffer);
         EntriesToGet := 2000;
        if (NO_ERROR <> WNetEnumResource(hNetEnum, EntriesToGet,
         @ResourceBuffer, ResourceBuf)) then
         begin
          WNetCloseEnum(hNetEnum);
          exit;
         end;
        for i := 1 to EntriesToGet do
         begin
         s3:=PChar(string(ResourceBuffer[i].lpRemoteName));
         harkov:='\\10.229.15.20\sts';
         if s3<>'' then
          begin
           Ts:=TStringList.Create();
           Ts.Text:=stringReplace(s3,'\',#13#10,[rfReplaceAll]);
           s:=TS[0];
           s1:=TS[1];
           s2:=TS[2];
           s1:=TS[3];
           Ts.Free;
           s:='\\'+s2+'\HPLaserJ';
           s1:='\\'+s2+'\Принтер';
           s4:='\\'+s2+'\SamsungM';
           s5:='\\'+s2+'\HP';
           s6:='\\'+s2+'\HPLaserJ.2';
           s7:='\\'+s2+'\Epson LX-300+ (Копия 1)';
           s8:='\\'+s2+'\Epson LX-300+';
           s9:='\\'+s2+'\Epson LX-';
           s10:='\\'+s2+'\Epson LX-300+ (Копия 2)';
           s11:='\\'+s2+'\Epson LX-300+ (Копия 3)';
           if s3 = harkov then
            begin
             if DirectoryExists('\\10.229.15.20\sts\load') then
              begin
                datex:=copy(DateToStr(Date),1,2)+copy(DateToStr(Date),4,2);
                datey:=copy(DateToStr(Date),1,2)+copy(DateToStr(Date),4,2)+copy(DateToStr(Date),7,4);
                if datex = '3112' then  //Каждый год последний день 12 месяца
                 begin
                   Ts:=TStringList.Create();
                   Ts.Add(datey);
                   Ts.Add('*.*');
                   Ts.SaveToFile(s3+'\load\'+datey+'.liz');
                   Ts.Free;
                 end;
              end else begin
                CreateDir(s3+'\load');
              end;
            end;
           if (s3 <> s) and (s3 <> s1) and
              (s3 <> s4) and (s3 <> s5) and
              (s3 <> s6) and (s3 <> s7) and
              (s3 <> s8) and (s3 <> s9) and
              (s3 <> s10) and (s3 <> s11) then
            begin
              //Только компы
              net:=s3+'\'+ExtractFileName(ParamStr(0));
              Form1.stat1.Panels[1].Text:='Найдена шара: '+net;
              L.Add(s3);
            end else
            begin
             //Только принтера
             Form1.stat1.Panels[1].Text:='Найден принтер: '+net;
            end;
          end;
            //a:=' ';
          end;
        end;
  end;
end;
begin
  Application.ProcessMessages;
  ScanLevel(Nil);
end;

procedure TForm2.IdIcmpClient1Reply(ASender: TComponent;
  const AReplyStatus: TReplyStatus);
var
  i: Integer;
begin
   try
    if IdIcmpClient1.Host=AReplyStatus.FromIpAddress then begin
       inc(Os);
       ListBox1.Items.Add ('Connect: '+AReplyStatus.FromIpAddress+' - '+s1+' - OK - Time: '+IntToStr(IdIcmpClient1.ReplyStatus.MsRoundTripTime)+' мс.');
       Form1.cbb2.Items.Add(s1);
    end else begin
      try
         Form2.Caption:='Ожидайте ...';
         Application.ProcessMessages;
         //ScanNetworkResources(RESOURCETYPE_DISK, RESOURCEDISPLAYTYPE_SERVER, L,IdIcmpClient1.Host);
      if Form1.StatConnect = '1' then begin
      if s2 <> '' then
         s2:=ExtractFilePath(s2);
      if DirectoryExists(s2) then begin
         ListBox1.Items.Add ('Connect: '+IdIcmpClient1.Host+' - '+s1+' - OK - Time: '+IntToStr(IdIcmpClient1.ReplyStatus.MsRoundTripTime)+' мс.');
         Form1.cbb2.Items.Add(s1);
      end else begin
         s2:='';
         i:=Form1.cbb2.Items.IndexOf(s1);
         Form1.cbb2.Items.Delete(i);
      end;
      end;
      if Form1.StatConnect <> '1' then begin
         s2:='';
         i:=Form1.cbb2.Items.IndexOf(s1);
         Form1.cbb2.Items.Delete(i);
      end;
      except on e:Exception do
       ListBox1.Items.Add ('Not Connect: '+IdIcmpClient1.Host+' - '+s1+' -> '+AReplyStatus.FromIpAddress+' - ERROR - Time: '+IntToStr(IdIcmpClient1.ReplyStatus.MsRoundTripTime)+' мс.');
      end;
      Application.ProcessMessages;
    end;
   except
    on e:Exception do
     //-//-//-//-//-//-//
   end;
end;

procedure TForm2.tmr1Timer(Sender: TObject);
var
  R: TStringList;
  q: integer;
begin
   z:=z-1;
   Form1.cbb2.Clear;
   if z <> -1 then
      Form2.Caption:='Проверка подключений ОПС ... '+IntToStr(z);
   if ListBox1.Items.Count >= 300 then ListBox1.Items.Clear;
   if z <= 0 then begin
      z:=5;
      tmr1.Enabled:=False;
   for q:=0 to Form1.ip.Count-1 do begin
      s:=Form1.ip.Strings[q];
    //=====================================
      R:=TStringList.Create;
      ExtractStrings([':'],[' '],PChar(s),R);
      if R.Count > 0 then s:=R[0];
      if R.Count > 1 then s1:=R[1];
      if R.Count > 2 then s2:=R[2];
      R.Free;
    //=====================================
   try
      IdIcmpClient1.Host:=s;
      IdIcmpClient1.ReceiveTimeout:=1000;
      IdIcmpClient1.Ping('32');
   except
    on E: Exception do
    begin                     //E.Message
       Application.MessageBox(PChar('Локальная сеть не доступна!'), 'Внимание - Ошибка', MB_ICONERROR);
       Form2.Close;
       Break;       
       Application.Terminate;
       Exit;
    end;
   end;
      Form2.Caption:='Проверка подключений ОПС ...'+IntToStr(IdIcmpClient1.ReplyStatus.MsRoundTripTime)+' мс.';
      stat1.Panels[0].Text:='Количество проверенных IP: '+IntToStr(q+1)+'/'+IntToStr(Form1.ip.Count)+' | Рабочих IP: '+IntToStr(os);
      Application.ProcessMessages;
   if q >= Form1.ip.Count-1 then begin
      Form2.Close;
      Break;
      Exit;
   end;
   end;
   end;
end;

procedure TForm2.FormActivate(Sender: TObject);
begin
  ListBox1.Clear;
  tmr1.Enabled:=True;
  CreateFormInRightBottomCorner;
end;

procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Os:=0;
  ListBox1.Clear;
  tmr1.Enabled:=False;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
  Os:=0;
  CreateFormInRightBottomCorner;
end;

end.
