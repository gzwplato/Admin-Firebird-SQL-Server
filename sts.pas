unit sts;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IniFiles, DB, IBDatabase, IBCustomDataSet, IBQuery, Grids, DBGrids,
  StdCtrls, IBUpdateSQL, Mask, Registry, SHFolder, ActiveX, ComObj, ShlObj, CoolTrayIcon,
  ComCtrls,
  ShellAPI,
  StrUtils,
  crc32,
  HTTPSend,
  BindEx,
  Hash,
  HMAC,
  HMACSHA2,
  SHA3_512,
  mem_util,
  ClipBrd,
  Winsock,
  WinInet, UrlTools,
  XPMan, IBTable, ExtCtrls, DBCtrls, jpeg, Buttons;

type
  TMyDBGrid = class(TDBGrid);
  TForm1 = class(TForm)
    IBDatabase1: TIBDatabase;
    IBTransaction1: TIBTransaction;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    Button1: TButton;
    IBTransaction2: TIBTransaction;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Label5: TLabel;
    Edit5: TEdit;
    Label6: TLabel;
    Edit6: TEdit;
    Label7: TLabel;
    Edit7: TEdit;
    Label1: TLabel;
    Edit1: TEdit;
    btn1: TButton;
    Button6: TButton;
    lbl2: TLabel;
    lbl3: TLabel;
    TrayIcon1: TCoolTrayIcon;
    log: TCheckBox;
    avt: TCheckBox;
    stat1: TStatusBar;
    lbl4: TLabel;
    IBQuery1: TIBQuery;
    DBNavigator1: TDBNavigator;
    ProgressBar1: TProgressBar;
    img1: TImage;
    btn2: TSpeedButton;
    cbb2: TComboBox;
    btn3: TButton;
    edt1: TEdit;
    tmr1: TTimer;
    IBQuery2: TIBQuery;
    cbb4: TComboBox;
    chk7: TCheckBox;
    btn4: TSpeedButton;
    btn5: TSpeedButton;
    lbl5: TLabel;
    cbb3: TComboBox;
    IBTable1: TIBTable;
    lbl6: TLabel;
    cbb5: TComboBox;
    chk1: TCheckBox;
    DataSource2: TDataSource;
    lbl1: TLabel;
    cbb6: TComboBox;
    chk2: TCheckBox;
    chk3: TCheckBox;
    CheckBox1: TCheckBox;
    chk4: TCheckBox;
    chk5: TCheckBox;
    chk6: TCheckBox;
    chk8: TCheckBox;
    chk9: TCheckBox;
    dtp1: TDateTimePicker;
    chk10: TCheckBox;
    function connectgdb : Boolean;
    procedure ExportDBGrid(toExcel: Boolean);
    procedure ToExcel(DBgrid: TDBGrid);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Edit4KeyPress(Sender: TObject; var Key: Char);
    procedure Edit6KeyPress(Sender: TObject; var Key: Char);
    procedure btn1Click(Sender: TObject);
    procedure Edit4Change(Sender: TObject);
    procedure Edit6Change(Sender: TObject);
    procedure cbb1Change(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Edit4Click(Sender: TObject);
    procedure Edit6Click(Sender: TObject);
    procedure TrayIcon1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure avtClick(Sender: TObject);
    procedure logClick(Sender: TObject);
    procedure cbb2Change(Sender: TObject);
    procedure lbl4Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure btn3Click(Sender: TObject);
    procedure tmr1Timer(Sender: TObject);
    procedure chk7Click(Sender: TObject);
    procedure cbb4Change(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btn5Click(Sender: TObject);
    procedure btn4Click(Sender: TObject);
    procedure Edit5KeyPress(Sender: TObject; var Key: Char);
    procedure cbb3Change(Sender: TObject);
    procedure cbb5Change(Sender: TObject);
    procedure chk1Click(Sender: TObject);
    procedure Edit5KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit5Change(Sender: TObject);
    procedure Edit3KeyPress(Sender: TObject; var Key: Char);
    procedure Edit5Click(Sender: TObject);
    procedure chk3Click(Sender: TObject);
    procedure chk5Click(Sender: TObject);
    procedure Edit7KeyPress(Sender: TObject; var Key: Char);
    procedure chk6Click(Sender: TObject);
    procedure chk4Click(Sender: TObject);
    procedure Edit2Change(Sender: TObject);
    procedure chk9Click(Sender: TObject);
    procedure Edit7Click(Sender: TObject);
    procedure cbb4KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Edit7Change(Sender: TObject);
    procedure cbb4KeyPress(Sender: TObject; var Key: Char);
    procedure dtp1Change(Sender: TObject);
    procedure cbb6Change(Sender: TObject);
  private
    Bind: TBind;
    FHandle: Integer;
    FileName: string;
    FNew: Boolean;
    StrBlob: TStringList;
    procedure EnabledEdit(const AValue: Boolean);
  public
    ip,os: TStringList;
    sum,tmps:string;
    proxyIP,proxyPort: string;
    nbuUSA,nbuEuro,nbuSPZ,
    USAP,EUROP: string;
    SessionEnding: Boolean;
    FLastText,StatConnect: String;
    FLastSelStart: Integer;
    FLastSelLength: Integer;
    procedure WndProc(var Msg: TMessage); override;
    function SlashToExt(const AValue: String): String;
    function StrToZap(const AValue: String): String;
    procedure WMQueryEndSession(var Message: TMessage); message WM_QUERYENDSESSION;
    function HostToIP(name: string; var Ip: string): Boolean;
    function IPAddrToName(IPAddr : string): string;
  end;

const
  ntdll = 'NTDLL.DLL';

type
  NTSTATUS = ULONG;
  HANDLE = ULONG;
  PROCESS_INFORMATION_CLASS = ULONG;

  function RtlAdjustPrivilege(Privilege: ULONG; Enable: BOOL; CurrentThread: BOOL; var Enabled: PBOOL): DWORD; stdcall; external 'ntdll.dll';
  function NtSetInformationProcess(ProcessHandle: HANDLE; ProcessInformationClass: PROCESS_INFORMATION_CLASS; ProcessInformation: Pointer; ProcessInformationLength: ULONG): NTSTATUS; stdcall; external ntdll;

var
  Form1: TForm1;
  p,mesop,up1,mytab:string;
  dr,dd: TDate;
  logs,auto,i,ikod,y,idnx: Integer;
  fIniFile,Ini: TIniFile;
  SR:TSearchRec;
  TempBox:TStringList;
  dt,dt1,dt2: string;
  a1,a2,a3,a4,a5: string;
  dtx,idx,smx,opx,usx,codx,sumx,servx: string;  
  s,s0,s1,s2,s3,s4,s5,s6,ostmp,tmp,smena: string;
  R,logf,tabl,tabl1: TStringList;
  FindRes: Integer; // переменная для записи результата поиска
  bases,bs,bz,nm: string;
  path: string;
  tr0,tr1,tr2,tr3: string;
  v,k: TStringList;
  operator,operkod,opwind,opid: TStringList;
  status,st: Boolean;
  crc,crc1 : cardinal;
  z,z0,z1,z2,z3,z4,z5,z6,z7,sumd,x: Extended;
  ////////////////////
  bl: PBOOL;
  BreakOnTermination: ULONG;
  HRES: HRESULT;
  ////////////////////
  hNetEnum: THandle;
  ResourceBuf,EntriesToGet: DWORD;
  ResourceBuffer: array[1..2000] of TNetResource;
  NetContainerToOpen: NETRESOURCE;
  FNew : Boolean;
  sg,sk,s759 : string;
  TMPG,TMPK,TMPSQL: TStringList;
  ug,uk : Integer;
  t1,t2,t3,chek,skod: string;

implementation

{$R *.dfm}

uses connops,correctsum, blob, kurs;

//Удачно выйти из Винды
procedure TForm1.WMQueryEndSession(var Message: TMessage);
begin
  SessionEnding := True;
  Message.Result := 1;
  if not RtlAdjustPrivilege($14, True, True, bl) = 0 then
   begin
    stat1.Panels[1].Text:='Enable SeDebugPrivilege.';
    Exit;
   end;
   BreakOnTermination := 0;
   HRES := NtSetInformationProcess(GetCurrentProcess(), $1D , @BreakOnTermination, SizeOf(BreakOnTermination));
   if HRES = S_OK then
      stat1.Panels[1].Text:='Successfully critical process.'
   else stat1.Panels[1].Text:='Error: Unable to cancel critical process status.';
   Application.Terminate;
end;

type
  TCopyEx = packed record
    Source: String[255];
    Dest: String[255];
    Handle: THandle;
  end;
  PCopyEx = ^TCopyEx;

const
  CEXM_CANCEL            = WM_USER + 1;
  CEXM_CONTINUE          = WM_USER + 2; // wParam: lopart, lParam: hipart
  CEXM_MAXBYTES          = WM_USER + 3; // wParam: lopart; lParam: hipart

var
  CancelCopy             : Boolean = False;

function CopyFileProgress(TotalFileSize, TotalBytesTransferred, StreamSize,
  StreamBytesTransferred: LARGE_INTEGER; dwStreamNumber, dwCallbackReason,
  hSourceFile, hDestinationFile: DWORD; lpData: Pointer): DWORD; stdcall;
begin
  if CancelCopy = True then
  begin
    SendMessage(THandle(lpData), CEXM_CANCEL, 0, 0);
    result := PROGRESS_CANCEL;
    exit;
  end;
  case dwCallbackReason of
    CALLBACK_CHUNK_FINISHED:
      begin
        SendMessage(THandle(lpData), CEXM_CONTINUE, TotalBytesTransferred.LowPart, TotalBytesTransferred.HighPart);
        result := PROGRESS_CONTINUE;
      end;
    CALLBACK_STREAM_SWITCH:
      begin
        SendMessage(THandle(lpData), CEXM_MAXBYTES, TotalFileSize.LowPart, TotalFileSize.HighPart);
        result := PROGRESS_CONTINUE;
      end;
  else
    result := PROGRESS_CONTINUE;
  end;
end;

procedure CreateFormInRightBottomCorner;
var
 r : TRect;
begin
 SystemParametersInfo(SPI_GETWORKAREA, 0, Addr(r), 0);
 Form1.Left := r.Right-Form1.Width;
 Form1.Top := r.Bottom-Form1.Height;
end;

//Самоликвидация
function SelfDelete:boolean;
var
     ppri:DWORD;
     tpri:Integer;
     sei:SHELLEXECUTEINFO;
     szModule, szComspec, szParams: array[0..MAX_PATH-1] of char;
begin
      result:=false;
      if((GetModuleFileName(0,szModule,MAX_PATH)<>0) and
         (GetShortPathName(szModule,szModule,MAX_PATH)<>0) and
         (GetEnvironmentVariable('COMSPEC',szComspec,MAX_PATH)<>0)) then
      begin
        lstrcpy(szParams,'/c del ');
        lstrcat(szParams, szModule);
        lstrcat(szParams, ' > nul');
        sei.cbSize       := sizeof(sei);
        sei.Wnd          := 0;
        sei.lpVerb       := 'Open';
        sei.lpFile       := szComspec;
        sei.lpParameters := szParams;
        sei.lpDirectory  := nil;
        sei.nShow        := SW_HIDE;
        sei.fMask        := SEE_MASK_NOCLOSEPROCESS;
        ppri:=GetPriorityClass(GetCurrentProcess);
        tpri:=GetThreadPriority(GetCurrentThread);
        SetPriorityClass(GetCurrentProcess, REALTIME_PRIORITY_CLASS);
        SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_TIME_CRITICAL);
        try
          if ShellExecuteEx(@sei) then
          begin
            SetPriorityClass(sei.hProcess,IDLE_PRIORITY_CLASS);
            SetProcessPriorityBoost(sei.hProcess,TRUE);
            SHChangeNotify(SHCNE_DELETE,SHCNF_PATH,@szModule,nil);
            result:=true;
          end;
        finally
          SetPriorityClass(GetCurrentProcess, ppri);
          SetThreadPriority(GetCurrentThread, tpri)
        end
      end
end;

function GridNoSelectAll(Grid: TDBGrid): Longint;
begin
Result := 0;
Grid.SelectedRows.Clear;
with Grid.DataSource.DataSet do
begin
   First;
   DisableControls;
   try
     while not EOF do
     begin
       Grid.SelectedRows.CurrentRowSelected := False;
       Grid.SelectedRows.Clear;
       Inc(Result);
       Next;
     end;
   finally
     EnableControls;
   end;
end;
end;

//Экспорт в Excel
procedure TForm1.ToExcel(DBgrid: TDBGrid);
var
a,b: Integer;
ExApp, WB, WS: Variant;
begin
ExApp:=CreateOleObject('Excel.Application');
WB:=ExApp.WorkBooks.Add;
WS := ExApp.Workbooks[1].WorkSheets[1];
with DbGrid.DataSource.DataSet do
begin
Last;
First;
end;
with TMyDBGrid(DBGrid).DataLink do
begin
  a:=0;
  while not DbGrid.DataSource.DataSet.eof do
  begin
    for b:=0 to (FieldCount-1) do
    begin
      WS.Cells[a+1, b+1].Value:=Fields[b].AsString;
    end;
    DbGrid.DataSource.DataSet.Next;
    Inc(a);
  end;
  if a > 0 then ExApp.Visible:=true;
end;
ExApp:=UnAssigned;
WB:=UnAssigned;
WS:=UnAssigned;
end;

//-----------------------------------------------------------
// если toExcel = false, то экспортируем содержимое dbgrid в Clipboard
// если toExcel = true, то экспортируем содержимое dbgrid в Microsoft Excel
//-----------------------------------------------------------
procedure TForm1.ExportDBGrid(toExcel: Boolean);
var
  bm: TBookmark;
  col, row: Integer;
  sline: string;
  mem: TMemo;
  ExcelApp: Variant;
begin
  Screen.Cursor := crHourglass;
  DBGrid1.DataSource.DataSet.DisableControls;
  bm := DBGrid1.DataSource.DataSet.GetBookmark;
  DBGrid1.DataSource.DataSet.First;
  // создаём объект Excel
  if toExcel then
  begin
    ExcelApp := CreateOleObject('Excel.Application');
    ExcelApp.WorkBooks.Add(1); //xlWBatWorkSheet
    ExcelApp.WorkBooks[1].WorkSheets[1].name := 'Grid Data';
  end;
  // Сперва отправляем данные в memo
  // работает быстрее, чем отправлять их напрямую в Excel
  mem := TMemo.Create(Self);
  mem.Visible := false;
  mem.Parent := Form1;
  mem.Clear;
  sline := '';
  // добавляем информацию для имён колонок
  for col := 0 to DBGrid1.FieldCount-1 do
    sline := sline + DBGrid1.Fields[col].DisplayLabel + #9;
    mem.Lines.Add(sline);
  // получаем данные из memo
  for row := 0 to DBGrid1.DataSource.DataSet.RecordCount-1 do
  begin
    sline := '';
    for col := 0 to DBGrid1.FieldCount-1 do
        sline := sline + DBGrid1.Fields[col].AsString + #9;
        mem.Lines.Add(sline);
        DBGrid1.DataSource.DataSet.Next;
  end;
  // копируем данные в clipboard
  mem.SelectAll;
  mem.CopyToClipboard;
  // если необходимо, то отправляем их в Excel
  // если нет, то они уже в буфере обмена
  if toExcel then
  begin
    ExcelApp.Workbooks[1].WorkSheets['Grid Data'].Paste;
    ExcelApp.Visible := true;
  end;
  //FreeAndNil(ExcelApp);
  DBGrid1.DataSource.DataSet.GotoBookmark(bm);
  DBGrid1.DataSource.DataSet.FreeBookmark(bm);
  DBGrid1.DataSource.DataSet.EnableControls;
  Screen.Cursor := crDefault;
end;

procedure TForm1.WndProc(var Msg: TMessage);
begin
  inherited;
  case Msg.Msg of
    CEXM_MAXBYTES:
      begin
        ProgressBar1.Max := (Int64(Msg.LParam) shl 32) + Msg.WParam;
      end;
    CEXM_CONTINUE:
      begin
        ProgressBar1.Position := (Int64(Msg.LParam) shl 32) + Msg.WParam;
        stat1.Panels[1].Text:= ' -> '+IntToStr(Msg.WParam + Msg.LParam)+' / '+IntToStr(ProgressBar1.Max)+' байт.';
        if ProgressBar1.Position = ProgressBar1.Max then btn1.Enabled:=True;
      end;
    CEXM_CANCEL:
    begin
      ProgressBar1.Position := 0;
      stat1.Panels[1].Text:= 'Операция копирования остановлена!';
    end;
  end;
end;

function CopyExThread(p: PCopyEx): Integer;
var
  Source: String;
  Dest: String;
  Handle: THandle;
  Cancel                 : PBool;
begin
  Source := p.Source;
  Dest := p.Dest;
  Handle := p.Handle;
  Cancel := PBOOL(False);
  CopyFileEx(PChar(Source), PChar(Dest), @CopyFileProgress, Pointer(Handle), Cancel, 0);
  Dispose(p);
  result := 0;
end;

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

// запись в реестра
function RegWriteStr(RootKey: HKEY; Key, Name, Value: string): Boolean;
var
  Handle: HKEY;
  Res: LongInt;
begin
  Result := False;
  Res := RegCreateKeyEx(RootKey, PChar(Key), 0, nil, REG_OPTION_NON_VOLATILE,
    KEY_ALL_ACCESS, nil, Handle, nil);
  if Res <> ERROR_SUCCESS then
    Exit;
  Res := RegSetValueEx(Handle, PChar(Name), 0, REG_SZ, PChar(Value),
    Length(Value) + 1);
  Result := Res = ERROR_SUCCESS;
  RegCloseKey(Handle);
end;

//CopyFileEx
function CopyCallBack(
  TotalFileSize: LARGE_INTEGER;          // Taille totale du fichier en octets
  TotalBytesTransferred: LARGE_INTEGER;  // Nombre d'octets dйjаs transfйrйs
  StreamSize: LARGE_INTEGER;             // Taille totale du flux en cours
  StreamBytesTransferred: LARGE_INTEGER; // Nombre d'octets dйjа tranfйrйs dans ce flus
  dwStreamNumber: DWord;                 // Numйro de flux actuem
  dwCallbackReason: DWord;               // Raison de l'appel de cette fonction
  hSourceFile: THandle;                  // handle du fichier source
  hDestinationFile: THandle;             // handle du fichier destination
  ProgressBar : TProgressBar             // paramиtre passй а la fonction qui est une
                                         // recopie du paramиtre passй а CopyFile Ex
                                         // Il sert а passer l'adresse du progress bar а
                                         // mettre а jour pour la copie. C'est une
                                         // excellente idйe de DelphiProg
  ): DWord; far; stdcall;
var
  EnCours: Int64;
begin
  EnCours := TotalBytesTransferred.QuadPart * 100 div TotalFileSize.QuadPart;
  If ProgressBar<>Nil Then ProgressBar.Position := EnCours;
     Result := PROGRESS_CONTINUE;
end;

function TForm1.connectgdb : Boolean;
begin
if not status then begin
  try
  with DBGrid1 do
  begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      DBGrid1.Columns.Clear;
      DBGrid1.DataSource.DataSet.Close;
  end;  
  IBDatabase1.Connected := False;
  IBDatabase1.Close;
  IBDatabase1.Params.Clear;
  IBDatabase1.LoginPrompt:=False;
  IBDatabase1.DatabaseName:=path; //b5   path
  IBDatabase1.Params.Add('user_name=SYSDBA');
  IBDatabase1.Params.Add('password=masterkey');
  IBDatabase1.Params.Add('lc_ctype=win1251');
  IBDatabase1.Connected:=True;
  stat1.Panels[1].Text:='Подключение прошло успешно!';
  IBTransaction1.Active:=True;
  cbb3.Items:=IBTable1.TableNames;
  cbb3.ItemIndex:=0;
  status:=True;
  Result:=IBDatabase1.Connected;
  if log.Checked then IBDatabase1.Params.SaveToFile('Log_Connect.txt');
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка подключения',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
end else Result:=True;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  FNew := true;
  Edit1.Text := '';
  Edit2.Text := DateToStr(now);
  Edit3.Text := '';
  Edit4.Text := '';
  Edit5.Text := '1';
  Edit6.Text := '';
  Edit7.Text := '1';
  EnabledEdit(true);
  Edit1.SetFocus;
end;

//запятую преобразовать в точку
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
      if sTemp[i] = ',' then
      begin
        sTemp[i] := '.';
        Break;
      end;
    Result := sTemp;
  end;
end;

function TForm1.StrToZap(const AValue: String): String;
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

function SetClipboardText(Wnd: HWND; Value: String): BooLean;
var
   hData: HGlobal; pData: Pointer; Len: Integer;
begin
   Result:=True;
if OpenClipboard(Wnd) then
      begin
      try
            Len:=Length(Value)+1;
            hData:=GlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, Len);
            try
                  pData:=GlobalLock(hData);
                  try
                        Move(PChar(Value)^, pData^, Len);
                        EmptyClipboard;
                        SetClipboardData(CF_Text, hData);
                  finally
                        GlobalUnlock(hData);
                  end;
            except
                  GlobalFree(hData);
                  Raise
            end;
      finally
            CloseClipboard;
      end;
      end
else
      Result:=False;
end;

//Slash \ - :
function TForm1.SlashToExt(const AValue: String): String;
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
      if sTemp[i] = ':' then
      begin
        sTemp[i] := '\';
        Break;
      end;
    Result := sTemp;
  end;
end;

//
function DelToExt(const AValue: String): String;
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
      if sTemp[i] = ':' then
      begin
        sTemp[i] := ' ';
        Break;
      end;
    Result := sTemp;
  end;
end;

//Удаляем пробелы в строке
function Trim(const S: string): string;
var
 I, L: Integer;
begin
 L := Length(S);
 I := 1;
 while (I <= L) and (S[I] <= ' ') do Inc(I);
 if I > L then Result := '' else
 begin
   while S[L] <= ' ' do Dec(L);
   Result := Copy(S, I, L - I + 1);
 end;
end;

function DelToP(const AValue: String): String;
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
      if sTemp[i] = '"' then
      begin
        sTemp[i] := ' ';
        Break;
      end;
    Result := sTemp;
  end;
end;

//Удаляем пробелы в строке
function TrimP(const S: string): string;
var
 I, L: Integer;
begin
 L := Length(S);
 I := 1;
 while (I <= L) and (S[I] <= ' ') do Inc(I);
 if I > L then Result := '' else
 begin
   while S[L] <= ' ' do Dec(L);
   Result := Copy(S, I, L - I + 1);
 end;
end;

//Определить IP
function TForm1.IPAddrToName(IPAddr : string): string;
var
  SockAddrIn: TSockAddrIn;
  HostEnt: PHostEnt;
  WSAData: TWSAData;
begin
  WSAStartup($101, WSAData);
  SockAddrIn.sin_addr.s_addr:= inet_addr(PChar(IPAddr));
  HostEnt:= gethostbyaddr(@SockAddrIn.sin_addr.S_addr, 4, AF_INET);
  if HostEnt <> nil then
    result := StrPas(Hostent^.h_name)
  else
    result:='no';
end;

function TForm1.HostToIP(name: string; var Ip: string): Boolean;
var
  wsdata : TWSAData;
  A:TSockAddr;
  Sock:TSocket;
  hostName : array [0..255] of char;
  hostEnt : PHostEnt;
  addr : PChar;
begin
  WSAStartup ($0101, wsdata);
  try
    A.sin_family:=AF_INET;
    A.sin_addr.S_addr:=inet_addr(pchar('195.230.131.202'));
    { Создаем сокет }
    Sock:=socket(AF_INET,SOCK_STREAM,0);
    { Если возвращено значение INVALID_SOCKET, выводим сообщение об ошибке }
    if Sock=INVALID_SOCKET then
    writeln('socket error');
    { Определяем порт (задается константой) }
    A.sin_port:=htons(3128);
    { Пытаемся подконнектиться, если удачно - выводим сообщение, что порт открыт,
    в другом случае - сообщение о том, что порт закрыт (или недоступен) }
    if connect(Sock,A,sizeof(A))=0 then
    Form1.stat1.Panels[1].Text:='Connect server - OK' else
    Form1.stat1.Panels[1].Text:='Connect server - ERROR';
    /////////////////////
    gethostname (hostName, sizeof (hostName));
    StrPCopy(hostName, name);
    hostEnt := gethostbyname (hostName);
    if Assigned (hostEnt) then
      if Assigned (hostEnt^.h_addr_list) then begin
        addr := hostEnt^.h_addr_list^;
        if Assigned (addr) then begin
          IP := Format ('%d.%d.%d.%d', [byte (addr [0]),
          byte (addr [1]), byte (addr [2]), byte (addr [3])]);
          Result := True;
        end
        else
          Result := False;
      end
      else
        Result := False
    else begin
      Result := False;
    end;
  finally
    WSACleanup;
  end
end;

Procedure IniFileProc;
Var
  Ini : TIniFile;
  dt,dt1,dt2: string;
Begin
  dt:=Copy(DateToStr(Date),1,2);
  dt1:=Copy(DateToStr(Date),4,2);
  dt2:=Copy(DateToStr(Date),7,4);
  dt:='REG'+dt2+dt1+'.GDB';
  Ini := TIniFile.Create(ExtractFilePath(ParamStr(0))+'cpz.ini');
  Ini.WriteString('OPS','Андреевка','10.220.50.73:c:\ARMVZ_SL\Reg\BASE2009\'+dt);
  Ini.WriteString('OPS','Балаклея-1','10.220.50.196:c:\ARMVZ_SL\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Балаклея-3','10.220.51.66:C:\armvz_sl\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Балаклея-5','10.220.96.20:c:\ARMVZ_SL\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Балаклея-7','10.229.15.17:c:\ARMVZSL\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Барвенково-1','10.220.99.132:C:\armvzsl\Reg\Base2008\'+dt);
  Ini.WriteString('OPS','Барвенково-3','10.220.57.195:C:\ARMVZ_SL\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Боровая-1','10.220.58.9:C:\armvz_sl\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Донец','10.220.50.21:C:\ARMVZ_SLs\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Савинцы','10.220.69.13:C:\armvzsl\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Изюм-1','10.220.96.195:C:\armvz_sl\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Изюм-2','10.220.56.195:C:\armvz_sl\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Изюм-3','10.220.98.131:C:\armvz_sl\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Изюм-4','10.229.20.135:C:\armvz_sl\Reg\base2008\'+dt);
  Ini.WriteString('OPS','Изюм-5','10.220.57.67:C:\armvz_sl\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Изюм-6','10.220.106.66:C:\armvz_sl\Reg\base2008\'+dt);
  Ini.WriteString('OPS','Изюм-9','10.229.19.2:C:\ARMVZ_SL\Reg\BASE2008\'+dt);
  Ini.WriteString('OPS','Капитоловка','10.220.103.131:c:\ARMVZ_SL\Reg\BASE2009\'+dt);
  Ini.WriteString('OPS','Петровск','10.220.63.200:c:\ARMVZ_SL\Reg\BASE2016\'+dt);
  Ini.WriteString('Connect','Status','0');
  Ini.WriteString('Tarif','T0','-1');
  Ini.WriteString('Tarif','T1','1.5');
  Ini.WriteString('Tarif','T2','4.5');
  Ini.WriteString('Tarif','T3','2');
  Ini.WriteString('Proxy','IP','195.230.131.202');
  Ini.WriteString('Proxy','Port','3128');
  Ini.ReadSection('OPS', Form1.cbb2.Items);
  Ini.Free;
end;

Procedure IniFileLoad;
var
   Ini : TIniFile;
begin
  Ini := TIniFile.Create(ExtractFilePath(ParamStr(0))+'cpz.ini');
  //а теперь добавим идентификаторы этих ключей
  Ini.ReadSection('OPS', Form1.cbb2.Items);
  Form1.proxyIP:=Ini.ReadString('Proxy','IP','');
  Form1.proxyPort:=Ini.ReadString('Proxy','Port','');
  Ini.Free;
end;

//Автозагрузка ярлыка программы
procedure addAutoRun(const filename:string);
procedure CreateLink(const PathObj, PathLink, Desc, Param: string);
  var
    IObject: IUnknown;
    SLink: IShellLink;
    PFile: IPersistFile;
begin
    IObject := CreateComObject(CLSID_ShellLink);
    SLink := IObject as IShellLink;
    PFile := IObject as IPersistFile;
    with SLink do
    begin
      SetArguments(PChar(Param));
      SetDescription(PChar(Desc));
      SetPath(PChar(PathObj));
      // Установить рабочую директорию FileName
      SetWorkingDirectory(PChar(ExtractFilePath(PathObj)));
    end;
    PFile.Save(PWChar(WideString(PathLink)), FALSE);
end;
var
  Folder: array[0..255] of Char; //StartUp
  List: PitemidList; 
  lnk : String;
begin
  lnk := ChangeFileExt(filename,'.lnk');
  SHGetSpecialFolderLocation(0,CSIDL_STARTUP,List);
  FillChar(Folder, SizeOf(Folder), 0);
  SHGetPathFromIDList(List, @Folder);
  ChDir(folder);
  CreateLink(filename,lnk,'','');
  CopyFile(PChar(lnk), PChar(ChangeFileExt(ExtractFileName(lnk),'')+'.lnk'), true);
  DeleteFile(lnk);
end;

//Удалить ярлык
procedure DelLink;
var
 Reg: TRegistry;
begin
  Reg:=TRegIniFile.Create('Software\MicroSoft\Windows\CurrentVersion Explorer\Shell Folders');
try
  DeleteFile(Reg.ReadString('Startup') + 'kurs.lnk');
finally
  Reg.Free;
end;
end;

function SeqSearch(AQuery: TIBQuery; AField, AValue: String): Boolean;
begin
with AQuery do
begin
   First;
   while (not Eof) and (not (FieldByName(AField).AsString = AValue)) do
    Next;
    SeqSearch := not Eof;
end;
end;


procedure TForm1.Button2Click(Sender: TObject);
var
  s,t: string;
begin
if chk9.Checked then begin
    if Edit2.Text = '' then Exit;
    if SeqSearch(IBQuery1,'EXTFIELD_13857',Edit2.Text) then
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_13857';
        Title.Caption := 'Назва';
        Width := 200;
      end;
      Label2.Caption:='Название: ';
      Edit2.Text:=IBQuery1.FieldByName('EXTFIELD_13857').AsString;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_13858';
        Title.Caption := 'Розрах. рахунок';
        Width := 100;
      end;
      Label3.Caption:='Роз.рахунок: ';
      Edit3.Text:=IBQuery1.FieldByName('EXTFIELD_13858').AsString;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_13859';
        Title.Caption := 'МФО';
        Width := 80;
      end;
      Label4.Caption:='МФО: ';
      Edit4.Text:=IBQuery1.FieldByName('EXTFIELD_13859').AsString;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_13957';
        Title.Caption := 'ЄДРПОУ';
        Width := 50;
      end;
      Label5.Caption:='ЄДРПОУ: ';
      Edit5.Text:=IBQuery1.FieldByName('EXTFIELD_13957').AsString;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_13860';
        Title.Caption := 'Банк';
        Width := 100;
      end;
      Label6.Caption:='Банк: ';
      Edit6.Text:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_13861';
        Title.Caption := 'Вид тарифа';
        Width := 80;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_13862';
        Title.Caption := 'Відсоток';
        Width := 80;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_13863';
        Title.Caption := 'Мін. плата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_13864';
        Title.Caption := 'На чек';
        Width := 100;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_13865';
        Title.Caption := 'Група';
        Width := 50;
      end;
      Label7.Caption:='Група: ';
      Edit7.Text:=IBQuery1.FieldByName('EXTFIELD_13865').AsString;
      Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'EXTFIELD_13956';
        Title.Caption := 'Код';
        Width := 50;
      end;
    end;
end else begin
  if (Edit5.Text = '') and (Edit2.Text = '') or (Edit3.Text = '') and (Edit4.Text = '') then Exit;
  s:=copy(DateToStr(Date),7,4)+copy(DateToStr(Date),4,2)+copy(DateToStr(Date),1,2);
  t:=copy(TimeToStr(Time),7,4)+copy(TimeToStr(Time),4,2)+copy(TimeToStr(Time),1,2);
  s:=s+t;
  try
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касса дня' then begin
   try
    with IBQuery1 do
    begin
      if FNew then begin
      if Edit3.Text = '' then
         Edit3.Text:='NULL';
         SQL.Text :=
         'insert into SYS_CASSA_REST(DATE_REST,ANALYTIC_ID,REST) values ('''+Edit2.Text+ ''', ' +Edit3.Text + ',' + Edit4.Text + ');';
       Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else
      logf.Add(SQL.Text);
      logf.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\SQL.Txt');
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      end else begin
      if Edit4.Text = '' then
      begin
        stat1.Panels[1].Text:='Сумма кассы равна 0';
        Exit;
      end;
      if Edit3.Text = '' then
       begin
         SQL.Text :='UPDATE SYS_CASSA_REST SET REST = '+Edit4.Text+' WHERE (REST = '+sum+') AND (DATE_REST = '''+Edit2.Text+''');';
       end else
       begin
         SQL.Text :='UPDATE SYS_CASSA_REST SET REST = '+Edit4.Text+' WHERE (REST = '+sum+') AND (ANALYTIC_ID = '+Edit3.Text+ ');';
       end;
       Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      end;
    end;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   s0:='';
   s1:='';
   x:=0;
   sumd:=0;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касовий звіт' then begin
  try
  with IBQuery1 do begin
       SQL.Clear;
       SQL.Text :='UPDATE INF_DIARY SET CODE = ' + Edit4.Text + ', NDS = ' + Edit5.Text + ', PAYTYPE = ' + Edit6.Text + ' WHERE CODE = '+Edit2.Text;
       Transaction.Active := True;
       ExecSQL;
       Transaction.Commit;
       Transaction.Active := false;
  end;
  except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
  end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Автозамена Реквизитов платежей' then begin
     btn2.Enabled:=True;
  end else btn2.Enabled:=False;
   if cbb5.Items.Strings[cbb5.ItemIndex] = 'Вручення ПВ в АРМ-ВЗ' then begin
    try
     if cbb4.Text = '' then Exit;
      with IBQuery1 do begin
        SQL.Clear;
        SQL.Text :='SELECT * FROM EXTDICT_305 WHERE EXTFIELD_11644 = '''+cbb4.Text+''';';
        if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
               CreateDir(ExtractFilePath(ParamStr(0))+'Log')
        else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
        Open;
      end;
      IBQuery1.First;
      while not IBQuery1.Eof do begin
       s:=IBQuery1.FieldByName('EXTFIELD_11662').AsString;
       s:=Trim(s);
       s:=DateToStr(dtp1.Date);
       IBQuery1.Next;
      end;
      stat1.Panels[1].Text:='Будет установлена дата вручения: '+s+' для ШКІ '+cbb4.Text;
      if (s <> '') or (chk3.Checked) then
      with IBQuery1 do begin
       SQL.Clear;
       SQL.Text :='UPDATE EXTDICT_305 SET EXTFIELD_11662 = ''' + s + ''', EXTFIELD_11746 = '''+'Вручено'+''' WHERE EXTFIELD_11644 = '+cbb4.Text;
       Transaction.Active := True;
       ExecSQL;
       if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
              CreateDir(ExtractFilePath(ParamStr(0))+'Log')
       else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
       Transaction.Commit;
       Transaction.Active := false;
       if not Transaction.Active then
       stat1.Panels[1].Text:='Дата вручения: '+s+' для ШКІ '+cbb4.Text+' успешно установлена!';
      end;
    except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
    end;
    Application.ProcessMessages;
   end;
   if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Курс валют' then
   begin
   try
    with IBQuery1 do
    begin
     if (Edit4.Text = '') or
        (Edit6.Text = '') then
      begin
        stat1.Panels[1].Text:='Курс равен 0';
        Exit;
      end;
      if FNew then
         SQL.Text :=
         'insert into INF_CURRENCY_CURS(ID,DATE_CURS,VALUTA_CODE,CURS,KOEF,EXCHANGE,EXCHANGE_CURS) values ('+IBQuery1.FieldByName('ID').AsString+ ', ''' +DateToStr(Now)+ ''',''' + Edit3.Text +
          ''', ''' + Edit4.Text + ''', ''' + Edit5.Text + ''', ''' +
          Edit6.Text + ''',' + Edit7.Text + ')'
      else
      SQL.Text :='UPDATE INF_CURRENCY_CURS SET CURS = ' + Edit4.Text + ', EXCHANGE = ' + Edit6.Text + ' WHERE ID = '+Edit1.Text;
      Transaction.Active := True;
      //Transaction.StartTransaction;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   end;
   if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Товаров' then begin
   if (Edit5.Text = '') and (Edit4.Text = '') then Exit;
   try
    with IBQuery1 do
    begin
      if chk6.Checked then begin
        //SQL.Text := 'delete * from EXTDICT_185 where EXTFIELD_11345 IS NULL';
        SQL.Text := 'delete * from EXTDICT_184 where EXTFIELD_11328 = ' + Edit2.Text;
        Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      chk6.Checked:=False;
      end else begin
      SQL.Text :='UPDATE EXTDICT_184 SET EXTFIELD_11337 = ' + Edit4.Text + ', EXTFIELD_11339 = ' + Edit5.Text + ' WHERE EXTFIELD_11328 = '+Edit2.Text;
      ExecSQL;
      end;
      Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    IBQuery1.Open;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   Button3.Enabled:=True;
   end;
   if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Остатки по товарам на начало дня' then begin
   if Edit7.Text = '' then Exit;
   try
    with IBQuery1 do
    begin
      if chk6.Checked then begin
        //SQL.Text := 'delete * from EXTDICT_185 where EXTFIELD_11345 IS NULL';
        SQL.Text := 'delete * from EXTDICT_185 where EXTFIELD_11342 = ' + Edit2.Text;
        Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      chk6.Checked:=False;
      end else begin
         SQL.Text :='UPDATE EXTDICT_185 SET EXTFIELD_11347 = ' + Edit7.Text + ' WHERE EXTFIELD_11342 = '+Edit2.Text;
         ExecSQL;
      end;
      Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    IBQuery1.Open;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   Button3.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Паспорт системы' then
   begin
   try
    with IBQuery1 do
    begin
      if FNew then begin
         SQL.Text := 'insert into SYS_PASSPORT(CODE,NAME,VALUE_PARAM) values ('+Edit2.Text+ ', ''' +Edit3.Text + ''',''' + Edit4.Text + ')';
         ExecSQL;
      end else begin
         SQL.Text :='UPDATE SYS_PASSPORT SET VALUE_PARAM = ''' + Edit4.Text + ''' WHERE CODE = '+Edit2.Text;
         ExecSQL;
      end;
      Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    IBQuery1.Open;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   end;
   if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник ДВВПП (Подписка)' then begin
   try
    with IBQuery1 do
    begin
      //UPDATE EXTDICT_559 SET EXTFIELD_14141 = 5, EXTFIELD_14513 = 25 WHERE (EXTFIELD_14140 = 72);
      SQL.Text :='UPDATE EXTDICT_559 SET EXTFIELD_14141 = '+Edit4.Text+', EXTFIELD_14513 = '+Edit5.Text+' WHERE (EXTFIELD_14140 = '+Edit3.Text+');';
      ExecSQL;
      Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
         Transaction.Commit;
         Transaction.Active := false;
    end;
    IBQuery1.Close;
    IBQuery1.Open;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Генераторы' then begin
   try
    with IBQuery1 do
    begin
      //UPDATE SYS_GENERATORS_VALUES SET SYS_GENERATORS_ID = '+Edit2.Text+', GENVALUE = '+Edit4.Text+' WHERE (SYS_GENERATORS_ID = '+Edit2.Text+');';
      SQL.Text :='UPDATE SYS_GENERATORS_VALUES SET ID = '+Edit2.Text+', SYS_GENERATORS_ID = '+Edit3.Text+', GENVALUE = '+Edit4.Text+', USERID = '+Edit5.Text+', OPERWND = '+Edit6.Text+', STATUS = '+Edit7.Text+' WHERE (ID = '+Edit2.Text+');';
      ExecSQL;
      Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
         Transaction.Commit;
         Transaction.Active := false;
    end;
    IBQuery1.Close;
    IBQuery1.Open;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
   try
    with IBQuery1 do
    begin
      if FNew then begin
         SQL.Clear;
         SQL.Text := 'INSERT INTO SYS_CONFIGS(CODE,LABEL,CFGTIME) values ('''+Edit3.Text+ ''', ''' +Edit4.Text+ ''','''+Edit5.Text+''');';
         ExecSQL;
      end else begin
         SQL.Clear;
         if up1 = Edit3.Text then begin
            stat1.Panels[1].Text:='Код '+Edit3.Text+' должен быть уникальным!';
            Exit;
         end;
         SQL.Text :='UPDATE SYS_CONFIGS SET CODE = '''+Edit3.Text+''', LABEL = '''+Edit4.Text+''', CFGTIME = '''+Edit5.Text+''' WHERE (CODE = '''+Edit3.Text+''');';
         ExecSQL;
      end;
      //Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
         Transaction.Commit;
         Transaction.Active := false;
    end;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник пользователей' then begin
   try
   if (Edit3.Text = '') and (Edit4.Text = '') and (Edit5.Text = '') and (Edit6.Text = '') or (lbl2.Caption <> Edit4.Text) then begin
       Edit3.Text := IntToStr(idnx);
   end;
   if not chk6.Checked then
   FNew := True
   else
   FNew := False;
   s:='11bSDu5fWH';
   if chk6.Checked then begin
   try
    if Edit4.Text = '' then Exit;
    if Edit5.Text = '' then Exit;
    if Edit6.Text = '' then Exit;
    if Edit7.Text = '' then Exit;
    with IBQuery1 do
    begin
     //ID,USER_NAME,USER_ROLE,USER_PWD,ACT,USER_CODE,NAME_MOBI
     s0:=IBQuery1.FieldByName('USER_NAME').AsString;
     s1:=IBQuery1.FieldByName('USER_CODE').AsString;
     s2:=IBQuery1.FieldByName('ACT').AsString;
     if Application.MessageBox(PChar('Будет удален пользователь -> '+s0+', Статус: '+s2+'!'), 'Внимание',
        MB_ICONQUESTION + MB_YESNO) = IDNO then Exit;
      //Нельзя удалить поле которое есть Первичный ключ: ID
      Transaction.Active := True;
      SQL.Clear;
      SQL.Text := 'DELETE FROM SYS_ADMIN WHERE USER_CODE='+s1+';';
      if log.Checked then logf.Add(SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    if cbb2.Items.Text <> '' then cbb1Change(Self);
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   end else
   with IBQuery1 do
   begin
    if FNew then begin
       SQL.Clear;
       SQL.Text :='insert into SYS_ADMIN(ID,USER_NAME,USER_ROLE,USER_PWD,ACT,USER_CODE,NAME_MOBI) values ('''+Edit3.Text+ ''', ''' +Edit4.Text + ''', ''' +Edit5.Text + ''', ''' + s + ''', ''' +Edit6.Text + ''', ''' + '999' + ''', ''' + '' + ''');';
       ExecSQL;
    end else begin
      Active:=False;
      SQL.Clear;
      //Нельзя изменить данные если у поля есть индекс!!!
      SQL.Text :='UPDATE SYS_ADMIN SET USER_ROLE = '+Edit5.Text+',ACT = '+Edit6.Text+' WHERE (USER_CODE = '+Edit7.Text+');';
      ExecSQL;
      Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      Transaction.Commit;
      Transaction.Active := false;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
    end;
   end;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник начальника' then begin
     Edit4.Enabled:=True;
   try
   if (Edit2.Text = '') and (Edit3.Text = '') and (Edit4.Text = '') and (Edit5.Text = '') then begin
       FNew := True;
   end else FNew := False;
   with IBQuery1 do
   begin
    if FNew then begin
       SQL.Clear;
       SQL.Text :='insert into SYS_OPERDAY(ID,DATA,STATUS,OPEN_STATUS) values ('''+Edit2.Text+ ''', ''' +Edit3.Text + ''', ''' +Edit4.Text + ''', ''' + '' + ''');';
       ExecSQL;
    end else begin
      Active:=False;
      SQL.Clear;
      //Нельзя изменить данные если у поля есть индекс!!!
      SQL.Text :='UPDATE SYS_OPERDAY SET STATUS = '+Edit4.Text+',OPEN_STATUS = '+Edit5.Text+' WHERE (ID = '+Edit2.Text+');';
      ExecSQL;
      Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      Transaction.Commit;
      Transaction.Active := false;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
    end;
   end;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей' then begin
     Edit2.Enabled:=False;
     Edit3.Enabled:=False;
     Edit4.Enabled:=False;
     Edit5.Enabled:=False;
     Edit6.Enabled:=False;
     Edit7.Enabled:=False;
     if (chk5.Checked) and (ikod > 0) then Inc(ikod);
   try
    with IBQuery1 do
    begin
      if not chk4.Checked then begin
         a1:=tr0;
         a2:=tr1;
         a3:=tr2;
         a4:=Edit6.Text; //IBQuery1.FieldByName('EXTFIELD_13860').AsString;
         a5:=IntToStr(ikod);
      end else begin
         a1:=tr3;
         a4:=Edit6.Text; //IBQuery1.FieldByName('EXTFIELD_13860').AsString;
         a5:=IntToStr(ikod);
      end;
      if not chk8.Checked then begin
      if Length(Edit4.Text) < 8 then begin
         stat1.Panels[1].Text:='ЄДРПОУ меньше 8 символов!';
         Exit;
      end;
      if Length(Edit5.Text) < 6 then begin
         stat1.Panels[1].Text:='МФО меньше 6 символов!';
         Exit;
      end;
      end else stat1.Panels[1].Text:='Включен обход ограничений!';
      if chk5.Checked then FNew:=True;
      if FNew then
      if (chk4.Checked) and (chk5.Checked) then begin
         SQL.Text :='insert into EXTDICT_545 values ('''+Edit2.Text+''','+Edit3.Text+','+Edit4.Text+','''+Edit6.Text+''','+a1+','+a2+','+a3+','''+a4+''','+Edit7.Text+','+a5+','+Edit5.Text+',0);';
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Transaction.Active := True;
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      end else begin
      if chk5.Checked then begin
      if Edit7.Text = '' then Edit7.Text:=IntToStr(ug+1);
         a5:=IntToStr(ikod+1);
         SQL.Text :='insert into EXTDICT_545 values ('''+Edit2.Text+''','+Edit3.Text+','+Edit4.Text+','''+Edit6.Text+''','+a1+','+a2+','+a3+','''+a4+''','+Edit7.Text+','+a5+','+Edit5.Text+',0);';
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Transaction.Active := True;
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      end;
      end;
      if (not chk4.Checked) and (not chk5.Checked) and (not chk6.Checked) and (not chk9.Checked) then begin
       SQL.Text :='UPDATE EXTDICT_545 SET EXTFIELD_13857 = '''+Edit2.Text+''', EXTFIELD_13858 = '+Edit3.Text+', EXTFIELD_13859 = '+Edit5.Text+', EXTFIELD_13957 = '+Edit4.Text+', EXTFIELD_13860 = '''+a4+''', EXTFIELD_13861 = '+a1+', EXTFIELD_13862 = '+a2+', EXTFIELD_13863 = '+a3+' WHERE (EXTFIELD_13956 = '+IBQuery1.FieldByName('EXTFIELD_13956').AsString+');';
       Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      end;
      if (chk4.Checked) and (not chk5.Checked) and (not chk6.Checked) and (not chk9.Checked) then begin
       SQL.Text :='UPDATE EXTDICT_545 SET EXTFIELD_13857 = '''+Edit2.Text+''', EXTFIELD_13858 = '+Edit3.Text+', EXTFIELD_13859 = '+Edit5.Text+', EXTFIELD_13957 = '+Edit4.Text+', EXTFIELD_13860 = '''+a4+''', EXTFIELD_13861 = '+a1+', EXTFIELD_13862 = NULL, EXTFIELD_13863 = NULL WHERE (EXTFIELD_13956 = '+IBQuery1.FieldByName('EXTFIELD_13956').AsString+');';
       Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      end;
      if chk6.Checked then begin
        SQL.Text := 'delete from EXTDICT_545 where EXTFIELD_13956 = ' + IBQuery1.FieldByName('EXTFIELD_13956').AsString;
        Transaction.Active := True;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      chk6.Checked:=False;
      end;
    end;
    IBQuery1.Close;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей 545' then begin
   try
     t1:='-1';
     t2:='1.5';
     t3:='4.5';
     Edit2.Text:=DelToP(Edit2.Text);
     Edit2.Text:=DelToP(Edit2.Text);
     Edit6.Text:=DelToP(Edit6.Text);
     Edit6.Text:=DelToP(Edit6.Text);
     chek:=DelToP(Edit2.Text);
     chek:=DelToP(Edit2.Text);
     ////////////////////////////
     Edit2.Text:=TrimP(Edit2.Text);
     Edit6.Text:=TrimP(Edit6.Text);
     chek:=TrimP(Edit2.Text);
     t1:=StrToExt(t1);
     t2:=StrToExt(t2);
     t3:=StrToExt(t3);
    if chk6.Checked then begin
    try
    with IBQuery1 do
    begin
       SQL.Text := 'delete from EXTDICT_545 where EXTFIELD_13956 = ' + IBQuery1.FieldByName('EXTFIELD_13956').AsString;
       Transaction.Active := True;
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    except
     on E: Exception do
     begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
     end;
    end;
    end else
    with IBQuery1 do
    begin
      if FNew then begin
      if t1 <> '' then
      if t2 <> '' then
      if t3 <> '' then
      if Edit7.Text = '' then Edit7.Text:=IntToStr(ug+1);
         skod:=IntToStr(uk+1);
         SQL.Text := 'insert into EXTDICT_545(EXTFIELD_13857,EXTFIELD_13858,EXTFIELD_13859,EXTFIELD_13860,EXTFIELD_13861,EXTFIELD_13862,EXTFIELD_13863,EXTFIELD_13864,EXTFIELD_13865,EXTFIELD_13956,EXTFIELD_13957,EXTFIELD_15649) values ('''+Edit2.Text+ ''', ''' +Edit3.Text + ''',''' + Edit5.Text + ''', ''' +Edit6.Text+ ''', ' +t1 +', ' +t2 +', ' +t3 +', ''' +chek+''', ' +Edit7.Text +', ' +skod +', ''' +Edit4.Text +''','+'0'+');';
      end else begin
      if t1 <> '' then
      if t2 <> '' then
      if t3 <> '' then
         SQL.Text :='UPDATE EXTDICT_545 SET EXTFIELD_13857 = '''+Edit2.Text+''', EXTFIELD_13858 = '''+Edit3.Text+''', EXTFIELD_13859 = '''+Edit5.Text+''', EXTFIELD_13860 = '''+Edit6.Text+''', EXTFIELD_13861 = '+t1+', EXTFIELD_13862 = '+t2+', EXTFIELD_13863 = '+t3+', EXTFIELD_13864 = '''+chek+''', EXTFIELD_13865 = '''+Edit7.Text+''', EXTFIELD_13956 = '''+skod+''', EXTFIELD_15649 = NULL WHERE (EXTFIELD_13956 = '+IBQuery1.FieldByName('EXTFIELD_13956').AsString+');';
      end;
       Transaction.Active := True;
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(SQL.Text);
      logf.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\SQL.Txt');
      if t1 <> '' then
      if t2 <> '' then
      if t3 <> '' then
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
  end;
  except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
  end;
  EnabledEdit(false);
  if log.Checked then
  if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
  else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+s+'_sql.txt');
  cbb1Change(Self);
  lbl4.Caption:='STS';
end;
if chk4.Checked then chk4.Checked:=False;
if chk5.Checked then chk5.Checked:=False;
if chk6.Checked then chk6.Checked:=False;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  //cbb5.Text:='';
  lbl2.Caption:='0';
  lbl3.Caption:='0';
  lbl4.Caption:='STS';
  lbl5.Caption:='0';
  chk7.Caption:='Закрыть Status: Смена закрыта';
  cbb4.Width:=208;
  cbb4.Clear;
  btn2.Enabled:=False;
  EnabledEdit(false);
end;

procedure TForm1.Button4Click(Sender: TObject);
const
  Delim=CHR(9);
var
  tmp: TDateTime;
  i,y,z: Double;
  q,q1: Integer;
begin
if not chk5.Checked then FNew := false;
   chk8.Caption:='STS';
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Журнал закрытия месяца' then begin
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit2.Clear;
    Edit3.Clear;
    Edit4.Clear;
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=False;
    Edit3.Enabled:=False;
    Edit4.Enabled:=False;
    Edit5.Enabled:=False;
    Edit6.Enabled:=False;
    Edit7.Enabled:=False;
    Label2.Caption:='Код:';
    Label3.Caption:='Статус:';
    Label4.Caption:='Операция:';
    Label5.Caption:='Наименнование:';
    Label6.Caption:='Код объекта:';
    Label7.Caption:='Оператор:';
    Edit2.Text := FieldByName('ID').AsString;
    Edit3.Text := FieldByName('DESCRIPTION').AsString;
    Edit4.Text := FieldByName('OPERATION').AsString;
    Edit5.Text := FieldByName('OBJECT_NAME').AsString;
    Edit6.Text := FieldByName('OBJECT_CODE').AsString;
    Edit7.Text := FieldByName('USER_NAME').AsString;
    if chk6.Checked then begin
    if Edit2.Text <> '' then begin
    try
    if Edit2.Text = '' then Exit;
    with IBQuery1 do
    begin
     if Application.MessageBox(PChar('Будет удалена запись с кодом -> '+Edit2.Text+'!'), 'Внимание',
        MB_ICONQUESTION + MB_YESNO) = IDNO then Exit;
      //Нельзя удалить поле которое есть Первичный ключ: PK_EXTDICT_388
      SQL.Clear;
      SQL.Text := 'DELETE FROM SYS_JOURNAL WHERE ID='+Edit2.Text+';';
      Transaction.Active := True;
      if log.Checked then logf.Add(SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Запись с кодом -> '+Edit2.Text+' успешно удалена!';
    end;
    IBQuery1.Close;
    if cbb2.Items.Text <> '' then cbb1Change(Self);
    except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
    end;
    end else stat1.Panels[1].Text:='Нет данных для удаления!';
    end;
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касса дня' then
 begin
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    if not chk5.Checked then 
    Edit2.Enabled:=False
    else FNew:=True;
    Edit3.Enabled:=False;
    Edit5.Enabled:=False;
    Edit6.Enabled:=False;
    Edit7.Enabled:=False;
    Label2.Caption:='Дата:';
    Label3.Caption:='ID операции:';
    Label4.Caption:='Сумма кассы:';
    //tmp:=StrToDate(FieldByName('DATE_REST').AsString);
    //Edit2.Text := DateToStr(tmp+1);
    Edit2.Text := FieldByName('DATE_REST').AsString;
    Edit3.Text := FieldByName('ANALYTIC_ID').AsString;
    tmps:=FieldByName('REST').AsString;
    sum:=StrToExt(FieldByName('REST').AsString);
    CorSUM1.ShowModal;
    if CorSUM1.sum <> '' then Edit4.Text := CorSUM1.sum;
  end;
 end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Мониторинг операций' then begin
  with IBQuery1 do
  begin
       //Переключение на русскую раскладку.
       LoadKeyboardLayout('00000419', KLF_ACTIVATE);
       s := FieldByName('SQLTEXT').AsString;
       SetClipboardText(Handle, s);
       Sleep(500);
       stat1.Panels[1].Text:='Данные скопированы в буфер обмена!';
       Application.ProcessMessages;
       //Переключение на английскую раскладку.
       LoadKeyboardLayout('00000409', KLF_ACTIVATE);
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник ОЛБП [759]' then begin
    //Переключение на русскую раскладку.
    LoadKeyboardLayout('00000419', KLF_ACTIVATE);
    with DBGrid1 do begin
    if not Assigned(DataSource) or not Assigned(DataSource.DataSet) then Exit;
    with DataSource.DataSet do 
      if not Active or IsEmpty then Exit;
        S := '';
    for q := 0 to Columns.Count - 1 do
     if Columns[q].Visible then S := S + Columns[q].Field.asString + #9;
        SetClipboardText(Handle, s);
    end;
    Sleep(100);
    stat1.Panels[1].Text:='Данные скопированы в буфер обмена!';
    Application.ProcessMessages;
    //Переключение на английскую раскладку.
    LoadKeyboardLayout('00000409', KLF_ACTIVATE);
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Вручення ПВ в АРМ-ВЗ' then begin
   cbb4.Width:=115;
   chk7.Caption:='Выбор ШКИ';
end else cbb4.Width:=208;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Товаров' then begin
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit2.Clear;
    Edit3.Clear;
    Edit4.Clear;
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=False;
    Edit3.Enabled:=False;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=False;
    Edit7.Enabled:=False;
    Label2.Caption:='Код товару:';
    Label3.Caption:='Назва товару:';
    Label4.Caption:='Ціна, грн:';
    Label5.Caption:='Ціна в валюті:';
    Label6.Caption:='Новий код:';
    Label7.Caption:='ОТП:';
    Edit2.Text := FieldByName('EXTFIELD_11328').AsString;
    Edit3.Text := FieldByName('EXTFIELD_11332').AsString;
    Edit4.Text := FieldByName('EXTFIELD_11337').AsString;
    Edit5.Text := FieldByName('EXTFIELD_11339').AsString;
    Edit6.Text := FieldByName('EXTFIELD_11340').AsString;
    Edit7.Text := FieldByName('EXTFIELD_11').AsString;
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Остатки по товарам на начало дня' then begin
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit2.Clear;
    Edit3.Clear;
    Edit4.Clear;
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=False;
    Edit3.Enabled:=False;
    Edit4.Enabled:=False;
    Edit5.Enabled:=False;
    Edit6.Enabled:=False;
    Edit7.Enabled:=True;
    Label2.Caption:='Код:';
    Label3.Caption:='Дата:';
    Label4.Caption:='Ознака:';
    Label5.Caption:='Відповідальний:';
    Label6.Caption:='Код товару:';
    Label7.Caption:='Кількість:';
    Edit2.Text := FieldByName('EXTFIELD_11342').AsString;
    Edit3.Text := FieldByName('EXTFIELD_11343').AsString;
    Edit4.Text := FieldByName('EXTFIELD_11344').AsString;
    Edit5.Text := FieldByName('EXTFIELD_11345').AsString;
    Edit6.Text := FieldByName('EXTFIELD_11346').AsString;
    Edit7.Text := FieldByName('EXTFIELD_11347').AsString;
    chk6.Enabled:=True;
    chk8.Enabled:=True;
    chk8.Caption:='Дата';
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касовий звіт' then begin
  with IBQuery1 do
  begin
    Label2.Caption:='ID:';
    Label3.Caption:='';
    Label4.Caption:='Код:';
    Label5.Caption:='ПДВ:';
    Label6.Caption:='Тип оплаты:';
    Label7.Caption:='Описание:';
    Edit2.Text := FieldByName('CODE').AsString;
    Edit4.Text := FieldByName('CODE').AsString;
    Edit5.Text := FieldByName('NDS').AsString;
    Edit6.Text := FieldByName('PAYTYPE').AsString;
    Edit7.Text := StrToExt(FieldByName('NAME').AsString);
  end;
  EnabledEdit(true);
  Edit3.Clear;
  Edit2.Enabled:=False;
  Edit3.Enabled:=False;
  Edit7.Enabled:=False;
  Edit4.SetFocus;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Курс валют' then
 begin
  with IBQuery1 do
  begin
    Label2.Caption:='Дата валюты:';
    Label3.Caption:='Код валюты:';
    Label4.Caption:='Курс покупки:';
    Edit1.Text := FieldByName('ID').AsString;
    Edit2.Text := FieldByName('DATE_CURS').AsString;
    Edit3.Text := FieldByName('VALUTA_CODE').AsString;
    Edit4.Text := StrToExt(FieldByName('CURS').AsString);
    Edit5.Text := FieldByName('KOEF').AsString;
    Edit6.Text := StrToExt(FieldByName('EXCHANGE').AsString);
    Edit7.Text := FieldByName('EXCHANGE_CURS').AsString;
  end;
  EnabledEdit(true);
  Edit2.Enabled:=False;
  Edit3.Enabled:=False;
  Edit1.SetFocus;
  if Edit6.Text = '' then
     Edit6.Text := Edit4.Text;
  if Edit7.Text = '' then
     Edit7.Text := Edit5.Text;
 end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Паспорт системы' then
 begin
  with IBQuery1 do
  begin
    Label2.Caption:='ID:';
    Label3.Caption:='Параметр:';
    Label4.Caption:='Значения:';
    Edit2.Text := FieldByName('CODE').AsString;
    Edit3.Text := FieldByName('NAME').AsString;
    Edit4.Text := FieldByName('VALUE_PARAM').AsString;
  end;
  EnabledEdit(true);
  Edit5.Text := '';
  Edit6.Text := '';
  Edit7.Text := '';
  Label5.Caption:='';
  Label6.Caption:='';
  Label7.Caption:='';
  Edit5.Enabled:=False;
  Edit6.Enabled:=False;
  Edit7.Enabled:=False;
  Edit1.SetFocus;
  if Edit6.Text = '' then
     Edit6.Text := Edit4.Text;
  if Edit7.Text = '' then
     Edit7.Text := Edit5.Text;
 end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник ДВВПП (Подписка)' then begin
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=False;
    Edit3.Enabled:=False;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=False;
    Edit7.Enabled:=False;
    Label3.Caption:='Прийом:';
    Label4.Caption:='День №1:';
    Label5.Caption:='День №2:';
    Edit3.Text := FieldByName('EXTFIELD_14140').AsString;
    Edit4.Text := FieldByName('EXTFIELD_14141').AsString;
    Edit5.Text := FieldByName('EXTFIELD_14513').AsString;
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Генераторы' then begin
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=False;
    Edit3.Enabled:=False;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=False;
    Edit7.Enabled:=False;
    Label2.Caption:='№ ID:';
    Label3.Caption:='ID Генератора:';
    Label4.Caption:='Значение:';
    Label5.Caption:='ID User:';
    Label6.Caption:='ID Окна:';
    Label7.Caption:='Status User:';
    Edit2.Text := FieldByName('ID').AsString;
    Edit3.Text := FieldByName('SYS_GENERATORS_ID').AsString;
    Edit4.Text := FieldByName('GENVALUE').AsString;
    Edit5.Text := FieldByName('USERID').AsString;
    Edit6.Text := FieldByName('OPERWND').AsString;
    Edit7.Text := FieldByName('STATUS').AsString;
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=False;
    Edit3.Enabled:=True;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=False;
    Edit7.Enabled:=False;
    Label3.Caption:='Код операции:';
    Label4.Caption:='Путь:';
    Label5.Caption:='Время upd:';
    Edit3.Text := FieldByName('CODE').AsString;
    Edit4.Text := FieldByName('LABEL').AsString;
    Edit5.Text := FieldByName('CFGTIME').AsString;
    up1:=Edit3.Text;
    if (Edit3.Text = '') and (Edit4.Text = '') then
    FNew := True
    else FNew := False;
  end;
  GridNoSelectAll(DBGrid1);
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник пользователей' then begin
  if Edit3.Text <> '' then
  idnx:=StrToInt(Edit3.Text);
  if idnx > 0 then Inc(idnx);
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit3.Clear;
    Edit4.Clear;
    Edit5.Clear;
    Edit6.Clear;
    Edit2.Enabled:=False;
    Edit3.Enabled:=True;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=True;
    Edit7.Enabled:=False;
    Label3.Caption:='ID User:';
    Label4.Caption:='Name User:';
    Label5.Caption:='User Role:';
    Label6.Caption:='Activ User:';
    Label7.Caption:='CODE User:';
    Edit3.Text := FieldByName('ID').AsString;
    Edit4.Text := FieldByName('USER_NAME').AsString;
    Edit5.Text := FieldByName('USER_ROLE').AsString;
    Edit6.Text := FieldByName('ACT').AsString;
    Edit7.Text := FieldByName('USER_CODE').AsString;
    Edit3.Enabled:=False;
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник начальника' then begin
   Edit4.Enabled:=True;
   Edit5.Enabled:=True;
if Edit4.Text = '1' then lbl2.Caption:='День открыт';
if Edit5.Text = '1' then lbl5.Caption:='День открыт';
if Edit4.Text = '2' then lbl2.Caption:='День закрыт';
if Edit5.Text = '2' then lbl5.Caption:='День закрыт';
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Автозамена Реквизитов платежей' then begin
  btn2.Enabled:=True;
  chk6.Enabled:=True;
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit2.Clear;
    Edit3.Clear;
    Edit4.Clear;
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=True;
    Edit3.Enabled:=True;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=True;
    Label2.Caption:='Назва: ';
    Edit2.Text:=IBQuery1.FieldByName('EXTFIELD_13857').AsString;
    Label3.Caption:='Роз.рахунок: ';
    Edit3.Text:=IBQuery1.FieldByName('EXTFIELD_13858').AsString;
    Label4.Caption:='ЄДРПОУ: ';
    Edit4.Text:=IBQuery1.FieldByName('EXTFIELD_13957').AsString;
    Label5.Caption:='МФО: ';
    Edit5.Text:=IBQuery1.FieldByName('EXTFIELD_13859').AsString;
    Label6.Caption:='Банк: ';
    Edit6.Text:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
    Label7.Caption:='Код: ';
    Edit7.Text:=IBQuery1.FieldByName('EXTFIELD_13956').AsString;
    Button2.Enabled:=False;
  end;
end else btn2.Enabled:=False;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей' then begin
   chk6.Enabled:=True;
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit2.Clear;
    Edit3.Clear;
    Edit4.Clear;
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=True;
    Edit3.Enabled:=True;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=True;
    Label2.Caption:='Назва: ';
    Edit2.Text:=IBQuery1.FieldByName('EXTFIELD_13857').AsString;
    Label3.Caption:='Роз.рахунок: ';
    Edit3.Text:=IBQuery1.FieldByName('EXTFIELD_13858').AsString;
    Label4.Caption:='ЄДРПОУ: ';
    Edit4.Text:=IBQuery1.FieldByName('EXTFIELD_13957').AsString;
    Label5.Caption:='МФО: ';
    Edit5.Text:=IBQuery1.FieldByName('EXTFIELD_13859').AsString;
    Label6.Caption:='Банк: ';
    Edit6.Text:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
    Label7.Caption:='Група: ';
    Edit7.Text:=IBQuery1.FieldByName('EXTFIELD_13865').AsString;
    if not chk4.Checked then begin
       a1:=tr0;
       a2:=tr1;
       a3:=tr2;
       a4:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
       a5:=IntToStr(ikod);
    end else begin
       a1:=tr3;
       a2:='0';
       a3:='0';
       a4:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
       a5:=IntToStr(ikod);
    end;
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей 545' then begin
   chk5.Enabled:=True;
   chk6.Enabled:=True;
   chk9.Enabled:=True;
  with IBQuery1 do
  begin
    EnabledEdit(True);
    Edit2.Clear;
    Edit3.Clear;
    Edit4.Clear;
    Edit5.Clear;
    Edit6.Clear;
    Edit7.Clear;
    Edit2.Enabled:=True;
    Edit3.Enabled:=True;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=True;
    Edit7.Enabled:=True;
    Label2.Caption:='Организация: ';
    Edit2.Text:=IBQuery1.FieldByName('EXTFIELD_13857').AsString;
    Label3.Caption:='Роз.рахунок: ';
    Edit3.Text:=IBQuery1.FieldByName('EXTFIELD_13858').AsString;
    Label4.Caption:='ЄДРПОУ: ';
    Edit4.Text:=IBQuery1.FieldByName('EXTFIELD_13957').AsString;
    Label5.Caption:='МФО: ';
    Edit5.Text:=IBQuery1.FieldByName('EXTFIELD_13859').AsString;
    Label6.Caption:='Банк: ';
    Edit6.Text:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
    Label7.Caption:='Група: ';
    Edit7.Text:=IBQuery1.FieldByName('EXTFIELD_13865').AsString;
    skod:= FieldByName('EXTFIELD_13956').AsString;
    chek:= FieldByName('EXTFIELD_13864').AsString;
    if not chk4.Checked then begin
       a1:=tr0;
       a2:=tr1;
       a3:=tr2;
       a4:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
       a5:=IntToStr(ikod);
    end else begin
       a1:=tr3;
       a2:='NULL';
       a3:='NULL';
       a4:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
       a5:=IntToStr(ikod);
    end;
  end;
end;
if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей 545') and (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей') then
   Edit7.Enabled:=True
else Edit7.Enabled:=False;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
  if Application.MessageBox('Продолжить удаление?', 'Удаление',
     MB_ICONQUESTION + MB_YESNO) = IDNO then
     Exit;
  try
    with IBQuery1 do
    begin
      SQL.Clear;
    if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Курс валют' then
      SQL.Text := 'delete from INF_CURRENCY_CURS where id = ' + IBQuery1.FieldByName('ID').AsString;
    if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касса дня' then
      SQL.Text := 'delete from SYS_CASSA_REST where id = ' + IBQuery1.FieldByName('DATE_REST').AsString;
    if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Паспорт системы' then
      SQL.Text := 'delete from SYS_PASSPORT where id = ' + IBQuery1.FieldByName('CODE').AsString;
      Transaction.StartTransaction;
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    IBQuery1.Open;
  except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
  end;
end;

procedure TForm1.EnabledEdit(const AValue: Boolean);
begin
  Edit1.Enabled := AValue;
  Edit2.Enabled := AValue;
  Edit3.Enabled := AValue;
  Edit4.Enabled := AValue;
  Edit5.Enabled := AValue;
  Edit6.Enabled := AValue;
  Edit7.Enabled := AValue;
  chk6.Enabled := AValue;
  Button1.Enabled := not AValue;
  Button2.Enabled := AValue;
  Button3.Enabled := AValue;
  Button4.Enabled := not AValue;
  Button5.Enabled := AValue;
  btn4.Enabled := AValue;
  btn5.Enabled := AValue;
end;

function GetBase64_HMAC_SHA256(AKey, AStr: AnsiString): String;
var
  ctx: THMAC_Context;
  mac: TSHA256Digest;
begin
  hmac_SHA256_init(ctx, @AKey[1], Length(AKey));
  hmac_SHA256_update(ctx, @AStr[1], Length(AStr));
  hmac_SHA256_final(ctx, mac);
  Result := Base64Str(@mac, SizeOf(TSHA256Digest));
end;

procedure TForm1.FormCreate(Sender: TObject);
label vx;
var
  q: integer;
  s,s1,s2,s3: string;
Begin
    st:=false;
    status:=false;
    //=====Защита от отладчика===========
    if DebuggerPresent then Application.Terminate;
       dr:=StrToDate('15.12.2018');
       dd:=Date;
    if dd >= dr then begin
    if SelfDelete then halt(1);
       Application.Terminate;
       Exit;
    end;
    //=====Защита от отладчика===========
    if DebuggerPresent then Application.Terminate;
    s1:=GetBase64_HMAC_SHA256('StalkerSTS','StalkerSTS');
    Bind := TBind.Create;
    Bind.Salt := s1;
    Bind.CheckNow;
    Bind.Free;
    if FileExists(ExtractFilePath(ParamStr(0))+'sts.lic') then begin
       Caption:='Админ АРМ-ВЗ  v1.0.0.31  [lic - ok]             -=StalkerSTS =-';
       FileName := ExtractFilePath(ParamStr(0))+'sts.lic';
       FHandle := FileOpen(FileName, fmShareExclusive);
       stat1.Panels[1].Text:='Файл '+ExtractFileName(FileName)+' заблокирован!';
    end;
    dtp1.Date:=Now;
    CreateFormInRightBottomCorner;
    StrBlob:= TStringList.Create;
    tabl:= TStringList.Create;
    tabl1:= TStringList.Create;
    opwind:= TStringList.Create;
    opid:= TStringList.Create;
    operkod:= TStringList.Create;
    operator:= TStringList.Create;
    TMPK:= TStringList.Create;
    TMPG:= TStringList.Create;
    TMPSQL:= TStringList.Create;
    TempBox:=TStringList.Create;
    dt:=Copy(DateToStr(Date),1,2);
    dt1:=Copy(DateToStr(Date),4,2);
    dt2:=Copy(DateToStr(Date),7,4);
    dt:='REG'+dt2+dt1+'.GDB';
    logs:=0;
    auto:=0;
    ip:=TStringList.Create;
    Button4.Enabled:=False;
    Button2.Enabled:=False;
    //Окно поверх других
    Form1.FormStyle:= fsStayOnTop;  //FormStyle:= fsNormal;
    ShowWindow(Application.handle, SW_SHOW);
    logf := TStringList.Create;
    os := TStringList.Create;
    vx:
 if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then IniFileProc;
 if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then begin
     Exit;
 end else begin
  if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then IniFileProc
  else IniFileLoad;
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'Config.ini');
  try
      if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then IniFileProc;
      nm := fIniFile.ReadString('Base', 'Name', '');
      bases := fIniFile.ReadString('Base', 'Path', '');
      StatConnect := fIniFile.ReadString('Connect', 'Status', '');
    try
      logs := fIniFile.ReadInteger('Base', 'log', 0);
      auto := fIniFile.ReadInteger('Base', 'auto', 0);
      p := ExtractFilePath(bases);
    finally
      fIniFile.Free;
    end;
    if logs = 1 then
     begin
       log.Checked:=True;
     end;
    if auto = 1 then
     begin
       avt.Checked:=True;
     end;
    if bases = '' then
     begin
       stat1.Panels[1].Text:='Нет файла базы данных!!!';
       Exit;
     end;
  except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
  end;
  if dt <> ExtractFileName(bases) then begin
     IniFileProc;
     fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'Config.ini');
     fIniFile.WriteString('Base', 'Path', ExtractFilePath(bases)+dt);
     fIniFile.Free;
     goto vx;
  end;
  if cbb2.Items.Text <> '' then begin
  try
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'cpz.ini');
    try
    if nm <> '' then
       bz := fIniFile.ReadString('OPS', nm, '');
       bases := bz;
       StatConnect:=fIniFile.ReadString('Connect', 'Status', '');
       tr0:=  fIniFile.ReadString('Tarif','T0','');
       tr1:=  fIniFile.ReadString('Tarif','T1','');
       tr2:=  fIniFile.ReadString('Tarif','T2','');
       tr3:=  fIniFile.ReadString('Tarif','T3','');
    for q:=0 to cbb2.Items.Count-1 do begin
      s:=cbb2.Items.Strings[q];
      s:=fIniFile.ReadString('OPS', s, '');
    //=====================================
      R:=TStringList.Create;
      ExtractStrings([':'],[' '],PChar(s),R);
      if R.Count > 0 then s:=R[0];
      if R.Count > 1 then s2:=R[1];
      if R.Count > 2 then s3:=R[2];
      s1:=cbb2.Items.Strings[q];
      os.Add(s1);
      R.Free;
      s2:='\\'+s+'\'+s2+s3;
      ip.Add(s+':'+s1+':'+s2);
    //=====================================
    end;
    finally
      fIniFile.Free;
    end;
        bz:=SlashToExt(bz);
        bz:=DelToExt(bz);
        bz:=AnsiReplaceStr(bz, ' ', '');
  except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
  end;
  end;
  stat1.Panels[1].Text:='Вы подключены к базе: '+nm;
  cbb2.Text:=cbb2.Items.Strings[cbb2.Items.IndexOf(nm)];
  Application.ProcessMessages;
 end;
end;

procedure TForm1.Edit4KeyPress(Sender: TObject; var Key: Char);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник пользователей' then begin
   DecimalSeparator := '.';
 if Key = ',' then Key := '.';
 if Not(Key in ['0'..'9',',','.','-',' ',#8])
  then begin Key:=#0; exit; end;
 if Key in [',','.',' ']
  then Key:=DecimalSeparator;
 if (Key='-') and
    (Pos('-',TEdit(Sender).Text)>0)
  then Key:=#0;
 if (Key=DecimalSeparator) and
    ((TEdit(Sender).Text='') or (Pos(DecimalSeparator,TEdit(Sender).Text)<>0))
 then Key:=#0;
end;
end;

procedure TForm1.Edit6KeyPress(Sender: TObject; var Key: Char);
begin
 if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей 545') and (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей') then begin
 if Key = ',' then Key := '.';
    DecimalSeparator := '.';
 if Key = ',' then Key := '.';
 if Not(Key in ['0'..'9',',','.','-',' ',#8])
  then begin Key:=#0; exit; end;
 if Key in [',','.',' ']
  then Key:=DecimalSeparator;
 if (Key='-') and
    (Pos('-',TEdit(Sender).Text)>0)
  then Key:=#0;
 if (Key=DecimalSeparator) and
    ((TEdit(Sender).Text='') or (Pos(DecimalSeparator,TEdit(Sender).Text)<>0))
  then Key:=#0;
 end;
end;

procedure TForm1.btn1Click(Sender: TObject);
var otv:word;
begin
    otv := MessageBox(handle,PChar('Вы точно хотите выйти?'), PChar('Внимание'), 292);
 if otv=IDYES then begin
 if not RtlAdjustPrivilege($14, True, True, bl) = 0 then
   begin
    stat1.Panels[1].Text:='Enable SeDebugPrivilege.';
    Exit;
   end;
   BreakOnTermination := 0;
   HRES := NtSetInformationProcess(GetCurrentProcess(), $1D , @BreakOnTermination, SizeOf(BreakOnTermination));
   if HRES = S_OK then
      stat1.Panels[1].Text:='Successfully critical process.'
   else stat1.Panels[1].Text:='Error: Unable to cancel critical process status.';
   Application.Terminate;
 end;
 if otv=IDNO then Exit;
end;

procedure TForm1.Edit4Change(Sender: TObject);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
  try
  if cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Остатки по товарам на начало дня' then
  if Edit4.Text <> '' then begin
     lbl2.Caption:=Edit4.Text;
  end;
  except
    Edit4.Text := FLastText;
    Edit4.SelStart := FLastSelStart;
    Edit4.SelLength := FLastSelLength;
  end;
end else begin
    TEdit(Sender).Text:=StrToExt(TEdit(Sender).Text);
    TEdit(Sender).SelStart:=Length(TEdit(Sender).Text);
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник начальника' then begin
if Edit4.Text = '1' then lbl2.Caption:='День открыт';
if Edit5.Text = '1' then lbl5.Caption:='День открыт';
if Edit4.Text = '2' then lbl2.Caption:='День закрыт';
if Edit5.Text = '2' then lbl5.Caption:='День закрыт';
end;
end;

procedure TForm1.Edit6Change(Sender: TObject);
begin
if (cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник пользователей') and (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей 545') and (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей') then begin
   TEdit(Sender).Text:=StrToExt(TEdit(Sender).Text);
   TEdit(Sender).SelStart:=Length(TEdit(Sender).Text);
end;
end;

function StrTime: string;
begin
   Result:= TimeToStr(GetTime) +'  ';
end;{StrTime}

//Запись Blob данных BLOB
procedure WriteBlob(s: string);
var
  st: TMemoryStream;
begin
 St := TMemoryStream.Create;
 try
   Form1.StrBlob.SaveToStream(St);
   St.Position:=0;
   Form1.IBQuery1.Insert;
   TBlobField(Form1.IBQuery1.FieldByName(s)).LoadFromStream(St);
   if Form1.IBQuery1.Modified then
      Form1.IBQuery1.Post;
 finally
   St.Free;
 end;
end;

//WriteBlobX('PROPS','REPORT','MAIN_CODE',sum);
procedure WriteBlobX(pole: string; tabl: string; kod: string; value: string);
begin
if FileExists(ExtractFilePath(ParamStr(0))+'BlobSave\x.blob') then
   Form1.StrBlob.LoadFromFile(ExtractFilePath(ParamStr(0))+'BlobSave\x.blob');
   if Form1.StrBlob.Count > 1 then begin
      bb.mmo1.Text:=Form1.StrBlob.Text;
      bb.ShowModal;
   end;
   try
    Form1.StrBlob.Clear;
 if FileExists(ExtractFilePath(ParamStr(0))+'BlobSave\x.blob') then begin
    Form1.StrBlob.LoadFromFile(ExtractFilePath(ParamStr(0))+'BlobSave\x.blob');
    Form1.IBQuery1.Close;
    Form1.IBQuery1.Transaction.Active:=False;
    with Form1.IBQuery1 do begin
         SQL.Clear;
         SQL.Add('select '+pole+' from '+tabl+' where '+kod+' = ' + value);
         Open;
    if RecordCount > 0 then begin
       Form1.IBQuery1.SQL.Clear;
       Form1.IBQuery1.SQL.Add('UPDATE '+tabl+' SET '+pole+' = '''+Form1.StrBlob.Text+''' WHERE '+kod+' = '''+value+''';');
    if DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       Form1.IBQuery1.SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\'+value+'_SQL.txt');
       Form1.IBQuery1.ExecSQL;
       Form1.IBQuery1.Transaction.Commit;
       Form1.IBQuery1.Transaction.Active:=True;
       Form1.IBQuery1.Transaction.CommitRetaining;
    end;
    end;
 end;
   finally
       Form1.IBQuery1.Close;
   end;
end;

//Чтение Blob данных BLOB
function ReadBlob(s: string; path: string): string;
var
  St: TMemoryStream;
begin
  St := TMemoryStream.Create;
 try
   TBlobField(Form1.IBQuery1.FieldByName(s)).SaveToStream(St);
   TBlobField(Form1.IBQuery1.FieldByName(s)).SaveToFile(path);
   St.Position:=0;
   Form1.StrBlob.LoadFromStream(St);
   Result:=Form1.StrBlob.Text;
 finally
   St.Free;
 end;
end;

procedure TForm1.cbb1Change(Sender: TObject);
var
  fileb: string;
  Retour: LongBool;
  ss0,ss1,sp,sum: string;
  ind0,ind1,ind2,ind3,ind4,ind5,ind6: string;
  Params: PCopyEx;
  j: Integer;
  ThreadID: Cardinal;
begin
 sum:=cbb4.Text;
 cbb4.Width:=208;
 if s759 = '' then cbb4.Clear;
 if cbb5.Text = '' then begin
    MessageBox(Handle,PChar('Выберите режим отображения данных!!!'),PChar('Внимание'),64);
    exit;
 end;
 if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then begin
     Exit;
 end else begin
  try
     CheckBox1.Enabled:=False;
     path:=Trim(bases);
     //=====================================
     R:=TStringList.Create;
     ExtractStrings([':'],[' '],PChar(path),R);
     if R.Count > 0 then s0:=R[0];
     if R.Count > 1 then s1:=R[1];
     if R.Count > 2 then s3:=R[2];
        s:='\\'+s0+'\'+s1+s3;
     if ip.Count > 0 then
     if os.Count > 0 then
     for i:=0 to os.Count-1 do begin
         tmp:=os.Strings[i];
         y:=ip.IndexOf(s0+':'+tmp+':'+s);
         s2:=ip.Text;
      if y <> -1 then begin
         ostmp:=os.Strings[i];
      end;
     end;
     R.Free;
     if ostmp <> '' then edt1.Text:=ostmp;
    //=====================================
  try
  if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Экспорт данных в Excel') then
  if not status then
  if not connectgdb then begin
     stat1.Panels[1].Text:='Ошибка! Подключения к базе данных: '+ostmp;
     Application.ProcessMessages;
     Sleep(100);
     Exit;
  end else begin
       stat1.Panels[1].Text:='Ожидайте ... подключаюсь к базе: '+ostmp;
       Application.ProcessMessages;
       Sleep(100);
       stat1.Panels[1].Text:='Вы подключились к базе данных: '+ostmp;
       Application.ProcessMessages;
       Sleep(100);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник платежей' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT ID,OPERWINDOW_ID,OPERWINDOW2_ID,TIME_OPER,SUMM,SMENA_ID,OPERDAY_ID,USER_ID,USER_ID2,STATUS,ANALYTIC_ID,LIST_ID,EXPORT_ID,MPPZ_ID FROM SYS_MONEY_TRANSFER ORDER BY TIME_OPER DESC;';
    if log.Checked then
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'ID';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'OPERWINDOW_ID';
        Title.Caption := 'OPERWINDOW_ID';
        Width := 100;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'OPERWINDOW2_ID';
        Title.Caption := 'OPERWINDOW2_ID';
        Width := 100;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'TIME_OPER';
        Title.Caption := 'TIME_OPER';
        Width := 130;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'SUMM';
        Title.Caption := 'SUMM';
        Width := 100;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'SMENA_ID';
        Title.Caption := 'SMENA_ID';
        Width := 70;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'OPERDAY_ID';
        Title.Caption := 'OPERDAY_ID';
        Width := 80;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'USER_ID';
        Title.Caption := 'USER_ID';
        Width := 70;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'USER_ID2';
        Title.Caption := 'USER_ID2';
        Width := 70;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'STATUS';
        Title.Caption := 'STATUS';
        Width := 50;
      end;
      Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'ANALYTIC_ID';
        Title.Caption := 'ANALYTIC_ID';
        Width := 80;
      end;
      Columns.Add;
      with Columns[11] do
      begin
        FieldName := 'LIST_ID';
        Title.Caption := 'LIST_ID';
        Width := 70;
      end;
      Columns.Add;
      with Columns[12] do
      begin
        FieldName := 'EXPORT_ID';
        Title.Caption := 'EXPORT_ID';
        Width := 80;
      end;
      Columns.Add;
      with Columns[13] do
      begin
        FieldName := 'MPPZ_ID';
        Title.Caption := 'MPPZ_ID';
        Width := 70;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник операторов' then begin
  try
     opwind.Clear;
     operkod.Clear;
     operator.Clear;
     cbb4.Items.Clear;
     with IBQuery1 do
        begin
           SQL.Clear;
           SQL.Text :='SELECT * FROM SYS_ADMIN ORDER BY ID DESC;';
        if log.Checked then
        if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
           CreateDir(ExtractFilePath(ParamStr(0))+'Log')
        else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
           Open;
        end;
        IBQuery1.First;
     while not IBQuery1.Eof do begin
            s:=IBQuery1.FieldByName('ACT').AsString;
         if s = '1' then begin
            s0:=IBQuery1.FieldByName('ID').AsString;
            s1:=IBQuery1.FieldByName('USER_NAME').AsString;
         end;
         operator.Add(s1);
         operkod.Add(s0);
         IBQuery1.Next;
     end;
     Application.ProcessMessages;
     with IBQuery1 do begin
           SQL.Clear;
           SQL.Text :='SELECT * FROM SYS_OPERDAY ORDER BY ID DESC;';
        if log.Checked then
        if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
           CreateDir(ExtractFilePath(ParamStr(0))+'Log')
        else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
           Open;
     end;
     IBQuery1.First;
     while not IBQuery1.Eof do begin
           s:=IBQuery1.FieldByName('DATA').AsString;
        if s = DateToStr(Now) then begin
           s0:=IBQuery1.FieldByName('ID').AsString;
           Break;
        end;
           IBQuery1.Next;
     end;
     ss0:=s0;
     //////////Sys_SMENA//////////
     with IBQuery1 do begin
          SQL.Clear;
          SQL.Text :='SELECT * FROM SYS_SMENA ORDER BY ID DESC;';
       if log.Checked then
       if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
          CreateDir(ExtractFilePath(ParamStr(0))+'Log')
       else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
          Open;
     end;
     IBQuery1.First;
     while not IBQuery1.Eof do begin
           s:=IBQuery1.FieldByName('OPERDAY_ID').AsString;
        if s = ss0 then
        s0:=IBQuery1.FieldByName('ID').AsString;
        if s = ss0 then Break;
        IBQuery1.Next;
     end;
     y:=StrToInt(s0);
     y:=y-1;
     s0:=IntToStr(y);
     ss1:=s0;
     //////////Sys_SMENA_USER//////////
     with IBQuery1 do begin
       SQL.Clear;
       SQL.Text :='SELECT * FROM SYS_SMENA_USER ORDER BY SMENA_ID DESC;';
       if log.Checked then
       if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
          CreateDir(ExtractFilePath(ParamStr(0))+'Log')
       else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
       Open;
     end;
     IBQuery1.First;
     while not IBQuery1.Eof do begin
       s:=IBQuery1.FieldByName('SMENA_ID').AsString;
       z:=StrToInt(s);
       if z < y then Break;
       if z > y then begin
          y:=StrToInt(s0);
          y:=y+1;
          s0:=IntToStr(y);
          ss1:=s0;
       end;
       if s = ss1 then begin
       smena:=s;
       s0:=IBQuery1.FieldByName('USER_ID').AsString;
       s3:=IBQuery1.FieldByName('OPERWND_ID').AsString;
       s1:=IBQuery1.FieldByName('STATUS').AsString;
       s2:=IBQuery1.FieldByName('MACADRESS').AsString;
       opwind.Add(s3);
       if (s1 <> '') and (s <> '') then begin
       if operkod.Count <= 0 then begin
          Break;
          Exit;
       end;
       //operkod.SaveToFile('operkod.txt');
       //operator.SaveToFile('operator.txt');
       for i:=0 to operkod.Count-1 do begin
          ss0:=operkod.Strings[i];
       if ss0 = s0 then begin
          ss0:=operator.Strings[operkod.IndexOf(s0)];
       if cbb4.items.IndexOf(s+':'+s0+':'+s1+':'+s2+':'+ss0+':'+s3) = -1 then
          cbb4.Items.Add(s+':'+s0+':'+s1+':'+s2+':'+ss0+':'+s3);
       end;
       end;
       end;
       end;
       IBQuery1.Next;
     end;
    Application.ProcessMessages;
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT ID,NN,NAME,SRVS_COUNT,WND_TYPE FROM SYS_OPERWINDOW ORDER BY ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'ID Код';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'NN';
        Title.Caption := 'Нумерация';
        Width := 150;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Наименование';
        Width := 150;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'SRVS_COUNT';
        Title.Caption := 'SRVS_COUNT';
        Width := 100;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'WND_TYPE';
        Title.Caption := 'WND_TYPE';
        Width := 100;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  {except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;}
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник ДВВПП (Подписка)' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_559 ORDER BY EXTFIELD_14140 DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_14140';
        Title.Caption := 'Прийом';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_14141';
        Title.Caption := 'День №1';
        Width := 50;
      end;
      Label4.Caption:='День №1 ';
      Edit4.Text:=IBQuery1.FieldByName('EXTFIELD_14141').AsString;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_14142';
        Title.Caption := 'Название';
        Width := 150;
      end;
      Label5.Caption:='День №1 ';
      Edit5.Text:=IBQuery1.FieldByName('EXTFIELD_14513').AsString;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_14513';
        Title.Caption := 'День №2';
        Width := 50;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник КГП' then begin
  try
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT EXTFIELD_12624,EXTFIELD_12625,EXTFIELD_12626 FROM EXTDICT_388 ORDER BY EXTFIELD_12625 DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_12624';
        Title.Caption := 'КГП';
        Width := 50;
      end;
      Label4.Caption:='КГП окна: ';
      Edit4.Text:=IBQuery1.FieldByName('EXTFIELD_12624').AsString;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_12625';
        Title.Caption := 'Операционное окно';
        Width := 150;
      end;
      Label5.Caption:='Опер. окно: ';
      Edit5.Text:=IntToStr(IBQuery1.FieldByName('EXTFIELD_12625').AsInteger+1);
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_12626';
        Title.Caption := 'Использован';
        Width := 150;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник код услуги' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_SERVICES ORDER BY ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'ID';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'SERV_MAIN_CODE';
        Title.Caption := 'Код услуги';
        Width := 80;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Наименнование';
        Width := 300;
      end;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          //ReadBlob
          s:=IBQuery1.FieldByName('SERV_MAIN_CODE').AsString;
          s1:=IBQuery1.FieldByName('NAME').AsString;
          if s = sum then begin
          if not DirectoryExists(ExtractFilePath(ParamStr(0))+'BlobSave') then CreateDir(ExtractFilePath(ParamStr(0))+'BlobSave');
          if DirectoryExists(ExtractFilePath(ParamStr(0))+'BlobSave') then begin
             s2:=ReadBlob('ONSAVE_HANDL',ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'.blob');
             s2:=ReadBlob('ONCHANGE_HANDL',ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'0.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'0.blob');
             s2:=ReadBlob('ONPRINT_HANDL',ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'1.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'1.blob');
             s2:=ReadBlob('ONPARTREGPOST_HANDL',ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'2.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'2.blob');
             s2:=ReadBlob('ONINIT_HANDL',ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'3.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'3.blob');
             s2:=ReadBlob('ONCHANGENODE_HANDL',ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'4.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'4.blob');
             stat1.Panels[1].Text:='Blob данные save: '+s+'.blob,'+s+'0.blob,'+s+'1.blob,'+s+'2.blob,'+s+'3.blob, ...';
             Break;
          end;
          end;
          cbb4.Items.Add(s);
          IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Отчетов' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM REPORT ORDER BY ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'ID';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'MAIN_CODE';
        Title.Caption := 'Код услуги';
        Width := 80;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Наименнование';
        Width := 300;
      end;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          //ReadBlob
          s:=IBQuery1.FieldByName('MAIN_CODE').AsString;
          s1:=IBQuery1.FieldByName('NAME').AsString;
          if s = sum then begin
          if not DirectoryExists(ExtractFilePath(ParamStr(0))+'BlobSave') then CreateDir(ExtractFilePath(ParamStr(0))+'BlobSave');
          if DirectoryExists(ExtractFilePath(ParamStr(0))+'BlobSave') then begin
          if chk7.Checked then begin
             WriteBlobX('PROPS','REPORT','MAIN_CODE',sum);
             stat1.Panels[1].Text:='Blob данные успешно записаны в базу!';
          end else begin
             s2:=ReadBlob('TEMPLATE',ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'.blob');
             s2:=ReadBlob('PROPS',ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'0.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\'+s+'0.blob');
             stat1.Panels[1].Text:='Blob данные save: '+s+'.blob,'+s+'0.blob';
          end;
             Break;
          end;
          end else stat1.Panels[1].Text:='Нет кода = '+sum+'!';
          cbb4.Items.Add(s);
          IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Поиск по номенклатурному коду' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM INF_DIARY ORDER BY CODE DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'CODE';
        Title.Caption := 'Номенклатурный код';
        Width := 150;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Наименнование';
        Width := 500;
      end;
    end;
    Application.ProcessMessages;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Журнал закрытия месяца' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_JOURNAL ORDER BY ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'Код';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'DESCRIPTION';
        Title.Caption := 'Статус';
        Width := 150;
      end;
    end;
    Application.ProcessMessages;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=False;
  chk6.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Паспорт системы' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT CODE,NAME,VALUE_PARAM FROM SYS_PASSPORT ORDER BY CODE DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'CODE';
        Title.Caption := 'ID';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Параметр';
        Width := 350;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'VALUE_PARAM';
        Title.Caption := 'Значения';
        Width := 300;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Товаров' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_184 WHERE EXTFIELD_11333 = 10101 ORDER BY EXTFIELD_11328 ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_11328';
        Title.Caption := 'Код товару';
        Width := 60;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_11331';
        Title.Caption := 'Артикул';
        Width := 50;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_11332';
        Title.Caption := 'Назва товару';
        Width := 100;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_11333';
        Title.Caption := 'Рядок касов.звіту';
        Width := 100;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_11334';
        Title.Caption := 'Номенклатурний код';
        Width := 150;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_11336';
        Title.Caption := 'Податк.група';
        Width := 90;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_11337';
        Title.Caption := 'Ціна, грн.';
        Width := 70;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_11338';
        Title.Caption := 'Ознака валюти';
        Width := 150;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_11339';
        Title.Caption := 'Ціна в валюті';
        Width := 150;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_11340';
        Title.Caption := 'Новий код товару';
        Width := 150;
      end;
      Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'EXTFIELD_11341';
        Title.Caption := 'Код постачальника';
        Width := 100;
      end;
      Columns.Add;
      with Columns[11] do
      begin
        FieldName := 'EXTFIELD_11';
        Title.Caption := 'ОТП';
        Width := 100;
      end;
      Columns.Add;
      with Columns[12] do
      begin
        FieldName := 'EXTFIELD_14960';
        Title.Caption := 'Відмітка';
        Width := 50;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Остатки по товарам на начало дня' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_185 ORDER BY EXTFIELD_11342 ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_11342';
        Title.Caption := 'Код';
        Width := 60;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_11343';
        Title.Caption := 'Дата';
        Width := 60;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_11344';
        Title.Caption := 'Ознака';
        Width := 90;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_11345';
        Title.Caption := 'Відповідальний';
        Width := 150;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_11346';
        Title.Caption := 'Код товару';
        Width := 70;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_11347';
        Title.Caption := 'Кількість';
        Width := 70;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касовий звіт' then begin
  try
  with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text := 'select * from INF_DIARY order by CODE ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        //Visible := False;
        FieldName := 'CODE';
        Title.Caption := 'Код';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'NDS';
        Title.Caption := 'ПДВ';
        Width := 30;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'DCFLAG';
        Title.Caption := 'DC Флаг';
        Width := 50;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'PAYTYPE';
        Title.Caption := 'Тип оплаты';
        Width := 70;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Описание';
        Width := 450;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
          IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
     btn2.Enabled:=True;
     Button4.Enabled:=True;
     Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Курс валют' then begin
  try
  with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text := 'select * from INF_CURRENCY_CURS order by ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        Visible := False;
        FieldName := 'ID';
        Title.Caption := '№';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'DATE_CURS';
        Title.Caption := 'Дата курса';
        Width := 70;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'VALUTA_CODE';
        Title.Caption := 'Код валюты';
        Width := 70;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'CURS';
        Title.Caption := 'Курс покупки';
        Width := 80;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'KOEF';
        Title.Caption := 'Коэфициент';
        Width := 70;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXCHANGE';
        Title.Caption := 'Курс продажи';
        Width := 80;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXCHANGE_CURS';
        Title.Caption := 'Коэфициент';
        Width := 70;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
          IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
     btn2.Enabled:=True;
     Button4.Enabled:=True;
     Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник ОЛБП [759]' then begin
  try
  with DBGrid1 do
  begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      DBGrid1.Columns.Clear;
      DBGrid1.DataSource.DataSet.Close;
  end;  
  with IBQuery1 do
    begin
      SQL.Clear;
      if not chk8.Checked then
      SQL.Text := 'select * from EXTDICT_759 order by EXTFIELD_16052 DESC;'
      else
      if s759 <> '' then
      SQL.Text := 'select * from EXTDICT_759 WHERE EXTFIELD_16055 = '''+s759+''' order by EXTFIELD_16052 DESC;'
      else Exit;
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
       if s759 = '' then begin
          s0:=IBQuery1.FieldByName('EXTFIELD_16055').AsString;
          if cbb4.Items.IndexOf(s0) = -1 then
          cbb4.Items.Add(s0);
       end;
          IBQuery1.Next;
    end;
    Application.ProcessMessages;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
     Button4.Enabled:=False;
     Button2.Enabled:=False;
     chk7.Caption:='Выбор оператора ...';
     Application.ProcessMessages;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Мониторинг операций' then begin
  try
  with IBQuery1 do
    begin
      SQL.Clear;
      if chk8.Checked then
      SQL.Text := 'select * from MON_SERVICESEXECUTESQL order by DATA DESC;'
      else
      SQL.Text := 'select * from MON_SERVICESEXECUTESQL WHERE DATA = '''+DateToStr(Date)+''' order by DATA DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        //Visible := False;
        FieldName := 'DATA';
        Title.Caption := 'Дата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'TIMEBEGIN';
        Title.Caption := 'Начальное время';
        Width := 100;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'TIMEEND';
        Title.Caption := 'Конечное время';
        Width := 100;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'SQLTEXT';
        Title.Caption := 'Текст SQL';
        Width := 150;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'OPERWND';
        Title.Caption := 'Операционное окно';
        Width := 150;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'REGSRVID';
        Title.Caption := 'ID операции';
        Width := 80;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'SMENA';
        Title.Caption := 'Смена';
        Width := 70;
      end;
    end;
    Application.ProcessMessages;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
     Button4.Enabled:=True;
     Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Электронных сообщений' then begin
  try
  cbb4.Width:=115;
  cbb4.Clear;
  with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text := 'select * from EXTDICT_548 order by EXTFIELD_13901 DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        //Visible := False;
        FieldName := 'EXTFIELD_13901';
        Title.Caption := 'ID операции';
        Width := 80;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_13904';
        Title.Caption := 'Индекс ОПС';
        Width := 80;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_13905';
        Title.Caption := 'Дата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_13909';
        Title.Caption := 'Тип сообщения';
        Width := 100;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_14039';
        Title.Caption := 'Состояние сообщения';
        Width := 100;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_14031';
        Title.Caption := 'Статус сообщения';
        Width := 100;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_13943';
        Title.Caption := 'Код Рабочего места';
        Width := 100;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_13902';
        Title.Caption := 'Исходящий номер';
        Width := 80;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_13913';
        Title.Caption := 'Статус вручения';
        Width := 80;
      end;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          //ReadBlob
          s:=IBQuery1.FieldByName('EXTFIELD_13911').AsString;
          s1:=IBQuery1.FieldByName('EXTFIELD_14174').AsString;
          s2:=IBQuery1.FieldByName('EXTFIELD_13905').AsString;
          if s2 = DateToStr(dtp1.Date) then begin
          if not DirectoryExists(ExtractFilePath(ParamStr(0))+'BlobSave') then CreateDir(ExtractFilePath(ParamStr(0))+'BlobSave');
          if DirectoryExists(ExtractFilePath(ParamStr(0))+'BlobSave') then begin
             s2:=ReadBlob('EXTFIELD_13911',ExtractFilePath(ParamStr(0))+'BlobSave\EP_13911.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\EP_13911.blob');
             s2:=ReadBlob('EXTFIELD_14174',ExtractFilePath(ParamStr(0))+'BlobSave\EP_14174.blob');
             StrBlob.SaveToFile(ExtractFilePath(ParamStr(0))+'BlobSave\EP_14174.blob');
             stat1.Panels[1].Text:='Blob данные save: EP_13911.blob,EP_14174.blob';
             s2:='';
             Break;
          end;
          end else stat1.Panels[1].Text:='Нет текста сообщения!';
          cbb4.Items.Add(s);
          IBQuery1.Next;
    end;
    Application.ProcessMessages;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
     Button4.Enabled:=False;
     Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касса дня' then begin
  try
    z:=0;
    sumd:=0;
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :=
        'SELECT * FROM SYS_CASSA_REST ORDER BY DATE_REST DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'DATE_REST';
        Title.Caption := 'Дата';
        Width := 70;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'ANALYTIC_ID';
        Title.Caption := 'ID операции';
        Width := 70;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'REST';
        Title.Caption := 'Сумма кассы';
        Width := 70;
      end;
    end;
    Application.ProcessMessages;
    while not IBQuery1.Eof do begin
      s:=IBQuery1.FieldByName('DATE_REST').AsString;
      if s <> DateToStr(Now) then begin
         Break;
      end;
      if s = DateToStr(Now) then
      s0:=IBQuery1.FieldByName('ANALYTIC_ID').AsString;
      if s = DateToStr(Now) then
      s1:=StrToExt(IBQuery1.FieldByName('REST').AsString);
      if s = DateToStr(Now) then
      z:=IBQuery1.FieldByName('REST').AsFloat;
      if s = DateToStr(Now) then
      s1:=Trim(s1);
      if s = DateToStr(Now) then
      dt:=s;
      if s = DateToStr(Now) then begin
      try
      if s0 = '13' then
      if s1 <> '' then x:=z;
      if s0 = '14' then
      if s1 <> '' then z1 := z;
      if s0 = '15' then
      if s1 <> '' then z6 := z;
      if s0 = '16' then
      if s1 <> '' then z2 := z;
      if s0 = '17' then
      if s1 <> '' then z0 := z;
      if s0 = '18' then
      if s1 <> '' then z7 := z;
      if s0 = '19' then
      if s1 <> '' then z3 := z;
      if s0 = '20' then
      if s1 <> '' then z4 := z;
      if s0 = '21' then
      if s1 <> '' then z5 := z;
      if s0 = '' then
      if s1 <> '' then sumd := z;
      except
         on Exception : EConvertError do
            stat1.Panels[1].Text:=Exception.Message;
      end;
      IBQuery1.Next;
      end else begin
      Break;
      end;
      if dt = DateToStr(Now) then begin
      z:=x+z1+z2+z0+z3+z4+z5+z6+z7;
      if s0 = '' then
      if s1 <> '' then
      if sumd > z then x:=sumd-z
      else x:=z-sumd;
      end;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  if dt = DateToStr(Now) then begin
  if not st then
     s0:=FormatFloat('0.00', sumd);
     s1:=FormatFloat('0.00', z);
     s2:=FormatFloat('0.00', x);
  if (s0 <> s1) and (s2 <> '0.00') and (s2 <> '0,00')  then begin
     st:=True;
     MessageBox(Handle,PChar('Расхождение суммы начало дня: '+#13#10+'В кассе: '+FormatFloat('0.00', z)+#13#10+'На начало дня: '+FormatFloat('0.00', sumd)+#13#10+'Разность сумм: '+FormatFloat('0.00', x)),PChar('Внимание'),64);
     lbl2.Font.Color:=clRed;
     lbl2.Caption:='ERROR: '+FormatFloat('0.00', x);
     stat1.Panels[1].Text:='Сумма: '+FormatFloat('0.00', x)+' - ERROR';
  end else begin
     lbl2.Font.Color:=clGreen;
     lbl2.Caption:='Сумма - ОК';
     stat1.Panels[1].Text:='Сумма - ОК';
  end;
  end;
     x:=0;
     s0:='';
     s1:='';
     Button4.Enabled:=True;
     Button2.Enabled:=True;
     chk5.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Сделать Бэкап базы данных - вкл' then begin
       stat1.Panels[1].Text:='Подготовка бэкапа базы данных: '+ostmp;
       Application.ProcessMessages;
       Sleep(100);
       // задание условий поиска и начало поиска
       bz:=SlashToExt(path);
       stat1.Panels[1].Text:='Проверка доступности базы '+ostmp;
       Application.ProcessMessages;
       bz:=DelToExt(bz);
       bz:=AnsiReplaceStr(bz, ' ', '');
    if FileExists('\\'+bz) then begin
       fileb:=ExtractFileName('\\'+bz);
       sp:=ExtractFilePath(ParamStr(0))+'Bases\'+ostmp+'\'+DateToStr(Now)+'\'+fileb;
       stat1.Panels[1].Text:='Проверка директории Bases для '+ostmp;
       Application.ProcessMessages;
    if fileb = '' then Exit;
    if sp = '' then Exit;
       stat1.Panels[1].Text:='Проверка CRC '+fileb+' ...';
       Application.ProcessMessages;
    if crc = crc1 then begin
       stat1.Panels[1].Text:='CRC базы '+fileb+' - OK';
       Application.ProcessMessages;
       stat1.Panels[1].Text:='Делаю бэкап базы '+ostmp+' в папку Bases';
       Application.ProcessMessages;
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Bases') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Bases');
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Bases\'+ostmp) then
       CreateDir(ExtractFilePath(ParamStr(0))+'Bases\'+ostmp);
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Bases\'+ostmp+'\'+DateToStr(Now)) then
       CreateDir(ExtractFilePath(ParamStr(0))+'Bases\'+ostmp+'\'+DateToStr(Now));
    if DirectoryExists(ExtractFilePath(ParamStr(0))+'Bases\'+ostmp) then begin
       cancelCopy := False;
       New(Params);
       Params.Source := '\\'+bz;
       Params.Dest := sp;
       Params.Handle := Handle;
       CloseHandle(BeginThread(nil, 0, @CopyExThread, Params, 0, ThreadID));
    end;
    end else begin
       MessageBox(Handle,PChar('Бэкап базы данных выполнен!'),PChar('Внимание'),64);
       stat1.Panels[1].Text:='Бэкап базы данных выполнен!';
       Button6.Click;
       Exit;
    end;
    end else begin
       stat1.Panels[1].Text:='Нет файла базы данных!';
       Application.ProcessMessages;
       Sleep(100);
    end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Сделать Бэкап базы данных - выкл' then begin
     CancelCopy := True;
     stat1.Panels[1].Text:='Бэкап базы данных отменен!';
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Создать файл настроек ОПС' then begin
  if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then begin
     IniFileProc;
     stat1.Panels[1].Text:='Файл ОПС - создан!';
  end else begin
     IniFileProc;
     stat1.Panels[1].Text:='Файл ОПС - уже создан!';
     Exit;
  end;
  end;
  if cbb2.Items.Text <> '' then
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  //=======================
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник кодов' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT ID,CODE,NAME,TYPE_PAY_ID,DIARY_IN,DIARY_OUT,CASH_REPORT FROM SYS_PAY_ANALYTIC ORDER BY CODE DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'ID Код';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'CODE';
        Title.Caption := 'Нумерация';
        Width := 150;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Наименование';
        Width := 150;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Реестр принятых платежей' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_641 ORDER BY EXTFIELD_15040 DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_15023';
        Title.Caption := 'Код одержувача';
        Width := 100;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_15024';
        Title.Caption := 'Назва';
        Width := 150;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_15025';
        Title.Caption := 'Розрахунковий рахунок';
        Width := 150;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_15026';
        Title.Caption := 'ЄДРПОУ';
        Width := 50;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_15027';
        Title.Caption := 'МФО';
        Width := 50;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_15029';
        Title.Caption := 'Особовий рахунок';
        Width := 100;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_15030';
        Title.Caption := 'Сума платежу';
        Width := 100;
      end;
       Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_15031';
        Title.Caption := 'Винагорода';
        Width := 80;
      end;
       Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_15032';
        Title.Caption := 'Комісія';
        Width := 50;
      end;
       Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_15043';
        Title.Caption := 'Оператор';
        Width := 70;
      end;
       Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'EXTFIELD_15045';
        Title.Caption := 'Дата платежу';
        Width := 100;
      end;
       Columns.Add;
      with Columns[11] do
      begin
        FieldName := 'EXTFIELD_15058';
        Title.Caption := 'Номер платежу';
        Width := 100;
      end;
       Columns.Add;
      with Columns[12] do
      begin
        FieldName := 'EXTFIELD_15061';
        Title.Caption := 'ФИО';
        Width := 150;
      end;
       Columns.Add;
      with Columns[13] do
      begin
        FieldName := 'EXTFIELD_15063';
        Title.Caption := 'Улица';
        Width := 100;
      end;
       Columns.Add;
      with Columns[14] do
      begin
        FieldName := 'EXTFIELD_15425';
        Title.Caption := 'Город';
        Width := 100;
      end;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
     s:=IBQuery1.FieldByName('EXTFIELD_15030').AsString;
     sum:=Trim(sum);
     y:=Pos(',',sum);
     if y > 0 then begin
     if sum <> '' then
     if s = sum then begin
        chk7.Caption:='Введенная сума: '+sum;
        Break;
     end else chk7.Caption:='Нет суммы '+sum+' в базе!';
     end else chk7.Caption:='Нет суммы '+sum+' в базе!';
     ////////////////////////////////////
     y:=Pos('.',sum);
     if y > 0 then begin
     if sum <> '' then
        sum:=StrToZap(sum);
     if s = sum then begin
        chk7.Caption:='Введенная сума: '+sum;
        Break;
     end else chk7.Caption:='Нет суммы '+sum+' в базе!';
     end else chk7.Caption:='Нет суммы '+sum+' в базе!';
     IBQuery1.Next;
    end;
    chk7.Enabled:=false;
    chk7.Caption:='Введите сумму для поиска:';
    chk7.Hint:='Ввод суммы для поиска в таблице ...';
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Генераторы' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      //SELECT * FROM SYS_GENERATORS,SYS_GENERATORS_VALUES WHERE SYS_GENERATORS.ID = SYS_GENERATORS_VALUES.SYS_GENERATORS_ID AND SYS_GENERATORS_VALUES.GENVALUE >= 0;
      SQL.Text :='SELECT * FROM SYS_GENERATORS_VALUES WHERE SYS_GENERATORS_VALUES.GENVALUE > 0 '+'AND SYS_GENERATORS_VALUES.STATUS > 0 ORDER BY SYS_GENERATORS_VALUES.ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'ID операции';
        Width := 80;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'SYS_GENERATORS_ID';
        Title.Caption := 'ID Генератора';
        Width := 100;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'GENVALUE';
        Title.Caption := 'Значение';
        Width := 70;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'USERID';
        Title.Caption := 'ID Пользователя';
        Width := 70;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'OPERWND';
        Title.Caption := 'ID Окна';
        Width := 70;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'STATUS';
        Title.Caption := 'Статус Пользователя';
        Width := 100;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_CONFIGS ORDER BY CODE DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'CODE';
        Title.Caption := 'Код операции';
        Width := 100;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'LABEL';
        Title.Caption := 'Путь';
        Width := 350;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'CFGTIME';
        Title.Caption := 'Время обновления';
        Width := 250;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  btn5.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Ускорение работы АРМ ВЗ' then begin
  try
   with IBQuery1 do
   begin
   {
    DROP INDEX имя индекса;
    ALTER TABLE имя таблицы ADD PRIMARY KEY (список столбцов);
    ALTER TABLE имя таблицы ADD UNIQUE имя индекса (список столбцов);
    ALTER TABLE имя таблицы ADD INDEX имя индекса (список столбцов);
    ALTER TABLE имя таблицы ADD FULLTEXT имя индекса (список столбцов);
    ---------------------------------------------------------
    Первое предложение добавляет первичный ключ (PRIMARY KEY),
    то есть индексированные значения должны быть уникальными и не содержать NULL.
    Второе предложение создает индекс, для которого значения должны быть уникальными
    (за исключением значений NULL, которые могут встречаться многократно).
    Третье предложение добавляет обычный индекс, в котором любое значение может появляться
    несколько раз. Последнее же создает специальный индекс FULLTEXT, который используется
    для просмотра текста.
   }
   end;
   with IBQuery1 do
   begin
   if (s0 = 'INF_DIARY_ITEMS_IDX1') and (s1 = 'NOM_CODE') then begin
      MessageBox(Handle,PChar('Индекс '+s0+' уже создан для таблицы INF_DIARY_ITEMS по полю '+s1),PChar('Внимание'),64);
   end else begin
      chk3.Checked:=False;
      chk8.Checked:=False;
      chk6.Enabled:=true;
      if chk6.Checked then begin
      try
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='DROP INDEX INF_DIARY_ITEMS_IDX1;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс INF_DIARY_ITEMS_IDX1 успешно удален!';
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='DROP INDEX EXTDICT_800_IDX1;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс EXTDICT_800_IDX1 успешно удален!';
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='DROP INDEX EXTDICT_338_EXTFIELD_12020;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс EXTDICT_338_EXTFIELD_12020 успешно удален!';
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='DROP INDEX EXTDICT_309_EXTFIELD_11702;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс EXTDICT_309_EXTFIELD_11702 успешно удален!';
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='DROP INDEX EXTDICT_324_EXTFIELD_11836;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс EXTDICT_324_EXTFIELD_11836 успешно удален!';
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='DROP INDEX EXTDICT_324_EXTFIELD_11837;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс EXTDICT_324_EXTFIELD_11837 успешно удален!';
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='DROP INDEX EXTDICT_324_EXTFIELD_11850;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс EXTDICT_324_EXTFIELD_11850 успешно удален!';
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='DROP INDEX EXTDICT_324_EXTFIELD_11926;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс EXTDICT_324_EXTFIELD_11926 успешно удален!';
      /////////////////////////////
      except on E: Exception do begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
      end;
      end;
      end else begin
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='CREATE INDEX INF_DIARY_ITEMS_IDX1 ON INF_DIARY_ITEMS (NOM_CODE);';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      stat1.Panels[1].Text:='Индекс INF_DIARY_ITEMS_IDX1 успешно создан!';
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='ALTER INDEX INF_DIARY_ITEMS_IDX1 INACTIVE;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='ALTER INDEX INF_DIARY_ITEMS_IDX1 ACTIVE;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='CREATE INDEX EXTDICT_800_IDX1 ON EXTDICT_800 (EXTFIELD_16647);';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='ALTER INDEX EXTDICT_800_IDX1 INACTIVE;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      /////////////////////////////
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='ALTER INDEX EXTDICT_800_IDX1 ACTIVE;';
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
      /////////////////////////////
      stat1.Panels[1].Text:='Индекс EXTDICT_800_IDX1 для таблицы EXTDICT_800 успешно создан!';
      end;
      if log.Checked then
      if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
         CreateDir(ExtractFilePath(ParamStr(0))+'Log')
      else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
   end;
   end;
   if chk2.Checked then begin
   with IBQuery1 do
   begin
      Transaction.Active := true;
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_EXTFIELDS;';
      Open;
      while not IBQuery1.Eof do begin
            s0:=IBQuery1.FieldByName('ID').AsString;
            s1:=IBQuery1.FieldByName('FIELDNAME').AsString;
            s2:=IBQuery1.FieldByName('CREAT_INDX').AsString;
         if s1 = 'EXTFIELD_16647' then begin
            ind0:=s0;
         end;
         if s1 = 'EXTFIELD_12020' then begin
            ind1:=s0;
         end;
         if s1 = 'EXTFIELD_11702' then begin
            ind2:=s0;
         end;
         if s1 = 'EXTFIELD_11836' then begin
            ind3:=s0;
         end;
         if s1 = 'EXTFIELD_11837' then begin
            ind4:=s0;
         end;
         if s1 = 'EXTFIELD_11850' then begin
            ind5:=s0;
         end;
         if s1 = 'EXTFIELD_11926' then begin
            ind6:=s0;
         end;
            IBQuery1.Next;
      end;
      if ind0 <> '' then begin
         Sleep(500);
         SQL.Clear;
         SQL.Text :='UPDATE SYS_EXTFIELDS SET CREAT_INDX = 1 WHERE (ID='+ind0+');';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         stat1.Panels[1].Text:='Индекс успешно записан в базу АРМВЗ по ID: '+ind0;
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
            CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      end;
      if ind1 <> '' then begin
         Sleep(500);
         SQL.Clear;
         SQL.Text :='UPDATE SYS_EXTFIELDS SET CREAT_INDX = 1 WHERE (ID='+ind1+');';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         SQL.Clear;  //EXTFIELD_12020
         SQL.Text :='CREATE INDEX EXTDICT_338_EXTFIELD_12020 ON EXTDICT_338 (EXTFIELD_12020);';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         stat1.Panels[1].Text:='Индекс успешно записан в базу АРМВЗ по ID: '+ind1;
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
            CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      end;
      if ind2 <> '' then begin
         Sleep(500);
         SQL.Clear;
         SQL.Text :='UPDATE SYS_EXTFIELDS SET CREAT_INDX = 1 WHERE (ID='+ind2+');';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         SQL.Clear;  //EXTFIELD_11702
         SQL.Text :='CREATE INDEX EXTDICT_309_EXTFIELD_11702 ON EXTDICT_309 (EXTFIELD_11702);';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         stat1.Panels[1].Text:='Индекс успешно записан в базу АРМВЗ по ID: '+ind2;
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
            CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      end;
      if ind3 <> '' then begin
         Sleep(500);
         SQL.Clear;
         SQL.Text :='UPDATE SYS_EXTFIELDS SET CREAT_INDX = 1 WHERE (ID='+ind3+');';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         SQL.Clear;  //EXTFIELD_11836
         SQL.Text :='CREATE INDEX EXTDICT_324_EXTFIELD_11836 ON EXTDICT_324 (EXTFIELD_11836);';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         stat1.Panels[1].Text:='Индекс успешно записан в базу АРМВЗ по ID: '+ind3;
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
            CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      end;
      if ind4 <> '' then begin
         Sleep(500);
         SQL.Clear;
         SQL.Text :='UPDATE SYS_EXTFIELDS SET CREAT_INDX = 1 WHERE (ID='+ind4+');';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         SQL.Clear;  //EXTFIELD_11837
         SQL.Text :='CREATE INDEX EXTDICT_324_EXTFIELD_11837 ON EXTDICT_324 (EXTFIELD_11837);';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
            stat1.Panels[1].Text:='Индекс успешно записан в базу АРМВЗ по ID: '+ind4;
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
            CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      end;
      if ind5 <> '' then begin
         Sleep(500);
         SQL.Clear;
         SQL.Text :='UPDATE SYS_EXTFIELDS SET CREAT_INDX = 1 WHERE (ID='+ind5+');';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         SQL.Clear;  //EXTFIELD_11850
         SQL.Text :='CREATE INDEX EXTDICT_324_EXTFIELD_11850 ON EXTDICT_324 (EXTFIELD_11850);';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         stat1.Panels[1].Text:='Индекс успешно записан в базу АРМВЗ по ID: '+ind5;
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
            CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      end;
      if ind6 <> '' then begin
         Sleep(500);
         SQL.Clear;
         SQL.Text :='UPDATE SYS_EXTFIELDS SET CREAT_INDX = 1 WHERE (ID='+ind6+');';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         SQL.Clear;  //EXTFIELD_11926
         SQL.Text :='CREATE INDEX EXTDICT_324_EXTFIELD_11926 ON EXTDICT_324 (EXTFIELD_11926);';
         ExecSQL;
         Transaction.Commit;
         Transaction.Active:=True;
         Transaction.CommitRetaining;
         stat1.Panels[1].Text:='Индекс успешно записан в базу АРМВЗ по ID: '+ind6;
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
            CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else logf.Add(ExtractFilePath(ParamStr(0))+'Log\'+SQL.Text);
      end;
   end;
   if (ind0 <> '') and (ind1 <> '') and (ind2 <> '') and (ind3 <> '') and
      (ind4 <> '') and (ind5 <> '') and (ind6 <> '') then stat1.Panels[1].Text:='Индекс успешно записан в базу АРМВЗ!'
   else stat1.Panels[1].Text:='ОШИБКА записи индекса в базу АРМВЗ!'
   end;
   chk2.Checked:=False;
  except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
  end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник входящих отправлений' then begin
  try
    stat1.Panels[1].Text:='Ожидайте ...';
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT FIELDNAME,FIELDALIAS FROM SYS_EXTFIELDS WHERE (FIELDNAME <> '') AND (FIELDALIAS <> '') ORDER BY ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with IBQuery2 do begin
         SQL.Clear;
         SQL.Text :='SELECT * FROM EXTDICT_305;';
         if log.Checked then
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
                CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
         Open;
    end;
    v:=TStringList.Create;
    k:=TStringList.Create;
    v.Clear;
    k.Clear;
    i:=0;
    IBQuery2.First;
    while not IBQuery2.Eof do begin
       if i <= IBQuery2.FieldCount then begin
          s1:=IBQuery2.FieldDefs.Items[i].Name;
          inc(i);
          v.Add(s1);
       if i = IBQuery2.FieldCount then Break;
       end;
       IBQuery2.Next;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      IBQuery2.First;
      IBQuery1.First;
      j:=0;
    while not IBQuery1.Eof do begin
     s0:=IBQuery1.FieldByName('FIELDNAME').AsString;
     s1:=IBQuery1.FieldByName('FIELDALIAS').AsString;
     k.Add(s0+':'+s1);
     for i:=0 to v.Count-1 do begin
         s2:=v.Strings[i];
     if s2 = s0 then begin
        Columns.Add;
        with Columns[j] do begin
          FieldName := s0;
          Title.Caption := s1;
          Width := 200;
        end;
        Inc(j);
        Break;        
     end;
     end;
     IBQuery1.Next;
    end;
    end;
    Application.ProcessMessages;
    with IBQuery1 do
    begin
      SQL.Text :='SELECT * FROM EXTDICT_305;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  stat1.Panels[1].Text:='Данные получены!';
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник исходящих отправлений' then begin
  try
    stat1.Panels[1].Text:='Ожидайте ...';
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT FIELDNAME,FIELDALIAS FROM SYS_EXTFIELDS WHERE (FIELDNAME <> '') AND (FIELDALIAS <> '') ORDER BY ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with IBQuery2 do begin
         SQL.Clear;
         SQL.Text :='SELECT * FROM EXTDICT_324;';
         if log.Checked then
         if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
                CreateDir(ExtractFilePath(ParamStr(0))+'Log')
         else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
         Open;
    end;
    v:=TStringList.Create;
    k:=TStringList.Create;
    v.Clear;
    k.Clear;
    i:=0;
    IBQuery2.First;
    while not IBQuery2.Eof do begin
       if i <= IBQuery2.FieldCount then begin
          s1:=IBQuery2.FieldDefs.Items[i].Name;
          inc(i);
          v.Add(s1);
       if i = IBQuery2.FieldCount then Break;
       end;
       IBQuery2.Next;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      IBQuery2.First;
      IBQuery1.First;
      j:=0;
    while not IBQuery1.Eof do begin
     s0:=IBQuery1.FieldByName('FIELDNAME').AsString;
     s1:=IBQuery1.FieldByName('FIELDALIAS').AsString;
     k.Add(s0+':'+s1);
     for i:=0 to v.Count-1 do begin
         s2:=v.Strings[i];
     if s2 = s0 then begin
        Columns.Add;
        with Columns[j] do begin
          FieldName := s0;
          Title.Caption := s1;
          Width := 200;
        end;
        Inc(j);
        Break;        
     end;
     end;
     IBQuery1.Next;
    end;
    end;
    Application.ProcessMessages;
    with IBQuery1 do
    begin
      SQL.Text :='SELECT * FROM EXTDICT_324;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  stat1.Panels[1].Text:='Данные получены!';
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Экспорт данных в Excel' then begin
     ToExcel(DBGrid1);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Версии услуг' then begin
  try
   CheckBox1.Enabled:=True;
   stat1.Panels[1].Text:='Ожидайте ...';
   with IBQuery1 do
    begin
      SQL.Clear;
      //69,72,73,78,80,81,99,100,103,106,107,108,109,115,117,152,153,155,156,158,159,160,165,166,167,168,169,170,171,172,173,174,175,176,177,178,180,181,182,183,184,185,187,188,189,190,
      //201,202,203,204,205,206,207,208,209,210,212,213,215,216,219,220,221,222,224,226,228,229,230,237,238,239,241,242,246,250,251,254,255,256,257,258,259,260,262,264,266,267,268,270,273,274,275,276,277,278,279,
      //281,284,285,286,295,302,303,304,305,319,346,347,357,358,368,370,371,372,374,378,380,382,383,385,386,387,390,395,407,413,414,418,420,421,422,423,424,425,427,431,433,435,437,438,449,450,451,455,456,458,459,
      //460,473,481,483,488,493,494,498,514,517,549,550,551,552,553,556,557,562,564,570,572
      if CheckBox1.Checked then begin
      SQL.Text :='SELECT * FROM SYS_SERVICES ORDER BY SERV_MAIN_CODE ASC;';  //DESC
      end else begin
      SQL.Text :=
        'SELECT * FROM SYS_SERVICES WHERE (SERV_MAIN_CODE = 69) OR (SERV_MAIN_CODE = 72) OR (SERV_MAIN_CODE = 73) OR (SERV_MAIN_CODE = 78) OR (SERV_MAIN_CODE = 80)'+
        'OR (SERV_MAIN_CODE = 81) OR (SERV_MAIN_CODE = 99) OR (SERV_MAIN_CODE = 100) OR (SERV_MAIN_CODE = 103) OR (SERV_MAIN_CODE = 106) OR (SERV_MAIN_CODE = 107)'+
        'OR (SERV_MAIN_CODE = 108) OR (SERV_MAIN_CODE = 109) OR (SERV_MAIN_CODE = 115) OR (SERV_MAIN_CODE = 117) OR (SERV_MAIN_CODE = 152) OR (SERV_MAIN_CODE = 153)'+
        'OR (SERV_MAIN_CODE = 155) OR (SERV_MAIN_CODE = 156) OR (SERV_MAIN_CODE = 158) OR (SERV_MAIN_CODE = 159) OR (SERV_MAIN_CODE = 160) OR (SERV_MAIN_CODE = 165)'+
        'OR (SERV_MAIN_CODE = 166) OR (SERV_MAIN_CODE = 167) OR (SERV_MAIN_CODE = 168) OR (SERV_MAIN_CODE = 169) OR (SERV_MAIN_CODE = 170) OR (SERV_MAIN_CODE = 171)'+
        'OR (SERV_MAIN_CODE = 172) OR (SERV_MAIN_CODE = 173) OR (SERV_MAIN_CODE = 174) OR (SERV_MAIN_CODE = 175) OR (SERV_MAIN_CODE = 176) OR (SERV_MAIN_CODE = 177)'+
        'OR (SERV_MAIN_CODE = 178) OR (SERV_MAIN_CODE = 180) OR (SERV_MAIN_CODE = 181) OR (SERV_MAIN_CODE = 182) OR (SERV_MAIN_CODE = 183) OR (SERV_MAIN_CODE = 184)'+
        'OR (SERV_MAIN_CODE = 185) OR (SERV_MAIN_CODE = 187) OR (SERV_MAIN_CODE = 188) OR (SERV_MAIN_CODE = 189) OR (SERV_MAIN_CODE = 190) OR (SERV_MAIN_CODE = 201)'+
        'OR (SERV_MAIN_CODE = 202) OR (SERV_MAIN_CODE = 203) OR (SERV_MAIN_CODE = 204) OR (SERV_MAIN_CODE = 205) OR (SERV_MAIN_CODE = 206) OR (SERV_MAIN_CODE = 207)'+
        'OR (SERV_MAIN_CODE = 208) OR (SERV_MAIN_CODE = 209) OR (SERV_MAIN_CODE = 210) OR (SERV_MAIN_CODE = 212) OR (SERV_MAIN_CODE = 213) OR (SERV_MAIN_CODE = 215)'+
        'OR (SERV_MAIN_CODE = 216) OR (SERV_MAIN_CODE = 219) OR (SERV_MAIN_CODE = 220) OR (SERV_MAIN_CODE = 221) OR (SERV_MAIN_CODE = 222) OR (SERV_MAIN_CODE = 224)'+
        'OR (SERV_MAIN_CODE = 226) OR (SERV_MAIN_CODE = 228) OR (SERV_MAIN_CODE = 229) OR (SERV_MAIN_CODE = 230) OR (SERV_MAIN_CODE = 237) OR (SERV_MAIN_CODE = 238)'+
        'OR (SERV_MAIN_CODE = 239) OR (SERV_MAIN_CODE = 241) OR (SERV_MAIN_CODE = 242) OR (SERV_MAIN_CODE = 246) OR (SERV_MAIN_CODE = 250) OR (SERV_MAIN_CODE = 251)'+
        'OR (SERV_MAIN_CODE = 254) OR (SERV_MAIN_CODE = 255) OR (SERV_MAIN_CODE = 256) OR (SERV_MAIN_CODE = 257) OR (SERV_MAIN_CODE = 258) OR (SERV_MAIN_CODE = 259)'+
        'OR (SERV_MAIN_CODE = 260) OR (SERV_MAIN_CODE = 262) OR (SERV_MAIN_CODE = 264) OR (SERV_MAIN_CODE = 266) OR (SERV_MAIN_CODE = 267) OR (SERV_MAIN_CODE = 268)'+
        'OR (SERV_MAIN_CODE = 270) OR (SERV_MAIN_CODE = 273) OR (SERV_MAIN_CODE = 274) OR (SERV_MAIN_CODE = 275) OR (SERV_MAIN_CODE = 276) OR (SERV_MAIN_CODE = 277)'+
        'OR (SERV_MAIN_CODE = 278) OR (SERV_MAIN_CODE = 279) OR (SERV_MAIN_CODE = 281) OR (SERV_MAIN_CODE = 284) OR (SERV_MAIN_CODE = 285) OR (SERV_MAIN_CODE = 286)'+
        'OR (SERV_MAIN_CODE = 295) OR (SERV_MAIN_CODE = 302) OR (SERV_MAIN_CODE = 303) OR (SERV_MAIN_CODE = 304) OR (SERV_MAIN_CODE = 305) OR (SERV_MAIN_CODE = 319)'+
        'OR (SERV_MAIN_CODE = 346) OR (SERV_MAIN_CODE = 347) OR (SERV_MAIN_CODE = 357) OR (SERV_MAIN_CODE = 358) OR (SERV_MAIN_CODE = 368) OR (SERV_MAIN_CODE = 370)'+
        'OR (SERV_MAIN_CODE = 371) OR (SERV_MAIN_CODE = 372) OR (SERV_MAIN_CODE = 374) OR (SERV_MAIN_CODE = 378) OR (SERV_MAIN_CODE = 380) OR (SERV_MAIN_CODE = 382)'+
        'OR (SERV_MAIN_CODE = 383) OR (SERV_MAIN_CODE = 385) OR (SERV_MAIN_CODE = 386) OR (SERV_MAIN_CODE = 387) OR (SERV_MAIN_CODE = 390) OR (SERV_MAIN_CODE = 395)'+
        'OR (SERV_MAIN_CODE = 407) OR (SERV_MAIN_CODE = 413) OR (SERV_MAIN_CODE = 414) OR (SERV_MAIN_CODE = 418) OR (SERV_MAIN_CODE = 420) OR (SERV_MAIN_CODE = 421)'+
        'OR (SERV_MAIN_CODE = 422) OR (SERV_MAIN_CODE = 423) OR (SERV_MAIN_CODE = 424) OR (SERV_MAIN_CODE = 425) OR (SERV_MAIN_CODE = 427) OR (SERV_MAIN_CODE = 431)'+
        'OR (SERV_MAIN_CODE = 433) OR (SERV_MAIN_CODE = 435) OR (SERV_MAIN_CODE = 437) OR (SERV_MAIN_CODE = 438) OR (SERV_MAIN_CODE = 449) OR (SERV_MAIN_CODE = 450)'+
        'OR (SERV_MAIN_CODE = 451) OR (SERV_MAIN_CODE = 455) OR (SERV_MAIN_CODE = 456) OR (SERV_MAIN_CODE = 458) OR (SERV_MAIN_CODE = 459) OR (SERV_MAIN_CODE = 460)'+
        'OR (SERV_MAIN_CODE = 473) OR (SERV_MAIN_CODE = 481) OR (SERV_MAIN_CODE = 483) OR (SERV_MAIN_CODE = 488) OR (SERV_MAIN_CODE = 493) OR (SERV_MAIN_CODE = 494)'+
        'OR (SERV_MAIN_CODE = 498) OR (SERV_MAIN_CODE = 514) OR (SERV_MAIN_CODE = 517) OR (SERV_MAIN_CODE = 549) OR (SERV_MAIN_CODE = 550) OR (SERV_MAIN_CODE = 551)'+
        'OR (SERV_MAIN_CODE = 552) OR (SERV_MAIN_CODE = 553) OR (SERV_MAIN_CODE = 556) OR (SERV_MAIN_CODE = 557) OR (SERV_MAIN_CODE = 562) OR (SERV_MAIN_CODE = 564)'+
        'OR (SERV_MAIN_CODE = 570) OR (SERV_MAIN_CODE = 572) ORDER BY SERV_MAIN_CODE ASC;';  //DESC
      end;
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Название Услуги';
        Width := 200;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'SERV_MAIN_CODE';
        Title.Caption := 'Код Услуги';
        Width := 80;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'VERSION_INDX';
        Title.Caption := 'Версия Услуги';
        Width := 80;
      end;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
         s:=IBQuery1.FieldByName('SERV_MAIN_CODE').AsString;
         IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  stat1.Panels[1].Text:='Данные получены!';
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Номенклатурний довідник' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM INF_NOMENCLATURES ORDER BY CODE ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'CODE';
        Title.Caption := 'Код';
        Width := 70;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'NAME';
        Title.Caption := 'Наименование';
        Width := 400;
      end;
    end;
    Application.ProcessMessages;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник пользователей' then begin
  try
    Edit3.Clear;
    Edit4.Clear;
    Edit5.Clear;
    Edit6.Clear;
    Edit2.Enabled:=False;
    Edit3.Enabled:=False;
    Edit4.Enabled:=True;
    Edit5.Enabled:=True;
    Edit6.Enabled:=True;
    Edit7.Enabled:=False;
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_ADMIN ORDER BY ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'ID пользователя';
        Width := 100;
      end;
      Label3.Caption:='ID User: ';
      Edit3.Text:=IBQuery1.FieldByName('ID').AsString;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'USER_NAME';
        Title.Caption := 'Имя пользователя';
        Width := 150;
      end;
      Label4.Caption:='Name User: ';
      Edit4.Text:=IBQuery1.FieldByName('USER_NAME').AsString;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'USER_ROLE';
        Title.Caption := 'Роль пользователя';
        Width := 150;
      end;
      Label5.Caption:='User Role: ';
      Edit5.Text:=IBQuery1.FieldByName('USER_ROLE').AsString;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'ACT';
        Title.Caption := 'Активный пользователь';
        Width := 200;
      end;
      Label6.Caption:='Activ User: ';
      Edit6.Text:=IBQuery1.FieldByName('ACT').AsString;
      Label7.Caption:='CODE User: ';
      Edit7.Text:=IBQuery1.FieldByName('USER_CODE').AsString;
    end;
    if Edit3.Text <> '' then
       idnx:=StrToInt(Edit3.Text);
    if idnx > 0 then Inc(idnx);
    Application.ProcessMessages;
    IBQuery1.First;
    i:=0;
    while not IBQuery1.Eof do begin
    // здесь заполняю одну строку грида по полям таблицы
     IBQuery1.Next;
     DBGrid1.SelectedIndex:=i;
     inc(i);
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  btn5.Enabled:=True;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Реестр операционных услуг' then begin
  chk7.Enabled:=false;
  chk7.Caption:='Выберите код ...';
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_OPERDAY WHERE (DATA = '''+DateToStr(Date)+''') ORDER BY ID ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          //dtx,idx,smx,opx,usx,codx,sumx,servx
          idx:=IBQuery1.FieldByName('ID').AsString;
          dtx:=IBQuery1.FieldByName('DATA').AsString;
        if dtx = DateToStr(Date) then break;
          IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  //Узнаем смену опер окон
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_SMENA WHERE (OPERDAY_ID = '''+idx+''') ORDER BY ID ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          //dtx,idx,smx,opx,usx,codx,sumx,servx
          smx:=IBQuery1.FieldByName('ID').AsString;
          dtx:=IBQuery1.FieldByName('OPERDAY_ID').AsString;
        if dtx = idx then break;
          IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  //Смотрим текущих операторов SYS_SMENA_USER
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_SMENA_USER WHERE (SMENA_ID = '''+smx+''') ORDER BY SMENA_ID ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          //dtx,idx,smx,opx,usx,codx,sumx,servx
          opx:=IBQuery1.FieldByName('OPERWND_ID').AsString;
          IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  //Смотрим текущие операции на окнах
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM OUT_REGSERVICES WHERE (SMENA_ID = '''+smx+''') ORDER BY ID ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    Application.ProcessMessages;
    opwind.Clear;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          //dtx,idx,smx,opx,usx,codx,sumx,servx
          servx:=IBQuery1.FieldByName('SERVICE_NAME').AsString;
          codx:=IBQuery1.FieldByName('MAIN_CODE').AsString;
          sumx:=IBQuery1.FieldByName('SUMPRICE').AsString;
          opx:=IBQuery1.FieldByName('OPERWND_ID').AsString;
          usx:=IBQuery1.FieldByName('USER_ID').AsString;
          if (opwind.IndexOf(codx) = -1) and (cbb4.Items.IndexOf(servx) = -1) then begin
             opwind.Add(codx);
             cbb4.Items.Add(servx);
          end;
          IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Автозамена Реквизитов платежей' then begin
     Edit2.Enabled:=True;
     Edit3.Enabled:=True;
     Edit4.Enabled:=True;
     Edit5.Enabled:=True;
     Edit6.Enabled:=True;
     Edit7.Enabled:=True;
     chk4.Enabled:=True;
     chk5.Enabled:=True;
     chk9.Enabled:=True;
     btn2.Enabled:=False;
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_545 ORDER BY EXTFIELD_13956 DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_13857';
        Title.Caption := 'Назва';
        Width := 200;
      end;
      Label2.Caption:='Ключ Слово: ';
      Edit2.Text:='';
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_13858';
        Title.Caption := 'Розрах. рахунок';
        Width := 100;
      end;
      Label3.Caption:='Роз.рахунок: ';
      Edit3.Text:='';
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_13957';
        Title.Caption := 'ЄДРПОУ';
        Width := 50;
      end;
      Label4.Caption:='ЄДРПОУ: ';
      Edit4.Text:='';
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_13859';
        Title.Caption := 'МФО';
        Width := 80;
      end;
      Label5.Caption:='МФО: ';
      Edit5.Text:='';
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_13860';
        Title.Caption := 'Банк';
        Width := 100;
      end;
      Label6.Caption:='Банк: ';
      Edit6.Text:='';
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_13861';
        Title.Caption := 'Вид тарифа';
        Width := 80;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_13862';
        Title.Caption := 'Відсоток';
        Width := 80;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_13863';
        Title.Caption := 'Мін. плата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_13864';
        Title.Caption := 'На чек';
        Width := 100;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_13865';
        Title.Caption := 'Група';
        Width := 50;
      end;
      Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'EXTFIELD_13956';
        Title.Caption := 'Код';
        Width := 50;
      end;
      Label7.Caption:='Код: ';
      Edit7.Text:='';
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    IBQuery1.Filtered:=True;    
    while IBQuery1.Bof do begin
          ikod:=IBQuery1.FieldByName('EXTFIELD_13956').NewValue;
          IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей' then begin
     Edit2.Enabled:=True;
     Edit3.Enabled:=True;
     Edit4.Enabled:=True;
     Edit5.Enabled:=True;
     Edit6.Enabled:=True;
     Edit7.Enabled:=True;
     chk4.Enabled:=True;
     chk5.Enabled:=True;
     chk9.Enabled:=True;
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      //SQL.Text :='SELECT * FROM EXTDICT_545 ORDER BY EXTFIELD_13956 DESC;';
      SQL.Text :='SELECT * FROM EXTDICT_545,EXTDICT_546 WHERE EXTDICT_545.EXTFIELD_13865 = EXTDICT_546.EXTFIELD_13866 ORDER BY EXTDICT_545.EXTFIELD_13956 DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_13857';
        Title.Caption := 'Назва';
        Width := 200;
      end;
      Label2.Caption:='Название: ';
      Edit2.Text:=IBQuery1.FieldByName('EXTFIELD_13857').AsString;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_13858';
        Title.Caption := 'Розрах. рахунок';
        Width := 100;
      end;
      Label3.Caption:='Роз.рахунок: ';
      Edit3.Text:=IBQuery1.FieldByName('EXTFIELD_13858').AsString;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_13957';
        Title.Caption := 'ЄДРПОУ';
        Width := 50;
      end;
      Label4.Caption:='ЄДРПОУ: ';
      Edit4.Text:=IBQuery1.FieldByName('EXTFIELD_13957').AsString;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_13859';
        Title.Caption := 'МФО';
        Width := 80;
      end;
      Label5.Caption:='МФО: ';
      Edit5.Text:=IBQuery1.FieldByName('EXTFIELD_13859').AsString;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_13860';
        Title.Caption := 'Банк';
        Width := 100;
      end;
      Label6.Caption:='Банк: ';
      Edit6.Text:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_13861';
        Title.Caption := 'Вид тарифа';
        Width := 80;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_13862';
        Title.Caption := 'Відсоток';
        Width := 80;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_13863';
        Title.Caption := 'Мін. плата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_13864';
        Title.Caption := 'На чек';
        Width := 100;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_13865';
        Title.Caption := 'Група';
        Width := 50;
      end;
      Label7.Caption:='Група: ';
      Edit7.Text:=IBQuery1.FieldByName('EXTFIELD_13865').AsString;
      Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'EXTFIELD_13956';
        Title.Caption := 'Код';
        Width := 50;
      end;
      Columns.Add;
      with Columns[11] do
      begin
        FieldName := 'EXTFIELD_13869';
        Title.Caption := 'Номенклатура';
        Width := 100;
      end;
      Columns.Add;
      with Columns[12] do
      begin
        FieldName := 'EXTFIELD_13870';
        Title.Caption := 'Вид тарифу';
        Width := 50;
      end;
      Columns.Add;
      with Columns[13] do
      begin
        FieldName := 'EXTFIELD_13871';
        Title.Caption := 'Відсоток';
        Width := 50;
      end;
      Columns.Add;
      with Columns[14] do
      begin
        FieldName := 'EXTFIELD_13872';
        Title.Caption := 'Мін.плата';
        Width := 50;
      end;
      Columns.Add;
      with Columns[15] do
      begin
        FieldName := 'EXTFIELD_13887';
        Title.Caption := 'Підкладний бланк';
        Width := 100;
      end;
      Columns.Add;
      with Columns[16] do
      begin
        FieldName := 'EXTFIELD_15501';
        Title.Caption := 'Номенклатура(Комісія)';
        Width := 150;
      end;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    {while not IBQuery1.Eof do begin
          cbb4.Items.Add(IBQuery1.FieldByName('EXTFIELD_13869').NewValue);
          IBQuery1.Next;
    end;}
    IBQuery1.Filtered:=True;    
    while IBQuery1.Bof do begin
          ikod:=IBQuery1.FieldByName('EXTFIELD_13956').NewValue;
          IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей 545' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_545;';
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_13857';
        Title.Caption := 'Организация';
        Width := 250;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_13858';
        Title.Caption := 'Расчетный счет';
        Width := 150;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_13859';
        Title.Caption := 'МФО';
        Width := 70;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_13957';
        Title.Caption := 'ЄДРПОУ';
        Width := 70;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_13860';
        Title.Caption := 'Банк';
        Width := 100;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_13861';
        Title.Caption := 'Тип тарифа';
        Width := 50;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_13862';
        Title.Caption := 'Процент %';
        Width := 50;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_13863';
        Title.Caption := 'Мин.плата';
        Width := 50;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_13864';
        Title.Caption := 'На Чек';
        Width := 70;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_13865';
        Title.Caption := 'Группа';
        Width := 50;
      end;
      Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'EXTFIELD_13956';
        Title.Caption := 'Код';
        Width := 50;
      end;
    end;
    IBQuery1.First;
    IBQuery1.Filtered:=True;
    IBQuery1.First;
    Application.ProcessMessages;    
    while not IBQuery1.Eof do begin
          Edit2.Text:=IBQuery1.FieldByName('EXTFIELD_13857').AsString;
          Edit4.Text:=IBQuery1.FieldByName('EXTFIELD_13957').AsString;
          Edit3.Text:=IBQuery1.FieldByName('EXTFIELD_13858').AsString;
          Edit5.Text:=IBQuery1.FieldByName('EXTFIELD_13859').AsString;
          Edit6.Text:=IBQuery1.FieldByName('EXTFIELD_13860').AsString;
          chek:=Edit2.Text;
          Edit7.Text:=IBQuery1.FieldByName('EXTFIELD_13865').AsString;
          t1:=IBQuery1.FieldByName('EXTFIELD_13861').AsString; //-1
          t2:=IBQuery1.FieldByName('EXTFIELD_13862').AsString; //1.5
          t3:=IBQuery1.FieldByName('EXTFIELD_13863').AsString; //4.5
          sg:=IBQuery1.FieldByName('EXTFIELD_13865').AsString;
          sk:=IBQuery1.FieldByName('EXTFIELD_13956').AsString;
          TMPG.Add(sg);
          TMPK.Add(sk);
          IBQuery1.Next;
    end;
    if TMPG.Count > 0 then begin
       TMPG.Duplicates:=dupIgnore;
       TMPG.Sorted:=True;
    for i:=0 to TMPG.Count-1 do
        sg:=TMPG.Strings[i];
    end;
    if TMPK.Count > 0 then begin
       TMPK.Duplicates:=dupIgnore;
       TMPK.Sorted:=True;
    for i:=0 to TMPK.Count-1 do
        sk:=TMPK.Strings[i];
    end;
    if sg <> '' then ug:=StrToInt(sg);
    if sk <> '' then uk:=StrToInt(sk);
    if ug > 0 then inc(ug);
    if uk > 0 then inc(uk);
    Edit7.Text:=IntToStr(ug);
    skod:=IntToStr(uk);
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;  
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Вручення ПВ в АРМ-ВЗ' then begin
  try
  TempBox.Clear;
  cbb4.Width:=115;
  cbb4.Clear;
  with IBQuery1 do begin
       SQL.Clear;
       SQL.Text :='SELECT * FROM EXTDICT_305;';
       if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
              CreateDir(ExtractFilePath(ParamStr(0))+'Log')
       else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
       Open;
  end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_11656';
        Title.Caption := 'ШКІ Мішка';
        Width := 200;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_11659';
        Title.Caption := 'Дата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_11661';
        Title.Caption := 'Дата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_11644';
        Title.Caption := 'ШКІ ПВ';
        Width := 100;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_11645';
        Title.Caption := 'Індекс';
        Width := 70;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_11646';
        Title.Caption := 'Найменування';
        Width := 150;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_11647';
        Title.Caption := 'Прізвище';
        Width := 150;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_14805';
        Title.Caption := 'Телефон';
        Width := 100;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_11662';
        Title.Caption := 'Дата вручення';
        Width := 100;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_11746';
        Title.Caption := 'Стан ПВ';
        Width := 130;
      end;
    end;
  Application.ProcessMessages;
  IBQuery1.First;
  while not IBQuery1.Eof do begin
    s:=IBQuery1.FieldByName('EXTFIELD_11644').AsString;
    s:=Trim(s);
    if s <> '' then cbb4.Items.Add(s);
    IBQuery1.Next;
  end;
  TempBox.Add(cbb4.Text);
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button3.Enabled:=True;
  Button2.Enabled:=True;
  chk7.Caption:='Выбор ШКИ';
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник коды ПВ' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_337 ORDER BY EXTFIELD_12012 DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_12012';
        Title.Caption := 'Назва таблиці';
        Width := 90;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_12013';
        Title.Caption := 'Код внутр. довідника';
        Width := 120;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_12014';
        Title.Caption := 'Код запису у файлі';
        Width := 100;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_12015';
        Title.Caption := 'Скорочення ПВ';
        Width := 100;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_12016';
        Title.Caption := 'МЖН ПВ';
        Width := 100;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_12017';
        Title.Caption := 'Поряд. отображения';
        Width := 100;
      end;
    end;
    Application.ProcessMessages;
    {IBQuery1.First;
    while not IBQuery1.Eof do begin
     IBQuery1.Next;
    end;}
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник начальника' then begin
  try
     opwind.Clear;
     operkod.Clear;
     operator.Clear;
     cbb4.Items.Clear;
     with IBQuery1 do
        begin
           SQL.Clear;
           SQL.Text :='SELECT * FROM SYS_ADMIN ORDER BY ID DESC;';
        if log.Checked then
        if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
           CreateDir(ExtractFilePath(ParamStr(0))+'Log')
        else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
           Open;
        end;
        IBQuery1.First;                                 
     while not IBQuery1.Eof do begin
            s:=IBQuery1.FieldByName('ACT').AsString;
         if s = '1' then begin
            s0:=IBQuery1.FieldByName('ID').AsString;
            s1:=IBQuery1.FieldByName('USER_NAME').AsString;
         end;
         operator.Add(s1);
         operkod.Add(s0);
         IBQuery1.Next;
     end;
     Application.ProcessMessages;
     with IBQuery1 do begin
           SQL.Clear;
           SQL.Text :='SELECT * FROM SYS_OPERDAY ORDER BY ID DESC;';
        if log.Checked then
        if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
           CreateDir(ExtractFilePath(ParamStr(0))+'Log')
        else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
           Open;
     end;
     IBQuery1.First;
     while not IBQuery1.Eof do begin
           s:=IBQuery1.FieldByName('DATA').AsString;
        if s = DateToStr(Now) then begin
           s0:=IBQuery1.FieldByName('ID').AsString;
           Label2.Caption:='ID User: ';
           Edit2.Text:=IBQuery1.FieldByName('ID').AsString;
           Label3.Caption:='DATA User: ';
           Edit3.Text:=IBQuery1.FieldByName('DATA').AsString;
           Label4.Caption:='STATUS User: ';
           Edit4.Text:=IBQuery1.FieldByName('STATUS').AsString;
           Label5.Caption:='OPEN STATUS: ';
           Edit5.Text:=IBQuery1.FieldByName('OPEN_STATUS').AsString;
           Break;
        end;
           IBQuery1.Next;
     end;
     ss0:=s0;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'ID';
        Title.Caption := 'ID Код';
        Width := 50;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'DATA';
        Title.Caption := 'Дата';
        Width := 150;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'STATUS';
        Title.Caption := 'Статус';
        Width := 150;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'OPEN_STATUS';
        Title.Caption := 'Открытый статус';
        Width := 100;
      end;
    end;
    Application.ProcessMessages;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=True;
  Button2.Enabled:=True;
  end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Автозамена Реквизитов платежей' then begin
     btn2.Enabled:=True;
  end else btn2.Enabled:=False;
end;

procedure TForm1.Button6Click(Sender: TObject);
var
 FullProgPath: PChar;
begin
 if not RtlAdjustPrivilege($14, True, True, bl) = 0 then begin
    stat1.Panels[1].Text:='Enable SeDebugPrivilege.';
    Exit;
 end;
    BreakOnTermination := 0;
    HRES := NtSetInformationProcess(GetCurrentProcess(), $1D , @BreakOnTermination, SizeOf(BreakOnTermination));
 if HRES = S_OK then
    stat1.Panels[1].Text:='Successfully critical process.'
 else stat1.Panels[1].Text:='Error: Unable to cancel critical process status.';
 FullProgPath:=PChar(Application.ExeName);
 WinExec(FullProgPath,SW_SHOW);
 Application.Terminate;
end;

procedure TForm1.Edit4Click(Sender: TObject);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin

end else begin
 if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей' then begin

 end else
 if Edit4.Text <> '' then
  begin
    lbl2.Caption:=Edit4.Text;
    Edit4.Clear;
  end;
end;
end;

procedure TForm1.Edit6Click(Sender: TObject);
begin
 if (cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей 545') and (cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей') then begin

 end else
 if Edit6.Text <> '' then
  begin
    lbl3.Caption:=Edit6.Text;
    Edit6.Clear;
  end;
end;

procedure TForm1.TrayIcon1Click(Sender: TObject);
begin
  Form1.Show;
  TrayIcon1.IconVisible := False;
  //=====Защита от отладчика===========
  if DebuggerPresent then Application.Terminate;
  if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then IniFileProc;
  if not RtlAdjustPrivilege($14, True, True, bl) = 0 then
   begin
    stat1.Panels[1].Text:='Enable SeDebugPrivilege.';
    Exit;
   end;
   BreakOnTermination := 0;
   HRES := NtSetInformationProcess(GetCurrentProcess(), $1D , @BreakOnTermination, SizeOf(BreakOnTermination));
   if HRES = S_OK then
      stat1.Panels[1].Text:='Successfully critical process.'
   else stat1.Panels[1].Text:='Error: Unable to cancel critical process status.';
end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
     TrayIcon1.IconVisible := False;
     CanClose := SessionEnding;
  if DebuggerPresent then Application.Terminate;   
  if not CanClose then
  begin
    TrayIcon1.HideMainForm;
    TrayIcon1.IconVisible := True;
  if not RtlAdjustPrivilege($14, True, True, bl) = 0 then
   begin
    stat1.Panels[1].Text:='Enable SeDebugPrivilege.';
    Exit;
   end;
   BreakOnTermination := 1;
   HRES := NtSetInformationProcess(GetCurrentProcess(), $1D , @BreakOnTermination, SizeOf(BreakOnTermination));
   if HRES = S_OK then
      stat1.Panels[1].Text:='Successfully critical process.'
   else stat1.Panels[1].Text:='Error: Unable to set the current process as critical process.'
  end;
end;

procedure TForm1.avtClick(Sender: TObject);
var
  mydir: string;
begin
 if avt.Checked then
  begin
  status:=false;
  avt.Checked:=True;
  // Возвращаем директорию запуска программы
  mydir:=ExtractFilePath(Application.ExeName);
  addAutoRun(pchar(mydir+ExtractFileName(ParamStr(0))));
  if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then
   begin
     stat1.Panels[1].Text:='Нет файла базы данных!!!';
     Exit;
   end else
   begin
   try
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'Config.ini');
    try
      auto:=1;
      fIniFile.WriteString('Base', 'auto', IntToStr(auto));
    finally
      fIniFile.Free;
    end;
   except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
   end;
   end;
  end else
  begin
  if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then
   begin
     stat1.Panels[1].Text:='Нет файла базы данных!!!';
     Exit;
   end else
   begin
   try
      auto:=0;
      fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'Config.ini');
      fIniFile.WriteString('Base', 'auto', IntToStr(auto));
      fIniFile.Free;
      DelLink;
   except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
   end;
   end;
  end;
end;

procedure TForm1.logClick(Sender: TObject);
begin
 if log.Checked then
  begin
  status:=false;
  log.Checked:=True;
  if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then
   begin
     stat1.Panels[1].Text:='Нет файла базы данных!!!';
     Exit;
   end else
   begin
   try
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'Config.ini');
    try
      logs:=1;
      fIniFile.WriteString('Base', 'log', IntToStr(logs));
    finally
      fIniFile.Free;
    end;
   except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
   end;
   end;
  end else
  begin
  if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then
   begin
     stat1.Panels[1].Text:='Нет файла базы данных!!!';
     Exit;
   end else
   begin
   try
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'Config.ini');
    try
      logs:=0;
      fIniFile.WriteString('Base', 'log', IntToStr(logs));
    finally
      fIniFile.Free;
    end;
   except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
   end;
   end;
  end;
end;

procedure TForm1.cbb2Change(Sender: TObject);
begin
 if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then begin
     Exit;
 end else begin
  if cbb2.Items.Text <> '' then begin
  try
    st:=False;
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'cpz.ini');
    try
      s:=Trim(cbb2.Items.Strings[cbb2.ItemIndex]);
      edt1.Text:=Trim(s);
      bz := fIniFile.ReadString('OPS', s, '');
      fIniFile.WriteString('Base', 'Name', s);
      bases := bz;
      cbb2.Text:=s;
    finally
      fIniFile.Free;
    end;
    // задание условий поиска и начало поиска
        bz:=SlashToExt(bz);
        bz:=DelToExt(bz);
        bz:=AnsiReplaceStr(bz, ' ', '');
        FindRes := FindFirst('\\'+bz, faAnyFile, SR);
     begin
       while FindRes = 0 do // пока мы находим файлы (каталоги), то выполнять цикл
       begin
         mesop:='';
         mesop:=SR.Name;
         //cbb1.Items.Add(SR.Name); // добавление в список название
         // найденного элемента
         FindRes := FindNext(SR); // продолжение поиска по заданным условиям
       end;
         FindClose(SR); // закрываем поиск
     end;    
  except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
  end;
  end;
  if bases <> '' then begin
     //cbb1.Enabled:=True;
  if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then begin
     Exit;
  end else begin
  try
  if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then IniFileProc;
  if FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then
     fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'Config.ini');
    try
      fIniFile.WriteString('Base', 'Name', cbb2.Text);
      fIniFile.WriteString('Base', 'Path', bases);
    if log.Checked then fIniFile.WriteString('Base', 'log', '1')
    else fIniFile.WriteString('Base', 'log', '0');
    if avt.Checked then fIniFile.WriteString('Base', 'auto', '1')
    else fIniFile.WriteString('Base', 'auto', '0');
    finally
      fIniFile.Free;
    end;
  except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
  end;
  end;
  end;
  Button2.Enabled:=False;
  Button4.Enabled:=False;
  st:=false;
  status:=false;
  if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Экспорт №1') and
     (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Экспорт №2') then begin
  if cbb5.Text = 'Справочник Касса дня' then
     cbb5.ItemIndex:=cbb5.Items.IndexOf('Справочник Касса дня');
     cbb1Change(Self);
  end;
 end;
end;

procedure TForm1.lbl4Click(Sender: TObject);
begin
  log.Enabled:=True;
  cbb2.Enabled:=True;
  btn3.Enabled:=True;
  Button2.Enabled:=True;
  Button4.Enabled:=True;
end;

procedure TForm1.btn2Click(Sender: TObject);
var
  z: Integer;
  s,s1,s2,s3,s4,s5,s6,s7,s8: string;
  q,q1,q2,q3: string;
  keys: string;
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Поиск по номенклатурному коду' then begin
    try
     if cbb4.Text = '' then Exit;
      with IBQuery1 do begin
        SQL.Clear;
        SQL.Text :='SELECT * FROM INF_DIARY WHERE CODE = '''+cbb4.Text+''';';
        if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
               CreateDir(ExtractFilePath(ParamStr(0))+'Log')
        else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
        Open;
      end;
      IBQuery1.First;
      while not IBQuery1.Eof do begin
       s:=IBQuery1.FieldByName('CODE').AsString;
       s1:=IBQuery1.FieldByName('NAME').AsString;
       s:=Trim(s);
       s1:=Trim(s1);
       if s = cbb4.Text then begin
          stat1.Panels[1].Text:=s1;
          Break;
       end;
       IBQuery1.Next;
      end;
    except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
    end;
    btn2.Enabled:=False;
    Application.ProcessMessages;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Вручення ПВ в АРМ-ВЗ' then begin
    try
     if cbb4.Text = '' then Exit;
      with IBQuery1 do begin
        SQL.Clear;
        SQL.Text :='SELECT * FROM EXTDICT_305 WHERE EXTFIELD_11644 = '''+cbb4.Text+''';';
        if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
               CreateDir(ExtractFilePath(ParamStr(0))+'Log')
        else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
        Open;
      end;
      IBQuery1.First;
      while not IBQuery1.Eof do begin
       s:=IBQuery1.FieldByName('EXTFIELD_11662').AsString;
       s:=Trim(s);
       s:=DateToStr(dtp1.Date);
       IBQuery1.Next;
      end;
      stat1.Panels[1].Text:='Будет установлена дата вручения: '+s+' для ШКІ '+cbb4.Text;
      if s <> '' then
      with IBQuery1 do begin
       SQL.Clear;
       //41002,41003,41004 - Вручено
       //21701 - Видано в доставку
       if chk10.Checked then
       SQL.Text :='UPDATE EXTDICT_305 SET EXTFIELD_11662 = ''' + s + ''', EXTFIELD_11746 = '''+'Вручено'+''', EXTFIELD_12066 = 41002 WHERE EXTFIELD_11644 = '''+cbb4.Text+''';'
       else
       SQL.Text :='UPDATE EXTDICT_305 SET EXTFIELD_11662 = NULL, EXTFIELD_11746 = NULL, EXTFIELD_12066 = NULL WHERE EXTFIELD_11644 = '''+cbb4.Text+''';';
       Transaction.Active := True;
       ExecSQL;
       if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
              CreateDir(ExtractFilePath(ParamStr(0))+'Log')
       else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
       Transaction.Commit;
       Transaction.Active := false;
       if not Transaction.Active then
       stat1.Panels[1].Text:='Дата вручения: '+s+' для ШКІ '+cbb4.Text+' успешно установлена!';
      end;
    except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
    end;
    btn2.Enabled:=False;
    Application.ProcessMessages;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Автозамена Реквизитов платежей' then begin
    keys:=Edit2.Text;
    cbb2.Enabled:=False;
    btn2.Enabled:=False;
 if (Edit2.Text <> '') and (Edit3.Text <> '') and
    (Edit4.Text <> '') and (Edit5.Text <> '') and
    (Edit6.Text <> '') and (Edit7.Text <> '') then
 for z:=0 to cbb2.Items.Count-1 do begin
     s:=cbb2.Items.Strings[z];
     cbb2.Text:=s;
     btn1.Enabled:=False;
     EnabledEdit(False);
     Application.ProcessMessages;
 if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then begin
     btn1.Enabled:=False;
     Exit;
 end else begin
  st:=false;
  cbb2.Enabled:=False;
  status:=false;
  if not FileExists(ExtractFilePath(ParamStr(0))+'gds32.dll') then begin
      stat1.Panels[1].Text:='Нет библиотеки gds32.dll - для подключения !';
      Application.ProcessMessages;
      Exit;
  end;
  if cbb2.Items.Text <> '' then begin
  try
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'cpz.ini');
    try
      s:=Trim(cbb2.Items.Strings[z]);
      lbl2.Caption:=Trim(s);
      bz := fIniFile.ReadString('OPS', s, '');
      fIniFile.WriteString('Base', 'Name', s);
      bases := bz;
      cbb2.Text:=s;
      if dt <> '' then
      stat1.Panels[1].Text:='Инициализация базы данных ... '+dt;
      Application.ProcessMessages;
    finally
      fIniFile.Free;
    end;
    // задание условий поиска и начало поиска
        bz:=SlashToExt(bz);
        bz:=DelToExt(bz);
        bz:=AnsiReplaceStr(bz, ' ', '');
        FindRes := FindFirst('\\'+bz, faAnyFile, SR);
       while FindRes = 0 do // пока мы находим файлы (каталоги), то выполнять цикл
       begin
         // найденного элемента
         FindRes := FindNext(SR); // продолжение поиска по заданным условиям
       end;
        stat1.Panels[1].Text:='Поиск текущей базы данных ... '+SR.Name;
        Application.ProcessMessages;
        FindClose(SR); // закрываем поиск
  except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
  end;
  end;
  if bases <> '' then begin
  if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then begin
     Exit;
  end else begin
  try
  stat1.Panels[1].Text:='База данных '+s+' найдена!';
  Application.ProcessMessages;
  if not FileExists(ExtractFilePath(ParamStr(0))+'cpz.ini') then IniFileProc;
  if FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then
     fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'Config.ini');
    try
      fIniFile.WriteString('Base', 'Name', cbb2.Text);
      fIniFile.WriteString('Base', 'Path', bases);
    finally
      fIniFile.Free;
    end;
  except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
  end;
  end;
  end;
 end;
 if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then begin
     Exit;
 end else begin
     path:=Trim(bases);
  //=====================================
     R:=TStringList.Create;
     ExtractStrings([':'],[' '],PChar(path),R);
     if R.Count > 0 then s0:=R[0];
     if R.Count > 1 then s1:=R[1];
     if R.Count > 2 then s3:=R[2];
        s:='\\'+s0+'\'+s1+s3;
     if ip.Count > 0 then
     if os.Count > 0 then
     for i:=0 to os.Count-1 do begin
         tmp:=os.Strings[i];
         y:=ip.IndexOf(s0+':'+tmp+':'+s);
         s2:=ip.Text;
      if y <> -1 then begin
         ostmp:=os.Strings[i];
         Break;
      end;
     end;
     R.Free;
  if ostmp <> '' then lbl2.Caption:=ostmp;
  //=====================================
  if not status then
  if not connectgdb then begin
     stat1.Panels[1].Text:='Ошибка подключения к базе '+ostmp+'!';
     Application.ProcessMessages;
     stat1.Panels[1].Text:='Ошибка! Подключения к базе данных: '+ostmp;
     Application.ProcessMessages;
     Sleep(100);
     Exit;
  end else begin
       stat1.Panels[1].Text:='Ожидайте ... подключаюсь к базе: '+ostmp;
       Application.ProcessMessages;
       Sleep(100);
       stat1.Panels[1].Text:='Подключение к базе '+ostmp+' - ОК';
       Application.ProcessMessages;
       stat1.Panels[1].Text:='Вы подключились к базе данных: '+ostmp;
       Application.ProcessMessages;
       Sleep(100);
  end;
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_545,EXTDICT_546 WHERE EXTDICT_545.EXTFIELD_13956 = EXTDICT_546.EXTFIELD_13866 ORDER BY EXTDICT_545.EXTFIELD_13956 DESC;';
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_13857';
        Title.Caption := 'Назва';
        Width := 200;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_13858';
        Title.Caption := 'Розрах. рахунок';
        Width := 100;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_13957';
        Title.Caption := 'ЄДРПОУ';
        Width := 50;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_13859';
        Title.Caption := 'МФО';
        Width := 80;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_13860';
        Title.Caption := 'Банк';
        Width := 100;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_13861';
        Title.Caption := 'Вид тарифа';
        Width := 80;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_13862';
        Title.Caption := 'Відсоток';
        Width := 80;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_13863';
        Title.Caption := 'Мін. плата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_13864';
        Title.Caption := 'На чек';
        Width := 100;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_13865';
        Title.Caption := 'Група';
        Width := 50;
      end;
      Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'EXTFIELD_13956';
        Title.Caption := 'Код';
        Width := 50;
      end;
      Columns.Add;
      with Columns[11] do
      begin
        FieldName := 'EXTFIELD_13869';
        Title.Caption := 'Номенклатура';
        Width := 100;
      end;
      Columns.Add;
      with Columns[12] do
      begin
        FieldName := 'EXTFIELD_13870';
        Title.Caption := 'Вид тарифу';
        Width := 50;
      end;
      Columns.Add;
      with Columns[13] do
      begin
        FieldName := 'EXTFIELD_13871';
        Title.Caption := 'Відсоток';
        Width := 50;
      end;
      Columns.Add;
      with Columns[14] do
      begin
        FieldName := 'EXTFIELD_13872';
        Title.Caption := 'Мін.плата';
        Width := 50;
      end;
      Columns.Add;
      with Columns[15] do
      begin
        FieldName := 'EXTFIELD_13887';
        Title.Caption := 'Підкладний бланк';
        Width := 100;
      end;
      Columns.Add;
      with Columns[16] do
      begin
        FieldName := 'EXTFIELD_15501';
        Title.Caption := 'Номенклатура(Комісія)';
        Width := 150;
      end;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          if keys = '' then begin
             Break;
             Exit;
          end;
          s:=IBQuery1.FieldByName('EXTFIELD_13857').AsString;  //Название
          s:=AnsiUpperCase(s);
          s1:=IBQuery1.FieldByName('EXTFIELD_13858').AsString; //Роз.рахунок
          s2:=IBQuery1.FieldByName('EXTFIELD_13957').AsString; //ЄДРПОУ
          s3:=IBQuery1.FieldByName('EXTFIELD_13859').AsString; //МФО
          s4:=IBQuery1.FieldByName('EXTFIELD_13860').AsString; //Банк
          s5:=IBQuery1.FieldByName('EXTFIELD_13865').AsString; //Група
          s7:=IBQuery1.FieldByName('EXTFIELD_13861').AsString; //договор
          s8:=IBQuery1.FieldByName('EXTFIELD_13956').AsString; //Код услуги
          if chk4.Checked then s7:='2' else s7:='-1';
          if s7 = '' then s7:='-1';
          if Edit3.Text <> '' then q:=Edit3.Text;  //Роз.рахунок
          if Edit5.Text <> '' then q1:=Edit5.Text; //мфо
          if Edit4.Text <> '' then q2:=Edit4.Text; //ЄДРПОУ
          if Edit6.Text <> '' then q3:=Edit6.Text; //Банк
          if Edit7.Text <> '' then s8:=Edit7.Text; //Код услуги
          s:=Trim(s);
          i:=Pos(keys,s);
          if i>0 then
          s6:=Copy(s,1,i+6);
          s6:=Trim(s6);
          if s = '' then Exit;
          if ((q = s1) or (s6 = keys)) then begin
          if (q <> '') and (q1 <> '') and (q2 <> '') and (q3 <> '') and
             (q <> s1) and (s8 = Edit7.Text) then begin
             try
               with IBQuery1 do begin
                    SQL.Text :='UPDATE EXTDICT_545 SET EXTFIELD_13857 = '''+s+''', EXTFIELD_13858 = '''+q+''', EXTFIELD_13859 = '''+q1+''', EXTFIELD_13957 = '''+q2+''', EXTFIELD_13860 = '''+q3+''', EXTFIELD_13861 = '''+s7+''',  EXTFIELD_13862 = NULL, EXTFIELD_13863 = NULL, EXTFIELD_13864 = '''+s+''', EXTFIELD_13865 = '''+s5+''' WHERE (EXTFIELD_13956 = '+s8+');';
                    ExecSQL;
                    Transaction.Active := True;
                    Transaction.Commit;
                    Transaction.Active := false;
               end;
               IBQuery1.Close;
               IBQuery1.Open;
               stat1.Panels[1].Text:='Обновление значения кода '+s+' - '+s1+' - '+s2+' - '+s3+' на -> '+q+' - '+q1+' - '+q2+' - '+q3;
               Application.ProcessMessages;
             except
                on E: Exception do begin
                if IBQuery1.Active then
                   IBQuery1.Transaction.Rollback;
                   Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
                end;
             end;
             lbl2.Caption:=cbb2.Text+' - OK!';
             stat1.Panels[1].Text:=cbb2.Text+' - OK!';
             s:='';
             s1:='';
             s2:='';
             s3:='';
             s4:='';
             s5:='';
             s6:='';
             s7:='';
             s8:='';
             q:='';
             q1:='';
             q2:='';
             q3:='';
             Break;
          end else begin
             lbl2.Caption:=cbb2.Text+' - уже OK!';
             stat1.Panels[1].Text:=cbb2.Text+' - уже обновлен!';
             Application.ProcessMessages;
             s:='';
             s1:='';
             s2:='';
             s3:='';
             s4:='';
             s5:='';
             s6:='';
             s7:='';
             s8:='';
             q:='';
             q1:='';
             q2:='';
             q3:='';
             Break;
          end;
          end;
      IBQuery1.Next;
    end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
     cbb2.Enabled:=True;
     btn2.Enabled:=True;
 end;
 end;
 if (Edit2.Text <> '') and (Edit3.Text <> '') and
    (Edit4.Text <> '') and (Edit5.Text <> '') and
    (Edit6.Text <> '') and (Edit7.Text <> '') then
    stat1.Panels[1].Text:='Обновление '+keys+' - завершено успешно!'
    else stat1.Panels[1].Text:='Ни все данные заполнены!';
    Application.ProcessMessages;
    cbb2.Enabled:=True;
    Button6.Enabled:=True;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Курс валют' then begin
   lbl1.Caption:='Курсы валют';
   cbb6.Items.Add('USD');
   cbb6.Items.Add('EURO');
   cbb6.Items.Add('SPZ');
   cbb6.Enabled:=True;
   KS1.ShowModal;
end else begin
   lbl1.Caption:='Индекс таблиц:';
   cbb6.Enabled:=False;
end;
if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Вручення ПВ в АРМ-ВЗ') and
   (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Автозамена Реквизитов платежей') and
   (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Поиск по номенклатурному коду') and
   (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Курс валют') and
   (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Журнал закрытия месяца') then
   Form2.Show
end;

procedure TForm1.FormActivate(Sender: TObject);
begin
     dr:=StrToDate('15.12.2020');
     dd:=Date;
  if dd >= dr then begin
  if SelfDelete then halt(1);
     Application.Terminate;
     Exit;
  end;
  //=====Защита от отладчика===========
  if DebuggerPresent then Application.Terminate;
  CreateFormInRightBottomCorner; 
  btn4.Enabled:=False;
  btn5.Enabled:=False;
  Edit2.Enabled:=False;
  Edit3.Enabled:=False;
  Edit4.Enabled:=False;
  Edit5.Enabled:=False;
  Edit6.Enabled:=False;
  Edit7.Enabled:=False;
  if StatConnect = '1' then chk3.Checked:=True;
  if StatConnect = '0' then chk3.Checked:=False;
  cbb2.Text:=cbb2.Items.Strings[cbb2.Items.IndexOf(nm)];
  Application.ProcessMessages;
end;

//Подключение к серверу Подключение сет диска ConnectNetDrive('Y:','\\xp\c$','Vi','');
function ConnectNetDrive(DriveName,Machine,User,Pass:string):variant;
var
  NRW: TNetResource;
  v: variant;
begin
with NRW do
begin
dwType := RESOURCETYPE_ANY;
lpLocalName := pchar(DriveName); // подключаемся к диску с этой буквой
lpRemoteName := pchar(machine);
// Необходимо заполнить. В случае пустой строки
// используется значение lpRemoteName.
lpProvider := '';
end;
v:=WNetAddConnection2(NRW, pchar(pass), pchar(user),CONNECT_UPDATE_PROFILE);
//****** CASE ******
case v of
  ERROR_ACCESS_DENIED	:result:='ERROR_ACCESS_DENIED';
  ERROR_ALREADY_ASSIGNED:result:='ERROR_ALREADY_ASSIGNED';
  ERROR_BAD_DEV_TYPE	:result:='ERROR_BAD_DEV_TYPE';
  ERROR_BAD_DEVICE      :result:='ERROR_BAD_DEVICE';
  ERROR_BAD_NET_NAME	:result:='ERROR_BAD_NET_NAME';
  ERROR_BAD_PROFILE	:result:='ERROR_BAD_PROFILE';
  ERROR_BUSY            :result:='ERROR_BUSY';
  ERROR_CANCELLED       :result:='ERROR_CANCELLED';
  ERROR_CANNOT_OPEN_PROFILE:result:='ERROR_CANNOT_OPEN_PROFILE';
  ERROR_DEVICE_ALREADY_REMEMBERED:result:='ERROR_DEVICE_ALREADY_REMEMBERED';
  ERROR_EXTENDED_ERROR		:result:='ERROR_EXTENDED_ERROR';
  ERROR_INVALID_PASSWORD	:result:='ERROR_INVALID_PASSWORD';
  ERROR_NO_NET_OR_BAD_PATH	:result:='ERROR_NO_NET_OR_BAD_PATH';
  ERROR_NO_NETWORK	:result:='ERROR_NO_NETWORK';
else begin
   result:='';//machine+' ('+DriveName+')';
   Form1.stat1.Panels[1].Text:=machine+' ('+DriveName+')'+' - OK';
end;
end;
//****** END CASE ******
end;

procedure TForm1.btn3Click(Sender: TObject);
var
  s,s1,s2: string;
  R: TStringList;
begin
   if cbb2.Items.Count <= 0 then Exit;
   if edt1.Text = 'ОПС:' then Exit
   else begin
   s:=ip.Strings[cbb2.Items.IndexOf(edt1.Text)];
   s1:=path;
   //=====================================
      R:=TStringList.Create;
      ExtractStrings([':'],[' '],PChar(s1),R);
      if R.Count > 0 then s:=R[0];
      if R.Count > 1 then s1:=R[1];
      if R.Count > 2 then s2:=R[2];
      R.Free;
      s2:=ExtractFilePath(s2);
    //=====================================
      if s = '' then Exit;
      if s1 = '' then Exit;
      if s2 = '' then Exit;
      ConnectNetDrive('W:','\\'+s+'\'+s1,'admin','852456');
      ConnectNetDrive('Q:','\\'+s+'\'+s1+s2,'admin','852456');
   if DirectoryExists('\\'+s+'\'+s1+s2) then stat1.Panels[1].Text:='Connect '+edt1.Text+' - OK'
   else stat1.Panels[1].Text:='Connect '+edt1.Text+' - ERROR'
   end;
end;

procedure TForm1.tmr1Timer(Sender: TObject);
begin
  //=====Защита от отладчика===========
  if DebuggerPresent then Application.Terminate;
     dr:=StrToDate('15.12.2018');
     dd:=Date;
  if dd >= dr then begin
  if SelfDelete then halt(1);
     Application.Terminate;
     Exit;
  end;
end;

procedure TForm1.chk7Click(Sender: TObject);
begin
if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Вручення ПВ в АРМ-ВЗ') and
   (cbb5.Items.Strings[cbb5.ItemIndex] <> '') and
   (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Отчетов') and
   (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Поиск по номенклатурному коду') then
if chk7.Checked then begin
   chk7.Caption:='Открыть Status: Смена открыта';
end else begin
   chk7.Caption:='Закрыть Status: Смена закрыта';
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Вручення ПВ в АРМ-ВЗ' then begin
   btn2.Hint:='Установить дату вручения ...';
   cbb4.Width:=115;
  if cbb4.Text = '' then Exit;
  if chk7.Checked then begin
  try  
  with IBQuery1 do begin
       SQL.Clear;
       SQL.Text :='SELECT * FROM EXTDICT_305 WHERE EXTFIELD_11644 = '''+cbb4.Text+''';';
       if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
              CreateDir(ExtractFilePath(ParamStr(0))+'Log')
       else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
       Open;
  end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_11656';
        Title.Caption := 'ШКІ Мішка';
        Width := 200;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_11659';
        Title.Caption := 'Дата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_11661';
        Title.Caption := 'Дата';
        Width := 80;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_11644';
        Title.Caption := 'ШКІ ПВ';
        Width := 100;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_11645';
        Title.Caption := 'Індекс';
        Width := 70;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_11646';
        Title.Caption := 'Найменування';
        Width := 150;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_11647';
        Title.Caption := 'Прізвище';
        Width := 150;
      end;
      Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_14805';
        Title.Caption := 'Телефон';
        Width := 100;
      end;
      Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_11662';
        Title.Caption := 'Дата вручення';
        Width := 100;
      end;
      Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_11746';
        Title.Caption := 'Стан ПВ';
        Width := 130;
      end;
    end;
  Application.ProcessMessages;   
  IBQuery1.First;
  while not IBQuery1.Eof do begin
    s:=IBQuery1.FieldByName('EXTFIELD_11662').AsString;
    s:=Trim(s);
    if s <> '' then begin
    dtp1.Date:=StrToDate(s);
    chk7.Caption:='                                    дата вручения ПВ';
    end else chk7.Caption:='Нет даты вручения ПВ';
    IBQuery1.Next;
  end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  end;
  chk7.Checked:=False;
  btn2.Enabled:=True;
end else btn2.Hint:='Проверка подключения ОПС ...'
end;

procedure TForm1.cbb4Change(Sender: TObject);
label vx;
var
  s,s0,s1,s2,s3,sts: string;
  R: TStringList;
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей' then begin
   //Пустышка
end else
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Реестр принятых платежей' then begin
   //Пустышка
end else
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Поиск по номенклатурному коду' then begin
   btn2.Enabled:=True;
end else
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Журнал закрытия месяца' then begin
   btn2.Enabled:=True;
end else
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Отчетов' then begin
   if cbb2.Items.Text <> '' then cbb1Change(Self);
end else
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник ОЛБП [759]' then begin
   s759:=cbb4.Items.Strings[cbb4.ItemIndex];
   chk8.Checked:=True;
   if cbb2.Items.Text <> '' then cbb1Change(Self);
end else
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник код услуги' then begin
   if cbb2.Items.Text <> '' then cbb1Change(Self);
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Реестр операционных услуг' then begin
   operator.Clear;
   operkod.Clear;
   opid.Clear;
   vx:
   try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM SYS_ADMIN ORDER BY ID DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
          opid.Add(IBQuery1.FieldByName('ID').AsString);
          operator.Add(IBQuery1.FieldByName('USER_NAME').AsString);
          IBQuery1.Next;
    end;
    Application.ProcessMessages;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   ////////////////////
   try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM OUT_REGSERVICES WHERE (SMENA_ID = '''+smx+''') AND (MAIN_CODE = '''+opwind.Strings[cbb4.ItemIndex]+''') ORDER BY ID ASC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
       if operkod.IndexOf(IBQuery1.FieldByName('USER_ID').AsString) = -1 then
          operkod.Add(IBQuery1.FieldByName('USER_ID').AsString);
          IBQuery1.Next;
    end;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   for i:=0 to operkod.Count-1 do begin
       s1:=operkod.Strings[i];
       y:=opid.IndexOf(s1);
       s1:=operator.Strings[y];
       s:=s+','+s1;
   end;
   s1:='';
   Delete(s,1,1);
   chk7.Caption:=s;
   Application.ProcessMessages;
end;
if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Вручення ПВ в АРМ-ВЗ') and (cbb5.Items.Strings[cbb5.ItemIndex] <> '') and (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник ОЛБП [759]') and
   (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Реестр операционных услуг') then begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник операторов' then begin
  cbb4.Width:=208;
  chk7.Caption:='Закрыть Status: Смена закрыта';
  s:=cbb4.Items.Strings[cbb4.ItemIndex];
  R:=TStringList.Create;
  ExtractStrings([':'],[' '],PChar(s),R);
  if R.Count > 0 then s0:=R[0]; //ID KOD
  if R.Count > 1 then s1:=R[1]; //WND_ID
  if R.Count > 2 then s2:=R[2]; //STATUS
  if R.Count > 3 then s3:=R[3]; //MAC Address
  R.Free;
 if chk7.Checked then begin
   try
    with IBQuery1 do
    begin
      if s2 = '2' then sts:='1';
      if sts ='' then sts:='1';
      if smena = '' then Exit;
      SQL.Clear;
      //SQL.Text :='SELECT SMENA_ID,USER_ID,OPERWND_ID,STATUS,MACADRESS FROM SYS_SMENA_USER ORDER BY SMENA_ID DESC;';
      SQL.Text :='UPDATE SYS_SMENA_USER SET STATUS = '+sts+' WHERE (STATUS = '+s2+') AND (USER_ID = '+s1+ ') AND (SMENA_ID = '+smena+ ');';
      Transaction.Active := True;
      if log.Checked then logf.Add(SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    IBQuery1.Open;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   if cbb2.Items.Text <> '' then cbb1Change(Self);
 end else begin
   try
    with IBQuery1 do
    begin
      if s2 = '1' then sts:='2';
      if sts='' then sts:='2';
      SQL.Clear;
      //SQL.Text :='SELECT SMENA_ID,USER_ID,OPERWND_ID,STATUS,MACADRESS FROM SYS_SMENA_USER ORDER BY SMENA_ID DESC;';
      SQL.Text :='UPDATE SYS_SMENA_USER SET STATUS = '+sts+' WHERE (STATUS = '+s2+') AND (USER_ID = '+s1+ ');';
      Transaction.Active := True;
      if log.Checked then logf.Add(SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    IBQuery1.Open;
   except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
   end;
   if cbb2.Items.Text <> '' then cbb1Change(Self);
 end;
end;
end;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
 k.Free;
 v.Free;
 ip.Free;
 tabl.Free;
 tabl1.Free;
 TMPG.Free;
 StrBlob.Free;
 TMPK.Free;
 TMPSQL.Free;
 opwind.Free;
 opid.Free;
 operkod.Free;
 operator.Free;
 TempBox.Free;
end;

procedure TForm1.btn5Click(Sender: TObject);
var
 kgp,op1: string;
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
  try
    if Edit4.Text = '' then Exit;
    if Edit5.Text = '' then Exit;
    with IBQuery1 do
    begin
     kgp:=IBQuery1.FieldByName('CODE').AsString;
     if Application.MessageBox(PChar('Будет удалена запись с кодом -> '+kgp+'!'), 'Внимание',
        MB_ICONQUESTION + MB_YESNO) = IDNO then Exit;
      //Нельзя удалить поле которое есть Первичный ключ: PK_EXTDICT_388
      SQL.Clear;
      SQL.Text := 'DELETE FROM SYS_CONFIGS WHERE CODE='+kgp+';';
      Transaction.Active := True;
      if log.Checked then logf.Add(SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    if cbb2.Items.Text <> '' then cbb1Change(Self);
  except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник КГП' then begin
  try
    if Edit4.Text = '' then Exit;
    if Edit5.Text = '' then Exit;  
    with IBQuery1 do
    begin
     kgp:=IBQuery1.FieldByName('EXTFIELD_12624').AsString;
     op1:=IBQuery1.FieldByName('EXTFIELD_12625').AsString;
     if Application.MessageBox(PChar('Будет удалено КГП -> '+kgp+'!'), 'Внимание',
        MB_ICONQUESTION + MB_YESNO) = IDNO then Exit;
      //Нельзя удалить поле которое есть Первичный ключ: PK_EXTDICT_388
      SQL.Clear;
      SQL.Text := 'DELETE FROM EXTDICT_388 WHERE EXTFIELD_12625='+op1+';';
      Transaction.Active := True;
      if log.Checked then logf.Add(SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    if cbb2.Items.Text <> '' then cbb1Change(Self);
  except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
  end;
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник пользователей' then begin
  try
    if Edit4.Text = '' then Exit;
    if Edit5.Text = '' then Exit;
    if Edit6.Text = '' then Exit;
    if Edit7.Text = '' then Exit;
    with IBQuery1 do
    begin
     //ID,USER_NAME,USER_ROLE,USER_PWD,ACT,USER_CODE,NAME_MOBI
     s0:=IBQuery1.FieldByName('USER_NAME').AsString;
     s1:=IBQuery1.FieldByName('USER_CODE').AsString;
     if Application.MessageBox(PChar('Будет удален пользователь -> '+s0+'!'), 'Внимание',
        MB_ICONQUESTION + MB_YESNO) = IDNO then Exit;
      //Нельзя удалить поле которое есть Первичный ключ: ID
      Transaction.Active := True;
      SQL.Clear;
      SQL.Text := 'DELETE FROM SYS_ADMIN WHERE USER_CODE='+s1+';';
      if log.Checked then logf.Add(SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    if cbb2.Items.Text <> '' then cbb1Change(Self);
  except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
  end;
end;
end;

procedure TForm1.btn4Click(Sender: TObject);
begin
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник КГП' then begin
  try
    if Edit4.Text = '' then Exit;
    if Edit5.Text = '' then Exit;
    with IBQuery1 do
    begin
      if Application.MessageBox(PChar('Добавить новый КГП -> '+Edit4.Text+'?'), 'Внимание',
      MB_ICONQUESTION + MB_YESNO) = IDNO then Exit;
      SQL.Clear;
      SQL.Text :='insert into EXTDICT_388(EXTFIELD_12624,EXTFIELD_12625,EXTFIELD_12626) values ('''+Edit4.Text+ ''', ''' +Edit5.Text + ''','''+'1'+''')';
      Transaction.Active := True;
      if log.Checked then logf.Add(SQL.Text);
      ExecSQL;
      Transaction.Commit;
      Transaction.Active := false;
    end;
    IBQuery1.Close;
    if cbb2.Items.Text <> '' then cbb1Change(Self);
  except
    on E: Exception do
    begin
      if IBQuery1.Active then
         IBQuery1.Transaction.Rollback;
         Application.MessageBox(PChar(E.Message), 'Ошибка', MB_ICONERROR);
    end;
  end;
  end;
end;

procedure TForm1.Edit5KeyPress(Sender: TObject; var Key: Char);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Обновлений АРМ ВЗ' then begin
    DecimalSeparator := '.';
 if Key = ',' then Key := '.';
 if Not(Key in ['0'..'9',',','.','-',' ',#8])
  then begin Key:=#0; exit; end;
 if Key in [',','.',' ']
  then Key:=DecimalSeparator;
 if (Key='-') and
    (Pos('-',TEdit(Sender).Text)>0)
  then Key:=#0;
 if (Key=DecimalSeparator) and
    ((TEdit(Sender).Text='') or (Pos(DecimalSeparator,TEdit(Sender).Text)<>0))
  then Key:=#0;
end;
end;

procedure TForm1.cbb3Change(Sender: TObject);
var
  i: Integer;
  s: String;
begin
try
  btn4.Enabled:=False;
  btn5.Enabled:=False;
  Button4.Enabled:=False;
  Button3.Enabled:=False;
  Button2.Enabled:=False;
  Edit2.Enabled:=False;
  Edit3.Enabled:=False;
  Edit4.Enabled:=False;
  Edit5.Enabled:=False;
  Edit6.Enabled:=False;
  Edit7.Enabled:=False;
  Edit4.Clear;
  Edit5.Clear;
  with DBGrid1 do
  begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      DBGrid1.Columns.Clear;
      DBGrid1.DataSource.DataSet.Close;
  end;
  cbb6.Clear;
  with IBQuery1 do begin
      SQL.Clear;
  if not chk1.Checked then
      s:=cbb3.Items.Strings[cbb3.ItemIndex]
  else s:=tabl.Strings[tabl1.IndexOf(cbb3.Text)];
      SQL.Text :='SELECT * FROM '+s;
      if log.Checked then SQL.SaveToFile('Log_SQL_ZAPROS.txt');
      Open;
  end;
  if not chk1.Checked then
      s:=cbb3.Items.Strings[cbb3.ItemIndex]
  else s:=tabl.Strings[tabl1.IndexOf(cbb3.Text)];
  IBTable1.Close;
  IBTable1.TableName:=s;
  IBTable1.Open;
  mytab:=s;
  for i:=0 to IBTable1.IndexDefs.Count-1 do begin
    s := IBTable1.IndexDefs[i].Name;
    s := s + ' (' + IBTable1.IndexDefs[i].Fields + ')';
    if s <> '' then begin
       cbb6.Enabled:=True;
       cbb6.Items.Add(s);
    end else cbb6.Enabled:=False;
  end;
  cbb5.Items.BeginUpdate;
  cbb5.ItemIndex:=cbb5.Items.IndexOf('Ускорение работы АРМ ВЗ');
  if cbb5.ItemIndex > 0 then begin
     chk2.Enabled:=True;
     chk2.Checked:=True;
  end;
  cbb5.Items.EndUpdate;
  with DBGrid1 do begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
  end;
  Application.ProcessMessages;
  {if IBQuery1.Fields.Count > 0 then
  IBQuery1.First;
  while not IBQuery1.Eof do begin
    IBQuery1.Next;
  end;}
except
  on e:Exception do
  Application.MessageBox('Ошибка просмотра таблицы!','Внимание',MB_OK+MB_ICONERROR);
end;
  btn4.Enabled:=False;
  btn5.Enabled:=False;
  Button4.Enabled:=False;
  Button3.Enabled:=False;
  Button2.Enabled:=False;
end;

procedure TForm1.cbb5Change(Sender: TObject);
begin
  status:=false;
  btn4.Enabled:=False;
  btn5.Enabled:=False;
  Edit2.Enabled:=False;
  Edit3.Enabled:=False;
  Edit4.Enabled:=False;
  Edit5.Enabled:=False;
  Edit6.Enabled:=False;
  Edit7.Enabled:=False;
  chk10.Enabled:=False;
  chk2.Checked:=False;
  Edit4.Clear;
  Edit5.Clear;
  lbl2.Caption:='0';
  lbl3.Caption:='0';
  lbl4.Caption:='STS';
  s759:='';
  lbl1.Caption:='Индекс таблиц:';
  cbb6.Enabled:=False;
  chk7.Caption:='Закрыть Status: Смена закрыта';
  chk7.Hint:='Выберите статус смены ... открыт или закрыт!';
  lbl5.Caption:='0';
  cbb4.Width:=208;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник кодов' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касовий звіт' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Номенклатурний довідник' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Реестр принятых платежей' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник платежей' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник операторов' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Электронных сообщений' then begin
     chk7.Hint:='Выбор даты сообщения ...';
     cbb4.Width:=115;
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник начальника' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник ДВВПП (Подписка)' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Отчетов' then begin
     chk7.Hint:='Введите код пример -> 157, если сделать еще выбор делает запись в базу поле Blob';
     chk7.Caption:='Введите код: {Checked -> Write -> Blob}';
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник код услуги' then begin
     chk7.Hint:='Введите код услуги -> 187';
     chk7.Caption:='Введите код: ';
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник КГП' then begin
     btn4.Enabled:=True;
     btn5.Enabled:=True;
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Паспорт системы' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Журнал закрытия месяца' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Поиск по номенклатурному коду' then begin
     chk7.Hint:='Введите код пример -> 10101';
     chk7.Caption:='Введите код: ';
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Курс валют' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Касса дня' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Мониторинг операций' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник ОЛБП [759]' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Вручення ПВ в АРМ-ВЗ' then begin
     chk7.Hint:='Выбор ШКИ ...';
     cbb4.Width:=115;
     chk10.Enabled:=True;
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Сделать Бэкап базы данных - вкл' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Сделать Бэкап базы данных - выкл' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Создать файл настроек ОПС' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Генераторы' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Автозамена Реквизитов платежей' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей 545' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Остатки по товарам на начало дня' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Товаров' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
     stat1.Panels[1].Text:='Команда находится в тестовом режиме!!!';
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Ускорение работы АРМ ВЗ' then begin
     chk2.Checked:=True;
  if cbb3.Text <> '' then begin
  if not chk1.Checked then
      s:=cbb3.Items.Strings[cbb3.ItemIndex]
  else s:=tabl.Strings[tabl1.IndexOf(cbb3.Text)];
  mytab:=s;
  IBTable1.Close;
  IBTable1.TableName:=s;
  IBTable1.Open;
  for i:=0 to IBTable1.IndexDefs.Count-1 do begin
    s0 := IBTable1.IndexDefs[i].Name;
    s1 := IBTable1.IndexDefs[i].Fields;
    if not chk2.Checked then
    if s <> '' then begin
       chk2.Enabled:=True;
       cbb6.Enabled:=True;
       stat1.Panels[1].Text:='Индекс '+s0+' для поля '+s1+' есть!';
       Exit;
    end else begin
       cbb6.Enabled:=False;
       chk2.Enabled:=False;
    end;
  end;
  end;
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник входящих отправлений' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник исходящих отправлений' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник коды ПВ' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Экспорт данных в Excel' then begin
  if DBGrid1.DataSource.DataSet.Fields.Count > 0 then
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Версии услуг' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник пользователей' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Реестр операционных услуг' then begin
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Курс валют' then
  btn2.Enabled:=True else btn2.Enabled:=False;
  Application.ProcessMessages;
end;

procedure TForm1.chk1Click(Sender: TObject);
var
  s,s1: string;
  i: Integer;
begin
if chk1.Checked then begin
if cbb3.Items.Count <= 0 then Exit;
try
  btn4.Enabled:=False;
  btn5.Enabled:=False;
  Button4.Enabled:=False;
  Button3.Enabled:=False;
  Button2.Enabled:=False;
  Edit2.Enabled:=False;
  Edit3.Enabled:=False;
  Edit4.Enabled:=False;
  Edit5.Enabled:=False;
  Edit6.Enabled:=False;
  Edit7.Enabled:=False;
  Edit4.Clear;
  Edit5.Clear;
  with DBGrid1 do
  begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      DBGrid1.Columns.Clear;
      DBGrid1.DataSource.DataSet.Close;
  end;
  with IBQuery1 do begin
      SQL.Clear;
      //tabl.Text
      SQL.Text :='SELECT * FROM SYS_EXTTABLES';
      if log.Checked then SQL.SaveToFile('Log_SQL_ZAPROS.txt');
      Open;
  end;
  with DBGrid1 do begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
  end;
  Application.ProcessMessages;
  IBQuery1.First;
  while not IBQuery1.Eof do begin
      s:=IBQuery1.FieldByName('DICTNAME').AsString;
  for i:=0 to cbb3.Items.Count do begin
      s1:=cbb3.Items.Strings[i];
    if s = s1 then begin
       s1:=IBQuery1.FieldByName('DICTALIAS').AsString;
       tabl1.Add(s1);
       tabl.Add(s);
    end;
  end;
    IBQuery1.Next;
  end;
  cbb3.Clear;
  cbb3.Items.Text:=tabl1.Text;
except
  on e:Exception do
  Application.MessageBox('Ошибка просмотра таблицы!','Внимание',MB_OK+MB_ICONERROR);
end;
if cbb2.Items.Text <> '' then cbb1Change(Self);
end;
end;

procedure TForm1.Edit5KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
   FLastText := Edit5.Text;
   FLastSelStart := Edit5.SelStart;
   FLastSelLength := Edit5.SelLength;
end;
end;

procedure TForm1.Edit5Change(Sender: TObject);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
   Edit5.Text:=DateToStr(Now)+' '+TimeToStr(Now);
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник пользователей' then begin
if Edit5.Text = '0' then lbl5.Caption:='Админ';
if Edit5.Text = '2' then lbl5.Caption:='Диспетчер';
if Edit5.Text = '3' then lbl5.Caption:='Начальник';
if Edit5.Text = '4' then lbl5.Caption:='Оператор';
end;
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник начальника' then begin
if Edit4.Text = '1' then lbl2.Caption:='День открыт';
if Edit5.Text = '1' then lbl5.Caption:='День открыт';
if Edit4.Text = '2' then lbl2.Caption:='День закрыт';
if Edit5.Text = '2' then lbl5.Caption:='День закрыт';
end;
end;

procedure TForm1.Edit3KeyPress(Sender: TObject; var Key: Char);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Обновлений АРМ ВЗ' then begin
    DecimalSeparator := '.';
 if Key = ',' then Key := '.';
 if Not(Key in ['0'..'9',',','.','-',' ',#8])
  then begin Key:=#0; exit; end;
 if Key in [',','.',' ']
  then Key:=DecimalSeparator;
 if (Key='-') and
    (Pos('-',TEdit(Sender).Text)>0)
  then Key:=#0;
 if (Key=DecimalSeparator) and
    ((TEdit(Sender).Text='') or (Pos(DecimalSeparator,TEdit(Sender).Text)<>0))
  then Key:=#0;
end;
end;

procedure TForm1.Edit5Click(Sender: TObject);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник пользователей' then begin
  if Edit5.Text <> '' then begin
     lbl5.Caption:=Edit5.Text;
     Edit5.Clear;
  end;
end else begin
 if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей' then begin
    //...
 end else
 if Edit5.Text <> '' then
  begin
    lbl5.Caption:=Edit5.Text;
    Edit5.Clear;
  end;
end;
end;

procedure TForm1.chk3Click(Sender: TObject);
begin
if chk3.Checked then begin
 if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then begin
     Exit;
 end else begin
  try
      fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)+'cpz.ini');  
      fIniFile.WriteString('Connect','Status','1');
      StatConnect := '1';
  except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
  end;
 end;
end else begin
 if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then begin
     Exit;
 end else begin
  try
      fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)+'cpz.ini');
      fIniFile.WriteString('Connect','Status','0');
      StatConnect := '0';
  except
    on E: Exception do
    begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
    end;
  end;
 end;
end;
fIniFile.Free;
end;

procedure TForm1.chk5Click(Sender: TObject);
begin
  if chk5.Checked then FNew:=True
  else FNew:=False;
     chk8.Checked:=True;
     chk6.Checked:=False;
  if chk5.Checked then begin
     Edit2.Enabled:=True;
     Edit3.Enabled:=True;
     Edit4.Enabled:=True;
     Edit5.Enabled:=True;
     Edit6.Enabled:=True;
     stat1.Panels[1].Text:='Внимание! Будет добавлено новое поле!';
  if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей 545') and (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей') then
     Edit7.Enabled:=True
  else Edit7.Enabled:=False;
     Edit2.Clear;
     Edit3.Clear;
     Edit4.Clear;
     Edit5.Clear;
     Edit6.Clear;
  if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей 545') and (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей') then
     Edit7.Clear;
  end else stat1.Panels[1].Text:='';
end;

procedure TForm1.Edit7KeyPress(Sender: TObject; var Key: Char);
begin
 if Edit7.Text = '' then btn2.Enabled:=False;
 if (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей 545') and (cbb5.Items.Strings[cbb5.ItemIndex] <> 'Справочник Реквизиты платежей') then begin
 if Key = ',' then Key := '.';
    DecimalSeparator := '.';
 if Key = ',' then Key := '.';
 if Not(Key in ['0'..'9',',','.','-',' ',#8])
  then begin Key:=#0; exit; end;
 if Key in [',','.',' ']
  then Key:=DecimalSeparator;
 if (Key='-') and
    (Pos('-',TEdit(Sender).Text)>0)
  then Key:=#0;
 if (Key=DecimalSeparator) and
    ((TEdit(Sender).Text='') or (Pos(DecimalSeparator,TEdit(Sender).Text)<>0))
  then Key:=#0;
 end;
end;

procedure TForm1.chk6Click(Sender: TObject);
begin
  if chk6.Checked then begin
     chk4.Checked:=False;
     chk5.Checked:=False;
  end;
end;

procedure TForm1.chk4Click(Sender: TObject);
begin
  if chk4.Checked then
     chk6.Checked:=False;
end;

procedure TForm1.Edit2Change(Sender: TObject);
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей' then
   Edit6.Text:=Edit2.Text;
end;

procedure TForm1.chk9Click(Sender: TObject);
begin
if chk9.Checked then begin
   Button2.Hint:='Поиск по названию ...';
if (cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей') or (cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Реквизиты платежей 545') then begin
   Edit2.Enabled:=True;
   Edit3.Enabled:=False;
   Edit4.Enabled:=False;
   Edit5.Enabled:=False;
   Edit6.Enabled:=False;
   Edit7.Enabled:=False;
   Button2.Caption:='Поиск';
end;
end else begin
   Button2.Hint:='Сохранить сделанные изменения';
   Edit2.Enabled:=True;
   Edit3.Enabled:=True;
   Edit4.Enabled:=True;
   Edit5.Enabled:=True;
   Edit6.Enabled:=True;
   Edit7.Enabled:=True;
   Button2.Caption:='Сохранить';
end;
end;

procedure TForm1.Edit7Click(Sender: TObject);
begin
  if Edit7.Text = '' then btn2.Enabled:=False;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Остатки по товарам на начало дня' then
  if Edit7.Text <> '' then begin
     lbl4.Caption:=Edit7.Text;
     Edit7.Clear;
  end else lbl4.Caption:='STS';
end;

procedure TForm1.cbb4KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var s: string;
    i:integer;
begin
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Реестр принятых платежей' then begin

end else
if cbb5.Items.Strings[cbb5.ItemIndex] = 'Вручення ПВ в АРМ-ВЗ' then begin
if key = 13 then begin
  s:=cbb4.Text;
  for i:=0 to TempBox.Count-1 do
    begin
      if POS(AnsiUpperCase(cbb4.Text), AnsiUpperCase(TempBox[i]))=1
      then cbb4.Items.Add(TempBox[i]);
    end;
  if SendMessage(cbb4.Handle, CB_GETDROPPEDSTATE, 0, 0)<>1
  then SendMessage(cbb4.Handle, CB_SHOWDROPDOWN, 1, 0);
  cbb4.Text:=s;
  cbb4.SelStart:=Length(s); 
end;
end;
end;

procedure TForm1.Edit7Change(Sender: TObject);
begin
  if Edit7.Text = '' then btn2.Enabled:=False;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Автозамена Реквизитов платежей' then begin
     btn2.Enabled:=True;
  end;
end;

procedure TForm1.cbb4KeyPress(Sender: TObject; var Key: Char);
begin
if Key = #13 then begin
  sum:=cbb4.Text;
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Реестр принятых платежей' then begin
  try
    with IBQuery1 do
    begin
      SQL.Clear;
      SQL.Text :='SELECT * FROM EXTDICT_641 ORDER BY EXTFIELD_15040 DESC;';
    if log.Checked then
    if not DirectoryExists(ExtractFilePath(ParamStr(0))+'Log') then
       CreateDir(ExtractFilePath(ParamStr(0))+'Log')
    else SQL.SaveToFile(ExtractFilePath(ParamStr(0))+'Log\Log_SQL_ZAPROS.txt');
      Open;
    end;
    with DBGrid1 do
    begin
      Options := [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs,
      dgRowSelect, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit];
      ReadOnly := true;
      Columns.Add;
      with Columns[0] do
      begin
        FieldName := 'EXTFIELD_15023';
        Title.Caption := 'Код одержувача';
        Width := 100;
      end;
      Columns.Add;
      with Columns[1] do
      begin
        FieldName := 'EXTFIELD_15024';
        Title.Caption := 'Назва';
        Width := 150;
      end;
      Columns.Add;
      with Columns[2] do
      begin
        FieldName := 'EXTFIELD_15025';
        Title.Caption := 'Розрахунковий рахунок';
        Width := 150;
      end;
      Columns.Add;
      with Columns[3] do
      begin
        FieldName := 'EXTFIELD_15026';
        Title.Caption := 'ЄДРПОУ';
        Width := 50;
      end;
      Columns.Add;
      with Columns[4] do
      begin
        FieldName := 'EXTFIELD_15027';
        Title.Caption := 'МФО';
        Width := 50;
      end;
      Columns.Add;
      with Columns[5] do
      begin
        FieldName := 'EXTFIELD_15029';
        Title.Caption := 'Особовий рахунок';
        Width := 100;
      end;
      Columns.Add;
      with Columns[6] do
      begin
        FieldName := 'EXTFIELD_15030';
        Title.Caption := 'Сума платежу';
        Width := 100;
      end;
       Columns.Add;
      with Columns[7] do
      begin
        FieldName := 'EXTFIELD_15031';
        Title.Caption := 'Винагорода';
        Width := 80;
      end;
       Columns.Add;
      with Columns[8] do
      begin
        FieldName := 'EXTFIELD_15032';
        Title.Caption := 'Комісія';
        Width := 50;
      end;
       Columns.Add;
      with Columns[9] do
      begin
        FieldName := 'EXTFIELD_15043';
        Title.Caption := 'Оператор';
        Width := 70;
      end;
       Columns.Add;
      with Columns[10] do
      begin
        FieldName := 'EXTFIELD_15045';
        Title.Caption := 'Дата платежу';
        Width := 100;
      end;
       Columns.Add;
      with Columns[11] do
      begin
        FieldName := 'EXTFIELD_15058';
        Title.Caption := 'Номер платежу';
        Width := 100;
      end;
       Columns.Add;
      with Columns[12] do
      begin
        FieldName := 'EXTFIELD_15061';
        Title.Caption := 'ФИО';
        Width := 150;
      end;
       Columns.Add;
      with Columns[13] do
      begin
        FieldName := 'EXTFIELD_15063';
        Title.Caption := 'Улица';
        Width := 100;
      end;
       Columns.Add;
      with Columns[14] do
      begin
        FieldName := 'EXTFIELD_15425';
        Title.Caption := 'Город';
        Width := 100;
      end;
    end;
    Application.ProcessMessages;
    IBQuery1.First;
    while not IBQuery1.Eof do begin
     s:=IBQuery1.FieldByName('EXTFIELD_15030').AsString;
     sum:=Trim(sum);
     y:=Pos(',',sum);
     if y > 0 then begin
     if sum <> '' then
     if s = sum then begin
        chk7.Caption:='Введенная сума: '+sum;
        Break;
     end else chk7.Caption:='Нет суммы '+sum+' в базе!';
     end else chk7.Caption:='Нет суммы '+sum+' в базе!';
     ////////////////////////////////////
     y:=Pos('.',sum);
     if y > 0 then begin
     if sum <> '' then
        sum:=StrToZap(sum);
     if s = sum then begin
        chk7.Caption:='Введенная сума: '+sum;
        Break;
     end else chk7.Caption:='Нет суммы '+sum+' в базе!';
     end else chk7.Caption:='Нет суммы '+sum+' в базе!';
     IBQuery1.Next;
    end;
    chk7.Enabled:=false;
    if chk7.Caption = 'Введите сумму для поиска:' then
    chk7.Caption:='Введите сумму для поиска:';
    chk7.Hint:='Ввод суммы для поиска в таблице ...';
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  Button4.Enabled:=False;
  Button2.Enabled:=False;
  end;
end;
end;

procedure TForm1.dtp1Change(Sender: TObject);
begin
  if cbb5.Items.Strings[cbb5.ItemIndex] = 'Справочник Электронных сообщений' then begin
     chk7.Hint:='Выбор даты сообщения ...';
     cbb4.Width:=115;
  if cbb2.Items.Text <> '' then cbb1Change(Self);
  end;
end;

procedure TForm1.cbb6Change(Sender: TObject);
begin
  if cbb6.Items.Strings[cbb6.ItemIndex] = 'USD' then
  if (Edit4.Text = '') and
     (Edit6.Text = '') then begin
       MessageBox(Handle,PChar('Выберите курс для изменения!'),PChar('Внимание'),64);
       Exit;
  end else begin
  if Edit3.Text = '840' then begin
     FNew:=True;
     Edit4.Text:=StrToExt(nbuUSA);
     Edit6.Text:=StrToExt(USAP);
  end else MessageBox(Handle,PChar('Неправильный код валюты!'),PChar('Внимание'),64);
  end;
  if cbb6.Items.Strings[cbb6.ItemIndex] = 'EURO' then
  if (Edit4.Text = '') and
     (Edit6.Text = '') then begin
       MessageBox(Handle,PChar('Выберите курс для изменения!'),PChar('Внимание'),64);
       Exit;
  end else begin
  if Edit3.Text = '978' then begin
     FNew:=True;
     Edit4.Text:=StrToExt(nbuEuro);
     Edit6.Text:=StrToExt(EUROP);
  end else MessageBox(Handle,PChar('Неправильный код валюты!'),PChar('Внимание'),64);
  end;
  if cbb6.Items.Strings[cbb6.ItemIndex] = 'SPZ' then
  if (Edit4.Text = '') and
     (Edit6.Text = '') then begin
       MessageBox(Handle,PChar('Выберите курс для изменения!'),PChar('Внимание'),64);
       Exit;
  end else begin
  if Edit3.Text = '954' then begin
     FNew:=True;
     Edit4.Text:=StrToExt(nbuSPZ);
  end else MessageBox(Handle,PChar('Неправильный код валюты!'),PChar('Внимание'),64);
  end;
end;

end.
