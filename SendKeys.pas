unit SendKeys;

 

interface

 

uses

Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs;

 

type

TSendKeys = class(TComponent)

private

   fhandle:HWND;

   L:Longint;

   fchild: boolean;

   fChildText: string;

   procedure SetIsChildWindow(const Value: boolean);

   procedure SetChildText(const Value: string);

   procedure SetWindowHandle(const Value: HWND);

protected

 

public

 

published

   Procedure GetWindowHandle(Text:String);

   Procedure SendKeys(buffer:string);

   Property WindowHandle:HWND read fhandle write SetWindowHandle;

   Property IsChildWindow:boolean read fchild write SetIsChildWindow;

   Property ChildWindowText:string read fChildText write SetChildText;

end;

 

procedure Register;

 

implementation

 

var temps:string;{é utilizado para ser acessivel pelas funcs q sao

                 utilizadas como callbacks}

   HTemp:Hwnd;

   ChildText:string;

   ChildWindow:boolean;

 

procedure Register;

begin

RegisterComponents('Standard', [TSendKeys]);

end;

 

{ TSendKeys }

 

 

function PRVGetChildHandle(H:HWND; L: Integer): LongBool;

var p:pchar;

   I:integer;

   s:string;

begin

I:=length(ChildText)+2;

GetMem(p,i+1);

SendMessage(H,WM_GetText,i,integer(p));

s:=strpcopy(p,s);

if pos(ChildText,s)<>0 then

  begin

    HTemp:=H;

    Result:=False

  end else

   Result:=True;

FreeMem(p);

end;

 

function PRVSendKeys(H: HWND; L: Integer): LongBool;stdcall;

var s:string;

   i:integer;

begin

i:=length(temps);

if i<>0 then

begin

   SetLength(s,i+2);

   GetWindowText(H, pchar(s),i+2);

   if Pos(temps,string(s))<>0 then

   begin

     Result:=false;

     if ChildWindow then

       EnumChildWindows(H,@PRVGetChildHandle,L)

     else

       HTemp:=H;

   end

   else

     Result:=True;

end else

   Result:=False;

end;

 

procedure TSendKeys.GetWindowHandle(Text: String);

begin

temps:=Text;

ChildText:=fChildText;

ChildWindow:=fChild;

EnumWindows(@PRVSendKeys,L);

fHandle:=HTemp;

end;

 

 

procedure TSendKeys.SendKeys(buffer: string);

var i:integer;

   w:word;

   D:DWORD;

   P:^DWORD;

begin

P:=@D;

SystemParametersInfo(                      //get flashing timeout on win98

        SPI_GETFOREGROUNDLOCKTIMEOUT,

        0,

        P,

        0);

SetForeGroundWindow(fHandle);

for i:=1 to length(buffer) do

begin

   w:=VkKeyScan(buffer[i]);

   keybd_event(w,0,0,0);

   keybd_event(w,0,KEYEVENTF_KEYUP,0);

end;

SystemParametersInfo(                     //set flashing TimeOut=0

        SPI_SETFOREGROUNDLOCKTIMEOUT,

        0,

        nil,

        0);

SetForegroundWindow(TWinControl(TComponent(Self).Owner).Handle);

                                           //->typecast working...

SystemParametersInfo(                     //set flashing TimeOut=previous value

        SPI_SETFOREGROUNDLOCKTIMEOUT,

        D,

        nil,

        0);

end;

 

procedure TSendKeys.SetChildText(const Value: string);

begin

fChildText := Value;

end;

 

procedure TSendKeys.SetIsChildWindow(const Value: boolean);

begin

fchild := Value;

end;

 

procedure TSendKeys.SetWindowHandle(const Value:HWND);

begin

fHandle:=WindowHandle;

end;

 

 

 

end.