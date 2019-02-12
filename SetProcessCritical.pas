program SetProcessCritical;

{$APPTYPE CONSOLE}

uses
  Windows;

const
  ntdll = 'NTDLL.DLL';

type
  NTSTATUS = ULONG;
  HANDLE = ULONG;
  PROCESS_INFORMATION_CLASS = ULONG;

  function RtlAdjustPrivilege(Privilege: ULONG; Enable: BOOL; CurrentThread: BOOL; var Enabled: PBOOL): DWORD; stdcall; external 'ntdll.dll';
  function NtSetInformationProcess(ProcessHandle: HANDLE; ProcessInformationClass: PROCESS_INFORMATION_CLASS; ProcessInformation: Pointer; ProcessInformationLength: ULONG): NTSTATUS; stdcall; external ntdll;

var
  Cmd: string[10];
  bl: PBOOL;
  BreakOnTermination: ULONG;
  HRES: HRESULT;
begin
  if not RtlAdjustPrivilege($14, True, True, bl) = 0 then
  begin
    writeln('Unable to enable SeDebugPrivilege. Make sure you are running this program as administrator.');
    Exit;
  end;
  writeln('Commands:' + #13#10 +
          'on - Set the current process as critical process.' + #13#10 +
          'off - Cancel the critical process status.' + #13#10 +
          'exit - Terminate the current process.');
  while True do
  begin
    Readln(cmd);
    if Cmd = 'on' then
    begin
      BreakOnTermination := 1;
      HRES := NtSetInformationProcess(GetCurrentProcess(), $1D , @BreakOnTermination, SizeOf(BreakOnTermination));
      if HRES = S_OK then
        writeln('Successfully set the current process as critical process.')
      else
        writeln('Error: Unable to set the current process as critical process.')
    end
    else if Cmd = 'off' then
    begin
      BreakOnTermination := 0;
      HRES := NtSetInformationProcess(GetCurrentProcess(), $1D , @BreakOnTermination, SizeOf(BreakOnTermination));
      if HRES = S_OK then
        writeln('Successfully canceled critical process status.')
      else
        writeln('Error: Unable to cancel critical process status.')
    end
    else if Cmd = 'exit' then
    begin
      Break;
    end;
  end;
  BreakOnTermination := 0;
  NtSetInformationProcess(GetCurrentProcess(), $1D , @BreakOnTermination, SizeOf(BreakOnTermination));
end.