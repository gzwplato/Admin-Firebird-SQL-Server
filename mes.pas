unit mes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IniFiles, StdCtrls;

type
  TForm2 = class(TForm)
    cbb1: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure cbb1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  p:string;
  SR:TSearchRec;
  FindRes: Integer; // ���������� ��� ������ ���������� ������
  bases,bs: string;

implementation

uses sts;

{$R *.dfm}

procedure TForm2.FormCreate(Sender: TObject);
var
  fIniFile: TIniFile;
begin
 if not FileExists(ExtractFilePath(ParamStr(0))+'Config.ini') then
  begin
     Exit;
  end else
  begin
  try
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'Config.ini');
    try
      bases := fIniFile.ReadString('Base', 'Path', '');
      p := ExtractFilePath(bases);
    finally
      fIniFile.Free;
    end;
    if bases = '' then
     begin
       MessageBox(0,'��� ����� ���� ������!!!','��������',64);
       Exit;
     end;
    // ������� ������� ������ � ������ ������
        FindRes := FindFirst(p+'\*.GDB', faAnyFile, SR);
     begin
       while FindRes = 0 do // ���� �� ������� ����� (��������), �� ��������� ����
       begin
         cbb1.Items.Add(SR.Name); // ���������� � ������ ��������
          // ���������� ��������
         FindRes := FindNext(SR); // ����������� ������ �� �������� ��������
       end;
         FindClose(SR); // ��������� �����
     end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), '������',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
  end;
end;

procedure TForm2.cbb1Change(Sender: TObject);
var
  fIniFile: TIniFile;
  FullProgPath: PChar;
begin
  try
    fIniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)
        + 'Config.ini');
    try
      p := ExtractFilePath(bases);
      bs := cbb1.Text;
      p:=p+bs;
      fIniFile.WriteString('Base', 'Path', p);
    finally
      fIniFile.Free;
    end;
      Form2.Close;
      FullProgPath := PChar(Application.ExeName);
      WinExec(FullProgPath, SW_SHOW); // Or better use the CreateProcess function
      Application.Terminate; // or: Close;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), '������',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
end;

end.
