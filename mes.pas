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
  FindRes: Integer; // переменная для записи результата поиска
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
       MessageBox(0,'Нет файла базы данных!!!','Внимание',64);
       Exit;
     end;
    // задание условий поиска и начало поиска
        FindRes := FindFirst(p+'\*.GDB', faAnyFile, SR);
     begin
       while FindRes = 0 do // пока мы находим файлы (каталоги), то выполнять цикл
       begin
         cbb1.Items.Add(SR.Name); // добавление в список название
          // найденного элемента
         FindRes := FindNext(SR); // продолжение поиска по заданным условиям
       end;
         FindClose(SR); // закрываем поиск
     end;
  except on E: Exception do
     begin
        Application.MessageBox(PChar(E.Message), 'Ошибка',
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
        Application.MessageBox(PChar(E.Message), 'Ошибка',
        MB_ICONERROR + MB_TOPMOST);
        Halt;
     end;
  end;
end;

end.
