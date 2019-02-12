unit bindex;

interface

uses Classes, Windows, SysUtils, Zlib, MD5Hash;

type
  TBind = class(TObject)
    private
      FDisk: char;
      FKeyFile: string;
      FSalt: string;

      function FReadSerial(AFileName: string): string;
      procedure FWriteSerial(AFileName, AString: string);
    public
      constructor Create;
      destructor Destroy; override;

      procedure CreateKeyFile(ASerial: string);
      procedure CheckNow;
      function GetDiskSerial: string;

      property Disk: char write FDisk;
      property KeyFile: string write FKeyFile;
      property Salt: string write FSalt;
  end;

implementation

{ TBind }

constructor TBind.Create;
begin
  FDisk := 'C';
  FKeyFile := 'sts.lic';
  FSalt := '';
end;

destructor TBind.Destroy;
begin
  inherited Destroy;
end;

procedure TBind.CreateKeyFile(ASerial: string);
begin
  FWriteSerial(FKeyFile, MD5(ASerial));
end;

procedure TBind.CheckNow;
var
  FileSerial: string;
  DiskSerial: string;
begin
  try
    DiskSerial := GetDiskSerial;
    FileSerial := FReadSerial(FKeyFile);
  finally
    if (FileSerial <> MD5(DiskSerial)) then
    Halt;
  end;
end;

function TBind.GetDiskSerial: string;
var
  VN: array[0..255] of char;
  SN, VW, SW: DWORD;
begin
  GetVolumeInformation(PChar(FDisk + ':\'), VN, SizeOf(VN), @SN, VW, SW, nil, 0);
  Result := MD5(IntToStr(SN) + FSalt);
end;

function TBind.FReadSerial(AFileName: string): string;
var
  FileStream: TFileStream;
  ZlibStream: TDecompressionStream;
  Buffer: array of byte;
const
  BuffSize = 1024;
begin
  try
    FileStream := TFileStream.Create(AFileName, fmOpenRead);
    ZlibStream := TDecompressionStream.Create(FileStream);
    SetLength(Buffer, BuffSize);
    ZlibStream.Read(Buffer[0], BuffSize);
    ZlibStream.Free;
    FileStream.Free;
    Result := Trim(string(Buffer));
  except
    Halt;
  end;
end;

procedure TBind.FWriteSerial(AFileName, AString: string);
var
  FileStream: TFileStream;
  ZlibStream: TCompressionStream;
  Stream: TStringStream;
begin
  try
    Stream := TStringStream.Create(AString);
    FileStream := TFileStream.Create(AFileName, fmCreate);
    ZlibStream := TCompressionStream.Create(clMax, FileStream);
    ZlibStream.CopyFrom(Stream, Stream.Size);
    ZlibStream.Free;
    FileStream.Free;
    Stream.Free;
  except
    Halt;
  end;
end;

end.
