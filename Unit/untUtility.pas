unit untUtility;

interface
uses
  Windows,
  Classes,Dialogs,
  SysUtils;

procedure SplitString(str: string; Delimiter: string; var strList: TStringList);
function CutLeftString(str: string; Num: Integer):string; overload;
function CutLeftString(str: string; LeftStr: string):string; overload;
function RemoveQuotation(str: string; QuotationMark: string = '"'): string;
function GetTempDirectory: string;
function GetWinCaption(Handle: Integer):string;
function GetWinClassName(Handle: Integer):string;

implementation

procedure SplitString(str: string; Delimiter: string; var strList: TStringList);
var
  DelimiterPos: Integer;
begin
  //Count := ExtractStrings([Delimiter], [' '], PChar(str), strList);

  strList.Clear;
  if str = '' then Exit;

  DelimiterPos := pos(Delimiter, str);
  while DelimiterPos > 0 do
  begin
    strList.Add(Copy(str, 1, DelimiterPos - 1));
    Delete(str, 1, DelimiterPos + Length(Delimiter) - 1);
    DelimiterPos := Pos(Delimiter, str);
  end;
  strList.Add(str);
end;

function CutLeftString(str: string; Num: Integer):string; overload;
begin
  Result := str;

  if Num >= Length(str) then Exit;

  Result := Copy(str, Num + 1, Length(str) - Num);
end;

function CutLeftString(str: string; LeftStr: string):string; overload;
begin
  Result := str;

  if Length(LeftStr) >= Length(str) then Exit;

  Result := Copy(str, Length(LeftStr) + 1, Length(str) - Length(LeftStr));
end;

function RemoveQuotation(str: string; QuotationMark: string = '"'): string;
var
  Len: Integer;
begin
  Result := str;

  // 先净身
  str := Trim(str);

  // 取得引号的长度
  Len := Length(QuotationMark);

  // 长度不够，直接退出
  if Length(str) <= 2 * Len then Exit;

  // 如果前后不是引号，直接退出
  if (Copy(str, 1, Len) <> QuotationMark)
    or (Copy(str, Length(str) - Len + 1, Len) <> QuotationMark) then
    Exit;

  // 去除前后的引号
  str := Copy(str, Len + 1, Length(str) - 2 * Len);

  Result := str;
end;

function GetTempDirectory: String;
var
  TempDir: array[0..255] of Char;
begin
  GetTempPath(255, @TempDir);
  Result := StrPas(TempDir);
end;

function GetWinCaption(Handle: Integer):string;
var
  Caption: array[0..MAX_PATH] of Char;
begin
  Result := '';

  if GetWindowText(Handle, Caption, SizeOf(Caption) - 1) = 0 then Exit;

  Result := Caption;
end;

function GetWinClassName(Handle: Integer):string;
var
  ClassName: array[0..MAX_PATH] of Char;
begin
  Result := '';

  if (GetClassName(Handle, ClassName, SizeOf(ClassName)) = 0) then Exit;

  Result := ClassName;
end;
end.
