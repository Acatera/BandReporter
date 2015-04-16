unit Acatera.Utils;

interface

function CenterStr(Str: string; Width: SmallInt): string;
function CenterStr2(Str: string; Width: SmallInt): string;
function LPad(Str: string; Width: SmallInt): string;
function RPad(Str: string; Width: SmallInt): string;

function VarIsNumber(Value: Variant): Boolean;
function VarIsString(Value: Variant): Boolean;

implementation

uses
  SysUtils, StrUtils, Variants;

function CenterStr(Str: string; Width: SmallInt): string;
begin
  if (Length(Str) > Width) then
    Result := Copy(Str, 1, Width)
  else begin
    Result := Format('%*s', [Width, Str + Format('%-*s', [(Width - Length(Str)) div 2, ''])]);
  end;
end;

function CenterStr2(Str: string; Width: SmallInt): string;
var
  PadL: SmallInt;
  PadR: SmallInt;
  Len: SmallInt;
begin
  if (Length(Str) > Width) then
    Result := Copy(Str, 1, Width)
  else begin
    Len := Length(Str);
    PadL := (Width - Len) div 2;
    if ((PadL * 2)<(Width - Len)) then
      PadR := Width - Len - PadL
    else
      PadR := PadL;
    Result := Format('%*s%s%*s', [PadL, ' ', Str, PadR, ' ']);
  end;
end;

function LPad(Str: string; Width: SmallInt): string;
begin
  if (Length(Str) > Width) then
    Str := Copy(Str, 1, Width);
  Result := Format('%*s', [Width, Str]);
end;

function RPad(Str: string; Width: SmallInt): string;
begin
  if (Length(Str) > Width) then
    Str := Copy(Str, 1, Width);
  Result := Format('%-*s', [Width, Str]);
end;

function VarIsNumber(Value: Variant): Boolean;
var
  VariantType: Integer;
begin
  VariantType := VarType(Value) and VarTypeMask;

  case VariantType of
    varSmallInt, varInteger, varSingle, varDouble, varCurrency, varByte, varWord, varLongWord, varInt64:
      Result := True;
  else
    Result := False;
  end;
end;

function VarIsString(Value: Variant): Boolean;
var
  VariantType: Integer;
begin
  VariantType := VarType(Value) and VarTypeMask;

  case VariantType of
    varOleStr, varStrArg, varString:
      Result := True;
  else
    Result := False;
  end;
end;

end.

