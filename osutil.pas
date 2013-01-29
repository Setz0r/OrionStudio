unit osutil;

interface

uses SysUtils;

  function StrToReal(chaine: string): Real;
  function RealToStr(inreal: Extended; digits: Byte): String;
  function Parse(str: String; delimiter: Char; param_num: Integer): String;
  function ParseString(instring:string;datastr:string) : string;
  function compareInvoices(Item1:Pointer;Item2:Pointer) : Integer;

type
  TOrionInvoice = record
    invoiceno : string;
    date : string;
    natnum : string;
    tx : string;
    total : real;
  end;

var
  custa : string;
  custb : string;
  custc : string;
  custd : string;
  custe : string;

implementation

{ Converst a String to a Real Integer }
function StrToReal(chaine: string): Real;
var
  r: Real;
  Pos: Integer;
begin
  Val(chaine, r, Pos);
  if Pos > 0 then Val(Copy(chaine, 1, Pos - 1), r, Pos);
  Result := r;
end;

{ Converts a Real Integer to a String }
function RealToStr(inreal: Extended; digits: Byte): string;
var
  S: string;
begin
  Str(inreal: 0: digits, S);
  realToStr := S;
end;

{ Takes a delimited string and returns the n'th field }
function Parse(str: String; delimiter: Char; param_num: Integer): String;
var
  c, x, y : LongInt;
begin
  x := 1;   // param number that we're currently 'in'
  y := 0;   // position of previous delimiter
  for c := 1 to Length(str) do
    if str[c] = delimiter then  // this char is a delimiter
    begin
      if x = param_num then
        Break;
      inc(x);
      y := c;
    end;
  if x = param_num then
    Result := Copy(str, y + 1, c - y - 1)
  else
    Result := '';
end;

{ Takes a string and replaces tokens with data }
function ParseString(instring:string;datastr:string) : string;
var
  newstring : string;
begin
{
 'OP'   F1
 'OQ'   F2
 'OR'   F3
 'OS'   F4
 '[15~' F5
 '[17~' F6
 '[18~' F7
 '[19~' F8
 '[20~' F9
 '[21~' F10
 '[23~' F11
 '[24~' F12
}
  newstring := StringReplace(instring,'[tab]',#9,[rfReplaceAll]);
  newstring := StringReplace(newstring,'[tab-]',#27#9,[rfReplaceAll]);
  newstring := StringReplace(newstring,'[enter]',#13#10,[rfReplaceAll]);
  newstring := StringReplace(newstring,'[f1]',#27'OP',[rfReplaceAll]);
  newstring := StringReplace(newstring,'[f2]',#27'OQ',[rfReplaceAll]);
  newstring := StringReplace(newstring,'[f3]',#27'OR',[rfReplaceAll]);
  newstring := StringReplace(newstring,'[f4]',#27'OS',[rfReplaceAll]);
  newstring := StringReplace(newstring,'[f8]',#27'[19~',[rfReplaceAll]);
  newstring := StringReplace(newstring,'[f12]',#27'[24~',[rfReplaceAll]);
  if (Length(datastr) > 0) then
  begin
    newstring := StringReplace(newstring,'{param1}',Parse(datastr,',',2),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param2}',Parse(datastr,',',3),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param3}',Parse(datastr,',',4),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param4}',Parse(datastr,',',5),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param5}',Parse(datastr,',',6),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param6}',Parse(datastr,',',7),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param7}',Parse(datastr,',',8),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param8}',Parse(datastr,',',9),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param9}',Parse(datastr,',',10),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param10}',Parse(datastr,',',11),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param11}',Parse(datastr,',',12),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param12}',Parse(datastr,',',13),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param13}',Parse(datastr,',',14),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param14}',Parse(datastr,',',15),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{param15}',Parse(datastr,',',16),[rfReplaceAll]);
    newstring := StringReplace(newstring,'{custa}',custa,[rfReplaceAll]);
    newstring := StringReplace(newstring,'{custb}',custb,[rfReplaceAll]);
    newstring := StringReplace(newstring,'{custc}',custc,[rfReplaceAll]);
    newstring := StringReplace(newstring,'{custd}',custd,[rfReplaceAll]);
    newstring := StringReplace(newstring,'{custe}',custe,[rfReplaceAll]);
  end;
  ParseString := newstring;
end;

function compareInvoices(Item1:Pointer;Item2:Pointer) : Integer;
var
  inva,invb : ^TOrionInvoice;
begin
  Result := 0;
  inva := Item1;
  invb := Item2;
  if (inva.total < invb.total) then Result := 1;
  if (inva.total = invb.total) then Result := 0;
  if (inva.total > invb.total) then Result := -1;
end;

end.
