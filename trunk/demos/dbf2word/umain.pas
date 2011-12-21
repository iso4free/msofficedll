unit umain; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, DBGrids, IniFiles, dbf, db, Grids;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    CheckBox1: TCheckBox;
    Datasource1: TDatasource;
    Dbf1: TDbf;
    DBGrid1: TDBGrid;
    Edit1: TEdit;
    Edit2: TEdit;
    Label1: TLabel;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure CheckBox1Change(Sender: TObject);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end; 

var
  Form1: TForm1;
  INI  : TINIFile;
  oldDecSep: Char;
  oldThSep : Char;

implementation

{$R *.lfm}

uses uofficedll, variants;

{ TForm1 }

function currency2Str (value: currency): string;
{*}
const hundreds: array [0..9] of string = ('',' сто',' двісті',' триста',' чотириста',' п''ятсот',
					  ' шістсот',' сімсот',' вісімсот',' дев''ятсот');

tens: array [0..9] of string = ('','',' двадцять',' тридцять',' сорок',' п''ятдесят',' шістдесят',
				' сімдесят',' вісімдесят',' дев''яносто');

ones: array [0..19] of string = ('','','',' три',' чотири',' п''ять',' шість',' сім',' вісім',
				  ' дев''ять',' десять',' одиннадцять',' дванадцять',' тринадцять',
                                  ' чотирнадцять',' п''ятнадцять',' шістнадцять',' сімнадцять',
                                  ' вісімнадцять',' дев''ятнадцять');
razryad: array [0..6] of string = ('',' тисяч',' мільйон',' мільярд',' трильйон',' квадрильйон',
				   ' квінтільйон');
{*}
var s: string; i: integer; val: integer;


function shortNum(s: string; raz: integer): string;
begin
Result:=hundreds[StrToInt(s[1])];
if StrToInt(s)=0 then Exit;
if s[2]<>'1' then begin
Result:=Result+tens[StrToInt(s[2])];
case StrToInt(s[3]) of
{*}
1: if raz=1 then Result:=Result+' одна' else Result:=Result+' одна';
2: if raz=1 then Result:=Result+' дві' else Result:=Result+' дві';
{!}
else Result:=Result+ones[StrToInt(s[3])];
end;
Result:=Result+razryad[raz];
case StrToInt(s[3]) of
{*}
0,5,6,7,8,9: if raz>1 then Result:=Result+'ів';
1: if raz=1 then Result:=Result+'а';
2,3,4: if raz=1 then Result:=Result+'і' else if raz>1 then Result:=Result+'а';
end;
{!}
end else begin
Result:=Result+ones[StrToInt(Copy(s,2,2))];
Result:=Result+razryad[raz];
if raz>1 then Result:=Result+'ів';
end;
end;

begin
{+}
//перевірка, чи сума від'ємна
if value<0 then begin
 Result := 'мінус ';
 value := System.Abs(Value);
end;
{!}
val:=Trunc(value);
{*}
if val=0 then begin Result:='нуль грн. 00 коп.'; Exit; end;
{!}
s:=IntToStr(val); Result:=''; i:=0;
while Length(s)>0 do begin
Result:=shortNum(Copy('00'+s,Length('00'+s)-2,3),i)+Result;
if Length(s)>3 then s:=Copy(s,1,Length(s)-3) else s:='';
inc(i);
end;
s:=IntToStr(Trunc((value-val)*100+0.5));
{*}
if s='0' then s := '00';
Result:=Result+' грн. '+s+' коп.';
{!}
end;


procedure TForm1.FormCreate(Sender: TObject);
begin
   oldDecSep:=DecimalSeparator;
   DecimalSeparator:=',';
   oldThSep:=ThousandSeparator;
   ThousandSeparator:='.';
  //очистити всі поля форми
  Edit1.Clear;
  //Задати стартовий каталог для вибору файла каталог програми
  OpenDialog1.InitialDir:=ExtractFilePath(Application.ExeName);
  //відкрити файл настройок
  INI:=TIniFile.Create(ChangeFileExt(Application.ExeName,'.ini'));
  CheckBox1.Checked:=INI.ReadBool('SECTION1','DOS CODEPAGE',true);
  Edit2.Text:=INI.ReadString('SECTION2','LISTNUM','1');
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  Dbf1.Active:=false;
  INI.Destroy;
  DecimalSeparator  := oldDecSep;
  ThousandSeparator := oldThSep;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  //відкрити діалог вибору файла DBF
  if OpenDialog1.Execute then begin
    //якщо файл вибрано, внести його ім'я в поле на формі
    Edit1.Text:=ExtractFileName(OpenDialog1.FileName);
  end;
  //відкрити DBF
  if OpenDialog1.FileName<>'' then begin
   Dbf1.FilePathFull:=ExtractFilePath(OpenDialog1.FileName);
   Dbf1.TableName:=Edit1.Text;
   Dbf1.Active:=true;
   Button2.Enabled:=true;
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var W : Variant;
   sum: Currency;
   i  : Integer;
begin
 //обнуляємо суму
 sum := 0;
 //Створюємо новий документ з полями зліва 2 см, з інших сторін по 1 см
 NewDocument(w,false);
 PageMargins(2,1,1,1,w);
 //Шапка таблиці вирівнювання по центру жирним шрифтом
 ParagraphAlign(wdAlignParagraphCenter,w);
 FontBold(true,w);
 AddText(Utf8ToAnsi('СПИСОК № '+Edit2.Text+#13),w);
 AddText(Utf8ToAnsi('для зарахувань на карткові рахунки працівників облнаркодиспансера'+#13),w);
 AddText(Utf8ToAnsi('від зарплати за січень 2011 р.'+#13),w);
 FontBold(false,w);
 ParagraphAlign(wdAlignParagraphLeft,w);
 //створюємо таблицю 6 колонок 2 рядки (решта рядків додається в процесі формування)
 //і заповнюємо шапку
 CreateTable(6,2,w);
 AddText(Utf8ToAnsi('№ п/п'),w);
 SetColWidth(20,w);
 GotoRight(1,w);
 AddText(Utf8ToAnsi('Таб. №'),w);
 SetColWidth(30,w);
 GotoRight(1,w);
 AddText(Utf8ToAnsi('№ рах.'),w);
 GotoRight(1,w);
 AddText(Utf8ToAnsi('П.І.П-б.'),w);
 SetColWidth(180,w);
 GotoRight(1,w);
 AddText(Utf8ToAnsi('Ід. код'),w);
 SetColWidth(70,w);
 GotoRight(1,w);
 AddText(Utf8ToAnsi('Сума'),w);
 SetColWidth(90,w);
 GotoRight(1,w);
 //встановлюємо виборку ненульових сум
 Dbf1.Filter:='RLSUM>0';
 Dbf1.Filtered:=true;
 //і йдемо на початок бази
 Dbf1.First;
 //виводимо в циклі вибрані записи в таблицю
 for i:= 1 to Dbf1.ExactRecordCount do begin
  AddText(IntToStr(i)+'.',w);
  GotoRight(1,w);
  AddText(Dbf1.FieldByName('LSTBL').AsString,w);
  GotoRight(1,w);
  AddText(Dbf1.FieldByName('CARD_NO').AsString,w);
  GotoRight(1,w);
  if CheckBox1.Checked then AddText(Utf8ToAnsi(ConsoleToUTF8(Dbf1.FieldByName('FAM').AsString+' '+
                                                      Dbf1.FieldByName('NAME').AsString+' '+
                                                      Dbf1.FieldByName('OT').AsString)),w)
                       else AddText(Dbf1.FieldByName('FAM').AsString+' '+
                                    Dbf1.FieldByName('NAME').AsString+' '+
                                    Dbf1.FieldByName('OT').AsString,w);
  GotoRight(1,w);
  AddText(Dbf1.FieldByName('INN').AsString,w);
  GotoRight(1,w);
  AddText(Utf8ToAnsi(FloatToStrF(Dbf1.FieldByName('RLSUM').AsCurrency,ffNumber,5,2)+' грн.'),w);
  sum:=sum+Dbf1.FieldByName('RLSUM').AsCurrency;
  GotoRight(1,w);
  Dbf1.Next;
  Application.ProcessMessages;
 end;
 //виводимо підсумки
  MergeCellsR(5,w);
  AddText(Utf8ToAnsi('ВСЬОГО: '+FloatToStrF(sum,ffNumber,10,2)+' грн. ('+currency2Str(sum)+')'),w);
  ExitTable(w);
  //підписуємо відомість
  AddText(#13#13,w);
  AddTabPosition(1,w);
  AddTabPosition(12,w);
  AddText(Utf8ToAnsi(#9+'Головний лікар'+#9+'Ю.С.ШУГАЛЕЙ'+#13#13),w);
  AddText(Utf8ToAnsi(#9+'Головний бухгалтер'+#9+'Л.О.БАЛАБАН'),w);
  InsertFooter(Utf8ToAnsi('Демонстраційний приклад формування документу MS Word програмно на основі файла DBF'),w);
  //зберігаємо документ
  SaveDocAs(Utf8ToAnsi(ExtractFilePath(Application.ExeName)+'список'+Edit2.Text+'.doc'),w);
  //закриваємо ворд
  CloseWord(w);
  w:=Unassigned;
  //знімаємо фільтр з DBF
  Dbf1.Filtered:=false;
  ShowMessage('Список сформовано і збережено в папку програми з іменем "'+'список'+Edit2.Text+'.doc"');
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
 INI.WriteBool('SECTION1','DOS CODEPAGE',CheckBox1.Checked);
 INI.WriteString('SECTION2','LISTNUM',Edit2.Text);
 Application.Terminate;
end;

procedure TForm1.CheckBox1Change(Sender: TObject);
begin
 DBGrid1.Refresh;
end;

procedure TForm1.DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
F: TField;
S: ShortString;
begin
f := Column.Field;
s := F.DisplayText;
if (F.FieldName = 'FAM') or (f.FieldName = 'NAME') or (f.FieldName='OT') then
begin
  DBGrid1.Canvas.Brush.Color := clWindow;
  DBGrid1.Canvas.Font.Color  := clBlack;
  DBGrid1.Canvas.FillRect(Rect);
  if CheckBox1.Checked then DBGrid1.Canvas.TextOut(Rect.Left, Rect.Top, ConsoleToUTF8(s))
                       else DBGrid1.Canvas.TextOut(Rect.Left, Rect.Top, AnsiToUtf8(s));
end
else if (f.FieldName='RLSUM') then begin
  DBGrid1.Canvas.Brush.Color := clGreen;
  DBGrid1.Canvas.Font.Color  := clYellow;
  DBGrid1.Canvas.FillRect(Rect);
  DBGrid1.Canvas.TextOut(Rect.Left, Rect.Top, FloatToStrF(F.AsCurrency,ffCurrency,14,2)+' грн.');
end;
end;

end.

