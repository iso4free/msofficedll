{ ***************************************************************************
 *                                                                         *
 *   This source is free software; you can redistribute it and/or modify   *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 *   This code is distributed in the hope that it will be useful, but      *
 *   WITHOUT ANY WARRANTY; without even the implied warranty of            *
 *   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU     *
 *   General Public License for more details.                              *
 *                                                                         *
 *   A copy of the GNU General Public License is available on the World    *
 *   Wide Web at <http://www.gnu.org/copyleft/gpl.html>. You can also      *
 *   obtain it by writing to the Free Software Foundation,                 *
 *   Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.        *
 *                                                                         *
 ***************************************************************************
}

unit umsoffice;
{* Модуль для автоматизации создания отчётов в M$ Word XP/2003 в Delphi/Lazarus.
Для использования необходимо в проект подключить модуль uofficedll.pas
(в Lazarus также необходимо подключить модуль Variants и модуль ComObj).}
interface
  uses sysutils,variants,comobj;
 {$I const.inc}

 function CentimetersToPoints(cm : Real) : Real;
 {* Переводит сантиметры в точки}

 procedure NewDocument(var Wrd : Variant; visible : Boolean);
 {* Создает новый документ.  Параметр visible указывает, будет ли отображаться окно Word'а на экране}


//******************************************************************************
 procedure PageMargins(l,r,t,b : Single; var wrd : Variant);
 {* Устанавливает поля страницы}

 procedure PageOrientation(orientation : Integer; var wrd : Variant);
  {* Устанавливает альбомную/книжную ориентацию страницы.
  wdOrientPortrait = 0;     //книжна
  wdOrientLandscape = 1;    //альбомна      }

 procedure HFDistance(h,f : Single; var wrd : Variant);
  {* Устанавливает отступы для верхнего и нижнего колонтитулов}

 Procedure PageSize (w,h : Single; var wrd : Variant);
  {* Устанавливает размер страницы в сантиметрах. Для А4 w=29.7, h=21}

 procedure SetOnPage(style : Integer; var Wrd : Variant);
 {* Устанавливает количество и размещение листов на странице.
  wdNormalPage = 0;   //1 к 1 - стандартно
  wdTwoOnOne = 1;     //2 стор. на 1
  wdBookFold = 2;     //брошюра }

 procedure PageAlign(align : Integer; var wrd : Variant);
  {* Устанавливает вертикальное выравнивание текста на странице.
  wdAlignVerticalTop = 0;       //по верхнему краю
  wdAlignVerticalCenter = 1;    //по центру
  wdAlignVerticalJustify = 2;   //разтянуть по висоте
  wdAlignVerticalBottom = 3;    //по нижнему краю    }

 procedure NewPage(wrd : Variant);
{* Вставить разрыв страницы в документ}

//******************************************************************************
 procedure FontName(name : ShortString; var wrd : Variant);
  {* Устанавливает название шрифта (напр. 'Courier New')}

 procedure FontSize(sz : Integer; var Wrd : Variant);
 {* Устанавливает размер для текущего шрифта}

 procedure FontBold(Bold : Boolean; var Wrd : Variant);
 {* Устанавливает жирность для текущего шрифта}

 procedure FontItalic(Italic : Boolean; var Wrd : Variant);
 {* Устанавливает курсив для текущего шрифта}

 procedure FontUnderlined(underlined : Boolean; var wrd : Variant);
 {* Устанавливает подчеркивание для текущего шрифта}

 procedure FontShadowed(Shadowed : Boolean; var Wrd : Variant);
 {* Устанавливает тень для текущего шрифта}

 procedure FontColor(Color : Integer; var Wrd : Variant);
 {* Устанавливает цвет для текущего  шрифта
  wdAuto = 0;
  wdBlack = 1;
  wdBlue = 2;
  wdTurquoise = 3;
  wdBrightGreen = 4;
  wdPink = 5;
  wdRed = 6;
  wdYellow = 7;
  wdWhite = 8;
  wdDarkBlue = 9;
  wdTeal = 10;
  wdGreen = 11;
  wdViolet = 12;
  wdDarkRed = 13;
  wdDarkYellow = 14;
  wdGray50 = 15;
  wdGray25 = 16;
  wdByAuthor = -1;
  wdNoHighlight = 0; }

 procedure FontSuperScript(super : Boolean; var Wrd : Variant);
 {* Устанавливает верхние индексы}

 procedure FontSubScript(sub : Boolean; var Wrd : Variant);
 {* Устанавливает нижние индексы}

 procedure FontSpacing(spacing : Single; var Wrd : Variant);
 {* Устанавливает межсимвольный интервал: 1, 1.5 и т.д.}

 procedure FontScaling(scaling : Integer; var Wrd : Variant);
 {* Установливает масштаб для текущего шрифта в %}

 procedure FontPosition(position : Single; var Wrd : Variant);
 {* Устанавливает смещение текста вверх (положительные значения)
   или вниз (отрицалельные значения) в пт}

 procedure AddText(s : ShortString; var wrd : Variant);
{* Вставить строку}

 procedure AddParagraph(var wrd : Variant);
 {* Начать новый абзац}

//******************************************************************************
 procedure ParagraphAlign(align : Integer; var Wrd : Variant);
{* Установить выравнивание абзаца по ширине
  wdAlignParagraphLeft = 0;
  wdAlignParagraphCenter = 1;
  wdAlignParagraphRight = 2;
  wdAlignParagraphJustify = 3;}

 procedure ParagraphLineSpace(space : Integer; var wrd : Variant);
 {* Установить междустрочный интервал:
  wdLineSpaceSingle = 0;
  wdLineSpace1pt5 = 1;
  wdLineSpaceDouble = 2;
  wdLineSpaceAtLeast = 3;
  wdLineSpaceExactly = 4;
  wdLineSpaceMultiple = 5;}

 procedure ParagraphIndents(l,r : Single; var wrd : Variant);
 {* Установить отступы абзаца слева и справа (в см)}

 procedure ParagraphSpaces(top,bottom : Single; var wrd : Variant);
 {* Установить отступы абзаца сверху и снизу}

 Procedure ParagraphFirstLine(indent : Single; var wrd : Variant);
 {* Установить отступ первой строки абзаца (в см)}

 //*****************************************************************************
 procedure AddTabPosition(pos : Single; var wrd : Variant);
{* Вставить позицию табуляции в pos см}

 procedure DefaultTabPos(pos : Single; var wrd : Variant);
 {* Установить позицию табуляции по умолчанию в pos см}

 procedure ClearAllTabs(var wrd : Variant);
 {* Очистить все позиции табуляции}

 //*****************************************************************************
 procedure CreateTable(Col,Row : Integer; var wrd : Variant);
 {* Создать таблицу с атрибутами по умолчанию}

 Procedure SetColWidth(wid : Single; var wrd : Variant);
 {* Установить ширину текущего столбца в таблице.
  Использовать после вызова AddText(), иначе значение игнорируется}

 Procedure SetRowHeight(h : Single; var wrd : Variant);
 {* Установить высоту текущей строки в таблице.}

 procedure GotoRight(cells : Integer; var wrd : Variant);
 {* Перейти на cells ячеек вправо. Если таблица закончилась, вставляет новую строку}

 procedure GotoLeft(cells : Integer; var wrd : Variant);
 {* Перейти на cells ячеек влево.}

 procedure GotoUp(lines : Integer; var wrd : Variant);
 {* Перейти на lines строк вверх}

 procedure GotoDown(lines : Integer; var wrd : Variant);
 {* Перейти на lines строк вниз.}

 Procedure MergeCellsR(count : Integer; var Wrd : Variant);
 {* Обьединить указанные ячейки вправо. Курсор должен находиться в первой из них}

 Procedure MergeCellsD(count : Integer; var Wrd : Variant);
 {* Обьединить указанные ячейки вниз. Курсор должен находиться в первой из них}

 procedure DeleteRow(var Wrd : Variant);
 {* Удалить текущую строку таблицы}

 procedure DeleteCol(var Wrd : Variant);
 {* Удалить текущий столбец таблицы}

 Procedure ExitTable(wrd : Variant);
{* Выйти из таблицы}

 procedure CellTextOrientation(orient : Integer;wrd : Variant);
 {* Направление текста в таблице:
  wdTextOrientationHorizontal = 0;
  wdTextOrientationUpward = 2;
  wdTextOrientationDownward = 3;
  wdTextOrientationVerticalFarEast = 1;
  wdTextOrientationHorizontalRotatedFarEast = 4;}

 procedure InsertHeader(hdr : ShortString; var wrd : Variant);
{* Вставить верхний колонтитул}

 procedure InsertFooter(ftr : ShortString; var wrd : Variant);
{* Вставить нижний колонтитул}

procedure SetWordVisible(var wrd : Variant; Visible : Boolean);
{*сделать Ворд видимым на экране}

function CheckWordVersion(var wrd : Variant):Boolean;
{*проверить может ли использоваться установленная версия Ворда}

Function SaveDocAs(file_:Shortstring; var wrd : Variant):boolean;
{*сохранить созданный документ с указанным именем и путем}

Function CloseDoc(var wrd : Variant):boolean;
{*закрыть документ}

Function CloseWord(var wrd : Variant):boolean;
{*выход из Word'a}

Function PrintDialogWord(var wrd : Variant):boolean;
{*вызов диалога печати}

Procedure CreateTableEx(Col,Row : Integer; DefaultTableBehavior,AutoFitBehavior : Integer; var wrd : Variant);
{*создать таблицу с расширенными параметрами.
  arg1:
  wdWord8TableBehavior = 0;
  wdWord9TableBehavior = 1;
  arg2:
  wdAutoFitFixed = 0;
  wdAutoFitContent = 1;
  wdAutoFitWindow = 2;
}


{******************************************************************************}
{                                                                              }
{                   Функции для автоматизации M$ Excel                         }
{                                                                              }
{******************************************************************************}

 procedure NewXlsDocument(var xls : Variant; visible : Boolean);
 {*Создает новый документ Excel}

 procedure OpenXlsDocument(var xls : Variant; xlsfile : ShortString);
 {*Открыть указанный документ}

 function GetXlsWorkBook(var xls : Variant; idx : Integer): Variant;
 {*Получить ссылку на активную книгу}

 function GetXlsWorkBookSheet(var WorkBook : Variant; idx : Integer):Variant;
 {*Получить ссылку на активный лист}

 procedure SetCellValue(var Xls : Variant; CellName : ShortString; value : ShortString);
 {*Записать значение как текст в указанную ячейку}

 procedure SetCellValueInteger(var Xls : Variant; CellName : ShortString; value : Integer);
 {*Записать значение как целое число в указанную ячейку}

 procedure SetCellValueFloat(var Xls : Variant; CellName : ShortString; value : Double);
 {*Записать значение как число с плавающей точкой в указанную ячейку}

 procedure SetCellValueDate(var Xls : Variant; CellName : ShortString; value : TDatetime);
 {*Записать значение как дату в указанную ячейку}

 procedure SetCellValueCurrency(var Xls : Variant; CellName : ShortString; value : Currency);
 {*Записать значение с денежным форматом в указанную ячейку}

 procedure SetCellValueFormat(var Xls : Variant; CellName : ShortString; valueformat : ShortString);
 {*Указать формат данных заданной ячейки}

 {******************************************************************************}
 {                                                                              }
 {                   Функции для автоматизации M$ Outlook                       }
 {              (Взято из библиотеки outlookdll от EmeraldMan)                  }
 {******************************************************************************}

 procedure OutLookConnect(var OL: Variant);
 {*Подключаемся к OutLook}


 procedure OutLookNewFolder(var OL: Variant; s: ShortString);
 {*Новая папка контактов}

 procedure OutLookNewContact(var OL: Variant; folder:ShortString; name:ShortString);
 {*Новый контакт
 folder - название папки в OutLook контактах;
 name - имя контакта
 Если папки folder не существует, то будет создана.
 Это вполне рабочая функция, правда добавляет только имя;
 Для добавления всего остального проишите
 OutlookContact.LastName
 OutlookContact.MiddleName
 OutlookContact.CompanyName
 Contact.HomeTelephoneNumber
 Contact.Email1Address
 и т.д. кому что нужно}

 procedure OutLookDisConnect(var OL: Variant);
 {*Отключаемся от OutLook}



implementation
function CentimetersToPoints(cm : Real) : Real;StdCall;
 {*Переводит сантиметры в точки}
 begin
  result := cm * cm2p;
 end;


{******************************************************************************}
{                                                                              }
{                   Функции для автоматизации M$ Word                          }
{                                                                              }
{******************************************************************************}
 procedure NewDocument(var Wrd : Variant; visible : Boolean);
 {*Создает новый документ}
 begin
  Wrd := CreateOleObject('Word.Application');
  Wrd.Visible := Visible;
  Wrd.Documents.Add;
  Wrd.Application.WindowState := wdWindowStateMaximize;
 end;

//***********************************************************************************
 procedure PageMargins(l,r,t,b : Single; var wrd : Variant);
 {*Устанавливает поля страницы}
 begin
  Wrd.ActiveDocument.PageSetup.LeftMargin := CentimetersToPoints(l);
  Wrd.ActiveDocument.PageSetup.RightMargin := CentimetersToPoints(r);
  Wrd.ActiveDocument.PageSetup.TopMargin := CentimetersToPoints(t);
  Wrd.ActiveDocument.PageSetup.BottomMargin := CentimetersToPoints(b);
 end;

 procedure PageOrientation(orientation : Integer; var wrd : Variant);
 {*Устанавливает альбомную/книжную ориентацию страницы }
 begin
  Wrd.ActiveDocument.PageSetup.Orientation := orientation;
 end;

 procedure HFDistance(h,f : Single; var wrd : Variant);
 {*Устанавливает отступы для верхнего и нижнего колонтитулов}
 begin
  Wrd.ActiveDocument.PageSetup.HeaderDistance := CentimetersToPoints(h);
  Wrd.ActiveDocument.PageSetup.FooterDistance := CentimetersToPoints(f);
 end;

 Procedure PageSize (w,h : Single; var wrd : Variant);
 {*Устанавливает размер страницы в сантиметрах. Для А4 w=29.7, h=21}
 begin
  Wrd.ActiveDocument.PageSetup.PageWidth := CentimetersToPoints(w);
  Wrd.ActiveDocument.PageSetup.PageHeight := CentimetersToPoints(h);
 end;

 procedure SetOnPage(style : Integer; var Wrd : Variant);
 {*Устанавливает количество и размещение листов на странице.
  wdNormalPage = 0;   //1 к 1 - стандартно
  wdTwoOnOne = 1;     //2 стор. на 1
  wdBookFold = 2;     //брошюра }
 begin
  Wrd.ActiveDocument.PageSetup.MirrorMargins := false;
  Wrd.ActiveDocument.PageSetup.TwoPagesOnOne := false;
  Wrd.ActiveDocument.PageSetup.BookFoldPrinting := false;
  case style of
   wdTwoOnOne : Wrd.ActiveDocument.PageSetup.TwoPagesOnOne := true;
   wdBookFold : Wrd.ActiveDocument.PageSetup.BookFoldPrintiong := true;
  end;
 end;

 procedure PageAlign(align : Integer; var wrd : Variant);
 {*Устанавливает вертикальное выравнивание текста на странице.
  wdAlignVerticalTop = 0;       //по верхнему краю
  wdAlignVerticalCenter = 1;    //по центру
  wdAlignVerticalJustify = 2;   //разтянуть по висоте
  wdAlignVerticalBottom = 3;    //по нижнему краю    }
 begin
  Wrd.ActiveDocument.PageSetup.VerticalAlignment := align;
 end;

 procedure NewPage(wrd : Variant);
{*Вставить разрыв страницы в документ}
 begin
  wrd.Selection.InsertBreak(wdPageBreak);
 end;

//*******************************************************************************
 procedure FontName(name : ShortString; var wrd : Variant);
 {*Устанавливает название шрифта}
 begin
  Wrd.Selection.Font.Name := name;
 end;

 procedure FontSize(sz : Integer; var Wrd : Variant);
 {*Устанавливает размер для текущего шрифта}
 begin
  Wrd.Selection.Font.Size := sz;
 end;

 procedure FontBold(Bold : Boolean; var Wrd : Variant);
 {*Устанавливает жирность для текущего шрифта}
 begin
  Wrd.Selection.Font.Bold := Bold;
 end;

 procedure FontItalic(Italic : Boolean; var Wrd : Variant);
 {*Устанавливает курсив для текущего шрифта}
 begin
  Wrd.Selection.Font.Italic := Italic;
 end;

 procedure FontUnderlined(underlined : Boolean; var wrd : Variant);
 {*Устанавливает подчеркивание для текущего шрифта}
 begin
  if underlined then wrd.Selection.Font.Underline := wdUnderlineSingle
                else wrd.Selection.Font.Underline := wdUnderlineNone;
 end;

 procedure FontShadowed(Shadowed : Boolean; var Wrd : Variant);
 {*Устанавливает тень для текущего шрифта}
 begin
  Wrd.Selection.Font.Shadow := Shadowed;
 end;

 procedure FontColor(Color : Integer; var Wrd : Variant);
 {*Устанавливает цвет для текущего  шрифта
  wdAuto = 0;
  wdBlack = 1;
  wdBlue = 2;
  wdTurquoise = 3;
  wdBrightGreen = 4;
  wdPink = 5;
  wdRed = 6;
  wdYellow = 7;
  wdWhite = 8;
  wdDarkBlue = 9;
  wdTeal = 10;
  wdGreen = 11;
  wdViolet = 12;
  wdDarkRed = 13;
  wdDarkYellow = 14;
  wdGray50 = 15;
  wdGray25 = 16;
  wdByAuthor = -1;
  wdNoHighlight = 0; }
 begin
  Wrd.Selection.Font.Color := Color;
 end;

 procedure FontSuperScript(super : Boolean; var Wrd : Variant);
 {*Устанавливает верхние индексы}
 begin
  Wrd.Selection.Font.Superscript := super;
 end;

 procedure FontSubScript(sub : Boolean; var Wrd : Variant);
 {*Устанавливает нижние индексы}
 begin
  Wrd.Selection.Font.Subscript := sub;
 end;

 procedure FontSpacing(spacing : Single; var Wrd : Variant);
 {*Устанавливает межсимвольный интервал: 1, 1.5 и т.д.}
 begin
  Wrd.Selection.Font.Spacing := spacing;
 end;

 procedure FontScaling(scaling : Integer; var Wrd : Variant);
 {*Установливает масштаб для текущего шрифта в %}
 begin
  Wrd.Selection.Font.Scaling := scaling;
 end;

 procedure FontPosition(position : Single; var Wrd : Variant);
 {*Устанавливает смещение текста вверх (положительные значения)
   или вниз (отрицалельные значения) в пт}
 begin
  Wrd.Selection.Font.Position := position;
 end;

 procedure AddText(s : ShortString; var wrd : Variant);
 {*Вставить строку}
 begin
  Wrd.Selection.TypeText(s);
 end;

 procedure AddParagraph(var wrd : Variant);
 {*Начать новый абзац}
 begin
  Wrd.Selection.TypeParagraph;
 end;

//***************************************************************************************
 procedure ParagraphAlign(align : Integer; var wrd : Variant);
 {*Установить выравнивание абзаца по ширине
  wdAlignParagraphLeft = 0;
  wdAlignParagraphCenter = 1;
  wdAlignParagraphRight = 2;
  wdAlignParagraphJustify = 3;}
 begin
  Wrd.Selection.ParagraphFormat.Alignment := align;
 end;

 procedure ParagraphLineSpace(space : Integer; var wrd : Variant);
 {*Установить междустрочный интервал:
  wdLineSpaceSingle = 0;
  wdLineSpace1pt5 = 1;
  wdLineSpaceDouble = 2;
  wdLineSpaceAtLeast = 3;
  wdLineSpaceExactly = 4;
  wdLineSpaceMultiple = 5;         }
 begin
  Wrd.Selection.ParagraphFormat.LineSpacingRule := space;
 end;

 procedure ParagraphIndents(l,r : Single; var wrd : Variant);
 {*Установить отступы абзаца слева и справа (в см)}
 begin
  Wrd.Selection.ParagraphFormat.LeftIndent := CentimetersToPoints(l);
  Wrd.Selection.ParagraphFormat.RightIndent := CentimetersToPoints(r);
 end;

 procedure ParagraphSpaces(top,bottom : Single; var wrd : Variant);
 {*Установить отступы абзаца сверху и снизу}
 begin
  Wrd.Selection.ParagraphFormat.SpaceBefore := top;
  Wrd.Selection.ParagraphFormat.SpaceBeforeAuto := false;
  Wrd.Selection.ParagraphFormat.SpaceAfter := bottom;
  Wrd.Selection.ParagraphFormat.SpaceAfterAuto := false;
 end;

 Procedure ParagraphFirstLine(indent : Single; var wrd : Variant);
 {*Установить отступ первой строки абзаца (в см)}
 begin
   wrd.Selection.ParagraphFormat.FirstLineIndent := CentimetersToPoints(indent);
 end;

//*********************************************************************************
 procedure AddTabPosition(pos : Single; var wrd : Variant);
 {*Вставить позицию табуляции в pos см}
 begin
  Wrd.Selection.ParagraphFormat.TabStops.Add(CentimetersToPoints(pos),wdAlignTabLeft,wdTabLeaderSpaces);
 end;

 procedure DefaultTabPos(pos : Single; var wrd : Variant);
 {*Установить позицию табуляции по умолчанию в pos см}
 begin
  Wrd.Selection.ParagraphFormat.DefaultTabStop := CentimetersToPoints(pos);
 end;

 procedure ClearAllTabs(var wrd : Variant);
 {*Очистить все позиции табуляции}
 begin
  Wrd.Selection.ParagraphFormat.TabStops.ClearAll;
 end;

//***********************************************************************************
 procedure CreateTable(Col,Row : Integer; var wrd : Variant);
 {*Создать таблицу с атрибутами по умолчанию}
 begin
   Wrd.ActiveDocument.Tables.Add(Wrd.Selection.Range,row,col,wdWord9TableBehavior,wdAutoFitWindow);
 end;

 Procedure SetColWidth(wid : Single; var wrd : Variant);
 {*Установить ширину текущего столбца в таблице.
  Использовать после вызова AddText(), иначе значение игнорируется}
 begin
  Wrd.Selection.Columns.SetWidth(wid, wdAdjustProportional);
 end;

 procedure SetRowHeight(h: Single; var wrd: Variant);
 begin
   Wrd.Selection.Rows.SetHeight(h,wdRowHeightAuto);
 end;

 procedure GotoRight(cells : Integer; var wrd : Variant);
 {*Перейти на cells ячеек вправо. Если таблица закончилась, вставляет новую строку}
 var i : Integer;
 begin
  for i := 1 to cells do Wrd.Selection.MoveRight(wdCell);
 end;

 procedure GotoLeft(cells : Integer; var wrd : Variant);
 {*Перейти на cells ячеек влево.}
 var i : Integer;
 begin
  for i := 1 to cells do Wrd.Selection.MoveLeft(wdCell);
 end;

 procedure GotoUp(lines : Integer; var wrd : Variant);
 {*Перейти на lines строк вверх}
 begin
  Wrd.Selection.MoveUp(wdLine,lines);
 end;

 procedure GotoDown(lines : Integer; var wrd : Variant);
 {*Перейти на lines строк вниз.}
 begin
  Wrd.Selection.MoveDown(wdLine,lines);
 end;

 Procedure MergeCellsR(count : Integer; var Wrd : Variant);
 {*Обьединить указанные ячейки вправо. Курсор должен находиться в первой из них}
 begin
  Wrd.Selection.MoveRight(wdCharacter,count,wdExtend);
  Wrd.Selection.Cells.Merge;
 end;

 Procedure MergeCellsD(count : Integer; var Wrd : Variant);
 {*Обьединить указанные ячейки вниз. Курсор должен находиться в первой из них}
 begin
  Wrd.Selection.MoveDown(wdLine,count,wdExtend);
  Wrd.Selection.Cells.Merge;
 end;

 procedure DeleteRow(var Wrd : Variant);
 {*Удалить текущую строку таблицы}
 begin
  Wrd.Selection.Rows.Delete;
 end;

 procedure DeleteCol(var Wrd : Variant);
 {*Удалить текущий столбец таблицы}
 begin
  Wrd.Selection.Columns.Delete;
 end;

 Procedure ExitTable(wrd : Variant);
 {*Выйти из таблицы}
 begin
  Wrd.Selection.MoveDown(wdLine,1);
 end;

 procedure CellTextOrientation(orient : Integer;wrd : Variant);
 {*Направление текста в таблице
  wdTextOrientationHorizontal = 0;
  wdTextOrientationUpward = 2;
  wdTextOrientationDownward = 3;
  wdTextOrientationVerticalFarEast = 1;
  wdTextOrientationHorizontalRotatedFarEast = 4;}
 begin
  wrd.Selection.Orientation := orient;
 end;

 procedure InsertHeader(hdr : ShortString; var wrd : Variant);
{*Вставить верхний колонтитул}
 begin
  wrd.ActiveWindow.ActivePane.View.SeekView := wdSeekCurrentPageHeader;
  wrd.Selection.ParagraphFormat.Alignment := wdAlignParagraphRight;
  FontColor(wdDarkYellow,wrd);
  FontSize(6,wrd);
  wrd.Selection.TypeText(hdr);
  wrd.ActiveWindow.ActivePane.View.SeekView := wdSeekMainDocument;
 end;

procedure InsertFooter(ftr : ShortString; var wrd : Variant);
{*Вставить нижний колонтитул}
 begin
  wrd.ActiveWindow.ActivePane.View.SeekView := wdSeekCurrentPageFooter;
  wrd.Selection.ParagraphFormat.Alignment := wdAlignParagraphRight;
  FontColor(wdDarkYellow,wrd);
  FontSize(6,wrd);
  wrd.Selection.TypeText(ftr);
  wrd.ActiveWindow.ActivePane.View.SeekView := wdSeekMainDocument;
 end;

procedure SetWordVisible(var wrd : Variant; Visible : Boolean);
{*сделать Ворд видимым на экране}
begin
 wrd.Visible := Visible;
end;

function CheckWordVersion(var wrd : Variant):Boolean;
{*проверить может ли использоваться установленная версия Ворда}
begin
 result := false;
 if wrd.Version>9 then result := true;
end;

Function SaveDocAs(file_:Shortstring; var wrd : Variant):boolean;
{*сохранить созданный документ с указанным именем и путем}
begin
 SaveDocAs:=true;
 try
 Wrd.ActiveDocument.SaveAs(file_);
 except
 SaveDocAs:=false;
 end;
End;

Function CloseDoc(var wrd : Variant):boolean;
{*закрыть документ}
begin
 CloseDoc:=true;
 try
  Wrd.ActiveDocument.Close;
 except
  CloseDoc:=false;
 end;
End;

Function CloseWord(var wrd : Variant):boolean;
{*выход из Word'a}
begin
 CloseWord:=true;
 try
  Wrd.Quit;
 except
  CloseWord:=false;
 end;
End;

Function PrintDialogWord(var wrd : Variant):boolean;
{*вызов диалога печати}
begin
 PrintDialogWord:=true;
 try
  Wrd.Dialogs.Item(wdDialogFilePrint).Show;
 except
  PrintDialogWord:=false;
 end;
End;

Procedure CreateTableEx(Col,Row : Integer; DefaultTableBehavior,AutoFitBehavior : Integer; var wrd : Variant);
{*создать таблицу с расширенными параметрами.
  arg1:
  wdWord8TableBehavior = 0;
  wdWord9TableBehavior = 1;
  arg2:
  wdAutoFitFixed = 0;
  wdAutoFitContent = 1;
  wdAutoFitWindow = 2;
}
begin
 Wrd.ActiveDocument.Tables.Add(Wrd.Selection.Range,row,col,DefaultTableBehavior,AutoFitBehavior);
end;

{******************************************************************************}
{                                                                              }
{                   Функции для автоматизации M$ Excel                         }
{                                                                              }
{******************************************************************************}

 procedure NewXlsDocument(var xls : Variant; visible : Boolean);
 {*Создает новый документ Excel}
 begin
  xls := CreateOleObject('Excel.Application');
  xls.Visible := Visible;
  xls.WorkBooks.Add;
 end;

 procedure OpenXlsDocument(var xls : Variant; xlsfile : ShortString);
 {*Открыть указанный документ}
 begin
  xls := CreateOleObject('Excel.Application');
  if xlsfile<>'' then xls.WorkBooks.Open(xlsfile);
  xls.Visible := true;
 end;

 function GetXlsWorkBook(var xls : Variant; idx : Integer): Variant;
 {*Получить ссылку на активную книгу}
 var workbooks : Variant;
 begin
   result := Null;
   WorkBooks := xls.WorkBooks;
   if idx<workbooks.Count then Result := WorkBooks.item[idx];
 end;

 function GetXlsWorkBookSheet(var WorkBook : Variant; idx : Integer):Variant;
 {*Получить ссылку на активную страницу}
 begin
  result := Null;
  if idx<WorkBook.Sheets.Count then Result := WorkBook.Sheets.Item[idx];
 end;

 procedure SetCellValue(var Xls : Variant; CellName : ShortString; value : ShortString);
 {*Записать значение как текст в указанную ячейку}
 begin
  Xls.ActiveSheet.Range[CellName].Value := value;
 end;

 procedure SetCellValueInteger(var Xls : Variant; CellName : ShortString; value : Integer);
 {*Записать значение как целое число в указанную ячейку}
 begin
 //Xls.ActiveSheet.Range[CellName].NumberFormat:='0';
  Xls.ActiveSheet.Range[CellName].Value := value;
 end;

 procedure SetCellValueFloat(var Xls : Variant; CellName : ShortString; value : Double);
 {*Записать значение как число с плавающей точкой в указанную ячейку}
 var Range : Variant;
 begin
  Range := Xls.ActiveSheet.Range[CellName];
  Range.Value:=value;
//  Range.NumberFormat:='General';
 end;

 procedure SetCellValueCurrency(var Xls : Variant; CellName : ShortString; value : Currency);
 {*Записать значение как денежное в указанную ячейку}
 begin
  Xls.ActiveSheet.Range[CellName].Value := value;
//  Xls.ActiveSheet.Range[CellName].NumberFormat:='0.00';
 end;

 procedure SetCellValueFormat(var Xls : Variant; CellName : ShortString; valueformat : ShortString);
 {*Указать формат данных заданной ячейки}
 begin
  Xls.ActiveSheet.Range[CellName].NumberFormat:=valueformat;
 end;

 procedure SetCellValueDate(var Xls : Variant; CellName : ShortString; value : TDatetime);
 {*Записать значение как дату в указанную ячейку}
 begin
//  Xls.ActiveSheet.Range[CellName].NumberFormat:='dd.mm.yyyy';
  Xls.ActiveSheet.Range[CellName].Value := Value;
 end;

 {******************************************************************************}
 {                                                                              }
 {                   Функции для автоматизации M$ Outlook                       }
 {              (Взято из библиотеки outlookdll от EmeraldMan)                  }
 {******************************************************************************}

procedure OutLookConnect(var OL: Variant);
{*Подключаемся к OutLook}
begin
  OL := CreateOleObject('Outlook.Application');
end;

procedure OutLookNewFolder(var OL: Variant; s: ShortString);
{*Новая папка контактов}
begin
  OL.GetNameSpace('MAPI').GetDefaultFolder(olFolderContacts).AddFolder(s);
end;

procedure OutLookNewContact(var OL: Variant; folder:ShortString; name:ShortString);
{*Новый контакт
folder - название папки в OutLook контактах;
name - имя контакта
Если папки folder не существует, то будет создана.
Это вполне рабочая функция, правда добавляет только имя;
Для добавления всего остального проишите
OutlookContact.LastName
OutlookContact.MiddleName
OutlookContact.CompanyName
Contact.HomeTelephoneNumber
Contact.Email1Address
и т.д. кому что нужно}
var
  NameSpace : OleVariant;
  ContactsRoot : OleVariant;
  ContactsFolder : OleVariant;
  OutlookContact : OleVariant;
  SubFolderName : string;
  Position : integer;
  Found : boolean;
  Counter : integer;
  TestContactFolder : OleVariant;
begin
  // Get name space
  NameSpace := OL.GetNameSpace('MAPI');
   // Get root contacts folder
  ContactsRoot := NameSpace.GetDefaultFolder(olFolderContacts);
   // Iterate to subfolder
  ContactsFolder := ContactsRoot;
  while folder <> '' do begin
    // Extract next subfolder
    Position := Pos('\', folder);
      if Position > 0 then begin
        SubFolderName := Copy(folder, 1, Position - 1);
        folder := Copy(folder, Position + 1, Length(folder));
      end
      else begin
        SubFolderName := folder;
        folder := '';
      end;
      if SubFolderName = '' then Break;
      // Search subfolder
      Found := False;
      for Counter := 1 to ContactsFolder.Folders.Count do begin
        TestContactFolder := ContactsRoot.Folders.Item(Counter);
        if LowerCase(TestContactFolder.Name) = LowerCase(SubFolderName) then begin
          ContactsFolder := TestContactFolder;
          Found := True;
          Break;
        end;
      end;
     // If not found create
     if not Found then ContactsFolder := ContactsFolder.Folders.Add(SubFolderName);
  end;
  // Create contact item
  OutlookContact := ContactsFolder.Items.Add;
  // Fill contact information
  OutlookContact.FirstName := Name;
  OutlookContact.Save;
end;

procedure OutLookDisConnect(var OL: Variant);
{*Отключаемся от OutLook}
begin
  OL := Unassigned;
end;


end.
