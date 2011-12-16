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

unit uofficedll;
{* Динамическая библиотека для автоматизации создания отчётов в M$ Word XP/2003 в Delphi/Lazarus.
Для использования необходимо в проект подключить модуль uofficedll.pas
(в Lazarus также необходимо подключить модуль Variants и модуль ComObj).}
interface
  uses sysutils;
 {$I const.inc}

 function CentimetersToPoints(cm : Real) : Real; stdcall; external 'msofficedll.dll' name 'CentimetersToPoints';
 {* Переводит сантиметры в точки}

 procedure NewDocument(var Wrd : Variant; visible : Boolean); stdcall; external 'msofficedll.dll' name 'NewDocument';
 {* Создает новый документ.  Параметр visible указывает, будет ли отображаться окно Word'а на экране}

//******************************************************************************
 procedure PageMargins(l,r,t,b : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'PageMargins';
 {* Устанавливает поля страницы}

 procedure PageOrientation(orientation : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'PageOrientation';
  {* Устанавливает альбомную/книжную ориентацию страницы.
  wdOrientPortrait = 0;     //книжна
  wdOrientLandscape = 1;    //альбомна      }

 procedure HFDistance(h,f : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'HFDistance';
  {* Устанавливает отступы для верхнего и нижнего колонтитулов}

 Procedure PageSize (w,h : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'PageSize';
  {* Устанавливает размер страницы в сантиметрах. Для А4 w=29.7, h=21}

 procedure SetOnPage(style : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'SetOnPage';
 {* Устанавливает количество и размещение листов на странице.
  wdNormalPage = 0;   //1 к 1 - стандартно
  wdTwoOnOne = 1;     //2 стор. на 1
  wdBookFold = 2;     //брошюра }

 procedure PageAlign(align : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'PageAlign';
  {* Устанавливает вертикальное выравнивание текста на странице.
  wdAlignVerticalTop = 0;       //по верхнему краю
  wdAlignVerticalCenter = 1;    //по центру
  wdAlignVerticalJustify = 2;   //разтянуть по висоте
  wdAlignVerticalBottom = 3;    //по нижнему краю    }

 procedure NewPage(wrd : Variant); stdcall; external 'msofficedll.dll' name 'NewPage';
{* Вставить разрыв страницы в документ}

//******************************************************************************
 procedure FontName(name : ShortString; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontName';
  {* Устанавливает название шрифта (напр. 'Courier New')}

 procedure FontSize(sz : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontSize';
 {* Устанавливает размер для текущего шрифта}

 procedure FontBold(Bold : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontBold';
 {* Устанавливает жирность для текущего шрифта}

 procedure FontItalic(Italic : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontItalic';
 {* Устанавливает курсив для текущего шрифта}

 procedure FontUnderlined(underlined : Boolean; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontUnderlined';
 {* Устанавливает подчеркивание для текущего шрифта}

 procedure FontShadowed(Shadowed : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontShadowed';
 {* Устанавливает тень для текущего шрифта}

 procedure FontColor(Color : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontColor';
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

 procedure FontSuperScript(super : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontSuperScript';
 {* Устанавливает верхние индексы}

 procedure FontSubScript(sub : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontSubScript';
 {* Устанавливает нижние индексы}

 procedure FontSpacing(spacing : Single; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontSpacing';
 {* Устанавливает межсимвольный интервал: 1, 1.5 и т.д.}

 procedure FontScaling(scaling : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontScaling';
 {* Установливает масштаб для текущего шрифта в %}

 procedure FontPosition(position : Single; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontPosition';
 {* Устанавливает смещение текста вверх (положительные значения)
   или вниз (отрицалельные значения) в пт}

 procedure AddText(s : ShortString; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'AddText';
{* Вставить строку}

 procedure AddParagraph(var wrd : Variant); stdcall; external 'msofficedll.dll' name 'AddParagraph';
 {* Начать новый абзац}

//******************************************************************************
 procedure ParagraphAlign(align : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphAlign';
{* Установить выравнивание абзаца по ширине
  wdAlignParagraphLeft = 0;
  wdAlignParagraphCenter = 1;
  wdAlignParagraphRight = 2;
  wdAlignParagraphJustify = 3;}

 procedure ParagraphLineSpace(space : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphLineSpace';
 {* Установить междустрочный интервал:
  wdLineSpaceSingle = 0;
  wdLineSpace1pt5 = 1;
  wdLineSpaceDouble = 2;
  wdLineSpaceAtLeast = 3;
  wdLineSpaceExactly = 4;
  wdLineSpaceMultiple = 5;}

 procedure ParagraphIndents(l,r : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphIndents';
 {* Установить отступы абзаца слева и справа (в см)}

 procedure ParagraphSpaces(top,bottom : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphSpaces';
 {* Установить отступы абзаца сверху и снизу}

 Procedure ParagraphFirstLine(indent : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphFirstLine';
 {* Установить отступ первой строки абзаца (в см)}

 //*****************************************************************************
 procedure AddTabPosition(pos : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'AddTabPosition';
{* Вставить позицию табуляции в pos см}

 procedure DefaultTabPos(pos : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'DefaultTabPos';
 {* Установить позицию табуляции по умолчанию в pos см}

 procedure ClearAllTabs(var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ClearAllTabs';
 {* Очистить все позиции табуляции}

 //*****************************************************************************
 procedure CreateTable(Col,Row : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'CreateTable';
 {* Создать таблицу с атрибутами по умолчанию}

 Procedure SetColWidth(wid : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'SetColWidth';
 {* Установить ширину текущего столбца в таблице.
  Использовать после вызова AddText(), иначе значение игнорируется}

 Procedure SetRowHeight(h : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'SetRowHeight';
 {* Установить высоту текущей строки в таблице.}

 procedure GotoRight(cells : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'GotoRight';
 {* Перейти на cells ячеек вправо. Если таблица закончилась, вставляет новую строку}

 procedure GotoLeft(cells : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'GotoLeft';
 {* Перейти на cells ячеек влево.}

 procedure GotoUp(lines : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'GotoUp';
 {* Перейти на lines строк вверх}

 procedure GotoDown(lines : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'GotoDown';
 {* Перейти на lines строк вниз.}

 Procedure MergeCellsR(count : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'MergeCellsR';
 {* Обьединить указанные ячейки вправо. Курсор должен находиться в первой из них}

 Procedure MergeCellsD(count : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'MergeCellsD';
 {* Обьединить указанные ячейки вниз. Курсор должен находиться в первой из них}

 procedure DeleteRow(var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'DeleteRow';
 {* Удалить текущую строку таблицы}

 procedure DeleteCol(var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'DeleteCol';
 {* Удалить текущий столбец таблицы}

 Procedure ExitTable(wrd : Variant); stdcall; external 'msofficedll.dll' name 'ExitTable';
{* Выйти из таблицы}

 procedure CellTextOrientation(orient : Integer;wrd : Variant); stdcall; external 'msofficedll.dll' name 'CellTextOrientation';
 {* Направление текста в таблице:
  wdTextOrientationHorizontal = 0;
  wdTextOrientationUpward = 2;
  wdTextOrientationDownward = 3;
  wdTextOrientationVerticalFarEast = 1;
  wdTextOrientationHorizontalRotatedFarEast = 4;}

 procedure InsertHeader(hdr : ShortString; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'InsertHeader';
{* Вставить верхний колонтитул}

 procedure InsertFooter(ftr : ShortString; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'InsertFooter';
{* Вставить нижний колонтитул}

procedure SetWordVisible(var wrd : Variant; Visible : Boolean); stdcall; external 'msofficedll.dll' name 'SetWordVisible';
{*сделать Ворд видимым на экране}

function CheckWordVersion(var wrd : Variant):Boolean; stdcall; external 'msofficedll.dll' name 'CheckWordVersion';
{*проверить может ли использоваться установленная версия Ворда}

Function SaveDocAs(file_:Shortstring; var wrd : Variant):boolean; stdcall; external 'msofficedll.dll' name 'SaveDocAs';
{*сохранить созданный документ с указанным именем и путем}

Function CloseDoc(var wrd : Variant):boolean; stdcall; external 'msofficedll.dll' name 'CloseDoc';
{*закрыть документ}

Function CloseWord(var wrd : Variant):boolean; stdcall; external 'msofficedll.dll' name 'CloseWord';
{*выход из Word'a}

Function PrintDialogWord(var wrd : Variant):boolean; stdcall; external 'msofficedll.dll' name 'PrintDialogWord';
{*вызов диалога печати}

Procedure CreateTableEx(Col,Row : Integer; DefaultTableBehavior,AutoFitBehavior : Integer; var wrd : Variant);stdcall; external 'msofficedll.dll' name 'CreateTableEx';
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

 procedure NewXlsDocument(var xls : Variant; visible : Boolean);StdCall; external 'msofficedll.dll' name 'NewXlsDocument';
 {*Создает новый документ Excel}

 procedure OpenXlsDocument(var xls : Variant; xlsfile : ShortString);StdCall; external 'msofficedll.dll' name 'OpenXlsDocument';
 {*Открыть указанный документ}

 function GetXlsWorkBook(var xls : Variant; idx : Integer): Variant;StdCall; external 'msofficedll.dll' name 'GetXlsWorkBook';
 {*Получить ссылку на активную книгу}

 function GetXlsWorkBookSheet(var WorkBook : Variant; idx : Integer):Variant;StdCall; external 'msofficedll.dll' name 'GetXlsWorkBookSheet';
 {*Получить ссылку на активный лист}

 procedure SetCellValue(var Xls : Variant; CellName : ShortString; value : ShortString);StdCall; external 'msofficedll.dll' name 'SetCellValue';
 {*Записать значение как текст в указанную ячейку}

 procedure SetCellValueInteger(var Xls : Variant; CellName : ShortString; value : Integer);StdCall; external 'msofficedll.dll' name 'SetCellValueInteger';
 {*Записать значение как целое число в указанную ячейку}

 procedure SetCellValueFloat(var Xls : Variant; CellName : ShortString; value : Double);StdCall; external 'msofficedll.dll' name 'SetCellValueFloat';
 {*Записать значение как число с плавающей точкой в указанную ячейку}

 procedure SetCellValueDate(var Xls : Variant; CellName : ShortString; value : TDatetime);StdCall; external 'msofficedll.dll' name 'SetCellValueDate';
 {*Записать значение как дату в указанную ячейку}

 procedure SetCellValueCurrency(var Xls : Variant; CellName : ShortString; value : Currency);StdCall; external 'msofficedll.dll' name 'SetCellValueFloat';
 {*Записать значение с денежным форматом в указанную ячейку}

 procedure SetCellValueFormat(var Xls : Variant; CellName : ShortString; valueformat : ShortString);StdCall; external 'msofficedll.dll' name 'SetCellValueFormat';
 {*Указать формат данных заданной ячейки}

 {******************************************************************************}
 {                                                                              }
 {                   Функции для автоматизации M$ Outlook                       }
 {              (Взято из библиотеки outlookdll от EmeraldMan)                  }
 {******************************************************************************}

 procedure OutLookConnect(var OL: Variant);StdCall; external 'msofficedll.dll' name 'OutLookConnect';
 {*Подключаемся к OutLook}


 procedure OutLookNewFolder(var OL: Variant; s: ShortString);StdCall; external 'msofficedll.dll' name 'OutLookNewFolder';
 {*Новая папка контактов}

 procedure OutLookNewContact(var OL: Variant; folder:ShortString; name:ShortString);StdCall; external 'msofficedll.dll' name 'OutLookNewContact';
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

 procedure OutLookDisConnect(var OL: Variant);StdCall; external 'msofficedll.dll' name 'OutLookDisConnect';
 {*Отключаемся от OutLook}



implementation

end.
