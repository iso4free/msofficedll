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

program testdll;

{$mode objfpc}{$H+}

uses
  Classes, SysUtils, variants, comobj, umsoffice;
 var w : Variant;
   _xl : Variant;
begin
 NewDocument(w,true);
 PageOrientation(wdOrientLandscape,w);
 PageMargins(2,2,3,2,w);
 HFDistance(1,0,w);
 FontName('Arial',w);
 FontSize(16,w);
 AddText(Utf8ToAnsi('Тест формирования докуменнта в MSWord из Lazarus'),w);
 CreateTable(3,3,w);
 MergeCellsD(2,w);
 ParagraphAlign(wdAlignParagraphCenter,w);
 FontBold(true,w);
 AddText(Utf8ToAnsi('Обьединение ячеек вниз'),w);
 GotoRight(1,w);
 MergeCellsR(2,w);
 ParagraphAlign(wdAlignParagraphCenter,w);
 FontBold(true,w);
 AddText(Utf8ToAnsi('Обьединение ячеек вправо'),w);
 GotoRight(1,w);
 SetWordVisible(w,true);
 SaveDocAs('testword.doc',w);
 CloseDoc(w);
 CloseWord(w);
 NewXlsDocument(_xl,true);
 SetCellValue(_xl,'A1',Utf8ToAnsi('Текстовое значение'));
 SetCellValue(_xl,'A2',Utf8ToAnsi('пример вывода текста в ячейку'));
 SetCellValue(_xl,'B1',Utf8ToAnsi('Целочисельное значение'));
 SetCellValueInteger(_xl,'B2',12345);
 SetCellValue(_xl,'C1',Utf8ToAnsi('Число с плавающей запятой'));
 SetCellValueFloat(_xl,'C2',123.456);
 SetCellValue(_xl,'D1',Utf8ToAnsi('Дата'));
 SetCellValueDate(_xl,'D2',Now);
 SetCellValue(_xl,'E1',Utf8ToAnsi('Деньги'));
 SetCellValueFloat(_xl,'E2',123.45);
 SetCellValueFormat(_xl,'E2',Utf8ToAnsi('0.00 грн.'));
 end.
