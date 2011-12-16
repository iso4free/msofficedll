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
{* ������������ ���������� ��� ������������� �������� ������� � M$ Word XP/2003 � Delphi/Lazarus.
��� ������������� ���������� � ������ ���������� ������ uofficedll.pas
(� Lazarus ����� ���������� ���������� ������ Variants � ������ ComObj).}
interface
  uses sysutils;
 {$I const.inc}

 function CentimetersToPoints(cm : Real) : Real; stdcall; external 'msofficedll.dll' name 'CentimetersToPoints';
 {* ��������� ���������� � �����}

 procedure NewDocument(var Wrd : Variant; visible : Boolean); stdcall; external 'msofficedll.dll' name 'NewDocument';
 {* ������� ����� ��������.  �������� visible ���������, ����� �� ������������ ���� Word'� �� ������}

//******************************************************************************
 procedure PageMargins(l,r,t,b : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'PageMargins';
 {* ������������� ���� ��������}

 procedure PageOrientation(orientation : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'PageOrientation';
  {* ������������� ���������/������� ���������� ��������.
  wdOrientPortrait = 0;     //������
  wdOrientLandscape = 1;    //��������      }

 procedure HFDistance(h,f : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'HFDistance';
  {* ������������� ������� ��� �������� � ������� ������������}

 Procedure PageSize (w,h : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'PageSize';
  {* ������������� ������ �������� � �����������. ��� �4 w=29.7, h=21}

 procedure SetOnPage(style : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'SetOnPage';
 {* ������������� ���������� � ���������� ������ �� ��������.
  wdNormalPage = 0;   //1 � 1 - ����������
  wdTwoOnOne = 1;     //2 ����. �� 1
  wdBookFold = 2;     //������� }

 procedure PageAlign(align : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'PageAlign';
  {* ������������� ������������ ������������ ������ �� ��������.
  wdAlignVerticalTop = 0;       //�� �������� ����
  wdAlignVerticalCenter = 1;    //�� ������
  wdAlignVerticalJustify = 2;   //��������� �� ������
  wdAlignVerticalBottom = 3;    //�� ������� ����    }

 procedure NewPage(wrd : Variant); stdcall; external 'msofficedll.dll' name 'NewPage';
{* �������� ������ �������� � ��������}

//******************************************************************************
 procedure FontName(name : ShortString; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontName';
  {* ������������� �������� ������ (����. 'Courier New')}

 procedure FontSize(sz : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontSize';
 {* ������������� ������ ��� �������� ������}

 procedure FontBold(Bold : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontBold';
 {* ������������� �������� ��� �������� ������}

 procedure FontItalic(Italic : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontItalic';
 {* ������������� ������ ��� �������� ������}

 procedure FontUnderlined(underlined : Boolean; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontUnderlined';
 {* ������������� ������������� ��� �������� ������}

 procedure FontShadowed(Shadowed : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontShadowed';
 {* ������������� ���� ��� �������� ������}

 procedure FontColor(Color : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontColor';
 {* ������������� ���� ��� ��������  ������
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
 {* ������������� ������� �������}

 procedure FontSubScript(sub : Boolean; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontSubScript';
 {* ������������� ������ �������}

 procedure FontSpacing(spacing : Single; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontSpacing';
 {* ������������� ������������� ��������: 1, 1.5 � �.�.}

 procedure FontScaling(scaling : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontScaling';
 {* ������������� ������� ��� �������� ������ � %}

 procedure FontPosition(position : Single; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'FontPosition';
 {* ������������� �������� ������ ����� (������������� ��������)
   ��� ���� (������������� ��������) � ��}

 procedure AddText(s : ShortString; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'AddText';
{* �������� ������}

 procedure AddParagraph(var wrd : Variant); stdcall; external 'msofficedll.dll' name 'AddParagraph';
 {* ������ ����� �����}

//******************************************************************************
 procedure ParagraphAlign(align : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphAlign';
{* ���������� ������������ ������ �� ������
  wdAlignParagraphLeft = 0;
  wdAlignParagraphCenter = 1;
  wdAlignParagraphRight = 2;
  wdAlignParagraphJustify = 3;}

 procedure ParagraphLineSpace(space : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphLineSpace';
 {* ���������� ������������� ��������:
  wdLineSpaceSingle = 0;
  wdLineSpace1pt5 = 1;
  wdLineSpaceDouble = 2;
  wdLineSpaceAtLeast = 3;
  wdLineSpaceExactly = 4;
  wdLineSpaceMultiple = 5;}

 procedure ParagraphIndents(l,r : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphIndents';
 {* ���������� ������� ������ ����� � ������ (� ��)}

 procedure ParagraphSpaces(top,bottom : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphSpaces';
 {* ���������� ������� ������ ������ � �����}

 Procedure ParagraphFirstLine(indent : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ParagraphFirstLine';
 {* ���������� ������ ������ ������ ������ (� ��)}

 //*****************************************************************************
 procedure AddTabPosition(pos : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'AddTabPosition';
{* �������� ������� ��������� � pos ��}

 procedure DefaultTabPos(pos : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'DefaultTabPos';
 {* ���������� ������� ��������� �� ��������� � pos ��}

 procedure ClearAllTabs(var wrd : Variant); stdcall; external 'msofficedll.dll' name 'ClearAllTabs';
 {* �������� ��� ������� ���������}

 //*****************************************************************************
 procedure CreateTable(Col,Row : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'CreateTable';
 {* ������� ������� � ���������� �� ���������}

 Procedure SetColWidth(wid : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'SetColWidth';
 {* ���������� ������ �������� ������� � �������.
  ������������ ����� ������ AddText(), ����� �������� ������������}

 Procedure SetRowHeight(h : Single; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'SetRowHeight';
 {* ���������� ������ ������� ������ � �������.}

 procedure GotoRight(cells : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'GotoRight';
 {* ������� �� cells ����� ������. ���� ������� �����������, ��������� ����� ������}

 procedure GotoLeft(cells : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'GotoLeft';
 {* ������� �� cells ����� �����.}

 procedure GotoUp(lines : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'GotoUp';
 {* ������� �� lines ����� �����}

 procedure GotoDown(lines : Integer; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'GotoDown';
 {* ������� �� lines ����� ����.}

 Procedure MergeCellsR(count : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'MergeCellsR';
 {* ���������� ��������� ������ ������. ������ ������ ���������� � ������ �� ���}

 Procedure MergeCellsD(count : Integer; var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'MergeCellsD';
 {* ���������� ��������� ������ ����. ������ ������ ���������� � ������ �� ���}

 procedure DeleteRow(var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'DeleteRow';
 {* ������� ������� ������ �������}

 procedure DeleteCol(var Wrd : Variant); stdcall; external 'msofficedll.dll' name 'DeleteCol';
 {* ������� ������� ������� �������}

 Procedure ExitTable(wrd : Variant); stdcall; external 'msofficedll.dll' name 'ExitTable';
{* ����� �� �������}

 procedure CellTextOrientation(orient : Integer;wrd : Variant); stdcall; external 'msofficedll.dll' name 'CellTextOrientation';
 {* ����������� ������ � �������:
  wdTextOrientationHorizontal = 0;
  wdTextOrientationUpward = 2;
  wdTextOrientationDownward = 3;
  wdTextOrientationVerticalFarEast = 1;
  wdTextOrientationHorizontalRotatedFarEast = 4;}

 procedure InsertHeader(hdr : ShortString; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'InsertHeader';
{* �������� ������� ����������}

 procedure InsertFooter(ftr : ShortString; var wrd : Variant); stdcall; external 'msofficedll.dll' name 'InsertFooter';
{* �������� ������ ����������}

procedure SetWordVisible(var wrd : Variant; Visible : Boolean); stdcall; external 'msofficedll.dll' name 'SetWordVisible';
{*������� ���� ������� �� ������}

function CheckWordVersion(var wrd : Variant):Boolean; stdcall; external 'msofficedll.dll' name 'CheckWordVersion';
{*��������� ����� �� �������������� ������������� ������ �����}

Function SaveDocAs(file_:Shortstring; var wrd : Variant):boolean; stdcall; external 'msofficedll.dll' name 'SaveDocAs';
{*��������� ��������� �������� � ��������� ������ � �����}

Function CloseDoc(var wrd : Variant):boolean; stdcall; external 'msofficedll.dll' name 'CloseDoc';
{*������� ��������}

Function CloseWord(var wrd : Variant):boolean; stdcall; external 'msofficedll.dll' name 'CloseWord';
{*����� �� Word'a}

Function PrintDialogWord(var wrd : Variant):boolean; stdcall; external 'msofficedll.dll' name 'PrintDialogWord';
{*����� ������� ������}

Procedure CreateTableEx(Col,Row : Integer; DefaultTableBehavior,AutoFitBehavior : Integer; var wrd : Variant);stdcall; external 'msofficedll.dll' name 'CreateTableEx';
{*������� ������� � ������������ �����������.
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
{                   ������� ��� ������������� M$ Excel                         }
{                                                                              }
{******************************************************************************}

 procedure NewXlsDocument(var xls : Variant; visible : Boolean);StdCall; external 'msofficedll.dll' name 'NewXlsDocument';
 {*������� ����� �������� Excel}

 procedure OpenXlsDocument(var xls : Variant; xlsfile : ShortString);StdCall; external 'msofficedll.dll' name 'OpenXlsDocument';
 {*������� ��������� ��������}

 function GetXlsWorkBook(var xls : Variant; idx : Integer): Variant;StdCall; external 'msofficedll.dll' name 'GetXlsWorkBook';
 {*�������� ������ �� �������� �����}

 function GetXlsWorkBookSheet(var WorkBook : Variant; idx : Integer):Variant;StdCall; external 'msofficedll.dll' name 'GetXlsWorkBookSheet';
 {*�������� ������ �� �������� ����}

 procedure SetCellValue(var Xls : Variant; CellName : ShortString; value : ShortString);StdCall; external 'msofficedll.dll' name 'SetCellValue';
 {*�������� �������� ��� ����� � ��������� ������}

 procedure SetCellValueInteger(var Xls : Variant; CellName : ShortString; value : Integer);StdCall; external 'msofficedll.dll' name 'SetCellValueInteger';
 {*�������� �������� ��� ����� ����� � ��������� ������}

 procedure SetCellValueFloat(var Xls : Variant; CellName : ShortString; value : Double);StdCall; external 'msofficedll.dll' name 'SetCellValueFloat';
 {*�������� �������� ��� ����� � ��������� ������ � ��������� ������}

 procedure SetCellValueDate(var Xls : Variant; CellName : ShortString; value : TDatetime);StdCall; external 'msofficedll.dll' name 'SetCellValueDate';
 {*�������� �������� ��� ���� � ��������� ������}

 procedure SetCellValueCurrency(var Xls : Variant; CellName : ShortString; value : Currency);StdCall; external 'msofficedll.dll' name 'SetCellValueFloat';
 {*�������� �������� � �������� �������� � ��������� ������}

 procedure SetCellValueFormat(var Xls : Variant; CellName : ShortString; valueformat : ShortString);StdCall; external 'msofficedll.dll' name 'SetCellValueFormat';
 {*������� ������ ������ �������� ������}

 {******************************************************************************}
 {                                                                              }
 {                   ������� ��� ������������� M$ Outlook                       }
 {              (����� �� ���������� outlookdll �� EmeraldMan)                  }
 {******************************************************************************}

 procedure OutLookConnect(var OL: Variant);StdCall; external 'msofficedll.dll' name 'OutLookConnect';
 {*������������ � OutLook}


 procedure OutLookNewFolder(var OL: Variant; s: ShortString);StdCall; external 'msofficedll.dll' name 'OutLookNewFolder';
 {*����� ����� ���������}

 procedure OutLookNewContact(var OL: Variant; folder:ShortString; name:ShortString);StdCall; external 'msofficedll.dll' name 'OutLookNewContact';
 {*����� �������
 folder - �������� ����� � OutLook ���������;
 name - ��� ��������
 ���� ����� folder �� ����������, �� ����� �������.
 ��� ������ ������� �������, ������ ��������� ������ ���;
 ��� ���������� ����� ���������� ��������
 OutlookContact.LastName
 OutlookContact.MiddleName
 OutlookContact.CompanyName
 Contact.HomeTelephoneNumber
 Contact.Email1Address
 � �.�. ���� ��� �����}

 procedure OutLookDisConnect(var OL: Variant);StdCall; external 'msofficedll.dll' name 'OutLookDisConnect';
 {*����������� �� OutLook}



implementation

end.
