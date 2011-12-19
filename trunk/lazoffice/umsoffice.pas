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
{* ������ ��� ������������� �������� ������� � M$ Word XP/2003 � Delphi/Lazarus.
��� ������������� ���������� � ������ ���������� ������ uofficedll.pas
(� Lazarus ����� ���������� ���������� ������ Variants � ������ ComObj).}
interface
  uses sysutils,variants,comobj;
 {$I const.inc}

 function CentimetersToPoints(cm : Real) : Real;
 {* ��������� ���������� � �����}

 procedure NewDocument(var Wrd : Variant; visible : Boolean);
 {* ������� ����� ��������.  �������� visible ���������, ����� �� ������������ ���� Word'� �� ������}


//******************************************************************************
 procedure PageMargins(l,r,t,b : Single; var wrd : Variant);
 {* ������������� ���� ��������}

 procedure PageOrientation(orientation : Integer; var wrd : Variant);
  {* ������������� ���������/������� ���������� ��������.
  wdOrientPortrait = 0;     //������
  wdOrientLandscape = 1;    //��������      }

 procedure HFDistance(h,f : Single; var wrd : Variant);
  {* ������������� ������� ��� �������� � ������� ������������}

 Procedure PageSize (w,h : Single; var wrd : Variant);
  {* ������������� ������ �������� � �����������. ��� �4 w=29.7, h=21}

 procedure SetOnPage(style : Integer; var Wrd : Variant);
 {* ������������� ���������� � ���������� ������ �� ��������.
  wdNormalPage = 0;   //1 � 1 - ����������
  wdTwoOnOne = 1;     //2 ����. �� 1
  wdBookFold = 2;     //������� }

 procedure PageAlign(align : Integer; var wrd : Variant);
  {* ������������� ������������ ������������ ������ �� ��������.
  wdAlignVerticalTop = 0;       //�� �������� ����
  wdAlignVerticalCenter = 1;    //�� ������
  wdAlignVerticalJustify = 2;   //��������� �� ������
  wdAlignVerticalBottom = 3;    //�� ������� ����    }

 procedure NewPage(wrd : Variant);
{* �������� ������ �������� � ��������}

//******************************************************************************
 procedure FontName(name : ShortString; var wrd : Variant);
  {* ������������� �������� ������ (����. 'Courier New')}

 procedure FontSize(sz : Integer; var Wrd : Variant);
 {* ������������� ������ ��� �������� ������}

 procedure FontBold(Bold : Boolean; var Wrd : Variant);
 {* ������������� �������� ��� �������� ������}

 procedure FontItalic(Italic : Boolean; var Wrd : Variant);
 {* ������������� ������ ��� �������� ������}

 procedure FontUnderlined(underlined : Boolean; var wrd : Variant);
 {* ������������� ������������� ��� �������� ������}

 procedure FontShadowed(Shadowed : Boolean; var Wrd : Variant);
 {* ������������� ���� ��� �������� ������}

 procedure FontColor(Color : Integer; var Wrd : Variant);
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

 procedure FontSuperScript(super : Boolean; var Wrd : Variant);
 {* ������������� ������� �������}

 procedure FontSubScript(sub : Boolean; var Wrd : Variant);
 {* ������������� ������ �������}

 procedure FontSpacing(spacing : Single; var Wrd : Variant);
 {* ������������� ������������� ��������: 1, 1.5 � �.�.}

 procedure FontScaling(scaling : Integer; var Wrd : Variant);
 {* ������������� ������� ��� �������� ������ � %}

 procedure FontPosition(position : Single; var Wrd : Variant);
 {* ������������� �������� ������ ����� (������������� ��������)
   ��� ���� (������������� ��������) � ��}

 procedure AddText(s : ShortString; var wrd : Variant);
{* �������� ������}

 procedure AddParagraph(var wrd : Variant);
 {* ������ ����� �����}

//******************************************************************************
 procedure ParagraphAlign(align : Integer; var Wrd : Variant);
{* ���������� ������������ ������ �� ������
  wdAlignParagraphLeft = 0;
  wdAlignParagraphCenter = 1;
  wdAlignParagraphRight = 2;
  wdAlignParagraphJustify = 3;}

 procedure ParagraphLineSpace(space : Integer; var wrd : Variant);
 {* ���������� ������������� ��������:
  wdLineSpaceSingle = 0;
  wdLineSpace1pt5 = 1;
  wdLineSpaceDouble = 2;
  wdLineSpaceAtLeast = 3;
  wdLineSpaceExactly = 4;
  wdLineSpaceMultiple = 5;}

 procedure ParagraphIndents(l,r : Single; var wrd : Variant);
 {* ���������� ������� ������ ����� � ������ (� ��)}

 procedure ParagraphSpaces(top,bottom : Single; var wrd : Variant);
 {* ���������� ������� ������ ������ � �����}

 Procedure ParagraphFirstLine(indent : Single; var wrd : Variant);
 {* ���������� ������ ������ ������ ������ (� ��)}

 //*****************************************************************************
 procedure AddTabPosition(pos : Single; var wrd : Variant);
{* �������� ������� ��������� � pos ��}

 procedure DefaultTabPos(pos : Single; var wrd : Variant);
 {* ���������� ������� ��������� �� ��������� � pos ��}

 procedure ClearAllTabs(var wrd : Variant);
 {* �������� ��� ������� ���������}

 //*****************************************************************************
 procedure CreateTable(Col,Row : Integer; var wrd : Variant);
 {* ������� ������� � ���������� �� ���������}

 Procedure SetColWidth(wid : Single; var wrd : Variant);
 {* ���������� ������ �������� ������� � �������.
  ������������ ����� ������ AddText(), ����� �������� ������������}

 Procedure SetRowHeight(h : Single; var wrd : Variant);
 {* ���������� ������ ������� ������ � �������.}

 procedure GotoRight(cells : Integer; var wrd : Variant);
 {* ������� �� cells ����� ������. ���� ������� �����������, ��������� ����� ������}

 procedure GotoLeft(cells : Integer; var wrd : Variant);
 {* ������� �� cells ����� �����.}

 procedure GotoUp(lines : Integer; var wrd : Variant);
 {* ������� �� lines ����� �����}

 procedure GotoDown(lines : Integer; var wrd : Variant);
 {* ������� �� lines ����� ����.}

 Procedure MergeCellsR(count : Integer; var Wrd : Variant);
 {* ���������� ��������� ������ ������. ������ ������ ���������� � ������ �� ���}

 Procedure MergeCellsD(count : Integer; var Wrd : Variant);
 {* ���������� ��������� ������ ����. ������ ������ ���������� � ������ �� ���}

 procedure DeleteRow(var Wrd : Variant);
 {* ������� ������� ������ �������}

 procedure DeleteCol(var Wrd : Variant);
 {* ������� ������� ������� �������}

 Procedure ExitTable(wrd : Variant);
{* ����� �� �������}

 procedure CellTextOrientation(orient : Integer;wrd : Variant);
 {* ����������� ������ � �������:
  wdTextOrientationHorizontal = 0;
  wdTextOrientationUpward = 2;
  wdTextOrientationDownward = 3;
  wdTextOrientationVerticalFarEast = 1;
  wdTextOrientationHorizontalRotatedFarEast = 4;}

 procedure InsertHeader(hdr : ShortString; var wrd : Variant);
{* �������� ������� ����������}

 procedure InsertFooter(ftr : ShortString; var wrd : Variant);
{* �������� ������ ����������}

procedure SetWordVisible(var wrd : Variant; Visible : Boolean);
{*������� ���� ������� �� ������}

function CheckWordVersion(var wrd : Variant):Boolean;
{*��������� ����� �� �������������� ������������� ������ �����}

Function SaveDocAs(file_:Shortstring; var wrd : Variant):boolean;
{*��������� ��������� �������� � ��������� ������ � �����}

Function CloseDoc(var wrd : Variant):boolean;
{*������� ��������}

Function CloseWord(var wrd : Variant):boolean;
{*����� �� Word'a}

Function PrintDialogWord(var wrd : Variant):boolean;
{*����� ������� ������}

Procedure CreateTableEx(Col,Row : Integer; DefaultTableBehavior,AutoFitBehavior : Integer; var wrd : Variant);
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

 procedure NewXlsDocument(var xls : Variant; visible : Boolean);
 {*������� ����� �������� Excel}

 procedure OpenXlsDocument(var xls : Variant; xlsfile : ShortString);
 {*������� ��������� ��������}

 function GetXlsWorkBook(var xls : Variant; idx : Integer): Variant;
 {*�������� ������ �� �������� �����}

 function GetXlsWorkBookSheet(var WorkBook : Variant; idx : Integer):Variant;
 {*�������� ������ �� �������� ����}

 procedure SetCellValue(var Xls : Variant; CellName : ShortString; value : ShortString);
 {*�������� �������� ��� ����� � ��������� ������}

 procedure SetCellValueInteger(var Xls : Variant; CellName : ShortString; value : Integer);
 {*�������� �������� ��� ����� ����� � ��������� ������}

 procedure SetCellValueFloat(var Xls : Variant; CellName : ShortString; value : Double);
 {*�������� �������� ��� ����� � ��������� ������ � ��������� ������}

 procedure SetCellValueDate(var Xls : Variant; CellName : ShortString; value : TDatetime);
 {*�������� �������� ��� ���� � ��������� ������}

 procedure SetCellValueCurrency(var Xls : Variant; CellName : ShortString; value : Currency);
 {*�������� �������� � �������� �������� � ��������� ������}

 procedure SetCellValueFormat(var Xls : Variant; CellName : ShortString; valueformat : ShortString);
 {*������� ������ ������ �������� ������}

 {******************************************************************************}
 {                                                                              }
 {                   ������� ��� ������������� M$ Outlook                       }
 {              (����� �� ���������� outlookdll �� EmeraldMan)                  }
 {******************************************************************************}

 procedure OutLookConnect(var OL: Variant);
 {*������������ � OutLook}


 procedure OutLookNewFolder(var OL: Variant; s: ShortString);
 {*����� ����� ���������}

 procedure OutLookNewContact(var OL: Variant; folder:ShortString; name:ShortString);
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

 procedure OutLookDisConnect(var OL: Variant);
 {*����������� �� OutLook}



implementation
function CentimetersToPoints(cm : Real) : Real;StdCall;
 {*��������� ���������� � �����}
 begin
  result := cm * cm2p;
 end;


{******************************************************************************}
{                                                                              }
{                   ������� ��� ������������� M$ Word                          }
{                                                                              }
{******************************************************************************}
 procedure NewDocument(var Wrd : Variant; visible : Boolean);
 {*������� ����� ��������}
 begin
  Wrd := CreateOleObject('Word.Application');
  Wrd.Visible := Visible;
  Wrd.Documents.Add;
  Wrd.Application.WindowState := wdWindowStateMaximize;
 end;

//***********************************************************************************
 procedure PageMargins(l,r,t,b : Single; var wrd : Variant);
 {*������������� ���� ��������}
 begin
  Wrd.ActiveDocument.PageSetup.LeftMargin := CentimetersToPoints(l);
  Wrd.ActiveDocument.PageSetup.RightMargin := CentimetersToPoints(r);
  Wrd.ActiveDocument.PageSetup.TopMargin := CentimetersToPoints(t);
  Wrd.ActiveDocument.PageSetup.BottomMargin := CentimetersToPoints(b);
 end;

 procedure PageOrientation(orientation : Integer; var wrd : Variant);
 {*������������� ���������/������� ���������� �������� }
 begin
  Wrd.ActiveDocument.PageSetup.Orientation := orientation;
 end;

 procedure HFDistance(h,f : Single; var wrd : Variant);
 {*������������� ������� ��� �������� � ������� ������������}
 begin
  Wrd.ActiveDocument.PageSetup.HeaderDistance := CentimetersToPoints(h);
  Wrd.ActiveDocument.PageSetup.FooterDistance := CentimetersToPoints(f);
 end;

 Procedure PageSize (w,h : Single; var wrd : Variant);
 {*������������� ������ �������� � �����������. ��� �4 w=29.7, h=21}
 begin
  Wrd.ActiveDocument.PageSetup.PageWidth := CentimetersToPoints(w);
  Wrd.ActiveDocument.PageSetup.PageHeight := CentimetersToPoints(h);
 end;

 procedure SetOnPage(style : Integer; var Wrd : Variant);
 {*������������� ���������� � ���������� ������ �� ��������.
  wdNormalPage = 0;   //1 � 1 - ����������
  wdTwoOnOne = 1;     //2 ����. �� 1
  wdBookFold = 2;     //������� }
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
 {*������������� ������������ ������������ ������ �� ��������.
  wdAlignVerticalTop = 0;       //�� �������� ����
  wdAlignVerticalCenter = 1;    //�� ������
  wdAlignVerticalJustify = 2;   //��������� �� ������
  wdAlignVerticalBottom = 3;    //�� ������� ����    }
 begin
  Wrd.ActiveDocument.PageSetup.VerticalAlignment := align;
 end;

 procedure NewPage(wrd : Variant);
{*�������� ������ �������� � ��������}
 begin
  wrd.Selection.InsertBreak(wdPageBreak);
 end;

//*******************************************************************************
 procedure FontName(name : ShortString; var wrd : Variant);
 {*������������� �������� ������}
 begin
  Wrd.Selection.Font.Name := name;
 end;

 procedure FontSize(sz : Integer; var Wrd : Variant);
 {*������������� ������ ��� �������� ������}
 begin
  Wrd.Selection.Font.Size := sz;
 end;

 procedure FontBold(Bold : Boolean; var Wrd : Variant);
 {*������������� �������� ��� �������� ������}
 begin
  Wrd.Selection.Font.Bold := Bold;
 end;

 procedure FontItalic(Italic : Boolean; var Wrd : Variant);
 {*������������� ������ ��� �������� ������}
 begin
  Wrd.Selection.Font.Italic := Italic;
 end;

 procedure FontUnderlined(underlined : Boolean; var wrd : Variant);
 {*������������� ������������� ��� �������� ������}
 begin
  if underlined then wrd.Selection.Font.Underline := wdUnderlineSingle
                else wrd.Selection.Font.Underline := wdUnderlineNone;
 end;

 procedure FontShadowed(Shadowed : Boolean; var Wrd : Variant);
 {*������������� ���� ��� �������� ������}
 begin
  Wrd.Selection.Font.Shadow := Shadowed;
 end;

 procedure FontColor(Color : Integer; var Wrd : Variant);
 {*������������� ���� ��� ��������  ������
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
 {*������������� ������� �������}
 begin
  Wrd.Selection.Font.Superscript := super;
 end;

 procedure FontSubScript(sub : Boolean; var Wrd : Variant);
 {*������������� ������ �������}
 begin
  Wrd.Selection.Font.Subscript := sub;
 end;

 procedure FontSpacing(spacing : Single; var Wrd : Variant);
 {*������������� ������������� ��������: 1, 1.5 � �.�.}
 begin
  Wrd.Selection.Font.Spacing := spacing;
 end;

 procedure FontScaling(scaling : Integer; var Wrd : Variant);
 {*������������� ������� ��� �������� ������ � %}
 begin
  Wrd.Selection.Font.Scaling := scaling;
 end;

 procedure FontPosition(position : Single; var Wrd : Variant);
 {*������������� �������� ������ ����� (������������� ��������)
   ��� ���� (������������� ��������) � ��}
 begin
  Wrd.Selection.Font.Position := position;
 end;

 procedure AddText(s : ShortString; var wrd : Variant);
 {*�������� ������}
 begin
  Wrd.Selection.TypeText(s);
 end;

 procedure AddParagraph(var wrd : Variant);
 {*������ ����� �����}
 begin
  Wrd.Selection.TypeParagraph;
 end;

//***************************************************************************************
 procedure ParagraphAlign(align : Integer; var wrd : Variant);
 {*���������� ������������ ������ �� ������
  wdAlignParagraphLeft = 0;
  wdAlignParagraphCenter = 1;
  wdAlignParagraphRight = 2;
  wdAlignParagraphJustify = 3;}
 begin
  Wrd.Selection.ParagraphFormat.Alignment := align;
 end;

 procedure ParagraphLineSpace(space : Integer; var wrd : Variant);
 {*���������� ������������� ��������:
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
 {*���������� ������� ������ ����� � ������ (� ��)}
 begin
  Wrd.Selection.ParagraphFormat.LeftIndent := CentimetersToPoints(l);
  Wrd.Selection.ParagraphFormat.RightIndent := CentimetersToPoints(r);
 end;

 procedure ParagraphSpaces(top,bottom : Single; var wrd : Variant);
 {*���������� ������� ������ ������ � �����}
 begin
  Wrd.Selection.ParagraphFormat.SpaceBefore := top;
  Wrd.Selection.ParagraphFormat.SpaceBeforeAuto := false;
  Wrd.Selection.ParagraphFormat.SpaceAfter := bottom;
  Wrd.Selection.ParagraphFormat.SpaceAfterAuto := false;
 end;

 Procedure ParagraphFirstLine(indent : Single; var wrd : Variant);
 {*���������� ������ ������ ������ ������ (� ��)}
 begin
   wrd.Selection.ParagraphFormat.FirstLineIndent := CentimetersToPoints(indent);
 end;

//*********************************************************************************
 procedure AddTabPosition(pos : Single; var wrd : Variant);
 {*�������� ������� ��������� � pos ��}
 begin
  Wrd.Selection.ParagraphFormat.TabStops.Add(CentimetersToPoints(pos),wdAlignTabLeft,wdTabLeaderSpaces);
 end;

 procedure DefaultTabPos(pos : Single; var wrd : Variant);
 {*���������� ������� ��������� �� ��������� � pos ��}
 begin
  Wrd.Selection.ParagraphFormat.DefaultTabStop := CentimetersToPoints(pos);
 end;

 procedure ClearAllTabs(var wrd : Variant);
 {*�������� ��� ������� ���������}
 begin
  Wrd.Selection.ParagraphFormat.TabStops.ClearAll;
 end;

//***********************************************************************************
 procedure CreateTable(Col,Row : Integer; var wrd : Variant);
 {*������� ������� � ���������� �� ���������}
 begin
   Wrd.ActiveDocument.Tables.Add(Wrd.Selection.Range,row,col,wdWord9TableBehavior,wdAutoFitWindow);
 end;

 Procedure SetColWidth(wid : Single; var wrd : Variant);
 {*���������� ������ �������� ������� � �������.
  ������������ ����� ������ AddText(), ����� �������� ������������}
 begin
  Wrd.Selection.Columns.SetWidth(wid, wdAdjustProportional);
 end;

 procedure SetRowHeight(h: Single; var wrd: Variant);
 begin
   Wrd.Selection.Rows.SetHeight(h,wdRowHeightAuto);
 end;

 procedure GotoRight(cells : Integer; var wrd : Variant);
 {*������� �� cells ����� ������. ���� ������� �����������, ��������� ����� ������}
 var i : Integer;
 begin
  for i := 1 to cells do Wrd.Selection.MoveRight(wdCell);
 end;

 procedure GotoLeft(cells : Integer; var wrd : Variant);
 {*������� �� cells ����� �����.}
 var i : Integer;
 begin
  for i := 1 to cells do Wrd.Selection.MoveLeft(wdCell);
 end;

 procedure GotoUp(lines : Integer; var wrd : Variant);
 {*������� �� lines ����� �����}
 begin
  Wrd.Selection.MoveUp(wdLine,lines);
 end;

 procedure GotoDown(lines : Integer; var wrd : Variant);
 {*������� �� lines ����� ����.}
 begin
  Wrd.Selection.MoveDown(wdLine,lines);
 end;

 Procedure MergeCellsR(count : Integer; var Wrd : Variant);
 {*���������� ��������� ������ ������. ������ ������ ���������� � ������ �� ���}
 begin
  Wrd.Selection.MoveRight(wdCharacter,count,wdExtend);
  Wrd.Selection.Cells.Merge;
 end;

 Procedure MergeCellsD(count : Integer; var Wrd : Variant);
 {*���������� ��������� ������ ����. ������ ������ ���������� � ������ �� ���}
 begin
  Wrd.Selection.MoveDown(wdLine,count,wdExtend);
  Wrd.Selection.Cells.Merge;
 end;

 procedure DeleteRow(var Wrd : Variant);
 {*������� ������� ������ �������}
 begin
  Wrd.Selection.Rows.Delete;
 end;

 procedure DeleteCol(var Wrd : Variant);
 {*������� ������� ������� �������}
 begin
  Wrd.Selection.Columns.Delete;
 end;

 Procedure ExitTable(wrd : Variant);
 {*����� �� �������}
 begin
  Wrd.Selection.MoveDown(wdLine,1);
 end;

 procedure CellTextOrientation(orient : Integer;wrd : Variant);
 {*����������� ������ � �������
  wdTextOrientationHorizontal = 0;
  wdTextOrientationUpward = 2;
  wdTextOrientationDownward = 3;
  wdTextOrientationVerticalFarEast = 1;
  wdTextOrientationHorizontalRotatedFarEast = 4;}
 begin
  wrd.Selection.Orientation := orient;
 end;

 procedure InsertHeader(hdr : ShortString; var wrd : Variant);
{*�������� ������� ����������}
 begin
  wrd.ActiveWindow.ActivePane.View.SeekView := wdSeekCurrentPageHeader;
  wrd.Selection.ParagraphFormat.Alignment := wdAlignParagraphRight;
  FontColor(wdDarkYellow,wrd);
  FontSize(6,wrd);
  wrd.Selection.TypeText(hdr);
  wrd.ActiveWindow.ActivePane.View.SeekView := wdSeekMainDocument;
 end;

procedure InsertFooter(ftr : ShortString; var wrd : Variant);
{*�������� ������ ����������}
 begin
  wrd.ActiveWindow.ActivePane.View.SeekView := wdSeekCurrentPageFooter;
  wrd.Selection.ParagraphFormat.Alignment := wdAlignParagraphRight;
  FontColor(wdDarkYellow,wrd);
  FontSize(6,wrd);
  wrd.Selection.TypeText(ftr);
  wrd.ActiveWindow.ActivePane.View.SeekView := wdSeekMainDocument;
 end;

procedure SetWordVisible(var wrd : Variant; Visible : Boolean);
{*������� ���� ������� �� ������}
begin
 wrd.Visible := Visible;
end;

function CheckWordVersion(var wrd : Variant):Boolean;
{*��������� ����� �� �������������� ������������� ������ �����}
begin
 result := false;
 if wrd.Version>9 then result := true;
end;

Function SaveDocAs(file_:Shortstring; var wrd : Variant):boolean;
{*��������� ��������� �������� � ��������� ������ � �����}
begin
 SaveDocAs:=true;
 try
 Wrd.ActiveDocument.SaveAs(file_);
 except
 SaveDocAs:=false;
 end;
End;

Function CloseDoc(var wrd : Variant):boolean;
{*������� ��������}
begin
 CloseDoc:=true;
 try
  Wrd.ActiveDocument.Close;
 except
  CloseDoc:=false;
 end;
End;

Function CloseWord(var wrd : Variant):boolean;
{*����� �� Word'a}
begin
 CloseWord:=true;
 try
  Wrd.Quit;
 except
  CloseWord:=false;
 end;
End;

Function PrintDialogWord(var wrd : Variant):boolean;
{*����� ������� ������}
begin
 PrintDialogWord:=true;
 try
  Wrd.Dialogs.Item(wdDialogFilePrint).Show;
 except
  PrintDialogWord:=false;
 end;
End;

Procedure CreateTableEx(Col,Row : Integer; DefaultTableBehavior,AutoFitBehavior : Integer; var wrd : Variant);
{*������� ������� � ������������ �����������.
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
{                   ������� ��� ������������� M$ Excel                         }
{                                                                              }
{******************************************************************************}

 procedure NewXlsDocument(var xls : Variant; visible : Boolean);
 {*������� ����� �������� Excel}
 begin
  xls := CreateOleObject('Excel.Application');
  xls.Visible := Visible;
  xls.WorkBooks.Add;
 end;

 procedure OpenXlsDocument(var xls : Variant; xlsfile : ShortString);
 {*������� ��������� ��������}
 begin
  xls := CreateOleObject('Excel.Application');
  if xlsfile<>'' then xls.WorkBooks.Open(xlsfile);
  xls.Visible := true;
 end;

 function GetXlsWorkBook(var xls : Variant; idx : Integer): Variant;
 {*�������� ������ �� �������� �����}
 var workbooks : Variant;
 begin
   result := Null;
   WorkBooks := xls.WorkBooks;
   if idx<workbooks.Count then Result := WorkBooks.item[idx];
 end;

 function GetXlsWorkBookSheet(var WorkBook : Variant; idx : Integer):Variant;
 {*�������� ������ �� �������� ��������}
 begin
  result := Null;
  if idx<WorkBook.Sheets.Count then Result := WorkBook.Sheets.Item[idx];
 end;

 procedure SetCellValue(var Xls : Variant; CellName : ShortString; value : ShortString);
 {*�������� �������� ��� ����� � ��������� ������}
 begin
  Xls.ActiveSheet.Range[CellName].Value := value;
 end;

 procedure SetCellValueInteger(var Xls : Variant; CellName : ShortString; value : Integer);
 {*�������� �������� ��� ����� ����� � ��������� ������}
 begin
 //Xls.ActiveSheet.Range[CellName].NumberFormat:='0';
  Xls.ActiveSheet.Range[CellName].Value := value;
 end;

 procedure SetCellValueFloat(var Xls : Variant; CellName : ShortString; value : Double);
 {*�������� �������� ��� ����� � ��������� ������ � ��������� ������}
 var Range : Variant;
 begin
  Range := Xls.ActiveSheet.Range[CellName];
  Range.Value:=value;
//  Range.NumberFormat:='General';
 end;

 procedure SetCellValueCurrency(var Xls : Variant; CellName : ShortString; value : Currency);
 {*�������� �������� ��� �������� � ��������� ������}
 begin
  Xls.ActiveSheet.Range[CellName].Value := value;
//  Xls.ActiveSheet.Range[CellName].NumberFormat:='0.00';
 end;

 procedure SetCellValueFormat(var Xls : Variant; CellName : ShortString; valueformat : ShortString);
 {*������� ������ ������ �������� ������}
 begin
  Xls.ActiveSheet.Range[CellName].NumberFormat:=valueformat;
 end;

 procedure SetCellValueDate(var Xls : Variant; CellName : ShortString; value : TDatetime);
 {*�������� �������� ��� ���� � ��������� ������}
 begin
//  Xls.ActiveSheet.Range[CellName].NumberFormat:='dd.mm.yyyy';
  Xls.ActiveSheet.Range[CellName].Value := Value;
 end;

 {******************************************************************************}
 {                                                                              }
 {                   ������� ��� ������������� M$ Outlook                       }
 {              (����� �� ���������� outlookdll �� EmeraldMan)                  }
 {******************************************************************************}

procedure OutLookConnect(var OL: Variant);
{*������������ � OutLook}
begin
  OL := CreateOleObject('Outlook.Application');
end;

procedure OutLookNewFolder(var OL: Variant; s: ShortString);
{*����� ����� ���������}
begin
  OL.GetNameSpace('MAPI').GetDefaultFolder(olFolderContacts).AddFolder(s);
end;

procedure OutLookNewContact(var OL: Variant; folder:ShortString; name:ShortString);
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
{*����������� �� OutLook}
begin
  OL := Unassigned;
end;


end.
