VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'+==+==+==+==+==+==+==+==+==+==+==+==+==+
' MSHFlexGrid Printer class
' Author: Opal Raj Ghimire, buna48@hotmail.com
' http://geocities.com/opalraj/vb
' You are free to use this class in your projects.
' This class comes with absolutely
' NO WAR ranty ! Use it at your own risk !!!
' +==+==+==+==+==+==+==+==+==+==+==+==+==+
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Dim objFlex As Object
Dim lRowsFrom As Long, lRowsTo As Long
Dim FinalX As Long, FinalY As Long
'***********Publics***********
Public GridPrint As Boolean
Public VSpace As Long, HSpace As Long
Public RoundCorX As Long, RoundCorY As Long
Public GridPenStyle As Long
Public BackColor As Long
Public DrawBoarder As Boolean
Public BoarderStyle As Long, BoarderColor As Long, BoarderWidth As Long, BoarderDistance As Long
Public PosLeft As Long, PosTop As Long
Public PixHeight As Integer, PixWidth As Integer
Public PicturePrint As Boolean
Public EnableCellBackColor As Boolean
'**********Const**************
Const DT_LEFT = 0
Const DT_TOP = 0
Const DT_CENTER = 1
Const DT_RIGHT = 2
Const DT_VCENTER = 4
Const DT_BOTTOM = 8
Const DT_WORDBREAK = 16
Const DT_SINGLELINE = 32
Const DT_NOPREFIX = 2048
Const DT_END_ELLIPSIS = 32768
Const DT_MODIFYSTRING = 65536
Const DT_WORD_ELLIPSIS = 262144
Const DT_RTLREADING = 131072  'For RightToLeft people
Const DT_CALCRECT = 1024



Public Sub PrintOut(Obj As Object)
Dim lRows As Long, lCols As Long, cellHeight As Long, cellWidth As Long
Dim tmpLeft As Long, rectBox As RECT, rectBox2 As RECT, Rect4Pic As RECT, TmpRect As RECT
Dim lDrawWidth As Long, sCellText As String, cellFont As StdFont, lTextColor, iAlignment As Integer
Dim iCellAlign As Integer, DT_Code As Long, bWordWrap As Boolean, lGridLineColor As Long, lOldPen As Long, lNewPen As Long, lOldBrush As Long, lNewBrush As Long
Dim BorX As Long, BorY As Long, BorDx As Long, BorDy As Long, opal As POINTAPI, TmpDrawWidth As Long, TmpBoarderWidth As Long
Dim TmpBoarderDistance As Long, BackupPosLeft As Long, BackupPosTop As Long, RetHeight As Long
Dim BackUpScaleMode As Long, YFac As Single, XFac As Single, lCellBackCol As Long
'*********************************************************
If Not ValidObj(Obj) Then Exit Sub
Set cellFont = New StdFont
BackUpScaleMode = Obj.ScaleMode
Obj.ScaleMode = vbPixels
Obj.Print
With objFlex
        lDrawWidth = .GridLineWidth
        bWordWrap = .WordWrap
        .Redraw = False
        lGridLineColor = .GridColor
End With 'objFlex
BackupPosLeft = PosLeft: BackupPosTop = PosTop
TmpBoarderWidth = BoarderWidth: TmpDrawWidth = lDrawWidth: TmpBoarderDistance = BoarderDistance

YFac = 1: XFac = 1
If TypeName(Obj) = "Printer" Then
        XFac = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
        YFac = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY
        lDrawWidth = lDrawWidth * XFac
        BoarderWidth = BoarderWidth * XFac
End If  'TypeName(obj) = "Printer"
If DrawBoarder Then
        PosTop = PosTop + (BoarderDistance * YFac)
        PosLeft = PosLeft + (BoarderDistance * XFac)
        BorX = PosLeft: BorY = PosTop
End If 'DrawBoarder
If GridPrint Then  'Creating the pen
         lNewPen = CreatePen(GridPenStyle, lDrawWidth, lGridLineColor)
         lOldPen = SelectObject(Obj.hDC, lNewPen)
End If 'GridPrint
For lRows = lRowsFrom To lRowsTo
        cellHeight = Obj.ScaleY(objFlex.RowHeight(lRows), vbTwips, vbPixels)
        tmpLeft = PosLeft
                For lCols = 0 To objFlex.Cols - 1
                        With objFlex
                            .Col = lCols: .Row = lRows
                            sCellText = objFlex.Text
                            cellFont.Name = .CellFontName: cellFont.Size = .CellFontSize
                            cellFont.Bold = .CellFontBold: cellFont.Underline = .CellFontUnderline
                            cellFont.Strikethrough = .CellFontStrikeThrough: cellFont.Italic = .CellFontItalic
                            lTextColor = .CellForeColor: iCellAlign = .CellAlignment
                                 If bWordWrap Then DT_Code = DT_WORDBREAK Or DT_NOPREFIX Or DT_MODIFYSTRING Else DT_Code = DT_SINGLELINE Or DT_NOPREFIX
                                    Select Case iCellAlign
                                        Case 0 'this should be Case 9 in case of MSFlex Grid Control, see case 9 too
                                            DT_Code = DT_Code Or DT_LEFT Or DT_TOP
                                        Case 1
                                            DT_Code = DT_Code Or DT_VCENTER Or DT_LEFT
                                        Case 2
                                            DT_Code = DT_Code Or DT_LEFT Or DT_BOTTOM
                                        Case 3
                                            DT_Code = DT_Code Or DT_CENTER Or DT_TOP
                                        Case 4
                                            DT_Code = DT_Code Or DT_VCENTER Or DT_CENTER
                                        Case 5
                                            DT_Code = DT_Code Or DT_CENTER Or DT_BOTTOM
                                        Case 6
                                            DT_Code = DT_Code Or DT_RIGHT Or DT_TOP
                                        Case 7
                                            DT_Code = DT_Code Or DT_RIGHT Or DT_VCENTER
                                        Case 8
                                            DT_Code = DT_Code Or DT_BOTTOM Or DT_RIGHT
                                        Case 9 'this should be Case 0 in case of MSFlex Grid Control
                                            If IsNumeric(sCellText) Or IsDate(sCellText) Then DT_Code = DT_Code Or DT_RIGHT Or DT_VCENTER Else DT_Code = DT_Code Or DT_LEFT Or DT_VCENTER
                                    End Select 'iCellAlign
                        End With 'objFlex
                        cellWidth = Obj.ScaleX(objFlex.ColWidth(lCols), vbTwips, vbPixels)
                        With rectBox
                                .Left = PosLeft: .Top = PosTop
                                .Right = PosLeft + cellWidth: .Bottom = PosTop + cellHeight
                                 If GridPrint Then
                                        lCellBackCol = BackColor
                                        If EnableCellBackColor Then lCellBackCol = objFlex.CellBackColor
                                        lNewBrush = CreateSolidBrush(lCellBackCol)
                                        lOldBrush = SelectObject(Obj.hDC, lNewBrush)
                                        RoundRect Obj.hDC, .Left, .Top, .Right + lDrawWidth, .Bottom + lDrawWidth, Obj.ScaleX(RoundCorX, vbTwips, vbPixels), Obj.ScaleY(RoundCorY, vbTwips, vbPixels)
                                        SelectObject Obj.hDC, lOldBrush: DeleteObject lNewBrush
                                 End If 'bGridPrint
                                 'Making rectBox a bit smaller, so that the text should not touch the lines
                                 TmpRect = rectBox   'to use it while aligning picture
                                .Left = .Left + (3 * XFac): .Right = .Right - (3 * XFac): .Top = .Top + (3 * YFac): .Bottom = .Bottom - (3 * YFac)
                      End With  ' rectBox
                      Set Obj.Font = cellFont
                      SetTextColor Obj.hDC, lTextColor
                      If bWordWrap And (iCellAlign <> 3 Or iCellAlign <> 6) Then
                                rectBox2 = rectBox
                                RetHeight = DrawText(Obj.hDC, sCellText, Len(sCellText), rectBox2, DT_CALCRECT Or DT_Code)
                                rectBox2.Left = 0: rectBox2.Top = 0
                                rectBox2.Bottom = RetHeight: rectBox2.Right = rectBox.Right - rectBox.Left
                                AlignRECT rectBox, rectBox2, ByVal iCellAlign
                                DrawText Obj.hDC, sCellText, Len(sCellText), rectBox2, DT_Code 'Or DT_RTLREADING  'removing first comment mark may help right to left system, can't check that in my computer
                       Else
                                DrawText Obj.hDC, sCellText, Len(sCellText), rectBox, DT_Code 'Or DT_RTLREADING  'removing first comment mark may help right to left system, can't check that in my computer
                       End If 'bWordWrap And (TmpJu...
                        '************Picture************
                        If PicturePrint Then
                                If objFlex.CellPicture <> 0 Then
                                      
                                        TmpRect.Left = TmpRect.Left + (2 * XFac): TmpRect.Right = TmpRect.Right - (2 * XFac)
                                        TmpRect.Top = TmpRect.Top + (2 * YFac): TmpRect.Bottom = TmpRect.Bottom - (2 * YFac)
                                        RetHeight = objFlex.CellPictureAlignment 'RetHeight=Picture Alignment
                                           If RetHeight = 9 Then
                                                   Obj.PaintPicture objFlex.CellPicture, rectBox.Left, rectBox.Top, (rectBox.Right - rectBox.Left), (rectBox.Bottom - rectBox.Top), 0, 0, , , vbSrcCopy     'vbSrcAnd
                                            Else
                                                    Rect4Pic.Left = 0: Rect4Pic.Top = 0
                                                    Rect4Pic.Right = (PixWidth + 1) * XFac
                                                    Rect4Pic.Bottom = (PixHeight + 1) * YFac
                                                    AlignRECT TmpRect, Rect4Pic, ByVal RetHeight
                                                    Obj.PaintPicture objFlex.CellPicture, Rect4Pic.Left, Rect4Pic.Top, , , 0, 0, Rect4Pic.Right - Rect4Pic.Left, Rect4Pic.Bottom - Rect4Pic.Top, vbSrcCopy   'vbSrcAnd
                                            End If  'RetHeight = 9
                                End If 'objFlex.CellPicture <> 0
                        End If 'bPicturePrint
                        '*********End of Picture****************
                        PosLeft = PosLeft + cellWidth + (VSpace * XFac) 'VSpace = Vertical Spacing
                        If (cellWidth = 0 And VSpace <> 0) Then PosLeft = PosLeft - (VSpace * XFac)
                Next lCols    'For lCols = 1 To flex.cols
        PosTop = PosTop + cellHeight + (HSpace * YFac) 'Horizontal Spacing
        If (cellHeight = 0 And HSpace <> 0) Then PosTop = PosTop - (HSpace * YFac)
        PosLeft = tmpLeft
Next lRows   'For lRows = lRowsFrom To lRowsTo
BorDx = rectBox.Right + (3 * XFac) + lDrawWidth
BorDy = rectBox.Bottom + (3 * YFac) + lDrawWidth
FinalX = BorDx: FinalY = BorDy
If GridPrint Then SelectObject Obj.hDC, lOldPen: DeleteObject lNewPen

If DrawBoarder Then
        lNewPen = CreatePen(BoarderStyle, BoarderWidth, BoarderColor)
        lOldPen = SelectObject(Obj.hDC, lNewPen)
        BoarderDistance = BoarderDistance * XFac
        BorX = BorX - BoarderDistance: BorY = BorY - BoarderDistance
        BorDx = BorDx + BoarderDistance: BorDy = BorDy + BoarderDistance
        MoveToEx Obj.hDC, BorDx, BorDy, opal  'Starting point Dx Dy
        LineTo Obj.hDC, BorX, BorDy        'MoveRight
        LineTo Obj.hDC, BorX, BorY          'MoveUp
        LineTo Obj.hDC, BorDx, BorY        'Moveleft
        LineTo Obj.hDC, BorDx, BorDy      'MoveDown to starting point
        SelectObject Obj.hDC, lOldPen
        DeleteObject lNewPen
        FinalX = BorDx: FinalY = BorDy
End If
objFlex.Redraw = True
Set cellFont = Nothing
BoarderWidth = TmpBoarderWidth: lDrawWidth = TmpDrawWidth
BoarderDistance = TmpBoarderDistance
Obj.ScaleMode = BackUpScaleMode
PosLeft = BackupPosLeft:  PosTop = BackupPosTop
End Sub


Public Property Set FlexName(ByVal NewValue As Object)
If TypeName(NewValue) = "MSHFlexGrid" Or TypeName(NewValue) = "MSFlexGrid" Then
    Set objFlex = NewValue
    objFlex.Refresh
Else
    MsgBox "Value to 'FlexName' property is invalid; it should be name of the MSHFlexGrid control or MSFlexGrid control. ", vbCritical, "Invalid Property"
End If
End Property
Public Property Let RowsTo(ByVal NewValue As Long)
lRowsTo = NewValue
End Property

Public Property Let RowsFrom(ByVal NewValue As Long)
lRowsFrom = NewValue
End Property

Private Sub Class_Initialize()
lRowsFrom = 0: lRowsTo = 1
GridPrint = True
VSpace = 0: HSpace = 0   'V=Vertical  H=Horizontal
RoundCorX = 10: RoundCorY = 10: GridPenStyle = 0
BackColor = vbWhite: DrawBoarder = True
BoarderStyle = 0: BoarderColor = vbBlack: BoarderWidth = 1: BoarderDistance = 5
PosLeft = 10: PosTop = 10
PicturePrint = False
PixHeight = 32
PixWidth = 32
EnableCellBackColor = False
End Sub

Public Property Get CurX() As Long
CurX = FinalX
End Property
Public Property Get CurY() As Long
CurY = FinalY
End Property

Private Function ValidObj(objOP As Object) As Boolean
Dim Msg As String
If TypeOf objOP Is Form Or TypeOf objOP Is PictureBox Or TypeOf objOP Is Printer Then
        ValidObj = True
Else
        ValidObj = False
        Msg = "Invalid object " + vbCrLf
        Msg = Msg + vbCrLf + vbCrLf
        Msg = Msg + "The valid objects are Printer, PictureBox and Form" + vbCrLf
        MsgBox Msg, vbCritical, "Invalid Object"
End If
End Function

Private Sub Class_Terminate()
Set objFlex = Nothing
End Sub
'This draws Horizontal ruler
'This method can be removed, if ur not using it
Public Sub DrawRulerH(Obj As Object, X As Long, Y As Long, Length As Long, sUnit As String)
Dim Height As Long, Devide As Long, SmallSteps As Long, bSwitch As Boolean
Dim h1 As Long, TwpPerUnit As Long, TmpK1 As Long, K1 As Long, K2 As Long, N As Long
Dim BuFore As Long, BuDrawStyle As Long, BuDrawWidth As Long, BuDrawMode As Long, BUScaleMode As Long
Dim BuFontS As Long
If Not ValidObj(Obj) Then Exit Sub

'Back Up
BuDrawStyle = Obj.DrawStyle: BuDrawWidth = Obj.DrawWidth
BuDrawMode = Obj.DrawMode: BuFontS = Obj.FontSize
BUScaleMode = Obj.ScaleMode
BuFore = Obj.ForeColor
'Set
Obj.DrawStyle = vbSolid: Obj.DrawWidth = 1: Obj.DrawMode = vbCopyPen
Obj.ScaleMode = vbTwips: Obj.FontSize = 8
Obj.FontBold = False: Obj.FontItalic = False
Obj.FontStrikethru = False
Obj.FontUnderline = False
SetTextColor Obj.hDC, vbBlack
X = X * Screen.TwipsPerPixelX
Y = Y * Screen.TwipsPerPixelY

Height = 400

If UCase(sUnit) = "INCH" Then
TwpPerUnit = 1440
Devide = 20
Else
'CM
TwpPerUnit = 567
Devide = 6
End If

SmallSteps = TwpPerUnit / Devide

Y = Y - (Height + 50)

Obj.Line (X, Y)-(X + Length, Y)
Obj.Line (X, Y + Height)-(X + Length, Y + Height)
N = -1
For K1 = 0 To Length Step TwpPerUnit
bSwitch = True: N = N + 1
Obj.Line (X + K1, Y)-(X + K1, Y + Height + 50), vbRed
''''''
Obj.Line (X + K1, Y + Height + 50)-(X + K1 - 50, Y + Height + 50 - 100), vbRed
Obj.Line (X + K1, Y + Height + 50)-(X + K1 + 50, Y + Height + 50 - 100), vbRed

Obj.CurrentY = Y + 30
Obj.CurrentX = Obj.CurrentX + 20 - Obj.TextWidth(Str(N)) / 2
If K1 = 0 Then Obj.Print "0 " + Mid(sUnit, 1, 2) + "." Else Obj.Print N

Obj.Line (X + K1 + TwpPerUnit / 2, Y)-(X + K1 + TwpPerUnit / 2, Y + Height)
    
    For K2 = TmpK1 + SmallSteps To K1 - SmallSteps Step SmallSteps
     bSwitch = Not bSwitch
    
     If bSwitch Then
     Obj.Line (X + K2, Y + Height - 200)-(X + K2, Y + Height)
     Else
     Obj.Line (X + K2, Y + Height - 100)-(X + K2, Y + Height)
     End If
     Next K2
    TmpK1 = K1

Next K1
'restore the things
   SetTextColor Obj.hDC, BuFore
Obj.DrawStyle = BuDrawStyle: Obj.DrawWidth = BuDrawWidth: Obj.DrawMode = BuDrawMode
Obj.ScaleMode = BUScaleMode: Obj.FontSize = BuFontS
End Sub
'This draws vertical ruler
'This method can be removed, if ur not using it
Public Sub DrawRulerV(Obj As Object, X As Long, Y As Long, Length As Long, sUnit As String)
Dim Height As Long, Devide As Long
Dim SmallSteps  As Long, bSwitch As Boolean
Dim h1 As Long, TwpPerUnit As Long
Dim TmpK1 As Long
Dim K1 As Long, K2 As Long, N As Long

Dim BuDrawStyle As Long, BuDrawWidth As Long, BuDrawMode As Long, BUScaleMode As Long
Dim BuFontS As Long
Dim BuFore As Long
If Not ValidObj(Obj) Then Exit Sub
'Back Up
BuDrawStyle = Obj.DrawStyle: BuDrawWidth = Obj.DrawWidth
BuDrawMode = Obj.DrawMode: BUScaleMode = Obj.ScaleMode
BuFontS = Obj.FontSize
BuFore = Obj.ForeColor
'Set New
Obj.DrawStyle = vbSolid: Obj.DrawWidth = 1: Obj.DrawMode = vbCopyPen
Obj.ScaleMode = vbTwips: Obj.FontSize = 8: Obj.FontBold = False: Obj.FontItalic = False
Obj.FontStrikethru = False
Obj.FontUnderline = False

SetTextColor Obj.hDC, vbBlack
X = X * Screen.TwipsPerPixelX: Y = Y * Screen.TwipsPerPixelY
Height = 400
'Length = Obj.Height

If UCase(sUnit) = "INCH" Then
TwpPerUnit = 1440
Devide = 20
Else
'CM
TwpPerUnit = 567
Devide = 6
End If

SmallSteps = TwpPerUnit / Devide
X = X - (Height + 50)


Obj.Line (X, Y)-(X, Y + Length)
Obj.Line (X + Height, Y)-(X + Height, Y + Length)
N = -1
For K1 = 0 To Length Step TwpPerUnit
bSwitch = True: N = N + 1
Obj.Line (X, Y + K1)-(X + Height + 50, Y + K1), vbRed
'Aero Heads
Obj.Line (X + Height + 50, Y + K1)-(X + Height - 50, Y + K1 - 50), vbRed
Obj.Line (X + Height + 50, Y + K1)-(X + Height - 50, Y + K1 + 50), vbRed

Obj.CurrentX = X - 30 + Obj.TextWidth(Str(N)) / 2
If K1 = 0 Then Obj.Print "0 " + Mid(sUnit, 1, 2) + "." Else Obj.Print N

Obj.Line (X, Y + K1 + TwpPerUnit / 2)-(X + Height, Y + K1 + TwpPerUnit / 2)
    
    For K2 = TmpK1 + SmallSteps To K1 - SmallSteps Step SmallSteps
     bSwitch = Not bSwitch
    
     If bSwitch Then
     Obj.Line (X + Height - 200, Y + K2)-(X + Height, Y + K2)
     Else
     Obj.Line (X + Height - 100, Y + K2)-(X + Height, Y + K2)
     End If
     Next K2
    TmpK1 = K1

Next K1
Obj.DrawStyle = BuDrawStyle: Obj.DrawWidth = BuDrawWidth: Obj.DrawMode = BuDrawMode
Obj.ScaleMode = BUScaleMode: Obj.FontSize = BuFontS
SetTextColor Obj.hDC, BuFore
End Sub
'Code Generator
'This method can be removed if ur not using it
Public Sub ColSetupCode()

Dim C As Long, R As Long
Dim cd As String

cd = cd + "'You can place this code in Form's load event" + vbCrLf + vbCrLf
cd = cd + "'MSFlex Grid Control does not require these code but MSHFlex does" + vbCrLf + vbCrLf
cd = cd + "With " + objFlex.Name + vbCrLf

For C = 0 To objFlex.Cols - 1
R = objFlex.ColWidth(C)
If R < 1 Then cd = cd + "'You must set positive value to follwing line" + vbCrLf
cd = cd + ".ColWidth(" + Str(C) + ")  =" + Str(R) + vbCrLf
Next C
cd = cd + "End With"
Clipboard.Clear ' Clear Clipboard.
Clipboard.SetText cd
MsgBox "Column setup code has been copied to cilipboard"

End Sub

'Calculation of the height of the output
'U should set RowsFrom and RowsTo B4 calling this function
'Returns in Pixels
Public Function GetHeight(Obj As Object) As Long
Dim Mul As Single, lRow As Long, lTotHeight As Long
Dim lNum_of_Rows As Long, TmpHt As Long, BUScaleMode As Long 'BU=BackUp
Mul = 1
BUScaleMode = Obj.ScaleMode
Obj.ScaleMode = vbPixels
If TypeName(Obj) = "Printer" Then Mul = Screen.TwipsPerPixelY / Printer.TwipsPerPixelY
lNum_of_Rows = lRowsTo - lRowsFrom

For lRow = lRowsFrom To lRowsTo
        TmpHt = objFlex.RowHeight(lRow)
            If TmpHt = 0 Then lNum_of_Rows = lNum_of_Rows - 1
        lTotHeight = lTotHeight + TmpHt
Next

lTotHeight = Obj.ScaleY(lTotHeight, vbTwips, vbPixels) + (objFlex.GridLineWidth * Mul)
lTotHeight = lTotHeight + (lNum_of_Rows * HSpace * Mul)

If DrawBoarder Then
        lTotHeight = lTotHeight + (2 * BoarderDistance * Mul)
End If

GetHeight = lTotHeight
Obj.ScaleMode = BUScaleMode

End Function

'Calculation of the Width of the output
'returns in Pixels
Public Function GetWidth(Obj As Object) As Long
Dim lCol As Long, Mul As Single, lCols As Long
Dim lTotWidth As Long, lNum_of_Cols As Long
Dim TmpHt As Long, BUScaleMode As Long
Mul = 1
BUScaleMode = Obj.ScaleMode
Obj.ScaleMode = vbPixels
If TypeName(Obj) = "Printer" Then Mul = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
lNum_of_Cols = objFlex.Cols

For lCols = 0 To lNum_of_Cols - 1
TmpHt = objFlex.ColWidth(lCols)
            If TmpHt = 0 Then lNum_of_Cols = lNum_of_Cols - 1
        lTotWidth = lTotWidth + TmpHt
Next

lTotWidth = Obj.ScaleX(lTotWidth, vbTwips, vbPixels) + (objFlex.GridLineWidth * Mul)
lTotWidth = lTotWidth + (lNum_of_Cols * VSpace * Mul)

If DrawBoarder Then
        lTotWidth = lTotWidth + (2 * BoarderDistance * Mul)
End If

GetWidth = lTotWidth
Obj.ScaleMode = BUScaleMode

End Function
'----------------------------
' AlignRect
' Aligns SmallRect within BigRect as per given Alignment
'----------------------------
'0   3   6            0=Left Top
'1   4   7            4=Center Center
'2   5   8            8=Right Bottom
Private Sub AlignRECT(BigRECT As RECT, SmallRECT As RECT, ByVal Alignment As Integer)
Dim SmallWidth As Long, SmallHeight As Long
Dim BigWidth As Long, BigHeight As Long

SmallWidth = SmallRECT.Right - SmallRECT.Left
SmallHeight = SmallRECT.Bottom - SmallRECT.Top

BigWidth = BigRECT.Right - BigRECT.Left
BigHeight = BigRECT.Bottom - BigRECT.Top

If SmallHeight > BigHeight Then
        SmallRECT.Top = BigRECT.Top: SmallRECT.Bottom = BigRECT.Bottom
        SmallHeight = SmallRECT.Bottom - SmallRECT.Top
End If

If SmallWidth > BigWidth Then
        SmallRECT.Left = BigRECT.Left: SmallRECT.Right = BigRECT.Right
        SmallWidth = SmallRECT.Right - SmallRECT.Left
End If

With SmallRECT
    Select Case Alignment
                Case 0  'Left top
                        .Left = BigRECT.Left
                        .Top = BigRECT.Top
                        .Right = .Left + SmallWidth
                        .Bottom = .Top + SmallHeight
                Case 1  'Left Center
                        .Left = BigRECT.Left
                        .Top = ((BigHeight - SmallHeight) / 2) + BigRECT.Top
                        .Right = .Left + SmallWidth
                        .Bottom = .Top + SmallHeight
                Case 2  'Left bottom
                        .Top = BigRECT.Bottom - SmallHeight
                        .Left = BigRECT.Left
                        .Right = .Left + SmallWidth
                        .Bottom = .Top + SmallHeight
                Case 3  'Center Top
                        .Left = ((BigWidth - SmallWidth) / 2) + BigRECT.Left
                        .Top = BigRECT.Top
                        .Right = .Left + SmallWidth
                        .Bottom = .Top + SmallHeight
                Case 4, 9   'Center Center
                        .Top = ((BigHeight - SmallHeight) / 2) + BigRECT.Top
                        .Left = ((BigWidth - SmallWidth) / 2) + BigRECT.Left
                        .Right = .Left + SmallWidth
                        .Bottom = .Top + SmallHeight
                Case 5  'Center Bottom
                        .Top = BigRECT.Bottom - SmallHeight
                        .Left = ((BigWidth - SmallWidth) / 2) + BigRECT.Left
                        .Right = .Left + SmallWidth
                        .Bottom = .Top + SmallHeight
                Case 6  'Right Top
                        .Left = BigRECT.Right - SmallWidth
                        .Top = BigRECT.Top
                        .Right = .Left + SmallWidth
                        .Bottom = .Top + SmallHeight
                Case 7  'Right Center
                        .Left = BigRECT.Right - SmallWidth
                        .Top = ((BigHeight - SmallHeight) / 2) + BigRECT.Top
                        .Right = .Left + SmallWidth
                        .Bottom = .Top + SmallHeight
                Case 8  'Right Bottom
                        .Top = BigRECT.Bottom - SmallHeight
                        .Left = BigRECT.Right - SmallWidth
                        .Right = BigRECT.Right
                        .Bottom = BigRECT.Bottom
                Case Else
                        .Left = BigRECT.Left
                        .Top = BigRECT.Top
                        .Right = BigRECT.Right
                        .Bottom = BigRECT.Bottom
      End Select 'Alignment
End With 'SmallRect
End Sub
