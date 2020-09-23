VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Demo FlexPrinter Class"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   FillColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex 
      Height          =   2850
      Left            =   135
      TabIndex        =   59
      Top             =   405
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   5027
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FillStyle       =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Pictures"
      Height          =   240
      Left            =   8055
      TabIndex        =   58
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      Caption         =   "Picture"
      Height          =   1905
      Left            =   7830
      TabIndex        =   54
      Top             =   45
      Width           =   1590
      Begin VB.CommandButton Command10 
         Caption         =   "Delete"
         Height          =   375
         Left            =   135
         TabIndex        =   57
         Top             =   1035
         Width           =   1365
      End
      Begin VB.CommandButton Command9 
         Height          =   780
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "  RIGHT CLICK to Change LEFT CLICK to Apply     "
         Top             =   225
         Width           =   1050
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pict Alignments"
         Height          =   375
         Left            =   135
         TabIndex        =   55
         ToolTipText     =   "      changes cell picture alignments    "
         Top             =   1440
         Width           =   1365
      End
   End
   Begin VB.CheckBox Check4 
      Caption         =   "back color"
      Height          =   240
      Left            =   6885
      TabIndex        =   53
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   420
      Left            =   9540
      TabIndex        =   50
      Top             =   7290
      Width           =   1500
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Boarder"
      Height          =   240
      Left            =   8055
      TabIndex        =   42
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1230
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "WordWrap"
      Height          =   240
      Left            =   6885
      TabIndex        =   41
      Top             =   2520
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Text Alignments"
      Height          =   330
      Left            =   8010
      TabIndex        =   40
      ToolTipText     =   "   Changes cell text alignmets   "
      Top             =   2070
      Width           =   1365
   End
   Begin VB.PictureBox Picture3 
      Height          =   4650
      Left            =   90
      ScaleHeight     =   4590
      ScaleWidth      =   9135
      TabIndex        =   38
      Top             =   3240
      Width           =   9195
      Begin VB.PictureBox PIC 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   9045
         Left            =   135
         ScaleHeight     =   601
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   803
         TabIndex        =   39
         Top             =   135
         Width           =   12075
         Begin ComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   6
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrintGrid.frx":0000
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrintGrid.frx":0852
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrintGrid.frx":10A4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrintGrid.frx":18F6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrintGrid.frx":2148
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "PrintGrid.frx":299A
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Inch"
      Height          =   195
      Left            =   9360
      TabIndex        =   36
      Top             =   6075
      Width           =   870
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cm"
      Height          =   195
      Left            =   10305
      TabIndex        =   35
      Top             =   6075
      Value           =   -1  'True
      Width           =   825
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   9405
      TabIndex        =   34
      Text            =   "8000"
      Top             =   5220
      Width           =   1230
   End
   Begin VB.CheckBox Hor 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Horizontal Ruler"
      Height          =   285
      Left            =   9405
      TabIndex        =   32
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CheckBox Ver 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Verticle Ruler"
      Height          =   285
      Left            =   9405
      TabIndex        =   31
      Top             =   5535
      Width           =   2130
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rows"
      Height          =   825
      Left            =   9270
      TabIndex        =   26
      Top             =   4050
      Width           =   1770
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1035
         TabIndex        =   30
         ToolTipText     =   " No. of final  row "
         Top             =   450
         Width           =   645
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1035
         TabIndex        =   29
         ToolTipText     =   " No. of initial  row "
         Top             =   135
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "To"
         Height          =   240
         Left            =   45
         TabIndex        =   28
         Top             =   540
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "From"
         Height          =   195
         Left            =   45
         TabIndex        =   27
         Top             =   225
         Width           =   870
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grid "
      Height          =   240
      Left            =   6885
      TabIndex        =   25
      Top             =   2115
      Value           =   1  'Checked
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   10305
      TabIndex        =   23
      Top             =   3780
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   10305
      TabIndex        =   22
      ToolTipText     =   " Value to round the cornor of rectangle"
      Top             =   3465
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   10305
      TabIndex        =   21
      ToolTipText     =   " Value to round the cornor of rectangle"
      Top             =   3195
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   10305
      TabIndex        =   20
      ToolTipText     =   "Horizontle space between Rows"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   10305
      TabIndex        =   19
      ToolTipText     =   "Verticle Space between columns"
      Top             =   2610
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   10305
      TabIndex        =   18
      Top             =   2295
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   10305
      TabIndex        =   17
      Top             =   2025
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   10305
      TabIndex        =   16
      Top             =   1125
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   10305
      TabIndex        =   15
      Top             =   720
      Width           =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Boarder"
      Height          =   1905
      Left            =   9450
      TabIndex        =   10
      Top             =   45
      Width           =   1590
      Begin VB.PictureBox Picture2 
         Height          =   285
         Left            =   855
         ScaleHeight     =   225
         ScaleWidth      =   540
         TabIndex        =   24
         ToolTipText     =   "      Drag and drop color here      "
         Top             =   1485
         Width           =   600
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   855
         TabIndex        =   14
         Top             =   225
         Width           =   600
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Color"
         Height          =   240
         Left            =   90
         TabIndex        =   37
         Top             =   1575
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Width"
         Height          =   285
         Left            =   45
         TabIndex        =   13
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Style"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Distance"
         Height          =   195
         Left            =   45
         TabIndex        =   11
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      DragIcon        =   "PrintGrid.frx":326C
      Height          =   1545
      Left            =   6930
      Picture         =   "PrintGrid.frx":36AE
      ScaleHeight     =   1485
      ScaleWidth      =   765
      TabIndex        =   9
      ToolTipText     =   "  Click to change Text color Drag to change Boarder color   "
      Top             =   405
      Width           =   825
      Begin VB.OptionButton Option4 
         Caption         =   "Bk Col"
         Height          =   195
         Left            =   0
         TabIndex        =   52
         ToolTipText     =   "   Back Color   "
         Top             =   1260
         Width           =   1365
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Txt Col"
         Height          =   195
         Left            =   0
         TabIndex        =   51
         ToolTipText     =   "   Text Color   "
         Top             =   1035
         Value           =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Print"
      Height          =   420
      Left            =   9315
      TabIndex        =   8
      Top             =   6840
      Width           =   1500
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Preview/Refresh"
      Height          =   465
      Left            =   9495
      TabIndex        =   7
      Top             =   6345
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6345
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   45
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "StrikeThorugh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   285
      Left            =   4860
      TabIndex        =   5
      Top             =   45
      Width           =   1365
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Underline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   4
      Top             =   45
      Width           =   915
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Italics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3150
      TabIndex        =   3
      Top             =   45
      Width           =   825
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2340
      TabIndex        =   2
      Top             =   45
      Width           =   825
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FontSize -"
      Height          =   285
      Left            =   1395
      TabIndex        =   1
      Top             =   45
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FontSize +"
      Height          =   285
      Left            =   450
      TabIndex        =   0
      Top             =   45
      Width           =   960
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "H Space"
      Height          =   240
      Left            =   9180
      TabIndex        =   49
      Top             =   2970
      Width           =   1050
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "V Space"
      Height          =   240
      Left            =   9135
      TabIndex        =   48
      Top             =   2655
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Round X"
      Height          =   240
      Left            =   9135
      TabIndex        =   47
      Top             =   3285
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Round Y"
      Height          =   240
      Left            =   9090
      TabIndex        =   46
      Top             =   3555
      Width           =   1140
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Left"
      Height          =   240
      Left            =   9135
      TabIndex        =   45
      Top             =   2115
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Top"
      Height          =   240
      Left            =   9135
      TabIndex        =   44
      Top             =   2385
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Style"
      Height          =   240
      Left            =   9045
      TabIndex        =   43
      Top             =   3825
      Width           =   1185
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ruler Length"
      Height          =   240
      Left            =   9405
      TabIndex        =   33
      Top             =   4950
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'A sample to show use of FlexPrinter class
 
Dim CCC As Long

Dim M As FlexPrinter

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Sub Check1_Click()
Command12_Click
End Sub

Private Sub Check2_Click()
Command12_Click
End Sub

Private Sub Check3_Click()
Flex.WordWrap = Check3.Value

End Sub

Private Sub Check4_Click()
Command12_Click
End Sub

Private Sub Check5_Click()
Command12_Click
End Sub


Private Sub Combo1_Click()
Flex.CellFontName = Combo1.Text
Command12_Click
End Sub



Private Sub Command1_Click()
Static X As Integer
X = X + 1
Flex.CellAlignment = X
If X > 8 Then X = -1

End Sub







Private Sub Command10_Click()
Set Flex.CellPicture = Nothing
End Sub

Private Sub Command12_Click()
PIC.Cls

Dim MEs As String
If Option1.Value = True Then MEs = "CM" Else MEs = "INCH"
With M
.RowsFrom = Val(Text2)
.RowsTo = Val(Text3)

'to center
'.PosTop = (PIC.ScaleHeight - .GetHeight(PIC)) / 2 '
'.PosLeft = (PIC.ScaleWidth - .GetWidth(PIC)) / 2 '

.PosTop = Val(Text1(4).Text)
.PosLeft = Val(Text1(3).Text)

.HSpace = Val(Text1(6).Text)
.VSpace = Val(Text1(5).Text)

.RoundCorX = Val(Text1(7).Text)
.RoundCorY = Val(Text1(8).Text)

.GridPenStyle = Val(Text1(9).Text)
.GridPrint = Check2.Value
.DrawBoarder = Check1.Value
.BoarderColor = Picture2.BackColor
.BoarderStyle = Val(Text1(1).Text)
.BoarderWidth = Val(Text1(2).Text)
.BoarderDistance = Val(Text1(0).Text)
.RowsFrom = Val(Text2.Text)
.RowsTo = Val(Text3.Text)
.EnableCellBackColor = Check4.Value
.PicturePrint = Check5.Value

.PrintOut PIC

If Ver.Value = vbChecked Then
.DrawRulerV PIC, Val(Text1(3)), Val(Text1(4)), Val(Text4.Text), MEs
End If
If Hor.Value = vbChecked Then
.DrawRulerH PIC, Val(Text1(3)), Val(Text1(4)), Val(Text4.Text), MEs

End If

End With

'these two lines to check height and width of picture in picturebox
'PIC.Line (M.PosLeft, M.PosTop - 5)-(M.GetWidth(PIC) + M.PosLeft, M.PosTop - 5)
'PIC.Line (M.PosLeft - 5, M.PosTop)-(M.PosLeft - 5, M.GetHeight(PIC) + M.PosTop)
Flex.SetFocus
End Sub



Private Sub Command13_Click()
Static X1 As Integer
X1 = X1 + 1
Flex.CellPictureAlignment = X1
If X1 > 8 Then X1 = -1

End Sub

Private Sub Command14_Click()
Printer.PaperSize = 9
Printer.Orientation = 1
Printer.ScaleMode = 3
M.RowsFrom = Val(Text2)
M.RowsTo = Val(Text3)

M.PosTop = (Printer.ScaleHeight - M.GetHeight(Printer)) / 2
M.PosLeft = (Printer.ScaleWidth - M.GetWidth(Printer)) / 2

M.PrintOut Printer

Printer.EndDoc
End Sub



Private Sub Command2_Click()

Flex.CellFontSize = Flex.CellFontSize + 2

End Sub





Private Sub Command3_Click()
MsgBox "An Ordinary MSHFlex Grid Printer" + vbCrLf + "Opal Raj Ghimire" + vbCrLf + "Kathmandu, Nepal" + vbCrLf + vbCrLf + "buna48@hotmail.com", , "Flex Printer"
'M.ColSetupCode


End Sub

Private Sub Command4_Click()
Flex.CellFontSize = Flex.CellFontSize - 2
If Flex.CellFontSize < 8 Then Flex.CellFontSize = 8

End Sub

Private Sub Command5_Click()
Flex.CellFontBold = Not Flex.CellFontBold

End Sub

Private Sub Command6_Click()
Flex.CellFontItalic = Not Flex.CellFontItalic

End Sub

Private Sub Command7_Click()
Flex.CellFontUnderline = Not Flex.CellFontUnderline

End Sub



Private Sub Command8_Click()

Flex.CellFontStrikeThrough = Not Flex.CellFontStrikeThrough

End Sub



Private Sub Command9_Click()
Set Flex.CellPicture = Command9.Picture
End Sub

Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


Static R As Integer
If Button = 2 Then
R = R + 1
Command9.Picture = ImageList1.ListImages(R).Picture
If R = ImageList1.ListImages.Count Then R = 0

End If
End Sub


Private Sub Form_Load()
Dim I As Integer
Dim CC As Control
For Each CC In Form1

If TypeName(CC) <> "ImageList" Then
CC.FontName = "MS Sans Serif"
CC.FontSize = 8
End If
Next

With Flex

.ColWidth(0) = 300
.ColWidth(1) = 1920
.ColWidth(2) = 1920
.ColWidth(3) = 1920
.ColWidth(4) = 1000
.TextMatrix(0, 0) = "Sr"
.TextMatrix(0, 1) = "Things"
.TextMatrix(0, 2) = "Time"
.TextMatrix(0, 3) = "Date"
.AddItem "3" + Chr(9) + "cisaB lausiV" + Chr(9) + "2002" + Chr(9) + "U        Me" + Chr(9) + "Hello World!", 1
.AddItem "2" + Chr(9) + "Dumped again!" + Chr(9) + "03:20 PM" + Chr(9) + "But" + Chr(9) + "Explore the site" + Chr(9), 1
.AddItem "1" + Chr(9) + "http://geocities.com/ opalraj/ vb" + Chr(9) + "$$$" + Chr(9) + "I       U" + Chr(9) + "e", 1
.RowHeight(0) = 280
.RowHeight(1) = 1200
.RowHeight(2) = 600
.RowHeight(3) = 600
.RowHeight(4) = 600
End With
Check3.Value = IIf(Flex.WordWrap, vbChecked, vbUnchecked)
 For I = 0 To Printer.FontCount - 1  ' Determine number of fonts.
     Combo1.AddItem Printer.Fonts(I)  ' Put each font into list box.
    Next I
'set up
Text1(0) = "5"
Text1(1) = "0"
Text1(2) = "1"
Text1(3) = "50"
Text1(4) = "50"
Text1(5) = "0"
Text1(6) = "0"
Text1(7) = "10"
Text1(8) = "10"
Text1(9) = "0"
'Text1(10) = "0"
'Text1(11) = "1"

Set M = New FlexPrinter
Set M.FlexName = Flex

'Check1.Value = vbChecked
Check3.Value = vbChecked
Picture2.BackColor = vbBlack
 Text2 = 0: M.RowsFrom = Val(Text2)
 Text3 = 4: M.RowsTo = Val(Text3)
Flex.BackColorFixed = RGB(225, 225, 225) '12632256
Command9.Picture = ImageList1.ListImages(1).Picture

'***********
SetUpFlex
'***********

End Sub


Private Sub Flex_Click()
Combo1 = Flex.CellFontName

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set M = Nothing
End Sub



Private Sub Hor_Click()
Command12_Click
End Sub

Private Sub Label14_Click()
Flex.GridLineWidth = 4
End Sub

Private Sub Option1_Click()
Command12_Click
End Sub

Private Sub Option2_Click()
Command12_Click
End Sub



Private Sub PIC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PIC.CurrentX = 50: PIC.CurrentY = 5
PIC.Print "Please click on MSHFlex Grid Control"
End Sub

Private Sub PIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        ReleaseCapture
        Call SendMessage(PIC.hwnd, &H112, &HF012, 0)
End If

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CCC = Picture1.Point(X, Y)
If Option3.Value = True Then
Flex.CellForeColor = CCC
Else
Flex.CellBackColor = CCC
End If

If Button = 1 Then Picture1.Drag

End Sub

Private Sub Picture2_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Name = "Picture1" Then
Picture2.BackColor = CCC
Command12_Click
End If
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Command12_Click
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
If IsNumeric(Text1(Index).Text) = False Then
MsgBox "Numeric Expected"
Cancel = True
End If

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command12_Click

End Sub

Private Sub Text2_Validate(Cancel As Boolean)
If IsNumeric(Text2.Text) = False Then
MsgBox "Numeric Expected"
Cancel = True
End If

End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command12_Click
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If IsNumeric(Text3.Text) = False Then
MsgBox "Numeric Expected"
Cancel = True
End If

End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command12_Click
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If IsNumeric(Text4.Text) = False Then
MsgBox "Numeric Expected"
Cancel = True
End If

End Sub

Private Sub Ver_Click()
Command12_Click
End Sub

Private Sub SetUpFlex()
With Flex
    .Redraw = False
    .Row = 1: .Col = 1
    .CellFontItalic = True
    .CellAlignment = 8
    .CellFontSize = 10
    .CellForeColor = vbBlue
     .CellBackColor = vbYellow
     Set .CellPicture = ImageList1.ListImages(3).Picture
    
    .Row = 3: .Col = 2
    .CellFontBold = True
    .CellBackColor = vbWhite
    .CellForeColor = vbBlue
    .CellFontSize = 24
    .CellAlignment = 4
    .CellFontName = "Times New Roman"
    
    .Row = 1: .Col = 3
    .CellFontBold = True
    .CellBackColor = vbWhite
    .CellForeColor = vbBlue
    .CellFontSize = 20
    .CellFontName = "Times New Roman"
    .CellAlignment = 4
    .CellPictureAlignment = 4
    Set .CellPicture = ImageList1.ListImages(2).Picture
    
     .Row = 1: .Col = 4
     .CellFontName = "Times New Roman"
    .CellAlignment = 4
    .CellFontSize = 48
    .CellForeColor = vbRed
    .CellBackColor = 10092288
    .Row = 2: .Col = 1
     .CellAlignment = 1
    .CellPictureAlignment = 7
     Set .CellPicture = ImageList1.ListImages(4).Picture
      .Row = 1: .Col = 2
      Set .CellPicture = ImageList1.ListImages(1).Picture
     .CellAlignment = 8
     .CellFontSize = 18
     .CellPictureAlignment = 4
     .Row = 2: .Col = 2
    .CellBackColor = 10092288
    .Text = Now()
    
     .Row = 2: .Col = 3
    'but
    .CellBackColor = 16776960
    .CellAlignment = 7
    .CellFontSize = 24
    .CellFontBold = True
    .CellForeColor = vbWhite
    
    .Row = 2: .Col = 4
    .CellBackColor = 65280
    .CellFontName = "Times New Romam"
    
    .Row = 3: .Col = 3
    'u  me
      .CellFontSize = 18
     Set .CellPicture = ImageList1.ListImages(1).Picture
      .CellAlignment = 4
    .CellPictureAlignment = 4
    .CellFontBold = True
      .Row = 3: .Col = 1
       .CellBackColor = 65280
    .CellFontName = "Times New Roman"
    .CellFontBold = True
    .CellAlignment = 7
    .CellFontSize = 14
    .CellForeColor = vbRed
      .Row = 4: .Col = 2
    .Text = "FlexPrinter2" + vbCrLf + "with pictures"
    .CellFontBold = True: .CellFontItalic = True
    .CellBackColor = 10092543
    .CellForeColor = 16711935
      .Row = 4: .Col = 4
     Set .CellPicture = ImageList1.ListImages(5).Picture
     .CellPictureAlignment = 4
     .CellBackColor = 16777113
     .Row = 4: .Col = 1
     Set .CellPicture = ImageList1.ListImages(6).Picture
     .CellPictureAlignment = 9

    
    
    .Redraw = True
End With
End Sub
