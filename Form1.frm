VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text Formatting API"
   ClientHeight    =   4950
   ClientLeft      =   5430
   ClientTop       =   -45
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMargin 
      Height          =   315
      Left            =   5700
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   630
      Width           =   765
   End
   Begin VB.TextBox txtSource 
      Height          =   3465
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1260
      Width           =   3135
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3435
      Left            =   3570
      ScaleHeight     =   3435
      ScaleWidth      =   3135
      TabIndex        =   7
      Top             =   1260
      Width           =   3135
   End
   Begin VB.Frame fraAlignType 
      Caption         =   "Alignmente Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   180
         ScaleHeight     =   345
         ScaleWidth      =   4575
         TabIndex        =   2
         Top             =   240
         Width           =   4575
         Begin VB.OptionButton optAlignType 
            Caption         =   "Left"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   60
            Width           =   690
         End
         Begin VB.OptionButton optAlignType 
            Caption         =   "Right"
            Height          =   195
            Index           =   1
            Left            =   1185
            TabIndex        =   5
            Top             =   60
            Width           =   765
         End
         Begin VB.OptionButton optAlignType 
            Caption         =   "Center"
            Height          =   195
            Index           =   2
            Left            =   2325
            TabIndex        =   4
            Top             =   60
            Width           =   840
         End
         Begin VB.OptionButton optAlignType 
            Caption         =   "Justified"
            Height          =   195
            Index           =   3
            Left            =   3540
            TabIndex        =   3
            Top             =   60
            Value           =   -1  'True
            Width           =   840
         End
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Margin:"
      Height          =   195
      Index           =   1
      Left            =   5130
      TabIndex        =   10
      Top             =   660
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "How to align text on various formatting. Choose the Alignment type: "
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   4980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' DateTime  : 30/03/2004 01.08
' Author    : Giorgio Brausi
' Project   : TextJustify
' Purpose   : Print formatted text with various style
'---------------------------------------------------------------------------------------
'           This sample show how to change the text formatting from
'           a TextBox to a PictureBox.
'           This my code was originally implemented with VB 3.0,
'           and now i review the code for VB 6.0.
'           The Justify() funstion derive from the title:
'           "Programming Windows 3.1" - 3rd Edition (1997)
'           by Charles Petzold  (1st edition 1992)
'           Copyright © Jackson Libri (italian edition)
'           Chapter 14 - Text and Font: Formatting Text section
'           from pag. 638 to 647
'           The code is well commented, so you can understan all!
'---------------------------------------------------------------------------------------
Option Explicit



Private Sub Form_Load()
    Dim num As Integer, s As String, i As Integer
    num = FreeFile
    
    '/ load orginal sample text file
    Open App.Path & "\JUSTIFY.ASC" For Binary As #num
    s = Space(LOF(1))
    Get #1, , s
    txtSource.Text = s
    Close
    
    ' repare the margin choice
    ' he margin is applied to the Left, Top and
    ' Right borders (not to the bottom)
    With cboMargin
        .Text = ""
        For i = 0 To 5
            .AddItem i * 5
        Next i
        .ListIndex = 2  '/ set margin = 10
    End With
    
    ' set justified alignment
    optAlignType_Click 3
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure   : JustifyText (NOT USED!)
' DateTime    : 30/03/2004 01.44
' Author      : Giorgio Brausi
' Purpose     : This is a old routine which print a single text line.
' Descritpion : Just for educational purpose.
'---------------------------------------------------------------------------------------
'
Public Sub JustifyText(ByVal lpString As String, oPicDest As Object)
'Dim lRet As Long, lpSize As Size
'Dim nBreakExtra As Long, nBreakCount As Long, nCount As Integer
'Dim lWidth As Long, arString() As String
'Dim hdc As Long
'
'    On Error Resume Next
''    If TypeOf oPicDest Is PictureBox Then
''        oPicDest.Cls
''    End If
'    hdc = oPicDest.hdc
'
'    If Err.Number <> 0 Then
'        MsgBox "The object named " & picDest.Name & " has not a HDC property!" & vbCrLf & "Action aborted.", vbCritical
'        Exit Sub
'    End If
'
'    nCount = Len(lpString)         ' string len
'    lWidth = oPicDest.ScaleWidth   ' width of output device
'
'    ' reset the error index
'    lRet = SetTextJustification(hdc, 0, 0)
'
'
'    ' compute the string width and height
'    ' lpSize structure (cx and cy) will contain this value
'    lRet = GetTextExtentPoint32(hdc, lpString, nCount, lpSize)
'
'    If lpSize.cx > lWidth Then
'        MsgBox "This string is too long! Can't be printed."
'        Exit Sub
'    End If
'
'
'    ' calc spaces number
'    arString = Split(lpString, " ")
'    nBreakCount = UBound(arString)
'
'    ' Compute the need spaces to fill the width
'    nBreakExtra = lWidth - lpSize.cx
'
'    ' set justification and print string
'    lRet = SetTextJustification(hdc, nBreakExtra, nBreakCount)
'    TextOut hdc, 0, 50, lpString, nCount

End Sub

Private Sub Form_Resize()
    ' prevent error
    If Me.WindowState = vbMinimized Then Exit Sub
    
    ' resize textbox and picturebox in according to windows size
    txtSource.Move 180, txtSource.Top, Width / 2 - 180, Height - txtSource.Top - 540
    picDest.Move Width / 2 + 180, picDest.Top, Width / 2 - 360, Height - picDest.Top - 540

    Call optAlignType_Click(nAlign)
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure   : optAlignType_Click
' DateTime    : 30/03/2004 01.36
' Author      : Giorgio Brausi
' Purpose     : Set and draw the formatted text
' Descritpion : call the Justify() routine
'---------------------------------------------------------------------------------------
'
Private Sub optAlignType_Click(Index As Integer)
    nAlign = Index
    
    picDest.ScaleMode = vbPixels    '!!! API need Pixels
    
    '/ set rectangle for print
    Dim rc As RECT
    With rc
        .Left = 0
        .Top = 0
        .Right = picDest.ScaleWidth
        .Bottom = picDest.ScaleHeight
    End With
    picDest.Cls
    
    Call Justify(picDest.hdc, txtSource, rc, nAlign, cboMargin.Text)
    
End Sub


