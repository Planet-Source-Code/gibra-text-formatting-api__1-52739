Attribute VB_Name = "modJustify"
'---------------------------------------------------------------------------------------
' Module    : modJustify
' DateTime  : 30/03/2004 01.19
' Author    : Giorgio Brausi
' Project   : TextJustify
' Purpose   : This module provide the Justify routine used by Form1.
' Info      : Print formatted text with various style
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

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Type SIZE
    cx As Long
    cy As Long
End Type

Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextJustification Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long

Public nAlign As Integer    ' see cboMargin_Click on Form1
Const IDM_LEFT = 0
Const IDM_RIGHT = 1
Const IDM_CENTER = 2
Const IDM_JUST = 3


'---------------------------------------------------------------------------------------
' Procedure   : Justify
' DateTime    : 30/03/2004 02.06
' Author      : Giorgio Brausi
' Purpose     : Print a formatted text with different alignments
' Descritpion : need other explanations? I don't think...
'---------------------------------------------------------------------------------------
'
Public Sub Justify(ByVal hdc As Long, _
    ByVal lpText As String, _
    ByRef pRc As RECT, _
    ByVal nAlign As Integer, _
    ByVal iMargin As Integer)
     
    Dim dwExtent As Long
    Dim lpSize As SIZE
    Dim xStart As Integer, yStart As Integer
    Dim nBreakCount As Integer
    Dim i As Integer
    Dim w As Long               ' rectangle width
    Dim lRet As Long            ' functions return value
    Dim arString() As String    ' split array
    Dim nCount As Integer
    Dim nBreakExtra As Integer
    i = 1                       ' start from 1st char
    
    w = Int(pRc.Right - iMargin * 2) ' output width
    yStart = pRc.Top + iMargin  ' set the Y start position
    xStart = pRc.Left           ' set the X start position
    lpText = Trim(lpText)       ' remove leading blanks
    
    Dim iPos As Integer, iStart As Integer
    Dim sTmp As String
    Dim bFinish As Boolean
                
    iStart = 1
    Do
        nBreakCount = 0 ' for each line
        Do
           iStart = iPos
           iPos = InStr(iStart + 1, lpText, " ")    ' start to find space one by one
            If iPos Then
                ' get the incremental string
                sTmp = Mid(lpText, 1, iPos)
            Else
                ' there is no other spaces, this is the last line
                ' ---------------------------------------------------
                ' *** I dont' undestand because the below code
                'sTmp = lpText
                ' *** don't work correctly. The rest of space string
                ' *** is filled by mistake chars!
                ' This is the workaround: I have add a blanks spaces!
                ' ---------------------------------------------------
                sTmp = lpText + Space(w - Len(lpText))
                
                ' notify this is the last line
                bFinish = True
                Exit Do
            End If
            
            nCount = Len(sTmp)                      ' length of string
            lRet = SetTextJustification(hdc, 0, 0)  ' reset error index
            
            ' find the string extent
            lRet = GetTextExtentPoint32(hdc, sTmp, nCount, lpSize)
            
        Loop While lpSize.cx < w
        
        ' is string too long? Can't be printed!
        If iStart = 1 Then
            MsgBox "This string is too long!"
            Exit Do
        End If
        
        ' At this point: lpSize.cx is > w, therefore need:
        ' 1. return to previous space and retrieve the string
        ' 2. compute compute again the string extent
        If iPos <> 0 Then
            iPos = InStr(1, StrReverse(Trim(sTmp)), " ")
            sTmp = Left(Trim(sTmp), Len(Trim(sTmp)) - iPos)
            nCount = Len(sTmp)
            lRet = SetTextJustification(hdc, 0, 0)  '/ reset error index
            '/ get the word extent
            lRet = GetTextExtentPoint32(hdc, Trim(sTmp), nCount, lpSize)
        End If
                
        If Not bFinish Then
            arString = Split(sTmp, " ")             ' find how many spaces
            nBreakCount = UBound(arString)
            nBreakExtra = w - lpSize.cx             ' compute width difference
        End If
        
        Select Case (nAlign)                    ' use alignment for xStart
            Case IDM_LEFT
                xStart = pRc.Left
            Case IDM_RIGHT
                xStart = w - lpSize.cx           ' LOWORD(dwExtent)
            Case IDM_CENTER
                xStart = (w - lpSize.cx) / 2     ' LOWORD(dwExtent)) / 2
            Case IDM_JUST
                '/ don't need to justify the last line
                If Not bFinish Then
                    Call SetTextJustification(hdc, nBreakExtra, nBreakCount)
                End If
                xStart = pRc.Left
        End Select
        xStart = xStart + iMargin
        
        ' print string
        
        TextOut hdc, xStart, yStart, sTmp, nCount
        
        yStart = yStart + lpSize.cy             ' increment row to print next line
        
        ' restart formatting for the next line
        iPos = iStart
        lpText = Trim(Mid(lpText, Len(sTmp) + 1))
    
    Loop While Len(lpText) > 0


End Sub

