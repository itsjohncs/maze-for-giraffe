VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   744
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Make"
      Height          =   495
      Left            =   7680
      TabIndex        =   4
      Top             =   7560
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   8040
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   3
      Top             =   5520
      Width           =   1980
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   2535
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2880
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   120
      ScaleHeight     =   7170
      ScaleWidth      =   7170
      TabIndex        =   1
      Top             =   120
      Width           =   7200
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   2535
      Left            =   7440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemoryWord Lib "kernel32" Alias "RtlMoveMemory" (Destination As Integer, ByVal Source As Long, ByVal Length As Long)
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long

Dim hBuffer As Long, hSheet As Long, tSheet As IPictureDisp, MapSize As Integer

Private Sub TileBlt(ByVal DestDC As Long, ByVal SrcDC As Long, ByVal Tile As Integer, ByVal X As Integer, ByVal Y As Integer, Optional Transparent As Long, Optional ByVal Width As Integer = 32, Optional ByVal Height As Integer = 32)
If Transparent Then
    TransparentBlt DestDC, _
                   X, _
                   Y, _
                   Width, _
                   Height, _
                   SrcDC, _
                   (Tile Mod (Picture2.ScaleWidth / Width)) * Width, _
                   (Height * Int(Tile / (Picture2.ScaleWidth / Width))), _
                   Width, _
                   Height, _
                   Transparent
Else
    BitBlt DestDC, _
           X, _
           Y, _
           Width, _
           Height, _
           SrcDC, _
           (Tile Mod (Picture2.ScaleWidth / Width)) * Width, _
           (Height * Int(Tile / (Picture2.ScaleWidth / Width))), _
           vbSrcCopy
End If
End Sub

Public Function LOWORD(dw As Long) As Integer
CopyMemoryWord LOWORD, VarPtr(dw), 2
End Function

Public Function HIWORD(dw As Long) As Integer
CopyMemoryWord HIWORD, VarPtr(dw) + 2, 2
End Function

Public Function MAKELONG(ByVal iLoWord As Integer, ByVal iHiWord As Integer) As Long
MAKELONG = (iHiWord * &H10000) + (iLoWord And &HFFFF&)
End Function
Private Sub Command1_Click()
Dim iCount As Integer, sTemp As String
Dim aData() As String, aData2() As String
aData = Split(Text1, ".")
'sTemp = ""
aData2 = Split(Text2, ".")
'For iCount = 0 To UBound(aData)
'    sTemp = sTemp & Chr(Val(aData(iCount)) + 1) & Chr(Val(aData2(iCount)) + 1)
'Next iCount
Kill App.Path & "\Map.txt"
Open App.Path & "\Map.txt" For Random As #1 Len = 1
    For iCount = 0 To UBound(aData) * 2 + 1
        If ((iCount + 1) / 2 <> Int((iCount + 1) / 2)) Then
            Put #1, iCount + 1, CByte(Val(aData(Int((iCount) / 2))))
        Else
            Put #1, iCount + 1, CByte(Val(aData2(Int((iCount) / 2))))
        End If
    Next iCount
Close #1
End Sub

Private Sub Form_Load()
Dim sFile As String, b8File As Byte, iCount As Integer
MapSize = 15
Open App.Path & "\Map.txt" For Random As #1 Len = 1
    ReDim tTile(LOF(1) / 2)
    For iCount = 1 To LOF(1)
        Get #1, iCount, b8File
        If (iCount / 2 <> Int(iCount / 2)) Then
            Text1 = Text1 & b8File & "."
        Else
            Text2 = Text2 & b8File & "."
        End If
    Next iCount
Close #1
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1 = Text1 & "." & Int(X / 32) + Int(Y / 32) * 4
End Sub

Private Sub Text1_Change()
On Error GoTo Leave
Dim iCountX As Integer, iCountY As Integer, iCount As Integer, aHold() As String
Picture1.Cls
aHold = Split(Text1, ".")
For iCountY = 0 To Picture1.Height - 32 Step 32
    For iCountX = 0 To Picture1.Width - 32 Step 32
        TileBlt Picture1.hdc, Picture2.hdc, _
            Val(aHold(iCount)), iCountX, iCountY
        iCount = iCount + 1
    Next iCountX
Next iCountY
Leave:
End Sub
