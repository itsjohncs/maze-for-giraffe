Attribute VB_Name = "Drawing"
Option Explicit

Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Public hBuffer As Long, hBitMap As Long, hSheet As Long, tSheet As IPictureDisp

Public Sub TileBlt(ByVal DestDC As Long, ByVal SrcDC As Long, ByVal Tile As Integer, ByVal X As Integer, ByVal Y As Integer, Optional Transparent As Long, Optional ByVal Width As Integer = 32, Optional ByVal Height As Integer = 32)
If Transparent Then
    TransparentBlt DestDC, _
                   X, _
                   Y, _
                   Width, _
                   Height, _
                   SrcDC, _
                   (Tile Mod ((4 * 32) / Width)) * Width, _
                   (Height * Int(Tile / ((4 * 32) / Width))), _
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
           (Tile Mod ((4 * 32) / Width)) * Width, _
           (Height * Int(Tile / ((4 * 32) / Width))), _
           vbSrcCopy
End If
End Sub

Public Function InitDrawing() As Boolean
hBuffer = CreateCompatibleDC(GetDC(0)) 'Create the buffer that we will draw into (Gets rid of annoying flickering)
If (hBuffer = 0) Then Exit Function
hSheet = CreateCompatibleDC(GetDC(0)) 'Holds the sprite sheet
If (hSheet = 0) Then Exit Function
hBitMap = CreateCompatibleBitmap(GetDC(0), 32 * MapSize, 32 * MapSize)
If (hBitMap = 0) Then Exit Function
SelectObject hBuffer, hBitMap
Set tSheet = LoadResPicture(1, vbResBitmap) 'Load the sprite sheet
SelectObject hSheet, tSheet.Handle 'Place the sprite sheet picture into its very own shiny device context, isn't that cozy
InitDrawing = True
End Function

Public Sub DrawMap(ByVal Dest As Long)
Dim iCountX As Integer, iCountY As Integer
BitBlt hBuffer, 0, 0, 32 * MapSize, 32 * MapSize, 0, 0, 0, vbWhiteness
For iCountY = 0 To MapSize - 1
    For iCountX = 0 To MapSize - 1
        TileBlt hBuffer, hSheet, tTile(iCountY * MapSize + iCountX).Sprite, iCountX * 32, iCountY * 32
    Next iCountX
Next iCountY
BitBlt Dest, 0, 0, MapSize * 32, MapSize * 32, hBuffer, 0, 0, vbSrcCopy
End Sub

Public Sub CleanUp()
DeleteDC hBuffer
DeleteDC hSheet
DeleteObject hBitMap
End Sub
