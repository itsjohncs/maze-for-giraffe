Attribute VB_Name = "Map"
Option Explicit

Public Type Tile
    Sprite As Integer
    Object As Integer
End Type
Public Type XYStruct
    X As Integer
    Y As Integer
End Type
Public tTile() As Tile, MapSize As Integer, Player As XYStruct
Public Function LoadMap(ByVal FileName As String) As Boolean
Dim sFile As String, iCount As Integer, b8File As Byte
Open App.Path & "\Map.txt" For Random As #1 Len = 1
    ReDim tTile(LOF(1) / 2)
    MapSize = Sqr(LOF(1) / 2)
    For iCount = 1 To LOF(1)
        Get #1, iCount, b8File
        If (iCount / 2 <> Int(iCount / 2)) Then
            tTile(Int((iCount) / 2)).Sprite = b8File
        Else
            tTile(Int((iCount) / 2)).Object = b8File
            If (b8File = 2) Then
                Player.X = (iCount / 2) Mod 15
                Player.Y = Int((iCount / 2 + 1) / 15)
            End If
        End If
    Next iCount
Close #1
frmMain.Width = 32 * MapSize * Screen.TwipsPerPixelX + 120
frmMain.Height = 32 * MapSize * Screen.TwipsPerPixelY + 510
End Function
