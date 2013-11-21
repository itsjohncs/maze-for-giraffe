VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Labyrinthus ob Camelopardum"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   581
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDialog 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   720
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   383
      TabIndex        =   0
      Top             =   5040
      Width           =   5775
      Begin VB.TextBox txtPrompt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblDlg 
         BackColor       =   &H0080C0FF&
         Caption         =   "Dialog Text"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bBook As Boolean, bRobe As Boolean, bLamb As Boolean, bKnife As Boolean, bDone As Boolean, bClose As Boolean

Private Sub Command1_Click()
DrawMap Me.hdc
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim t8Player As XYStruct
t8Player = Player
Select Case KeyCode
    Case vbKeyDown, vbKeyS: t8Player.Y = t8Player.Y + 1
    Case vbKeyUp, vbKeyW: t8Player.Y = t8Player.Y - 1
    Case vbKeyLeft, vbKeyA: t8Player.X = t8Player.X - 1
    Case vbKeyRight, vbKeyD: t8Player.X = t8Player.X + 1
    Case Else: Exit Sub
End Select
If (t8Player.X > MapSize - 1 Or _
    t8Player.Y > MapSize - 1 Or _
    t8Player.X < 0 Or t8Player.Y < 0) Then Exit Sub
Debug.Print tTile(t8Player.X + t8Player.Y * MapSize).Object
Select Case tTile(t8Player.X + t8Player.Y * MapSize).Object
    Case 0 'Open space
        tTile(Player.X + Player.Y * MapSize).Object = 0
        tTile(Player.X + Player.Y * MapSize).Sprite = 3
        
        tTile(t8Player.X + t8Player.Y * MapSize).Object = 2
        tTile(t8Player.X + t8Player.Y * MapSize).Sprite = 12
        
        Player = t8Player
    Case 3
        lblDlg.Caption = "I AM THE SEER OF DAVIS! YOU HAVE COME FOR GUIDANCE IN THIS YOUR HOUR OF NEED!!! IN ORDER TO PLEASE THE GODS YOU MUST SACRIFICE A LAMB ON THE CERMONIAL ALTAR IN PRIESTS CLOTHES. ONCE THIS HAS BEEN DONE THE WAY WILL BE OPENED TO YOU!!!!"
        TileBlt picIcon.hdc, hSheet, 8, 0, 0
        picDialog.Visible = True
    Case 4
        If bRobe Then
            tTile(t8Player.X + t8Player.Y * MapSize).Object = 0
            tTile(t8Player.X + t8Player.Y * MapSize).Sprite = 3
            lblDlg.Caption = "You can feel Zeus' love as the lighting stops for you"
            TileBlt picIcon.hdc, hSheet, 1, 0, 0
            picDialog.Visible = True
        Else
            lblDlg.Caption = "Ya... About that.. theres some lighting in the way here..."
            TileBlt picIcon.hdc, hSheet, 1, 0, 0
            picDialog.Visible = True
        End If
    Case 5
        If bBook Then
            tTile(4 * 15 + 12).Object = 0
            tTile(4 * 15 + 12).Sprite = 3
            lblDlg.Caption = "The fire abates for the holder of the book of fire"
            TileBlt picIcon.hdc, hSheet, 4, 0, 0
            picDialog.Visible = True
        Else
            lblDlg.Caption = "That fire looks kind of hot..."
            TileBlt picIcon.hdc, hSheet, 4, 0, 0
            picDialog.Visible = True
        End If
    Case 6
        If bDone Then
            tTile(t8Player.X + t8Player.Y * MapSize).Object = 0
            tTile(t8Player.X + t8Player.Y * MapSize).Sprite = 3
            lblDlg.Caption = "The water solidifies under your feet"
            TileBlt picIcon.hdc, hSheet, 11, 0, 0
            picDialog.Visible = True
        Else
            lblDlg.Caption = "I can see the giraffe in the distance, but the water is blocking my way.."
            TileBlt picIcon.hdc, hSheet, 11, 0, 0
            picDialog.Visible = True
        End If
        Case 7
        tTile(t8Player.X + t8Player.Y * MapSize).Object = 0
        tTile(t8Player.X + t8Player.Y * MapSize).Sprite = 3
        lblDlg.Caption = "Hmm, a lamb, these are usually pretty good for sacrifices.."
        bLamb = True
        TileBlt picIcon.hdc, hSheet, 2, 0, 0
        picDialog.Visible = True
    Case 8
        tTile(1 * 15 + 5).Object = 0
        tTile(1 * 15 + 5).Sprite = 3
        lblDlg.Caption = "With these robes on I can probably please the gods easier.."
        bRobe = True
        TileBlt picIcon.hdc, hSheet, 9, 0, 0
        picDialog.Visible = True
    Case 9
        tTile(t8Player.X + t8Player.Y * MapSize).Object = 0
        tTile(t8Player.X + t8Player.Y * MapSize).Sprite = 3
        lblDlg.Caption = "Whats this knife doing back here?"
        bKnife = True
        TileBlt picIcon.hdc, hSheet, 5, 0, 0
        picDialog.Visible = True
    Case 10
        lblDlg.Caption = "AAEAEAMAAEARUMISAS"
        TileBlt picIcon.hdc, hSheet, 6, 0, 0
        txtPrompt.Visible = True
        picDialog.Visible = True
        txtPrompt.SetFocus
    Case 11
        If (bLamb And bKnife And bRobe) Then
            tTile(t8Player.X + t8Player.Y * MapSize).Object = 0
            tTile(t8Player.X + t8Player.Y * MapSize).Sprite = 3
            lblDlg.Caption = "I have succesfully sacrificed to Poseidon and he has granted me the ability to walk above his domain."
            bDone = True
            TileBlt picIcon.hdc, hSheet, 10, 0, 0
            picDialog.Visible = True
        Else
            lblDlg.Caption = "I don't have everything I need to sacrifice to the gods yet."
            TileBlt picIcon.hdc, hSheet, 10, 0, 0
            picDialog.Visible = True
        End If
    Case 12
        lblDlg.Caption = "YAY!!! I'm the hero! I saved the spirit giraffe so now we can dominate in our dodgeball games again! Hip hip, horray! Give yourself a nice pat on the back for finishing! Maybe even a cookie! Ya, a cookie sounds great! Go get a cookie! (Hit any button to close)"
        TileBlt picIcon.hdc, hSheet, 7, 0, 0
        bClose = True
        picDialog.Visible = True
End Select

DrawMap Me.hdc
End Sub

Private Sub Form_Load()
LoadMap "Map.txt"
If Not InitDrawing Then Unload Me

lblDlg.Caption = "The spirit giraffe has been hidden in this temple! The gods must be pleased to reach the giraffe, GET TO IT!"
TileBlt picIcon.hdc, hSheet, 7, 0, 0
picDialog.Visible = True
End Sub

Private Sub Form_Paint()
DrawMap Me.hdc
End Sub

Private Sub Form_Unload(Cancel As Integer)
CleanUp
End Sub

Private Sub picDialog_KeyDown(KeyCode As Integer, Shift As Integer)
picDialog.Visible = False
If bClose Then Unload Me
End Sub

Private Sub picIcon_KeyDown(KeyCode As Integer, Shift As Integer)
picDialog_KeyDown KeyCode, Shift
End Sub

Private Sub txtPrompt_Change()
If (LCase(txtPrompt) = "is") Then txtPrompt_KeyDown vbKeyReturn, 0
End Sub

Private Sub txtPrompt_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn) Then
    If (LCase(txtPrompt) = "is") Then
        txtPrompt.Visible = False
        lblDlg.Caption = "You have solved the riddle of the book of fire! It will now open its secrets to you (Side effects may include attack by eagle)"
        DoEvents
        picDialog.Visible = True
        tTile(8 * 15 + 1).Object = 0
        tTile(8 * 15 + 1).Sprite = 3
        bBook = True
    Else
        txtPrompt.Visible = False
        txtPrompt.Text = ""
        picDialog.Visible = False
    End If
End If
End Sub

Private Sub txtPrompt_LostFocus()
txtPrompt.Visible = False
txtPrompt.Text = ""
picDialog.Visible = False
End Sub
