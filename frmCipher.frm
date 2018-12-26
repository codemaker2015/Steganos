VERSION 5.00
Begin VB.Form frmCipher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Cipher Classics"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7470
   Icon            =   "frmCipher.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMain 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   7215
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2760
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtK1 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtGroupsPerLine 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "10"
         Top             =   250
         Width           =   375
      End
      Begin VB.TextBox txtK2 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtK3 
         Height          =   285
         Left            =   2160
         TabIndex        =   5
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Caption         =   "Key1"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Groups per line"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Right Justify
         Caption         =   "Key2"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lbl3 
         Alignment       =   1  'Right Justify
         Caption         =   "Key3"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.Menu mnuEncode 
      Caption         =   "&Encode"
   End
   Begin VB.Menu mnuDecode 
      Caption         =   "&Decode"
   End
   Begin VB.Menu mnuUndo 
      Caption         =   "&Undo"
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuClipBoard 
      Caption         =   "&To ClipBoard"
   End
   Begin VB.Menu mnuPaperVersion 
      Caption         =   "Paper &Version Help"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmCipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strPaperVersion As String
Private strUndo As String

Private Sub Form_Load()
Me.Combo1.AddItem "ADFGVX Cipher"
Me.Combo1.AddItem "Ceasar Shift ( ROT cipher )"
Me.Combo1.AddItem "Columnar Transposition"
Me.Combo1.AddItem "Double Columnar Transposition"
Me.Combo1.AddItem "Playfair"
Me.Combo1.AddItem "Straddling Checkerboard"
Me.Combo1.AddItem "Straddling Checkerboard with Double Columnar"
Me.Combo1.AddItem "Vigenére"
Me.Combo1.ListIndex = 0
If Printer.DeviceName <> "" Then
   PrinterPresent = True
   Else
   PrinterPresent = False
   End If
End Sub

Private Sub Combo1_Click()
Select Case Me.Combo1.ListIndex
Case 0
    'ADFGVX
    Me.txtK2.Visible = True
    Me.txtK3.Visible = False
    Me.lbl1.Caption = "Square Key"
    Me.lbl2.Caption = "Columnar Key"
    Me.lbl3.Caption = ""
    strPaperVersion = "ADFGVX"
Case 1
    'Ceasar Shift
    Me.txtK2.Visible = False
    Me.txtK3.Visible = False
    Me.lbl1.Caption = "Shift Key"
    Me.lbl2.Caption = ""
    Me.lbl3.Caption = ""
    strPaperVersion = "Ceasar Shift"
Case 2
    'Columnar Transposition
    Me.txtK2.Visible = False
    Me.txtK3.Visible = False
    Me.lbl1.Caption = "Columnar Key"
    Me.lbl2.Caption = ""
    Me.lbl3.Caption = ""
    strPaperVersion = "Columnar Transposition"
Case 3
    'Double Columnar Transposition
    Me.txtK2.Visible = True
    Me.txtK3.Visible = False
    Me.lbl1.Caption = "1st Columnar Key"
    Me.lbl2.Caption = "2nd Columnar Key"
    Me.lbl3.Caption = ""
    strPaperVersion = "Double Columnar"
Case 4
    'Plaifair
    Me.txtK2.Visible = False
    Me.txtK3.Visible = False
    Me.lbl1.Caption = "Playfair key"
    Me.lbl2.Caption = ""
    Me.lbl3.Caption = ""
    strPaperVersion = "Playfair"
Case 5
    'Straddling Checkerboard
    Me.txtK2.Visible = False
    Me.txtK3.Visible = False
    Me.lbl1.Caption = "Checkerboard Key"
    Me.lbl2.Caption = ""
    Me.lbl3.Caption = ""
    strPaperVersion = "Straddling Checkerboard"
Case 6
    'Straddling Checkerboard with Double Columnar
    Me.txtK2.Visible = True
    Me.txtK3.Visible = True
    Me.lbl1.Caption = "Checkerboard Key"
    Me.lbl2.Caption = "1st Columnar Key"
    Me.lbl3.Caption = "2nd Columnar Key"
    strPaperVersion = "Checkerboard with Columnar"
Case 7
    'Vigenére
    Me.txtK2.Visible = False
    Me.txtK3.Visible = False
    Me.lbl1.Caption = "Vigenére Key"
    Me.lbl2.Caption = ""
    Me.lbl3.Caption = ""
    strPaperVersion = "Vigenere"
End Select
Me.txtK1.Text = ""
Me.txtK2.Text = ""
Me.txtK3.Text = ""
If frmPaperVersion.Visible = True Then loadPaperVersion (strPaperVersion)
End Sub

Private Sub mnuEncode_Click()
Dim tmpOutput As String
strUndo = Me.txtMain.Text

If Len(Me.txtMain.Text) > 10000 Then
    MsgBox "The text is exceeds the limit of 10,000 characters", vbCritical
    Exit Sub
    End If

Screen.MousePointer = 11
Select Case Me.Combo1.ListIndex
Case 0
    'ADFGVX
    tmpOutput = EncodeADFGVX(Me.txtMain.Text, Me.txtK1.Text, Me.txtK2.Text)
Case 1
    'Ceasar Shift
    tmpOutput = EncodeCeasar(Me.txtMain.Text, Me.txtK1.Text)
Case 2
    'Columnar Transposition
    tmpOutput = EncodeColumnar(Me.txtMain.Text, Me.txtK1.Text)
Case 3
    'Double Columnar Transposition
    tmpOutput = EncodeDoubleColumnar(Me.txtMain.Text, Me.txtK1.Text, Me.txtK2.Text)
Case 4
    'Playfair
    tmpOutput = EncodePlayFair(Me.txtMain.Text, Me.txtK1.Text)
Case 5
    'Straddling Checkerboard
    tmpOutput = EncodeCheckerBoard(Me.txtMain.Text, Me.txtK1.Text)
Case 6
    'Straddling Checkerboard with Double Columnar
    tmpOutput = EncodeCheckAndColumnar(Me.txtMain.Text, Me.txtK1.Text, Me.txtK2.Text, Me.txtK3.Text)
Case 7
    'Vigenére
    tmpOutput = EncodeVigenere(Me.txtMain.Text, Me.txtK1.Text)
End Select
If tmpOutput <> "" Then Me.txtMain.Text = MakeGroups(tmpOutput, True, Val(Me.txtGroupsPerLine.Text))
Screen.MousePointer = 0
End Sub

Private Sub mnuDecode_Click()
Dim tmpOutput As String
strUndo = Me.txtMain.Text
Screen.MousePointer = 11
Select Case Me.Combo1.ListIndex
Case 0
    'ADFGVX
    tmpOutput = DecodeADFGVX(Me.txtMain.Text, Me.txtK1.Text, Me.txtK2.Text)
Case 1
    'Ceasar Shift
    tmpOutput = DecodeCeasar(Me.txtMain.Text, Me.txtK1.Text)
Case 2
    'Columnar Transposition
    tmpOutput = DecodeColumnar(Me.txtMain.Text, Me.txtK1.Text)
Case 3
    'Double Columnar Transposition
    tmpOutput = DecodeDoubleColumnar(Me.txtMain.Text, Me.txtK1.Text, Me.txtK2.Text)
Case 4
    'Playfair
    tmpOutput = DecodePlayFair(Me.txtMain.Text, Me.txtK1.Text)
Case 5
    'Straddling Checkerboard
    tmpOutput = DecodeCheckerBoard(Me.txtMain.Text, Me.txtK1.Text)
Case 6
    'Straddling Checkerboard with Double Columnar
    tmpOutput = DecodeCheckAndColumnar(Me.txtMain.Text, Me.txtK1.Text, Me.txtK2.Text, Me.txtK3.Text)
Case 7
    'Vigenére
    tmpOutput = DecodeVigenere(Me.txtMain.Text, Me.txtK1.Text)
End Select
If tmpOutput <> "" Then Me.txtMain.Text = tmpOutput
Screen.MousePointer = 0
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmPaperVersion
Unload frmAbout
End Sub

Private Sub mnuPaperVersion_Click()
loadPaperVersion (strPaperVersion)
frmPaperVersion.Show
End Sub

Private Sub mnuPrint_Click()
If PrinterPresent = False Then Exit Sub
Call PrintString("*** BEGIN OF MESSAGE ***" & vbCrLf & vbCrLf & Me.txtMain & vbCrLf & vbCrLf & "*** END OF MESSAGE ***", 10, 10, 5, 5)
End Sub

Private Sub mnuUndo_Click()
Me.txtMain.Text = strUndo
End Sub

Private Sub mnuClipBoard_Click()
On Error Resume Next
Clipboard.SetText Me.txtMain.Text
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show (vbModal)
End Sub

Private Sub txtGroupsPerLine_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
End Sub

Private Sub txtGroupsPerLine_GotFocus()
txtGroupsPerLine.SelStart = 0
txtGroupsPerLine.SelLength = Len(txtGroupsPerLine.Text)
End Sub


Private Sub txtMain_KeyPress(KeyAscii As Integer)
'limit text input
Select Case Me.Combo1.ListIndex
Case 0
    'ADFGVX
    KeyAscii = TestKey(KeyAscii, True, True, False)
Case 1
    'Ceasar Shift
    KeyAscii = TestKey(KeyAscii, False, False, False)
Case 2
    'Columnar Transposition
    KeyAscii = TestKey(KeyAscii, False, True, False)
Case 3
    'Double Columnar Transposition
    KeyAscii = TestKey(KeyAscii, False, True, False)
Case 4
    'Playfair
    KeyAscii = TestKey(KeyAscii, False, True, False)
    If KeyAscii = Asc("J") Then KeyAscii = Asc("I")
Case 5
    'Straddling Checkerboard
    KeyAscii = TestKey(KeyAscii, False, True, True)
Case 6
    'Straddling Checkerboard with Double Columnar
    KeyAscii = TestKey(KeyAscii, False, True, True)
Case 7
    'Vigenére
    KeyAscii = TestKey(KeyAscii, False, True, False)
End Select
End Sub

Private Sub txtK1_KeyPress(KeyAscii As Integer)
KeyAscii = TestKey(KeyAscii, False, False, False)
End Sub

Private Sub txtK2_KeyPress(KeyAscii As Integer)
KeyAscii = TestKey(KeyAscii, False, False, False)
End Sub

Private Sub txtK3_KeyPress(KeyAscii As Integer)
KeyAscii = TestKey(KeyAscii, False, False, False)
End Sub


