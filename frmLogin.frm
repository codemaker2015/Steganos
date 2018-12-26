VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Steganos - Login"
   ClientHeight    =   13740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   28680
   LinkTopic       =   "Form2"
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   13740
   ScaleWidth      =   28680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Login"
      Height          =   2655
      Left            =   11513
      TabIndex        =   0
      Top             =   5543
      Width           =   5655
      Begin VB.CommandButton cmdLogin 
         Appearance      =   0  'Flat
         Caption         =   "Login"
         Height          =   495
         Left            =   3240
         TabIndex        =   5
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.Image imgShow 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4680
         Picture         =   "frmLogin.frx":16F26
         Stretch         =   -1  'True
         Top             =   1150
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Passwotrd: "
         Height          =   375
         Left            =   255
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         Height          =   375
         Left            =   255
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Steganos"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   675
      Left            =   12960
      TabIndex        =   6
      Top             =   4680
      Width           =   3015
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Width = 5000
    Text1.Height = 5000
End Sub

Private Sub cmdLogin_Click()
    'Login code
    If txtUsername = "abcd" And txtPassword.Text = "abcd" Then
        Unload Me
        frmFileManagerDrive.Show
    Else
        MsgBox "Wrong Username or Password", vbCritical, "Steganography - Login"
    End If
End Sub

Private Sub imgShow_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    txtPassword.PasswordChar = ""
End Sub

Private Sub imgShow_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
    txtPassword.PasswordChar = "*"
End Sub
