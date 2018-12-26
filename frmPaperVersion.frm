VERSION 5.00
Begin VB.Form frmPaperVersion 
   Caption         =   "Help"
   ClientHeight    =   4740
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8325
   Icon            =   "frmPaperVersion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "&Copy To ClipBoard"
   End
End
Attribute VB_Name = "frmPaperVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
    Me.Text1.Height = Me.Height - 675
    Me.Text1.Width = Me.Width - 100
    End If
End Sub

Private Sub mnuCopy_Click()
On Error Resume Next
Clipboard.SetText Me.Text1.Text
End Sub

Private Sub mnuPrint_Click()
If PrinterPresent = False Then Exit Sub
Call PrintString(String(Len(Me.Caption), "-") & vbCrLf & Me.Caption & vbCrLf & String(Len(Me.Caption), "-") & vbCrLf & vbCrLf & Me.Text1.Text, 10, 10, 5, 5)
End Sub
