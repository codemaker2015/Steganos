VERSION 5.00
Begin VB.Form frmPreview 
   BackColor       =   &H8000000E&
   Caption         =   "Preview"
   ClientHeight    =   13740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   28680
   LinkTopic       =   "Form3"
   ScaleHeight     =   13740
   ScaleWidth      =   28680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image imgPrevious 
      Height          =   495
      Left            =   360
      Picture         =   "frmPreview.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Image imgPreview 
      Height          =   12000
      Left            =   6338
      Top             =   870
      Width           =   16005
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    imgPreview.Picture = LoadPicture(file)
End Sub

Private Sub imgPreview_Click()
    If GetAttr(file) = 2 And vbHidden Then
        frmOptions.lblHide.Caption = "Show"
    Else
        frmOptions.lblHide.Caption = "Hide"
    End If
    If getextension(file) = "jpg" Or getextension(file) = "JPG" Then
        frmOptions.lblSteganography.Visible = True
        frmOptions.lblSteganography.Enabled = True
    Else
        frmOptions.lblSteganography.Visible = False
        frmOptions.lblSteganography.Enabled = False
    End If
    
    If getextension(file) = "txt" Or getextension(file) = "TXT" Then
        frmOptions.lblCryptography.Visible = True
        frmOptions.lblCryptography.Enabled = True
    Else
        frmOptions.lblCryptography.Visible = False
        frmOptions.lblCryptography.Enabled = False
    End If
    
    frmOptions.Left = imgPreview.Left + imgPreview.Width / 2
    frmOptions.Top = imgPreview.Top + imgPreview.Height / 2
    frmOptions.Show

End Sub

Private Sub imgPrevious_Click()
    Unload Me
    
End Sub
