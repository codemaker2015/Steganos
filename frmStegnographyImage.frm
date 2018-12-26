VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStegnographyImage 
   Caption         =   "Steganos - Stegnography"
   ClientHeight    =   12390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   26625
   LinkTopic       =   "Form1"
   ScaleHeight     =   12390
   ScaleWidth      =   26625
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Steganography With Image"
      Height          =   5775
      Left            =   7005
      TabIndex        =   0
      Top             =   3308
      Width           =   12615
      Begin VB.CommandButton cmdRecover 
         Caption         =   "Recover"
         Default         =   -1  'True
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtNumBits 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Text            =   "2"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.PictureBox picVisible 
         Height          =   2655
         Left            =   1560
         ScaleHeight     =   2595
         ScaleWidth      =   2235
         TabIndex        =   4
         Top             =   1680
         Width           =   2295
      End
      Begin VB.PictureBox picHidden 
         Height          =   2655
         Left            =   3960
         ScaleHeight     =   2595
         ScaleWidth      =   2235
         TabIndex        =   3
         Top             =   1680
         Width           =   2295
      End
      Begin VB.PictureBox picCombined 
         Height          =   2655
         Left            =   6360
         ScaleHeight     =   2595
         ScaleWidth      =   2235
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.PictureBox picRecovered 
         Height          =   2655
         Left            =   8760
         ScaleHeight     =   2595
         ScaleWidth      =   2235
         TabIndex        =   1
         Top             =   1680
         Width           =   2295
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8280
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Cover Image"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   11
         Top             =   1440
         Width           =   2280
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hidden Image Bits:"
         Height          =   195
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Hidden Image"
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   9
         Top             =   1440
         Width           =   2280
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Combined Image"
         Height          =   195
         Index           =   2
         Left            =   6360
         TabIndex        =   8
         Top             =   1440
         Width           =   2280
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Recovered Image"
         Height          =   195
         Index           =   3
         Left            =   8760
         TabIndex        =   7
         Top             =   1440
         Width           =   2280
      End
   End
   Begin VB.Image imgPrevious 
      Height          =   495
      Left            =   240
      Picture         =   "frmStegnographyImage.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmStegnographyImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRecover_Click()
    Dim num_bits As Integer

    MousePointer = vbHourglass

    ' Hide the image.
    num_bits = Val(txtNumBits.Text)

    ' Recover the hidden image.
    RecoverImage picCombined, picRecovered, num_bits

    MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    picVisible.AutoRedraw = True
    picHidden.AutoRedraw = True
    picCombined.AutoRedraw = True
    picRecovered.AutoRedraw = True
    
    picVisible.ScaleMode = vbPixels
    picHidden.ScaleMode = vbPixels
    picCombined.ScaleMode = vbPixels
    picRecovered.ScaleMode = vbPixels
    picVisible.Picture = LoadPicture(file)
End Sub

' Hide and then recover the image.
Private Sub cmdGo_Click()
    Dim num_bits As Integer

    MousePointer = vbHourglass

    ' Hide the image.
    num_bits = Val(txtNumBits.Text)
    HideImage picVisible, picHidden, picCombined, num_bits

    ' Recover the hidden image.
    'RecoverImage picCombined, picRecovered, num_bits
    picRecovered.Picture = LoadPicture()

    MousePointer = vbDefault
End Sub

Private Sub imgPrevious_Click()
    Unload Me
End Sub

Private Sub picCombined_Change()
    'AutoSizeToPicture picCombined
End Sub

Private Sub picHidden_Change()
    AutoSizeToPicture picHidden
End Sub


Private Sub picHidden_Click()
    CommonDialog1.ShowOpen

    picHidden.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub picRecovered_Change()
    'AutoSizeToPicture picRecovered
End Sub

Private Sub picVisible_Change()
    AutoSizeToPicture picVisible
End Sub



