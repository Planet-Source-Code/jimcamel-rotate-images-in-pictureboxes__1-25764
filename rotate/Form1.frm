VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Image Rotate by Adrian ""JimCamel"" Clark"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Text            =   "30"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load New Image"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Auto Rotate"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   1200
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "45"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rotate"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   3000
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Y:"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "X:"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Destination Image:"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Source Image:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Degrees:"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = False
Picture2.Cls
RotateSurface Picture1, Picture2, Text1.Text, Text2.Text, Text3.Text
End Sub

Private Sub Command2_Click()
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Command3_Click()
CommonDialog1.Filter = "Graphics Files|*.bmp;*.ico;*.gif;*.jpg;|All Files (*.*)|*.*"
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Timer1_Timer()
Text1.Text = Text1.Text + 5
Picture2.Cls
RotateSurface Picture1, Picture2, Text1.Text, Text2.Text, Text3.Text
End Sub
