VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImagem 
      Height          =   3690
      Left            =   315
      ScaleHeight     =   3630
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   315
      Width           =   5475
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6930
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btAbrir 
      Caption         =   "Abrir"
      Height          =   330
      Left            =   315
      TabIndex        =   0
      Top             =   4200
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Path As String
Dim pic As StdPicture
Dim intWidth As Integer
Dim intHeight As Integer
Dim sngRatio As Single

Private Sub btAbrir_Click()

    CommonDialog.Filter = "Images (*.bmp) | *.bmp|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Selecione o arquivo"
    CommonDialog.ShowOpen
    Path = CommonDialog.FileName
    
    Set pic = LoadPicture(Path)
    intWidth = ScaleX(pic.Width, vbHimetric, vbTwips)
    intHeight = ScaleY(pic.Height, vbHimetric, vbTwips)
    
    sngRatio = picImagem.Width / intWidth
    If intHeight * sngRatio > picImagem.Height Then
        sngRatio = picImagem.Height / intHeight
    End If
    
    MsgBox Path
    picImagem.AutoRedraw = True
    picImagem.Picture = LoadPicture(Path)
    picImagem.Height = intHeight
    picImagem.Width = intWidth
    
    picImagem.PaintPicture pic, 0, 0, intWidth * sngRatio, intHeight * sngRatio
    'MsgBox intHeight
    'Form1.Height = intHeight + 1000 + btAbrir.Height
    'btAbrir.Move 0#, 0#, -(CSng(intHeight)), 0#
    'MsgBox Form1.Height
End Sub
