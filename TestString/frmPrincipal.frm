VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Testes com String"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3045
      TabIndex        =   1
      Top             =   210
      Width           =   1380
   End
   Begin VB.TextBox txtStr1 
      Height          =   435
      Left            =   420
      TabIndex        =   0
      Top             =   210
      Width           =   2430
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btPrint_Click()
    Dim strLen As Integer
    Dim index As Integer
    strLen = Len(txtStr1)
    
    index = 1
    
    Do While index <= strLen
        ' Gets an individual char in string
        Print Mid(txtStr1.Text, index, 1)
        index = index + 1
    Loop

End Sub

