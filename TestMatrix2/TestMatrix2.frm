VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkItens 
      Caption         =   "Macarrão"
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   3375
   End
   Begin VB.CheckBox chkItens 
      Caption         =   "Feijão"
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CheckBox chkItens 
      Caption         =   "Arroz"
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblTest 
      Height          =   135
      Left            =   960
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkItens_Click(Index As Integer)
    lblTest.Caption = chkItens(Index).Caption
End Sub
