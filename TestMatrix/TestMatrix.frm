VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdComando 
      Caption         =   "OK"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ListBox lstItens 
      Height          =   2985
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdComando_Click()
    MsgBox lstItens.List(2)
End Sub

Private Sub Form_Load()
    lstItens.AddItem "Arroz"
    lstItens.AddItem "Feijão"
    lstItens.AddItem "Macarrão"
End Sub
