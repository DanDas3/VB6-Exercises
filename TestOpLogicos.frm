VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtValor1 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblMsg 
      Caption         =   "Insira o valor"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub cmdCalcular_Click()
    If txtValor1 Mod 2 = 0 Then
        MsgBox "Par"
    Else
        MsgBox "Ímpar"
    End If
    
End Sub

