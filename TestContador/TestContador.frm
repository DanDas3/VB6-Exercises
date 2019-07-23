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
   Begin VB.CommandButton cmdComecar 
      Caption         =   "Começar"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtContador 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdComecar_Click()
    Dim Contador As Integer
    Contador = 0
    
    Do While Contador <= 10000
        txtContador.Text = Contador
        Contador = Contador + 1
    Loop
End Sub
