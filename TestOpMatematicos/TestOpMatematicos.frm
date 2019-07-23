VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      Height          =   975
      Left            =   2760
      TabIndex        =   0
      Top             =   3480
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalcular_Click()
    Dim Indice As Integer
    Dim Valor As Currency
    
    Valor = 2425
    Indice = 2.5
    Print Valor * Indice / 100
End Sub
