VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTesteVariavel2 
      Caption         =   "Teste de Variável 2"
      Height          =   855
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdTesteVariavel 
      Caption         =   "Teste de Variável"
      Height          =   855
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblTexto2 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblTexto1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Teste As String

Private Sub cmdTesteVariavel2_Click()
    Teste = "Segundo Botão"
    lblTexto2.Caption = Teste
End Sub

Private Sub cmdTesteVariavel_Click()
    
    Teste = "Texto Aleatorio"
    lblTexto1.Caption = Teste
End Sub
