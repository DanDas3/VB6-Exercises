VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecutar 
      Caption         =   "Executar"
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtNumero 
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label lblResultado6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4680
      TabIndex        =   7
      Top             =   4200
      Width           =   45
   End
   Begin VB.Label lblResultado5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4800
      TabIndex        =   6
      Top             =   3720
      Width           =   45
   End
   Begin VB.Label lblResultado4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4800
      TabIndex        =   5
      Top             =   3000
      Width           =   45
   End
   Begin VB.Label lblResultado3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   4080
      Width           =   45
   End
   Begin VB.Label lblResultado2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Width           =   45
   End
   Begin VB.Label lblResultado1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   3000
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExecutar_Click()
    lblResultado1.Caption = "Função ABS: " & Abs(txtNumero.Text)
    lblResultado2.Caption = "Função Fix: " & Fix(txtNumero.Text)
    lblResultado3.Caption = "Função Int: " & Int(txtNumero.Text)
    lblResultado4.Caption = "Função Sgn: " & Sgn(txtNumero.Text)
    lblResultado5.Caption = "Função Sqr: " & Sqr(txtNumero.Text)
    lblResultado6.Caption = "Função Tan: " & Tan(txtNumero.Text)
End Sub

