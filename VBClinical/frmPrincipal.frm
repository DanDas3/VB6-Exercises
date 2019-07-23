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
   Begin VB.CommandButton btSair 
      Caption         =   "Sair"
      Height          =   645
      Left            =   2415
      TabIndex        =   5
      Top             =   1470
      Width           =   1590
   End
   Begin VB.CommandButton btEntrar 
      Caption         =   "Entrar"
      Height          =   645
      Left            =   525
      TabIndex        =   4
      Top             =   1470
      Width           =   1590
   End
   Begin VB.TextBox txtSenha 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1365
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   1590
   End
   Begin VB.TextBox txtLogin 
      Height          =   330
      Left            =   1365
      TabIndex        =   2
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   630
      TabIndex        =   1
      Top             =   840
      Width           =   465
   End
   Begin VB.Label lblLogin 
      AutoSize        =   -1  'True
      Caption         =   "Login"
      Height          =   195
      Left            =   735
      TabIndex        =   0
      Top             =   420
      Width           =   390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

