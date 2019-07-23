VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChecar 
      Caption         =   "Processar"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChecar_Click()
    Dim DataDigitada As Date
    DataDigitada = txtData.Text
    
    If DataDigitada > Date Then
        MsgBox "Data posterior"
    ElseIf DataDigitada = Date Then
        MsgBox "Data atual"
    Else
        MsgBox "Data anterior"
    End If
    Print Date
    
End Sub
