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
   Begin VB.TextBox txtBox 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      Caption         =   "Texto de uma Label"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
Print "Botão cancelar pressionado"
End Sub

Private Sub cmdOK_Click()
Print "Botão OK pressionado"
End Sub
