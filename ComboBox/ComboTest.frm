VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btAlterar 
      Caption         =   "Alterar"
      Height          =   330
      Left            =   1260
      TabIndex        =   3
      Top             =   2730
      Width           =   1065
   End
   Begin VB.TextBox txtIndex 
      Height          =   330
      Left            =   1365
      TabIndex        =   2
      Top             =   2100
      Width           =   540
   End
   Begin VB.ComboBox cboItem 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btAlterar_Click()
    cboItem.ListIndex = txtIndex
End Sub

Private Sub Form_Load()
    cboItem.AddItem "Maria José"
    cboItem.AddItem "Roberto Carlos"
    cboItem.AddItem "Alberto Maia"
    
End Sub

Private Sub cboItem_Click()
    lblItem.Caption = "Nome selecionado: " & cboItem.Text
End Sub
