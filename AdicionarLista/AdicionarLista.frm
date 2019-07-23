VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstItens 
      Height          =   1815
      Left            =   2880
      TabIndex        =   4
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton cmdLimparTudo 
      Caption         =   "Limpar Tudo"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "Adicionar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtItem 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdicionar_Click()
    lstItens.AddItem txtItem.Text
    txtItem.Text = ""
    
End Sub

Private Sub cmdLimparTudo_Click()
    lstItens.Clear
    
End Sub

Private Sub cmdRemover_Click()
    If lstItens.ListIndex >= 0 Then
        lstItens.RemoveItem lstItens.ListIndex
    End If
    cmdRemover.Enabled = False
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        End
    End If
End Sub

Private Sub lstItens_Click()
    cmdRemover.Enabled = True
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        End
    End If
End Sub

Private Sub txtItem_Change()
    If txtItem.Text = "" Then
        cmdAdicionar.Enabled = False
    Else
        cmdAdicionar.Enabled = True
    End If
    
End Sub
