VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstDireita 
      Height          =   1425
      Left            =   4440
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdMoverEsquerda 
      Caption         =   "<"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton cmdMoverDireita 
      Caption         =   ">"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.ListBox lstEsquerda 
      Height          =   1425
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "Adicionar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtItem 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdicionar_Click()
    lstEsquerda.AddItem txtItem.Text
    txtItem.Text = ""
End Sub

Private Sub Form_Load()
    
End Sub

Private Sub txtItem_Change()
    If txtItem.Text = "" Then
        cmdAdicionar.Enabled = False
    Else
        cmdAdicionar.Enabled = True
    End If
    
End Sub

Private Sub cmdMoverDireita_Click()
    If lstEsquerda.ListIndex >= 0 Then
        lstDireita.AddItem lstEsquerda.Text
        lstEsquerda.RemoveItem lstEsquerda.ListIndex
    End If
    
End Sub

Private Sub cmdMoverEsquerda_Click()
    If lstDireita.ListIndex >= 0 Then
        lstEsquerda.AddItem lstDireita
        
        lstDireita.RemoveItem lstDireita.ListIndex
    End If
    
End Sub

