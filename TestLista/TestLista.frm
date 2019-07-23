VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOrdenar 
      Caption         =   "Ordenar"
      Height          =   855
      Left            =   3480
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdTexto 
      Caption         =   "Texto"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdIndice 
      Caption         =   "Índice"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuantidade 
      Caption         =   "Quantidade"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   2055
   End
   Begin VB.ListBox lstNomes 
      Columns         =   2
      Height          =   1620
      ItemData        =   "TestLista.frx":0000
      Left            =   3000
      List            =   "TestLista.frx":0016
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label lblTexto 
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblIndice 
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblQuantidade 
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdIndice_Click()
    lblIndice.Caption = lstNomes.ListIndex
    
End Sub

Private Sub cmdQuantidade_Click()
    lblQuantidade.Caption = lstNomes.ListCount
    
End Sub

Private Sub cmdTexto_Click()
    lblTexto = lstNomes.Text
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        End
    End If
End Sub

Private Sub lstNomes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        End
    End If
    
End Sub
