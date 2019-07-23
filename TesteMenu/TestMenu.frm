VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQualquer 
      Caption         =   "OK"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Index           =   0
      Begin VB.Menu mnuMercadoria 
         Caption         =   "&Mercadoria"
      End
      Begin VB.Menu mnuFornecedor 
         Caption         =   "&Fornecedor"
      End
      Begin VB.Menu mnuSeparacao 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuLancamento 
      Caption         =   "&Lançamento"
      Begin VB.Menu mnuEntrada 
         Caption         =   "&Entrada"
      End
      Begin VB.Menu mnuSaida 
         Caption         =   "&Saída"
      End
   End
   Begin VB.Menu mnuRelatorio 
      Caption         =   "&Relatório"
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdQualquer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuLancamento
    End If
End Sub

Private Sub mnuEntrada_Click()
    MsgBox "Entrada de Mercadorias"
End Sub

Private Sub mnuSaida_Click()
    Dim Msg As String
    Msg = "Saída de Mercadorias"
    MsgBox Msg
End Sub

Private Sub mnuSair_Click()
    End
End Sub
