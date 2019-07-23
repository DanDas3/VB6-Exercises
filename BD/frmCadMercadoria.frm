VERSION 5.00
Begin VB.Form frmCadMercadoria 
   Caption         =   "Cadastro de Mercadoria"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProximo 
      Caption         =   "Próximo >>"
      Height          =   495
      Left            =   2040
      TabIndex        =   16
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<< Anterior"
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   4095
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   8175
      Begin VB.TextBox txtValorVenda 
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtQuantidadeEstoque 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtDescricao 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtCodigo 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblValorVenda 
         AutoSize        =   -1  'True
         Caption         =   "Valor da Venda"
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblQuantidadeEstoque 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade no estoque"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1440
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   615
      Left            =   8880
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   615
      Left            =   8880
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   615
      Left            =   8880
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   615
      Left            =   8880
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   615
      Left            =   8880
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   615
      Left            =   8880
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmCadMercadoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BancoDeDados As Database
Dim TBMercadoria As Recordset

Private Sub cmdConsultar_Click()
    Dim ProcurarCodigo As String
    
    ProcurarCodigo = InputBox("Digite o código a ser procurado")
    
    TBMercadoria.Seek "=", ProcurarCodigo
    
    If TBMercadoria.NoMatch = True Then
        MsgBox "Mercadoria não cadastrada", vbOKOnly, "ESTOQUE"
        TBMercadoria.MovePrevious
        
    End If
    AtualizarFormulario
    
End Sub

Private Sub cmdExcluir_Click()
    If TBMercadoria.EOF = False Then
        If MsgBox("Confirma a exclusão do produto?", vbYesNo) = vbYes Then
            TBMercadoria.Delete
            cmdAnterior_Click
            
        End If
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Path = "C:\Users\danilo.araujo\Documents\VB6\Project\BD"
    Set BancoDeDados = OpenDatabase(Path & "\Estoque.MDB")
    Set TBMercadoria = BancoDeDados.OpenRecordset("Mercadoria", dbOpenTable)

    TBMercadoria.Index = "IndCodigo"
    cmdGravar.Enabled = False
    Frame1.Enabled = False
    If TBMercadoria.EOF = False Then
        AtualizarFormulario
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TBMercadoria.Close
    BancoDeDados.Close
    
End Sub

Private Sub cmdIncluir_Click()
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdConsultar.Enabled = False
    cmdExcluir.Enabled = False
    cmdAnterior.Enabled = False
    cmdProximo.Enabled = False
    cmdGravar.Enabled = True
    cmdSair.Enabled = True
    
    LimpaFormulario
    Frame1.Enabled = True
    txtCodigo.SetFocus
    
End Sub

Private Sub cmdAlterar_Click()
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdConsultar.Enabled = False
    cmdExcluir.Enabled = False
    cmdAnterior.Enabled = False
    cmdProximo.Enabled = False
    cmdGravar.Enabled = True
    cmdSair.Enabled = True
    
    Frame1.Enabled = True
    txtCodigo.Enabled = False
    txtDescricao.SetFocus
    TBMercadoria.Edit
End Sub

Private Sub cmdGravar_Click()
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    cmdConsultar.Enabled = True
    cmdExcluir.Enabled = True
    cmdAnterior.Enabled = True
    cmdProximo.Enabled = True
    cmdGravar.Enabled = False
    cmdSair.Enabled = True
    
    Frame1.Enabled = False
    txtCodigo.Enabled = True
    
    AtualizaCampos
    TBMercadoria.Update
    LimpaFormulario
End Sub


Private Sub txtCodigo_LostFocus()
    txtCodigo.Text = Format(txtCodigo.Text, "000")
    
    TBMercadoria.Seek "=", txtCodigo.Text
    
    If TBMercadoria.NoMatch = False Then
        MsgBox "Mercadoria já existente. Tente outro Código"
        
        AtualizarFormulario
        cmdIncluir.Enabled = True
        cmdAlterar.Enabled = True
        cmdConsultar.Enabled = True
        cmdExcluir.Enabled = True
        cmdAnterior.Enabled = True
        cmdProximo.Enabled = True
        cmdGravar.Enabled = False
        cmdSair.Enabled = True
        Frame1.Enabled = False
    Else
        TBMercadoria.AddNew
    End If
    
End Sub

Private Sub txtValorVenda_LostFocus()
    txtValorVenda.Text = Format(txtValorVenda.Text, "Standard")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys ("{TAB}")
        KeyAscii = 0
    End If
    
End Sub

Private Function AtualizaCampos()
    TBMercadoria("Codigo") = txtCodigo
    TBMercadoria("Descrição") = txtDescricao
    TBMercadoria("Quantidade") = txtQuantidadeEstoque
    TBMercadoria("Valor") = txtValorVenda
    
End Function

Private Function AtualizarFormulario()
    If TBMercadoria.EOF = False And TBMercadoria.BOF = False Then
        txtCodigo = TBMercadoria("Codigo")
        txtDescricao = TBMercadoria("Descrição")
        txtQuantidadeEstoque = TBMercadoria("Quantidade")
        txtValorVenda = Format(TBMercadoria("Valor"), "Standard")
    End If
End Function

Private Function LimpaFormulario()
    txtCodigo = ""
    txtDescricao = ""
    txtQuantidadeEstoque = ""
    txtValorVenda = ""
    
End Function

Private Sub cmdProximo_Click()
    
    If TBMercadoria.EOF = True Then
        If TBMercadoria.BOF = False Then
            TBMercadoria.MovePrevious
        End If
    Else
        TBMercadoria.MoveNext
    End If
    AtualizarFormulario
    
End Sub

Private Sub cmdAnterior_Click()
    
    If TBMercadoria.BOF = True Then
        If TBMercadoria.EOF = False Then
            TBMercadoria.MoveNext
        End If
    Else
        TBMercadoria.MovePrevious
    End If
    AtualizarFormulario
    
End Sub

