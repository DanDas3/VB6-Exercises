VERSION 5.00
Begin VB.Form frmSistemaVendas 
   Caption         =   "Private Function preencheLista()"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btProcurar 
      Caption         =   "Procurar"
      Height          =   540
      Left            =   5985
      TabIndex        =   13
      Top             =   105
      Width           =   1695
   End
   Begin VB.CommandButton btCancelar 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   5985
      TabIndex        =   12
      Top             =   1470
      Width           =   1695
   End
   Begin VB.CommandButton btConfirmarVenda 
      Caption         =   "Confirmar"
      Height          =   540
      Left            =   5985
      TabIndex        =   11
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtCodigoProduto 
      Height          =   330
      Left            =   1785
      TabIndex        =   1
      Top             =   210
      Width           =   1590
   End
   Begin VB.Frame fraDetalhesProduto 
      Height          =   3165
      Left            =   210
      TabIndex        =   2
      Top             =   735
      Width           =   5055
      Begin VB.TextBox txtQuantidade 
         Height          =   330
         Left            =   1575
         TabIndex        =   8
         Top             =   1155
         Width           =   1590
      End
      Begin VB.TextBox txtPrecoTotal 
         Height          =   330
         Left            =   1575
         TabIndex        =   10
         Top             =   1575
         Width           =   1590
      End
      Begin VB.TextBox txtNomeFabricante 
         Height          =   330
         Left            =   1575
         TabIndex        =   6
         Top             =   735
         Width           =   2640
      End
      Begin VB.TextBox txtNomeProduto 
         Height          =   330
         Left            =   1575
         TabIndex        =   4
         Top             =   315
         Width           =   2640
      End
      Begin VB.Label lblPrecoTotal 
         AutoSize        =   -1  'True
         Caption         =   "Preço Total"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   1575
         Width           =   825
      End
      Begin VB.Label lblQuantidade 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   1155
         Width           =   825
      End
      Begin VB.Label lblNomeFabricante 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   735
         Width           =   450
      End
      Begin VB.Label lblNomeProduto 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Produto"
         Height          =   195
         Left            =   210
         TabIndex        =   3
         Top             =   315
         Width           =   1245
      End
   End
   Begin VB.Label lblCodigoSmartphone 
      AutoSize        =   -1  'True
      Caption         =   "Código do Produto"
      Height          =   195
      Left            =   315
      TabIndex        =   0
      Top             =   210
      Width           =   1320
   End
End
Attribute VB_Name = "frmSistemaVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BancoDados As Database
Dim TBSmartphone As Recordset
Dim TBFabricante As Recordset
Dim Index As Long

Private Sub btConfirmarVenda_Click()
    Dim qtdVendidos As Integer
    Dim qtdEstoque As Integer
    
    TBSmartphone.Edit
    TBSmartphone("vendidos") = TBSmartphone("vendidos") + txtQuantidade
    TBSmartphone("estoque") = TBSmartphone("estoque") - txtQuantidade
    TBSmartphone.Update
    MsgBox "Venda realizada com sucesso!", vbOKOnly + vbInformation
End Sub

Private Sub btProcurar_Click()
    txtCodigoProduto_LostFocus
    
End Sub

Private Sub Form_Load()
    Set BancoDados = OpenDatabase(PathDatabase)
    Set TBSmartphone = BancoDados.OpenRecordset("smartphone", dbOpenTable)
    Set TBFabricante = BancoDados.OpenRecordset("fabricante", dbOpenTable)
    
    TBSmartphone.Index = "indCodigoSmartphone"
    TBFabricante.Index = "indCodigoFabricante"
    
    txtCodigoProduto = ""
    fraDetalhesProduto.Enabled = False
    btConfirmarVenda.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TBSmartphone.Close
    TBFabricante.Close
    BancoDados.Close
End Sub

Private Sub txtCodigoProduto_LostFocus()
    If txtCodigoProduto <> "" Then
        TBSmartphone.Seek "=", txtCodigoProduto.Text
        'MsgBox "Produto: " & TBSmartphone("nome") & " Fabricante: " & TBSmartphone("fabricante")
        If TBSmartphone.NoMatch Then
            MsgBox "Código inválido", vbOKOnly + vbInformation, "Aviso"
            txtCodigoProduto = ""
        Else
            TBFabricante.Seek "=", TBSmartphone("fabricante") + 1
            fraDetalhesProduto.Enabled = True
            btConfirmarVenda.Enabled = True
            txtQuantidade = "0"
            txtNomeProduto = TBSmartphone("nome")
            txtNomeFabricante = TBFabricante("nome")
            txtPrecoTotal = TBSmartphone("preco") * txtQuantidade
            
            btConfirmarVenda.Enabled = True
        End If
    End If
End Sub

Private Sub txtQuantidade_lostfocus()
    txtPrecoTotal = txtQuantidade * TBSmartphone("preco")
End Sub
