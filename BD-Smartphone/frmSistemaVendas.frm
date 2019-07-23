VERSION 5.00
Begin VB.Form frmSistemaVendas 
   Caption         =   "Sistema Vendas"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btCancelar 
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   6300
      TabIndex        =   12
      Top             =   1575
      Width           =   1065
   End
   Begin VB.CommandButton btConfirmaVenda 
      Caption         =   "Confirmar"
      Height          =   435
      Left            =   6300
      TabIndex        =   11
      Top             =   945
      Width           =   1065
   End
   Begin VB.Frame fraDetalhesProduto 
      Caption         =   "Produto"
      Height          =   2010
      Left            =   420
      TabIndex        =   2
      Top             =   840
      Width           =   5370
      Begin VB.TextBox txtPrecoTotal 
         Height          =   330
         Left            =   1575
         TabIndex        =   10
         Top             =   1470
         Width           =   1695
      End
      Begin VB.TextBox txtQuantidadeCompra 
         Height          =   330
         Left            =   1575
         TabIndex        =   8
         Top             =   1050
         Width           =   1905
      End
      Begin VB.TextBox txtNomeFabricante 
         Height          =   330
         Left            =   1575
         TabIndex        =   6
         Top             =   630
         Width           =   1800
      End
      Begin VB.TextBox txtNomeProduto 
         Height          =   330
         Left            =   1575
         TabIndex        =   4
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label lblPrecoTotal 
         AutoSize        =   -1  'True
         Caption         =   "Preço Total"
         Height          =   195
         Left            =   105
         TabIndex        =   9
         Top             =   1470
         Width           =   825
      End
      Begin VB.Label lblQuantidadeCompra 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   1050
         Width           =   825
      End
      Begin VB.Label lblNomeFabricante 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   630
         Width           =   450
      End
      Begin VB.Label lblNomeProduto 
         AutoSize        =   -1  'True
         Caption         =   "Nome do produto"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   210
         Width           =   1230
      End
   End
   Begin VB.TextBox txtCodigoProduto 
      Height          =   330
      Left            =   1995
      TabIndex        =   1
      Top             =   315
      Width           =   1485
   End
   Begin VB.Label lblCodigoProduto 
      AutoSize        =   -1  'True
      Caption         =   "Código do Produto"
      Height          =   195
      Left            =   525
      TabIndex        =   0
      Top             =   420
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

Private Sub Form_Load()
    Set BancoDados = OpenDatabase(PathDatabase)
    Set TBSmartphone = BancoDados.OpenRecordset("smartphone", dbOpenTable)
    Set TBFabricante = BancoDados.OpenRecordset("fabricante", dbOpenTable)
    
    TBSmartphone.Index = "indCodigoSmartphone"
    TBFabricante.Index = "indCodigoFabricante"
    
    txtCodigoProduto = ""
    fraDetalhesProduto.Enabled = False
    btConfirmaVenda.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TBSmartphone.Close
    TBFabricante.Close
    BancoDados.Close
End Sub

Private Sub txtCodigoProduto_LostFocus()
    
    TBSmartphone.Seek "=", txtCodigoProduto.Text
    MsgBox "Produto: " & TBSmartphone("nome") & " Fabricante: " & TBSmartphone("fabricante")
    If TBSmartphone.NoMatch Then
        MsgBox "Código inválido", vbOKOnly + vbInformation, "Aviso"
    Else
        TBFabricante.Seek "=", TBSmartphone("fabricante")
        fraDetalhesProduto.Enabled = True
        btConfirmaVenda.Enabled = True
        
        txtNomeProduto = TBSmartphone("nome")
        txtNomeFabricante = TBFabricante("nome")
    End If
End Sub

