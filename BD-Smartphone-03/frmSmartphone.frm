VERSION 5.00
Begin VB.Form frmSmartphone 
   Caption         =   "Cadastro de Smartphones"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dataSmartphones 
      Caption         =   "Smartphone"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\danilo.araujo\Documents\VB6\Project\BD-Smartphone-03\smartphone.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "smartphone"
      Top             =   3465
      Width           =   6210
   End
   Begin VB.CommandButton btSair 
      Caption         =   "Sair"
      Height          =   405
      Left            =   5460
      TabIndex        =   17
      Top             =   2940
      Width           =   960
   End
   Begin VB.CommandButton btGravar 
      Caption         =   "Gravar"
      Height          =   405
      Left            =   4410
      TabIndex        =   16
      Top             =   2940
      Width           =   960
   End
   Begin VB.CommandButton btExcluir 
      Caption         =   "Excluir"
      Height          =   405
      Left            =   3360
      TabIndex        =   15
      Top             =   2940
      Width           =   960
   End
   Begin VB.CommandButton btConsultar 
      Caption         =   "Consultar"
      Height          =   405
      Left            =   2310
      TabIndex        =   14
      Top             =   2940
      Width           =   960
   End
   Begin VB.CommandButton btAlterar 
      Caption         =   "Alterar"
      Height          =   405
      Left            =   1260
      TabIndex        =   13
      Top             =   2940
      Width           =   960
   End
   Begin VB.CommandButton btIncluir 
      Caption         =   "Incluir"
      Height          =   405
      Left            =   210
      TabIndex        =   12
      Top             =   2940
      Width           =   960
   End
   Begin VB.TextBox txtVendidos 
      DataField       =   "vendidos"
      DataSource      =   "dataSmartphones"
      Height          =   330
      Left            =   1155
      TabIndex        =   11
      Top             =   2310
      Width           =   1275
   End
   Begin VB.TextBox txtEstoque 
      DataField       =   "estoque"
      DataSource      =   "dataSmartphones"
      Height          =   330
      Left            =   1155
      TabIndex        =   10
      Top             =   1890
      Width           =   1275
   End
   Begin VB.TextBox txtPreco 
      DataField       =   "preco"
      DataSource      =   "dataSmartphones"
      Height          =   330
      Left            =   1155
      TabIndex        =   9
      Top             =   1470
      Width           =   1275
   End
   Begin VB.ComboBox cmbFabricantes 
      Height          =   315
      Left            =   1155
      TabIndex        =   8
      Text            =   "cmbFabricantes"
      Top             =   1050
      Width           =   2325
   End
   Begin VB.TextBox txtNome 
      DataField       =   "nome"
      DataSource      =   "dataSmartphones"
      Height          =   330
      Left            =   1155
      TabIndex        =   7
      Top             =   630
      Width           =   2325
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "codigo"
      DataSource      =   "dataSmartphones"
      Height          =   330
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label lblVendidos 
      AutoSize        =   -1  'True
      Caption         =   "Vendidos"
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   2310
      Width           =   660
   End
   Begin VB.Label lblEstoque 
      AutoSize        =   -1  'True
      Caption         =   "Estoque"
      Height          =   195
      Left            =   210
      TabIndex        =   4
      Top             =   1890
      Width           =   585
   End
   Begin VB.Label lblPreco 
      AutoSize        =   -1  'True
      Caption         =   "Preço"
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   1470
      Width           =   420
   End
   Begin VB.Label lblFabricante 
      AutoSize        =   -1  'True
      Caption         =   "Fabricante"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   1050
      Width           =   750
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   630
      Width           =   420
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   495
   End
End
Attribute VB_Name = "frmSmartphone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FabricanteDB As Database
Dim TBFabricante As Recordset

Private Function preencheLista()
    TBFabricante.MoveFirst
    Do While TBFabricante.EOF = False
        cmbFabricantes.AddItem TBFabricante("nome")
        TBFabricante.MoveNext
    Loop
End Function


Private Sub btAlterar_Click()

    btIncluir.Enabled = False
    btAlterar.Enabled = False
    btConsultar.Enabled = False
    btExcluir.Enabled = False
    btGravar.Enabled = True
    btSair.Enabled = True
    
    dataSmartphones.Recordset.Edit
End Sub

Private Sub btConsultar_Click()
    dataSmartphones.Recordset.Seek "=", InputBox("Digite o código do produto")
    
    If dataSmartphones.Recordset.NoMatch Then
        MsgBox "Produto não cadastrado"
    Else
        
    End If
End Sub

Private Sub btExcluir_Click()
    dataSmartphones.Recordset.Delete
End Sub

Private Sub btGravar_Click()
    'Na form de venda adicionar mais 1 no combo list dos fabricantes
    dataSmartphones.Recordset("fabricante") = cmbFabricantes.ListIndex
    dataSmartphones.Recordset.Update
End Sub

Private Sub btIncluir_Click()
    dataSmartphones.Recordset.AddNew
    
    btIncluir.Enabled = False
    btAlterar.Enabled = False
    btConsultar.Enabled = False
    btExcluir.Enabled = False
    btGravar.Enabled = True
    btSair.Enabled = True
    
    txtNome.SetFocus
End Sub

Private Sub btSair_Click()
    Unload Me
End Sub

Private Sub dataSmartphones_Reposition()
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'This will display the current record position
    'for dynasets and snapshots
    dataSmartphones.Caption = "Record: " & (dataSmartphones.Recordset.AbsolutePosition + 1)
    cmbFabricantes.ListIndex = dataSmartphones.Recordset("fabricante")
    'for the table object you must set the index property when
    'the recordset gets created and use the following line
    'dataFabricante.Caption = "Record: " & (dataFabricante.Recordset.RecordCount * (dataFabricante.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Form_Load()
    Set FabricanteDB = OpenDatabase(PathDatabase)
    Set TBFabricante = FabricanteDB.OpenRecordset("fabricante", dbOpenTable)
    
    preencheLista
    
    TBFabricante.Close
    FabricanteDB.Close
End Sub
