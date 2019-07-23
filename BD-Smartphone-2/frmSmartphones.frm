VERSION 5.00
Begin VB.Form frmSmartphones 
   Caption         =   "Smartphones"
   ClientHeight    =   3705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dataSmartphone 
      Caption         =   "Smartphone"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\danilo.araujo\Documents\VB6\Project\BD-Smartphone\smartphone.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "smartphone"
      Top             =   3150
      Width           =   4740
   End
   Begin VB.CommandButton btFechar 
      Caption         =   "Fechar"
      Height          =   330
      Left            =   3990
      TabIndex        =   16
      Top             =   2730
      Width           =   855
   End
   Begin VB.CommandButton btExcluir 
      Caption         =   "Excluir"
      Height          =   330
      Left            =   3045
      TabIndex        =   15
      Top             =   2730
      Width           =   855
   End
   Begin VB.CommandButton btAlterar 
      Caption         =   "Alterar"
      Height          =   330
      Left            =   2100
      TabIndex        =   14
      Top             =   2730
      Width           =   855
   End
   Begin VB.CommandButton btAtualizar 
      Caption         =   "Atualizar"
      Height          =   330
      Left            =   1155
      TabIndex        =   13
      Top             =   2730
      Width           =   855
   End
   Begin VB.CommandButton btNovo 
      Caption         =   "Novo"
      Height          =   330
      Left            =   210
      TabIndex        =   12
      Top             =   2730
      Width           =   855
   End
   Begin VB.TextBox txtVendidos 
      DataField       =   "vendidos"
      DataSource      =   "dataSmartphone"
      Height          =   330
      Left            =   945
      TabIndex        =   11
      Top             =   2310
      Width           =   1485
   End
   Begin VB.TextBox txtEstoque 
      DataField       =   "estoque"
      DataSource      =   "dataSmartphone"
      Height          =   330
      Left            =   945
      TabIndex        =   10
      Top             =   1890
      Width           =   1485
   End
   Begin VB.TextBox txtPreco 
      DataField       =   "preco"
      DataSource      =   "dataSmartphone"
      Height          =   330
      Left            =   945
      TabIndex        =   9
      Top             =   1470
      Width           =   1485
   End
   Begin VB.ComboBox lstFabricantes 
      Height          =   315
      Left            =   945
      TabIndex        =   8
      Top             =   1050
      Width           =   2010
   End
   Begin VB.TextBox txtNomeFabricante 
      DataField       =   "nome"
      DataSource      =   "dataSmartphone"
      Height          =   330
      Left            =   945
      TabIndex        =   7
      Top             =   630
      Width           =   2220
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "codigo"
      DataSource      =   "dataSmartphone"
      Height          =   330
      Left            =   945
      TabIndex        =   6
      Top             =   210
      Width           =   1485
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
      Left            =   315
      TabIndex        =   3
      Top             =   1470
      Width           =   420
   End
   Begin VB.Label lblFabricante 
      AutoSize        =   -1  'True
      Caption         =   "Fabricante"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   1050
      Width           =   750
   End
   Begin VB.Label lblNomeSmartphone 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   315
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
Attribute VB_Name = "frmSmartphones"
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
        lstFabricantes.AddItem TBFabricante("nome")
        TBFabricante.MoveNext
    Loop
End Function

Private Sub btAlterar_Click()
    dataSmartphone.UpdateRecord
    dataSmartphone.Recordset.Bookmark = dataSmartphone.Recordset.LastModified
End Sub

Private Sub btAtualizar_Click()
    dataSmartphone.Recordset("fabricante") = lstFabricantes.ListIndex
    dataSmartphone.Refresh
End Sub

Private Sub btExcluir_Click()
    dataSmartphone.Recordset.Delete
    dataSmartphone.Recordset.MoveNext
End Sub

Private Sub btFechar_Click()
    Unload Me
End Sub

Private Sub btNovo_Click()
    dataSmartphone.Recordset.AddNew
End Sub

Private Sub dataSmartphone_Validate(Action As Integer, Save As Integer)
    'This is where you put validation code
    'This event gets called when the following actions occur
    Select Case Action
        Case vbDataActionMoveFirst
        Case vbDataActionMovePrevious
        Case vbDataActionMoveNext
        Case vbDataActionMoveLast
        Case vbDataActionAddNew
        Case vbDataActionUpdate
        Case vbDataActionDelete
        Case vbDataActionFind
        Case vbDataActionBookmark
        Case vbDataActionClose
      End Select
      Screen.MousePointer = vbHourglass
End Sub

Private Sub dataSmartphone_Reposition()
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'This will display the current record position
    'for dynasets and snapshots
    dataSmartphone.Caption = "Record: " & (dataSmartphone.Recordset.AbsolutePosition + 1)
    lstFabricantes.ListIndex = dataSmartphone.Recordset("fabricante")
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
