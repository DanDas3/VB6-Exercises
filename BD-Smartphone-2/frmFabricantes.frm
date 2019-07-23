VERSION 5.00
Begin VB.Form frmFabricantes 
   Caption         =   "Fabricantes"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dataFabricantes 
      Caption         =   "Fabricantes"
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
      RecordSource    =   "fabricante"
      Top             =   1890
      Width           =   4635
   End
   Begin VB.CommandButton btFechar 
      Caption         =   "Fechar"
      Height          =   330
      Left            =   3990
      TabIndex        =   8
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton btAlterar 
      Caption         =   "Alterar"
      Height          =   330
      Left            =   2100
      TabIndex        =   7
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton btExcluir 
      Caption         =   "Excluir"
      Height          =   330
      Left            =   3045
      TabIndex        =   6
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton btAtualizar 
      Caption         =   "Atualizar"
      Height          =   330
      Left            =   1155
      TabIndex        =   5
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton btNovo 
      Caption         =   "Novo"
      Height          =   330
      Left            =   210
      TabIndex        =   4
      Top             =   1365
      Width           =   855
   End
   Begin VB.TextBox txtNomeFabricante 
      DataField       =   "nome"
      DataSource      =   "dataFabricantes"
      Height          =   330
      Left            =   945
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "codigo"
      DataSource      =   "dataFabricantes"
      Height          =   330
      Left            =   945
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   1905
   End
   Begin VB.Label lblNomeFabricante 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   420
      Width           =   495
   End
End
Attribute VB_Name = "frmFabricantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btAlterar_Click()
    dataFabricante.UpdateRecord
    dataFabricante.Recordset.Bookmark = dataFabricante.Recordset.LastModified
End Sub

Private Sub btAtualizar_Click()
    dataFabricantes.Refresh
End Sub

Private Sub btExcluir_Click()
    dataFabricantes.Recordset.Delete
    dataFabricantes.Recordset.MoveNext
End Sub

Private Sub btFechar_Click()
    Unload Me
End Sub

Private Sub btNovo_Click()
    dataFabricantes.Recordset.AddNew
End Sub

Private Sub dataFabricantes_Reposition()
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'This will display the current record position
    'for dynasets and snapshots
    dataFabricantes.Caption = "Record: " & (dataFabricantes.Recordset.AbsolutePosition + 1)
    'for the table object you must set the index property when
    'the recordset gets created and use the following line
    'dataFabricante.Caption = "Record: " & (dataFabricante.Recordset.RecordCount * (dataFabricante.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub dataFabricantes_Validate(Action As Integer, Save As Integer)
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

