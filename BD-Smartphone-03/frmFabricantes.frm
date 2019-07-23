VERSION 5.00
Begin VB.Form frmFabricantes 
   Caption         =   "Cadastro de Fabricantes"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dataFabricantes 
      Caption         =   "Fabricante"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\danilo.araujo\Documents\VB6\Project\BD-Smartphone-03\smartphone.mdb"
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
      Width           =   6285
   End
   Begin VB.CommandButton btSair 
      Caption         =   "Sair"
      Height          =   405
      Left            =   5460
      TabIndex        =   9
      Top             =   1260
      Width           =   960
   End
   Begin VB.CommandButton btGravar 
      Caption         =   "Gravar"
      Height          =   405
      Left            =   4410
      TabIndex        =   8
      Top             =   1260
      Width           =   960
   End
   Begin VB.CommandButton btExcluir 
      Caption         =   "Excluir"
      Height          =   405
      Left            =   3360
      TabIndex        =   7
      Top             =   1260
      Width           =   960
   End
   Begin VB.CommandButton btConsultar 
      Caption         =   "Consultar"
      Height          =   405
      Left            =   2310
      TabIndex        =   6
      Top             =   1260
      Width           =   960
   End
   Begin VB.CommandButton btAlterar 
      Caption         =   "Alterar"
      Height          =   405
      Left            =   1260
      TabIndex        =   5
      Top             =   1260
      Width           =   960
   End
   Begin VB.CommandButton btIncluir 
      Caption         =   "Incluir"
      Height          =   405
      Left            =   210
      TabIndex        =   4
      Top             =   1260
      Width           =   960
   End
   Begin VB.TextBox txtNomeFabricante 
      DataField       =   "nome"
      DataSource      =   "dataFabricantes"
      Height          =   330
      Left            =   945
      TabIndex        =   3
      Top             =   630
      Width           =   4740
   End
   Begin VB.TextBox txtCodigo 
      DataField       =   "codigo"
      DataSource      =   "dataFabricantes"
      Height          =   330
      Left            =   945
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
   Begin VB.Label lblNomeFabricante 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   210
      TabIndex        =   2
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
Attribute VB_Name = "frmFabricantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btAlterar_Click()
    btIncluir.Enabled = False
    btAlterar.Enabled = False
    btConsultar.Enabled = False
    btExcluir.Enabled = False
    btGravar.Enabled = True
    btSair.Enabled = True
    
    dataFabricantes.Recordset.Edit
End Sub

Private Sub btConsultar_Click()
    dataFabricantes.Recordset.Seek "=", InputBox("Insira o código do fabricante")
    If dataFabricantes.Recordset.NoMatch Then
        MsgBox "Fabricante não cadastrado"
    Else
    End If
End Sub

Private Sub btExcluir_Click()
    
    If MsgBox("Confirma a exclusão do produto?", vbYesNo) = vbYes Then
        dataFabricantes.Recordset.Delete
    End If
End Sub

Private Sub btGravar_Click()

    btIncluir.Enabled = True
    btAlterar.Enabled = True
    btConsultar.Enabled = True
    btExcluir.Enabled = True
    btGravar.Enabled = False
    btSair.Enabled = True
    dataFabricantes.Recordset.Update
End Sub

Private Sub btIncluir_Click()
    dataFabricantes.Recordset.AddNew
    btIncluir.Enabled = False
    btAlterar.Enabled = False
    btConsultar.Enabled = False
    btExcluir.Enabled = False
    btGravar.Enabled = True
    btSair.Enabled = True
    
    txtNomeFabricante.SetFocus
End Sub

Private Sub btSair_Click()
    Unload Me
End Sub

