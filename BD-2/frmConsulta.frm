VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConsulta 
   Caption         =   "Form2"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8925
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btFechar 
      Caption         =   "Fechar"
      Height          =   435
      Left            =   5145
      TabIndex        =   1
      Top             =   3045
      Width           =   1380
   End
   Begin VB.Data datVendas 
      Caption         =   "Dados das Vendas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\danilo.araujo\Documents\VB6\Project\BD\Estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Vendas"
      Top             =   3045
      Width           =   3585
   End
   Begin MSDBGrid.DBGrid gridVendas 
      Bindings        =   "frmConsulta.frx":0000
      Height          =   2745
      Left            =   105
      OleObjectBlob   =   "frmConsulta.frx":0018
      TabIndex        =   0
      Top             =   105
      Width           =   6630
   End
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btFechar_Click()
    Hide
End Sub
