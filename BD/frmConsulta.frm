VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConsulta 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   540
      Left            =   7350
      TabIndex        =   1
      Top             =   4935
      Width           =   1275
   End
   Begin MSDBGrid.DBGrid gridVendas 
      Bindings        =   "frmConsulta.frx":0000
      Height          =   4530
      Left            =   315
      OleObjectBlob   =   "frmConsulta.frx":0018
      TabIndex        =   0
      Top             =   210
      Width           =   6840
   End
   Begin VB.Data datVendas 
      Caption         =   "Dados das Vendas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\danilo.araujo\Documents\VB6\Project\BD\Estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Vendas"
      Top             =   4935
      Width           =   6840
   End
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
    Hide
End Sub

