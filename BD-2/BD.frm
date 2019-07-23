VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btConsultar 
      Caption         =   "Consultar"
      Height          =   435
      Left            =   4515
      TabIndex        =   10
      Top             =   4200
      Width           =   1275
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   435
      Left            =   5985
      TabIndex        =   9
      Top             =   4200
      Width           =   1380
   End
   Begin VB.Data dataMercadoria 
      Caption         =   "Mercadoria"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\danilo.araujo\Documents\VB6\Project\BD\Estoque.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Mercadoria"
      Top             =   4200
      Width           =   2640
   End
   Begin VB.Frame Frame 
      Height          =   3690
      Left            =   315
      TabIndex        =   0
      Top             =   315
      Width           =   7050
      Begin VB.TextBox txtValorVenda 
         DataField       =   "Valor"
         DataSource      =   "dataMercadoria"
         Height          =   390
         Left            =   1995
         TabIndex        =   8
         Top             =   2205
         Width           =   2745
      End
      Begin VB.TextBox txtQuantidadeEstoque 
         DataField       =   "Quantidade"
         DataSource      =   "dataMercadoria"
         Height          =   390
         Left            =   1995
         TabIndex        =   6
         Top             =   1575
         Width           =   1800
      End
      Begin VB.TextBox txtDescricao 
         DataField       =   "Descrição"
         DataSource      =   "dataMercadoria"
         Height          =   390
         Left            =   1995
         TabIndex        =   4
         Top             =   945
         Width           =   4005
      End
      Begin VB.TextBox txtCodigo 
         DataField       =   "Codigo"
         DataSource      =   "dataMercadoria"
         Height          =   390
         Left            =   1995
         TabIndex        =   2
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblValorVenda 
         AutoSize        =   -1  'True
         Caption         =   "Valor de Venda"
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label lblQuantidadeEstoque 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade no Estoque"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   1575
         Width           =   1680
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   1155
         TabIndex        =   3
         Top             =   945
         Width           =   720
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   1365
         TabIndex        =   1
         Top             =   315
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btConsultar_Click()
    frmConsulta.Show vbModal
    
End Sub
