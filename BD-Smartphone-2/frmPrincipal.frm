VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Gerenciamento de Produtos"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btSistemaVendas 
      Caption         =   "Sistema Vendas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3990
      TabIndex        =   3
      Top             =   2100
      Width           =   2640
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3360
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btBancoDados 
      Caption         =   "Banco de Dados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   630
      TabIndex        =   2
      Top             =   2100
      Width           =   2640
   End
   Begin VB.CommandButton btSmartphones 
      Caption         =   "Smartphones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3990
      TabIndex        =   1
      Top             =   630
      Width           =   2640
   End
   Begin VB.CommandButton btFabricantes 
      Caption         =   "Fabricantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   630
      TabIndex        =   0
      Top             =   630
      Width           =   2640
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btBancoDados_Click()
    CommonDialog.Filter = "Microsoft Access Database (*.mdb) | *.mdb|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Selecione o arquivo"
    CommonDialog.ShowOpen
    PathDatabase = CommonDialog.FileName
End Sub

Private Sub btFabricantes_Click()
    frmFabricantes.Show
End Sub

Private Sub btSistemaVendas_Click()
    frmSistemaVendas.Show
End Sub

Private Sub btSmartphones_Click()
    frmSmartphones.Show
End Sub
