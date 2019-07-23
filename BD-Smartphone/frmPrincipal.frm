VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Sistema de Controle de Produtos"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2415
      Top             =   1890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btLocalizaBancoDados 
      Caption         =   "Banco de Dados"
      Height          =   540
      Left            =   315
      TabIndex        =   3
      Top             =   1890
      Width           =   1905
   End
   Begin VB.CommandButton btSmartphone 
      Caption         =   "Smartphones"
      Height          =   540
      Left            =   2415
      TabIndex        =   2
      Top             =   1155
      Width           =   1950
   End
   Begin VB.CommandButton btFabricante 
      Caption         =   "Fabricantes"
      Height          =   540
      Left            =   315
      TabIndex        =   1
      Top             =   1155
      Width           =   1905
   End
   Begin VB.Label lblControleProdutos 
      AutoSize        =   -1  'True
      Caption         =   "Controle de Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   630
      TabIndex        =   0
      Top             =   315
      Width           =   3435
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btFabricante_Click()
    frmFabricante.Show
End Sub

Private Sub btLocalizaBancoDados_Click()
    CommonDialog.Filter = "Microsoft Access Database (*.mdb) | *.mdb|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Selecione o arquivo"
    CommonDialog.ShowOpen
    PathDatabase = CommonDialog.FileName
End Sub

Private Sub btSmartphone_Click()
    frmSmartphone.Show
End Sub

