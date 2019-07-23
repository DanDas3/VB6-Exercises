VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Tela Inicial - Sistema de Vendas"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btSistemaVendas 
      Caption         =   "Sistema de Vendas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   2100
      TabIndex        =   2
      Top             =   2205
      Width           =   3690
   End
   Begin VB.CommandButton btSmartphones 
      Caption         =   "Smartphones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4305
      TabIndex        =   1
      Top             =   315
      Width           =   3690
   End
   Begin VB.CommandButton btFabricantes 
      Caption         =   "Fabricantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   3690
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btFabricantes_Click()
    frmFabricantes.Show
End Sub

Private Sub btSistemaVendas_Click()
    frmSistemaVendas.Show
End Sub

Private Sub btSmartphones_Click()
    frmSmartphone.Show
End Sub

Private Sub Form_Load()
    PathDatabase = "C:\Users\danilo.araujo\Documents\VB6\Project\BD-Smartphone-03\smartphone.mdb"
End Sub
