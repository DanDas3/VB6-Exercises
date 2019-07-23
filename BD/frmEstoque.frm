VERSION 5.00
Begin VB.Form frmEstoque 
   Caption         =   "Controle de Estoque"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuMercadoria 
         Caption         =   "&Mercadoria"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuCliente 
         Caption         =   "&Cliente"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuLancamento 
      Caption         =   "&Lançamento"
      Begin VB.Menu mnuVenda 
         Caption         =   "Vendas"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "frmEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BancoDeDados As Database
Dim TBMercadoria As Recordset


Private Sub mnuCliente_Click()
    frmCadClientes.Show
End Sub

Private Sub mnuMercadoria_Click()
    frmCadMercadoria.Show
    
End Sub

Private Sub mnuSair_Click()
    Unload frmCadMercadoria
    Unload frmCadClientes
    Unload frmCadVendas
    Unload Me
End Sub

Private Sub mnuVenda_Click()
    frmCadVendas.Show
    
End Sub
