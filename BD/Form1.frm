VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btMercadoria 
      Caption         =   "Mercadoria"
      Height          =   330
      Left            =   525
      TabIndex        =   0
      Top             =   840
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btMercadoria_Click()
    frmMercadorias.Show
End Sub
