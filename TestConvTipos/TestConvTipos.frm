VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtParcelas 
      Height          =   405
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "Calcular"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtData 
      Height          =   405
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblDataPagamento 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   1560
      Width           =   45
   End
   Begin VB.Label txtPagamento 
      Caption         =   "Data do Pagamento:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblParcelas 
      AutoSize        =   -1  'True
      Caption         =   "Digite o N de parcelas"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Digite a data da compra"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalcular_Click()
    Dim DataPagamento As Date
    DataPagamento = CDate(txtData.Text) + txtParcelas
    
    If Weekday(DataPagamento) = 1 Then
        DataPagamento = DataPagamento + 1
    End If
    
    Print Month(DataPagamento)
    
    lblDataPagamento.Caption = DataPagamento
    
End Sub
