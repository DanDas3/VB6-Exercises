VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNomeConjugue 
      Height          =   405
      Left            =   360
      TabIndex        =   14
      Top             =   5640
      Width           =   4695
   End
   Begin VB.Frame fraSetor 
      Caption         =   "Setor"
      Height          =   1335
      Left            =   3360
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
      Begin VB.OptionButton optAdministrativo 
         Caption         =   "Administrativo"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optIndustrial 
         Caption         =   "Industrial"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraEstadoCivil 
      Caption         =   "Estado Civil"
      Height          =   1815
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   2295
      Begin VB.OptionButton optEstadoCivil 
         Caption         =   "Viúvo"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optEstadoCivil 
         Caption         =   "Solteiro"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optEstadoCivil 
         Caption         =   "Casado"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtCargo 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox txtNomeFuncionario 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblNomeConjugue 
      Caption         =   "Nome do Conjugue"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label lblCargo 
      AutoSize        =   -1  'True
      Caption         =   "Cargo"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label lblNomeFuncionario 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Funcionário"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optEstadoCivil_Click(Index As Integer)
    lblNomeConjugue.Enabled = optEstadoCivil(0).Value
    txtNomeConjugue.Enabled = optEstadoCivil(0).Value
    
    If Index = 0 Then
        txtNomeConjugue.SetFocus
    End If
End Sub
