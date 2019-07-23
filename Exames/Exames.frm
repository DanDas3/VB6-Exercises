VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8925
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraExame 
      Caption         =   "Exame"
      Height          =   855
      Left            =   4680
      TabIndex        =   9
      Top             =   1200
      Width           =   3735
      Begin VB.CheckBox chkEspermograma 
         Caption         =   "Espermograma"
         Height          =   195
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkSangue 
         Caption         =   "Sangue"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox chkGravidez 
         Caption         =   "Gravidez"
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkRaioX 
         Caption         =   "Raio X"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fraSexo 
      Caption         =   "Sexo"
      Enabled         =   0   'False
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   3855
      Begin VB.OptionButton optFeminino 
         Caption         =   "Feminino"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optMasculino 
         Caption         =   "Masculino"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2520
      Width           =   7935
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtNome 
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      MaxLength       =   30
      TabIndex        =   0
      Top             =   480
      Width           =   7935
   End
   Begin VB.Label lblRelatorio 
      Caption         =   "Relatório"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblNomePaciente 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancelar_Click()
    End
End Sub

Private Sub optMasculino_Click()
    chkEspermograma.Enabled = True
    chkGravidez.Enabled = False
    chkRaioX.Enabled = True
    chkSangue.Enabled = True
    
End Sub

Private Sub optFeminino_Click()
    chkEspermograma.Enabled = False
    chkGravidez.Enabled = True
    chkRaioX.Enabled = True
    chkSangue.Enabled = True
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub txtNome_Change()
    If txtNome.Text = "" Then
        fraSexo.Enabled = False
        
    Else
        fraSexo.Enabled = True
        optMasculino.Value = True
        
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
        KeyAscii = 0
    End If
    
End Sub
