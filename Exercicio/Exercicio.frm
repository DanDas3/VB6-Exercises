VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btFormNovo 
      Caption         =   "Form Novo"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtNumero2 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtNumero1 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtMsg2 
      Height          =   375
      Left            =   360
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1080
      Width           =   6615
   End
   Begin VB.CommandButton btSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtMsg1 
      Height          =   375
      Left            =   360
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label lblNumero2 
      AutoSize        =   -1  'True
      Caption         =   "Número 2"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   690
   End
   Begin VB.Label lblNumero1 
      AutoSize        =   -1  'True
      Caption         =   "Número 1"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form1
' Author    : danilo.araujo
' Date      : 02/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit
Dim texto1OK As Boolean
Dim texto2OK As Boolean
Dim valor1OK As Boolean
Dim valor2OK As Boolean

'---------------------------------------------------------------------------------------
' Procedure : btOK_Click
' Author    : danilo.araujo
' Date      : 02/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub btOK_Click()
    Dim texto1 As String
    Dim texto2 As String
    Dim valorA As Integer
    Dim valorB As Integer
    texto1 = txtMsg1.Text
    texto2 = txtMsg2.Text
    
    
    If texto1OK = False Or texto2OK = False Or valor1OK = False Or valor2OK = False Then
        MsgBox "Preencha todos os campos", vbOKOnly, "Erro"
    Else
        valorA = CInt(txtNumero1)
        valorB = CInt(txtNumero2)
        MsgBox texto1 & Chr(vbKeyReturn) & texto2 & Chr(vbKeyReturn) & (valorA + valorB), vbOKOnly, "Mensagem"
    End If
End Sub

Private Sub btSair_Click()
    Dim comando As Integer
    
    comando = MsgBox("Você deseja realmente sair", vbYesNo + vbExclamation, "Confirmar saída")
    
    If comando = vbYes Then
        End
    End If
End Sub

Private Sub Form_Load()
    texto1OK = False
    texto2OK = False
    valor1OK = False
    valor2OK = False
End Sub

Private Sub txtMsg1_Change()
    If txtMsg1.Text = "" Then
        texto1OK = False
    Else
        texto1OK = True
    End If
    
    If texto1OK = False Or texto2OK = False Or valor1OK = False Or valor2OK = False Then
        btFormNovo.Visible = False
    Else
        btFormNovo.Visible = True
    End If
End Sub

Private Sub txtMsg2_Change()
    If txtMsg2.Text = "" Then
        texto2OK = False
    Else
        texto2OK = True
    End If
    
    If texto1OK = False Or texto2OK = False Or valor1OK = False Or valor2OK = False Then
        btFormNovo.Visible = False
    Else
        btFormNovo.Visible = True
    End If
End Sub

Private Sub txtNumero1_Change()
    If txtNumero1.Text = "" Then
        valor1OK = False
    Else
        valor1OK = True
    End If

    If texto1OK = False Or texto2OK = False Or valor1OK = False Or valor2OK = False Then
        btFormNovo.Visible = False
    Else
        btFormNovo.Visible = True
    End If
End Sub

Private Sub txtNumero1_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> vbKeyBack) Then
        If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNumero2_Change()
    If txtNumero2.Text = "" Then
        valor2OK = False
    Else
        valor2OK = True
    End If
    
    If texto1OK = False Or texto2OK = False Or valor1OK = False Or valor2OK = False Then
        btFormNovo.Visible = False
    Else
        btFormNovo.Visible = True
    End If
End Sub

Private Sub txtNumero2_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> vbKeyBack) Then
        If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
            KeyAscii = 0
        End If
    End If
End Sub
