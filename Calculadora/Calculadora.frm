VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSin 
      Caption         =   "sin"
      Height          =   435
      Left            =   3255
      TabIndex        =   25
      Top             =   1155
      Width           =   540
   End
   Begin VB.Frame fraEspaco 
      Height          =   435
      Left            =   3255
      TabIndex        =   24
      Top             =   630
      Width           =   540
   End
   Begin VB.CommandButton btIgual 
      Caption         =   "="
      Height          =   960
      Left            =   2625
      TabIndex        =   23
      Top             =   2205
      Width           =   540
   End
   Begin VB.CommandButton btInversor 
      Caption         =   "1/x"
      Height          =   435
      Left            =   2625
      TabIndex        =   22
      Top             =   1680
      Width           =   540
   End
   Begin VB.CommandButton btPercentagem 
      Caption         =   "%"
      Height          =   435
      Left            =   2625
      TabIndex        =   21
      Top             =   1155
      Width           =   540
   End
   Begin VB.CommandButton btSomador 
      Caption         =   "+"
      Height          =   435
      Left            =   1995
      TabIndex        =   20
      Top             =   2730
      Width           =   540
   End
   Begin VB.CommandButton btSubtrador 
      Caption         =   "-"
      Height          =   435
      Left            =   1995
      TabIndex        =   19
      Top             =   2205
      Width           =   540
   End
   Begin VB.CommandButton btMultiplicador 
      Caption         =   "*"
      Height          =   435
      Left            =   1995
      TabIndex        =   18
      Top             =   1680
      Width           =   540
   End
   Begin VB.CommandButton btDivisor 
      Caption         =   "/"
      Height          =   435
      Left            =   1995
      TabIndex        =   17
      Top             =   1155
      Width           =   540
   End
   Begin VB.CommandButton btVirgula 
      Caption         =   ","
      Height          =   435
      Left            =   1365
      TabIndex        =   16
      Top             =   2730
      Width           =   540
   End
   Begin VB.CommandButton btZero 
      Caption         =   "0"
      Height          =   435
      Left            =   105
      TabIndex        =   15
      Top             =   2730
      Width           =   1170
   End
   Begin VB.CommandButton btTres 
      Caption         =   "3"
      Height          =   435
      Left            =   1365
      TabIndex        =   14
      Top             =   2205
      Width           =   540
   End
   Begin VB.CommandButton btSeis 
      Caption         =   "6"
      Height          =   435
      Left            =   1365
      TabIndex        =   13
      Top             =   1680
      Width           =   540
   End
   Begin VB.CommandButton btNove 
      Caption         =   "9"
      Height          =   435
      Left            =   1365
      TabIndex        =   12
      Top             =   1155
      Width           =   540
   End
   Begin VB.CommandButton btDois 
      Caption         =   "2"
      Height          =   435
      Left            =   735
      TabIndex        =   11
      Top             =   2205
      Width           =   540
   End
   Begin VB.CommandButton btCinco 
      Caption         =   "5"
      Height          =   435
      Left            =   735
      TabIndex        =   10
      Top             =   1680
      Width           =   540
   End
   Begin VB.CommandButton btOito 
      Caption         =   "8"
      Height          =   435
      Left            =   735
      TabIndex        =   9
      Top             =   1155
      Width           =   540
   End
   Begin VB.CommandButton btUm 
      Caption         =   "1"
      Height          =   435
      Left            =   105
      TabIndex        =   8
      Top             =   2205
      Width           =   540
   End
   Begin VB.CommandButton btQuatro 
      Caption         =   "4"
      Height          =   435
      Left            =   105
      TabIndex        =   7
      Top             =   1680
      Width           =   540
   End
   Begin VB.CommandButton btSete 
      Caption         =   "7"
      Height          =   435
      Left            =   105
      TabIndex        =   6
      Top             =   1155
      Width           =   540
   End
   Begin VB.CommandButton btSqrt 
      Caption         =   "v"
      Height          =   435
      Left            =   2625
      TabIndex        =   5
      Top             =   630
      Width           =   540
   End
   Begin VB.CommandButton btMaisMenos 
      Caption         =   "±"
      Height          =   435
      Left            =   1995
      TabIndex        =   4
      Top             =   630
      Width           =   540
   End
   Begin VB.CommandButton btClear 
      Caption         =   "C"
      Height          =   435
      Left            =   1365
      TabIndex        =   3
      Top             =   630
      Width           =   540
   End
   Begin VB.CommandButton btCE 
      Caption         =   "CE"
      Height          =   435
      Left            =   735
      TabIndex        =   2
      Top             =   630
      Width           =   540
   End
   Begin VB.CommandButton btBackspace 
      Height          =   435
      Left            =   105
      Picture         =   "Calculadora.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   630
      Width           =   540
   End
   Begin VB.TextBox txtDisplay 
      Alignment       =   1  'Right Justify
      Height          =   435
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   105
      Width           =   3060
   End
   Begin VB.Menu mnExibir 
      Caption         =   "Exibir"
      Begin VB.Menu mnClassico 
         Caption         =   "Clássico"
      End
      Begin VB.Menu mnCientifica 
         Caption         =   "Científica"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim data1 As Double
Dim resultado As Double
Dim operador As Integer
Dim igualPressionado As Boolean
Dim heightClassico As Integer
Dim widthClassico As Integer
Dim heightCientifico As Integer
Dim widthCientifico As Integer

Private Function Append(Key As Integer)
    txtDisplay.Text = txtDisplay & Chr(Key)
End Function

Private Function LimpaDisplay()
    txtDisplay = "0"
End Function

Private Function SetOperador(op As Integer)
    operador = op
    data1 = Val(txtDisplay.Text)
    LimpaDisplay
End Function

Private Function SetIgualEvento(pressionado As Boolean)
    igualPressionado = pressionado
End Function

Private Function BotaoNumericoPressionado(Key As Integer)
    If igualPressionado = True Then
        txtDisplay = ""
        igualPressionado = False
        
    End If
    
    If (StrComp(txtDisplay.Text, "0") = 0) Then
        txtDisplay.Text = ""
    End If
    Append Key
    
End Function

Private Sub btDivisor_Click()
    SetOperador vbKeyDivide
End Sub

Private Sub btIgual_Click()

    Select Case operador
    Case vbKeyAdd
        resultado = data1 + Val(txtDisplay.Text)
    Case vbKeyMenu
        resultado = data1 - Val(txtDisplay.Text)
    Case vbKeyMultiply
        resultado = data1 * Val(txtDisplay.Text)
    Case vbKeyDivide
        resultado = data1 / Val(txtDisplay.Text)
    End Select
    
    txtDisplay.Text = resultado
    
    
End Sub

Private Sub btMultiplicador_Click()
    SetOperador vbKeyMultiply
End Sub

Private Sub btSomador_Click()
    operador = vbKeyAdd
    SetOperador operador
End Sub

Private Sub btSubtrador_Click()
    SetOperador vbKeyMenu
End Sub

Private Sub Form_Load()
    LimpaDisplay
    igualPressionado = False
    
    heightClassico = 4155
    widthClassico = 3525
    
    Form1.Height = heightClassico
    Form1.Width = widthClassico
    
    heightCientifico = heightClassico + 100
    widthCientifico = widthClassico + 100
End Sub

Private Sub btBackspace_Click()
txtDisplay.Text = Left(txtDisplay.Text, (Len(txtDisplay) - 1))
End Sub


Private Sub btCinco_Click()
    BotaoNumericoPressionado vbKey5
End Sub

Private Sub btDois_Click()
    BotaoNumericoPressionado vbKey2
End Sub

Private Sub btNove_Click()
    BotaoNumericoPressionado vbKey9
End Sub

Private Sub btOito_Click()
    BotaoNumericoPressionado vbKey8
End Sub

Private Sub btQuatro_Click()
    BotaoNumericoPressionado vbKey4
End Sub

Private Sub btSeis_Click()
    BotaoNumericoPressionado vbKey6
End Sub

Private Sub btSete_Click()
    BotaoNumericoPressionado vbKey7
End Sub

Private Sub btTres_Click()
    BotaoNumericoPressionado vbKey3
End Sub

Private Sub btUm_Click()
    BotaoNumericoPressionado vbKey1
End Sub

Private Sub btZero_Click()
    BotaoNumericoPressionado vbKey0
End Sub

Private Sub mnCientifica_Click()
    Form1.Height = heightCientifico
    Form1.Width = widthCientifico + cmdSin.Width
End Sub

Private Sub mnClassico_Click()
    Form1.Height = heightClassico
    Form1.Width = widthClassico
End Sub
