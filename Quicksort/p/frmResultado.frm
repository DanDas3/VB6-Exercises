VERSION 5.00
Begin VB.Form frmResultado 
   Caption         =   "Resultado"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   ScaleHeight     =   5145
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   1785
      TabIndex        =   1
      Top             =   4515
      Width           =   1590
   End
   Begin VB.TextBox txtResultado 
      Height          =   4215
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   105
      Width           =   5055
   End
End
Attribute VB_Name = "frmResultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function troca(vetor() As Integer, A As Integer, B As Integer)
    Dim Aux As Integer
    
    Aux = vetor(A)
    vetor(A) = vetor(B)
    vetor(B) = Aux
End Function

Function particiona(vetor() As Integer, Inicio As Integer, Fim As Integer) As Integer
    Dim Pivot As Integer
    Dim i As Integer
    Dim j As Integer
    
    Pivot = 0
    i = Inicio + 1
    j = Fim
    Do While (i <= j)
        Do While (vetor(i) <= vetor(Pivot) And i <= Fim)
            i = i + 1
        Loop
        
        Do While (vetor(j) > vetor(Pivot))
            j = j - 1
        Loop
        
        troca vetor, i, j
        If i <= j Then
            i = i + 1
            j = j - 1
        End If
    Loop
    
    troca vetor, Pivot, j
    
    particiona = j
End Function

Function Quick(vetor() As Integer, Inicio As Integer, Fim As Integer)
    Dim meio As Integer
    
    If (Inicio < Fim) Then
        meio = particiona(vetor, Inicio, Fim)
        
        Quick vetor, Inicio, meio - 1
        Quick vetor, meio + 1, Fim
    End If
End Function

Private Sub btOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim index As Integer
    
    index = 0
    Quick vetor, 0, qtdElementos - 1
    
    Do While (index < qtdElementos)
        txtResultado.Text = txtResultado.Text & Str(vetor(index)) & vbCrLf
        
        index = index + 1
    Loop
End Sub
