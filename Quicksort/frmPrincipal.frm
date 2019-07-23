VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3360
      Top             =   1785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btAbrirArquivo 
      Caption         =   "Abrir"
      Height          =   540
      Left            =   2625
      TabIndex        =   1
      Top             =   1155
      Width           =   1695
   End
   Begin VB.CommandButton btOrdenar 
      Caption         =   "Ordenar"
      Height          =   540
      Left            =   420
      TabIndex        =   0
      Top             =   1155
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function FileToString(strFilename As String) As String
    Dim line As String
    
    Dim index As Integer
    Dim iFile As Integer
    
    index = 0
    iFile = FreeFile
    
    Open strFilename For Input As #iFile
      'FileToString = StrConv(InputB(LOF(iFile), iFile), vbUnicode)
      Line Input #iFile, line
      qtdElementos = line
      
      ReDim vetor(qtdElementos)
      
      Do Until EOF(iFile)
        Line Input #iFile, line
        vetor(index) = line
        index = index + 1
      Loop
    Close #iFile
    MsgBox "Operação concluída com sucesso!"
End Function

Private Sub btAbrirArquivo_Click()
    CommonDialog.Filter = "Text Files (*.txt) | *.txt|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Selecione o arquivo"
    CommonDialog.ShowOpen
    
    FileToString CommonDialog.FileName
End Sub

Private Sub btOrdenar_Click()
    frmResultado.Show
    
End Sub

