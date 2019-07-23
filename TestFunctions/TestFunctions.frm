VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   9990
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Dim Var As String
    Var = "    Danilo de Araújo Silva    "
    Print InStr(1, UCase(Var), "SI", 0)
    Print String(30, "Santa série D")
    Print Len(Var)
    Print Len(Trim(Var))
    
    Dim Var2 As Variant
    Var2 = Array("Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sab")
    Print Var2(Weekday(Date) - 1)
    Dim A As Integer
    Dim B As Integer
    
    A = 10
    B = 10
    
    Print IIf(A > B, A & " é maior do que " & B, A & " é menor ou igual a " & B)
    
    Print "Este programa está sendo executado na pasta: " & CurDir
    Var = CurDir & "\VB6.exe"
    Print FileLen(Var) / 1024 / 1024
    
    Print "Tipo da variável Var é: " & TypeName(Var)
    InputBox "Digite algo", "Teste"
    Print Var
    Print MsgBox("Teste", vbYesNoCancel + vbExclamation, "Teste")
    
    Print "Telefone: " & Format(8182165922#, "(##) 9####-####")
    Print Format(123456, "Currency")
    Print Format(1, "Yes/No")
    Print Format("26/06/19", "c")
End Sub

