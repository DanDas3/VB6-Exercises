VERSION 5.00
Begin VB.Form frmSmartphone 
   Caption         =   "smartphone"
   ClientHeight    =   2640
   ClientLeft      =   1170
   ClientTop       =   465
   ClientWidth     =   5400
   LinkTopic       =   "Form2"
   ScaleHeight     =   2640
   ScaleWidth      =   5400
   Begin VB.ListBox lstFabricantes 
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   680
      Width           =   2850
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4440
      TabIndex        =   15
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   3360
      TabIndex        =   14
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   2280
      TabIndex        =   13
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   1200
      TabIndex        =   12
      Top             =   1980
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   1980
      Width           =   975
   End
   Begin VB.Data dataSmartphone 
      Align           =   2  'Align Bottom
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Users\danilo.araujo\Documents\VB6\Project\BD\smartphone.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "smartphone"
      Top             =   2295
      Width           =   5400
   End
   Begin VB.TextBox txtFields 
      DataField       =   "vendidos"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   10
      Top             =   1640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "estoque"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "preco"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   6
      Top             =   1000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "nome"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "codigo"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   40
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "vendidos:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   1660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "estoque:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "preco:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "fabricante:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "nome:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "codigo:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmSmartphone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FabricanteDB As Database
Dim TBFabricante As Recordset

Private Function preencheLista()
    TBFabricante.MoveFirst
    Do While TBFabricante.EOF = False
        lstFabricantes.AddItem TBFabricante("nome")
        TBFabricante.MoveNext
    Loop
End Function

Private Sub cmdAdd_Click()
  dataSmartphone.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  'this may produce an error if you delete the last
  'record or the only record in the recordset
  dataSmartphone.Recordset.Delete
  dataSmartphone.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  'this is really only needed for multi user apps
  dataSmartphone.Refresh
End Sub

Private Sub cmdUpdate_Click()
  dataSmartphone.UpdateRecord
  dataSmartphone.Recordset.Bookmark = dataSmartphone.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub dataSmartphone_Error(DataErr As Integer, Response As Integer)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Error$(DataErr)
  Response = 0  'throw away the error
End Sub

Private Sub dataSmartphone_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position
  'for dynasets and snapshots
  dataSmartphone.Caption = "Record: " & (dataSmartphone.Recordset.AbsolutePosition + 1)
  'for the table object you must set the index property when
  'the recordset gets created and use the following line
  'dataSmartphone.Caption = "Record: " & (dataSmartphone.Recordset.RecordCount * (dataSmartphone.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub dataSmartphone_Validate(Action As Integer, Save As Integer)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Load()
    Set FabricanteDB = OpenDatabase(PathDatabase)
    Set TBFabricante = FabricanteDB.OpenRecordset("fabricante", dbOpenTable)
    
    preencheLista
    
    TBFabricante.Close
    FabricanteDB.Close
End Sub
