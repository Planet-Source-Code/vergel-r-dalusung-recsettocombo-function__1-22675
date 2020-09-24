VERSION 5.00
Begin VB.Form frmRecsetToCombo 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   825
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   555
      Width           =   3180
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      Height          =   375
      Left            =   855
      TabIndex        =   3
      Top             =   1260
      Width           =   2985
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   375
      Left            =   810
      TabIndex        =   2
      Top             =   2340
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   2985
   End
End
Attribute VB_Name = "frmRecsetToCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cnConn As ADODB.Connection
Private rsRecSet As ADODB.Recordset

Private Sub Form_Load()
    Label1.Caption = "Query Start: " & Time()
    Dim strConnect As String
    Dim blnRetVal As Boolean
    
    'replace with your own provider, database path and filename and password
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\TestData\vbt001.mdb;Jet OLEDB:Database Password = allow;"
    Set cnConn = New ADODB.Connection
    Set rsRecSet = New ADODB.Recordset
    
    cnConn.Open strConnect
    rsRecSet.CursorLocation = adUseClient
    'replace with your own table name
    rsRecSet.Open "SELECT * FROM ARCustomers", cnConn, adOpenKeyset, adLockOptimistic, adCmdText
    
    'usage: RecsetToCombo(ComboBox.hWnd, recordset object, column to display)
    Label3.Caption = "Using Windows API Calls"
    blnRetVal = RecsetToCombo(Combo1.hWnd, rsRecSet, 0)
    
'    Call UseAddItem   'uncomment this function and comment out RecsetToCombo
                       'to compare processing time

    Combo1.ListIndex = 0
    Label2.Caption = "Query End : " & Time()
End Sub

Private Sub UseAddItem(Optional ByVal intCol As Integer = 0)
    'use this to compare processing time with RecsetToCombo
    Label3.Caption = "Using AddItem Method"
    If rsRecSet.BOF And rsRecSet.EOF Then Exit Sub
    rsRecSet.MoveFirst

    Do Until rsRecSet.EOF
        If Not IsNull(rsRecSet(intCol).Value) Then
            Combo1.AddItem rsRecSet(intCol).Value
        End If
        rsRecSet.MoveNext
    Loop
End Sub


