Attribute VB_Name = "modRecSetToCombo"
Option Explicit
'modRecsetToCombo
'Coded by Legrev3@aol.com
'Populates a combo box through API calls. Faster than AddItem
'April 24, 2001

'**  Function Declarations:
#If Win32 Then
Private Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Private Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock As Long)

'**  Constant Definitions:
Private Const CB_ADDSTRING& = &H143
Private Const CB_RESETCONTENT& = &H14B
#End If 'WIN32

Public Function RecsetToCombo(hWnd As Long, rsRecSet As ADODB.Recordset, intCol As Integer) As Boolean
    Dim lngRetVal As Long
    
    Call LockWindowUpdate(hWnd)
    On Error GoTo LocalErrHandler:
    If rsRecSet.BOF And rsRecSet.EOF Then
        MsgBox "There are no records to display.", vbInformation + vbOKOnly
        Call LockWindowUpdate(0&)
        Exit Function
    End If

    rsRecSet.MoveFirst
    Call SendMessageBynum(hWnd, CB_RESETCONTENT, 0, 0)
    
    Do Until rsRecSet.EOF
        If Not IsNull(rsRecSet(intCol).Value) Then
            lngRetVal = SendMessageByString(hWnd, CB_ADDSTRING, 0, rsRecSet(intCol).Value)
        End If
        rsRecSet.MoveNext
    Loop
    RecsetToCombo = True
    Call LockWindowUpdate(0&)
    Exit Function
LocalErrHandler:
    MsgBox "Error in filling combo box: " & vbCrLf & _
        Err.Number & "  " & Err.Description, vbCritical + vbOKOnly
    Err.Clear
    RecsetToCombo = False
    Call LockWindowUpdate(0&)
End Function




