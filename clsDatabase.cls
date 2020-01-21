Public Database As String
Public password As String

'Declare Variable
Private Conn As New Connection
Private Rec As Recordset
Public Sub Connect()
   
   'if database not found then create new database
1
    If Dir(Database) = vbNullString Then
        On Error GoTo 1
       Call CreateDatabase
        Exit Sub
    End If
   
   'Create Connection with Access Database
  
   If Conn.State = 0 Then
        Conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Database & _
        ";Jet OLEDB:Database Password=" & password
        Conn.Open
   End If

   
End Sub
Public Sub Disconnect()
    If Conn.State = 1 Then
        Conn.Close
    End If
End Sub
Private Sub CreateDatabase()
    Dim obj As Access.Application
    Set obj = New Access.Application
    obj.DBEngine.CreateDatabase Database, DB_LANG_GENERAL
    'create password db
    
    Set mydb = obj.DBEngine.OpenDatabase(Database, True)
    mydb.newpassword "", password
    obj.Quit
    
    Call Connect
End Sub

Public Sub Execute(Query As String, Optional Prompt As Boolean = False)
    On Error Resume Next
    Conn.Execute (Query)
    
    If Err.Number <> 0 And Prompt Then
        MsgBox Err.Description
    End If
    
End Sub

Public Function Getdata(Query As String) As Recordset
    Set Rec = New Recordset
    Rec.Open Query, Conn, adOpenKeyset, adLockOptimistic
    Set Getdata = Rec
End Function

Public Function ShowData(TableName As String, ListBox As MSForms.ListBox, Optional Filter As String = "") As String
    Getdata (TableName)
    
    If Rec Is Nothing Then
        MsgBox "Table not found", vbCritical, "Alert"
        Exit Function
    End If
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(Rec.Source)

    If Err.Number <> 0 Then
        Err.Clear
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = Rec.Source
    End If
    
    'clearcontent
    ws.Range("a:z").ClearContents
        
    If Not Filter = "" Then Rec.Filter = Filter
    Err.Clear
    
    'create header
    For i = 0 To Rec.Fields.Count - 1
        ws.Cells(2, i + 1).value = Rec.Fields(i).Name
    Next
    
    'populate record
    ws.Cells(3, 1).CopyFromRecordset Rec
    
    'showing data in listbox
    ws.Activate
    Dim iRow, iCol As Integer
    iRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(0, 0).Row
    iCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Offset(0, 0).Column
    
    If iRow < 3 Then iRow = 3
    
    ShowData = "a3:" & Chr(64 + iCol) & iRow
    
    With ListBox
        .RowSource = ShowData
        .ColumnHeads = True
        .ColumnCount = iCol
    End With
    
End Function

Public Function AddItem(TableName As String) As Recordset
    Getdata TableName
    Set AddItem = Rec
    Rec.AddNew
End Function

Public Function EditItem(TableName As String, Filter As String) As Recordset
    Getdata "select * from " & TableName & " where " & Filter
    Set EditItem = Rec
End Function

Public Function DeleteItem(TableName As String, Filter As String) As Recordset
    Getdata "delete from " & TableName & " where " & Filter
    Set DeleteItem = Rec
End Function

