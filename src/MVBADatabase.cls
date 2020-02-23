VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MVBADatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private conString           As String
Private conn                As ADODB.Connection
Private Const MODULE_NAME   As String = "MVBADatabase"

Private Sub Class_Initialize()
    conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
End Sub

Private Sub OpenConnection()
    If conn Is Nothing Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = conString
    End If
    
    If conn.State <> 1 Then
        conn.Open
    End If
    'Debug.Print "Connection Opened"
End Sub

Public Function GetRecordSet(ByVal sqlQuery As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    OpenConnection
    Set rs = New ADODB.Recordset
    
    'Debug.Print sqlQuery
    rs.Open sqlQuery, conn
    
    'Debug.Print rs.GetString
    
    Set GetRecordSet = rs
    
End Function

Public Sub ExecuteQuery(ByVal sqlQuery As String)
    Const PROC_NAME As String = "ExecuteQuery"
    OpenConnection
    
    'Debug.Print sqlQuery
    WriteToLog MODULE_NAME, PROC_NAME, "Executing Query: " & sqlQuery
    
    With New ADODB.Command
        Set .ActiveConnection = conn
        .CommandText = sqlQuery
        .Execute
    End With
    
End Sub
Private Sub Class_Terminate()
    If conn.State Then
        conn.Close
    End If
    
    If Not conn Is Nothing Then
        Set conn = Nothing
    End If
    
    'Debug.Print "Connection closed"
End Sub
