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

Private Const MODULE_NAME   As String = "MVBADatabase"

Private Type T_MVBADatabase
    ConnectionString As String
    connection As ADODB.connection
End Type

Private this As T_MVBADatabase

Private Sub Class_Initialize()
    this.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
End Sub

Public Property Get ConnectionString() As String
    ConnectionString = this.ConnectionString
End Property

Public Property Let ConnectionString(ByVal value As String)
    this.ConnectionString = value
End Property

Private Sub OpenConnection()
    If this.connection Is Nothing Then
        Set this.connection = New ADODB.connection
        this.connection.ConnectionString = this.ConnectionString
    End If
    
    If this.connection.State <> 1 Then
        this.connection.Open
    End If
    'Debug.Print "Connection Opened"
End Sub

Public Function GetRecordSet(ByVal sqlQuery As String) As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    OpenConnection
    Set rs = New ADODB.Recordset
    
    'Debug.Print sqlQuery
    rs.Open sqlQuery, this.connection
    
    'Debug.Print rs.GetString
    
    Set GetRecordSet = rs
    
End Function

Public Sub ExecuteQuery(ByVal sqlQuery As String)
    Const PROC_NAME As String = "ExecuteQuery"
    OpenConnection
    
    'Debug.Print sqlQuery
    WriteToLog MODULE_NAME, PROC_NAME, "Executing Query: " & sqlQuery
    
    With New ADODB.Command
        Set .ActiveConnection = this.connection
        .CommandText = sqlQuery
        .Execute
    End With
    
End Sub
Private Sub Class_Terminate()
    If this.connection.State Then
        this.connection.Close
    End If
    
    If Not this.connection Is Nothing Then
        Set this.connection = Nothing
    End If
    
    'Debug.Print "Connection closed"
End Sub
