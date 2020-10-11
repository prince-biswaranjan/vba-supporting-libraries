Attribute VB_Name = "MVBALogger"

' =================================================================================================
' Module      : MVBALogger
' Type        : Module
' Description : Module to write logs
' -------------------------------------------------------------------------------------------------
' Properties  : LogFilePath         :   The path where the log would be written to
'               IsEnabled           :   If logging is enabled or not
'               GlobalLoggingLevel  :   Set the global logging level. Anything logged at a higher
'                                       level would not be logged
' -------------------------------------------------------------------------------------------------
' Procedures  : GetLogFilePath          Return Type :   String
'               WriteToLog              Return Type :   Void
'               GetLoglevelDescription  Return Type :   String
' -------------------------------------------------------------------------------------------------
' Events      : XXX
' -------------------------------------------------------------------------------------------------
' Dependencies: MVBAFileSystem
' -------------------------------------------------------------------------------------------------
' References  : XXX
' -------------------------------------------------------------------------------------------------
' Comments    :
' =================================================================================================

' -----------------------------------------------
' Option statements
' -----------------------------------------------

Option Explicit
Option Private Module

' -----------------------------------------------
' Interface declarations
' -----------------------------------------------

'Implements IUnknown

' -----------------------------------------------
' Constant declarations
' -----------------------------------------------
' Global Level
' ----------------------

'Public Const GLOBAL_CONST As String = ""

' ----------------------
' Module Level
' ----------------------

Private Const MODULE_NAME           As String = "MVBALogger"
Private Const MULTI_LINE_SEPARATOR  As String = "\n"

' -----------------------------------------------
' Enumeration declarations
' -----------------------------------------------
' Global Level
' ----------------------

'Public Enum enuGlobal
'    enuGItem = 0
'End Enum

Public Enum LoggingLevel
    Fatal
    Warning
    Information
    Verbose
End Enum

' ----------------------
' Module Level
' ----------------------

'Private Enum enuMod
'    enuMItem = 0
'End Enum


' -----------------------------------------------
' Type declarations
' -----------------------------------------------
' Global Level
' ----------------------

'Public Type TPublic
'    PublicID    As Integer
'End Type

' ----------------------
' Module Level
' ----------------------

'Private Type TPrivate
'    PrivateID   As Integer
'End Type

Private Type TLogger
    LogFilePath As String
    IsEnabled As Boolean
    GlobalLoggingLevel As LoggingLevel
End Type

' -----------------------------------------------
' Event declarations
' -----------------------------------------------

'[Public] Event EventName(ByVal Arg As String)

' -----------------------------------------------
' External Function declarations
' -----------------------------------------------

'#If VBA7 Then
'    Private Declare PtrSafe _
'            Function <FunctionName> _
'            Lib "user32" _
'            Alias "" (ByVal hWnd As LongPtr) As LongPtr
'#Else
'    Private Declare _
'            Function <FunctionName> _
'            Lib "user32" _
'            Alias "" (ByVal hWnd As Long) As Long
'#End If

' -----------------------------------------------
' Variable declarations
' -----------------------------------------------
' Global Level
' ----------------------

'Public gsVar    As String

' ----------------------
' Module Level
' ----------------------

'Private msVar   As String

Private this As TLogger


Public Property Let LogFilePath(ByVal value As String)
    this.LogFilePath = value
End Property

Public Property Get LogFilePath() As String
    LogFilePath = this.LogFilePath
End Property

Public Property Let IsEnabled(ByVal value As Boolean)
    this.IsEnabled = value
End Property

Public Property Get IsEnabled() As Boolean
    IsEnabled = this.IsEnabled
End Property

Public Property Let GlobalLoggingLevel(ByVal value As LoggingLevel)
    this.GlobalLoggingLevel = value
End Property

Public Property Get GlobalLoggingLevel() As LoggingLevel
    GlobalLoggingLevel = this.GlobalLoggingLevel
End Property

Private Function GetLogFilePath() As String
' =================================================================================================
' Description : Function to return the log file path
'

' Return Type : String
'
' Comments    : Returns set file path or a temp folder path if no log file path is set
' =================================================================================================

    Const PROCEDURE_NAME    As String = "GetLogFilePath"

    Dim tempFolderPath      As String
    Dim rtnLogFilePath      As String

    '----------------------------------------------------------------------------------------------
    
    If this.LogFilePath = vbNullString Then
        tempFolderPath = Environ$("Temp")

        rtnLogFilePath = FormatString("{0}\{1}_{2}.log", tempFolderPath, ThisWorkbook.Name, Format$(Now, "yyyymmdd"))
    Else
        rtnLogFilePath = FormatString("{0}\{1}_{2}.log", this.LogFilePath, ThisWorkbook.Name, Format$(Now, "yyyymmdd"))
    End If
    
    '----------------------------------------------------------------------------------------------
    
    GetLogFilePath = rtnLogFilePath
End Function


Public Sub WriteToLog(ByVal module As String, _
                        ByVal procedure As String, _
                        ByVal Message As String, _
                Optional ByVal logLevel As LoggingLevel = Information, _
                Optional ByVal errNumber As Long, _
                Optional ByVal Source As String)
' =================================================================================================
' Description : Procedure to write to logs
'
' Parameter : module (String): Name of the calling module
' Parameter : procedure (String): Name of the calling procedure
' Parameter : message (String): Message to be written
' Parameter : logLevel (LoggingLevel): Logging level of the message
' Parameter : errNumber (Long): Error number (Optional)
' Parameter : source (String): Error Source (Optional)
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME    As String = "WriteToLog"

    Dim computerName        As String
    Dim userName            As String
    Dim iFileNum            As Integer
    Dim logFile             As String
    Dim logFileHeader       As String
    Dim logMessage          As String
    Dim logEntry            As String

    '----------------------------------------------------------------------------------------------
    
    'Exit if logging not required
    '----------------------------
    If (Not this.IsEnabled Or this.GlobalLoggingLevel < logLevel) Then
        Exit Sub
    End If
    
    'Variable(s) Initialization
    '--------------------------
    logFile = GetLogFilePath()
    computerName = Environ$("COMPUTERNAME")
    userName = Environ$("USERNAME")
    
    
    'If the log doesn't exist create
    'an empty file with header row
    '-------------------------------
    iFileNum = FreeFile()
    
    If Not FileExists(logFile) Then
        
        logFileHeader = Concat(vbTab, _
                               "Date", _
                               "Time", _
                               "Computer", _
                               "User", _
                               "LogLevel", _
                               "Source", _
                               "Module", _
                               "Procedure", _
                               "Message", _
                               "ErrorNumber")
        
        Open logFile For Append As #iFileNum
        Print #iFileNum, logFileHeader
    Else
        Open logFile For Append As #iFileNum
    End If
    
    'Remove new line characters
    '--------------------------
    logMessage = RemoveLineBreaks(Message, MULTI_LINE_SEPARATOR)
    
    logEntry = Concat(vbTab, _
                      Format$(Now(), "dd/mmm/yyyy"), _
                      Format$(Now(), "HH:MM:SS"), _
                      computerName, _
                      userName, _
                      GetLoglevelDescription(logLevel), _
                      IIf(Len(Source) > 0, Source, ThisWorkbook.Name), _
                      module, _
                      procedure, _
                      logMessage, _
                      IIf(errNumber, errNumber, vbNullString))
    
    Print #iFileNum, logEntry
    
    Close #iFileNum
    
    '----------------------------------------------------------------------------------------------
    
PROC_EXIT:
    
    Exit Sub
    
    '----------------------------------------------------------------------------------------------
    
PROC_ERR:
    
    Resume PROC_EXIT

End Sub


Private Function GetLoglevelDescription(ByVal level As LoggingLevel) As String
' =================================================================================================
' Description : Converts logging level enum to userfriendly strings
'
' Parameter : level (LoggingLevel): logging level enum value

' Return Type : String
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME    As String = "GetLoglevelDescription"

    Dim sRtn                As String


    '----------------------------------------------------------------------------------------------
    
    Select Case level
        Case Fatal
            sRtn = "[FATAL]"
        Case Warning
            sRtn = "[WARN]"
        Case Information
            sRtn = "[INFO]"
        Case Verbose
            sRtn = "[DEBUG]"
        Case Else
            sRtn = "[Unknown]"
    End Select


    '----------------------------------------------------------------------------------------------

PROC_EXIT:
    
    GetLoglevelDescription = sRtn

    Exit Function
    
    '----------------------------------------------------------------------------------------------
    
PROC_ERR:

    Resume PROC_EXIT
    
    
End Function





