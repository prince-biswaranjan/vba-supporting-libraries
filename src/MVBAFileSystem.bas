Attribute VB_Name = "MVBAFileSystem"

' =================================================================================================
' Module      : MVBAFileSystem
' Type        : Module
' Description : Module to handle all file system interactions
' -------------------------------------------------------------------------------------------------
' Properties  : XXX
' -------------------------------------------------------------------------------------------------
' Procedures  : XXX
' -------------------------------------------------------------------------------------------------
' Events      : XXX
' -------------------------------------------------------------------------------------------------
' Dependencies: XXX
' -------------------------------------------------------------------------------------------------
' References  : XXX
' -------------------------------------------------------------------------------------------------
' Comments    :
' =================================================================================================

' -----------------------------------------------
' Option statements
' -----------------------------------------------

'Option Base \{0 | 1}
'Option Compare \{Binary | Text | Database} ' Microsoft Access only
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

Private Const MODULE_NAME As String = "MVBAFileSystem"

' -----------------------------------------------
' Enumeration declarations
' -----------------------------------------------
' Global Level
' ----------------------

'Public Enum enuGlobal
'    enuGItem = 0
'End Enum

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


Public Function FileExists(ByVal filePath As String) As Boolean
' =================================================================================================
' Description : Checks if the file exists at the given path
'
' Parameter : filePath (String): Path of the file whose existence is to be checked

' Return Type : Boolean
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME    As String = "FileExists"

    Dim bRtn    As Boolean
    Dim fso As Scripting.FileSystemObject

    '----------------------------------------------------------------------------------------------
    
    Set fso = New Scripting.FileSystemObject
    
    bRtn = fso.FileExists(filePath)

    '----------------------------------------------------------------------------------------------

PROC_EXIT:
    
    FileExists = bRtn
    Set fso = Nothing
    Exit Function
    
    '----------------------------------------------------------------------------------------------
    
PROC_ERR:
    Resume PROC_EXIT
    
End Function
