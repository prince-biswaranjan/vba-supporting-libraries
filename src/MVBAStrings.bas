Attribute VB_Name = "MVBAStrings"

' =================================================================================================
' Module      : MVBAStrings
' Type        : Module
' Description : Module to handle all string operations
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

Private Const MODULE_NAME As String = "MVBAStrings"

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


Public Function FormatString(ByVal inputString As String, ParamArray replacements() As Variant) As String
' =================================================================================================
' Description : Returns a formatted string. Mimics the String.Format behaviour of C#
'
' Parameter : inputString (): Input string with the placeholders
' Parameter : replacements (Variant()): Replacement array for the placeholders

' Return Type : String
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME As String = "FormatString"

    Dim formattedString    As String
    Dim Index As Long
    Dim placeholder As String
    Dim replacement As String

    '----------------------------------------------------------------------------------------------
    
    formattedString = inputString
    
    'Replace all placeholders
    '------------------------
    
    For Index = LBound(replacements) To UBound(replacements)
        
        placeholder = "{" & Index & "}"
        replacement = replacements(Index)
        
        formattedString = Replace(formattedString, placeholder, replacement)
        
    Next Index

    '----------------------------------------------------------------------------------------------
    
    FormatString = formattedString
End Function


Public Function Concat(ByVal Delimiter As String, ParamArray Params() As Variant) As String
' =================================================================================================
' Description : Concatenate string values with a delimiter
'
' Parameter : delimiter (String): Delimiter to be used
' Parameter : params (Variant()): Values to be concatenated

' Return Type : String
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME    As String = "Concat"

    Dim sRtn    As String
    Dim parameter As Variant

    '----------------------------------------------------------------------------------------------
    
    For Each parameter In Params
        
        If Len(sRtn) = 0 Then
            'First item
            '----------
            sRtn = CStr(parameter)
        Else
            'Other Items
            '-----------
            sRtn = FormatString("{0}{1}{2}", sRtn, Delimiter, CStr(parameter))
        End If
        
    Next parameter


    '----------------------------------------------------------------------------------------------


PROC_EXIT:
    
    Concat = sRtn

    Exit Function
    
    '----------------------------------------------------------------------------------------------
    
PROC_ERR:

    Resume PROC_EXIT
    
    
End Function

Public Function RemoveLineBreaks(ByVal inputString As String, Optional ByVal multiLineSeparator As String = vbNullString) As String
' =================================================================================================
' Description : Function to remove new line characters from input string
'
' Parameter : inputString (String): String from which new line characters are to be removed
' Parameter : multiLineSeparator (String): Separator for concating the multiple lines

' Return Type : String
'
' Comments    :
' =================================================================================================

    Const PROCEDURE_NAME    As String = "RemoveLineBreaks"

    Dim sRtn                As String
    Dim newLineCharacters   As Variant
    Dim newLineChar         As Variant

    '----------------------------------------------------------------------------------------------
    
    'List all new line characters
    '----------------------------
    newLineCharacters = Array(vbNewLine, _
                              vbCr, _
                              vbCrLf, _
                              vbLf)
    
    sRtn = inputString
    
    For Each newLineChar In newLineCharacters
        sRtn = Replace(sRtn, newLineChar, multiLineSeparator)
    Next newLineChar
    
    '----------------------------------------------------------------------------------------------

PROC_EXIT:
    
    RemoveLineBreaks = sRtn

    Exit Function
    
    '----------------------------------------------------------------------------------------------
    
PROC_ERR:

    Resume PROC_EXIT
    

End Function
