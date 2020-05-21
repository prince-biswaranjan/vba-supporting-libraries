Attribute VB_Name = "MMSExcelApp"
' ==========================================================================
' Module      : MMSExcelApp
' Type        : Module
' Description : Support for working with the Application object
' --------------------------------------------------------------------------
' Procedures  : GetApplicationProperties
'               ResetApplicationProperties
'               SetApplicationProperties
' ==========================================================================

' -----------------------------------
' Option statements
' -----------------------------------

Option Explicit
Option Private Module

' -----------------------------------
' Type declarations
' -----------------------------------
' Global Level
' ----------------

Public Type TApplicationProperties
    Calculation     As XlCalculation
    Cursor          As XlMousePointer
    DisplayAlerts   As Boolean
    EnableEvents    As Boolean
    ReferenceStyle  As XlReferenceStyle
    ScreenUpdating  As Boolean
    StatusBar       As Variant
End Type

Public Sub GetApplicationProperties(ByRef Properties _
                                       As TApplicationProperties, _
                           Optional ByVal UseDefaults As Boolean)
' ==========================================================================
' Description : Copy the Application object properties to a structure.
'
' Parameters  : Properties      The structure to populate
'               UseDefaults     Load with default values
' ==========================================================================

    If UseDefaults Then
        With Properties
            .Calculation = xlCalculationAutomatic
            .Cursor = xlDefault
            .DisplayAlerts = True
            .EnableEvents = True
            .ReferenceStyle = xlA1
            .ScreenUpdating = True
            .StatusBar = False
        End With
    Else
        With Properties
            .Calculation = Application.Calculation
            .Cursor = Application.Cursor
            .DisplayAlerts = Application.DisplayAlerts
            .EnableEvents = Application.EnableEvents
            .ReferenceStyle = Application.ReferenceStyle
            .ScreenUpdating = Application.ScreenUpdating
            .StatusBar = Application.StatusBar
        End With
    End If

End Sub

Public Sub ResetApplicationProperties()
' ==========================================================================
' Description : Reset the Application properties back to the default values.
' ==========================================================================

    Dim udtProps As TApplicationProperties

    Call GetApplicationProperties(udtProps, True)
    Call SetApplicationProperties(udtProps)

End Sub

Public Sub SetApplicationProperties(ByRef Properties _
                                          As TApplicationProperties)
' ==========================================================================
' Description : Populate application properties from structure values
'
' Parameters  : Properties  The structure containing the values to use.
' ==========================================================================

    With Application
        .Calculation = Properties.Calculation
        .Cursor = Properties.Cursor
        .DisplayAlerts = Properties.DisplayAlerts
        .EnableEvents = Properties.EnableEvents
        .ReferenceStyle = Properties.ReferenceStyle
        .ScreenUpdating = Properties.ScreenUpdating

        ' Need to pass a boolean, not a variant
        ' -------------------------------------
        If (UCase(CStr(Properties.StatusBar)) = "FALSE") Then
            .StatusBar = False
        Else
            .StatusBar = Properties.StatusBar
        End If
    End With

End Sub
