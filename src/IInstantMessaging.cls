VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IInstantMessaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get CountryCode() As String
End Property

Public Property Let CountryCode(ByVal value As String)
End Property

Public Property Get PhoneNumber() As String
End Property

Public Property Let PhoneNumber(ByVal value As String)
End Property

Public Property Get FullPhoneNumber() As String
End Property

Public Property Get Message() As String
End Property

Public Property Let Message(ByVal value As String)
End Property

Public Sub SendMessage()
End Sub
