VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhatsAppDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IInstantMessaging

Private Const whatsAppPartialPath As String = "\AppData\Local\WhatsApp\WhatsApp.exe"

Private Type TWhatsAppDesktop
    CountryCode         As String
    PhoneNumber         As String
    FullPhoneNumber     As String
    Message             As String
    ExecutableFullPath  As String
End Type

Private this As TWhatsAppDesktop

Private Sub Class_Initialize()
    this.ExecutableFullPath = Environ$("USERPROFILE") & whatsAppPartialPath
    this.CountryCode = "91"
End Sub

Public Property Get IInstantMessaging_CountryCode() As String
    IInstantMessaging_CountryCode = this.CountryCode
End Property

Public Property Let IInstantMessaging_CountryCode(ByVal value As String)
    this.CountryCode = value
    UpdateFullPhoneNumber
End Property

Public Property Get IInstantMessaging_PhoneNumber() As String
    IInstantMessaging_PhoneNumber = this.PhoneNumber
End Property

Public Property Let IInstantMessaging_PhoneNumber(ByVal value As String)
    this.PhoneNumber = value
    UpdateFullPhoneNumber
End Property

Public Property Get IInstantMessaging_FullPhoneNumber() As String
    IInstantMessaging_FullPhoneNumber = this.FullPhoneNumber
End Property

Public Property Get IInstantMessaging_Message() As String
    IInstantMessaging_Message = this.Message
End Property

Public Property Let IInstantMessaging_Message(ByVal value As String)
    this.Message = value
End Property

Public Sub IInstantMessaging_SendMessage()
    SendMessage
End Sub

Private Sub UpdateFullPhoneNumber()
    this.FullPhoneNumber = this.CountryCode & this.PhoneNumber
End Sub

Private Sub SendMessage()
    'Check if whatsapp is installed
    If Not FileExists(this.ExecutableFullPath) Then
        MsgBox "WhatsApp not installed. Download & install from https://www.whatsapp.com/download"
        Exit Sub
    End If
    
    'Build command
    Dim whatsappCommand As String
    whatsappCommand = FormatString("{0} whatsapp://send?phone={1}&text=""{2}""", _
                        this.ExecutableFullPath, _
                        this.FullPhoneNumber, _
                        WorksheetFunction.EncodeURL(this.Message))
    
    'Send message
    Shell whatsappCommand
End Sub


