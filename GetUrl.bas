Attribute VB_Name = "GetUrl"
Option Explicit

Declare Function shellexecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Const WM_USER = &H400
Const EM_LIMITTEXT = WM_USER + 21
 Const WM_GETTEXT = &HD
 Const WM_GETTEXTLENGTH = &HE
 Const EM_GETLINECOUNT = &HBA
 Const EM_LINEINDEX = &HBB
 Const EM_LINELENGTH = &HC1
 'Used to determine what OS Version
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Const WINNT As Integer = 2
Public Const WIN98 As Integer = 1
Public Function getVersion() As Integer
  Dim udtOSInfo As OSVERSIONINFO
  Dim intRetVal As Integer
         
  'Initialize the type's buffer sizes
    With udtOSInfo
        .dwOSVersionInfoSize = 148
        .szCSDVersion = Space$(128)
    End With
    
  'Make an API Call to Retrieve the OSVersion info
    intRetVal = GetVersionExA(udtOSInfo)
  
  'Set the return value
    getVersion = udtOSInfo.dwPlatformId
End Function
Public Function GetTheURL() As String

 On Error GoTo CallErrorA
    Dim iPos As Integer
    Dim sClassName As String
    Dim GetAddressText As String
    Dim lhwnd As Long
    Dim WindowCaption As Long

   lhwnd = 0
   sClassName = ("IEFrame") ' (Dialog)")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)

         Select Case getVersion()
        Case WIN98 'Windows 95/98
            sClassName = ("WorkerA")
        Case WINNT 'Windows NT
            sClassName = ("WorkerW")
    End Select
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("ReBarWindow32")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("ComboBoxEx32")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("ComboBox")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
   sClassName = ("Edit")
   lhwnd = FindWindowEx(lhwnd, 0, sClassName, vbNullString)
        
        WindowCaption& = lhwnd
        Dim buffer As String, TextLength As Long
        TextLength& = SendMessage(WindowCaption&, WM_GETTEXTLENGTH, 0&, 0&)
        buffer$ = String(TextLength&, 0&)
        Call SendMessageByString(WindowCaption&, WM_GETTEXT, TextLength& + 1, buffer$)
        GetTheURL = GetJustURL(buffer$)

   Exit Function
CallErrorA:
    MsgBox Err.Description
    Err.Clear

End Function

Private Function GetJustURL(DataX As String) As String
    'Parse the request and get the website address
Dim I As Long
Dim y As Long
   For I = 1 To Len(DataX)
        If Mid(DataX, I, 2) = "//" Then 'The website starts after the //(HTTP://)
            Exit For
        End If
    Next
    
    For y = I + 2 To Len(DataX)
        If Mid(DataX, y, 1) = "/" Then 'The website ends at the /(www.planet-source-code.com/) 'In the case of www.planet-source-code.com/Junk/Cheese.html  , the /junk/Cheese.html is part of the request, not part of the website URL, so we connect to the website URL and send the request.
            GetJustURL = Mid(DataX, I + 2, y - (I + 2))
            Exit For
        End If
    Next
End Function

