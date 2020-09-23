VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Google Power 2"
   ClientHeight    =   1125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option4 
      Caption         =   "Current Website"
      Height          =   195
      Left            =   5325
      TabIndex        =   7
      Top             =   113
      Width           =   1590
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Patent Search"
      Height          =   210
      Left            =   3840
      TabIndex        =   6
      Top             =   105
      Width           =   1365
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Website References"
      Height          =   210
      Left            =   1950
      TabIndex        =   5
      Top             =   105
      Width           =   1845
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   2475
      Width           =   345
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   592
      Top             =   1725
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Images"
      Height          =   210
      Left            =   1020
      TabIndex        =   4
      Top             =   105
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Web"
      Height          =   210
      Left            =   60
      TabIndex        =   3
      Top             =   105
      Width           =   735
   End
   Begin VB.TextBox Capt1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1590
      TabIndex        =   2
      Text            =   "Search Criteria Here"
      Top             =   420
      Width           =   3675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search Now"
      Height          =   255
      Left            =   2535
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   750
      Width           =   1815
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long


Dim GPreface As String
Dim GDir As String
Dim SavText As String
Private Sub Command1_Click()
On Error Resume Next
Text2.SetFocus
Me.WindowState = 1
If Option4.Value = True Then
Shell "C:\Program Files\Internet Explorer\IEXPLORE " & GDir & "?&q=" & Capt1.Text & " " & GPreface & GetTheURL, vbMaximizedFocus
Else
Shell "C:\Program Files\Internet Explorer\IEXPLORE " & GDir & "?&q=" & GPreface & Capt1.Text, vbMaximizedFocus
End If
Capt1.Text = SavText
End Sub
Private Sub Form_Load()
If App.PrevInstance Then End
Call FormOnTop(Me.hWnd, True)
btnFlat Command1
Option1 = True
End Sub
Private Sub Option1_Click()
On Error Resume Next
Text2.SetFocus
GDir = "www.google.com/search"
GPreface = ""
Capt1.Text = "Search Criteria Here"
SavText = "Search Criteria Here"
End Sub
Private Sub Option2_Click()
On Error Resume Next
Text2.SetFocus
GDir = "www.google.com/images"
GPreface = ""
Capt1.Text = "Search Criteria Here"
SavText = "Search Criteria Here"
End Sub
Private Sub Option3_Click()
On Error Resume Next
Text2.SetFocus
GDir = "www.google.com/search"
GPreface = "link:"
Capt1.Text = "Type The URL Here"
SavText = "Type The URL Here"
End Sub
Private Sub Option4_Click()
On Error Resume Next
Text2.SetFocus
GDir = "www.google.com/search"
GPreface = "site:"
Capt1.Text = "Search Criteria Here"
SavText = "Search Criteria Here"
End Sub
Private Sub Option5_Click()
On Error Resume Next
Text2.SetFocus
GDir = "www.google.com/search"
GPreface = "patent "
Capt1.Text = "Patent Number Here"
SavText = "Patent Number Here"
End Sub
Function btnFlat(Button As CommandButton)
SetWindowLong Button.hWnd, GWL_STYLE, WS_CHILD Or BS_FLAT
Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
Private Sub Timer1_Timer()
If Option4.Value = True Then
 If GetTheURL = "" Then
 Capt1.Enabled = False
 Me.Caption = "Google Power 2 - Their Are No Open Browsers !"
 Else
 Capt1.Enabled = True
 Me.Caption = "Searching " & GetTheURL
 End If
Else
Me.Caption = "Google Power 2"
End If
End Sub
Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
' Example: Call FormOnTop(me.hWnd, True)
On Error GoTo Goof
Goof:

    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    

    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub
