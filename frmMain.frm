VERSION 5.00
Begin VB.Form frmMouseXY 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   210
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   2040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   2040
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrTimer 
      Index           =   1
      Interval        =   1
      Left            =   720
      Top             =   600
   End
   Begin VB.Timer tmrTimer 
      Index           =   0
      Interval        =   1
      Left            =   360
      Top             =   600
   End
   Begin VB.Image imgCloseUp 
      Height          =   210
      Left            =   960
      Picture         =   "frmMain.frx":1042
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCloseDown 
      Height          =   210
      Left            =   720
      Picture         =   "frmMain.frx":10B5
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image btnClose 
      Height          =   210
      Left            =   1800
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   240
   End
   Begin VB.Menu mnuMain 
      Caption         =   "1"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "Enable Mouse Erradication"
      End
   End
End
Attribute VB_Name = "frmMouseXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Name : MouseXY
' Author : Rudy Alex Kohn
' Purpose : Display mouse X/Y coordinates
' Extra : Build-In Fake Erratic Movement (optional :)

' Info -:-
'   Displays Mouse XY coordinates in a small box, you can drag/close it just like a normal window.
'   The display is only updated if movement is detected, furthermore, it don't take any CPU ;)

' Requirements : Windows 95+ (Works on all windows versions above)
'                A Mouse :D

Option Explicit
' Formdrag >
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' <
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private pt As POINTAPI
Private iLoop As Integer
Private iLoop2 As Integer

Private bErratic As Boolean ' State of erratic mouse movement


Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub Form_Load()
    Dim Ret As String, NC As Long
    Ret = String$(255, 0)
    NC = GetPrivateProfileString("MouseXY", "Font", "Tahoma", Ret, 255, App.Path & "\MouseXY.ini")
    If NC <> 0 Then Ret = Left$(Ret, NC)
    ' set font
    Me.Font = Ret

    Ret = String$(255, 0)
    NC = GetPrivateProfileString("MouseXY", "Back", App.Path & "\default.jpg", Ret, 255, App.Path & "\MouseXY.ini")
    If NC <> 0 Then Ret = Left$(Ret, NC)
    ' load background
    Me.Picture = LoadPicture(Ret)

    Ret = String$(255, 0)
    NC = GetPrivateProfileString("MouseXY", "Front", "255:255:255", Ret, 255, App.Path & "\MouseXY.ini")
    If NC <> 0 Then Ret = Left$(Ret, NC)
    ' set color to font here
    Dim nColor
    nColor = Split(Ret, ":")
    On Error GoTo errHandl:
    Me.ForeColor = RGB(nColor(0), nColor(1), nColor(2))

    Ret = String$(255, 0)
    NC = GetPrivateProfileString("MouseXY", "Erratic", "0", Ret, 255, App.Path & "\MouseXY.ini")
    If NC <> 0 Then Ret = Left$(Ret, NC)
    bErratic = CBool(Ret)

    Set btnClose.Picture = imgCloseUp.Picture ' Close pic
    Current = 0
With Me
    .Top = 0
    .Left = Screen.Width \ 2
End With
    mnu1.Checked = bErratic
    NC = 0
    Ret = vbNullString
    Erase nColor
    Exit Sub

errHandl:
    Me.ForeColor = RGB(255, 255, 255)
    Resume Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        ' Drag 'me' as long as the mouse button is pressed down
        FormDrag Me
    Case 2
        PopupMenu mnuMain
    End Select
End Sub

Private Sub btnClose_Click()
    ' Clean up and quit
    Set frmMouseXY = Nothing
    Erase MouseX
    Erase MouseY
    Unload Me
    End
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    SetCursorPos MouseX(Current - 1000), MouseY(Current - 1000)
End Sub

Private Sub FormDrag(frm As Object)
' This Sub allows you to move the form
  ReleaseCapture
  SendMessage frm.hwnd, &HA1, 2, 0&
End Sub

Private Sub btnClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Close button is pressed down
With btnClose
    Select Case Button
    Case 1
        If .Picture <> imgCloseDown.Picture Then Set .Picture = imgCloseDown.Picture
    Case 0
        If .Picture <> imgCloseUp.Picture Then Set .Picture = imgCloseUp.Picture
    End Select
End With
End Sub
Private Sub btnClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Close button released
    Set btnClose.Picture = imgCloseUp.Picture
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "MouseXY", "Options", "Erratic Movement", bErratic
End Sub

Private Sub mnu1_Click()
    mnu1.Checked = Not mnu1.Checked
    bErratic = Not bErratic
End Sub

Private Sub tmrTimer_Timer(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        GetCursorPos pt
With pt
        MouseX(Current) = .X
        MouseY(Current) = .Y
End With
        ' Print the x/y cordinates to the form
        Dim sTmp As String
        Me.Cls    'clear
        sTmp = CStr(MouseX(Current)) & " ," & CStr(MouseY(Current))
With Me
        .CurrentX = (.Width - (btnClose.Width + .TextWidth(sTmp))) / 2    ' make sure text is allways in center X
End With
        Me.Print CStr(MouseX(Current)) & " ," & CStr(MouseY(Current))
        Select Case Current
        Case Is < 9999
            Current = Current + 1
        Case Else
'        SaveXY
            Current = 0
        End Select
    Case 1
        If bErratic = False Then Exit Sub
        If GetTickCount Mod 2 Then
            SetCursorPos MouseX(Current - 1), MouseY(Current - 1)
        ElseIf GetTickCount Mod 1 Then
            For iLoop = 1 To 2
                SetCursorPos MouseX(Current) + iLoop, MouseY(Current) + iLoop2
                For iLoop2 = 1 To 2
                    SetCursorPos MouseX(Current - iLoop2), MouseY(Current - iLoop)
                Next iLoop2
            Next iLoop
'        a = True
        End If
    End Select
    DoEvents
End Sub
