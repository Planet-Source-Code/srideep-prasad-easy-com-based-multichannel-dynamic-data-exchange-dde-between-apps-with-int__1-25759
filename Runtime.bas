Attribute VB_Name = "Module1"
'Copyright(C) Srideep Prasad
'This code is only for private/non-commercial use


'I cannot give a line by line explanation unfortunately
'because the execution pattern of the code is highly non-linear
'and "zigzag"
'However I may mention that for each active data channel
'2 "dummy" windows are created and loaded - One
'for the server side and the other at the client side
'These windows are used to store object pointers against
'their Window handles as virtual properties
'When a particular action is to be undertaken, say the firing
'of an
'event, a WM_SIZE message is transmitted to the appripriate
'dummy Window after setting the "Action" property appropriately
'The WM_SIZE event is captured by VB and translated to
'The FORM_RESIZE event. At the FORM_RESIZE event code
'the "Action" property is retrieved and depending on its
'value appropriate events are fired



Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
          (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" _
          (ByVal hwnd As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long

Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const THREAD_BASE_PRIORITY_MIN = -2

Public Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Public Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN

Public Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Public Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)



'//UDT required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
 cbSize As Long             '//size of this UDT
 hwnd As Long               '//handle of the app
 uId As Long                '//unused (set to vbNull)
 uFlags As Long             '//Flags needed for actions
 uCallBackMessage As Long   '//WM we are going to subclass
 hIcon As Long              '//Icon we're going to use for the systray
 szTip As String * 64       '//ToolTip for the mouse_over of the icon.
End Type


'//Constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0             '//Flag : "ALL NEW nid"
Public Const NIM_MODIFY = &H1          '//Flag : "ONLY MODIFYING nid"
Public Const NIM_DELETE = &H2          '//Flag : "DELETE THE CURRENT nid"
Public Const NIF_MESSAGE = &H1         '//Flag : "Message in nid is valid"
Public Const NIF_ICON = &H2            '//Flag : "Icon in nid is valid"
Public Const NIF_TIP = &H4             '//Flag : "Tip in nid is valid"
Public Const WM_MOUSEMOVE = &H200      '//This is our CallBack Message
Public Const WM_LBUTTONDOWN = &H201    '//LButton down
Public Const WM_LBUTTONUP = &H202      '//LButton up
Public Const WM_LBUTTONDBLCLK = &H203  '//LDouble-click
Public Const WM_RBUTTONDOWN = &H204    '//RButton down
Public Const WM_RBUTTONUP = &H205      '//RButton up
Public Const WM_RBUTTONDBLCLK = &H206  '//RDouble-click

Public nid As NOTIFYICONDATA       '//global UDT for the systray function

Public Const WM_CLOSE = &H10

Public Const WM_SIZE = &H5
Public Enum CRAction
    acDefault = 0
    acConnectEvent = 1
    acDisConnectEvent = 2
End Enum

Public Enum SRAction
    sacDefault = 0
    sacChannelClose = 1
    sacChannelOpen = 2
End Enum


Function GetObj(Ptr As Long)
'Retrieves an Object from its pointer
Dim TObj As Object
CopyMemory TObj, Ptr, 4
Set GetObj = TObj
CopyMemory TObj, 0&, 4
End Function

Function ObjectPtr(Obj As Object)
'Returns a pointer to an object
Dim lpObj As Long
CopyMemory lpObj, Obj, 4
ObjectPtr = lpObj
End Function


Sub cLoadGlobalWin()
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
If Wnd = 0 Then
    Dim GFrm As New GlobalWin
    Load GFrm
    GFrm.Visible = False
    Set GFrm = Nothing
End If
End Sub

Sub SetGlobalProp(PropName As String, PVal As Long)
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
SetProp Wnd, PropName, PVal
End Sub

Function GetGlobalProp(PropName As String) As Long
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
GetGlobalProp = GetProp(Wnd, PropName)
End Function

Sub RemoveGlobalProp(PropName As String)
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
RemoveProp Wnd, PropName
End Sub

Sub IncrementClientCount()
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
SetProp Wnd, "ClientCount", GetGlobalProp("ClientCount") + 1
End Sub

Sub DecrementClientCount()
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
SetProp Wnd, "ClientCount", GetGlobalProp("ClientCount") - 1
Call ChkTerminate
End Sub

Sub IncrementServerCount()
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
SetProp Wnd, "ServerCount", GetGlobalProp("ServerCount") + 1
End Sub

Sub DecrementServerCount()
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
SetProp Wnd, "ServerCount", GetGlobalProp("ServerCount") - 1
Call ChkTerminate
End Sub

Sub ChkTerminate()
    If GetGlobalProp("ServerCount") = 0 And GetGlobalProp("ClientCount") = 0 Then
        Dim Wnd As Long
        Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
        SendMessage Wnd, WM_CLOSE, 0, 0
    End If
End Sub


Function GetGlobalWindow() As GlobalWin
Dim Wnd As Long
Wnd = FindWindow(vbNullString, "InterCommVB II Hidden Connection Manager Helper Window")
Set GetGlobalWindow = GetObj(GetProp(Wnd, "GlobalWin"))
End Function


Sub Main()
cLoadGlobalWin
End Sub
