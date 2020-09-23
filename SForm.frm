VERSION 5.00
Begin VB.Form SForm 
   Caption         =   "Server Side Hidden Window"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5670
   Icon            =   "SForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer CrashTimer 
      Interval        =   15000
      Left            =   30
      Top             =   3390
   End
End
Attribute VB_Name = "SForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ServerClass As IServer, ID As Long

Private Sub CrashTimer_Timer()

If IsWindow(GetProp(Me.hwnd, "ServerWindow")) = 0 Then
    Dim CFormHandle As Long
    CFormHandle = FindWindow(vbNullString, "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(ID))
    
    
    SetProp Me.hwnd, "Busy", 1
    MsgBox "A previous server process had connected to the data channel [ID#" & ID & "].But since the server has apparently crashed or terminated unexpectedly, InterCommVB II will now attempt to terminate the connection and and recover wasted resources", vbInformation Or vbOKOnly, "InterCommVB II Critical Message"
    SetProp CFormHandle, "Action", 2
    PostMessage CFormHandle, WM_SIZE, 0, 0
    Set ServerClass = Nothing
    SendMessage Me.hwnd, WM_CLOSE, 0, 0
End If
        
End Sub

Private Sub Form_Load()
IncrementServerCount
SetProp Me.hwnd, "SForm", ObjectPtr(Me)
End Sub

Sub SetServerClass(SObj As IServer)
    Set ServerClass = SObj
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Here we inform the user when Windows attempts to terminate
'InterCommVB II during shutdown with open data channels since
'Doing so could cause the dependent apps to crash !
    If UnloadMode = vbAppWindows Then
    If GetProp(FindWindow(vbNullString, App.Title), "Confirm") = 0 Then
        Dim Choice As VbMsgBoxResult, Wnd As Long
            Choice = MsgBox("Windows is attempting to shut down InterCommVB II.Doing so will result in data loss or unexpected crashes of applications dependent on InterCommVB II's COM based DDE Services. Do you want to allow the shut down ?", vbCritical Or vbYesNo, "InterCommVB II Critical Message")
            If Choice = vbNo Then Cancel = True
            If Choice = vbYes Then
                SetProp FindWindow(vbNullString, App.Title), "Confirm", 1
            End If
        
    End If
    End If

End Sub

Private Sub Form_Resize()
'This event code is fired whenever a class posts the WM_SIZE
'message to this window
'Depending on the "Action" property, the appropriate event
'is fired
On Local Error GoTo 900

Dim SAction As SRAction
SAction = GetProp(Me.hwnd, "Action")
If SAction = sacDefault Then
    If ServerClass Is Nothing Then
    Else
       ServerClass.RaiseSuccess
    End If
End If

If SAction = sacChannelClose Then
    ServerClass.RaiseChannelClose
End If

If SAction = sacChannelOpen Then
    ServerClass.RaiseChannelOpen
End If

'Reset the "Action" property
SetProp Me.hwnd, "Action", 0
GoTo 950
900 ServerClass.RaiseIntError Err.Number, Err.Description
    Err.Clear
950 End Sub

Private Sub Form_Unload(Cancel As Integer)
    DecrementServerCount
    Set ServerClass = Nothing
End Sub


Sub SetID(iID As Long)
    ID = iID
End Sub

Sub SetInterval(Interval As Long)
    CrashTimer.Interval = Interval
End Sub

