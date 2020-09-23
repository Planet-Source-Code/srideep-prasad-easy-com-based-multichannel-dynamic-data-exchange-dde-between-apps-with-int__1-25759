VERSION 5.00
Begin VB.Form CForm 
   Caption         =   "Client Side Hidden Window"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "DF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer CrashTimer 
      Interval        =   10000
      Left            =   0
      Top             =   2760
   End
End
Attribute VB_Name = "CForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The timer fires the timer_event every 15 sec only (at default setting)
'So there is NO PERFORMANCE Problem !
Dim ClientClass As IClient, bData As String
Dim ID As Long, SFormHandle As Long

Private Sub CrashTimer_Timer()
If IsWindow(GetProp(Me.hwnd, "ClientWindow")) = 0 Then
    SetProp Me.hwnd, "Busy", 1
    MsgBox "A previous client process had registered the data channel [ID#" & ID & "].But since the client has apparently crashed or terminated unexpectedly, InterCommVB II will now attempt to terminate the data channel and free wasted resources", vbInformation Or vbOKOnly, "InterCommVB II Critical Message"
    SetProp SFormHandle, "Action", 1
    PostMessage SFormHandle, WM_SIZE, 0, 0
    Set ClientClass = Nothing
    SendMessage Me.hwnd, WM_CLOSE, 0, 0
End If
End Sub

Private Sub Form_Load()
IncrementClientCount
SetProp Me.hwnd, "CForm", ObjPtr(Me)
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

On Local Error GoTo 700

Dim Action As CRAction
Action = GetProp(Me.hwnd, "Action")

If Action = acDefault Then
        If ClientClass Is Nothing Then
        Else
                ClientClass.RaiseDataEvent bData
                bData = ""
                SetProp Me.hwnd, "Busy", 0
        End If
End If

If Action = acConnectEvent Then
    ClientClass.RaiseConnectEvent
End If

If Action = acDisConnectEvent Then
    ClientClass.RaiseDisconnectEvent
End If

SetProp Me.hwnd, "Action", 0
GoTo 750
700 ClientClass.RaiseIntError Err.Number, Err.Description
    Err.Clear
750 End Sub

Private Sub Form_Unload(Cancel As Integer)
    DecrementClientCount
    Set ClientClass = Nothing
End Sub


Sub SetClient(CObj As IClient)
Set ClientClass = CObj
End Sub

Sub SetData(oData As String)
    SetProp Me.hwnd, "Busy", 1
    bData = oData
End Sub

Sub SetServerFormHandle(Handle As Long)
    SFormHandle = Handle
End Sub

Sub SetID(iID As Long)
ID = iID
End Sub

Sub SetInterval(Interval As Long)
CrashTimer.Interval = Interval
End Sub
