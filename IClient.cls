VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim CliID As Long
Dim DF As CForm
Dim SWnd As Long
Public Event OnDataArrival(bData As String)
Attribute OnDataArrival.VB_Description = "Event fires when data is transmitted by the server process"
Public Event OnServerConnect(ByVal ChannelID As Long)
Attribute OnServerConnect.VB_Description = "Event fires when the server process connects to the active data channel"
Public Event OnServerDisconnect(ByVal ChannelID As Long)
Attribute OnServerDisconnect.VB_Description = "Event fires when the server process disconnects to the active data channel"
Public Event OnChannelRegistrationSuccess(ByVal ChannelID As Long)
Public Event OnChannelRegistrationFailure(ByVal ChannelID As Long, Reason As String)
Public Event OnVBInternalError(ByVal ErrCode As Long, ByVal ErrDesc As String)
Attribute OnVBInternalError.VB_Description = "Event fires when an unexpected internal error occurs"

Sub RegisterDataChannel(ChannelID As Long, WindowHandle As Long)
Attribute RegisterDataChannel.VB_Description = "Registers a specified data channel and prepares it for data input"
Dim Wnd As Long
If ChannelID < 0 Then
    Err.Raise vbObjectError + 530, , "Cannot register Client with given ID"
    Exit Sub
End If
Wnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))
'Is data channel already open ?
'If yes generate an error
If Wnd <> 0 Then
    If (GetProp(Wnd, "ClientWindow") <> 0) And (IsWindow(GetProp(Wnd, "ClientWindow")) = 0) Then
        MsgBox "A previous client process had already registered the requested data channel [ID#]" & ChannelID & "].But since the client has apparently crashed or terminated unexpectedly, InterCommVB II will now attempt to redirect the channel to the current client", vbInformation Or vbOKOnly, "InterCommVB II Critical Message"
        SendMessage Wnd, WM_CLOSE, 0, 0
    Else
        RaiseEvent OnChannelRegistrationFailure(ChannelID, "Cannot register Client with given ID")
        Exit Sub
    End If
End If
'Load the Hidden client side form and set its internal variables

If IsWindow(WindowHandle) = 0 Then
    RaiseEvent OnChannelRegistrationFailure(ChannelID, "Invalid window handle specified")
    Exit Sub
End If
On Local Error GoTo 500

Set DF = New CForm
Load DF
DF.Visible = False
DF.SetClient Me
'Set a unique caption to the hidden form to enable identification
DF.Caption = "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID)
CliID = ChannelID

SetProp DF.hwnd, "ClientWindow", WindowHandle
'Store a pointer to this class as the property "ClientClass"
SetProp DF.hwnd, "ClientClass", ObjectPtr(Me)

'Is the Server side interface initialized
SWnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))
If SWnd <> 0 Then
'Yes - So this means that after the server app had connected to
'this data channel, the client app disconnected from the channel
'Now we will set the "Action" property to sacChannelOpen = 2
'So that we can fire the OnChannelReopen() event at the Server side
    SetProp SWnd, "Action", 2
    
'Transmit the WM_SIZE message to the Server side form
'so that the FORM_RESIZE event will fire
'The Action property will than be validated and the appropriate
'event will be fired at the Server side
    PostMessage SWnd, WM_SIZE, 0, 0
    DF.SetServerFormHandle SWnd
End If
RaiseEvent OnChannelRegistrationSuccess(ChannelID)
DF.SetID ChannelID
Set GWin = GetGlobalWindow()
GoTo 550
500 RaiseEvent OnVBInternalError(Err.Number, Err.Description)
    Err.Clear
550 End Sub

Function GetNextValidID() As Long
Attribute GetNextValidID.VB_Description = "Gets the next available valid Data Channel ID"
'This function returns the next valid Data channel ID number
Dim Wnd As Long, Iter As Long
Iter = -1
Do
    Iter = Iter + 1
    Wnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(Iter))
Loop Until Wnd = 0
GetNextValidID = Iter
End Function

Function IsChannelRegistered(ChannelID As Long) As Boolean
Attribute IsChannelRegistered.VB_Description = "Returns True if the specified channel is registered"
'This function returns True if a particular Data Channel is registered
    Dim Wnd As Long
    Wnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))
    If Wnd = 0 Then
        IsChannelRegistered = False
    Else
        IsChannelRegistered = True
    End If
End Function


Sub UnregisterChannel()
Attribute UnregisterChannel.VB_Description = "Unregisters the currently registered data channel"
'This function unregisters the data channel and closes it
Unload DF
CliID = 0
'We now set the Action property to 1 or sacChannelClose
'and attempt to fire the OnChannelClose event at the server app end
SetProp SWnd, "Action", 1
PostMessage SWnd, WM_SIZE, 0, 0
SWnd = 0
Set DF = Nothing
70 End Sub

Friend Sub SetServerWnd(SHandle As Long)
'This sub is called by the IServer class to pass the
'handle of the Server side window to this class
'We can then post messages to the window
SWnd = SHandle
DF.SetServerFormHandle SWnd
End Sub

Friend Sub RaiseDataEvent(bData As String)
'This sub is called by the Client side form
'whenever data is transmitted so as to raise the
'OnDataArrival event

    RaiseEvent OnDataArrival(bData)
    
'We post the WM_SIZE message to the server window
'and cause the OnTransmissionSuccess event to fire
'at the server side
    If IsWindow(SWnd) <> 0 Then
        PostMessage SWnd, WM_SIZE, 0, 0
    Else
        Dim GWin As GlobalWin
        Set GWin = GetGlobalWindow()
        GWin.NormaliseIcon
        RaiseEvent OnServerDisconnect(CliID)
    End If
110 End Sub


Friend Sub RaiseConnectEvent()
'This sub is called by the client side form
'whenever the server app connects to this data channel

    RaiseEvent OnServerConnect(CliID)
End Sub
Friend Sub RaiseDisconnectEvent()
'This sub is called by the client side form
'whenever the server app disconnects to this data channel
    
    SWnd = 0
    RaiseEvent OnServerDisconnect(CliID)
End Sub

Friend Sub RaiseIntError(ErrNo As Long, ErrDesc As String)
    RaiseEvent OnVBInternalError(ErrNo, ErrDesc)
End Sub

Sub SetCrashDetectInterval(Interval As Long)
    If Interval < 0 Then Exit Sub
    If CFrm Is Nothing Then
    Else
        Dim SFrm As SForm
        Set SFrm = GetObject(GetProp("SForm", SWnd))
        If SFrm Is Nothing Then
        Else
            SFrm.SetInterval Interval
        End If
        
        DF.SetInterval Interval
    End If
End Sub

