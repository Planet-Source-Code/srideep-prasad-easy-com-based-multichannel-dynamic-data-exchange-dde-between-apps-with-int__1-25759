VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event OnTransmissionSuccess()
Attribute OnTransmissionSuccess.VB_Description = "Event fires when an attempt to transmit data to the client app succeeds"
Public Event OnTransmissionFailure(Reason As String)
Attribute OnTransmissionFailure.VB_Description = "Event fires when an attempt to transmit data to the client app fails"
Public Event OnConnectionWait()
Attribute OnConnectionWait.VB_Description = "Event fires when the server instantiates a data channel and is is beginning to wait for the client side interface to be instantiated by client process/thread"
Public Event OnConnectionSuccess()
Attribute OnConnectionSuccess.VB_Description = "Event fires when an attempt to connect to a data channel succeeds"
Public Event OnConnectionFailure(Reason As String)
Attribute OnConnectionFailure.VB_Description = "Event fires when an attempt to connect to a data channel fails"
Public Event OnChannelReOpen(ByVal ChannelID As Long)
Attribute OnChannelReOpen.VB_Description = "Event fires when the client reopens  the data channel associated with the given instance of InterCommVB's server side interface"
Public Event OnChannelOpen(ByVal ChannelID As Long)
Attribute OnChannelOpen.VB_Description = "Event fires when the client opens the associated data channel "
Public Event OnChannelClose(ByVal ChannelID As Long)
Attribute OnChannelClose.VB_Description = "Event fires when the client closes the data channel associated with the given instance of InterCommVB's server side interface"
Public Event OnVBInternalError(ByVal ErrCode As Long, ByVal ErrDesc As String)
Attribute OnVBInternalError.VB_Description = "Event fires when an unexpected internal error occurs"
Dim CWnd As Long, CFrm As CForm, CID As Long
Dim ServerInit As Boolean, GWin As GlobalWin
Dim SFrm As SForm, ClientClass As IClient
Sub ConnectToDataChannel(ChannelID As Long, WindowHandle As Long, Optional IgnoreMissingClient As Boolean = False)
Attribute ConnectToDataChannel.VB_Description = "Connects to the specified data channel"
On Local Error GoTo 10

If CWnd <> 0 Then
    Call DisconnectFromChannel
End If

If IsWindow(WindowHandle) = 0 Then
    RaiseEvent OnConnectionFailure("Invalid Window handle specified")
    GoTo 150
    Exit Sub
End If

Dim Wnd As Long
Dim DFPtr As Long, DF As CForm, STwnd As Long
'Check to see whether another app has hooked on to the
'data channel specified
STwnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))


If STwnd <> 0 Then
    If (GetProp(STwnd, "ServerWindow") <> 0) And (IsWindow(GetProp(STwnd, "ServerWindow")) = 0) Then
    'If the server had crashed then free the resources and redirect
    'the connection to this server
        MsgBox "A previous server process had connected to the requested data channel [ID#" & ChannelID & "].Since the server had apparently crashed or terminated unexpectedly, InterCommVB II will now attempt to redirect the connection to this server", vbInformation Or vbOKOnly, "InterCommVB II Critical Message"
        SendMessage STwnd, WM_CLOSE, 0, 0
    Else
        'If yes, then generate the error notification
        RaiseEvent OnConnectionFailure("Another process or thread has already connected to this channel")
        GoTo 150
        Exit Sub
    End If
End If
    
'Find handle of the Client side data helper Window
Wnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))

If (Wnd = 0) And (IgnoreMissingClient = True) Then
    ServerInit = True
    Call IgnoreClientAndConnect(WindowHandle, ChannelID)
    GoTo 150
End If
ServerInit = False

If Wnd = 0 Then
    'If no such Window is found, then raise the error notification
    RaiseEvent OnConnectionFailure("Unable to find data channel")
    GoTo 150
    Exit Sub
End If

If GetProp(Wnd, "ClientClass") = 0 Then
    'If the object pointer properties are not found
    'then again raise an error event
    RaiseEvent OnConnectionFailure("Unable to get pointer to client side COM interface[IClient]")
    GoTo 150
    Exit Sub
End If

If GetProp(Wnd, "Busy") = 1 Then
    'If the "Busy" property is 1 (it is set during transmission)
    'then the channel may be still busy
    'We must Wait for it to free up
    RaiseEvent OnConnectionFailure("This communication channel is busy - A previous request is pending")
    GoTo 150
    Exit Sub
End If
CWnd = Wnd

'Retrieve the pointer to the client side class and helper form
'and from the pointers retrieve references to the object
'themselves
DFPtr = GetProp(Wnd, "CForm")
Set CFrm = GetObj(DFPtr)

'Also load the Server side helper Window and
'set its unique caption to enable idenficiation
Set SFrm = New SForm
Load SFrm
SFrm.Visible = False
SetProp SFrm.hwnd, "ServerWindow", WindowHandle
SFrm.Caption = "InterCommVB Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID)
SFrm.SetServerClass Me
Set ClientClass = GetObj(GetProp(Wnd, "ClientClass"))
ClientClass.SetServerWnd SFrm.hwnd
'Notifiy the success of the connection
RaiseEvent OnConnectionSuccess

'Set the Action property of the Client side Window to 1 (acConnectEvent)
'and post the WM_SIZE message so that the OnServer connect
'event will fire at the client side
SetProp CWnd, "Action", 1
PostMessage CWnd, WM_SIZE, 0, 0
SFrm.SetID ChannelID
Set GWin = GetGlobalWindow()
GoTo 20
10  RaiseEvent OnVBInternalError(Err.Number, Err.Description)
    Err.Clear

20  CID = ChannelID
150 End Sub

Sub TransmitToClient(bData As String)
Attribute TransmitToClient.VB_Description = "Transmits the specified data to the client process"

On Local Error GoTo 30

If (CWnd = 0) And (CID <> 0) Then
    AttemptReConnect CID
        If CWnd <> 0 Then
            CFrm.SetData bData
            SetProp CWnd, "Action", 0
            PostMessage CWnd, WM_SIZE, 0, 0
        Else
            GoTo 250
        End If
End If

If GetProp(CWnd, "Busy") = 1 Then
    'If the data channel is busy (a transaction is in progress)
    'then raise an error event
    RaiseEvent OnTransmissionFailure("This communication channel is busy - A previous request is pending")
    GoTo 50
    Exit Sub
End If

250 If (IsWindow(CWnd) = 0) Or (IsWindow(GetProp(CWnd, "ClientWindow")) = 0) Then
        'This is actually unnecessary !
        'If some problems with the Window handle  then
        'raise an error event
        Set CFrm = Nothing
        RaiseEvent OnTransmissionFailure("The Data channel is not open")
        If (IsWindow(GetProp(CWnd, "ClientWindow")) = 0) Then SendMessage CWnd, WM_CLOSE, 0, 0
        CWnd = 0
        Set ClientClass = Nothing
        GoTo 50
        Exit Sub
    End If
    
'Set the internal variables of the client side helper form
'and raise the OnDataArrival event at the client side
If GWin.GetBusyFlag = False Then GWin.HighlightIcon
CFrm.SetData bData
SetProp CWnd, "Action", 0
PostMessage CWnd, WM_SIZE, 0, 0
GoTo 50
30 RaiseEvent OnVBInternalError(Err.Number, Err.Description)
   Err.Clear
50 End Sub

Function IsChannelRegistered(ChannelID As Long) As Boolean
Attribute IsChannelRegistered.VB_Description = "Returns true if the specified data channel is registered"
'Returns True if the specified channel is registered
    Dim Wnd As Long
    Wnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))
    If Wnd = 0 Then
        IsChannelRegistered = False
    Else
        IsChannelRegistered = True
    End If
End Function

Sub DisconnectFromChannel()
Attribute DisconnectFromChannel.VB_Description = "DisConnects from  the specified data channel"
'Disconnectes the server from a particular server app
'from the specified channel
    Set CFrm = Nothing
    Set GWin = Nothing
    If SFrm Is Nothing Then
    Else
        Unload SFrm
        Set SFrm = Nothing
    End If
    Set ClientClass = Nothing

'Fire the OnServerDisconnect event at the client side
    If GetProp(CWnd, "Busy") = 0 Then
        SetProp CWnd, "Action", 2
        PostMessage CWnd, WM_SIZE, 0, 0
    End If
    CID = 0
    CWnd = 0
    
End Sub
Function IsConnected(ChannelID As Long) As Boolean
Attribute IsConnected.VB_Description = "Returns true if a server app has connected to a given data channel"
'Returns whether any server app has hooked on to a specified
'data channel
Dim STwnd As Long
    STwnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))
    If STwnd = 0 Then
        IsConnected = False
    Else
        IsConnected = True
    End If
End Function


Friend Sub RaiseSuccess()
'This sub is called by the server side helper window
'when the transmission is a success
    RaiseEvent OnTransmissionSuccess
    If GWin.GetBusyFlag = False Then GWin.NormaliseIcon
End Sub

Friend Sub RaiseChannelClose()
'This sub is called by the server side helper window
'when the client closes the data channel

        Set CFrm = Nothing
        CWnd = 0
        Set ClientClass = Nothing
        RaiseEvent OnChannelClose(CID)
End Sub

Friend Sub RaiseChannelOpen()
'This sub is called by the server side helper window
'when the client reopens/opens the data channel on which this class
'is hooked

    AttemptReConnect CID
If ServerInit = False Then
    RaiseEvent OnChannelReOpen(CID)
Else
    RaiseEvent OnChannelOpen(CID)
    ServerInit = False
End If
End Sub

Friend Sub RaiseIntError(ErrNo As Long, ErrDesc As String)
    RaiseEvent OnVBInternalError(ErrNo, ErrDesc)
End Sub

Friend Sub AttemptReConnect(ChannelID As Long)
On Local Error GoTo 400

'This sub is called to attempt a reconnection operation to the client
'in case the connection is unexpectedly closed

Dim Wnd As Long
Dim DFPtr As Long, DF As CForm
Wnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Client]:[Hidden Window]" & CStr(ChannelID))
SWnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))
If Wnd = 0 Then
    GoTo 350
    Exit Sub
End If

If GetProp(Wnd, "ClientClass") = 0 Then
    GoTo 350
    Exit Sub
End If

If GetProp(Wnd, "Busy") = 1 Then
    GoTo 350
    Exit Sub
End If
CWnd = Wnd
DFPtr = GetProp(Wnd, "CForm")

Set CFrm = GetObj(DFPtr)
Set ClientClass = GetObj(GetProp(Wnd, "ClientClass"))


ClientClass.SetServerWnd SFrm.hwnd
CID = ChannelID

GoTo 350
400 RaiseEvent OnVBInternalError(Err.Number, Err.Description)
    Err.Clear
350 End Sub


Friend Sub IgnoreClientAndConnect(WindowHandle As Long, ChannelID As Long)
On Local Error GoTo 1000

If CWnd <> 0 Then
    Call DisconnectFromChannel
End If

If IsWindow(WindowHandle) = 0 Then
    RaiseEvent OnConnectionFailure("Invalid Window handle specified")
    GoTo 15000
    Exit Sub
End If

Dim Wnd As Long
Dim DFPtr As Long, DF As CForm, STwnd As Long
'Check to see whether another app has hooked on to the
'data channel specified
STwnd = FindWindow(vbNullString, "InterCommVB Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID))

If STwnd <> 0 Then
    If (GetProp(STwnd, "ServerWindow") <> 0) And (IsWindow(GetProp(STwnd, "ServerWindow")) = 0) Then
    'If the server had crashed then free the resources and redirect
    'the connection to this server
        MsgBox "A previous server process had connected to the requested data channel [ID#" & ChannelID & "].Since the server had apparently crashed or terminated unexpectedly, InterCommVB II will now attempt to redirect the connection to this server", vbInformation Or vbOKOnly, "InterCommVB II Critical Message"
        SendMessage STwnd, WM_CLOSE, 0, 0
    Else
        'If yes, then generate the error notification
        RaiseEvent OnConnectionFailure("Another process or thread has already connected to this channel")
        GoTo 15000
        Exit Sub
    End If
End If
    
CWnd = 0

'Also load the Server side helper Window and
'set its unique caption to enable idenficiation
Set SFrm = New SForm
Load SFrm
SFrm.Visible = False
SetProp SFrm.hwnd, "ServerWindow", WindowHandle
SFrm.Caption = "InterCommVB Data Channel ID[Server]:[Hidden Window]" & CStr(ChannelID)
SFrm.SetServerClass Me
SFrm.SetID ChannelID
RaiseEvent OnConnectionWait
Set GWin = GetGlobalWindow()
GoTo 2000
1000 RaiseEvent OnVBInternalError(Err.Number, Err.Description)
     Err.Clear

2000  CID = ChannelID
15000 End Sub

Sub SetCrashDetectInterval(Interval As Long)
    If Interval < 0 Then Exit Sub
    If SFrm Is Nothing Then
    Else
        SFrm.SetInterval Interval
        CFrm.SetInterval Interval
    End If
End Sub
