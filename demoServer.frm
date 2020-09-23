VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InterCommVB Server Application (Demo)"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Vote 
      Caption         =   "&Vote Now...."
      Height          =   330
      Left            =   3450
      TabIndex        =   7
      Top             =   2265
      Width           =   1860
   End
   Begin VB.CommandButton Transmit 
      Caption         =   "Transmit Now >>"
      Height          =   330
      Left            =   5325
      TabIndex        =   5
      Top             =   2265
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type any Text that you want transmitted"
      Height          =   1620
      Left            =   15
      TabIndex        =   1
      Top             =   405
      Width           =   7440
      Begin VB.CheckBox Dyn 
         Caption         =   "&Dynamic Transmissiion"
         Height          =   195
         Left            =   5100
         TabIndex        =   4
         Top             =   1335
         Value           =   1  'Checked
         Width           =   2880
      End
      Begin VB.TextBox SData 
         Height          =   1020
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   225
         Width           =   7200
      End
   End
   Begin VB.Label ProcID 
      Caption         =   "Current Process ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   6
      Top             =   2325
      Width           =   4380
   End
   Begin VB.Label Lastres 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   2070
      Width           =   7440
   End
   Begin VB.Label Label3 
      Caption         =   "Welcome to the InterCommVB Demonstration !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7530
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   0
      Top             =   300
      Width           =   7590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Const SW_SHOWNORMAL = 1
Const CodeID = 25759


Dim WithEvents Server As InterCommVB.IServer
Attribute Server.VB_VarHelpID = -1
Private Sub Form_Load()
Lastres.Caption = "Result of last transmission:N/A"
'Initialize the IServer object
Set Server = New IServer
ProcID.Caption = "Current Process ID:" & CStr(GetCurrentProcessId())

'The last parameter if set to true, ignores a missing
'data channel and fired the OnConnectionWait event
'It then waits for the client to connect to the client side
'interface and causes the OnChannelOpen() event to fire
Server.ConnectToDataChannel 1, Me.hwnd, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Destroy the IServer Object
Server.DisconnectFromChannel
Set Server = Nothing
End Sub

Private Sub SData_Change()
If Dyn.Value = 1 Then
'Transmit text in textbox to client registered with id# 1
Server.TransmitToClient SData.Text
End If

End Sub

Private Sub Server_OnChannelClose(ByVal ChannelID As Long)
    MsgBox "The Data channel has been closed by the client - No more transmission possible", vbExclamation Or vbOKOnly, "Channel ID:" & ChannelID
End Sub

Private Sub Server_OnChannelOpen(ByVal ChannelID As Long)
    MsgBox "The client app had opened the data channel - Communication is now possible"
End Sub

Private Sub Server_OnChannelReOpen(ByVal ChannelID As Long)
    MsgBox "The data channel has been re-opened by the client - Data transfer is now possible", vbExclamation Or vbOKOnly, "Channel ID:" & ChannelID
End Sub

Private Sub Server_OnConnectionFailure(Reason As String)
    MsgBox "Unable to connect to client - Reason:" & Reason
End Sub

Private Sub Server_OnConnectionSuccess()
    MsgBox "Connection to client established"
End Sub

Private Sub Server_OnConnectionWait()
    MsgBox "The client app had not initialized the data channel as yet. But the communication channel has been initiated and communication will be possible the moment the client initializes the data reception service"
End Sub

Private Sub Server_OnTransmissionFailure(Reason As String)
    'This event fires when transmission fails

    Lastres.Caption = "Result of last transmission:Failure     " & "Reason:" & Reason
End Sub

Private Sub Server_OnTransmissionSuccess()
    'This event fires when transmission is successful
    Lastres.Caption = "Result of last transmission:Success"
    
End Sub

Private Sub Server_OnVBInternalError(ByVal ErrCode As Long, ByVal ErrDesc As String)
    MsgBox "Internal Error - " & ErrDesc, vbCritical Or vbOKOnly, "Error Code:" & ErrCode
End Sub

Private Sub Transmit_Click()
'Transmit text in textbox to client registered with id# 1
Server.TransmitToClient SData.Text
End Sub


Sub VoteNow(URL As String)
    Dim Res As Long
    Dim TFile As String, Browser As String, Dum As String
    
    TFile = App.Path + "\test.htm"
    Open TFile For Output As #1
    Close
    Browser = String(255, " ")
    Res = FindExecutable(TFile, Dum, Browser)
    Browser = Trim$(Browser)
    
    If Len(Browser) = 0 Then
        MsgBox "Cannot find browser"
        Exit Sub
    End If
    
    Res = ShellExecute(Me.hwnd, "open", Browser, URL, Dum, SW_SHOWNORMAL)
    If Res <= 32 Then
        MsgBox "Cannot open web page"
        Exit Sub
    End If
End Sub

Private Sub Vote_Click()
VoteNow ("http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=25759&lngWId=1")
End Sub
