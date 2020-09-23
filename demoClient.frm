VERSION 5.00
Begin VB.Form Client 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InterCommVB Client Application (Demo)"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Data Recieved from Server Application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   15
      TabIndex        =   5
      Top             =   1920
      Width           =   7590
      Begin VB.TextBox CData 
         Height          =   2115
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   240
         Width           =   7395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   4275
      Width           =   7605
   End
   Begin VB.PictureBox Logo 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   -15
      Picture         =   "demoClient.frx":0000
      ScaleHeight     =   750
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   4470
      Width           =   3060
   End
   Begin VB.Label Label1 
      Caption         =   $"demoClient.frx":0AE4
      Height          =   675
      Left            =   3075
      TabIndex        =   8
      Top             =   4470
      Width           =   4530
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
      Height          =   210
      Left            =   0
      TabIndex        =   7
      Top             =   1695
      Width           =   7545
   End
   Begin VB.Label Label5 
      Caption         =   $"demoClient.frx":0B91
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   1110
      Width           =   7575
   End
   Begin VB.Label Label4 
      Caption         =   $"demoClient.frx":0C63
      Height          =   840
      Left            =   15
      TabIndex        =   3
      Top             =   495
      Width           =   7590
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   15
      Top             =   375
      Width           =   7590
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
      Left            =   15
      TabIndex        =   2
      Top             =   90
      Width           =   7530
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Dim WithEvents Client As InterCommVB.IClient
Attribute Client.VB_VarHelpID = -1

Private Sub Client_OnChannelRegistrationFailure(ByVal ChannelID As Long, Reason As String)
    MsgBox "Data channel could not be registered - Reason :" & Reason
    End
End Sub

Private Sub Client_OnChannelRegistrationSuccess(ByVal ChannelID As Long)
    MsgBox "Data communication channel was registered successfully"
End Sub

Private Sub Client_OnDataArrival(bData As String)
'Event fired when data is transmitted by some server app
CData.Text = bData

End Sub

Private Sub Client_OnServerConnect(ByVal ChannelID As Long)
    MsgBox "The server application has established a connection", vbExclamation Or vbOKOnly, "Channel ID:" & ChannelID
End Sub


Private Sub Client_OnServerDisconnect(ByVal ChannelID As Long)
    MsgBox "The server application has disconnected from the data channel", vbExclamation Or vbOKOnly, "Channel ID:" & ChannelID
End Sub

Private Sub Client_OnVBInternalError(ByVal ErrCode As Long, ByVal ErrDesc As String)
    MsgBox "Internal Error - " & ErrDesc, vbCritical Or vbOKOnly, "Error Code:" & ErrCode
End Sub

Private Sub Form_Load()
ProcID.Caption = "Current Process ID:" & CStr(GetCurrentProcessId())

'Initialize the Client Object Variable
Set Client = New IClient
     
'Register this client with a unique ID
Client.RegisterDataChannel 1, Me.hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unregister this client before quitting

Client.UnregisterChannel
Set Client = Nothing
End Sub

