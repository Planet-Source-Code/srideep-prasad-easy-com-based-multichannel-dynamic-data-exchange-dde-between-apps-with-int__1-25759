VERSION 5.00
Begin VB.Form GlobalWin 
   AutoRedraw      =   -1  'True
   Caption         =   "InterCommVB II Hidden Connection Manager Helper Window"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Active 
      Height          =   480
      Left            =   1410
      Picture         =   "GlobalWin.frx":0000
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Inactive 
      Height          =   480
      Left            =   675
      Picture         =   "GlobalWin.frx":0442
      Top             =   2145
      Width           =   480
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "&Tray Menu"
      Begin VB.Menu About 
         Caption         =   "&About InterCommVB II...."
      End
      Begin VB.Menu SEP1 
         Caption         =   "-"
      End
      Begin VB.Menu Terminate 
         Caption         =   "&Terminate InterCommVB...."
      End
      Begin VB.Menu SEP2 
         Caption         =   "-"
      End
      Begin VB.Menu CStat 
         Caption         =   "&Connection Status...."
      End
      Begin VB.Menu SEP3 
         Caption         =   "-"
      End
      Begin VB.Menu Credit 
         Caption         =   "&Credits...."
      End
   End
End
Attribute VB_Name = "GlobalWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tray As NOTIFYICONDATA, Busy As Boolean

Private Sub About_Click()
    MsgBox "InterCommVB II COM Based DDE Server" & Chr$(13) & "Created By:Srideep Prasad" & Chr$(13) & Chr$(13) & "Copyright(C) Srideep Prasad" & Chr$(13) & "Please e-mail all suggestions/bug reports to srideepprasad@yahoo.com", vbInformation Or vbOKOnly, "About InterCommVB II"
End Sub

Private Sub Credit_Click()
    MsgBox "Following are the names of a few people who have helped by submitting bug reports" & Chr$(13) & Chr$(13) & "Jan Alsenz - For notifying the IsChannelRegistered() method bug" & Chr$(13) & "Jon - For reporting the resource leak bug that manifested when the client or server crashed", vbExclamation Or vbOKOnly, "Credits"
End Sub

Private Sub CStat_Click()
    MsgBox "Total Server Processes Registered:" & GetProp(Me.hwnd, "ServerCount") & Chr$(13) & "Total Client Processes Registered:" & GetProp(Me.hwnd, "ClientCount") & Chr$(13), vbInformation, "InterCommVB II Connection Status"
End Sub

Private Sub Form_Load()
SetProp Me.hwnd, "GlobalWin", ObjectPtr(Me)
With Tray
    .cbSize = Len(Tray)
    .hwnd = Me.hwnd
    .hIcon = Inactive.Picture
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .szTip = "InterCommVB II DDE Server" & vbNullChar
    .uCallBackMessage = WM_MOUSEMOVE
End With
Shell_NotifyIcon NIM_ADD, Tray
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case X
    Case WM_RBUTTONUP
        Me.PopupMenu Me.TrayMenu
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, Tray
End Sub

Sub NormaliseIcon()
With Tray
    .cbSize = Len(Tray)
    .hwnd = Me.hwnd
    .hIcon = Inactive.Picture
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .szTip = "InterCommVB II DDE Server" & vbNullChar
    .uCallBackMessage = WM_MOUSEMOVE
End With
Shell_NotifyIcon NIM_MODIFY, Tray
Busy = False
End Sub

Sub HighlightIcon()

With Tray
    .cbSize = Len(Tray)
    .hwnd = Me.hwnd
    .hIcon = Active.Picture
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .szTip = "InterCommVB II DDE Server" & vbNullChar
    .uCallBackMessage = WM_MOUSEMOVE
End With
Shell_NotifyIcon NIM_MODIFY, Tray
Busy = True
End Sub

Function GetBusyFlag() As Boolean
    GetBusyFlag = GWin
End Function

Private Sub Terminate_Click()
Dim Choice As VbMsgBoxResult
Choice = MsgBox("Terminating InterCommVB II will cause dependent processes to crash, terminate unexpectedly or can result in data loss. Do you still want to terminate InterCommVB ?", vbYesNo Or vbDefaultButton2 Or vbCritical, "InterCommVB II Critical Message")
If Choice = vbYes Then End
End Sub
