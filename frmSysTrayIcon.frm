VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Adding an icon to the system tray.
'by Peh Tee Howe, 2002


'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA

' create a form named Form1
' add a commond dialog control named CommonDialog1
' add two command buttons named Command1 and Command2

Private Sub Command1_Click()
   'Click this button to add an icon to the taskbar status area.

   'Set the individual values of the NOTIFYICONDATA data type.
   nid.cbSize = Len(nid)
   nid.hWnd = Form1.hWnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = Form1.Icon
   nid.szTip = "Taskbar Status Area Sample Program" & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Command2_Click()
   'Click this button to delete the added icon from the taskbar
   'status area by calling the Shell_NotifyIcon function.
   Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Load()
   'Set the captions of the command button when the form loads.
   Command1.Caption = "Add an Icon"
   Command2.Caption = "Delete Icon"
End Sub

Private Sub Form_Terminate()
   'Delete the added icon from the taskbar status area when the
   'program ends.
   Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_MouseMove _
   (Button As Integer, _
    Shift As Integer, _
    X As Single, _
    Y As Single)
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
       CommonDialog1.DialogTitle = "Select an Icon"
       sFilter = "Icon Files (*.ico)|*.ico"
       sFilter = sFilter & "|All Files (*.*)|*.*"
       CommonDialog1.Filter = sFilter
       CommonDialog1.ShowOpen
       If CommonDialog1.FileName <> "" Then
          Form1.Icon = LoadPicture(CommonDialog1.FileName)
          nid.hIcon = Form1.Icon
          Shell_NotifyIcon NIM_MODIFY, nid
       End If
       Case WM_RBUTTONDOWN
          Dim ToolTipString As String
          ToolTipString = InputBox("Enter the new ToolTip:", "Change ToolTip")
          If ToolTipString <> "" Then
             nid.szTip = ToolTipString & vbNullChar
             Shell_NotifyIcon NIM_MODIFY, nid
          End If
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select
End Sub
