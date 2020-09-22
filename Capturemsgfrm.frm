VERSION 5.00
Begin VB.Form Capturemsgfrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Capturemsgfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Capturemsgfrm.frx":030A
   ScaleHeight     =   2760
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Height          =   200
      Left            =   1800
      MouseIcon       =   "Capturemsgfrm.frx":6B2A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1920
      Width           =   200
   End
   Begin VB.CheckBox Check1 
      Height          =   200
      Left            =   2520
      MouseIcon       =   "Capturemsgfrm.frx":6E34
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1920
      Width           =   200
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   3720
      MouseIcon       =   "Capturemsgfrm.frx":713E
      MousePointer    =   99  'Custom
      Picture         =   "Capturemsgfrm.frx":7448
      ScaleHeight     =   120
      ScaleWidth      =   510
      TabIndex        =   6
      ToolTipText     =   " Move "
      Top             =   160
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1560
      Picture         =   "Capturemsgfrm.frx":77CA
      ScaleHeight     =   330
      ScaleWidth      =   1425
      TabIndex        =   2
      Top             =   2160
      Width           =   1425
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1560
      MouseIcon       =   "Capturemsgfrm.frx":90CC
      MousePointer    =   99  'Custom
      Picture         =   "Capturemsgfrm.frx":93D6
      ScaleHeight     =   330
      ScaleWidth      =   1425
      TabIndex        =   1
      ToolTipText     =   " Continue "
      Top             =   2160
      Width           =   1425
   End
   Begin BmptoIcon.Fader Fader1 
      Left            =   5160
      Top             =   1200
      _ExtentX        =   979
      _ExtentY        =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide The Taskbar"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide Desktop Icons"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Capture Mode"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   1200
   End
   Begin VB.Label lblscreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "               After You  Close This Dialog               The WHOLE Screen Will Be Captured."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblactive 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"Capturemsgfrm.frx":ACD8
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "Capturemsgfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, Ip As Any) As Long
 Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'show/hide desk - icons
Private TaskBar As Long, StartButton As Long, Icons As Long, NotificationArea As Long, StartButtonCaption As Long, Clock As Long

Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12
Dim blnAboveVer4 As Boolean

Private Sub captr()


Bmp2Icon.WindowState = 0
Shell_NotifyIcon NIM_DELETE, IconData
Bmp2Icon.Show
Bmp2Icon.Pictureimage.Top = 0: Bmp2Icon.Pictureimage.Left = 0
Bmp2Icon.Set_Scrolls
Bmp2Icon.HScroll1.Value = 1
Bmp2Icon.VScroll1.Value = 1
Check1.Value = 0
Check2.Value = 0
ShowWindow Icons, 4
ShowWindow TaskBar, 4
End Sub




Private Sub Check1_Click()
If Check1.Value = 1 Then
ShowWindow Icons, 0
Else
Check1.Value = 0
ShowWindow Icons, 4
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
ShowWindow TaskBar, 0
Else
Check2.Value = 0
ShowWindow TaskBar, 4
End If

End Sub

Private Sub Form_Activate()
Fader1.FadeIn
 
End Sub

Private Sub Form_Load()

'to hide / show and /or taskbar / desktop
On Error Resume Next
    TaskBar = FindWindow("Shell_TrayWnd", vbNullString)
    Icons = FindWindowEx(0&, 0&, "Progman", vbNullString)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = True
Picture3.Visible = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
End Sub

Private Sub Picture1_Click()
Fader1.FadeOut
Capturemsgfrm.Hide
 If lblactive.Visible = True Then
 
 Dim bEndTime As Date
       bEndTime = DateAdd("s", 3, Now)
        Do Until Now > bEndTime
            DoEvents
         Loop
        Set Bmp2Icon.Pictureimage.Picture = CaptureActiveWindow()
       
        captr
  Else
  
  If lblscreen.Visible = True Then
    
    Clipboard.Clear
     
        If blnAboveVer4 Then
        keybd_event VK_SNAPSHOT, 0, 0, 0
    Else
  
   keybd_event VK_SNAPSHOT, 1, 0, 0
   End If
        
        
         Set Bmp2Icon.Pictureimage.Picture = CaptureScreen()
         
         captr
  End If
  End If
  
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = False
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Capturemsgfrm.ZOrder 0

ReleaseCapture
      
      SendMessage Capturemsgfrm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
