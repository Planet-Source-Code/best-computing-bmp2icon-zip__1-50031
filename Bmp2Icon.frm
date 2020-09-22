VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Bmp2Icon 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   9720
   DrawMode        =   6  'Mask Pen Not
   Icon            =   "Bmp2Icon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Bmp2Icon.frx":0CCA
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   StartUpPosition =   2  'CenterScreen
   Begin BmptoIcon.IconX IconX 
      Left            =   5280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9240
      Picture         =   "Bmp2Icon.frx":72C5
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   36
      Top             =   60
      Width           =   315
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9240
      MouseIcon       =   "Bmp2Icon.frx":77E7
      MousePointer    =   99  'Custom
      Picture         =   "Bmp2Icon.frx":7AF1
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   35
      ToolTipText     =   " Close "
      Top             =   60
      Width           =   315
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8880
      MouseIcon       =   "Bmp2Icon.frx":7FE6
      MousePointer    =   99  'Custom
      Picture         =   "Bmp2Icon.frx":82F0
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   32
      ToolTipText     =   " AutoDock Color Save Options Box "
      Top             =   600
      Width           =   330
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9960
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9960
      Top             =   360
   End
   Begin VB.PictureBox saveoptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   900
      MouseIcon       =   "Bmp2Icon.frx":890A
      MousePointer    =   99  'Custom
      Picture         =   "Bmp2Icon.frx":8C14
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   203
      TabIndex        =   22
      ToolTipText     =   " Move "
      Top             =   1080
      Width           =   3045
      Begin VB.PictureBox Picture16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   840
         Picture         =   "Bmp2Icon.frx":98E5
         ScaleHeight     =   330
         ScaleWidth      =   1425
         TabIndex        =   30
         ToolTipText     =   " Cancel & Close Save "
         Top             =   2640
         Width           =   1425
      End
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   840
         MouseIcon       =   "Bmp2Icon.frx":B1E7
         Picture         =   "Bmp2Icon.frx":B4F1
         ScaleHeight     =   330
         ScaleWidth      =   1425
         TabIndex        =   29
         ToolTipText     =   " Cancel & Close Save "
         Top             =   2640
         Width           =   1425
      End
      Begin VB.PictureBox header 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         MouseIcon       =   "Bmp2Icon.frx":CDF3
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":D0FD
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   225
         TabIndex        =   23
         Top             =   0
         Width           =   3375
         Begin VB.PictureBox Picture17 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2820
            MouseIcon       =   "Bmp2Icon.frx":1060F
            MousePointer    =   99  'Custom
            Picture         =   "Bmp2Icon.frx":10919
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   14
            TabIndex        =   31
            ToolTipText     =   " Cancel & Close Save "
            Top             =   30
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "    Color Save Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00933E28&
            Height          =   255
            Left            =   0
            MouseIcon       =   "Bmp2Icon.frx":10BC3
            MousePointer    =   99  'Custom
            TabIndex        =   28
            ToolTipText     =   " Move "
            Top             =   60
            Width           =   3015
         End
      End
      Begin VB.Label bitsave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 24 Bit (   TrueColor   ) "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   555
         TabIndex        =   27
         ToolTipText     =   " Save as... ( 24 Bit Color ) "
         Top             =   2160
         Width           =   2025
      End
      Begin VB.Label bitsave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  8 Bit (  256 Colors   ) "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   570
         TabIndex        =   26
         ToolTipText     =   " Save as... ( 8 Bit Color ) "
         Top             =   1560
         Width           =   1995
      End
      Begin VB.Label bitsave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "     4 Bit (    16 Colors   )    "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   25
         ToolTipText     =   " Save as... ( 4 Bit Color ) "
         Top             =   960
         Width           =   2355
      End
      Begin VB.Label bitsave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 1 Bit ( Monochrome )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   24
         ToolTipText     =   " Save as... ( 1 Bit Color ) "
         Top             =   480
         Width           =   1875
      End
   End
   Begin VB.PictureBox PicMask 
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   9960
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      Top             =   2040
      Width           =   495
   End
   Begin VB.PictureBox IconXctl 
      Height          =   480
      Left            =   9960
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   33
      Top             =   1440
      Width           =   1200
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      Picture         =   "Bmp2Icon.frx":10ECD
      ScaleHeight     =   330
      ScaleWidth      =   210
      TabIndex        =   20
      Top             =   6960
      Width           =   210
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9300
      Picture         =   "Bmp2Icon.frx":112D7
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   19
      Top             =   6915
      Width           =   330
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9300
      Picture         =   "Bmp2Icon.frx":118AD
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   18
      Top             =   1140
      Width           =   330
   End
   Begin VB.PictureBox buttonbase 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   120
      Picture         =   "Bmp2Icon.frx":11E83
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   8
      Top             =   525
      Width           =   4815
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1320
         MouseIcon       =   "Bmp2Icon.frx":12537
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":12841
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   17
         ToolTipText     =   " Save As Options... "
         Top             =   60
         Width           =   330
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4320
         MouseIcon       =   "Bmp2Icon.frx":12E93
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":1319D
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   16
         ToolTipText     =   " About / Credits "
         Top             =   60
         Width           =   330
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3720
         MouseIcon       =   "Bmp2Icon.frx":137B7
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":13AC1
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   15
         ToolTipText     =   " Help (See Code ) "
         Top             =   60
         Width           =   330
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1320
         MouseIcon       =   "Bmp2Icon.frx":140DB
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":143E5
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   14
         ToolTipText     =   " Save As Options... "
         Top             =   60
         Width           =   330
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   720
         MouseIcon       =   "Bmp2Icon.frx":149FF
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":14D09
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   13
         ToolTipText     =   " Open SOURCE Picture "
         Top             =   60
         Width           =   330
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         MouseIcon       =   "Bmp2Icon.frx":15323
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":1562D
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   12
         ToolTipText     =   " Exit "
         Top             =   60
         Width           =   330
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3120
         MouseIcon       =   "Bmp2Icon.frx":15C47
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":15F51
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   11
         ToolTipText     =   " Capture ICON - SIZE "
         Top             =   60
         Width           =   330
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2520
         MouseIcon       =   "Bmp2Icon.frx":1656B
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":16875
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   10
         ToolTipText     =   " Capture CLIENT Area "
         Top             =   60
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1920
         MouseIcon       =   "Bmp2Icon.frx":16E8F
         MousePointer    =   99  'Custom
         Picture         =   "Bmp2Icon.frx":17199
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   22
         TabIndex        =   9
         ToolTipText     =   " Capture ENTIRE Screen "
         Top             =   60
         Width           =   330
      End
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   5985
      Left            =   9330
      MouseIcon       =   "Bmp2Icon.frx":177B3
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1185
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   -45
      MouseIcon       =   "Bmp2Icon.frx":17ABD
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   6960
      Width           =   9600
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   8640
      MouseIcon       =   "Bmp2Icon.frx":17DC7
      MousePointer    =   99  'Custom
      Picture         =   "Bmp2Icon.frx":180D1
      ScaleHeight     =   120
      ScaleWidth      =   510
      TabIndex        =   5
      ToolTipText     =   " Move "
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   6960
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      ToolTipText     =   " Captured Image "
      Top             =   555
      Width           =   510
   End
   Begin VB.PictureBox Pic 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Height          =   5475
      Left            =   210
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   0
      Top             =   1440
      Width           =   9075
      Begin VB.PictureBox Pictureimage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         DrawMode        =   6  'Mask Pen Not
         DrawStyle       =   3  'Dash-Dot
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         MouseIcon       =   "Bmp2Icon.frx":18453
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   1
         ToolTipText     =   " Right Click To Move Picture "
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComDlg.CommonDialog CDbmp 
      Left            =   9960
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bmp2Icon Ver 1.0 - Best Computing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   75
      Width           =   3030
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   930
      Left            =   6240
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Captured Image"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6525
      TabIndex        =   3
      Top             =   1170
      Width           =   1365
   End
End
Attribute VB_Name = "Bmp2Icon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, Ip As Any) As Long
 Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12
Dim blnAboveVer4 As Boolean
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long
    
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long
    
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateIconIndirect Lib "user32" (icoinfo As ICONINFO) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As _
     PictDesc, riid As Guid, ByVal fown As Long, ipic As IPicture) As Long
     
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, _
     icoinfo As ICONINFO) As Long
     
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
    ByVal crColor As Long) As Long
    
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight _
    As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
    
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hBmMask As Long
    hbmColor As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Const Ttl = "Bmp2Icon"
Const PICTYPE_BITMAP = 1
Const PICTYPE_ICON = 3
Dim IGuid As Guid
Dim hdcMono
Dim bmpMono
Dim bmpMonoTemp
Const stdW = 32
Const stdH = 32
Dim mresult
Dim x1 As Single, y1 As Single, X2 As Single, Y2 As Single
Dim NoPicFlag As Boolean, RegionFlag As Boolean
Dim MYMouseDownX, MYMouseDownY As Integer
Dim MYNewLeft, MYNewTop As Integer
Dim MYVScrollMax, MYHScrollMax  As Integer
Dim MYVScrollMin, MYHScrollMin  As Integer
' Button parameter masks
Const LEFT_BUTTON = 1
Const RIGHT_BUTTON = 2
Const MIDDLE_BUTTON = 4


Public Sub Set_Scrolls()
If Pictureimage.width > Pic.width Then
          HScroll1.Enabled = True
           HScroll1.Max = Pictureimage.width - Pic.width
           HScroll1.Min = 1
           HScroll1.SmallChange = 1
           HScroll1.LargeChange = 100
       Else
          HScroll1.Enabled = False
       End If
If Pictureimage.height > Pic.height Then
           VScroll1.Enabled = True
           VScroll1.Max = Pictureimage.height - Pic.height
           VScroll1.Min = 1
           VScroll1.SmallChange = 1
           VScroll1.LargeChange = 100
       Else
           VScroll1.Enabled = False
       End If
End Sub

Private Sub captr()
MYHScrollMax = -(Pictureimage.width - Pic.width)
  
    MYVScrollMax = -(Pictureimage.height - Pic.height)
    MYHScrollMin = 0
    MYVScrollMin = 0
Pictureimage.Top = 0: Pictureimage.Left = 0
Bmp2Icon.WindowState = 0
Shell_NotifyIcon NIM_DELETE, IconData
Bmp2Icon.Show
End Sub



Private Sub ScreenCapture() 'capture icon sized
 Picture4.Cls
 Label3.Enabled = False
 Dim hDt&
    DoEvents
    hDt = GetDesktopWindow()
    hDtDc = GetDC(hDt)
    IconCapture.Show 1
    ReleaseDC hDt, hDtDc
    
End Sub



Private Sub mnuFileOpen_Click()
 
 On Error Resume Next
    CDbmp.DialogTitle = Ttl & "  :  Select the Picture to Open"
    CDbmp.Filter = "All Image Files (*.bmp;*.jpg;*.gif;*.ico;*.dib;*.wmf)|*.bmp;*.jpg;*.gif;*.ico;*.dib;*.wmf|Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg|GIF (*.gif)|*.gif|Icon (*.ico)|*.ico|DIB (*.dib)|*.dib|Windows Meta File (*.wmf)|*.wmf"
    CDbmp.Flags = 4 + cdlOFNFileMustExist
    CDbmp.Filename = "Image.bmp"
    
    CDbmp.ShowOpen
   If Error = cdlCancel Then _
       Exit Sub
    Pictureimage.Left = 0: Pictureimage.Top = 0
    Pictureimage.Picture = LoadPicture(CDbmp.Filename)
 
   
    Set_Scrolls
      
       HScroll1.Value = 1
       VScroll1.Value = 1
 End Sub


Private Sub bitsave_Click(Index As Integer)
Select Case Index
Case 0
 On Error Resume Next
 CDbmp.Filter = "Icon Files (*.ico)|*.ico"
 CDbmp.Flags = 4 + cdlOFNOverwritePrompt
 CDbmp.DialogTitle = "Bmp2Icon ( 1 Bit ) -  Type An Icon Name To Save as"
 CDbmp.Filename = "Icon001.ico"
 CDbmp.ShowSave
 If Err = cdlCancel Then Exit Sub
 IconX.SavePicToIcon Picture4.hdc, Picture4.Point(0, 0), CDbmp.Filename, Save1Bit

Case 1
 On Error Resume Next
 CDbmp.Filter = "Icon Files (*.ico)|*.ico"
 CDbmp.Flags = 4 + cdlOFNOverwritePrompt
 CDbmp.DialogTitle = "Bmp2Icon ( 4 Bit )  -  Type An Icon Name To Save as"
 CDbmp.Filename = "Icon004.ico"
 CDbmp.ShowSave
 If Err = cdlCancel Then Exit Sub
 IconX.SavePicToIcon Picture4.hdc, Picture4.Point(0, 0), CDbmp.Filename, Save4Bits

Case 2
 On Error Resume Next
 CDbmp.Filter = "Icon Files (*.ico)|*.ico"
 CDbmp.Flags = 4 + cdlOFNOverwritePrompt
 CDbmp.DialogTitle = "Bmp2Icon ( 8 Bit )  -  Type An Icon Name To Save as"
 CDbmp.Filename = "Icon008.ico"
 CDbmp.ShowSave
 If Err = cdlCancel Then Exit Sub
 IconX.SavePicToIcon Picture4.hdc, Picture4.Point(0, 0), CDbmp.Filename, Save8Bits '''"C:\Test Icon5.ico", Save8Bits

Case 3
 On Error Resume Next
 CDbmp.Filter = "Icon Files (*.ico)|*.ico"
 CDbmp.Flags = 4 + cdlOFNOverwritePrompt
 CDbmp.DialogTitle = "Bmp2Icon ( 24 Bit )  -  Type An Icon Name To Save as"
 CDbmp.Filename = "Icon0024.ico"
 CDbmp.ShowSave
 If Err = cdlCancel Then Exit Sub
 IconX.SavePicToIcon Picture4.hdc, Picture4.Point(0, 0), CDbmp.Filename, SaveTrueColors

End Select

End Sub

Private Sub bitsave_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
bitsave(0).ForeColor = &HC0FFC0
bitsave(1).ForeColor = &HFFFFFF
bitsave(2).ForeColor = &HFFFFFF
bitsave(3).ForeColor = &HFFFFFF



Case 1
bitsave(1).ForeColor = &HC0FFC0
bitsave(0).ForeColor = &HFFFFFF
bitsave(2).ForeColor = &HFFFFFF
bitsave(3).ForeColor = &HFFFFFF
Case 2
bitsave(2).ForeColor = &HC0FFC0
bitsave(0).ForeColor = &HFFFFFF
bitsave(1).ForeColor = &HFFFFFF
bitsave(3).ForeColor = &HFFFFFF
Case 3
bitsave(3).ForeColor = &HC0FFC0
bitsave(0).ForeColor = &HFFFFFF
bitsave(2).ForeColor = &HFFFFFF
bitsave(1).ForeColor = &HFFFFFF
End Select
End Sub

Private Sub buttonbase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 0
Picture2.BorderStyle = 0
Picture6.BorderStyle = 0
Picture9.BorderStyle = 0
Picture10.BorderStyle = 0
Picture5.BorderStyle = 0
Picture7.BorderStyle = 0
Picture8.BorderStyle = 0
Picture11.BorderStyle = 0
End Sub






Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
Picture18.BorderStyle = 0
Picture20.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
SelectObject bmpMono, bmpMonoTemp
    DeleteObject bmpMono
    DeleteDC hdcMono
    End
End Sub

Private Sub HScroll1_Change()
HScroll1_Scroll
End Sub

Private Sub HScroll1_GotFocus()
Pic.SetFocus
End Sub

Private Sub HScroll1_Scroll()
Pictureimage.Left = -HScroll1.Value
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


saveoptions.ZOrder 0
ReleaseCapture
      SendMessage saveoptions.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
End Sub

Private Sub lblMessage_Click()

End Sub

Private Sub Picture1_Click()
Picture1.BorderStyle = 0
 Set Pictureimage.Picture = Nothing
        Bmp2Icon.WindowState = 1
        'The user has minimized his window
        Call Shell_NotifyIcon(NIM_ADD, IconData)
          ' Add the form's icon to the tray
          Picture11.Visible = True
          Picture8.BorderStyle = 0
          Label3.Enabled = False
          Picture4.Cls
        Bmp2Icon.Hide
          ' Hide the button at the taskbar
          Capturemsgfrm.Check1.Visible = True
        Capturemsgfrm.Check2.Visible = True
        Capturemsgfrm.Label3.Visible = True
        Capturemsgfrm.Label4.Visible = True
          Capturemsgfrm.Check1.Enabled = True
        Capturemsgfrm.Check2.Enabled = True
        Capturemsgfrm.Label3.Enabled = True
        Capturemsgfrm.Label4.Enabled = True
        Capturemsgfrm.Label1.Caption = "SCREEN Capture Mode"
      Capturemsgfrm.Show
       Capturemsgfrm.lblactive.Visible = False
       Capturemsgfrm.lblscreen.Visible = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.BorderStyle = 1
End Sub

Private Sub Picture10_Click()
Picture10.BorderStyle = 0
frmAbout.Show 1
End Sub

Private Sub Picture10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture10.BorderStyle = 1
End Sub

Private Sub Picture11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture11.BorderStyle = 1
End Sub

Private Sub Picture15_Click()
Timer2.Enabled = True
End Sub

Private Sub Picture16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture16.Visible = False
End Sub

Private Sub Picture17_Click()
Timer2.Enabled = True
End Sub

Private Sub Picture18_Click()
saveoptions.Move 60, 72, 203, 20
Picture17.Visible = False
End Sub

Private Sub Picture18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture18.BorderStyle = 1
End Sub

Private Sub Picture19_Click()
Dim Msg
    Msg = "Save Changes To Current Icon Image?"
    Select Case MsgBox(Msg, vbQuestion + vbYesNoCancel, Ttl)
    Case vbYes
        Timer1.Enabled = True
    Case vbNo
        Picture5.BorderStyle = 0
        Pictureimage.Cls
        Picture4.Cls
        Unload Bmp2Icon
       
End Select



            
      
End Sub

Private Sub Picture2_Click()
 Picture2.BorderStyle = 0
 Set Pictureimage.Picture = Nothing
       Bmp2Icon.WindowState = 1
      'The user has minimized his window
       Call Shell_NotifyIcon(NIM_ADD, IconData)
          ' Add the form's icon to the tray
          Picture11.Visible = True
          Picture8.BorderStyle = 0
          Label3.Enabled = False
          Picture4.Cls
       Bmp2Icon.Hide
          ' Hide the button at the taskbar
        Capturemsgfrm.Check1.Visible = False
        Capturemsgfrm.Check2.Visible = False
        Capturemsgfrm.Label3.Visible = False
        Capturemsgfrm.Label4.Visible = False
        Capturemsgfrm.Label1.Caption = "Active CLIENT Capture Mode"
       Capturemsgfrm.Show
       Capturemsgfrm.lblactive.Visible = True
       Capturemsgfrm.lblscreen.Visible = False
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.BorderStyle = 1
End Sub

Private Sub Picture20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture20.Visible = False
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Bmp2Icon.ZOrder 0
ReleaseCapture
      SendMessage Bmp2Icon.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


Private Sub Picture5_Click()
Dim Msg
    Msg = "Save Changes To Current Icon Image?"
    Select Case MsgBox(Msg, vbQuestion + vbYesNoCancel, Ttl)
    Case vbYes
        Timer1.Enabled = True
    Case vbNo
        Picture5.BorderStyle = 0
        Pictureimage.Cls
        Picture4.Cls
        Unload Bmp2Icon
       
End Select
      

End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture5.BorderStyle = 1

End Sub

Private Sub Picture6_Click()
Picture6.BorderStyle = 0
ScreenCapture
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture6.BorderStyle = 1
End Sub

Private Sub Picture7_Click()
Picture7.BorderStyle = 0
mnuFileOpen_Click
End Sub

Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture7.BorderStyle = 1

End Sub

Private Sub Picture8_Click()
Picture8.BorderStyle = 0
Timer1.Enabled = True

End Sub

Private Sub Picture8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture8.BorderStyle = 1

End Sub

Private Sub Picture9_Click()
'To shell to your HTM
Picture9.BorderStyle = 0
 ShellExecute Me.hWnd, vbNullString, App.Path & "\yourhelpfilehere.htm", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture9.BorderStyle = 1
End Sub

Private Sub Pictureimage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
        Case RIGHT_BUTTON
           If Pictureimage = LoadPicture() Then
        Exit Sub
        Else
           Pictureimage.MousePointer = vbCustom
            MYMouseDownX = X
            MYMouseDownY = Y
        
        End If
        Case LEFT_BUTTON
       End Select
End Sub

Private Sub Pictureimage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case RIGHT_BUTTON
        MYHScrollMax = -(Pictureimage.width - Pic.width)
  
    MYVScrollMax = -(Pictureimage.height - Pic.height)
    MYHScrollMin = 0
    MYVScrollMin = 0
        MYNewLeft = Pictureimage.Left - (MYMouseDownX - X)
       
        If MYNewLeft > MYHScrollMax And MYNewLeft < 0 Then
            Pictureimage.Left = MYNewLeft
       HScroll1.Value = HScroll1.Value + MYMouseDownX - X
        End If
        
        MYNewTop = Pictureimage.Top - (MYMouseDownY - Y)
        If MYNewTop > MYVScrollMax And MYNewTop < 0 Then
            Pictureimage.Top = MYNewTop
        VScroll1.Value = VScroll1.Value + MYMouseDownY - Y
        End If
  Case LEFT_BUTTON
       End Select

End Sub



Private Sub Pictureimage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pictureimage.MousePointer = vbDefault
End Sub




Private Sub saveoptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture16.Visible = True
bitsave(0).ForeColor = &HFFFFFF
bitsave(1).ForeColor = &HFFFFFF
bitsave(2).ForeColor = &HFFFFFF
bitsave(3).ForeColor = &HFFFFFF
saveoptions.ZOrder 0
ReleaseCapture
      SendMessage saveoptions.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Private Sub Timer1_Timer()

Dim height1
For height1 = 20 To 215 Step 2

If saveoptions.height < 215 Then

saveoptions.Refresh

saveoptions.height = saveoptions.height + 2
End If
Next height1
Picture17.Visible = True
Timer1.Enabled = False

End Sub

Private Sub Timer2_Timer()
Dim height1
For height1 = 215 To 20 Step -0.04
If saveoptions.height > 20 Then

saveoptions.height = saveoptions.height - 0.04
End If
Next height1
Picture17.Visible = False
Timer2.Enabled = False
End Sub

Private Sub VScroll1_Change()
VScroll1_Scroll
End Sub

Private Sub VScroll1_GotFocus()
Pic.SetFocus
End Sub

Private Sub VScroll1_Scroll()
Pictureimage.Top = -VScroll1.Value
End Sub
