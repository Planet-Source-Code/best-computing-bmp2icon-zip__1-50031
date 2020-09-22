VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer ReDrawTimer 
      Interval        =   1400
      Left            =   6120
      Top             =   4320
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2100
      Left            =   240
      ScaleHeight     =   2100
      ScaleWidth      =   4080
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.PictureBox picOut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2760
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2760
      ScaleWidth      =   4590
      TabIndex        =   1
      Top             =   0
      Width           =   4590
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4200
         Picture         =   "frmAbout.frx":6820
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   6
         Top             =   40
         Width           =   315
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   120
         Left            =   3480
         MouseIcon       =   "frmAbout.frx":6D42
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":704C
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   5
         ToolTipText     =   " Move "
         Top             =   120
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4200
         MouseIcon       =   "frmAbout.frx":7ACE
         MousePointer    =   99  'Custom
         Picture         =   "frmAbout.frx":7DD8
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   3
         ToolTipText     =   " And The Piper Closes "
         Top             =   40
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "About / Credits"
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
         TabIndex        =   7
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   135
         Left            =   3480
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin BmptoIcon.Fader Fader1 
      Left            =   4920
      Top             =   480
      _ExtentX        =   979
      _ExtentY        =   450
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type CreditLine
    Text As String
    Bold As Boolean
End Type

Private marrLines(1 To 3000) As CreditLine 'the credits
Private mlngNumLines As Long            'how many lines of credits there are
Private msngYPos As Single              'the current Y pos

Private mlngTextAreaWidth As Long       'dimensions of the text area,
Private mlngTextAreaHeight As Long      'stored as module level variables
Private mlngTextAreaTop As Long         'to speed stuff up
Private mlngTextAreaLeft As Long
Private mlngTextAreaBottom As Long
Private mlngTextAreaRight As Long



Private Sub SetUpVariables()

    AddCreditLine "CopyRight ©2003 Best Computing", True
    AddCreditLine String$(58, "-")
    AddCreditLine "Credit's Go To The Following People: ", True
    AddCreditLine "Without Their Programming Skills", True
    AddCreditLine "This Suite Of Programmes", True
    AddCreditLine "Could Not Have Been Realized", True
    AddCreditLine String$(27, "-")
    AddCreditLine "1:"
    
    AddCreditLine "Stephan Swertvaegher", True
    AddCreditLine "Location", True
    AddCreditLine "Unknown", True
    AddCreditLine
    AddCreditLine "2:"
    
    AddCreditLine "Michael Drotar", True
    AddCreditLine "Location ", True
    AddCreditLine "Unknown", True
    AddCreditLine
    AddCreditLine "3:"
    
    AddCreditLine "Norm Cook", True
    AddCreditLine "guinn@ netjava.com", True
    AddCreditLine
    AddCreditLine "4:"
    
    AddCreditLine "Herman Lui", True
    AddCreditLine "Location", True
    AddCreditLine "Unknown", True
    AddCreditLine
    AddCreditLine "5:"
    
    AddCreditLine "KN", True
    AddCreditLine "kyriakosnicola@yahoo.com", True
    AddCreditLine
    AddCreditLine "6:"
    
    AddCreditLine "Carles P.V.", True
    AddCreditLine "carles_pv@ terra.es", True
    AddCreditLine
    AddCreditLine "7:"
   
    AddCreditLine "Stuart Pennington.", True
    AddCreditLine "Location", True
    AddCreditLine "England", True
    AddCreditLine
    AddCreditLine "8:"
    
    AddCreditLine "CarlHarvey@ Videotron.ca", True
    AddCreditLine
    AddCreditLine "9:"
   
    AddCreditLine "Loc Nquven", True
    AddCreditLine "Location", True
    AddCreditLine "Unknown", True
    AddCreditLine
    AddCreditLine "10:"
  
    AddCreditLine "Keral C Patel", True
    AddCreditLine "Location", True
    AddCreditLine "India", True
    AddCreditLine "keral82@keral.com", True
    AddCreditLine
    AddCreditLine "11:"
    
    AddCreditLine "Steve McMahon", True
    AddCreditLine "Location", True
    AddCreditLine "Vbaccelerator", True
    AddCreditLine "steve@vbaccelerator.com", True
    AddCreditLine
    AddCreditLine String$(70, "-")
    AddCreditLine "Many Thanks To All The Above", True
    AddCreditLine "Richard Best - Best Computing", True
    AddCreditLine "rabc1950@hotmail.com", True
    AddCreditLine String$(70, "-")
    AddCreditLine "END OF CREDITS", True
   
   
    
 
End Sub

Private Sub AddCreditLine(Optional pstrText As String = "", Optional pbolBold As Boolean = False)
    'bump the line count up by one
    mlngNumLines = mlngNumLines + 1
    
    'make sure they've given us something
    If Len(pstrText) > 0 Then
        With marrLines(mlngNumLines)
            .Text = pstrText
            .Bold = pbolBold
        End With
    End If

End Sub



Private Sub Form_Load()
    
    Dim lstrVersion As String
    
   
    
    'make sure the buffer picture is the same size as the back buffer picture
    picBuffer.Move 0, 0, picBackBuffer.width, picBackBuffer.height
    
    'make sure everything is dealing in pixels, not twips
    Me.ScaleMode = vbPixels
    picBuffer.ScaleMode = vbPixels
    picOut.ScaleMode = vbPixels
    picBackBuffer.ScaleMode = vbPixels
    'set a few properties of the buffer
    'picBuffer.ForeColor = vbYellow
    picBuffer.BackColor = vbWhite
    picBuffer.AutoRedraw = True
    'hide the buffer
    picBuffer.Visible = False
    
    'grab the dimensions of the background area
    mlngTextAreaHeight = picBackBuffer.height
    mlngTextAreaLeft = picBackBuffer.Left
    mlngTextAreaTop = picBackBuffer.Top
    mlngTextAreaWidth = picBackBuffer.width
    mlngTextAreaBottom = mlngTextAreaTop + mlngTextAreaHeight
    mlngTextAreaRight = mlngTextAreaLeft + mlngTextAreaWidth
    
    'copy a chunk of the main picture to the background buffer *before* we start drawing on the main picture
    BitBlt picBackBuffer.hdc, 0, 0, mlngTextAreaWidth, mlngTextAreaHeight, picOut.hdc, mlngTextAreaLeft, mlngTextAreaTop, SRCCOPY
    
    'set the initial horizontal drawing position to be about 1/4 the way down the drawing area
    'this gives the user time to go "huh? its scrolling? whats that first line say?" before the first line disappears
    msngYPos = CLng(mlngTextAreaHeight * (1 / 4))
    
    'setup the text and stuff to display
    SetUpVariables
    
    '20
    ReDrawTimer.Interval = 20 'speed
    ReDrawTimer.Enabled = True

End Sub

Private Sub Form_Activate()
    Fader1.FadeIn
    ReDrawTimer.Enabled = True
    'draw the credits when they swap back to this form
    DrawCredits
End Sub

Private Sub DrawCredits()
    Dim llngCount As Long
    Dim llngFontSize As Long
    
    'not a whole lot we can do about errors
    On Error Resume Next
    
    'Draw the background to the buffer. It's only had to be written once, so we'll just re-blit it over again and agin.
    BitBlt picBuffer.hdc, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBackBuffer.hdc, 0, 0, SRCCOPY
    
    'remember where we are supposed to start from
    picBuffer.CurrentY = msngYPos
    
    'Do the following for each line of text in our credits message...
    For llngCount = 1 To mlngNumLines
        'set bolding based on what the array says
        picBuffer.FontBold = marrLines(llngCount).Bold
        
        'Set the starting location of where to print the text
        picBuffer.CurrentY = picBuffer.CurrentY + 2 '(this is to bump the line spacing a bit)
        picBuffer.CurrentX = (picBuffer.ScaleWidth - picBuffer.TextWidth(marrLines(llngCount).Text)) / 2
        
        'Send the text to the buffer now
        picBuffer.Print marrLines(llngCount).Text
    Next
    
    'Ok, now that we have painted the entire buffer as we see fit for this pass, we blast the entire
    'finished image directly to our output picturebox control.
    BitBlt picOut.hdc, mlngTextAreaLeft, mlngTextAreaTop, mlngTextAreaWidth, mlngTextAreaHeight, picBuffer.hdc, 0, 0, SRCCOPY
    
   
    
    'force a refresh of pic out
    picOut.Refresh
    
    If picBuffer.CurrentY < -5 Then
       
        'if the last line is above the top there's no more text to scroll
        'and its time to reset the draw position to the height of the text area
        msngYPos = mlngTextAreaHeight
    Else
        
'still some room left to go up, move up the text area by a pixel
        msngYPos = msngYPos - 1
    
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = True
End Sub

Private Sub Picout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ReDrawTimer.Enabled = True Then
ReDrawTimer.Enabled = False
Else
ReDrawTimer.Enabled = True
End If
End Sub

Private Sub Picout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = False
Picture3.Visible = True
End Sub

Private Sub Picture1_Click()
ReDrawTimer.Enabled = False
Fader1.FadeOut
frmAbout.Hide

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

frmAbout.ZOrder 0

ReleaseCapture
      
      SendMessage frmAbout.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
End Sub

Private Sub ReDrawTimer_Timer()
    DrawCredits
End Sub


