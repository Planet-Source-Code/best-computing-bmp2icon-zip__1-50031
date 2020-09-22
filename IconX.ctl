VERSION 5.00
Begin VB.UserControl IconX 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   InvisibleAtRuntime=   -1  'True
   Picture         =   "IconX.ctx":0000
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   47
   ToolboxBitmap   =   "IconX.ctx":0ECA
   Begin VB.PictureBox picTemp 
      BackColor       =   &H00404080&
      Height          =   2055
      Left            =   2880
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   2400
      Width           =   2535
   End
End
Attribute VB_Name = "IconX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Private TempVal As String                 ' Temp Value, very useful
Private TempVal2 As String                 ' Temp Value, very useful

  

Public Function SavePicToIcon(PicturehDC As Long, TransperantColor As Long, TargetFileName As String, SaveBits As IconSaveTypes) As Boolean
ExportPicToIcon PicturehDC, TransperantColor, TargetFileName, SaveBits
End Function
Private Sub UserControl_Paint()
UserControl.width = 32 * 15
UserControl.height = 32 * 15
End Sub

