VERSION 5.00
Begin VB.Form IconCapture 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2895
   ClientLeft      =   2310
   ClientTop       =   2160
   ClientWidth     =   5070
   ControlBox      =   0   'False
   Icon            =   "IconCapture.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "IconCapture.frx":030A
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "IconCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Private w&, h&
Private Rct As RECT
Private Sub Form_Activate()

    SetRect Rct, 0, 0, w - 31, h - 31
    ClipCursor Rct

End Sub
Private Sub Form_Load()

    Move 0, 0, Screen.width, Screen.height
    DoEvents
    w = GetSystemMetrics(0)
    h = GetSystemMetrics(1)
    BitBlt hdc, 0, 0, w, h, hDtDc, 0, 0, vbSrcCopy

End Sub
Private Sub Form_MouseUp(Button%, Shift%, X!, Y!)
Bmp2Icon.Picture6.BorderStyle = 0
    With Bmp2Icon
        .PicMask.Line (0, 0)-(32, 32), 0, BF
         BitBlt .Picture4.hdc, 0, 0, 32, 32, hdc, X, Y, vbSrcCopy
   
    End With

    FreeCursor 0
    Unload Me
Bmp2Icon.Picture11.Visible = False
Bmp2Icon.Label3.Enabled = True
End Sub
