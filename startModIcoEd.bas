Attribute VB_Name = "startModIcoEd"
Option Explicit
 
 'down Only one instance of exe
 Public Const GW_HWNDPREV = 3
 Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
      Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) _
         As Long
      Declare Function GetWindow Lib "user32" _
        (ByVal hWnd As Long, ByVal wCmd As Long) As Long
      Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As Long) As Long
 'up Only one instance of exe
'down show/hide desktop
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_SHOW = 5
Public Const SW_SHOWNORMAL As Long = 1
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWDEFAULT As Long = 10
Public Const SW_MINIMIZE As Long = 6
Public Const SW_HIDE = 0
'up show/hide desktop

'Function to get OS
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'Function to Start a external program/file
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

'Function to apply WinXP style controls
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As _
    INITCOMMONCONTROLSEX_TYPE) As Long

'For XP style controls
Public Const ICC_INTERNET_CLASSES = &H800

'For OS detection
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'For 'Start a external program/file'
Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

'For OS detection
Private Type OSVERSIONINFO
    OSVSize As Long
    dwVerMajor As Long
    dwVerMinor As Long
    dwBuildNumber As Long ' NT: Build Number, 9x: High-Order has Major/Minor ver, Low-Order has build
    PlatformID As Long
    szCSDVersion As String * 128    '#### NT: ie- "Service Pack 3", 9x: 'arbitrary additional information'
End Type

'For OS detection
Public Enum cnWin32Ver
    UnknownOS = 0
    Win95 = 1
    Win98 = 2
    WinME = 3
    WinNT4 = 4
    Win2k = 5
    WinXP = 6
End Enum

'For XP Style controls
Public Type INITCOMMONCONTROLSEX_TYPE
    dwSize As Long
    dwICC As Long
End Type





'ORIGINAL FROM HERE
Public ColorArray&(255)

Public hDtDc&

Public hGridDc&
Private hGrid&, hGridOld&

Public hTileDc&
Private hTileOld&

Public hToolsDc&
Private hToolsOld&

Public hBgDc&
Private hBg&, hBgOld&

Public MouseColors&(1)

Type IconColors
     icRed As Integer
     icGreen As Integer
     icBlue As Integer
End Type

Public Defaults(1023) As IconColors
Public EditVals(1023) As IconColors
Public OutputVals(1023) As IconColors

Type BITMAPINFOHEADER
     biSize As Long
     biWidth As Long
     biHeight As Long
     biPlanes As Integer
     biBitCount As Integer
     biCompression As Long
     biSizeImage As Long
     biXPelsPerMeter As Long
     biYPelsPerMeter As Long
     biClrUsed As Long
     biClrImportant As Long
End Type

Type ICONDIR
     idReserved As Integer
     idType As Integer
     idCount As Integer
End Type

Type ICONDIRENTRY
     bWidth As Byte
     bHeight As Byte
     bColorCount As Byte
     bReserved As Byte
     wPlanes As Integer
     wBitCount As Integer
     dwBytesInRes As Long
     dwImageOffset As Long
End Type



Public BIH As BITMAPINFOHEADER
Public ID As ICONDIR
Public IDE As ICONDIRENTRY

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function ClipCursor Lib "user32" (lpRect As RECT) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Declare Function FreeCursor Lib "user32" Alias "ClipCursor" (ByVal Zero As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal crColor As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function CreateIconFromResource Lib "user32" _
   (ByVal presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, _
    ByVal dwVer As Long) As Long
Declare Function DrawIcon Lib "user32" _
   (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) _
    As Long


Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegCreateKeyEx& Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal Reserved&, ByVal lpClass$, ByVal dwOptions&, ByVal samDesired&, ByVal SecAtts&, phkResult&, lpdwDisp&)
Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey&, ByVal lpValueName$)
Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
Declare Function RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, lpReserved&, lpType&, ByVal lpData$, lpcbData&)
Declare Function RegSetValueEx& Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal Reserved&, ByVal dwType&, ByVal lpData$, ByVal cbData&)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_ALL_ACCESS = (&H1F0000 Or &H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H20) And (Not &H100000)
Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'These three constants specify what you want to do
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const DI_NORMAL = &H3
'Public Const RGN_OR = 2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Type newhdr  'GENERAL HEADER ***FIRST THING WHAT WE GOT .ICO OR .CUR
 Reserved As Integer  'MUST BE 0
 restype As Integer  '1---ICON,2---CURSOR
 rescount As Integer 'HOW MANY ICON/CURSOR
End Type
Public Type iconresdir
 width As Byte
 height As Byte
 colorcount As Byte
 Reserved As Byte  'MUST BE 0
End Type

Public Type resdir
 ICONDIR As iconresdir
 planes As Integer
 bitcount As Integer
 bytesinres As Long
 entrypoint As Long
End Type
Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public IconData As NOTIFYICONDATA
Public Const RGN_OR = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const Ttl = "Launcher"
Private RFL$(1 To 4)
Private K%, M%, Rv&, P%, Q%
Private Buffer$, hKey&, PaFn$, TempFile$




Private Sub Main()
        Dim comctls As INITCOMMONCONTROLSEX_TYPE
    Dim retval As Long
    Dim CanProceed As Boolean
    
    CanProceed = IsManifestFile 'CanProceed = True if MANIFEST FILE EXISTS
    'Hence - No need to auto restart the application
    If Val(Win32Ver) > 5 Then 'IF WINDOWS XP
        If MakeMANIFESTfile Then 'IF MANIFEST FILE WAS SUCCESSFULLY CREATED
            
            'Apply XP Style Controls
            With comctls
                .dwSize = Len(comctls)
                .dwICC = ICC_INTERNET_CLASSES
            End With
            retval = InitCommonControlsEx(comctls)
            '#######################
            
        Else
            'Manifest file -->wasn't<--- successfully created
            'Hence - can't apply xp style controls
            'Hence - No need to auto restart the application
            CanProceed = True
        End If
    Else
        'Not WinXP
        'Hence - can't apply xp style controls
        'Hence - No need to auto restart the application
        CanProceed = True
    End If
    If CanProceed Then   'Can continue with this exe session.
    'BestComputing
    'SetUpIconDblClick
        
        Bmp2Icon.Show
        
        
    Else
    
    
        'The application needs to be auto restarted
    
        'USE THIS CODE IF YOU WANT ONE INSTANCE OF YOUR PROGRAM TO RUN:
        SaveSetting App.EXEName, "Settings", "CanRun", "YES"
        '###################################################
        
        'START THE EXE FILE AGAIN: (Shelldocument = True if program was successfully
        'restarted // False if it wasn't
If ShellDocument(App.Path & "\" & App.EXEName & ".exe", , , , START_NORMAL) Then
            End ' End this program --> New one auto started
            
        Else 'Wasn't able to start exe file (Proceed in current exe session)
        
            'USE THIS CODE IF YOU WANT ONE INSTANCE OF YOUR PROGRAM TO RUN:
            SaveSetting App.EXEName, "Settings", "CanRun", "NO"
            '##################################################
        'BestComputing
        'SetUpIconDblClick
        Bmp2Icon.Show
            
        End If
        
        
        
    End If
        


   

End Sub
'#################################################################
'THE BELOW CODE CREATES A MANIFEST FILE:
'MakeMANIFESTfile returns True if it was able to create the file
'THIS CODE IS TO RUN AN EXTERNAL APPLICATION / DOCUMENT:

'SHELLDOCUMENT will return a True if the file was run
'Else False if not

Public Function ShellDocument(sDocName As String, _
                    Optional ByVal Action As String = "Open", _
                    Optional ByVal Parameters As String = vbNullString, _
                    Optional ByVal Directory As String = vbNullString, _
                    Optional ByVal WindowState As StartWindowState) As Boolean
    Dim Response
    Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
    Select Case Response
        Case Is < 33
            ShellDocument = False
        Case Else
            ShellDocument = True
    End Select
End Function







'THE BELOW CODE IS TO DETECT OPERATING SYSTEM
'WE USE THIS TO SEE IF THE USER IS RUNNING WIN XP

'# Public subs/functions
'# Returns the asso. cnWin32Ver eNum value of the current Win32 OS

Public Function Win32Ver() As cnWin32Ver
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)
   
        '#### If the API returned a valid value
    If GetVersionEx(oOSV) = 1 Then
        
            '#### If we're running WinXP
            '####    If VER_PLATFORM_WIN32_NT, dwVerMajor is 5 and dwVerMinor is 1, it's WinXP
        If (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1) Then
           Win32Ver = WinXP

            '#### If we're running WinNT2000 (NT5)
            '####    If VER_PLATFORM_WIN32_NT, dwVerMajor is 5 and dwVerMinor is 0, it's Win2k
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0) Then
           Win32Ver = Win2k

            '#### If we're running WinNT4
            '####    If VER_PLATFORM_WIN32_NT and dwVerMajor is 4
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 4) Then
           Win32Ver = WinNT4

            '#### If we're running Windows ME
            '####    If VER_PLATFORM_WIN32_WINDOWS and
            '####    dwVerMajor = 4,  and dwVerMinor > 0, return true
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 90) Then
           Win32Ver = WinME

            '#### If we're running Win98
            '####    If VER_PLATFORM_WIN32_WINDOWS and
            '####    dwVerMajor => 4, or dwVerMajor = 4 and
            '####    dwVerMinor > 0, return true
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (oOSV.dwVerMajor > 4) Or (oOSV.dwVerMajor = 4 And oOSV.dwVerMinor > 0) Then
           Win32Ver = Win98

            '#### If we're running Win95
            '####    If VER_PLATFORM_WIN32_WINDOWS and
            '####    dwVerMajor = 4, and dwVerMinor = 0,
        ElseIf (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 0) Then
           Win32Ver = Win95

            '#### Else the OS is not reconized by this function
        Else
            Win32Ver = UnknownOS
        End If
    
        '#### Else the OS is not reconized by this function
    Else
        Win32Ver = UnknownOS
    End If
End Function


'#########################################################
'# Returns true if the OS is WinNT4, Win2k or WinXP
'#########################################################
Public Function isNT() As Boolean
        '#### Determine the return value of Win32Ver() and set the return value accordingly
    Select Case Win32Ver()
        Case WinNT4, Win2k, WinXP
            isNT = True
        Case Else
            isNT = False
    End Select
End Function


'#########################################################
'# Returns true if the OS is Win95, Win98 or WinME
'#########################################################
Public Function is9x() As Boolean
        '#### Determine the return value of Win32Ver() and set the return value accordingly
    Select Case Win32Ver()
        Case Win95, Win98, WinME
            is9x = True
        Case Else
            is9x = False
    End Select
End Function


'#########################################################
'# Returns true if the OS is WinXP
'#########################################################
Public Function isWinXP() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinXP = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 1)
    End If
End Function


'#########################################################
'# Returns true if the OS is Win2k
'#########################################################
Public Function isWin2k() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWin2k = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 5 And oOSV.dwVerMinor = 0)
    End If
End Function


'#########################################################
'# Returns true if the OS is WinNT4
'#########################################################
Public Function isWinNT4() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinNT4 = (oOSV.PlatformID = VER_PLATFORM_WIN32_NT And oOSV.dwVerMajor = 4)
    End If
End Function


'#########################################################
'# Returns true if the OS is WinME
'#########################################################
Public Function isWinME() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
        isWinME = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 90)
    End If
End Function



'# Returns true if the OS is Win98

Public Function isWin98() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
         isWin98 = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And (oOSV.dwVerMajor > 4) Or (oOSV.dwVerMajor = 4 And oOSV.dwVerMinor > 0)
    End If
End Function



'# Returns true if the OS is Win95

Public Function isWin95() As Boolean
    Dim oOSV As OSVERSIONINFO
    oOSV.OSVSize = Len(oOSV)

        '#### If the API returned a valid value
    If (GetVersionEx(oOSV) = 1) Then
         isWin95 = (oOSV.PlatformID = VER_PLATFORM_WIN32_WINDOWS And oOSV.dwVerMajor = 4 And oOSV.dwVerMinor = 0)
    End If
End Function
'#################################################################################



Public Property Get MakeMANIFESTfile() As Boolean
    
    MakeMANIFESTfile = False
    
    On Local Error GoTo MakeMANIFESTfile_Err
    
    Dim ManifestFileName As String
    Dim NewFreeFile As Integer
    
    ManifestFileName = App.Path & "\" & App.EXEName & ".exe.MANIFEST"
    NewFreeFile = FreeFile
    
    'Note:  CHR(34)   =   "
    
    Open ManifestFileName For Output As NewFreeFile
        Print #NewFreeFile, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
        Print #NewFreeFile, "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">"
        Print #NewFreeFile, "<assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "x86" & Chr(34) & " name=" & Chr(34) & "prjThemed" & Chr(34) & " type=" & Chr(34) & "Win32" & Chr(34) & " />"
        Print #NewFreeFile, "<dependency>"
        Print #NewFreeFile, "<dependentAssembly>"
        Print #NewFreeFile, "<assemblyIdentity type=" & Chr(34) & "Win32" & Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "x86" & Chr(34) & " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & " language=" & Chr(34) & "*" & Chr(34) & " />"
        Print #NewFreeFile, "</dependentAssembly>"
        Print #NewFreeFile, "</dependency>"
        Print #NewFreeFile, "</assembly>"
    Close NewFreeFile
    
    MakeMANIFESTfile = True 'FILE CREATED OK
    
    Exit Property
MakeMANIFESTfile_Err:
    MakeMANIFESTfile = False 'ERROR CREATING FILE
End Property
'THE BELOW CODE CHECK THE EXISTANCE OF THE MANIFEST FILE:
'IF FILE EXISTS THEN IsManifestFile returns True

'THIS CODE SIMPLY TRIES TO OPEN THE FILE -
'IF THERE IS AN ERROR THEN THE FILE DOESN'T EXIST

Public Property Get IsManifestFile() As Boolean
    
    IsManifestFile = False
    
    On Local Error GoTo IsManifestFile_Err
    
    Dim ManifestFileName As String
    Dim NewFreeFile As Integer
    
    ManifestFileName = App.Path & "\" & App.EXEName & ".exe.MANIFEST"
    NewFreeFile = FreeFile
    
    Open ManifestFileName For Input Access Read As NewFreeFile
    Close NewFreeFile
    
    IsManifestFile = True 'FILE DOES EXIST
    
    Exit Property
    
IsManifestFile_Err:
    IsManifestFile = False 'FILE DOESN'T EXIST

End Property


Public Sub CreateGrid(Pic As PictureBox)

    Dim K%

    For K = 0 To 321 Step 10
        Pic.Line (K, 0)-(K, 321)
        Pic.Line (0, K)-(321, K)
    Next

    hGridDc = CreateCompatibleDC(0)
    hGrid = CreateCompatibleBitmap(Pic.hdc, 321, 321)
    hGridOld = SelectObject(hGridDc, hGrid)
    BitBlt hGridDc, 0, 0, 321, 321, Pic.hdc, 0, 0, vbSrcCopy

End Sub

Public Sub CreateTools()

    hToolsDc = CreateCompatibleDC(0)
    hToolsOld = SelectObject(hToolsDc, LoadResPicture(20, 0))

End Sub
Public Sub DestroyTools()

    DeleteObject SelectObject(hToolsDc, hToolsOld)
    DeleteDC hToolsDc

End Sub
Public Sub DestroyGrid()

    DeleteObject SelectObject(hGridDc, hGridOld)
    DeleteDC hGridDc

End Sub
Public Sub DestroyTile()

    DeleteObject SelectObject(hTileDc, hTileOld)
    DeleteDC hTileDc

    DeleteObject SelectObject(hBgDc, hBgOld)
    DeleteDC hBgDc

End Sub
Public Sub AdjustToEdge(P%, Q%)

    P = (P \ 10) * 10
    Q = (Q \ 10) * 10

End Sub
Public Sub AdjustToNearestEdge(P%, Q%)

   If P Mod 10 >= 5 Then
       P = (P \ 10) * 10 + 10
   Else
       P = (P \ 10) * 10
   End If

   If Q Mod 10 >= 5 Then
      Q = (Q \ 10) * 10 + 10
   Else
      Q = (Q \ 10) * 10
   End If

End Sub
Public Sub AdjustToCentre(P%, Q%)

    P = (P \ 10) * 10 + 5
    Q = (Q \ 10) * 10 + 5

End Sub
Public Sub ConfineCoords(P%, Q%)

    If P < 0 Then P = 0
    If P > 320 Then P = 320
    If Q < 0 Then Q = 0
    If Q > 320 Then Q = 320

End Sub


Public Sub PrepIconHeader()

    ID.idReserved = 0
    ID.idType = 1
    ID.idCount = 1

    IDE.bWidth = 32
    IDE.bHeight = 32
    IDE.bColorCount = 0
    IDE.bReserved = 0
    IDE.wPlanes = 1
    IDE.wBitCount = 24
    IDE.dwBytesInRes = 3240
    IDE.dwImageOffset = 22

    BIH.biSize = 40
    BIH.biWidth = 32
    BIH.biHeight = 64
    BIH.biPlanes = 1
    BIH.biBitCount = 24
    BIH.biCompression = 0
    BIH.biSizeImage = 3200
    BIH.biXPelsPerMeter = 0
    BIH.biYPelsPerMeter = 0
    BIH.biClrUsed = 0
    BIH.biClrImportant = 0

End Sub

