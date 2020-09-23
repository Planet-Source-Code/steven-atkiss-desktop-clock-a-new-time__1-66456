Attribute VB_Name = "ModDTC_ProMo"
Option Explicit

'Skin Procedure Flow
'1   New Skin        Loads Skin And Sets Windows Dimensions
'2   Trans Colour    Set The Alarm Image Trans Colour
'3   Draw Time       Draw The Next Time Interval On The Canvas
'4   Re-Draw Clock   Copy The Canvas Image To The Display Image And Re-Skin The Form

Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFOEX) As Long


'----[ API's ]----'
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long
Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205


Public TrayI As NOTIFYICONDATA




'----[ Constants ]----'
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_BADDB = 1009&
Public Const ERROR_BADKEY = 1010&
Public Const ERROR_CANTOPEN = 1011&
Public Const ERROR_CANTREAD = 1012&
Public Const ERROR_CANTWRITE = 1013&
Public Const ERROR_OUTOFMEMORY = 14&
Public Const ERROR_INVALID_PARAMETER = 87&
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const ERROR_MORE_DATA = 234&
Public Const KEY_QUERY_VALUE = &H1&
Public Const KEY_SET_VALUE = &H2&
Public Const KEY_CREATE_SUB_KEY = &H4&
Public Const KEY_ENUMERATE_SUB_KEYS = &H8&
Public Const KEY_NOTIFY = &H10&
Public Const KEY_CREATE_LINK = &H20&
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const SUREGKEY = "Software\Microsoft\Windows\CurrentVersion\Run"

'----[ Enums ]----'
Public Enum rcMainKey       'root keys constants
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum


Public Enum rcRegType       'data types constants
    REG_NONE = 0
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_LINK = 6
    REG_MULTI_SZ = 7
    REG_RESOURCE_LIST = 8
    REG_FULL_RESOURCE_DESCRIPTOR = 9
    REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

'----[ Dim's ]----'
Private HKey             As Long
Private mainKey          As Long
Private sKey             As String
Private lBufferSize      As Long
Private lDataSize        As Long
Private ByteArray()      As Byte
Private createNoExists   As Boolean

    Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
    End Type


    Public Type OSVERSIONINFOEX
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
        wServicePackMajor As Integer
        wServicePackMinor As Integer
        wSuiteMask As Integer
        wProductType As Byte
        wReserved As Byte
    End Type
    
    Public Const VER_PLATFORM_WIN32s = 0
    Public Const VER_PLATFORM_WIN32_WINDOWS = 1
    Public Const VER_PLATFORM_WIN32_NT = 2
    
    Public OSBase As String

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const RGN_OR = 2

Public Type SkinA
    Skin As String
    Width As Long
    Height As Long
    DigitX As Long
    DigitY As Long
End Type

Public Type ClockA
    X As Long
    Y As Long
    DispAlarm As Boolean
    RemLocation As Boolean
    StartWithWindows As Boolean
    AlwaysOntop As Boolean
End Type

Public Type AlarmA
    AlarmTime As String
    AlarmSound As String
    AlarmSet As Boolean
    AlarmType As Byte
    AlarmNote As String
    AlarmActivated As Boolean
    SnoozeTime As String
    AlarmVolume As Integer
End Type

Public Type Attribs
    Clock As ClockA
    Skin As SkinA
    Alarm As AlarmA
End Type

Public Attribs As Attribs
Public SymTransCol As Long
Public ClkTransCol As Long

Public FSys As New FileSystemObject
Public TempSkin As String, TempSound As String, TempString As String
Public REGKEY As String
Public MaskPrepared As Boolean
Public TimeCheck As Boolean

Public SnoozeTime As Single

Public Remain As String

Public Function CreateKey(ByVal sPath As String) As Long
    
    HKey = GetKeys(sPath, sKey) 'get keys
    
    'try to create key
    If (RegCreateKey(HKey, sKey, mainKey) = ERROR_SUCCESS) Then
        RegCloseKey mainKey
        CreateKey = mainKey 'success
    Else
        CreateKey = 0 'error
    End If
    
    RegCloseKey HKey
    
End Function

Public Function KeyExists(ByVal sPath As String) As Boolean

    HKey = GetKeys(sPath, sKey)
    
    'try to open key
    If (RegOpenKeyEx(HKey, sKey, 0, KEY_ALL_ACCESS, mainKey) = ERROR_SUCCESS) Then
        KeyExists = True 'if we open it than it exists ;o)
        RegCloseKey mainKey 'close key
    Else
        KeyExists = False ' noup, the key don't exists
    End If

    RegCloseKey HKey
End Function
Public Function GetREGSZVal(Key As String, PropertyName As String) As String
On Error Resume Next

    Dim HKey As Long
    Dim sPath As String
    Dim sKey As String
    Dim C As Long
    Dim r As Long
    Dim S As String
    Dim T As Long
    
    HKey = GetKeys(Key, sKey)
    
    r = RegOpenKeyEx(HKey, sKey, 0, KEY_READ, HKey)

    C = 255
    S = String(C, Chr(0))
    r = RegQueryValueEx(HKey, PropertyName, 0, T, S, C)
   
    GetREGSZVal = StripNonChar(Trim(Left(S, C - 1)))
    
    RegCloseKey HKey
    
End Function








Public Function GetKeys(sPath As String, sKey As String) As rcMainKey
Dim pos As Long, mk As String
    
    'replace long with short root constants
    sPath = Replace$(sPath, "HKEY_CURRENT_USER", "HKCU", , , 1)
    sPath = Replace$(sPath, "HKEY_LOCAL_MACHINE", "HKLM", , , 1)
    sPath = Replace$(sPath, "HKEY_CLASSES_ROOT", "HKCR", , , 1)
    sPath = Replace$(sPath, "HKEY_USERS", "HKUS", , , 1)
    sPath = Replace$(sPath, "HKEY_PERFORMANCE_DATA", "HKPD", , , 1)
    sPath = Replace$(sPath, "HKEY_DYN_DATA", "HKDD", , , 1)
    sPath = Replace$(sPath, "HKEY_CURRENT_CONFIG", "HKCC", , , 1)
    
    pos = InStr(1, sPath, "\") 'get pos of first slash

    If (pos = 0) Then 'writting to root
        mk = UCase$(sPath)
        sKey = ""
    Else
        mk = UCase$(Left$(sPath, 4)) 'get hkey
        sKey = Right$(sPath, Len(sPath) - pos) 'get path
    End If
    
    Select Case mk 'return main key handle
        Case "HKCU": GetKeys = HKEY_CURRENT_USER
        Case "HKLM": GetKeys = HKEY_LOCAL_MACHINE
        Case "HKCR": GetKeys = HKEY_CLASSES_ROOT
        Case "HKUS": GetKeys = HKEY_USERS
        Case "HKPD": GetKeys = HKEY_PERFORMANCE_DATA
        Case "HKDD": GetKeys = HKEY_DYN_DATA
        Case "HKCC": GetKeys = HKEY_CURRENT_CONFIG
    End Select
    
End Function

Public Function WriteString(ByVal sPath As String, ByVal sName As String, _
                                                   ByVal sValue As String) As Long
                            

    If (KeyExists(sPath) = False) Then 'if key don't exists,
        If (createNoExists = False) Then 'and if CreateKeyIfDoesntExists = True
            CreateKey sPath  ' then create it ;o)
        Else
            WriteString = 0 'error!
            Exit Function
        End If
    End If
    
    HKey = GetKeys(sPath, sKey) 'parse keys
    
    If (sName = "@") Then sName = "" '(Default)
    
    'try to open key
    If (RegOpenKeyEx(HKey, sKey, 0, KEY_WRITE, mainKey) = ERROR_SUCCESS) Then
        'try to write data
        If (RegSetValueEx(mainKey, sName, 0, REG_SZ, ByVal sValue, Len(sValue)) = ERROR_SUCCESS) Then
            RegCloseKey mainKey 'close key
            WriteString = mainKey 'success!
        Else
            WriteString = 0 'error writting data
      End If
    Else
         WriteString = 0 'error opening key
    End If
    
End Function

Function SaveDword(ByVal HKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(HKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function

Public Sub NewSkin()
    
    FrmMain.Visible = False
    FrmMain.Timer.Enabled = False
    
    If Attribs.Skin.Skin <> "Default" And Attribs.Skin.Skin <> "" Then
        FrmMain.PicSkin.Picture = LoadPicture(Attribs.Skin.Skin)
    End If
    
    Attribs.Skin.DigitX = FrmMain.PicSkin.ScaleWidth / 11
    Attribs.Skin.DigitY = FrmMain.PicSkin.ScaleHeight
    
    With FrmMain
        .PicCanvas.Picture = LoadPicture("")
        .PicCanvas.Cls
    
        .PicCanvas.Width = ((Attribs.Skin.DigitX * 8) + (.PicAlarm.ScaleWidth / 2)) * Screen.TwipsPerPixelX
        .PicCanvas.Height = .PicSkin.Height + .PicAlarm.Height
        
        .PicDisplay.ScaleWidth = .PicCanvas.ScaleWidth
        .PicDisplay.Height = .PicCanvas.Height
    End With
    
    SetToTransCol
        FrmMain.PicCanvas.BackColor = ClkTransCol
    DrawTime
    
End Sub

Public Sub SetToTransCol()
    
    Dim X As Single, Y As Single
    
    With FrmMain
    
    SymTransCol = .PicAlarm.Point(0, 0)
    ClkTransCol = .PicSkin.Point(0, 0)
    
    For X = 0 To .PicAlarm.ScaleWidth
        For Y = 0 To .PicAlarm.ScaleHeight
            If .PicAlarm.Point(X, Y) = SymTransCol Then
                SetPixelV .PicAlarm.hdc, X, Y, ClkTransCol
            Else
                SetPixelV .PicAlarm.hdc, X, Y, .PicAlarm.Point(X, Y)
            End If
        Next Y
    Next X
    
    .PicAlarm.Picture = .PicAlarm.Image
    .PicAlarm.Refresh
    
    End With
    
End Sub

Public Sub DrawTime()

    Dim TPos As Single
    Dim TIncr As String
    TIncr = TimeSerial(Hour(Now), Minute(Now), Second(Now) + 1)
    
    With FrmMain
        .PicCanvas.Cls
        .PicCanvas.Picture = LoadPicture("")
        For TPos = 1 To Len(TIncr)
            If Mid$(TIncr, TPos, 1) = ":" Then
                BitBlt .PicCanvas.hdc, ((TPos - 1) * Attribs.Skin.DigitX) + (.PicAlarm.ScaleWidth / 2), 0, Attribs.Skin.DigitX, Attribs.Skin.DigitY, .PicSkin.hdc, Attribs.Skin.DigitX * 10, 0, vbSrcCopy
            Else
                BitBlt .PicCanvas.hdc, ((TPos - 1) * Attribs.Skin.DigitX) + (.PicAlarm.ScaleWidth / 2), 0, Attribs.Skin.DigitX, Attribs.Skin.DigitY, .PicSkin.hdc, Attribs.Skin.DigitX * Val(Mid$(TIncr, TPos, 1)), 0, vbSrcCopy
            End If
        Next TPos
        
        If Attribs.Clock.DispAlarm = True Then
            BitBlt .PicCanvas.hdc, 0, 0, .PicAlarm.ScaleWidth / 2, .PicAlarm.ScaleHeight + 2, .PicAlarm.hdc, (.PicAlarm.ScaleWidth / 2) * Cbl(Attribs.Alarm.AlarmSet), 0, vbSrcCopy
        End If
        
        .PicCanvas.Picture = .PicCanvas.Image
        
    End With
    
    MaskPrepared = True
    
    
    
End Sub


Public Sub SkinForm(Frm As Form, MaskPic As PictureBox, Optional TransColor As Long)
    
    Dim Retr As Long
    Dim RgnFinal As Long
    Dim RgnTmp As Long
    
    Dim hHeight As Long
    Dim wWidth As Long
    
    Dim Col As Long
    Dim Start As Long
    Dim RowR As Long
    
    MaskPic.AutoSize = True
    MaskPic.AutoRedraw = True

    With Frm
        .Height = MaskPic.Height
        .Width = MaskPic.Width
    End With
    
    If TransColor < 1 Then
        TransColor = GetPixel(MaskPic.hdc, 0, 0)
    End If
    
    hHeight = MaskPic.Height / Screen.TwipsPerPixelY
    wWidth = MaskPic.Width / Screen.TwipsPerPixelX
    
    RgnFinal = CreateRectRgn(0, 0, 0, 0)

    For RowR = 0 To hHeight - 1
        Col = 0
        Do While Col < wWidth
            Do While Col < wWidth And GetPixel(MaskPic.hdc, Col, RowR) = TransColor
                Col = Col + 1
            Loop
            If Col < wWidth Then
                Start = Col
                Do While Col < wWidth And GetPixel(MaskPic.hdc, Col, RowR) <> TransColor
                    Col = Col + 1
                Loop
                
                If Col > wWidth Then Col = wWidth
                RgnTmp = CreateRectRgn(Start, RowR, Col, RowR + 1)
                Retr = CombineRgn(RgnFinal, RgnFinal, RgnTmp, RGN_OR)
                DeleteObject (RgnTmp)
            End If
        Loop
    Next RowR
    
    Retr = SetWindowRgn(Frm.hWnd, RgnFinal, True)
     
End Sub

Public Function Cbl(BoolValue As Boolean) As Single
    
    If BoolValue = True Then
        Cbl = 1
    Else
        Cbl = 0
    End If
    
End Function

Public Sub ReDrawClock()
    
    FrmMain.PicDisplay.Picture = LoadPicture("")
    FrmMain.PicDisplay.Cls
    FrmMain.PicDisplay.Picture = FrmMain.PicCanvas.Picture
    SkinForm FrmMain, FrmMain.PicDisplay
    
    FrmMain.PicDisplay.Refresh
    DrawTime
    MaskPrepared = False
    If FrmMain.Timer.Enabled = False Then FrmMain.Timer.Enabled = True
    
End Sub
Public Function sOperatingSystemString() As String

    Dim sOSString As String
    Dim lMaj As Long
    Dim lMin As Long
    Dim lPID As Long
    
    Dim osvVersionInfo As OSVERSIONINFO
    Dim osvexVersionInfo As OSVERSIONINFOEX
    osvVersionInfo.dwOSVersionInfoSize = Len(osvVersionInfo)
    osvexVersionInfo.dwOSVersionInfoSize = Len(osvexVersionInfo)

    If GetVersionEx(osvVersionInfo) <> 0 Then
        lMaj = osvVersionInfo.dwMajorVersion
        lMin = osvVersionInfo.dwMinorVersion
        lPID = osvVersionInfo.dwPlatformId

        Select Case lPID
            Case VER_PLATFORM_WIN32_WINDOWS '9x/ME
            OSBase = "9x/Me"
            If lMaj = 4 And lMin = 0 Then sOSString = "Windows 95,"
            If lMaj = 4 And lMin = 10 Then sOSString = "Windows 98,"
            If lMaj = 4 And lMin = 90 Then sOSString = "Windows ME,"


            If InStr(osvVersionInfo.szCSDVersion, "A") > 0 Then
                sOSString = sOSString & " Second Edition,"
            End If


            If InStr(osvVersionInfo.szCSDVersion, "C") > 0 Then
                sOSString = sOSString & " OSR2,"
            End If
            Case VER_PLATFORM_WIN32_NT 'NT Based
            OSBase = "NT"

            If GetVersionExA(osvexVersionInfo) <> 0 Then
                lMaj = osvexVersionInfo.dwMajorVersion
                lMin = osvexVersionInfo.dwMinorVersion
                lPID = osvexVersionInfo.dwPlatformId
                If lMaj = 4 And lMin = 0 Then sOSString = "Windows NT 4.0,"
                If lMaj = 5 And lMin = 0 Then sOSString = "Windows 2000,"
                If lMaj = 5 And lMin = 1 Then sOSString = "Windows XP,"
                If lMaj = 5 And lMin = 2 Then sOSString = "Windows Server 2003 Family,"
                


                If osvexVersionInfo.wSuiteMask = &H300 Then
                    sOSString = sOSString & " Home Edition,"
                Else
                    If lMin = 1 Then
                        sOSString = sOSString & " Professional,"
                    End If
                End If
                
                sOSString = sOSString & " " & osvexVersionInfo.szCSDVersion
                
                sOSString = Left(sOSString, InStr(sOSString, Chr(0)) - 1)
                
                sOSString = sOSString & ", Version "
                
                sOSString = sOSString & osvexVersionInfo.dwMajorVersion
                sOSString = sOSString & "."
                sOSString = sOSString & osvexVersionInfo.dwMinorVersion
                sOSString = sOSString & "."
                sOSString = sOSString & osvexVersionInfo.dwBuildNumber & " (" & OSBase & ")"
                
            End If
        End Select
    sOperatingSystemString = sOSString
    Exit Function
Else
    sOperatingSystemString = ""
    Exit Function
End If

End Function



Public Sub GetRegSettings()
    
    
    
    If StripNonChar(GetREGSZVal(REGKEY, "Install Date")) = "" Then
        WriteString REGKEY, "Install Date", FormatDateTime(Date$, vbLongDate)
        WriteString REGKEY, "Display Alarm", "True"
        WriteString REGKEY, "Clock X", 0
        WriteString REGKEY, "Clock Y", 0
        WriteString REGKEY, "Skin File", "Default"
        
        If FSys.FileExists(App.path & "\Alarm.Mp3") = True Then
            WriteString REGKEY, "Alarm File", App.path & "\Alarm.Mp3"
        Else
            WriteString REGKEY, "Alarm File", "None"
        End If
        
        WriteString REGKEY, "Alarm Set", "False"
        WriteString REGKEY, "Alarm Time", "00:00:00"
        WriteString REGKEY, "Alarm Type", 2 '0=Visual 1=Audio 2=Both
        WriteString REGKEY, "Alarm Note", "Your Alarm Time Has Been Reached."
        WriteString REGKEY, "Rem Clock Location", "True"
        WriteString REGKEY, "Start With Windows", "True"
        WriteString REGKEY, "Always Ontop", "True"
        WriteString REGKEY, "Alarm Volume", 1000
    End If
    
    Attribs.Clock.DispAlarm = GetREGSZVal(REGKEY, "Display Alarm")
    Attribs.Clock.RemLocation = GetREGSZVal(REGKEY, "Rem Clock Location")
    Attribs.Clock.X = Val(GetREGSZVal(REGKEY, "Clock X"))
    Attribs.Clock.Y = Val(GetREGSZVal(REGKEY, "Clock Y"))
    Attribs.Clock.StartWithWindows = GetREGSZVal(REGKEY, "Start With Windows")
    Attribs.Clock.AlwaysOntop = GetREGSZVal(REGKEY, "Always Ontop")
    
    Attribs.Skin.Skin = GetREGSZVal(REGKEY, "Skin File")
    
    If FSys.FileExists(Attribs.Skin.Skin) = False Then
        Attribs.Skin.Skin = "Default"
    End If
    
    Attribs.Alarm.AlarmNote = GetREGSZVal(REGKEY, "Alarm Note")
    Attribs.Alarm.AlarmSet = GetREGSZVal(REGKEY, "Alarm Set")
    Attribs.Alarm.AlarmSound = GetREGSZVal(REGKEY, "Alarm File")
    Attribs.Alarm.AlarmTime = GetREGSZVal(REGKEY, "Alarm Time")
    Attribs.Alarm.AlarmType = Val(GetREGSZVal(REGKEY, "Alarm Type"))
    
    If FSys.FileExists(Attribs.Alarm.AlarmSound) = False Then
        Attribs.Alarm.AlarmSound = "None"
        SetRegSettings
    End If
    
    Dim path As Long
    
    If Attribs.Clock.StartWithWindows = True Then
        If RegOpenKeyEx(HKEY_CURRENT_USER, SUREGKEY, 0, KEY_WRITE, path) Then Exit Sub
        RegSetValueEx path, App.Title & ".Exe", 0, REG_SZ, ByVal App.path & "\" & App.Title & ".Exe", Len(App.path & "\" & App.Title & ".Exe")
    Else
        If RegOpenKeyEx(HKEY_CURRENT_USER, SUREGKEY, 0, KEY_WRITE, path) Then Exit Sub
        RegDeleteValue path, App.Title & ".Exe"
    End If
    
    Attribs.Alarm.AlarmVolume = Val(GetREGSZVal(REGKEY, "Alarm Volume"))
    
End Sub

Public Sub SetRegSettings()
    
        If Trim(Attribs.Skin.Skin) = "" Then Attribs.Skin.Skin = "Default"
        If Trim(Attribs.Alarm.AlarmSound) = "" Then Attribs.Alarm.AlarmSound = "None"
        
        WriteString REGKEY, "Display Alarm", Attribs.Clock.DispAlarm
        
        If Attribs.Clock.RemLocation = True Then
            WriteString REGKEY, "Clock X", FrmMain.Left
            WriteString REGKEY, "Clock Y", FrmMain.Top
        End If
        
        WriteString REGKEY, "Skin File", Attribs.Skin.Skin
        WriteString REGKEY, "Alarm File", Attribs.Alarm.AlarmSound
        WriteString REGKEY, "Alarm Set", Attribs.Alarm.AlarmSet
        WriteString REGKEY, "Alarm Time", Attribs.Alarm.AlarmTime
        WriteString REGKEY, "Alarm Type", Attribs.Alarm.AlarmType
        WriteString REGKEY, "Alarm Note", Attribs.Alarm.AlarmNote
        WriteString REGKEY, "Rem Clock Location", Attribs.Clock.RemLocation
        WriteString REGKEY, "Start With Windows", Attribs.Clock.StartWithWindows
        WriteString REGKEY, "Always Ontop", Attribs.Clock.AlwaysOntop
        WriteString REGKEY, "Alarm Volume", Attribs.Alarm.AlarmVolume
        
End Sub


Public Sub ApplySettings()

    FrmMain.Left = Attribs.Clock.X
    FrmMain.Top = Attribs.Clock.Y
    
End Sub

Public Function StripNonChar(Text As String, Optional Replace As Boolean = False, Optional ReplaceChar As String = " ") As String

    Dim LP As Single
    Dim TempStr As String
    
    For LP = 1 To Len(Text)
        If Asc(Mid$(Text, LP, 1)) > 31 And Asc(Mid$(Text, LP, 1)) < 127 Then
            TempStr = TempStr & Mid$(Text, LP, 1)
        Else
            If Replace = True Then
                If ReplaceChar = "" Then ReplaceChar = " "
                TempStr = TempStr & ReplaceChar
            End If
        End If
    Next LP
    
    StripNonChar = Trim(TempStr)

End Function

Public Sub SkinPicture(Pic As PictureBox, MaskPic As PictureBox, Optional TransColor As Long)
    
    Dim Retr As Long
    Dim RgnFinal As Long
    Dim RgnTmp As Long
    
    Dim hHeight As Long
    Dim wWidth As Long
    
    Dim Col As Long
    Dim Start As Long
    Dim RowR As Long
    
    MaskPic.AutoSize = True
    MaskPic.AutoRedraw = True

    With Pic
        .Height = MaskPic.Height ' - 30
        .Width = MaskPic.Width ' - 30
    End With
    
    If TransColor < 1 Then
        TransColor = GetPixel(MaskPic.hdc, 0, 0)
    End If
    
    hHeight = MaskPic.Height / Screen.TwipsPerPixelY
    wWidth = MaskPic.Width / Screen.TwipsPerPixelX
    RgnFinal = CreateRectRgn(0, 0, 0, 0)

    For RowR = 0 To hHeight - 1
        
        Col = 0

        Do While Col < wWidth

            Do While Col < wWidth And GetPixel(MaskPic.hdc, Col, RowR) = TransColor
                Col = Col + 1
            Loop

            If Col < wWidth Then
                Start = Col
                Do While Col < wWidth And GetPixel(MaskPic.hdc, Col, RowR) <> TransColor
                    Col = Col + 1
                Loop
                
                If Col > wWidth Then Col = wWidth
                
                RgnTmp = CreateRectRgn(Start, RowR, Col, RowR + 1)
                Retr = CombineRgn(RgnFinal, RgnFinal, RgnTmp, RGN_OR)
                DeleteObject (RgnTmp)
                
            End If
        Loop
        
    Next RowR
    
    Retr = SetWindowRgn(Pic.hWnd, RgnFinal, True)
    
End Sub

Public Function GetFileName(FileName As String, Optional FilePath As String) As String
On Error GoTo FileNameError

    GetFileName = Right$(FileName, Len(FileName) - InStrRev(FileName, "\"))
    GetFileName = Left$(GetFileName, InStrRev(GetFileName, ".") - 1)
    FilePath = Left$(FileName, InStrRev(FileName, "\"))
    Exit Function

FileNameError:
GetFileName = ""
End Function

Public Sub DoButton(Target As PictureBox, Source As PictureClip, State As Boolean)
    
    Dim sSplit() As String
    sSplit = Split(Target.Tag, "\")
    
    If State = True Then
        Target.Picture = Source.GraphicCell(Val(sSplit(1)))
    Else
        Target.Picture = Source.GraphicCell(Val(sSplit(0)))
    End If
    
End Sub

Public Sub ApplyEnabled(Button As PictureBox, Value As Boolean)
    
    Dim sSplit() As String
    sSplit = Split(Button.Tag, "\")
    
    If Value = True Then
        Button.Enabled = True
        Button.Picture = FrmMain.StdClip.GraphicCell(Val(sSplit(0)))
    Else
        Button.Enabled = False
        Button.Picture = FrmMain.StdClip.GraphicCell(6)
    End If
    
End Sub

Public Sub AlarmSetup()
      
    If Attribs.Alarm.AlarmSet = False Then
        FrmSettings.FrameAlarm.Caption = "Alarm Settings: Alarm Is Not Active."
    Else
        FrmSettings.FrameAlarm.Caption = "Alarm Settings: Alarm Is Set For " & Left$(Attribs.Alarm.AlarmTime, 5) & ". Remain(" & TimeDiff(Time$, Attribs.Alarm.AlarmTime) & ")"
    End If
    
    Remain = TimeDiff(Time$, Attribs.Alarm.AlarmTime)
    
End Sub


Public Function TimeDiff(Time1 As String, Time2 As String) As String
    
    If TimeCheck = True Then Exit Function
    TimeCheck = True
    
    
    Dim H1 As Single, H2 As Single, M1 As Single, M2 As Single
    Dim HD As Single, MD As Single
    
    H1 = Hour(Time1): H2 = Hour(Time2)
    M1 = Minute(Time1): M2 = Minute(Time2)
    
    If M2 < 0 Then M2 = 59
    
    If M2 < M1 Then
        MD = (60 - M1) + M2
        H2 = H2 - 1
    Else
        MD = M2 - M1
    End If
    
    If H2 < H1 Then
        HD = (24 - H1) + H2
    Else
        HD = H2 - H1
    End If
        
    TimeDiff = Format(HD, "00") & ":" & Format(MD, "00")
    
    TimeCheck = False
    
End Function

Public Function MCISend(SendString As String, FromForm As Form) As String

    Dim RetStr As String * 255

    Call mciSendString(SendString, RetStr, 255, FromForm.hWnd)
    MCISend = Replace(RetStr, Chr(0), "")

End Function

Public Function Snooze(sTime As String, sDuration As Single) As String
    
    Dim H As Single, M As Single
    Dim LP As Single
    
    H = Hour(sTime)
    M = Minute(sTime)
    
    For LP = 1 To 5
        M = M + 1
        If M > 59 Then
            M = 0
            H = H + 1
            If H > 23 Then H = 0
        End If
    Next LP
    
    Snooze = Format(H, "00") & ":" & Format(M, "00") & ":" & Format(Second(sTime), "00")
    
End Function
