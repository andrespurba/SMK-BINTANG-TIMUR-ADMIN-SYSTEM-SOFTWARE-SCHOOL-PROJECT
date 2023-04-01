VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form login
  Caption = "Form1"
  BackColor = &H404040&
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  BorderStyle = 0 'None
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 0
  ClientTop = 0
  ClientWidth = 14955
  ClientHeight = 9030
  StartUpPosition = 2 'CenterScreen
  Begin VB.Timer Timer3
    Interval = 10
    Left = 3630
    Top = 8220
  End
  Begin VB.Timer Timer2
    Interval = 10
    Left = 3210
    Top = 8220
  End
  Begin VB.Timer Timer1
    Interval = 10
    Left = 2790
    Top = 8220
  End
  Begin Project1.jcbutton login
    Left = 11220
    Top = 7080
    Width = 1425
    Height = 465
    TabIndex = 8
    OleObjectBlob = "login.frx":0000
  End
  Begin MSAdodcLib.Adodc adodata
    Left = 1590
    Top = 8220
    Width = 1200
    Height = 435
    Visible = 0   'False
    OleObjectBlob = "login.frx":01DC
  End
  Begin MSDataGridLib.DataGrid dgdata
    Left = 0
    Top = 8220
    Width = 1575
    Height = 465
    Visible = 0   'False
    TabIndex = 6
    OleObjectBlob = "login.frx":0514
  End
  Begin VB.Frame fpw
    Caption = "Frame1"
    BackColor = &HFFFFFF&
    Left = 9930
    Top = 4680
    Width = 4305
    Height = 1755
    TabIndex = 3
    BorderStyle = 0 'None
    Begin VB.TextBox txtpw
      BackColor = &HFFFFFF&
      Left = 450
      Top = 1170
      Width = 2775
      Height = 405
      TabIndex = 4
      BorderStyle = 0 'None
      PasswordChar = "X"
    End
    Begin VB.Label Label5
      Caption = "P A S S W O R D"
      ForeColor = &H808080&
      Left = 1020
      Top = 30
      Width = 1995
      Height = 315
      TabIndex = 10
      BackStyle = 0 'Transparent
      BeginProperty Font
        Name = "Trebuchet MS"
        Size = 12
        Charset = 0
        Weight = 700
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Image o
      Picture = "login.frx":06C0
      Left = 3270
      Top = 1170
      Width = 465
      Height = 255
      Stretch = -1  'True
      Appearance = 0 'Flat
    End
    Begin VB.Image x
      Picture = "login.frx":1185
      Left = 3270
      Top = 1110
      Width = 465
      Height = 375
      Stretch = -1  'True
      Appearance = 0 'Flat
    End
    Begin VB.Shape Shape19
      BorderColor = &H808080&
      Left = 30
      Top = 990
      Width = 3885
      Height = 675
      Shape = 4
    End
  End
  Begin VB.Frame fus
    Caption = "Frame1"
    BackColor = &HFFFFFF&
    Left = 9930
    Top = 4710
    Width = 4305
    Height = 1875
    TabIndex = 1
    BorderStyle = 0 'None
    Begin VB.TextBox txtnama
      BackColor = &HFFFFFF&
      Left = 480
      Top = 1110
      Width = 2985
      Height = 405
      TabIndex = 2
      BorderStyle = 0 'None
      BeginProperty Font
        Name = "Trebuchet MS"
        Size = 12
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label Label3
      Caption = "U S E R N A M E"
      ForeColor = &H808080&
      Left = 1050
      Top = 0
      Width = 1995
      Height = 585
      TabIndex = 9
      BackStyle = 0 'Transparent
      BeginProperty Font
        Name = "Trebuchet MS"
        Size = 12
        Charset = 0
        Weight = 700
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Shape Shape20
      BorderColor = &H808080&
      Left = 30
      Top = 960
      Width = 3885
      Height = 675
      Shape = 4
    End
  End
  Begin VB.Image back
    Picture = "login.frx":1DAD
    Left = 0
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Shape shpp
    Left = 0
    Top = 450
    Width = 15
    Height = 75
    Shape = 4
    BorderStyle = 0 'Transparent
    FillColor = &HFF0000&
    FillStyle = 0
  End
  Begin VB.Label Label4
    Caption = "Login System"
    ForeColor = &H80FF80&
    Left = 7080
    Top = 0
    Width = 3135
    Height = 585
    TabIndex = 7
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 12
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Label Label1
    Caption = "PENDAFTARAN SISWA/i SMK BARU BINTANG TIMUR"
    ForeColor = &H8000&
    Left = 1290
    Top = 720
    Width = 3135
    Height = 585
    TabIndex = 5
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 12
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Shape Shape18
    Left = 7350
    Top = 2460
    Width = 165
    Height = 165
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &HFFFF&
    FillStyle = 0
  End
  Begin VB.Shape Shape17
    Left = 7770
    Top = 2460
    Width = 165
    Height = 165
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &H80FF&
    FillStyle = 0
  End
  Begin VB.Shape Shape16
    Left = 8130
    Top = 2460
    Width = 165
    Height = 165
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &HFF&
    FillStyle = 0
  End
  Begin VB.Line Line2
    BorderColor = &HC0C0C0&
    X1 = 9540
    Y1 = 6960
    X2 = 14250
    Y2 = 6960
  End
  Begin VB.Line Line1
    BorderColor = &HC0C0C0&
    X1 = 9540
    Y1 = 2940
    X2 = 14250
    Y2 = 2940
  End
  Begin VB.Image Image6
End

Attribute VB_Name = "login"

'VA: 6DE6D8
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
'VA: 6DE690
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'VA: 6DE648
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'VA: 6DE604
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'VA: 6DE5C0
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
'VA: 6DE574
Private Declare Sub _TrackMouseEvent Lib "comctl32"()
'VA: 6DE518
Private Declare Sub TrackMouseEvent Lib "user32"()
'VA: 6DE4D0
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
'VA: 6DE488
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'VA: 6DE444
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'VA: 6DE3F8
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'VA: 6DE388
Private Declare Function SetLayout Lib "gdi32" (ByVal hdc As Long, ByVal dword As Long) As Long
'VA: 6DE344
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
'VA: 6DE300
Private Declare Sub DrawTextW Lib "user32"()
'VA: 6DE2BC
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'VA: 6DE278
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'VA: 6DE234
Private Declare Function GetCapture Lib "user32" () As Long
'VA: 6DE1D8
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
'VA: 6DE194
Private Declare Function ReleaseCapture Lib "user32" () As Long
'VA: 6DE14C
Private Declare Sub IsAppThemed Lib "uxtheme"()
'VA: 6DE108
Private Declare Sub GetThemeBackgroundRegion Lib "uxtheme"()
'VA: 6DE0B4
Private Declare Sub DrawThemeBackground Lib "uxtheme"()
'VA: 6DE068
Private Declare Sub CloseThemeData Lib "uxtheme"()
'VA: 6DE018
Private Declare Sub OpenThemeData Lib "uxtheme"()
'VA: 6DDFC0
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'VA: 6DDF78
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'VA: 6DDF30
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'VA: 6DDEE8
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'VA: 6DDEA0
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'VA: 6DDE54
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'VA: 6DDE0C
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'VA: 6DDDB4
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'VA: 6DDD70
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'VA: 6DDD2C
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
'VA: 6DDCD8
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'VA: 6DDC7C
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'VA: 6DDC34
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, ByVal pt As POINTAPI) As Long
'VA: 6DDBEC
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long
'VA: 6DDBA8
Private Declare Sub TransparentBlt Lib "msimg32"()
'VA: 6DDB50
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'VA: 6DDB10
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'VA: 6DDAA8
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'VA: 6DDA60
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'VA: 6DDA18
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
'VA: 6DD9D0
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'VA: 6DD98C
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
'VA: 6DD948
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
'VA: 6DD8F8
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'VA: 6DD8B4
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
'VA: 6DD868
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'VA: 6DD820
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'VA: 6DD7E0
Private Declare Sub OleTranslateColor Lib "olepro32"()
'VA: 6DD780
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'VA: 6DD71C
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
'VA: 6DD6D4
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'VA: 6DD68C
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'VA: 6DD648
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'VA: 6DD604
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'VA: 6DD5B8
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'VA: 6DD568
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'VA: 6DD520
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'VA: 6DD4DC
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'VA: 6DD498
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'VA: 6DD440
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'VA: 6DD3F0
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'VA: 6DD388
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
'VA: 6DD340
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
'VA: 6DD2F4
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'VA: 6DD2B0
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
'VA: 6DD26C
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long


Private Sub login_UnknownEvent_9 'DE9530
  loc_00DE9530: push ebp
  loc_00DE9531: mov ebp, esp
  loc_00DE9533: sub esp, 0000000Ch
  loc_00DE9536: push 00402836h ; __vbaExceptHandler
  loc_00DE953B: mov eax, fs:[00000000h]
  loc_00DE9541: push eax
  loc_00DE9542: mov fs:[00000000h], esp
  loc_00DE9549: sub esp, 000000B8h
  loc_00DE954F: push ebx
  loc_00DE9550: push esi
  loc_00DE9551: push edi
  loc_00DE9552: mov var_C, esp
  loc_00DE9555: mov var_8, 004012B0h
  loc_00DE955C: mov esi, Me
  loc_00DE955F: mov eax, esi
  loc_00DE9561: and eax, 00000001h
  loc_00DE9564: mov var_4, eax
  loc_00DE9567: and esi, FFFFFFFEh
  loc_00DE956A: push esi
  loc_00DE956B: mov Me, esi
  loc_00DE956E: mov ecx, [esi]
  loc_00DE9570: call [ecx+00000004h]
  loc_00DE9573: mov edx, [esi]
  loc_00DE9575: xor ebx, ebx
  loc_00DE9577: push esi
  loc_00DE9578: mov var_24, ebx
  loc_00DE957B: mov var_28, ebx
  loc_00DE957E: mov var_2C, ebx
  loc_00DE9581: mov var_30, ebx
  loc_00DE9584: mov var_34, ebx
  loc_00DE9587: mov var_44, ebx
  loc_00DE958A: mov var_54, ebx
  loc_00DE958D: mov var_64, ebx
  loc_00DE9590: mov var_74, ebx
  loc_00DE9593: mov var_84, ebx
  loc_00DE9599: mov var_94, ebx
  loc_00DE959F: mov var_B8, ebx
  loc_00DE95A5: call [edx+00000324h]
  loc_00DE95AB: mov edi, [004010ACh] ; __vbaObjSet
  loc_00DE95B1: push eax
  loc_00DE95B2: lea eax, var_30
  loc_00DE95B5: push eax
  loc_00DE95B6: call edi
  loc_00DE95B8: mov ecx, [eax]
  loc_00DE95BA: lea edx, var_28
  loc_00DE95BD: push edx
  loc_00DE95BE: push eax
  loc_00DE95BF: mov var_BC, eax
  loc_00DE95C5: call [ecx+000000A0h]
  loc_00DE95CB: cmp eax, ebx
  loc_00DE95CD: fnclex
  loc_00DE95CF: jge 00DE95E9h
  loc_00DE95D1: mov ecx, var_BC
  loc_00DE95D7: push 000000A0h
  loc_00DE95DC: push 006DCB70h
  loc_00DE95E1: push ecx
  loc_00DE95E2: push eax
  loc_00DE95E3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE95E9: mov edx, var_28
  loc_00DE95EC: mov ebx, [0040106Ch] ; __vbaStrCat
  loc_00DE95F2: push 006DCB24h ; "select * from login where username='"
  loc_00DE95F7: push edx
  loc_00DE95F8: call ebx
  loc_00DE95FA: mov edx, eax
  loc_00DE95FC: lea ecx, var_2C
  loc_00DE95FF: call [00401228h] ; __vbaStrMove
  loc_00DE9605: push eax
  loc_00DE9606: push 006DCB84h ; "'"
  loc_00DE960B: call ebx
  loc_00DE960D: lea edx, var_44
  loc_00DE9610: lea ecx, var_24
  loc_00DE9613: mov var_3C, eax
  loc_00DE9616: mov var_44, 00000008h
  loc_00DE961D: call [0040101Ch] ; __vbaVarMove
  loc_00DE9623: lea eax, var_2C
  loc_00DE9626: lea ecx, var_28
  loc_00DE9629: push eax
  loc_00DE962A: push ecx
  loc_00DE962B: push 00000002h
  loc_00DE962D: call [004011BCh] ; __vbaFreeStrList
  loc_00DE9633: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00DE9639: add esp, 0000000Ch
  loc_00DE963C: lea ecx, var_30
  loc_00DE963F: call ebx
  loc_00DE9641: mov edx, [esi]
  loc_00DE9643: push esi
  loc_00DE9644: call [edx+0000030Ch]
  loc_00DE964A: push eax
  loc_00DE964B: lea eax, var_30
  loc_00DE964E: push eax
  loc_00DE964F: call edi
  loc_00DE9651: mov ecx, [eax]
  loc_00DE9653: lea edx, var_28
  loc_00DE9656: push edx
  loc_00DE9657: push eax
  loc_00DE9658: mov var_BC, eax
  loc_00DE965E: call [ecx+000000A0h]
  loc_00DE9664: test eax, eax
  loc_00DE9666: fnclex
  loc_00DE9668: jge 00DE9682h
  loc_00DE966A: mov ecx, var_BC
  loc_00DE9670: push 000000A0h
  loc_00DE9675: push 006DCB70h
  loc_00DE967A: push ecx
  loc_00DE967B: push eax
  loc_00DE967C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9682: mov edx, var_28
  loc_00DE9685: push 006DCB8Ch ; "select * from login where password='"
  loc_00DE968A: push edx
  loc_00DE968B: call [0040106Ch] ; __vbaStrCat
  loc_00DE9691: mov edx, eax
  loc_00DE9693: lea ecx, var_2C
  loc_00DE9696: call [00401228h] ; __vbaStrMove
  loc_00DE969C: push eax
  loc_00DE969D: push 006DCB84h ; "'"
  loc_00DE96A2: call [0040106Ch] ; __vbaStrCat
  loc_00DE96A8: lea edx, var_44
  loc_00DE96AB: lea ecx, var_24
  loc_00DE96AE: mov var_3C, eax
  loc_00DE96B1: mov var_44, 00000008h
  loc_00DE96B8: call [0040101Ch] ; __vbaVarMove
  loc_00DE96BE: lea eax, var_2C
  loc_00DE96C1: lea ecx, var_28
  loc_00DE96C4: push eax
  loc_00DE96C5: push ecx
  loc_00DE96C6: push 00000002h
  loc_00DE96C8: call [004011BCh] ; __vbaFreeStrList
  loc_00DE96CE: add esp, 0000000Ch
  loc_00DE96D1: lea ecx, var_30
  loc_00DE96D4: call ebx
  loc_00DE96D6: lea edx, var_24
  loc_00DE96D9: push edx
  loc_00DE96DA: call [00401230h] ; __vbaStrVarCopy
  loc_00DE96E0: sub esp, 00000010h
  loc_00DE96E3: mov ecx, 00000008h
  loc_00DE96E8: mov edx, esp
  loc_00DE96EA: mov var_44, ecx
  loc_00DE96ED: mov var_3C, eax
  loc_00DE96F0: push 0000000Eh
  loc_00DE96F2: mov [edx], ecx
  loc_00DE96F4: mov ecx, var_40
  loc_00DE96F7: push esi
  loc_00DE96F8: mov [edx+00000004h], ecx
  loc_00DE96FB: mov ecx, [esi]
  loc_00DE96FD: mov [edx+00000008h], eax
  loc_00DE9700: mov eax, var_38
  loc_00DE9703: mov [edx+0000000Ch], eax
  loc_00DE9706: call [ecx+000003A8h]
  loc_00DE970C: lea edx, var_30
  loc_00DE970F: push eax
  loc_00DE9710: push edx
  loc_00DE9711: call edi
  loc_00DE9713: push eax
  loc_00DE9714: call [00401238h] ; __vbaLateIdSt
  loc_00DE971A: lea ecx, var_30
  loc_00DE971D: call ebx
  loc_00DE971F: lea ecx, var_44
  loc_00DE9722: call [00401024h] ; __vbaFreeVar
  loc_00DE9728: mov eax, [esi]
  loc_00DE972A: push 00000000h
  loc_00DE972C: push 00000019h
  loc_00DE972E: push esi
  loc_00DE972F: call [eax+000003A8h]
  loc_00DE9735: lea ecx, var_30
  loc_00DE9738: push eax
  loc_00DE9739: push ecx
  loc_00DE973A: call edi
  loc_00DE973C: push eax
  loc_00DE973D: call [00401030h] ; __vbaLateIdCall
  loc_00DE9743: add esp, 0000000Ch
  loc_00DE9746: lea ecx, var_30
  loc_00DE9749: call ebx
  loc_00DE974B: mov edx, [esi]
  loc_00DE974D: push 006DCBD8h
  loc_00DE9752: push 00000000h
  loc_00DE9754: push 00000018h
  loc_00DE9756: push esi
  loc_00DE9757: call [edx+000003A8h]
  loc_00DE975D: push eax
  loc_00DE975E: lea eax, var_30
  loc_00DE9761: push eax
  loc_00DE9762: call edi
  loc_00DE9764: lea ecx, var_44
  loc_00DE9767: push eax
  loc_00DE9768: push ecx
  loc_00DE9769: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00DE976F: add esp, 00000010h
  loc_00DE9772: push eax
  loc_00DE9773: call [00401120h] ; __vbaCastObjVar
  loc_00DE9779: lea edx, var_34
  loc_00DE977C: push eax
  loc_00DE977D: push edx
  loc_00DE977E: call edi
  loc_00DE9780: mov ecx, [eax]
  loc_00DE9782: lea edx, var_B8
  loc_00DE9788: push edx
  loc_00DE9789: push eax
  loc_00DE978A: mov var_BC, eax
  loc_00DE9790: call [ecx+00000050h]
  loc_00DE9793: test eax, eax
  loc_00DE9795: fnclex
  loc_00DE9797: jge 00DE97AEh
  loc_00DE9799: mov ecx, var_BC
  loc_00DE979F: push 00000050h
  loc_00DE97A1: push 006DCBE8h
  loc_00DE97A6: push ecx
  loc_00DE97A7: push eax
  loc_00DE97A8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE97AE: mov edx, var_B8
  loc_00DE97B4: lea eax, var_34
  loc_00DE97B7: lea ecx, var_30
  loc_00DE97BA: push eax
  loc_00DE97BB: push ecx
  loc_00DE97BC: push 00000002h
  loc_00DE97BE: mov var_C4, edx
  loc_00DE97C4: call [00401048h] ; __vbaFreeObjList
  loc_00DE97CA: add esp, 0000000Ch
  loc_00DE97CD: lea ecx, var_44
  loc_00DE97D0: call [00401024h] ; __vbaFreeVar
  loc_00DE97D6: cmp var_C4, 0000h
  loc_00DE97DE: mov edx, [esi]
  loc_00DE97E0: push esi
  loc_00DE97E1: jz 00DE9A02h
  loc_00DE97E7: call [edx+00000300h]
  loc_00DE97ED: push eax
  loc_00DE97EE: lea eax, var_30
  loc_00DE97F1: push eax
  loc_00DE97F2: call edi
  loc_00DE97F4: mov ecx, [eax]
  loc_00DE97F6: push FFFFFFFFh
  loc_00DE97F8: push eax
  loc_00DE97F9: mov var_BC, eax
  loc_00DE97FF: call [ecx+0000005Ch]
  loc_00DE9802: test eax, eax
  loc_00DE9804: fnclex
  loc_00DE9806: jge 00DE981Dh
  loc_00DE9808: mov edx, var_BC
  loc_00DE980E: push 0000005Ch
  loc_00DE9810: push 006DCAE0h
  loc_00DE9815: push edx
  loc_00DE9816: push eax
  loc_00DE9817: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE981D: lea ecx, var_30
  loc_00DE9820: call ebx
  loc_00DE9822: mov ecx, 80020004h
  loc_00DE9827: mov eax, 0000000Ah
  loc_00DE982C: mov var_6C, ecx
  loc_00DE982F: mov var_5C, ecx
  loc_00DE9832: lea edx, var_94
  loc_00DE9838: lea ecx, var_54
  loc_00DE983B: mov var_74, eax
  loc_00DE983E: mov var_64, eax
  loc_00DE9841: mov var_8C, 006DCC68h ; "ERROR 01"
  loc_00DE984B: mov var_94, 00000008h
  loc_00DE9855: call [004011F0h] ; __vbaVarDup
  loc_00DE985B: lea edx, var_84
  loc_00DE9861: lea ecx, var_44
  loc_00DE9864: mov var_7C, 006DCBFCh ; "Username atau Password Yang Anda Masukkan Salah ! "
  loc_00DE986B: mov var_84, 00000008h
  loc_00DE9875: call [004011F0h] ; __vbaVarDup
  loc_00DE987B: lea eax, var_74
  loc_00DE987E: lea ecx, var_64
  loc_00DE9881: push eax
  loc_00DE9882: lea edx, var_54
  loc_00DE9885: push ecx
  loc_00DE9886: push edx
  loc_00DE9887: lea eax, var_44
  loc_00DE988A: push 00000010h
  loc_00DE988C: push eax
  loc_00DE988D: call [004010A8h] ; rtcMsgBox
  loc_00DE9893: lea ecx, var_74
  loc_00DE9896: lea edx, var_64
  loc_00DE9899: push ecx
  loc_00DE989A: lea eax, var_54
  loc_00DE989D: push edx
  loc_00DE989E: lea ecx, var_44
  loc_00DE98A1: push eax
  loc_00DE98A2: push ecx
  loc_00DE98A3: push 00000004h
  loc_00DE98A5: call [00401038h] ; __vbaFreeVarList
  loc_00DE98AB: mov edx, [esi]
  loc_00DE98AD: add esp, 00000014h
  loc_00DE98B0: push esi
  loc_00DE98B1: call [edx+00000308h]
  loc_00DE98B7: push eax
  loc_00DE98B8: lea eax, var_30
  loc_00DE98BB: push eax
  loc_00DE98BC: call edi
  loc_00DE98BE: mov ecx, [eax]
  loc_00DE98C0: push 00000000h
  loc_00DE98C2: push eax
  loc_00DE98C3: mov var_BC, eax
  loc_00DE98C9: call [ecx+0000009Ch]
  loc_00DE98CF: test eax, eax
  loc_00DE98D1: fnclex
  loc_00DE98D3: jge 00DE98EDh
  loc_00DE98D5: mov edx, var_BC
  loc_00DE98DB: push 0000009Ch
  loc_00DE98E0: push 006DCAD0h
  loc_00DE98E5: push edx
  loc_00DE98E6: push eax
  loc_00DE98E7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE98ED: lea ecx, var_30
  loc_00DE98F0: call ebx
  loc_00DE98F2: mov eax, [esi]
  loc_00DE98F4: push esi
  loc_00DE98F5: call [eax+00000320h]
  loc_00DE98FB: lea ecx, var_30
  loc_00DE98FE: push eax
  loc_00DE98FF: push ecx
  loc_00DE9900: call edi
  loc_00DE9902: mov edx, [eax]
  loc_00DE9904: push FFFFFFFFh
  loc_00DE9906: push eax
  loc_00DE9907: mov var_BC, eax
  loc_00DE990D: call [edx+0000009Ch]
  loc_00DE9913: test eax, eax
  loc_00DE9915: fnclex
  loc_00DE9917: jge 00DE9931h
  loc_00DE9919: mov ecx, var_BC
  loc_00DE991F: push 0000009Ch
  loc_00DE9924: push 006DCAD0h
  loc_00DE9929: push ecx
  loc_00DE992A: push eax
  loc_00DE992B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9931: lea ecx, var_30
  loc_00DE9934: call ebx
  loc_00DE9936: mov edx, [esi]
  loc_00DE9938: push esi
  loc_00DE9939: call [edx+0000030Ch]
  loc_00DE993F: push eax
  loc_00DE9940: lea eax, var_30
  loc_00DE9943: push eax
  loc_00DE9944: call edi
  loc_00DE9946: mov ecx, [eax]
  loc_00DE9948: push 006DCC80h
  loc_00DE994D: push eax
  loc_00DE994E: mov var_BC, eax
  loc_00DE9954: call [ecx+000000A4h]
  loc_00DE995A: test eax, eax
  loc_00DE995C: fnclex
  loc_00DE995E: jge 00DE9978h
  loc_00DE9960: mov edx, var_BC
  loc_00DE9966: push 000000A4h
  loc_00DE996B: push 006DCB70h
  loc_00DE9970: push edx
  loc_00DE9971: push eax
  loc_00DE9972: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9978: lea ecx, var_30
  loc_00DE997B: call ebx
  loc_00DE997D: mov eax, [esi]
  loc_00DE997F: push esi
  loc_00DE9980: call [eax+00000324h]
  loc_00DE9986: lea ecx, var_30
  loc_00DE9989: push eax
  loc_00DE998A: push ecx
  loc_00DE998B: call edi
  loc_00DE998D: mov esi, eax
  loc_00DE998F: push 006DCC80h
  loc_00DE9994: push esi
  loc_00DE9995: mov edx, [esi]
  loc_00DE9997: call [edx+000000A4h]
  loc_00DE999D: test eax, eax
  loc_00DE999F: fnclex
  loc_00DE99A1: jge 00DE99B5h
  loc_00DE99A3: push 000000A4h
  loc_00DE99A8: push 006DCB70h
  loc_00DE99AD: push esi
  loc_00DE99AE: push eax
  loc_00DE99AF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE99B5: lea ecx, var_30
  loc_00DE99B8: call ebx
  loc_00DE99BA: mov esi, [004011F0h] ; __vbaVarDup
  loc_00DE99C0: mov ecx, 80020004h
  loc_00DE99C5: mov var_6C, ecx
  loc_00DE99C8: mov eax, 0000000Ah
  loc_00DE99CD: mov var_5C, ecx
  loc_00DE99D0: mov edi, 00000008h
  loc_00DE99D5: lea edx, var_94
  loc_00DE99DB: lea ecx, var_54
  loc_00DE99DE: mov var_74, eax
  loc_00DE99E1: mov var_64, eax
  loc_00DE99E4: mov var_8C, 006DCC68h ; "ERROR 01"
  loc_00DE99EE: mov var_94, edi
  loc_00DE99F4: call __vbaVarDup
  loc_00DE99F6: mov var_7C, 006DCC88h ; "Daftar Terlebih Dahulu ! "
  loc_00DE99FD: jmp 00DE9F28h
  loc_00DE9A02: call [edx+0000030Ch]
  loc_00DE9A08: push eax
  loc_00DE9A09: lea eax, var_30
  loc_00DE9A0C: push eax
  loc_00DE9A0D: call edi
  loc_00DE9A0F: mov ecx, [eax]
  loc_00DE9A11: lea edx, var_28
  loc_00DE9A14: push edx
  loc_00DE9A15: push eax
  loc_00DE9A16: mov var_BC, eax
  loc_00DE9A1C: call [ecx+000000A0h]
  loc_00DE9A22: test eax, eax
  loc_00DE9A24: fnclex
  loc_00DE9A26: jge 00DE9A40h
  loc_00DE9A28: mov ecx, var_BC
  loc_00DE9A2E: push 000000A0h
  loc_00DE9A33: push 006DCB70h
  loc_00DE9A38: push ecx
  loc_00DE9A39: push eax
  loc_00DE9A3A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9A40: mov edx, var_28
  loc_00DE9A43: push edx
  loc_00DE9A44: push 006DCC80h
  loc_00DE9A49: call [00401104h] ; __vbaStrCmp
  loc_00DE9A4F: neg eax
  loc_00DE9A51: sbb eax, eax
  loc_00DE9A53: lea ecx, var_28
  loc_00DE9A56: inc eax
  loc_00DE9A57: neg eax
  loc_00DE9A59: mov var_C4, eax
  loc_00DE9A5F: call [00401258h] ; __vbaFreeStr
  loc_00DE9A65: lea ecx, var_30
  loc_00DE9A68: call ebx
  loc_00DE9A6A: cmp var_C4, 0000h
  loc_00DE9A72: jz 00DE9C96h
  loc_00DE9A78: mov eax, [esi]
  loc_00DE9A7A: push esi
  loc_00DE9A7B: call [eax+000002FCh]
  loc_00DE9A81: lea ecx, var_30
  loc_00DE9A84: push eax
  loc_00DE9A85: push ecx
  loc_00DE9A86: call edi
  loc_00DE9A88: mov edx, [eax]
  loc_00DE9A8A: push FFFFFFFFh
  loc_00DE9A8C: push eax
  loc_00DE9A8D: mov var_BC, eax
  loc_00DE9A93: call [edx+0000005Ch]
  loc_00DE9A96: test eax, eax
  loc_00DE9A98: fnclex
  loc_00DE9A9A: jge 00DE9AB1h
  loc_00DE9A9C: mov ecx, var_BC
  loc_00DE9AA2: push 0000005Ch
  loc_00DE9AA4: push 006DCAE0h
  loc_00DE9AA9: push ecx
  loc_00DE9AAA: push eax
  loc_00DE9AAB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9AB1: lea ecx, var_30
  loc_00DE9AB4: call ebx
  loc_00DE9AB6: mov ecx, 80020004h
  loc_00DE9ABB: mov eax, 0000000Ah
  loc_00DE9AC0: mov var_6C, ecx
  loc_00DE9AC3: mov var_5C, ecx
  loc_00DE9AC6: lea edx, var_94
  loc_00DE9ACC: lea ecx, var_54
  loc_00DE9ACF: mov var_74, eax
  loc_00DE9AD2: mov var_64, eax
  loc_00DE9AD5: mov var_8C, 006DCC68h ; "ERROR 01"
  loc_00DE9ADF: mov var_94, 00000008h
  loc_00DE9AE9: call [004011F0h] ; __vbaVarDup
  loc_00DE9AEF: lea edx, var_84
  loc_00DE9AF5: lea ecx, var_44
  loc_00DE9AF8: mov var_7C, 006DCBFCh ; "Username atau Password Yang Anda Masukkan Salah ! "
  loc_00DE9AFF: mov var_84, 00000008h
  loc_00DE9B09: call [004011F0h] ; __vbaVarDup
  loc_00DE9B0F: lea edx, var_74
  loc_00DE9B12: lea eax, var_64
  loc_00DE9B15: push edx
  loc_00DE9B16: lea ecx, var_54
  loc_00DE9B19: push eax
  loc_00DE9B1A: push ecx
  loc_00DE9B1B: lea edx, var_44
  loc_00DE9B1E: push 00000010h
  loc_00DE9B20: push edx
  loc_00DE9B21: call [004010A8h] ; rtcMsgBox
  loc_00DE9B27: lea eax, var_74
  loc_00DE9B2A: lea ecx, var_64
  loc_00DE9B2D: push eax
  loc_00DE9B2E: lea edx, var_54
  loc_00DE9B31: push ecx
  loc_00DE9B32: lea eax, var_44
  loc_00DE9B35: push edx
  loc_00DE9B36: push eax
  loc_00DE9B37: push 00000004h
  loc_00DE9B39: call [00401038h] ; __vbaFreeVarList
  loc_00DE9B3F: mov ecx, [esi]
  loc_00DE9B41: add esp, 00000014h
  loc_00DE9B44: push esi
  loc_00DE9B45: call [ecx+00000308h]
  loc_00DE9B4B: lea edx, var_30
  loc_00DE9B4E: push eax
  loc_00DE9B4F: push edx
  loc_00DE9B50: call edi
  loc_00DE9B52: mov ecx, [eax]
  loc_00DE9B54: push 00000000h
  loc_00DE9B56: push eax
  loc_00DE9B57: mov var_BC, eax
  loc_00DE9B5D: call [ecx+0000009Ch]
  loc_00DE9B63: test eax, eax
  loc_00DE9B65: fnclex
  loc_00DE9B67: jge 00DE9B81h
  loc_00DE9B69: mov edx, var_BC
  loc_00DE9B6F: push 0000009Ch
  loc_00DE9B74: push 006DCAD0h
  loc_00DE9B79: push edx
  loc_00DE9B7A: push eax
  loc_00DE9B7B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9B81: lea ecx, var_30
  loc_00DE9B84: call ebx
  loc_00DE9B86: mov eax, [esi]
  loc_00DE9B88: push esi
  loc_00DE9B89: call [eax+00000320h]
  loc_00DE9B8F: lea ecx, var_30
  loc_00DE9B92: push eax
  loc_00DE9B93: push ecx
  loc_00DE9B94: call edi
  loc_00DE9B96: mov edx, [eax]
  loc_00DE9B98: push FFFFFFFFh
  loc_00DE9B9A: push eax
  loc_00DE9B9B: mov var_BC, eax
  loc_00DE9BA1: call [edx+0000009Ch]
  loc_00DE9BA7: test eax, eax
  loc_00DE9BA9: fnclex
  loc_00DE9BAB: jge 00DE9BC5h
  loc_00DE9BAD: mov ecx, var_BC
  loc_00DE9BB3: push 0000009Ch
  loc_00DE9BB8: push 006DCAD0h
  loc_00DE9BBD: push ecx
  loc_00DE9BBE: push eax
  loc_00DE9BBF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9BC5: lea ecx, var_30
  loc_00DE9BC8: call ebx
  loc_00DE9BCA: mov edx, [esi]
  loc_00DE9BCC: push esi
  loc_00DE9BCD: call [edx+0000030Ch]
  loc_00DE9BD3: push eax
  loc_00DE9BD4: lea eax, var_30
  loc_00DE9BD7: push eax
  loc_00DE9BD8: call edi
  loc_00DE9BDA: mov ecx, [eax]
  loc_00DE9BDC: push 006DCC80h
  loc_00DE9BE1: push eax
  loc_00DE9BE2: mov var_BC, eax
  loc_00DE9BE8: call [ecx+000000A4h]
  loc_00DE9BEE: test eax, eax
  loc_00DE9BF0: fnclex
  loc_00DE9BF2: jge 00DE9C0Ch
  loc_00DE9BF4: mov edx, var_BC
  loc_00DE9BFA: push 000000A4h
  loc_00DE9BFF: push 006DCB70h
  loc_00DE9C04: push edx
  loc_00DE9C05: push eax
  loc_00DE9C06: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9C0C: lea ecx, var_30
  loc_00DE9C0F: call ebx
  loc_00DE9C11: mov eax, [esi]
  loc_00DE9C13: push esi
  loc_00DE9C14: call [eax+00000324h]
  loc_00DE9C1A: lea ecx, var_30
  loc_00DE9C1D: push eax
  loc_00DE9C1E: push ecx
  loc_00DE9C1F: call edi
  loc_00DE9C21: mov esi, eax
  loc_00DE9C23: push 006DCC80h
  loc_00DE9C28: push esi
  loc_00DE9C29: mov edx, [esi]
  loc_00DE9C2B: call [edx+000000A4h]
  loc_00DE9C31: test eax, eax
  loc_00DE9C33: fnclex
  loc_00DE9C35: jge 00DE9C49h
  loc_00DE9C37: push 000000A4h
  loc_00DE9C3C: push 006DCB70h
  loc_00DE9C41: push esi
  loc_00DE9C42: push eax
  loc_00DE9C43: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9C49: lea ecx, var_30
  loc_00DE9C4C: call ebx
  loc_00DE9C4E: mov esi, [004011F0h] ; __vbaVarDup
  loc_00DE9C54: mov ecx, 80020004h
  loc_00DE9C59: mov var_6C, ecx
  loc_00DE9C5C: mov eax, 0000000Ah
  loc_00DE9C61: mov var_5C, ecx
  loc_00DE9C64: mov edi, 00000008h
  loc_00DE9C69: lea edx, var_94
  loc_00DE9C6F: lea ecx, var_54
  loc_00DE9C72: mov var_74, eax
  loc_00DE9C75: mov var_64, eax
  loc_00DE9C78: mov var_8C, 006DCC68h ; "ERROR 01"
  loc_00DE9C82: mov var_94, edi
  loc_00DE9C88: call __vbaVarDup
  loc_00DE9C8A: mov var_7C, 006DCCC0h ; "Masukkan Password Terlebih Dahulu ! "
  loc_00DE9C91: jmp 00DE9F28h
  loc_00DE9C96: mov edx, [esi]
  loc_00DE9C98: push esi
  loc_00DE9C99: call [edx+00000324h]
  loc_00DE9C9F: push eax
  loc_00DE9CA0: lea eax, var_30
  loc_00DE9CA3: push eax
  loc_00DE9CA4: call edi
  loc_00DE9CA6: mov ecx, [eax]
  loc_00DE9CA8: lea edx, var_28
  loc_00DE9CAB: push edx
  loc_00DE9CAC: push eax
  loc_00DE9CAD: mov var_BC, eax
  loc_00DE9CB3: call [ecx+000000A0h]
  loc_00DE9CB9: test eax, eax
  loc_00DE9CBB: fnclex
  loc_00DE9CBD: jge 00DE9CD7h
  loc_00DE9CBF: mov ecx, var_BC
  loc_00DE9CC5: push 000000A0h
  loc_00DE9CCA: push 006DCB70h
  loc_00DE9CCF: push ecx
  loc_00DE9CD0: push eax
  loc_00DE9CD1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9CD7: mov edx, var_28
  loc_00DE9CDA: push edx
  loc_00DE9CDB: push 006DCC80h
  loc_00DE9CE0: call [00401104h] ; __vbaStrCmp
  loc_00DE9CE6: neg eax
  loc_00DE9CE8: sbb eax, eax
  loc_00DE9CEA: lea ecx, var_28
  loc_00DE9CED: inc eax
  loc_00DE9CEE: neg eax
  loc_00DE9CF0: mov var_C4, eax
  loc_00DE9CF6: call [00401258h] ; __vbaFreeStr
  loc_00DE9CFC: lea ecx, var_30
  loc_00DE9CFF: call ebx
  loc_00DE9D01: cmp var_C4, 0000h
  loc_00DE9D09: jz 00DE9F6Eh
  loc_00DE9D0F: mov eax, [esi]
  loc_00DE9D11: push esi
  loc_00DE9D12: call [eax+00000300h]
  loc_00DE9D18: lea ecx, var_30
  loc_00DE9D1B: push eax
  loc_00DE9D1C: push ecx
  loc_00DE9D1D: call edi
  loc_00DE9D1F: mov edx, [eax]
  loc_00DE9D21: push FFFFFFFFh
  loc_00DE9D23: push eax
  loc_00DE9D24: mov var_BC, eax
  loc_00DE9D2A: call [edx+0000005Ch]
  loc_00DE9D2D: test eax, eax
  loc_00DE9D2F: fnclex
  loc_00DE9D31: jge 00DE9D48h
  loc_00DE9D33: mov ecx, var_BC
  loc_00DE9D39: push 0000005Ch
  loc_00DE9D3B: push 006DCAE0h
  loc_00DE9D40: push ecx
  loc_00DE9D41: push eax
  loc_00DE9D42: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9D48: lea ecx, var_30
  loc_00DE9D4B: call ebx
  loc_00DE9D4D: mov ecx, 80020004h
  loc_00DE9D52: mov eax, 0000000Ah
  loc_00DE9D57: mov var_6C, ecx
  loc_00DE9D5A: mov var_5C, ecx
  loc_00DE9D5D: lea edx, var_94
  loc_00DE9D63: lea ecx, var_54
  loc_00DE9D66: mov var_74, eax
  loc_00DE9D69: mov var_64, eax
  loc_00DE9D6C: mov var_8C, 006DCC68h ; "ERROR 01"
  loc_00DE9D76: mov var_94, 00000008h
  loc_00DE9D80: call [004011F0h] ; __vbaVarDup
  loc_00DE9D86: lea edx, var_84
  loc_00DE9D8C: lea ecx, var_44
  loc_00DE9D8F: mov var_7C, 006DCBFCh ; "Username atau Password Yang Anda Masukkan Salah ! "
  loc_00DE9D96: mov var_84, 00000008h
  loc_00DE9DA0: call [004011F0h] ; __vbaVarDup
  loc_00DE9DA6: lea edx, var_74
  loc_00DE9DA9: lea eax, var_64
  loc_00DE9DAC: push edx
  loc_00DE9DAD: lea ecx, var_54
  loc_00DE9DB0: push eax
  loc_00DE9DB1: push ecx
  loc_00DE9DB2: lea edx, var_44
  loc_00DE9DB5: push 00000010h
  loc_00DE9DB7: push edx
  loc_00DE9DB8: call [004010A8h] ; rtcMsgBox
  loc_00DE9DBE: lea eax, var_74
  loc_00DE9DC1: lea ecx, var_64
  loc_00DE9DC4: push eax
  loc_00DE9DC5: lea edx, var_54
  loc_00DE9DC8: push ecx
  loc_00DE9DC9: lea eax, var_44
  loc_00DE9DCC: push edx
  loc_00DE9DCD: push eax
  loc_00DE9DCE: push 00000004h
  loc_00DE9DD0: call [00401038h] ; __vbaFreeVarList
  loc_00DE9DD6: mov ecx, [esi]
  loc_00DE9DD8: add esp, 00000014h
  loc_00DE9DDB: push esi
  loc_00DE9DDC: call [ecx+00000308h]
  loc_00DE9DE2: lea edx, var_30
  loc_00DE9DE5: push eax
  loc_00DE9DE6: push edx
  loc_00DE9DE7: call edi
  loc_00DE9DE9: mov ecx, [eax]
  loc_00DE9DEB: push 00000000h
  loc_00DE9DED: push eax
  loc_00DE9DEE: mov var_BC, eax
  loc_00DE9DF4: call [ecx+0000009Ch]
  loc_00DE9DFA: test eax, eax
  loc_00DE9DFC: fnclex
  loc_00DE9DFE: jge 00DE9E18h
  loc_00DE9E00: mov edx, var_BC
  loc_00DE9E06: push 0000009Ch
  loc_00DE9E0B: push 006DCAD0h
  loc_00DE9E10: push edx
  loc_00DE9E11: push eax
  loc_00DE9E12: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9E18: lea ecx, var_30
  loc_00DE9E1B: call ebx
  loc_00DE9E1D: mov eax, [esi]
  loc_00DE9E1F: push esi
  loc_00DE9E20: call [eax+00000320h]
  loc_00DE9E26: lea ecx, var_30
  loc_00DE9E29: push eax
  loc_00DE9E2A: push ecx
  loc_00DE9E2B: call edi
  loc_00DE9E2D: mov edx, [eax]
  loc_00DE9E2F: push FFFFFFFFh
  loc_00DE9E31: push eax
  loc_00DE9E32: mov var_BC, eax
  loc_00DE9E38: call [edx+0000009Ch]
  loc_00DE9E3E: test eax, eax
  loc_00DE9E40: fnclex
  loc_00DE9E42: jge 00DE9E5Ch
  loc_00DE9E44: mov ecx, var_BC
  loc_00DE9E4A: push 0000009Ch
  loc_00DE9E4F: push 006DCAD0h
  loc_00DE9E54: push ecx
  loc_00DE9E55: push eax
  loc_00DE9E56: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9E5C: lea ecx, var_30
  loc_00DE9E5F: call ebx
  loc_00DE9E61: mov edx, [esi]
  loc_00DE9E63: push esi
  loc_00DE9E64: call [edx+0000030Ch]
  loc_00DE9E6A: push eax
  loc_00DE9E6B: lea eax, var_30
  loc_00DE9E6E: push eax
  loc_00DE9E6F: call edi
  loc_00DE9E71: mov ecx, [eax]
  loc_00DE9E73: push 006DCC80h
  loc_00DE9E78: push eax
  loc_00DE9E79: mov var_BC, eax
  loc_00DE9E7F: call [ecx+000000A4h]
  loc_00DE9E85: test eax, eax
  loc_00DE9E87: fnclex
  loc_00DE9E89: jge 00DE9EA3h
  loc_00DE9E8B: mov edx, var_BC
  loc_00DE9E91: push 000000A4h
  loc_00DE9E96: push 006DCB70h
  loc_00DE9E9B: push edx
  loc_00DE9E9C: push eax
  loc_00DE9E9D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9EA3: lea ecx, var_30
  loc_00DE9EA6: call ebx
  loc_00DE9EA8: mov eax, [esi]
  loc_00DE9EAA: push esi
  loc_00DE9EAB: call [eax+00000324h]
  loc_00DE9EB1: lea ecx, var_30
  loc_00DE9EB4: push eax
  loc_00DE9EB5: push ecx
  loc_00DE9EB6: call edi
  loc_00DE9EB8: mov esi, eax
  loc_00DE9EBA: push 006DCC80h
  loc_00DE9EBF: push esi
  loc_00DE9EC0: mov edx, [esi]
  loc_00DE9EC2: call [edx+000000A4h]
  loc_00DE9EC8: test eax, eax
  loc_00DE9ECA: fnclex
  loc_00DE9ECC: jge 00DE9EE0h
  loc_00DE9ECE: push 000000A4h
  loc_00DE9ED3: push 006DCB70h
  loc_00DE9ED8: push esi
  loc_00DE9ED9: push eax
  loc_00DE9EDA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9EE0: lea ecx, var_30
  loc_00DE9EE3: call ebx
  loc_00DE9EE5: mov esi, [004011F0h] ; __vbaVarDup
  loc_00DE9EEB: mov ecx, 80020004h
  loc_00DE9EF0: mov var_6C, ecx
  loc_00DE9EF3: mov eax, 0000000Ah
  loc_00DE9EF8: mov var_5C, ecx
  loc_00DE9EFB: mov edi, 00000008h
  loc_00DE9F00: lea edx, var_94
  loc_00DE9F06: lea ecx, var_54
  loc_00DE9F09: mov var_74, eax
  loc_00DE9F0C: mov var_64, eax
  loc_00DE9F0F: mov var_8C, 006DCC68h ; "ERROR 01"
  loc_00DE9F19: mov var_94, edi
  loc_00DE9F1F: call __vbaVarDup
  loc_00DE9F21: mov var_7C, 006DCD10h ; "Masukkan Username Terlebih Dahulu ! "
  loc_00DE9F28: lea edx, var_84
  loc_00DE9F2E: lea ecx, var_44
  loc_00DE9F31: mov var_84, edi
  loc_00DE9F37: call __vbaVarDup
  loc_00DE9F39: lea eax, var_74
  loc_00DE9F3C: lea ecx, var_64
  loc_00DE9F3F: push eax
  loc_00DE9F40: lea edx, var_54
  loc_00DE9F43: push ecx
  loc_00DE9F44: push edx
  loc_00DE9F45: lea eax, var_44
  loc_00DE9F48: push 00000010h
  loc_00DE9F4A: push eax
  loc_00DE9F4B: call [004010A8h] ; rtcMsgBox
  loc_00DE9F51: lea ecx, var_74
  loc_00DE9F54: lea edx, var_64
  loc_00DE9F57: push ecx
  loc_00DE9F58: lea eax, var_54
  loc_00DE9F5B: push edx
  loc_00DE9F5C: lea ecx, var_44
  loc_00DE9F5F: push eax
  loc_00DE9F60: push ecx
  loc_00DE9F61: push 00000004h
  loc_00DE9F63: call [00401038h] ; __vbaFreeVarList
  loc_00DE9F69: add esp, 00000014h
  loc_00DE9F6C: jmp 00DE9FA2h
  loc_00DE9F6E: mov edx, [esi]
  loc_00DE9F70: push esi
  loc_00DE9F71: call [edx+00000304h]
  loc_00DE9F77: push eax
  loc_00DE9F78: lea eax, var_30
  loc_00DE9F7B: push eax
  loc_00DE9F7C: call edi
  loc_00DE9F7E: mov esi, eax
  loc_00DE9F80: push FFFFFFFFh
  loc_00DE9F82: push esi
  loc_00DE9F83: mov ecx, [esi]
  loc_00DE9F85: call [ecx+0000005Ch]
  loc_00DE9F88: test eax, eax
  loc_00DE9F8A: fnclex
  loc_00DE9F8C: jge 00DE9F9Dh
  loc_00DE9F8E: push 0000005Ch
  loc_00DE9F90: push 006DCAE0h
  loc_00DE9F95: push esi
  loc_00DE9F96: push eax
  loc_00DE9F97: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9F9D: lea ecx, var_30
  loc_00DE9FA0: call ebx
  loc_00DE9FA2: mov var_4, 00000000h
  loc_00DE9FA9: push 00DE9FF6h
  loc_00DE9FAE: jmp 00DE9FECh
  loc_00DE9FB0: lea edx, var_2C
  loc_00DE9FB3: lea eax, var_28
  loc_00DE9FB6: push edx
  loc_00DE9FB7: push eax
  loc_00DE9FB8: push 00000002h
  loc_00DE9FBA: call [004011BCh] ; __vbaFreeStrList
  loc_00DE9FC0: lea ecx, var_34
  loc_00DE9FC3: lea edx, var_30
  loc_00DE9FC6: push ecx
  loc_00DE9FC7: push edx
  loc_00DE9FC8: push 00000002h
  loc_00DE9FCA: call [00401048h] ; __vbaFreeObjList
  loc_00DE9FD0: lea eax, var_74
  loc_00DE9FD3: lea ecx, var_64
  loc_00DE9FD6: push eax
  loc_00DE9FD7: lea edx, var_54
  loc_00DE9FDA: push ecx
  loc_00DE9FDB: lea eax, var_44
  loc_00DE9FDE: push edx
  loc_00DE9FDF: push eax
  loc_00DE9FE0: push 00000004h
  loc_00DE9FE2: call [00401038h] ; __vbaFreeVarList
  loc_00DE9FE8: add esp, 0000002Ch
  loc_00DE9FEB: ret
  loc_00DE9FEC: lea ecx, var_24
  loc_00DE9FEF: call [00401024h] ; __vbaFreeVar
  loc_00DE9FF5: ret
  loc_00DE9FF6: mov eax, Me
  loc_00DE9FF9: push eax
  loc_00DE9FFA: mov ecx, [eax]
  loc_00DE9FFC: call [ecx+00000008h]
  loc_00DE9FFF: mov eax, var_4
  loc_00DEA002: mov ecx, var_14
  loc_00DEA005: pop edi
  loc_00DEA006: pop esi
  loc_00DEA007: mov fs:[00000000h], ecx
  loc_00DEA00E: pop ebx
  loc_00DEA00F: mov esp, ebp
  loc_00DEA011: pop ebp
  loc_00DEA012: retn 0004h
End Sub

Private Sub back_Click() 'DE8FA0
  loc_00DE8FA0: push ebp
  loc_00DE8FA1: mov ebp, esp
  loc_00DE8FA3: sub esp, 0000000Ch
  loc_00DE8FA6: push 00402836h ; __vbaExceptHandler
  loc_00DE8FAB: mov eax, fs:[00000000h]
  loc_00DE8FB1: push eax
  loc_00DE8FB2: mov fs:[00000000h], esp
  loc_00DE8FB9: sub esp, 00000034h
  loc_00DE8FBC: push ebx
  loc_00DE8FBD: push esi
  loc_00DE8FBE: push edi
  loc_00DE8FBF: mov var_C, esp
  loc_00DE8FC2: mov var_8, 00401278h
  loc_00DE8FC9: mov edi, Me
  loc_00DE8FCC: mov eax, edi
  loc_00DE8FCE: and eax, 00000001h
  loc_00DE8FD1: mov var_4, eax
  loc_00DE8FD4: and edi, FFFFFFFEh
  loc_00DE8FD7: push edi
  loc_00DE8FD8: mov Me, edi
  loc_00DE8FDB: mov ecx, [edi]
  loc_00DE8FDD: call [ecx+00000004h]
  loc_00DE8FE0: mov eax, [00E3D9CCh]
  loc_00DE8FE5: mov var_18, 00000000h
  loc_00DE8FEC: test eax, eax
  loc_00DE8FEE: jnz 00DE9000h
  loc_00DE8FF0: push 00E3D9CCh
  loc_00DE8FF5: push 006DC960h
  loc_00DE8FFA: call [004011A0h] ; __vbaNew2
  loc_00DE9000: mov esi, [00E3D9CCh]
  loc_00DE9006: lea edx, var_18
  loc_00DE9009: push edi
  loc_00DE900A: push edx
  loc_00DE900B: mov ebx, [esi]
  loc_00DE900D: call [004010B4h] ; __vbaObjSetAddref
  loc_00DE9013: push eax
  loc_00DE9014: push esi
  loc_00DE9015: call [ebx+00000010h]
  loc_00DE9018: test eax, eax
  loc_00DE901A: fnclex
  loc_00DE901C: jge 00DE902Dh
  loc_00DE901E: push 00000010h
  loc_00DE9020: push 006DC950h
  loc_00DE9025: push esi
  loc_00DE9026: push eax
  loc_00DE9027: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE902D: lea ecx, var_18
  loc_00DE9030: call [00401254h] ; __vbaFreeObj
  loc_00DE9036: mov eax, [00E3D024h]
  loc_00DE903B: test eax, eax
  loc_00DE903D: jnz 00DE904Fh
  loc_00DE903F: push 00E3D024h
  loc_00DE9044: push 006CE120h
  loc_00DE9049: call [004011A0h] ; __vbaNew2
  loc_00DE904F: sub esp, 00000010h
  loc_00DE9052: mov ecx, 0000000Ah
  loc_00DE9057: mov ebx, esp
  loc_00DE9059: mov var_28, ecx
  loc_00DE905C: mov eax, 80020004h
  loc_00DE9061: sub esp, 00000010h
  loc_00DE9064: mov [ebx], ecx
  loc_00DE9066: mov ecx, var_34
  loc_00DE9069: mov edx, eax
  loc_00DE906B: mov esi, [00E3D024h]
  loc_00DE9071: mov [ebx+00000004h], ecx
  loc_00DE9074: mov ecx, esp
  loc_00DE9076: mov edi, [esi]
  loc_00DE9078: push esi
  loc_00DE9079: mov [ebx+00000008h], eax
  loc_00DE907C: mov eax, var_2C
  loc_00DE907F: mov [ebx+0000000Ch], eax
  loc_00DE9082: mov eax, var_28
  loc_00DE9085: mov [ecx], eax
  loc_00DE9087: mov eax, var_24
  loc_00DE908A: mov [ecx+00000004h], eax
  loc_00DE908D: mov [ecx+00000008h], edx
  loc_00DE9090: mov edx, var_1C
  loc_00DE9093: mov [ecx+0000000Ch], edx
  loc_00DE9096: call [edi+000002B0h]
  loc_00DE909C: test eax, eax
  loc_00DE909E: fnclex
  loc_00DE90A0: jge 00DE90B4h
  loc_00DE90A2: push 000002B0h
  loc_00DE90A7: push 006DC970h
  loc_00DE90AC: push esi
  loc_00DE90AD: push eax
  loc_00DE90AE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE90B4: mov var_4, 00000000h
  loc_00DE90BB: push 00DE90CDh
  loc_00DE90C0: jmp 00DE90CCh
  loc_00DE90C2: lea ecx, var_18
  loc_00DE90C5: call [00401254h] ; __vbaFreeObj
  loc_00DE90CB: ret
  loc_00DE90CC: ret
  loc_00DE90CD: mov eax, Me
  loc_00DE90D0: push eax
  loc_00DE90D1: mov ecx, [eax]
  loc_00DE90D3: call [ecx+00000008h]
  loc_00DE90D6: mov eax, var_4
  loc_00DE90D9: mov ecx, var_14
  loc_00DE90DC: pop edi
  loc_00DE90DD: pop esi
  loc_00DE90DE: mov fs:[00000000h], ecx
  loc_00DE90E5: pop ebx
  loc_00DE90E6: mov esp, ebp
  loc_00DE90E8: pop ebp
  loc_00DE90E9: retn 0004h
End Sub

Private Sub o_Click() 'DEA020
  loc_00DEA020: push ebp
  loc_00DEA021: mov ebp, esp
  loc_00DEA023: sub esp, 0000000Ch
  loc_00DEA026: push 00402836h ; __vbaExceptHandler
  loc_00DEA02B: mov eax, fs:[00000000h]
  loc_00DEA031: push eax
  loc_00DEA032: mov fs:[00000000h], esp
  loc_00DEA039: sub esp, 000000ACh
  loc_00DEA03F: push ebx
  loc_00DEA040: push esi
  loc_00DEA041: push edi
  loc_00DEA042: mov var_C, esp
  loc_00DEA045: mov var_8, 004012C0h
  loc_00DEA04C: mov esi, Me
  loc_00DEA04F: mov eax, esi
  loc_00DEA051: and eax, 00000001h
  loc_00DEA054: mov var_4, eax
  loc_00DEA057: and esi, FFFFFFFEh
  loc_00DEA05A: push esi
  loc_00DEA05B: mov Me, esi
  loc_00DEA05E: mov ecx, [esi]
  loc_00DEA060: call [ecx+00000004h]
  loc_00DEA063: mov edx, [esi]
  loc_00DEA065: xor ebx, ebx
  loc_00DEA067: push esi
  loc_00DEA068: mov var_24, ebx
  loc_00DEA06B: mov var_28, ebx
  loc_00DEA06E: mov var_2C, ebx
  loc_00DEA071: mov var_3C, ebx
  loc_00DEA074: mov var_4C, ebx
  loc_00DEA077: mov var_5C, ebx
  loc_00DEA07A: mov var_6C, ebx
  loc_00DEA07D: mov var_7C, ebx
  loc_00DEA080: mov var_8C, ebx
  loc_00DEA086: call [edx+0000030Ch]
  loc_00DEA08C: push eax
  loc_00DEA08D: lea eax, var_2C
  loc_00DEA090: push eax
  loc_00DEA091: call [004010ACh] ; __vbaObjSet
  loc_00DEA097: mov edi, eax
  loc_00DEA099: lea edx, var_28
  loc_00DEA09C: push edx
  loc_00DEA09D: push edi
  loc_00DEA09E: mov ecx, [edi]
  loc_00DEA0A0: call [ecx+00000158h]
  loc_00DEA0A6: cmp eax, ebx
  loc_00DEA0A8: fnclex
  loc_00DEA0AA: jge 00DEA0BEh
  loc_00DEA0AC: push 00000158h
  loc_00DEA0B1: push 006DCB70h
  loc_00DEA0B6: push edi
  loc_00DEA0B7: push eax
  loc_00DEA0B8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA0BE: mov eax, var_28
  loc_00DEA0C1: push eax
  loc_00DEA0C2: push 006DCD60h
  loc_00DEA0C7: call [00401104h] ; __vbaStrCmp
  loc_00DEA0CD: mov edi, eax
  loc_00DEA0CF: lea ecx, var_28
  loc_00DEA0D2: neg edi
  loc_00DEA0D4: sbb edi, edi
  loc_00DEA0D6: inc edi
  loc_00DEA0D7: neg edi
  loc_00DEA0D9: call [00401258h] ; __vbaFreeStr
  loc_00DEA0DF: lea ecx, var_2C
  loc_00DEA0E2: call [00401254h] ; __vbaFreeObj
  loc_00DEA0E8: cmp di, bx
  loc_00DEA0EB: jz 00DEA1D1h
  loc_00DEA0F1: mov ecx, [esi]
  loc_00DEA0F3: push esi
  loc_00DEA0F4: call [ecx+0000030Ch]
  loc_00DEA0FA: lea edx, var_2C
  loc_00DEA0FD: push eax
  loc_00DEA0FE: push edx
  loc_00DEA0FF: call [004010ACh] ; __vbaObjSet
  loc_00DEA105: mov edi, eax
  loc_00DEA107: lea eax, var_24
  loc_00DEA10A: lea ecx, var_28
  loc_00DEA10D: push eax
  loc_00DEA10E: mov ebx, [edi]
  loc_00DEA110: push ecx
  loc_00DEA111: call [00401180h] ; __vbaStrVarVal
  loc_00DEA117: push eax
  loc_00DEA118: push edi
  loc_00DEA119: call [ebx+0000015Ch]
  loc_00DEA11F: xor ebx, ebx
  loc_00DEA121: cmp eax, ebx
  loc_00DEA123: fnclex
  loc_00DEA125: jge 00DEA139h
  loc_00DEA127: push 0000015Ch
  loc_00DEA12C: push 006DCB70h
  loc_00DEA131: push edi
  loc_00DEA132: push eax
  loc_00DEA133: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA139: lea ecx, var_28
  loc_00DEA13C: call [00401258h] ; __vbaFreeStr
  loc_00DEA142: lea ecx, var_2C
  loc_00DEA145: call [00401254h] ; __vbaFreeObj
  loc_00DEA14B: mov edx, [esi]
  loc_00DEA14D: push esi
  loc_00DEA14E: call [edx+00000314h]
  loc_00DEA154: push eax
  loc_00DEA155: lea eax, var_2C
  loc_00DEA158: push eax
  loc_00DEA159: call [004010ACh] ; __vbaObjSet
  loc_00DEA15F: mov edi, eax
  loc_00DEA161: push ebx
  loc_00DEA162: push edi
  loc_00DEA163: mov ecx, [edi]
  loc_00DEA165: call [ecx+0000008Ch]
  loc_00DEA16B: cmp eax, ebx
  loc_00DEA16D: fnclex
  loc_00DEA16F: jge 00DEA183h
  loc_00DEA171: push 0000008Ch
  loc_00DEA176: push 006DCAC0h
  loc_00DEA17B: push edi
  loc_00DEA17C: push eax
  loc_00DEA17D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA183: mov edi, [00401254h] ; __vbaFreeObj
  loc_00DEA189: lea ecx, var_2C
  loc_00DEA18C: call edi
  loc_00DEA18E: mov edx, [esi]
  loc_00DEA190: push esi
  loc_00DEA191: call [edx+00000318h]
  loc_00DEA197: push eax
  loc_00DEA198: lea eax, var_2C
  loc_00DEA19B: push eax
  loc_00DEA19C: call [004010ACh] ; __vbaObjSet
  loc_00DEA1A2: mov esi, eax
  loc_00DEA1A4: push FFFFFFFFh
  loc_00DEA1A6: push esi
  loc_00DEA1A7: mov ecx, [esi]
  loc_00DEA1A9: call [ecx+0000008Ch]
  loc_00DEA1AF: cmp eax, ebx
  loc_00DEA1B1: fnclex
  loc_00DEA1B3: jge 00DEA1C7h
  loc_00DEA1B5: push 0000008Ch
  loc_00DEA1BA: push 006DCAC0h
  loc_00DEA1BF: push esi
  loc_00DEA1C0: push eax
  loc_00DEA1C1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA1C7: lea ecx, var_2C
  loc_00DEA1CA: call edi
  loc_00DEA1CC: jmp 00DEA252h
  loc_00DEA1D1: mov esi, [004011F0h] ; __vbaVarDup
  loc_00DEA1D7: mov ecx, 80020004h
  loc_00DEA1DC: mov var_64, ecx
  loc_00DEA1DF: mov eax, 0000000Ah
  loc_00DEA1E4: mov var_54, ecx
  loc_00DEA1E7: mov edi, 00000008h
  loc_00DEA1EC: lea edx, var_8C
  loc_00DEA1F2: lea ecx, var_4C
  loc_00DEA1F5: mov var_6C, eax
  loc_00DEA1F8: mov var_5C, eax
  loc_00DEA1FB: mov var_84, 006DCD7Ch ; "Beware OF THE DOG"
  loc_00DEA205: mov var_8C, edi
  loc_00DEA20B: call __vbaVarDup
  loc_00DEA20D: lea edx, var_7C
  loc_00DEA210: lea ecx, var_3C
  loc_00DEA213: mov var_74, 006DCD68h ; " error"
  loc_00DEA21A: mov var_7C, edi
  loc_00DEA21D: call __vbaVarDup
  loc_00DEA21F: lea edx, var_6C
  loc_00DEA222: lea eax, var_5C
  loc_00DEA225: push edx
  loc_00DEA226: lea ecx, var_4C
  loc_00DEA229: push eax
  loc_00DEA22A: push ecx
  loc_00DEA22B: lea edx, var_3C
  loc_00DEA22E: push 00000016h
  loc_00DEA230: push edx
  loc_00DEA231: call [004010A8h] ; rtcMsgBox
  loc_00DEA237: lea eax, var_6C
  loc_00DEA23A: lea ecx, var_5C
  loc_00DEA23D: push eax
  loc_00DEA23E: lea edx, var_4C
  loc_00DEA241: push ecx
  loc_00DEA242: lea eax, var_3C
  loc_00DEA245: push edx
  loc_00DEA246: push eax
  loc_00DEA247: push 00000004h
  loc_00DEA249: call [00401038h] ; __vbaFreeVarList
  loc_00DEA24F: add esp, 00000014h
  loc_00DEA252: mov var_4, ebx
  loc_00DEA255: push 00DEA294h
  loc_00DEA25A: jmp 00DEA28Ah
  loc_00DEA25C: lea ecx, var_28
  loc_00DEA25F: call [00401258h] ; __vbaFreeStr
  loc_00DEA265: lea ecx, var_2C
  loc_00DEA268: call [00401254h] ; __vbaFreeObj
  loc_00DEA26E: lea ecx, var_6C
  loc_00DEA271: lea edx, var_5C
  loc_00DEA274: push ecx
  loc_00DEA275: lea eax, var_4C
  loc_00DEA278: push edx
  loc_00DEA279: lea ecx, var_3C
  loc_00DEA27C: push eax
  loc_00DEA27D: push ecx
  loc_00DEA27E: push 00000004h
  loc_00DEA280: call [00401038h] ; __vbaFreeVarList
  loc_00DEA286: add esp, 00000014h
  loc_00DEA289: ret
  loc_00DEA28A: lea ecx, var_24
  loc_00DEA28D: call [00401024h] ; __vbaFreeVar
  loc_00DEA293: ret
  loc_00DEA294: mov eax, Me
  loc_00DEA297: push eax
  loc_00DEA298: mov edx, [eax]
  loc_00DEA29A: call [edx+00000008h]
  loc_00DEA29D: mov eax, var_4
  loc_00DEA2A0: mov ecx, var_14
  loc_00DEA2A3: pop edi
  loc_00DEA2A4: pop esi
  loc_00DEA2A5: mov fs:[00000000h], ecx
  loc_00DEA2AC: pop ebx
  loc_00DEA2AD: mov esp, ebp
  loc_00DEA2AF: pop ebp
  loc_00DEA2B0: retn 0004h
End Sub

Private Sub x_Click() 'DEB390
  loc_00DEB390: push ebp
  loc_00DEB391: mov ebp, esp
  loc_00DEB393: sub esp, 0000000Ch
  loc_00DEB396: push 00402836h ; __vbaExceptHandler
  loc_00DEB39B: mov eax, fs:[00000000h]
  loc_00DEB3A1: push eax
  loc_00DEB3A2: mov fs:[00000000h], esp
  loc_00DEB3A9: sub esp, 000000B0h
  loc_00DEB3AF: push ebx
  loc_00DEB3B0: push esi
  loc_00DEB3B1: push edi
  loc_00DEB3B2: mov var_C, esp
  loc_00DEB3B5: mov var_8, 00401328h
  loc_00DEB3BC: mov esi, Me
  loc_00DEB3BF: mov eax, esi
  loc_00DEB3C1: and eax, 00000001h
  loc_00DEB3C4: mov var_4, eax
  loc_00DEB3C7: and esi, FFFFFFFEh
  loc_00DEB3CA: push esi
  loc_00DEB3CB: mov Me, esi
  loc_00DEB3CE: mov ecx, [esi]
  loc_00DEB3D0: call [ecx+00000004h]
  loc_00DEB3D3: mov edx, [esi]
  loc_00DEB3D5: xor eax, eax
  loc_00DEB3D7: push esi
  loc_00DEB3D8: mov var_24, eax
  loc_00DEB3DB: mov var_28, eax
  loc_00DEB3DE: mov var_2C, eax
  loc_00DEB3E1: mov var_3C, eax
  loc_00DEB3E4: mov var_4C, eax
  loc_00DEB3E7: mov var_5C, eax
  loc_00DEB3EA: mov var_6C, eax
  loc_00DEB3ED: mov var_7C, eax
  loc_00DEB3F0: mov var_8C, eax
  loc_00DEB3F6: call [edx+0000030Ch]
  loc_00DEB3FC: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00DEB402: push eax
  loc_00DEB403: lea eax, var_2C
  loc_00DEB406: push eax
  loc_00DEB407: call ebx
  loc_00DEB409: mov edi, eax
  loc_00DEB40B: lea edx, var_28
  loc_00DEB40E: push edx
  loc_00DEB40F: push edi
  loc_00DEB410: mov ecx, [edi]
  loc_00DEB412: call [ecx+00000158h]
  loc_00DEB418: test eax, eax
  loc_00DEB41A: fnclex
  loc_00DEB41C: jge 00DEB430h
  loc_00DEB41E: push 00000158h
  loc_00DEB423: push 006DCB70h
  loc_00DEB428: push edi
  loc_00DEB429: push eax
  loc_00DEB42A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB430: mov eax, var_28
  loc_00DEB433: lea ecx, var_24
  loc_00DEB436: mov var_34, eax
  loc_00DEB439: lea eax, var_3C
  loc_00DEB43C: push eax
  loc_00DEB43D: push ecx
  loc_00DEB43E: mov var_28, 00000000h
  loc_00DEB445: mov var_3C, 00008008h
  loc_00DEB44C: call [00401108h] ; __vbaVarTstEq
  loc_00DEB452: lea ecx, var_2C
  loc_00DEB455: mov edi, eax
  loc_00DEB457: call [00401254h] ; __vbaFreeObj
  loc_00DEB45D: lea ecx, var_3C
  loc_00DEB460: call [00401024h] ; __vbaFreeVar
  loc_00DEB466: test di, di
  loc_00DEB469: jz 00DEB52Fh
  loc_00DEB46F: mov edx, [esi]
  loc_00DEB471: push esi
  loc_00DEB472: call [edx+0000030Ch]
  loc_00DEB478: push eax
  loc_00DEB479: lea eax, var_2C
  loc_00DEB47C: push eax
  loc_00DEB47D: call ebx
  loc_00DEB47F: mov edi, eax
  loc_00DEB481: push 006DCD60h
  loc_00DEB486: push edi
  loc_00DEB487: mov ecx, [edi]
  loc_00DEB489: call [ecx+0000015Ch]
  loc_00DEB48F: test eax, eax
  loc_00DEB491: fnclex
  loc_00DEB493: jge 00DEB4A7h
  loc_00DEB495: push 0000015Ch
  loc_00DEB49A: push 006DCB70h
  loc_00DEB49F: push edi
  loc_00DEB4A0: push eax
  loc_00DEB4A1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB4A7: lea ecx, var_2C
  loc_00DEB4AA: call [00401254h] ; __vbaFreeObj
  loc_00DEB4B0: mov edx, [esi]
  loc_00DEB4B2: push esi
  loc_00DEB4B3: call [edx+00000314h]
  loc_00DEB4B9: push eax
  loc_00DEB4BA: lea eax, var_2C
  loc_00DEB4BD: push eax
  loc_00DEB4BE: call ebx
  loc_00DEB4C0: mov edi, eax
  loc_00DEB4C2: push FFFFFFFFh
  loc_00DEB4C4: push edi
  loc_00DEB4C5: mov ecx, [edi]
  loc_00DEB4C7: call [ecx+0000008Ch]
  loc_00DEB4CD: test eax, eax
  loc_00DEB4CF: fnclex
  loc_00DEB4D1: jge 00DEB4E5h
  loc_00DEB4D3: push 0000008Ch
  loc_00DEB4D8: push 006DCAC0h
  loc_00DEB4DD: push edi
  loc_00DEB4DE: push eax
  loc_00DEB4DF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB4E5: mov edi, [00401254h] ; __vbaFreeObj
  loc_00DEB4EB: lea ecx, var_2C
  loc_00DEB4EE: call edi
  loc_00DEB4F0: mov edx, [esi]
  loc_00DEB4F2: push esi
  loc_00DEB4F3: call [edx+00000318h]
  loc_00DEB4F9: push eax
  loc_00DEB4FA: lea eax, var_2C
  loc_00DEB4FD: push eax
  loc_00DEB4FE: call ebx
  loc_00DEB500: mov esi, eax
  loc_00DEB502: push 00000000h
  loc_00DEB504: push esi
  loc_00DEB505: mov ecx, [esi]
  loc_00DEB507: call [ecx+0000008Ch]
  loc_00DEB50D: test eax, eax
  loc_00DEB50F: fnclex
  loc_00DEB511: jge 00DEB525h
  loc_00DEB513: push 0000008Ch
  loc_00DEB518: push 006DCAC0h
  loc_00DEB51D: push esi
  loc_00DEB51E: push eax
  loc_00DEB51F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB525: lea ecx, var_2C
  loc_00DEB528: call edi
  loc_00DEB52A: jmp 00DEB5B0h
  loc_00DEB52F: mov esi, [004011F0h] ; __vbaVarDup
  loc_00DEB535: mov ecx, 80020004h
  loc_00DEB53A: mov var_64, ecx
  loc_00DEB53D: mov eax, 0000000Ah
  loc_00DEB542: mov var_54, ecx
  loc_00DEB545: mov edi, 00000008h
  loc_00DEB54A: lea edx, var_8C
  loc_00DEB550: lea ecx, var_4C
  loc_00DEB553: mov var_6C, eax
  loc_00DEB556: mov var_5C, eax
  loc_00DEB559: mov var_84, 006DCD7Ch ; "Beware OF THE DOG"
  loc_00DEB563: mov var_8C, edi
  loc_00DEB569: call __vbaVarDup
  loc_00DEB56B: lea edx, var_7C
  loc_00DEB56E: lea ecx, var_3C
  loc_00DEB571: mov var_74, 006DCD68h ; " error"
  loc_00DEB578: mov var_7C, edi
  loc_00DEB57B: call __vbaVarDup
  loc_00DEB57D: lea edx, var_6C
  loc_00DEB580: lea eax, var_5C
  loc_00DEB583: push edx
  loc_00DEB584: lea ecx, var_4C
  loc_00DEB587: push eax
  loc_00DEB588: push ecx
  loc_00DEB589: lea edx, var_3C
  loc_00DEB58C: push 00000016h
  loc_00DEB58E: push edx
  loc_00DEB58F: call [004010A8h] ; rtcMsgBox
  loc_00DEB595: lea eax, var_6C
  loc_00DEB598: lea ecx, var_5C
  loc_00DEB59B: push eax
  loc_00DEB59C: lea edx, var_4C
  loc_00DEB59F: push ecx
  loc_00DEB5A0: lea eax, var_3C
  loc_00DEB5A3: push edx
  loc_00DEB5A4: push eax
  loc_00DEB5A5: push 00000004h
  loc_00DEB5A7: call [00401038h] ; __vbaFreeVarList
  loc_00DEB5AD: add esp, 00000014h
  loc_00DEB5B0: mov var_4, 00000000h
  loc_00DEB5B7: push 00DEB5F6h
  loc_00DEB5BC: jmp 00DEB5ECh
  loc_00DEB5BE: lea ecx, var_28
  loc_00DEB5C1: call [00401258h] ; __vbaFreeStr
  loc_00DEB5C7: lea ecx, var_2C
  loc_00DEB5CA: call [00401254h] ; __vbaFreeObj
  loc_00DEB5D0: lea ecx, var_6C
  loc_00DEB5D3: lea edx, var_5C
  loc_00DEB5D6: push ecx
  loc_00DEB5D7: lea eax, var_4C
  loc_00DEB5DA: push edx
  loc_00DEB5DB: lea ecx, var_3C
  loc_00DEB5DE: push eax
  loc_00DEB5DF: push ecx
  loc_00DEB5E0: push 00000004h
  loc_00DEB5E2: call [00401038h] ; __vbaFreeVarList
  loc_00DEB5E8: add esp, 00000014h
  loc_00DEB5EB: ret
  loc_00DEB5EC: lea ecx, var_24
  loc_00DEB5EF: call [00401024h] ; __vbaFreeVar
  loc_00DEB5F5: ret
  loc_00DEB5F6: mov eax, Me
  loc_00DEB5F9: push eax
  loc_00DEB5FA: mov edx, [eax]
  loc_00DEB5FC: call [edx+00000008h]
  loc_00DEB5FF: mov eax, var_4
  loc_00DEB602: mov ecx, var_14
  loc_00DEB605: pop edi
  loc_00DEB606: pop esi
  loc_00DEB607: mov fs:[00000000h], ecx
  loc_00DEB60E: pop ebx
  loc_00DEB60F: mov esp, ebp
  loc_00DEB611: pop ebp
  loc_00DEB612: retn 0004h
End Sub

Private Sub Timer1_Timer() 'DEA2C0
  loc_00DEA2C0: push ebp
  loc_00DEA2C1: mov ebp, esp
  loc_00DEA2C3: sub esp, 0000000Ch
  loc_00DEA2C6: push 00402836h ; __vbaExceptHandler
  loc_00DEA2CB: mov eax, fs:[00000000h]
  loc_00DEA2D1: push eax
  loc_00DEA2D2: mov fs:[00000000h], esp
  loc_00DEA2D9: sub esp, 00000064h
  loc_00DEA2DC: push ebx
  loc_00DEA2DD: push esi
  loc_00DEA2DE: push edi
  loc_00DEA2DF: mov var_C, esp
  loc_00DEA2E2: mov var_8, 004012E8h
  loc_00DEA2E9: mov ebx, Me
  loc_00DEA2EC: mov eax, ebx
  loc_00DEA2EE: and eax, 00000001h
  loc_00DEA2F1: mov var_4, eax
  loc_00DEA2F4: and ebx, FFFFFFFEh
  loc_00DEA2F7: push ebx
  loc_00DEA2F8: mov Me, ebx
  loc_00DEA2FB: mov ecx, [ebx]
  loc_00DEA2FD: call [ecx+00000004h]
  loc_00DEA300: xor edi, edi
  loc_00DEA302: lea esi, [ebx+00000034h]
  loc_00DEA305: mov var_40, edi
  loc_00DEA308: lea edx, var_40
  loc_00DEA30B: mov ecx, esi
  loc_00DEA30D: mov var_18, edi
  loc_00DEA310: mov var_1C, edi
  loc_00DEA313: mov var_20, edi
  loc_00DEA316: mov var_30, edi
  loc_00DEA319: mov var_50, edi
  loc_00DEA31C: mov var_54, edi
  loc_00DEA31F: mov var_38, edi
  loc_00DEA322: mov var_40, 00000002h
  loc_00DEA329: call [0040101Ch] ; __vbaVarMove
  loc_00DEA32F: lea eax, [ebx+00000044h]
  loc_00DEA332: lea edx, var_40
  loc_00DEA335: push eax
  loc_00DEA336: push edx
  loc_00DEA337: mov var_38, 00000064h
  loc_00DEA33E: mov var_40, 00008002h
  loc_00DEA345: mov var_70, eax
  loc_00DEA348: call [00401088h] ; __vbaVarTstLe
  loc_00DEA34E: test ax, ax
  loc_00DEA351: jz 00DEA9E4h
  loc_00DEA357: lea eax, var_30
  loc_00DEA35A: mov var_28, 80020004h
  loc_00DEA361: push eax
  loc_00DEA362: mov var_30, 0000000Ah
  loc_00DEA369: call [004010A0h] ; rtcRandomize
  loc_00DEA36F: mov edi, [00401024h] ; __vbaFreeVar
  loc_00DEA375: lea ecx, var_30
  loc_00DEA378: call edi
  loc_00DEA37A: lea ecx, var_30
  loc_00DEA37D: mov var_28, 80020004h
  loc_00DEA384: push ecx
  loc_00DEA385: mov var_30, 0000000Ah
  loc_00DEA38C: call [0040109Ch] ; rtcRandomNext
  loc_00DEA392: fstp real4 ptr var_54
  loc_00DEA395: fld real4 ptr var_54
  loc_00DEA398: fmul st0, real4 ptr [004012E4h]
  loc_00DEA39E: lea ecx, var_30
  loc_00DEA3A1: fadd st0, real4 ptr [004012E0h]
  loc_00DEA3A7: fstp real8 ptr [ebx+00000054h]
  loc_00DEA3AA: fnstsw ax
  loc_00DEA3AC: test al, 0Dh
  loc_00DEA3AE: jnz 00DEAA35h
  loc_00DEA3B4: call edi
  loc_00DEA3B6: fld real8 ptr [ebx+00000054h]
  loc_00DEA3B9: cmp [00E3D000h], 00000000h
  loc_00DEA3C0: jnz 00DEA3CAh
  loc_00DEA3C2: fdiv st0, real8 ptr [004012D8h]
  loc_00DEA3C8: jmp 00DEA3DBh
  loc_00DEA3CA: push [004012DCh]
  loc_00DEA3D0: push [004012D8h]
  loc_00DEA3D6: call 00402854h ; _adj_fdiv_m64
  loc_00DEA3DB: lea edx, var_40
  loc_00DEA3DE: mov ecx, esi
  loc_00DEA3E0: mov var_40, 00000005h
  loc_00DEA3E7: fstp real8 ptr var_38
  loc_00DEA3EA: fnstsw ax
  loc_00DEA3EC: test al, 0Dh
  loc_00DEA3EE: jnz 00DEAA35h
  loc_00DEA3F4: call [0040101Ch] ; __vbaVarMove
  loc_00DEA3FA: push 00000000h
  loc_00DEA3FC: lea edx, var_30
  loc_00DEA3FF: push esi
  loc_00DEA400: push edx
  loc_00DEA401: call [00401170h] ; rtcRound
  loc_00DEA407: lea edx, var_30
  loc_00DEA40A: mov ecx, esi
  loc_00DEA40C: call [0040101Ch] ; __vbaVarMove
  loc_00DEA412: lea ecx, var_30
  loc_00DEA415: call edi
  loc_00DEA417: mov eax, var_70
  loc_00DEA41A: lea ecx, var_30
  loc_00DEA41D: push eax
  loc_00DEA41E: push esi
  loc_00DEA41F: push ecx
  loc_00DEA420: call [004011E0h] ; __vbaVarAdd
  loc_00DEA426: mov edx, eax
  loc_00DEA428: mov ecx, esi
  loc_00DEA42A: call [0040101Ch] ; __vbaVarMove
  loc_00DEA430: lea ecx, var_30
  loc_00DEA433: call edi
  loc_00DEA435: lea edx, var_30
  loc_00DEA438: push esi
  loc_00DEA439: push edx
  loc_00DEA43A: call [004011FCh] ; rtcVarStrFromVar
  loc_00DEA440: lea ecx, var_30
  loc_00DEA443: lea eax, [ebx+0000005Ch]
  loc_00DEA446: push ecx
  loc_00DEA447: mov var_74, eax
  loc_00DEA44A: call [00401034h] ; __vbaStrVarMove
  loc_00DEA450: mov edx, eax
  loc_00DEA452: lea ecx, var_18
  loc_00DEA455: call [00401228h] ; __vbaStrMove
  loc_00DEA45B: mov ecx, var_74
  loc_00DEA45E: mov edx, eax
  loc_00DEA460: call [004011B0h] ; __vbaStrCopy
  loc_00DEA466: lea ecx, var_18
  loc_00DEA469: call [00401258h] ; __vbaFreeStr
  loc_00DEA46F: lea ecx, var_30
  loc_00DEA472: call edi
  loc_00DEA474: lea edx, var_40
  loc_00DEA477: push esi
  loc_00DEA478: push edx
  loc_00DEA479: mov var_38, 00000064h
  loc_00DEA480: mov var_40, 00008002h
  loc_00DEA487: call [00401210h] ; __vbaVarTstGe
  loc_00DEA48D: test ax, ax
  loc_00DEA490: mov eax, [ebx]
  loc_00DEA492: push ebx
  loc_00DEA493: jz 00DEA92Fh
  loc_00DEA499: call [eax+00000334h]
  loc_00DEA49F: mov edi, [004010ACh] ; __vbaObjSet
  loc_00DEA4A5: lea ecx, var_1C
  loc_00DEA4A8: push eax
  loc_00DEA4A9: push ecx
  loc_00DEA4AA: call edi
  loc_00DEA4AC: mov esi, eax
  loc_00DEA4AE: push 4669AC00h
  loc_00DEA4B3: push esi
  loc_00DEA4B4: mov edx, [esi]
  loc_00DEA4B6: call [edx+0000007Ch]
  loc_00DEA4B9: test eax, eax
  loc_00DEA4BB: fnclex
  loc_00DEA4BD: jge 00DEA4CEh
  loc_00DEA4BF: push 0000007Ch
  loc_00DEA4C1: push 006DCDA0h
  loc_00DEA4C6: push esi
  loc_00DEA4C7: push eax
  loc_00DEA4C8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA4CE: lea ecx, var_1C
  loc_00DEA4D1: call [00401254h] ; __vbaFreeObj
  loc_00DEA4D7: mov eax, [00E3D024h]
  loc_00DEA4DC: test eax, eax
  loc_00DEA4DE: jnz 00DEA4F0h
  loc_00DEA4E0: push 00E3D024h
  loc_00DEA4E5: push 006CE120h
  loc_00DEA4EA: call [004011A0h] ; __vbaNew2
  loc_00DEA4F0: sub esp, 00000010h
  loc_00DEA4F3: mov eax, 0000000Ah
  loc_00DEA4F8: mov edx, esp
  loc_00DEA4FA: mov var_40, eax
  loc_00DEA4FD: sub esp, 00000010h
  loc_00DEA500: mov esi, [00E3D024h]
  loc_00DEA506: mov [edx], eax
  loc_00DEA508: mov eax, var_4C
  loc_00DEA50B: mov var_38, 80020004h
  loc_00DEA512: mov ecx, [esi]
  loc_00DEA514: mov [edx+00000004h], eax
  loc_00DEA517: mov eax, 80020004h
  loc_00DEA51C: mov [edx+00000008h], eax
  loc_00DEA51F: mov eax, var_44
  loc_00DEA522: mov [edx+0000000Ch], eax
  loc_00DEA525: mov eax, var_40
  loc_00DEA528: mov edx, esp
  loc_00DEA52A: push esi
  loc_00DEA52B: mov [edx], eax
  loc_00DEA52D: mov eax, var_3C
  loc_00DEA530: mov [edx+00000004h], eax
  loc_00DEA533: mov eax, var_38
  loc_00DEA536: mov [edx+00000008h], eax
  loc_00DEA539: mov eax, var_34
  loc_00DEA53C: mov [edx+0000000Ch], eax
  loc_00DEA53F: call [ecx+000002B0h]
  loc_00DEA545: test eax, eax
  loc_00DEA547: fnclex
  loc_00DEA549: jge 00DEA55Dh
  loc_00DEA54B: push 000002B0h
  loc_00DEA550: push 006DC970h
  loc_00DEA555: push esi
  loc_00DEA556: push eax
  loc_00DEA557: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA55D: mov eax, [00E3D024h]
  loc_00DEA562: mov var_38, FFFFFFFFh
  loc_00DEA569: test eax, eax
  loc_00DEA56B: mov var_40, 0000000Bh
  loc_00DEA572: jnz 00DEA589h
  loc_00DEA574: push 00E3D024h
  loc_00DEA579: push 006CE120h
  loc_00DEA57E: call [004011A0h] ; __vbaNew2
  loc_00DEA584: mov eax, [00E3D024h]
  loc_00DEA589: mov edx, var_40
  loc_00DEA58C: sub esp, 00000010h
  loc_00DEA58F: mov ecx, esp
  loc_00DEA591: push 80010007h
  loc_00DEA596: push eax
  loc_00DEA597: mov [ecx], edx
  loc_00DEA599: mov edx, var_3C
  loc_00DEA59C: mov [ecx+00000004h], edx
  loc_00DEA59F: mov edx, var_38
  loc_00DEA5A2: mov [ecx+00000008h], edx
  loc_00DEA5A5: mov edx, var_34
  loc_00DEA5A8: mov [ecx+0000000Ch], edx
  loc_00DEA5AB: mov ecx, [eax]
  loc_00DEA5AD: call [ecx+00000318h]
  loc_00DEA5B3: lea edx, var_1C
  loc_00DEA5B6: push eax
  loc_00DEA5B7: push edx
  loc_00DEA5B8: call edi
  loc_00DEA5BA: mov esi, [00401238h] ; __vbaLateIdSt
  loc_00DEA5C0: push eax
  loc_00DEA5C1: call __vbaLateIdSt
  loc_00DEA5C3: lea ecx, var_1C
  loc_00DEA5C6: call [00401254h] ; __vbaFreeObj
  loc_00DEA5CC: mov eax, [00E3D024h]
  loc_00DEA5D1: mov var_38, FFFFFFFFh
  loc_00DEA5D8: test eax, eax
  loc_00DEA5DA: mov var_40, 0000000Bh
  loc_00DEA5E1: jnz 00DEA5F8h
  loc_00DEA5E3: push 00E3D024h
  loc_00DEA5E8: push 006CE120h
  loc_00DEA5ED: call [004011A0h] ; __vbaNew2
  loc_00DEA5F3: mov eax, [00E3D024h]
  loc_00DEA5F8: mov edx, var_40
  loc_00DEA5FB: sub esp, 00000010h
  loc_00DEA5FE: mov ecx, esp
  loc_00DEA600: push 80010007h
  loc_00DEA605: push eax
  loc_00DEA606: mov [ecx], edx
  loc_00DEA608: mov edx, var_3C
  loc_00DEA60B: mov [ecx+00000004h], edx
  loc_00DEA60E: mov edx, var_38
  loc_00DEA611: mov [ecx+00000008h], edx
  loc_00DEA614: mov edx, var_34
  loc_00DEA617: mov [ecx+0000000Ch], edx
  loc_00DEA61A: mov ecx, [eax]
  loc_00DEA61C: call [ecx+0000031Ch]
  loc_00DEA622: lea edx, var_1C
  loc_00DEA625: push eax
  loc_00DEA626: push edx
  loc_00DEA627: call edi
  loc_00DEA629: push eax
  loc_00DEA62A: call __vbaLateIdSt
  loc_00DEA62C: lea ecx, var_1C
  loc_00DEA62F: call [00401254h] ; __vbaFreeObj
  loc_00DEA635: mov eax, [00E3D024h]
  loc_00DEA63A: mov var_38, FFFFFFFFh
  loc_00DEA641: test eax, eax
  loc_00DEA643: mov var_40, 0000000Bh
  loc_00DEA64A: jnz 00DEA661h
  loc_00DEA64C: push 00E3D024h
  loc_00DEA651: push 006CE120h
  loc_00DEA656: call [004011A0h] ; __vbaNew2
  loc_00DEA65C: mov eax, [00E3D024h]
  loc_00DEA661: mov edx, var_40
  loc_00DEA664: sub esp, 00000010h
  loc_00DEA667: mov ecx, esp
  loc_00DEA669: push 80010007h
  loc_00DEA66E: push eax
  loc_00DEA66F: mov [ecx], edx
  loc_00DEA671: mov edx, var_3C
  loc_00DEA674: mov [ecx+00000004h], edx
  loc_00DEA677: mov edx, var_38
  loc_00DEA67A: mov [ecx+00000008h], edx
  loc_00DEA67D: mov edx, var_34
  loc_00DEA680: mov [ecx+0000000Ch], edx
  loc_00DEA683: mov ecx, [eax]
  loc_00DEA685: call [ecx+00000320h]
  loc_00DEA68B: lea edx, var_1C
  loc_00DEA68E: push eax
  loc_00DEA68F: push edx
  loc_00DEA690: call edi
  loc_00DEA692: push eax
  loc_00DEA693: call __vbaLateIdSt
  loc_00DEA695: lea ecx, var_1C
  loc_00DEA698: call [00401254h] ; __vbaFreeObj
  loc_00DEA69E: mov eax, [00E3D024h]
  loc_00DEA6A3: mov var_38, 006DCDB4h ; " A D M I N "
  loc_00DEA6AA: test eax, eax
  loc_00DEA6AC: mov var_40, 00000008h
  loc_00DEA6B3: jnz 00DEA6CAh
  loc_00DEA6B5: push 00E3D024h
  loc_00DEA6BA: push 006CE120h
  loc_00DEA6BF: call [004011A0h] ; __vbaNew2
  loc_00DEA6C5: mov eax, [00E3D024h]
  loc_00DEA6CA: mov edx, var_40
  loc_00DEA6CD: sub esp, 00000010h
  loc_00DEA6D0: mov ecx, esp
  loc_00DEA6D2: push FFFFFDFAh
  loc_00DEA6D7: push eax
  loc_00DEA6D8: mov [ecx], edx
  loc_00DEA6DA: mov edx, var_3C
  loc_00DEA6DD: mov [ecx+00000004h], edx
  loc_00DEA6E0: mov edx, var_38
  loc_00DEA6E3: mov [ecx+00000008h], edx
  loc_00DEA6E6: mov edx, var_34
  loc_00DEA6E9: mov [ecx+0000000Ch], edx
  loc_00DEA6EC: mov ecx, [eax]
  loc_00DEA6EE: call [ecx+00000324h]
  loc_00DEA6F4: lea edx, var_1C
  loc_00DEA6F7: push eax
  loc_00DEA6F8: push edx
  loc_00DEA6F9: call edi
  loc_00DEA6FB: push eax
  loc_00DEA6FC: call __vbaLateIdSt
  loc_00DEA6FE: lea ecx, var_1C
  loc_00DEA701: call [00401254h] ; __vbaFreeObj
  loc_00DEA707: mov eax, [00E3D024h]
  loc_00DEA70C: mov var_38, 00000000h
  loc_00DEA713: test eax, eax
  loc_00DEA715: mov var_40, 0000000Bh
  loc_00DEA71C: jnz 00DEA733h
  loc_00DEA71E: push 00E3D024h
  loc_00DEA723: push 006CE120h
  loc_00DEA728: call [004011A0h] ; __vbaNew2
  loc_00DEA72E: mov eax, [00E3D024h]
  loc_00DEA733: mov edx, var_40
  loc_00DEA736: sub esp, 00000010h
  loc_00DEA739: mov ecx, esp
  loc_00DEA73B: push 8001000Dh
  loc_00DEA740: push eax
  loc_00DEA741: mov [ecx], edx
  loc_00DEA743: mov edx, var_3C
  loc_00DEA746: mov [ecx+00000004h], edx
  loc_00DEA749: mov edx, var_38
  loc_00DEA74C: mov [ecx+00000008h], edx
  loc_00DEA74F: mov edx, var_34
  loc_00DEA752: mov [ecx+0000000Ch], edx
  loc_00DEA755: mov ecx, [eax]
  loc_00DEA757: call [ecx+00000324h]
  loc_00DEA75D: lea edx, var_1C
  loc_00DEA760: push eax
  loc_00DEA761: push edx
  loc_00DEA762: call edi
  loc_00DEA764: push eax
  loc_00DEA765: call __vbaLateIdSt
  loc_00DEA767: lea ecx, var_1C
  loc_00DEA76A: call [00401254h] ; __vbaFreeObj
  loc_00DEA770: mov eax, [00E3D024h]
  loc_00DEA775: mov var_38, 00404040h
  loc_00DEA77C: test eax, eax
  loc_00DEA77E: mov var_40, 00000003h
  loc_00DEA785: jnz 00DEA79Ch
  loc_00DEA787: push 00E3D024h
  loc_00DEA78C: push 006CE120h
  loc_00DEA791: call [004011A0h] ; __vbaNew2
  loc_00DEA797: mov eax, [00E3D024h]
  loc_00DEA79C: mov edx, var_40
  loc_00DEA79F: sub esp, 00000010h
  loc_00DEA7A2: mov ecx, esp
  loc_00DEA7A4: push FFFFFE0Bh
  loc_00DEA7A9: push eax
  loc_00DEA7AA: mov [ecx], edx
  loc_00DEA7AC: mov edx, var_3C
  loc_00DEA7AF: mov [ecx+00000004h], edx
  loc_00DEA7B2: mov edx, var_38
  loc_00DEA7B5: mov [ecx+00000008h], edx
  loc_00DEA7B8: mov edx, var_34
  loc_00DEA7BB: mov [ecx+0000000Ch], edx
  loc_00DEA7BE: mov ecx, [eax]
  loc_00DEA7C0: call [ecx+00000324h]
  loc_00DEA7C6: lea edx, var_1C
  loc_00DEA7C9: push eax
  loc_00DEA7CA: push edx
  loc_00DEA7CB: call edi
  loc_00DEA7CD: push eax
  loc_00DEA7CE: call __vbaLateIdSt
  loc_00DEA7D0: lea ecx, var_1C
  loc_00DEA7D3: call [00401254h] ; __vbaFreeObj
  loc_00DEA7D9: mov eax, [00E3D024h]
  loc_00DEA7DE: mov var_38, 00E0E0E0h
  loc_00DEA7E5: test eax, eax
  loc_00DEA7E7: mov var_40, 00000003h
  loc_00DEA7EE: jnz 00DEA805h
  loc_00DEA7F0: push 00E3D024h
  loc_00DEA7F5: push 006CE120h
  loc_00DEA7FA: call [004011A0h] ; __vbaNew2
  loc_00DEA800: mov eax, [00E3D024h]
  loc_00DEA805: mov edx, var_40
  loc_00DEA808: sub esp, 00000010h
  loc_00DEA80B: mov ecx, esp
  loc_00DEA80D: push FFFFFDFFh
  loc_00DEA812: push eax
  loc_00DEA813: mov [ecx], edx
  loc_00DEA815: mov edx, var_3C
  loc_00DEA818: mov [ecx+00000004h], edx
  loc_00DEA81B: mov edx, var_38
  loc_00DEA81E: mov [ecx+00000008h], edx
  loc_00DEA821: mov edx, var_34
  loc_00DEA824: mov [ecx+0000000Ch], edx
  loc_00DEA827: mov ecx, [eax]
  loc_00DEA829: call [ecx+00000324h]
  loc_00DEA82F: lea edx, var_1C
  loc_00DEA832: push eax
  loc_00DEA833: push edx
  loc_00DEA834: call edi
  loc_00DEA836: push eax
  loc_00DEA837: call __vbaLateIdSt
  loc_00DEA839: lea ecx, var_1C
  loc_00DEA83C: call [00401254h] ; __vbaFreeObj
  loc_00DEA842: mov eax, [ebx]
  loc_00DEA844: push ebx
  loc_00DEA845: call [eax+00000334h]
  loc_00DEA84B: lea ecx, var_1C
  loc_00DEA84E: push eax
  loc_00DEA84F: push ecx
  loc_00DEA850: call edi
  loc_00DEA852: mov esi, eax
  loc_00DEA854: lea eax, var_54
  loc_00DEA857: push eax
  loc_00DEA858: push esi
  loc_00DEA859: mov edx, [esi]
  loc_00DEA85B: call [edx+00000078h]
  loc_00DEA85E: test eax, eax
  loc_00DEA860: fnclex
  loc_00DEA862: jge 00DEA873h
  loc_00DEA864: push 00000078h
  loc_00DEA866: push 006DCDA0h
  loc_00DEA86B: push esi
  loc_00DEA86C: push eax
  loc_00DEA86D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA873: cmp var_54, 4669AC00h
  loc_00DEA87A: jnz 00DEA883h
  loc_00DEA87C: mov eax, 00000001h
  loc_00DEA881: jmp 00DEA885h
  loc_00DEA883: xor eax, eax
  loc_00DEA885: neg eax
  loc_00DEA887: lea ecx, var_1C
  loc_00DEA88A: mov si, ax
  loc_00DEA88D: call [00401254h] ; __vbaFreeObj
  loc_00DEA893: test si, si
  loc_00DEA896: jz 00DEA9E2h
  loc_00DEA89C: mov ecx, [ebx]
  loc_00DEA89E: push ebx
  loc_00DEA89F: call [ecx+00000334h]
  loc_00DEA8A5: lea edx, var_1C
  loc_00DEA8A8: push eax
  loc_00DEA8A9: push edx
  loc_00DEA8AA: call edi
  loc_00DEA8AC: mov edi, eax
  loc_00DEA8AE: push 3F800000h
  loc_00DEA8B3: push edi
  loc_00DEA8B4: mov eax, [edi]
  loc_00DEA8B6: call [eax+0000007Ch]
  loc_00DEA8B9: test eax, eax
  loc_00DEA8BB: fnclex
  loc_00DEA8BD: jge 00DEA8CEh
  loc_00DEA8BF: push 0000007Ch
  loc_00DEA8C1: push 006DCDA0h
  loc_00DEA8C6: push edi
  loc_00DEA8C7: push eax
  loc_00DEA8C8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA8CE: mov esi, [00401254h] ; __vbaFreeObj
  loc_00DEA8D4: lea ecx, var_1C
  loc_00DEA8D7: call __vbaFreeObj
  loc_00DEA8D9: mov eax, [00E3D9CCh]
  loc_00DEA8DE: test eax, eax
  loc_00DEA8E0: jnz 00DEA8F2h
  loc_00DEA8E2: push 00E3D9CCh
  loc_00DEA8E7: push 006DC960h
  loc_00DEA8EC: call [004011A0h] ; __vbaNew2
  loc_00DEA8F2: mov edi, [00E3D9CCh]
  loc_00DEA8F8: lea ecx, var_1C
  loc_00DEA8FB: push ebx
  loc_00DEA8FC: push ecx
  loc_00DEA8FD: mov edx, [edi]
  loc_00DEA8FF: mov var_78, edx
  loc_00DEA902: call [004010B4h] ; __vbaObjSetAddref
  loc_00DEA908: mov edx, var_78
  loc_00DEA90B: push eax
  loc_00DEA90C: push edi
  loc_00DEA90D: call [edx+00000010h]
  loc_00DEA910: test eax, eax
  loc_00DEA912: fnclex
  loc_00DEA914: jge 00DEA925h
  loc_00DEA916: push 00000010h
  loc_00DEA918: push 006DC950h
  loc_00DEA91D: push edi
  loc_00DEA91E: push eax
  loc_00DEA91F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA925: lea ecx, var_1C
  loc_00DEA928: call __vbaFreeObj
  loc_00DEA92A: jmp 00DEA9E2h
  loc_00DEA92F: call [eax+00000334h]
  loc_00DEA935: mov esi, [004010ACh] ; __vbaObjSet
  loc_00DEA93B: lea ecx, var_20
  loc_00DEA93E: push eax
  loc_00DEA93F: push ecx
  loc_00DEA940: call __vbaObjSet
  loc_00DEA942: mov edx, [ebx]
  loc_00DEA944: push ebx
  loc_00DEA945: mov edi, eax
  loc_00DEA947: call [edx+00000334h]
  loc_00DEA94D: push eax
  loc_00DEA94E: lea eax, var_1C
  loc_00DEA951: push eax
  loc_00DEA952: call __vbaObjSet
  loc_00DEA954: mov esi, eax
  loc_00DEA956: lea edx, var_54
  loc_00DEA959: push edx
  loc_00DEA95A: push esi
  loc_00DEA95B: mov ecx, [esi]
  loc_00DEA95D: call [ecx+00000078h]
  loc_00DEA960: test eax, eax
  loc_00DEA962: fnclex
  loc_00DEA964: jge 00DEA975h
  loc_00DEA966: push 00000078h
  loc_00DEA968: push 006DCDA0h
  loc_00DEA96D: push esi
  loc_00DEA96E: push eax
  loc_00DEA96F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA975: fld real4 ptr var_54
  loc_00DEA978: fadd st0, real4 ptr [004012D0h]
  loc_00DEA97E: mov ecx, [edi]
  loc_00DEA980: push ecx
  loc_00DEA981: fnstsw ax
  loc_00DEA983: test al, 0Dh
  loc_00DEA985: jnz 00DEAA35h
  loc_00DEA98B: fstp real4 ptr [esp]
  loc_00DEA98E: push edi
  loc_00DEA98F: call [ecx+0000007Ch]
  loc_00DEA992: test eax, eax
  loc_00DEA994: fnclex
  loc_00DEA996: jge 00DEA9A7h
  loc_00DEA998: push 0000007Ch
  loc_00DEA99A: push 006DCDA0h
  loc_00DEA99F: push edi
  loc_00DEA9A0: push eax
  loc_00DEA9A1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEA9A7: lea edx, var_20
  loc_00DEA9AA: lea eax, var_1C
  loc_00DEA9AD: push edx
  loc_00DEA9AE: push eax
  loc_00DEA9AF: push 00000002h
  loc_00DEA9B1: call [00401048h] ; __vbaFreeObjList
  loc_00DEA9B7: mov ecx, var_74
  loc_00DEA9BA: add esp, 0000000Ch
  loc_00DEA9BD: mov edx, [ecx]
  loc_00DEA9BF: push edx
  loc_00DEA9C0: call [004011A4h] ; __vbaR8Str
  loc_00DEA9C6: call [0040124Ch] ; __vbaFPInt
  loc_00DEA9CC: mov ecx, var_70
  loc_00DEA9CF: lea edx, var_40
  loc_00DEA9D2: fstp real8 ptr var_38
  loc_00DEA9D5: mov var_40, 00000005h
  loc_00DEA9DC: call [0040101Ch] ; __vbaVarMove
  loc_00DEA9E2: xor edi, edi
  loc_00DEA9E4: mov var_4, edi
  loc_00DEA9E7: fwait
  loc_00DEA9E8: push 00DEAA16h
  loc_00DEA9ED: jmp 00DEAA15h
  loc_00DEA9EF: lea ecx, var_18
  loc_00DEA9F2: call [00401258h] ; __vbaFreeStr
  loc_00DEA9F8: lea eax, var_20
  loc_00DEA9FB: lea ecx, var_1C
  loc_00DEA9FE: push eax
  loc_00DEA9FF: push ecx
  loc_00DEAA00: push 00000002h
  loc_00DEAA02: call [00401048h] ; __vbaFreeObjList
  loc_00DEAA08: add esp, 0000000Ch
  loc_00DEAA0B: lea ecx, var_30
  loc_00DEAA0E: call [00401024h] ; __vbaFreeVar
  loc_00DEAA14: ret
  loc_00DEAA15: ret
  loc_00DEAA16: mov eax, Me
  loc_00DEAA19: push eax
  loc_00DEAA1A: mov edx, [eax]
  loc_00DEAA1C: call [edx+00000008h]
  loc_00DEAA1F: mov eax, var_4
  loc_00DEAA22: mov ecx, var_14
  loc_00DEAA25: pop edi
  loc_00DEAA26: pop esi
  loc_00DEAA27: mov fs:[00000000h], ecx
  loc_00DEAA2E: pop ebx
  loc_00DEAA2F: mov esp, ebp
  loc_00DEAA31: pop ebp
  loc_00DEAA32: retn 0004h
End Sub

Private Sub Timer2_Timer() 'DEAA40
  loc_00DEAA40: push ebp
  loc_00DEAA41: mov ebp, esp
  loc_00DEAA43: sub esp, 0000000Ch
  loc_00DEAA46: push 00402836h ; __vbaExceptHandler
  loc_00DEAA4B: mov eax, fs:[00000000h]
  loc_00DEAA51: push eax
  loc_00DEAA52: mov fs:[00000000h], esp
  loc_00DEAA59: sub esp, 00000050h
  loc_00DEAA5C: push ebx
  loc_00DEAA5D: push esi
  loc_00DEAA5E: push edi
  loc_00DEAA5F: mov var_C, esp
  loc_00DEAA62: mov var_8, 004012F8h
  loc_00DEAA69: mov esi, Me
  loc_00DEAA6C: mov eax, esi
  loc_00DEAA6E: and eax, 00000001h
  loc_00DEAA71: mov var_4, eax
  loc_00DEAA74: and esi, FFFFFFFEh
  loc_00DEAA77: push esi
  loc_00DEAA78: mov Me, esi
  loc_00DEAA7B: mov ecx, [esi]
  loc_00DEAA7D: call [ecx+00000004h]
  loc_00DEAA80: xor ebx, ebx
  loc_00DEAA82: lea edi, [esi+00000034h]
  loc_00DEAA85: mov var_40, ebx
  loc_00DEAA88: lea edx, var_40
  loc_00DEAA8B: mov ecx, edi
  loc_00DEAA8D: mov var_18, ebx
  loc_00DEAA90: mov var_1C, ebx
  loc_00DEAA93: mov var_20, ebx
  loc_00DEAA96: mov var_30, ebx
  loc_00DEAA99: mov var_44, ebx
  loc_00DEAA9C: mov var_38, ebx
  loc_00DEAA9F: mov var_40, 00000002h
  loc_00DEAAA6: call [0040101Ch] ; __vbaVarMove
  loc_00DEAAAC: lea eax, [esi+00000044h]
  loc_00DEAAAF: lea edx, var_40
  loc_00DEAAB2: push eax
  loc_00DEAAB3: push edx
  loc_00DEAAB4: mov var_38, 00000064h
  loc_00DEAABB: mov var_40, 00008002h
  loc_00DEAAC2: mov var_60, eax
  loc_00DEAAC5: call [00401088h] ; __vbaVarTstLe
  loc_00DEAACB: test ax, ax
  loc_00DEAACE: jz 00DEADDCh
  loc_00DEAAD4: lea eax, var_30
  loc_00DEAAD7: mov var_28, 80020004h
  loc_00DEAADE: push eax
  loc_00DEAADF: mov var_30, 0000000Ah
  loc_00DEAAE6: call [004010A0h] ; rtcRandomize
  loc_00DEAAEC: mov ebx, [00401024h] ; __vbaFreeVar
  loc_00DEAAF2: lea ecx, var_30
  loc_00DEAAF5: call ebx
  loc_00DEAAF7: lea ecx, var_30
  loc_00DEAAFA: mov var_28, 80020004h
  loc_00DEAB01: push ecx
  loc_00DEAB02: mov var_30, 0000000Ah
  loc_00DEAB09: call [0040109Ch] ; rtcRandomNext
  loc_00DEAB0F: fstp real4 ptr var_44
  loc_00DEAB12: fld real4 ptr var_44
  loc_00DEAB15: fmul st0, real4 ptr [004012E4h]
  loc_00DEAB1B: lea ecx, var_30
  loc_00DEAB1E: fadd st0, real4 ptr [004012E0h]
  loc_00DEAB24: fstp real8 ptr [esi+00000054h]
  loc_00DEAB27: fnstsw ax
  loc_00DEAB29: test al, 0Dh
  loc_00DEAB2B: jnz 00DEAE2Dh
  loc_00DEAB31: call ebx
  loc_00DEAB33: fld real8 ptr [esi+00000054h]
  loc_00DEAB36: cmp [00E3D000h], 00000000h
  loc_00DEAB3D: jnz 00DEAB47h
  loc_00DEAB3F: fdiv st0, real8 ptr [004012D8h]
  loc_00DEAB45: jmp 00DEAB58h
  loc_00DEAB47: push [004012DCh]
  loc_00DEAB4D: push [004012D8h]
  loc_00DEAB53: call 00402854h ; _adj_fdiv_m64
  loc_00DEAB58: lea edx, var_40
  loc_00DEAB5B: mov ecx, edi
  loc_00DEAB5D: mov var_40, 00000005h
  loc_00DEAB64: fstp real8 ptr var_38
  loc_00DEAB67: fnstsw ax
  loc_00DEAB69: test al, 0Dh
  loc_00DEAB6B: jnz 00DEAE2Dh
  loc_00DEAB71: call [0040101Ch] ; __vbaVarMove
  loc_00DEAB77: push 00000000h
  loc_00DEAB79: lea edx, var_30
  loc_00DEAB7C: push edi
  loc_00DEAB7D: push edx
  loc_00DEAB7E: call [00401170h] ; rtcRound
  loc_00DEAB84: lea edx, var_30
  loc_00DEAB87: mov ecx, edi
  loc_00DEAB89: call [0040101Ch] ; __vbaVarMove
  loc_00DEAB8F: lea ecx, var_30
  loc_00DEAB92: call ebx
  loc_00DEAB94: mov eax, var_60
  loc_00DEAB97: lea ecx, var_30
  loc_00DEAB9A: push eax
  loc_00DEAB9B: push edi
  loc_00DEAB9C: push ecx
  loc_00DEAB9D: call [004011E0h] ; __vbaVarAdd
  loc_00DEABA3: mov edx, eax
  loc_00DEABA5: mov ecx, edi
  loc_00DEABA7: call [0040101Ch] ; __vbaVarMove
  loc_00DEABAD: lea ecx, var_30
  loc_00DEABB0: call ebx
  loc_00DEABB2: lea edx, var_30
  loc_00DEABB5: push edi
  loc_00DEABB6: push edx
  loc_00DEABB7: call [004011FCh] ; rtcVarStrFromVar
  loc_00DEABBD: lea ecx, var_30
  loc_00DEABC0: lea eax, [esi+0000005Ch]
  loc_00DEABC3: push ecx
  loc_00DEABC4: mov var_64, eax
  loc_00DEABC7: call [00401034h] ; __vbaStrVarMove
  loc_00DEABCD: mov edx, eax
  loc_00DEABCF: lea ecx, var_18
  loc_00DEABD2: call [00401228h] ; __vbaStrMove
  loc_00DEABD8: mov ecx, var_64
  loc_00DEABDB: mov edx, eax
  loc_00DEABDD: call [004011B0h] ; __vbaStrCopy
  loc_00DEABE3: lea ecx, var_18
  loc_00DEABE6: call [00401258h] ; __vbaFreeStr
  loc_00DEABEC: lea ecx, var_30
  loc_00DEABEF: call ebx
  loc_00DEABF1: lea edx, var_40
  loc_00DEABF4: push edi
  loc_00DEABF5: push edx
  loc_00DEABF6: mov var_38, 00000064h
  loc_00DEABFD: mov var_40, 00008002h
  loc_00DEAC04: call [00401210h] ; __vbaVarTstGe
  loc_00DEAC0A: test ax, ax
  loc_00DEAC0D: jz 00DEAD24h
  loc_00DEAC13: mov eax, [esi]
  loc_00DEAC15: push esi
  loc_00DEAC16: call [eax+00000334h]
  loc_00DEAC1C: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00DEAC22: lea ecx, var_1C
  loc_00DEAC25: push eax
  loc_00DEAC26: push ecx
  loc_00DEAC27: call ebx
  loc_00DEAC29: mov edi, eax
  loc_00DEAC2B: push 4669AC00h
  loc_00DEAC30: push edi
  loc_00DEAC31: mov edx, [edi]
  loc_00DEAC33: call [edx+0000007Ch]
  loc_00DEAC36: test eax, eax
  loc_00DEAC38: fnclex
  loc_00DEAC3A: jge 00DEAC4Bh
  loc_00DEAC3C: push 0000007Ch
  loc_00DEAC3E: push 006DCDA0h
  loc_00DEAC43: push edi
  loc_00DEAC44: push eax
  loc_00DEAC45: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEAC4B: lea ecx, var_1C
  loc_00DEAC4E: call [00401254h] ; __vbaFreeObj
  loc_00DEAC54: mov eax, [esi]
  loc_00DEAC56: push esi
  loc_00DEAC57: call [eax+00000334h]
  loc_00DEAC5D: lea ecx, var_1C
  loc_00DEAC60: push eax
  loc_00DEAC61: push ecx
  loc_00DEAC62: call ebx
  loc_00DEAC64: mov edi, eax
  loc_00DEAC66: lea eax, var_44
  loc_00DEAC69: push eax
  loc_00DEAC6A: push edi
  loc_00DEAC6B: mov edx, [edi]
  loc_00DEAC6D: call [edx+00000078h]
  loc_00DEAC70: test eax, eax
  loc_00DEAC72: fnclex
  loc_00DEAC74: jge 00DEAC85h
  loc_00DEAC76: push 00000078h
  loc_00DEAC78: push 006DCDA0h
  loc_00DEAC7D: push edi
  loc_00DEAC7E: push eax
  loc_00DEAC7F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEAC85: cmp var_44, 4669AC00h
  loc_00DEAC8C: jnz 00DEAC95h
  loc_00DEAC8E: mov eax, 00000001h
  loc_00DEAC93: jmp 00DEAC97h
  loc_00DEAC95: xor eax, eax
  loc_00DEAC97: neg eax
  loc_00DEAC99: lea ecx, var_1C
  loc_00DEAC9C: mov di, ax
  loc_00DEAC9F: call [00401254h] ; __vbaFreeObj
  loc_00DEACA5: test di, di
  loc_00DEACA8: jz 00DEADDAh
  loc_00DEACAE: mov ecx, [esi]
  loc_00DEACB0: push esi
  loc_00DEACB1: call [ecx+00000334h]
  loc_00DEACB7: lea edx, var_1C
  loc_00DEACBA: push eax
  loc_00DEACBB: push edx
  loc_00DEACBC: call ebx
  loc_00DEACBE: mov edi, eax
  loc_00DEACC0: push 3F800000h
  loc_00DEACC5: push edi
  loc_00DEACC6: mov eax, [edi]
  loc_00DEACC8: call [eax+0000007Ch]
  loc_00DEACCB: test eax, eax
  loc_00DEACCD: fnclex
  loc_00DEACCF: jge 00DEACE0h
  loc_00DEACD1: push 0000007Ch
  loc_00DEACD3: push 006DCDA0h
  loc_00DEACD8: push edi
  loc_00DEACD9: push eax
  loc_00DEACDA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEACE0: mov edi, [00401254h] ; __vbaFreeObj
  loc_00DEACE6: lea ecx, var_1C
  loc_00DEACE9: call edi
  loc_00DEACEB: mov ecx, [esi]
  loc_00DEACED: push esi
  loc_00DEACEE: call [ecx+00000300h]
  loc_00DEACF4: lea edx, var_1C
  loc_00DEACF7: push eax
  loc_00DEACF8: push edx
  loc_00DEACF9: call ebx
  loc_00DEACFB: mov esi, eax
  loc_00DEACFD: push 00000000h
  loc_00DEACFF: push esi
  loc_00DEAD00: mov eax, [esi]
  loc_00DEAD02: call [eax+0000005Ch]
  loc_00DEAD05: test eax, eax
  loc_00DEAD07: fnclex
  loc_00DEAD09: jge 00DEAD1Ah
  loc_00DEAD0B: push 0000005Ch
  loc_00DEAD0D: push 006DCAE0h
  loc_00DEAD12: push esi
  loc_00DEAD13: push eax
  loc_00DEAD14: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEAD1A: lea ecx, var_1C
  loc_00DEAD1D: call edi
  loc_00DEAD1F: jmp 00DEADDAh
  loc_00DEAD24: mov ecx, [esi]
  loc_00DEAD26: push esi
  loc_00DEAD27: call [ecx+00000334h]
  loc_00DEAD2D: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00DEAD33: lea edx, var_20
  loc_00DEAD36: push eax
  loc_00DEAD37: push edx
  loc_00DEAD38: call ebx
  loc_00DEAD3A: mov edi, eax
  loc_00DEAD3C: mov eax, [esi]
  loc_00DEAD3E: push esi
  loc_00DEAD3F: call [eax+00000334h]
  loc_00DEAD45: lea ecx, var_1C
  loc_00DEAD48: push eax
  loc_00DEAD49: push ecx
  loc_00DEAD4A: call ebx
  loc_00DEAD4C: mov esi, eax
  loc_00DEAD4E: lea eax, var_44
  loc_00DEAD51: push eax
  loc_00DEAD52: push esi
  loc_00DEAD53: mov edx, [esi]
  loc_00DEAD55: call [edx+00000078h]
  loc_00DEAD58: test eax, eax
  loc_00DEAD5A: fnclex
  loc_00DEAD5C: jge 00DEAD6Dh
  loc_00DEAD5E: push 00000078h
  loc_00DEAD60: push 006DCDA0h
  loc_00DEAD65: push esi
  loc_00DEAD66: push eax
  loc_00DEAD67: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEAD6D: fld real4 ptr var_44
  loc_00DEAD70: fadd st0, real4 ptr [004012D0h]
  loc_00DEAD76: mov ecx, [edi]
  loc_00DEAD78: push ecx
  loc_00DEAD79: fnstsw ax
  loc_00DEAD7B: test al, 0Dh
  loc_00DEAD7D: jnz 00DEAE2Dh
  loc_00DEAD83: fstp real4 ptr [esp]
  loc_00DEAD86: push edi
  loc_00DEAD87: call [ecx+0000007Ch]
  loc_00DEAD8A: test eax, eax
  loc_00DEAD8C: fnclex
  loc_00DEAD8E: jge 00DEAD9Fh
  loc_00DEAD90: push 0000007Ch
  loc_00DEAD92: push 006DCDA0h
  loc_00DEAD97: push edi
  loc_00DEAD98: push eax
  loc_00DEAD99: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEAD9F: lea edx, var_20
  loc_00DEADA2: lea eax, var_1C
  loc_00DEADA5: push edx
  loc_00DEADA6: push eax
  loc_00DEADA7: push 00000002h
  loc_00DEADA9: call [00401048h] ; __vbaFreeObjList
  loc_00DEADAF: mov ecx, var_64
  loc_00DEADB2: add esp, 0000000Ch
  loc_00DEADB5: mov edx, [ecx]
  loc_00DEADB7: push edx
  loc_00DEADB8: call [004011A4h] ; __vbaR8Str
  loc_00DEADBE: call [0040124Ch] ; __vbaFPInt
  loc_00DEADC4: mov ecx, var_60
  loc_00DEADC7: lea edx, var_40
  loc_00DEADCA: fstp real8 ptr var_38
  loc_00DEADCD: mov var_40, 00000005h
  loc_00DEADD4: call [0040101Ch] ; __vbaVarMove
  loc_00DEADDA: xor ebx, ebx
  loc_00DEADDC: mov var_4, ebx
  loc_00DEADDF: fwait
  loc_00DEADE0: push 00DEAE0Eh
  loc_00DEADE5: jmp 00DEAE0Dh
  loc_00DEADE7: lea ecx, var_18
  loc_00DEADEA: call [00401258h] ; __vbaFreeStr
  loc_00DEADF0: lea eax, var_20
  loc_00DEADF3: lea ecx, var_1C
  loc_00DEADF6: push eax
  loc_00DEADF7: push ecx
  loc_00DEADF8: push 00000002h
  loc_00DEADFA: call [00401048h] ; __vbaFreeObjList
  loc_00DEAE00: add esp, 0000000Ch
  loc_00DEAE03: lea ecx, var_30
  loc_00DEAE06: call [00401024h] ; __vbaFreeVar
  loc_00DEAE0C: ret
  loc_00DEAE0D: ret
  loc_00DEAE0E: mov eax, Me
  loc_00DEAE11: push eax
  loc_00DEAE12: mov edx, [eax]
  loc_00DEAE14: call [edx+00000008h]
  loc_00DEAE17: mov eax, var_4
  loc_00DEAE1A: mov ecx, var_14
  loc_00DEAE1D: pop edi
  loc_00DEAE1E: pop esi
  loc_00DEAE1F: mov fs:[00000000h], ecx
  loc_00DEAE26: pop ebx
  loc_00DEAE27: mov esp, ebp
  loc_00DEAE29: pop ebp
  loc_00DEAE2A: retn 0004h
End Sub

Private Sub Timer3_Timer() 'DEAE40
  loc_00DEAE40: push ebp
  loc_00DEAE41: mov ebp, esp
  loc_00DEAE43: sub esp, 0000000Ch
  loc_00DEAE46: push 00402836h ; __vbaExceptHandler
  loc_00DEAE4B: mov eax, fs:[00000000h]
  loc_00DEAE51: push eax
  loc_00DEAE52: mov fs:[00000000h], esp
  loc_00DEAE59: sub esp, 00000050h
  loc_00DEAE5C: push ebx
  loc_00DEAE5D: push esi
  loc_00DEAE5E: push edi
  loc_00DEAE5F: mov var_C, esp
  loc_00DEAE62: mov var_8, 00401308h
  loc_00DEAE69: mov esi, Me
  loc_00DEAE6C: mov eax, esi
  loc_00DEAE6E: and eax, 00000001h
  loc_00DEAE71: mov var_4, eax
  loc_00DEAE74: and esi, FFFFFFFEh
  loc_00DEAE77: push esi
  loc_00DEAE78: mov Me, esi
  loc_00DEAE7B: mov ecx, [esi]
  loc_00DEAE7D: call [ecx+00000004h]
  loc_00DEAE80: xor ebx, ebx
  loc_00DEAE82: lea edi, [esi+00000034h]
  loc_00DEAE85: mov var_40, ebx
  loc_00DEAE88: lea edx, var_40
  loc_00DEAE8B: mov ecx, edi
  loc_00DEAE8D: mov var_18, ebx
  loc_00DEAE90: mov var_1C, ebx
  loc_00DEAE93: mov var_20, ebx
  loc_00DEAE96: mov var_30, ebx
  loc_00DEAE99: mov var_44, ebx
  loc_00DEAE9C: mov var_38, ebx
  loc_00DEAE9F: mov var_40, 00000002h
  loc_00DEAEA6: call [0040101Ch] ; __vbaVarMove
  loc_00DEAEAC: lea eax, [esi+00000044h]
  loc_00DEAEAF: lea edx, var_40
  loc_00DEAEB2: push eax
  loc_00DEAEB3: push edx
  loc_00DEAEB4: mov var_38, 00000064h
  loc_00DEAEBB: mov var_40, 00008002h
  loc_00DEAEC2: mov var_60, eax
  loc_00DEAEC5: call [00401088h] ; __vbaVarTstLe
  loc_00DEAECB: test ax, ax
  loc_00DEAECE: jz 00DEB1DCh
  loc_00DEAED4: lea eax, var_30
  loc_00DEAED7: mov var_28, 80020004h
  loc_00DEAEDE: push eax
  loc_00DEAEDF: mov var_30, 0000000Ah
  loc_00DEAEE6: call [004010A0h] ; rtcRandomize
  loc_00DEAEEC: mov ebx, [00401024h] ; __vbaFreeVar
  loc_00DEAEF2: lea ecx, var_30
  loc_00DEAEF5: call ebx
  loc_00DEAEF7: lea ecx, var_30
  loc_00DEAEFA: mov var_28, 80020004h
  loc_00DEAF01: push ecx
  loc_00DEAF02: mov var_30, 0000000Ah
  loc_00DEAF09: call [0040109Ch] ; rtcRandomNext
  loc_00DEAF0F: fstp real4 ptr var_44
  loc_00DEAF12: fld real4 ptr var_44
  loc_00DEAF15: fmul st0, real4 ptr [004012E4h]
  loc_00DEAF1B: lea ecx, var_30
  loc_00DEAF1E: fadd st0, real4 ptr [004012E0h]
  loc_00DEAF24: fstp real8 ptr [esi+00000054h]
  loc_00DEAF27: fnstsw ax
  loc_00DEAF29: test al, 0Dh
  loc_00DEAF2B: jnz 00DEB22Dh
  loc_00DEAF31: call ebx
  loc_00DEAF33: fld real8 ptr [esi+00000054h]
  loc_00DEAF36: cmp [00E3D000h], 00000000h
  loc_00DEAF3D: jnz 00DEAF47h
  loc_00DEAF3F: fdiv st0, real8 ptr [004012D8h]
  loc_00DEAF45: jmp 00DEAF58h
  loc_00DEAF47: push [004012DCh]
  loc_00DEAF4D: push [004012D8h]
  loc_00DEAF53: call 00402854h ; _adj_fdiv_m64
  loc_00DEAF58: lea edx, var_40
  loc_00DEAF5B: mov ecx, edi
  loc_00DEAF5D: mov var_40, 00000005h
  loc_00DEAF64: fstp real8 ptr var_38
  loc_00DEAF67: fnstsw ax
  loc_00DEAF69: test al, 0Dh
  loc_00DEAF6B: jnz 00DEB22Dh
  loc_00DEAF71: call [0040101Ch] ; __vbaVarMove
  loc_00DEAF77: push 00000000h
  loc_00DEAF79: lea edx, var_30
  loc_00DEAF7C: push edi
  loc_00DEAF7D: push edx
  loc_00DEAF7E: call [00401170h] ; rtcRound
  loc_00DEAF84: lea edx, var_30
  loc_00DEAF87: mov ecx, edi
  loc_00DEAF89: call [0040101Ch] ; __vbaVarMove
  loc_00DEAF8F: lea ecx, var_30
  loc_00DEAF92: call ebx
  loc_00DEAF94: mov eax, var_60
  loc_00DEAF97: lea ecx, var_30
  loc_00DEAF9A: push eax
  loc_00DEAF9B: push edi
  loc_00DEAF9C: push ecx
  loc_00DEAF9D: call [004011E0h] ; __vbaVarAdd
  loc_00DEAFA3: mov edx, eax
  loc_00DEAFA5: mov ecx, edi
  loc_00DEAFA7: call [0040101Ch] ; __vbaVarMove
  loc_00DEAFAD: lea ecx, var_30
  loc_00DEAFB0: call ebx
  loc_00DEAFB2: lea edx, var_30
  loc_00DEAFB5: push edi
  loc_00DEAFB6: push edx
  loc_00DEAFB7: call [004011FCh] ; rtcVarStrFromVar
  loc_00DEAFBD: lea ecx, var_30
  loc_00DEAFC0: lea eax, [esi+0000005Ch]
  loc_00DEAFC3: push ecx
  loc_00DEAFC4: mov var_64, eax
  loc_00DEAFC7: call [00401034h] ; __vbaStrVarMove
  loc_00DEAFCD: mov edx, eax
  loc_00DEAFCF: lea ecx, var_18
  loc_00DEAFD2: call [00401228h] ; __vbaStrMove
  loc_00DEAFD8: mov ecx, var_64
  loc_00DEAFDB: mov edx, eax
  loc_00DEAFDD: call [004011B0h] ; __vbaStrCopy
  loc_00DEAFE3: lea ecx, var_18
  loc_00DEAFE6: call [00401258h] ; __vbaFreeStr
  loc_00DEAFEC: lea ecx, var_30
  loc_00DEAFEF: call ebx
  loc_00DEAFF1: lea edx, var_40
  loc_00DEAFF4: push edi
  loc_00DEAFF5: push edx
  loc_00DEAFF6: mov var_38, 00000064h
  loc_00DEAFFD: mov var_40, 00008002h
  loc_00DEB004: call [00401210h] ; __vbaVarTstGe
  loc_00DEB00A: test ax, ax
  loc_00DEB00D: jz 00DEB124h
  loc_00DEB013: mov eax, [esi]
  loc_00DEB015: push esi
  loc_00DEB016: call [eax+00000334h]
  loc_00DEB01C: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00DEB022: lea ecx, var_1C
  loc_00DEB025: push eax
  loc_00DEB026: push ecx
  loc_00DEB027: call ebx
  loc_00DEB029: mov edi, eax
  loc_00DEB02B: push 4669AC00h
  loc_00DEB030: push edi
  loc_00DEB031: mov edx, [edi]
  loc_00DEB033: call [edx+0000007Ch]
  loc_00DEB036: test eax, eax
  loc_00DEB038: fnclex
  loc_00DEB03A: jge 00DEB04Bh
  loc_00DEB03C: push 0000007Ch
  loc_00DEB03E: push 006DCDA0h
  loc_00DEB043: push edi
  loc_00DEB044: push eax
  loc_00DEB045: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB04B: lea ecx, var_1C
  loc_00DEB04E: call [00401254h] ; __vbaFreeObj
  loc_00DEB054: mov eax, [esi]
  loc_00DEB056: push esi
  loc_00DEB057: call [eax+00000334h]
  loc_00DEB05D: lea ecx, var_1C
  loc_00DEB060: push eax
  loc_00DEB061: push ecx
  loc_00DEB062: call ebx
  loc_00DEB064: mov edi, eax
  loc_00DEB066: lea eax, var_44
  loc_00DEB069: push eax
  loc_00DEB06A: push edi
  loc_00DEB06B: mov edx, [edi]
  loc_00DEB06D: call [edx+00000078h]
  loc_00DEB070: test eax, eax
  loc_00DEB072: fnclex
  loc_00DEB074: jge 00DEB085h
  loc_00DEB076: push 00000078h
  loc_00DEB078: push 006DCDA0h
  loc_00DEB07D: push edi
  loc_00DEB07E: push eax
  loc_00DEB07F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB085: cmp var_44, 4669AC00h
  loc_00DEB08C: jnz 00DEB095h
  loc_00DEB08E: mov eax, 00000001h
  loc_00DEB093: jmp 00DEB097h
  loc_00DEB095: xor eax, eax
  loc_00DEB097: neg eax
  loc_00DEB099: lea ecx, var_1C
  loc_00DEB09C: mov di, ax
  loc_00DEB09F: call [00401254h] ; __vbaFreeObj
  loc_00DEB0A5: test di, di
  loc_00DEB0A8: jz 00DEB1DAh
  loc_00DEB0AE: mov ecx, [esi]
  loc_00DEB0B0: push esi
  loc_00DEB0B1: call [ecx+00000334h]
  loc_00DEB0B7: lea edx, var_1C
  loc_00DEB0BA: push eax
  loc_00DEB0BB: push edx
  loc_00DEB0BC: call ebx
  loc_00DEB0BE: mov edi, eax
  loc_00DEB0C0: push 3F800000h
  loc_00DEB0C5: push edi
  loc_00DEB0C6: mov eax, [edi]
  loc_00DEB0C8: call [eax+0000007Ch]
  loc_00DEB0CB: test eax, eax
  loc_00DEB0CD: fnclex
  loc_00DEB0CF: jge 00DEB0E0h
  loc_00DEB0D1: push 0000007Ch
  loc_00DEB0D3: push 006DCDA0h
  loc_00DEB0D8: push edi
  loc_00DEB0D9: push eax
  loc_00DEB0DA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB0E0: mov edi, [00401254h] ; __vbaFreeObj
  loc_00DEB0E6: lea ecx, var_1C
  loc_00DEB0E9: call edi
  loc_00DEB0EB: mov ecx, [esi]
  loc_00DEB0ED: push esi
  loc_00DEB0EE: call [ecx+000002FCh]
  loc_00DEB0F4: lea edx, var_1C
  loc_00DEB0F7: push eax
  loc_00DEB0F8: push edx
  loc_00DEB0F9: call ebx
  loc_00DEB0FB: mov esi, eax
  loc_00DEB0FD: push 00000000h
  loc_00DEB0FF: push esi
  loc_00DEB100: mov eax, [esi]
  loc_00DEB102: call [eax+0000005Ch]
  loc_00DEB105: test eax, eax
  loc_00DEB107: fnclex
  loc_00DEB109: jge 00DEB11Ah
  loc_00DEB10B: push 0000005Ch
  loc_00DEB10D: push 006DCAE0h
  loc_00DEB112: push esi
  loc_00DEB113: push eax
  loc_00DEB114: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB11A: lea ecx, var_1C
  loc_00DEB11D: call edi
  loc_00DEB11F: jmp 00DEB1DAh
  loc_00DEB124: mov ecx, [esi]
  loc_00DEB126: push esi
  loc_00DEB127: call [ecx+00000334h]
  loc_00DEB12D: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00DEB133: lea edx, var_20
  loc_00DEB136: push eax
  loc_00DEB137: push edx
  loc_00DEB138: call ebx
  loc_00DEB13A: mov edi, eax
  loc_00DEB13C: mov eax, [esi]
  loc_00DEB13E: push esi
  loc_00DEB13F: call [eax+00000334h]
  loc_00DEB145: lea ecx, var_1C
  loc_00DEB148: push eax
  loc_00DEB149: push ecx
  loc_00DEB14A: call ebx
  loc_00DEB14C: mov esi, eax
  loc_00DEB14E: lea eax, var_44
  loc_00DEB151: push eax
  loc_00DEB152: push esi
  loc_00DEB153: mov edx, [esi]
  loc_00DEB155: call [edx+00000078h]
  loc_00DEB158: test eax, eax
  loc_00DEB15A: fnclex
  loc_00DEB15C: jge 00DEB16Dh
  loc_00DEB15E: push 00000078h
  loc_00DEB160: push 006DCDA0h
  loc_00DEB165: push esi
  loc_00DEB166: push eax
  loc_00DEB167: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB16D: fld real4 ptr var_44
  loc_00DEB170: fadd st0, real4 ptr [004012D0h]
  loc_00DEB176: mov ecx, [edi]
  loc_00DEB178: push ecx
  loc_00DEB179: fnstsw ax
  loc_00DEB17B: test al, 0Dh
  loc_00DEB17D: jnz 00DEB22Dh
  loc_00DEB183: fstp real4 ptr [esp]
  loc_00DEB186: push edi
  loc_00DEB187: call [ecx+0000007Ch]
  loc_00DEB18A: test eax, eax
  loc_00DEB18C: fnclex
  loc_00DEB18E: jge 00DEB19Fh
  loc_00DEB190: push 0000007Ch
  loc_00DEB192: push 006DCDA0h
  loc_00DEB197: push edi
  loc_00DEB198: push eax
  loc_00DEB199: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB19F: lea edx, var_20
  loc_00DEB1A2: lea eax, var_1C
  loc_00DEB1A5: push edx
  loc_00DEB1A6: push eax
  loc_00DEB1A7: push 00000002h
  loc_00DEB1A9: call [00401048h] ; __vbaFreeObjList
  loc_00DEB1AF: mov ecx, var_64
  loc_00DEB1B2: add esp, 0000000Ch
  loc_00DEB1B5: mov edx, [ecx]
  loc_00DEB1B7: push edx
  loc_00DEB1B8: call [004011A4h] ; __vbaR8Str
  loc_00DEB1BE: call [0040124Ch] ; __vbaFPInt
  loc_00DEB1C4: mov ecx, var_60
  loc_00DEB1C7: lea edx, var_40
  loc_00DEB1CA: fstp real8 ptr var_38
  loc_00DEB1CD: mov var_40, 00000005h
  loc_00DEB1D4: call [0040101Ch] ; __vbaVarMove
  loc_00DEB1DA: xor ebx, ebx
  loc_00DEB1DC: mov var_4, ebx
  loc_00DEB1DF: fwait
  loc_00DEB1E0: push 00DEB20Eh
  loc_00DEB1E5: jmp 00DEB20Dh
  loc_00DEB1E7: lea ecx, var_18
  loc_00DEB1EA: call [00401258h] ; __vbaFreeStr
  loc_00DEB1F0: lea eax, var_20
  loc_00DEB1F3: lea ecx, var_1C
  loc_00DEB1F6: push eax
  loc_00DEB1F7: push ecx
  loc_00DEB1F8: push 00000002h
  loc_00DEB1FA: call [00401048h] ; __vbaFreeObjList
  loc_00DEB200: add esp, 0000000Ch
  loc_00DEB203: lea ecx, var_30
  loc_00DEB206: call [00401024h] ; __vbaFreeVar
  loc_00DEB20C: ret
  loc_00DEB20D: ret
  loc_00DEB20E: mov eax, Me
  loc_00DEB211: push eax
  loc_00DEB212: mov edx, [eax]
  loc_00DEB214: call [edx+00000008h]
  loc_00DEB217: mov eax, var_4
  loc_00DEB21A: mov ecx, var_14
  loc_00DEB21D: pop edi
  loc_00DEB21E: pop esi
  loc_00DEB21F: mov fs:[00000000h], ecx
  loc_00DEB226: pop ebx
  loc_00DEB227: mov esp, ebp
  loc_00DEB229: pop ebp
  loc_00DEB22A: retn 0004h
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer) 'DEB240
  loc_00DEB240: push ebp
  loc_00DEB241: mov ebp, esp
  loc_00DEB243: sub esp, 0000000Ch
  loc_00DEB246: push 00402836h ; __vbaExceptHandler
  loc_00DEB24B: mov eax, fs:[00000000h]
  loc_00DEB251: push eax
  loc_00DEB252: mov fs:[00000000h], esp
  loc_00DEB259: sub esp, 00000014h
  loc_00DEB25C: push ebx
  loc_00DEB25D: push esi
  loc_00DEB25E: push edi
  loc_00DEB25F: mov var_C, esp
  loc_00DEB262: mov var_8, 00401318h
  loc_00DEB269: mov esi, Me
  loc_00DEB26C: mov eax, esi
  loc_00DEB26E: and eax, 00000001h
  loc_00DEB271: mov var_4, eax
  loc_00DEB274: and esi, FFFFFFFEh
  loc_00DEB277: push esi
  loc_00DEB278: mov Me, esi
  loc_00DEB27B: mov ecx, [esi]
  loc_00DEB27D: call [ecx+00000004h]
  loc_00DEB280: mov edx, KeyAscii
  loc_00DEB283: mov var_18, 00000000h
  loc_00DEB28A: cmp [edx], 000Dh
  loc_00DEB28E: jnz 00DEB350h
  loc_00DEB294: mov eax, [esi]
  loc_00DEB296: push esi
  loc_00DEB297: call [eax+00000320h]
  loc_00DEB29D: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00DEB2A3: lea ecx, var_18
  loc_00DEB2A6: push eax
  loc_00DEB2A7: push ecx
  loc_00DEB2A8: call ebx
  loc_00DEB2AA: mov edi, eax
  loc_00DEB2AC: push 00000000h
  loc_00DEB2AE: push edi
  loc_00DEB2AF: mov edx, [edi]
  loc_00DEB2B1: call [edx+0000009Ch]
  loc_00DEB2B7: test eax, eax
  loc_00DEB2B9: fnclex
  loc_00DEB2BB: jge 00DEB2CFh
  loc_00DEB2BD: push 0000009Ch
  loc_00DEB2C2: push 006DCAD0h
  loc_00DEB2C7: push edi
  loc_00DEB2C8: push eax
  loc_00DEB2C9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB2CF: lea ecx, var_18
  loc_00DEB2D2: call [00401254h] ; __vbaFreeObj
  loc_00DEB2D8: mov eax, [esi]
  loc_00DEB2DA: push esi
  loc_00DEB2DB: call [eax+00000308h]
  loc_00DEB2E1: lea ecx, var_18
  loc_00DEB2E4: push eax
  loc_00DEB2E5: push ecx
  loc_00DEB2E6: call ebx
  loc_00DEB2E8: mov edi, eax
  loc_00DEB2EA: push FFFFFFFFh
  loc_00DEB2EC: push edi
  loc_00DEB2ED: mov edx, [edi]
  loc_00DEB2EF: call [edx+0000009Ch]
  loc_00DEB2F5: test eax, eax
  loc_00DEB2F7: fnclex
  loc_00DEB2F9: jge 00DEB30Dh
  loc_00DEB2FB: push 0000009Ch
  loc_00DEB300: push 006DCAD0h
  loc_00DEB305: push edi
  loc_00DEB306: push eax
  loc_00DEB307: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB30D: mov edi, [00401254h] ; __vbaFreeObj
  loc_00DEB313: lea ecx, var_18
  loc_00DEB316: call edi
  loc_00DEB318: mov eax, [esi]
  loc_00DEB31A: push esi
  loc_00DEB31B: call [eax+0000030Ch]
  loc_00DEB321: lea ecx, var_18
  loc_00DEB324: push eax
  loc_00DEB325: push ecx
  loc_00DEB326: call ebx
  loc_00DEB328: mov esi, eax
  loc_00DEB32A: push esi
  loc_00DEB32B: mov edx, [esi]
  loc_00DEB32D: call [edx+00000204h]
  loc_00DEB333: test eax, eax
  loc_00DEB335: fnclex
  loc_00DEB337: jge 00DEB34Bh
  loc_00DEB339: push 00000204h
  loc_00DEB33E: push 006DCB70h
  loc_00DEB343: push esi
  loc_00DEB344: push eax
  loc_00DEB345: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DEB34B: lea ecx, var_18
  loc_00DEB34E: call edi
  loc_00DEB350: mov var_4, 00000000h
  loc_00DEB357: push 00DEB369h
  loc_00DEB35C: jmp 00DEB368h
  loc_00DEB35E: lea ecx, var_18
  loc_00DEB361: call [00401254h] ; __vbaFreeObj
  loc_00DEB367: ret
  loc_00DEB368: ret
  loc_00DEB369: mov eax, Me
  loc_00DEB36C: push eax
  loc_00DEB36D: mov ecx, [eax]
  loc_00DEB36F: call [ecx+00000008h]
  loc_00DEB372: mov eax, var_4
  loc_00DEB375: mov ecx, var_14
  loc_00DEB378: pop edi
  loc_00DEB379: pop esi
  loc_00DEB37A: mov fs:[00000000h], ecx
  loc_00DEB381: pop ebx
  loc_00DEB382: mov esp, ebp
  loc_00DEB384: pop ebp
  loc_00DEB385: retn 0008h
End Sub

Private Sub Form_Load() 'DE90F0
  loc_00DE90F0: push ebp
  loc_00DE90F1: mov ebp, esp
  loc_00DE90F3: sub esp, 0000000Ch
  loc_00DE90F6: push 00402836h ; __vbaExceptHandler
  loc_00DE90FB: mov eax, fs:[00000000h]
  loc_00DE9101: push eax
  loc_00DE9102: mov fs:[00000000h], esp
  loc_00DE9109: sub esp, 00000014h
  loc_00DE910C: push ebx
  loc_00DE910D: push esi
  loc_00DE910E: push edi
  loc_00DE910F: mov var_C, esp
  loc_00DE9112: mov var_8, 00401288h
  loc_00DE9119: mov esi, Me
  loc_00DE911C: mov eax, esi
  loc_00DE911E: and eax, 00000001h
  loc_00DE9121: mov var_4, eax
  loc_00DE9124: and esi, FFFFFFFEh
  loc_00DE9127: push esi
  loc_00DE9128: mov Me, esi
  loc_00DE912B: mov ecx, [esi]
  loc_00DE912D: call [ecx+00000004h]
  loc_00DE9130: mov edx, [esi]
  loc_00DE9132: push esi
  loc_00DE9133: mov var_18, 00000000h
  loc_00DE913A: call [edx+00000314h]
  loc_00DE9140: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00DE9146: push eax
  loc_00DE9147: lea eax, var_18
  loc_00DE914A: push eax
  loc_00DE914B: call ebx
  loc_00DE914D: mov edi, eax
  loc_00DE914F: push FFFFFFFFh
  loc_00DE9151: push edi
  loc_00DE9152: mov ecx, [edi]
  loc_00DE9154: call [ecx+0000008Ch]
  loc_00DE915A: test eax, eax
  loc_00DE915C: fnclex
  loc_00DE915E: jge 00DE9172h
  loc_00DE9160: push 0000008Ch
  loc_00DE9165: push 006DCAC0h
  loc_00DE916A: push edi
  loc_00DE916B: push eax
  loc_00DE916C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9172: lea ecx, var_18
  loc_00DE9175: call [00401254h] ; __vbaFreeObj
  loc_00DE917B: mov edx, [esi]
  loc_00DE917D: push esi
  loc_00DE917E: call [edx+00000318h]
  loc_00DE9184: push eax
  loc_00DE9185: lea eax, var_18
  loc_00DE9188: push eax
  loc_00DE9189: call ebx
  loc_00DE918B: mov edi, eax
  loc_00DE918D: push 00000000h
  loc_00DE918F: push edi
  loc_00DE9190: mov ecx, [edi]
  loc_00DE9192: call [ecx+0000008Ch]
  loc_00DE9198: test eax, eax
  loc_00DE919A: fnclex
  loc_00DE919C: jge 00DE91B0h
  loc_00DE919E: push 0000008Ch
  loc_00DE91A3: push 006DCAC0h
  loc_00DE91A8: push edi
  loc_00DE91A9: push eax
  loc_00DE91AA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE91B0: lea ecx, var_18
  loc_00DE91B3: call [00401254h] ; __vbaFreeObj
  loc_00DE91B9: mov edx, [esi]
  loc_00DE91BB: push esi
  loc_00DE91BC: call [edx+00000308h]
  loc_00DE91C2: push eax
  loc_00DE91C3: lea eax, var_18
  loc_00DE91C6: push eax
  loc_00DE91C7: call ebx
  loc_00DE91C9: mov edi, eax
  loc_00DE91CB: push 00000000h
  loc_00DE91CD: push edi
  loc_00DE91CE: mov ecx, [edi]
  loc_00DE91D0: call [ecx+0000009Ch]
  loc_00DE91D6: test eax, eax
  loc_00DE91D8: fnclex
  loc_00DE91DA: jge 00DE91EEh
  loc_00DE91DC: push 0000009Ch
  loc_00DE91E1: push 006DCAD0h
  loc_00DE91E6: push edi
  loc_00DE91E7: push eax
  loc_00DE91E8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE91EE: lea ecx, var_18
  loc_00DE91F1: call [00401254h] ; __vbaFreeObj
  loc_00DE91F7: mov edx, [esi]
  loc_00DE91F9: push esi
  loc_00DE91FA: call [edx+00000304h]
  loc_00DE9200: push eax
  loc_00DE9201: lea eax, var_18
  loc_00DE9204: push eax
  loc_00DE9205: call ebx
  loc_00DE9207: mov edi, eax
  loc_00DE9209: push 00000000h
  loc_00DE920B: push edi
  loc_00DE920C: mov ecx, [edi]
  loc_00DE920E: call [ecx+0000005Ch]
  loc_00DE9211: test eax, eax
  loc_00DE9213: fnclex
  loc_00DE9215: jge 00DE9226h
  loc_00DE9217: push 0000005Ch
  loc_00DE9219: push 006DCAE0h
  loc_00DE921E: push edi
  loc_00DE921F: push eax
  loc_00DE9220: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9226: lea ecx, var_18
  loc_00DE9229: call [00401254h] ; __vbaFreeObj
  loc_00DE922F: mov edx, [esi]
  loc_00DE9231: push esi
  loc_00DE9232: call [edx+00000300h]
  loc_00DE9238: push eax
  loc_00DE9239: lea eax, var_18
  loc_00DE923C: push eax
  loc_00DE923D: call ebx
  loc_00DE923F: mov edi, eax
  loc_00DE9241: push 00000000h
  loc_00DE9243: push edi
  loc_00DE9244: mov ecx, [edi]
  loc_00DE9246: call [ecx+0000005Ch]
  loc_00DE9249: test eax, eax
  loc_00DE924B: fnclex
  loc_00DE924D: jge 00DE925Eh
  loc_00DE924F: push 0000005Ch
  loc_00DE9251: push 006DCAE0h
  loc_00DE9256: push edi
  loc_00DE9257: push eax
  loc_00DE9258: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE925E: mov edi, [00401254h] ; __vbaFreeObj
  loc_00DE9264: lea ecx, var_18
  loc_00DE9267: call edi
  loc_00DE9269: mov edx, [esi]
  loc_00DE926B: push esi
  loc_00DE926C: call [edx+000002FCh]
  loc_00DE9272: push eax
  loc_00DE9273: lea eax, var_18
  loc_00DE9276: push eax
  loc_00DE9277: call ebx
  loc_00DE9279: mov esi, eax
  loc_00DE927B: push 00000000h
  loc_00DE927D: push esi
  loc_00DE927E: mov ecx, [esi]
  loc_00DE9280: call [ecx+0000005Ch]
  loc_00DE9283: test eax, eax
  loc_00DE9285: fnclex
  loc_00DE9287: jge 00DE9298h
  loc_00DE9289: push 0000005Ch
  loc_00DE928B: push 006DCAE0h
  loc_00DE9290: push esi
  loc_00DE9291: push eax
  loc_00DE9292: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9298: lea ecx, var_18
  loc_00DE929B: call edi
  loc_00DE929D: mov var_4, 00000000h
  loc_00DE92A4: push 00DE92B6h
  loc_00DE92A9: jmp 00DE92B5h
  loc_00DE92AB: lea ecx, var_18
  loc_00DE92AE: call [00401254h] ; __vbaFreeObj
  loc_00DE92B4: ret
  loc_00DE92B5: ret
  loc_00DE92B6: mov eax, Me
  loc_00DE92B9: push eax
  loc_00DE92BA: mov edx, [eax]
  loc_00DE92BC: call [edx+00000008h]
  loc_00DE92BF: mov eax, var_4
  loc_00DE92C2: mov ecx, var_14
  loc_00DE92C5: pop edi
  loc_00DE92C6: pop esi
  loc_00DE92C7: mov fs:[00000000h], ecx
  loc_00DE92CE: pop ebx
  loc_00DE92CF: mov esp, ebp
  loc_00DE92D1: pop ebp
  loc_00DE92D2: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'DE92E0
  loc_00DE92E0: push ebp
  loc_00DE92E1: mov ebp, esp
  loc_00DE92E3: sub esp, 0000000Ch
  loc_00DE92E6: push 00402836h ; __vbaExceptHandler
  loc_00DE92EB: mov eax, fs:[00000000h]
  loc_00DE92F1: push eax
  loc_00DE92F2: mov fs:[00000000h], esp
  loc_00DE92F9: sub esp, 0000005Ch
  loc_00DE92FC: push ebx
  loc_00DE92FD: push esi
  loc_00DE92FE: push edi
  loc_00DE92FF: mov var_C, esp
  loc_00DE9302: mov var_8, 004012A0h
  loc_00DE9309: mov esi, Me
  loc_00DE930C: mov eax, esi
  loc_00DE930E: and eax, 00000001h
  loc_00DE9311: mov var_4, eax
  loc_00DE9314: and esi, FFFFFFFEh
  loc_00DE9317: push esi
  loc_00DE9318: mov Me, esi
  loc_00DE931B: mov ecx, [esi]
  loc_00DE931D: call [ecx+00000004h]
  loc_00DE9320: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9326: xor eax, eax
  loc_00DE9328: mov var_18, eax
  loc_00DE932B: mov var_4C, eax
  loc_00DE932E: mov var_50, eax
  loc_00DE9331: mov edx, [esi]
  loc_00DE9333: lea eax, var_4C
  loc_00DE9336: push eax
  loc_00DE9337: push esi
  loc_00DE9338: call [edx+00000070h]
  loc_00DE933B: test eax, eax
  loc_00DE933D: fnclex
  loc_00DE933F: jge 00DE934Ch
  loc_00DE9341: push 00000070h
  loc_00DE9343: push 006DC604h
  loc_00DE9348: push esi
  loc_00DE9349: push eax
  loc_00DE934A: call ebx
  loc_00DE934C: fld real4 ptr var_4C
  loc_00DE934F: fadd st0, real4 ptr [00401298h]
  loc_00DE9355: mov ecx, [esi]
  loc_00DE9357: push ecx
  loc_00DE9358: fnstsw ax
  loc_00DE935A: test al, 0Dh
  loc_00DE935C: jnz 00DE9520h
  loc_00DE9362: fstp real4 ptr [esp]
  loc_00DE9365: push esi
  loc_00DE9366: call [ecx+00000074h]
  loc_00DE9369: test eax, eax
  loc_00DE936B: fnclex
  loc_00DE936D: jge 00DE937Ah
  loc_00DE936F: push 00000074h
  loc_00DE9371: push 006DC604h
  loc_00DE9376: push esi
  loc_00DE9377: push eax
  loc_00DE9378: call ebx
  loc_00DE937A: mov edx, [esi]
  loc_00DE937C: lea eax, var_4C
  loc_00DE937F: push eax
  loc_00DE9380: push esi
  loc_00DE9381: call [edx+00000070h]
  loc_00DE9384: test eax, eax
  loc_00DE9386: fnclex
  loc_00DE9388: jge 00DE9395h
  loc_00DE938A: push 00000070h
  loc_00DE938C: push 006DC604h
  loc_00DE9391: push esi
  loc_00DE9392: push eax
  loc_00DE9393: call ebx
  loc_00DE9395: mov ecx, [esi]
  loc_00DE9397: lea edx, var_50
  loc_00DE939A: push edx
  loc_00DE939B: push esi
  loc_00DE939C: call [ecx+00000078h]
  loc_00DE939F: test eax, eax
  loc_00DE93A1: fnclex
  loc_00DE93A3: jge 00DE93B0h
  loc_00DE93A5: push 00000078h
  loc_00DE93A7: push 006DC604h
  loc_00DE93AC: push esi
  loc_00DE93AD: push eax
  loc_00DE93AE: call ebx
  loc_00DE93B0: sub esp, 00000010h
  loc_00DE93B3: mov ecx, 0000000Ah
  loc_00DE93B8: mov ebx, esp
  loc_00DE93BA: mov eax, 80020004h
  loc_00DE93BF: mov edx, eax
  loc_00DE93C1: sub esp, 00000010h
  loc_00DE93C4: mov [ebx], ecx
  loc_00DE93C6: mov ecx, var_44
  loc_00DE93C9: mov edi, [esi]
  loc_00DE93CB: mov [ebx+00000004h], ecx
  loc_00DE93CE: mov ecx, esp
  loc_00DE93D0: sub esp, 00000010h
  loc_00DE93D3: mov [ebx+00000008h], eax
  loc_00DE93D6: mov eax, var_3C
  loc_00DE93D9: mov [ebx+0000000Ch], eax
  loc_00DE93DC: mov eax, 0000000Ah
  loc_00DE93E1: mov [ecx], eax
  loc_00DE93E3: mov eax, var_34
  loc_00DE93E6: mov [ecx+00000004h], eax
  loc_00DE93E9: mov eax, 00000004h
  loc_00DE93EE: mov [ecx+00000008h], edx
  loc_00DE93F1: mov edx, var_2C
  loc_00DE93F4: mov [ecx+0000000Ch], edx
  loc_00DE93F7: mov edx, var_24
  loc_00DE93FA: mov ecx, esp
  loc_00DE93FC: mov [ecx], eax
  loc_00DE93FE: mov eax, var_50
  loc_00DE9401: mov [ecx+00000004h], edx
  loc_00DE9404: mov [ecx+00000008h], eax
  loc_00DE9407: mov eax, var_1C
  loc_00DE940A: mov [ecx+0000000Ch], eax
  loc_00DE940D: mov ecx, var_4C
  loc_00DE9410: push ecx
  loc_00DE9411: push esi
  loc_00DE9412: call [edi+000002A4h]
  loc_00DE9418: test eax, eax
  loc_00DE941A: fnclex
  loc_00DE941C: jge 00DE9434h
  loc_00DE941E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE9424: push 000002A4h
  loc_00DE9429: push 006DC604h
  loc_00DE942E: push esi
  loc_00DE942F: push eax
  loc_00DE9430: call ebx
  loc_00DE9432: jmp 00DE943Ah
  loc_00DE9434: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE943A: call [004010BCh] ; rtcDoEvents
  loc_00DE9440: mov edx, [esi]
  loc_00DE9442: lea eax, var_4C
  loc_00DE9445: push eax
  loc_00DE9446: push esi
  loc_00DE9447: call [edx+00000070h]
  loc_00DE944A: test eax, eax
  loc_00DE944C: fnclex
  loc_00DE944E: jge 00DE945Bh
  loc_00DE9450: push 00000070h
  loc_00DE9452: push 006DC604h
  loc_00DE9457: push esi
  loc_00DE9458: push eax
  loc_00DE9459: call ebx
  loc_00DE945B: mov eax, [00E3D9CCh]
  loc_00DE9460: test eax, eax
  loc_00DE9462: jnz 00DE9474h
  loc_00DE9464: push 00E3D9CCh
  loc_00DE9469: push 006DC960h
  loc_00DE946E: call [004011A0h] ; __vbaNew2
  loc_00DE9474: mov edi, [00E3D9CCh]
  loc_00DE947A: lea edx, var_18
  loc_00DE947D: push edx
  loc_00DE947E: push edi
  loc_00DE947F: mov ecx, [edi]
  loc_00DE9481: call [ecx+00000018h]
  loc_00DE9484: test eax, eax
  loc_00DE9486: fnclex
  loc_00DE9488: jge 00DE9495h
  loc_00DE948A: push 00000018h
  loc_00DE948C: push 006DC950h
  loc_00DE9491: push edi
  loc_00DE9492: push eax
  loc_00DE9493: call ebx
  loc_00DE9495: mov eax, var_18
  loc_00DE9498: lea edx, var_50
  loc_00DE949B: push edx
  loc_00DE949C: push eax
  loc_00DE949D: mov ecx, [eax]
  loc_00DE949F: mov edi, eax
  loc_00DE94A1: call [ecx+00000098h]
  loc_00DE94A7: test eax, eax
  loc_00DE94A9: fnclex
  loc_00DE94AB: jge 00DE94BBh
  loc_00DE94AD: push 00000098h
  loc_00DE94B2: push 006DCAF0h
  loc_00DE94B7: push edi
  loc_00DE94B8: push eax
  loc_00DE94B9: call ebx
  loc_00DE94BB: fld real4 ptr var_4C
  loc_00DE94BE: fcomp real4 ptr var_50
  loc_00DE94C1: fnstsw ax
  loc_00DE94C3: test ah, 41h
  loc_00DE94C6: jz 00DE94CFh
  loc_00DE94C8: mov eax, 00000001h
  loc_00DE94CD: jmp 00DE94D1h
  loc_00DE94CF: xor eax, eax
  loc_00DE94D1: neg eax
  loc_00DE94D3: lea ecx, var_18
  loc_00DE94D6: mov edi, eax
  loc_00DE94D8: call [00401254h] ; __vbaFreeObj
  loc_00DE94DE: test di, di
  loc_00DE94E1: jnz 00DE9331h
  loc_00DE94E7: mov var_4, 00000000h
  loc_00DE94EE: fwait
  loc_00DE94EF: push 00DE9501h
  loc_00DE94F4: jmp 00DE9500h
  loc_00DE94F6: lea ecx, var_18
  loc_00DE94F9: call [00401254h] ; __vbaFreeObj
  loc_00DE94FF: ret
  loc_00DE9500: ret
  loc_00DE9501: mov eax, Me
  loc_00DE9504: push eax
  loc_00DE9505: mov ecx, [eax]
  loc_00DE9507: call [ecx+00000008h]
  loc_00DE950A: mov eax, var_4
  loc_00DE950D: mov ecx, var_14
  loc_00DE9510: pop edi
  loc_00DE9511: pop esi
  loc_00DE9512: mov fs:[00000000h], ecx
  loc_00DE9519: pop ebx
  loc_00DE951A: mov esp, ebp
  loc_00DE951C: pop ebp
  loc_00DE951D: retn 0008h
End Sub

Private Sub Proc_0_10_DE8E00
  loc_00DE8E00: push ebp
  loc_00DE8E01: mov ebp, esp
  loc_00DE8E03: sub esp, 00000008h
  loc_00DE8E06: push 00402836h ; __vbaExceptHandler
  loc_00DE8E0B: mov eax, fs:[00000000h]
  loc_00DE8E11: push eax
  loc_00DE8E12: mov fs:[00000000h], esp
  loc_00DE8E19: sub esp, 000000B0h
  loc_00DE8E1F: push ebx
  loc_00DE8E20: push esi
  loc_00DE8E21: push edi
  loc_00DE8E22: mov var_8, esp
  loc_00DE8E25: mov var_4, 00401268h
  loc_00DE8E2C: mov edi, [004011F0h] ; __vbaVarDup
  loc_00DE8E32: mov ecx, 80020004h
  loc_00DE8E37: xor esi, esi
  loc_00DE8E39: mov var_5C, ecx
  loc_00DE8E3C: mov eax, 0000000Ah
  loc_00DE8E41: mov var_4C, ecx
  loc_00DE8E44: mov ebx, 00000008h
  loc_00DE8E49: lea edx, var_84
  loc_00DE8E4F: lea ecx, var_44
  loc_00DE8E52: mov var_20, esi
  loc_00DE8E55: mov var_24, esi
  loc_00DE8E58: mov var_34, esi
  loc_00DE8E5B: mov var_44, esi
  loc_00DE8E5E: mov var_74, esi
  loc_00DE8E61: mov var_B4, esi
  loc_00DE8E67: mov var_64, eax
  loc_00DE8E6A: mov var_54, eax
  loc_00DE8E6D: mov var_7C, 006DC934h ; "Alert"
  loc_00DE8E74: mov var_84, ebx
  loc_00DE8E7A: call edi
  loc_00DE8E7C: lea edx, var_74
  loc_00DE8E7F: lea ecx, var_34
  loc_00DE8E82: mov var_6C, 006DC8E0h ; "Apakah anda ingin keluar dari program?"
  loc_00DE8E89: mov var_74, ebx
  loc_00DE8E8C: call edi
  loc_00DE8E8E: lea eax, var_64
  loc_00DE8E91: lea ecx, var_54
  loc_00DE8E94: push eax
  loc_00DE8E95: lea edx, var_44
  loc_00DE8E98: push ecx
  loc_00DE8E99: push edx
  loc_00DE8E9A: lea eax, var_34
  loc_00DE8E9D: push 00000044h
  loc_00DE8E9F: push eax
  loc_00DE8EA0: call [004010A8h] ; rtcMsgBox
  loc_00DE8EA6: lea edx, var_B4
  loc_00DE8EAC: lea ecx, var_20
  loc_00DE8EAF: mov var_AC, eax
  loc_00DE8EB5: mov var_B4, 00000003h
  loc_00DE8EBF: call [0040101Ch] ; __vbaVarMove
  loc_00DE8EC5: lea ecx, var_64
  loc_00DE8EC8: lea edx, var_54
  loc_00DE8ECB: push ecx
  loc_00DE8ECC: lea eax, var_44
  loc_00DE8ECF: push edx
  loc_00DE8ED0: lea ecx, var_34
  loc_00DE8ED3: push eax
  loc_00DE8ED4: push ecx
  loc_00DE8ED5: push 00000004h
  loc_00DE8ED7: call [00401038h] ; __vbaFreeVarList
  loc_00DE8EDD: add esp, 00000014h
  loc_00DE8EE0: lea edx, var_20
  loc_00DE8EE3: lea eax, var_74
  loc_00DE8EE6: mov var_6C, 00000006h
  loc_00DE8EED: push edx
  loc_00DE8EEE: push eax
  loc_00DE8EEF: mov var_74, 00008003h
  loc_00DE8EF6: call [00401108h] ; __vbaVarTstEq
  loc_00DE8EFC: test ax, ax
  loc_00DE8EFF: jz 00DE8F52h
  loc_00DE8F01: cmp [00E3D9CCh], esi
  loc_00DE8F07: jnz 00DE8F19h
  loc_00DE8F09: push 00E3D9CCh
  loc_00DE8F0E: push 006DC960h
  loc_00DE8F13: call [004011A0h] ; __vbaNew2
  loc_00DE8F19: mov ecx, Me
  loc_00DE8F1C: mov edi, [00E3D9CCh]
  loc_00DE8F22: lea edx, var_24
  loc_00DE8F25: push ecx
  loc_00DE8F26: mov ebx, [edi]
  loc_00DE8F28: push edx
  loc_00DE8F29: call [004010B4h] ; __vbaObjSetAddref
  loc_00DE8F2F: push eax
  loc_00DE8F30: push edi
  loc_00DE8F31: call [ebx+00000010h]
  loc_00DE8F34: cmp eax, esi
  loc_00DE8F36: fnclex
  loc_00DE8F38: jge 00DE8F49h
  loc_00DE8F3A: push 00000010h
  loc_00DE8F3C: push 006DC950h
  loc_00DE8F41: push edi
  loc_00DE8F42: push eax
  loc_00DE8F43: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00DE8F49: lea ecx, var_24
  loc_00DE8F4C: call [00401254h] ; __vbaFreeObj
  loc_00DE8F52: push 00DE8F88h
  loc_00DE8F57: jmp 00DE8F7Eh
  loc_00DE8F59: lea ecx, var_24
  loc_00DE8F5C: call [00401254h] ; __vbaFreeObj
  loc_00DE8F62: lea eax, var_64
  loc_00DE8F65: lea ecx, var_54
  loc_00DE8F68: push eax
  loc_00DE8F69: lea edx, var_44
  loc_00DE8F6C: push ecx
  loc_00DE8F6D: lea eax, var_34
  loc_00DE8F70: push edx
  loc_00DE8F71: push eax
  loc_00DE8F72: push 00000004h
  loc_00DE8F74: call [00401038h] ; __vbaFreeVarList
  loc_00DE8F7A: add esp, 00000014h
  loc_00DE8F7D: ret
  loc_00DE8F7E: lea ecx, var_20
  loc_00DE8F81: call [00401024h] ; __vbaFreeVar
  loc_00DE8F87: ret
  loc_00DE8F88: mov ecx, var_10
  loc_00DE8F8B: pop edi
  loc_00DE8F8C: pop esi
  loc_00DE8F8D: xor eax, eax
  loc_00DE8F8F: mov fs:[00000000h], ecx
  loc_00DE8F96: pop ebx
  loc_00DE8F97: mov esp, ebp
  loc_00DE8F99: pop ebp
  loc_00DE8F9A: retn 0004h
End Sub
