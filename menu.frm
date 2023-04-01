VERSION 5.00
Object = "{6BF52A50-394A-11d3-B15300C04F79FAA6}#1.0#0"; "%SystemRoot%\system32\wmp.dll"
Begin VB.Form menu
  Caption = "Form1"
  BackColor = &HFFFFFF&
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  BorderStyle = 0 'None
  Icon = "menu.frx":0000
  LinkTopic = "Form1"
  ClientLeft = 0
  ClientTop = 0
  ClientWidth = 6825
  ClientHeight = 3870
  StartUpPosition = 2 'CenterScreen
  Begin VB.Timer Timer4
    Interval = 10
    Left = 5310
    Top = 3450
  End
  Begin VB.Frame fcetak
    Caption = "Frame1"
    BackColor = &H80FF80&
    Left = 1680
    Top = 1290
    Width = 2385
    Height = 1275
    TabIndex = 8
    BorderStyle = 0 'None
    Begin Project1.jcbutton jcbutton1
      Left = 0
      Top = 60
      Width = 2385
      Height = 375
      TabIndex = 9
      OleObjectBlob = "menu.frx":0BDA
    End
    Begin Project1.jcbutton jcbutton2
      Left = -30
      Top = 480
      Width = 2415
      Height = 345
      TabIndex = 10
      OleObjectBlob = "menu.frx":0E26
    End
    Begin Project1.jcbutton jcbutton3
      Left = -30
      Top = 870
      Width = 2415
      Height = 345
      TabIndex = 11
      OleObjectBlob = "menu.frx":106A
    End
  End
  Begin VB.Timer Timer1
    Left = 6390
    Top = 3420
  End
  Begin VB.Frame fmenu
    Caption = "Frame1"
    BackColor = &H404040&
    Left = -60
    Top = 510
    Width = 1755
    Height = 3375
    TabIndex = 2
    BorderStyle = 0 'None
    Begin Project1.jcbutton daftar
      Left = 0
      Top = 390
      Width = 1755
      Height = 435
      TabIndex = 3
      OleObjectBlob = "menu.frx":12C2
    End
    Begin Project1.jcbutton biaya
      Left = 30
      Top = 1200
      Width = 1755
      Height = 435
      TabIndex = 4
      OleObjectBlob = "menu.frx":150A
    End
    Begin Project1.jcbutton cetak
      Left = 30
      Top = 780
      Width = 1755
      Height = 435
      TabIndex = 5
      OleObjectBlob = "menu.frx":1752
    End
    Begin Project1.jcbutton logind
      Left = 30
      Top = 0
      Width = 1755
      Height = 435
      TabIndex = 6
      OleObjectBlob = "menu.frx":1992
    End
    Begin VB.Image Image1
      Picture = "menu.frx":1BC6
      Left = 300
      Top = 2940
      Width = 375
      Height = 360
      Stretch = -1  'True
    End
    Begin VB.Image ext
      Picture = "menu.frx":2008
      Left = 1200
      Top = 3000
      Width = 255
      Height = 225
      Stretch = -1  'True
    End
  End
  Begin Project1.jcbutton menus
    Left = 0
    Top = 270
    Width = 1665
    Height = 225
    TabIndex = 1
    OleObjectBlob = "menu.frx":271F
  End
  Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1
    Left = 5820
    Top = 3420
    Width = 525
    Height = 435
    Visible = 0   'False
    TabIndex = 7
    OleObjectBlob = "menu.frx":294F
  End
  Begin VB.Shape Shape3
    Left = -30
    Top = 3480
    Width = 6945
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &HE0E0E0&
    FillStyle = 0
  End
  Begin VB.Image Image3
End

Attribute VB_Name = "menu"


Private Sub logind_UnknownEvent_9 'E068A0
  loc_00E068A0: push ebp
  loc_00E068A1: mov ebp, esp
  loc_00E068A3: sub esp, 0000000Ch
  loc_00E068A6: push 00402836h ; __vbaExceptHandler
  loc_00E068AB: mov eax, fs:[00000000h]
  loc_00E068B1: push eax
  loc_00E068B2: mov fs:[00000000h], esp
  loc_00E068B9: sub esp, 00000034h
  loc_00E068BC: push ebx
  loc_00E068BD: push esi
  loc_00E068BE: push edi
  loc_00E068BF: mov var_C, esp
  loc_00E068C2: mov var_8, 00401FE0h
  loc_00E068C9: mov esi, Me
  loc_00E068CC: mov eax, esi
  loc_00E068CE: and eax, 00000001h
  loc_00E068D1: mov var_4, eax
  loc_00E068D4: and esi, FFFFFFFEh
  loc_00E068D7: push esi
  loc_00E068D8: mov Me, esi
  loc_00E068DB: mov ecx, [esi]
  loc_00E068DD: call [ecx+00000004h]
  loc_00E068E0: mov eax, [00E3D024h]
  loc_00E068E5: mov var_18, 00000000h
  loc_00E068EC: test eax, eax
  loc_00E068EE: jnz 00E06900h
  loc_00E068F0: push 00E3D024h
  loc_00E068F5: push 006CE120h
  loc_00E068FA: call [004011A0h] ; __vbaNew2
  loc_00E06900: mov edi, [00E3D024h]
  loc_00E06906: push edi
  loc_00E06907: mov edx, [edi]
  loc_00E06909: call [edx+000002B4h]
  loc_00E0690F: test eax, eax
  loc_00E06911: fnclex
  loc_00E06913: jge 00E06927h
  loc_00E06915: push 000002B4h
  loc_00E0691A: push 006DC970h
  loc_00E0691F: push edi
  loc_00E06920: push eax
  loc_00E06921: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06927: mov eax, [00E3D010h]
  loc_00E0692C: test eax, eax
  loc_00E0692E: jnz 00E06940h
  loc_00E06930: push 00E3D010h
  loc_00E06935: push 006D058Ch
  loc_00E0693A: call [004011A0h] ; __vbaNew2
  loc_00E06940: sub esp, 00000010h
  loc_00E06943: mov ecx, 0000000Ah
  loc_00E06948: mov ebx, esp
  loc_00E0694A: mov var_28, ecx
  loc_00E0694D: mov eax, 80020004h
  loc_00E06952: sub esp, 00000010h
  loc_00E06955: mov [ebx], ecx
  loc_00E06957: mov ecx, var_34
  loc_00E0695A: mov var_20, eax
  loc_00E0695D: mov edi, [00E3D010h]
  loc_00E06963: mov [ebx+00000004h], ecx
  loc_00E06966: mov ecx, esp
  loc_00E06968: mov edx, [edi]
  loc_00E0696A: push edi
  loc_00E0696B: mov [ebx+00000008h], eax
  loc_00E0696E: mov eax, var_2C
  loc_00E06971: mov [ebx+0000000Ch], eax
  loc_00E06974: mov eax, var_28
  loc_00E06977: mov ebx, var_24
  loc_00E0697A: mov [ecx], eax
  loc_00E0697C: mov eax, var_20
  loc_00E0697F: mov [ecx+00000004h], ebx
  loc_00E06982: mov [ecx+00000008h], eax
  loc_00E06985: mov eax, var_1C
  loc_00E06988: mov [ecx+0000000Ch], eax
  loc_00E0698B: call [edx+000002B0h]
  loc_00E06991: test eax, eax
  loc_00E06993: fnclex
  loc_00E06995: jge 00E069A9h
  loc_00E06997: push 000002B0h
  loc_00E0699C: push 006DC604h
  loc_00E069A1: push edi
  loc_00E069A2: push eax
  loc_00E069A3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E069A9: sub esp, 00000010h
  loc_00E069AC: mov ecx, 00000003h
  loc_00E069B1: mov edx, esp
  loc_00E069B3: mov eax, 00008000h
  loc_00E069B8: push FFFFFE0Bh
  loc_00E069BD: push esi
  loc_00E069BE: mov [edx], ecx
  loc_00E069C0: mov ecx, [esi]
  loc_00E069C2: mov [edx+00000004h], ebx
  loc_00E069C5: mov [edx+00000008h], eax
  loc_00E069C8: mov eax, var_1C
  loc_00E069CB: mov [edx+0000000Ch], eax
  loc_00E069CE: call [ecx+00000330h]
  loc_00E069D4: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E069DA: lea edx, var_18
  loc_00E069DD: push eax
  loc_00E069DE: push edx
  loc_00E069DF: call edi
  loc_00E069E1: push eax
  loc_00E069E2: call [00401238h] ; __vbaLateIdSt
  loc_00E069E8: lea ecx, var_18
  loc_00E069EB: call [00401254h] ; __vbaFreeObj
  loc_00E069F1: sub esp, 00000010h
  loc_00E069F4: mov ecx, 00000003h
  loc_00E069F9: mov edx, esp
  loc_00E069FB: mov eax, 0080FF80h
  loc_00E06A00: push FFFFFDFFh
  loc_00E06A05: push esi
  loc_00E06A06: mov [edx], ecx
  loc_00E06A08: mov ecx, [esi]
  loc_00E06A0A: mov [edx+00000004h], ebx
  loc_00E06A0D: mov [edx+00000008h], eax
  loc_00E06A10: mov eax, var_1C
  loc_00E06A13: mov [edx+0000000Ch], eax
  loc_00E06A16: call [ecx+00000330h]
  loc_00E06A1C: lea edx, var_18
  loc_00E06A1F: push eax
  loc_00E06A20: push edx
  loc_00E06A21: call edi
  loc_00E06A23: push eax
  loc_00E06A24: call [00401238h] ; __vbaLateIdSt
  loc_00E06A2A: lea ecx, var_18
  loc_00E06A2D: call [00401254h] ; __vbaFreeObj
  loc_00E06A33: sub esp, 00000010h
  loc_00E06A36: mov ecx, 00000008h
  loc_00E06A3B: mov edx, esp
  loc_00E06A3D: mov eax, 006E0398h ; "SHOW"
  loc_00E06A42: push FFFFFDFAh
  loc_00E06A47: push esi
  loc_00E06A48: mov [edx], ecx
  loc_00E06A4A: mov ecx, [esi]
  loc_00E06A4C: mov [edx+00000004h], ebx
  loc_00E06A4F: mov [edx+00000008h], eax
  loc_00E06A52: mov eax, var_1C
  loc_00E06A55: mov [edx+0000000Ch], eax
  loc_00E06A58: call [ecx+00000330h]
  loc_00E06A5E: lea edx, var_18
  loc_00E06A61: push eax
  loc_00E06A62: push edx
  loc_00E06A63: call edi
  loc_00E06A65: push eax
  loc_00E06A66: call [00401238h] ; __vbaLateIdSt
  loc_00E06A6C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E06A72: lea ecx, var_18
  loc_00E06A75: call ebx
  loc_00E06A77: mov eax, [esi]
  loc_00E06A79: push esi
  loc_00E06A7A: call [eax+00000314h]
  loc_00E06A80: lea ecx, var_18
  loc_00E06A83: push eax
  loc_00E06A84: push ecx
  loc_00E06A85: call edi
  loc_00E06A87: mov esi, eax
  loc_00E06A89: push 00000000h
  loc_00E06A8B: push esi
  loc_00E06A8C: mov edx, [esi]
  loc_00E06A8E: call [edx+0000009Ch]
  loc_00E06A94: test eax, eax
  loc_00E06A96: fnclex
  loc_00E06A98: jge 00E06AACh
  loc_00E06A9A: push 0000009Ch
  loc_00E06A9F: push 006DCAD0h
  loc_00E06AA4: push esi
  loc_00E06AA5: push eax
  loc_00E06AA6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06AAC: lea ecx, var_18
  loc_00E06AAF: call ebx
  loc_00E06AB1: mov var_4, 00000000h
  loc_00E06AB8: push 00E06ACAh
  loc_00E06ABD: jmp 00E06AC9h
  loc_00E06ABF: lea ecx, var_18
  loc_00E06AC2: call [00401254h] ; __vbaFreeObj
  loc_00E06AC8: ret
  loc_00E06AC9: ret
  loc_00E06ACA: mov eax, Me
  loc_00E06ACD: push eax
  loc_00E06ACE: mov ecx, [eax]
  loc_00E06AD0: call [ecx+00000008h]
  loc_00E06AD3: mov eax, var_4
  loc_00E06AD6: mov ecx, var_14
  loc_00E06AD9: pop edi
  loc_00E06ADA: pop esi
  loc_00E06ADB: mov fs:[00000000h], ecx
  loc_00E06AE2: pop ebx
  loc_00E06AE3: mov esp, ebp
  loc_00E06AE5: pop ebp
  loc_00E06AE6: retn 0004h
End Sub

Private Sub jcbutton2_UnknownEvent_9 'E06680
  loc_00E06680: push ebp
  loc_00E06681: mov ebp, esp
  loc_00E06683: sub esp, 0000000Ch
  loc_00E06686: push 00402836h ; __vbaExceptHandler
  loc_00E0668B: mov eax, fs:[00000000h]
  loc_00E06691: push eax
  loc_00E06692: mov fs:[00000000h], esp
  loc_00E06699: sub esp, 00000030h
  loc_00E0669C: push ebx
  loc_00E0669D: push esi
  loc_00E0669E: push edi
  loc_00E0669F: mov var_C, esp
  loc_00E066A2: mov var_8, 00401FD0h
  loc_00E066A9: mov eax, Me
  loc_00E066AC: mov ecx, eax
  loc_00E066AE: and ecx, 00000001h
  loc_00E066B1: mov var_4, ecx
  loc_00E066B4: and al, FEh
  loc_00E066B6: push eax
  loc_00E066B7: mov Me, eax
  loc_00E066BA: mov edx, [eax]
  loc_00E066BC: call [edx+00000004h]
  loc_00E066BF: mov eax, [00E3D09Ch]
  loc_00E066C4: test eax, eax
  loc_00E066C6: jnz 00E066D8h
  loc_00E066C8: push 00E3D09Ch
  loc_00E066CD: push 006CECE4h
  loc_00E066D2: call [004011A0h] ; __vbaNew2
  loc_00E066D8: sub esp, 00000010h
  loc_00E066DB: mov ecx, 0000000Ah
  loc_00E066E0: mov ebx, esp
  loc_00E066E2: mov var_24, ecx
  loc_00E066E5: mov eax, 80020004h
  loc_00E066EA: sub esp, 00000010h
  loc_00E066ED: mov [ebx], ecx
  loc_00E066EF: mov ecx, var_30
  loc_00E066F2: mov edx, eax
  loc_00E066F4: mov esi, [00E3D09Ch]
  loc_00E066FA: mov [ebx+00000004h], ecx
  loc_00E066FD: mov ecx, esp
  loc_00E066FF: mov edi, [esi]
  loc_00E06701: push esi
  loc_00E06702: mov [ebx+00000008h], eax
  loc_00E06705: mov eax, var_28
  loc_00E06708: mov [ebx+0000000Ch], eax
  loc_00E0670B: mov eax, var_24
  loc_00E0670E: mov [ecx], eax
  loc_00E06710: mov eax, var_20
  loc_00E06713: mov [ecx+00000004h], eax
  loc_00E06716: mov [ecx+00000008h], edx
  loc_00E06719: mov edx, var_18
  loc_00E0671C: mov [ecx+0000000Ch], edx
  loc_00E0671F: call [edi+000002B0h]
  loc_00E06725: test eax, eax
  loc_00E06727: fnclex
  loc_00E06729: jge 00E0673Dh
  loc_00E0672B: push 000002B0h
  loc_00E06730: push 006DFEF8h
  loc_00E06735: push esi
  loc_00E06736: push eax
  loc_00E06737: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0673D: mov var_4, 00000000h
  loc_00E06744: mov eax, Me
  loc_00E06747: push eax
  loc_00E06748: mov ecx, [eax]
  loc_00E0674A: call [ecx+00000008h]
  loc_00E0674D: mov eax, var_4
  loc_00E06750: mov ecx, var_14
  loc_00E06753: pop edi
  loc_00E06754: pop esi
  loc_00E06755: mov fs:[00000000h], ecx
  loc_00E0675C: pop ebx
  loc_00E0675D: mov esp, ebp
  loc_00E0675F: pop ebp
  loc_00E06760: retn 0004h
End Sub

Private Sub Form_Load() 'E06140
  loc_00E06140: push ebp
  loc_00E06141: mov ebp, esp
  loc_00E06143: sub esp, 0000000Ch
  loc_00E06146: push 00402836h ; __vbaExceptHandler
  loc_00E0614B: mov eax, fs:[00000000h]
  loc_00E06151: push eax
  loc_00E06152: mov fs:[00000000h], esp
  loc_00E06159: sub esp, 00000054h
  loc_00E0615C: push ebx
  loc_00E0615D: push esi
  loc_00E0615E: push edi
  loc_00E0615F: mov var_C, esp
  loc_00E06162: mov var_8, 00401FB0h
  loc_00E06169: mov esi, Me
  loc_00E0616C: mov eax, esi
  loc_00E0616E: and eax, 00000001h
  loc_00E06171: mov var_4, eax
  loc_00E06174: and esi, FFFFFFFEh
  loc_00E06177: push esi
  loc_00E06178: mov Me, esi
  loc_00E0617B: mov ecx, [esi]
  loc_00E0617D: call [ecx+00000004h]
  loc_00E06180: mov eax, [00E3D9CCh]
  loc_00E06185: xor ebx, ebx
  loc_00E06187: cmp eax, ebx
  loc_00E06189: mov var_18, ebx
  loc_00E0618C: mov var_1C, ebx
  loc_00E0618F: mov var_20, ebx
  loc_00E06192: mov var_30, ebx
  loc_00E06195: jnz 00E061A7h
  loc_00E06197: push 00E3D9CCh
  loc_00E0619C: push 006DC960h
  loc_00E061A1: call [004011A0h] ; __vbaNew2
  loc_00E061A7: mov edi, [00E3D9CCh]
  loc_00E061AD: lea eax, var_1C
  loc_00E061B0: push eax
  loc_00E061B1: push edi
  loc_00E061B2: mov edx, [edi]
  loc_00E061B4: call [edx+00000014h]
  loc_00E061B7: cmp eax, ebx
  loc_00E061B9: fnclex
  loc_00E061BB: jge 00E061CCh
  loc_00E061BD: push 00000014h
  loc_00E061BF: push 006DC950h
  loc_00E061C4: push edi
  loc_00E061C5: push eax
  loc_00E061C6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E061CC: mov eax, var_1C
  loc_00E061CF: lea edx, var_18
  loc_00E061D2: push edx
  loc_00E061D3: push eax
  loc_00E061D4: mov ecx, [eax]
  loc_00E061D6: mov edi, eax
  loc_00E061D8: call [ecx+00000050h]
  loc_00E061DB: cmp eax, ebx
  loc_00E061DD: fnclex
  loc_00E061DF: jge 00E061F0h
  loc_00E061E1: push 00000050h
  loc_00E061E3: push 006DEB4Ch
  loc_00E061E8: push edi
  loc_00E061E9: push eax
  loc_00E061EA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E061F0: mov eax, var_18
  loc_00E061F3: push eax
  loc_00E061F4: push 006E0380h ; "\SDD.mp4 "
  loc_00E061F9: call [0040106Ch] ; __vbaStrCat
  loc_00E061FF: sub esp, 00000010h
  loc_00E06202: mov ecx, 00000008h
  loc_00E06207: mov edx, esp
  loc_00E06209: mov var_30, ecx
  loc_00E0620C: mov var_28, eax
  loc_00E0620F: push 00000001h
  loc_00E06211: mov [edx], ecx
  loc_00E06213: mov ecx, var_2C
  loc_00E06216: push esi
  loc_00E06217: mov [edx+00000004h], ecx
  loc_00E0621A: mov ecx, [esi]
  loc_00E0621C: mov [edx+00000008h], eax
  loc_00E0621F: mov eax, var_24
  loc_00E06222: mov [edx+0000000Ch], eax
  loc_00E06225: call [ecx+00000350h]
  loc_00E0622B: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E06231: lea edx, var_20
  loc_00E06234: push eax
  loc_00E06235: push edx
  loc_00E06236: call edi
  loc_00E06238: push eax
  loc_00E06239: call [00401238h] ; __vbaLateIdSt
  loc_00E0623F: lea ecx, var_18
  loc_00E06242: call [00401258h] ; __vbaFreeStr
  loc_00E06248: lea eax, var_20
  loc_00E0624B: lea ecx, var_1C
  loc_00E0624E: push eax
  loc_00E0624F: push ecx
  loc_00E06250: push 00000002h
  loc_00E06252: call [00401048h] ; __vbaFreeObjList
  loc_00E06258: add esp, 0000000Ch
  loc_00E0625B: lea ecx, var_30
  loc_00E0625E: call [00401024h] ; __vbaFreeVar
  loc_00E06264: mov edx, [esi]
  loc_00E06266: push esi
  loc_00E06267: call [edx+00000314h]
  loc_00E0626D: push eax
  loc_00E0626E: lea eax, var_1C
  loc_00E06271: push eax
  loc_00E06272: call edi
  loc_00E06274: mov ebx, eax
  loc_00E06276: push 00000000h
  loc_00E06278: push ebx
  loc_00E06279: mov ecx, [ebx]
  loc_00E0627B: call [ecx+0000009Ch]
  loc_00E06281: test eax, eax
  loc_00E06283: fnclex
  loc_00E06285: jge 00E06299h
  loc_00E06287: push 0000009Ch
  loc_00E0628C: push 006DCAD0h
  loc_00E06291: push ebx
  loc_00E06292: push eax
  loc_00E06293: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06299: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E0629F: lea ecx, var_1C
  loc_00E062A2: call ebx
  loc_00E062A4: sub esp, 00000010h
  loc_00E062A7: mov ecx, 0000000Bh
  loc_00E062AC: mov edx, esp
  loc_00E062AE: xor eax, eax
  loc_00E062B0: push 80010007h
  loc_00E062B5: push esi
  loc_00E062B6: mov [edx], ecx
  loc_00E062B8: mov ecx, var_3C
  loc_00E062BB: mov [edx+00000004h], ecx
  loc_00E062BE: mov ecx, [esi]
  loc_00E062C0: mov [edx+00000008h], eax
  loc_00E062C3: mov eax, var_34
  loc_00E062C6: mov [edx+0000000Ch], eax
  loc_00E062C9: call [ecx+00000318h]
  loc_00E062CF: lea edx, var_1C
  loc_00E062D2: push eax
  loc_00E062D3: push edx
  loc_00E062D4: call edi
  loc_00E062D6: push eax
  loc_00E062D7: call [00401238h] ; __vbaLateIdSt
  loc_00E062DD: lea ecx, var_1C
  loc_00E062E0: call ebx
  loc_00E062E2: sub esp, 00000010h
  loc_00E062E5: mov ecx, 0000000Bh
  loc_00E062EA: mov edx, esp
  loc_00E062EC: xor eax, eax
  loc_00E062EE: push 80010007h
  loc_00E062F3: push esi
  loc_00E062F4: mov [edx], ecx
  loc_00E062F6: mov ecx, var_3C
  loc_00E062F9: mov [edx+00000004h], ecx
  loc_00E062FC: mov ecx, [esi]
  loc_00E062FE: mov [edx+00000008h], eax
  loc_00E06301: mov eax, var_34
  loc_00E06304: mov [edx+0000000Ch], eax
  loc_00E06307: call [ecx+0000031Ch]
  loc_00E0630D: lea edx, var_1C
  loc_00E06310: push eax
  loc_00E06311: push edx
  loc_00E06312: call edi
  loc_00E06314: push eax
  loc_00E06315: call [00401238h] ; __vbaLateIdSt
  loc_00E0631B: lea ecx, var_1C
  loc_00E0631E: call ebx
  loc_00E06320: sub esp, 00000010h
  loc_00E06323: mov ecx, 0000000Bh
  loc_00E06328: mov edx, esp
  loc_00E0632A: xor eax, eax
  loc_00E0632C: push 80010007h
  loc_00E06331: push esi
  loc_00E06332: mov [edx], ecx
  loc_00E06334: mov ecx, var_3C
  loc_00E06337: mov [edx+00000004h], ecx
  loc_00E0633A: mov ecx, [esi]
  loc_00E0633C: mov [edx+00000008h], eax
  loc_00E0633F: mov eax, var_34
  loc_00E06342: mov [edx+0000000Ch], eax
  loc_00E06345: call [ecx+00000320h]
  loc_00E0634B: lea edx, var_1C
  loc_00E0634E: push eax
  loc_00E0634F: push edx
  loc_00E06350: call edi
  loc_00E06352: push eax
  loc_00E06353: call [00401238h] ; __vbaLateIdSt
  loc_00E06359: lea ecx, var_1C
  loc_00E0635C: call ebx
  loc_00E0635E: mov eax, [esi]
  loc_00E06360: push esi
  loc_00E06361: call [eax+00000300h]
  loc_00E06367: lea ecx, var_1C
  loc_00E0636A: push eax
  loc_00E0636B: push ecx
  loc_00E0636C: call edi
  loc_00E0636E: mov edi, eax
  loc_00E06370: push 00000000h
  loc_00E06372: push edi
  loc_00E06373: mov edx, [edi]
  loc_00E06375: call [edx+0000009Ch]
  loc_00E0637B: test eax, eax
  loc_00E0637D: fnclex
  loc_00E0637F: jge 00E06393h
  loc_00E06381: push 0000009Ch
  loc_00E06386: push 006DCAD0h
  loc_00E0638B: push edi
  loc_00E0638C: push eax
  loc_00E0638D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06393: lea ecx, var_1C
  loc_00E06396: call ebx
  loc_00E06398: mov eax, [esi]
  loc_00E0639A: push 41200000h
  loc_00E0639F: push esi
  loc_00E063A0: call [eax+0000008Ch]
  loc_00E063A6: test eax, eax
  loc_00E063A8: fnclex
  loc_00E063AA: jge 00E063BEh
  loc_00E063AC: push 0000008Ch
  loc_00E063B1: push 006DC970h
  loc_00E063B6: push esi
  loc_00E063B7: push eax
  loc_00E063B8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E063BE: mov var_4, 00000000h
  loc_00E063C5: fwait
  loc_00E063C6: push 00E063F4h
  loc_00E063CB: jmp 00E063F3h
  loc_00E063CD: lea ecx, var_18
  loc_00E063D0: call [00401258h] ; __vbaFreeStr
  loc_00E063D6: lea ecx, var_20
  loc_00E063D9: lea edx, var_1C
  loc_00E063DC: push ecx
  loc_00E063DD: push edx
  loc_00E063DE: push 00000002h
  loc_00E063E0: call [00401048h] ; __vbaFreeObjList
  loc_00E063E6: add esp, 0000000Ch
  loc_00E063E9: lea ecx, var_30
  loc_00E063EC: call [00401024h] ; __vbaFreeVar
  loc_00E063F2: ret
  loc_00E063F3: ret
  loc_00E063F4: mov eax, Me
  loc_00E063F7: push eax
  loc_00E063F8: mov ecx, [eax]
  loc_00E063FA: call [ecx+00000008h]
  loc_00E063FD: mov eax, var_4
  loc_00E06400: mov ecx, var_14
  loc_00E06403: pop edi
  loc_00E06404: pop esi
  loc_00E06405: mov fs:[00000000h], ecx
  loc_00E0640C: pop ebx
  loc_00E0640D: mov esp, ebp
  loc_00E0640F: pop ebp
  loc_00E06410: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E05EF0
  loc_00E05EF0: push ebp
  loc_00E05EF1: mov ebp, esp
  loc_00E05EF3: sub esp, 0000000Ch
  loc_00E05EF6: push 00402836h ; __vbaExceptHandler
  loc_00E05EFB: mov eax, fs:[00000000h]
  loc_00E05F01: push eax
  loc_00E05F02: mov fs:[00000000h], esp
  loc_00E05F09: sub esp, 0000005Ch
  loc_00E05F0C: push ebx
  loc_00E05F0D: push esi
  loc_00E05F0E: push edi
  loc_00E05F0F: mov var_C, esp
  loc_00E05F12: mov var_8, 00401FA0h
  loc_00E05F19: mov esi, Me
  loc_00E05F1C: mov eax, esi
  loc_00E05F1E: and eax, 00000001h
  loc_00E05F21: mov var_4, eax
  loc_00E05F24: and esi, FFFFFFFEh
  loc_00E05F27: push esi
  loc_00E05F28: mov Me, esi
  loc_00E05F2B: mov ecx, [esi]
  loc_00E05F2D: call [ecx+00000004h]
  loc_00E05F30: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05F36: xor eax, eax
  loc_00E05F38: mov var_18, eax
  loc_00E05F3B: mov var_4C, eax
  loc_00E05F3E: mov var_50, eax
  loc_00E05F41: mov edx, [esi]
  loc_00E05F43: lea eax, var_4C
  loc_00E05F46: push eax
  loc_00E05F47: push esi
  loc_00E05F48: call [edx+00000070h]
  loc_00E05F4B: test eax, eax
  loc_00E05F4D: fnclex
  loc_00E05F4F: jge 00E05F5Ch
  loc_00E05F51: push 00000070h
  loc_00E05F53: push 006DC970h
  loc_00E05F58: push esi
  loc_00E05F59: push eax
  loc_00E05F5A: call ebx
  loc_00E05F5C: fld real4 ptr var_4C
  loc_00E05F5F: fadd st0, real4 ptr [00401298h]
  loc_00E05F65: mov ecx, [esi]
  loc_00E05F67: push ecx
  loc_00E05F68: fnstsw ax
  loc_00E05F6A: test al, 0Dh
  loc_00E05F6C: jnz 00E06130h
  loc_00E05F72: fstp real4 ptr [esp]
  loc_00E05F75: push esi
  loc_00E05F76: call [ecx+00000074h]
  loc_00E05F79: test eax, eax
  loc_00E05F7B: fnclex
  loc_00E05F7D: jge 00E05F8Ah
  loc_00E05F7F: push 00000074h
  loc_00E05F81: push 006DC970h
  loc_00E05F86: push esi
  loc_00E05F87: push eax
  loc_00E05F88: call ebx
  loc_00E05F8A: mov edx, [esi]
  loc_00E05F8C: lea eax, var_4C
  loc_00E05F8F: push eax
  loc_00E05F90: push esi
  loc_00E05F91: call [edx+00000070h]
  loc_00E05F94: test eax, eax
  loc_00E05F96: fnclex
  loc_00E05F98: jge 00E05FA5h
  loc_00E05F9A: push 00000070h
  loc_00E05F9C: push 006DC970h
  loc_00E05FA1: push esi
  loc_00E05FA2: push eax
  loc_00E05FA3: call ebx
  loc_00E05FA5: mov ecx, [esi]
  loc_00E05FA7: lea edx, var_50
  loc_00E05FAA: push edx
  loc_00E05FAB: push esi
  loc_00E05FAC: call [ecx+00000078h]
  loc_00E05FAF: test eax, eax
  loc_00E05FB1: fnclex
  loc_00E05FB3: jge 00E05FC0h
  loc_00E05FB5: push 00000078h
  loc_00E05FB7: push 006DC970h
  loc_00E05FBC: push esi
  loc_00E05FBD: push eax
  loc_00E05FBE: call ebx
  loc_00E05FC0: sub esp, 00000010h
  loc_00E05FC3: mov ecx, 0000000Ah
  loc_00E05FC8: mov ebx, esp
  loc_00E05FCA: mov eax, 80020004h
  loc_00E05FCF: mov edx, eax
  loc_00E05FD1: sub esp, 00000010h
  loc_00E05FD4: mov [ebx], ecx
  loc_00E05FD6: mov ecx, var_44
  loc_00E05FD9: mov edi, [esi]
  loc_00E05FDB: mov [ebx+00000004h], ecx
  loc_00E05FDE: mov ecx, esp
  loc_00E05FE0: sub esp, 00000010h
  loc_00E05FE3: mov [ebx+00000008h], eax
  loc_00E05FE6: mov eax, var_3C
  loc_00E05FE9: mov [ebx+0000000Ch], eax
  loc_00E05FEC: mov eax, 0000000Ah
  loc_00E05FF1: mov [ecx], eax
  loc_00E05FF3: mov eax, var_34
  loc_00E05FF6: mov [ecx+00000004h], eax
  loc_00E05FF9: mov eax, 00000004h
  loc_00E05FFE: mov [ecx+00000008h], edx
  loc_00E06001: mov edx, var_2C
  loc_00E06004: mov [ecx+0000000Ch], edx
  loc_00E06007: mov edx, var_24
  loc_00E0600A: mov ecx, esp
  loc_00E0600C: mov [ecx], eax
  loc_00E0600E: mov eax, var_50
  loc_00E06011: mov [ecx+00000004h], edx
  loc_00E06014: mov [ecx+00000008h], eax
  loc_00E06017: mov eax, var_1C
  loc_00E0601A: mov [ecx+0000000Ch], eax
  loc_00E0601D: mov ecx, var_4C
  loc_00E06020: push ecx
  loc_00E06021: push esi
  loc_00E06022: call [edi+000002A4h]
  loc_00E06028: test eax, eax
  loc_00E0602A: fnclex
  loc_00E0602C: jge 00E06044h
  loc_00E0602E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06034: push 000002A4h
  loc_00E06039: push 006DC970h
  loc_00E0603E: push esi
  loc_00E0603F: push eax
  loc_00E06040: call ebx
  loc_00E06042: jmp 00E0604Ah
  loc_00E06044: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0604A: call [004010BCh] ; rtcDoEvents
  loc_00E06050: mov edx, [esi]
  loc_00E06052: lea eax, var_4C
  loc_00E06055: push eax
  loc_00E06056: push esi
  loc_00E06057: call [edx+00000070h]
  loc_00E0605A: test eax, eax
  loc_00E0605C: fnclex
  loc_00E0605E: jge 00E0606Bh
  loc_00E06060: push 00000070h
  loc_00E06062: push 006DC970h
  loc_00E06067: push esi
  loc_00E06068: push eax
  loc_00E06069: call ebx
  loc_00E0606B: mov eax, [00E3D9CCh]
  loc_00E06070: test eax, eax
  loc_00E06072: jnz 00E06084h
  loc_00E06074: push 00E3D9CCh
  loc_00E06079: push 006DC960h
  loc_00E0607E: call [004011A0h] ; __vbaNew2
  loc_00E06084: mov edi, [00E3D9CCh]
  loc_00E0608A: lea edx, var_18
  loc_00E0608D: push edx
  loc_00E0608E: push edi
  loc_00E0608F: mov ecx, [edi]
  loc_00E06091: call [ecx+00000018h]
  loc_00E06094: test eax, eax
  loc_00E06096: fnclex
  loc_00E06098: jge 00E060A5h
  loc_00E0609A: push 00000018h
  loc_00E0609C: push 006DC950h
  loc_00E060A1: push edi
  loc_00E060A2: push eax
  loc_00E060A3: call ebx
  loc_00E060A5: mov eax, var_18
  loc_00E060A8: lea edx, var_50
  loc_00E060AB: push edx
  loc_00E060AC: push eax
  loc_00E060AD: mov ecx, [eax]
  loc_00E060AF: mov edi, eax
  loc_00E060B1: call [ecx+00000098h]
  loc_00E060B7: test eax, eax
  loc_00E060B9: fnclex
  loc_00E060BB: jge 00E060CBh
  loc_00E060BD: push 00000098h
  loc_00E060C2: push 006DCAF0h
  loc_00E060C7: push edi
  loc_00E060C8: push eax
  loc_00E060C9: call ebx
  loc_00E060CB: fld real4 ptr var_4C
  loc_00E060CE: fcomp real4 ptr var_50
  loc_00E060D1: fnstsw ax
  loc_00E060D3: test ah, 41h
  loc_00E060D6: jz 00E060DFh
  loc_00E060D8: mov eax, 00000001h
  loc_00E060DD: jmp 00E060E1h
  loc_00E060DF: xor eax, eax
  loc_00E060E1: neg eax
  loc_00E060E3: lea ecx, var_18
  loc_00E060E6: mov edi, eax
  loc_00E060E8: call [00401254h] ; __vbaFreeObj
  loc_00E060EE: test di, di
  loc_00E060F1: jnz 00E05F41h
  loc_00E060F7: mov var_4, 00000000h
  loc_00E060FE: fwait
  loc_00E060FF: push 00E06111h
  loc_00E06104: jmp 00E06110h
  loc_00E06106: lea ecx, var_18
  loc_00E06109: call [00401254h] ; __vbaFreeObj
  loc_00E0610F: ret
  loc_00E06110: ret
  loc_00E06111: mov eax, Me
  loc_00E06114: push eax
  loc_00E06115: mov ecx, [eax]
  loc_00E06117: call [ecx+00000008h]
  loc_00E0611A: mov eax, var_4
  loc_00E0611D: mov ecx, var_14
  loc_00E06120: pop edi
  loc_00E06121: pop esi
  loc_00E06122: mov fs:[00000000h], ecx
  loc_00E06129: pop ebx
  loc_00E0612A: mov esp, ebp
  loc_00E0612C: pop ebp
  loc_00E0612D: retn 0008h
End Sub

Private Sub jcbutton3_UnknownEvent_9 'E06770
  loc_00E06770: push ebp
  loc_00E06771: mov ebp, esp
  loc_00E06773: sub esp, 0000000Ch
  loc_00E06776: push 00402836h ; __vbaExceptHandler
  loc_00E0677B: mov eax, fs:[00000000h]
  loc_00E06781: push eax
  loc_00E06782: mov fs:[00000000h], esp
  loc_00E06789: sub esp, 00000030h
  loc_00E0678C: push ebx
  loc_00E0678D: push esi
  loc_00E0678E: push edi
  loc_00E0678F: mov var_C, esp
  loc_00E06792: mov var_8, 00401FD8h
  loc_00E06799: mov eax, Me
  loc_00E0679C: mov ecx, eax
  loc_00E0679E: and ecx, 00000001h
  loc_00E067A1: mov var_4, ecx
  loc_00E067A4: and al, FEh
  loc_00E067A6: push eax
  loc_00E067A7: mov Me, eax
  loc_00E067AA: mov edx, [eax]
  loc_00E067AC: call [edx+00000004h]
  loc_00E067AF: mov eax, [00E3D0C4h]
  loc_00E067B4: test eax, eax
  loc_00E067B6: jnz 00E067C8h
  loc_00E067B8: push 00E3D0C4h
  loc_00E067BD: push 006CD7E4h
  loc_00E067C2: call [004011A0h] ; __vbaNew2
  loc_00E067C8: sub esp, 00000010h
  loc_00E067CB: mov ecx, 0000000Ah
  loc_00E067D0: mov ebx, esp
  loc_00E067D2: mov var_24, ecx
  loc_00E067D5: mov eax, 80020004h
  loc_00E067DA: sub esp, 00000010h
  loc_00E067DD: mov [ebx], ecx
  loc_00E067DF: mov ecx, var_30
  loc_00E067E2: mov edx, eax
  loc_00E067E4: mov esi, [00E3D0C4h]
  loc_00E067EA: mov [ebx+00000004h], ecx
  loc_00E067ED: mov ecx, esp
  loc_00E067EF: mov edi, [esi]
  loc_00E067F1: push esi
  loc_00E067F2: mov [ebx+00000008h], eax
  loc_00E067F5: mov eax, var_28
  loc_00E067F8: mov [ebx+0000000Ch], eax
  loc_00E067FB: mov eax, var_24
  loc_00E067FE: mov [ecx], eax
  loc_00E06800: mov eax, var_20
  loc_00E06803: mov [ecx+00000004h], eax
  loc_00E06806: mov [ecx+00000008h], edx
  loc_00E06809: mov edx, var_18
  loc_00E0680C: mov [ecx+0000000Ch], edx
  loc_00E0680F: call [edi+000002B0h]
  loc_00E06815: test eax, eax
  loc_00E06817: fnclex
  loc_00E06819: jge 00E0682Dh
  loc_00E0681B: push 000002B0h
  loc_00E06820: push 006DFFC8h
  loc_00E06825: push esi
  loc_00E06826: push eax
  loc_00E06827: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0682D: mov eax, [00E3D024h]
  loc_00E06832: test eax, eax
  loc_00E06834: jnz 00E06846h
  loc_00E06836: push 00E3D024h
  loc_00E0683B: push 006CE120h
  loc_00E06840: call [004011A0h] ; __vbaNew2
  loc_00E06846: mov esi, [00E3D024h]
  loc_00E0684C: push esi
  loc_00E0684D: mov eax, [esi]
  loc_00E0684F: call [eax+000002B4h]
  loc_00E06855: test eax, eax
  loc_00E06857: fnclex
  loc_00E06859: jge 00E0686Dh
  loc_00E0685B: push 000002B4h
  loc_00E06860: push 006DC970h
  loc_00E06865: push esi
  loc_00E06866: push eax
  loc_00E06867: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0686D: mov var_4, 00000000h
  loc_00E06874: mov eax, Me
  loc_00E06877: push eax
  loc_00E06878: mov ecx, [eax]
  loc_00E0687A: call [ecx+00000008h]
  loc_00E0687D: mov eax, var_4
  loc_00E06880: mov ecx, var_14
  loc_00E06883: pop edi
  loc_00E06884: pop esi
  loc_00E06885: mov fs:[00000000h], ecx
  loc_00E0688C: pop ebx
  loc_00E0688D: mov esp, ebp
  loc_00E0688F: pop ebp
  loc_00E06890: retn 0004h
End Sub

Private Sub Image1_Click() 'E06420
  loc_00E06420: push ebp
  loc_00E06421: mov ebp, esp
  loc_00E06423: sub esp, 0000000Ch
  loc_00E06426: push 00402836h ; __vbaExceptHandler
  loc_00E0642B: mov eax, fs:[00000000h]
  loc_00E06431: push eax
  loc_00E06432: mov fs:[00000000h], esp
  loc_00E06439: sub esp, 00000030h
  loc_00E0643C: push ebx
  loc_00E0643D: push esi
  loc_00E0643E: push edi
  loc_00E0643F: mov var_C, esp
  loc_00E06442: mov var_8, 00401FC0h
  loc_00E06449: mov eax, Me
  loc_00E0644C: mov ecx, eax
  loc_00E0644E: and ecx, 00000001h
  loc_00E06451: mov var_4, ecx
  loc_00E06454: and al, FEh
  loc_00E06456: push eax
  loc_00E06457: mov Me, eax
  loc_00E0645A: mov edx, [eax]
  loc_00E0645C: call [edx+00000004h]
  loc_00E0645F: mov eax, [00E3D024h]
  loc_00E06464: test eax, eax
  loc_00E06466: jnz 00E06478h
  loc_00E06468: push 00E3D024h
  loc_00E0646D: push 006CE120h
  loc_00E06472: call [004011A0h] ; __vbaNew2
  loc_00E06478: mov esi, [00E3D024h]
  loc_00E0647E: push esi
  loc_00E0647F: mov eax, [esi]
  loc_00E06481: call [eax+000002B4h]
  loc_00E06487: test eax, eax
  loc_00E06489: fnclex
  loc_00E0648B: jge 00E0649Fh
  loc_00E0648D: push 000002B4h
  loc_00E06492: push 006DC970h
  loc_00E06497: push esi
  loc_00E06498: push eax
  loc_00E06499: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0649F: mov eax, [00E3D088h]
  loc_00E064A4: test eax, eax
  loc_00E064A6: jnz 00E064B8h
  loc_00E064A8: push 00E3D088h
  loc_00E064AD: push 006CCEC4h
  loc_00E064B2: call [004011A0h] ; __vbaNew2
  loc_00E064B8: sub esp, 00000010h
  loc_00E064BB: mov ecx, 0000000Ah
  loc_00E064C0: mov ebx, esp
  loc_00E064C2: mov var_24, ecx
  loc_00E064C5: mov eax, 80020004h
  loc_00E064CA: sub esp, 00000010h
  loc_00E064CD: mov [ebx], ecx
  loc_00E064CF: mov ecx, var_30
  loc_00E064D2: mov edx, eax
  loc_00E064D4: mov esi, [00E3D088h]
  loc_00E064DA: mov [ebx+00000004h], ecx
  loc_00E064DD: mov ecx, esp
  loc_00E064DF: mov edi, [esi]
  loc_00E064E1: push esi
  loc_00E064E2: mov [ebx+00000008h], eax
  loc_00E064E5: mov eax, var_28
  loc_00E064E8: mov [ebx+0000000Ch], eax
  loc_00E064EB: mov eax, var_24
  loc_00E064EE: mov [ecx], eax
  loc_00E064F0: mov eax, var_20
  loc_00E064F3: mov [ecx+00000004h], eax
  loc_00E064F6: mov [ecx+00000008h], edx
  loc_00E064F9: mov edx, var_18
  loc_00E064FC: mov [ecx+0000000Ch], edx
  loc_00E064FF: call [edi+000002B0h]
  loc_00E06505: test eax, eax
  loc_00E06507: fnclex
  loc_00E06509: jge 00E0651Dh
  loc_00E0650B: push 000002B0h
  loc_00E06510: push 006DFE8Ch
  loc_00E06515: push esi
  loc_00E06516: push eax
  loc_00E06517: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0651D: mov var_4, 00000000h
  loc_00E06524: mov eax, Me
  loc_00E06527: push eax
  loc_00E06528: mov ecx, [eax]
  loc_00E0652A: call [ecx+00000008h]
  loc_00E0652D: mov eax, var_4
  loc_00E06530: mov ecx, var_14
  loc_00E06533: pop edi
  loc_00E06534: pop esi
  loc_00E06535: mov fs:[00000000h], ecx
  loc_00E0653C: pop ebx
  loc_00E0653D: mov esp, ebp
  loc_00E0653F: pop ebp
  loc_00E06540: retn 0004h
End Sub

Private Sub ext_Click() 'E055D0
  loc_00E055D0: push ebp
  loc_00E055D1: mov ebp, esp
  loc_00E055D3: sub esp, 0000000Ch
  loc_00E055D6: push 00402836h ; __vbaExceptHandler
  loc_00E055DB: mov eax, fs:[00000000h]
  loc_00E055E1: push eax
  loc_00E055E2: mov fs:[00000000h], esp
  loc_00E055E9: sub esp, 000000B8h
  loc_00E055EF: push ebx
  loc_00E055F0: push esi
  loc_00E055F1: push edi
  loc_00E055F2: mov var_C, esp
  loc_00E055F5: mov var_8, 00401F90h
  loc_00E055FC: mov ebx, Me
  loc_00E055FF: mov eax, ebx
  loc_00E05601: and eax, 00000001h
  loc_00E05604: mov var_4, eax
  loc_00E05607: and ebx, FFFFFFFEh
  loc_00E0560A: push ebx
  loc_00E0560B: mov Me, ebx
  loc_00E0560E: mov ecx, [ebx]
  loc_00E05610: call [ecx+00000004h]
  loc_00E05613: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E05619: xor eax, eax
  loc_00E0561B: mov ecx, 80020004h
  loc_00E05620: mov var_24, eax
  loc_00E05623: mov var_28, eax
  loc_00E05626: mov var_38, eax
  loc_00E05629: mov var_48, eax
  loc_00E0562C: mov var_58, eax
  loc_00E0562F: mov var_68, eax
  loc_00E05632: mov var_78, eax
  loc_00E05635: mov var_88, eax
  loc_00E0563B: mov var_B8, eax
  loc_00E05641: mov var_60, ecx
  loc_00E05644: mov eax, 0000000Ah
  loc_00E05649: mov var_50, ecx
  loc_00E0564C: mov edi, 00000008h
  loc_00E05651: lea edx, var_88
  loc_00E05657: lea ecx, var_48
  loc_00E0565A: mov var_68, eax
  loc_00E0565D: mov var_58, eax
  loc_00E05660: mov var_80, 006DC934h ; "Alert"
  loc_00E05667: mov var_88, edi
  loc_00E0566D: call __vbaVarDup
  loc_00E0566F: lea edx, var_78
  loc_00E05672: lea ecx, var_38
  loc_00E05675: mov var_70, 006DC8E0h ; "Apakah anda ingin keluar dari program?"
  loc_00E0567C: mov var_78, edi
  loc_00E0567F: call __vbaVarDup
  loc_00E05681: lea edx, var_68
  loc_00E05684: lea eax, var_58
  loc_00E05687: push edx
  loc_00E05688: lea ecx, var_48
  loc_00E0568B: push eax
  loc_00E0568C: push ecx
  loc_00E0568D: lea edx, var_38
  loc_00E05690: push 00000044h
  loc_00E05692: push edx
  loc_00E05693: call [004010A8h] ; rtcMsgBox
  loc_00E05699: lea edx, var_B8
  loc_00E0569F: lea ecx, var_24
  loc_00E056A2: mov var_B0, eax
  loc_00E056A8: mov var_B8, 00000003h
  loc_00E056B2: call [0040101Ch] ; __vbaVarMove
  loc_00E056B8: lea eax, var_68
  loc_00E056BB: lea ecx, var_58
  loc_00E056BE: push eax
  loc_00E056BF: lea edx, var_48
  loc_00E056C2: push ecx
  loc_00E056C3: lea eax, var_38
  loc_00E056C6: push edx
  loc_00E056C7: push eax
  loc_00E056C8: push 00000004h
  loc_00E056CA: call [00401038h] ; __vbaFreeVarList
  loc_00E056D0: add esp, 00000014h
  loc_00E056D3: lea ecx, var_24
  loc_00E056D6: lea edx, var_78
  loc_00E056D9: mov var_70, 00000006h
  loc_00E056E0: push ecx
  loc_00E056E1: push edx
  loc_00E056E2: mov var_78, 00008003h
  loc_00E056E9: call [00401108h] ; __vbaVarTstEq
  loc_00E056EF: test ax, ax
  loc_00E056F2: jz 00E05E90h
  loc_00E056F8: mov eax, [00E3D9CCh]
  loc_00E056FD: test eax, eax
  loc_00E056FF: jnz 00E05711h
  loc_00E05701: push 00E3D9CCh
  loc_00E05706: push 006DC960h
  loc_00E0570B: call [004011A0h] ; __vbaNew2
  loc_00E05711: mov eax, [00E3D088h]
  loc_00E05716: mov esi, [00E3D9CCh]
  loc_00E0571C: test eax, eax
  loc_00E0571E: jnz 00E05730h
  loc_00E05720: push 00E3D088h
  loc_00E05725: push 006CCEC4h
  loc_00E0572A: call [004011A0h] ; __vbaNew2
  loc_00E05730: mov eax, [00E3D088h]
  loc_00E05735: mov edi, [esi]
  loc_00E05737: lea ecx, var_28
  loc_00E0573A: push eax
  loc_00E0573B: push ecx
  loc_00E0573C: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05742: push eax
  loc_00E05743: push esi
  loc_00E05744: call [edi+00000010h]
  loc_00E05747: test eax, eax
  loc_00E05749: fnclex
  loc_00E0574B: jge 00E0575Ch
  loc_00E0574D: push 00000010h
  loc_00E0574F: push 006DC950h
  loc_00E05754: push esi
  loc_00E05755: push eax
  loc_00E05756: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0575C: lea ecx, var_28
  loc_00E0575F: call [00401254h] ; __vbaFreeObj
  loc_00E05765: mov eax, [00E3D9CCh]
  loc_00E0576A: test eax, eax
  loc_00E0576C: jnz 00E0577Eh
  loc_00E0576E: push 00E3D9CCh
  loc_00E05773: push 006DC960h
  loc_00E05778: call [004011A0h] ; __vbaNew2
  loc_00E0577E: mov eax, [00E3D010h]
  loc_00E05783: mov esi, [00E3D9CCh]
  loc_00E05789: test eax, eax
  loc_00E0578B: jnz 00E0579Dh
  loc_00E0578D: push 00E3D010h
  loc_00E05792: push 006D058Ch
  loc_00E05797: call [004011A0h] ; __vbaNew2
  loc_00E0579D: mov edx, [00E3D010h]
  loc_00E057A3: mov edi, [esi]
  loc_00E057A5: lea eax, var_28
  loc_00E057A8: push edx
  loc_00E057A9: push eax
  loc_00E057AA: call [004010B4h] ; __vbaObjSetAddref
  loc_00E057B0: push eax
  loc_00E057B1: push esi
  loc_00E057B2: call [edi+00000010h]
  loc_00E057B5: test eax, eax
  loc_00E057B7: fnclex
  loc_00E057B9: jge 00E057CAh
  loc_00E057BB: push 00000010h
  loc_00E057BD: push 006DC950h
  loc_00E057C2: push esi
  loc_00E057C3: push eax
  loc_00E057C4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E057CA: lea ecx, var_28
  loc_00E057CD: call [00401254h] ; __vbaFreeObj
  loc_00E057D3: mov eax, [00E3D9CCh]
  loc_00E057D8: test eax, eax
  loc_00E057DA: jnz 00E057F0h
  loc_00E057DC: mov edi, [004011A0h] ; __vbaNew2
  loc_00E057E2: push 00E3D9CCh
  loc_00E057E7: push 006DC960h
  loc_00E057EC: call edi
  loc_00E057EE: jmp 00E057F6h
  loc_00E057F0: mov edi, [004011A0h] ; __vbaNew2
  loc_00E057F6: mov esi, [00E3D9CCh]
  loc_00E057FC: lea ecx, var_28
  loc_00E057FF: push ebx
  loc_00E05800: push ecx
  loc_00E05801: mov edx, [esi]
  loc_00E05803: mov var_CC, edx
  loc_00E05809: call [004010B4h] ; __vbaObjSetAddref
  loc_00E0580F: mov edx, var_CC
  loc_00E05815: push eax
  loc_00E05816: push esi
  loc_00E05817: call [edx+00000010h]
  loc_00E0581A: test eax, eax
  loc_00E0581C: fnclex
  loc_00E0581E: jge 00E0582Fh
  loc_00E05820: push 00000010h
  loc_00E05822: push 006DC950h
  loc_00E05827: push esi
  loc_00E05828: push eax
  loc_00E05829: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0582F: lea ecx, var_28
  loc_00E05832: call [00401254h] ; __vbaFreeObj
  loc_00E05838: mov eax, [00E3D9CCh]
  loc_00E0583D: test eax, eax
  loc_00E0583F: jnz 00E0584Dh
  loc_00E05841: push 00E3D9CCh
  loc_00E05846: push 006DC960h
  loc_00E0584B: call edi
  loc_00E0584D: mov eax, [00E3D060h]
  loc_00E05852: mov esi, [00E3D9CCh]
  loc_00E05858: test eax, eax
  loc_00E0585A: jnz 00E05868h
  loc_00E0585C: push 00E3D060h
  loc_00E05861: push 006D87C0h
  loc_00E05866: call edi
  loc_00E05868: mov eax, [00E3D060h]
  loc_00E0586D: mov ebx, [esi]
  loc_00E0586F: lea ecx, var_28
  loc_00E05872: push eax
  loc_00E05873: push ecx
  loc_00E05874: call [004010B4h] ; __vbaObjSetAddref
  loc_00E0587A: push eax
  loc_00E0587B: push esi
  loc_00E0587C: call [ebx+00000010h]
  loc_00E0587F: test eax, eax
  loc_00E05881: fnclex
  loc_00E05883: jge 00E05894h
  loc_00E05885: push 00000010h
  loc_00E05887: push 006DC950h
  loc_00E0588C: push esi
  loc_00E0588D: push eax
  loc_00E0588E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05894: lea ecx, var_28
  loc_00E05897: call [00401254h] ; __vbaFreeObj
  loc_00E0589D: mov eax, [00E3D9CCh]
  loc_00E058A2: test eax, eax
  loc_00E058A4: jnz 00E058B2h
  loc_00E058A6: push 00E3D9CCh
  loc_00E058AB: push 006DC960h
  loc_00E058B0: call edi
  loc_00E058B2: mov eax, [00E3D09Ch]
  loc_00E058B7: mov esi, [00E3D9CCh]
  loc_00E058BD: test eax, eax
  loc_00E058BF: jnz 00E058CDh
  loc_00E058C1: push 00E3D09Ch
  loc_00E058C6: push 006CECE4h
  loc_00E058CB: call edi
  loc_00E058CD: mov edx, [00E3D09Ch]
  loc_00E058D3: mov ebx, [esi]
  loc_00E058D5: lea eax, var_28
  loc_00E058D8: push edx
  loc_00E058D9: push eax
  loc_00E058DA: call [004010B4h] ; __vbaObjSetAddref
  loc_00E058E0: push eax
  loc_00E058E1: push esi
  loc_00E058E2: call [ebx+00000010h]
  loc_00E058E5: test eax, eax
  loc_00E058E7: fnclex
  loc_00E058E9: jge 00E058FAh
  loc_00E058EB: push 00000010h
  loc_00E058ED: push 006DC950h
  loc_00E058F2: push esi
  loc_00E058F3: push eax
  loc_00E058F4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E058FA: lea ecx, var_28
  loc_00E058FD: call [00401254h] ; __vbaFreeObj
  loc_00E05903: mov eax, [00E3D9CCh]
  loc_00E05908: test eax, eax
  loc_00E0590A: jnz 00E05918h
  loc_00E0590C: push 00E3D9CCh
  loc_00E05911: push 006DC960h
  loc_00E05916: call edi
  loc_00E05918: mov eax, [00E3D0B0h]
  loc_00E0591D: mov esi, [00E3D9CCh]
  loc_00E05923: test eax, eax
  loc_00E05925: jnz 00E05933h
  loc_00E05927: push 00E3D0B0h
  loc_00E0592C: push 006CF8E0h
  loc_00E05931: call edi
  loc_00E05933: mov ecx, [00E3D0B0h]
  loc_00E05939: mov ebx, [esi]
  loc_00E0593B: lea edx, var_28
  loc_00E0593E: push ecx
  loc_00E0593F: push edx
  loc_00E05940: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05946: push eax
  loc_00E05947: push esi
  loc_00E05948: call [ebx+00000010h]
  loc_00E0594B: test eax, eax
  loc_00E0594D: fnclex
  loc_00E0594F: jge 00E05960h
  loc_00E05951: push 00000010h
  loc_00E05953: push 006DC950h
  loc_00E05958: push esi
  loc_00E05959: push eax
  loc_00E0595A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05960: lea ecx, var_28
  loc_00E05963: call [00401254h] ; __vbaFreeObj
  loc_00E05969: mov eax, [00E3D9CCh]
  loc_00E0596E: test eax, eax
  loc_00E05970: jnz 00E0597Eh
  loc_00E05972: push 00E3D9CCh
  loc_00E05977: push 006DC960h
  loc_00E0597C: call edi
  loc_00E0597E: mov eax, [00E3D0C4h]
  loc_00E05983: mov esi, [00E3D9CCh]
  loc_00E05989: test eax, eax
  loc_00E0598B: jnz 00E05999h
  loc_00E0598D: push 00E3D0C4h
  loc_00E05992: push 006CD7E4h
  loc_00E05997: call edi
  loc_00E05999: mov eax, [00E3D0C4h]
  loc_00E0599E: mov ebx, [esi]
  loc_00E059A0: lea ecx, var_28
  loc_00E059A3: push eax
  loc_00E059A4: push ecx
  loc_00E059A5: call [004010B4h] ; __vbaObjSetAddref
  loc_00E059AB: push eax
  loc_00E059AC: push esi
  loc_00E059AD: call [ebx+00000010h]
  loc_00E059B0: test eax, eax
  loc_00E059B2: fnclex
  loc_00E059B4: jge 00E059C5h
  loc_00E059B6: push 00000010h
  loc_00E059B8: push 006DC950h
  loc_00E059BD: push esi
  loc_00E059BE: push eax
  loc_00E059BF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E059C5: lea ecx, var_28
  loc_00E059C8: call [00401254h] ; __vbaFreeObj
  loc_00E059CE: mov eax, [00E3D9CCh]
  loc_00E059D3: test eax, eax
  loc_00E059D5: jnz 00E059E3h
  loc_00E059D7: push 00E3D9CCh
  loc_00E059DC: push 006DC960h
  loc_00E059E1: call edi
  loc_00E059E3: mov eax, [00E3D074h]
  loc_00E059E8: mov esi, [00E3D9CCh]
  loc_00E059EE: test eax, eax
  loc_00E059F0: jnz 00E059FEh
  loc_00E059F2: push 00E3D074h
  loc_00E059F7: push 006D57D0h
  loc_00E059FC: call edi
  loc_00E059FE: mov edx, [00E3D074h]
  loc_00E05A04: mov ebx, [esi]
  loc_00E05A06: lea eax, var_28
  loc_00E05A09: push edx
  loc_00E05A0A: push eax
  loc_00E05A0B: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05A11: push eax
  loc_00E05A12: push esi
  loc_00E05A13: call [ebx+00000010h]
  loc_00E05A16: test eax, eax
  loc_00E05A18: fnclex
  loc_00E05A1A: jge 00E05A2Bh
  loc_00E05A1C: push 00000010h
  loc_00E05A1E: push 006DC950h
  loc_00E05A23: push esi
  loc_00E05A24: push eax
  loc_00E05A25: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05A2B: lea ecx, var_28
  loc_00E05A2E: call [00401254h] ; __vbaFreeObj
  loc_00E05A34: mov eax, [00E3D9CCh]
  loc_00E05A39: test eax, eax
  loc_00E05A3B: jnz 00E05A49h
  loc_00E05A3D: push 00E3D9CCh
  loc_00E05A42: push 006DC960h
  loc_00E05A47: call edi
  loc_00E05A49: mov eax, [00E3D010h]
  loc_00E05A4E: mov esi, [00E3D9CCh]
  loc_00E05A54: test eax, eax
  loc_00E05A56: jnz 00E05A64h
  loc_00E05A58: push 00E3D010h
  loc_00E05A5D: push 006D058Ch
  loc_00E05A62: call edi
  loc_00E05A64: mov ecx, [00E3D010h]
  loc_00E05A6A: mov ebx, [esi]
  loc_00E05A6C: lea edx, var_28
  loc_00E05A6F: push ecx
  loc_00E05A70: push edx
  loc_00E05A71: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05A77: push eax
  loc_00E05A78: push esi
  loc_00E05A79: call [ebx+00000010h]
  loc_00E05A7C: test eax, eax
  loc_00E05A7E: fnclex
  loc_00E05A80: jge 00E05A91h
  loc_00E05A82: push 00000010h
  loc_00E05A84: push 006DC950h
  loc_00E05A89: push esi
  loc_00E05A8A: push eax
  loc_00E05A8B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05A91: lea ecx, var_28
  loc_00E05A94: call [00401254h] ; __vbaFreeObj
  loc_00E05A9A: mov eax, [00E3D9CCh]
  loc_00E05A9F: test eax, eax
  loc_00E05AA1: jnz 00E05AAFh
  loc_00E05AA3: push 00E3D9CCh
  loc_00E05AA8: push 006DC960h
  loc_00E05AAD: call edi
  loc_00E05AAF: mov eax, [00E3D024h]
  loc_00E05AB4: mov esi, [00E3D9CCh]
  loc_00E05ABA: test eax, eax
  loc_00E05ABC: jnz 00E05ACAh
  loc_00E05ABE: push 00E3D024h
  loc_00E05AC3: push 006CE120h
  loc_00E05AC8: call edi
  loc_00E05ACA: mov eax, [00E3D024h]
  loc_00E05ACF: mov ebx, [esi]
  loc_00E05AD1: lea ecx, var_28
  loc_00E05AD4: push eax
  loc_00E05AD5: push ecx
  loc_00E05AD6: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05ADC: push eax
  loc_00E05ADD: push esi
  loc_00E05ADE: call [ebx+00000010h]
  loc_00E05AE1: test eax, eax
  loc_00E05AE3: fnclex
  loc_00E05AE5: jge 00E05AF6h
  loc_00E05AE7: push 00000010h
  loc_00E05AE9: push 006DC950h
  loc_00E05AEE: push esi
  loc_00E05AEF: push eax
  loc_00E05AF0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05AF6: lea ecx, var_28
  loc_00E05AF9: call [00401254h] ; __vbaFreeObj
  loc_00E05AFF: mov eax, [00E3D9CCh]
  loc_00E05B04: test eax, eax
  loc_00E05B06: jnz 00E05B14h
  loc_00E05B08: push 00E3D9CCh
  loc_00E05B0D: push 006DC960h
  loc_00E05B12: call edi
  loc_00E05B14: mov eax, [00E3D0D8h]
  loc_00E05B19: mov esi, [00E3D9CCh]
  loc_00E05B1F: test eax, eax
  loc_00E05B21: jnz 00E05B2Fh
  loc_00E05B23: push 00E3D0D8h
  loc_00E05B28: push 006D300Ch
  loc_00E05B2D: call edi
  loc_00E05B2F: mov edx, [00E3D0D8h]
  loc_00E05B35: mov ebx, [esi]
  loc_00E05B37: lea eax, var_28
  loc_00E05B3A: push edx
  loc_00E05B3B: push eax
  loc_00E05B3C: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05B42: push eax
  loc_00E05B43: push esi
  loc_00E05B44: call [ebx+00000010h]
  loc_00E05B47: test eax, eax
  loc_00E05B49: fnclex
  loc_00E05B4B: jge 00E05B5Ch
  loc_00E05B4D: push 00000010h
  loc_00E05B4F: push 006DC950h
  loc_00E05B54: push esi
  loc_00E05B55: push eax
  loc_00E05B56: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05B5C: lea ecx, var_28
  loc_00E05B5F: call [00401254h] ; __vbaFreeObj
  loc_00E05B65: mov eax, [00E3D9CCh]
  loc_00E05B6A: test eax, eax
  loc_00E05B6C: jnz 00E05B7Ah
  loc_00E05B6E: push 00E3D9CCh
  loc_00E05B73: push 006DC960h
  loc_00E05B78: call edi
  loc_00E05B7A: mov eax, [00E3D0ECh]
  loc_00E05B7F: mov esi, [00E3D9CCh]
  loc_00E05B85: test eax, eax
  loc_00E05B87: jnz 00E05B95h
  loc_00E05B89: push 00E3D0ECh
  loc_00E05B8E: push 006CC808h
  loc_00E05B93: call edi
  loc_00E05B95: mov ecx, [00E3D0ECh]
  loc_00E05B9B: mov ebx, [esi]
  loc_00E05B9D: lea edx, var_28
  loc_00E05BA0: push ecx
  loc_00E05BA1: push edx
  loc_00E05BA2: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05BA8: push eax
  loc_00E05BA9: push esi
  loc_00E05BAA: call [ebx+00000010h]
  loc_00E05BAD: test eax, eax
  loc_00E05BAF: fnclex
  loc_00E05BB1: jge 00E05BC2h
  loc_00E05BB3: push 00000010h
  loc_00E05BB5: push 006DC950h
  loc_00E05BBA: push esi
  loc_00E05BBB: push eax
  loc_00E05BBC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05BC2: lea ecx, var_28
  loc_00E05BC5: call [00401254h] ; __vbaFreeObj
  loc_00E05BCB: mov eax, [00E3D9CCh]
  loc_00E05BD0: test eax, eax
  loc_00E05BD2: jnz 00E05BE0h
  loc_00E05BD4: push 00E3D9CCh
  loc_00E05BD9: push 006DC960h
  loc_00E05BDE: call edi
  loc_00E05BE0: mov eax, [00E3D100h]
  loc_00E05BE5: mov esi, [00E3D9CCh]
  loc_00E05BEB: test eax, eax
  loc_00E05BED: jnz 00E05BFBh
  loc_00E05BEF: push 00E3D100h
  loc_00E05BF4: push 006CC14Ch
  loc_00E05BF9: call edi
  loc_00E05BFB: mov eax, [00E3D100h]
  loc_00E05C00: mov ebx, [esi]
  loc_00E05C02: lea ecx, var_28
  loc_00E05C05: push eax
  loc_00E05C06: push ecx
  loc_00E05C07: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05C0D: push eax
  loc_00E05C0E: push esi
  loc_00E05C0F: call [ebx+00000010h]
  loc_00E05C12: test eax, eax
  loc_00E05C14: fnclex
  loc_00E05C16: jge 00E05C27h
  loc_00E05C18: push 00000010h
  loc_00E05C1A: push 006DC950h
  loc_00E05C1F: push esi
  loc_00E05C20: push eax
  loc_00E05C21: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05C27: lea ecx, var_28
  loc_00E05C2A: call [00401254h] ; __vbaFreeObj
  loc_00E05C30: mov eax, [00E3D9CCh]
  loc_00E05C35: test eax, eax
  loc_00E05C37: jnz 00E05C45h
  loc_00E05C39: push 00E3D9CCh
  loc_00E05C3E: push 006DC960h
  loc_00E05C43: call edi
  loc_00E05C45: mov eax, [00E3D114h]
  loc_00E05C4A: mov esi, [00E3D9CCh]
  loc_00E05C50: test eax, eax
  loc_00E05C52: jnz 00E05C60h
  loc_00E05C54: push 00E3D114h
  loc_00E05C59: push 006CB640h
  loc_00E05C5E: call edi
  loc_00E05C60: mov edx, [00E3D114h]
  loc_00E05C66: mov ebx, [esi]
  loc_00E05C68: lea eax, var_28
  loc_00E05C6B: push edx
  loc_00E05C6C: push eax
  loc_00E05C6D: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05C73: push eax
  loc_00E05C74: push esi
  loc_00E05C75: call [ebx+00000010h]
  loc_00E05C78: test eax, eax
  loc_00E05C7A: fnclex
  loc_00E05C7C: jge 00E05C8Dh
  loc_00E05C7E: push 00000010h
  loc_00E05C80: push 006DC950h
  loc_00E05C85: push esi
  loc_00E05C86: push eax
  loc_00E05C87: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05C8D: lea ecx, var_28
  loc_00E05C90: call [00401254h] ; __vbaFreeObj
  loc_00E05C96: mov eax, [00E3D9CCh]
  loc_00E05C9B: test eax, eax
  loc_00E05C9D: jnz 00E05CABh
  loc_00E05C9F: push 00E3D9CCh
  loc_00E05CA4: push 006DC960h
  loc_00E05CA9: call edi
  loc_00E05CAB: mov eax, [00E3D128h]
  loc_00E05CB0: mov esi, [00E3D9CCh]
  loc_00E05CB6: test eax, eax
  loc_00E05CB8: jnz 00E05CC6h
  loc_00E05CBA: push 00E3D128h
  loc_00E05CBF: push 006CB548h
  loc_00E05CC4: call edi
  loc_00E05CC6: mov ecx, [00E3D128h]
  loc_00E05CCC: mov ebx, [esi]
  loc_00E05CCE: lea edx, var_28
  loc_00E05CD1: push ecx
  loc_00E05CD2: push edx
  loc_00E05CD3: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05CD9: push eax
  loc_00E05CDA: push esi
  loc_00E05CDB: call [ebx+00000010h]
  loc_00E05CDE: test eax, eax
  loc_00E05CE0: fnclex
  loc_00E05CE2: jge 00E05CF3h
  loc_00E05CE4: push 00000010h
  loc_00E05CE6: push 006DC950h
  loc_00E05CEB: push esi
  loc_00E05CEC: push eax
  loc_00E05CED: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05CF3: lea ecx, var_28
  loc_00E05CF6: call [00401254h] ; __vbaFreeObj
  loc_00E05CFC: mov eax, [00E3D9CCh]
  loc_00E05D01: test eax, eax
  loc_00E05D03: jnz 00E05D11h
  loc_00E05D05: push 00E3D9CCh
  loc_00E05D0A: push 006DC960h
  loc_00E05D0F: call edi
  loc_00E05D11: mov eax, [00E3D13Ch]
  loc_00E05D16: mov esi, [00E3D9CCh]
  loc_00E05D1C: test eax, eax
  loc_00E05D1E: jnz 00E05D2Ch
  loc_00E05D20: push 00E3D13Ch
  loc_00E05D25: push 006CB450h
  loc_00E05D2A: call edi
  loc_00E05D2C: mov eax, [00E3D13Ch]
  loc_00E05D31: mov ebx, [esi]
  loc_00E05D33: lea ecx, var_28
  loc_00E05D36: push eax
  loc_00E05D37: push ecx
  loc_00E05D38: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05D3E: push eax
  loc_00E05D3F: push esi
  loc_00E05D40: call [ebx+00000010h]
  loc_00E05D43: test eax, eax
  loc_00E05D45: fnclex
  loc_00E05D47: jge 00E05D58h
  loc_00E05D49: push 00000010h
  loc_00E05D4B: push 006DC950h
  loc_00E05D50: push esi
  loc_00E05D51: push eax
  loc_00E05D52: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05D58: lea ecx, var_28
  loc_00E05D5B: call [00401254h] ; __vbaFreeObj
  loc_00E05D61: mov eax, [00E3D9CCh]
  loc_00E05D66: test eax, eax
  loc_00E05D68: jnz 00E05D76h
  loc_00E05D6A: push 00E3D9CCh
  loc_00E05D6F: push 006DC960h
  loc_00E05D74: call edi
  loc_00E05D76: mov eax, [00E3D150h]
  loc_00E05D7B: mov esi, [00E3D9CCh]
  loc_00E05D81: test eax, eax
  loc_00E05D83: jnz 00E05D91h
  loc_00E05D85: push 00E3D150h
  loc_00E05D8A: push 006CB358h
  loc_00E05D8F: call edi
  loc_00E05D91: mov edx, [00E3D150h]
  loc_00E05D97: mov ebx, [esi]
  loc_00E05D99: lea eax, var_28
  loc_00E05D9C: push edx
  loc_00E05D9D: push eax
  loc_00E05D9E: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05DA4: push eax
  loc_00E05DA5: push esi
  loc_00E05DA6: call [ebx+00000010h]
  loc_00E05DA9: test eax, eax
  loc_00E05DAB: fnclex
  loc_00E05DAD: jge 00E05DBEh
  loc_00E05DAF: push 00000010h
  loc_00E05DB1: push 006DC950h
  loc_00E05DB6: push esi
  loc_00E05DB7: push eax
  loc_00E05DB8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05DBE: lea ecx, var_28
  loc_00E05DC1: call [00401254h] ; __vbaFreeObj
  loc_00E05DC7: mov eax, [00E3D9CCh]
  loc_00E05DCC: test eax, eax
  loc_00E05DCE: jnz 00E05DDCh
  loc_00E05DD0: push 00E3D9CCh
  loc_00E05DD5: push 006DC960h
  loc_00E05DDA: call edi
  loc_00E05DDC: mov eax, [00E3D038h]
  loc_00E05DE1: mov esi, [00E3D9CCh]
  loc_00E05DE7: test eax, eax
  loc_00E05DE9: jnz 00E05DF7h
  loc_00E05DEB: push 00E3D038h
  loc_00E05DF0: push 006CB260h
  loc_00E05DF5: call edi
  loc_00E05DF7: mov ecx, [00E3D038h]
  loc_00E05DFD: mov ebx, [esi]
  loc_00E05DFF: lea edx, var_28
  loc_00E05E02: push ecx
  loc_00E05E03: push edx
  loc_00E05E04: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05E0A: push eax
  loc_00E05E0B: push esi
  loc_00E05E0C: call [ebx+00000010h]
  loc_00E05E0F: test eax, eax
  loc_00E05E11: fnclex
  loc_00E05E13: jge 00E05E24h
  loc_00E05E15: push 00000010h
  loc_00E05E17: push 006DC950h
  loc_00E05E1C: push esi
  loc_00E05E1D: push eax
  loc_00E05E1E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05E24: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E05E2A: lea ecx, var_28
  loc_00E05E2D: call ebx
  loc_00E05E2F: mov eax, [00E3D9CCh]
  loc_00E05E34: test eax, eax
  loc_00E05E36: jnz 00E05E44h
  loc_00E05E38: push 00E3D9CCh
  loc_00E05E3D: push 006DC960h
  loc_00E05E42: call edi
  loc_00E05E44: mov eax, [00E3D164h]
  loc_00E05E49: mov esi, [00E3D9CCh]
  loc_00E05E4F: test eax, eax
  loc_00E05E51: jnz 00E05E5Fh
  loc_00E05E53: push 00E3D164h
  loc_00E05E58: push 006CB168h
  loc_00E05E5D: call edi
  loc_00E05E5F: mov eax, [00E3D164h]
  loc_00E05E64: mov edi, [esi]
  loc_00E05E66: lea ecx, var_28
  loc_00E05E69: push eax
  loc_00E05E6A: push ecx
  loc_00E05E6B: call [004010B4h] ; __vbaObjSetAddref
  loc_00E05E71: push eax
  loc_00E05E72: push esi
  loc_00E05E73: call [edi+00000010h]
  loc_00E05E76: test eax, eax
  loc_00E05E78: fnclex
  loc_00E05E7A: jge 00E05E8Bh
  loc_00E05E7C: push 00000010h
  loc_00E05E7E: push 006DC950h
  loc_00E05E83: push esi
  loc_00E05E84: push eax
  loc_00E05E85: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05E8B: lea ecx, var_28
  loc_00E05E8E: call ebx
  loc_00E05E90: mov var_4, 00000000h
  loc_00E05E97: push 00E05ECDh
  loc_00E05E9C: jmp 00E05EC3h
  loc_00E05E9E: lea ecx, var_28
  loc_00E05EA1: call [00401254h] ; __vbaFreeObj
  loc_00E05EA7: lea edx, var_68
  loc_00E05EAA: lea eax, var_58
  loc_00E05EAD: push edx
  loc_00E05EAE: lea ecx, var_48
  loc_00E05EB1: push eax
  loc_00E05EB2: lea edx, var_38
  loc_00E05EB5: push ecx
  loc_00E05EB6: push edx
  loc_00E05EB7: push 00000004h
  loc_00E05EB9: call [00401038h] ; __vbaFreeVarList
  loc_00E05EBF: add esp, 00000014h
  loc_00E05EC2: ret
  loc_00E05EC3: lea ecx, var_24
  loc_00E05EC6: call [00401024h] ; __vbaFreeVar
  loc_00E05ECC: ret
  loc_00E05ECD: mov eax, Me
  loc_00E05ED0: push eax
  loc_00E05ED1: mov ecx, [eax]
  loc_00E05ED3: call [ecx+00000008h]
  loc_00E05ED6: mov eax, var_4
  loc_00E05ED9: mov ecx, var_14
  loc_00E05EDC: pop edi
  loc_00E05EDD: pop esi
  loc_00E05EDE: mov fs:[00000000h], ecx
  loc_00E05EE5: pop ebx
  loc_00E05EE6: mov esp, ebp
  loc_00E05EE8: pop ebp
  loc_00E05EE9: retn 0004h
End Sub

Private Sub menus_UnknownEvent_9 'E06AF0
  loc_00E06AF0: push ebp
  loc_00E06AF1: mov ebp, esp
  loc_00E06AF3: sub esp, 0000000Ch
  loc_00E06AF6: push 00402836h ; __vbaExceptHandler
  loc_00E06AFB: mov eax, fs:[00000000h]
  loc_00E06B01: push eax
  loc_00E06B02: mov fs:[00000000h], esp
  loc_00E06B09: sub esp, 00000044h
  loc_00E06B0C: push ebx
  loc_00E06B0D: push esi
  loc_00E06B0E: push edi
  loc_00E06B0F: mov var_C, esp
  loc_00E06B12: mov var_8, 00401FF0h
  loc_00E06B19: mov esi, Me
  loc_00E06B1C: mov eax, esi
  loc_00E06B1E: and eax, 00000001h
  loc_00E06B21: mov var_4, eax
  loc_00E06B24: and esi, FFFFFFFEh
  loc_00E06B27: push esi
  loc_00E06B28: mov Me, esi
  loc_00E06B2B: mov ecx, [esi]
  loc_00E06B2D: call [ecx+00000004h]
  loc_00E06B30: mov edx, [esi]
  loc_00E06B32: xor eax, eax
  loc_00E06B34: push eax
  loc_00E06B35: push FFFFFE0Bh
  loc_00E06B3A: push esi
  loc_00E06B3B: mov var_18, eax
  loc_00E06B3E: mov var_28, eax
  loc_00E06B41: call [edx+00000330h]
  loc_00E06B47: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E06B4D: push eax
  loc_00E06B4E: lea eax, var_18
  loc_00E06B51: push eax
  loc_00E06B52: call edi
  loc_00E06B54: lea ecx, var_28
  loc_00E06B57: push eax
  loc_00E06B58: push ecx
  loc_00E06B59: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E06B5F: add esp, 00000010h
  loc_00E06B62: push eax
  loc_00E06B63: call [004011D0h] ; __vbaI4Var
  loc_00E06B69: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E06B6F: xor edx, edx
  loc_00E06B71: cmp eax, 00404040h
  loc_00E06B76: lea ecx, var_18
  loc_00E06B79: setz dl
  loc_00E06B7C: neg edx
  loc_00E06B7E: mov var_4C, dx
  loc_00E06B82: call ebx
  loc_00E06B84: lea ecx, var_28
  loc_00E06B87: call [00401024h] ; __vbaFreeVar
  loc_00E06B8D: cmp var_4C, 0000h
  loc_00E06B92: jz 00E06C99h
  loc_00E06B98: sub esp, 00000010h
  loc_00E06B9B: mov ecx, 00000003h
  loc_00E06BA0: mov edx, esp
  loc_00E06BA2: mov eax, 00008000h
  loc_00E06BA7: push FFFFFE0Bh
  loc_00E06BAC: push esi
  loc_00E06BAD: mov [edx], ecx
  loc_00E06BAF: mov ecx, var_34
  loc_00E06BB2: mov [edx+00000004h], ecx
  loc_00E06BB5: mov ecx, [esi]
  loc_00E06BB7: mov [edx+00000008h], eax
  loc_00E06BBA: mov eax, var_2C
  loc_00E06BBD: mov [edx+0000000Ch], eax
  loc_00E06BC0: call [ecx+00000330h]
  loc_00E06BC6: lea edx, var_18
  loc_00E06BC9: push eax
  loc_00E06BCA: push edx
  loc_00E06BCB: call edi
  loc_00E06BCD: push eax
  loc_00E06BCE: call [00401238h] ; __vbaLateIdSt
  loc_00E06BD4: lea ecx, var_18
  loc_00E06BD7: call ebx
  loc_00E06BD9: sub esp, 00000010h
  loc_00E06BDC: mov ecx, 00000003h
  loc_00E06BE1: mov edx, esp
  loc_00E06BE3: mov eax, 0080FF80h
  loc_00E06BE8: push FFFFFDFFh
  loc_00E06BED: push esi
  loc_00E06BEE: mov [edx], ecx
  loc_00E06BF0: mov ecx, var_34
  loc_00E06BF3: mov [edx+00000004h], ecx
  loc_00E06BF6: mov ecx, [esi]
  loc_00E06BF8: mov [edx+00000008h], eax
  loc_00E06BFB: mov eax, var_2C
  loc_00E06BFE: mov [edx+0000000Ch], eax
  loc_00E06C01: call [ecx+00000330h]
  loc_00E06C07: lea edx, var_18
  loc_00E06C0A: push eax
  loc_00E06C0B: push edx
  loc_00E06C0C: call edi
  loc_00E06C0E: push eax
  loc_00E06C0F: call [00401238h] ; __vbaLateIdSt
  loc_00E06C15: lea ecx, var_18
  loc_00E06C18: call ebx
  loc_00E06C1A: sub esp, 00000010h
  loc_00E06C1D: mov ecx, 00000008h
  loc_00E06C22: mov edx, esp
  loc_00E06C24: mov eax, 006E0398h ; "SHOW"
  loc_00E06C29: push FFFFFDFAh
  loc_00E06C2E: push esi
  loc_00E06C2F: mov [edx], ecx
  loc_00E06C31: mov ecx, var_34
  loc_00E06C34: mov [edx+00000004h], ecx
  loc_00E06C37: mov ecx, [esi]
  loc_00E06C39: mov [edx+00000008h], eax
  loc_00E06C3C: mov eax, var_2C
  loc_00E06C3F: mov [edx+0000000Ch], eax
  loc_00E06C42: call [ecx+00000330h]
  loc_00E06C48: lea edx, var_18
  loc_00E06C4B: push eax
  loc_00E06C4C: push edx
  loc_00E06C4D: call edi
  loc_00E06C4F: push eax
  loc_00E06C50: call [00401238h] ; __vbaLateIdSt
  loc_00E06C56: lea ecx, var_18
  loc_00E06C59: call ebx
  loc_00E06C5B: mov eax, [esi]
  loc_00E06C5D: push esi
  loc_00E06C5E: call [eax+00000314h]
  loc_00E06C64: lea ecx, var_18
  loc_00E06C67: push eax
  loc_00E06C68: push ecx
  loc_00E06C69: call edi
  loc_00E06C6B: mov esi, eax
  loc_00E06C6D: push 00000000h
  loc_00E06C6F: push esi
  loc_00E06C70: mov edx, [esi]
  loc_00E06C72: call [edx+0000009Ch]
  loc_00E06C78: test eax, eax
  loc_00E06C7A: fnclex
  loc_00E06C7C: jge 00E06D95h
  loc_00E06C82: push 0000009Ch
  loc_00E06C87: push 006DCAD0h
  loc_00E06C8C: push esi
  loc_00E06C8D: push eax
  loc_00E06C8E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06C94: jmp 00E06D95h
  loc_00E06C99: mov eax, [esi]
  loc_00E06C9B: push esi
  loc_00E06C9C: call [eax+00000314h]
  loc_00E06CA2: lea ecx, var_18
  loc_00E06CA5: push eax
  loc_00E06CA6: push ecx
  loc_00E06CA7: call edi
  loc_00E06CA9: mov edx, [eax]
  loc_00E06CAB: push FFFFFFFFh
  loc_00E06CAD: push eax
  loc_00E06CAE: mov var_4C, eax
  loc_00E06CB1: call [edx+0000009Ch]
  loc_00E06CB7: test eax, eax
  loc_00E06CB9: fnclex
  loc_00E06CBB: jge 00E06CD2h
  loc_00E06CBD: mov ecx, var_4C
  loc_00E06CC0: push 0000009Ch
  loc_00E06CC5: push 006DCAD0h
  loc_00E06CCA: push ecx
  loc_00E06CCB: push eax
  loc_00E06CCC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06CD2: lea ecx, var_18
  loc_00E06CD5: call ebx
  loc_00E06CD7: sub esp, 00000010h
  loc_00E06CDA: mov ecx, 00000003h
  loc_00E06CDF: mov edx, esp
  loc_00E06CE1: mov eax, 00404040h
  loc_00E06CE6: push FFFFFE0Bh
  loc_00E06CEB: push esi
  loc_00E06CEC: mov [edx], ecx
  loc_00E06CEE: mov ecx, var_34
  loc_00E06CF1: mov [edx+00000004h], ecx
  loc_00E06CF4: mov ecx, [esi]
  loc_00E06CF6: mov [edx+00000008h], eax
  loc_00E06CF9: mov eax, var_2C
  loc_00E06CFC: mov [edx+0000000Ch], eax
  loc_00E06CFF: call [ecx+00000330h]
  loc_00E06D05: lea edx, var_18
  loc_00E06D08: push eax
  loc_00E06D09: push edx
  loc_00E06D0A: call edi
  loc_00E06D0C: push eax
  loc_00E06D0D: call [00401238h] ; __vbaLateIdSt
  loc_00E06D13: lea ecx, var_18
  loc_00E06D16: call ebx
  loc_00E06D18: sub esp, 00000010h
  loc_00E06D1B: mov ecx, 00000003h
  loc_00E06D20: mov edx, esp
  loc_00E06D22: mov eax, 00E0E0E0h
  loc_00E06D27: push FFFFFDFFh
  loc_00E06D2C: push esi
  loc_00E06D2D: mov [edx], ecx
  loc_00E06D2F: mov ecx, var_34
  loc_00E06D32: mov [edx+00000004h], ecx
  loc_00E06D35: mov ecx, [esi]
  loc_00E06D37: mov [edx+00000008h], eax
  loc_00E06D3A: mov eax, var_2C
  loc_00E06D3D: mov [edx+0000000Ch], eax
  loc_00E06D40: call [ecx+00000330h]
  loc_00E06D46: lea edx, var_18
  loc_00E06D49: push eax
  loc_00E06D4A: push edx
  loc_00E06D4B: call edi
  loc_00E06D4D: push eax
  loc_00E06D4E: call [00401238h] ; __vbaLateIdSt
  loc_00E06D54: lea ecx, var_18
  loc_00E06D57: call ebx
  loc_00E06D59: sub esp, 00000010h
  loc_00E06D5C: mov ecx, 00000008h
  loc_00E06D61: mov edx, esp
  loc_00E06D63: mov eax, 006E03A8h ; "CLOSE"
  loc_00E06D68: push FFFFFDFAh
  loc_00E06D6D: push esi
  loc_00E06D6E: mov [edx], ecx
  loc_00E06D70: mov ecx, var_34
  loc_00E06D73: mov [edx+00000004h], ecx
  loc_00E06D76: mov ecx, [esi]
  loc_00E06D78: mov [edx+00000008h], eax
  loc_00E06D7B: mov eax, var_2C
  loc_00E06D7E: mov [edx+0000000Ch], eax
  loc_00E06D81: call [ecx+00000330h]
  loc_00E06D87: lea edx, var_18
  loc_00E06D8A: push eax
  loc_00E06D8B: push edx
  loc_00E06D8C: call edi
  loc_00E06D8E: push eax
  loc_00E06D8F: call [00401238h] ; __vbaLateIdSt
  loc_00E06D95: lea ecx, var_18
  loc_00E06D98: call ebx
  loc_00E06D9A: mov var_4, 00000000h
  loc_00E06DA1: push 00E06DBCh
  loc_00E06DA6: jmp 00E06DBBh
  loc_00E06DA8: lea ecx, var_18
  loc_00E06DAB: call [00401254h] ; __vbaFreeObj
  loc_00E06DB1: lea ecx, var_28
  loc_00E06DB4: call [00401024h] ; __vbaFreeVar
  loc_00E06DBA: ret
  loc_00E06DBB: ret
  loc_00E06DBC: mov eax, Me
  loc_00E06DBF: push eax
  loc_00E06DC0: mov ecx, [eax]
  loc_00E06DC2: call [ecx+00000008h]
  loc_00E06DC5: mov eax, var_4
  loc_00E06DC8: mov ecx, var_14
  loc_00E06DCB: pop edi
  loc_00E06DCC: pop esi
  loc_00E06DCD: mov fs:[00000000h], ecx
  loc_00E06DD4: pop ebx
  loc_00E06DD5: mov esp, ebp
  loc_00E06DD7: pop ebp
  loc_00E06DD8: retn 0004h
End Sub

Private Sub daftar_UnknownEvent_9 'E054A0
  loc_00E054A0: push ebp
  loc_00E054A1: mov ebp, esp
  loc_00E054A3: sub esp, 0000000Ch
  loc_00E054A6: push 00402836h ; __vbaExceptHandler
  loc_00E054AB: mov eax, fs:[00000000h]
  loc_00E054B1: push eax
  loc_00E054B2: mov fs:[00000000h], esp
  loc_00E054B9: sub esp, 00000030h
  loc_00E054BC: push ebx
  loc_00E054BD: push esi
  loc_00E054BE: push edi
  loc_00E054BF: mov var_C, esp
  loc_00E054C2: mov var_8, 00401F88h
  loc_00E054C9: mov eax, Me
  loc_00E054CC: mov ecx, eax
  loc_00E054CE: and ecx, 00000001h
  loc_00E054D1: mov var_4, ecx
  loc_00E054D4: and al, FEh
  loc_00E054D6: push eax
  loc_00E054D7: mov Me, eax
  loc_00E054DA: mov edx, [eax]
  loc_00E054DC: call [edx+00000004h]
  loc_00E054DF: mov eax, [00E3D074h]
  loc_00E054E4: test eax, eax
  loc_00E054E6: jnz 00E054F8h
  loc_00E054E8: push 00E3D074h
  loc_00E054ED: push 006D57D0h
  loc_00E054F2: call [004011A0h] ; __vbaNew2
  loc_00E054F8: sub esp, 00000010h
  loc_00E054FB: mov ecx, 0000000Ah
  loc_00E05500: mov ebx, esp
  loc_00E05502: mov var_24, ecx
  loc_00E05505: mov eax, 80020004h
  loc_00E0550A: sub esp, 00000010h
  loc_00E0550D: mov [ebx], ecx
  loc_00E0550F: mov ecx, var_30
  loc_00E05512: mov edx, eax
  loc_00E05514: mov esi, [00E3D074h]
  loc_00E0551A: mov [ebx+00000004h], ecx
  loc_00E0551D: mov ecx, esp
  loc_00E0551F: mov edi, [esi]
  loc_00E05521: push esi
  loc_00E05522: mov [ebx+00000008h], eax
  loc_00E05525: mov eax, var_28
  loc_00E05528: mov [ebx+0000000Ch], eax
  loc_00E0552B: mov eax, var_24
  loc_00E0552E: mov [ecx], eax
  loc_00E05530: mov eax, var_20
  loc_00E05533: mov [ecx+00000004h], eax
  loc_00E05536: mov [ecx+00000008h], edx
  loc_00E05539: mov edx, var_18
  loc_00E0553C: mov [ecx+0000000Ch], edx
  loc_00E0553F: call [edi+000002B0h]
  loc_00E05545: test eax, eax
  loc_00E05547: fnclex
  loc_00E05549: jge 00E0555Dh
  loc_00E0554B: push 000002B0h
  loc_00E05550: push 006DFCB8h
  loc_00E05555: push esi
  loc_00E05556: push eax
  loc_00E05557: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0555D: mov eax, [00E3D024h]
  loc_00E05562: test eax, eax
  loc_00E05564: jnz 00E05576h
  loc_00E05566: push 00E3D024h
  loc_00E0556B: push 006CE120h
  loc_00E05570: call [004011A0h] ; __vbaNew2
  loc_00E05576: mov esi, [00E3D024h]
  loc_00E0557C: push esi
  loc_00E0557D: mov eax, [esi]
  loc_00E0557F: call [eax+000002B4h]
  loc_00E05585: test eax, eax
  loc_00E05587: fnclex
  loc_00E05589: jge 00E0559Dh
  loc_00E0558B: push 000002B4h
  loc_00E05590: push 006DC970h
  loc_00E05595: push esi
  loc_00E05596: push eax
  loc_00E05597: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0559D: mov var_4, 00000000h
  loc_00E055A4: mov eax, Me
  loc_00E055A7: push eax
  loc_00E055A8: mov ecx, [eax]
  loc_00E055AA: call [ecx+00000008h]
  loc_00E055AD: mov eax, var_4
  loc_00E055B0: mov ecx, var_14
  loc_00E055B3: pop edi
  loc_00E055B4: pop esi
  loc_00E055B5: mov fs:[00000000h], ecx
  loc_00E055BC: pop ebx
  loc_00E055BD: mov esp, ebp
  loc_00E055BF: pop ebp
  loc_00E055C0: retn 0004h
End Sub

Private Sub biaya_UnknownEvent_9 'E05040
  loc_00E05040: push ebp
  loc_00E05041: mov ebp, esp
  loc_00E05043: sub esp, 0000000Ch
  loc_00E05046: push 00402836h ; __vbaExceptHandler
  loc_00E0504B: mov eax, fs:[00000000h]
  loc_00E05051: push eax
  loc_00E05052: mov fs:[00000000h], esp
  loc_00E05059: sub esp, 00000030h
  loc_00E0505C: push ebx
  loc_00E0505D: push esi
  loc_00E0505E: push edi
  loc_00E0505F: mov var_C, esp
  loc_00E05062: mov var_8, 00401F70h
  loc_00E05069: mov eax, Me
  loc_00E0506C: mov ecx, eax
  loc_00E0506E: and ecx, 00000001h
  loc_00E05071: mov var_4, ecx
  loc_00E05074: and al, FEh
  loc_00E05076: push eax
  loc_00E05077: mov Me, eax
  loc_00E0507A: mov edx, [eax]
  loc_00E0507C: call [edx+00000004h]
  loc_00E0507F: mov eax, [00E3D024h]
  loc_00E05084: test eax, eax
  loc_00E05086: jnz 00E05098h
  loc_00E05088: push 00E3D024h
  loc_00E0508D: push 006CE120h
  loc_00E05092: call [004011A0h] ; __vbaNew2
  loc_00E05098: mov esi, [00E3D024h]
  loc_00E0509E: push esi
  loc_00E0509F: mov eax, [esi]
  loc_00E050A1: call [eax+000002B4h]
  loc_00E050A7: test eax, eax
  loc_00E050A9: fnclex
  loc_00E050AB: jge 00E050BFh
  loc_00E050AD: push 000002B4h
  loc_00E050B2: push 006DC970h
  loc_00E050B7: push esi
  loc_00E050B8: push eax
  loc_00E050B9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E050BF: mov eax, [00E3D060h]
  loc_00E050C4: test eax, eax
  loc_00E050C6: jnz 00E050D8h
  loc_00E050C8: push 00E3D060h
  loc_00E050CD: push 006D87C0h
  loc_00E050D2: call [004011A0h] ; __vbaNew2
  loc_00E050D8: sub esp, 00000010h
  loc_00E050DB: mov ecx, 0000000Ah
  loc_00E050E0: mov ebx, esp
  loc_00E050E2: mov var_24, ecx
  loc_00E050E5: mov eax, 80020004h
  loc_00E050EA: sub esp, 00000010h
  loc_00E050ED: mov [ebx], ecx
  loc_00E050EF: mov ecx, var_30
  loc_00E050F2: mov edx, eax
  loc_00E050F4: mov esi, [00E3D060h]
  loc_00E050FA: mov [ebx+00000004h], ecx
  loc_00E050FD: mov ecx, esp
  loc_00E050FF: mov edi, [esi]
  loc_00E05101: push esi
  loc_00E05102: mov [ebx+00000008h], eax
  loc_00E05105: mov eax, var_28
  loc_00E05108: mov [ebx+0000000Ch], eax
  loc_00E0510B: mov eax, var_24
  loc_00E0510E: mov [ecx], eax
  loc_00E05110: mov eax, var_20
  loc_00E05113: mov [ecx+00000004h], eax
  loc_00E05116: mov [ecx+00000008h], edx
  loc_00E05119: mov edx, var_18
  loc_00E0511C: mov [ecx+0000000Ch], edx
  loc_00E0511F: call [edi+000002B0h]
  loc_00E05125: test eax, eax
  loc_00E05127: fnclex
  loc_00E05129: jge 00E0513Dh
  loc_00E0512B: push 000002B0h
  loc_00E05130: push 006DFA4Ch
  loc_00E05135: push esi
  loc_00E05136: push eax
  loc_00E05137: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0513D: mov var_4, 00000000h
  loc_00E05144: mov eax, Me
  loc_00E05147: push eax
  loc_00E05148: mov ecx, [eax]
  loc_00E0514A: call [ecx+00000008h]
  loc_00E0514D: mov eax, var_4
  loc_00E05150: mov ecx, var_14
  loc_00E05153: pop edi
  loc_00E05154: pop esi
  loc_00E05155: mov fs:[00000000h], ecx
  loc_00E0515C: pop ebx
  loc_00E0515D: mov esp, ebp
  loc_00E0515F: pop ebp
  loc_00E05160: retn 0004h
End Sub

Private Sub Timer1_Timer() 'E06DE0
  loc_00E06DE0: push ebp
  loc_00E06DE1: mov ebp, esp
  loc_00E06DE3: sub esp, 0000000Ch
  loc_00E06DE6: push 00402836h ; __vbaExceptHandler
  loc_00E06DEB: mov eax, fs:[00000000h]
  loc_00E06DF1: push eax
  loc_00E06DF2: mov fs:[00000000h], esp
  loc_00E06DF9: sub esp, 00000028h
  loc_00E06DFC: push ebx
  loc_00E06DFD: push esi
  loc_00E06DFE: push edi
  loc_00E06DFF: mov var_C, esp
  loc_00E06E02: mov var_8, 00402000h
  loc_00E06E09: mov esi, Me
  loc_00E06E0C: mov eax, esi
  loc_00E06E0E: and eax, 00000001h
  loc_00E06E11: mov var_4, eax
  loc_00E06E14: and esi, FFFFFFFEh
  loc_00E06E17: push esi
  loc_00E06E18: mov Me, esi
  loc_00E06E1B: mov ecx, [esi]
  loc_00E06E1D: call [ecx+00000004h]
  loc_00E06E20: mov edx, [esi]
  loc_00E06E22: xor edi, edi
  loc_00E06E24: push 006E03B4h
  loc_00E06E29: push edi
  loc_00E06E2A: push 00000004h
  loc_00E06E2C: push esi
  loc_00E06E2D: mov var_18, edi
  loc_00E06E30: mov var_1C, edi
  loc_00E06E33: mov var_2C, edi
  loc_00E06E36: call [edx+00000350h]
  loc_00E06E3C: mov esi, [004010ACh] ; __vbaObjSet
  loc_00E06E42: push eax
  loc_00E06E43: lea eax, var_18
  loc_00E06E46: push eax
  loc_00E06E47: call __vbaObjSet
  loc_00E06E49: lea ecx, var_2C
  loc_00E06E4C: push eax
  loc_00E06E4D: push ecx
  loc_00E06E4E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E06E54: add esp, 00000010h
  loc_00E06E57: push eax
  loc_00E06E58: call [00401120h] ; __vbaCastObjVar
  loc_00E06E5E: lea edx, var_1C
  loc_00E06E61: push eax
  loc_00E06E62: push edx
  loc_00E06E63: call __vbaObjSet
  loc_00E06E65: mov esi, eax
  loc_00E06E67: push esi
  loc_00E06E68: mov eax, [esi]
  loc_00E06E6A: call [eax+00000020h]
  loc_00E06E6D: cmp eax, edi
  loc_00E06E6F: fnclex
  loc_00E06E71: jge 00E06E82h
  loc_00E06E73: push 00000020h
  loc_00E06E75: push 006E03B4h
  loc_00E06E7A: push esi
  loc_00E06E7B: push eax
  loc_00E06E7C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06E82: lea ecx, var_1C
  loc_00E06E85: lea edx, var_18
  loc_00E06E88: push ecx
  loc_00E06E89: push edx
  loc_00E06E8A: push 00000002h
  loc_00E06E8C: call [00401048h] ; __vbaFreeObjList
  loc_00E06E92: add esp, 0000000Ch
  loc_00E06E95: lea ecx, var_2C
  loc_00E06E98: call [00401024h] ; __vbaFreeVar
  loc_00E06E9E: mov var_4, edi
  loc_00E06EA1: push 00E06EC6h
  loc_00E06EA6: jmp 00E06EC5h
  loc_00E06EA8: lea eax, var_1C
  loc_00E06EAB: lea ecx, var_18
  loc_00E06EAE: push eax
  loc_00E06EAF: push ecx
  loc_00E06EB0: push 00000002h
  loc_00E06EB2: call [00401048h] ; __vbaFreeObjList
  loc_00E06EB8: add esp, 0000000Ch
  loc_00E06EBB: lea ecx, var_2C
  loc_00E06EBE: call [00401024h] ; __vbaFreeVar
  loc_00E06EC4: ret
  loc_00E06EC5: ret
  loc_00E06EC6: mov eax, Me
  loc_00E06EC9: push eax
  loc_00E06ECA: mov edx, [eax]
  loc_00E06ECC: call [edx+00000008h]
  loc_00E06ECF: mov eax, var_4
  loc_00E06ED2: mov ecx, var_14
  loc_00E06ED5: pop edi
  loc_00E06ED6: pop esi
  loc_00E06ED7: mov fs:[00000000h], ecx
  loc_00E06EDE: pop ebx
  loc_00E06EDF: mov esp, ebp
  loc_00E06EE1: pop ebp
  loc_00E06EE2: retn 0004h
End Sub

Private Sub Timer4_Timer() 'E06EF0
  loc_00E06EF0: push ebp
  loc_00E06EF1: mov ebp, esp
  loc_00E06EF3: sub esp, 0000000Ch
  loc_00E06EF6: push 00402836h ; __vbaExceptHandler
  loc_00E06EFB: mov eax, fs:[00000000h]
  loc_00E06F01: push eax
  loc_00E06F02: mov fs:[00000000h], esp
  loc_00E06F09: sub esp, 00000018h
  loc_00E06F0C: push ebx
  loc_00E06F0D: push esi
  loc_00E06F0E: push edi
  loc_00E06F0F: mov var_C, esp
  loc_00E06F12: mov var_8, 00402018h
  loc_00E06F19: mov esi, Me
  loc_00E06F1C: mov eax, esi
  loc_00E06F1E: and eax, 00000001h
  loc_00E06F21: mov var_4, eax
  loc_00E06F24: and esi, FFFFFFFEh
  loc_00E06F27: push esi
  loc_00E06F28: mov Me, esi
  loc_00E06F2B: mov ecx, [esi]
  loc_00E06F2D: call [ecx+00000004h]
  loc_00E06F30: mov edx, [esi]
  loc_00E06F32: lea eax, var_1C
  loc_00E06F35: xor ebx, ebx
  loc_00E06F37: push eax
  loc_00E06F38: push esi
  loc_00E06F39: mov var_18, ebx
  loc_00E06F3C: mov var_1C, ebx
  loc_00E06F3F: call [edx+00000088h]
  loc_00E06F45: cmp eax, ebx
  loc_00E06F47: fnclex
  loc_00E06F49: jge 00E06F61h
  loc_00E06F4B: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06F51: push 00000088h
  loc_00E06F56: push 006DC970h
  loc_00E06F5B: push esi
  loc_00E06F5C: push eax
  loc_00E06F5D: call edi
  loc_00E06F5F: jmp 00E06F67h
  loc_00E06F61: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E06F67: fld real4 ptr var_1C
  loc_00E06F6A: fadd st0, real4 ptr [004012E0h]
  loc_00E06F70: mov ecx, [esi]
  loc_00E06F72: push ecx
  loc_00E06F73: fnstsw ax
  loc_00E06F75: test al, 0Dh
  loc_00E06F77: jnz 00E07073h
  loc_00E06F7D: fstp real4 ptr [esp]
  loc_00E06F80: push esi
  loc_00E06F81: call [ecx+0000008Ch]
  loc_00E06F87: cmp eax, ebx
  loc_00E06F89: fnclex
  loc_00E06F8B: jge 00E06F9Bh
  loc_00E06F8D: push 0000008Ch
  loc_00E06F92: push 006DC970h
  loc_00E06F97: push esi
  loc_00E06F98: push eax
  loc_00E06F99: call edi
  loc_00E06F9B: mov edx, [esi]
  loc_00E06F9D: push esi
  loc_00E06F9E: call [edx+000006F8h]
  loc_00E06FA4: cmp eax, ebx
  loc_00E06FA6: jge 00E06FB6h
  loc_00E06FA8: push 000006F8h
  loc_00E06FAD: push 006DC9A4h
  loc_00E06FB2: push esi
  loc_00E06FB3: push eax
  loc_00E06FB4: call edi
  loc_00E06FB6: mov eax, [esi]
  loc_00E06FB8: lea ecx, var_1C
  loc_00E06FBB: push ecx
  loc_00E06FBC: push esi
  loc_00E06FBD: call [eax+00000088h]
  loc_00E06FC3: cmp eax, ebx
  loc_00E06FC5: fnclex
  loc_00E06FC7: jge 00E06FD7h
  loc_00E06FC9: push 00000088h
  loc_00E06FCE: push 006DC970h
  loc_00E06FD3: push esi
  loc_00E06FD4: push eax
  loc_00E06FD5: call edi
  loc_00E06FD7: fld real4 ptr var_1C
  loc_00E06FDA: fcomp real4 ptr [00402010h]
  loc_00E06FE0: fnstsw ax
  loc_00E06FE2: test ah, 01h
  loc_00E06FE5: jnz 00E0703Ah
  loc_00E06FE7: mov edx, [esi]
  loc_00E06FE9: push esi
  loc_00E06FEA: call [edx+000002FCh]
  loc_00E06FF0: push eax
  loc_00E06FF1: lea eax, var_18
  loc_00E06FF4: push eax
  loc_00E06FF5: call [004010ACh] ; __vbaObjSet
  loc_00E06FFB: mov ebx, eax
  loc_00E06FFD: push 00000000h
  loc_00E06FFF: push ebx
  loc_00E07000: mov ecx, [ebx]
  loc_00E07002: call [ecx+0000005Ch]
  loc_00E07005: test eax, eax
  loc_00E07007: fnclex
  loc_00E07009: jge 00E07016h
  loc_00E0700B: push 0000005Ch
  loc_00E0700D: push 006DCAE0h
  loc_00E07012: push ebx
  loc_00E07013: push eax
  loc_00E07014: call edi
  loc_00E07016: lea ecx, var_18
  loc_00E07019: call [00401254h] ; __vbaFreeObj
  loc_00E0701F: mov edx, [esi]
  loc_00E07021: push esi
  loc_00E07022: call [edx+000006F8h]
  loc_00E07028: test eax, eax
  loc_00E0702A: jge 00E0703Ah
  loc_00E0702C: push 000006F8h
  loc_00E07031: push 006DC9A4h
  loc_00E07036: push esi
  loc_00E07037: push eax
  loc_00E07038: call edi
  loc_00E0703A: mov var_4, 00000000h
  loc_00E07041: fwait
  loc_00E07042: push 00E07054h
  loc_00E07047: jmp 00E07053h
  loc_00E07049: lea ecx, var_18
  loc_00E0704C: call [00401254h] ; __vbaFreeObj
  loc_00E07052: ret
  loc_00E07053: ret
  loc_00E07054: mov eax, Me
  loc_00E07057: push eax
  loc_00E07058: mov ecx, [eax]
  loc_00E0705A: call [ecx+00000008h]
  loc_00E0705D: mov eax, var_4
  loc_00E07060: mov ecx, var_14
  loc_00E07063: pop edi
  loc_00E07064: pop esi
  loc_00E07065: mov fs:[00000000h], ecx
  loc_00E0706C: pop ebx
  loc_00E0706D: mov esp, ebp
  loc_00E0706F: pop ebp
  loc_00E07070: retn 0004h
End Sub

Private Sub jcbutton1_UnknownEvent_9 'E06550
  loc_00E06550: push ebp
  loc_00E06551: mov ebp, esp
  loc_00E06553: sub esp, 0000000Ch
  loc_00E06556: push 00402836h ; __vbaExceptHandler
  loc_00E0655B: mov eax, fs:[00000000h]
  loc_00E06561: push eax
  loc_00E06562: mov fs:[00000000h], esp
  loc_00E06569: sub esp, 00000030h
  loc_00E0656C: push ebx
  loc_00E0656D: push esi
  loc_00E0656E: push edi
  loc_00E0656F: mov var_C, esp
  loc_00E06572: mov var_8, 00401FC8h
  loc_00E06579: mov eax, Me
  loc_00E0657C: mov ecx, eax
  loc_00E0657E: and ecx, 00000001h
  loc_00E06581: mov var_4, ecx
  loc_00E06584: and al, FEh
  loc_00E06586: push eax
  loc_00E06587: mov Me, eax
  loc_00E0658A: mov edx, [eax]
  loc_00E0658C: call [edx+00000004h]
  loc_00E0658F: mov eax, [00E3D0B0h]
  loc_00E06594: test eax, eax
  loc_00E06596: jnz 00E065A8h
  loc_00E06598: push 00E3D0B0h
  loc_00E0659D: push 006CF8E0h
  loc_00E065A2: call [004011A0h] ; __vbaNew2
  loc_00E065A8: sub esp, 00000010h
  loc_00E065AB: mov ecx, 0000000Ah
  loc_00E065B0: mov ebx, esp
  loc_00E065B2: mov var_24, ecx
  loc_00E065B5: mov eax, 80020004h
  loc_00E065BA: sub esp, 00000010h
  loc_00E065BD: mov [ebx], ecx
  loc_00E065BF: mov ecx, var_30
  loc_00E065C2: mov edx, eax
  loc_00E065C4: mov esi, [00E3D0B0h]
  loc_00E065CA: mov [ebx+00000004h], ecx
  loc_00E065CD: mov ecx, esp
  loc_00E065CF: mov edi, [esi]
  loc_00E065D1: push esi
  loc_00E065D2: mov [ebx+00000008h], eax
  loc_00E065D5: mov eax, var_28
  loc_00E065D8: mov [ebx+0000000Ch], eax
  loc_00E065DB: mov eax, var_24
  loc_00E065DE: mov [ecx], eax
  loc_00E065E0: mov eax, var_20
  loc_00E065E3: mov [ecx+00000004h], eax
  loc_00E065E6: mov [ecx+00000008h], edx
  loc_00E065E9: mov edx, var_18
  loc_00E065EC: mov [ecx+0000000Ch], edx
  loc_00E065EF: call [edi+000002B0h]
  loc_00E065F5: test eax, eax
  loc_00E065F7: fnclex
  loc_00E065F9: jge 00E0660Dh
  loc_00E065FB: push 000002B0h
  loc_00E06600: push 006DFDE0h
  loc_00E06605: push esi
  loc_00E06606: push eax
  loc_00E06607: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0660D: mov eax, [00E3D024h]
  loc_00E06612: test eax, eax
  loc_00E06614: jnz 00E06626h
  loc_00E06616: push 00E3D024h
  loc_00E0661B: push 006CE120h
  loc_00E06620: call [004011A0h] ; __vbaNew2
  loc_00E06626: mov esi, [00E3D024h]
  loc_00E0662C: push esi
  loc_00E0662D: mov eax, [esi]
  loc_00E0662F: call [eax+000002B4h]
  loc_00E06635: test eax, eax
  loc_00E06637: fnclex
  loc_00E06639: jge 00E0664Dh
  loc_00E0663B: push 000002B4h
  loc_00E06640: push 006DC970h
  loc_00E06645: push esi
  loc_00E06646: push eax
  loc_00E06647: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0664D: mov var_4, 00000000h
  loc_00E06654: mov eax, Me
  loc_00E06657: push eax
  loc_00E06658: mov ecx, [eax]
  loc_00E0665A: call [ecx+00000008h]
  loc_00E0665D: mov eax, var_4
  loc_00E06660: mov ecx, var_14
  loc_00E06663: pop edi
  loc_00E06664: pop esi
  loc_00E06665: mov fs:[00000000h], ecx
  loc_00E0666C: pop ebx
  loc_00E0666D: mov esp, ebp
  loc_00E0666F: pop ebp
  loc_00E06670: retn 0004h
End Sub

Private Sub cetak_UnknownEvent_9 'E05170
  loc_00E05170: push ebp
  loc_00E05171: mov ebp, esp
  loc_00E05173: sub esp, 0000000Ch
  loc_00E05176: push 00402836h ; __vbaExceptHandler
  loc_00E0517B: mov eax, fs:[00000000h]
  loc_00E05181: push eax
  loc_00E05182: mov fs:[00000000h], esp
  loc_00E05189: sub esp, 00000044h
  loc_00E0518C: push ebx
  loc_00E0518D: push esi
  loc_00E0518E: push edi
  loc_00E0518F: mov var_C, esp
  loc_00E05192: mov var_8, 00401F78h
  loc_00E05199: mov esi, Me
  loc_00E0519C: mov eax, esi
  loc_00E0519E: and eax, 00000001h
  loc_00E051A1: mov var_4, eax
  loc_00E051A4: and esi, FFFFFFFEh
  loc_00E051A7: push esi
  loc_00E051A8: mov Me, esi
  loc_00E051AB: mov ecx, [esi]
  loc_00E051AD: call [ecx+00000004h]
  loc_00E051B0: mov edx, [esi]
  loc_00E051B2: xor eax, eax
  loc_00E051B4: push eax
  loc_00E051B5: push FFFFFE0Bh
  loc_00E051BA: push esi
  loc_00E051BB: mov var_18, eax
  loc_00E051BE: mov var_28, eax
  loc_00E051C1: mov var_38, eax
  loc_00E051C4: call [edx+00000320h]
  loc_00E051CA: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E051D0: push eax
  loc_00E051D1: lea eax, var_18
  loc_00E051D4: push eax
  loc_00E051D5: call edi
  loc_00E051D7: lea ecx, var_28
  loc_00E051DA: push eax
  loc_00E051DB: push ecx
  loc_00E051DC: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E051E2: add esp, 00000010h
  loc_00E051E5: push eax
  loc_00E051E6: call [004011D0h] ; __vbaI4Var
  loc_00E051EC: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E051F2: xor edx, edx
  loc_00E051F4: cmp eax, 0080FF80h
  loc_00E051F9: lea ecx, var_18
  loc_00E051FC: setz dl
  loc_00E051FF: neg edx
  loc_00E05201: mov var_4C, dx
  loc_00E05205: call ebx
  loc_00E05207: lea ecx, var_28
  loc_00E0520A: call [00401024h] ; __vbaFreeVar
  loc_00E05210: cmp var_4C, 0000h
  loc_00E05215: jz 00E0530Ah
  loc_00E0521B: sub esp, 00000010h
  loc_00E0521E: mov ecx, 00000003h
  loc_00E05223: mov edx, esp
  loc_00E05225: mov eax, 00404040h
  loc_00E0522A: push FFFFFE0Bh
  loc_00E0522F: push esi
  loc_00E05230: mov [edx], ecx
  loc_00E05232: mov ecx, var_34
  loc_00E05235: mov [edx+00000004h], ecx
  loc_00E05238: mov ecx, [esi]
  loc_00E0523A: mov [edx+00000008h], eax
  loc_00E0523D: mov eax, var_2C
  loc_00E05240: mov [edx+0000000Ch], eax
  loc_00E05243: call [ecx+00000320h]
  loc_00E05249: lea edx, var_18
  loc_00E0524C: push eax
  loc_00E0524D: push edx
  loc_00E0524E: call edi
  loc_00E05250: push eax
  loc_00E05251: call [00401238h] ; __vbaLateIdSt
  loc_00E05257: lea ecx, var_18
  loc_00E0525A: call ebx
  loc_00E0525C: sub esp, 00000010h
  loc_00E0525F: mov ecx, 00000003h
  loc_00E05264: mov edx, esp
  loc_00E05266: mov eax, 00E0E0E0h
  loc_00E0526B: push FFFFFDFFh
  loc_00E05270: push esi
  loc_00E05271: mov [edx], ecx
  loc_00E05273: mov ecx, var_34
  loc_00E05276: mov [edx+00000004h], ecx
  loc_00E05279: mov ecx, [esi]
  loc_00E0527B: mov [edx+00000008h], eax
  loc_00E0527E: mov eax, var_2C
  loc_00E05281: mov [edx+0000000Ch], eax
  loc_00E05284: call [ecx+00000320h]
  loc_00E0528A: lea edx, var_18
  loc_00E0528D: push eax
  loc_00E0528E: push edx
  loc_00E0528F: call edi
  loc_00E05291: push eax
  loc_00E05292: call [00401238h] ; __vbaLateIdSt
  loc_00E05298: lea ecx, var_18
  loc_00E0529B: call ebx
  loc_00E0529D: sub esp, 00000010h
  loc_00E052A0: mov ecx, 00000008h
  loc_00E052A5: mov edx, esp
  loc_00E052A7: mov eax, 006DFC80h ; "Pilih------>"
  loc_00E052AC: push FFFFFDFAh
  loc_00E052B1: push esi
  loc_00E052B2: mov [edx], ecx
  loc_00E052B4: mov ecx, var_34
  loc_00E052B7: mov [edx+00000004h], ecx
  loc_00E052BA: mov ecx, [esi]
  loc_00E052BC: mov [edx+00000008h], eax
  loc_00E052BF: mov eax, var_2C
  loc_00E052C2: mov [edx+0000000Ch], eax
  loc_00E052C5: call [ecx+00000320h]
  loc_00E052CB: lea edx, var_18
  loc_00E052CE: push eax
  loc_00E052CF: push edx
  loc_00E052D0: call edi
  loc_00E052D2: push eax
  loc_00E052D3: call [00401238h] ; __vbaLateIdSt
  loc_00E052D9: lea ecx, var_18
  loc_00E052DC: call ebx
  loc_00E052DE: mov eax, [esi]
  loc_00E052E0: push esi
  loc_00E052E1: call [eax+00000300h]
  loc_00E052E7: lea ecx, var_18
  loc_00E052EA: push eax
  loc_00E052EB: push ecx
  loc_00E052EC: call edi
  loc_00E052EE: mov esi, eax
  loc_00E052F0: push FFFFFFFFh
  loc_00E052F2: push esi
  loc_00E052F3: mov edx, [esi]
  loc_00E052F5: call [edx+0000009Ch]
  loc_00E052FB: test eax, eax
  loc_00E052FD: fnclex
  loc_00E052FF: jge 00E05457h
  loc_00E05305: jmp 00E05445h
  loc_00E0530A: mov eax, [esi]
  loc_00E0530C: push 00000000h
  loc_00E0530E: push FFFFFE0Bh
  loc_00E05313: push esi
  loc_00E05314: call [eax+00000320h]
  loc_00E0531A: lea ecx, var_18
  loc_00E0531D: push eax
  loc_00E0531E: push ecx
  loc_00E0531F: call edi
  loc_00E05321: lea edx, var_28
  loc_00E05324: push eax
  loc_00E05325: push edx
  loc_00E05326: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0532C: add esp, 00000010h
  loc_00E0532F: push eax
  loc_00E05330: call [004011D0h] ; __vbaI4Var
  loc_00E05336: xor ecx, ecx
  loc_00E05338: cmp eax, 00404040h
  loc_00E0533D: setz cl
  loc_00E05340: neg ecx
  loc_00E05342: mov var_4C, cx
  loc_00E05346: lea ecx, var_18
  loc_00E05349: call ebx
  loc_00E0534B: lea ecx, var_28
  loc_00E0534E: call [00401024h] ; __vbaFreeVar
  loc_00E05354: cmp var_4C, 0000h
  loc_00E05359: jz 00E0545Ch
  loc_00E0535F: sub esp, 00000010h
  loc_00E05362: mov ecx, 00000003h
  loc_00E05367: mov edx, esp
  loc_00E05369: mov eax, 0080FF80h
  loc_00E0536E: push FFFFFE0Bh
  loc_00E05373: push esi
  loc_00E05374: mov [edx], ecx
  loc_00E05376: mov ecx, var_34
  loc_00E05379: mov [edx+00000004h], ecx
  loc_00E0537C: mov ecx, [esi]
  loc_00E0537E: mov [edx+00000008h], eax
  loc_00E05381: mov eax, var_2C
  loc_00E05384: mov [edx+0000000Ch], eax
  loc_00E05387: call [ecx+00000320h]
  loc_00E0538D: lea edx, var_18
  loc_00E05390: push eax
  loc_00E05391: push edx
  loc_00E05392: call edi
  loc_00E05394: push eax
  loc_00E05395: call [00401238h] ; __vbaLateIdSt
  loc_00E0539B: lea ecx, var_18
  loc_00E0539E: call ebx
  loc_00E053A0: sub esp, 00000010h
  loc_00E053A3: mov ecx, 00000003h
  loc_00E053A8: mov edx, esp
  loc_00E053AA: mov eax, 00004000h
  loc_00E053AF: push FFFFFDFFh
  loc_00E053B4: push esi
  loc_00E053B5: mov [edx], ecx
  loc_00E053B7: mov ecx, var_34
  loc_00E053BA: mov [edx+00000004h], ecx
  loc_00E053BD: mov ecx, [esi]
  loc_00E053BF: mov [edx+00000008h], eax
  loc_00E053C2: mov eax, var_2C
  loc_00E053C5: mov [edx+0000000Ch], eax
  loc_00E053C8: call [ecx+00000320h]
  loc_00E053CE: lea edx, var_18
  loc_00E053D1: push eax
  loc_00E053D2: push edx
  loc_00E053D3: call edi
  loc_00E053D5: push eax
  loc_00E053D6: call [00401238h] ; __vbaLateIdSt
  loc_00E053DC: lea ecx, var_18
  loc_00E053DF: call ebx
  loc_00E053E1: sub esp, 00000010h
  loc_00E053E4: mov ecx, 00000008h
  loc_00E053E9: mov edx, esp
  loc_00E053EB: mov eax, 006DFCA0h ; "Pencetakaan"
  loc_00E053F0: push FFFFFDFAh
  loc_00E053F5: push esi
  loc_00E053F6: mov [edx], ecx
  loc_00E053F8: mov ecx, var_34
  loc_00E053FB: mov [edx+00000004h], ecx
  loc_00E053FE: mov ecx, [esi]
  loc_00E05400: mov [edx+00000008h], eax
  loc_00E05403: mov eax, var_2C
  loc_00E05406: mov [edx+0000000Ch], eax
  loc_00E05409: call [ecx+00000320h]
  loc_00E0540F: lea edx, var_18
  loc_00E05412: push eax
  loc_00E05413: push edx
  loc_00E05414: call edi
  loc_00E05416: push eax
  loc_00E05417: call [00401238h] ; __vbaLateIdSt
  loc_00E0541D: lea ecx, var_18
  loc_00E05420: call ebx
  loc_00E05422: mov eax, [esi]
  loc_00E05424: push esi
  loc_00E05425: call [eax+00000300h]
  loc_00E0542B: lea ecx, var_18
  loc_00E0542E: push eax
  loc_00E0542F: push ecx
  loc_00E05430: call edi
  loc_00E05432: mov esi, eax
  loc_00E05434: push 00000000h
  loc_00E05436: push esi
  loc_00E05437: mov edx, [esi]
  loc_00E05439: call [edx+0000009Ch]
  loc_00E0543F: test eax, eax
  loc_00E05441: fnclex
  loc_00E05443: jge 00E05457h
  loc_00E05445: push 0000009Ch
  loc_00E0544A: push 006DCAD0h
  loc_00E0544F: push esi
  loc_00E05450: push eax
  loc_00E05451: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E05457: lea ecx, var_18
  loc_00E0545A: call ebx
  loc_00E0545C: mov var_4, 00000000h
  loc_00E05463: push 00E0547Eh
  loc_00E05468: jmp 00E0547Dh
  loc_00E0546A: lea ecx, var_18
  loc_00E0546D: call [00401254h] ; __vbaFreeObj
  loc_00E05473: lea ecx, var_28
  loc_00E05476: call [00401024h] ; __vbaFreeVar
  loc_00E0547C: ret
  loc_00E0547D: ret
  loc_00E0547E: mov eax, Me
  loc_00E05481: push eax
  loc_00E05482: mov ecx, [eax]
  loc_00E05484: call [ecx+00000008h]
  loc_00E05487: mov eax, var_4
  loc_00E0548A: mov ecx, var_14
  loc_00E0548D: pop edi
  loc_00E0548E: pop esi
  loc_00E0548F: mov fs:[00000000h], ecx
  loc_00E05496: pop ebx
  loc_00E05497: mov esp, ebp
  loc_00E05499: pop ebp
  loc_00E0549A: retn 0004h
End Sub

Public Sub Tengah() 'E07080
  loc_00E07080: push ebp
  loc_00E07081: mov ebp, esp
  loc_00E07083: sub esp, 0000000Ch
  loc_00E07086: push 00402836h ; __vbaExceptHandler
  loc_00E0708B: mov eax, fs:[00000000h]
  loc_00E07091: push eax
  loc_00E07092: mov fs:[00000000h], esp
  loc_00E07099: sub esp, 0000002Ch
  loc_00E0709C: push ebx
  loc_00E0709D: push esi
  loc_00E0709E: push edi
  loc_00E0709F: mov var_C, esp
  loc_00E070A2: mov var_8, 00402030h
  loc_00E070A9: xor edi, edi
  loc_00E070AB: mov var_4, edi
  loc_00E070AE: mov esi, Me
  loc_00E070B1: push esi
  loc_00E070B2: mov eax, [esi]
  loc_00E070B4: call [eax+00000004h]
  loc_00E070B7: mov ecx, [esi]
  loc_00E070B9: lea edx, var_20
  loc_00E070BC: push edx
  loc_00E070BD: push esi
  loc_00E070BE: mov var_18, edi
  loc_00E070C1: mov var_1C, edi
  loc_00E070C4: mov var_20, edi
  loc_00E070C7: call [ecx+00000080h]
  loc_00E070CD: cmp eax, edi
  loc_00E070CF: fnclex
  loc_00E070D1: jge 00E070E9h
  loc_00E070D3: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E070D9: push 00000080h
  loc_00E070DE: push 006DC970h
  loc_00E070E3: push esi
  loc_00E070E4: push eax
  loc_00E070E5: call ebx
  loc_00E070E7: jmp 00E070EFh
  loc_00E070E9: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E070EF: cmp [00E3D9CCh], edi
  loc_00E070F5: jnz 00E07107h
  loc_00E070F7: push 00E3D9CCh
  loc_00E070FC: push 006DC960h
  loc_00E07101: call [004011A0h] ; __vbaNew2
  loc_00E07107: mov edi, [00E3D9CCh]
  loc_00E0710D: lea ecx, var_18
  loc_00E07110: push ecx
  loc_00E07111: push edi
  loc_00E07112: mov eax, [edi]
  loc_00E07114: call [eax+00000018h]
  loc_00E07117: test eax, eax
  loc_00E07119: fnclex
  loc_00E0711B: jge 00E07128h
  loc_00E0711D: push 00000018h
  loc_00E0711F: push 006DC950h
  loc_00E07124: push edi
  loc_00E07125: push eax
  loc_00E07126: call ebx
  loc_00E07128: mov eax, var_18
  loc_00E0712B: lea ecx, var_1C
  loc_00E0712E: push ecx
  loc_00E0712F: push eax
  loc_00E07130: mov edx, [eax]
  loc_00E07132: mov edi, eax
  loc_00E07134: call [edx+00000098h]
  loc_00E0713A: test eax, eax
  loc_00E0713C: fnclex
  loc_00E0713E: jge 00E0714Eh
  loc_00E07140: push 00000098h
  loc_00E07145: push 006DCAF0h
  loc_00E0714A: push edi
  loc_00E0714B: push eax
  loc_00E0714C: call ebx
  loc_00E0714E: fld real4 ptr var_1C
  loc_00E07151: fsub st0, real4 ptr var_20
  loc_00E07154: mov edx, [esi]
  loc_00E07156: push ecx
  loc_00E07157: cmp [00E3D000h], 00000000h
  loc_00E0715E: jnz 00E07168h
  loc_00E07160: fdiv st0, real4 ptr [00402028h]
  loc_00E07166: jmp 00E07173h
  loc_00E07168: push [00402028h]
  loc_00E0716E: call 00402848h ; _adj_fdiv_m32
  loc_00E07173: fnstsw ax
  loc_00E07175: test al, 0Dh
  loc_00E07177: jnz 00E07297h
  loc_00E0717D: fstp real4 ptr [esp]
  loc_00E07180: push esi
  loc_00E07181: call [edx+00000074h]
  loc_00E07184: test eax, eax
  loc_00E07186: fnclex
  loc_00E07188: jge 00E07195h
  loc_00E0718A: push 00000074h
  loc_00E0718C: push 006DC970h
  loc_00E07191: push esi
  loc_00E07192: push eax
  loc_00E07193: call ebx
  loc_00E07195: lea ecx, var_18
  loc_00E07198: call [00401254h] ; __vbaFreeObj
  loc_00E0719E: mov eax, [esi]
  loc_00E071A0: lea ecx, var_20
  loc_00E071A3: push ecx
  loc_00E071A4: push esi
  loc_00E071A5: call [eax+00000088h]
  loc_00E071AB: test eax, eax
  loc_00E071AD: fnclex
  loc_00E071AF: jge 00E071BFh
  loc_00E071B1: push 00000088h
  loc_00E071B6: push 006DC970h
  loc_00E071BB: push esi
  loc_00E071BC: push eax
  loc_00E071BD: call ebx
  loc_00E071BF: mov eax, [00E3D9CCh]
  loc_00E071C4: test eax, eax
  loc_00E071C6: jnz 00E071D8h
  loc_00E071C8: push 00E3D9CCh
  loc_00E071CD: push 006DC960h
  loc_00E071D2: call [004011A0h] ; __vbaNew2
  loc_00E071D8: mov edi, [00E3D9CCh]
  loc_00E071DE: lea eax, var_18
  loc_00E071E1: push eax
  loc_00E071E2: push edi
  loc_00E071E3: mov edx, [edi]
  loc_00E071E5: call [edx+00000018h]
  loc_00E071E8: test eax, eax
  loc_00E071EA: fnclex
  loc_00E071EC: jge 00E071F9h
  loc_00E071EE: push 00000018h
  loc_00E071F0: push 006DC950h
  loc_00E071F5: push edi
  loc_00E071F6: push eax
  loc_00E071F7: call ebx
  loc_00E071F9: mov eax, var_18
  loc_00E071FC: lea edx, var_1C
  loc_00E071FF: push edx
  loc_00E07200: push eax
  loc_00E07201: mov ecx, [eax]
  loc_00E07203: mov edi, eax
  loc_00E07205: call [ecx+00000050h]
  loc_00E07208: test eax, eax
  loc_00E0720A: fnclex
  loc_00E0720C: jge 00E07219h
  loc_00E0720E: push 00000050h
  loc_00E07210: push 006DCAF0h
  loc_00E07215: push edi
  loc_00E07216: push eax
  loc_00E07217: call ebx
  loc_00E07219: fld real4 ptr var_1C
  loc_00E0721C: fsub st0, real4 ptr var_20
  loc_00E0721F: mov ecx, [esi]
  loc_00E07221: push ecx
  loc_00E07222: cmp [00E3D000h], 00000000h
  loc_00E07229: jnz 00E07233h
  loc_00E0722B: fdiv st0, real4 ptr [00402028h]
  loc_00E07231: jmp 00E0723Eh
  loc_00E07233: push [00402028h]
  loc_00E07239: call 00402848h ; _adj_fdiv_m32
  loc_00E0723E: fnstsw ax
  loc_00E07240: test al, 0Dh
  loc_00E07242: jnz 00E07297h
  loc_00E07244: fstp real4 ptr [esp]
  loc_00E07247: push esi
  loc_00E07248: call [ecx+0000007Ch]
  loc_00E0724B: test eax, eax
  loc_00E0724D: fnclex
  loc_00E0724F: jge 00E0725Ch
  loc_00E07251: push 0000007Ch
  loc_00E07253: push 006DC970h
  loc_00E07258: push esi
  loc_00E07259: push eax
  loc_00E0725A: call ebx
  loc_00E0725C: lea ecx, var_18
  loc_00E0725F: call [00401254h] ; __vbaFreeObj
  loc_00E07265: fwait
  loc_00E07266: push 00E07278h
  loc_00E0726B: jmp 00E07277h
  loc_00E0726D: lea ecx, var_18
  loc_00E07270: call [00401254h] ; __vbaFreeObj
  loc_00E07276: ret
  loc_00E07277: ret
  loc_00E07278: mov eax, Me
  loc_00E0727B: push eax
  loc_00E0727C: mov edx, [eax]
  loc_00E0727E: call [edx+00000008h]
  loc_00E07281: mov eax, var_4
  loc_00E07284: mov ecx, var_14
  loc_00E07287: pop edi
  loc_00E07288: pop esi
  loc_00E07289: mov fs:[00000000h], ecx
  loc_00E07290: pop ebx
  loc_00E07291: mov esp, ebp
  loc_00E07293: pop ebp
  loc_00E07294: retn 0004h
End Sub
