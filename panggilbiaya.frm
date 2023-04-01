VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form panggilbiaya
  Caption = "Form1"
  BackColor = &H404040&
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  BorderStyle = 0 'None
  'Icon = n/a
  LinkTopic = "Form1"
  MaxButton = 0   'False
  MinButton = 0   'False
  ClientLeft = 0
  ClientTop = 0
  ClientWidth = 11685
  ClientHeight = 5775
  ShowInTaskbar = 0   'False
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fcari
    Caption = "Cari Disini"
    BackColor = &HE0E0E0&
    ForeColor = &H404040&
    Left = 90
    Top = 870
    Width = 11475
    Height = 1695
    TabIndex = 0
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.OptionButton optstatus
      Caption = "Status Pembayaran"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 5310
      Top = 540
      Width = 2265
      Height = 300
      TabIndex = 6
      BeginProperty Font
        Name = "Trebuchet MS"
        Size = 11.25
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.OptionButton optnibu
      Caption = "Nama"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 330
      Top = 510
      Width = 1455
      Height = 300
      TabIndex = 3
      BeginProperty Font
        Name = "Trebuchet MS"
        Size = 11.25
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.OptionButton optno
      Caption = "No. Daftar"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 2640
      Top = 510
      Width = 1455
      Height = 300
      TabIndex = 2
      BeginProperty Font
        Name = "Trebuchet MS"
        Size = 11.25
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = 0 'False
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.TextBox txtcari
      BackColor = &HC0C0C0&
      Left = 240
      Top = 1110
      Width = 10905
      Height = 345
      TabIndex = 1
      BorderStyle = 0 'None
    End
    Begin VB.Shape Shape2
      Left = 300
      Top = 1140
      Width = 10935
      Height = 405
      BorderStyle = 0 'Transparent
      FillColor = &H404040&
      FillStyle = 0
    End
  End
  Begin MSAdodcLib.Adodc adopen
    Left = 240
    Top = 4980
    Width = 1200
    Height = 330
    Visible = 0   'False
    OleObjectBlob = "panggilbiaya.frx":0000
  End
  Begin MSDataGridLib.DataGrid dgpen
    Left = 90
    Top = 2670
    Width = 11535
    Height = 3015
    TabIndex = 4
    OleObjectBlob = "panggilbiaya.frx":0332
  End
  Begin VB.Shape Shape3
    Left = 0
    Top = 450
    Width = 14685
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Label Label4
    Caption = "Riwayat Pembayaran"
    ForeColor = &H80FF80&
    Left = 4710
    Top = 30
    Width = 2745
    Height = 315
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
  Begin VB.Image back
    Picture = "panggilbiaya.frx":04DD
    Left = 30
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Shape Shape1
    Left = 0
    Top = 0
    Width = 14685
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
End

Attribute VB_Name = "panggilbiaya"


Private Sub txtcari_KeyPress(KeyAscii As Integer) 'E32570
  loc_00E32570: push ebp
  loc_00E32571: mov ebp, esp
  loc_00E32573: sub esp, 0000000Ch
  loc_00E32576: push 00402836h ; __vbaExceptHandler
  loc_00E3257B: mov eax, fs:[00000000h]
  loc_00E32581: push eax
  loc_00E32582: mov fs:[00000000h], esp
  loc_00E32589: sub esp, 000000B4h
  loc_00E3258F: push ebx
  loc_00E32590: push esi
  loc_00E32591: push edi
  loc_00E32592: mov var_C, esp
  loc_00E32595: mov var_8, 004025E0h
  loc_00E3259C: mov esi, Me
  loc_00E3259F: mov eax, esi
  loc_00E325A1: and eax, 00000001h
  loc_00E325A4: mov var_4, eax
  loc_00E325A7: and esi, FFFFFFFEh
  loc_00E325AA: push esi
  loc_00E325AB: mov Me, esi
  loc_00E325AE: mov ecx, [esi]
  loc_00E325B0: call [ecx+00000004h]
  loc_00E325B3: mov edx, [esi]
  loc_00E325B5: xor ebx, ebx
  loc_00E325B7: push esi
  loc_00E325B8: mov var_24, ebx
  loc_00E325BB: mov var_28, ebx
  loc_00E325BE: mov var_2C, ebx
  loc_00E325C1: mov var_30, ebx
  loc_00E325C4: mov var_40, ebx
  loc_00E325C7: mov var_50, ebx
  loc_00E325CA: mov var_60, ebx
  loc_00E325CD: mov var_70, ebx
  loc_00E325D0: mov var_80, ebx
  loc_00E325D3: mov var_90, ebx
  loc_00E325D9: mov var_B4, ebx
  loc_00E325DF: call [edx+00000308h]
  loc_00E325E5: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E325EB: push eax
  loc_00E325EC: lea eax, var_30
  loc_00E325EF: push eax
  loc_00E325F0: call edi
  loc_00E325F2: mov ecx, [eax]
  loc_00E325F4: lea edx, var_B4
  loc_00E325FA: push edx
  loc_00E325FB: push eax
  loc_00E325FC: mov var_B8, eax
  loc_00E32602: call [ecx+000000E0h]
  loc_00E32608: cmp eax, ebx
  loc_00E3260A: fnclex
  loc_00E3260C: jge 00E32626h
  loc_00E3260E: mov ecx, var_B8
  loc_00E32614: push 000000E0h
  loc_00E32619: push 006E03D4h
  loc_00E3261E: push ecx
  loc_00E3261F: push eax
  loc_00E32620: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32626: mov edx, var_B4
  loc_00E3262C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E32632: lea ecx, var_30
  loc_00E32635: mov var_C0, edx
  loc_00E3263B: call ebx
  loc_00E3263D: cmp var_C0, 0000h
  loc_00E32645: jz 00E326ADh
  loc_00E32647: mov eax, [esi]
  loc_00E32649: push esi
  loc_00E3264A: call [eax+0000030Ch]
  loc_00E32650: lea ecx, var_30
  loc_00E32653: push eax
  loc_00E32654: push ecx
  loc_00E32655: call edi
  loc_00E32657: mov edx, [eax]
  loc_00E32659: lea ecx, var_28
  loc_00E3265C: push ecx
  loc_00E3265D: push eax
  loc_00E3265E: mov var_B8, eax
  loc_00E32664: call [edx+000000A0h]
  loc_00E3266A: test eax, eax
  loc_00E3266C: fnclex
  loc_00E3266E: jge 00E32688h
  loc_00E32670: mov edx, var_B8
  loc_00E32676: push 000000A0h
  loc_00E3267B: push 006DCB70h
  loc_00E32680: push edx
  loc_00E32681: push eax
  loc_00E32682: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32688: mov eax, var_28
  loc_00E3268B: push 006E1958h ; "select * from lc where no_daftar like '" & Chr(37)
  loc_00E32690: push eax
  loc_00E32691: call [0040106Ch] ; __vbaStrCat
  loc_00E32697: mov edx, eax
  loc_00E32699: lea ecx, var_2C
  loc_00E3269C: call [00401228h] ; __vbaStrMove
  loc_00E326A2: push eax
  loc_00E326A3: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E326A8: jmp 00E32836h
  loc_00E326AD: mov edx, [esi]
  loc_00E326AF: push esi
  loc_00E326B0: call [edx+00000304h]
  loc_00E326B6: push eax
  loc_00E326B7: lea eax, var_30
  loc_00E326BA: push eax
  loc_00E326BB: call edi
  loc_00E326BD: mov ecx, [eax]
  loc_00E326BF: lea edx, var_B4
  loc_00E326C5: push edx
  loc_00E326C6: push eax
  loc_00E326C7: mov var_B8, eax
  loc_00E326CD: call [ecx+000000E0h]
  loc_00E326D3: test eax, eax
  loc_00E326D5: fnclex
  loc_00E326D7: jge 00E326F1h
  loc_00E326D9: mov ecx, var_B8
  loc_00E326DF: push 000000E0h
  loc_00E326E4: push 006E03D4h
  loc_00E326E9: push ecx
  loc_00E326EA: push eax
  loc_00E326EB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E326F1: mov edx, var_B4
  loc_00E326F7: lea ecx, var_30
  loc_00E326FA: mov var_C0, edx
  loc_00E32700: call ebx
  loc_00E32702: cmp var_C0, 0000h
  loc_00E3270A: jz 00E32772h
  loc_00E3270C: mov eax, [esi]
  loc_00E3270E: push esi
  loc_00E3270F: call [eax+0000030Ch]
  loc_00E32715: lea ecx, var_30
  loc_00E32718: push eax
  loc_00E32719: push ecx
  loc_00E3271A: call edi
  loc_00E3271C: mov edx, [eax]
  loc_00E3271E: lea ecx, var_28
  loc_00E32721: push ecx
  loc_00E32722: push eax
  loc_00E32723: mov var_B8, eax
  loc_00E32729: call [edx+000000A0h]
  loc_00E3272F: test eax, eax
  loc_00E32731: fnclex
  loc_00E32733: jge 00E3274Dh
  loc_00E32735: mov edx, var_B8
  loc_00E3273B: push 000000A0h
  loc_00E32740: push 006DCB70h
  loc_00E32745: push edx
  loc_00E32746: push eax
  loc_00E32747: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3274D: mov eax, var_28
  loc_00E32750: push 006E19DCh ; "select * from lc where nama_siswa like '" & Chr(37)
  loc_00E32755: push eax
  loc_00E32756: call [0040106Ch] ; __vbaStrCat
  loc_00E3275C: mov edx, eax
  loc_00E3275E: lea ecx, var_2C
  loc_00E32761: call [00401228h] ; __vbaStrMove
  loc_00E32767: push eax
  loc_00E32768: push 006E1A34h ; Chr(37) & "' order by nama_siswa asc"
  loc_00E3276D: jmp 00E32836h
  loc_00E32772: mov edx, [esi]
  loc_00E32774: push esi
  loc_00E32775: call [edx+00000300h]
  loc_00E3277B: push eax
  loc_00E3277C: lea eax, var_30
  loc_00E3277F: push eax
  loc_00E32780: call edi
  loc_00E32782: mov ecx, [eax]
  loc_00E32784: lea edx, var_B4
  loc_00E3278A: push edx
  loc_00E3278B: push eax
  loc_00E3278C: mov var_B8, eax
  loc_00E32792: call [ecx+000000E0h]
  loc_00E32798: test eax, eax
  loc_00E3279A: fnclex
  loc_00E3279C: jge 00E327B6h
  loc_00E3279E: mov ecx, var_B8
  loc_00E327A4: push 000000E0h
  loc_00E327A9: push 006E03D4h
  loc_00E327AE: push ecx
  loc_00E327AF: push eax
  loc_00E327B0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E327B6: mov edx, var_B4
  loc_00E327BC: lea ecx, var_30
  loc_00E327BF: mov var_C0, edx
  loc_00E327C5: call ebx
  loc_00E327C7: cmp var_C0, 0000h
  loc_00E327CF: jz 00E328DFh
  loc_00E327D5: mov eax, [esi]
  loc_00E327D7: push esi
  loc_00E327D8: call [eax+0000030Ch]
  loc_00E327DE: lea ecx, var_30
  loc_00E327E1: push eax
  loc_00E327E2: push ecx
  loc_00E327E3: call edi
  loc_00E327E5: mov edx, [eax]
  loc_00E327E7: lea ecx, var_28
  loc_00E327EA: push ecx
  loc_00E327EB: push eax
  loc_00E327EC: mov var_B8, eax
  loc_00E327F2: call [edx+000000A0h]
  loc_00E327F8: test eax, eax
  loc_00E327FA: fnclex
  loc_00E327FC: jge 00E32816h
  loc_00E327FE: mov edx, var_B8
  loc_00E32804: push 000000A0h
  loc_00E32809: push 006DCB70h
  loc_00E3280E: push edx
  loc_00E3280F: push eax
  loc_00E32810: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32816: mov eax, var_28
  loc_00E32819: push 006E1C38h ; "select * from lc where metode_pembayaran like '" & Chr(37)
  loc_00E3281E: push eax
  loc_00E3281F: call [0040106Ch] ; __vbaStrCat
  loc_00E32825: mov edx, eax
  loc_00E32827: lea ecx, var_2C
  loc_00E3282A: call [00401228h] ; __vbaStrMove
  loc_00E32830: push eax
  loc_00E32831: push 006E1CA0h ; Chr(37) & "' order by metode_pembayaran asc"
  loc_00E32836: call [0040106Ch] ; __vbaStrCat
  loc_00E3283C: lea edx, var_40
  loc_00E3283F: lea ecx, var_24
  loc_00E32842: mov var_38, eax
  loc_00E32845: mov var_40, 00000008h
  loc_00E3284C: call [0040101Ch] ; __vbaVarMove
  loc_00E32852: lea ecx, var_2C
  loc_00E32855: lea edx, var_28
  loc_00E32858: push ecx
  loc_00E32859: push edx
  loc_00E3285A: push 00000002h
  loc_00E3285C: call [004011BCh] ; __vbaFreeStrList
  loc_00E32862: add esp, 0000000Ch
  loc_00E32865: lea ecx, var_30
  loc_00E32868: call ebx
  loc_00E3286A: lea eax, var_24
  loc_00E3286D: push eax
  loc_00E3286E: call [00401230h] ; __vbaStrVarCopy
  loc_00E32874: sub esp, 00000010h
  loc_00E32877: mov ecx, 00000008h
  loc_00E3287C: mov edx, esp
  loc_00E3287E: mov var_40, ecx
  loc_00E32881: mov var_38, eax
  loc_00E32884: push 0000000Eh
  loc_00E32886: mov [edx], ecx
  loc_00E32888: mov ecx, var_3C
  loc_00E3288B: push esi
  loc_00E3288C: mov [edx+00000004h], ecx
  loc_00E3288F: mov ecx, [esi]
  loc_00E32891: mov [edx+00000008h], eax
  loc_00E32894: mov eax, var_34
  loc_00E32897: mov [edx+0000000Ch], eax
  loc_00E3289A: call [ecx+00000324h]
  loc_00E328A0: lea edx, var_30
  loc_00E328A3: push eax
  loc_00E328A4: push edx
  loc_00E328A5: call edi
  loc_00E328A7: push eax
  loc_00E328A8: call [00401238h] ; __vbaLateIdSt
  loc_00E328AE: lea ecx, var_30
  loc_00E328B1: call ebx
  loc_00E328B3: lea ecx, var_40
  loc_00E328B6: call [00401024h] ; __vbaFreeVar
  loc_00E328BC: mov eax, [esi]
  loc_00E328BE: push 00000000h
  loc_00E328C0: push 00000019h
  loc_00E328C2: push esi
  loc_00E328C3: call [eax+00000324h]
  loc_00E328C9: lea ecx, var_30
  loc_00E328CC: push eax
  loc_00E328CD: push ecx
  loc_00E328CE: call edi
  loc_00E328D0: push eax
  loc_00E328D1: call [00401030h] ; __vbaLateIdCall
  loc_00E328D7: add esp, 0000000Ch
  loc_00E328DA: jmp 00E32A08h
  loc_00E328DF: mov edx, [esi]
  loc_00E328E1: push esi
  loc_00E328E2: call [edx+00000308h]
  loc_00E328E8: push eax
  loc_00E328E9: lea eax, var_30
  loc_00E328EC: push eax
  loc_00E328ED: call edi
  loc_00E328EF: mov ecx, [eax]
  loc_00E328F1: lea edx, var_B4
  loc_00E328F7: push edx
  loc_00E328F8: push eax
  loc_00E328F9: mov var_B8, eax
  loc_00E328FF: call [ecx+000000E0h]
  loc_00E32905: test eax, eax
  loc_00E32907: fnclex
  loc_00E32909: jge 00E32923h
  loc_00E3290B: mov ecx, var_B8
  loc_00E32911: push 000000E0h
  loc_00E32916: push 006E03D4h
  loc_00E3291B: push ecx
  loc_00E3291C: push eax
  loc_00E3291D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32923: xor edx, edx
  loc_00E32925: lea ecx, var_30
  loc_00E32928: cmp var_B4, dx
  loc_00E3292F: setz dl
  loc_00E32932: neg edx
  loc_00E32934: mov var_C0, edx
  loc_00E3293A: call ebx
  loc_00E3293C: cmp var_C0, 0000h
  loc_00E32944: jz 00E32A0Dh
  loc_00E3294A: mov ecx, 80020004h
  loc_00E3294F: mov eax, 0000000Ah
  loc_00E32954: mov var_68, ecx
  loc_00E32957: mov var_58, ecx
  loc_00E3295A: lea edx, var_90
  loc_00E32960: lea ecx, var_50
  loc_00E32963: mov var_70, eax
  loc_00E32966: mov var_60, eax
  loc_00E32969: mov var_88, 006E1AE8h ; "Need To Do !1"
  loc_00E32973: mov var_90, 00000008h
  loc_00E3297D: call [004011F0h] ; __vbaVarDup
  loc_00E32983: lea edx, var_80
  loc_00E32986: lea ecx, var_40
  loc_00E32989: mov var_78, 006E1A70h ; "Silahkan Pilih Item Yang ingin di cari terlebih dahulu !!"
  loc_00E32990: mov var_80, 00000008h
  loc_00E32997: call [004011F0h] ; __vbaVarDup
  loc_00E3299D: lea eax, var_70
  loc_00E329A0: lea ecx, var_60
  loc_00E329A3: push eax
  loc_00E329A4: lea edx, var_50
  loc_00E329A7: push ecx
  loc_00E329A8: push edx
  loc_00E329A9: lea eax, var_40
  loc_00E329AC: push 00000010h
  loc_00E329AE: push eax
  loc_00E329AF: call [004010A8h] ; rtcMsgBox
  loc_00E329B5: lea ecx, var_70
  loc_00E329B8: lea edx, var_60
  loc_00E329BB: push ecx
  loc_00E329BC: lea eax, var_50
  loc_00E329BF: push edx
  loc_00E329C0: lea ecx, var_40
  loc_00E329C3: push eax
  loc_00E329C4: push ecx
  loc_00E329C5: push 00000004h
  loc_00E329C7: call [00401038h] ; __vbaFreeVarList
  loc_00E329CD: mov edx, [esi]
  loc_00E329CF: add esp, 00000014h
  loc_00E329D2: push esi
  loc_00E329D3: call [edx+0000030Ch]
  loc_00E329D9: push eax
  loc_00E329DA: lea eax, var_30
  loc_00E329DD: push eax
  loc_00E329DE: call edi
  loc_00E329E0: mov esi, eax
  loc_00E329E2: push 006DCC80h
  loc_00E329E7: push esi
  loc_00E329E8: mov ecx, [esi]
  loc_00E329EA: call [ecx+000000A4h]
  loc_00E329F0: test eax, eax
  loc_00E329F2: fnclex
  loc_00E329F4: jge 00E32A08h
  loc_00E329F6: push 000000A4h
  loc_00E329FB: push 006DCB70h
  loc_00E32A00: push esi
  loc_00E32A01: push eax
  loc_00E32A02: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32A08: lea ecx, var_30
  loc_00E32A0B: call ebx
  loc_00E32A0D: mov var_4, 00000000h
  loc_00E32A14: push 00E32A5Dh
  loc_00E32A19: jmp 00E32A53h
  loc_00E32A1B: lea edx, var_2C
  loc_00E32A1E: lea eax, var_28
  loc_00E32A21: push edx
  loc_00E32A22: push eax
  loc_00E32A23: push 00000002h
  loc_00E32A25: call [004011BCh] ; __vbaFreeStrList
  loc_00E32A2B: add esp, 0000000Ch
  loc_00E32A2E: lea ecx, var_30
  loc_00E32A31: call [00401254h] ; __vbaFreeObj
  loc_00E32A37: lea ecx, var_70
  loc_00E32A3A: lea edx, var_60
  loc_00E32A3D: push ecx
  loc_00E32A3E: lea eax, var_50
  loc_00E32A41: push edx
  loc_00E32A42: lea ecx, var_40
  loc_00E32A45: push eax
  loc_00E32A46: push ecx
  loc_00E32A47: push 00000004h
  loc_00E32A49: call [00401038h] ; __vbaFreeVarList
  loc_00E32A4F: add esp, 00000014h
  loc_00E32A52: ret
  loc_00E32A53: lea ecx, var_24
  loc_00E32A56: call [00401024h] ; __vbaFreeVar
  loc_00E32A5C: ret
  loc_00E32A5D: mov eax, Me
  loc_00E32A60: push eax
  loc_00E32A61: mov edx, [eax]
  loc_00E32A63: call [edx+00000008h]
  loc_00E32A66: mov eax, var_4
  loc_00E32A69: mov ecx, var_14
  loc_00E32A6C: pop edi
  loc_00E32A6D: pop esi
  loc_00E32A6E: mov fs:[00000000h], ecx
  loc_00E32A75: pop ebx
  loc_00E32A76: mov esp, ebp
  loc_00E32A78: pop ebp
  loc_00E32A79: retn 0008h
End Sub

Private Sub Form_Load() 'E32380
  loc_00E32380: push ebp
  loc_00E32381: mov ebp, esp
  loc_00E32383: sub esp, 0000000Ch
  loc_00E32386: push 00402836h ; __vbaExceptHandler
  loc_00E3238B: mov eax, fs:[00000000h]
  loc_00E32391: push eax
  loc_00E32392: mov fs:[00000000h], esp
  loc_00E32399: sub esp, 00000040h
  loc_00E3239C: push ebx
  loc_00E3239D: push esi
  loc_00E3239E: push edi
  loc_00E3239F: mov var_C, esp
  loc_00E323A2: mov var_8, 004025D0h
  loc_00E323A9: mov esi, Me
  loc_00E323AC: mov eax, esi
  loc_00E323AE: and eax, 00000001h
  loc_00E323B1: mov var_4, eax
  loc_00E323B4: and esi, FFFFFFFEh
  loc_00E323B7: push esi
  loc_00E323B8: mov Me, esi
  loc_00E323BB: mov ecx, [esi]
  loc_00E323BD: call [ecx+00000004h]
  loc_00E323C0: xor eax, eax
  loc_00E323C2: lea edx, var_4C
  loc_00E323C5: mov var_4C, eax
  loc_00E323C8: lea ecx, var_24
  loc_00E323CB: mov var_24, eax
  loc_00E323CE: mov var_28, eax
  loc_00E323D1: mov var_2C, eax
  loc_00E323D4: mov var_3C, eax
  loc_00E323D7: mov var_44, 006E15ECh ; "select * from lc"
  loc_00E323DE: mov var_4C, 00000008h
  loc_00E323E5: call [00401204h] ; __vbaVarCopy
  loc_00E323EB: lea edx, var_24
  loc_00E323EE: push edx
  loc_00E323EF: call [00401230h] ; __vbaStrVarCopy
  loc_00E323F5: sub esp, 00000010h
  loc_00E323F8: mov ecx, 00000008h
  loc_00E323FD: mov edx, esp
  loc_00E323FF: mov var_3C, ecx
  loc_00E32402: mov var_34, eax
  loc_00E32405: push 0000000Eh
  loc_00E32407: mov [edx], ecx
  loc_00E32409: mov ecx, var_38
  loc_00E3240C: push esi
  loc_00E3240D: mov [edx+00000004h], ecx
  loc_00E32410: mov ecx, [esi]
  loc_00E32412: mov [edx+00000008h], eax
  loc_00E32415: mov eax, var_30
  loc_00E32418: mov [edx+0000000Ch], eax
  loc_00E3241B: call [ecx+00000324h]
  loc_00E32421: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E32427: lea edx, var_28
  loc_00E3242A: push eax
  loc_00E3242B: push edx
  loc_00E3242C: call edi
  loc_00E3242E: push eax
  loc_00E3242F: call [00401238h] ; __vbaLateIdSt
  loc_00E32435: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E3243B: lea ecx, var_28
  loc_00E3243E: call ebx
  loc_00E32440: lea ecx, var_3C
  loc_00E32443: call [00401024h] ; __vbaFreeVar
  loc_00E32449: mov eax, [esi]
  loc_00E3244B: push 00000000h
  loc_00E3244D: push 00000019h
  loc_00E3244F: push esi
  loc_00E32450: call [eax+00000324h]
  loc_00E32456: lea ecx, var_28
  loc_00E32459: push eax
  loc_00E3245A: push ecx
  loc_00E3245B: call edi
  loc_00E3245D: push eax
  loc_00E3245E: call [00401030h] ; __vbaLateIdCall
  loc_00E32464: add esp, 0000000Ch
  loc_00E32467: lea ecx, var_28
  loc_00E3246A: call ebx
  loc_00E3246C: mov edx, [esi]
  loc_00E3246E: push 006E05D8h
  loc_00E32473: push esi
  loc_00E32474: call [edx+00000324h]
  loc_00E3247A: push eax
  loc_00E3247B: lea eax, var_28
  loc_00E3247E: push eax
  loc_00E3247F: call edi
  loc_00E32481: push eax
  loc_00E32482: call [00401224h] ; __vbaCastObj
  loc_00E32488: mov var_34, eax
  loc_00E3248B: mov var_3C, 0000000Dh
  loc_00E32492: lea ecx, var_3C
  loc_00E32495: push ecx
  loc_00E32496: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3249C: mov ecx, [eax]
  loc_00E3249E: sub esp, 00000010h
  loc_00E324A1: mov edx, esp
  loc_00E324A3: push 00000000h
  loc_00E324A5: push 0000002Ah
  loc_00E324A7: mov [edx], ecx
  loc_00E324A9: mov ecx, [eax+00000004h]
  loc_00E324AC: push esi
  loc_00E324AD: mov [edx+00000004h], ecx
  loc_00E324B0: mov ecx, [eax+00000008h]
  loc_00E324B3: mov eax, [eax+0000000Ch]
  loc_00E324B6: mov [edx+00000008h], ecx
  loc_00E324B9: mov ecx, [esi]
  loc_00E324BB: mov [edx+0000000Ch], eax
  loc_00E324BE: call [ecx+00000328h]
  loc_00E324C4: lea edx, var_2C
  loc_00E324C7: push eax
  loc_00E324C8: push edx
  loc_00E324C9: call edi
  loc_00E324CB: push eax
  loc_00E324CC: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E324D2: lea eax, var_2C
  loc_00E324D5: lea ecx, var_28
  loc_00E324D8: push eax
  loc_00E324D9: push ecx
  loc_00E324DA: push 00000002h
  loc_00E324DC: call [00401048h] ; __vbaFreeObjList
  loc_00E324E2: add esp, 00000028h
  loc_00E324E5: lea ecx, var_3C
  loc_00E324E8: call [00401024h] ; __vbaFreeVar
  loc_00E324EE: mov edx, [esi]
  loc_00E324F0: push 00000000h
  loc_00E324F2: push FFFFFDDAh
  loc_00E324F7: push esi
  loc_00E324F8: call [edx+00000328h]
  loc_00E324FE: push eax
  loc_00E324FF: lea eax, var_28
  loc_00E32502: push eax
  loc_00E32503: call edi
  loc_00E32505: push eax
  loc_00E32506: call [00401030h] ; __vbaLateIdCall
  loc_00E3250C: add esp, 0000000Ch
  loc_00E3250F: lea ecx, var_28
  loc_00E32512: call ebx
  loc_00E32514: mov var_4, 00000000h
  loc_00E3251B: push 00E32549h
  loc_00E32520: jmp 00E3253Fh
  loc_00E32522: lea ecx, var_2C
  loc_00E32525: lea edx, var_28
  loc_00E32528: push ecx
  loc_00E32529: push edx
  loc_00E3252A: push 00000002h
  loc_00E3252C: call [00401048h] ; __vbaFreeObjList
  loc_00E32532: add esp, 0000000Ch
  loc_00E32535: lea ecx, var_3C
  loc_00E32538: call [00401024h] ; __vbaFreeVar
  loc_00E3253E: ret
  loc_00E3253F: lea ecx, var_24
  loc_00E32542: call [00401024h] ; __vbaFreeVar
  loc_00E32548: ret
  loc_00E32549: mov eax, Me
  loc_00E3254C: push eax
  loc_00E3254D: mov ecx, [eax]
  loc_00E3254F: call [ecx+00000008h]
  loc_00E32552: mov eax, var_4
  loc_00E32555: mov ecx, var_14
  loc_00E32558: pop edi
  loc_00E32559: pop esi
  loc_00E3255A: mov fs:[00000000h], ecx
  loc_00E32561: pop ebx
  loc_00E32562: mov esp, ebp
  loc_00E32564: pop ebp
  loc_00E32565: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E32130
  loc_00E32130: push ebp
  loc_00E32131: mov ebp, esp
  loc_00E32133: sub esp, 0000000Ch
  loc_00E32136: push 00402836h ; __vbaExceptHandler
  loc_00E3213B: mov eax, fs:[00000000h]
  loc_00E32141: push eax
  loc_00E32142: mov fs:[00000000h], esp
  loc_00E32149: sub esp, 0000005Ch
  loc_00E3214C: push ebx
  loc_00E3214D: push esi
  loc_00E3214E: push edi
  loc_00E3214F: mov var_C, esp
  loc_00E32152: mov var_8, 004025C0h
  loc_00E32159: mov esi, Me
  loc_00E3215C: mov eax, esi
  loc_00E3215E: and eax, 00000001h
  loc_00E32161: mov var_4, eax
  loc_00E32164: and esi, FFFFFFFEh
  loc_00E32167: push esi
  loc_00E32168: mov Me, esi
  loc_00E3216B: mov ecx, [esi]
  loc_00E3216D: call [ecx+00000004h]
  loc_00E32170: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32176: xor eax, eax
  loc_00E32178: mov var_18, eax
  loc_00E3217B: mov var_4C, eax
  loc_00E3217E: mov var_50, eax
  loc_00E32181: mov edx, [esi]
  loc_00E32183: lea eax, var_4C
  loc_00E32186: push eax
  loc_00E32187: push esi
  loc_00E32188: call [edx+00000070h]
  loc_00E3218B: test eax, eax
  loc_00E3218D: fnclex
  loc_00E3218F: jge 00E3219Ch
  loc_00E32191: push 00000070h
  loc_00E32193: push 006E00E8h
  loc_00E32198: push esi
  loc_00E32199: push eax
  loc_00E3219A: call ebx
  loc_00E3219C: fld real4 ptr var_4C
  loc_00E3219F: fadd st0, real4 ptr [00401298h]
  loc_00E321A5: mov ecx, [esi]
  loc_00E321A7: push ecx
  loc_00E321A8: fnstsw ax
  loc_00E321AA: test al, 0Dh
  loc_00E321AC: jnz 00E32370h
  loc_00E321B2: fstp real4 ptr [esp]
  loc_00E321B5: push esi
  loc_00E321B6: call [ecx+00000074h]
  loc_00E321B9: test eax, eax
  loc_00E321BB: fnclex
  loc_00E321BD: jge 00E321CAh
  loc_00E321BF: push 00000074h
  loc_00E321C1: push 006E00E8h
  loc_00E321C6: push esi
  loc_00E321C7: push eax
  loc_00E321C8: call ebx
  loc_00E321CA: mov edx, [esi]
  loc_00E321CC: lea eax, var_4C
  loc_00E321CF: push eax
  loc_00E321D0: push esi
  loc_00E321D1: call [edx+00000070h]
  loc_00E321D4: test eax, eax
  loc_00E321D6: fnclex
  loc_00E321D8: jge 00E321E5h
  loc_00E321DA: push 00000070h
  loc_00E321DC: push 006E00E8h
  loc_00E321E1: push esi
  loc_00E321E2: push eax
  loc_00E321E3: call ebx
  loc_00E321E5: mov ecx, [esi]
  loc_00E321E7: lea edx, var_50
  loc_00E321EA: push edx
  loc_00E321EB: push esi
  loc_00E321EC: call [ecx+00000078h]
  loc_00E321EF: test eax, eax
  loc_00E321F1: fnclex
  loc_00E321F3: jge 00E32200h
  loc_00E321F5: push 00000078h
  loc_00E321F7: push 006E00E8h
  loc_00E321FC: push esi
  loc_00E321FD: push eax
  loc_00E321FE: call ebx
  loc_00E32200: sub esp, 00000010h
  loc_00E32203: mov ecx, 0000000Ah
  loc_00E32208: mov ebx, esp
  loc_00E3220A: mov eax, 80020004h
  loc_00E3220F: mov edx, eax
  loc_00E32211: sub esp, 00000010h
  loc_00E32214: mov [ebx], ecx
  loc_00E32216: mov ecx, var_44
  loc_00E32219: mov edi, [esi]
  loc_00E3221B: mov [ebx+00000004h], ecx
  loc_00E3221E: mov ecx, esp
  loc_00E32220: sub esp, 00000010h
  loc_00E32223: mov [ebx+00000008h], eax
  loc_00E32226: mov eax, var_3C
  loc_00E32229: mov [ebx+0000000Ch], eax
  loc_00E3222C: mov eax, 0000000Ah
  loc_00E32231: mov [ecx], eax
  loc_00E32233: mov eax, var_34
  loc_00E32236: mov [ecx+00000004h], eax
  loc_00E32239: mov eax, 00000004h
  loc_00E3223E: mov [ecx+00000008h], edx
  loc_00E32241: mov edx, var_2C
  loc_00E32244: mov [ecx+0000000Ch], edx
  loc_00E32247: mov edx, var_24
  loc_00E3224A: mov ecx, esp
  loc_00E3224C: mov [ecx], eax
  loc_00E3224E: mov eax, var_50
  loc_00E32251: mov [ecx+00000004h], edx
  loc_00E32254: mov [ecx+00000008h], eax
  loc_00E32257: mov eax, var_1C
  loc_00E3225A: mov [ecx+0000000Ch], eax
  loc_00E3225D: mov ecx, var_4C
  loc_00E32260: push ecx
  loc_00E32261: push esi
  loc_00E32262: call [edi+000002A4h]
  loc_00E32268: test eax, eax
  loc_00E3226A: fnclex
  loc_00E3226C: jge 00E32284h
  loc_00E3226E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32274: push 000002A4h
  loc_00E32279: push 006E00E8h
  loc_00E3227E: push esi
  loc_00E3227F: push eax
  loc_00E32280: call ebx
  loc_00E32282: jmp 00E3228Ah
  loc_00E32284: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3228A: call [004010BCh] ; rtcDoEvents
  loc_00E32290: mov edx, [esi]
  loc_00E32292: lea eax, var_4C
  loc_00E32295: push eax
  loc_00E32296: push esi
  loc_00E32297: call [edx+00000070h]
  loc_00E3229A: test eax, eax
  loc_00E3229C: fnclex
  loc_00E3229E: jge 00E322ABh
  loc_00E322A0: push 00000070h
  loc_00E322A2: push 006E00E8h
  loc_00E322A7: push esi
  loc_00E322A8: push eax
  loc_00E322A9: call ebx
  loc_00E322AB: mov eax, [00E3D9CCh]
  loc_00E322B0: test eax, eax
  loc_00E322B2: jnz 00E322C4h
  loc_00E322B4: push 00E3D9CCh
  loc_00E322B9: push 006DC960h
  loc_00E322BE: call [004011A0h] ; __vbaNew2
  loc_00E322C4: mov edi, [00E3D9CCh]
  loc_00E322CA: lea edx, var_18
  loc_00E322CD: push edx
  loc_00E322CE: push edi
  loc_00E322CF: mov ecx, [edi]
  loc_00E322D1: call [ecx+00000018h]
  loc_00E322D4: test eax, eax
  loc_00E322D6: fnclex
  loc_00E322D8: jge 00E322E5h
  loc_00E322DA: push 00000018h
  loc_00E322DC: push 006DC950h
  loc_00E322E1: push edi
  loc_00E322E2: push eax
  loc_00E322E3: call ebx
  loc_00E322E5: mov eax, var_18
  loc_00E322E8: lea edx, var_50
  loc_00E322EB: push edx
  loc_00E322EC: push eax
  loc_00E322ED: mov ecx, [eax]
  loc_00E322EF: mov edi, eax
  loc_00E322F1: call [ecx+00000098h]
  loc_00E322F7: test eax, eax
  loc_00E322F9: fnclex
  loc_00E322FB: jge 00E3230Bh
  loc_00E322FD: push 00000098h
  loc_00E32302: push 006DCAF0h
  loc_00E32307: push edi
  loc_00E32308: push eax
  loc_00E32309: call ebx
  loc_00E3230B: fld real4 ptr var_4C
  loc_00E3230E: fcomp real4 ptr var_50
  loc_00E32311: fnstsw ax
  loc_00E32313: test ah, 41h
  loc_00E32316: jz 00E3231Fh
  loc_00E32318: mov eax, 00000001h
  loc_00E3231D: jmp 00E32321h
  loc_00E3231F: xor eax, eax
  loc_00E32321: neg eax
  loc_00E32323: lea ecx, var_18
  loc_00E32326: mov edi, eax
  loc_00E32328: call [00401254h] ; __vbaFreeObj
  loc_00E3232E: test di, di
  loc_00E32331: jnz 00E32181h
  loc_00E32337: mov var_4, 00000000h
  loc_00E3233E: fwait
  loc_00E3233F: push 00E32351h
  loc_00E32344: jmp 00E32350h
  loc_00E32346: lea ecx, var_18
  loc_00E32349: call [00401254h] ; __vbaFreeObj
  loc_00E3234F: ret
  loc_00E32350: ret
  loc_00E32351: mov eax, Me
  loc_00E32354: push eax
  loc_00E32355: mov ecx, [eax]
  loc_00E32357: call [ecx+00000008h]
  loc_00E3235A: mov eax, var_4
  loc_00E3235D: mov ecx, var_14
  loc_00E32360: pop edi
  loc_00E32361: pop esi
  loc_00E32362: mov fs:[00000000h], ecx
  loc_00E32369: pop ebx
  loc_00E3236A: mov esp, ebp
  loc_00E3236C: pop ebp
  loc_00E3236D: retn 0008h
End Sub

Private Sub back_Click() 'E30540
  loc_00E30540: push ebp
  loc_00E30541: mov ebp, esp
  loc_00E30543: sub esp, 0000000Ch
  loc_00E30546: push 00402836h ; __vbaExceptHandler
  loc_00E3054B: mov eax, fs:[00000000h]
  loc_00E30551: push eax
  loc_00E30552: mov fs:[00000000h], esp
  loc_00E30559: sub esp, 00000018h
  loc_00E3055C: push ebx
  loc_00E3055D: push esi
  loc_00E3055E: push edi
  loc_00E3055F: mov var_C, esp
  loc_00E30562: mov var_8, 004025A0h
  loc_00E30569: mov edi, Me
  loc_00E3056C: mov eax, edi
  loc_00E3056E: and eax, 00000001h
  loc_00E30571: mov var_4, eax
  loc_00E30574: and edi, FFFFFFFEh
  loc_00E30577: push edi
  loc_00E30578: mov Me, edi
  loc_00E3057B: mov ecx, [edi]
  loc_00E3057D: call [ecx+00000004h]
  loc_00E30580: mov eax, [00E3D9CCh]
  loc_00E30585: xor ebx, ebx
  loc_00E30587: cmp eax, ebx
  loc_00E30589: mov var_18, ebx
  loc_00E3058C: jnz 00E3059Eh
  loc_00E3058E: push 00E3D9CCh
  loc_00E30593: push 006DC960h
  loc_00E30598: call [004011A0h] ; __vbaNew2
  loc_00E3059E: mov esi, [00E3D9CCh]
  loc_00E305A4: lea eax, var_18
  loc_00E305A7: push edi
  loc_00E305A8: push eax
  loc_00E305A9: mov edx, [esi]
  loc_00E305AB: mov var_2C, edx
  loc_00E305AE: call [004010B4h] ; __vbaObjSetAddref
  loc_00E305B4: mov ecx, var_2C
  loc_00E305B7: push eax
  loc_00E305B8: push esi
  loc_00E305B9: call [ecx+00000010h]
  loc_00E305BC: cmp eax, ebx
  loc_00E305BE: fnclex
  loc_00E305C0: jge 00E305D1h
  loc_00E305C2: push 00000010h
  loc_00E305C4: push 006DC950h
  loc_00E305C9: push esi
  loc_00E305CA: push eax
  loc_00E305CB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E305D1: lea ecx, var_18
  loc_00E305D4: call [00401254h] ; __vbaFreeObj
  loc_00E305DA: mov var_4, ebx
  loc_00E305DD: push 00E305EFh
  loc_00E305E2: jmp 00E305EEh
  loc_00E305E4: lea ecx, var_18
  loc_00E305E7: call [00401254h] ; __vbaFreeObj
  loc_00E305ED: ret
  loc_00E305EE: ret
  loc_00E305EF: mov eax, Me
  loc_00E305F2: push eax
  loc_00E305F3: mov edx, [eax]
  loc_00E305F5: call [edx+00000008h]
  loc_00E305F8: mov eax, var_4
  loc_00E305FB: mov ecx, var_14
  loc_00E305FE: pop edi
  loc_00E305FF: pop esi
  loc_00E30600: mov fs:[00000000h], ecx
  loc_00E30607: pop ebx
  loc_00E30608: mov esp, ebp
  loc_00E3060A: pop ebp
  loc_00E3060B: retn 0004h
End Sub

Private Sub dgpen_Click() 'E30610
  loc_00E30610: push ebp
  loc_00E30611: mov ebp, esp
  loc_00E30613: sub esp, 0000000Ch
  loc_00E30616: push 00402836h ; __vbaExceptHandler
  loc_00E3061B: mov eax, fs:[00000000h]
  loc_00E30621: push eax
  loc_00E30622: mov fs:[00000000h], esp
  loc_00E30629: sub esp, 00000100h
  loc_00E3062F: push ebx
  loc_00E30630: push esi
  loc_00E30631: push edi
  loc_00E30632: mov var_C, esp
  loc_00E30635: mov var_8, 004025B0h
  loc_00E3063C: mov ebx, Me
  loc_00E3063F: mov eax, ebx
  loc_00E30641: and eax, 00000001h
  loc_00E30644: mov var_4, eax
  loc_00E30647: and ebx, FFFFFFFEh
  loc_00E3064A: push ebx
  loc_00E3064B: mov Me, ebx
  loc_00E3064E: mov ecx, [ebx]
  loc_00E30650: call [ecx+00000004h]
  loc_00E30653: mov edx, [ebx]
  loc_00E30655: xor eax, eax
  loc_00E30657: push 006DCBD8h
  loc_00E3065C: push eax
  loc_00E3065D: push 00000018h
  loc_00E3065F: push ebx
  loc_00E30660: mov var_24, eax
  loc_00E30663: mov var_28, eax
  loc_00E30666: mov var_2C, eax
  loc_00E30669: mov var_30, eax
  loc_00E3066C: mov var_34, eax
  loc_00E3066F: mov var_38, eax
  loc_00E30672: mov var_3C, eax
  loc_00E30675: mov var_40, eax
  loc_00E30678: mov var_50, eax
  loc_00E3067B: mov var_60, eax
  loc_00E3067E: mov var_70, eax
  loc_00E30681: mov var_80, eax
  loc_00E30684: mov var_90, eax
  loc_00E3068A: mov var_A0, eax
  loc_00E30690: call [edx+00000324h]
  loc_00E30696: mov esi, [004010ACh] ; __vbaObjSet
  loc_00E3069C: push eax
  loc_00E3069D: lea eax, var_30
  loc_00E306A0: push eax
  loc_00E306A1: call __vbaObjSet
  loc_00E306A3: lea ecx, var_50
  loc_00E306A6: push eax
  loc_00E306A7: push ecx
  loc_00E306A8: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E306AE: add esp, 00000010h
  loc_00E306B1: push eax
  loc_00E306B2: call [00401120h] ; __vbaCastObjVar
  loc_00E306B8: lea edx, var_34
  loc_00E306BB: push eax
  loc_00E306BC: push edx
  loc_00E306BD: call __vbaObjSet
  loc_00E306BF: mov edi, eax
  loc_00E306C1: lea ecx, var_38
  loc_00E306C4: push ecx
  loc_00E306C5: push edi
  loc_00E306C6: mov eax, [edi]
  loc_00E306C8: call [eax+00000054h]
  loc_00E306CB: test eax, eax
  loc_00E306CD: fnclex
  loc_00E306CF: jge 00E306E0h
  loc_00E306D1: push 00000054h
  loc_00E306D3: push 006DCBE8h
  loc_00E306D8: push edi
  loc_00E306D9: push eax
  loc_00E306DA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E306E0: lea edi, var_3C
  loc_00E306E3: mov eax, var_38
  loc_00E306E6: push edi
  loc_00E306E7: mov ecx, 00000002h
  loc_00E306EC: sub esp, 00000010h
  loc_00E306EF: mov var_90, ecx
  loc_00E306F5: mov edi, esp
  loc_00E306F7: mov var_88, 00000005h
  loc_00E30701: mov edx, [eax]
  loc_00E30703: push eax
  loc_00E30704: mov [edi], ecx
  loc_00E30706: mov ecx, var_8C
  loc_00E3070C: mov var_CC, eax
  loc_00E30712: mov [edi+00000004h], ecx
  loc_00E30715: mov ecx, var_88
  loc_00E3071B: mov [edi+00000008h], ecx
  loc_00E3071E: mov ecx, var_84
  loc_00E30724: mov [edi+0000000Ch], ecx
  loc_00E30727: call [edx+00000028h]
  loc_00E3072A: test eax, eax
  loc_00E3072C: fnclex
  loc_00E3072E: jge 00E30745h
  loc_00E30730: mov edx, var_CC
  loc_00E30736: push 00000028h
  loc_00E30738: push 006E09E8h
  loc_00E3073D: push edx
  loc_00E3073E: push eax
  loc_00E3073F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30745: mov eax, var_3C
  loc_00E30748: lea edx, var_60
  loc_00E3074B: push edx
  loc_00E3074C: push eax
  loc_00E3074D: mov ecx, [eax]
  loc_00E3074F: mov edi, eax
  loc_00E30751: call [ecx+00000034h]
  loc_00E30754: test eax, eax
  loc_00E30756: fnclex
  loc_00E30758: jge 00E30769h
  loc_00E3075A: push 00000034h
  loc_00E3075C: push 006E09F8h
  loc_00E30761: push edi
  loc_00E30762: push eax
  loc_00E30763: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30769: lea eax, var_60
  loc_00E3076C: lea ecx, var_A0
  loc_00E30772: push eax
  loc_00E30773: push ecx
  loc_00E30774: mov var_98, 006E1728h ; "Lunas"
  loc_00E3077E: mov var_A0, 00008008h
  loc_00E30788: call [00401108h] ; __vbaVarTstEq
  loc_00E3078E: mov di, ax
  loc_00E30791: lea edx, var_3C
  loc_00E30794: lea eax, var_38
  loc_00E30797: push edx
  loc_00E30798: lea ecx, var_34
  loc_00E3079B: push eax
  loc_00E3079C: lea edx, var_30
  loc_00E3079F: push ecx
  loc_00E307A0: push edx
  loc_00E307A1: push 00000004h
  loc_00E307A3: call [00401048h] ; __vbaFreeObjList
  loc_00E307A9: lea eax, var_60
  loc_00E307AC: lea ecx, var_50
  loc_00E307AF: push eax
  loc_00E307B0: push ecx
  loc_00E307B1: push 00000002h
  loc_00E307B3: call [00401038h] ; __vbaFreeVarList
  loc_00E307B9: add esp, 00000020h
  loc_00E307BC: test di, di
  loc_00E307BF: jnz 00E30900h
  loc_00E307C5: mov ecx, [ebx]
  loc_00E307C7: push 006DCBD8h
  loc_00E307CC: push 00000000h
  loc_00E307CE: push 00000018h
  loc_00E307D0: push ebx
  loc_00E307D1: call [ecx+00000324h]
  loc_00E307D7: lea edx, var_30
  loc_00E307DA: push eax
  loc_00E307DB: push edx
  loc_00E307DC: call __vbaObjSet
  loc_00E307DE: push eax
  loc_00E307DF: lea eax, var_50
  loc_00E307E2: push eax
  loc_00E307E3: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E307E9: add esp, 00000010h
  loc_00E307EC: push eax
  loc_00E307ED: call [00401120h] ; __vbaCastObjVar
  loc_00E307F3: lea ecx, var_34
  loc_00E307F6: push eax
  loc_00E307F7: push ecx
  loc_00E307F8: call __vbaObjSet
  loc_00E307FA: mov edi, eax
  loc_00E307FC: lea eax, var_38
  loc_00E307FF: push eax
  loc_00E30800: push edi
  loc_00E30801: mov edx, [edi]
  loc_00E30803: call [edx+00000054h]
  loc_00E30806: test eax, eax
  loc_00E30808: fnclex
  loc_00E3080A: jge 00E3081Bh
  loc_00E3080C: push 00000054h
  loc_00E3080E: push 006DCBE8h
  loc_00E30813: push edi
  loc_00E30814: push eax
  loc_00E30815: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3081B: lea edi, var_3C
  loc_00E3081E: mov eax, var_38
  loc_00E30821: push edi
  loc_00E30822: mov ecx, 00000002h
  loc_00E30827: sub esp, 00000010h
  loc_00E3082A: mov var_90, ecx
  loc_00E30830: mov edi, esp
  loc_00E30832: mov var_88, 00000005h
  loc_00E3083C: mov edx, [eax]
  loc_00E3083E: push eax
  loc_00E3083F: mov [edi], ecx
  loc_00E30841: mov ecx, var_8C
  loc_00E30847: mov var_CC, eax
  loc_00E3084D: mov [edi+00000004h], ecx
  loc_00E30850: mov ecx, var_88
  loc_00E30856: mov [edi+00000008h], ecx
  loc_00E30859: mov ecx, var_84
  loc_00E3085F: mov [edi+0000000Ch], ecx
  loc_00E30862: call [edx+00000028h]
  loc_00E30865: test eax, eax
  loc_00E30867: fnclex
  loc_00E30869: jge 00E30880h
  loc_00E3086B: mov edx, var_CC
  loc_00E30871: push 00000028h
  loc_00E30873: push 006E09E8h
  loc_00E30878: push edx
  loc_00E30879: push eax
  loc_00E3087A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30880: mov eax, var_3C
  loc_00E30883: lea edx, var_60
  loc_00E30886: push edx
  loc_00E30887: push eax
  loc_00E30888: mov ecx, [eax]
  loc_00E3088A: mov edi, eax
  loc_00E3088C: call [ecx+00000034h]
  loc_00E3088F: test eax, eax
  loc_00E30891: fnclex
  loc_00E30893: jge 00E308A4h
  loc_00E30895: push 00000034h
  loc_00E30897: push 006E09F8h
  loc_00E3089C: push edi
  loc_00E3089D: push eax
  loc_00E3089E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E308A4: lea eax, var_60
  loc_00E308A7: lea ecx, var_A0
  loc_00E308AD: push eax
  loc_00E308AE: push ecx
  loc_00E308AF: mov var_98, 006E1C10h ; "Mencicil (Lunas)"
  loc_00E308B9: mov var_A0, 00008008h
  loc_00E308C3: call [00401108h] ; __vbaVarTstEq
  loc_00E308C9: mov di, ax
  loc_00E308CC: lea edx, var_3C
  loc_00E308CF: lea eax, var_38
  loc_00E308D2: push edx
  loc_00E308D3: lea ecx, var_34
  loc_00E308D6: push eax
  loc_00E308D7: lea edx, var_30
  loc_00E308DA: push ecx
  loc_00E308DB: push edx
  loc_00E308DC: push 00000004h
  loc_00E308DE: call [00401048h] ; __vbaFreeObjList
  loc_00E308E4: lea eax, var_60
  loc_00E308E7: lea ecx, var_50
  loc_00E308EA: push eax
  loc_00E308EB: push ecx
  loc_00E308EC: push 00000002h
  loc_00E308EE: call [00401038h] ; __vbaFreeVarList
  loc_00E308F4: add esp, 00000020h
  loc_00E308F7: test di, di
  loc_00E308FA: jz 00E3098Fh
  loc_00E30900: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E30906: mov ecx, 80020004h
  loc_00E3090B: mov var_78, ecx
  loc_00E3090E: mov eax, 0000000Ah
  loc_00E30913: mov var_68, ecx
  loc_00E30916: mov edi, 00000008h
  loc_00E3091B: lea edx, var_A0
  loc_00E30921: lea ecx, var_60
  loc_00E30924: mov var_80, eax
  loc_00E30927: mov var_70, eax
  loc_00E3092A: mov var_98, 006E16F0h ; "SMK RK BT PS"
  loc_00E30934: mov var_A0, edi
  loc_00E3093A: call __vbaVarDup
  loc_00E3093C: lea edx, var_90
  loc_00E30942: lea ecx, var_50
  loc_00E30945: mov var_88, 006E1B9Ch ; "Data calon peserta didik telah melunasi Tunggakan nya !"
  loc_00E3094F: mov var_90, edi
  loc_00E30955: call __vbaVarDup
  loc_00E30957: lea edx, var_80
  loc_00E3095A: lea eax, var_70
  loc_00E3095D: push edx
  loc_00E3095E: lea ecx, var_60
  loc_00E30961: push eax
  loc_00E30962: push ecx
  loc_00E30963: lea edx, var_50
  loc_00E30966: push 00000040h
  loc_00E30968: push edx
  loc_00E30969: call [004010A8h] ; rtcMsgBox
  loc_00E3096F: lea eax, var_80
  loc_00E30972: lea ecx, var_70
  loc_00E30975: push eax
  loc_00E30976: lea edx, var_60
  loc_00E30979: push ecx
  loc_00E3097A: lea eax, var_50
  loc_00E3097D: push edx
  loc_00E3097E: push eax
  loc_00E3097F: push 00000004h
  loc_00E30981: call [00401038h] ; __vbaFreeVarList
  loc_00E30987: add esp, 00000014h
  loc_00E3098A: jmp 00E320ABh
  loc_00E3098F: mov eax, [00E3D0D8h]
  loc_00E30994: test eax, eax
  loc_00E30996: jnz 00E309ADh
  loc_00E30998: push 00E3D0D8h
  loc_00E3099D: push 006D300Ch
  loc_00E309A2: call [004011A0h] ; __vbaNew2
  loc_00E309A8: mov eax, [00E3D0D8h]
  loc_00E309AD: mov ecx, [eax]
  loc_00E309AF: push eax
  loc_00E309B0: call [ecx+00000430h]
  loc_00E309B6: lea edx, var_40
  loc_00E309B9: push eax
  loc_00E309BA: push edx
  loc_00E309BB: call __vbaVarDup
  loc_00E309BD: push 006DCBD8h
  loc_00E309C2: mov var_DC, eax
  loc_00E309C8: mov eax, [ebx]
  loc_00E309CA: push 00000000h
  loc_00E309CC: push 00000018h
  loc_00E309CE: push ebx
  loc_00E309CF: call [eax+00000324h]
  loc_00E309D5: lea ecx, var_30
  loc_00E309D8: push eax
  loc_00E309D9: push ecx
  loc_00E309DA: call __vbaVarDup
  loc_00E309DC: lea edx, var_50
  loc_00E309DF: push eax
  loc_00E309E0: push edx
  loc_00E309E1: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E309E7: add esp, 00000010h
  loc_00E309EA: push eax
  loc_00E309EB: call [00401120h] ; __vbaCastObjVar
  loc_00E309F1: push eax
  loc_00E309F2: lea eax, var_34
  loc_00E309F5: push eax
  loc_00E309F6: call __vbaVarDup
  loc_00E309F8: mov edi, eax
  loc_00E309FA: lea edx, var_38
  loc_00E309FD: push edx
  loc_00E309FE: push edi
  loc_00E309FF: mov ecx, [edi]
  loc_00E30A01: call [ecx+00000054h]
  loc_00E30A04: test eax, eax
  loc_00E30A06: fnclex
  loc_00E30A08: jge 00E30A19h
  loc_00E30A0A: push 00000054h
  loc_00E30A0C: push 006DCBE8h
  loc_00E30A11: push edi
  loc_00E30A12: push eax
  loc_00E30A13: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30A19: lea edi, var_3C
  loc_00E30A1C: mov eax, var_38
  loc_00E30A1F: push edi
  loc_00E30A20: mov ecx, 00000002h
  loc_00E30A25: sub esp, 00000010h
  loc_00E30A28: mov var_90, ecx
  loc_00E30A2E: mov edi, esp
  loc_00E30A30: mov var_88, 00000000h
  loc_00E30A3A: mov edx, [eax]
  loc_00E30A3C: push eax
  loc_00E30A3D: mov [edi], ecx
  loc_00E30A3F: mov ecx, var_8C
  loc_00E30A45: mov var_CC, eax
  loc_00E30A4B: mov [edi+00000004h], ecx
  loc_00E30A4E: mov ecx, var_88
  loc_00E30A54: mov [edi+00000008h], ecx
  loc_00E30A57: mov ecx, var_84
  loc_00E30A5D: mov [edi+0000000Ch], ecx
  loc_00E30A60: call [edx+00000028h]
  loc_00E30A63: test eax, eax
  loc_00E30A65: fnclex
  loc_00E30A67: jge 00E30A7Eh
  loc_00E30A69: mov edx, var_CC
  loc_00E30A6F: push 00000028h
  loc_00E30A71: push 006E09E8h
  loc_00E30A76: push edx
  loc_00E30A77: push eax
  loc_00E30A78: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30A7E: mov eax, var_3C
  loc_00E30A81: lea edx, var_60
  loc_00E30A84: push edx
  loc_00E30A85: push eax
  loc_00E30A86: mov ecx, [eax]
  loc_00E30A88: mov edi, eax
  loc_00E30A8A: call [ecx+00000034h]
  loc_00E30A8D: test eax, eax
  loc_00E30A8F: fnclex
  loc_00E30A91: jge 00E30AA2h
  loc_00E30A93: push 00000034h
  loc_00E30A95: push 006E09F8h
  loc_00E30A9A: push edi
  loc_00E30A9B: push eax
  loc_00E30A9C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30AA2: mov eax, var_DC
  loc_00E30AA8: lea ecx, var_60
  loc_00E30AAB: push ecx
  loc_00E30AAC: mov edi, [eax]
  loc_00E30AAE: call [00401034h] ; __vbaStrVarMove
  loc_00E30AB4: mov edx, eax
  loc_00E30AB6: lea ecx, var_28
  loc_00E30AB9: call [00401228h] ; __vbaStrMove
  loc_00E30ABF: mov edx, edi
  loc_00E30AC1: mov edi, var_DC
  loc_00E30AC7: push eax
  loc_00E30AC8: push edi
  loc_00E30AC9: call [edx+00000054h]
  loc_00E30ACC: test eax, eax
  loc_00E30ACE: fnclex
  loc_00E30AD0: jge 00E30AE1h
  loc_00E30AD2: push 00000054h
  loc_00E30AD4: push 006E0410h
  loc_00E30AD9: push edi
  loc_00E30ADA: push eax
  loc_00E30ADB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30AE1: lea ecx, var_28
  loc_00E30AE4: call [00401258h] ; __vbaFreeStr
  loc_00E30AEA: lea eax, var_40
  loc_00E30AED: lea ecx, var_3C
  loc_00E30AF0: push eax
  loc_00E30AF1: lea edx, var_38
  loc_00E30AF4: push ecx
  loc_00E30AF5: lea eax, var_34
  loc_00E30AF8: push edx
  loc_00E30AF9: lea ecx, var_30
  loc_00E30AFC: push eax
  loc_00E30AFD: push ecx
  loc_00E30AFE: push 00000005h
  loc_00E30B00: call [00401048h] ; __vbaFreeObjList
  loc_00E30B06: lea edx, var_60
  loc_00E30B09: lea eax, var_50
  loc_00E30B0C: push edx
  loc_00E30B0D: push eax
  loc_00E30B0E: push 00000002h
  loc_00E30B10: call [00401038h] ; __vbaFreeVarList
  loc_00E30B16: mov eax, [00E3D0D8h]
  loc_00E30B1B: add esp, 00000024h
  loc_00E30B1E: test eax, eax
  loc_00E30B20: jnz 00E30B37h
  loc_00E30B22: push 00E3D0D8h
  loc_00E30B27: push 006D300Ch
  loc_00E30B2C: call [004011A0h] ; __vbaNew2
  loc_00E30B32: mov eax, [00E3D0D8h]
  loc_00E30B37: mov ecx, [eax]
  loc_00E30B39: push eax
  loc_00E30B3A: call [ecx+00000434h]
  loc_00E30B40: lea edx, var_40
  loc_00E30B43: push eax
  loc_00E30B44: push edx
  loc_00E30B45: call __vbaVarDup
  loc_00E30B47: push 006DCBD8h
  loc_00E30B4C: mov var_DC, eax
  loc_00E30B52: mov eax, [ebx]
  loc_00E30B54: push 00000000h
  loc_00E30B56: push 00000018h
  loc_00E30B58: push ebx
  loc_00E30B59: call [eax+00000324h]
  loc_00E30B5F: lea ecx, var_30
  loc_00E30B62: push eax
  loc_00E30B63: push ecx
  loc_00E30B64: call __vbaVarDup
  loc_00E30B66: lea edx, var_50
  loc_00E30B69: push eax
  loc_00E30B6A: push edx
  loc_00E30B6B: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E30B71: add esp, 00000010h
  loc_00E30B74: push eax
  loc_00E30B75: call [00401120h] ; __vbaCastObjVar
  loc_00E30B7B: push eax
  loc_00E30B7C: lea eax, var_34
  loc_00E30B7F: push eax
  loc_00E30B80: call __vbaVarDup
  loc_00E30B82: mov edi, eax
  loc_00E30B84: lea edx, var_38
  loc_00E30B87: push edx
  loc_00E30B88: push edi
  loc_00E30B89: mov ecx, [edi]
  loc_00E30B8B: call [ecx+00000054h]
  loc_00E30B8E: test eax, eax
  loc_00E30B90: fnclex
  loc_00E30B92: jge 00E30BA3h
  loc_00E30B94: push 00000054h
  loc_00E30B96: push 006DCBE8h
  loc_00E30B9B: push edi
  loc_00E30B9C: push eax
  loc_00E30B9D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30BA3: lea edi, var_3C
  loc_00E30BA6: mov eax, var_38
  loc_00E30BA9: push edi
  loc_00E30BAA: mov ecx, 00000002h
  loc_00E30BAF: sub esp, 00000010h
  loc_00E30BB2: mov var_90, ecx
  loc_00E30BB8: mov edi, esp
  loc_00E30BBA: mov var_88, 00000001h
  loc_00E30BC4: mov edx, [eax]
  loc_00E30BC6: push eax
  loc_00E30BC7: mov [edi], ecx
  loc_00E30BC9: mov ecx, var_8C
  loc_00E30BCF: mov var_CC, eax
  loc_00E30BD5: mov [edi+00000004h], ecx
  loc_00E30BD8: mov ecx, var_88
  loc_00E30BDE: mov [edi+00000008h], ecx
  loc_00E30BE1: mov ecx, var_84
  loc_00E30BE7: mov [edi+0000000Ch], ecx
  loc_00E30BEA: call [edx+00000028h]
  loc_00E30BED: test eax, eax
  loc_00E30BEF: fnclex
  loc_00E30BF1: jge 00E30C08h
  loc_00E30BF3: mov edx, var_CC
  loc_00E30BF9: push 00000028h
  loc_00E30BFB: push 006E09E8h
  loc_00E30C00: push edx
  loc_00E30C01: push eax
  loc_00E30C02: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30C08: mov eax, var_3C
  loc_00E30C0B: lea edx, var_60
  loc_00E30C0E: push edx
  loc_00E30C0F: push eax
  loc_00E30C10: mov ecx, [eax]
  loc_00E30C12: mov edi, eax
  loc_00E30C14: call [ecx+00000034h]
  loc_00E30C17: test eax, eax
  loc_00E30C19: fnclex
  loc_00E30C1B: jge 00E30C2Ch
  loc_00E30C1D: push 00000034h
  loc_00E30C1F: push 006E09F8h
  loc_00E30C24: push edi
  loc_00E30C25: push eax
  loc_00E30C26: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30C2C: mov eax, var_DC
  loc_00E30C32: lea ecx, var_60
  loc_00E30C35: push ecx
  loc_00E30C36: mov edi, [eax]
  loc_00E30C38: call [00401034h] ; __vbaStrVarMove
  loc_00E30C3E: mov edx, eax
  loc_00E30C40: lea ecx, var_28
  loc_00E30C43: call [00401228h] ; __vbaStrMove
  loc_00E30C49: mov edx, edi
  loc_00E30C4B: mov edi, var_DC
  loc_00E30C51: push eax
  loc_00E30C52: push edi
  loc_00E30C53: call [edx+00000054h]
  loc_00E30C56: test eax, eax
  loc_00E30C58: fnclex
  loc_00E30C5A: jge 00E30C6Bh
  loc_00E30C5C: push 00000054h
  loc_00E30C5E: push 006E0410h
  loc_00E30C63: push edi
  loc_00E30C64: push eax
  loc_00E30C65: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30C6B: lea ecx, var_28
  loc_00E30C6E: call [00401258h] ; __vbaFreeStr
  loc_00E30C74: lea eax, var_40
  loc_00E30C77: lea ecx, var_3C
  loc_00E30C7A: push eax
  loc_00E30C7B: lea edx, var_38
  loc_00E30C7E: push ecx
  loc_00E30C7F: lea eax, var_34
  loc_00E30C82: push edx
  loc_00E30C83: lea ecx, var_30
  loc_00E30C86: push eax
  loc_00E30C87: push ecx
  loc_00E30C88: push 00000005h
  loc_00E30C8A: call [00401048h] ; __vbaFreeObjList
  loc_00E30C90: lea edx, var_60
  loc_00E30C93: lea eax, var_50
  loc_00E30C96: push edx
  loc_00E30C97: push eax
  loc_00E30C98: push 00000002h
  loc_00E30C9A: call [00401038h] ; __vbaFreeVarList
  loc_00E30CA0: mov eax, [00E3D0D8h]
  loc_00E30CA5: add esp, 00000024h
  loc_00E30CA8: test eax, eax
  loc_00E30CAA: jnz 00E30CC1h
  loc_00E30CAC: push 00E3D0D8h
  loc_00E30CB1: push 006D300Ch
  loc_00E30CB6: call [004011A0h] ; __vbaNew2
  loc_00E30CBC: mov eax, [00E3D0D8h]
  loc_00E30CC1: mov ecx, [eax]
  loc_00E30CC3: push eax
  loc_00E30CC4: call [ecx+00000438h]
  loc_00E30CCA: lea edx, var_40
  loc_00E30CCD: push eax
  loc_00E30CCE: push edx
  loc_00E30CCF: call __vbaVarDup
  loc_00E30CD1: push 006DCBD8h
  loc_00E30CD6: mov var_DC, eax
  loc_00E30CDC: mov eax, [ebx]
  loc_00E30CDE: push 00000000h
  loc_00E30CE0: push 00000018h
  loc_00E30CE2: push ebx
  loc_00E30CE3: call [eax+00000324h]
  loc_00E30CE9: lea ecx, var_30
  loc_00E30CEC: push eax
  loc_00E30CED: push ecx
  loc_00E30CEE: call __vbaVarDup
  loc_00E30CF0: lea edx, var_50
  loc_00E30CF3: push eax
  loc_00E30CF4: push edx
  loc_00E30CF5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E30CFB: add esp, 00000010h
  loc_00E30CFE: push eax
  loc_00E30CFF: call [00401120h] ; __vbaCastObjVar
  loc_00E30D05: push eax
  loc_00E30D06: lea eax, var_34
  loc_00E30D09: push eax
  loc_00E30D0A: call __vbaVarDup
  loc_00E30D0C: mov edi, eax
  loc_00E30D0E: lea edx, var_38
  loc_00E30D11: push edx
  loc_00E30D12: push edi
  loc_00E30D13: mov ecx, [edi]
  loc_00E30D15: call [ecx+00000054h]
  loc_00E30D18: test eax, eax
  loc_00E30D1A: fnclex
  loc_00E30D1C: jge 00E30D2Dh
  loc_00E30D1E: push 00000054h
  loc_00E30D20: push 006DCBE8h
  loc_00E30D25: push edi
  loc_00E30D26: push eax
  loc_00E30D27: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30D2D: lea edi, var_3C
  loc_00E30D30: mov eax, var_38
  loc_00E30D33: push edi
  loc_00E30D34: mov ecx, 00000002h
  loc_00E30D39: sub esp, 00000010h
  loc_00E30D3C: mov var_88, ecx
  loc_00E30D42: mov edi, esp
  loc_00E30D44: mov var_90, ecx
  loc_00E30D4A: mov edx, [eax]
  loc_00E30D4C: push eax
  loc_00E30D4D: mov [edi], ecx
  loc_00E30D4F: mov ecx, var_8C
  loc_00E30D55: mov var_CC, eax
  loc_00E30D5B: mov [edi+00000004h], ecx
  loc_00E30D5E: mov ecx, var_88
  loc_00E30D64: mov [edi+00000008h], ecx
  loc_00E30D67: mov ecx, var_84
  loc_00E30D6D: mov [edi+0000000Ch], ecx
  loc_00E30D70: call [edx+00000028h]
  loc_00E30D73: test eax, eax
  loc_00E30D75: fnclex
  loc_00E30D77: jge 00E30D8Eh
  loc_00E30D79: mov edx, var_CC
  loc_00E30D7F: push 00000028h
  loc_00E30D81: push 006E09E8h
  loc_00E30D86: push edx
  loc_00E30D87: push eax
  loc_00E30D88: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30D8E: mov eax, var_3C
  loc_00E30D91: lea edx, var_60
  loc_00E30D94: push edx
  loc_00E30D95: push eax
  loc_00E30D96: mov ecx, [eax]
  loc_00E30D98: mov edi, eax
  loc_00E30D9A: call [ecx+00000034h]
  loc_00E30D9D: test eax, eax
  loc_00E30D9F: fnclex
  loc_00E30DA1: jge 00E30DB2h
  loc_00E30DA3: push 00000034h
  loc_00E30DA5: push 006E09F8h
  loc_00E30DAA: push edi
  loc_00E30DAB: push eax
  loc_00E30DAC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30DB2: mov eax, var_DC
  loc_00E30DB8: lea ecx, var_60
  loc_00E30DBB: push ecx
  loc_00E30DBC: mov edi, [eax]
  loc_00E30DBE: call [00401034h] ; __vbaStrVarMove
  loc_00E30DC4: mov edx, eax
  loc_00E30DC6: lea ecx, var_28
  loc_00E30DC9: call [00401228h] ; __vbaStrMove
  loc_00E30DCF: mov edx, edi
  loc_00E30DD1: mov edi, var_DC
  loc_00E30DD7: push eax
  loc_00E30DD8: push edi
  loc_00E30DD9: call [edx+00000054h]
  loc_00E30DDC: test eax, eax
  loc_00E30DDE: fnclex
  loc_00E30DE0: jge 00E30DF1h
  loc_00E30DE2: push 00000054h
  loc_00E30DE4: push 006E0410h
  loc_00E30DE9: push edi
  loc_00E30DEA: push eax
  loc_00E30DEB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30DF1: lea ecx, var_28
  loc_00E30DF4: call [00401258h] ; __vbaFreeStr
  loc_00E30DFA: lea eax, var_40
  loc_00E30DFD: lea ecx, var_3C
  loc_00E30E00: push eax
  loc_00E30E01: lea edx, var_38
  loc_00E30E04: push ecx
  loc_00E30E05: lea eax, var_34
  loc_00E30E08: push edx
  loc_00E30E09: lea ecx, var_30
  loc_00E30E0C: push eax
  loc_00E30E0D: push ecx
  loc_00E30E0E: push 00000005h
  loc_00E30E10: call [00401048h] ; __vbaFreeObjList
  loc_00E30E16: lea edx, var_60
  loc_00E30E19: lea eax, var_50
  loc_00E30E1C: push edx
  loc_00E30E1D: push eax
  loc_00E30E1E: push 00000002h
  loc_00E30E20: call [00401038h] ; __vbaFreeVarList
  loc_00E30E26: mov eax, [00E3D0D8h]
  loc_00E30E2B: add esp, 00000024h
  loc_00E30E2E: test eax, eax
  loc_00E30E30: jnz 00E30E47h
  loc_00E30E32: push 00E3D0D8h
  loc_00E30E37: push 006D300Ch
  loc_00E30E3C: call [004011A0h] ; __vbaNew2
  loc_00E30E42: mov eax, [00E3D0D8h]
  loc_00E30E47: mov ecx, [eax]
  loc_00E30E49: push eax
  loc_00E30E4A: call [ecx+0000044Ch]
  loc_00E30E50: lea edx, var_40
  loc_00E30E53: push eax
  loc_00E30E54: push edx
  loc_00E30E55: call __vbaVarDup
  loc_00E30E57: push 006DCBD8h
  loc_00E30E5C: mov var_DC, eax
  loc_00E30E62: mov eax, [ebx]
  loc_00E30E64: push 00000000h
  loc_00E30E66: push 00000018h
  loc_00E30E68: push ebx
  loc_00E30E69: call [eax+00000324h]
  loc_00E30E6F: lea ecx, var_30
  loc_00E30E72: push eax
  loc_00E30E73: push ecx
  loc_00E30E74: call __vbaVarDup
  loc_00E30E76: lea edx, var_50
  loc_00E30E79: push eax
  loc_00E30E7A: push edx
  loc_00E30E7B: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E30E81: add esp, 00000010h
  loc_00E30E84: push eax
  loc_00E30E85: call [00401120h] ; __vbaCastObjVar
  loc_00E30E8B: push eax
  loc_00E30E8C: lea eax, var_34
  loc_00E30E8F: push eax
  loc_00E30E90: call __vbaVarDup
  loc_00E30E92: mov edi, eax
  loc_00E30E94: lea edx, var_38
  loc_00E30E97: push edx
  loc_00E30E98: push edi
  loc_00E30E99: mov ecx, [edi]
  loc_00E30E9B: call [ecx+00000054h]
  loc_00E30E9E: test eax, eax
  loc_00E30EA0: fnclex
  loc_00E30EA2: jge 00E30EB3h
  loc_00E30EA4: push 00000054h
  loc_00E30EA6: push 006DCBE8h
  loc_00E30EAB: push edi
  loc_00E30EAC: push eax
  loc_00E30EAD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30EB3: lea edi, var_3C
  loc_00E30EB6: mov eax, var_38
  loc_00E30EB9: push edi
  loc_00E30EBA: mov ecx, 00000002h
  loc_00E30EBF: sub esp, 00000010h
  loc_00E30EC2: mov var_90, ecx
  loc_00E30EC8: mov edi, esp
  loc_00E30ECA: mov var_88, 00000003h
  loc_00E30ED4: mov edx, [eax]
  loc_00E30ED6: push eax
  loc_00E30ED7: mov [edi], ecx
  loc_00E30ED9: mov ecx, var_8C
  loc_00E30EDF: mov var_CC, eax
  loc_00E30EE5: mov [edi+00000004h], ecx
  loc_00E30EE8: mov ecx, var_88
  loc_00E30EEE: mov [edi+00000008h], ecx
  loc_00E30EF1: mov ecx, var_84
  loc_00E30EF7: mov [edi+0000000Ch], ecx
  loc_00E30EFA: call [edx+00000028h]
  loc_00E30EFD: test eax, eax
  loc_00E30EFF: fnclex
  loc_00E30F01: jge 00E30F18h
  loc_00E30F03: mov edx, var_CC
  loc_00E30F09: push 00000028h
  loc_00E30F0B: push 006E09E8h
  loc_00E30F10: push edx
  loc_00E30F11: push eax
  loc_00E30F12: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30F18: mov eax, var_3C
  loc_00E30F1B: lea edx, var_60
  loc_00E30F1E: push edx
  loc_00E30F1F: push eax
  loc_00E30F20: mov ecx, [eax]
  loc_00E30F22: mov edi, eax
  loc_00E30F24: call [ecx+00000034h]
  loc_00E30F27: test eax, eax
  loc_00E30F29: fnclex
  loc_00E30F2B: jge 00E30F3Ch
  loc_00E30F2D: push 00000034h
  loc_00E30F2F: push 006E09F8h
  loc_00E30F34: push edi
  loc_00E30F35: push eax
  loc_00E30F36: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30F3C: mov eax, var_DC
  loc_00E30F42: lea ecx, var_60
  loc_00E30F45: push ecx
  loc_00E30F46: mov edi, [eax]
  loc_00E30F48: call [00401034h] ; __vbaStrVarMove
  loc_00E30F4E: mov edx, eax
  loc_00E30F50: lea ecx, var_28
  loc_00E30F53: call [00401228h] ; __vbaStrMove
  loc_00E30F59: mov edx, edi
  loc_00E30F5B: mov edi, var_DC
  loc_00E30F61: push eax
  loc_00E30F62: push edi
  loc_00E30F63: call [edx+00000054h]
  loc_00E30F66: test eax, eax
  loc_00E30F68: fnclex
  loc_00E30F6A: jge 00E30F7Bh
  loc_00E30F6C: push 00000054h
  loc_00E30F6E: push 006E0410h
  loc_00E30F73: push edi
  loc_00E30F74: push eax
  loc_00E30F75: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30F7B: lea ecx, var_28
  loc_00E30F7E: call [00401258h] ; __vbaFreeStr
  loc_00E30F84: lea eax, var_40
  loc_00E30F87: lea ecx, var_3C
  loc_00E30F8A: push eax
  loc_00E30F8B: lea edx, var_38
  loc_00E30F8E: push ecx
  loc_00E30F8F: lea eax, var_34
  loc_00E30F92: push edx
  loc_00E30F93: lea ecx, var_30
  loc_00E30F96: push eax
  loc_00E30F97: push ecx
  loc_00E30F98: push 00000005h
  loc_00E30F9A: call [00401048h] ; __vbaFreeObjList
  loc_00E30FA0: lea edx, var_60
  loc_00E30FA3: lea eax, var_50
  loc_00E30FA6: push edx
  loc_00E30FA7: push eax
  loc_00E30FA8: push 00000002h
  loc_00E30FAA: call [00401038h] ; __vbaFreeVarList
  loc_00E30FB0: mov eax, [00E3D0D8h]
  loc_00E30FB5: add esp, 00000024h
  loc_00E30FB8: test eax, eax
  loc_00E30FBA: jnz 00E30FD1h
  loc_00E30FBC: push 00E3D0D8h
  loc_00E30FC1: push 006D300Ch
  loc_00E30FC6: call [004011A0h] ; __vbaNew2
  loc_00E30FCC: mov eax, [00E3D0D8h]
  loc_00E30FD1: mov ecx, [eax]
  loc_00E30FD3: push eax
  loc_00E30FD4: call [ecx+00000450h]
  loc_00E30FDA: lea edx, var_40
  loc_00E30FDD: push eax
  loc_00E30FDE: push edx
  loc_00E30FDF: call __vbaVarDup
  loc_00E30FE1: push 006DCBD8h
  loc_00E30FE6: mov var_DC, eax
  loc_00E30FEC: mov eax, [ebx]
  loc_00E30FEE: push 00000000h
  loc_00E30FF0: push 00000018h
  loc_00E30FF2: push ebx
  loc_00E30FF3: call [eax+00000324h]
  loc_00E30FF9: lea ecx, var_30
  loc_00E30FFC: push eax
  loc_00E30FFD: push ecx
  loc_00E30FFE: call __vbaVarDup
  loc_00E31000: lea edx, var_50
  loc_00E31003: push eax
  loc_00E31004: push edx
  loc_00E31005: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E3100B: add esp, 00000010h
  loc_00E3100E: push eax
  loc_00E3100F: call [00401120h] ; __vbaCastObjVar
  loc_00E31015: push eax
  loc_00E31016: lea eax, var_34
  loc_00E31019: push eax
  loc_00E3101A: call __vbaVarDup
  loc_00E3101C: mov edi, eax
  loc_00E3101E: lea edx, var_38
  loc_00E31021: push edx
  loc_00E31022: push edi
  loc_00E31023: mov ecx, [edi]
  loc_00E31025: call [ecx+00000054h]
  loc_00E31028: test eax, eax
  loc_00E3102A: fnclex
  loc_00E3102C: jge 00E3103Dh
  loc_00E3102E: push 00000054h
  loc_00E31030: push 006DCBE8h
  loc_00E31035: push edi
  loc_00E31036: push eax
  loc_00E31037: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3103D: lea edi, var_3C
  loc_00E31040: mov eax, var_38
  loc_00E31043: push edi
  loc_00E31044: mov ecx, 00000002h
  loc_00E31049: sub esp, 00000010h
  loc_00E3104C: mov var_90, ecx
  loc_00E31052: mov edi, esp
  loc_00E31054: mov var_88, 00000004h
  loc_00E3105E: mov edx, [eax]
  loc_00E31060: push eax
  loc_00E31061: mov [edi], ecx
  loc_00E31063: mov ecx, var_8C
  loc_00E31069: mov var_CC, eax
  loc_00E3106F: mov [edi+00000004h], ecx
  loc_00E31072: mov ecx, var_88
  loc_00E31078: mov [edi+00000008h], ecx
  loc_00E3107B: mov ecx, var_84
  loc_00E31081: mov [edi+0000000Ch], ecx
  loc_00E31084: call [edx+00000028h]
  loc_00E31087: test eax, eax
  loc_00E31089: fnclex
  loc_00E3108B: jge 00E310A2h
  loc_00E3108D: mov edx, var_CC
  loc_00E31093: push 00000028h
  loc_00E31095: push 006E09E8h
  loc_00E3109A: push edx
  loc_00E3109B: push eax
  loc_00E3109C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E310A2: mov eax, var_3C
  loc_00E310A5: lea edx, var_60
  loc_00E310A8: push edx
  loc_00E310A9: push eax
  loc_00E310AA: mov ecx, [eax]
  loc_00E310AC: mov edi, eax
  loc_00E310AE: call [ecx+00000034h]
  loc_00E310B1: test eax, eax
  loc_00E310B3: fnclex
  loc_00E310B5: jge 00E310C6h
  loc_00E310B7: push 00000034h
  loc_00E310B9: push 006E09F8h
  loc_00E310BE: push edi
  loc_00E310BF: push eax
  loc_00E310C0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E310C6: mov eax, var_DC
  loc_00E310CC: lea ecx, var_60
  loc_00E310CF: push ecx
  loc_00E310D0: mov edi, [eax]
  loc_00E310D2: call [00401034h] ; __vbaStrVarMove
  loc_00E310D8: mov edx, eax
  loc_00E310DA: lea ecx, var_28
  loc_00E310DD: call [00401228h] ; __vbaStrMove
  loc_00E310E3: mov edx, edi
  loc_00E310E5: mov edi, var_DC
  loc_00E310EB: push eax
  loc_00E310EC: push edi
  loc_00E310ED: call [edx+00000054h]
  loc_00E310F0: test eax, eax
  loc_00E310F2: fnclex
  loc_00E310F4: jge 00E31105h
  loc_00E310F6: push 00000054h
  loc_00E310F8: push 006E0410h
  loc_00E310FD: push edi
  loc_00E310FE: push eax
  loc_00E310FF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31105: lea ecx, var_28
  loc_00E31108: call [00401258h] ; __vbaFreeStr
  loc_00E3110E: lea eax, var_40
  loc_00E31111: lea ecx, var_3C
  loc_00E31114: push eax
  loc_00E31115: lea edx, var_38
  loc_00E31118: push ecx
  loc_00E31119: lea eax, var_34
  loc_00E3111C: push edx
  loc_00E3111D: lea ecx, var_30
  loc_00E31120: push eax
  loc_00E31121: push ecx
  loc_00E31122: push 00000005h
  loc_00E31124: call [00401048h] ; __vbaFreeObjList
  loc_00E3112A: lea edx, var_60
  loc_00E3112D: lea eax, var_50
  loc_00E31130: push edx
  loc_00E31131: push eax
  loc_00E31132: push 00000002h
  loc_00E31134: call [00401038h] ; __vbaFreeVarList
  loc_00E3113A: mov eax, [00E3D0D8h]
  loc_00E3113F: add esp, 00000024h
  loc_00E31142: test eax, eax
  loc_00E31144: jnz 00E3115Bh
  loc_00E31146: push 00E3D0D8h
  loc_00E3114B: push 006D300Ch
  loc_00E31150: call [004011A0h] ; __vbaNew2
  loc_00E31156: mov eax, [00E3D0D8h]
  loc_00E3115B: mov ecx, [eax]
  loc_00E3115D: push eax
  loc_00E3115E: call [ecx+00000424h]
  loc_00E31164: lea edx, var_40
  loc_00E31167: push eax
  loc_00E31168: push edx
  loc_00E31169: call __vbaVarDup
  loc_00E3116B: push 006DCBD8h
  loc_00E31170: mov var_DC, eax
  loc_00E31176: mov eax, [ebx]
  loc_00E31178: push 00000000h
  loc_00E3117A: push 00000018h
  loc_00E3117C: push ebx
  loc_00E3117D: call [eax+00000324h]
  loc_00E31183: lea ecx, var_30
  loc_00E31186: push eax
  loc_00E31187: push ecx
  loc_00E31188: call __vbaVarDup
  loc_00E3118A: lea edx, var_50
  loc_00E3118D: push eax
  loc_00E3118E: push edx
  loc_00E3118F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E31195: add esp, 00000010h
  loc_00E31198: push eax
  loc_00E31199: call [00401120h] ; __vbaCastObjVar
  loc_00E3119F: push eax
  loc_00E311A0: lea eax, var_34
  loc_00E311A3: push eax
  loc_00E311A4: call __vbaVarDup
  loc_00E311A6: mov edi, eax
  loc_00E311A8: lea edx, var_38
  loc_00E311AB: push edx
  loc_00E311AC: push edi
  loc_00E311AD: mov ecx, [edi]
  loc_00E311AF: call [ecx+00000054h]
  loc_00E311B2: test eax, eax
  loc_00E311B4: fnclex
  loc_00E311B6: jge 00E311C7h
  loc_00E311B8: push 00000054h
  loc_00E311BA: push 006DCBE8h
  loc_00E311BF: push edi
  loc_00E311C0: push eax
  loc_00E311C1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E311C7: lea edi, var_3C
  loc_00E311CA: mov eax, var_38
  loc_00E311CD: push edi
  loc_00E311CE: mov ecx, 00000002h
  loc_00E311D3: sub esp, 00000010h
  loc_00E311D6: mov var_90, ecx
  loc_00E311DC: mov edi, esp
  loc_00E311DE: mov var_88, 00000005h
  loc_00E311E8: mov edx, [eax]
  loc_00E311EA: push eax
  loc_00E311EB: mov [edi], ecx
  loc_00E311ED: mov ecx, var_8C
  loc_00E311F3: mov var_CC, eax
  loc_00E311F9: mov [edi+00000004h], ecx
  loc_00E311FC: mov ecx, var_88
  loc_00E31202: mov [edi+00000008h], ecx
  loc_00E31205: mov ecx, var_84
  loc_00E3120B: mov [edi+0000000Ch], ecx
  loc_00E3120E: call [edx+00000028h]
  loc_00E31211: test eax, eax
  loc_00E31213: fnclex
  loc_00E31215: jge 00E3122Ch
  loc_00E31217: mov edx, var_CC
  loc_00E3121D: push 00000028h
  loc_00E3121F: push 006E09E8h
  loc_00E31224: push edx
  loc_00E31225: push eax
  loc_00E31226: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3122C: mov eax, var_3C
  loc_00E3122F: lea edx, var_60
  loc_00E31232: push edx
  loc_00E31233: push eax
  loc_00E31234: mov ecx, [eax]
  loc_00E31236: mov edi, eax
  loc_00E31238: call [ecx+00000034h]
  loc_00E3123B: test eax, eax
  loc_00E3123D: fnclex
  loc_00E3123F: jge 00E31250h
  loc_00E31241: push 00000034h
  loc_00E31243: push 006E09F8h
  loc_00E31248: push edi
  loc_00E31249: push eax
  loc_00E3124A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31250: mov eax, var_DC
  loc_00E31256: lea ecx, var_60
  loc_00E31259: push ecx
  loc_00E3125A: mov edi, [eax]
  loc_00E3125C: call [00401034h] ; __vbaStrVarMove
  loc_00E31262: mov edx, eax
  loc_00E31264: lea ecx, var_28
  loc_00E31267: call [00401228h] ; __vbaStrMove
  loc_00E3126D: mov edx, edi
  loc_00E3126F: mov edi, var_DC
  loc_00E31275: push eax
  loc_00E31276: push edi
  loc_00E31277: call [edx+00000054h]
  loc_00E3127A: test eax, eax
  loc_00E3127C: fnclex
  loc_00E3127E: jge 00E3128Fh
  loc_00E31280: push 00000054h
  loc_00E31282: push 006E0410h
  loc_00E31287: push edi
  loc_00E31288: push eax
  loc_00E31289: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3128F: lea ecx, var_28
  loc_00E31292: call [00401258h] ; __vbaFreeStr
  loc_00E31298: lea eax, var_40
  loc_00E3129B: lea ecx, var_3C
  loc_00E3129E: push eax
  loc_00E3129F: lea edx, var_38
  loc_00E312A2: push ecx
  loc_00E312A3: lea eax, var_34
  loc_00E312A6: push edx
  loc_00E312A7: lea ecx, var_30
  loc_00E312AA: push eax
  loc_00E312AB: push ecx
  loc_00E312AC: push 00000005h
  loc_00E312AE: call [00401048h] ; __vbaFreeObjList
  loc_00E312B4: lea edx, var_60
  loc_00E312B7: lea eax, var_50
  loc_00E312BA: push edx
  loc_00E312BB: push eax
  loc_00E312BC: push 00000002h
  loc_00E312BE: call [00401038h] ; __vbaFreeVarList
  loc_00E312C4: mov eax, [00E3D0D8h]
  loc_00E312C9: add esp, 00000024h
  loc_00E312CC: test eax, eax
  loc_00E312CE: jnz 00E312E5h
  loc_00E312D0: push 00E3D0D8h
  loc_00E312D5: push 006D300Ch
  loc_00E312DA: call [004011A0h] ; __vbaNew2
  loc_00E312E0: mov eax, [00E3D0D8h]
  loc_00E312E5: mov ecx, [eax]
  loc_00E312E7: push eax
  loc_00E312E8: call [ecx+00000378h]
  loc_00E312EE: lea edx, var_40
  loc_00E312F1: push eax
  loc_00E312F2: push edx
  loc_00E312F3: call __vbaVarDup
  loc_00E312F5: push 006DCBD8h
  loc_00E312FA: mov var_DC, eax
  loc_00E31300: mov eax, [ebx]
  loc_00E31302: push 00000000h
  loc_00E31304: push 00000018h
  loc_00E31306: push ebx
  loc_00E31307: call [eax+00000324h]
  loc_00E3130D: lea ecx, var_30
  loc_00E31310: push eax
  loc_00E31311: push ecx
  loc_00E31312: call __vbaVarDup
  loc_00E31314: lea edx, var_50
  loc_00E31317: push eax
  loc_00E31318: push edx
  loc_00E31319: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E3131F: add esp, 00000010h
  loc_00E31322: push eax
  loc_00E31323: call [00401120h] ; __vbaCastObjVar
  loc_00E31329: push eax
  loc_00E3132A: lea eax, var_34
  loc_00E3132D: push eax
  loc_00E3132E: call __vbaVarDup
  loc_00E31330: mov edi, eax
  loc_00E31332: lea edx, var_38
  loc_00E31335: push edx
  loc_00E31336: push edi
  loc_00E31337: mov ecx, [edi]
  loc_00E31339: call [ecx+00000054h]
  loc_00E3133C: test eax, eax
  loc_00E3133E: fnclex
  loc_00E31340: jge 00E31351h
  loc_00E31342: push 00000054h
  loc_00E31344: push 006DCBE8h
  loc_00E31349: push edi
  loc_00E3134A: push eax
  loc_00E3134B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31351: lea edi, var_3C
  loc_00E31354: mov eax, var_38
  loc_00E31357: push edi
  loc_00E31358: mov ecx, 00000002h
  loc_00E3135D: sub esp, 00000010h
  loc_00E31360: mov var_90, ecx
  loc_00E31366: mov edi, esp
  loc_00E31368: mov var_88, 00000006h
  loc_00E31372: mov edx, [eax]
  loc_00E31374: push eax
  loc_00E31375: mov [edi], ecx
  loc_00E31377: mov ecx, var_8C
  loc_00E3137D: mov var_CC, eax
  loc_00E31383: mov [edi+00000004h], ecx
  loc_00E31386: mov ecx, var_88
  loc_00E3138C: mov [edi+00000008h], ecx
  loc_00E3138F: mov ecx, var_84
  loc_00E31395: mov [edi+0000000Ch], ecx
  loc_00E31398: call [edx+00000028h]
  loc_00E3139B: test eax, eax
  loc_00E3139D: fnclex
  loc_00E3139F: jge 00E313B6h
  loc_00E313A1: mov edx, var_CC
  loc_00E313A7: push 00000028h
  loc_00E313A9: push 006E09E8h
  loc_00E313AE: push edx
  loc_00E313AF: push eax
  loc_00E313B0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E313B6: mov eax, var_3C
  loc_00E313B9: lea edx, var_60
  loc_00E313BC: push edx
  loc_00E313BD: push eax
  loc_00E313BE: mov ecx, [eax]
  loc_00E313C0: mov edi, eax
  loc_00E313C2: call [ecx+00000034h]
  loc_00E313C5: test eax, eax
  loc_00E313C7: fnclex
  loc_00E313C9: jge 00E313DAh
  loc_00E313CB: push 00000034h
  loc_00E313CD: push 006E09F8h
  loc_00E313D2: push edi
  loc_00E313D3: push eax
  loc_00E313D4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E313DA: mov eax, var_DC
  loc_00E313E0: lea ecx, var_60
  loc_00E313E3: push ecx
  loc_00E313E4: mov edi, [eax]
  loc_00E313E6: call [00401034h] ; __vbaStrVarMove
  loc_00E313EC: mov edx, eax
  loc_00E313EE: lea ecx, var_28
  loc_00E313F1: call [00401228h] ; __vbaStrMove
  loc_00E313F7: mov edx, edi
  loc_00E313F9: mov edi, var_DC
  loc_00E313FF: push eax
  loc_00E31400: push edi
  loc_00E31401: call [edx+000000A4h]
  loc_00E31407: test eax, eax
  loc_00E31409: fnclex
  loc_00E3140B: jge 00E3141Fh
  loc_00E3140D: push 000000A4h
  loc_00E31412: push 006DCB70h
  loc_00E31417: push edi
  loc_00E31418: push eax
  loc_00E31419: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3141F: lea ecx, var_28
  loc_00E31422: call [00401258h] ; __vbaFreeStr
  loc_00E31428: lea eax, var_40
  loc_00E3142B: lea ecx, var_3C
  loc_00E3142E: push eax
  loc_00E3142F: lea edx, var_38
  loc_00E31432: push ecx
  loc_00E31433: lea eax, var_34
  loc_00E31436: push edx
  loc_00E31437: lea ecx, var_30
  loc_00E3143A: push eax
  loc_00E3143B: push ecx
  loc_00E3143C: push 00000005h
  loc_00E3143E: call [00401048h] ; __vbaFreeObjList
  loc_00E31444: lea edx, var_60
  loc_00E31447: lea eax, var_50
  loc_00E3144A: push edx
  loc_00E3144B: push eax
  loc_00E3144C: push 00000002h
  loc_00E3144E: call [00401038h] ; __vbaFreeVarList
  loc_00E31454: mov eax, [00E3D0D8h]
  loc_00E31459: add esp, 00000024h
  loc_00E3145C: test eax, eax
  loc_00E3145E: jnz 00E31475h
  loc_00E31460: push 00E3D0D8h
  loc_00E31465: push 006D300Ch
  loc_00E3146A: call [004011A0h] ; __vbaNew2
  loc_00E31470: mov eax, [00E3D0D8h]
  loc_00E31475: mov ecx, [eax]
  loc_00E31477: push eax
  loc_00E31478: call [ecx+0000039Ch]
  loc_00E3147E: lea edx, var_40
  loc_00E31481: push eax
  loc_00E31482: push edx
  loc_00E31483: call __vbaVarDup
  loc_00E31485: push 006DCBD8h
  loc_00E3148A: mov var_DC, eax
  loc_00E31490: mov eax, [ebx]
  loc_00E31492: push 00000000h
  loc_00E31494: push 00000018h
  loc_00E31496: push ebx
  loc_00E31497: call [eax+00000324h]
  loc_00E3149D: lea ecx, var_30
  loc_00E314A0: push eax
  loc_00E314A1: push ecx
  loc_00E314A2: call __vbaVarDup
  loc_00E314A4: lea edx, var_50
  loc_00E314A7: push eax
  loc_00E314A8: push edx
  loc_00E314A9: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E314AF: add esp, 00000010h
  loc_00E314B2: push eax
  loc_00E314B3: call [00401120h] ; __vbaCastObjVar
  loc_00E314B9: push eax
  loc_00E314BA: lea eax, var_34
  loc_00E314BD: push eax
  loc_00E314BE: call __vbaVarDup
  loc_00E314C0: mov edi, eax
  loc_00E314C2: lea edx, var_38
  loc_00E314C5: push edx
  loc_00E314C6: push edi
  loc_00E314C7: mov ecx, [edi]
  loc_00E314C9: call [ecx+00000054h]
  loc_00E314CC: test eax, eax
  loc_00E314CE: fnclex
  loc_00E314D0: jge 00E314E1h
  loc_00E314D2: push 00000054h
  loc_00E314D4: push 006DCBE8h
  loc_00E314D9: push edi
  loc_00E314DA: push eax
  loc_00E314DB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E314E1: lea edi, var_3C
  loc_00E314E4: mov eax, var_38
  loc_00E314E7: push edi
  loc_00E314E8: mov ecx, 00000002h
  loc_00E314ED: sub esp, 00000010h
  loc_00E314F0: mov var_90, ecx
  loc_00E314F6: mov edi, esp
  loc_00E314F8: mov var_88, 00000007h
  loc_00E31502: mov edx, [eax]
  loc_00E31504: push eax
  loc_00E31505: mov [edi], ecx
  loc_00E31507: mov ecx, var_8C
  loc_00E3150D: mov var_CC, eax
  loc_00E31513: mov [edi+00000004h], ecx
  loc_00E31516: mov ecx, var_88
  loc_00E3151C: mov [edi+00000008h], ecx
  loc_00E3151F: mov ecx, var_84
  loc_00E31525: mov [edi+0000000Ch], ecx
  loc_00E31528: call [edx+00000028h]
  loc_00E3152B: test eax, eax
  loc_00E3152D: fnclex
  loc_00E3152F: jge 00E31546h
  loc_00E31531: mov edx, var_CC
  loc_00E31537: push 00000028h
  loc_00E31539: push 006E09E8h
  loc_00E3153E: push edx
  loc_00E3153F: push eax
  loc_00E31540: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31546: mov eax, var_3C
  loc_00E31549: lea edx, var_60
  loc_00E3154C: push edx
  loc_00E3154D: push eax
  loc_00E3154E: mov ecx, [eax]
  loc_00E31550: mov edi, eax
  loc_00E31552: call [ecx+00000034h]
  loc_00E31555: test eax, eax
  loc_00E31557: fnclex
  loc_00E31559: jge 00E3156Ah
  loc_00E3155B: push 00000034h
  loc_00E3155D: push 006E09F8h
  loc_00E31562: push edi
  loc_00E31563: push eax
  loc_00E31564: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3156A: mov eax, var_DC
  loc_00E31570: lea ecx, var_60
  loc_00E31573: push ecx
  loc_00E31574: mov edi, [eax]
  loc_00E31576: call [00401034h] ; __vbaStrVarMove
  loc_00E3157C: mov edx, eax
  loc_00E3157E: lea ecx, var_28
  loc_00E31581: call [00401228h] ; __vbaStrMove
  loc_00E31587: mov edx, edi
  loc_00E31589: mov edi, var_DC
  loc_00E3158F: push eax
  loc_00E31590: push edi
  loc_00E31591: call [edx+000000A4h]
  loc_00E31597: test eax, eax
  loc_00E31599: fnclex
  loc_00E3159B: jge 00E315AFh
  loc_00E3159D: push 000000A4h
  loc_00E315A2: push 006DCB70h
  loc_00E315A7: push edi
  loc_00E315A8: push eax
  loc_00E315A9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E315AF: lea ecx, var_28
  loc_00E315B2: call [00401258h] ; __vbaFreeStr
  loc_00E315B8: lea eax, var_40
  loc_00E315BB: lea ecx, var_3C
  loc_00E315BE: push eax
  loc_00E315BF: lea edx, var_38
  loc_00E315C2: push ecx
  loc_00E315C3: lea eax, var_34
  loc_00E315C6: push edx
  loc_00E315C7: lea ecx, var_30
  loc_00E315CA: push eax
  loc_00E315CB: push ecx
  loc_00E315CC: push 00000005h
  loc_00E315CE: call [00401048h] ; __vbaFreeObjList
  loc_00E315D4: lea edx, var_60
  loc_00E315D7: lea eax, var_50
  loc_00E315DA: push edx
  loc_00E315DB: push eax
  loc_00E315DC: push 00000002h
  loc_00E315DE: call [00401038h] ; __vbaFreeVarList
  loc_00E315E4: mov eax, [00E3D0D8h]
  loc_00E315E9: add esp, 00000024h
  loc_00E315EC: test eax, eax
  loc_00E315EE: jnz 00E31605h
  loc_00E315F0: push 00E3D0D8h
  loc_00E315F5: push 006D300Ch
  loc_00E315FA: call [004011A0h] ; __vbaNew2
  loc_00E31600: mov eax, [00E3D0D8h]
  loc_00E31605: mov ecx, [eax]
  loc_00E31607: push eax
  loc_00E31608: call [ecx+00000368h]
  loc_00E3160E: lea edx, var_40
  loc_00E31611: push eax
  loc_00E31612: push edx
  loc_00E31613: call __vbaVarDup
  loc_00E31615: push 006DCBD8h
  loc_00E3161A: mov var_DC, eax
  loc_00E31620: mov eax, [ebx]
  loc_00E31622: push 00000000h
  loc_00E31624: push 00000018h
  loc_00E31626: push ebx
  loc_00E31627: call [eax+00000324h]
  loc_00E3162D: lea ecx, var_30
  loc_00E31630: push eax
  loc_00E31631: push ecx
  loc_00E31632: call __vbaVarDup
  loc_00E31634: lea edx, var_50
  loc_00E31637: push eax
  loc_00E31638: push edx
  loc_00E31639: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E3163F: add esp, 00000010h
  loc_00E31642: push eax
  loc_00E31643: call [00401120h] ; __vbaCastObjVar
  loc_00E31649: push eax
  loc_00E3164A: lea eax, var_34
  loc_00E3164D: push eax
  loc_00E3164E: call __vbaVarDup
  loc_00E31650: mov edi, eax
  loc_00E31652: lea edx, var_38
  loc_00E31655: push edx
  loc_00E31656: push edi
  loc_00E31657: mov ecx, [edi]
  loc_00E31659: call [ecx+00000054h]
  loc_00E3165C: test eax, eax
  loc_00E3165E: fnclex
  loc_00E31660: jge 00E31671h
  loc_00E31662: push 00000054h
  loc_00E31664: push 006DCBE8h
  loc_00E31669: push edi
  loc_00E3166A: push eax
  loc_00E3166B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31671: lea edi, var_3C
  loc_00E31674: mov eax, var_38
  loc_00E31677: push edi
  loc_00E31678: mov ecx, 00000002h
  loc_00E3167D: sub esp, 00000010h
  loc_00E31680: mov var_90, ecx
  loc_00E31686: mov edi, esp
  loc_00E31688: mov var_88, 00000008h
  loc_00E31692: mov edx, [eax]
  loc_00E31694: push eax
  loc_00E31695: mov [edi], ecx
  loc_00E31697: mov ecx, var_8C
  loc_00E3169D: mov var_CC, eax
  loc_00E316A3: mov [edi+00000004h], ecx
  loc_00E316A6: mov ecx, var_88
  loc_00E316AC: mov [edi+00000008h], ecx
  loc_00E316AF: mov ecx, var_84
  loc_00E316B5: mov [edi+0000000Ch], ecx
  loc_00E316B8: call [edx+00000028h]
  loc_00E316BB: test eax, eax
  loc_00E316BD: fnclex
  loc_00E316BF: jge 00E316D6h
  loc_00E316C1: mov edx, var_CC
  loc_00E316C7: push 00000028h
  loc_00E316C9: push 006E09E8h
  loc_00E316CE: push edx
  loc_00E316CF: push eax
  loc_00E316D0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E316D6: mov eax, var_3C
  loc_00E316D9: lea edx, var_60
  loc_00E316DC: push edx
  loc_00E316DD: push eax
  loc_00E316DE: mov ecx, [eax]
  loc_00E316E0: mov edi, eax
  loc_00E316E2: call [ecx+00000034h]
  loc_00E316E5: test eax, eax
  loc_00E316E7: fnclex
  loc_00E316E9: jge 00E316FAh
  loc_00E316EB: push 00000034h
  loc_00E316ED: push 006E09F8h
  loc_00E316F2: push edi
  loc_00E316F3: push eax
  loc_00E316F4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E316FA: mov eax, var_DC
  loc_00E31700: lea ecx, var_60
  loc_00E31703: push ecx
  loc_00E31704: mov edi, [eax]
  loc_00E31706: call [00401034h] ; __vbaStrVarMove
  loc_00E3170C: mov edx, eax
  loc_00E3170E: lea ecx, var_28
  loc_00E31711: call [00401228h] ; __vbaStrMove
  loc_00E31717: mov edx, edi
  loc_00E31719: mov edi, var_DC
  loc_00E3171F: push eax
  loc_00E31720: push edi
  loc_00E31721: call [edx+00000054h]
  loc_00E31724: test eax, eax
  loc_00E31726: fnclex
  loc_00E31728: jge 00E31739h
  loc_00E3172A: push 00000054h
  loc_00E3172C: push 006E0410h
  loc_00E31731: push edi
  loc_00E31732: push eax
  loc_00E31733: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31739: lea ecx, var_28
  loc_00E3173C: call [00401258h] ; __vbaFreeStr
  loc_00E31742: lea eax, var_40
  loc_00E31745: lea ecx, var_3C
  loc_00E31748: push eax
  loc_00E31749: lea edx, var_38
  loc_00E3174C: push ecx
  loc_00E3174D: lea eax, var_34
  loc_00E31750: push edx
  loc_00E31751: lea ecx, var_30
  loc_00E31754: push eax
  loc_00E31755: push ecx
  loc_00E31756: push 00000005h
  loc_00E31758: call [00401048h] ; __vbaFreeObjList
  loc_00E3175E: lea edx, var_60
  loc_00E31761: lea eax, var_50
  loc_00E31764: push edx
  loc_00E31765: push eax
  loc_00E31766: push 00000002h
  loc_00E31768: call [00401038h] ; __vbaFreeVarList
  loc_00E3176E: mov eax, [00E3D0D8h]
  loc_00E31773: add esp, 00000024h
  loc_00E31776: test eax, eax
  loc_00E31778: jnz 00E3178Fh
  loc_00E3177A: push 00E3D0D8h
  loc_00E3177F: push 006D300Ch
  loc_00E31784: call [004011A0h] ; __vbaNew2
  loc_00E3178A: mov eax, [00E3D0D8h]
  loc_00E3178F: mov ecx, [eax]
  loc_00E31791: push eax
  loc_00E31792: call [ecx+00000388h]
  loc_00E31798: lea edx, var_40
  loc_00E3179B: push eax
  loc_00E3179C: push edx
  loc_00E3179D: call __vbaVarDup
  loc_00E3179F: push 006DCBD8h
  loc_00E317A4: mov var_DC, eax
  loc_00E317AA: mov eax, [ebx]
  loc_00E317AC: push 00000000h
  loc_00E317AE: push 00000018h
  loc_00E317B0: push ebx
  loc_00E317B1: call [eax+00000324h]
  loc_00E317B7: lea ecx, var_30
  loc_00E317BA: push eax
  loc_00E317BB: push ecx
  loc_00E317BC: call __vbaVarDup
  loc_00E317BE: lea edx, var_50
  loc_00E317C1: push eax
  loc_00E317C2: push edx
  loc_00E317C3: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E317C9: add esp, 00000010h
  loc_00E317CC: push eax
  loc_00E317CD: call [00401120h] ; __vbaCastObjVar
  loc_00E317D3: push eax
  loc_00E317D4: lea eax, var_34
  loc_00E317D7: push eax
  loc_00E317D8: call __vbaVarDup
  loc_00E317DA: mov edi, eax
  loc_00E317DC: lea edx, var_38
  loc_00E317DF: push edx
  loc_00E317E0: push edi
  loc_00E317E1: mov ecx, [edi]
  loc_00E317E3: call [ecx+00000054h]
  loc_00E317E6: test eax, eax
  loc_00E317E8: fnclex
  loc_00E317EA: jge 00E317FBh
  loc_00E317EC: push 00000054h
  loc_00E317EE: push 006DCBE8h
  loc_00E317F3: push edi
  loc_00E317F4: push eax
  loc_00E317F5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E317FB: lea edi, var_3C
  loc_00E317FE: mov eax, var_38
  loc_00E31801: push edi
  loc_00E31802: mov ecx, 00000002h
  loc_00E31807: sub esp, 00000010h
  loc_00E3180A: mov var_90, ecx
  loc_00E31810: mov edi, esp
  loc_00E31812: mov var_88, 00000009h
  loc_00E3181C: mov edx, [eax]
  loc_00E3181E: push eax
  loc_00E3181F: mov [edi], ecx
  loc_00E31821: mov ecx, var_8C
  loc_00E31827: mov var_CC, eax
  loc_00E3182D: mov [edi+00000004h], ecx
  loc_00E31830: mov ecx, var_88
  loc_00E31836: mov [edi+00000008h], ecx
  loc_00E31839: mov ecx, var_84
  loc_00E3183F: mov [edi+0000000Ch], ecx
  loc_00E31842: call [edx+00000028h]
  loc_00E31845: test eax, eax
  loc_00E31847: fnclex
  loc_00E31849: jge 00E31860h
  loc_00E3184B: mov edx, var_CC
  loc_00E31851: push 00000028h
  loc_00E31853: push 006E09E8h
  loc_00E31858: push edx
  loc_00E31859: push eax
  loc_00E3185A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31860: mov eax, var_3C
  loc_00E31863: lea edx, var_60
  loc_00E31866: push edx
  loc_00E31867: push eax
  loc_00E31868: mov ecx, [eax]
  loc_00E3186A: mov edi, eax
  loc_00E3186C: call [ecx+00000034h]
  loc_00E3186F: test eax, eax
  loc_00E31871: fnclex
  loc_00E31873: jge 00E31884h
  loc_00E31875: push 00000034h
  loc_00E31877: push 006E09F8h
  loc_00E3187C: push edi
  loc_00E3187D: push eax
  loc_00E3187E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31884: mov eax, var_DC
  loc_00E3188A: lea ecx, var_60
  loc_00E3188D: push ecx
  loc_00E3188E: mov edi, [eax]
  loc_00E31890: call [00401034h] ; __vbaStrVarMove
  loc_00E31896: mov edx, eax
  loc_00E31898: lea ecx, var_28
  loc_00E3189B: call [00401228h] ; __vbaStrMove
  loc_00E318A1: mov edx, edi
  loc_00E318A3: mov edi, var_DC
  loc_00E318A9: push eax
  loc_00E318AA: push edi
  loc_00E318AB: call [edx+00000054h]
  loc_00E318AE: test eax, eax
  loc_00E318B0: fnclex
  loc_00E318B2: jge 00E318C3h
  loc_00E318B4: push 00000054h
  loc_00E318B6: push 006E0410h
  loc_00E318BB: push edi
  loc_00E318BC: push eax
  loc_00E318BD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E318C3: lea ecx, var_28
  loc_00E318C6: call [00401258h] ; __vbaFreeStr
  loc_00E318CC: lea eax, var_40
  loc_00E318CF: lea ecx, var_3C
  loc_00E318D2: push eax
  loc_00E318D3: lea edx, var_38
  loc_00E318D6: push ecx
  loc_00E318D7: lea eax, var_34
  loc_00E318DA: push edx
  loc_00E318DB: lea ecx, var_30
  loc_00E318DE: push eax
  loc_00E318DF: push ecx
  loc_00E318E0: push 00000005h
  loc_00E318E2: call [00401048h] ; __vbaFreeObjList
  loc_00E318E8: lea edx, var_60
  loc_00E318EB: lea eax, var_50
  loc_00E318EE: push edx
  loc_00E318EF: push eax
  loc_00E318F0: push 00000002h
  loc_00E318F2: call [00401038h] ; __vbaFreeVarList
  loc_00E318F8: mov eax, [00E3D0D8h]
  loc_00E318FD: add esp, 00000024h
  loc_00E31900: test eax, eax
  loc_00E31902: jnz 00E31919h
  loc_00E31904: push 00E3D0D8h
  loc_00E31909: push 006D300Ch
  loc_00E3190E: call [004011A0h] ; __vbaNew2
  loc_00E31914: mov eax, [00E3D0D8h]
  loc_00E31919: mov ecx, [eax]
  loc_00E3191B: push eax
  loc_00E3191C: call [ecx+00000304h]
  loc_00E31922: lea edx, var_40
  loc_00E31925: push eax
  loc_00E31926: push edx
  loc_00E31927: call __vbaVarDup
  loc_00E31929: push 006DCBD8h
  loc_00E3192E: mov var_DC, eax
  loc_00E31934: mov eax, [ebx]
  loc_00E31936: push 00000000h
  loc_00E31938: push 00000018h
  loc_00E3193A: push ebx
  loc_00E3193B: call [eax+00000324h]
  loc_00E31941: lea ecx, var_30
  loc_00E31944: push eax
  loc_00E31945: push ecx
  loc_00E31946: call __vbaVarDup
  loc_00E31948: lea edx, var_50
  loc_00E3194B: push eax
  loc_00E3194C: push edx
  loc_00E3194D: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E31953: add esp, 00000010h
  loc_00E31956: push eax
  loc_00E31957: call [00401120h] ; __vbaCastObjVar
  loc_00E3195D: push eax
  loc_00E3195E: lea eax, var_34
  loc_00E31961: push eax
  loc_00E31962: call __vbaVarDup
  loc_00E31964: mov edi, eax
  loc_00E31966: lea edx, var_38
  loc_00E31969: push edx
  loc_00E3196A: push edi
  loc_00E3196B: mov ecx, [edi]
  loc_00E3196D: call [ecx+00000054h]
  loc_00E31970: test eax, eax
  loc_00E31972: fnclex
  loc_00E31974: jge 00E31985h
  loc_00E31976: push 00000054h
  loc_00E31978: push 006DCBE8h
  loc_00E3197D: push edi
  loc_00E3197E: push eax
  loc_00E3197F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31985: lea ebx, var_3C
  loc_00E31988: mov eax, var_38
  loc_00E3198B: push ebx
  loc_00E3198C: mov edx, 00000002h
  loc_00E31991: sub esp, 00000010h
  loc_00E31994: mov var_90, edx
  loc_00E3199A: mov ebx, esp
  loc_00E3199C: mov ecx, 0000000Bh
  loc_00E319A1: mov var_88, ecx
  loc_00E319A7: mov edi, [eax]
  loc_00E319A9: mov [ebx], edx
  loc_00E319AB: mov edx, var_8C
  loc_00E319B1: push eax
  loc_00E319B2: mov var_CC, eax
  loc_00E319B8: mov [ebx+00000004h], edx
  loc_00E319BB: mov [ebx+00000008h], ecx
  loc_00E319BE: mov ecx, var_84
  loc_00E319C4: mov [ebx+0000000Ch], ecx
  loc_00E319C7: call [edi+00000028h]
  loc_00E319CA: test eax, eax
  loc_00E319CC: fnclex
  loc_00E319CE: jge 00E319E5h
  loc_00E319D0: mov edx, var_CC
  loc_00E319D6: push 00000028h
  loc_00E319D8: push 006E09E8h
  loc_00E319DD: push edx
  loc_00E319DE: push eax
  loc_00E319DF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E319E5: mov eax, var_3C
  loc_00E319E8: lea edx, var_60
  loc_00E319EB: push edx
  loc_00E319EC: push eax
  loc_00E319ED: mov ecx, [eax]
  loc_00E319EF: mov edi, eax
  loc_00E319F1: call [ecx+00000034h]
  loc_00E319F4: test eax, eax
  loc_00E319F6: fnclex
  loc_00E319F8: jge 00E31A09h
  loc_00E319FA: push 00000034h
  loc_00E319FC: push 006E09F8h
  loc_00E31A01: push edi
  loc_00E31A02: push eax
  loc_00E31A03: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31A09: mov edi, var_DC
  loc_00E31A0F: lea eax, var_60
  loc_00E31A12: push eax
  loc_00E31A13: mov ebx, [edi]
  loc_00E31A15: call [00401034h] ; __vbaStrVarMove
  loc_00E31A1B: mov edx, eax
  loc_00E31A1D: lea ecx, var_28
  loc_00E31A20: call [00401228h] ; __vbaStrMove
  loc_00E31A26: push eax
  loc_00E31A27: push edi
  loc_00E31A28: call [ebx+000000A4h]
  loc_00E31A2E: test eax, eax
  loc_00E31A30: fnclex
  loc_00E31A32: jge 00E31A46h
  loc_00E31A34: push 000000A4h
  loc_00E31A39: push 006DCB70h
  loc_00E31A3E: push edi
  loc_00E31A3F: push eax
  loc_00E31A40: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31A46: lea ecx, var_28
  loc_00E31A49: call [00401258h] ; __vbaFreeStr
  loc_00E31A4F: lea ecx, var_40
  loc_00E31A52: lea edx, var_3C
  loc_00E31A55: push ecx
  loc_00E31A56: lea eax, var_38
  loc_00E31A59: push edx
  loc_00E31A5A: lea ecx, var_34
  loc_00E31A5D: push eax
  loc_00E31A5E: lea edx, var_30
  loc_00E31A61: push ecx
  loc_00E31A62: push edx
  loc_00E31A63: push 00000005h
  loc_00E31A65: call [00401048h] ; __vbaFreeObjList
  loc_00E31A6B: lea eax, var_60
  loc_00E31A6E: lea ecx, var_50
  loc_00E31A71: push eax
  loc_00E31A72: push ecx
  loc_00E31A73: push 00000002h
  loc_00E31A75: call [00401038h] ; __vbaFreeVarList
  loc_00E31A7B: mov eax, [00E3D0D8h]
  loc_00E31A80: add esp, 00000024h
  loc_00E31A83: test eax, eax
  loc_00E31A85: jnz 00E31A9Ch
  loc_00E31A87: push 00E3D0D8h
  loc_00E31A8C: push 006D300Ch
  loc_00E31A91: call [004011A0h] ; __vbaNew2
  loc_00E31A97: mov eax, [00E3D0D8h]
  loc_00E31A9C: mov edx, [eax]
  loc_00E31A9E: push eax
  loc_00E31A9F: call [edx+00000354h]
  loc_00E31AA5: push eax
  loc_00E31AA6: lea eax, var_40
  loc_00E31AA9: push eax
  loc_00E31AAA: call __vbaVarDup
  loc_00E31AAC: mov var_DC, eax
  loc_00E31AB2: mov eax, Me
  loc_00E31AB5: push 006DCBD8h
  loc_00E31ABA: push 00000000h
  loc_00E31ABC: mov ecx, [eax]
  loc_00E31ABE: push 00000018h
  loc_00E31AC0: push eax
  loc_00E31AC1: call [ecx+00000324h]
  loc_00E31AC7: lea edx, var_30
  loc_00E31ACA: push eax
  loc_00E31ACB: push edx
  loc_00E31ACC: call __vbaVarDup
  loc_00E31ACE: push eax
  loc_00E31ACF: lea eax, var_50
  loc_00E31AD2: push eax
  loc_00E31AD3: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E31AD9: add esp, 00000010h
  loc_00E31ADC: push eax
  loc_00E31ADD: call [00401120h] ; __vbaCastObjVar
  loc_00E31AE3: lea ecx, var_34
  loc_00E31AE6: push eax
  loc_00E31AE7: push ecx
  loc_00E31AE8: call __vbaVarDup
  loc_00E31AEA: mov edi, eax
  loc_00E31AEC: lea eax, var_38
  loc_00E31AEF: push eax
  loc_00E31AF0: push edi
  loc_00E31AF1: mov edx, [edi]
  loc_00E31AF3: call [edx+00000054h]
  loc_00E31AF6: test eax, eax
  loc_00E31AF8: fnclex
  loc_00E31AFA: jge 00E31B0Bh
  loc_00E31AFC: push 00000054h
  loc_00E31AFE: push 006DCBE8h
  loc_00E31B03: push edi
  loc_00E31B04: push eax
  loc_00E31B05: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31B0B: lea ebx, var_3C
  loc_00E31B0E: mov eax, var_38
  loc_00E31B11: push ebx
  loc_00E31B12: mov edx, 00000002h
  loc_00E31B17: sub esp, 00000010h
  loc_00E31B1A: mov var_90, edx
  loc_00E31B20: mov ebx, esp
  loc_00E31B22: mov ecx, 0000000Dh
  loc_00E31B27: mov var_88, ecx
  loc_00E31B2D: mov edi, [eax]
  loc_00E31B2F: mov [ebx], edx
  loc_00E31B31: mov edx, var_8C
  loc_00E31B37: push eax
  loc_00E31B38: mov var_CC, eax
  loc_00E31B3E: mov [ebx+00000004h], edx
  loc_00E31B41: mov [ebx+00000008h], ecx
  loc_00E31B44: mov ecx, var_84
  loc_00E31B4A: mov [ebx+0000000Ch], ecx
  loc_00E31B4D: call [edi+00000028h]
  loc_00E31B50: test eax, eax
  loc_00E31B52: fnclex
  loc_00E31B54: jge 00E31B6Bh
  loc_00E31B56: mov edx, var_CC
  loc_00E31B5C: push 00000028h
  loc_00E31B5E: push 006E09E8h
  loc_00E31B63: push edx
  loc_00E31B64: push eax
  loc_00E31B65: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31B6B: mov eax, var_3C
  loc_00E31B6E: lea edx, var_60
  loc_00E31B71: push edx
  loc_00E31B72: push eax
  loc_00E31B73: mov ecx, [eax]
  loc_00E31B75: mov edi, eax
  loc_00E31B77: call [ecx+00000034h]
  loc_00E31B7A: test eax, eax
  loc_00E31B7C: fnclex
  loc_00E31B7E: jge 00E31B8Fh
  loc_00E31B80: push 00000034h
  loc_00E31B82: push 006E09F8h
  loc_00E31B87: push edi
  loc_00E31B88: push eax
  loc_00E31B89: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31B8F: mov eax, var_DC
  loc_00E31B95: mov edi, [00401034h] ; __vbaStrVarMove
  loc_00E31B9B: lea ecx, var_60
  loc_00E31B9E: mov ebx, [eax]
  loc_00E31BA0: push ecx
  loc_00E31BA1: call edi
  loc_00E31BA3: mov edx, eax
  loc_00E31BA5: lea ecx, var_28
  loc_00E31BA8: call [00401228h] ; __vbaStrMove
  loc_00E31BAE: mov edx, ebx
  loc_00E31BB0: mov ebx, var_DC
  loc_00E31BB6: push eax
  loc_00E31BB7: push ebx
  loc_00E31BB8: call [edx+00000054h]
  loc_00E31BBB: test eax, eax
  loc_00E31BBD: fnclex
  loc_00E31BBF: jge 00E31BD0h
  loc_00E31BC1: push 00000054h
  loc_00E31BC3: push 006E0410h
  loc_00E31BC8: push ebx
  loc_00E31BC9: push eax
  loc_00E31BCA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31BD0: lea ecx, var_28
  loc_00E31BD3: call [00401258h] ; __vbaFreeStr
  loc_00E31BD9: lea eax, var_40
  loc_00E31BDC: lea ecx, var_3C
  loc_00E31BDF: push eax
  loc_00E31BE0: lea edx, var_38
  loc_00E31BE3: push ecx
  loc_00E31BE4: lea eax, var_34
  loc_00E31BE7: push edx
  loc_00E31BE8: lea ecx, var_30
  loc_00E31BEB: push eax
  loc_00E31BEC: push ecx
  loc_00E31BED: push 00000005h
  loc_00E31BEF: call [00401048h] ; __vbaFreeObjList
  loc_00E31BF5: lea edx, var_60
  loc_00E31BF8: lea eax, var_50
  loc_00E31BFB: push edx
  loc_00E31BFC: push eax
  loc_00E31BFD: push 00000002h
  loc_00E31BFF: call [00401038h] ; __vbaFreeVarList
  loc_00E31C05: mov ebx, Me
  loc_00E31C08: add esp, 00000024h
  loc_00E31C0B: mov ecx, [ebx]
  loc_00E31C0D: push 006E1958h ; "select * from lc where no_daftar like '" & Chr(37)
  loc_00E31C12: push 00000000h
  loc_00E31C14: push 00000000h
  loc_00E31C16: push ebx
  loc_00E31C17: call [ecx+00000328h]
  loc_00E31C1D: lea edx, var_30
  loc_00E31C20: push eax
  loc_00E31C21: push edx
  loc_00E31C22: call __vbaVarDup
  loc_00E31C24: push eax
  loc_00E31C25: lea eax, var_50
  loc_00E31C28: push eax
  loc_00E31C29: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E31C2F: add esp, 00000010h
  loc_00E31C32: push eax
  loc_00E31C33: call edi
  loc_00E31C35: mov edx, eax
  loc_00E31C37: lea ecx, var_28
  loc_00E31C3A: call [00401228h] ; __vbaStrMove
  loc_00E31C40: mov edi, [0040106Ch] ; __vbaStrCat
  loc_00E31C46: push eax
  loc_00E31C47: call edi
  loc_00E31C49: mov edx, eax
  loc_00E31C4B: lea ecx, var_2C
  loc_00E31C4E: call [00401228h] ; __vbaStrMove
  loc_00E31C54: push eax
  loc_00E31C55: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E31C5A: call edi
  loc_00E31C5C: lea edx, var_60
  loc_00E31C5F: lea ecx, var_24
  loc_00E31C62: mov var_58, eax
  loc_00E31C65: mov var_60, 00000008h
  loc_00E31C6C: call [0040101Ch] ; __vbaVarMove
  loc_00E31C72: lea ecx, var_2C
  loc_00E31C75: lea edx, var_28
  loc_00E31C78: push ecx
  loc_00E31C79: push edx
  loc_00E31C7A: push 00000002h
  loc_00E31C7C: call [004011BCh] ; __vbaFreeStrList
  loc_00E31C82: add esp, 0000000Ch
  loc_00E31C85: lea ecx, var_30
  loc_00E31C88: call [00401254h] ; __vbaFreeObj
  loc_00E31C8E: mov edi, [00401024h] ; __vbaFreeVar
  loc_00E31C94: lea ecx, var_50
  loc_00E31C97: call edi
  loc_00E31C99: lea eax, var_24
  loc_00E31C9C: push eax
  loc_00E31C9D: call [00401230h] ; __vbaStrVarCopy
  loc_00E31CA3: mov var_48, eax
  loc_00E31CA6: mov eax, [00E3D0D8h]
  loc_00E31CAB: test eax, eax
  loc_00E31CAD: mov var_50, 00000008h
  loc_00E31CB4: jnz 00E31CCBh
  loc_00E31CB6: push 00E3D0D8h
  loc_00E31CBB: push 006D300Ch
  loc_00E31CC0: call [004011A0h] ; __vbaNew2
  loc_00E31CC6: mov eax, [00E3D0D8h]
  loc_00E31CCB: mov edx, var_50
  loc_00E31CCE: sub esp, 00000010h
  loc_00E31CD1: mov ecx, esp
  loc_00E31CD3: push 0000000Eh
  loc_00E31CD5: push eax
  loc_00E31CD6: mov [ecx], edx
  loc_00E31CD8: mov edx, var_4C
  loc_00E31CDB: mov [ecx+00000004h], edx
  loc_00E31CDE: mov edx, var_48
  loc_00E31CE1: mov [ecx+00000008h], edx
  loc_00E31CE4: mov edx, var_44
  loc_00E31CE7: mov [ecx+0000000Ch], edx
  loc_00E31CEA: mov ecx, [eax]
  loc_00E31CEC: call [ecx+000004A8h]
  loc_00E31CF2: lea edx, var_30
  loc_00E31CF5: push eax
  loc_00E31CF6: push edx
  loc_00E31CF7: call __vbaVarDup
  loc_00E31CF9: push eax
  loc_00E31CFA: call [00401238h] ; __vbaLateIdSt
  loc_00E31D00: lea ecx, var_30
  loc_00E31D03: call [00401254h] ; __vbaFreeObj
  loc_00E31D09: lea ecx, var_50
  loc_00E31D0C: call edi
  loc_00E31D0E: mov eax, [00E3D0D8h]
  loc_00E31D13: test eax, eax
  loc_00E31D15: jnz 00E31D2Ch
  loc_00E31D17: push 00E3D0D8h
  loc_00E31D1C: push 006D300Ch
  loc_00E31D21: call [004011A0h] ; __vbaNew2
  loc_00E31D27: mov eax, [00E3D0D8h]
  loc_00E31D2C: mov ecx, [eax]
  loc_00E31D2E: push 00000000h
  loc_00E31D30: push 00000019h
  loc_00E31D32: push eax
  loc_00E31D33: call [ecx+000004A8h]
  loc_00E31D39: lea edx, var_30
  loc_00E31D3C: push eax
  loc_00E31D3D: push edx
  loc_00E31D3E: call __vbaVarDup
  loc_00E31D40: mov edi, [00401030h] ; __vbaLateIdCall
  loc_00E31D46: push eax
  loc_00E31D47: call edi
  loc_00E31D49: add esp, 0000000Ch
  loc_00E31D4C: lea ecx, var_30
  loc_00E31D4F: call [00401254h] ; __vbaFreeObj
  loc_00E31D55: mov eax, [ebx]
  loc_00E31D57: push 00000000h
  loc_00E31D59: push FFFFFDDAh
  loc_00E31D5E: push ebx
  loc_00E31D5F: call [eax+00000328h]
  loc_00E31D65: lea ecx, var_30
  loc_00E31D68: push eax
  loc_00E31D69: push ecx
  loc_00E31D6A: call __vbaVarDup
  loc_00E31D6C: push eax
  loc_00E31D6D: call edi
  loc_00E31D6F: add esp, 0000000Ch
  loc_00E31D72: lea ecx, var_30
  loc_00E31D75: call [00401254h] ; __vbaFreeObj
  loc_00E31D7B: mov eax, [00E3D0D8h]
  loc_00E31D80: test eax, eax
  loc_00E31D82: jnz 00E31D99h
  loc_00E31D84: push 00E3D0D8h
  loc_00E31D89: push 006D300Ch
  loc_00E31D8E: call [004011A0h] ; __vbaNew2
  loc_00E31D94: mov eax, [00E3D0D8h]
  loc_00E31D99: mov edx, [eax]
  loc_00E31D9B: push eax
  loc_00E31D9C: call [edx+0000039Ch]
  loc_00E31DA2: push eax
  loc_00E31DA3: lea eax, var_34
  loc_00E31DA6: push eax
  loc_00E31DA7: call __vbaVarDup
  loc_00E31DA9: mov ebx, eax
  loc_00E31DAB: mov eax, [00E3D0D8h]
  loc_00E31DB0: test eax, eax
  loc_00E31DB2: jnz 00E31DC9h
  loc_00E31DB4: push 00E3D0D8h
  loc_00E31DB9: push 006D300Ch
  loc_00E31DBE: call [004011A0h] ; __vbaNew2
  loc_00E31DC4: mov eax, [00E3D0D8h]
  loc_00E31DC9: mov ecx, [eax]
  loc_00E31DCB: push eax
  loc_00E31DCC: call [ecx+0000039Ch]
  loc_00E31DD2: lea edx, var_30
  loc_00E31DD5: push eax
  loc_00E31DD6: push edx
  loc_00E31DD7: call __vbaVarDup
  loc_00E31DD9: mov edi, eax
  loc_00E31DDB: lea ecx, var_28
  loc_00E31DDE: push ecx
  loc_00E31DDF: push edi
  loc_00E31DE0: mov eax, [edi]
  loc_00E31DE2: call [eax+000000A0h]
  loc_00E31DE8: test eax, eax
  loc_00E31DEA: fnclex
  loc_00E31DEC: jge 00E31E00h
  loc_00E31DEE: push 000000A0h
  loc_00E31DF3: push 006DCB70h
  loc_00E31DF8: push edi
  loc_00E31DF9: push eax
  loc_00E31DFA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31E00: mov edx, var_28
  loc_00E31E03: mov edi, [ebx]
  loc_00E31E05: push edx
  loc_00E31E06: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E31E0C: fadd st0, real8 ptr [00401850h]
  loc_00E31E12: sub esp, 00000008h
  loc_00E31E15: fnstsw ax
  loc_00E31E17: test al, 0Dh
  loc_00E31E19: jnz 00E3212Bh
  loc_00E31E1F: fstp real8 ptr [esp]
  loc_00E31E22: call [00401134h] ; __vbaStrR8
  loc_00E31E28: mov edx, eax
  loc_00E31E2A: lea ecx, var_2C
  loc_00E31E2D: call [00401228h] ; __vbaStrMove
  loc_00E31E33: push eax
  loc_00E31E34: push ebx
  loc_00E31E35: call [edi+000000A4h]
  loc_00E31E3B: test eax, eax
  loc_00E31E3D: fnclex
  loc_00E31E3F: jge 00E31E53h
  loc_00E31E41: push 000000A4h
  loc_00E31E46: push 006DCB70h
  loc_00E31E4B: push ebx
  loc_00E31E4C: push eax
  loc_00E31E4D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31E53: lea eax, var_2C
  loc_00E31E56: lea ecx, var_28
  loc_00E31E59: push eax
  loc_00E31E5A: push ecx
  loc_00E31E5B: push 00000002h
  loc_00E31E5D: call [004011BCh] ; __vbaFreeStrList
  loc_00E31E63: lea edx, var_34
  loc_00E31E66: lea eax, var_30
  loc_00E31E69: push edx
  loc_00E31E6A: push eax
  loc_00E31E6B: push 00000002h
  loc_00E31E6D: call [00401048h] ; __vbaFreeObjList
  loc_00E31E73: mov eax, [00E3D0D8h]
  loc_00E31E78: add esp, 00000018h
  loc_00E31E7B: test eax, eax
  loc_00E31E7D: jnz 00E31E94h
  loc_00E31E7F: push 00E3D0D8h
  loc_00E31E84: push 006D300Ch
  loc_00E31E89: call [004011A0h] ; __vbaNew2
  loc_00E31E8F: mov eax, [00E3D0D8h]
  loc_00E31E94: mov ecx, [eax]
  loc_00E31E96: push eax
  loc_00E31E97: call [ecx+00000378h]
  loc_00E31E9D: lea edx, var_30
  loc_00E31EA0: push eax
  loc_00E31EA1: push edx
  loc_00E31EA2: call __vbaVarDup
  loc_00E31EA4: mov edi, eax
  loc_00E31EA6: push 006DCC80h
  loc_00E31EAB: push edi
  loc_00E31EAC: mov eax, [edi]
  loc_00E31EAE: call [eax+000000A4h]
  loc_00E31EB4: test eax, eax
  loc_00E31EB6: fnclex
  loc_00E31EB8: jge 00E31ECCh
  loc_00E31EBA: push 000000A4h
  loc_00E31EBF: push 006DCB70h
  loc_00E31EC4: push edi
  loc_00E31EC5: push eax
  loc_00E31EC6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31ECC: lea ecx, var_30
  loc_00E31ECF: call [00401254h] ; __vbaFreeObj
  loc_00E31ED5: mov eax, [00E3D0D8h]
  loc_00E31EDA: test eax, eax
  loc_00E31EDC: jnz 00E31EF7h
  loc_00E31EDE: mov ebx, [004011A0h] ; __vbaNew2
  loc_00E31EE4: push 00E3D0D8h
  loc_00E31EE9: push 006D300Ch
  loc_00E31EEE: call ebx
  loc_00E31EF0: mov eax, [00E3D0D8h]
  loc_00E31EF5: jmp 00E31EFDh
  loc_00E31EF7: mov ebx, [004011A0h] ; __vbaNew2
  loc_00E31EFD: mov ecx, [eax]
  loc_00E31EFF: push eax
  loc_00E31F00: call [ecx+000003B0h]
  loc_00E31F06: lea edx, var_30
  loc_00E31F09: push eax
  loc_00E31F0A: push edx
  loc_00E31F0B: call __vbaVarDup
  loc_00E31F0D: mov edi, eax
  loc_00E31F0F: push 006E0934h
  loc_00E31F14: push edi
  loc_00E31F15: mov eax, [edi]
  loc_00E31F17: call [eax+00000054h]
  loc_00E31F1A: test eax, eax
  loc_00E31F1C: fnclex
  loc_00E31F1E: jge 00E31F2Fh
  loc_00E31F20: push 00000054h
  loc_00E31F22: push 006E0410h
  loc_00E31F27: push edi
  loc_00E31F28: push eax
  loc_00E31F29: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31F2F: lea ecx, var_30
  loc_00E31F32: call [00401254h] ; __vbaFreeObj
  loc_00E31F38: mov eax, [00E3D0D8h]
  loc_00E31F3D: test eax, eax
  loc_00E31F3F: jnz 00E31F52h
  loc_00E31F41: push 00E3D0D8h
  loc_00E31F46: push 006D300Ch
  loc_00E31F4B: call ebx
  loc_00E31F4D: mov eax, [00E3D0D8h]
  loc_00E31F52: mov ecx, [eax]
  loc_00E31F54: push eax
  loc_00E31F55: call [ecx+00000350h]
  loc_00E31F5B: lea edx, var_30
  loc_00E31F5E: push eax
  loc_00E31F5F: push edx
  loc_00E31F60: call __vbaVarDup
  loc_00E31F62: mov edi, eax
  loc_00E31F64: push FFFFFFFFh
  loc_00E31F66: push edi
  loc_00E31F67: mov eax, [edi]
  loc_00E31F69: call [eax+0000009Ch]
  loc_00E31F6F: test eax, eax
  loc_00E31F71: fnclex
  loc_00E31F73: jge 00E31F87h
  loc_00E31F75: push 0000009Ch
  loc_00E31F7A: push 006DCAD0h
  loc_00E31F7F: push edi
  loc_00E31F80: push eax
  loc_00E31F81: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31F87: lea ecx, var_30
  loc_00E31F8A: call [00401254h] ; __vbaFreeObj
  loc_00E31F90: mov eax, [00E3D0D8h]
  loc_00E31F95: test eax, eax
  loc_00E31F97: jnz 00E31FAAh
  loc_00E31F99: push 00E3D0D8h
  loc_00E31F9E: push 006D300Ch
  loc_00E31FA3: call ebx
  loc_00E31FA5: mov eax, [00E3D0D8h]
  loc_00E31FAA: mov ecx, [eax]
  loc_00E31FAC: push eax
  loc_00E31FAD: call [ecx+00000364h]
  loc_00E31FB3: lea edx, var_30
  loc_00E31FB6: push eax
  loc_00E31FB7: push edx
  loc_00E31FB8: call __vbaVarDup
  loc_00E31FBA: mov edi, eax
  loc_00E31FBC: push 00000000h
  loc_00E31FBE: push edi
  loc_00E31FBF: mov eax, [edi]
  loc_00E31FC1: call [eax+0000009Ch]
  loc_00E31FC7: test eax, eax
  loc_00E31FC9: fnclex
  loc_00E31FCB: jge 00E31FDFh
  loc_00E31FCD: push 0000009Ch
  loc_00E31FD2: push 006DCAD0h
  loc_00E31FD7: push edi
  loc_00E31FD8: push eax
  loc_00E31FD9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E31FDF: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E31FE5: lea ecx, var_30
  loc_00E31FE8: call edi
  loc_00E31FEA: mov eax, [00E3D0D8h]
  loc_00E31FEF: mov var_88, FFFFFFFFh
  loc_00E31FF9: test eax, eax
  loc_00E31FFB: mov var_90, 0000000Bh
  loc_00E32005: jnz 00E32018h
  loc_00E32007: push 00E3D0D8h
  loc_00E3200C: push 006D300Ch
  loc_00E32011: call ebx
  loc_00E32013: mov eax, [00E3D0D8h]
  loc_00E32018: mov edx, var_90
  loc_00E3201E: sub esp, 00000010h
  loc_00E32021: mov ecx, esp
  loc_00E32023: push 80010007h
  loc_00E32028: push eax
  loc_00E32029: mov [ecx], edx
  loc_00E3202B: mov edx, var_8C
  loc_00E32031: mov [ecx+00000004h], edx
  loc_00E32034: mov edx, var_88
  loc_00E3203A: mov [ecx+00000008h], edx
  loc_00E3203D: mov edx, var_84
  loc_00E32043: mov [ecx+0000000Ch], edx
  loc_00E32046: mov ecx, [eax]
  loc_00E32048: call [ecx+00000458h]
  loc_00E3204E: lea edx, var_30
  loc_00E32051: push eax
  loc_00E32052: push edx
  loc_00E32053: call __vbaVarDup
  loc_00E32055: push eax
  loc_00E32056: call [00401238h] ; __vbaLateIdSt
  loc_00E3205C: lea ecx, var_30
  loc_00E3205F: call edi
  loc_00E32061: mov eax, [00E3D9CCh]
  loc_00E32066: test eax, eax
  loc_00E32068: jnz 00E32076h
  loc_00E3206A: push 00E3D9CCh
  loc_00E3206F: push 006DC960h
  loc_00E32074: call ebx
  loc_00E32076: mov eax, Me
  loc_00E32079: mov esi, [00E3D9CCh]
  loc_00E3207F: lea ecx, var_30
  loc_00E32082: push eax
  loc_00E32083: mov ebx, [esi]
  loc_00E32085: push ecx
  loc_00E32086: call [004010B4h] ; __vbaObjSetAddref
  loc_00E3208C: push eax
  loc_00E3208D: push esi
  loc_00E3208E: call [ebx+00000010h]
  loc_00E32091: test eax, eax
  loc_00E32093: fnclex
  loc_00E32095: jge 00E320A6h
  loc_00E32097: push 00000010h
  loc_00E32099: push 006DC950h
  loc_00E3209E: push esi
  loc_00E3209F: push eax
  loc_00E320A0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E320A6: lea ecx, var_30
  loc_00E320A9: call edi
  loc_00E320AB: mov var_4, 00000000h
  loc_00E320B2: fwait
  loc_00E320B3: push 00E3210Ch
  loc_00E320B8: jmp 00E32102h
  loc_00E320BA: lea edx, var_2C
  loc_00E320BD: lea eax, var_28
  loc_00E320C0: push edx
  loc_00E320C1: push eax
  loc_00E320C2: push 00000002h
  loc_00E320C4: call [004011BCh] ; __vbaFreeStrList
  loc_00E320CA: lea ecx, var_40
  loc_00E320CD: lea edx, var_3C
  loc_00E320D0: push ecx
  loc_00E320D1: lea eax, var_38
  loc_00E320D4: push edx
  loc_00E320D5: lea ecx, var_34
  loc_00E320D8: push eax
  loc_00E320D9: lea edx, var_30
  loc_00E320DC: push ecx
  loc_00E320DD: push edx
  loc_00E320DE: push 00000005h
  loc_00E320E0: call [00401048h] ; __vbaFreeObjList
  loc_00E320E6: lea eax, var_80
  loc_00E320E9: lea ecx, var_70
  loc_00E320EC: push eax
  loc_00E320ED: lea edx, var_60
  loc_00E320F0: push ecx
  loc_00E320F1: lea eax, var_50
  loc_00E320F4: push edx
  loc_00E320F5: push eax
  loc_00E320F6: push 00000004h
  loc_00E320F8: call [00401038h] ; __vbaFreeVarList
  loc_00E320FE: add esp, 00000038h
  loc_00E32101: ret
  loc_00E32102: lea ecx, var_24
  loc_00E32105: call [00401024h] ; __vbaFreeVar
  loc_00E3210B: ret
  loc_00E3210C: mov eax, Me
  loc_00E3210F: push eax
  loc_00E32110: mov ecx, [eax]
  loc_00E32112: call [ecx+00000008h]
  loc_00E32115: mov eax, var_4
  loc_00E32118: mov ecx, var_14
  loc_00E3211B: pop edi
  loc_00E3211C: pop esi
  loc_00E3211D: mov fs:[00000000h], ecx
  loc_00E32124: pop ebx
  loc_00E32125: mov esp, ebp
  loc_00E32127: pop ebp
  loc_00E32128: retn 0004h
End Sub
