VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form panggilbio
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
  ClientWidth = 11655
  ClientHeight = 5745
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fcari
    Caption = "Cari Disini"
    BackColor = &HE0E0E0&
    ForeColor = &H404040&
    Left = 90
    Top = 870
    Width = 11475
    Height = 1695
    TabIndex = 2
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.OptionButton optnama
      Caption = "Nama Siswa"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 5100
      Top = 510
      Width = 2445
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
    Begin VB.TextBox txtcari
      BackColor = &HC0C0C0&
      Left = 240
      Top = 1110
      Width = 10905
      Height = 345
      TabIndex = 5
      BorderStyle = 0 'None
    End
    Begin VB.OptionButton optno
      Caption = "No. Daftar"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 2640
      Top = 510
      Width = 1455
      Height = 300
      TabIndex = 4
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
      Caption = "Nama Ibu"
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
    OleObjectBlob = "panggilbio.frx":0000
  End
  Begin MSDataGridLib.DataGrid dgpen
    Left = 90
    Top = 2640
    Width = 11475
    Height = 3015
    TabIndex = 0
    OleObjectBlob = "panggilbio.frx":0332
  End
  Begin VB.Image back
    Picture = "panggilbio.frx":04DD
    Left = 30
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Label Label4
    Caption = "Call Biodata"
    ForeColor = &H80FF80&
    Left = 5220
    Top = 60
    Width = 2745
    Height = 315
    TabIndex = 1
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
  Begin VB.Shape Shape3
    Left = 0
    Top = 450
    Width = 14685
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
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

Attribute VB_Name = "panggilbio"


Private Sub txtcari_KeyPress(KeyAscii As Integer) 'E23870
  loc_00E23870: push ebp
  loc_00E23871: mov ebp, esp
  loc_00E23873: sub esp, 0000000Ch
  loc_00E23876: push 00402836h ; __vbaExceptHandler
  loc_00E2387B: mov eax, fs:[00000000h]
  loc_00E23881: push eax
  loc_00E23882: mov fs:[00000000h], esp
  loc_00E23889: sub esp, 00000054h
  loc_00E2388C: push ebx
  loc_00E2388D: push esi
  loc_00E2388E: push edi
  loc_00E2388F: mov var_C, esp
  loc_00E23892: mov var_8, 00402460h
  loc_00E23899: mov esi, Me
  loc_00E2389C: mov eax, esi
  loc_00E2389E: and eax, 00000001h
  loc_00E238A1: mov var_4, eax
  loc_00E238A4: and esi, FFFFFFFEh
  loc_00E238A7: push esi
  loc_00E238A8: mov Me, esi
  loc_00E238AB: mov ecx, [esi]
  loc_00E238AD: call [ecx+00000004h]
  loc_00E238B0: mov edx, [esi]
  loc_00E238B2: xor eax, eax
  loc_00E238B4: push esi
  loc_00E238B5: mov var_24, eax
  loc_00E238B8: mov var_28, eax
  loc_00E238BB: mov var_2C, eax
  loc_00E238BE: mov var_30, eax
  loc_00E238C1: mov var_40, eax
  loc_00E238C4: mov var_54, eax
  loc_00E238C7: call [edx+00000308h]
  loc_00E238CD: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E238D3: push eax
  loc_00E238D4: lea eax, var_30
  loc_00E238D7: push eax
  loc_00E238D8: call edi
  loc_00E238DA: mov ebx, eax
  loc_00E238DC: lea edx, var_54
  loc_00E238DF: push edx
  loc_00E238E0: push ebx
  loc_00E238E1: mov ecx, [ebx]
  loc_00E238E3: call [ecx+000000E0h]
  loc_00E238E9: test eax, eax
  loc_00E238EB: fnclex
  loc_00E238ED: jge 00E23901h
  loc_00E238EF: push 000000E0h
  loc_00E238F4: push 006E03D4h
  loc_00E238F9: push ebx
  loc_00E238FA: push eax
  loc_00E238FB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23901: mov eax, var_54
  loc_00E23904: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E2390A: lea ecx, var_30
  loc_00E2390D: mov var_60, eax
  loc_00E23910: call ebx
  loc_00E23912: cmp var_60, 0000h
  loc_00E23917: jz 00E239B5h
  loc_00E2391D: mov ecx, [esi]
  loc_00E2391F: push esi
  loc_00E23920: call [ecx+00000304h]
  loc_00E23926: lea edx, var_30
  loc_00E23929: push eax
  loc_00E2392A: push edx
  loc_00E2392B: call edi
  loc_00E2392D: mov ecx, [eax]
  loc_00E2392F: lea edx, var_28
  loc_00E23932: push edx
  loc_00E23933: push eax
  loc_00E23934: mov var_58, eax
  loc_00E23937: call [ecx+000000A0h]
  loc_00E2393D: test eax, eax
  loc_00E2393F: fnclex
  loc_00E23941: jge 00E23958h
  loc_00E23943: mov ecx, var_58
  loc_00E23946: push 000000A0h
  loc_00E2394B: push 006DCB70h
  loc_00E23950: push ecx
  loc_00E23951: push eax
  loc_00E23952: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23958: mov edx, var_28
  loc_00E2395B: push 006E0C48h ; "select * from biodata where no_daftar like '" & Chr(37)
  loc_00E23960: push edx
  loc_00E23961: call [0040106Ch] ; __vbaStrCat
  loc_00E23967: mov edx, eax
  loc_00E23969: lea ecx, var_2C
  loc_00E2396C: call [00401228h] ; __vbaStrMove
  loc_00E23972: push eax
  loc_00E23973: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E23978: call [0040106Ch] ; __vbaStrCat
  loc_00E2397E: lea edx, var_40
  loc_00E23981: lea ecx, var_24
  loc_00E23984: mov var_38, eax
  loc_00E23987: mov var_40, 00000008h
  loc_00E2398E: call [0040101Ch] ; __vbaVarMove
  loc_00E23994: lea eax, var_2C
  loc_00E23997: lea ecx, var_28
  loc_00E2399A: push eax
  loc_00E2399B: push ecx
  loc_00E2399C: push 00000002h
  loc_00E2399E: call [004011BCh] ; __vbaFreeStrList
  loc_00E239A4: add esp, 0000000Ch
  loc_00E239A7: lea ecx, var_30
  loc_00E239AA: call ebx
  loc_00E239AC: lea edx, var_24
  loc_00E239AF: push edx
  loc_00E239B0: jmp 00E23B46h
  loc_00E239B5: mov edx, [esi]
  loc_00E239B7: push esi
  loc_00E239B8: call [edx+0000030Ch]
  loc_00E239BE: push eax
  loc_00E239BF: lea eax, var_30
  loc_00E239C2: push eax
  loc_00E239C3: call edi
  loc_00E239C5: mov ecx, [eax]
  loc_00E239C7: lea edx, var_54
  loc_00E239CA: push edx
  loc_00E239CB: push eax
  loc_00E239CC: mov var_58, eax
  loc_00E239CF: call [ecx+000000E0h]
  loc_00E239D5: test eax, eax
  loc_00E239D7: fnclex
  loc_00E239D9: jge 00E239F0h
  loc_00E239DB: mov ecx, var_58
  loc_00E239DE: push 000000E0h
  loc_00E239E3: push 006E03D4h
  loc_00E239E8: push ecx
  loc_00E239E9: push eax
  loc_00E239EA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E239F0: mov edx, var_54
  loc_00E239F3: lea ecx, var_30
  loc_00E239F6: mov var_60, edx
  loc_00E239F9: call ebx
  loc_00E239FB: cmp var_60, 0000h
  loc_00E23A00: jz 00E23A62h
  loc_00E23A02: mov eax, [esi]
  loc_00E23A04: push esi
  loc_00E23A05: call [eax+00000304h]
  loc_00E23A0B: lea ecx, var_30
  loc_00E23A0E: push eax
  loc_00E23A0F: push ecx
  loc_00E23A10: call edi
  loc_00E23A12: mov edx, [eax]
  loc_00E23A14: lea ecx, var_28
  loc_00E23A17: push ecx
  loc_00E23A18: push eax
  loc_00E23A19: mov var_58, eax
  loc_00E23A1C: call [edx+000000A0h]
  loc_00E23A22: test eax, eax
  loc_00E23A24: fnclex
  loc_00E23A26: jge 00E23A3Dh
  loc_00E23A28: mov edx, var_58
  loc_00E23A2B: push 000000A0h
  loc_00E23A30: push 006DCB70h
  loc_00E23A35: push edx
  loc_00E23A36: push eax
  loc_00E23A37: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23A3D: mov eax, var_28
  loc_00E23A40: push 006E0D20h ; "select * from biodata where nama_ibu like '" & Chr(37)
  loc_00E23A45: push eax
  loc_00E23A46: call [0040106Ch] ; __vbaStrCat
  loc_00E23A4C: mov edx, eax
  loc_00E23A4E: lea ecx, var_2C
  loc_00E23A51: call [00401228h] ; __vbaStrMove
  loc_00E23A57: push eax
  loc_00E23A58: push 006E0D80h ; Chr(37) & "' order by nama_ibu asc"
  loc_00E23A5D: jmp 00E23B0Eh
  loc_00E23A62: mov edx, [esi]
  loc_00E23A64: push esi
  loc_00E23A65: call [edx+00000300h]
  loc_00E23A6B: push eax
  loc_00E23A6C: lea eax, var_30
  loc_00E23A6F: push eax
  loc_00E23A70: call edi
  loc_00E23A72: mov ecx, [eax]
  loc_00E23A74: lea edx, var_54
  loc_00E23A77: push edx
  loc_00E23A78: push eax
  loc_00E23A79: mov var_58, eax
  loc_00E23A7C: call [ecx+000000E0h]
  loc_00E23A82: test eax, eax
  loc_00E23A84: fnclex
  loc_00E23A86: jge 00E23A9Dh
  loc_00E23A88: mov ecx, var_58
  loc_00E23A8B: push 000000E0h
  loc_00E23A90: push 006E03D4h
  loc_00E23A95: push ecx
  loc_00E23A96: push eax
  loc_00E23A97: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23A9D: mov edx, var_54
  loc_00E23AA0: lea ecx, var_30
  loc_00E23AA3: mov var_60, edx
  loc_00E23AA6: call ebx
  loc_00E23AA8: cmp var_60, 0000h
  loc_00E23AAD: jz 00E23BB7h
  loc_00E23AB3: mov eax, [esi]
  loc_00E23AB5: push esi
  loc_00E23AB6: call [eax+00000304h]
  loc_00E23ABC: lea ecx, var_30
  loc_00E23ABF: push eax
  loc_00E23AC0: push ecx
  loc_00E23AC1: call edi
  loc_00E23AC3: mov edx, [eax]
  loc_00E23AC5: lea ecx, var_28
  loc_00E23AC8: push ecx
  loc_00E23AC9: push eax
  loc_00E23ACA: mov var_58, eax
  loc_00E23ACD: call [edx+000000A0h]
  loc_00E23AD3: test eax, eax
  loc_00E23AD5: fnclex
  loc_00E23AD7: jge 00E23AEEh
  loc_00E23AD9: mov edx, var_58
  loc_00E23ADC: push 000000A0h
  loc_00E23AE1: push 006DCB70h
  loc_00E23AE6: push edx
  loc_00E23AE7: push eax
  loc_00E23AE8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23AEE: mov eax, var_28
  loc_00E23AF1: push 006E0DB8h ; "select * from biodata where nama like '" & Chr(37)
  loc_00E23AF6: push eax
  loc_00E23AF7: call [0040106Ch] ; __vbaStrCat
  loc_00E23AFD: mov edx, eax
  loc_00E23AFF: lea ecx, var_2C
  loc_00E23B02: call [00401228h] ; __vbaStrMove
  loc_00E23B08: push eax
  loc_00E23B09: push 006E0E10h ; Chr(37) & "' order by nama asc"
  loc_00E23B0E: call [0040106Ch] ; __vbaStrCat
  loc_00E23B14: lea edx, var_40
  loc_00E23B17: lea ecx, var_24
  loc_00E23B1A: mov var_38, eax
  loc_00E23B1D: mov var_40, 00000008h
  loc_00E23B24: call [0040101Ch] ; __vbaVarMove
  loc_00E23B2A: lea ecx, var_2C
  loc_00E23B2D: lea edx, var_28
  loc_00E23B30: push ecx
  loc_00E23B31: push edx
  loc_00E23B32: push 00000002h
  loc_00E23B34: call [004011BCh] ; __vbaFreeStrList
  loc_00E23B3A: add esp, 0000000Ch
  loc_00E23B3D: lea ecx, var_30
  loc_00E23B40: call ebx
  loc_00E23B42: lea eax, var_24
  loc_00E23B45: push eax
  loc_00E23B46: call [00401230h] ; __vbaStrVarCopy
  loc_00E23B4C: sub esp, 00000010h
  loc_00E23B4F: mov ecx, 00000008h
  loc_00E23B54: mov edx, esp
  loc_00E23B56: mov var_40, ecx
  loc_00E23B59: mov var_38, eax
  loc_00E23B5C: push 0000000Eh
  loc_00E23B5E: mov [edx], ecx
  loc_00E23B60: mov ecx, var_3C
  loc_00E23B63: push esi
  loc_00E23B64: mov [edx+00000004h], ecx
  loc_00E23B67: mov ecx, [esi]
  loc_00E23B69: mov [edx+00000008h], eax
  loc_00E23B6C: mov eax, var_34
  loc_00E23B6F: mov [edx+0000000Ch], eax
  loc_00E23B72: call [ecx+00000324h]
  loc_00E23B78: lea edx, var_30
  loc_00E23B7B: push eax
  loc_00E23B7C: push edx
  loc_00E23B7D: call edi
  loc_00E23B7F: push eax
  loc_00E23B80: call [00401238h] ; __vbaLateIdSt
  loc_00E23B86: lea ecx, var_30
  loc_00E23B89: call ebx
  loc_00E23B8B: lea ecx, var_40
  loc_00E23B8E: call [00401024h] ; __vbaFreeVar
  loc_00E23B94: mov eax, [esi]
  loc_00E23B96: push 00000000h
  loc_00E23B98: push 00000019h
  loc_00E23B9A: push esi
  loc_00E23B9B: call [eax+00000324h]
  loc_00E23BA1: lea ecx, var_30
  loc_00E23BA4: push eax
  loc_00E23BA5: push ecx
  loc_00E23BA6: call edi
  loc_00E23BA8: push eax
  loc_00E23BA9: call [00401030h] ; __vbaLateIdCall
  loc_00E23BAF: add esp, 0000000Ch
  loc_00E23BB2: lea ecx, var_30
  loc_00E23BB5: call ebx
  loc_00E23BB7: mov var_4, 00000000h
  loc_00E23BBE: push 00E23BF5h
  loc_00E23BC3: jmp 00E23BEBh
  loc_00E23BC5: lea edx, var_2C
  loc_00E23BC8: lea eax, var_28
  loc_00E23BCB: push edx
  loc_00E23BCC: push eax
  loc_00E23BCD: push 00000002h
  loc_00E23BCF: call [004011BCh] ; __vbaFreeStrList
  loc_00E23BD5: add esp, 0000000Ch
  loc_00E23BD8: lea ecx, var_30
  loc_00E23BDB: call [00401254h] ; __vbaFreeObj
  loc_00E23BE1: lea ecx, var_40
  loc_00E23BE4: call [00401024h] ; __vbaFreeVar
  loc_00E23BEA: ret
  loc_00E23BEB: lea ecx, var_24
  loc_00E23BEE: call [00401024h] ; __vbaFreeVar
  loc_00E23BF4: ret
  loc_00E23BF5: mov eax, Me
  loc_00E23BF8: push eax
  loc_00E23BF9: mov ecx, [eax]
  loc_00E23BFB: call [ecx+00000008h]
  loc_00E23BFE: mov eax, var_4
  loc_00E23C01: mov ecx, var_14
  loc_00E23C04: pop edi
  loc_00E23C05: pop esi
  loc_00E23C06: mov fs:[00000000h], ecx
  loc_00E23C0D: pop ebx
  loc_00E23C0E: mov esp, ebp
  loc_00E23C10: pop ebp
  loc_00E23C11: retn 0008h
End Sub

Private Sub Form_Load() 'E23680
  loc_00E23680: push ebp
  loc_00E23681: mov ebp, esp
  loc_00E23683: sub esp, 0000000Ch
  loc_00E23686: push 00402836h ; __vbaExceptHandler
  loc_00E2368B: mov eax, fs:[00000000h]
  loc_00E23691: push eax
  loc_00E23692: mov fs:[00000000h], esp
  loc_00E23699: sub esp, 00000040h
  loc_00E2369C: push ebx
  loc_00E2369D: push esi
  loc_00E2369E: push edi
  loc_00E2369F: mov var_C, esp
  loc_00E236A2: mov var_8, 00402450h
  loc_00E236A9: mov esi, Me
  loc_00E236AC: mov eax, esi
  loc_00E236AE: and eax, 00000001h
  loc_00E236B1: mov var_4, eax
  loc_00E236B4: and esi, FFFFFFFEh
  loc_00E236B7: push esi
  loc_00E236B8: mov Me, esi
  loc_00E236BB: mov ecx, [esi]
  loc_00E236BD: call [ecx+00000004h]
  loc_00E236C0: xor eax, eax
  loc_00E236C2: lea edx, var_4C
  loc_00E236C5: mov var_4C, eax
  loc_00E236C8: lea ecx, var_24
  loc_00E236CB: mov var_24, eax
  loc_00E236CE: mov var_28, eax
  loc_00E236D1: mov var_2C, eax
  loc_00E236D4: mov var_3C, eax
  loc_00E236D7: mov var_44, 006E05ACh ; "select * from biodata"
  loc_00E236DE: mov var_4C, 00000008h
  loc_00E236E5: call [00401204h] ; __vbaVarCopy
  loc_00E236EB: lea edx, var_24
  loc_00E236EE: push edx
  loc_00E236EF: call [00401230h] ; __vbaStrVarCopy
  loc_00E236F5: sub esp, 00000010h
  loc_00E236F8: mov ecx, 00000008h
  loc_00E236FD: mov edx, esp
  loc_00E236FF: mov var_3C, ecx
  loc_00E23702: mov var_34, eax
  loc_00E23705: push 0000000Eh
  loc_00E23707: mov [edx], ecx
  loc_00E23709: mov ecx, var_38
  loc_00E2370C: push esi
  loc_00E2370D: mov [edx+00000004h], ecx
  loc_00E23710: mov ecx, [esi]
  loc_00E23712: mov [edx+00000008h], eax
  loc_00E23715: mov eax, var_30
  loc_00E23718: mov [edx+0000000Ch], eax
  loc_00E2371B: call [ecx+00000324h]
  loc_00E23721: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E23727: lea edx, var_28
  loc_00E2372A: push eax
  loc_00E2372B: push edx
  loc_00E2372C: call edi
  loc_00E2372E: push eax
  loc_00E2372F: call [00401238h] ; __vbaLateIdSt
  loc_00E23735: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E2373B: lea ecx, var_28
  loc_00E2373E: call ebx
  loc_00E23740: lea ecx, var_3C
  loc_00E23743: call [00401024h] ; __vbaFreeVar
  loc_00E23749: mov eax, [esi]
  loc_00E2374B: push 00000000h
  loc_00E2374D: push 00000019h
  loc_00E2374F: push esi
  loc_00E23750: call [eax+00000324h]
  loc_00E23756: lea ecx, var_28
  loc_00E23759: push eax
  loc_00E2375A: push ecx
  loc_00E2375B: call edi
  loc_00E2375D: push eax
  loc_00E2375E: call [00401030h] ; __vbaLateIdCall
  loc_00E23764: add esp, 0000000Ch
  loc_00E23767: lea ecx, var_28
  loc_00E2376A: call ebx
  loc_00E2376C: mov edx, [esi]
  loc_00E2376E: push 006E05D8h
  loc_00E23773: push esi
  loc_00E23774: call [edx+00000324h]
  loc_00E2377A: push eax
  loc_00E2377B: lea eax, var_28
  loc_00E2377E: push eax
  loc_00E2377F: call edi
  loc_00E23781: push eax
  loc_00E23782: call [00401224h] ; __vbaCastObj
  loc_00E23788: mov var_34, eax
  loc_00E2378B: mov var_3C, 0000000Dh
  loc_00E23792: lea ecx, var_3C
  loc_00E23795: push ecx
  loc_00E23796: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E2379C: mov ecx, [eax]
  loc_00E2379E: sub esp, 00000010h
  loc_00E237A1: mov edx, esp
  loc_00E237A3: push 00000000h
  loc_00E237A5: push 0000002Ah
  loc_00E237A7: mov [edx], ecx
  loc_00E237A9: mov ecx, [eax+00000004h]
  loc_00E237AC: push esi
  loc_00E237AD: mov [edx+00000004h], ecx
  loc_00E237B0: mov ecx, [eax+00000008h]
  loc_00E237B3: mov eax, [eax+0000000Ch]
  loc_00E237B6: mov [edx+00000008h], ecx
  loc_00E237B9: mov ecx, [esi]
  loc_00E237BB: mov [edx+0000000Ch], eax
  loc_00E237BE: call [ecx+00000328h]
  loc_00E237C4: lea edx, var_2C
  loc_00E237C7: push eax
  loc_00E237C8: push edx
  loc_00E237C9: call edi
  loc_00E237CB: push eax
  loc_00E237CC: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E237D2: lea eax, var_2C
  loc_00E237D5: lea ecx, var_28
  loc_00E237D8: push eax
  loc_00E237D9: push ecx
  loc_00E237DA: push 00000002h
  loc_00E237DC: call [00401048h] ; __vbaFreeObjList
  loc_00E237E2: add esp, 00000028h
  loc_00E237E5: lea ecx, var_3C
  loc_00E237E8: call [00401024h] ; __vbaFreeVar
  loc_00E237EE: mov edx, [esi]
  loc_00E237F0: push 00000000h
  loc_00E237F2: push FFFFFDDAh
  loc_00E237F7: push esi
  loc_00E237F8: call [edx+00000328h]
  loc_00E237FE: push eax
  loc_00E237FF: lea eax, var_28
  loc_00E23802: push eax
  loc_00E23803: call edi
  loc_00E23805: push eax
  loc_00E23806: call [00401030h] ; __vbaLateIdCall
  loc_00E2380C: add esp, 0000000Ch
  loc_00E2380F: lea ecx, var_28
  loc_00E23812: call ebx
  loc_00E23814: mov var_4, 00000000h
  loc_00E2381B: push 00E23849h
  loc_00E23820: jmp 00E2383Fh
  loc_00E23822: lea ecx, var_2C
  loc_00E23825: lea edx, var_28
  loc_00E23828: push ecx
  loc_00E23829: push edx
  loc_00E2382A: push 00000002h
  loc_00E2382C: call [00401048h] ; __vbaFreeObjList
  loc_00E23832: add esp, 0000000Ch
  loc_00E23835: lea ecx, var_3C
  loc_00E23838: call [00401024h] ; __vbaFreeVar
  loc_00E2383E: ret
  loc_00E2383F: lea ecx, var_24
  loc_00E23842: call [00401024h] ; __vbaFreeVar
  loc_00E23848: ret
  loc_00E23849: mov eax, Me
  loc_00E2384C: push eax
  loc_00E2384D: mov ecx, [eax]
  loc_00E2384F: call [ecx+00000008h]
  loc_00E23852: mov eax, var_4
  loc_00E23855: mov ecx, var_14
  loc_00E23858: pop edi
  loc_00E23859: pop esi
  loc_00E2385A: mov fs:[00000000h], ecx
  loc_00E23861: pop ebx
  loc_00E23862: mov esp, ebp
  loc_00E23864: pop ebp
  loc_00E23865: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E1E9B0
  loc_00E1E9B0: push ebp
  loc_00E1E9B1: mov ebp, esp
  loc_00E1E9B3: sub esp, 0000000Ch
  loc_00E1E9B6: push 00402836h ; __vbaExceptHandler
  loc_00E1E9BB: mov eax, fs:[00000000h]
  loc_00E1E9C1: push eax
  loc_00E1E9C2: mov fs:[00000000h], esp
  loc_00E1E9C9: sub esp, 0000005Ch
  loc_00E1E9CC: push ebx
  loc_00E1E9CD: push esi
  loc_00E1E9CE: push edi
  loc_00E1E9CF: mov var_C, esp
  loc_00E1E9D2: mov var_8, 00402420h
  loc_00E1E9D9: mov esi, Me
  loc_00E1E9DC: mov eax, esi
  loc_00E1E9DE: and eax, 00000001h
  loc_00E1E9E1: mov var_4, eax
  loc_00E1E9E4: and esi, FFFFFFFEh
  loc_00E1E9E7: push esi
  loc_00E1E9E8: mov Me, esi
  loc_00E1E9EB: mov ecx, [esi]
  loc_00E1E9ED: call [ecx+00000004h]
  loc_00E1E9F0: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E9F6: xor eax, eax
  loc_00E1E9F8: mov var_18, eax
  loc_00E1E9FB: mov var_4C, eax
  loc_00E1E9FE: mov var_50, eax
  loc_00E1EA01: mov edx, [esi]
  loc_00E1EA03: lea eax, var_4C
  loc_00E1EA06: push eax
  loc_00E1EA07: push esi
  loc_00E1EA08: call [edx+00000070h]
  loc_00E1EA0B: test eax, eax
  loc_00E1EA0D: fnclex
  loc_00E1EA0F: jge 00E1EA1Ch
  loc_00E1EA11: push 00000070h
  loc_00E1EA13: push 006E0134h
  loc_00E1EA18: push esi
  loc_00E1EA19: push eax
  loc_00E1EA1A: call ebx
  loc_00E1EA1C: fld real4 ptr var_4C
  loc_00E1EA1F: fadd st0, real4 ptr [00401298h]
  loc_00E1EA25: mov ecx, [esi]
  loc_00E1EA27: push ecx
  loc_00E1EA28: fnstsw ax
  loc_00E1EA2A: test al, 0Dh
  loc_00E1EA2C: jnz 00E1EBF0h
  loc_00E1EA32: fstp real4 ptr [esp]
  loc_00E1EA35: push esi
  loc_00E1EA36: call [ecx+00000074h]
  loc_00E1EA39: test eax, eax
  loc_00E1EA3B: fnclex
  loc_00E1EA3D: jge 00E1EA4Ah
  loc_00E1EA3F: push 00000074h
  loc_00E1EA41: push 006E0134h
  loc_00E1EA46: push esi
  loc_00E1EA47: push eax
  loc_00E1EA48: call ebx
  loc_00E1EA4A: mov edx, [esi]
  loc_00E1EA4C: lea eax, var_4C
  loc_00E1EA4F: push eax
  loc_00E1EA50: push esi
  loc_00E1EA51: call [edx+00000070h]
  loc_00E1EA54: test eax, eax
  loc_00E1EA56: fnclex
  loc_00E1EA58: jge 00E1EA65h
  loc_00E1EA5A: push 00000070h
  loc_00E1EA5C: push 006E0134h
  loc_00E1EA61: push esi
  loc_00E1EA62: push eax
  loc_00E1EA63: call ebx
  loc_00E1EA65: mov ecx, [esi]
  loc_00E1EA67: lea edx, var_50
  loc_00E1EA6A: push edx
  loc_00E1EA6B: push esi
  loc_00E1EA6C: call [ecx+00000078h]
  loc_00E1EA6F: test eax, eax
  loc_00E1EA71: fnclex
  loc_00E1EA73: jge 00E1EA80h
  loc_00E1EA75: push 00000078h
  loc_00E1EA77: push 006E0134h
  loc_00E1EA7C: push esi
  loc_00E1EA7D: push eax
  loc_00E1EA7E: call ebx
  loc_00E1EA80: sub esp, 00000010h
  loc_00E1EA83: mov ecx, 0000000Ah
  loc_00E1EA88: mov ebx, esp
  loc_00E1EA8A: mov eax, 80020004h
  loc_00E1EA8F: mov edx, eax
  loc_00E1EA91: sub esp, 00000010h
  loc_00E1EA94: mov [ebx], ecx
  loc_00E1EA96: mov ecx, var_44
  loc_00E1EA99: mov edi, [esi]
  loc_00E1EA9B: mov [ebx+00000004h], ecx
  loc_00E1EA9E: mov ecx, esp
  loc_00E1EAA0: sub esp, 00000010h
  loc_00E1EAA3: mov [ebx+00000008h], eax
  loc_00E1EAA6: mov eax, var_3C
  loc_00E1EAA9: mov [ebx+0000000Ch], eax
  loc_00E1EAAC: mov eax, 0000000Ah
  loc_00E1EAB1: mov [ecx], eax
  loc_00E1EAB3: mov eax, var_34
  loc_00E1EAB6: mov [ecx+00000004h], eax
  loc_00E1EAB9: mov eax, 00000004h
  loc_00E1EABE: mov [ecx+00000008h], edx
  loc_00E1EAC1: mov edx, var_2C
  loc_00E1EAC4: mov [ecx+0000000Ch], edx
  loc_00E1EAC7: mov edx, var_24
  loc_00E1EACA: mov ecx, esp
  loc_00E1EACC: mov [ecx], eax
  loc_00E1EACE: mov eax, var_50
  loc_00E1EAD1: mov [ecx+00000004h], edx
  loc_00E1EAD4: mov [ecx+00000008h], eax
  loc_00E1EAD7: mov eax, var_1C
  loc_00E1EADA: mov [ecx+0000000Ch], eax
  loc_00E1EADD: mov ecx, var_4C
  loc_00E1EAE0: push ecx
  loc_00E1EAE1: push esi
  loc_00E1EAE2: call [edi+000002A4h]
  loc_00E1EAE8: test eax, eax
  loc_00E1EAEA: fnclex
  loc_00E1EAEC: jge 00E1EB04h
  loc_00E1EAEE: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EAF4: push 000002A4h
  loc_00E1EAF9: push 006E0134h
  loc_00E1EAFE: push esi
  loc_00E1EAFF: push eax
  loc_00E1EB00: call ebx
  loc_00E1EB02: jmp 00E1EB0Ah
  loc_00E1EB04: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EB0A: call [004010BCh] ; rtcDoEvents
  loc_00E1EB10: mov edx, [esi]
  loc_00E1EB12: lea eax, var_4C
  loc_00E1EB15: push eax
  loc_00E1EB16: push esi
  loc_00E1EB17: call [edx+00000070h]
  loc_00E1EB1A: test eax, eax
  loc_00E1EB1C: fnclex
  loc_00E1EB1E: jge 00E1EB2Bh
  loc_00E1EB20: push 00000070h
  loc_00E1EB22: push 006E0134h
  loc_00E1EB27: push esi
  loc_00E1EB28: push eax
  loc_00E1EB29: call ebx
  loc_00E1EB2B: mov eax, [00E3D9CCh]
  loc_00E1EB30: test eax, eax
  loc_00E1EB32: jnz 00E1EB44h
  loc_00E1EB34: push 00E3D9CCh
  loc_00E1EB39: push 006DC960h
  loc_00E1EB3E: call [004011A0h] ; __vbaNew2
  loc_00E1EB44: mov edi, [00E3D9CCh]
  loc_00E1EB4A: lea edx, var_18
  loc_00E1EB4D: push edx
  loc_00E1EB4E: push edi
  loc_00E1EB4F: mov ecx, [edi]
  loc_00E1EB51: call [ecx+00000018h]
  loc_00E1EB54: test eax, eax
  loc_00E1EB56: fnclex
  loc_00E1EB58: jge 00E1EB65h
  loc_00E1EB5A: push 00000018h
  loc_00E1EB5C: push 006DC950h
  loc_00E1EB61: push edi
  loc_00E1EB62: push eax
  loc_00E1EB63: call ebx
  loc_00E1EB65: mov eax, var_18
  loc_00E1EB68: lea edx, var_50
  loc_00E1EB6B: push edx
  loc_00E1EB6C: push eax
  loc_00E1EB6D: mov ecx, [eax]
  loc_00E1EB6F: mov edi, eax
  loc_00E1EB71: call [ecx+00000098h]
  loc_00E1EB77: test eax, eax
  loc_00E1EB79: fnclex
  loc_00E1EB7B: jge 00E1EB8Bh
  loc_00E1EB7D: push 00000098h
  loc_00E1EB82: push 006DCAF0h
  loc_00E1EB87: push edi
  loc_00E1EB88: push eax
  loc_00E1EB89: call ebx
  loc_00E1EB8B: fld real4 ptr var_4C
  loc_00E1EB8E: fcomp real4 ptr var_50
  loc_00E1EB91: fnstsw ax
  loc_00E1EB93: test ah, 41h
  loc_00E1EB96: jz 00E1EB9Fh
  loc_00E1EB98: mov eax, 00000001h
  loc_00E1EB9D: jmp 00E1EBA1h
  loc_00E1EB9F: xor eax, eax
  loc_00E1EBA1: neg eax
  loc_00E1EBA3: lea ecx, var_18
  loc_00E1EBA6: mov edi, eax
  loc_00E1EBA8: call [00401254h] ; __vbaFreeObj
  loc_00E1EBAE: test di, di
  loc_00E1EBB1: jnz 00E1EA01h
  loc_00E1EBB7: mov var_4, 00000000h
  loc_00E1EBBE: fwait
  loc_00E1EBBF: push 00E1EBD1h
  loc_00E1EBC4: jmp 00E1EBD0h
  loc_00E1EBC6: lea ecx, var_18
  loc_00E1EBC9: call [00401254h] ; __vbaFreeObj
  loc_00E1EBCF: ret
  loc_00E1EBD0: ret
  loc_00E1EBD1: mov eax, Me
  loc_00E1EBD4: push eax
  loc_00E1EBD5: mov ecx, [eax]
  loc_00E1EBD7: call [ecx+00000008h]
  loc_00E1EBDA: mov eax, var_4
  loc_00E1EBDD: mov ecx, var_14
  loc_00E1EBE0: pop edi
  loc_00E1EBE1: pop esi
  loc_00E1EBE2: mov fs:[00000000h], ecx
  loc_00E1EBE9: pop ebx
  loc_00E1EBEA: mov esp, ebp
  loc_00E1EBEC: pop ebp
  loc_00E1EBED: retn 0008h
End Sub

Private Sub back_Click() 'E1EC00
  loc_00E1EC00: push ebp
  loc_00E1EC01: mov ebp, esp
  loc_00E1EC03: sub esp, 0000000Ch
  loc_00E1EC06: push 00402836h ; __vbaExceptHandler
  loc_00E1EC0B: mov eax, fs:[00000000h]
  loc_00E1EC11: push eax
  loc_00E1EC12: mov fs:[00000000h], esp
  loc_00E1EC19: sub esp, 00000018h
  loc_00E1EC1C: push ebx
  loc_00E1EC1D: push esi
  loc_00E1EC1E: push edi
  loc_00E1EC1F: mov var_C, esp
  loc_00E1EC22: mov var_8, 00402430h
  loc_00E1EC29: mov edi, Me
  loc_00E1EC2C: mov eax, edi
  loc_00E1EC2E: and eax, 00000001h
  loc_00E1EC31: mov var_4, eax
  loc_00E1EC34: and edi, FFFFFFFEh
  loc_00E1EC37: push edi
  loc_00E1EC38: mov Me, edi
  loc_00E1EC3B: mov ecx, [edi]
  loc_00E1EC3D: call [ecx+00000004h]
  loc_00E1EC40: mov eax, [00E3D9CCh]
  loc_00E1EC45: xor ebx, ebx
  loc_00E1EC47: cmp eax, ebx
  loc_00E1EC49: mov var_18, ebx
  loc_00E1EC4C: jnz 00E1EC5Eh
  loc_00E1EC4E: push 00E3D9CCh
  loc_00E1EC53: push 006DC960h
  loc_00E1EC58: call [004011A0h] ; __vbaNew2
  loc_00E1EC5E: mov esi, [00E3D9CCh]
  loc_00E1EC64: lea eax, var_18
  loc_00E1EC67: push edi
  loc_00E1EC68: push eax
  loc_00E1EC69: mov edx, [esi]
  loc_00E1EC6B: mov var_2C, edx
  loc_00E1EC6E: call [004010B4h] ; __vbaObjSetAddref
  loc_00E1EC74: mov ecx, var_2C
  loc_00E1EC77: push eax
  loc_00E1EC78: push esi
  loc_00E1EC79: call [ecx+00000010h]
  loc_00E1EC7C: cmp eax, ebx
  loc_00E1EC7E: fnclex
  loc_00E1EC80: jge 00E1EC91h
  loc_00E1EC82: push 00000010h
  loc_00E1EC84: push 006DC950h
  loc_00E1EC89: push esi
  loc_00E1EC8A: push eax
  loc_00E1EC8B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EC91: lea ecx, var_18
  loc_00E1EC94: call [00401254h] ; __vbaFreeObj
  loc_00E1EC9A: mov var_4, ebx
  loc_00E1EC9D: push 00E1ECAFh
  loc_00E1ECA2: jmp 00E1ECAEh
  loc_00E1ECA4: lea ecx, var_18
  loc_00E1ECA7: call [00401254h] ; __vbaFreeObj
  loc_00E1ECAD: ret
  loc_00E1ECAE: ret
  loc_00E1ECAF: mov eax, Me
  loc_00E1ECB2: push eax
  loc_00E1ECB3: mov edx, [eax]
  loc_00E1ECB5: call [edx+00000008h]
  loc_00E1ECB8: mov eax, var_4
  loc_00E1ECBB: mov ecx, var_14
  loc_00E1ECBE: pop edi
  loc_00E1ECBF: pop esi
  loc_00E1ECC0: mov fs:[00000000h], ecx
  loc_00E1ECC7: pop ebx
  loc_00E1ECC8: mov esp, ebp
  loc_00E1ECCA: pop ebp
  loc_00E1ECCB: retn 0004h
End Sub

Private Sub dgpen_Click() 'E1ECD0
  loc_00E1ECD0: push ebp
  loc_00E1ECD1: mov ebp, esp
  loc_00E1ECD3: sub esp, 0000000Ch
  loc_00E1ECD6: push 00402836h ; __vbaExceptHandler
  loc_00E1ECDB: mov eax, fs:[00000000h]
  loc_00E1ECE1: push eax
  loc_00E1ECE2: mov fs:[00000000h], esp
  loc_00E1ECE9: sub esp, 000000D0h
  loc_00E1ECEF: push ebx
  loc_00E1ECF0: push esi
  loc_00E1ECF1: push edi
  loc_00E1ECF2: mov var_C, esp
  loc_00E1ECF5: mov var_8, 00402440h
  loc_00E1ECFC: mov ebx, Me
  loc_00E1ECFF: mov eax, ebx
  loc_00E1ED01: and eax, 00000001h
  loc_00E1ED04: mov var_4, eax
  loc_00E1ED07: and ebx, FFFFFFFEh
  loc_00E1ED0A: push ebx
  loc_00E1ED0B: mov Me, ebx
  loc_00E1ED0E: mov ecx, [ebx]
  loc_00E1ED10: call [ecx+00000004h]
  loc_00E1ED13: mov eax, [00E3D060h]
  loc_00E1ED18: xor ecx, ecx
  loc_00E1ED1A: cmp eax, ecx
  loc_00E1ED1C: mov var_18, ecx
  loc_00E1ED1F: mov var_1C, ecx
  loc_00E1ED22: mov var_20, ecx
  loc_00E1ED25: mov var_24, ecx
  loc_00E1ED28: mov var_28, ecx
  loc_00E1ED2B: mov var_2C, ecx
  loc_00E1ED2E: mov var_30, ecx
  loc_00E1ED31: mov var_40, ecx
  loc_00E1ED34: mov var_50, ecx
  loc_00E1ED37: mov var_60, ecx
  loc_00E1ED3A: mov var_70, ecx
  loc_00E1ED3D: mov var_80, ecx
  loc_00E1ED40: mov var_90, ecx
  loc_00E1ED46: jnz 00E1ED5Dh
  loc_00E1ED48: push 00E3D060h
  loc_00E1ED4D: push 006D87C0h
  loc_00E1ED52: call [004011A0h] ; __vbaNew2
  loc_00E1ED58: mov eax, [00E3D060h]
  loc_00E1ED5D: mov edx, [eax]
  loc_00E1ED5F: push eax
  loc_00E1ED60: call [edx+00000520h]
  loc_00E1ED66: mov esi, [004010ACh] ; __vbaObjSet
  loc_00E1ED6C: push eax
  loc_00E1ED6D: lea eax, var_30
  loc_00E1ED70: push eax
  loc_00E1ED71: call __vbaObjSet
  loc_00E1ED73: mov ecx, [ebx]
  loc_00E1ED75: push 006DCBD8h
  loc_00E1ED7A: push 00000000h
  loc_00E1ED7C: push 00000018h
  loc_00E1ED7E: push ebx
  loc_00E1ED7F: mov var_CC, eax
  loc_00E1ED85: call [ecx+00000324h]
  loc_00E1ED8B: lea edx, var_20
  loc_00E1ED8E: push eax
  loc_00E1ED8F: push edx
  loc_00E1ED90: call __vbaObjSet
  loc_00E1ED92: push eax
  loc_00E1ED93: lea eax, var_40
  loc_00E1ED96: push eax
  loc_00E1ED97: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1ED9D: add esp, 00000010h
  loc_00E1EDA0: push eax
  loc_00E1EDA1: call [00401120h] ; __vbaCastObjVar
  loc_00E1EDA7: lea ecx, var_24
  loc_00E1EDAA: push eax
  loc_00E1EDAB: push ecx
  loc_00E1EDAC: call __vbaObjSet
  loc_00E1EDAE: mov edi, eax
  loc_00E1EDB0: lea eax, var_28
  loc_00E1EDB3: push eax
  loc_00E1EDB4: push edi
  loc_00E1EDB5: mov edx, [edi]
  loc_00E1EDB7: call [edx+00000054h]
  loc_00E1EDBA: test eax, eax
  loc_00E1EDBC: fnclex
  loc_00E1EDBE: jge 00E1EDCFh
  loc_00E1EDC0: push 00000054h
  loc_00E1EDC2: push 006DCBE8h
  loc_00E1EDC7: push edi
  loc_00E1EDC8: push eax
  loc_00E1EDC9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EDCF: lea edi, var_2C
  loc_00E1EDD2: mov eax, var_28
  loc_00E1EDD5: push edi
  loc_00E1EDD6: mov ecx, 00000002h
  loc_00E1EDDB: sub esp, 00000010h
  loc_00E1EDDE: mov var_80, ecx
  loc_00E1EDE1: mov edi, esp
  loc_00E1EDE3: mov var_78, 00000000h
  loc_00E1EDEA: mov edx, [eax]
  loc_00E1EDEC: push eax
  loc_00E1EDED: mov [edi], ecx
  loc_00E1EDEF: mov ecx, var_7C
  loc_00E1EDF2: mov var_BC, eax
  loc_00E1EDF8: mov [edi+00000004h], ecx
  loc_00E1EDFB: mov ecx, var_78
  loc_00E1EDFE: mov [edi+00000008h], ecx
  loc_00E1EE01: mov ecx, var_74
  loc_00E1EE04: mov [edi+0000000Ch], ecx
  loc_00E1EE07: call [edx+00000028h]
  loc_00E1EE0A: test eax, eax
  loc_00E1EE0C: fnclex
  loc_00E1EE0E: jge 00E1EE25h
  loc_00E1EE10: mov edx, var_BC
  loc_00E1EE16: push 00000028h
  loc_00E1EE18: push 006E09E8h
  loc_00E1EE1D: push edx
  loc_00E1EE1E: push eax
  loc_00E1EE1F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EE25: mov eax, var_2C
  loc_00E1EE28: lea edx, var_50
  loc_00E1EE2B: push edx
  loc_00E1EE2C: push eax
  loc_00E1EE2D: mov ecx, [eax]
  loc_00E1EE2F: mov edi, eax
  loc_00E1EE31: call [ecx+00000034h]
  loc_00E1EE34: test eax, eax
  loc_00E1EE36: fnclex
  loc_00E1EE38: jge 00E1EE49h
  loc_00E1EE3A: push 00000034h
  loc_00E1EE3C: push 006E09F8h
  loc_00E1EE41: push edi
  loc_00E1EE42: push eax
  loc_00E1EE43: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EE49: mov eax, var_CC
  loc_00E1EE4F: lea ecx, var_50
  loc_00E1EE52: push ecx
  loc_00E1EE53: mov edi, [eax]
  loc_00E1EE55: call [00401034h] ; __vbaStrVarMove
  loc_00E1EE5B: mov edx, eax
  loc_00E1EE5D: lea ecx, var_18
  loc_00E1EE60: call [00401228h] ; __vbaStrMove
  loc_00E1EE66: mov edx, edi
  loc_00E1EE68: mov edi, var_CC
  loc_00E1EE6E: push eax
  loc_00E1EE6F: push edi
  loc_00E1EE70: call [edx+00000054h]
  loc_00E1EE73: test eax, eax
  loc_00E1EE75: fnclex
  loc_00E1EE77: jge 00E1EE88h
  loc_00E1EE79: push 00000054h
  loc_00E1EE7B: push 006E0410h
  loc_00E1EE80: push edi
  loc_00E1EE81: push eax
  loc_00E1EE82: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EE88: lea ecx, var_18
  loc_00E1EE8B: call [00401258h] ; __vbaFreeStr
  loc_00E1EE91: lea eax, var_30
  loc_00E1EE94: lea ecx, var_2C
  loc_00E1EE97: push eax
  loc_00E1EE98: lea edx, var_28
  loc_00E1EE9B: push ecx
  loc_00E1EE9C: lea eax, var_24
  loc_00E1EE9F: push edx
  loc_00E1EEA0: lea ecx, var_20
  loc_00E1EEA3: push eax
  loc_00E1EEA4: push ecx
  loc_00E1EEA5: push 00000005h
  loc_00E1EEA7: call [00401048h] ; __vbaFreeObjList
  loc_00E1EEAD: lea edx, var_50
  loc_00E1EEB0: lea eax, var_40
  loc_00E1EEB3: push edx
  loc_00E1EEB4: push eax
  loc_00E1EEB5: push 00000002h
  loc_00E1EEB7: call [00401038h] ; __vbaFreeVarList
  loc_00E1EEBD: mov eax, [00E3D060h]
  loc_00E1EEC2: add esp, 00000024h
  loc_00E1EEC5: test eax, eax
  loc_00E1EEC7: jnz 00E1EEDEh
  loc_00E1EEC9: push 00E3D060h
  loc_00E1EECE: push 006D87C0h
  loc_00E1EED3: call [004011A0h] ; __vbaNew2
  loc_00E1EED9: mov eax, [00E3D060h]
  loc_00E1EEDE: mov ecx, [eax]
  loc_00E1EEE0: push eax
  loc_00E1EEE1: call [ecx+0000051Ch]
  loc_00E1EEE7: lea edx, var_30
  loc_00E1EEEA: push eax
  loc_00E1EEEB: push edx
  loc_00E1EEEC: call __vbaObjSet
  loc_00E1EEEE: push 006DCBD8h
  loc_00E1EEF3: mov var_CC, eax
  loc_00E1EEF9: mov eax, [ebx]
  loc_00E1EEFB: push 00000000h
  loc_00E1EEFD: push 00000018h
  loc_00E1EEFF: push ebx
  loc_00E1EF00: call [eax+00000324h]
  loc_00E1EF06: lea ecx, var_20
  loc_00E1EF09: push eax
  loc_00E1EF0A: push ecx
  loc_00E1EF0B: call __vbaObjSet
  loc_00E1EF0D: lea edx, var_40
  loc_00E1EF10: push eax
  loc_00E1EF11: push edx
  loc_00E1EF12: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1EF18: add esp, 00000010h
  loc_00E1EF1B: push eax
  loc_00E1EF1C: call [00401120h] ; __vbaCastObjVar
  loc_00E1EF22: push eax
  loc_00E1EF23: lea eax, var_24
  loc_00E1EF26: push eax
  loc_00E1EF27: call __vbaObjSet
  loc_00E1EF29: mov edi, eax
  loc_00E1EF2B: lea edx, var_28
  loc_00E1EF2E: push edx
  loc_00E1EF2F: push edi
  loc_00E1EF30: mov ecx, [edi]
  loc_00E1EF32: call [ecx+00000054h]
  loc_00E1EF35: test eax, eax
  loc_00E1EF37: fnclex
  loc_00E1EF39: jge 00E1EF4Ah
  loc_00E1EF3B: push 00000054h
  loc_00E1EF3D: push 006DCBE8h
  loc_00E1EF42: push edi
  loc_00E1EF43: push eax
  loc_00E1EF44: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EF4A: lea ebx, var_2C
  loc_00E1EF4D: mov eax, var_28
  loc_00E1EF50: push ebx
  loc_00E1EF51: mov edx, 00000002h
  loc_00E1EF56: sub esp, 00000010h
  loc_00E1EF59: mov var_80, edx
  loc_00E1EF5C: mov ebx, esp
  loc_00E1EF5E: mov ecx, 00000001h
  loc_00E1EF63: mov var_78, ecx
  loc_00E1EF66: mov edi, [eax]
  loc_00E1EF68: mov [ebx], edx
  loc_00E1EF6A: mov edx, var_7C
  loc_00E1EF6D: push eax
  loc_00E1EF6E: mov var_BC, eax
  loc_00E1EF74: mov [ebx+00000004h], edx
  loc_00E1EF77: mov [ebx+00000008h], ecx
  loc_00E1EF7A: mov ecx, var_74
  loc_00E1EF7D: mov [ebx+0000000Ch], ecx
  loc_00E1EF80: call [edi+00000028h]
  loc_00E1EF83: test eax, eax
  loc_00E1EF85: fnclex
  loc_00E1EF87: jge 00E1EF9Eh
  loc_00E1EF89: mov edx, var_BC
  loc_00E1EF8F: push 00000028h
  loc_00E1EF91: push 006E09E8h
  loc_00E1EF96: push edx
  loc_00E1EF97: push eax
  loc_00E1EF98: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EF9E: mov eax, var_2C
  loc_00E1EFA1: lea edx, var_50
  loc_00E1EFA4: push edx
  loc_00E1EFA5: push eax
  loc_00E1EFA6: mov ecx, [eax]
  loc_00E1EFA8: mov edi, eax
  loc_00E1EFAA: call [ecx+00000034h]
  loc_00E1EFAD: test eax, eax
  loc_00E1EFAF: fnclex
  loc_00E1EFB1: jge 00E1EFC2h
  loc_00E1EFB3: push 00000034h
  loc_00E1EFB5: push 006E09F8h
  loc_00E1EFBA: push edi
  loc_00E1EFBB: push eax
  loc_00E1EFBC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EFC2: mov edi, var_CC
  loc_00E1EFC8: lea eax, var_50
  loc_00E1EFCB: push eax
  loc_00E1EFCC: mov ebx, [edi]
  loc_00E1EFCE: call [00401034h] ; __vbaStrVarMove
  loc_00E1EFD4: mov edx, eax
  loc_00E1EFD6: lea ecx, var_18
  loc_00E1EFD9: call [00401228h] ; __vbaStrMove
  loc_00E1EFDF: push eax
  loc_00E1EFE0: push edi
  loc_00E1EFE1: call [ebx+00000054h]
  loc_00E1EFE4: test eax, eax
  loc_00E1EFE6: fnclex
  loc_00E1EFE8: jge 00E1EFF9h
  loc_00E1EFEA: push 00000054h
  loc_00E1EFEC: push 006E0410h
  loc_00E1EFF1: push edi
  loc_00E1EFF2: push eax
  loc_00E1EFF3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1EFF9: lea ecx, var_18
  loc_00E1EFFC: call [00401258h] ; __vbaFreeStr
  loc_00E1F002: lea ecx, var_30
  loc_00E1F005: lea edx, var_2C
  loc_00E1F008: push ecx
  loc_00E1F009: lea eax, var_28
  loc_00E1F00C: push edx
  loc_00E1F00D: lea ecx, var_24
  loc_00E1F010: push eax
  loc_00E1F011: lea edx, var_20
  loc_00E1F014: push ecx
  loc_00E1F015: push edx
  loc_00E1F016: push 00000005h
  loc_00E1F018: call [00401048h] ; __vbaFreeObjList
  loc_00E1F01E: lea eax, var_50
  loc_00E1F021: lea ecx, var_40
  loc_00E1F024: push eax
  loc_00E1F025: push ecx
  loc_00E1F026: push 00000002h
  loc_00E1F028: call [00401038h] ; __vbaFreeVarList
  loc_00E1F02E: mov eax, [00E3D060h]
  loc_00E1F033: add esp, 00000024h
  loc_00E1F036: test eax, eax
  loc_00E1F038: jnz 00E1F04Fh
  loc_00E1F03A: push 00E3D060h
  loc_00E1F03F: push 006D87C0h
  loc_00E1F044: call [004011A0h] ; __vbaNew2
  loc_00E1F04A: mov eax, [00E3D060h]
  loc_00E1F04F: mov edx, [eax]
  loc_00E1F051: push eax
  loc_00E1F052: call [edx+00000518h]
  loc_00E1F058: push eax
  loc_00E1F059: lea eax, var_30
  loc_00E1F05C: push eax
  loc_00E1F05D: call __vbaObjSet
  loc_00E1F05F: mov ebx, Me
  loc_00E1F062: push 006DCBD8h
  loc_00E1F067: push 00000000h
  loc_00E1F069: push 00000018h
  loc_00E1F06B: mov ecx, [ebx]
  loc_00E1F06D: push ebx
  loc_00E1F06E: mov var_CC, eax
  loc_00E1F074: call [ecx+00000324h]
  loc_00E1F07A: lea edx, var_20
  loc_00E1F07D: push eax
  loc_00E1F07E: push edx
  loc_00E1F07F: call __vbaObjSet
  loc_00E1F081: push eax
  loc_00E1F082: lea eax, var_40
  loc_00E1F085: push eax
  loc_00E1F086: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1F08C: add esp, 00000010h
  loc_00E1F08F: push eax
  loc_00E1F090: call [00401120h] ; __vbaCastObjVar
  loc_00E1F096: lea ecx, var_24
  loc_00E1F099: push eax
  loc_00E1F09A: push ecx
  loc_00E1F09B: call __vbaObjSet
  loc_00E1F09D: mov edi, eax
  loc_00E1F09F: lea eax, var_28
  loc_00E1F0A2: push eax
  loc_00E1F0A3: push edi
  loc_00E1F0A4: mov edx, [edi]
  loc_00E1F0A6: call [edx+00000054h]
  loc_00E1F0A9: test eax, eax
  loc_00E1F0AB: fnclex
  loc_00E1F0AD: jge 00E1F0BEh
  loc_00E1F0AF: push 00000054h
  loc_00E1F0B1: push 006DCBE8h
  loc_00E1F0B6: push edi
  loc_00E1F0B7: push eax
  loc_00E1F0B8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F0BE: lea edi, var_2C
  loc_00E1F0C1: mov eax, var_28
  loc_00E1F0C4: push edi
  loc_00E1F0C5: mov ecx, 00000002h
  loc_00E1F0CA: sub esp, 00000010h
  loc_00E1F0CD: mov var_80, ecx
  loc_00E1F0D0: mov edi, esp
  loc_00E1F0D2: mov var_78, 00000013h
  loc_00E1F0D9: mov edx, [eax]
  loc_00E1F0DB: push eax
  loc_00E1F0DC: mov [edi], ecx
  loc_00E1F0DE: mov ecx, var_7C
  loc_00E1F0E1: mov var_BC, eax
  loc_00E1F0E7: mov [edi+00000004h], ecx
  loc_00E1F0EA: mov ecx, var_78
  loc_00E1F0ED: mov [edi+00000008h], ecx
  loc_00E1F0F0: mov ecx, var_74
  loc_00E1F0F3: mov [edi+0000000Ch], ecx
  loc_00E1F0F6: call [edx+00000028h]
  loc_00E1F0F9: test eax, eax
  loc_00E1F0FB: fnclex
  loc_00E1F0FD: jge 00E1F114h
  loc_00E1F0FF: mov edx, var_BC
  loc_00E1F105: push 00000028h
  loc_00E1F107: push 006E09E8h
  loc_00E1F10C: push edx
  loc_00E1F10D: push eax
  loc_00E1F10E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F114: mov eax, var_2C
  loc_00E1F117: lea edx, var_50
  loc_00E1F11A: push edx
  loc_00E1F11B: push eax
  loc_00E1F11C: mov ecx, [eax]
  loc_00E1F11E: mov edi, eax
  loc_00E1F120: call [ecx+00000034h]
  loc_00E1F123: test eax, eax
  loc_00E1F125: fnclex
  loc_00E1F127: jge 00E1F138h
  loc_00E1F129: push 00000034h
  loc_00E1F12B: push 006E09F8h
  loc_00E1F130: push edi
  loc_00E1F131: push eax
  loc_00E1F132: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F138: mov eax, var_CC
  loc_00E1F13E: lea ecx, var_50
  loc_00E1F141: push ecx
  loc_00E1F142: mov edi, [eax]
  loc_00E1F144: call [00401034h] ; __vbaStrVarMove
  loc_00E1F14A: mov edx, eax
  loc_00E1F14C: lea ecx, var_18
  loc_00E1F14F: call [00401228h] ; __vbaStrMove
  loc_00E1F155: mov edx, edi
  loc_00E1F157: mov edi, var_CC
  loc_00E1F15D: push eax
  loc_00E1F15E: push edi
  loc_00E1F15F: call [edx+00000054h]
  loc_00E1F162: test eax, eax
  loc_00E1F164: fnclex
  loc_00E1F166: jge 00E1F177h
  loc_00E1F168: push 00000054h
  loc_00E1F16A: push 006E0410h
  loc_00E1F16F: push edi
  loc_00E1F170: push eax
  loc_00E1F171: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F177: lea ecx, var_18
  loc_00E1F17A: call [00401258h] ; __vbaFreeStr
  loc_00E1F180: lea eax, var_30
  loc_00E1F183: lea ecx, var_2C
  loc_00E1F186: push eax
  loc_00E1F187: lea edx, var_28
  loc_00E1F18A: push ecx
  loc_00E1F18B: lea eax, var_24
  loc_00E1F18E: push edx
  loc_00E1F18F: lea ecx, var_20
  loc_00E1F192: push eax
  loc_00E1F193: push ecx
  loc_00E1F194: push 00000005h
  loc_00E1F196: call [00401048h] ; __vbaFreeObjList
  loc_00E1F19C: lea edx, var_50
  loc_00E1F19F: lea eax, var_40
  loc_00E1F1A2: push edx
  loc_00E1F1A3: push eax
  loc_00E1F1A4: push 00000002h
  loc_00E1F1A6: call [00401038h] ; __vbaFreeVarList
  loc_00E1F1AC: mov eax, [00E3D060h]
  loc_00E1F1B1: add esp, 00000024h
  loc_00E1F1B4: test eax, eax
  loc_00E1F1B6: jnz 00E1F1CDh
  loc_00E1F1B8: push 00E3D060h
  loc_00E1F1BD: push 006D87C0h
  loc_00E1F1C2: call [004011A0h] ; __vbaNew2
  loc_00E1F1C8: mov eax, [00E3D060h]
  loc_00E1F1CD: mov ecx, [eax]
  loc_00E1F1CF: push eax
  loc_00E1F1D0: call [ecx+000004E4h]
  loc_00E1F1D6: lea edx, var_30
  loc_00E1F1D9: push eax
  loc_00E1F1DA: push edx
  loc_00E1F1DB: call __vbaObjSet
  loc_00E1F1DD: push 006DCBD8h
  loc_00E1F1E2: mov var_CC, eax
  loc_00E1F1E8: mov eax, [ebx]
  loc_00E1F1EA: push 00000000h
  loc_00E1F1EC: push 00000018h
  loc_00E1F1EE: push ebx
  loc_00E1F1EF: call [eax+00000324h]
  loc_00E1F1F5: lea ecx, var_20
  loc_00E1F1F8: push eax
  loc_00E1F1F9: push ecx
  loc_00E1F1FA: call __vbaObjSet
  loc_00E1F1FC: lea edx, var_40
  loc_00E1F1FF: push eax
  loc_00E1F200: push edx
  loc_00E1F201: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1F207: add esp, 00000010h
  loc_00E1F20A: push eax
  loc_00E1F20B: call [00401120h] ; __vbaCastObjVar
  loc_00E1F211: push eax
  loc_00E1F212: lea eax, var_24
  loc_00E1F215: push eax
  loc_00E1F216: call __vbaObjSet
  loc_00E1F218: mov edi, eax
  loc_00E1F21A: lea edx, var_28
  loc_00E1F21D: push edx
  loc_00E1F21E: push edi
  loc_00E1F21F: mov ecx, [edi]
  loc_00E1F221: call [ecx+00000054h]
  loc_00E1F224: test eax, eax
  loc_00E1F226: fnclex
  loc_00E1F228: jge 00E1F239h
  loc_00E1F22A: push 00000054h
  loc_00E1F22C: push 006DCBE8h
  loc_00E1F231: push edi
  loc_00E1F232: push eax
  loc_00E1F233: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F239: lea edi, var_2C
  loc_00E1F23C: mov eax, var_28
  loc_00E1F23F: push edi
  loc_00E1F240: mov ecx, 00000002h
  loc_00E1F245: sub esp, 00000010h
  loc_00E1F248: mov var_78, ecx
  loc_00E1F24B: mov edi, esp
  loc_00E1F24D: mov var_80, ecx
  loc_00E1F250: mov edx, [eax]
  loc_00E1F252: push eax
  loc_00E1F253: mov [edi], ecx
  loc_00E1F255: mov ecx, var_7C
  loc_00E1F258: mov var_BC, eax
  loc_00E1F25E: mov [edi+00000004h], ecx
  loc_00E1F261: mov ecx, var_78
  loc_00E1F264: mov [edi+00000008h], ecx
  loc_00E1F267: mov ecx, var_74
  loc_00E1F26A: mov [edi+0000000Ch], ecx
  loc_00E1F26D: call [edx+00000028h]
  loc_00E1F270: test eax, eax
  loc_00E1F272: fnclex
  loc_00E1F274: jge 00E1F28Bh
  loc_00E1F276: mov edx, var_BC
  loc_00E1F27C: push 00000028h
  loc_00E1F27E: push 006E09E8h
  loc_00E1F283: push edx
  loc_00E1F284: push eax
  loc_00E1F285: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F28B: mov eax, var_2C
  loc_00E1F28E: lea edx, var_50
  loc_00E1F291: push edx
  loc_00E1F292: push eax
  loc_00E1F293: mov ecx, [eax]
  loc_00E1F295: mov edi, eax
  loc_00E1F297: call [ecx+00000034h]
  loc_00E1F29A: test eax, eax
  loc_00E1F29C: fnclex
  loc_00E1F29E: jge 00E1F2AFh
  loc_00E1F2A0: push 00000034h
  loc_00E1F2A2: push 006E09F8h
  loc_00E1F2A7: push edi
  loc_00E1F2A8: push eax
  loc_00E1F2A9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F2AF: mov eax, var_CC
  loc_00E1F2B5: lea ecx, var_50
  loc_00E1F2B8: push ecx
  loc_00E1F2B9: mov edi, [eax]
  loc_00E1F2BB: call [00401034h] ; __vbaStrVarMove
  loc_00E1F2C1: mov edx, eax
  loc_00E1F2C3: lea ecx, var_18
  loc_00E1F2C6: call [00401228h] ; __vbaStrMove
  loc_00E1F2CC: mov edx, edi
  loc_00E1F2CE: mov edi, var_CC
  loc_00E1F2D4: push eax
  loc_00E1F2D5: push edi
  loc_00E1F2D6: call [edx+00000054h]
  loc_00E1F2D9: test eax, eax
  loc_00E1F2DB: fnclex
  loc_00E1F2DD: jge 00E1F2EEh
  loc_00E1F2DF: push 00000054h
  loc_00E1F2E1: push 006E0410h
  loc_00E1F2E6: push edi
  loc_00E1F2E7: push eax
  loc_00E1F2E8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F2EE: lea ecx, var_18
  loc_00E1F2F1: call [00401258h] ; __vbaFreeStr
  loc_00E1F2F7: lea eax, var_30
  loc_00E1F2FA: lea ecx, var_2C
  loc_00E1F2FD: push eax
  loc_00E1F2FE: lea edx, var_28
  loc_00E1F301: push ecx
  loc_00E1F302: lea eax, var_24
  loc_00E1F305: push edx
  loc_00E1F306: lea ecx, var_20
  loc_00E1F309: push eax
  loc_00E1F30A: push ecx
  loc_00E1F30B: push 00000005h
  loc_00E1F30D: call [00401048h] ; __vbaFreeObjList
  loc_00E1F313: lea edx, var_50
  loc_00E1F316: lea eax, var_40
  loc_00E1F319: push edx
  loc_00E1F31A: push eax
  loc_00E1F31B: push 00000002h
  loc_00E1F31D: call [00401038h] ; __vbaFreeVarList
  loc_00E1F323: mov eax, [00E3D060h]
  loc_00E1F328: add esp, 00000024h
  loc_00E1F32B: test eax, eax
  loc_00E1F32D: jnz 00E1F344h
  loc_00E1F32F: push 00E3D060h
  loc_00E1F334: push 006D87C0h
  loc_00E1F339: call [004011A0h] ; __vbaNew2
  loc_00E1F33F: mov eax, [00E3D060h]
  loc_00E1F344: mov ecx, [eax]
  loc_00E1F346: push eax
  loc_00E1F347: call [ecx+000004E0h]
  loc_00E1F34D: lea edx, var_30
  loc_00E1F350: push eax
  loc_00E1F351: push edx
  loc_00E1F352: call __vbaObjSet
  loc_00E1F354: push 006DCBD8h
  loc_00E1F359: mov var_CC, eax
  loc_00E1F35F: mov eax, [ebx]
  loc_00E1F361: push 00000000h
  loc_00E1F363: push 00000018h
  loc_00E1F365: push ebx
  loc_00E1F366: call [eax+00000324h]
  loc_00E1F36C: lea ecx, var_20
  loc_00E1F36F: push eax
  loc_00E1F370: push ecx
  loc_00E1F371: call __vbaObjSet
  loc_00E1F373: lea edx, var_40
  loc_00E1F376: push eax
  loc_00E1F377: push edx
  loc_00E1F378: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1F37E: add esp, 00000010h
  loc_00E1F381: push eax
  loc_00E1F382: call [00401120h] ; __vbaCastObjVar
  loc_00E1F388: push eax
  loc_00E1F389: lea eax, var_24
  loc_00E1F38C: push eax
  loc_00E1F38D: call __vbaObjSet
  loc_00E1F38F: mov edi, eax
  loc_00E1F391: lea edx, var_28
  loc_00E1F394: push edx
  loc_00E1F395: push edi
  loc_00E1F396: mov ecx, [edi]
  loc_00E1F398: call [ecx+00000054h]
  loc_00E1F39B: test eax, eax
  loc_00E1F39D: fnclex
  loc_00E1F39F: jge 00E1F3B0h
  loc_00E1F3A1: push 00000054h
  loc_00E1F3A3: push 006DCBE8h
  loc_00E1F3A8: push edi
  loc_00E1F3A9: push eax
  loc_00E1F3AA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F3B0: lea ebx, var_2C
  loc_00E1F3B3: mov eax, var_28
  loc_00E1F3B6: push ebx
  loc_00E1F3B7: mov edx, 00000002h
  loc_00E1F3BC: sub esp, 00000010h
  loc_00E1F3BF: mov var_80, edx
  loc_00E1F3C2: mov ebx, esp
  loc_00E1F3C4: mov ecx, 00000007h
  loc_00E1F3C9: mov var_78, ecx
  loc_00E1F3CC: mov edi, [eax]
  loc_00E1F3CE: mov [ebx], edx
  loc_00E1F3D0: mov edx, var_7C
  loc_00E1F3D3: push eax
  loc_00E1F3D4: mov var_BC, eax
  loc_00E1F3DA: mov [ebx+00000004h], edx
  loc_00E1F3DD: mov [ebx+00000008h], ecx
  loc_00E1F3E0: mov ecx, var_74
  loc_00E1F3E3: mov [ebx+0000000Ch], ecx
  loc_00E1F3E6: call [edi+00000028h]
  loc_00E1F3E9: test eax, eax
  loc_00E1F3EB: fnclex
  loc_00E1F3ED: jge 00E1F404h
  loc_00E1F3EF: mov edx, var_BC
  loc_00E1F3F5: push 00000028h
  loc_00E1F3F7: push 006E09E8h
  loc_00E1F3FC: push edx
  loc_00E1F3FD: push eax
  loc_00E1F3FE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F404: mov eax, var_2C
  loc_00E1F407: lea edx, var_50
  loc_00E1F40A: push edx
  loc_00E1F40B: push eax
  loc_00E1F40C: mov ecx, [eax]
  loc_00E1F40E: mov edi, eax
  loc_00E1F410: call [ecx+00000034h]
  loc_00E1F413: test eax, eax
  loc_00E1F415: fnclex
  loc_00E1F417: jge 00E1F428h
  loc_00E1F419: push 00000034h
  loc_00E1F41B: push 006E09F8h
  loc_00E1F420: push edi
  loc_00E1F421: push eax
  loc_00E1F422: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F428: mov edi, var_CC
  loc_00E1F42E: lea eax, var_50
  loc_00E1F431: push eax
  loc_00E1F432: mov ebx, [edi]
  loc_00E1F434: call [00401034h] ; __vbaStrVarMove
  loc_00E1F43A: mov edx, eax
  loc_00E1F43C: lea ecx, var_18
  loc_00E1F43F: call [00401228h] ; __vbaStrMove
  loc_00E1F445: push eax
  loc_00E1F446: push edi
  loc_00E1F447: call [ebx+00000054h]
  loc_00E1F44A: test eax, eax
  loc_00E1F44C: fnclex
  loc_00E1F44E: jge 00E1F463h
  loc_00E1F450: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F456: push 00000054h
  loc_00E1F458: push 006E0410h
  loc_00E1F45D: push edi
  loc_00E1F45E: push eax
  loc_00E1F45F: call ebx
  loc_00E1F461: jmp 00E1F469h
  loc_00E1F463: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F469: lea ecx, var_18
  loc_00E1F46C: call [00401258h] ; __vbaFreeStr
  loc_00E1F472: lea ecx, var_30
  loc_00E1F475: lea edx, var_2C
  loc_00E1F478: push ecx
  loc_00E1F479: lea eax, var_28
  loc_00E1F47C: push edx
  loc_00E1F47D: lea ecx, var_24
  loc_00E1F480: push eax
  loc_00E1F481: lea edx, var_20
  loc_00E1F484: push ecx
  loc_00E1F485: push edx
  loc_00E1F486: push 00000005h
  loc_00E1F488: call [00401048h] ; __vbaFreeObjList
  loc_00E1F48E: lea eax, var_50
  loc_00E1F491: lea ecx, var_40
  loc_00E1F494: push eax
  loc_00E1F495: push ecx
  loc_00E1F496: push 00000002h
  loc_00E1F498: call [00401038h] ; __vbaFreeVarList
  loc_00E1F49E: mov eax, [00E3D060h]
  loc_00E1F4A3: add esp, 00000024h
  loc_00E1F4A6: test eax, eax
  loc_00E1F4A8: jnz 00E1F4BFh
  loc_00E1F4AA: push 00E3D060h
  loc_00E1F4AF: push 006D87C0h
  loc_00E1F4B4: call [004011A0h] ; __vbaNew2
  loc_00E1F4BA: mov eax, [00E3D060h]
  loc_00E1F4BF: mov edx, [eax]
  loc_00E1F4C1: push eax
  loc_00E1F4C2: call [edx+000004E4h]
  loc_00E1F4C8: push eax
  loc_00E1F4C9: lea eax, var_20
  loc_00E1F4CC: push eax
  loc_00E1F4CD: call __vbaObjSet
  loc_00E1F4CF: mov edi, eax
  loc_00E1F4D1: lea edx, var_18
  loc_00E1F4D4: push edx
  loc_00E1F4D5: push edi
  loc_00E1F4D6: mov ecx, [edi]
  loc_00E1F4D8: call [ecx+00000050h]
  loc_00E1F4DB: test eax, eax
  loc_00E1F4DD: fnclex
  loc_00E1F4DF: jge 00E1F4ECh
  loc_00E1F4E1: push 00000050h
  loc_00E1F4E3: push 006E0410h
  loc_00E1F4E8: push edi
  loc_00E1F4E9: push eax
  loc_00E1F4EA: call ebx
  loc_00E1F4EC: mov eax, [00E3D060h]
  loc_00E1F4F1: test eax, eax
  loc_00E1F4F3: jnz 00E1F50Ah
  loc_00E1F4F5: push 00E3D060h
  loc_00E1F4FA: push 006D87C0h
  loc_00E1F4FF: call [004011A0h] ; __vbaNew2
  loc_00E1F505: mov eax, [00E3D060h]
  loc_00E1F50A: mov ecx, [eax]
  loc_00E1F50C: push eax
  loc_00E1F50D: call [ecx+00000518h]
  loc_00E1F513: lea edx, var_24
  loc_00E1F516: push eax
  loc_00E1F517: push edx
  loc_00E1F518: call __vbaObjSet
  loc_00E1F51A: mov edi, eax
  loc_00E1F51C: lea ecx, var_1C
  loc_00E1F51F: push ecx
  loc_00E1F520: push edi
  loc_00E1F521: mov eax, [edi]
  loc_00E1F523: call [eax+00000050h]
  loc_00E1F526: test eax, eax
  loc_00E1F528: fnclex
  loc_00E1F52A: jge 00E1F537h
  loc_00E1F52C: push 00000050h
  loc_00E1F52E: push 006E0410h
  loc_00E1F533: push edi
  loc_00E1F534: push eax
  loc_00E1F535: call ebx
  loc_00E1F537: mov edx, var_1C
  loc_00E1F53A: mov ebx, [00401104h] ; __vbaStrCmp
  loc_00E1F540: push edx
  loc_00E1F541: push 006E0670h ; "Rekayasa Perangkat Lunak"
  loc_00E1F546: call ebx
  loc_00E1F548: mov edi, eax
  loc_00E1F54A: mov eax, var_18
  loc_00E1F54D: neg edi
  loc_00E1F54F: sbb edi, edi
  loc_00E1F551: push eax
  loc_00E1F552: inc edi
  loc_00E1F553: push 006E0B24h ; "Laki - laki"
  loc_00E1F558: neg edi
  loc_00E1F55A: call ebx
  loc_00E1F55C: neg eax
  loc_00E1F55E: sbb eax, eax
  loc_00E1F560: lea ecx, var_1C
  loc_00E1F563: inc eax
  loc_00E1F564: lea edx, var_18
  loc_00E1F567: push ecx
  loc_00E1F568: push edx
  loc_00E1F569: neg eax
  loc_00E1F56B: push 00000002h
  loc_00E1F56D: and edi, eax
  loc_00E1F56F: call [004011BCh] ; __vbaFreeStrList
  loc_00E1F575: lea eax, var_24
  loc_00E1F578: lea ecx, var_20
  loc_00E1F57B: push eax
  loc_00E1F57C: push ecx
  loc_00E1F57D: push 00000002h
  loc_00E1F57F: call [00401048h] ; __vbaFreeObjList
  loc_00E1F585: mov eax, [00E3D060h]
  loc_00E1F58A: add esp, 00000018h
  loc_00E1F58D: test di, di
  loc_00E1F590: jz 00E200B0h
  loc_00E1F596: test eax, eax
  loc_00E1F598: jnz 00E1F5AFh
  loc_00E1F59A: push 00E3D060h
  loc_00E1F59F: push 006D87C0h
  loc_00E1F5A4: call [004011A0h] ; __vbaNew2
  loc_00E1F5AA: mov eax, [00E3D060h]
  loc_00E1F5AF: mov edx, [eax]
  loc_00E1F5B1: push eax
  loc_00E1F5B2: call [edx+000004E0h]
  loc_00E1F5B8: push eax
  loc_00E1F5B9: lea eax, var_20
  loc_00E1F5BC: push eax
  loc_00E1F5BD: call __vbaObjSet
  loc_00E1F5BF: mov edi, eax
  loc_00E1F5C1: lea edx, var_18
  loc_00E1F5C4: push edx
  loc_00E1F5C5: push edi
  loc_00E1F5C6: mov ecx, [edi]
  loc_00E1F5C8: call [ecx+00000050h]
  loc_00E1F5CB: test eax, eax
  loc_00E1F5CD: fnclex
  loc_00E1F5CF: jge 00E1F5E0h
  loc_00E1F5D1: push 00000050h
  loc_00E1F5D3: push 006E0410h
  loc_00E1F5D8: push edi
  loc_00E1F5D9: push eax
  loc_00E1F5DA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F5E0: mov eax, var_18
  loc_00E1F5E3: push eax
  loc_00E1F5E4: push 006E05FCh ; "Katholik"
  loc_00E1F5E9: call ebx
  loc_00E1F5EB: mov edi, eax
  loc_00E1F5ED: lea ecx, var_18
  loc_00E1F5F0: neg edi
  loc_00E1F5F2: sbb edi, edi
  loc_00E1F5F4: inc edi
  loc_00E1F5F5: neg edi
  loc_00E1F5F7: call [00401258h] ; __vbaFreeStr
  loc_00E1F5FD: lea ecx, var_20
  loc_00E1F600: call [00401254h] ; __vbaFreeObj
  loc_00E1F606: mov eax, [00E3D060h]
  loc_00E1F60B: test di, di
  loc_00E1F60E: jz 00E1FA8Fh
  loc_00E1F614: test eax, eax
  loc_00E1F616: jnz 00E1F62Dh
  loc_00E1F618: push 00E3D060h
  loc_00E1F61D: push 006D87C0h
  loc_00E1F622: call [004011A0h] ; __vbaNew2
  loc_00E1F628: mov eax, [00E3D060h]
  loc_00E1F62D: mov ecx, [eax]
  loc_00E1F62F: push eax
  loc_00E1F630: call [ecx+00000474h]
  loc_00E1F636: lea edx, var_20
  loc_00E1F639: push eax
  loc_00E1F63A: push edx
  loc_00E1F63B: call __vbaObjSet
  loc_00E1F63D: mov edi, eax
  loc_00E1F63F: push 000927C0h
  loc_00E1F644: mov ebx, [edi]
  loc_00E1F646: call [00401018h] ; __vbaStrI4
  loc_00E1F64C: mov edx, eax
  loc_00E1F64E: lea ecx, var_18
  loc_00E1F651: call [00401228h] ; __vbaStrMove
  loc_00E1F657: push eax
  loc_00E1F658: push edi
  loc_00E1F659: call [ebx+00000054h]
  loc_00E1F65C: test eax, eax
  loc_00E1F65E: fnclex
  loc_00E1F660: jge 00E1F671h
  loc_00E1F662: push 00000054h
  loc_00E1F664: push 006E0410h
  loc_00E1F669: push edi
  loc_00E1F66A: push eax
  loc_00E1F66B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F671: lea ecx, var_18
  loc_00E1F674: call [00401258h] ; __vbaFreeStr
  loc_00E1F67A: lea ecx, var_20
  loc_00E1F67D: call [00401254h] ; __vbaFreeObj
  loc_00E1F683: mov eax, [00E3D060h]
  loc_00E1F688: test eax, eax
  loc_00E1F68A: jnz 00E1F6A1h
  loc_00E1F68C: push 00E3D060h
  loc_00E1F691: push 006D87C0h
  loc_00E1F696: call [004011A0h] ; __vbaNew2
  loc_00E1F69C: mov eax, [00E3D060h]
  loc_00E1F6A1: mov ecx, [eax]
  loc_00E1F6A3: push eax
  loc_00E1F6A4: call [ecx+00000470h]
  loc_00E1F6AA: lea edx, var_20
  loc_00E1F6AD: push eax
  loc_00E1F6AE: push edx
  loc_00E1F6AF: call __vbaObjSet
  loc_00E1F6B1: mov edi, eax
  loc_00E1F6B3: push 0003A980h
  loc_00E1F6B8: mov ebx, [edi]
  loc_00E1F6BA: call [00401018h] ; __vbaStrI4
  loc_00E1F6C0: mov edx, eax
  loc_00E1F6C2: lea ecx, var_18
  loc_00E1F6C5: call [00401228h] ; __vbaStrMove
  loc_00E1F6CB: push eax
  loc_00E1F6CC: push edi
  loc_00E1F6CD: call [ebx+00000054h]
  loc_00E1F6D0: test eax, eax
  loc_00E1F6D2: fnclex
  loc_00E1F6D4: jge 00E1F6E5h
  loc_00E1F6D6: push 00000054h
  loc_00E1F6D8: push 006E0410h
  loc_00E1F6DD: push edi
  loc_00E1F6DE: push eax
  loc_00E1F6DF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F6E5: lea ecx, var_18
  loc_00E1F6E8: call [00401258h] ; __vbaFreeStr
  loc_00E1F6EE: lea ecx, var_20
  loc_00E1F6F1: call [00401254h] ; __vbaFreeObj
  loc_00E1F6F7: mov eax, [00E3D060h]
  loc_00E1F6FC: test eax, eax
  loc_00E1F6FE: jnz 00E1F715h
  loc_00E1F700: push 00E3D060h
  loc_00E1F705: push 006D87C0h
  loc_00E1F70A: call [004011A0h] ; __vbaNew2
  loc_00E1F710: mov eax, [00E3D060h]
  loc_00E1F715: mov ecx, [eax]
  loc_00E1F717: push eax
  loc_00E1F718: call [ecx+0000046Ch]
  loc_00E1F71E: lea edx, var_20
  loc_00E1F721: push eax
  loc_00E1F722: push edx
  loc_00E1F723: call __vbaObjSet
  loc_00E1F725: mov edi, eax
  loc_00E1F727: push 00077A10h
  loc_00E1F72C: mov ebx, [edi]
  loc_00E1F72E: call [00401018h] ; __vbaStrI4
  loc_00E1F734: mov edx, eax
  loc_00E1F736: lea ecx, var_18
  loc_00E1F739: call [00401228h] ; __vbaStrMove
  loc_00E1F73F: push eax
  loc_00E1F740: push edi
  loc_00E1F741: call [ebx+00000054h]
  loc_00E1F744: test eax, eax
  loc_00E1F746: fnclex
  loc_00E1F748: jge 00E1F759h
  loc_00E1F74A: push 00000054h
  loc_00E1F74C: push 006E0410h
  loc_00E1F751: push edi
  loc_00E1F752: push eax
  loc_00E1F753: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F759: lea ecx, var_18
  loc_00E1F75C: call [00401258h] ; __vbaFreeStr
  loc_00E1F762: lea ecx, var_20
  loc_00E1F765: call [00401254h] ; __vbaFreeObj
  loc_00E1F76B: mov eax, [00E3D060h]
  loc_00E1F770: test eax, eax
  loc_00E1F772: jnz 00E1F789h
  loc_00E1F774: push 00E3D060h
  loc_00E1F779: push 006D87C0h
  loc_00E1F77E: call [004011A0h] ; __vbaNew2
  loc_00E1F784: mov eax, [00E3D060h]
  loc_00E1F789: mov ecx, [eax]
  loc_00E1F78B: push eax
  loc_00E1F78C: call [ecx+00000454h]
  loc_00E1F792: lea edx, var_20
  loc_00E1F795: push eax
  loc_00E1F796: push edx
  loc_00E1F797: call __vbaObjSet
  loc_00E1F799: mov edi, eax
  loc_00E1F79B: push 000124F8h
  loc_00E1F7A0: mov ebx, [edi]
  loc_00E1F7A2: call [00401018h] ; __vbaStrI4
  loc_00E1F7A8: mov edx, eax
  loc_00E1F7AA: lea ecx, var_18
  loc_00E1F7AD: call [00401228h] ; __vbaStrMove
  loc_00E1F7B3: push eax
  loc_00E1F7B4: push edi
  loc_00E1F7B5: call [ebx+00000054h]
  loc_00E1F7B8: test eax, eax
  loc_00E1F7BA: fnclex
  loc_00E1F7BC: jge 00E1F7CDh
  loc_00E1F7BE: push 00000054h
  loc_00E1F7C0: push 006E0410h
  loc_00E1F7C5: push edi
  loc_00E1F7C6: push eax
  loc_00E1F7C7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F7CD: lea ecx, var_18
  loc_00E1F7D0: call [00401258h] ; __vbaFreeStr
  loc_00E1F7D6: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E1F7DC: lea ecx, var_20
  loc_00E1F7DF: call ebx
  loc_00E1F7E1: mov eax, [00E3D060h]
  loc_00E1F7E6: test eax, eax
  loc_00E1F7E8: jnz 00E1F7FFh
  loc_00E1F7EA: push 00E3D060h
  loc_00E1F7EF: push 006D87C0h
  loc_00E1F7F4: call [004011A0h] ; __vbaNew2
  loc_00E1F7FA: mov eax, [00E3D060h]
  loc_00E1F7FF: mov ecx, [eax]
  loc_00E1F801: push eax
  loc_00E1F802: call [ecx+0000042Ch]
  loc_00E1F808: lea edx, var_20
  loc_00E1F80B: push eax
  loc_00E1F80C: push edx
  loc_00E1F80D: call __vbaObjSet
  loc_00E1F80F: mov edi, eax
  loc_00E1F811: push 00000000h
  loc_00E1F813: push edi
  loc_00E1F814: mov eax, [edi]
  loc_00E1F816: call [eax+0000009Ch]
  loc_00E1F81C: test eax, eax
  loc_00E1F81E: fnclex
  loc_00E1F820: jge 00E1F834h
  loc_00E1F822: push 0000009Ch
  loc_00E1F827: push 006DCAD0h
  loc_00E1F82C: push edi
  loc_00E1F82D: push eax
  loc_00E1F82E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F834: lea ecx, var_20
  loc_00E1F837: call ebx
  loc_00E1F839: mov eax, [00E3D060h]
  loc_00E1F83E: test eax, eax
  loc_00E1F840: jnz 00E1F857h
  loc_00E1F842: push 00E3D060h
  loc_00E1F847: push 006D87C0h
  loc_00E1F84C: call [004011A0h] ; __vbaNew2
  loc_00E1F852: mov eax, [00E3D060h]
  loc_00E1F857: mov ecx, [eax]
  loc_00E1F859: push eax
  loc_00E1F85A: call [ecx+000003FCh]
  loc_00E1F860: lea edx, var_20
  loc_00E1F863: push eax
  loc_00E1F864: push edx
  loc_00E1F865: call __vbaObjSet
  loc_00E1F867: mov edi, eax
  loc_00E1F869: push 00027100h
  loc_00E1F86E: mov ebx, [edi]
  loc_00E1F870: call [00401018h] ; __vbaStrI4
  loc_00E1F876: mov edx, eax
  loc_00E1F878: lea ecx, var_18
  loc_00E1F87B: call [00401228h] ; __vbaStrMove
  loc_00E1F881: push eax
  loc_00E1F882: push edi
  loc_00E1F883: call [ebx+00000054h]
  loc_00E1F886: test eax, eax
  loc_00E1F888: fnclex
  loc_00E1F88A: jge 00E1F89Bh
  loc_00E1F88C: push 00000054h
  loc_00E1F88E: push 006E0410h
  loc_00E1F893: push edi
  loc_00E1F894: push eax
  loc_00E1F895: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F89B: lea ecx, var_18
  loc_00E1F89E: call [00401258h] ; __vbaFreeStr
  loc_00E1F8A4: lea ecx, var_20
  loc_00E1F8A7: call [00401254h] ; __vbaFreeObj
  loc_00E1F8AD: mov eax, [00E3D060h]
  loc_00E1F8B2: test eax, eax
  loc_00E1F8B4: jnz 00E1F8CBh
  loc_00E1F8B6: push 00E3D060h
  loc_00E1F8BB: push 006D87C0h
  loc_00E1F8C0: call [004011A0h] ; __vbaNew2
  loc_00E1F8C6: mov eax, [00E3D060h]
  loc_00E1F8CB: mov ecx, [eax]
  loc_00E1F8CD: push eax
  loc_00E1F8CE: call [ecx+000003ECh]
  loc_00E1F8D4: lea edx, var_20
  loc_00E1F8D7: push eax
  loc_00E1F8D8: push edx
  loc_00E1F8D9: call __vbaObjSet
  loc_00E1F8DB: mov edi, eax
  loc_00E1F8DD: push 00035B60h
  loc_00E1F8E2: mov ebx, [edi]
  loc_00E1F8E4: call [00401018h] ; __vbaStrI4
  loc_00E1F8EA: mov edx, eax
  loc_00E1F8EC: lea ecx, var_18
  loc_00E1F8EF: call [00401228h] ; __vbaStrMove
  loc_00E1F8F5: push eax
  loc_00E1F8F6: push edi
  loc_00E1F8F7: call [ebx+00000054h]
  loc_00E1F8FA: test eax, eax
  loc_00E1F8FC: fnclex
  loc_00E1F8FE: jge 00E1F90Fh
  loc_00E1F900: push 00000054h
  loc_00E1F902: push 006E0410h
  loc_00E1F907: push edi
  loc_00E1F908: push eax
  loc_00E1F909: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F90F: lea ecx, var_18
  loc_00E1F912: call [00401258h] ; __vbaFreeStr
  loc_00E1F918: lea ecx, var_20
  loc_00E1F91B: call [00401254h] ; __vbaFreeObj
  loc_00E1F921: mov eax, [00E3D060h]
  loc_00E1F926: test eax, eax
  loc_00E1F928: jnz 00E1F93Fh
  loc_00E1F92A: push 00E3D060h
  loc_00E1F92F: push 006D87C0h
  loc_00E1F934: call [004011A0h] ; __vbaNew2
  loc_00E1F93A: mov eax, [00E3D060h]
  loc_00E1F93F: mov ecx, [eax]
  loc_00E1F941: push eax
  loc_00E1F942: call [ecx+000003DCh]
  loc_00E1F948: lea edx, var_20
  loc_00E1F94B: push eax
  loc_00E1F94C: push edx
  loc_00E1F94D: call __vbaObjSet
  loc_00E1F94F: mov edi, eax
  loc_00E1F951: push 0000FDE8h
  loc_00E1F956: mov ebx, [edi]
  loc_00E1F958: call [00401018h] ; __vbaStrI4
  loc_00E1F95E: mov edx, eax
  loc_00E1F960: lea ecx, var_18
  loc_00E1F963: call [00401228h] ; __vbaStrMove
  loc_00E1F969: push eax
  loc_00E1F96A: push edi
  loc_00E1F96B: call [ebx+00000054h]
  loc_00E1F96E: test eax, eax
  loc_00E1F970: fnclex
  loc_00E1F972: jge 00E1F983h
  loc_00E1F974: push 00000054h
  loc_00E1F976: push 006E0410h
  loc_00E1F97B: push edi
  loc_00E1F97C: push eax
  loc_00E1F97D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F983: lea ecx, var_18
  loc_00E1F986: call [00401258h] ; __vbaFreeStr
  loc_00E1F98C: lea ecx, var_20
  loc_00E1F98F: call [00401254h] ; __vbaFreeObj
  loc_00E1F995: mov eax, [00E3D060h]
  loc_00E1F99A: test eax, eax
  loc_00E1F99C: jnz 00E1F9B3h
  loc_00E1F99E: push 00E3D060h
  loc_00E1F9A3: push 006D87C0h
  loc_00E1F9A8: call [004011A0h] ; __vbaNew2
  loc_00E1F9AE: mov eax, [00E3D060h]
  loc_00E1F9B3: mov ecx, [eax]
  loc_00E1F9B5: push eax
  loc_00E1F9B6: call [ecx+000003CCh]
  loc_00E1F9BC: lea edx, var_20
  loc_00E1F9BF: push eax
  loc_00E1F9C0: push edx
  loc_00E1F9C1: call __vbaObjSet
  loc_00E1F9C3: mov edi, eax
  loc_00E1F9C5: push 0001E848h
  loc_00E1F9CA: mov ebx, [edi]
  loc_00E1F9CC: call [00401018h] ; __vbaStrI4
  loc_00E1F9D2: mov edx, eax
  loc_00E1F9D4: lea ecx, var_18
  loc_00E1F9D7: call [00401228h] ; __vbaStrMove
  loc_00E1F9DD: push eax
  loc_00E1F9DE: push edi
  loc_00E1F9DF: call [ebx+00000054h]
  loc_00E1F9E2: test eax, eax
  loc_00E1F9E4: fnclex
  loc_00E1F9E6: jge 00E1F9F7h
  loc_00E1F9E8: push 00000054h
  loc_00E1F9EA: push 006E0410h
  loc_00E1F9EF: push edi
  loc_00E1F9F0: push eax
  loc_00E1F9F1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1F9F7: lea ecx, var_18
  loc_00E1F9FA: call [00401258h] ; __vbaFreeStr
  loc_00E1FA00: lea ecx, var_20
  loc_00E1FA03: call [00401254h] ; __vbaFreeObj
  loc_00E1FA09: mov eax, [00E3D060h]
  loc_00E1FA0E: test eax, eax
  loc_00E1FA10: jnz 00E1FA27h
  loc_00E1FA12: push 00E3D060h
  loc_00E1FA17: push 006D87C0h
  loc_00E1FA1C: call [004011A0h] ; __vbaNew2
  loc_00E1FA22: mov eax, [00E3D060h]
  loc_00E1FA27: mov ecx, [eax]
  loc_00E1FA29: push eax
  loc_00E1FA2A: call [ecx+000003BCh]
  loc_00E1FA30: lea edx, var_20
  loc_00E1FA33: push eax
  loc_00E1FA34: push edx
  loc_00E1FA35: call __vbaObjSet
  loc_00E1FA37: mov edi, eax
  loc_00E1FA39: push 00027100h
  loc_00E1FA3E: mov ebx, [edi]
  loc_00E1FA40: call [00401018h] ; __vbaStrI4
  loc_00E1FA46: mov edx, eax
  loc_00E1FA48: lea ecx, var_18
  loc_00E1FA4B: call [00401228h] ; __vbaStrMove
  loc_00E1FA51: push eax
  loc_00E1FA52: push edi
  loc_00E1FA53: call [ebx+00000054h]
  loc_00E1FA56: test eax, eax
  loc_00E1FA58: fnclex
  loc_00E1FA5A: jge 00E1FA6Bh
  loc_00E1FA5C: push 00000054h
  loc_00E1FA5E: push 006E0410h
  loc_00E1FA63: push edi
  loc_00E1FA64: push eax
  loc_00E1FA65: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FA6B: lea ecx, var_18
  loc_00E1FA6E: call [00401258h] ; __vbaFreeStr
  loc_00E1FA74: lea ecx, var_20
  loc_00E1FA77: call [00401254h] ; __vbaFreeObj
  loc_00E1FA7D: mov eax, [00E3D060h]
  loc_00E1FA82: test eax, eax
  loc_00E1FA84: jnz 00E22660h
  loc_00E1FA8A: jmp 00E2264Bh
  loc_00E1FA8F: test eax, eax
  loc_00E1FA91: jnz 00E1FAA8h
  loc_00E1FA93: push 00E3D060h
  loc_00E1FA98: push 006D87C0h
  loc_00E1FA9D: call [004011A0h] ; __vbaNew2
  loc_00E1FAA3: mov eax, [00E3D060h]
  loc_00E1FAA8: mov ecx, [eax]
  loc_00E1FAAA: push eax
  loc_00E1FAAB: call [ecx+00000474h]
  loc_00E1FAB1: lea edx, var_20
  loc_00E1FAB4: push eax
  loc_00E1FAB5: push edx
  loc_00E1FAB6: call __vbaObjSet
  loc_00E1FAB8: mov edi, eax
  loc_00E1FABA: push 000AAE60h
  loc_00E1FABF: mov ebx, [edi]
  loc_00E1FAC1: call [00401018h] ; __vbaStrI4
  loc_00E1FAC7: mov edx, eax
  loc_00E1FAC9: lea ecx, var_18
  loc_00E1FACC: call [00401228h] ; __vbaStrMove
  loc_00E1FAD2: push eax
  loc_00E1FAD3: push edi
  loc_00E1FAD4: call [ebx+00000054h]
  loc_00E1FAD7: test eax, eax
  loc_00E1FAD9: fnclex
  loc_00E1FADB: jge 00E1FAECh
  loc_00E1FADD: push 00000054h
  loc_00E1FADF: push 006E0410h
  loc_00E1FAE4: push edi
  loc_00E1FAE5: push eax
  loc_00E1FAE6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FAEC: lea ecx, var_18
  loc_00E1FAEF: call [00401258h] ; __vbaFreeStr
  loc_00E1FAF5: lea ecx, var_20
  loc_00E1FAF8: call [00401254h] ; __vbaFreeObj
  loc_00E1FAFE: mov eax, [00E3D060h]
  loc_00E1FB03: test eax, eax
  loc_00E1FB05: jnz 00E1FB1Ch
  loc_00E1FB07: push 00E3D060h
  loc_00E1FB0C: push 006D87C0h
  loc_00E1FB11: call [004011A0h] ; __vbaNew2
  loc_00E1FB17: mov eax, [00E3D060h]
  loc_00E1FB1C: mov ecx, [eax]
  loc_00E1FB1E: push eax
  loc_00E1FB1F: call [ecx+00000470h]
  loc_00E1FB25: lea edx, var_20
  loc_00E1FB28: push eax
  loc_00E1FB29: push edx
  loc_00E1FB2A: call __vbaObjSet
  loc_00E1FB2C: mov edi, eax
  loc_00E1FB2E: push 0003A980h
  loc_00E1FB33: mov ebx, [edi]
  loc_00E1FB35: call [00401018h] ; __vbaStrI4
  loc_00E1FB3B: mov edx, eax
  loc_00E1FB3D: lea ecx, var_18
  loc_00E1FB40: call [00401228h] ; __vbaStrMove
  loc_00E1FB46: push eax
  loc_00E1FB47: push edi
  loc_00E1FB48: call [ebx+00000054h]
  loc_00E1FB4B: test eax, eax
  loc_00E1FB4D: fnclex
  loc_00E1FB4F: jge 00E1FB60h
  loc_00E1FB51: push 00000054h
  loc_00E1FB53: push 006E0410h
  loc_00E1FB58: push edi
  loc_00E1FB59: push eax
  loc_00E1FB5A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FB60: lea ecx, var_18
  loc_00E1FB63: call [00401258h] ; __vbaFreeStr
  loc_00E1FB69: lea ecx, var_20
  loc_00E1FB6C: call [00401254h] ; __vbaFreeObj
  loc_00E1FB72: mov eax, [00E3D060h]
  loc_00E1FB77: test eax, eax
  loc_00E1FB79: jnz 00E1FB90h
  loc_00E1FB7B: push 00E3D060h
  loc_00E1FB80: push 006D87C0h
  loc_00E1FB85: call [004011A0h] ; __vbaNew2
  loc_00E1FB8B: mov eax, [00E3D060h]
  loc_00E1FB90: mov ecx, [eax]
  loc_00E1FB92: push eax
  loc_00E1FB93: call [ecx+0000046Ch]
  loc_00E1FB99: lea edx, var_20
  loc_00E1FB9C: push eax
  loc_00E1FB9D: push edx
  loc_00E1FB9E: call __vbaObjSet
  loc_00E1FBA0: mov edi, eax
  loc_00E1FBA2: push 00077A10h
  loc_00E1FBA7: mov ebx, [edi]
  loc_00E1FBA9: call [00401018h] ; __vbaStrI4
  loc_00E1FBAF: mov edx, eax
  loc_00E1FBB1: lea ecx, var_18
  loc_00E1FBB4: call [00401228h] ; __vbaStrMove
  loc_00E1FBBA: push eax
  loc_00E1FBBB: push edi
  loc_00E1FBBC: call [ebx+00000054h]
  loc_00E1FBBF: test eax, eax
  loc_00E1FBC1: fnclex
  loc_00E1FBC3: jge 00E1FBD4h
  loc_00E1FBC5: push 00000054h
  loc_00E1FBC7: push 006E0410h
  loc_00E1FBCC: push edi
  loc_00E1FBCD: push eax
  loc_00E1FBCE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FBD4: lea ecx, var_18
  loc_00E1FBD7: call [00401258h] ; __vbaFreeStr
  loc_00E1FBDD: lea ecx, var_20
  loc_00E1FBE0: call [00401254h] ; __vbaFreeObj
  loc_00E1FBE6: mov eax, [00E3D060h]
  loc_00E1FBEB: test eax, eax
  loc_00E1FBED: jnz 00E1FC04h
  loc_00E1FBEF: push 00E3D060h
  loc_00E1FBF4: push 006D87C0h
  loc_00E1FBF9: call [004011A0h] ; __vbaNew2
  loc_00E1FBFF: mov eax, [00E3D060h]
  loc_00E1FC04: mov ecx, [eax]
  loc_00E1FC06: push eax
  loc_00E1FC07: call [ecx+00000454h]
  loc_00E1FC0D: lea edx, var_20
  loc_00E1FC10: push eax
  loc_00E1FC11: push edx
  loc_00E1FC12: call __vbaObjSet
  loc_00E1FC14: mov edi, eax
  loc_00E1FC16: push 000124F8h
  loc_00E1FC1B: mov ebx, [edi]
  loc_00E1FC1D: call [00401018h] ; __vbaStrI4
  loc_00E1FC23: mov edx, eax
  loc_00E1FC25: lea ecx, var_18
  loc_00E1FC28: call [00401228h] ; __vbaStrMove
  loc_00E1FC2E: push eax
  loc_00E1FC2F: push edi
  loc_00E1FC30: call [ebx+00000054h]
  loc_00E1FC33: test eax, eax
  loc_00E1FC35: fnclex
  loc_00E1FC37: jge 00E1FC48h
  loc_00E1FC39: push 00000054h
  loc_00E1FC3B: push 006E0410h
  loc_00E1FC40: push edi
  loc_00E1FC41: push eax
  loc_00E1FC42: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FC48: lea ecx, var_18
  loc_00E1FC4B: call [00401258h] ; __vbaFreeStr
  loc_00E1FC51: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E1FC57: lea ecx, var_20
  loc_00E1FC5A: call ebx
  loc_00E1FC5C: mov eax, [00E3D060h]
  loc_00E1FC61: test eax, eax
  loc_00E1FC63: jnz 00E1FC7Ah
  loc_00E1FC65: push 00E3D060h
  loc_00E1FC6A: push 006D87C0h
  loc_00E1FC6F: call [004011A0h] ; __vbaNew2
  loc_00E1FC75: mov eax, [00E3D060h]
  loc_00E1FC7A: mov ecx, [eax]
  loc_00E1FC7C: push eax
  loc_00E1FC7D: call [ecx+0000042Ch]
  loc_00E1FC83: lea edx, var_20
  loc_00E1FC86: push eax
  loc_00E1FC87: push edx
  loc_00E1FC88: call __vbaObjSet
  loc_00E1FC8A: mov edi, eax
  loc_00E1FC8C: push 00000000h
  loc_00E1FC8E: push edi
  loc_00E1FC8F: mov eax, [edi]
  loc_00E1FC91: call [eax+0000009Ch]
  loc_00E1FC97: test eax, eax
  loc_00E1FC99: fnclex
  loc_00E1FC9B: jge 00E1FCAFh
  loc_00E1FC9D: push 0000009Ch
  loc_00E1FCA2: push 006DCAD0h
  loc_00E1FCA7: push edi
  loc_00E1FCA8: push eax
  loc_00E1FCA9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FCAF: lea ecx, var_20
  loc_00E1FCB2: call ebx
  loc_00E1FCB4: mov eax, [00E3D060h]
  loc_00E1FCB9: test eax, eax
  loc_00E1FCBB: jnz 00E1FCD2h
  loc_00E1FCBD: push 00E3D060h
  loc_00E1FCC2: push 006D87C0h
  loc_00E1FCC7: call [004011A0h] ; __vbaNew2
  loc_00E1FCCD: mov eax, [00E3D060h]
  loc_00E1FCD2: mov ecx, [eax]
  loc_00E1FCD4: push eax
  loc_00E1FCD5: call [ecx+000003FCh]
  loc_00E1FCDB: lea edx, var_20
  loc_00E1FCDE: push eax
  loc_00E1FCDF: push edx
  loc_00E1FCE0: call __vbaObjSet
  loc_00E1FCE2: mov edi, eax
  loc_00E1FCE4: push 00027100h
  loc_00E1FCE9: mov ebx, [edi]
  loc_00E1FCEB: call [00401018h] ; __vbaStrI4
  loc_00E1FCF1: mov edx, eax
  loc_00E1FCF3: lea ecx, var_18
  loc_00E1FCF6: call [00401228h] ; __vbaStrMove
  loc_00E1FCFC: push eax
  loc_00E1FCFD: push edi
  loc_00E1FCFE: call [ebx+00000054h]
  loc_00E1FD01: test eax, eax
  loc_00E1FD03: fnclex
  loc_00E1FD05: jge 00E1FD16h
  loc_00E1FD07: push 00000054h
  loc_00E1FD09: push 006E0410h
  loc_00E1FD0E: push edi
  loc_00E1FD0F: push eax
  loc_00E1FD10: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FD16: lea ecx, var_18
  loc_00E1FD19: call [00401258h] ; __vbaFreeStr
  loc_00E1FD1F: lea ecx, var_20
  loc_00E1FD22: call [00401254h] ; __vbaFreeObj
  loc_00E1FD28: mov eax, [00E3D060h]
  loc_00E1FD2D: test eax, eax
  loc_00E1FD2F: jnz 00E1FD46h
  loc_00E1FD31: push 00E3D060h
  loc_00E1FD36: push 006D87C0h
  loc_00E1FD3B: call [004011A0h] ; __vbaNew2
  loc_00E1FD41: mov eax, [00E3D060h]
  loc_00E1FD46: mov ecx, [eax]
  loc_00E1FD48: push eax
  loc_00E1FD49: call [ecx+000003ECh]
  loc_00E1FD4F: lea edx, var_20
  loc_00E1FD52: push eax
  loc_00E1FD53: push edx
  loc_00E1FD54: call __vbaObjSet
  loc_00E1FD56: mov edi, eax
  loc_00E1FD58: push 00035B60h
  loc_00E1FD5D: mov ebx, [edi]
  loc_00E1FD5F: call [00401018h] ; __vbaStrI4
  loc_00E1FD65: mov edx, eax
  loc_00E1FD67: lea ecx, var_18
  loc_00E1FD6A: call [00401228h] ; __vbaStrMove
  loc_00E1FD70: push eax
  loc_00E1FD71: push edi
  loc_00E1FD72: call [ebx+00000054h]
  loc_00E1FD75: test eax, eax
  loc_00E1FD77: fnclex
  loc_00E1FD79: jge 00E1FD8Ah
  loc_00E1FD7B: push 00000054h
  loc_00E1FD7D: push 006E0410h
  loc_00E1FD82: push edi
  loc_00E1FD83: push eax
  loc_00E1FD84: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FD8A: lea ecx, var_18
  loc_00E1FD8D: call [00401258h] ; __vbaFreeStr
  loc_00E1FD93: lea ecx, var_20
  loc_00E1FD96: call [00401254h] ; __vbaFreeObj
  loc_00E1FD9C: mov eax, [00E3D060h]
  loc_00E1FDA1: test eax, eax
  loc_00E1FDA3: jnz 00E1FDBAh
  loc_00E1FDA5: push 00E3D060h
  loc_00E1FDAA: push 006D87C0h
  loc_00E1FDAF: call [004011A0h] ; __vbaNew2
  loc_00E1FDB5: mov eax, [00E3D060h]
  loc_00E1FDBA: mov ecx, [eax]
  loc_00E1FDBC: push eax
  loc_00E1FDBD: call [ecx+000003DCh]
  loc_00E1FDC3: lea edx, var_20
  loc_00E1FDC6: push eax
  loc_00E1FDC7: push edx
  loc_00E1FDC8: call __vbaObjSet
  loc_00E1FDCA: mov edi, eax
  loc_00E1FDCC: push 0000FDE8h
  loc_00E1FDD1: mov ebx, [edi]
  loc_00E1FDD3: call [00401018h] ; __vbaStrI4
  loc_00E1FDD9: mov edx, eax
  loc_00E1FDDB: lea ecx, var_18
  loc_00E1FDDE: call [00401228h] ; __vbaStrMove
  loc_00E1FDE4: push eax
  loc_00E1FDE5: push edi
  loc_00E1FDE6: call [ebx+00000054h]
  loc_00E1FDE9: test eax, eax
  loc_00E1FDEB: fnclex
  loc_00E1FDED: jge 00E1FDFEh
  loc_00E1FDEF: push 00000054h
  loc_00E1FDF1: push 006E0410h
  loc_00E1FDF6: push edi
  loc_00E1FDF7: push eax
  loc_00E1FDF8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FDFE: lea ecx, var_18
  loc_00E1FE01: call [00401258h] ; __vbaFreeStr
  loc_00E1FE07: lea ecx, var_20
  loc_00E1FE0A: call [00401254h] ; __vbaFreeObj
  loc_00E1FE10: mov eax, [00E3D060h]
  loc_00E1FE15: test eax, eax
  loc_00E1FE17: jnz 00E1FE2Eh
  loc_00E1FE19: push 00E3D060h
  loc_00E1FE1E: push 006D87C0h
  loc_00E1FE23: call [004011A0h] ; __vbaNew2
  loc_00E1FE29: mov eax, [00E3D060h]
  loc_00E1FE2E: mov ecx, [eax]
  loc_00E1FE30: push eax
  loc_00E1FE31: call [ecx+000003CCh]
  loc_00E1FE37: lea edx, var_20
  loc_00E1FE3A: push eax
  loc_00E1FE3B: push edx
  loc_00E1FE3C: call __vbaObjSet
  loc_00E1FE3E: mov edi, eax
  loc_00E1FE40: push 0001E848h
  loc_00E1FE45: mov ebx, [edi]
  loc_00E1FE47: call [00401018h] ; __vbaStrI4
  loc_00E1FE4D: mov edx, eax
  loc_00E1FE4F: lea ecx, var_18
  loc_00E1FE52: call [00401228h] ; __vbaStrMove
  loc_00E1FE58: push eax
  loc_00E1FE59: push edi
  loc_00E1FE5A: call [ebx+00000054h]
  loc_00E1FE5D: test eax, eax
  loc_00E1FE5F: fnclex
  loc_00E1FE61: jge 00E1FE72h
  loc_00E1FE63: push 00000054h
  loc_00E1FE65: push 006E0410h
  loc_00E1FE6A: push edi
  loc_00E1FE6B: push eax
  loc_00E1FE6C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FE72: lea ecx, var_18
  loc_00E1FE75: call [00401258h] ; __vbaFreeStr
  loc_00E1FE7B: lea ecx, var_20
  loc_00E1FE7E: call [00401254h] ; __vbaFreeObj
  loc_00E1FE84: mov eax, [00E3D060h]
  loc_00E1FE89: test eax, eax
  loc_00E1FE8B: jnz 00E1FEA2h
  loc_00E1FE8D: push 00E3D060h
  loc_00E1FE92: push 006D87C0h
  loc_00E1FE97: call [004011A0h] ; __vbaNew2
  loc_00E1FE9D: mov eax, [00E3D060h]
  loc_00E1FEA2: mov ecx, [eax]
  loc_00E1FEA4: push eax
  loc_00E1FEA5: call [ecx+000003BCh]
  loc_00E1FEAB: lea edx, var_20
  loc_00E1FEAE: push eax
  loc_00E1FEAF: push edx
  loc_00E1FEB0: call __vbaObjSet
  loc_00E1FEB2: mov edi, eax
  loc_00E1FEB4: push 00027100h
  loc_00E1FEB9: mov ebx, [edi]
  loc_00E1FEBB: call [00401018h] ; __vbaStrI4
  loc_00E1FEC1: mov edx, eax
  loc_00E1FEC3: lea ecx, var_18
  loc_00E1FEC6: call [00401228h] ; __vbaStrMove
  loc_00E1FECC: push eax
  loc_00E1FECD: push edi
  loc_00E1FECE: call [ebx+00000054h]
  loc_00E1FED1: test eax, eax
  loc_00E1FED3: fnclex
  loc_00E1FED5: jge 00E1FEE6h
  loc_00E1FED7: push 00000054h
  loc_00E1FED9: push 006E0410h
  loc_00E1FEDE: push edi
  loc_00E1FEDF: push eax
  loc_00E1FEE0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FEE6: lea ecx, var_18
  loc_00E1FEE9: call [00401258h] ; __vbaFreeStr
  loc_00E1FEEF: lea ecx, var_20
  loc_00E1FEF2: call [00401254h] ; __vbaFreeObj
  loc_00E1FEF8: mov eax, [00E3D060h]
  loc_00E1FEFD: test eax, eax
  loc_00E1FEFF: jnz 00E1FF16h
  loc_00E1FF01: push 00E3D060h
  loc_00E1FF06: push 006D87C0h
  loc_00E1FF0B: call [004011A0h] ; __vbaNew2
  loc_00E1FF11: mov eax, [00E3D060h]
  loc_00E1FF16: mov ecx, [eax]
  loc_00E1FF18: push eax
  loc_00E1FF19: call [ecx+000003A8h]
  loc_00E1FF1F: lea edx, var_20
  loc_00E1FF22: push eax
  loc_00E1FF23: push edx
  loc_00E1FF24: call __vbaObjSet
  loc_00E1FF26: mov edi, eax
  loc_00E1FF28: push 00003A98h
  loc_00E1FF2D: mov ebx, [edi]
  loc_00E1FF2F: call [0040100Ch] ; __vbaStrI2
  loc_00E1FF35: mov edx, eax
  loc_00E1FF37: lea ecx, var_18
  loc_00E1FF3A: call [00401228h] ; __vbaStrMove
  loc_00E1FF40: push eax
  loc_00E1FF41: push edi
  loc_00E1FF42: call [ebx+00000054h]
  loc_00E1FF45: test eax, eax
  loc_00E1FF47: fnclex
  loc_00E1FF49: jge 00E1FF5Ah
  loc_00E1FF4B: push 00000054h
  loc_00E1FF4D: push 006E0410h
  loc_00E1FF52: push edi
  loc_00E1FF53: push eax
  loc_00E1FF54: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FF5A: lea ecx, var_18
  loc_00E1FF5D: call [00401258h] ; __vbaFreeStr
  loc_00E1FF63: lea ecx, var_20
  loc_00E1FF66: call [00401254h] ; __vbaFreeObj
  loc_00E1FF6C: mov eax, [00E3D060h]
  loc_00E1FF71: test eax, eax
  loc_00E1FF73: jnz 00E1FF8Ah
  loc_00E1FF75: push 00E3D060h
  loc_00E1FF7A: push 006D87C0h
  loc_00E1FF7F: call [004011A0h] ; __vbaNew2
  loc_00E1FF85: mov eax, [00E3D060h]
  loc_00E1FF8A: mov ecx, [eax]
  loc_00E1FF8C: push eax
  loc_00E1FF8D: call [ecx+00000398h]
  loc_00E1FF93: lea edx, var_20
  loc_00E1FF96: push eax
  loc_00E1FF97: push edx
  loc_00E1FF98: call __vbaObjSet
  loc_00E1FF9A: mov edi, eax
  loc_00E1FF9C: push 00003A98h
  loc_00E1FFA1: mov ebx, [edi]
  loc_00E1FFA3: call [0040100Ch] ; __vbaStrI2
  loc_00E1FFA9: mov edx, eax
  loc_00E1FFAB: lea ecx, var_18
  loc_00E1FFAE: call [00401228h] ; __vbaStrMove
  loc_00E1FFB4: push eax
  loc_00E1FFB5: push edi
  loc_00E1FFB6: call [ebx+00000054h]
  loc_00E1FFB9: test eax, eax
  loc_00E1FFBB: fnclex
  loc_00E1FFBD: jge 00E1FFCEh
  loc_00E1FFBF: push 00000054h
  loc_00E1FFC1: push 006E0410h
  loc_00E1FFC6: push edi
  loc_00E1FFC7: push eax
  loc_00E1FFC8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1FFCE: lea ecx, var_18
  loc_00E1FFD1: call [00401258h] ; __vbaFreeStr
  loc_00E1FFD7: lea ecx, var_20
  loc_00E1FFDA: call [00401254h] ; __vbaFreeObj
  loc_00E1FFE0: mov eax, [00E3D060h]
  loc_00E1FFE5: test eax, eax
  loc_00E1FFE7: jnz 00E1FFFEh
  loc_00E1FFE9: push 00E3D060h
  loc_00E1FFEE: push 006D87C0h
  loc_00E1FFF3: call [004011A0h] ; __vbaNew2
  loc_00E1FFF9: mov eax, [00E3D060h]
  loc_00E1FFFE: mov ecx, [eax]
  loc_00E20000: push eax
  loc_00E20001: call [ecx+00000388h]
  loc_00E20007: lea edx, var_20
  loc_00E2000A: push eax
  loc_00E2000B: push edx
  loc_00E2000C: call __vbaObjSet
  loc_00E2000E: mov edi, eax
  loc_00E20010: push 00001388h
  loc_00E20015: mov ebx, [edi]
  loc_00E20017: call [0040100Ch] ; __vbaStrI2
  loc_00E2001D: mov edx, eax
  loc_00E2001F: lea ecx, var_18
  loc_00E20022: call [00401228h] ; __vbaStrMove
  loc_00E20028: push eax
  loc_00E20029: push edi
  loc_00E2002A: call [ebx+00000054h]
  loc_00E2002D: test eax, eax
  loc_00E2002F: fnclex
  loc_00E20031: jge 00E20042h
  loc_00E20033: push 00000054h
  loc_00E20035: push 006E0410h
  loc_00E2003A: push edi
  loc_00E2003B: push eax
  loc_00E2003C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20042: lea ecx, var_18
  loc_00E20045: call [00401258h] ; __vbaFreeStr
  loc_00E2004B: lea ecx, var_20
  loc_00E2004E: call [00401254h] ; __vbaFreeObj
  loc_00E20054: mov eax, [00E3D060h]
  loc_00E20059: test eax, eax
  loc_00E2005B: jnz 00E20072h
  loc_00E2005D: push 00E3D060h
  loc_00E20062: push 006D87C0h
  loc_00E20067: call [004011A0h] ; __vbaNew2
  loc_00E2006D: mov eax, [00E3D060h]
  loc_00E20072: mov ecx, [eax]
  loc_00E20074: push eax
  loc_00E20075: call [ecx+00000378h]
  loc_00E2007B: lea edx, var_20
  loc_00E2007E: push eax
  loc_00E2007F: push edx
  loc_00E20080: call __vbaObjSet
  loc_00E20082: mov edi, eax
  loc_00E20084: push 00001388h
  loc_00E20089: mov ebx, [edi]
  loc_00E2008B: call [0040100Ch] ; __vbaStrI2
  loc_00E20091: mov edx, eax
  loc_00E20093: lea ecx, var_18
  loc_00E20096: call [00401228h] ; __vbaStrMove
  loc_00E2009C: push eax
  loc_00E2009D: push edi
  loc_00E2009E: call [ebx+00000054h]
  loc_00E200A1: test eax, eax
  loc_00E200A3: fnclex
  loc_00E200A5: jge 00E22F1Dh
  loc_00E200AB: jmp 00E22F0Eh
  loc_00E200B0: test eax, eax
  loc_00E200B2: jnz 00E200C9h
  loc_00E200B4: push 00E3D060h
  loc_00E200B9: push 006D87C0h
  loc_00E200BE: call [004011A0h] ; __vbaNew2
  loc_00E200C4: mov eax, [00E3D060h]
  loc_00E200C9: mov ecx, [eax]
  loc_00E200CB: push eax
  loc_00E200CC: call [ecx+000004E4h]
  loc_00E200D2: lea edx, var_20
  loc_00E200D5: push eax
  loc_00E200D6: push edx
  loc_00E200D7: call __vbaObjSet
  loc_00E200D9: mov edi, eax
  loc_00E200DB: lea ecx, var_18
  loc_00E200DE: push ecx
  loc_00E200DF: push edi
  loc_00E200E0: mov eax, [edi]
  loc_00E200E2: call [eax+00000050h]
  loc_00E200E5: test eax, eax
  loc_00E200E7: fnclex
  loc_00E200E9: jge 00E200FAh
  loc_00E200EB: push 00000050h
  loc_00E200ED: push 006E0410h
  loc_00E200F2: push edi
  loc_00E200F3: push eax
  loc_00E200F4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E200FA: mov eax, [00E3D060h]
  loc_00E200FF: test eax, eax
  loc_00E20101: jnz 00E20118h
  loc_00E20103: push 00E3D060h
  loc_00E20108: push 006D87C0h
  loc_00E2010D: call [004011A0h] ; __vbaNew2
  loc_00E20113: mov eax, [00E3D060h]
  loc_00E20118: mov edx, [eax]
  loc_00E2011A: push eax
  loc_00E2011B: call [edx+00000518h]
  loc_00E20121: push eax
  loc_00E20122: lea eax, var_24
  loc_00E20125: push eax
  loc_00E20126: call __vbaObjSet
  loc_00E20128: mov edi, eax
  loc_00E2012A: lea edx, var_1C
  loc_00E2012D: push edx
  loc_00E2012E: push edi
  loc_00E2012F: mov ecx, [edi]
  loc_00E20131: call [ecx+00000050h]
  loc_00E20134: test eax, eax
  loc_00E20136: fnclex
  loc_00E20138: jge 00E20149h
  loc_00E2013A: push 00000050h
  loc_00E2013C: push 006E0410h
  loc_00E20141: push edi
  loc_00E20142: push eax
  loc_00E20143: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20149: mov eax, var_1C
  loc_00E2014C: push eax
  loc_00E2014D: push 006E0670h ; "Rekayasa Perangkat Lunak"
  loc_00E20152: call ebx
  loc_00E20154: mov ecx, var_18
  loc_00E20157: mov edi, eax
  loc_00E20159: neg edi
  loc_00E2015B: sbb edi, edi
  loc_00E2015D: push ecx
  loc_00E2015E: inc edi
  loc_00E2015F: push 006E0B40h ; "Perempuan"
  loc_00E20164: neg edi
  loc_00E20166: call ebx
  loc_00E20168: neg eax
  loc_00E2016A: sbb eax, eax
  loc_00E2016C: lea edx, var_1C
  loc_00E2016F: inc eax
  loc_00E20170: push edx
  loc_00E20171: neg eax
  loc_00E20173: and edi, eax
  loc_00E20175: lea eax, var_18
  loc_00E20178: push eax
  loc_00E20179: push 00000002h
  loc_00E2017B: call [004011BCh] ; __vbaFreeStrList
  loc_00E20181: lea ecx, var_24
  loc_00E20184: lea edx, var_20
  loc_00E20187: push ecx
  loc_00E20188: push edx
  loc_00E20189: push 00000002h
  loc_00E2018B: call [00401048h] ; __vbaFreeObjList
  loc_00E20191: mov eax, [00E3D060h]
  loc_00E20196: add esp, 00000018h
  loc_00E20199: test di, di
  loc_00E2019C: jz 00E20CBCh
  loc_00E201A2: test eax, eax
  loc_00E201A4: jnz 00E201BBh
  loc_00E201A6: push 00E3D060h
  loc_00E201AB: push 006D87C0h
  loc_00E201B0: call [004011A0h] ; __vbaNew2
  loc_00E201B6: mov eax, [00E3D060h]
  loc_00E201BB: mov ecx, [eax]
  loc_00E201BD: push eax
  loc_00E201BE: call [ecx+000004E0h]
  loc_00E201C4: lea edx, var_20
  loc_00E201C7: push eax
  loc_00E201C8: push edx
  loc_00E201C9: call __vbaObjSet
  loc_00E201CB: mov edi, eax
  loc_00E201CD: lea ecx, var_18
  loc_00E201D0: push ecx
  loc_00E201D1: push edi
  loc_00E201D2: mov eax, [edi]
  loc_00E201D4: call [eax+00000050h]
  loc_00E201D7: test eax, eax
  loc_00E201D9: fnclex
  loc_00E201DB: jge 00E201ECh
  loc_00E201DD: push 00000050h
  loc_00E201DF: push 006E0410h
  loc_00E201E4: push edi
  loc_00E201E5: push eax
  loc_00E201E6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E201EC: mov edx, var_18
  loc_00E201EF: push edx
  loc_00E201F0: push 006E05FCh ; "Katholik"
  loc_00E201F5: call ebx
  loc_00E201F7: mov edi, eax
  loc_00E201F9: lea ecx, var_18
  loc_00E201FC: neg edi
  loc_00E201FE: sbb edi, edi
  loc_00E20200: inc edi
  loc_00E20201: neg edi
  loc_00E20203: call [00401258h] ; __vbaFreeStr
  loc_00E20209: lea ecx, var_20
  loc_00E2020C: call [00401254h] ; __vbaFreeObj
  loc_00E20212: mov eax, [00E3D060h]
  loc_00E20217: test di, di
  loc_00E2021A: jz 00E2070Fh
  loc_00E20220: test eax, eax
  loc_00E20222: jnz 00E20239h
  loc_00E20224: push 00E3D060h
  loc_00E20229: push 006D87C0h
  loc_00E2022E: call [004011A0h] ; __vbaNew2
  loc_00E20234: mov eax, [00E3D060h]
  loc_00E20239: mov ecx, [eax]
  loc_00E2023B: push eax
  loc_00E2023C: call [ecx+00000474h]
  loc_00E20242: lea edx, var_20
  loc_00E20245: push eax
  loc_00E20246: push edx
  loc_00E20247: call __vbaObjSet
  loc_00E20249: mov edi, eax
  loc_00E2024B: push 000927C0h
  loc_00E20250: mov ebx, [edi]
  loc_00E20252: call [00401018h] ; __vbaStrI4
  loc_00E20258: mov edx, eax
  loc_00E2025A: lea ecx, var_18
  loc_00E2025D: call [00401228h] ; __vbaStrMove
  loc_00E20263: push eax
  loc_00E20264: push edi
  loc_00E20265: call [ebx+00000054h]
  loc_00E20268: test eax, eax
  loc_00E2026A: fnclex
  loc_00E2026C: jge 00E2027Dh
  loc_00E2026E: push 00000054h
  loc_00E20270: push 006E0410h
  loc_00E20275: push edi
  loc_00E20276: push eax
  loc_00E20277: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2027D: lea ecx, var_18
  loc_00E20280: call [00401258h] ; __vbaFreeStr
  loc_00E20286: lea ecx, var_20
  loc_00E20289: call [00401254h] ; __vbaFreeObj
  loc_00E2028F: mov eax, [00E3D060h]
  loc_00E20294: test eax, eax
  loc_00E20296: jnz 00E202ADh
  loc_00E20298: push 00E3D060h
  loc_00E2029D: push 006D87C0h
  loc_00E202A2: call [004011A0h] ; __vbaNew2
  loc_00E202A8: mov eax, [00E3D060h]
  loc_00E202AD: mov ecx, [eax]
  loc_00E202AF: push eax
  loc_00E202B0: call [ecx+00000470h]
  loc_00E202B6: lea edx, var_20
  loc_00E202B9: push eax
  loc_00E202BA: push edx
  loc_00E202BB: call __vbaObjSet
  loc_00E202BD: mov edi, eax
  loc_00E202BF: push 0003A980h
  loc_00E202C4: mov ebx, [edi]
  loc_00E202C6: call [00401018h] ; __vbaStrI4
  loc_00E202CC: mov edx, eax
  loc_00E202CE: lea ecx, var_18
  loc_00E202D1: call [00401228h] ; __vbaStrMove
  loc_00E202D7: push eax
  loc_00E202D8: push edi
  loc_00E202D9: call [ebx+00000054h]
  loc_00E202DC: test eax, eax
  loc_00E202DE: fnclex
  loc_00E202E0: jge 00E202F1h
  loc_00E202E2: push 00000054h
  loc_00E202E4: push 006E0410h
  loc_00E202E9: push edi
  loc_00E202EA: push eax
  loc_00E202EB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E202F1: lea ecx, var_18
  loc_00E202F4: call [00401258h] ; __vbaFreeStr
  loc_00E202FA: lea ecx, var_20
  loc_00E202FD: call [00401254h] ; __vbaFreeObj
  loc_00E20303: mov eax, [00E3D060h]
  loc_00E20308: test eax, eax
  loc_00E2030A: jnz 00E20321h
  loc_00E2030C: push 00E3D060h
  loc_00E20311: push 006D87C0h
  loc_00E20316: call [004011A0h] ; __vbaNew2
  loc_00E2031C: mov eax, [00E3D060h]
  loc_00E20321: mov ecx, [eax]
  loc_00E20323: push eax
  loc_00E20324: call [ecx+0000046Ch]
  loc_00E2032A: lea edx, var_20
  loc_00E2032D: push eax
  loc_00E2032E: push edx
  loc_00E2032F: call __vbaObjSet
  loc_00E20331: mov edi, eax
  loc_00E20333: push 00077A10h
  loc_00E20338: mov ebx, [edi]
  loc_00E2033A: call [00401018h] ; __vbaStrI4
  loc_00E20340: mov edx, eax
  loc_00E20342: lea ecx, var_18
  loc_00E20345: call [00401228h] ; __vbaStrMove
  loc_00E2034B: push eax
  loc_00E2034C: push edi
  loc_00E2034D: call [ebx+00000054h]
  loc_00E20350: test eax, eax
  loc_00E20352: fnclex
  loc_00E20354: jge 00E20365h
  loc_00E20356: push 00000054h
  loc_00E20358: push 006E0410h
  loc_00E2035D: push edi
  loc_00E2035E: push eax
  loc_00E2035F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20365: lea ecx, var_18
  loc_00E20368: call [00401258h] ; __vbaFreeStr
  loc_00E2036E: lea ecx, var_20
  loc_00E20371: call [00401254h] ; __vbaFreeObj
  loc_00E20377: mov eax, [00E3D060h]
  loc_00E2037C: test eax, eax
  loc_00E2037E: jnz 00E20395h
  loc_00E20380: push 00E3D060h
  loc_00E20385: push 006D87C0h
  loc_00E2038A: call [004011A0h] ; __vbaNew2
  loc_00E20390: mov eax, [00E3D060h]
  loc_00E20395: mov ecx, [eax]
  loc_00E20397: push eax
  loc_00E20398: call [ecx+00000454h]
  loc_00E2039E: lea edx, var_20
  loc_00E203A1: push eax
  loc_00E203A2: push edx
  loc_00E203A3: call __vbaObjSet
  loc_00E203A5: mov edi, eax
  loc_00E203A7: push 000124F8h
  loc_00E203AC: mov ebx, [edi]
  loc_00E203AE: call [00401018h] ; __vbaStrI4
  loc_00E203B4: mov edx, eax
  loc_00E203B6: lea ecx, var_18
  loc_00E203B9: call [00401228h] ; __vbaStrMove
  loc_00E203BF: push eax
  loc_00E203C0: push edi
  loc_00E203C1: call [ebx+00000054h]
  loc_00E203C4: test eax, eax
  loc_00E203C6: fnclex
  loc_00E203C8: jge 00E203D9h
  loc_00E203CA: push 00000054h
  loc_00E203CC: push 006E0410h
  loc_00E203D1: push edi
  loc_00E203D2: push eax
  loc_00E203D3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E203D9: lea ecx, var_18
  loc_00E203DC: call [00401258h] ; __vbaFreeStr
  loc_00E203E2: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E203E8: lea ecx, var_20
  loc_00E203EB: call ebx
  loc_00E203ED: mov eax, [00E3D060h]
  loc_00E203F2: test eax, eax
  loc_00E203F4: jnz 00E2040Bh
  loc_00E203F6: push 00E3D060h
  loc_00E203FB: push 006D87C0h
  loc_00E20400: call [004011A0h] ; __vbaNew2
  loc_00E20406: mov eax, [00E3D060h]
  loc_00E2040B: mov ecx, [eax]
  loc_00E2040D: push eax
  loc_00E2040E: call [ecx+0000042Ch]
  loc_00E20414: lea edx, var_20
  loc_00E20417: push eax
  loc_00E20418: push edx
  loc_00E20419: call __vbaObjSet
  loc_00E2041B: mov edi, eax
  loc_00E2041D: push 00000000h
  loc_00E2041F: push edi
  loc_00E20420: mov eax, [edi]
  loc_00E20422: call [eax+0000009Ch]
  loc_00E20428: test eax, eax
  loc_00E2042A: fnclex
  loc_00E2042C: jge 00E20440h
  loc_00E2042E: push 0000009Ch
  loc_00E20433: push 006DCAD0h
  loc_00E20438: push edi
  loc_00E20439: push eax
  loc_00E2043A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20440: lea ecx, var_20
  loc_00E20443: call ebx
  loc_00E20445: mov eax, [00E3D060h]
  loc_00E2044A: test eax, eax
  loc_00E2044C: jnz 00E20463h
  loc_00E2044E: push 00E3D060h
  loc_00E20453: push 006D87C0h
  loc_00E20458: call [004011A0h] ; __vbaNew2
  loc_00E2045E: mov eax, [00E3D060h]
  loc_00E20463: mov ecx, [eax]
  loc_00E20465: push eax
  loc_00E20466: call [ecx+000003FCh]
  loc_00E2046C: lea edx, var_20
  loc_00E2046F: push eax
  loc_00E20470: push edx
  loc_00E20471: call __vbaObjSet
  loc_00E20473: mov edi, eax
  loc_00E20475: push 00023668h
  loc_00E2047A: mov ebx, [edi]
  loc_00E2047C: call [00401018h] ; __vbaStrI4
  loc_00E20482: mov edx, eax
  loc_00E20484: lea ecx, var_18
  loc_00E20487: call [00401228h] ; __vbaStrMove
  loc_00E2048D: push eax
  loc_00E2048E: push edi
  loc_00E2048F: call [ebx+00000054h]
  loc_00E20492: test eax, eax
  loc_00E20494: fnclex
  loc_00E20496: jge 00E204A7h
  loc_00E20498: push 00000054h
  loc_00E2049A: push 006E0410h
  loc_00E2049F: push edi
  loc_00E204A0: push eax
  loc_00E204A1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E204A7: lea ecx, var_18
  loc_00E204AA: call [00401258h] ; __vbaFreeStr
  loc_00E204B0: lea ecx, var_20
  loc_00E204B3: call [00401254h] ; __vbaFreeObj
  loc_00E204B9: mov eax, [00E3D060h]
  loc_00E204BE: test eax, eax
  loc_00E204C0: jnz 00E204D7h
  loc_00E204C2: push 00E3D060h
  loc_00E204C7: push 006D87C0h
  loc_00E204CC: call [004011A0h] ; __vbaNew2
  loc_00E204D2: mov eax, [00E3D060h]
  loc_00E204D7: mov ecx, [eax]
  loc_00E204D9: push eax
  loc_00E204DA: call [ecx+000003ECh]
  loc_00E204E0: lea edx, var_20
  loc_00E204E3: push eax
  loc_00E204E4: push edx
  loc_00E204E5: call __vbaObjSet
  loc_00E204E7: mov edi, eax
  loc_00E204E9: push 00030D40h
  loc_00E204EE: mov ebx, [edi]
  loc_00E204F0: call [00401018h] ; __vbaStrI4
  loc_00E204F6: mov edx, eax
  loc_00E204F8: lea ecx, var_18
  loc_00E204FB: call [00401228h] ; __vbaStrMove
  loc_00E20501: push eax
  loc_00E20502: push edi
  loc_00E20503: call [ebx+00000054h]
  loc_00E20506: test eax, eax
  loc_00E20508: fnclex
  loc_00E2050A: jge 00E2051Bh
  loc_00E2050C: push 00000054h
  loc_00E2050E: push 006E0410h
  loc_00E20513: push edi
  loc_00E20514: push eax
  loc_00E20515: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2051B: lea ecx, var_18
  loc_00E2051E: call [00401258h] ; __vbaFreeStr
  loc_00E20524: lea ecx, var_20
  loc_00E20527: call [00401254h] ; __vbaFreeObj
  loc_00E2052D: mov eax, [00E3D060h]
  loc_00E20532: test eax, eax
  loc_00E20534: jnz 00E2054Bh
  loc_00E20536: push 00E3D060h
  loc_00E2053B: push 006D87C0h
  loc_00E20540: call [004011A0h] ; __vbaNew2
  loc_00E20546: mov eax, [00E3D060h]
  loc_00E2054B: mov ecx, [eax]
  loc_00E2054D: push eax
  loc_00E2054E: call [ecx+000003DCh]
  loc_00E20554: lea edx, var_20
  loc_00E20557: push eax
  loc_00E20558: push edx
  loc_00E20559: call __vbaObjSet
  loc_00E2055B: mov edi, eax
  loc_00E2055D: push 0000FDE8h
  loc_00E20562: mov ebx, [edi]
  loc_00E20564: call [00401018h] ; __vbaStrI4
  loc_00E2056A: mov edx, eax
  loc_00E2056C: lea ecx, var_18
  loc_00E2056F: call [00401228h] ; __vbaStrMove
  loc_00E20575: push eax
  loc_00E20576: push edi
  loc_00E20577: call [ebx+00000054h]
  loc_00E2057A: test eax, eax
  loc_00E2057C: fnclex
  loc_00E2057E: jge 00E2058Fh
  loc_00E20580: push 00000054h
  loc_00E20582: push 006E0410h
  loc_00E20587: push edi
  loc_00E20588: push eax
  loc_00E20589: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2058F: lea ecx, var_18
  loc_00E20592: call [00401258h] ; __vbaFreeStr
  loc_00E20598: lea ecx, var_20
  loc_00E2059B: call [00401254h] ; __vbaFreeObj
  loc_00E205A1: mov eax, [00E3D060h]
  loc_00E205A6: test eax, eax
  loc_00E205A8: jnz 00E205BFh
  loc_00E205AA: push 00E3D060h
  loc_00E205AF: push 006D87C0h
  loc_00E205B4: call [004011A0h] ; __vbaNew2
  loc_00E205BA: mov eax, [00E3D060h]
  loc_00E205BF: mov ecx, [eax]
  loc_00E205C1: push eax
  loc_00E205C2: call [ecx+000003CCh]
  loc_00E205C8: lea edx, var_20
  loc_00E205CB: push eax
  loc_00E205CC: push edx
  loc_00E205CD: call __vbaObjSet
  loc_00E205CF: mov edi, eax
  loc_00E205D1: push 0001E848h
  loc_00E205D6: mov ebx, [edi]
  loc_00E205D8: call [00401018h] ; __vbaStrI4
  loc_00E205DE: mov edx, eax
  loc_00E205E0: lea ecx, var_18
  loc_00E205E3: call [00401228h] ; __vbaStrMove
  loc_00E205E9: push eax
  loc_00E205EA: push edi
  loc_00E205EB: call [ebx+00000054h]
  loc_00E205EE: test eax, eax
  loc_00E205F0: fnclex
  loc_00E205F2: jge 00E20603h
  loc_00E205F4: push 00000054h
  loc_00E205F6: push 006E0410h
  loc_00E205FB: push edi
  loc_00E205FC: push eax
  loc_00E205FD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20603: lea ecx, var_18
  loc_00E20606: call [00401258h] ; __vbaFreeStr
  loc_00E2060C: lea ecx, var_20
  loc_00E2060F: call [00401254h] ; __vbaFreeObj
  loc_00E20615: mov eax, [00E3D060h]
  loc_00E2061A: test eax, eax
  loc_00E2061C: jnz 00E20633h
  loc_00E2061E: push 00E3D060h
  loc_00E20623: push 006D87C0h
  loc_00E20628: call [004011A0h] ; __vbaNew2
  loc_00E2062E: mov eax, [00E3D060h]
  loc_00E20633: mov ecx, [eax]
  loc_00E20635: push eax
  loc_00E20636: call [ecx+000003BCh]
  loc_00E2063C: lea edx, var_20
  loc_00E2063F: push eax
  loc_00E20640: push edx
  loc_00E20641: call __vbaObjSet
  loc_00E20643: mov edi, eax
  loc_00E20645: push 00023668h
  loc_00E2064A: mov ebx, [edi]
  loc_00E2064C: call [00401018h] ; __vbaStrI4
  loc_00E20652: mov edx, eax
  loc_00E20654: lea ecx, var_18
  loc_00E20657: call [00401228h] ; __vbaStrMove
  loc_00E2065D: push eax
  loc_00E2065E: push edi
  loc_00E2065F: call [ebx+00000054h]
  loc_00E20662: test eax, eax
  loc_00E20664: fnclex
  loc_00E20666: jge 00E20677h
  loc_00E20668: push 00000054h
  loc_00E2066A: push 006E0410h
  loc_00E2066F: push edi
  loc_00E20670: push eax
  loc_00E20671: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20677: lea ecx, var_18
  loc_00E2067A: call [00401258h] ; __vbaFreeStr
  loc_00E20680: lea ecx, var_20
  loc_00E20683: call [00401254h] ; __vbaFreeObj
  loc_00E20689: mov eax, [00E3D060h]
  loc_00E2068E: test eax, eax
  loc_00E20690: jnz 00E206A7h
  loc_00E20692: push 00E3D060h
  loc_00E20697: push 006D87C0h
  loc_00E2069C: call [004011A0h] ; __vbaNew2
  loc_00E206A2: mov eax, [00E3D060h]
  loc_00E206A7: mov ecx, [eax]
  loc_00E206A9: push eax
  loc_00E206AA: call [ecx+000003A8h]
  loc_00E206B0: lea edx, var_20
  loc_00E206B3: push eax
  loc_00E206B4: push edx
  loc_00E206B5: call __vbaObjSet
  loc_00E206B7: mov edi, eax
  loc_00E206B9: push 00003A98h
  loc_00E206BE: mov ebx, [edi]
  loc_00E206C0: call [0040100Ch] ; __vbaStrI2
  loc_00E206C6: mov edx, eax
  loc_00E206C8: lea ecx, var_18
  loc_00E206CB: call [00401228h] ; __vbaStrMove
  loc_00E206D1: push eax
  loc_00E206D2: push edi
  loc_00E206D3: call [ebx+00000054h]
  loc_00E206D6: test eax, eax
  loc_00E206D8: fnclex
  loc_00E206DA: jge 00E206EBh
  loc_00E206DC: push 00000054h
  loc_00E206DE: push 006E0410h
  loc_00E206E3: push edi
  loc_00E206E4: push eax
  loc_00E206E5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E206EB: lea ecx, var_18
  loc_00E206EE: call [00401258h] ; __vbaFreeStr
  loc_00E206F4: lea ecx, var_20
  loc_00E206F7: call [00401254h] ; __vbaFreeObj
  loc_00E206FD: mov eax, [00E3D060h]
  loc_00E20702: test eax, eax
  loc_00E20704: jnz 00E21884h
  loc_00E2070A: jmp 00E2186Fh
  loc_00E2070F: test eax, eax
  loc_00E20711: jnz 00E20728h
  loc_00E20713: push 00E3D060h
  loc_00E20718: push 006D87C0h
  loc_00E2071D: call [004011A0h] ; __vbaNew2
  loc_00E20723: mov eax, [00E3D060h]
  loc_00E20728: mov ecx, [eax]
  loc_00E2072A: push eax
  loc_00E2072B: call [ecx+00000474h]
  loc_00E20731: lea edx, var_20
  loc_00E20734: push eax
  loc_00E20735: push edx
  loc_00E20736: call __vbaObjSet
  loc_00E20738: mov edi, eax
  loc_00E2073A: push 000AAE60h
  loc_00E2073F: mov ebx, [edi]
  loc_00E20741: call [00401018h] ; __vbaStrI4
  loc_00E20747: mov edx, eax
  loc_00E20749: lea ecx, var_18
  loc_00E2074C: call [00401228h] ; __vbaStrMove
  loc_00E20752: push eax
  loc_00E20753: push edi
  loc_00E20754: call [ebx+00000054h]
  loc_00E20757: test eax, eax
  loc_00E20759: fnclex
  loc_00E2075B: jge 00E2076Ch
  loc_00E2075D: push 00000054h
  loc_00E2075F: push 006E0410h
  loc_00E20764: push edi
  loc_00E20765: push eax
  loc_00E20766: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2076C: lea ecx, var_18
  loc_00E2076F: call [00401258h] ; __vbaFreeStr
  loc_00E20775: lea ecx, var_20
  loc_00E20778: call [00401254h] ; __vbaFreeObj
  loc_00E2077E: mov eax, [00E3D060h]
  loc_00E20783: test eax, eax
  loc_00E20785: jnz 00E2079Ch
  loc_00E20787: push 00E3D060h
  loc_00E2078C: push 006D87C0h
  loc_00E20791: call [004011A0h] ; __vbaNew2
  loc_00E20797: mov eax, [00E3D060h]
  loc_00E2079C: mov ecx, [eax]
  loc_00E2079E: push eax
  loc_00E2079F: call [ecx+00000470h]
  loc_00E207A5: lea edx, var_20
  loc_00E207A8: push eax
  loc_00E207A9: push edx
  loc_00E207AA: call __vbaObjSet
  loc_00E207AC: mov edi, eax
  loc_00E207AE: push 0003A980h
  loc_00E207B3: mov ebx, [edi]
  loc_00E207B5: call [00401018h] ; __vbaStrI4
  loc_00E207BB: mov edx, eax
  loc_00E207BD: lea ecx, var_18
  loc_00E207C0: call [00401228h] ; __vbaStrMove
  loc_00E207C6: push eax
  loc_00E207C7: push edi
  loc_00E207C8: call [ebx+00000054h]
  loc_00E207CB: test eax, eax
  loc_00E207CD: fnclex
  loc_00E207CF: jge 00E207E0h
  loc_00E207D1: push 00000054h
  loc_00E207D3: push 006E0410h
  loc_00E207D8: push edi
  loc_00E207D9: push eax
  loc_00E207DA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E207E0: lea ecx, var_18
  loc_00E207E3: call [00401258h] ; __vbaFreeStr
  loc_00E207E9: lea ecx, var_20
  loc_00E207EC: call [00401254h] ; __vbaFreeObj
  loc_00E207F2: mov eax, [00E3D060h]
  loc_00E207F7: test eax, eax
  loc_00E207F9: jnz 00E20810h
  loc_00E207FB: push 00E3D060h
  loc_00E20800: push 006D87C0h
  loc_00E20805: call [004011A0h] ; __vbaNew2
  loc_00E2080B: mov eax, [00E3D060h]
  loc_00E20810: mov ecx, [eax]
  loc_00E20812: push eax
  loc_00E20813: call [ecx+0000046Ch]
  loc_00E20819: lea edx, var_20
  loc_00E2081C: push eax
  loc_00E2081D: push edx
  loc_00E2081E: call __vbaObjSet
  loc_00E20820: mov edi, eax
  loc_00E20822: push 00077A10h
  loc_00E20827: mov ebx, [edi]
  loc_00E20829: call [00401018h] ; __vbaStrI4
  loc_00E2082F: mov edx, eax
  loc_00E20831: lea ecx, var_18
  loc_00E20834: call [00401228h] ; __vbaStrMove
  loc_00E2083A: push eax
  loc_00E2083B: push edi
  loc_00E2083C: call [ebx+00000054h]
  loc_00E2083F: test eax, eax
  loc_00E20841: fnclex
  loc_00E20843: jge 00E20854h
  loc_00E20845: push 00000054h
  loc_00E20847: push 006E0410h
  loc_00E2084C: push edi
  loc_00E2084D: push eax
  loc_00E2084E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20854: lea ecx, var_18
  loc_00E20857: call [00401258h] ; __vbaFreeStr
  loc_00E2085D: lea ecx, var_20
  loc_00E20860: call [00401254h] ; __vbaFreeObj
  loc_00E20866: mov eax, [00E3D060h]
  loc_00E2086B: test eax, eax
  loc_00E2086D: jnz 00E20884h
  loc_00E2086F: push 00E3D060h
  loc_00E20874: push 006D87C0h
  loc_00E20879: call [004011A0h] ; __vbaNew2
  loc_00E2087F: mov eax, [00E3D060h]
  loc_00E20884: mov ecx, [eax]
  loc_00E20886: push eax
  loc_00E20887: call [ecx+00000454h]
  loc_00E2088D: lea edx, var_20
  loc_00E20890: push eax
  loc_00E20891: push edx
  loc_00E20892: call __vbaObjSet
  loc_00E20894: mov edi, eax
  loc_00E20896: push 000124F8h
  loc_00E2089B: mov ebx, [edi]
  loc_00E2089D: call [00401018h] ; __vbaStrI4
  loc_00E208A3: mov edx, eax
  loc_00E208A5: lea ecx, var_18
  loc_00E208A8: call [00401228h] ; __vbaStrMove
  loc_00E208AE: push eax
  loc_00E208AF: push edi
  loc_00E208B0: call [ebx+00000054h]
  loc_00E208B3: test eax, eax
  loc_00E208B5: fnclex
  loc_00E208B7: jge 00E208C8h
  loc_00E208B9: push 00000054h
  loc_00E208BB: push 006E0410h
  loc_00E208C0: push edi
  loc_00E208C1: push eax
  loc_00E208C2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E208C8: lea ecx, var_18
  loc_00E208CB: call [00401258h] ; __vbaFreeStr
  loc_00E208D1: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E208D7: lea ecx, var_20
  loc_00E208DA: call ebx
  loc_00E208DC: mov eax, [00E3D060h]
  loc_00E208E1: test eax, eax
  loc_00E208E3: jnz 00E208FAh
  loc_00E208E5: push 00E3D060h
  loc_00E208EA: push 006D87C0h
  loc_00E208EF: call [004011A0h] ; __vbaNew2
  loc_00E208F5: mov eax, [00E3D060h]
  loc_00E208FA: mov ecx, [eax]
  loc_00E208FC: push eax
  loc_00E208FD: call [ecx+0000042Ch]
  loc_00E20903: lea edx, var_20
  loc_00E20906: push eax
  loc_00E20907: push edx
  loc_00E20908: call __vbaObjSet
  loc_00E2090A: mov edi, eax
  loc_00E2090C: push 00000000h
  loc_00E2090E: push edi
  loc_00E2090F: mov eax, [edi]
  loc_00E20911: call [eax+0000009Ch]
  loc_00E20917: test eax, eax
  loc_00E20919: fnclex
  loc_00E2091B: jge 00E2092Fh
  loc_00E2091D: push 0000009Ch
  loc_00E20922: push 006DCAD0h
  loc_00E20927: push edi
  loc_00E20928: push eax
  loc_00E20929: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2092F: lea ecx, var_20
  loc_00E20932: call ebx
  loc_00E20934: mov eax, [00E3D060h]
  loc_00E20939: test eax, eax
  loc_00E2093B: jnz 00E20952h
  loc_00E2093D: push 00E3D060h
  loc_00E20942: push 006D87C0h
  loc_00E20947: call [004011A0h] ; __vbaNew2
  loc_00E2094D: mov eax, [00E3D060h]
  loc_00E20952: mov ecx, [eax]
  loc_00E20954: push eax
  loc_00E20955: call [ecx+000003FCh]
  loc_00E2095B: lea edx, var_20
  loc_00E2095E: push eax
  loc_00E2095F: push edx
  loc_00E20960: call __vbaObjSet
  loc_00E20962: mov edi, eax
  loc_00E20964: push 00023668h
  loc_00E20969: mov ebx, [edi]
  loc_00E2096B: call [00401018h] ; __vbaStrI4
  loc_00E20971: mov edx, eax
  loc_00E20973: lea ecx, var_18
  loc_00E20976: call [00401228h] ; __vbaStrMove
  loc_00E2097C: push eax
  loc_00E2097D: push edi
  loc_00E2097E: call [ebx+00000054h]
  loc_00E20981: test eax, eax
  loc_00E20983: fnclex
  loc_00E20985: jge 00E20996h
  loc_00E20987: push 00000054h
  loc_00E20989: push 006E0410h
  loc_00E2098E: push edi
  loc_00E2098F: push eax
  loc_00E20990: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20996: lea ecx, var_18
  loc_00E20999: call [00401258h] ; __vbaFreeStr
  loc_00E2099F: lea ecx, var_20
  loc_00E209A2: call [00401254h] ; __vbaFreeObj
  loc_00E209A8: mov eax, [00E3D060h]
  loc_00E209AD: test eax, eax
  loc_00E209AF: jnz 00E209C6h
  loc_00E209B1: push 00E3D060h
  loc_00E209B6: push 006D87C0h
  loc_00E209BB: call [004011A0h] ; __vbaNew2
  loc_00E209C1: mov eax, [00E3D060h]
  loc_00E209C6: mov ecx, [eax]
  loc_00E209C8: push eax
  loc_00E209C9: call [ecx+000003ECh]
  loc_00E209CF: lea edx, var_20
  loc_00E209D2: push eax
  loc_00E209D3: push edx
  loc_00E209D4: call __vbaObjSet
  loc_00E209D6: mov edi, eax
  loc_00E209D8: push 00030D40h
  loc_00E209DD: mov ebx, [edi]
  loc_00E209DF: call [00401018h] ; __vbaStrI4
  loc_00E209E5: mov edx, eax
  loc_00E209E7: lea ecx, var_18
  loc_00E209EA: call [00401228h] ; __vbaStrMove
  loc_00E209F0: push eax
  loc_00E209F1: push edi
  loc_00E209F2: call [ebx+00000054h]
  loc_00E209F5: test eax, eax
  loc_00E209F7: fnclex
  loc_00E209F9: jge 00E20A0Ah
  loc_00E209FB: push 00000054h
  loc_00E209FD: push 006E0410h
  loc_00E20A02: push edi
  loc_00E20A03: push eax
  loc_00E20A04: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20A0A: lea ecx, var_18
  loc_00E20A0D: call [00401258h] ; __vbaFreeStr
  loc_00E20A13: lea ecx, var_20
  loc_00E20A16: call [00401254h] ; __vbaFreeObj
  loc_00E20A1C: mov eax, [00E3D060h]
  loc_00E20A21: test eax, eax
  loc_00E20A23: jnz 00E20A3Ah
  loc_00E20A25: push 00E3D060h
  loc_00E20A2A: push 006D87C0h
  loc_00E20A2F: call [004011A0h] ; __vbaNew2
  loc_00E20A35: mov eax, [00E3D060h]
  loc_00E20A3A: mov ecx, [eax]
  loc_00E20A3C: push eax
  loc_00E20A3D: call [ecx+000003DCh]
  loc_00E20A43: lea edx, var_20
  loc_00E20A46: push eax
  loc_00E20A47: push edx
  loc_00E20A48: call __vbaObjSet
  loc_00E20A4A: mov edi, eax
  loc_00E20A4C: push 0000FDE8h
  loc_00E20A51: mov ebx, [edi]
  loc_00E20A53: call [00401018h] ; __vbaStrI4
  loc_00E20A59: mov edx, eax
  loc_00E20A5B: lea ecx, var_18
  loc_00E20A5E: call [00401228h] ; __vbaStrMove
  loc_00E20A64: push eax
  loc_00E20A65: push edi
  loc_00E20A66: call [ebx+00000054h]
  loc_00E20A69: test eax, eax
  loc_00E20A6B: fnclex
  loc_00E20A6D: jge 00E20A7Eh
  loc_00E20A6F: push 00000054h
  loc_00E20A71: push 006E0410h
  loc_00E20A76: push edi
  loc_00E20A77: push eax
  loc_00E20A78: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20A7E: lea ecx, var_18
  loc_00E20A81: call [00401258h] ; __vbaFreeStr
  loc_00E20A87: lea ecx, var_20
  loc_00E20A8A: call [00401254h] ; __vbaFreeObj
  loc_00E20A90: mov eax, [00E3D060h]
  loc_00E20A95: test eax, eax
  loc_00E20A97: jnz 00E20AAEh
  loc_00E20A99: push 00E3D060h
  loc_00E20A9E: push 006D87C0h
  loc_00E20AA3: call [004011A0h] ; __vbaNew2
  loc_00E20AA9: mov eax, [00E3D060h]
  loc_00E20AAE: mov ecx, [eax]
  loc_00E20AB0: push eax
  loc_00E20AB1: call [ecx+000003CCh]
  loc_00E20AB7: lea edx, var_20
  loc_00E20ABA: push eax
  loc_00E20ABB: push edx
  loc_00E20ABC: call __vbaObjSet
  loc_00E20ABE: mov edi, eax
  loc_00E20AC0: push 0001E848h
  loc_00E20AC5: mov ebx, [edi]
  loc_00E20AC7: call [00401018h] ; __vbaStrI4
  loc_00E20ACD: mov edx, eax
  loc_00E20ACF: lea ecx, var_18
  loc_00E20AD2: call [00401228h] ; __vbaStrMove
  loc_00E20AD8: push eax
  loc_00E20AD9: push edi
  loc_00E20ADA: call [ebx+00000054h]
  loc_00E20ADD: test eax, eax
  loc_00E20ADF: fnclex
  loc_00E20AE1: jge 00E20AF2h
  loc_00E20AE3: push 00000054h
  loc_00E20AE5: push 006E0410h
  loc_00E20AEA: push edi
  loc_00E20AEB: push eax
  loc_00E20AEC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20AF2: lea ecx, var_18
  loc_00E20AF5: call [00401258h] ; __vbaFreeStr
  loc_00E20AFB: lea ecx, var_20
  loc_00E20AFE: call [00401254h] ; __vbaFreeObj
  loc_00E20B04: mov eax, [00E3D060h]
  loc_00E20B09: test eax, eax
  loc_00E20B0B: jnz 00E20B22h
  loc_00E20B0D: push 00E3D060h
  loc_00E20B12: push 006D87C0h
  loc_00E20B17: call [004011A0h] ; __vbaNew2
  loc_00E20B1D: mov eax, [00E3D060h]
  loc_00E20B22: mov ecx, [eax]
  loc_00E20B24: push eax
  loc_00E20B25: call [ecx+000003BCh]
  loc_00E20B2B: lea edx, var_20
  loc_00E20B2E: push eax
  loc_00E20B2F: push edx
  loc_00E20B30: call __vbaObjSet
  loc_00E20B32: mov edi, eax
  loc_00E20B34: push 00023668h
  loc_00E20B39: mov ebx, [edi]
  loc_00E20B3B: call [00401018h] ; __vbaStrI4
  loc_00E20B41: mov edx, eax
  loc_00E20B43: lea ecx, var_18
  loc_00E20B46: call [00401228h] ; __vbaStrMove
  loc_00E20B4C: push eax
  loc_00E20B4D: push edi
  loc_00E20B4E: call [ebx+00000054h]
  loc_00E20B51: test eax, eax
  loc_00E20B53: fnclex
  loc_00E20B55: jge 00E20B66h
  loc_00E20B57: push 00000054h
  loc_00E20B59: push 006E0410h
  loc_00E20B5E: push edi
  loc_00E20B5F: push eax
  loc_00E20B60: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20B66: lea ecx, var_18
  loc_00E20B69: call [00401258h] ; __vbaFreeStr
  loc_00E20B6F: lea ecx, var_20
  loc_00E20B72: call [00401254h] ; __vbaFreeObj
  loc_00E20B78: mov eax, [00E3D060h]
  loc_00E20B7D: test eax, eax
  loc_00E20B7F: jnz 00E20B96h
  loc_00E20B81: push 00E3D060h
  loc_00E20B86: push 006D87C0h
  loc_00E20B8B: call [004011A0h] ; __vbaNew2
  loc_00E20B91: mov eax, [00E3D060h]
  loc_00E20B96: mov ecx, [eax]
  loc_00E20B98: push eax
  loc_00E20B99: call [ecx+000003A8h]
  loc_00E20B9F: lea edx, var_20
  loc_00E20BA2: push eax
  loc_00E20BA3: push edx
  loc_00E20BA4: call __vbaObjSet
  loc_00E20BA6: mov edi, eax
  loc_00E20BA8: push 00003A98h
  loc_00E20BAD: mov ebx, [edi]
  loc_00E20BAF: call [0040100Ch] ; __vbaStrI2
  loc_00E20BB5: mov edx, eax
  loc_00E20BB7: lea ecx, var_18
  loc_00E20BBA: call [00401228h] ; __vbaStrMove
  loc_00E20BC0: push eax
  loc_00E20BC1: push edi
  loc_00E20BC2: call [ebx+00000054h]
  loc_00E20BC5: test eax, eax
  loc_00E20BC7: fnclex
  loc_00E20BC9: jge 00E20BDAh
  loc_00E20BCB: push 00000054h
  loc_00E20BCD: push 006E0410h
  loc_00E20BD2: push edi
  loc_00E20BD3: push eax
  loc_00E20BD4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20BDA: lea ecx, var_18
  loc_00E20BDD: call [00401258h] ; __vbaFreeStr
  loc_00E20BE3: lea ecx, var_20
  loc_00E20BE6: call [00401254h] ; __vbaFreeObj
  loc_00E20BEC: mov eax, [00E3D060h]
  loc_00E20BF1: test eax, eax
  loc_00E20BF3: jnz 00E20C0Ah
  loc_00E20BF5: push 00E3D060h
  loc_00E20BFA: push 006D87C0h
  loc_00E20BFF: call [004011A0h] ; __vbaNew2
  loc_00E20C05: mov eax, [00E3D060h]
  loc_00E20C0A: mov ecx, [eax]
  loc_00E20C0C: push eax
  loc_00E20C0D: call [ecx+00000398h]
  loc_00E20C13: lea edx, var_20
  loc_00E20C16: push eax
  loc_00E20C17: push edx
  loc_00E20C18: call __vbaObjSet
  loc_00E20C1A: mov edi, eax
  loc_00E20C1C: push 00003A98h
  loc_00E20C21: mov ebx, [edi]
  loc_00E20C23: call [0040100Ch] ; __vbaStrI2
  loc_00E20C29: mov edx, eax
  loc_00E20C2B: lea ecx, var_18
  loc_00E20C2E: call [00401228h] ; __vbaStrMove
  loc_00E20C34: push eax
  loc_00E20C35: push edi
  loc_00E20C36: call [ebx+00000054h]
  loc_00E20C39: test eax, eax
  loc_00E20C3B: fnclex
  loc_00E20C3D: jge 00E20C4Eh
  loc_00E20C3F: push 00000054h
  loc_00E20C41: push 006E0410h
  loc_00E20C46: push edi
  loc_00E20C47: push eax
  loc_00E20C48: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20C4E: lea ecx, var_18
  loc_00E20C51: call [00401258h] ; __vbaFreeStr
  loc_00E20C57: lea ecx, var_20
  loc_00E20C5A: call [00401254h] ; __vbaFreeObj
  loc_00E20C60: mov eax, [00E3D060h]
  loc_00E20C65: test eax, eax
  loc_00E20C67: jnz 00E20C7Eh
  loc_00E20C69: push 00E3D060h
  loc_00E20C6E: push 006D87C0h
  loc_00E20C73: call [004011A0h] ; __vbaNew2
  loc_00E20C79: mov eax, [00E3D060h]
  loc_00E20C7E: mov ecx, [eax]
  loc_00E20C80: push eax
  loc_00E20C81: call [ecx+00000388h]
  loc_00E20C87: lea edx, var_20
  loc_00E20C8A: push eax
  loc_00E20C8B: push edx
  loc_00E20C8C: call __vbaObjSet
  loc_00E20C8E: mov edi, eax
  loc_00E20C90: push 00001388h
  loc_00E20C95: mov ebx, [edi]
  loc_00E20C97: call [0040100Ch] ; __vbaStrI2
  loc_00E20C9D: mov edx, eax
  loc_00E20C9F: lea ecx, var_18
  loc_00E20CA2: call [00401228h] ; __vbaStrMove
  loc_00E20CA8: push eax
  loc_00E20CA9: push edi
  loc_00E20CAA: call [ebx+00000054h]
  loc_00E20CAD: test eax, eax
  loc_00E20CAF: fnclex
  loc_00E20CB1: jge 00E220CDh
  loc_00E20CB7: jmp 00E220BEh
  loc_00E20CBC: test eax, eax
  loc_00E20CBE: jnz 00E20CD5h
  loc_00E20CC0: push 00E3D060h
  loc_00E20CC5: push 006D87C0h
  loc_00E20CCA: call [004011A0h] ; __vbaNew2
  loc_00E20CD0: mov eax, [00E3D060h]
  loc_00E20CD5: mov ecx, [eax]
  loc_00E20CD7: push eax
  loc_00E20CD8: call [ecx+000004E4h]
  loc_00E20CDE: lea edx, var_20
  loc_00E20CE1: push eax
  loc_00E20CE2: push edx
  loc_00E20CE3: call __vbaObjSet
  loc_00E20CE5: mov edi, eax
  loc_00E20CE7: lea ecx, var_18
  loc_00E20CEA: push ecx
  loc_00E20CEB: push edi
  loc_00E20CEC: mov eax, [edi]
  loc_00E20CEE: call [eax+00000050h]
  loc_00E20CF1: test eax, eax
  loc_00E20CF3: fnclex
  loc_00E20CF5: jge 00E20D06h
  loc_00E20CF7: push 00000050h
  loc_00E20CF9: push 006E0410h
  loc_00E20CFE: push edi
  loc_00E20CFF: push eax
  loc_00E20D00: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20D06: mov eax, [00E3D060h]
  loc_00E20D0B: test eax, eax
  loc_00E20D0D: jnz 00E20D24h
  loc_00E20D0F: push 00E3D060h
  loc_00E20D14: push 006D87C0h
  loc_00E20D19: call [004011A0h] ; __vbaNew2
  loc_00E20D1F: mov eax, [00E3D060h]
  loc_00E20D24: mov edx, [eax]
  loc_00E20D26: push eax
  loc_00E20D27: call [edx+00000518h]
  loc_00E20D2D: push eax
  loc_00E20D2E: lea eax, var_24
  loc_00E20D31: push eax
  loc_00E20D32: call __vbaObjSet
  loc_00E20D34: mov edi, eax
  loc_00E20D36: lea edx, var_1C
  loc_00E20D39: push edx
  loc_00E20D3A: push edi
  loc_00E20D3B: mov ecx, [edi]
  loc_00E20D3D: call [ecx+00000050h]
  loc_00E20D40: test eax, eax
  loc_00E20D42: fnclex
  loc_00E20D44: jge 00E20D55h
  loc_00E20D46: push 00000050h
  loc_00E20D48: push 006E0410h
  loc_00E20D4D: push edi
  loc_00E20D4E: push eax
  loc_00E20D4F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20D55: mov eax, var_1C
  loc_00E20D58: push eax
  loc_00E20D59: push 006E049Ch ; "Tata Busana"
  loc_00E20D5E: call ebx
  loc_00E20D60: mov ecx, var_18
  loc_00E20D63: mov edi, eax
  loc_00E20D65: neg edi
  loc_00E20D67: sbb edi, edi
  loc_00E20D69: push ecx
  loc_00E20D6A: inc edi
  loc_00E20D6B: push 006E0B40h ; "Perempuan"
  loc_00E20D70: neg edi
  loc_00E20D72: call ebx
  loc_00E20D74: neg eax
  loc_00E20D76: sbb eax, eax
  loc_00E20D78: lea edx, var_1C
  loc_00E20D7B: inc eax
  loc_00E20D7C: push edx
  loc_00E20D7D: neg eax
  loc_00E20D7F: and edi, eax
  loc_00E20D81: lea eax, var_18
  loc_00E20D84: push eax
  loc_00E20D85: push 00000002h
  loc_00E20D87: call [004011BCh] ; __vbaFreeStrList
  loc_00E20D8D: lea ecx, var_24
  loc_00E20D90: lea edx, var_20
  loc_00E20D93: push ecx
  loc_00E20D94: push edx
  loc_00E20D95: push 00000002h
  loc_00E20D97: call [00401048h] ; __vbaFreeObjList
  loc_00E20D9D: mov eax, [00E3D060h]
  loc_00E20DA2: add esp, 00000018h
  loc_00E20DA5: test di, di
  loc_00E20DA8: jz 00E218C2h
  loc_00E20DAE: test eax, eax
  loc_00E20DB0: jnz 00E20DC7h
  loc_00E20DB2: push 00E3D060h
  loc_00E20DB7: push 006D87C0h
  loc_00E20DBC: call [004011A0h] ; __vbaNew2
  loc_00E20DC2: mov eax, [00E3D060h]
  loc_00E20DC7: mov ecx, [eax]
  loc_00E20DC9: push eax
  loc_00E20DCA: call [ecx+000004E0h]
  loc_00E20DD0: lea edx, var_20
  loc_00E20DD3: push eax
  loc_00E20DD4: push edx
  loc_00E20DD5: call __vbaObjSet
  loc_00E20DD7: mov edi, eax
  loc_00E20DD9: lea ecx, var_18
  loc_00E20DDC: push ecx
  loc_00E20DDD: push edi
  loc_00E20DDE: mov eax, [edi]
  loc_00E20DE0: call [eax+00000050h]
  loc_00E20DE3: test eax, eax
  loc_00E20DE5: fnclex
  loc_00E20DE7: jge 00E20DF8h
  loc_00E20DE9: push 00000050h
  loc_00E20DEB: push 006E0410h
  loc_00E20DF0: push edi
  loc_00E20DF1: push eax
  loc_00E20DF2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20DF8: mov edx, var_18
  loc_00E20DFB: push edx
  loc_00E20DFC: push 006E05FCh ; "Katholik"
  loc_00E20E01: call ebx
  loc_00E20E03: mov edi, eax
  loc_00E20E05: lea ecx, var_18
  loc_00E20E08: neg edi
  loc_00E20E0A: sbb edi, edi
  loc_00E20E0C: inc edi
  loc_00E20E0D: neg edi
  loc_00E20E0F: call [00401258h] ; __vbaFreeStr
  loc_00E20E15: lea ecx, var_20
  loc_00E20E18: call [00401254h] ; __vbaFreeObj
  loc_00E20E1E: mov eax, [00E3D060h]
  loc_00E20E23: test di, di
  loc_00E20E26: jz 00E2138Ch
  loc_00E20E2C: test eax, eax
  loc_00E20E2E: jnz 00E20E45h
  loc_00E20E30: push 00E3D060h
  loc_00E20E35: push 006D87C0h
  loc_00E20E3A: call [004011A0h] ; __vbaNew2
  loc_00E20E40: mov eax, [00E3D060h]
  loc_00E20E45: mov ecx, [eax]
  loc_00E20E47: push eax
  loc_00E20E48: call [ecx+00000474h]
  loc_00E20E4E: lea edx, var_20
  loc_00E20E51: push eax
  loc_00E20E52: push edx
  loc_00E20E53: call __vbaObjSet
  loc_00E20E55: mov edi, eax
  loc_00E20E57: push 000927C0h
  loc_00E20E5C: mov ebx, [edi]
  loc_00E20E5E: call [00401018h] ; __vbaStrI4
  loc_00E20E64: mov edx, eax
  loc_00E20E66: lea ecx, var_18
  loc_00E20E69: call [00401228h] ; __vbaStrMove
  loc_00E20E6F: push eax
  loc_00E20E70: push edi
  loc_00E20E71: call [ebx+00000054h]
  loc_00E20E74: test eax, eax
  loc_00E20E76: fnclex
  loc_00E20E78: jge 00E20E89h
  loc_00E20E7A: push 00000054h
  loc_00E20E7C: push 006E0410h
  loc_00E20E81: push edi
  loc_00E20E82: push eax
  loc_00E20E83: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20E89: lea ecx, var_18
  loc_00E20E8C: call [00401258h] ; __vbaFreeStr
  loc_00E20E92: lea ecx, var_20
  loc_00E20E95: call [00401254h] ; __vbaFreeObj
  loc_00E20E9B: mov eax, [00E3D060h]
  loc_00E20EA0: test eax, eax
  loc_00E20EA2: jnz 00E20EB9h
  loc_00E20EA4: push 00E3D060h
  loc_00E20EA9: push 006D87C0h
  loc_00E20EAE: call [004011A0h] ; __vbaNew2
  loc_00E20EB4: mov eax, [00E3D060h]
  loc_00E20EB9: mov ecx, [eax]
  loc_00E20EBB: push eax
  loc_00E20EBC: call [ecx+00000470h]
  loc_00E20EC2: lea edx, var_20
  loc_00E20EC5: push eax
  loc_00E20EC6: push edx
  loc_00E20EC7: call __vbaObjSet
  loc_00E20EC9: mov edi, eax
  loc_00E20ECB: push 00030D40h
  loc_00E20ED0: mov ebx, [edi]
  loc_00E20ED2: call [00401018h] ; __vbaStrI4
  loc_00E20ED8: mov edx, eax
  loc_00E20EDA: lea ecx, var_18
  loc_00E20EDD: call [00401228h] ; __vbaStrMove
  loc_00E20EE3: push eax
  loc_00E20EE4: push edi
  loc_00E20EE5: call [ebx+00000054h]
  loc_00E20EE8: test eax, eax
  loc_00E20EEA: fnclex
  loc_00E20EEC: jge 00E20EFDh
  loc_00E20EEE: push 00000054h
  loc_00E20EF0: push 006E0410h
  loc_00E20EF5: push edi
  loc_00E20EF6: push eax
  loc_00E20EF7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20EFD: lea ecx, var_18
  loc_00E20F00: call [00401258h] ; __vbaFreeStr
  loc_00E20F06: lea ecx, var_20
  loc_00E20F09: call [00401254h] ; __vbaFreeObj
  loc_00E20F0F: mov eax, [00E3D060h]
  loc_00E20F14: test eax, eax
  loc_00E20F16: jnz 00E20F2Dh
  loc_00E20F18: push 00E3D060h
  loc_00E20F1D: push 006D87C0h
  loc_00E20F22: call [004011A0h] ; __vbaNew2
  loc_00E20F28: mov eax, [00E3D060h]
  loc_00E20F2D: mov ecx, [eax]
  loc_00E20F2F: push eax
  loc_00E20F30: call [ecx+0000046Ch]
  loc_00E20F36: lea edx, var_20
  loc_00E20F39: push eax
  loc_00E20F3A: push edx
  loc_00E20F3B: call __vbaObjSet
  loc_00E20F3D: mov edi, eax
  loc_00E20F3F: push 00077A10h
  loc_00E20F44: mov ebx, [edi]
  loc_00E20F46: call [00401018h] ; __vbaStrI4
  loc_00E20F4C: mov edx, eax
  loc_00E20F4E: lea ecx, var_18
  loc_00E20F51: call [00401228h] ; __vbaStrMove
  loc_00E20F57: push eax
  loc_00E20F58: push edi
  loc_00E20F59: call [ebx+00000054h]
  loc_00E20F5C: test eax, eax
  loc_00E20F5E: fnclex
  loc_00E20F60: jge 00E20F71h
  loc_00E20F62: push 00000054h
  loc_00E20F64: push 006E0410h
  loc_00E20F69: push edi
  loc_00E20F6A: push eax
  loc_00E20F6B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20F71: lea ecx, var_18
  loc_00E20F74: call [00401258h] ; __vbaFreeStr
  loc_00E20F7A: lea ecx, var_20
  loc_00E20F7D: call [00401254h] ; __vbaFreeObj
  loc_00E20F83: mov eax, [00E3D060h]
  loc_00E20F88: test eax, eax
  loc_00E20F8A: jnz 00E20FA1h
  loc_00E20F8C: push 00E3D060h
  loc_00E20F91: push 006D87C0h
  loc_00E20F96: call [004011A0h] ; __vbaNew2
  loc_00E20F9C: mov eax, [00E3D060h]
  loc_00E20FA1: mov ecx, [eax]
  loc_00E20FA3: push eax
  loc_00E20FA4: call [ecx+00000454h]
  loc_00E20FAA: lea edx, var_20
  loc_00E20FAD: push eax
  loc_00E20FAE: push edx
  loc_00E20FAF: call __vbaObjSet
  loc_00E20FB1: mov edi, eax
  loc_00E20FB3: push 000124F8h
  loc_00E20FB8: mov ebx, [edi]
  loc_00E20FBA: call [00401018h] ; __vbaStrI4
  loc_00E20FC0: mov edx, eax
  loc_00E20FC2: lea ecx, var_18
  loc_00E20FC5: call [00401228h] ; __vbaStrMove
  loc_00E20FCB: push eax
  loc_00E20FCC: push edi
  loc_00E20FCD: call [ebx+00000054h]
  loc_00E20FD0: test eax, eax
  loc_00E20FD2: fnclex
  loc_00E20FD4: jge 00E20FE5h
  loc_00E20FD6: push 00000054h
  loc_00E20FD8: push 006E0410h
  loc_00E20FDD: push edi
  loc_00E20FDE: push eax
  loc_00E20FDF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E20FE5: lea ecx, var_18
  loc_00E20FE8: call [00401258h] ; __vbaFreeStr
  loc_00E20FEE: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E20FF4: lea ecx, var_20
  loc_00E20FF7: call ebx
  loc_00E20FF9: mov eax, [00E3D060h]
  loc_00E20FFE: test eax, eax
  loc_00E21000: jnz 00E21017h
  loc_00E21002: push 00E3D060h
  loc_00E21007: push 006D87C0h
  loc_00E2100C: call [004011A0h] ; __vbaNew2
  loc_00E21012: mov eax, [00E3D060h]
  loc_00E21017: mov ecx, [eax]
  loc_00E21019: push eax
  loc_00E2101A: call [ecx+0000042Ch]
  loc_00E21020: lea edx, var_20
  loc_00E21023: push eax
  loc_00E21024: push edx
  loc_00E21025: call __vbaObjSet
  loc_00E21027: mov edi, eax
  loc_00E21029: push 00000000h
  loc_00E2102B: push edi
  loc_00E2102C: mov eax, [edi]
  loc_00E2102E: call [eax+0000009Ch]
  loc_00E21034: test eax, eax
  loc_00E21036: fnclex
  loc_00E21038: jge 00E2104Ch
  loc_00E2103A: push 0000009Ch
  loc_00E2103F: push 006DCAD0h
  loc_00E21044: push edi
  loc_00E21045: push eax
  loc_00E21046: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2104C: lea ecx, var_20
  loc_00E2104F: call ebx
  loc_00E21051: mov eax, [00E3D060h]
  loc_00E21056: test eax, eax
  loc_00E21058: jnz 00E2106Fh
  loc_00E2105A: push 00E3D060h
  loc_00E2105F: push 006D87C0h
  loc_00E21064: call [004011A0h] ; __vbaNew2
  loc_00E2106A: mov eax, [00E3D060h]
  loc_00E2106F: mov ecx, [eax]
  loc_00E21071: push eax
  loc_00E21072: call [ecx+000003FCh]
  loc_00E21078: lea edx, var_20
  loc_00E2107B: push eax
  loc_00E2107C: push edx
  loc_00E2107D: call __vbaObjSet
  loc_00E2107F: mov edi, eax
  loc_00E21081: push 000186A0h
  loc_00E21086: mov ebx, [edi]
  loc_00E21088: call [00401018h] ; __vbaStrI4
  loc_00E2108E: mov edx, eax
  loc_00E21090: lea ecx, var_18
  loc_00E21093: call [00401228h] ; __vbaStrMove
  loc_00E21099: push eax
  loc_00E2109A: push edi
  loc_00E2109B: call [ebx+00000054h]
  loc_00E2109E: test eax, eax
  loc_00E210A0: fnclex
  loc_00E210A2: jge 00E210B3h
  loc_00E210A4: push 00000054h
  loc_00E210A6: push 006E0410h
  loc_00E210AB: push edi
  loc_00E210AC: push eax
  loc_00E210AD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E210B3: lea ecx, var_18
  loc_00E210B6: call [00401258h] ; __vbaFreeStr
  loc_00E210BC: lea ecx, var_20
  loc_00E210BF: call [00401254h] ; __vbaFreeObj
  loc_00E210C5: mov eax, [00E3D060h]
  loc_00E210CA: test eax, eax
  loc_00E210CC: jnz 00E210E3h
  loc_00E210CE: push 00E3D060h
  loc_00E210D3: push 006D87C0h
  loc_00E210D8: call [004011A0h] ; __vbaNew2
  loc_00E210DE: mov eax, [00E3D060h]
  loc_00E210E3: mov ecx, [eax]
  loc_00E210E5: push eax
  loc_00E210E6: call [ecx+000003ECh]
  loc_00E210EC: lea edx, var_20
  loc_00E210EF: push eax
  loc_00E210F0: push edx
  loc_00E210F1: call __vbaObjSet
  loc_00E210F3: mov edi, eax
  loc_00E210F5: push 00000000h
  loc_00E210F7: mov ebx, [edi]
  loc_00E210F9: call [0040100Ch] ; __vbaStrI2
  loc_00E210FF: mov edx, eax
  loc_00E21101: lea ecx, var_18
  loc_00E21104: call [00401228h] ; __vbaStrMove
  loc_00E2110A: push eax
  loc_00E2110B: push edi
  loc_00E2110C: call [ebx+00000054h]
  loc_00E2110F: test eax, eax
  loc_00E21111: fnclex
  loc_00E21113: jge 00E21124h
  loc_00E21115: push 00000054h
  loc_00E21117: push 006E0410h
  loc_00E2111C: push edi
  loc_00E2111D: push eax
  loc_00E2111E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21124: lea ecx, var_18
  loc_00E21127: call [00401258h] ; __vbaFreeStr
  loc_00E2112D: lea ecx, var_20
  loc_00E21130: call [00401254h] ; __vbaFreeObj
  loc_00E21136: mov eax, [00E3D060h]
  loc_00E2113B: test eax, eax
  loc_00E2113D: jnz 00E21154h
  loc_00E2113F: push 00E3D060h
  loc_00E21144: push 006D87C0h
  loc_00E21149: call [004011A0h] ; __vbaNew2
  loc_00E2114F: mov eax, [00E3D060h]
  loc_00E21154: mov ecx, [eax]
  loc_00E21156: push eax
  loc_00E21157: call [ecx+000003DCh]
  loc_00E2115D: lea edx, var_20
  loc_00E21160: push eax
  loc_00E21161: push edx
  loc_00E21162: call __vbaObjSet
  loc_00E21164: mov edi, eax
  loc_00E21166: push 000124F8h
  loc_00E2116B: mov ebx, [edi]
  loc_00E2116D: call [00401018h] ; __vbaStrI4
  loc_00E21173: mov edx, eax
  loc_00E21175: lea ecx, var_18
  loc_00E21178: call [00401228h] ; __vbaStrMove
  loc_00E2117E: push eax
  loc_00E2117F: push edi
  loc_00E21180: call [ebx+00000054h]
  loc_00E21183: test eax, eax
  loc_00E21185: fnclex
  loc_00E21187: jge 00E21198h
  loc_00E21189: push 00000054h
  loc_00E2118B: push 006E0410h
  loc_00E21190: push edi
  loc_00E21191: push eax
  loc_00E21192: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21198: lea ecx, var_18
  loc_00E2119B: call [00401258h] ; __vbaFreeStr
  loc_00E211A1: lea ecx, var_20
  loc_00E211A4: call [00401254h] ; __vbaFreeObj
  loc_00E211AA: mov eax, [00E3D060h]
  loc_00E211AF: test eax, eax
  loc_00E211B1: jnz 00E211C8h
  loc_00E211B3: push 00E3D060h
  loc_00E211B8: push 006D87C0h
  loc_00E211BD: call [004011A0h] ; __vbaNew2
  loc_00E211C3: mov eax, [00E3D060h]
  loc_00E211C8: mov ecx, [eax]
  loc_00E211CA: push eax
  loc_00E211CB: call [ecx+000003CCh]
  loc_00E211D1: lea edx, var_20
  loc_00E211D4: push eax
  loc_00E211D5: push edx
  loc_00E211D6: call __vbaObjSet
  loc_00E211D8: mov edi, eax
  loc_00E211DA: push 000186A0h
  loc_00E211DF: mov ebx, [edi]
  loc_00E211E1: call [00401018h] ; __vbaStrI4
  loc_00E211E7: mov edx, eax
  loc_00E211E9: lea ecx, var_18
  loc_00E211EC: call [00401228h] ; __vbaStrMove
  loc_00E211F2: push eax
  loc_00E211F3: push edi
  loc_00E211F4: call [ebx+00000054h]
  loc_00E211F7: test eax, eax
  loc_00E211F9: fnclex
  loc_00E211FB: jge 00E2120Ch
  loc_00E211FD: push 00000054h
  loc_00E211FF: push 006E0410h
  loc_00E21204: push edi
  loc_00E21205: push eax
  loc_00E21206: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2120C: lea ecx, var_18
  loc_00E2120F: call [00401258h] ; __vbaFreeStr
  loc_00E21215: lea ecx, var_20
  loc_00E21218: call [00401254h] ; __vbaFreeObj
  loc_00E2121E: mov eax, [00E3D060h]
  loc_00E21223: test eax, eax
  loc_00E21225: jnz 00E2123Ch
  loc_00E21227: push 00E3D060h
  loc_00E2122C: push 006D87C0h
  loc_00E21231: call [004011A0h] ; __vbaNew2
  loc_00E21237: mov eax, [00E3D060h]
  loc_00E2123C: mov ecx, [eax]
  loc_00E2123E: push eax
  loc_00E2123F: call [ecx+000003BCh]
  loc_00E21245: lea edx, var_20
  loc_00E21248: push eax
  loc_00E21249: push edx
  loc_00E2124A: call __vbaObjSet
  loc_00E2124C: mov edi, eax
  loc_00E2124E: push 000186A0h
  loc_00E21253: mov ebx, [edi]
  loc_00E21255: call [00401018h] ; __vbaStrI4
  loc_00E2125B: mov edx, eax
  loc_00E2125D: lea ecx, var_18
  loc_00E21260: call [00401228h] ; __vbaStrMove
  loc_00E21266: push eax
  loc_00E21267: push edi
  loc_00E21268: call [ebx+00000054h]
  loc_00E2126B: test eax, eax
  loc_00E2126D: fnclex
  loc_00E2126F: jge 00E21280h
  loc_00E21271: push 00000054h
  loc_00E21273: push 006E0410h
  loc_00E21278: push edi
  loc_00E21279: push eax
  loc_00E2127A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21280: lea ecx, var_18
  loc_00E21283: call [00401258h] ; __vbaFreeStr
  loc_00E21289: lea ecx, var_20
  loc_00E2128C: call [00401254h] ; __vbaFreeObj
  loc_00E21292: mov eax, [00E3D060h]
  loc_00E21297: test eax, eax
  loc_00E21299: jnz 00E212B0h
  loc_00E2129B: push 00E3D060h
  loc_00E212A0: push 006D87C0h
  loc_00E212A5: call [004011A0h] ; __vbaNew2
  loc_00E212AB: mov eax, [00E3D060h]
  loc_00E212B0: mov ecx, [eax]
  loc_00E212B2: push eax
  loc_00E212B3: call [ecx+000003A8h]
  loc_00E212B9: lea edx, var_20
  loc_00E212BC: push eax
  loc_00E212BD: push edx
  loc_00E212BE: call __vbaObjSet
  loc_00E212C0: mov edi, eax
  loc_00E212C2: push 00003A98h
  loc_00E212C7: mov ebx, [edi]
  loc_00E212C9: call [0040100Ch] ; __vbaStrI2
  loc_00E212CF: mov edx, eax
  loc_00E212D1: lea ecx, var_18
  loc_00E212D4: call [00401228h] ; __vbaStrMove
  loc_00E212DA: push eax
  loc_00E212DB: push edi
  loc_00E212DC: call [ebx+00000054h]
  loc_00E212DF: test eax, eax
  loc_00E212E1: fnclex
  loc_00E212E3: jge 00E212F4h
  loc_00E212E5: push 00000054h
  loc_00E212E7: push 006E0410h
  loc_00E212EC: push edi
  loc_00E212ED: push eax
  loc_00E212EE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E212F4: lea ecx, var_18
  loc_00E212F7: call [00401258h] ; __vbaFreeStr
  loc_00E212FD: lea ecx, var_20
  loc_00E21300: call [00401254h] ; __vbaFreeObj
  loc_00E21306: mov eax, [00E3D060h]
  loc_00E2130B: test eax, eax
  loc_00E2130D: jnz 00E21324h
  loc_00E2130F: push 00E3D060h
  loc_00E21314: push 006D87C0h
  loc_00E21319: call [004011A0h] ; __vbaNew2
  loc_00E2131F: mov eax, [00E3D060h]
  loc_00E21324: mov ecx, [eax]
  loc_00E21326: push eax
  loc_00E21327: call [ecx+00000398h]
  loc_00E2132D: lea edx, var_20
  loc_00E21330: push eax
  loc_00E21331: push edx
  loc_00E21332: call __vbaObjSet
  loc_00E21334: mov edi, eax
  loc_00E21336: push 00003A98h
  loc_00E2133B: mov ebx, [edi]
  loc_00E2133D: call [0040100Ch] ; __vbaStrI2
  loc_00E21343: mov edx, eax
  loc_00E21345: lea ecx, var_18
  loc_00E21348: call [00401228h] ; __vbaStrMove
  loc_00E2134E: push eax
  loc_00E2134F: push edi
  loc_00E21350: call [ebx+00000054h]
  loc_00E21353: test eax, eax
  loc_00E21355: fnclex
  loc_00E21357: jge 00E21368h
  loc_00E21359: push 00000054h
  loc_00E2135B: push 006E0410h
  loc_00E21360: push edi
  loc_00E21361: push eax
  loc_00E21362: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21368: lea ecx, var_18
  loc_00E2136B: call [00401258h] ; __vbaFreeStr
  loc_00E21371: lea ecx, var_20
  loc_00E21374: call [00401254h] ; __vbaFreeObj
  loc_00E2137A: mov eax, [00E3D060h]
  loc_00E2137F: test eax, eax
  loc_00E21381: jnz 00E20C7Eh
  loc_00E21387: jmp 00E20C69h
  loc_00E2138C: test eax, eax
  loc_00E2138E: jnz 00E213A5h
  loc_00E21390: push 00E3D060h
  loc_00E21395: push 006D87C0h
  loc_00E2139A: call [004011A0h] ; __vbaNew2
  loc_00E213A0: mov eax, [00E3D060h]
  loc_00E213A5: mov ecx, [eax]
  loc_00E213A7: push eax
  loc_00E213A8: call [ecx+00000474h]
  loc_00E213AE: lea edx, var_20
  loc_00E213B1: push eax
  loc_00E213B2: push edx
  loc_00E213B3: call __vbaObjSet
  loc_00E213B5: mov edi, eax
  loc_00E213B7: push 000AAE60h
  loc_00E213BC: mov ebx, [edi]
  loc_00E213BE: call [00401018h] ; __vbaStrI4
  loc_00E213C4: mov edx, eax
  loc_00E213C6: lea ecx, var_18
  loc_00E213C9: call [00401228h] ; __vbaStrMove
  loc_00E213CF: push eax
  loc_00E213D0: push edi
  loc_00E213D1: call [ebx+00000054h]
  loc_00E213D4: test eax, eax
  loc_00E213D6: fnclex
  loc_00E213D8: jge 00E213E9h
  loc_00E213DA: push 00000054h
  loc_00E213DC: push 006E0410h
  loc_00E213E1: push edi
  loc_00E213E2: push eax
  loc_00E213E3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E213E9: lea ecx, var_18
  loc_00E213EC: call [00401258h] ; __vbaFreeStr
  loc_00E213F2: lea ecx, var_20
  loc_00E213F5: call [00401254h] ; __vbaFreeObj
  loc_00E213FB: mov eax, [00E3D060h]
  loc_00E21400: test eax, eax
  loc_00E21402: jnz 00E21419h
  loc_00E21404: push 00E3D060h
  loc_00E21409: push 006D87C0h
  loc_00E2140E: call [004011A0h] ; __vbaNew2
  loc_00E21414: mov eax, [00E3D060h]
  loc_00E21419: mov ecx, [eax]
  loc_00E2141B: push eax
  loc_00E2141C: call [ecx+00000470h]
  loc_00E21422: lea edx, var_20
  loc_00E21425: push eax
  loc_00E21426: push edx
  loc_00E21427: call __vbaObjSet
  loc_00E21429: mov edi, eax
  loc_00E2142B: push 00030D40h
  loc_00E21430: mov ebx, [edi]
  loc_00E21432: call [00401018h] ; __vbaStrI4
  loc_00E21438: mov edx, eax
  loc_00E2143A: lea ecx, var_18
  loc_00E2143D: call [00401228h] ; __vbaStrMove
  loc_00E21443: push eax
  loc_00E21444: push edi
  loc_00E21445: call [ebx+00000054h]
  loc_00E21448: test eax, eax
  loc_00E2144A: fnclex
  loc_00E2144C: jge 00E2145Dh
  loc_00E2144E: push 00000054h
  loc_00E21450: push 006E0410h
  loc_00E21455: push edi
  loc_00E21456: push eax
  loc_00E21457: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2145D: lea ecx, var_18
  loc_00E21460: call [00401258h] ; __vbaFreeStr
  loc_00E21466: lea ecx, var_20
  loc_00E21469: call [00401254h] ; __vbaFreeObj
  loc_00E2146F: mov eax, [00E3D060h]
  loc_00E21474: test eax, eax
  loc_00E21476: jnz 00E2148Dh
  loc_00E21478: push 00E3D060h
  loc_00E2147D: push 006D87C0h
  loc_00E21482: call [004011A0h] ; __vbaNew2
  loc_00E21488: mov eax, [00E3D060h]
  loc_00E2148D: mov ecx, [eax]
  loc_00E2148F: push eax
  loc_00E21490: call [ecx+0000046Ch]
  loc_00E21496: lea edx, var_20
  loc_00E21499: push eax
  loc_00E2149A: push edx
  loc_00E2149B: call __vbaObjSet
  loc_00E2149D: mov edi, eax
  loc_00E2149F: push 00077A10h
  loc_00E214A4: mov ebx, [edi]
  loc_00E214A6: call [00401018h] ; __vbaStrI4
  loc_00E214AC: mov edx, eax
  loc_00E214AE: lea ecx, var_18
  loc_00E214B1: call [00401228h] ; __vbaStrMove
  loc_00E214B7: push eax
  loc_00E214B8: push edi
  loc_00E214B9: call [ebx+00000054h]
  loc_00E214BC: test eax, eax
  loc_00E214BE: fnclex
  loc_00E214C0: jge 00E214D1h
  loc_00E214C2: push 00000054h
  loc_00E214C4: push 006E0410h
  loc_00E214C9: push edi
  loc_00E214CA: push eax
  loc_00E214CB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E214D1: lea ecx, var_18
  loc_00E214D4: call [00401258h] ; __vbaFreeStr
  loc_00E214DA: lea ecx, var_20
  loc_00E214DD: call [00401254h] ; __vbaFreeObj
  loc_00E214E3: mov eax, [00E3D060h]
  loc_00E214E8: test eax, eax
  loc_00E214EA: jnz 00E21501h
  loc_00E214EC: push 00E3D060h
  loc_00E214F1: push 006D87C0h
  loc_00E214F6: call [004011A0h] ; __vbaNew2
  loc_00E214FC: mov eax, [00E3D060h]
  loc_00E21501: mov ecx, [eax]
  loc_00E21503: push eax
  loc_00E21504: call [ecx+00000454h]
  loc_00E2150A: lea edx, var_20
  loc_00E2150D: push eax
  loc_00E2150E: push edx
  loc_00E2150F: call __vbaObjSet
  loc_00E21511: mov edi, eax
  loc_00E21513: push 000124F8h
  loc_00E21518: mov ebx, [edi]
  loc_00E2151A: call [00401018h] ; __vbaStrI4
  loc_00E21520: mov edx, eax
  loc_00E21522: lea ecx, var_18
  loc_00E21525: call [00401228h] ; __vbaStrMove
  loc_00E2152B: push eax
  loc_00E2152C: push edi
  loc_00E2152D: call [ebx+00000054h]
  loc_00E21530: test eax, eax
  loc_00E21532: fnclex
  loc_00E21534: jge 00E21545h
  loc_00E21536: push 00000054h
  loc_00E21538: push 006E0410h
  loc_00E2153D: push edi
  loc_00E2153E: push eax
  loc_00E2153F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21545: lea ecx, var_18
  loc_00E21548: call [00401258h] ; __vbaFreeStr
  loc_00E2154E: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E21554: lea ecx, var_20
  loc_00E21557: call ebx
  loc_00E21559: mov eax, [00E3D060h]
  loc_00E2155E: test eax, eax
  loc_00E21560: jnz 00E21577h
  loc_00E21562: push 00E3D060h
  loc_00E21567: push 006D87C0h
  loc_00E2156C: call [004011A0h] ; __vbaNew2
  loc_00E21572: mov eax, [00E3D060h]
  loc_00E21577: mov ecx, [eax]
  loc_00E21579: push eax
  loc_00E2157A: call [ecx+0000042Ch]
  loc_00E21580: lea edx, var_20
  loc_00E21583: push eax
  loc_00E21584: push edx
  loc_00E21585: call __vbaObjSet
  loc_00E21587: mov edi, eax
  loc_00E21589: push 00000000h
  loc_00E2158B: push edi
  loc_00E2158C: mov eax, [edi]
  loc_00E2158E: call [eax+0000009Ch]
  loc_00E21594: test eax, eax
  loc_00E21596: fnclex
  loc_00E21598: jge 00E215ACh
  loc_00E2159A: push 0000009Ch
  loc_00E2159F: push 006DCAD0h
  loc_00E215A4: push edi
  loc_00E215A5: push eax
  loc_00E215A6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E215AC: lea ecx, var_20
  loc_00E215AF: call ebx
  loc_00E215B1: mov eax, [00E3D060h]
  loc_00E215B6: test eax, eax
  loc_00E215B8: jnz 00E215CFh
  loc_00E215BA: push 00E3D060h
  loc_00E215BF: push 006D87C0h
  loc_00E215C4: call [004011A0h] ; __vbaNew2
  loc_00E215CA: mov eax, [00E3D060h]
  loc_00E215CF: mov ecx, [eax]
  loc_00E215D1: push eax
  loc_00E215D2: call [ecx+000003FCh]
  loc_00E215D8: lea edx, var_20
  loc_00E215DB: push eax
  loc_00E215DC: push edx
  loc_00E215DD: call __vbaObjSet
  loc_00E215DF: mov edi, eax
  loc_00E215E1: push 000186A0h
  loc_00E215E6: mov ebx, [edi]
  loc_00E215E8: call [00401018h] ; __vbaStrI4
  loc_00E215EE: mov edx, eax
  loc_00E215F0: lea ecx, var_18
  loc_00E215F3: call [00401228h] ; __vbaStrMove
  loc_00E215F9: push eax
  loc_00E215FA: push edi
  loc_00E215FB: call [ebx+00000054h]
  loc_00E215FE: test eax, eax
  loc_00E21600: fnclex
  loc_00E21602: jge 00E21613h
  loc_00E21604: push 00000054h
  loc_00E21606: push 006E0410h
  loc_00E2160B: push edi
  loc_00E2160C: push eax
  loc_00E2160D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21613: lea ecx, var_18
  loc_00E21616: call [00401258h] ; __vbaFreeStr
  loc_00E2161C: lea ecx, var_20
  loc_00E2161F: call [00401254h] ; __vbaFreeObj
  loc_00E21625: mov eax, [00E3D060h]
  loc_00E2162A: test eax, eax
  loc_00E2162C: jnz 00E21643h
  loc_00E2162E: push 00E3D060h
  loc_00E21633: push 006D87C0h
  loc_00E21638: call [004011A0h] ; __vbaNew2
  loc_00E2163E: mov eax, [00E3D060h]
  loc_00E21643: mov ecx, [eax]
  loc_00E21645: push eax
  loc_00E21646: call [ecx+000003ECh]
  loc_00E2164C: lea edx, var_20
  loc_00E2164F: push eax
  loc_00E21650: push edx
  loc_00E21651: call __vbaObjSet
  loc_00E21653: mov edi, eax
  loc_00E21655: push 00000000h
  loc_00E21657: mov ebx, [edi]
  loc_00E21659: call [0040100Ch] ; __vbaStrI2
  loc_00E2165F: mov edx, eax
  loc_00E21661: lea ecx, var_18
  loc_00E21664: call [00401228h] ; __vbaStrMove
  loc_00E2166A: push eax
  loc_00E2166B: push edi
  loc_00E2166C: call [ebx+00000054h]
  loc_00E2166F: test eax, eax
  loc_00E21671: fnclex
  loc_00E21673: jge 00E21684h
  loc_00E21675: push 00000054h
  loc_00E21677: push 006E0410h
  loc_00E2167C: push edi
  loc_00E2167D: push eax
  loc_00E2167E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21684: lea ecx, var_18
  loc_00E21687: call [00401258h] ; __vbaFreeStr
  loc_00E2168D: lea ecx, var_20
  loc_00E21690: call [00401254h] ; __vbaFreeObj
  loc_00E21696: mov eax, [00E3D060h]
  loc_00E2169B: test eax, eax
  loc_00E2169D: jnz 00E216B4h
  loc_00E2169F: push 00E3D060h
  loc_00E216A4: push 006D87C0h
  loc_00E216A9: call [004011A0h] ; __vbaNew2
  loc_00E216AF: mov eax, [00E3D060h]
  loc_00E216B4: mov ecx, [eax]
  loc_00E216B6: push eax
  loc_00E216B7: call [ecx+000003DCh]
  loc_00E216BD: lea edx, var_20
  loc_00E216C0: push eax
  loc_00E216C1: push edx
  loc_00E216C2: call __vbaObjSet
  loc_00E216C4: mov edi, eax
  loc_00E216C6: push 000124F8h
  loc_00E216CB: mov ebx, [edi]
  loc_00E216CD: call [00401018h] ; __vbaStrI4
  loc_00E216D3: mov edx, eax
  loc_00E216D5: lea ecx, var_18
  loc_00E216D8: call [00401228h] ; __vbaStrMove
  loc_00E216DE: push eax
  loc_00E216DF: push edi
  loc_00E216E0: call [ebx+00000054h]
  loc_00E216E3: test eax, eax
  loc_00E216E5: fnclex
  loc_00E216E7: jge 00E216F8h
  loc_00E216E9: push 00000054h
  loc_00E216EB: push 006E0410h
  loc_00E216F0: push edi
  loc_00E216F1: push eax
  loc_00E216F2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E216F8: lea ecx, var_18
  loc_00E216FB: call [00401258h] ; __vbaFreeStr
  loc_00E21701: lea ecx, var_20
  loc_00E21704: call [00401254h] ; __vbaFreeObj
  loc_00E2170A: mov eax, [00E3D060h]
  loc_00E2170F: test eax, eax
  loc_00E21711: jnz 00E21728h
  loc_00E21713: push 00E3D060h
  loc_00E21718: push 006D87C0h
  loc_00E2171D: call [004011A0h] ; __vbaNew2
  loc_00E21723: mov eax, [00E3D060h]
  loc_00E21728: mov ecx, [eax]
  loc_00E2172A: push eax
  loc_00E2172B: call [ecx+000003CCh]
  loc_00E21731: lea edx, var_20
  loc_00E21734: push eax
  loc_00E21735: push edx
  loc_00E21736: call __vbaObjSet
  loc_00E21738: mov edi, eax
  loc_00E2173A: push 000186A0h
  loc_00E2173F: mov ebx, [edi]
  loc_00E21741: call [00401018h] ; __vbaStrI4
  loc_00E21747: mov edx, eax
  loc_00E21749: lea ecx, var_18
  loc_00E2174C: call [00401228h] ; __vbaStrMove
  loc_00E21752: push eax
  loc_00E21753: push edi
  loc_00E21754: call [ebx+00000054h]
  loc_00E21757: test eax, eax
  loc_00E21759: fnclex
  loc_00E2175B: jge 00E2176Ch
  loc_00E2175D: push 00000054h
  loc_00E2175F: push 006E0410h
  loc_00E21764: push edi
  loc_00E21765: push eax
  loc_00E21766: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2176C: lea ecx, var_18
  loc_00E2176F: call [00401258h] ; __vbaFreeStr
  loc_00E21775: lea ecx, var_20
  loc_00E21778: call [00401254h] ; __vbaFreeObj
  loc_00E2177E: mov eax, [00E3D060h]
  loc_00E21783: test eax, eax
  loc_00E21785: jnz 00E2179Ch
  loc_00E21787: push 00E3D060h
  loc_00E2178C: push 006D87C0h
  loc_00E21791: call [004011A0h] ; __vbaNew2
  loc_00E21797: mov eax, [00E3D060h]
  loc_00E2179C: mov ecx, [eax]
  loc_00E2179E: push eax
  loc_00E2179F: call [ecx+000003BCh]
  loc_00E217A5: lea edx, var_20
  loc_00E217A8: push eax
  loc_00E217A9: push edx
  loc_00E217AA: call __vbaObjSet
  loc_00E217AC: mov edi, eax
  loc_00E217AE: push 000186A0h
  loc_00E217B3: mov ebx, [edi]
  loc_00E217B5: call [00401018h] ; __vbaStrI4
  loc_00E217BB: mov edx, eax
  loc_00E217BD: lea ecx, var_18
  loc_00E217C0: call [00401228h] ; __vbaStrMove
  loc_00E217C6: push eax
  loc_00E217C7: push edi
  loc_00E217C8: call [ebx+00000054h]
  loc_00E217CB: test eax, eax
  loc_00E217CD: fnclex
  loc_00E217CF: jge 00E217E0h
  loc_00E217D1: push 00000054h
  loc_00E217D3: push 006E0410h
  loc_00E217D8: push edi
  loc_00E217D9: push eax
  loc_00E217DA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E217E0: lea ecx, var_18
  loc_00E217E3: call [00401258h] ; __vbaFreeStr
  loc_00E217E9: lea ecx, var_20
  loc_00E217EC: call [00401254h] ; __vbaFreeObj
  loc_00E217F2: mov eax, [00E3D060h]
  loc_00E217F7: test eax, eax
  loc_00E217F9: jnz 00E21810h
  loc_00E217FB: push 00E3D060h
  loc_00E21800: push 006D87C0h
  loc_00E21805: call [004011A0h] ; __vbaNew2
  loc_00E2180B: mov eax, [00E3D060h]
  loc_00E21810: mov ecx, [eax]
  loc_00E21812: push eax
  loc_00E21813: call [ecx+000003A8h]
  loc_00E21819: lea edx, var_20
  loc_00E2181C: push eax
  loc_00E2181D: push edx
  loc_00E2181E: call __vbaObjSet
  loc_00E21820: mov edi, eax
  loc_00E21822: push 00003A98h
  loc_00E21827: mov ebx, [edi]
  loc_00E21829: call [0040100Ch] ; __vbaStrI2
  loc_00E2182F: mov edx, eax
  loc_00E21831: lea ecx, var_18
  loc_00E21834: call [00401228h] ; __vbaStrMove
  loc_00E2183A: push eax
  loc_00E2183B: push edi
  loc_00E2183C: call [ebx+00000054h]
  loc_00E2183F: test eax, eax
  loc_00E21841: fnclex
  loc_00E21843: jge 00E21854h
  loc_00E21845: push 00000054h
  loc_00E21847: push 006E0410h
  loc_00E2184C: push edi
  loc_00E2184D: push eax
  loc_00E2184E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21854: lea ecx, var_18
  loc_00E21857: call [00401258h] ; __vbaFreeStr
  loc_00E2185D: lea ecx, var_20
  loc_00E21860: call [00401254h] ; __vbaFreeObj
  loc_00E21866: mov eax, [00E3D060h]
  loc_00E2186B: test eax, eax
  loc_00E2186D: jnz 00E21884h
  loc_00E2186F: push 00E3D060h
  loc_00E21874: push 006D87C0h
  loc_00E21879: call [004011A0h] ; __vbaNew2
  loc_00E2187F: mov eax, [00E3D060h]
  loc_00E21884: mov ecx, [eax]
  loc_00E21886: push eax
  loc_00E21887: call [ecx+00000398h]
  loc_00E2188D: lea edx, var_20
  loc_00E21890: push eax
  loc_00E21891: push edx
  loc_00E21892: call __vbaObjSet
  loc_00E21894: mov edi, eax
  loc_00E21896: push 00003A98h
  loc_00E2189B: mov ebx, [edi]
  loc_00E2189D: call [0040100Ch] ; __vbaStrI2
  loc_00E218A3: mov edx, eax
  loc_00E218A5: lea ecx, var_18
  loc_00E218A8: call [00401228h] ; __vbaStrMove
  loc_00E218AE: push eax
  loc_00E218AF: push edi
  loc_00E218B0: call [ebx+00000054h]
  loc_00E218B3: test eax, eax
  loc_00E218B5: fnclex
  loc_00E218B7: jge 00E21368h
  loc_00E218BD: jmp 00E21359h
  loc_00E218C2: test eax, eax
  loc_00E218C4: jnz 00E218DBh
  loc_00E218C6: push 00E3D060h
  loc_00E218CB: push 006D87C0h
  loc_00E218D0: call [004011A0h] ; __vbaNew2
  loc_00E218D6: mov eax, [00E3D060h]
  loc_00E218DB: mov ecx, [eax]
  loc_00E218DD: push eax
  loc_00E218DE: call [ecx+000004E4h]
  loc_00E218E4: lea edx, var_20
  loc_00E218E7: push eax
  loc_00E218E8: push edx
  loc_00E218E9: call __vbaObjSet
  loc_00E218EB: mov edi, eax
  loc_00E218ED: lea ecx, var_18
  loc_00E218F0: push ecx
  loc_00E218F1: push edi
  loc_00E218F2: mov eax, [edi]
  loc_00E218F4: call [eax+00000050h]
  loc_00E218F7: test eax, eax
  loc_00E218F9: fnclex
  loc_00E218FB: jge 00E2190Ch
  loc_00E218FD: push 00000050h
  loc_00E218FF: push 006E0410h
  loc_00E21904: push edi
  loc_00E21905: push eax
  loc_00E21906: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2190C: mov eax, [00E3D060h]
  loc_00E21911: test eax, eax
  loc_00E21913: jnz 00E2192Ah
  loc_00E21915: push 00E3D060h
  loc_00E2191A: push 006D87C0h
  loc_00E2191F: call [004011A0h] ; __vbaNew2
  loc_00E21925: mov eax, [00E3D060h]
  loc_00E2192A: mov edx, [eax]
  loc_00E2192C: push eax
  loc_00E2192D: call [edx+00000518h]
  loc_00E21933: push eax
  loc_00E21934: lea eax, var_24
  loc_00E21937: push eax
  loc_00E21938: call __vbaObjSet
  loc_00E2193A: mov edi, eax
  loc_00E2193C: lea edx, var_1C
  loc_00E2193F: push edx
  loc_00E21940: push edi
  loc_00E21941: mov ecx, [edi]
  loc_00E21943: call [ecx+00000050h]
  loc_00E21946: test eax, eax
  loc_00E21948: fnclex
  loc_00E2194A: jge 00E2195Bh
  loc_00E2194C: push 00000050h
  loc_00E2194E: push 006E0410h
  loc_00E21953: push edi
  loc_00E21954: push eax
  loc_00E21955: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2195B: mov eax, var_1C
  loc_00E2195E: push eax
  loc_00E2195F: push 006E06A8h ; "Farmasi"
  loc_00E21964: call ebx
  loc_00E21966: mov ecx, var_18
  loc_00E21969: mov edi, eax
  loc_00E2196B: neg edi
  loc_00E2196D: sbb edi, edi
  loc_00E2196F: push ecx
  loc_00E21970: inc edi
  loc_00E21971: push 006E0B24h ; "Laki - laki"
  loc_00E21976: neg edi
  loc_00E21978: call ebx
  loc_00E2197A: neg eax
  loc_00E2197C: sbb eax, eax
  loc_00E2197E: lea edx, var_1C
  loc_00E21981: inc eax
  loc_00E21982: push edx
  loc_00E21983: neg eax
  loc_00E21985: and edi, eax
  loc_00E21987: lea eax, var_18
  loc_00E2198A: push eax
  loc_00E2198B: push 00000002h
  loc_00E2198D: call [004011BCh] ; __vbaFreeStrList
  loc_00E21993: lea ecx, var_24
  loc_00E21996: lea edx, var_20
  loc_00E21999: push ecx
  loc_00E2199A: push edx
  loc_00E2199B: push 00000002h
  loc_00E2199D: call [00401048h] ; __vbaFreeObjList
  loc_00E219A3: mov eax, [00E3D060h]
  loc_00E219A8: add esp, 00000018h
  loc_00E219AB: test di, di
  loc_00E219AE: jz 00E2269Eh
  loc_00E219B4: test eax, eax
  loc_00E219B6: jnz 00E219CDh
  loc_00E219B8: push 00E3D060h
  loc_00E219BD: push 006D87C0h
  loc_00E219C2: call [004011A0h] ; __vbaNew2
  loc_00E219C8: mov eax, [00E3D060h]
  loc_00E219CD: mov ecx, [eax]
  loc_00E219CF: push eax
  loc_00E219D0: call [ecx+000004E0h]
  loc_00E219D6: lea edx, var_20
  loc_00E219D9: push eax
  loc_00E219DA: push edx
  loc_00E219DB: call __vbaObjSet
  loc_00E219DD: mov edi, eax
  loc_00E219DF: lea ecx, var_18
  loc_00E219E2: push ecx
  loc_00E219E3: push edi
  loc_00E219E4: mov eax, [edi]
  loc_00E219E6: call [eax+00000050h]
  loc_00E219E9: test eax, eax
  loc_00E219EB: fnclex
  loc_00E219ED: jge 00E219FEh
  loc_00E219EF: push 00000050h
  loc_00E219F1: push 006E0410h
  loc_00E219F6: push edi
  loc_00E219F7: push eax
  loc_00E219F8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E219FE: mov edx, var_18
  loc_00E21A01: push edx
  loc_00E21A02: push 006E05FCh ; "Katholik"
  loc_00E21A07: call ebx
  loc_00E21A09: mov edi, eax
  loc_00E21A0B: lea ecx, var_18
  loc_00E21A0E: neg edi
  loc_00E21A10: sbb edi, edi
  loc_00E21A12: inc edi
  loc_00E21A13: neg edi
  loc_00E21A15: call [00401258h] ; __vbaFreeStr
  loc_00E21A1B: lea ecx, var_20
  loc_00E21A1E: call [00401254h] ; __vbaFreeObj
  loc_00E21A24: mov eax, [00E3D060h]
  loc_00E21A29: test di, di
  loc_00E21A2C: jz 00E220F1h
  loc_00E21A32: test eax, eax
  loc_00E21A34: jnz 00E21A4Bh
  loc_00E21A36: push 00E3D060h
  loc_00E21A3B: push 006D87C0h
  loc_00E21A40: call [004011A0h] ; __vbaNew2
  loc_00E21A46: mov eax, [00E3D060h]
  loc_00E21A4B: mov ecx, [eax]
  loc_00E21A4D: push eax
  loc_00E21A4E: call [ecx+00000474h]
  loc_00E21A54: lea edx, var_20
  loc_00E21A57: push eax
  loc_00E21A58: push edx
  loc_00E21A59: call __vbaObjSet
  loc_00E21A5B: mov edi, eax
  loc_00E21A5D: push 000927C0h
  loc_00E21A62: mov ebx, [edi]
  loc_00E21A64: call [00401018h] ; __vbaStrI4
  loc_00E21A6A: mov edx, eax
  loc_00E21A6C: lea ecx, var_18
  loc_00E21A6F: call [00401228h] ; __vbaStrMove
  loc_00E21A75: push eax
  loc_00E21A76: push edi
  loc_00E21A77: call [ebx+00000054h]
  loc_00E21A7A: test eax, eax
  loc_00E21A7C: fnclex
  loc_00E21A7E: jge 00E21A8Fh
  loc_00E21A80: push 00000054h
  loc_00E21A82: push 006E0410h
  loc_00E21A87: push edi
  loc_00E21A88: push eax
  loc_00E21A89: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21A8F: lea ecx, var_18
  loc_00E21A92: call [00401258h] ; __vbaFreeStr
  loc_00E21A98: lea ecx, var_20
  loc_00E21A9B: call [00401254h] ; __vbaFreeObj
  loc_00E21AA1: mov eax, [00E3D060h]
  loc_00E21AA6: test eax, eax
  loc_00E21AA8: jnz 00E21ABFh
  loc_00E21AAA: push 00E3D060h
  loc_00E21AAF: push 006D87C0h
  loc_00E21AB4: call [004011A0h] ; __vbaNew2
  loc_00E21ABA: mov eax, [00E3D060h]
  loc_00E21ABF: mov ecx, [eax]
  loc_00E21AC1: push eax
  loc_00E21AC2: call [ecx+00000470h]
  loc_00E21AC8: lea edx, var_20
  loc_00E21ACB: push eax
  loc_00E21ACC: push edx
  loc_00E21ACD: call __vbaObjSet
  loc_00E21ACF: mov edi, eax
  loc_00E21AD1: push 0003A980h
  loc_00E21AD6: mov ebx, [edi]
  loc_00E21AD8: call [00401018h] ; __vbaStrI4
  loc_00E21ADE: mov edx, eax
  loc_00E21AE0: lea ecx, var_18
  loc_00E21AE3: call [00401228h] ; __vbaStrMove
  loc_00E21AE9: push eax
  loc_00E21AEA: push edi
  loc_00E21AEB: call [ebx+00000054h]
  loc_00E21AEE: test eax, eax
  loc_00E21AF0: fnclex
  loc_00E21AF2: jge 00E21B03h
  loc_00E21AF4: push 00000054h
  loc_00E21AF6: push 006E0410h
  loc_00E21AFB: push edi
  loc_00E21AFC: push eax
  loc_00E21AFD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21B03: lea ecx, var_18
  loc_00E21B06: call [00401258h] ; __vbaFreeStr
  loc_00E21B0C: lea ecx, var_20
  loc_00E21B0F: call [00401254h] ; __vbaFreeObj
  loc_00E21B15: mov eax, [00E3D060h]
  loc_00E21B1A: test eax, eax
  loc_00E21B1C: jnz 00E21B33h
  loc_00E21B1E: push 00E3D060h
  loc_00E21B23: push 006D87C0h
  loc_00E21B28: call [004011A0h] ; __vbaNew2
  loc_00E21B2E: mov eax, [00E3D060h]
  loc_00E21B33: mov ecx, [eax]
  loc_00E21B35: push eax
  loc_00E21B36: call [ecx+0000046Ch]
  loc_00E21B3C: lea edx, var_20
  loc_00E21B3F: push eax
  loc_00E21B40: push edx
  loc_00E21B41: call __vbaObjSet
  loc_00E21B43: mov edi, eax
  loc_00E21B45: push 00077A10h
  loc_00E21B4A: mov ebx, [edi]
  loc_00E21B4C: call [00401018h] ; __vbaStrI4
  loc_00E21B52: mov edx, eax
  loc_00E21B54: lea ecx, var_18
  loc_00E21B57: call [00401228h] ; __vbaStrMove
  loc_00E21B5D: push eax
  loc_00E21B5E: push edi
  loc_00E21B5F: call [ebx+00000054h]
  loc_00E21B62: test eax, eax
  loc_00E21B64: fnclex
  loc_00E21B66: jge 00E21B77h
  loc_00E21B68: push 00000054h
  loc_00E21B6A: push 006E0410h
  loc_00E21B6F: push edi
  loc_00E21B70: push eax
  loc_00E21B71: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21B77: lea ecx, var_18
  loc_00E21B7A: call [00401258h] ; __vbaFreeStr
  loc_00E21B80: lea ecx, var_20
  loc_00E21B83: call [00401254h] ; __vbaFreeObj
  loc_00E21B89: mov eax, [00E3D060h]
  loc_00E21B8E: test eax, eax
  loc_00E21B90: jnz 00E21BA7h
  loc_00E21B92: push 00E3D060h
  loc_00E21B97: push 006D87C0h
  loc_00E21B9C: call [004011A0h] ; __vbaNew2
  loc_00E21BA2: mov eax, [00E3D060h]
  loc_00E21BA7: mov ecx, [eax]
  loc_00E21BA9: push eax
  loc_00E21BAA: call [ecx+00000454h]
  loc_00E21BB0: lea edx, var_20
  loc_00E21BB3: push eax
  loc_00E21BB4: push edx
  loc_00E21BB5: call __vbaObjSet
  loc_00E21BB7: mov edi, eax
  loc_00E21BB9: push 000124F8h
  loc_00E21BBE: mov ebx, [edi]
  loc_00E21BC0: call [00401018h] ; __vbaStrI4
  loc_00E21BC6: mov edx, eax
  loc_00E21BC8: lea ecx, var_18
  loc_00E21BCB: call [00401228h] ; __vbaStrMove
  loc_00E21BD1: push eax
  loc_00E21BD2: push edi
  loc_00E21BD3: call [ebx+00000054h]
  loc_00E21BD6: test eax, eax
  loc_00E21BD8: fnclex
  loc_00E21BDA: jge 00E21BEBh
  loc_00E21BDC: push 00000054h
  loc_00E21BDE: push 006E0410h
  loc_00E21BE3: push edi
  loc_00E21BE4: push eax
  loc_00E21BE5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21BEB: lea ecx, var_18
  loc_00E21BEE: call [00401258h] ; __vbaFreeStr
  loc_00E21BF4: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E21BFA: lea ecx, var_20
  loc_00E21BFD: call ebx
  loc_00E21BFF: mov eax, [00E3D060h]
  loc_00E21C04: test eax, eax
  loc_00E21C06: jnz 00E21C1Dh
  loc_00E21C08: push 00E3D060h
  loc_00E21C0D: push 006D87C0h
  loc_00E21C12: call [004011A0h] ; __vbaNew2
  loc_00E21C18: mov eax, [00E3D060h]
  loc_00E21C1D: mov ecx, [eax]
  loc_00E21C1F: push eax
  loc_00E21C20: call [ecx+0000042Ch]
  loc_00E21C26: lea edx, var_20
  loc_00E21C29: push eax
  loc_00E21C2A: push edx
  loc_00E21C2B: call __vbaObjSet
  loc_00E21C2D: mov edi, eax
  loc_00E21C2F: push FFFFFFFFh
  loc_00E21C31: push edi
  loc_00E21C32: mov eax, [edi]
  loc_00E21C34: call [eax+0000009Ch]
  loc_00E21C3A: test eax, eax
  loc_00E21C3C: fnclex
  loc_00E21C3E: jge 00E21C52h
  loc_00E21C40: push 0000009Ch
  loc_00E21C45: push 006DCAD0h
  loc_00E21C4A: push edi
  loc_00E21C4B: push eax
  loc_00E21C4C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21C52: lea ecx, var_20
  loc_00E21C55: call ebx
  loc_00E21C57: mov eax, [00E3D060h]
  loc_00E21C5C: test eax, eax
  loc_00E21C5E: jnz 00E21C75h
  loc_00E21C60: push 00E3D060h
  loc_00E21C65: push 006D87C0h
  loc_00E21C6A: call [004011A0h] ; __vbaNew2
  loc_00E21C70: mov eax, [00E3D060h]
  loc_00E21C75: mov ecx, [eax]
  loc_00E21C77: push eax
  loc_00E21C78: call [ecx+00000440h]
  loc_00E21C7E: lea edx, var_20
  loc_00E21C81: push eax
  loc_00E21C82: push edx
  loc_00E21C83: call __vbaObjSet
  loc_00E21C85: mov edi, eax
  loc_00E21C87: push 0007A120h
  loc_00E21C8C: mov ebx, [edi]
  loc_00E21C8E: call [00401018h] ; __vbaStrI4
  loc_00E21C94: mov edx, eax
  loc_00E21C96: lea ecx, var_18
  loc_00E21C99: call [00401228h] ; __vbaStrMove
  loc_00E21C9F: push eax
  loc_00E21CA0: push edi
  loc_00E21CA1: call [ebx+00000054h]
  loc_00E21CA4: test eax, eax
  loc_00E21CA6: fnclex
  loc_00E21CA8: jge 00E21CB9h
  loc_00E21CAA: push 00000054h
  loc_00E21CAC: push 006E0410h
  loc_00E21CB1: push edi
  loc_00E21CB2: push eax
  loc_00E21CB3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21CB9: lea ecx, var_18
  loc_00E21CBC: call [00401258h] ; __vbaFreeStr
  loc_00E21CC2: lea ecx, var_20
  loc_00E21CC5: call [00401254h] ; __vbaFreeObj
  loc_00E21CCB: mov eax, [00E3D060h]
  loc_00E21CD0: test eax, eax
  loc_00E21CD2: jnz 00E21CE9h
  loc_00E21CD4: push 00E3D060h
  loc_00E21CD9: push 006D87C0h
  loc_00E21CDE: call [004011A0h] ; __vbaNew2
  loc_00E21CE4: mov eax, [00E3D060h]
  loc_00E21CE9: mov ecx, [eax]
  loc_00E21CEB: push eax
  loc_00E21CEC: call [ecx+00000434h]
  loc_00E21CF2: lea edx, var_20
  loc_00E21CF5: push eax
  loc_00E21CF6: push edx
  loc_00E21CF7: call __vbaObjSet
  loc_00E21CF9: mov edi, eax
  loc_00E21CFB: push 00030D40h
  loc_00E21D00: mov ebx, [edi]
  loc_00E21D02: call [00401018h] ; __vbaStrI4
  loc_00E21D08: mov edx, eax
  loc_00E21D0A: lea ecx, var_18
  loc_00E21D0D: call [00401228h] ; __vbaStrMove
  loc_00E21D13: push eax
  loc_00E21D14: push edi
  loc_00E21D15: call [ebx+00000054h]
  loc_00E21D18: test eax, eax
  loc_00E21D1A: fnclex
  loc_00E21D1C: jge 00E21D2Dh
  loc_00E21D1E: push 00000054h
  loc_00E21D20: push 006E0410h
  loc_00E21D25: push edi
  loc_00E21D26: push eax
  loc_00E21D27: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21D2D: lea ecx, var_18
  loc_00E21D30: call [00401258h] ; __vbaFreeStr
  loc_00E21D36: lea ecx, var_20
  loc_00E21D39: call [00401254h] ; __vbaFreeObj
  loc_00E21D3F: mov eax, [00E3D060h]
  loc_00E21D44: test eax, eax
  loc_00E21D46: jnz 00E21D5Dh
  loc_00E21D48: push 00E3D060h
  loc_00E21D4D: push 006D87C0h
  loc_00E21D52: call [004011A0h] ; __vbaNew2
  loc_00E21D58: mov eax, [00E3D060h]
  loc_00E21D5D: mov ecx, [eax]
  loc_00E21D5F: push eax
  loc_00E21D60: call [ecx+000003FCh]
  loc_00E21D66: lea edx, var_20
  loc_00E21D69: push eax
  loc_00E21D6A: push edx
  loc_00E21D6B: call __vbaObjSet
  loc_00E21D6D: mov edi, eax
  loc_00E21D6F: push 00027100h
  loc_00E21D74: mov ebx, [edi]
  loc_00E21D76: call [00401018h] ; __vbaStrI4
  loc_00E21D7C: mov edx, eax
  loc_00E21D7E: lea ecx, var_18
  loc_00E21D81: call [00401228h] ; __vbaStrMove
  loc_00E21D87: push eax
  loc_00E21D88: push edi
  loc_00E21D89: call [ebx+00000054h]
  loc_00E21D8C: test eax, eax
  loc_00E21D8E: fnclex
  loc_00E21D90: jge 00E21DA1h
  loc_00E21D92: push 00000054h
  loc_00E21D94: push 006E0410h
  loc_00E21D99: push edi
  loc_00E21D9A: push eax
  loc_00E21D9B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21DA1: lea ecx, var_18
  loc_00E21DA4: call [00401258h] ; __vbaFreeStr
  loc_00E21DAA: lea ecx, var_20
  loc_00E21DAD: call [00401254h] ; __vbaFreeObj
  loc_00E21DB3: mov eax, [00E3D060h]
  loc_00E21DB8: test eax, eax
  loc_00E21DBA: jnz 00E21DD1h
  loc_00E21DBC: push 00E3D060h
  loc_00E21DC1: push 006D87C0h
  loc_00E21DC6: call [004011A0h] ; __vbaNew2
  loc_00E21DCC: mov eax, [00E3D060h]
  loc_00E21DD1: mov ecx, [eax]
  loc_00E21DD3: push eax
  loc_00E21DD4: call [ecx+000003ECh]
  loc_00E21DDA: lea edx, var_20
  loc_00E21DDD: push eax
  loc_00E21DDE: push edx
  loc_00E21DDF: call __vbaObjSet
  loc_00E21DE1: mov edi, eax
  loc_00E21DE3: push 00035B60h
  loc_00E21DE8: mov ebx, [edi]
  loc_00E21DEA: call [00401018h] ; __vbaStrI4
  loc_00E21DF0: mov edx, eax
  loc_00E21DF2: lea ecx, var_18
  loc_00E21DF5: call [00401228h] ; __vbaStrMove
  loc_00E21DFB: push eax
  loc_00E21DFC: push edi
  loc_00E21DFD: call [ebx+00000054h]
  loc_00E21E00: test eax, eax
  loc_00E21E02: fnclex
  loc_00E21E04: jge 00E21E15h
  loc_00E21E06: push 00000054h
  loc_00E21E08: push 006E0410h
  loc_00E21E0D: push edi
  loc_00E21E0E: push eax
  loc_00E21E0F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21E15: lea ecx, var_18
  loc_00E21E18: call [00401258h] ; __vbaFreeStr
  loc_00E21E1E: lea ecx, var_20
  loc_00E21E21: call [00401254h] ; __vbaFreeObj
  loc_00E21E27: mov eax, [00E3D060h]
  loc_00E21E2C: test eax, eax
  loc_00E21E2E: jnz 00E21E45h
  loc_00E21E30: push 00E3D060h
  loc_00E21E35: push 006D87C0h
  loc_00E21E3A: call [004011A0h] ; __vbaNew2
  loc_00E21E40: mov eax, [00E3D060h]
  loc_00E21E45: mov ecx, [eax]
  loc_00E21E47: push eax
  loc_00E21E48: call [ecx+000003DCh]
  loc_00E21E4E: lea edx, var_20
  loc_00E21E51: push eax
  loc_00E21E52: push edx
  loc_00E21E53: call __vbaObjSet
  loc_00E21E55: mov edi, eax
  loc_00E21E57: push 0000FDE8h
  loc_00E21E5C: mov ebx, [edi]
  loc_00E21E5E: call [00401018h] ; __vbaStrI4
  loc_00E21E64: mov edx, eax
  loc_00E21E66: lea ecx, var_18
  loc_00E21E69: call [00401228h] ; __vbaStrMove
  loc_00E21E6F: push eax
  loc_00E21E70: push edi
  loc_00E21E71: call [ebx+00000054h]
  loc_00E21E74: test eax, eax
  loc_00E21E76: fnclex
  loc_00E21E78: jge 00E21E89h
  loc_00E21E7A: push 00000054h
  loc_00E21E7C: push 006E0410h
  loc_00E21E81: push edi
  loc_00E21E82: push eax
  loc_00E21E83: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21E89: lea ecx, var_18
  loc_00E21E8C: call [00401258h] ; __vbaFreeStr
  loc_00E21E92: lea ecx, var_20
  loc_00E21E95: call [00401254h] ; __vbaFreeObj
  loc_00E21E9B: mov eax, [00E3D060h]
  loc_00E21EA0: test eax, eax
  loc_00E21EA2: jnz 00E21EB9h
  loc_00E21EA4: push 00E3D060h
  loc_00E21EA9: push 006D87C0h
  loc_00E21EAE: call [004011A0h] ; __vbaNew2
  loc_00E21EB4: mov eax, [00E3D060h]
  loc_00E21EB9: mov ecx, [eax]
  loc_00E21EBB: push eax
  loc_00E21EBC: call [ecx+000003CCh]
  loc_00E21EC2: lea edx, var_20
  loc_00E21EC5: push eax
  loc_00E21EC6: push edx
  loc_00E21EC7: call __vbaObjSet
  loc_00E21EC9: mov edi, eax
  loc_00E21ECB: push 0001E848h
  loc_00E21ED0: mov ebx, [edi]
  loc_00E21ED2: call [00401018h] ; __vbaStrI4
  loc_00E21ED8: mov edx, eax
  loc_00E21EDA: lea ecx, var_18
  loc_00E21EDD: call [00401228h] ; __vbaStrMove
  loc_00E21EE3: push eax
  loc_00E21EE4: push edi
  loc_00E21EE5: call [ebx+00000054h]
  loc_00E21EE8: test eax, eax
  loc_00E21EEA: fnclex
  loc_00E21EEC: jge 00E21EFDh
  loc_00E21EEE: push 00000054h
  loc_00E21EF0: push 006E0410h
  loc_00E21EF5: push edi
  loc_00E21EF6: push eax
  loc_00E21EF7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21EFD: lea ecx, var_18
  loc_00E21F00: call [00401258h] ; __vbaFreeStr
  loc_00E21F06: lea ecx, var_20
  loc_00E21F09: call [00401254h] ; __vbaFreeObj
  loc_00E21F0F: mov eax, [00E3D060h]
  loc_00E21F14: test eax, eax
  loc_00E21F16: jnz 00E21F2Dh
  loc_00E21F18: push 00E3D060h
  loc_00E21F1D: push 006D87C0h
  loc_00E21F22: call [004011A0h] ; __vbaNew2
  loc_00E21F28: mov eax, [00E3D060h]
  loc_00E21F2D: mov ecx, [eax]
  loc_00E21F2F: push eax
  loc_00E21F30: call [ecx+000003BCh]
  loc_00E21F36: lea edx, var_20
  loc_00E21F39: push eax
  loc_00E21F3A: push edx
  loc_00E21F3B: call __vbaObjSet
  loc_00E21F3D: mov edi, eax
  loc_00E21F3F: push 00027100h
  loc_00E21F44: mov ebx, [edi]
  loc_00E21F46: call [00401018h] ; __vbaStrI4
  loc_00E21F4C: mov edx, eax
  loc_00E21F4E: lea ecx, var_18
  loc_00E21F51: call [00401228h] ; __vbaStrMove
  loc_00E21F57: push eax
  loc_00E21F58: push edi
  loc_00E21F59: call [ebx+00000054h]
  loc_00E21F5C: test eax, eax
  loc_00E21F5E: fnclex
  loc_00E21F60: jge 00E21F71h
  loc_00E21F62: push 00000054h
  loc_00E21F64: push 006E0410h
  loc_00E21F69: push edi
  loc_00E21F6A: push eax
  loc_00E21F6B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21F71: lea ecx, var_18
  loc_00E21F74: call [00401258h] ; __vbaFreeStr
  loc_00E21F7A: lea ecx, var_20
  loc_00E21F7D: call [00401254h] ; __vbaFreeObj
  loc_00E21F83: mov eax, [00E3D060h]
  loc_00E21F88: test eax, eax
  loc_00E21F8A: jnz 00E21FA1h
  loc_00E21F8C: push 00E3D060h
  loc_00E21F91: push 006D87C0h
  loc_00E21F96: call [004011A0h] ; __vbaNew2
  loc_00E21F9C: mov eax, [00E3D060h]
  loc_00E21FA1: mov ecx, [eax]
  loc_00E21FA3: push eax
  loc_00E21FA4: call [ecx+000003A8h]
  loc_00E21FAA: lea edx, var_20
  loc_00E21FAD: push eax
  loc_00E21FAE: push edx
  loc_00E21FAF: call __vbaObjSet
  loc_00E21FB1: mov edi, eax
  loc_00E21FB3: push 00003A98h
  loc_00E21FB8: mov ebx, [edi]
  loc_00E21FBA: call [0040100Ch] ; __vbaStrI2
  loc_00E21FC0: mov edx, eax
  loc_00E21FC2: lea ecx, var_18
  loc_00E21FC5: call [00401228h] ; __vbaStrMove
  loc_00E21FCB: push eax
  loc_00E21FCC: push edi
  loc_00E21FCD: call [ebx+00000054h]
  loc_00E21FD0: test eax, eax
  loc_00E21FD2: fnclex
  loc_00E21FD4: jge 00E21FE5h
  loc_00E21FD6: push 00000054h
  loc_00E21FD8: push 006E0410h
  loc_00E21FDD: push edi
  loc_00E21FDE: push eax
  loc_00E21FDF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E21FE5: lea ecx, var_18
  loc_00E21FE8: call [00401258h] ; __vbaFreeStr
  loc_00E21FEE: lea ecx, var_20
  loc_00E21FF1: call [00401254h] ; __vbaFreeObj
  loc_00E21FF7: mov eax, [00E3D060h]
  loc_00E21FFC: test eax, eax
  loc_00E21FFE: jnz 00E22015h
  loc_00E22000: push 00E3D060h
  loc_00E22005: push 006D87C0h
  loc_00E2200A: call [004011A0h] ; __vbaNew2
  loc_00E22010: mov eax, [00E3D060h]
  loc_00E22015: mov ecx, [eax]
  loc_00E22017: push eax
  loc_00E22018: call [ecx+00000398h]
  loc_00E2201E: lea edx, var_20
  loc_00E22021: push eax
  loc_00E22022: push edx
  loc_00E22023: call __vbaObjSet
  loc_00E22025: mov edi, eax
  loc_00E22027: push 00003A98h
  loc_00E2202C: mov ebx, [edi]
  loc_00E2202E: call [0040100Ch] ; __vbaStrI2
  loc_00E22034: mov edx, eax
  loc_00E22036: lea ecx, var_18
  loc_00E22039: call [00401228h] ; __vbaStrMove
  loc_00E2203F: push eax
  loc_00E22040: push edi
  loc_00E22041: call [ebx+00000054h]
  loc_00E22044: test eax, eax
  loc_00E22046: fnclex
  loc_00E22048: jge 00E22059h
  loc_00E2204A: push 00000054h
  loc_00E2204C: push 006E0410h
  loc_00E22051: push edi
  loc_00E22052: push eax
  loc_00E22053: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22059: lea ecx, var_18
  loc_00E2205C: call [00401258h] ; __vbaFreeStr
  loc_00E22062: lea ecx, var_20
  loc_00E22065: call [00401254h] ; __vbaFreeObj
  loc_00E2206B: mov eax, [00E3D060h]
  loc_00E22070: test eax, eax
  loc_00E22072: jnz 00E22089h
  loc_00E22074: push 00E3D060h
  loc_00E22079: push 006D87C0h
  loc_00E2207E: call [004011A0h] ; __vbaNew2
  loc_00E22084: mov eax, [00E3D060h]
  loc_00E22089: mov ecx, [eax]
  loc_00E2208B: push eax
  loc_00E2208C: call [ecx+00000388h]
  loc_00E22092: lea edx, var_20
  loc_00E22095: push eax
  loc_00E22096: push edx
  loc_00E22097: call __vbaObjSet
  loc_00E22099: mov edi, eax
  loc_00E2209B: push 00001388h
  loc_00E220A0: mov ebx, [edi]
  loc_00E220A2: call [0040100Ch] ; __vbaStrI2
  loc_00E220A8: mov edx, eax
  loc_00E220AA: lea ecx, var_18
  loc_00E220AD: call [00401228h] ; __vbaStrMove
  loc_00E220B3: push eax
  loc_00E220B4: push edi
  loc_00E220B5: call [ebx+00000054h]
  loc_00E220B8: test eax, eax
  loc_00E220BA: fnclex
  loc_00E220BC: jge 00E220CDh
  loc_00E220BE: push 00000054h
  loc_00E220C0: push 006E0410h
  loc_00E220C5: push edi
  loc_00E220C6: push eax
  loc_00E220C7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E220CD: lea ecx, var_18
  loc_00E220D0: call [00401258h] ; __vbaFreeStr
  loc_00E220D6: lea ecx, var_20
  loc_00E220D9: call [00401254h] ; __vbaFreeObj
  loc_00E220DF: mov eax, [00E3D060h]
  loc_00E220E4: test eax, eax
  loc_00E220E6: jnz 00E20072h
  loc_00E220EC: jmp 00E2005Dh
  loc_00E220F1: test eax, eax
  loc_00E220F3: jnz 00E2210Ah
  loc_00E220F5: push 00E3D060h
  loc_00E220FA: push 006D87C0h
  loc_00E220FF: call [004011A0h] ; __vbaNew2
  loc_00E22105: mov eax, [00E3D060h]
  loc_00E2210A: mov ecx, [eax]
  loc_00E2210C: push eax
  loc_00E2210D: call [ecx+00000474h]
  loc_00E22113: lea edx, var_20
  loc_00E22116: push eax
  loc_00E22117: push edx
  loc_00E22118: call __vbaObjSet
  loc_00E2211A: mov edi, eax
  loc_00E2211C: push 000AAE60h
  loc_00E22121: mov ebx, [edi]
  loc_00E22123: call [00401018h] ; __vbaStrI4
  loc_00E22129: mov edx, eax
  loc_00E2212B: lea ecx, var_18
  loc_00E2212E: call [00401228h] ; __vbaStrMove
  loc_00E22134: push eax
  loc_00E22135: push edi
  loc_00E22136: call [ebx+00000054h]
  loc_00E22139: test eax, eax
  loc_00E2213B: fnclex
  loc_00E2213D: jge 00E2214Eh
  loc_00E2213F: push 00000054h
  loc_00E22141: push 006E0410h
  loc_00E22146: push edi
  loc_00E22147: push eax
  loc_00E22148: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2214E: lea ecx, var_18
  loc_00E22151: call [00401258h] ; __vbaFreeStr
  loc_00E22157: lea ecx, var_20
  loc_00E2215A: call [00401254h] ; __vbaFreeObj
  loc_00E22160: mov eax, [00E3D060h]
  loc_00E22165: test eax, eax
  loc_00E22167: jnz 00E2217Eh
  loc_00E22169: push 00E3D060h
  loc_00E2216E: push 006D87C0h
  loc_00E22173: call [004011A0h] ; __vbaNew2
  loc_00E22179: mov eax, [00E3D060h]
  loc_00E2217E: mov ecx, [eax]
  loc_00E22180: push eax
  loc_00E22181: call [ecx+00000470h]
  loc_00E22187: lea edx, var_20
  loc_00E2218A: push eax
  loc_00E2218B: push edx
  loc_00E2218C: call __vbaObjSet
  loc_00E2218E: mov edi, eax
  loc_00E22190: push 0003A980h
  loc_00E22195: mov ebx, [edi]
  loc_00E22197: call [00401018h] ; __vbaStrI4
  loc_00E2219D: mov edx, eax
  loc_00E2219F: lea ecx, var_18
  loc_00E221A2: call [00401228h] ; __vbaStrMove
  loc_00E221A8: push eax
  loc_00E221A9: push edi
  loc_00E221AA: call [ebx+00000054h]
  loc_00E221AD: test eax, eax
  loc_00E221AF: fnclex
  loc_00E221B1: jge 00E221C2h
  loc_00E221B3: push 00000054h
  loc_00E221B5: push 006E0410h
  loc_00E221BA: push edi
  loc_00E221BB: push eax
  loc_00E221BC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E221C2: lea ecx, var_18
  loc_00E221C5: call [00401258h] ; __vbaFreeStr
  loc_00E221CB: lea ecx, var_20
  loc_00E221CE: call [00401254h] ; __vbaFreeObj
  loc_00E221D4: mov eax, [00E3D060h]
  loc_00E221D9: test eax, eax
  loc_00E221DB: jnz 00E221F2h
  loc_00E221DD: push 00E3D060h
  loc_00E221E2: push 006D87C0h
  loc_00E221E7: call [004011A0h] ; __vbaNew2
  loc_00E221ED: mov eax, [00E3D060h]
  loc_00E221F2: mov ecx, [eax]
  loc_00E221F4: push eax
  loc_00E221F5: call [ecx+0000046Ch]
  loc_00E221FB: lea edx, var_20
  loc_00E221FE: push eax
  loc_00E221FF: push edx
  loc_00E22200: call __vbaObjSet
  loc_00E22202: mov edi, eax
  loc_00E22204: push 00077A10h
  loc_00E22209: mov ebx, [edi]
  loc_00E2220B: call [00401018h] ; __vbaStrI4
  loc_00E22211: mov edx, eax
  loc_00E22213: lea ecx, var_18
  loc_00E22216: call [00401228h] ; __vbaStrMove
  loc_00E2221C: push eax
  loc_00E2221D: push edi
  loc_00E2221E: call [ebx+00000054h]
  loc_00E22221: test eax, eax
  loc_00E22223: fnclex
  loc_00E22225: jge 00E22236h
  loc_00E22227: push 00000054h
  loc_00E22229: push 006E0410h
  loc_00E2222E: push edi
  loc_00E2222F: push eax
  loc_00E22230: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22236: lea ecx, var_18
  loc_00E22239: call [00401258h] ; __vbaFreeStr
  loc_00E2223F: lea ecx, var_20
  loc_00E22242: call [00401254h] ; __vbaFreeObj
  loc_00E22248: mov eax, [00E3D060h]
  loc_00E2224D: test eax, eax
  loc_00E2224F: jnz 00E22266h
  loc_00E22251: push 00E3D060h
  loc_00E22256: push 006D87C0h
  loc_00E2225B: call [004011A0h] ; __vbaNew2
  loc_00E22261: mov eax, [00E3D060h]
  loc_00E22266: mov ecx, [eax]
  loc_00E22268: push eax
  loc_00E22269: call [ecx+00000454h]
  loc_00E2226F: lea edx, var_20
  loc_00E22272: push eax
  loc_00E22273: push edx
  loc_00E22274: call __vbaObjSet
  loc_00E22276: mov edi, eax
  loc_00E22278: push 000124F8h
  loc_00E2227D: mov ebx, [edi]
  loc_00E2227F: call [00401018h] ; __vbaStrI4
  loc_00E22285: mov edx, eax
  loc_00E22287: lea ecx, var_18
  loc_00E2228A: call [00401228h] ; __vbaStrMove
  loc_00E22290: push eax
  loc_00E22291: push edi
  loc_00E22292: call [ebx+00000054h]
  loc_00E22295: test eax, eax
  loc_00E22297: fnclex
  loc_00E22299: jge 00E222AAh
  loc_00E2229B: push 00000054h
  loc_00E2229D: push 006E0410h
  loc_00E222A2: push edi
  loc_00E222A3: push eax
  loc_00E222A4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E222AA: lea ecx, var_18
  loc_00E222AD: call [00401258h] ; __vbaFreeStr
  loc_00E222B3: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E222B9: lea ecx, var_20
  loc_00E222BC: call ebx
  loc_00E222BE: mov eax, [00E3D060h]
  loc_00E222C3: test eax, eax
  loc_00E222C5: jnz 00E222DCh
  loc_00E222C7: push 00E3D060h
  loc_00E222CC: push 006D87C0h
  loc_00E222D1: call [004011A0h] ; __vbaNew2
  loc_00E222D7: mov eax, [00E3D060h]
  loc_00E222DC: mov ecx, [eax]
  loc_00E222DE: push eax
  loc_00E222DF: call [ecx+0000042Ch]
  loc_00E222E5: lea edx, var_20
  loc_00E222E8: push eax
  loc_00E222E9: push edx
  loc_00E222EA: call __vbaObjSet
  loc_00E222EC: mov edi, eax
  loc_00E222EE: push FFFFFFFFh
  loc_00E222F0: push edi
  loc_00E222F1: mov eax, [edi]
  loc_00E222F3: call [eax+0000009Ch]
  loc_00E222F9: test eax, eax
  loc_00E222FB: fnclex
  loc_00E222FD: jge 00E22311h
  loc_00E222FF: push 0000009Ch
  loc_00E22304: push 006DCAD0h
  loc_00E22309: push edi
  loc_00E2230A: push eax
  loc_00E2230B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22311: lea ecx, var_20
  loc_00E22314: call ebx
  loc_00E22316: mov eax, [00E3D060h]
  loc_00E2231B: test eax, eax
  loc_00E2231D: jnz 00E22334h
  loc_00E2231F: push 00E3D060h
  loc_00E22324: push 006D87C0h
  loc_00E22329: call [004011A0h] ; __vbaNew2
  loc_00E2232F: mov eax, [00E3D060h]
  loc_00E22334: mov ecx, [eax]
  loc_00E22336: push eax
  loc_00E22337: call [ecx+00000434h]
  loc_00E2233D: lea edx, var_20
  loc_00E22340: push eax
  loc_00E22341: push edx
  loc_00E22342: call __vbaObjSet
  loc_00E22344: mov edi, eax
  loc_00E22346: push 00030D40h
  loc_00E2234B: mov ebx, [edi]
  loc_00E2234D: call [00401018h] ; __vbaStrI4
  loc_00E22353: mov edx, eax
  loc_00E22355: lea ecx, var_18
  loc_00E22358: call [00401228h] ; __vbaStrMove
  loc_00E2235E: push eax
  loc_00E2235F: push edi
  loc_00E22360: call [ebx+00000054h]
  loc_00E22363: test eax, eax
  loc_00E22365: fnclex
  loc_00E22367: jge 00E22378h
  loc_00E22369: push 00000054h
  loc_00E2236B: push 006E0410h
  loc_00E22370: push edi
  loc_00E22371: push eax
  loc_00E22372: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22378: lea ecx, var_18
  loc_00E2237B: call [00401258h] ; __vbaFreeStr
  loc_00E22381: lea ecx, var_20
  loc_00E22384: call [00401254h] ; __vbaFreeObj
  loc_00E2238A: mov eax, [00E3D060h]
  loc_00E2238F: test eax, eax
  loc_00E22391: jnz 00E223A8h
  loc_00E22393: push 00E3D060h
  loc_00E22398: push 006D87C0h
  loc_00E2239D: call [004011A0h] ; __vbaNew2
  loc_00E223A3: mov eax, [00E3D060h]
  loc_00E223A8: mov ecx, [eax]
  loc_00E223AA: push eax
  loc_00E223AB: call [ecx+00000440h]
  loc_00E223B1: lea edx, var_20
  loc_00E223B4: push eax
  loc_00E223B5: push edx
  loc_00E223B6: call __vbaObjSet
  loc_00E223B8: mov edi, eax
  loc_00E223BA: push 0007A120h
  loc_00E223BF: mov ebx, [edi]
  loc_00E223C1: call [00401018h] ; __vbaStrI4
  loc_00E223C7: mov edx, eax
  loc_00E223C9: lea ecx, var_18
  loc_00E223CC: call [00401228h] ; __vbaStrMove
  loc_00E223D2: push eax
  loc_00E223D3: push edi
  loc_00E223D4: call [ebx+00000054h]
  loc_00E223D7: test eax, eax
  loc_00E223D9: fnclex
  loc_00E223DB: jge 00E223ECh
  loc_00E223DD: push 00000054h
  loc_00E223DF: push 006E0410h
  loc_00E223E4: push edi
  loc_00E223E5: push eax
  loc_00E223E6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E223EC: lea ecx, var_18
  loc_00E223EF: call [00401258h] ; __vbaFreeStr
  loc_00E223F5: lea ecx, var_20
  loc_00E223F8: call [00401254h] ; __vbaFreeObj
  loc_00E223FE: mov eax, [00E3D060h]
  loc_00E22403: test eax, eax
  loc_00E22405: jnz 00E2241Ch
  loc_00E22407: push 00E3D060h
  loc_00E2240C: push 006D87C0h
  loc_00E22411: call [004011A0h] ; __vbaNew2
  loc_00E22417: mov eax, [00E3D060h]
  loc_00E2241C: mov ecx, [eax]
  loc_00E2241E: push eax
  loc_00E2241F: call [ecx+000003FCh]
  loc_00E22425: lea edx, var_20
  loc_00E22428: push eax
  loc_00E22429: push edx
  loc_00E2242A: call __vbaObjSet
  loc_00E2242C: mov edi, eax
  loc_00E2242E: push 00027100h
  loc_00E22433: mov ebx, [edi]
  loc_00E22435: call [00401018h] ; __vbaStrI4
  loc_00E2243B: mov edx, eax
  loc_00E2243D: lea ecx, var_18
  loc_00E22440: call [00401228h] ; __vbaStrMove
  loc_00E22446: push eax
  loc_00E22447: push edi
  loc_00E22448: call [ebx+00000054h]
  loc_00E2244B: test eax, eax
  loc_00E2244D: fnclex
  loc_00E2244F: jge 00E22460h
  loc_00E22451: push 00000054h
  loc_00E22453: push 006E0410h
  loc_00E22458: push edi
  loc_00E22459: push eax
  loc_00E2245A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22460: lea ecx, var_18
  loc_00E22463: call [00401258h] ; __vbaFreeStr
  loc_00E22469: lea ecx, var_20
  loc_00E2246C: call [00401254h] ; __vbaFreeObj
  loc_00E22472: mov eax, [00E3D060h]
  loc_00E22477: test eax, eax
  loc_00E22479: jnz 00E22490h
  loc_00E2247B: push 00E3D060h
  loc_00E22480: push 006D87C0h
  loc_00E22485: call [004011A0h] ; __vbaNew2
  loc_00E2248B: mov eax, [00E3D060h]
  loc_00E22490: mov ecx, [eax]
  loc_00E22492: push eax
  loc_00E22493: call [ecx+000003ECh]
  loc_00E22499: lea edx, var_20
  loc_00E2249C: push eax
  loc_00E2249D: push edx
  loc_00E2249E: call __vbaObjSet
  loc_00E224A0: mov edi, eax
  loc_00E224A2: push 00035B60h
  loc_00E224A7: mov ebx, [edi]
  loc_00E224A9: call [00401018h] ; __vbaStrI4
  loc_00E224AF: mov edx, eax
  loc_00E224B1: lea ecx, var_18
  loc_00E224B4: call [00401228h] ; __vbaStrMove
  loc_00E224BA: push eax
  loc_00E224BB: push edi
  loc_00E224BC: call [ebx+00000054h]
  loc_00E224BF: test eax, eax
  loc_00E224C1: fnclex
  loc_00E224C3: jge 00E224D4h
  loc_00E224C5: push 00000054h
  loc_00E224C7: push 006E0410h
  loc_00E224CC: push edi
  loc_00E224CD: push eax
  loc_00E224CE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E224D4: lea ecx, var_18
  loc_00E224D7: call [00401258h] ; __vbaFreeStr
  loc_00E224DD: lea ecx, var_20
  loc_00E224E0: call [00401254h] ; __vbaFreeObj
  loc_00E224E6: mov eax, [00E3D060h]
  loc_00E224EB: test eax, eax
  loc_00E224ED: jnz 00E22504h
  loc_00E224EF: push 00E3D060h
  loc_00E224F4: push 006D87C0h
  loc_00E224F9: call [004011A0h] ; __vbaNew2
  loc_00E224FF: mov eax, [00E3D060h]
  loc_00E22504: mov ecx, [eax]
  loc_00E22506: push eax
  loc_00E22507: call [ecx+000003DCh]
  loc_00E2250D: lea edx, var_20
  loc_00E22510: push eax
  loc_00E22511: push edx
  loc_00E22512: call __vbaObjSet
  loc_00E22514: mov edi, eax
  loc_00E22516: push 0000FDE8h
  loc_00E2251B: mov ebx, [edi]
  loc_00E2251D: call [00401018h] ; __vbaStrI4
  loc_00E22523: mov edx, eax
  loc_00E22525: lea ecx, var_18
  loc_00E22528: call [00401228h] ; __vbaStrMove
  loc_00E2252E: push eax
  loc_00E2252F: push edi
  loc_00E22530: call [ebx+00000054h]
  loc_00E22533: test eax, eax
  loc_00E22535: fnclex
  loc_00E22537: jge 00E22548h
  loc_00E22539: push 00000054h
  loc_00E2253B: push 006E0410h
  loc_00E22540: push edi
  loc_00E22541: push eax
  loc_00E22542: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22548: lea ecx, var_18
  loc_00E2254B: call [00401258h] ; __vbaFreeStr
  loc_00E22551: lea ecx, var_20
  loc_00E22554: call [00401254h] ; __vbaFreeObj
  loc_00E2255A: mov eax, [00E3D060h]
  loc_00E2255F: test eax, eax
  loc_00E22561: jnz 00E22578h
  loc_00E22563: push 00E3D060h
  loc_00E22568: push 006D87C0h
  loc_00E2256D: call [004011A0h] ; __vbaNew2
  loc_00E22573: mov eax, [00E3D060h]
  loc_00E22578: mov ecx, [eax]
  loc_00E2257A: push eax
  loc_00E2257B: call [ecx+000003CCh]
  loc_00E22581: lea edx, var_20
  loc_00E22584: push eax
  loc_00E22585: push edx
  loc_00E22586: call __vbaObjSet
  loc_00E22588: mov edi, eax
  loc_00E2258A: push 0001E848h
  loc_00E2258F: mov ebx, [edi]
  loc_00E22591: call [00401018h] ; __vbaStrI4
  loc_00E22597: mov edx, eax
  loc_00E22599: lea ecx, var_18
  loc_00E2259C: call [00401228h] ; __vbaStrMove
  loc_00E225A2: push eax
  loc_00E225A3: push edi
  loc_00E225A4: call [ebx+00000054h]
  loc_00E225A7: test eax, eax
  loc_00E225A9: fnclex
  loc_00E225AB: jge 00E225BCh
  loc_00E225AD: push 00000054h
  loc_00E225AF: push 006E0410h
  loc_00E225B4: push edi
  loc_00E225B5: push eax
  loc_00E225B6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E225BC: lea ecx, var_18
  loc_00E225BF: call [00401258h] ; __vbaFreeStr
  loc_00E225C5: lea ecx, var_20
  loc_00E225C8: call [00401254h] ; __vbaFreeObj
  loc_00E225CE: mov eax, [00E3D060h]
  loc_00E225D3: test eax, eax
  loc_00E225D5: jnz 00E225ECh
  loc_00E225D7: push 00E3D060h
  loc_00E225DC: push 006D87C0h
  loc_00E225E1: call [004011A0h] ; __vbaNew2
  loc_00E225E7: mov eax, [00E3D060h]
  loc_00E225EC: mov ecx, [eax]
  loc_00E225EE: push eax
  loc_00E225EF: call [ecx+000003BCh]
  loc_00E225F5: lea edx, var_20
  loc_00E225F8: push eax
  loc_00E225F9: push edx
  loc_00E225FA: call __vbaObjSet
  loc_00E225FC: mov edi, eax
  loc_00E225FE: push 00027100h
  loc_00E22603: mov ebx, [edi]
  loc_00E22605: call [00401018h] ; __vbaStrI4
  loc_00E2260B: mov edx, eax
  loc_00E2260D: lea ecx, var_18
  loc_00E22610: call [00401228h] ; __vbaStrMove
  loc_00E22616: push eax
  loc_00E22617: push edi
  loc_00E22618: call [ebx+00000054h]
  loc_00E2261B: test eax, eax
  loc_00E2261D: fnclex
  loc_00E2261F: jge 00E22630h
  loc_00E22621: push 00000054h
  loc_00E22623: push 006E0410h
  loc_00E22628: push edi
  loc_00E22629: push eax
  loc_00E2262A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22630: lea ecx, var_18
  loc_00E22633: call [00401258h] ; __vbaFreeStr
  loc_00E22639: lea ecx, var_20
  loc_00E2263C: call [00401254h] ; __vbaFreeObj
  loc_00E22642: mov eax, [00E3D060h]
  loc_00E22647: test eax, eax
  loc_00E22649: jnz 00E22660h
  loc_00E2264B: push 00E3D060h
  loc_00E22650: push 006D87C0h
  loc_00E22655: call [004011A0h] ; __vbaNew2
  loc_00E2265B: mov eax, [00E3D060h]
  loc_00E22660: mov ecx, [eax]
  loc_00E22662: push eax
  loc_00E22663: call [ecx+000003A8h]
  loc_00E22669: lea edx, var_20
  loc_00E2266C: push eax
  loc_00E2266D: push edx
  loc_00E2266E: call __vbaObjSet
  loc_00E22670: mov edi, eax
  loc_00E22672: push 00003A98h
  loc_00E22677: mov ebx, [edi]
  loc_00E22679: call [0040100Ch] ; __vbaStrI2
  loc_00E2267F: mov edx, eax
  loc_00E22681: lea ecx, var_18
  loc_00E22684: call [00401228h] ; __vbaStrMove
  loc_00E2268A: push eax
  loc_00E2268B: push edi
  loc_00E2268C: call [ebx+00000054h]
  loc_00E2268F: test eax, eax
  loc_00E22691: fnclex
  loc_00E22693: jge 00E206EBh
  loc_00E22699: jmp 00E206DCh
  loc_00E2269E: test eax, eax
  loc_00E226A0: jnz 00E226B7h
  loc_00E226A2: push 00E3D060h
  loc_00E226A7: push 006D87C0h
  loc_00E226AC: call [004011A0h] ; __vbaNew2
  loc_00E226B2: mov eax, [00E3D060h]
  loc_00E226B7: mov ecx, [eax]
  loc_00E226B9: push eax
  loc_00E226BA: call [ecx+000004E4h]
  loc_00E226C0: lea edx, var_20
  loc_00E226C3: push eax
  loc_00E226C4: push edx
  loc_00E226C5: call __vbaObjSet
  loc_00E226C7: mov edi, eax
  loc_00E226C9: lea ecx, var_18
  loc_00E226CC: push ecx
  loc_00E226CD: push edi
  loc_00E226CE: mov eax, [edi]
  loc_00E226D0: call [eax+00000050h]
  loc_00E226D3: test eax, eax
  loc_00E226D5: fnclex
  loc_00E226D7: jge 00E226E8h
  loc_00E226D9: push 00000050h
  loc_00E226DB: push 006E0410h
  loc_00E226E0: push edi
  loc_00E226E1: push eax
  loc_00E226E2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E226E8: mov eax, [00E3D060h]
  loc_00E226ED: test eax, eax
  loc_00E226EF: jnz 00E22706h
  loc_00E226F1: push 00E3D060h
  loc_00E226F6: push 006D87C0h
  loc_00E226FB: call [004011A0h] ; __vbaNew2
  loc_00E22701: mov eax, [00E3D060h]
  loc_00E22706: mov edx, [eax]
  loc_00E22708: push eax
  loc_00E22709: call [edx+00000518h]
  loc_00E2270F: push eax
  loc_00E22710: lea eax, var_24
  loc_00E22713: push eax
  loc_00E22714: call __vbaObjSet
  loc_00E22716: mov edi, eax
  loc_00E22718: lea edx, var_1C
  loc_00E2271B: push edx
  loc_00E2271C: push edi
  loc_00E2271D: mov ecx, [edi]
  loc_00E2271F: call [ecx+00000050h]
  loc_00E22722: test eax, eax
  loc_00E22724: fnclex
  loc_00E22726: jge 00E22737h
  loc_00E22728: push 00000050h
  loc_00E2272A: push 006E0410h
  loc_00E2272F: push edi
  loc_00E22730: push eax
  loc_00E22731: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22737: mov eax, var_1C
  loc_00E2273A: push eax
  loc_00E2273B: push 006E06A8h ; "Farmasi"
  loc_00E22740: call ebx
  loc_00E22742: mov ecx, var_18
  loc_00E22745: mov edi, eax
  loc_00E22747: neg edi
  loc_00E22749: sbb edi, edi
  loc_00E2274B: push ecx
  loc_00E2274C: inc edi
  loc_00E2274D: push 006E0B40h ; "Perempuan"
  loc_00E22752: neg edi
  loc_00E22754: call ebx
  loc_00E22756: neg eax
  loc_00E22758: sbb eax, eax
  loc_00E2275A: lea edx, var_1C
  loc_00E2275D: inc eax
  loc_00E2275E: push edx
  loc_00E2275F: neg eax
  loc_00E22761: and edi, eax
  loc_00E22763: lea eax, var_18
  loc_00E22766: push eax
  loc_00E22767: push 00000002h
  loc_00E22769: call [004011BCh] ; __vbaFreeStrList
  loc_00E2276F: lea ecx, var_24
  loc_00E22772: lea edx, var_20
  loc_00E22775: push ecx
  loc_00E22776: push edx
  loc_00E22777: push 00000002h
  loc_00E22779: call [00401048h] ; __vbaFreeObjList
  loc_00E2277F: add esp, 00000018h
  loc_00E22782: test di, di
  loc_00E22785: jz 00E2346Fh
  loc_00E2278B: mov eax, [00E3D060h]
  loc_00E22790: test eax, eax
  loc_00E22792: jnz 00E227A9h
  loc_00E22794: push 00E3D060h
  loc_00E22799: push 006D87C0h
  loc_00E2279E: call [004011A0h] ; __vbaNew2
  loc_00E227A4: mov eax, [00E3D060h]
  loc_00E227A9: mov ecx, [eax]
  loc_00E227AB: push eax
  loc_00E227AC: call [ecx+000004E0h]
  loc_00E227B2: lea edx, var_20
  loc_00E227B5: push eax
  loc_00E227B6: push edx
  loc_00E227B7: call __vbaObjSet
  loc_00E227B9: mov edi, eax
  loc_00E227BB: lea ecx, var_18
  loc_00E227BE: push ecx
  loc_00E227BF: push edi
  loc_00E227C0: mov eax, [edi]
  loc_00E227C2: call [eax+00000050h]
  loc_00E227C5: test eax, eax
  loc_00E227C7: fnclex
  loc_00E227C9: jge 00E227DAh
  loc_00E227CB: push 00000050h
  loc_00E227CD: push 006E0410h
  loc_00E227D2: push edi
  loc_00E227D3: push eax
  loc_00E227D4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E227DA: mov edx, var_18
  loc_00E227DD: push edx
  loc_00E227DE: push 006E05FCh ; "Katholik"
  loc_00E227E3: call ebx
  loc_00E227E5: mov edi, eax
  loc_00E227E7: lea ecx, var_18
  loc_00E227EA: neg edi
  loc_00E227EC: sbb edi, edi
  loc_00E227EE: inc edi
  loc_00E227EF: neg edi
  loc_00E227F1: call [00401258h] ; __vbaFreeStr
  loc_00E227F7: lea ecx, var_20
  loc_00E227FA: call [00401254h] ; __vbaFreeObj
  loc_00E22800: mov eax, [00E3D060h]
  loc_00E22805: test di, di
  loc_00E22808: jz 00E22F36h
  loc_00E2280E: test eax, eax
  loc_00E22810: jnz 00E22827h
  loc_00E22812: push 00E3D060h
  loc_00E22817: push 006D87C0h
  loc_00E2281C: call [004011A0h] ; __vbaNew2
  loc_00E22822: mov eax, [00E3D060h]
  loc_00E22827: mov ecx, [eax]
  loc_00E22829: push eax
  loc_00E2282A: call [ecx+00000474h]
  loc_00E22830: lea edx, var_20
  loc_00E22833: push eax
  loc_00E22834: push edx
  loc_00E22835: call __vbaObjSet
  loc_00E22837: mov edi, eax
  loc_00E22839: push 000927C0h
  loc_00E2283E: mov ebx, [edi]
  loc_00E22840: call [00401018h] ; __vbaStrI4
  loc_00E22846: mov edx, eax
  loc_00E22848: lea ecx, var_18
  loc_00E2284B: call [00401228h] ; __vbaStrMove
  loc_00E22851: push eax
  loc_00E22852: push edi
  loc_00E22853: call [ebx+00000054h]
  loc_00E22856: test eax, eax
  loc_00E22858: fnclex
  loc_00E2285A: jge 00E2286Bh
  loc_00E2285C: push 00000054h
  loc_00E2285E: push 006E0410h
  loc_00E22863: push edi
  loc_00E22864: push eax
  loc_00E22865: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2286B: lea ecx, var_18
  loc_00E2286E: call [00401258h] ; __vbaFreeStr
  loc_00E22874: lea ecx, var_20
  loc_00E22877: call [00401254h] ; __vbaFreeObj
  loc_00E2287D: mov eax, [00E3D060h]
  loc_00E22882: test eax, eax
  loc_00E22884: jnz 00E2289Bh
  loc_00E22886: push 00E3D060h
  loc_00E2288B: push 006D87C0h
  loc_00E22890: call [004011A0h] ; __vbaNew2
  loc_00E22896: mov eax, [00E3D060h]
  loc_00E2289B: mov ecx, [eax]
  loc_00E2289D: push eax
  loc_00E2289E: call [ecx+00000470h]
  loc_00E228A4: lea edx, var_20
  loc_00E228A7: push eax
  loc_00E228A8: push edx
  loc_00E228A9: call __vbaObjSet
  loc_00E228AB: mov edi, eax
  loc_00E228AD: push 0003A980h
  loc_00E228B2: mov ebx, [edi]
  loc_00E228B4: call [00401018h] ; __vbaStrI4
  loc_00E228BA: mov edx, eax
  loc_00E228BC: lea ecx, var_18
  loc_00E228BF: call [00401228h] ; __vbaStrMove
  loc_00E228C5: push eax
  loc_00E228C6: push edi
  loc_00E228C7: call [ebx+00000054h]
  loc_00E228CA: test eax, eax
  loc_00E228CC: fnclex
  loc_00E228CE: jge 00E228DFh
  loc_00E228D0: push 00000054h
  loc_00E228D2: push 006E0410h
  loc_00E228D7: push edi
  loc_00E228D8: push eax
  loc_00E228D9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E228DF: lea ecx, var_18
  loc_00E228E2: call [00401258h] ; __vbaFreeStr
  loc_00E228E8: lea ecx, var_20
  loc_00E228EB: call [00401254h] ; __vbaFreeObj
  loc_00E228F1: mov eax, [00E3D060h]
  loc_00E228F6: test eax, eax
  loc_00E228F8: jnz 00E2290Fh
  loc_00E228FA: push 00E3D060h
  loc_00E228FF: push 006D87C0h
  loc_00E22904: call [004011A0h] ; __vbaNew2
  loc_00E2290A: mov eax, [00E3D060h]
  loc_00E2290F: mov ecx, [eax]
  loc_00E22911: push eax
  loc_00E22912: call [ecx+0000046Ch]
  loc_00E22918: lea edx, var_20
  loc_00E2291B: push eax
  loc_00E2291C: push edx
  loc_00E2291D: call __vbaObjSet
  loc_00E2291F: mov edi, eax
  loc_00E22921: push 00077A10h
  loc_00E22926: mov ebx, [edi]
  loc_00E22928: call [00401018h] ; __vbaStrI4
  loc_00E2292E: mov edx, eax
  loc_00E22930: lea ecx, var_18
  loc_00E22933: call [00401228h] ; __vbaStrMove
  loc_00E22939: push eax
  loc_00E2293A: push edi
  loc_00E2293B: call [ebx+00000054h]
  loc_00E2293E: test eax, eax
  loc_00E22940: fnclex
  loc_00E22942: jge 00E22953h
  loc_00E22944: push 00000054h
  loc_00E22946: push 006E0410h
  loc_00E2294B: push edi
  loc_00E2294C: push eax
  loc_00E2294D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22953: lea ecx, var_18
  loc_00E22956: call [00401258h] ; __vbaFreeStr
  loc_00E2295C: lea ecx, var_20
  loc_00E2295F: call [00401254h] ; __vbaFreeObj
  loc_00E22965: mov eax, [00E3D060h]
  loc_00E2296A: test eax, eax
  loc_00E2296C: jnz 00E22983h
  loc_00E2296E: push 00E3D060h
  loc_00E22973: push 006D87C0h
  loc_00E22978: call [004011A0h] ; __vbaNew2
  loc_00E2297E: mov eax, [00E3D060h]
  loc_00E22983: mov ecx, [eax]
  loc_00E22985: push eax
  loc_00E22986: call [ecx+00000454h]
  loc_00E2298C: lea edx, var_20
  loc_00E2298F: push eax
  loc_00E22990: push edx
  loc_00E22991: call __vbaObjSet
  loc_00E22993: mov edi, eax
  loc_00E22995: push 000124F8h
  loc_00E2299A: mov ebx, [edi]
  loc_00E2299C: call [00401018h] ; __vbaStrI4
  loc_00E229A2: mov edx, eax
  loc_00E229A4: lea ecx, var_18
  loc_00E229A7: call [00401228h] ; __vbaStrMove
  loc_00E229AD: push eax
  loc_00E229AE: push edi
  loc_00E229AF: call [ebx+00000054h]
  loc_00E229B2: test eax, eax
  loc_00E229B4: fnclex
  loc_00E229B6: jge 00E229C7h
  loc_00E229B8: push 00000054h
  loc_00E229BA: push 006E0410h
  loc_00E229BF: push edi
  loc_00E229C0: push eax
  loc_00E229C1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E229C7: lea ecx, var_18
  loc_00E229CA: call [00401258h] ; __vbaFreeStr
  loc_00E229D0: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E229D6: lea ecx, var_20
  loc_00E229D9: call ebx
  loc_00E229DB: mov eax, [00E3D060h]
  loc_00E229E0: test eax, eax
  loc_00E229E2: jnz 00E229F9h
  loc_00E229E4: push 00E3D060h
  loc_00E229E9: push 006D87C0h
  loc_00E229EE: call [004011A0h] ; __vbaNew2
  loc_00E229F4: mov eax, [00E3D060h]
  loc_00E229F9: mov ecx, [eax]
  loc_00E229FB: push eax
  loc_00E229FC: call [ecx+0000042Ch]
  loc_00E22A02: lea edx, var_20
  loc_00E22A05: push eax
  loc_00E22A06: push edx
  loc_00E22A07: call __vbaObjSet
  loc_00E22A09: mov edi, eax
  loc_00E22A0B: push FFFFFFFFh
  loc_00E22A0D: push edi
  loc_00E22A0E: mov eax, [edi]
  loc_00E22A10: call [eax+0000009Ch]
  loc_00E22A16: test eax, eax
  loc_00E22A18: fnclex
  loc_00E22A1A: jge 00E22A2Eh
  loc_00E22A1C: push 0000009Ch
  loc_00E22A21: push 006DCAD0h
  loc_00E22A26: push edi
  loc_00E22A27: push eax
  loc_00E22A28: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22A2E: lea ecx, var_20
  loc_00E22A31: call ebx
  loc_00E22A33: mov eax, [00E3D060h]
  loc_00E22A38: test eax, eax
  loc_00E22A3A: jnz 00E22A51h
  loc_00E22A3C: push 00E3D060h
  loc_00E22A41: push 006D87C0h
  loc_00E22A46: call [004011A0h] ; __vbaNew2
  loc_00E22A4C: mov eax, [00E3D060h]
  loc_00E22A51: mov ecx, [eax]
  loc_00E22A53: push eax
  loc_00E22A54: call [ecx+00000434h]
  loc_00E22A5A: lea edx, var_20
  loc_00E22A5D: push eax
  loc_00E22A5E: push edx
  loc_00E22A5F: call __vbaObjSet
  loc_00E22A61: mov edi, eax
  loc_00E22A63: push 00030D40h
  loc_00E22A68: mov ebx, [edi]
  loc_00E22A6A: call [00401018h] ; __vbaStrI4
  loc_00E22A70: mov edx, eax
  loc_00E22A72: lea ecx, var_18
  loc_00E22A75: call [00401228h] ; __vbaStrMove
  loc_00E22A7B: push eax
  loc_00E22A7C: push edi
  loc_00E22A7D: call [ebx+00000054h]
  loc_00E22A80: test eax, eax
  loc_00E22A82: fnclex
  loc_00E22A84: jge 00E22A95h
  loc_00E22A86: push 00000054h
  loc_00E22A88: push 006E0410h
  loc_00E22A8D: push edi
  loc_00E22A8E: push eax
  loc_00E22A8F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22A95: lea ecx, var_18
  loc_00E22A98: call [00401258h] ; __vbaFreeStr
  loc_00E22A9E: lea ecx, var_20
  loc_00E22AA1: call [00401254h] ; __vbaFreeObj
  loc_00E22AA7: mov eax, [00E3D060h]
  loc_00E22AAC: test eax, eax
  loc_00E22AAE: jnz 00E22AC5h
  loc_00E22AB0: push 00E3D060h
  loc_00E22AB5: push 006D87C0h
  loc_00E22ABA: call [004011A0h] ; __vbaNew2
  loc_00E22AC0: mov eax, [00E3D060h]
  loc_00E22AC5: mov ecx, [eax]
  loc_00E22AC7: push eax
  loc_00E22AC8: call [ecx+000003FCh]
  loc_00E22ACE: lea edx, var_20
  loc_00E22AD1: push eax
  loc_00E22AD2: push edx
  loc_00E22AD3: call __vbaObjSet
  loc_00E22AD5: mov edi, eax
  loc_00E22AD7: push 00023668h
  loc_00E22ADC: mov ebx, [edi]
  loc_00E22ADE: call [00401018h] ; __vbaStrI4
  loc_00E22AE4: mov edx, eax
  loc_00E22AE6: lea ecx, var_18
  loc_00E22AE9: call [00401228h] ; __vbaStrMove
  loc_00E22AEF: push eax
  loc_00E22AF0: push edi
  loc_00E22AF1: call [ebx+00000054h]
  loc_00E22AF4: test eax, eax
  loc_00E22AF6: fnclex
  loc_00E22AF8: jge 00E22B09h
  loc_00E22AFA: push 00000054h
  loc_00E22AFC: push 006E0410h
  loc_00E22B01: push edi
  loc_00E22B02: push eax
  loc_00E22B03: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22B09: lea ecx, var_18
  loc_00E22B0C: call [00401258h] ; __vbaFreeStr
  loc_00E22B12: lea ecx, var_20
  loc_00E22B15: call [00401254h] ; __vbaFreeObj
  loc_00E22B1B: mov eax, [00E3D060h]
  loc_00E22B20: test eax, eax
  loc_00E22B22: jnz 00E22B39h
  loc_00E22B24: push 00E3D060h
  loc_00E22B29: push 006D87C0h
  loc_00E22B2E: call [004011A0h] ; __vbaNew2
  loc_00E22B34: mov eax, [00E3D060h]
  loc_00E22B39: mov ecx, [eax]
  loc_00E22B3B: push eax
  loc_00E22B3C: call [ecx+000003ECh]
  loc_00E22B42: lea edx, var_20
  loc_00E22B45: push eax
  loc_00E22B46: push edx
  loc_00E22B47: call __vbaObjSet
  loc_00E22B49: mov edi, eax
  loc_00E22B4B: push 00030D40h
  loc_00E22B50: mov ebx, [edi]
  loc_00E22B52: call [00401018h] ; __vbaStrI4
  loc_00E22B58: mov edx, eax
  loc_00E22B5A: lea ecx, var_18
  loc_00E22B5D: call [00401228h] ; __vbaStrMove
  loc_00E22B63: push eax
  loc_00E22B64: push edi
  loc_00E22B65: call [ebx+00000054h]
  loc_00E22B68: test eax, eax
  loc_00E22B6A: fnclex
  loc_00E22B6C: jge 00E22B7Dh
  loc_00E22B6E: push 00000054h
  loc_00E22B70: push 006E0410h
  loc_00E22B75: push edi
  loc_00E22B76: push eax
  loc_00E22B77: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22B7D: lea ecx, var_18
  loc_00E22B80: call [00401258h] ; __vbaFreeStr
  loc_00E22B86: lea ecx, var_20
  loc_00E22B89: call [00401254h] ; __vbaFreeObj
  loc_00E22B8F: mov eax, [00E3D060h]
  loc_00E22B94: test eax, eax
  loc_00E22B96: jnz 00E22BADh
  loc_00E22B98: push 00E3D060h
  loc_00E22B9D: push 006D87C0h
  loc_00E22BA2: call [004011A0h] ; __vbaNew2
  loc_00E22BA8: mov eax, [00E3D060h]
  loc_00E22BAD: mov ecx, [eax]
  loc_00E22BAF: push eax
  loc_00E22BB0: call [ecx+000003DCh]
  loc_00E22BB6: lea edx, var_20
  loc_00E22BB9: push eax
  loc_00E22BBA: push edx
  loc_00E22BBB: call __vbaObjSet
  loc_00E22BBD: mov edi, eax
  loc_00E22BBF: push 0000FDE8h
  loc_00E22BC4: mov ebx, [edi]
  loc_00E22BC6: call [00401018h] ; __vbaStrI4
  loc_00E22BCC: mov edx, eax
  loc_00E22BCE: lea ecx, var_18
  loc_00E22BD1: call [00401228h] ; __vbaStrMove
  loc_00E22BD7: push eax
  loc_00E22BD8: push edi
  loc_00E22BD9: call [ebx+00000054h]
  loc_00E22BDC: test eax, eax
  loc_00E22BDE: fnclex
  loc_00E22BE0: jge 00E22BF1h
  loc_00E22BE2: push 00000054h
  loc_00E22BE4: push 006E0410h
  loc_00E22BE9: push edi
  loc_00E22BEA: push eax
  loc_00E22BEB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22BF1: lea ecx, var_18
  loc_00E22BF4: call [00401258h] ; __vbaFreeStr
  loc_00E22BFA: lea ecx, var_20
  loc_00E22BFD: call [00401254h] ; __vbaFreeObj
  loc_00E22C03: mov eax, [00E3D060h]
  loc_00E22C08: test eax, eax
  loc_00E22C0A: jnz 00E22C21h
  loc_00E22C0C: push 00E3D060h
  loc_00E22C11: push 006D87C0h
  loc_00E22C16: call [004011A0h] ; __vbaNew2
  loc_00E22C1C: mov eax, [00E3D060h]
  loc_00E22C21: mov ecx, [eax]
  loc_00E22C23: push eax
  loc_00E22C24: call [ecx+000003CCh]
  loc_00E22C2A: lea edx, var_20
  loc_00E22C2D: push eax
  loc_00E22C2E: push edx
  loc_00E22C2F: call __vbaObjSet
  loc_00E22C31: mov edi, eax
  loc_00E22C33: push 0001E848h
  loc_00E22C38: mov ebx, [edi]
  loc_00E22C3A: call [00401018h] ; __vbaStrI4
  loc_00E22C40: mov edx, eax
  loc_00E22C42: lea ecx, var_18
  loc_00E22C45: call [00401228h] ; __vbaStrMove
  loc_00E22C4B: push eax
  loc_00E22C4C: push edi
  loc_00E22C4D: call [ebx+00000054h]
  loc_00E22C50: test eax, eax
  loc_00E22C52: fnclex
  loc_00E22C54: jge 00E22C65h
  loc_00E22C56: push 00000054h
  loc_00E22C58: push 006E0410h
  loc_00E22C5D: push edi
  loc_00E22C5E: push eax
  loc_00E22C5F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22C65: lea ecx, var_18
  loc_00E22C68: call [00401258h] ; __vbaFreeStr
  loc_00E22C6E: lea ecx, var_20
  loc_00E22C71: call [00401254h] ; __vbaFreeObj
  loc_00E22C77: mov eax, [00E3D060h]
  loc_00E22C7C: test eax, eax
  loc_00E22C7E: jnz 00E22C95h
  loc_00E22C80: push 00E3D060h
  loc_00E22C85: push 006D87C0h
  loc_00E22C8A: call [004011A0h] ; __vbaNew2
  loc_00E22C90: mov eax, [00E3D060h]
  loc_00E22C95: mov ecx, [eax]
  loc_00E22C97: push eax
  loc_00E22C98: call [ecx+000003BCh]
  loc_00E22C9E: lea edx, var_20
  loc_00E22CA1: push eax
  loc_00E22CA2: push edx
  loc_00E22CA3: call __vbaObjSet
  loc_00E22CA5: mov edi, eax
  loc_00E22CA7: push 00023668h
  loc_00E22CAC: mov ebx, [edi]
  loc_00E22CAE: call [00401018h] ; __vbaStrI4
  loc_00E22CB4: mov edx, eax
  loc_00E22CB6: lea ecx, var_18
  loc_00E22CB9: call [00401228h] ; __vbaStrMove
  loc_00E22CBF: push eax
  loc_00E22CC0: push edi
  loc_00E22CC1: call [ebx+00000054h]
  loc_00E22CC4: test eax, eax
  loc_00E22CC6: fnclex
  loc_00E22CC8: jge 00E22CD9h
  loc_00E22CCA: push 00000054h
  loc_00E22CCC: push 006E0410h
  loc_00E22CD1: push edi
  loc_00E22CD2: push eax
  loc_00E22CD3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22CD9: lea ecx, var_18
  loc_00E22CDC: call [00401258h] ; __vbaFreeStr
  loc_00E22CE2: lea ecx, var_20
  loc_00E22CE5: call [00401254h] ; __vbaFreeObj
  loc_00E22CEB: mov eax, [00E3D060h]
  loc_00E22CF0: test eax, eax
  loc_00E22CF2: jnz 00E22D09h
  loc_00E22CF4: push 00E3D060h
  loc_00E22CF9: push 006D87C0h
  loc_00E22CFE: call [004011A0h] ; __vbaNew2
  loc_00E22D04: mov eax, [00E3D060h]
  loc_00E22D09: mov ecx, [eax]
  loc_00E22D0B: push eax
  loc_00E22D0C: call [ecx+00000440h]
  loc_00E22D12: lea edx, var_20
  loc_00E22D15: push eax
  loc_00E22D16: push edx
  loc_00E22D17: call __vbaObjSet
  loc_00E22D19: mov edi, eax
  loc_00E22D1B: push 0007A120h
  loc_00E22D20: mov ebx, [edi]
  loc_00E22D22: call [00401018h] ; __vbaStrI4
  loc_00E22D28: mov edx, eax
  loc_00E22D2A: lea ecx, var_18
  loc_00E22D2D: call [00401228h] ; __vbaStrMove
  loc_00E22D33: push eax
  loc_00E22D34: push edi
  loc_00E22D35: call [ebx+00000054h]
  loc_00E22D38: test eax, eax
  loc_00E22D3A: fnclex
  loc_00E22D3C: jge 00E22D4Dh
  loc_00E22D3E: push 00000054h
  loc_00E22D40: push 006E0410h
  loc_00E22D45: push edi
  loc_00E22D46: push eax
  loc_00E22D47: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22D4D: lea ecx, var_18
  loc_00E22D50: call [00401258h] ; __vbaFreeStr
  loc_00E22D56: lea ecx, var_20
  loc_00E22D59: call [00401254h] ; __vbaFreeObj
  loc_00E22D5F: mov eax, [00E3D060h]
  loc_00E22D64: test eax, eax
  loc_00E22D66: jnz 00E22D7Dh
  loc_00E22D68: push 00E3D060h
  loc_00E22D6D: push 006D87C0h
  loc_00E22D72: call [004011A0h] ; __vbaNew2
  loc_00E22D78: mov eax, [00E3D060h]
  loc_00E22D7D: mov ecx, [eax]
  loc_00E22D7F: push eax
  loc_00E22D80: call [ecx+000003A8h]
  loc_00E22D86: lea edx, var_20
  loc_00E22D89: push eax
  loc_00E22D8A: push edx
  loc_00E22D8B: call __vbaObjSet
  loc_00E22D8D: mov edi, eax
  loc_00E22D8F: push 00003A98h
  loc_00E22D94: mov ebx, [edi]
  loc_00E22D96: call [0040100Ch] ; __vbaStrI2
  loc_00E22D9C: mov edx, eax
  loc_00E22D9E: lea ecx, var_18
  loc_00E22DA1: call [00401228h] ; __vbaStrMove
  loc_00E22DA7: push eax
  loc_00E22DA8: push edi
  loc_00E22DA9: call [ebx+00000054h]
  loc_00E22DAC: test eax, eax
  loc_00E22DAE: fnclex
  loc_00E22DB0: jge 00E22DC1h
  loc_00E22DB2: push 00000054h
  loc_00E22DB4: push 006E0410h
  loc_00E22DB9: push edi
  loc_00E22DBA: push eax
  loc_00E22DBB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22DC1: lea ecx, var_18
  loc_00E22DC4: call [00401258h] ; __vbaFreeStr
  loc_00E22DCA: lea ecx, var_20
  loc_00E22DCD: call [00401254h] ; __vbaFreeObj
  loc_00E22DD3: mov eax, [00E3D060h]
  loc_00E22DD8: test eax, eax
  loc_00E22DDA: jnz 00E22DF1h
  loc_00E22DDC: push 00E3D060h
  loc_00E22DE1: push 006D87C0h
  loc_00E22DE6: call [004011A0h] ; __vbaNew2
  loc_00E22DEC: mov eax, [00E3D060h]
  loc_00E22DF1: mov ecx, [eax]
  loc_00E22DF3: push eax
  loc_00E22DF4: call [ecx+00000398h]
  loc_00E22DFA: lea edx, var_20
  loc_00E22DFD: push eax
  loc_00E22DFE: push edx
  loc_00E22DFF: call __vbaObjSet
  loc_00E22E01: mov edi, eax
  loc_00E22E03: push 00003A98h
  loc_00E22E08: mov ebx, [edi]
  loc_00E22E0A: call [0040100Ch] ; __vbaStrI2
  loc_00E22E10: mov edx, eax
  loc_00E22E12: lea ecx, var_18
  loc_00E22E15: call [00401228h] ; __vbaStrMove
  loc_00E22E1B: push eax
  loc_00E22E1C: push edi
  loc_00E22E1D: call [ebx+00000054h]
  loc_00E22E20: test eax, eax
  loc_00E22E22: fnclex
  loc_00E22E24: jge 00E22E35h
  loc_00E22E26: push 00000054h
  loc_00E22E28: push 006E0410h
  loc_00E22E2D: push edi
  loc_00E22E2E: push eax
  loc_00E22E2F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22E35: lea ecx, var_18
  loc_00E22E38: call [00401258h] ; __vbaFreeStr
  loc_00E22E3E: lea ecx, var_20
  loc_00E22E41: call [00401254h] ; __vbaFreeObj
  loc_00E22E47: mov eax, [00E3D060h]
  loc_00E22E4C: test eax, eax
  loc_00E22E4E: jnz 00E22E65h
  loc_00E22E50: push 00E3D060h
  loc_00E22E55: push 006D87C0h
  loc_00E22E5A: call [004011A0h] ; __vbaNew2
  loc_00E22E60: mov eax, [00E3D060h]
  loc_00E22E65: mov ecx, [eax]
  loc_00E22E67: push eax
  loc_00E22E68: call [ecx+00000388h]
  loc_00E22E6E: lea edx, var_20
  loc_00E22E71: push eax
  loc_00E22E72: push edx
  loc_00E22E73: call __vbaObjSet
  loc_00E22E75: mov edi, eax
  loc_00E22E77: push 00001388h
  loc_00E22E7C: mov ebx, [edi]
  loc_00E22E7E: call [0040100Ch] ; __vbaStrI2
  loc_00E22E84: mov edx, eax
  loc_00E22E86: lea ecx, var_18
  loc_00E22E89: call [00401228h] ; __vbaStrMove
  loc_00E22E8F: push eax
  loc_00E22E90: push edi
  loc_00E22E91: call [ebx+00000054h]
  loc_00E22E94: test eax, eax
  loc_00E22E96: fnclex
  loc_00E22E98: jge 00E22EA9h
  loc_00E22E9A: push 00000054h
  loc_00E22E9C: push 006E0410h
  loc_00E22EA1: push edi
  loc_00E22EA2: push eax
  loc_00E22EA3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22EA9: lea ecx, var_18
  loc_00E22EAC: call [00401258h] ; __vbaFreeStr
  loc_00E22EB2: lea ecx, var_20
  loc_00E22EB5: call [00401254h] ; __vbaFreeObj
  loc_00E22EBB: mov eax, [00E3D060h]
  loc_00E22EC0: test eax, eax
  loc_00E22EC2: jnz 00E22ED9h
  loc_00E22EC4: push 00E3D060h
  loc_00E22EC9: push 006D87C0h
  loc_00E22ECE: call [004011A0h] ; __vbaNew2
  loc_00E22ED4: mov eax, [00E3D060h]
  loc_00E22ED9: mov ecx, [eax]
  loc_00E22EDB: push eax
  loc_00E22EDC: call [ecx+00000378h]
  loc_00E22EE2: lea edx, var_20
  loc_00E22EE5: push eax
  loc_00E22EE6: push edx
  loc_00E22EE7: call __vbaObjSet
  loc_00E22EE9: mov edi, eax
  loc_00E22EEB: push 00001388h
  loc_00E22EF0: mov ebx, [edi]
  loc_00E22EF2: call [0040100Ch] ; __vbaStrI2
  loc_00E22EF8: mov edx, eax
  loc_00E22EFA: lea ecx, var_18
  loc_00E22EFD: call [00401228h] ; __vbaStrMove
  loc_00E22F03: push eax
  loc_00E22F04: push edi
  loc_00E22F05: call [ebx+00000054h]
  loc_00E22F08: test eax, eax
  loc_00E22F0A: fnclex
  loc_00E22F0C: jge 00E22F1Dh
  loc_00E22F0E: push 00000054h
  loc_00E22F10: push 006E0410h
  loc_00E22F15: push edi
  loc_00E22F16: push eax
  loc_00E22F17: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22F1D: lea ecx, var_18
  loc_00E22F20: call [00401258h] ; __vbaFreeStr
  loc_00E22F26: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E22F2C: lea ecx, var_20
  loc_00E22F2F: call edi
  loc_00E22F31: jmp 00E234F6h
  loc_00E22F36: test eax, eax
  loc_00E22F38: jnz 00E22F4Fh
  loc_00E22F3A: push 00E3D060h
  loc_00E22F3F: push 006D87C0h
  loc_00E22F44: call [004011A0h] ; __vbaNew2
  loc_00E22F4A: mov eax, [00E3D060h]
  loc_00E22F4F: mov ecx, [eax]
  loc_00E22F51: push eax
  loc_00E22F52: call [ecx+00000474h]
  loc_00E22F58: lea edx, var_20
  loc_00E22F5B: push eax
  loc_00E22F5C: push edx
  loc_00E22F5D: call __vbaObjSet
  loc_00E22F5F: mov edi, eax
  loc_00E22F61: push 000AAE60h
  loc_00E22F66: mov ebx, [edi]
  loc_00E22F68: call [00401018h] ; __vbaStrI4
  loc_00E22F6E: mov edx, eax
  loc_00E22F70: lea ecx, var_18
  loc_00E22F73: call [00401228h] ; __vbaStrMove
  loc_00E22F79: push eax
  loc_00E22F7A: push edi
  loc_00E22F7B: call [ebx+00000054h]
  loc_00E22F7E: test eax, eax
  loc_00E22F80: fnclex
  loc_00E22F82: jge 00E22F93h
  loc_00E22F84: push 00000054h
  loc_00E22F86: push 006E0410h
  loc_00E22F8B: push edi
  loc_00E22F8C: push eax
  loc_00E22F8D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E22F93: lea ecx, var_18
  loc_00E22F96: call [00401258h] ; __vbaFreeStr
  loc_00E22F9C: lea ecx, var_20
  loc_00E22F9F: call [00401254h] ; __vbaFreeObj
  loc_00E22FA5: mov eax, [00E3D060h]
  loc_00E22FAA: test eax, eax
  loc_00E22FAC: jnz 00E22FC3h
  loc_00E22FAE: push 00E3D060h
  loc_00E22FB3: push 006D87C0h
  loc_00E22FB8: call [004011A0h] ; __vbaNew2
  loc_00E22FBE: mov eax, [00E3D060h]
  loc_00E22FC3: mov ecx, [eax]
  loc_00E22FC5: push eax
  loc_00E22FC6: call [ecx+00000470h]
  loc_00E22FCC: lea edx, var_20
  loc_00E22FCF: push eax
  loc_00E22FD0: push edx
  loc_00E22FD1: call __vbaObjSet
  loc_00E22FD3: mov edi, eax
  loc_00E22FD5: push 0003A980h
  loc_00E22FDA: mov ebx, [edi]
  loc_00E22FDC: call [00401018h] ; __vbaStrI4
  loc_00E22FE2: mov edx, eax
  loc_00E22FE4: lea ecx, var_18
  loc_00E22FE7: call [00401228h] ; __vbaStrMove
  loc_00E22FED: push eax
  loc_00E22FEE: push edi
  loc_00E22FEF: call [ebx+00000054h]
  loc_00E22FF2: test eax, eax
  loc_00E22FF4: fnclex
  loc_00E22FF6: jge 00E23007h
  loc_00E22FF8: push 00000054h
  loc_00E22FFA: push 006E0410h
  loc_00E22FFF: push edi
  loc_00E23000: push eax
  loc_00E23001: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23007: lea ecx, var_18
  loc_00E2300A: call [00401258h] ; __vbaFreeStr
  loc_00E23010: lea ecx, var_20
  loc_00E23013: call [00401254h] ; __vbaFreeObj
  loc_00E23019: mov eax, [00E3D060h]
  loc_00E2301E: test eax, eax
  loc_00E23020: jnz 00E23037h
  loc_00E23022: push 00E3D060h
  loc_00E23027: push 006D87C0h
  loc_00E2302C: call [004011A0h] ; __vbaNew2
  loc_00E23032: mov eax, [00E3D060h]
  loc_00E23037: mov ecx, [eax]
  loc_00E23039: push eax
  loc_00E2303A: call [ecx+0000046Ch]
  loc_00E23040: lea edx, var_20
  loc_00E23043: push eax
  loc_00E23044: push edx
  loc_00E23045: call __vbaObjSet
  loc_00E23047: mov edi, eax
  loc_00E23049: push 00077A10h
  loc_00E2304E: mov ebx, [edi]
  loc_00E23050: call [00401018h] ; __vbaStrI4
  loc_00E23056: mov edx, eax
  loc_00E23058: lea ecx, var_18
  loc_00E2305B: call [00401228h] ; __vbaStrMove
  loc_00E23061: push eax
  loc_00E23062: push edi
  loc_00E23063: call [ebx+00000054h]
  loc_00E23066: test eax, eax
  loc_00E23068: fnclex
  loc_00E2306A: jge 00E2307Bh
  loc_00E2306C: push 00000054h
  loc_00E2306E: push 006E0410h
  loc_00E23073: push edi
  loc_00E23074: push eax
  loc_00E23075: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2307B: lea ecx, var_18
  loc_00E2307E: call [00401258h] ; __vbaFreeStr
  loc_00E23084: lea ecx, var_20
  loc_00E23087: call [00401254h] ; __vbaFreeObj
  loc_00E2308D: mov eax, [00E3D060h]
  loc_00E23092: test eax, eax
  loc_00E23094: jnz 00E230ABh
  loc_00E23096: push 00E3D060h
  loc_00E2309B: push 006D87C0h
  loc_00E230A0: call [004011A0h] ; __vbaNew2
  loc_00E230A6: mov eax, [00E3D060h]
  loc_00E230AB: mov ecx, [eax]
  loc_00E230AD: push eax
  loc_00E230AE: call [ecx+00000454h]
  loc_00E230B4: lea edx, var_20
  loc_00E230B7: push eax
  loc_00E230B8: push edx
  loc_00E230B9: call __vbaObjSet
  loc_00E230BB: mov edi, eax
  loc_00E230BD: push 000124F8h
  loc_00E230C2: mov ebx, [edi]
  loc_00E230C4: call [00401018h] ; __vbaStrI4
  loc_00E230CA: mov edx, eax
  loc_00E230CC: lea ecx, var_18
  loc_00E230CF: call [00401228h] ; __vbaStrMove
  loc_00E230D5: push eax
  loc_00E230D6: push edi
  loc_00E230D7: call [ebx+00000054h]
  loc_00E230DA: test eax, eax
  loc_00E230DC: fnclex
  loc_00E230DE: jge 00E230EFh
  loc_00E230E0: push 00000054h
  loc_00E230E2: push 006E0410h
  loc_00E230E7: push edi
  loc_00E230E8: push eax
  loc_00E230E9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E230EF: lea ecx, var_18
  loc_00E230F2: call [00401258h] ; __vbaFreeStr
  loc_00E230F8: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E230FE: lea ecx, var_20
  loc_00E23101: call ebx
  loc_00E23103: mov eax, [00E3D060h]
  loc_00E23108: test eax, eax
  loc_00E2310A: jnz 00E23121h
  loc_00E2310C: push 00E3D060h
  loc_00E23111: push 006D87C0h
  loc_00E23116: call [004011A0h] ; __vbaNew2
  loc_00E2311C: mov eax, [00E3D060h]
  loc_00E23121: mov ecx, [eax]
  loc_00E23123: push eax
  loc_00E23124: call [ecx+0000042Ch]
  loc_00E2312A: lea edx, var_20
  loc_00E2312D: push eax
  loc_00E2312E: push edx
  loc_00E2312F: call __vbaObjSet
  loc_00E23131: mov edi, eax
  loc_00E23133: push FFFFFFFFh
  loc_00E23135: push edi
  loc_00E23136: mov eax, [edi]
  loc_00E23138: call [eax+0000009Ch]
  loc_00E2313E: test eax, eax
  loc_00E23140: fnclex
  loc_00E23142: jge 00E23156h
  loc_00E23144: push 0000009Ch
  loc_00E23149: push 006DCAD0h
  loc_00E2314E: push edi
  loc_00E2314F: push eax
  loc_00E23150: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23156: lea ecx, var_20
  loc_00E23159: call ebx
  loc_00E2315B: mov eax, [00E3D060h]
  loc_00E23160: test eax, eax
  loc_00E23162: jnz 00E23179h
  loc_00E23164: push 00E3D060h
  loc_00E23169: push 006D87C0h
  loc_00E2316E: call [004011A0h] ; __vbaNew2
  loc_00E23174: mov eax, [00E3D060h]
  loc_00E23179: mov ecx, [eax]
  loc_00E2317B: push eax
  loc_00E2317C: call [ecx+00000434h]
  loc_00E23182: lea edx, var_20
  loc_00E23185: push eax
  loc_00E23186: push edx
  loc_00E23187: call __vbaObjSet
  loc_00E23189: mov edi, eax
  loc_00E2318B: push 00030D40h
  loc_00E23190: mov ebx, [edi]
  loc_00E23192: call [00401018h] ; __vbaStrI4
  loc_00E23198: mov edx, eax
  loc_00E2319A: lea ecx, var_18
  loc_00E2319D: call [00401228h] ; __vbaStrMove
  loc_00E231A3: push eax
  loc_00E231A4: push edi
  loc_00E231A5: call [ebx+00000054h]
  loc_00E231A8: test eax, eax
  loc_00E231AA: fnclex
  loc_00E231AC: jge 00E231BDh
  loc_00E231AE: push 00000054h
  loc_00E231B0: push 006E0410h
  loc_00E231B5: push edi
  loc_00E231B6: push eax
  loc_00E231B7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E231BD: lea ecx, var_18
  loc_00E231C0: call [00401258h] ; __vbaFreeStr
  loc_00E231C6: lea ecx, var_20
  loc_00E231C9: call [00401254h] ; __vbaFreeObj
  loc_00E231CF: mov eax, [00E3D060h]
  loc_00E231D4: test eax, eax
  loc_00E231D6: jnz 00E231EDh
  loc_00E231D8: push 00E3D060h
  loc_00E231DD: push 006D87C0h
  loc_00E231E2: call [004011A0h] ; __vbaNew2
  loc_00E231E8: mov eax, [00E3D060h]
  loc_00E231ED: mov ecx, [eax]
  loc_00E231EF: push eax
  loc_00E231F0: call [ecx+00000440h]
  loc_00E231F6: lea edx, var_20
  loc_00E231F9: push eax
  loc_00E231FA: push edx
  loc_00E231FB: call __vbaObjSet
  loc_00E231FD: mov edi, eax
  loc_00E231FF: push 0007A120h
  loc_00E23204: mov ebx, [edi]
  loc_00E23206: call [00401018h] ; __vbaStrI4
  loc_00E2320C: mov edx, eax
  loc_00E2320E: lea ecx, var_18
  loc_00E23211: call [00401228h] ; __vbaStrMove
  loc_00E23217: push eax
  loc_00E23218: push edi
  loc_00E23219: call [ebx+00000054h]
  loc_00E2321C: test eax, eax
  loc_00E2321E: fnclex
  loc_00E23220: jge 00E23231h
  loc_00E23222: push 00000054h
  loc_00E23224: push 006E0410h
  loc_00E23229: push edi
  loc_00E2322A: push eax
  loc_00E2322B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23231: lea ecx, var_18
  loc_00E23234: call [00401258h] ; __vbaFreeStr
  loc_00E2323A: lea ecx, var_20
  loc_00E2323D: call [00401254h] ; __vbaFreeObj
  loc_00E23243: mov eax, [00E3D060h]
  loc_00E23248: test eax, eax
  loc_00E2324A: jnz 00E23261h
  loc_00E2324C: push 00E3D060h
  loc_00E23251: push 006D87C0h
  loc_00E23256: call [004011A0h] ; __vbaNew2
  loc_00E2325C: mov eax, [00E3D060h]
  loc_00E23261: mov ecx, [eax]
  loc_00E23263: push eax
  loc_00E23264: call [ecx+000003FCh]
  loc_00E2326A: lea edx, var_20
  loc_00E2326D: push eax
  loc_00E2326E: push edx
  loc_00E2326F: call __vbaObjSet
  loc_00E23271: mov edi, eax
  loc_00E23273: push 00023668h
  loc_00E23278: mov ebx, [edi]
  loc_00E2327A: call [00401018h] ; __vbaStrI4
  loc_00E23280: mov edx, eax
  loc_00E23282: lea ecx, var_18
  loc_00E23285: call [00401228h] ; __vbaStrMove
  loc_00E2328B: push eax
  loc_00E2328C: push edi
  loc_00E2328D: call [ebx+00000054h]
  loc_00E23290: test eax, eax
  loc_00E23292: fnclex
  loc_00E23294: jge 00E232A5h
  loc_00E23296: push 00000054h
  loc_00E23298: push 006E0410h
  loc_00E2329D: push edi
  loc_00E2329E: push eax
  loc_00E2329F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E232A5: lea ecx, var_18
  loc_00E232A8: call [00401258h] ; __vbaFreeStr
  loc_00E232AE: lea ecx, var_20
  loc_00E232B1: call [00401254h] ; __vbaFreeObj
  loc_00E232B7: mov eax, [00E3D060h]
  loc_00E232BC: test eax, eax
  loc_00E232BE: jnz 00E232D5h
  loc_00E232C0: push 00E3D060h
  loc_00E232C5: push 006D87C0h
  loc_00E232CA: call [004011A0h] ; __vbaNew2
  loc_00E232D0: mov eax, [00E3D060h]
  loc_00E232D5: mov ecx, [eax]
  loc_00E232D7: push eax
  loc_00E232D8: call [ecx+000003ECh]
  loc_00E232DE: lea edx, var_20
  loc_00E232E1: push eax
  loc_00E232E2: push edx
  loc_00E232E3: call __vbaObjSet
  loc_00E232E5: mov edi, eax
  loc_00E232E7: push 00030D40h
  loc_00E232EC: mov ebx, [edi]
  loc_00E232EE: call [00401018h] ; __vbaStrI4
  loc_00E232F4: mov edx, eax
  loc_00E232F6: lea ecx, var_18
  loc_00E232F9: call [00401228h] ; __vbaStrMove
  loc_00E232FF: push eax
  loc_00E23300: push edi
  loc_00E23301: call [ebx+00000054h]
  loc_00E23304: test eax, eax
  loc_00E23306: fnclex
  loc_00E23308: jge 00E23319h
  loc_00E2330A: push 00000054h
  loc_00E2330C: push 006E0410h
  loc_00E23311: push edi
  loc_00E23312: push eax
  loc_00E23313: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23319: lea ecx, var_18
  loc_00E2331C: call [00401258h] ; __vbaFreeStr
  loc_00E23322: lea ecx, var_20
  loc_00E23325: call [00401254h] ; __vbaFreeObj
  loc_00E2332B: mov eax, [00E3D060h]
  loc_00E23330: test eax, eax
  loc_00E23332: jnz 00E23349h
  loc_00E23334: push 00E3D060h
  loc_00E23339: push 006D87C0h
  loc_00E2333E: call [004011A0h] ; __vbaNew2
  loc_00E23344: mov eax, [00E3D060h]
  loc_00E23349: mov ecx, [eax]
  loc_00E2334B: push eax
  loc_00E2334C: call [ecx+000003DCh]
  loc_00E23352: lea edx, var_20
  loc_00E23355: push eax
  loc_00E23356: push edx
  loc_00E23357: call __vbaObjSet
  loc_00E23359: mov edi, eax
  loc_00E2335B: push 0000FDE8h
  loc_00E23360: mov ebx, [edi]
  loc_00E23362: call [00401018h] ; __vbaStrI4
  loc_00E23368: mov edx, eax
  loc_00E2336A: lea ecx, var_18
  loc_00E2336D: call [00401228h] ; __vbaStrMove
  loc_00E23373: push eax
  loc_00E23374: push edi
  loc_00E23375: call [ebx+00000054h]
  loc_00E23378: test eax, eax
  loc_00E2337A: fnclex
  loc_00E2337C: jge 00E2338Dh
  loc_00E2337E: push 00000054h
  loc_00E23380: push 006E0410h
  loc_00E23385: push edi
  loc_00E23386: push eax
  loc_00E23387: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2338D: lea ecx, var_18
  loc_00E23390: call [00401258h] ; __vbaFreeStr
  loc_00E23396: lea ecx, var_20
  loc_00E23399: call [00401254h] ; __vbaFreeObj
  loc_00E2339F: mov eax, [00E3D060h]
  loc_00E233A4: test eax, eax
  loc_00E233A6: jnz 00E233BDh
  loc_00E233A8: push 00E3D060h
  loc_00E233AD: push 006D87C0h
  loc_00E233B2: call [004011A0h] ; __vbaNew2
  loc_00E233B8: mov eax, [00E3D060h]
  loc_00E233BD: mov ecx, [eax]
  loc_00E233BF: push eax
  loc_00E233C0: call [ecx+000003CCh]
  loc_00E233C6: lea edx, var_20
  loc_00E233C9: push eax
  loc_00E233CA: push edx
  loc_00E233CB: call __vbaObjSet
  loc_00E233CD: mov edi, eax
  loc_00E233CF: push 0001E848h
  loc_00E233D4: mov ebx, [edi]
  loc_00E233D6: call [00401018h] ; __vbaStrI4
  loc_00E233DC: mov edx, eax
  loc_00E233DE: lea ecx, var_18
  loc_00E233E1: call [00401228h] ; __vbaStrMove
  loc_00E233E7: push eax
  loc_00E233E8: push edi
  loc_00E233E9: call [ebx+00000054h]
  loc_00E233EC: test eax, eax
  loc_00E233EE: fnclex
  loc_00E233F0: jge 00E23401h
  loc_00E233F2: push 00000054h
  loc_00E233F4: push 006E0410h
  loc_00E233F9: push edi
  loc_00E233FA: push eax
  loc_00E233FB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23401: lea ecx, var_18
  loc_00E23404: call [00401258h] ; __vbaFreeStr
  loc_00E2340A: lea ecx, var_20
  loc_00E2340D: call [00401254h] ; __vbaFreeObj
  loc_00E23413: mov eax, [00E3D060h]
  loc_00E23418: test eax, eax
  loc_00E2341A: jnz 00E23431h
  loc_00E2341C: push 00E3D060h
  loc_00E23421: push 006D87C0h
  loc_00E23426: call [004011A0h] ; __vbaNew2
  loc_00E2342C: mov eax, [00E3D060h]
  loc_00E23431: mov ecx, [eax]
  loc_00E23433: push eax
  loc_00E23434: call [ecx+000003BCh]
  loc_00E2343A: lea edx, var_20
  loc_00E2343D: push eax
  loc_00E2343E: push edx
  loc_00E2343F: call __vbaObjSet
  loc_00E23441: mov edi, eax
  loc_00E23443: push 00023668h
  loc_00E23448: mov ebx, [edi]
  loc_00E2344A: call [00401018h] ; __vbaStrI4
  loc_00E23450: mov edx, eax
  loc_00E23452: lea ecx, var_18
  loc_00E23455: call [00401228h] ; __vbaStrMove
  loc_00E2345B: push eax
  loc_00E2345C: push edi
  loc_00E2345D: call [ebx+00000054h]
  loc_00E23460: test eax, eax
  loc_00E23462: fnclex
  loc_00E23464: jge 00E1FA6Bh
  loc_00E2346A: jmp 00E1FA5Ch
  loc_00E2346F: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E23475: mov ecx, 80020004h
  loc_00E2347A: mov var_68, ecx
  loc_00E2347D: mov eax, 0000000Ah
  loc_00E23482: mov var_58, ecx
  loc_00E23485: mov ebx, 00000008h
  loc_00E2348A: lea edx, var_90
  loc_00E23490: lea ecx, var_50
  loc_00E23493: mov var_70, eax
  loc_00E23496: mov var_60, eax
  loc_00E23499: mov var_88, 006E15D4h ; "Not Valid"
  loc_00E234A3: mov var_90, ebx
  loc_00E234A9: call edi
  loc_00E234AB: lea edx, var_80
  loc_00E234AE: lea ecx, var_40
  loc_00E234B1: mov var_78, 006E1550h ; "Data Calon Peserta Didik Tidak Valid, atau Mengalami [CRASH], !"
  loc_00E234B8: mov var_80, ebx
  loc_00E234BB: call edi
  loc_00E234BD: lea eax, var_70
  loc_00E234C0: lea ecx, var_60
  loc_00E234C3: push eax
  loc_00E234C4: lea edx, var_50
  loc_00E234C7: push ecx
  loc_00E234C8: push edx
  loc_00E234C9: lea eax, var_40
  loc_00E234CC: push 00000010h
  loc_00E234CE: push eax
  loc_00E234CF: call [004010A8h] ; rtcMsgBox
  loc_00E234D5: lea ecx, var_70
  loc_00E234D8: lea edx, var_60
  loc_00E234DB: push ecx
  loc_00E234DC: lea eax, var_50
  loc_00E234DF: push edx
  loc_00E234E0: lea ecx, var_40
  loc_00E234E3: push eax
  loc_00E234E4: push ecx
  loc_00E234E5: push 00000004h
  loc_00E234E7: call [00401038h] ; __vbaFreeVarList
  loc_00E234ED: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E234F3: add esp, 00000014h
  loc_00E234F6: mov eax, [00E3D060h]
  loc_00E234FB: mov var_78, FFFFFFFFh
  loc_00E23502: test eax, eax
  loc_00E23504: mov var_80, 0000000Bh
  loc_00E2350B: jnz 00E23526h
  loc_00E2350D: mov ebx, [004011A0h] ; __vbaNew2
  loc_00E23513: push 00E3D060h
  loc_00E23518: push 006D87C0h
  loc_00E2351D: call ebx
  loc_00E2351F: mov eax, [00E3D060h]
  loc_00E23524: jmp 00E2352Ch
  loc_00E23526: mov ebx, [004011A0h] ; __vbaNew2
  loc_00E2352C: mov ecx, var_80
  loc_00E2352F: sub esp, 00000010h
  loc_00E23532: mov edx, esp
  loc_00E23534: push 8001000Dh
  loc_00E23539: push eax
  loc_00E2353A: mov [edx], ecx
  loc_00E2353C: mov ecx, var_7C
  loc_00E2353F: mov [edx+00000004h], ecx
  loc_00E23542: mov ecx, var_78
  loc_00E23545: mov [edx+00000008h], ecx
  loc_00E23548: mov ecx, var_74
  loc_00E2354B: mov [edx+0000000Ch], ecx
  loc_00E2354E: mov edx, [eax]
  loc_00E23550: call [edx+000004C8h]
  loc_00E23556: push eax
  loc_00E23557: lea eax, var_20
  loc_00E2355A: push eax
  loc_00E2355B: call __vbaObjSet
  loc_00E2355D: push eax
  loc_00E2355E: call [00401238h] ; __vbaLateIdSt
  loc_00E23564: lea ecx, var_20
  loc_00E23567: call edi
  loc_00E23569: mov eax, [00E3D060h]
  loc_00E2356E: test eax, eax
  loc_00E23570: jnz 00E23583h
  loc_00E23572: push 00E3D060h
  loc_00E23577: push 006D87C0h
  loc_00E2357C: call ebx
  loc_00E2357E: mov eax, [00E3D060h]
  loc_00E23583: mov ecx, [eax]
  loc_00E23585: push eax
  loc_00E23586: call [ecx+000004A0h]
  loc_00E2358C: lea edx, var_20
  loc_00E2358F: push eax
  loc_00E23590: push edx
  loc_00E23591: call __vbaObjSet
  loc_00E23593: mov esi, eax
  loc_00E23595: push 00000000h
  loc_00E23597: push esi
  loc_00E23598: mov eax, [esi]
  loc_00E2359A: call [eax+0000008Ch]
  loc_00E235A0: test eax, eax
  loc_00E235A2: fnclex
  loc_00E235A4: jge 00E235B8h
  loc_00E235A6: push 0000008Ch
  loc_00E235AB: push 006DCDA0h
  loc_00E235B0: push esi
  loc_00E235B1: push eax
  loc_00E235B2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E235B8: lea ecx, var_20
  loc_00E235BB: call edi
  loc_00E235BD: mov eax, [00E3D9CCh]
  loc_00E235C2: test eax, eax
  loc_00E235C4: jnz 00E235D2h
  loc_00E235C6: push 00E3D9CCh
  loc_00E235CB: push 006DC960h
  loc_00E235D0: call ebx
  loc_00E235D2: mov ecx, Me
  loc_00E235D5: mov esi, [00E3D9CCh]
  loc_00E235DB: lea edx, var_20
  loc_00E235DE: push ecx
  loc_00E235DF: mov ebx, [esi]
  loc_00E235E1: push edx
  loc_00E235E2: call [004010B4h] ; __vbaObjSetAddref
  loc_00E235E8: push eax
  loc_00E235E9: push esi
  loc_00E235EA: call [ebx+00000010h]
  loc_00E235ED: test eax, eax
  loc_00E235EF: fnclex
  loc_00E235F1: jge 00E23602h
  loc_00E235F3: push 00000010h
  loc_00E235F5: push 006DC950h
  loc_00E235FA: push esi
  loc_00E235FB: push eax
  loc_00E235FC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E23602: lea ecx, var_20
  loc_00E23605: call edi
  loc_00E23607: mov var_4, 00000000h
  loc_00E2360E: push 00E2365Eh
  loc_00E23613: jmp 00E2365Dh
  loc_00E23615: lea eax, var_1C
  loc_00E23618: lea ecx, var_18
  loc_00E2361B: push eax
  loc_00E2361C: push ecx
  loc_00E2361D: push 00000002h
  loc_00E2361F: call [004011BCh] ; __vbaFreeStrList
  loc_00E23625: lea edx, var_30
  loc_00E23628: lea eax, var_2C
  loc_00E2362B: push edx
  loc_00E2362C: lea ecx, var_28
  loc_00E2362F: push eax
  loc_00E23630: lea edx, var_24
  loc_00E23633: push ecx
  loc_00E23634: lea eax, var_20
  loc_00E23637: push edx
  loc_00E23638: push eax
  loc_00E23639: push 00000005h
  loc_00E2363B: call [00401048h] ; __vbaFreeObjList
  loc_00E23641: lea ecx, var_70
  loc_00E23644: lea edx, var_60
  loc_00E23647: push ecx
  loc_00E23648: lea eax, var_50
  loc_00E2364B: push edx
  loc_00E2364C: lea ecx, var_40
  loc_00E2364F: push eax
  loc_00E23650: push ecx
  loc_00E23651: push 00000004h
  loc_00E23653: call [00401038h] ; __vbaFreeVarList
  loc_00E23659: add esp, 00000038h
  loc_00E2365C: ret
  loc_00E2365D: ret
  loc_00E2365E: mov eax, Me
  loc_00E23661: push eax
  loc_00E23662: mov edx, [eax]
  loc_00E23664: call [edx+00000008h]
  loc_00E23667: mov eax, var_4
  loc_00E2366A: mov ecx, var_14
  loc_00E2366D: pop edi
  loc_00E2366E: pop esi
  loc_00E2366F: mov fs:[00000000h], ecx
  loc_00E23676: pop ebx
  loc_00E23677: mov esp, ebp
  loc_00E23679: pop ebp
  loc_00E2367A: retn 0004h
End Sub
