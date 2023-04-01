VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form cetakbio
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
  ClientWidth = 11685
  ClientHeight = 5760
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame ftk
    Caption = "Frame1"
    BackColor = &H404040&
    Left = 8760
    Top = 2370
    Width = 2505
    Height = 1845
    TabIndex = 12
    BorderStyle = 0 'None
    Begin Project1.jcbutton jcbutton3
      Left = -30
      Top = 480
      Width = 2565
      Height = 705
      TabIndex = 15
      OleObjectBlob = "cetakbio.frx":0000
    End
    Begin Project1.jcbutton jcbutton4
      Left = -30
      Top = 1170
      Width = 2565
      Height = 705
      TabIndex = 16
      OleObjectBlob = "cetakbio.frx":0208
    End
    Begin VB.Shape Shape6
      Left = 0
      Top = 420
      Width = 14685
      Height = 75
      BorderStyle = 0 'Transparent
      FillColor = &HC0C0C0&
      FillStyle = 0
    End
    Begin VB.Label Label2
      Caption = "-PILIH-"
      ForeColor = &H80FF80&
      Left = 840
      Top = 30
      Width = 1305
      Height = 315
      TabIndex = 13
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
  End
  Begin Project1.jcbutton ctk
    Left = 90
    Top = 2730
    Width = 11475
    Height = 555
    TabIndex = 11
    OleObjectBlob = "cetakbio.frx":043C
  End
  Begin VB.OptionButton optjur
    Caption = "Jurusan"
    BackColor = &HE0E0E0&
    ForeColor = &H404040&
    Left = 7710
    Top = 1410
    Width = 1515
    Height = 300
    TabIndex = 9
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
  Begin VB.Frame fcari
    Caption = "Cetak Berdasarkan"
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
    Begin VB.OptionButton optagama
      Caption = "Agama"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 9300
      Top = 600
      Width = 1695
      Height = 300
      TabIndex = 10
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
    Begin VB.OptionButton optasal
      Caption = "Asal Sekolah"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 5580
      Top = 540
      Width = 1695
      Height = 300
      TabIndex = 8
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
      Left = 3930
      Top = 540
      Width = 1455
      Height = 300
      TabIndex = 7
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
      TabIndex = 3
      BorderStyle = 0 'None
    End
    Begin VB.OptionButton optno
      Caption = "No. Daftar"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 2160
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
    Begin VB.OptionButton optnama
      Caption = "Nama"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 900
      Top = 510
      Width = 1455
      Height = 300
      TabIndex = 1
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
    Begin Project1.jcbutton batal
      Left = 5580
      Top = 870
      Width = 675
      Height = 225
      TabIndex = 14
      OleObjectBlob = "cetakbio.frx":0644
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
    OleObjectBlob = "cetakbio.frx":0878
  End
  Begin MSDataGridLib.DataGrid dgpen
    Left = 90
    Top = 3450
    Width = 11475
    Height = 2265
    TabIndex = 4
    OleObjectBlob = "cetakbio.frx":0BAA
  End
  Begin VB.Shape Shape33
    BorderColor = &HE0E0E0&
    Left = -60
    Top = 2640
    Width = 15015
    Height = 750
    BorderStyle = 3 'Dot
    FillColor = &HFFFF&
  End
  Begin VB.Label Label1
    Caption = "-BIODATA-"
    ForeColor = &H4000&
    Left = 5460
    Top = 510
    Width = 1305
    Height = 315
    TabIndex = 6
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
    Picture = "cetakbio.frx":0D55
    Left = 30
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Label Label4
    Caption = "Print System"
    ForeColor = &H80FF80&
    Left = 5370
    Top = 30
    Width = 1815
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
  Begin VB.Shape Shape4
    Left = 0
    Top = 510
    Width = 14685
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
End

Attribute VB_Name = "cetakbio"


Private Sub optnibu_Click() 'E351A0
  loc_00E351A0: push ebp
  loc_00E351A1: mov ebp, esp
  loc_00E351A3: sub esp, 0000000Ch
  loc_00E351A6: push 00402836h ; __vbaExceptHandler
  loc_00E351AB: mov eax, fs:[00000000h]
  loc_00E351B1: push eax
  loc_00E351B2: mov fs:[00000000h], esp
  loc_00E351B9: sub esp, 00000048h
  loc_00E351BC: push ebx
  loc_00E351BD: push esi
  loc_00E351BE: push edi
  loc_00E351BF: mov var_C, esp
  loc_00E351C2: mov var_8, 004026A0h
  loc_00E351C9: mov esi, Me
  loc_00E351CC: mov eax, esi
  loc_00E351CE: and eax, 00000001h
  loc_00E351D1: mov var_4, eax
  loc_00E351D4: and esi, FFFFFFFEh
  loc_00E351D7: push esi
  loc_00E351D8: mov Me, esi
  loc_00E351DB: mov ecx, [esi]
  loc_00E351DD: call [ecx+00000004h]
  loc_00E351E0: sub esp, 00000010h
  loc_00E351E3: mov ecx, 0000000Bh
  loc_00E351E8: mov edx, esp
  loc_00E351EA: xor eax, eax
  loc_00E351EC: mov var_18, eax
  loc_00E351EF: mov var_1C, eax
  loc_00E351F2: mov [edx], ecx
  loc_00E351F4: mov ecx, var_38
  loc_00E351F7: mov var_2C, eax
  loc_00E351FA: or eax, FFFFFFFFh
  loc_00E351FD: mov [edx+00000004h], ecx
  loc_00E35200: mov ecx, [esi]
  loc_00E35202: push 80010007h
  loc_00E35207: push esi
  loc_00E35208: mov [edx+00000008h], eax
  loc_00E3520B: mov eax, var_30
  loc_00E3520E: mov [edx+0000000Ch], eax
  loc_00E35211: call [ecx+00000334h]
  loc_00E35217: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3521D: lea edx, var_18
  loc_00E35220: push eax
  loc_00E35221: push edx
  loc_00E35222: call edi
  loc_00E35224: push eax
  loc_00E35225: call [00401238h] ; __vbaLateIdSt
  loc_00E3522B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E35231: lea ecx, var_18
  loc_00E35234: call ebx
  loc_00E35236: mov eax, [esi]
  loc_00E35238: push esi
  loc_00E35239: call [eax+00000328h]
  loc_00E3523F: lea ecx, var_18
  loc_00E35242: push eax
  loc_00E35243: push ecx
  loc_00E35244: call edi
  loc_00E35246: mov edx, [eax]
  loc_00E35248: push eax
  loc_00E35249: mov var_50, eax
  loc_00E3524C: call [edx+00000204h]
  loc_00E35252: test eax, eax
  loc_00E35254: fnclex
  loc_00E35256: jge 00E3526Dh
  loc_00E35258: mov ecx, var_50
  loc_00E3525B: push 00000204h
  loc_00E35260: push 006DCB70h
  loc_00E35265: push ecx
  loc_00E35266: push eax
  loc_00E35267: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3526D: lea ecx, var_18
  loc_00E35270: call ebx
  loc_00E35272: mov edx, [esi]
  loc_00E35274: push esi
  loc_00E35275: call [edx+0000032Ch]
  loc_00E3527B: push eax
  loc_00E3527C: lea eax, var_18
  loc_00E3527F: push eax
  loc_00E35280: call edi
  loc_00E35282: mov ecx, [eax]
  loc_00E35284: push 00000000h
  loc_00E35286: push eax
  loc_00E35287: mov var_50, eax
  loc_00E3528A: call [ecx+0000009Ch]
  loc_00E35290: test eax, eax
  loc_00E35292: fnclex
  loc_00E35294: jge 00E352ABh
  loc_00E35296: mov edx, var_50
  loc_00E35299: push 0000009Ch
  loc_00E3529E: push 006E03D4h
  loc_00E352A3: push edx
  loc_00E352A4: push eax
  loc_00E352A5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E352AB: lea ecx, var_18
  loc_00E352AE: call ebx
  loc_00E352B0: mov eax, [esi]
  loc_00E352B2: push esi
  loc_00E352B3: call [eax+00000330h]
  loc_00E352B9: lea ecx, var_18
  loc_00E352BC: push eax
  loc_00E352BD: push ecx
  loc_00E352BE: call edi
  loc_00E352C0: mov edx, [eax]
  loc_00E352C2: push 00000000h
  loc_00E352C4: push eax
  loc_00E352C5: mov var_50, eax
  loc_00E352C8: call [edx+0000009Ch]
  loc_00E352CE: test eax, eax
  loc_00E352D0: fnclex
  loc_00E352D2: jge 00E352E9h
  loc_00E352D4: mov ecx, var_50
  loc_00E352D7: push 0000009Ch
  loc_00E352DC: push 006E03D4h
  loc_00E352E1: push ecx
  loc_00E352E2: push eax
  loc_00E352E3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E352E9: lea ecx, var_18
  loc_00E352EC: call ebx
  loc_00E352EE: mov edx, [esi]
  loc_00E352F0: push esi
  loc_00E352F1: call [edx+0000031Ch]
  loc_00E352F7: push eax
  loc_00E352F8: lea eax, var_18
  loc_00E352FB: push eax
  loc_00E352FC: call edi
  loc_00E352FE: mov ecx, [eax]
  loc_00E35300: push 00000000h
  loc_00E35302: push eax
  loc_00E35303: mov var_50, eax
  loc_00E35306: call [ecx+0000009Ch]
  loc_00E3530C: test eax, eax
  loc_00E3530E: fnclex
  loc_00E35310: jge 00E35327h
  loc_00E35312: mov edx, var_50
  loc_00E35315: push 0000009Ch
  loc_00E3531A: push 006E03D4h
  loc_00E3531F: push edx
  loc_00E35320: push eax
  loc_00E35321: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35327: lea ecx, var_18
  loc_00E3532A: call ebx
  loc_00E3532C: mov eax, [esi]
  loc_00E3532E: push esi
  loc_00E3532F: call [eax+00000314h]
  loc_00E35335: lea ecx, var_18
  loc_00E35338: push eax
  loc_00E35339: push ecx
  loc_00E3533A: call edi
  loc_00E3533C: mov edx, [eax]
  loc_00E3533E: push 00000000h
  loc_00E35340: push eax
  loc_00E35341: mov var_50, eax
  loc_00E35344: call [edx+0000009Ch]
  loc_00E3534A: test eax, eax
  loc_00E3534C: fnclex
  loc_00E3534E: jge 00E35365h
  loc_00E35350: mov ecx, var_50
  loc_00E35353: push 0000009Ch
  loc_00E35358: push 006E03D4h
  loc_00E3535D: push ecx
  loc_00E3535E: push eax
  loc_00E3535F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35365: lea ecx, var_18
  loc_00E35368: call ebx
  loc_00E3536A: mov edx, [esi]
  loc_00E3536C: push esi
  loc_00E3536D: call [edx+00000320h]
  loc_00E35373: push eax
  loc_00E35374: lea eax, var_18
  loc_00E35377: push eax
  loc_00E35378: call edi
  loc_00E3537A: mov ecx, [eax]
  loc_00E3537C: push 00000000h
  loc_00E3537E: push eax
  loc_00E3537F: mov var_50, eax
  loc_00E35382: call [ecx+0000009Ch]
  loc_00E35388: test eax, eax
  loc_00E3538A: fnclex
  loc_00E3538C: jge 00E353A3h
  loc_00E3538E: mov edx, var_50
  loc_00E35391: push 0000009Ch
  loc_00E35396: push 006E03D4h
  loc_00E3539B: push edx
  loc_00E3539C: push eax
  loc_00E3539D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E353A3: lea ecx, var_18
  loc_00E353A6: call ebx
  loc_00E353A8: sub esp, 00000010h
  loc_00E353AB: mov ecx, 00000008h
  loc_00E353B0: mov edx, esp
  loc_00E353B2: mov eax, 006E0A78h ; "biodata"
  loc_00E353B7: push 0000000Eh
  loc_00E353B9: push esi
  loc_00E353BA: mov [edx], ecx
  loc_00E353BC: mov ecx, var_38
  loc_00E353BF: mov [edx+00000004h], ecx
  loc_00E353C2: mov ecx, [esi]
  loc_00E353C4: mov [edx+00000008h], eax
  loc_00E353C7: mov eax, var_30
  loc_00E353CA: mov [edx+0000000Ch], eax
  loc_00E353CD: call [ecx+00000358h]
  loc_00E353D3: lea edx, var_18
  loc_00E353D6: push eax
  loc_00E353D7: push edx
  loc_00E353D8: call edi
  loc_00E353DA: push eax
  loc_00E353DB: call [00401238h] ; __vbaLateIdSt
  loc_00E353E1: lea ecx, var_18
  loc_00E353E4: call ebx
  loc_00E353E6: mov eax, [esi]
  loc_00E353E8: push 00000000h
  loc_00E353EA: push 00000019h
  loc_00E353EC: push esi
  loc_00E353ED: call [eax+00000358h]
  loc_00E353F3: lea ecx, var_18
  loc_00E353F6: push eax
  loc_00E353F7: push ecx
  loc_00E353F8: call edi
  loc_00E353FA: push eax
  loc_00E353FB: call [00401030h] ; __vbaLateIdCall
  loc_00E35401: add esp, 0000000Ch
  loc_00E35404: lea ecx, var_18
  loc_00E35407: call ebx
  loc_00E35409: mov edx, [esi]
  loc_00E3540B: push 006E05D8h
  loc_00E35410: push esi
  loc_00E35411: call [edx+00000358h]
  loc_00E35417: push eax
  loc_00E35418: lea eax, var_18
  loc_00E3541B: push eax
  loc_00E3541C: call edi
  loc_00E3541E: push eax
  loc_00E3541F: call [00401224h] ; __vbaCastObj
  loc_00E35425: lea ecx, var_2C
  loc_00E35428: mov var_24, eax
  loc_00E3542B: push ecx
  loc_00E3542C: mov var_2C, 0000000Dh
  loc_00E35433: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E35439: mov ecx, [eax]
  loc_00E3543B: sub esp, 00000010h
  loc_00E3543E: mov edx, esp
  loc_00E35440: push 00000000h
  loc_00E35442: push 0000002Ah
  loc_00E35444: mov [edx], ecx
  loc_00E35446: mov ecx, [eax+00000004h]
  loc_00E35449: push esi
  loc_00E3544A: mov [edx+00000004h], ecx
  loc_00E3544D: mov ecx, [eax+00000008h]
  loc_00E35450: mov eax, [eax+0000000Ch]
  loc_00E35453: mov [edx+00000008h], ecx
  loc_00E35456: mov ecx, [esi]
  loc_00E35458: mov [edx+0000000Ch], eax
  loc_00E3545B: call [ecx+0000035Ch]
  loc_00E35461: lea edx, var_1C
  loc_00E35464: push eax
  loc_00E35465: push edx
  loc_00E35466: call edi
  loc_00E35468: push eax
  loc_00E35469: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3546F: lea eax, var_1C
  loc_00E35472: lea ecx, var_18
  loc_00E35475: push eax
  loc_00E35476: push ecx
  loc_00E35477: push 00000002h
  loc_00E35479: call [00401048h] ; __vbaFreeObjList
  loc_00E3547F: add esp, 00000028h
  loc_00E35482: lea ecx, var_2C
  loc_00E35485: call [00401024h] ; __vbaFreeVar
  loc_00E3548B: sub esp, 00000010h
  loc_00E3548E: mov ecx, 0000000Bh
  loc_00E35493: mov edx, esp
  loc_00E35495: xor eax, eax
  loc_00E35497: push 00000006h
  loc_00E35499: push esi
  loc_00E3549A: mov [edx], ecx
  loc_00E3549C: mov ecx, var_38
  loc_00E3549F: mov [edx+00000004h], ecx
  loc_00E354A2: mov ecx, [esi]
  loc_00E354A4: mov [edx+00000008h], eax
  loc_00E354A7: mov eax, var_30
  loc_00E354AA: mov [edx+0000000Ch], eax
  loc_00E354AD: call [ecx+0000035Ch]
  loc_00E354B3: lea edx, var_18
  loc_00E354B6: push eax
  loc_00E354B7: push edx
  loc_00E354B8: call edi
  loc_00E354BA: push eax
  loc_00E354BB: call [00401238h] ; __vbaLateIdSt
  loc_00E354C1: lea ecx, var_18
  loc_00E354C4: call ebx
  loc_00E354C6: sub esp, 00000010h
  loc_00E354C9: mov ecx, 0000000Bh
  loc_00E354CE: mov edx, esp
  loc_00E354D0: xor eax, eax
  loc_00E354D2: push 8001000Eh
  loc_00E354D7: push esi
  loc_00E354D8: mov [edx], ecx
  loc_00E354DA: mov ecx, var_38
  loc_00E354DD: mov [edx+00000004h], ecx
  loc_00E354E0: mov ecx, [esi]
  loc_00E354E2: mov [edx+00000008h], eax
  loc_00E354E5: mov eax, var_30
  loc_00E354E8: mov [edx+0000000Ch], eax
  loc_00E354EB: call [ecx+0000035Ch]
  loc_00E354F1: lea edx, var_18
  loc_00E354F4: push eax
  loc_00E354F5: push edx
  loc_00E354F6: call edi
  loc_00E354F8: push eax
  loc_00E354F9: call [00401238h] ; __vbaLateIdSt
  loc_00E354FF: lea ecx, var_18
  loc_00E35502: call ebx
  loc_00E35504: mov eax, [esi]
  loc_00E35506: push 00000000h
  loc_00E35508: push FFFFFDDAh
  loc_00E3550D: push esi
  loc_00E3550E: call [eax+0000035Ch]
  loc_00E35514: lea ecx, var_18
  loc_00E35517: push eax
  loc_00E35518: push ecx
  loc_00E35519: call edi
  loc_00E3551B: push eax
  loc_00E3551C: call [00401030h] ; __vbaLateIdCall
  loc_00E35522: add esp, 0000000Ch
  loc_00E35525: lea ecx, var_18
  loc_00E35528: call ebx
  loc_00E3552A: mov var_4, 00000000h
  loc_00E35531: push 00E35556h
  loc_00E35536: jmp 00E35555h
  loc_00E35538: lea edx, var_1C
  loc_00E3553B: lea eax, var_18
  loc_00E3553E: push edx
  loc_00E3553F: push eax
  loc_00E35540: push 00000002h
  loc_00E35542: call [00401048h] ; __vbaFreeObjList
  loc_00E35548: add esp, 0000000Ch
  loc_00E3554B: lea ecx, var_2C
  loc_00E3554E: call [00401024h] ; __vbaFreeVar
  loc_00E35554: ret
  loc_00E35555: ret
  loc_00E35556: mov eax, Me
  loc_00E35559: push eax
  loc_00E3555A: mov ecx, [eax]
  loc_00E3555C: call [ecx+00000008h]
  loc_00E3555F: mov eax, var_4
  loc_00E35562: mov ecx, var_14
  loc_00E35565: pop edi
  loc_00E35566: pop esi
  loc_00E35567: mov fs:[00000000h], ecx
  loc_00E3556E: pop ebx
  loc_00E3556F: mov esp, ebp
  loc_00E35571: pop ebp
  loc_00E35572: retn 0004h
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer) 'E35960
  loc_00E35960: push ebp
  loc_00E35961: mov ebp, esp
  loc_00E35963: sub esp, 0000000Ch
  loc_00E35966: push 00402836h ; __vbaExceptHandler
  loc_00E3596B: mov eax, fs:[00000000h]
  loc_00E35971: push eax
  loc_00E35972: mov fs:[00000000h], esp
  loc_00E35979: sub esp, 000000B4h
  loc_00E3597F: push ebx
  loc_00E35980: push esi
  loc_00E35981: push edi
  loc_00E35982: mov var_C, esp
  loc_00E35985: mov var_8, 004026C0h
  loc_00E3598C: mov esi, Me
  loc_00E3598F: mov eax, esi
  loc_00E35991: and eax, 00000001h
  loc_00E35994: mov var_4, eax
  loc_00E35997: and esi, FFFFFFFEh
  loc_00E3599A: push esi
  loc_00E3599B: mov Me, esi
  loc_00E3599E: mov ecx, [esi]
  loc_00E359A0: call [ecx+00000004h]
  loc_00E359A3: mov edx, [esi]
  loc_00E359A5: xor ebx, ebx
  loc_00E359A7: push esi
  loc_00E359A8: mov var_24, ebx
  loc_00E359AB: mov var_28, ebx
  loc_00E359AE: mov var_2C, ebx
  loc_00E359B1: mov var_30, ebx
  loc_00E359B4: mov var_40, ebx
  loc_00E359B7: mov var_50, ebx
  loc_00E359BA: mov var_60, ebx
  loc_00E359BD: mov var_70, ebx
  loc_00E359C0: mov var_80, ebx
  loc_00E359C3: mov var_90, ebx
  loc_00E359C9: mov var_B4, ebx
  loc_00E359CF: call [edx+0000032Ch]
  loc_00E359D5: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E359DB: push eax
  loc_00E359DC: lea eax, var_30
  loc_00E359DF: push eax
  loc_00E359E0: call edi
  loc_00E359E2: mov ecx, [eax]
  loc_00E359E4: lea edx, var_B4
  loc_00E359EA: push edx
  loc_00E359EB: push eax
  loc_00E359EC: mov var_B8, eax
  loc_00E359F2: call [ecx+000000E0h]
  loc_00E359F8: cmp eax, ebx
  loc_00E359FA: fnclex
  loc_00E359FC: jge 00E35A16h
  loc_00E359FE: mov ecx, var_B8
  loc_00E35A04: push 000000E0h
  loc_00E35A09: push 006E03D4h
  loc_00E35A0E: push ecx
  loc_00E35A0F: push eax
  loc_00E35A10: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35A16: mov edx, var_B4
  loc_00E35A1C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E35A22: lea ecx, var_30
  loc_00E35A25: mov var_C0, edx
  loc_00E35A2B: call ebx
  loc_00E35A2D: cmp var_C0, 0000h
  loc_00E35A35: jz 00E35A9Dh
  loc_00E35A37: mov eax, [esi]
  loc_00E35A39: push esi
  loc_00E35A3A: call [eax+00000328h]
  loc_00E35A40: lea ecx, var_30
  loc_00E35A43: push eax
  loc_00E35A44: push ecx
  loc_00E35A45: call edi
  loc_00E35A47: mov edx, [eax]
  loc_00E35A49: lea ecx, var_28
  loc_00E35A4C: push ecx
  loc_00E35A4D: push eax
  loc_00E35A4E: mov var_B8, eax
  loc_00E35A54: call [edx+000000A0h]
  loc_00E35A5A: test eax, eax
  loc_00E35A5C: fnclex
  loc_00E35A5E: jge 00E35A78h
  loc_00E35A60: mov edx, var_B8
  loc_00E35A66: push 000000A0h
  loc_00E35A6B: push 006DCB70h
  loc_00E35A70: push edx
  loc_00E35A71: push eax
  loc_00E35A72: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35A78: mov eax, var_28
  loc_00E35A7B: push 006E0C48h ; "select * from biodata where no_daftar like '" & Chr(37)
  loc_00E35A80: push eax
  loc_00E35A81: call [0040106Ch] ; __vbaStrCat
  loc_00E35A87: mov edx, eax
  loc_00E35A89: lea ecx, var_2C
  loc_00E35A8C: call [00401228h] ; __vbaStrMove
  loc_00E35A92: push eax
  loc_00E35A93: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E35A98: jmp 00E35E75h
  loc_00E35A9D: mov edx, [esi]
  loc_00E35A9F: push esi
  loc_00E35AA0: call [edx+00000330h]
  loc_00E35AA6: push eax
  loc_00E35AA7: lea eax, var_30
  loc_00E35AAA: push eax
  loc_00E35AAB: call edi
  loc_00E35AAD: mov ecx, [eax]
  loc_00E35AAF: lea edx, var_B4
  loc_00E35AB5: push edx
  loc_00E35AB6: push eax
  loc_00E35AB7: mov var_B8, eax
  loc_00E35ABD: call [ecx+000000E0h]
  loc_00E35AC3: test eax, eax
  loc_00E35AC5: fnclex
  loc_00E35AC7: jge 00E35AE1h
  loc_00E35AC9: mov ecx, var_B8
  loc_00E35ACF: push 000000E0h
  loc_00E35AD4: push 006E03D4h
  loc_00E35AD9: push ecx
  loc_00E35ADA: push eax
  loc_00E35ADB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35AE1: mov edx, var_B4
  loc_00E35AE7: lea ecx, var_30
  loc_00E35AEA: mov var_C0, edx
  loc_00E35AF0: call ebx
  loc_00E35AF2: cmp var_C0, 0000h
  loc_00E35AFA: jz 00E35B62h
  loc_00E35AFC: mov eax, [esi]
  loc_00E35AFE: push esi
  loc_00E35AFF: call [eax+00000328h]
  loc_00E35B05: lea ecx, var_30
  loc_00E35B08: push eax
  loc_00E35B09: push ecx
  loc_00E35B0A: call edi
  loc_00E35B0C: mov edx, [eax]
  loc_00E35B0E: lea ecx, var_28
  loc_00E35B11: push ecx
  loc_00E35B12: push eax
  loc_00E35B13: mov var_B8, eax
  loc_00E35B19: call [edx+000000A0h]
  loc_00E35B1F: test eax, eax
  loc_00E35B21: fnclex
  loc_00E35B23: jge 00E35B3Dh
  loc_00E35B25: mov edx, var_B8
  loc_00E35B2B: push 000000A0h
  loc_00E35B30: push 006DCB70h
  loc_00E35B35: push edx
  loc_00E35B36: push eax
  loc_00E35B37: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35B3D: mov eax, var_28
  loc_00E35B40: push 006E0DB8h ; "select * from biodata where nama like '" & Chr(37)
  loc_00E35B45: push eax
  loc_00E35B46: call [0040106Ch] ; __vbaStrCat
  loc_00E35B4C: mov edx, eax
  loc_00E35B4E: lea ecx, var_2C
  loc_00E35B51: call [00401228h] ; __vbaStrMove
  loc_00E35B57: push eax
  loc_00E35B58: push 006E0E10h ; Chr(37) & "' order by nama asc"
  loc_00E35B5D: jmp 00E35E75h
  loc_00E35B62: mov edx, [esi]
  loc_00E35B64: push esi
  loc_00E35B65: call [edx+00000320h]
  loc_00E35B6B: push eax
  loc_00E35B6C: lea eax, var_30
  loc_00E35B6F: push eax
  loc_00E35B70: call edi
  loc_00E35B72: mov ecx, [eax]
  loc_00E35B74: lea edx, var_B4
  loc_00E35B7A: push edx
  loc_00E35B7B: push eax
  loc_00E35B7C: mov var_B8, eax
  loc_00E35B82: call [ecx+000000E0h]
  loc_00E35B88: test eax, eax
  loc_00E35B8A: fnclex
  loc_00E35B8C: jge 00E35BA6h
  loc_00E35B8E: mov ecx, var_B8
  loc_00E35B94: push 000000E0h
  loc_00E35B99: push 006E03D4h
  loc_00E35B9E: push ecx
  loc_00E35B9F: push eax
  loc_00E35BA0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35BA6: mov edx, var_B4
  loc_00E35BAC: lea ecx, var_30
  loc_00E35BAF: mov var_C0, edx
  loc_00E35BB5: call ebx
  loc_00E35BB7: cmp var_C0, 0000h
  loc_00E35BBF: jz 00E35C27h
  loc_00E35BC1: mov eax, [esi]
  loc_00E35BC3: push esi
  loc_00E35BC4: call [eax+00000328h]
  loc_00E35BCA: lea ecx, var_30
  loc_00E35BCD: push eax
  loc_00E35BCE: push ecx
  loc_00E35BCF: call edi
  loc_00E35BD1: mov edx, [eax]
  loc_00E35BD3: lea ecx, var_28
  loc_00E35BD6: push ecx
  loc_00E35BD7: push eax
  loc_00E35BD8: mov var_B8, eax
  loc_00E35BDE: call [edx+000000A0h]
  loc_00E35BE4: test eax, eax
  loc_00E35BE6: fnclex
  loc_00E35BE8: jge 00E35C02h
  loc_00E35BEA: mov edx, var_B8
  loc_00E35BF0: push 000000A0h
  loc_00E35BF5: push 006DCB70h
  loc_00E35BFA: push edx
  loc_00E35BFB: push eax
  loc_00E35BFC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35C02: mov eax, var_28
  loc_00E35C05: push 006E1D80h ; "select * from biodata where asal_sekolah like '" & Chr(37)
  loc_00E35C0A: push eax
  loc_00E35C0B: call [0040106Ch] ; __vbaStrCat
  loc_00E35C11: mov edx, eax
  loc_00E35C13: lea ecx, var_2C
  loc_00E35C16: call [00401228h] ; __vbaStrMove
  loc_00E35C1C: push eax
  loc_00E35C1D: push 006E1DE8h ; Chr(37) & "' order by asal_sekolah asc"
  loc_00E35C22: jmp 00E35E75h
  loc_00E35C27: mov edx, [esi]
  loc_00E35C29: push esi
  loc_00E35C2A: call [edx+00000324h]
  loc_00E35C30: push eax
  loc_00E35C31: lea eax, var_30
  loc_00E35C34: push eax
  loc_00E35C35: call edi
  loc_00E35C37: mov ecx, [eax]
  loc_00E35C39: lea edx, var_B4
  loc_00E35C3F: push edx
  loc_00E35C40: push eax
  loc_00E35C41: mov var_B8, eax
  loc_00E35C47: call [ecx+000000E0h]
  loc_00E35C4D: test eax, eax
  loc_00E35C4F: fnclex
  loc_00E35C51: jge 00E35C6Bh
  loc_00E35C53: mov ecx, var_B8
  loc_00E35C59: push 000000E0h
  loc_00E35C5E: push 006E03D4h
  loc_00E35C63: push ecx
  loc_00E35C64: push eax
  loc_00E35C65: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35C6B: mov edx, var_B4
  loc_00E35C71: lea ecx, var_30
  loc_00E35C74: mov var_C0, edx
  loc_00E35C7A: call ebx
  loc_00E35C7C: cmp var_C0, 0000h
  loc_00E35C84: jz 00E35CECh
  loc_00E35C86: mov eax, [esi]
  loc_00E35C88: push esi
  loc_00E35C89: call [eax+00000328h]
  loc_00E35C8F: lea ecx, var_30
  loc_00E35C92: push eax
  loc_00E35C93: push ecx
  loc_00E35C94: call edi
  loc_00E35C96: mov edx, [eax]
  loc_00E35C98: lea ecx, var_28
  loc_00E35C9B: push ecx
  loc_00E35C9C: push eax
  loc_00E35C9D: mov var_B8, eax
  loc_00E35CA3: call [edx+000000A0h]
  loc_00E35CA9: test eax, eax
  loc_00E35CAB: fnclex
  loc_00E35CAD: jge 00E35CC7h
  loc_00E35CAF: mov edx, var_B8
  loc_00E35CB5: push 000000A0h
  loc_00E35CBA: push 006DCB70h
  loc_00E35CBF: push edx
  loc_00E35CC0: push eax
  loc_00E35CC1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35CC7: mov eax, var_28
  loc_00E35CCA: push 006E0D20h ; "select * from biodata where nama_ibu like '" & Chr(37)
  loc_00E35CCF: push eax
  loc_00E35CD0: call [0040106Ch] ; __vbaStrCat
  loc_00E35CD6: mov edx, eax
  loc_00E35CD8: lea ecx, var_2C
  loc_00E35CDB: call [00401228h] ; __vbaStrMove
  loc_00E35CE1: push eax
  loc_00E35CE2: push 006E0D80h ; Chr(37) & "' order by nama_ibu asc"
  loc_00E35CE7: jmp 00E35E75h
  loc_00E35CEC: mov edx, [esi]
  loc_00E35CEE: push esi
  loc_00E35CEF: call [edx+00000314h]
  loc_00E35CF5: push eax
  loc_00E35CF6: lea eax, var_30
  loc_00E35CF9: push eax
  loc_00E35CFA: call edi
  loc_00E35CFC: mov ecx, [eax]
  loc_00E35CFE: lea edx, var_B4
  loc_00E35D04: push edx
  loc_00E35D05: push eax
  loc_00E35D06: mov var_B8, eax
  loc_00E35D0C: call [ecx+000000E0h]
  loc_00E35D12: test eax, eax
  loc_00E35D14: fnclex
  loc_00E35D16: jge 00E35D30h
  loc_00E35D18: mov ecx, var_B8
  loc_00E35D1E: push 000000E0h
  loc_00E35D23: push 006E03D4h
  loc_00E35D28: push ecx
  loc_00E35D29: push eax
  loc_00E35D2A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35D30: mov edx, var_B4
  loc_00E35D36: lea ecx, var_30
  loc_00E35D39: mov var_C0, edx
  loc_00E35D3F: call ebx
  loc_00E35D41: cmp var_C0, 0000h
  loc_00E35D49: jz 00E35DB1h
  loc_00E35D4B: mov eax, [esi]
  loc_00E35D4D: push esi
  loc_00E35D4E: call [eax+00000328h]
  loc_00E35D54: lea ecx, var_30
  loc_00E35D57: push eax
  loc_00E35D58: push ecx
  loc_00E35D59: call edi
  loc_00E35D5B: mov edx, [eax]
  loc_00E35D5D: lea ecx, var_28
  loc_00E35D60: push ecx
  loc_00E35D61: push eax
  loc_00E35D62: mov var_B8, eax
  loc_00E35D68: call [edx+000000A0h]
  loc_00E35D6E: test eax, eax
  loc_00E35D70: fnclex
  loc_00E35D72: jge 00E35D8Ch
  loc_00E35D74: mov edx, var_B8
  loc_00E35D7A: push 000000A0h
  loc_00E35D7F: push 006DCB70h
  loc_00E35D84: push edx
  loc_00E35D85: push eax
  loc_00E35D86: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35D8C: mov eax, var_28
  loc_00E35D8F: push 006E1E28h ; "select * from biodata where jurusan_dipilih like '" & Chr(37)
  loc_00E35D94: push eax
  loc_00E35D95: call [0040106Ch] ; __vbaStrCat
  loc_00E35D9B: mov edx, eax
  loc_00E35D9D: lea ecx, var_2C
  loc_00E35DA0: call [00401228h] ; __vbaStrMove
  loc_00E35DA6: push eax
  loc_00E35DA7: push 006E1E94h ; Chr(37) & "' order by jurusan_dipilih asc"
  loc_00E35DAC: jmp 00E35E75h
  loc_00E35DB1: mov edx, [esi]
  loc_00E35DB3: push esi
  loc_00E35DB4: call [edx+0000031Ch]
  loc_00E35DBA: push eax
  loc_00E35DBB: lea eax, var_30
  loc_00E35DBE: push eax
  loc_00E35DBF: call edi
  loc_00E35DC1: mov ecx, [eax]
  loc_00E35DC3: lea edx, var_B4
  loc_00E35DC9: push edx
  loc_00E35DCA: push eax
  loc_00E35DCB: mov var_B8, eax
  loc_00E35DD1: call [ecx+000000E0h]
  loc_00E35DD7: test eax, eax
  loc_00E35DD9: fnclex
  loc_00E35DDB: jge 00E35DF5h
  loc_00E35DDD: mov ecx, var_B8
  loc_00E35DE3: push 000000E0h
  loc_00E35DE8: push 006E03D4h
  loc_00E35DED: push ecx
  loc_00E35DEE: push eax
  loc_00E35DEF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35DF5: mov edx, var_B4
  loc_00E35DFB: lea ecx, var_30
  loc_00E35DFE: mov var_C0, edx
  loc_00E35E04: call ebx
  loc_00E35E06: cmp var_C0, 0000h
  loc_00E35E0E: jz 00E35F1Eh
  loc_00E35E14: mov eax, [esi]
  loc_00E35E16: push esi
  loc_00E35E17: call [eax+00000328h]
  loc_00E35E1D: lea ecx, var_30
  loc_00E35E20: push eax
  loc_00E35E21: push ecx
  loc_00E35E22: call edi
  loc_00E35E24: mov edx, [eax]
  loc_00E35E26: lea ecx, var_28
  loc_00E35E29: push ecx
  loc_00E35E2A: push eax
  loc_00E35E2B: mov var_B8, eax
  loc_00E35E31: call [edx+000000A0h]
  loc_00E35E37: test eax, eax
  loc_00E35E39: fnclex
  loc_00E35E3B: jge 00E35E55h
  loc_00E35E3D: mov edx, var_B8
  loc_00E35E43: push 000000A0h
  loc_00E35E48: push 006DCB70h
  loc_00E35E4D: push edx
  loc_00E35E4E: push eax
  loc_00E35E4F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35E55: mov eax, var_28
  loc_00E35E58: push 006E1ED8h ; "select * from biodata where agama like '" & Chr(37)
  loc_00E35E5D: push eax
  loc_00E35E5E: call [0040106Ch] ; __vbaStrCat
  loc_00E35E64: mov edx, eax
  loc_00E35E66: lea ecx, var_2C
  loc_00E35E69: call [00401228h] ; __vbaStrMove
  loc_00E35E6F: push eax
  loc_00E35E70: push 006E1F30h ; Chr(37) & "' order by agama asc"
  loc_00E35E75: call [0040106Ch] ; __vbaStrCat
  loc_00E35E7B: lea edx, var_40
  loc_00E35E7E: lea ecx, var_24
  loc_00E35E81: mov var_38, eax
  loc_00E35E84: mov var_40, 00000008h
  loc_00E35E8B: call [0040101Ch] ; __vbaVarMove
  loc_00E35E91: lea ecx, var_2C
  loc_00E35E94: lea edx, var_28
  loc_00E35E97: push ecx
  loc_00E35E98: push edx
  loc_00E35E99: push 00000002h
  loc_00E35E9B: call [004011BCh] ; __vbaFreeStrList
  loc_00E35EA1: add esp, 0000000Ch
  loc_00E35EA4: lea ecx, var_30
  loc_00E35EA7: call ebx
  loc_00E35EA9: lea eax, var_24
  loc_00E35EAC: push eax
  loc_00E35EAD: call [00401230h] ; __vbaStrVarCopy
  loc_00E35EB3: sub esp, 00000010h
  loc_00E35EB6: mov ecx, 00000008h
  loc_00E35EBB: mov edx, esp
  loc_00E35EBD: mov var_40, ecx
  loc_00E35EC0: mov var_38, eax
  loc_00E35EC3: push 0000000Eh
  loc_00E35EC5: mov [edx], ecx
  loc_00E35EC7: mov ecx, var_3C
  loc_00E35ECA: push esi
  loc_00E35ECB: mov [edx+00000004h], ecx
  loc_00E35ECE: mov ecx, [esi]
  loc_00E35ED0: mov [edx+00000008h], eax
  loc_00E35ED3: mov eax, var_34
  loc_00E35ED6: mov [edx+0000000Ch], eax
  loc_00E35ED9: call [ecx+00000358h]
  loc_00E35EDF: lea edx, var_30
  loc_00E35EE2: push eax
  loc_00E35EE3: push edx
  loc_00E35EE4: call edi
  loc_00E35EE6: push eax
  loc_00E35EE7: call [00401238h] ; __vbaLateIdSt
  loc_00E35EED: lea ecx, var_30
  loc_00E35EF0: call ebx
  loc_00E35EF2: lea ecx, var_40
  loc_00E35EF5: call [00401024h] ; __vbaFreeVar
  loc_00E35EFB: mov eax, [esi]
  loc_00E35EFD: push 00000000h
  loc_00E35EFF: push 00000019h
  loc_00E35F01: push esi
  loc_00E35F02: call [eax+00000358h]
  loc_00E35F08: lea ecx, var_30
  loc_00E35F0B: push eax
  loc_00E35F0C: push ecx
  loc_00E35F0D: call edi
  loc_00E35F0F: push eax
  loc_00E35F10: call [00401030h] ; __vbaLateIdCall
  loc_00E35F16: add esp, 0000000Ch
  loc_00E35F19: jmp 00E36033h
  loc_00E35F1E: mov edx, [esi]
  loc_00E35F20: push 00000000h
  loc_00E35F22: push 80010007h
  loc_00E35F27: push esi
  loc_00E35F28: call [edx+00000334h]
  loc_00E35F2E: push eax
  loc_00E35F2F: lea eax, var_30
  loc_00E35F32: push eax
  loc_00E35F33: call edi
  loc_00E35F35: lea ecx, var_40
  loc_00E35F38: push eax
  loc_00E35F39: push ecx
  loc_00E35F3A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E35F40: add esp, 00000010h
  loc_00E35F43: push eax
  loc_00E35F44: call [004010CCh] ; __vbaBoolVar
  loc_00E35F4A: neg ax
  loc_00E35F4D: sbb eax, eax
  loc_00E35F4F: lea ecx, var_30
  loc_00E35F52: inc eax
  loc_00E35F53: neg eax
  loc_00E35F55: mov var_B8, ax
  loc_00E35F5C: call ebx
  loc_00E35F5E: lea ecx, var_40
  loc_00E35F61: call [00401024h] ; __vbaFreeVar
  loc_00E35F67: cmp var_B8, 0000h
  loc_00E35F6F: jz 00E36038h
  loc_00E35F75: mov ecx, 80020004h
  loc_00E35F7A: mov eax, 0000000Ah
  loc_00E35F7F: mov var_68, ecx
  loc_00E35F82: mov var_58, ecx
  loc_00E35F85: lea edx, var_90
  loc_00E35F8B: lea ecx, var_50
  loc_00E35F8E: mov var_70, eax
  loc_00E35F91: mov var_60, eax
  loc_00E35F94: mov var_88, 006E16F0h ; "SMK RK BT PS"
  loc_00E35F9E: mov var_90, 00000008h
  loc_00E35FA8: call [004011F0h] ; __vbaVarDup
  loc_00E35FAE: lea edx, var_80
  loc_00E35FB1: lea ecx, var_40
  loc_00E35FB4: mov var_78, 006E1CE8h ; "Silahkan Pilih Item Data yang Ingin dicari !"
  loc_00E35FBB: mov var_80, 00000008h
  loc_00E35FC2: call [004011F0h] ; __vbaVarDup
  loc_00E35FC8: lea edx, var_70
  loc_00E35FCB: lea eax, var_60
  loc_00E35FCE: push edx
  loc_00E35FCF: lea ecx, var_50
  loc_00E35FD2: push eax
  loc_00E35FD3: push ecx
  loc_00E35FD4: lea edx, var_40
  loc_00E35FD7: push 00000040h
  loc_00E35FD9: push edx
  loc_00E35FDA: call [004010A8h] ; rtcMsgBox
  loc_00E35FE0: lea eax, var_70
  loc_00E35FE3: lea ecx, var_60
  loc_00E35FE6: push eax
  loc_00E35FE7: lea edx, var_50
  loc_00E35FEA: push ecx
  loc_00E35FEB: lea eax, var_40
  loc_00E35FEE: push edx
  loc_00E35FEF: push eax
  loc_00E35FF0: push 00000004h
  loc_00E35FF2: call [00401038h] ; __vbaFreeVarList
  loc_00E35FF8: mov ecx, [esi]
  loc_00E35FFA: add esp, 00000014h
  loc_00E35FFD: push esi
  loc_00E35FFE: call [ecx+00000328h]
  loc_00E36004: lea edx, var_30
  loc_00E36007: push eax
  loc_00E36008: push edx
  loc_00E36009: call edi
  loc_00E3600B: mov esi, eax
  loc_00E3600D: push 006DCC80h
  loc_00E36012: push esi
  loc_00E36013: mov eax, [esi]
  loc_00E36015: call [eax+000000A4h]
  loc_00E3601B: test eax, eax
  loc_00E3601D: fnclex
  loc_00E3601F: jge 00E36033h
  loc_00E36021: push 000000A4h
  loc_00E36026: push 006DCB70h
  loc_00E3602B: push esi
  loc_00E3602C: push eax
  loc_00E3602D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36033: lea ecx, var_30
  loc_00E36036: call ebx
  loc_00E36038: mov var_4, 00000000h
  loc_00E3603F: push 00E36088h
  loc_00E36044: jmp 00E3607Eh
  loc_00E36046: lea ecx, var_2C
  loc_00E36049: lea edx, var_28
  loc_00E3604C: push ecx
  loc_00E3604D: push edx
  loc_00E3604E: push 00000002h
  loc_00E36050: call [004011BCh] ; __vbaFreeStrList
  loc_00E36056: add esp, 0000000Ch
  loc_00E36059: lea ecx, var_30
  loc_00E3605C: call [00401254h] ; __vbaFreeObj
  loc_00E36062: lea eax, var_70
  loc_00E36065: lea ecx, var_60
  loc_00E36068: push eax
  loc_00E36069: lea edx, var_50
  loc_00E3606C: push ecx
  loc_00E3606D: lea eax, var_40
  loc_00E36070: push edx
  loc_00E36071: push eax
  loc_00E36072: push 00000004h
  loc_00E36074: call [00401038h] ; __vbaFreeVarList
  loc_00E3607A: add esp, 00000014h
  loc_00E3607D: ret
  loc_00E3607E: lea ecx, var_24
  loc_00E36081: call [00401024h] ; __vbaFreeVar
  loc_00E36087: ret
  loc_00E36088: mov eax, Me
  loc_00E3608B: push eax
  loc_00E3608C: mov ecx, [eax]
  loc_00E3608E: call [ecx+00000008h]
  loc_00E36091: mov eax, var_4
  loc_00E36094: mov ecx, var_14
  loc_00E36097: pop edi
  loc_00E36098: pop esi
  loc_00E36099: mov fs:[00000000h], ecx
  loc_00E360A0: pop ebx
  loc_00E360A1: mov esp, ebp
  loc_00E360A3: pop ebp
  loc_00E360A4: retn 0008h
End Sub

Private Sub Form_Load() 'E337F0
  loc_00E337F0: push ebp
  loc_00E337F1: mov ebp, esp
  loc_00E337F3: sub esp, 0000000Ch
  loc_00E337F6: push 00402836h ; __vbaExceptHandler
  loc_00E337FB: mov eax, fs:[00000000h]
  loc_00E33801: push eax
  loc_00E33802: mov fs:[00000000h], esp
  loc_00E33809: sub esp, 00000058h
  loc_00E3380C: push ebx
  loc_00E3380D: push esi
  loc_00E3380E: push edi
  loc_00E3380F: mov var_C, esp
  loc_00E33812: mov var_8, 00402620h
  loc_00E33819: mov esi, Me
  loc_00E3381C: mov eax, esi
  loc_00E3381E: and eax, 00000001h
  loc_00E33821: mov var_4, eax
  loc_00E33824: and esi, FFFFFFFEh
  loc_00E33827: push esi
  loc_00E33828: mov Me, esi
  loc_00E3382B: mov ecx, [esi]
  loc_00E3382D: call [ecx+00000004h]
  loc_00E33830: mov edx, [esi]
  loc_00E33832: xor eax, eax
  loc_00E33834: push esi
  loc_00E33835: mov var_24, eax
  loc_00E33838: mov var_28, eax
  loc_00E3383B: mov var_2C, eax
  loc_00E3383E: mov var_3C, eax
  loc_00E33841: mov var_4C, eax
  loc_00E33844: call [edx+0000031Ch]
  loc_00E3384A: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E33850: push eax
  loc_00E33851: lea eax, var_28
  loc_00E33854: push eax
  loc_00E33855: call edi
  loc_00E33857: mov ebx, eax
  loc_00E33859: push 00000000h
  loc_00E3385B: push ebx
  loc_00E3385C: mov ecx, [ebx]
  loc_00E3385E: call [ecx+000000E4h]
  loc_00E33864: test eax, eax
  loc_00E33866: fnclex
  loc_00E33868: jge 00E3387Ch
  loc_00E3386A: push 000000E4h
  loc_00E3386F: push 006E03D4h
  loc_00E33874: push ebx
  loc_00E33875: push eax
  loc_00E33876: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3387C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E33882: lea ecx, var_28
  loc_00E33885: call ebx
  loc_00E33887: mov edx, [esi]
  loc_00E33889: push esi
  loc_00E3388A: call [edx+0000032Ch]
  loc_00E33890: push eax
  loc_00E33891: lea eax, var_28
  loc_00E33894: push eax
  loc_00E33895: call edi
  loc_00E33897: mov ecx, [eax]
  loc_00E33899: push 00000000h
  loc_00E3389B: push eax
  loc_00E3389C: mov var_60, eax
  loc_00E3389F: call [ecx+000000E4h]
  loc_00E338A5: test eax, eax
  loc_00E338A7: fnclex
  loc_00E338A9: jge 00E338C0h
  loc_00E338AB: mov edx, var_60
  loc_00E338AE: push 000000E4h
  loc_00E338B3: push 006E03D4h
  loc_00E338B8: push edx
  loc_00E338B9: push eax
  loc_00E338BA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E338C0: lea ecx, var_28
  loc_00E338C3: call ebx
  loc_00E338C5: mov eax, [esi]
  loc_00E338C7: push esi
  loc_00E338C8: call [eax+00000330h]
  loc_00E338CE: lea ecx, var_28
  loc_00E338D1: push eax
  loc_00E338D2: push ecx
  loc_00E338D3: call edi
  loc_00E338D5: mov edx, [eax]
  loc_00E338D7: push 00000000h
  loc_00E338D9: push eax
  loc_00E338DA: mov var_60, eax
  loc_00E338DD: call [edx+000000E4h]
  loc_00E338E3: test eax, eax
  loc_00E338E5: fnclex
  loc_00E338E7: jge 00E338FEh
  loc_00E338E9: mov ecx, var_60
  loc_00E338EC: push 000000E4h
  loc_00E338F1: push 006E03D4h
  loc_00E338F6: push ecx
  loc_00E338F7: push eax
  loc_00E338F8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E338FE: lea ecx, var_28
  loc_00E33901: call ebx
  loc_00E33903: mov edx, [esi]
  loc_00E33905: push esi
  loc_00E33906: call [edx+00000324h]
  loc_00E3390C: push eax
  loc_00E3390D: lea eax, var_28
  loc_00E33910: push eax
  loc_00E33911: call edi
  loc_00E33913: mov ecx, [eax]
  loc_00E33915: push 00000000h
  loc_00E33917: push eax
  loc_00E33918: mov var_60, eax
  loc_00E3391B: call [ecx+000000E4h]
  loc_00E33921: test eax, eax
  loc_00E33923: fnclex
  loc_00E33925: jge 00E3393Ch
  loc_00E33927: mov edx, var_60
  loc_00E3392A: push 000000E4h
  loc_00E3392F: push 006E03D4h
  loc_00E33934: push edx
  loc_00E33935: push eax
  loc_00E33936: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3393C: lea ecx, var_28
  loc_00E3393F: call ebx
  loc_00E33941: mov eax, [esi]
  loc_00E33943: push esi
  loc_00E33944: call [eax+00000314h]
  loc_00E3394A: lea ecx, var_28
  loc_00E3394D: push eax
  loc_00E3394E: push ecx
  loc_00E3394F: call edi
  loc_00E33951: mov edx, [eax]
  loc_00E33953: push 00000000h
  loc_00E33955: push eax
  loc_00E33956: mov var_60, eax
  loc_00E33959: call [edx+000000E4h]
  loc_00E3395F: test eax, eax
  loc_00E33961: fnclex
  loc_00E33963: jge 00E3397Ah
  loc_00E33965: mov ecx, var_60
  loc_00E33968: push 000000E4h
  loc_00E3396D: push 006E03D4h
  loc_00E33972: push ecx
  loc_00E33973: push eax
  loc_00E33974: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3397A: lea ecx, var_28
  loc_00E3397D: call ebx
  loc_00E3397F: mov edx, [esi]
  loc_00E33981: push esi
  loc_00E33982: call [edx+00000320h]
  loc_00E33988: push eax
  loc_00E33989: lea eax, var_28
  loc_00E3398C: push eax
  loc_00E3398D: call edi
  loc_00E3398F: mov ecx, [eax]
  loc_00E33991: push 00000000h
  loc_00E33993: push eax
  loc_00E33994: mov var_60, eax
  loc_00E33997: call [ecx+000000E4h]
  loc_00E3399D: test eax, eax
  loc_00E3399F: fnclex
  loc_00E339A1: jge 00E339B8h
  loc_00E339A3: mov edx, var_60
  loc_00E339A6: push 000000E4h
  loc_00E339AB: push 006E03D4h
  loc_00E339B0: push edx
  loc_00E339B1: push eax
  loc_00E339B2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E339B8: lea ecx, var_28
  loc_00E339BB: call ebx
  loc_00E339BD: sub esp, 00000010h
  loc_00E339C0: mov ecx, 0000000Bh
  loc_00E339C5: mov edx, esp
  loc_00E339C7: mov var_4C, ecx
  loc_00E339CA: xor eax, eax
  loc_00E339CC: push 80010007h
  loc_00E339D1: mov [edx], ecx
  loc_00E339D3: mov ecx, var_48
  loc_00E339D6: mov var_44, eax
  loc_00E339D9: push esi
  loc_00E339DA: mov [edx+00000004h], ecx
  loc_00E339DD: mov ecx, [esi]
  loc_00E339DF: mov [edx+00000008h], eax
  loc_00E339E2: mov eax, var_40
  loc_00E339E5: mov [edx+0000000Ch], eax
  loc_00E339E8: call [ecx+00000334h]
  loc_00E339EE: lea edx, var_28
  loc_00E339F1: push eax
  loc_00E339F2: push edx
  loc_00E339F3: call edi
  loc_00E339F5: push eax
  loc_00E339F6: call [00401238h] ; __vbaLateIdSt
  loc_00E339FC: lea ecx, var_28
  loc_00E339FF: call ebx
  loc_00E33A01: mov eax, [esi]
  loc_00E33A03: push esi
  loc_00E33A04: call [eax+000002FCh]
  loc_00E33A0A: lea ecx, var_28
  loc_00E33A0D: push eax
  loc_00E33A0E: push ecx
  loc_00E33A0F: call edi
  loc_00E33A11: mov edx, [eax]
  loc_00E33A13: push 00000000h
  loc_00E33A15: push eax
  loc_00E33A16: mov var_60, eax
  loc_00E33A19: call [edx+0000009Ch]
  loc_00E33A1F: test eax, eax
  loc_00E33A21: fnclex
  loc_00E33A23: jge 00E33A3Ah
  loc_00E33A25: mov ecx, var_60
  loc_00E33A28: push 0000009Ch
  loc_00E33A2D: push 006DCAD0h
  loc_00E33A32: push ecx
  loc_00E33A33: push eax
  loc_00E33A34: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33A3A: lea ecx, var_28
  loc_00E33A3D: call ebx
  loc_00E33A3F: lea edx, var_4C
  loc_00E33A42: lea ecx, var_24
  loc_00E33A45: mov var_44, 006E05ACh ; "select * from biodata"
  loc_00E33A4C: mov var_4C, 00000008h
  loc_00E33A53: call [00401204h] ; __vbaVarCopy
  loc_00E33A59: lea edx, var_24
  loc_00E33A5C: push edx
  loc_00E33A5D: call [00401230h] ; __vbaStrVarCopy
  loc_00E33A63: sub esp, 00000010h
  loc_00E33A66: mov ecx, 00000008h
  loc_00E33A6B: mov edx, esp
  loc_00E33A6D: mov var_3C, ecx
  loc_00E33A70: mov var_34, eax
  loc_00E33A73: push 0000000Eh
  loc_00E33A75: mov [edx], ecx
  loc_00E33A77: mov ecx, var_38
  loc_00E33A7A: push esi
  loc_00E33A7B: mov [edx+00000004h], ecx
  loc_00E33A7E: mov ecx, [esi]
  loc_00E33A80: mov [edx+00000008h], eax
  loc_00E33A83: mov eax, var_30
  loc_00E33A86: mov [edx+0000000Ch], eax
  loc_00E33A89: call [ecx+00000358h]
  loc_00E33A8F: lea edx, var_28
  loc_00E33A92: push eax
  loc_00E33A93: push edx
  loc_00E33A94: call edi
  loc_00E33A96: push eax
  loc_00E33A97: call [00401238h] ; __vbaLateIdSt
  loc_00E33A9D: lea ecx, var_28
  loc_00E33AA0: call ebx
  loc_00E33AA2: lea ecx, var_3C
  loc_00E33AA5: call [00401024h] ; __vbaFreeVar
  loc_00E33AAB: mov eax, [esi]
  loc_00E33AAD: push 00000000h
  loc_00E33AAF: push 00000019h
  loc_00E33AB1: push esi
  loc_00E33AB2: call [eax+00000358h]
  loc_00E33AB8: lea ecx, var_28
  loc_00E33ABB: push eax
  loc_00E33ABC: push ecx
  loc_00E33ABD: call edi
  loc_00E33ABF: push eax
  loc_00E33AC0: call [00401030h] ; __vbaLateIdCall
  loc_00E33AC6: add esp, 0000000Ch
  loc_00E33AC9: lea ecx, var_28
  loc_00E33ACC: call ebx
  loc_00E33ACE: mov edx, [esi]
  loc_00E33AD0: push 006E05D8h
  loc_00E33AD5: push esi
  loc_00E33AD6: call [edx+00000358h]
  loc_00E33ADC: push eax
  loc_00E33ADD: lea eax, var_28
  loc_00E33AE0: push eax
  loc_00E33AE1: call edi
  loc_00E33AE3: push eax
  loc_00E33AE4: call [00401224h] ; __vbaCastObj
  loc_00E33AEA: lea ecx, var_3C
  loc_00E33AED: mov var_34, eax
  loc_00E33AF0: push ecx
  loc_00E33AF1: mov var_3C, 0000000Dh
  loc_00E33AF8: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E33AFE: mov ecx, [eax]
  loc_00E33B00: sub esp, 00000010h
  loc_00E33B03: mov edx, esp
  loc_00E33B05: push 00000000h
  loc_00E33B07: push 0000002Ah
  loc_00E33B09: mov [edx], ecx
  loc_00E33B0B: mov ecx, [eax+00000004h]
  loc_00E33B0E: push esi
  loc_00E33B0F: mov [edx+00000004h], ecx
  loc_00E33B12: mov ecx, [eax+00000008h]
  loc_00E33B15: mov eax, [eax+0000000Ch]
  loc_00E33B18: mov [edx+00000008h], ecx
  loc_00E33B1B: mov ecx, [esi]
  loc_00E33B1D: mov [edx+0000000Ch], eax
  loc_00E33B20: call [ecx+0000035Ch]
  loc_00E33B26: push eax
  loc_00E33B27: lea edx, var_2C
  loc_00E33B2A: push edx
  loc_00E33B2B: call edi
  loc_00E33B2D: push eax
  loc_00E33B2E: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E33B34: lea eax, var_2C
  loc_00E33B37: lea ecx, var_28
  loc_00E33B3A: push eax
  loc_00E33B3B: push ecx
  loc_00E33B3C: push 00000002h
  loc_00E33B3E: call [00401048h] ; __vbaFreeObjList
  loc_00E33B44: add esp, 00000028h
  loc_00E33B47: lea ecx, var_3C
  loc_00E33B4A: call [00401024h] ; __vbaFreeVar
  loc_00E33B50: mov edx, [esi]
  loc_00E33B52: push 00000000h
  loc_00E33B54: push FFFFFDDAh
  loc_00E33B59: push esi
  loc_00E33B5A: call [edx+0000035Ch]
  loc_00E33B60: push eax
  loc_00E33B61: lea eax, var_28
  loc_00E33B64: push eax
  loc_00E33B65: call edi
  loc_00E33B67: push eax
  loc_00E33B68: call [00401030h] ; __vbaLateIdCall
  loc_00E33B6E: add esp, 0000000Ch
  loc_00E33B71: lea ecx, var_28
  loc_00E33B74: call ebx
  loc_00E33B76: mov var_4, 00000000h
  loc_00E33B7D: push 00E33BABh
  loc_00E33B82: jmp 00E33BA1h
  loc_00E33B84: lea ecx, var_2C
  loc_00E33B87: lea edx, var_28
  loc_00E33B8A: push ecx
  loc_00E33B8B: push edx
  loc_00E33B8C: push 00000002h
  loc_00E33B8E: call [00401048h] ; __vbaFreeObjList
  loc_00E33B94: add esp, 0000000Ch
  loc_00E33B97: lea ecx, var_3C
  loc_00E33B9A: call [00401024h] ; __vbaFreeVar
  loc_00E33BA0: ret
  loc_00E33BA1: lea ecx, var_24
  loc_00E33BA4: call [00401024h] ; __vbaFreeVar
  loc_00E33BAA: ret
  loc_00E33BAB: mov eax, Me
  loc_00E33BAE: push eax
  loc_00E33BAF: mov ecx, [eax]
  loc_00E33BB1: call [ecx+00000008h]
  loc_00E33BB4: mov eax, var_4
  loc_00E33BB7: mov ecx, var_14
  loc_00E33BBA: pop edi
  loc_00E33BBB: pop esi
  loc_00E33BBC: mov fs:[00000000h], ecx
  loc_00E33BC3: pop ebx
  loc_00E33BC4: mov esp, ebp
  loc_00E33BC6: pop ebp
  loc_00E33BC7: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E33BD0
  loc_00E33BD0: push ebp
  loc_00E33BD1: mov ebp, esp
  loc_00E33BD3: sub esp, 0000000Ch
  loc_00E33BD6: push 00402836h ; __vbaExceptHandler
  loc_00E33BDB: mov eax, fs:[00000000h]
  loc_00E33BE1: push eax
  loc_00E33BE2: mov fs:[00000000h], esp
  loc_00E33BE9: sub esp, 0000005Ch
  loc_00E33BEC: push ebx
  loc_00E33BED: push esi
  loc_00E33BEE: push edi
  loc_00E33BEF: mov var_C, esp
  loc_00E33BF2: mov var_8, 00402630h
  loc_00E33BF9: mov esi, Me
  loc_00E33BFC: mov eax, esi
  loc_00E33BFE: and eax, 00000001h
  loc_00E33C01: mov var_4, eax
  loc_00E33C04: and esi, FFFFFFFEh
  loc_00E33C07: push esi
  loc_00E33C08: mov Me, esi
  loc_00E33C0B: mov ecx, [esi]
  loc_00E33C0D: call [ecx+00000004h]
  loc_00E33C10: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33C16: xor eax, eax
  loc_00E33C18: mov var_18, eax
  loc_00E33C1B: mov var_4C, eax
  loc_00E33C1E: mov var_50, eax
  loc_00E33C21: mov edx, [esi]
  loc_00E33C23: lea eax, var_4C
  loc_00E33C26: push eax
  loc_00E33C27: push esi
  loc_00E33C28: call [edx+00000070h]
  loc_00E33C2B: test eax, eax
  loc_00E33C2D: fnclex
  loc_00E33C2F: jge 00E33C3Ch
  loc_00E33C31: push 00000070h
  loc_00E33C33: push 006DFDE0h
  loc_00E33C38: push esi
  loc_00E33C39: push eax
  loc_00E33C3A: call ebx
  loc_00E33C3C: fld real4 ptr var_4C
  loc_00E33C3F: fadd st0, real4 ptr [00401298h]
  loc_00E33C45: mov ecx, [esi]
  loc_00E33C47: push ecx
  loc_00E33C48: fnstsw ax
  loc_00E33C4A: test al, 0Dh
  loc_00E33C4C: jnz 00E33E10h
  loc_00E33C52: fstp real4 ptr [esp]
  loc_00E33C55: push esi
  loc_00E33C56: call [ecx+00000074h]
  loc_00E33C59: test eax, eax
  loc_00E33C5B: fnclex
  loc_00E33C5D: jge 00E33C6Ah
  loc_00E33C5F: push 00000074h
  loc_00E33C61: push 006DFDE0h
  loc_00E33C66: push esi
  loc_00E33C67: push eax
  loc_00E33C68: call ebx
  loc_00E33C6A: mov edx, [esi]
  loc_00E33C6C: lea eax, var_4C
  loc_00E33C6F: push eax
  loc_00E33C70: push esi
  loc_00E33C71: call [edx+00000070h]
  loc_00E33C74: test eax, eax
  loc_00E33C76: fnclex
  loc_00E33C78: jge 00E33C85h
  loc_00E33C7A: push 00000070h
  loc_00E33C7C: push 006DFDE0h
  loc_00E33C81: push esi
  loc_00E33C82: push eax
  loc_00E33C83: call ebx
  loc_00E33C85: mov ecx, [esi]
  loc_00E33C87: lea edx, var_50
  loc_00E33C8A: push edx
  loc_00E33C8B: push esi
  loc_00E33C8C: call [ecx+00000078h]
  loc_00E33C8F: test eax, eax
  loc_00E33C91: fnclex
  loc_00E33C93: jge 00E33CA0h
  loc_00E33C95: push 00000078h
  loc_00E33C97: push 006DFDE0h
  loc_00E33C9C: push esi
  loc_00E33C9D: push eax
  loc_00E33C9E: call ebx
  loc_00E33CA0: sub esp, 00000010h
  loc_00E33CA3: mov ecx, 0000000Ah
  loc_00E33CA8: mov ebx, esp
  loc_00E33CAA: mov eax, 80020004h
  loc_00E33CAF: mov edx, eax
  loc_00E33CB1: sub esp, 00000010h
  loc_00E33CB4: mov [ebx], ecx
  loc_00E33CB6: mov ecx, var_44
  loc_00E33CB9: mov edi, [esi]
  loc_00E33CBB: mov [ebx+00000004h], ecx
  loc_00E33CBE: mov ecx, esp
  loc_00E33CC0: sub esp, 00000010h
  loc_00E33CC3: mov [ebx+00000008h], eax
  loc_00E33CC6: mov eax, var_3C
  loc_00E33CC9: mov [ebx+0000000Ch], eax
  loc_00E33CCC: mov eax, 0000000Ah
  loc_00E33CD1: mov [ecx], eax
  loc_00E33CD3: mov eax, var_34
  loc_00E33CD6: mov [ecx+00000004h], eax
  loc_00E33CD9: mov eax, 00000004h
  loc_00E33CDE: mov [ecx+00000008h], edx
  loc_00E33CE1: mov edx, var_2C
  loc_00E33CE4: mov [ecx+0000000Ch], edx
  loc_00E33CE7: mov edx, var_24
  loc_00E33CEA: mov ecx, esp
  loc_00E33CEC: mov [ecx], eax
  loc_00E33CEE: mov eax, var_50
  loc_00E33CF1: mov [ecx+00000004h], edx
  loc_00E33CF4: mov [ecx+00000008h], eax
  loc_00E33CF7: mov eax, var_1C
  loc_00E33CFA: mov [ecx+0000000Ch], eax
  loc_00E33CFD: mov ecx, var_4C
  loc_00E33D00: push ecx
  loc_00E33D01: push esi
  loc_00E33D02: call [edi+000002A4h]
  loc_00E33D08: test eax, eax
  loc_00E33D0A: fnclex
  loc_00E33D0C: jge 00E33D24h
  loc_00E33D0E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33D14: push 000002A4h
  loc_00E33D19: push 006DFDE0h
  loc_00E33D1E: push esi
  loc_00E33D1F: push eax
  loc_00E33D20: call ebx
  loc_00E33D22: jmp 00E33D2Ah
  loc_00E33D24: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33D2A: call [004010BCh] ; rtcDoEvents
  loc_00E33D30: mov edx, [esi]
  loc_00E33D32: lea eax, var_4C
  loc_00E33D35: push eax
  loc_00E33D36: push esi
  loc_00E33D37: call [edx+00000070h]
  loc_00E33D3A: test eax, eax
  loc_00E33D3C: fnclex
  loc_00E33D3E: jge 00E33D4Bh
  loc_00E33D40: push 00000070h
  loc_00E33D42: push 006DFDE0h
  loc_00E33D47: push esi
  loc_00E33D48: push eax
  loc_00E33D49: call ebx
  loc_00E33D4B: mov eax, [00E3D9CCh]
  loc_00E33D50: test eax, eax
  loc_00E33D52: jnz 00E33D64h
  loc_00E33D54: push 00E3D9CCh
  loc_00E33D59: push 006DC960h
  loc_00E33D5E: call [004011A0h] ; __vbaNew2
  loc_00E33D64: mov edi, [00E3D9CCh]
  loc_00E33D6A: lea edx, var_18
  loc_00E33D6D: push edx
  loc_00E33D6E: push edi
  loc_00E33D6F: mov ecx, [edi]
  loc_00E33D71: call [ecx+00000018h]
  loc_00E33D74: test eax, eax
  loc_00E33D76: fnclex
  loc_00E33D78: jge 00E33D85h
  loc_00E33D7A: push 00000018h
  loc_00E33D7C: push 006DC950h
  loc_00E33D81: push edi
  loc_00E33D82: push eax
  loc_00E33D83: call ebx
  loc_00E33D85: mov eax, var_18
  loc_00E33D88: lea edx, var_50
  loc_00E33D8B: push edx
  loc_00E33D8C: push eax
  loc_00E33D8D: mov ecx, [eax]
  loc_00E33D8F: mov edi, eax
  loc_00E33D91: call [ecx+00000098h]
  loc_00E33D97: test eax, eax
  loc_00E33D99: fnclex
  loc_00E33D9B: jge 00E33DABh
  loc_00E33D9D: push 00000098h
  loc_00E33DA2: push 006DCAF0h
  loc_00E33DA7: push edi
  loc_00E33DA8: push eax
  loc_00E33DA9: call ebx
  loc_00E33DAB: fld real4 ptr var_4C
  loc_00E33DAE: fcomp real4 ptr var_50
  loc_00E33DB1: fnstsw ax
  loc_00E33DB3: test ah, 41h
  loc_00E33DB6: jz 00E33DBFh
  loc_00E33DB8: mov eax, 00000001h
  loc_00E33DBD: jmp 00E33DC1h
  loc_00E33DBF: xor eax, eax
  loc_00E33DC1: neg eax
  loc_00E33DC3: lea ecx, var_18
  loc_00E33DC6: mov edi, eax
  loc_00E33DC8: call [00401254h] ; __vbaFreeObj
  loc_00E33DCE: test di, di
  loc_00E33DD1: jnz 00E33C21h
  loc_00E33DD7: mov var_4, 00000000h
  loc_00E33DDE: fwait
  loc_00E33DDF: push 00E33DF1h
  loc_00E33DE4: jmp 00E33DF0h
  loc_00E33DE6: lea ecx, var_18
  loc_00E33DE9: call [00401254h] ; __vbaFreeObj
  loc_00E33DEF: ret
  loc_00E33DF0: ret
  loc_00E33DF1: mov eax, Me
  loc_00E33DF4: push eax
  loc_00E33DF5: mov ecx, [eax]
  loc_00E33DF7: call [ecx+00000008h]
  loc_00E33DFA: mov eax, var_4
  loc_00E33DFD: mov ecx, var_14
  loc_00E33E00: pop edi
  loc_00E33E01: pop esi
  loc_00E33E02: mov fs:[00000000h], ecx
  loc_00E33E09: pop ebx
  loc_00E33E0A: mov esp, ebp
  loc_00E33E0C: pop ebp
  loc_00E33E0D: retn 0008h
End Sub

Private Sub back_Click() 'E32A80
  loc_00E32A80: push ebp
  loc_00E32A81: mov ebp, esp
  loc_00E32A83: sub esp, 0000000Ch
  loc_00E32A86: push 00402836h ; __vbaExceptHandler
  loc_00E32A8B: mov eax, fs:[00000000h]
  loc_00E32A91: push eax
  loc_00E32A92: mov fs:[00000000h], esp
  loc_00E32A99: sub esp, 00000034h
  loc_00E32A9C: push ebx
  loc_00E32A9D: push esi
  loc_00E32A9E: push edi
  loc_00E32A9F: mov var_C, esp
  loc_00E32AA2: mov var_8, 004025F0h
  loc_00E32AA9: mov eax, Me
  loc_00E32AAC: mov ecx, eax
  loc_00E32AAE: and ecx, 00000001h
  loc_00E32AB1: mov var_4, ecx
  loc_00E32AB4: and al, FEh
  loc_00E32AB6: push eax
  loc_00E32AB7: mov Me, eax
  loc_00E32ABA: mov edx, [eax]
  loc_00E32ABC: call [edx+00000004h]
  loc_00E32ABF: mov eax, [00E3D024h]
  loc_00E32AC4: mov var_18, 00000000h
  loc_00E32ACB: test eax, eax
  loc_00E32ACD: jnz 00E32ADFh
  loc_00E32ACF: push 00E3D024h
  loc_00E32AD4: push 006CE120h
  loc_00E32AD9: call [004011A0h] ; __vbaNew2
  loc_00E32ADF: sub esp, 00000010h
  loc_00E32AE2: mov ecx, 0000000Ah
  loc_00E32AE7: mov ebx, esp
  loc_00E32AE9: mov var_28, ecx
  loc_00E32AEC: mov eax, 80020004h
  loc_00E32AF1: sub esp, 00000010h
  loc_00E32AF4: mov [ebx], ecx
  loc_00E32AF6: mov ecx, var_34
  loc_00E32AF9: mov edx, eax
  loc_00E32AFB: mov esi, [00E3D024h]
  loc_00E32B01: mov [ebx+00000004h], ecx
  loc_00E32B04: mov ecx, esp
  loc_00E32B06: mov edi, [esi]
  loc_00E32B08: push esi
  loc_00E32B09: mov [ebx+00000008h], eax
  loc_00E32B0C: mov eax, var_2C
  loc_00E32B0F: mov [ebx+0000000Ch], eax
  loc_00E32B12: mov eax, var_28
  loc_00E32B15: mov [ecx], eax
  loc_00E32B17: mov eax, var_24
  loc_00E32B1A: mov [ecx+00000004h], eax
  loc_00E32B1D: mov [ecx+00000008h], edx
  loc_00E32B20: mov edx, var_1C
  loc_00E32B23: mov [ecx+0000000Ch], edx
  loc_00E32B26: call [edi+000002B0h]
  loc_00E32B2C: test eax, eax
  loc_00E32B2E: fnclex
  loc_00E32B30: jge 00E32B44h
  loc_00E32B32: push 000002B0h
  loc_00E32B37: push 006DC970h
  loc_00E32B3C: push esi
  loc_00E32B3D: push eax
  loc_00E32B3E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32B44: mov eax, [00E3D9CCh]
  loc_00E32B49: test eax, eax
  loc_00E32B4B: jnz 00E32B5Dh
  loc_00E32B4D: push 00E3D9CCh
  loc_00E32B52: push 006DC960h
  loc_00E32B57: call [004011A0h] ; __vbaNew2
  loc_00E32B5D: mov eax, Me
  loc_00E32B60: mov esi, [00E3D9CCh]
  loc_00E32B66: lea ecx, var_18
  loc_00E32B69: push eax
  loc_00E32B6A: mov edi, [esi]
  loc_00E32B6C: push ecx
  loc_00E32B6D: call [004010B4h] ; __vbaObjSetAddref
  loc_00E32B73: push eax
  loc_00E32B74: push esi
  loc_00E32B75: call [edi+00000010h]
  loc_00E32B78: test eax, eax
  loc_00E32B7A: fnclex
  loc_00E32B7C: jge 00E32B8Dh
  loc_00E32B7E: push 00000010h
  loc_00E32B80: push 006DC950h
  loc_00E32B85: push esi
  loc_00E32B86: push eax
  loc_00E32B87: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32B8D: lea ecx, var_18
  loc_00E32B90: call [00401254h] ; __vbaFreeObj
  loc_00E32B96: mov var_4, 00000000h
  loc_00E32B9D: push 00E32BAFh
  loc_00E32BA2: jmp 00E32BAEh
  loc_00E32BA4: lea ecx, var_18
  loc_00E32BA7: call [00401254h] ; __vbaFreeObj
  loc_00E32BAD: ret
  loc_00E32BAE: ret
  loc_00E32BAF: mov eax, Me
  loc_00E32BB2: push eax
  loc_00E32BB3: mov edx, [eax]
  loc_00E32BB5: call [edx+00000008h]
  loc_00E32BB8: mov eax, var_4
  loc_00E32BBB: mov ecx, var_14
  loc_00E32BBE: pop edi
  loc_00E32BBF: pop esi
  loc_00E32BC0: mov fs:[00000000h], ecx
  loc_00E32BC7: pop ebx
  loc_00E32BC8: mov esp, ebp
  loc_00E32BCA: pop ebp
  loc_00E32BCB: retn 0004h
End Sub

Private Sub jcbutton4_UnknownEvent_9 'E34000
  loc_00E34000: push ebp
  loc_00E34001: mov ebp, esp
  loc_00E34003: sub esp, 0000000Ch
  loc_00E34006: push 00402836h ; __vbaExceptHandler
  loc_00E3400B: mov eax, fs:[00000000h]
  loc_00E34011: push eax
  loc_00E34012: mov fs:[00000000h], esp
  loc_00E34019: sub esp, 00000018h
  loc_00E3401C: push ebx
  loc_00E3401D: push esi
  loc_00E3401E: push edi
  loc_00E3401F: mov var_C, esp
  loc_00E34022: mov var_8, 00402650h
  loc_00E34029: mov edi, Me
  loc_00E3402C: mov eax, edi
  loc_00E3402E: and eax, 00000001h
  loc_00E34031: mov var_4, eax
  loc_00E34034: and edi, FFFFFFFEh
  loc_00E34037: push edi
  loc_00E34038: mov Me, edi
  loc_00E3403B: mov ecx, [edi]
  loc_00E3403D: call [ecx+00000004h]
  loc_00E34040: mov ecx, [00E3D128h]
  loc_00E34046: xor eax, eax
  loc_00E34048: cmp ecx, eax
  loc_00E3404A: mov var_18, eax
  loc_00E3404D: mov var_1C, eax
  loc_00E34050: jnz 00E34062h
  loc_00E34052: push 00E3D128h
  loc_00E34057: push 006CB548h
  loc_00E3405C: call [004011A0h] ; __vbaNew2
  loc_00E34062: mov esi, [00E3D128h]
  loc_00E34068: mov edx, [edi]
  loc_00E3406A: push 006E05D8h
  loc_00E3406F: push edi
  loc_00E34070: mov ebx, [esi]
  loc_00E34072: call [edx+00000358h]
  loc_00E34078: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3407E: push eax
  loc_00E3407F: lea eax, var_18
  loc_00E34082: push eax
  loc_00E34083: call edi
  loc_00E34085: push eax
  loc_00E34086: call [00401224h] ; __vbaCastObj
  loc_00E3408C: lea ecx, var_1C
  loc_00E3408F: push eax
  loc_00E34090: push ecx
  loc_00E34091: call edi
  loc_00E34093: push eax
  loc_00E34094: push esi
  loc_00E34095: call [ebx+00000028h]
  loc_00E34098: test eax, eax
  loc_00E3409A: fnclex
  loc_00E3409C: jge 00E340ADh
  loc_00E3409E: push 00000028h
  loc_00E340A0: push 006DD034h
  loc_00E340A5: push esi
  loc_00E340A6: push eax
  loc_00E340A7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E340AD: lea edx, var_1C
  loc_00E340B0: lea eax, var_18
  loc_00E340B3: push edx
  loc_00E340B4: push eax
  loc_00E340B5: push 00000002h
  loc_00E340B7: call [00401048h] ; __vbaFreeObjList
  loc_00E340BD: mov eax, [00E3D128h]
  loc_00E340C2: add esp, 0000000Ch
  loc_00E340C5: test eax, eax
  loc_00E340C7: jnz 00E340DDh
  loc_00E340C9: mov edi, [004011A0h] ; __vbaNew2
  loc_00E340CF: push 00E3D128h
  loc_00E340D4: push 006CB548h
  loc_00E340D9: call edi
  loc_00E340DB: jmp 00E340E3h
  loc_00E340DD: mov edi, [004011A0h] ; __vbaNew2
  loc_00E340E3: mov esi, [00E3D128h]
  loc_00E340E9: push esi
  loc_00E340EA: mov ecx, [esi]
  loc_00E340EC: call [ecx+00000088h]
  loc_00E340F2: test eax, eax
  loc_00E340F4: fnclex
  loc_00E340F6: jge 00E3410Ah
  loc_00E340F8: push 00000088h
  loc_00E340FD: push 006DD034h
  loc_00E34102: push esi
  loc_00E34103: push eax
  loc_00E34104: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3410A: mov eax, [00E3D9CCh]
  loc_00E3410F: test eax, eax
  loc_00E34111: jnz 00E3411Fh
  loc_00E34113: push 00E3D9CCh
  loc_00E34118: push 006DC960h
  loc_00E3411D: call edi
  loc_00E3411F: mov eax, [00E3D128h]
  loc_00E34124: mov esi, [00E3D9CCh]
  loc_00E3412A: test eax, eax
  loc_00E3412C: jnz 00E3413Ah
  loc_00E3412E: push 00E3D128h
  loc_00E34133: push 006CB548h
  loc_00E34138: call edi
  loc_00E3413A: mov edx, [00E3D128h]
  loc_00E34140: mov ebx, [esi]
  loc_00E34142: lea eax, var_18
  loc_00E34145: push edx
  loc_00E34146: push eax
  loc_00E34147: call [004010B4h] ; __vbaObjSetAddref
  loc_00E3414D: push eax
  loc_00E3414E: push esi
  loc_00E3414F: call [ebx+0000000Ch]
  loc_00E34152: test eax, eax
  loc_00E34154: fnclex
  loc_00E34156: jge 00E34167h
  loc_00E34158: push 0000000Ch
  loc_00E3415A: push 006DC950h
  loc_00E3415F: push esi
  loc_00E34160: push eax
  loc_00E34161: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34167: lea ecx, var_18
  loc_00E3416A: call [00401254h] ; __vbaFreeObj
  loc_00E34170: mov eax, [00E3D128h]
  loc_00E34175: test eax, eax
  loc_00E34177: jnz 00E34185h
  loc_00E34179: push 00E3D128h
  loc_00E3417E: push 006CB548h
  loc_00E34183: call edi
  loc_00E34185: mov ecx, [00E3D128h]
  loc_00E3418B: push 00000000h
  loc_00E3418D: push 80011003h
  loc_00E34192: push ecx
  loc_00E34193: call [00401030h] ; __vbaLateIdCall
  loc_00E34199: add esp, 0000000Ch
  loc_00E3419C: mov var_4, 00000000h
  loc_00E341A3: push 00E341BFh
  loc_00E341A8: jmp 00E341BEh
  loc_00E341AA: lea edx, var_1C
  loc_00E341AD: lea eax, var_18
  loc_00E341B0: push edx
  loc_00E341B1: push eax
  loc_00E341B2: push 00000002h
  loc_00E341B4: call [00401048h] ; __vbaFreeObjList
  loc_00E341BA: add esp, 0000000Ch
  loc_00E341BD: ret
  loc_00E341BE: ret
  loc_00E341BF: mov eax, Me
  loc_00E341C2: push eax
  loc_00E341C3: mov ecx, [eax]
  loc_00E341C5: call [ecx+00000008h]
  loc_00E341C8: mov eax, var_4
  loc_00E341CB: mov ecx, var_14
  loc_00E341CE: pop edi
  loc_00E341CF: pop esi
  loc_00E341D0: mov fs:[00000000h], ecx
  loc_00E341D7: pop ebx
  loc_00E341D8: mov esp, ebp
  loc_00E341DA: pop ebp
  loc_00E341DB: retn 0004h
End Sub

Private Sub optagama_Click() 'E341E0
  loc_00E341E0: push ebp
  loc_00E341E1: mov ebp, esp
  loc_00E341E3: sub esp, 0000000Ch
  loc_00E341E6: push 00402836h ; __vbaExceptHandler
  loc_00E341EB: mov eax, fs:[00000000h]
  loc_00E341F1: push eax
  loc_00E341F2: mov fs:[00000000h], esp
  loc_00E341F9: sub esp, 00000048h
  loc_00E341FC: push ebx
  loc_00E341FD: push esi
  loc_00E341FE: push edi
  loc_00E341FF: mov var_C, esp
  loc_00E34202: mov var_8, 00402660h
  loc_00E34209: mov esi, Me
  loc_00E3420C: mov eax, esi
  loc_00E3420E: and eax, 00000001h
  loc_00E34211: mov var_4, eax
  loc_00E34214: and esi, FFFFFFFEh
  loc_00E34217: push esi
  loc_00E34218: mov Me, esi
  loc_00E3421B: mov ecx, [esi]
  loc_00E3421D: call [ecx+00000004h]
  loc_00E34220: sub esp, 00000010h
  loc_00E34223: mov ecx, 0000000Bh
  loc_00E34228: mov edx, esp
  loc_00E3422A: xor eax, eax
  loc_00E3422C: mov var_18, eax
  loc_00E3422F: mov var_1C, eax
  loc_00E34232: mov [edx], ecx
  loc_00E34234: mov ecx, var_38
  loc_00E34237: mov var_2C, eax
  loc_00E3423A: or eax, FFFFFFFFh
  loc_00E3423D: mov [edx+00000004h], ecx
  loc_00E34240: mov ecx, [esi]
  loc_00E34242: push 80010007h
  loc_00E34247: push esi
  loc_00E34248: mov [edx+00000008h], eax
  loc_00E3424B: mov eax, var_30
  loc_00E3424E: mov [edx+0000000Ch], eax
  loc_00E34251: call [ecx+00000334h]
  loc_00E34257: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3425D: lea edx, var_18
  loc_00E34260: push eax
  loc_00E34261: push edx
  loc_00E34262: call edi
  loc_00E34264: push eax
  loc_00E34265: call [00401238h] ; __vbaLateIdSt
  loc_00E3426B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E34271: lea ecx, var_18
  loc_00E34274: call ebx
  loc_00E34276: mov eax, [esi]
  loc_00E34278: push esi
  loc_00E34279: call [eax+00000328h]
  loc_00E3427F: lea ecx, var_18
  loc_00E34282: push eax
  loc_00E34283: push ecx
  loc_00E34284: call edi
  loc_00E34286: mov edx, [eax]
  loc_00E34288: push eax
  loc_00E34289: mov var_50, eax
  loc_00E3428C: call [edx+00000204h]
  loc_00E34292: test eax, eax
  loc_00E34294: fnclex
  loc_00E34296: jge 00E342ADh
  loc_00E34298: mov ecx, var_50
  loc_00E3429B: push 00000204h
  loc_00E342A0: push 006DCB70h
  loc_00E342A5: push ecx
  loc_00E342A6: push eax
  loc_00E342A7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E342AD: lea ecx, var_18
  loc_00E342B0: call ebx
  loc_00E342B2: mov edx, [esi]
  loc_00E342B4: push esi
  loc_00E342B5: call [edx+0000032Ch]
  loc_00E342BB: push eax
  loc_00E342BC: lea eax, var_18
  loc_00E342BF: push eax
  loc_00E342C0: call edi
  loc_00E342C2: mov ecx, [eax]
  loc_00E342C4: push 00000000h
  loc_00E342C6: push eax
  loc_00E342C7: mov var_50, eax
  loc_00E342CA: call [ecx+0000009Ch]
  loc_00E342D0: test eax, eax
  loc_00E342D2: fnclex
  loc_00E342D4: jge 00E342EBh
  loc_00E342D6: mov edx, var_50
  loc_00E342D9: push 0000009Ch
  loc_00E342DE: push 006E03D4h
  loc_00E342E3: push edx
  loc_00E342E4: push eax
  loc_00E342E5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E342EB: lea ecx, var_18
  loc_00E342EE: call ebx
  loc_00E342F0: mov eax, [esi]
  loc_00E342F2: push esi
  loc_00E342F3: call [eax+00000330h]
  loc_00E342F9: lea ecx, var_18
  loc_00E342FC: push eax
  loc_00E342FD: push ecx
  loc_00E342FE: call edi
  loc_00E34300: mov edx, [eax]
  loc_00E34302: push 00000000h
  loc_00E34304: push eax
  loc_00E34305: mov var_50, eax
  loc_00E34308: call [edx+0000009Ch]
  loc_00E3430E: test eax, eax
  loc_00E34310: fnclex
  loc_00E34312: jge 00E34329h
  loc_00E34314: mov ecx, var_50
  loc_00E34317: push 0000009Ch
  loc_00E3431C: push 006E03D4h
  loc_00E34321: push ecx
  loc_00E34322: push eax
  loc_00E34323: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34329: lea ecx, var_18
  loc_00E3432C: call ebx
  loc_00E3432E: mov edx, [esi]
  loc_00E34330: push esi
  loc_00E34331: call [edx+00000324h]
  loc_00E34337: push eax
  loc_00E34338: lea eax, var_18
  loc_00E3433B: push eax
  loc_00E3433C: call edi
  loc_00E3433E: mov ecx, [eax]
  loc_00E34340: push 00000000h
  loc_00E34342: push eax
  loc_00E34343: mov var_50, eax
  loc_00E34346: call [ecx+0000009Ch]
  loc_00E3434C: test eax, eax
  loc_00E3434E: fnclex
  loc_00E34350: jge 00E34367h
  loc_00E34352: mov edx, var_50
  loc_00E34355: push 0000009Ch
  loc_00E3435A: push 006E03D4h
  loc_00E3435F: push edx
  loc_00E34360: push eax
  loc_00E34361: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34367: lea ecx, var_18
  loc_00E3436A: call ebx
  loc_00E3436C: mov eax, [esi]
  loc_00E3436E: push esi
  loc_00E3436F: call [eax+00000314h]
  loc_00E34375: lea ecx, var_18
  loc_00E34378: push eax
  loc_00E34379: push ecx
  loc_00E3437A: call edi
  loc_00E3437C: mov edx, [eax]
  loc_00E3437E: push 00000000h
  loc_00E34380: push eax
  loc_00E34381: mov var_50, eax
  loc_00E34384: call [edx+0000009Ch]
  loc_00E3438A: test eax, eax
  loc_00E3438C: fnclex
  loc_00E3438E: jge 00E343A5h
  loc_00E34390: mov ecx, var_50
  loc_00E34393: push 0000009Ch
  loc_00E34398: push 006E03D4h
  loc_00E3439D: push ecx
  loc_00E3439E: push eax
  loc_00E3439F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E343A5: lea ecx, var_18
  loc_00E343A8: call ebx
  loc_00E343AA: mov edx, [esi]
  loc_00E343AC: push esi
  loc_00E343AD: call [edx+00000320h]
  loc_00E343B3: push eax
  loc_00E343B4: lea eax, var_18
  loc_00E343B7: push eax
  loc_00E343B8: call edi
  loc_00E343BA: mov ecx, [eax]
  loc_00E343BC: push 00000000h
  loc_00E343BE: push eax
  loc_00E343BF: mov var_50, eax
  loc_00E343C2: call [ecx+0000009Ch]
  loc_00E343C8: test eax, eax
  loc_00E343CA: fnclex
  loc_00E343CC: jge 00E343E3h
  loc_00E343CE: mov edx, var_50
  loc_00E343D1: push 0000009Ch
  loc_00E343D6: push 006E03D4h
  loc_00E343DB: push edx
  loc_00E343DC: push eax
  loc_00E343DD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E343E3: lea ecx, var_18
  loc_00E343E6: call ebx
  loc_00E343E8: sub esp, 00000010h
  loc_00E343EB: mov ecx, 00000008h
  loc_00E343F0: mov edx, esp
  loc_00E343F2: mov eax, 006E0A78h ; "biodata"
  loc_00E343F7: push 0000000Eh
  loc_00E343F9: push esi
  loc_00E343FA: mov [edx], ecx
  loc_00E343FC: mov ecx, var_38
  loc_00E343FF: mov [edx+00000004h], ecx
  loc_00E34402: mov ecx, [esi]
  loc_00E34404: mov [edx+00000008h], eax
  loc_00E34407: mov eax, var_30
  loc_00E3440A: mov [edx+0000000Ch], eax
  loc_00E3440D: call [ecx+00000358h]
  loc_00E34413: lea edx, var_18
  loc_00E34416: push eax
  loc_00E34417: push edx
  loc_00E34418: call edi
  loc_00E3441A: push eax
  loc_00E3441B: call [00401238h] ; __vbaLateIdSt
  loc_00E34421: lea ecx, var_18
  loc_00E34424: call ebx
  loc_00E34426: mov eax, [esi]
  loc_00E34428: push 00000000h
  loc_00E3442A: push 00000019h
  loc_00E3442C: push esi
  loc_00E3442D: call [eax+00000358h]
  loc_00E34433: lea ecx, var_18
  loc_00E34436: push eax
  loc_00E34437: push ecx
  loc_00E34438: call edi
  loc_00E3443A: push eax
  loc_00E3443B: call [00401030h] ; __vbaLateIdCall
  loc_00E34441: add esp, 0000000Ch
  loc_00E34444: lea ecx, var_18
  loc_00E34447: call ebx
  loc_00E34449: mov edx, [esi]
  loc_00E3444B: push 006E05D8h
  loc_00E34450: push esi
  loc_00E34451: call [edx+00000358h]
  loc_00E34457: push eax
  loc_00E34458: lea eax, var_18
  loc_00E3445B: push eax
  loc_00E3445C: call edi
  loc_00E3445E: push eax
  loc_00E3445F: call [00401224h] ; __vbaCastObj
  loc_00E34465: lea ecx, var_2C
  loc_00E34468: mov var_24, eax
  loc_00E3446B: push ecx
  loc_00E3446C: mov var_2C, 0000000Dh
  loc_00E34473: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E34479: mov ecx, [eax]
  loc_00E3447B: sub esp, 00000010h
  loc_00E3447E: mov edx, esp
  loc_00E34480: push 00000000h
  loc_00E34482: push 0000002Ah
  loc_00E34484: mov [edx], ecx
  loc_00E34486: mov ecx, [eax+00000004h]
  loc_00E34489: push esi
  loc_00E3448A: mov [edx+00000004h], ecx
  loc_00E3448D: mov ecx, [eax+00000008h]
  loc_00E34490: mov eax, [eax+0000000Ch]
  loc_00E34493: mov [edx+00000008h], ecx
  loc_00E34496: mov ecx, [esi]
  loc_00E34498: mov [edx+0000000Ch], eax
  loc_00E3449B: call [ecx+0000035Ch]
  loc_00E344A1: lea edx, var_1C
  loc_00E344A4: push eax
  loc_00E344A5: push edx
  loc_00E344A6: call edi
  loc_00E344A8: push eax
  loc_00E344A9: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E344AF: lea eax, var_1C
  loc_00E344B2: lea ecx, var_18
  loc_00E344B5: push eax
  loc_00E344B6: push ecx
  loc_00E344B7: push 00000002h
  loc_00E344B9: call [00401048h] ; __vbaFreeObjList
  loc_00E344BF: add esp, 00000028h
  loc_00E344C2: lea ecx, var_2C
  loc_00E344C5: call [00401024h] ; __vbaFreeVar
  loc_00E344CB: sub esp, 00000010h
  loc_00E344CE: mov ecx, 0000000Bh
  loc_00E344D3: mov edx, esp
  loc_00E344D5: xor eax, eax
  loc_00E344D7: push 00000006h
  loc_00E344D9: push esi
  loc_00E344DA: mov [edx], ecx
  loc_00E344DC: mov ecx, var_38
  loc_00E344DF: mov [edx+00000004h], ecx
  loc_00E344E2: mov ecx, [esi]
  loc_00E344E4: mov [edx+00000008h], eax
  loc_00E344E7: mov eax, var_30
  loc_00E344EA: mov [edx+0000000Ch], eax
  loc_00E344ED: call [ecx+0000035Ch]
  loc_00E344F3: lea edx, var_18
  loc_00E344F6: push eax
  loc_00E344F7: push edx
  loc_00E344F8: call edi
  loc_00E344FA: push eax
  loc_00E344FB: call [00401238h] ; __vbaLateIdSt
  loc_00E34501: lea ecx, var_18
  loc_00E34504: call ebx
  loc_00E34506: sub esp, 00000010h
  loc_00E34509: mov ecx, 0000000Bh
  loc_00E3450E: mov edx, esp
  loc_00E34510: xor eax, eax
  loc_00E34512: push 8001000Eh
  loc_00E34517: push esi
  loc_00E34518: mov [edx], ecx
  loc_00E3451A: mov ecx, var_38
  loc_00E3451D: mov [edx+00000004h], ecx
  loc_00E34520: mov ecx, [esi]
  loc_00E34522: mov [edx+00000008h], eax
  loc_00E34525: mov eax, var_30
  loc_00E34528: mov [edx+0000000Ch], eax
  loc_00E3452B: call [ecx+0000035Ch]
  loc_00E34531: lea edx, var_18
  loc_00E34534: push eax
  loc_00E34535: push edx
  loc_00E34536: call edi
  loc_00E34538: push eax
  loc_00E34539: call [00401238h] ; __vbaLateIdSt
  loc_00E3453F: lea ecx, var_18
  loc_00E34542: call ebx
  loc_00E34544: mov eax, [esi]
  loc_00E34546: push 00000000h
  loc_00E34548: push FFFFFDDAh
  loc_00E3454D: push esi
  loc_00E3454E: call [eax+0000035Ch]
  loc_00E34554: lea ecx, var_18
  loc_00E34557: push eax
  loc_00E34558: push ecx
  loc_00E34559: call edi
  loc_00E3455B: push eax
  loc_00E3455C: call [00401030h] ; __vbaLateIdCall
  loc_00E34562: add esp, 0000000Ch
  loc_00E34565: lea ecx, var_18
  loc_00E34568: call ebx
  loc_00E3456A: mov var_4, 00000000h
  loc_00E34571: push 00E34596h
  loc_00E34576: jmp 00E34595h
  loc_00E34578: lea edx, var_1C
  loc_00E3457B: lea eax, var_18
  loc_00E3457E: push edx
  loc_00E3457F: push eax
  loc_00E34580: push 00000002h
  loc_00E34582: call [00401048h] ; __vbaFreeObjList
  loc_00E34588: add esp, 0000000Ch
  loc_00E3458B: lea ecx, var_2C
  loc_00E3458E: call [00401024h] ; __vbaFreeVar
  loc_00E34594: ret
  loc_00E34595: ret
  loc_00E34596: mov eax, Me
  loc_00E34599: push eax
  loc_00E3459A: mov ecx, [eax]
  loc_00E3459C: call [ecx+00000008h]
  loc_00E3459F: mov eax, var_4
  loc_00E345A2: mov ecx, var_14
  loc_00E345A5: pop edi
  loc_00E345A6: pop esi
  loc_00E345A7: mov fs:[00000000h], ecx
  loc_00E345AE: pop ebx
  loc_00E345AF: mov esp, ebp
  loc_00E345B1: pop ebp
  loc_00E345B2: retn 0004h
End Sub

Private Sub optjur_Click() 'E349A0
  loc_00E349A0: push ebp
  loc_00E349A1: mov ebp, esp
  loc_00E349A3: sub esp, 0000000Ch
  loc_00E349A6: push 00402836h ; __vbaExceptHandler
  loc_00E349AB: mov eax, fs:[00000000h]
  loc_00E349B1: push eax
  loc_00E349B2: mov fs:[00000000h], esp
  loc_00E349B9: sub esp, 00000048h
  loc_00E349BC: push ebx
  loc_00E349BD: push esi
  loc_00E349BE: push edi
  loc_00E349BF: mov var_C, esp
  loc_00E349C2: mov var_8, 00402680h
  loc_00E349C9: mov esi, Me
  loc_00E349CC: mov eax, esi
  loc_00E349CE: and eax, 00000001h
  loc_00E349D1: mov var_4, eax
  loc_00E349D4: and esi, FFFFFFFEh
  loc_00E349D7: push esi
  loc_00E349D8: mov Me, esi
  loc_00E349DB: mov ecx, [esi]
  loc_00E349DD: call [ecx+00000004h]
  loc_00E349E0: sub esp, 00000010h
  loc_00E349E3: mov ecx, 0000000Bh
  loc_00E349E8: mov edx, esp
  loc_00E349EA: xor eax, eax
  loc_00E349EC: mov var_18, eax
  loc_00E349EF: mov var_1C, eax
  loc_00E349F2: mov [edx], ecx
  loc_00E349F4: mov ecx, var_38
  loc_00E349F7: mov var_2C, eax
  loc_00E349FA: or eax, FFFFFFFFh
  loc_00E349FD: mov [edx+00000004h], ecx
  loc_00E34A00: mov ecx, [esi]
  loc_00E34A02: push 80010007h
  loc_00E34A07: push esi
  loc_00E34A08: mov [edx+00000008h], eax
  loc_00E34A0B: mov eax, var_30
  loc_00E34A0E: mov [edx+0000000Ch], eax
  loc_00E34A11: call [ecx+00000334h]
  loc_00E34A17: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E34A1D: lea edx, var_18
  loc_00E34A20: push eax
  loc_00E34A21: push edx
  loc_00E34A22: call edi
  loc_00E34A24: push eax
  loc_00E34A25: call [00401238h] ; __vbaLateIdSt
  loc_00E34A2B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E34A31: lea ecx, var_18
  loc_00E34A34: call ebx
  loc_00E34A36: mov eax, [esi]
  loc_00E34A38: push esi
  loc_00E34A39: call [eax+00000328h]
  loc_00E34A3F: lea ecx, var_18
  loc_00E34A42: push eax
  loc_00E34A43: push ecx
  loc_00E34A44: call edi
  loc_00E34A46: mov edx, [eax]
  loc_00E34A48: push eax
  loc_00E34A49: mov var_50, eax
  loc_00E34A4C: call [edx+00000204h]
  loc_00E34A52: test eax, eax
  loc_00E34A54: fnclex
  loc_00E34A56: jge 00E34A6Dh
  loc_00E34A58: mov ecx, var_50
  loc_00E34A5B: push 00000204h
  loc_00E34A60: push 006DCB70h
  loc_00E34A65: push ecx
  loc_00E34A66: push eax
  loc_00E34A67: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34A6D: lea ecx, var_18
  loc_00E34A70: call ebx
  loc_00E34A72: mov edx, [esi]
  loc_00E34A74: push esi
  loc_00E34A75: call [edx+0000032Ch]
  loc_00E34A7B: push eax
  loc_00E34A7C: lea eax, var_18
  loc_00E34A7F: push eax
  loc_00E34A80: call edi
  loc_00E34A82: mov ecx, [eax]
  loc_00E34A84: push 00000000h
  loc_00E34A86: push eax
  loc_00E34A87: mov var_50, eax
  loc_00E34A8A: call [ecx+0000009Ch]
  loc_00E34A90: test eax, eax
  loc_00E34A92: fnclex
  loc_00E34A94: jge 00E34AABh
  loc_00E34A96: mov edx, var_50
  loc_00E34A99: push 0000009Ch
  loc_00E34A9E: push 006E03D4h
  loc_00E34AA3: push edx
  loc_00E34AA4: push eax
  loc_00E34AA5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34AAB: lea ecx, var_18
  loc_00E34AAE: call ebx
  loc_00E34AB0: mov eax, [esi]
  loc_00E34AB2: push esi
  loc_00E34AB3: call [eax+00000330h]
  loc_00E34AB9: lea ecx, var_18
  loc_00E34ABC: push eax
  loc_00E34ABD: push ecx
  loc_00E34ABE: call edi
  loc_00E34AC0: mov edx, [eax]
  loc_00E34AC2: push 00000000h
  loc_00E34AC4: push eax
  loc_00E34AC5: mov var_50, eax
  loc_00E34AC8: call [edx+0000009Ch]
  loc_00E34ACE: test eax, eax
  loc_00E34AD0: fnclex
  loc_00E34AD2: jge 00E34AE9h
  loc_00E34AD4: mov ecx, var_50
  loc_00E34AD7: push 0000009Ch
  loc_00E34ADC: push 006E03D4h
  loc_00E34AE1: push ecx
  loc_00E34AE2: push eax
  loc_00E34AE3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34AE9: lea ecx, var_18
  loc_00E34AEC: call ebx
  loc_00E34AEE: mov edx, [esi]
  loc_00E34AF0: push esi
  loc_00E34AF1: call [edx+00000324h]
  loc_00E34AF7: push eax
  loc_00E34AF8: lea eax, var_18
  loc_00E34AFB: push eax
  loc_00E34AFC: call edi
  loc_00E34AFE: mov ecx, [eax]
  loc_00E34B00: push 00000000h
  loc_00E34B02: push eax
  loc_00E34B03: mov var_50, eax
  loc_00E34B06: call [ecx+0000009Ch]
  loc_00E34B0C: test eax, eax
  loc_00E34B0E: fnclex
  loc_00E34B10: jge 00E34B27h
  loc_00E34B12: mov edx, var_50
  loc_00E34B15: push 0000009Ch
  loc_00E34B1A: push 006E03D4h
  loc_00E34B1F: push edx
  loc_00E34B20: push eax
  loc_00E34B21: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34B27: lea ecx, var_18
  loc_00E34B2A: call ebx
  loc_00E34B2C: mov eax, [esi]
  loc_00E34B2E: push esi
  loc_00E34B2F: call [eax+0000031Ch]
  loc_00E34B35: lea ecx, var_18
  loc_00E34B38: push eax
  loc_00E34B39: push ecx
  loc_00E34B3A: call edi
  loc_00E34B3C: mov edx, [eax]
  loc_00E34B3E: push 00000000h
  loc_00E34B40: push eax
  loc_00E34B41: mov var_50, eax
  loc_00E34B44: call [edx+0000009Ch]
  loc_00E34B4A: test eax, eax
  loc_00E34B4C: fnclex
  loc_00E34B4E: jge 00E34B65h
  loc_00E34B50: mov ecx, var_50
  loc_00E34B53: push 0000009Ch
  loc_00E34B58: push 006E03D4h
  loc_00E34B5D: push ecx
  loc_00E34B5E: push eax
  loc_00E34B5F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34B65: lea ecx, var_18
  loc_00E34B68: call ebx
  loc_00E34B6A: mov edx, [esi]
  loc_00E34B6C: push esi
  loc_00E34B6D: call [edx+00000320h]
  loc_00E34B73: push eax
  loc_00E34B74: lea eax, var_18
  loc_00E34B77: push eax
  loc_00E34B78: call edi
  loc_00E34B7A: mov ecx, [eax]
  loc_00E34B7C: push 00000000h
  loc_00E34B7E: push eax
  loc_00E34B7F: mov var_50, eax
  loc_00E34B82: call [ecx+0000009Ch]
  loc_00E34B88: test eax, eax
  loc_00E34B8A: fnclex
  loc_00E34B8C: jge 00E34BA3h
  loc_00E34B8E: mov edx, var_50
  loc_00E34B91: push 0000009Ch
  loc_00E34B96: push 006E03D4h
  loc_00E34B9B: push edx
  loc_00E34B9C: push eax
  loc_00E34B9D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34BA3: lea ecx, var_18
  loc_00E34BA6: call ebx
  loc_00E34BA8: sub esp, 00000010h
  loc_00E34BAB: mov ecx, 00000008h
  loc_00E34BB0: mov edx, esp
  loc_00E34BB2: mov eax, 006E0A78h ; "biodata"
  loc_00E34BB7: push 0000000Eh
  loc_00E34BB9: push esi
  loc_00E34BBA: mov [edx], ecx
  loc_00E34BBC: mov ecx, var_38
  loc_00E34BBF: mov [edx+00000004h], ecx
  loc_00E34BC2: mov ecx, [esi]
  loc_00E34BC4: mov [edx+00000008h], eax
  loc_00E34BC7: mov eax, var_30
  loc_00E34BCA: mov [edx+0000000Ch], eax
  loc_00E34BCD: call [ecx+00000358h]
  loc_00E34BD3: lea edx, var_18
  loc_00E34BD6: push eax
  loc_00E34BD7: push edx
  loc_00E34BD8: call edi
  loc_00E34BDA: push eax
  loc_00E34BDB: call [00401238h] ; __vbaLateIdSt
  loc_00E34BE1: lea ecx, var_18
  loc_00E34BE4: call ebx
  loc_00E34BE6: mov eax, [esi]
  loc_00E34BE8: push 00000000h
  loc_00E34BEA: push 00000019h
  loc_00E34BEC: push esi
  loc_00E34BED: call [eax+00000358h]
  loc_00E34BF3: lea ecx, var_18
  loc_00E34BF6: push eax
  loc_00E34BF7: push ecx
  loc_00E34BF8: call edi
  loc_00E34BFA: push eax
  loc_00E34BFB: call [00401030h] ; __vbaLateIdCall
  loc_00E34C01: add esp, 0000000Ch
  loc_00E34C04: lea ecx, var_18
  loc_00E34C07: call ebx
  loc_00E34C09: mov edx, [esi]
  loc_00E34C0B: push 006E05D8h
  loc_00E34C10: push esi
  loc_00E34C11: call [edx+00000358h]
  loc_00E34C17: push eax
  loc_00E34C18: lea eax, var_18
  loc_00E34C1B: push eax
  loc_00E34C1C: call edi
  loc_00E34C1E: push eax
  loc_00E34C1F: call [00401224h] ; __vbaCastObj
  loc_00E34C25: lea ecx, var_2C
  loc_00E34C28: mov var_24, eax
  loc_00E34C2B: push ecx
  loc_00E34C2C: mov var_2C, 0000000Dh
  loc_00E34C33: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E34C39: mov ecx, [eax]
  loc_00E34C3B: sub esp, 00000010h
  loc_00E34C3E: mov edx, esp
  loc_00E34C40: push 00000000h
  loc_00E34C42: push 0000002Ah
  loc_00E34C44: mov [edx], ecx
  loc_00E34C46: mov ecx, [eax+00000004h]
  loc_00E34C49: push esi
  loc_00E34C4A: mov [edx+00000004h], ecx
  loc_00E34C4D: mov ecx, [eax+00000008h]
  loc_00E34C50: mov eax, [eax+0000000Ch]
  loc_00E34C53: mov [edx+00000008h], ecx
  loc_00E34C56: mov ecx, [esi]
  loc_00E34C58: mov [edx+0000000Ch], eax
  loc_00E34C5B: call [ecx+0000035Ch]
  loc_00E34C61: lea edx, var_1C
  loc_00E34C64: push eax
  loc_00E34C65: push edx
  loc_00E34C66: call edi
  loc_00E34C68: push eax
  loc_00E34C69: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E34C6F: lea eax, var_1C
  loc_00E34C72: lea ecx, var_18
  loc_00E34C75: push eax
  loc_00E34C76: push ecx
  loc_00E34C77: push 00000002h
  loc_00E34C79: call [00401048h] ; __vbaFreeObjList
  loc_00E34C7F: add esp, 00000028h
  loc_00E34C82: lea ecx, var_2C
  loc_00E34C85: call [00401024h] ; __vbaFreeVar
  loc_00E34C8B: sub esp, 00000010h
  loc_00E34C8E: mov ecx, 0000000Bh
  loc_00E34C93: mov edx, esp
  loc_00E34C95: xor eax, eax
  loc_00E34C97: push 00000006h
  loc_00E34C99: push esi
  loc_00E34C9A: mov [edx], ecx
  loc_00E34C9C: mov ecx, var_38
  loc_00E34C9F: mov [edx+00000004h], ecx
  loc_00E34CA2: mov ecx, [esi]
  loc_00E34CA4: mov [edx+00000008h], eax
  loc_00E34CA7: mov eax, var_30
  loc_00E34CAA: mov [edx+0000000Ch], eax
  loc_00E34CAD: call [ecx+0000035Ch]
  loc_00E34CB3: lea edx, var_18
  loc_00E34CB6: push eax
  loc_00E34CB7: push edx
  loc_00E34CB8: call edi
  loc_00E34CBA: push eax
  loc_00E34CBB: call [00401238h] ; __vbaLateIdSt
  loc_00E34CC1: lea ecx, var_18
  loc_00E34CC4: call ebx
  loc_00E34CC6: sub esp, 00000010h
  loc_00E34CC9: mov ecx, 0000000Bh
  loc_00E34CCE: mov edx, esp
  loc_00E34CD0: xor eax, eax
  loc_00E34CD2: push 8001000Eh
  loc_00E34CD7: push esi
  loc_00E34CD8: mov [edx], ecx
  loc_00E34CDA: mov ecx, var_38
  loc_00E34CDD: mov [edx+00000004h], ecx
  loc_00E34CE0: mov ecx, [esi]
  loc_00E34CE2: mov [edx+00000008h], eax
  loc_00E34CE5: mov eax, var_30
  loc_00E34CE8: mov [edx+0000000Ch], eax
  loc_00E34CEB: call [ecx+0000035Ch]
  loc_00E34CF1: lea edx, var_18
  loc_00E34CF4: push eax
  loc_00E34CF5: push edx
  loc_00E34CF6: call edi
  loc_00E34CF8: push eax
  loc_00E34CF9: call [00401238h] ; __vbaLateIdSt
  loc_00E34CFF: lea ecx, var_18
  loc_00E34D02: call ebx
  loc_00E34D04: mov eax, [esi]
  loc_00E34D06: push 00000000h
  loc_00E34D08: push FFFFFDDAh
  loc_00E34D0D: push esi
  loc_00E34D0E: call [eax+0000035Ch]
  loc_00E34D14: lea ecx, var_18
  loc_00E34D17: push eax
  loc_00E34D18: push ecx
  loc_00E34D19: call edi
  loc_00E34D1B: push eax
  loc_00E34D1C: call [00401030h] ; __vbaLateIdCall
  loc_00E34D22: add esp, 0000000Ch
  loc_00E34D25: lea ecx, var_18
  loc_00E34D28: call ebx
  loc_00E34D2A: mov var_4, 00000000h
  loc_00E34D31: push 00E34D56h
  loc_00E34D36: jmp 00E34D55h
  loc_00E34D38: lea edx, var_1C
  loc_00E34D3B: lea eax, var_18
  loc_00E34D3E: push edx
  loc_00E34D3F: push eax
  loc_00E34D40: push 00000002h
  loc_00E34D42: call [00401048h] ; __vbaFreeObjList
  loc_00E34D48: add esp, 0000000Ch
  loc_00E34D4B: lea ecx, var_2C
  loc_00E34D4E: call [00401024h] ; __vbaFreeVar
  loc_00E34D54: ret
  loc_00E34D55: ret
  loc_00E34D56: mov eax, Me
  loc_00E34D59: push eax
  loc_00E34D5A: mov ecx, [eax]
  loc_00E34D5C: call [ecx+00000008h]
  loc_00E34D5F: mov eax, var_4
  loc_00E34D62: mov ecx, var_14
  loc_00E34D65: pop edi
  loc_00E34D66: pop esi
  loc_00E34D67: mov fs:[00000000h], ecx
  loc_00E34D6E: pop ebx
  loc_00E34D6F: mov esp, ebp
  loc_00E34D71: pop ebp
  loc_00E34D72: retn 0004h
End Sub

Private Sub optasal_Click() 'E345C0
  loc_00E345C0: push ebp
  loc_00E345C1: mov ebp, esp
  loc_00E345C3: sub esp, 0000000Ch
  loc_00E345C6: push 00402836h ; __vbaExceptHandler
  loc_00E345CB: mov eax, fs:[00000000h]
  loc_00E345D1: push eax
  loc_00E345D2: mov fs:[00000000h], esp
  loc_00E345D9: sub esp, 00000048h
  loc_00E345DC: push ebx
  loc_00E345DD: push esi
  loc_00E345DE: push edi
  loc_00E345DF: mov var_C, esp
  loc_00E345E2: mov var_8, 00402670h
  loc_00E345E9: mov esi, Me
  loc_00E345EC: mov eax, esi
  loc_00E345EE: and eax, 00000001h
  loc_00E345F1: mov var_4, eax
  loc_00E345F4: and esi, FFFFFFFEh
  loc_00E345F7: push esi
  loc_00E345F8: mov Me, esi
  loc_00E345FB: mov ecx, [esi]
  loc_00E345FD: call [ecx+00000004h]
  loc_00E34600: sub esp, 00000010h
  loc_00E34603: mov ecx, 0000000Bh
  loc_00E34608: mov edx, esp
  loc_00E3460A: xor eax, eax
  loc_00E3460C: mov var_18, eax
  loc_00E3460F: mov var_1C, eax
  loc_00E34612: mov [edx], ecx
  loc_00E34614: mov ecx, var_38
  loc_00E34617: mov var_2C, eax
  loc_00E3461A: or eax, FFFFFFFFh
  loc_00E3461D: mov [edx+00000004h], ecx
  loc_00E34620: mov ecx, [esi]
  loc_00E34622: push 80010007h
  loc_00E34627: push esi
  loc_00E34628: mov [edx+00000008h], eax
  loc_00E3462B: mov eax, var_30
  loc_00E3462E: mov [edx+0000000Ch], eax
  loc_00E34631: call [ecx+00000334h]
  loc_00E34637: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3463D: lea edx, var_18
  loc_00E34640: push eax
  loc_00E34641: push edx
  loc_00E34642: call edi
  loc_00E34644: push eax
  loc_00E34645: call [00401238h] ; __vbaLateIdSt
  loc_00E3464B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E34651: lea ecx, var_18
  loc_00E34654: call ebx
  loc_00E34656: mov eax, [esi]
  loc_00E34658: push esi
  loc_00E34659: call [eax+00000328h]
  loc_00E3465F: lea ecx, var_18
  loc_00E34662: push eax
  loc_00E34663: push ecx
  loc_00E34664: call edi
  loc_00E34666: mov edx, [eax]
  loc_00E34668: push eax
  loc_00E34669: mov var_50, eax
  loc_00E3466C: call [edx+00000204h]
  loc_00E34672: test eax, eax
  loc_00E34674: fnclex
  loc_00E34676: jge 00E3468Dh
  loc_00E34678: mov ecx, var_50
  loc_00E3467B: push 00000204h
  loc_00E34680: push 006DCB70h
  loc_00E34685: push ecx
  loc_00E34686: push eax
  loc_00E34687: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3468D: lea ecx, var_18
  loc_00E34690: call ebx
  loc_00E34692: mov edx, [esi]
  loc_00E34694: push esi
  loc_00E34695: call [edx+0000032Ch]
  loc_00E3469B: push eax
  loc_00E3469C: lea eax, var_18
  loc_00E3469F: push eax
  loc_00E346A0: call edi
  loc_00E346A2: mov ecx, [eax]
  loc_00E346A4: push 00000000h
  loc_00E346A6: push eax
  loc_00E346A7: mov var_50, eax
  loc_00E346AA: call [ecx+0000009Ch]
  loc_00E346B0: test eax, eax
  loc_00E346B2: fnclex
  loc_00E346B4: jge 00E346CBh
  loc_00E346B6: mov edx, var_50
  loc_00E346B9: push 0000009Ch
  loc_00E346BE: push 006E03D4h
  loc_00E346C3: push edx
  loc_00E346C4: push eax
  loc_00E346C5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E346CB: lea ecx, var_18
  loc_00E346CE: call ebx
  loc_00E346D0: mov eax, [esi]
  loc_00E346D2: push esi
  loc_00E346D3: call [eax+00000330h]
  loc_00E346D9: lea ecx, var_18
  loc_00E346DC: push eax
  loc_00E346DD: push ecx
  loc_00E346DE: call edi
  loc_00E346E0: mov edx, [eax]
  loc_00E346E2: push 00000000h
  loc_00E346E4: push eax
  loc_00E346E5: mov var_50, eax
  loc_00E346E8: call [edx+0000009Ch]
  loc_00E346EE: test eax, eax
  loc_00E346F0: fnclex
  loc_00E346F2: jge 00E34709h
  loc_00E346F4: mov ecx, var_50
  loc_00E346F7: push 0000009Ch
  loc_00E346FC: push 006E03D4h
  loc_00E34701: push ecx
  loc_00E34702: push eax
  loc_00E34703: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34709: lea ecx, var_18
  loc_00E3470C: call ebx
  loc_00E3470E: mov edx, [esi]
  loc_00E34710: push esi
  loc_00E34711: call [edx+00000324h]
  loc_00E34717: push eax
  loc_00E34718: lea eax, var_18
  loc_00E3471B: push eax
  loc_00E3471C: call edi
  loc_00E3471E: mov ecx, [eax]
  loc_00E34720: push 00000000h
  loc_00E34722: push eax
  loc_00E34723: mov var_50, eax
  loc_00E34726: call [ecx+0000009Ch]
  loc_00E3472C: test eax, eax
  loc_00E3472E: fnclex
  loc_00E34730: jge 00E34747h
  loc_00E34732: mov edx, var_50
  loc_00E34735: push 0000009Ch
  loc_00E3473A: push 006E03D4h
  loc_00E3473F: push edx
  loc_00E34740: push eax
  loc_00E34741: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34747: lea ecx, var_18
  loc_00E3474A: call ebx
  loc_00E3474C: mov eax, [esi]
  loc_00E3474E: push esi
  loc_00E3474F: call [eax+00000314h]
  loc_00E34755: lea ecx, var_18
  loc_00E34758: push eax
  loc_00E34759: push ecx
  loc_00E3475A: call edi
  loc_00E3475C: mov edx, [eax]
  loc_00E3475E: push 00000000h
  loc_00E34760: push eax
  loc_00E34761: mov var_50, eax
  loc_00E34764: call [edx+0000009Ch]
  loc_00E3476A: test eax, eax
  loc_00E3476C: fnclex
  loc_00E3476E: jge 00E34785h
  loc_00E34770: mov ecx, var_50
  loc_00E34773: push 0000009Ch
  loc_00E34778: push 006E03D4h
  loc_00E3477D: push ecx
  loc_00E3477E: push eax
  loc_00E3477F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34785: lea ecx, var_18
  loc_00E34788: call ebx
  loc_00E3478A: mov edx, [esi]
  loc_00E3478C: push esi
  loc_00E3478D: call [edx+0000031Ch]
  loc_00E34793: push eax
  loc_00E34794: lea eax, var_18
  loc_00E34797: push eax
  loc_00E34798: call edi
  loc_00E3479A: mov ecx, [eax]
  loc_00E3479C: push 00000000h
  loc_00E3479E: push eax
  loc_00E3479F: mov var_50, eax
  loc_00E347A2: call [ecx+0000009Ch]
  loc_00E347A8: test eax, eax
  loc_00E347AA: fnclex
  loc_00E347AC: jge 00E347C3h
  loc_00E347AE: mov edx, var_50
  loc_00E347B1: push 0000009Ch
  loc_00E347B6: push 006E03D4h
  loc_00E347BB: push edx
  loc_00E347BC: push eax
  loc_00E347BD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E347C3: lea ecx, var_18
  loc_00E347C6: call ebx
  loc_00E347C8: sub esp, 00000010h
  loc_00E347CB: mov ecx, 00000008h
  loc_00E347D0: mov edx, esp
  loc_00E347D2: mov eax, 006E0A78h ; "biodata"
  loc_00E347D7: push 0000000Eh
  loc_00E347D9: push esi
  loc_00E347DA: mov [edx], ecx
  loc_00E347DC: mov ecx, var_38
  loc_00E347DF: mov [edx+00000004h], ecx
  loc_00E347E2: mov ecx, [esi]
  loc_00E347E4: mov [edx+00000008h], eax
  loc_00E347E7: mov eax, var_30
  loc_00E347EA: mov [edx+0000000Ch], eax
  loc_00E347ED: call [ecx+00000358h]
  loc_00E347F3: lea edx, var_18
  loc_00E347F6: push eax
  loc_00E347F7: push edx
  loc_00E347F8: call edi
  loc_00E347FA: push eax
  loc_00E347FB: call [00401238h] ; __vbaLateIdSt
  loc_00E34801: lea ecx, var_18
  loc_00E34804: call ebx
  loc_00E34806: mov eax, [esi]
  loc_00E34808: push 00000000h
  loc_00E3480A: push 00000019h
  loc_00E3480C: push esi
  loc_00E3480D: call [eax+00000358h]
  loc_00E34813: lea ecx, var_18
  loc_00E34816: push eax
  loc_00E34817: push ecx
  loc_00E34818: call edi
  loc_00E3481A: push eax
  loc_00E3481B: call [00401030h] ; __vbaLateIdCall
  loc_00E34821: add esp, 0000000Ch
  loc_00E34824: lea ecx, var_18
  loc_00E34827: call ebx
  loc_00E34829: mov edx, [esi]
  loc_00E3482B: push 006E05D8h
  loc_00E34830: push esi
  loc_00E34831: call [edx+00000358h]
  loc_00E34837: push eax
  loc_00E34838: lea eax, var_18
  loc_00E3483B: push eax
  loc_00E3483C: call edi
  loc_00E3483E: push eax
  loc_00E3483F: call [00401224h] ; __vbaCastObj
  loc_00E34845: lea ecx, var_2C
  loc_00E34848: mov var_24, eax
  loc_00E3484B: push ecx
  loc_00E3484C: mov var_2C, 0000000Dh
  loc_00E34853: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E34859: mov ecx, [eax]
  loc_00E3485B: sub esp, 00000010h
  loc_00E3485E: mov edx, esp
  loc_00E34860: push 00000000h
  loc_00E34862: push 0000002Ah
  loc_00E34864: mov [edx], ecx
  loc_00E34866: mov ecx, [eax+00000004h]
  loc_00E34869: push esi
  loc_00E3486A: mov [edx+00000004h], ecx
  loc_00E3486D: mov ecx, [eax+00000008h]
  loc_00E34870: mov eax, [eax+0000000Ch]
  loc_00E34873: mov [edx+00000008h], ecx
  loc_00E34876: mov ecx, [esi]
  loc_00E34878: mov [edx+0000000Ch], eax
  loc_00E3487B: call [ecx+0000035Ch]
  loc_00E34881: lea edx, var_1C
  loc_00E34884: push eax
  loc_00E34885: push edx
  loc_00E34886: call edi
  loc_00E34888: push eax
  loc_00E34889: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3488F: lea eax, var_1C
  loc_00E34892: lea ecx, var_18
  loc_00E34895: push eax
  loc_00E34896: push ecx
  loc_00E34897: push 00000002h
  loc_00E34899: call [00401048h] ; __vbaFreeObjList
  loc_00E3489F: add esp, 00000028h
  loc_00E348A2: lea ecx, var_2C
  loc_00E348A5: call [00401024h] ; __vbaFreeVar
  loc_00E348AB: sub esp, 00000010h
  loc_00E348AE: mov ecx, 0000000Bh
  loc_00E348B3: mov edx, esp
  loc_00E348B5: xor eax, eax
  loc_00E348B7: push 00000006h
  loc_00E348B9: push esi
  loc_00E348BA: mov [edx], ecx
  loc_00E348BC: mov ecx, var_38
  loc_00E348BF: mov [edx+00000004h], ecx
  loc_00E348C2: mov ecx, [esi]
  loc_00E348C4: mov [edx+00000008h], eax
  loc_00E348C7: mov eax, var_30
  loc_00E348CA: mov [edx+0000000Ch], eax
  loc_00E348CD: call [ecx+0000035Ch]
  loc_00E348D3: lea edx, var_18
  loc_00E348D6: push eax
  loc_00E348D7: push edx
  loc_00E348D8: call edi
  loc_00E348DA: push eax
  loc_00E348DB: call [00401238h] ; __vbaLateIdSt
  loc_00E348E1: lea ecx, var_18
  loc_00E348E4: call ebx
  loc_00E348E6: sub esp, 00000010h
  loc_00E348E9: mov ecx, 0000000Bh
  loc_00E348EE: mov edx, esp
  loc_00E348F0: xor eax, eax
  loc_00E348F2: push 8001000Eh
  loc_00E348F7: push esi
  loc_00E348F8: mov [edx], ecx
  loc_00E348FA: mov ecx, var_38
  loc_00E348FD: mov [edx+00000004h], ecx
  loc_00E34900: mov ecx, [esi]
  loc_00E34902: mov [edx+00000008h], eax
  loc_00E34905: mov eax, var_30
  loc_00E34908: mov [edx+0000000Ch], eax
  loc_00E3490B: call [ecx+0000035Ch]
  loc_00E34911: lea edx, var_18
  loc_00E34914: push eax
  loc_00E34915: push edx
  loc_00E34916: call edi
  loc_00E34918: push eax
  loc_00E34919: call [00401238h] ; __vbaLateIdSt
  loc_00E3491F: lea ecx, var_18
  loc_00E34922: call ebx
  loc_00E34924: mov eax, [esi]
  loc_00E34926: push 00000000h
  loc_00E34928: push FFFFFDDAh
  loc_00E3492D: push esi
  loc_00E3492E: call [eax+0000035Ch]
  loc_00E34934: lea ecx, var_18
  loc_00E34937: push eax
  loc_00E34938: push ecx
  loc_00E34939: call edi
  loc_00E3493B: push eax
  loc_00E3493C: call [00401030h] ; __vbaLateIdCall
  loc_00E34942: add esp, 0000000Ch
  loc_00E34945: lea ecx, var_18
  loc_00E34948: call ebx
  loc_00E3494A: mov var_4, 00000000h
  loc_00E34951: push 00E34976h
  loc_00E34956: jmp 00E34975h
  loc_00E34958: lea edx, var_1C
  loc_00E3495B: lea eax, var_18
  loc_00E3495E: push edx
  loc_00E3495F: push eax
  loc_00E34960: push 00000002h
  loc_00E34962: call [00401048h] ; __vbaFreeObjList
  loc_00E34968: add esp, 0000000Ch
  loc_00E3496B: lea ecx, var_2C
  loc_00E3496E: call [00401024h] ; __vbaFreeVar
  loc_00E34974: ret
  loc_00E34975: ret
  loc_00E34976: mov eax, Me
  loc_00E34979: push eax
  loc_00E3497A: mov ecx, [eax]
  loc_00E3497C: call [ecx+00000008h]
  loc_00E3497F: mov eax, var_4
  loc_00E34982: mov ecx, var_14
  loc_00E34985: pop edi
  loc_00E34986: pop esi
  loc_00E34987: mov fs:[00000000h], ecx
  loc_00E3498E: pop ebx
  loc_00E3498F: mov esp, ebp
  loc_00E34991: pop ebp
  loc_00E34992: retn 0004h
End Sub

Private Sub batal_UnknownEvent_9 'E32BD0
  loc_00E32BD0: push ebp
  loc_00E32BD1: mov ebp, esp
  loc_00E32BD3: sub esp, 0000000Ch
  loc_00E32BD6: push 00402836h ; __vbaExceptHandler
  loc_00E32BDB: mov eax, fs:[00000000h]
  loc_00E32BE1: push eax
  loc_00E32BE2: mov fs:[00000000h], esp
  loc_00E32BE9: sub esp, 00000098h
  loc_00E32BEF: push ebx
  loc_00E32BF0: push esi
  loc_00E32BF1: push edi
  loc_00E32BF2: mov var_C, esp
  loc_00E32BF5: mov var_8, 00402600h
  loc_00E32BFC: mov esi, Me
  loc_00E32BFF: mov eax, esi
  loc_00E32C01: and eax, 00000001h
  loc_00E32C04: mov var_4, eax
  loc_00E32C07: and esi, FFFFFFFEh
  loc_00E32C0A: push esi
  loc_00E32C0B: mov Me, esi
  loc_00E32C0E: mov ecx, [esi]
  loc_00E32C10: call [ecx+00000004h]
  loc_00E32C13: mov edx, [esi]
  loc_00E32C15: xor eax, eax
  loc_00E32C17: push esi
  loc_00E32C18: mov var_18, eax
  loc_00E32C1B: mov var_1C, eax
  loc_00E32C1E: mov var_2C, eax
  loc_00E32C21: mov var_3C, eax
  loc_00E32C24: mov var_4C, eax
  loc_00E32C27: mov var_5C, eax
  loc_00E32C2A: mov var_6C, eax
  loc_00E32C2D: mov var_7C, eax
  loc_00E32C30: call [edx+0000031Ch]
  loc_00E32C36: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E32C3C: push eax
  loc_00E32C3D: lea eax, var_18
  loc_00E32C40: push eax
  loc_00E32C41: call edi
  loc_00E32C43: mov ebx, eax
  loc_00E32C45: push FFFFFFFFh
  loc_00E32C47: push ebx
  loc_00E32C48: mov ecx, [ebx]
  loc_00E32C4A: call [ecx+0000009Ch]
  loc_00E32C50: test eax, eax
  loc_00E32C52: fnclex
  loc_00E32C54: jge 00E32C68h
  loc_00E32C56: push 0000009Ch
  loc_00E32C5B: push 006E03D4h
  loc_00E32C60: push ebx
  loc_00E32C61: push eax
  loc_00E32C62: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32C68: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E32C6E: lea ecx, var_18
  loc_00E32C71: call ebx
  loc_00E32C73: mov edx, [esi]
  loc_00E32C75: push esi
  loc_00E32C76: call [edx+0000032Ch]
  loc_00E32C7C: push eax
  loc_00E32C7D: lea eax, var_18
  loc_00E32C80: push eax
  loc_00E32C81: call edi
  loc_00E32C83: mov ecx, [eax]
  loc_00E32C85: push FFFFFFFFh
  loc_00E32C87: push eax
  loc_00E32C88: mov var_A0, eax
  loc_00E32C8E: call [ecx+0000009Ch]
  loc_00E32C94: test eax, eax
  loc_00E32C96: fnclex
  loc_00E32C98: jge 00E32CB2h
  loc_00E32C9A: mov edx, var_A0
  loc_00E32CA0: push 0000009Ch
  loc_00E32CA5: push 006E03D4h
  loc_00E32CAA: push edx
  loc_00E32CAB: push eax
  loc_00E32CAC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32CB2: lea ecx, var_18
  loc_00E32CB5: call ebx
  loc_00E32CB7: mov eax, [esi]
  loc_00E32CB9: push esi
  loc_00E32CBA: call [eax+00000330h]
  loc_00E32CC0: lea ecx, var_18
  loc_00E32CC3: push eax
  loc_00E32CC4: push ecx
  loc_00E32CC5: call edi
  loc_00E32CC7: mov edx, [eax]
  loc_00E32CC9: push FFFFFFFFh
  loc_00E32CCB: push eax
  loc_00E32CCC: mov var_A0, eax
  loc_00E32CD2: call [edx+0000009Ch]
  loc_00E32CD8: test eax, eax
  loc_00E32CDA: fnclex
  loc_00E32CDC: jge 00E32CF6h
  loc_00E32CDE: mov ecx, var_A0
  loc_00E32CE4: push 0000009Ch
  loc_00E32CE9: push 006E03D4h
  loc_00E32CEE: push ecx
  loc_00E32CEF: push eax
  loc_00E32CF0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32CF6: lea ecx, var_18
  loc_00E32CF9: call ebx
  loc_00E32CFB: mov edx, [esi]
  loc_00E32CFD: push esi
  loc_00E32CFE: call [edx+00000324h]
  loc_00E32D04: push eax
  loc_00E32D05: lea eax, var_18
  loc_00E32D08: push eax
  loc_00E32D09: call edi
  loc_00E32D0B: mov ecx, [eax]
  loc_00E32D0D: push FFFFFFFFh
  loc_00E32D0F: push eax
  loc_00E32D10: mov var_A0, eax
  loc_00E32D16: call [ecx+0000009Ch]
  loc_00E32D1C: test eax, eax
  loc_00E32D1E: fnclex
  loc_00E32D20: jge 00E32D3Ah
  loc_00E32D22: mov edx, var_A0
  loc_00E32D28: push 0000009Ch
  loc_00E32D2D: push 006E03D4h
  loc_00E32D32: push edx
  loc_00E32D33: push eax
  loc_00E32D34: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32D3A: lea ecx, var_18
  loc_00E32D3D: call ebx
  loc_00E32D3F: mov eax, [esi]
  loc_00E32D41: push esi
  loc_00E32D42: call [eax+00000314h]
  loc_00E32D48: lea ecx, var_18
  loc_00E32D4B: push eax
  loc_00E32D4C: push ecx
  loc_00E32D4D: call edi
  loc_00E32D4F: mov edx, [eax]
  loc_00E32D51: push FFFFFFFFh
  loc_00E32D53: push eax
  loc_00E32D54: mov var_A0, eax
  loc_00E32D5A: call [edx+0000009Ch]
  loc_00E32D60: test eax, eax
  loc_00E32D62: fnclex
  loc_00E32D64: jge 00E32D7Eh
  loc_00E32D66: mov ecx, var_A0
  loc_00E32D6C: push 0000009Ch
  loc_00E32D71: push 006E03D4h
  loc_00E32D76: push ecx
  loc_00E32D77: push eax
  loc_00E32D78: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32D7E: lea ecx, var_18
  loc_00E32D81: call ebx
  loc_00E32D83: mov edx, [esi]
  loc_00E32D85: push esi
  loc_00E32D86: call [edx+00000320h]
  loc_00E32D8C: push eax
  loc_00E32D8D: lea eax, var_18
  loc_00E32D90: push eax
  loc_00E32D91: call edi
  loc_00E32D93: mov ecx, [eax]
  loc_00E32D95: push FFFFFFFFh
  loc_00E32D97: push eax
  loc_00E32D98: mov var_A0, eax
  loc_00E32D9E: call [ecx+0000009Ch]
  loc_00E32DA4: test eax, eax
  loc_00E32DA6: fnclex
  loc_00E32DA8: jge 00E32DC2h
  loc_00E32DAA: mov edx, var_A0
  loc_00E32DB0: push 0000009Ch
  loc_00E32DB5: push 006E03D4h
  loc_00E32DBA: push edx
  loc_00E32DBB: push eax
  loc_00E32DBC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32DC2: lea ecx, var_18
  loc_00E32DC5: call ebx
  loc_00E32DC7: mov eax, [esi]
  loc_00E32DC9: push esi
  loc_00E32DCA: call [eax+0000031Ch]
  loc_00E32DD0: lea ecx, var_18
  loc_00E32DD3: push eax
  loc_00E32DD4: push ecx
  loc_00E32DD5: call edi
  loc_00E32DD7: mov edx, [eax]
  loc_00E32DD9: push 00000000h
  loc_00E32DDB: push eax
  loc_00E32DDC: mov var_A0, eax
  loc_00E32DE2: call [edx+000000E4h]
  loc_00E32DE8: test eax, eax
  loc_00E32DEA: fnclex
  loc_00E32DEC: jge 00E32E06h
  loc_00E32DEE: mov ecx, var_A0
  loc_00E32DF4: push 000000E4h
  loc_00E32DF9: push 006E03D4h
  loc_00E32DFE: push ecx
  loc_00E32DFF: push eax
  loc_00E32E00: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32E06: lea ecx, var_18
  loc_00E32E09: call ebx
  loc_00E32E0B: mov edx, [esi]
  loc_00E32E0D: push esi
  loc_00E32E0E: call [edx+00000328h]
  loc_00E32E14: push eax
  loc_00E32E15: lea eax, var_18
  loc_00E32E18: push eax
  loc_00E32E19: call edi
  loc_00E32E1B: mov ecx, [eax]
  loc_00E32E1D: push eax
  loc_00E32E1E: mov var_A0, eax
  loc_00E32E24: call [ecx+00000204h]
  loc_00E32E2A: test eax, eax
  loc_00E32E2C: fnclex
  loc_00E32E2E: jge 00E32E48h
  loc_00E32E30: mov edx, var_A0
  loc_00E32E36: push 00000204h
  loc_00E32E3B: push 006DCB70h
  loc_00E32E40: push edx
  loc_00E32E41: push eax
  loc_00E32E42: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E32E48: lea ecx, var_18
  loc_00E32E4B: call ebx
  loc_00E32E4D: sub esp, 00000010h
  loc_00E32E50: mov ecx, 00000008h
  loc_00E32E55: mov edx, esp
  loc_00E32E57: mov var_6C, ecx
  loc_00E32E5A: mov eax, 006E0A78h ; "biodata"
  loc_00E32E5F: push 0000000Eh
  loc_00E32E61: mov [edx], ecx
  loc_00E32E63: mov ecx, var_68
  loc_00E32E66: mov var_64, eax
  loc_00E32E69: push esi
  loc_00E32E6A: mov [edx+00000004h], ecx
  loc_00E32E6D: mov ecx, [esi]
  loc_00E32E6F: mov [edx+00000008h], eax
  loc_00E32E72: mov eax, var_60
  loc_00E32E75: mov [edx+0000000Ch], eax
  loc_00E32E78: call [ecx+00000358h]
  loc_00E32E7E: lea edx, var_18
  loc_00E32E81: push eax
  loc_00E32E82: push edx
  loc_00E32E83: call edi
  loc_00E32E85: push eax
  loc_00E32E86: call [00401238h] ; __vbaLateIdSt
  loc_00E32E8C: lea ecx, var_18
  loc_00E32E8F: call ebx
  loc_00E32E91: mov eax, [esi]
  loc_00E32E93: push 00000000h
  loc_00E32E95: push 00000019h
  loc_00E32E97: push esi
  loc_00E32E98: call [eax+00000358h]
  loc_00E32E9E: lea ecx, var_18
  loc_00E32EA1: push eax
  loc_00E32EA2: push ecx
  loc_00E32EA3: call edi
  loc_00E32EA5: push eax
  loc_00E32EA6: call [00401030h] ; __vbaLateIdCall
  loc_00E32EAC: add esp, 0000000Ch
  loc_00E32EAF: lea ecx, var_18
  loc_00E32EB2: call ebx
  loc_00E32EB4: mov edx, [esi]
  loc_00E32EB6: push 006E05D8h
  loc_00E32EBB: push esi
  loc_00E32EBC: call [edx+00000358h]
  loc_00E32EC2: push eax
  loc_00E32EC3: lea eax, var_18
  loc_00E32EC6: push eax
  loc_00E32EC7: call edi
  loc_00E32EC9: push eax
  loc_00E32ECA: call [00401224h] ; __vbaCastObj
  loc_00E32ED0: lea ecx, var_2C
  loc_00E32ED3: mov var_24, eax
  loc_00E32ED6: push ecx
  loc_00E32ED7: mov var_2C, 0000000Dh
  loc_00E32EDE: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E32EE4: mov ecx, [eax]
  loc_00E32EE6: sub esp, 00000010h
  loc_00E32EE9: mov edx, esp
  loc_00E32EEB: push 00000000h
  loc_00E32EED: push 0000002Ah
  loc_00E32EEF: mov [edx], ecx
  loc_00E32EF1: mov ecx, [eax+00000004h]
  loc_00E32EF4: push esi
  loc_00E32EF5: mov [edx+00000004h], ecx
  loc_00E32EF8: mov ecx, [eax+00000008h]
  loc_00E32EFB: mov eax, [eax+0000000Ch]
  loc_00E32EFE: mov [edx+00000008h], ecx
  loc_00E32F01: mov ecx, [esi]
  loc_00E32F03: mov [edx+0000000Ch], eax
  loc_00E32F06: call [ecx+0000035Ch]
  loc_00E32F0C: lea edx, var_1C
  loc_00E32F0F: push eax
  loc_00E32F10: push edx
  loc_00E32F11: call edi
  loc_00E32F13: push eax
  loc_00E32F14: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E32F1A: lea eax, var_1C
  loc_00E32F1D: lea ecx, var_18
  loc_00E32F20: push eax
  loc_00E32F21: push ecx
  loc_00E32F22: push 00000002h
  loc_00E32F24: call [00401048h] ; __vbaFreeObjList
  loc_00E32F2A: add esp, 00000028h
  loc_00E32F2D: lea ecx, var_2C
  loc_00E32F30: call [00401024h] ; __vbaFreeVar
  loc_00E32F36: sub esp, 00000010h
  loc_00E32F39: mov ecx, 0000000Bh
  loc_00E32F3E: mov edx, esp
  loc_00E32F40: mov var_6C, ecx
  loc_00E32F43: xor eax, eax
  loc_00E32F45: push 00000006h
  loc_00E32F47: mov [edx], ecx
  loc_00E32F49: mov ecx, var_68
  loc_00E32F4C: mov var_64, eax
  loc_00E32F4F: push esi
  loc_00E32F50: mov [edx+00000004h], ecx
  loc_00E32F53: mov ecx, [esi]
  loc_00E32F55: mov [edx+00000008h], eax
  loc_00E32F58: mov eax, var_60
  loc_00E32F5B: mov [edx+0000000Ch], eax
  loc_00E32F5E: call [ecx+0000035Ch]
  loc_00E32F64: lea edx, var_18
  loc_00E32F67: push eax
  loc_00E32F68: push edx
  loc_00E32F69: call edi
  loc_00E32F6B: push eax
  loc_00E32F6C: call [00401238h] ; __vbaLateIdSt
  loc_00E32F72: lea ecx, var_18
  loc_00E32F75: call ebx
  loc_00E32F77: sub esp, 00000010h
  loc_00E32F7A: mov ecx, 0000000Bh
  loc_00E32F7F: mov edx, esp
  loc_00E32F81: mov var_6C, ecx
  loc_00E32F84: xor eax, eax
  loc_00E32F86: push 8001000Eh
  loc_00E32F8B: mov [edx], ecx
  loc_00E32F8D: mov ecx, var_68
  loc_00E32F90: mov var_64, eax
  loc_00E32F93: push esi
  loc_00E32F94: mov [edx+00000004h], ecx
  loc_00E32F97: mov ecx, [esi]
  loc_00E32F99: mov [edx+00000008h], eax
  loc_00E32F9C: mov eax, var_60
  loc_00E32F9F: mov [edx+0000000Ch], eax
  loc_00E32FA2: call [ecx+0000035Ch]
  loc_00E32FA8: lea edx, var_18
  loc_00E32FAB: push eax
  loc_00E32FAC: push edx
  loc_00E32FAD: call edi
  loc_00E32FAF: push eax
  loc_00E32FB0: call [00401238h] ; __vbaLateIdSt
  loc_00E32FB6: lea ecx, var_18
  loc_00E32FB9: call ebx
  loc_00E32FBB: mov eax, [esi]
  loc_00E32FBD: push 00000000h
  loc_00E32FBF: push FFFFFDDAh
  loc_00E32FC4: push esi
  loc_00E32FC5: call [eax+0000035Ch]
  loc_00E32FCB: lea ecx, var_18
  loc_00E32FCE: push eax
  loc_00E32FCF: push ecx
  loc_00E32FD0: call edi
  loc_00E32FD2: push eax
  loc_00E32FD3: call [00401030h] ; __vbaLateIdCall
  loc_00E32FD9: add esp, 0000000Ch
  loc_00E32FDC: lea ecx, var_18
  loc_00E32FDF: call ebx
  loc_00E32FE1: mov edx, [esi]
  loc_00E32FE3: push esi
  loc_00E32FE4: call [edx+0000032Ch]
  loc_00E32FEA: push eax
  loc_00E32FEB: lea eax, var_18
  loc_00E32FEE: push eax
  loc_00E32FEF: call edi
  loc_00E32FF1: mov ecx, [eax]
  loc_00E32FF3: push 00000000h
  loc_00E32FF5: push eax
  loc_00E32FF6: mov var_A0, eax
  loc_00E32FFC: call [ecx+000000E4h]
  loc_00E33002: test eax, eax
  loc_00E33004: fnclex
  loc_00E33006: jge 00E33020h
  loc_00E33008: mov edx, var_A0
  loc_00E3300E: push 000000E4h
  loc_00E33013: push 006E03D4h
  loc_00E33018: push edx
  loc_00E33019: push eax
  loc_00E3301A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33020: lea ecx, var_18
  loc_00E33023: call ebx
  loc_00E33025: mov eax, [esi]
  loc_00E33027: push esi
  loc_00E33028: call [eax+00000330h]
  loc_00E3302E: lea ecx, var_18
  loc_00E33031: push eax
  loc_00E33032: push ecx
  loc_00E33033: call edi
  loc_00E33035: mov edx, [eax]
  loc_00E33037: push 00000000h
  loc_00E33039: push eax
  loc_00E3303A: mov var_A0, eax
  loc_00E33040: call [edx+000000E4h]
  loc_00E33046: test eax, eax
  loc_00E33048: fnclex
  loc_00E3304A: jge 00E33064h
  loc_00E3304C: mov ecx, var_A0
  loc_00E33052: push 000000E4h
  loc_00E33057: push 006E03D4h
  loc_00E3305C: push ecx
  loc_00E3305D: push eax
  loc_00E3305E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33064: lea ecx, var_18
  loc_00E33067: call ebx
  loc_00E33069: mov edx, [esi]
  loc_00E3306B: push esi
  loc_00E3306C: call [edx+00000324h]
  loc_00E33072: push eax
  loc_00E33073: lea eax, var_18
  loc_00E33076: push eax
  loc_00E33077: call edi
  loc_00E33079: mov ecx, [eax]
  loc_00E3307B: push 00000000h
  loc_00E3307D: push eax
  loc_00E3307E: mov var_A0, eax
  loc_00E33084: call [ecx+000000E4h]
  loc_00E3308A: test eax, eax
  loc_00E3308C: fnclex
  loc_00E3308E: jge 00E330A8h
  loc_00E33090: mov edx, var_A0
  loc_00E33096: push 000000E4h
  loc_00E3309B: push 006E03D4h
  loc_00E330A0: push edx
  loc_00E330A1: push eax
  loc_00E330A2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E330A8: lea ecx, var_18
  loc_00E330AB: call ebx
  loc_00E330AD: mov eax, [esi]
  loc_00E330AF: push esi
  loc_00E330B0: call [eax+00000314h]
  loc_00E330B6: lea ecx, var_18
  loc_00E330B9: push eax
  loc_00E330BA: push ecx
  loc_00E330BB: call edi
  loc_00E330BD: mov edx, [eax]
  loc_00E330BF: push 00000000h
  loc_00E330C1: push eax
  loc_00E330C2: mov var_A0, eax
  loc_00E330C8: call [edx+000000E4h]
  loc_00E330CE: test eax, eax
  loc_00E330D0: fnclex
  loc_00E330D2: jge 00E330ECh
  loc_00E330D4: mov ecx, var_A0
  loc_00E330DA: push 000000E4h
  loc_00E330DF: push 006E03D4h
  loc_00E330E4: push ecx
  loc_00E330E5: push eax
  loc_00E330E6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E330EC: lea ecx, var_18
  loc_00E330EF: call ebx
  loc_00E330F1: mov edx, [esi]
  loc_00E330F3: push esi
  loc_00E330F4: call [edx+00000320h]
  loc_00E330FA: push eax
  loc_00E330FB: lea eax, var_18
  loc_00E330FE: push eax
  loc_00E330FF: call edi
  loc_00E33101: mov ecx, [eax]
  loc_00E33103: push 00000000h
  loc_00E33105: push eax
  loc_00E33106: mov var_A0, eax
  loc_00E3310C: call [ecx+000000E4h]
  loc_00E33112: test eax, eax
  loc_00E33114: fnclex
  loc_00E33116: jge 00E33130h
  loc_00E33118: mov edx, var_A0
  loc_00E3311E: push 000000E4h
  loc_00E33123: push 006E03D4h
  loc_00E33128: push edx
  loc_00E33129: push eax
  loc_00E3312A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33130: lea ecx, var_18
  loc_00E33133: call ebx
  loc_00E33135: mov ecx, 80020004h
  loc_00E3313A: mov eax, 0000000Ah
  loc_00E3313F: mov var_54, ecx
  loc_00E33142: mov var_44, ecx
  loc_00E33145: lea edx, var_7C
  loc_00E33148: lea ecx, var_3C
  loc_00E3314B: mov var_5C, eax
  loc_00E3314E: mov var_4C, eax
  loc_00E33151: mov var_74, 006E16F0h ; "SMK RK BT PS"
  loc_00E33158: mov var_7C, 00000008h
  loc_00E3315F: call [004011F0h] ; __vbaVarDup
  loc_00E33165: lea edx, var_6C
  loc_00E33168: lea ecx, var_2C
  loc_00E3316B: mov var_64, 006E1CE8h ; "Silahkan Pilih Item Data yang Ingin dicari !"
  loc_00E33172: mov var_6C, 00000008h
  loc_00E33179: call [004011F0h] ; __vbaVarDup
  loc_00E3317F: lea eax, var_5C
  loc_00E33182: lea ecx, var_4C
  loc_00E33185: push eax
  loc_00E33186: lea edx, var_3C
  loc_00E33189: push ecx
  loc_00E3318A: push edx
  loc_00E3318B: lea eax, var_2C
  loc_00E3318E: push 00000040h
  loc_00E33190: push eax
  loc_00E33191: call [004010A8h] ; rtcMsgBox
  loc_00E33197: lea ecx, var_5C
  loc_00E3319A: lea edx, var_4C
  loc_00E3319D: push ecx
  loc_00E3319E: lea eax, var_3C
  loc_00E331A1: push edx
  loc_00E331A2: lea ecx, var_2C
  loc_00E331A5: push eax
  loc_00E331A6: push ecx
  loc_00E331A7: push 00000004h
  loc_00E331A9: call [00401038h] ; __vbaFreeVarList
  loc_00E331AF: add esp, 00000004h
  loc_00E331B2: mov ecx, 0000000Bh
  loc_00E331B7: mov edx, esp
  loc_00E331B9: mov var_6C, ecx
  loc_00E331BC: xor eax, eax
  loc_00E331BE: push 80010007h
  loc_00E331C3: mov [edx], ecx
  loc_00E331C5: mov ecx, var_68
  loc_00E331C8: mov var_64, eax
  loc_00E331CB: push esi
  loc_00E331CC: mov [edx+00000004h], ecx
  loc_00E331CF: mov ecx, [esi]
  loc_00E331D1: mov [edx+00000008h], eax
  loc_00E331D4: mov eax, var_60
  loc_00E331D7: mov [edx+0000000Ch], eax
  loc_00E331DA: call [ecx+00000334h]
  loc_00E331E0: lea edx, var_18
  loc_00E331E3: push eax
  loc_00E331E4: push edx
  loc_00E331E5: call edi
  loc_00E331E7: push eax
  loc_00E331E8: call [00401238h] ; __vbaLateIdSt
  loc_00E331EE: lea ecx, var_18
  loc_00E331F1: call ebx
  loc_00E331F3: mov var_4, 00000000h
  loc_00E331FA: push 00E3322Eh
  loc_00E331FF: jmp 00E3322Dh
  loc_00E33201: lea eax, var_1C
  loc_00E33204: lea ecx, var_18
  loc_00E33207: push eax
  loc_00E33208: push ecx
  loc_00E33209: push 00000002h
  loc_00E3320B: call [00401048h] ; __vbaFreeObjList
  loc_00E33211: lea edx, var_5C
  loc_00E33214: lea eax, var_4C
  loc_00E33217: push edx
  loc_00E33218: lea ecx, var_3C
  loc_00E3321B: push eax
  loc_00E3321C: lea edx, var_2C
  loc_00E3321F: push ecx
  loc_00E33220: push edx
  loc_00E33221: push 00000004h
  loc_00E33223: call [00401038h] ; __vbaFreeVarList
  loc_00E33229: add esp, 00000020h
  loc_00E3322C: ret
  loc_00E3322D: ret
  loc_00E3322E: mov eax, Me
  loc_00E33231: push eax
  loc_00E33232: mov ecx, [eax]
  loc_00E33234: call [ecx+00000008h]
  loc_00E33237: mov eax, var_4
  loc_00E3323A: mov ecx, var_14
  loc_00E3323D: pop edi
  loc_00E3323E: pop esi
  loc_00E3323F: mov fs:[00000000h], ecx
  loc_00E33246: pop ebx
  loc_00E33247: mov esp, ebp
  loc_00E33249: pop ebp
  loc_00E3324A: retn 0004h
End Sub

Private Sub ctk_UnknownEvent_9 'E33250
  loc_00E33250: push ebp
  loc_00E33251: mov ebp, esp
  loc_00E33253: sub esp, 0000000Ch
  loc_00E33256: push 00402836h ; __vbaExceptHandler
  loc_00E3325B: mov eax, fs:[00000000h]
  loc_00E33261: push eax
  loc_00E33262: mov fs:[00000000h], esp
  loc_00E33269: sub esp, 00000058h
  loc_00E3326C: push ebx
  loc_00E3326D: push esi
  loc_00E3326E: push edi
  loc_00E3326F: mov var_C, esp
  loc_00E33272: mov var_8, 00402610h
  loc_00E33279: mov esi, Me
  loc_00E3327C: mov eax, esi
  loc_00E3327E: and eax, 00000001h
  loc_00E33281: mov var_4, eax
  loc_00E33284: and esi, FFFFFFFEh
  loc_00E33287: push esi
  loc_00E33288: mov Me, esi
  loc_00E3328B: mov ecx, [esi]
  loc_00E3328D: call [ecx+00000004h]
  loc_00E33290: mov edx, [esi]
  loc_00E33292: xor eax, eax
  loc_00E33294: push eax
  loc_00E33295: push FFFFFDFAh
  loc_00E3329A: push esi
  loc_00E3329B: mov var_18, eax
  loc_00E3329E: mov var_1C, eax
  loc_00E332A1: mov var_20, eax
  loc_00E332A4: mov var_30, eax
  loc_00E332A7: mov var_40, eax
  loc_00E332AA: mov var_54, eax
  loc_00E332AD: call [edx+00000310h]
  loc_00E332B3: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E332B9: push eax
  loc_00E332BA: lea eax, var_1C
  loc_00E332BD: push eax
  loc_00E332BE: call edi
  loc_00E332C0: lea ecx, var_30
  loc_00E332C3: push eax
  loc_00E332C4: push ecx
  loc_00E332C5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E332CB: add esp, 00000010h
  loc_00E332CE: push eax
  loc_00E332CF: call [00401034h] ; __vbaStrVarMove
  loc_00E332D5: mov edx, eax
  loc_00E332D7: lea ecx, var_18
  loc_00E332DA: call [00401228h] ; __vbaStrMove
  loc_00E332E0: push eax
  loc_00E332E1: push 006E1D48h ; "C E T A K"
  loc_00E332E6: call [00401104h] ; __vbaStrCmp
  loc_00E332EC: mov ebx, eax
  loc_00E332EE: lea ecx, var_18
  loc_00E332F1: neg ebx
  loc_00E332F3: sbb ebx, ebx
  loc_00E332F5: inc ebx
  loc_00E332F6: neg ebx
  loc_00E332F8: call [00401258h] ; __vbaFreeStr
  loc_00E332FE: lea ecx, var_1C
  loc_00E33301: call [00401254h] ; __vbaFreeObj
  loc_00E33307: lea ecx, var_30
  loc_00E3330A: call [00401024h] ; __vbaFreeVar
  loc_00E33310: test bx, bx
  loc_00E33313: jz 00E336A0h
  loc_00E33319: mov edx, [esi]
  loc_00E3331B: push esi
  loc_00E3331C: call [edx+00000330h]
  loc_00E33322: push eax
  loc_00E33323: lea eax, var_1C
  loc_00E33326: push eax
  loc_00E33327: call edi
  loc_00E33329: mov ebx, eax
  loc_00E3332B: lea edx, var_54
  loc_00E3332E: push edx
  loc_00E3332F: push ebx
  loc_00E33330: mov ecx, [ebx]
  loc_00E33332: call [ecx+000000E0h]
  loc_00E33338: test eax, eax
  loc_00E3333A: fnclex
  loc_00E3333C: jge 00E33350h
  loc_00E3333E: push 000000E0h
  loc_00E33343: push 006E03D4h
  loc_00E33348: push ebx
  loc_00E33349: push eax
  loc_00E3334A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33350: mov eax, var_54
  loc_00E33353: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E33359: lea ecx, var_1C
  loc_00E3335C: mov var_60, eax
  loc_00E3335F: call ebx
  loc_00E33361: cmp var_60, 0000h
  loc_00E33366: jz 00E333D5h
  loc_00E33368: sub esp, 00000010h
  loc_00E3336B: mov ecx, 00000008h
  loc_00E33370: mov edx, esp
  loc_00E33372: mov eax, 006E1D60h ; "Close Menu"
  loc_00E33377: push FFFFFDFAh
  loc_00E3337C: push esi
  loc_00E3337D: mov [edx], ecx
  loc_00E3337F: mov ecx, var_3C
  loc_00E33382: mov [edx+00000004h], ecx
  loc_00E33385: mov ecx, [esi]
  loc_00E33387: mov [edx+00000008h], eax
  loc_00E3338A: mov eax, var_34
  loc_00E3338D: mov [edx+0000000Ch], eax
  loc_00E33390: call [ecx+00000310h]
  loc_00E33396: lea edx, var_1C
  loc_00E33399: push eax
  loc_00E3339A: push edx
  loc_00E3339B: call edi
  loc_00E3339D: push eax
  loc_00E3339E: call [00401238h] ; __vbaLateIdSt
  loc_00E333A4: lea ecx, var_1C
  loc_00E333A7: call ebx
  loc_00E333A9: mov eax, [esi]
  loc_00E333AB: push esi
  loc_00E333AC: call [eax+000002FCh]
  loc_00E333B2: lea ecx, var_1C
  loc_00E333B5: push eax
  loc_00E333B6: push ecx
  loc_00E333B7: call edi
  loc_00E333B9: mov esi, eax
  loc_00E333BB: push FFFFFFFFh
  loc_00E333BD: push esi
  loc_00E333BE: mov edx, [esi]
  loc_00E333C0: call [edx+0000009Ch]
  loc_00E333C6: test eax, eax
  loc_00E333C8: fnclex
  loc_00E333CA: jge 00E3378Ch
  loc_00E333D0: jmp 00E3377Ah
  loc_00E333D5: mov eax, [esi]
  loc_00E333D7: push esi
  loc_00E333D8: call [eax+0000032Ch]
  loc_00E333DE: lea ecx, var_1C
  loc_00E333E1: push eax
  loc_00E333E2: push ecx
  loc_00E333E3: call edi
  loc_00E333E5: mov edx, [eax]
  loc_00E333E7: lea ecx, var_54
  loc_00E333EA: push ecx
  loc_00E333EB: push eax
  loc_00E333EC: mov var_58, eax
  loc_00E333EF: call [edx+000000E0h]
  loc_00E333F5: test eax, eax
  loc_00E333F7: fnclex
  loc_00E333F9: jge 00E33410h
  loc_00E333FB: mov edx, var_58
  loc_00E333FE: push 000000E0h
  loc_00E33403: push 006E03D4h
  loc_00E33408: push edx
  loc_00E33409: push eax
  loc_00E3340A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33410: mov eax, var_54
  loc_00E33413: lea ecx, var_1C
  loc_00E33416: mov var_60, eax
  loc_00E33419: call ebx
  loc_00E3341B: cmp var_60, 0000h
  loc_00E33420: jz 00E3348Fh
  loc_00E33422: sub esp, 00000010h
  loc_00E33425: mov ecx, 00000008h
  loc_00E3342A: mov edx, esp
  loc_00E3342C: mov eax, 006E1D60h ; "Close Menu"
  loc_00E33431: push FFFFFDFAh
  loc_00E33436: push esi
  loc_00E33437: mov [edx], ecx
  loc_00E33439: mov ecx, var_3C
  loc_00E3343C: mov [edx+00000004h], ecx
  loc_00E3343F: mov ecx, [esi]
  loc_00E33441: mov [edx+00000008h], eax
  loc_00E33444: mov eax, var_34
  loc_00E33447: mov [edx+0000000Ch], eax
  loc_00E3344A: call [ecx+00000310h]
  loc_00E33450: lea edx, var_1C
  loc_00E33453: push eax
  loc_00E33454: push edx
  loc_00E33455: call edi
  loc_00E33457: push eax
  loc_00E33458: call [00401238h] ; __vbaLateIdSt
  loc_00E3345E: lea ecx, var_1C
  loc_00E33461: call ebx
  loc_00E33463: mov eax, [esi]
  loc_00E33465: push esi
  loc_00E33466: call [eax+000002FCh]
  loc_00E3346C: lea ecx, var_1C
  loc_00E3346F: push eax
  loc_00E33470: push ecx
  loc_00E33471: call edi
  loc_00E33473: mov esi, eax
  loc_00E33475: push FFFFFFFFh
  loc_00E33477: push esi
  loc_00E33478: mov edx, [esi]
  loc_00E3347A: call [edx+0000009Ch]
  loc_00E33480: test eax, eax
  loc_00E33482: fnclex
  loc_00E33484: jge 00E3378Ch
  loc_00E3348A: jmp 00E3377Ah
  loc_00E3348F: mov eax, [esi]
  loc_00E33491: push esi
  loc_00E33492: call [eax+00000324h]
  loc_00E33498: lea ecx, var_1C
  loc_00E3349B: push eax
  loc_00E3349C: push ecx
  loc_00E3349D: call edi
  loc_00E3349F: mov edx, [eax]
  loc_00E334A1: lea ecx, var_54
  loc_00E334A4: push ecx
  loc_00E334A5: push eax
  loc_00E334A6: mov var_58, eax
  loc_00E334A9: call [edx+000000E0h]
  loc_00E334AF: test eax, eax
  loc_00E334B1: fnclex
  loc_00E334B3: jge 00E334CAh
  loc_00E334B5: mov edx, var_58
  loc_00E334B8: push 000000E0h
  loc_00E334BD: push 006E03D4h
  loc_00E334C2: push edx
  loc_00E334C3: push eax
  loc_00E334C4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E334CA: mov eax, var_54
  loc_00E334CD: lea ecx, var_1C
  loc_00E334D0: mov var_60, eax
  loc_00E334D3: call ebx
  loc_00E334D5: cmp var_60, 0000h
  loc_00E334DA: jz 00E33549h
  loc_00E334DC: sub esp, 00000010h
  loc_00E334DF: mov ecx, 00000008h
  loc_00E334E4: mov edx, esp
  loc_00E334E6: mov eax, 006E1D60h ; "Close Menu"
  loc_00E334EB: push FFFFFDFAh
  loc_00E334F0: push esi
  loc_00E334F1: mov [edx], ecx
  loc_00E334F3: mov ecx, var_3C
  loc_00E334F6: mov [edx+00000004h], ecx
  loc_00E334F9: mov ecx, [esi]
  loc_00E334FB: mov [edx+00000008h], eax
  loc_00E334FE: mov eax, var_34
  loc_00E33501: mov [edx+0000000Ch], eax
  loc_00E33504: call [ecx+00000310h]
  loc_00E3350A: lea edx, var_1C
  loc_00E3350D: push eax
  loc_00E3350E: push edx
  loc_00E3350F: call edi
  loc_00E33511: push eax
  loc_00E33512: call [00401238h] ; __vbaLateIdSt
  loc_00E33518: lea ecx, var_1C
  loc_00E3351B: call ebx
  loc_00E3351D: mov eax, [esi]
  loc_00E3351F: push esi
  loc_00E33520: call [eax+000002FCh]
  loc_00E33526: lea ecx, var_1C
  loc_00E33529: push eax
  loc_00E3352A: push ecx
  loc_00E3352B: call edi
  loc_00E3352D: mov esi, eax
  loc_00E3352F: push FFFFFFFFh
  loc_00E33531: push esi
  loc_00E33532: mov edx, [esi]
  loc_00E33534: call [edx+0000009Ch]
  loc_00E3353A: test eax, eax
  loc_00E3353C: fnclex
  loc_00E3353E: jge 00E3378Ch
  loc_00E33544: jmp 00E3377Ah
  loc_00E33549: mov eax, [00E3D128h]
  loc_00E3354E: test eax, eax
  loc_00E33550: jnz 00E33562h
  loc_00E33552: push 00E3D128h
  loc_00E33557: push 006CB548h
  loc_00E3355C: call [004011A0h] ; __vbaNew2
  loc_00E33562: mov ebx, [00E3D128h]
  loc_00E33568: mov eax, [esi]
  loc_00E3356A: push 006E05D8h
  loc_00E3356F: push esi
  loc_00E33570: mov edx, [ebx]
  loc_00E33572: mov var_6C, edx
  loc_00E33575: call [eax+00000358h]
  loc_00E3357B: lea ecx, var_1C
  loc_00E3357E: push eax
  loc_00E3357F: push ecx
  loc_00E33580: call edi
  loc_00E33582: push eax
  loc_00E33583: call [00401224h] ; __vbaCastObj
  loc_00E33589: lea edx, var_20
  loc_00E3358C: push eax
  loc_00E3358D: push edx
  loc_00E3358E: call edi
  loc_00E33590: push eax
  loc_00E33591: mov eax, var_6C
  loc_00E33594: push ebx
  loc_00E33595: call [eax+00000028h]
  loc_00E33598: test eax, eax
  loc_00E3359A: fnclex
  loc_00E3359C: jge 00E335ADh
  loc_00E3359E: push 00000028h
  loc_00E335A0: push 006DD034h
  loc_00E335A5: push ebx
  loc_00E335A6: push eax
  loc_00E335A7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E335AD: lea ecx, var_20
  loc_00E335B0: lea edx, var_1C
  loc_00E335B3: push ecx
  loc_00E335B4: push edx
  loc_00E335B5: push 00000002h
  loc_00E335B7: call [00401048h] ; __vbaFreeObjList
  loc_00E335BD: mov eax, [00E3D128h]
  loc_00E335C2: add esp, 0000000Ch
  loc_00E335C5: test eax, eax
  loc_00E335C7: jnz 00E335DDh
  loc_00E335C9: mov edi, [004011A0h] ; __vbaNew2
  loc_00E335CF: push 00E3D128h
  loc_00E335D4: push 006CB548h
  loc_00E335D9: call edi
  loc_00E335DB: jmp 00E335E3h
  loc_00E335DD: mov edi, [004011A0h] ; __vbaNew2
  loc_00E335E3: mov esi, [00E3D128h]
  loc_00E335E9: push esi
  loc_00E335EA: mov eax, [esi]
  loc_00E335EC: call [eax+00000088h]
  loc_00E335F2: test eax, eax
  loc_00E335F4: fnclex
  loc_00E335F6: jge 00E3360Ah
  loc_00E335F8: push 00000088h
  loc_00E335FD: push 006DD034h
  loc_00E33602: push esi
  loc_00E33603: push eax
  loc_00E33604: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3360A: mov eax, [00E3D9CCh]
  loc_00E3360F: test eax, eax
  loc_00E33611: jnz 00E3361Fh
  loc_00E33613: push 00E3D9CCh
  loc_00E33618: push 006DC960h
  loc_00E3361D: call edi
  loc_00E3361F: mov eax, [00E3D128h]
  loc_00E33624: mov esi, [00E3D9CCh]
  loc_00E3362A: test eax, eax
  loc_00E3362C: jnz 00E3363Ah
  loc_00E3362E: push 00E3D128h
  loc_00E33633: push 006CB548h
  loc_00E33638: call edi
  loc_00E3363A: mov ecx, [00E3D128h]
  loc_00E33640: mov ebx, [esi]
  loc_00E33642: lea edx, var_1C
  loc_00E33645: push ecx
  loc_00E33646: push edx
  loc_00E33647: call [004010B4h] ; __vbaObjSetAddref
  loc_00E3364D: push eax
  loc_00E3364E: push esi
  loc_00E3364F: call [ebx+0000000Ch]
  loc_00E33652: test eax, eax
  loc_00E33654: fnclex
  loc_00E33656: jge 00E33667h
  loc_00E33658: push 0000000Ch
  loc_00E3365A: push 006DC950h
  loc_00E3365F: push esi
  loc_00E33660: push eax
  loc_00E33661: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33667: lea ecx, var_1C
  loc_00E3366A: call [00401254h] ; __vbaFreeObj
  loc_00E33670: mov eax, [00E3D128h]
  loc_00E33675: test eax, eax
  loc_00E33677: jnz 00E33685h
  loc_00E33679: push 00E3D128h
  loc_00E3367E: push 006CB548h
  loc_00E33683: call edi
  loc_00E33685: mov eax, [00E3D128h]
  loc_00E3368A: push 00000000h
  loc_00E3368C: push 80011003h
  loc_00E33691: push eax
  loc_00E33692: call [00401030h] ; __vbaLateIdCall
  loc_00E33698: add esp, 0000000Ch
  loc_00E3369B: jmp 00E33791h
  loc_00E336A0: mov ecx, [esi]
  loc_00E336A2: push 00000000h
  loc_00E336A4: push FFFFFDFAh
  loc_00E336A9: push esi
  loc_00E336AA: call [ecx+00000310h]
  loc_00E336B0: lea edx, var_1C
  loc_00E336B3: push eax
  loc_00E336B4: push edx
  loc_00E336B5: call edi
  loc_00E336B7: push eax
  loc_00E336B8: lea eax, var_30
  loc_00E336BB: push eax
  loc_00E336BC: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E336C2: add esp, 00000010h
  loc_00E336C5: push eax
  loc_00E336C6: call [00401034h] ; __vbaStrVarMove
  loc_00E336CC: mov edx, eax
  loc_00E336CE: lea ecx, var_18
  loc_00E336D1: call [00401228h] ; __vbaStrMove
  loc_00E336D7: push eax
  loc_00E336D8: push 006E1D60h ; "Close Menu"
  loc_00E336DD: call [00401104h] ; __vbaStrCmp
  loc_00E336E3: mov ebx, eax
  loc_00E336E5: lea ecx, var_18
  loc_00E336E8: neg ebx
  loc_00E336EA: sbb ebx, ebx
  loc_00E336EC: inc ebx
  loc_00E336ED: neg ebx
  loc_00E336EF: call [00401258h] ; __vbaFreeStr
  loc_00E336F5: lea ecx, var_1C
  loc_00E336F8: call [00401254h] ; __vbaFreeObj
  loc_00E336FE: lea ecx, var_30
  loc_00E33701: call [00401024h] ; __vbaFreeVar
  loc_00E33707: test bx, bx
  loc_00E3370A: jz 00E33791h
  loc_00E33710: sub esp, 00000010h
  loc_00E33713: mov ecx, 00000008h
  loc_00E33718: mov edx, esp
  loc_00E3371A: mov eax, 006E1D48h ; "C E T A K"
  loc_00E3371F: push FFFFFDFAh
  loc_00E33724: push esi
  loc_00E33725: mov [edx], ecx
  loc_00E33727: mov ecx, var_3C
  loc_00E3372A: mov [edx+00000004h], ecx
  loc_00E3372D: mov ecx, [esi]
  loc_00E3372F: mov [edx+00000008h], eax
  loc_00E33732: mov eax, var_34
  loc_00E33735: mov [edx+0000000Ch], eax
  loc_00E33738: call [ecx+00000310h]
  loc_00E3373E: lea edx, var_1C
  loc_00E33741: push eax
  loc_00E33742: push edx
  loc_00E33743: call edi
  loc_00E33745: push eax
  loc_00E33746: call [00401238h] ; __vbaLateIdSt
  loc_00E3374C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E33752: lea ecx, var_1C
  loc_00E33755: call ebx
  loc_00E33757: mov eax, [esi]
  loc_00E33759: push esi
  loc_00E3375A: call [eax+000002FCh]
  loc_00E33760: lea ecx, var_1C
  loc_00E33763: push eax
  loc_00E33764: push ecx
  loc_00E33765: call edi
  loc_00E33767: mov esi, eax
  loc_00E33769: push 00000000h
  loc_00E3376B: push esi
  loc_00E3376C: mov edx, [esi]
  loc_00E3376E: call [edx+0000009Ch]
  loc_00E33774: test eax, eax
  loc_00E33776: fnclex
  loc_00E33778: jge 00E3378Ch
  loc_00E3377A: push 0000009Ch
  loc_00E3377F: push 006DCAD0h
  loc_00E33784: push esi
  loc_00E33785: push eax
  loc_00E33786: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3378C: lea ecx, var_1C
  loc_00E3378F: call ebx
  loc_00E33791: mov var_4, 00000000h
  loc_00E33798: push 00E337C6h
  loc_00E3379D: jmp 00E337C5h
  loc_00E3379F: lea ecx, var_18
  loc_00E337A2: call [00401258h] ; __vbaFreeStr
  loc_00E337A8: lea eax, var_20
  loc_00E337AB: lea ecx, var_1C
  loc_00E337AE: push eax
  loc_00E337AF: push ecx
  loc_00E337B0: push 00000002h
  loc_00E337B2: call [00401048h] ; __vbaFreeObjList
  loc_00E337B8: add esp, 0000000Ch
  loc_00E337BB: lea ecx, var_30
  loc_00E337BE: call [00401024h] ; __vbaFreeVar
  loc_00E337C4: ret
  loc_00E337C5: ret
  loc_00E337C6: mov eax, Me
  loc_00E337C9: push eax
  loc_00E337CA: mov edx, [eax]
  loc_00E337CC: call [edx+00000008h]
  loc_00E337CF: mov eax, var_4
  loc_00E337D2: mov ecx, var_14
  loc_00E337D5: pop edi
  loc_00E337D6: pop esi
  loc_00E337D7: mov fs:[00000000h], ecx
  loc_00E337DE: pop ebx
  loc_00E337DF: mov esp, ebp
  loc_00E337E1: pop ebp
  loc_00E337E2: retn 0004h
End Sub

Private Sub optnama_Click() 'E34D80
  loc_00E34D80: push ebp
  loc_00E34D81: mov ebp, esp
  loc_00E34D83: sub esp, 0000000Ch
  loc_00E34D86: push 00402836h ; __vbaExceptHandler
  loc_00E34D8B: mov eax, fs:[00000000h]
  loc_00E34D91: push eax
  loc_00E34D92: mov fs:[00000000h], esp
  loc_00E34D99: sub esp, 00000048h
  loc_00E34D9C: push ebx
  loc_00E34D9D: push esi
  loc_00E34D9E: push edi
  loc_00E34D9F: mov var_C, esp
  loc_00E34DA2: mov var_8, 00402690h
  loc_00E34DA9: mov esi, Me
  loc_00E34DAC: mov eax, esi
  loc_00E34DAE: and eax, 00000001h
  loc_00E34DB1: mov var_4, eax
  loc_00E34DB4: and esi, FFFFFFFEh
  loc_00E34DB7: push esi
  loc_00E34DB8: mov Me, esi
  loc_00E34DBB: mov ecx, [esi]
  loc_00E34DBD: call [ecx+00000004h]
  loc_00E34DC0: sub esp, 00000010h
  loc_00E34DC3: mov ecx, 0000000Bh
  loc_00E34DC8: mov edx, esp
  loc_00E34DCA: xor eax, eax
  loc_00E34DCC: mov var_18, eax
  loc_00E34DCF: mov var_1C, eax
  loc_00E34DD2: mov [edx], ecx
  loc_00E34DD4: mov ecx, var_38
  loc_00E34DD7: mov var_2C, eax
  loc_00E34DDA: or eax, FFFFFFFFh
  loc_00E34DDD: mov [edx+00000004h], ecx
  loc_00E34DE0: mov ecx, [esi]
  loc_00E34DE2: push 80010007h
  loc_00E34DE7: push esi
  loc_00E34DE8: mov [edx+00000008h], eax
  loc_00E34DEB: mov eax, var_30
  loc_00E34DEE: mov [edx+0000000Ch], eax
  loc_00E34DF1: call [ecx+00000334h]
  loc_00E34DF7: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E34DFD: lea edx, var_18
  loc_00E34E00: push eax
  loc_00E34E01: push edx
  loc_00E34E02: call edi
  loc_00E34E04: push eax
  loc_00E34E05: call [00401238h] ; __vbaLateIdSt
  loc_00E34E0B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E34E11: lea ecx, var_18
  loc_00E34E14: call ebx
  loc_00E34E16: mov eax, [esi]
  loc_00E34E18: push esi
  loc_00E34E19: call [eax+00000328h]
  loc_00E34E1F: lea ecx, var_18
  loc_00E34E22: push eax
  loc_00E34E23: push ecx
  loc_00E34E24: call edi
  loc_00E34E26: mov edx, [eax]
  loc_00E34E28: push eax
  loc_00E34E29: mov var_50, eax
  loc_00E34E2C: call [edx+00000204h]
  loc_00E34E32: test eax, eax
  loc_00E34E34: fnclex
  loc_00E34E36: jge 00E34E4Dh
  loc_00E34E38: mov ecx, var_50
  loc_00E34E3B: push 00000204h
  loc_00E34E40: push 006DCB70h
  loc_00E34E45: push ecx
  loc_00E34E46: push eax
  loc_00E34E47: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34E4D: lea ecx, var_18
  loc_00E34E50: call ebx
  loc_00E34E52: mov edx, [esi]
  loc_00E34E54: push esi
  loc_00E34E55: call [edx+0000032Ch]
  loc_00E34E5B: push eax
  loc_00E34E5C: lea eax, var_18
  loc_00E34E5F: push eax
  loc_00E34E60: call edi
  loc_00E34E62: mov ecx, [eax]
  loc_00E34E64: push 00000000h
  loc_00E34E66: push eax
  loc_00E34E67: mov var_50, eax
  loc_00E34E6A: call [ecx+0000009Ch]
  loc_00E34E70: test eax, eax
  loc_00E34E72: fnclex
  loc_00E34E74: jge 00E34E8Bh
  loc_00E34E76: mov edx, var_50
  loc_00E34E79: push 0000009Ch
  loc_00E34E7E: push 006E03D4h
  loc_00E34E83: push edx
  loc_00E34E84: push eax
  loc_00E34E85: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34E8B: lea ecx, var_18
  loc_00E34E8E: call ebx
  loc_00E34E90: mov eax, [esi]
  loc_00E34E92: push esi
  loc_00E34E93: call [eax+0000031Ch]
  loc_00E34E99: lea ecx, var_18
  loc_00E34E9C: push eax
  loc_00E34E9D: push ecx
  loc_00E34E9E: call edi
  loc_00E34EA0: mov edx, [eax]
  loc_00E34EA2: push 00000000h
  loc_00E34EA4: push eax
  loc_00E34EA5: mov var_50, eax
  loc_00E34EA8: call [edx+0000009Ch]
  loc_00E34EAE: test eax, eax
  loc_00E34EB0: fnclex
  loc_00E34EB2: jge 00E34EC9h
  loc_00E34EB4: mov ecx, var_50
  loc_00E34EB7: push 0000009Ch
  loc_00E34EBC: push 006E03D4h
  loc_00E34EC1: push ecx
  loc_00E34EC2: push eax
  loc_00E34EC3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34EC9: lea ecx, var_18
  loc_00E34ECC: call ebx
  loc_00E34ECE: mov edx, [esi]
  loc_00E34ED0: push esi
  loc_00E34ED1: call [edx+00000324h]
  loc_00E34ED7: push eax
  loc_00E34ED8: lea eax, var_18
  loc_00E34EDB: push eax
  loc_00E34EDC: call edi
  loc_00E34EDE: mov ecx, [eax]
  loc_00E34EE0: push 00000000h
  loc_00E34EE2: push eax
  loc_00E34EE3: mov var_50, eax
  loc_00E34EE6: call [ecx+0000009Ch]
  loc_00E34EEC: test eax, eax
  loc_00E34EEE: fnclex
  loc_00E34EF0: jge 00E34F07h
  loc_00E34EF2: mov edx, var_50
  loc_00E34EF5: push 0000009Ch
  loc_00E34EFA: push 006E03D4h
  loc_00E34EFF: push edx
  loc_00E34F00: push eax
  loc_00E34F01: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34F07: lea ecx, var_18
  loc_00E34F0A: call ebx
  loc_00E34F0C: mov eax, [esi]
  loc_00E34F0E: push esi
  loc_00E34F0F: call [eax+00000314h]
  loc_00E34F15: lea ecx, var_18
  loc_00E34F18: push eax
  loc_00E34F19: push ecx
  loc_00E34F1A: call edi
  loc_00E34F1C: mov edx, [eax]
  loc_00E34F1E: push 00000000h
  loc_00E34F20: push eax
  loc_00E34F21: mov var_50, eax
  loc_00E34F24: call [edx+0000009Ch]
  loc_00E34F2A: test eax, eax
  loc_00E34F2C: fnclex
  loc_00E34F2E: jge 00E34F45h
  loc_00E34F30: mov ecx, var_50
  loc_00E34F33: push 0000009Ch
  loc_00E34F38: push 006E03D4h
  loc_00E34F3D: push ecx
  loc_00E34F3E: push eax
  loc_00E34F3F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34F45: lea ecx, var_18
  loc_00E34F48: call ebx
  loc_00E34F4A: mov edx, [esi]
  loc_00E34F4C: push esi
  loc_00E34F4D: call [edx+00000320h]
  loc_00E34F53: push eax
  loc_00E34F54: lea eax, var_18
  loc_00E34F57: push eax
  loc_00E34F58: call edi
  loc_00E34F5A: mov ecx, [eax]
  loc_00E34F5C: push 00000000h
  loc_00E34F5E: push eax
  loc_00E34F5F: mov var_50, eax
  loc_00E34F62: call [ecx+0000009Ch]
  loc_00E34F68: test eax, eax
  loc_00E34F6A: fnclex
  loc_00E34F6C: jge 00E34F83h
  loc_00E34F6E: mov edx, var_50
  loc_00E34F71: push 0000009Ch
  loc_00E34F76: push 006E03D4h
  loc_00E34F7B: push edx
  loc_00E34F7C: push eax
  loc_00E34F7D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34F83: lea ecx, var_18
  loc_00E34F86: call ebx
  loc_00E34F88: mov eax, [esi]
  loc_00E34F8A: push esi
  loc_00E34F8B: call [eax+00000328h]
  loc_00E34F91: lea ecx, var_18
  loc_00E34F94: push eax
  loc_00E34F95: push ecx
  loc_00E34F96: call edi
  loc_00E34F98: mov edx, [eax]
  loc_00E34F9A: push 006DCC80h
  loc_00E34F9F: push eax
  loc_00E34FA0: mov var_50, eax
  loc_00E34FA3: call [edx+000000A4h]
  loc_00E34FA9: test eax, eax
  loc_00E34FAB: fnclex
  loc_00E34FAD: jge 00E34FC4h
  loc_00E34FAF: mov ecx, var_50
  loc_00E34FB2: push 000000A4h
  loc_00E34FB7: push 006DCB70h
  loc_00E34FBC: push ecx
  loc_00E34FBD: push eax
  loc_00E34FBE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E34FC4: lea ecx, var_18
  loc_00E34FC7: call ebx
  loc_00E34FC9: sub esp, 00000010h
  loc_00E34FCC: mov ecx, 00000008h
  loc_00E34FD1: mov edx, esp
  loc_00E34FD3: mov eax, 006E0A78h ; "biodata"
  loc_00E34FD8: push 0000000Eh
  loc_00E34FDA: push esi
  loc_00E34FDB: mov [edx], ecx
  loc_00E34FDD: mov ecx, var_38
  loc_00E34FE0: mov [edx+00000004h], ecx
  loc_00E34FE3: mov ecx, [esi]
  loc_00E34FE5: mov [edx+00000008h], eax
  loc_00E34FE8: mov eax, var_30
  loc_00E34FEB: mov [edx+0000000Ch], eax
  loc_00E34FEE: call [ecx+00000358h]
  loc_00E34FF4: lea edx, var_18
  loc_00E34FF7: push eax
  loc_00E34FF8: push edx
  loc_00E34FF9: call edi
  loc_00E34FFB: push eax
  loc_00E34FFC: call [00401238h] ; __vbaLateIdSt
  loc_00E35002: lea ecx, var_18
  loc_00E35005: call ebx
  loc_00E35007: mov eax, [esi]
  loc_00E35009: push 00000000h
  loc_00E3500B: push 00000019h
  loc_00E3500D: push esi
  loc_00E3500E: call [eax+00000358h]
  loc_00E35014: lea ecx, var_18
  loc_00E35017: push eax
  loc_00E35018: push ecx
  loc_00E35019: call edi
  loc_00E3501B: push eax
  loc_00E3501C: call [00401030h] ; __vbaLateIdCall
  loc_00E35022: add esp, 0000000Ch
  loc_00E35025: lea ecx, var_18
  loc_00E35028: call ebx
  loc_00E3502A: mov edx, [esi]
  loc_00E3502C: push 006E05D8h
  loc_00E35031: push esi
  loc_00E35032: call [edx+00000358h]
  loc_00E35038: push eax
  loc_00E35039: lea eax, var_18
  loc_00E3503C: push eax
  loc_00E3503D: call edi
  loc_00E3503F: push eax
  loc_00E35040: call [00401224h] ; __vbaCastObj
  loc_00E35046: lea ecx, var_2C
  loc_00E35049: mov var_24, eax
  loc_00E3504C: push ecx
  loc_00E3504D: mov var_2C, 0000000Dh
  loc_00E35054: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3505A: mov ecx, [eax]
  loc_00E3505C: sub esp, 00000010h
  loc_00E3505F: mov edx, esp
  loc_00E35061: push 00000000h
  loc_00E35063: push 0000002Ah
  loc_00E35065: mov [edx], ecx
  loc_00E35067: mov ecx, [eax+00000004h]
  loc_00E3506A: push esi
  loc_00E3506B: mov [edx+00000004h], ecx
  loc_00E3506E: mov ecx, [eax+00000008h]
  loc_00E35071: mov eax, [eax+0000000Ch]
  loc_00E35074: mov [edx+00000008h], ecx
  loc_00E35077: mov ecx, [esi]
  loc_00E35079: mov [edx+0000000Ch], eax
  loc_00E3507C: call [ecx+0000035Ch]
  loc_00E35082: lea edx, var_1C
  loc_00E35085: push eax
  loc_00E35086: push edx
  loc_00E35087: call edi
  loc_00E35089: push eax
  loc_00E3508A: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E35090: lea eax, var_1C
  loc_00E35093: lea ecx, var_18
  loc_00E35096: push eax
  loc_00E35097: push ecx
  loc_00E35098: push 00000002h
  loc_00E3509A: call [00401048h] ; __vbaFreeObjList
  loc_00E350A0: add esp, 00000028h
  loc_00E350A3: lea ecx, var_2C
  loc_00E350A6: call [00401024h] ; __vbaFreeVar
  loc_00E350AC: sub esp, 00000010h
  loc_00E350AF: mov ecx, 0000000Bh
  loc_00E350B4: mov edx, esp
  loc_00E350B6: xor eax, eax
  loc_00E350B8: push 00000006h
  loc_00E350BA: push esi
  loc_00E350BB: mov [edx], ecx
  loc_00E350BD: mov ecx, var_38
  loc_00E350C0: mov [edx+00000004h], ecx
  loc_00E350C3: mov ecx, [esi]
  loc_00E350C5: mov [edx+00000008h], eax
  loc_00E350C8: mov eax, var_30
  loc_00E350CB: mov [edx+0000000Ch], eax
  loc_00E350CE: call [ecx+0000035Ch]
  loc_00E350D4: lea edx, var_18
  loc_00E350D7: push eax
  loc_00E350D8: push edx
  loc_00E350D9: call edi
  loc_00E350DB: push eax
  loc_00E350DC: call [00401238h] ; __vbaLateIdSt
  loc_00E350E2: lea ecx, var_18
  loc_00E350E5: call ebx
  loc_00E350E7: sub esp, 00000010h
  loc_00E350EA: mov ecx, 0000000Bh
  loc_00E350EF: mov edx, esp
  loc_00E350F1: xor eax, eax
  loc_00E350F3: push 8001000Eh
  loc_00E350F8: push esi
  loc_00E350F9: mov [edx], ecx
  loc_00E350FB: mov ecx, var_38
  loc_00E350FE: mov [edx+00000004h], ecx
  loc_00E35101: mov ecx, [esi]
  loc_00E35103: mov [edx+00000008h], eax
  loc_00E35106: mov eax, var_30
  loc_00E35109: mov [edx+0000000Ch], eax
  loc_00E3510C: call [ecx+0000035Ch]
  loc_00E35112: lea edx, var_18
  loc_00E35115: push eax
  loc_00E35116: push edx
  loc_00E35117: call edi
  loc_00E35119: push eax
  loc_00E3511A: call [00401238h] ; __vbaLateIdSt
  loc_00E35120: lea ecx, var_18
  loc_00E35123: call ebx
  loc_00E35125: mov eax, [esi]
  loc_00E35127: push 00000000h
  loc_00E35129: push FFFFFDDAh
  loc_00E3512E: push esi
  loc_00E3512F: call [eax+0000035Ch]
  loc_00E35135: lea ecx, var_18
  loc_00E35138: push eax
  loc_00E35139: push ecx
  loc_00E3513A: call edi
  loc_00E3513C: push eax
  loc_00E3513D: call [00401030h] ; __vbaLateIdCall
  loc_00E35143: add esp, 0000000Ch
  loc_00E35146: lea ecx, var_18
  loc_00E35149: call ebx
  loc_00E3514B: mov var_4, 00000000h
  loc_00E35152: push 00E35177h
  loc_00E35157: jmp 00E35176h
  loc_00E35159: lea edx, var_1C
  loc_00E3515C: lea eax, var_18
  loc_00E3515F: push edx
  loc_00E35160: push eax
  loc_00E35161: push 00000002h
  loc_00E35163: call [00401048h] ; __vbaFreeObjList
  loc_00E35169: add esp, 0000000Ch
  loc_00E3516C: lea ecx, var_2C
  loc_00E3516F: call [00401024h] ; __vbaFreeVar
  loc_00E35175: ret
  loc_00E35176: ret
  loc_00E35177: mov eax, Me
  loc_00E3517A: push eax
  loc_00E3517B: mov ecx, [eax]
  loc_00E3517D: call [ecx+00000008h]
  loc_00E35180: mov eax, var_4
  loc_00E35183: mov ecx, var_14
  loc_00E35186: pop edi
  loc_00E35187: pop esi
  loc_00E35188: mov fs:[00000000h], ecx
  loc_00E3518F: pop ebx
  loc_00E35190: mov esp, ebp
  loc_00E35192: pop ebp
  loc_00E35193: retn 0004h
End Sub

Private Sub optno_Click() 'E35580
  loc_00E35580: push ebp
  loc_00E35581: mov ebp, esp
  loc_00E35583: sub esp, 0000000Ch
  loc_00E35586: push 00402836h ; __vbaExceptHandler
  loc_00E3558B: mov eax, fs:[00000000h]
  loc_00E35591: push eax
  loc_00E35592: mov fs:[00000000h], esp
  loc_00E35599: sub esp, 00000048h
  loc_00E3559C: push ebx
  loc_00E3559D: push esi
  loc_00E3559E: push edi
  loc_00E3559F: mov var_C, esp
  loc_00E355A2: mov var_8, 004026B0h
  loc_00E355A9: mov esi, Me
  loc_00E355AC: mov eax, esi
  loc_00E355AE: and eax, 00000001h
  loc_00E355B1: mov var_4, eax
  loc_00E355B4: and esi, FFFFFFFEh
  loc_00E355B7: push esi
  loc_00E355B8: mov Me, esi
  loc_00E355BB: mov ecx, [esi]
  loc_00E355BD: call [ecx+00000004h]
  loc_00E355C0: sub esp, 00000010h
  loc_00E355C3: mov ecx, 0000000Bh
  loc_00E355C8: mov edx, esp
  loc_00E355CA: xor eax, eax
  loc_00E355CC: mov var_18, eax
  loc_00E355CF: mov var_1C, eax
  loc_00E355D2: mov [edx], ecx
  loc_00E355D4: mov ecx, var_38
  loc_00E355D7: mov var_2C, eax
  loc_00E355DA: or eax, FFFFFFFFh
  loc_00E355DD: mov [edx+00000004h], ecx
  loc_00E355E0: mov ecx, [esi]
  loc_00E355E2: push 80010007h
  loc_00E355E7: push esi
  loc_00E355E8: mov [edx+00000008h], eax
  loc_00E355EB: mov eax, var_30
  loc_00E355EE: mov [edx+0000000Ch], eax
  loc_00E355F1: call [ecx+00000334h]
  loc_00E355F7: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E355FD: lea edx, var_18
  loc_00E35600: push eax
  loc_00E35601: push edx
  loc_00E35602: call edi
  loc_00E35604: push eax
  loc_00E35605: call [00401238h] ; __vbaLateIdSt
  loc_00E3560B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E35611: lea ecx, var_18
  loc_00E35614: call ebx
  loc_00E35616: mov eax, [esi]
  loc_00E35618: push esi
  loc_00E35619: call [eax+00000328h]
  loc_00E3561F: lea ecx, var_18
  loc_00E35622: push eax
  loc_00E35623: push ecx
  loc_00E35624: call edi
  loc_00E35626: mov edx, [eax]
  loc_00E35628: push eax
  loc_00E35629: mov var_50, eax
  loc_00E3562C: call [edx+00000204h]
  loc_00E35632: test eax, eax
  loc_00E35634: fnclex
  loc_00E35636: jge 00E3564Dh
  loc_00E35638: mov ecx, var_50
  loc_00E3563B: push 00000204h
  loc_00E35640: push 006DCB70h
  loc_00E35645: push ecx
  loc_00E35646: push eax
  loc_00E35647: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3564D: lea ecx, var_18
  loc_00E35650: call ebx
  loc_00E35652: mov edx, [esi]
  loc_00E35654: push esi
  loc_00E35655: call [edx+0000031Ch]
  loc_00E3565B: push eax
  loc_00E3565C: lea eax, var_18
  loc_00E3565F: push eax
  loc_00E35660: call edi
  loc_00E35662: mov ecx, [eax]
  loc_00E35664: push 00000000h
  loc_00E35666: push eax
  loc_00E35667: mov var_50, eax
  loc_00E3566A: call [ecx+0000009Ch]
  loc_00E35670: test eax, eax
  loc_00E35672: fnclex
  loc_00E35674: jge 00E3568Bh
  loc_00E35676: mov edx, var_50
  loc_00E35679: push 0000009Ch
  loc_00E3567E: push 006E03D4h
  loc_00E35683: push edx
  loc_00E35684: push eax
  loc_00E35685: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3568B: lea ecx, var_18
  loc_00E3568E: call ebx
  loc_00E35690: mov eax, [esi]
  loc_00E35692: push esi
  loc_00E35693: call [eax+00000330h]
  loc_00E35699: lea ecx, var_18
  loc_00E3569C: push eax
  loc_00E3569D: push ecx
  loc_00E3569E: call edi
  loc_00E356A0: mov edx, [eax]
  loc_00E356A2: push 00000000h
  loc_00E356A4: push eax
  loc_00E356A5: mov var_50, eax
  loc_00E356A8: call [edx+0000009Ch]
  loc_00E356AE: test eax, eax
  loc_00E356B0: fnclex
  loc_00E356B2: jge 00E356C9h
  loc_00E356B4: mov ecx, var_50
  loc_00E356B7: push 0000009Ch
  loc_00E356BC: push 006E03D4h
  loc_00E356C1: push ecx
  loc_00E356C2: push eax
  loc_00E356C3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E356C9: lea ecx, var_18
  loc_00E356CC: call ebx
  loc_00E356CE: mov edx, [esi]
  loc_00E356D0: push esi
  loc_00E356D1: call [edx+00000324h]
  loc_00E356D7: push eax
  loc_00E356D8: lea eax, var_18
  loc_00E356DB: push eax
  loc_00E356DC: call edi
  loc_00E356DE: mov ecx, [eax]
  loc_00E356E0: push 00000000h
  loc_00E356E2: push eax
  loc_00E356E3: mov var_50, eax
  loc_00E356E6: call [ecx+0000009Ch]
  loc_00E356EC: test eax, eax
  loc_00E356EE: fnclex
  loc_00E356F0: jge 00E35707h
  loc_00E356F2: mov edx, var_50
  loc_00E356F5: push 0000009Ch
  loc_00E356FA: push 006E03D4h
  loc_00E356FF: push edx
  loc_00E35700: push eax
  loc_00E35701: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35707: lea ecx, var_18
  loc_00E3570A: call ebx
  loc_00E3570C: mov eax, [esi]
  loc_00E3570E: push esi
  loc_00E3570F: call [eax+00000314h]
  loc_00E35715: lea ecx, var_18
  loc_00E35718: push eax
  loc_00E35719: push ecx
  loc_00E3571A: call edi
  loc_00E3571C: mov edx, [eax]
  loc_00E3571E: push 00000000h
  loc_00E35720: push eax
  loc_00E35721: mov var_50, eax
  loc_00E35724: call [edx+0000009Ch]
  loc_00E3572A: test eax, eax
  loc_00E3572C: fnclex
  loc_00E3572E: jge 00E35745h
  loc_00E35730: mov ecx, var_50
  loc_00E35733: push 0000009Ch
  loc_00E35738: push 006E03D4h
  loc_00E3573D: push ecx
  loc_00E3573E: push eax
  loc_00E3573F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35745: lea ecx, var_18
  loc_00E35748: call ebx
  loc_00E3574A: mov edx, [esi]
  loc_00E3574C: push esi
  loc_00E3574D: call [edx+00000320h]
  loc_00E35753: push eax
  loc_00E35754: lea eax, var_18
  loc_00E35757: push eax
  loc_00E35758: call edi
  loc_00E3575A: mov ecx, [eax]
  loc_00E3575C: push 00000000h
  loc_00E3575E: push eax
  loc_00E3575F: mov var_50, eax
  loc_00E35762: call [ecx+0000009Ch]
  loc_00E35768: test eax, eax
  loc_00E3576A: fnclex
  loc_00E3576C: jge 00E35783h
  loc_00E3576E: mov edx, var_50
  loc_00E35771: push 0000009Ch
  loc_00E35776: push 006E03D4h
  loc_00E3577B: push edx
  loc_00E3577C: push eax
  loc_00E3577D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E35783: lea ecx, var_18
  loc_00E35786: call ebx
  loc_00E35788: sub esp, 00000010h
  loc_00E3578B: mov ecx, 00000008h
  loc_00E35790: mov edx, esp
  loc_00E35792: mov eax, 006E0A78h ; "biodata"
  loc_00E35797: push 0000000Eh
  loc_00E35799: push esi
  loc_00E3579A: mov [edx], ecx
  loc_00E3579C: mov ecx, var_38
  loc_00E3579F: mov [edx+00000004h], ecx
  loc_00E357A2: mov ecx, [esi]
  loc_00E357A4: mov [edx+00000008h], eax
  loc_00E357A7: mov eax, var_30
  loc_00E357AA: mov [edx+0000000Ch], eax
  loc_00E357AD: call [ecx+00000358h]
  loc_00E357B3: lea edx, var_18
  loc_00E357B6: push eax
  loc_00E357B7: push edx
  loc_00E357B8: call edi
  loc_00E357BA: push eax
  loc_00E357BB: call [00401238h] ; __vbaLateIdSt
  loc_00E357C1: lea ecx, var_18
  loc_00E357C4: call ebx
  loc_00E357C6: mov eax, [esi]
  loc_00E357C8: push 00000000h
  loc_00E357CA: push 00000019h
  loc_00E357CC: push esi
  loc_00E357CD: call [eax+00000358h]
  loc_00E357D3: lea ecx, var_18
  loc_00E357D6: push eax
  loc_00E357D7: push ecx
  loc_00E357D8: call edi
  loc_00E357DA: push eax
  loc_00E357DB: call [00401030h] ; __vbaLateIdCall
  loc_00E357E1: add esp, 0000000Ch
  loc_00E357E4: lea ecx, var_18
  loc_00E357E7: call ebx
  loc_00E357E9: mov edx, [esi]
  loc_00E357EB: push 006E05D8h
  loc_00E357F0: push esi
  loc_00E357F1: call [edx+00000358h]
  loc_00E357F7: push eax
  loc_00E357F8: lea eax, var_18
  loc_00E357FB: push eax
  loc_00E357FC: call edi
  loc_00E357FE: push eax
  loc_00E357FF: call [00401224h] ; __vbaCastObj
  loc_00E35805: lea ecx, var_2C
  loc_00E35808: mov var_24, eax
  loc_00E3580B: push ecx
  loc_00E3580C: mov var_2C, 0000000Dh
  loc_00E35813: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E35819: mov ecx, [eax]
  loc_00E3581B: sub esp, 00000010h
  loc_00E3581E: mov edx, esp
  loc_00E35820: push 00000000h
  loc_00E35822: push 0000002Ah
  loc_00E35824: mov [edx], ecx
  loc_00E35826: mov ecx, [eax+00000004h]
  loc_00E35829: push esi
  loc_00E3582A: mov [edx+00000004h], ecx
  loc_00E3582D: mov ecx, [eax+00000008h]
  loc_00E35830: mov eax, [eax+0000000Ch]
  loc_00E35833: mov [edx+00000008h], ecx
  loc_00E35836: mov ecx, [esi]
  loc_00E35838: mov [edx+0000000Ch], eax
  loc_00E3583B: call [ecx+0000035Ch]
  loc_00E35841: lea edx, var_1C
  loc_00E35844: push eax
  loc_00E35845: push edx
  loc_00E35846: call edi
  loc_00E35848: push eax
  loc_00E35849: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3584F: lea eax, var_1C
  loc_00E35852: lea ecx, var_18
  loc_00E35855: push eax
  loc_00E35856: push ecx
  loc_00E35857: push 00000002h
  loc_00E35859: call [00401048h] ; __vbaFreeObjList
  loc_00E3585F: add esp, 00000028h
  loc_00E35862: lea ecx, var_2C
  loc_00E35865: call [00401024h] ; __vbaFreeVar
  loc_00E3586B: sub esp, 00000010h
  loc_00E3586E: mov ecx, 0000000Bh
  loc_00E35873: mov edx, esp
  loc_00E35875: xor eax, eax
  loc_00E35877: push 00000006h
  loc_00E35879: push esi
  loc_00E3587A: mov [edx], ecx
  loc_00E3587C: mov ecx, var_38
  loc_00E3587F: mov [edx+00000004h], ecx
  loc_00E35882: mov ecx, [esi]
  loc_00E35884: mov [edx+00000008h], eax
  loc_00E35887: mov eax, var_30
  loc_00E3588A: mov [edx+0000000Ch], eax
  loc_00E3588D: call [ecx+0000035Ch]
  loc_00E35893: lea edx, var_18
  loc_00E35896: push eax
  loc_00E35897: push edx
  loc_00E35898: call edi
  loc_00E3589A: push eax
  loc_00E3589B: call [00401238h] ; __vbaLateIdSt
  loc_00E358A1: lea ecx, var_18
  loc_00E358A4: call ebx
  loc_00E358A6: sub esp, 00000010h
  loc_00E358A9: mov ecx, 0000000Bh
  loc_00E358AE: mov edx, esp
  loc_00E358B0: xor eax, eax
  loc_00E358B2: push 8001000Eh
  loc_00E358B7: push esi
  loc_00E358B8: mov [edx], ecx
  loc_00E358BA: mov ecx, var_38
  loc_00E358BD: mov [edx+00000004h], ecx
  loc_00E358C0: mov ecx, [esi]
  loc_00E358C2: mov [edx+00000008h], eax
  loc_00E358C5: mov eax, var_30
  loc_00E358C8: mov [edx+0000000Ch], eax
  loc_00E358CB: call [ecx+0000035Ch]
  loc_00E358D1: lea edx, var_18
  loc_00E358D4: push eax
  loc_00E358D5: push edx
  loc_00E358D6: call edi
  loc_00E358D8: push eax
  loc_00E358D9: call [00401238h] ; __vbaLateIdSt
  loc_00E358DF: lea ecx, var_18
  loc_00E358E2: call ebx
  loc_00E358E4: mov eax, [esi]
  loc_00E358E6: push 00000000h
  loc_00E358E8: push FFFFFDDAh
  loc_00E358ED: push esi
  loc_00E358EE: call [eax+0000035Ch]
  loc_00E358F4: lea ecx, var_18
  loc_00E358F7: push eax
  loc_00E358F8: push ecx
  loc_00E358F9: call edi
  loc_00E358FB: push eax
  loc_00E358FC: call [00401030h] ; __vbaLateIdCall
  loc_00E35902: add esp, 0000000Ch
  loc_00E35905: lea ecx, var_18
  loc_00E35908: call ebx
  loc_00E3590A: mov var_4, 00000000h
  loc_00E35911: push 00E35936h
  loc_00E35916: jmp 00E35935h
  loc_00E35918: lea edx, var_1C
  loc_00E3591B: lea eax, var_18
  loc_00E3591E: push edx
  loc_00E3591F: push eax
  loc_00E35920: push 00000002h
  loc_00E35922: call [00401048h] ; __vbaFreeObjList
  loc_00E35928: add esp, 0000000Ch
  loc_00E3592B: lea ecx, var_2C
  loc_00E3592E: call [00401024h] ; __vbaFreeVar
  loc_00E35934: ret
  loc_00E35935: ret
  loc_00E35936: mov eax, Me
  loc_00E35939: push eax
  loc_00E3593A: mov ecx, [eax]
  loc_00E3593C: call [ecx+00000008h]
  loc_00E3593F: mov eax, var_4
  loc_00E35942: mov ecx, var_14
  loc_00E35945: pop edi
  loc_00E35946: pop esi
  loc_00E35947: mov fs:[00000000h], ecx
  loc_00E3594E: pop ebx
  loc_00E3594F: mov esp, ebp
  loc_00E35951: pop ebp
  loc_00E35952: retn 0004h
End Sub

Private Sub jcbutton3_UnknownEvent_9 'E33E20
  loc_00E33E20: push ebp
  loc_00E33E21: mov ebp, esp
  loc_00E33E23: sub esp, 0000000Ch
  loc_00E33E26: push 00402836h ; __vbaExceptHandler
  loc_00E33E2B: mov eax, fs:[00000000h]
  loc_00E33E31: push eax
  loc_00E33E32: mov fs:[00000000h], esp
  loc_00E33E39: sub esp, 00000018h
  loc_00E33E3C: push ebx
  loc_00E33E3D: push esi
  loc_00E33E3E: push edi
  loc_00E33E3F: mov var_C, esp
  loc_00E33E42: mov var_8, 00402640h
  loc_00E33E49: mov edi, Me
  loc_00E33E4C: mov eax, edi
  loc_00E33E4E: and eax, 00000001h
  loc_00E33E51: mov var_4, eax
  loc_00E33E54: and edi, FFFFFFFEh
  loc_00E33E57: push edi
  loc_00E33E58: mov Me, edi
  loc_00E33E5B: mov ecx, [edi]
  loc_00E33E5D: call [ecx+00000004h]
  loc_00E33E60: mov ecx, [00E3D13Ch]
  loc_00E33E66: xor eax, eax
  loc_00E33E68: cmp ecx, eax
  loc_00E33E6A: mov var_18, eax
  loc_00E33E6D: mov var_1C, eax
  loc_00E33E70: jnz 00E33E82h
  loc_00E33E72: push 00E3D13Ch
  loc_00E33E77: push 006CB450h
  loc_00E33E7C: call [004011A0h] ; __vbaNew2
  loc_00E33E82: mov esi, [00E3D13Ch]
  loc_00E33E88: mov edx, [edi]
  loc_00E33E8A: push 006E05D8h
  loc_00E33E8F: push edi
  loc_00E33E90: mov ebx, [esi]
  loc_00E33E92: call [edx+00000358h]
  loc_00E33E98: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E33E9E: push eax
  loc_00E33E9F: lea eax, var_18
  loc_00E33EA2: push eax
  loc_00E33EA3: call edi
  loc_00E33EA5: push eax
  loc_00E33EA6: call [00401224h] ; __vbaCastObj
  loc_00E33EAC: lea ecx, var_1C
  loc_00E33EAF: push eax
  loc_00E33EB0: push ecx
  loc_00E33EB1: call edi
  loc_00E33EB3: push eax
  loc_00E33EB4: push esi
  loc_00E33EB5: call [ebx+00000028h]
  loc_00E33EB8: test eax, eax
  loc_00E33EBA: fnclex
  loc_00E33EBC: jge 00E33ECDh
  loc_00E33EBE: push 00000028h
  loc_00E33EC0: push 006DD034h
  loc_00E33EC5: push esi
  loc_00E33EC6: push eax
  loc_00E33EC7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33ECD: lea edx, var_1C
  loc_00E33ED0: lea eax, var_18
  loc_00E33ED3: push edx
  loc_00E33ED4: push eax
  loc_00E33ED5: push 00000002h
  loc_00E33ED7: call [00401048h] ; __vbaFreeObjList
  loc_00E33EDD: mov eax, [00E3D13Ch]
  loc_00E33EE2: add esp, 0000000Ch
  loc_00E33EE5: test eax, eax
  loc_00E33EE7: jnz 00E33EFDh
  loc_00E33EE9: mov edi, [004011A0h] ; __vbaNew2
  loc_00E33EEF: push 00E3D13Ch
  loc_00E33EF4: push 006CB450h
  loc_00E33EF9: call edi
  loc_00E33EFB: jmp 00E33F03h
  loc_00E33EFD: mov edi, [004011A0h] ; __vbaNew2
  loc_00E33F03: mov esi, [00E3D13Ch]
  loc_00E33F09: push esi
  loc_00E33F0A: mov ecx, [esi]
  loc_00E33F0C: call [ecx+00000088h]
  loc_00E33F12: test eax, eax
  loc_00E33F14: fnclex
  loc_00E33F16: jge 00E33F2Ah
  loc_00E33F18: push 00000088h
  loc_00E33F1D: push 006DD034h
  loc_00E33F22: push esi
  loc_00E33F23: push eax
  loc_00E33F24: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33F2A: mov eax, [00E3D9CCh]
  loc_00E33F2F: test eax, eax
  loc_00E33F31: jnz 00E33F3Fh
  loc_00E33F33: push 00E3D9CCh
  loc_00E33F38: push 006DC960h
  loc_00E33F3D: call edi
  loc_00E33F3F: mov eax, [00E3D13Ch]
  loc_00E33F44: mov esi, [00E3D9CCh]
  loc_00E33F4A: test eax, eax
  loc_00E33F4C: jnz 00E33F5Ah
  loc_00E33F4E: push 00E3D13Ch
  loc_00E33F53: push 006CB450h
  loc_00E33F58: call edi
  loc_00E33F5A: mov edx, [00E3D13Ch]
  loc_00E33F60: mov ebx, [esi]
  loc_00E33F62: lea eax, var_18
  loc_00E33F65: push edx
  loc_00E33F66: push eax
  loc_00E33F67: call [004010B4h] ; __vbaObjSetAddref
  loc_00E33F6D: push eax
  loc_00E33F6E: push esi
  loc_00E33F6F: call [ebx+0000000Ch]
  loc_00E33F72: test eax, eax
  loc_00E33F74: fnclex
  loc_00E33F76: jge 00E33F87h
  loc_00E33F78: push 0000000Ch
  loc_00E33F7A: push 006DC950h
  loc_00E33F7F: push esi
  loc_00E33F80: push eax
  loc_00E33F81: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E33F87: lea ecx, var_18
  loc_00E33F8A: call [00401254h] ; __vbaFreeObj
  loc_00E33F90: mov eax, [00E3D13Ch]
  loc_00E33F95: test eax, eax
  loc_00E33F97: jnz 00E33FA5h
  loc_00E33F99: push 00E3D13Ch
  loc_00E33F9E: push 006CB450h
  loc_00E33FA3: call edi
  loc_00E33FA5: mov ecx, [00E3D13Ch]
  loc_00E33FAB: push 00000000h
  loc_00E33FAD: push 80011003h
  loc_00E33FB2: push ecx
  loc_00E33FB3: call [00401030h] ; __vbaLateIdCall
  loc_00E33FB9: add esp, 0000000Ch
  loc_00E33FBC: mov var_4, 00000000h
  loc_00E33FC3: push 00E33FDFh
  loc_00E33FC8: jmp 00E33FDEh
  loc_00E33FCA: lea edx, var_1C
  loc_00E33FCD: lea eax, var_18
  loc_00E33FD0: push edx
  loc_00E33FD1: push eax
  loc_00E33FD2: push 00000002h
  loc_00E33FD4: call [00401048h] ; __vbaFreeObjList
  loc_00E33FDA: add esp, 0000000Ch
  loc_00E33FDD: ret
  loc_00E33FDE: ret
  loc_00E33FDF: mov eax, Me
  loc_00E33FE2: push eax
  loc_00E33FE3: mov ecx, [eax]
  loc_00E33FE5: call [ecx+00000008h]
  loc_00E33FE8: mov eax, var_4
  loc_00E33FEB: mov ecx, var_14
  loc_00E33FEE: pop edi
  loc_00E33FEF: pop esi
  loc_00E33FF0: mov fs:[00000000h], ecx
  loc_00E33FF7: pop ebx
  loc_00E33FF8: mov esp, ebp
  loc_00E33FFA: pop ebp
  loc_00E33FFB: retn 0004h
End Sub
