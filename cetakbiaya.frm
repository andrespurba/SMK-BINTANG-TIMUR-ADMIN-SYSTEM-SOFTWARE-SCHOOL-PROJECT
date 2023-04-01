VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form cetakbiaya
  Caption = "Form1"
  BackColor = &H404040&
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  DrawStyle = 1
  BorderStyle = 0 'None
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 0
  ClientTop = 0
  ClientWidth = 11745
  ClientHeight = 5805
  ShowInTaskbar = 0   'False
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame ftk
    Caption = "Frame1"
    BackColor = &H404040&
    Left = 8820
    Top = 2370
    Width = 2505
    Height = 1845
    TabIndex = 0
    BorderStyle = 0 'None
    Begin Project1.jcbutton seluruh
      Left = -30
      Top = 1140
      Width = 2565
      Height = 705
      TabIndex = 1
      OleObjectBlob = "cetakbiaya.frx":0000
    End
    Begin Project1.jcbutton tunggal
      Left = 0
      Top = 480
      Width = 2565
      Height = 705
      TabIndex = 13
      OleObjectBlob = "cetakbiaya.frx":0234
    End
    Begin VB.Label Label2
      Caption = "-PILIH-"
      ForeColor = &H80FF80&
      Left = 870
      Top = 30
      Width = 1305
      Height = 315
      TabIndex = 2
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
    Begin VB.Shape Shape6
      Left = 0
      Top = 420
      Width = 14685
      Height = 75
      BorderStyle = 0 'Transparent
      FillColor = &HC0C0C0&
      FillStyle = 0
    End
  End
  Begin VB.Frame fcari
    Caption = "Cetak Berdasarkan"
    BackColor = &HE0E0E0&
    ForeColor = &H404040&
    Left = 120
    Top = 870
    Width = 11475
    Height = 1695
    TabIndex = 4
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
      Left = 8280
      Top = 510
      Width = 2265
      Height = 300
      TabIndex = 15
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
    Begin VB.OptionButton optjur
      Caption = "Jurusan"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 4830
      Top = 480
      Width = 1515
      Height = 300
      TabIndex = 14
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
      Left = 1290
      Top = 450
      Width = 1455
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
    Begin VB.OptionButton optno
      Caption = "No. Daftar"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 2910
      Top = 450
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
      TabIndex = 6
      BorderStyle = 0 'None
    End
    Begin VB.OptionButton optagama
      Caption = "Agama"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 6660
      Top = 510
      Width = 1695
      Height = 300
      TabIndex = 5
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
      Top = 900
      Width = 675
      Height = 225
      TabIndex = 9
      OleObjectBlob = "cetakbiaya.frx":043C
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
  Begin Project1.jcbutton ctk
    Left = 180
    Top = 2730
    Width = 11475
    Height = 555
    TabIndex = 3
    OleObjectBlob = "cetakbiaya.frx":0670
  End
  Begin MSAdodcLib.Adodc adopen
    Left = 300
    Top = 4980
    Width = 1200
    Height = 330
    Visible = 0   'False
    OleObjectBlob = "cetakbiaya.frx":0878
  End
  Begin MSDataGridLib.DataGrid dgpen
    Left = 150
    Top = 3450
    Width = 11475
    Height = 2265
    TabIndex = 10
    OleObjectBlob = "cetakbiaya.frx":0BAA
  End
  Begin VB.Shape Shape3
    Left = -120
    Top = 450
    Width = 14685
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Label Label4
    Caption = "Print System"
    ForeColor = &H80FF80&
    Left = 5190
    Top = 60
    Width = 1815
    Height = 315
    TabIndex = 12
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
    Picture = "cetakbiaya.frx":0D55
    Left = 90
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Label Label1
    Caption = "-RINCIAN BIAYA-"
    ForeColor = &H4000&
    Left = 4950
    Top = 510
    Width = 1995
    Height = 315
    TabIndex = 11
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
  Begin VB.Shape Shape33
    BorderColor = &HE0E0E0&
    Left = -120
    Top = 2640
    Width = 12495
    Height = 750
    BorderStyle = 3 'Dot
    FillColor = &HFFFF&
  End
  Begin VB.Shape Shape1
    Left = -120
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

Attribute VB_Name = "cetakbiaya"


Private Sub Form_Load() 'E37240
  loc_00E37240: push ebp
  loc_00E37241: mov ebp, esp
  loc_00E37243: sub esp, 0000000Ch
  loc_00E37246: push 00402836h ; __vbaExceptHandler
  loc_00E3724B: mov eax, fs:[00000000h]
  loc_00E37251: push eax
  loc_00E37252: mov fs:[00000000h], esp
  loc_00E37259: sub esp, 00000058h
  loc_00E3725C: push ebx
  loc_00E3725D: push esi
  loc_00E3725E: push edi
  loc_00E3725F: mov var_C, esp
  loc_00E37262: mov var_8, 00402710h
  loc_00E37269: mov esi, Me
  loc_00E3726C: mov eax, esi
  loc_00E3726E: and eax, 00000001h
  loc_00E37271: mov var_4, eax
  loc_00E37274: and esi, FFFFFFFEh
  loc_00E37277: push esi
  loc_00E37278: mov Me, esi
  loc_00E3727B: mov ecx, [esi]
  loc_00E3727D: call [ecx+00000004h]
  loc_00E37280: xor eax, eax
  loc_00E37282: lea edx, var_4C
  loc_00E37285: mov var_4C, eax
  loc_00E37288: lea ecx, var_24
  loc_00E3728B: mov var_24, eax
  loc_00E3728E: mov var_28, eax
  loc_00E37291: mov var_2C, eax
  loc_00E37294: mov var_3C, eax
  loc_00E37297: mov var_44, 006E15ECh ; "select * from lc"
  loc_00E3729E: mov var_4C, 00000008h
  loc_00E372A5: call [00401204h] ; __vbaVarCopy
  loc_00E372AB: lea edx, var_24
  loc_00E372AE: push edx
  loc_00E372AF: call [00401230h] ; __vbaStrVarCopy
  loc_00E372B5: sub esp, 00000010h
  loc_00E372B8: mov ecx, 00000008h
  loc_00E372BD: mov edx, esp
  loc_00E372BF: mov var_3C, ecx
  loc_00E372C2: mov var_34, eax
  loc_00E372C5: push 0000000Eh
  loc_00E372C7: mov [edx], ecx
  loc_00E372C9: mov ecx, var_38
  loc_00E372CC: push esi
  loc_00E372CD: mov [edx+00000004h], ecx
  loc_00E372D0: mov ecx, [esi]
  loc_00E372D2: mov [edx+00000008h], eax
  loc_00E372D5: mov eax, var_30
  loc_00E372D8: mov [edx+0000000Ch], eax
  loc_00E372DB: call [ecx+00000354h]
  loc_00E372E1: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E372E7: lea edx, var_28
  loc_00E372EA: push eax
  loc_00E372EB: push edx
  loc_00E372EC: call edi
  loc_00E372EE: push eax
  loc_00E372EF: call [00401238h] ; __vbaLateIdSt
  loc_00E372F5: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E372FB: lea ecx, var_28
  loc_00E372FE: call ebx
  loc_00E37300: lea ecx, var_3C
  loc_00E37303: call [00401024h] ; __vbaFreeVar
  loc_00E37309: mov eax, [esi]
  loc_00E3730B: push 00000000h
  loc_00E3730D: push 00000019h
  loc_00E3730F: push esi
  loc_00E37310: call [eax+00000354h]
  loc_00E37316: lea ecx, var_28
  loc_00E37319: push eax
  loc_00E3731A: push ecx
  loc_00E3731B: call edi
  loc_00E3731D: push eax
  loc_00E3731E: call [00401030h] ; __vbaLateIdCall
  loc_00E37324: add esp, 0000000Ch
  loc_00E37327: lea ecx, var_28
  loc_00E3732A: call ebx
  loc_00E3732C: mov edx, [esi]
  loc_00E3732E: push 006E05D8h
  loc_00E37333: push esi
  loc_00E37334: call [edx+00000354h]
  loc_00E3733A: push eax
  loc_00E3733B: lea eax, var_28
  loc_00E3733E: push eax
  loc_00E3733F: call edi
  loc_00E37341: push eax
  loc_00E37342: call [00401224h] ; __vbaCastObj
  loc_00E37348: mov var_34, eax
  loc_00E3734B: mov var_3C, 0000000Dh
  loc_00E37352: lea ecx, var_3C
  loc_00E37355: push ecx
  loc_00E37356: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3735C: mov ecx, [eax]
  loc_00E3735E: sub esp, 00000010h
  loc_00E37361: mov edx, esp
  loc_00E37363: push 00000000h
  loc_00E37365: push 0000002Ah
  loc_00E37367: mov [edx], ecx
  loc_00E37369: mov ecx, [eax+00000004h]
  loc_00E3736C: push esi
  loc_00E3736D: mov [edx+00000004h], ecx
  loc_00E37370: mov ecx, [eax+00000008h]
  loc_00E37373: mov eax, [eax+0000000Ch]
  loc_00E37376: mov [edx+00000008h], ecx
  loc_00E37379: mov ecx, [esi]
  loc_00E3737B: mov [edx+0000000Ch], eax
  loc_00E3737E: call [ecx+00000358h]
  loc_00E37384: lea edx, var_2C
  loc_00E37387: push eax
  loc_00E37388: push edx
  loc_00E37389: call edi
  loc_00E3738B: push eax
  loc_00E3738C: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E37392: lea eax, var_2C
  loc_00E37395: lea ecx, var_28
  loc_00E37398: push eax
  loc_00E37399: push ecx
  loc_00E3739A: push 00000002h
  loc_00E3739C: call [00401048h] ; __vbaFreeObjList
  loc_00E373A2: add esp, 00000028h
  loc_00E373A5: lea ecx, var_3C
  loc_00E373A8: call [00401024h] ; __vbaFreeVar
  loc_00E373AE: mov edx, [esi]
  loc_00E373B0: push 00000000h
  loc_00E373B2: push FFFFFDDAh
  loc_00E373B7: push esi
  loc_00E373B8: call [edx+00000358h]
  loc_00E373BE: push eax
  loc_00E373BF: lea eax, var_28
  loc_00E373C2: push eax
  loc_00E373C3: call edi
  loc_00E373C5: push eax
  loc_00E373C6: call [00401030h] ; __vbaLateIdCall
  loc_00E373CC: add esp, 0000000Ch
  loc_00E373CF: lea ecx, var_28
  loc_00E373D2: call ebx
  loc_00E373D4: mov ecx, [esi]
  loc_00E373D6: push esi
  loc_00E373D7: call [ecx+000002FCh]
  loc_00E373DD: lea edx, var_28
  loc_00E373E0: push eax
  loc_00E373E1: push edx
  loc_00E373E2: call edi
  loc_00E373E4: mov ecx, [eax]
  loc_00E373E6: push 00000000h
  loc_00E373E8: push eax
  loc_00E373E9: mov var_60, eax
  loc_00E373EC: call [ecx+0000009Ch]
  loc_00E373F2: test eax, eax
  loc_00E373F4: fnclex
  loc_00E373F6: jge 00E3740Dh
  loc_00E373F8: mov edx, var_60
  loc_00E373FB: push 0000009Ch
  loc_00E37400: push 006DCAD0h
  loc_00E37405: push edx
  loc_00E37406: push eax
  loc_00E37407: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3740D: lea ecx, var_28
  loc_00E37410: call ebx
  loc_00E37412: sub esp, 00000010h
  loc_00E37415: mov ecx, 0000000Bh
  loc_00E3741A: mov edx, esp
  loc_00E3741C: mov var_4C, ecx
  loc_00E3741F: xor eax, eax
  loc_00E37421: push 80010007h
  loc_00E37426: mov [edx], ecx
  loc_00E37428: mov ecx, var_48
  loc_00E3742B: mov var_44, eax
  loc_00E3742E: push esi
  loc_00E3742F: mov [edx+00000004h], ecx
  loc_00E37432: mov ecx, [esi]
  loc_00E37434: mov [edx+00000008h], eax
  loc_00E37437: mov eax, var_40
  loc_00E3743A: mov [edx+0000000Ch], eax
  loc_00E3743D: call [ecx+0000032Ch]
  loc_00E37443: lea edx, var_28
  loc_00E37446: push eax
  loc_00E37447: push edx
  loc_00E37448: call edi
  loc_00E3744A: push eax
  loc_00E3744B: call [00401238h] ; __vbaLateIdSt
  loc_00E37451: lea ecx, var_28
  loc_00E37454: call ebx
  loc_00E37456: mov eax, [esi]
  loc_00E37458: push esi
  loc_00E37459: call [eax+00000328h]
  loc_00E3745F: lea ecx, var_28
  loc_00E37462: push eax
  loc_00E37463: push ecx
  loc_00E37464: call edi
  loc_00E37466: mov edx, [eax]
  loc_00E37468: push 00000000h
  loc_00E3746A: push eax
  loc_00E3746B: mov var_60, eax
  loc_00E3746E: call [edx+000000E4h]
  loc_00E37474: test eax, eax
  loc_00E37476: fnclex
  loc_00E37478: jge 00E3748Fh
  loc_00E3747A: mov ecx, var_60
  loc_00E3747D: push 000000E4h
  loc_00E37482: push 006E03D4h
  loc_00E37487: push ecx
  loc_00E37488: push eax
  loc_00E37489: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3748F: lea ecx, var_28
  loc_00E37492: call ebx
  loc_00E37494: mov edx, [esi]
  loc_00E37496: push esi
  loc_00E37497: call [edx+00000320h]
  loc_00E3749D: push eax
  loc_00E3749E: lea eax, var_28
  loc_00E374A1: push eax
  loc_00E374A2: call edi
  loc_00E374A4: mov ecx, [eax]
  loc_00E374A6: push 00000000h
  loc_00E374A8: push eax
  loc_00E374A9: mov var_60, eax
  loc_00E374AC: call [ecx+000000E4h]
  loc_00E374B2: test eax, eax
  loc_00E374B4: fnclex
  loc_00E374B6: jge 00E374CDh
  loc_00E374B8: mov edx, var_60
  loc_00E374BB: push 000000E4h
  loc_00E374C0: push 006E03D4h
  loc_00E374C5: push edx
  loc_00E374C6: push eax
  loc_00E374C7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E374CD: lea ecx, var_28
  loc_00E374D0: call ebx
  loc_00E374D2: mov eax, [esi]
  loc_00E374D4: push esi
  loc_00E374D5: call [eax+0000031Ch]
  loc_00E374DB: lea ecx, var_28
  loc_00E374DE: push eax
  loc_00E374DF: push ecx
  loc_00E374E0: call edi
  loc_00E374E2: mov edx, [eax]
  loc_00E374E4: push 00000000h
  loc_00E374E6: push eax
  loc_00E374E7: mov var_60, eax
  loc_00E374EA: call [edx+000000E4h]
  loc_00E374F0: test eax, eax
  loc_00E374F2: fnclex
  loc_00E374F4: jge 00E3750Bh
  loc_00E374F6: mov ecx, var_60
  loc_00E374F9: push 000000E4h
  loc_00E374FE: push 006E03D4h
  loc_00E37503: push ecx
  loc_00E37504: push eax
  loc_00E37505: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3750B: lea ecx, var_28
  loc_00E3750E: call ebx
  loc_00E37510: mov edx, [esi]
  loc_00E37512: push esi
  loc_00E37513: call [edx+00000318h]
  loc_00E37519: push eax
  loc_00E3751A: lea eax, var_28
  loc_00E3751D: push eax
  loc_00E3751E: call edi
  loc_00E37520: mov ecx, [eax]
  loc_00E37522: push 00000000h
  loc_00E37524: push eax
  loc_00E37525: mov var_60, eax
  loc_00E37528: call [ecx+000000E4h]
  loc_00E3752E: test eax, eax
  loc_00E37530: fnclex
  loc_00E37532: jge 00E37549h
  loc_00E37534: mov edx, var_60
  loc_00E37537: push 000000E4h
  loc_00E3753C: push 006E03D4h
  loc_00E37541: push edx
  loc_00E37542: push eax
  loc_00E37543: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37549: lea ecx, var_28
  loc_00E3754C: call ebx
  loc_00E3754E: sub esp, 00000010h
  loc_00E37551: mov ecx, 0000000Bh
  loc_00E37556: mov edx, esp
  loc_00E37558: mov var_4C, ecx
  loc_00E3755B: xor eax, eax
  loc_00E3755D: push 80010007h
  loc_00E37562: mov [edx], ecx
  loc_00E37564: mov ecx, var_48
  loc_00E37567: mov var_44, eax
  loc_00E3756A: push esi
  loc_00E3756B: mov [edx+00000004h], ecx
  loc_00E3756E: mov ecx, [esi]
  loc_00E37570: mov [edx+00000008h], eax
  loc_00E37573: mov eax, var_40
  loc_00E37576: mov [edx+0000000Ch], eax
  loc_00E37579: call [ecx+0000032Ch]
  loc_00E3757F: lea edx, var_28
  loc_00E37582: push eax
  loc_00E37583: push edx
  loc_00E37584: call edi
  loc_00E37586: push eax
  loc_00E37587: call [00401238h] ; __vbaLateIdSt
  loc_00E3758D: lea ecx, var_28
  loc_00E37590: call ebx
  loc_00E37592: mov eax, [esi]
  loc_00E37594: push esi
  loc_00E37595: call [eax+000002FCh]
  loc_00E3759B: lea ecx, var_28
  loc_00E3759E: push eax
  loc_00E3759F: push ecx
  loc_00E375A0: call edi
  loc_00E375A2: mov esi, eax
  loc_00E375A4: push 00000000h
  loc_00E375A6: push esi
  loc_00E375A7: mov edx, [esi]
  loc_00E375A9: call [edx+0000009Ch]
  loc_00E375AF: test eax, eax
  loc_00E375B1: fnclex
  loc_00E375B3: jge 00E375C7h
  loc_00E375B5: push 0000009Ch
  loc_00E375BA: push 006DCAD0h
  loc_00E375BF: push esi
  loc_00E375C0: push eax
  loc_00E375C1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E375C7: lea ecx, var_28
  loc_00E375CA: call ebx
  loc_00E375CC: mov var_4, 00000000h
  loc_00E375D3: push 00E37601h
  loc_00E375D8: jmp 00E375F7h
  loc_00E375DA: lea eax, var_2C
  loc_00E375DD: lea ecx, var_28
  loc_00E375E0: push eax
  loc_00E375E1: push ecx
  loc_00E375E2: push 00000002h
  loc_00E375E4: call [00401048h] ; __vbaFreeObjList
  loc_00E375EA: add esp, 0000000Ch
  loc_00E375ED: lea ecx, var_3C
  loc_00E375F0: call [00401024h] ; __vbaFreeVar
  loc_00E375F6: ret
  loc_00E375F7: lea ecx, var_24
  loc_00E375FA: call [00401024h] ; __vbaFreeVar
  loc_00E37600: ret
  loc_00E37601: mov eax, Me
  loc_00E37604: push eax
  loc_00E37605: mov edx, [eax]
  loc_00E37607: call [edx+00000008h]
  loc_00E3760A: mov eax, var_4
  loc_00E3760D: mov ecx, var_14
  loc_00E37610: pop edi
  loc_00E37611: pop esi
  loc_00E37612: mov fs:[00000000h], ecx
  loc_00E37619: pop ebx
  loc_00E3761A: mov esp, ebp
  loc_00E3761C: pop ebp
  loc_00E3761D: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E36FF0
  loc_00E36FF0: push ebp
  loc_00E36FF1: mov ebp, esp
  loc_00E36FF3: sub esp, 0000000Ch
  loc_00E36FF6: push 00402836h ; __vbaExceptHandler
  loc_00E36FFB: mov eax, fs:[00000000h]
  loc_00E37001: push eax
  loc_00E37002: mov fs:[00000000h], esp
  loc_00E37009: sub esp, 0000005Ch
  loc_00E3700C: push ebx
  loc_00E3700D: push esi
  loc_00E3700E: push edi
  loc_00E3700F: mov var_C, esp
  loc_00E37012: mov var_8, 00402700h
  loc_00E37019: mov esi, Me
  loc_00E3701C: mov eax, esi
  loc_00E3701E: and eax, 00000001h
  loc_00E37021: mov var_4, eax
  loc_00E37024: and esi, FFFFFFFEh
  loc_00E37027: push esi
  loc_00E37028: mov Me, esi
  loc_00E3702B: mov ecx, [esi]
  loc_00E3702D: call [ecx+00000004h]
  loc_00E37030: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37036: xor eax, eax
  loc_00E37038: mov var_18, eax
  loc_00E3703B: mov var_4C, eax
  loc_00E3703E: mov var_50, eax
  loc_00E37041: mov edx, [esi]
  loc_00E37043: lea eax, var_4C
  loc_00E37046: push eax
  loc_00E37047: push esi
  loc_00E37048: call [edx+00000070h]
  loc_00E3704B: test eax, eax
  loc_00E3704D: fnclex
  loc_00E3704F: jge 00E3705Ch
  loc_00E37051: push 00000070h
  loc_00E37053: push 006DFEF8h
  loc_00E37058: push esi
  loc_00E37059: push eax
  loc_00E3705A: call ebx
  loc_00E3705C: fld real4 ptr var_4C
  loc_00E3705F: fadd st0, real4 ptr [00401298h]
  loc_00E37065: mov ecx, [esi]
  loc_00E37067: push ecx
  loc_00E37068: fnstsw ax
  loc_00E3706A: test al, 0Dh
  loc_00E3706C: jnz 00E37230h
  loc_00E37072: fstp real4 ptr [esp]
  loc_00E37075: push esi
  loc_00E37076: call [ecx+00000074h]
  loc_00E37079: test eax, eax
  loc_00E3707B: fnclex
  loc_00E3707D: jge 00E3708Ah
  loc_00E3707F: push 00000074h
  loc_00E37081: push 006DFEF8h
  loc_00E37086: push esi
  loc_00E37087: push eax
  loc_00E37088: call ebx
  loc_00E3708A: mov edx, [esi]
  loc_00E3708C: lea eax, var_4C
  loc_00E3708F: push eax
  loc_00E37090: push esi
  loc_00E37091: call [edx+00000070h]
  loc_00E37094: test eax, eax
  loc_00E37096: fnclex
  loc_00E37098: jge 00E370A5h
  loc_00E3709A: push 00000070h
  loc_00E3709C: push 006DFEF8h
  loc_00E370A1: push esi
  loc_00E370A2: push eax
  loc_00E370A3: call ebx
  loc_00E370A5: mov ecx, [esi]
  loc_00E370A7: lea edx, var_50
  loc_00E370AA: push edx
  loc_00E370AB: push esi
  loc_00E370AC: call [ecx+00000078h]
  loc_00E370AF: test eax, eax
  loc_00E370B1: fnclex
  loc_00E370B3: jge 00E370C0h
  loc_00E370B5: push 00000078h
  loc_00E370B7: push 006DFEF8h
  loc_00E370BC: push esi
  loc_00E370BD: push eax
  loc_00E370BE: call ebx
  loc_00E370C0: sub esp, 00000010h
  loc_00E370C3: mov ecx, 0000000Ah
  loc_00E370C8: mov ebx, esp
  loc_00E370CA: mov eax, 80020004h
  loc_00E370CF: mov edx, eax
  loc_00E370D1: sub esp, 00000010h
  loc_00E370D4: mov [ebx], ecx
  loc_00E370D6: mov ecx, var_44
  loc_00E370D9: mov edi, [esi]
  loc_00E370DB: mov [ebx+00000004h], ecx
  loc_00E370DE: mov ecx, esp
  loc_00E370E0: sub esp, 00000010h
  loc_00E370E3: mov [ebx+00000008h], eax
  loc_00E370E6: mov eax, var_3C
  loc_00E370E9: mov [ebx+0000000Ch], eax
  loc_00E370EC: mov eax, 0000000Ah
  loc_00E370F1: mov [ecx], eax
  loc_00E370F3: mov eax, var_34
  loc_00E370F6: mov [ecx+00000004h], eax
  loc_00E370F9: mov eax, 00000004h
  loc_00E370FE: mov [ecx+00000008h], edx
  loc_00E37101: mov edx, var_2C
  loc_00E37104: mov [ecx+0000000Ch], edx
  loc_00E37107: mov edx, var_24
  loc_00E3710A: mov ecx, esp
  loc_00E3710C: mov [ecx], eax
  loc_00E3710E: mov eax, var_50
  loc_00E37111: mov [ecx+00000004h], edx
  loc_00E37114: mov [ecx+00000008h], eax
  loc_00E37117: mov eax, var_1C
  loc_00E3711A: mov [ecx+0000000Ch], eax
  loc_00E3711D: mov ecx, var_4C
  loc_00E37120: push ecx
  loc_00E37121: push esi
  loc_00E37122: call [edi+000002A4h]
  loc_00E37128: test eax, eax
  loc_00E3712A: fnclex
  loc_00E3712C: jge 00E37144h
  loc_00E3712E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37134: push 000002A4h
  loc_00E37139: push 006DFEF8h
  loc_00E3713E: push esi
  loc_00E3713F: push eax
  loc_00E37140: call ebx
  loc_00E37142: jmp 00E3714Ah
  loc_00E37144: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3714A: call [004010BCh] ; rtcDoEvents
  loc_00E37150: mov edx, [esi]
  loc_00E37152: lea eax, var_4C
  loc_00E37155: push eax
  loc_00E37156: push esi
  loc_00E37157: call [edx+00000070h]
  loc_00E3715A: test eax, eax
  loc_00E3715C: fnclex
  loc_00E3715E: jge 00E3716Bh
  loc_00E37160: push 00000070h
  loc_00E37162: push 006DFEF8h
  loc_00E37167: push esi
  loc_00E37168: push eax
  loc_00E37169: call ebx
  loc_00E3716B: mov eax, [00E3D9CCh]
  loc_00E37170: test eax, eax
  loc_00E37172: jnz 00E37184h
  loc_00E37174: push 00E3D9CCh
  loc_00E37179: push 006DC960h
  loc_00E3717E: call [004011A0h] ; __vbaNew2
  loc_00E37184: mov edi, [00E3D9CCh]
  loc_00E3718A: lea edx, var_18
  loc_00E3718D: push edx
  loc_00E3718E: push edi
  loc_00E3718F: mov ecx, [edi]
  loc_00E37191: call [ecx+00000018h]
  loc_00E37194: test eax, eax
  loc_00E37196: fnclex
  loc_00E37198: jge 00E371A5h
  loc_00E3719A: push 00000018h
  loc_00E3719C: push 006DC950h
  loc_00E371A1: push edi
  loc_00E371A2: push eax
  loc_00E371A3: call ebx
  loc_00E371A5: mov eax, var_18
  loc_00E371A8: lea edx, var_50
  loc_00E371AB: push edx
  loc_00E371AC: push eax
  loc_00E371AD: mov ecx, [eax]
  loc_00E371AF: mov edi, eax
  loc_00E371B1: call [ecx+00000098h]
  loc_00E371B7: test eax, eax
  loc_00E371B9: fnclex
  loc_00E371BB: jge 00E371CBh
  loc_00E371BD: push 00000098h
  loc_00E371C2: push 006DCAF0h
  loc_00E371C7: push edi
  loc_00E371C8: push eax
  loc_00E371C9: call ebx
  loc_00E371CB: fld real4 ptr var_4C
  loc_00E371CE: fcomp real4 ptr var_50
  loc_00E371D1: fnstsw ax
  loc_00E371D3: test ah, 41h
  loc_00E371D6: jz 00E371DFh
  loc_00E371D8: mov eax, 00000001h
  loc_00E371DD: jmp 00E371E1h
  loc_00E371DF: xor eax, eax
  loc_00E371E1: neg eax
  loc_00E371E3: lea ecx, var_18
  loc_00E371E6: mov edi, eax
  loc_00E371E8: call [00401254h] ; __vbaFreeObj
  loc_00E371EE: test di, di
  loc_00E371F1: jnz 00E37041h
  loc_00E371F7: mov var_4, 00000000h
  loc_00E371FE: fwait
  loc_00E371FF: push 00E37211h
  loc_00E37204: jmp 00E37210h
  loc_00E37206: lea ecx, var_18
  loc_00E37209: call [00401254h] ; __vbaFreeObj
  loc_00E3720F: ret
  loc_00E37210: ret
  loc_00E37211: mov eax, Me
  loc_00E37214: push eax
  loc_00E37215: mov ecx, [eax]
  loc_00E37217: call [ecx+00000008h]
  loc_00E3721A: mov eax, var_4
  loc_00E3721D: mov ecx, var_14
  loc_00E37220: pop edi
  loc_00E37221: pop esi
  loc_00E37222: mov fs:[00000000h], ecx
  loc_00E37229: pop ebx
  loc_00E3722A: mov esp, ebp
  loc_00E3722C: pop ebp
  loc_00E3722D: retn 0008h
End Sub

Private Sub ctk_UnknownEvent_9 'E36A40
  loc_00E36A40: push ebp
  loc_00E36A41: mov ebp, esp
  loc_00E36A43: sub esp, 0000000Ch
  loc_00E36A46: push 00402836h ; __vbaExceptHandler
  loc_00E36A4B: mov eax, fs:[00000000h]
  loc_00E36A51: push eax
  loc_00E36A52: mov fs:[00000000h], esp
  loc_00E36A59: sub esp, 0000005Ch
  loc_00E36A5C: push ebx
  loc_00E36A5D: push esi
  loc_00E36A5E: push edi
  loc_00E36A5F: mov var_C, esp
  loc_00E36A62: mov var_8, 004026F0h
  loc_00E36A69: mov esi, Me
  loc_00E36A6C: mov eax, esi
  loc_00E36A6E: and eax, 00000001h
  loc_00E36A71: mov var_4, eax
  loc_00E36A74: and esi, FFFFFFFEh
  loc_00E36A77: push esi
  loc_00E36A78: mov Me, esi
  loc_00E36A7B: mov ecx, [esi]
  loc_00E36A7D: call [ecx+00000004h]
  loc_00E36A80: mov edx, [esi]
  loc_00E36A82: xor eax, eax
  loc_00E36A84: push eax
  loc_00E36A85: push FFFFFDFAh
  loc_00E36A8A: push esi
  loc_00E36A8B: mov var_18, eax
  loc_00E36A8E: mov var_1C, eax
  loc_00E36A91: mov var_20, eax
  loc_00E36A94: mov var_30, eax
  loc_00E36A97: mov var_40, eax
  loc_00E36A9A: mov var_54, eax
  loc_00E36A9D: call [edx+00000334h]
  loc_00E36AA3: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E36AA9: push eax
  loc_00E36AAA: lea eax, var_1C
  loc_00E36AAD: push eax
  loc_00E36AAE: call edi
  loc_00E36AB0: lea ecx, var_30
  loc_00E36AB3: push eax
  loc_00E36AB4: push ecx
  loc_00E36AB5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E36ABB: add esp, 00000010h
  loc_00E36ABE: push eax
  loc_00E36ABF: call [00401034h] ; __vbaStrVarMove
  loc_00E36AC5: mov edx, eax
  loc_00E36AC7: lea ecx, var_18
  loc_00E36ACA: call [00401228h] ; __vbaStrMove
  loc_00E36AD0: push eax
  loc_00E36AD1: push 006E1D48h ; "C E T A K"
  loc_00E36AD6: call [00401104h] ; __vbaStrCmp
  loc_00E36ADC: neg eax
  loc_00E36ADE: sbb eax, eax
  loc_00E36AE0: lea ecx, var_18
  loc_00E36AE3: inc eax
  loc_00E36AE4: neg eax
  loc_00E36AE6: mov var_58, ax
  loc_00E36AEA: call [00401258h] ; __vbaFreeStr
  loc_00E36AF0: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E36AF6: lea ecx, var_1C
  loc_00E36AF9: call ebx
  loc_00E36AFB: lea ecx, var_30
  loc_00E36AFE: call [00401024h] ; __vbaFreeVar
  loc_00E36B04: cmp var_58, 0000h
  loc_00E36B09: jz 00E36EA8h
  loc_00E36B0F: sub esp, 00000010h
  loc_00E36B12: mov ecx, 00000008h
  loc_00E36B17: mov edx, esp
  loc_00E36B19: mov eax, 006E1D60h ; "Close Menu"
  loc_00E36B1E: push FFFFFDFAh
  loc_00E36B23: push esi
  loc_00E36B24: mov [edx], ecx
  loc_00E36B26: mov ecx, var_3C
  loc_00E36B29: mov [edx+00000004h], ecx
  loc_00E36B2C: mov ecx, [esi]
  loc_00E36B2E: mov [edx+00000008h], eax
  loc_00E36B31: mov eax, var_34
  loc_00E36B34: mov [edx+0000000Ch], eax
  loc_00E36B37: call [ecx+00000334h]
  loc_00E36B3D: lea edx, var_1C
  loc_00E36B40: push eax
  loc_00E36B41: push edx
  loc_00E36B42: call edi
  loc_00E36B44: push eax
  loc_00E36B45: call [00401238h] ; __vbaLateIdSt
  loc_00E36B4B: lea ecx, var_1C
  loc_00E36B4E: call ebx
  loc_00E36B50: mov eax, [esi]
  loc_00E36B52: push esi
  loc_00E36B53: call [eax+00000320h]
  loc_00E36B59: lea ecx, var_1C
  loc_00E36B5C: push eax
  loc_00E36B5D: push ecx
  loc_00E36B5E: call edi
  loc_00E36B60: mov edx, [eax]
  loc_00E36B62: lea ecx, var_54
  loc_00E36B65: push ecx
  loc_00E36B66: push eax
  loc_00E36B67: mov var_58, eax
  loc_00E36B6A: call [edx+000000E0h]
  loc_00E36B70: test eax, eax
  loc_00E36B72: fnclex
  loc_00E36B74: jge 00E36B8Bh
  loc_00E36B76: mov edx, var_58
  loc_00E36B79: push 000000E0h
  loc_00E36B7E: push 006E03D4h
  loc_00E36B83: push edx
  loc_00E36B84: push eax
  loc_00E36B85: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36B8B: xor eax, eax
  loc_00E36B8D: cmp var_54, FFFFFFh
  loc_00E36B92: lea ecx, var_1C
  loc_00E36B95: setz al
  loc_00E36B98: neg eax
  loc_00E36B9A: mov var_60, eax
  loc_00E36B9D: call ebx
  loc_00E36B9F: cmp var_60, 0000h
  loc_00E36BA4: jnz 00E36C5Ah
  loc_00E36BAA: mov eax, [esi]
  loc_00E36BAC: push esi
  loc_00E36BAD: call [eax+0000031Ch]
  loc_00E36BB3: lea ecx, var_1C
  loc_00E36BB6: push eax
  loc_00E36BB7: push ecx
  loc_00E36BB8: call edi
  loc_00E36BBA: mov edx, [eax]
  loc_00E36BBC: lea ecx, var_54
  loc_00E36BBF: push ecx
  loc_00E36BC0: push eax
  loc_00E36BC1: mov var_58, eax
  loc_00E36BC4: call [edx+000000E0h]
  loc_00E36BCA: test eax, eax
  loc_00E36BCC: fnclex
  loc_00E36BCE: jge 00E36BE5h
  loc_00E36BD0: mov edx, var_58
  loc_00E36BD3: push 000000E0h
  loc_00E36BD8: push 006E03D4h
  loc_00E36BDD: push edx
  loc_00E36BDE: push eax
  loc_00E36BDF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36BE5: xor eax, eax
  loc_00E36BE7: cmp var_54, FFFFFFh
  loc_00E36BEC: lea ecx, var_1C
  loc_00E36BEF: setz al
  loc_00E36BF2: neg eax
  loc_00E36BF4: mov var_60, eax
  loc_00E36BF7: call ebx
  loc_00E36BF9: cmp var_60, 0000h
  loc_00E36BFE: jnz 00E36C5Ah
  loc_00E36C00: mov eax, [esi]
  loc_00E36C02: push esi
  loc_00E36C03: call [eax+00000328h]
  loc_00E36C09: lea ecx, var_1C
  loc_00E36C0C: push eax
  loc_00E36C0D: push ecx
  loc_00E36C0E: call edi
  loc_00E36C10: mov edx, [eax]
  loc_00E36C12: lea ecx, var_54
  loc_00E36C15: push ecx
  loc_00E36C16: push eax
  loc_00E36C17: mov var_58, eax
  loc_00E36C1A: call [edx+000000E0h]
  loc_00E36C20: test eax, eax
  loc_00E36C22: fnclex
  loc_00E36C24: jge 00E36C3Bh
  loc_00E36C26: mov edx, var_58
  loc_00E36C29: push 000000E0h
  loc_00E36C2E: push 006E03D4h
  loc_00E36C33: push edx
  loc_00E36C34: push eax
  loc_00E36C35: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36C3B: xor eax, eax
  loc_00E36C3D: cmp var_54, FFFFFFh
  loc_00E36C42: lea ecx, var_1C
  loc_00E36C45: setz al
  loc_00E36C48: neg eax
  loc_00E36C4A: mov var_60, eax
  loc_00E36C4D: call ebx
  loc_00E36C4F: cmp var_60, 0000h
  loc_00E36C54: jz 00E36D04h
  loc_00E36C5A: sub esp, 00000010h
  loc_00E36C5D: mov ecx, 0000000Bh
  loc_00E36C62: mov edx, esp
  loc_00E36C64: or eax, FFFFFFFFh
  loc_00E36C67: push 80010007h
  loc_00E36C6C: push esi
  loc_00E36C6D: mov [edx], ecx
  loc_00E36C6F: mov ecx, var_3C
  loc_00E36C72: mov [edx+00000004h], ecx
  loc_00E36C75: mov ecx, [esi]
  loc_00E36C77: mov [edx+00000008h], eax
  loc_00E36C7A: mov eax, var_34
  loc_00E36C7D: mov [edx+0000000Ch], eax
  loc_00E36C80: call [ecx+00000304h]
  loc_00E36C86: lea edx, var_1C
  loc_00E36C89: push eax
  loc_00E36C8A: push edx
  loc_00E36C8B: call edi
  loc_00E36C8D: push eax
  loc_00E36C8E: call [00401238h] ; __vbaLateIdSt
  loc_00E36C94: lea ecx, var_1C
  loc_00E36C97: call ebx
  loc_00E36C99: sub esp, 00000010h
  loc_00E36C9C: mov ecx, 0000000Bh
  loc_00E36CA1: mov edx, esp
  loc_00E36CA3: or eax, FFFFFFFFh
  loc_00E36CA6: push 80010007h
  loc_00E36CAB: push esi
  loc_00E36CAC: mov [edx], ecx
  loc_00E36CAE: mov ecx, var_3C
  loc_00E36CB1: mov [edx+00000004h], ecx
  loc_00E36CB4: mov ecx, [esi]
  loc_00E36CB6: mov [edx+00000008h], eax
  loc_00E36CB9: mov eax, var_34
  loc_00E36CBC: mov [edx+0000000Ch], eax
  loc_00E36CBF: call [ecx+00000300h]
  loc_00E36CC5: lea edx, var_1C
  loc_00E36CC8: push eax
  loc_00E36CC9: push edx
  loc_00E36CCA: call edi
  loc_00E36CCC: push eax
  loc_00E36CCD: call [00401238h] ; __vbaLateIdSt
  loc_00E36CD3: lea ecx, var_1C
  loc_00E36CD6: call ebx
  loc_00E36CD8: mov eax, [esi]
  loc_00E36CDA: push esi
  loc_00E36CDB: call [eax+000002FCh]
  loc_00E36CE1: lea ecx, var_1C
  loc_00E36CE4: push eax
  loc_00E36CE5: push ecx
  loc_00E36CE6: call edi
  loc_00E36CE8: mov esi, eax
  loc_00E36CEA: push FFFFFFFFh
  loc_00E36CEC: push esi
  loc_00E36CED: mov edx, [esi]
  loc_00E36CEF: call [edx+0000009Ch]
  loc_00E36CF5: test eax, eax
  loc_00E36CF7: fnclex
  loc_00E36CF9: jge 00E36F8Ah
  loc_00E36CFF: jmp 00E36F78h
  loc_00E36D04: mov eax, [00E3D114h]
  loc_00E36D09: test eax, eax
  loc_00E36D0B: jnz 00E36D1Dh
  loc_00E36D0D: push 00E3D114h
  loc_00E36D12: push 006CB640h
  loc_00E36D17: call [004011A0h] ; __vbaNew2
  loc_00E36D1D: mov eax, [00E3D114h]
  loc_00E36D22: push 006E05D8h
  loc_00E36D27: mov var_58, eax
  loc_00E36D2A: push esi
  loc_00E36D2B: mov ebx, [eax]
  loc_00E36D2D: mov eax, [esi]
  loc_00E36D2F: call [eax+00000354h]
  loc_00E36D35: lea ecx, var_1C
  loc_00E36D38: push eax
  loc_00E36D39: push ecx
  loc_00E36D3A: call edi
  loc_00E36D3C: push eax
  loc_00E36D3D: call [00401224h] ; __vbaCastObj
  loc_00E36D43: lea edx, var_20
  loc_00E36D46: push eax
  loc_00E36D47: push edx
  loc_00E36D48: call edi
  loc_00E36D4A: mov var_6C, ebx
  loc_00E36D4D: mov ebx, var_58
  loc_00E36D50: push eax
  loc_00E36D51: mov eax, var_6C
  loc_00E36D54: push ebx
  loc_00E36D55: call [eax+00000028h]
  loc_00E36D58: test eax, eax
  loc_00E36D5A: fnclex
  loc_00E36D5C: jge 00E36D6Dh
  loc_00E36D5E: push 00000028h
  loc_00E36D60: push 006DD034h
  loc_00E36D65: push ebx
  loc_00E36D66: push eax
  loc_00E36D67: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36D6D: lea ecx, var_20
  loc_00E36D70: lea edx, var_1C
  loc_00E36D73: push ecx
  loc_00E36D74: push edx
  loc_00E36D75: push 00000002h
  loc_00E36D77: call [00401048h] ; __vbaFreeObjList
  loc_00E36D7D: mov eax, [00E3D114h]
  loc_00E36D82: add esp, 0000000Ch
  loc_00E36D85: test eax, eax
  loc_00E36D87: jnz 00E36D99h
  loc_00E36D89: push 00E3D114h
  loc_00E36D8E: push 006CB640h
  loc_00E36D93: call [004011A0h] ; __vbaNew2
  loc_00E36D99: mov ebx, [00E3D114h]
  loc_00E36D9F: push ebx
  loc_00E36DA0: mov eax, [ebx]
  loc_00E36DA2: call [eax+00000088h]
  loc_00E36DA8: test eax, eax
  loc_00E36DAA: fnclex
  loc_00E36DAC: jge 00E36DC0h
  loc_00E36DAE: push 00000088h
  loc_00E36DB3: push 006DD034h
  loc_00E36DB8: push ebx
  loc_00E36DB9: push eax
  loc_00E36DBA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36DC0: mov eax, [00E3D9CCh]
  loc_00E36DC5: test eax, eax
  loc_00E36DC7: jnz 00E36DD9h
  loc_00E36DC9: push 00E3D9CCh
  loc_00E36DCE: push 006DC960h
  loc_00E36DD3: call [004011A0h] ; __vbaNew2
  loc_00E36DD9: mov eax, [00E3D114h]
  loc_00E36DDE: mov ebx, [00E3D9CCh]
  loc_00E36DE4: test eax, eax
  loc_00E36DE6: mov var_58, ebx
  loc_00E36DE9: jnz 00E36DFBh
  loc_00E36DEB: push 00E3D114h
  loc_00E36DF0: push 006CB640h
  loc_00E36DF5: call [004011A0h] ; __vbaNew2
  loc_00E36DFB: mov ecx, [00E3D114h]
  loc_00E36E01: mov ebx, [ebx]
  loc_00E36E03: lea edx, var_1C
  loc_00E36E06: push ecx
  loc_00E36E07: push edx
  loc_00E36E08: call [004010B4h] ; __vbaObjSetAddref
  loc_00E36E0E: mov var_70, ebx
  loc_00E36E11: mov ebx, var_58
  loc_00E36E14: push eax
  loc_00E36E15: mov eax, var_70
  loc_00E36E18: push ebx
  loc_00E36E19: call [eax+0000000Ch]
  loc_00E36E1C: test eax, eax
  loc_00E36E1E: fnclex
  loc_00E36E20: jge 00E36E31h
  loc_00E36E22: push 0000000Ch
  loc_00E36E24: push 006DC950h
  loc_00E36E29: push ebx
  loc_00E36E2A: push eax
  loc_00E36E2B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36E31: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E36E37: lea ecx, var_1C
  loc_00E36E3A: call ebx
  loc_00E36E3C: mov eax, [00E3D114h]
  loc_00E36E41: test eax, eax
  loc_00E36E43: jnz 00E36E55h
  loc_00E36E45: push 00E3D114h
  loc_00E36E4A: push 006CB640h
  loc_00E36E4F: call [004011A0h] ; __vbaNew2
  loc_00E36E55: mov ecx, [00E3D114h]
  loc_00E36E5B: push 00000000h
  loc_00E36E5D: push 80011003h
  loc_00E36E62: push ecx
  loc_00E36E63: call [00401030h] ; __vbaLateIdCall
  loc_00E36E69: mov ecx, 00000008h
  loc_00E36E6E: mov eax, 006E1D48h ; "C E T A K"
  loc_00E36E73: push ecx
  loc_00E36E74: mov edx, esp
  loc_00E36E76: push FFFFFDFAh
  loc_00E36E7B: push esi
  loc_00E36E7C: mov [edx], ecx
  loc_00E36E7E: mov ecx, var_3C
  loc_00E36E81: mov [edx+00000004h], ecx
  loc_00E36E84: mov ecx, [esi]
  loc_00E36E86: mov [edx+00000008h], eax
  loc_00E36E89: mov eax, var_34
  loc_00E36E8C: mov [edx+0000000Ch], eax
  loc_00E36E8F: call [ecx+00000334h]
  loc_00E36E95: lea edx, var_1C
  loc_00E36E98: push eax
  loc_00E36E99: push edx
  loc_00E36E9A: call edi
  loc_00E36E9C: push eax
  loc_00E36E9D: call [00401238h] ; __vbaLateIdSt
  loc_00E36EA3: jmp 00E36F8Ah
  loc_00E36EA8: mov eax, [esi]
  loc_00E36EAA: push 00000000h
  loc_00E36EAC: push FFFFFDFAh
  loc_00E36EB1: push esi
  loc_00E36EB2: call [eax+00000334h]
  loc_00E36EB8: lea ecx, var_1C
  loc_00E36EBB: push eax
  loc_00E36EBC: push ecx
  loc_00E36EBD: call edi
  loc_00E36EBF: lea edx, var_30
  loc_00E36EC2: push eax
  loc_00E36EC3: push edx
  loc_00E36EC4: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E36ECA: add esp, 00000010h
  loc_00E36ECD: push eax
  loc_00E36ECE: call [00401034h] ; __vbaStrVarMove
  loc_00E36ED4: mov edx, eax
  loc_00E36ED6: lea ecx, var_18
  loc_00E36ED9: call [00401228h] ; __vbaStrMove
  loc_00E36EDF: push eax
  loc_00E36EE0: push 006E1D60h ; "Close Menu"
  loc_00E36EE5: call [00401104h] ; __vbaStrCmp
  loc_00E36EEB: neg eax
  loc_00E36EED: sbb eax, eax
  loc_00E36EEF: lea ecx, var_18
  loc_00E36EF2: inc eax
  loc_00E36EF3: neg eax
  loc_00E36EF5: mov var_58, ax
  loc_00E36EF9: call [00401258h] ; __vbaFreeStr
  loc_00E36EFF: lea ecx, var_1C
  loc_00E36F02: call ebx
  loc_00E36F04: lea ecx, var_30
  loc_00E36F07: call [00401024h] ; __vbaFreeVar
  loc_00E36F0D: cmp var_58, 0000h
  loc_00E36F12: jz 00E36F8Fh
  loc_00E36F14: sub esp, 00000010h
  loc_00E36F17: mov ecx, 00000008h
  loc_00E36F1C: mov edx, esp
  loc_00E36F1E: mov eax, 006E1D48h ; "C E T A K"
  loc_00E36F23: push FFFFFDFAh
  loc_00E36F28: push esi
  loc_00E36F29: mov [edx], ecx
  loc_00E36F2B: mov ecx, var_3C
  loc_00E36F2E: mov [edx+00000004h], ecx
  loc_00E36F31: mov ecx, [esi]
  loc_00E36F33: mov [edx+00000008h], eax
  loc_00E36F36: mov eax, var_34
  loc_00E36F39: mov [edx+0000000Ch], eax
  loc_00E36F3C: call [ecx+00000334h]
  loc_00E36F42: lea edx, var_1C
  loc_00E36F45: push eax
  loc_00E36F46: push edx
  loc_00E36F47: call edi
  loc_00E36F49: push eax
  loc_00E36F4A: call [00401238h] ; __vbaLateIdSt
  loc_00E36F50: lea ecx, var_1C
  loc_00E36F53: call ebx
  loc_00E36F55: mov eax, [esi]
  loc_00E36F57: push esi
  loc_00E36F58: call [eax+000002FCh]
  loc_00E36F5E: lea ecx, var_1C
  loc_00E36F61: push eax
  loc_00E36F62: push ecx
  loc_00E36F63: call edi
  loc_00E36F65: mov esi, eax
  loc_00E36F67: push 00000000h
  loc_00E36F69: push esi
  loc_00E36F6A: mov edx, [esi]
  loc_00E36F6C: call [edx+0000009Ch]
  loc_00E36F72: test eax, eax
  loc_00E36F74: fnclex
  loc_00E36F76: jge 00E36F8Ah
  loc_00E36F78: push 0000009Ch
  loc_00E36F7D: push 006DCAD0h
  loc_00E36F82: push esi
  loc_00E36F83: push eax
  loc_00E36F84: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36F8A: lea ecx, var_1C
  loc_00E36F8D: call ebx
  loc_00E36F8F: mov var_4, 00000000h
  loc_00E36F96: push 00E36FC4h
  loc_00E36F9B: jmp 00E36FC3h
  loc_00E36F9D: lea ecx, var_18
  loc_00E36FA0: call [00401258h] ; __vbaFreeStr
  loc_00E36FA6: lea eax, var_20
  loc_00E36FA9: lea ecx, var_1C
  loc_00E36FAC: push eax
  loc_00E36FAD: push ecx
  loc_00E36FAE: push 00000002h
  loc_00E36FB0: call [00401048h] ; __vbaFreeObjList
  loc_00E36FB6: add esp, 0000000Ch
  loc_00E36FB9: lea ecx, var_30
  loc_00E36FBC: call [00401024h] ; __vbaFreeVar
  loc_00E36FC2: ret
  loc_00E36FC3: ret
  loc_00E36FC4: mov eax, Me
  loc_00E36FC7: push eax
  loc_00E36FC8: mov edx, [eax]
  loc_00E36FCA: call [edx+00000008h]
  loc_00E36FCD: mov eax, var_4
  loc_00E36FD0: mov ecx, var_14
  loc_00E36FD3: pop edi
  loc_00E36FD4: pop esi
  loc_00E36FD5: mov fs:[00000000h], ecx
  loc_00E36FDC: pop ebx
  loc_00E36FDD: mov esp, ebp
  loc_00E36FDF: pop ebp
  loc_00E36FE0: retn 0004h
End Sub

Private Sub tunggal_UnknownEvent_9 'E38780
  loc_00E38780: push ebp
  loc_00E38781: mov ebp, esp
  loc_00E38783: sub esp, 0000000Ch
  loc_00E38786: push 00402836h ; __vbaExceptHandler
  loc_00E3878B: mov eax, fs:[00000000h]
  loc_00E38791: push eax
  loc_00E38792: mov fs:[00000000h], esp
  loc_00E38799: sub esp, 00000018h
  loc_00E3879C: push ebx
  loc_00E3879D: push esi
  loc_00E3879E: push edi
  loc_00E3879F: mov var_C, esp
  loc_00E387A2: mov var_8, 00402770h
  loc_00E387A9: mov edi, Me
  loc_00E387AC: mov eax, edi
  loc_00E387AE: and eax, 00000001h
  loc_00E387B1: mov var_4, eax
  loc_00E387B4: and edi, FFFFFFFEh
  loc_00E387B7: push edi
  loc_00E387B8: mov Me, edi
  loc_00E387BB: mov ecx, [edi]
  loc_00E387BD: call [ecx+00000004h]
  loc_00E387C0: mov ecx, [00E3D150h]
  loc_00E387C6: xor eax, eax
  loc_00E387C8: cmp ecx, eax
  loc_00E387CA: mov var_18, eax
  loc_00E387CD: mov var_1C, eax
  loc_00E387D0: jnz 00E387E2h
  loc_00E387D2: push 00E3D150h
  loc_00E387D7: push 006CB358h
  loc_00E387DC: call [004011A0h] ; __vbaNew2
  loc_00E387E2: mov esi, [00E3D150h]
  loc_00E387E8: mov edx, [edi]
  loc_00E387EA: push 006E05D8h
  loc_00E387EF: push edi
  loc_00E387F0: mov ebx, [esi]
  loc_00E387F2: call [edx+00000354h]
  loc_00E387F8: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E387FE: push eax
  loc_00E387FF: lea eax, var_18
  loc_00E38802: push eax
  loc_00E38803: call edi
  loc_00E38805: push eax
  loc_00E38806: call [00401224h] ; __vbaCastObj
  loc_00E3880C: lea ecx, var_1C
  loc_00E3880F: push eax
  loc_00E38810: push ecx
  loc_00E38811: call edi
  loc_00E38813: push eax
  loc_00E38814: push esi
  loc_00E38815: call [ebx+00000028h]
  loc_00E38818: test eax, eax
  loc_00E3881A: fnclex
  loc_00E3881C: jge 00E3882Dh
  loc_00E3881E: push 00000028h
  loc_00E38820: push 006DD034h
  loc_00E38825: push esi
  loc_00E38826: push eax
  loc_00E38827: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3882D: lea edx, var_1C
  loc_00E38830: lea eax, var_18
  loc_00E38833: push edx
  loc_00E38834: push eax
  loc_00E38835: push 00000002h
  loc_00E38837: call [00401048h] ; __vbaFreeObjList
  loc_00E3883D: mov eax, [00E3D150h]
  loc_00E38842: add esp, 0000000Ch
  loc_00E38845: test eax, eax
  loc_00E38847: jnz 00E3885Dh
  loc_00E38849: mov edi, [004011A0h] ; __vbaNew2
  loc_00E3884F: push 00E3D150h
  loc_00E38854: push 006CB358h
  loc_00E38859: call edi
  loc_00E3885B: jmp 00E38863h
  loc_00E3885D: mov edi, [004011A0h] ; __vbaNew2
  loc_00E38863: mov esi, [00E3D150h]
  loc_00E38869: push esi
  loc_00E3886A: mov ecx, [esi]
  loc_00E3886C: call [ecx+00000088h]
  loc_00E38872: test eax, eax
  loc_00E38874: fnclex
  loc_00E38876: jge 00E3888Ah
  loc_00E38878: push 00000088h
  loc_00E3887D: push 006DD034h
  loc_00E38882: push esi
  loc_00E38883: push eax
  loc_00E38884: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3888A: mov eax, [00E3D9CCh]
  loc_00E3888F: test eax, eax
  loc_00E38891: jnz 00E3889Fh
  loc_00E38893: push 00E3D9CCh
  loc_00E38898: push 006DC960h
  loc_00E3889D: call edi
  loc_00E3889F: mov eax, [00E3D150h]
  loc_00E388A4: mov esi, [00E3D9CCh]
  loc_00E388AA: test eax, eax
  loc_00E388AC: jnz 00E388BAh
  loc_00E388AE: push 00E3D150h
  loc_00E388B3: push 006CB358h
  loc_00E388B8: call edi
  loc_00E388BA: mov edx, [00E3D150h]
  loc_00E388C0: mov ebx, [esi]
  loc_00E388C2: lea eax, var_18
  loc_00E388C5: push edx
  loc_00E388C6: push eax
  loc_00E388C7: call [004010B4h] ; __vbaObjSetAddref
  loc_00E388CD: push eax
  loc_00E388CE: push esi
  loc_00E388CF: call [ebx+0000000Ch]
  loc_00E388D2: test eax, eax
  loc_00E388D4: fnclex
  loc_00E388D6: jge 00E388E7h
  loc_00E388D8: push 0000000Ch
  loc_00E388DA: push 006DC950h
  loc_00E388DF: push esi
  loc_00E388E0: push eax
  loc_00E388E1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E388E7: lea ecx, var_18
  loc_00E388EA: call [00401254h] ; __vbaFreeObj
  loc_00E388F0: mov eax, [00E3D150h]
  loc_00E388F5: test eax, eax
  loc_00E388F7: jnz 00E38905h
  loc_00E388F9: push 00E3D150h
  loc_00E388FE: push 006CB358h
  loc_00E38903: call edi
  loc_00E38905: mov ecx, [00E3D150h]
  loc_00E3890B: push 00000000h
  loc_00E3890D: push 80011003h
  loc_00E38912: push ecx
  loc_00E38913: call [00401030h] ; __vbaLateIdCall
  loc_00E38919: add esp, 0000000Ch
  loc_00E3891C: mov var_4, 00000000h
  loc_00E38923: push 00E3893Fh
  loc_00E38928: jmp 00E3893Eh
  loc_00E3892A: lea edx, var_1C
  loc_00E3892D: lea eax, var_18
  loc_00E38930: push edx
  loc_00E38931: push eax
  loc_00E38932: push 00000002h
  loc_00E38934: call [00401048h] ; __vbaFreeObjList
  loc_00E3893A: add esp, 0000000Ch
  loc_00E3893D: ret
  loc_00E3893E: ret
  loc_00E3893F: mov eax, Me
  loc_00E38942: push eax
  loc_00E38943: mov ecx, [eax]
  loc_00E38945: call [ecx+00000008h]
  loc_00E38948: mov eax, var_4
  loc_00E3894B: mov ecx, var_14
  loc_00E3894E: pop edi
  loc_00E3894F: pop esi
  loc_00E38950: mov fs:[00000000h], ecx
  loc_00E38957: pop ebx
  loc_00E38958: mov esp, ebp
  loc_00E3895A: pop ebp
  loc_00E3895B: retn 0004h
End Sub

Private Sub seluruh_UnknownEvent_9 'E385A0
  loc_00E385A0: push ebp
  loc_00E385A1: mov ebp, esp
  loc_00E385A3: sub esp, 0000000Ch
  loc_00E385A6: push 00402836h ; __vbaExceptHandler
  loc_00E385AB: mov eax, fs:[00000000h]
  loc_00E385B1: push eax
  loc_00E385B2: mov fs:[00000000h], esp
  loc_00E385B9: sub esp, 00000018h
  loc_00E385BC: push ebx
  loc_00E385BD: push esi
  loc_00E385BE: push edi
  loc_00E385BF: mov var_C, esp
  loc_00E385C2: mov var_8, 00402760h
  loc_00E385C9: mov edi, Me
  loc_00E385CC: mov eax, edi
  loc_00E385CE: and eax, 00000001h
  loc_00E385D1: mov var_4, eax
  loc_00E385D4: and edi, FFFFFFFEh
  loc_00E385D7: push edi
  loc_00E385D8: mov Me, edi
  loc_00E385DB: mov ecx, [edi]
  loc_00E385DD: call [ecx+00000004h]
  loc_00E385E0: mov ecx, [00E3D114h]
  loc_00E385E6: xor eax, eax
  loc_00E385E8: cmp ecx, eax
  loc_00E385EA: mov var_18, eax
  loc_00E385ED: mov var_1C, eax
  loc_00E385F0: jnz 00E38602h
  loc_00E385F2: push 00E3D114h
  loc_00E385F7: push 006CB640h
  loc_00E385FC: call [004011A0h] ; __vbaNew2
  loc_00E38602: mov esi, [00E3D114h]
  loc_00E38608: mov edx, [edi]
  loc_00E3860A: push 006E05D8h
  loc_00E3860F: push edi
  loc_00E38610: mov ebx, [esi]
  loc_00E38612: call [edx+00000354h]
  loc_00E38618: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3861E: push eax
  loc_00E3861F: lea eax, var_18
  loc_00E38622: push eax
  loc_00E38623: call edi
  loc_00E38625: push eax
  loc_00E38626: call [00401224h] ; __vbaCastObj
  loc_00E3862C: lea ecx, var_1C
  loc_00E3862F: push eax
  loc_00E38630: push ecx
  loc_00E38631: call edi
  loc_00E38633: push eax
  loc_00E38634: push esi
  loc_00E38635: call [ebx+00000028h]
  loc_00E38638: test eax, eax
  loc_00E3863A: fnclex
  loc_00E3863C: jge 00E3864Dh
  loc_00E3863E: push 00000028h
  loc_00E38640: push 006DD034h
  loc_00E38645: push esi
  loc_00E38646: push eax
  loc_00E38647: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3864D: lea edx, var_1C
  loc_00E38650: lea eax, var_18
  loc_00E38653: push edx
  loc_00E38654: push eax
  loc_00E38655: push 00000002h
  loc_00E38657: call [00401048h] ; __vbaFreeObjList
  loc_00E3865D: mov eax, [00E3D114h]
  loc_00E38662: add esp, 0000000Ch
  loc_00E38665: test eax, eax
  loc_00E38667: jnz 00E3867Dh
  loc_00E38669: mov edi, [004011A0h] ; __vbaNew2
  loc_00E3866F: push 00E3D114h
  loc_00E38674: push 006CB640h
  loc_00E38679: call edi
  loc_00E3867B: jmp 00E38683h
  loc_00E3867D: mov edi, [004011A0h] ; __vbaNew2
  loc_00E38683: mov esi, [00E3D114h]
  loc_00E38689: push esi
  loc_00E3868A: mov ecx, [esi]
  loc_00E3868C: call [ecx+00000088h]
  loc_00E38692: test eax, eax
  loc_00E38694: fnclex
  loc_00E38696: jge 00E386AAh
  loc_00E38698: push 00000088h
  loc_00E3869D: push 006DD034h
  loc_00E386A2: push esi
  loc_00E386A3: push eax
  loc_00E386A4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E386AA: mov eax, [00E3D9CCh]
  loc_00E386AF: test eax, eax
  loc_00E386B1: jnz 00E386BFh
  loc_00E386B3: push 00E3D9CCh
  loc_00E386B8: push 006DC960h
  loc_00E386BD: call edi
  loc_00E386BF: mov eax, [00E3D114h]
  loc_00E386C4: mov esi, [00E3D9CCh]
  loc_00E386CA: test eax, eax
  loc_00E386CC: jnz 00E386DAh
  loc_00E386CE: push 00E3D114h
  loc_00E386D3: push 006CB640h
  loc_00E386D8: call edi
  loc_00E386DA: mov edx, [00E3D114h]
  loc_00E386E0: mov ebx, [esi]
  loc_00E386E2: lea eax, var_18
  loc_00E386E5: push edx
  loc_00E386E6: push eax
  loc_00E386E7: call [004010B4h] ; __vbaObjSetAddref
  loc_00E386ED: push eax
  loc_00E386EE: push esi
  loc_00E386EF: call [ebx+0000000Ch]
  loc_00E386F2: test eax, eax
  loc_00E386F4: fnclex
  loc_00E386F6: jge 00E38707h
  loc_00E386F8: push 0000000Ch
  loc_00E386FA: push 006DC950h
  loc_00E386FF: push esi
  loc_00E38700: push eax
  loc_00E38701: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38707: lea ecx, var_18
  loc_00E3870A: call [00401254h] ; __vbaFreeObj
  loc_00E38710: mov eax, [00E3D114h]
  loc_00E38715: test eax, eax
  loc_00E38717: jnz 00E38725h
  loc_00E38719: push 00E3D114h
  loc_00E3871E: push 006CB640h
  loc_00E38723: call edi
  loc_00E38725: mov ecx, [00E3D114h]
  loc_00E3872B: push 00000000h
  loc_00E3872D: push 80011003h
  loc_00E38732: push ecx
  loc_00E38733: call [00401030h] ; __vbaLateIdCall
  loc_00E38739: add esp, 0000000Ch
  loc_00E3873C: mov var_4, 00000000h
  loc_00E38743: push 00E3875Fh
  loc_00E38748: jmp 00E3875Eh
  loc_00E3874A: lea edx, var_1C
  loc_00E3874D: lea eax, var_18
  loc_00E38750: push edx
  loc_00E38751: push eax
  loc_00E38752: push 00000002h
  loc_00E38754: call [00401048h] ; __vbaFreeObjList
  loc_00E3875A: add esp, 0000000Ch
  loc_00E3875D: ret
  loc_00E3875E: ret
  loc_00E3875F: mov eax, Me
  loc_00E38762: push eax
  loc_00E38763: mov ecx, [eax]
  loc_00E38765: call [ecx+00000008h]
  loc_00E38768: mov eax, var_4
  loc_00E3876B: mov ecx, var_14
  loc_00E3876E: pop edi
  loc_00E3876F: pop esi
  loc_00E38770: mov fs:[00000000h], ecx
  loc_00E38777: pop ebx
  loc_00E38778: mov esp, ebp
  loc_00E3877A: pop ebp
  loc_00E3877B: retn 0004h
End Sub

Private Sub optnama_Click() 'E37DE0
  loc_00E37DE0: push ebp
  loc_00E37DE1: mov ebp, esp
  loc_00E37DE3: sub esp, 0000000Ch
  loc_00E37DE6: push 00402836h ; __vbaExceptHandler
  loc_00E37DEB: mov eax, fs:[00000000h]
  loc_00E37DF1: push eax
  loc_00E37DF2: mov fs:[00000000h], esp
  loc_00E37DF9: sub esp, 00000048h
  loc_00E37DFC: push ebx
  loc_00E37DFD: push esi
  loc_00E37DFE: push edi
  loc_00E37DFF: mov var_C, esp
  loc_00E37E02: mov var_8, 00402740h
  loc_00E37E09: mov esi, Me
  loc_00E37E0C: mov eax, esi
  loc_00E37E0E: and eax, 00000001h
  loc_00E37E11: mov var_4, eax
  loc_00E37E14: and esi, FFFFFFFEh
  loc_00E37E17: push esi
  loc_00E37E18: mov Me, esi
  loc_00E37E1B: mov ecx, [esi]
  loc_00E37E1D: call [ecx+00000004h]
  loc_00E37E20: sub esp, 00000010h
  loc_00E37E23: mov ecx, 0000000Bh
  loc_00E37E28: mov edx, esp
  loc_00E37E2A: xor eax, eax
  loc_00E37E2C: mov var_18, eax
  loc_00E37E2F: mov var_1C, eax
  loc_00E37E32: mov [edx], ecx
  loc_00E37E34: mov ecx, var_38
  loc_00E37E37: mov var_2C, eax
  loc_00E37E3A: or eax, FFFFFFFFh
  loc_00E37E3D: mov [edx+00000004h], ecx
  loc_00E37E40: mov ecx, [esi]
  loc_00E37E42: push 80010007h
  loc_00E37E47: push esi
  loc_00E37E48: mov [edx+00000008h], eax
  loc_00E37E4B: mov eax, var_30
  loc_00E37E4E: mov [edx+0000000Ch], eax
  loc_00E37E51: call [ecx+0000032Ch]
  loc_00E37E57: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E37E5D: lea edx, var_18
  loc_00E37E60: push eax
  loc_00E37E61: push edx
  loc_00E37E62: call edi
  loc_00E37E64: push eax
  loc_00E37E65: call [00401238h] ; __vbaLateIdSt
  loc_00E37E6B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E37E71: lea ecx, var_18
  loc_00E37E74: call ebx
  loc_00E37E76: mov eax, [esi]
  loc_00E37E78: push esi
  loc_00E37E79: call [eax+00000324h]
  loc_00E37E7F: lea ecx, var_18
  loc_00E37E82: push eax
  loc_00E37E83: push ecx
  loc_00E37E84: call edi
  loc_00E37E86: mov edx, [eax]
  loc_00E37E88: push eax
  loc_00E37E89: mov var_50, eax
  loc_00E37E8C: call [edx+00000204h]
  loc_00E37E92: test eax, eax
  loc_00E37E94: fnclex
  loc_00E37E96: jge 00E37EADh
  loc_00E37E98: mov ecx, var_50
  loc_00E37E9B: push 00000204h
  loc_00E37EA0: push 006DCB70h
  loc_00E37EA5: push ecx
  loc_00E37EA6: push eax
  loc_00E37EA7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37EAD: lea ecx, var_18
  loc_00E37EB0: call ebx
  loc_00E37EB2: mov edx, [esi]
  loc_00E37EB4: push esi
  loc_00E37EB5: call [edx+00000320h]
  loc_00E37EBB: push eax
  loc_00E37EBC: lea eax, var_18
  loc_00E37EBF: push eax
  loc_00E37EC0: call edi
  loc_00E37EC2: mov ecx, [eax]
  loc_00E37EC4: push 00000000h
  loc_00E37EC6: push eax
  loc_00E37EC7: mov var_50, eax
  loc_00E37ECA: call [ecx+0000009Ch]
  loc_00E37ED0: test eax, eax
  loc_00E37ED2: fnclex
  loc_00E37ED4: jge 00E37EEBh
  loc_00E37ED6: mov edx, var_50
  loc_00E37ED9: push 0000009Ch
  loc_00E37EDE: push 006E03D4h
  loc_00E37EE3: push edx
  loc_00E37EE4: push eax
  loc_00E37EE5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37EEB: lea ecx, var_18
  loc_00E37EEE: call ebx
  loc_00E37EF0: mov eax, [esi]
  loc_00E37EF2: push esi
  loc_00E37EF3: call [eax+00000314h]
  loc_00E37EF9: lea ecx, var_18
  loc_00E37EFC: push eax
  loc_00E37EFD: push ecx
  loc_00E37EFE: call edi
  loc_00E37F00: mov edx, [eax]
  loc_00E37F02: push 00000000h
  loc_00E37F04: push eax
  loc_00E37F05: mov var_50, eax
  loc_00E37F08: call [edx+0000009Ch]
  loc_00E37F0E: test eax, eax
  loc_00E37F10: fnclex
  loc_00E37F12: jge 00E37F29h
  loc_00E37F14: mov ecx, var_50
  loc_00E37F17: push 0000009Ch
  loc_00E37F1C: push 006E03D4h
  loc_00E37F21: push ecx
  loc_00E37F22: push eax
  loc_00E37F23: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37F29: lea ecx, var_18
  loc_00E37F2C: call ebx
  loc_00E37F2E: mov edx, [esi]
  loc_00E37F30: push esi
  loc_00E37F31: call [edx+00000328h]
  loc_00E37F37: push eax
  loc_00E37F38: lea eax, var_18
  loc_00E37F3B: push eax
  loc_00E37F3C: call edi
  loc_00E37F3E: mov ecx, [eax]
  loc_00E37F40: push 00000000h
  loc_00E37F42: push eax
  loc_00E37F43: mov var_50, eax
  loc_00E37F46: call [ecx+0000009Ch]
  loc_00E37F4C: test eax, eax
  loc_00E37F4E: fnclex
  loc_00E37F50: jge 00E37F67h
  loc_00E37F52: mov edx, var_50
  loc_00E37F55: push 0000009Ch
  loc_00E37F5A: push 006E03D4h
  loc_00E37F5F: push edx
  loc_00E37F60: push eax
  loc_00E37F61: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37F67: lea ecx, var_18
  loc_00E37F6A: call ebx
  loc_00E37F6C: mov eax, [esi]
  loc_00E37F6E: push esi
  loc_00E37F6F: call [eax+00000318h]
  loc_00E37F75: lea ecx, var_18
  loc_00E37F78: push eax
  loc_00E37F79: push ecx
  loc_00E37F7A: call edi
  loc_00E37F7C: mov edx, [eax]
  loc_00E37F7E: push 00000000h
  loc_00E37F80: push eax
  loc_00E37F81: mov var_50, eax
  loc_00E37F84: call [edx+0000009Ch]
  loc_00E37F8A: test eax, eax
  loc_00E37F8C: fnclex
  loc_00E37F8E: jge 00E37FA5h
  loc_00E37F90: mov ecx, var_50
  loc_00E37F93: push 0000009Ch
  loc_00E37F98: push 006E03D4h
  loc_00E37F9D: push ecx
  loc_00E37F9E: push eax
  loc_00E37F9F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37FA5: lea ecx, var_18
  loc_00E37FA8: call ebx
  loc_00E37FAA: mov edx, [esi]
  loc_00E37FAC: push esi
  loc_00E37FAD: call [edx+00000324h]
  loc_00E37FB3: push eax
  loc_00E37FB4: lea eax, var_18
  loc_00E37FB7: push eax
  loc_00E37FB8: call edi
  loc_00E37FBA: mov ecx, [eax]
  loc_00E37FBC: push 006DCC80h
  loc_00E37FC1: push eax
  loc_00E37FC2: mov var_50, eax
  loc_00E37FC5: call [ecx+000000A4h]
  loc_00E37FCB: test eax, eax
  loc_00E37FCD: fnclex
  loc_00E37FCF: jge 00E37FE6h
  loc_00E37FD1: mov edx, var_50
  loc_00E37FD4: push 000000A4h
  loc_00E37FD9: push 006DCB70h
  loc_00E37FDE: push edx
  loc_00E37FDF: push eax
  loc_00E37FE0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37FE6: lea ecx, var_18
  loc_00E37FE9: call ebx
  loc_00E37FEB: sub esp, 00000010h
  loc_00E37FEE: mov ecx, 00000008h
  loc_00E37FF3: mov edx, esp
  loc_00E37FF5: mov eax, 006E1738h ; "lc"
  loc_00E37FFA: push 0000000Eh
  loc_00E37FFC: push esi
  loc_00E37FFD: mov [edx], ecx
  loc_00E37FFF: mov ecx, var_38
  loc_00E38002: mov [edx+00000004h], ecx
  loc_00E38005: mov ecx, [esi]
  loc_00E38007: mov [edx+00000008h], eax
  loc_00E3800A: mov eax, var_30
  loc_00E3800D: mov [edx+0000000Ch], eax
  loc_00E38010: call [ecx+00000354h]
  loc_00E38016: lea edx, var_18
  loc_00E38019: push eax
  loc_00E3801A: push edx
  loc_00E3801B: call edi
  loc_00E3801D: push eax
  loc_00E3801E: call [00401238h] ; __vbaLateIdSt
  loc_00E38024: lea ecx, var_18
  loc_00E38027: call ebx
  loc_00E38029: mov eax, [esi]
  loc_00E3802B: push 00000000h
  loc_00E3802D: push 00000019h
  loc_00E3802F: push esi
  loc_00E38030: call [eax+00000354h]
  loc_00E38036: lea ecx, var_18
  loc_00E38039: push eax
  loc_00E3803A: push ecx
  loc_00E3803B: call edi
  loc_00E3803D: push eax
  loc_00E3803E: call [00401030h] ; __vbaLateIdCall
  loc_00E38044: add esp, 0000000Ch
  loc_00E38047: lea ecx, var_18
  loc_00E3804A: call ebx
  loc_00E3804C: mov edx, [esi]
  loc_00E3804E: push 006E05D8h
  loc_00E38053: push esi
  loc_00E38054: call [edx+00000354h]
  loc_00E3805A: push eax
  loc_00E3805B: lea eax, var_18
  loc_00E3805E: push eax
  loc_00E3805F: call edi
  loc_00E38061: push eax
  loc_00E38062: call [00401224h] ; __vbaCastObj
  loc_00E38068: lea ecx, var_2C
  loc_00E3806B: mov var_24, eax
  loc_00E3806E: push ecx
  loc_00E3806F: mov var_2C, 0000000Dh
  loc_00E38076: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3807C: mov ecx, [eax]
  loc_00E3807E: sub esp, 00000010h
  loc_00E38081: mov edx, esp
  loc_00E38083: push 00000000h
  loc_00E38085: push 0000002Ah
  loc_00E38087: mov [edx], ecx
  loc_00E38089: mov ecx, [eax+00000004h]
  loc_00E3808C: push esi
  loc_00E3808D: mov [edx+00000004h], ecx
  loc_00E38090: mov ecx, [eax+00000008h]
  loc_00E38093: mov eax, [eax+0000000Ch]
  loc_00E38096: mov [edx+00000008h], ecx
  loc_00E38099: mov ecx, [esi]
  loc_00E3809B: mov [edx+0000000Ch], eax
  loc_00E3809E: call [ecx+00000358h]
  loc_00E380A4: lea edx, var_1C
  loc_00E380A7: push eax
  loc_00E380A8: push edx
  loc_00E380A9: call edi
  loc_00E380AB: push eax
  loc_00E380AC: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E380B2: lea eax, var_1C
  loc_00E380B5: lea ecx, var_18
  loc_00E380B8: push eax
  loc_00E380B9: push ecx
  loc_00E380BA: push 00000002h
  loc_00E380BC: call [00401048h] ; __vbaFreeObjList
  loc_00E380C2: add esp, 00000028h
  loc_00E380C5: lea ecx, var_2C
  loc_00E380C8: call [00401024h] ; __vbaFreeVar
  loc_00E380CE: sub esp, 00000010h
  loc_00E380D1: mov ecx, 0000000Bh
  loc_00E380D6: mov edx, esp
  loc_00E380D8: xor eax, eax
  loc_00E380DA: push 00000006h
  loc_00E380DC: push esi
  loc_00E380DD: mov [edx], ecx
  loc_00E380DF: mov ecx, var_38
  loc_00E380E2: mov [edx+00000004h], ecx
  loc_00E380E5: mov ecx, [esi]
  loc_00E380E7: mov [edx+00000008h], eax
  loc_00E380EA: mov eax, var_30
  loc_00E380ED: mov [edx+0000000Ch], eax
  loc_00E380F0: call [ecx+00000358h]
  loc_00E380F6: lea edx, var_18
  loc_00E380F9: push eax
  loc_00E380FA: push edx
  loc_00E380FB: call edi
  loc_00E380FD: push eax
  loc_00E380FE: call [00401238h] ; __vbaLateIdSt
  loc_00E38104: lea ecx, var_18
  loc_00E38107: call ebx
  loc_00E38109: sub esp, 00000010h
  loc_00E3810C: mov ecx, 0000000Bh
  loc_00E38111: mov edx, esp
  loc_00E38113: xor eax, eax
  loc_00E38115: push 8001000Eh
  loc_00E3811A: push esi
  loc_00E3811B: mov [edx], ecx
  loc_00E3811D: mov ecx, var_38
  loc_00E38120: mov [edx+00000004h], ecx
  loc_00E38123: mov ecx, [esi]
  loc_00E38125: mov [edx+00000008h], eax
  loc_00E38128: mov eax, var_30
  loc_00E3812B: mov [edx+0000000Ch], eax
  loc_00E3812E: call [ecx+00000358h]
  loc_00E38134: lea edx, var_18
  loc_00E38137: push eax
  loc_00E38138: push edx
  loc_00E38139: call edi
  loc_00E3813B: push eax
  loc_00E3813C: call [00401238h] ; __vbaLateIdSt
  loc_00E38142: lea ecx, var_18
  loc_00E38145: call ebx
  loc_00E38147: mov eax, [esi]
  loc_00E38149: push 00000000h
  loc_00E3814B: push FFFFFDDAh
  loc_00E38150: push esi
  loc_00E38151: call [eax+00000358h]
  loc_00E38157: lea ecx, var_18
  loc_00E3815A: push eax
  loc_00E3815B: push ecx
  loc_00E3815C: call edi
  loc_00E3815E: push eax
  loc_00E3815F: call [00401030h] ; __vbaLateIdCall
  loc_00E38165: add esp, 0000000Ch
  loc_00E38168: lea ecx, var_18
  loc_00E3816B: call ebx
  loc_00E3816D: mov var_4, 00000000h
  loc_00E38174: push 00E38199h
  loc_00E38179: jmp 00E38198h
  loc_00E3817B: lea edx, var_1C
  loc_00E3817E: lea eax, var_18
  loc_00E38181: push edx
  loc_00E38182: push eax
  loc_00E38183: push 00000002h
  loc_00E38185: call [00401048h] ; __vbaFreeObjList
  loc_00E3818B: add esp, 0000000Ch
  loc_00E3818E: lea ecx, var_2C
  loc_00E38191: call [00401024h] ; __vbaFreeVar
  loc_00E38197: ret
  loc_00E38198: ret
  loc_00E38199: mov eax, Me
  loc_00E3819C: push eax
  loc_00E3819D: mov ecx, [eax]
  loc_00E3819F: call [ecx+00000008h]
  loc_00E381A2: mov eax, var_4
  loc_00E381A5: mov ecx, var_14
  loc_00E381A8: pop edi
  loc_00E381A9: pop esi
  loc_00E381AA: mov fs:[00000000h], ecx
  loc_00E381B1: pop ebx
  loc_00E381B2: mov esp, ebp
  loc_00E381B4: pop ebp
  loc_00E381B5: retn 0004h
End Sub

Private Sub back_Click() 'E360B0
  loc_00E360B0: push ebp
  loc_00E360B1: mov ebp, esp
  loc_00E360B3: sub esp, 0000000Ch
  loc_00E360B6: push 00402836h ; __vbaExceptHandler
  loc_00E360BB: mov eax, fs:[00000000h]
  loc_00E360C1: push eax
  loc_00E360C2: mov fs:[00000000h], esp
  loc_00E360C9: sub esp, 0000003Ch
  loc_00E360CC: push ebx
  loc_00E360CD: push esi
  loc_00E360CE: push edi
  loc_00E360CF: mov var_C, esp
  loc_00E360D2: mov var_8, 004026D0h
  loc_00E360D9: mov eax, Me
  loc_00E360DC: mov ecx, eax
  loc_00E360DE: and ecx, 00000001h
  loc_00E360E1: mov var_4, ecx
  loc_00E360E4: and al, FEh
  loc_00E360E6: push eax
  loc_00E360E7: mov Me, eax
  loc_00E360EA: mov edx, [eax]
  loc_00E360EC: call [edx+00000004h]
  loc_00E360EF: mov eax, [00E3D024h]
  loc_00E360F4: mov var_18, 00000000h
  loc_00E360FB: test eax, eax
  loc_00E360FD: jnz 00E3610Fh
  loc_00E360FF: push 00E3D024h
  loc_00E36104: push 006CE120h
  loc_00E36109: call [004011A0h] ; __vbaNew2
  loc_00E3610F: sub esp, 00000010h
  loc_00E36112: mov ecx, 0000000Ah
  loc_00E36117: mov esi, esp
  loc_00E36119: mov var_28, ecx
  loc_00E3611C: mov eax, 80020004h
  loc_00E36121: mov ebx, [00E3D024h]
  loc_00E36127: mov [esi], ecx
  loc_00E36129: mov ecx, var_34
  loc_00E3612C: mov edx, eax
  loc_00E3612E: sub esp, 00000010h
  loc_00E36131: mov [esi+00000004h], ecx
  loc_00E36134: mov edi, [ebx]
  loc_00E36136: mov ecx, esp
  loc_00E36138: mov var_4C, edi
  loc_00E3613B: mov [esi+00000008h], eax
  loc_00E3613E: mov eax, var_2C
  loc_00E36141: mov edi, var_1C
  loc_00E36144: push ebx
  loc_00E36145: mov [esi+0000000Ch], eax
  loc_00E36148: mov eax, var_28
  loc_00E3614B: mov esi, var_24
  loc_00E3614E: mov [ecx], eax
  loc_00E36150: mov [ecx+00000004h], esi
  loc_00E36153: mov [ecx+00000008h], edx
  loc_00E36156: mov [ecx+0000000Ch], edi
  loc_00E36159: mov ecx, var_4C
  loc_00E3615C: call [ecx+000002B0h]
  loc_00E36162: test eax, eax
  loc_00E36164: fnclex
  loc_00E36166: jge 00E3617Ah
  loc_00E36168: push 000002B0h
  loc_00E3616D: push 006DC970h
  loc_00E36172: push ebx
  loc_00E36173: push eax
  loc_00E36174: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3617A: mov eax, [00E3D9CCh]
  loc_00E3617F: test eax, eax
  loc_00E36181: jnz 00E36193h
  loc_00E36183: push 00E3D9CCh
  loc_00E36188: push 006DC960h
  loc_00E3618D: call [004011A0h] ; __vbaNew2
  loc_00E36193: mov eax, [00E3D9CCh]
  loc_00E36198: mov edx, Me
  loc_00E3619B: mov var_3C, eax
  loc_00E3619E: push edx
  loc_00E3619F: mov ebx, [eax]
  loc_00E361A1: lea eax, var_18
  loc_00E361A4: push eax
  loc_00E361A5: call [004010B4h] ; __vbaObjSetAddref
  loc_00E361AB: mov ecx, ebx
  loc_00E361AD: mov ebx, var_3C
  loc_00E361B0: push eax
  loc_00E361B1: push ebx
  loc_00E361B2: call [ecx+00000010h]
  loc_00E361B5: test eax, eax
  loc_00E361B7: fnclex
  loc_00E361B9: jge 00E361CAh
  loc_00E361BB: push 00000010h
  loc_00E361BD: push 006DC950h
  loc_00E361C2: push ebx
  loc_00E361C3: push eax
  loc_00E361C4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E361CA: lea ecx, var_18
  loc_00E361CD: call [00401254h] ; __vbaFreeObj
  loc_00E361D3: mov eax, [00E3D024h]
  loc_00E361D8: test eax, eax
  loc_00E361DA: jnz 00E361ECh
  loc_00E361DC: push 00E3D024h
  loc_00E361E1: push 006CE120h
  loc_00E361E6: call [004011A0h] ; __vbaNew2
  loc_00E361EC: sub esp, 00000010h
  loc_00E361EF: mov eax, 0000000Ah
  loc_00E361F4: mov ecx, esp
  loc_00E361F6: mov var_28, eax
  loc_00E361F9: sub esp, 00000010h
  loc_00E361FC: mov ebx, [00E3D024h]
  loc_00E36202: mov [ecx], eax
  loc_00E36204: mov eax, var_34
  loc_00E36207: mov edx, [ebx]
  loc_00E36209: mov var_20, 80020004h
  loc_00E36210: mov [ecx+00000004h], eax
  loc_00E36213: mov eax, 80020004h
  loc_00E36218: mov [ecx+00000008h], eax
  loc_00E3621B: mov eax, var_2C
  loc_00E3621E: mov [ecx+0000000Ch], eax
  loc_00E36221: mov eax, var_28
  loc_00E36224: mov ecx, esp
  loc_00E36226: push ebx
  loc_00E36227: mov [ecx], eax
  loc_00E36229: mov eax, var_20
  loc_00E3622C: mov [ecx+00000004h], esi
  loc_00E3622F: mov [ecx+00000008h], eax
  loc_00E36232: mov [ecx+0000000Ch], edi
  loc_00E36235: call [edx+000002B0h]
  loc_00E3623B: test eax, eax
  loc_00E3623D: fnclex
  loc_00E3623F: jge 00E36253h
  loc_00E36241: push 000002B0h
  loc_00E36246: push 006DC970h
  loc_00E3624B: push ebx
  loc_00E3624C: push eax
  loc_00E3624D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36253: mov eax, [00E3D024h]
  loc_00E36258: or ebx, FFFFFFFFh
  loc_00E3625B: test eax, eax
  loc_00E3625D: jnz 00E36274h
  loc_00E3625F: push 00E3D024h
  loc_00E36264: push 006CE120h
  loc_00E36269: call [004011A0h] ; __vbaNew2
  loc_00E3626F: mov eax, [00E3D024h]
  loc_00E36274: sub esp, 00000010h
  loc_00E36277: mov ecx, 0000000Bh
  loc_00E3627C: mov edx, esp
  loc_00E3627E: push 80010007h
  loc_00E36283: push eax
  loc_00E36284: mov [edx], ecx
  loc_00E36286: mov ecx, [eax]
  loc_00E36288: mov [edx+00000004h], esi
  loc_00E3628B: mov [edx+00000008h], ebx
  loc_00E3628E: mov [edx+0000000Ch], edi
  loc_00E36291: call [ecx+00000318h]
  loc_00E36297: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E3629D: lea edx, var_18
  loc_00E362A0: push eax
  loc_00E362A1: push edx
  loc_00E362A2: call ebx
  loc_00E362A4: push eax
  loc_00E362A5: call [00401238h] ; __vbaLateIdSt
  loc_00E362AB: lea ecx, var_18
  loc_00E362AE: call [00401254h] ; __vbaFreeObj
  loc_00E362B4: mov eax, [00E3D024h]
  loc_00E362B9: test eax, eax
  loc_00E362BB: jnz 00E362D2h
  loc_00E362BD: push 00E3D024h
  loc_00E362C2: push 006CE120h
  loc_00E362C7: call [004011A0h] ; __vbaNew2
  loc_00E362CD: mov eax, [00E3D024h]
  loc_00E362D2: sub esp, 00000010h
  loc_00E362D5: mov ecx, 0000000Bh
  loc_00E362DA: mov edx, esp
  loc_00E362DC: push 80010007h
  loc_00E362E1: push eax
  loc_00E362E2: mov [edx], ecx
  loc_00E362E4: or ecx, FFFFFFFFh
  loc_00E362E7: mov [edx+00000004h], esi
  loc_00E362EA: mov [edx+00000008h], ecx
  loc_00E362ED: mov ecx, [eax]
  loc_00E362EF: mov [edx+0000000Ch], edi
  loc_00E362F2: call [ecx+0000031Ch]
  loc_00E362F8: lea edx, var_18
  loc_00E362FB: push eax
  loc_00E362FC: push edx
  loc_00E362FD: call ebx
  loc_00E362FF: push eax
  loc_00E36300: call [00401238h] ; __vbaLateIdSt
  loc_00E36306: lea ecx, var_18
  loc_00E36309: call [00401254h] ; __vbaFreeObj
  loc_00E3630F: mov eax, [00E3D024h]
  loc_00E36314: mov var_20, FFFFFFFFh
  loc_00E3631B: test eax, eax
  loc_00E3631D: jnz 00E36334h
  loc_00E3631F: push 00E3D024h
  loc_00E36324: push 006CE120h
  loc_00E36329: call [004011A0h] ; __vbaNew2
  loc_00E3632F: mov eax, [00E3D024h]
  loc_00E36334: sub esp, 00000010h
  loc_00E36337: mov ecx, 0000000Bh
  loc_00E3633C: mov edx, esp
  loc_00E3633E: push 80010007h
  loc_00E36343: push eax
  loc_00E36344: mov [edx], ecx
  loc_00E36346: mov ecx, var_20
  loc_00E36349: mov [edx+00000004h], esi
  loc_00E3634C: mov [edx+00000008h], ecx
  loc_00E3634F: mov [edx+0000000Ch], edi
  loc_00E36352: mov edx, [eax]
  loc_00E36354: call [edx+00000320h]
  loc_00E3635A: push eax
  loc_00E3635B: lea eax, var_18
  loc_00E3635E: push eax
  loc_00E3635F: call ebx
  loc_00E36361: push eax
  loc_00E36362: call [00401238h] ; __vbaLateIdSt
  loc_00E36368: lea ecx, var_18
  loc_00E3636B: call [00401254h] ; __vbaFreeObj
  loc_00E36371: mov eax, [00E3D024h]
  loc_00E36376: test eax, eax
  loc_00E36378: jnz 00E3638Fh
  loc_00E3637A: push 00E3D024h
  loc_00E3637F: push 006CE120h
  loc_00E36384: call [004011A0h] ; __vbaNew2
  loc_00E3638A: mov eax, [00E3D024h]
  loc_00E3638F: sub esp, 00000010h
  loc_00E36392: mov ecx, 00000008h
  loc_00E36397: mov edx, esp
  loc_00E36399: push FFFFFDFAh
  loc_00E3639E: push eax
  loc_00E3639F: mov [edx], ecx
  loc_00E363A1: mov ecx, 006DCDB4h ; " A D M I N "
  loc_00E363A6: mov [edx+00000004h], esi
  loc_00E363A9: mov [edx+00000008h], ecx
  loc_00E363AC: mov ecx, [eax]
  loc_00E363AE: mov [edx+0000000Ch], edi
  loc_00E363B1: call [ecx+00000324h]
  loc_00E363B7: lea edx, var_18
  loc_00E363BA: push eax
  loc_00E363BB: push edx
  loc_00E363BC: call ebx
  loc_00E363BE: push eax
  loc_00E363BF: call [00401238h] ; __vbaLateIdSt
  loc_00E363C5: lea ecx, var_18
  loc_00E363C8: call [00401254h] ; __vbaFreeObj
  loc_00E363CE: mov eax, [00E3D024h]
  loc_00E363D3: mov var_28, 0000000Bh
  loc_00E363DA: test eax, eax
  loc_00E363DC: jnz 00E363F3h
  loc_00E363DE: push 00E3D024h
  loc_00E363E3: push 006CE120h
  loc_00E363E8: call [004011A0h] ; __vbaNew2
  loc_00E363EE: mov eax, [00E3D024h]
  loc_00E363F3: mov ecx, var_28
  loc_00E363F6: sub esp, 00000010h
  loc_00E363F9: mov edx, esp
  loc_00E363FB: push 8001000Dh
  loc_00E36400: push eax
  loc_00E36401: mov [edx], ecx
  loc_00E36403: xor ecx, ecx
  loc_00E36405: mov [edx+00000004h], esi
  loc_00E36408: mov [edx+00000008h], ecx
  loc_00E3640B: mov [edx+0000000Ch], edi
  loc_00E3640E: mov edx, [eax]
  loc_00E36410: call [edx+00000324h]
  loc_00E36416: push eax
  loc_00E36417: lea eax, var_18
  loc_00E3641A: push eax
  loc_00E3641B: call ebx
  loc_00E3641D: push eax
  loc_00E3641E: call [00401238h] ; __vbaLateIdSt
  loc_00E36424: lea ecx, var_18
  loc_00E36427: call [00401254h] ; __vbaFreeObj
  loc_00E3642D: mov eax, [00E3D024h]
  loc_00E36432: test eax, eax
  loc_00E36434: jnz 00E3644Bh
  loc_00E36436: push 00E3D024h
  loc_00E3643B: push 006CE120h
  loc_00E36440: call [004011A0h] ; __vbaNew2
  loc_00E36446: mov eax, [00E3D024h]
  loc_00E3644B: sub esp, 00000010h
  loc_00E3644E: mov ecx, 00000003h
  loc_00E36453: mov edx, esp
  loc_00E36455: push FFFFFE0Bh
  loc_00E3645A: push eax
  loc_00E3645B: mov [edx], ecx
  loc_00E3645D: mov ecx, 00404040h
  loc_00E36462: mov [edx+00000004h], esi
  loc_00E36465: mov [edx+00000008h], ecx
  loc_00E36468: mov ecx, [eax]
  loc_00E3646A: mov [edx+0000000Ch], edi
  loc_00E3646D: call [ecx+00000324h]
  loc_00E36473: lea edx, var_18
  loc_00E36476: push eax
  loc_00E36477: push edx
  loc_00E36478: call ebx
  loc_00E3647A: push eax
  loc_00E3647B: call [00401238h] ; __vbaLateIdSt
  loc_00E36481: lea ecx, var_18
  loc_00E36484: call [00401254h] ; __vbaFreeObj
  loc_00E3648A: mov eax, [00E3D024h]
  loc_00E3648F: mov var_28, 00000003h
  loc_00E36496: test eax, eax
  loc_00E36498: jnz 00E364AFh
  loc_00E3649A: push 00E3D024h
  loc_00E3649F: push 006CE120h
  loc_00E364A4: call [004011A0h] ; __vbaNew2
  loc_00E364AA: mov eax, [00E3D024h]
  loc_00E364AF: mov ecx, var_28
  loc_00E364B2: sub esp, 00000010h
  loc_00E364B5: mov edx, esp
  loc_00E364B7: push FFFFFDFFh
  loc_00E364BC: push eax
  loc_00E364BD: mov [edx], ecx
  loc_00E364BF: mov ecx, 00E0E0E0h
  loc_00E364C4: mov [edx+00000004h], esi
  loc_00E364C7: mov [edx+00000008h], ecx
  loc_00E364CA: mov [edx+0000000Ch], edi
  loc_00E364CD: mov edx, [eax]
  loc_00E364CF: call [edx+00000324h]
  loc_00E364D5: push eax
  loc_00E364D6: lea eax, var_18
  loc_00E364D9: push eax
  loc_00E364DA: call ebx
  loc_00E364DC: push eax
  loc_00E364DD: call [00401238h] ; __vbaLateIdSt
  loc_00E364E3: lea ecx, var_18
  loc_00E364E6: call [00401254h] ; __vbaFreeObj
  loc_00E364EC: mov var_4, 00000000h
  loc_00E364F3: push 00E36505h
  loc_00E364F8: jmp 00E36504h
  loc_00E364FA: lea ecx, var_18
  loc_00E364FD: call [00401254h] ; __vbaFreeObj
  loc_00E36503: ret
  loc_00E36504: ret
  loc_00E36505: mov eax, Me
  loc_00E36508: push eax
  loc_00E36509: mov ecx, [eax]
  loc_00E3650B: call [ecx+00000008h]
  loc_00E3650E: mov eax, var_4
  loc_00E36511: mov ecx, var_14
  loc_00E36514: pop edi
  loc_00E36515: pop esi
  loc_00E36516: mov fs:[00000000h], ecx
  loc_00E3651D: pop ebx
  loc_00E3651E: mov esp, ebp
  loc_00E36520: pop ebp
  loc_00E36521: retn 0004h
End Sub

Private Sub optagama_Click() 'E37620
  loc_00E37620: push ebp
  loc_00E37621: mov ebp, esp
  loc_00E37623: sub esp, 0000000Ch
  loc_00E37626: push 00402836h ; __vbaExceptHandler
  loc_00E3762B: mov eax, fs:[00000000h]
  loc_00E37631: push eax
  loc_00E37632: mov fs:[00000000h], esp
  loc_00E37639: sub esp, 00000048h
  loc_00E3763C: push ebx
  loc_00E3763D: push esi
  loc_00E3763E: push edi
  loc_00E3763F: mov var_C, esp
  loc_00E37642: mov var_8, 00402720h
  loc_00E37649: mov esi, Me
  loc_00E3764C: mov eax, esi
  loc_00E3764E: and eax, 00000001h
  loc_00E37651: mov var_4, eax
  loc_00E37654: and esi, FFFFFFFEh
  loc_00E37657: push esi
  loc_00E37658: mov Me, esi
  loc_00E3765B: mov ecx, [esi]
  loc_00E3765D: call [ecx+00000004h]
  loc_00E37660: sub esp, 00000010h
  loc_00E37663: mov ecx, 0000000Bh
  loc_00E37668: mov edx, esp
  loc_00E3766A: xor eax, eax
  loc_00E3766C: mov var_18, eax
  loc_00E3766F: mov var_1C, eax
  loc_00E37672: mov [edx], ecx
  loc_00E37674: mov ecx, var_38
  loc_00E37677: mov var_2C, eax
  loc_00E3767A: or eax, FFFFFFFFh
  loc_00E3767D: mov [edx+00000004h], ecx
  loc_00E37680: mov ecx, [esi]
  loc_00E37682: push 80010007h
  loc_00E37687: push esi
  loc_00E37688: mov [edx+00000008h], eax
  loc_00E3768B: mov eax, var_30
  loc_00E3768E: mov [edx+0000000Ch], eax
  loc_00E37691: call [ecx+0000032Ch]
  loc_00E37697: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3769D: lea edx, var_18
  loc_00E376A0: push eax
  loc_00E376A1: push edx
  loc_00E376A2: call edi
  loc_00E376A4: push eax
  loc_00E376A5: call [00401238h] ; __vbaLateIdSt
  loc_00E376AB: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E376B1: lea ecx, var_18
  loc_00E376B4: call ebx
  loc_00E376B6: mov eax, [esi]
  loc_00E376B8: push esi
  loc_00E376B9: call [eax+00000324h]
  loc_00E376BF: lea ecx, var_18
  loc_00E376C2: push eax
  loc_00E376C3: push ecx
  loc_00E376C4: call edi
  loc_00E376C6: mov edx, [eax]
  loc_00E376C8: push eax
  loc_00E376C9: mov var_50, eax
  loc_00E376CC: call [edx+00000204h]
  loc_00E376D2: test eax, eax
  loc_00E376D4: fnclex
  loc_00E376D6: jge 00E376EDh
  loc_00E376D8: mov ecx, var_50
  loc_00E376DB: push 00000204h
  loc_00E376E0: push 006DCB70h
  loc_00E376E5: push ecx
  loc_00E376E6: push eax
  loc_00E376E7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E376ED: lea ecx, var_18
  loc_00E376F0: call ebx
  loc_00E376F2: mov edx, [esi]
  loc_00E376F4: push esi
  loc_00E376F5: call [edx+00000320h]
  loc_00E376FB: push eax
  loc_00E376FC: lea eax, var_18
  loc_00E376FF: push eax
  loc_00E37700: call edi
  loc_00E37702: mov ecx, [eax]
  loc_00E37704: push 00000000h
  loc_00E37706: push eax
  loc_00E37707: mov var_50, eax
  loc_00E3770A: call [ecx+0000009Ch]
  loc_00E37710: test eax, eax
  loc_00E37712: fnclex
  loc_00E37714: jge 00E3772Bh
  loc_00E37716: mov edx, var_50
  loc_00E37719: push 0000009Ch
  loc_00E3771E: push 006E03D4h
  loc_00E37723: push edx
  loc_00E37724: push eax
  loc_00E37725: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3772B: lea ecx, var_18
  loc_00E3772E: call ebx
  loc_00E37730: mov eax, [esi]
  loc_00E37732: push esi
  loc_00E37733: call [eax+0000031Ch]
  loc_00E37739: lea ecx, var_18
  loc_00E3773C: push eax
  loc_00E3773D: push ecx
  loc_00E3773E: call edi
  loc_00E37740: mov edx, [eax]
  loc_00E37742: push 00000000h
  loc_00E37744: push eax
  loc_00E37745: mov var_50, eax
  loc_00E37748: call [edx+0000009Ch]
  loc_00E3774E: test eax, eax
  loc_00E37750: fnclex
  loc_00E37752: jge 00E37769h
  loc_00E37754: mov ecx, var_50
  loc_00E37757: push 0000009Ch
  loc_00E3775C: push 006E03D4h
  loc_00E37761: push ecx
  loc_00E37762: push eax
  loc_00E37763: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37769: lea ecx, var_18
  loc_00E3776C: call ebx
  loc_00E3776E: mov edx, [esi]
  loc_00E37770: push esi
  loc_00E37771: call [edx+00000318h]
  loc_00E37777: push eax
  loc_00E37778: lea eax, var_18
  loc_00E3777B: push eax
  loc_00E3777C: call edi
  loc_00E3777E: mov ecx, [eax]
  loc_00E37780: push 00000000h
  loc_00E37782: push eax
  loc_00E37783: mov var_50, eax
  loc_00E37786: call [ecx+0000009Ch]
  loc_00E3778C: test eax, eax
  loc_00E3778E: fnclex
  loc_00E37790: jge 00E377A7h
  loc_00E37792: mov edx, var_50
  loc_00E37795: push 0000009Ch
  loc_00E3779A: push 006E03D4h
  loc_00E3779F: push edx
  loc_00E377A0: push eax
  loc_00E377A1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E377A7: lea ecx, var_18
  loc_00E377AA: call ebx
  loc_00E377AC: mov eax, [esi]
  loc_00E377AE: push esi
  loc_00E377AF: call [eax+00000324h]
  loc_00E377B5: lea ecx, var_18
  loc_00E377B8: push eax
  loc_00E377B9: push ecx
  loc_00E377BA: call edi
  loc_00E377BC: mov edx, [eax]
  loc_00E377BE: push 006DCC80h
  loc_00E377C3: push eax
  loc_00E377C4: mov var_50, eax
  loc_00E377C7: call [edx+000000A4h]
  loc_00E377CD: test eax, eax
  loc_00E377CF: fnclex
  loc_00E377D1: jge 00E377E8h
  loc_00E377D3: mov ecx, var_50
  loc_00E377D6: push 000000A4h
  loc_00E377DB: push 006DCB70h
  loc_00E377E0: push ecx
  loc_00E377E1: push eax
  loc_00E377E2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E377E8: lea ecx, var_18
  loc_00E377EB: call ebx
  loc_00E377ED: mov edx, [esi]
  loc_00E377EF: push esi
  loc_00E377F0: call [edx+00000314h]
  loc_00E377F6: push eax
  loc_00E377F7: lea eax, var_18
  loc_00E377FA: push eax
  loc_00E377FB: call edi
  loc_00E377FD: mov ecx, [eax]
  loc_00E377FF: push 00000000h
  loc_00E37801: push eax
  loc_00E37802: mov var_50, eax
  loc_00E37805: call [ecx+0000009Ch]
  loc_00E3780B: test eax, eax
  loc_00E3780D: fnclex
  loc_00E3780F: jge 00E37826h
  loc_00E37811: mov edx, var_50
  loc_00E37814: push 0000009Ch
  loc_00E37819: push 006E03D4h
  loc_00E3781E: push edx
  loc_00E3781F: push eax
  loc_00E37820: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37826: lea ecx, var_18
  loc_00E37829: call ebx
  loc_00E3782B: sub esp, 00000010h
  loc_00E3782E: mov ecx, 00000008h
  loc_00E37833: mov edx, esp
  loc_00E37835: mov eax, 006E1738h ; "lc"
  loc_00E3783A: push 0000000Eh
  loc_00E3783C: push esi
  loc_00E3783D: mov [edx], ecx
  loc_00E3783F: mov ecx, var_38
  loc_00E37842: mov [edx+00000004h], ecx
  loc_00E37845: mov ecx, [esi]
  loc_00E37847: mov [edx+00000008h], eax
  loc_00E3784A: mov eax, var_30
  loc_00E3784D: mov [edx+0000000Ch], eax
  loc_00E37850: call [ecx+00000354h]
  loc_00E37856: lea edx, var_18
  loc_00E37859: push eax
  loc_00E3785A: push edx
  loc_00E3785B: call edi
  loc_00E3785D: push eax
  loc_00E3785E: call [00401238h] ; __vbaLateIdSt
  loc_00E37864: lea ecx, var_18
  loc_00E37867: call ebx
  loc_00E37869: mov eax, [esi]
  loc_00E3786B: push 00000000h
  loc_00E3786D: push 00000019h
  loc_00E3786F: push esi
  loc_00E37870: call [eax+00000354h]
  loc_00E37876: lea ecx, var_18
  loc_00E37879: push eax
  loc_00E3787A: push ecx
  loc_00E3787B: call edi
  loc_00E3787D: push eax
  loc_00E3787E: call [00401030h] ; __vbaLateIdCall
  loc_00E37884: add esp, 0000000Ch
  loc_00E37887: lea ecx, var_18
  loc_00E3788A: call ebx
  loc_00E3788C: mov edx, [esi]
  loc_00E3788E: push 006E05D8h
  loc_00E37893: push esi
  loc_00E37894: call [edx+00000354h]
  loc_00E3789A: push eax
  loc_00E3789B: lea eax, var_18
  loc_00E3789E: push eax
  loc_00E3789F: call edi
  loc_00E378A1: push eax
  loc_00E378A2: call [00401224h] ; __vbaCastObj
  loc_00E378A8: lea ecx, var_2C
  loc_00E378AB: mov var_24, eax
  loc_00E378AE: push ecx
  loc_00E378AF: mov var_2C, 0000000Dh
  loc_00E378B6: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E378BC: mov ecx, [eax]
  loc_00E378BE: sub esp, 00000010h
  loc_00E378C1: mov edx, esp
  loc_00E378C3: push 00000000h
  loc_00E378C5: push 0000002Ah
  loc_00E378C7: mov [edx], ecx
  loc_00E378C9: mov ecx, [eax+00000004h]
  loc_00E378CC: push esi
  loc_00E378CD: mov [edx+00000004h], ecx
  loc_00E378D0: mov ecx, [eax+00000008h]
  loc_00E378D3: mov eax, [eax+0000000Ch]
  loc_00E378D6: mov [edx+00000008h], ecx
  loc_00E378D9: mov ecx, [esi]
  loc_00E378DB: mov [edx+0000000Ch], eax
  loc_00E378DE: call [ecx+00000358h]
  loc_00E378E4: lea edx, var_1C
  loc_00E378E7: push eax
  loc_00E378E8: push edx
  loc_00E378E9: call edi
  loc_00E378EB: push eax
  loc_00E378EC: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E378F2: lea eax, var_1C
  loc_00E378F5: lea ecx, var_18
  loc_00E378F8: push eax
  loc_00E378F9: push ecx
  loc_00E378FA: push 00000002h
  loc_00E378FC: call [00401048h] ; __vbaFreeObjList
  loc_00E37902: add esp, 00000028h
  loc_00E37905: lea ecx, var_2C
  loc_00E37908: call [00401024h] ; __vbaFreeVar
  loc_00E3790E: sub esp, 00000010h
  loc_00E37911: mov ecx, 0000000Bh
  loc_00E37916: mov edx, esp
  loc_00E37918: xor eax, eax
  loc_00E3791A: push 00000006h
  loc_00E3791C: push esi
  loc_00E3791D: mov [edx], ecx
  loc_00E3791F: mov ecx, var_38
  loc_00E37922: mov [edx+00000004h], ecx
  loc_00E37925: mov ecx, [esi]
  loc_00E37927: mov [edx+00000008h], eax
  loc_00E3792A: mov eax, var_30
  loc_00E3792D: mov [edx+0000000Ch], eax
  loc_00E37930: call [ecx+00000358h]
  loc_00E37936: lea edx, var_18
  loc_00E37939: push eax
  loc_00E3793A: push edx
  loc_00E3793B: call edi
  loc_00E3793D: push eax
  loc_00E3793E: call [00401238h] ; __vbaLateIdSt
  loc_00E37944: lea ecx, var_18
  loc_00E37947: call ebx
  loc_00E37949: sub esp, 00000010h
  loc_00E3794C: mov ecx, 0000000Bh
  loc_00E37951: mov edx, esp
  loc_00E37953: xor eax, eax
  loc_00E37955: push 8001000Eh
  loc_00E3795A: push esi
  loc_00E3795B: mov [edx], ecx
  loc_00E3795D: mov ecx, var_38
  loc_00E37960: mov [edx+00000004h], ecx
  loc_00E37963: mov ecx, [esi]
  loc_00E37965: mov [edx+00000008h], eax
  loc_00E37968: mov eax, var_30
  loc_00E3796B: mov [edx+0000000Ch], eax
  loc_00E3796E: call [ecx+00000358h]
  loc_00E37974: lea edx, var_18
  loc_00E37977: push eax
  loc_00E37978: push edx
  loc_00E37979: call edi
  loc_00E3797B: push eax
  loc_00E3797C: call [00401238h] ; __vbaLateIdSt
  loc_00E37982: lea ecx, var_18
  loc_00E37985: call ebx
  loc_00E37987: mov eax, [esi]
  loc_00E37989: push 00000000h
  loc_00E3798B: push FFFFFDDAh
  loc_00E37990: push esi
  loc_00E37991: call [eax+00000358h]
  loc_00E37997: lea ecx, var_18
  loc_00E3799A: push eax
  loc_00E3799B: push ecx
  loc_00E3799C: call edi
  loc_00E3799E: push eax
  loc_00E3799F: call [00401030h] ; __vbaLateIdCall
  loc_00E379A5: add esp, 0000000Ch
  loc_00E379A8: lea ecx, var_18
  loc_00E379AB: call ebx
  loc_00E379AD: mov var_4, 00000000h
  loc_00E379B4: push 00E379D9h
  loc_00E379B9: jmp 00E379D8h
  loc_00E379BB: lea edx, var_1C
  loc_00E379BE: lea eax, var_18
  loc_00E379C1: push edx
  loc_00E379C2: push eax
  loc_00E379C3: push 00000002h
  loc_00E379C5: call [00401048h] ; __vbaFreeObjList
  loc_00E379CB: add esp, 0000000Ch
  loc_00E379CE: lea ecx, var_2C
  loc_00E379D1: call [00401024h] ; __vbaFreeVar
  loc_00E379D7: ret
  loc_00E379D8: ret
  loc_00E379D9: mov eax, Me
  loc_00E379DC: push eax
  loc_00E379DD: mov ecx, [eax]
  loc_00E379DF: call [ecx+00000008h]
  loc_00E379E2: mov eax, var_4
  loc_00E379E5: mov ecx, var_14
  loc_00E379E8: pop edi
  loc_00E379E9: pop esi
  loc_00E379EA: mov fs:[00000000h], ecx
  loc_00E379F1: pop ebx
  loc_00E379F2: mov esp, ebp
  loc_00E379F4: pop ebp
  loc_00E379F5: retn 0004h
End Sub

Private Sub optjur_Click() 'E37A00
  loc_00E37A00: push ebp
  loc_00E37A01: mov ebp, esp
  loc_00E37A03: sub esp, 0000000Ch
  loc_00E37A06: push 00402836h ; __vbaExceptHandler
  loc_00E37A0B: mov eax, fs:[00000000h]
  loc_00E37A11: push eax
  loc_00E37A12: mov fs:[00000000h], esp
  loc_00E37A19: sub esp, 00000048h
  loc_00E37A1C: push ebx
  loc_00E37A1D: push esi
  loc_00E37A1E: push edi
  loc_00E37A1F: mov var_C, esp
  loc_00E37A22: mov var_8, 00402730h
  loc_00E37A29: mov esi, Me
  loc_00E37A2C: mov eax, esi
  loc_00E37A2E: and eax, 00000001h
  loc_00E37A31: mov var_4, eax
  loc_00E37A34: and esi, FFFFFFFEh
  loc_00E37A37: push esi
  loc_00E37A38: mov Me, esi
  loc_00E37A3B: mov ecx, [esi]
  loc_00E37A3D: call [ecx+00000004h]
  loc_00E37A40: sub esp, 00000010h
  loc_00E37A43: mov ecx, 0000000Bh
  loc_00E37A48: mov edx, esp
  loc_00E37A4A: xor eax, eax
  loc_00E37A4C: mov var_18, eax
  loc_00E37A4F: mov var_1C, eax
  loc_00E37A52: mov [edx], ecx
  loc_00E37A54: mov ecx, var_38
  loc_00E37A57: mov var_2C, eax
  loc_00E37A5A: or eax, FFFFFFFFh
  loc_00E37A5D: mov [edx+00000004h], ecx
  loc_00E37A60: mov ecx, [esi]
  loc_00E37A62: push 80010007h
  loc_00E37A67: push esi
  loc_00E37A68: mov [edx+00000008h], eax
  loc_00E37A6B: mov eax, var_30
  loc_00E37A6E: mov [edx+0000000Ch], eax
  loc_00E37A71: call [ecx+0000032Ch]
  loc_00E37A77: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E37A7D: lea edx, var_18
  loc_00E37A80: push eax
  loc_00E37A81: push edx
  loc_00E37A82: call edi
  loc_00E37A84: push eax
  loc_00E37A85: call [00401238h] ; __vbaLateIdSt
  loc_00E37A8B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E37A91: lea ecx, var_18
  loc_00E37A94: call ebx
  loc_00E37A96: mov eax, [esi]
  loc_00E37A98: push esi
  loc_00E37A99: call [eax+00000324h]
  loc_00E37A9F: lea ecx, var_18
  loc_00E37AA2: push eax
  loc_00E37AA3: push ecx
  loc_00E37AA4: call edi
  loc_00E37AA6: mov edx, [eax]
  loc_00E37AA8: push eax
  loc_00E37AA9: mov var_50, eax
  loc_00E37AAC: call [edx+00000204h]
  loc_00E37AB2: test eax, eax
  loc_00E37AB4: fnclex
  loc_00E37AB6: jge 00E37ACDh
  loc_00E37AB8: mov ecx, var_50
  loc_00E37ABB: push 00000204h
  loc_00E37AC0: push 006DCB70h
  loc_00E37AC5: push ecx
  loc_00E37AC6: push eax
  loc_00E37AC7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37ACD: lea ecx, var_18
  loc_00E37AD0: call ebx
  loc_00E37AD2: mov edx, [esi]
  loc_00E37AD4: push esi
  loc_00E37AD5: call [edx+00000320h]
  loc_00E37ADB: push eax
  loc_00E37ADC: lea eax, var_18
  loc_00E37ADF: push eax
  loc_00E37AE0: call edi
  loc_00E37AE2: mov ecx, [eax]
  loc_00E37AE4: push 00000000h
  loc_00E37AE6: push eax
  loc_00E37AE7: mov var_50, eax
  loc_00E37AEA: call [ecx+0000009Ch]
  loc_00E37AF0: test eax, eax
  loc_00E37AF2: fnclex
  loc_00E37AF4: jge 00E37B0Bh
  loc_00E37AF6: mov edx, var_50
  loc_00E37AF9: push 0000009Ch
  loc_00E37AFE: push 006E03D4h
  loc_00E37B03: push edx
  loc_00E37B04: push eax
  loc_00E37B05: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37B0B: lea ecx, var_18
  loc_00E37B0E: call ebx
  loc_00E37B10: mov eax, [esi]
  loc_00E37B12: push esi
  loc_00E37B13: call [eax+00000328h]
  loc_00E37B19: lea ecx, var_18
  loc_00E37B1C: push eax
  loc_00E37B1D: push ecx
  loc_00E37B1E: call edi
  loc_00E37B20: mov edx, [eax]
  loc_00E37B22: push 00000000h
  loc_00E37B24: push eax
  loc_00E37B25: mov var_50, eax
  loc_00E37B28: call [edx+0000009Ch]
  loc_00E37B2E: test eax, eax
  loc_00E37B30: fnclex
  loc_00E37B32: jge 00E37B49h
  loc_00E37B34: mov ecx, var_50
  loc_00E37B37: push 0000009Ch
  loc_00E37B3C: push 006E03D4h
  loc_00E37B41: push ecx
  loc_00E37B42: push eax
  loc_00E37B43: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37B49: lea ecx, var_18
  loc_00E37B4C: call ebx
  loc_00E37B4E: mov edx, [esi]
  loc_00E37B50: push esi
  loc_00E37B51: call [edx+00000314h]
  loc_00E37B57: push eax
  loc_00E37B58: lea eax, var_18
  loc_00E37B5B: push eax
  loc_00E37B5C: call edi
  loc_00E37B5E: mov ecx, [eax]
  loc_00E37B60: push 00000000h
  loc_00E37B62: push eax
  loc_00E37B63: mov var_50, eax
  loc_00E37B66: call [ecx+0000009Ch]
  loc_00E37B6C: test eax, eax
  loc_00E37B6E: fnclex
  loc_00E37B70: jge 00E37B87h
  loc_00E37B72: mov edx, var_50
  loc_00E37B75: push 0000009Ch
  loc_00E37B7A: push 006E03D4h
  loc_00E37B7F: push edx
  loc_00E37B80: push eax
  loc_00E37B81: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37B87: lea ecx, var_18
  loc_00E37B8A: call ebx
  loc_00E37B8C: mov eax, [esi]
  loc_00E37B8E: push esi
  loc_00E37B8F: call [eax+0000031Ch]
  loc_00E37B95: lea ecx, var_18
  loc_00E37B98: push eax
  loc_00E37B99: push ecx
  loc_00E37B9A: call edi
  loc_00E37B9C: mov edx, [eax]
  loc_00E37B9E: push 00000000h
  loc_00E37BA0: push eax
  loc_00E37BA1: mov var_50, eax
  loc_00E37BA4: call [edx+0000009Ch]
  loc_00E37BAA: test eax, eax
  loc_00E37BAC: fnclex
  loc_00E37BAE: jge 00E37BC5h
  loc_00E37BB0: mov ecx, var_50
  loc_00E37BB3: push 0000009Ch
  loc_00E37BB8: push 006E03D4h
  loc_00E37BBD: push ecx
  loc_00E37BBE: push eax
  loc_00E37BBF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37BC5: lea ecx, var_18
  loc_00E37BC8: call ebx
  loc_00E37BCA: mov edx, [esi]
  loc_00E37BCC: push esi
  loc_00E37BCD: call [edx+00000324h]
  loc_00E37BD3: push eax
  loc_00E37BD4: lea eax, var_18
  loc_00E37BD7: push eax
  loc_00E37BD8: call edi
  loc_00E37BDA: mov ecx, [eax]
  loc_00E37BDC: push 006DCC80h
  loc_00E37BE1: push eax
  loc_00E37BE2: mov var_50, eax
  loc_00E37BE5: call [ecx+000000A4h]
  loc_00E37BEB: test eax, eax
  loc_00E37BED: fnclex
  loc_00E37BEF: jge 00E37C06h
  loc_00E37BF1: mov edx, var_50
  loc_00E37BF4: push 000000A4h
  loc_00E37BF9: push 006DCB70h
  loc_00E37BFE: push edx
  loc_00E37BFF: push eax
  loc_00E37C00: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E37C06: lea ecx, var_18
  loc_00E37C09: call ebx
  loc_00E37C0B: sub esp, 00000010h
  loc_00E37C0E: mov ecx, 00000008h
  loc_00E37C13: mov edx, esp
  loc_00E37C15: mov eax, 006E1738h ; "lc"
  loc_00E37C1A: push 0000000Eh
  loc_00E37C1C: push esi
  loc_00E37C1D: mov [edx], ecx
  loc_00E37C1F: mov ecx, var_38
  loc_00E37C22: mov [edx+00000004h], ecx
  loc_00E37C25: mov ecx, [esi]
  loc_00E37C27: mov [edx+00000008h], eax
  loc_00E37C2A: mov eax, var_30
  loc_00E37C2D: mov [edx+0000000Ch], eax
  loc_00E37C30: call [ecx+00000354h]
  loc_00E37C36: lea edx, var_18
  loc_00E37C39: push eax
  loc_00E37C3A: push edx
  loc_00E37C3B: call edi
  loc_00E37C3D: push eax
  loc_00E37C3E: call [00401238h] ; __vbaLateIdSt
  loc_00E37C44: lea ecx, var_18
  loc_00E37C47: call ebx
  loc_00E37C49: mov eax, [esi]
  loc_00E37C4B: push 00000000h
  loc_00E37C4D: push 00000019h
  loc_00E37C4F: push esi
  loc_00E37C50: call [eax+00000354h]
  loc_00E37C56: lea ecx, var_18
  loc_00E37C59: push eax
  loc_00E37C5A: push ecx
  loc_00E37C5B: call edi
  loc_00E37C5D: push eax
  loc_00E37C5E: call [00401030h] ; __vbaLateIdCall
  loc_00E37C64: add esp, 0000000Ch
  loc_00E37C67: lea ecx, var_18
  loc_00E37C6A: call ebx
  loc_00E37C6C: mov edx, [esi]
  loc_00E37C6E: push 006E05D8h
  loc_00E37C73: push esi
  loc_00E37C74: call [edx+00000354h]
  loc_00E37C7A: push eax
  loc_00E37C7B: lea eax, var_18
  loc_00E37C7E: push eax
  loc_00E37C7F: call edi
  loc_00E37C81: push eax
  loc_00E37C82: call [00401224h] ; __vbaCastObj
  loc_00E37C88: lea ecx, var_2C
  loc_00E37C8B: mov var_24, eax
  loc_00E37C8E: push ecx
  loc_00E37C8F: mov var_2C, 0000000Dh
  loc_00E37C96: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E37C9C: mov ecx, [eax]
  loc_00E37C9E: sub esp, 00000010h
  loc_00E37CA1: mov edx, esp
  loc_00E37CA3: push 00000000h
  loc_00E37CA5: push 0000002Ah
  loc_00E37CA7: mov [edx], ecx
  loc_00E37CA9: mov ecx, [eax+00000004h]
  loc_00E37CAC: push esi
  loc_00E37CAD: mov [edx+00000004h], ecx
  loc_00E37CB0: mov ecx, [eax+00000008h]
  loc_00E37CB3: mov eax, [eax+0000000Ch]
  loc_00E37CB6: mov [edx+00000008h], ecx
  loc_00E37CB9: mov ecx, [esi]
  loc_00E37CBB: mov [edx+0000000Ch], eax
  loc_00E37CBE: call [ecx+00000358h]
  loc_00E37CC4: lea edx, var_1C
  loc_00E37CC7: push eax
  loc_00E37CC8: push edx
  loc_00E37CC9: call edi
  loc_00E37CCB: push eax
  loc_00E37CCC: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E37CD2: lea eax, var_1C
  loc_00E37CD5: lea ecx, var_18
  loc_00E37CD8: push eax
  loc_00E37CD9: push ecx
  loc_00E37CDA: push 00000002h
  loc_00E37CDC: call [00401048h] ; __vbaFreeObjList
  loc_00E37CE2: add esp, 00000028h
  loc_00E37CE5: lea ecx, var_2C
  loc_00E37CE8: call [00401024h] ; __vbaFreeVar
  loc_00E37CEE: sub esp, 00000010h
  loc_00E37CF1: mov ecx, 0000000Bh
  loc_00E37CF6: mov edx, esp
  loc_00E37CF8: xor eax, eax
  loc_00E37CFA: push 00000006h
  loc_00E37CFC: push esi
  loc_00E37CFD: mov [edx], ecx
  loc_00E37CFF: mov ecx, var_38
  loc_00E37D02: mov [edx+00000004h], ecx
  loc_00E37D05: mov ecx, [esi]
  loc_00E37D07: mov [edx+00000008h], eax
  loc_00E37D0A: mov eax, var_30
  loc_00E37D0D: mov [edx+0000000Ch], eax
  loc_00E37D10: call [ecx+00000358h]
  loc_00E37D16: lea edx, var_18
  loc_00E37D19: push eax
  loc_00E37D1A: push edx
  loc_00E37D1B: call edi
  loc_00E37D1D: push eax
  loc_00E37D1E: call [00401238h] ; __vbaLateIdSt
  loc_00E37D24: lea ecx, var_18
  loc_00E37D27: call ebx
  loc_00E37D29: sub esp, 00000010h
  loc_00E37D2C: mov ecx, 0000000Bh
  loc_00E37D31: mov edx, esp
  loc_00E37D33: xor eax, eax
  loc_00E37D35: push 8001000Eh
  loc_00E37D3A: push esi
  loc_00E37D3B: mov [edx], ecx
  loc_00E37D3D: mov ecx, var_38
  loc_00E37D40: mov [edx+00000004h], ecx
  loc_00E37D43: mov ecx, [esi]
  loc_00E37D45: mov [edx+00000008h], eax
  loc_00E37D48: mov eax, var_30
  loc_00E37D4B: mov [edx+0000000Ch], eax
  loc_00E37D4E: call [ecx+00000358h]
  loc_00E37D54: lea edx, var_18
  loc_00E37D57: push eax
  loc_00E37D58: push edx
  loc_00E37D59: call edi
  loc_00E37D5B: push eax
  loc_00E37D5C: call [00401238h] ; __vbaLateIdSt
  loc_00E37D62: lea ecx, var_18
  loc_00E37D65: call ebx
  loc_00E37D67: mov eax, [esi]
  loc_00E37D69: push 00000000h
  loc_00E37D6B: push FFFFFDDAh
  loc_00E37D70: push esi
  loc_00E37D71: call [eax+00000358h]
  loc_00E37D77: lea ecx, var_18
  loc_00E37D7A: push eax
  loc_00E37D7B: push ecx
  loc_00E37D7C: call edi
  loc_00E37D7E: push eax
  loc_00E37D7F: call [00401030h] ; __vbaLateIdCall
  loc_00E37D85: add esp, 0000000Ch
  loc_00E37D88: lea ecx, var_18
  loc_00E37D8B: call ebx
  loc_00E37D8D: mov var_4, 00000000h
  loc_00E37D94: push 00E37DB9h
  loc_00E37D99: jmp 00E37DB8h
  loc_00E37D9B: lea edx, var_1C
  loc_00E37D9E: lea eax, var_18
  loc_00E37DA1: push edx
  loc_00E37DA2: push eax
  loc_00E37DA3: push 00000002h
  loc_00E37DA5: call [00401048h] ; __vbaFreeObjList
  loc_00E37DAB: add esp, 0000000Ch
  loc_00E37DAE: lea ecx, var_2C
  loc_00E37DB1: call [00401024h] ; __vbaFreeVar
  loc_00E37DB7: ret
  loc_00E37DB8: ret
  loc_00E37DB9: mov eax, Me
  loc_00E37DBC: push eax
  loc_00E37DBD: mov ecx, [eax]
  loc_00E37DBF: call [ecx+00000008h]
  loc_00E37DC2: mov eax, var_4
  loc_00E37DC5: mov ecx, var_14
  loc_00E37DC8: pop edi
  loc_00E37DC9: pop esi
  loc_00E37DCA: mov fs:[00000000h], ecx
  loc_00E37DD1: pop ebx
  loc_00E37DD2: mov esp, ebp
  loc_00E37DD4: pop ebp
  loc_00E37DD5: retn 0004h
End Sub

Private Sub batal_UnknownEvent_9 'E36530
  loc_00E36530: push ebp
  loc_00E36531: mov ebp, esp
  loc_00E36533: sub esp, 0000000Ch
  loc_00E36536: push 00402836h ; __vbaExceptHandler
  loc_00E3653B: mov eax, fs:[00000000h]
  loc_00E36541: push eax
  loc_00E36542: mov fs:[00000000h], esp
  loc_00E36549: sub esp, 00000048h
  loc_00E3654C: push ebx
  loc_00E3654D: push esi
  loc_00E3654E: push edi
  loc_00E3654F: mov var_C, esp
  loc_00E36552: mov var_8, 004026E0h
  loc_00E36559: mov esi, Me
  loc_00E3655C: mov eax, esi
  loc_00E3655E: and eax, 00000001h
  loc_00E36561: mov var_4, eax
  loc_00E36564: and esi, FFFFFFFEh
  loc_00E36567: push esi
  loc_00E36568: mov Me, esi
  loc_00E3656B: mov ecx, [esi]
  loc_00E3656D: call [ecx+00000004h]
  loc_00E36570: mov edx, [esi]
  loc_00E36572: xor eax, eax
  loc_00E36574: push esi
  loc_00E36575: mov var_18, eax
  loc_00E36578: mov var_1C, eax
  loc_00E3657B: mov var_2C, eax
  loc_00E3657E: call [edx+0000031Ch]
  loc_00E36584: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3658A: push eax
  loc_00E3658B: lea eax, var_18
  loc_00E3658E: push eax
  loc_00E3658F: call edi
  loc_00E36591: mov ebx, eax
  loc_00E36593: push 00000000h
  loc_00E36595: push ebx
  loc_00E36596: mov ecx, [ebx]
  loc_00E36598: call [ecx+000000E4h]
  loc_00E3659E: test eax, eax
  loc_00E365A0: fnclex
  loc_00E365A2: jge 00E365B6h
  loc_00E365A4: push 000000E4h
  loc_00E365A9: push 006E03D4h
  loc_00E365AE: push ebx
  loc_00E365AF: push eax
  loc_00E365B0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E365B6: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E365BC: lea ecx, var_18
  loc_00E365BF: call ebx
  loc_00E365C1: mov edx, [esi]
  loc_00E365C3: push esi
  loc_00E365C4: call [edx+00000320h]
  loc_00E365CA: push eax
  loc_00E365CB: lea eax, var_18
  loc_00E365CE: push eax
  loc_00E365CF: call edi
  loc_00E365D1: mov ecx, [eax]
  loc_00E365D3: push 00000000h
  loc_00E365D5: push eax
  loc_00E365D6: mov var_50, eax
  loc_00E365D9: call [ecx+000000E4h]
  loc_00E365DF: test eax, eax
  loc_00E365E1: fnclex
  loc_00E365E3: jge 00E365FAh
  loc_00E365E5: mov edx, var_50
  loc_00E365E8: push 000000E4h
  loc_00E365ED: push 006E03D4h
  loc_00E365F2: push edx
  loc_00E365F3: push eax
  loc_00E365F4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E365FA: lea ecx, var_18
  loc_00E365FD: call ebx
  loc_00E365FF: mov eax, [esi]
  loc_00E36601: push esi
  loc_00E36602: call [eax+00000318h]
  loc_00E36608: lea ecx, var_18
  loc_00E3660B: push eax
  loc_00E3660C: push ecx
  loc_00E3660D: call edi
  loc_00E3660F: mov edx, [eax]
  loc_00E36611: push 00000000h
  loc_00E36613: push eax
  loc_00E36614: mov var_50, eax
  loc_00E36617: call [edx+000000E4h]
  loc_00E3661D: test eax, eax
  loc_00E3661F: fnclex
  loc_00E36621: jge 00E36638h
  loc_00E36623: mov ecx, var_50
  loc_00E36626: push 000000E4h
  loc_00E3662B: push 006E03D4h
  loc_00E36630: push ecx
  loc_00E36631: push eax
  loc_00E36632: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36638: lea ecx, var_18
  loc_00E3663B: call ebx
  loc_00E3663D: mov edx, [esi]
  loc_00E3663F: push esi
  loc_00E36640: call [edx+00000328h]
  loc_00E36646: push eax
  loc_00E36647: lea eax, var_18
  loc_00E3664A: push eax
  loc_00E3664B: call edi
  loc_00E3664D: mov ecx, [eax]
  loc_00E3664F: push 00000000h
  loc_00E36651: push eax
  loc_00E36652: mov var_50, eax
  loc_00E36655: call [ecx+0000009Ch]
  loc_00E3665B: test eax, eax
  loc_00E3665D: fnclex
  loc_00E3665F: jge 00E36676h
  loc_00E36661: mov edx, var_50
  loc_00E36664: push 0000009Ch
  loc_00E36669: push 006E03D4h
  loc_00E3666E: push edx
  loc_00E3666F: push eax
  loc_00E36670: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36676: lea ecx, var_18
  loc_00E36679: call ebx
  loc_00E3667B: sub esp, 00000010h
  loc_00E3667E: mov ecx, 0000000Bh
  loc_00E36683: mov edx, esp
  loc_00E36685: xor eax, eax
  loc_00E36687: push 80010007h
  loc_00E3668C: push esi
  loc_00E3668D: mov [edx], ecx
  loc_00E3668F: mov ecx, var_38
  loc_00E36692: mov [edx+00000004h], ecx
  loc_00E36695: mov ecx, [esi]
  loc_00E36697: mov [edx+00000008h], eax
  loc_00E3669A: mov eax, var_30
  loc_00E3669D: mov [edx+0000000Ch], eax
  loc_00E366A0: call [ecx+0000032Ch]
  loc_00E366A6: lea edx, var_18
  loc_00E366A9: push eax
  loc_00E366AA: push edx
  loc_00E366AB: call edi
  loc_00E366AD: push eax
  loc_00E366AE: call [00401238h] ; __vbaLateIdSt
  loc_00E366B4: lea ecx, var_18
  loc_00E366B7: call ebx
  loc_00E366B9: mov eax, [esi]
  loc_00E366BB: push esi
  loc_00E366BC: call [eax+0000031Ch]
  loc_00E366C2: lea ecx, var_18
  loc_00E366C5: push eax
  loc_00E366C6: push ecx
  loc_00E366C7: call edi
  loc_00E366C9: mov edx, [eax]
  loc_00E366CB: push FFFFFFFFh
  loc_00E366CD: push eax
  loc_00E366CE: mov var_50, eax
  loc_00E366D1: call [edx+0000009Ch]
  loc_00E366D7: test eax, eax
  loc_00E366D9: fnclex
  loc_00E366DB: jge 00E366F2h
  loc_00E366DD: mov ecx, var_50
  loc_00E366E0: push 0000009Ch
  loc_00E366E5: push 006E03D4h
  loc_00E366EA: push ecx
  loc_00E366EB: push eax
  loc_00E366EC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E366F2: lea ecx, var_18
  loc_00E366F5: call ebx
  loc_00E366F7: mov edx, [esi]
  loc_00E366F9: push esi
  loc_00E366FA: call [edx+00000320h]
  loc_00E36700: push eax
  loc_00E36701: lea eax, var_18
  loc_00E36704: push eax
  loc_00E36705: call edi
  loc_00E36707: mov ecx, [eax]
  loc_00E36709: push FFFFFFFFh
  loc_00E3670B: push eax
  loc_00E3670C: mov var_50, eax
  loc_00E3670F: call [ecx+0000009Ch]
  loc_00E36715: test eax, eax
  loc_00E36717: fnclex
  loc_00E36719: jge 00E36730h
  loc_00E3671B: mov edx, var_50
  loc_00E3671E: push 0000009Ch
  loc_00E36723: push 006E03D4h
  loc_00E36728: push edx
  loc_00E36729: push eax
  loc_00E3672A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36730: lea ecx, var_18
  loc_00E36733: call ebx
  loc_00E36735: mov eax, [esi]
  loc_00E36737: push esi
  loc_00E36738: call [eax+00000318h]
  loc_00E3673E: lea ecx, var_18
  loc_00E36741: push eax
  loc_00E36742: push ecx
  loc_00E36743: call edi
  loc_00E36745: mov edx, [eax]
  loc_00E36747: push FFFFFFFFh
  loc_00E36749: push eax
  loc_00E3674A: mov var_50, eax
  loc_00E3674D: call [edx+0000009Ch]
  loc_00E36753: test eax, eax
  loc_00E36755: fnclex
  loc_00E36757: jge 00E3676Eh
  loc_00E36759: mov ecx, var_50
  loc_00E3675C: push 0000009Ch
  loc_00E36761: push 006E03D4h
  loc_00E36766: push ecx
  loc_00E36767: push eax
  loc_00E36768: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3676E: lea ecx, var_18
  loc_00E36771: call ebx
  loc_00E36773: mov edx, [esi]
  loc_00E36775: push esi
  loc_00E36776: call [edx+00000328h]
  loc_00E3677C: push eax
  loc_00E3677D: lea eax, var_18
  loc_00E36780: push eax
  loc_00E36781: call edi
  loc_00E36783: mov ecx, [eax]
  loc_00E36785: push FFFFFFFFh
  loc_00E36787: push eax
  loc_00E36788: mov var_50, eax
  loc_00E3678B: call [ecx+0000009Ch]
  loc_00E36791: test eax, eax
  loc_00E36793: fnclex
  loc_00E36795: jge 00E367ACh
  loc_00E36797: mov edx, var_50
  loc_00E3679A: push 0000009Ch
  loc_00E3679F: push 006E03D4h
  loc_00E367A4: push edx
  loc_00E367A5: push eax
  loc_00E367A6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E367AC: lea ecx, var_18
  loc_00E367AF: call ebx
  loc_00E367B1: mov eax, [esi]
  loc_00E367B3: push esi
  loc_00E367B4: call [eax+00000324h]
  loc_00E367BA: lea ecx, var_18
  loc_00E367BD: push eax
  loc_00E367BE: push ecx
  loc_00E367BF: call edi
  loc_00E367C1: mov edx, [eax]
  loc_00E367C3: push 006DCC80h
  loc_00E367C8: push eax
  loc_00E367C9: mov var_50, eax
  loc_00E367CC: call [edx+000000A4h]
  loc_00E367D2: test eax, eax
  loc_00E367D4: fnclex
  loc_00E367D6: jge 00E367EDh
  loc_00E367D8: mov ecx, var_50
  loc_00E367DB: push 000000A4h
  loc_00E367E0: push 006DCB70h
  loc_00E367E5: push ecx
  loc_00E367E6: push eax
  loc_00E367E7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E367ED: lea ecx, var_18
  loc_00E367F0: call ebx
  loc_00E367F2: mov edx, [esi]
  loc_00E367F4: push esi
  loc_00E367F5: call [edx+00000314h]
  loc_00E367FB: push eax
  loc_00E367FC: lea eax, var_18
  loc_00E367FF: push eax
  loc_00E36800: call edi
  loc_00E36802: mov ecx, [eax]
  loc_00E36804: push FFFFFFFFh
  loc_00E36806: push eax
  loc_00E36807: mov var_50, eax
  loc_00E3680A: call [ecx+0000009Ch]
  loc_00E36810: test eax, eax
  loc_00E36812: fnclex
  loc_00E36814: jge 00E3682Bh
  loc_00E36816: mov edx, var_50
  loc_00E36819: push 0000009Ch
  loc_00E3681E: push 006E03D4h
  loc_00E36823: push edx
  loc_00E36824: push eax
  loc_00E36825: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3682B: lea ecx, var_18
  loc_00E3682E: call ebx
  loc_00E36830: mov eax, [esi]
  loc_00E36832: push esi
  loc_00E36833: call [eax+00000324h]
  loc_00E36839: lea ecx, var_18
  loc_00E3683C: push eax
  loc_00E3683D: push ecx
  loc_00E3683E: call edi
  loc_00E36840: mov edx, [eax]
  loc_00E36842: push eax
  loc_00E36843: mov var_50, eax
  loc_00E36846: call [edx+00000204h]
  loc_00E3684C: test eax, eax
  loc_00E3684E: fnclex
  loc_00E36850: jge 00E36867h
  loc_00E36852: mov ecx, var_50
  loc_00E36855: push 00000204h
  loc_00E3685A: push 006DCB70h
  loc_00E3685F: push ecx
  loc_00E36860: push eax
  loc_00E36861: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E36867: lea ecx, var_18
  loc_00E3686A: call ebx
  loc_00E3686C: sub esp, 00000010h
  loc_00E3686F: mov ecx, 00000008h
  loc_00E36874: mov edx, esp
  loc_00E36876: mov eax, 006E1738h ; "lc"
  loc_00E3687B: push 0000000Eh
  loc_00E3687D: push esi
  loc_00E3687E: mov [edx], ecx
  loc_00E36880: mov ecx, var_38
  loc_00E36883: mov [edx+00000004h], ecx
  loc_00E36886: mov ecx, [esi]
  loc_00E36888: mov [edx+00000008h], eax
  loc_00E3688B: mov eax, var_30
  loc_00E3688E: mov [edx+0000000Ch], eax
  loc_00E36891: call [ecx+00000354h]
  loc_00E36897: lea edx, var_18
  loc_00E3689A: push eax
  loc_00E3689B: push edx
  loc_00E3689C: call edi
  loc_00E3689E: push eax
  loc_00E3689F: call [00401238h] ; __vbaLateIdSt
  loc_00E368A5: lea ecx, var_18
  loc_00E368A8: call ebx
  loc_00E368AA: mov eax, [esi]
  loc_00E368AC: push 00000000h
  loc_00E368AE: push 00000019h
  loc_00E368B0: push esi
  loc_00E368B1: call [eax+00000354h]
  loc_00E368B7: lea ecx, var_18
  loc_00E368BA: push eax
  loc_00E368BB: push ecx
  loc_00E368BC: call edi
  loc_00E368BE: push eax
  loc_00E368BF: call [00401030h] ; __vbaLateIdCall
  loc_00E368C5: add esp, 0000000Ch
  loc_00E368C8: lea ecx, var_18
  loc_00E368CB: call ebx
  loc_00E368CD: mov edx, [esi]
  loc_00E368CF: push 006E05D8h
  loc_00E368D4: push esi
  loc_00E368D5: call [edx+00000354h]
  loc_00E368DB: push eax
  loc_00E368DC: lea eax, var_18
  loc_00E368DF: push eax
  loc_00E368E0: call edi
  loc_00E368E2: push eax
  loc_00E368E3: call [00401224h] ; __vbaCastObj
  loc_00E368E9: lea ecx, var_2C
  loc_00E368EC: mov var_24, eax
  loc_00E368EF: push ecx
  loc_00E368F0: mov var_2C, 0000000Dh
  loc_00E368F7: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E368FD: mov ecx, [eax]
  loc_00E368FF: sub esp, 00000010h
  loc_00E36902: mov edx, esp
  loc_00E36904: push 00000000h
  loc_00E36906: push 0000002Ah
  loc_00E36908: mov [edx], ecx
  loc_00E3690A: mov ecx, [eax+00000004h]
  loc_00E3690D: push esi
  loc_00E3690E: mov [edx+00000004h], ecx
  loc_00E36911: mov ecx, [eax+00000008h]
  loc_00E36914: mov eax, [eax+0000000Ch]
  loc_00E36917: mov [edx+00000008h], ecx
  loc_00E3691A: mov ecx, [esi]
  loc_00E3691C: mov [edx+0000000Ch], eax
  loc_00E3691F: call [ecx+00000358h]
  loc_00E36925: lea edx, var_1C
  loc_00E36928: push eax
  loc_00E36929: push edx
  loc_00E3692A: call edi
  loc_00E3692C: push eax
  loc_00E3692D: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E36933: lea eax, var_1C
  loc_00E36936: lea ecx, var_18
  loc_00E36939: push eax
  loc_00E3693A: push ecx
  loc_00E3693B: push 00000002h
  loc_00E3693D: call [00401048h] ; __vbaFreeObjList
  loc_00E36943: add esp, 00000028h
  loc_00E36946: lea ecx, var_2C
  loc_00E36949: call [00401024h] ; __vbaFreeVar
  loc_00E3694F: sub esp, 00000010h
  loc_00E36952: mov ecx, 0000000Bh
  loc_00E36957: mov edx, esp
  loc_00E36959: xor eax, eax
  loc_00E3695B: push 00000006h
  loc_00E3695D: push esi
  loc_00E3695E: mov [edx], ecx
  loc_00E36960: mov ecx, var_38
  loc_00E36963: mov [edx+00000004h], ecx
  loc_00E36966: mov ecx, [esi]
  loc_00E36968: mov [edx+00000008h], eax
  loc_00E3696B: mov eax, var_30
  loc_00E3696E: mov [edx+0000000Ch], eax
  loc_00E36971: call [ecx+00000358h]
  loc_00E36977: lea edx, var_18
  loc_00E3697A: push eax
  loc_00E3697B: push edx
  loc_00E3697C: call edi
  loc_00E3697E: push eax
  loc_00E3697F: call [00401238h] ; __vbaLateIdSt
  loc_00E36985: lea ecx, var_18
  loc_00E36988: call ebx
  loc_00E3698A: sub esp, 00000010h
  loc_00E3698D: mov ecx, 0000000Bh
  loc_00E36992: mov edx, esp
  loc_00E36994: xor eax, eax
  loc_00E36996: push 8001000Eh
  loc_00E3699B: push esi
  loc_00E3699C: mov [edx], ecx
  loc_00E3699E: mov ecx, var_38
  loc_00E369A1: mov [edx+00000004h], ecx
  loc_00E369A4: mov ecx, [esi]
  loc_00E369A6: mov [edx+00000008h], eax
  loc_00E369A9: mov eax, var_30
  loc_00E369AC: mov [edx+0000000Ch], eax
  loc_00E369AF: call [ecx+00000358h]
  loc_00E369B5: lea edx, var_18
  loc_00E369B8: push eax
  loc_00E369B9: push edx
  loc_00E369BA: call edi
  loc_00E369BC: push eax
  loc_00E369BD: call [00401238h] ; __vbaLateIdSt
  loc_00E369C3: lea ecx, var_18
  loc_00E369C6: call ebx
  loc_00E369C8: mov eax, [esi]
  loc_00E369CA: push 00000000h
  loc_00E369CC: push FFFFFDDAh
  loc_00E369D1: push esi
  loc_00E369D2: call [eax+00000358h]
  loc_00E369D8: lea ecx, var_18
  loc_00E369DB: push eax
  loc_00E369DC: push ecx
  loc_00E369DD: call edi
  loc_00E369DF: push eax
  loc_00E369E0: call [00401030h] ; __vbaLateIdCall
  loc_00E369E6: add esp, 0000000Ch
  loc_00E369E9: lea ecx, var_18
  loc_00E369EC: call ebx
  loc_00E369EE: mov var_4, 00000000h
  loc_00E369F5: push 00E36A1Ah
  loc_00E369FA: jmp 00E36A19h
  loc_00E369FC: lea edx, var_1C
  loc_00E369FF: lea eax, var_18
  loc_00E36A02: push edx
  loc_00E36A03: push eax
  loc_00E36A04: push 00000002h
  loc_00E36A06: call [00401048h] ; __vbaFreeObjList
  loc_00E36A0C: add esp, 0000000Ch
  loc_00E36A0F: lea ecx, var_2C
  loc_00E36A12: call [00401024h] ; __vbaFreeVar
  loc_00E36A18: ret
  loc_00E36A19: ret
  loc_00E36A1A: mov eax, Me
  loc_00E36A1D: push eax
  loc_00E36A1E: mov ecx, [eax]
  loc_00E36A20: call [ecx+00000008h]
  loc_00E36A23: mov eax, var_4
  loc_00E36A26: mov ecx, var_14
  loc_00E36A29: pop edi
  loc_00E36A2A: pop esi
  loc_00E36A2B: mov fs:[00000000h], ecx
  loc_00E36A32: pop ebx
  loc_00E36A33: mov esp, ebp
  loc_00E36A35: pop ebp
  loc_00E36A36: retn 0004h
End Sub

Private Sub optno_Click() 'E381C0
  loc_00E381C0: push ebp
  loc_00E381C1: mov ebp, esp
  loc_00E381C3: sub esp, 0000000Ch
  loc_00E381C6: push 00402836h ; __vbaExceptHandler
  loc_00E381CB: mov eax, fs:[00000000h]
  loc_00E381D1: push eax
  loc_00E381D2: mov fs:[00000000h], esp
  loc_00E381D9: sub esp, 00000048h
  loc_00E381DC: push ebx
  loc_00E381DD: push esi
  loc_00E381DE: push edi
  loc_00E381DF: mov var_C, esp
  loc_00E381E2: mov var_8, 00402750h
  loc_00E381E9: mov esi, Me
  loc_00E381EC: mov eax, esi
  loc_00E381EE: and eax, 00000001h
  loc_00E381F1: mov var_4, eax
  loc_00E381F4: and esi, FFFFFFFEh
  loc_00E381F7: push esi
  loc_00E381F8: mov Me, esi
  loc_00E381FB: mov ecx, [esi]
  loc_00E381FD: call [ecx+00000004h]
  loc_00E38200: sub esp, 00000010h
  loc_00E38203: mov ecx, 0000000Bh
  loc_00E38208: mov edx, esp
  loc_00E3820A: xor eax, eax
  loc_00E3820C: mov var_18, eax
  loc_00E3820F: mov var_1C, eax
  loc_00E38212: mov [edx], ecx
  loc_00E38214: mov ecx, var_38
  loc_00E38217: mov var_2C, eax
  loc_00E3821A: or eax, FFFFFFFFh
  loc_00E3821D: mov [edx+00000004h], ecx
  loc_00E38220: mov ecx, [esi]
  loc_00E38222: push 80010007h
  loc_00E38227: push esi
  loc_00E38228: mov [edx+00000008h], eax
  loc_00E3822B: mov eax, var_30
  loc_00E3822E: mov [edx+0000000Ch], eax
  loc_00E38231: call [ecx+0000032Ch]
  loc_00E38237: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3823D: lea edx, var_18
  loc_00E38240: push eax
  loc_00E38241: push edx
  loc_00E38242: call edi
  loc_00E38244: push eax
  loc_00E38245: call [00401238h] ; __vbaLateIdSt
  loc_00E3824B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E38251: lea ecx, var_18
  loc_00E38254: call ebx
  loc_00E38256: mov eax, [esi]
  loc_00E38258: push esi
  loc_00E38259: call [eax+00000324h]
  loc_00E3825F: lea ecx, var_18
  loc_00E38262: push eax
  loc_00E38263: push ecx
  loc_00E38264: call edi
  loc_00E38266: mov edx, [eax]
  loc_00E38268: push eax
  loc_00E38269: mov var_50, eax
  loc_00E3826C: call [edx+00000204h]
  loc_00E38272: test eax, eax
  loc_00E38274: fnclex
  loc_00E38276: jge 00E3828Dh
  loc_00E38278: mov ecx, var_50
  loc_00E3827B: push 00000204h
  loc_00E38280: push 006DCB70h
  loc_00E38285: push ecx
  loc_00E38286: push eax
  loc_00E38287: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3828D: lea ecx, var_18
  loc_00E38290: call ebx
  loc_00E38292: mov edx, [esi]
  loc_00E38294: push esi
  loc_00E38295: call [edx+0000031Ch]
  loc_00E3829B: push eax
  loc_00E3829C: lea eax, var_18
  loc_00E3829F: push eax
  loc_00E382A0: call edi
  loc_00E382A2: mov ecx, [eax]
  loc_00E382A4: push 00000000h
  loc_00E382A6: push eax
  loc_00E382A7: mov var_50, eax
  loc_00E382AA: call [ecx+0000009Ch]
  loc_00E382B0: test eax, eax
  loc_00E382B2: fnclex
  loc_00E382B4: jge 00E382CBh
  loc_00E382B6: mov edx, var_50
  loc_00E382B9: push 0000009Ch
  loc_00E382BE: push 006E03D4h
  loc_00E382C3: push edx
  loc_00E382C4: push eax
  loc_00E382C5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E382CB: lea ecx, var_18
  loc_00E382CE: call ebx
  loc_00E382D0: mov eax, [esi]
  loc_00E382D2: push esi
  loc_00E382D3: call [eax+00000314h]
  loc_00E382D9: lea ecx, var_18
  loc_00E382DC: push eax
  loc_00E382DD: push ecx
  loc_00E382DE: call edi
  loc_00E382E0: mov edx, [eax]
  loc_00E382E2: push 00000000h
  loc_00E382E4: push eax
  loc_00E382E5: mov var_50, eax
  loc_00E382E8: call [edx+0000009Ch]
  loc_00E382EE: test eax, eax
  loc_00E382F0: fnclex
  loc_00E382F2: jge 00E38309h
  loc_00E382F4: mov ecx, var_50
  loc_00E382F7: push 0000009Ch
  loc_00E382FC: push 006E03D4h
  loc_00E38301: push ecx
  loc_00E38302: push eax
  loc_00E38303: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38309: lea ecx, var_18
  loc_00E3830C: call ebx
  loc_00E3830E: mov edx, [esi]
  loc_00E38310: push esi
  loc_00E38311: call [edx+00000328h]
  loc_00E38317: push eax
  loc_00E38318: lea eax, var_18
  loc_00E3831B: push eax
  loc_00E3831C: call edi
  loc_00E3831E: mov ecx, [eax]
  loc_00E38320: push 00000000h
  loc_00E38322: push eax
  loc_00E38323: mov var_50, eax
  loc_00E38326: call [ecx+0000009Ch]
  loc_00E3832C: test eax, eax
  loc_00E3832E: fnclex
  loc_00E38330: jge 00E38347h
  loc_00E38332: mov edx, var_50
  loc_00E38335: push 0000009Ch
  loc_00E3833A: push 006E03D4h
  loc_00E3833F: push edx
  loc_00E38340: push eax
  loc_00E38341: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38347: lea ecx, var_18
  loc_00E3834A: call ebx
  loc_00E3834C: mov eax, [esi]
  loc_00E3834E: push esi
  loc_00E3834F: call [eax+00000318h]
  loc_00E38355: lea ecx, var_18
  loc_00E38358: push eax
  loc_00E38359: push ecx
  loc_00E3835A: call edi
  loc_00E3835C: mov edx, [eax]
  loc_00E3835E: push 00000000h
  loc_00E38360: push eax
  loc_00E38361: mov var_50, eax
  loc_00E38364: call [edx+0000009Ch]
  loc_00E3836A: test eax, eax
  loc_00E3836C: fnclex
  loc_00E3836E: jge 00E38385h
  loc_00E38370: mov ecx, var_50
  loc_00E38373: push 0000009Ch
  loc_00E38378: push 006E03D4h
  loc_00E3837D: push ecx
  loc_00E3837E: push eax
  loc_00E3837F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38385: lea ecx, var_18
  loc_00E38388: call ebx
  loc_00E3838A: mov edx, [esi]
  loc_00E3838C: push esi
  loc_00E3838D: call [edx+00000324h]
  loc_00E38393: push eax
  loc_00E38394: lea eax, var_18
  loc_00E38397: push eax
  loc_00E38398: call edi
  loc_00E3839A: mov ecx, [eax]
  loc_00E3839C: push 006DCC80h
  loc_00E383A1: push eax
  loc_00E383A2: mov var_50, eax
  loc_00E383A5: call [ecx+000000A4h]
  loc_00E383AB: test eax, eax
  loc_00E383AD: fnclex
  loc_00E383AF: jge 00E383C6h
  loc_00E383B1: mov edx, var_50
  loc_00E383B4: push 000000A4h
  loc_00E383B9: push 006DCB70h
  loc_00E383BE: push edx
  loc_00E383BF: push eax
  loc_00E383C0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E383C6: lea ecx, var_18
  loc_00E383C9: call ebx
  loc_00E383CB: sub esp, 00000010h
  loc_00E383CE: mov ecx, 00000008h
  loc_00E383D3: mov edx, esp
  loc_00E383D5: mov eax, 006E1738h ; "lc"
  loc_00E383DA: push 0000000Eh
  loc_00E383DC: push esi
  loc_00E383DD: mov [edx], ecx
  loc_00E383DF: mov ecx, var_38
  loc_00E383E2: mov [edx+00000004h], ecx
  loc_00E383E5: mov ecx, [esi]
  loc_00E383E7: mov [edx+00000008h], eax
  loc_00E383EA: mov eax, var_30
  loc_00E383ED: mov [edx+0000000Ch], eax
  loc_00E383F0: call [ecx+00000354h]
  loc_00E383F6: lea edx, var_18
  loc_00E383F9: push eax
  loc_00E383FA: push edx
  loc_00E383FB: call edi
  loc_00E383FD: push eax
  loc_00E383FE: call [00401238h] ; __vbaLateIdSt
  loc_00E38404: lea ecx, var_18
  loc_00E38407: call ebx
  loc_00E38409: mov eax, [esi]
  loc_00E3840B: push 00000000h
  loc_00E3840D: push 00000019h
  loc_00E3840F: push esi
  loc_00E38410: call [eax+00000354h]
  loc_00E38416: lea ecx, var_18
  loc_00E38419: push eax
  loc_00E3841A: push ecx
  loc_00E3841B: call edi
  loc_00E3841D: push eax
  loc_00E3841E: call [00401030h] ; __vbaLateIdCall
  loc_00E38424: add esp, 0000000Ch
  loc_00E38427: lea ecx, var_18
  loc_00E3842A: call ebx
  loc_00E3842C: mov edx, [esi]
  loc_00E3842E: push 006E05D8h
  loc_00E38433: push esi
  loc_00E38434: call [edx+00000354h]
  loc_00E3843A: push eax
  loc_00E3843B: lea eax, var_18
  loc_00E3843E: push eax
  loc_00E3843F: call edi
  loc_00E38441: push eax
  loc_00E38442: call [00401224h] ; __vbaCastObj
  loc_00E38448: lea ecx, var_2C
  loc_00E3844B: mov var_24, eax
  loc_00E3844E: push ecx
  loc_00E3844F: mov var_2C, 0000000Dh
  loc_00E38456: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3845C: mov ecx, [eax]
  loc_00E3845E: sub esp, 00000010h
  loc_00E38461: mov edx, esp
  loc_00E38463: push 00000000h
  loc_00E38465: push 0000002Ah
  loc_00E38467: mov [edx], ecx
  loc_00E38469: mov ecx, [eax+00000004h]
  loc_00E3846C: push esi
  loc_00E3846D: mov [edx+00000004h], ecx
  loc_00E38470: mov ecx, [eax+00000008h]
  loc_00E38473: mov eax, [eax+0000000Ch]
  loc_00E38476: mov [edx+00000008h], ecx
  loc_00E38479: mov ecx, [esi]
  loc_00E3847B: mov [edx+0000000Ch], eax
  loc_00E3847E: call [ecx+00000358h]
  loc_00E38484: lea edx, var_1C
  loc_00E38487: push eax
  loc_00E38488: push edx
  loc_00E38489: call edi
  loc_00E3848B: push eax
  loc_00E3848C: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E38492: lea eax, var_1C
  loc_00E38495: lea ecx, var_18
  loc_00E38498: push eax
  loc_00E38499: push ecx
  loc_00E3849A: push 00000002h
  loc_00E3849C: call [00401048h] ; __vbaFreeObjList
  loc_00E384A2: add esp, 00000028h
  loc_00E384A5: lea ecx, var_2C
  loc_00E384A8: call [00401024h] ; __vbaFreeVar
  loc_00E384AE: sub esp, 00000010h
  loc_00E384B1: mov ecx, 0000000Bh
  loc_00E384B6: mov edx, esp
  loc_00E384B8: xor eax, eax
  loc_00E384BA: push 00000006h
  loc_00E384BC: push esi
  loc_00E384BD: mov [edx], ecx
  loc_00E384BF: mov ecx, var_38
  loc_00E384C2: mov [edx+00000004h], ecx
  loc_00E384C5: mov ecx, [esi]
  loc_00E384C7: mov [edx+00000008h], eax
  loc_00E384CA: mov eax, var_30
  loc_00E384CD: mov [edx+0000000Ch], eax
  loc_00E384D0: call [ecx+00000358h]
  loc_00E384D6: lea edx, var_18
  loc_00E384D9: push eax
  loc_00E384DA: push edx
  loc_00E384DB: call edi
  loc_00E384DD: push eax
  loc_00E384DE: call [00401238h] ; __vbaLateIdSt
  loc_00E384E4: lea ecx, var_18
  loc_00E384E7: call ebx
  loc_00E384E9: sub esp, 00000010h
  loc_00E384EC: mov ecx, 0000000Bh
  loc_00E384F1: mov edx, esp
  loc_00E384F3: xor eax, eax
  loc_00E384F5: push 8001000Eh
  loc_00E384FA: push esi
  loc_00E384FB: mov [edx], ecx
  loc_00E384FD: mov ecx, var_38
  loc_00E38500: mov [edx+00000004h], ecx
  loc_00E38503: mov ecx, [esi]
  loc_00E38505: mov [edx+00000008h], eax
  loc_00E38508: mov eax, var_30
  loc_00E3850B: mov [edx+0000000Ch], eax
  loc_00E3850E: call [ecx+00000358h]
  loc_00E38514: lea edx, var_18
  loc_00E38517: push eax
  loc_00E38518: push edx
  loc_00E38519: call edi
  loc_00E3851B: push eax
  loc_00E3851C: call [00401238h] ; __vbaLateIdSt
  loc_00E38522: lea ecx, var_18
  loc_00E38525: call ebx
  loc_00E38527: mov eax, [esi]
  loc_00E38529: push 00000000h
  loc_00E3852B: push FFFFFDDAh
  loc_00E38530: push esi
  loc_00E38531: call [eax+00000358h]
  loc_00E38537: lea ecx, var_18
  loc_00E3853A: push eax
  loc_00E3853B: push ecx
  loc_00E3853C: call edi
  loc_00E3853E: push eax
  loc_00E3853F: call [00401030h] ; __vbaLateIdCall
  loc_00E38545: add esp, 0000000Ch
  loc_00E38548: lea ecx, var_18
  loc_00E3854B: call ebx
  loc_00E3854D: mov var_4, 00000000h
  loc_00E38554: push 00E38579h
  loc_00E38559: jmp 00E38578h
  loc_00E3855B: lea edx, var_1C
  loc_00E3855E: lea eax, var_18
  loc_00E38561: push edx
  loc_00E38562: push eax
  loc_00E38563: push 00000002h
  loc_00E38565: call [00401048h] ; __vbaFreeObjList
  loc_00E3856B: add esp, 0000000Ch
  loc_00E3856E: lea ecx, var_2C
  loc_00E38571: call [00401024h] ; __vbaFreeVar
  loc_00E38577: ret
  loc_00E38578: ret
  loc_00E38579: mov eax, Me
  loc_00E3857C: push eax
  loc_00E3857D: mov ecx, [eax]
  loc_00E3857F: call [ecx+00000008h]
  loc_00E38582: mov eax, var_4
  loc_00E38585: mov ecx, var_14
  loc_00E38588: pop edi
  loc_00E38589: pop esi
  loc_00E3858A: mov fs:[00000000h], ecx
  loc_00E38591: pop ebx
  loc_00E38592: mov esp, ebp
  loc_00E38594: pop ebp
  loc_00E38595: retn 0004h
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer) 'E38960
  loc_00E38960: push ebp
  loc_00E38961: mov ebp, esp
  loc_00E38963: sub esp, 0000000Ch
  loc_00E38966: push 00402836h ; __vbaExceptHandler
  loc_00E3896B: mov eax, fs:[00000000h]
  loc_00E38971: push eax
  loc_00E38972: mov fs:[00000000h], esp
  loc_00E38979: sub esp, 000000B4h
  loc_00E3897F: push ebx
  loc_00E38980: push esi
  loc_00E38981: push edi
  loc_00E38982: mov var_C, esp
  loc_00E38985: mov var_8, 00402780h
  loc_00E3898C: mov esi, Me
  loc_00E3898F: mov eax, esi
  loc_00E38991: and eax, 00000001h
  loc_00E38994: mov var_4, eax
  loc_00E38997: and esi, FFFFFFFEh
  loc_00E3899A: push esi
  loc_00E3899B: mov Me, esi
  loc_00E3899E: mov ecx, [esi]
  loc_00E389A0: call [ecx+00000004h]
  loc_00E389A3: mov edx, [esi]
  loc_00E389A5: xor ebx, ebx
  loc_00E389A7: push esi
  loc_00E389A8: mov var_24, ebx
  loc_00E389AB: mov var_28, ebx
  loc_00E389AE: mov var_2C, ebx
  loc_00E389B1: mov var_30, ebx
  loc_00E389B4: mov var_40, ebx
  loc_00E389B7: mov var_50, ebx
  loc_00E389BA: mov var_60, ebx
  loc_00E389BD: mov var_70, ebx
  loc_00E389C0: mov var_80, ebx
  loc_00E389C3: mov var_90, ebx
  loc_00E389C9: mov var_B4, ebx
  loc_00E389CF: call [edx+00000320h]
  loc_00E389D5: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E389DB: push eax
  loc_00E389DC: lea eax, var_30
  loc_00E389DF: push eax
  loc_00E389E0: call edi
  loc_00E389E2: mov ecx, [eax]
  loc_00E389E4: lea edx, var_B4
  loc_00E389EA: push edx
  loc_00E389EB: push eax
  loc_00E389EC: mov var_B8, eax
  loc_00E389F2: call [ecx+000000E0h]
  loc_00E389F8: cmp eax, ebx
  loc_00E389FA: fnclex
  loc_00E389FC: jge 00E38A16h
  loc_00E389FE: mov ecx, var_B8
  loc_00E38A04: push 000000E0h
  loc_00E38A09: push 006E03D4h
  loc_00E38A0E: push ecx
  loc_00E38A0F: push eax
  loc_00E38A10: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38A16: mov edx, var_B4
  loc_00E38A1C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E38A22: lea ecx, var_30
  loc_00E38A25: mov var_C0, edx
  loc_00E38A2B: call ebx
  loc_00E38A2D: cmp var_C0, 0000h
  loc_00E38A35: jz 00E38A9Dh
  loc_00E38A37: mov eax, [esi]
  loc_00E38A39: push esi
  loc_00E38A3A: call [eax+00000324h]
  loc_00E38A40: lea ecx, var_30
  loc_00E38A43: push eax
  loc_00E38A44: push ecx
  loc_00E38A45: call edi
  loc_00E38A47: mov edx, [eax]
  loc_00E38A49: lea ecx, var_28
  loc_00E38A4C: push ecx
  loc_00E38A4D: push eax
  loc_00E38A4E: mov var_B8, eax
  loc_00E38A54: call [edx+000000A0h]
  loc_00E38A5A: test eax, eax
  loc_00E38A5C: fnclex
  loc_00E38A5E: jge 00E38A78h
  loc_00E38A60: mov edx, var_B8
  loc_00E38A66: push 000000A0h
  loc_00E38A6B: push 006DCB70h
  loc_00E38A70: push edx
  loc_00E38A71: push eax
  loc_00E38A72: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38A78: mov eax, var_28
  loc_00E38A7B: push 006E1958h ; "select * from lc where no_daftar like '" & Chr(37)
  loc_00E38A80: push eax
  loc_00E38A81: call [0040106Ch] ; __vbaStrCat
  loc_00E38A87: mov edx, eax
  loc_00E38A89: lea ecx, var_2C
  loc_00E38A8C: call [00401228h] ; __vbaStrMove
  loc_00E38A92: push eax
  loc_00E38A93: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E38A98: jmp 00E38DB0h
  loc_00E38A9D: mov edx, [esi]
  loc_00E38A9F: push esi
  loc_00E38AA0: call [edx+0000031Ch]
  loc_00E38AA6: push eax
  loc_00E38AA7: lea eax, var_30
  loc_00E38AAA: push eax
  loc_00E38AAB: call edi
  loc_00E38AAD: mov ecx, [eax]
  loc_00E38AAF: lea edx, var_B4
  loc_00E38AB5: push edx
  loc_00E38AB6: push eax
  loc_00E38AB7: mov var_B8, eax
  loc_00E38ABD: call [ecx+000000E0h]
  loc_00E38AC3: test eax, eax
  loc_00E38AC5: fnclex
  loc_00E38AC7: jge 00E38AE1h
  loc_00E38AC9: mov ecx, var_B8
  loc_00E38ACF: push 000000E0h
  loc_00E38AD4: push 006E03D4h
  loc_00E38AD9: push ecx
  loc_00E38ADA: push eax
  loc_00E38ADB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38AE1: mov edx, var_B4
  loc_00E38AE7: lea ecx, var_30
  loc_00E38AEA: mov var_C0, edx
  loc_00E38AF0: call ebx
  loc_00E38AF2: cmp var_C0, 0000h
  loc_00E38AFA: jz 00E38B62h
  loc_00E38AFC: mov eax, [esi]
  loc_00E38AFE: push esi
  loc_00E38AFF: call [eax+00000324h]
  loc_00E38B05: lea ecx, var_30
  loc_00E38B08: push eax
  loc_00E38B09: push ecx
  loc_00E38B0A: call edi
  loc_00E38B0C: mov edx, [eax]
  loc_00E38B0E: lea ecx, var_28
  loc_00E38B11: push ecx
  loc_00E38B12: push eax
  loc_00E38B13: mov var_B8, eax
  loc_00E38B19: call [edx+000000A0h]
  loc_00E38B1F: test eax, eax
  loc_00E38B21: fnclex
  loc_00E38B23: jge 00E38B3Dh
  loc_00E38B25: mov edx, var_B8
  loc_00E38B2B: push 000000A0h
  loc_00E38B30: push 006DCB70h
  loc_00E38B35: push edx
  loc_00E38B36: push eax
  loc_00E38B37: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38B3D: mov eax, var_28
  loc_00E38B40: push 006E19DCh ; "select * from lc where nama_siswa like '" & Chr(37)
  loc_00E38B45: push eax
  loc_00E38B46: call [0040106Ch] ; __vbaStrCat
  loc_00E38B4C: mov edx, eax
  loc_00E38B4E: lea ecx, var_2C
  loc_00E38B51: call [00401228h] ; __vbaStrMove
  loc_00E38B57: push eax
  loc_00E38B58: push 006E1A34h ; Chr(37) & "' order by nama_siswa asc"
  loc_00E38B5D: jmp 00E38DB0h
  loc_00E38B62: mov edx, [esi]
  loc_00E38B64: push esi
  loc_00E38B65: call [edx+00000314h]
  loc_00E38B6B: push eax
  loc_00E38B6C: lea eax, var_30
  loc_00E38B6F: push eax
  loc_00E38B70: call edi
  loc_00E38B72: mov ecx, [eax]
  loc_00E38B74: lea edx, var_B4
  loc_00E38B7A: push edx
  loc_00E38B7B: push eax
  loc_00E38B7C: mov var_B8, eax
  loc_00E38B82: call [ecx+000000E0h]
  loc_00E38B88: test eax, eax
  loc_00E38B8A: fnclex
  loc_00E38B8C: jge 00E38BA6h
  loc_00E38B8E: mov ecx, var_B8
  loc_00E38B94: push 000000E0h
  loc_00E38B99: push 006E03D4h
  loc_00E38B9E: push ecx
  loc_00E38B9F: push eax
  loc_00E38BA0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38BA6: mov edx, var_B4
  loc_00E38BAC: lea ecx, var_30
  loc_00E38BAF: mov var_C0, edx
  loc_00E38BB5: call ebx
  loc_00E38BB7: cmp var_C0, 0000h
  loc_00E38BBF: jz 00E38C27h
  loc_00E38BC1: mov eax, [esi]
  loc_00E38BC3: push esi
  loc_00E38BC4: call [eax+00000324h]
  loc_00E38BCA: lea ecx, var_30
  loc_00E38BCD: push eax
  loc_00E38BCE: push ecx
  loc_00E38BCF: call edi
  loc_00E38BD1: mov edx, [eax]
  loc_00E38BD3: lea ecx, var_28
  loc_00E38BD6: push ecx
  loc_00E38BD7: push eax
  loc_00E38BD8: mov var_B8, eax
  loc_00E38BDE: call [edx+000000A0h]
  loc_00E38BE4: test eax, eax
  loc_00E38BE6: fnclex
  loc_00E38BE8: jge 00E38C02h
  loc_00E38BEA: mov edx, var_B8
  loc_00E38BF0: push 000000A0h
  loc_00E38BF5: push 006DCB70h
  loc_00E38BFA: push edx
  loc_00E38BFB: push eax
  loc_00E38BFC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38C02: mov eax, var_28
  loc_00E38C05: push 006E1C38h ; "select * from lc where metode_pembayaran like '" & Chr(37)
  loc_00E38C0A: push eax
  loc_00E38C0B: call [0040106Ch] ; __vbaStrCat
  loc_00E38C11: mov edx, eax
  loc_00E38C13: lea ecx, var_2C
  loc_00E38C16: call [00401228h] ; __vbaStrMove
  loc_00E38C1C: push eax
  loc_00E38C1D: push 006E1CA0h ; Chr(37) & "' order by metode_pembayaran asc"
  loc_00E38C22: jmp 00E38DB0h
  loc_00E38C27: mov edx, [esi]
  loc_00E38C29: push esi
  loc_00E38C2A: call [edx+00000318h]
  loc_00E38C30: push eax
  loc_00E38C31: lea eax, var_30
  loc_00E38C34: push eax
  loc_00E38C35: call edi
  loc_00E38C37: mov ecx, [eax]
  loc_00E38C39: lea edx, var_B4
  loc_00E38C3F: push edx
  loc_00E38C40: push eax
  loc_00E38C41: mov var_B8, eax
  loc_00E38C47: call [ecx+000000E0h]
  loc_00E38C4D: test eax, eax
  loc_00E38C4F: fnclex
  loc_00E38C51: jge 00E38C6Bh
  loc_00E38C53: mov ecx, var_B8
  loc_00E38C59: push 000000E0h
  loc_00E38C5E: push 006E03D4h
  loc_00E38C63: push ecx
  loc_00E38C64: push eax
  loc_00E38C65: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38C6B: mov edx, var_B4
  loc_00E38C71: lea ecx, var_30
  loc_00E38C74: mov var_C0, edx
  loc_00E38C7A: call ebx
  loc_00E38C7C: cmp var_C0, 0000h
  loc_00E38C84: jz 00E38CECh
  loc_00E38C86: mov eax, [esi]
  loc_00E38C88: push esi
  loc_00E38C89: call [eax+00000324h]
  loc_00E38C8F: lea ecx, var_30
  loc_00E38C92: push eax
  loc_00E38C93: push ecx
  loc_00E38C94: call edi
  loc_00E38C96: mov edx, [eax]
  loc_00E38C98: lea ecx, var_28
  loc_00E38C9B: push ecx
  loc_00E38C9C: push eax
  loc_00E38C9D: mov var_B8, eax
  loc_00E38CA3: call [edx+000000A0h]
  loc_00E38CA9: test eax, eax
  loc_00E38CAB: fnclex
  loc_00E38CAD: jge 00E38CC7h
  loc_00E38CAF: mov edx, var_B8
  loc_00E38CB5: push 000000A0h
  loc_00E38CBA: push 006DCB70h
  loc_00E38CBF: push edx
  loc_00E38CC0: push eax
  loc_00E38CC1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38CC7: mov eax, var_28
  loc_00E38CCA: push 006E1F60h ; "select * from lc where jurusan like '" & Chr(37)
  loc_00E38CCF: push eax
  loc_00E38CD0: call [0040106Ch] ; __vbaStrCat
  loc_00E38CD6: mov edx, eax
  loc_00E38CD8: lea ecx, var_2C
  loc_00E38CDB: call [00401228h] ; __vbaStrMove
  loc_00E38CE1: push eax
  loc_00E38CE2: push 006E1FB4h ; Chr(37) & "' order by jurusan asc"
  loc_00E38CE7: jmp 00E38DB0h
  loc_00E38CEC: mov edx, [esi]
  loc_00E38CEE: push esi
  loc_00E38CEF: call [edx+00000328h]
  loc_00E38CF5: push eax
  loc_00E38CF6: lea eax, var_30
  loc_00E38CF9: push eax
  loc_00E38CFA: call edi
  loc_00E38CFC: mov ecx, [eax]
  loc_00E38CFE: lea edx, var_B4
  loc_00E38D04: push edx
  loc_00E38D05: push eax
  loc_00E38D06: mov var_B8, eax
  loc_00E38D0C: call [ecx+000000E0h]
  loc_00E38D12: test eax, eax
  loc_00E38D14: fnclex
  loc_00E38D16: jge 00E38D30h
  loc_00E38D18: mov ecx, var_B8
  loc_00E38D1E: push 000000E0h
  loc_00E38D23: push 006E03D4h
  loc_00E38D28: push ecx
  loc_00E38D29: push eax
  loc_00E38D2A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38D30: mov edx, var_B4
  loc_00E38D36: lea ecx, var_30
  loc_00E38D39: mov var_C0, edx
  loc_00E38D3F: call ebx
  loc_00E38D41: cmp var_C0, 0000h
  loc_00E38D49: jz 00E38E59h
  loc_00E38D4F: mov eax, [esi]
  loc_00E38D51: push esi
  loc_00E38D52: call [eax+00000324h]
  loc_00E38D58: lea ecx, var_30
  loc_00E38D5B: push eax
  loc_00E38D5C: push ecx
  loc_00E38D5D: call edi
  loc_00E38D5F: mov edx, [eax]
  loc_00E38D61: lea ecx, var_28
  loc_00E38D64: push ecx
  loc_00E38D65: push eax
  loc_00E38D66: mov var_B8, eax
  loc_00E38D6C: call [edx+000000A0h]
  loc_00E38D72: test eax, eax
  loc_00E38D74: fnclex
  loc_00E38D76: jge 00E38D90h
  loc_00E38D78: mov edx, var_B8
  loc_00E38D7E: push 000000A0h
  loc_00E38D83: push 006DCB70h
  loc_00E38D88: push edx
  loc_00E38D89: push eax
  loc_00E38D8A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38D90: mov eax, var_28
  loc_00E38D93: push 006E1FE8h ; "select * from lc where agama like '" & Chr(37)
  loc_00E38D98: push eax
  loc_00E38D99: call [0040106Ch] ; __vbaStrCat
  loc_00E38D9F: mov edx, eax
  loc_00E38DA1: lea ecx, var_2C
  loc_00E38DA4: call [00401228h] ; __vbaStrMove
  loc_00E38DAA: push eax
  loc_00E38DAB: push 006E1F30h ; Chr(37) & "' order by agama asc"
  loc_00E38DB0: call [0040106Ch] ; __vbaStrCat
  loc_00E38DB6: lea edx, var_40
  loc_00E38DB9: lea ecx, var_24
  loc_00E38DBC: mov var_38, eax
  loc_00E38DBF: mov var_40, 00000008h
  loc_00E38DC6: call [0040101Ch] ; __vbaVarMove
  loc_00E38DCC: lea ecx, var_2C
  loc_00E38DCF: lea edx, var_28
  loc_00E38DD2: push ecx
  loc_00E38DD3: push edx
  loc_00E38DD4: push 00000002h
  loc_00E38DD6: call [004011BCh] ; __vbaFreeStrList
  loc_00E38DDC: add esp, 0000000Ch
  loc_00E38DDF: lea ecx, var_30
  loc_00E38DE2: call ebx
  loc_00E38DE4: lea eax, var_24
  loc_00E38DE7: push eax
  loc_00E38DE8: call [00401230h] ; __vbaStrVarCopy
  loc_00E38DEE: sub esp, 00000010h
  loc_00E38DF1: mov ecx, 00000008h
  loc_00E38DF6: mov edx, esp
  loc_00E38DF8: mov var_40, ecx
  loc_00E38DFB: mov var_38, eax
  loc_00E38DFE: push 0000000Eh
  loc_00E38E00: mov [edx], ecx
  loc_00E38E02: mov ecx, var_3C
  loc_00E38E05: push esi
  loc_00E38E06: mov [edx+00000004h], ecx
  loc_00E38E09: mov ecx, [esi]
  loc_00E38E0B: mov [edx+00000008h], eax
  loc_00E38E0E: mov eax, var_34
  loc_00E38E11: mov [edx+0000000Ch], eax
  loc_00E38E14: call [ecx+00000354h]
  loc_00E38E1A: lea edx, var_30
  loc_00E38E1D: push eax
  loc_00E38E1E: push edx
  loc_00E38E1F: call edi
  loc_00E38E21: push eax
  loc_00E38E22: call [00401238h] ; __vbaLateIdSt
  loc_00E38E28: lea ecx, var_30
  loc_00E38E2B: call ebx
  loc_00E38E2D: lea ecx, var_40
  loc_00E38E30: call [00401024h] ; __vbaFreeVar
  loc_00E38E36: mov eax, [esi]
  loc_00E38E38: push 00000000h
  loc_00E38E3A: push 00000019h
  loc_00E38E3C: push esi
  loc_00E38E3D: call [eax+00000354h]
  loc_00E38E43: lea ecx, var_30
  loc_00E38E46: push eax
  loc_00E38E47: push ecx
  loc_00E38E48: call edi
  loc_00E38E4A: push eax
  loc_00E38E4B: call [00401030h] ; __vbaLateIdCall
  loc_00E38E51: add esp, 0000000Ch
  loc_00E38E54: jmp 00E38F6Eh
  loc_00E38E59: mov edx, [esi]
  loc_00E38E5B: push 00000000h
  loc_00E38E5D: push 80010007h
  loc_00E38E62: push esi
  loc_00E38E63: call [edx+0000032Ch]
  loc_00E38E69: push eax
  loc_00E38E6A: lea eax, var_30
  loc_00E38E6D: push eax
  loc_00E38E6E: call edi
  loc_00E38E70: lea ecx, var_40
  loc_00E38E73: push eax
  loc_00E38E74: push ecx
  loc_00E38E75: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E38E7B: add esp, 00000010h
  loc_00E38E7E: push eax
  loc_00E38E7F: call [004010CCh] ; __vbaBoolVar
  loc_00E38E85: neg ax
  loc_00E38E88: sbb eax, eax
  loc_00E38E8A: lea ecx, var_30
  loc_00E38E8D: inc eax
  loc_00E38E8E: neg eax
  loc_00E38E90: mov var_B8, ax
  loc_00E38E97: call ebx
  loc_00E38E99: lea ecx, var_40
  loc_00E38E9C: call [00401024h] ; __vbaFreeVar
  loc_00E38EA2: cmp var_B8, 0000h
  loc_00E38EAA: jz 00E38F73h
  loc_00E38EB0: mov ecx, 80020004h
  loc_00E38EB5: mov eax, 0000000Ah
  loc_00E38EBA: mov var_68, ecx
  loc_00E38EBD: mov var_58, ecx
  loc_00E38EC0: lea edx, var_90
  loc_00E38EC6: lea ecx, var_50
  loc_00E38EC9: mov var_70, eax
  loc_00E38ECC: mov var_60, eax
  loc_00E38ECF: mov var_88, 006E16F0h ; "SMK RK BT PS"
  loc_00E38ED9: mov var_90, 00000008h
  loc_00E38EE3: call [004011F0h] ; __vbaVarDup
  loc_00E38EE9: lea edx, var_80
  loc_00E38EEC: lea ecx, var_40
  loc_00E38EEF: mov var_78, 006E1CE8h ; "Silahkan Pilih Item Data yang Ingin dicari !"
  loc_00E38EF6: mov var_80, 00000008h
  loc_00E38EFD: call [004011F0h] ; __vbaVarDup
  loc_00E38F03: lea edx, var_70
  loc_00E38F06: lea eax, var_60
  loc_00E38F09: push edx
  loc_00E38F0A: lea ecx, var_50
  loc_00E38F0D: push eax
  loc_00E38F0E: push ecx
  loc_00E38F0F: lea edx, var_40
  loc_00E38F12: push 00000040h
  loc_00E38F14: push edx
  loc_00E38F15: call [004010A8h] ; rtcMsgBox
  loc_00E38F1B: lea eax, var_70
  loc_00E38F1E: lea ecx, var_60
  loc_00E38F21: push eax
  loc_00E38F22: lea edx, var_50
  loc_00E38F25: push ecx
  loc_00E38F26: lea eax, var_40
  loc_00E38F29: push edx
  loc_00E38F2A: push eax
  loc_00E38F2B: push 00000004h
  loc_00E38F2D: call [00401038h] ; __vbaFreeVarList
  loc_00E38F33: mov ecx, [esi]
  loc_00E38F35: add esp, 00000014h
  loc_00E38F38: push esi
  loc_00E38F39: call [ecx+00000324h]
  loc_00E38F3F: lea edx, var_30
  loc_00E38F42: push eax
  loc_00E38F43: push edx
  loc_00E38F44: call edi
  loc_00E38F46: mov esi, eax
  loc_00E38F48: push 006DCC80h
  loc_00E38F4D: push esi
  loc_00E38F4E: mov eax, [esi]
  loc_00E38F50: call [eax+000000A4h]
  loc_00E38F56: test eax, eax
  loc_00E38F58: fnclex
  loc_00E38F5A: jge 00E38F6Eh
  loc_00E38F5C: push 000000A4h
  loc_00E38F61: push 006DCB70h
  loc_00E38F66: push esi
  loc_00E38F67: push eax
  loc_00E38F68: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E38F6E: lea ecx, var_30
  loc_00E38F71: call ebx
  loc_00E38F73: mov var_4, 00000000h
  loc_00E38F7A: push 00E38FC3h
  loc_00E38F7F: jmp 00E38FB9h
  loc_00E38F81: lea ecx, var_2C
  loc_00E38F84: lea edx, var_28
  loc_00E38F87: push ecx
  loc_00E38F88: push edx
  loc_00E38F89: push 00000002h
  loc_00E38F8B: call [004011BCh] ; __vbaFreeStrList
  loc_00E38F91: add esp, 0000000Ch
  loc_00E38F94: lea ecx, var_30
  loc_00E38F97: call [00401254h] ; __vbaFreeObj
  loc_00E38F9D: lea eax, var_70
  loc_00E38FA0: lea ecx, var_60
  loc_00E38FA3: push eax
  loc_00E38FA4: lea edx, var_50
  loc_00E38FA7: push ecx
  loc_00E38FA8: lea eax, var_40
  loc_00E38FAB: push edx
  loc_00E38FAC: push eax
  loc_00E38FAD: push 00000004h
  loc_00E38FAF: call [00401038h] ; __vbaFreeVarList
  loc_00E38FB5: add esp, 00000014h
  loc_00E38FB8: ret
  loc_00E38FB9: lea ecx, var_24
  loc_00E38FBC: call [00401024h] ; __vbaFreeVar
  loc_00E38FC2: ret
  loc_00E38FC3: mov eax, Me
  loc_00E38FC6: push eax
  loc_00E38FC7: mov ecx, [eax]
  loc_00E38FC9: call [ecx+00000008h]
  loc_00E38FCC: mov eax, var_4
  loc_00E38FCF: mov ecx, var_14
  loc_00E38FD2: pop edi
  loc_00E38FD3: pop esi
  loc_00E38FD4: mov fs:[00000000h], ecx
  loc_00E38FDB: pop ebx
  loc_00E38FDC: mov esp, ebp
  loc_00E38FDE: pop ebp
  loc_00E38FDF: retn 0008h
End Sub
