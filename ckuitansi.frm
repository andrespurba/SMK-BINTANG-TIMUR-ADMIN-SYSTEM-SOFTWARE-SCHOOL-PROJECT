VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form ckuitansi
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
  ClientWidth = 11730
  ClientHeight = 5805
  ShowInTaskbar = 0   'False
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame fcari
    Caption = "Cetak Berdasarkan"
    BackColor = &HE0E0E0&
    ForeColor = &H404040&
    Left = 150
    Top = 870
    Width = 11475
    Height = 1695
    TabIndex = 1
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.OptionButton optjur
      Caption = "Jurusan"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 6300
      Top = 540
      Width = 1515
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
    Begin VB.OptionButton optnama
      Caption = "Nama"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 2700
      Top = 510
      Width = 1455
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
    Begin VB.OptionButton optno
      Caption = "No. Daftar"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 4380
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
    Begin VB.TextBox txtcari
      BackColor = &HC0C0C0&
      Left = 240
      Top = 1110
      Width = 10905
      Height = 345
      TabIndex = 3
      BorderStyle = 0 'None
    End
    Begin VB.OptionButton optagama
      Caption = "Agama"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 7980
      Top = 570
      Width = 1695
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
    Begin Project1.jcbutton batal
      Left = 5580
      Top = 870
      Width = 675
      Height = 225
      TabIndex = 6
      OleObjectBlob = "ckuitansi.frx":0000
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
    Left = 150
    Top = 2760
    Width = 11475
    Height = 555
    TabIndex = 0
    OleObjectBlob = "ckuitansi.frx":0234
  End
  Begin MSAdodcLib.Adodc adopen
    Left = 300
    Top = 4980
    Width = 1200
    Height = 330
    Visible = 0   'False
    OleObjectBlob = "ckuitansi.frx":043C
  End
  Begin MSDataGridLib.DataGrid dgpen
    Left = 150
    Top = 3450
    Width = 11475
    Height = 2265
    TabIndex = 7
    OleObjectBlob = "ckuitansi.frx":076E
  End
  Begin VB.Shape Shape4
    Left = -30
    Top = 510
    Width = 14685
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Shape Shape3
    Left = -30
    Top = 450
    Width = 14685
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Label Label4
    Caption = "TANDA PEMBAYARAN"
    ForeColor = &H80FF80&
    Left = 4890
    Top = 60
    Width = 2535
    Height = 315
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
  Begin VB.Image back
    Picture = "ckuitansi.frx":0919
    Left = 0
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Label Label1
    Caption = "-BIODATA-"
    ForeColor = &H4000&
    Left = 5520
    Top = 510
    Width = 1305
    Height = 315
    TabIndex = 8
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
    Left = -90
    Top = 2640
    Width = 15015
    Height = 750
    BorderStyle = 3 'Dot
    FillColor = &HFFFF&
  End
  Begin VB.Shape Shape1
    Left = -30
    Top = 0
    Width = 14685
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
End

Attribute VB_Name = "ckuitansi"


Private Sub back_Click() 'E38FF0
  loc_00E38FF0: push ebp
  loc_00E38FF1: mov ebp, esp
  loc_00E38FF3: sub esp, 0000000Ch
  loc_00E38FF6: push 00402836h ; __vbaExceptHandler
  loc_00E38FFB: mov eax, fs:[00000000h]
  loc_00E39001: push eax
  loc_00E39002: mov fs:[00000000h], esp
  loc_00E39009: sub esp, 0000003Ch
  loc_00E3900C: push ebx
  loc_00E3900D: push esi
  loc_00E3900E: push edi
  loc_00E3900F: mov var_C, esp
  loc_00E39012: mov var_8, 00402790h
  loc_00E39019: mov eax, Me
  loc_00E3901C: mov ecx, eax
  loc_00E3901E: and ecx, 00000001h
  loc_00E39021: mov var_4, ecx
  loc_00E39024: and al, FEh
  loc_00E39026: push eax
  loc_00E39027: mov Me, eax
  loc_00E3902A: mov edx, [eax]
  loc_00E3902C: call [edx+00000004h]
  loc_00E3902F: mov eax, [00E3D024h]
  loc_00E39034: mov var_18, 00000000h
  loc_00E3903B: test eax, eax
  loc_00E3903D: jnz 00E3904Fh
  loc_00E3903F: push 00E3D024h
  loc_00E39044: push 006CE120h
  loc_00E39049: call [004011A0h] ; __vbaNew2
  loc_00E3904F: sub esp, 00000010h
  loc_00E39052: mov ecx, 0000000Ah
  loc_00E39057: mov esi, esp
  loc_00E39059: mov var_28, ecx
  loc_00E3905C: mov eax, 80020004h
  loc_00E39061: mov ebx, [00E3D024h]
  loc_00E39067: mov [esi], ecx
  loc_00E39069: mov ecx, var_34
  loc_00E3906C: mov edx, eax
  loc_00E3906E: sub esp, 00000010h
  loc_00E39071: mov [esi+00000004h], ecx
  loc_00E39074: mov edi, [ebx]
  loc_00E39076: mov ecx, esp
  loc_00E39078: mov var_4C, edi
  loc_00E3907B: mov [esi+00000008h], eax
  loc_00E3907E: mov eax, var_2C
  loc_00E39081: mov edi, var_1C
  loc_00E39084: push ebx
  loc_00E39085: mov [esi+0000000Ch], eax
  loc_00E39088: mov eax, var_28
  loc_00E3908B: mov esi, var_24
  loc_00E3908E: mov [ecx], eax
  loc_00E39090: mov [ecx+00000004h], esi
  loc_00E39093: mov [ecx+00000008h], edx
  loc_00E39096: mov [ecx+0000000Ch], edi
  loc_00E39099: mov ecx, var_4C
  loc_00E3909C: call [ecx+000002B0h]
  loc_00E390A2: test eax, eax
  loc_00E390A4: fnclex
  loc_00E390A6: jge 00E390BAh
  loc_00E390A8: push 000002B0h
  loc_00E390AD: push 006DC970h
  loc_00E390B2: push ebx
  loc_00E390B3: push eax
  loc_00E390B4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E390BA: mov eax, [00E3D9CCh]
  loc_00E390BF: test eax, eax
  loc_00E390C1: jnz 00E390D3h
  loc_00E390C3: push 00E3D9CCh
  loc_00E390C8: push 006DC960h
  loc_00E390CD: call [004011A0h] ; __vbaNew2
  loc_00E390D3: mov eax, [00E3D9CCh]
  loc_00E390D8: mov edx, Me
  loc_00E390DB: mov var_3C, eax
  loc_00E390DE: push edx
  loc_00E390DF: mov ebx, [eax]
  loc_00E390E1: lea eax, var_18
  loc_00E390E4: push eax
  loc_00E390E5: call [004010B4h] ; __vbaObjSetAddref
  loc_00E390EB: mov ecx, ebx
  loc_00E390ED: mov ebx, var_3C
  loc_00E390F0: push eax
  loc_00E390F1: push ebx
  loc_00E390F2: call [ecx+00000010h]
  loc_00E390F5: test eax, eax
  loc_00E390F7: fnclex
  loc_00E390F9: jge 00E3910Ah
  loc_00E390FB: push 00000010h
  loc_00E390FD: push 006DC950h
  loc_00E39102: push ebx
  loc_00E39103: push eax
  loc_00E39104: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3910A: lea ecx, var_18
  loc_00E3910D: call [00401254h] ; __vbaFreeObj
  loc_00E39113: mov eax, [00E3D024h]
  loc_00E39118: test eax, eax
  loc_00E3911A: jnz 00E3912Ch
  loc_00E3911C: push 00E3D024h
  loc_00E39121: push 006CE120h
  loc_00E39126: call [004011A0h] ; __vbaNew2
  loc_00E3912C: sub esp, 00000010h
  loc_00E3912F: mov eax, 0000000Ah
  loc_00E39134: mov ecx, esp
  loc_00E39136: mov var_28, eax
  loc_00E39139: sub esp, 00000010h
  loc_00E3913C: mov ebx, [00E3D024h]
  loc_00E39142: mov [ecx], eax
  loc_00E39144: mov eax, var_34
  loc_00E39147: mov edx, [ebx]
  loc_00E39149: mov var_20, 80020004h
  loc_00E39150: mov [ecx+00000004h], eax
  loc_00E39153: mov eax, 80020004h
  loc_00E39158: mov [ecx+00000008h], eax
  loc_00E3915B: mov eax, var_2C
  loc_00E3915E: mov [ecx+0000000Ch], eax
  loc_00E39161: mov eax, var_28
  loc_00E39164: mov ecx, esp
  loc_00E39166: push ebx
  loc_00E39167: mov [ecx], eax
  loc_00E39169: mov eax, var_20
  loc_00E3916C: mov [ecx+00000004h], esi
  loc_00E3916F: mov [ecx+00000008h], eax
  loc_00E39172: mov [ecx+0000000Ch], edi
  loc_00E39175: call [edx+000002B0h]
  loc_00E3917B: test eax, eax
  loc_00E3917D: fnclex
  loc_00E3917F: jge 00E39193h
  loc_00E39181: push 000002B0h
  loc_00E39186: push 006DC970h
  loc_00E3918B: push ebx
  loc_00E3918C: push eax
  loc_00E3918D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39193: mov eax, [00E3D024h]
  loc_00E39198: or ebx, FFFFFFFFh
  loc_00E3919B: test eax, eax
  loc_00E3919D: jnz 00E391B4h
  loc_00E3919F: push 00E3D024h
  loc_00E391A4: push 006CE120h
  loc_00E391A9: call [004011A0h] ; __vbaNew2
  loc_00E391AF: mov eax, [00E3D024h]
  loc_00E391B4: sub esp, 00000010h
  loc_00E391B7: mov ecx, 0000000Bh
  loc_00E391BC: mov edx, esp
  loc_00E391BE: push 80010007h
  loc_00E391C3: push eax
  loc_00E391C4: mov [edx], ecx
  loc_00E391C6: mov ecx, [eax]
  loc_00E391C8: mov [edx+00000004h], esi
  loc_00E391CB: mov [edx+00000008h], ebx
  loc_00E391CE: mov [edx+0000000Ch], edi
  loc_00E391D1: call [ecx+00000318h]
  loc_00E391D7: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E391DD: lea edx, var_18
  loc_00E391E0: push eax
  loc_00E391E1: push edx
  loc_00E391E2: call ebx
  loc_00E391E4: push eax
  loc_00E391E5: call [00401238h] ; __vbaLateIdSt
  loc_00E391EB: lea ecx, var_18
  loc_00E391EE: call [00401254h] ; __vbaFreeObj
  loc_00E391F4: mov eax, [00E3D024h]
  loc_00E391F9: test eax, eax
  loc_00E391FB: jnz 00E39212h
  loc_00E391FD: push 00E3D024h
  loc_00E39202: push 006CE120h
  loc_00E39207: call [004011A0h] ; __vbaNew2
  loc_00E3920D: mov eax, [00E3D024h]
  loc_00E39212: sub esp, 00000010h
  loc_00E39215: mov ecx, 0000000Bh
  loc_00E3921A: mov edx, esp
  loc_00E3921C: push 80010007h
  loc_00E39221: push eax
  loc_00E39222: mov [edx], ecx
  loc_00E39224: or ecx, FFFFFFFFh
  loc_00E39227: mov [edx+00000004h], esi
  loc_00E3922A: mov [edx+00000008h], ecx
  loc_00E3922D: mov ecx, [eax]
  loc_00E3922F: mov [edx+0000000Ch], edi
  loc_00E39232: call [ecx+0000031Ch]
  loc_00E39238: lea edx, var_18
  loc_00E3923B: push eax
  loc_00E3923C: push edx
  loc_00E3923D: call ebx
  loc_00E3923F: push eax
  loc_00E39240: call [00401238h] ; __vbaLateIdSt
  loc_00E39246: lea ecx, var_18
  loc_00E39249: call [00401254h] ; __vbaFreeObj
  loc_00E3924F: mov eax, [00E3D024h]
  loc_00E39254: mov var_20, FFFFFFFFh
  loc_00E3925B: test eax, eax
  loc_00E3925D: jnz 00E39274h
  loc_00E3925F: push 00E3D024h
  loc_00E39264: push 006CE120h
  loc_00E39269: call [004011A0h] ; __vbaNew2
  loc_00E3926F: mov eax, [00E3D024h]
  loc_00E39274: sub esp, 00000010h
  loc_00E39277: mov ecx, 0000000Bh
  loc_00E3927C: mov edx, esp
  loc_00E3927E: push 80010007h
  loc_00E39283: push eax
  loc_00E39284: mov [edx], ecx
  loc_00E39286: mov ecx, var_20
  loc_00E39289: mov [edx+00000004h], esi
  loc_00E3928C: mov [edx+00000008h], ecx
  loc_00E3928F: mov [edx+0000000Ch], edi
  loc_00E39292: mov edx, [eax]
  loc_00E39294: call [edx+00000320h]
  loc_00E3929A: push eax
  loc_00E3929B: lea eax, var_18
  loc_00E3929E: push eax
  loc_00E3929F: call ebx
  loc_00E392A1: push eax
  loc_00E392A2: call [00401238h] ; __vbaLateIdSt
  loc_00E392A8: lea ecx, var_18
  loc_00E392AB: call [00401254h] ; __vbaFreeObj
  loc_00E392B1: mov eax, [00E3D024h]
  loc_00E392B6: test eax, eax
  loc_00E392B8: jnz 00E392CFh
  loc_00E392BA: push 00E3D024h
  loc_00E392BF: push 006CE120h
  loc_00E392C4: call [004011A0h] ; __vbaNew2
  loc_00E392CA: mov eax, [00E3D024h]
  loc_00E392CF: sub esp, 00000010h
  loc_00E392D2: mov ecx, 00000008h
  loc_00E392D7: mov edx, esp
  loc_00E392D9: push FFFFFDFAh
  loc_00E392DE: push eax
  loc_00E392DF: mov [edx], ecx
  loc_00E392E1: mov ecx, 006DCDB4h ; " A D M I N "
  loc_00E392E6: mov [edx+00000004h], esi
  loc_00E392E9: mov [edx+00000008h], ecx
  loc_00E392EC: mov ecx, [eax]
  loc_00E392EE: mov [edx+0000000Ch], edi
  loc_00E392F1: call [ecx+00000324h]
  loc_00E392F7: lea edx, var_18
  loc_00E392FA: push eax
  loc_00E392FB: push edx
  loc_00E392FC: call ebx
  loc_00E392FE: push eax
  loc_00E392FF: call [00401238h] ; __vbaLateIdSt
  loc_00E39305: lea ecx, var_18
  loc_00E39308: call [00401254h] ; __vbaFreeObj
  loc_00E3930E: mov eax, [00E3D024h]
  loc_00E39313: mov var_28, 0000000Bh
  loc_00E3931A: test eax, eax
  loc_00E3931C: jnz 00E39333h
  loc_00E3931E: push 00E3D024h
  loc_00E39323: push 006CE120h
  loc_00E39328: call [004011A0h] ; __vbaNew2
  loc_00E3932E: mov eax, [00E3D024h]
  loc_00E39333: mov ecx, var_28
  loc_00E39336: sub esp, 00000010h
  loc_00E39339: mov edx, esp
  loc_00E3933B: push 8001000Dh
  loc_00E39340: push eax
  loc_00E39341: mov [edx], ecx
  loc_00E39343: xor ecx, ecx
  loc_00E39345: mov [edx+00000004h], esi
  loc_00E39348: mov [edx+00000008h], ecx
  loc_00E3934B: mov [edx+0000000Ch], edi
  loc_00E3934E: mov edx, [eax]
  loc_00E39350: call [edx+00000324h]
  loc_00E39356: push eax
  loc_00E39357: lea eax, var_18
  loc_00E3935A: push eax
  loc_00E3935B: call ebx
  loc_00E3935D: push eax
  loc_00E3935E: call [00401238h] ; __vbaLateIdSt
  loc_00E39364: lea ecx, var_18
  loc_00E39367: call [00401254h] ; __vbaFreeObj
  loc_00E3936D: mov eax, [00E3D024h]
  loc_00E39372: test eax, eax
  loc_00E39374: jnz 00E3938Bh
  loc_00E39376: push 00E3D024h
  loc_00E3937B: push 006CE120h
  loc_00E39380: call [004011A0h] ; __vbaNew2
  loc_00E39386: mov eax, [00E3D024h]
  loc_00E3938B: sub esp, 00000010h
  loc_00E3938E: mov ecx, 00000003h
  loc_00E39393: mov edx, esp
  loc_00E39395: push FFFFFE0Bh
  loc_00E3939A: push eax
  loc_00E3939B: mov [edx], ecx
  loc_00E3939D: mov ecx, 00404040h
  loc_00E393A2: mov [edx+00000004h], esi
  loc_00E393A5: mov [edx+00000008h], ecx
  loc_00E393A8: mov ecx, [eax]
  loc_00E393AA: mov [edx+0000000Ch], edi
  loc_00E393AD: call [ecx+00000324h]
  loc_00E393B3: lea edx, var_18
  loc_00E393B6: push eax
  loc_00E393B7: push edx
  loc_00E393B8: call ebx
  loc_00E393BA: push eax
  loc_00E393BB: call [00401238h] ; __vbaLateIdSt
  loc_00E393C1: lea ecx, var_18
  loc_00E393C4: call [00401254h] ; __vbaFreeObj
  loc_00E393CA: mov eax, [00E3D024h]
  loc_00E393CF: mov var_28, 00000003h
  loc_00E393D6: test eax, eax
  loc_00E393D8: jnz 00E393EFh
  loc_00E393DA: push 00E3D024h
  loc_00E393DF: push 006CE120h
  loc_00E393E4: call [004011A0h] ; __vbaNew2
  loc_00E393EA: mov eax, [00E3D024h]
  loc_00E393EF: mov ecx, var_28
  loc_00E393F2: sub esp, 00000010h
  loc_00E393F5: mov edx, esp
  loc_00E393F7: push FFFFFDFFh
  loc_00E393FC: push eax
  loc_00E393FD: mov [edx], ecx
  loc_00E393FF: mov ecx, 00E0E0E0h
  loc_00E39404: mov [edx+00000004h], esi
  loc_00E39407: mov [edx+00000008h], ecx
  loc_00E3940A: mov [edx+0000000Ch], edi
  loc_00E3940D: mov edx, [eax]
  loc_00E3940F: call [edx+00000324h]
  loc_00E39415: push eax
  loc_00E39416: lea eax, var_18
  loc_00E39419: push eax
  loc_00E3941A: call ebx
  loc_00E3941C: push eax
  loc_00E3941D: call [00401238h] ; __vbaLateIdSt
  loc_00E39423: lea ecx, var_18
  loc_00E39426: call [00401254h] ; __vbaFreeObj
  loc_00E3942C: mov var_4, 00000000h
  loc_00E39433: push 00E39445h
  loc_00E39438: jmp 00E39444h
  loc_00E3943A: lea ecx, var_18
  loc_00E3943D: call [00401254h] ; __vbaFreeObj
  loc_00E39443: ret
  loc_00E39444: ret
  loc_00E39445: mov eax, Me
  loc_00E39448: push eax
  loc_00E39449: mov ecx, [eax]
  loc_00E3944B: call [ecx+00000008h]
  loc_00E3944E: mov eax, var_4
  loc_00E39451: mov ecx, var_14
  loc_00E39454: pop edi
  loc_00E39455: pop esi
  loc_00E39456: mov fs:[00000000h], ecx
  loc_00E3945D: pop ebx
  loc_00E3945E: mov esp, ebp
  loc_00E39460: pop ebp
  loc_00E39461: retn 0004h
End Sub

Private Sub optnama_Click() 'E3B470
  loc_00E3B470: push ebp
  loc_00E3B471: mov ebp, esp
  loc_00E3B473: sub esp, 0000000Ch
  loc_00E3B476: push 00402836h ; __vbaExceptHandler
  loc_00E3B47B: mov eax, fs:[00000000h]
  loc_00E3B481: push eax
  loc_00E3B482: mov fs:[00000000h], esp
  loc_00E3B489: sub esp, 00000048h
  loc_00E3B48C: push ebx
  loc_00E3B48D: push esi
  loc_00E3B48E: push edi
  loc_00E3B48F: mov var_C, esp
  loc_00E3B492: mov var_8, 00402810h
  loc_00E3B499: mov esi, Me
  loc_00E3B49C: mov eax, esi
  loc_00E3B49E: and eax, 00000001h
  loc_00E3B4A1: mov var_4, eax
  loc_00E3B4A4: and esi, FFFFFFFEh
  loc_00E3B4A7: push esi
  loc_00E3B4A8: mov Me, esi
  loc_00E3B4AB: mov ecx, [esi]
  loc_00E3B4AD: call [ecx+00000004h]
  loc_00E3B4B0: sub esp, 00000010h
  loc_00E3B4B3: mov ecx, 0000000Bh
  loc_00E3B4B8: mov edx, esp
  loc_00E3B4BA: xor eax, eax
  loc_00E3B4BC: mov var_18, eax
  loc_00E3B4BF: mov var_1C, eax
  loc_00E3B4C2: mov [edx], ecx
  loc_00E3B4C4: mov ecx, var_38
  loc_00E3B4C7: mov var_2C, eax
  loc_00E3B4CA: or eax, FFFFFFFFh
  loc_00E3B4CD: mov [edx+00000004h], ecx
  loc_00E3B4D0: mov ecx, [esi]
  loc_00E3B4D2: push 80010007h
  loc_00E3B4D7: push esi
  loc_00E3B4D8: mov [edx+00000008h], eax
  loc_00E3B4DB: mov eax, var_30
  loc_00E3B4DE: mov [edx+0000000Ch], eax
  loc_00E3B4E1: call [ecx+00000314h]
  loc_00E3B4E7: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3B4ED: lea edx, var_18
  loc_00E3B4F0: push eax
  loc_00E3B4F1: push edx
  loc_00E3B4F2: call edi
  loc_00E3B4F4: push eax
  loc_00E3B4F5: call [00401238h] ; __vbaLateIdSt
  loc_00E3B4FB: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E3B501: lea ecx, var_18
  loc_00E3B504: call ebx
  loc_00E3B506: mov eax, [esi]
  loc_00E3B508: push esi
  loc_00E3B509: call [eax+0000030Ch]
  loc_00E3B50F: lea ecx, var_18
  loc_00E3B512: push eax
  loc_00E3B513: push ecx
  loc_00E3B514: call edi
  loc_00E3B516: mov edx, [eax]
  loc_00E3B518: push eax
  loc_00E3B519: mov var_50, eax
  loc_00E3B51C: call [edx+00000204h]
  loc_00E3B522: test eax, eax
  loc_00E3B524: fnclex
  loc_00E3B526: jge 00E3B53Dh
  loc_00E3B528: mov ecx, var_50
  loc_00E3B52B: push 00000204h
  loc_00E3B530: push 006DCB70h
  loc_00E3B535: push ecx
  loc_00E3B536: push eax
  loc_00E3B537: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3B53D: lea ecx, var_18
  loc_00E3B540: call ebx
  loc_00E3B542: mov edx, [esi]
  loc_00E3B544: push esi
  loc_00E3B545: call [edx+00000308h]
  loc_00E3B54B: push eax
  loc_00E3B54C: lea eax, var_18
  loc_00E3B54F: push eax
  loc_00E3B550: call edi
  loc_00E3B552: mov ecx, [eax]
  loc_00E3B554: push 00000000h
  loc_00E3B556: push eax
  loc_00E3B557: mov var_50, eax
  loc_00E3B55A: call [ecx+0000009Ch]
  loc_00E3B560: test eax, eax
  loc_00E3B562: fnclex
  loc_00E3B564: jge 00E3B57Bh
  loc_00E3B566: mov edx, var_50
  loc_00E3B569: push 0000009Ch
  loc_00E3B56E: push 006E03D4h
  loc_00E3B573: push edx
  loc_00E3B574: push eax
  loc_00E3B575: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3B57B: lea ecx, var_18
  loc_00E3B57E: call ebx
  loc_00E3B580: mov eax, [esi]
  loc_00E3B582: push esi
  loc_00E3B583: call [eax+00000310h]
  loc_00E3B589: lea ecx, var_18
  loc_00E3B58C: push eax
  loc_00E3B58D: push ecx
  loc_00E3B58E: call edi
  loc_00E3B590: mov edx, [eax]
  loc_00E3B592: push 00000000h
  loc_00E3B594: push eax
  loc_00E3B595: mov var_50, eax
  loc_00E3B598: call [edx+0000009Ch]
  loc_00E3B59E: test eax, eax
  loc_00E3B5A0: fnclex
  loc_00E3B5A2: jge 00E3B5B9h
  loc_00E3B5A4: mov ecx, var_50
  loc_00E3B5A7: push 0000009Ch
  loc_00E3B5AC: push 006E03D4h
  loc_00E3B5B1: push ecx
  loc_00E3B5B2: push eax
  loc_00E3B5B3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3B5B9: lea ecx, var_18
  loc_00E3B5BC: call ebx
  loc_00E3B5BE: mov edx, [esi]
  loc_00E3B5C0: push esi
  loc_00E3B5C1: call [edx+00000300h]
  loc_00E3B5C7: push eax
  loc_00E3B5C8: lea eax, var_18
  loc_00E3B5CB: push eax
  loc_00E3B5CC: call edi
  loc_00E3B5CE: mov ecx, [eax]
  loc_00E3B5D0: push 00000000h
  loc_00E3B5D2: push eax
  loc_00E3B5D3: mov var_50, eax
  loc_00E3B5D6: call [ecx+0000009Ch]
  loc_00E3B5DC: test eax, eax
  loc_00E3B5DE: fnclex
  loc_00E3B5E0: jge 00E3B5F7h
  loc_00E3B5E2: mov edx, var_50
  loc_00E3B5E5: push 0000009Ch
  loc_00E3B5EA: push 006E03D4h
  loc_00E3B5EF: push edx
  loc_00E3B5F0: push eax
  loc_00E3B5F1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3B5F7: lea ecx, var_18
  loc_00E3B5FA: call ebx
  loc_00E3B5FC: sub esp, 00000010h
  loc_00E3B5FF: mov ecx, 00000008h
  loc_00E3B604: mov edx, esp
  loc_00E3B606: mov eax, 006E1738h ; "lc"
  loc_00E3B60B: push 0000000Eh
  loc_00E3B60D: push esi
  loc_00E3B60E: mov [edx], ecx
  loc_00E3B610: mov ecx, var_38
  loc_00E3B613: mov [edx+00000004h], ecx
  loc_00E3B616: mov ecx, [esi]
  loc_00E3B618: mov [edx+00000008h], eax
  loc_00E3B61B: mov eax, var_30
  loc_00E3B61E: mov [edx+0000000Ch], eax
  loc_00E3B621: call [ecx+0000033Ch]
  loc_00E3B627: lea edx, var_18
  loc_00E3B62A: push eax
  loc_00E3B62B: push edx
  loc_00E3B62C: call edi
  loc_00E3B62E: push eax
  loc_00E3B62F: call [00401238h] ; __vbaLateIdSt
  loc_00E3B635: lea ecx, var_18
  loc_00E3B638: call ebx
  loc_00E3B63A: mov eax, [esi]
  loc_00E3B63C: push 00000000h
  loc_00E3B63E: push 00000019h
  loc_00E3B640: push esi
  loc_00E3B641: call [eax+0000033Ch]
  loc_00E3B647: lea ecx, var_18
  loc_00E3B64A: push eax
  loc_00E3B64B: push ecx
  loc_00E3B64C: call edi
  loc_00E3B64E: push eax
  loc_00E3B64F: call [00401030h] ; __vbaLateIdCall
  loc_00E3B655: add esp, 0000000Ch
  loc_00E3B658: lea ecx, var_18
  loc_00E3B65B: call ebx
  loc_00E3B65D: mov edx, [esi]
  loc_00E3B65F: push 006E05D8h
  loc_00E3B664: push esi
  loc_00E3B665: call [edx+0000033Ch]
  loc_00E3B66B: push eax
  loc_00E3B66C: lea eax, var_18
  loc_00E3B66F: push eax
  loc_00E3B670: call edi
  loc_00E3B672: push eax
  loc_00E3B673: call [00401224h] ; __vbaCastObj
  loc_00E3B679: lea ecx, var_2C
  loc_00E3B67C: mov var_24, eax
  loc_00E3B67F: push ecx
  loc_00E3B680: mov var_2C, 0000000Dh
  loc_00E3B687: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3B68D: mov ecx, [eax]
  loc_00E3B68F: sub esp, 00000010h
  loc_00E3B692: mov edx, esp
  loc_00E3B694: push 00000000h
  loc_00E3B696: push 0000002Ah
  loc_00E3B698: mov [edx], ecx
  loc_00E3B69A: mov ecx, [eax+00000004h]
  loc_00E3B69D: push esi
  loc_00E3B69E: mov [edx+00000004h], ecx
  loc_00E3B6A1: mov ecx, [eax+00000008h]
  loc_00E3B6A4: mov eax, [eax+0000000Ch]
  loc_00E3B6A7: mov [edx+00000008h], ecx
  loc_00E3B6AA: mov ecx, [esi]
  loc_00E3B6AC: mov [edx+0000000Ch], eax
  loc_00E3B6AF: call [ecx+00000340h]
  loc_00E3B6B5: lea edx, var_1C
  loc_00E3B6B8: push eax
  loc_00E3B6B9: push edx
  loc_00E3B6BA: call edi
  loc_00E3B6BC: push eax
  loc_00E3B6BD: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3B6C3: lea eax, var_1C
  loc_00E3B6C6: lea ecx, var_18
  loc_00E3B6C9: push eax
  loc_00E3B6CA: push ecx
  loc_00E3B6CB: push 00000002h
  loc_00E3B6CD: call [00401048h] ; __vbaFreeObjList
  loc_00E3B6D3: add esp, 00000028h
  loc_00E3B6D6: lea ecx, var_2C
  loc_00E3B6D9: call [00401024h] ; __vbaFreeVar
  loc_00E3B6DF: sub esp, 00000010h
  loc_00E3B6E2: mov ecx, 0000000Bh
  loc_00E3B6E7: mov edx, esp
  loc_00E3B6E9: xor eax, eax
  loc_00E3B6EB: push 00000006h
  loc_00E3B6ED: push esi
  loc_00E3B6EE: mov [edx], ecx
  loc_00E3B6F0: mov ecx, var_38
  loc_00E3B6F3: mov [edx+00000004h], ecx
  loc_00E3B6F6: mov ecx, [esi]
  loc_00E3B6F8: mov [edx+00000008h], eax
  loc_00E3B6FB: mov eax, var_30
  loc_00E3B6FE: mov [edx+0000000Ch], eax
  loc_00E3B701: call [ecx+00000340h]
  loc_00E3B707: lea edx, var_18
  loc_00E3B70A: push eax
  loc_00E3B70B: push edx
  loc_00E3B70C: call edi
  loc_00E3B70E: push eax
  loc_00E3B70F: call [00401238h] ; __vbaLateIdSt
  loc_00E3B715: lea ecx, var_18
  loc_00E3B718: call ebx
  loc_00E3B71A: sub esp, 00000010h
  loc_00E3B71D: mov ecx, 0000000Bh
  loc_00E3B722: mov edx, esp
  loc_00E3B724: xor eax, eax
  loc_00E3B726: push 8001000Eh
  loc_00E3B72B: push esi
  loc_00E3B72C: mov [edx], ecx
  loc_00E3B72E: mov ecx, var_38
  loc_00E3B731: mov [edx+00000004h], ecx
  loc_00E3B734: mov ecx, [esi]
  loc_00E3B736: mov [edx+00000008h], eax
  loc_00E3B739: mov eax, var_30
  loc_00E3B73C: mov [edx+0000000Ch], eax
  loc_00E3B73F: call [ecx+00000340h]
  loc_00E3B745: lea edx, var_18
  loc_00E3B748: push eax
  loc_00E3B749: push edx
  loc_00E3B74A: call edi
  loc_00E3B74C: push eax
  loc_00E3B74D: call [00401238h] ; __vbaLateIdSt
  loc_00E3B753: lea ecx, var_18
  loc_00E3B756: call ebx
  loc_00E3B758: mov eax, [esi]
  loc_00E3B75A: push 00000000h
  loc_00E3B75C: push FFFFFDDAh
  loc_00E3B761: push esi
  loc_00E3B762: call [eax+00000340h]
  loc_00E3B768: lea ecx, var_18
  loc_00E3B76B: push eax
  loc_00E3B76C: push ecx
  loc_00E3B76D: call edi
  loc_00E3B76F: push eax
  loc_00E3B770: call [00401030h] ; __vbaLateIdCall
  loc_00E3B776: add esp, 0000000Ch
  loc_00E3B779: lea ecx, var_18
  loc_00E3B77C: call ebx
  loc_00E3B77E: mov var_4, 00000000h
  loc_00E3B785: push 00E3B7AAh
  loc_00E3B78A: jmp 00E3B7A9h
  loc_00E3B78C: lea edx, var_1C
  loc_00E3B78F: lea eax, var_18
  loc_00E3B792: push edx
  loc_00E3B793: push eax
  loc_00E3B794: push 00000002h
  loc_00E3B796: call [00401048h] ; __vbaFreeObjList
  loc_00E3B79C: add esp, 0000000Ch
  loc_00E3B79F: lea ecx, var_2C
  loc_00E3B7A2: call [00401024h] ; __vbaFreeVar
  loc_00E3B7A8: ret
  loc_00E3B7A9: ret
  loc_00E3B7AA: mov eax, Me
  loc_00E3B7AD: push eax
  loc_00E3B7AE: mov ecx, [eax]
  loc_00E3B7B0: call [ecx+00000008h]
  loc_00E3B7B3: mov eax, var_4
  loc_00E3B7B6: mov ecx, var_14
  loc_00E3B7B9: pop edi
  loc_00E3B7BA: pop esi
  loc_00E3B7BB: mov fs:[00000000h], ecx
  loc_00E3B7C2: pop ebx
  loc_00E3B7C3: mov esp, ebp
  loc_00E3B7C5: pop ebp
  loc_00E3B7C6: retn 0004h
End Sub

Private Sub optno_Click() 'E3B7D0
  loc_00E3B7D0: push ebp
  loc_00E3B7D1: mov ebp, esp
  loc_00E3B7D3: sub esp, 0000000Ch
  loc_00E3B7D6: push 00402836h ; __vbaExceptHandler
  loc_00E3B7DB: mov eax, fs:[00000000h]
  loc_00E3B7E1: push eax
  loc_00E3B7E2: mov fs:[00000000h], esp
  loc_00E3B7E9: sub esp, 00000048h
  loc_00E3B7EC: push ebx
  loc_00E3B7ED: push esi
  loc_00E3B7EE: push edi
  loc_00E3B7EF: mov var_C, esp
  loc_00E3B7F2: mov var_8, 00402820h
  loc_00E3B7F9: mov esi, Me
  loc_00E3B7FC: mov eax, esi
  loc_00E3B7FE: and eax, 00000001h
  loc_00E3B801: mov var_4, eax
  loc_00E3B804: and esi, FFFFFFFEh
  loc_00E3B807: push esi
  loc_00E3B808: mov Me, esi
  loc_00E3B80B: mov ecx, [esi]
  loc_00E3B80D: call [ecx+00000004h]
  loc_00E3B810: sub esp, 00000010h
  loc_00E3B813: mov ecx, 00000008h
  loc_00E3B818: mov edx, esp
  loc_00E3B81A: xor eax, eax
  loc_00E3B81C: mov var_18, eax
  loc_00E3B81F: mov var_1C, eax
  loc_00E3B822: mov [edx], ecx
  loc_00E3B824: mov ecx, var_38
  loc_00E3B827: mov var_2C, eax
  loc_00E3B82A: mov eax, 006E1738h ; "lc"
  loc_00E3B82F: mov [edx+00000004h], ecx
  loc_00E3B832: mov ecx, [esi]
  loc_00E3B834: push 0000000Eh
  loc_00E3B836: push esi
  loc_00E3B837: mov [edx+00000008h], eax
  loc_00E3B83A: mov eax, var_30
  loc_00E3B83D: mov [edx+0000000Ch], eax
  loc_00E3B840: call [ecx+0000033Ch]
  loc_00E3B846: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3B84C: lea edx, var_18
  loc_00E3B84F: push eax
  loc_00E3B850: push edx
  loc_00E3B851: call edi
  loc_00E3B853: push eax
  loc_00E3B854: call [00401238h] ; __vbaLateIdSt
  loc_00E3B85A: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E3B860: lea ecx, var_18
  loc_00E3B863: call ebx
  loc_00E3B865: mov eax, [esi]
  loc_00E3B867: push 00000000h
  loc_00E3B869: push 00000019h
  loc_00E3B86B: push esi
  loc_00E3B86C: call [eax+0000033Ch]
  loc_00E3B872: lea ecx, var_18
  loc_00E3B875: push eax
  loc_00E3B876: push ecx
  loc_00E3B877: call edi
  loc_00E3B879: push eax
  loc_00E3B87A: call [00401030h] ; __vbaLateIdCall
  loc_00E3B880: add esp, 0000000Ch
  loc_00E3B883: lea ecx, var_18
  loc_00E3B886: call ebx
  loc_00E3B888: mov edx, [esi]
  loc_00E3B88A: push 006E05D8h
  loc_00E3B88F: push esi
  loc_00E3B890: call [edx+0000033Ch]
  loc_00E3B896: push eax
  loc_00E3B897: lea eax, var_18
  loc_00E3B89A: push eax
  loc_00E3B89B: call edi
  loc_00E3B89D: push eax
  loc_00E3B89E: call [00401224h] ; __vbaCastObj
  loc_00E3B8A4: lea ecx, var_2C
  loc_00E3B8A7: mov var_24, eax
  loc_00E3B8AA: push ecx
  loc_00E3B8AB: mov var_2C, 0000000Dh
  loc_00E3B8B2: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3B8B8: mov ecx, [eax]
  loc_00E3B8BA: sub esp, 00000010h
  loc_00E3B8BD: mov edx, esp
  loc_00E3B8BF: mov [edx], ecx
  loc_00E3B8C1: mov ecx, [eax+00000004h]
  loc_00E3B8C4: mov [edx+00000004h], ecx
  loc_00E3B8C7: mov ecx, [eax+00000008h]
  loc_00E3B8CA: mov eax, [eax+0000000Ch]
  loc_00E3B8CD: mov [edx+00000008h], ecx
  loc_00E3B8D0: mov [edx+0000000Ch], eax
  loc_00E3B8D3: mov ecx, [esi]
  loc_00E3B8D5: push 00000000h
  loc_00E3B8D7: push 0000002Ah
  loc_00E3B8D9: push esi
  loc_00E3B8DA: call [ecx+00000340h]
  loc_00E3B8E0: lea edx, var_1C
  loc_00E3B8E3: push eax
  loc_00E3B8E4: push edx
  loc_00E3B8E5: call edi
  loc_00E3B8E7: push eax
  loc_00E3B8E8: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3B8EE: lea eax, var_1C
  loc_00E3B8F1: lea ecx, var_18
  loc_00E3B8F4: push eax
  loc_00E3B8F5: push ecx
  loc_00E3B8F6: push 00000002h
  loc_00E3B8F8: call [00401048h] ; __vbaFreeObjList
  loc_00E3B8FE: add esp, 00000028h
  loc_00E3B901: lea ecx, var_2C
  loc_00E3B904: call [00401024h] ; __vbaFreeVar
  loc_00E3B90A: sub esp, 00000010h
  loc_00E3B90D: mov ecx, 0000000Bh
  loc_00E3B912: mov edx, esp
  loc_00E3B914: xor eax, eax
  loc_00E3B916: push 00000006h
  loc_00E3B918: push esi
  loc_00E3B919: mov [edx], ecx
  loc_00E3B91B: mov ecx, var_38
  loc_00E3B91E: mov [edx+00000004h], ecx
  loc_00E3B921: mov ecx, [esi]
  loc_00E3B923: mov [edx+00000008h], eax
  loc_00E3B926: mov eax, var_30
  loc_00E3B929: mov [edx+0000000Ch], eax
  loc_00E3B92C: call [ecx+00000340h]
  loc_00E3B932: lea edx, var_18
  loc_00E3B935: push eax
  loc_00E3B936: push edx
  loc_00E3B937: call edi
  loc_00E3B939: push eax
  loc_00E3B93A: call [00401238h] ; __vbaLateIdSt
  loc_00E3B940: lea ecx, var_18
  loc_00E3B943: call ebx
  loc_00E3B945: sub esp, 00000010h
  loc_00E3B948: mov ecx, 0000000Bh
  loc_00E3B94D: mov edx, esp
  loc_00E3B94F: xor eax, eax
  loc_00E3B951: push 8001000Eh
  loc_00E3B956: push esi
  loc_00E3B957: mov [edx], ecx
  loc_00E3B959: mov ecx, var_38
  loc_00E3B95C: mov [edx+00000004h], ecx
  loc_00E3B95F: mov ecx, [esi]
  loc_00E3B961: mov [edx+00000008h], eax
  loc_00E3B964: mov eax, var_30
  loc_00E3B967: mov [edx+0000000Ch], eax
  loc_00E3B96A: call [ecx+00000340h]
  loc_00E3B970: lea edx, var_18
  loc_00E3B973: push eax
  loc_00E3B974: push edx
  loc_00E3B975: call edi
  loc_00E3B977: push eax
  loc_00E3B978: call [00401238h] ; __vbaLateIdSt
  loc_00E3B97E: lea ecx, var_18
  loc_00E3B981: call ebx
  loc_00E3B983: mov eax, [esi]
  loc_00E3B985: push 00000000h
  loc_00E3B987: push FFFFFDDAh
  loc_00E3B98C: push esi
  loc_00E3B98D: call [eax+00000340h]
  loc_00E3B993: lea ecx, var_18
  loc_00E3B996: push eax
  loc_00E3B997: push ecx
  loc_00E3B998: call edi
  loc_00E3B99A: push eax
  loc_00E3B99B: call [00401030h] ; __vbaLateIdCall
  loc_00E3B9A1: add esp, 0000000Ch
  loc_00E3B9A4: lea ecx, var_18
  loc_00E3B9A7: call ebx
  loc_00E3B9A9: or eax, FFFFFFFFh
  loc_00E3B9AC: mov ecx, 0000000Bh
  loc_00E3B9B1: sub esp, 00000010h
  loc_00E3B9B4: mov edx, esp
  loc_00E3B9B6: push 80010007h
  loc_00E3B9BB: push esi
  loc_00E3B9BC: mov [edx], ecx
  loc_00E3B9BE: mov ecx, var_38
  loc_00E3B9C1: mov [edx+00000004h], ecx
  loc_00E3B9C4: mov ecx, [esi]
  loc_00E3B9C6: mov [edx+00000008h], eax
  loc_00E3B9C9: mov eax, var_30
  loc_00E3B9CC: mov [edx+0000000Ch], eax
  loc_00E3B9CF: call [ecx+00000314h]
  loc_00E3B9D5: lea edx, var_18
  loc_00E3B9D8: push eax
  loc_00E3B9D9: push edx
  loc_00E3B9DA: call edi
  loc_00E3B9DC: push eax
  loc_00E3B9DD: call [00401238h] ; __vbaLateIdSt
  loc_00E3B9E3: lea ecx, var_18
  loc_00E3B9E6: call ebx
  loc_00E3B9E8: mov eax, [esi]
  loc_00E3B9EA: push esi
  loc_00E3B9EB: call [eax+0000030Ch]
  loc_00E3B9F1: lea ecx, var_18
  loc_00E3B9F4: push eax
  loc_00E3B9F5: push ecx
  loc_00E3B9F6: call edi
  loc_00E3B9F8: mov edx, [eax]
  loc_00E3B9FA: push eax
  loc_00E3B9FB: mov var_50, eax
  loc_00E3B9FE: call [edx+00000204h]
  loc_00E3BA04: test eax, eax
  loc_00E3BA06: fnclex
  loc_00E3BA08: jge 00E3BA1Fh
  loc_00E3BA0A: mov ecx, var_50
  loc_00E3BA0D: push 00000204h
  loc_00E3BA12: push 006DCB70h
  loc_00E3BA17: push ecx
  loc_00E3BA18: push eax
  loc_00E3BA19: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3BA1F: lea ecx, var_18
  loc_00E3BA22: call ebx
  loc_00E3BA24: mov edx, [esi]
  loc_00E3BA26: push esi
  loc_00E3BA27: call [edx+00000304h]
  loc_00E3BA2D: push eax
  loc_00E3BA2E: lea eax, var_18
  loc_00E3BA31: push eax
  loc_00E3BA32: call edi
  loc_00E3BA34: mov ecx, [eax]
  loc_00E3BA36: push 00000000h
  loc_00E3BA38: push eax
  loc_00E3BA39: mov var_50, eax
  loc_00E3BA3C: call [ecx+0000009Ch]
  loc_00E3BA42: test eax, eax
  loc_00E3BA44: fnclex
  loc_00E3BA46: jge 00E3BA5Dh
  loc_00E3BA48: mov edx, var_50
  loc_00E3BA4B: push 0000009Ch
  loc_00E3BA50: push 006E03D4h
  loc_00E3BA55: push edx
  loc_00E3BA56: push eax
  loc_00E3BA57: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3BA5D: lea ecx, var_18
  loc_00E3BA60: call ebx
  loc_00E3BA62: mov eax, [esi]
  loc_00E3BA64: push esi
  loc_00E3BA65: call [eax+00000310h]
  loc_00E3BA6B: lea ecx, var_18
  loc_00E3BA6E: push eax
  loc_00E3BA6F: push ecx
  loc_00E3BA70: call edi
  loc_00E3BA72: mov edx, [eax]
  loc_00E3BA74: push 00000000h
  loc_00E3BA76: push eax
  loc_00E3BA77: mov var_50, eax
  loc_00E3BA7A: call [edx+0000009Ch]
  loc_00E3BA80: test eax, eax
  loc_00E3BA82: fnclex
  loc_00E3BA84: jge 00E3BA9Bh
  loc_00E3BA86: mov ecx, var_50
  loc_00E3BA89: push 0000009Ch
  loc_00E3BA8E: push 006E03D4h
  loc_00E3BA93: push ecx
  loc_00E3BA94: push eax
  loc_00E3BA95: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3BA9B: lea ecx, var_18
  loc_00E3BA9E: call ebx
  loc_00E3BAA0: mov edx, [esi]
  loc_00E3BAA2: push esi
  loc_00E3BAA3: call [edx+00000300h]
  loc_00E3BAA9: push eax
  loc_00E3BAAA: lea eax, var_18
  loc_00E3BAAD: push eax
  loc_00E3BAAE: call edi
  loc_00E3BAB0: mov esi, eax
  loc_00E3BAB2: push 00000000h
  loc_00E3BAB4: push esi
  loc_00E3BAB5: mov ecx, [esi]
  loc_00E3BAB7: call [ecx+0000009Ch]
  loc_00E3BABD: test eax, eax
  loc_00E3BABF: fnclex
  loc_00E3BAC1: jge 00E3BAD5h
  loc_00E3BAC3: push 0000009Ch
  loc_00E3BAC8: push 006E03D4h
  loc_00E3BACD: push esi
  loc_00E3BACE: push eax
  loc_00E3BACF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3BAD5: lea ecx, var_18
  loc_00E3BAD8: call ebx
  loc_00E3BADA: mov var_4, 00000000h
  loc_00E3BAE1: push 00E3BB06h
  loc_00E3BAE6: jmp 00E3BB05h
  loc_00E3BAE8: lea edx, var_1C
  loc_00E3BAEB: lea eax, var_18
  loc_00E3BAEE: push edx
  loc_00E3BAEF: push eax
  loc_00E3BAF0: push 00000002h
  loc_00E3BAF2: call [00401048h] ; __vbaFreeObjList
  loc_00E3BAF8: add esp, 0000000Ch
  loc_00E3BAFB: lea ecx, var_2C
  loc_00E3BAFE: call [00401024h] ; __vbaFreeVar
  loc_00E3BB04: ret
  loc_00E3BB05: ret
  loc_00E3BB06: mov eax, Me
  loc_00E3BB09: push eax
  loc_00E3BB0A: mov ecx, [eax]
  loc_00E3BB0C: call [ecx+00000008h]
  loc_00E3BB0F: mov eax, var_4
  loc_00E3BB12: mov ecx, var_14
  loc_00E3BB15: pop edi
  loc_00E3BB16: pop esi
  loc_00E3BB17: mov fs:[00000000h], ecx
  loc_00E3BB1E: pop ebx
  loc_00E3BB1F: mov esp, ebp
  loc_00E3BB21: pop ebp
  loc_00E3BB22: retn 0004h
End Sub

Private Sub optjur_Click() 'E3B110
  loc_00E3B110: push ebp
  loc_00E3B111: mov ebp, esp
  loc_00E3B113: sub esp, 0000000Ch
  loc_00E3B116: push 00402836h ; __vbaExceptHandler
  loc_00E3B11B: mov eax, fs:[00000000h]
  loc_00E3B121: push eax
  loc_00E3B122: mov fs:[00000000h], esp
  loc_00E3B129: sub esp, 00000048h
  loc_00E3B12C: push ebx
  loc_00E3B12D: push esi
  loc_00E3B12E: push edi
  loc_00E3B12F: mov var_C, esp
  loc_00E3B132: mov var_8, 00402800h
  loc_00E3B139: mov esi, Me
  loc_00E3B13C: mov eax, esi
  loc_00E3B13E: and eax, 00000001h
  loc_00E3B141: mov var_4, eax
  loc_00E3B144: and esi, FFFFFFFEh
  loc_00E3B147: push esi
  loc_00E3B148: mov Me, esi
  loc_00E3B14B: mov ecx, [esi]
  loc_00E3B14D: call [ecx+00000004h]
  loc_00E3B150: sub esp, 00000010h
  loc_00E3B153: mov ecx, 0000000Bh
  loc_00E3B158: mov edx, esp
  loc_00E3B15A: xor eax, eax
  loc_00E3B15C: mov var_18, eax
  loc_00E3B15F: mov var_1C, eax
  loc_00E3B162: mov [edx], ecx
  loc_00E3B164: mov ecx, var_38
  loc_00E3B167: mov var_2C, eax
  loc_00E3B16A: or eax, FFFFFFFFh
  loc_00E3B16D: mov [edx+00000004h], ecx
  loc_00E3B170: mov ecx, [esi]
  loc_00E3B172: push 80010007h
  loc_00E3B177: push esi
  loc_00E3B178: mov [edx+00000008h], eax
  loc_00E3B17B: mov eax, var_30
  loc_00E3B17E: mov [edx+0000000Ch], eax
  loc_00E3B181: call [ecx+00000314h]
  loc_00E3B187: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3B18D: lea edx, var_18
  loc_00E3B190: push eax
  loc_00E3B191: push edx
  loc_00E3B192: call edi
  loc_00E3B194: push eax
  loc_00E3B195: call [00401238h] ; __vbaLateIdSt
  loc_00E3B19B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E3B1A1: lea ecx, var_18
  loc_00E3B1A4: call ebx
  loc_00E3B1A6: mov eax, [esi]
  loc_00E3B1A8: push esi
  loc_00E3B1A9: call [eax+0000030Ch]
  loc_00E3B1AF: lea ecx, var_18
  loc_00E3B1B2: push eax
  loc_00E3B1B3: push ecx
  loc_00E3B1B4: call edi
  loc_00E3B1B6: mov edx, [eax]
  loc_00E3B1B8: push eax
  loc_00E3B1B9: mov var_50, eax
  loc_00E3B1BC: call [edx+00000204h]
  loc_00E3B1C2: test eax, eax
  loc_00E3B1C4: fnclex
  loc_00E3B1C6: jge 00E3B1DDh
  loc_00E3B1C8: mov ecx, var_50
  loc_00E3B1CB: push 00000204h
  loc_00E3B1D0: push 006DCB70h
  loc_00E3B1D5: push ecx
  loc_00E3B1D6: push eax
  loc_00E3B1D7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3B1DD: lea ecx, var_18
  loc_00E3B1E0: call ebx
  loc_00E3B1E2: mov edx, [esi]
  loc_00E3B1E4: push esi
  loc_00E3B1E5: call [edx+00000308h]
  loc_00E3B1EB: push eax
  loc_00E3B1EC: lea eax, var_18
  loc_00E3B1EF: push eax
  loc_00E3B1F0: call edi
  loc_00E3B1F2: mov ecx, [eax]
  loc_00E3B1F4: push 00000000h
  loc_00E3B1F6: push eax
  loc_00E3B1F7: mov var_50, eax
  loc_00E3B1FA: call [ecx+0000009Ch]
  loc_00E3B200: test eax, eax
  loc_00E3B202: fnclex
  loc_00E3B204: jge 00E3B21Bh
  loc_00E3B206: mov edx, var_50
  loc_00E3B209: push 0000009Ch
  loc_00E3B20E: push 006E03D4h
  loc_00E3B213: push edx
  loc_00E3B214: push eax
  loc_00E3B215: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3B21B: lea ecx, var_18
  loc_00E3B21E: call ebx
  loc_00E3B220: mov eax, [esi]
  loc_00E3B222: push esi
  loc_00E3B223: call [eax+00000310h]
  loc_00E3B229: lea ecx, var_18
  loc_00E3B22C: push eax
  loc_00E3B22D: push ecx
  loc_00E3B22E: call edi
  loc_00E3B230: mov edx, [eax]
  loc_00E3B232: push 00000000h
  loc_00E3B234: push eax
  loc_00E3B235: mov var_50, eax
  loc_00E3B238: call [edx+0000009Ch]
  loc_00E3B23E: test eax, eax
  loc_00E3B240: fnclex
  loc_00E3B242: jge 00E3B259h
  loc_00E3B244: mov ecx, var_50
  loc_00E3B247: push 0000009Ch
  loc_00E3B24C: push 006E03D4h
  loc_00E3B251: push ecx
  loc_00E3B252: push eax
  loc_00E3B253: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3B259: lea ecx, var_18
  loc_00E3B25C: call ebx
  loc_00E3B25E: mov edx, [esi]
  loc_00E3B260: push esi
  loc_00E3B261: call [edx+00000304h]
  loc_00E3B267: push eax
  loc_00E3B268: lea eax, var_18
  loc_00E3B26B: push eax
  loc_00E3B26C: call edi
  loc_00E3B26E: mov ecx, [eax]
  loc_00E3B270: push 00000000h
  loc_00E3B272: push eax
  loc_00E3B273: mov var_50, eax
  loc_00E3B276: call [ecx+0000009Ch]
  loc_00E3B27C: test eax, eax
  loc_00E3B27E: fnclex
  loc_00E3B280: jge 00E3B297h
  loc_00E3B282: mov edx, var_50
  loc_00E3B285: push 0000009Ch
  loc_00E3B28A: push 006E03D4h
  loc_00E3B28F: push edx
  loc_00E3B290: push eax
  loc_00E3B291: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3B297: lea ecx, var_18
  loc_00E3B29A: call ebx
  loc_00E3B29C: sub esp, 00000010h
  loc_00E3B29F: mov ecx, 00000008h
  loc_00E3B2A4: mov edx, esp
  loc_00E3B2A6: mov eax, 006E1738h ; "lc"
  loc_00E3B2AB: push 0000000Eh
  loc_00E3B2AD: push esi
  loc_00E3B2AE: mov [edx], ecx
  loc_00E3B2B0: mov ecx, var_38
  loc_00E3B2B3: mov [edx+00000004h], ecx
  loc_00E3B2B6: mov ecx, [esi]
  loc_00E3B2B8: mov [edx+00000008h], eax
  loc_00E3B2BB: mov eax, var_30
  loc_00E3B2BE: mov [edx+0000000Ch], eax
  loc_00E3B2C1: call [ecx+0000033Ch]
  loc_00E3B2C7: lea edx, var_18
  loc_00E3B2CA: push eax
  loc_00E3B2CB: push edx
  loc_00E3B2CC: call edi
  loc_00E3B2CE: push eax
  loc_00E3B2CF: call [00401238h] ; __vbaLateIdSt
  loc_00E3B2D5: lea ecx, var_18
  loc_00E3B2D8: call ebx
  loc_00E3B2DA: mov eax, [esi]
  loc_00E3B2DC: push 00000000h
  loc_00E3B2DE: push 00000019h
  loc_00E3B2E0: push esi
  loc_00E3B2E1: call [eax+0000033Ch]
  loc_00E3B2E7: lea ecx, var_18
  loc_00E3B2EA: push eax
  loc_00E3B2EB: push ecx
  loc_00E3B2EC: call edi
  loc_00E3B2EE: push eax
  loc_00E3B2EF: call [00401030h] ; __vbaLateIdCall
  loc_00E3B2F5: add esp, 0000000Ch
  loc_00E3B2F8: lea ecx, var_18
  loc_00E3B2FB: call ebx
  loc_00E3B2FD: mov edx, [esi]
  loc_00E3B2FF: push 006E05D8h
  loc_00E3B304: push esi
  loc_00E3B305: call [edx+0000033Ch]
  loc_00E3B30B: push eax
  loc_00E3B30C: lea eax, var_18
  loc_00E3B30F: push eax
  loc_00E3B310: call edi
  loc_00E3B312: push eax
  loc_00E3B313: call [00401224h] ; __vbaCastObj
  loc_00E3B319: lea ecx, var_2C
  loc_00E3B31C: mov var_24, eax
  loc_00E3B31F: push ecx
  loc_00E3B320: mov var_2C, 0000000Dh
  loc_00E3B327: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3B32D: mov ecx, [eax]
  loc_00E3B32F: sub esp, 00000010h
  loc_00E3B332: mov edx, esp
  loc_00E3B334: push 00000000h
  loc_00E3B336: push 0000002Ah
  loc_00E3B338: mov [edx], ecx
  loc_00E3B33A: mov ecx, [eax+00000004h]
  loc_00E3B33D: push esi
  loc_00E3B33E: mov [edx+00000004h], ecx
  loc_00E3B341: mov ecx, [eax+00000008h]
  loc_00E3B344: mov eax, [eax+0000000Ch]
  loc_00E3B347: mov [edx+00000008h], ecx
  loc_00E3B34A: mov ecx, [esi]
  loc_00E3B34C: mov [edx+0000000Ch], eax
  loc_00E3B34F: call [ecx+00000340h]
  loc_00E3B355: lea edx, var_1C
  loc_00E3B358: push eax
  loc_00E3B359: push edx
  loc_00E3B35A: call edi
  loc_00E3B35C: push eax
  loc_00E3B35D: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3B363: lea eax, var_1C
  loc_00E3B366: lea ecx, var_18
  loc_00E3B369: push eax
  loc_00E3B36A: push ecx
  loc_00E3B36B: push 00000002h
  loc_00E3B36D: call [00401048h] ; __vbaFreeObjList
  loc_00E3B373: add esp, 00000028h
  loc_00E3B376: lea ecx, var_2C
  loc_00E3B379: call [00401024h] ; __vbaFreeVar
  loc_00E3B37F: sub esp, 00000010h
  loc_00E3B382: mov ecx, 0000000Bh
  loc_00E3B387: mov edx, esp
  loc_00E3B389: xor eax, eax
  loc_00E3B38B: push 00000006h
  loc_00E3B38D: push esi
  loc_00E3B38E: mov [edx], ecx
  loc_00E3B390: mov ecx, var_38
  loc_00E3B393: mov [edx+00000004h], ecx
  loc_00E3B396: mov ecx, [esi]
  loc_00E3B398: mov [edx+00000008h], eax
  loc_00E3B39B: mov eax, var_30
  loc_00E3B39E: mov [edx+0000000Ch], eax
  loc_00E3B3A1: call [ecx+00000340h]
  loc_00E3B3A7: lea edx, var_18
  loc_00E3B3AA: push eax
  loc_00E3B3AB: push edx
  loc_00E3B3AC: call edi
  loc_00E3B3AE: push eax
  loc_00E3B3AF: call [00401238h] ; __vbaLateIdSt
  loc_00E3B3B5: lea ecx, var_18
  loc_00E3B3B8: call ebx
  loc_00E3B3BA: sub esp, 00000010h
  loc_00E3B3BD: mov ecx, 0000000Bh
  loc_00E3B3C2: mov edx, esp
  loc_00E3B3C4: xor eax, eax
  loc_00E3B3C6: push 8001000Eh
  loc_00E3B3CB: push esi
  loc_00E3B3CC: mov [edx], ecx
  loc_00E3B3CE: mov ecx, var_38
  loc_00E3B3D1: mov [edx+00000004h], ecx
  loc_00E3B3D4: mov ecx, [esi]
  loc_00E3B3D6: mov [edx+00000008h], eax
  loc_00E3B3D9: mov eax, var_30
  loc_00E3B3DC: mov [edx+0000000Ch], eax
  loc_00E3B3DF: call [ecx+00000340h]
  loc_00E3B3E5: lea edx, var_18
  loc_00E3B3E8: push eax
  loc_00E3B3E9: push edx
  loc_00E3B3EA: call edi
  loc_00E3B3EC: push eax
  loc_00E3B3ED: call [00401238h] ; __vbaLateIdSt
  loc_00E3B3F3: lea ecx, var_18
  loc_00E3B3F6: call ebx
  loc_00E3B3F8: mov eax, [esi]
  loc_00E3B3FA: push 00000000h
  loc_00E3B3FC: push FFFFFDDAh
  loc_00E3B401: push esi
  loc_00E3B402: call [eax+00000340h]
  loc_00E3B408: lea ecx, var_18
  loc_00E3B40B: push eax
  loc_00E3B40C: push ecx
  loc_00E3B40D: call edi
  loc_00E3B40F: push eax
  loc_00E3B410: call [00401030h] ; __vbaLateIdCall
  loc_00E3B416: add esp, 0000000Ch
  loc_00E3B419: lea ecx, var_18
  loc_00E3B41C: call ebx
  loc_00E3B41E: mov var_4, 00000000h
  loc_00E3B425: push 00E3B44Ah
  loc_00E3B42A: jmp 00E3B449h
  loc_00E3B42C: lea edx, var_1C
  loc_00E3B42F: lea eax, var_18
  loc_00E3B432: push edx
  loc_00E3B433: push eax
  loc_00E3B434: push 00000002h
  loc_00E3B436: call [00401048h] ; __vbaFreeObjList
  loc_00E3B43C: add esp, 0000000Ch
  loc_00E3B43F: lea ecx, var_2C
  loc_00E3B442: call [00401024h] ; __vbaFreeVar
  loc_00E3B448: ret
  loc_00E3B449: ret
  loc_00E3B44A: mov eax, Me
  loc_00E3B44D: push eax
  loc_00E3B44E: mov ecx, [eax]
  loc_00E3B450: call [ecx+00000008h]
  loc_00E3B453: mov eax, var_4
  loc_00E3B456: mov ecx, var_14
  loc_00E3B459: pop edi
  loc_00E3B45A: pop esi
  loc_00E3B45B: mov fs:[00000000h], ecx
  loc_00E3B462: pop ebx
  loc_00E3B463: mov esp, ebp
  loc_00E3B465: pop ebp
  loc_00E3B466: retn 0004h
End Sub

Private Sub optagama_Click() 'E3AD70
  loc_00E3AD70: push ebp
  loc_00E3AD71: mov ebp, esp
  loc_00E3AD73: sub esp, 0000000Ch
  loc_00E3AD76: push 00402836h ; __vbaExceptHandler
  loc_00E3AD7B: mov eax, fs:[00000000h]
  loc_00E3AD81: push eax
  loc_00E3AD82: mov fs:[00000000h], esp
  loc_00E3AD89: sub esp, 00000048h
  loc_00E3AD8C: push ebx
  loc_00E3AD8D: push esi
  loc_00E3AD8E: push edi
  loc_00E3AD8F: mov var_C, esp
  loc_00E3AD92: mov var_8, 004027F0h
  loc_00E3AD99: mov esi, Me
  loc_00E3AD9C: mov eax, esi
  loc_00E3AD9E: and eax, 00000001h
  loc_00E3ADA1: mov var_4, eax
  loc_00E3ADA4: and esi, FFFFFFFEh
  loc_00E3ADA7: push esi
  loc_00E3ADA8: mov Me, esi
  loc_00E3ADAB: mov ecx, [esi]
  loc_00E3ADAD: call [ecx+00000004h]
  loc_00E3ADB0: sub esp, 00000010h
  loc_00E3ADB3: mov ecx, 0000000Bh
  loc_00E3ADB8: mov edx, esp
  loc_00E3ADBA: xor eax, eax
  loc_00E3ADBC: mov var_18, eax
  loc_00E3ADBF: mov var_1C, eax
  loc_00E3ADC2: mov [edx], ecx
  loc_00E3ADC4: mov ecx, var_38
  loc_00E3ADC7: mov var_2C, eax
  loc_00E3ADCA: or eax, FFFFFFFFh
  loc_00E3ADCD: mov [edx+00000004h], ecx
  loc_00E3ADD0: mov ecx, [esi]
  loc_00E3ADD2: push 80010007h
  loc_00E3ADD7: push esi
  loc_00E3ADD8: mov [edx+00000008h], eax
  loc_00E3ADDB: mov eax, var_30
  loc_00E3ADDE: mov [edx+0000000Ch], eax
  loc_00E3ADE1: call [ecx+00000314h]
  loc_00E3ADE7: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3ADED: lea edx, var_18
  loc_00E3ADF0: push eax
  loc_00E3ADF1: push edx
  loc_00E3ADF2: call edi
  loc_00E3ADF4: push eax
  loc_00E3ADF5: call [00401238h] ; __vbaLateIdSt
  loc_00E3ADFB: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E3AE01: lea ecx, var_18
  loc_00E3AE04: call ebx
  loc_00E3AE06: mov eax, [esi]
  loc_00E3AE08: push esi
  loc_00E3AE09: call [eax+0000030Ch]
  loc_00E3AE0F: lea ecx, var_18
  loc_00E3AE12: push eax
  loc_00E3AE13: push ecx
  loc_00E3AE14: call edi
  loc_00E3AE16: mov edx, [eax]
  loc_00E3AE18: push eax
  loc_00E3AE19: mov var_50, eax
  loc_00E3AE1C: call [edx+00000204h]
  loc_00E3AE22: test eax, eax
  loc_00E3AE24: fnclex
  loc_00E3AE26: jge 00E3AE3Dh
  loc_00E3AE28: mov ecx, var_50
  loc_00E3AE2B: push 00000204h
  loc_00E3AE30: push 006DCB70h
  loc_00E3AE35: push ecx
  loc_00E3AE36: push eax
  loc_00E3AE37: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AE3D: lea ecx, var_18
  loc_00E3AE40: call ebx
  loc_00E3AE42: mov edx, [esi]
  loc_00E3AE44: push esi
  loc_00E3AE45: call [edx+00000308h]
  loc_00E3AE4B: push eax
  loc_00E3AE4C: lea eax, var_18
  loc_00E3AE4F: push eax
  loc_00E3AE50: call edi
  loc_00E3AE52: mov ecx, [eax]
  loc_00E3AE54: push 00000000h
  loc_00E3AE56: push eax
  loc_00E3AE57: mov var_50, eax
  loc_00E3AE5A: call [ecx+0000009Ch]
  loc_00E3AE60: test eax, eax
  loc_00E3AE62: fnclex
  loc_00E3AE64: jge 00E3AE7Bh
  loc_00E3AE66: mov edx, var_50
  loc_00E3AE69: push 0000009Ch
  loc_00E3AE6E: push 006E03D4h
  loc_00E3AE73: push edx
  loc_00E3AE74: push eax
  loc_00E3AE75: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AE7B: lea ecx, var_18
  loc_00E3AE7E: call ebx
  loc_00E3AE80: mov eax, [esi]
  loc_00E3AE82: push esi
  loc_00E3AE83: call [eax+00000304h]
  loc_00E3AE89: lea ecx, var_18
  loc_00E3AE8C: push eax
  loc_00E3AE8D: push ecx
  loc_00E3AE8E: call edi
  loc_00E3AE90: mov edx, [eax]
  loc_00E3AE92: push 00000000h
  loc_00E3AE94: push eax
  loc_00E3AE95: mov var_50, eax
  loc_00E3AE98: call [edx+0000009Ch]
  loc_00E3AE9E: test eax, eax
  loc_00E3AEA0: fnclex
  loc_00E3AEA2: jge 00E3AEB9h
  loc_00E3AEA4: mov ecx, var_50
  loc_00E3AEA7: push 0000009Ch
  loc_00E3AEAC: push 006E03D4h
  loc_00E3AEB1: push ecx
  loc_00E3AEB2: push eax
  loc_00E3AEB3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AEB9: lea ecx, var_18
  loc_00E3AEBC: call ebx
  loc_00E3AEBE: mov edx, [esi]
  loc_00E3AEC0: push esi
  loc_00E3AEC1: call [edx+00000300h]
  loc_00E3AEC7: push eax
  loc_00E3AEC8: lea eax, var_18
  loc_00E3AECB: push eax
  loc_00E3AECC: call edi
  loc_00E3AECE: mov ecx, [eax]
  loc_00E3AED0: push 00000000h
  loc_00E3AED2: push eax
  loc_00E3AED3: mov var_50, eax
  loc_00E3AED6: call [ecx+0000009Ch]
  loc_00E3AEDC: test eax, eax
  loc_00E3AEDE: fnclex
  loc_00E3AEE0: jge 00E3AEF7h
  loc_00E3AEE2: mov edx, var_50
  loc_00E3AEE5: push 0000009Ch
  loc_00E3AEEA: push 006E03D4h
  loc_00E3AEEF: push edx
  loc_00E3AEF0: push eax
  loc_00E3AEF1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AEF7: lea ecx, var_18
  loc_00E3AEFA: call ebx
  loc_00E3AEFC: sub esp, 00000010h
  loc_00E3AEFF: mov ecx, 00000008h
  loc_00E3AF04: mov edx, esp
  loc_00E3AF06: mov eax, 006E1738h ; "lc"
  loc_00E3AF0B: push 0000000Eh
  loc_00E3AF0D: push esi
  loc_00E3AF0E: mov [edx], ecx
  loc_00E3AF10: mov ecx, var_38
  loc_00E3AF13: mov [edx+00000004h], ecx
  loc_00E3AF16: mov ecx, [esi]
  loc_00E3AF18: mov [edx+00000008h], eax
  loc_00E3AF1B: mov eax, var_30
  loc_00E3AF1E: mov [edx+0000000Ch], eax
  loc_00E3AF21: call [ecx+0000033Ch]
  loc_00E3AF27: lea edx, var_18
  loc_00E3AF2A: push eax
  loc_00E3AF2B: push edx
  loc_00E3AF2C: call edi
  loc_00E3AF2E: push eax
  loc_00E3AF2F: call [00401238h] ; __vbaLateIdSt
  loc_00E3AF35: lea ecx, var_18
  loc_00E3AF38: call ebx
  loc_00E3AF3A: mov eax, [esi]
  loc_00E3AF3C: push 00000000h
  loc_00E3AF3E: push 00000019h
  loc_00E3AF40: push esi
  loc_00E3AF41: call [eax+0000033Ch]
  loc_00E3AF47: lea ecx, var_18
  loc_00E3AF4A: push eax
  loc_00E3AF4B: push ecx
  loc_00E3AF4C: call edi
  loc_00E3AF4E: push eax
  loc_00E3AF4F: call [00401030h] ; __vbaLateIdCall
  loc_00E3AF55: add esp, 0000000Ch
  loc_00E3AF58: lea ecx, var_18
  loc_00E3AF5B: call ebx
  loc_00E3AF5D: mov edx, [esi]
  loc_00E3AF5F: push esi
  loc_00E3AF60: call [edx+0000030Ch]
  loc_00E3AF66: push eax
  loc_00E3AF67: lea eax, var_18
  loc_00E3AF6A: push eax
  loc_00E3AF6B: call edi
  loc_00E3AF6D: mov ecx, [eax]
  loc_00E3AF6F: push 006DCC80h
  loc_00E3AF74: push eax
  loc_00E3AF75: mov var_50, eax
  loc_00E3AF78: call [ecx+000000A4h]
  loc_00E3AF7E: test eax, eax
  loc_00E3AF80: fnclex
  loc_00E3AF82: jge 00E3AF99h
  loc_00E3AF84: mov edx, var_50
  loc_00E3AF87: push 000000A4h
  loc_00E3AF8C: push 006DCB70h
  loc_00E3AF91: push edx
  loc_00E3AF92: push eax
  loc_00E3AF93: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AF99: lea ecx, var_18
  loc_00E3AF9C: call ebx
  loc_00E3AF9E: mov eax, [esi]
  loc_00E3AFA0: push 006E05D8h
  loc_00E3AFA5: push esi
  loc_00E3AFA6: call [eax+0000033Ch]
  loc_00E3AFAC: lea ecx, var_18
  loc_00E3AFAF: push eax
  loc_00E3AFB0: push ecx
  loc_00E3AFB1: call edi
  loc_00E3AFB3: push eax
  loc_00E3AFB4: call [00401224h] ; __vbaCastObj
  loc_00E3AFBA: lea edx, var_2C
  loc_00E3AFBD: mov var_24, eax
  loc_00E3AFC0: push edx
  loc_00E3AFC1: mov var_2C, 0000000Dh
  loc_00E3AFC8: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3AFCE: mov edx, [eax]
  loc_00E3AFD0: sub esp, 00000010h
  loc_00E3AFD3: mov ecx, esp
  loc_00E3AFD5: push 00000000h
  loc_00E3AFD7: push 0000002Ah
  loc_00E3AFD9: mov [ecx], edx
  loc_00E3AFDB: mov edx, [eax+00000004h]
  loc_00E3AFDE: push esi
  loc_00E3AFDF: mov [ecx+00000004h], edx
  loc_00E3AFE2: mov edx, [eax+00000008h]
  loc_00E3AFE5: mov eax, [eax+0000000Ch]
  loc_00E3AFE8: mov [ecx+00000008h], edx
  loc_00E3AFEB: mov [ecx+0000000Ch], eax
  loc_00E3AFEE: mov ecx, [esi]
  loc_00E3AFF0: call [ecx+00000340h]
  loc_00E3AFF6: lea edx, var_1C
  loc_00E3AFF9: push eax
  loc_00E3AFFA: push edx
  loc_00E3AFFB: call edi
  loc_00E3AFFD: push eax
  loc_00E3AFFE: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3B004: lea eax, var_1C
  loc_00E3B007: lea ecx, var_18
  loc_00E3B00A: push eax
  loc_00E3B00B: push ecx
  loc_00E3B00C: push 00000002h
  loc_00E3B00E: call [00401048h] ; __vbaFreeObjList
  loc_00E3B014: add esp, 00000028h
  loc_00E3B017: lea ecx, var_2C
  loc_00E3B01A: call [00401024h] ; __vbaFreeVar
  loc_00E3B020: sub esp, 00000010h
  loc_00E3B023: mov ecx, 0000000Bh
  loc_00E3B028: mov edx, esp
  loc_00E3B02A: xor eax, eax
  loc_00E3B02C: push 00000006h
  loc_00E3B02E: push esi
  loc_00E3B02F: mov [edx], ecx
  loc_00E3B031: mov ecx, var_38
  loc_00E3B034: mov [edx+00000004h], ecx
  loc_00E3B037: mov ecx, [esi]
  loc_00E3B039: mov [edx+00000008h], eax
  loc_00E3B03C: mov eax, var_30
  loc_00E3B03F: mov [edx+0000000Ch], eax
  loc_00E3B042: call [ecx+00000340h]
  loc_00E3B048: lea edx, var_18
  loc_00E3B04B: push eax
  loc_00E3B04C: push edx
  loc_00E3B04D: call edi
  loc_00E3B04F: push eax
  loc_00E3B050: call [00401238h] ; __vbaLateIdSt
  loc_00E3B056: lea ecx, var_18
  loc_00E3B059: call ebx
  loc_00E3B05B: xor eax, eax
  loc_00E3B05D: sub esp, 00000010h
  loc_00E3B060: mov edx, esp
  loc_00E3B062: mov ecx, 0000000Bh
  loc_00E3B067: push 8001000Eh
  loc_00E3B06C: mov [edx], ecx
  loc_00E3B06E: mov ecx, var_38
  loc_00E3B071: mov [edx+00000004h], ecx
  loc_00E3B074: mov ecx, [esi]
  loc_00E3B076: mov [edx+00000008h], eax
  loc_00E3B079: mov eax, var_30
  loc_00E3B07C: mov [edx+0000000Ch], eax
  loc_00E3B07F: push esi
  loc_00E3B080: call [ecx+00000340h]
  loc_00E3B086: lea edx, var_18
  loc_00E3B089: push eax
  loc_00E3B08A: push edx
  loc_00E3B08B: call edi
  loc_00E3B08D: push eax
  loc_00E3B08E: call [00401238h] ; __vbaLateIdSt
  loc_00E3B094: lea ecx, var_18
  loc_00E3B097: call ebx
  loc_00E3B099: mov eax, [esi]
  loc_00E3B09B: push 00000000h
  loc_00E3B09D: push FFFFFDDAh
  loc_00E3B0A2: push esi
  loc_00E3B0A3: call [eax+00000340h]
  loc_00E3B0A9: lea ecx, var_18
  loc_00E3B0AC: push eax
  loc_00E3B0AD: push ecx
  loc_00E3B0AE: call edi
  loc_00E3B0B0: push eax
  loc_00E3B0B1: call [00401030h] ; __vbaLateIdCall
  loc_00E3B0B7: add esp, 0000000Ch
  loc_00E3B0BA: lea ecx, var_18
  loc_00E3B0BD: call ebx
  loc_00E3B0BF: mov var_4, 00000000h
  loc_00E3B0C6: push 00E3B0EBh
  loc_00E3B0CB: jmp 00E3B0EAh
  loc_00E3B0CD: lea edx, var_1C
  loc_00E3B0D0: lea eax, var_18
  loc_00E3B0D3: push edx
  loc_00E3B0D4: push eax
  loc_00E3B0D5: push 00000002h
  loc_00E3B0D7: call [00401048h] ; __vbaFreeObjList
  loc_00E3B0DD: add esp, 0000000Ch
  loc_00E3B0E0: lea ecx, var_2C
  loc_00E3B0E3: call [00401024h] ; __vbaFreeVar
  loc_00E3B0E9: ret
  loc_00E3B0EA: ret
  loc_00E3B0EB: mov eax, Me
  loc_00E3B0EE: push eax
  loc_00E3B0EF: mov ecx, [eax]
  loc_00E3B0F1: call [ecx+00000008h]
  loc_00E3B0F4: mov eax, var_4
  loc_00E3B0F7: mov ecx, var_14
  loc_00E3B0FA: pop edi
  loc_00E3B0FB: pop esi
  loc_00E3B0FC: mov fs:[00000000h], ecx
  loc_00E3B103: pop ebx
  loc_00E3B104: mov esp, ebp
  loc_00E3B106: pop ebp
  loc_00E3B107: retn 0004h
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer) 'E3A340
  loc_00E3A340: push ebp
  loc_00E3A341: mov ebp, esp
  loc_00E3A343: sub esp, 0000000Ch
  loc_00E3A346: push 00402836h ; __vbaExceptHandler
  loc_00E3A34B: mov eax, fs:[00000000h]
  loc_00E3A351: push eax
  loc_00E3A352: mov fs:[00000000h], esp
  loc_00E3A359: sub esp, 000000B4h
  loc_00E3A35F: push ebx
  loc_00E3A360: push esi
  loc_00E3A361: push edi
  loc_00E3A362: mov var_C, esp
  loc_00E3A365: mov var_8, 004027D0h
  loc_00E3A36C: mov esi, Me
  loc_00E3A36F: mov eax, esi
  loc_00E3A371: and eax, 00000001h
  loc_00E3A374: mov var_4, eax
  loc_00E3A377: and esi, FFFFFFFEh
  loc_00E3A37A: push esi
  loc_00E3A37B: mov Me, esi
  loc_00E3A37E: mov ecx, [esi]
  loc_00E3A380: call [ecx+00000004h]
  loc_00E3A383: mov edx, [esi]
  loc_00E3A385: xor ebx, ebx
  loc_00E3A387: push esi
  loc_00E3A388: mov var_24, ebx
  loc_00E3A38B: mov var_28, ebx
  loc_00E3A38E: mov var_2C, ebx
  loc_00E3A391: mov var_30, ebx
  loc_00E3A394: mov var_40, ebx
  loc_00E3A397: mov var_50, ebx
  loc_00E3A39A: mov var_60, ebx
  loc_00E3A39D: mov var_70, ebx
  loc_00E3A3A0: mov var_80, ebx
  loc_00E3A3A3: mov var_90, ebx
  loc_00E3A3A9: mov var_B4, ebx
  loc_00E3A3AF: call [edx+00000308h]
  loc_00E3A3B5: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3A3BB: push eax
  loc_00E3A3BC: lea eax, var_30
  loc_00E3A3BF: push eax
  loc_00E3A3C0: call edi
  loc_00E3A3C2: mov ecx, [eax]
  loc_00E3A3C4: lea edx, var_B4
  loc_00E3A3CA: push edx
  loc_00E3A3CB: push eax
  loc_00E3A3CC: mov var_B8, eax
  loc_00E3A3D2: call [ecx+000000E0h]
  loc_00E3A3D8: cmp eax, ebx
  loc_00E3A3DA: fnclex
  loc_00E3A3DC: jge 00E3A3F6h
  loc_00E3A3DE: mov ecx, var_B8
  loc_00E3A3E4: push 000000E0h
  loc_00E3A3E9: push 006E03D4h
  loc_00E3A3EE: push ecx
  loc_00E3A3EF: push eax
  loc_00E3A3F0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A3F6: mov edx, var_B4
  loc_00E3A3FC: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E3A402: lea ecx, var_30
  loc_00E3A405: mov var_C0, edx
  loc_00E3A40B: call ebx
  loc_00E3A40D: cmp var_C0, 0000h
  loc_00E3A415: jz 00E3A47Dh
  loc_00E3A417: mov eax, [esi]
  loc_00E3A419: push esi
  loc_00E3A41A: call [eax+0000030Ch]
  loc_00E3A420: lea ecx, var_30
  loc_00E3A423: push eax
  loc_00E3A424: push ecx
  loc_00E3A425: call edi
  loc_00E3A427: mov edx, [eax]
  loc_00E3A429: lea ecx, var_28
  loc_00E3A42C: push ecx
  loc_00E3A42D: push eax
  loc_00E3A42E: mov var_B8, eax
  loc_00E3A434: call [edx+000000A0h]
  loc_00E3A43A: test eax, eax
  loc_00E3A43C: fnclex
  loc_00E3A43E: jge 00E3A458h
  loc_00E3A440: mov edx, var_B8
  loc_00E3A446: push 000000A0h
  loc_00E3A44B: push 006DCB70h
  loc_00E3A450: push edx
  loc_00E3A451: push eax
  loc_00E3A452: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A458: mov eax, var_28
  loc_00E3A45B: push 006E1958h ; "select * from lc where no_daftar like '" & Chr(37)
  loc_00E3A460: push eax
  loc_00E3A461: call [0040106Ch] ; __vbaStrCat
  loc_00E3A467: mov edx, eax
  loc_00E3A469: lea ecx, var_2C
  loc_00E3A46C: call [00401228h] ; __vbaStrMove
  loc_00E3A472: push eax
  loc_00E3A473: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E3A478: jmp 00E3A6CBh
  loc_00E3A47D: mov edx, [esi]
  loc_00E3A47F: push esi
  loc_00E3A480: call [edx+00000304h]
  loc_00E3A486: push eax
  loc_00E3A487: lea eax, var_30
  loc_00E3A48A: push eax
  loc_00E3A48B: call edi
  loc_00E3A48D: mov ecx, [eax]
  loc_00E3A48F: lea edx, var_B4
  loc_00E3A495: push edx
  loc_00E3A496: push eax
  loc_00E3A497: mov var_B8, eax
  loc_00E3A49D: call [ecx+000000E0h]
  loc_00E3A4A3: test eax, eax
  loc_00E3A4A5: fnclex
  loc_00E3A4A7: jge 00E3A4C1h
  loc_00E3A4A9: mov ecx, var_B8
  loc_00E3A4AF: push 000000E0h
  loc_00E3A4B4: push 006E03D4h
  loc_00E3A4B9: push ecx
  loc_00E3A4BA: push eax
  loc_00E3A4BB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A4C1: mov edx, var_B4
  loc_00E3A4C7: lea ecx, var_30
  loc_00E3A4CA: mov var_C0, edx
  loc_00E3A4D0: call ebx
  loc_00E3A4D2: cmp var_C0, 0000h
  loc_00E3A4DA: jz 00E3A542h
  loc_00E3A4DC: mov eax, [esi]
  loc_00E3A4DE: push esi
  loc_00E3A4DF: call [eax+0000030Ch]
  loc_00E3A4E5: lea ecx, var_30
  loc_00E3A4E8: push eax
  loc_00E3A4E9: push ecx
  loc_00E3A4EA: call edi
  loc_00E3A4EC: mov edx, [eax]
  loc_00E3A4EE: lea ecx, var_28
  loc_00E3A4F1: push ecx
  loc_00E3A4F2: push eax
  loc_00E3A4F3: mov var_B8, eax
  loc_00E3A4F9: call [edx+000000A0h]
  loc_00E3A4FF: test eax, eax
  loc_00E3A501: fnclex
  loc_00E3A503: jge 00E3A51Dh
  loc_00E3A505: mov edx, var_B8
  loc_00E3A50B: push 000000A0h
  loc_00E3A510: push 006DCB70h
  loc_00E3A515: push edx
  loc_00E3A516: push eax
  loc_00E3A517: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A51D: mov eax, var_28
  loc_00E3A520: push 006E19DCh ; "select * from lc where nama_siswa like '" & Chr(37)
  loc_00E3A525: push eax
  loc_00E3A526: call [0040106Ch] ; __vbaStrCat
  loc_00E3A52C: mov edx, eax
  loc_00E3A52E: lea ecx, var_2C
  loc_00E3A531: call [00401228h] ; __vbaStrMove
  loc_00E3A537: push eax
  loc_00E3A538: push 006E1A34h ; Chr(37) & "' order by nama_siswa asc"
  loc_00E3A53D: jmp 00E3A6CBh
  loc_00E3A542: mov edx, [esi]
  loc_00E3A544: push esi
  loc_00E3A545: call [edx+00000300h]
  loc_00E3A54B: push eax
  loc_00E3A54C: lea eax, var_30
  loc_00E3A54F: push eax
  loc_00E3A550: call edi
  loc_00E3A552: mov ecx, [eax]
  loc_00E3A554: lea edx, var_B4
  loc_00E3A55A: push edx
  loc_00E3A55B: push eax
  loc_00E3A55C: mov var_B8, eax
  loc_00E3A562: call [ecx+000000E0h]
  loc_00E3A568: test eax, eax
  loc_00E3A56A: fnclex
  loc_00E3A56C: jge 00E3A586h
  loc_00E3A56E: mov ecx, var_B8
  loc_00E3A574: push 000000E0h
  loc_00E3A579: push 006E03D4h
  loc_00E3A57E: push ecx
  loc_00E3A57F: push eax
  loc_00E3A580: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A586: mov edx, var_B4
  loc_00E3A58C: lea ecx, var_30
  loc_00E3A58F: mov var_C0, edx
  loc_00E3A595: call ebx
  loc_00E3A597: cmp var_C0, 0000h
  loc_00E3A59F: jz 00E3A607h
  loc_00E3A5A1: mov eax, [esi]
  loc_00E3A5A3: push esi
  loc_00E3A5A4: call [eax+0000030Ch]
  loc_00E3A5AA: lea ecx, var_30
  loc_00E3A5AD: push eax
  loc_00E3A5AE: push ecx
  loc_00E3A5AF: call edi
  loc_00E3A5B1: mov edx, [eax]
  loc_00E3A5B3: lea ecx, var_28
  loc_00E3A5B6: push ecx
  loc_00E3A5B7: push eax
  loc_00E3A5B8: mov var_B8, eax
  loc_00E3A5BE: call [edx+000000A0h]
  loc_00E3A5C4: test eax, eax
  loc_00E3A5C6: fnclex
  loc_00E3A5C8: jge 00E3A5E2h
  loc_00E3A5CA: mov edx, var_B8
  loc_00E3A5D0: push 000000A0h
  loc_00E3A5D5: push 006DCB70h
  loc_00E3A5DA: push edx
  loc_00E3A5DB: push eax
  loc_00E3A5DC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A5E2: mov eax, var_28
  loc_00E3A5E5: push 006E1F60h ; "select * from lc where jurusan like '" & Chr(37)
  loc_00E3A5EA: push eax
  loc_00E3A5EB: call [0040106Ch] ; __vbaStrCat
  loc_00E3A5F1: mov edx, eax
  loc_00E3A5F3: lea ecx, var_2C
  loc_00E3A5F6: call [00401228h] ; __vbaStrMove
  loc_00E3A5FC: push eax
  loc_00E3A5FD: push 006E1FB4h ; Chr(37) & "' order by jurusan asc"
  loc_00E3A602: jmp 00E3A6CBh
  loc_00E3A607: mov edx, [esi]
  loc_00E3A609: push esi
  loc_00E3A60A: call [edx+00000310h]
  loc_00E3A610: push eax
  loc_00E3A611: lea eax, var_30
  loc_00E3A614: push eax
  loc_00E3A615: call edi
  loc_00E3A617: mov ecx, [eax]
  loc_00E3A619: lea edx, var_B4
  loc_00E3A61F: push edx
  loc_00E3A620: push eax
  loc_00E3A621: mov var_B8, eax
  loc_00E3A627: call [ecx+000000E0h]
  loc_00E3A62D: test eax, eax
  loc_00E3A62F: fnclex
  loc_00E3A631: jge 00E3A64Bh
  loc_00E3A633: mov ecx, var_B8
  loc_00E3A639: push 000000E0h
  loc_00E3A63E: push 006E03D4h
  loc_00E3A643: push ecx
  loc_00E3A644: push eax
  loc_00E3A645: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A64B: mov edx, var_B4
  loc_00E3A651: lea ecx, var_30
  loc_00E3A654: mov var_C0, edx
  loc_00E3A65A: call ebx
  loc_00E3A65C: cmp var_C0, 0000h
  loc_00E3A664: jz 00E3A774h
  loc_00E3A66A: mov eax, [esi]
  loc_00E3A66C: push esi
  loc_00E3A66D: call [eax+0000030Ch]
  loc_00E3A673: lea ecx, var_30
  loc_00E3A676: push eax
  loc_00E3A677: push ecx
  loc_00E3A678: call edi
  loc_00E3A67A: mov edx, [eax]
  loc_00E3A67C: lea ecx, var_28
  loc_00E3A67F: push ecx
  loc_00E3A680: push eax
  loc_00E3A681: mov var_B8, eax
  loc_00E3A687: call [edx+000000A0h]
  loc_00E3A68D: test eax, eax
  loc_00E3A68F: fnclex
  loc_00E3A691: jge 00E3A6ABh
  loc_00E3A693: mov edx, var_B8
  loc_00E3A699: push 000000A0h
  loc_00E3A69E: push 006DCB70h
  loc_00E3A6A3: push edx
  loc_00E3A6A4: push eax
  loc_00E3A6A5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A6AB: mov eax, var_28
  loc_00E3A6AE: push 006E1FE8h ; "select * from lc where agama like '" & Chr(37)
  loc_00E3A6B3: push eax
  loc_00E3A6B4: call [0040106Ch] ; __vbaStrCat
  loc_00E3A6BA: mov edx, eax
  loc_00E3A6BC: lea ecx, var_2C
  loc_00E3A6BF: call [00401228h] ; __vbaStrMove
  loc_00E3A6C5: push eax
  loc_00E3A6C6: push 006E1F30h ; Chr(37) & "' order by agama asc"
  loc_00E3A6CB: call [0040106Ch] ; __vbaStrCat
  loc_00E3A6D1: lea edx, var_40
  loc_00E3A6D4: lea ecx, var_24
  loc_00E3A6D7: mov var_38, eax
  loc_00E3A6DA: mov var_40, 00000008h
  loc_00E3A6E1: call [0040101Ch] ; __vbaVarMove
  loc_00E3A6E7: lea ecx, var_2C
  loc_00E3A6EA: lea edx, var_28
  loc_00E3A6ED: push ecx
  loc_00E3A6EE: push edx
  loc_00E3A6EF: push 00000002h
  loc_00E3A6F1: call [004011BCh] ; __vbaFreeStrList
  loc_00E3A6F7: add esp, 0000000Ch
  loc_00E3A6FA: lea ecx, var_30
  loc_00E3A6FD: call ebx
  loc_00E3A6FF: lea eax, var_24
  loc_00E3A702: push eax
  loc_00E3A703: call [00401230h] ; __vbaStrVarCopy
  loc_00E3A709: sub esp, 00000010h
  loc_00E3A70C: mov ecx, 00000008h
  loc_00E3A711: mov edx, esp
  loc_00E3A713: mov var_40, ecx
  loc_00E3A716: mov var_38, eax
  loc_00E3A719: push 0000000Eh
  loc_00E3A71B: mov [edx], ecx
  loc_00E3A71D: mov ecx, var_3C
  loc_00E3A720: push esi
  loc_00E3A721: mov [edx+00000004h], ecx
  loc_00E3A724: mov ecx, [esi]
  loc_00E3A726: mov [edx+00000008h], eax
  loc_00E3A729: mov eax, var_34
  loc_00E3A72C: mov [edx+0000000Ch], eax
  loc_00E3A72F: call [ecx+0000033Ch]
  loc_00E3A735: lea edx, var_30
  loc_00E3A738: push eax
  loc_00E3A739: push edx
  loc_00E3A73A: call edi
  loc_00E3A73C: push eax
  loc_00E3A73D: call [00401238h] ; __vbaLateIdSt
  loc_00E3A743: lea ecx, var_30
  loc_00E3A746: call ebx
  loc_00E3A748: lea ecx, var_40
  loc_00E3A74B: call [00401024h] ; __vbaFreeVar
  loc_00E3A751: mov eax, [esi]
  loc_00E3A753: push 00000000h
  loc_00E3A755: push 00000019h
  loc_00E3A757: push esi
  loc_00E3A758: call [eax+0000033Ch]
  loc_00E3A75E: lea ecx, var_30
  loc_00E3A761: push eax
  loc_00E3A762: push ecx
  loc_00E3A763: call edi
  loc_00E3A765: push eax
  loc_00E3A766: call [00401030h] ; __vbaLateIdCall
  loc_00E3A76C: add esp, 0000000Ch
  loc_00E3A76F: jmp 00E3A889h
  loc_00E3A774: mov edx, [esi]
  loc_00E3A776: push 00000000h
  loc_00E3A778: push 80010007h
  loc_00E3A77D: push esi
  loc_00E3A77E: call [edx+00000314h]
  loc_00E3A784: push eax
  loc_00E3A785: lea eax, var_30
  loc_00E3A788: push eax
  loc_00E3A789: call edi
  loc_00E3A78B: lea ecx, var_40
  loc_00E3A78E: push eax
  loc_00E3A78F: push ecx
  loc_00E3A790: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E3A796: add esp, 00000010h
  loc_00E3A799: push eax
  loc_00E3A79A: call [004010CCh] ; __vbaBoolVar
  loc_00E3A7A0: neg ax
  loc_00E3A7A3: sbb eax, eax
  loc_00E3A7A5: lea ecx, var_30
  loc_00E3A7A8: inc eax
  loc_00E3A7A9: neg eax
  loc_00E3A7AB: mov var_B8, ax
  loc_00E3A7B2: call ebx
  loc_00E3A7B4: lea ecx, var_40
  loc_00E3A7B7: call [00401024h] ; __vbaFreeVar
  loc_00E3A7BD: cmp var_B8, 0000h
  loc_00E3A7C5: jz 00E3A88Eh
  loc_00E3A7CB: mov ecx, 80020004h
  loc_00E3A7D0: mov eax, 0000000Ah
  loc_00E3A7D5: mov var_68, ecx
  loc_00E3A7D8: mov var_58, ecx
  loc_00E3A7DB: lea edx, var_90
  loc_00E3A7E1: lea ecx, var_50
  loc_00E3A7E4: mov var_70, eax
  loc_00E3A7E7: mov var_60, eax
  loc_00E3A7EA: mov var_88, 006E16F0h ; "SMK RK BT PS"
  loc_00E3A7F4: mov var_90, 00000008h
  loc_00E3A7FE: call [004011F0h] ; __vbaVarDup
  loc_00E3A804: lea edx, var_80
  loc_00E3A807: lea ecx, var_40
  loc_00E3A80A: mov var_78, 006E1CE8h ; "Silahkan Pilih Item Data yang Ingin dicari !"
  loc_00E3A811: mov var_80, 00000008h
  loc_00E3A818: call [004011F0h] ; __vbaVarDup
  loc_00E3A81E: lea edx, var_70
  loc_00E3A821: lea eax, var_60
  loc_00E3A824: push edx
  loc_00E3A825: lea ecx, var_50
  loc_00E3A828: push eax
  loc_00E3A829: push ecx
  loc_00E3A82A: lea edx, var_40
  loc_00E3A82D: push 00000040h
  loc_00E3A82F: push edx
  loc_00E3A830: call [004010A8h] ; rtcMsgBox
  loc_00E3A836: lea eax, var_70
  loc_00E3A839: lea ecx, var_60
  loc_00E3A83C: push eax
  loc_00E3A83D: lea edx, var_50
  loc_00E3A840: push ecx
  loc_00E3A841: lea eax, var_40
  loc_00E3A844: push edx
  loc_00E3A845: push eax
  loc_00E3A846: push 00000004h
  loc_00E3A848: call [00401038h] ; __vbaFreeVarList
  loc_00E3A84E: mov ecx, [esi]
  loc_00E3A850: add esp, 00000014h
  loc_00E3A853: push esi
  loc_00E3A854: call [ecx+0000030Ch]
  loc_00E3A85A: lea edx, var_30
  loc_00E3A85D: push eax
  loc_00E3A85E: push edx
  loc_00E3A85F: call edi
  loc_00E3A861: mov esi, eax
  loc_00E3A863: push 006DCC80h
  loc_00E3A868: push esi
  loc_00E3A869: mov eax, [esi]
  loc_00E3A86B: call [eax+000000A4h]
  loc_00E3A871: test eax, eax
  loc_00E3A873: fnclex
  loc_00E3A875: jge 00E3A889h
  loc_00E3A877: push 000000A4h
  loc_00E3A87C: push 006DCB70h
  loc_00E3A881: push esi
  loc_00E3A882: push eax
  loc_00E3A883: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A889: lea ecx, var_30
  loc_00E3A88C: call ebx
  loc_00E3A88E: mov var_4, 00000000h
  loc_00E3A895: push 00E3A8DEh
  loc_00E3A89A: jmp 00E3A8D4h
  loc_00E3A89C: lea ecx, var_2C
  loc_00E3A89F: lea edx, var_28
  loc_00E3A8A2: push ecx
  loc_00E3A8A3: push edx
  loc_00E3A8A4: push 00000002h
  loc_00E3A8A6: call [004011BCh] ; __vbaFreeStrList
  loc_00E3A8AC: add esp, 0000000Ch
  loc_00E3A8AF: lea ecx, var_30
  loc_00E3A8B2: call [00401254h] ; __vbaFreeObj
  loc_00E3A8B8: lea eax, var_70
  loc_00E3A8BB: lea ecx, var_60
  loc_00E3A8BE: push eax
  loc_00E3A8BF: lea edx, var_50
  loc_00E3A8C2: push ecx
  loc_00E3A8C3: lea eax, var_40
  loc_00E3A8C6: push edx
  loc_00E3A8C7: push eax
  loc_00E3A8C8: push 00000004h
  loc_00E3A8CA: call [00401038h] ; __vbaFreeVarList
  loc_00E3A8D0: add esp, 00000014h
  loc_00E3A8D3: ret
  loc_00E3A8D4: lea ecx, var_24
  loc_00E3A8D7: call [00401024h] ; __vbaFreeVar
  loc_00E3A8DD: ret
  loc_00E3A8DE: mov eax, Me
  loc_00E3A8E1: push eax
  loc_00E3A8E2: mov ecx, [eax]
  loc_00E3A8E4: call [ecx+00000008h]
  loc_00E3A8E7: mov eax, var_4
  loc_00E3A8EA: mov ecx, var_14
  loc_00E3A8ED: pop edi
  loc_00E3A8EE: pop esi
  loc_00E3A8EF: mov fs:[00000000h], ecx
  loc_00E3A8F6: pop ebx
  loc_00E3A8F7: mov esp, ebp
  loc_00E3A8F9: pop ebp
  loc_00E3A8FA: retn 0008h
End Sub

Private Sub Form_Load() 'E39EC0
  loc_00E39EC0: push ebp
  loc_00E39EC1: mov ebp, esp
  loc_00E39EC3: sub esp, 0000000Ch
  loc_00E39EC6: push 00402836h ; __vbaExceptHandler
  loc_00E39ECB: mov eax, fs:[00000000h]
  loc_00E39ED1: push eax
  loc_00E39ED2: mov fs:[00000000h], esp
  loc_00E39ED9: sub esp, 00000050h
  loc_00E39EDC: push ebx
  loc_00E39EDD: push esi
  loc_00E39EDE: push edi
  loc_00E39EDF: mov var_C, esp
  loc_00E39EE2: mov var_8, 004027B0h
  loc_00E39EE9: mov esi, Me
  loc_00E39EEC: mov eax, esi
  loc_00E39EEE: and eax, 00000001h
  loc_00E39EF1: mov var_4, eax
  loc_00E39EF4: and esi, FFFFFFFEh
  loc_00E39EF7: push esi
  loc_00E39EF8: mov Me, esi
  loc_00E39EFB: mov ecx, [esi]
  loc_00E39EFD: call [ecx+00000004h]
  loc_00E39F00: xor eax, eax
  loc_00E39F02: lea edx, var_4C
  loc_00E39F05: mov var_4C, eax
  loc_00E39F08: lea ecx, var_24
  loc_00E39F0B: mov var_24, eax
  loc_00E39F0E: mov var_28, eax
  loc_00E39F11: mov var_2C, eax
  loc_00E39F14: mov var_3C, eax
  loc_00E39F17: mov var_44, 006E15ECh ; "select * from lc"
  loc_00E39F1E: mov var_4C, 00000008h
  loc_00E39F25: call [00401204h] ; __vbaVarCopy
  loc_00E39F2B: lea edx, var_24
  loc_00E39F2E: push edx
  loc_00E39F2F: call [00401230h] ; __vbaStrVarCopy
  loc_00E39F35: sub esp, 00000010h
  loc_00E39F38: mov ecx, 00000008h
  loc_00E39F3D: mov edx, esp
  loc_00E39F3F: mov var_3C, ecx
  loc_00E39F42: mov var_34, eax
  loc_00E39F45: push 0000000Eh
  loc_00E39F47: mov [edx], ecx
  loc_00E39F49: mov ecx, var_38
  loc_00E39F4C: push esi
  loc_00E39F4D: mov [edx+00000004h], ecx
  loc_00E39F50: mov ecx, [esi]
  loc_00E39F52: mov [edx+00000008h], eax
  loc_00E39F55: mov eax, var_30
  loc_00E39F58: mov [edx+0000000Ch], eax
  loc_00E39F5B: call [ecx+0000033Ch]
  loc_00E39F61: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E39F67: lea edx, var_28
  loc_00E39F6A: push eax
  loc_00E39F6B: push edx
  loc_00E39F6C: call edi
  loc_00E39F6E: push eax
  loc_00E39F6F: call [00401238h] ; __vbaLateIdSt
  loc_00E39F75: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E39F7B: lea ecx, var_28
  loc_00E39F7E: call ebx
  loc_00E39F80: lea ecx, var_3C
  loc_00E39F83: call [00401024h] ; __vbaFreeVar
  loc_00E39F89: mov eax, [esi]
  loc_00E39F8B: push 00000000h
  loc_00E39F8D: push 00000019h
  loc_00E39F8F: push esi
  loc_00E39F90: call [eax+0000033Ch]
  loc_00E39F96: lea ecx, var_28
  loc_00E39F99: push eax
  loc_00E39F9A: push ecx
  loc_00E39F9B: call edi
  loc_00E39F9D: push eax
  loc_00E39F9E: call [00401030h] ; __vbaLateIdCall
  loc_00E39FA4: add esp, 0000000Ch
  loc_00E39FA7: lea ecx, var_28
  loc_00E39FAA: call ebx
  loc_00E39FAC: mov edx, [esi]
  loc_00E39FAE: push 006E05D8h
  loc_00E39FB3: push esi
  loc_00E39FB4: call [edx+0000033Ch]
  loc_00E39FBA: push eax
  loc_00E39FBB: lea eax, var_28
  loc_00E39FBE: push eax
  loc_00E39FBF: call edi
  loc_00E39FC1: push eax
  loc_00E39FC2: call [00401224h] ; __vbaCastObj
  loc_00E39FC8: mov var_34, eax
  loc_00E39FCB: mov var_3C, 0000000Dh
  loc_00E39FD2: lea ecx, var_3C
  loc_00E39FD5: push ecx
  loc_00E39FD6: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E39FDC: mov ecx, [eax]
  loc_00E39FDE: sub esp, 00000010h
  loc_00E39FE1: mov edx, esp
  loc_00E39FE3: push 00000000h
  loc_00E39FE5: push 0000002Ah
  loc_00E39FE7: mov [edx], ecx
  loc_00E39FE9: mov ecx, [eax+00000004h]
  loc_00E39FEC: push esi
  loc_00E39FED: mov [edx+00000004h], ecx
  loc_00E39FF0: mov ecx, [eax+00000008h]
  loc_00E39FF3: mov eax, [eax+0000000Ch]
  loc_00E39FF6: mov [edx+00000008h], ecx
  loc_00E39FF9: mov ecx, [esi]
  loc_00E39FFB: mov [edx+0000000Ch], eax
  loc_00E39FFE: call [ecx+00000340h]
  loc_00E3A004: lea edx, var_2C
  loc_00E3A007: push eax
  loc_00E3A008: push edx
  loc_00E3A009: call edi
  loc_00E3A00B: push eax
  loc_00E3A00C: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3A012: lea eax, var_2C
  loc_00E3A015: lea ecx, var_28
  loc_00E3A018: push eax
  loc_00E3A019: push ecx
  loc_00E3A01A: push 00000002h
  loc_00E3A01C: call [00401048h] ; __vbaFreeObjList
  loc_00E3A022: add esp, 00000028h
  loc_00E3A025: lea ecx, var_3C
  loc_00E3A028: call [00401024h] ; __vbaFreeVar
  loc_00E3A02E: mov edx, [esi]
  loc_00E3A030: push 00000000h
  loc_00E3A032: push FFFFFDDAh
  loc_00E3A037: push esi
  loc_00E3A038: call [edx+00000340h]
  loc_00E3A03E: push eax
  loc_00E3A03F: lea eax, var_28
  loc_00E3A042: push eax
  loc_00E3A043: call edi
  loc_00E3A045: push eax
  loc_00E3A046: call [00401030h] ; __vbaLateIdCall
  loc_00E3A04C: add esp, 0000000Ch
  loc_00E3A04F: lea ecx, var_28
  loc_00E3A052: call ebx
  loc_00E3A054: sub esp, 00000010h
  loc_00E3A057: mov ecx, 0000000Bh
  loc_00E3A05C: mov edx, esp
  loc_00E3A05E: mov var_4C, ecx
  loc_00E3A061: xor eax, eax
  loc_00E3A063: push 80010007h
  loc_00E3A068: mov [edx], ecx
  loc_00E3A06A: mov ecx, var_48
  loc_00E3A06D: mov var_44, eax
  loc_00E3A070: push esi
  loc_00E3A071: mov [edx+00000004h], ecx
  loc_00E3A074: mov ecx, [esi]
  loc_00E3A076: mov [edx+00000008h], eax
  loc_00E3A079: mov eax, var_40
  loc_00E3A07C: mov [edx+0000000Ch], eax
  loc_00E3A07F: call [ecx+00000314h]
  loc_00E3A085: lea edx, var_28
  loc_00E3A088: push eax
  loc_00E3A089: push edx
  loc_00E3A08A: call edi
  loc_00E3A08C: push eax
  loc_00E3A08D: call [00401238h] ; __vbaLateIdSt
  loc_00E3A093: lea ecx, var_28
  loc_00E3A096: call ebx
  loc_00E3A098: mov var_4, 00000000h
  loc_00E3A09F: push 00E3A0CDh
  loc_00E3A0A4: jmp 00E3A0C3h
  loc_00E3A0A6: lea eax, var_2C
  loc_00E3A0A9: lea ecx, var_28
  loc_00E3A0AC: push eax
  loc_00E3A0AD: push ecx
  loc_00E3A0AE: push 00000002h
  loc_00E3A0B0: call [00401048h] ; __vbaFreeObjList
  loc_00E3A0B6: add esp, 0000000Ch
  loc_00E3A0B9: lea ecx, var_3C
  loc_00E3A0BC: call [00401024h] ; __vbaFreeVar
  loc_00E3A0C2: ret
  loc_00E3A0C3: lea ecx, var_24
  loc_00E3A0C6: call [00401024h] ; __vbaFreeVar
  loc_00E3A0CC: ret
  loc_00E3A0CD: mov eax, Me
  loc_00E3A0D0: push eax
  loc_00E3A0D1: mov edx, [eax]
  loc_00E3A0D3: call [edx+00000008h]
  loc_00E3A0D6: mov eax, var_4
  loc_00E3A0D9: mov ecx, var_14
  loc_00E3A0DC: pop edi
  loc_00E3A0DD: pop esi
  loc_00E3A0DE: mov fs:[00000000h], ecx
  loc_00E3A0E5: pop ebx
  loc_00E3A0E6: mov esp, ebp
  loc_00E3A0E8: pop ebp
  loc_00E3A0E9: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E3A0F0
  loc_00E3A0F0: push ebp
  loc_00E3A0F1: mov ebp, esp
  loc_00E3A0F3: sub esp, 0000000Ch
  loc_00E3A0F6: push 00402836h ; __vbaExceptHandler
  loc_00E3A0FB: mov eax, fs:[00000000h]
  loc_00E3A101: push eax
  loc_00E3A102: mov fs:[00000000h], esp
  loc_00E3A109: sub esp, 0000005Ch
  loc_00E3A10C: push ebx
  loc_00E3A10D: push esi
  loc_00E3A10E: push edi
  loc_00E3A10F: mov var_C, esp
  loc_00E3A112: mov var_8, 004027C0h
  loc_00E3A119: mov esi, Me
  loc_00E3A11C: mov eax, esi
  loc_00E3A11E: and eax, 00000001h
  loc_00E3A121: mov var_4, eax
  loc_00E3A124: and esi, FFFFFFFEh
  loc_00E3A127: push esi
  loc_00E3A128: mov Me, esi
  loc_00E3A12B: mov ecx, [esi]
  loc_00E3A12D: call [ecx+00000004h]
  loc_00E3A130: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A136: xor eax, eax
  loc_00E3A138: mov var_18, eax
  loc_00E3A13B: mov var_4C, eax
  loc_00E3A13E: mov var_50, eax
  loc_00E3A141: mov edx, [esi]
  loc_00E3A143: lea eax, var_4C
  loc_00E3A146: push eax
  loc_00E3A147: push esi
  loc_00E3A148: call [edx+00000070h]
  loc_00E3A14B: test eax, eax
  loc_00E3A14D: fnclex
  loc_00E3A14F: jge 00E3A15Ch
  loc_00E3A151: push 00000070h
  loc_00E3A153: push 006DFFC8h
  loc_00E3A158: push esi
  loc_00E3A159: push eax
  loc_00E3A15A: call ebx
  loc_00E3A15C: fld real4 ptr var_4C
  loc_00E3A15F: fadd st0, real4 ptr [00401298h]
  loc_00E3A165: mov ecx, [esi]
  loc_00E3A167: push ecx
  loc_00E3A168: fnstsw ax
  loc_00E3A16A: test al, 0Dh
  loc_00E3A16C: jnz 00E3A330h
  loc_00E3A172: fstp real4 ptr [esp]
  loc_00E3A175: push esi
  loc_00E3A176: call [ecx+00000074h]
  loc_00E3A179: test eax, eax
  loc_00E3A17B: fnclex
  loc_00E3A17D: jge 00E3A18Ah
  loc_00E3A17F: push 00000074h
  loc_00E3A181: push 006DFFC8h
  loc_00E3A186: push esi
  loc_00E3A187: push eax
  loc_00E3A188: call ebx
  loc_00E3A18A: mov edx, [esi]
  loc_00E3A18C: lea eax, var_4C
  loc_00E3A18F: push eax
  loc_00E3A190: push esi
  loc_00E3A191: call [edx+00000070h]
  loc_00E3A194: test eax, eax
  loc_00E3A196: fnclex
  loc_00E3A198: jge 00E3A1A5h
  loc_00E3A19A: push 00000070h
  loc_00E3A19C: push 006DFFC8h
  loc_00E3A1A1: push esi
  loc_00E3A1A2: push eax
  loc_00E3A1A3: call ebx
  loc_00E3A1A5: mov ecx, [esi]
  loc_00E3A1A7: lea edx, var_50
  loc_00E3A1AA: push edx
  loc_00E3A1AB: push esi
  loc_00E3A1AC: call [ecx+00000078h]
  loc_00E3A1AF: test eax, eax
  loc_00E3A1B1: fnclex
  loc_00E3A1B3: jge 00E3A1C0h
  loc_00E3A1B5: push 00000078h
  loc_00E3A1B7: push 006DFFC8h
  loc_00E3A1BC: push esi
  loc_00E3A1BD: push eax
  loc_00E3A1BE: call ebx
  loc_00E3A1C0: sub esp, 00000010h
  loc_00E3A1C3: mov ecx, 0000000Ah
  loc_00E3A1C8: mov ebx, esp
  loc_00E3A1CA: mov eax, 80020004h
  loc_00E3A1CF: mov edx, eax
  loc_00E3A1D1: sub esp, 00000010h
  loc_00E3A1D4: mov [ebx], ecx
  loc_00E3A1D6: mov ecx, var_44
  loc_00E3A1D9: mov edi, [esi]
  loc_00E3A1DB: mov [ebx+00000004h], ecx
  loc_00E3A1DE: mov ecx, esp
  loc_00E3A1E0: sub esp, 00000010h
  loc_00E3A1E3: mov [ebx+00000008h], eax
  loc_00E3A1E6: mov eax, var_3C
  loc_00E3A1E9: mov [ebx+0000000Ch], eax
  loc_00E3A1EC: mov eax, 0000000Ah
  loc_00E3A1F1: mov [ecx], eax
  loc_00E3A1F3: mov eax, var_34
  loc_00E3A1F6: mov [ecx+00000004h], eax
  loc_00E3A1F9: mov eax, 00000004h
  loc_00E3A1FE: mov [ecx+00000008h], edx
  loc_00E3A201: mov edx, var_2C
  loc_00E3A204: mov [ecx+0000000Ch], edx
  loc_00E3A207: mov edx, var_24
  loc_00E3A20A: mov ecx, esp
  loc_00E3A20C: mov [ecx], eax
  loc_00E3A20E: mov eax, var_50
  loc_00E3A211: mov [ecx+00000004h], edx
  loc_00E3A214: mov [ecx+00000008h], eax
  loc_00E3A217: mov eax, var_1C
  loc_00E3A21A: mov [ecx+0000000Ch], eax
  loc_00E3A21D: mov ecx, var_4C
  loc_00E3A220: push ecx
  loc_00E3A221: push esi
  loc_00E3A222: call [edi+000002A4h]
  loc_00E3A228: test eax, eax
  loc_00E3A22A: fnclex
  loc_00E3A22C: jge 00E3A244h
  loc_00E3A22E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A234: push 000002A4h
  loc_00E3A239: push 006DFFC8h
  loc_00E3A23E: push esi
  loc_00E3A23F: push eax
  loc_00E3A240: call ebx
  loc_00E3A242: jmp 00E3A24Ah
  loc_00E3A244: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A24A: call [004010BCh] ; rtcDoEvents
  loc_00E3A250: mov edx, [esi]
  loc_00E3A252: lea eax, var_4C
  loc_00E3A255: push eax
  loc_00E3A256: push esi
  loc_00E3A257: call [edx+00000070h]
  loc_00E3A25A: test eax, eax
  loc_00E3A25C: fnclex
  loc_00E3A25E: jge 00E3A26Bh
  loc_00E3A260: push 00000070h
  loc_00E3A262: push 006DFFC8h
  loc_00E3A267: push esi
  loc_00E3A268: push eax
  loc_00E3A269: call ebx
  loc_00E3A26B: mov eax, [00E3D9CCh]
  loc_00E3A270: test eax, eax
  loc_00E3A272: jnz 00E3A284h
  loc_00E3A274: push 00E3D9CCh
  loc_00E3A279: push 006DC960h
  loc_00E3A27E: call [004011A0h] ; __vbaNew2
  loc_00E3A284: mov edi, [00E3D9CCh]
  loc_00E3A28A: lea edx, var_18
  loc_00E3A28D: push edx
  loc_00E3A28E: push edi
  loc_00E3A28F: mov ecx, [edi]
  loc_00E3A291: call [ecx+00000018h]
  loc_00E3A294: test eax, eax
  loc_00E3A296: fnclex
  loc_00E3A298: jge 00E3A2A5h
  loc_00E3A29A: push 00000018h
  loc_00E3A29C: push 006DC950h
  loc_00E3A2A1: push edi
  loc_00E3A2A2: push eax
  loc_00E3A2A3: call ebx
  loc_00E3A2A5: mov eax, var_18
  loc_00E3A2A8: lea edx, var_50
  loc_00E3A2AB: push edx
  loc_00E3A2AC: push eax
  loc_00E3A2AD: mov ecx, [eax]
  loc_00E3A2AF: mov edi, eax
  loc_00E3A2B1: call [ecx+00000098h]
  loc_00E3A2B7: test eax, eax
  loc_00E3A2B9: fnclex
  loc_00E3A2BB: jge 00E3A2CBh
  loc_00E3A2BD: push 00000098h
  loc_00E3A2C2: push 006DCAF0h
  loc_00E3A2C7: push edi
  loc_00E3A2C8: push eax
  loc_00E3A2C9: call ebx
  loc_00E3A2CB: fld real4 ptr var_4C
  loc_00E3A2CE: fcomp real4 ptr var_50
  loc_00E3A2D1: fnstsw ax
  loc_00E3A2D3: test ah, 41h
  loc_00E3A2D6: jz 00E3A2DFh
  loc_00E3A2D8: mov eax, 00000001h
  loc_00E3A2DD: jmp 00E3A2E1h
  loc_00E3A2DF: xor eax, eax
  loc_00E3A2E1: neg eax
  loc_00E3A2E3: lea ecx, var_18
  loc_00E3A2E6: mov edi, eax
  loc_00E3A2E8: call [00401254h] ; __vbaFreeObj
  loc_00E3A2EE: test di, di
  loc_00E3A2F1: jnz 00E3A141h
  loc_00E3A2F7: mov var_4, 00000000h
  loc_00E3A2FE: fwait
  loc_00E3A2FF: push 00E3A311h
  loc_00E3A304: jmp 00E3A310h
  loc_00E3A306: lea ecx, var_18
  loc_00E3A309: call [00401254h] ; __vbaFreeObj
  loc_00E3A30F: ret
  loc_00E3A310: ret
  loc_00E3A311: mov eax, Me
  loc_00E3A314: push eax
  loc_00E3A315: mov ecx, [eax]
  loc_00E3A317: call [ecx+00000008h]
  loc_00E3A31A: mov eax, var_4
  loc_00E3A31D: mov ecx, var_14
  loc_00E3A320: pop edi
  loc_00E3A321: pop esi
  loc_00E3A322: mov fs:[00000000h], ecx
  loc_00E3A329: pop ebx
  loc_00E3A32A: mov esp, ebp
  loc_00E3A32C: pop ebp
  loc_00E3A32D: retn 0008h
End Sub

Private Sub ctk_UnknownEvent_9 'E39470
  loc_00E39470: push ebp
  loc_00E39471: mov ebp, esp
  loc_00E39473: sub esp, 0000000Ch
  loc_00E39476: push 00402836h ; __vbaExceptHandler
  loc_00E3947B: mov eax, fs:[00000000h]
  loc_00E39481: push eax
  loc_00E39482: mov fs:[00000000h], esp
  loc_00E39489: sub esp, 000000C0h
  loc_00E3948F: push ebx
  loc_00E39490: push esi
  loc_00E39491: push edi
  loc_00E39492: mov var_C, esp
  loc_00E39495: mov var_8, 004027A0h
  loc_00E3949C: mov esi, Me
  loc_00E3949F: mov eax, esi
  loc_00E394A1: and eax, 00000001h
  loc_00E394A4: mov var_4, eax
  loc_00E394A7: and esi, FFFFFFFEh
  loc_00E394AA: push esi
  loc_00E394AB: mov Me, esi
  loc_00E394AE: mov ecx, [esi]
  loc_00E394B0: call [ecx+00000004h]
  loc_00E394B3: mov edx, [esi]
  loc_00E394B5: xor eax, eax
  loc_00E394B7: push eax
  loc_00E394B8: push 80010007h
  loc_00E394BD: push esi
  loc_00E394BE: mov var_18, eax
  loc_00E394C1: mov var_1C, eax
  loc_00E394C4: mov var_20, eax
  loc_00E394C7: mov var_24, eax
  loc_00E394CA: mov var_34, eax
  loc_00E394CD: mov var_44, eax
  loc_00E394D0: mov var_54, eax
  loc_00E394D3: mov var_64, eax
  loc_00E394D6: mov var_74, eax
  loc_00E394D9: mov var_84, eax
  loc_00E394DF: call [edx+00000314h]
  loc_00E394E5: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E394EB: push eax
  loc_00E394EC: lea eax, var_18
  loc_00E394EF: push eax
  loc_00E394F0: call ebx
  loc_00E394F2: lea ecx, var_34
  loc_00E394F5: push eax
  loc_00E394F6: push ecx
  loc_00E394F7: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E394FD: add esp, 00000010h
  loc_00E39500: push eax
  loc_00E39501: call [004010CCh] ; __vbaBoolVar
  loc_00E39507: mov di, ax
  loc_00E3950A: lea ecx, var_18
  loc_00E3950D: neg di
  loc_00E39510: sbb edi, edi
  loc_00E39512: inc edi
  loc_00E39513: neg edi
  loc_00E39515: call [00401254h] ; __vbaFreeObj
  loc_00E3951B: lea ecx, var_34
  loc_00E3951E: call [00401024h] ; __vbaFreeVar
  loc_00E39524: test di, di
  loc_00E39527: jz 00E395A1h
  loc_00E39529: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E3952F: mov ecx, 80020004h
  loc_00E39534: mov var_5C, ecx
  loc_00E39537: mov eax, 0000000Ah
  loc_00E3953C: mov var_4C, ecx
  loc_00E3953F: mov esi, 00000008h
  loc_00E39544: lea edx, var_84
  loc_00E3954A: lea ecx, var_44
  loc_00E3954D: mov var_64, eax
  loc_00E39550: mov var_54, eax
  loc_00E39553: mov var_7C, 006E20A4h ; "Alert !"
  loc_00E3955A: mov var_84, esi
  loc_00E39560: call edi
  loc_00E39562: lea edx, var_74
  loc_00E39565: lea ecx, var_34
  loc_00E39568: mov var_6C, 006E2038h ; "Tentukan Item yang ingin di cetak terlebih dahulu !"
  loc_00E3956F: mov var_74, esi
  loc_00E39572: call edi
  loc_00E39574: lea edx, var_64
  loc_00E39577: lea eax, var_54
  loc_00E3957A: push edx
  loc_00E3957B: lea ecx, var_44
  loc_00E3957E: push eax
  loc_00E3957F: push ecx
  loc_00E39580: lea edx, var_34
  loc_00E39583: push 00000010h
  loc_00E39585: push edx
  loc_00E39586: call [004010A8h] ; rtcMsgBox
  loc_00E3958C: lea eax, var_64
  loc_00E3958F: lea ecx, var_54
  loc_00E39592: push eax
  loc_00E39593: lea edx, var_44
  loc_00E39596: push ecx
  loc_00E39597: lea eax, var_34
  loc_00E3959A: push edx
  loc_00E3959B: push eax
  loc_00E3959C: jmp 00E39E4Fh
  loc_00E395A1: mov ecx, [esi]
  loc_00E395A3: push 00000000h
  loc_00E395A5: push 80010007h
  loc_00E395AA: push esi
  loc_00E395AB: call [ecx+00000314h]
  loc_00E395B1: lea edx, var_18
  loc_00E395B4: push eax
  loc_00E395B5: push edx
  loc_00E395B6: call ebx
  loc_00E395B8: push eax
  loc_00E395B9: lea eax, var_34
  loc_00E395BC: push eax
  loc_00E395BD: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E395C3: add esp, 00000010h
  loc_00E395C6: push eax
  loc_00E395C7: call [004010CCh] ; __vbaBoolVar
  loc_00E395CD: xor ecx, ecx
  loc_00E395CF: cmp ax, FFFFFFh
  loc_00E395D3: setz cl
  loc_00E395D6: neg ecx
  loc_00E395D8: mov di, cx
  loc_00E395DB: lea ecx, var_18
  loc_00E395DE: call [00401254h] ; __vbaFreeObj
  loc_00E395E4: lea ecx, var_34
  loc_00E395E7: call [00401024h] ; __vbaFreeVar
  loc_00E395ED: test di, di
  loc_00E395F0: jz 00E39DDCh
  loc_00E395F6: mov edx, [esi]
  loc_00E395F8: push 006DCBD8h
  loc_00E395FD: push 00000000h
  loc_00E395FF: push 00000018h
  loc_00E39601: push esi
  loc_00E39602: call [edx+0000033Ch]
  loc_00E39608: push eax
  loc_00E39609: lea eax, var_18
  loc_00E3960C: push eax
  loc_00E3960D: call ebx
  loc_00E3960F: lea ecx, var_34
  loc_00E39612: push eax
  loc_00E39613: push ecx
  loc_00E39614: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E3961A: add esp, 00000010h
  loc_00E3961D: push eax
  loc_00E3961E: call [00401120h] ; __vbaCastObjVar
  loc_00E39624: lea edx, var_1C
  loc_00E39627: push eax
  loc_00E39628: push edx
  loc_00E39629: call ebx
  loc_00E3962B: mov edi, eax
  loc_00E3962D: lea ecx, var_20
  loc_00E39630: push ecx
  loc_00E39631: push edi
  loc_00E39632: mov eax, [edi]
  loc_00E39634: call [eax+00000054h]
  loc_00E39637: test eax, eax
  loc_00E39639: fnclex
  loc_00E3963B: jge 00E3964Ch
  loc_00E3963D: push 00000054h
  loc_00E3963F: push 006DCBE8h
  loc_00E39644: push edi
  loc_00E39645: push eax
  loc_00E39646: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3964C: lea edi, var_24
  loc_00E3964F: mov eax, var_20
  loc_00E39652: push edi
  loc_00E39653: mov ecx, 00000002h
  loc_00E39658: sub esp, 00000010h
  loc_00E3965B: mov var_74, ecx
  loc_00E3965E: mov edi, esp
  loc_00E39660: mov var_6C, 00000005h
  loc_00E39667: mov edx, [eax]
  loc_00E39669: push eax
  loc_00E3966A: mov [edi], ecx
  loc_00E3966C: mov ecx, var_70
  loc_00E3966F: mov var_B0, eax
  loc_00E39675: mov [edi+00000004h], ecx
  loc_00E39678: mov ecx, var_6C
  loc_00E3967B: mov [edi+00000008h], ecx
  loc_00E3967E: mov ecx, var_68
  loc_00E39681: mov [edi+0000000Ch], ecx
  loc_00E39684: call [edx+00000028h]
  loc_00E39687: test eax, eax
  loc_00E39689: fnclex
  loc_00E3968B: jge 00E396A2h
  loc_00E3968D: mov edx, var_B0
  loc_00E39693: push 00000028h
  loc_00E39695: push 006E09E8h
  loc_00E3969A: push edx
  loc_00E3969B: push eax
  loc_00E3969C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E396A2: mov eax, var_24
  loc_00E396A5: lea edx, var_44
  loc_00E396A8: push edx
  loc_00E396A9: push eax
  loc_00E396AA: mov ecx, [eax]
  loc_00E396AC: mov edi, eax
  loc_00E396AE: call [ecx+00000034h]
  loc_00E396B1: test eax, eax
  loc_00E396B3: fnclex
  loc_00E396B5: jge 00E396C6h
  loc_00E396B7: push 00000034h
  loc_00E396B9: push 006E09F8h
  loc_00E396BE: push edi
  loc_00E396BF: push eax
  loc_00E396C0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E396C6: lea eax, var_44
  loc_00E396C9: lea ecx, var_84
  loc_00E396CF: push eax
  loc_00E396D0: push ecx
  loc_00E396D1: mov var_7C, 006E1728h ; "Lunas"
  loc_00E396D8: mov var_84, 00008008h
  loc_00E396E2: call [00401108h] ; __vbaVarTstEq
  loc_00E396E8: mov edi, eax
  loc_00E396EA: lea edx, var_24
  loc_00E396ED: lea eax, var_20
  loc_00E396F0: push edx
  loc_00E396F1: lea ecx, var_1C
  loc_00E396F4: push eax
  loc_00E396F5: lea edx, var_18
  loc_00E396F8: push ecx
  loc_00E396F9: push edx
  loc_00E396FA: push 00000004h
  loc_00E396FC: call [00401048h] ; __vbaFreeObjList
  loc_00E39702: lea eax, var_44
  loc_00E39705: lea ecx, var_34
  loc_00E39708: push eax
  loc_00E39709: push ecx
  loc_00E3970A: push 00000002h
  loc_00E3970C: call [00401038h] ; __vbaFreeVarList
  loc_00E39712: add esp, 00000020h
  loc_00E39715: test di, di
  loc_00E39718: jz 00E3987Bh
  loc_00E3971E: mov eax, [00E3D164h]
  loc_00E39723: test eax, eax
  loc_00E39725: jnz 00E39737h
  loc_00E39727: push 00E3D164h
  loc_00E3972C: push 006CB168h
  loc_00E39731: call [004011A0h] ; __vbaNew2
  loc_00E39737: mov edi, [00E3D164h]
  loc_00E3973D: mov eax, [esi]
  loc_00E3973F: push 006E05D8h
  loc_00E39744: push esi
  loc_00E39745: mov edx, [edi]
  loc_00E39747: mov var_CC, edx
  loc_00E3974D: call [eax+0000033Ch]
  loc_00E39753: lea ecx, var_18
  loc_00E39756: push eax
  loc_00E39757: push ecx
  loc_00E39758: call ebx
  loc_00E3975A: push eax
  loc_00E3975B: call [00401224h] ; __vbaCastObj
  loc_00E39761: lea edx, var_1C
  loc_00E39764: push eax
  loc_00E39765: push edx
  loc_00E39766: call ebx
  loc_00E39768: push eax
  loc_00E39769: mov eax, var_CC
  loc_00E3976F: push edi
  loc_00E39770: call [eax+00000028h]
  loc_00E39773: test eax, eax
  loc_00E39775: fnclex
  loc_00E39777: jge 00E39788h
  loc_00E39779: push 00000028h
  loc_00E3977B: push 006DD034h
  loc_00E39780: push edi
  loc_00E39781: push eax
  loc_00E39782: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39788: lea ecx, var_1C
  loc_00E3978B: lea edx, var_18
  loc_00E3978E: push ecx
  loc_00E3978F: push edx
  loc_00E39790: push 00000002h
  loc_00E39792: call [00401048h] ; __vbaFreeObjList
  loc_00E39798: mov eax, [00E3D164h]
  loc_00E3979D: add esp, 0000000Ch
  loc_00E397A0: test eax, eax
  loc_00E397A2: jnz 00E397B8h
  loc_00E397A4: mov edi, [004011A0h] ; __vbaNew2
  loc_00E397AA: push 00E3D164h
  loc_00E397AF: push 006CB168h
  loc_00E397B4: call edi
  loc_00E397B6: jmp 00E397BEh
  loc_00E397B8: mov edi, [004011A0h] ; __vbaNew2
  loc_00E397BE: mov esi, [00E3D164h]
  loc_00E397C4: push esi
  loc_00E397C5: mov eax, [esi]
  loc_00E397C7: call [eax+00000088h]
  loc_00E397CD: test eax, eax
  loc_00E397CF: fnclex
  loc_00E397D1: jge 00E397E5h
  loc_00E397D3: push 00000088h
  loc_00E397D8: push 006DD034h
  loc_00E397DD: push esi
  loc_00E397DE: push eax
  loc_00E397DF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E397E5: mov eax, [00E3D9CCh]
  loc_00E397EA: test eax, eax
  loc_00E397EC: jnz 00E397FAh
  loc_00E397EE: push 00E3D9CCh
  loc_00E397F3: push 006DC960h
  loc_00E397F8: call edi
  loc_00E397FA: mov eax, [00E3D164h]
  loc_00E397FF: mov esi, [00E3D9CCh]
  loc_00E39805: test eax, eax
  loc_00E39807: jnz 00E39815h
  loc_00E39809: push 00E3D164h
  loc_00E3980E: push 006CB168h
  loc_00E39813: call edi
  loc_00E39815: mov ecx, [00E3D164h]
  loc_00E3981B: mov ebx, [esi]
  loc_00E3981D: lea edx, var_18
  loc_00E39820: push ecx
  loc_00E39821: push edx
  loc_00E39822: call [004010B4h] ; __vbaObjSetAddref
  loc_00E39828: push eax
  loc_00E39829: push esi
  loc_00E3982A: call [ebx+0000000Ch]
  loc_00E3982D: test eax, eax
  loc_00E3982F: fnclex
  loc_00E39831: jge 00E39842h
  loc_00E39833: push 0000000Ch
  loc_00E39835: push 006DC950h
  loc_00E3983A: push esi
  loc_00E3983B: push eax
  loc_00E3983C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39842: lea ecx, var_18
  loc_00E39845: call [00401254h] ; __vbaFreeObj
  loc_00E3984B: mov eax, [00E3D164h]
  loc_00E39850: test eax, eax
  loc_00E39852: jnz 00E39860h
  loc_00E39854: push 00E3D164h
  loc_00E39859: push 006CB168h
  loc_00E3985E: call edi
  loc_00E39860: mov eax, [00E3D164h]
  loc_00E39865: push 00000000h
  loc_00E39867: push 80011003h
  loc_00E3986C: push eax
  loc_00E3986D: call [00401030h] ; __vbaLateIdCall
  loc_00E39873: add esp, 0000000Ch
  loc_00E39876: jmp 00E39E5Ah
  loc_00E3987B: mov ecx, [esi]
  loc_00E3987D: push 006DCBD8h
  loc_00E39882: push 00000000h
  loc_00E39884: push 00000018h
  loc_00E39886: push esi
  loc_00E39887: call [ecx+0000033Ch]
  loc_00E3988D: lea edx, var_18
  loc_00E39890: push eax
  loc_00E39891: push edx
  loc_00E39892: call ebx
  loc_00E39894: push eax
  loc_00E39895: lea eax, var_34
  loc_00E39898: push eax
  loc_00E39899: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E3989F: add esp, 00000010h
  loc_00E398A2: push eax
  loc_00E398A3: call [00401120h] ; __vbaCastObjVar
  loc_00E398A9: lea ecx, var_1C
  loc_00E398AC: push eax
  loc_00E398AD: push ecx
  loc_00E398AE: call ebx
  loc_00E398B0: mov edi, eax
  loc_00E398B2: lea eax, var_20
  loc_00E398B5: push eax
  loc_00E398B6: push edi
  loc_00E398B7: mov edx, [edi]
  loc_00E398B9: call [edx+00000054h]
  loc_00E398BC: test eax, eax
  loc_00E398BE: fnclex
  loc_00E398C0: jge 00E398D1h
  loc_00E398C2: push 00000054h
  loc_00E398C4: push 006DCBE8h
  loc_00E398C9: push edi
  loc_00E398CA: push eax
  loc_00E398CB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E398D1: lea edi, var_24
  loc_00E398D4: mov eax, var_20
  loc_00E398D7: push edi
  loc_00E398D8: mov ecx, 00000002h
  loc_00E398DD: sub esp, 00000010h
  loc_00E398E0: mov var_74, ecx
  loc_00E398E3: mov edi, esp
  loc_00E398E5: mov var_6C, 00000005h
  loc_00E398EC: mov edx, [eax]
  loc_00E398EE: push eax
  loc_00E398EF: mov [edi], ecx
  loc_00E398F1: mov ecx, var_70
  loc_00E398F4: mov var_B0, eax
  loc_00E398FA: mov [edi+00000004h], ecx
  loc_00E398FD: mov ecx, var_6C
  loc_00E39900: mov [edi+00000008h], ecx
  loc_00E39903: mov ecx, var_68
  loc_00E39906: mov [edi+0000000Ch], ecx
  loc_00E39909: call [edx+00000028h]
  loc_00E3990C: test eax, eax
  loc_00E3990E: fnclex
  loc_00E39910: jge 00E39927h
  loc_00E39912: mov edx, var_B0
  loc_00E39918: push 00000028h
  loc_00E3991A: push 006E09E8h
  loc_00E3991F: push edx
  loc_00E39920: push eax
  loc_00E39921: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39927: mov eax, var_24
  loc_00E3992A: lea edx, var_44
  loc_00E3992D: push edx
  loc_00E3992E: push eax
  loc_00E3992F: mov ecx, [eax]
  loc_00E39931: mov edi, eax
  loc_00E39933: call [ecx+00000034h]
  loc_00E39936: test eax, eax
  loc_00E39938: fnclex
  loc_00E3993A: jge 00E3994Bh
  loc_00E3993C: push 00000034h
  loc_00E3993E: push 006E09F8h
  loc_00E39943: push edi
  loc_00E39944: push eax
  loc_00E39945: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3994B: lea eax, var_44
  loc_00E3994E: lea ecx, var_84
  loc_00E39954: push eax
  loc_00E39955: push ecx
  loc_00E39956: mov var_7C, 006E1710h ; "Mencicil"
  loc_00E3995D: mov var_84, 00008008h
  loc_00E39967: call [00401108h] ; __vbaVarTstEq
  loc_00E3996D: mov edi, eax
  loc_00E3996F: lea edx, var_24
  loc_00E39972: lea eax, var_20
  loc_00E39975: push edx
  loc_00E39976: lea ecx, var_1C
  loc_00E39979: push eax
  loc_00E3997A: lea edx, var_18
  loc_00E3997D: push ecx
  loc_00E3997E: push edx
  loc_00E3997F: push 00000004h
  loc_00E39981: call [00401048h] ; __vbaFreeObjList
  loc_00E39987: lea eax, var_44
  loc_00E3998A: lea ecx, var_34
  loc_00E3998D: push eax
  loc_00E3998E: push ecx
  loc_00E3998F: push 00000002h
  loc_00E39991: call [00401038h] ; __vbaFreeVarList
  loc_00E39997: add esp, 00000020h
  loc_00E3999A: test di, di
  loc_00E3999D: jz 00E39B00h
  loc_00E399A3: mov eax, [00E3D038h]
  loc_00E399A8: test eax, eax
  loc_00E399AA: jnz 00E399BCh
  loc_00E399AC: push 00E3D038h
  loc_00E399B1: push 006CB260h
  loc_00E399B6: call [004011A0h] ; __vbaNew2
  loc_00E399BC: mov edi, [00E3D038h]
  loc_00E399C2: mov eax, [esi]
  loc_00E399C4: push 006E05D8h
  loc_00E399C9: push esi
  loc_00E399CA: mov edx, [edi]
  loc_00E399CC: mov var_D0, edx
  loc_00E399D2: call [eax+0000033Ch]
  loc_00E399D8: lea ecx, var_18
  loc_00E399DB: push eax
  loc_00E399DC: push ecx
  loc_00E399DD: call ebx
  loc_00E399DF: push eax
  loc_00E399E0: call [00401224h] ; __vbaCastObj
  loc_00E399E6: lea edx, var_1C
  loc_00E399E9: push eax
  loc_00E399EA: push edx
  loc_00E399EB: call ebx
  loc_00E399ED: push eax
  loc_00E399EE: mov eax, var_D0
  loc_00E399F4: push edi
  loc_00E399F5: call [eax+00000028h]
  loc_00E399F8: test eax, eax
  loc_00E399FA: fnclex
  loc_00E399FC: jge 00E39A0Dh
  loc_00E399FE: push 00000028h
  loc_00E39A00: push 006DD034h
  loc_00E39A05: push edi
  loc_00E39A06: push eax
  loc_00E39A07: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39A0D: lea ecx, var_1C
  loc_00E39A10: lea edx, var_18
  loc_00E39A13: push ecx
  loc_00E39A14: push edx
  loc_00E39A15: push 00000002h
  loc_00E39A17: call [00401048h] ; __vbaFreeObjList
  loc_00E39A1D: mov eax, [00E3D038h]
  loc_00E39A22: add esp, 0000000Ch
  loc_00E39A25: test eax, eax
  loc_00E39A27: jnz 00E39A3Dh
  loc_00E39A29: mov edi, [004011A0h] ; __vbaNew2
  loc_00E39A2F: push 00E3D038h
  loc_00E39A34: push 006CB260h
  loc_00E39A39: call edi
  loc_00E39A3B: jmp 00E39A43h
  loc_00E39A3D: mov edi, [004011A0h] ; __vbaNew2
  loc_00E39A43: mov esi, [00E3D038h]
  loc_00E39A49: push esi
  loc_00E39A4A: mov eax, [esi]
  loc_00E39A4C: call [eax+00000088h]
  loc_00E39A52: test eax, eax
  loc_00E39A54: fnclex
  loc_00E39A56: jge 00E39A6Ah
  loc_00E39A58: push 00000088h
  loc_00E39A5D: push 006DD034h
  loc_00E39A62: push esi
  loc_00E39A63: push eax
  loc_00E39A64: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39A6A: mov eax, [00E3D9CCh]
  loc_00E39A6F: test eax, eax
  loc_00E39A71: jnz 00E39A7Fh
  loc_00E39A73: push 00E3D9CCh
  loc_00E39A78: push 006DC960h
  loc_00E39A7D: call edi
  loc_00E39A7F: mov eax, [00E3D038h]
  loc_00E39A84: mov esi, [00E3D9CCh]
  loc_00E39A8A: test eax, eax
  loc_00E39A8C: jnz 00E39A9Ah
  loc_00E39A8E: push 00E3D038h
  loc_00E39A93: push 006CB260h
  loc_00E39A98: call edi
  loc_00E39A9A: mov ecx, [00E3D038h]
  loc_00E39AA0: mov ebx, [esi]
  loc_00E39AA2: lea edx, var_18
  loc_00E39AA5: push ecx
  loc_00E39AA6: push edx
  loc_00E39AA7: call [004010B4h] ; __vbaObjSetAddref
  loc_00E39AAD: push eax
  loc_00E39AAE: push esi
  loc_00E39AAF: call [ebx+0000000Ch]
  loc_00E39AB2: test eax, eax
  loc_00E39AB4: fnclex
  loc_00E39AB6: jge 00E39AC7h
  loc_00E39AB8: push 0000000Ch
  loc_00E39ABA: push 006DC950h
  loc_00E39ABF: push esi
  loc_00E39AC0: push eax
  loc_00E39AC1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39AC7: lea ecx, var_18
  loc_00E39ACA: call [00401254h] ; __vbaFreeObj
  loc_00E39AD0: mov eax, [00E3D038h]
  loc_00E39AD5: test eax, eax
  loc_00E39AD7: jnz 00E39AE5h
  loc_00E39AD9: push 00E3D038h
  loc_00E39ADE: push 006CB260h
  loc_00E39AE3: call edi
  loc_00E39AE5: mov eax, [00E3D038h]
  loc_00E39AEA: push 00000000h
  loc_00E39AEC: push 80011003h
  loc_00E39AF1: push eax
  loc_00E39AF2: call [00401030h] ; __vbaLateIdCall
  loc_00E39AF8: add esp, 0000000Ch
  loc_00E39AFB: jmp 00E39E5Ah
  loc_00E39B00: mov ecx, [esi]
  loc_00E39B02: push 006DCBD8h
  loc_00E39B07: push 00000000h
  loc_00E39B09: push 00000018h
  loc_00E39B0B: push esi
  loc_00E39B0C: call [ecx+0000033Ch]
  loc_00E39B12: lea edx, var_18
  loc_00E39B15: push eax
  loc_00E39B16: push edx
  loc_00E39B17: call ebx
  loc_00E39B19: push eax
  loc_00E39B1A: lea eax, var_34
  loc_00E39B1D: push eax
  loc_00E39B1E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E39B24: add esp, 00000010h
  loc_00E39B27: push eax
  loc_00E39B28: call [00401120h] ; __vbaCastObjVar
  loc_00E39B2E: lea ecx, var_1C
  loc_00E39B31: push eax
  loc_00E39B32: push ecx
  loc_00E39B33: call ebx
  loc_00E39B35: mov edi, eax
  loc_00E39B37: lea eax, var_20
  loc_00E39B3A: push eax
  loc_00E39B3B: push edi
  loc_00E39B3C: mov edx, [edi]
  loc_00E39B3E: call [edx+00000054h]
  loc_00E39B41: test eax, eax
  loc_00E39B43: fnclex
  loc_00E39B45: jge 00E39B56h
  loc_00E39B47: push 00000054h
  loc_00E39B49: push 006DCBE8h
  loc_00E39B4E: push edi
  loc_00E39B4F: push eax
  loc_00E39B50: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39B56: lea edi, var_24
  loc_00E39B59: mov eax, var_20
  loc_00E39B5C: push edi
  loc_00E39B5D: mov ecx, 00000002h
  loc_00E39B62: sub esp, 00000010h
  loc_00E39B65: mov var_74, ecx
  loc_00E39B68: mov edi, esp
  loc_00E39B6A: mov var_6C, 00000005h
  loc_00E39B71: mov edx, [eax]
  loc_00E39B73: push eax
  loc_00E39B74: mov [edi], ecx
  loc_00E39B76: mov ecx, var_70
  loc_00E39B79: mov var_B0, eax
  loc_00E39B7F: mov [edi+00000004h], ecx
  loc_00E39B82: mov ecx, var_6C
  loc_00E39B85: mov [edi+00000008h], ecx
  loc_00E39B88: mov ecx, var_68
  loc_00E39B8B: mov [edi+0000000Ch], ecx
  loc_00E39B8E: call [edx+00000028h]
  loc_00E39B91: test eax, eax
  loc_00E39B93: fnclex
  loc_00E39B95: jge 00E39BACh
  loc_00E39B97: mov edx, var_B0
  loc_00E39B9D: push 00000028h
  loc_00E39B9F: push 006E09E8h
  loc_00E39BA4: push edx
  loc_00E39BA5: push eax
  loc_00E39BA6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39BAC: mov eax, var_24
  loc_00E39BAF: lea edx, var_44
  loc_00E39BB2: push edx
  loc_00E39BB3: push eax
  loc_00E39BB4: mov ecx, [eax]
  loc_00E39BB6: mov edi, eax
  loc_00E39BB8: call [ecx+00000034h]
  loc_00E39BBB: test eax, eax
  loc_00E39BBD: fnclex
  loc_00E39BBF: jge 00E39BD0h
  loc_00E39BC1: push 00000034h
  loc_00E39BC3: push 006E09F8h
  loc_00E39BC8: push edi
  loc_00E39BC9: push eax
  loc_00E39BCA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39BD0: lea eax, var_44
  loc_00E39BD3: lea ecx, var_84
  loc_00E39BD9: push eax
  loc_00E39BDA: push ecx
  loc_00E39BDB: mov var_7C, 006E1C10h ; "Mencicil (Lunas)"
  loc_00E39BE2: mov var_84, 00008008h
  loc_00E39BEC: call [00401108h] ; __vbaVarTstEq
  loc_00E39BF2: mov edi, eax
  loc_00E39BF4: lea edx, var_24
  loc_00E39BF7: lea eax, var_20
  loc_00E39BFA: push edx
  loc_00E39BFB: lea ecx, var_1C
  loc_00E39BFE: push eax
  loc_00E39BFF: lea edx, var_18
  loc_00E39C02: push ecx
  loc_00E39C03: push edx
  loc_00E39C04: push 00000004h
  loc_00E39C06: call [00401048h] ; __vbaFreeObjList
  loc_00E39C0C: lea eax, var_44
  loc_00E39C0F: lea ecx, var_34
  loc_00E39C12: push eax
  loc_00E39C13: push ecx
  loc_00E39C14: push 00000002h
  loc_00E39C16: call [00401038h] ; __vbaFreeVarList
  loc_00E39C1C: add esp, 00000020h
  loc_00E39C1F: test di, di
  loc_00E39C22: jz 00E39D67h
  loc_00E39C28: mov eax, [00E3D164h]
  loc_00E39C2D: test eax, eax
  loc_00E39C2F: jnz 00E39C41h
  loc_00E39C31: push 00E3D164h
  loc_00E39C36: push 006CB168h
  loc_00E39C3B: call [004011A0h] ; __vbaNew2
  loc_00E39C41: mov edi, [00E3D164h]
  loc_00E39C47: mov eax, [esi]
  loc_00E39C49: push 006E05D8h
  loc_00E39C4E: push esi
  loc_00E39C4F: mov edx, [edi]
  loc_00E39C51: mov var_D4, edx
  loc_00E39C57: call [eax+0000033Ch]
  loc_00E39C5D: lea ecx, var_18
  loc_00E39C60: push eax
  loc_00E39C61: push ecx
  loc_00E39C62: call ebx
  loc_00E39C64: push eax
  loc_00E39C65: call [00401224h] ; __vbaCastObj
  loc_00E39C6B: lea edx, var_1C
  loc_00E39C6E: push eax
  loc_00E39C6F: push edx
  loc_00E39C70: call ebx
  loc_00E39C72: push eax
  loc_00E39C73: mov eax, var_D4
  loc_00E39C79: push edi
  loc_00E39C7A: call [eax+00000028h]
  loc_00E39C7D: test eax, eax
  loc_00E39C7F: fnclex
  loc_00E39C81: jge 00E39C92h
  loc_00E39C83: push 00000028h
  loc_00E39C85: push 006DD034h
  loc_00E39C8A: push edi
  loc_00E39C8B: push eax
  loc_00E39C8C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39C92: lea ecx, var_1C
  loc_00E39C95: lea edx, var_18
  loc_00E39C98: push ecx
  loc_00E39C99: push edx
  loc_00E39C9A: push 00000002h
  loc_00E39C9C: call [00401048h] ; __vbaFreeObjList
  loc_00E39CA2: mov eax, [00E3D164h]
  loc_00E39CA7: add esp, 0000000Ch
  loc_00E39CAA: test eax, eax
  loc_00E39CAC: jnz 00E39CC2h
  loc_00E39CAE: mov edi, [004011A0h] ; __vbaNew2
  loc_00E39CB4: push 00E3D164h
  loc_00E39CB9: push 006CB168h
  loc_00E39CBE: call edi
  loc_00E39CC0: jmp 00E39CC8h
  loc_00E39CC2: mov edi, [004011A0h] ; __vbaNew2
  loc_00E39CC8: mov esi, [00E3D164h]
  loc_00E39CCE: push esi
  loc_00E39CCF: mov eax, [esi]
  loc_00E39CD1: call [eax+00000088h]
  loc_00E39CD7: test eax, eax
  loc_00E39CD9: fnclex
  loc_00E39CDB: jge 00E39CEFh
  loc_00E39CDD: push 00000088h
  loc_00E39CE2: push 006DD034h
  loc_00E39CE7: push esi
  loc_00E39CE8: push eax
  loc_00E39CE9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39CEF: mov eax, [00E3D9CCh]
  loc_00E39CF4: test eax, eax
  loc_00E39CF6: jnz 00E39D04h
  loc_00E39CF8: push 00E3D9CCh
  loc_00E39CFD: push 006DC960h
  loc_00E39D02: call edi
  loc_00E39D04: mov eax, [00E3D164h]
  loc_00E39D09: mov esi, [00E3D9CCh]
  loc_00E39D0F: test eax, eax
  loc_00E39D11: jnz 00E39D1Fh
  loc_00E39D13: push 00E3D164h
  loc_00E39D18: push 006CB168h
  loc_00E39D1D: call edi
  loc_00E39D1F: mov ecx, [00E3D164h]
  loc_00E39D25: mov ebx, [esi]
  loc_00E39D27: lea edx, var_18
  loc_00E39D2A: push ecx
  loc_00E39D2B: push edx
  loc_00E39D2C: call [004010B4h] ; __vbaObjSetAddref
  loc_00E39D32: push eax
  loc_00E39D33: push esi
  loc_00E39D34: call [ebx+0000000Ch]
  loc_00E39D37: test eax, eax
  loc_00E39D39: fnclex
  loc_00E39D3B: jge 00E39D4Ch
  loc_00E39D3D: push 0000000Ch
  loc_00E39D3F: push 006DC950h
  loc_00E39D44: push esi
  loc_00E39D45: push eax
  loc_00E39D46: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E39D4C: lea ecx, var_18
  loc_00E39D4F: call [00401254h] ; __vbaFreeObj
  loc_00E39D55: mov eax, [00E3D164h]
  loc_00E39D5A: test eax, eax
  loc_00E39D5C: jnz 00E39860h
  loc_00E39D62: jmp 00E39854h
  loc_00E39D67: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E39D6D: mov ecx, 80020004h
  loc_00E39D72: mov var_5C, ecx
  loc_00E39D75: mov eax, 0000000Ah
  loc_00E39D7A: mov var_4C, ecx
  loc_00E39D7D: mov esi, 00000008h
  loc_00E39D82: lea edx, var_84
  loc_00E39D88: lea ecx, var_44
  loc_00E39D8B: mov var_64, eax
  loc_00E39D8E: mov var_54, eax
  loc_00E39D91: mov var_7C, 006E20D4h ; "ERROR 404"
  loc_00E39D98: mov var_84, esi
  loc_00E39D9E: call edi
  loc_00E39DA0: lea edx, var_74
  loc_00E39DA3: lea ecx, var_34
  loc_00E39DA6: mov var_6C, 006E20B8h ; "Ada [CRASH]"
  loc_00E39DAD: mov var_74, esi
  loc_00E39DB0: call edi
  loc_00E39DB2: lea ecx, var_64
  loc_00E39DB5: lea edx, var_54
  loc_00E39DB8: push ecx
  loc_00E39DB9: lea eax, var_44
  loc_00E39DBC: push edx
  loc_00E39DBD: push eax
  loc_00E39DBE: lea ecx, var_34
  loc_00E39DC1: push 00000010h
  loc_00E39DC3: push ecx
  loc_00E39DC4: call [004010A8h] ; rtcMsgBox
  loc_00E39DCA: lea edx, var_64
  loc_00E39DCD: lea eax, var_54
  loc_00E39DD0: push edx
  loc_00E39DD1: lea ecx, var_44
  loc_00E39DD4: push eax
  loc_00E39DD5: lea edx, var_34
  loc_00E39DD8: push ecx
  loc_00E39DD9: push edx
  loc_00E39DDA: jmp 00E39E4Fh
  loc_00E39DDC: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E39DE2: mov ecx, 80020004h
  loc_00E39DE7: mov var_5C, ecx
  loc_00E39DEA: mov eax, 0000000Ah
  loc_00E39DEF: mov var_4C, ecx
  loc_00E39DF2: mov esi, 00000008h
  loc_00E39DF7: lea edx, var_84
  loc_00E39DFD: lea ecx, var_44
  loc_00E39E00: mov var_64, eax
  loc_00E39E03: mov var_54, eax
  loc_00E39E06: mov var_7C, 006E20D4h ; "ERROR 404"
  loc_00E39E0D: mov var_84, esi
  loc_00E39E13: call edi
  loc_00E39E15: lea edx, var_74
  loc_00E39E18: lea ecx, var_34
  loc_00E39E1B: mov var_6C, 006E20B8h ; "Ada [CRASH]"
  loc_00E39E22: mov var_74, esi
  loc_00E39E25: call edi
  loc_00E39E27: lea eax, var_64
  loc_00E39E2A: lea ecx, var_54
  loc_00E39E2D: push eax
  loc_00E39E2E: lea edx, var_44
  loc_00E39E31: push ecx
  loc_00E39E32: push edx
  loc_00E39E33: lea eax, var_34
  loc_00E39E36: push 00000010h
  loc_00E39E38: push eax
  loc_00E39E39: call [004010A8h] ; rtcMsgBox
  loc_00E39E3F: lea ecx, var_64
  loc_00E39E42: lea edx, var_54
  loc_00E39E45: push ecx
  loc_00E39E46: lea eax, var_44
  loc_00E39E49: push edx
  loc_00E39E4A: lea ecx, var_34
  loc_00E39E4D: push eax
  loc_00E39E4E: push ecx
  loc_00E39E4F: push 00000004h
  loc_00E39E51: call [00401038h] ; __vbaFreeVarList
  loc_00E39E57: add esp, 00000014h
  loc_00E39E5A: mov var_4, 00000000h
  loc_00E39E61: push 00E39E9Dh
  loc_00E39E66: jmp 00E39E9Ch
  loc_00E39E68: lea edx, var_24
  loc_00E39E6B: lea eax, var_20
  loc_00E39E6E: push edx
  loc_00E39E6F: lea ecx, var_1C
  loc_00E39E72: push eax
  loc_00E39E73: lea edx, var_18
  loc_00E39E76: push ecx
  loc_00E39E77: push edx
  loc_00E39E78: push 00000004h
  loc_00E39E7A: call [00401048h] ; __vbaFreeObjList
  loc_00E39E80: lea eax, var_64
  loc_00E39E83: lea ecx, var_54
  loc_00E39E86: push eax
  loc_00E39E87: lea edx, var_44
  loc_00E39E8A: push ecx
  loc_00E39E8B: lea eax, var_34
  loc_00E39E8E: push edx
  loc_00E39E8F: push eax
  loc_00E39E90: push 00000004h
  loc_00E39E92: call [00401038h] ; __vbaFreeVarList
  loc_00E39E98: add esp, 00000028h
  loc_00E39E9B: ret
  loc_00E39E9C: ret
  loc_00E39E9D: mov eax, Me
  loc_00E39EA0: push eax
  loc_00E39EA1: mov ecx, [eax]
  loc_00E39EA3: call [ecx+00000008h]
  loc_00E39EA6: mov eax, var_4
  loc_00E39EA9: mov ecx, var_14
  loc_00E39EAC: pop edi
  loc_00E39EAD: pop esi
  loc_00E39EAE: mov fs:[00000000h], ecx
  loc_00E39EB5: pop ebx
  loc_00E39EB6: mov esp, ebp
  loc_00E39EB8: pop ebp
  loc_00E39EB9: retn 0004h
End Sub

Private Sub batal_UnknownEvent_9 'E3A900
  loc_00E3A900: push ebp
  loc_00E3A901: mov ebp, esp
  loc_00E3A903: sub esp, 0000000Ch
  loc_00E3A906: push 00402836h ; __vbaExceptHandler
  loc_00E3A90B: mov eax, fs:[00000000h]
  loc_00E3A911: push eax
  loc_00E3A912: mov fs:[00000000h], esp
  loc_00E3A919: sub esp, 00000048h
  loc_00E3A91C: push ebx
  loc_00E3A91D: push esi
  loc_00E3A91E: push edi
  loc_00E3A91F: mov var_C, esp
  loc_00E3A922: mov var_8, 004027E0h
  loc_00E3A929: mov esi, Me
  loc_00E3A92C: mov eax, esi
  loc_00E3A92E: and eax, 00000001h
  loc_00E3A931: mov var_4, eax
  loc_00E3A934: and esi, FFFFFFFEh
  loc_00E3A937: push esi
  loc_00E3A938: mov Me, esi
  loc_00E3A93B: mov ecx, [esi]
  loc_00E3A93D: call [ecx+00000004h]
  loc_00E3A940: mov edx, [esi]
  loc_00E3A942: xor eax, eax
  loc_00E3A944: push esi
  loc_00E3A945: mov var_18, eax
  loc_00E3A948: mov var_1C, eax
  loc_00E3A94B: mov var_2C, eax
  loc_00E3A94E: call [edx+00000304h]
  loc_00E3A954: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E3A95A: push eax
  loc_00E3A95B: lea eax, var_18
  loc_00E3A95E: push eax
  loc_00E3A95F: call edi
  loc_00E3A961: mov ebx, eax
  loc_00E3A963: push 00000000h
  loc_00E3A965: push ebx
  loc_00E3A966: mov ecx, [ebx]
  loc_00E3A968: call [ecx+000000E4h]
  loc_00E3A96E: test eax, eax
  loc_00E3A970: fnclex
  loc_00E3A972: jge 00E3A986h
  loc_00E3A974: push 000000E4h
  loc_00E3A979: push 006E03D4h
  loc_00E3A97E: push ebx
  loc_00E3A97F: push eax
  loc_00E3A980: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A986: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E3A98C: lea ecx, var_18
  loc_00E3A98F: call ebx
  loc_00E3A991: mov edx, [esi]
  loc_00E3A993: push esi
  loc_00E3A994: call [edx+00000308h]
  loc_00E3A99A: push eax
  loc_00E3A99B: lea eax, var_18
  loc_00E3A99E: push eax
  loc_00E3A99F: call edi
  loc_00E3A9A1: mov ecx, [eax]
  loc_00E3A9A3: push 00000000h
  loc_00E3A9A5: push eax
  loc_00E3A9A6: mov var_50, eax
  loc_00E3A9A9: call [ecx+000000E4h]
  loc_00E3A9AF: test eax, eax
  loc_00E3A9B1: fnclex
  loc_00E3A9B3: jge 00E3A9CAh
  loc_00E3A9B5: mov edx, var_50
  loc_00E3A9B8: push 000000E4h
  loc_00E3A9BD: push 006E03D4h
  loc_00E3A9C2: push edx
  loc_00E3A9C3: push eax
  loc_00E3A9C4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3A9CA: lea ecx, var_18
  loc_00E3A9CD: call ebx
  loc_00E3A9CF: mov eax, [esi]
  loc_00E3A9D1: push esi
  loc_00E3A9D2: call [eax+00000300h]
  loc_00E3A9D8: lea ecx, var_18
  loc_00E3A9DB: push eax
  loc_00E3A9DC: push ecx
  loc_00E3A9DD: call edi
  loc_00E3A9DF: mov edx, [eax]
  loc_00E3A9E1: push 00000000h
  loc_00E3A9E3: push eax
  loc_00E3A9E4: mov var_50, eax
  loc_00E3A9E7: call [edx+000000E4h]
  loc_00E3A9ED: test eax, eax
  loc_00E3A9EF: fnclex
  loc_00E3A9F1: jge 00E3AA08h
  loc_00E3A9F3: mov ecx, var_50
  loc_00E3A9F6: push 000000E4h
  loc_00E3A9FB: push 006E03D4h
  loc_00E3AA00: push ecx
  loc_00E3AA01: push eax
  loc_00E3AA02: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AA08: lea ecx, var_18
  loc_00E3AA0B: call ebx
  loc_00E3AA0D: mov edx, [esi]
  loc_00E3AA0F: push esi
  loc_00E3AA10: call [edx+00000310h]
  loc_00E3AA16: push eax
  loc_00E3AA17: lea eax, var_18
  loc_00E3AA1A: push eax
  loc_00E3AA1B: call edi
  loc_00E3AA1D: mov ecx, [eax]
  loc_00E3AA1F: push 00000000h
  loc_00E3AA21: push eax
  loc_00E3AA22: mov var_50, eax
  loc_00E3AA25: call [ecx+0000009Ch]
  loc_00E3AA2B: test eax, eax
  loc_00E3AA2D: fnclex
  loc_00E3AA2F: jge 00E3AA46h
  loc_00E3AA31: mov edx, var_50
  loc_00E3AA34: push 0000009Ch
  loc_00E3AA39: push 006E03D4h
  loc_00E3AA3E: push edx
  loc_00E3AA3F: push eax
  loc_00E3AA40: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AA46: lea ecx, var_18
  loc_00E3AA49: call ebx
  loc_00E3AA4B: sub esp, 00000010h
  loc_00E3AA4E: mov ecx, 0000000Bh
  loc_00E3AA53: mov edx, esp
  loc_00E3AA55: xor eax, eax
  loc_00E3AA57: push 80010007h
  loc_00E3AA5C: push esi
  loc_00E3AA5D: mov [edx], ecx
  loc_00E3AA5F: mov ecx, var_38
  loc_00E3AA62: mov [edx+00000004h], ecx
  loc_00E3AA65: mov ecx, [esi]
  loc_00E3AA67: mov [edx+00000008h], eax
  loc_00E3AA6A: mov eax, var_30
  loc_00E3AA6D: mov [edx+0000000Ch], eax
  loc_00E3AA70: call [ecx+00000314h]
  loc_00E3AA76: lea edx, var_18
  loc_00E3AA79: push eax
  loc_00E3AA7A: push edx
  loc_00E3AA7B: call edi
  loc_00E3AA7D: push eax
  loc_00E3AA7E: call [00401238h] ; __vbaLateIdSt
  loc_00E3AA84: lea ecx, var_18
  loc_00E3AA87: call ebx
  loc_00E3AA89: mov eax, [esi]
  loc_00E3AA8B: push esi
  loc_00E3AA8C: call [eax+00000304h]
  loc_00E3AA92: lea ecx, var_18
  loc_00E3AA95: push eax
  loc_00E3AA96: push ecx
  loc_00E3AA97: call edi
  loc_00E3AA99: mov edx, [eax]
  loc_00E3AA9B: push FFFFFFFFh
  loc_00E3AA9D: push eax
  loc_00E3AA9E: mov var_50, eax
  loc_00E3AAA1: call [edx+0000009Ch]
  loc_00E3AAA7: test eax, eax
  loc_00E3AAA9: fnclex
  loc_00E3AAAB: jge 00E3AAC2h
  loc_00E3AAAD: mov ecx, var_50
  loc_00E3AAB0: push 0000009Ch
  loc_00E3AAB5: push 006E03D4h
  loc_00E3AABA: push ecx
  loc_00E3AABB: push eax
  loc_00E3AABC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AAC2: lea ecx, var_18
  loc_00E3AAC5: call ebx
  loc_00E3AAC7: mov edx, [esi]
  loc_00E3AAC9: push esi
  loc_00E3AACA: call [edx+00000308h]
  loc_00E3AAD0: push eax
  loc_00E3AAD1: lea eax, var_18
  loc_00E3AAD4: push eax
  loc_00E3AAD5: call edi
  loc_00E3AAD7: mov ecx, [eax]
  loc_00E3AAD9: push FFFFFFFFh
  loc_00E3AADB: push eax
  loc_00E3AADC: mov var_50, eax
  loc_00E3AADF: call [ecx+0000009Ch]
  loc_00E3AAE5: test eax, eax
  loc_00E3AAE7: fnclex
  loc_00E3AAE9: jge 00E3AB00h
  loc_00E3AAEB: mov edx, var_50
  loc_00E3AAEE: push 0000009Ch
  loc_00E3AAF3: push 006E03D4h
  loc_00E3AAF8: push edx
  loc_00E3AAF9: push eax
  loc_00E3AAFA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AB00: lea ecx, var_18
  loc_00E3AB03: call ebx
  loc_00E3AB05: mov eax, [esi]
  loc_00E3AB07: push esi
  loc_00E3AB08: call [eax+00000300h]
  loc_00E3AB0E: lea ecx, var_18
  loc_00E3AB11: push eax
  loc_00E3AB12: push ecx
  loc_00E3AB13: call edi
  loc_00E3AB15: mov edx, [eax]
  loc_00E3AB17: push FFFFFFFFh
  loc_00E3AB19: push eax
  loc_00E3AB1A: mov var_50, eax
  loc_00E3AB1D: call [edx+0000009Ch]
  loc_00E3AB23: test eax, eax
  loc_00E3AB25: fnclex
  loc_00E3AB27: jge 00E3AB3Eh
  loc_00E3AB29: mov ecx, var_50
  loc_00E3AB2C: push 0000009Ch
  loc_00E3AB31: push 006E03D4h
  loc_00E3AB36: push ecx
  loc_00E3AB37: push eax
  loc_00E3AB38: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AB3E: lea ecx, var_18
  loc_00E3AB41: call ebx
  loc_00E3AB43: mov edx, [esi]
  loc_00E3AB45: push esi
  loc_00E3AB46: call [edx+00000310h]
  loc_00E3AB4C: push eax
  loc_00E3AB4D: lea eax, var_18
  loc_00E3AB50: push eax
  loc_00E3AB51: call edi
  loc_00E3AB53: mov ecx, [eax]
  loc_00E3AB55: push FFFFFFFFh
  loc_00E3AB57: push eax
  loc_00E3AB58: mov var_50, eax
  loc_00E3AB5B: call [ecx+0000009Ch]
  loc_00E3AB61: test eax, eax
  loc_00E3AB63: fnclex
  loc_00E3AB65: jge 00E3AB7Ch
  loc_00E3AB67: mov edx, var_50
  loc_00E3AB6A: push 0000009Ch
  loc_00E3AB6F: push 006E03D4h
  loc_00E3AB74: push edx
  loc_00E3AB75: push eax
  loc_00E3AB76: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AB7C: lea ecx, var_18
  loc_00E3AB7F: call ebx
  loc_00E3AB81: mov eax, [esi]
  loc_00E3AB83: push esi
  loc_00E3AB84: call [eax+0000030Ch]
  loc_00E3AB8A: lea ecx, var_18
  loc_00E3AB8D: push eax
  loc_00E3AB8E: push ecx
  loc_00E3AB8F: call edi
  loc_00E3AB91: mov edx, [eax]
  loc_00E3AB93: push 006DCC80h
  loc_00E3AB98: push eax
  loc_00E3AB99: mov var_50, eax
  loc_00E3AB9C: call [edx+000000A4h]
  loc_00E3ABA2: test eax, eax
  loc_00E3ABA4: fnclex
  loc_00E3ABA6: jge 00E3ABBDh
  loc_00E3ABA8: mov ecx, var_50
  loc_00E3ABAB: push 000000A4h
  loc_00E3ABB0: push 006DCB70h
  loc_00E3ABB5: push ecx
  loc_00E3ABB6: push eax
  loc_00E3ABB7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3ABBD: lea ecx, var_18
  loc_00E3ABC0: call ebx
  loc_00E3ABC2: mov edx, [esi]
  loc_00E3ABC4: push 006E05D8h
  loc_00E3ABC9: push esi
  loc_00E3ABCA: call [edx+0000033Ch]
  loc_00E3ABD0: push eax
  loc_00E3ABD1: lea eax, var_18
  loc_00E3ABD4: push eax
  loc_00E3ABD5: call edi
  loc_00E3ABD7: push eax
  loc_00E3ABD8: call [00401224h] ; __vbaCastObj
  loc_00E3ABDE: lea ecx, var_2C
  loc_00E3ABE1: mov var_24, eax
  loc_00E3ABE4: push ecx
  loc_00E3ABE5: mov var_2C, 0000000Dh
  loc_00E3ABEC: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E3ABF2: mov ecx, [eax]
  loc_00E3ABF4: sub esp, 00000010h
  loc_00E3ABF7: mov edx, esp
  loc_00E3ABF9: push 00000000h
  loc_00E3ABFB: push 0000002Ah
  loc_00E3ABFD: mov [edx], ecx
  loc_00E3ABFF: mov ecx, [eax+00000004h]
  loc_00E3AC02: push esi
  loc_00E3AC03: mov [edx+00000004h], ecx
  loc_00E3AC06: mov ecx, [eax+00000008h]
  loc_00E3AC09: mov eax, [eax+0000000Ch]
  loc_00E3AC0C: mov [edx+00000008h], ecx
  loc_00E3AC0F: mov ecx, [esi]
  loc_00E3AC11: mov [edx+0000000Ch], eax
  loc_00E3AC14: call [ecx+00000340h]
  loc_00E3AC1A: lea edx, var_1C
  loc_00E3AC1D: push eax
  loc_00E3AC1E: push edx
  loc_00E3AC1F: call edi
  loc_00E3AC21: push eax
  loc_00E3AC22: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3AC28: lea eax, var_1C
  loc_00E3AC2B: lea ecx, var_18
  loc_00E3AC2E: push eax
  loc_00E3AC2F: push ecx
  loc_00E3AC30: push 00000002h
  loc_00E3AC32: call [00401048h] ; __vbaFreeObjList
  loc_00E3AC38: add esp, 00000028h
  loc_00E3AC3B: lea ecx, var_2C
  loc_00E3AC3E: call [00401024h] ; __vbaFreeVar
  loc_00E3AC44: sub esp, 00000010h
  loc_00E3AC47: mov ecx, 0000000Bh
  loc_00E3AC4C: mov edx, esp
  loc_00E3AC4E: xor eax, eax
  loc_00E3AC50: push 00000006h
  loc_00E3AC52: push esi
  loc_00E3AC53: mov [edx], ecx
  loc_00E3AC55: mov ecx, var_38
  loc_00E3AC58: mov [edx+00000004h], ecx
  loc_00E3AC5B: mov ecx, [esi]
  loc_00E3AC5D: mov [edx+00000008h], eax
  loc_00E3AC60: mov eax, var_30
  loc_00E3AC63: mov [edx+0000000Ch], eax
  loc_00E3AC66: call [ecx+00000340h]
  loc_00E3AC6C: lea edx, var_18
  loc_00E3AC6F: push eax
  loc_00E3AC70: push edx
  loc_00E3AC71: call edi
  loc_00E3AC73: push eax
  loc_00E3AC74: call [00401238h] ; __vbaLateIdSt
  loc_00E3AC7A: lea ecx, var_18
  loc_00E3AC7D: call ebx
  loc_00E3AC7F: xor eax, eax
  loc_00E3AC81: sub esp, 00000010h
  loc_00E3AC84: mov edx, esp
  loc_00E3AC86: mov ecx, 0000000Bh
  loc_00E3AC8B: push 8001000Eh
  loc_00E3AC90: mov [edx], ecx
  loc_00E3AC92: mov ecx, var_38
  loc_00E3AC95: mov [edx+00000004h], ecx
  loc_00E3AC98: mov ecx, [esi]
  loc_00E3AC9A: mov [edx+00000008h], eax
  loc_00E3AC9D: mov eax, var_30
  loc_00E3ACA0: mov [edx+0000000Ch], eax
  loc_00E3ACA3: push esi
  loc_00E3ACA4: call [ecx+00000340h]
  loc_00E3ACAA: lea edx, var_18
  loc_00E3ACAD: push eax
  loc_00E3ACAE: push edx
  loc_00E3ACAF: call edi
  loc_00E3ACB1: push eax
  loc_00E3ACB2: call [00401238h] ; __vbaLateIdSt
  loc_00E3ACB8: lea ecx, var_18
  loc_00E3ACBB: call ebx
  loc_00E3ACBD: mov eax, [esi]
  loc_00E3ACBF: push 00000000h
  loc_00E3ACC1: push FFFFFDDAh
  loc_00E3ACC6: push esi
  loc_00E3ACC7: call [eax+00000340h]
  loc_00E3ACCD: lea ecx, var_18
  loc_00E3ACD0: push eax
  loc_00E3ACD1: push ecx
  loc_00E3ACD2: call edi
  loc_00E3ACD4: push eax
  loc_00E3ACD5: call [00401030h] ; __vbaLateIdCall
  loc_00E3ACDB: add esp, 0000000Ch
  loc_00E3ACDE: lea ecx, var_18
  loc_00E3ACE1: call ebx
  loc_00E3ACE3: mov edx, [esi]
  loc_00E3ACE5: push esi
  loc_00E3ACE6: call [edx+0000030Ch]
  loc_00E3ACEC: push eax
  loc_00E3ACED: lea eax, var_18
  loc_00E3ACF0: push eax
  loc_00E3ACF1: call edi
  loc_00E3ACF3: mov esi, eax
  loc_00E3ACF5: push esi
  loc_00E3ACF6: mov ecx, [esi]
  loc_00E3ACF8: call [ecx+00000204h]
  loc_00E3ACFE: test eax, eax
  loc_00E3AD00: fnclex
  loc_00E3AD02: jge 00E3AD16h
  loc_00E3AD04: push 00000204h
  loc_00E3AD09: push 006DCB70h
  loc_00E3AD0E: push esi
  loc_00E3AD0F: push eax
  loc_00E3AD10: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3AD16: lea ecx, var_18
  loc_00E3AD19: call ebx
  loc_00E3AD1B: mov var_4, 00000000h
  loc_00E3AD22: push 00E3AD47h
  loc_00E3AD27: jmp 00E3AD46h
  loc_00E3AD29: lea edx, var_1C
  loc_00E3AD2C: lea eax, var_18
  loc_00E3AD2F: push edx
  loc_00E3AD30: push eax
  loc_00E3AD31: push 00000002h
  loc_00E3AD33: call [00401048h] ; __vbaFreeObjList
  loc_00E3AD39: add esp, 0000000Ch
  loc_00E3AD3C: lea ecx, var_2C
  loc_00E3AD3F: call [00401024h] ; __vbaFreeVar
  loc_00E3AD45: ret
  loc_00E3AD46: ret
  loc_00E3AD47: mov eax, Me
  loc_00E3AD4A: push eax
  loc_00E3AD4B: mov ecx, [eax]
  loc_00E3AD4D: call [ecx+00000008h]
  loc_00E3AD50: mov eax, var_4
  loc_00E3AD53: mov ecx, var_14
  loc_00E3AD56: pop edi
  loc_00E3AD57: pop esi
  loc_00E3AD58: mov fs:[00000000h], ecx
  loc_00E3AD5F: pop ebx
  loc_00E3AD60: mov esp, ebp
  loc_00E3AD62: pop ebp
  loc_00E3AD63: retn 0004h
End Sub
