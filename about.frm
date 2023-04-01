VERSION 5.00
Begin VB.Form about
  Caption = "Form1"
  BackColor = &H404040&
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  BorderStyle = 0 'None
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 13860
  ClientTop = 2190
  ClientWidth = 8790
  ClientHeight = 6270
  StartUpPosition = 2 'CenterScreen
  Begin VB.TextBox Text4
    BackColor = &H404040&
    ForeColor = &H80FF80&
    Left = 4950
    Top = 4410
    Width = 3165
    Height = 270
    Text = "SMK Swasta RK Bintang Timur "
    TabIndex = 11
    BorderStyle = 0 'None
  End
  Begin VB.TextBox Text3
    BackColor = &H404040&
    ForeColor = &H80FF80&
    Left = 4980
    Top = 3720
    Width = 3255
    Height = 330
    Text = "Buat Kreasi Logo, Menggambar"
    TabIndex = 9
    BorderStyle = 0 'None
  End
  Begin VB.TextBox Text2
    BackColor = &H404040&
    ForeColor = &H80FF80&
    Left = 4980
    Top = 3060
    Width = 2775
    Height = 330
    Text = "Kristen Protestan"
    TabIndex = 7
    BorderStyle = 0 'None
  End
  Begin VB.TextBox Text1
    BackColor = &H404040&
    ForeColor = &H80FF80&
    Left = 4980
    Top = 2340
    Width = 2775
    Height = 330
    Text = "XI - Ti B / RPL"
    TabIndex = 5
    BorderStyle = 0 'None
  End
  Begin VB.TextBox nama
    BackColor = &H404040&
    ForeColor = &H80FF80&
    Left = 4980
    Top = 1590
    Width = 2775
    Height = 330
    Text = "Andres Taruli Purba"
    TabIndex = 3
    BorderStyle = 0 'None
  End
  Begin VB.Shape Shape14
    Left = 5940
    Top = 5310
    Width = 435
    Height = 435
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &H808080&
    FillStyle = 0
  End
  Begin VB.Shape Shape13
    BorderColor = &HFFFFFF&
    Left = 4230
    Top = 5310
    Width = 435
    Height = 435
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &H808080&
    FillStyle = 0
  End
  Begin VB.Shape Shape11
    Left = 2580
    Top = 5310
    Width = 435
    Height = 435
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &H808080&
    FillStyle = 0
  End
  Begin VB.Shape Shape12
    BorderColor = &H80FF80&
    Left = 0
    Top = 6150
    Width = 15015
    Height = 60
    BorderStyle = 3 'Dot
    FillColor = &HFFFF&
  End
  Begin VB.Shape Shape10
    BorderColor = &H808080&
    Left = 3780
    Top = 4410
    Width = 1035
    Height = 375
    FillColor = &H4000&
  End
  Begin VB.Line Line5
    BorderColor = &H808080&
    X1 = 4800
    Y1 = 4770
    X2 = 7860
    Y2 = 4770
  End
  Begin VB.Label Label7
    Caption = "Sekolah"
    ForeColor = &HC0C0C0&
    Left = 3870
    Top = 4410
    Width = 855
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
  Begin VB.Shape Shape9
    BorderColor = &H808080&
    Left = 3810
    Top = 3720
    Width = 1035
    Height = 375
    FillColor = &H4000&
  End
  Begin VB.Line Line4
    BorderColor = &H808080&
    X1 = 4830
    Y1 = 4080
    X2 = 7890
    Y2 = 4080
  End
  Begin VB.Label Label6
    Caption = "Hobi"
    ForeColor = &HC0C0C0&
    Left = 3990
    Top = 3720
    Width = 795
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
  Begin VB.Shape Shape8
    BorderColor = &H808080&
    Left = 3810
    Top = 3060
    Width = 1035
    Height = 375
    FillColor = &H4000&
  End
  Begin VB.Line Line3
    BorderColor = &H808080&
    X1 = 4830
    Y1 = 3420
    X2 = 7890
    Y2 = 3420
  End
  Begin VB.Label Label5
    Caption = "Agama"
    ForeColor = &HC0C0C0&
    Left = 3990
    Top = 3060
    Width = 795
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
  Begin VB.Shape Shape7
    BorderColor = &H808080&
    Left = 3810
    Top = 2340
    Width = 1035
    Height = 375
    FillColor = &H4000&
  End
  Begin VB.Line Line2
    BorderColor = &H808080&
    X1 = 4830
    Y1 = 2700
    X2 = 7890
    Y2 = 2700
  End
  Begin VB.Label Label3
    Caption = "Kelas"
    ForeColor = &HC0C0C0&
    Left = 3990
    Top = 2340
    Width = 795
    Height = 315
    TabIndex = 4
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
  Begin VB.Line Line1
    BorderColor = &H808080&
    X1 = 4830
    Y1 = 1950
    X2 = 7890
    Y2 = 1950
  End
  Begin VB.Label Label2
    Caption = "B I O D A T A"
    ForeColor = &H808080&
    Left = 6960
    Top = 510
    Width = 1845
    Height = 315
    TabIndex = 2
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 15.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Label Label1
    Caption = "Nama"
    ForeColor = &HC0C0C0&
    Left = 3990
    Top = 1590
    Width = 795
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
  Begin VB.Shape Shape5
    Left = 630
    Top = 810
    Width = 2625
    Height = 285
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Label Label4
    Caption = "About Programmer"
    ForeColor = &H80FF80&
    Left = 3420
    Top = 60
    Width = 2445
    Height = 315
    TabIndex = 0
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
  Begin VB.Image Image1
    Picture = "about.frx":0000
    Left = 390
    Top = 1440
    Width = 3015
    Height = 3285
    Stretch = -1  'True
  End
  Begin VB.Image back
    Picture = "about.frx":00012198
    Left = 0
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Shape Shape2
    Left = 0
    Top = 480
    Width = 8865
    Height = 45
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Shape Shape1
    Left = 0
    Top = 0
    Width = 8865
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Shape Shape3
    Left = 240
    Top = 1290
    Width = 3315
    Height = 3585
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Shape Shape4
    Left = 450
    Top = 1080
    Width = 2955
    Height = 3645
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Shape Shape6
    BorderColor = &H808080&
    Left = 3810
    Top = 1590
    Width = 1035
    Height = 375
    FillColor = &H4000&
  End
End

Attribute VB_Name = "about"


Private Sub back_Click() 'E072A0
  loc_00E072A0: push ebp
  loc_00E072A1: mov ebp, esp
  loc_00E072A3: sub esp, 0000000Ch
  loc_00E072A6: push 00402836h ; __vbaExceptHandler
  loc_00E072AB: mov eax, fs:[00000000h]
  loc_00E072B1: push eax
  loc_00E072B2: mov fs:[00000000h], esp
  loc_00E072B9: sub esp, 00000034h
  loc_00E072BC: push ebx
  loc_00E072BD: push esi
  loc_00E072BE: push edi
  loc_00E072BF: mov var_C, esp
  loc_00E072C2: mov var_8, 00402040h
  loc_00E072C9: mov edi, Me
  loc_00E072CC: mov eax, edi
  loc_00E072CE: and eax, 00000001h
  loc_00E072D1: mov var_4, eax
  loc_00E072D4: and edi, FFFFFFFEh
  loc_00E072D7: push edi
  loc_00E072D8: mov Me, edi
  loc_00E072DB: mov ecx, [edi]
  loc_00E072DD: call [ecx+00000004h]
  loc_00E072E0: mov eax, [00E3D9CCh]
  loc_00E072E5: mov var_18, 00000000h
  loc_00E072EC: test eax, eax
  loc_00E072EE: jnz 00E07300h
  loc_00E072F0: push 00E3D9CCh
  loc_00E072F5: push 006DC960h
  loc_00E072FA: call [004011A0h] ; __vbaNew2
  loc_00E07300: mov esi, [00E3D9CCh]
  loc_00E07306: lea edx, var_18
  loc_00E07309: push edi
  loc_00E0730A: push edx
  loc_00E0730B: mov ebx, [esi]
  loc_00E0730D: call [004010B4h] ; __vbaObjSetAddref
  loc_00E07313: push eax
  loc_00E07314: push esi
  loc_00E07315: call [ebx+00000010h]
  loc_00E07318: test eax, eax
  loc_00E0731A: fnclex
  loc_00E0731C: jge 00E0732Dh
  loc_00E0731E: push 00000010h
  loc_00E07320: push 006DC950h
  loc_00E07325: push esi
  loc_00E07326: push eax
  loc_00E07327: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0732D: lea ecx, var_18
  loc_00E07330: call [00401254h] ; __vbaFreeObj
  loc_00E07336: mov eax, [00E3D024h]
  loc_00E0733B: test eax, eax
  loc_00E0733D: jnz 00E0734Fh
  loc_00E0733F: push 00E3D024h
  loc_00E07344: push 006CE120h
  loc_00E07349: call [004011A0h] ; __vbaNew2
  loc_00E0734F: sub esp, 00000010h
  loc_00E07352: mov ecx, 0000000Ah
  loc_00E07357: mov ebx, esp
  loc_00E07359: mov var_28, ecx
  loc_00E0735C: mov eax, 80020004h
  loc_00E07361: sub esp, 00000010h
  loc_00E07364: mov [ebx], ecx
  loc_00E07366: mov ecx, var_34
  loc_00E07369: mov edx, eax
  loc_00E0736B: mov esi, [00E3D024h]
  loc_00E07371: mov [ebx+00000004h], ecx
  loc_00E07374: mov ecx, esp
  loc_00E07376: mov edi, [esi]
  loc_00E07378: push esi
  loc_00E07379: mov [ebx+00000008h], eax
  loc_00E0737C: mov eax, var_2C
  loc_00E0737F: mov [ebx+0000000Ch], eax
  loc_00E07382: mov eax, var_28
  loc_00E07385: mov [ecx], eax
  loc_00E07387: mov eax, var_24
  loc_00E0738A: mov [ecx+00000004h], eax
  loc_00E0738D: mov [ecx+00000008h], edx
  loc_00E07390: mov edx, var_1C
  loc_00E07393: mov [ecx+0000000Ch], edx
  loc_00E07396: call [edi+000002B0h]
  loc_00E0739C: test eax, eax
  loc_00E0739E: fnclex
  loc_00E073A0: jge 00E073B4h
  loc_00E073A2: push 000002B0h
  loc_00E073A7: push 006DC970h
  loc_00E073AC: push esi
  loc_00E073AD: push eax
  loc_00E073AE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E073B4: mov var_4, 00000000h
  loc_00E073BB: push 00E073CDh
  loc_00E073C0: jmp 00E073CCh
  loc_00E073C2: lea ecx, var_18
  loc_00E073C5: call [00401254h] ; __vbaFreeObj
  loc_00E073CB: ret
  loc_00E073CC: ret
  loc_00E073CD: mov eax, Me
  loc_00E073D0: push eax
  loc_00E073D1: mov ecx, [eax]
  loc_00E073D3: call [ecx+00000008h]
  loc_00E073D6: mov eax, var_4
  loc_00E073D9: mov ecx, var_14
  loc_00E073DC: pop edi
  loc_00E073DD: pop esi
  loc_00E073DE: mov fs:[00000000h], ecx
  loc_00E073E5: pop ebx
  loc_00E073E6: mov esp, ebp
  loc_00E073E8: pop ebp
  loc_00E073E9: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E073F0
  loc_00E073F0: push ebp
  loc_00E073F1: mov ebp, esp
  loc_00E073F3: sub esp, 0000000Ch
  loc_00E073F6: push 00402836h ; __vbaExceptHandler
  loc_00E073FB: mov eax, fs:[00000000h]
  loc_00E07401: push eax
  loc_00E07402: mov fs:[00000000h], esp
  loc_00E07409: sub esp, 0000005Ch
  loc_00E0740C: push ebx
  loc_00E0740D: push esi
  loc_00E0740E: push edi
  loc_00E0740F: mov var_C, esp
  loc_00E07412: mov var_8, 00402050h
  loc_00E07419: mov esi, Me
  loc_00E0741C: mov eax, esi
  loc_00E0741E: and eax, 00000001h
  loc_00E07421: mov var_4, eax
  loc_00E07424: and esi, FFFFFFFEh
  loc_00E07427: push esi
  loc_00E07428: mov Me, esi
  loc_00E0742B: mov ecx, [esi]
  loc_00E0742D: call [ecx+00000004h]
  loc_00E07430: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07436: xor eax, eax
  loc_00E07438: mov var_18, eax
  loc_00E0743B: mov var_4C, eax
  loc_00E0743E: mov var_50, eax
  loc_00E07441: mov edx, [esi]
  loc_00E07443: lea eax, var_4C
  loc_00E07446: push eax
  loc_00E07447: push esi
  loc_00E07448: call [edx+00000070h]
  loc_00E0744B: test eax, eax
  loc_00E0744D: fnclex
  loc_00E0744F: jge 00E0745Ch
  loc_00E07451: push 00000070h
  loc_00E07453: push 006DFE8Ch
  loc_00E07458: push esi
  loc_00E07459: push eax
  loc_00E0745A: call ebx
  loc_00E0745C: fld real4 ptr var_4C
  loc_00E0745F: fadd st0, real4 ptr [00401298h]
  loc_00E07465: mov ecx, [esi]
  loc_00E07467: push ecx
  loc_00E07468: fnstsw ax
  loc_00E0746A: test al, 0Dh
  loc_00E0746C: jnz 00E07630h
  loc_00E07472: fstp real4 ptr [esp]
  loc_00E07475: push esi
  loc_00E07476: call [ecx+00000074h]
  loc_00E07479: test eax, eax
  loc_00E0747B: fnclex
  loc_00E0747D: jge 00E0748Ah
  loc_00E0747F: push 00000074h
  loc_00E07481: push 006DFE8Ch
  loc_00E07486: push esi
  loc_00E07487: push eax
  loc_00E07488: call ebx
  loc_00E0748A: mov edx, [esi]
  loc_00E0748C: lea eax, var_4C
  loc_00E0748F: push eax
  loc_00E07490: push esi
  loc_00E07491: call [edx+00000070h]
  loc_00E07494: test eax, eax
  loc_00E07496: fnclex
  loc_00E07498: jge 00E074A5h
  loc_00E0749A: push 00000070h
  loc_00E0749C: push 006DFE8Ch
  loc_00E074A1: push esi
  loc_00E074A2: push eax
  loc_00E074A3: call ebx
  loc_00E074A5: mov ecx, [esi]
  loc_00E074A7: lea edx, var_50
  loc_00E074AA: push edx
  loc_00E074AB: push esi
  loc_00E074AC: call [ecx+00000078h]
  loc_00E074AF: test eax, eax
  loc_00E074B1: fnclex
  loc_00E074B3: jge 00E074C0h
  loc_00E074B5: push 00000078h
  loc_00E074B7: push 006DFE8Ch
  loc_00E074BC: push esi
  loc_00E074BD: push eax
  loc_00E074BE: call ebx
  loc_00E074C0: sub esp, 00000010h
  loc_00E074C3: mov ecx, 0000000Ah
  loc_00E074C8: mov ebx, esp
  loc_00E074CA: mov eax, 80020004h
  loc_00E074CF: mov edx, eax
  loc_00E074D1: sub esp, 00000010h
  loc_00E074D4: mov [ebx], ecx
  loc_00E074D6: mov ecx, var_44
  loc_00E074D9: mov edi, [esi]
  loc_00E074DB: mov [ebx+00000004h], ecx
  loc_00E074DE: mov ecx, esp
  loc_00E074E0: sub esp, 00000010h
  loc_00E074E3: mov [ebx+00000008h], eax
  loc_00E074E6: mov eax, var_3C
  loc_00E074E9: mov [ebx+0000000Ch], eax
  loc_00E074EC: mov eax, 0000000Ah
  loc_00E074F1: mov [ecx], eax
  loc_00E074F3: mov eax, var_34
  loc_00E074F6: mov [ecx+00000004h], eax
  loc_00E074F9: mov eax, 00000004h
  loc_00E074FE: mov [ecx+00000008h], edx
  loc_00E07501: mov edx, var_2C
  loc_00E07504: mov [ecx+0000000Ch], edx
  loc_00E07507: mov edx, var_24
  loc_00E0750A: mov ecx, esp
  loc_00E0750C: mov [ecx], eax
  loc_00E0750E: mov eax, var_50
  loc_00E07511: mov [ecx+00000004h], edx
  loc_00E07514: mov [ecx+00000008h], eax
  loc_00E07517: mov eax, var_1C
  loc_00E0751A: mov [ecx+0000000Ch], eax
  loc_00E0751D: mov ecx, var_4C
  loc_00E07520: push ecx
  loc_00E07521: push esi
  loc_00E07522: call [edi+000002A4h]
  loc_00E07528: test eax, eax
  loc_00E0752A: fnclex
  loc_00E0752C: jge 00E07544h
  loc_00E0752E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07534: push 000002A4h
  loc_00E07539: push 006DFE8Ch
  loc_00E0753E: push esi
  loc_00E0753F: push eax
  loc_00E07540: call ebx
  loc_00E07542: jmp 00E0754Ah
  loc_00E07544: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0754A: call [004010BCh] ; rtcDoEvents
  loc_00E07550: mov edx, [esi]
  loc_00E07552: lea eax, var_4C
  loc_00E07555: push eax
  loc_00E07556: push esi
  loc_00E07557: call [edx+00000070h]
  loc_00E0755A: test eax, eax
  loc_00E0755C: fnclex
  loc_00E0755E: jge 00E0756Bh
  loc_00E07560: push 00000070h
  loc_00E07562: push 006DFE8Ch
  loc_00E07567: push esi
  loc_00E07568: push eax
  loc_00E07569: call ebx
  loc_00E0756B: mov eax, [00E3D9CCh]
  loc_00E07570: test eax, eax
  loc_00E07572: jnz 00E07584h
  loc_00E07574: push 00E3D9CCh
  loc_00E07579: push 006DC960h
  loc_00E0757E: call [004011A0h] ; __vbaNew2
  loc_00E07584: mov edi, [00E3D9CCh]
  loc_00E0758A: lea edx, var_18
  loc_00E0758D: push edx
  loc_00E0758E: push edi
  loc_00E0758F: mov ecx, [edi]
  loc_00E07591: call [ecx+00000018h]
  loc_00E07594: test eax, eax
  loc_00E07596: fnclex
  loc_00E07598: jge 00E075A5h
  loc_00E0759A: push 00000018h
  loc_00E0759C: push 006DC950h
  loc_00E075A1: push edi
  loc_00E075A2: push eax
  loc_00E075A3: call ebx
  loc_00E075A5: mov eax, var_18
  loc_00E075A8: lea edx, var_50
  loc_00E075AB: push edx
  loc_00E075AC: push eax
  loc_00E075AD: mov ecx, [eax]
  loc_00E075AF: mov edi, eax
  loc_00E075B1: call [ecx+00000098h]
  loc_00E075B7: test eax, eax
  loc_00E075B9: fnclex
  loc_00E075BB: jge 00E075CBh
  loc_00E075BD: push 00000098h
  loc_00E075C2: push 006DCAF0h
  loc_00E075C7: push edi
  loc_00E075C8: push eax
  loc_00E075C9: call ebx
  loc_00E075CB: fld real4 ptr var_4C
  loc_00E075CE: fcomp real4 ptr var_50
  loc_00E075D1: fnstsw ax
  loc_00E075D3: test ah, 41h
  loc_00E075D6: jz 00E075DFh
  loc_00E075D8: mov eax, 00000001h
  loc_00E075DD: jmp 00E075E1h
  loc_00E075DF: xor eax, eax
  loc_00E075E1: neg eax
  loc_00E075E3: lea ecx, var_18
  loc_00E075E6: mov edi, eax
  loc_00E075E8: call [00401254h] ; __vbaFreeObj
  loc_00E075EE: test di, di
  loc_00E075F1: jnz 00E07441h
  loc_00E075F7: mov var_4, 00000000h
  loc_00E075FE: fwait
  loc_00E075FF: push 00E07611h
  loc_00E07604: jmp 00E07610h
  loc_00E07606: lea ecx, var_18
  loc_00E07609: call [00401254h] ; __vbaFreeObj
  loc_00E0760F: ret
  loc_00E07610: ret
  loc_00E07611: mov eax, Me
  loc_00E07614: push eax
  loc_00E07615: mov ecx, [eax]
  loc_00E07617: call [ecx+00000008h]
  loc_00E0761A: mov eax, var_4
  loc_00E0761D: mov ecx, var_14
  loc_00E07620: pop edi
  loc_00E07621: pop esi
  loc_00E07622: mov fs:[00000000h], ecx
  loc_00E07629: pop ebx
  loc_00E0762A: mov esp, ebp
  loc_00E0762C: pop ebp
  loc_00E0762D: retn 0008h
End Sub
