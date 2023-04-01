VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form biayaf
  Caption = "Jurusan Yang Dipilih"
  BackColor = &H404040&
  ForeColor = &H0&
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  BorderStyle = 0 'None
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 0
  ClientTop = 0
  ClientWidth = 14685
  ClientHeight = 11145
  StartUpPosition = 2 'CenterScreen
  Begin VB.Timer Timer2
    Interval = 10
    Left = 1770
    Top = 10380
  End
  Begin VB.Frame Frame4
    BackColor = &H404040&
    Left = 10560
    Top = 3750
    Width = 4005
    Height = 5385
    TabIndex = 67
    BorderStyle = 0 'None
    Begin VB.Frame fcari
      Caption = "Cari Disini"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 390
      Top = 3960
      Width = 3315
      Height = 1305
      TabIndex = 79
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
        Left = 150
        Top = 510
        Width = 1725
        Height = 300
        TabIndex = 82
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
        Left = 150
        Top = 240
        Width = 1455
        Height = 300
        TabIndex = 81
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
        Left = 270
        Top = 840
        Width = 2835
        Height = 345
        TabIndex = 80
        BorderStyle = 0 'None
      End
      Begin Project1.jcbutton jcbutton2
        Left = 3060
        Top = 0
        Width = 225
        Height = 225
        TabIndex = 83
        OleObjectBlob = "biayaf.frx":0000
      End
    End
    Begin Project1.jcbutton totalkan
      Left = 180
      Top = 1470
      Width = 3645
      Height = 375
      TabIndex = 68
      OleObjectBlob = "biayaf.frx":0220
    End
    Begin VB.Frame fbayar
      Caption = "Frame5"
      BackColor = &H404040&
      Left = 30
      Top = 3480
      Width = 3945
      Height = 1755
      TabIndex = 76
      BorderStyle = 0 'None
      Begin Project1.jcbutton metodes
        Left = 1590
        Top = 660
        Width = 885
        Height = 855
        TabIndex = 77
        OleObjectBlob = "biayaf.frx":044C
      End
      Begin VB.Label Label40
        Caption = "*Pilih Metode Pembayaran Sekarang"
        ForeColor = &HFFFFFF&
        Left = 180
        Top = 90
        Width = 3615
        Height = 315
        TabIndex = 78
        BackStyle = 0 'Transparent
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
      Begin VB.Shape Shape34
        Left = 1500
        Top = 570
        Width = 1095
        Height = 1035
        Shape = 4
        BorderStyle = 0 'Transparent
        FillColor = &H808080&
        FillStyle = 0
      End
    End
    Begin VB.Shape shitung
      BorderColor = &HFF&
      Left = 120
      Top = 1410
      Width = 3795
      Height = 525
      BorderWidth = 2
      FillColor = &H8000&
    End
    Begin VB.Shape scari
      Left = -180
      Top = 3600
      Width = 6675
      Height = 1545
      BorderStyle = 0 'Transparent
      FillColor = &H8000&
      FillStyle = 0
    End
    Begin VB.Label Label39
      Caption = "*Terbilang"
      ForeColor = &H80FF80&
      Left = 1380
      Top = 1980
      Width = 1305
      Height = 315
      TabIndex = 74
      BackStyle = 0 'Transparent
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
    Begin VB.Line Line23
      BorderColor = &HE0E0E0&
      X1 = 150
      Y1 = 3270
      X2 = 3750
      Y2 = 3270
    End
    Begin VB.Shape Shape33
      BorderColor = &H80FF80&
      Left = 0
      Top = 3360
      Width = 15015
      Height = 60
      BorderStyle = 3 'Dot
      FillColor = &HFFFF&
    End
    Begin VB.Label tr
      Caption = "Total"
      ForeColor = &HFFFFFF&
      Left = 180
      Top = 2580
      Width = 3675
      Height = 675
      TabIndex = 73
      BackStyle = 0 'Transparent
      BeginProperty Font
        Name = "Trebuchet MS"
        Size = 9.75
        Charset = 0
        Weight = 400
        Underline = 0 'False
        Italic = -1 'True
        Strikethrough = 0 'False
      EndProperty
    End
    Begin VB.Label total
      Caption = "0"
      ForeColor = &HFFFFFF&
      Left = 2220
      Top = 840
      Width = 1815
      Height = 315
      TabIndex = 72
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
    Begin VB.Shape Shape32
      BorderColor = &H80FF80&
      Left = -60
      Top = 1380
      Width = 15015
      Height = 540
      BorderStyle = 3 'Dot
      FillColor = &HFFFF&
    End
    Begin VB.Label Label38
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 1800
      Top = 840
      Width = 405
      Height = 315
      TabIndex = 71
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
    Begin VB.Line Line20
      BorderColor = &HE0E0E0&
      X1 = 2220
      Y1 = 1170
      X2 = 3930
      Y2 = 1170
    End
    Begin VB.Label Label37
      Caption = "Total Biaya"
      ForeColor = &H80FF80&
      Left = 240
      Top = 870
      Width = 1485
      Height = 315
      TabIndex = 70
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
    Begin VB.Label Label15
      Caption = "Perhitungan Biaya"
      ForeColor = &HFFFFFF&
      Left = 180
      Top = 30
      Width = 2325
      Height = 315
      TabIndex = 69
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
    Begin VB.Shape Shape29
      Left = 2850
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &HC0C0FF&
      FillStyle = 0
    End
    Begin VB.Shape Shape28
      Left = 3660
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &HFF&
      FillStyle = 0
    End
    Begin VB.Shape Shape27
      Left = 3270
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &H8080FF&
      FillStyle = 0
    End
    Begin VB.Shape Shape26
      Left = 0
      Top = 420
      Width = 12255
      Height = 75
      BorderStyle = 0 'Transparent
      FillColor = &H80FF80&
      FillStyle = 0
    End
    Begin VB.Shape Shape25
      BorderColor = &H80FF80&
      Left = 0
      Top = 510
      Width = 15015
      Height = 60
      BorderStyle = 3 'Dot
      FillColor = &HFFFF&
    End
    Begin VB.Shape Shape30
      Left = 0
      Top = 0
      Width = 7125
      Height = 435
      BorderStyle = 0 'Transparent
      FillColor = &H4000&
      FillStyle = 0
    End
  End
  Begin VB.Frame Frame3
    BackColor = &H404040&
    Left = 5010
    Top = 3750
    Width = 5505
    Height = 5385
    TabIndex = 17
    BorderStyle = 0 'None
    Begin VB.Line Line19
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 5220
      X2 = 5250
      Y2 = 5220
    End
    Begin VB.Label pn
      Caption = "Papan Nama"
      ForeColor = &HFFFFFF&
      Left = 3570
      Top = 4860
      Width = 1815
      Height = 315
      TabIndex = 58
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
    Begin VB.Label Label35
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 4860
      Width = 405
      Height = 315
      TabIndex = 57
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
    Begin VB.Label Label34
      Caption = "Papan Nama"
      ForeColor = &H80FF80&
      Left = 180
      Top = 4830
      Width = 2745
      Height = 315
      TabIndex = 56
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
    Begin VB.Line Line18
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 4800
      X2 = 5250
      Y2 = 4800
    End
    Begin VB.Label atribut
      Caption = "Atribut"
      ForeColor = &HFFFFFF&
      Left = 3570
      Top = 4440
      Width = 1815
      Height = 315
      TabIndex = 55
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
    Begin VB.Label Label33
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 4440
      Width = 405
      Height = 315
      TabIndex = 54
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
    Begin VB.Label Label32
      Caption = "Atribut"
      ForeColor = &H80FF80&
      Left = 180
      Top = 4410
      Width = 2745
      Height = 315
      TabIndex = 53
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
    Begin VB.Line Line17
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 4350
      X2 = 5250
      Y2 = 4350
    End
    Begin VB.Label dasi
      Caption = "Dasi Sekolah"
      ForeColor = &HFFFFFF&
      Left = 3570
      Top = 3990
      Width = 1815
      Height = 315
      TabIndex = 52
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
    Begin VB.Label Label29
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 3990
      Width = 405
      Height = 315
      TabIndex = 51
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
    Begin VB.Label Label26
      Caption = "Dasi Sekolah"
      ForeColor = &H80FF80&
      Left = 180
      Top = 3960
      Width = 2745
      Height = 315
      TabIndex = 50
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
    Begin VB.Line Line16
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 3870
      X2 = 5250
      Y2 = 3870
    End
    Begin VB.Label topi
      Caption = "Topi"
      ForeColor = &HFFFFFF&
      Left = 3540
      Top = 3540
      Width = 1815
      Height = 315
      TabIndex = 49
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
    Begin VB.Label Label23
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 3540
      Width = 405
      Height = 315
      TabIndex = 48
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
    Begin VB.Label Label20
      Caption = "Topi"
      ForeColor = &H80FF80&
      Left = 180
      Top = 3570
      Width = 2745
      Height = 315
      TabIndex = 47
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
    Begin VB.Shape Shape31
      BorderColor = &H80FF80&
      Left = 0
      Top = 3360
      Width = 15015
      Height = 60
      BorderStyle = 3 'Dot
      FillColor = &HFFFF&
    End
    Begin VB.Line Line14
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 3120
      X2 = 5250
      Y2 = 3120
    End
    Begin VB.Label pramuka
      Caption = "Baju Pramuka"
      ForeColor = &HFFFFFF&
      Left = 3540
      Top = 2790
      Width = 1815
      Height = 315
      TabIndex = 43
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
    Begin VB.Label Label28
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 2790
      Width = 405
      Height = 315
      TabIndex = 42
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
    Begin VB.Label Label27
      Caption = "Pramuka"
      ForeColor = &H80FF80&
      Left = 180
      Top = 2820
      Width = 2745
      Height = 315
      TabIndex = 41
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
    Begin VB.Line Line13
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 2640
      X2 = 5250
      Y2 = 2640
    End
    Begin VB.Label olahraga
      Caption = "Baju Olahraga"
      ForeColor = &HFFFFFF&
      Left = 3540
      Top = 2310
      Width = 1815
      Height = 315
      TabIndex = 40
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
    Begin VB.Label Label25
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 2310
      Width = 405
      Height = 315
      TabIndex = 39
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
    Begin VB.Label Label24
      Caption = "Olahraga"
      ForeColor = &H80FF80&
      Left = 180
      Top = 2340
      Width = 2745
      Height = 315
      TabIndex = 38
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
    Begin VB.Line Line12
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 2160
      X2 = 5250
      Y2 = 2160
    End
    Begin VB.Label batik
      Caption = "Baju Batik"
      ForeColor = &HFFFFFF&
      Left = 3540
      Top = 1830
      Width = 1815
      Height = 315
      TabIndex = 37
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
    Begin VB.Label Label22
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 1830
      Width = 405
      Height = 315
      TabIndex = 36
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
    Begin VB.Label Label21
      Caption = "Batik"
      ForeColor = &H80FF80&
      Left = 180
      Top = 1860
      Width = 2745
      Height = 315
      TabIndex = 35
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
    Begin VB.Line Line11
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 1650
      X2 = 5250
      Y2 = 1650
    End
    Begin VB.Label bajuprak
      Caption = "Baju Praktek"
      ForeColor = &HFFFFFF&
      Left = 3540
      Top = 1320
      Width = 1815
      Height = 315
      TabIndex = 34
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
    Begin VB.Label Label19
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 1320
      Width = 405
      Height = 315
      TabIndex = 33
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
    Begin VB.Label Label18
      Caption = "Praktek"
      ForeColor = &H80FF80&
      Left = 180
      Top = 1350
      Width = 2745
      Height = 315
      TabIndex = 32
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
    Begin VB.Line Line10
      BorderColor = &HE0E0E0&
      X1 = 3540
      Y1 = 1140
      X2 = 5250
      Y2 = 1140
    End
    Begin VB.Label abu
      Caption = "Putih Abu-abu"
      ForeColor = &HFFFFFF&
      Left = 3540
      Top = 810
      Width = 1815
      Height = 315
      TabIndex = 31
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
    Begin VB.Label Label17
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 810
      Width = 405
      Height = 315
      TabIndex = 30
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
    Begin VB.Label Label16
      Caption = "Putih Abu - abu"
      ForeColor = &H80FF80&
      Left = 180
      Top = 840
      Width = 2745
      Height = 315
      TabIndex = 29
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
    Begin VB.Shape Shape24
      BorderColor = &H80FF80&
      Left = 0
      Top = 510
      Width = 15015
      Height = 60
      BorderStyle = 3 'Dot
      FillColor = &HFFFF&
    End
    Begin VB.Shape Shape22
      Left = 0
      Top = 420
      Width = 12255
      Height = 75
      BorderStyle = 0 'Transparent
      FillColor = &H80FF80&
      FillStyle = 0
    End
    Begin VB.Shape Shape21
      Left = 4860
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &H80FFFF&
      FillStyle = 0
    End
    Begin VB.Shape Shape20
      Left = 5250
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &H80FF&
      FillStyle = 0
    End
    Begin VB.Shape Shape19
      Left = 4440
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &HC0FFFF&
      FillStyle = 0
    End
    Begin VB.Label Label2
      Caption = "Rincian Biaya Seragam Lengkap"
      ForeColor = &H80FF80&
      Left = 180
      Top = 30
      Width = 4125
      Height = 315
      TabIndex = 18
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
    Begin VB.Shape Shape13
      Left = 0
      Top = 0
      Width = 7125
      Height = 435
      BorderStyle = 0 'Transparent
      FillColor = &H4000&
      FillStyle = 0
    End
  End
  Begin Project1.jcbutton refreshh
    Left = 8640
    Top = 1890
    Width = 1125
    Height = 885
    TabIndex = 12
    OleObjectBlob = "biayaf.frx":1820
  End
  Begin VB.Frame Frame1
    BackColor = &H404040&
    Left = 120
    Top = 3750
    Width = 4845
    Height = 5385
    TabIndex = 9
    BorderStyle = 0 'None
    Begin VB.Frame farmasi
      Caption = "Jurusan Yang Dipilih"
      BackColor = &H404040&
      Left = 0
      Top = 3240
      Width = 4845
      Height = 1605
      TabIndex = 59
      BorderStyle = 0 'None
      Begin VB.Line Line25
        BorderColor = &HE0E0E0&
        X1 = 3150
        Y1 = 1140
        X2 = 4860
        Y2 = 1140
      End
      Begin VB.Label jas
        Caption = "Jas"
        ForeColor = &HFFFFFF&
        Left = 3150
        Top = 810
        Width = 2445
        Height = 315
        TabIndex = 66
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
      Begin VB.Label Label36
        Caption = "Rp"
        ForeColor = &HFFFFFF&
        Left = 2670
        Top = 870
        Width = 405
        Height = 315
        TabIndex = 65
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
      Begin VB.Label Label14
        Caption = "Uang Jas Praktek Farmasi"
        ForeColor = &H80FF80&
        Left = 210
        Top = 750
        Width = 2085
        Height = 705
        TabIndex = 64
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
      Begin VB.Label praktekfarmasi
        Caption = "Praktek"
        ForeColor = &HFFFFFF&
        Left = 3150
        Top = 120
        Width = 2445
        Height = 315
        TabIndex = 62
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
      Begin VB.Line Line8
        BorderColor = &HE0E0E0&
        X1 = 3090
        Y1 = 420
        X2 = 4800
        Y2 = 420
      End
      Begin VB.Label Label12
        Caption = "Uang Praktek Farmasi"
        ForeColor = &H80FF80&
        Left = 210
        Top = 0
        Width = 2085
        Height = 705
        TabIndex = 61
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
      Begin VB.Label Label13
        Caption = "Rp"
        ForeColor = &HFFFFFF&
        Left = 2670
        Top = 120
        Width = 405
        Height = 315
        TabIndex = 60
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
    Begin VB.Line Line15
      BorderColor = &HE0E0E0&
      X1 = 3090
      Y1 = 2940
      X2 = 4800
      Y2 = 2940
    End
    Begin VB.Label dasar
      Caption = "Dasar"
      ForeColor = &HFFFFFF&
      Left = 3180
      Top = 2640
      Width = 2445
      Height = 315
      TabIndex = 46
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
    Begin VB.Label Label31
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 2670
      Top = 2610
      Width = 405
      Height = 315
      TabIndex = 45
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
    Begin VB.Label Label30
      Caption = "Uji Kompetensi Dasar"
      ForeColor = &H80FF80&
      Left = 210
      Top = 2490
      Width = 2085
      Height = 705
      TabIndex = 44
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
    Begin VB.Line Line7
      BorderColor = &HE0E0E0&
      X1 = 3090
      Y1 = 2370
      X2 = 4800
      Y2 = 2370
    End
    Begin VB.Line Line6
      BorderColor = &HE0E0E0&
      X1 = 3090
      Y1 = 1800
      X2 = 4800
      Y2 = 1800
    End
    Begin VB.Line Line5
      BorderColor = &HE0E0E0&
      X1 = 3090
      Y1 = 1140
      X2 = 4800
      Y2 = 1140
    End
    Begin VB.Label tahun
      Caption = "Tahunan"
      ForeColor = &HFFFFFF&
      Left = 3180
      Top = 2070
      Width = 2445
      Height = 315
      TabIndex = 27
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
    Begin VB.Label spp
      Caption = "SPP"
      ForeColor = &HFFFFFF&
      Left = 3180
      Top = 1470
      Width = 2445
      Height = 315
      TabIndex = 26
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
    Begin VB.Label pem
      Caption = "Pembangunan"
      ForeColor = &HFFFFFF&
      Left = 3120
      Top = 780
      Width = 2445
      Height = 315
      TabIndex = 25
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
    Begin VB.Label Label11
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 2670
      Top = 2040
      Width = 405
      Height = 315
      TabIndex = 24
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
    Begin VB.Label Label10
      Caption = "Uang Tahunan"
      ForeColor = &H80FF80&
      Left = 210
      Top = 2040
      Width = 2085
      Height = 315
      TabIndex = 23
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
    Begin VB.Label Label9
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 2670
      Top = 1440
      Width = 405
      Height = 315
      TabIndex = 22
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
    Begin VB.Label Label8
      Caption = "Rp"
      ForeColor = &HFFFFFF&
      Left = 2700
      Top = 840
      Width = 405
      Height = 315
      TabIndex = 21
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
    Begin VB.Label Label7
      Caption = "Uang SPP"
      ForeColor = &H80FF80&
      Left = 210
      Top = 1440
      Width = 2085
      Height = 315
      TabIndex = 20
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
    Begin VB.Shape Shape23
      BorderColor = &H80FF80&
      Left = 0
      Top = 510
      Width = 15015
      Height = 60
      BorderStyle = 3 'Dot
      FillColor = &HFFFF&
    End
    Begin VB.Label Label6
      Caption = "Uang Pembangunan"
      ForeColor = &H80FF80&
      Left = 210
      Top = 840
      Width = 2745
      Height = 315
      TabIndex = 19
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
      Left = 0
      Top = 420
      Width = 12255
      Height = 75
      BorderStyle = 0 'Transparent
      FillColor = &H80FF80&
      FillStyle = 0
    End
    Begin VB.Shape Shape17
      Left = 3960
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &H80FF80&
      FillStyle = 0
    End
    Begin VB.Shape Shape16
      Left = 4350
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &H8000&
      FillStyle = 0
    End
    Begin VB.Shape target
      Left = 3540
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &HC0FFC0&
      FillStyle = 0
    End
    Begin VB.Label Label1
      Caption = "Rincian Biaya Sekolah"
      ForeColor = &H80FF80&
      Left = 210
      Top = 30
      Width = 2745
      Height = 315
      TabIndex = 14
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
      Left = 0
      Top = 0
      Width = 7155
      Height = 435
      BorderStyle = 0 'Transparent
      FillColor = &H4000&
      FillStyle = 0
    End
  End
  Begin Project1.jcbutton call
    Left = 4020
    Top = 2550
    Width = 2625
    Height = 435
    TabIndex = 8
    OleObjectBlob = "biayaf.frx":264C
  End
  Begin VB.Timer Timer1
    Interval = 100
    Left = 1350
    Top = 10380
  End
  Begin MSAdodcLib.Adodc adopen
    Left = 150
    Top = 10380
    Width = 1200
    Height = 345
    Visible = 0   'False
    OleObjectBlob = "biayaf.frx":2CD0
  End
  Begin MSDataGridLib.DataGrid dgpen
    Left = 90
    Top = 9300
    Width = 14415
    Height = 1725
    TabIndex = 2
    OleObjectBlob = "biayaf.frx":300C
  End
  Begin Project1.jcbutton simpan
    Left = 9870
    Top = 1890
    Width = 1095
    Height = 885
    TabIndex = 10
    OleObjectBlob = "biayaf.frx":31E7
  End
  Begin Project1.jcbutton hapus
    Left = 11040
    Top = 1890
    Width = 1125
    Height = 885
    TabIndex = 11
    OleObjectBlob = "biayaf.frx":3C43
  End
  Begin VB.Frame Frame2
    Caption = "Frame1"
    BackColor = &H4000&
    Left = 660
    Top = 3150
    Width = 8745
    Height = 345
    TabIndex = 15
    BorderStyle = 0 'None
    Begin VB.Label Label54
      Caption = "Silahkan Input Data Biaya Calon Peseta Didik Secara lengkap."
      ForeColor = &H80FF80&
      Left = 750
      Top = 0
      Width = 7155
      Height = 315
      TabIndex = 16
      BackStyle = 0 'Transparent
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
  End
  Begin Project1.jcbutton cari
    Left = 12240
    Top = 1890
    Width = 1095
    Height = 885
    TabIndex = 75
    OleObjectBlob = "biayaf.frx":4B1F
  End
  Begin Project1.jcbutton metodeff
    Left = 13410
    Top = 1890
    Width = 1035
    Height = 885
    TabIndex = 84
    OleObjectBlob = "biayaf.frx":5BBB
  End
  Begin VB.Shape scall
    BorderColor = &HFF&
    Left = 3960
    Top = 2490
    Width = 2775
    Height = 555
    BorderWidth = 2
  End
  Begin VB.Shape shpp
    Left = 0
    Top = 510
    Width = 15
    Height = 75
    Shape = 4
    BorderStyle = 0 'Transparent
    FillColor = &HFF0000&
    FillStyle = 0
  End
  Begin VB.Shape ssimpan
    BorderColor = &HFF&
    Left = 9810
    Top = 1830
    Width = 1215
    Height = 1035
    BorderWidth = 2
    FillColor = &H8000&
  End
  Begin VB.Shape Shape35
    Left = 0
    Top = 9180
    Width = 14775
    Height = 2205
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Line Line24
    BorderColor = &H80FF80&
    X1 = 4020
    Y1 = 2400
    X2 = 6450
    Y2 = 2400
  End
  Begin VB.Label agama
    Caption = "Agama"
    ForeColor = &HFFFFFF&
    Left = 4020
    Top = 2070
    Width = 2445
    Height = 315
    TabIndex = 63
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Label jenisk
    Caption = "Jenis Kelamin"
    ForeColor = &HFFFFFF&
    Left = 4020
    Top = 1620
    Width = 2445
    Height = 315
    TabIndex = 28
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Line Line9
    BorderColor = &H80FF80&
    X1 = 4020
    Y1 = 1950
    X2 = 6450
    Y2 = 1950
  End
  Begin VB.Shape Shape12
    Left = 0
    Top = 3150
    Width = 14685
    Height = 375
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Shape Shape11
    Left = 0
    Top = 3510
    Width = 14685
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Shape Shape10
    Left = 0
    Top = 3090
    Width = 14715
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Line Line4
    BorderColor = &HE0E0E0&
    X1 = 8640
    Y1 = 2880
    X2 = 14400
    Y2 = 2880
  End
  Begin VB.Shape Shape8
    Left = 8400
    Top = 1470
    Width = 6285
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Label Label5
    Caption = "Tombol Navigasi"
    ForeColor = &HFFFFFF&
    Left = 8580
    Top = 1560
    Width = 1965
    Height = 315
    TabIndex = 13
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Shape Shape7
    Left = 8400
    Top = 1470
    Width = 6285
    Height = 1545
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Shape Shape6
    Left = -420
    Top = 1470
    Width = 7095
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Line Line3
    BorderColor = &H80FF80&
    X1 = 1230
    Y1 = 2760
    X2 = 3660
    Y2 = 2760
  End
  Begin VB.Line Line2
    BorderColor = &H80FF80&
    X1 = 1230
    Y1 = 2400
    X2 = 3660
    Y2 = 2400
  End
  Begin VB.Line Line1
    BorderColor = &H80FF80&
    X1 = 1230
    Y1 = 1980
    X2 = 3660
    Y2 = 1980
  End
  Begin VB.Label jurus
    Caption = "Jurusan Yang Dipilih"
    ForeColor = &HFFFFFF&
    Left = 1230
    Top = 2460
    Width = 2445
    Height = 315
    TabIndex = 7
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Label namas
    Caption = "Nama Siswa"
    ForeColor = &HFFFFFF&
    Left = 1230
    Top = 2070
    Width = 2445
    Height = 315
    TabIndex = 6
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Label nodafs
    Caption = "No. Daftar"
    ForeColor = &HFFFFFF&
    Left = 1230
    Top = 1650
    Width = 2445
    Height = 315
    TabIndex = 5
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Shape Shape5
    BorderColor = &HFFFFFF&
    Left = 210
    Top = 1740
    Width = 795
    Height = 885
    BorderWidth = 2
  End
  Begin VB.Image Image1
    Picture = "biayaf.frx":6F8F
    Left = 180
    Top = 1770
    Width = 915
    Height = 855
    Stretch = -1  'True
  End
  Begin VB.Shape Shape4
    Left = 0
    Top = 1470
    Width = 6675
    Height = 1545
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Label lbljam
    Caption = "Labeljam"
    ForeColor = &H80FF80&
    Left = 13860
    Top = 60
    Width = 735
    Height = 375
    TabIndex = 4
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Digital-7"
      Size = 12
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Shape Shape15
    Left = 13710
    Top = 60
    Width = 45
    Height = 285
    BorderStyle = 0 'Transparent
    FillColor = &HFF00&
    FillStyle = 0
  End
  Begin VB.Label lbltgl
    Caption = "Labeltanggal"
    ForeColor = &H80FF80&
    Left = 12540
    Top = 60
    Width = 1185
    Height = 375
    TabIndex = 3
    BackStyle = 0 'Transparent
    BeginProperty Font
      Name = "Digital-7"
      Size = 12
      Charset = 0
      Weight = 400
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
  End
  Begin VB.Label Label3
    Caption = "SMK Swasta RK Bintang Timur Pematang Siantar"
    ForeColor = &H4000&
    Left = 1290
    Top = 600
    Width = 3465
    Height = 675
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
  Begin VB.Image Image2
    Picture = "biayaf.frx":7B46
    Left = 12900
    Top = 540
    Width = 705
    Height = 675
    Stretch = -1  'True
  End
  Begin VB.Image Image3
    Picture = "biayaf.frx":D682
    Left = 210
    Top = 570
    Width = 915
    Height = 855
    Stretch = -1  'True
  End
  Begin VB.Image Image4
    Picture = "biayaf.frx":00040964
    Left = 13710
    Top = 540
    Width = 1305
    Height = 735
    Stretch = -1  'True
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
    Caption = "Administration System"
    ForeColor = &H80FF80&
    Left = 6210
    Top = 60
    Width = 2745
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
  Begin VB.Image back
    Picture = "biayaf.frx":00047E3A
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
  Begin VB.Shape Shape2
    Left = 0
    Top = 450
    Width = 14715
    Height = 1035
    BorderStyle = 0 'Transparent
    FillColor = &HFFFFFF&
    FillStyle = 0
  End
  Begin VB.Image Image5
End

Attribute VB_Name = "biayaf"


Private Sub txtcari_KeyPress(KeyAscii As Integer) 'E1E580
  loc_00E1E580: push ebp
  loc_00E1E581: mov ebp, esp
  loc_00E1E583: sub esp, 0000000Ch
  loc_00E1E586: push 00402836h ; __vbaExceptHandler
  loc_00E1E58B: mov eax, fs:[00000000h]
  loc_00E1E591: push eax
  loc_00E1E592: mov fs:[00000000h], esp
  loc_00E1E599: sub esp, 000000ACh
  loc_00E1E59F: push ebx
  loc_00E1E5A0: push esi
  loc_00E1E5A1: push edi
  loc_00E1E5A2: mov var_C, esp
  loc_00E1E5A5: mov var_8, 00402410h
  loc_00E1E5AC: mov esi, Me
  loc_00E1E5AF: mov eax, esi
  loc_00E1E5B1: and eax, 00000001h
  loc_00E1E5B4: mov var_4, eax
  loc_00E1E5B7: and esi, FFFFFFFEh
  loc_00E1E5BA: push esi
  loc_00E1E5BB: mov Me, esi
  loc_00E1E5BE: mov ecx, [esi]
  loc_00E1E5C0: call [ecx+00000004h]
  loc_00E1E5C3: mov edx, [esi]
  loc_00E1E5C5: xor ebx, ebx
  loc_00E1E5C7: push esi
  loc_00E1E5C8: mov var_18, ebx
  loc_00E1E5CB: mov var_1C, ebx
  loc_00E1E5CE: mov var_20, ebx
  loc_00E1E5D1: mov var_24, ebx
  loc_00E1E5D4: mov var_34, ebx
  loc_00E1E5D7: mov var_44, ebx
  loc_00E1E5DA: mov var_54, ebx
  loc_00E1E5DD: mov var_64, ebx
  loc_00E1E5E0: mov var_74, ebx
  loc_00E1E5E3: mov var_84, ebx
  loc_00E1E5E9: mov var_A8, ebx
  loc_00E1E5EF: call [edx+00000308h]
  loc_00E1E5F5: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E1E5FB: push eax
  loc_00E1E5FC: lea eax, var_24
  loc_00E1E5FF: push eax
  loc_00E1E600: call edi
  loc_00E1E602: mov ecx, [eax]
  loc_00E1E604: lea edx, var_A8
  loc_00E1E60A: push edx
  loc_00E1E60B: push eax
  loc_00E1E60C: mov var_AC, eax
  loc_00E1E612: call [ecx+000000E0h]
  loc_00E1E618: cmp eax, ebx
  loc_00E1E61A: fnclex
  loc_00E1E61C: jge 00E1E636h
  loc_00E1E61E: mov ecx, var_AC
  loc_00E1E624: push 000000E0h
  loc_00E1E629: push 006E03D4h
  loc_00E1E62E: push ecx
  loc_00E1E62F: push eax
  loc_00E1E630: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E636: mov edx, var_A8
  loc_00E1E63C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E1E642: lea ecx, var_24
  loc_00E1E645: mov var_B4, edx
  loc_00E1E64B: call ebx
  loc_00E1E64D: cmp var_B4, 0000h
  loc_00E1E655: jz 00E1E6A2h
  loc_00E1E657: mov eax, [esi]
  loc_00E1E659: push esi
  loc_00E1E65A: call [eax+00000310h]
  loc_00E1E660: lea ecx, var_24
  loc_00E1E663: push eax
  loc_00E1E664: push ecx
  loc_00E1E665: call edi
  loc_00E1E667: mov edx, [eax]
  loc_00E1E669: lea ecx, var_18
  loc_00E1E66C: push ecx
  loc_00E1E66D: push eax
  loc_00E1E66E: mov var_AC, eax
  loc_00E1E674: call [edx+000000A0h]
  loc_00E1E67A: test eax, eax
  loc_00E1E67C: fnclex
  loc_00E1E67E: jge 00E1E698h
  loc_00E1E680: mov edx, var_AC
  loc_00E1E686: push 000000A0h
  loc_00E1E68B: push 006DCB70h
  loc_00E1E690: push edx
  loc_00E1E691: push eax
  loc_00E1E692: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E698: push 006E1428h ; "select * from biaya where nama like '" & Chr(37)
  loc_00E1E69D: jmp 00E1E74Bh
  loc_00E1E6A2: mov edx, [esi]
  loc_00E1E6A4: push esi
  loc_00E1E6A5: call [edx+0000030Ch]
  loc_00E1E6AB: push eax
  loc_00E1E6AC: lea eax, var_24
  loc_00E1E6AF: push eax
  loc_00E1E6B0: call edi
  loc_00E1E6B2: mov ecx, [eax]
  loc_00E1E6B4: lea edx, var_A8
  loc_00E1E6BA: push edx
  loc_00E1E6BB: push eax
  loc_00E1E6BC: mov var_AC, eax
  loc_00E1E6C2: call [ecx+000000E0h]
  loc_00E1E6C8: test eax, eax
  loc_00E1E6CA: fnclex
  loc_00E1E6CC: jge 00E1E6E6h
  loc_00E1E6CE: mov ecx, var_AC
  loc_00E1E6D4: push 000000E0h
  loc_00E1E6D9: push 006E03D4h
  loc_00E1E6DE: push ecx
  loc_00E1E6DF: push eax
  loc_00E1E6E0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E6E6: mov edx, var_A8
  loc_00E1E6EC: lea ecx, var_24
  loc_00E1E6EF: mov var_B4, edx
  loc_00E1E6F5: call ebx
  loc_00E1E6F7: cmp var_B4, 0000h
  loc_00E1E6FF: jz 00E1E808h
  loc_00E1E705: mov eax, [esi]
  loc_00E1E707: push esi
  loc_00E1E708: call [eax+00000310h]
  loc_00E1E70E: lea ecx, var_24
  loc_00E1E711: push eax
  loc_00E1E712: push ecx
  loc_00E1E713: call edi
  loc_00E1E715: mov edx, [eax]
  loc_00E1E717: lea ecx, var_18
  loc_00E1E71A: push ecx
  loc_00E1E71B: push eax
  loc_00E1E71C: mov var_AC, eax
  loc_00E1E722: call [edx+000000A0h]
  loc_00E1E728: test eax, eax
  loc_00E1E72A: fnclex
  loc_00E1E72C: jge 00E1E746h
  loc_00E1E72E: mov edx, var_AC
  loc_00E1E734: push 000000A0h
  loc_00E1E739: push 006DCB70h
  loc_00E1E73E: push edx
  loc_00E1E73F: push eax
  loc_00E1E740: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E746: push 006E147Ch ; "select * from biaya where no_daftar like '" & Chr(37)
  loc_00E1E74B: mov eax, var_18
  loc_00E1E74E: push eax
  loc_00E1E74F: call [0040106Ch] ; __vbaStrCat
  loc_00E1E755: mov edx, eax
  loc_00E1E757: lea ecx, var_1C
  loc_00E1E75A: call [00401228h] ; __vbaStrMove
  loc_00E1E760: push eax
  loc_00E1E761: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E1E766: call [0040106Ch] ; __vbaStrCat
  loc_00E1E76C: mov edx, eax
  loc_00E1E76E: lea ecx, var_20
  loc_00E1E771: call [00401228h] ; __vbaStrMove
  loc_00E1E777: mov edx, eax
  loc_00E1E779: lea ecx, [esi+00000034h]
  loc_00E1E77C: call [004011B0h] ; __vbaStrCopy
  loc_00E1E782: lea ecx, var_20
  loc_00E1E785: lea edx, var_1C
  loc_00E1E788: push ecx
  loc_00E1E789: lea eax, var_18
  loc_00E1E78C: push edx
  loc_00E1E78D: push eax
  loc_00E1E78E: push 00000003h
  loc_00E1E790: call [004011BCh] ; __vbaFreeStrList
  loc_00E1E796: add esp, 00000010h
  loc_00E1E799: lea ecx, var_24
  loc_00E1E79C: call ebx
  loc_00E1E79E: sub esp, 00000010h
  loc_00E1E7A1: mov ecx, 00004008h
  loc_00E1E7A6: mov edx, esp
  loc_00E1E7A8: mov var_74, ecx
  loc_00E1E7AB: lea eax, [esi+00000034h]
  loc_00E1E7AE: push 0000000Eh
  loc_00E1E7B0: mov [edx], ecx
  loc_00E1E7B2: mov ecx, var_70
  loc_00E1E7B5: mov var_6C, eax
  loc_00E1E7B8: push esi
  loc_00E1E7B9: mov [edx+00000004h], ecx
  loc_00E1E7BC: mov ecx, [esi]
  loc_00E1E7BE: mov [edx+00000008h], eax
  loc_00E1E7C1: mov eax, var_68
  loc_00E1E7C4: mov [edx+0000000Ch], eax
  loc_00E1E7C7: call [ecx+00000564h]
  loc_00E1E7CD: lea edx, var_24
  loc_00E1E7D0: push eax
  loc_00E1E7D1: push edx
  loc_00E1E7D2: call edi
  loc_00E1E7D4: push eax
  loc_00E1E7D5: call [00401238h] ; __vbaLateIdSt
  loc_00E1E7DB: lea ecx, var_24
  loc_00E1E7DE: call ebx
  loc_00E1E7E0: mov eax, [esi]
  loc_00E1E7E2: push 00000000h
  loc_00E1E7E4: push 00000019h
  loc_00E1E7E6: push esi
  loc_00E1E7E7: call [eax+00000564h]
  loc_00E1E7ED: lea ecx, var_24
  loc_00E1E7F0: push eax
  loc_00E1E7F1: push ecx
  loc_00E1E7F2: call edi
  loc_00E1E7F4: push eax
  loc_00E1E7F5: call [00401030h] ; __vbaLateIdCall
  loc_00E1E7FB: add esp, 0000000Ch
  loc_00E1E7FE: lea ecx, var_24
  loc_00E1E801: call ebx
  loc_00E1E803: jmp 00E1E941h
  loc_00E1E808: mov edx, [esi]
  loc_00E1E80A: push esi
  loc_00E1E80B: call [edx+00000310h]
  loc_00E1E811: push eax
  loc_00E1E812: lea eax, var_24
  loc_00E1E815: push eax
  loc_00E1E816: call edi
  loc_00E1E818: mov edi, eax
  loc_00E1E81A: lea edx, var_18
  loc_00E1E81D: push edx
  loc_00E1E81E: push edi
  loc_00E1E81F: mov ecx, [edi]
  loc_00E1E821: call [ecx+000000A0h]
  loc_00E1E827: test eax, eax
  loc_00E1E829: fnclex
  loc_00E1E82B: jge 00E1E83Fh
  loc_00E1E82D: push 000000A0h
  loc_00E1E832: push 006DCB70h
  loc_00E1E837: push edi
  loc_00E1E838: push eax
  loc_00E1E839: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E83F: mov eax, var_18
  loc_00E1E842: push eax
  loc_00E1E843: push 006E0708h ; "Keluar"
  loc_00E1E848: call [00401104h] ; __vbaStrCmp
  loc_00E1E84E: mov edi, eax
  loc_00E1E850: lea ecx, var_18
  loc_00E1E853: neg edi
  loc_00E1E855: sbb edi, edi
  loc_00E1E857: inc edi
  loc_00E1E858: neg edi
  loc_00E1E85A: call [00401258h] ; __vbaFreeStr
  loc_00E1E860: lea ecx, var_24
  loc_00E1E863: call ebx
  loc_00E1E865: test di, di
  loc_00E1E868: jz 00E1E8C3h
  loc_00E1E86A: mov eax, [00E3D9CCh]
  loc_00E1E86F: test eax, eax
  loc_00E1E871: jnz 00E1E883h
  loc_00E1E873: push 00E3D9CCh
  loc_00E1E878: push 006DC960h
  loc_00E1E87D: call [004011A0h] ; __vbaNew2
  loc_00E1E883: mov edi, [00E3D9CCh]
  loc_00E1E889: lea ecx, var_24
  loc_00E1E88C: push esi
  loc_00E1E88D: push ecx
  loc_00E1E88E: mov edx, [edi]
  loc_00E1E890: mov var_C0, edx
  loc_00E1E896: call [004010B4h] ; __vbaObjSetAddref
  loc_00E1E89C: mov edx, var_C0
  loc_00E1E8A2: push eax
  loc_00E1E8A3: push edi
  loc_00E1E8A4: call [edx+00000010h]
  loc_00E1E8A7: test eax, eax
  loc_00E1E8A9: fnclex
  loc_00E1E8AB: jge 00E1E8BCh
  loc_00E1E8AD: push 00000010h
  loc_00E1E8AF: push 006DC950h
  loc_00E1E8B4: push edi
  loc_00E1E8B5: push eax
  loc_00E1E8B6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E8BC: lea ecx, var_24
  loc_00E1E8BF: call ebx
  loc_00E1E8C1: jmp 00E1E941h
  loc_00E1E8C3: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E1E8C9: mov ecx, 80020004h
  loc_00E1E8CE: mov var_5C, ecx
  loc_00E1E8D1: mov eax, 0000000Ah
  loc_00E1E8D6: mov var_4C, ecx
  loc_00E1E8D9: mov edi, 00000008h
  loc_00E1E8DE: lea edx, var_84
  loc_00E1E8E4: lea ecx, var_44
  loc_00E1E8E7: mov var_64, eax
  loc_00E1E8EA: mov var_54, eax
  loc_00E1E8ED: mov var_7C, 006E07B8h ; "ERROR 001"
  loc_00E1E8F4: mov var_84, edi
  loc_00E1E8FA: call __vbaVarDup
  loc_00E1E8FC: lea edx, var_74
  loc_00E1E8FF: lea ecx, var_34
  loc_00E1E902: mov var_6C, 006E14D8h ; "Ada [CRASH], !!"
  loc_00E1E909: mov var_74, edi
  loc_00E1E90C: call __vbaVarDup
  loc_00E1E90E: lea eax, var_64
  loc_00E1E911: lea ecx, var_54
  loc_00E1E914: push eax
  loc_00E1E915: lea edx, var_44
  loc_00E1E918: push ecx
  loc_00E1E919: push edx
  loc_00E1E91A: lea eax, var_34
  loc_00E1E91D: push 00000040h
  loc_00E1E91F: push eax
  loc_00E1E920: call [004010A8h] ; rtcMsgBox
  loc_00E1E926: lea ecx, var_64
  loc_00E1E929: lea edx, var_54
  loc_00E1E92C: push ecx
  loc_00E1E92D: lea eax, var_44
  loc_00E1E930: push edx
  loc_00E1E931: lea ecx, var_34
  loc_00E1E934: push eax
  loc_00E1E935: push ecx
  loc_00E1E936: push 00000004h
  loc_00E1E938: call [00401038h] ; __vbaFreeVarList
  loc_00E1E93E: add esp, 00000014h
  loc_00E1E941: mov var_4, 00000000h
  loc_00E1E948: push 00E1E98Ch
  loc_00E1E94D: jmp 00E1E98Bh
  loc_00E1E94F: lea edx, var_20
  loc_00E1E952: lea eax, var_1C
  loc_00E1E955: push edx
  loc_00E1E956: lea ecx, var_18
  loc_00E1E959: push eax
  loc_00E1E95A: push ecx
  loc_00E1E95B: push 00000003h
  loc_00E1E95D: call [004011BCh] ; __vbaFreeStrList
  loc_00E1E963: add esp, 00000010h
  loc_00E1E966: lea ecx, var_24
  loc_00E1E969: call [00401254h] ; __vbaFreeObj
  loc_00E1E96F: lea edx, var_64
  loc_00E1E972: lea eax, var_54
  loc_00E1E975: push edx
  loc_00E1E976: lea ecx, var_44
  loc_00E1E979: push eax
  loc_00E1E97A: lea edx, var_34
  loc_00E1E97D: push ecx
  loc_00E1E97E: push edx
  loc_00E1E97F: push 00000004h
  loc_00E1E981: call [00401038h] ; __vbaFreeVarList
  loc_00E1E987: add esp, 00000014h
  loc_00E1E98A: ret
  loc_00E1E98B: ret
  loc_00E1E98C: mov eax, Me
  loc_00E1E98F: push eax
  loc_00E1E990: mov ecx, [eax]
  loc_00E1E992: call [ecx+00000008h]
  loc_00E1E995: mov eax, var_4
  loc_00E1E998: mov ecx, var_14
  loc_00E1E99B: pop edi
  loc_00E1E99C: pop esi
  loc_00E1E99D: mov fs:[00000000h], ecx
  loc_00E1E9A4: pop ebx
  loc_00E1E9A5: mov esp, ebp
  loc_00E1E9A7: pop ebp
  loc_00E1E9A8: retn 0008h
End Sub

Private Sub optno_Click() 'E1A140
  loc_00E1A140: push ebp
  loc_00E1A141: mov ebp, esp
  loc_00E1A143: sub esp, 0000000Ch
  loc_00E1A146: push 00402836h ; __vbaExceptHandler
  loc_00E1A14B: mov eax, fs:[00000000h]
  loc_00E1A151: push eax
  loc_00E1A152: mov fs:[00000000h], esp
  loc_00E1A159: sub esp, 00000014h
  loc_00E1A15C: push ebx
  loc_00E1A15D: push esi
  loc_00E1A15E: push edi
  loc_00E1A15F: mov var_C, esp
  loc_00E1A162: mov var_8, 004023A0h
  loc_00E1A169: mov esi, Me
  loc_00E1A16C: mov eax, esi
  loc_00E1A16E: and eax, 00000001h
  loc_00E1A171: mov var_4, eax
  loc_00E1A174: and esi, FFFFFFFEh
  loc_00E1A177: push esi
  loc_00E1A178: mov Me, esi
  loc_00E1A17B: mov ecx, [esi]
  loc_00E1A17D: call [ecx+00000004h]
  loc_00E1A180: mov edx, [esi]
  loc_00E1A182: xor edi, edi
  loc_00E1A184: push esi
  loc_00E1A185: mov var_18, edi
  loc_00E1A188: call [edx+00000310h]
  loc_00E1A18E: push eax
  loc_00E1A18F: lea eax, var_18
  loc_00E1A192: push eax
  loc_00E1A193: call [004010ACh] ; __vbaObjSet
  loc_00E1A199: mov esi, eax
  loc_00E1A19B: push esi
  loc_00E1A19C: mov ecx, [esi]
  loc_00E1A19E: call [ecx+00000204h]
  loc_00E1A1A4: cmp eax, edi
  loc_00E1A1A6: fnclex
  loc_00E1A1A8: jge 00E1A1BCh
  loc_00E1A1AA: push 00000204h
  loc_00E1A1AF: push 006DCB70h
  loc_00E1A1B4: push esi
  loc_00E1A1B5: push eax
  loc_00E1A1B6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1A1BC: lea ecx, var_18
  loc_00E1A1BF: call [00401254h] ; __vbaFreeObj
  loc_00E1A1C5: mov var_4, edi
  loc_00E1A1C8: push 00E1A1DAh
  loc_00E1A1CD: jmp 00E1A1D9h
  loc_00E1A1CF: lea ecx, var_18
  loc_00E1A1D2: call [00401254h] ; __vbaFreeObj
  loc_00E1A1D8: ret
  loc_00E1A1D9: ret
  loc_00E1A1DA: mov eax, Me
  loc_00E1A1DD: push eax
  loc_00E1A1DE: mov edx, [eax]
  loc_00E1A1E0: call [edx+00000008h]
  loc_00E1A1E3: mov eax, var_4
  loc_00E1A1E6: mov ecx, var_14
  loc_00E1A1E9: pop edi
  loc_00E1A1EA: pop esi
  loc_00E1A1EB: mov fs:[00000000h], ecx
  loc_00E1A1F2: pop ebx
  loc_00E1A1F3: mov esp, ebp
  loc_00E1A1F5: pop ebp
  loc_00E1A1F6: retn 0004h
End Sub

Private Sub Form_Load() 'E18F80
  loc_00E18F80: push ebp
  loc_00E18F81: mov ebp, esp
  loc_00E18F83: sub esp, 0000000Ch
  loc_00E18F86: push 00402836h ; __vbaExceptHandler
  loc_00E18F8B: mov eax, fs:[00000000h]
  loc_00E18F91: push eax
  loc_00E18F92: mov fs:[00000000h], esp
  loc_00E18F99: sub esp, 00000044h
  loc_00E18F9C: push ebx
  loc_00E18F9D: push esi
  loc_00E18F9E: push edi
  loc_00E18F9F: mov var_C, esp
  loc_00E18FA2: mov var_8, 00402330h
  loc_00E18FA9: mov esi, Me
  loc_00E18FAC: mov eax, esi
  loc_00E18FAE: and eax, 00000001h
  loc_00E18FB1: mov var_4, eax
  loc_00E18FB4: and esi, FFFFFFFEh
  loc_00E18FB7: push esi
  loc_00E18FB8: mov Me, esi
  loc_00E18FBB: mov ecx, [esi]
  loc_00E18FBD: call [ecx+00000004h]
  loc_00E18FC0: lea edi, [esi+00000034h]
  loc_00E18FC3: xor eax, eax
  loc_00E18FC5: mov edx, 006E1140h ; "select * from biaya"
  loc_00E18FCA: mov ecx, edi
  loc_00E18FCC: mov var_18, eax
  loc_00E18FCF: mov var_1C, eax
  loc_00E18FD2: mov var_20, eax
  loc_00E18FD5: mov var_30, eax
  loc_00E18FD8: call [004011B0h] ; __vbaStrCopy
  loc_00E18FDE: sub esp, 00000010h
  loc_00E18FE1: mov eax, 00004008h
  loc_00E18FE6: mov edx, esp
  loc_00E18FE8: mov ecx, var_34
  loc_00E18FEB: push 0000000Eh
  loc_00E18FED: push esi
  loc_00E18FEE: mov [edx], eax
  loc_00E18FF0: mov eax, var_3C
  loc_00E18FF3: mov [edx+00000004h], eax
  loc_00E18FF6: mov [edx+00000008h], edi
  loc_00E18FF9: mov [edx+0000000Ch], ecx
  loc_00E18FFC: mov edx, [esi]
  loc_00E18FFE: call [edx+00000564h]
  loc_00E19004: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E1900A: push eax
  loc_00E1900B: lea eax, var_1C
  loc_00E1900E: push eax
  loc_00E1900F: call edi
  loc_00E19011: push eax
  loc_00E19012: call [00401238h] ; __vbaLateIdSt
  loc_00E19018: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E1901E: lea ecx, var_1C
  loc_00E19021: call ebx
  loc_00E19023: mov ecx, [esi]
  loc_00E19025: push 00000000h
  loc_00E19027: push 00000019h
  loc_00E19029: push esi
  loc_00E1902A: call [ecx+00000564h]
  loc_00E19030: lea edx, var_1C
  loc_00E19033: push eax
  loc_00E19034: push edx
  loc_00E19035: call edi
  loc_00E19037: push eax
  loc_00E19038: call [00401030h] ; __vbaLateIdCall
  loc_00E1903E: add esp, 0000000Ch
  loc_00E19041: lea ecx, var_1C
  loc_00E19044: call ebx
  loc_00E19046: mov eax, [esi]
  loc_00E19048: push 006E05D8h
  loc_00E1904D: push esi
  loc_00E1904E: call [eax+00000564h]
  loc_00E19054: lea ecx, var_1C
  loc_00E19057: push eax
  loc_00E19058: push ecx
  loc_00E19059: call edi
  loc_00E1905B: push eax
  loc_00E1905C: call [00401224h] ; __vbaCastObj
  loc_00E19062: lea edx, var_30
  loc_00E19065: mov var_28, eax
  loc_00E19068: push edx
  loc_00E19069: mov var_30, 0000000Dh
  loc_00E19070: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E19076: mov edx, [eax]
  loc_00E19078: sub esp, 00000010h
  loc_00E1907B: mov ecx, esp
  loc_00E1907D: mov [ecx], edx
  loc_00E1907F: mov edx, [eax+00000004h]
  loc_00E19082: mov [ecx+00000004h], edx
  loc_00E19085: mov edx, [eax+00000008h]
  loc_00E19088: push 00000000h
  loc_00E1908A: mov [ecx+00000008h], edx
  loc_00E1908D: push 0000002Ah
  loc_00E1908F: mov eax, [eax+0000000Ch]
  loc_00E19092: push esi
  loc_00E19093: mov [ecx+0000000Ch], eax
  loc_00E19096: mov ecx, [esi]
  loc_00E19098: call [ecx+00000568h]
  loc_00E1909E: lea edx, var_20
  loc_00E190A1: push eax
  loc_00E190A2: push edx
  loc_00E190A3: call edi
  loc_00E190A5: push eax
  loc_00E190A6: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E190AC: lea eax, var_20
  loc_00E190AF: lea ecx, var_1C
  loc_00E190B2: push eax
  loc_00E190B3: push ecx
  loc_00E190B4: push 00000002h
  loc_00E190B6: call [00401048h] ; __vbaFreeObjList
  loc_00E190BC: add esp, 00000028h
  loc_00E190BF: lea ecx, var_30
  loc_00E190C2: call [00401024h] ; __vbaFreeVar
  loc_00E190C8: mov edx, [esi]
  loc_00E190CA: push 00000000h
  loc_00E190CC: push FFFFFDDAh
  loc_00E190D1: push esi
  loc_00E190D2: call [edx+00000568h]
  loc_00E190D8: push eax
  loc_00E190D9: lea eax, var_1C
  loc_00E190DC: push eax
  loc_00E190DD: call edi
  loc_00E190DF: push eax
  loc_00E190E0: call [00401030h] ; __vbaLateIdCall
  loc_00E190E6: add esp, 0000000Ch
  loc_00E190E9: lea ecx, var_1C
  loc_00E190EC: call ebx
  loc_00E190EE: mov ecx, [esi]
  loc_00E190F0: push esi
  loc_00E190F1: call [ecx+00000538h]
  loc_00E190F7: lea edx, var_1C
  loc_00E190FA: push eax
  loc_00E190FB: push edx
  loc_00E190FC: call edi
  loc_00E190FE: mov ebx, eax
  loc_00E19100: lea eax, var_30
  loc_00E19103: push eax
  loc_00E19104: mov var_44, ebx
  loc_00E19107: call [004011D8h] ; rtcGetDateVar
  loc_00E1910D: mov ebx, [ebx]
  loc_00E1910F: lea ecx, var_30
  loc_00E19112: lea edx, var_18
  loc_00E19115: push ecx
  loc_00E19116: push edx
  loc_00E19117: call [00401180h] ; __vbaStrVarVal
  loc_00E1911D: mov var_54, ebx
  loc_00E19120: mov ebx, var_44
  loc_00E19123: push eax
  loc_00E19124: mov eax, var_54
  loc_00E19127: push ebx
  loc_00E19128: call [eax+00000054h]
  loc_00E1912B: test eax, eax
  loc_00E1912D: fnclex
  loc_00E1912F: jge 00E19140h
  loc_00E19131: push 00000054h
  loc_00E19133: push 006E0410h
  loc_00E19138: push ebx
  loc_00E19139: push eax
  loc_00E1913A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19140: lea ecx, var_18
  loc_00E19143: call [00401258h] ; __vbaFreeStr
  loc_00E19149: lea ecx, var_1C
  loc_00E1914C: call [00401254h] ; __vbaFreeObj
  loc_00E19152: lea ecx, var_30
  loc_00E19155: call [00401024h] ; __vbaFreeVar
  loc_00E1915B: mov ecx, [esi]
  loc_00E1915D: push esi
  loc_00E1915E: call [ecx+00000530h]
  loc_00E19164: lea edx, var_1C
  loc_00E19167: push eax
  loc_00E19168: push edx
  loc_00E19169: call edi
  loc_00E1916B: mov ebx, eax
  loc_00E1916D: lea eax, var_30
  loc_00E19170: push eax
  loc_00E19171: mov var_44, ebx
  loc_00E19174: call [004011E8h] ; rtcGetTimeVar
  loc_00E1917A: mov ebx, [ebx]
  loc_00E1917C: lea ecx, var_30
  loc_00E1917F: lea edx, var_18
  loc_00E19182: push ecx
  loc_00E19183: push edx
  loc_00E19184: call [00401180h] ; __vbaStrVarVal
  loc_00E1918A: mov var_58, ebx
  loc_00E1918D: mov ebx, var_44
  loc_00E19190: push eax
  loc_00E19191: mov eax, var_58
  loc_00E19194: push ebx
  loc_00E19195: call [eax+00000054h]
  loc_00E19198: test eax, eax
  loc_00E1919A: fnclex
  loc_00E1919C: jge 00E191ADh
  loc_00E1919E: push 00000054h
  loc_00E191A0: push 006E0410h
  loc_00E191A5: push ebx
  loc_00E191A6: push eax
  loc_00E191A7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E191AD: lea ecx, var_18
  loc_00E191B0: call [00401258h] ; __vbaFreeStr
  loc_00E191B6: lea ecx, var_1C
  loc_00E191B9: call [00401254h] ; __vbaFreeObj
  loc_00E191BF: lea ecx, var_30
  loc_00E191C2: call [00401024h] ; __vbaFreeVar
  loc_00E191C8: mov ecx, [esi]
  loc_00E191CA: push esi
  loc_00E191CB: call [ecx+0000031Ch]
  loc_00E191D1: lea edx, var_1C
  loc_00E191D4: push eax
  loc_00E191D5: push edx
  loc_00E191D6: call edi
  loc_00E191D8: mov ebx, eax
  loc_00E191DA: push 00000000h
  loc_00E191DC: push ebx
  loc_00E191DD: mov eax, [ebx]
  loc_00E191DF: call [eax+0000009Ch]
  loc_00E191E5: test eax, eax
  loc_00E191E7: fnclex
  loc_00E191E9: jge 00E191FDh
  loc_00E191EB: push 0000009Ch
  loc_00E191F0: push 006DCAD0h
  loc_00E191F5: push ebx
  loc_00E191F6: push eax
  loc_00E191F7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E191FD: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E19203: lea ecx, var_1C
  loc_00E19206: call ebx
  loc_00E19208: mov ecx, [esi]
  loc_00E1920A: push esi
  loc_00E1920B: call [ecx+000004D4h]
  loc_00E19211: lea edx, var_1C
  loc_00E19214: push eax
  loc_00E19215: push edx
  loc_00E19216: call edi
  loc_00E19218: mov ecx, [eax]
  loc_00E1921A: push 00000000h
  loc_00E1921C: push eax
  loc_00E1921D: mov var_44, eax
  loc_00E19220: call [ecx+0000008Ch]
  loc_00E19226: test eax, eax
  loc_00E19228: fnclex
  loc_00E1922A: jge 00E19241h
  loc_00E1922C: mov edx, var_44
  loc_00E1922F: push 0000008Ch
  loc_00E19234: push 006DCDA0h
  loc_00E19239: push edx
  loc_00E1923A: push eax
  loc_00E1923B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19241: lea ecx, var_1C
  loc_00E19244: call ebx
  loc_00E19246: mov eax, [esi]
  loc_00E19248: push esi
  loc_00E19249: call [eax+0000032Ch]
  loc_00E1924F: lea ecx, var_1C
  loc_00E19252: push eax
  loc_00E19253: push ecx
  loc_00E19254: call edi
  loc_00E19256: mov edx, [eax]
  loc_00E19258: push 00000000h
  loc_00E1925A: push eax
  loc_00E1925B: mov var_44, eax
  loc_00E1925E: call [edx+0000008Ch]
  loc_00E19264: test eax, eax
  loc_00E19266: fnclex
  loc_00E19268: jge 00E1927Fh
  loc_00E1926A: mov ecx, var_44
  loc_00E1926D: push 0000008Ch
  loc_00E19272: push 006DCDA0h
  loc_00E19277: push ecx
  loc_00E19278: push eax
  loc_00E19279: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1927F: lea ecx, var_1C
  loc_00E19282: call ebx
  loc_00E19284: mov edx, [esi]
  loc_00E19286: push esi
  loc_00E19287: call [edx+00000304h]
  loc_00E1928D: push eax
  loc_00E1928E: lea eax, var_1C
  loc_00E19291: push eax
  loc_00E19292: call edi
  loc_00E19294: mov ecx, [eax]
  loc_00E19296: push 00000000h
  loc_00E19298: push eax
  loc_00E19299: mov var_44, eax
  loc_00E1929C: call [ecx+0000009Ch]
  loc_00E192A2: test eax, eax
  loc_00E192A4: fnclex
  loc_00E192A6: jge 00E192BDh
  loc_00E192A8: mov edx, var_44
  loc_00E192AB: push 0000009Ch
  loc_00E192B0: push 006DCAD0h
  loc_00E192B5: push edx
  loc_00E192B6: push eax
  loc_00E192B7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E192BD: lea ecx, var_1C
  loc_00E192C0: call ebx
  loc_00E192C2: mov eax, [esi]
  loc_00E192C4: push esi
  loc_00E192C5: call [eax+000004CCh]
  loc_00E192CB: lea ecx, var_1C
  loc_00E192CE: push eax
  loc_00E192CF: push ecx
  loc_00E192D0: call edi
  loc_00E192D2: mov edx, [eax]
  loc_00E192D4: push 00000000h
  loc_00E192D6: push eax
  loc_00E192D7: mov var_44, eax
  loc_00E192DA: call [edx+0000008Ch]
  loc_00E192E0: test eax, eax
  loc_00E192E2: fnclex
  loc_00E192E4: jge 00E192FBh
  loc_00E192E6: mov ecx, var_44
  loc_00E192E9: push 0000008Ch
  loc_00E192EE: push 006DCDA0h
  loc_00E192F3: push ecx
  loc_00E192F4: push eax
  loc_00E192F5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E192FB: lea ecx, var_1C
  loc_00E192FE: call ebx
  loc_00E19300: mov edx, [esi]
  loc_00E19302: push esi
  loc_00E19303: call [edx+000002FCh]
  loc_00E19309: push eax
  loc_00E1930A: lea eax, var_1C
  loc_00E1930D: push eax
  loc_00E1930E: call edi
  loc_00E19310: mov ecx, [eax]
  loc_00E19312: push 00000000h
  loc_00E19314: push eax
  loc_00E19315: mov var_44, eax
  loc_00E19318: call [ecx+0000005Ch]
  loc_00E1931B: test eax, eax
  loc_00E1931D: fnclex
  loc_00E1931F: jge 00E19333h
  loc_00E19321: mov edx, var_44
  loc_00E19324: push 0000005Ch
  loc_00E19326: push 006DCAE0h
  loc_00E1932B: push edx
  loc_00E1932C: push eax
  loc_00E1932D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19333: lea ecx, var_1C
  loc_00E19336: call ebx
  loc_00E19338: mov eax, [esi]
  loc_00E1933A: push esi
  loc_00E1933B: call [eax+0000056Ch]
  loc_00E19341: lea ecx, var_1C
  loc_00E19344: push eax
  loc_00E19345: push ecx
  loc_00E19346: call edi
  loc_00E19348: mov esi, eax
  loc_00E1934A: push 00000000h
  loc_00E1934C: push esi
  loc_00E1934D: mov edx, [esi]
  loc_00E1934F: call [edx+0000008Ch]
  loc_00E19355: test eax, eax
  loc_00E19357: fnclex
  loc_00E19359: jge 00E1936Dh
  loc_00E1935B: push 0000008Ch
  loc_00E19360: push 006DCDA0h
  loc_00E19365: push esi
  loc_00E19366: push eax
  loc_00E19367: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1936D: lea ecx, var_1C
  loc_00E19370: call ebx
  loc_00E19372: mov var_4, 00000000h
  loc_00E19379: push 00E193A7h
  loc_00E1937E: jmp 00E193A6h
  loc_00E19380: lea ecx, var_18
  loc_00E19383: call [00401258h] ; __vbaFreeStr
  loc_00E19389: lea eax, var_20
  loc_00E1938C: lea ecx, var_1C
  loc_00E1938F: push eax
  loc_00E19390: push ecx
  loc_00E19391: push 00000002h
  loc_00E19393: call [00401048h] ; __vbaFreeObjList
  loc_00E19399: add esp, 0000000Ch
  loc_00E1939C: lea ecx, var_30
  loc_00E1939F: call [00401024h] ; __vbaFreeVar
  loc_00E193A5: ret
  loc_00E193A6: ret
  loc_00E193A7: mov eax, Me
  loc_00E193AA: push eax
  loc_00E193AB: mov edx, [eax]
  loc_00E193AD: call [edx+00000008h]
  loc_00E193B0: mov eax, var_4
  loc_00E193B3: mov ecx, var_14
  loc_00E193B6: pop edi
  loc_00E193B7: pop esi
  loc_00E193B8: mov fs:[00000000h], ecx
  loc_00E193BF: pop ebx
  loc_00E193C0: mov esp, ebp
  loc_00E193C2: pop ebp
  loc_00E193C3: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E193D0
  loc_00E193D0: push ebp
  loc_00E193D1: mov ebp, esp
  loc_00E193D3: sub esp, 0000000Ch
  loc_00E193D6: push 00402836h ; __vbaExceptHandler
  loc_00E193DB: mov eax, fs:[00000000h]
  loc_00E193E1: push eax
  loc_00E193E2: mov fs:[00000000h], esp
  loc_00E193E9: sub esp, 0000005Ch
  loc_00E193EC: push ebx
  loc_00E193ED: push esi
  loc_00E193EE: push edi
  loc_00E193EF: mov var_C, esp
  loc_00E193F2: mov var_8, 00402340h
  loc_00E193F9: mov esi, Me
  loc_00E193FC: mov eax, esi
  loc_00E193FE: and eax, 00000001h
  loc_00E19401: mov var_4, eax
  loc_00E19404: and esi, FFFFFFFEh
  loc_00E19407: push esi
  loc_00E19408: mov Me, esi
  loc_00E1940B: mov ecx, [esi]
  loc_00E1940D: call [ecx+00000004h]
  loc_00E19410: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19416: xor eax, eax
  loc_00E19418: mov var_18, eax
  loc_00E1941B: mov var_4C, eax
  loc_00E1941E: mov var_50, eax
  loc_00E19421: mov edx, [esi]
  loc_00E19423: lea eax, var_4C
  loc_00E19426: push eax
  loc_00E19427: push esi
  loc_00E19428: call [edx+00000070h]
  loc_00E1942B: test eax, eax
  loc_00E1942D: fnclex
  loc_00E1942F: jge 00E1943Ch
  loc_00E19431: push 00000070h
  loc_00E19433: push 006DFA4Ch
  loc_00E19438: push esi
  loc_00E19439: push eax
  loc_00E1943A: call ebx
  loc_00E1943C: fld real4 ptr var_4C
  loc_00E1943F: fadd st0, real4 ptr [00401298h]
  loc_00E19445: mov ecx, [esi]
  loc_00E19447: push ecx
  loc_00E19448: fnstsw ax
  loc_00E1944A: test al, 0Dh
  loc_00E1944C: jnz 00E19610h
  loc_00E19452: fstp real4 ptr [esp]
  loc_00E19455: push esi
  loc_00E19456: call [ecx+00000074h]
  loc_00E19459: test eax, eax
  loc_00E1945B: fnclex
  loc_00E1945D: jge 00E1946Ah
  loc_00E1945F: push 00000074h
  loc_00E19461: push 006DFA4Ch
  loc_00E19466: push esi
  loc_00E19467: push eax
  loc_00E19468: call ebx
  loc_00E1946A: mov edx, [esi]
  loc_00E1946C: lea eax, var_4C
  loc_00E1946F: push eax
  loc_00E19470: push esi
  loc_00E19471: call [edx+00000070h]
  loc_00E19474: test eax, eax
  loc_00E19476: fnclex
  loc_00E19478: jge 00E19485h
  loc_00E1947A: push 00000070h
  loc_00E1947C: push 006DFA4Ch
  loc_00E19481: push esi
  loc_00E19482: push eax
  loc_00E19483: call ebx
  loc_00E19485: mov ecx, [esi]
  loc_00E19487: lea edx, var_50
  loc_00E1948A: push edx
  loc_00E1948B: push esi
  loc_00E1948C: call [ecx+00000078h]
  loc_00E1948F: test eax, eax
  loc_00E19491: fnclex
  loc_00E19493: jge 00E194A0h
  loc_00E19495: push 00000078h
  loc_00E19497: push 006DFA4Ch
  loc_00E1949C: push esi
  loc_00E1949D: push eax
  loc_00E1949E: call ebx
  loc_00E194A0: sub esp, 00000010h
  loc_00E194A3: mov ecx, 0000000Ah
  loc_00E194A8: mov ebx, esp
  loc_00E194AA: mov eax, 80020004h
  loc_00E194AF: mov edx, eax
  loc_00E194B1: sub esp, 00000010h
  loc_00E194B4: mov [ebx], ecx
  loc_00E194B6: mov ecx, var_44
  loc_00E194B9: mov edi, [esi]
  loc_00E194BB: mov [ebx+00000004h], ecx
  loc_00E194BE: mov ecx, esp
  loc_00E194C0: sub esp, 00000010h
  loc_00E194C3: mov [ebx+00000008h], eax
  loc_00E194C6: mov eax, var_3C
  loc_00E194C9: mov [ebx+0000000Ch], eax
  loc_00E194CC: mov eax, 0000000Ah
  loc_00E194D1: mov [ecx], eax
  loc_00E194D3: mov eax, var_34
  loc_00E194D6: mov [ecx+00000004h], eax
  loc_00E194D9: mov eax, 00000004h
  loc_00E194DE: mov [ecx+00000008h], edx
  loc_00E194E1: mov edx, var_2C
  loc_00E194E4: mov [ecx+0000000Ch], edx
  loc_00E194E7: mov edx, var_24
  loc_00E194EA: mov ecx, esp
  loc_00E194EC: mov [ecx], eax
  loc_00E194EE: mov eax, var_50
  loc_00E194F1: mov [ecx+00000004h], edx
  loc_00E194F4: mov [ecx+00000008h], eax
  loc_00E194F7: mov eax, var_1C
  loc_00E194FA: mov [ecx+0000000Ch], eax
  loc_00E194FD: mov ecx, var_4C
  loc_00E19500: push ecx
  loc_00E19501: push esi
  loc_00E19502: call [edi+000002A4h]
  loc_00E19508: test eax, eax
  loc_00E1950A: fnclex
  loc_00E1950C: jge 00E19524h
  loc_00E1950E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19514: push 000002A4h
  loc_00E19519: push 006DFA4Ch
  loc_00E1951E: push esi
  loc_00E1951F: push eax
  loc_00E19520: call ebx
  loc_00E19522: jmp 00E1952Ah
  loc_00E19524: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1952A: call [004010BCh] ; rtcDoEvents
  loc_00E19530: mov edx, [esi]
  loc_00E19532: lea eax, var_4C
  loc_00E19535: push eax
  loc_00E19536: push esi
  loc_00E19537: call [edx+00000070h]
  loc_00E1953A: test eax, eax
  loc_00E1953C: fnclex
  loc_00E1953E: jge 00E1954Bh
  loc_00E19540: push 00000070h
  loc_00E19542: push 006DFA4Ch
  loc_00E19547: push esi
  loc_00E19548: push eax
  loc_00E19549: call ebx
  loc_00E1954B: mov eax, [00E3D9CCh]
  loc_00E19550: test eax, eax
  loc_00E19552: jnz 00E19564h
  loc_00E19554: push 00E3D9CCh
  loc_00E19559: push 006DC960h
  loc_00E1955E: call [004011A0h] ; __vbaNew2
  loc_00E19564: mov edi, [00E3D9CCh]
  loc_00E1956A: lea edx, var_18
  loc_00E1956D: push edx
  loc_00E1956E: push edi
  loc_00E1956F: mov ecx, [edi]
  loc_00E19571: call [ecx+00000018h]
  loc_00E19574: test eax, eax
  loc_00E19576: fnclex
  loc_00E19578: jge 00E19585h
  loc_00E1957A: push 00000018h
  loc_00E1957C: push 006DC950h
  loc_00E19581: push edi
  loc_00E19582: push eax
  loc_00E19583: call ebx
  loc_00E19585: mov eax, var_18
  loc_00E19588: lea edx, var_50
  loc_00E1958B: push edx
  loc_00E1958C: push eax
  loc_00E1958D: mov ecx, [eax]
  loc_00E1958F: mov edi, eax
  loc_00E19591: call [ecx+00000098h]
  loc_00E19597: test eax, eax
  loc_00E19599: fnclex
  loc_00E1959B: jge 00E195ABh
  loc_00E1959D: push 00000098h
  loc_00E195A2: push 006DCAF0h
  loc_00E195A7: push edi
  loc_00E195A8: push eax
  loc_00E195A9: call ebx
  loc_00E195AB: fld real4 ptr var_4C
  loc_00E195AE: fcomp real4 ptr var_50
  loc_00E195B1: fnstsw ax
  loc_00E195B3: test ah, 41h
  loc_00E195B6: jz 00E195BFh
  loc_00E195B8: mov eax, 00000001h
  loc_00E195BD: jmp 00E195C1h
  loc_00E195BF: xor eax, eax
  loc_00E195C1: neg eax
  loc_00E195C3: lea ecx, var_18
  loc_00E195C6: mov edi, eax
  loc_00E195C8: call [00401254h] ; __vbaFreeObj
  loc_00E195CE: test di, di
  loc_00E195D1: jnz 00E19421h
  loc_00E195D7: mov var_4, 00000000h
  loc_00E195DE: fwait
  loc_00E195DF: push 00E195F1h
  loc_00E195E4: jmp 00E195F0h
  loc_00E195E6: lea ecx, var_18
  loc_00E195E9: call [00401254h] ; __vbaFreeObj
  loc_00E195EF: ret
  loc_00E195F0: ret
  loc_00E195F1: mov eax, Me
  loc_00E195F4: push eax
  loc_00E195F5: mov ecx, [eax]
  loc_00E195F7: call [ecx+00000008h]
  loc_00E195FA: mov eax, var_4
  loc_00E195FD: mov ecx, var_14
  loc_00E19600: pop edi
  loc_00E19601: pop esi
  loc_00E19602: mov fs:[00000000h], ecx
  loc_00E19609: pop ebx
  loc_00E1960A: mov esp, ebp
  loc_00E1960C: pop ebp
  loc_00E1960D: retn 0008h
End Sub

Private Sub call_UnknownEvent_9 'E18C30
  loc_00E18C30: push ebp
  loc_00E18C31: mov ebp, esp
  loc_00E18C33: sub esp, 0000000Ch
  loc_00E18C36: push 00402836h ; __vbaExceptHandler
  loc_00E18C3B: mov eax, fs:[00000000h]
  loc_00E18C41: push eax
  loc_00E18C42: mov fs:[00000000h], esp
  loc_00E18C49: sub esp, 00000034h
  loc_00E18C4C: push ebx
  loc_00E18C4D: push esi
  loc_00E18C4E: push edi
  loc_00E18C4F: mov var_C, esp
  loc_00E18C52: mov var_8, 00402310h
  loc_00E18C59: mov esi, Me
  loc_00E18C5C: mov eax, esi
  loc_00E18C5E: and eax, 00000001h
  loc_00E18C61: mov var_4, eax
  loc_00E18C64: and esi, FFFFFFFEh
  loc_00E18C67: push esi
  loc_00E18C68: mov Me, esi
  loc_00E18C6B: mov ecx, [esi]
  loc_00E18C6D: call [ecx+00000004h]
  loc_00E18C70: mov eax, [00E3D100h]
  loc_00E18C75: mov var_18, 00000000h
  loc_00E18C7C: test eax, eax
  loc_00E18C7E: jnz 00E18C90h
  loc_00E18C80: push 00E3D100h
  loc_00E18C85: push 006CC14Ch
  loc_00E18C8A: call [004011A0h] ; __vbaNew2
  loc_00E18C90: sub esp, 00000010h
  loc_00E18C93: mov ecx, 0000000Ah
  loc_00E18C98: mov ebx, esp
  loc_00E18C9A: mov var_28, ecx
  loc_00E18C9D: mov eax, 80020004h
  loc_00E18CA2: sub esp, 00000010h
  loc_00E18CA5: mov [ebx], ecx
  loc_00E18CA7: mov ecx, var_34
  loc_00E18CAA: mov var_20, eax
  loc_00E18CAD: mov edi, [00E3D100h]
  loc_00E18CB3: mov [ebx+00000004h], ecx
  loc_00E18CB6: mov ecx, esp
  loc_00E18CB8: mov edx, [edi]
  loc_00E18CBA: push edi
  loc_00E18CBB: mov [ebx+00000008h], eax
  loc_00E18CBE: mov eax, var_2C
  loc_00E18CC1: mov [ebx+0000000Ch], eax
  loc_00E18CC4: mov eax, var_28
  loc_00E18CC7: mov [ecx], eax
  loc_00E18CC9: mov eax, var_24
  loc_00E18CCC: mov [ecx+00000004h], eax
  loc_00E18CCF: mov eax, var_20
  loc_00E18CD2: mov [ecx+00000008h], eax
  loc_00E18CD5: mov eax, var_1C
  loc_00E18CD8: mov [ecx+0000000Ch], eax
  loc_00E18CDB: call [edx+000002B0h]
  loc_00E18CE1: test eax, eax
  loc_00E18CE3: fnclex
  loc_00E18CE5: jge 00E18CF9h
  loc_00E18CE7: push 000002B0h
  loc_00E18CEC: push 006E0134h
  loc_00E18CF1: push edi
  loc_00E18CF2: push eax
  loc_00E18CF3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E18CF9: mov ecx, [esi]
  loc_00E18CFB: push esi
  loc_00E18CFC: call [ecx+000004CCh]
  loc_00E18D02: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E18D08: lea edx, var_18
  loc_00E18D0B: push eax
  loc_00E18D0C: push edx
  loc_00E18D0D: call ebx
  loc_00E18D0F: mov edi, eax
  loc_00E18D11: push 00000000h
  loc_00E18D13: push edi
  loc_00E18D14: mov eax, [edi]
  loc_00E18D16: call [eax+0000008Ch]
  loc_00E18D1C: test eax, eax
  loc_00E18D1E: fnclex
  loc_00E18D20: jge 00E18D34h
  loc_00E18D22: push 0000008Ch
  loc_00E18D27: push 006DCDA0h
  loc_00E18D2C: push edi
  loc_00E18D2D: push eax
  loc_00E18D2E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E18D34: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E18D3A: lea ecx, var_18
  loc_00E18D3D: call edi
  loc_00E18D3F: mov ecx, [esi]
  loc_00E18D41: push esi
  loc_00E18D42: call [ecx+000004A0h]
  loc_00E18D48: lea edx, var_18
  loc_00E18D4B: push eax
  loc_00E18D4C: push edx
  loc_00E18D4D: call ebx
  loc_00E18D4F: mov ebx, eax
  loc_00E18D51: push 00000000h
  loc_00E18D53: push ebx
  loc_00E18D54: mov eax, [ebx]
  loc_00E18D56: call [eax+0000008Ch]
  loc_00E18D5C: test eax, eax
  loc_00E18D5E: fnclex
  loc_00E18D60: jge 00E18D74h
  loc_00E18D62: push 0000008Ch
  loc_00E18D67: push 006DCDA0h
  loc_00E18D6C: push ebx
  loc_00E18D6D: push eax
  loc_00E18D6E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E18D74: lea ecx, var_18
  loc_00E18D77: call edi
  loc_00E18D79: mov ecx, [esi]
  loc_00E18D7B: push esi
  loc_00E18D7C: call [ecx+00000340h]
  loc_00E18D82: lea edx, var_18
  loc_00E18D85: push eax
  loc_00E18D86: push edx
  loc_00E18D87: call [004010ACh] ; __vbaObjSet
  loc_00E18D8D: mov ebx, eax
  loc_00E18D8F: push 006E1138h ; "-"
  loc_00E18D94: push ebx
  loc_00E18D95: mov eax, [ebx]
  loc_00E18D97: call [eax+00000054h]
  loc_00E18D9A: test eax, eax
  loc_00E18D9C: fnclex
  loc_00E18D9E: jge 00E18DAFh
  loc_00E18DA0: push 00000054h
  loc_00E18DA2: push 006E0410h
  loc_00E18DA7: push ebx
  loc_00E18DA8: push eax
  loc_00E18DA9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E18DAF: lea ecx, var_18
  loc_00E18DB2: call edi
  loc_00E18DB4: sub esp, 00000010h
  loc_00E18DB7: mov ecx, 0000000Bh
  loc_00E18DBC: mov edx, esp
  loc_00E18DBE: or eax, FFFFFFFFh
  loc_00E18DC1: push 80010007h
  loc_00E18DC6: push esi
  loc_00E18DC7: mov [edx], ecx
  loc_00E18DC9: mov ecx, var_24
  loc_00E18DCC: mov [edx+00000004h], ecx
  loc_00E18DCF: mov ecx, [esi]
  loc_00E18DD1: mov [edx+00000008h], eax
  loc_00E18DD4: mov eax, var_1C
  loc_00E18DD7: mov [edx+0000000Ch], eax
  loc_00E18DDA: call [ecx+000004C8h]
  loc_00E18DE0: lea edx, var_18
  loc_00E18DE3: push eax
  loc_00E18DE4: push edx
  loc_00E18DE5: call [004010ACh] ; __vbaObjSet
  loc_00E18DEB: push eax
  loc_00E18DEC: call [00401238h] ; __vbaLateIdSt
  loc_00E18DF2: lea ecx, var_18
  loc_00E18DF5: call edi
  loc_00E18DF7: mov var_4, 00000000h
  loc_00E18DFE: push 00E18E10h
  loc_00E18E03: jmp 00E18E0Fh
  loc_00E18E05: lea ecx, var_18
  loc_00E18E08: call [00401254h] ; __vbaFreeObj
  loc_00E18E0E: ret
  loc_00E18E0F: ret
  loc_00E18E10: mov eax, Me
  loc_00E18E13: push eax
  loc_00E18E14: mov ecx, [eax]
  loc_00E18E16: call [ecx+00000008h]
  loc_00E18E19: mov eax, var_4
  loc_00E18E1C: mov ecx, var_14
  loc_00E18E1F: pop edi
  loc_00E18E20: pop esi
  loc_00E18E21: mov fs:[00000000h], ecx
  loc_00E18E28: pop ebx
  loc_00E18E29: mov esp, ebp
  loc_00E18E2B: pop ebp
  loc_00E18E2C: retn 0004h
End Sub

Private Sub Timer1_Timer() 'E1CF40
  loc_00E1CF40: push ebp
  loc_00E1CF41: mov ebp, esp
  loc_00E1CF43: sub esp, 0000000Ch
  loc_00E1CF46: push 00402836h ; __vbaExceptHandler
  loc_00E1CF4B: mov eax, fs:[00000000h]
  loc_00E1CF51: push eax
  loc_00E1CF52: mov fs:[00000000h], esp
  loc_00E1CF59: sub esp, 00000040h
  loc_00E1CF5C: push ebx
  loc_00E1CF5D: push esi
  loc_00E1CF5E: push edi
  loc_00E1CF5F: mov var_C, esp
  loc_00E1CF62: mov var_8, 004023D0h
  loc_00E1CF69: mov esi, Me
  loc_00E1CF6C: mov eax, esi
  loc_00E1CF6E: and eax, 00000001h
  loc_00E1CF71: mov var_4, eax
  loc_00E1CF74: and esi, FFFFFFFEh
  loc_00E1CF77: push esi
  loc_00E1CF78: mov Me, esi
  loc_00E1CF7B: mov ecx, [esi]
  loc_00E1CF7D: call [ecx+00000004h]
  loc_00E1CF80: mov edx, [esi]
  loc_00E1CF82: xor eax, eax
  loc_00E1CF84: push esi
  loc_00E1CF85: mov var_18, eax
  loc_00E1CF88: mov var_1C, eax
  loc_00E1CF8B: mov var_20, eax
  loc_00E1CF8E: mov var_30, eax
  loc_00E1CF91: mov var_34, eax
  loc_00E1CF94: mov var_38, eax
  loc_00E1CF97: call [edx+00000538h]
  loc_00E1CF9D: push eax
  loc_00E1CF9E: lea eax, var_1C
  loc_00E1CFA1: push eax
  loc_00E1CFA2: call [004010ACh] ; __vbaObjSet
  loc_00E1CFA8: lea ecx, var_30
  loc_00E1CFAB: mov edi, eax
  loc_00E1CFAD: push ecx
  loc_00E1CFAE: call [004011D8h] ; rtcGetDateVar
  loc_00E1CFB4: mov ebx, [edi]
  loc_00E1CFB6: lea edx, var_30
  loc_00E1CFB9: lea eax, var_18
  loc_00E1CFBC: push edx
  loc_00E1CFBD: push eax
  loc_00E1CFBE: call [00401180h] ; __vbaStrVarVal
  loc_00E1CFC4: push eax
  loc_00E1CFC5: push edi
  loc_00E1CFC6: call [ebx+00000054h]
  loc_00E1CFC9: test eax, eax
  loc_00E1CFCB: fnclex
  loc_00E1CFCD: jge 00E1CFDEh
  loc_00E1CFCF: push 00000054h
  loc_00E1CFD1: push 006E0410h
  loc_00E1CFD6: push edi
  loc_00E1CFD7: push eax
  loc_00E1CFD8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1CFDE: lea ecx, var_18
  loc_00E1CFE1: call [00401258h] ; __vbaFreeStr
  loc_00E1CFE7: lea ecx, var_1C
  loc_00E1CFEA: call [00401254h] ; __vbaFreeObj
  loc_00E1CFF0: lea ecx, var_30
  loc_00E1CFF3: call [00401024h] ; __vbaFreeVar
  loc_00E1CFF9: mov ecx, [esi]
  loc_00E1CFFB: push esi
  loc_00E1CFFC: call [ecx+00000530h]
  loc_00E1D002: lea edx, var_1C
  loc_00E1D005: push eax
  loc_00E1D006: push edx
  loc_00E1D007: call [004010ACh] ; __vbaObjSet
  loc_00E1D00D: mov edi, eax
  loc_00E1D00F: lea eax, var_30
  loc_00E1D012: push eax
  loc_00E1D013: call [004011E8h] ; rtcGetTimeVar
  loc_00E1D019: mov ebx, [edi]
  loc_00E1D01B: lea ecx, var_30
  loc_00E1D01E: lea edx, var_18
  loc_00E1D021: push ecx
  loc_00E1D022: push edx
  loc_00E1D023: call [00401180h] ; __vbaStrVarVal
  loc_00E1D029: push eax
  loc_00E1D02A: push edi
  loc_00E1D02B: call [ebx+00000054h]
  loc_00E1D02E: test eax, eax
  loc_00E1D030: fnclex
  loc_00E1D032: jge 00E1D047h
  loc_00E1D034: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D03A: push 00000054h
  loc_00E1D03C: push 006E0410h
  loc_00E1D041: push edi
  loc_00E1D042: push eax
  loc_00E1D043: call ebx
  loc_00E1D045: jmp 00E1D04Dh
  loc_00E1D047: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D04D: lea ecx, var_18
  loc_00E1D050: call [00401258h] ; __vbaFreeStr
  loc_00E1D056: lea ecx, var_1C
  loc_00E1D059: call [00401254h] ; __vbaFreeObj
  loc_00E1D05F: lea ecx, var_30
  loc_00E1D062: call [00401024h] ; __vbaFreeVar
  loc_00E1D068: mov eax, [esi]
  loc_00E1D06A: push esi
  loc_00E1D06B: call [eax+000004C0h]
  loc_00E1D071: lea ecx, var_20
  loc_00E1D074: push eax
  loc_00E1D075: push ecx
  loc_00E1D076: call [004010ACh] ; __vbaObjSet
  loc_00E1D07C: mov edi, eax
  loc_00E1D07E: lea eax, var_38
  loc_00E1D081: push eax
  loc_00E1D082: push edi
  loc_00E1D083: mov edx, [edi]
  loc_00E1D085: call [edx+00000080h]
  loc_00E1D08B: test eax, eax
  loc_00E1D08D: fnclex
  loc_00E1D08F: jge 00E1D09Fh
  loc_00E1D091: push 00000080h
  loc_00E1D096: push 006E0410h
  loc_00E1D09B: push edi
  loc_00E1D09C: push eax
  loc_00E1D09D: call ebx
  loc_00E1D09F: mov ecx, [esi]
  loc_00E1D0A1: push esi
  loc_00E1D0A2: call [ecx+000004C0h]
  loc_00E1D0A8: lea edx, var_1C
  loc_00E1D0AB: push eax
  loc_00E1D0AC: push edx
  loc_00E1D0AD: call [004010ACh] ; __vbaObjSet
  loc_00E1D0B3: mov edi, eax
  loc_00E1D0B5: lea ecx, var_34
  loc_00E1D0B8: push ecx
  loc_00E1D0B9: push edi
  loc_00E1D0BA: mov eax, [edi]
  loc_00E1D0BC: call [eax+00000070h]
  loc_00E1D0BF: test eax, eax
  loc_00E1D0C1: fnclex
  loc_00E1D0C3: jge 00E1D0D0h
  loc_00E1D0C5: push 00000070h
  loc_00E1D0C7: push 006E0410h
  loc_00E1D0CC: push edi
  loc_00E1D0CD: push eax
  loc_00E1D0CE: call ebx
  loc_00E1D0D0: fld real4 ptr var_38
  loc_00E1D0D3: fadd st0, real4 ptr var_34
  loc_00E1D0D6: fnstsw ax
  loc_00E1D0D8: test al, 0Dh
  loc_00E1D0DA: jnz 00E1D250h
  loc_00E1D0E0: call [004010C0h] ; __vbaFpR4
  loc_00E1D0E6: fcomp real4 ptr [00401888h]
  loc_00E1D0EC: fnstsw ax
  loc_00E1D0EE: test ah, 41h
  loc_00E1D0F1: jz 00E1D0FAh
  loc_00E1D0F3: mov eax, 00000001h
  loc_00E1D0F8: jmp 00E1D0FCh
  loc_00E1D0FA: xor eax, eax
  loc_00E1D0FC: neg eax
  loc_00E1D0FE: mov edi, eax
  loc_00E1D100: lea edx, var_20
  loc_00E1D103: lea eax, var_1C
  loc_00E1D106: push edx
  loc_00E1D107: push eax
  loc_00E1D108: push 00000002h
  loc_00E1D10A: call [00401048h] ; __vbaFreeObjList
  loc_00E1D110: add esp, 0000000Ch
  loc_00E1D113: test di, di
  loc_00E1D116: jz 00E1D173h
  loc_00E1D118: mov ecx, [esi]
  loc_00E1D11A: push esi
  loc_00E1D11B: call [ecx+000004C0h]
  loc_00E1D121: lea edx, var_1C
  loc_00E1D124: push eax
  loc_00E1D125: push edx
  loc_00E1D126: call [004010ACh] ; __vbaObjSet
  loc_00E1D12C: lea ecx, var_34
  loc_00E1D12F: mov edi, eax
  loc_00E1D131: mov eax, [esi]
  loc_00E1D133: push ecx
  loc_00E1D134: push esi
  loc_00E1D135: call [eax+00000080h]
  loc_00E1D13B: test eax, eax
  loc_00E1D13D: fnclex
  loc_00E1D13F: jge 00E1D14Fh
  loc_00E1D141: push 00000080h
  loc_00E1D146: push 006DFA4Ch
  loc_00E1D14B: push esi
  loc_00E1D14C: push eax
  loc_00E1D14D: call ebx
  loc_00E1D14F: mov eax, var_34
  loc_00E1D152: mov edx, [edi]
  loc_00E1D154: push eax
  loc_00E1D155: push edi
  loc_00E1D156: call [edx+00000074h]
  loc_00E1D159: test eax, eax
  loc_00E1D15B: fnclex
  loc_00E1D15D: jge 00E1D16Ah
  loc_00E1D15F: push 00000074h
  loc_00E1D161: push 006E0410h
  loc_00E1D166: push edi
  loc_00E1D167: push eax
  loc_00E1D168: call ebx
  loc_00E1D16A: lea ecx, var_1C
  loc_00E1D16D: call [00401254h] ; __vbaFreeObj
  loc_00E1D173: mov ecx, [esi]
  loc_00E1D175: push esi
  loc_00E1D176: call [ecx+000004C0h]
  loc_00E1D17C: lea edx, var_20
  loc_00E1D17F: push eax
  loc_00E1D180: push edx
  loc_00E1D181: call [004010ACh] ; __vbaObjSet
  loc_00E1D187: mov edi, eax
  loc_00E1D189: mov eax, [esi]
  loc_00E1D18B: push esi
  loc_00E1D18C: call [eax+000004C0h]
  loc_00E1D192: lea ecx, var_1C
  loc_00E1D195: push eax
  loc_00E1D196: push ecx
  loc_00E1D197: call [004010ACh] ; __vbaObjSet
  loc_00E1D19D: mov esi, eax
  loc_00E1D19F: lea eax, var_34
  loc_00E1D1A2: push eax
  loc_00E1D1A3: push esi
  loc_00E1D1A4: mov edx, [esi]
  loc_00E1D1A6: call [edx+00000070h]
  loc_00E1D1A9: test eax, eax
  loc_00E1D1AB: fnclex
  loc_00E1D1AD: jge 00E1D1BAh
  loc_00E1D1AF: push 00000070h
  loc_00E1D1B1: push 006E0410h
  loc_00E1D1B6: push esi
  loc_00E1D1B7: push eax
  loc_00E1D1B8: call ebx
  loc_00E1D1BA: fld real4 ptr var_34
  loc_00E1D1BD: fsub st0, real4 ptr [00402290h]
  loc_00E1D1C3: mov ecx, [edi]
  loc_00E1D1C5: push ecx
  loc_00E1D1C6: fnstsw ax
  loc_00E1D1C8: test al, 0Dh
  loc_00E1D1CA: jnz 00E1D250h
  loc_00E1D1D0: fstp real4 ptr [esp]
  loc_00E1D1D3: push edi
  loc_00E1D1D4: call [ecx+00000074h]
  loc_00E1D1D7: test eax, eax
  loc_00E1D1D9: fnclex
  loc_00E1D1DB: jge 00E1D1E8h
  loc_00E1D1DD: push 00000074h
  loc_00E1D1DF: push 006E0410h
  loc_00E1D1E4: push edi
  loc_00E1D1E5: push eax
  loc_00E1D1E6: call ebx
  loc_00E1D1E8: lea edx, var_20
  loc_00E1D1EB: lea eax, var_1C
  loc_00E1D1EE: push edx
  loc_00E1D1EF: push eax
  loc_00E1D1F0: push 00000002h
  loc_00E1D1F2: call [00401048h] ; __vbaFreeObjList
  loc_00E1D1F8: add esp, 0000000Ch
  loc_00E1D1FB: mov var_4, 00000000h
  loc_00E1D202: fwait
  loc_00E1D203: push 00E1D231h
  loc_00E1D208: jmp 00E1D230h
  loc_00E1D20A: lea ecx, var_18
  loc_00E1D20D: call [00401258h] ; __vbaFreeStr
  loc_00E1D213: lea ecx, var_20
  loc_00E1D216: lea edx, var_1C
  loc_00E1D219: push ecx
  loc_00E1D21A: push edx
  loc_00E1D21B: push 00000002h
  loc_00E1D21D: call [00401048h] ; __vbaFreeObjList
  loc_00E1D223: add esp, 0000000Ch
  loc_00E1D226: lea ecx, var_30
  loc_00E1D229: call [00401024h] ; __vbaFreeVar
  loc_00E1D22F: ret
  loc_00E1D230: ret
  loc_00E1D231: mov eax, Me
  loc_00E1D234: push eax
  loc_00E1D235: mov ecx, [eax]
  loc_00E1D237: call [ecx+00000008h]
  loc_00E1D23A: mov eax, var_4
  loc_00E1D23D: mov ecx, var_14
  loc_00E1D240: pop edi
  loc_00E1D241: pop esi
  loc_00E1D242: mov fs:[00000000h], ecx
  loc_00E1D249: pop ebx
  loc_00E1D24A: mov esp, ebp
  loc_00E1D24C: pop ebp
  loc_00E1D24D: retn 0004h
End Sub

Private Sub Timer2_Timer() 'E1D260
  loc_00E1D260: push ebp
  loc_00E1D261: mov ebp, esp
  loc_00E1D263: sub esp, 0000000Ch
  loc_00E1D266: push 00402836h ; __vbaExceptHandler
  loc_00E1D26B: mov eax, fs:[00000000h]
  loc_00E1D271: push eax
  loc_00E1D272: mov fs:[00000000h], esp
  loc_00E1D279: sub esp, 000000B0h
  loc_00E1D27F: push ebx
  loc_00E1D280: push esi
  loc_00E1D281: push edi
  loc_00E1D282: mov var_C, esp
  loc_00E1D285: mov var_8, 004023E0h
  loc_00E1D28C: mov esi, Me
  loc_00E1D28F: mov eax, esi
  loc_00E1D291: and eax, 00000001h
  loc_00E1D294: mov var_4, eax
  loc_00E1D297: and esi, FFFFFFFEh
  loc_00E1D29A: push esi
  loc_00E1D29B: mov Me, esi
  loc_00E1D29E: mov ecx, [esi]
  loc_00E1D2A0: call [ecx+00000004h]
  loc_00E1D2A3: xor ebx, ebx
  loc_00E1D2A5: lea edi, [esi+00000038h]
  loc_00E1D2A8: mov var_70, ebx
  loc_00E1D2AB: lea edx, var_70
  loc_00E1D2AE: mov ecx, edi
  loc_00E1D2B0: mov var_18, ebx
  loc_00E1D2B3: mov var_1C, ebx
  loc_00E1D2B6: mov var_20, ebx
  loc_00E1D2B9: mov var_30, ebx
  loc_00E1D2BC: mov var_40, ebx
  loc_00E1D2BF: mov var_50, ebx
  loc_00E1D2C2: mov var_60, ebx
  loc_00E1D2C5: mov var_80, ebx
  loc_00E1D2C8: mov var_A4, ebx
  loc_00E1D2CE: mov var_68, ebx
  loc_00E1D2D1: mov var_70, 00000002h
  loc_00E1D2D8: call [0040101Ch] ; __vbaVarMove
  loc_00E1D2DE: lea eax, [esi+00000048h]
  loc_00E1D2E1: lea edx, var_70
  loc_00E1D2E4: push eax
  loc_00E1D2E5: push edx
  loc_00E1D2E6: mov var_68, 00000064h
  loc_00E1D2ED: mov var_70, 00008002h
  loc_00E1D2F4: mov var_C0, eax
  loc_00E1D2FA: call [00401088h] ; __vbaVarTstLe
  loc_00E1D300: test ax, ax
  loc_00E1D303: jz 00E1D71Bh
  loc_00E1D309: lea eax, var_30
  loc_00E1D30C: mov var_28, 80020004h
  loc_00E1D313: push eax
  loc_00E1D314: mov var_30, 0000000Ah
  loc_00E1D31B: call [004010A0h] ; rtcRandomize
  loc_00E1D321: mov ebx, [00401024h] ; __vbaFreeVar
  loc_00E1D327: lea ecx, var_30
  loc_00E1D32A: call ebx
  loc_00E1D32C: lea ecx, var_30
  loc_00E1D32F: mov var_28, 80020004h
  loc_00E1D336: push ecx
  loc_00E1D337: mov var_30, 0000000Ah
  loc_00E1D33E: call [0040109Ch] ; rtcRandomNext
  loc_00E1D344: fstp real4 ptr var_A4
  loc_00E1D34A: fld real4 ptr var_A4
  loc_00E1D350: fmul st0, real4 ptr [004012E4h]
  loc_00E1D356: lea ecx, var_30
  loc_00E1D359: fadd st0, real4 ptr [004012E0h]
  loc_00E1D35F: fstp real8 ptr [esi+00000058h]
  loc_00E1D362: fnstsw ax
  loc_00E1D364: test al, 0Dh
  loc_00E1D366: jnz 00E1D77Bh
  loc_00E1D36C: call ebx
  loc_00E1D36E: fld real8 ptr [esi+00000058h]
  loc_00E1D371: cmp [00E3D000h], 00000000h
  loc_00E1D378: jnz 00E1D382h
  loc_00E1D37A: fdiv st0, real8 ptr [004012D8h]
  loc_00E1D380: jmp 00E1D393h
  loc_00E1D382: push [004012DCh]
  loc_00E1D388: push [004012D8h]
  loc_00E1D38E: call 00402854h ; _adj_fdiv_m64
  loc_00E1D393: lea edx, var_70
  loc_00E1D396: mov ecx, edi
  loc_00E1D398: mov var_70, 00000005h
  loc_00E1D39F: fstp real8 ptr var_68
  loc_00E1D3A2: fnstsw ax
  loc_00E1D3A4: test al, 0Dh
  loc_00E1D3A6: jnz 00E1D77Bh
  loc_00E1D3AC: call [0040101Ch] ; __vbaVarMove
  loc_00E1D3B2: push 00000000h
  loc_00E1D3B4: lea edx, var_30
  loc_00E1D3B7: push edi
  loc_00E1D3B8: push edx
  loc_00E1D3B9: call [00401170h] ; rtcRound
  loc_00E1D3BF: lea edx, var_30
  loc_00E1D3C2: mov ecx, edi
  loc_00E1D3C4: call [0040101Ch] ; __vbaVarMove
  loc_00E1D3CA: lea ecx, var_30
  loc_00E1D3CD: call ebx
  loc_00E1D3CF: mov eax, var_C0
  loc_00E1D3D5: lea ecx, var_30
  loc_00E1D3D8: push eax
  loc_00E1D3D9: push edi
  loc_00E1D3DA: push ecx
  loc_00E1D3DB: call [004011E0h] ; __vbaVarAdd
  loc_00E1D3E1: mov edx, eax
  loc_00E1D3E3: mov ecx, edi
  loc_00E1D3E5: call [0040101Ch] ; __vbaVarMove
  loc_00E1D3EB: lea ecx, var_30
  loc_00E1D3EE: call ebx
  loc_00E1D3F0: lea edx, var_30
  loc_00E1D3F3: push edi
  loc_00E1D3F4: push edx
  loc_00E1D3F5: call [004011FCh] ; rtcVarStrFromVar
  loc_00E1D3FB: lea ecx, var_30
  loc_00E1D3FE: lea eax, [esi+00000060h]
  loc_00E1D401: push ecx
  loc_00E1D402: mov var_C4, eax
  loc_00E1D408: call [00401034h] ; __vbaStrVarMove
  loc_00E1D40E: mov edx, eax
  loc_00E1D410: lea ecx, var_18
  loc_00E1D413: call [00401228h] ; __vbaStrMove
  loc_00E1D419: mov ecx, var_C4
  loc_00E1D41F: mov edx, eax
  loc_00E1D421: call [004011B0h] ; __vbaStrCopy
  loc_00E1D427: lea ecx, var_18
  loc_00E1D42A: call [00401258h] ; __vbaFreeStr
  loc_00E1D430: lea ecx, var_30
  loc_00E1D433: call ebx
  loc_00E1D435: lea edx, var_70
  loc_00E1D438: push edi
  loc_00E1D439: push edx
  loc_00E1D43A: mov var_68, 00000064h
  loc_00E1D441: mov var_70, 00008002h
  loc_00E1D448: call [00401210h] ; __vbaVarTstGe
  loc_00E1D44E: test ax, ax
  loc_00E1D451: jz 00E1D657h
  loc_00E1D457: mov eax, [esi]
  loc_00E1D459: push esi
  loc_00E1D45A: call [eax+000004D0h]
  loc_00E1D460: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E1D466: lea ecx, var_1C
  loc_00E1D469: push eax
  loc_00E1D46A: push ecx
  loc_00E1D46B: call ebx
  loc_00E1D46D: mov edi, eax
  loc_00E1D46F: push 46657400h
  loc_00E1D474: push edi
  loc_00E1D475: mov edx, [edi]
  loc_00E1D477: call [edx+0000007Ch]
  loc_00E1D47A: test eax, eax
  loc_00E1D47C: fnclex
  loc_00E1D47E: jge 00E1D48Fh
  loc_00E1D480: push 0000007Ch
  loc_00E1D482: push 006DCDA0h
  loc_00E1D487: push edi
  loc_00E1D488: push eax
  loc_00E1D489: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D48F: lea ecx, var_1C
  loc_00E1D492: call [00401254h] ; __vbaFreeObj
  loc_00E1D498: mov eax, [esi]
  loc_00E1D49A: push esi
  loc_00E1D49B: call [eax+000004D0h]
  loc_00E1D4A1: lea ecx, var_1C
  loc_00E1D4A4: push eax
  loc_00E1D4A5: push ecx
  loc_00E1D4A6: call ebx
  loc_00E1D4A8: mov edi, eax
  loc_00E1D4AA: lea eax, var_A4
  loc_00E1D4B0: push eax
  loc_00E1D4B1: push edi
  loc_00E1D4B2: mov edx, [edi]
  loc_00E1D4B4: call [edx+00000078h]
  loc_00E1D4B7: test eax, eax
  loc_00E1D4B9: fnclex
  loc_00E1D4BB: jge 00E1D4CCh
  loc_00E1D4BD: push 00000078h
  loc_00E1D4BF: push 006DCDA0h
  loc_00E1D4C4: push edi
  loc_00E1D4C5: push eax
  loc_00E1D4C6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D4CC: cmp var_A4, 46657400h
  loc_00E1D4D6: jnz 00E1D4DFh
  loc_00E1D4D8: mov eax, 00000001h
  loc_00E1D4DD: jmp 00E1D4E1h
  loc_00E1D4DF: xor eax, eax
  loc_00E1D4E1: neg eax
  loc_00E1D4E3: lea ecx, var_1C
  loc_00E1D4E6: mov di, ax
  loc_00E1D4E9: call [00401254h] ; __vbaFreeObj
  loc_00E1D4EF: test di, di
  loc_00E1D4F2: jz 00E1D5E9h
  loc_00E1D4F8: mov ecx, [esi]
  loc_00E1D4FA: push esi
  loc_00E1D4FB: call [ecx+000004D0h]
  loc_00E1D501: lea edx, var_1C
  loc_00E1D504: push eax
  loc_00E1D505: push edx
  loc_00E1D506: call ebx
  loc_00E1D508: mov edi, eax
  loc_00E1D50A: push 00000000h
  loc_00E1D50C: push edi
  loc_00E1D50D: mov eax, [edi]
  loc_00E1D50F: call [eax+0000008Ch]
  loc_00E1D515: test eax, eax
  loc_00E1D517: fnclex
  loc_00E1D519: jge 00E1D52Dh
  loc_00E1D51B: push 0000008Ch
  loc_00E1D520: push 006DCDA0h
  loc_00E1D525: push edi
  loc_00E1D526: push eax
  loc_00E1D527: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D52D: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E1D533: lea ecx, var_1C
  loc_00E1D536: call edi
  loc_00E1D538: mov ecx, [esi]
  loc_00E1D53A: push esi
  loc_00E1D53B: call [ecx+000002FCh]
  loc_00E1D541: lea edx, var_1C
  loc_00E1D544: push eax
  loc_00E1D545: push edx
  loc_00E1D546: call ebx
  loc_00E1D548: mov esi, eax
  loc_00E1D54A: push 00000000h
  loc_00E1D54C: push esi
  loc_00E1D54D: mov eax, [esi]
  loc_00E1D54F: call [eax+0000005Ch]
  loc_00E1D552: test eax, eax
  loc_00E1D554: fnclex
  loc_00E1D556: jge 00E1D567h
  loc_00E1D558: push 0000005Ch
  loc_00E1D55A: push 006DCAE0h
  loc_00E1D55F: push esi
  loc_00E1D560: push eax
  loc_00E1D561: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D567: lea ecx, var_1C
  loc_00E1D56A: call edi
  loc_00E1D56C: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E1D572: mov ecx, 80020004h
  loc_00E1D577: mov var_58, ecx
  loc_00E1D57A: mov eax, 0000000Ah
  loc_00E1D57F: mov var_48, ecx
  loc_00E1D582: mov edi, 00000008h
  loc_00E1D587: lea edx, var_80
  loc_00E1D58A: lea ecx, var_40
  loc_00E1D58D: mov var_60, eax
  loc_00E1D590: mov var_50, eax
  loc_00E1D593: mov var_78, 006E0C30h ; "INFO 001"
  loc_00E1D59A: mov var_80, edi
  loc_00E1D59D: call __vbaVarDup
  loc_00E1D59F: lea edx, var_70
  loc_00E1D5A2: lea ecx, var_30
  loc_00E1D5A5: mov var_68, 006E0BFCh ; "Data Telah Tersimpan !"
  loc_00E1D5AC: mov var_70, edi
  loc_00E1D5AF: call __vbaVarDup
  loc_00E1D5B1: lea ecx, var_60
  loc_00E1D5B4: lea edx, var_50
  loc_00E1D5B7: push ecx
  loc_00E1D5B8: lea eax, var_40
  loc_00E1D5BB: push edx
  loc_00E1D5BC: push eax
  loc_00E1D5BD: lea ecx, var_30
  loc_00E1D5C0: push 00000040h
  loc_00E1D5C2: push ecx
  loc_00E1D5C3: call [004010A8h] ; rtcMsgBox
  loc_00E1D5C9: lea edx, var_60
  loc_00E1D5CC: lea eax, var_50
  loc_00E1D5CF: push edx
  loc_00E1D5D0: lea ecx, var_40
  loc_00E1D5D3: push eax
  loc_00E1D5D4: lea edx, var_30
  loc_00E1D5D7: push ecx
  loc_00E1D5D8: push edx
  loc_00E1D5D9: push 00000004h
  loc_00E1D5DB: call [00401038h] ; __vbaFreeVarList
  loc_00E1D5E1: add esp, 00000014h
  loc_00E1D5E4: jmp 00E1D719h
  loc_00E1D5E9: mov ecx, 80020004h
  loc_00E1D5EE: mov eax, 0000000Ah
  loc_00E1D5F3: mov var_58, ecx
  loc_00E1D5F6: mov var_48, ecx
  loc_00E1D5F9: mov var_38, ecx
  loc_00E1D5FC: lea edx, var_70
  loc_00E1D5FF: lea ecx, var_30
  loc_00E1D602: mov var_60, eax
  loc_00E1D605: mov var_50, eax
  loc_00E1D608: mov var_40, eax
  loc_00E1D60B: mov var_68, 006E12C8h ; "Ada Error Di TImer 2 !"
  loc_00E1D612: mov var_70, 00000008h
  loc_00E1D619: call [004011F0h] ; __vbaVarDup
  loc_00E1D61F: lea eax, var_60
  loc_00E1D622: lea ecx, var_50
  loc_00E1D625: push eax
  loc_00E1D626: lea edx, var_40
  loc_00E1D629: push ecx
  loc_00E1D62A: push edx
  loc_00E1D62B: lea eax, var_30
  loc_00E1D62E: push 00000000h
  loc_00E1D630: push eax
  loc_00E1D631: call [004010A8h] ; rtcMsgBox
  loc_00E1D637: lea ecx, var_60
  loc_00E1D63A: lea edx, var_50
  loc_00E1D63D: push ecx
  loc_00E1D63E: lea eax, var_40
  loc_00E1D641: push edx
  loc_00E1D642: lea ecx, var_30
  loc_00E1D645: push eax
  loc_00E1D646: push ecx
  loc_00E1D647: push 00000004h
  loc_00E1D649: call [00401038h] ; __vbaFreeVarList
  loc_00E1D64F: add esp, 00000014h
  loc_00E1D652: jmp 00E1D719h
  loc_00E1D657: mov edx, [esi]
  loc_00E1D659: push esi
  loc_00E1D65A: call [edx+000004D0h]
  loc_00E1D660: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E1D666: push eax
  loc_00E1D667: lea eax, var_20
  loc_00E1D66A: push eax
  loc_00E1D66B: call ebx
  loc_00E1D66D: mov ecx, [esi]
  loc_00E1D66F: push esi
  loc_00E1D670: mov edi, eax
  loc_00E1D672: call [ecx+000004D0h]
  loc_00E1D678: lea edx, var_1C
  loc_00E1D67B: push eax
  loc_00E1D67C: push edx
  loc_00E1D67D: call ebx
  loc_00E1D67F: mov esi, eax
  loc_00E1D681: lea ecx, var_A4
  loc_00E1D687: push ecx
  loc_00E1D688: push esi
  loc_00E1D689: mov eax, [esi]
  loc_00E1D68B: call [eax+00000078h]
  loc_00E1D68E: test eax, eax
  loc_00E1D690: fnclex
  loc_00E1D692: jge 00E1D6A3h
  loc_00E1D694: push 00000078h
  loc_00E1D696: push 006DCDA0h
  loc_00E1D69B: push esi
  loc_00E1D69C: push eax
  loc_00E1D69D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D6A3: fld real4 ptr var_A4
  loc_00E1D6A9: fadd st0, real4 ptr [004012D0h]
  loc_00E1D6AF: mov edx, [edi]
  loc_00E1D6B1: push ecx
  loc_00E1D6B2: fnstsw ax
  loc_00E1D6B4: test al, 0Dh
  loc_00E1D6B6: jnz 00E1D77Bh
  loc_00E1D6BC: fstp real4 ptr [esp]
  loc_00E1D6BF: push edi
  loc_00E1D6C0: call [edx+0000007Ch]
  loc_00E1D6C3: test eax, eax
  loc_00E1D6C5: fnclex
  loc_00E1D6C7: jge 00E1D6D8h
  loc_00E1D6C9: push 0000007Ch
  loc_00E1D6CB: push 006DCDA0h
  loc_00E1D6D0: push edi
  loc_00E1D6D1: push eax
  loc_00E1D6D2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D6D8: lea eax, var_20
  loc_00E1D6DB: lea ecx, var_1C
  loc_00E1D6DE: push eax
  loc_00E1D6DF: push ecx
  loc_00E1D6E0: push 00000002h
  loc_00E1D6E2: call [00401048h] ; __vbaFreeObjList
  loc_00E1D6E8: mov edx, var_C4
  loc_00E1D6EE: add esp, 0000000Ch
  loc_00E1D6F1: mov eax, [edx]
  loc_00E1D6F3: push eax
  loc_00E1D6F4: call [004011A4h] ; __vbaR8Str
  loc_00E1D6FA: call [0040124Ch] ; __vbaFPInt
  loc_00E1D700: mov ecx, var_C0
  loc_00E1D706: lea edx, var_70
  loc_00E1D709: fstp real8 ptr var_68
  loc_00E1D70C: mov var_70, 00000005h
  loc_00E1D713: call [0040101Ch] ; __vbaVarMove
  loc_00E1D719: xor ebx, ebx
  loc_00E1D71B: mov var_4, ebx
  loc_00E1D71E: fwait
  loc_00E1D71F: push 00E1D75Ch
  loc_00E1D724: jmp 00E1D75Bh
  loc_00E1D726: lea ecx, var_18
  loc_00E1D729: call [00401258h] ; __vbaFreeStr
  loc_00E1D72F: lea ecx, var_20
  loc_00E1D732: lea edx, var_1C
  loc_00E1D735: push ecx
  loc_00E1D736: push edx
  loc_00E1D737: push 00000002h
  loc_00E1D739: call [00401048h] ; __vbaFreeObjList
  loc_00E1D73F: lea eax, var_60
  loc_00E1D742: lea ecx, var_50
  loc_00E1D745: push eax
  loc_00E1D746: lea edx, var_40
  loc_00E1D749: push ecx
  loc_00E1D74A: lea eax, var_30
  loc_00E1D74D: push edx
  loc_00E1D74E: push eax
  loc_00E1D74F: push 00000004h
  loc_00E1D751: call [00401038h] ; __vbaFreeVarList
  loc_00E1D757: add esp, 00000020h
  loc_00E1D75A: ret
  loc_00E1D75B: ret
  loc_00E1D75C: mov eax, Me
  loc_00E1D75F: push eax
  loc_00E1D760: mov ecx, [eax]
  loc_00E1D762: call [ecx+00000008h]
  loc_00E1D765: mov eax, var_4
  loc_00E1D768: mov ecx, var_14
  loc_00E1D76B: pop edi
  loc_00E1D76C: pop esi
  loc_00E1D76D: mov fs:[00000000h], ecx
  loc_00E1D774: pop ebx
  loc_00E1D775: mov esp, ebp
  loc_00E1D777: pop ebp
  loc_00E1D778: retn 0004h
End Sub

Private Sub metodeff_UnknownEvent_9 'E198B0
  loc_00E198B0: push ebp
  loc_00E198B1: mov ebp, esp
  loc_00E198B3: sub esp, 0000000Ch
  loc_00E198B6: push 00402836h ; __vbaExceptHandler
  loc_00E198BB: mov eax, fs:[00000000h]
  loc_00E198C1: push eax
  loc_00E198C2: mov fs:[00000000h], esp
  loc_00E198C9: sub esp, 0000009Ch
  loc_00E198CF: push ebx
  loc_00E198D0: push esi
  loc_00E198D1: push edi
  loc_00E198D2: mov var_C, esp
  loc_00E198D5: mov var_8, 00402370h
  loc_00E198DC: mov esi, Me
  loc_00E198DF: mov eax, esi
  loc_00E198E1: and eax, 00000001h
  loc_00E198E4: mov var_4, eax
  loc_00E198E7: and esi, FFFFFFFEh
  loc_00E198EA: push esi
  loc_00E198EB: mov Me, esi
  loc_00E198EE: mov ecx, [esi]
  loc_00E198F0: call [ecx+00000004h]
  loc_00E198F3: mov edx, [esi]
  loc_00E198F5: xor eax, eax
  loc_00E198F7: push esi
  loc_00E198F8: mov var_18, eax
  loc_00E198FB: mov var_1C, eax
  loc_00E198FE: mov var_2C, eax
  loc_00E19901: mov var_3C, eax
  loc_00E19904: mov var_4C, eax
  loc_00E19907: mov var_5C, eax
  loc_00E1990A: mov var_6C, eax
  loc_00E1990D: mov var_7C, eax
  loc_00E19910: call [edx+00000340h]
  loc_00E19916: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E1991C: push eax
  loc_00E1991D: lea eax, var_1C
  loc_00E19920: push eax
  loc_00E19921: call ebx
  loc_00E19923: mov edi, eax
  loc_00E19925: lea edx, var_18
  loc_00E19928: push edx
  loc_00E19929: push edi
  loc_00E1992A: mov ecx, [edi]
  loc_00E1992C: call [ecx+00000050h]
  loc_00E1992F: test eax, eax
  loc_00E19931: fnclex
  loc_00E19933: jge 00E19944h
  loc_00E19935: push 00000050h
  loc_00E19937: push 006E0410h
  loc_00E1993C: push edi
  loc_00E1993D: push eax
  loc_00E1993E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19944: mov eax, var_18
  loc_00E19947: push eax
  loc_00E19948: push 006E1138h ; "-"
  loc_00E1994D: call [00401104h] ; __vbaStrCmp
  loc_00E19953: mov edi, eax
  loc_00E19955: lea ecx, var_18
  loc_00E19958: neg edi
  loc_00E1995A: sbb edi, edi
  loc_00E1995C: inc edi
  loc_00E1995D: neg edi
  loc_00E1995F: call [00401258h] ; __vbaFreeStr
  loc_00E19965: lea ecx, var_1C
  loc_00E19968: call [00401254h] ; __vbaFreeObj
  loc_00E1996E: test di, di
  loc_00E19971: jz 00E19A28h
  loc_00E19977: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E1997D: mov edx, 80020004h
  loc_00E19982: mov ecx, 0000000Ah
  loc_00E19987: mov var_54, edx
  loc_00E1998A: mov var_5C, ecx
  loc_00E1998D: mov var_44, edx
  loc_00E19990: mov var_4C, ecx
  loc_00E19993: lea edx, var_7C
  loc_00E19996: lea ecx, var_3C
  loc_00E19999: mov var_74, 006E07B8h ; "ERROR 001"
  loc_00E199A0: mov var_7C, 00000008h
  loc_00E199A7: call edi
  loc_00E199A9: lea edx, var_6C
  loc_00E199AC: lea ecx, var_2C
  loc_00E199AF: mov var_64, 006E116Ch ; "Silahkan Hitung jumlah keseluruhan terlebih dahulu !"
  loc_00E199B6: mov var_6C, 00000008h
  loc_00E199BD: call edi
  loc_00E199BF: lea ecx, var_5C
  loc_00E199C2: lea edx, var_4C
  loc_00E199C5: push ecx
  loc_00E199C6: lea eax, var_3C
  loc_00E199C9: push edx
  loc_00E199CA: push eax
  loc_00E199CB: lea ecx, var_2C
  loc_00E199CE: push 00000010h
  loc_00E199D0: push ecx
  loc_00E199D1: call [004010A8h] ; rtcMsgBox
  loc_00E199D7: lea edx, var_5C
  loc_00E199DA: lea eax, var_4C
  loc_00E199DD: push edx
  loc_00E199DE: lea ecx, var_3C
  loc_00E199E1: push eax
  loc_00E199E2: lea edx, var_2C
  loc_00E199E5: push ecx
  loc_00E199E6: push edx
  loc_00E199E7: push 00000004h
  loc_00E199E9: call [00401038h] ; __vbaFreeVarList
  loc_00E199EF: mov eax, [esi]
  loc_00E199F1: add esp, 00000014h
  loc_00E199F4: push esi
  loc_00E199F5: call [eax+0000056Ch]
  loc_00E199FB: lea ecx, var_1C
  loc_00E199FE: push eax
  loc_00E199FF: push ecx
  loc_00E19A00: call ebx
  loc_00E19A02: mov esi, eax
  loc_00E19A04: push FFFFFFFFh
  loc_00E19A06: push esi
  loc_00E19A07: mov edx, [esi]
  loc_00E19A09: call [edx+0000008Ch]
  loc_00E19A0F: test eax, eax
  loc_00E19A11: fnclex
  loc_00E19A13: jge 00E19B13h
  loc_00E19A19: push 0000008Ch
  loc_00E19A1E: push 006DCDA0h
  loc_00E19A23: jmp 00E19B0Bh
  loc_00E19A28: mov eax, [00E3D0D8h]
  loc_00E19A2D: test eax, eax
  loc_00E19A2F: jnz 00E19A41h
  loc_00E19A31: push 00E3D0D8h
  loc_00E19A36: push 006D300Ch
  loc_00E19A3B: call [004011A0h] ; __vbaNew2
  loc_00E19A41: sub esp, 00000010h
  loc_00E19A44: mov ecx, 0000000Ah
  loc_00E19A49: mov ebx, esp
  loc_00E19A4B: mov var_7C, ecx
  loc_00E19A4E: mov var_6C, ecx
  loc_00E19A51: mov edx, 80020004h
  loc_00E19A56: mov [ebx], ecx
  loc_00E19A58: mov ecx, var_78
  loc_00E19A5B: mov eax, edx
  loc_00E19A5D: sub esp, 00000010h
  loc_00E19A60: mov [ebx+00000004h], ecx
  loc_00E19A63: mov var_74, eax
  loc_00E19A66: mov esi, [00E3D0D8h]
  loc_00E19A6C: mov ecx, esp
  loc_00E19A6E: mov [ebx+00000008h], eax
  loc_00E19A71: mov eax, var_70
  loc_00E19A74: mov var_64, edx
  loc_00E19A77: mov edi, [esi]
  loc_00E19A79: mov [ebx+0000000Ch], eax
  loc_00E19A7C: mov eax, var_6C
  loc_00E19A7F: mov [ecx], eax
  loc_00E19A81: mov eax, var_68
  loc_00E19A84: push esi
  loc_00E19A85: mov [ecx+00000004h], eax
  loc_00E19A88: mov [ecx+00000008h], edx
  loc_00E19A8B: mov edx, var_60
  loc_00E19A8E: mov [ecx+0000000Ch], edx
  loc_00E19A91: call [edi+000002B0h]
  loc_00E19A97: test eax, eax
  loc_00E19A99: fnclex
  loc_00E19A9B: jge 00E19AAFh
  loc_00E19A9D: push 000002B0h
  loc_00E19AA2: push 006E0014h
  loc_00E19AA7: push esi
  loc_00E19AA8: push eax
  loc_00E19AA9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19AAF: mov eax, [00E3D9CCh]
  loc_00E19AB4: test eax, eax
  loc_00E19AB6: jnz 00E19AC8h
  loc_00E19AB8: push 00E3D9CCh
  loc_00E19ABD: push 006DC960h
  loc_00E19AC2: call [004011A0h] ; __vbaNew2
  loc_00E19AC8: mov eax, [00E3D060h]
  loc_00E19ACD: mov esi, [00E3D9CCh]
  loc_00E19AD3: test eax, eax
  loc_00E19AD5: jnz 00E19AE7h
  loc_00E19AD7: push 00E3D060h
  loc_00E19ADC: push 006D87C0h
  loc_00E19AE1: call [004011A0h] ; __vbaNew2
  loc_00E19AE7: mov eax, [00E3D060h]
  loc_00E19AEC: mov edi, [esi]
  loc_00E19AEE: lea ecx, var_1C
  loc_00E19AF1: push eax
  loc_00E19AF2: push ecx
  loc_00E19AF3: call [004010B4h] ; __vbaObjSetAddref
  loc_00E19AF9: push eax
  loc_00E19AFA: push esi
  loc_00E19AFB: call [edi+00000010h]
  loc_00E19AFE: test eax, eax
  loc_00E19B00: fnclex
  loc_00E19B02: jge 00E19B13h
  loc_00E19B04: push 00000010h
  loc_00E19B06: push 006DC950h
  loc_00E19B0B: push esi
  loc_00E19B0C: push eax
  loc_00E19B0D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19B13: lea ecx, var_1C
  loc_00E19B16: call [00401254h] ; __vbaFreeObj
  loc_00E19B1C: mov var_4, 00000000h
  loc_00E19B23: push 00E19B59h
  loc_00E19B28: jmp 00E19B58h
  loc_00E19B2A: lea ecx, var_18
  loc_00E19B2D: call [00401258h] ; __vbaFreeStr
  loc_00E19B33: lea ecx, var_1C
  loc_00E19B36: call [00401254h] ; __vbaFreeObj
  loc_00E19B3C: lea edx, var_5C
  loc_00E19B3F: lea eax, var_4C
  loc_00E19B42: push edx
  loc_00E19B43: lea ecx, var_3C
  loc_00E19B46: push eax
  loc_00E19B47: lea edx, var_2C
  loc_00E19B4A: push ecx
  loc_00E19B4B: push edx
  loc_00E19B4C: push 00000004h
  loc_00E19B4E: call [00401038h] ; __vbaFreeVarList
  loc_00E19B54: add esp, 00000014h
  loc_00E19B57: ret
  loc_00E19B58: ret
  loc_00E19B59: mov eax, Me
  loc_00E19B5C: push eax
  loc_00E19B5D: mov ecx, [eax]
  loc_00E19B5F: call [ecx+00000008h]
  loc_00E19B62: mov eax, var_4
  loc_00E19B65: mov ecx, var_14
  loc_00E19B68: pop edi
  loc_00E19B69: pop esi
  loc_00E19B6A: mov fs:[00000000h], ecx
  loc_00E19B71: pop ebx
  loc_00E19B72: mov esp, ebp
  loc_00E19B74: pop ebp
  loc_00E19B75: retn 0004h
End Sub

Private Sub metodes_UnknownEvent_9 'E19B80
  loc_00E19B80: push ebp
  loc_00E19B81: mov ebp, esp
  loc_00E19B83: sub esp, 0000000Ch
  loc_00E19B86: push 00402836h ; __vbaExceptHandler
  loc_00E19B8B: mov eax, fs:[00000000h]
  loc_00E19B91: push eax
  loc_00E19B92: mov fs:[00000000h], esp
  loc_00E19B99: sub esp, 00000044h
  loc_00E19B9C: push ebx
  loc_00E19B9D: push esi
  loc_00E19B9E: push edi
  loc_00E19B9F: mov var_C, esp
  loc_00E19BA2: mov var_8, 00402380h
  loc_00E19BA9: mov eax, Me
  loc_00E19BAC: mov ecx, eax
  loc_00E19BAE: and ecx, 00000001h
  loc_00E19BB1: mov var_4, ecx
  loc_00E19BB4: and al, FEh
  loc_00E19BB6: push eax
  loc_00E19BB7: mov Me, eax
  loc_00E19BBA: mov edx, [eax]
  loc_00E19BBC: call [edx+00000004h]
  loc_00E19BBF: mov ecx, [00E3D0D8h]
  loc_00E19BC5: xor eax, eax
  loc_00E19BC7: cmp ecx, eax
  loc_00E19BC9: mov var_18, eax
  loc_00E19BCC: mov var_1C, eax
  loc_00E19BCF: mov var_20, eax
  loc_00E19BD2: jnz 00E19BE4h
  loc_00E19BD4: push 00E3D0D8h
  loc_00E19BD9: push 006D300Ch
  loc_00E19BDE: call [004011A0h] ; __vbaNew2
  loc_00E19BE4: sub esp, 00000010h
  loc_00E19BE7: mov ecx, 0000000Ah
  loc_00E19BEC: mov ebx, esp
  loc_00E19BEE: mov var_30, ecx
  loc_00E19BF1: mov eax, 80020004h
  loc_00E19BF6: sub esp, 00000010h
  loc_00E19BF9: mov [ebx], ecx
  loc_00E19BFB: mov ecx, var_3C
  loc_00E19BFE: mov edx, eax
  loc_00E19C00: mov esi, [00E3D0D8h]
  loc_00E19C06: mov [ebx+00000004h], ecx
  loc_00E19C09: mov ecx, esp
  loc_00E19C0B: mov edi, [esi]
  loc_00E19C0D: push esi
  loc_00E19C0E: mov [ebx+00000008h], eax
  loc_00E19C11: mov eax, var_34
  loc_00E19C14: mov [ebx+0000000Ch], eax
  loc_00E19C17: mov eax, var_30
  loc_00E19C1A: mov [ecx], eax
  loc_00E19C1C: mov eax, var_2C
  loc_00E19C1F: mov [ecx+00000004h], eax
  loc_00E19C22: mov [ecx+00000008h], edx
  loc_00E19C25: mov edx, var_24
  loc_00E19C28: mov [ecx+0000000Ch], edx
  loc_00E19C2B: call [edi+000002B0h]
  loc_00E19C31: test eax, eax
  loc_00E19C33: fnclex
  loc_00E19C35: jge 00E19C49h
  loc_00E19C37: push 000002B0h
  loc_00E19C3C: push 006E0014h
  loc_00E19C41: push esi
  loc_00E19C42: push eax
  loc_00E19C43: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19C49: mov eax, [00E3D060h]
  loc_00E19C4E: test eax, eax
  loc_00E19C50: jnz 00E19C62h
  loc_00E19C52: push 00E3D060h
  loc_00E19C57: push 006D87C0h
  loc_00E19C5C: call [004011A0h] ; __vbaNew2
  loc_00E19C62: mov esi, [00E3D060h]
  loc_00E19C68: push esi
  loc_00E19C69: mov eax, [esi]
  loc_00E19C6B: call [eax+000002B4h]
  loc_00E19C71: test eax, eax
  loc_00E19C73: fnclex
  loc_00E19C75: jge 00E19C89h
  loc_00E19C77: push 000002B4h
  loc_00E19C7C: push 006DFA4Ch
  loc_00E19C81: push esi
  loc_00E19C82: push eax
  loc_00E19C83: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19C89: mov eax, [00E3D0D8h]
  loc_00E19C8E: test eax, eax
  loc_00E19C90: jnz 00E19CA7h
  loc_00E19C92: push 00E3D0D8h
  loc_00E19C97: push 006D300Ch
  loc_00E19C9C: call [004011A0h] ; __vbaNew2
  loc_00E19CA2: mov eax, [00E3D0D8h]
  loc_00E19CA7: mov ecx, [eax]
  loc_00E19CA9: push eax
  loc_00E19CAA: call [ecx+00000434h]
  loc_00E19CB0: mov esi, [004010ACh] ; __vbaObjSet
  loc_00E19CB6: lea edx, var_20
  loc_00E19CB9: push eax
  loc_00E19CBA: push edx
  loc_00E19CBB: call __vbaObjSet
  loc_00E19CBD: mov ebx, Me
  loc_00E19CC0: mov var_4C, eax
  loc_00E19CC3: push ebx
  loc_00E19CC4: mov eax, [ebx]
  loc_00E19CC6: call [eax+0000051Ch]
  loc_00E19CCC: lea ecx, var_1C
  loc_00E19CCF: push eax
  loc_00E19CD0: push ecx
  loc_00E19CD1: call __vbaObjSet
  loc_00E19CD3: mov edi, eax
  loc_00E19CD5: lea eax, var_18
  loc_00E19CD8: push eax
  loc_00E19CD9: push edi
  loc_00E19CDA: mov edx, [edi]
  loc_00E19CDC: call [edx+00000050h]
  loc_00E19CDF: test eax, eax
  loc_00E19CE1: fnclex
  loc_00E19CE3: jge 00E19CF4h
  loc_00E19CE5: push 00000050h
  loc_00E19CE7: push 006E0410h
  loc_00E19CEC: push edi
  loc_00E19CED: push eax
  loc_00E19CEE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19CF4: mov edi, var_4C
  loc_00E19CF7: mov edx, var_18
  loc_00E19CFA: push edx
  loc_00E19CFB: push edi
  loc_00E19CFC: mov ecx, [edi]
  loc_00E19CFE: call [ecx+00000054h]
  loc_00E19D01: test eax, eax
  loc_00E19D03: fnclex
  loc_00E19D05: jge 00E19D16h
  loc_00E19D07: push 00000054h
  loc_00E19D09: push 006E0410h
  loc_00E19D0E: push edi
  loc_00E19D0F: push eax
  loc_00E19D10: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19D16: lea ecx, var_18
  loc_00E19D19: call [00401258h] ; __vbaFreeStr
  loc_00E19D1F: lea eax, var_20
  loc_00E19D22: lea ecx, var_1C
  loc_00E19D25: push eax
  loc_00E19D26: push ecx
  loc_00E19D27: push 00000002h
  loc_00E19D29: call [00401048h] ; __vbaFreeObjList
  loc_00E19D2F: mov eax, [00E3D0D8h]
  loc_00E19D34: add esp, 0000000Ch
  loc_00E19D37: test eax, eax
  loc_00E19D39: jnz 00E19D50h
  loc_00E19D3B: push 00E3D0D8h
  loc_00E19D40: push 006D300Ch
  loc_00E19D45: call [004011A0h] ; __vbaNew2
  loc_00E19D4B: mov eax, [00E3D0D8h]
  loc_00E19D50: mov edx, [eax]
  loc_00E19D52: push eax
  loc_00E19D53: call [edx+00000430h]
  loc_00E19D59: push eax
  loc_00E19D5A: lea eax, var_20
  loc_00E19D5D: push eax
  loc_00E19D5E: call __vbaObjSet
  loc_00E19D60: mov ecx, [ebx]
  loc_00E19D62: push ebx
  loc_00E19D63: mov var_4C, eax
  loc_00E19D66: call [ecx+00000520h]
  loc_00E19D6C: lea edx, var_1C
  loc_00E19D6F: push eax
  loc_00E19D70: push edx
  loc_00E19D71: call __vbaObjSet
  loc_00E19D73: mov edi, eax
  loc_00E19D75: lea ecx, var_18
  loc_00E19D78: push ecx
  loc_00E19D79: push edi
  loc_00E19D7A: mov eax, [edi]
  loc_00E19D7C: call [eax+00000050h]
  loc_00E19D7F: test eax, eax
  loc_00E19D81: fnclex
  loc_00E19D83: jge 00E19D94h
  loc_00E19D85: push 00000050h
  loc_00E19D87: push 006E0410h
  loc_00E19D8C: push edi
  loc_00E19D8D: push eax
  loc_00E19D8E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19D94: mov edi, var_4C
  loc_00E19D97: mov eax, var_18
  loc_00E19D9A: push eax
  loc_00E19D9B: push edi
  loc_00E19D9C: mov edx, [edi]
  loc_00E19D9E: call [edx+00000054h]
  loc_00E19DA1: test eax, eax
  loc_00E19DA3: fnclex
  loc_00E19DA5: jge 00E19DB6h
  loc_00E19DA7: push 00000054h
  loc_00E19DA9: push 006E0410h
  loc_00E19DAE: push edi
  loc_00E19DAF: push eax
  loc_00E19DB0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19DB6: lea ecx, var_18
  loc_00E19DB9: call [00401258h] ; __vbaFreeStr
  loc_00E19DBF: lea ecx, var_20
  loc_00E19DC2: lea edx, var_1C
  loc_00E19DC5: push ecx
  loc_00E19DC6: push edx
  loc_00E19DC7: push 00000002h
  loc_00E19DC9: call [00401048h] ; __vbaFreeObjList
  loc_00E19DCF: mov eax, [00E3D0D8h]
  loc_00E19DD4: add esp, 0000000Ch
  loc_00E19DD7: test eax, eax
  loc_00E19DD9: jnz 00E19DF0h
  loc_00E19DDB: push 00E3D0D8h
  loc_00E19DE0: push 006D300Ch
  loc_00E19DE5: call [004011A0h] ; __vbaNew2
  loc_00E19DEB: mov eax, [00E3D0D8h]
  loc_00E19DF0: mov ecx, [eax]
  loc_00E19DF2: push eax
  loc_00E19DF3: call [ecx+00000438h]
  loc_00E19DF9: lea edx, var_20
  loc_00E19DFC: push eax
  loc_00E19DFD: push edx
  loc_00E19DFE: call __vbaObjSet
  loc_00E19E00: mov var_4C, eax
  loc_00E19E03: mov eax, [ebx]
  loc_00E19E05: push ebx
  loc_00E19E06: call [eax+00000518h]
  loc_00E19E0C: lea ecx, var_1C
  loc_00E19E0F: push eax
  loc_00E19E10: push ecx
  loc_00E19E11: call __vbaObjSet
  loc_00E19E13: mov edi, eax
  loc_00E19E15: lea eax, var_18
  loc_00E19E18: push eax
  loc_00E19E19: push edi
  loc_00E19E1A: mov edx, [edi]
  loc_00E19E1C: call [edx+00000050h]
  loc_00E19E1F: test eax, eax
  loc_00E19E21: fnclex
  loc_00E19E23: jge 00E19E34h
  loc_00E19E25: push 00000050h
  loc_00E19E27: push 006E0410h
  loc_00E19E2C: push edi
  loc_00E19E2D: push eax
  loc_00E19E2E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19E34: mov edi, var_4C
  loc_00E19E37: mov edx, var_18
  loc_00E19E3A: push edx
  loc_00E19E3B: push edi
  loc_00E19E3C: mov ecx, [edi]
  loc_00E19E3E: call [ecx+00000054h]
  loc_00E19E41: test eax, eax
  loc_00E19E43: fnclex
  loc_00E19E45: jge 00E19E56h
  loc_00E19E47: push 00000054h
  loc_00E19E49: push 006E0410h
  loc_00E19E4E: push edi
  loc_00E19E4F: push eax
  loc_00E19E50: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19E56: lea ecx, var_18
  loc_00E19E59: call [00401258h] ; __vbaFreeStr
  loc_00E19E5F: lea eax, var_20
  loc_00E19E62: lea ecx, var_1C
  loc_00E19E65: push eax
  loc_00E19E66: push ecx
  loc_00E19E67: push 00000002h
  loc_00E19E69: call [00401048h] ; __vbaFreeObjList
  loc_00E19E6F: mov eax, [00E3D0D8h]
  loc_00E19E74: add esp, 0000000Ch
  loc_00E19E77: test eax, eax
  loc_00E19E79: jnz 00E19E90h
  loc_00E19E7B: push 00E3D0D8h
  loc_00E19E80: push 006D300Ch
  loc_00E19E85: call [004011A0h] ; __vbaNew2
  loc_00E19E8B: mov eax, [00E3D0D8h]
  loc_00E19E90: mov edx, [eax]
  loc_00E19E92: push eax
  loc_00E19E93: call [edx+0000044Ch]
  loc_00E19E99: push eax
  loc_00E19E9A: lea eax, var_20
  loc_00E19E9D: push eax
  loc_00E19E9E: call __vbaObjSet
  loc_00E19EA0: mov ecx, [ebx]
  loc_00E19EA2: push ebx
  loc_00E19EA3: mov var_4C, eax
  loc_00E19EA6: call [ecx+000004E4h]
  loc_00E19EAC: lea edx, var_1C
  loc_00E19EAF: push eax
  loc_00E19EB0: push edx
  loc_00E19EB1: call __vbaObjSet
  loc_00E19EB3: mov edi, eax
  loc_00E19EB5: lea ecx, var_18
  loc_00E19EB8: push ecx
  loc_00E19EB9: push edi
  loc_00E19EBA: mov eax, [edi]
  loc_00E19EBC: call [eax+00000050h]
  loc_00E19EBF: test eax, eax
  loc_00E19EC1: fnclex
  loc_00E19EC3: jge 00E19ED4h
  loc_00E19EC5: push 00000050h
  loc_00E19EC7: push 006E0410h
  loc_00E19ECC: push edi
  loc_00E19ECD: push eax
  loc_00E19ECE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19ED4: mov edi, var_4C
  loc_00E19ED7: mov eax, var_18
  loc_00E19EDA: push eax
  loc_00E19EDB: push edi
  loc_00E19EDC: mov edx, [edi]
  loc_00E19EDE: call [edx+00000054h]
  loc_00E19EE1: test eax, eax
  loc_00E19EE3: fnclex
  loc_00E19EE5: jge 00E19EF6h
  loc_00E19EE7: push 00000054h
  loc_00E19EE9: push 006E0410h
  loc_00E19EEE: push edi
  loc_00E19EEF: push eax
  loc_00E19EF0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19EF6: lea ecx, var_18
  loc_00E19EF9: call [00401258h] ; __vbaFreeStr
  loc_00E19EFF: lea ecx, var_20
  loc_00E19F02: lea edx, var_1C
  loc_00E19F05: push ecx
  loc_00E19F06: push edx
  loc_00E19F07: push 00000002h
  loc_00E19F09: call [00401048h] ; __vbaFreeObjList
  loc_00E19F0F: mov eax, [00E3D0D8h]
  loc_00E19F14: add esp, 0000000Ch
  loc_00E19F17: test eax, eax
  loc_00E19F19: jnz 00E19F30h
  loc_00E19F1B: push 00E3D0D8h
  loc_00E19F20: push 006D300Ch
  loc_00E19F25: call [004011A0h] ; __vbaNew2
  loc_00E19F2B: mov eax, [00E3D0D8h]
  loc_00E19F30: mov ecx, [eax]
  loc_00E19F32: push eax
  loc_00E19F33: call [ecx+00000450h]
  loc_00E19F39: lea edx, var_20
  loc_00E19F3C: push eax
  loc_00E19F3D: push edx
  loc_00E19F3E: call __vbaObjSet
  loc_00E19F40: mov var_4C, eax
  loc_00E19F43: mov eax, [ebx]
  loc_00E19F45: push ebx
  loc_00E19F46: call [eax+000004E0h]
  loc_00E19F4C: lea ecx, var_1C
  loc_00E19F4F: push eax
  loc_00E19F50: push ecx
  loc_00E19F51: call __vbaObjSet
  loc_00E19F53: mov edi, eax
  loc_00E19F55: lea eax, var_18
  loc_00E19F58: push eax
  loc_00E19F59: push edi
  loc_00E19F5A: mov edx, [edi]
  loc_00E19F5C: call [edx+00000050h]
  loc_00E19F5F: test eax, eax
  loc_00E19F61: fnclex
  loc_00E19F63: jge 00E19F74h
  loc_00E19F65: push 00000050h
  loc_00E19F67: push 006E0410h
  loc_00E19F6C: push edi
  loc_00E19F6D: push eax
  loc_00E19F6E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19F74: mov edi, var_4C
  loc_00E19F77: mov edx, var_18
  loc_00E19F7A: push edx
  loc_00E19F7B: push edi
  loc_00E19F7C: mov ecx, [edi]
  loc_00E19F7E: call [ecx+00000054h]
  loc_00E19F81: test eax, eax
  loc_00E19F83: fnclex
  loc_00E19F85: jge 00E19F96h
  loc_00E19F87: push 00000054h
  loc_00E19F89: push 006E0410h
  loc_00E19F8E: push edi
  loc_00E19F8F: push eax
  loc_00E19F90: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19F96: lea ecx, var_18
  loc_00E19F99: call [00401258h] ; __vbaFreeStr
  loc_00E19F9F: lea eax, var_20
  loc_00E19FA2: lea ecx, var_1C
  loc_00E19FA5: push eax
  loc_00E19FA6: push ecx
  loc_00E19FA7: push 00000002h
  loc_00E19FA9: call [00401048h] ; __vbaFreeObjList
  loc_00E19FAF: mov eax, [00E3D0D8h]
  loc_00E19FB4: add esp, 0000000Ch
  loc_00E19FB7: test eax, eax
  loc_00E19FB9: jnz 00E19FD0h
  loc_00E19FBB: push 00E3D0D8h
  loc_00E19FC0: push 006D300Ch
  loc_00E19FC5: call [004011A0h] ; __vbaNew2
  loc_00E19FCB: mov eax, [00E3D0D8h]
  loc_00E19FD0: mov edx, [eax]
  loc_00E19FD2: push eax
  loc_00E19FD3: call [edx+00000368h]
  loc_00E19FD9: push eax
  loc_00E19FDA: lea eax, var_20
  loc_00E19FDD: push eax
  loc_00E19FDE: call __vbaObjSet
  loc_00E19FE0: mov ecx, [ebx]
  loc_00E19FE2: push ebx
  loc_00E19FE3: mov edi, eax
  loc_00E19FE5: call [ecx+00000340h]
  loc_00E19FEB: lea edx, var_1C
  loc_00E19FEE: push eax
  loc_00E19FEF: push edx
  loc_00E19FF0: call __vbaObjSet
  loc_00E19FF2: mov esi, eax
  loc_00E19FF4: lea ecx, var_18
  loc_00E19FF7: push ecx
  loc_00E19FF8: push esi
  loc_00E19FF9: mov eax, [esi]
  loc_00E19FFB: call [eax+00000050h]
  loc_00E19FFE: test eax, eax
  loc_00E1A000: fnclex
  loc_00E1A002: jge 00E1A013h
  loc_00E1A004: push 00000050h
  loc_00E1A006: push 006E0410h
  loc_00E1A00B: push esi
  loc_00E1A00C: push eax
  loc_00E1A00D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1A013: mov eax, var_18
  loc_00E1A016: mov edx, [edi]
  loc_00E1A018: push eax
  loc_00E1A019: push edi
  loc_00E1A01A: call [edx+00000054h]
  loc_00E1A01D: test eax, eax
  loc_00E1A01F: fnclex
  loc_00E1A021: jge 00E1A032h
  loc_00E1A023: push 00000054h
  loc_00E1A025: push 006E0410h
  loc_00E1A02A: push edi
  loc_00E1A02B: push eax
  loc_00E1A02C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1A032: lea ecx, var_18
  loc_00E1A035: call [00401258h] ; __vbaFreeStr
  loc_00E1A03B: lea ecx, var_20
  loc_00E1A03E: lea edx, var_1C
  loc_00E1A041: push ecx
  loc_00E1A042: push edx
  loc_00E1A043: push 00000002h
  loc_00E1A045: call [00401048h] ; __vbaFreeObjList
  loc_00E1A04B: add esp, 0000000Ch
  loc_00E1A04E: mov var_4, 00000000h
  loc_00E1A055: push 00E1A07Ah
  loc_00E1A05A: jmp 00E1A079h
  loc_00E1A05C: lea ecx, var_18
  loc_00E1A05F: call [00401258h] ; __vbaFreeStr
  loc_00E1A065: lea eax, var_20
  loc_00E1A068: lea ecx, var_1C
  loc_00E1A06B: push eax
  loc_00E1A06C: push ecx
  loc_00E1A06D: push 00000002h
  loc_00E1A06F: call [00401048h] ; __vbaFreeObjList
  loc_00E1A075: add esp, 0000000Ch
  loc_00E1A078: ret
  loc_00E1A079: ret
  loc_00E1A07A: mov eax, Me
  loc_00E1A07D: push eax
  loc_00E1A07E: mov edx, [eax]
  loc_00E1A080: call [edx+00000008h]
  loc_00E1A083: mov eax, var_4
  loc_00E1A086: mov ecx, var_14
  loc_00E1A089: pop edi
  loc_00E1A08A: pop esi
  loc_00E1A08B: mov fs:[00000000h], ecx
  loc_00E1A092: pop ebx
  loc_00E1A093: mov esp, ebp
  loc_00E1A095: pop ebp
  loc_00E1A096: retn 0004h
End Sub

Private Sub jcbutton2_UnknownEvent_9 'E19760
  loc_00E19760: push ebp
  loc_00E19761: mov ebp, esp
  loc_00E19763: sub esp, 0000000Ch
  loc_00E19766: push 00402836h ; __vbaExceptHandler
  loc_00E1976B: mov eax, fs:[00000000h]
  loc_00E19771: push eax
  loc_00E19772: mov fs:[00000000h], esp
  loc_00E19779: sub esp, 00000034h
  loc_00E1977C: push ebx
  loc_00E1977D: push esi
  loc_00E1977E: push edi
  loc_00E1977F: mov var_C, esp
  loc_00E19782: mov var_8, 00402360h
  loc_00E19789: mov esi, Me
  loc_00E1978C: mov eax, esi
  loc_00E1978E: and eax, 00000001h
  loc_00E19791: mov var_4, eax
  loc_00E19794: and esi, FFFFFFFEh
  loc_00E19797: push esi
  loc_00E19798: mov Me, esi
  loc_00E1979B: mov ecx, [esi]
  loc_00E1979D: call [ecx+00000004h]
  loc_00E197A0: mov edx, [esi]
  loc_00E197A2: push esi
  loc_00E197A3: mov var_18, 00000000h
  loc_00E197AA: call [edx+00000304h]
  loc_00E197B0: push eax
  loc_00E197B1: lea eax, var_18
  loc_00E197B4: push eax
  loc_00E197B5: call [004010ACh] ; __vbaObjSet
  loc_00E197BB: mov edi, eax
  loc_00E197BD: push 00000000h
  loc_00E197BF: push edi
  loc_00E197C0: mov ecx, [edi]
  loc_00E197C2: call [ecx+0000009Ch]
  loc_00E197C8: test eax, eax
  loc_00E197CA: fnclex
  loc_00E197CC: jge 00E197E0h
  loc_00E197CE: push 0000009Ch
  loc_00E197D3: push 006DCAD0h
  loc_00E197D8: push edi
  loc_00E197D9: push eax
  loc_00E197DA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E197E0: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E197E6: lea ecx, var_18
  loc_00E197E9: call ebx
  loc_00E197EB: mov edx, [esi]
  loc_00E197ED: push esi
  loc_00E197EE: call [edx+0000032Ch]
  loc_00E197F4: push eax
  loc_00E197F5: lea eax, var_18
  loc_00E197F8: push eax
  loc_00E197F9: call [004010ACh] ; __vbaObjSet
  loc_00E197FF: mov edi, eax
  loc_00E19801: push 00000000h
  loc_00E19803: push edi
  loc_00E19804: mov ecx, [edi]
  loc_00E19806: call [ecx+0000008Ch]
  loc_00E1980C: test eax, eax
  loc_00E1980E: fnclex
  loc_00E19810: jge 00E19824h
  loc_00E19812: push 0000008Ch
  loc_00E19817: push 006DCDA0h
  loc_00E1981C: push edi
  loc_00E1981D: push eax
  loc_00E1981E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E19824: lea ecx, var_18
  loc_00E19827: call ebx
  loc_00E19829: sub esp, 00000010h
  loc_00E1982C: mov ecx, 0000000Bh
  loc_00E19831: mov edx, esp
  loc_00E19833: or eax, FFFFFFFFh
  loc_00E19836: push 8001000Dh
  loc_00E1983B: push esi
  loc_00E1983C: mov [edx], ecx
  loc_00E1983E: mov ecx, var_24
  loc_00E19841: mov [edx+00000004h], ecx
  loc_00E19844: mov ecx, [esi]
  loc_00E19846: mov [edx+00000008h], eax
  loc_00E19849: mov eax, var_1C
  loc_00E1984C: mov [edx+0000000Ch], eax
  loc_00E1984F: call [ecx+000004C4h]
  loc_00E19855: lea edx, var_18
  loc_00E19858: push eax
  loc_00E19859: push edx
  loc_00E1985A: call [004010ACh] ; __vbaObjSet
  loc_00E19860: push eax
  loc_00E19861: call [00401238h] ; __vbaLateIdSt
  loc_00E19867: lea ecx, var_18
  loc_00E1986A: call ebx
  loc_00E1986C: mov var_4, 00000000h
  loc_00E19873: push 00E19885h
  loc_00E19878: jmp 00E19884h
  loc_00E1987A: lea ecx, var_18
  loc_00E1987D: call [00401254h] ; __vbaFreeObj
  loc_00E19883: ret
  loc_00E19884: ret
  loc_00E19885: mov eax, Me
  loc_00E19888: push eax
  loc_00E19889: mov ecx, [eax]
  loc_00E1988B: call [ecx+00000008h]
  loc_00E1988E: mov eax, var_4
  loc_00E19891: mov ecx, var_14
  loc_00E19894: pop edi
  loc_00E19895: pop esi
  loc_00E19896: mov fs:[00000000h], ecx
  loc_00E1989D: pop ebx
  loc_00E1989E: mov esp, ebp
  loc_00E198A0: pop ebp
  loc_00E198A1: retn 0004h
End Sub

Private Sub simpan_UnknownEvent_9 'E1A430
  loc_00E1A430: push ebp
  loc_00E1A431: mov ebp, esp
  loc_00E1A433: sub esp, 0000000Ch
  loc_00E1A436: push 00402836h ; __vbaExceptHandler
  loc_00E1A43B: mov eax, fs:[00000000h]
  loc_00E1A441: push eax
  loc_00E1A442: mov fs:[00000000h], esp
  loc_00E1A449: sub esp, 00000124h
  loc_00E1A44F: push ebx
  loc_00E1A450: push esi
  loc_00E1A451: push edi
  loc_00E1A452: mov var_C, esp
  loc_00E1A455: mov var_8, 004023C0h
  loc_00E1A45C: mov esi, Me
  loc_00E1A45F: mov eax, esi
  loc_00E1A461: and eax, 00000001h
  loc_00E1A464: mov var_4, eax
  loc_00E1A467: and esi, FFFFFFFEh
  loc_00E1A46A: push esi
  loc_00E1A46B: mov Me, esi
  loc_00E1A46E: mov ecx, [esi]
  loc_00E1A470: call [ecx+00000004h]
  loc_00E1A473: mov edx, [esi]
  loc_00E1A475: xor ebx, ebx
  loc_00E1A477: push esi
  loc_00E1A478: mov var_18, ebx
  loc_00E1A47B: mov var_1C, ebx
  loc_00E1A47E: mov var_20, ebx
  loc_00E1A481: mov var_24, ebx
  loc_00E1A484: mov var_28, ebx
  loc_00E1A487: mov var_2C, ebx
  loc_00E1A48A: mov var_30, ebx
  loc_00E1A48D: mov var_34, ebx
  loc_00E1A490: mov var_44, ebx
  loc_00E1A493: mov var_54, ebx
  loc_00E1A496: mov var_64, ebx
  loc_00E1A499: mov var_74, ebx
  loc_00E1A49C: mov var_84, ebx
  loc_00E1A4A2: mov var_94, ebx
  loc_00E1A4A8: mov var_B8, ebx
  loc_00E1A4AE: call [edx+00000520h]
  loc_00E1A4B4: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E1A4BA: push eax
  loc_00E1A4BB: lea eax, var_24
  loc_00E1A4BE: push eax
  loc_00E1A4BF: call edi
  loc_00E1A4C1: mov ecx, [eax]
  loc_00E1A4C3: lea edx, var_18
  loc_00E1A4C6: push edx
  loc_00E1A4C7: push eax
  loc_00E1A4C8: mov var_BC, eax
  loc_00E1A4CE: call [ecx+00000050h]
  loc_00E1A4D1: cmp eax, ebx
  loc_00E1A4D3: fnclex
  loc_00E1A4D5: jge 00E1A4ECh
  loc_00E1A4D7: mov ecx, var_BC
  loc_00E1A4DD: push 00000050h
  loc_00E1A4DF: push 006E0410h
  loc_00E1A4E4: push ecx
  loc_00E1A4E5: push eax
  loc_00E1A4E6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1A4EC: mov edx, var_18
  loc_00E1A4EF: push 006E11ECh ; "select * from biaya where no_daftar ='"
  loc_00E1A4F4: push edx
  loc_00E1A4F5: lea ebx, [esi+00000034h]
  loc_00E1A4F8: call [0040106Ch] ; __vbaStrCat
  loc_00E1A4FE: mov edx, eax
  loc_00E1A500: lea ecx, var_1C
  loc_00E1A503: call [00401228h] ; __vbaStrMove
  loc_00E1A509: push eax
  loc_00E1A50A: push 006DCB84h ; "'"
  loc_00E1A50F: call [0040106Ch] ; __vbaStrCat
  loc_00E1A515: mov edx, eax
  loc_00E1A517: lea ecx, var_20
  loc_00E1A51A: call [00401228h] ; __vbaStrMove
  loc_00E1A520: mov edx, eax
  loc_00E1A522: mov ecx, ebx
  loc_00E1A524: call [004011B0h] ; __vbaStrCopy
  loc_00E1A52A: lea eax, var_20
  loc_00E1A52D: lea ecx, var_1C
  loc_00E1A530: push eax
  loc_00E1A531: lea edx, var_18
  loc_00E1A534: push ecx
  loc_00E1A535: push edx
  loc_00E1A536: push 00000003h
  loc_00E1A538: call [004011BCh] ; __vbaFreeStrList
  loc_00E1A53E: add esp, 00000010h
  loc_00E1A541: lea ecx, var_24
  loc_00E1A544: call [00401254h] ; __vbaFreeObj
  loc_00E1A54A: mov edx, var_80
  loc_00E1A54D: sub esp, 00000010h
  loc_00E1A550: mov ecx, esp
  loc_00E1A552: mov eax, 00004008h
  loc_00E1A557: mov var_84, eax
  loc_00E1A55D: push 0000000Eh
  loc_00E1A55F: mov [ecx], eax
  loc_00E1A561: mov eax, var_78
  loc_00E1A564: push esi
  loc_00E1A565: mov var_7C, ebx
  loc_00E1A568: mov [ecx+00000004h], edx
  loc_00E1A56B: mov [ecx+00000008h], ebx
  loc_00E1A56E: mov [ecx+0000000Ch], eax
  loc_00E1A571: mov ecx, [esi]
  loc_00E1A573: call [ecx+00000564h]
  loc_00E1A579: lea edx, var_24
  loc_00E1A57C: push eax
  loc_00E1A57D: push edx
  loc_00E1A57E: call edi
  loc_00E1A580: push eax
  loc_00E1A581: call [00401238h] ; __vbaLateIdSt
  loc_00E1A587: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E1A58D: lea ecx, var_24
  loc_00E1A590: call ebx
  loc_00E1A592: mov eax, [esi]
  loc_00E1A594: push 00000000h
  loc_00E1A596: push 00000019h
  loc_00E1A598: push esi
  loc_00E1A599: call [eax+00000564h]
  loc_00E1A59F: lea ecx, var_24
  loc_00E1A5A2: push eax
  loc_00E1A5A3: push ecx
  loc_00E1A5A4: call edi
  loc_00E1A5A6: push eax
  loc_00E1A5A7: call [00401030h] ; __vbaLateIdCall
  loc_00E1A5AD: add esp, 0000000Ch
  loc_00E1A5B0: lea ecx, var_24
  loc_00E1A5B3: call ebx
  loc_00E1A5B5: mov edx, [esi]
  loc_00E1A5B7: push 006DCBD8h
  loc_00E1A5BC: push 00000000h
  loc_00E1A5BE: push 00000018h
  loc_00E1A5C0: push esi
  loc_00E1A5C1: call [edx+00000564h]
  loc_00E1A5C7: push eax
  loc_00E1A5C8: lea eax, var_24
  loc_00E1A5CB: push eax
  loc_00E1A5CC: call edi
  loc_00E1A5CE: lea ecx, var_44
  loc_00E1A5D1: push eax
  loc_00E1A5D2: push ecx
  loc_00E1A5D3: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1A5D9: add esp, 00000010h
  loc_00E1A5DC: push eax
  loc_00E1A5DD: call [00401120h] ; __vbaCastObjVar
  loc_00E1A5E3: lea edx, var_28
  loc_00E1A5E6: push eax
  loc_00E1A5E7: push edx
  loc_00E1A5E8: call edi
  loc_00E1A5EA: mov ebx, eax
  loc_00E1A5EC: lea ecx, var_B8
  loc_00E1A5F2: push ecx
  loc_00E1A5F3: push ebx
  loc_00E1A5F4: mov eax, [ebx]
  loc_00E1A5F6: call [eax+00000050h]
  loc_00E1A5F9: test eax, eax
  loc_00E1A5FB: fnclex
  loc_00E1A5FD: jge 00E1A60Eh
  loc_00E1A5FF: push 00000050h
  loc_00E1A601: push 006DCBE8h
  loc_00E1A606: push ebx
  loc_00E1A607: push eax
  loc_00E1A608: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1A60E: mov bx, var_B8
  loc_00E1A615: lea edx, var_28
  loc_00E1A618: lea eax, var_24
  loc_00E1A61B: push edx
  loc_00E1A61C: push eax
  loc_00E1A61D: push 00000002h
  loc_00E1A61F: call [00401048h] ; __vbaFreeObjList
  loc_00E1A625: add esp, 0000000Ch
  loc_00E1A628: lea ecx, var_44
  loc_00E1A62B: call [00401024h] ; __vbaFreeVar
  loc_00E1A631: test bx, bx
  loc_00E1A634: jz 00E1CDF6h
  loc_00E1A63A: mov ecx, [esi]
  loc_00E1A63C: push esi
  loc_00E1A63D: call [ecx+00000340h]
  loc_00E1A643: lea edx, var_24
  loc_00E1A646: push eax
  loc_00E1A647: push edx
  loc_00E1A648: call edi
  loc_00E1A64A: mov ebx, eax
  loc_00E1A64C: lea ecx, var_18
  loc_00E1A64F: push ecx
  loc_00E1A650: push ebx
  loc_00E1A651: mov eax, [ebx]
  loc_00E1A653: call [eax+00000050h]
  loc_00E1A656: test eax, eax
  loc_00E1A658: fnclex
  loc_00E1A65A: jge 00E1A66Fh
  loc_00E1A65C: push 00000050h
  loc_00E1A65E: push 006E0410h
  loc_00E1A663: push ebx
  loc_00E1A664: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1A66A: push eax
  loc_00E1A66B: call ebx
  loc_00E1A66D: jmp 00E1A675h
  loc_00E1A66F: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1A675: mov edx, var_18
  loc_00E1A678: push edx
  loc_00E1A679: push 006E1138h ; "-"
  loc_00E1A67E: call [00401104h] ; __vbaStrCmp
  loc_00E1A684: neg eax
  loc_00E1A686: sbb eax, eax
  loc_00E1A688: lea ecx, var_18
  loc_00E1A68B: inc eax
  loc_00E1A68C: neg eax
  loc_00E1A68E: mov var_C4, ax
  loc_00E1A695: call [00401258h] ; __vbaFreeStr
  loc_00E1A69B: lea ecx, var_24
  loc_00E1A69E: call [00401254h] ; __vbaFreeObj
  loc_00E1A6A4: cmp var_C4, 0000h
  loc_00E1A6AC: jz 00E1A73Eh
  loc_00E1A6B2: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E1A6B8: mov ecx, 0000000Ah
  loc_00E1A6BD: mov eax, 80020004h
  loc_00E1A6C2: mov var_74, ecx
  loc_00E1A6C5: mov var_64, ecx
  loc_00E1A6C8: mov edi, 00000008h
  loc_00E1A6CD: lea edx, var_94
  loc_00E1A6D3: lea ecx, var_54
  loc_00E1A6D6: mov var_6C, eax
  loc_00E1A6D9: mov var_5C, eax
  loc_00E1A6DC: mov var_8C, 006E07B8h ; "ERROR 001"
  loc_00E1A6E6: mov var_94, edi
  loc_00E1A6EC: call __vbaVarDup
  loc_00E1A6EE: lea edx, var_84
  loc_00E1A6F4: lea ecx, var_44
  loc_00E1A6F7: mov var_7C, 006E1240h ; "Harap Isi Seluruh Data Sebelum Menyimpan nya !"
  loc_00E1A6FE: mov var_84, edi
  loc_00E1A704: call __vbaVarDup
  loc_00E1A706: lea eax, var_74
  loc_00E1A709: lea ecx, var_64
  loc_00E1A70C: push eax
  loc_00E1A70D: lea edx, var_54
  loc_00E1A710: push ecx
  loc_00E1A711: push edx
  loc_00E1A712: lea eax, var_44
  loc_00E1A715: push 00000040h
  loc_00E1A717: push eax
  loc_00E1A718: call [004010A8h] ; rtcMsgBox
  loc_00E1A71E: lea ecx, var_74
  loc_00E1A721: lea edx, var_64
  loc_00E1A724: push ecx
  loc_00E1A725: lea eax, var_54
  loc_00E1A728: push edx
  loc_00E1A729: lea ecx, var_44
  loc_00E1A72C: push eax
  loc_00E1A72D: push ecx
  loc_00E1A72E: push 00000004h
  loc_00E1A730: call [00401038h] ; __vbaFreeVarList
  loc_00E1A736: add esp, 00000014h
  loc_00E1A739: jmp 00E1CEBEh
  loc_00E1A73E: mov edx, [esi]
  loc_00E1A740: push esi
  loc_00E1A741: call [edx+00000340h]
  loc_00E1A747: push eax
  loc_00E1A748: lea eax, var_24
  loc_00E1A74B: push eax
  loc_00E1A74C: call edi
  loc_00E1A74E: mov ecx, [eax]
  loc_00E1A750: lea edx, var_18
  loc_00E1A753: push edx
  loc_00E1A754: push eax
  loc_00E1A755: mov var_BC, eax
  loc_00E1A75B: call [ecx+00000050h]
  loc_00E1A75E: test eax, eax
  loc_00E1A760: fnclex
  loc_00E1A762: jge 00E1A775h
  loc_00E1A764: mov ecx, var_BC
  loc_00E1A76A: push 00000050h
  loc_00E1A76C: push 006E0410h
  loc_00E1A771: push ecx
  loc_00E1A772: push eax
  loc_00E1A773: call ebx
  loc_00E1A775: mov edx, var_18
  loc_00E1A778: push edx
  loc_00E1A779: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1A77F: call [004010D8h] ; __vbaFpR8
  loc_00E1A785: fcomp real8 ptr [004022E8h]
  loc_00E1A78B: fnstsw ax
  loc_00E1A78D: test ah, 40h
  loc_00E1A790: jz 00E1A799h
  loc_00E1A792: mov eax, 00000001h
  loc_00E1A797: jmp 00E1A79Bh
  loc_00E1A799: xor eax, eax
  loc_00E1A79B: neg eax
  loc_00E1A79D: lea ecx, var_18
  loc_00E1A7A0: mov var_C4, ax
  loc_00E1A7A7: call [00401258h] ; __vbaFreeStr
  loc_00E1A7AD: lea ecx, var_24
  loc_00E1A7B0: call [00401254h] ; __vbaFreeObj
  loc_00E1A7B6: cmp var_C4, 0000h
  loc_00E1A7BE: jnz 00E1A6B2h
  loc_00E1A7C4: mov edx, [esi]
  loc_00E1A7C6: push 006DCBD8h
  loc_00E1A7CB: push 00000000h
  loc_00E1A7CD: push 00000018h
  loc_00E1A7CF: push esi
  loc_00E1A7D0: call [edx+00000564h]
  loc_00E1A7D6: push eax
  loc_00E1A7D7: lea eax, var_24
  loc_00E1A7DA: push eax
  loc_00E1A7DB: call edi
  loc_00E1A7DD: lea ecx, var_44
  loc_00E1A7E0: push eax
  loc_00E1A7E1: push ecx
  loc_00E1A7E2: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1A7E8: add esp, 00000010h
  loc_00E1A7EB: push eax
  loc_00E1A7EC: call [00401120h] ; __vbaCastObjVar
  loc_00E1A7F2: lea edx, var_28
  loc_00E1A7F5: push eax
  loc_00E1A7F6: push edx
  loc_00E1A7F7: call edi
  loc_00E1A7F9: mov edx, 80020004h
  loc_00E1A7FE: sub esp, 00000010h
  loc_00E1A801: mov var_8C, edx
  loc_00E1A807: mov var_7C, edx
  loc_00E1A80A: mov edx, esp
  loc_00E1A80C: mov ecx, 0000000Ah
  loc_00E1A811: mov var_94, ecx
  loc_00E1A817: mov var_84, ecx
  loc_00E1A81D: mov [edx], ecx
  loc_00E1A81F: mov ecx, var_90
  loc_00E1A825: sub esp, 00000010h
  loc_00E1A828: mov var_BC, eax
  loc_00E1A82E: mov [edx+00000004h], ecx
  loc_00E1A831: mov ecx, var_8C
  loc_00E1A837: mov eax, [eax]
  loc_00E1A839: mov [edx+00000008h], ecx
  loc_00E1A83C: mov ecx, var_88
  loc_00E1A842: mov [edx+0000000Ch], ecx
  loc_00E1A845: mov ecx, var_84
  loc_00E1A84B: mov edx, esp
  loc_00E1A84D: mov [edx], ecx
  loc_00E1A84F: mov ecx, var_80
  loc_00E1A852: mov [edx+00000004h], ecx
  loc_00E1A855: mov ecx, var_7C
  loc_00E1A858: mov [edx+00000008h], ecx
  loc_00E1A85B: mov ecx, var_78
  loc_00E1A85E: mov [edx+0000000Ch], ecx
  loc_00E1A861: mov edx, var_BC
  loc_00E1A867: push edx
  loc_00E1A868: call [eax+00000078h]
  loc_00E1A86B: test eax, eax
  loc_00E1A86D: fnclex
  loc_00E1A86F: jge 00E1A882h
  loc_00E1A871: mov ecx, var_BC
  loc_00E1A877: push 00000078h
  loc_00E1A879: push 006DCBE8h
  loc_00E1A87E: push ecx
  loc_00E1A87F: push eax
  loc_00E1A880: call ebx
  loc_00E1A882: lea edx, var_28
  loc_00E1A885: lea eax, var_24
  loc_00E1A888: push edx
  loc_00E1A889: push eax
  loc_00E1A88A: push 00000002h
  loc_00E1A88C: call [00401048h] ; __vbaFreeObjList
  loc_00E1A892: add esp, 0000000Ch
  loc_00E1A895: lea ecx, var_44
  loc_00E1A898: call [00401024h] ; __vbaFreeVar
  loc_00E1A89E: mov ecx, [esi]
  loc_00E1A8A0: push 006DCBD8h
  loc_00E1A8A5: push 00000000h
  loc_00E1A8A7: push 00000018h
  loc_00E1A8A9: push esi
  loc_00E1A8AA: call [ecx+00000564h]
  loc_00E1A8B0: lea edx, var_28
  loc_00E1A8B3: push eax
  loc_00E1A8B4: push edx
  loc_00E1A8B5: call edi
  loc_00E1A8B7: push eax
  loc_00E1A8B8: lea eax, var_44
  loc_00E1A8BB: push eax
  loc_00E1A8BC: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1A8C2: add esp, 00000010h
  loc_00E1A8C5: push eax
  loc_00E1A8C6: call [00401120h] ; __vbaCastObjVar
  loc_00E1A8CC: lea ecx, var_2C
  loc_00E1A8CF: push eax
  loc_00E1A8D0: push ecx
  loc_00E1A8D1: call edi
  loc_00E1A8D3: mov edx, [eax]
  loc_00E1A8D5: lea ecx, var_30
  loc_00E1A8D8: push ecx
  loc_00E1A8D9: push eax
  loc_00E1A8DA: mov var_C4, eax
  loc_00E1A8E0: call [edx+00000054h]
  loc_00E1A8E3: test eax, eax
  loc_00E1A8E5: fnclex
  loc_00E1A8E7: jge 00E1A8FAh
  loc_00E1A8E9: mov edx, var_C4
  loc_00E1A8EF: push 00000054h
  loc_00E1A8F1: push 006DCBE8h
  loc_00E1A8F6: push edx
  loc_00E1A8F7: push eax
  loc_00E1A8F8: call ebx
  loc_00E1A8FA: lea edx, var_34
  loc_00E1A8FD: mov eax, 00000002h
  loc_00E1A902: push edx
  loc_00E1A903: mov ecx, var_30
  loc_00E1A906: sub esp, 00000010h
  loc_00E1A909: mov var_84, eax
  loc_00E1A90F: mov edx, esp
  loc_00E1A911: mov var_7C, 00000000h
  loc_00E1A918: mov var_CC, ecx
  loc_00E1A91E: mov ecx, [ecx]
  loc_00E1A920: mov [edx], eax
  loc_00E1A922: mov eax, var_80
  loc_00E1A925: mov [edx+00000004h], eax
  loc_00E1A928: mov eax, var_7C
  loc_00E1A92B: mov [edx+00000008h], eax
  loc_00E1A92E: mov eax, var_78
  loc_00E1A931: mov [edx+0000000Ch], eax
  loc_00E1A934: mov edx, var_30
  loc_00E1A937: push edx
  loc_00E1A938: call [ecx+00000028h]
  loc_00E1A93B: test eax, eax
  loc_00E1A93D: fnclex
  loc_00E1A93F: jge 00E1A952h
  loc_00E1A941: mov ecx, var_CC
  loc_00E1A947: push 00000028h
  loc_00E1A949: push 006E09E8h
  loc_00E1A94E: push ecx
  loc_00E1A94F: push eax
  loc_00E1A950: call ebx
  loc_00E1A952: mov edx, var_34
  loc_00E1A955: mov eax, [esi]
  loc_00E1A957: push esi
  loc_00E1A958: mov var_D4, edx
  loc_00E1A95E: call [eax+00000520h]
  loc_00E1A964: lea ecx, var_24
  loc_00E1A967: push eax
  loc_00E1A968: push ecx
  loc_00E1A969: call edi
  loc_00E1A96B: mov edx, [eax]
  loc_00E1A96D: lea ecx, var_18
  loc_00E1A970: push ecx
  loc_00E1A971: push eax
  loc_00E1A972: mov var_BC, eax
  loc_00E1A978: call [edx+00000050h]
  loc_00E1A97B: test eax, eax
  loc_00E1A97D: fnclex
  loc_00E1A97F: jge 00E1A992h
  loc_00E1A981: mov edx, var_BC
  loc_00E1A987: push 00000050h
  loc_00E1A989: push 006E0410h
  loc_00E1A98E: push edx
  loc_00E1A98F: push eax
  loc_00E1A990: call ebx
  loc_00E1A992: mov eax, var_18
  loc_00E1A995: mov ecx, var_D4
  loc_00E1A99B: mov var_4C, eax
  loc_00E1A99E: mov eax, 00000008h
  loc_00E1A9A3: mov var_18, 00000000h
  loc_00E1A9AA: mov var_54, eax
  loc_00E1A9AD: mov edx, [ecx]
  loc_00E1A9AF: sub esp, 00000010h
  loc_00E1A9B2: mov ecx, esp
  loc_00E1A9B4: mov [ecx], eax
  loc_00E1A9B6: mov eax, var_50
  loc_00E1A9B9: mov [ecx+00000004h], eax
  loc_00E1A9BC: mov eax, var_4C
  loc_00E1A9BF: mov [ecx+00000008h], eax
  loc_00E1A9C2: mov eax, var_48
  loc_00E1A9C5: mov [ecx+0000000Ch], eax
  loc_00E1A9C8: mov ecx, var_D4
  loc_00E1A9CE: push ecx
  loc_00E1A9CF: call [edx+00000038h]
  loc_00E1A9D2: test eax, eax
  loc_00E1A9D4: fnclex
  loc_00E1A9D6: jge 00E1A9E9h
  loc_00E1A9D8: mov edx, var_D4
  loc_00E1A9DE: push 00000038h
  loc_00E1A9E0: push 006E09F8h
  loc_00E1A9E5: push edx
  loc_00E1A9E6: push eax
  loc_00E1A9E7: call ebx
  loc_00E1A9E9: lea eax, var_34
  loc_00E1A9EC: lea ecx, var_30
  loc_00E1A9EF: push eax
  loc_00E1A9F0: lea edx, var_2C
  loc_00E1A9F3: push ecx
  loc_00E1A9F4: lea eax, var_28
  loc_00E1A9F7: push edx
  loc_00E1A9F8: lea ecx, var_24
  loc_00E1A9FB: push eax
  loc_00E1A9FC: push ecx
  loc_00E1A9FD: push 00000005h
  loc_00E1A9FF: call [00401048h] ; __vbaFreeObjList
  loc_00E1AA05: lea edx, var_54
  loc_00E1AA08: lea eax, var_44
  loc_00E1AA0B: push edx
  loc_00E1AA0C: push eax
  loc_00E1AA0D: push 00000002h
  loc_00E1AA0F: call [00401038h] ; __vbaFreeVarList
  loc_00E1AA15: mov ecx, [esi]
  loc_00E1AA17: add esp, 00000024h
  loc_00E1AA1A: push 006DCBD8h
  loc_00E1AA1F: push 00000000h
  loc_00E1AA21: push 00000018h
  loc_00E1AA23: push esi
  loc_00E1AA24: call [ecx+00000564h]
  loc_00E1AA2A: lea edx, var_28
  loc_00E1AA2D: push eax
  loc_00E1AA2E: push edx
  loc_00E1AA2F: call edi
  loc_00E1AA31: push eax
  loc_00E1AA32: lea eax, var_44
  loc_00E1AA35: push eax
  loc_00E1AA36: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1AA3C: add esp, 00000010h
  loc_00E1AA3F: push eax
  loc_00E1AA40: call [00401120h] ; __vbaCastObjVar
  loc_00E1AA46: lea ecx, var_2C
  loc_00E1AA49: push eax
  loc_00E1AA4A: push ecx
  loc_00E1AA4B: call edi
  loc_00E1AA4D: mov edx, [eax]
  loc_00E1AA4F: lea ecx, var_30
  loc_00E1AA52: push ecx
  loc_00E1AA53: push eax
  loc_00E1AA54: mov var_C4, eax
  loc_00E1AA5A: call [edx+00000054h]
  loc_00E1AA5D: test eax, eax
  loc_00E1AA5F: fnclex
  loc_00E1AA61: jge 00E1AA74h
  loc_00E1AA63: mov edx, var_C4
  loc_00E1AA69: push 00000054h
  loc_00E1AA6B: push 006DCBE8h
  loc_00E1AA70: push edx
  loc_00E1AA71: push eax
  loc_00E1AA72: call ebx
  loc_00E1AA74: lea edx, var_34
  loc_00E1AA77: mov eax, 00000002h
  loc_00E1AA7C: push edx
  loc_00E1AA7D: mov ecx, var_30
  loc_00E1AA80: sub esp, 00000010h
  loc_00E1AA83: mov var_84, eax
  loc_00E1AA89: mov edx, esp
  loc_00E1AA8B: mov var_7C, 00000001h
  loc_00E1AA92: mov var_CC, ecx
  loc_00E1AA98: mov ecx, [ecx]
  loc_00E1AA9A: mov [edx], eax
  loc_00E1AA9C: mov eax, var_80
  loc_00E1AA9F: mov [edx+00000004h], eax
  loc_00E1AAA2: mov eax, var_7C
  loc_00E1AAA5: mov [edx+00000008h], eax
  loc_00E1AAA8: mov eax, var_78
  loc_00E1AAAB: mov [edx+0000000Ch], eax
  loc_00E1AAAE: mov edx, var_30
  loc_00E1AAB1: push edx
  loc_00E1AAB2: call [ecx+00000028h]
  loc_00E1AAB5: test eax, eax
  loc_00E1AAB7: fnclex
  loc_00E1AAB9: jge 00E1AACCh
  loc_00E1AABB: mov ecx, var_CC
  loc_00E1AAC1: push 00000028h
  loc_00E1AAC3: push 006E09E8h
  loc_00E1AAC8: push ecx
  loc_00E1AAC9: push eax
  loc_00E1AACA: call ebx
  loc_00E1AACC: mov edx, var_34
  loc_00E1AACF: mov eax, [esi]
  loc_00E1AAD1: push esi
  loc_00E1AAD2: mov var_D4, edx
  loc_00E1AAD8: call [eax+0000051Ch]
  loc_00E1AADE: lea ecx, var_24
  loc_00E1AAE1: push eax
  loc_00E1AAE2: push ecx
  loc_00E1AAE3: call edi
  loc_00E1AAE5: mov edx, [eax]
  loc_00E1AAE7: lea ecx, var_18
  loc_00E1AAEA: push ecx
  loc_00E1AAEB: push eax
  loc_00E1AAEC: mov var_BC, eax
  loc_00E1AAF2: call [edx+00000050h]
  loc_00E1AAF5: test eax, eax
  loc_00E1AAF7: fnclex
  loc_00E1AAF9: jge 00E1AB0Ch
  loc_00E1AAFB: mov edx, var_BC
  loc_00E1AB01: push 00000050h
  loc_00E1AB03: push 006E0410h
  loc_00E1AB08: push edx
  loc_00E1AB09: push eax
  loc_00E1AB0A: call ebx
  loc_00E1AB0C: mov eax, var_18
  loc_00E1AB0F: mov ecx, var_D4
  loc_00E1AB15: mov var_4C, eax
  loc_00E1AB18: mov eax, 00000008h
  loc_00E1AB1D: mov var_18, 00000000h
  loc_00E1AB24: mov var_54, eax
  loc_00E1AB27: mov edx, [ecx]
  loc_00E1AB29: sub esp, 00000010h
  loc_00E1AB2C: mov ecx, esp
  loc_00E1AB2E: mov [ecx], eax
  loc_00E1AB30: mov eax, var_50
  loc_00E1AB33: mov [ecx+00000004h], eax
  loc_00E1AB36: mov eax, var_4C
  loc_00E1AB39: mov [ecx+00000008h], eax
  loc_00E1AB3C: mov eax, var_48
  loc_00E1AB3F: mov [ecx+0000000Ch], eax
  loc_00E1AB42: mov ecx, var_D4
  loc_00E1AB48: push ecx
  loc_00E1AB49: call [edx+00000038h]
  loc_00E1AB4C: test eax, eax
  loc_00E1AB4E: fnclex
  loc_00E1AB50: jge 00E1AB63h
  loc_00E1AB52: mov edx, var_D4
  loc_00E1AB58: push 00000038h
  loc_00E1AB5A: push 006E09F8h
  loc_00E1AB5F: push edx
  loc_00E1AB60: push eax
  loc_00E1AB61: call ebx
  loc_00E1AB63: lea eax, var_34
  loc_00E1AB66: lea ecx, var_30
  loc_00E1AB69: push eax
  loc_00E1AB6A: lea edx, var_2C
  loc_00E1AB6D: push ecx
  loc_00E1AB6E: lea eax, var_28
  loc_00E1AB71: push edx
  loc_00E1AB72: lea ecx, var_24
  loc_00E1AB75: push eax
  loc_00E1AB76: push ecx
  loc_00E1AB77: push 00000005h
  loc_00E1AB79: call [00401048h] ; __vbaFreeObjList
  loc_00E1AB7F: lea edx, var_54
  loc_00E1AB82: lea eax, var_44
  loc_00E1AB85: push edx
  loc_00E1AB86: push eax
  loc_00E1AB87: push 00000002h
  loc_00E1AB89: call [00401038h] ; __vbaFreeVarList
  loc_00E1AB8F: mov ecx, [esi]
  loc_00E1AB91: add esp, 00000024h
  loc_00E1AB94: push 006DCBD8h
  loc_00E1AB99: push 00000000h
  loc_00E1AB9B: push 00000018h
  loc_00E1AB9D: push esi
  loc_00E1AB9E: call [ecx+00000564h]
  loc_00E1ABA4: lea edx, var_28
  loc_00E1ABA7: push eax
  loc_00E1ABA8: push edx
  loc_00E1ABA9: call edi
  loc_00E1ABAB: push eax
  loc_00E1ABAC: lea eax, var_44
  loc_00E1ABAF: push eax
  loc_00E1ABB0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1ABB6: add esp, 00000010h
  loc_00E1ABB9: push eax
  loc_00E1ABBA: call [00401120h] ; __vbaCastObjVar
  loc_00E1ABC0: lea ecx, var_2C
  loc_00E1ABC3: push eax
  loc_00E1ABC4: push ecx
  loc_00E1ABC5: call edi
  loc_00E1ABC7: mov edx, [eax]
  loc_00E1ABC9: lea ecx, var_30
  loc_00E1ABCC: push ecx
  loc_00E1ABCD: push eax
  loc_00E1ABCE: mov var_C4, eax
  loc_00E1ABD4: call [edx+00000054h]
  loc_00E1ABD7: test eax, eax
  loc_00E1ABD9: fnclex
  loc_00E1ABDB: jge 00E1ABEEh
  loc_00E1ABDD: mov edx, var_C4
  loc_00E1ABE3: push 00000054h
  loc_00E1ABE5: push 006DCBE8h
  loc_00E1ABEA: push edx
  loc_00E1ABEB: push eax
  loc_00E1ABEC: call ebx
  loc_00E1ABEE: lea edx, var_34
  loc_00E1ABF1: mov eax, 00000002h
  loc_00E1ABF6: push edx
  loc_00E1ABF7: mov ecx, var_30
  loc_00E1ABFA: sub esp, 00000010h
  loc_00E1ABFD: mov var_7C, eax
  loc_00E1AC00: mov edx, esp
  loc_00E1AC02: mov var_84, eax
  loc_00E1AC08: mov var_CC, ecx
  loc_00E1AC0E: mov ecx, [ecx]
  loc_00E1AC10: mov [edx], eax
  loc_00E1AC12: mov eax, var_80
  loc_00E1AC15: mov [edx+00000004h], eax
  loc_00E1AC18: mov eax, var_7C
  loc_00E1AC1B: mov [edx+00000008h], eax
  loc_00E1AC1E: mov eax, var_78
  loc_00E1AC21: mov [edx+0000000Ch], eax
  loc_00E1AC24: mov edx, var_30
  loc_00E1AC27: push edx
  loc_00E1AC28: call [ecx+00000028h]
  loc_00E1AC2B: test eax, eax
  loc_00E1AC2D: fnclex
  loc_00E1AC2F: jge 00E1AC42h
  loc_00E1AC31: mov ecx, var_CC
  loc_00E1AC37: push 00000028h
  loc_00E1AC39: push 006E09E8h
  loc_00E1AC3E: push ecx
  loc_00E1AC3F: push eax
  loc_00E1AC40: call ebx
  loc_00E1AC42: mov edx, var_34
  loc_00E1AC45: mov eax, [esi]
  loc_00E1AC47: push esi
  loc_00E1AC48: mov var_D4, edx
  loc_00E1AC4E: call [eax+00000518h]
  loc_00E1AC54: lea ecx, var_24
  loc_00E1AC57: push eax
  loc_00E1AC58: push ecx
  loc_00E1AC59: call edi
  loc_00E1AC5B: mov edx, [eax]
  loc_00E1AC5D: lea ecx, var_18
  loc_00E1AC60: push ecx
  loc_00E1AC61: push eax
  loc_00E1AC62: mov var_BC, eax
  loc_00E1AC68: call [edx+00000050h]
  loc_00E1AC6B: test eax, eax
  loc_00E1AC6D: fnclex
  loc_00E1AC6F: jge 00E1AC82h
  loc_00E1AC71: mov edx, var_BC
  loc_00E1AC77: push 00000050h
  loc_00E1AC79: push 006E0410h
  loc_00E1AC7E: push edx
  loc_00E1AC7F: push eax
  loc_00E1AC80: call ebx
  loc_00E1AC82: mov eax, var_18
  loc_00E1AC85: mov ecx, var_D4
  loc_00E1AC8B: mov var_4C, eax
  loc_00E1AC8E: mov eax, 00000008h
  loc_00E1AC93: mov var_18, 00000000h
  loc_00E1AC9A: mov var_54, eax
  loc_00E1AC9D: mov edx, [ecx]
  loc_00E1AC9F: sub esp, 00000010h
  loc_00E1ACA2: mov ecx, esp
  loc_00E1ACA4: mov [ecx], eax
  loc_00E1ACA6: mov eax, var_50
  loc_00E1ACA9: mov [ecx+00000004h], eax
  loc_00E1ACAC: mov eax, var_4C
  loc_00E1ACAF: mov [ecx+00000008h], eax
  loc_00E1ACB2: mov eax, var_48
  loc_00E1ACB5: mov [ecx+0000000Ch], eax
  loc_00E1ACB8: mov ecx, var_D4
  loc_00E1ACBE: push ecx
  loc_00E1ACBF: call [edx+00000038h]
  loc_00E1ACC2: test eax, eax
  loc_00E1ACC4: fnclex
  loc_00E1ACC6: jge 00E1ACD9h
  loc_00E1ACC8: mov edx, var_D4
  loc_00E1ACCE: push 00000038h
  loc_00E1ACD0: push 006E09F8h
  loc_00E1ACD5: push edx
  loc_00E1ACD6: push eax
  loc_00E1ACD7: call ebx
  loc_00E1ACD9: lea eax, var_34
  loc_00E1ACDC: lea ecx, var_30
  loc_00E1ACDF: push eax
  loc_00E1ACE0: lea edx, var_2C
  loc_00E1ACE3: push ecx
  loc_00E1ACE4: lea eax, var_28
  loc_00E1ACE7: push edx
  loc_00E1ACE8: lea ecx, var_24
  loc_00E1ACEB: push eax
  loc_00E1ACEC: push ecx
  loc_00E1ACED: push 00000005h
  loc_00E1ACEF: call [00401048h] ; __vbaFreeObjList
  loc_00E1ACF5: lea edx, var_54
  loc_00E1ACF8: lea eax, var_44
  loc_00E1ACFB: push edx
  loc_00E1ACFC: push eax
  loc_00E1ACFD: push 00000002h
  loc_00E1ACFF: call [00401038h] ; __vbaFreeVarList
  loc_00E1AD05: mov ecx, [esi]
  loc_00E1AD07: add esp, 00000024h
  loc_00E1AD0A: push 006DCBD8h
  loc_00E1AD0F: push 00000000h
  loc_00E1AD11: push 00000018h
  loc_00E1AD13: push esi
  loc_00E1AD14: call [ecx+00000564h]
  loc_00E1AD1A: lea edx, var_28
  loc_00E1AD1D: push eax
  loc_00E1AD1E: push edx
  loc_00E1AD1F: call edi
  loc_00E1AD21: push eax
  loc_00E1AD22: lea eax, var_44
  loc_00E1AD25: push eax
  loc_00E1AD26: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1AD2C: add esp, 00000010h
  loc_00E1AD2F: push eax
  loc_00E1AD30: call [00401120h] ; __vbaCastObjVar
  loc_00E1AD36: lea ecx, var_2C
  loc_00E1AD39: push eax
  loc_00E1AD3A: push ecx
  loc_00E1AD3B: call edi
  loc_00E1AD3D: mov edx, [eax]
  loc_00E1AD3F: lea ecx, var_30
  loc_00E1AD42: push ecx
  loc_00E1AD43: push eax
  loc_00E1AD44: mov var_C4, eax
  loc_00E1AD4A: call [edx+00000054h]
  loc_00E1AD4D: test eax, eax
  loc_00E1AD4F: fnclex
  loc_00E1AD51: jge 00E1AD64h
  loc_00E1AD53: mov edx, var_C4
  loc_00E1AD59: push 00000054h
  loc_00E1AD5B: push 006DCBE8h
  loc_00E1AD60: push edx
  loc_00E1AD61: push eax
  loc_00E1AD62: call ebx
  loc_00E1AD64: lea edx, var_34
  loc_00E1AD67: mov eax, 00000002h
  loc_00E1AD6C: push edx
  loc_00E1AD6D: mov ecx, var_30
  loc_00E1AD70: sub esp, 00000010h
  loc_00E1AD73: mov var_84, eax
  loc_00E1AD79: mov edx, esp
  loc_00E1AD7B: mov var_7C, 00000003h
  loc_00E1AD82: mov var_CC, ecx
  loc_00E1AD88: mov ecx, [ecx]
  loc_00E1AD8A: mov [edx], eax
  loc_00E1AD8C: mov eax, var_80
  loc_00E1AD8F: mov [edx+00000004h], eax
  loc_00E1AD92: mov eax, var_7C
  loc_00E1AD95: mov [edx+00000008h], eax
  loc_00E1AD98: mov eax, var_78
  loc_00E1AD9B: mov [edx+0000000Ch], eax
  loc_00E1AD9E: mov edx, var_30
  loc_00E1ADA1: push edx
  loc_00E1ADA2: call [ecx+00000028h]
  loc_00E1ADA5: test eax, eax
  loc_00E1ADA7: fnclex
  loc_00E1ADA9: jge 00E1ADBCh
  loc_00E1ADAB: mov ecx, var_CC
  loc_00E1ADB1: push 00000028h
  loc_00E1ADB3: push 006E09E8h
  loc_00E1ADB8: push ecx
  loc_00E1ADB9: push eax
  loc_00E1ADBA: call ebx
  loc_00E1ADBC: mov edx, var_34
  loc_00E1ADBF: mov eax, [esi]
  loc_00E1ADC1: push esi
  loc_00E1ADC2: mov var_D4, edx
  loc_00E1ADC8: call [eax+000004E4h]
  loc_00E1ADCE: lea ecx, var_24
  loc_00E1ADD1: push eax
  loc_00E1ADD2: push ecx
  loc_00E1ADD3: call edi
  loc_00E1ADD5: mov edx, [eax]
  loc_00E1ADD7: lea ecx, var_18
  loc_00E1ADDA: push ecx
  loc_00E1ADDB: push eax
  loc_00E1ADDC: mov var_BC, eax
  loc_00E1ADE2: call [edx+00000050h]
  loc_00E1ADE5: test eax, eax
  loc_00E1ADE7: fnclex
  loc_00E1ADE9: jge 00E1ADFCh
  loc_00E1ADEB: mov edx, var_BC
  loc_00E1ADF1: push 00000050h
  loc_00E1ADF3: push 006E0410h
  loc_00E1ADF8: push edx
  loc_00E1ADF9: push eax
  loc_00E1ADFA: call ebx
  loc_00E1ADFC: mov eax, var_18
  loc_00E1ADFF: mov ecx, var_D4
  loc_00E1AE05: mov var_4C, eax
  loc_00E1AE08: mov eax, 00000008h
  loc_00E1AE0D: mov var_18, 00000000h
  loc_00E1AE14: mov var_54, eax
  loc_00E1AE17: mov edx, [ecx]
  loc_00E1AE19: sub esp, 00000010h
  loc_00E1AE1C: mov ecx, esp
  loc_00E1AE1E: mov [ecx], eax
  loc_00E1AE20: mov eax, var_50
  loc_00E1AE23: mov [ecx+00000004h], eax
  loc_00E1AE26: mov eax, var_4C
  loc_00E1AE29: mov [ecx+00000008h], eax
  loc_00E1AE2C: mov eax, var_48
  loc_00E1AE2F: mov [ecx+0000000Ch], eax
  loc_00E1AE32: mov ecx, var_D4
  loc_00E1AE38: push ecx
  loc_00E1AE39: call [edx+00000038h]
  loc_00E1AE3C: test eax, eax
  loc_00E1AE3E: fnclex
  loc_00E1AE40: jge 00E1AE53h
  loc_00E1AE42: mov edx, var_D4
  loc_00E1AE48: push 00000038h
  loc_00E1AE4A: push 006E09F8h
  loc_00E1AE4F: push edx
  loc_00E1AE50: push eax
  loc_00E1AE51: call ebx
  loc_00E1AE53: lea eax, var_34
  loc_00E1AE56: lea ecx, var_30
  loc_00E1AE59: push eax
  loc_00E1AE5A: lea edx, var_2C
  loc_00E1AE5D: push ecx
  loc_00E1AE5E: lea eax, var_28
  loc_00E1AE61: push edx
  loc_00E1AE62: lea ecx, var_24
  loc_00E1AE65: push eax
  loc_00E1AE66: push ecx
  loc_00E1AE67: push 00000005h
  loc_00E1AE69: call [00401048h] ; __vbaFreeObjList
  loc_00E1AE6F: lea edx, var_54
  loc_00E1AE72: lea eax, var_44
  loc_00E1AE75: push edx
  loc_00E1AE76: push eax
  loc_00E1AE77: push 00000002h
  loc_00E1AE79: call [00401038h] ; __vbaFreeVarList
  loc_00E1AE7F: mov ecx, [esi]
  loc_00E1AE81: add esp, 00000024h
  loc_00E1AE84: push 006DCBD8h
  loc_00E1AE89: push 00000000h
  loc_00E1AE8B: push 00000018h
  loc_00E1AE8D: push esi
  loc_00E1AE8E: call [ecx+00000564h]
  loc_00E1AE94: lea edx, var_28
  loc_00E1AE97: push eax
  loc_00E1AE98: push edx
  loc_00E1AE99: call edi
  loc_00E1AE9B: push eax
  loc_00E1AE9C: lea eax, var_44
  loc_00E1AE9F: push eax
  loc_00E1AEA0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1AEA6: add esp, 00000010h
  loc_00E1AEA9: push eax
  loc_00E1AEAA: call [00401120h] ; __vbaCastObjVar
  loc_00E1AEB0: lea ecx, var_2C
  loc_00E1AEB3: push eax
  loc_00E1AEB4: push ecx
  loc_00E1AEB5: call edi
  loc_00E1AEB7: mov edx, [eax]
  loc_00E1AEB9: lea ecx, var_30
  loc_00E1AEBC: push ecx
  loc_00E1AEBD: push eax
  loc_00E1AEBE: mov var_C4, eax
  loc_00E1AEC4: call [edx+00000054h]
  loc_00E1AEC7: test eax, eax
  loc_00E1AEC9: fnclex
  loc_00E1AECB: jge 00E1AEDEh
  loc_00E1AECD: mov edx, var_C4
  loc_00E1AED3: push 00000054h
  loc_00E1AED5: push 006DCBE8h
  loc_00E1AEDA: push edx
  loc_00E1AEDB: push eax
  loc_00E1AEDC: call ebx
  loc_00E1AEDE: lea edx, var_34
  loc_00E1AEE1: mov eax, 00000002h
  loc_00E1AEE6: push edx
  loc_00E1AEE7: mov ecx, var_30
  loc_00E1AEEA: sub esp, 00000010h
  loc_00E1AEED: mov var_84, eax
  loc_00E1AEF3: mov edx, esp
  loc_00E1AEF5: mov var_7C, 00000004h
  loc_00E1AEFC: mov var_CC, ecx
  loc_00E1AF02: mov ecx, [ecx]
  loc_00E1AF04: mov [edx], eax
  loc_00E1AF06: mov eax, var_80
  loc_00E1AF09: mov [edx+00000004h], eax
  loc_00E1AF0C: mov eax, var_7C
  loc_00E1AF0F: mov [edx+00000008h], eax
  loc_00E1AF12: mov eax, var_78
  loc_00E1AF15: mov [edx+0000000Ch], eax
  loc_00E1AF18: mov edx, var_30
  loc_00E1AF1B: push edx
  loc_00E1AF1C: call [ecx+00000028h]
  loc_00E1AF1F: test eax, eax
  loc_00E1AF21: fnclex
  loc_00E1AF23: jge 00E1AF36h
  loc_00E1AF25: mov ecx, var_CC
  loc_00E1AF2B: push 00000028h
  loc_00E1AF2D: push 006E09E8h
  loc_00E1AF32: push ecx
  loc_00E1AF33: push eax
  loc_00E1AF34: call ebx
  loc_00E1AF36: mov edx, var_34
  loc_00E1AF39: mov eax, [esi]
  loc_00E1AF3B: push esi
  loc_00E1AF3C: mov var_D4, edx
  loc_00E1AF42: call [eax+000004E0h]
  loc_00E1AF48: lea ecx, var_24
  loc_00E1AF4B: push eax
  loc_00E1AF4C: push ecx
  loc_00E1AF4D: call edi
  loc_00E1AF4F: mov edx, [eax]
  loc_00E1AF51: lea ecx, var_18
  loc_00E1AF54: push ecx
  loc_00E1AF55: push eax
  loc_00E1AF56: mov var_BC, eax
  loc_00E1AF5C: call [edx+00000050h]
  loc_00E1AF5F: test eax, eax
  loc_00E1AF61: fnclex
  loc_00E1AF63: jge 00E1AF76h
  loc_00E1AF65: mov edx, var_BC
  loc_00E1AF6B: push 00000050h
  loc_00E1AF6D: push 006E0410h
  loc_00E1AF72: push edx
  loc_00E1AF73: push eax
  loc_00E1AF74: call ebx
  loc_00E1AF76: mov eax, var_18
  loc_00E1AF79: mov ecx, var_D4
  loc_00E1AF7F: mov var_4C, eax
  loc_00E1AF82: mov eax, 00000008h
  loc_00E1AF87: mov var_18, 00000000h
  loc_00E1AF8E: mov var_54, eax
  loc_00E1AF91: mov edx, [ecx]
  loc_00E1AF93: sub esp, 00000010h
  loc_00E1AF96: mov ecx, esp
  loc_00E1AF98: mov [ecx], eax
  loc_00E1AF9A: mov eax, var_50
  loc_00E1AF9D: mov [ecx+00000004h], eax
  loc_00E1AFA0: mov eax, var_4C
  loc_00E1AFA3: mov [ecx+00000008h], eax
  loc_00E1AFA6: mov eax, var_48
  loc_00E1AFA9: mov [ecx+0000000Ch], eax
  loc_00E1AFAC: mov ecx, var_D4
  loc_00E1AFB2: push ecx
  loc_00E1AFB3: call [edx+00000038h]
  loc_00E1AFB6: test eax, eax
  loc_00E1AFB8: fnclex
  loc_00E1AFBA: jge 00E1AFCDh
  loc_00E1AFBC: mov edx, var_D4
  loc_00E1AFC2: push 00000038h
  loc_00E1AFC4: push 006E09F8h
  loc_00E1AFC9: push edx
  loc_00E1AFCA: push eax
  loc_00E1AFCB: call ebx
  loc_00E1AFCD: lea eax, var_34
  loc_00E1AFD0: lea ecx, var_30
  loc_00E1AFD3: push eax
  loc_00E1AFD4: lea edx, var_2C
  loc_00E1AFD7: push ecx
  loc_00E1AFD8: lea eax, var_28
  loc_00E1AFDB: push edx
  loc_00E1AFDC: lea ecx, var_24
  loc_00E1AFDF: push eax
  loc_00E1AFE0: push ecx
  loc_00E1AFE1: push 00000005h
  loc_00E1AFE3: call [00401048h] ; __vbaFreeObjList
  loc_00E1AFE9: lea edx, var_54
  loc_00E1AFEC: lea eax, var_44
  loc_00E1AFEF: push edx
  loc_00E1AFF0: push eax
  loc_00E1AFF1: push 00000002h
  loc_00E1AFF3: call [00401038h] ; __vbaFreeVarList
  loc_00E1AFF9: mov ecx, [esi]
  loc_00E1AFFB: add esp, 00000024h
  loc_00E1AFFE: push 006DCBD8h
  loc_00E1B003: push 00000000h
  loc_00E1B005: push 00000018h
  loc_00E1B007: push esi
  loc_00E1B008: call [ecx+00000564h]
  loc_00E1B00E: lea edx, var_28
  loc_00E1B011: push eax
  loc_00E1B012: push edx
  loc_00E1B013: call edi
  loc_00E1B015: push eax
  loc_00E1B016: lea eax, var_44
  loc_00E1B019: push eax
  loc_00E1B01A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1B020: add esp, 00000010h
  loc_00E1B023: push eax
  loc_00E1B024: call [00401120h] ; __vbaCastObjVar
  loc_00E1B02A: lea ecx, var_2C
  loc_00E1B02D: push eax
  loc_00E1B02E: push ecx
  loc_00E1B02F: call edi
  loc_00E1B031: mov edx, [eax]
  loc_00E1B033: lea ecx, var_30
  loc_00E1B036: push ecx
  loc_00E1B037: push eax
  loc_00E1B038: mov var_C4, eax
  loc_00E1B03E: call [edx+00000054h]
  loc_00E1B041: test eax, eax
  loc_00E1B043: fnclex
  loc_00E1B045: jge 00E1B058h
  loc_00E1B047: mov edx, var_C4
  loc_00E1B04D: push 00000054h
  loc_00E1B04F: push 006DCBE8h
  loc_00E1B054: push edx
  loc_00E1B055: push eax
  loc_00E1B056: call ebx
  loc_00E1B058: lea edx, var_34
  loc_00E1B05B: mov eax, 00000002h
  loc_00E1B060: push edx
  loc_00E1B061: mov ecx, var_30
  loc_00E1B064: sub esp, 00000010h
  loc_00E1B067: mov var_84, eax
  loc_00E1B06D: mov edx, esp
  loc_00E1B06F: mov var_7C, 00000005h
  loc_00E1B076: mov var_CC, ecx
  loc_00E1B07C: mov ecx, [ecx]
  loc_00E1B07E: mov [edx], eax
  loc_00E1B080: mov eax, var_80
  loc_00E1B083: mov [edx+00000004h], eax
  loc_00E1B086: mov eax, var_7C
  loc_00E1B089: mov [edx+00000008h], eax
  loc_00E1B08C: mov eax, var_78
  loc_00E1B08F: mov [edx+0000000Ch], eax
  loc_00E1B092: mov edx, var_30
  loc_00E1B095: push edx
  loc_00E1B096: call [ecx+00000028h]
  loc_00E1B099: test eax, eax
  loc_00E1B09B: fnclex
  loc_00E1B09D: jge 00E1B0B0h
  loc_00E1B09F: mov ecx, var_CC
  loc_00E1B0A5: push 00000028h
  loc_00E1B0A7: push 006E09E8h
  loc_00E1B0AC: push ecx
  loc_00E1B0AD: push eax
  loc_00E1B0AE: call ebx
  loc_00E1B0B0: mov edx, var_34
  loc_00E1B0B3: mov eax, [esi]
  loc_00E1B0B5: push esi
  loc_00E1B0B6: mov var_D4, edx
  loc_00E1B0BC: call [eax+00000474h]
  loc_00E1B0C2: lea ecx, var_24
  loc_00E1B0C5: push eax
  loc_00E1B0C6: push ecx
  loc_00E1B0C7: call edi
  loc_00E1B0C9: mov edx, [eax]
  loc_00E1B0CB: lea ecx, var_18
  loc_00E1B0CE: push ecx
  loc_00E1B0CF: push eax
  loc_00E1B0D0: mov var_BC, eax
  loc_00E1B0D6: call [edx+00000050h]
  loc_00E1B0D9: test eax, eax
  loc_00E1B0DB: fnclex
  loc_00E1B0DD: jge 00E1B0F0h
  loc_00E1B0DF: mov edx, var_BC
  loc_00E1B0E5: push 00000050h
  loc_00E1B0E7: push 006E0410h
  loc_00E1B0EC: push edx
  loc_00E1B0ED: push eax
  loc_00E1B0EE: call ebx
  loc_00E1B0F0: mov eax, var_18
  loc_00E1B0F3: mov ecx, var_D4
  loc_00E1B0F9: mov var_4C, eax
  loc_00E1B0FC: mov eax, 00000008h
  loc_00E1B101: mov var_18, 00000000h
  loc_00E1B108: mov var_54, eax
  loc_00E1B10B: mov edx, [ecx]
  loc_00E1B10D: sub esp, 00000010h
  loc_00E1B110: mov ecx, esp
  loc_00E1B112: mov [ecx], eax
  loc_00E1B114: mov eax, var_50
  loc_00E1B117: mov [ecx+00000004h], eax
  loc_00E1B11A: mov eax, var_4C
  loc_00E1B11D: mov [ecx+00000008h], eax
  loc_00E1B120: mov eax, var_48
  loc_00E1B123: mov [ecx+0000000Ch], eax
  loc_00E1B126: mov ecx, var_D4
  loc_00E1B12C: push ecx
  loc_00E1B12D: call [edx+00000038h]
  loc_00E1B130: test eax, eax
  loc_00E1B132: fnclex
  loc_00E1B134: jge 00E1B147h
  loc_00E1B136: mov edx, var_D4
  loc_00E1B13C: push 00000038h
  loc_00E1B13E: push 006E09F8h
  loc_00E1B143: push edx
  loc_00E1B144: push eax
  loc_00E1B145: call ebx
  loc_00E1B147: lea eax, var_34
  loc_00E1B14A: lea ecx, var_30
  loc_00E1B14D: push eax
  loc_00E1B14E: lea edx, var_2C
  loc_00E1B151: push ecx
  loc_00E1B152: lea eax, var_28
  loc_00E1B155: push edx
  loc_00E1B156: lea ecx, var_24
  loc_00E1B159: push eax
  loc_00E1B15A: push ecx
  loc_00E1B15B: push 00000005h
  loc_00E1B15D: call [00401048h] ; __vbaFreeObjList
  loc_00E1B163: lea edx, var_54
  loc_00E1B166: lea eax, var_44
  loc_00E1B169: push edx
  loc_00E1B16A: push eax
  loc_00E1B16B: push 00000002h
  loc_00E1B16D: call [00401038h] ; __vbaFreeVarList
  loc_00E1B173: mov ecx, [esi]
  loc_00E1B175: add esp, 00000024h
  loc_00E1B178: push 006DCBD8h
  loc_00E1B17D: push 00000000h
  loc_00E1B17F: push 00000018h
  loc_00E1B181: push esi
  loc_00E1B182: call [ecx+00000564h]
  loc_00E1B188: lea edx, var_28
  loc_00E1B18B: push eax
  loc_00E1B18C: push edx
  loc_00E1B18D: call edi
  loc_00E1B18F: push eax
  loc_00E1B190: lea eax, var_44
  loc_00E1B193: push eax
  loc_00E1B194: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1B19A: add esp, 00000010h
  loc_00E1B19D: push eax
  loc_00E1B19E: call [00401120h] ; __vbaCastObjVar
  loc_00E1B1A4: lea ecx, var_2C
  loc_00E1B1A7: push eax
  loc_00E1B1A8: push ecx
  loc_00E1B1A9: call edi
  loc_00E1B1AB: mov edx, [eax]
  loc_00E1B1AD: lea ecx, var_30
  loc_00E1B1B0: push ecx
  loc_00E1B1B1: push eax
  loc_00E1B1B2: mov var_C4, eax
  loc_00E1B1B8: call [edx+00000054h]
  loc_00E1B1BB: test eax, eax
  loc_00E1B1BD: fnclex
  loc_00E1B1BF: jge 00E1B1D2h
  loc_00E1B1C1: mov edx, var_C4
  loc_00E1B1C7: push 00000054h
  loc_00E1B1C9: push 006DCBE8h
  loc_00E1B1CE: push edx
  loc_00E1B1CF: push eax
  loc_00E1B1D0: call ebx
  loc_00E1B1D2: lea edx, var_34
  loc_00E1B1D5: mov eax, 00000002h
  loc_00E1B1DA: push edx
  loc_00E1B1DB: mov ecx, var_30
  loc_00E1B1DE: sub esp, 00000010h
  loc_00E1B1E1: mov var_84, eax
  loc_00E1B1E7: mov edx, esp
  loc_00E1B1E9: mov var_7C, 00000006h
  loc_00E1B1F0: mov var_CC, ecx
  loc_00E1B1F6: mov ecx, [ecx]
  loc_00E1B1F8: mov [edx], eax
  loc_00E1B1FA: mov eax, var_80
  loc_00E1B1FD: mov [edx+00000004h], eax
  loc_00E1B200: mov eax, var_7C
  loc_00E1B203: mov [edx+00000008h], eax
  loc_00E1B206: mov eax, var_78
  loc_00E1B209: mov [edx+0000000Ch], eax
  loc_00E1B20C: mov edx, var_30
  loc_00E1B20F: push edx
  loc_00E1B210: call [ecx+00000028h]
  loc_00E1B213: test eax, eax
  loc_00E1B215: fnclex
  loc_00E1B217: jge 00E1B22Ah
  loc_00E1B219: mov ecx, var_CC
  loc_00E1B21F: push 00000028h
  loc_00E1B221: push 006E09E8h
  loc_00E1B226: push ecx
  loc_00E1B227: push eax
  loc_00E1B228: call ebx
  loc_00E1B22A: mov edx, var_34
  loc_00E1B22D: mov eax, [esi]
  loc_00E1B22F: push esi
  loc_00E1B230: mov var_D4, edx
  loc_00E1B236: call [eax+00000470h]
  loc_00E1B23C: lea ecx, var_24
  loc_00E1B23F: push eax
  loc_00E1B240: push ecx
  loc_00E1B241: call edi
  loc_00E1B243: mov edx, [eax]
  loc_00E1B245: lea ecx, var_18
  loc_00E1B248: push ecx
  loc_00E1B249: push eax
  loc_00E1B24A: mov var_BC, eax
  loc_00E1B250: call [edx+00000050h]
  loc_00E1B253: test eax, eax
  loc_00E1B255: fnclex
  loc_00E1B257: jge 00E1B26Ah
  loc_00E1B259: mov edx, var_BC
  loc_00E1B25F: push 00000050h
  loc_00E1B261: push 006E0410h
  loc_00E1B266: push edx
  loc_00E1B267: push eax
  loc_00E1B268: call ebx
  loc_00E1B26A: mov eax, var_18
  loc_00E1B26D: mov ecx, var_D4
  loc_00E1B273: mov var_4C, eax
  loc_00E1B276: mov eax, 00000008h
  loc_00E1B27B: mov var_18, 00000000h
  loc_00E1B282: mov var_54, eax
  loc_00E1B285: mov edx, [ecx]
  loc_00E1B287: sub esp, 00000010h
  loc_00E1B28A: mov ecx, esp
  loc_00E1B28C: mov [ecx], eax
  loc_00E1B28E: mov eax, var_50
  loc_00E1B291: mov [ecx+00000004h], eax
  loc_00E1B294: mov eax, var_4C
  loc_00E1B297: mov [ecx+00000008h], eax
  loc_00E1B29A: mov eax, var_48
  loc_00E1B29D: mov [ecx+0000000Ch], eax
  loc_00E1B2A0: mov ecx, var_D4
  loc_00E1B2A6: push ecx
  loc_00E1B2A7: call [edx+00000038h]
  loc_00E1B2AA: test eax, eax
  loc_00E1B2AC: fnclex
  loc_00E1B2AE: jge 00E1B2C1h
  loc_00E1B2B0: mov edx, var_D4
  loc_00E1B2B6: push 00000038h
  loc_00E1B2B8: push 006E09F8h
  loc_00E1B2BD: push edx
  loc_00E1B2BE: push eax
  loc_00E1B2BF: call ebx
  loc_00E1B2C1: lea eax, var_34
  loc_00E1B2C4: lea ecx, var_30
  loc_00E1B2C7: push eax
  loc_00E1B2C8: lea edx, var_2C
  loc_00E1B2CB: push ecx
  loc_00E1B2CC: lea eax, var_28
  loc_00E1B2CF: push edx
  loc_00E1B2D0: lea ecx, var_24
  loc_00E1B2D3: push eax
  loc_00E1B2D4: push ecx
  loc_00E1B2D5: push 00000005h
  loc_00E1B2D7: call [00401048h] ; __vbaFreeObjList
  loc_00E1B2DD: lea edx, var_54
  loc_00E1B2E0: lea eax, var_44
  loc_00E1B2E3: push edx
  loc_00E1B2E4: push eax
  loc_00E1B2E5: push 00000002h
  loc_00E1B2E7: call [00401038h] ; __vbaFreeVarList
  loc_00E1B2ED: mov ecx, [esi]
  loc_00E1B2EF: add esp, 00000024h
  loc_00E1B2F2: push 006DCBD8h
  loc_00E1B2F7: push 00000000h
  loc_00E1B2F9: push 00000018h
  loc_00E1B2FB: push esi
  loc_00E1B2FC: call [ecx+00000564h]
  loc_00E1B302: lea edx, var_28
  loc_00E1B305: push eax
  loc_00E1B306: push edx
  loc_00E1B307: call edi
  loc_00E1B309: push eax
  loc_00E1B30A: lea eax, var_44
  loc_00E1B30D: push eax
  loc_00E1B30E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1B314: add esp, 00000010h
  loc_00E1B317: push eax
  loc_00E1B318: call [00401120h] ; __vbaCastObjVar
  loc_00E1B31E: lea ecx, var_2C
  loc_00E1B321: push eax
  loc_00E1B322: push ecx
  loc_00E1B323: call edi
  loc_00E1B325: mov edx, [eax]
  loc_00E1B327: lea ecx, var_30
  loc_00E1B32A: push ecx
  loc_00E1B32B: push eax
  loc_00E1B32C: mov var_C4, eax
  loc_00E1B332: call [edx+00000054h]
  loc_00E1B335: test eax, eax
  loc_00E1B337: fnclex
  loc_00E1B339: jge 00E1B34Ch
  loc_00E1B33B: mov edx, var_C4
  loc_00E1B341: push 00000054h
  loc_00E1B343: push 006DCBE8h
  loc_00E1B348: push edx
  loc_00E1B349: push eax
  loc_00E1B34A: call ebx
  loc_00E1B34C: lea edx, var_34
  loc_00E1B34F: mov eax, 00000002h
  loc_00E1B354: push edx
  loc_00E1B355: mov ecx, var_30
  loc_00E1B358: sub esp, 00000010h
  loc_00E1B35B: mov var_84, eax
  loc_00E1B361: mov edx, esp
  loc_00E1B363: mov var_7C, 00000007h
  loc_00E1B36A: mov var_CC, ecx
  loc_00E1B370: mov ecx, [ecx]
  loc_00E1B372: mov [edx], eax
  loc_00E1B374: mov eax, var_80
  loc_00E1B377: mov [edx+00000004h], eax
  loc_00E1B37A: mov eax, var_7C
  loc_00E1B37D: mov [edx+00000008h], eax
  loc_00E1B380: mov eax, var_78
  loc_00E1B383: mov [edx+0000000Ch], eax
  loc_00E1B386: mov edx, var_30
  loc_00E1B389: push edx
  loc_00E1B38A: call [ecx+00000028h]
  loc_00E1B38D: test eax, eax
  loc_00E1B38F: fnclex
  loc_00E1B391: jge 00E1B3A4h
  loc_00E1B393: mov ecx, var_CC
  loc_00E1B399: push 00000028h
  loc_00E1B39B: push 006E09E8h
  loc_00E1B3A0: push ecx
  loc_00E1B3A1: push eax
  loc_00E1B3A2: call ebx
  loc_00E1B3A4: mov edx, var_34
  loc_00E1B3A7: mov eax, [esi]
  loc_00E1B3A9: push esi
  loc_00E1B3AA: mov var_D4, edx
  loc_00E1B3B0: call [eax+0000046Ch]
  loc_00E1B3B6: lea ecx, var_24
  loc_00E1B3B9: push eax
  loc_00E1B3BA: push ecx
  loc_00E1B3BB: call edi
  loc_00E1B3BD: mov edx, [eax]
  loc_00E1B3BF: lea ecx, var_18
  loc_00E1B3C2: push ecx
  loc_00E1B3C3: push eax
  loc_00E1B3C4: mov var_BC, eax
  loc_00E1B3CA: call [edx+00000050h]
  loc_00E1B3CD: test eax, eax
  loc_00E1B3CF: fnclex
  loc_00E1B3D1: jge 00E1B3E4h
  loc_00E1B3D3: mov edx, var_BC
  loc_00E1B3D9: push 00000050h
  loc_00E1B3DB: push 006E0410h
  loc_00E1B3E0: push edx
  loc_00E1B3E1: push eax
  loc_00E1B3E2: call ebx
  loc_00E1B3E4: mov eax, var_18
  loc_00E1B3E7: mov ecx, var_D4
  loc_00E1B3ED: mov var_4C, eax
  loc_00E1B3F0: mov eax, 00000008h
  loc_00E1B3F5: mov var_18, 00000000h
  loc_00E1B3FC: mov var_54, eax
  loc_00E1B3FF: mov edx, [ecx]
  loc_00E1B401: sub esp, 00000010h
  loc_00E1B404: mov ecx, esp
  loc_00E1B406: mov [ecx], eax
  loc_00E1B408: mov eax, var_50
  loc_00E1B40B: mov [ecx+00000004h], eax
  loc_00E1B40E: mov eax, var_4C
  loc_00E1B411: mov [ecx+00000008h], eax
  loc_00E1B414: mov eax, var_48
  loc_00E1B417: mov [ecx+0000000Ch], eax
  loc_00E1B41A: mov ecx, var_D4
  loc_00E1B420: push ecx
  loc_00E1B421: call [edx+00000038h]
  loc_00E1B424: test eax, eax
  loc_00E1B426: fnclex
  loc_00E1B428: jge 00E1B43Bh
  loc_00E1B42A: mov edx, var_D4
  loc_00E1B430: push 00000038h
  loc_00E1B432: push 006E09F8h
  loc_00E1B437: push edx
  loc_00E1B438: push eax
  loc_00E1B439: call ebx
  loc_00E1B43B: lea eax, var_34
  loc_00E1B43E: lea ecx, var_30
  loc_00E1B441: push eax
  loc_00E1B442: lea edx, var_2C
  loc_00E1B445: push ecx
  loc_00E1B446: lea eax, var_28
  loc_00E1B449: push edx
  loc_00E1B44A: lea ecx, var_24
  loc_00E1B44D: push eax
  loc_00E1B44E: push ecx
  loc_00E1B44F: push 00000005h
  loc_00E1B451: call [00401048h] ; __vbaFreeObjList
  loc_00E1B457: lea edx, var_54
  loc_00E1B45A: lea eax, var_44
  loc_00E1B45D: push edx
  loc_00E1B45E: push eax
  loc_00E1B45F: push 00000002h
  loc_00E1B461: call [00401038h] ; __vbaFreeVarList
  loc_00E1B467: mov ecx, [esi]
  loc_00E1B469: add esp, 00000024h
  loc_00E1B46C: push 006DCBD8h
  loc_00E1B471: push 00000000h
  loc_00E1B473: push 00000018h
  loc_00E1B475: push esi
  loc_00E1B476: call [ecx+00000564h]
  loc_00E1B47C: lea edx, var_28
  loc_00E1B47F: push eax
  loc_00E1B480: push edx
  loc_00E1B481: call edi
  loc_00E1B483: push eax
  loc_00E1B484: lea eax, var_44
  loc_00E1B487: push eax
  loc_00E1B488: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1B48E: add esp, 00000010h
  loc_00E1B491: push eax
  loc_00E1B492: call [00401120h] ; __vbaCastObjVar
  loc_00E1B498: lea ecx, var_2C
  loc_00E1B49B: push eax
  loc_00E1B49C: push ecx
  loc_00E1B49D: call edi
  loc_00E1B49F: mov edx, [eax]
  loc_00E1B4A1: lea ecx, var_30
  loc_00E1B4A4: push ecx
  loc_00E1B4A5: push eax
  loc_00E1B4A6: mov var_C4, eax
  loc_00E1B4AC: call [edx+00000054h]
  loc_00E1B4AF: test eax, eax
  loc_00E1B4B1: fnclex
  loc_00E1B4B3: jge 00E1B4C6h
  loc_00E1B4B5: mov edx, var_C4
  loc_00E1B4BB: push 00000054h
  loc_00E1B4BD: push 006DCBE8h
  loc_00E1B4C2: push edx
  loc_00E1B4C3: push eax
  loc_00E1B4C4: call ebx
  loc_00E1B4C6: lea edx, var_34
  loc_00E1B4C9: mov eax, 00000002h
  loc_00E1B4CE: push edx
  loc_00E1B4CF: mov ecx, var_30
  loc_00E1B4D2: sub esp, 00000010h
  loc_00E1B4D5: mov var_84, eax
  loc_00E1B4DB: mov edx, esp
  loc_00E1B4DD: mov var_7C, 00000008h
  loc_00E1B4E4: mov var_CC, ecx
  loc_00E1B4EA: mov ecx, [ecx]
  loc_00E1B4EC: mov [edx], eax
  loc_00E1B4EE: mov eax, var_80
  loc_00E1B4F1: mov [edx+00000004h], eax
  loc_00E1B4F4: mov eax, var_7C
  loc_00E1B4F7: mov [edx+00000008h], eax
  loc_00E1B4FA: mov eax, var_78
  loc_00E1B4FD: mov [edx+0000000Ch], eax
  loc_00E1B500: mov edx, var_30
  loc_00E1B503: push edx
  loc_00E1B504: call [ecx+00000028h]
  loc_00E1B507: test eax, eax
  loc_00E1B509: fnclex
  loc_00E1B50B: jge 00E1B51Eh
  loc_00E1B50D: mov ecx, var_CC
  loc_00E1B513: push 00000028h
  loc_00E1B515: push 006E09E8h
  loc_00E1B51A: push ecx
  loc_00E1B51B: push eax
  loc_00E1B51C: call ebx
  loc_00E1B51E: mov edx, var_34
  loc_00E1B521: mov eax, [esi]
  loc_00E1B523: push esi
  loc_00E1B524: mov var_D4, edx
  loc_00E1B52A: call [eax+00000454h]
  loc_00E1B530: lea ecx, var_24
  loc_00E1B533: push eax
  loc_00E1B534: push ecx
  loc_00E1B535: call edi
  loc_00E1B537: mov edx, [eax]
  loc_00E1B539: lea ecx, var_18
  loc_00E1B53C: push ecx
  loc_00E1B53D: push eax
  loc_00E1B53E: mov var_BC, eax
  loc_00E1B544: call [edx+00000050h]
  loc_00E1B547: test eax, eax
  loc_00E1B549: fnclex
  loc_00E1B54B: jge 00E1B55Eh
  loc_00E1B54D: mov edx, var_BC
  loc_00E1B553: push 00000050h
  loc_00E1B555: push 006E0410h
  loc_00E1B55A: push edx
  loc_00E1B55B: push eax
  loc_00E1B55C: call ebx
  loc_00E1B55E: mov eax, var_18
  loc_00E1B561: mov ecx, var_D4
  loc_00E1B567: mov var_4C, eax
  loc_00E1B56A: mov eax, 00000008h
  loc_00E1B56F: mov var_18, 00000000h
  loc_00E1B576: mov var_54, eax
  loc_00E1B579: mov edx, [ecx]
  loc_00E1B57B: sub esp, 00000010h
  loc_00E1B57E: mov ecx, esp
  loc_00E1B580: mov [ecx], eax
  loc_00E1B582: mov eax, var_50
  loc_00E1B585: mov [ecx+00000004h], eax
  loc_00E1B588: mov eax, var_4C
  loc_00E1B58B: mov [ecx+00000008h], eax
  loc_00E1B58E: mov eax, var_48
  loc_00E1B591: mov [ecx+0000000Ch], eax
  loc_00E1B594: mov ecx, var_D4
  loc_00E1B59A: push ecx
  loc_00E1B59B: call [edx+00000038h]
  loc_00E1B59E: test eax, eax
  loc_00E1B5A0: fnclex
  loc_00E1B5A2: jge 00E1B5B5h
  loc_00E1B5A4: mov edx, var_D4
  loc_00E1B5AA: push 00000038h
  loc_00E1B5AC: push 006E09F8h
  loc_00E1B5B1: push edx
  loc_00E1B5B2: push eax
  loc_00E1B5B3: call ebx
  loc_00E1B5B5: lea eax, var_34
  loc_00E1B5B8: lea ecx, var_30
  loc_00E1B5BB: push eax
  loc_00E1B5BC: lea edx, var_2C
  loc_00E1B5BF: push ecx
  loc_00E1B5C0: lea eax, var_28
  loc_00E1B5C3: push edx
  loc_00E1B5C4: lea ecx, var_24
  loc_00E1B5C7: push eax
  loc_00E1B5C8: push ecx
  loc_00E1B5C9: push 00000005h
  loc_00E1B5CB: call [00401048h] ; __vbaFreeObjList
  loc_00E1B5D1: lea edx, var_54
  loc_00E1B5D4: lea eax, var_44
  loc_00E1B5D7: push edx
  loc_00E1B5D8: push eax
  loc_00E1B5D9: push 00000002h
  loc_00E1B5DB: call [00401038h] ; __vbaFreeVarList
  loc_00E1B5E1: mov ecx, [esi]
  loc_00E1B5E3: add esp, 00000024h
  loc_00E1B5E6: push esi
  loc_00E1B5E7: call [ecx+0000042Ch]
  loc_00E1B5ED: lea edx, var_24
  loc_00E1B5F0: push eax
  loc_00E1B5F1: push edx
  loc_00E1B5F2: call edi
  loc_00E1B5F4: mov ecx, [eax]
  loc_00E1B5F6: lea edx, var_B8
  loc_00E1B5FC: push edx
  loc_00E1B5FD: push eax
  loc_00E1B5FE: mov var_BC, eax
  loc_00E1B604: call [ecx+00000098h]
  loc_00E1B60A: test eax, eax
  loc_00E1B60C: fnclex
  loc_00E1B60E: jge 00E1B624h
  loc_00E1B610: mov ecx, var_BC
  loc_00E1B616: push 00000098h
  loc_00E1B61B: push 006DCAD0h
  loc_00E1B620: push ecx
  loc_00E1B621: push eax
  loc_00E1B622: call ebx
  loc_00E1B624: xor edx, edx
  loc_00E1B626: cmp var_B8, FFFFFFh
  loc_00E1B62E: lea ecx, var_24
  loc_00E1B631: setz dl
  loc_00E1B634: neg edx
  loc_00E1B636: mov var_C4, dx
  loc_00E1B63D: call [00401254h] ; __vbaFreeObj
  loc_00E1B643: cmp var_C4, 0000h
  loc_00E1B64B: mov eax, [esi]
  loc_00E1B64D: push 006DCBD8h
  loc_00E1B652: push 00000000h
  loc_00E1B654: push 00000018h
  loc_00E1B656: push esi
  loc_00E1B657: jz 00E1B94Ah
  loc_00E1B65D: call [eax+00000564h]
  loc_00E1B663: lea ecx, var_28
  loc_00E1B666: push eax
  loc_00E1B667: push ecx
  loc_00E1B668: call edi
  loc_00E1B66A: lea edx, var_44
  loc_00E1B66D: push eax
  loc_00E1B66E: push edx
  loc_00E1B66F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1B675: add esp, 00000010h
  loc_00E1B678: push eax
  loc_00E1B679: call [00401120h] ; __vbaCastObjVar
  loc_00E1B67F: push eax
  loc_00E1B680: lea eax, var_2C
  loc_00E1B683: push eax
  loc_00E1B684: call edi
  loc_00E1B686: mov ecx, [eax]
  loc_00E1B688: lea edx, var_30
  loc_00E1B68B: push edx
  loc_00E1B68C: push eax
  loc_00E1B68D: mov var_C4, eax
  loc_00E1B693: call [ecx+00000054h]
  loc_00E1B696: test eax, eax
  loc_00E1B698: fnclex
  loc_00E1B69A: jge 00E1B6ADh
  loc_00E1B69C: mov ecx, var_C4
  loc_00E1B6A2: push 00000054h
  loc_00E1B6A4: push 006DCBE8h
  loc_00E1B6A9: push ecx
  loc_00E1B6AA: push eax
  loc_00E1B6AB: call ebx
  loc_00E1B6AD: mov ecx, var_30
  loc_00E1B6B0: mov eax, 00000002h
  loc_00E1B6B5: mov var_7C, 00000009h
  loc_00E1B6BC: mov var_84, eax
  loc_00E1B6C2: mov edx, [ecx]
  loc_00E1B6C4: mov var_CC, ecx
  loc_00E1B6CA: lea ecx, var_34
  loc_00E1B6CD: push ecx
  loc_00E1B6CE: sub esp, 00000010h
  loc_00E1B6D1: mov ecx, esp
  loc_00E1B6D3: mov [ecx], eax
  loc_00E1B6D5: mov eax, var_80
  loc_00E1B6D8: mov [ecx+00000004h], eax
  loc_00E1B6DB: mov eax, var_7C
  loc_00E1B6DE: mov [ecx+00000008h], eax
  loc_00E1B6E1: mov eax, var_78
  loc_00E1B6E4: mov [ecx+0000000Ch], eax
  loc_00E1B6E7: mov ecx, var_30
  loc_00E1B6EA: push ecx
  loc_00E1B6EB: call [edx+00000028h]
  loc_00E1B6EE: test eax, eax
  loc_00E1B6F0: fnclex
  loc_00E1B6F2: jge 00E1B705h
  loc_00E1B6F4: mov edx, var_CC
  loc_00E1B6FA: push 00000028h
  loc_00E1B6FC: push 006E09E8h
  loc_00E1B701: push edx
  loc_00E1B702: push eax
  loc_00E1B703: call ebx
  loc_00E1B705: mov eax, var_34
  loc_00E1B708: mov ecx, [esi]
  loc_00E1B70A: push esi
  loc_00E1B70B: mov var_D4, eax
  loc_00E1B711: call [ecx+00000440h]
  loc_00E1B717: lea edx, var_24
  loc_00E1B71A: push eax
  loc_00E1B71B: push edx
  loc_00E1B71C: call edi
  loc_00E1B71E: mov ecx, [eax]
  loc_00E1B720: lea edx, var_18
  loc_00E1B723: push edx
  loc_00E1B724: push eax
  loc_00E1B725: mov var_BC, eax
  loc_00E1B72B: call [ecx+00000050h]
  loc_00E1B72E: test eax, eax
  loc_00E1B730: fnclex
  loc_00E1B732: jge 00E1B745h
  loc_00E1B734: mov ecx, var_BC
  loc_00E1B73A: push 00000050h
  loc_00E1B73C: push 006E0410h
  loc_00E1B741: push ecx
  loc_00E1B742: push eax
  loc_00E1B743: call ebx
  loc_00E1B745: mov eax, var_18
  loc_00E1B748: mov edx, var_D4
  loc_00E1B74E: mov var_4C, eax
  loc_00E1B751: mov eax, 00000008h
  loc_00E1B756: mov var_18, 00000000h
  loc_00E1B75D: mov var_54, eax
  loc_00E1B760: mov ecx, [edx]
  loc_00E1B762: sub esp, 00000010h
  loc_00E1B765: mov edx, esp
  loc_00E1B767: mov [edx], eax
  loc_00E1B769: mov eax, var_50
  loc_00E1B76C: mov [edx+00000004h], eax
  loc_00E1B76F: mov eax, var_4C
  loc_00E1B772: mov [edx+00000008h], eax
  loc_00E1B775: mov eax, var_48
  loc_00E1B778: mov [edx+0000000Ch], eax
  loc_00E1B77B: mov edx, var_D4
  loc_00E1B781: push edx
  loc_00E1B782: call [ecx+00000038h]
  loc_00E1B785: test eax, eax
  loc_00E1B787: fnclex
  loc_00E1B789: jge 00E1B79Ch
  loc_00E1B78B: mov ecx, var_D4
  loc_00E1B791: push 00000038h
  loc_00E1B793: push 006E09F8h
  loc_00E1B798: push ecx
  loc_00E1B799: push eax
  loc_00E1B79A: call ebx
  loc_00E1B79C: lea edx, var_34
  loc_00E1B79F: lea eax, var_30
  loc_00E1B7A2: push edx
  loc_00E1B7A3: lea ecx, var_2C
  loc_00E1B7A6: push eax
  loc_00E1B7A7: lea edx, var_28
  loc_00E1B7AA: push ecx
  loc_00E1B7AB: lea eax, var_24
  loc_00E1B7AE: push edx
  loc_00E1B7AF: push eax
  loc_00E1B7B0: push 00000005h
  loc_00E1B7B2: call [00401048h] ; __vbaFreeObjList
  loc_00E1B7B8: lea ecx, var_54
  loc_00E1B7BB: lea edx, var_44
  loc_00E1B7BE: push ecx
  loc_00E1B7BF: push edx
  loc_00E1B7C0: push 00000002h
  loc_00E1B7C2: call [00401038h] ; __vbaFreeVarList
  loc_00E1B7C8: mov eax, [esi]
  loc_00E1B7CA: add esp, 00000024h
  loc_00E1B7CD: push 006DCBD8h
  loc_00E1B7D2: push 00000000h
  loc_00E1B7D4: push 00000018h
  loc_00E1B7D6: push esi
  loc_00E1B7D7: call [eax+00000564h]
  loc_00E1B7DD: lea ecx, var_28
  loc_00E1B7E0: push eax
  loc_00E1B7E1: push ecx
  loc_00E1B7E2: call edi
  loc_00E1B7E4: lea edx, var_44
  loc_00E1B7E7: push eax
  loc_00E1B7E8: push edx
  loc_00E1B7E9: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1B7EF: add esp, 00000010h
  loc_00E1B7F2: push eax
  loc_00E1B7F3: call [00401120h] ; __vbaCastObjVar
  loc_00E1B7F9: push eax
  loc_00E1B7FA: lea eax, var_2C
  loc_00E1B7FD: push eax
  loc_00E1B7FE: call edi
  loc_00E1B800: mov ecx, [eax]
  loc_00E1B802: lea edx, var_30
  loc_00E1B805: push edx
  loc_00E1B806: push eax
  loc_00E1B807: mov var_C4, eax
  loc_00E1B80D: call [ecx+00000054h]
  loc_00E1B810: test eax, eax
  loc_00E1B812: fnclex
  loc_00E1B814: jge 00E1B827h
  loc_00E1B816: mov ecx, var_C4
  loc_00E1B81C: push 00000054h
  loc_00E1B81E: push 006DCBE8h
  loc_00E1B823: push ecx
  loc_00E1B824: push eax
  loc_00E1B825: call ebx
  loc_00E1B827: mov ecx, var_30
  loc_00E1B82A: mov eax, 00000002h
  loc_00E1B82F: mov var_7C, 0000000Ah
  loc_00E1B836: mov var_84, eax
  loc_00E1B83C: mov edx, [ecx]
  loc_00E1B83E: mov var_CC, ecx
  loc_00E1B844: lea ecx, var_34
  loc_00E1B847: push ecx
  loc_00E1B848: sub esp, 00000010h
  loc_00E1B84B: mov ecx, esp
  loc_00E1B84D: mov [ecx], eax
  loc_00E1B84F: mov eax, var_80
  loc_00E1B852: mov [ecx+00000004h], eax
  loc_00E1B855: mov eax, var_7C
  loc_00E1B858: mov [ecx+00000008h], eax
  loc_00E1B85B: mov eax, var_78
  loc_00E1B85E: mov [ecx+0000000Ch], eax
  loc_00E1B861: mov ecx, var_30
  loc_00E1B864: push ecx
  loc_00E1B865: call [edx+00000028h]
  loc_00E1B868: test eax, eax
  loc_00E1B86A: fnclex
  loc_00E1B86C: jge 00E1B87Fh
  loc_00E1B86E: mov edx, var_CC
  loc_00E1B874: push 00000028h
  loc_00E1B876: push 006E09E8h
  loc_00E1B87B: push edx
  loc_00E1B87C: push eax
  loc_00E1B87D: call ebx
  loc_00E1B87F: mov eax, var_34
  loc_00E1B882: mov ecx, [esi]
  loc_00E1B884: push esi
  loc_00E1B885: mov var_D4, eax
  loc_00E1B88B: call [ecx+00000434h]
  loc_00E1B891: lea edx, var_24
  loc_00E1B894: push eax
  loc_00E1B895: push edx
  loc_00E1B896: call edi
  loc_00E1B898: mov ecx, [eax]
  loc_00E1B89A: lea edx, var_18
  loc_00E1B89D: push edx
  loc_00E1B89E: push eax
  loc_00E1B89F: mov var_BC, eax
  loc_00E1B8A5: call [ecx+00000050h]
  loc_00E1B8A8: test eax, eax
  loc_00E1B8AA: fnclex
  loc_00E1B8AC: jge 00E1B8BFh
  loc_00E1B8AE: mov ecx, var_BC
  loc_00E1B8B4: push 00000050h
  loc_00E1B8B6: push 006E0410h
  loc_00E1B8BB: push ecx
  loc_00E1B8BC: push eax
  loc_00E1B8BD: call ebx
  loc_00E1B8BF: mov eax, var_18
  loc_00E1B8C2: mov edx, var_D4
  loc_00E1B8C8: mov var_4C, eax
  loc_00E1B8CB: mov eax, 00000008h
  loc_00E1B8D0: mov var_18, 00000000h
  loc_00E1B8D7: mov var_54, eax
  loc_00E1B8DA: mov ecx, [edx]
  loc_00E1B8DC: sub esp, 00000010h
  loc_00E1B8DF: mov edx, esp
  loc_00E1B8E1: mov [edx], eax
  loc_00E1B8E3: mov eax, var_50
  loc_00E1B8E6: mov [edx+00000004h], eax
  loc_00E1B8E9: mov eax, var_4C
  loc_00E1B8EC: mov [edx+00000008h], eax
  loc_00E1B8EF: mov eax, var_48
  loc_00E1B8F2: mov [edx+0000000Ch], eax
  loc_00E1B8F5: mov edx, var_D4
  loc_00E1B8FB: push edx
  loc_00E1B8FC: call [ecx+00000038h]
  loc_00E1B8FF: test eax, eax
  loc_00E1B901: fnclex
  loc_00E1B903: jge 00E1B916h
  loc_00E1B905: mov ecx, var_D4
  loc_00E1B90B: push 00000038h
  loc_00E1B90D: push 006E09F8h
  loc_00E1B912: push ecx
  loc_00E1B913: push eax
  loc_00E1B914: call ebx
  loc_00E1B916: lea edx, var_34
  loc_00E1B919: lea eax, var_30
  loc_00E1B91C: push edx
  loc_00E1B91D: lea ecx, var_2C
  loc_00E1B920: push eax
  loc_00E1B921: lea edx, var_28
  loc_00E1B924: push ecx
  loc_00E1B925: lea eax, var_24
  loc_00E1B928: push edx
  loc_00E1B929: push eax
  loc_00E1B92A: push 00000005h
  loc_00E1B92C: call [00401048h] ; __vbaFreeObjList
  loc_00E1B932: lea ecx, var_54
  loc_00E1B935: lea edx, var_44
  loc_00E1B938: push ecx
  loc_00E1B939: push edx
  loc_00E1B93A: push 00000002h
  loc_00E1B93C: call [00401038h] ; __vbaFreeVarList
  loc_00E1B942: add esp, 00000024h
  loc_00E1B945: jmp 00E1BBAEh
  loc_00E1B94A: call [eax+00000564h]
  loc_00E1B950: lea ecx, var_24
  loc_00E1B953: push eax
  loc_00E1B954: push ecx
  loc_00E1B955: call edi
  loc_00E1B957: lea edx, var_44
  loc_00E1B95A: push eax
  loc_00E1B95B: push edx
  loc_00E1B95C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1B962: add esp, 00000010h
  loc_00E1B965: push eax
  loc_00E1B966: call [00401120h] ; __vbaCastObjVar
  loc_00E1B96C: push eax
  loc_00E1B96D: lea eax, var_28
  loc_00E1B970: push eax
  loc_00E1B971: call edi
  loc_00E1B973: mov ecx, [eax]
  loc_00E1B975: lea edx, var_2C
  loc_00E1B978: push edx
  loc_00E1B979: push eax
  loc_00E1B97A: mov var_BC, eax
  loc_00E1B980: call [ecx+00000054h]
  loc_00E1B983: test eax, eax
  loc_00E1B985: fnclex
  loc_00E1B987: jge 00E1B99Ah
  loc_00E1B989: mov ecx, var_BC
  loc_00E1B98F: push 00000054h
  loc_00E1B991: push 006DCBE8h
  loc_00E1B996: push ecx
  loc_00E1B997: push eax
  loc_00E1B998: call ebx
  loc_00E1B99A: mov ecx, var_2C
  loc_00E1B99D: mov eax, 00000002h
  loc_00E1B9A2: mov var_7C, 00000009h
  loc_00E1B9A9: mov var_84, eax
  loc_00E1B9AF: mov edx, [ecx]
  loc_00E1B9B1: mov var_C4, ecx
  loc_00E1B9B7: lea ecx, var_30
  loc_00E1B9BA: push ecx
  loc_00E1B9BB: sub esp, 00000010h
  loc_00E1B9BE: mov ecx, esp
  loc_00E1B9C0: mov [ecx], eax
  loc_00E1B9C2: mov eax, var_80
  loc_00E1B9C5: mov [ecx+00000004h], eax
  loc_00E1B9C8: mov eax, var_7C
  loc_00E1B9CB: mov [ecx+00000008h], eax
  loc_00E1B9CE: mov eax, var_78
  loc_00E1B9D1: mov [ecx+0000000Ch], eax
  loc_00E1B9D4: mov ecx, var_2C
  loc_00E1B9D7: push ecx
  loc_00E1B9D8: call [edx+00000028h]
  loc_00E1B9DB: test eax, eax
  loc_00E1B9DD: fnclex
  loc_00E1B9DF: jge 00E1B9F2h
  loc_00E1B9E1: mov edx, var_C4
  loc_00E1B9E7: push 00000028h
  loc_00E1B9E9: push 006E09E8h
  loc_00E1B9EE: push edx
  loc_00E1B9EF: push eax
  loc_00E1B9F0: call ebx
  loc_00E1B9F2: sub esp, 00000010h
  loc_00E1B9F5: mov eax, 00000002h
  loc_00E1B9FA: mov edx, esp
  loc_00E1B9FC: mov ecx, var_30
  loc_00E1B9FF: mov var_94, eax
  loc_00E1BA05: mov var_8C, 00000000h
  loc_00E1BA0F: mov [edx], eax
  loc_00E1BA11: mov eax, var_90
  loc_00E1BA17: mov var_CC, ecx
  loc_00E1BA1D: mov ecx, [ecx]
  loc_00E1BA1F: mov [edx+00000004h], eax
  loc_00E1BA22: mov eax, var_8C
  loc_00E1BA28: mov [edx+00000008h], eax
  loc_00E1BA2B: mov eax, var_88
  loc_00E1BA31: mov [edx+0000000Ch], eax
  loc_00E1BA34: mov edx, var_30
  loc_00E1BA37: push edx
  loc_00E1BA38: call [ecx+00000038h]
  loc_00E1BA3B: test eax, eax
  loc_00E1BA3D: fnclex
  loc_00E1BA3F: jge 00E1BA52h
  loc_00E1BA41: mov ecx, var_CC
  loc_00E1BA47: push 00000038h
  loc_00E1BA49: push 006E09F8h
  loc_00E1BA4E: push ecx
  loc_00E1BA4F: push eax
  loc_00E1BA50: call ebx
  loc_00E1BA52: lea edx, var_30
  loc_00E1BA55: lea eax, var_2C
  loc_00E1BA58: push edx
  loc_00E1BA59: lea ecx, var_28
  loc_00E1BA5C: push eax
  loc_00E1BA5D: lea edx, var_24
  loc_00E1BA60: push ecx
  loc_00E1BA61: push edx
  loc_00E1BA62: push 00000004h
  loc_00E1BA64: call [00401048h] ; __vbaFreeObjList
  loc_00E1BA6A: add esp, 00000014h
  loc_00E1BA6D: lea ecx, var_44
  loc_00E1BA70: call [00401024h] ; __vbaFreeVar
  loc_00E1BA76: mov eax, [esi]
  loc_00E1BA78: push 006DCBD8h
  loc_00E1BA7D: push 00000000h
  loc_00E1BA7F: push 00000018h
  loc_00E1BA81: push esi
  loc_00E1BA82: call [eax+00000564h]
  loc_00E1BA88: lea ecx, var_24
  loc_00E1BA8B: push eax
  loc_00E1BA8C: push ecx
  loc_00E1BA8D: call edi
  loc_00E1BA8F: lea edx, var_44
  loc_00E1BA92: push eax
  loc_00E1BA93: push edx
  loc_00E1BA94: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1BA9A: add esp, 00000010h
  loc_00E1BA9D: push eax
  loc_00E1BA9E: call [00401120h] ; __vbaCastObjVar
  loc_00E1BAA4: push eax
  loc_00E1BAA5: lea eax, var_28
  loc_00E1BAA8: push eax
  loc_00E1BAA9: call edi
  loc_00E1BAAB: mov ecx, [eax]
  loc_00E1BAAD: lea edx, var_2C
  loc_00E1BAB0: push edx
  loc_00E1BAB1: push eax
  loc_00E1BAB2: mov var_BC, eax
  loc_00E1BAB8: call [ecx+00000054h]
  loc_00E1BABB: test eax, eax
  loc_00E1BABD: fnclex
  loc_00E1BABF: jge 00E1BAD2h
  loc_00E1BAC1: mov ecx, var_BC
  loc_00E1BAC7: push 00000054h
  loc_00E1BAC9: push 006DCBE8h
  loc_00E1BACE: push ecx
  loc_00E1BACF: push eax
  loc_00E1BAD0: call ebx
  loc_00E1BAD2: mov ecx, var_2C
  loc_00E1BAD5: mov eax, 00000002h
  loc_00E1BADA: mov var_7C, 0000000Ah
  loc_00E1BAE1: mov var_84, eax
  loc_00E1BAE7: mov edx, [ecx]
  loc_00E1BAE9: mov var_C4, ecx
  loc_00E1BAEF: lea ecx, var_30
  loc_00E1BAF2: push ecx
  loc_00E1BAF3: sub esp, 00000010h
  loc_00E1BAF6: mov ecx, esp
  loc_00E1BAF8: mov [ecx], eax
  loc_00E1BAFA: mov eax, var_80
  loc_00E1BAFD: mov [ecx+00000004h], eax
  loc_00E1BB00: mov eax, var_7C
  loc_00E1BB03: mov [ecx+00000008h], eax
  loc_00E1BB06: mov eax, var_78
  loc_00E1BB09: mov [ecx+0000000Ch], eax
  loc_00E1BB0C: mov ecx, var_2C
  loc_00E1BB0F: push ecx
  loc_00E1BB10: call [edx+00000028h]
  loc_00E1BB13: test eax, eax
  loc_00E1BB15: fnclex
  loc_00E1BB17: jge 00E1BB2Ah
  loc_00E1BB19: mov edx, var_C4
  loc_00E1BB1F: push 00000028h
  loc_00E1BB21: push 006E09E8h
  loc_00E1BB26: push edx
  loc_00E1BB27: push eax
  loc_00E1BB28: call ebx
  loc_00E1BB2A: sub esp, 00000010h
  loc_00E1BB2D: mov eax, 00000002h
  loc_00E1BB32: mov edx, esp
  loc_00E1BB34: mov ecx, var_30
  loc_00E1BB37: mov var_94, eax
  loc_00E1BB3D: mov var_8C, 00000000h
  loc_00E1BB47: mov [edx], eax
  loc_00E1BB49: mov eax, var_90
  loc_00E1BB4F: mov var_CC, ecx
  loc_00E1BB55: mov ecx, [ecx]
  loc_00E1BB57: mov [edx+00000004h], eax
  loc_00E1BB5A: mov eax, var_8C
  loc_00E1BB60: mov [edx+00000008h], eax
  loc_00E1BB63: mov eax, var_88
  loc_00E1BB69: mov [edx+0000000Ch], eax
  loc_00E1BB6C: mov edx, var_30
  loc_00E1BB6F: push edx
  loc_00E1BB70: call [ecx+00000038h]
  loc_00E1BB73: test eax, eax
  loc_00E1BB75: fnclex
  loc_00E1BB77: jge 00E1BB8Ah
  loc_00E1BB79: mov ecx, var_CC
  loc_00E1BB7F: push 00000038h
  loc_00E1BB81: push 006E09F8h
  loc_00E1BB86: push ecx
  loc_00E1BB87: push eax
  loc_00E1BB88: call ebx
  loc_00E1BB8A: lea edx, var_30
  loc_00E1BB8D: lea eax, var_2C
  loc_00E1BB90: push edx
  loc_00E1BB91: lea ecx, var_28
  loc_00E1BB94: push eax
  loc_00E1BB95: lea edx, var_24
  loc_00E1BB98: push ecx
  loc_00E1BB99: push edx
  loc_00E1BB9A: push 00000004h
  loc_00E1BB9C: call [00401048h] ; __vbaFreeObjList
  loc_00E1BBA2: add esp, 00000014h
  loc_00E1BBA5: lea ecx, var_44
  loc_00E1BBA8: call [00401024h] ; __vbaFreeVar
  loc_00E1BBAE: mov eax, [esi]
  loc_00E1BBB0: push 006DCBD8h
  loc_00E1BBB5: push 00000000h
  loc_00E1BBB7: push 00000018h
  loc_00E1BBB9: push esi
  loc_00E1BBBA: call [eax+00000564h]
  loc_00E1BBC0: lea ecx, var_28
  loc_00E1BBC3: push eax
  loc_00E1BBC4: push ecx
  loc_00E1BBC5: call edi
  loc_00E1BBC7: lea edx, var_44
  loc_00E1BBCA: push eax
  loc_00E1BBCB: push edx
  loc_00E1BBCC: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1BBD2: add esp, 00000010h
  loc_00E1BBD5: push eax
  loc_00E1BBD6: call [00401120h] ; __vbaCastObjVar
  loc_00E1BBDC: push eax
  loc_00E1BBDD: lea eax, var_2C
  loc_00E1BBE0: push eax
  loc_00E1BBE1: call edi
  loc_00E1BBE3: mov ecx, [eax]
  loc_00E1BBE5: lea edx, var_30
  loc_00E1BBE8: push edx
  loc_00E1BBE9: push eax
  loc_00E1BBEA: mov var_C4, eax
  loc_00E1BBF0: call [ecx+00000054h]
  loc_00E1BBF3: test eax, eax
  loc_00E1BBF5: fnclex
  loc_00E1BBF7: jge 00E1BC0Ah
  loc_00E1BBF9: mov ecx, var_C4
  loc_00E1BBFF: push 00000054h
  loc_00E1BC01: push 006DCBE8h
  loc_00E1BC06: push ecx
  loc_00E1BC07: push eax
  loc_00E1BC08: call ebx
  loc_00E1BC0A: mov ecx, var_30
  loc_00E1BC0D: mov eax, 00000002h
  loc_00E1BC12: mov var_7C, 0000000Bh
  loc_00E1BC19: mov var_84, eax
  loc_00E1BC1F: mov edx, [ecx]
  loc_00E1BC21: mov var_CC, ecx
  loc_00E1BC27: lea ecx, var_34
  loc_00E1BC2A: push ecx
  loc_00E1BC2B: sub esp, 00000010h
  loc_00E1BC2E: mov ecx, esp
  loc_00E1BC30: mov [ecx], eax
  loc_00E1BC32: mov eax, var_80
  loc_00E1BC35: mov [ecx+00000004h], eax
  loc_00E1BC38: mov eax, var_7C
  loc_00E1BC3B: mov [ecx+00000008h], eax
  loc_00E1BC3E: mov eax, var_78
  loc_00E1BC41: mov [ecx+0000000Ch], eax
  loc_00E1BC44: mov ecx, var_30
  loc_00E1BC47: push ecx
  loc_00E1BC48: call [edx+00000028h]
  loc_00E1BC4B: test eax, eax
  loc_00E1BC4D: fnclex
  loc_00E1BC4F: jge 00E1BC62h
  loc_00E1BC51: mov edx, var_CC
  loc_00E1BC57: push 00000028h
  loc_00E1BC59: push 006E09E8h
  loc_00E1BC5E: push edx
  loc_00E1BC5F: push eax
  loc_00E1BC60: call ebx
  loc_00E1BC62: mov eax, var_34
  loc_00E1BC65: mov ecx, [esi]
  loc_00E1BC67: push esi
  loc_00E1BC68: mov var_D4, eax
  loc_00E1BC6E: call [ecx+000003FCh]
  loc_00E1BC74: lea edx, var_24
  loc_00E1BC77: push eax
  loc_00E1BC78: push edx
  loc_00E1BC79: call edi
  loc_00E1BC7B: mov ecx, [eax]
  loc_00E1BC7D: lea edx, var_18
  loc_00E1BC80: push edx
  loc_00E1BC81: push eax
  loc_00E1BC82: mov var_BC, eax
  loc_00E1BC88: call [ecx+00000050h]
  loc_00E1BC8B: test eax, eax
  loc_00E1BC8D: fnclex
  loc_00E1BC8F: jge 00E1BCA2h
  loc_00E1BC91: mov ecx, var_BC
  loc_00E1BC97: push 00000050h
  loc_00E1BC99: push 006E0410h
  loc_00E1BC9E: push ecx
  loc_00E1BC9F: push eax
  loc_00E1BCA0: call ebx
  loc_00E1BCA2: mov eax, var_18
  loc_00E1BCA5: mov edx, var_D4
  loc_00E1BCAB: mov var_4C, eax
  loc_00E1BCAE: mov eax, 00000008h
  loc_00E1BCB3: mov var_18, 00000000h
  loc_00E1BCBA: mov var_54, eax
  loc_00E1BCBD: mov ecx, [edx]
  loc_00E1BCBF: sub esp, 00000010h
  loc_00E1BCC2: mov edx, esp
  loc_00E1BCC4: mov [edx], eax
  loc_00E1BCC6: mov eax, var_50
  loc_00E1BCC9: mov [edx+00000004h], eax
  loc_00E1BCCC: mov eax, var_4C
  loc_00E1BCCF: mov [edx+00000008h], eax
  loc_00E1BCD2: mov eax, var_48
  loc_00E1BCD5: mov [edx+0000000Ch], eax
  loc_00E1BCD8: mov edx, var_D4
  loc_00E1BCDE: push edx
  loc_00E1BCDF: call [ecx+00000038h]
  loc_00E1BCE2: test eax, eax
  loc_00E1BCE4: fnclex
  loc_00E1BCE6: jge 00E1BCF9h
  loc_00E1BCE8: mov ecx, var_D4
  loc_00E1BCEE: push 00000038h
  loc_00E1BCF0: push 006E09F8h
  loc_00E1BCF5: push ecx
  loc_00E1BCF6: push eax
  loc_00E1BCF7: call ebx
  loc_00E1BCF9: lea edx, var_34
  loc_00E1BCFC: lea eax, var_30
  loc_00E1BCFF: push edx
  loc_00E1BD00: lea ecx, var_2C
  loc_00E1BD03: push eax
  loc_00E1BD04: lea edx, var_28
  loc_00E1BD07: push ecx
  loc_00E1BD08: lea eax, var_24
  loc_00E1BD0B: push edx
  loc_00E1BD0C: push eax
  loc_00E1BD0D: push 00000005h
  loc_00E1BD0F: call [00401048h] ; __vbaFreeObjList
  loc_00E1BD15: lea ecx, var_54
  loc_00E1BD18: lea edx, var_44
  loc_00E1BD1B: push ecx
  loc_00E1BD1C: push edx
  loc_00E1BD1D: push 00000002h
  loc_00E1BD1F: call [00401038h] ; __vbaFreeVarList
  loc_00E1BD25: mov eax, [esi]
  loc_00E1BD27: add esp, 00000024h
  loc_00E1BD2A: push 006DCBD8h
  loc_00E1BD2F: push 00000000h
  loc_00E1BD31: push 00000018h
  loc_00E1BD33: push esi
  loc_00E1BD34: call [eax+00000564h]
  loc_00E1BD3A: lea ecx, var_28
  loc_00E1BD3D: push eax
  loc_00E1BD3E: push ecx
  loc_00E1BD3F: call edi
  loc_00E1BD41: lea edx, var_44
  loc_00E1BD44: push eax
  loc_00E1BD45: push edx
  loc_00E1BD46: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1BD4C: add esp, 00000010h
  loc_00E1BD4F: push eax
  loc_00E1BD50: call [00401120h] ; __vbaCastObjVar
  loc_00E1BD56: push eax
  loc_00E1BD57: lea eax, var_2C
  loc_00E1BD5A: push eax
  loc_00E1BD5B: call edi
  loc_00E1BD5D: mov ecx, [eax]
  loc_00E1BD5F: lea edx, var_30
  loc_00E1BD62: push edx
  loc_00E1BD63: push eax
  loc_00E1BD64: mov var_C4, eax
  loc_00E1BD6A: call [ecx+00000054h]
  loc_00E1BD6D: test eax, eax
  loc_00E1BD6F: fnclex
  loc_00E1BD71: jge 00E1BD84h
  loc_00E1BD73: mov ecx, var_C4
  loc_00E1BD79: push 00000054h
  loc_00E1BD7B: push 006DCBE8h
  loc_00E1BD80: push ecx
  loc_00E1BD81: push eax
  loc_00E1BD82: call ebx
  loc_00E1BD84: mov ecx, var_30
  loc_00E1BD87: mov eax, 00000002h
  loc_00E1BD8C: mov var_7C, 0000000Ch
  loc_00E1BD93: mov var_84, eax
  loc_00E1BD99: mov edx, [ecx]
  loc_00E1BD9B: mov var_CC, ecx
  loc_00E1BDA1: lea ecx, var_34
  loc_00E1BDA4: push ecx
  loc_00E1BDA5: sub esp, 00000010h
  loc_00E1BDA8: mov ecx, esp
  loc_00E1BDAA: mov [ecx], eax
  loc_00E1BDAC: mov eax, var_80
  loc_00E1BDAF: mov [ecx+00000004h], eax
  loc_00E1BDB2: mov eax, var_7C
  loc_00E1BDB5: mov [ecx+00000008h], eax
  loc_00E1BDB8: mov eax, var_78
  loc_00E1BDBB: mov [ecx+0000000Ch], eax
  loc_00E1BDBE: mov ecx, var_30
  loc_00E1BDC1: push ecx
  loc_00E1BDC2: call [edx+00000028h]
  loc_00E1BDC5: test eax, eax
  loc_00E1BDC7: fnclex
  loc_00E1BDC9: jge 00E1BDDCh
  loc_00E1BDCB: mov edx, var_CC
  loc_00E1BDD1: push 00000028h
  loc_00E1BDD3: push 006E09E8h
  loc_00E1BDD8: push edx
  loc_00E1BDD9: push eax
  loc_00E1BDDA: call ebx
  loc_00E1BDDC: mov eax, var_34
  loc_00E1BDDF: mov ecx, [esi]
  loc_00E1BDE1: push esi
  loc_00E1BDE2: mov var_D4, eax
  loc_00E1BDE8: call [ecx+000003ECh]
  loc_00E1BDEE: lea edx, var_24
  loc_00E1BDF1: push eax
  loc_00E1BDF2: push edx
  loc_00E1BDF3: call edi
  loc_00E1BDF5: mov ecx, [eax]
  loc_00E1BDF7: lea edx, var_18
  loc_00E1BDFA: push edx
  loc_00E1BDFB: push eax
  loc_00E1BDFC: mov var_BC, eax
  loc_00E1BE02: call [ecx+00000050h]
  loc_00E1BE05: test eax, eax
  loc_00E1BE07: fnclex
  loc_00E1BE09: jge 00E1BE1Ch
  loc_00E1BE0B: mov ecx, var_BC
  loc_00E1BE11: push 00000050h
  loc_00E1BE13: push 006E0410h
  loc_00E1BE18: push ecx
  loc_00E1BE19: push eax
  loc_00E1BE1A: call ebx
  loc_00E1BE1C: mov eax, var_18
  loc_00E1BE1F: mov edx, var_D4
  loc_00E1BE25: mov var_4C, eax
  loc_00E1BE28: mov eax, 00000008h
  loc_00E1BE2D: mov var_18, 00000000h
  loc_00E1BE34: mov var_54, eax
  loc_00E1BE37: mov ecx, [edx]
  loc_00E1BE39: sub esp, 00000010h
  loc_00E1BE3C: mov edx, esp
  loc_00E1BE3E: mov [edx], eax
  loc_00E1BE40: mov eax, var_50
  loc_00E1BE43: mov [edx+00000004h], eax
  loc_00E1BE46: mov eax, var_4C
  loc_00E1BE49: mov [edx+00000008h], eax
  loc_00E1BE4C: mov eax, var_48
  loc_00E1BE4F: mov [edx+0000000Ch], eax
  loc_00E1BE52: mov edx, var_D4
  loc_00E1BE58: push edx
  loc_00E1BE59: call [ecx+00000038h]
  loc_00E1BE5C: test eax, eax
  loc_00E1BE5E: fnclex
  loc_00E1BE60: jge 00E1BE73h
  loc_00E1BE62: mov ecx, var_D4
  loc_00E1BE68: push 00000038h
  loc_00E1BE6A: push 006E09F8h
  loc_00E1BE6F: push ecx
  loc_00E1BE70: push eax
  loc_00E1BE71: call ebx
  loc_00E1BE73: lea edx, var_34
  loc_00E1BE76: lea eax, var_30
  loc_00E1BE79: push edx
  loc_00E1BE7A: lea ecx, var_2C
  loc_00E1BE7D: push eax
  loc_00E1BE7E: lea edx, var_28
  loc_00E1BE81: push ecx
  loc_00E1BE82: lea eax, var_24
  loc_00E1BE85: push edx
  loc_00E1BE86: push eax
  loc_00E1BE87: push 00000005h
  loc_00E1BE89: call [00401048h] ; __vbaFreeObjList
  loc_00E1BE8F: lea ecx, var_54
  loc_00E1BE92: lea edx, var_44
  loc_00E1BE95: push ecx
  loc_00E1BE96: push edx
  loc_00E1BE97: push 00000002h
  loc_00E1BE99: call [00401038h] ; __vbaFreeVarList
  loc_00E1BE9F: mov eax, [esi]
  loc_00E1BEA1: add esp, 00000024h
  loc_00E1BEA4: push 006DCBD8h
  loc_00E1BEA9: push 00000000h
  loc_00E1BEAB: push 00000018h
  loc_00E1BEAD: push esi
  loc_00E1BEAE: call [eax+00000564h]
  loc_00E1BEB4: lea ecx, var_28
  loc_00E1BEB7: push eax
  loc_00E1BEB8: push ecx
  loc_00E1BEB9: call edi
  loc_00E1BEBB: lea edx, var_44
  loc_00E1BEBE: push eax
  loc_00E1BEBF: push edx
  loc_00E1BEC0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1BEC6: add esp, 00000010h
  loc_00E1BEC9: push eax
  loc_00E1BECA: call [00401120h] ; __vbaCastObjVar
  loc_00E1BED0: push eax
  loc_00E1BED1: lea eax, var_2C
  loc_00E1BED4: push eax
  loc_00E1BED5: call edi
  loc_00E1BED7: mov ecx, [eax]
  loc_00E1BED9: lea edx, var_30
  loc_00E1BEDC: push edx
  loc_00E1BEDD: push eax
  loc_00E1BEDE: mov var_C4, eax
  loc_00E1BEE4: call [ecx+00000054h]
  loc_00E1BEE7: test eax, eax
  loc_00E1BEE9: fnclex
  loc_00E1BEEB: jge 00E1BEFEh
  loc_00E1BEED: mov ecx, var_C4
  loc_00E1BEF3: push 00000054h
  loc_00E1BEF5: push 006DCBE8h
  loc_00E1BEFA: push ecx
  loc_00E1BEFB: push eax
  loc_00E1BEFC: call ebx
  loc_00E1BEFE: mov ecx, var_30
  loc_00E1BF01: mov eax, 00000002h
  loc_00E1BF06: mov var_7C, 0000000Dh
  loc_00E1BF0D: mov var_84, eax
  loc_00E1BF13: mov edx, [ecx]
  loc_00E1BF15: mov var_CC, ecx
  loc_00E1BF1B: lea ecx, var_34
  loc_00E1BF1E: push ecx
  loc_00E1BF1F: sub esp, 00000010h
  loc_00E1BF22: mov ecx, esp
  loc_00E1BF24: mov [ecx], eax
  loc_00E1BF26: mov eax, var_80
  loc_00E1BF29: mov [ecx+00000004h], eax
  loc_00E1BF2C: mov eax, var_7C
  loc_00E1BF2F: mov [ecx+00000008h], eax
  loc_00E1BF32: mov eax, var_78
  loc_00E1BF35: mov [ecx+0000000Ch], eax
  loc_00E1BF38: mov ecx, var_30
  loc_00E1BF3B: push ecx
  loc_00E1BF3C: call [edx+00000028h]
  loc_00E1BF3F: test eax, eax
  loc_00E1BF41: fnclex
  loc_00E1BF43: jge 00E1BF56h
  loc_00E1BF45: mov edx, var_CC
  loc_00E1BF4B: push 00000028h
  loc_00E1BF4D: push 006E09E8h
  loc_00E1BF52: push edx
  loc_00E1BF53: push eax
  loc_00E1BF54: call ebx
  loc_00E1BF56: mov eax, var_34
  loc_00E1BF59: mov ecx, [esi]
  loc_00E1BF5B: push esi
  loc_00E1BF5C: mov var_D4, eax
  loc_00E1BF62: call [ecx+000003DCh]
  loc_00E1BF68: lea edx, var_24
  loc_00E1BF6B: push eax
  loc_00E1BF6C: push edx
  loc_00E1BF6D: call edi
  loc_00E1BF6F: mov ecx, [eax]
  loc_00E1BF71: lea edx, var_18
  loc_00E1BF74: push edx
  loc_00E1BF75: push eax
  loc_00E1BF76: mov var_BC, eax
  loc_00E1BF7C: call [ecx+00000050h]
  loc_00E1BF7F: test eax, eax
  loc_00E1BF81: fnclex
  loc_00E1BF83: jge 00E1BF96h
  loc_00E1BF85: mov ecx, var_BC
  loc_00E1BF8B: push 00000050h
  loc_00E1BF8D: push 006E0410h
  loc_00E1BF92: push ecx
  loc_00E1BF93: push eax
  loc_00E1BF94: call ebx
  loc_00E1BF96: mov eax, var_18
  loc_00E1BF99: mov edx, var_D4
  loc_00E1BF9F: mov var_4C, eax
  loc_00E1BFA2: mov eax, 00000008h
  loc_00E1BFA7: mov var_18, 00000000h
  loc_00E1BFAE: mov var_54, eax
  loc_00E1BFB1: mov ecx, [edx]
  loc_00E1BFB3: sub esp, 00000010h
  loc_00E1BFB6: mov edx, esp
  loc_00E1BFB8: mov [edx], eax
  loc_00E1BFBA: mov eax, var_50
  loc_00E1BFBD: mov [edx+00000004h], eax
  loc_00E1BFC0: mov eax, var_4C
  loc_00E1BFC3: mov [edx+00000008h], eax
  loc_00E1BFC6: mov eax, var_48
  loc_00E1BFC9: mov [edx+0000000Ch], eax
  loc_00E1BFCC: mov edx, var_D4
  loc_00E1BFD2: push edx
  loc_00E1BFD3: call [ecx+00000038h]
  loc_00E1BFD6: test eax, eax
  loc_00E1BFD8: fnclex
  loc_00E1BFDA: jge 00E1BFEDh
  loc_00E1BFDC: mov ecx, var_D4
  loc_00E1BFE2: push 00000038h
  loc_00E1BFE4: push 006E09F8h
  loc_00E1BFE9: push ecx
  loc_00E1BFEA: push eax
  loc_00E1BFEB: call ebx
  loc_00E1BFED: lea edx, var_34
  loc_00E1BFF0: lea eax, var_30
  loc_00E1BFF3: push edx
  loc_00E1BFF4: lea ecx, var_2C
  loc_00E1BFF7: push eax
  loc_00E1BFF8: lea edx, var_28
  loc_00E1BFFB: push ecx
  loc_00E1BFFC: lea eax, var_24
  loc_00E1BFFF: push edx
  loc_00E1C000: push eax
  loc_00E1C001: push 00000005h
  loc_00E1C003: call [00401048h] ; __vbaFreeObjList
  loc_00E1C009: lea ecx, var_54
  loc_00E1C00C: lea edx, var_44
  loc_00E1C00F: push ecx
  loc_00E1C010: push edx
  loc_00E1C011: push 00000002h
  loc_00E1C013: call [00401038h] ; __vbaFreeVarList
  loc_00E1C019: mov eax, [esi]
  loc_00E1C01B: add esp, 00000024h
  loc_00E1C01E: push 006DCBD8h
  loc_00E1C023: push 00000000h
  loc_00E1C025: push 00000018h
  loc_00E1C027: push esi
  loc_00E1C028: call [eax+00000564h]
  loc_00E1C02E: lea ecx, var_28
  loc_00E1C031: push eax
  loc_00E1C032: push ecx
  loc_00E1C033: call edi
  loc_00E1C035: lea edx, var_44
  loc_00E1C038: push eax
  loc_00E1C039: push edx
  loc_00E1C03A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1C040: add esp, 00000010h
  loc_00E1C043: push eax
  loc_00E1C044: call [00401120h] ; __vbaCastObjVar
  loc_00E1C04A: push eax
  loc_00E1C04B: lea eax, var_2C
  loc_00E1C04E: push eax
  loc_00E1C04F: call edi
  loc_00E1C051: mov ecx, [eax]
  loc_00E1C053: lea edx, var_30
  loc_00E1C056: push edx
  loc_00E1C057: push eax
  loc_00E1C058: mov var_C4, eax
  loc_00E1C05E: call [ecx+00000054h]
  loc_00E1C061: test eax, eax
  loc_00E1C063: fnclex
  loc_00E1C065: jge 00E1C078h
  loc_00E1C067: mov ecx, var_C4
  loc_00E1C06D: push 00000054h
  loc_00E1C06F: push 006DCBE8h
  loc_00E1C074: push ecx
  loc_00E1C075: push eax
  loc_00E1C076: call ebx
  loc_00E1C078: mov ecx, var_30
  loc_00E1C07B: mov eax, 00000002h
  loc_00E1C080: mov var_7C, 0000000Eh
  loc_00E1C087: mov var_84, eax
  loc_00E1C08D: mov edx, [ecx]
  loc_00E1C08F: mov var_CC, ecx
  loc_00E1C095: lea ecx, var_34
  loc_00E1C098: push ecx
  loc_00E1C099: sub esp, 00000010h
  loc_00E1C09C: mov ecx, esp
  loc_00E1C09E: mov [ecx], eax
  loc_00E1C0A0: mov eax, var_80
  loc_00E1C0A3: mov [ecx+00000004h], eax
  loc_00E1C0A6: mov eax, var_7C
  loc_00E1C0A9: mov [ecx+00000008h], eax
  loc_00E1C0AC: mov eax, var_78
  loc_00E1C0AF: mov [ecx+0000000Ch], eax
  loc_00E1C0B2: mov ecx, var_30
  loc_00E1C0B5: push ecx
  loc_00E1C0B6: call [edx+00000028h]
  loc_00E1C0B9: test eax, eax
  loc_00E1C0BB: fnclex
  loc_00E1C0BD: jge 00E1C0D0h
  loc_00E1C0BF: mov edx, var_CC
  loc_00E1C0C5: push 00000028h
  loc_00E1C0C7: push 006E09E8h
  loc_00E1C0CC: push edx
  loc_00E1C0CD: push eax
  loc_00E1C0CE: call ebx
  loc_00E1C0D0: mov eax, var_34
  loc_00E1C0D3: mov ecx, [esi]
  loc_00E1C0D5: push esi
  loc_00E1C0D6: mov var_D4, eax
  loc_00E1C0DC: call [ecx+000003CCh]
  loc_00E1C0E2: lea edx, var_24
  loc_00E1C0E5: push eax
  loc_00E1C0E6: push edx
  loc_00E1C0E7: call edi
  loc_00E1C0E9: mov ecx, [eax]
  loc_00E1C0EB: lea edx, var_18
  loc_00E1C0EE: push edx
  loc_00E1C0EF: push eax
  loc_00E1C0F0: mov var_BC, eax
  loc_00E1C0F6: call [ecx+00000050h]
  loc_00E1C0F9: test eax, eax
  loc_00E1C0FB: fnclex
  loc_00E1C0FD: jge 00E1C110h
  loc_00E1C0FF: mov ecx, var_BC
  loc_00E1C105: push 00000050h
  loc_00E1C107: push 006E0410h
  loc_00E1C10C: push ecx
  loc_00E1C10D: push eax
  loc_00E1C10E: call ebx
  loc_00E1C110: mov eax, var_18
  loc_00E1C113: mov edx, var_D4
  loc_00E1C119: mov var_4C, eax
  loc_00E1C11C: mov eax, 00000008h
  loc_00E1C121: mov var_18, 00000000h
  loc_00E1C128: mov var_54, eax
  loc_00E1C12B: mov ecx, [edx]
  loc_00E1C12D: sub esp, 00000010h
  loc_00E1C130: mov edx, esp
  loc_00E1C132: mov [edx], eax
  loc_00E1C134: mov eax, var_50
  loc_00E1C137: mov [edx+00000004h], eax
  loc_00E1C13A: mov eax, var_4C
  loc_00E1C13D: mov [edx+00000008h], eax
  loc_00E1C140: mov eax, var_48
  loc_00E1C143: mov [edx+0000000Ch], eax
  loc_00E1C146: mov edx, var_D4
  loc_00E1C14C: push edx
  loc_00E1C14D: call [ecx+00000038h]
  loc_00E1C150: test eax, eax
  loc_00E1C152: fnclex
  loc_00E1C154: jge 00E1C167h
  loc_00E1C156: mov ecx, var_D4
  loc_00E1C15C: push 00000038h
  loc_00E1C15E: push 006E09F8h
  loc_00E1C163: push ecx
  loc_00E1C164: push eax
  loc_00E1C165: call ebx
  loc_00E1C167: lea edx, var_34
  loc_00E1C16A: lea eax, var_30
  loc_00E1C16D: push edx
  loc_00E1C16E: lea ecx, var_2C
  loc_00E1C171: push eax
  loc_00E1C172: lea edx, var_28
  loc_00E1C175: push ecx
  loc_00E1C176: lea eax, var_24
  loc_00E1C179: push edx
  loc_00E1C17A: push eax
  loc_00E1C17B: push 00000005h
  loc_00E1C17D: call [00401048h] ; __vbaFreeObjList
  loc_00E1C183: lea ecx, var_54
  loc_00E1C186: lea edx, var_44
  loc_00E1C189: push ecx
  loc_00E1C18A: push edx
  loc_00E1C18B: push 00000002h
  loc_00E1C18D: call [00401038h] ; __vbaFreeVarList
  loc_00E1C193: mov eax, [esi]
  loc_00E1C195: add esp, 00000024h
  loc_00E1C198: push 006DCBD8h
  loc_00E1C19D: push 00000000h
  loc_00E1C19F: push 00000018h
  loc_00E1C1A1: push esi
  loc_00E1C1A2: call [eax+00000564h]
  loc_00E1C1A8: lea ecx, var_28
  loc_00E1C1AB: push eax
  loc_00E1C1AC: push ecx
  loc_00E1C1AD: call edi
  loc_00E1C1AF: lea edx, var_44
  loc_00E1C1B2: push eax
  loc_00E1C1B3: push edx
  loc_00E1C1B4: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1C1BA: add esp, 00000010h
  loc_00E1C1BD: push eax
  loc_00E1C1BE: call [00401120h] ; __vbaCastObjVar
  loc_00E1C1C4: push eax
  loc_00E1C1C5: lea eax, var_2C
  loc_00E1C1C8: push eax
  loc_00E1C1C9: call edi
  loc_00E1C1CB: mov ecx, [eax]
  loc_00E1C1CD: lea edx, var_30
  loc_00E1C1D0: push edx
  loc_00E1C1D1: push eax
  loc_00E1C1D2: mov var_C4, eax
  loc_00E1C1D8: call [ecx+00000054h]
  loc_00E1C1DB: test eax, eax
  loc_00E1C1DD: fnclex
  loc_00E1C1DF: jge 00E1C1F2h
  loc_00E1C1E1: mov ecx, var_C4
  loc_00E1C1E7: push 00000054h
  loc_00E1C1E9: push 006DCBE8h
  loc_00E1C1EE: push ecx
  loc_00E1C1EF: push eax
  loc_00E1C1F0: call ebx
  loc_00E1C1F2: mov ecx, var_30
  loc_00E1C1F5: mov eax, 00000002h
  loc_00E1C1FA: mov var_7C, 0000000Fh
  loc_00E1C201: mov var_84, eax
  loc_00E1C207: mov edx, [ecx]
  loc_00E1C209: mov var_CC, ecx
  loc_00E1C20F: lea ecx, var_34
  loc_00E1C212: push ecx
  loc_00E1C213: sub esp, 00000010h
  loc_00E1C216: mov ecx, esp
  loc_00E1C218: mov [ecx], eax
  loc_00E1C21A: mov eax, var_80
  loc_00E1C21D: mov [ecx+00000004h], eax
  loc_00E1C220: mov eax, var_7C
  loc_00E1C223: mov [ecx+00000008h], eax
  loc_00E1C226: mov eax, var_78
  loc_00E1C229: mov [ecx+0000000Ch], eax
  loc_00E1C22C: mov ecx, var_30
  loc_00E1C22F: push ecx
  loc_00E1C230: call [edx+00000028h]
  loc_00E1C233: test eax, eax
  loc_00E1C235: fnclex
  loc_00E1C237: jge 00E1C24Ah
  loc_00E1C239: mov edx, var_CC
  loc_00E1C23F: push 00000028h
  loc_00E1C241: push 006E09E8h
  loc_00E1C246: push edx
  loc_00E1C247: push eax
  loc_00E1C248: call ebx
  loc_00E1C24A: mov eax, var_34
  loc_00E1C24D: mov ecx, [esi]
  loc_00E1C24F: push esi
  loc_00E1C250: mov var_D4, eax
  loc_00E1C256: call [ecx+000003BCh]
  loc_00E1C25C: lea edx, var_24
  loc_00E1C25F: push eax
  loc_00E1C260: push edx
  loc_00E1C261: call edi
  loc_00E1C263: mov ecx, [eax]
  loc_00E1C265: lea edx, var_18
  loc_00E1C268: push edx
  loc_00E1C269: push eax
  loc_00E1C26A: mov var_BC, eax
  loc_00E1C270: call [ecx+00000050h]
  loc_00E1C273: test eax, eax
  loc_00E1C275: fnclex
  loc_00E1C277: jge 00E1C28Ah
  loc_00E1C279: mov ecx, var_BC
  loc_00E1C27F: push 00000050h
  loc_00E1C281: push 006E0410h
  loc_00E1C286: push ecx
  loc_00E1C287: push eax
  loc_00E1C288: call ebx
  loc_00E1C28A: mov eax, var_18
  loc_00E1C28D: mov edx, var_D4
  loc_00E1C293: mov var_4C, eax
  loc_00E1C296: mov eax, 00000008h
  loc_00E1C29B: mov var_18, 00000000h
  loc_00E1C2A2: mov var_54, eax
  loc_00E1C2A5: mov ecx, [edx]
  loc_00E1C2A7: sub esp, 00000010h
  loc_00E1C2AA: mov edx, esp
  loc_00E1C2AC: mov [edx], eax
  loc_00E1C2AE: mov eax, var_50
  loc_00E1C2B1: mov [edx+00000004h], eax
  loc_00E1C2B4: mov eax, var_4C
  loc_00E1C2B7: mov [edx+00000008h], eax
  loc_00E1C2BA: mov eax, var_48
  loc_00E1C2BD: mov [edx+0000000Ch], eax
  loc_00E1C2C0: mov edx, var_D4
  loc_00E1C2C6: push edx
  loc_00E1C2C7: call [ecx+00000038h]
  loc_00E1C2CA: test eax, eax
  loc_00E1C2CC: fnclex
  loc_00E1C2CE: jge 00E1C2E1h
  loc_00E1C2D0: mov ecx, var_D4
  loc_00E1C2D6: push 00000038h
  loc_00E1C2D8: push 006E09F8h
  loc_00E1C2DD: push ecx
  loc_00E1C2DE: push eax
  loc_00E1C2DF: call ebx
  loc_00E1C2E1: lea edx, var_34
  loc_00E1C2E4: lea eax, var_30
  loc_00E1C2E7: push edx
  loc_00E1C2E8: lea ecx, var_2C
  loc_00E1C2EB: push eax
  loc_00E1C2EC: lea edx, var_28
  loc_00E1C2EF: push ecx
  loc_00E1C2F0: lea eax, var_24
  loc_00E1C2F3: push edx
  loc_00E1C2F4: push eax
  loc_00E1C2F5: push 00000005h
  loc_00E1C2F7: call [00401048h] ; __vbaFreeObjList
  loc_00E1C2FD: lea ecx, var_54
  loc_00E1C300: lea edx, var_44
  loc_00E1C303: push ecx
  loc_00E1C304: push edx
  loc_00E1C305: push 00000002h
  loc_00E1C307: call [00401038h] ; __vbaFreeVarList
  loc_00E1C30D: mov eax, [esi]
  loc_00E1C30F: add esp, 00000024h
  loc_00E1C312: push 006DCBD8h
  loc_00E1C317: push 00000000h
  loc_00E1C319: push 00000018h
  loc_00E1C31B: push esi
  loc_00E1C31C: call [eax+00000564h]
  loc_00E1C322: lea ecx, var_28
  loc_00E1C325: push eax
  loc_00E1C326: push ecx
  loc_00E1C327: call edi
  loc_00E1C329: lea edx, var_44
  loc_00E1C32C: push eax
  loc_00E1C32D: push edx
  loc_00E1C32E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1C334: add esp, 00000010h
  loc_00E1C337: push eax
  loc_00E1C338: call [00401120h] ; __vbaCastObjVar
  loc_00E1C33E: push eax
  loc_00E1C33F: lea eax, var_2C
  loc_00E1C342: push eax
  loc_00E1C343: call edi
  loc_00E1C345: mov ecx, [eax]
  loc_00E1C347: lea edx, var_30
  loc_00E1C34A: push edx
  loc_00E1C34B: push eax
  loc_00E1C34C: mov var_C4, eax
  loc_00E1C352: call [ecx+00000054h]
  loc_00E1C355: test eax, eax
  loc_00E1C357: fnclex
  loc_00E1C359: jge 00E1C36Ch
  loc_00E1C35B: mov ecx, var_C4
  loc_00E1C361: push 00000054h
  loc_00E1C363: push 006DCBE8h
  loc_00E1C368: push ecx
  loc_00E1C369: push eax
  loc_00E1C36A: call ebx
  loc_00E1C36C: mov ecx, var_30
  loc_00E1C36F: mov eax, 00000002h
  loc_00E1C374: mov var_7C, 00000010h
  loc_00E1C37B: mov var_84, eax
  loc_00E1C381: mov edx, [ecx]
  loc_00E1C383: mov var_CC, ecx
  loc_00E1C389: lea ecx, var_34
  loc_00E1C38C: push ecx
  loc_00E1C38D: sub esp, 00000010h
  loc_00E1C390: mov ecx, esp
  loc_00E1C392: mov [ecx], eax
  loc_00E1C394: mov eax, var_80
  loc_00E1C397: mov [ecx+00000004h], eax
  loc_00E1C39A: mov eax, var_7C
  loc_00E1C39D: mov [ecx+00000008h], eax
  loc_00E1C3A0: mov eax, var_78
  loc_00E1C3A3: mov [ecx+0000000Ch], eax
  loc_00E1C3A6: mov ecx, var_30
  loc_00E1C3A9: push ecx
  loc_00E1C3AA: call [edx+00000028h]
  loc_00E1C3AD: test eax, eax
  loc_00E1C3AF: fnclex
  loc_00E1C3B1: jge 00E1C3C4h
  loc_00E1C3B3: mov edx, var_CC
  loc_00E1C3B9: push 00000028h
  loc_00E1C3BB: push 006E09E8h
  loc_00E1C3C0: push edx
  loc_00E1C3C1: push eax
  loc_00E1C3C2: call ebx
  loc_00E1C3C4: mov eax, var_34
  loc_00E1C3C7: mov ecx, [esi]
  loc_00E1C3C9: push esi
  loc_00E1C3CA: mov var_D4, eax
  loc_00E1C3D0: call [ecx+000003A8h]
  loc_00E1C3D6: lea edx, var_24
  loc_00E1C3D9: push eax
  loc_00E1C3DA: push edx
  loc_00E1C3DB: call edi
  loc_00E1C3DD: mov ecx, [eax]
  loc_00E1C3DF: lea edx, var_18
  loc_00E1C3E2: push edx
  loc_00E1C3E3: push eax
  loc_00E1C3E4: mov var_BC, eax
  loc_00E1C3EA: call [ecx+00000050h]
  loc_00E1C3ED: test eax, eax
  loc_00E1C3EF: fnclex
  loc_00E1C3F1: jge 00E1C404h
  loc_00E1C3F3: mov ecx, var_BC
  loc_00E1C3F9: push 00000050h
  loc_00E1C3FB: push 006E0410h
  loc_00E1C400: push ecx
  loc_00E1C401: push eax
  loc_00E1C402: call ebx
  loc_00E1C404: mov eax, var_18
  loc_00E1C407: mov edx, var_D4
  loc_00E1C40D: mov var_4C, eax
  loc_00E1C410: mov eax, 00000008h
  loc_00E1C415: mov var_18, 00000000h
  loc_00E1C41C: mov var_54, eax
  loc_00E1C41F: mov ecx, [edx]
  loc_00E1C421: sub esp, 00000010h
  loc_00E1C424: mov edx, esp
  loc_00E1C426: mov [edx], eax
  loc_00E1C428: mov eax, var_50
  loc_00E1C42B: mov [edx+00000004h], eax
  loc_00E1C42E: mov eax, var_4C
  loc_00E1C431: mov [edx+00000008h], eax
  loc_00E1C434: mov eax, var_48
  loc_00E1C437: mov [edx+0000000Ch], eax
  loc_00E1C43A: mov edx, var_D4
  loc_00E1C440: push edx
  loc_00E1C441: call [ecx+00000038h]
  loc_00E1C444: test eax, eax
  loc_00E1C446: fnclex
  loc_00E1C448: jge 00E1C45Bh
  loc_00E1C44A: mov ecx, var_D4
  loc_00E1C450: push 00000038h
  loc_00E1C452: push 006E09F8h
  loc_00E1C457: push ecx
  loc_00E1C458: push eax
  loc_00E1C459: call ebx
  loc_00E1C45B: lea edx, var_34
  loc_00E1C45E: lea eax, var_30
  loc_00E1C461: push edx
  loc_00E1C462: lea ecx, var_2C
  loc_00E1C465: push eax
  loc_00E1C466: lea edx, var_28
  loc_00E1C469: push ecx
  loc_00E1C46A: lea eax, var_24
  loc_00E1C46D: push edx
  loc_00E1C46E: push eax
  loc_00E1C46F: push 00000005h
  loc_00E1C471: call [00401048h] ; __vbaFreeObjList
  loc_00E1C477: lea ecx, var_54
  loc_00E1C47A: lea edx, var_44
  loc_00E1C47D: push ecx
  loc_00E1C47E: push edx
  loc_00E1C47F: push 00000002h
  loc_00E1C481: call [00401038h] ; __vbaFreeVarList
  loc_00E1C487: mov eax, [esi]
  loc_00E1C489: add esp, 00000024h
  loc_00E1C48C: push 006DCBD8h
  loc_00E1C491: push 00000000h
  loc_00E1C493: push 00000018h
  loc_00E1C495: push esi
  loc_00E1C496: call [eax+00000564h]
  loc_00E1C49C: lea ecx, var_28
  loc_00E1C49F: push eax
  loc_00E1C4A0: push ecx
  loc_00E1C4A1: call edi
  loc_00E1C4A3: lea edx, var_44
  loc_00E1C4A6: push eax
  loc_00E1C4A7: push edx
  loc_00E1C4A8: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1C4AE: add esp, 00000010h
  loc_00E1C4B1: push eax
  loc_00E1C4B2: call [00401120h] ; __vbaCastObjVar
  loc_00E1C4B8: push eax
  loc_00E1C4B9: lea eax, var_2C
  loc_00E1C4BC: push eax
  loc_00E1C4BD: call edi
  loc_00E1C4BF: mov ecx, [eax]
  loc_00E1C4C1: lea edx, var_30
  loc_00E1C4C4: push edx
  loc_00E1C4C5: push eax
  loc_00E1C4C6: mov var_C4, eax
  loc_00E1C4CC: call [ecx+00000054h]
  loc_00E1C4CF: test eax, eax
  loc_00E1C4D1: fnclex
  loc_00E1C4D3: jge 00E1C4E6h
  loc_00E1C4D5: mov ecx, var_C4
  loc_00E1C4DB: push 00000054h
  loc_00E1C4DD: push 006DCBE8h
  loc_00E1C4E2: push ecx
  loc_00E1C4E3: push eax
  loc_00E1C4E4: call ebx
  loc_00E1C4E6: mov ecx, var_30
  loc_00E1C4E9: mov eax, 00000002h
  loc_00E1C4EE: mov var_7C, 00000011h
  loc_00E1C4F5: mov var_84, eax
  loc_00E1C4FB: mov edx, [ecx]
  loc_00E1C4FD: mov var_CC, ecx
  loc_00E1C503: lea ecx, var_34
  loc_00E1C506: push ecx
  loc_00E1C507: sub esp, 00000010h
  loc_00E1C50A: mov ecx, esp
  loc_00E1C50C: mov [ecx], eax
  loc_00E1C50E: mov eax, var_80
  loc_00E1C511: mov [ecx+00000004h], eax
  loc_00E1C514: mov eax, var_7C
  loc_00E1C517: mov [ecx+00000008h], eax
  loc_00E1C51A: mov eax, var_78
  loc_00E1C51D: mov [ecx+0000000Ch], eax
  loc_00E1C520: mov ecx, var_30
  loc_00E1C523: push ecx
  loc_00E1C524: call [edx+00000028h]
  loc_00E1C527: test eax, eax
  loc_00E1C529: fnclex
  loc_00E1C52B: jge 00E1C53Eh
  loc_00E1C52D: mov edx, var_CC
  loc_00E1C533: push 00000028h
  loc_00E1C535: push 006E09E8h
  loc_00E1C53A: push edx
  loc_00E1C53B: push eax
  loc_00E1C53C: call ebx
  loc_00E1C53E: mov eax, var_34
  loc_00E1C541: mov ecx, [esi]
  loc_00E1C543: push esi
  loc_00E1C544: mov var_D4, eax
  loc_00E1C54A: call [ecx+00000398h]
  loc_00E1C550: lea edx, var_24
  loc_00E1C553: push eax
  loc_00E1C554: push edx
  loc_00E1C555: call edi
  loc_00E1C557: mov ecx, [eax]
  loc_00E1C559: lea edx, var_18
  loc_00E1C55C: push edx
  loc_00E1C55D: push eax
  loc_00E1C55E: mov var_BC, eax
  loc_00E1C564: call [ecx+00000050h]
  loc_00E1C567: test eax, eax
  loc_00E1C569: fnclex
  loc_00E1C56B: jge 00E1C57Eh
  loc_00E1C56D: mov ecx, var_BC
  loc_00E1C573: push 00000050h
  loc_00E1C575: push 006E0410h
  loc_00E1C57A: push ecx
  loc_00E1C57B: push eax
  loc_00E1C57C: call ebx
  loc_00E1C57E: mov eax, var_18
  loc_00E1C581: mov edx, var_D4
  loc_00E1C587: mov var_4C, eax
  loc_00E1C58A: mov eax, 00000008h
  loc_00E1C58F: mov var_18, 00000000h
  loc_00E1C596: mov var_54, eax
  loc_00E1C599: mov ecx, [edx]
  loc_00E1C59B: sub esp, 00000010h
  loc_00E1C59E: mov edx, esp
  loc_00E1C5A0: mov [edx], eax
  loc_00E1C5A2: mov eax, var_50
  loc_00E1C5A5: mov [edx+00000004h], eax
  loc_00E1C5A8: mov eax, var_4C
  loc_00E1C5AB: mov [edx+00000008h], eax
  loc_00E1C5AE: mov eax, var_48
  loc_00E1C5B1: mov [edx+0000000Ch], eax
  loc_00E1C5B4: mov edx, var_D4
  loc_00E1C5BA: push edx
  loc_00E1C5BB: call [ecx+00000038h]
  loc_00E1C5BE: test eax, eax
  loc_00E1C5C0: fnclex
  loc_00E1C5C2: jge 00E1C5D5h
  loc_00E1C5C4: mov ecx, var_D4
  loc_00E1C5CA: push 00000038h
  loc_00E1C5CC: push 006E09F8h
  loc_00E1C5D1: push ecx
  loc_00E1C5D2: push eax
  loc_00E1C5D3: call ebx
  loc_00E1C5D5: lea edx, var_34
  loc_00E1C5D8: lea eax, var_30
  loc_00E1C5DB: push edx
  loc_00E1C5DC: lea ecx, var_2C
  loc_00E1C5DF: push eax
  loc_00E1C5E0: lea edx, var_28
  loc_00E1C5E3: push ecx
  loc_00E1C5E4: lea eax, var_24
  loc_00E1C5E7: push edx
  loc_00E1C5E8: push eax
  loc_00E1C5E9: push 00000005h
  loc_00E1C5EB: call [00401048h] ; __vbaFreeObjList
  loc_00E1C5F1: lea ecx, var_54
  loc_00E1C5F4: lea edx, var_44
  loc_00E1C5F7: push ecx
  loc_00E1C5F8: push edx
  loc_00E1C5F9: push 00000002h
  loc_00E1C5FB: call [00401038h] ; __vbaFreeVarList
  loc_00E1C601: mov eax, [esi]
  loc_00E1C603: add esp, 00000024h
  loc_00E1C606: push 006DCBD8h
  loc_00E1C60B: push 00000000h
  loc_00E1C60D: push 00000018h
  loc_00E1C60F: push esi
  loc_00E1C610: call [eax+00000564h]
  loc_00E1C616: lea ecx, var_28
  loc_00E1C619: push eax
  loc_00E1C61A: push ecx
  loc_00E1C61B: call edi
  loc_00E1C61D: lea edx, var_44
  loc_00E1C620: push eax
  loc_00E1C621: push edx
  loc_00E1C622: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1C628: add esp, 00000010h
  loc_00E1C62B: push eax
  loc_00E1C62C: call [00401120h] ; __vbaCastObjVar
  loc_00E1C632: push eax
  loc_00E1C633: lea eax, var_2C
  loc_00E1C636: push eax
  loc_00E1C637: call edi
  loc_00E1C639: mov ecx, [eax]
  loc_00E1C63B: lea edx, var_30
  loc_00E1C63E: push edx
  loc_00E1C63F: push eax
  loc_00E1C640: mov var_C4, eax
  loc_00E1C646: call [ecx+00000054h]
  loc_00E1C649: test eax, eax
  loc_00E1C64B: fnclex
  loc_00E1C64D: jge 00E1C660h
  loc_00E1C64F: mov ecx, var_C4
  loc_00E1C655: push 00000054h
  loc_00E1C657: push 006DCBE8h
  loc_00E1C65C: push ecx
  loc_00E1C65D: push eax
  loc_00E1C65E: call ebx
  loc_00E1C660: mov ecx, var_30
  loc_00E1C663: mov eax, 00000002h
  loc_00E1C668: mov var_7C, 00000012h
  loc_00E1C66F: mov var_84, eax
  loc_00E1C675: mov edx, [ecx]
  loc_00E1C677: mov var_CC, ecx
  loc_00E1C67D: lea ecx, var_34
  loc_00E1C680: push ecx
  loc_00E1C681: sub esp, 00000010h
  loc_00E1C684: mov ecx, esp
  loc_00E1C686: mov [ecx], eax
  loc_00E1C688: mov eax, var_80
  loc_00E1C68B: mov [ecx+00000004h], eax
  loc_00E1C68E: mov eax, var_7C
  loc_00E1C691: mov [ecx+00000008h], eax
  loc_00E1C694: mov eax, var_78
  loc_00E1C697: mov [ecx+0000000Ch], eax
  loc_00E1C69A: mov ecx, var_30
  loc_00E1C69D: push ecx
  loc_00E1C69E: call [edx+00000028h]
  loc_00E1C6A1: test eax, eax
  loc_00E1C6A3: fnclex
  loc_00E1C6A5: jge 00E1C6B8h
  loc_00E1C6A7: mov edx, var_CC
  loc_00E1C6AD: push 00000028h
  loc_00E1C6AF: push 006E09E8h
  loc_00E1C6B4: push edx
  loc_00E1C6B5: push eax
  loc_00E1C6B6: call ebx
  loc_00E1C6B8: mov eax, var_34
  loc_00E1C6BB: mov ecx, [esi]
  loc_00E1C6BD: push esi
  loc_00E1C6BE: mov var_D4, eax
  loc_00E1C6C4: call [ecx+00000388h]
  loc_00E1C6CA: lea edx, var_24
  loc_00E1C6CD: push eax
  loc_00E1C6CE: push edx
  loc_00E1C6CF: call edi
  loc_00E1C6D1: mov ecx, [eax]
  loc_00E1C6D3: lea edx, var_18
  loc_00E1C6D6: push edx
  loc_00E1C6D7: push eax
  loc_00E1C6D8: mov var_BC, eax
  loc_00E1C6DE: call [ecx+00000050h]
  loc_00E1C6E1: test eax, eax
  loc_00E1C6E3: fnclex
  loc_00E1C6E5: jge 00E1C6F8h
  loc_00E1C6E7: mov ecx, var_BC
  loc_00E1C6ED: push 00000050h
  loc_00E1C6EF: push 006E0410h
  loc_00E1C6F4: push ecx
  loc_00E1C6F5: push eax
  loc_00E1C6F6: call ebx
  loc_00E1C6F8: mov eax, var_18
  loc_00E1C6FB: mov edx, var_D4
  loc_00E1C701: mov var_4C, eax
  loc_00E1C704: mov eax, 00000008h
  loc_00E1C709: mov var_18, 00000000h
  loc_00E1C710: mov var_54, eax
  loc_00E1C713: mov ecx, [edx]
  loc_00E1C715: sub esp, 00000010h
  loc_00E1C718: mov edx, esp
  loc_00E1C71A: mov [edx], eax
  loc_00E1C71C: mov eax, var_50
  loc_00E1C71F: mov [edx+00000004h], eax
  loc_00E1C722: mov eax, var_4C
  loc_00E1C725: mov [edx+00000008h], eax
  loc_00E1C728: mov eax, var_48
  loc_00E1C72B: mov [edx+0000000Ch], eax
  loc_00E1C72E: mov edx, var_D4
  loc_00E1C734: push edx
  loc_00E1C735: call [ecx+00000038h]
  loc_00E1C738: test eax, eax
  loc_00E1C73A: fnclex
  loc_00E1C73C: jge 00E1C74Fh
  loc_00E1C73E: mov ecx, var_D4
  loc_00E1C744: push 00000038h
  loc_00E1C746: push 006E09F8h
  loc_00E1C74B: push ecx
  loc_00E1C74C: push eax
  loc_00E1C74D: call ebx
  loc_00E1C74F: lea edx, var_34
  loc_00E1C752: lea eax, var_30
  loc_00E1C755: push edx
  loc_00E1C756: lea ecx, var_2C
  loc_00E1C759: push eax
  loc_00E1C75A: lea edx, var_28
  loc_00E1C75D: push ecx
  loc_00E1C75E: lea eax, var_24
  loc_00E1C761: push edx
  loc_00E1C762: push eax
  loc_00E1C763: push 00000005h
  loc_00E1C765: call [00401048h] ; __vbaFreeObjList
  loc_00E1C76B: lea ecx, var_54
  loc_00E1C76E: lea edx, var_44
  loc_00E1C771: push ecx
  loc_00E1C772: push edx
  loc_00E1C773: push 00000002h
  loc_00E1C775: call [00401038h] ; __vbaFreeVarList
  loc_00E1C77B: mov eax, [esi]
  loc_00E1C77D: add esp, 00000024h
  loc_00E1C780: push 006DCBD8h
  loc_00E1C785: push 00000000h
  loc_00E1C787: push 00000018h
  loc_00E1C789: push esi
  loc_00E1C78A: call [eax+00000564h]
  loc_00E1C790: lea ecx, var_28
  loc_00E1C793: push eax
  loc_00E1C794: push ecx
  loc_00E1C795: call edi
  loc_00E1C797: lea edx, var_44
  loc_00E1C79A: push eax
  loc_00E1C79B: push edx
  loc_00E1C79C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1C7A2: add esp, 00000010h
  loc_00E1C7A5: push eax
  loc_00E1C7A6: call [00401120h] ; __vbaCastObjVar
  loc_00E1C7AC: push eax
  loc_00E1C7AD: lea eax, var_2C
  loc_00E1C7B0: push eax
  loc_00E1C7B1: call edi
  loc_00E1C7B3: mov ecx, [eax]
  loc_00E1C7B5: lea edx, var_30
  loc_00E1C7B8: push edx
  loc_00E1C7B9: push eax
  loc_00E1C7BA: mov var_C4, eax
  loc_00E1C7C0: call [ecx+00000054h]
  loc_00E1C7C3: test eax, eax
  loc_00E1C7C5: fnclex
  loc_00E1C7C7: jge 00E1C7DAh
  loc_00E1C7C9: mov ecx, var_C4
  loc_00E1C7CF: push 00000054h
  loc_00E1C7D1: push 006DCBE8h
  loc_00E1C7D6: push ecx
  loc_00E1C7D7: push eax
  loc_00E1C7D8: call ebx
  loc_00E1C7DA: mov ecx, var_30
  loc_00E1C7DD: mov eax, 00000002h
  loc_00E1C7E2: mov var_7C, 00000013h
  loc_00E1C7E9: mov var_84, eax
  loc_00E1C7EF: mov edx, [ecx]
  loc_00E1C7F1: mov var_CC, ecx
  loc_00E1C7F7: lea ecx, var_34
  loc_00E1C7FA: push ecx
  loc_00E1C7FB: sub esp, 00000010h
  loc_00E1C7FE: mov ecx, esp
  loc_00E1C800: mov [ecx], eax
  loc_00E1C802: mov eax, var_80
  loc_00E1C805: mov [ecx+00000004h], eax
  loc_00E1C808: mov eax, var_7C
  loc_00E1C80B: mov [ecx+00000008h], eax
  loc_00E1C80E: mov eax, var_78
  loc_00E1C811: mov [ecx+0000000Ch], eax
  loc_00E1C814: mov ecx, var_30
  loc_00E1C817: push ecx
  loc_00E1C818: call [edx+00000028h]
  loc_00E1C81B: test eax, eax
  loc_00E1C81D: fnclex
  loc_00E1C81F: jge 00E1C832h
  loc_00E1C821: mov edx, var_CC
  loc_00E1C827: push 00000028h
  loc_00E1C829: push 006E09E8h
  loc_00E1C82E: push edx
  loc_00E1C82F: push eax
  loc_00E1C830: call ebx
  loc_00E1C832: mov eax, var_34
  loc_00E1C835: mov ecx, [esi]
  loc_00E1C837: push esi
  loc_00E1C838: mov var_D4, eax
  loc_00E1C83E: call [ecx+00000378h]
  loc_00E1C844: lea edx, var_24
  loc_00E1C847: push eax
  loc_00E1C848: push edx
  loc_00E1C849: call edi
  loc_00E1C84B: mov ecx, [eax]
  loc_00E1C84D: lea edx, var_18
  loc_00E1C850: push edx
  loc_00E1C851: push eax
  loc_00E1C852: mov var_BC, eax
  loc_00E1C858: call [ecx+00000050h]
  loc_00E1C85B: test eax, eax
  loc_00E1C85D: fnclex
  loc_00E1C85F: jge 00E1C872h
  loc_00E1C861: mov ecx, var_BC
  loc_00E1C867: push 00000050h
  loc_00E1C869: push 006E0410h
  loc_00E1C86E: push ecx
  loc_00E1C86F: push eax
  loc_00E1C870: call ebx
  loc_00E1C872: mov eax, var_18
  loc_00E1C875: mov edx, var_D4
  loc_00E1C87B: mov var_4C, eax
  loc_00E1C87E: mov eax, 00000008h
  loc_00E1C883: mov var_18, 00000000h
  loc_00E1C88A: mov var_54, eax
  loc_00E1C88D: mov ecx, [edx]
  loc_00E1C88F: sub esp, 00000010h
  loc_00E1C892: mov edx, esp
  loc_00E1C894: mov [edx], eax
  loc_00E1C896: mov eax, var_50
  loc_00E1C899: mov [edx+00000004h], eax
  loc_00E1C89C: mov eax, var_4C
  loc_00E1C89F: mov [edx+00000008h], eax
  loc_00E1C8A2: mov eax, var_48
  loc_00E1C8A5: mov [edx+0000000Ch], eax
  loc_00E1C8A8: mov edx, var_D4
  loc_00E1C8AE: push edx
  loc_00E1C8AF: call [ecx+00000038h]
  loc_00E1C8B2: test eax, eax
  loc_00E1C8B4: fnclex
  loc_00E1C8B6: jge 00E1C8C9h
  loc_00E1C8B8: mov ecx, var_D4
  loc_00E1C8BE: push 00000038h
  loc_00E1C8C0: push 006E09F8h
  loc_00E1C8C5: push ecx
  loc_00E1C8C6: push eax
  loc_00E1C8C7: call ebx
  loc_00E1C8C9: lea edx, var_34
  loc_00E1C8CC: lea eax, var_30
  loc_00E1C8CF: push edx
  loc_00E1C8D0: lea ecx, var_2C
  loc_00E1C8D3: push eax
  loc_00E1C8D4: lea edx, var_28
  loc_00E1C8D7: push ecx
  loc_00E1C8D8: lea eax, var_24
  loc_00E1C8DB: push edx
  loc_00E1C8DC: push eax
  loc_00E1C8DD: push 00000005h
  loc_00E1C8DF: call [00401048h] ; __vbaFreeObjList
  loc_00E1C8E5: lea ecx, var_54
  loc_00E1C8E8: lea edx, var_44
  loc_00E1C8EB: push ecx
  loc_00E1C8EC: push edx
  loc_00E1C8ED: push 00000002h
  loc_00E1C8EF: call [00401038h] ; __vbaFreeVarList
  loc_00E1C8F5: mov eax, [esi]
  loc_00E1C8F7: add esp, 00000024h
  loc_00E1C8FA: push 006DCBD8h
  loc_00E1C8FF: push 00000000h
  loc_00E1C901: push 00000018h
  loc_00E1C903: push esi
  loc_00E1C904: call [eax+00000564h]
  loc_00E1C90A: lea ecx, var_28
  loc_00E1C90D: push eax
  loc_00E1C90E: push ecx
  loc_00E1C90F: call edi
  loc_00E1C911: lea edx, var_44
  loc_00E1C914: push eax
  loc_00E1C915: push edx
  loc_00E1C916: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1C91C: add esp, 00000010h
  loc_00E1C91F: push eax
  loc_00E1C920: call [00401120h] ; __vbaCastObjVar
  loc_00E1C926: push eax
  loc_00E1C927: lea eax, var_2C
  loc_00E1C92A: push eax
  loc_00E1C92B: call edi
  loc_00E1C92D: mov ecx, [eax]
  loc_00E1C92F: lea edx, var_30
  loc_00E1C932: push edx
  loc_00E1C933: push eax
  loc_00E1C934: mov var_C4, eax
  loc_00E1C93A: call [ecx+00000054h]
  loc_00E1C93D: test eax, eax
  loc_00E1C93F: fnclex
  loc_00E1C941: jge 00E1C954h
  loc_00E1C943: mov ecx, var_C4
  loc_00E1C949: push 00000054h
  loc_00E1C94B: push 006DCBE8h
  loc_00E1C950: push ecx
  loc_00E1C951: push eax
  loc_00E1C952: call ebx
  loc_00E1C954: mov ecx, var_30
  loc_00E1C957: mov eax, 00000002h
  loc_00E1C95C: mov var_7C, 00000014h
  loc_00E1C963: mov var_84, eax
  loc_00E1C969: mov edx, [ecx]
  loc_00E1C96B: mov var_CC, ecx
  loc_00E1C971: lea ecx, var_34
  loc_00E1C974: push ecx
  loc_00E1C975: sub esp, 00000010h
  loc_00E1C978: mov ecx, esp
  loc_00E1C97A: mov [ecx], eax
  loc_00E1C97C: mov eax, var_80
  loc_00E1C97F: mov [ecx+00000004h], eax
  loc_00E1C982: mov eax, var_7C
  loc_00E1C985: mov [ecx+00000008h], eax
  loc_00E1C988: mov eax, var_78
  loc_00E1C98B: mov [ecx+0000000Ch], eax
  loc_00E1C98E: mov ecx, var_30
  loc_00E1C991: push ecx
  loc_00E1C992: call [edx+00000028h]
  loc_00E1C995: test eax, eax
  loc_00E1C997: fnclex
  loc_00E1C999: jge 00E1C9ACh
  loc_00E1C99B: mov edx, var_CC
  loc_00E1C9A1: push 00000028h
  loc_00E1C9A3: push 006E09E8h
  loc_00E1C9A8: push edx
  loc_00E1C9A9: push eax
  loc_00E1C9AA: call ebx
  loc_00E1C9AC: mov eax, var_34
  loc_00E1C9AF: mov ecx, [esi]
  loc_00E1C9B1: push esi
  loc_00E1C9B2: mov var_D4, eax
  loc_00E1C9B8: call [ecx+00000340h]
  loc_00E1C9BE: lea edx, var_24
  loc_00E1C9C1: push eax
  loc_00E1C9C2: push edx
  loc_00E1C9C3: call edi
  loc_00E1C9C5: mov ecx, [eax]
  loc_00E1C9C7: lea edx, var_18
  loc_00E1C9CA: push edx
  loc_00E1C9CB: push eax
  loc_00E1C9CC: mov var_BC, eax
  loc_00E1C9D2: call [ecx+00000050h]
  loc_00E1C9D5: test eax, eax
  loc_00E1C9D7: fnclex
  loc_00E1C9D9: jge 00E1C9ECh
  loc_00E1C9DB: mov ecx, var_BC
  loc_00E1C9E1: push 00000050h
  loc_00E1C9E3: push 006E0410h
  loc_00E1C9E8: push ecx
  loc_00E1C9E9: push eax
  loc_00E1C9EA: call ebx
  loc_00E1C9EC: mov eax, var_18
  loc_00E1C9EF: mov edx, var_D4
  loc_00E1C9F5: mov var_4C, eax
  loc_00E1C9F8: mov eax, 00000008h
  loc_00E1C9FD: mov var_18, 00000000h
  loc_00E1CA04: mov var_54, eax
  loc_00E1CA07: mov ecx, [edx]
  loc_00E1CA09: sub esp, 00000010h
  loc_00E1CA0C: mov edx, esp
  loc_00E1CA0E: mov [edx], eax
  loc_00E1CA10: mov eax, var_50
  loc_00E1CA13: mov [edx+00000004h], eax
  loc_00E1CA16: mov eax, var_4C
  loc_00E1CA19: mov [edx+00000008h], eax
  loc_00E1CA1C: mov eax, var_48
  loc_00E1CA1F: mov [edx+0000000Ch], eax
  loc_00E1CA22: mov edx, var_D4
  loc_00E1CA28: push edx
  loc_00E1CA29: call [ecx+00000038h]
  loc_00E1CA2C: test eax, eax
  loc_00E1CA2E: fnclex
  loc_00E1CA30: jge 00E1CA43h
  loc_00E1CA32: mov ecx, var_D4
  loc_00E1CA38: push 00000038h
  loc_00E1CA3A: push 006E09F8h
  loc_00E1CA3F: push ecx
  loc_00E1CA40: push eax
  loc_00E1CA41: call ebx
  loc_00E1CA43: lea edx, var_34
  loc_00E1CA46: lea eax, var_30
  loc_00E1CA49: push edx
  loc_00E1CA4A: lea ecx, var_2C
  loc_00E1CA4D: push eax
  loc_00E1CA4E: lea edx, var_28
  loc_00E1CA51: push ecx
  loc_00E1CA52: lea eax, var_24
  loc_00E1CA55: push edx
  loc_00E1CA56: push eax
  loc_00E1CA57: push 00000005h
  loc_00E1CA59: call [00401048h] ; __vbaFreeObjList
  loc_00E1CA5F: lea ecx, var_54
  loc_00E1CA62: lea edx, var_44
  loc_00E1CA65: push ecx
  loc_00E1CA66: push edx
  loc_00E1CA67: push 00000002h
  loc_00E1CA69: call [00401038h] ; __vbaFreeVarList
  loc_00E1CA6F: mov eax, [esi]
  loc_00E1CA71: add esp, 00000024h
  loc_00E1CA74: push 006DCBD8h
  loc_00E1CA79: push 00000000h
  loc_00E1CA7B: push 00000018h
  loc_00E1CA7D: push esi
  loc_00E1CA7E: call [eax+00000564h]
  loc_00E1CA84: lea ecx, var_28
  loc_00E1CA87: push eax
  loc_00E1CA88: push ecx
  loc_00E1CA89: call edi
  loc_00E1CA8B: lea edx, var_44
  loc_00E1CA8E: push eax
  loc_00E1CA8F: push edx
  loc_00E1CA90: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1CA96: add esp, 00000010h
  loc_00E1CA99: push eax
  loc_00E1CA9A: call [00401120h] ; __vbaCastObjVar
  loc_00E1CAA0: push eax
  loc_00E1CAA1: lea eax, var_2C
  loc_00E1CAA4: push eax
  loc_00E1CAA5: call edi
  loc_00E1CAA7: mov ecx, [eax]
  loc_00E1CAA9: lea edx, var_30
  loc_00E1CAAC: push edx
  loc_00E1CAAD: push eax
  loc_00E1CAAE: mov var_C4, eax
  loc_00E1CAB4: call [ecx+00000054h]
  loc_00E1CAB7: test eax, eax
  loc_00E1CAB9: fnclex
  loc_00E1CABB: jge 00E1CACEh
  loc_00E1CABD: mov ecx, var_C4
  loc_00E1CAC3: push 00000054h
  loc_00E1CAC5: push 006DCBE8h
  loc_00E1CACA: push ecx
  loc_00E1CACB: push eax
  loc_00E1CACC: call ebx
  loc_00E1CACE: mov ecx, var_30
  loc_00E1CAD1: mov eax, 00000002h
  loc_00E1CAD6: mov var_7C, 00000015h
  loc_00E1CADD: mov var_84, eax
  loc_00E1CAE3: mov edx, [ecx]
  loc_00E1CAE5: mov var_CC, ecx
  loc_00E1CAEB: lea ecx, var_34
  loc_00E1CAEE: push ecx
  loc_00E1CAEF: sub esp, 00000010h
  loc_00E1CAF2: mov ecx, esp
  loc_00E1CAF4: mov [ecx], eax
  loc_00E1CAF6: mov eax, var_80
  loc_00E1CAF9: mov [ecx+00000004h], eax
  loc_00E1CAFC: mov eax, var_7C
  loc_00E1CAFF: mov [ecx+00000008h], eax
  loc_00E1CB02: mov eax, var_78
  loc_00E1CB05: mov [ecx+0000000Ch], eax
  loc_00E1CB08: mov ecx, var_30
  loc_00E1CB0B: push ecx
  loc_00E1CB0C: call [edx+00000028h]
  loc_00E1CB0F: test eax, eax
  loc_00E1CB11: fnclex
  loc_00E1CB13: jge 00E1CB26h
  loc_00E1CB15: mov edx, var_CC
  loc_00E1CB1B: push 00000028h
  loc_00E1CB1D: push 006E09E8h
  loc_00E1CB22: push edx
  loc_00E1CB23: push eax
  loc_00E1CB24: call ebx
  loc_00E1CB26: mov eax, var_34
  loc_00E1CB29: mov ecx, [esi]
  loc_00E1CB2B: push esi
  loc_00E1CB2C: mov var_D4, eax
  loc_00E1CB32: call [ecx+0000033Ch]
  loc_00E1CB38: lea edx, var_24
  loc_00E1CB3B: push eax
  loc_00E1CB3C: push edx
  loc_00E1CB3D: call edi
  loc_00E1CB3F: mov ecx, [eax]
  loc_00E1CB41: lea edx, var_18
  loc_00E1CB44: push edx
  loc_00E1CB45: push eax
  loc_00E1CB46: mov var_BC, eax
  loc_00E1CB4C: call [ecx+00000050h]
  loc_00E1CB4F: test eax, eax
  loc_00E1CB51: fnclex
  loc_00E1CB53: jge 00E1CB66h
  loc_00E1CB55: mov ecx, var_BC
  loc_00E1CB5B: push 00000050h
  loc_00E1CB5D: push 006E0410h
  loc_00E1CB62: push ecx
  loc_00E1CB63: push eax
  loc_00E1CB64: call ebx
  loc_00E1CB66: mov eax, var_18
  loc_00E1CB69: mov edx, var_D4
  loc_00E1CB6F: mov var_4C, eax
  loc_00E1CB72: mov eax, 00000008h
  loc_00E1CB77: mov var_18, 00000000h
  loc_00E1CB7E: mov var_54, eax
  loc_00E1CB81: mov ecx, [edx]
  loc_00E1CB83: sub esp, 00000010h
  loc_00E1CB86: mov edx, esp
  loc_00E1CB88: mov [edx], eax
  loc_00E1CB8A: mov eax, var_50
  loc_00E1CB8D: mov [edx+00000004h], eax
  loc_00E1CB90: mov eax, var_4C
  loc_00E1CB93: mov [edx+00000008h], eax
  loc_00E1CB96: mov eax, var_48
  loc_00E1CB99: mov [edx+0000000Ch], eax
  loc_00E1CB9C: mov edx, var_D4
  loc_00E1CBA2: push edx
  loc_00E1CBA3: call [ecx+00000038h]
  loc_00E1CBA6: test eax, eax
  loc_00E1CBA8: fnclex
  loc_00E1CBAA: jge 00E1CBBDh
  loc_00E1CBAC: mov ecx, var_D4
  loc_00E1CBB2: push 00000038h
  loc_00E1CBB4: push 006E09F8h
  loc_00E1CBB9: push ecx
  loc_00E1CBBA: push eax
  loc_00E1CBBB: call ebx
  loc_00E1CBBD: lea edx, var_34
  loc_00E1CBC0: lea eax, var_30
  loc_00E1CBC3: push edx
  loc_00E1CBC4: lea ecx, var_2C
  loc_00E1CBC7: push eax
  loc_00E1CBC8: lea edx, var_28
  loc_00E1CBCB: push ecx
  loc_00E1CBCC: lea eax, var_24
  loc_00E1CBCF: push edx
  loc_00E1CBD0: push eax
  loc_00E1CBD1: push 00000005h
  loc_00E1CBD3: call [00401048h] ; __vbaFreeObjList
  loc_00E1CBD9: lea ecx, var_54
  loc_00E1CBDC: lea edx, var_44
  loc_00E1CBDF: push ecx
  loc_00E1CBE0: push edx
  loc_00E1CBE1: push 00000002h
  loc_00E1CBE3: call [00401038h] ; __vbaFreeVarList
  loc_00E1CBE9: mov eax, [esi]
  loc_00E1CBEB: add esp, 00000024h
  loc_00E1CBEE: push esi
  loc_00E1CBEF: call [eax+000004D4h]
  loc_00E1CBF5: lea ecx, var_24
  loc_00E1CBF8: push eax
  loc_00E1CBF9: push ecx
  loc_00E1CBFA: call edi
  loc_00E1CBFC: mov edx, [eax]
  loc_00E1CBFE: push 00000000h
  loc_00E1CC00: push eax
  loc_00E1CC01: mov var_BC, eax
  loc_00E1CC07: call [edx+0000008Ch]
  loc_00E1CC0D: test eax, eax
  loc_00E1CC0F: fnclex
  loc_00E1CC11: jge 00E1CC27h
  loc_00E1CC13: mov ecx, var_BC
  loc_00E1CC19: push 0000008Ch
  loc_00E1CC1E: push 006DCDA0h
  loc_00E1CC23: push ecx
  loc_00E1CC24: push eax
  loc_00E1CC25: call ebx
  loc_00E1CC27: lea ecx, var_24
  loc_00E1CC2A: call [00401254h] ; __vbaFreeObj
  loc_00E1CC30: mov edx, [esi]
  loc_00E1CC32: push 006DCBD8h
  loc_00E1CC37: push 00000000h
  loc_00E1CC39: push 00000018h
  loc_00E1CC3B: push esi
  loc_00E1CC3C: call [edx+00000564h]
  loc_00E1CC42: push eax
  loc_00E1CC43: lea eax, var_24
  loc_00E1CC46: push eax
  loc_00E1CC47: call edi
  loc_00E1CC49: lea ecx, var_44
  loc_00E1CC4C: push eax
  loc_00E1CC4D: push ecx
  loc_00E1CC4E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1CC54: add esp, 00000010h
  loc_00E1CC57: push eax
  loc_00E1CC58: call [00401120h] ; __vbaCastObjVar
  loc_00E1CC5E: lea edx, var_28
  loc_00E1CC61: push eax
  loc_00E1CC62: push edx
  loc_00E1CC63: call edi
  loc_00E1CC65: sub esp, 00000010h
  loc_00E1CC68: mov ecx, 0000000Ah
  loc_00E1CC6D: mov edx, esp
  loc_00E1CC6F: mov var_94, ecx
  loc_00E1CC75: mov var_84, ecx
  loc_00E1CC7B: mov var_8C, 80020004h
  loc_00E1CC85: mov [edx], ecx
  loc_00E1CC87: mov ecx, var_90
  loc_00E1CC8D: sub esp, 00000010h
  loc_00E1CC90: mov var_7C, 80020004h
  loc_00E1CC97: mov [edx+00000004h], ecx
  loc_00E1CC9A: mov ecx, var_8C
  loc_00E1CCA0: mov var_BC, eax
  loc_00E1CCA6: mov eax, [eax]
  loc_00E1CCA8: mov [edx+00000008h], ecx
  loc_00E1CCAB: mov ecx, var_88
  loc_00E1CCB1: mov [edx+0000000Ch], ecx
  loc_00E1CCB4: mov ecx, var_84
  loc_00E1CCBA: mov edx, esp
  loc_00E1CCBC: mov [edx], ecx
  loc_00E1CCBE: mov ecx, var_80
  loc_00E1CCC1: mov [edx+00000004h], ecx
  loc_00E1CCC4: mov ecx, var_7C
  loc_00E1CCC7: mov [edx+00000008h], ecx
  loc_00E1CCCA: mov ecx, var_78
  loc_00E1CCCD: mov [edx+0000000Ch], ecx
  loc_00E1CCD0: mov edx, var_BC
  loc_00E1CCD6: push edx
  loc_00E1CCD7: call [eax+000000ACh]
  loc_00E1CCDD: test eax, eax
  loc_00E1CCDF: fnclex
  loc_00E1CCE1: jge 00E1CCF7h
  loc_00E1CCE3: mov ecx, var_BC
  loc_00E1CCE9: push 000000ACh
  loc_00E1CCEE: push 006DCBE8h
  loc_00E1CCF3: push ecx
  loc_00E1CCF4: push eax
  loc_00E1CCF5: call ebx
  loc_00E1CCF7: lea edx, var_28
  loc_00E1CCFA: lea eax, var_24
  loc_00E1CCFD: push edx
  loc_00E1CCFE: push eax
  loc_00E1CCFF: push 00000002h
  loc_00E1CD01: call [00401048h] ; __vbaFreeObjList
  loc_00E1CD07: add esp, 0000000Ch
  loc_00E1CD0A: lea ecx, var_44
  loc_00E1CD0D: call [00401024h] ; __vbaFreeVar
  loc_00E1CD13: mov ecx, [esi]
  loc_00E1CD15: push 00000000h
  loc_00E1CD17: push 00000019h
  loc_00E1CD19: push esi
  loc_00E1CD1A: call [ecx+00000564h]
  loc_00E1CD20: lea edx, var_24
  loc_00E1CD23: push eax
  loc_00E1CD24: push edx
  loc_00E1CD25: call edi
  loc_00E1CD27: push eax
  loc_00E1CD28: call [00401030h] ; __vbaLateIdCall
  loc_00E1CD2E: add esp, 0000000Ch
  loc_00E1CD31: lea ecx, var_24
  loc_00E1CD34: call [00401254h] ; __vbaFreeObj
  loc_00E1CD3A: mov eax, [esi]
  loc_00E1CD3C: push esi
  loc_00E1CD3D: call [eax+0000031Ch]
  loc_00E1CD43: lea ecx, var_24
  loc_00E1CD46: push eax
  loc_00E1CD47: push ecx
  loc_00E1CD48: call edi
  loc_00E1CD4A: mov edx, [eax]
  loc_00E1CD4C: push FFFFFFFFh
  loc_00E1CD4E: push eax
  loc_00E1CD4F: mov var_BC, eax
  loc_00E1CD55: call [edx+0000009Ch]
  loc_00E1CD5B: test eax, eax
  loc_00E1CD5D: fnclex
  loc_00E1CD5F: jge 00E1CD75h
  loc_00E1CD61: mov ecx, var_BC
  loc_00E1CD67: push 0000009Ch
  loc_00E1CD6C: push 006DCAD0h
  loc_00E1CD71: push ecx
  loc_00E1CD72: push eax
  loc_00E1CD73: call ebx
  loc_00E1CD75: lea ecx, var_24
  loc_00E1CD78: call [00401254h] ; __vbaFreeObj
  loc_00E1CD7E: mov edx, [esi]
  loc_00E1CD80: push esi
  loc_00E1CD81: call [edx+000002FCh]
  loc_00E1CD87: push eax
  loc_00E1CD88: lea eax, var_24
  loc_00E1CD8B: push eax
  loc_00E1CD8C: call edi
  loc_00E1CD8E: mov ecx, [eax]
  loc_00E1CD90: push FFFFFFFFh
  loc_00E1CD92: push eax
  loc_00E1CD93: mov var_BC, eax
  loc_00E1CD99: call [ecx+0000005Ch]
  loc_00E1CD9C: test eax, eax
  loc_00E1CD9E: fnclex
  loc_00E1CDA0: jge 00E1CDB3h
  loc_00E1CDA2: mov edx, var_BC
  loc_00E1CDA8: push 0000005Ch
  loc_00E1CDAA: push 006DCAE0h
  loc_00E1CDAF: push edx
  loc_00E1CDB0: push eax
  loc_00E1CDB1: call ebx
  loc_00E1CDB3: lea ecx, var_24
  loc_00E1CDB6: call [00401254h] ; __vbaFreeObj
  loc_00E1CDBC: mov eax, [esi]
  loc_00E1CDBE: push esi
  loc_00E1CDBF: call [eax+0000031Ch]
  loc_00E1CDC5: lea ecx, var_24
  loc_00E1CDC8: push eax
  loc_00E1CDC9: push ecx
  loc_00E1CDCA: call edi
  loc_00E1CDCC: mov esi, eax
  loc_00E1CDCE: push FFFFFFFFh
  loc_00E1CDD0: push esi
  loc_00E1CDD1: mov edx, [esi]
  loc_00E1CDD3: call [edx+0000009Ch]
  loc_00E1CDD9: test eax, eax
  loc_00E1CDDB: fnclex
  loc_00E1CDDD: jge 00E1CEB5h
  loc_00E1CDE3: push 0000009Ch
  loc_00E1CDE8: push 006DCAD0h
  loc_00E1CDED: push esi
  loc_00E1CDEE: push eax
  loc_00E1CDEF: call ebx
  loc_00E1CDF1: jmp 00E1CEB5h
  loc_00E1CDF6: mov ebx, [004011F0h] ; __vbaVarDup
  loc_00E1CDFC: mov ecx, 80020004h
  loc_00E1CE01: mov var_6C, ecx
  loc_00E1CE04: mov eax, 0000000Ah
  loc_00E1CE09: mov var_5C, ecx
  loc_00E1CE0C: lea edx, var_94
  loc_00E1CE12: lea ecx, var_54
  loc_00E1CE15: mov var_74, eax
  loc_00E1CE18: mov var_64, eax
  loc_00E1CE1B: mov var_8C, 006E12A4h ; "INFORMATION 002"
  loc_00E1CE25: mov var_94, 00000008h
  loc_00E1CE2F: call ebx
  loc_00E1CE31: lea edx, var_84
  loc_00E1CE37: lea ecx, var_44
  loc_00E1CE3A: mov var_7C, 006E0BD4h ; "Data Sudah Ada !"
  loc_00E1CE41: mov var_84, 00000008h
  loc_00E1CE4B: call ebx
  loc_00E1CE4D: lea eax, var_74
  loc_00E1CE50: lea ecx, var_64
  loc_00E1CE53: push eax
  loc_00E1CE54: lea edx, var_54
  loc_00E1CE57: push ecx
  loc_00E1CE58: push edx
  loc_00E1CE59: lea eax, var_44
  loc_00E1CE5C: push 00000040h
  loc_00E1CE5E: push eax
  loc_00E1CE5F: call [004010A8h] ; rtcMsgBox
  loc_00E1CE65: lea ecx, var_74
  loc_00E1CE68: lea edx, var_64
  loc_00E1CE6B: push ecx
  loc_00E1CE6C: lea eax, var_54
  loc_00E1CE6F: push edx
  loc_00E1CE70: lea ecx, var_44
  loc_00E1CE73: push eax
  loc_00E1CE74: push ecx
  loc_00E1CE75: push 00000004h
  loc_00E1CE77: call [00401038h] ; __vbaFreeVarList
  loc_00E1CE7D: mov edx, [esi]
  loc_00E1CE7F: add esp, 00000014h
  loc_00E1CE82: push esi
  loc_00E1CE83: call [edx+0000031Ch]
  loc_00E1CE89: push eax
  loc_00E1CE8A: lea eax, var_24
  loc_00E1CE8D: push eax
  loc_00E1CE8E: call edi
  loc_00E1CE90: mov esi, eax
  loc_00E1CE92: push FFFFFFFFh
  loc_00E1CE94: push esi
  loc_00E1CE95: mov ecx, [esi]
  loc_00E1CE97: call [ecx+0000009Ch]
  loc_00E1CE9D: test eax, eax
  loc_00E1CE9F: fnclex
  loc_00E1CEA1: jge 00E1CEB5h
  loc_00E1CEA3: push 0000009Ch
  loc_00E1CEA8: push 006DCAD0h
  loc_00E1CEAD: push esi
  loc_00E1CEAE: push eax
  loc_00E1CEAF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1CEB5: lea ecx, var_24
  loc_00E1CEB8: call [00401254h] ; __vbaFreeObj
  loc_00E1CEBE: mov var_4, 00000000h
  loc_00E1CEC5: fwait
  loc_00E1CEC6: push 00E1CF1Ah
  loc_00E1CECB: jmp 00E1CF19h
  loc_00E1CECD: lea edx, var_20
  loc_00E1CED0: lea eax, var_1C
  loc_00E1CED3: push edx
  loc_00E1CED4: lea ecx, var_18
  loc_00E1CED7: push eax
  loc_00E1CED8: push ecx
  loc_00E1CED9: push 00000003h
  loc_00E1CEDB: call [004011BCh] ; __vbaFreeStrList
  loc_00E1CEE1: lea edx, var_34
  loc_00E1CEE4: lea eax, var_30
  loc_00E1CEE7: push edx
  loc_00E1CEE8: lea ecx, var_2C
  loc_00E1CEEB: push eax
  loc_00E1CEEC: lea edx, var_28
  loc_00E1CEEF: push ecx
  loc_00E1CEF0: lea eax, var_24
  loc_00E1CEF3: push edx
  loc_00E1CEF4: push eax
  loc_00E1CEF5: push 00000005h
  loc_00E1CEF7: call [00401048h] ; __vbaFreeObjList
  loc_00E1CEFD: lea ecx, var_74
  loc_00E1CF00: lea edx, var_64
  loc_00E1CF03: push ecx
  loc_00E1CF04: lea eax, var_54
  loc_00E1CF07: push edx
  loc_00E1CF08: lea ecx, var_44
  loc_00E1CF0B: push eax
  loc_00E1CF0C: push ecx
  loc_00E1CF0D: push 00000004h
  loc_00E1CF0F: call [00401038h] ; __vbaFreeVarList
  loc_00E1CF15: add esp, 0000003Ch
  loc_00E1CF18: ret
  loc_00E1CF19: ret
  loc_00E1CF1A: mov eax, Me
  loc_00E1CF1D: push eax
  loc_00E1CF1E: mov edx, [eax]
  loc_00E1CF20: call [edx+00000008h]
  loc_00E1CF23: mov eax, var_4
  loc_00E1CF26: mov ecx, var_14
  loc_00E1CF29: pop edi
  loc_00E1CF2A: pop esi
  loc_00E1CF2B: mov fs:[00000000h], ecx
  loc_00E1CF32: pop ebx
  loc_00E1CF33: mov esp, ebp
  loc_00E1CF35: pop ebp
  loc_00E1CF36: retn 0004h
End Sub

Private Sub cari_UnknownEvent_9 'E18E30
  loc_00E18E30: push ebp
  loc_00E18E31: mov ebp, esp
  loc_00E18E33: sub esp, 0000000Ch
  loc_00E18E36: push 00402836h ; __vbaExceptHandler
  loc_00E18E3B: mov eax, fs:[00000000h]
  loc_00E18E41: push eax
  loc_00E18E42: mov fs:[00000000h], esp
  loc_00E18E49: sub esp, 00000034h
  loc_00E18E4C: push ebx
  loc_00E18E4D: push esi
  loc_00E18E4E: push edi
  loc_00E18E4F: mov var_C, esp
  loc_00E18E52: mov var_8, 00402320h
  loc_00E18E59: mov esi, Me
  loc_00E18E5C: mov eax, esi
  loc_00E18E5E: and eax, 00000001h
  loc_00E18E61: mov var_4, eax
  loc_00E18E64: and esi, FFFFFFFEh
  loc_00E18E67: push esi
  loc_00E18E68: mov Me, esi
  loc_00E18E6B: mov ecx, [esi]
  loc_00E18E6D: call [ecx+00000004h]
  loc_00E18E70: mov edx, [esi]
  loc_00E18E72: push esi
  loc_00E18E73: mov var_18, 00000000h
  loc_00E18E7A: call [edx+00000304h]
  loc_00E18E80: push eax
  loc_00E18E81: lea eax, var_18
  loc_00E18E84: push eax
  loc_00E18E85: call [004010ACh] ; __vbaObjSet
  loc_00E18E8B: mov edi, eax
  loc_00E18E8D: push FFFFFFFFh
  loc_00E18E8F: push edi
  loc_00E18E90: mov ecx, [edi]
  loc_00E18E92: call [ecx+0000009Ch]
  loc_00E18E98: test eax, eax
  loc_00E18E9A: fnclex
  loc_00E18E9C: jge 00E18EB0h
  loc_00E18E9E: push 0000009Ch
  loc_00E18EA3: push 006DCAD0h
  loc_00E18EA8: push edi
  loc_00E18EA9: push eax
  loc_00E18EAA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E18EB0: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E18EB6: lea ecx, var_18
  loc_00E18EB9: call ebx
  loc_00E18EBB: mov edx, [esi]
  loc_00E18EBD: push esi
  loc_00E18EBE: call [edx+0000032Ch]
  loc_00E18EC4: push eax
  loc_00E18EC5: lea eax, var_18
  loc_00E18EC8: push eax
  loc_00E18EC9: call [004010ACh] ; __vbaObjSet
  loc_00E18ECF: mov edi, eax
  loc_00E18ED1: push FFFFFFFFh
  loc_00E18ED3: push edi
  loc_00E18ED4: mov ecx, [edi]
  loc_00E18ED6: call [ecx+0000008Ch]
  loc_00E18EDC: test eax, eax
  loc_00E18EDE: fnclex
  loc_00E18EE0: jge 00E18EF4h
  loc_00E18EE2: push 0000008Ch
  loc_00E18EE7: push 006DCDA0h
  loc_00E18EEC: push edi
  loc_00E18EED: push eax
  loc_00E18EEE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E18EF4: lea ecx, var_18
  loc_00E18EF7: call ebx
  loc_00E18EF9: sub esp, 00000010h
  loc_00E18EFC: mov ecx, 0000000Bh
  loc_00E18F01: mov edx, esp
  loc_00E18F03: xor eax, eax
  loc_00E18F05: push 8001000Dh
  loc_00E18F0A: push esi
  loc_00E18F0B: mov [edx], ecx
  loc_00E18F0D: mov ecx, var_24
  loc_00E18F10: mov [edx+00000004h], ecx
  loc_00E18F13: mov ecx, [esi]
  loc_00E18F15: mov [edx+00000008h], eax
  loc_00E18F18: mov eax, var_1C
  loc_00E18F1B: mov [edx+0000000Ch], eax
  loc_00E18F1E: call [ecx+000004C4h]
  loc_00E18F24: lea edx, var_18
  loc_00E18F27: push eax
  loc_00E18F28: push edx
  loc_00E18F29: call [004010ACh] ; __vbaObjSet
  loc_00E18F2F: push eax
  loc_00E18F30: call [00401238h] ; __vbaLateIdSt
  loc_00E18F36: lea ecx, var_18
  loc_00E18F39: call ebx
  loc_00E18F3B: mov var_4, 00000000h
  loc_00E18F42: push 00E18F54h
  loc_00E18F47: jmp 00E18F53h
  loc_00E18F49: lea ecx, var_18
  loc_00E18F4C: call [00401254h] ; __vbaFreeObj
  loc_00E18F52: ret
  loc_00E18F53: ret
  loc_00E18F54: mov eax, Me
  loc_00E18F57: push eax
  loc_00E18F58: mov ecx, [eax]
  loc_00E18F5A: call [ecx+00000008h]
  loc_00E18F5D: mov eax, var_4
  loc_00E18F60: mov ecx, var_14
  loc_00E18F63: pop edi
  loc_00E18F64: pop esi
  loc_00E18F65: mov fs:[00000000h], ecx
  loc_00E18F6C: pop ebx
  loc_00E18F6D: mov esp, ebp
  loc_00E18F6F: pop ebp
  loc_00E18F70: retn 0004h
End Sub

Private Sub totalkan_UnknownEvent_9 'E1D780
  loc_00E1D780: push ebp
  loc_00E1D781: mov ebp, esp
  loc_00E1D783: sub esp, 0000000Ch
  loc_00E1D786: push 00402836h ; __vbaExceptHandler
  loc_00E1D78B: mov eax, fs:[00000000h]
  loc_00E1D791: push eax
  loc_00E1D792: mov fs:[00000000h], esp
  loc_00E1D799: sub esp, 000001E8h
  loc_00E1D79F: push ebx
  loc_00E1D7A0: push esi
  loc_00E1D7A1: push edi
  loc_00E1D7A2: mov var_C, esp
  loc_00E1D7A5: mov var_8, 004023F0h
  loc_00E1D7AC: mov esi, Me
  loc_00E1D7AF: mov eax, esi
  loc_00E1D7B1: and eax, 00000001h
  loc_00E1D7B4: mov var_4, eax
  loc_00E1D7B7: and esi, FFFFFFFEh
  loc_00E1D7BA: push esi
  loc_00E1D7BB: mov Me, esi
  loc_00E1D7BE: mov ecx, [esi]
  loc_00E1D7C0: call [ecx+00000004h]
  loc_00E1D7C3: mov edx, [esi]
  loc_00E1D7C5: xor ebx, ebx
  loc_00E1D7C7: push esi
  loc_00E1D7C8: mov var_18, ebx
  loc_00E1D7CB: mov var_1C, ebx
  loc_00E1D7CE: mov var_20, ebx
  loc_00E1D7D1: mov var_24, ebx
  loc_00E1D7D4: mov var_28, ebx
  loc_00E1D7D7: mov var_2C, ebx
  loc_00E1D7DA: mov var_30, ebx
  loc_00E1D7DD: mov var_34, ebx
  loc_00E1D7E0: mov var_38, ebx
  loc_00E1D7E3: mov var_3C, ebx
  loc_00E1D7E6: mov var_40, ebx
  loc_00E1D7E9: mov var_44, ebx
  loc_00E1D7EC: mov var_48, ebx
  loc_00E1D7EF: mov var_4C, ebx
  loc_00E1D7F2: mov var_50, ebx
  loc_00E1D7F5: mov var_54, ebx
  loc_00E1D7F8: mov var_58, ebx
  loc_00E1D7FB: mov var_5C, ebx
  loc_00E1D7FE: mov var_60, ebx
  loc_00E1D801: mov var_64, ebx
  loc_00E1D804: mov var_68, ebx
  loc_00E1D807: mov var_6C, ebx
  loc_00E1D80A: mov var_70, ebx
  loc_00E1D80D: mov var_74, ebx
  loc_00E1D810: mov var_78, ebx
  loc_00E1D813: mov var_7C, ebx
  loc_00E1D816: mov var_80, ebx
  loc_00E1D819: mov var_84, ebx
  loc_00E1D81F: mov var_88, ebx
  loc_00E1D825: mov var_8C, ebx
  loc_00E1D82B: mov var_9C, ebx
  loc_00E1D831: mov var_AC, ebx
  loc_00E1D837: mov var_BC, ebx
  loc_00E1D83D: mov var_CC, ebx
  loc_00E1D843: mov var_DC, ebx
  loc_00E1D849: mov var_EC, ebx
  loc_00E1D84F: call [edx+00000340h]
  loc_00E1D855: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E1D85B: push eax
  loc_00E1D85C: lea eax, var_54
  loc_00E1D85F: push eax
  loc_00E1D860: call edi
  loc_00E1D862: mov esi, eax
  loc_00E1D864: lea edx, var_18
  loc_00E1D867: push edx
  loc_00E1D868: push esi
  loc_00E1D869: mov ecx, [esi]
  loc_00E1D86B: call [ecx+00000050h]
  loc_00E1D86E: cmp eax, ebx
  loc_00E1D870: fnclex
  loc_00E1D872: jge 00E1D883h
  loc_00E1D874: push 00000050h
  loc_00E1D876: push 006E0410h
  loc_00E1D87B: push esi
  loc_00E1D87C: push eax
  loc_00E1D87D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D883: mov eax, var_18
  loc_00E1D886: push eax
  loc_00E1D887: push 006E0934h
  loc_00E1D88C: call [00401104h] ; __vbaStrCmp
  loc_00E1D892: mov esi, eax
  loc_00E1D894: lea ecx, var_18
  loc_00E1D897: neg esi
  loc_00E1D899: sbb esi, esi
  loc_00E1D89B: inc esi
  loc_00E1D89C: neg esi
  loc_00E1D89E: call [00401258h] ; __vbaFreeStr
  loc_00E1D8A4: lea ecx, var_54
  loc_00E1D8A7: call [00401254h] ; __vbaFreeObj
  loc_00E1D8AD: cmp si, bx
  loc_00E1D8B0: jz 00E1D9B0h
  loc_00E1D8B6: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E1D8BC: mov ecx, 80020004h
  loc_00E1D8C1: mov var_C4, ecx
  loc_00E1D8C7: mov eax, 0000000Ah
  loc_00E1D8CC: mov var_B4, ecx
  loc_00E1D8D2: mov ebx, 00000008h
  loc_00E1D8D7: lea edx, var_EC
  loc_00E1D8DD: lea ecx, var_AC
  loc_00E1D8E3: mov var_CC, eax
  loc_00E1D8E9: mov var_BC, eax
  loc_00E1D8EF: mov var_E4, 006E0B08h ; "Need To Do"
  loc_00E1D8F9: mov var_EC, ebx
  loc_00E1D8FF: call __vbaVarDup
  loc_00E1D901: lea edx, var_DC
  loc_00E1D907: lea ecx, var_9C
  loc_00E1D90D: mov var_D4, 006E12FCh ; "Anda Belum Memasukkan Data Calon Peserta Didik, Silahkan Input sebelum Melakukan penyimpanan !"
  loc_00E1D917: mov var_DC, ebx
  loc_00E1D91D: call __vbaVarDup
  loc_00E1D91F: lea ecx, var_CC
  loc_00E1D925: lea edx, var_BC
  loc_00E1D92B: push ecx
  loc_00E1D92C: lea eax, var_AC
  loc_00E1D932: push edx
  loc_00E1D933: push eax
  loc_00E1D934: lea ecx, var_9C
  loc_00E1D93A: push 00000040h
  loc_00E1D93C: push ecx
  loc_00E1D93D: call [004010A8h] ; rtcMsgBox
  loc_00E1D943: lea edx, var_CC
  loc_00E1D949: lea eax, var_BC
  loc_00E1D94F: push edx
  loc_00E1D950: lea ecx, var_AC
  loc_00E1D956: push eax
  loc_00E1D957: lea edx, var_9C
  loc_00E1D95D: push ecx
  loc_00E1D95E: push edx
  loc_00E1D95F: push 00000004h
  loc_00E1D961: call [00401038h] ; __vbaFreeVarList
  loc_00E1D967: mov eax, Me
  loc_00E1D96A: add esp, 00000014h
  loc_00E1D96D: mov ecx, [eax]
  loc_00E1D96F: push eax
  loc_00E1D970: call [ecx+000004CCh]
  loc_00E1D976: lea edx, var_54
  loc_00E1D979: push eax
  loc_00E1D97A: push edx
  loc_00E1D97B: call edi
  loc_00E1D97D: mov esi, eax
  loc_00E1D97F: push FFFFFFFFh
  loc_00E1D981: push esi
  loc_00E1D982: mov eax, [esi]
  loc_00E1D984: call [eax+0000008Ch]
  loc_00E1D98A: test eax, eax
  loc_00E1D98C: fnclex
  loc_00E1D98E: jge 00E1D9A2h
  loc_00E1D990: push 0000008Ch
  loc_00E1D995: push 006DCDA0h
  loc_00E1D99A: push esi
  loc_00E1D99B: push eax
  loc_00E1D99C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D9A2: lea ecx, var_54
  loc_00E1D9A5: call [00401254h] ; __vbaFreeObj
  loc_00E1D9AB: jmp 00E1E25Bh
  loc_00E1D9B0: mov eax, [00E3D060h]
  loc_00E1D9B5: cmp eax, ebx
  loc_00E1D9B7: jnz 00E1D9CEh
  loc_00E1D9B9: push 00E3D060h
  loc_00E1D9BE: push 006D87C0h
  loc_00E1D9C3: call [004011A0h] ; __vbaNew2
  loc_00E1D9C9: mov eax, [00E3D060h]
  loc_00E1D9CE: mov ecx, [eax]
  loc_00E1D9D0: push eax
  loc_00E1D9D1: call [ecx+00000470h]
  loc_00E1D9D7: lea edx, var_58
  loc_00E1D9DA: push eax
  loc_00E1D9DB: push edx
  loc_00E1D9DC: call edi
  loc_00E1D9DE: mov esi, eax
  loc_00E1D9E0: lea ecx, var_1C
  loc_00E1D9E3: push ecx
  loc_00E1D9E4: push esi
  loc_00E1D9E5: mov eax, [esi]
  loc_00E1D9E7: call [eax+00000050h]
  loc_00E1D9EA: cmp eax, ebx
  loc_00E1D9EC: fnclex
  loc_00E1D9EE: jge 00E1D9FFh
  loc_00E1D9F0: push 00000050h
  loc_00E1D9F2: push 006E0410h
  loc_00E1D9F7: push esi
  loc_00E1D9F8: push eax
  loc_00E1D9F9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1D9FF: mov edx, var_1C
  loc_00E1DA02: push edx
  loc_00E1DA03: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DA09: mov eax, [00E3D060h]
  loc_00E1DA0E: fstp real8 ptr var_114
  loc_00E1DA14: cmp eax, ebx
  loc_00E1DA16: jnz 00E1DA2Dh
  loc_00E1DA18: push 00E3D060h
  loc_00E1DA1D: push 006D87C0h
  loc_00E1DA22: call [004011A0h] ; __vbaNew2
  loc_00E1DA28: mov eax, [00E3D060h]
  loc_00E1DA2D: mov ecx, [eax]
  loc_00E1DA2F: push eax
  loc_00E1DA30: call [ecx+0000046Ch]
  loc_00E1DA36: lea edx, var_5C
  loc_00E1DA39: push eax
  loc_00E1DA3A: push edx
  loc_00E1DA3B: call edi
  loc_00E1DA3D: mov esi, eax
  loc_00E1DA3F: lea ecx, var_20
  loc_00E1DA42: push ecx
  loc_00E1DA43: push esi
  loc_00E1DA44: mov eax, [esi]
  loc_00E1DA46: call [eax+00000050h]
  loc_00E1DA49: cmp eax, ebx
  loc_00E1DA4B: fnclex
  loc_00E1DA4D: jge 00E1DA5Eh
  loc_00E1DA4F: push 00000050h
  loc_00E1DA51: push 006E0410h
  loc_00E1DA56: push esi
  loc_00E1DA57: push eax
  loc_00E1DA58: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DA5E: mov edx, var_20
  loc_00E1DA61: push edx
  loc_00E1DA62: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DA68: mov eax, [00E3D060h]
  loc_00E1DA6D: fstp real8 ptr var_11C
  loc_00E1DA73: cmp eax, ebx
  loc_00E1DA75: jnz 00E1DA8Ch
  loc_00E1DA77: push 00E3D060h
  loc_00E1DA7C: push 006D87C0h
  loc_00E1DA81: call [004011A0h] ; __vbaNew2
  loc_00E1DA87: mov eax, [00E3D060h]
  loc_00E1DA8C: mov ecx, [eax]
  loc_00E1DA8E: push eax
  loc_00E1DA8F: call [ecx+00000454h]
  loc_00E1DA95: lea edx, var_60
  loc_00E1DA98: push eax
  loc_00E1DA99: push edx
  loc_00E1DA9A: call edi
  loc_00E1DA9C: mov esi, eax
  loc_00E1DA9E: lea ecx, var_24
  loc_00E1DAA1: push ecx
  loc_00E1DAA2: push esi
  loc_00E1DAA3: mov eax, [esi]
  loc_00E1DAA5: call [eax+00000050h]
  loc_00E1DAA8: cmp eax, ebx
  loc_00E1DAAA: fnclex
  loc_00E1DAAC: jge 00E1DABDh
  loc_00E1DAAE: push 00000050h
  loc_00E1DAB0: push 006E0410h
  loc_00E1DAB5: push esi
  loc_00E1DAB6: push eax
  loc_00E1DAB7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DABD: mov edx, var_24
  loc_00E1DAC0: push edx
  loc_00E1DAC1: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DAC7: mov eax, [00E3D060h]
  loc_00E1DACC: fstp real8 ptr var_124
  loc_00E1DAD2: cmp eax, ebx
  loc_00E1DAD4: jnz 00E1DAEBh
  loc_00E1DAD6: push 00E3D060h
  loc_00E1DADB: push 006D87C0h
  loc_00E1DAE0: call [004011A0h] ; __vbaNew2
  loc_00E1DAE6: mov eax, [00E3D060h]
  loc_00E1DAEB: mov ecx, [eax]
  loc_00E1DAED: push eax
  loc_00E1DAEE: call [ecx+00000440h]
  loc_00E1DAF4: lea edx, var_64
  loc_00E1DAF7: push eax
  loc_00E1DAF8: push edx
  loc_00E1DAF9: call edi
  loc_00E1DAFB: mov esi, eax
  loc_00E1DAFD: lea ecx, var_28
  loc_00E1DB00: push ecx
  loc_00E1DB01: push esi
  loc_00E1DB02: mov eax, [esi]
  loc_00E1DB04: call [eax+00000050h]
  loc_00E1DB07: cmp eax, ebx
  loc_00E1DB09: fnclex
  loc_00E1DB0B: jge 00E1DB1Ch
  loc_00E1DB0D: push 00000050h
  loc_00E1DB0F: push 006E0410h
  loc_00E1DB14: push esi
  loc_00E1DB15: push eax
  loc_00E1DB16: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DB1C: mov edx, var_28
  loc_00E1DB1F: push edx
  loc_00E1DB20: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DB26: mov eax, [00E3D060h]
  loc_00E1DB2B: fstp real8 ptr var_12C
  loc_00E1DB31: cmp eax, ebx
  loc_00E1DB33: jnz 00E1DB4Ah
  loc_00E1DB35: push 00E3D060h
  loc_00E1DB3A: push 006D87C0h
  loc_00E1DB3F: call [004011A0h] ; __vbaNew2
  loc_00E1DB45: mov eax, [00E3D060h]
  loc_00E1DB4A: mov ecx, [eax]
  loc_00E1DB4C: push eax
  loc_00E1DB4D: call [ecx+000003FCh]
  loc_00E1DB53: lea edx, var_68
  loc_00E1DB56: push eax
  loc_00E1DB57: push edx
  loc_00E1DB58: call edi
  loc_00E1DB5A: mov esi, eax
  loc_00E1DB5C: lea ecx, var_2C
  loc_00E1DB5F: push ecx
  loc_00E1DB60: push esi
  loc_00E1DB61: mov eax, [esi]
  loc_00E1DB63: call [eax+00000050h]
  loc_00E1DB66: cmp eax, ebx
  loc_00E1DB68: fnclex
  loc_00E1DB6A: jge 00E1DB7Bh
  loc_00E1DB6C: push 00000050h
  loc_00E1DB6E: push 006E0410h
  loc_00E1DB73: push esi
  loc_00E1DB74: push eax
  loc_00E1DB75: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DB7B: mov edx, var_2C
  loc_00E1DB7E: push edx
  loc_00E1DB7F: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DB85: mov eax, [00E3D060h]
  loc_00E1DB8A: fstp real8 ptr var_134
  loc_00E1DB90: cmp eax, ebx
  loc_00E1DB92: jnz 00E1DBA9h
  loc_00E1DB94: push 00E3D060h
  loc_00E1DB99: push 006D87C0h
  loc_00E1DB9E: call [004011A0h] ; __vbaNew2
  loc_00E1DBA4: mov eax, [00E3D060h]
  loc_00E1DBA9: mov ecx, [eax]
  loc_00E1DBAB: push eax
  loc_00E1DBAC: call [ecx+000003ECh]
  loc_00E1DBB2: lea edx, var_6C
  loc_00E1DBB5: push eax
  loc_00E1DBB6: push edx
  loc_00E1DBB7: call edi
  loc_00E1DBB9: mov esi, eax
  loc_00E1DBBB: lea ecx, var_30
  loc_00E1DBBE: push ecx
  loc_00E1DBBF: push esi
  loc_00E1DBC0: mov eax, [esi]
  loc_00E1DBC2: call [eax+00000050h]
  loc_00E1DBC5: cmp eax, ebx
  loc_00E1DBC7: fnclex
  loc_00E1DBC9: jge 00E1DBDAh
  loc_00E1DBCB: push 00000050h
  loc_00E1DBCD: push 006E0410h
  loc_00E1DBD2: push esi
  loc_00E1DBD3: push eax
  loc_00E1DBD4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DBDA: mov edx, var_30
  loc_00E1DBDD: push edx
  loc_00E1DBDE: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DBE4: mov eax, [00E3D060h]
  loc_00E1DBE9: fstp real8 ptr var_13C
  loc_00E1DBEF: cmp eax, ebx
  loc_00E1DBF1: jnz 00E1DC08h
  loc_00E1DBF3: push 00E3D060h
  loc_00E1DBF8: push 006D87C0h
  loc_00E1DBFD: call [004011A0h] ; __vbaNew2
  loc_00E1DC03: mov eax, [00E3D060h]
  loc_00E1DC08: mov ecx, [eax]
  loc_00E1DC0A: push eax
  loc_00E1DC0B: call [ecx+000003DCh]
  loc_00E1DC11: lea edx, var_70
  loc_00E1DC14: push eax
  loc_00E1DC15: push edx
  loc_00E1DC16: call edi
  loc_00E1DC18: mov esi, eax
  loc_00E1DC1A: lea ecx, var_34
  loc_00E1DC1D: push ecx
  loc_00E1DC1E: push esi
  loc_00E1DC1F: mov eax, [esi]
  loc_00E1DC21: call [eax+00000050h]
  loc_00E1DC24: cmp eax, ebx
  loc_00E1DC26: fnclex
  loc_00E1DC28: jge 00E1DC39h
  loc_00E1DC2A: push 00000050h
  loc_00E1DC2C: push 006E0410h
  loc_00E1DC31: push esi
  loc_00E1DC32: push eax
  loc_00E1DC33: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DC39: mov edx, var_34
  loc_00E1DC3C: push edx
  loc_00E1DC3D: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DC43: mov eax, [00E3D060h]
  loc_00E1DC48: fstp real8 ptr var_144
  loc_00E1DC4E: cmp eax, ebx
  loc_00E1DC50: jnz 00E1DC67h
  loc_00E1DC52: push 00E3D060h
  loc_00E1DC57: push 006D87C0h
  loc_00E1DC5C: call [004011A0h] ; __vbaNew2
  loc_00E1DC62: mov eax, [00E3D060h]
  loc_00E1DC67: mov ecx, [eax]
  loc_00E1DC69: push eax
  loc_00E1DC6A: call [ecx+000003CCh]
  loc_00E1DC70: lea edx, var_74
  loc_00E1DC73: push eax
  loc_00E1DC74: push edx
  loc_00E1DC75: call edi
  loc_00E1DC77: mov esi, eax
  loc_00E1DC79: lea ecx, var_38
  loc_00E1DC7C: push ecx
  loc_00E1DC7D: push esi
  loc_00E1DC7E: mov eax, [esi]
  loc_00E1DC80: call [eax+00000050h]
  loc_00E1DC83: cmp eax, ebx
  loc_00E1DC85: fnclex
  loc_00E1DC87: jge 00E1DC98h
  loc_00E1DC89: push 00000050h
  loc_00E1DC8B: push 006E0410h
  loc_00E1DC90: push esi
  loc_00E1DC91: push eax
  loc_00E1DC92: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DC98: mov edx, var_38
  loc_00E1DC9B: push edx
  loc_00E1DC9C: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DCA2: mov eax, [00E3D060h]
  loc_00E1DCA7: fstp real8 ptr var_14C
  loc_00E1DCAD: cmp eax, ebx
  loc_00E1DCAF: jnz 00E1DCC6h
  loc_00E1DCB1: push 00E3D060h
  loc_00E1DCB6: push 006D87C0h
  loc_00E1DCBB: call [004011A0h] ; __vbaNew2
  loc_00E1DCC1: mov eax, [00E3D060h]
  loc_00E1DCC6: mov ecx, [eax]
  loc_00E1DCC8: push eax
  loc_00E1DCC9: call [ecx+000003BCh]
  loc_00E1DCCF: lea edx, var_78
  loc_00E1DCD2: push eax
  loc_00E1DCD3: push edx
  loc_00E1DCD4: call edi
  loc_00E1DCD6: mov esi, eax
  loc_00E1DCD8: lea ecx, var_3C
  loc_00E1DCDB: push ecx
  loc_00E1DCDC: push esi
  loc_00E1DCDD: mov eax, [esi]
  loc_00E1DCDF: call [eax+00000050h]
  loc_00E1DCE2: cmp eax, ebx
  loc_00E1DCE4: fnclex
  loc_00E1DCE6: jge 00E1DCF7h
  loc_00E1DCE8: push 00000050h
  loc_00E1DCEA: push 006E0410h
  loc_00E1DCEF: push esi
  loc_00E1DCF0: push eax
  loc_00E1DCF1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DCF7: mov edx, var_3C
  loc_00E1DCFA: push edx
  loc_00E1DCFB: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DD01: mov eax, [00E3D060h]
  loc_00E1DD06: fstp real8 ptr var_154
  loc_00E1DD0C: cmp eax, ebx
  loc_00E1DD0E: jnz 00E1DD25h
  loc_00E1DD10: push 00E3D060h
  loc_00E1DD15: push 006D87C0h
  loc_00E1DD1A: call [004011A0h] ; __vbaNew2
  loc_00E1DD20: mov eax, [00E3D060h]
  loc_00E1DD25: mov ecx, [eax]
  loc_00E1DD27: push eax
  loc_00E1DD28: call [ecx+000003A8h]
  loc_00E1DD2E: lea edx, var_7C
  loc_00E1DD31: push eax
  loc_00E1DD32: push edx
  loc_00E1DD33: call edi
  loc_00E1DD35: mov esi, eax
  loc_00E1DD37: lea ecx, var_40
  loc_00E1DD3A: push ecx
  loc_00E1DD3B: push esi
  loc_00E1DD3C: mov eax, [esi]
  loc_00E1DD3E: call [eax+00000050h]
  loc_00E1DD41: cmp eax, ebx
  loc_00E1DD43: fnclex
  loc_00E1DD45: jge 00E1DD56h
  loc_00E1DD47: push 00000050h
  loc_00E1DD49: push 006E0410h
  loc_00E1DD4E: push esi
  loc_00E1DD4F: push eax
  loc_00E1DD50: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DD56: mov edx, var_40
  loc_00E1DD59: push edx
  loc_00E1DD5A: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DD60: mov eax, [00E3D060h]
  loc_00E1DD65: fstp real8 ptr var_15C
  loc_00E1DD6B: cmp eax, ebx
  loc_00E1DD6D: jnz 00E1DD84h
  loc_00E1DD6F: push 00E3D060h
  loc_00E1DD74: push 006D87C0h
  loc_00E1DD79: call [004011A0h] ; __vbaNew2
  loc_00E1DD7F: mov eax, [00E3D060h]
  loc_00E1DD84: mov ecx, [eax]
  loc_00E1DD86: push eax
  loc_00E1DD87: call [ecx+00000398h]
  loc_00E1DD8D: lea edx, var_80
  loc_00E1DD90: push eax
  loc_00E1DD91: push edx
  loc_00E1DD92: call edi
  loc_00E1DD94: mov esi, eax
  loc_00E1DD96: lea ecx, var_44
  loc_00E1DD99: push ecx
  loc_00E1DD9A: push esi
  loc_00E1DD9B: mov eax, [esi]
  loc_00E1DD9D: call [eax+00000050h]
  loc_00E1DDA0: cmp eax, ebx
  loc_00E1DDA2: fnclex
  loc_00E1DDA4: jge 00E1DDB5h
  loc_00E1DDA6: push 00000050h
  loc_00E1DDA8: push 006E0410h
  loc_00E1DDAD: push esi
  loc_00E1DDAE: push eax
  loc_00E1DDAF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DDB5: mov edx, var_44
  loc_00E1DDB8: push edx
  loc_00E1DDB9: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DDBF: mov eax, [00E3D060h]
  loc_00E1DDC4: fstp real8 ptr var_164
  loc_00E1DDCA: cmp eax, ebx
  loc_00E1DDCC: jnz 00E1DDE3h
  loc_00E1DDCE: push 00E3D060h
  loc_00E1DDD3: push 006D87C0h
  loc_00E1DDD8: call [004011A0h] ; __vbaNew2
  loc_00E1DDDE: mov eax, [00E3D060h]
  loc_00E1DDE3: mov ecx, [eax]
  loc_00E1DDE5: push eax
  loc_00E1DDE6: call [ecx+00000388h]
  loc_00E1DDEC: lea edx, var_84
  loc_00E1DDF2: push eax
  loc_00E1DDF3: push edx
  loc_00E1DDF4: call edi
  loc_00E1DDF6: mov esi, eax
  loc_00E1DDF8: lea ecx, var_48
  loc_00E1DDFB: push ecx
  loc_00E1DDFC: push esi
  loc_00E1DDFD: mov eax, [esi]
  loc_00E1DDFF: call [eax+00000050h]
  loc_00E1DE02: cmp eax, ebx
  loc_00E1DE04: fnclex
  loc_00E1DE06: jge 00E1DE17h
  loc_00E1DE08: push 00000050h
  loc_00E1DE0A: push 006E0410h
  loc_00E1DE0F: push esi
  loc_00E1DE10: push eax
  loc_00E1DE11: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DE17: mov edx, var_48
  loc_00E1DE1A: push edx
  loc_00E1DE1B: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DE21: mov eax, [00E3D060h]
  loc_00E1DE26: fstp real8 ptr var_16C
  loc_00E1DE2C: cmp eax, ebx
  loc_00E1DE2E: jnz 00E1DE45h
  loc_00E1DE30: push 00E3D060h
  loc_00E1DE35: push 006D87C0h
  loc_00E1DE3A: call [004011A0h] ; __vbaNew2
  loc_00E1DE40: mov eax, [00E3D060h]
  loc_00E1DE45: mov ecx, [eax]
  loc_00E1DE47: push eax
  loc_00E1DE48: call [ecx+00000378h]
  loc_00E1DE4E: lea edx, var_88
  loc_00E1DE54: push eax
  loc_00E1DE55: push edx
  loc_00E1DE56: call edi
  loc_00E1DE58: mov esi, eax
  loc_00E1DE5A: lea ecx, var_4C
  loc_00E1DE5D: push ecx
  loc_00E1DE5E: push esi
  loc_00E1DE5F: mov eax, [esi]
  loc_00E1DE61: call [eax+00000050h]
  loc_00E1DE64: cmp eax, ebx
  loc_00E1DE66: fnclex
  loc_00E1DE68: jge 00E1DE79h
  loc_00E1DE6A: push 00000050h
  loc_00E1DE6C: push 006E0410h
  loc_00E1DE71: push esi
  loc_00E1DE72: push eax
  loc_00E1DE73: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DE79: mov edx, var_4C
  loc_00E1DE7C: push edx
  loc_00E1DE7D: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DE83: mov eax, [00E3D060h]
  loc_00E1DE88: fstp real8 ptr var_174
  loc_00E1DE8E: cmp eax, ebx
  loc_00E1DE90: jnz 00E1DEA7h
  loc_00E1DE92: push 00E3D060h
  loc_00E1DE97: push 006D87C0h
  loc_00E1DE9C: call [004011A0h] ; __vbaNew2
  loc_00E1DEA2: mov eax, [00E3D060h]
  loc_00E1DEA7: mov ecx, [eax]
  loc_00E1DEA9: push eax
  loc_00E1DEAA: call [ecx+00000340h]
  loc_00E1DEB0: lea edx, var_8C
  loc_00E1DEB6: push eax
  loc_00E1DEB7: push edx
  loc_00E1DEB8: call edi
  loc_00E1DEBA: mov var_1E8, eax
  loc_00E1DEC0: mov eax, [00E3D060h]
  loc_00E1DEC5: cmp eax, ebx
  loc_00E1DEC7: jnz 00E1DEDEh
  loc_00E1DEC9: push 00E3D060h
  loc_00E1DECE: push 006D87C0h
  loc_00E1DED3: call [004011A0h] ; __vbaNew2
  loc_00E1DED9: mov eax, [00E3D060h]
  loc_00E1DEDE: mov ecx, [eax]
  loc_00E1DEE0: push eax
  loc_00E1DEE1: call [ecx+00000474h]
  loc_00E1DEE7: lea edx, var_54
  loc_00E1DEEA: push eax
  loc_00E1DEEB: push edx
  loc_00E1DEEC: call edi
  loc_00E1DEEE: mov esi, eax
  loc_00E1DEF0: lea ecx, var_18
  loc_00E1DEF3: push ecx
  loc_00E1DEF4: push esi
  loc_00E1DEF5: mov eax, [esi]
  loc_00E1DEF7: call [eax+00000050h]
  loc_00E1DEFA: cmp eax, ebx
  loc_00E1DEFC: fnclex
  loc_00E1DEFE: jge 00E1DF0Fh
  loc_00E1DF00: push 00000050h
  loc_00E1DF02: push 006E0410h
  loc_00E1DF07: push esi
  loc_00E1DF08: push eax
  loc_00E1DF09: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DF0F: mov edx, var_1E8
  loc_00E1DF15: mov eax, var_18
  loc_00E1DF18: push eax
  loc_00E1DF19: mov esi, [edx]
  loc_00E1DF1B: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1DF21: fadd st0, real8 ptr var_114
  loc_00E1DF27: sub esp, 00000008h
  loc_00E1DF2A: fadd st0, real8 ptr var_11C
  loc_00E1DF30: fadd st0, real8 ptr var_124
  loc_00E1DF36: fadd st0, real8 ptr var_12C
  loc_00E1DF3C: fadd st0, real8 ptr var_134
  loc_00E1DF42: fadd st0, real8 ptr var_13C
  loc_00E1DF48: fadd st0, real8 ptr var_144
  loc_00E1DF4E: fadd st0, real8 ptr var_14C
  loc_00E1DF54: fadd st0, real8 ptr var_154
  loc_00E1DF5A: fadd st0, real8 ptr var_15C
  loc_00E1DF60: fadd st0, real8 ptr var_164
  loc_00E1DF66: fadd st0, real8 ptr var_16C
  loc_00E1DF6C: fadd st0, real8 ptr var_174
  loc_00E1DF72: fnstsw ax
  loc_00E1DF74: test al, 0Dh
  loc_00E1DF76: jnz 00E1E34Ch
  loc_00E1DF7C: fstp real8 ptr [esp]
  loc_00E1DF7F: call [00401134h] ; __vbaStrR8
  loc_00E1DF85: mov edx, eax
  loc_00E1DF87: lea ecx, var_50
  loc_00E1DF8A: call [00401228h] ; __vbaStrMove
  loc_00E1DF90: mov ecx, esi
  loc_00E1DF92: mov esi, var_1E8
  loc_00E1DF98: push eax
  loc_00E1DF99: push esi
  loc_00E1DF9A: call [ecx+00000054h]
  loc_00E1DF9D: cmp eax, ebx
  loc_00E1DF9F: fnclex
  loc_00E1DFA1: jge 00E1DFB2h
  loc_00E1DFA3: push 00000054h
  loc_00E1DFA5: push 006E0410h
  loc_00E1DFAA: push esi
  loc_00E1DFAB: push eax
  loc_00E1DFAC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1DFB2: lea edx, var_50
  loc_00E1DFB5: lea eax, var_4C
  loc_00E1DFB8: push edx
  loc_00E1DFB9: lea ecx, var_48
  loc_00E1DFBC: push eax
  loc_00E1DFBD: lea edx, var_44
  loc_00E1DFC0: push ecx
  loc_00E1DFC1: lea eax, var_40
  loc_00E1DFC4: push edx
  loc_00E1DFC5: lea ecx, var_3C
  loc_00E1DFC8: push eax
  loc_00E1DFC9: lea edx, var_38
  loc_00E1DFCC: push ecx
  loc_00E1DFCD: lea eax, var_34
  loc_00E1DFD0: push edx
  loc_00E1DFD1: lea ecx, var_30
  loc_00E1DFD4: push eax
  loc_00E1DFD5: lea edx, var_2C
  loc_00E1DFD8: push ecx
  loc_00E1DFD9: lea eax, var_28
  loc_00E1DFDC: push edx
  loc_00E1DFDD: lea ecx, var_24
  loc_00E1DFE0: push eax
  loc_00E1DFE1: lea edx, var_20
  loc_00E1DFE4: push ecx
  loc_00E1DFE5: lea eax, var_1C
  loc_00E1DFE8: push edx
  loc_00E1DFE9: lea ecx, var_18
  loc_00E1DFEC: push eax
  loc_00E1DFED: push ecx
  loc_00E1DFEE: push 0000000Fh
  loc_00E1DFF0: call [004011BCh] ; __vbaFreeStrList
  loc_00E1DFF6: lea edx, var_8C
  loc_00E1DFFC: lea eax, var_88
  loc_00E1E002: push edx
  loc_00E1E003: lea ecx, var_84
  loc_00E1E009: push eax
  loc_00E1E00A: lea edx, var_80
  loc_00E1E00D: push ecx
  loc_00E1E00E: lea eax, var_7C
  loc_00E1E011: push edx
  loc_00E1E012: lea ecx, var_78
  loc_00E1E015: push eax
  loc_00E1E016: lea edx, var_74
  loc_00E1E019: push ecx
  loc_00E1E01A: lea eax, var_70
  loc_00E1E01D: push edx
  loc_00E1E01E: lea ecx, var_6C
  loc_00E1E021: push eax
  loc_00E1E022: lea edx, var_68
  loc_00E1E025: push ecx
  loc_00E1E026: lea eax, var_64
  loc_00E1E029: push edx
  loc_00E1E02A: lea ecx, var_60
  loc_00E1E02D: push eax
  loc_00E1E02E: lea edx, var_5C
  loc_00E1E031: push ecx
  loc_00E1E032: lea eax, var_58
  loc_00E1E035: push edx
  loc_00E1E036: lea ecx, var_54
  loc_00E1E039: push eax
  loc_00E1E03A: push ecx
  loc_00E1E03B: push 0000000Fh
  loc_00E1E03D: call [00401048h] ; __vbaFreeObjList
  loc_00E1E043: mov esi, Me
  loc_00E1E046: add esp, 00000080h
  loc_00E1E04C: mov edx, [esi]
  loc_00E1E04E: push esi
  loc_00E1E04F: call [edx+0000033Ch]
  loc_00E1E055: push eax
  loc_00E1E056: lea eax, var_58
  loc_00E1E059: push eax
  loc_00E1E05A: call edi
  loc_00E1E05C: mov ecx, [esi]
  loc_00E1E05E: push esi
  loc_00E1E05F: mov var_184, eax
  loc_00E1E065: call [ecx+00000340h]
  loc_00E1E06B: lea edx, var_54
  loc_00E1E06E: push eax
  loc_00E1E06F: push edx
  loc_00E1E070: call edi
  loc_00E1E072: mov esi, eax
  loc_00E1E074: lea ecx, var_18
  loc_00E1E077: push ecx
  loc_00E1E078: push esi
  loc_00E1E079: mov eax, [esi]
  loc_00E1E07B: call [eax+00000050h]
  loc_00E1E07E: cmp eax, ebx
  loc_00E1E080: fnclex
  loc_00E1E082: jge 00E1E093h
  loc_00E1E084: push 00000050h
  loc_00E1E086: push 006E0410h
  loc_00E1E08B: push esi
  loc_00E1E08C: push eax
  loc_00E1E08D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E093: mov edx, var_18
  loc_00E1E096: lea ecx, var_1C
  loc_00E1E099: mov var_18, ebx
  loc_00E1E09C: call [00401228h] ; __vbaStrMove
  loc_00E1E0A2: mov esi, Me
  loc_00E1E0A5: lea eax, var_20
  loc_00E1E0A8: lea ecx, var_1C
  loc_00E1E0AB: push eax
  loc_00E1E0AC: mov edx, [esi]
  loc_00E1E0AE: push ecx
  loc_00E1E0AF: push esi
  loc_00E1E0B0: call [edx+000006F8h]
  loc_00E1E0B6: cmp eax, ebx
  loc_00E1E0B8: jge 00E1E0CCh
  loc_00E1E0BA: push 000006F8h
  loc_00E1E0BF: push 006DFA7Ch ; "^¿ÏG≤ê"
  loc_00E1E0C4: push esi
  loc_00E1E0C5: push eax
  loc_00E1E0C6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E0CC: mov esi, var_184
  loc_00E1E0D2: mov eax, var_20
  loc_00E1E0D5: push eax
  loc_00E1E0D6: push esi
  loc_00E1E0D7: mov edx, [esi]
  loc_00E1E0D9: call [edx+00000054h]
  loc_00E1E0DC: cmp eax, ebx
  loc_00E1E0DE: fnclex
  loc_00E1E0E0: jge 00E1E0F1h
  loc_00E1E0E2: push 00000054h
  loc_00E1E0E4: push 006E0410h
  loc_00E1E0E9: push esi
  loc_00E1E0EA: push eax
  loc_00E1E0EB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E0F1: lea ecx, var_20
  loc_00E1E0F4: lea edx, var_1C
  loc_00E1E0F7: push ecx
  loc_00E1E0F8: push edx
  loc_00E1E0F9: push 00000002h
  loc_00E1E0FB: call [004011BCh] ; __vbaFreeStrList
  loc_00E1E101: lea eax, var_58
  loc_00E1E104: lea ecx, var_54
  loc_00E1E107: push eax
  loc_00E1E108: push ecx
  loc_00E1E109: push 00000002h
  loc_00E1E10B: call [00401048h] ; __vbaFreeObjList
  loc_00E1E111: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E1E117: mov ecx, 80020004h
  loc_00E1E11C: mov var_C4, ecx
  loc_00E1E122: mov eax, 0000000Ah
  loc_00E1E127: mov var_B4, ecx
  loc_00E1E12D: mov ebx, 00000008h
  loc_00E1E132: add esp, 00000018h
  loc_00E1E135: lea edx, var_EC
  loc_00E1E13B: lea ecx, var_AC
  loc_00E1E141: mov var_CC, eax
  loc_00E1E147: mov var_BC, eax
  loc_00E1E14D: mov var_E4, 006E0B08h ; "Need To Do"
  loc_00E1E157: mov var_EC, ebx
  loc_00E1E15D: call __vbaVarDup
  loc_00E1E15F: lea edx, var_DC
  loc_00E1E165: lea ecx, var_9C
  loc_00E1E16B: mov var_D4, 006E13C0h ; "Silahkan Simpan Data yang Telah di Input Sekarang"
  loc_00E1E175: mov var_DC, ebx
  loc_00E1E17B: call __vbaVarDup
  loc_00E1E17D: lea edx, var_CC
  loc_00E1E183: lea eax, var_BC
  loc_00E1E189: push edx
  loc_00E1E18A: lea ecx, var_AC
  loc_00E1E190: push eax
  loc_00E1E191: push ecx
  loc_00E1E192: lea edx, var_9C
  loc_00E1E198: push 00000040h
  loc_00E1E19A: push edx
  loc_00E1E19B: call [004010A8h] ; rtcMsgBox
  loc_00E1E1A1: lea eax, var_CC
  loc_00E1E1A7: lea ecx, var_BC
  loc_00E1E1AD: push eax
  loc_00E1E1AE: lea edx, var_AC
  loc_00E1E1B4: push ecx
  loc_00E1E1B5: lea eax, var_9C
  loc_00E1E1BB: push edx
  loc_00E1E1BC: push eax
  loc_00E1E1BD: push 00000004h
  loc_00E1E1BF: call [00401038h] ; __vbaFreeVarList
  loc_00E1E1C5: mov ebx, Me
  loc_00E1E1C8: add esp, 00000014h
  loc_00E1E1CB: mov ecx, [ebx]
  loc_00E1E1CD: push ebx
  loc_00E1E1CE: call [ecx+000004D4h]
  loc_00E1E1D4: lea edx, var_54
  loc_00E1E1D7: push eax
  loc_00E1E1D8: push edx
  loc_00E1E1D9: call edi
  loc_00E1E1DB: mov esi, eax
  loc_00E1E1DD: push FFFFFFFFh
  loc_00E1E1DF: push esi
  loc_00E1E1E0: mov eax, [esi]
  loc_00E1E1E2: call [eax+0000008Ch]
  loc_00E1E1E8: test eax, eax
  loc_00E1E1EA: fnclex
  loc_00E1E1EC: jge 00E1E200h
  loc_00E1E1EE: push 0000008Ch
  loc_00E1E1F3: push 006DCDA0h
  loc_00E1E1F8: push esi
  loc_00E1E1F9: push eax
  loc_00E1E1FA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E200: mov esi, [00401254h] ; __vbaFreeObj
  loc_00E1E206: lea ecx, var_54
  loc_00E1E209: call __vbaFreeObj
  loc_00E1E20B: sub esp, 00000010h
  loc_00E1E20E: mov ecx, 0000000Bh
  loc_00E1E213: mov edx, esp
  loc_00E1E215: mov var_DC, ecx
  loc_00E1E21B: xor eax, eax
  loc_00E1E21D: push 8001000Dh
  loc_00E1E222: mov [edx], ecx
  loc_00E1E224: mov ecx, var_D8
  loc_00E1E22A: mov var_D4, eax
  loc_00E1E230: push ebx
  loc_00E1E231: mov [edx+00000004h], ecx
  loc_00E1E234: mov ecx, [ebx]
  loc_00E1E236: mov [edx+00000008h], eax
  loc_00E1E239: mov eax, var_D0
  loc_00E1E23F: mov [edx+0000000Ch], eax
  loc_00E1E242: call [ecx+000004C8h]
  loc_00E1E248: lea edx, var_54
  loc_00E1E24B: push eax
  loc_00E1E24C: push edx
  loc_00E1E24D: call edi
  loc_00E1E24F: push eax
  loc_00E1E250: call [00401238h] ; __vbaLateIdSt
  loc_00E1E256: lea ecx, var_54
  loc_00E1E259: call __vbaFreeObj
  loc_00E1E25B: mov var_4, 00000000h
  loc_00E1E262: fwait
  loc_00E1E263: push 00E1E32Dh
  loc_00E1E268: jmp 00E1E32Ch
  loc_00E1E26D: lea eax, var_50
  loc_00E1E270: lea ecx, var_4C
  loc_00E1E273: push eax
  loc_00E1E274: lea edx, var_48
  loc_00E1E277: push ecx
  loc_00E1E278: lea eax, var_44
  loc_00E1E27B: push edx
  loc_00E1E27C: lea ecx, var_40
  loc_00E1E27F: push eax
  loc_00E1E280: lea edx, var_3C
  loc_00E1E283: push ecx
  loc_00E1E284: lea eax, var_38
  loc_00E1E287: push edx
  loc_00E1E288: lea ecx, var_34
  loc_00E1E28B: push eax
  loc_00E1E28C: lea edx, var_30
  loc_00E1E28F: push ecx
  loc_00E1E290: lea eax, var_2C
  loc_00E1E293: push edx
  loc_00E1E294: lea ecx, var_28
  loc_00E1E297: push eax
  loc_00E1E298: lea edx, var_24
  loc_00E1E29B: push ecx
  loc_00E1E29C: lea eax, var_20
  loc_00E1E29F: push edx
  loc_00E1E2A0: lea ecx, var_1C
  loc_00E1E2A3: push eax
  loc_00E1E2A4: lea edx, var_18
  loc_00E1E2A7: push ecx
  loc_00E1E2A8: push edx
  loc_00E1E2A9: push 0000000Fh
  loc_00E1E2AB: call [004011BCh] ; __vbaFreeStrList
  loc_00E1E2B1: lea eax, var_8C
  loc_00E1E2B7: lea ecx, var_88
  loc_00E1E2BD: push eax
  loc_00E1E2BE: lea edx, var_84
  loc_00E1E2C4: push ecx
  loc_00E1E2C5: lea eax, var_80
  loc_00E1E2C8: push edx
  loc_00E1E2C9: lea ecx, var_7C
  loc_00E1E2CC: push eax
  loc_00E1E2CD: lea edx, var_78
  loc_00E1E2D0: push ecx
  loc_00E1E2D1: lea eax, var_74
  loc_00E1E2D4: push edx
  loc_00E1E2D5: lea ecx, var_70
  loc_00E1E2D8: push eax
  loc_00E1E2D9: lea edx, var_6C
  loc_00E1E2DC: push ecx
  loc_00E1E2DD: lea eax, var_68
  loc_00E1E2E0: push edx
  loc_00E1E2E1: lea ecx, var_64
  loc_00E1E2E4: push eax
  loc_00E1E2E5: lea edx, var_60
  loc_00E1E2E8: push ecx
  loc_00E1E2E9: lea eax, var_5C
  loc_00E1E2EC: push edx
  loc_00E1E2ED: lea ecx, var_58
  loc_00E1E2F0: push eax
  loc_00E1E2F1: lea edx, var_54
  loc_00E1E2F4: push ecx
  loc_00E1E2F5: push edx
  loc_00E1E2F6: push 0000000Fh
  loc_00E1E2F8: call [00401048h] ; __vbaFreeObjList
  loc_00E1E2FE: add esp, 00000080h
  loc_00E1E304: lea eax, var_CC
  loc_00E1E30A: lea ecx, var_BC
  loc_00E1E310: lea edx, var_AC
  loc_00E1E316: push eax
  loc_00E1E317: push ecx
  loc_00E1E318: lea eax, var_9C
  loc_00E1E31E: push edx
  loc_00E1E31F: push eax
  loc_00E1E320: push 00000004h
  loc_00E1E322: call [00401038h] ; __vbaFreeVarList
  loc_00E1E328: add esp, 00000014h
  loc_00E1E32B: ret
  loc_00E1E32C: ret
  loc_00E1E32D: mov eax, Me
  loc_00E1E330: push eax
  loc_00E1E331: mov ecx, [eax]
  loc_00E1E333: call [ecx+00000008h]
  loc_00E1E336: mov eax, var_4
  loc_00E1E339: mov ecx, var_14
  loc_00E1E33C: pop edi
  loc_00E1E33D: pop esi
  loc_00E1E33E: mov fs:[00000000h], ecx
  loc_00E1E345: pop ebx
  loc_00E1E346: mov esp, ebp
  loc_00E1E348: pop ebp
  loc_00E1E349: retn 0004h
End Sub

Private Sub back_Click() 'E18830
  loc_00E18830: push ebp
  loc_00E18831: mov ebp, esp
  loc_00E18833: sub esp, 0000000Ch
  loc_00E18836: push 00402836h ; __vbaExceptHandler
  loc_00E1883B: mov eax, fs:[00000000h]
  loc_00E18841: push eax
  loc_00E18842: mov fs:[00000000h], esp
  loc_00E18849: sub esp, 00000038h
  loc_00E1884C: push ebx
  loc_00E1884D: push esi
  loc_00E1884E: push edi
  loc_00E1884F: mov var_C, esp
  loc_00E18852: mov var_8, 00402300h
  loc_00E18859: mov edi, Me
  loc_00E1885C: mov eax, edi
  loc_00E1885E: and eax, 00000001h
  loc_00E18861: mov var_4, eax
  loc_00E18864: and edi, FFFFFFFEh
  loc_00E18867: push edi
  loc_00E18868: mov Me, edi
  loc_00E1886B: mov ecx, [edi]
  loc_00E1886D: call [ecx+00000004h]
  loc_00E18870: mov eax, [00E3D9CCh]
  loc_00E18875: mov var_18, 00000000h
  loc_00E1887C: test eax, eax
  loc_00E1887E: jnz 00E18890h
  loc_00E18880: push 00E3D9CCh
  loc_00E18885: push 006DC960h
  loc_00E1888A: call [004011A0h] ; __vbaNew2
  loc_00E18890: mov esi, [00E3D9CCh]
  loc_00E18896: lea edx, var_18
  loc_00E18899: push edi
  loc_00E1889A: push edx
  loc_00E1889B: mov ebx, [esi]
  loc_00E1889D: call [004010B4h] ; __vbaObjSetAddref
  loc_00E188A3: push eax
  loc_00E188A4: push esi
  loc_00E188A5: call [ebx+00000010h]
  loc_00E188A8: test eax, eax
  loc_00E188AA: fnclex
  loc_00E188AC: jge 00E188BDh
  loc_00E188AE: push 00000010h
  loc_00E188B0: push 006DC950h
  loc_00E188B5: push esi
  loc_00E188B6: push eax
  loc_00E188B7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E188BD: lea ecx, var_18
  loc_00E188C0: call [00401254h] ; __vbaFreeObj
  loc_00E188C6: mov eax, [00E3D024h]
  loc_00E188CB: test eax, eax
  loc_00E188CD: jnz 00E188DFh
  loc_00E188CF: push 00E3D024h
  loc_00E188D4: push 006CE120h
  loc_00E188D9: call [004011A0h] ; __vbaNew2
  loc_00E188DF: sub esp, 00000010h
  loc_00E188E2: mov ecx, 0000000Ah
  loc_00E188E7: mov esi, esp
  loc_00E188E9: mov var_28, ecx
  loc_00E188EC: mov eax, 80020004h
  loc_00E188F1: mov ebx, [00E3D024h]
  loc_00E188F7: mov [esi], ecx
  loc_00E188F9: mov ecx, var_34
  loc_00E188FC: mov edx, eax
  loc_00E188FE: sub esp, 00000010h
  loc_00E18901: mov [esi+00000004h], ecx
  loc_00E18904: mov edi, [ebx]
  loc_00E18906: mov ecx, esp
  loc_00E18908: mov var_4C, edi
  loc_00E1890B: mov [esi+00000008h], eax
  loc_00E1890E: mov eax, var_2C
  loc_00E18911: mov edi, var_1C
  loc_00E18914: push ebx
  loc_00E18915: mov [esi+0000000Ch], eax
  loc_00E18918: mov eax, var_28
  loc_00E1891B: mov esi, var_24
  loc_00E1891E: mov [ecx], eax
  loc_00E18920: mov [ecx+00000004h], esi
  loc_00E18923: mov [ecx+00000008h], edx
  loc_00E18926: mov [ecx+0000000Ch], edi
  loc_00E18929: mov ecx, var_4C
  loc_00E1892C: call [ecx+000002B0h]
  loc_00E18932: test eax, eax
  loc_00E18934: fnclex
  loc_00E18936: jge 00E1894Ah
  loc_00E18938: push 000002B0h
  loc_00E1893D: push 006DC970h
  loc_00E18942: push ebx
  loc_00E18943: push eax
  loc_00E18944: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1894A: mov eax, [00E3D024h]
  loc_00E1894F: or ebx, FFFFFFFFh
  loc_00E18952: test eax, eax
  loc_00E18954: jnz 00E1896Bh
  loc_00E18956: push 00E3D024h
  loc_00E1895B: push 006CE120h
  loc_00E18960: call [004011A0h] ; __vbaNew2
  loc_00E18966: mov eax, [00E3D024h]
  loc_00E1896B: sub esp, 00000010h
  loc_00E1896E: mov ecx, 0000000Bh
  loc_00E18973: mov edx, esp
  loc_00E18975: push 80010007h
  loc_00E1897A: push eax
  loc_00E1897B: mov [edx], ecx
  loc_00E1897D: mov ecx, [eax]
  loc_00E1897F: mov [edx+00000004h], esi
  loc_00E18982: mov [edx+00000008h], ebx
  loc_00E18985: mov [edx+0000000Ch], edi
  loc_00E18988: call [ecx+00000318h]
  loc_00E1898E: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E18994: lea edx, var_18
  loc_00E18997: push eax
  loc_00E18998: push edx
  loc_00E18999: call ebx
  loc_00E1899B: push eax
  loc_00E1899C: call [00401238h] ; __vbaLateIdSt
  loc_00E189A2: lea ecx, var_18
  loc_00E189A5: call [00401254h] ; __vbaFreeObj
  loc_00E189AB: mov eax, [00E3D024h]
  loc_00E189B0: test eax, eax
  loc_00E189B2: jnz 00E189C9h
  loc_00E189B4: push 00E3D024h
  loc_00E189B9: push 006CE120h
  loc_00E189BE: call [004011A0h] ; __vbaNew2
  loc_00E189C4: mov eax, [00E3D024h]
  loc_00E189C9: sub esp, 00000010h
  loc_00E189CC: mov ecx, 0000000Bh
  loc_00E189D1: mov edx, esp
  loc_00E189D3: push 80010007h
  loc_00E189D8: push eax
  loc_00E189D9: mov [edx], ecx
  loc_00E189DB: or ecx, FFFFFFFFh
  loc_00E189DE: mov [edx+00000004h], esi
  loc_00E189E1: mov [edx+00000008h], ecx
  loc_00E189E4: mov ecx, [eax]
  loc_00E189E6: mov [edx+0000000Ch], edi
  loc_00E189E9: call [ecx+0000031Ch]
  loc_00E189EF: lea edx, var_18
  loc_00E189F2: push eax
  loc_00E189F3: push edx
  loc_00E189F4: call ebx
  loc_00E189F6: push eax
  loc_00E189F7: call [00401238h] ; __vbaLateIdSt
  loc_00E189FD: lea ecx, var_18
  loc_00E18A00: call [00401254h] ; __vbaFreeObj
  loc_00E18A06: mov eax, [00E3D024h]
  loc_00E18A0B: mov var_20, FFFFFFFFh
  loc_00E18A12: test eax, eax
  loc_00E18A14: jnz 00E18A2Bh
  loc_00E18A16: push 00E3D024h
  loc_00E18A1B: push 006CE120h
  loc_00E18A20: call [004011A0h] ; __vbaNew2
  loc_00E18A26: mov eax, [00E3D024h]
  loc_00E18A2B: sub esp, 00000010h
  loc_00E18A2E: mov ecx, 0000000Bh
  loc_00E18A33: mov edx, esp
  loc_00E18A35: push 80010007h
  loc_00E18A3A: push eax
  loc_00E18A3B: mov [edx], ecx
  loc_00E18A3D: mov ecx, var_20
  loc_00E18A40: mov [edx+00000004h], esi
  loc_00E18A43: mov [edx+00000008h], ecx
  loc_00E18A46: mov [edx+0000000Ch], edi
  loc_00E18A49: mov edx, [eax]
  loc_00E18A4B: call [edx+00000320h]
  loc_00E18A51: push eax
  loc_00E18A52: lea eax, var_18
  loc_00E18A55: push eax
  loc_00E18A56: call ebx
  loc_00E18A58: push eax
  loc_00E18A59: call [00401238h] ; __vbaLateIdSt
  loc_00E18A5F: lea ecx, var_18
  loc_00E18A62: call [00401254h] ; __vbaFreeObj
  loc_00E18A68: mov eax, [00E3D024h]
  loc_00E18A6D: test eax, eax
  loc_00E18A6F: jnz 00E18A86h
  loc_00E18A71: push 00E3D024h
  loc_00E18A76: push 006CE120h
  loc_00E18A7B: call [004011A0h] ; __vbaNew2
  loc_00E18A81: mov eax, [00E3D024h]
  loc_00E18A86: sub esp, 00000010h
  loc_00E18A89: mov ecx, 00000008h
  loc_00E18A8E: mov edx, esp
  loc_00E18A90: push FFFFFDFAh
  loc_00E18A95: push eax
  loc_00E18A96: mov [edx], ecx
  loc_00E18A98: mov ecx, 006DCDB4h ; " A D M I N "
  loc_00E18A9D: mov [edx+00000004h], esi
  loc_00E18AA0: mov [edx+00000008h], ecx
  loc_00E18AA3: mov ecx, [eax]
  loc_00E18AA5: mov [edx+0000000Ch], edi
  loc_00E18AA8: call [ecx+00000324h]
  loc_00E18AAE: lea edx, var_18
  loc_00E18AB1: push eax
  loc_00E18AB2: push edx
  loc_00E18AB3: call ebx
  loc_00E18AB5: push eax
  loc_00E18AB6: call [00401238h] ; __vbaLateIdSt
  loc_00E18ABC: lea ecx, var_18
  loc_00E18ABF: call [00401254h] ; __vbaFreeObj
  loc_00E18AC5: mov eax, [00E3D024h]
  loc_00E18ACA: mov var_28, 0000000Bh
  loc_00E18AD1: test eax, eax
  loc_00E18AD3: jnz 00E18AEAh
  loc_00E18AD5: push 00E3D024h
  loc_00E18ADA: push 006CE120h
  loc_00E18ADF: call [004011A0h] ; __vbaNew2
  loc_00E18AE5: mov eax, [00E3D024h]
  loc_00E18AEA: mov ecx, var_28
  loc_00E18AED: sub esp, 00000010h
  loc_00E18AF0: mov edx, esp
  loc_00E18AF2: push 8001000Dh
  loc_00E18AF7: push eax
  loc_00E18AF8: mov [edx], ecx
  loc_00E18AFA: xor ecx, ecx
  loc_00E18AFC: mov [edx+00000004h], esi
  loc_00E18AFF: mov [edx+00000008h], ecx
  loc_00E18B02: mov [edx+0000000Ch], edi
  loc_00E18B05: mov edx, [eax]
  loc_00E18B07: call [edx+00000324h]
  loc_00E18B0D: push eax
  loc_00E18B0E: lea eax, var_18
  loc_00E18B11: push eax
  loc_00E18B12: call ebx
  loc_00E18B14: push eax
  loc_00E18B15: call [00401238h] ; __vbaLateIdSt
  loc_00E18B1B: lea ecx, var_18
  loc_00E18B1E: call [00401254h] ; __vbaFreeObj
  loc_00E18B24: mov eax, [00E3D024h]
  loc_00E18B29: test eax, eax
  loc_00E18B2B: jnz 00E18B42h
  loc_00E18B2D: push 00E3D024h
  loc_00E18B32: push 006CE120h
  loc_00E18B37: call [004011A0h] ; __vbaNew2
  loc_00E18B3D: mov eax, [00E3D024h]
  loc_00E18B42: sub esp, 00000010h
  loc_00E18B45: mov ecx, 00000003h
  loc_00E18B4A: mov edx, esp
  loc_00E18B4C: push FFFFFE0Bh
  loc_00E18B51: push eax
  loc_00E18B52: mov [edx], ecx
  loc_00E18B54: mov ecx, 00404040h
  loc_00E18B59: mov [edx+00000004h], esi
  loc_00E18B5C: mov [edx+00000008h], ecx
  loc_00E18B5F: mov ecx, [eax]
  loc_00E18B61: mov [edx+0000000Ch], edi
  loc_00E18B64: call [ecx+00000324h]
  loc_00E18B6A: lea edx, var_18
  loc_00E18B6D: push eax
  loc_00E18B6E: push edx
  loc_00E18B6F: call ebx
  loc_00E18B71: push eax
  loc_00E18B72: call [00401238h] ; __vbaLateIdSt
  loc_00E18B78: lea ecx, var_18
  loc_00E18B7B: call [00401254h] ; __vbaFreeObj
  loc_00E18B81: mov eax, [00E3D024h]
  loc_00E18B86: mov var_28, 00000003h
  loc_00E18B8D: test eax, eax
  loc_00E18B8F: jnz 00E18BA6h
  loc_00E18B91: push 00E3D024h
  loc_00E18B96: push 006CE120h
  loc_00E18B9B: call [004011A0h] ; __vbaNew2
  loc_00E18BA1: mov eax, [00E3D024h]
  loc_00E18BA6: mov ecx, var_28
  loc_00E18BA9: sub esp, 00000010h
  loc_00E18BAC: mov edx, esp
  loc_00E18BAE: push FFFFFDFFh
  loc_00E18BB3: push eax
  loc_00E18BB4: mov [edx], ecx
  loc_00E18BB6: mov ecx, 00E0E0E0h
  loc_00E18BBB: mov [edx+00000004h], esi
  loc_00E18BBE: mov [edx+00000008h], ecx
  loc_00E18BC1: mov [edx+0000000Ch], edi
  loc_00E18BC4: mov edx, [eax]
  loc_00E18BC6: call [edx+00000324h]
  loc_00E18BCC: push eax
  loc_00E18BCD: lea eax, var_18
  loc_00E18BD0: push eax
  loc_00E18BD1: call ebx
  loc_00E18BD3: push eax
  loc_00E18BD4: call [00401238h] ; __vbaLateIdSt
  loc_00E18BDA: lea ecx, var_18
  loc_00E18BDD: call [00401254h] ; __vbaFreeObj
  loc_00E18BE3: mov var_4, 00000000h
  loc_00E18BEA: push 00E18BFCh
  loc_00E18BEF: jmp 00E18BFBh
  loc_00E18BF1: lea ecx, var_18
  loc_00E18BF4: call [00401254h] ; __vbaFreeObj
  loc_00E18BFA: ret
  loc_00E18BFB: ret
  loc_00E18BFC: mov eax, Me
  loc_00E18BFF: push eax
  loc_00E18C00: mov ecx, [eax]
  loc_00E18C02: call [ecx+00000008h]
  loc_00E18C05: mov eax, var_4
  loc_00E18C08: mov ecx, var_14
  loc_00E18C0B: pop edi
  loc_00E18C0C: pop esi
  loc_00E18C0D: mov fs:[00000000h], ecx
  loc_00E18C14: pop ebx
  loc_00E18C15: mov esp, ebp
  loc_00E18C17: pop ebp
  loc_00E18C18: retn 0004h
End Sub

Private Sub hapus_UnknownEvent_9 'E19620
  loc_00E19620: push ebp
  loc_00E19621: mov ebp, esp
  loc_00E19623: sub esp, 0000000Ch
  loc_00E19626: push 00402836h ; __vbaExceptHandler
  loc_00E1962B: mov eax, fs:[00000000h]
  loc_00E19631: push eax
  loc_00E19632: mov fs:[00000000h], esp
  loc_00E19639: sub esp, 00000028h
  loc_00E1963C: push ebx
  loc_00E1963D: push esi
  loc_00E1963E: push edi
  loc_00E1963F: mov var_C, esp
  loc_00E19642: mov var_8, 00402350h
  loc_00E19649: mov esi, Me
  loc_00E1964C: mov eax, esi
  loc_00E1964E: and eax, 00000001h
  loc_00E19651: mov var_4, eax
  loc_00E19654: and esi, FFFFFFFEh
  loc_00E19657: push esi
  loc_00E19658: mov Me, esi
  loc_00E1965B: mov ecx, [esi]
  loc_00E1965D: call [ecx+00000004h]
  loc_00E19660: mov edx, [esi]
  loc_00E19662: xor eax, eax
  loc_00E19664: push 006DCBD8h
  loc_00E19669: push eax
  loc_00E1966A: push 00000018h
  loc_00E1966C: push esi
  loc_00E1966D: mov var_18, eax
  loc_00E19670: mov var_1C, eax
  loc_00E19673: mov var_2C, eax
  loc_00E19676: call [edx+00000564h]
  loc_00E1967C: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E19682: push eax
  loc_00E19683: lea eax, var_18
  loc_00E19686: push eax
  loc_00E19687: call ebx
  loc_00E19689: lea ecx, var_2C
  loc_00E1968C: push eax
  loc_00E1968D: push ecx
  loc_00E1968E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E19694: add esp, 00000010h
  loc_00E19697: push eax
  loc_00E19698: call [00401120h] ; __vbaCastObjVar
  loc_00E1969E: lea edx, var_1C
  loc_00E196A1: push eax
  loc_00E196A2: push edx
  loc_00E196A3: call ebx
  loc_00E196A5: mov edi, eax
  loc_00E196A7: push 00000001h
  loc_00E196A9: push edi
  loc_00E196AA: mov eax, [edi]
  loc_00E196AC: call [eax+00000084h]
  loc_00E196B2: test eax, eax
  loc_00E196B4: fnclex
  loc_00E196B6: jge 00E196CAh
  loc_00E196B8: push 00000084h
  loc_00E196BD: push 006DCBE8h
  loc_00E196C2: push edi
  loc_00E196C3: push eax
  loc_00E196C4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E196CA: lea ecx, var_1C
  loc_00E196CD: lea edx, var_18
  loc_00E196D0: push ecx
  loc_00E196D1: push edx
  loc_00E196D2: push 00000002h
  loc_00E196D4: call [00401048h] ; __vbaFreeObjList
  loc_00E196DA: add esp, 0000000Ch
  loc_00E196DD: lea ecx, var_2C
  loc_00E196E0: call [00401024h] ; __vbaFreeVar
  loc_00E196E6: mov eax, [esi]
  loc_00E196E8: push 00000000h
  loc_00E196EA: push FFFFFDDAh
  loc_00E196EF: push esi
  loc_00E196F0: call [eax+00000568h]
  loc_00E196F6: lea ecx, var_18
  loc_00E196F9: push eax
  loc_00E196FA: push ecx
  loc_00E196FB: call ebx
  loc_00E196FD: push eax
  loc_00E196FE: call [00401030h] ; __vbaLateIdCall
  loc_00E19704: add esp, 0000000Ch
  loc_00E19707: lea ecx, var_18
  loc_00E1970A: call [00401254h] ; __vbaFreeObj
  loc_00E19710: mov var_4, 00000000h
  loc_00E19717: push 00E1973Ch
  loc_00E1971C: jmp 00E1973Bh
  loc_00E1971E: lea edx, var_1C
  loc_00E19721: lea eax, var_18
  loc_00E19724: push edx
  loc_00E19725: push eax
  loc_00E19726: push 00000002h
  loc_00E19728: call [00401048h] ; __vbaFreeObjList
  loc_00E1972E: add esp, 0000000Ch
  loc_00E19731: lea ecx, var_2C
  loc_00E19734: call [00401024h] ; __vbaFreeVar
  loc_00E1973A: ret
  loc_00E1973B: ret
  loc_00E1973C: mov eax, Me
  loc_00E1973F: push eax
  loc_00E19740: mov ecx, [eax]
  loc_00E19742: call [ecx+00000008h]
  loc_00E19745: mov eax, var_4
  loc_00E19748: mov ecx, var_14
  loc_00E1974B: pop edi
  loc_00E1974C: pop esi
  loc_00E1974D: mov fs:[00000000h], ecx
  loc_00E19754: pop ebx
  loc_00E19755: mov esp, ebp
  loc_00E19757: pop ebp
  loc_00E19758: retn 0004h
End Sub

Private Sub refreshh_UnknownEvent_9 'E1A200
  loc_00E1A200: push ebp
  loc_00E1A201: mov ebp, esp
  loc_00E1A203: sub esp, 0000000Ch
  loc_00E1A206: push 00402836h ; __vbaExceptHandler
  loc_00E1A20B: mov eax, fs:[00000000h]
  loc_00E1A211: push eax
  loc_00E1A212: mov fs:[00000000h], esp
  loc_00E1A219: sub esp, 00000040h
  loc_00E1A21C: push ebx
  loc_00E1A21D: push esi
  loc_00E1A21E: push edi
  loc_00E1A21F: mov var_C, esp
  loc_00E1A222: mov var_8, 004023B0h
  loc_00E1A229: mov esi, Me
  loc_00E1A22C: mov eax, esi
  loc_00E1A22E: and eax, 00000001h
  loc_00E1A231: mov var_4, eax
  loc_00E1A234: and esi, FFFFFFFEh
  loc_00E1A237: push esi
  loc_00E1A238: mov Me, esi
  loc_00E1A23B: mov ecx, [esi]
  loc_00E1A23D: call [ecx+00000004h]
  loc_00E1A240: sub esp, 00000010h
  loc_00E1A243: mov ecx, 00000008h
  loc_00E1A248: mov edx, esp
  loc_00E1A24A: xor eax, eax
  loc_00E1A24C: mov var_18, eax
  loc_00E1A24F: mov var_1C, eax
  loc_00E1A252: mov [edx], ecx
  loc_00E1A254: mov ecx, var_38
  loc_00E1A257: mov var_2C, eax
  loc_00E1A25A: mov eax, 006E11DCh ; "biaya"
  loc_00E1A25F: mov [edx+00000004h], ecx
  loc_00E1A262: mov ecx, [esi]
  loc_00E1A264: push 0000000Eh
  loc_00E1A266: push esi
  loc_00E1A267: mov [edx+00000008h], eax
  loc_00E1A26A: mov eax, var_30
  loc_00E1A26D: mov [edx+0000000Ch], eax
  loc_00E1A270: call [ecx+00000564h]
  loc_00E1A276: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E1A27C: lea edx, var_18
  loc_00E1A27F: push eax
  loc_00E1A280: push edx
  loc_00E1A281: call edi
  loc_00E1A283: push eax
  loc_00E1A284: call [00401238h] ; __vbaLateIdSt
  loc_00E1A28A: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E1A290: lea ecx, var_18
  loc_00E1A293: call ebx
  loc_00E1A295: mov eax, [esi]
  loc_00E1A297: push 00000000h
  loc_00E1A299: push 00000019h
  loc_00E1A29B: push esi
  loc_00E1A29C: call [eax+00000564h]
  loc_00E1A2A2: lea ecx, var_18
  loc_00E1A2A5: push eax
  loc_00E1A2A6: push ecx
  loc_00E1A2A7: call edi
  loc_00E1A2A9: push eax
  loc_00E1A2AA: call [00401030h] ; __vbaLateIdCall
  loc_00E1A2B0: add esp, 0000000Ch
  loc_00E1A2B3: lea ecx, var_18
  loc_00E1A2B6: call ebx
  loc_00E1A2B8: mov edx, [esi]
  loc_00E1A2BA: push 006E05D8h
  loc_00E1A2BF: push esi
  loc_00E1A2C0: call [edx+00000564h]
  loc_00E1A2C6: push eax
  loc_00E1A2C7: lea eax, var_18
  loc_00E1A2CA: push eax
  loc_00E1A2CB: call edi
  loc_00E1A2CD: push eax
  loc_00E1A2CE: call [00401224h] ; __vbaCastObj
  loc_00E1A2D4: lea ecx, var_2C
  loc_00E1A2D7: mov var_24, eax
  loc_00E1A2DA: push ecx
  loc_00E1A2DB: mov var_2C, 0000000Dh
  loc_00E1A2E2: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E1A2E8: mov ecx, [eax]
  loc_00E1A2EA: sub esp, 00000010h
  loc_00E1A2ED: mov edx, esp
  loc_00E1A2EF: mov [edx], ecx
  loc_00E1A2F1: mov ecx, [eax+00000004h]
  loc_00E1A2F4: mov [edx+00000004h], ecx
  loc_00E1A2F7: mov ecx, [eax+00000008h]
  loc_00E1A2FA: mov eax, [eax+0000000Ch]
  loc_00E1A2FD: mov [edx+00000008h], ecx
  loc_00E1A300: mov [edx+0000000Ch], eax
  loc_00E1A303: mov ecx, [esi]
  loc_00E1A305: push 00000000h
  loc_00E1A307: push 0000002Ah
  loc_00E1A309: push esi
  loc_00E1A30A: call [ecx+00000568h]
  loc_00E1A310: lea edx, var_1C
  loc_00E1A313: push eax
  loc_00E1A314: push edx
  loc_00E1A315: call edi
  loc_00E1A317: push eax
  loc_00E1A318: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E1A31E: lea eax, var_1C
  loc_00E1A321: lea ecx, var_18
  loc_00E1A324: push eax
  loc_00E1A325: push ecx
  loc_00E1A326: push 00000002h
  loc_00E1A328: call [00401048h] ; __vbaFreeObjList
  loc_00E1A32E: add esp, 00000028h
  loc_00E1A331: lea ecx, var_2C
  loc_00E1A334: call [00401024h] ; __vbaFreeVar
  loc_00E1A33A: sub esp, 00000010h
  loc_00E1A33D: mov ecx, 0000000Bh
  loc_00E1A342: mov edx, esp
  loc_00E1A344: xor eax, eax
  loc_00E1A346: push 00000006h
  loc_00E1A348: push esi
  loc_00E1A349: mov [edx], ecx
  loc_00E1A34B: mov ecx, var_38
  loc_00E1A34E: mov [edx+00000004h], ecx
  loc_00E1A351: mov ecx, [esi]
  loc_00E1A353: mov [edx+00000008h], eax
  loc_00E1A356: mov eax, var_30
  loc_00E1A359: mov [edx+0000000Ch], eax
  loc_00E1A35C: call [ecx+00000568h]
  loc_00E1A362: lea edx, var_18
  loc_00E1A365: push eax
  loc_00E1A366: push edx
  loc_00E1A367: call edi
  loc_00E1A369: push eax
  loc_00E1A36A: call [00401238h] ; __vbaLateIdSt
  loc_00E1A370: lea ecx, var_18
  loc_00E1A373: call ebx
  loc_00E1A375: sub esp, 00000010h
  loc_00E1A378: mov ecx, 0000000Bh
  loc_00E1A37D: mov edx, esp
  loc_00E1A37F: xor eax, eax
  loc_00E1A381: push 8001000Eh
  loc_00E1A386: push esi
  loc_00E1A387: mov [edx], ecx
  loc_00E1A389: mov ecx, var_38
  loc_00E1A38C: mov [edx+00000004h], ecx
  loc_00E1A38F: mov ecx, [esi]
  loc_00E1A391: mov [edx+00000008h], eax
  loc_00E1A394: mov eax, var_30
  loc_00E1A397: mov [edx+0000000Ch], eax
  loc_00E1A39A: call [ecx+00000568h]
  loc_00E1A3A0: lea edx, var_18
  loc_00E1A3A3: push eax
  loc_00E1A3A4: push edx
  loc_00E1A3A5: call edi
  loc_00E1A3A7: push eax
  loc_00E1A3A8: call [00401238h] ; __vbaLateIdSt
  loc_00E1A3AE: lea ecx, var_18
  loc_00E1A3B1: call ebx
  loc_00E1A3B3: mov eax, [esi]
  loc_00E1A3B5: push 00000000h
  loc_00E1A3B7: push FFFFFDDAh
  loc_00E1A3BC: push esi
  loc_00E1A3BD: call [eax+00000568h]
  loc_00E1A3C3: lea ecx, var_18
  loc_00E1A3C6: push eax
  loc_00E1A3C7: push ecx
  loc_00E1A3C8: call edi
  loc_00E1A3CA: push eax
  loc_00E1A3CB: call [00401030h] ; __vbaLateIdCall
  loc_00E1A3D1: add esp, 0000000Ch
  loc_00E1A3D4: lea ecx, var_18
  loc_00E1A3D7: call ebx
  loc_00E1A3D9: mov var_4, 00000000h
  loc_00E1A3E0: push 00E1A405h
  loc_00E1A3E5: jmp 00E1A404h
  loc_00E1A3E7: lea edx, var_1C
  loc_00E1A3EA: lea eax, var_18
  loc_00E1A3ED: push edx
  loc_00E1A3EE: push eax
  loc_00E1A3EF: push 00000002h
  loc_00E1A3F1: call [00401048h] ; __vbaFreeObjList
  loc_00E1A3F7: add esp, 0000000Ch
  loc_00E1A3FA: lea ecx, var_2C
  loc_00E1A3FD: call [00401024h] ; __vbaFreeVar
  loc_00E1A403: ret
  loc_00E1A404: ret
  loc_00E1A405: mov eax, Me
  loc_00E1A408: push eax
  loc_00E1A409: mov ecx, [eax]
  loc_00E1A40B: call [ecx+00000008h]
  loc_00E1A40E: mov eax, var_4
  loc_00E1A411: mov ecx, var_14
  loc_00E1A414: pop edi
  loc_00E1A415: pop esi
  loc_00E1A416: mov fs:[00000000h], ecx
  loc_00E1A41D: pop ebx
  loc_00E1A41E: mov esp, ebp
  loc_00E1A420: pop ebp
  loc_00E1A421: retn 0004h
End Sub

Public Function Terbilang(Angka) 'E176F0
  loc_00E176F0: push ebp
  loc_00E176F1: mov ebp, esp
  loc_00E176F3: sub esp, 0000000Ch
  loc_00E176F6: push 00402836h ; __vbaExceptHandler
  loc_00E176FB: mov eax, fs:[00000000h]
  loc_00E17701: push eax
  loc_00E17702: mov fs:[00000000h], esp
  loc_00E17709: sub esp, 00000190h
  loc_00E1770F: push ebx
  loc_00E17710: push esi
  loc_00E17711: push edi
  loc_00E17712: mov var_C, esp
  loc_00E17715: mov var_8, 004022F0h
  loc_00E1771C: xor esi, esi
  loc_00E1771E: mov var_4, esi
  loc_00E17721: mov eax, Me
  loc_00E17724: push eax
  loc_00E17725: mov ecx, [eax]
  loc_00E17727: call [ecx+00000004h]
  loc_00E1772A: mov edx, arg_10
  loc_00E1772D: mov eax, Angka
  loc_00E17730: mov var_28, esi
  loc_00E17733: mov var_2C, esi
  loc_00E17736: mov [edx], esi
  loc_00E17738: mov ecx, [eax]
  loc_00E1773A: push ecx
  loc_00E1773B: mov var_34, esi
  loc_00E1773E: mov var_44, esi
  loc_00E17741: mov var_48, esi
  loc_00E17744: mov var_4C, esi
  loc_00E17747: mov var_50, esi
  loc_00E1774A: mov var_60, esi
  loc_00E1774D: mov var_70, esi
  loc_00E17750: mov var_74, esi
  loc_00E17753: mov var_78, esi
  loc_00E17756: mov var_88, esi
  loc_00E1775C: mov var_8C, esi
  loc_00E17762: mov var_90, esi
  loc_00E17768: mov var_A0, esi
  loc_00E1776E: mov var_B0, esi
  loc_00E17774: mov var_C0, esi
  loc_00E1777A: mov var_D0, esi
  loc_00E17780: mov var_E0, esi
  loc_00E17786: mov var_F0, esi
  loc_00E1778C: mov var_100, esi
  loc_00E17792: mov var_110, esi
  loc_00E17798: mov var_120, esi
  loc_00E1779E: mov var_130, esi
  loc_00E177A4: mov var_140, esi
  loc_00E177AA: mov var_150, esi
  loc_00E177B0: mov var_160, esi
  loc_00E177B6: mov var_170, esi
  loc_00E177BC: mov var_184, esi
  loc_00E177C2: call [0040102Ch] ; __vbaLenBstr
  loc_00E177C8: mov ecx, eax
  loc_00E177CA: call [00401110h] ; __vbaI2I4
  loc_00E177D0: mov esi, [004011D4h] ; __vbaVarCmpEq
  loc_00E177D6: mov var_18C, eax
  loc_00E177DC: mov var_18, 00000001h
  loc_00E177E3: mov dx, var_18C
  loc_00E177EA: cmp var_18, dx
  loc_00E177EE: jg 00E17931h
  loc_00E177F4: movsx edi, var_18
  loc_00E177F8: mov eax, Angka
  loc_00E177FB: lea ecx, var_A0
  loc_00E17801: mov var_128, eax
  loc_00E17807: push ecx
  loc_00E17808: lea edx, var_130
  loc_00E1780E: push edi
  loc_00E1780F: lea eax, var_B0
  loc_00E17815: push edx
  loc_00E17816: push eax
  loc_00E17817: mov var_98, 00000001h
  loc_00E17821: mov var_A0, 00000002h
  loc_00E1782B: mov var_130, 00004008h
  loc_00E17835: call [004010F0h] ; rtcMidCharVar
  loc_00E1783B: lea ecx, var_B0
  loc_00E17841: lea edx, var_150
  loc_00E17847: push ecx
  loc_00E17848: lea eax, var_C0
  loc_00E1784E: push edx
  loc_00E1784F: push eax
  loc_00E17850: mov var_148, 006E0EB0h ; "."
  loc_00E1785A: mov var_150, 00008008h
  loc_00E17864: call __vbaVarCmpEq
  loc_00E17866: lea ecx, var_D0
  loc_00E1786C: push eax
  loc_00E1786D: push ecx
  loc_00E1786E: call [004011B8h] ; __vbaVarNot
  loc_00E17874: push eax
  loc_00E17875: call [004010DCh] ; __vbaBoolVarNull
  loc_00E1787B: mov ebx, eax
  loc_00E1787D: lea edx, var_B0
  loc_00E17883: lea eax, var_A0
  loc_00E17889: push edx
  loc_00E1788A: push eax
  loc_00E1788B: push 00000002h
  loc_00E1788D: call [00401038h] ; __vbaFreeVarList
  loc_00E17893: add esp, 0000000Ch
  loc_00E17896: test bx, bx
  loc_00E17899: jz 00E1791Ah
  loc_00E1789B: mov ecx, Angka
  loc_00E1789E: lea edx, var_A0
  loc_00E178A4: mov var_128, ecx
  loc_00E178AA: push edx
  loc_00E178AB: lea eax, var_130
  loc_00E178B1: push edi
  loc_00E178B2: lea ecx, var_B0
  loc_00E178B8: push eax
  loc_00E178B9: push ecx
  loc_00E178BA: mov var_98, 00000001h
  loc_00E178C4: mov var_A0, 00000002h
  loc_00E178CE: mov var_130, 00004008h
  loc_00E178D8: call [004010F0h] ; rtcMidCharVar
  loc_00E178DE: lea edx, var_60
  loc_00E178E1: lea eax, var_B0
  loc_00E178E7: push edx
  loc_00E178E8: lea ecx, var_C0
  loc_00E178EE: push eax
  loc_00E178EF: push ecx
  loc_00E178F0: call [004011E0h] ; __vbaVarAdd
  loc_00E178F6: mov edx, eax
  loc_00E178F8: lea ecx, var_60
  loc_00E178FB: call [0040101Ch] ; __vbaVarMove
  loc_00E17901: lea edx, var_B0
  loc_00E17907: lea eax, var_A0
  loc_00E1790D: push edx
  loc_00E1790E: push eax
  loc_00E1790F: push 00000002h
  loc_00E17911: call [00401038h] ; __vbaFreeVarList
  loc_00E17917: add esp, 0000000Ch
  loc_00E1791A: mov eax, 00000001h
  loc_00E1791F: add ax, var_18
  loc_00E17923: jo 00E18821h
  loc_00E17929: mov var_18, eax
  loc_00E1792C: jmp 00E177E3h
  loc_00E17931: mov edi, [004010D0h] ; rtcLeftTrimVar
  loc_00E17937: lea ecx, var_60
  loc_00E1793A: lea edx, var_A0
  loc_00E17940: push ecx
  loc_00E17941: push edx
  loc_00E17942: call edi
  loc_00E17944: lea eax, var_A0
  loc_00E1794A: lea ecx, var_B0
  loc_00E17950: push eax
  loc_00E17951: push ecx
  loc_00E17952: mov var_128, 00000000h
  loc_00E1795C: mov var_130, 00008002h
  loc_00E17966: call [00401080h] ; __vbaLenVar
  loc_00E1796C: lea edx, var_130
  loc_00E17972: push eax
  loc_00E17973: push edx
  loc_00E17974: call [00401108h] ; __vbaVarTstEq
  loc_00E1797A: lea ecx, var_A0
  loc_00E17980: mov ebx, eax
  loc_00E17982: call [00401024h] ; __vbaFreeVar
  loc_00E17988: test bx, bx
  loc_00E1798B: jz 00E179B5h
  loc_00E1798D: lea edx, var_130
  loc_00E17993: lea ecx, var_44
  loc_00E17996: mov var_128, 006E0EB8h ; "Nol Rupiah"
  loc_00E179A0: mov var_130, 00000008h
  loc_00E179AA: call [00401204h] ; __vbaVarCopy
  loc_00E179B0: jmp 00E18735h
  loc_00E179B5: lea edx, var_60
  loc_00E179B8: lea ecx, var_A0
  loc_00E179BE: call [00401204h] ; __vbaVarCopy
  loc_00E179C4: lea eax, var_A0
  loc_00E179CA: lea ecx, var_B0
  loc_00E179D0: push eax
  loc_00E179D1: push ecx
  loc_00E179D2: call [004010E4h] ; rtcRightTrimVar
  loc_00E179D8: lea edx, var_B0
  loc_00E179DE: lea eax, var_C0
  loc_00E179E4: push edx
  loc_00E179E5: push eax
  loc_00E179E6: call edi
  loc_00E179E8: lea ecx, var_C0
  loc_00E179EE: push ecx
  loc_00E179EF: call [00401034h] ; __vbaStrVarMove
  loc_00E179F5: mov edx, eax
  loc_00E179F7: lea ecx, var_34
  loc_00E179FA: call [00401228h] ; __vbaStrMove
  loc_00E17A00: mov ebx, [00401038h] ; __vbaFreeVarList
  loc_00E17A06: lea edx, var_C0
  loc_00E17A0C: lea eax, var_B0
  loc_00E17A12: push edx
  loc_00E17A13: lea ecx, var_A0
  loc_00E17A19: push eax
  loc_00E17A1A: push ecx
  loc_00E17A1B: push 00000003h
  loc_00E17A1D: call ebx
  loc_00E17A1F: mov edx, Angka
  loc_00E17A22: add esp, 00000010h
  loc_00E17A25: lea eax, var_A0
  loc_00E17A2B: mov var_128, edx
  loc_00E17A31: push eax
  loc_00E17A32: lea ecx, var_130
  loc_00E17A38: push 0000000Fh
  loc_00E17A3A: lea edx, var_B0
  loc_00E17A40: mov edi, 00000002h
  loc_00E17A45: push ecx
  loc_00E17A46: push edx
  loc_00E17A47: mov var_98, edi
  loc_00E17A4D: mov var_A0, edi
  loc_00E17A53: mov var_130, 00004008h
  loc_00E17A5D: call [004010F0h] ; rtcMidCharVar
  loc_00E17A63: lea eax, var_B0
  loc_00E17A69: push edi
  loc_00E17A6A: lea ecx, var_C0
  loc_00E17A70: push eax
  loc_00E17A71: push ecx
  loc_00E17A72: call [0040122Ch] ; rtcRightCharVar
  loc_00E17A78: lea edx, var_C0
  loc_00E17A7E: lea eax, var_90
  loc_00E17A84: push edx
  loc_00E17A85: push eax
  loc_00E17A86: call [00401180h] ; __vbaStrVarVal
  loc_00E17A8C: push eax
  loc_00E17A8D: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E17A93: call [00401200h] ; __vbaFpI2
  loc_00E17A99: lea ecx, var_90
  loc_00E17A9F: mov edi, eax
  loc_00E17AA1: call [00401258h] ; __vbaFreeStr
  loc_00E17AA7: lea ecx, var_C0
  loc_00E17AAD: lea edx, var_B0
  loc_00E17AB3: push ecx
  loc_00E17AB4: lea eax, var_A0
  loc_00E17ABA: push edx
  loc_00E17ABB: push eax
  loc_00E17ABC: push 00000003h
  loc_00E17ABE: call ebx
  loc_00E17AC0: add esp, 00000010h
  loc_00E17AC3: test di, di
  loc_00E17AC6: jnz 00E17AD6h
  loc_00E17AC8: mov edx, 006DCC80h
  loc_00E17ACD: lea ecx, var_48
  loc_00E17AD0: call [004011B0h] ; __vbaStrCopy
  loc_00E17AD6: mov edi, [0040101Ch] ; __vbaVarMove
  loc_00E17ADC: mov ebx, 00000002h
  loc_00E17AE1: lea edx, var_130
  loc_00E17AE7: lea ecx, var_88
  loc_00E17AED: mov var_128, 00000000h
  loc_00E17AF7: mov var_130, ebx
  loc_00E17AFD: call edi
  loc_00E17AFF: lea edx, var_130
  loc_00E17B05: lea ecx, var_70
  loc_00E17B08: mov var_128, 00000000h
  loc_00E17B12: mov var_130, ebx
  loc_00E17B18: call edi
  loc_00E17B1A: mov edx, 006DCC80h
  loc_00E17B1F: lea ecx, var_2C
  loc_00E17B22: call [004011B0h] ; __vbaStrCopy
  loc_00E17B28: mov edi, [00401118h] ; __vbaVarOr
  loc_00E17B2E: mov ebx, 00008002h
  loc_00E17B33: mov ecx, var_34
  loc_00E17B36: push ecx
  loc_00E17B37: call [0040102Ch] ; __vbaLenBstr
  loc_00E17B3D: mov var_128, eax
  loc_00E17B43: lea edx, var_88
  loc_00E17B49: lea eax, var_130
  loc_00E17B4F: push edx
  loc_00E17B50: push eax
  loc_00E17B51: mov var_130, 00008003h
  loc_00E17B5B: call [004010D4h] ; __vbaVarTstLt
  loc_00E17B61: test ax, ax
  loc_00E17B64: jz 00E185B1h
  loc_00E17B6A: lea ecx, var_88
  loc_00E17B70: lea edx, var_130
  loc_00E17B76: push ecx
  loc_00E17B77: lea eax, var_A0
  loc_00E17B7D: push edx
  loc_00E17B7E: push eax
  loc_00E17B7F: mov var_128, 00000001h
  loc_00E17B89: mov var_130, 00000002h
  loc_00E17B93: call [004011E0h] ; __vbaVarAdd
  loc_00E17B99: mov edx, eax
  loc_00E17B9B: lea ecx, var_88
  loc_00E17BA1: call [0040101Ch] ; __vbaVarMove
  loc_00E17BA7: lea edx, var_A0
  loc_00E17BAD: lea eax, var_88
  loc_00E17BB3: lea ecx, var_34
  loc_00E17BB6: push edx
  loc_00E17BB7: push eax
  loc_00E17BB8: mov var_98, 00000001h
  loc_00E17BC2: mov var_A0, 00000002h
  loc_00E17BCC: mov var_128, ecx
  loc_00E17BD2: mov var_130, 00004008h
  loc_00E17BDC: call [004011D0h] ; __vbaI4Var
  loc_00E17BE2: lea ecx, var_130
  loc_00E17BE8: push eax
  loc_00E17BE9: lea edx, var_B0
  loc_00E17BEF: push ecx
  loc_00E17BF0: push edx
  loc_00E17BF1: call [004010F0h] ; rtcMidCharVar
  loc_00E17BF7: lea eax, var_B0
  loc_00E17BFD: push eax
  loc_00E17BFE: call [00401034h] ; __vbaStrVarMove
  loc_00E17C04: mov edx, eax
  loc_00E17C06: lea ecx, var_8C
  loc_00E17C0C: call [00401228h] ; __vbaStrMove
  loc_00E17C12: lea ecx, var_B0
  loc_00E17C18: lea edx, var_A0
  loc_00E17C1E: push ecx
  loc_00E17C1F: push edx
  loc_00E17C20: push 00000002h
  loc_00E17C22: call [00401038h] ; __vbaFreeVarList
  loc_00E17C28: mov eax, var_8C
  loc_00E17C2E: add esp, 0000000Ch
  loc_00E17C31: push eax
  loc_00E17C32: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E17C38: fstp real8 ptr var_128
  loc_00E17C3E: lea ecx, var_70
  loc_00E17C41: lea edx, var_130
  loc_00E17C47: push ecx
  loc_00E17C48: lea eax, var_A0
  loc_00E17C4E: push edx
  loc_00E17C4F: push eax
  loc_00E17C50: mov var_130, 00000005h
  loc_00E17C5A: call [004011E0h] ; __vbaVarAdd
  loc_00E17C60: mov edx, eax
  loc_00E17C62: lea ecx, var_70
  loc_00E17C65: call [0040101Ch] ; __vbaVarMove
  loc_00E17C6B: mov ecx, var_34
  loc_00E17C6E: push ecx
  loc_00E17C6F: call [0040102Ch] ; __vbaLenBstr
  loc_00E17C75: mov var_128, eax
  loc_00E17C7B: lea edx, var_130
  loc_00E17C81: lea eax, var_88
  loc_00E17C87: push edx
  loc_00E17C88: lea ecx, var_A0
  loc_00E17C8E: push eax
  loc_00E17C8F: push ecx
  loc_00E17C90: mov var_130, 00000003h
  loc_00E17C9A: mov var_138, 00000001h
  loc_00E17CA4: mov var_140, 00000002h
  loc_00E17CAE: call [00401004h] ; __vbaVarSub
  loc_00E17CB4: push eax
  loc_00E17CB5: lea edx, var_140
  loc_00E17CBB: lea eax, var_B0
  loc_00E17CC1: push edx
  loc_00E17CC2: push eax
  loc_00E17CC3: call [004011E0h] ; __vbaVarAdd
  loc_00E17CC9: mov edx, eax
  loc_00E17CCB: lea ecx, var_28
  loc_00E17CCE: call [0040101Ch] ; __vbaVarMove
  loc_00E17CD4: mov ecx, var_8C
  loc_00E17CDA: push ecx
  loc_00E17CDB: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E17CE1: fstp real8 ptr var_194
  loc_00E17CE7: mov eax, var_194
  loc_00E17CED: mov ecx, var_190
  loc_00E17CF3: test eax, eax
  loc_00E17CF5: jnz 00E1816Ch
  loc_00E17CFB: cmp ecx, 3FF00000h
  loc_00E17D01: jnz 00E1816Ch
  loc_00E17D07: lea edx, var_28
  loc_00E17D0A: lea eax, var_130
  loc_00E17D10: push edx
  loc_00E17D11: lea ecx, var_A0
  loc_00E17D17: push eax
  loc_00E17D18: push ecx
  loc_00E17D19: mov var_128, 00000001h
  loc_00E17D23: mov var_130, ebx
  loc_00E17D29: mov var_138, 00000007h
  loc_00E17D33: mov var_140, ebx
  loc_00E17D39: mov var_148, 0000000Ah
  loc_00E17D43: mov var_150, ebx
  loc_00E17D49: mov var_158, 0000000Dh
  loc_00E17D53: mov var_160, ebx
  loc_00E17D59: call __vbaVarCmpEq
  loc_00E17D5B: push eax
  loc_00E17D5C: lea edx, var_28
  loc_00E17D5F: lea eax, var_140
  loc_00E17D65: push edx
  loc_00E17D66: lea ecx, var_B0
  loc_00E17D6C: push eax
  loc_00E17D6D: push ecx
  loc_00E17D6E: call __vbaVarCmpEq
  loc_00E17D70: lea edx, var_C0
  loc_00E17D76: push eax
  loc_00E17D77: push edx
  loc_00E17D78: call edi
  loc_00E17D7A: push eax
  loc_00E17D7B: lea eax, var_28
  loc_00E17D7E: lea ecx, var_150
  loc_00E17D84: push eax
  loc_00E17D85: lea edx, var_D0
  loc_00E17D8B: push ecx
  loc_00E17D8C: push edx
  loc_00E17D8D: call __vbaVarCmpEq
  loc_00E17D8F: push eax
  loc_00E17D90: lea eax, var_E0
  loc_00E17D96: push eax
  loc_00E17D97: call edi
  loc_00E17D99: lea ecx, var_28
  loc_00E17D9C: push eax
  loc_00E17D9D: lea edx, var_160
  loc_00E17DA3: push ecx
  loc_00E17DA4: lea eax, var_F0
  loc_00E17DAA: push edx
  loc_00E17DAB: push eax
  loc_00E17DAC: call __vbaVarCmpEq
  loc_00E17DAE: lea ecx, var_100
  loc_00E17DB4: push eax
  loc_00E17DB5: push ecx
  loc_00E17DB6: call edi
  loc_00E17DB8: push eax
  loc_00E17DB9: call [004010DCh] ; __vbaBoolVarNull
  loc_00E17DBF: test ax, ax
  loc_00E17DC2: jz 00E17DCEh
  loc_00E17DC4: mov edx, 006E0ED4h ; "Satu "
  loc_00E17DC9: jmp 00E1820Eh
  loc_00E17DCE: lea edx, var_28
  loc_00E17DD1: lea eax, var_130
  loc_00E17DD7: push edx
  loc_00E17DD8: push eax
  loc_00E17DD9: mov var_128, 00000004h
  loc_00E17DE3: mov var_130, ebx
  loc_00E17DE9: call [00401108h] ; __vbaVarTstEq
  loc_00E17DEF: test ax, ax
  loc_00E17DF2: jz 00E17E27h
  loc_00E17DF4: lea ecx, var_88
  loc_00E17DFA: lea edx, var_130
  loc_00E17E00: push ecx
  loc_00E17E01: push edx
  loc_00E17E02: mov var_128, 00000001h
  loc_00E17E0C: mov var_130, ebx
  loc_00E17E12: call [00401108h] ; __vbaVarTstEq
  loc_00E17E18: test ax, ax
  loc_00E17E1B: jz 00E17DC4h
  loc_00E17E1D: mov edx, 006E0EE4h ; "Se"
  loc_00E17E22: jmp 00E1820Eh
  loc_00E17E27: lea eax, var_28
  loc_00E17E2A: lea ecx, var_130
  loc_00E17E30: push eax
  loc_00E17E31: lea edx, var_A0
  loc_00E17E37: push ecx
  loc_00E17E38: push edx
  loc_00E17E39: mov var_128, 00000002h
  loc_00E17E43: mov var_130, ebx
  loc_00E17E49: mov var_138, 00000005h
  loc_00E17E53: mov var_140, ebx
  loc_00E17E59: mov var_148, 00000008h
  loc_00E17E63: mov var_150, ebx
  loc_00E17E69: mov var_158, 0000000Bh
  loc_00E17E73: mov var_160, ebx
  loc_00E17E79: mov var_168, 0000000Eh
  loc_00E17E83: mov var_170, ebx
  loc_00E17E89: call __vbaVarCmpEq
  loc_00E17E8B: push eax
  loc_00E17E8C: lea eax, var_28
  loc_00E17E8F: lea ecx, var_140
  loc_00E17E95: push eax
  loc_00E17E96: lea edx, var_B0
  loc_00E17E9C: push ecx
  loc_00E17E9D: push edx
  loc_00E17E9E: call __vbaVarCmpEq
  loc_00E17EA0: push eax
  loc_00E17EA1: lea eax, var_C0
  loc_00E17EA7: push eax
  loc_00E17EA8: call edi
  loc_00E17EAA: lea ecx, var_28
  loc_00E17EAD: push eax
  loc_00E17EAE: lea edx, var_150
  loc_00E17EB4: push ecx
  loc_00E17EB5: lea eax, var_D0
  loc_00E17EBB: push edx
  loc_00E17EBC: push eax
  loc_00E17EBD: call __vbaVarCmpEq
  loc_00E17EBF: lea ecx, var_E0
  loc_00E17EC5: push eax
  loc_00E17EC6: push ecx
  loc_00E17EC7: call edi
  loc_00E17EC9: push eax
  loc_00E17ECA: lea edx, var_28
  loc_00E17ECD: lea eax, var_160
  loc_00E17ED3: push edx
  loc_00E17ED4: lea ecx, var_F0
  loc_00E17EDA: push eax
  loc_00E17EDB: push ecx
  loc_00E17EDC: call __vbaVarCmpEq
  loc_00E17EDE: lea edx, var_100
  loc_00E17EE4: push eax
  loc_00E17EE5: push edx
  loc_00E17EE6: call edi
  loc_00E17EE8: push eax
  loc_00E17EE9: lea eax, var_28
  loc_00E17EEC: lea ecx, var_170
  loc_00E17EF2: push eax
  loc_00E17EF3: lea edx, var_110
  loc_00E17EF9: push ecx
  loc_00E17EFA: push edx
  loc_00E17EFB: call __vbaVarCmpEq
  loc_00E17EFD: push eax
  loc_00E17EFE: lea eax, var_120
  loc_00E17F04: push eax
  loc_00E17F05: call edi
  loc_00E17F07: push eax
  loc_00E17F08: call [004010DCh] ; __vbaBoolVarNull
  loc_00E17F0E: test ax, ax
  loc_00E17F11: jz 00E17E1Dh
  loc_00E17F17: lea ecx, var_88
  loc_00E17F1D: lea edx, var_130
  loc_00E17F23: push ecx
  loc_00E17F24: lea eax, var_A0
  loc_00E17F2A: push edx
  loc_00E17F2B: push eax
  loc_00E17F2C: mov var_128, 00000001h
  loc_00E17F36: mov var_130, 00000002h
  loc_00E17F40: call [004011E0h] ; __vbaVarAdd
  loc_00E17F46: mov edx, eax
  loc_00E17F48: lea ecx, var_88
  loc_00E17F4E: call [0040101Ch] ; __vbaVarMove
  loc_00E17F54: lea edx, var_A0
  loc_00E17F5A: lea eax, var_88
  loc_00E17F60: lea ecx, var_34
  loc_00E17F63: push edx
  loc_00E17F64: push eax
  loc_00E17F65: mov var_98, 00000001h
  loc_00E17F6F: mov var_A0, 00000002h
  loc_00E17F79: mov var_128, ecx
  loc_00E17F7F: mov var_130, 00004008h
  loc_00E17F89: call [004011D0h] ; __vbaI4Var
  loc_00E17F8F: lea ecx, var_130
  loc_00E17F95: push eax
  loc_00E17F96: lea edx, var_B0
  loc_00E17F9C: push ecx
  loc_00E17F9D: push edx
  loc_00E17F9E: call [004010F0h] ; rtcMidCharVar
  loc_00E17FA4: lea eax, var_B0
  loc_00E17FAA: push eax
  loc_00E17FAB: call [00401034h] ; __vbaStrVarMove
  loc_00E17FB1: mov edx, eax
  loc_00E17FB3: lea ecx, var_8C
  loc_00E17FB9: call [00401228h] ; __vbaStrMove
  loc_00E17FBF: lea ecx, var_B0
  loc_00E17FC5: lea edx, var_A0
  loc_00E17FCB: push ecx
  loc_00E17FCC: push edx
  loc_00E17FCD: push 00000002h
  loc_00E17FCF: call [00401038h] ; __vbaFreeVarList
  loc_00E17FD5: mov eax, var_34
  loc_00E17FD8: add esp, 0000000Ch
  loc_00E17FDB: push eax
  loc_00E17FDC: call [0040102Ch] ; __vbaLenBstr
  loc_00E17FE2: lea ecx, var_130
  loc_00E17FE8: mov var_128, eax
  loc_00E17FEE: lea edx, var_88
  loc_00E17FF4: push ecx
  loc_00E17FF5: lea eax, var_A0
  loc_00E17FFB: push edx
  loc_00E17FFC: push eax
  loc_00E17FFD: mov var_130, 00000003h
  loc_00E18007: mov var_138, 00000001h
  loc_00E18011: mov var_140, 00000002h
  loc_00E1801B: call [00401004h] ; __vbaVarSub
  loc_00E18021: lea ecx, var_140
  loc_00E18027: push eax
  loc_00E18028: lea edx, var_B0
  loc_00E1802E: push ecx
  loc_00E1802F: push edx
  loc_00E18030: call [004011E0h] ; __vbaVarAdd
  loc_00E18036: mov edx, eax
  loc_00E18038: lea ecx, var_28
  loc_00E1803B: call [0040101Ch] ; __vbaVarMove
  loc_00E18041: mov edx, 006DCC80h
  loc_00E18046: lea ecx, var_50
  loc_00E18049: call [004011B0h] ; __vbaStrCopy
  loc_00E1804F: mov eax, var_8C
  loc_00E18055: push eax
  loc_00E18056: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E1805C: fstp real8 ptr var_19C
  loc_00E18062: fld real8 ptr var_19C
  loc_00E18068: fcomp real8 ptr [004022E8h]
  loc_00E1806E: fnstsw ax
  loc_00E18070: test ah, 40h
  loc_00E18073: jz 00E1807Fh
  loc_00E18075: mov edx, 006E0EF0h ; "Sepuluh "
  loc_00E1807A: jmp 00E1820Eh
  loc_00E1807F: mov ecx, var_19C
  loc_00E18085: mov eax, var_198
  loc_00E1808B: test ecx, ecx
  loc_00E1808D: jnz 00E180A0h
  loc_00E1808F: cmp eax, 3FF00000h
  loc_00E18094: jnz 00E180A0h
  loc_00E18096: mov edx, 006E0F08h ; "Sebelas "
  loc_00E1809B: jmp 00E1820Eh
  loc_00E180A0: test ecx, ecx
  loc_00E180A2: jnz 00E18217h
  loc_00E180A8: cmp eax, 40000000h
  loc_00E180AD: jnz 00E180B9h
  loc_00E180AF: mov edx, 006E0F20h ; "Dua Belas "
  loc_00E180B4: jmp 00E1820Eh
  loc_00E180B9: test ecx, ecx
  loc_00E180BB: jnz 00E18217h
  loc_00E180C1: cmp eax, 40080000h
  loc_00E180C6: jnz 00E180D2h
  loc_00E180C8: mov edx, 006E0F4Ch ; "Tiga Belas "
  loc_00E180CD: jmp 00E1820Eh
  loc_00E180D2: test ecx, ecx
  loc_00E180D4: jnz 00E18217h
  loc_00E180DA: cmp eax, 40100000h
  loc_00E180DF: jnz 00E180EBh
  loc_00E180E1: mov edx, 006E0F68h ; "Empat Belas "
  loc_00E180E6: jmp 00E1820Eh
  loc_00E180EB: test ecx, ecx
  loc_00E180ED: jnz 00E18217h
  loc_00E180F3: cmp eax, 40140000h
  loc_00E180F8: jnz 00E18104h
  loc_00E180FA: mov edx, 006E0F88h ; "Lima Belas "
  loc_00E180FF: jmp 00E1820Eh
  loc_00E18104: test ecx, ecx
  loc_00E18106: jnz 00E18217h
  loc_00E1810C: cmp eax, 40180000h
  loc_00E18111: jnz 00E1811Dh
  loc_00E18113: mov edx, 006E0FA4h ; "Enam Belas "
  loc_00E18118: jmp 00E1820Eh
  loc_00E1811D: test ecx, ecx
  loc_00E1811F: jnz 00E18217h
  loc_00E18125: cmp eax, 401C0000h
  loc_00E1812A: jnz 00E18136h
  loc_00E1812C: mov edx, 006E0FC0h ; "Tujuh Belas "
  loc_00E18131: jmp 00E1820Eh
  loc_00E18136: test ecx, ecx
  loc_00E18138: jnz 00E18217h
  loc_00E1813E: cmp eax, 40200000h
  loc_00E18143: jnz 00E1814Fh
  loc_00E18145: mov edx, 006E0FE0h ; "Delapan Belas "
  loc_00E1814A: jmp 00E1820Eh
  loc_00E1814F: test ecx, ecx
  loc_00E18151: jnz 00E18217h
  loc_00E18157: cmp eax, 40220000h
  loc_00E1815C: jnz 00E18217h
  loc_00E18162: mov edx, 006E1004h ; "Sembilan Belas "
  loc_00E18167: jmp 00E1820Eh
  loc_00E1816C: test eax, eax
  loc_00E1816E: jnz 00E18209h
  loc_00E18174: cmp ecx, 40000000h
  loc_00E1817A: jnz 00E18186h
  loc_00E1817C: mov edx, 006E1028h ; "Dua "
  loc_00E18181: jmp 00E1820Eh
  loc_00E18186: test eax, eax
  loc_00E18188: jnz 00E18209h
  loc_00E1818A: cmp ecx, 40080000h
  loc_00E18190: jnz 00E18199h
  loc_00E18192: mov edx, 006E1038h ; "Tiga "
  loc_00E18197: jmp 00E1820Eh
  loc_00E18199: test eax, eax
  loc_00E1819B: jnz 00E18209h
  loc_00E1819D: cmp ecx, 40100000h
  loc_00E181A3: jnz 00E181ACh
  loc_00E181A5: mov edx, 006E1048h ; "Empat "
  loc_00E181AA: jmp 00E1820Eh
  loc_00E181AC: test eax, eax
  loc_00E181AE: jnz 00E18209h
  loc_00E181B0: cmp ecx, 40140000h
  loc_00E181B6: jnz 00E181BFh
  loc_00E181B8: mov edx, 006E105Ch ; "Lima "
  loc_00E181BD: jmp 00E1820Eh
  loc_00E181BF: test eax, eax
  loc_00E181C1: jnz 00E18209h
  loc_00E181C3: cmp ecx, 40180000h
  loc_00E181C9: jnz 00E181D2h
  loc_00E181CB: mov edx, 006E106Ch ; "Enam "
  loc_00E181D0: jmp 00E1820Eh
  loc_00E181D2: test eax, eax
  loc_00E181D4: jnz 00E18209h
  loc_00E181D6: cmp ecx, 401C0000h
  loc_00E181DC: jnz 00E181E5h
  loc_00E181DE: mov edx, 006E107Ch ; "Tujuh "
  loc_00E181E3: jmp 00E1820Eh
  loc_00E181E5: test eax, eax
  loc_00E181E7: jnz 00E18209h
  loc_00E181E9: cmp ecx, 40200000h
  loc_00E181EF: jnz 00E181F8h
  loc_00E181F1: mov edx, 006E1090h ; "Delapan "
  loc_00E181F6: jmp 00E1820Eh
  loc_00E181F8: test eax, eax
  loc_00E181FA: jnz 00E18209h
  loc_00E181FC: cmp ecx, 40220000h
  loc_00E18202: mov edx, 006E10A8h ; "Sembilan "
  loc_00E18207: jz 00E1820Eh
  loc_00E18209: mov edx, 006DCC80h
  loc_00E1820E: lea ecx, var_4C
  loc_00E18211: call [004011B0h] ; __vbaStrCopy
  loc_00E18217: mov ecx, var_8C
  loc_00E1821D: push ecx
  loc_00E1821E: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E18224: call [004010D8h] ; __vbaFpR8
  loc_00E1822A: fcomp real8 ptr [004022E8h]
  loc_00E18230: fnstsw ax
  loc_00E18232: test ah, 41h
  loc_00E18235: jnz 00E18422h
  loc_00E1823B: lea edx, var_28
  loc_00E1823E: lea eax, var_130
  loc_00E18244: push edx
  loc_00E18245: lea ecx, var_A0
  loc_00E1824B: push eax
  loc_00E1824C: push ecx
  loc_00E1824D: mov var_128, 00000002h
  loc_00E18257: mov var_130, ebx
  loc_00E1825D: mov var_138, 00000005h
  loc_00E18267: mov var_140, ebx
  loc_00E1826D: mov var_148, 00000008h
  loc_00E18277: mov var_150, ebx
  loc_00E1827D: mov var_158, 0000000Bh
  loc_00E18287: mov var_160, ebx
  loc_00E1828D: mov var_168, 0000000Eh
  loc_00E18297: mov var_170, ebx
  loc_00E1829D: call __vbaVarCmpEq
  loc_00E1829F: push eax
  loc_00E182A0: lea edx, var_28
  loc_00E182A3: lea eax, var_140
  loc_00E182A9: push edx
  loc_00E182AA: lea ecx, var_B0
  loc_00E182B0: push eax
  loc_00E182B1: push ecx
  loc_00E182B2: call __vbaVarCmpEq
  loc_00E182B4: lea edx, var_C0
  loc_00E182BA: push eax
  loc_00E182BB: push edx
  loc_00E182BC: call edi
  loc_00E182BE: push eax
  loc_00E182BF: lea eax, var_28
  loc_00E182C2: lea ecx, var_150
  loc_00E182C8: push eax
  loc_00E182C9: lea edx, var_D0
  loc_00E182CF: push ecx
  loc_00E182D0: push edx
  loc_00E182D1: call __vbaVarCmpEq
  loc_00E182D3: push eax
  loc_00E182D4: lea eax, var_E0
  loc_00E182DA: push eax
  loc_00E182DB: call edi
  loc_00E182DD: lea ecx, var_28
  loc_00E182E0: push eax
  loc_00E182E1: lea edx, var_160
  loc_00E182E7: push ecx
  loc_00E182E8: lea eax, var_F0
  loc_00E182EE: push edx
  loc_00E182EF: push eax
  loc_00E182F0: call __vbaVarCmpEq
  loc_00E182F2: lea ecx, var_100
  loc_00E182F8: push eax
  loc_00E182F9: push ecx
  loc_00E182FA: call edi
  loc_00E182FC: push eax
  loc_00E182FD: lea edx, var_28
  loc_00E18300: lea eax, var_170
  loc_00E18306: push edx
  loc_00E18307: lea ecx, var_110
  loc_00E1830D: push eax
  loc_00E1830E: push ecx
  loc_00E1830F: call __vbaVarCmpEq
  loc_00E18311: lea edx, var_120
  loc_00E18317: push eax
  loc_00E18318: push edx
  loc_00E18319: call edi
  loc_00E1831B: push eax
  loc_00E1831C: call [004010DCh] ; __vbaBoolVarNull
  loc_00E18322: test ax, ax
  loc_00E18325: jz 00E18331h
  loc_00E18327: mov edx, 006E10C0h ; "Puluh "
  loc_00E1832C: jmp 00E18427h
  loc_00E18331: lea eax, var_28
  loc_00E18334: lea ecx, var_130
  loc_00E1833A: push eax
  loc_00E1833B: lea edx, var_A0
  loc_00E18341: push ecx
  loc_00E18342: push edx
  loc_00E18343: mov var_128, 00000003h
  loc_00E1834D: mov var_130, ebx
  loc_00E18353: mov var_138, 00000006h
  loc_00E1835D: mov var_140, ebx
  loc_00E18363: mov var_148, 00000009h
  loc_00E1836D: mov var_150, ebx
  loc_00E18373: mov var_158, 0000000Ch
  loc_00E1837D: mov var_160, ebx
  loc_00E18383: mov var_168, 0000000Fh
  loc_00E1838D: mov var_170, ebx
  loc_00E18393: call __vbaVarCmpEq
  loc_00E18395: push eax
  loc_00E18396: lea eax, var_28
  loc_00E18399: lea ecx, var_140
  loc_00E1839F: push eax
  loc_00E183A0: lea edx, var_B0
  loc_00E183A6: push ecx
  loc_00E183A7: push edx
  loc_00E183A8: call __vbaVarCmpEq
  loc_00E183AA: push eax
  loc_00E183AB: lea eax, var_C0
  loc_00E183B1: push eax
  loc_00E183B2: call edi
  loc_00E183B4: lea ecx, var_28
  loc_00E183B7: push eax
  loc_00E183B8: lea edx, var_150
  loc_00E183BE: push ecx
  loc_00E183BF: lea eax, var_D0
  loc_00E183C5: push edx
  loc_00E183C6: push eax
  loc_00E183C7: call __vbaVarCmpEq
  loc_00E183C9: lea ecx, var_E0
  loc_00E183CF: push eax
  loc_00E183D0: push ecx
  loc_00E183D1: call edi
  loc_00E183D3: push eax
  loc_00E183D4: lea edx, var_28
  loc_00E183D7: lea eax, var_160
  loc_00E183DD: push edx
  loc_00E183DE: lea ecx, var_F0
  loc_00E183E4: push eax
  loc_00E183E5: push ecx
  loc_00E183E6: call __vbaVarCmpEq
  loc_00E183E8: lea edx, var_100
  loc_00E183EE: push eax
  loc_00E183EF: push edx
  loc_00E183F0: call edi
  loc_00E183F2: push eax
  loc_00E183F3: lea eax, var_28
  loc_00E183F6: lea ecx, var_170
  loc_00E183FC: push eax
  loc_00E183FD: lea edx, var_110
  loc_00E18403: push ecx
  loc_00E18404: push edx
  loc_00E18405: call __vbaVarCmpEq
  loc_00E18407: push eax
  loc_00E18408: lea eax, var_120
  loc_00E1840E: push eax
  loc_00E1840F: call edi
  loc_00E18411: push eax
  loc_00E18412: call [004010DCh] ; __vbaBoolVarNull
  loc_00E18418: test ax, ax
  loc_00E1841B: mov edx, 006E10D4h ; "Ratus "
  loc_00E18420: jnz 00E18427h
  loc_00E18422: mov edx, 006DCC80h
  loc_00E18427: lea ecx, var_50
  loc_00E1842A: call [004011B0h] ; __vbaStrCopy
  loc_00E18430: lea ecx, var_70
  loc_00E18433: lea edx, var_130
  loc_00E18439: push ecx
  loc_00E1843A: push edx
  loc_00E1843B: mov var_128, 00000000h
  loc_00E18445: mov var_130, ebx
  loc_00E1844B: call [00401008h] ; __vbaVarTstGt
  loc_00E18451: test ax, ax
  loc_00E18454: jz 00E1856Eh
  loc_00E1845A: lea edx, var_28
  loc_00E1845D: lea ecx, var_184
  loc_00E18463: call [00401204h] ; __vbaVarCopy
  loc_00E18469: lea eax, var_184
  loc_00E1846F: lea ecx, var_130
  loc_00E18475: push eax
  loc_00E18476: push ecx
  loc_00E18477: mov var_128, 00000004h
  loc_00E18481: mov var_130, ebx
  loc_00E18487: call [00401108h] ; __vbaVarTstEq
  loc_00E1848D: test ax, ax
  loc_00E18490: jz 00E184A0h
  loc_00E18492: mov edx, var_50
  loc_00E18495: push edx
  loc_00E18496: push 006E10E8h ; "Ribu "
  loc_00E1849B: jmp 00E1853Ah
  loc_00E184A0: lea eax, var_184
  loc_00E184A6: lea ecx, var_130
  loc_00E184AC: push eax
  loc_00E184AD: push ecx
  loc_00E184AE: mov var_128, 00000007h
  loc_00E184B8: mov var_130, ebx
  loc_00E184BE: call [00401108h] ; __vbaVarTstEq
  loc_00E184C4: test ax, ax
  loc_00E184C7: jz 00E184D4h
  loc_00E184C9: mov edx, var_50
  loc_00E184CC: push edx
  loc_00E184CD: push 006E0F3Ch ; "Juta "
  loc_00E184D2: jmp 00E1853Ah
  loc_00E184D4: lea eax, var_184
  loc_00E184DA: lea ecx, var_130
  loc_00E184E0: push eax
  loc_00E184E1: push ecx
  loc_00E184E2: mov var_128, 0000000Ah
  loc_00E184EC: mov var_130, ebx
  loc_00E184F2: call [00401108h] ; __vbaVarTstEq
  loc_00E184F8: test ax, ax
  loc_00E184FB: jz 00E18508h
  loc_00E184FD: mov edx, var_50
  loc_00E18500: push edx
  loc_00E18501: push 006E10F8h ; "Milyar "
  loc_00E18506: jmp 00E1853Ah
  loc_00E18508: lea eax, var_184
  loc_00E1850E: lea ecx, var_130
  loc_00E18514: push eax
  loc_00E18515: push ecx
  loc_00E18516: mov var_128, 0000000Dh
  loc_00E18520: mov var_130, ebx
  loc_00E18526: call [00401108h] ; __vbaVarTstEq
  loc_00E1852C: test ax, ax
  loc_00E1852F: jz 00E1856Eh
  loc_00E18531: mov edx, var_50
  loc_00E18534: push edx
  loc_00E18535: push 006E110Ch ; "Trilyun "
  loc_00E1853A: call [0040106Ch] ; __vbaStrCat
  loc_00E18540: mov edx, eax
  loc_00E18542: lea ecx, var_50
  loc_00E18545: call [00401228h] ; __vbaStrMove
  loc_00E1854B: lea edx, var_130
  loc_00E18551: lea ecx, var_70
  loc_00E18554: mov var_128, 00000000h
  loc_00E1855E: mov var_130, 00000002h
  loc_00E18568: call [0040101Ch] ; __vbaVarMove
  loc_00E1856E: mov eax, var_2C
  loc_00E18571: mov ecx, var_4C
  loc_00E18574: push eax
  loc_00E18575: push ecx
  loc_00E18576: call [0040106Ch] ; __vbaStrCat
  loc_00E1857C: mov edx, eax
  loc_00E1857E: lea ecx, var_90
  loc_00E18584: call [00401228h] ; __vbaStrMove
  loc_00E1858A: mov edx, var_50
  loc_00E1858D: push eax
  loc_00E1858E: push edx
  loc_00E1858F: call [0040106Ch] ; __vbaStrCat
  loc_00E18595: mov edx, eax
  loc_00E18597: lea ecx, var_2C
  loc_00E1859A: call [00401228h] ; __vbaStrMove
  loc_00E185A0: lea ecx, var_90
  loc_00E185A6: call [00401258h] ; __vbaFreeStr
  loc_00E185AC: jmp 00E17B33h
  loc_00E185B1: mov eax, var_2C
  loc_00E185B4: mov ecx, var_48
  loc_00E185B7: mov edi, [0040106Ch] ; __vbaStrCat
  loc_00E185BD: push eax
  loc_00E185BE: push ecx
  loc_00E185BF: call edi
  loc_00E185C1: mov esi, [00401228h] ; __vbaStrMove
  loc_00E185C7: mov edx, eax
  loc_00E185C9: lea ecx, var_2C
  loc_00E185CC: call __vbaStrMove
  loc_00E185CE: mov edx, var_2C
  loc_00E185D1: push edx
  loc_00E185D2: push 006E1124h ; "Rupiah "
  loc_00E185D7: call edi
  loc_00E185D9: mov edx, eax
  loc_00E185DB: lea ecx, var_78
  loc_00E185DE: call __vbaStrMove
  loc_00E185E0: lea ecx, var_130
  loc_00E185E6: lea edx, var_A0
  loc_00E185EC: lea eax, var_78
  loc_00E185EF: mov ebx, 00004008h
  loc_00E185F4: push ecx
  loc_00E185F5: push edx
  loc_00E185F6: mov var_128, eax
  loc_00E185FC: mov var_130, ebx
  loc_00E18602: call [00401054h] ; rtcLowerCaseVar
  loc_00E18608: mov edi, [00401034h] ; __vbaStrVarMove
  loc_00E1860E: lea eax, var_A0
  loc_00E18614: push eax
  loc_00E18615: call edi
  loc_00E18617: mov edx, eax
  loc_00E18619: lea ecx, var_78
  loc_00E1861C: call __vbaStrMove
  loc_00E1861E: lea ecx, var_A0
  loc_00E18624: call [00401024h] ; __vbaFreeVar
  loc_00E1862A: lea edx, var_130
  loc_00E18630: push 00000001h
  loc_00E18632: lea eax, var_A0
  loc_00E18638: lea ecx, var_78
  loc_00E1863B: push edx
  loc_00E1863C: push eax
  loc_00E1863D: mov var_128, ecx
  loc_00E18643: mov var_130, ebx
  loc_00E18649: call [0040121Ch] ; rtcLeftCharVar
  loc_00E1864F: lea ecx, var_A0
  loc_00E18655: lea edx, var_B0
  loc_00E1865B: push ecx
  loc_00E1865C: push edx
  loc_00E1865D: call [004010FCh] ; rtcUpperCaseVar
  loc_00E18663: lea eax, var_B0
  loc_00E18669: push eax
  loc_00E1866A: call edi
  loc_00E1866C: mov edx, eax
  loc_00E1866E: lea ecx, var_74
  loc_00E18671: call __vbaStrMove
  loc_00E18673: mov ebx, [00401038h] ; __vbaFreeVarList
  loc_00E18679: lea ecx, var_B0
  loc_00E1867F: lea edx, var_A0
  loc_00E18685: push ecx
  loc_00E18686: push edx
  loc_00E18687: push 00000002h
  loc_00E18689: call ebx
  loc_00E1868B: mov ecx, var_78
  loc_00E1868E: mov eax, var_74
  loc_00E18691: add esp, 0000000Ch
  loc_00E18694: mov var_148, eax
  loc_00E1869A: mov var_150, 00000008h
  loc_00E186A4: push ecx
  loc_00E186A5: call [0040102Ch] ; __vbaLenBstr
  loc_00E186AB: sub eax, 00000001h
  loc_00E186AE: lea edx, var_78
  loc_00E186B1: jo 00E18821h
  loc_00E186B7: mov var_98, eax
  loc_00E186BD: lea eax, var_A0
  loc_00E186C3: push eax
  loc_00E186C4: lea ecx, var_130
  loc_00E186CA: push 00000002h
  loc_00E186CC: mov var_A0, 00000003h
  loc_00E186D6: mov var_128, edx
  loc_00E186DC: mov var_130, 00004008h
  loc_00E186E6: push ecx
  loc_00E186E7: lea edx, var_B0
  loc_00E186ED: push edx
  loc_00E186EE: call [004010F0h] ; rtcMidCharVar
  loc_00E186F4: lea eax, var_150
  loc_00E186FA: lea ecx, var_B0
  loc_00E18700: push eax
  loc_00E18701: lea edx, var_C0
  loc_00E18707: push ecx
  loc_00E18708: push edx
  loc_00E18709: call [00401184h] ; __vbaVarCat
  loc_00E1870F: push eax
  loc_00E18710: call edi
  loc_00E18712: mov edx, eax
  loc_00E18714: lea ecx, var_78
  loc_00E18717: call __vbaStrMove
  loc_00E18719: lea eax, var_C0
  loc_00E1871F: lea ecx, var_B0
  loc_00E18725: push eax
  loc_00E18726: lea edx, var_A0
  loc_00E1872C: push ecx
  loc_00E1872D: push edx
  loc_00E1872E: push 00000003h
  loc_00E18730: call ebx
  loc_00E18732: add esp, 00000010h
  loc_00E18735: fwait
  loc_00E18736: push 00E187FAh
  loc_00E1873B: jmp 00E187A3h
  loc_00E1873D: test var_4, 04h
  loc_00E18741: jz 00E1874Ch
  loc_00E18743: lea ecx, var_78
  loc_00E18746: call [00401258h] ; __vbaFreeStr
  loc_00E1874C: lea ecx, var_90
  loc_00E18752: call [00401258h] ; __vbaFreeStr
  loc_00E18758: lea eax, var_120
  loc_00E1875E: lea ecx, var_110
  loc_00E18764: push eax
  loc_00E18765: lea edx, var_100
  loc_00E1876B: push ecx
  loc_00E1876C: lea eax, var_F0
  loc_00E18772: push edx
  loc_00E18773: lea ecx, var_E0
  loc_00E18779: push eax
  loc_00E1877A: lea edx, var_D0
  loc_00E18780: push ecx
  loc_00E18781: lea eax, var_C0
  loc_00E18787: push edx
  loc_00E18788: lea ecx, var_B0
  loc_00E1878E: push eax
  loc_00E1878F: lea edx, var_A0
  loc_00E18795: push ecx
  loc_00E18796: push edx
  loc_00E18797: push 00000009h
  loc_00E18799: call [00401038h] ; __vbaFreeVarList
  loc_00E1879F: add esp, 00000028h
  loc_00E187A2: ret
  loc_00E187A3: mov edi, [00401024h] ; __vbaFreeVar
  loc_00E187A9: lea ecx, var_184
  loc_00E187AF: call edi
  loc_00E187B1: lea ecx, var_28
  loc_00E187B4: call edi
  loc_00E187B6: mov esi, [00401258h] ; __vbaFreeStr
  loc_00E187BC: lea ecx, var_2C
  loc_00E187BF: call __vbaFreeStr
  loc_00E187C1: lea ecx, var_34
  loc_00E187C4: call __vbaFreeStr
  loc_00E187C6: lea ecx, var_44
  loc_00E187C9: call edi
  loc_00E187CB: lea ecx, var_48
  loc_00E187CE: call __vbaFreeStr
  loc_00E187D0: lea ecx, var_4C
  loc_00E187D3: call __vbaFreeStr
  loc_00E187D5: lea ecx, var_50
  loc_00E187D8: call __vbaFreeStr
  loc_00E187DA: lea ecx, var_60
  loc_00E187DD: call edi
  loc_00E187DF: lea ecx, var_70
  loc_00E187E2: call edi
  loc_00E187E4: lea ecx, var_74
  loc_00E187E7: call __vbaFreeStr
  loc_00E187E9: lea ecx, var_88
  loc_00E187EF: call edi
  loc_00E187F1: lea ecx, var_8C
  loc_00E187F7: call __vbaFreeStr
  loc_00E187F9: ret
  loc_00E187FA: mov eax, Me
  loc_00E187FD: push eax
  loc_00E187FE: mov ecx, [eax]
  loc_00E18800: call [ecx+00000008h]
  loc_00E18803: mov edx, arg_10
  loc_00E18806: mov eax, var_78
  loc_00E18809: mov [edx], eax
  loc_00E1880B: mov eax, var_4
  loc_00E1880E: mov ecx, var_14
  loc_00E18811: pop edi
  loc_00E18812: pop esi
  loc_00E18813: mov fs:[00000000h], ecx
  loc_00E1881A: pop ebx
  loc_00E1881B: mov esp, ebp
  loc_00E1881D: pop ebp
  loc_00E1881E: retn 000Ch
End Function

Private Sub Proc_6_17_E18C20
  loc_00E18C20: xor eax, eax
  loc_00E18C22: retn 0004h
End Sub

Private Sub Proc_6_18_E1A0A0
  loc_00E1A0A0: push ebp
  loc_00E1A0A1: mov ebp, esp
  loc_00E1A0A3: sub esp, 00000008h
  loc_00E1A0A6: push 00402836h ; __vbaExceptHandler
  loc_00E1A0AB: mov eax, fs:[00000000h]
  loc_00E1A0B1: push eax
  loc_00E1A0B2: mov fs:[00000000h], esp
  loc_00E1A0B9: sub esp, 00000010h
  loc_00E1A0BC: push ebx
  loc_00E1A0BD: push esi
  loc_00E1A0BE: push edi
  loc_00E1A0BF: mov var_8, esp
  loc_00E1A0C2: mov var_4, 00402390h
  loc_00E1A0C9: mov eax, Me
  loc_00E1A0CC: mov var_14, 00000000h
  loc_00E1A0D3: push eax
  loc_00E1A0D4: mov ecx, [eax]
  loc_00E1A0D6: call [ecx+00000310h]
  loc_00E1A0DC: lea edx, var_14
  loc_00E1A0DF: push eax
  loc_00E1A0E0: push edx
  loc_00E1A0E1: call [004010ACh] ; __vbaObjSet
  loc_00E1A0E7: mov esi, eax
  loc_00E1A0E9: push esi
  loc_00E1A0EA: mov eax, [esi]
  loc_00E1A0EC: call [eax+00000204h]
  loc_00E1A0F2: test eax, eax
  loc_00E1A0F4: fnclex
  loc_00E1A0F6: jge 00E1A10Ah
  loc_00E1A0F8: push 00000204h
  loc_00E1A0FD: push 006DCB70h
  loc_00E1A102: push esi
  loc_00E1A103: push eax
  loc_00E1A104: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1A10A: lea ecx, var_14
  loc_00E1A10D: call [00401254h] ; __vbaFreeObj
  loc_00E1A113: push 00E1A125h
  loc_00E1A118: jmp 00E1A124h
  loc_00E1A11A: lea ecx, var_14
  loc_00E1A11D: call [00401254h] ; __vbaFreeObj
  loc_00E1A123: ret
  loc_00E1A124: ret
  loc_00E1A125: mov ecx, var_10
  loc_00E1A128: pop edi
  loc_00E1A129: pop esi
  loc_00E1A12A: xor eax, eax
  loc_00E1A12C: mov fs:[00000000h], ecx
  loc_00E1A133: pop ebx
  loc_00E1A134: mov esp, ebp
  loc_00E1A136: pop ebp
  loc_00E1A137: retn 0004h
End Sub

Private Sub Proc_6_19_E1E360
  loc_00E1E360: push ebp
  loc_00E1E361: mov ebp, esp
  loc_00E1E363: sub esp, 00000008h
  loc_00E1E366: push 00402836h ; __vbaExceptHandler
  loc_00E1E36B: mov eax, fs:[00000000h]
  loc_00E1E371: push eax
  loc_00E1E372: mov fs:[00000000h], esp
  loc_00E1E379: sub esp, 00000050h
  loc_00E1E37C: push ebx
  loc_00E1E37D: push esi
  loc_00E1E37E: push edi
  loc_00E1E37F: mov var_8, esp
  loc_00E1E382: mov var_4, 00402400h
  loc_00E1E389: mov esi, Me
  loc_00E1E38C: xor eax, eax
  loc_00E1E38E: mov var_14, eax
  loc_00E1E391: mov var_18, eax
  loc_00E1E394: mov var_1C, eax
  loc_00E1E397: mov var_20, eax
  loc_00E1E39A: mov var_24, eax
  loc_00E1E39D: mov var_34, eax
  loc_00E1E3A0: mov eax, [esi]
  loc_00E1E3A2: push esi
  loc_00E1E3A3: call [eax+00000520h]
  loc_00E1E3A9: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E1E3AF: lea ecx, var_20
  loc_00E1E3B2: push eax
  loc_00E1E3B3: push ecx
  loc_00E1E3B4: call edi
  loc_00E1E3B6: mov ebx, eax
  loc_00E1E3B8: lea eax, var_14
  loc_00E1E3BB: push eax
  loc_00E1E3BC: push ebx
  loc_00E1E3BD: mov edx, [ebx]
  loc_00E1E3BF: call [edx+00000050h]
  loc_00E1E3C2: test eax, eax
  loc_00E1E3C4: fnclex
  loc_00E1E3C6: jge 00E1E3D7h
  loc_00E1E3C8: push 00000050h
  loc_00E1E3CA: push 006E0410h
  loc_00E1E3CF: push ebx
  loc_00E1E3D0: push eax
  loc_00E1E3D1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E3D7: mov ecx, var_14
  loc_00E1E3DA: push 006E11ECh ; "select * from biaya where no_daftar ='"
  loc_00E1E3DF: push ecx
  loc_00E1E3E0: lea ebx, [esi+00000034h]
  loc_00E1E3E3: call [0040106Ch] ; __vbaStrCat
  loc_00E1E3E9: mov edx, eax
  loc_00E1E3EB: lea ecx, var_18
  loc_00E1E3EE: call [00401228h] ; __vbaStrMove
  loc_00E1E3F4: push eax
  loc_00E1E3F5: push 006DCB84h ; "'"
  loc_00E1E3FA: call [0040106Ch] ; __vbaStrCat
  loc_00E1E400: mov edx, eax
  loc_00E1E402: lea ecx, var_1C
  loc_00E1E405: call [00401228h] ; __vbaStrMove
  loc_00E1E40B: mov edx, eax
  loc_00E1E40D: mov ecx, ebx
  loc_00E1E40F: call [004011B0h] ; __vbaStrCopy
  loc_00E1E415: lea edx, var_1C
  loc_00E1E418: lea eax, var_18
  loc_00E1E41B: push edx
  loc_00E1E41C: lea ecx, var_14
  loc_00E1E41F: push eax
  loc_00E1E420: push ecx
  loc_00E1E421: push 00000003h
  loc_00E1E423: call [004011BCh] ; __vbaFreeStrList
  loc_00E1E429: add esp, 00000010h
  loc_00E1E42C: lea ecx, var_20
  loc_00E1E42F: call [00401254h] ; __vbaFreeObj
  loc_00E1E435: sub esp, 00000010h
  loc_00E1E438: mov eax, 00004008h
  loc_00E1E43D: mov edx, esp
  loc_00E1E43F: mov ecx, [esi]
  loc_00E1E441: push 0000000Eh
  loc_00E1E443: push esi
  loc_00E1E444: mov [edx], eax
  loc_00E1E446: mov eax, var_40
  loc_00E1E449: mov [edx+00000004h], eax
  loc_00E1E44C: mov [edx+00000008h], ebx
  loc_00E1E44F: mov ebx, var_38
  loc_00E1E452: mov [edx+0000000Ch], ebx
  loc_00E1E455: call [ecx+00000564h]
  loc_00E1E45B: lea edx, var_20
  loc_00E1E45E: push eax
  loc_00E1E45F: push edx
  loc_00E1E460: call edi
  loc_00E1E462: push eax
  loc_00E1E463: call [00401238h] ; __vbaLateIdSt
  loc_00E1E469: lea ecx, var_20
  loc_00E1E46C: call [00401254h] ; __vbaFreeObj
  loc_00E1E472: mov eax, [esi]
  loc_00E1E474: push 006DCBD8h
  loc_00E1E479: push 00000000h
  loc_00E1E47B: push 00000018h
  loc_00E1E47D: push esi
  loc_00E1E47E: call [eax+00000564h]
  loc_00E1E484: lea ecx, var_20
  loc_00E1E487: push eax
  loc_00E1E488: push ecx
  loc_00E1E489: call edi
  loc_00E1E48B: lea edx, var_34
  loc_00E1E48E: push eax
  loc_00E1E48F: push edx
  loc_00E1E490: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1E496: add esp, 00000010h
  loc_00E1E499: push eax
  loc_00E1E49A: call [00401120h] ; __vbaCastObjVar
  loc_00E1E4A0: push eax
  loc_00E1E4A1: lea eax, var_24
  loc_00E1E4A4: push eax
  loc_00E1E4A5: call edi
  loc_00E1E4A7: sub esp, 00000010h
  loc_00E1E4AA: mov ecx, 0000000Ah
  loc_00E1E4AF: mov edi, esp
  loc_00E1E4B1: mov esi, eax
  loc_00E1E4B3: mov var_44, ecx
  loc_00E1E4B6: mov eax, 80020004h
  loc_00E1E4BB: mov edx, [esi]
  loc_00E1E4BD: mov [edi], ecx
  loc_00E1E4BF: mov ecx, var_50
  loc_00E1E4C2: mov var_3C, eax
  loc_00E1E4C5: mov [edi+00000004h], ecx
  loc_00E1E4C8: mov [edi+00000008h], eax
  loc_00E1E4CB: mov eax, var_48
  loc_00E1E4CE: sub esp, 00000010h
  loc_00E1E4D1: mov [edi+0000000Ch], eax
  loc_00E1E4D4: mov eax, var_44
  loc_00E1E4D7: mov ecx, esp
  loc_00E1E4D9: push esi
  loc_00E1E4DA: mov [ecx], eax
  loc_00E1E4DC: mov eax, var_40
  loc_00E1E4DF: mov [ecx+00000004h], eax
  loc_00E1E4E2: mov eax, var_3C
  loc_00E1E4E5: mov [ecx+00000008h], eax
  loc_00E1E4E8: mov [ecx+0000000Ch], ebx
  loc_00E1E4EB: call [edx+000000ACh]
  loc_00E1E4F1: test eax, eax
  loc_00E1E4F3: fnclex
  loc_00E1E4F5: jge 00E1E509h
  loc_00E1E4F7: push 000000ACh
  loc_00E1E4FC: push 006DCBE8h
  loc_00E1E501: push esi
  loc_00E1E502: push eax
  loc_00E1E503: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1E509: lea ecx, var_24
  loc_00E1E50C: lea edx, var_20
  loc_00E1E50F: push ecx
  loc_00E1E510: push edx
  loc_00E1E511: push 00000002h
  loc_00E1E513: call [00401048h] ; __vbaFreeObjList
  loc_00E1E519: add esp, 0000000Ch
  loc_00E1E51C: lea ecx, var_34
  loc_00E1E51F: call [00401024h] ; __vbaFreeVar
  loc_00E1E525: push 00E1E55Eh
  loc_00E1E52A: jmp 00E1E55Dh
  loc_00E1E52C: lea eax, var_1C
  loc_00E1E52F: lea ecx, var_18
  loc_00E1E532: push eax
  loc_00E1E533: lea edx, var_14
  loc_00E1E536: push ecx
  loc_00E1E537: push edx
  loc_00E1E538: push 00000003h
  loc_00E1E53A: call [004011BCh] ; __vbaFreeStrList
  loc_00E1E540: lea eax, var_24
  loc_00E1E543: lea ecx, var_20
  loc_00E1E546: push eax
  loc_00E1E547: push ecx
  loc_00E1E548: push 00000002h
  loc_00E1E54A: call [00401048h] ; __vbaFreeObjList
  loc_00E1E550: add esp, 0000001Ch
  loc_00E1E553: lea ecx, var_34
  loc_00E1E556: call [00401024h] ; __vbaFreeVar
  loc_00E1E55C: ret
  loc_00E1E55D: ret
  loc_00E1E55E: mov ecx, var_10
  loc_00E1E561: pop edi
  loc_00E1E562: pop esi
  loc_00E1E563: xor eax, eax
  loc_00E1E565: mov fs:[00000000h], ecx
  loc_00E1E56C: pop ebx
  loc_00E1E56D: mov esp, ebp
  loc_00E1E56F: pop ebp
  loc_00E1E570: retn 0004h
End Sub
