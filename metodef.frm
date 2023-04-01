VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form metodef
  BackColor = &H404040&
  ScaleMode = 1
  AutoRedraw = False
  FontTransparent = True
  BorderStyle = 0 'None
  'Icon = n/a
  LinkTopic = "Form1"
  ClientLeft = 0
  ClientTop = 0
  ClientWidth = 14640
  ClientHeight = 9780
  StartUpPosition = 2 'CenterScreen
  Begin VB.TextBox ubayar
    Left = 12510
    Top = 4920
    Width = 1815
    Height = 435
    Visible = 0   'False
    TabIndex = 60
  End
  Begin VB.TextBox tr
    Left = 12510
    Top = 4380
    Width = 1815
    Height = 435
    Visible = 0   'False
    TabIndex = 59
  End
  Begin VB.TextBox daftar
    BackColor = &HFFFFFF&
    Left = 4680
    Top = 3090
    Width = 1665
    Height = 330
    TabIndex = 48
    BorderStyle = 0 'None
    Alignment = 2 'Center
  End
  Begin VB.Frame fcari
    Caption = "Cari Disini"
    BackColor = &HE0E0E0&
    ForeColor = &H404040&
    Left = 3600
    Top = 4860
    Width = 8115
    Height = 1305
    TabIndex = 42
    BeginProperty Font
      Name = "Trebuchet MS"
      Size = 9.75
      Charset = 0
      Weight = 700
      Underline = 0 'False
      Italic = 0 'False
      Strikethrough = 0 'False
    EndProperty
    Begin VB.TextBox txtcari
      BackColor = &HC0C0C0&
      Left = 240
      Top = 840
      Width = 7665
      Height = 345
      TabIndex = 46
      BorderStyle = 0 'None
    End
    Begin VB.OptionButton optno
      Caption = "No. Daftar"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 2130
      Top = 360
      Width = 1455
      Height = 300
      TabIndex = 45
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
      Left = 210
      Top = 330
      Width = 1455
      Height = 300
      TabIndex = 44
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
    Begin Project1.jcbutton jcbutton2
      Left = 7890
      Top = 0
      Width = 225
      Height = 225
      TabIndex = 43
      OleObjectBlob = "metodef.frx":0000
    End
    Begin VB.Shape Shape17
      Left = 390
      Top = 900
      Width = 7635
      Height = 375
      BorderStyle = 0 'Transparent
      FillColor = &H404040&
      FillStyle = 0
    End
  End
  Begin VB.Frame flc
    Caption = "&H0080FF80&"
    BackColor = &H4000&
    Left = 5190
    Top = 4320
    Width = 4995
    Height = 2565
    TabIndex = 0
    BorderStyle = 0 'None
    Begin Project1.jcbutton metodes
      Left = 690
      Top = 450
      Width = 1335
      Height = 1335
      TabIndex = 1
      OleObjectBlob = "metodef.frx":0220
    End
    Begin Project1.jcbutton jcbutton1
      Left = 3000
      Top = 480
      Width = 1335
      Height = 1335
      TabIndex = 2
      OleObjectBlob = "metodef.frx":15F4
    End
    Begin VB.Shape Shape12
      Left = 4710
      Top = 0
      Width = 135
      Height = 2625
      BorderStyle = 0 'Transparent
      FillColor = &H808080&
      FillStyle = 0
    End
    Begin VB.Shape Shape7
      Left = 4830
      Top = 0
      Width = 195
      Height = 2625
      BorderStyle = 0 'Transparent
      FillColor = &HE0E0E0&
      FillStyle = 0
    End
    Begin VB.Shape Shape6
      Left = 150
      Top = 0
      Width = 105
      Height = 2625
      BorderStyle = 0 'Transparent
      FillColor = &H808080&
      FillStyle = 0
    End
    Begin VB.Shape Shape5
      Left = -30
      Top = 0
      Width = 195
      Height = 2625
      BorderStyle = 0 'Transparent
      FillColor = &HE0E0E0&
      FillStyle = 0
    End
    Begin VB.Label Label1
      Caption = "Lunas"
      ForeColor = &H80FF80&
      Left = 1050
      Top = 1860
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
    Begin VB.Label Label2
      Caption = "Mencicil"
      ForeColor = &HFFFF&
      Left = 3240
      Top = 1890
      Width = 1035
      Height = 315
      TabIndex = 3
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
    Begin VB.Shape Shape4
      Left = 2790
      Top = 300
      Width = 1755
      Height = 2055
      BorderStyle = 0 'Transparent
      FillColor = &H40&
      FillStyle = 0
    End
    Begin VB.Shape Shape3
      Left = 480
      Top = 330
      Width = 1755
      Height = 1995
      BorderStyle = 0 'Transparent
      FillColor = &H8000&
      FillStyle = 0
    End
  End
  Begin VB.Frame fhitung
    BackColor = &H404040&
    Left = 2790
    Top = 3630
    Width = 9615
    Height = 3885
    TabIndex = 21
    BorderStyle = 0 'None
    Begin VB.Frame fawal
      Caption = "Frame1"
      BackColor = &H404040&
      Left = 330
      Top = 750
      Width = 4185
      Height = 615
      TabIndex = 55
      BorderStyle = 0 'None
      Begin VB.Label cicilanawal
        Caption = "0"
        ForeColor = &HFFFFFF&
        Left = 2400
        Top = 0
        Width = 1575
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
      Begin VB.Label Label16
        Caption = "Rp"
        ForeColor = &HFFFFFF&
        Left = 1980
        Top = 0
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
      Begin VB.Line Line8
        BorderColor = &HE0E0E0&
        X1 = 2400
        Y1 = 330
        X2 = 4110
        Y2 = 330
      End
      Begin VB.Label Label17
        Caption = "Sisa Cicilan Sebelumnya"
        ForeColor = &H80FF80&
        Left = 0
        Top = -30
        Width = 2115
        Height = 615
        TabIndex = 56
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
    End
    Begin VB.Frame ftotal
      Caption = "Frame1"
      BackColor = &H404040&
      Left = 480
      Top = 690
      Width = 4065
      Height = 495
      TabIndex = 51
      BorderStyle = 0 'None
      Begin VB.Label total
        Caption = "0"
        ForeColor = &HFFFFFF&
        Left = 2130
        Top = 0
        Width = 1815
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
      Begin VB.Label Label38
        Caption = "Rp"
        ForeColor = &HFFFFFF&
        Left = 1710
        Top = 0
        Width = 405
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
      Begin VB.Line Line20
        BorderColor = &HE0E0E0&
        X1 = 2130
        Y1 = 330
        X2 = 3840
        Y2 = 330
      End
      Begin VB.Label Label37
        Caption = "Total Biaya"
        ForeColor = &H80FF80&
        Left = 0
        Top = 30
        Width = 1815
        Height = 315
        TabIndex = 52
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
    End
    Begin VB.TextBox bayar
      BackColor = &H404040&
      ForeColor = &H80FFFF&
      Left = 7200
      Top = 750
      Width = 1725
      Height = 330
      TabIndex = 34
      BorderStyle = 0 'None
    End
    Begin VB.Frame flunas
      Caption = "Frame1"
      BackColor = &H404040&
      Left = 600
      Top = 1920
      Width = 3945
      Height = 1695
      TabIndex = 30
      BorderStyle = 0 'None
      Begin VB.Label Label14
        Caption = "Rp"
        ForeColor = &HFFFFFF&
        Left = 1710
        Top = 30
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
      Begin VB.Image ilunas
        Picture = "metodef.frx":29C8
        Left = 2490
        Top = 390
        Width = 1395
        Height = 1275
        Stretch = -1  'True
      End
      Begin VB.Label kembali
        Caption = "-"
        ForeColor = &HFFFFFF&
        Left = 2040
        Top = 30
        Width = 1815
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
      Begin VB.Line Line10
        BorderColor = &HE0E0E0&
        X1 = 2010
        Y1 = 360
        X2 = 3720
        Y2 = 360
      End
      Begin VB.Label Label12
        Caption = "Uang Kembali"
        ForeColor = &H80FF80&
        Left = 120
        Top = 60
        Width = 1725
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
      Begin VB.Label lunas
        Caption = "LUNAS"
        ForeColor = &H8080FF&
        Left = 540
        Top = 780
        Width = 1125
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
    End
    Begin VB.Frame fcicil
      Caption = "Frame1"
      BackColor = &H404040&
      ForeColor = &H404040&
      Left = 5160
      Top = 2100
      Width = 3945
      Height = 1365
      TabIndex = 27
      BorderStyle = 0 'None
      Begin VB.TextBox cicil
        BackColor = &H404040&
        ForeColor = &H80FF80&
        Left = 1980
        Top = 0
        Width = 1725
        Height = 330
        Text = "0"
        TabIndex = 38
        BorderStyle = 0 'None
      End
      Begin Project1.jcbutton callx
        Left = 60
        Top = 960
        Width = 3675
        Height = 375
        TabIndex = 39
        OleObjectBlob = "metodef.frx":0001B62D
      End
      Begin VB.Label Label9
        Caption = "Sisa Cicilan"
        ForeColor = &H80FF80&
        Left = 60
        Top = 450
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
      Begin VB.Label Label8
        Caption = "Rp"
        ForeColor = &HFFFFFF&
        Left = 1590
        Top = 420
        Width = 405
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
      Begin VB.Line Line7
        BorderColor = &HE0E0E0&
        X1 = 1950
        Y1 = 780
        X2 = 3660
        Y2 = 780
      End
      Begin VB.Label sisa
        Caption = "0"
        ForeColor = &HFFFFFF&
        Left = 1980
        Top = 450
        Width = 1815
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
      Begin VB.Label Label10
        Caption = "Cicilan Ke"
        ForeColor = &H8080FF&
        Left = 60
        Top = 30
        Width = 1485
        Height = 315
        TabIndex = 28
        BackStyle = 0 'Transparent
        BeginProperty Font
          Name = "Trebuchet MS"
          Size = 14.25
          Charset = 0
          Weight = 700
          Underline = 0 'False
          Italic = 0 'False
          Strikethrough = 0 'False
        EndProperty
      End
      Begin VB.Line Line6
        BorderColor = &HE0E0E0&
        X1 = 1950
        Y1 = 330
        X2 = 3660
        Y2 = 330
      End
    End
    Begin VB.Frame fbayar
      Caption = "Frame1"
      BackColor = &H404040&
      Left = 5280
      Top = 780
      Width = 3885
      Height = 405
      TabIndex = 24
      BorderStyle = 0 'None
      Begin VB.Label Label7
        Caption = "Rp"
        ForeColor = &HFFFFFF&
        Left = 1530
        Top = 0
        Width = 405
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
      Begin VB.Line Line5
        BorderColor = &HE0E0E0&
        X1 = 1950
        Y1 = 330
        X2 = 3660
        Y2 = 330
      End
      Begin VB.Label Label6
        Caption = "Uang Bayar"
        ForeColor = &H80FF80&
        Left = 0
        Top = 0
        Width = 1485
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
    End
    Begin Project1.jcbutton totalkan
      Left = 3090
      Top = 1470
      Width = 3645
      Height = 375
      TabIndex = 22
      OleObjectBlob = "metodef.frx":0001BCB1
    End
    Begin VB.Image Image7
      Picture = "metodef.frx":0001BEDD
      Left = 0
      Top = 0
      Width = 435
      Height = 450
      Stretch = -1  'True
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
    Begin VB.Shape Shape26
      Left = 0
      Top = 420
      Width = 12255
      Height = 75
      BorderStyle = 0 'Transparent
      FillColor = &H80FF80&
      FillStyle = 0
    End
    Begin VB.Shape Shape27
      Left = 8910
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &H8080FF&
      FillStyle = 0
    End
    Begin VB.Shape Shape28
      Left = 9300
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &HFF&
      FillStyle = 0
    End
    Begin VB.Shape Shape29
      Left = 8490
      Top = 120
      Width = 225
      Height = 195
      Shape = 3
      BorderStyle = 0 'Transparent
      FillColor = &HC0C0FF&
      FillStyle = 0
    End
    Begin VB.Label Label15
      Caption = "Silahkan Input Data Yang Dperlukan"
      ForeColor = &H80FF80&
      Left = 540
      Top = 60
      Width = 4065
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
    Begin VB.Shape Shape32
      BorderColor = &H80FF80&
      Left = -60
      Top = 1380
      Width = 15015
      Height = 540
      BorderStyle = 3 'Dot
      FillColor = &HFFFF&
    End
    Begin VB.Shape Shape33
      BorderColor = &H80FF80&
      Left = 0
      Top = 3660
      Width = 15015
      Height = 60
      BorderStyle = 3 'Dot
      FillColor = &HFFFF&
    End
    Begin VB.Shape Shape30
      Left = 0
      Top = 0
      Width = 9735
      Height = 435
      BorderStyle = 0 'Transparent
      FillColor = &H4000&
      FillStyle = 0
    End
  End
  Begin VB.Timer Timer1
    Interval = 100
    Left = 1230
    Top = 6360
  End
  Begin VB.Frame Frame4
    Caption = "Frame4"
    BackColor = &H4000&
    Left = 8430
    Top = 1530
    Width = 6195
    Height = 1575
    TabIndex = 15
    BorderStyle = 0 'None
    Begin Project1.jcbutton update
      Left = 1800
      Top = 360
      Width = 1425
      Height = 885
      TabIndex = 50
      OleObjectBlob = "metodef.frx":0001D987
    End
    Begin Project1.jcbutton refreshh
      Left = 300
      Top = 360
      Width = 1425
      Height = 885
      TabIndex = 16
      OleObjectBlob = "metodef.frx":0001E8CF
    End
    Begin Project1.jcbutton simpan
      Left = 1800
      Top = 360
      Width = 1425
      Height = 885
      TabIndex = 17
      OleObjectBlob = "metodef.frx":0001F6FB
    End
    Begin Project1.jcbutton hapus
      Left = 3300
      Top = 360
      Width = 1365
      Height = 885
      TabIndex = 18
      OleObjectBlob = "metodef.frx":00020157
    End
    Begin Project1.jcbutton cari
      Left = 4740
      Top = 360
      Width = 1275
      Height = 885
      TabIndex = 19
      OleObjectBlob = "metodef.frx":00021033
    End
    Begin VB.Shape ssimpan
      BorderColor = &HFF&
      Left = 1740
      Top = 300
      Width = 1545
      Height = 1035
      BorderWidth = 2
    End
    Begin VB.Label Label5
      Caption = "Tombol Navigasi"
      ForeColor = &HFFFFFF&
      Left = 240
      Top = 30
      Width = 1965
      Height = 315
      TabIndex = 20
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
    Begin VB.Line Line4
      BorderColor = &HE0E0E0&
      X1 = 270
      Y1 = 1350
      X2 = 6030
      Y2 = 1350
    End
  End
  Begin VB.Frame Frame3
    Caption = "Frame3"
    BackColor = &H4000&
    Left = 0
    Top = 1530
    Width = 6885
    Height = 1545
    TabIndex = 9
    BorderStyle = 0 'None
    Begin VB.Label tipe
      Caption = "........"
      ForeColor = &H8080FF&
      Left = 4920
      Top = 1140
      Width = 1965
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
    Begin VB.Image Image1
      Picture = "metodef.frx":000220CF
      Left = 150
      Top = 270
      Width = 915
      Height = 855
      Stretch = -1  'True
    End
    Begin VB.Shape Shape2
      BorderColor = &HFFFFFF&
      Left = 180
      Top = 240
      Width = 795
      Height = 885
      BorderWidth = 2
    End
    Begin VB.Label nodafs
      Caption = "No. Daftar"
      ForeColor = &HFFFFFF&
      Left = 1200
      Top = 150
      Width = 2445
      Height = 315
      TabIndex = 14
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
      Left = 1200
      Top = 570
      Width = 2445
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
    Begin VB.Label jurus
      Caption = "Jurusan Yang Dipilih"
      ForeColor = &HFFFFFF&
      Left = 1200
      Top = 960
      Width = 2445
      Height = 315
      TabIndex = 12
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
    Begin VB.Line Line1
      BorderColor = &H80FF80&
      X1 = 1200
      Y1 = 480
      X2 = 3630
      Y2 = 480
    End
    Begin VB.Line Line2
      BorderColor = &H80FF80&
      X1 = 1200
      Y1 = 900
      X2 = 3630
      Y2 = 900
    End
    Begin VB.Line Line3
      BorderColor = &H80FF80&
      X1 = 1200
      Y1 = 1260
      X2 = 3630
      Y2 = 1260
    End
    Begin VB.Line Line9
      BorderColor = &H80FF80&
      X1 = 3990
      Y1 = 450
      X2 = 6420
      Y2 = 450
    End
    Begin VB.Label jenisk
      Caption = "Jenis Kelamin"
      ForeColor = &HFFFFFF&
      Left = 3990
      Top = 120
      Width = 2445
      Height = 315
      TabIndex = 11
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
    Begin VB.Label agama
      Caption = "Agama"
      ForeColor = &HFFFFFF&
      Left = 3990
      Top = 570
      Width = 2445
      Height = 315
      TabIndex = 10
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
    Begin VB.Line Line24
      BorderColor = &H80FF80&
      X1 = 3990
      Y1 = 900
      X2 = 6420
      Y2 = 900
    End
  End
  Begin MSAdodcLib.Adodc adopen
    Left = 480
    Top = 7980
    Width = 1200
    Height = 330
    Visible = 0   'False
    OleObjectBlob = "metodef.frx":00022C86
  End
  Begin MSDataGridLib.DataGrid dgpen
    Left = 150
    Top = 8040
    Width = 14385
    Height = 1575
    TabIndex = 41
    OleObjectBlob = "metodef.frx":00022FB8
  End
  Begin Project1.jcbutton tgl
    Left = 6330
    Top = 3060
    Width = 555
    Height = 375
    TabIndex = 49
    OleObjectBlob = "metodef.frx":00023163
  End
  Begin VB.Label Label11
    Caption = "Tanggal Transaksi"
    ForeColor = &H4000&
    Left = 1920
    Top = 3090
    Width = 1815
    Height = 315
    TabIndex = 47
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
  Begin VB.Shape Shape18
    Left = 0
    Top = 3060
    Width = 6885
    Height = 375
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Shape Shape16
    Left = -30
    Top = 7860
    Width = 14685
    Height = 465
    BorderStyle = 0 'Transparent
    FillColor = &H808080&
    FillStyle = 0
  End
  Begin VB.Shape Shape13
    Left = -30
    Top = 8310
    Width = 14685
    Height = 1485
    BorderStyle = 0 'Transparent
    FillColor = &HC0C0C0&
    FillStyle = 0
  End
  Begin VB.Image back
    Picture = "metodef.frx":0002336F
    Left = 60
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Label Label4
    Caption = "Metode Pembayaran"
    ForeColor = &H80FF80&
    Left = 6420
    Top = 60
    Width = 2745
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
  Begin VB.Shape Shape9
    Left = -30
    Top = 450
    Width = 14685
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Image Image4
    Picture = "metodef.frx":00024E19
    Left = 13590
    Top = 540
    Width = 1305
    Height = 735
    Stretch = -1  'True
  End
  Begin VB.Image Image3
    Picture = "metodef.frx":0002C2EF
    Left = 420
    Top = 570
    Width = 915
    Height = 855
    Stretch = -1  'True
  End
  Begin VB.Image Image2
    Picture = "metodef.frx":0005F5D1
    Left = 12840
    Top = 600
    Width = 705
    Height = 675
    Stretch = -1  'True
  End
  Begin VB.Label Label3
    Caption = "SMK Swasta RK Bintang Timur Pematang Siantar"
    ForeColor = &H4000&
    Left = 1500
    Top = 600
    Width = 3465
    Height = 675
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
  Begin VB.Label lbltgl
    Caption = "Labeltanggal"
    ForeColor = &H80FF80&
    Left = 12240
    Top = 60
    Width = 1395
    Height = 375
    TabIndex = 6
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
    Left = 13650
    Top = 60
    Width = 45
    Height = 285
    BorderStyle = 0 'Transparent
    FillColor = &HFF00&
    FillStyle = 0
  End
  Begin VB.Label lbljam
    Caption = "Labeljam"
    ForeColor = &H80FF80&
    Left = 13710
    Top = 60
    Width = 1005
    Height = 375
    TabIndex = 5
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
  Begin VB.Shape Shape1
    Left = -210
    Top = 1470
    Width = 7095
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
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
  Begin VB.Shape Shape11
    Left = 0
    Top = 450
    Width = 14715
    Height = 1035
    BorderStyle = 0 'Transparent
    FillColor = &HFFFFFF&
    FillStyle = 0
  End
  Begin VB.Shape Shape10
    Left = 30
    Top = 0
    Width = 14685
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Image Image6
    Picture = "metodef.frx":0006510D
    Left = 270
    Top = 330
    Width = 14745
    Height = 9525
    Stretch = -1  'True
  End
End

Attribute VB_Name = "metodef"


Private Sub totalkan_UnknownEvent_9 'E2CB20
  loc_00E2CB20: push ebp
  loc_00E2CB21: mov ebp, esp
  loc_00E2CB23: sub esp, 0000000Ch
  loc_00E2CB26: push 00402836h ; __vbaExceptHandler
  loc_00E2CB2B: mov eax, fs:[00000000h]
  loc_00E2CB31: push eax
  loc_00E2CB32: mov fs:[00000000h], esp
  loc_00E2CB39: sub esp, 00000120h
  loc_00E2CB3F: push ebx
  loc_00E2CB40: push esi
  loc_00E2CB41: push edi
  loc_00E2CB42: mov var_C, esp
  loc_00E2CB45: mov var_8, 00402570h
  loc_00E2CB4C: mov esi, Me
  loc_00E2CB4F: mov eax, esi
  loc_00E2CB51: and eax, 00000001h
  loc_00E2CB54: mov var_4, eax
  loc_00E2CB57: and esi, FFFFFFFEh
  loc_00E2CB5A: push esi
  loc_00E2CB5B: mov Me, esi
  loc_00E2CB5E: mov ecx, [esi]
  loc_00E2CB60: call [ecx+00000004h]
  loc_00E2CB63: mov edx, [esi]
  loc_00E2CB65: xor ebx, ebx
  loc_00E2CB67: push esi
  loc_00E2CB68: mov var_18, ebx
  loc_00E2CB6B: mov var_1C, ebx
  loc_00E2CB6E: mov var_20, ebx
  loc_00E2CB71: mov var_24, ebx
  loc_00E2CB74: mov var_28, ebx
  loc_00E2CB77: mov var_2C, ebx
  loc_00E2CB7A: mov var_30, ebx
  loc_00E2CB7D: mov var_34, ebx
  loc_00E2CB80: mov var_38, ebx
  loc_00E2CB83: mov var_3C, ebx
  loc_00E2CB86: mov var_40, ebx
  loc_00E2CB89: mov var_50, ebx
  loc_00E2CB8C: mov var_60, ebx
  loc_00E2CB8F: mov var_70, ebx
  loc_00E2CB92: mov var_80, ebx
  loc_00E2CB95: mov var_90, ebx
  loc_00E2CB9B: mov var_A0, ebx
  loc_00E2CBA1: mov var_C4, ebx
  loc_00E2CBA7: call [edx+0000037Ch]
  loc_00E2CBAD: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E2CBB3: push eax
  loc_00E2CBB4: lea eax, var_28
  loc_00E2CBB7: push eax
  loc_00E2CBB8: call edi
  loc_00E2CBBA: mov ecx, [eax]
  loc_00E2CBBC: lea edx, var_C4
  loc_00E2CBC2: push edx
  loc_00E2CBC3: push eax
  loc_00E2CBC4: mov var_D8, eax
  loc_00E2CBCA: call [ecx+00000098h]
  loc_00E2CBD0: cmp eax, ebx
  loc_00E2CBD2: fnclex
  loc_00E2CBD4: jge 00E2CBEEh
  loc_00E2CBD6: mov ecx, var_D8
  loc_00E2CBDC: push 00000098h
  loc_00E2CBE1: push 006DCAD0h
  loc_00E2CBE6: push ecx
  loc_00E2CBE7: push eax
  loc_00E2CBE8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CBEE: xor edx, edx
  loc_00E2CBF0: cmp var_C4, FFFFFFh
  loc_00E2CBF8: lea ecx, var_28
  loc_00E2CBFB: setz dl
  loc_00E2CBFE: neg edx
  loc_00E2CC00: mov var_E0, dx
  loc_00E2CC07: call [00401254h] ; __vbaFreeObj
  loc_00E2CC0D: cmp var_E0, bx
  loc_00E2CC14: jz 00E2D203h
  loc_00E2CC1A: mov eax, [esi]
  loc_00E2CC1C: push esi
  loc_00E2CC1D: call [eax+00000368h]
  loc_00E2CC23: lea ecx, var_2C
  loc_00E2CC26: push eax
  loc_00E2CC27: push ecx
  loc_00E2CC28: call edi
  loc_00E2CC2A: mov edx, [eax]
  loc_00E2CC2C: lea ecx, var_1C
  loc_00E2CC2F: push ecx
  loc_00E2CC30: push eax
  loc_00E2CC31: mov var_D8, eax
  loc_00E2CC37: call [edx+00000050h]
  loc_00E2CC3A: cmp eax, ebx
  loc_00E2CC3C: fnclex
  loc_00E2CC3E: jge 00E2CC55h
  loc_00E2CC40: mov edx, var_D8
  loc_00E2CC46: push 00000050h
  loc_00E2CC48: push 006E0410h
  loc_00E2CC4D: push edx
  loc_00E2CC4E: push eax
  loc_00E2CC4F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CC55: mov eax, var_1C
  loc_00E2CC58: push eax
  loc_00E2CC59: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2CC5F: mov ecx, [esi]
  loc_00E2CC61: push esi
  loc_00E2CC62: fstp real8 ptr var_CC
  loc_00E2CC68: call [ecx+00000388h]
  loc_00E2CC6E: lea edx, var_30
  loc_00E2CC71: push eax
  loc_00E2CC72: push edx
  loc_00E2CC73: call edi
  loc_00E2CC75: mov var_E8, eax
  loc_00E2CC7B: mov eax, [esi]
  loc_00E2CC7D: push esi
  loc_00E2CC7E: call [eax+00000378h]
  loc_00E2CC84: lea ecx, var_28
  loc_00E2CC87: push eax
  loc_00E2CC88: push ecx
  loc_00E2CC89: call edi
  loc_00E2CC8B: mov edx, [eax]
  loc_00E2CC8D: lea ecx, var_18
  loc_00E2CC90: push ecx
  loc_00E2CC91: push eax
  loc_00E2CC92: mov var_E0, eax
  loc_00E2CC98: call [edx+000000A0h]
  loc_00E2CC9E: cmp eax, ebx
  loc_00E2CCA0: fnclex
  loc_00E2CCA2: jge 00E2CCBCh
  loc_00E2CCA4: mov edx, var_E0
  loc_00E2CCAA: push 000000A0h
  loc_00E2CCAF: push 006DCB70h
  loc_00E2CCB4: push edx
  loc_00E2CCB5: push eax
  loc_00E2CCB6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CCBC: mov eax, var_E8
  loc_00E2CCC2: mov ecx, var_18
  loc_00E2CCC5: push ecx
  loc_00E2CCC6: mov ebx, [eax]
  loc_00E2CCC8: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2CCCE: fsub st0, real8 ptr var_CC
  loc_00E2CCD4: sub esp, 00000008h
  loc_00E2CCD7: fnstsw ax
  loc_00E2CCD9: test al, 0Dh
  loc_00E2CCDB: jnz 00E2DCFAh
  loc_00E2CCE1: fstp real8 ptr [esp]
  loc_00E2CCE4: call [00401134h] ; __vbaStrR8
  loc_00E2CCEA: mov edx, eax
  loc_00E2CCEC: lea ecx, var_20
  loc_00E2CCEF: call [00401228h] ; __vbaStrMove
  loc_00E2CCF5: mov edx, ebx
  loc_00E2CCF7: mov ebx, var_E8
  loc_00E2CCFD: push eax
  loc_00E2CCFE: push ebx
  loc_00E2CCFF: call [edx+00000054h]
  loc_00E2CD02: test eax, eax
  loc_00E2CD04: fnclex
  loc_00E2CD06: jge 00E2CD17h
  loc_00E2CD08: push 00000054h
  loc_00E2CD0A: push 006E0410h
  loc_00E2CD0F: push ebx
  loc_00E2CD10: push eax
  loc_00E2CD11: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CD17: lea eax, var_20
  loc_00E2CD1A: lea ecx, var_1C
  loc_00E2CD1D: push eax
  loc_00E2CD1E: lea edx, var_18
  loc_00E2CD21: push ecx
  loc_00E2CD22: push edx
  loc_00E2CD23: push 00000003h
  loc_00E2CD25: call [004011BCh] ; __vbaFreeStrList
  loc_00E2CD2B: lea eax, var_30
  loc_00E2CD2E: lea ecx, var_2C
  loc_00E2CD31: push eax
  loc_00E2CD32: lea edx, var_28
  loc_00E2CD35: push ecx
  loc_00E2CD36: push edx
  loc_00E2CD37: push 00000003h
  loc_00E2CD39: call [00401048h] ; __vbaFreeObjList
  loc_00E2CD3F: mov eax, [esi]
  loc_00E2CD41: add esp, 00000020h
  loc_00E2CD44: push esi
  loc_00E2CD45: call [eax+000002FCh]
  loc_00E2CD4B: lea ecx, var_2C
  loc_00E2CD4E: push eax
  loc_00E2CD4F: push ecx
  loc_00E2CD50: call edi
  loc_00E2CD52: mov edx, [esi]
  loc_00E2CD54: push esi
  loc_00E2CD55: mov var_E0, eax
  loc_00E2CD5B: call [edx+00000378h]
  loc_00E2CD61: push eax
  loc_00E2CD62: lea eax, var_28
  loc_00E2CD65: push eax
  loc_00E2CD66: call edi
  loc_00E2CD68: mov ebx, eax
  loc_00E2CD6A: lea edx, var_18
  loc_00E2CD6D: push edx
  loc_00E2CD6E: push ebx
  loc_00E2CD6F: mov ecx, [ebx]
  loc_00E2CD71: call [ecx+000000A0h]
  loc_00E2CD77: test eax, eax
  loc_00E2CD79: fnclex
  loc_00E2CD7B: jge 00E2CD8Fh
  loc_00E2CD7D: push 000000A0h
  loc_00E2CD82: push 006DCB70h
  loc_00E2CD87: push ebx
  loc_00E2CD88: push eax
  loc_00E2CD89: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CD8F: mov eax, var_E0
  loc_00E2CD95: mov ecx, var_18
  loc_00E2CD98: push ecx
  loc_00E2CD99: mov ebx, [eax]
  loc_00E2CD9B: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2CDA1: sub esp, 00000008h
  loc_00E2CDA4: fstp real8 ptr [esp]
  loc_00E2CDA7: call [00401134h] ; __vbaStrR8
  loc_00E2CDAD: mov edx, eax
  loc_00E2CDAF: lea ecx, var_1C
  loc_00E2CDB2: call [00401228h] ; __vbaStrMove
  loc_00E2CDB8: mov edx, ebx
  loc_00E2CDBA: mov ebx, var_E0
  loc_00E2CDC0: push eax
  loc_00E2CDC1: push ebx
  loc_00E2CDC2: call [edx+000000A4h]
  loc_00E2CDC8: test eax, eax
  loc_00E2CDCA: fnclex
  loc_00E2CDCC: jge 00E2CDE0h
  loc_00E2CDCE: push 000000A4h
  loc_00E2CDD3: push 006DCB70h
  loc_00E2CDD8: push ebx
  loc_00E2CDD9: push eax
  loc_00E2CDDA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CDE0: lea eax, var_1C
  loc_00E2CDE3: lea ecx, var_18
  loc_00E2CDE6: push eax
  loc_00E2CDE7: push ecx
  loc_00E2CDE8: push 00000002h
  loc_00E2CDEA: call [004011BCh] ; __vbaFreeStrList
  loc_00E2CDF0: lea edx, var_2C
  loc_00E2CDF3: lea eax, var_28
  loc_00E2CDF6: push edx
  loc_00E2CDF7: push eax
  loc_00E2CDF8: push 00000002h
  loc_00E2CDFA: call [00401048h] ; __vbaFreeObjList
  loc_00E2CE00: mov ecx, [esi]
  loc_00E2CE02: add esp, 00000018h
  loc_00E2CE05: push esi
  loc_00E2CE06: call [ecx+00000388h]
  loc_00E2CE0C: lea edx, var_28
  loc_00E2CE0F: push eax
  loc_00E2CE10: push edx
  loc_00E2CE11: call edi
  loc_00E2CE13: mov ebx, eax
  loc_00E2CE15: lea ecx, var_18
  loc_00E2CE18: push ecx
  loc_00E2CE19: push ebx
  loc_00E2CE1A: mov eax, [ebx]
  loc_00E2CE1C: call [eax+00000050h]
  loc_00E2CE1F: test eax, eax
  loc_00E2CE21: fnclex
  loc_00E2CE23: jge 00E2CE34h
  loc_00E2CE25: push 00000050h
  loc_00E2CE27: push 006E0410h
  loc_00E2CE2C: push ebx
  loc_00E2CE2D: push eax
  loc_00E2CE2E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CE34: mov edx, var_18
  loc_00E2CE37: push edx
  loc_00E2CE38: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2CE3E: call [004010D8h] ; __vbaFpR8
  loc_00E2CE44: fcomp real8 ptr [004022E8h]
  loc_00E2CE4A: fnstsw ax
  loc_00E2CE4C: test ah, 40h
  loc_00E2CE4F: jz 00E2CE58h
  loc_00E2CE51: mov ebx, 00000001h
  loc_00E2CE56: jmp 00E2CE5Ah
  loc_00E2CE58: xor ebx, ebx
  loc_00E2CE5A: lea ecx, var_18
  loc_00E2CE5D: call [00401258h] ; __vbaFreeStr
  loc_00E2CE63: lea ecx, var_28
  loc_00E2CE66: call [00401254h] ; __vbaFreeObj
  loc_00E2CE6C: mov eax, [esi]
  loc_00E2CE6E: push esi
  loc_00E2CE6F: neg ebx
  loc_00E2CE71: test bx, bx
  loc_00E2CE74: jz 00E2CEE7h
  loc_00E2CE76: call [eax+00000384h]
  loc_00E2CE7C: lea ecx, var_28
  loc_00E2CE7F: push eax
  loc_00E2CE80: push ecx
  loc_00E2CE81: call edi
  loc_00E2CE83: mov ebx, eax
  loc_00E2CE85: push FFFFFFFFh
  loc_00E2CE87: push ebx
  loc_00E2CE88: mov edx, [ebx]
  loc_00E2CE8A: call [edx+0000008Ch]
  loc_00E2CE90: test eax, eax
  loc_00E2CE92: fnclex
  loc_00E2CE94: jge 00E2CEA8h
  loc_00E2CE96: push 0000008Ch
  loc_00E2CE9B: push 006DCAC0h
  loc_00E2CEA0: push ebx
  loc_00E2CEA1: push eax
  loc_00E2CEA2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CEA8: lea ecx, var_28
  loc_00E2CEAB: call [00401254h] ; __vbaFreeObj
  loc_00E2CEB1: mov eax, [esi]
  loc_00E2CEB3: push esi
  loc_00E2CEB4: call [eax+00000394h]
  loc_00E2CEBA: lea ecx, var_28
  loc_00E2CEBD: push eax
  loc_00E2CEBE: push ecx
  loc_00E2CEBF: call edi
  loc_00E2CEC1: mov ebx, eax
  loc_00E2CEC3: push FFFFFFFFh
  loc_00E2CEC5: push ebx
  loc_00E2CEC6: mov edx, [ebx]
  loc_00E2CEC8: call [edx+0000009Ch]
  loc_00E2CECE: test eax, eax
  loc_00E2CED0: fnclex
  loc_00E2CED2: jge 00E2DB9Dh
  loc_00E2CED8: push 0000009Ch
  loc_00E2CEDD: push 006E0410h
  loc_00E2CEE2: jmp 00E2DB95h
  loc_00E2CEE7: call [eax+00000388h]
  loc_00E2CEED: lea ecx, var_28
  loc_00E2CEF0: push eax
  loc_00E2CEF1: push ecx
  loc_00E2CEF2: call edi
  loc_00E2CEF4: mov ebx, eax
  loc_00E2CEF6: lea eax, var_18
  loc_00E2CEF9: push eax
  loc_00E2CEFA: push ebx
  loc_00E2CEFB: mov edx, [ebx]
  loc_00E2CEFD: call [edx+00000050h]
  loc_00E2CF00: test eax, eax
  loc_00E2CF02: fnclex
  loc_00E2CF04: jge 00E2CF15h
  loc_00E2CF06: push 00000050h
  loc_00E2CF08: push 006E0410h
  loc_00E2CF0D: push ebx
  loc_00E2CF0E: push eax
  loc_00E2CF0F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CF15: mov ecx, var_18
  loc_00E2CF18: push ecx
  loc_00E2CF19: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2CF1F: call [004010D8h] ; __vbaFpR8
  loc_00E2CF25: fcomp real8 ptr [004022E8h]
  loc_00E2CF2B: fnstsw ax
  loc_00E2CF2D: test ah, 41h
  loc_00E2CF30: jnz 00E2CF39h
  loc_00E2CF32: mov ebx, 00000001h
  loc_00E2CF37: jmp 00E2CF3Bh
  loc_00E2CF39: xor ebx, ebx
  loc_00E2CF3B: lea ecx, var_18
  loc_00E2CF3E: call [00401258h] ; __vbaFreeStr
  loc_00E2CF44: lea ecx, var_28
  loc_00E2CF47: call [00401254h] ; __vbaFreeObj
  loc_00E2CF4D: mov edx, [esi]
  loc_00E2CF4F: push esi
  loc_00E2CF50: neg ebx
  loc_00E2CF52: test bx, bx
  loc_00E2CF55: jz 00E2CFC8h
  loc_00E2CF57: call [edx+00000384h]
  loc_00E2CF5D: push eax
  loc_00E2CF5E: lea eax, var_28
  loc_00E2CF61: push eax
  loc_00E2CF62: call edi
  loc_00E2CF64: mov ebx, eax
  loc_00E2CF66: push FFFFFFFFh
  loc_00E2CF68: push ebx
  loc_00E2CF69: mov ecx, [ebx]
  loc_00E2CF6B: call [ecx+0000008Ch]
  loc_00E2CF71: test eax, eax
  loc_00E2CF73: fnclex
  loc_00E2CF75: jge 00E2CF89h
  loc_00E2CF77: push 0000008Ch
  loc_00E2CF7C: push 006DCAC0h
  loc_00E2CF81: push ebx
  loc_00E2CF82: push eax
  loc_00E2CF83: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CF89: lea ecx, var_28
  loc_00E2CF8C: call [00401254h] ; __vbaFreeObj
  loc_00E2CF92: mov edx, [esi]
  loc_00E2CF94: push esi
  loc_00E2CF95: call [edx+00000394h]
  loc_00E2CF9B: push eax
  loc_00E2CF9C: lea eax, var_28
  loc_00E2CF9F: push eax
  loc_00E2CFA0: call edi
  loc_00E2CFA2: mov ebx, eax
  loc_00E2CFA4: push FFFFFFFFh
  loc_00E2CFA6: push ebx
  loc_00E2CFA7: mov ecx, [ebx]
  loc_00E2CFA9: call [ecx+0000009Ch]
  loc_00E2CFAF: test eax, eax
  loc_00E2CFB1: fnclex
  loc_00E2CFB3: jge 00E2DB9Dh
  loc_00E2CFB9: push 0000009Ch
  loc_00E2CFBE: push 006E0410h
  loc_00E2CFC3: jmp 00E2DB95h
  loc_00E2CFC8: call [edx+00000368h]
  loc_00E2CFCE: push eax
  loc_00E2CFCF: lea eax, var_2C
  loc_00E2CFD2: push eax
  loc_00E2CFD3: call edi
  loc_00E2CFD5: mov ebx, eax
  loc_00E2CFD7: lea edx, var_1C
  loc_00E2CFDA: push edx
  loc_00E2CFDB: push ebx
  loc_00E2CFDC: mov ecx, [ebx]
  loc_00E2CFDE: call [ecx+00000050h]
  loc_00E2CFE1: test eax, eax
  loc_00E2CFE3: fnclex
  loc_00E2CFE5: jge 00E2CFF6h
  loc_00E2CFE7: push 00000050h
  loc_00E2CFE9: push 006E0410h
  loc_00E2CFEE: push ebx
  loc_00E2CFEF: push eax
  loc_00E2CFF0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CFF6: mov eax, var_1C
  loc_00E2CFF9: push eax
  loc_00E2CFFA: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D000: mov ecx, [esi]
  loc_00E2D002: push esi
  loc_00E2D003: fstp real8 ptr var_CC
  loc_00E2D009: call [ecx+00000378h]
  loc_00E2D00F: lea edx, var_28
  loc_00E2D012: push eax
  loc_00E2D013: push edx
  loc_00E2D014: call edi
  loc_00E2D016: mov ebx, eax
  loc_00E2D018: lea ecx, var_18
  loc_00E2D01B: push ecx
  loc_00E2D01C: push ebx
  loc_00E2D01D: mov eax, [ebx]
  loc_00E2D01F: call [eax+000000A0h]
  loc_00E2D025: test eax, eax
  loc_00E2D027: fnclex
  loc_00E2D029: jge 00E2D03Dh
  loc_00E2D02B: push 000000A0h
  loc_00E2D030: push 006DCB70h
  loc_00E2D035: push ebx
  loc_00E2D036: push eax
  loc_00E2D037: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D03D: mov edx, var_18
  loc_00E2D040: push edx
  loc_00E2D041: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D047: call [004010D8h] ; __vbaFpR8
  loc_00E2D04D: fstp real8 ptr var_120
  loc_00E2D053: fld real8 ptr var_CC
  loc_00E2D059: call [004010D8h] ; __vbaFpR8
  loc_00E2D05F: fcomp real8 ptr var_120
  loc_00E2D065: fnstsw ax
  loc_00E2D067: test ah, 41h
  loc_00E2D06A: jnz 00E2D073h
  loc_00E2D06C: mov ebx, 00000001h
  loc_00E2D071: jmp 00E2D075h
  loc_00E2D073: xor ebx, ebx
  loc_00E2D075: lea eax, var_1C
  loc_00E2D078: lea ecx, var_18
  loc_00E2D07B: push eax
  loc_00E2D07C: push ecx
  loc_00E2D07D: push 00000002h
  loc_00E2D07F: call [004011BCh] ; __vbaFreeStrList
  loc_00E2D085: lea edx, var_2C
  loc_00E2D088: lea eax, var_28
  loc_00E2D08B: push edx
  loc_00E2D08C: push eax
  loc_00E2D08D: push 00000002h
  loc_00E2D08F: call [00401048h] ; __vbaFreeObjList
  loc_00E2D095: neg ebx
  loc_00E2D097: add esp, 00000018h
  loc_00E2D09A: test bx, bx
  loc_00E2D09D: jz 00E2DBA6h
  loc_00E2D0A3: mov ebx, [004011F0h] ; __vbaVarDup
  loc_00E2D0A9: mov ecx, 0000000Ah
  loc_00E2D0AE: mov eax, 80020004h
  loc_00E2D0B3: mov var_80, ecx
  loc_00E2D0B6: mov var_70, ecx
  loc_00E2D0B9: lea edx, var_A0
  loc_00E2D0BF: lea ecx, var_60
  loc_00E2D0C2: mov var_78, eax
  loc_00E2D0C5: mov var_68, eax
  loc_00E2D0C8: mov var_98, 006E18E8h ; "Exhauted !"
  loc_00E2D0D2: mov var_A0, 00000008h
  loc_00E2D0DC: call ebx
  loc_00E2D0DE: lea edx, var_90
  loc_00E2D0E4: lea ecx, var_50
  loc_00E2D0E7: mov var_88, 006E189Ch ; "Tidak dapat melunasi Total biaya !"
  loc_00E2D0F1: mov var_90, 00000008h
  loc_00E2D0FB: call ebx
  loc_00E2D0FD: lea ecx, var_80
  loc_00E2D100: lea edx, var_70
  loc_00E2D103: push ecx
  loc_00E2D104: lea eax, var_60
  loc_00E2D107: push edx
  loc_00E2D108: push eax
  loc_00E2D109: lea ecx, var_50
  loc_00E2D10C: push 00000010h
  loc_00E2D10E: push ecx
  loc_00E2D10F: call [004010A8h] ; rtcMsgBox
  loc_00E2D115: lea edx, var_80
  loc_00E2D118: lea eax, var_70
  loc_00E2D11B: push edx
  loc_00E2D11C: lea ecx, var_60
  loc_00E2D11F: push eax
  loc_00E2D120: lea edx, var_50
  loc_00E2D123: push ecx
  loc_00E2D124: push edx
  loc_00E2D125: push 00000004h
  loc_00E2D127: call [00401038h] ; __vbaFreeVarList
  loc_00E2D12D: mov eax, [esi]
  loc_00E2D12F: add esp, 00000014h
  loc_00E2D132: push esi
  loc_00E2D133: call [eax+00000378h]
  loc_00E2D139: lea ecx, var_28
  loc_00E2D13C: push eax
  loc_00E2D13D: push ecx
  loc_00E2D13E: call edi
  loc_00E2D140: mov ebx, eax
  loc_00E2D142: push 006DCC80h
  loc_00E2D147: push ebx
  loc_00E2D148: mov edx, [ebx]
  loc_00E2D14A: call [edx+000000A4h]
  loc_00E2D150: test eax, eax
  loc_00E2D152: fnclex
  loc_00E2D154: jge 00E2D168h
  loc_00E2D156: push 000000A4h
  loc_00E2D15B: push 006DCB70h
  loc_00E2D160: push ebx
  loc_00E2D161: push eax
  loc_00E2D162: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D168: lea ecx, var_28
  loc_00E2D16B: call [00401254h] ; __vbaFreeObj
  loc_00E2D171: mov ebx, [004011F0h] ; __vbaVarDup
  loc_00E2D177: mov ecx, 0000000Ah
  loc_00E2D17C: mov eax, 80020004h
  loc_00E2D181: mov var_80, ecx
  loc_00E2D184: mov var_70, ecx
  loc_00E2D187: lea edx, var_A0
  loc_00E2D18D: lea ecx, var_60
  loc_00E2D190: mov var_78, eax
  loc_00E2D193: mov var_68, eax
  loc_00E2D196: mov var_98, 006E1948h ; "INFO"
  loc_00E2D1A0: mov var_A0, 00000008h
  loc_00E2D1AA: call ebx
  loc_00E2D1AC: lea edx, var_90
  loc_00E2D1B2: lea ecx, var_50
  loc_00E2D1B5: mov var_88, 006E1904h ; " Tentukan Kembali Uang bayar !"
  loc_00E2D1BF: mov var_90, 00000008h
  loc_00E2D1C9: call ebx
  loc_00E2D1CB: lea eax, var_80
  loc_00E2D1CE: lea ecx, var_70
  loc_00E2D1D1: push eax
  loc_00E2D1D2: lea edx, var_60
  loc_00E2D1D5: push ecx
  loc_00E2D1D6: push edx
  loc_00E2D1D7: lea eax, var_50
  loc_00E2D1DA: push 00000040h
  loc_00E2D1DC: push eax
  loc_00E2D1DD: call [004010A8h] ; rtcMsgBox
  loc_00E2D1E3: lea ecx, var_80
  loc_00E2D1E6: lea edx, var_70
  loc_00E2D1E9: push ecx
  loc_00E2D1EA: lea eax, var_60
  loc_00E2D1ED: push edx
  loc_00E2D1EE: lea ecx, var_50
  loc_00E2D1F1: push eax
  loc_00E2D1F2: push ecx
  loc_00E2D1F3: push 00000004h
  loc_00E2D1F5: call [00401038h] ; __vbaFreeVarList
  loc_00E2D1FB: add esp, 00000014h
  loc_00E2D1FE: jmp 00E2DBA6h
  loc_00E2D203: mov edx, [esi]
  loc_00E2D205: push esi
  loc_00E2D206: call [edx+00000398h]
  loc_00E2D20C: push eax
  loc_00E2D20D: lea eax, var_28
  loc_00E2D210: push eax
  loc_00E2D211: call edi
  loc_00E2D213: mov ecx, [eax]
  loc_00E2D215: lea edx, var_C4
  loc_00E2D21B: push edx
  loc_00E2D21C: push eax
  loc_00E2D21D: mov var_D8, eax
  loc_00E2D223: call [ecx+00000098h]
  loc_00E2D229: cmp eax, ebx
  loc_00E2D22B: fnclex
  loc_00E2D22D: jge 00E2D247h
  loc_00E2D22F: mov ecx, var_D8
  loc_00E2D235: push 00000098h
  loc_00E2D23A: push 006DCAD0h
  loc_00E2D23F: push ecx
  loc_00E2D240: push eax
  loc_00E2D241: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D247: xor edx, edx
  loc_00E2D249: cmp var_C4, FFFFFFh
  loc_00E2D251: lea ecx, var_28
  loc_00E2D254: setz dl
  loc_00E2D257: neg edx
  loc_00E2D259: mov var_E0, dx
  loc_00E2D260: call [00401254h] ; __vbaFreeObj
  loc_00E2D266: cmp var_E0, bx
  loc_00E2D26D: jz 00E2DBA6h
  loc_00E2D273: mov eax, [esi]
  loc_00E2D275: push esi
  loc_00E2D276: call [eax+000002FCh]
  loc_00E2D27C: lea ecx, var_2C
  loc_00E2D27F: push eax
  loc_00E2D280: push ecx
  loc_00E2D281: call edi
  loc_00E2D283: mov edx, [esi]
  loc_00E2D285: push esi
  loc_00E2D286: mov var_E0, eax
  loc_00E2D28C: call [edx+00000378h]
  loc_00E2D292: push eax
  loc_00E2D293: lea eax, var_28
  loc_00E2D296: push eax
  loc_00E2D297: call edi
  loc_00E2D299: mov ecx, [eax]
  loc_00E2D29B: lea edx, var_18
  loc_00E2D29E: push edx
  loc_00E2D29F: push eax
  loc_00E2D2A0: mov var_D8, eax
  loc_00E2D2A6: call [ecx+000000A0h]
  loc_00E2D2AC: cmp eax, ebx
  loc_00E2D2AE: fnclex
  loc_00E2D2B0: jge 00E2D2CAh
  loc_00E2D2B2: mov ecx, var_D8
  loc_00E2D2B8: push 000000A0h
  loc_00E2D2BD: push 006DCB70h
  loc_00E2D2C2: push ecx
  loc_00E2D2C3: push eax
  loc_00E2D2C4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D2CA: mov edx, var_E0
  loc_00E2D2D0: mov eax, var_18
  loc_00E2D2D3: push eax
  loc_00E2D2D4: mov ebx, [edx]
  loc_00E2D2D6: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D2DC: sub esp, 00000008h
  loc_00E2D2DF: fstp real8 ptr [esp]
  loc_00E2D2E2: call [00401134h] ; __vbaStrR8
  loc_00E2D2E8: mov edx, eax
  loc_00E2D2EA: lea ecx, var_1C
  loc_00E2D2ED: call [00401228h] ; __vbaStrMove
  loc_00E2D2F3: mov ecx, ebx
  loc_00E2D2F5: mov ebx, var_E0
  loc_00E2D2FB: push eax
  loc_00E2D2FC: push ebx
  loc_00E2D2FD: call [ecx+000000A4h]
  loc_00E2D303: test eax, eax
  loc_00E2D305: fnclex
  loc_00E2D307: jge 00E2D31Bh
  loc_00E2D309: push 000000A4h
  loc_00E2D30E: push 006DCB70h
  loc_00E2D313: push ebx
  loc_00E2D314: push eax
  loc_00E2D315: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D31B: lea edx, var_1C
  loc_00E2D31E: lea eax, var_18
  loc_00E2D321: push edx
  loc_00E2D322: push eax
  loc_00E2D323: push 00000002h
  loc_00E2D325: call [004011BCh] ; __vbaFreeStrList
  loc_00E2D32B: lea ecx, var_2C
  loc_00E2D32E: lea edx, var_28
  loc_00E2D331: push ecx
  loc_00E2D332: push edx
  loc_00E2D333: push 00000002h
  loc_00E2D335: call [00401048h] ; __vbaFreeObjList
  loc_00E2D33B: mov eax, [esi]
  loc_00E2D33D: add esp, 00000018h
  loc_00E2D340: push 00000000h
  loc_00E2D342: push 80010007h
  loc_00E2D347: push esi
  loc_00E2D348: call [eax+000003A0h]
  loc_00E2D34E: lea ecx, var_28
  loc_00E2D351: push eax
  loc_00E2D352: push ecx
  loc_00E2D353: call edi
  loc_00E2D355: lea edx, var_50
  loc_00E2D358: push eax
  loc_00E2D359: push edx
  loc_00E2D35A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2D360: add esp, 00000010h
  loc_00E2D363: push eax
  loc_00E2D364: call [004010CCh] ; __vbaBoolVar
  loc_00E2D36A: xor ebx, ebx
  loc_00E2D36C: cmp ax, FFFFFFh
  loc_00E2D370: setz bl
  loc_00E2D373: lea ecx, var_28
  loc_00E2D376: neg ebx
  loc_00E2D378: call [00401254h] ; __vbaFreeObj
  loc_00E2D37E: lea ecx, var_50
  loc_00E2D381: call [00401024h] ; __vbaFreeVar
  loc_00E2D387: test bx, bx
  loc_00E2D38A: jz 00E2D7F7h
  loc_00E2D390: mov eax, [esi]
  loc_00E2D392: push 006DCBD8h
  loc_00E2D397: push 00000000h
  loc_00E2D399: push 00000018h
  loc_00E2D39B: push esi
  loc_00E2D39C: call [eax+000004A8h]
  loc_00E2D3A2: lea ecx, var_2C
  loc_00E2D3A5: push eax
  loc_00E2D3A6: push ecx
  loc_00E2D3A7: call edi
  loc_00E2D3A9: lea edx, var_50
  loc_00E2D3AC: push eax
  loc_00E2D3AD: push edx
  loc_00E2D3AE: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2D3B4: add esp, 00000010h
  loc_00E2D3B7: push eax
  loc_00E2D3B8: call [00401120h] ; __vbaCastObjVar
  loc_00E2D3BE: push eax
  loc_00E2D3BF: lea eax, var_30
  loc_00E2D3C2: push eax
  loc_00E2D3C3: call edi
  loc_00E2D3C5: mov ebx, eax
  loc_00E2D3C7: lea edx, var_34
  loc_00E2D3CA: push edx
  loc_00E2D3CB: push ebx
  loc_00E2D3CC: mov ecx, [ebx]
  loc_00E2D3CE: call [ecx+00000054h]
  loc_00E2D3D1: test eax, eax
  loc_00E2D3D3: fnclex
  loc_00E2D3D5: jge 00E2D3E6h
  loc_00E2D3D7: push 00000054h
  loc_00E2D3D9: push 006DCBE8h
  loc_00E2D3DE: push ebx
  loc_00E2D3DF: push eax
  loc_00E2D3E0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D3E6: lea ebx, var_38
  loc_00E2D3E9: mov eax, var_34
  loc_00E2D3EC: push ebx
  loc_00E2D3ED: mov ecx, 00000002h
  loc_00E2D3F2: sub esp, 00000010h
  loc_00E2D3F5: mov var_90, ecx
  loc_00E2D3FB: mov ebx, esp
  loc_00E2D3FD: mov var_88, 00000006h
  loc_00E2D407: mov edx, [eax]
  loc_00E2D409: push eax
  loc_00E2D40A: mov [ebx], ecx
  loc_00E2D40C: mov ecx, var_8C
  loc_00E2D412: mov var_E0, eax
  loc_00E2D418: mov [ebx+00000004h], ecx
  loc_00E2D41B: mov ecx, var_88
  loc_00E2D421: mov [ebx+00000008h], ecx
  loc_00E2D424: mov ecx, var_84
  loc_00E2D42A: mov [ebx+0000000Ch], ecx
  loc_00E2D42D: call [edx+00000028h]
  loc_00E2D430: test eax, eax
  loc_00E2D432: fnclex
  loc_00E2D434: jge 00E2D44Bh
  loc_00E2D436: mov edx, var_E0
  loc_00E2D43C: push 00000028h
  loc_00E2D43E: push 006E09E8h
  loc_00E2D443: push edx
  loc_00E2D444: push eax
  loc_00E2D445: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D44B: mov eax, var_38
  loc_00E2D44E: lea edx, var_60
  loc_00E2D451: push edx
  loc_00E2D452: push eax
  loc_00E2D453: mov ecx, [eax]
  loc_00E2D455: mov ebx, eax
  loc_00E2D457: call [ecx+00000034h]
  loc_00E2D45A: test eax, eax
  loc_00E2D45C: fnclex
  loc_00E2D45E: jge 00E2D46Fh
  loc_00E2D460: push 00000034h
  loc_00E2D462: push 006E09F8h
  loc_00E2D467: push ebx
  loc_00E2D468: push eax
  loc_00E2D469: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D46F: lea eax, var_60
  loc_00E2D472: push eax
  loc_00E2D473: call [00401034h] ; __vbaStrVarMove
  loc_00E2D479: mov edx, eax
  loc_00E2D47B: lea ecx, var_1C
  loc_00E2D47E: call [00401228h] ; __vbaStrMove
  loc_00E2D484: push eax
  loc_00E2D485: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D48B: mov ecx, [esi]
  loc_00E2D48D: push esi
  loc_00E2D48E: fstp real8 ptr var_CC
  loc_00E2D494: call [ecx+000003B0h]
  loc_00E2D49A: lea edx, var_3C
  loc_00E2D49D: push eax
  loc_00E2D49E: push edx
  loc_00E2D49F: call edi
  loc_00E2D4A1: mov ebx, eax
  loc_00E2D4A3: lea ecx, var_20
  loc_00E2D4A6: push ecx
  loc_00E2D4A7: push ebx
  loc_00E2D4A8: mov eax, [ebx]
  loc_00E2D4AA: call [eax+00000050h]
  loc_00E2D4AD: test eax, eax
  loc_00E2D4AF: fnclex
  loc_00E2D4B1: jge 00E2D4C2h
  loc_00E2D4B3: push 00000050h
  loc_00E2D4B5: push 006E0410h
  loc_00E2D4BA: push ebx
  loc_00E2D4BB: push eax
  loc_00E2D4BC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D4C2: mov edx, var_20
  loc_00E2D4C5: push edx
  loc_00E2D4C6: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D4CC: mov eax, [esi]
  loc_00E2D4CE: push esi
  loc_00E2D4CF: fstp real8 ptr var_D4
  loc_00E2D4D5: call [eax+000002FCh]
  loc_00E2D4DB: lea ecx, var_40
  loc_00E2D4DE: push eax
  loc_00E2D4DF: push ecx
  loc_00E2D4E0: call edi
  loc_00E2D4E2: mov edx, [esi]
  loc_00E2D4E4: push esi
  loc_00E2D4E5: mov var_100, eax
  loc_00E2D4EB: call [edx+00000378h]
  loc_00E2D4F1: push eax
  loc_00E2D4F2: lea eax, var_28
  loc_00E2D4F5: push eax
  loc_00E2D4F6: call edi
  loc_00E2D4F8: mov ebx, eax
  loc_00E2D4FA: lea edx, var_18
  loc_00E2D4FD: push edx
  loc_00E2D4FE: push ebx
  loc_00E2D4FF: mov ecx, [ebx]
  loc_00E2D501: call [ecx+000000A0h]
  loc_00E2D507: test eax, eax
  loc_00E2D509: fnclex
  loc_00E2D50B: jge 00E2D51Fh
  loc_00E2D50D: push 000000A0h
  loc_00E2D512: push 006DCB70h
  loc_00E2D517: push ebx
  loc_00E2D518: push eax
  loc_00E2D519: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D51F: mov eax, var_100
  loc_00E2D525: mov ecx, var_18
  loc_00E2D528: push ecx
  loc_00E2D529: mov ebx, [eax]
  loc_00E2D52B: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D531: fadd st0, real8 ptr var_CC
  loc_00E2D537: sub esp, 00000008h
  loc_00E2D53A: fadd st0, real8 ptr var_D4
  loc_00E2D540: fnstsw ax
  loc_00E2D542: test al, 0Dh
  loc_00E2D544: jnz 00E2DCFAh
  loc_00E2D54A: fstp real8 ptr [esp]
  loc_00E2D54D: call [00401134h] ; __vbaStrR8
  loc_00E2D553: mov edx, eax
  loc_00E2D555: lea ecx, var_24
  loc_00E2D558: call [00401228h] ; __vbaStrMove
  loc_00E2D55E: mov edx, ebx
  loc_00E2D560: mov ebx, var_100
  loc_00E2D566: push eax
  loc_00E2D567: push ebx
  loc_00E2D568: call [edx+000000A4h]
  loc_00E2D56E: test eax, eax
  loc_00E2D570: fnclex
  loc_00E2D572: jge 00E2D586h
  loc_00E2D574: push 000000A4h
  loc_00E2D579: push 006DCB70h
  loc_00E2D57E: push ebx
  loc_00E2D57F: push eax
  loc_00E2D580: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D586: lea eax, var_24
  loc_00E2D589: lea ecx, var_20
  loc_00E2D58C: push eax
  loc_00E2D58D: lea edx, var_1C
  loc_00E2D590: push ecx
  loc_00E2D591: lea eax, var_18
  loc_00E2D594: push edx
  loc_00E2D595: push eax
  loc_00E2D596: push 00000004h
  loc_00E2D598: call [004011BCh] ; __vbaFreeStrList
  loc_00E2D59E: lea ecx, var_40
  loc_00E2D5A1: lea edx, var_3C
  loc_00E2D5A4: push ecx
  loc_00E2D5A5: lea eax, var_38
  loc_00E2D5A8: push edx
  loc_00E2D5A9: lea ecx, var_34
  loc_00E2D5AC: push eax
  loc_00E2D5AD: lea edx, var_30
  loc_00E2D5B0: push ecx
  loc_00E2D5B1: lea eax, var_2C
  loc_00E2D5B4: push edx
  loc_00E2D5B5: lea ecx, var_28
  loc_00E2D5B8: push eax
  loc_00E2D5B9: push ecx
  loc_00E2D5BA: push 00000007h
  loc_00E2D5BC: call [00401048h] ; __vbaFreeObjList
  loc_00E2D5C2: lea edx, var_60
  loc_00E2D5C5: lea eax, var_50
  loc_00E2D5C8: push edx
  loc_00E2D5C9: push eax
  loc_00E2D5CA: push 00000002h
  loc_00E2D5CC: call [00401038h] ; __vbaFreeVarList
  loc_00E2D5D2: mov ecx, [esi]
  loc_00E2D5D4: add esp, 00000040h
  loc_00E2D5D7: push esi
  loc_00E2D5D8: call [ecx+00000378h]
  loc_00E2D5DE: lea edx, var_2C
  loc_00E2D5E1: push eax
  loc_00E2D5E2: push edx
  loc_00E2D5E3: call edi
  loc_00E2D5E5: mov ebx, eax
  loc_00E2D5E7: lea ecx, var_1C
  loc_00E2D5EA: push ecx
  loc_00E2D5EB: push ebx
  loc_00E2D5EC: mov eax, [ebx]
  loc_00E2D5EE: call [eax+000000A0h]
  loc_00E2D5F4: test eax, eax
  loc_00E2D5F6: fnclex
  loc_00E2D5F8: jge 00E2D60Ch
  loc_00E2D5FA: push 000000A0h
  loc_00E2D5FF: push 006DCB70h
  loc_00E2D604: push ebx
  loc_00E2D605: push eax
  loc_00E2D606: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D60C: mov edx, var_1C
  loc_00E2D60F: push edx
  loc_00E2D610: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D616: mov eax, [esi]
  loc_00E2D618: push esi
  loc_00E2D619: fstp real8 ptr var_CC
  loc_00E2D61F: call [eax+000003B0h]
  loc_00E2D625: lea ecx, var_30
  loc_00E2D628: push eax
  loc_00E2D629: push ecx
  loc_00E2D62A: call edi
  loc_00E2D62C: mov edx, [esi]
  loc_00E2D62E: push esi
  loc_00E2D62F: mov var_E8, eax
  loc_00E2D635: call [edx+00000354h]
  loc_00E2D63B: push eax
  loc_00E2D63C: lea eax, var_28
  loc_00E2D63F: push eax
  loc_00E2D640: call edi
  loc_00E2D642: mov ebx, eax
  loc_00E2D644: lea edx, var_18
  loc_00E2D647: push edx
  loc_00E2D648: push ebx
  loc_00E2D649: mov ecx, [ebx]
  loc_00E2D64B: call [ecx+00000050h]
  loc_00E2D64E: test eax, eax
  loc_00E2D650: fnclex
  loc_00E2D652: jge 00E2D663h
  loc_00E2D654: push 00000050h
  loc_00E2D656: push 006E0410h
  loc_00E2D65B: push ebx
  loc_00E2D65C: push eax
  loc_00E2D65D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D663: mov eax, var_E8
  loc_00E2D669: mov ecx, var_18
  loc_00E2D66C: push ecx
  loc_00E2D66D: mov ebx, [eax]
  loc_00E2D66F: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D675: fsub st0, real8 ptr var_CC
  loc_00E2D67B: sub esp, 00000008h
  loc_00E2D67E: fnstsw ax
  loc_00E2D680: test al, 0Dh
  loc_00E2D682: jnz 00E2DCFAh
  loc_00E2D688: fstp real8 ptr [esp]
  loc_00E2D68B: call [00401134h] ; __vbaStrR8
  loc_00E2D691: mov edx, eax
  loc_00E2D693: lea ecx, var_20
  loc_00E2D696: call [00401228h] ; __vbaStrMove
  loc_00E2D69C: mov edx, ebx
  loc_00E2D69E: mov ebx, var_E8
  loc_00E2D6A4: push eax
  loc_00E2D6A5: push ebx
  loc_00E2D6A6: call [edx+00000054h]
  loc_00E2D6A9: test eax, eax
  loc_00E2D6AB: fnclex
  loc_00E2D6AD: jge 00E2D6BEh
  loc_00E2D6AF: push 00000054h
  loc_00E2D6B1: push 006E0410h
  loc_00E2D6B6: push ebx
  loc_00E2D6B7: push eax
  loc_00E2D6B8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D6BE: lea eax, var_20
  loc_00E2D6C1: lea ecx, var_1C
  loc_00E2D6C4: push eax
  loc_00E2D6C5: lea edx, var_18
  loc_00E2D6C8: push ecx
  loc_00E2D6C9: push edx
  loc_00E2D6CA: push 00000003h
  loc_00E2D6CC: call [004011BCh] ; __vbaFreeStrList
  loc_00E2D6D2: lea eax, var_30
  loc_00E2D6D5: lea ecx, var_2C
  loc_00E2D6D8: push eax
  loc_00E2D6D9: lea edx, var_28
  loc_00E2D6DC: push ecx
  loc_00E2D6DD: push edx
  loc_00E2D6DE: push 00000003h
  loc_00E2D6E0: call [00401048h] ; __vbaFreeObjList
  loc_00E2D6E6: mov eax, [esi]
  loc_00E2D6E8: add esp, 00000020h
  loc_00E2D6EB: push esi
  loc_00E2D6EC: call [eax+00000368h]
  loc_00E2D6F2: lea ecx, var_2C
  loc_00E2D6F5: push eax
  loc_00E2D6F6: push ecx
  loc_00E2D6F7: call edi
  loc_00E2D6F9: mov ebx, eax
  loc_00E2D6FB: lea eax, var_1C
  loc_00E2D6FE: push eax
  loc_00E2D6FF: push ebx
  loc_00E2D700: mov edx, [ebx]
  loc_00E2D702: call [edx+00000050h]
  loc_00E2D705: test eax, eax
  loc_00E2D707: fnclex
  loc_00E2D709: jge 00E2D71Ah
  loc_00E2D70B: push 00000050h
  loc_00E2D70D: push 006E0410h
  loc_00E2D712: push ebx
  loc_00E2D713: push eax
  loc_00E2D714: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D71A: mov ecx, var_1C
  loc_00E2D71D: push ecx
  loc_00E2D71E: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D724: mov edx, [esi]
  loc_00E2D726: push esi
  loc_00E2D727: fstp real8 ptr var_CC
  loc_00E2D72D: call [edx+00000388h]
  loc_00E2D733: push eax
  loc_00E2D734: lea eax, var_30
  loc_00E2D737: push eax
  loc_00E2D738: call edi
  loc_00E2D73A: mov ecx, [esi]
  loc_00E2D73C: push esi
  loc_00E2D73D: mov var_E8, eax
  loc_00E2D743: call [ecx+00000378h]
  loc_00E2D749: lea edx, var_28
  loc_00E2D74C: push eax
  loc_00E2D74D: push edx
  loc_00E2D74E: call edi
  loc_00E2D750: mov ebx, eax
  loc_00E2D752: lea ecx, var_18
  loc_00E2D755: push ecx
  loc_00E2D756: push ebx
  loc_00E2D757: mov eax, [ebx]
  loc_00E2D759: call [eax+000000A0h]
  loc_00E2D75F: test eax, eax
  loc_00E2D761: fnclex
  loc_00E2D763: jge 00E2D777h
  loc_00E2D765: push 000000A0h
  loc_00E2D76A: push 006DCB70h
  loc_00E2D76F: push ebx
  loc_00E2D770: push eax
  loc_00E2D771: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D777: mov edx, var_E8
  loc_00E2D77D: mov eax, var_18
  loc_00E2D780: push eax
  loc_00E2D781: mov ebx, [edx]
  loc_00E2D783: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D789: fsub st0, real8 ptr var_CC
  loc_00E2D78F: sub esp, 00000008h
  loc_00E2D792: fnstsw ax
  loc_00E2D794: test al, 0Dh
  loc_00E2D796: jnz 00E2DCFAh
  loc_00E2D79C: fstp real8 ptr [esp]
  loc_00E2D79F: call [00401134h] ; __vbaStrR8
  loc_00E2D7A5: mov edx, eax
  loc_00E2D7A7: lea ecx, var_20
  loc_00E2D7AA: call [00401228h] ; __vbaStrMove
  loc_00E2D7B0: mov ecx, ebx
  loc_00E2D7B2: mov ebx, var_E8
  loc_00E2D7B8: push eax
  loc_00E2D7B9: push ebx
  loc_00E2D7BA: call [ecx+00000054h]
  loc_00E2D7BD: test eax, eax
  loc_00E2D7BF: fnclex
  loc_00E2D7C1: jge 00E2D7D2h
  loc_00E2D7C3: push 00000054h
  loc_00E2D7C5: push 006E0410h
  loc_00E2D7CA: push ebx
  loc_00E2D7CB: push eax
  loc_00E2D7CC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D7D2: lea edx, var_20
  loc_00E2D7D5: lea eax, var_1C
  loc_00E2D7D8: push edx
  loc_00E2D7D9: lea ecx, var_18
  loc_00E2D7DC: push eax
  loc_00E2D7DD: push ecx
  loc_00E2D7DE: push 00000003h
  loc_00E2D7E0: call [004011BCh] ; __vbaFreeStrList
  loc_00E2D7E6: lea edx, var_30
  loc_00E2D7E9: lea eax, var_2C
  loc_00E2D7EC: push edx
  loc_00E2D7ED: lea ecx, var_28
  loc_00E2D7F0: push eax
  loc_00E2D7F1: push ecx
  loc_00E2D7F2: jmp 00E2D90Ah
  loc_00E2D7F7: mov edx, [esi]
  loc_00E2D7F9: push esi
  loc_00E2D7FA: call [edx+00000378h]
  loc_00E2D800: push eax
  loc_00E2D801: lea eax, var_2C
  loc_00E2D804: push eax
  loc_00E2D805: call edi
  loc_00E2D807: mov ebx, eax
  loc_00E2D809: lea edx, var_1C
  loc_00E2D80C: push edx
  loc_00E2D80D: push ebx
  loc_00E2D80E: mov ecx, [ebx]
  loc_00E2D810: call [ecx+000000A0h]
  loc_00E2D816: test eax, eax
  loc_00E2D818: fnclex
  loc_00E2D81A: jge 00E2D82Eh
  loc_00E2D81C: push 000000A0h
  loc_00E2D821: push 006DCB70h
  loc_00E2D826: push ebx
  loc_00E2D827: push eax
  loc_00E2D828: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D82E: mov eax, var_1C
  loc_00E2D831: push eax
  loc_00E2D832: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D838: mov ecx, [esi]
  loc_00E2D83A: push esi
  loc_00E2D83B: fstp real8 ptr var_CC
  loc_00E2D841: call [ecx+000003B0h]
  loc_00E2D847: lea edx, var_30
  loc_00E2D84A: push eax
  loc_00E2D84B: push edx
  loc_00E2D84C: call edi
  loc_00E2D84E: mov var_E8, eax
  loc_00E2D854: mov eax, [esi]
  loc_00E2D856: push esi
  loc_00E2D857: call [eax+00000368h]
  loc_00E2D85D: lea ecx, var_28
  loc_00E2D860: push eax
  loc_00E2D861: push ecx
  loc_00E2D862: call edi
  loc_00E2D864: mov ebx, eax
  loc_00E2D866: lea eax, var_18
  loc_00E2D869: push eax
  loc_00E2D86A: push ebx
  loc_00E2D86B: mov edx, [ebx]
  loc_00E2D86D: call [edx+00000050h]
  loc_00E2D870: test eax, eax
  loc_00E2D872: fnclex
  loc_00E2D874: jge 00E2D885h
  loc_00E2D876: push 00000050h
  loc_00E2D878: push 006E0410h
  loc_00E2D87D: push ebx
  loc_00E2D87E: push eax
  loc_00E2D87F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D885: mov ecx, var_E8
  loc_00E2D88B: mov edx, var_18
  loc_00E2D88E: push edx
  loc_00E2D88F: mov ebx, [ecx]
  loc_00E2D891: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D897: fsub st0, real8 ptr var_CC
  loc_00E2D89D: sub esp, 00000008h
  loc_00E2D8A0: fnstsw ax
  loc_00E2D8A2: test al, 0Dh
  loc_00E2D8A4: jnz 00E2DCFAh
  loc_00E2D8AA: fstp real8 ptr [esp]
  loc_00E2D8AD: call [00401134h] ; __vbaStrR8
  loc_00E2D8B3: mov edx, eax
  loc_00E2D8B5: lea ecx, var_20
  loc_00E2D8B8: call [00401228h] ; __vbaStrMove
  loc_00E2D8BE: mov var_134, ebx
  loc_00E2D8C4: mov ebx, var_E8
  loc_00E2D8CA: push eax
  loc_00E2D8CB: mov eax, var_134
  loc_00E2D8D1: push ebx
  loc_00E2D8D2: call [eax+00000054h]
  loc_00E2D8D5: test eax, eax
  loc_00E2D8D7: fnclex
  loc_00E2D8D9: jge 00E2D8EAh
  loc_00E2D8DB: push 00000054h
  loc_00E2D8DD: push 006E0410h
  loc_00E2D8E2: push ebx
  loc_00E2D8E3: push eax
  loc_00E2D8E4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D8EA: lea ecx, var_20
  loc_00E2D8ED: lea edx, var_1C
  loc_00E2D8F0: push ecx
  loc_00E2D8F1: lea eax, var_18
  loc_00E2D8F4: push edx
  loc_00E2D8F5: push eax
  loc_00E2D8F6: push 00000003h
  loc_00E2D8F8: call [004011BCh] ; __vbaFreeStrList
  loc_00E2D8FE: lea ecx, var_30
  loc_00E2D901: lea edx, var_2C
  loc_00E2D904: push ecx
  loc_00E2D905: lea eax, var_28
  loc_00E2D908: push edx
  loc_00E2D909: push eax
  loc_00E2D90A: push 00000003h
  loc_00E2D90C: call [00401048h] ; __vbaFreeObjList
  loc_00E2D912: mov ecx, [esi]
  loc_00E2D914: add esp, 00000020h
  loc_00E2D917: push esi
  loc_00E2D918: call [ecx+000003B0h]
  loc_00E2D91E: lea edx, var_28
  loc_00E2D921: push eax
  loc_00E2D922: push edx
  loc_00E2D923: call edi
  loc_00E2D925: mov ebx, eax
  loc_00E2D927: lea ecx, var_18
  loc_00E2D92A: push ecx
  loc_00E2D92B: push ebx
  loc_00E2D92C: mov eax, [ebx]
  loc_00E2D92E: call [eax+00000050h]
  loc_00E2D931: test eax, eax
  loc_00E2D933: fnclex
  loc_00E2D935: jge 00E2D946h
  loc_00E2D937: push 00000050h
  loc_00E2D939: push 006E0410h
  loc_00E2D93E: push ebx
  loc_00E2D93F: push eax
  loc_00E2D940: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D946: mov edx, var_18
  loc_00E2D949: push edx
  loc_00E2D94A: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2D950: call [004010D8h] ; __vbaFpR8
  loc_00E2D956: fcomp real8 ptr [004022E8h]
  loc_00E2D95C: fnstsw ax
  loc_00E2D95E: test ah, 40h
  loc_00E2D961: jz 00E2D96Ah
  loc_00E2D963: mov ebx, 00000001h
  loc_00E2D968: jmp 00E2D96Ch
  loc_00E2D96A: xor ebx, ebx
  loc_00E2D96C: lea ecx, var_18
  loc_00E2D96F: call [00401258h] ; __vbaFreeStr
  loc_00E2D975: lea ecx, var_28
  loc_00E2D978: call [00401254h] ; __vbaFreeObj
  loc_00E2D97E: mov eax, [esi]
  loc_00E2D980: push esi
  loc_00E2D981: neg ebx
  loc_00E2D983: test bx, bx
  loc_00E2D986: jz 00E2D9ECh
  loc_00E2D988: call [eax+00000424h]
  loc_00E2D98E: lea ecx, var_28
  loc_00E2D991: push eax
  loc_00E2D992: push ecx
  loc_00E2D993: call edi
  loc_00E2D995: mov ebx, eax
  loc_00E2D997: push 006E1728h ; "Lunas"
  loc_00E2D99C: push ebx
  loc_00E2D99D: mov edx, [ebx]
  loc_00E2D99F: call [edx+00000054h]
  loc_00E2D9A2: test eax, eax
  loc_00E2D9A4: fnclex
  loc_00E2D9A6: jge 00E2D9B7h
  loc_00E2D9A8: push 00000054h
  loc_00E2D9AA: push 006E0410h
  loc_00E2D9AF: push ebx
  loc_00E2D9B0: push eax
  loc_00E2D9B1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2D9B7: lea ecx, var_28
  loc_00E2D9BA: call [00401254h] ; __vbaFreeObj
  loc_00E2D9C0: mov eax, [esi]
  loc_00E2D9C2: push esi
  loc_00E2D9C3: call [eax+0000037Ch]
  loc_00E2D9C9: lea ecx, var_28
  loc_00E2D9CC: push eax
  loc_00E2D9CD: push ecx
  loc_00E2D9CE: call edi
  loc_00E2D9D0: mov ebx, eax
  loc_00E2D9D2: push 00000000h
  loc_00E2D9D4: push ebx
  loc_00E2D9D5: mov edx, [ebx]
  loc_00E2D9D7: call [edx+0000009Ch]
  loc_00E2D9DD: test eax, eax
  loc_00E2D9DF: fnclex
  loc_00E2D9E1: jge 00E2DB9Dh
  loc_00E2D9E7: jmp 00E2DB8Bh
  loc_00E2D9EC: call [eax+000003B0h]
  loc_00E2D9F2: lea ecx, var_28
  loc_00E2D9F5: push eax
  loc_00E2D9F6: push ecx
  loc_00E2D9F7: call edi
  loc_00E2D9F9: mov ebx, eax
  loc_00E2D9FB: lea eax, var_18
  loc_00E2D9FE: push eax
  loc_00E2D9FF: push ebx
  loc_00E2DA00: mov edx, [ebx]
  loc_00E2DA02: call [edx+00000050h]
  loc_00E2DA05: test eax, eax
  loc_00E2DA07: fnclex
  loc_00E2DA09: jge 00E2DA1Ah
  loc_00E2DA0B: push 00000050h
  loc_00E2DA0D: push 006E0410h
  loc_00E2DA12: push ebx
  loc_00E2DA13: push eax
  loc_00E2DA14: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DA1A: mov ecx, var_18
  loc_00E2DA1D: push ecx
  loc_00E2DA1E: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2DA24: call [004010D8h] ; __vbaFpR8
  loc_00E2DA2A: fcomp real8 ptr [004022E8h]
  loc_00E2DA30: fnstsw ax
  loc_00E2DA32: test ah, 01h
  loc_00E2DA35: jz 00E2DA3Eh
  loc_00E2DA37: mov ebx, 00000001h
  loc_00E2DA3C: jmp 00E2DA40h
  loc_00E2DA3E: xor ebx, ebx
  loc_00E2DA40: lea ecx, var_18
  loc_00E2DA43: call [00401258h] ; __vbaFreeStr
  loc_00E2DA49: lea ecx, var_28
  loc_00E2DA4C: call [00401254h] ; __vbaFreeObj
  loc_00E2DA52: mov edx, [esi]
  loc_00E2DA54: push esi
  loc_00E2DA55: neg ebx
  loc_00E2DA57: test bx, bx
  loc_00E2DA5A: jz 00E2DAC0h
  loc_00E2DA5C: call [edx+00000424h]
  loc_00E2DA62: push eax
  loc_00E2DA63: lea eax, var_28
  loc_00E2DA66: push eax
  loc_00E2DA67: call edi
  loc_00E2DA69: mov ebx, eax
  loc_00E2DA6B: push 006E1728h ; "Lunas"
  loc_00E2DA70: push ebx
  loc_00E2DA71: mov ecx, [ebx]
  loc_00E2DA73: call [ecx+00000054h]
  loc_00E2DA76: test eax, eax
  loc_00E2DA78: fnclex
  loc_00E2DA7A: jge 00E2DA8Bh
  loc_00E2DA7C: push 00000054h
  loc_00E2DA7E: push 006E0410h
  loc_00E2DA83: push ebx
  loc_00E2DA84: push eax
  loc_00E2DA85: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DA8B: lea ecx, var_28
  loc_00E2DA8E: call [00401254h] ; __vbaFreeObj
  loc_00E2DA94: mov edx, [esi]
  loc_00E2DA96: push esi
  loc_00E2DA97: call [edx+0000037Ch]
  loc_00E2DA9D: push eax
  loc_00E2DA9E: lea eax, var_28
  loc_00E2DAA1: push eax
  loc_00E2DAA2: call edi
  loc_00E2DAA4: mov ebx, eax
  loc_00E2DAA6: push 00000000h
  loc_00E2DAA8: push ebx
  loc_00E2DAA9: mov ecx, [ebx]
  loc_00E2DAAB: call [ecx+0000009Ch]
  loc_00E2DAB1: test eax, eax
  loc_00E2DAB3: fnclex
  loc_00E2DAB5: jge 00E2DB9Dh
  loc_00E2DABB: jmp 00E2DB8Bh
  loc_00E2DAC0: call [edx+000003B0h]
  loc_00E2DAC6: push eax
  loc_00E2DAC7: lea eax, var_28
  loc_00E2DACA: push eax
  loc_00E2DACB: call edi
  loc_00E2DACD: mov ebx, eax
  loc_00E2DACF: lea edx, var_18
  loc_00E2DAD2: push edx
  loc_00E2DAD3: push ebx
  loc_00E2DAD4: mov ecx, [ebx]
  loc_00E2DAD6: call [ecx+00000050h]
  loc_00E2DAD9: test eax, eax
  loc_00E2DADB: fnclex
  loc_00E2DADD: jge 00E2DAEEh
  loc_00E2DADF: push 00000050h
  loc_00E2DAE1: push 006E0410h
  loc_00E2DAE6: push ebx
  loc_00E2DAE7: push eax
  loc_00E2DAE8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DAEE: mov eax, var_18
  loc_00E2DAF1: push eax
  loc_00E2DAF2: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2DAF8: call [004010D8h] ; __vbaFpR8
  loc_00E2DAFE: fcomp real8 ptr [004022E8h]
  loc_00E2DB04: fnstsw ax
  loc_00E2DB06: test ah, 41h
  loc_00E2DB09: jnz 00E2DB12h
  loc_00E2DB0B: mov ebx, 00000001h
  loc_00E2DB10: jmp 00E2DB14h
  loc_00E2DB12: xor ebx, ebx
  loc_00E2DB14: lea ecx, var_18
  loc_00E2DB17: call [00401258h] ; __vbaFreeStr
  loc_00E2DB1D: lea ecx, var_28
  loc_00E2DB20: call [00401254h] ; __vbaFreeObj
  loc_00E2DB26: neg ebx
  loc_00E2DB28: test bx, bx
  loc_00E2DB2B: jz 00E2DBA6h
  loc_00E2DB2D: mov ecx, [esi]
  loc_00E2DB2F: push esi
  loc_00E2DB30: call [ecx+00000424h]
  loc_00E2DB36: lea edx, var_28
  loc_00E2DB39: push eax
  loc_00E2DB3A: push edx
  loc_00E2DB3B: call edi
  loc_00E2DB3D: mov ebx, eax
  loc_00E2DB3F: push 006E1710h ; "Mencicil"
  loc_00E2DB44: push ebx
  loc_00E2DB45: mov eax, [ebx]
  loc_00E2DB47: call [eax+00000054h]
  loc_00E2DB4A: test eax, eax
  loc_00E2DB4C: fnclex
  loc_00E2DB4E: jge 00E2DB5Fh
  loc_00E2DB50: push 00000054h
  loc_00E2DB52: push 006E0410h
  loc_00E2DB57: push ebx
  loc_00E2DB58: push eax
  loc_00E2DB59: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DB5F: lea ecx, var_28
  loc_00E2DB62: call [00401254h] ; __vbaFreeObj
  loc_00E2DB68: mov ecx, [esi]
  loc_00E2DB6A: push esi
  loc_00E2DB6B: call [ecx+0000037Ch]
  loc_00E2DB71: lea edx, var_28
  loc_00E2DB74: push eax
  loc_00E2DB75: push edx
  loc_00E2DB76: call edi
  loc_00E2DB78: mov ebx, eax
  loc_00E2DB7A: push 00000000h
  loc_00E2DB7C: push ebx
  loc_00E2DB7D: mov eax, [ebx]
  loc_00E2DB7F: call [eax+0000009Ch]
  loc_00E2DB85: test eax, eax
  loc_00E2DB87: fnclex
  loc_00E2DB89: jge 00E2DB9Dh
  loc_00E2DB8B: push 0000009Ch
  loc_00E2DB90: push 006DCAD0h
  loc_00E2DB95: push ebx
  loc_00E2DB96: push eax
  loc_00E2DB97: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DB9D: lea ecx, var_28
  loc_00E2DBA0: call [00401254h] ; __vbaFreeObj
  loc_00E2DBA6: mov ecx, [esi]
  loc_00E2DBA8: push esi
  loc_00E2DBA9: call [ecx+00000300h]
  loc_00E2DBAF: lea edx, var_2C
  loc_00E2DBB2: push eax
  loc_00E2DBB3: push edx
  loc_00E2DBB4: call edi
  loc_00E2DBB6: mov ebx, eax
  loc_00E2DBB8: mov eax, [esi]
  loc_00E2DBBA: push esi
  loc_00E2DBBB: call [eax+00000378h]
  loc_00E2DBC1: lea ecx, var_28
  loc_00E2DBC4: push eax
  loc_00E2DBC5: push ecx
  loc_00E2DBC6: call edi
  loc_00E2DBC8: mov edi, eax
  loc_00E2DBCA: lea eax, var_18
  loc_00E2DBCD: push eax
  loc_00E2DBCE: push edi
  loc_00E2DBCF: mov edx, [edi]
  loc_00E2DBD1: call [edx+000000A0h]
  loc_00E2DBD7: test eax, eax
  loc_00E2DBD9: fnclex
  loc_00E2DBDB: jge 00E2DBF3h
  loc_00E2DBDD: push 000000A0h
  loc_00E2DBE2: push 006DCB70h
  loc_00E2DBE7: push edi
  loc_00E2DBE8: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DBEE: push eax
  loc_00E2DBEF: call edi
  loc_00E2DBF1: jmp 00E2DBF9h
  loc_00E2DBF3: mov edi, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DBF9: mov edx, var_18
  loc_00E2DBFC: lea ecx, var_1C
  loc_00E2DBFF: mov var_18, 00000000h
  loc_00E2DC06: call [00401228h] ; __vbaStrMove
  loc_00E2DC0C: mov ecx, [esi]
  loc_00E2DC0E: lea edx, var_20
  loc_00E2DC11: lea eax, var_1C
  loc_00E2DC14: push edx
  loc_00E2DC15: push eax
  loc_00E2DC16: push esi
  loc_00E2DC17: call [ecx+000006F8h]
  loc_00E2DC1D: test eax, eax
  loc_00E2DC1F: jge 00E2DC2Fh
  loc_00E2DC21: push 000006F8h
  loc_00E2DC26: push 006E0044h
  loc_00E2DC2B: push esi
  loc_00E2DC2C: push eax
  loc_00E2DC2D: call edi
  loc_00E2DC2F: mov edx, var_20
  loc_00E2DC32: mov ecx, [ebx]
  loc_00E2DC34: push edx
  loc_00E2DC35: push ebx
  loc_00E2DC36: call [ecx+000000A4h]
  loc_00E2DC3C: test eax, eax
  loc_00E2DC3E: fnclex
  loc_00E2DC40: jge 00E2DC50h
  loc_00E2DC42: push 000000A4h
  loc_00E2DC47: push 006DCB70h
  loc_00E2DC4C: push ebx
  loc_00E2DC4D: push eax
  loc_00E2DC4E: call edi
  loc_00E2DC50: lea eax, var_20
  loc_00E2DC53: lea ecx, var_1C
  loc_00E2DC56: push eax
  loc_00E2DC57: push ecx
  loc_00E2DC58: push 00000002h
  loc_00E2DC5A: call [004011BCh] ; __vbaFreeStrList
  loc_00E2DC60: lea edx, var_2C
  loc_00E2DC63: lea eax, var_28
  loc_00E2DC66: push edx
  loc_00E2DC67: push eax
  loc_00E2DC68: push 00000002h
  loc_00E2DC6A: call [00401048h] ; __vbaFreeObjList
  loc_00E2DC70: add esp, 00000018h
  loc_00E2DC73: mov var_4, 00000000h
  loc_00E2DC7A: fwait
  loc_00E2DC7B: push 00E2DCDBh
  loc_00E2DC80: jmp 00E2DCDAh
  loc_00E2DC82: lea ecx, var_24
  loc_00E2DC85: lea edx, var_20
  loc_00E2DC88: push ecx
  loc_00E2DC89: lea eax, var_1C
  loc_00E2DC8C: push edx
  loc_00E2DC8D: lea ecx, var_18
  loc_00E2DC90: push eax
  loc_00E2DC91: push ecx
  loc_00E2DC92: push 00000004h
  loc_00E2DC94: call [004011BCh] ; __vbaFreeStrList
  loc_00E2DC9A: lea edx, var_40
  loc_00E2DC9D: lea eax, var_3C
  loc_00E2DCA0: push edx
  loc_00E2DCA1: lea ecx, var_38
  loc_00E2DCA4: push eax
  loc_00E2DCA5: lea edx, var_34
  loc_00E2DCA8: push ecx
  loc_00E2DCA9: lea eax, var_30
  loc_00E2DCAC: push edx
  loc_00E2DCAD: lea ecx, var_2C
  loc_00E2DCB0: push eax
  loc_00E2DCB1: lea edx, var_28
  loc_00E2DCB4: push ecx
  loc_00E2DCB5: push edx
  loc_00E2DCB6: push 00000007h
  loc_00E2DCB8: call [00401048h] ; __vbaFreeObjList
  loc_00E2DCBE: lea eax, var_80
  loc_00E2DCC1: lea ecx, var_70
  loc_00E2DCC4: push eax
  loc_00E2DCC5: lea edx, var_60
  loc_00E2DCC8: push ecx
  loc_00E2DCC9: lea eax, var_50
  loc_00E2DCCC: push edx
  loc_00E2DCCD: push eax
  loc_00E2DCCE: push 00000004h
  loc_00E2DCD0: call [00401038h] ; __vbaFreeVarList
  loc_00E2DCD6: add esp, 00000048h
  loc_00E2DCD9: ret
  loc_00E2DCDA: ret
  loc_00E2DCDB: mov eax, Me
  loc_00E2DCDE: push eax
  loc_00E2DCDF: mov ecx, [eax]
  loc_00E2DCE1: call [ecx+00000008h]
  loc_00E2DCE4: mov eax, var_4
  loc_00E2DCE7: mov ecx, var_14
  loc_00E2DCEA: pop edi
  loc_00E2DCEB: pop esi
  loc_00E2DCEC: mov fs:[00000000h], ecx
  loc_00E2DCF3: pop ebx
  loc_00E2DCF4: mov esp, ebp
  loc_00E2DCF6: pop ebp
  loc_00E2DCF7: retn 0004h
End Sub

Private Sub back_Click() 'E24D60
  loc_00E24D60: push ebp
  loc_00E24D61: mov ebp, esp
  loc_00E24D63: sub esp, 0000000Ch
  loc_00E24D66: push 00402836h ; __vbaExceptHandler
  loc_00E24D6B: mov eax, fs:[00000000h]
  loc_00E24D71: push eax
  loc_00E24D72: mov fs:[00000000h], esp
  loc_00E24D79: sub esp, 00000034h
  loc_00E24D7C: push ebx
  loc_00E24D7D: push esi
  loc_00E24D7E: push edi
  loc_00E24D7F: mov var_C, esp
  loc_00E24D82: mov var_8, 00402480h
  loc_00E24D89: mov edi, Me
  loc_00E24D8C: mov eax, edi
  loc_00E24D8E: and eax, 00000001h
  loc_00E24D91: mov var_4, eax
  loc_00E24D94: and edi, FFFFFFFEh
  loc_00E24D97: push edi
  loc_00E24D98: mov Me, edi
  loc_00E24D9B: mov ecx, [edi]
  loc_00E24D9D: call [ecx+00000004h]
  loc_00E24DA0: mov eax, [00E3D9CCh]
  loc_00E24DA5: mov var_18, 00000000h
  loc_00E24DAC: test eax, eax
  loc_00E24DAE: jnz 00E24DC0h
  loc_00E24DB0: push 00E3D9CCh
  loc_00E24DB5: push 006DC960h
  loc_00E24DBA: call [004011A0h] ; __vbaNew2
  loc_00E24DC0: mov esi, [00E3D9CCh]
  loc_00E24DC6: lea edx, var_18
  loc_00E24DC9: push edi
  loc_00E24DCA: push edx
  loc_00E24DCB: mov ebx, [esi]
  loc_00E24DCD: call [004010B4h] ; __vbaObjSetAddref
  loc_00E24DD3: push eax
  loc_00E24DD4: push esi
  loc_00E24DD5: call [ebx+00000010h]
  loc_00E24DD8: test eax, eax
  loc_00E24DDA: fnclex
  loc_00E24DDC: jge 00E24DEDh
  loc_00E24DDE: push 00000010h
  loc_00E24DE0: push 006DC950h
  loc_00E24DE5: push esi
  loc_00E24DE6: push eax
  loc_00E24DE7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E24DED: lea ecx, var_18
  loc_00E24DF0: call [00401254h] ; __vbaFreeObj
  loc_00E24DF6: mov eax, [00E3D060h]
  loc_00E24DFB: test eax, eax
  loc_00E24DFD: jnz 00E24E0Fh
  loc_00E24DFF: push 00E3D060h
  loc_00E24E04: push 006D87C0h
  loc_00E24E09: call [004011A0h] ; __vbaNew2
  loc_00E24E0F: sub esp, 00000010h
  loc_00E24E12: mov ecx, 0000000Ah
  loc_00E24E17: mov ebx, esp
  loc_00E24E19: mov var_28, ecx
  loc_00E24E1C: mov eax, 80020004h
  loc_00E24E21: sub esp, 00000010h
  loc_00E24E24: mov [ebx], ecx
  loc_00E24E26: mov ecx, var_34
  loc_00E24E29: mov edx, eax
  loc_00E24E2B: mov esi, [00E3D060h]
  loc_00E24E31: mov [ebx+00000004h], ecx
  loc_00E24E34: mov ecx, esp
  loc_00E24E36: mov edi, [esi]
  loc_00E24E38: push esi
  loc_00E24E39: mov [ebx+00000008h], eax
  loc_00E24E3C: mov eax, var_2C
  loc_00E24E3F: mov [ebx+0000000Ch], eax
  loc_00E24E42: mov eax, var_28
  loc_00E24E45: mov [ecx], eax
  loc_00E24E47: mov eax, var_24
  loc_00E24E4A: mov [ecx+00000004h], eax
  loc_00E24E4D: mov [ecx+00000008h], edx
  loc_00E24E50: mov edx, var_1C
  loc_00E24E53: mov [ecx+0000000Ch], edx
  loc_00E24E56: call [edi+000002B0h]
  loc_00E24E5C: test eax, eax
  loc_00E24E5E: fnclex
  loc_00E24E60: jge 00E24E74h
  loc_00E24E62: push 000002B0h
  loc_00E24E67: push 006DFA4Ch
  loc_00E24E6C: push esi
  loc_00E24E6D: push eax
  loc_00E24E6E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E24E74: mov var_4, 00000000h
  loc_00E24E7B: push 00E24E8Dh
  loc_00E24E80: jmp 00E24E8Ch
  loc_00E24E82: lea ecx, var_18
  loc_00E24E85: call [00401254h] ; __vbaFreeObj
  loc_00E24E8B: ret
  loc_00E24E8C: ret
  loc_00E24E8D: mov eax, Me
  loc_00E24E90: push eax
  loc_00E24E91: mov ecx, [eax]
  loc_00E24E93: call [ecx+00000008h]
  loc_00E24E96: mov eax, var_4
  loc_00E24E99: mov ecx, var_14
  loc_00E24E9C: pop edi
  loc_00E24E9D: pop esi
  loc_00E24E9E: mov fs:[00000000h], ecx
  loc_00E24EA5: pop ebx
  loc_00E24EA6: mov esp, ebp
  loc_00E24EA8: pop ebp
  loc_00E24EA9: retn 0004h
End Sub

Private Sub Form_Load() 'E25390
  loc_00E25390: push ebp
  loc_00E25391: mov ebp, esp
  loc_00E25393: sub esp, 0000000Ch
  loc_00E25396: push 00402836h ; __vbaExceptHandler
  loc_00E2539B: mov eax, fs:[00000000h]
  loc_00E253A1: push eax
  loc_00E253A2: mov fs:[00000000h], esp
  loc_00E253A9: sub esp, 00000064h
  loc_00E253AC: push ebx
  loc_00E253AD: push esi
  loc_00E253AE: push edi
  loc_00E253AF: mov var_C, esp
  loc_00E253B2: mov var_8, 004024C0h
  loc_00E253B9: mov esi, Me
  loc_00E253BC: mov eax, esi
  loc_00E253BE: and eax, 00000001h
  loc_00E253C1: mov var_4, eax
  loc_00E253C4: and esi, FFFFFFFEh
  loc_00E253C7: push esi
  loc_00E253C8: mov Me, esi
  loc_00E253CB: mov ecx, [esi]
  loc_00E253CD: call [ecx+00000004h]
  loc_00E253D0: sub esp, 00000010h
  loc_00E253D3: xor eax, eax
  loc_00E253D5: mov edx, esp
  loc_00E253D7: mov ecx, 0000000Bh
  loc_00E253DC: mov var_50, eax
  loc_00E253DF: mov var_50, ecx
  loc_00E253E2: mov [edx], ecx
  loc_00E253E4: mov ecx, var_4C
  loc_00E253E7: mov var_24, eax
  loc_00E253EA: mov var_28, eax
  loc_00E253ED: mov [edx+00000004h], ecx
  loc_00E253F0: mov ecx, [esi]
  loc_00E253F2: mov var_2C, eax
  loc_00E253F5: mov var_30, eax
  loc_00E253F8: mov var_40, eax
  loc_00E253FB: mov var_48, eax
  loc_00E253FE: mov [edx+00000008h], eax
  loc_00E25401: mov eax, var_44
  loc_00E25404: push 80010007h
  loc_00E25409: push esi
  loc_00E2540A: mov [edx+0000000Ch], eax
  loc_00E2540D: call [ecx+00000400h]
  loc_00E25413: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E25419: lea edx, var_2C
  loc_00E2541C: push eax
  loc_00E2541D: push edx
  loc_00E2541E: call edi
  loc_00E25420: push eax
  loc_00E25421: call [00401238h] ; __vbaLateIdSt
  loc_00E25427: lea ecx, var_2C
  loc_00E2542A: call [00401254h] ; __vbaFreeObj
  loc_00E25430: mov eax, [esi]
  loc_00E25432: push esi
  loc_00E25433: call [eax+00000378h]
  loc_00E25439: lea ecx, var_2C
  loc_00E2543C: push eax
  loc_00E2543D: push ecx
  loc_00E2543E: call edi
  loc_00E25440: mov ebx, eax
  loc_00E25442: push 006E0934h
  loc_00E25447: push ebx
  loc_00E25448: mov edx, [ebx]
  loc_00E2544A: call [edx+000000A4h]
  loc_00E25450: test eax, eax
  loc_00E25452: fnclex
  loc_00E25454: jge 00E25468h
  loc_00E25456: push 000000A4h
  loc_00E2545B: push 006DCB70h
  loc_00E25460: push ebx
  loc_00E25461: push eax
  loc_00E25462: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25468: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E2546E: lea ecx, var_2C
  loc_00E25471: call ebx
  loc_00E25473: sub esp, 00000010h
  loc_00E25476: mov ecx, 0000000Bh
  loc_00E2547B: mov edx, esp
  loc_00E2547D: mov var_50, ecx
  loc_00E25480: xor eax, eax
  loc_00E25482: push 80010007h
  loc_00E25487: mov [edx], ecx
  loc_00E25489: mov ecx, var_4C
  loc_00E2548C: mov var_48, eax
  loc_00E2548F: push esi
  loc_00E25490: mov [edx+00000004h], ecx
  loc_00E25493: mov ecx, [esi]
  loc_00E25495: mov [edx+00000008h], eax
  loc_00E25498: mov eax, var_44
  loc_00E2549B: mov [edx+0000000Ch], eax
  loc_00E2549E: call [ecx+00000458h]
  loc_00E254A4: lea edx, var_2C
  loc_00E254A7: push eax
  loc_00E254A8: push edx
  loc_00E254A9: call edi
  loc_00E254AB: push eax
  loc_00E254AC: call [00401238h] ; __vbaLateIdSt
  loc_00E254B2: lea ecx, var_2C
  loc_00E254B5: call ebx
  loc_00E254B7: mov eax, [esi]
  loc_00E254B9: push esi
  loc_00E254BA: call [eax+00000414h]
  loc_00E254C0: lea ecx, var_2C
  loc_00E254C3: push eax
  loc_00E254C4: push ecx
  loc_00E254C5: call edi
  loc_00E254C7: mov ebx, eax
  loc_00E254C9: push 00000000h
  loc_00E254CB: push ebx
  loc_00E254CC: mov edx, [ebx]
  loc_00E254CE: call [edx+0000008Ch]
  loc_00E254D4: test eax, eax
  loc_00E254D6: fnclex
  loc_00E254D8: jge 00E254ECh
  loc_00E254DA: push 0000008Ch
  loc_00E254DF: push 006DCDA0h
  loc_00E254E4: push ebx
  loc_00E254E5: push eax
  loc_00E254E6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E254EC: lea ecx, var_2C
  loc_00E254EF: call [00401254h] ; __vbaFreeObj
  loc_00E254F5: mov eax, [esi]
  loc_00E254F7: push esi
  loc_00E254F8: call [eax+0000039Ch]
  loc_00E254FE: lea ecx, var_2C
  loc_00E25501: push eax
  loc_00E25502: push ecx
  loc_00E25503: call edi
  loc_00E25505: mov ebx, eax
  loc_00E25507: push 006E1138h ; "-"
  loc_00E2550C: push ebx
  loc_00E2550D: mov edx, [ebx]
  loc_00E2550F: call [edx+000000A4h]
  loc_00E25515: test eax, eax
  loc_00E25517: fnclex
  loc_00E25519: jge 00E2552Dh
  loc_00E2551B: push 000000A4h
  loc_00E25520: push 006DCB70h
  loc_00E25525: push ebx
  loc_00E25526: push eax
  loc_00E25527: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2552D: lea ecx, var_2C
  loc_00E25530: call [00401254h] ; __vbaFreeObj
  loc_00E25536: mov eax, [esi]
  loc_00E25538: push esi
  loc_00E25539: call [eax+00000490h]
  loc_00E2553F: lea ecx, var_2C
  loc_00E25542: push eax
  loc_00E25543: push ecx
  loc_00E25544: call edi
  loc_00E25546: lea edx, var_40
  loc_00E25549: mov ebx, eax
  loc_00E2554B: push edx
  loc_00E2554C: mov var_64, ebx
  loc_00E2554F: call [004011E8h] ; rtcGetTimeVar
  loc_00E25555: mov ebx, [ebx]
  loc_00E25557: lea eax, var_40
  loc_00E2555A: lea ecx, var_28
  loc_00E2555D: push eax
  loc_00E2555E: push ecx
  loc_00E2555F: call [00401180h] ; __vbaStrVarVal
  loc_00E25565: mov edx, ebx
  loc_00E25567: mov ebx, var_64
  loc_00E2556A: push eax
  loc_00E2556B: push ebx
  loc_00E2556C: call [edx+00000054h]
  loc_00E2556F: test eax, eax
  loc_00E25571: fnclex
  loc_00E25573: jge 00E25584h
  loc_00E25575: push 00000054h
  loc_00E25577: push 006E0410h
  loc_00E2557C: push ebx
  loc_00E2557D: push eax
  loc_00E2557E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25584: lea ecx, var_28
  loc_00E25587: call [00401258h] ; __vbaFreeStr
  loc_00E2558D: lea ecx, var_2C
  loc_00E25590: call [00401254h] ; __vbaFreeObj
  loc_00E25596: lea ecx, var_40
  loc_00E25599: call [00401024h] ; __vbaFreeVar
  loc_00E2559F: mov eax, [esi]
  loc_00E255A1: push esi
  loc_00E255A2: call [eax+00000488h]
  loc_00E255A8: lea ecx, var_2C
  loc_00E255AB: push eax
  loc_00E255AC: push ecx
  loc_00E255AD: call edi
  loc_00E255AF: lea edx, var_40
  loc_00E255B2: mov ebx, eax
  loc_00E255B4: push edx
  loc_00E255B5: mov var_64, ebx
  loc_00E255B8: call [004011D8h] ; rtcGetDateVar
  loc_00E255BE: mov ebx, [ebx]
  loc_00E255C0: lea eax, var_40
  loc_00E255C3: lea ecx, var_28
  loc_00E255C6: push eax
  loc_00E255C7: push ecx
  loc_00E255C8: call [00401180h] ; __vbaStrVarVal
  loc_00E255CE: mov edx, ebx
  loc_00E255D0: mov ebx, var_64
  loc_00E255D3: push eax
  loc_00E255D4: push ebx
  loc_00E255D5: call [edx+00000054h]
  loc_00E255D8: test eax, eax
  loc_00E255DA: fnclex
  loc_00E255DC: jge 00E255EDh
  loc_00E255DE: push 00000054h
  loc_00E255E0: push 006E0410h
  loc_00E255E5: push ebx
  loc_00E255E6: push eax
  loc_00E255E7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E255ED: lea ecx, var_28
  loc_00E255F0: call [00401258h] ; __vbaFreeStr
  loc_00E255F6: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E255FC: lea ecx, var_2C
  loc_00E255FF: call ebx
  loc_00E25601: lea ecx, var_40
  loc_00E25604: call [00401024h] ; __vbaFreeVar
  loc_00E2560A: mov eax, [esi]
  loc_00E2560C: push esi
  loc_00E2560D: call [eax+00000320h]
  loc_00E25613: lea ecx, var_2C
  loc_00E25616: push eax
  loc_00E25617: push ecx
  loc_00E25618: call edi
  loc_00E2561A: mov edx, [eax]
  loc_00E2561C: push FFFFFFFFh
  loc_00E2561E: push eax
  loc_00E2561F: mov var_64, eax
  loc_00E25622: call [edx+0000009Ch]
  loc_00E25628: test eax, eax
  loc_00E2562A: fnclex
  loc_00E2562C: jge 00E25643h
  loc_00E2562E: mov ecx, var_64
  loc_00E25631: push 0000009Ch
  loc_00E25636: push 006DCAD0h
  loc_00E2563B: push ecx
  loc_00E2563C: push eax
  loc_00E2563D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25643: lea ecx, var_2C
  loc_00E25646: call ebx
  loc_00E25648: mov edx, [esi]
  loc_00E2564A: push esi
  loc_00E2564B: call [edx+0000034Ch]
  loc_00E25651: push eax
  loc_00E25652: lea eax, var_2C
  loc_00E25655: push eax
  loc_00E25656: call edi
  loc_00E25658: mov ecx, [eax]
  loc_00E2565A: push 00000000h
  loc_00E2565C: push eax
  loc_00E2565D: mov var_64, eax
  loc_00E25660: call [ecx+0000009Ch]
  loc_00E25666: test eax, eax
  loc_00E25668: fnclex
  loc_00E2566A: jge 00E25681h
  loc_00E2566C: mov edx, var_64
  loc_00E2566F: push 0000009Ch
  loc_00E25674: push 006DCAD0h
  loc_00E25679: push edx
  loc_00E2567A: push eax
  loc_00E2567B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25681: lea ecx, var_2C
  loc_00E25684: call ebx
  loc_00E25686: mov eax, [esi]
  loc_00E25688: push esi
  loc_00E25689: call [eax+00000394h]
  loc_00E2568F: lea ecx, var_2C
  loc_00E25692: push eax
  loc_00E25693: push ecx
  loc_00E25694: call edi
  loc_00E25696: mov edx, [eax]
  loc_00E25698: push 00000000h
  loc_00E2569A: push eax
  loc_00E2569B: mov var_64, eax
  loc_00E2569E: call [edx+0000009Ch]
  loc_00E256A4: test eax, eax
  loc_00E256A6: fnclex
  loc_00E256A8: jge 00E256BFh
  loc_00E256AA: mov ecx, var_64
  loc_00E256AD: push 0000009Ch
  loc_00E256B2: push 006E0410h
  loc_00E256B7: push ecx
  loc_00E256B8: push eax
  loc_00E256B9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E256BF: lea ecx, var_2C
  loc_00E256C2: call ebx
  loc_00E256C4: mov edx, [esi]
  loc_00E256C6: push esi
  loc_00E256C7: call [edx+00000384h]
  loc_00E256CD: push eax
  loc_00E256CE: lea eax, var_2C
  loc_00E256D1: push eax
  loc_00E256D2: call edi
  loc_00E256D4: mov ecx, [eax]
  loc_00E256D6: push 00000000h
  loc_00E256D8: push eax
  loc_00E256D9: mov var_64, eax
  loc_00E256DC: call [ecx+0000008Ch]
  loc_00E256E2: test eax, eax
  loc_00E256E4: fnclex
  loc_00E256E6: jge 00E256FDh
  loc_00E256E8: mov edx, var_64
  loc_00E256EB: push 0000008Ch
  loc_00E256F0: push 006DCAC0h
  loc_00E256F5: push edx
  loc_00E256F6: push eax
  loc_00E256F7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E256FD: lea ecx, var_2C
  loc_00E25700: call ebx
  loc_00E25702: lea edx, var_50
  loc_00E25705: lea ecx, var_24
  loc_00E25708: mov var_48, 006E15ECh ; "select * from lc"
  loc_00E2570F: mov var_50, 00000008h
  loc_00E25716: call [00401204h] ; __vbaVarCopy
  loc_00E2571C: lea eax, var_24
  loc_00E2571F: push eax
  loc_00E25720: call [00401230h] ; __vbaStrVarCopy
  loc_00E25726: sub esp, 00000010h
  loc_00E25729: mov ecx, 00000008h
  loc_00E2572E: mov edx, esp
  loc_00E25730: mov var_40, ecx
  loc_00E25733: mov var_38, eax
  loc_00E25736: push 0000000Eh
  loc_00E25738: mov [edx], ecx
  loc_00E2573A: mov ecx, var_3C
  loc_00E2573D: push esi
  loc_00E2573E: mov [edx+00000004h], ecx
  loc_00E25741: mov ecx, [esi]
  loc_00E25743: mov [edx+00000008h], eax
  loc_00E25746: mov eax, var_34
  loc_00E25749: mov [edx+0000000Ch], eax
  loc_00E2574C: call [ecx+000004A8h]
  loc_00E25752: lea edx, var_2C
  loc_00E25755: push eax
  loc_00E25756: push edx
  loc_00E25757: call edi
  loc_00E25759: push eax
  loc_00E2575A: call [00401238h] ; __vbaLateIdSt
  loc_00E25760: lea ecx, var_2C
  loc_00E25763: call ebx
  loc_00E25765: lea ecx, var_40
  loc_00E25768: call [00401024h] ; __vbaFreeVar
  loc_00E2576E: mov eax, [esi]
  loc_00E25770: push 00000000h
  loc_00E25772: push 00000019h
  loc_00E25774: push esi
  loc_00E25775: call [eax+000004A8h]
  loc_00E2577B: lea ecx, var_2C
  loc_00E2577E: push eax
  loc_00E2577F: push ecx
  loc_00E25780: call edi
  loc_00E25782: push eax
  loc_00E25783: call [00401030h] ; __vbaLateIdCall
  loc_00E25789: add esp, 0000000Ch
  loc_00E2578C: lea ecx, var_2C
  loc_00E2578F: call ebx
  loc_00E25791: mov edx, [esi]
  loc_00E25793: push 006E05D8h
  loc_00E25798: push esi
  loc_00E25799: call [edx+000004A8h]
  loc_00E2579F: push eax
  loc_00E257A0: lea eax, var_2C
  loc_00E257A3: push eax
  loc_00E257A4: call edi
  loc_00E257A6: push eax
  loc_00E257A7: call [00401224h] ; __vbaCastObj
  loc_00E257AD: lea ecx, var_40
  loc_00E257B0: mov var_38, eax
  loc_00E257B3: push ecx
  loc_00E257B4: mov var_40, 0000000Dh
  loc_00E257BB: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E257C1: mov ecx, [eax]
  loc_00E257C3: sub esp, 00000010h
  loc_00E257C6: mov edx, esp
  loc_00E257C8: push 00000000h
  loc_00E257CA: push 0000002Ah
  loc_00E257CC: mov [edx], ecx
  loc_00E257CE: mov ecx, [eax+00000004h]
  loc_00E257D1: push esi
  loc_00E257D2: mov [edx+00000004h], ecx
  loc_00E257D5: mov ecx, [eax+00000008h]
  loc_00E257D8: mov eax, [eax+0000000Ch]
  loc_00E257DB: mov [edx+00000008h], ecx
  loc_00E257DE: mov ecx, [esi]
  loc_00E257E0: mov [edx+0000000Ch], eax
  loc_00E257E3: call [ecx+000004ACh]
  loc_00E257E9: push eax
  loc_00E257EA: lea edx, var_30
  loc_00E257ED: push edx
  loc_00E257EE: call edi
  loc_00E257F0: push eax
  loc_00E257F1: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E257F7: lea eax, var_30
  loc_00E257FA: lea ecx, var_2C
  loc_00E257FD: push eax
  loc_00E257FE: push ecx
  loc_00E257FF: push 00000002h
  loc_00E25801: call [00401048h] ; __vbaFreeObjList
  loc_00E25807: add esp, 00000028h
  loc_00E2580A: lea ecx, var_40
  loc_00E2580D: call [00401024h] ; __vbaFreeVar
  loc_00E25813: mov edx, [esi]
  loc_00E25815: push 00000000h
  loc_00E25817: push FFFFFDDAh
  loc_00E2581C: push esi
  loc_00E2581D: call [edx+000004ACh]
  loc_00E25823: push eax
  loc_00E25824: lea eax, var_2C
  loc_00E25827: push eax
  loc_00E25828: call edi
  loc_00E2582A: push eax
  loc_00E2582B: call [00401030h] ; __vbaLateIdCall
  loc_00E25831: add esp, 0000000Ch
  loc_00E25834: lea ecx, var_2C
  loc_00E25837: call ebx
  loc_00E25839: mov ecx, [esi]
  loc_00E2583B: push esi
  loc_00E2583C: call [ecx+00000308h]
  loc_00E25842: lea edx, var_2C
  loc_00E25845: push eax
  loc_00E25846: push edx
  loc_00E25847: call edi
  loc_00E25849: mov ecx, [eax]
  loc_00E2584B: push 00000000h
  loc_00E2584D: push eax
  loc_00E2584E: mov var_64, eax
  loc_00E25851: call [ecx+0000009Ch]
  loc_00E25857: test eax, eax
  loc_00E25859: fnclex
  loc_00E2585B: jge 00E25872h
  loc_00E2585D: mov edx, var_64
  loc_00E25860: push 0000009Ch
  loc_00E25865: push 006DCAD0h
  loc_00E2586A: push edx
  loc_00E2586B: push eax
  loc_00E2586C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25872: lea ecx, var_2C
  loc_00E25875: call ebx
  loc_00E25877: sub esp, 00000010h
  loc_00E2587A: mov ecx, 0000000Bh
  loc_00E2587F: mov edx, esp
  loc_00E25881: mov var_50, ecx
  loc_00E25884: xor eax, eax
  loc_00E25886: push 8001000Dh
  loc_00E2588B: mov [edx], ecx
  loc_00E2588D: mov ecx, var_4C
  loc_00E25890: mov var_48, eax
  loc_00E25893: push esi
  loc_00E25894: mov [edx+00000004h], ecx
  loc_00E25897: mov ecx, [esi]
  loc_00E25899: mov [edx+00000008h], eax
  loc_00E2589C: mov eax, var_44
  loc_00E2589F: mov [edx+0000000Ch], eax
  loc_00E258A2: call [ecx+00000408h]
  loc_00E258A8: lea edx, var_2C
  loc_00E258AB: push eax
  loc_00E258AC: push edx
  loc_00E258AD: call edi
  loc_00E258AF: push eax
  loc_00E258B0: call [00401238h] ; __vbaLateIdSt
  loc_00E258B6: lea ecx, var_2C
  loc_00E258B9: call ebx
  loc_00E258BB: sub esp, 00000010h
  loc_00E258BE: mov ecx, 0000000Bh
  loc_00E258C3: mov edx, esp
  loc_00E258C5: mov var_50, ecx
  loc_00E258C8: xor eax, eax
  loc_00E258CA: push 80010007h
  loc_00E258CF: mov [edx], ecx
  loc_00E258D1: mov ecx, var_4C
  loc_00E258D4: mov var_48, eax
  loc_00E258D7: push esi
  loc_00E258D8: mov [edx+00000004h], ecx
  loc_00E258DB: mov ecx, [esi]
  loc_00E258DD: mov [edx+00000008h], eax
  loc_00E258E0: mov eax, var_44
  loc_00E258E3: mov [edx+0000000Ch], eax
  loc_00E258E6: call [ecx+000003A0h]
  loc_00E258EC: lea edx, var_2C
  loc_00E258EF: push eax
  loc_00E258F0: push edx
  loc_00E258F1: call edi
  loc_00E258F3: push eax
  loc_00E258F4: call [00401238h] ; __vbaLateIdSt
  loc_00E258FA: lea ecx, var_2C
  loc_00E258FD: call ebx
  loc_00E258FF: mov eax, [esi]
  loc_00E25901: push esi
  loc_00E25902: call [eax+00000350h]
  loc_00E25908: lea ecx, var_2C
  loc_00E2590B: push eax
  loc_00E2590C: push ecx
  loc_00E2590D: call edi
  loc_00E2590F: mov esi, eax
  loc_00E25911: push 00000000h
  loc_00E25913: push esi
  loc_00E25914: mov edx, [esi]
  loc_00E25916: call [edx+0000009Ch]
  loc_00E2591C: test eax, eax
  loc_00E2591E: fnclex
  loc_00E25920: jge 00E25934h
  loc_00E25922: push 0000009Ch
  loc_00E25927: push 006DCAD0h
  loc_00E2592C: push esi
  loc_00E2592D: push eax
  loc_00E2592E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25934: lea ecx, var_2C
  loc_00E25937: call ebx
  loc_00E25939: mov var_4, 00000000h
  loc_00E25940: push 00E25977h
  loc_00E25945: jmp 00E2596Dh
  loc_00E25947: lea ecx, var_28
  loc_00E2594A: call [00401258h] ; __vbaFreeStr
  loc_00E25950: lea eax, var_30
  loc_00E25953: lea ecx, var_2C
  loc_00E25956: push eax
  loc_00E25957: push ecx
  loc_00E25958: push 00000002h
  loc_00E2595A: call [00401048h] ; __vbaFreeObjList
  loc_00E25960: add esp, 0000000Ch
  loc_00E25963: lea ecx, var_40
  loc_00E25966: call [00401024h] ; __vbaFreeVar
  loc_00E2596C: ret
  loc_00E2596D: lea ecx, var_24
  loc_00E25970: call [00401024h] ; __vbaFreeVar
  loc_00E25976: ret
  loc_00E25977: mov eax, Me
  loc_00E2597A: push eax
  loc_00E2597B: mov edx, [eax]
  loc_00E2597D: call [edx+00000008h]
  loc_00E25980: mov eax, var_4
  loc_00E25983: mov ecx, var_14
  loc_00E25986: pop edi
  loc_00E25987: pop esi
  loc_00E25988: mov fs:[00000000h], ecx
  loc_00E2598F: pop ebx
  loc_00E25990: mov esp, ebp
  loc_00E25992: pop ebp
  loc_00E25993: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E259A0
  loc_00E259A0: push ebp
  loc_00E259A1: mov ebp, esp
  loc_00E259A3: sub esp, 0000000Ch
  loc_00E259A6: push 00402836h ; __vbaExceptHandler
  loc_00E259AB: mov eax, fs:[00000000h]
  loc_00E259B1: push eax
  loc_00E259B2: mov fs:[00000000h], esp
  loc_00E259B9: sub esp, 0000005Ch
  loc_00E259BC: push ebx
  loc_00E259BD: push esi
  loc_00E259BE: push edi
  loc_00E259BF: mov var_C, esp
  loc_00E259C2: mov var_8, 004024D0h
  loc_00E259C9: mov esi, Me
  loc_00E259CC: mov eax, esi
  loc_00E259CE: and eax, 00000001h
  loc_00E259D1: mov var_4, eax
  loc_00E259D4: and esi, FFFFFFFEh
  loc_00E259D7: push esi
  loc_00E259D8: mov Me, esi
  loc_00E259DB: mov ecx, [esi]
  loc_00E259DD: call [ecx+00000004h]
  loc_00E259E0: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E259E6: xor eax, eax
  loc_00E259E8: mov var_18, eax
  loc_00E259EB: mov var_4C, eax
  loc_00E259EE: mov var_50, eax
  loc_00E259F1: mov edx, [esi]
  loc_00E259F3: lea eax, var_4C
  loc_00E259F6: push eax
  loc_00E259F7: push esi
  loc_00E259F8: call [edx+00000070h]
  loc_00E259FB: test eax, eax
  loc_00E259FD: fnclex
  loc_00E259FF: jge 00E25A0Ch
  loc_00E25A01: push 00000070h
  loc_00E25A03: push 006E0014h
  loc_00E25A08: push esi
  loc_00E25A09: push eax
  loc_00E25A0A: call ebx
  loc_00E25A0C: fld real4 ptr var_4C
  loc_00E25A0F: fadd st0, real4 ptr [00401298h]
  loc_00E25A15: mov ecx, [esi]
  loc_00E25A17: push ecx
  loc_00E25A18: fnstsw ax
  loc_00E25A1A: test al, 0Dh
  loc_00E25A1C: jnz 00E25BE0h
  loc_00E25A22: fstp real4 ptr [esp]
  loc_00E25A25: push esi
  loc_00E25A26: call [ecx+00000074h]
  loc_00E25A29: test eax, eax
  loc_00E25A2B: fnclex
  loc_00E25A2D: jge 00E25A3Ah
  loc_00E25A2F: push 00000074h
  loc_00E25A31: push 006E0014h
  loc_00E25A36: push esi
  loc_00E25A37: push eax
  loc_00E25A38: call ebx
  loc_00E25A3A: mov edx, [esi]
  loc_00E25A3C: lea eax, var_4C
  loc_00E25A3F: push eax
  loc_00E25A40: push esi
  loc_00E25A41: call [edx+00000070h]
  loc_00E25A44: test eax, eax
  loc_00E25A46: fnclex
  loc_00E25A48: jge 00E25A55h
  loc_00E25A4A: push 00000070h
  loc_00E25A4C: push 006E0014h
  loc_00E25A51: push esi
  loc_00E25A52: push eax
  loc_00E25A53: call ebx
  loc_00E25A55: mov ecx, [esi]
  loc_00E25A57: lea edx, var_50
  loc_00E25A5A: push edx
  loc_00E25A5B: push esi
  loc_00E25A5C: call [ecx+00000078h]
  loc_00E25A5F: test eax, eax
  loc_00E25A61: fnclex
  loc_00E25A63: jge 00E25A70h
  loc_00E25A65: push 00000078h
  loc_00E25A67: push 006E0014h
  loc_00E25A6C: push esi
  loc_00E25A6D: push eax
  loc_00E25A6E: call ebx
  loc_00E25A70: sub esp, 00000010h
  loc_00E25A73: mov ecx, 0000000Ah
  loc_00E25A78: mov ebx, esp
  loc_00E25A7A: mov eax, 80020004h
  loc_00E25A7F: mov edx, eax
  loc_00E25A81: sub esp, 00000010h
  loc_00E25A84: mov [ebx], ecx
  loc_00E25A86: mov ecx, var_44
  loc_00E25A89: mov edi, [esi]
  loc_00E25A8B: mov [ebx+00000004h], ecx
  loc_00E25A8E: mov ecx, esp
  loc_00E25A90: sub esp, 00000010h
  loc_00E25A93: mov [ebx+00000008h], eax
  loc_00E25A96: mov eax, var_3C
  loc_00E25A99: mov [ebx+0000000Ch], eax
  loc_00E25A9C: mov eax, 0000000Ah
  loc_00E25AA1: mov [ecx], eax
  loc_00E25AA3: mov eax, var_34
  loc_00E25AA6: mov [ecx+00000004h], eax
  loc_00E25AA9: mov eax, 00000004h
  loc_00E25AAE: mov [ecx+00000008h], edx
  loc_00E25AB1: mov edx, var_2C
  loc_00E25AB4: mov [ecx+0000000Ch], edx
  loc_00E25AB7: mov edx, var_24
  loc_00E25ABA: mov ecx, esp
  loc_00E25ABC: mov [ecx], eax
  loc_00E25ABE: mov eax, var_50
  loc_00E25AC1: mov [ecx+00000004h], edx
  loc_00E25AC4: mov [ecx+00000008h], eax
  loc_00E25AC7: mov eax, var_1C
  loc_00E25ACA: mov [ecx+0000000Ch], eax
  loc_00E25ACD: mov ecx, var_4C
  loc_00E25AD0: push ecx
  loc_00E25AD1: push esi
  loc_00E25AD2: call [edi+000002A4h]
  loc_00E25AD8: test eax, eax
  loc_00E25ADA: fnclex
  loc_00E25ADC: jge 00E25AF4h
  loc_00E25ADE: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25AE4: push 000002A4h
  loc_00E25AE9: push 006E0014h
  loc_00E25AEE: push esi
  loc_00E25AEF: push eax
  loc_00E25AF0: call ebx
  loc_00E25AF2: jmp 00E25AFAh
  loc_00E25AF4: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25AFA: call [004010BCh] ; rtcDoEvents
  loc_00E25B00: mov edx, [esi]
  loc_00E25B02: lea eax, var_4C
  loc_00E25B05: push eax
  loc_00E25B06: push esi
  loc_00E25B07: call [edx+00000070h]
  loc_00E25B0A: test eax, eax
  loc_00E25B0C: fnclex
  loc_00E25B0E: jge 00E25B1Bh
  loc_00E25B10: push 00000070h
  loc_00E25B12: push 006E0014h
  loc_00E25B17: push esi
  loc_00E25B18: push eax
  loc_00E25B19: call ebx
  loc_00E25B1B: mov eax, [00E3D9CCh]
  loc_00E25B20: test eax, eax
  loc_00E25B22: jnz 00E25B34h
  loc_00E25B24: push 00E3D9CCh
  loc_00E25B29: push 006DC960h
  loc_00E25B2E: call [004011A0h] ; __vbaNew2
  loc_00E25B34: mov edi, [00E3D9CCh]
  loc_00E25B3A: lea edx, var_18
  loc_00E25B3D: push edx
  loc_00E25B3E: push edi
  loc_00E25B3F: mov ecx, [edi]
  loc_00E25B41: call [ecx+00000018h]
  loc_00E25B44: test eax, eax
  loc_00E25B46: fnclex
  loc_00E25B48: jge 00E25B55h
  loc_00E25B4A: push 00000018h
  loc_00E25B4C: push 006DC950h
  loc_00E25B51: push edi
  loc_00E25B52: push eax
  loc_00E25B53: call ebx
  loc_00E25B55: mov eax, var_18
  loc_00E25B58: lea edx, var_50
  loc_00E25B5B: push edx
  loc_00E25B5C: push eax
  loc_00E25B5D: mov ecx, [eax]
  loc_00E25B5F: mov edi, eax
  loc_00E25B61: call [ecx+00000098h]
  loc_00E25B67: test eax, eax
  loc_00E25B69: fnclex
  loc_00E25B6B: jge 00E25B7Bh
  loc_00E25B6D: push 00000098h
  loc_00E25B72: push 006DCAF0h
  loc_00E25B77: push edi
  loc_00E25B78: push eax
  loc_00E25B79: call ebx
  loc_00E25B7B: fld real4 ptr var_4C
  loc_00E25B7E: fcomp real4 ptr var_50
  loc_00E25B81: fnstsw ax
  loc_00E25B83: test ah, 41h
  loc_00E25B86: jz 00E25B8Fh
  loc_00E25B88: mov eax, 00000001h
  loc_00E25B8D: jmp 00E25B91h
  loc_00E25B8F: xor eax, eax
  loc_00E25B91: neg eax
  loc_00E25B93: lea ecx, var_18
  loc_00E25B96: mov edi, eax
  loc_00E25B98: call [00401254h] ; __vbaFreeObj
  loc_00E25B9E: test di, di
  loc_00E25BA1: jnz 00E259F1h
  loc_00E25BA7: mov var_4, 00000000h
  loc_00E25BAE: fwait
  loc_00E25BAF: push 00E25BC1h
  loc_00E25BB4: jmp 00E25BC0h
  loc_00E25BB6: lea ecx, var_18
  loc_00E25BB9: call [00401254h] ; __vbaFreeObj
  loc_00E25BBF: ret
  loc_00E25BC0: ret
  loc_00E25BC1: mov eax, Me
  loc_00E25BC4: push eax
  loc_00E25BC5: mov ecx, [eax]
  loc_00E25BC7: call [ecx+00000008h]
  loc_00E25BCA: mov eax, var_4
  loc_00E25BCD: mov ecx, var_14
  loc_00E25BD0: pop edi
  loc_00E25BD1: pop esi
  loc_00E25BD2: mov fs:[00000000h], ecx
  loc_00E25BD9: pop ebx
  loc_00E25BDA: mov esp, ebp
  loc_00E25BDC: pop ebp
  loc_00E25BDD: retn 0008h
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer) 'E2DD00
  loc_00E2DD00: push ebp
  loc_00E2DD01: mov ebp, esp
  loc_00E2DD03: sub esp, 0000000Ch
  loc_00E2DD06: push 00402836h ; __vbaExceptHandler
  loc_00E2DD0B: mov eax, fs:[00000000h]
  loc_00E2DD11: push eax
  loc_00E2DD12: mov fs:[00000000h], esp
  loc_00E2DD19: sub esp, 000000B8h
  loc_00E2DD1F: push ebx
  loc_00E2DD20: push esi
  loc_00E2DD21: push edi
  loc_00E2DD22: mov var_C, esp
  loc_00E2DD25: mov var_8, 00402580h
  loc_00E2DD2C: mov esi, Me
  loc_00E2DD2F: mov eax, esi
  loc_00E2DD31: and eax, 00000001h
  loc_00E2DD34: mov var_4, eax
  loc_00E2DD37: and esi, FFFFFFFEh
  loc_00E2DD3A: push esi
  loc_00E2DD3B: mov Me, esi
  loc_00E2DD3E: mov ecx, [esi]
  loc_00E2DD40: call [ecx+00000004h]
  loc_00E2DD43: mov edx, [esi]
  loc_00E2DD45: xor ebx, ebx
  loc_00E2DD47: push esi
  loc_00E2DD48: mov var_24, ebx
  loc_00E2DD4B: mov var_28, ebx
  loc_00E2DD4E: mov var_2C, ebx
  loc_00E2DD51: mov var_30, ebx
  loc_00E2DD54: mov var_34, ebx
  loc_00E2DD57: mov var_44, ebx
  loc_00E2DD5A: mov var_54, ebx
  loc_00E2DD5D: mov var_64, ebx
  loc_00E2DD60: mov var_74, ebx
  loc_00E2DD63: mov var_84, ebx
  loc_00E2DD69: mov var_94, ebx
  loc_00E2DD6F: mov var_B8, ebx
  loc_00E2DD75: call [edx+00000310h]
  loc_00E2DD7B: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E2DD81: push eax
  loc_00E2DD82: lea eax, var_30
  loc_00E2DD85: push eax
  loc_00E2DD86: call edi
  loc_00E2DD88: mov ecx, [eax]
  loc_00E2DD8A: lea edx, var_B8
  loc_00E2DD90: push edx
  loc_00E2DD91: push eax
  loc_00E2DD92: mov var_BC, eax
  loc_00E2DD98: call [ecx+000000E0h]
  loc_00E2DD9E: cmp eax, ebx
  loc_00E2DDA0: fnclex
  loc_00E2DDA2: jge 00E2DDBCh
  loc_00E2DDA4: mov ecx, var_BC
  loc_00E2DDAA: push 000000E0h
  loc_00E2DDAF: push 006E03D4h
  loc_00E2DDB4: push ecx
  loc_00E2DDB5: push eax
  loc_00E2DDB6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DDBC: mov edx, var_B8
  loc_00E2DDC2: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E2DDC8: lea ecx, var_30
  loc_00E2DDCB: mov var_C4, edx
  loc_00E2DDD1: call ebx
  loc_00E2DDD3: cmp var_C4, 0000h
  loc_00E2DDDB: jz 00E2DF65h
  loc_00E2DDE1: mov eax, [esi]
  loc_00E2DDE3: push esi
  loc_00E2DDE4: call [eax+0000030Ch]
  loc_00E2DDEA: lea ecx, var_30
  loc_00E2DDED: push eax
  loc_00E2DDEE: push ecx
  loc_00E2DDEF: call edi
  loc_00E2DDF1: mov edx, [eax]
  loc_00E2DDF3: lea ecx, var_28
  loc_00E2DDF6: push ecx
  loc_00E2DDF7: push eax
  loc_00E2DDF8: mov var_BC, eax
  loc_00E2DDFE: call [edx+000000A0h]
  loc_00E2DE04: test eax, eax
  loc_00E2DE06: fnclex
  loc_00E2DE08: jge 00E2DE22h
  loc_00E2DE0A: mov edx, var_BC
  loc_00E2DE10: push 000000A0h
  loc_00E2DE15: push 006DCB70h
  loc_00E2DE1A: push edx
  loc_00E2DE1B: push eax
  loc_00E2DE1C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DE22: mov eax, var_28
  loc_00E2DE25: push 006E1958h ; "select * from lc where no_daftar like '" & Chr(37)
  loc_00E2DE2A: push eax
  loc_00E2DE2B: call [0040106Ch] ; __vbaStrCat
  loc_00E2DE31: mov edx, eax
  loc_00E2DE33: lea ecx, var_2C
  loc_00E2DE36: call [00401228h] ; __vbaStrMove
  loc_00E2DE3C: push eax
  loc_00E2DE3D: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E2DE42: call [0040106Ch] ; __vbaStrCat
  loc_00E2DE48: lea edx, var_44
  loc_00E2DE4B: lea ecx, var_24
  loc_00E2DE4E: mov var_3C, eax
  loc_00E2DE51: mov var_44, 00000008h
  loc_00E2DE58: call [0040101Ch] ; __vbaVarMove
  loc_00E2DE5E: lea ecx, var_2C
  loc_00E2DE61: lea edx, var_28
  loc_00E2DE64: push ecx
  loc_00E2DE65: push edx
  loc_00E2DE66: push 00000002h
  loc_00E2DE68: call [004011BCh] ; __vbaFreeStrList
  loc_00E2DE6E: add esp, 0000000Ch
  loc_00E2DE71: lea ecx, var_30
  loc_00E2DE74: call ebx
  loc_00E2DE76: lea eax, var_24
  loc_00E2DE79: push eax
  loc_00E2DE7A: call [00401230h] ; __vbaStrVarCopy
  loc_00E2DE80: sub esp, 00000010h
  loc_00E2DE83: mov ecx, 00000008h
  loc_00E2DE88: mov edx, esp
  loc_00E2DE8A: mov var_44, ecx
  loc_00E2DE8D: mov var_3C, eax
  loc_00E2DE90: push 0000000Eh
  loc_00E2DE92: mov [edx], ecx
  loc_00E2DE94: mov ecx, var_40
  loc_00E2DE97: push esi
  loc_00E2DE98: mov [edx+00000004h], ecx
  loc_00E2DE9B: mov ecx, [esi]
  loc_00E2DE9D: mov [edx+00000008h], eax
  loc_00E2DEA0: mov eax, var_38
  loc_00E2DEA3: mov [edx+0000000Ch], eax
  loc_00E2DEA6: call [ecx+000004A8h]
  loc_00E2DEAC: lea edx, var_30
  loc_00E2DEAF: push eax
  loc_00E2DEB0: push edx
  loc_00E2DEB1: call edi
  loc_00E2DEB3: push eax
  loc_00E2DEB4: call [00401238h] ; __vbaLateIdSt
  loc_00E2DEBA: lea ecx, var_30
  loc_00E2DEBD: call ebx
  loc_00E2DEBF: lea ecx, var_44
  loc_00E2DEC2: call [00401024h] ; __vbaFreeVar
  loc_00E2DEC8: mov eax, [esi]
  loc_00E2DECA: push 00000000h
  loc_00E2DECC: push 00000019h
  loc_00E2DECE: push esi
  loc_00E2DECF: call [eax+000004A8h]
  loc_00E2DED5: lea ecx, var_30
  loc_00E2DED8: push eax
  loc_00E2DED9: push ecx
  loc_00E2DEDA: call edi
  loc_00E2DEDC: push eax
  loc_00E2DEDD: call [00401030h] ; __vbaLateIdCall
  loc_00E2DEE3: add esp, 0000000Ch
  loc_00E2DEE6: lea ecx, var_30
  loc_00E2DEE9: call ebx
  loc_00E2DEEB: mov edx, [esi]
  loc_00E2DEED: push 006DCBD8h
  loc_00E2DEF2: push 00000000h
  loc_00E2DEF4: push 00000018h
  loc_00E2DEF6: push esi
  loc_00E2DEF7: call [edx+000004A8h]
  loc_00E2DEFD: push eax
  loc_00E2DEFE: lea eax, var_30
  loc_00E2DF01: push eax
  loc_00E2DF02: call edi
  loc_00E2DF04: lea ecx, var_44
  loc_00E2DF07: push eax
  loc_00E2DF08: push ecx
  loc_00E2DF09: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2DF0F: add esp, 00000010h
  loc_00E2DF12: push eax
  loc_00E2DF13: call [00401120h] ; __vbaCastObjVar
  loc_00E2DF19: lea edx, var_34
  loc_00E2DF1C: push eax
  loc_00E2DF1D: push edx
  loc_00E2DF1E: call edi
  loc_00E2DF20: mov esi, eax
  loc_00E2DF22: lea ecx, var_B8
  loc_00E2DF28: push ecx
  loc_00E2DF29: push esi
  loc_00E2DF2A: mov eax, [esi]
  loc_00E2DF2C: call [eax+00000050h]
  loc_00E2DF2F: test eax, eax
  loc_00E2DF31: fnclex
  loc_00E2DF33: jge 00E2DF44h
  loc_00E2DF35: push 00000050h
  loc_00E2DF37: push 006DCBE8h
  loc_00E2DF3C: push esi
  loc_00E2DF3D: push eax
  loc_00E2DF3E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DF44: lea edx, var_34
  loc_00E2DF47: lea eax, var_30
  loc_00E2DF4A: push edx
  loc_00E2DF4B: push eax
  loc_00E2DF4C: push 00000002h
  loc_00E2DF4E: call [00401048h] ; __vbaFreeObjList
  loc_00E2DF54: add esp, 0000000Ch
  loc_00E2DF57: lea ecx, var_44
  loc_00E2DF5A: call [00401024h] ; __vbaFreeVar
  loc_00E2DF60: jmp 00E2E27Eh
  loc_00E2DF65: mov ecx, [esi]
  loc_00E2DF67: push esi
  loc_00E2DF68: call [ecx+00000314h]
  loc_00E2DF6E: lea edx, var_30
  loc_00E2DF71: push eax
  loc_00E2DF72: push edx
  loc_00E2DF73: call edi
  loc_00E2DF75: mov ecx, [eax]
  loc_00E2DF77: lea edx, var_B8
  loc_00E2DF7D: push edx
  loc_00E2DF7E: push eax
  loc_00E2DF7F: mov var_BC, eax
  loc_00E2DF85: call [ecx+000000E0h]
  loc_00E2DF8B: test eax, eax
  loc_00E2DF8D: fnclex
  loc_00E2DF8F: jge 00E2DFA9h
  loc_00E2DF91: mov ecx, var_BC
  loc_00E2DF97: push 000000E0h
  loc_00E2DF9C: push 006E03D4h
  loc_00E2DFA1: push ecx
  loc_00E2DFA2: push eax
  loc_00E2DFA3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2DFA9: mov edx, var_B8
  loc_00E2DFAF: lea ecx, var_30
  loc_00E2DFB2: mov var_C4, edx
  loc_00E2DFB8: call ebx
  loc_00E2DFBA: cmp var_C4, 0000h
  loc_00E2DFC2: jz 00E2E0D7h
  loc_00E2DFC8: mov eax, [esi]
  loc_00E2DFCA: push esi
  loc_00E2DFCB: call [eax+0000030Ch]
  loc_00E2DFD1: lea ecx, var_30
  loc_00E2DFD4: push eax
  loc_00E2DFD5: push ecx
  loc_00E2DFD6: call edi
  loc_00E2DFD8: mov edx, [eax]
  loc_00E2DFDA: lea ecx, var_28
  loc_00E2DFDD: push ecx
  loc_00E2DFDE: push eax
  loc_00E2DFDF: mov var_BC, eax
  loc_00E2DFE5: call [edx+000000A0h]
  loc_00E2DFEB: test eax, eax
  loc_00E2DFED: fnclex
  loc_00E2DFEF: jge 00E2E009h
  loc_00E2DFF1: mov edx, var_BC
  loc_00E2DFF7: push 000000A0h
  loc_00E2DFFC: push 006DCB70h
  loc_00E2E001: push edx
  loc_00E2E002: push eax
  loc_00E2E003: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E009: mov eax, var_28
  loc_00E2E00C: push 006E19DCh ; "select * from lc where nama_siswa like '" & Chr(37)
  loc_00E2E011: push eax
  loc_00E2E012: call [0040106Ch] ; __vbaStrCat
  loc_00E2E018: mov edx, eax
  loc_00E2E01A: lea ecx, var_2C
  loc_00E2E01D: call [00401228h] ; __vbaStrMove
  loc_00E2E023: push eax
  loc_00E2E024: push 006E1A34h ; Chr(37) & "' order by nama_siswa asc"
  loc_00E2E029: call [0040106Ch] ; __vbaStrCat
  loc_00E2E02F: lea edx, var_44
  loc_00E2E032: lea ecx, var_24
  loc_00E2E035: mov var_3C, eax
  loc_00E2E038: mov var_44, 00000008h
  loc_00E2E03F: call [0040101Ch] ; __vbaVarMove
  loc_00E2E045: lea ecx, var_2C
  loc_00E2E048: lea edx, var_28
  loc_00E2E04B: push ecx
  loc_00E2E04C: push edx
  loc_00E2E04D: push 00000002h
  loc_00E2E04F: call [004011BCh] ; __vbaFreeStrList
  loc_00E2E055: add esp, 0000000Ch
  loc_00E2E058: lea ecx, var_30
  loc_00E2E05B: call ebx
  loc_00E2E05D: lea eax, var_24
  loc_00E2E060: push eax
  loc_00E2E061: call [00401230h] ; __vbaStrVarCopy
  loc_00E2E067: sub esp, 00000010h
  loc_00E2E06A: mov ecx, 00000008h
  loc_00E2E06F: mov edx, esp
  loc_00E2E071: mov var_44, ecx
  loc_00E2E074: mov var_3C, eax
  loc_00E2E077: push 0000000Eh
  loc_00E2E079: mov [edx], ecx
  loc_00E2E07B: mov ecx, var_40
  loc_00E2E07E: push esi
  loc_00E2E07F: mov [edx+00000004h], ecx
  loc_00E2E082: mov ecx, [esi]
  loc_00E2E084: mov [edx+00000008h], eax
  loc_00E2E087: mov eax, var_38
  loc_00E2E08A: mov [edx+0000000Ch], eax
  loc_00E2E08D: call [ecx+000004A8h]
  loc_00E2E093: lea edx, var_30
  loc_00E2E096: push eax
  loc_00E2E097: push edx
  loc_00E2E098: call edi
  loc_00E2E09A: push eax
  loc_00E2E09B: call [00401238h] ; __vbaLateIdSt
  loc_00E2E0A1: lea ecx, var_30
  loc_00E2E0A4: call ebx
  loc_00E2E0A6: lea ecx, var_44
  loc_00E2E0A9: call [00401024h] ; __vbaFreeVar
  loc_00E2E0AF: mov eax, [esi]
  loc_00E2E0B1: push 00000000h
  loc_00E2E0B3: push 00000019h
  loc_00E2E0B5: push esi
  loc_00E2E0B6: call [eax+000004A8h]
  loc_00E2E0BC: lea ecx, var_30
  loc_00E2E0BF: push eax
  loc_00E2E0C0: push ecx
  loc_00E2E0C1: call edi
  loc_00E2E0C3: push eax
  loc_00E2E0C4: call [00401030h] ; __vbaLateIdCall
  loc_00E2E0CA: add esp, 0000000Ch
  loc_00E2E0CD: lea ecx, var_30
  loc_00E2E0D0: call ebx
  loc_00E2E0D2: jmp 00E2E27Eh
  loc_00E2E0D7: mov edx, [esi]
  loc_00E2E0D9: push esi
  loc_00E2E0DA: call [edx+00000310h]
  loc_00E2E0E0: push eax
  loc_00E2E0E1: lea eax, var_30
  loc_00E2E0E4: push eax
  loc_00E2E0E5: call edi
  loc_00E2E0E7: mov ecx, [eax]
  loc_00E2E0E9: lea edx, var_B8
  loc_00E2E0EF: push edx
  loc_00E2E0F0: push eax
  loc_00E2E0F1: mov var_BC, eax
  loc_00E2E0F7: call [ecx+000000E0h]
  loc_00E2E0FD: test eax, eax
  loc_00E2E0FF: fnclex
  loc_00E2E101: jge 00E2E11Bh
  loc_00E2E103: mov ecx, var_BC
  loc_00E2E109: push 000000E0h
  loc_00E2E10E: push 006E03D4h
  loc_00E2E113: push ecx
  loc_00E2E114: push eax
  loc_00E2E115: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E11B: xor edx, edx
  loc_00E2E11D: lea ecx, var_30
  loc_00E2E120: cmp var_B8, dx
  loc_00E2E127: setz dl
  loc_00E2E12A: neg edx
  loc_00E2E12C: mov var_C4, edx
  loc_00E2E132: call ebx
  loc_00E2E134: cmp var_C4, 0000h
  loc_00E2E13C: mov ecx, 80020004h
  loc_00E2E141: mov eax, 0000000Ah
  loc_00E2E146: mov var_6C, ecx
  loc_00E2E149: mov var_74, eax
  loc_00E2E14C: mov var_5C, ecx
  loc_00E2E14F: mov var_64, eax
  loc_00E2E152: jz 00E2E20Dh
  loc_00E2E158: lea edx, var_94
  loc_00E2E15E: lea ecx, var_54
  loc_00E2E161: mov var_8C, 006E1AE8h ; "Need To Do !1"
  loc_00E2E16B: mov var_94, 00000008h
  loc_00E2E175: call [004011F0h] ; __vbaVarDup
  loc_00E2E17B: lea edx, var_84
  loc_00E2E181: lea ecx, var_44
  loc_00E2E184: mov var_7C, 006E1A70h ; "Silahkan Pilih Item Yang ingin di cari terlebih dahulu !!"
  loc_00E2E18B: mov var_84, 00000008h
  loc_00E2E195: call [004011F0h] ; __vbaVarDup
  loc_00E2E19B: lea eax, var_74
  loc_00E2E19E: lea ecx, var_64
  loc_00E2E1A1: push eax
  loc_00E2E1A2: lea edx, var_54
  loc_00E2E1A5: push ecx
  loc_00E2E1A6: push edx
  loc_00E2E1A7: lea eax, var_44
  loc_00E2E1AA: push 00000010h
  loc_00E2E1AC: push eax
  loc_00E2E1AD: call [004010A8h] ; rtcMsgBox
  loc_00E2E1B3: lea ecx, var_74
  loc_00E2E1B6: lea edx, var_64
  loc_00E2E1B9: push ecx
  loc_00E2E1BA: lea eax, var_54
  loc_00E2E1BD: push edx
  loc_00E2E1BE: lea ecx, var_44
  loc_00E2E1C1: push eax
  loc_00E2E1C2: push ecx
  loc_00E2E1C3: push 00000004h
  loc_00E2E1C5: call [00401038h] ; __vbaFreeVarList
  loc_00E2E1CB: mov edx, [esi]
  loc_00E2E1CD: add esp, 00000014h
  loc_00E2E1D0: push esi
  loc_00E2E1D1: call [edx+0000030Ch]
  loc_00E2E1D7: push eax
  loc_00E2E1D8: lea eax, var_30
  loc_00E2E1DB: push eax
  loc_00E2E1DC: call edi
  loc_00E2E1DE: mov esi, eax
  loc_00E2E1E0: push 006DCC80h
  loc_00E2E1E5: push esi
  loc_00E2E1E6: mov ecx, [esi]
  loc_00E2E1E8: call [ecx+000000A4h]
  loc_00E2E1EE: test eax, eax
  loc_00E2E1F0: fnclex
  loc_00E2E1F2: jge 00E2E206h
  loc_00E2E1F4: push 000000A4h
  loc_00E2E1F9: push 006DCB70h
  loc_00E2E1FE: push esi
  loc_00E2E1FF: push eax
  loc_00E2E200: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E206: lea ecx, var_30
  loc_00E2E209: call ebx
  loc_00E2E20B: jmp 00E2E27Eh
  loc_00E2E20D: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E2E213: mov edi, 00000008h
  loc_00E2E218: lea edx, var_94
  loc_00E2E21E: lea ecx, var_54
  loc_00E2E221: mov var_8C, 006E0D04h ; "Not Found !"
  loc_00E2E22B: mov var_94, edi
  loc_00E2E231: call __vbaVarDup
  loc_00E2E233: lea edx, var_84
  loc_00E2E239: lea ecx, var_44
  loc_00E2E23C: mov var_7C, 006E0CE0h ; "Data Tidak Ada"
  loc_00E2E243: mov var_84, edi
  loc_00E2E249: call __vbaVarDup
  loc_00E2E24B: lea edx, var_74
  loc_00E2E24E: lea eax, var_64
  loc_00E2E251: push edx
  loc_00E2E252: lea ecx, var_54
  loc_00E2E255: push eax
  loc_00E2E256: push ecx
  loc_00E2E257: lea edx, var_44
  loc_00E2E25A: push 00000010h
  loc_00E2E25C: push edx
  loc_00E2E25D: call [004010A8h] ; rtcMsgBox
  loc_00E2E263: lea eax, var_74
  loc_00E2E266: lea ecx, var_64
  loc_00E2E269: push eax
  loc_00E2E26A: lea edx, var_54
  loc_00E2E26D: push ecx
  loc_00E2E26E: lea eax, var_44
  loc_00E2E271: push edx
  loc_00E2E272: push eax
  loc_00E2E273: push 00000004h
  loc_00E2E275: call [00401038h] ; __vbaFreeVarList
  loc_00E2E27B: add esp, 00000014h
  loc_00E2E27E: mov var_4, 00000000h
  loc_00E2E285: push 00E2E2D2h
  loc_00E2E28A: jmp 00E2E2C8h
  loc_00E2E28C: lea ecx, var_2C
  loc_00E2E28F: lea edx, var_28
  loc_00E2E292: push ecx
  loc_00E2E293: push edx
  loc_00E2E294: push 00000002h
  loc_00E2E296: call [004011BCh] ; __vbaFreeStrList
  loc_00E2E29C: lea eax, var_34
  loc_00E2E29F: lea ecx, var_30
  loc_00E2E2A2: push eax
  loc_00E2E2A3: push ecx
  loc_00E2E2A4: push 00000002h
  loc_00E2E2A6: call [00401048h] ; __vbaFreeObjList
  loc_00E2E2AC: lea edx, var_74
  loc_00E2E2AF: lea eax, var_64
  loc_00E2E2B2: push edx
  loc_00E2E2B3: lea ecx, var_54
  loc_00E2E2B6: push eax
  loc_00E2E2B7: lea edx, var_44
  loc_00E2E2BA: push ecx
  loc_00E2E2BB: push edx
  loc_00E2E2BC: push 00000004h
  loc_00E2E2BE: call [00401038h] ; __vbaFreeVarList
  loc_00E2E2C4: add esp, 0000002Ch
  loc_00E2E2C7: ret
  loc_00E2E2C8: lea ecx, var_24
  loc_00E2E2CB: call [00401024h] ; __vbaFreeVar
  loc_00E2E2D1: ret
  loc_00E2E2D2: mov eax, Me
  loc_00E2E2D5: push eax
  loc_00E2E2D6: mov ecx, [eax]
  loc_00E2E2D8: call [ecx+00000008h]
  loc_00E2E2DB: mov eax, var_4
  loc_00E2E2DE: mov ecx, var_14
  loc_00E2E2E1: pop edi
  loc_00E2E2E2: pop esi
  loc_00E2E2E3: mov fs:[00000000h], ecx
  loc_00E2E2EA: pop ebx
  loc_00E2E2EB: mov esp, ebp
  loc_00E2E2ED: pop ebp
  loc_00E2E2EE: retn 0008h
End Sub

Private Sub daftar_KeyPress(KeyAscii As Integer) 'E25230
  loc_00E25230: push ebp
  loc_00E25231: mov ebp, esp
  loc_00E25233: sub esp, 0000000Ch
  loc_00E25236: push 00402836h ; __vbaExceptHandler
  loc_00E2523B: mov eax, fs:[00000000h]
  loc_00E25241: push eax
  loc_00E25242: mov fs:[00000000h], esp
  loc_00E25249: sub esp, 00000020h
  loc_00E2524C: push ebx
  loc_00E2524D: push esi
  loc_00E2524E: push edi
  loc_00E2524F: mov var_C, esp
  loc_00E25252: mov var_8, 004024B0h
  loc_00E25259: mov edi, Me
  loc_00E2525C: mov eax, edi
  loc_00E2525E: and eax, 00000001h
  loc_00E25261: mov var_4, eax
  loc_00E25264: and edi, FFFFFFFEh
  loc_00E25267: push edi
  loc_00E25268: mov Me, edi
  loc_00E2526B: mov ecx, [edi]
  loc_00E2526D: call [ecx+00000004h]
  loc_00E25270: mov edx, KeyAscii
  loc_00E25273: xor ebx, ebx
  loc_00E25275: mov var_18, ebx
  loc_00E25278: mov var_1C, ebx
  loc_00E2527B: cmp [edx], 000Dh
  loc_00E2527F: jnz 00E25342h
  loc_00E25285: mov eax, [edi]
  loc_00E25287: push edi
  loc_00E25288: call [eax+00000304h]
  loc_00E2528E: lea ecx, var_1C
  loc_00E25291: push eax
  loc_00E25292: push ecx
  loc_00E25293: call [004010ACh] ; __vbaObjSet
  loc_00E25299: mov esi, eax
  loc_00E2529B: lea eax, var_18
  loc_00E2529E: push eax
  loc_00E2529F: push esi
  loc_00E252A0: mov edx, [esi]
  loc_00E252A2: call [edx+000000A0h]
  loc_00E252A8: cmp eax, ebx
  loc_00E252AA: fnclex
  loc_00E252AC: jge 00E252C0h
  loc_00E252AE: push 000000A0h
  loc_00E252B3: push 006DCB70h
  loc_00E252B8: push esi
  loc_00E252B9: push eax
  loc_00E252BA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E252C0: mov ecx, var_18
  loc_00E252C3: push ecx
  loc_00E252C4: push 006E0708h ; "Keluar"
  loc_00E252C9: call [00401104h] ; __vbaStrCmp
  loc_00E252CF: mov esi, eax
  loc_00E252D1: lea ecx, var_18
  loc_00E252D4: neg esi
  loc_00E252D6: sbb esi, esi
  loc_00E252D8: inc esi
  loc_00E252D9: neg esi
  loc_00E252DB: call [00401258h] ; __vbaFreeStr
  loc_00E252E1: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E252E7: lea ecx, var_1C
  loc_00E252EA: call ebx
  loc_00E252EC: test si, si
  loc_00E252EF: jz 00E25342h
  loc_00E252F1: mov eax, [00E3D9CCh]
  loc_00E252F6: test eax, eax
  loc_00E252F8: jnz 00E2530Ah
  loc_00E252FA: push 00E3D9CCh
  loc_00E252FF: push 006DC960h
  loc_00E25304: call [004011A0h] ; __vbaNew2
  loc_00E2530A: mov esi, [00E3D9CCh]
  loc_00E25310: lea eax, var_1C
  loc_00E25313: push edi
  loc_00E25314: push eax
  loc_00E25315: mov edx, [esi]
  loc_00E25317: mov var_34, edx
  loc_00E2531A: call [004010B4h] ; __vbaObjSetAddref
  loc_00E25320: mov ecx, var_34
  loc_00E25323: push eax
  loc_00E25324: push esi
  loc_00E25325: call [ecx+00000010h]
  loc_00E25328: test eax, eax
  loc_00E2532A: fnclex
  loc_00E2532C: jge 00E2533Dh
  loc_00E2532E: push 00000010h
  loc_00E25330: push 006DC950h
  loc_00E25335: push esi
  loc_00E25336: push eax
  loc_00E25337: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2533D: lea ecx, var_1C
  loc_00E25340: call ebx
  loc_00E25342: mov var_4, 00000000h
  loc_00E25349: push 00E25364h
  loc_00E2534E: jmp 00E25363h
  loc_00E25350: lea ecx, var_18
  loc_00E25353: call [00401258h] ; __vbaFreeStr
  loc_00E25359: lea ecx, var_1C
  loc_00E2535C: call [00401254h] ; __vbaFreeObj
  loc_00E25362: ret
  loc_00E25363: ret
  loc_00E25364: mov eax, Me
  loc_00E25367: push eax
  loc_00E25368: mov edx, [eax]
  loc_00E2536A: call [edx+00000008h]
  loc_00E2536D: mov eax, var_4
  loc_00E25370: mov ecx, var_14
  loc_00E25373: pop edi
  loc_00E25374: pop esi
  loc_00E25375: mov fs:[00000000h], ecx
  loc_00E2537C: pop ebx
  loc_00E2537D: mov esp, ebp
  loc_00E2537F: pop ebp
  loc_00E25380: retn 0008h
End Sub

Private Sub update_UnknownEvent_9 'E2E300
  loc_00E2E300: push ebp
  loc_00E2E301: mov ebp, esp
  loc_00E2E303: sub esp, 0000000Ch
  loc_00E2E306: push 00402836h ; __vbaExceptHandler
  loc_00E2E30B: mov eax, fs:[00000000h]
  loc_00E2E311: push eax
  loc_00E2E312: mov fs:[00000000h], esp
  loc_00E2E319: sub esp, 0000013Ch
  loc_00E2E31F: push ebx
  loc_00E2E320: push esi
  loc_00E2E321: push edi
  loc_00E2E322: mov var_C, esp
  loc_00E2E325: mov var_8, 00402590h
  loc_00E2E32C: mov esi, Me
  loc_00E2E32F: mov eax, esi
  loc_00E2E331: and eax, 00000001h
  loc_00E2E334: mov var_4, eax
  loc_00E2E337: and esi, FFFFFFFEh
  loc_00E2E33A: push esi
  loc_00E2E33B: mov Me, esi
  loc_00E2E33E: mov ecx, [esi]
  loc_00E2E340: call [ecx+00000004h]
  loc_00E2E343: mov edx, [esi]
  loc_00E2E345: xor ebx, ebx
  loc_00E2E347: push esi
  loc_00E2E348: mov var_24, ebx
  loc_00E2E34B: mov var_28, ebx
  loc_00E2E34E: mov var_2C, ebx
  loc_00E2E351: mov var_30, ebx
  loc_00E2E354: mov var_34, ebx
  loc_00E2E357: mov var_38, ebx
  loc_00E2E35A: mov var_3C, ebx
  loc_00E2E35D: mov var_40, ebx
  loc_00E2E360: mov var_44, ebx
  loc_00E2E363: mov var_48, ebx
  loc_00E2E366: mov var_4C, ebx
  loc_00E2E369: mov var_5C, ebx
  loc_00E2E36C: mov var_6C, ebx
  loc_00E2E36F: mov var_7C, ebx
  loc_00E2E372: mov var_8C, ebx
  loc_00E2E378: mov var_9C, ebx
  loc_00E2E37E: mov var_D0, ebx
  loc_00E2E384: call [edx+00000434h]
  loc_00E2E38A: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E2E390: push eax
  loc_00E2E391: lea eax, var_34
  loc_00E2E394: push eax
  loc_00E2E395: call edi
  loc_00E2E397: mov ecx, [eax]
  loc_00E2E399: lea edx, var_28
  loc_00E2E39C: push edx
  loc_00E2E39D: push eax
  loc_00E2E39E: mov var_E4, eax
  loc_00E2E3A4: call [ecx+00000050h]
  loc_00E2E3A7: cmp eax, ebx
  loc_00E2E3A9: fnclex
  loc_00E2E3AB: jge 00E2E3C2h
  loc_00E2E3AD: mov ecx, var_E4
  loc_00E2E3B3: push 00000050h
  loc_00E2E3B5: push 006E0410h
  loc_00E2E3BA: push ecx
  loc_00E2E3BB: push eax
  loc_00E2E3BC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E3C2: mov edx, var_28
  loc_00E2E3C5: mov ebx, [0040106Ch] ; __vbaStrCat
  loc_00E2E3CB: push 006E1B08h ; "select * from lc where nama_siswa ='"
  loc_00E2E3D0: push edx
  loc_00E2E3D1: call ebx
  loc_00E2E3D3: mov edx, eax
  loc_00E2E3D5: lea ecx, var_2C
  loc_00E2E3D8: call [00401228h] ; __vbaStrMove
  loc_00E2E3DE: push eax
  loc_00E2E3DF: push 006DCB84h ; "'"
  loc_00E2E3E4: call ebx
  loc_00E2E3E6: lea edx, var_5C
  loc_00E2E3E9: lea ecx, var_24
  loc_00E2E3EC: mov var_54, eax
  loc_00E2E3EF: mov var_5C, 00000008h
  loc_00E2E3F6: call [0040101Ch] ; __vbaVarMove
  loc_00E2E3FC: lea eax, var_2C
  loc_00E2E3FF: lea ecx, var_28
  loc_00E2E402: push eax
  loc_00E2E403: push ecx
  loc_00E2E404: push 00000002h
  loc_00E2E406: call [004011BCh] ; __vbaFreeStrList
  loc_00E2E40C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E2E412: add esp, 0000000Ch
  loc_00E2E415: lea ecx, var_34
  loc_00E2E418: call ebx
  loc_00E2E41A: lea edx, var_24
  loc_00E2E41D: push edx
  loc_00E2E41E: call [00401230h] ; __vbaStrVarCopy
  loc_00E2E424: sub esp, 00000010h
  loc_00E2E427: mov ecx, 00000008h
  loc_00E2E42C: mov edx, esp
  loc_00E2E42E: mov var_5C, ecx
  loc_00E2E431: mov var_54, eax
  loc_00E2E434: push 0000000Eh
  loc_00E2E436: mov [edx], ecx
  loc_00E2E438: mov ecx, var_58
  loc_00E2E43B: push esi
  loc_00E2E43C: mov [edx+00000004h], ecx
  loc_00E2E43F: mov ecx, [esi]
  loc_00E2E441: mov [edx+00000008h], eax
  loc_00E2E444: mov eax, var_50
  loc_00E2E447: mov [edx+0000000Ch], eax
  loc_00E2E44A: call [ecx+000004A8h]
  loc_00E2E450: lea edx, var_34
  loc_00E2E453: push eax
  loc_00E2E454: push edx
  loc_00E2E455: call edi
  loc_00E2E457: push eax
  loc_00E2E458: call [00401238h] ; __vbaLateIdSt
  loc_00E2E45E: lea ecx, var_34
  loc_00E2E461: call ebx
  loc_00E2E463: lea ecx, var_5C
  loc_00E2E466: call [00401024h] ; __vbaFreeVar
  loc_00E2E46C: mov eax, [esi]
  loc_00E2E46E: push 006DCBD8h
  loc_00E2E473: push 00000000h
  loc_00E2E475: push 00000018h
  loc_00E2E477: push esi
  loc_00E2E478: call [eax+000004A8h]
  loc_00E2E47E: lea ecx, var_34
  loc_00E2E481: push eax
  loc_00E2E482: push ecx
  loc_00E2E483: call edi
  loc_00E2E485: lea edx, var_5C
  loc_00E2E488: push eax
  loc_00E2E489: push edx
  loc_00E2E48A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2E490: add esp, 00000010h
  loc_00E2E493: push eax
  loc_00E2E494: call [00401120h] ; __vbaCastObjVar
  loc_00E2E49A: push eax
  loc_00E2E49B: lea eax, var_38
  loc_00E2E49E: push eax
  loc_00E2E49F: call edi
  loc_00E2E4A1: mov ebx, eax
  loc_00E2E4A3: mov eax, 0000000Ah
  loc_00E2E4A8: mov var_94, 80020004h
  loc_00E2E4B2: mov var_9C, eax
  loc_00E2E4B8: mov ecx, [ebx]
  loc_00E2E4BA: sub esp, 00000010h
  loc_00E2E4BD: mov edx, esp
  loc_00E2E4BF: sub esp, 00000010h
  loc_00E2E4C2: mov [edx], eax
  loc_00E2E4C4: mov eax, var_A8
  loc_00E2E4CA: mov [edx+00000004h], eax
  loc_00E2E4CD: mov eax, 80020004h
  loc_00E2E4D2: mov [edx+00000008h], eax
  loc_00E2E4D5: mov eax, var_A0
  loc_00E2E4DB: mov [edx+0000000Ch], eax
  loc_00E2E4DE: mov eax, var_9C
  loc_00E2E4E4: mov edx, esp
  loc_00E2E4E6: push ebx
  loc_00E2E4E7: mov [edx], eax
  loc_00E2E4E9: mov eax, var_98
  loc_00E2E4EF: mov [edx+00000004h], eax
  loc_00E2E4F2: mov eax, var_94
  loc_00E2E4F8: mov [edx+00000008h], eax
  loc_00E2E4FB: mov eax, var_90
  loc_00E2E501: mov [edx+0000000Ch], eax
  loc_00E2E504: call [ecx+000000ACh]
  loc_00E2E50A: test eax, eax
  loc_00E2E50C: fnclex
  loc_00E2E50E: jge 00E2E522h
  loc_00E2E510: push 000000ACh
  loc_00E2E515: push 006DCBE8h
  loc_00E2E51A: push ebx
  loc_00E2E51B: push eax
  loc_00E2E51C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E522: lea ecx, var_38
  loc_00E2E525: lea edx, var_34
  loc_00E2E528: push ecx
  loc_00E2E529: push edx
  loc_00E2E52A: push 00000002h
  loc_00E2E52C: call [00401048h] ; __vbaFreeObjList
  loc_00E2E532: add esp, 0000000Ch
  loc_00E2E535: lea ecx, var_5C
  loc_00E2E538: call [00401024h] ; __vbaFreeVar
  loc_00E2E53E: mov eax, [esi]
  loc_00E2E540: push 006DCBD8h
  loc_00E2E545: push 00000000h
  loc_00E2E547: push 00000018h
  loc_00E2E549: push esi
  loc_00E2E54A: call [eax+000004A8h]
  loc_00E2E550: lea ecx, var_38
  loc_00E2E553: push eax
  loc_00E2E554: push ecx
  loc_00E2E555: call edi
  loc_00E2E557: lea edx, var_5C
  loc_00E2E55A: push eax
  loc_00E2E55B: push edx
  loc_00E2E55C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2E562: add esp, 00000010h
  loc_00E2E565: push eax
  loc_00E2E566: call [00401120h] ; __vbaCastObjVar
  loc_00E2E56C: push eax
  loc_00E2E56D: lea eax, var_3C
  loc_00E2E570: push eax
  loc_00E2E571: call edi
  loc_00E2E573: mov ebx, eax
  loc_00E2E575: lea edx, var_40
  loc_00E2E578: push edx
  loc_00E2E579: push ebx
  loc_00E2E57A: mov ecx, [ebx]
  loc_00E2E57C: call [ecx+00000054h]
  loc_00E2E57F: test eax, eax
  loc_00E2E581: fnclex
  loc_00E2E583: jge 00E2E594h
  loc_00E2E585: push 00000054h
  loc_00E2E587: push 006DCBE8h
  loc_00E2E58C: push ebx
  loc_00E2E58D: push eax
  loc_00E2E58E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E594: lea ebx, var_44
  loc_00E2E597: mov eax, var_40
  loc_00E2E59A: push ebx
  loc_00E2E59B: mov ecx, 00000002h
  loc_00E2E5A0: sub esp, 00000010h
  loc_00E2E5A3: mov var_9C, ecx
  loc_00E2E5A9: mov ebx, esp
  loc_00E2E5AB: mov var_94, 00000000h
  loc_00E2E5B5: mov edx, [eax]
  loc_00E2E5B7: push eax
  loc_00E2E5B8: mov [ebx], ecx
  loc_00E2E5BA: mov ecx, var_98
  loc_00E2E5C0: mov var_F4, eax
  loc_00E2E5C6: mov [ebx+00000004h], ecx
  loc_00E2E5C9: mov ecx, var_94
  loc_00E2E5CF: mov [ebx+00000008h], ecx
  loc_00E2E5D2: mov ecx, var_90
  loc_00E2E5D8: mov [ebx+0000000Ch], ecx
  loc_00E2E5DB: call [edx+00000028h]
  loc_00E2E5DE: test eax, eax
  loc_00E2E5E0: fnclex
  loc_00E2E5E2: jge 00E2E5F9h
  loc_00E2E5E4: mov edx, var_F4
  loc_00E2E5EA: push 00000028h
  loc_00E2E5EC: push 006E09E8h
  loc_00E2E5F1: push edx
  loc_00E2E5F2: push eax
  loc_00E2E5F3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E5F9: mov eax, var_44
  loc_00E2E5FC: mov ecx, [esi]
  loc_00E2E5FE: push esi
  loc_00E2E5FF: mov var_FC, eax
  loc_00E2E605: call [ecx+00000430h]
  loc_00E2E60B: lea edx, var_34
  loc_00E2E60E: push eax
  loc_00E2E60F: push edx
  loc_00E2E610: call edi
  loc_00E2E612: mov ebx, eax
  loc_00E2E614: lea ecx, var_28
  loc_00E2E617: push ecx
  loc_00E2E618: push ebx
  loc_00E2E619: mov eax, [ebx]
  loc_00E2E61B: call [eax+00000050h]
  loc_00E2E61E: test eax, eax
  loc_00E2E620: fnclex
  loc_00E2E622: jge 00E2E633h
  loc_00E2E624: push 00000050h
  loc_00E2E626: push 006E0410h
  loc_00E2E62B: push ebx
  loc_00E2E62C: push eax
  loc_00E2E62D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E633: sub esp, 00000010h
  loc_00E2E636: mov eax, var_28
  loc_00E2E639: mov ebx, esp
  loc_00E2E63B: mov ecx, 00000008h
  loc_00E2E640: mov var_6C, ecx
  loc_00E2E643: mov edx, var_FC
  loc_00E2E649: mov [ebx], ecx
  loc_00E2E64B: mov ecx, var_68
  loc_00E2E64E: mov var_64, eax
  loc_00E2E651: mov var_28, 00000000h
  loc_00E2E658: mov edx, [edx]
  loc_00E2E65A: mov [ebx+00000004h], ecx
  loc_00E2E65D: mov [ebx+00000008h], eax
  loc_00E2E660: mov eax, var_60
  loc_00E2E663: mov [ebx+0000000Ch], eax
  loc_00E2E666: mov ebx, var_FC
  loc_00E2E66C: push ebx
  loc_00E2E66D: call [edx+00000038h]
  loc_00E2E670: test eax, eax
  loc_00E2E672: fnclex
  loc_00E2E674: jge 00E2E685h
  loc_00E2E676: push 00000038h
  loc_00E2E678: push 006E09F8h
  loc_00E2E67D: push ebx
  loc_00E2E67E: push eax
  loc_00E2E67F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E685: lea ecx, var_44
  loc_00E2E688: lea edx, var_40
  loc_00E2E68B: push ecx
  loc_00E2E68C: lea eax, var_3C
  loc_00E2E68F: push edx
  loc_00E2E690: lea ecx, var_38
  loc_00E2E693: push eax
  loc_00E2E694: lea edx, var_34
  loc_00E2E697: push ecx
  loc_00E2E698: push edx
  loc_00E2E699: push 00000005h
  loc_00E2E69B: call [00401048h] ; __vbaFreeObjList
  loc_00E2E6A1: lea eax, var_6C
  loc_00E2E6A4: lea ecx, var_5C
  loc_00E2E6A7: push eax
  loc_00E2E6A8: push ecx
  loc_00E2E6A9: push 00000002h
  loc_00E2E6AB: call [00401038h] ; __vbaFreeVarList
  loc_00E2E6B1: mov edx, [esi]
  loc_00E2E6B3: add esp, 00000024h
  loc_00E2E6B6: push 006DCBD8h
  loc_00E2E6BB: push 00000000h
  loc_00E2E6BD: push 00000018h
  loc_00E2E6BF: push esi
  loc_00E2E6C0: call [edx+000004A8h]
  loc_00E2E6C6: push eax
  loc_00E2E6C7: lea eax, var_38
  loc_00E2E6CA: push eax
  loc_00E2E6CB: call edi
  loc_00E2E6CD: lea ecx, var_5C
  loc_00E2E6D0: push eax
  loc_00E2E6D1: push ecx
  loc_00E2E6D2: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2E6D8: add esp, 00000010h
  loc_00E2E6DB: push eax
  loc_00E2E6DC: call [00401120h] ; __vbaCastObjVar
  loc_00E2E6E2: lea edx, var_3C
  loc_00E2E6E5: push eax
  loc_00E2E6E6: push edx
  loc_00E2E6E7: call edi
  loc_00E2E6E9: mov ebx, eax
  loc_00E2E6EB: lea ecx, var_40
  loc_00E2E6EE: push ecx
  loc_00E2E6EF: push ebx
  loc_00E2E6F0: mov eax, [ebx]
  loc_00E2E6F2: call [eax+00000054h]
  loc_00E2E6F5: test eax, eax
  loc_00E2E6F7: fnclex
  loc_00E2E6F9: jge 00E2E70Ah
  loc_00E2E6FB: push 00000054h
  loc_00E2E6FD: push 006DCBE8h
  loc_00E2E702: push ebx
  loc_00E2E703: push eax
  loc_00E2E704: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E70A: lea ebx, var_44
  loc_00E2E70D: mov eax, var_40
  loc_00E2E710: push ebx
  loc_00E2E711: mov ecx, 00000002h
  loc_00E2E716: sub esp, 00000010h
  loc_00E2E719: mov var_9C, ecx
  loc_00E2E71F: mov ebx, esp
  loc_00E2E721: mov var_94, 00000001h
  loc_00E2E72B: mov edx, [eax]
  loc_00E2E72D: push eax
  loc_00E2E72E: mov [ebx], ecx
  loc_00E2E730: mov ecx, var_98
  loc_00E2E736: mov var_F4, eax
  loc_00E2E73C: mov [ebx+00000004h], ecx
  loc_00E2E73F: mov ecx, var_94
  loc_00E2E745: mov [ebx+00000008h], ecx
  loc_00E2E748: mov ecx, var_90
  loc_00E2E74E: mov [ebx+0000000Ch], ecx
  loc_00E2E751: call [edx+00000028h]
  loc_00E2E754: test eax, eax
  loc_00E2E756: fnclex
  loc_00E2E758: jge 00E2E76Fh
  loc_00E2E75A: mov edx, var_F4
  loc_00E2E760: push 00000028h
  loc_00E2E762: push 006E09E8h
  loc_00E2E767: push edx
  loc_00E2E768: push eax
  loc_00E2E769: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E76F: mov eax, var_44
  loc_00E2E772: mov ecx, [esi]
  loc_00E2E774: push esi
  loc_00E2E775: mov var_FC, eax
  loc_00E2E77B: call [ecx+00000434h]
  loc_00E2E781: lea edx, var_34
  loc_00E2E784: push eax
  loc_00E2E785: push edx
  loc_00E2E786: call edi
  loc_00E2E788: mov ebx, eax
  loc_00E2E78A: lea ecx, var_28
  loc_00E2E78D: push ecx
  loc_00E2E78E: push ebx
  loc_00E2E78F: mov eax, [ebx]
  loc_00E2E791: call [eax+00000050h]
  loc_00E2E794: test eax, eax
  loc_00E2E796: fnclex
  loc_00E2E798: jge 00E2E7A9h
  loc_00E2E79A: push 00000050h
  loc_00E2E79C: push 006E0410h
  loc_00E2E7A1: push ebx
  loc_00E2E7A2: push eax
  loc_00E2E7A3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E7A9: sub esp, 00000010h
  loc_00E2E7AC: mov eax, var_28
  loc_00E2E7AF: mov ebx, esp
  loc_00E2E7B1: mov ecx, 00000008h
  loc_00E2E7B6: mov var_6C, ecx
  loc_00E2E7B9: mov edx, var_FC
  loc_00E2E7BF: mov [ebx], ecx
  loc_00E2E7C1: mov ecx, var_68
  loc_00E2E7C4: mov var_64, eax
  loc_00E2E7C7: mov var_28, 00000000h
  loc_00E2E7CE: mov edx, [edx]
  loc_00E2E7D0: mov [ebx+00000004h], ecx
  loc_00E2E7D3: mov [ebx+00000008h], eax
  loc_00E2E7D6: mov eax, var_60
  loc_00E2E7D9: mov [ebx+0000000Ch], eax
  loc_00E2E7DC: mov ebx, var_FC
  loc_00E2E7E2: push ebx
  loc_00E2E7E3: call [edx+00000038h]
  loc_00E2E7E6: test eax, eax
  loc_00E2E7E8: fnclex
  loc_00E2E7EA: jge 00E2E7FBh
  loc_00E2E7EC: push 00000038h
  loc_00E2E7EE: push 006E09F8h
  loc_00E2E7F3: push ebx
  loc_00E2E7F4: push eax
  loc_00E2E7F5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E7FB: lea ecx, var_44
  loc_00E2E7FE: lea edx, var_40
  loc_00E2E801: push ecx
  loc_00E2E802: lea eax, var_3C
  loc_00E2E805: push edx
  loc_00E2E806: lea ecx, var_38
  loc_00E2E809: push eax
  loc_00E2E80A: lea edx, var_34
  loc_00E2E80D: push ecx
  loc_00E2E80E: push edx
  loc_00E2E80F: push 00000005h
  loc_00E2E811: call [00401048h] ; __vbaFreeObjList
  loc_00E2E817: lea eax, var_6C
  loc_00E2E81A: lea ecx, var_5C
  loc_00E2E81D: push eax
  loc_00E2E81E: push ecx
  loc_00E2E81F: push 00000002h
  loc_00E2E821: call [00401038h] ; __vbaFreeVarList
  loc_00E2E827: mov edx, [esi]
  loc_00E2E829: add esp, 00000024h
  loc_00E2E82C: push 006DCBD8h
  loc_00E2E831: push 00000000h
  loc_00E2E833: push 00000018h
  loc_00E2E835: push esi
  loc_00E2E836: call [edx+000004A8h]
  loc_00E2E83C: push eax
  loc_00E2E83D: lea eax, var_38
  loc_00E2E840: push eax
  loc_00E2E841: call edi
  loc_00E2E843: lea ecx, var_5C
  loc_00E2E846: push eax
  loc_00E2E847: push ecx
  loc_00E2E848: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2E84E: add esp, 00000010h
  loc_00E2E851: push eax
  loc_00E2E852: call [00401120h] ; __vbaCastObjVar
  loc_00E2E858: lea edx, var_3C
  loc_00E2E85B: push eax
  loc_00E2E85C: push edx
  loc_00E2E85D: call edi
  loc_00E2E85F: mov ebx, eax
  loc_00E2E861: lea ecx, var_40
  loc_00E2E864: push ecx
  loc_00E2E865: push ebx
  loc_00E2E866: mov eax, [ebx]
  loc_00E2E868: call [eax+00000054h]
  loc_00E2E86B: test eax, eax
  loc_00E2E86D: fnclex
  loc_00E2E86F: jge 00E2E880h
  loc_00E2E871: push 00000054h
  loc_00E2E873: push 006DCBE8h
  loc_00E2E878: push ebx
  loc_00E2E879: push eax
  loc_00E2E87A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E880: lea ebx, var_44
  loc_00E2E883: mov eax, var_40
  loc_00E2E886: push ebx
  loc_00E2E887: mov ecx, 00000002h
  loc_00E2E88C: sub esp, 00000010h
  loc_00E2E88F: mov var_94, ecx
  loc_00E2E895: mov ebx, esp
  loc_00E2E897: mov var_9C, ecx
  loc_00E2E89D: mov edx, [eax]
  loc_00E2E89F: push eax
  loc_00E2E8A0: mov [ebx], ecx
  loc_00E2E8A2: mov ecx, var_98
  loc_00E2E8A8: mov var_F4, eax
  loc_00E2E8AE: mov [ebx+00000004h], ecx
  loc_00E2E8B1: mov ecx, var_94
  loc_00E2E8B7: mov [ebx+00000008h], ecx
  loc_00E2E8BA: mov ecx, var_90
  loc_00E2E8C0: mov [ebx+0000000Ch], ecx
  loc_00E2E8C3: call [edx+00000028h]
  loc_00E2E8C6: test eax, eax
  loc_00E2E8C8: fnclex
  loc_00E2E8CA: jge 00E2E8E1h
  loc_00E2E8CC: mov edx, var_F4
  loc_00E2E8D2: push 00000028h
  loc_00E2E8D4: push 006E09E8h
  loc_00E2E8D9: push edx
  loc_00E2E8DA: push eax
  loc_00E2E8DB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E8E1: mov eax, var_44
  loc_00E2E8E4: mov ecx, [esi]
  loc_00E2E8E6: push esi
  loc_00E2E8E7: mov var_FC, eax
  loc_00E2E8ED: call [ecx+00000438h]
  loc_00E2E8F3: lea edx, var_34
  loc_00E2E8F6: push eax
  loc_00E2E8F7: push edx
  loc_00E2E8F8: call edi
  loc_00E2E8FA: mov ebx, eax
  loc_00E2E8FC: lea ecx, var_28
  loc_00E2E8FF: push ecx
  loc_00E2E900: push ebx
  loc_00E2E901: mov eax, [ebx]
  loc_00E2E903: call [eax+00000050h]
  loc_00E2E906: test eax, eax
  loc_00E2E908: fnclex
  loc_00E2E90A: jge 00E2E91Bh
  loc_00E2E90C: push 00000050h
  loc_00E2E90E: push 006E0410h
  loc_00E2E913: push ebx
  loc_00E2E914: push eax
  loc_00E2E915: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E91B: sub esp, 00000010h
  loc_00E2E91E: mov eax, var_28
  loc_00E2E921: mov ebx, esp
  loc_00E2E923: mov ecx, 00000008h
  loc_00E2E928: mov var_6C, ecx
  loc_00E2E92B: mov edx, var_FC
  loc_00E2E931: mov [ebx], ecx
  loc_00E2E933: mov ecx, var_68
  loc_00E2E936: mov var_64, eax
  loc_00E2E939: mov var_28, 00000000h
  loc_00E2E940: mov edx, [edx]
  loc_00E2E942: mov [ebx+00000004h], ecx
  loc_00E2E945: mov [ebx+00000008h], eax
  loc_00E2E948: mov eax, var_60
  loc_00E2E94B: mov [ebx+0000000Ch], eax
  loc_00E2E94E: mov ebx, var_FC
  loc_00E2E954: push ebx
  loc_00E2E955: call [edx+00000038h]
  loc_00E2E958: test eax, eax
  loc_00E2E95A: fnclex
  loc_00E2E95C: jge 00E2E96Dh
  loc_00E2E95E: push 00000038h
  loc_00E2E960: push 006E09F8h
  loc_00E2E965: push ebx
  loc_00E2E966: push eax
  loc_00E2E967: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E96D: lea ecx, var_44
  loc_00E2E970: lea edx, var_40
  loc_00E2E973: push ecx
  loc_00E2E974: lea eax, var_3C
  loc_00E2E977: push edx
  loc_00E2E978: lea ecx, var_38
  loc_00E2E97B: push eax
  loc_00E2E97C: lea edx, var_34
  loc_00E2E97F: push ecx
  loc_00E2E980: push edx
  loc_00E2E981: push 00000005h
  loc_00E2E983: call [00401048h] ; __vbaFreeObjList
  loc_00E2E989: lea eax, var_6C
  loc_00E2E98C: lea ecx, var_5C
  loc_00E2E98F: push eax
  loc_00E2E990: push ecx
  loc_00E2E991: push 00000002h
  loc_00E2E993: call [00401038h] ; __vbaFreeVarList
  loc_00E2E999: mov edx, [esi]
  loc_00E2E99B: add esp, 00000024h
  loc_00E2E99E: push 006DCBD8h
  loc_00E2E9A3: push 00000000h
  loc_00E2E9A5: push 00000018h
  loc_00E2E9A7: push esi
  loc_00E2E9A8: call [edx+000004A8h]
  loc_00E2E9AE: push eax
  loc_00E2E9AF: lea eax, var_38
  loc_00E2E9B2: push eax
  loc_00E2E9B3: call edi
  loc_00E2E9B5: lea ecx, var_5C
  loc_00E2E9B8: push eax
  loc_00E2E9B9: push ecx
  loc_00E2E9BA: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2E9C0: add esp, 00000010h
  loc_00E2E9C3: push eax
  loc_00E2E9C4: call [00401120h] ; __vbaCastObjVar
  loc_00E2E9CA: lea edx, var_3C
  loc_00E2E9CD: push eax
  loc_00E2E9CE: push edx
  loc_00E2E9CF: call edi
  loc_00E2E9D1: mov ebx, eax
  loc_00E2E9D3: lea ecx, var_40
  loc_00E2E9D6: push ecx
  loc_00E2E9D7: push ebx
  loc_00E2E9D8: mov eax, [ebx]
  loc_00E2E9DA: call [eax+00000054h]
  loc_00E2E9DD: test eax, eax
  loc_00E2E9DF: fnclex
  loc_00E2E9E1: jge 00E2E9F2h
  loc_00E2E9E3: push 00000054h
  loc_00E2E9E5: push 006DCBE8h
  loc_00E2E9EA: push ebx
  loc_00E2E9EB: push eax
  loc_00E2E9EC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2E9F2: lea ebx, var_44
  loc_00E2E9F5: mov eax, var_40
  loc_00E2E9F8: push ebx
  loc_00E2E9F9: mov ecx, 00000002h
  loc_00E2E9FE: sub esp, 00000010h
  loc_00E2EA01: mov var_9C, ecx
  loc_00E2EA07: mov ebx, esp
  loc_00E2EA09: mov var_94, 00000003h
  loc_00E2EA13: mov edx, [eax]
  loc_00E2EA15: push eax
  loc_00E2EA16: mov [ebx], ecx
  loc_00E2EA18: mov ecx, var_98
  loc_00E2EA1E: mov var_F4, eax
  loc_00E2EA24: mov [ebx+00000004h], ecx
  loc_00E2EA27: mov ecx, var_94
  loc_00E2EA2D: mov [ebx+00000008h], ecx
  loc_00E2EA30: mov ecx, var_90
  loc_00E2EA36: mov [ebx+0000000Ch], ecx
  loc_00E2EA39: call [edx+00000028h]
  loc_00E2EA3C: test eax, eax
  loc_00E2EA3E: fnclex
  loc_00E2EA40: jge 00E2EA57h
  loc_00E2EA42: mov edx, var_F4
  loc_00E2EA48: push 00000028h
  loc_00E2EA4A: push 006E09E8h
  loc_00E2EA4F: push edx
  loc_00E2EA50: push eax
  loc_00E2EA51: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EA57: mov eax, var_44
  loc_00E2EA5A: mov ecx, [esi]
  loc_00E2EA5C: push esi
  loc_00E2EA5D: mov var_FC, eax
  loc_00E2EA63: call [ecx+0000044Ch]
  loc_00E2EA69: lea edx, var_34
  loc_00E2EA6C: push eax
  loc_00E2EA6D: push edx
  loc_00E2EA6E: call edi
  loc_00E2EA70: mov ebx, eax
  loc_00E2EA72: lea ecx, var_28
  loc_00E2EA75: push ecx
  loc_00E2EA76: push ebx
  loc_00E2EA77: mov eax, [ebx]
  loc_00E2EA79: call [eax+00000050h]
  loc_00E2EA7C: test eax, eax
  loc_00E2EA7E: fnclex
  loc_00E2EA80: jge 00E2EA91h
  loc_00E2EA82: push 00000050h
  loc_00E2EA84: push 006E0410h
  loc_00E2EA89: push ebx
  loc_00E2EA8A: push eax
  loc_00E2EA8B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EA91: sub esp, 00000010h
  loc_00E2EA94: mov eax, var_28
  loc_00E2EA97: mov ebx, esp
  loc_00E2EA99: mov ecx, 00000008h
  loc_00E2EA9E: mov var_6C, ecx
  loc_00E2EAA1: mov edx, var_FC
  loc_00E2EAA7: mov [ebx], ecx
  loc_00E2EAA9: mov ecx, var_68
  loc_00E2EAAC: mov var_64, eax
  loc_00E2EAAF: mov var_28, 00000000h
  loc_00E2EAB6: mov edx, [edx]
  loc_00E2EAB8: mov [ebx+00000004h], ecx
  loc_00E2EABB: mov [ebx+00000008h], eax
  loc_00E2EABE: mov eax, var_60
  loc_00E2EAC1: mov [ebx+0000000Ch], eax
  loc_00E2EAC4: mov ebx, var_FC
  loc_00E2EACA: push ebx
  loc_00E2EACB: call [edx+00000038h]
  loc_00E2EACE: test eax, eax
  loc_00E2EAD0: fnclex
  loc_00E2EAD2: jge 00E2EAE3h
  loc_00E2EAD4: push 00000038h
  loc_00E2EAD6: push 006E09F8h
  loc_00E2EADB: push ebx
  loc_00E2EADC: push eax
  loc_00E2EADD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EAE3: lea ecx, var_44
  loc_00E2EAE6: lea edx, var_40
  loc_00E2EAE9: push ecx
  loc_00E2EAEA: lea eax, var_3C
  loc_00E2EAED: push edx
  loc_00E2EAEE: lea ecx, var_38
  loc_00E2EAF1: push eax
  loc_00E2EAF2: lea edx, var_34
  loc_00E2EAF5: push ecx
  loc_00E2EAF6: push edx
  loc_00E2EAF7: push 00000005h
  loc_00E2EAF9: call [00401048h] ; __vbaFreeObjList
  loc_00E2EAFF: lea eax, var_6C
  loc_00E2EB02: lea ecx, var_5C
  loc_00E2EB05: push eax
  loc_00E2EB06: push ecx
  loc_00E2EB07: push 00000002h
  loc_00E2EB09: call [00401038h] ; __vbaFreeVarList
  loc_00E2EB0F: mov edx, [esi]
  loc_00E2EB11: add esp, 00000024h
  loc_00E2EB14: push 006DCBD8h
  loc_00E2EB19: push 00000000h
  loc_00E2EB1B: push 00000018h
  loc_00E2EB1D: push esi
  loc_00E2EB1E: call [edx+000004A8h]
  loc_00E2EB24: push eax
  loc_00E2EB25: lea eax, var_38
  loc_00E2EB28: push eax
  loc_00E2EB29: call edi
  loc_00E2EB2B: lea ecx, var_5C
  loc_00E2EB2E: push eax
  loc_00E2EB2F: push ecx
  loc_00E2EB30: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2EB36: add esp, 00000010h
  loc_00E2EB39: push eax
  loc_00E2EB3A: call [00401120h] ; __vbaCastObjVar
  loc_00E2EB40: lea edx, var_3C
  loc_00E2EB43: push eax
  loc_00E2EB44: push edx
  loc_00E2EB45: call edi
  loc_00E2EB47: mov ebx, eax
  loc_00E2EB49: lea ecx, var_40
  loc_00E2EB4C: push ecx
  loc_00E2EB4D: push ebx
  loc_00E2EB4E: mov eax, [ebx]
  loc_00E2EB50: call [eax+00000054h]
  loc_00E2EB53: test eax, eax
  loc_00E2EB55: fnclex
  loc_00E2EB57: jge 00E2EB68h
  loc_00E2EB59: push 00000054h
  loc_00E2EB5B: push 006DCBE8h
  loc_00E2EB60: push ebx
  loc_00E2EB61: push eax
  loc_00E2EB62: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EB68: lea ebx, var_44
  loc_00E2EB6B: mov eax, var_40
  loc_00E2EB6E: push ebx
  loc_00E2EB6F: mov ecx, 00000002h
  loc_00E2EB74: sub esp, 00000010h
  loc_00E2EB77: mov var_9C, ecx
  loc_00E2EB7D: mov ebx, esp
  loc_00E2EB7F: mov var_94, 00000004h
  loc_00E2EB89: mov edx, [eax]
  loc_00E2EB8B: push eax
  loc_00E2EB8C: mov [ebx], ecx
  loc_00E2EB8E: mov ecx, var_98
  loc_00E2EB94: mov var_F4, eax
  loc_00E2EB9A: mov [ebx+00000004h], ecx
  loc_00E2EB9D: mov ecx, var_94
  loc_00E2EBA3: mov [ebx+00000008h], ecx
  loc_00E2EBA6: mov ecx, var_90
  loc_00E2EBAC: mov [ebx+0000000Ch], ecx
  loc_00E2EBAF: call [edx+00000028h]
  loc_00E2EBB2: test eax, eax
  loc_00E2EBB4: fnclex
  loc_00E2EBB6: jge 00E2EBCDh
  loc_00E2EBB8: mov edx, var_F4
  loc_00E2EBBE: push 00000028h
  loc_00E2EBC0: push 006E09E8h
  loc_00E2EBC5: push edx
  loc_00E2EBC6: push eax
  loc_00E2EBC7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EBCD: mov eax, var_44
  loc_00E2EBD0: mov ecx, [esi]
  loc_00E2EBD2: push esi
  loc_00E2EBD3: mov var_FC, eax
  loc_00E2EBD9: call [ecx+00000450h]
  loc_00E2EBDF: lea edx, var_34
  loc_00E2EBE2: push eax
  loc_00E2EBE3: push edx
  loc_00E2EBE4: call edi
  loc_00E2EBE6: mov ebx, eax
  loc_00E2EBE8: lea ecx, var_28
  loc_00E2EBEB: push ecx
  loc_00E2EBEC: push ebx
  loc_00E2EBED: mov eax, [ebx]
  loc_00E2EBEF: call [eax+00000050h]
  loc_00E2EBF2: test eax, eax
  loc_00E2EBF4: fnclex
  loc_00E2EBF6: jge 00E2EC07h
  loc_00E2EBF8: push 00000050h
  loc_00E2EBFA: push 006E0410h
  loc_00E2EBFF: push ebx
  loc_00E2EC00: push eax
  loc_00E2EC01: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EC07: sub esp, 00000010h
  loc_00E2EC0A: mov eax, var_28
  loc_00E2EC0D: mov ebx, esp
  loc_00E2EC0F: mov ecx, 00000008h
  loc_00E2EC14: mov var_6C, ecx
  loc_00E2EC17: mov edx, var_FC
  loc_00E2EC1D: mov [ebx], ecx
  loc_00E2EC1F: mov ecx, var_68
  loc_00E2EC22: mov var_64, eax
  loc_00E2EC25: mov var_28, 00000000h
  loc_00E2EC2C: mov edx, [edx]
  loc_00E2EC2E: mov [ebx+00000004h], ecx
  loc_00E2EC31: mov [ebx+00000008h], eax
  loc_00E2EC34: mov eax, var_60
  loc_00E2EC37: mov [ebx+0000000Ch], eax
  loc_00E2EC3A: mov ebx, var_FC
  loc_00E2EC40: push ebx
  loc_00E2EC41: call [edx+00000038h]
  loc_00E2EC44: test eax, eax
  loc_00E2EC46: fnclex
  loc_00E2EC48: jge 00E2EC59h
  loc_00E2EC4A: push 00000038h
  loc_00E2EC4C: push 006E09F8h
  loc_00E2EC51: push ebx
  loc_00E2EC52: push eax
  loc_00E2EC53: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EC59: lea ecx, var_44
  loc_00E2EC5C: lea edx, var_40
  loc_00E2EC5F: push ecx
  loc_00E2EC60: lea eax, var_3C
  loc_00E2EC63: push edx
  loc_00E2EC64: lea ecx, var_38
  loc_00E2EC67: push eax
  loc_00E2EC68: lea edx, var_34
  loc_00E2EC6B: push ecx
  loc_00E2EC6C: push edx
  loc_00E2EC6D: push 00000005h
  loc_00E2EC6F: call [00401048h] ; __vbaFreeObjList
  loc_00E2EC75: lea eax, var_6C
  loc_00E2EC78: lea ecx, var_5C
  loc_00E2EC7B: push eax
  loc_00E2EC7C: push ecx
  loc_00E2EC7D: push 00000002h
  loc_00E2EC7F: call [00401038h] ; __vbaFreeVarList
  loc_00E2EC85: mov edx, [esi]
  loc_00E2EC87: add esp, 00000024h
  loc_00E2EC8A: push 006DCBD8h
  loc_00E2EC8F: push 00000000h
  loc_00E2EC91: push 00000018h
  loc_00E2EC93: push esi
  loc_00E2EC94: call [edx+000004A8h]
  loc_00E2EC9A: push eax
  loc_00E2EC9B: lea eax, var_38
  loc_00E2EC9E: push eax
  loc_00E2EC9F: call edi
  loc_00E2ECA1: lea ecx, var_5C
  loc_00E2ECA4: push eax
  loc_00E2ECA5: push ecx
  loc_00E2ECA6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2ECAC: add esp, 00000010h
  loc_00E2ECAF: push eax
  loc_00E2ECB0: call [00401120h] ; __vbaCastObjVar
  loc_00E2ECB6: lea edx, var_3C
  loc_00E2ECB9: push eax
  loc_00E2ECBA: push edx
  loc_00E2ECBB: call edi
  loc_00E2ECBD: mov ebx, eax
  loc_00E2ECBF: lea ecx, var_40
  loc_00E2ECC2: push ecx
  loc_00E2ECC3: push ebx
  loc_00E2ECC4: mov eax, [ebx]
  loc_00E2ECC6: call [eax+00000054h]
  loc_00E2ECC9: test eax, eax
  loc_00E2ECCB: fnclex
  loc_00E2ECCD: jge 00E2ECDEh
  loc_00E2ECCF: push 00000054h
  loc_00E2ECD1: push 006DCBE8h
  loc_00E2ECD6: push ebx
  loc_00E2ECD7: push eax
  loc_00E2ECD8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2ECDE: lea ebx, var_44
  loc_00E2ECE1: mov eax, var_40
  loc_00E2ECE4: push ebx
  loc_00E2ECE5: mov ecx, 00000002h
  loc_00E2ECEA: sub esp, 00000010h
  loc_00E2ECED: mov var_9C, ecx
  loc_00E2ECF3: mov ebx, esp
  loc_00E2ECF5: mov var_94, 00000005h
  loc_00E2ECFF: mov edx, [eax]
  loc_00E2ED01: push eax
  loc_00E2ED02: mov [ebx], ecx
  loc_00E2ED04: mov ecx, var_98
  loc_00E2ED0A: mov var_F4, eax
  loc_00E2ED10: mov [ebx+00000004h], ecx
  loc_00E2ED13: mov ecx, var_94
  loc_00E2ED19: mov [ebx+00000008h], ecx
  loc_00E2ED1C: mov ecx, var_90
  loc_00E2ED22: mov [ebx+0000000Ch], ecx
  loc_00E2ED25: call [edx+00000028h]
  loc_00E2ED28: test eax, eax
  loc_00E2ED2A: fnclex
  loc_00E2ED2C: jge 00E2ED43h
  loc_00E2ED2E: mov edx, var_F4
  loc_00E2ED34: push 00000028h
  loc_00E2ED36: push 006E09E8h
  loc_00E2ED3B: push edx
  loc_00E2ED3C: push eax
  loc_00E2ED3D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2ED43: mov eax, var_44
  loc_00E2ED46: mov ecx, [esi]
  loc_00E2ED48: push esi
  loc_00E2ED49: mov var_FC, eax
  loc_00E2ED4F: call [ecx+00000424h]
  loc_00E2ED55: lea edx, var_34
  loc_00E2ED58: push eax
  loc_00E2ED59: push edx
  loc_00E2ED5A: call edi
  loc_00E2ED5C: mov ebx, eax
  loc_00E2ED5E: lea ecx, var_28
  loc_00E2ED61: push ecx
  loc_00E2ED62: push ebx
  loc_00E2ED63: mov eax, [ebx]
  loc_00E2ED65: call [eax+00000050h]
  loc_00E2ED68: test eax, eax
  loc_00E2ED6A: fnclex
  loc_00E2ED6C: jge 00E2ED7Dh
  loc_00E2ED6E: push 00000050h
  loc_00E2ED70: push 006E0410h
  loc_00E2ED75: push ebx
  loc_00E2ED76: push eax
  loc_00E2ED77: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2ED7D: sub esp, 00000010h
  loc_00E2ED80: mov eax, var_28
  loc_00E2ED83: mov ebx, esp
  loc_00E2ED85: mov ecx, 00000008h
  loc_00E2ED8A: mov var_6C, ecx
  loc_00E2ED8D: mov edx, var_FC
  loc_00E2ED93: mov [ebx], ecx
  loc_00E2ED95: mov ecx, var_68
  loc_00E2ED98: mov var_64, eax
  loc_00E2ED9B: mov var_28, 00000000h
  loc_00E2EDA2: mov edx, [edx]
  loc_00E2EDA4: mov [ebx+00000004h], ecx
  loc_00E2EDA7: mov [ebx+00000008h], eax
  loc_00E2EDAA: mov eax, var_60
  loc_00E2EDAD: mov [ebx+0000000Ch], eax
  loc_00E2EDB0: mov ebx, var_FC
  loc_00E2EDB6: push ebx
  loc_00E2EDB7: call [edx+00000038h]
  loc_00E2EDBA: test eax, eax
  loc_00E2EDBC: fnclex
  loc_00E2EDBE: jge 00E2EDCFh
  loc_00E2EDC0: push 00000038h
  loc_00E2EDC2: push 006E09F8h
  loc_00E2EDC7: push ebx
  loc_00E2EDC8: push eax
  loc_00E2EDC9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EDCF: lea ecx, var_44
  loc_00E2EDD2: lea edx, var_40
  loc_00E2EDD5: push ecx
  loc_00E2EDD6: lea eax, var_3C
  loc_00E2EDD9: push edx
  loc_00E2EDDA: lea ecx, var_38
  loc_00E2EDDD: push eax
  loc_00E2EDDE: lea edx, var_34
  loc_00E2EDE1: push ecx
  loc_00E2EDE2: push edx
  loc_00E2EDE3: push 00000005h
  loc_00E2EDE5: call [00401048h] ; __vbaFreeObjList
  loc_00E2EDEB: lea eax, var_6C
  loc_00E2EDEE: lea ecx, var_5C
  loc_00E2EDF1: push eax
  loc_00E2EDF2: push ecx
  loc_00E2EDF3: push 00000002h
  loc_00E2EDF5: call [00401038h] ; __vbaFreeVarList
  loc_00E2EDFB: mov edx, [esi]
  loc_00E2EDFD: add esp, 00000024h
  loc_00E2EE00: push 006DCBD8h
  loc_00E2EE05: push 00000000h
  loc_00E2EE07: push 00000018h
  loc_00E2EE09: push esi
  loc_00E2EE0A: call [edx+000004A8h]
  loc_00E2EE10: push eax
  loc_00E2EE11: lea eax, var_34
  loc_00E2EE14: push eax
  loc_00E2EE15: call edi
  loc_00E2EE17: lea ecx, var_5C
  loc_00E2EE1A: push eax
  loc_00E2EE1B: push ecx
  loc_00E2EE1C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2EE22: add esp, 00000010h
  loc_00E2EE25: push eax
  loc_00E2EE26: call [00401120h] ; __vbaCastObjVar
  loc_00E2EE2C: lea edx, var_38
  loc_00E2EE2F: push eax
  loc_00E2EE30: push edx
  loc_00E2EE31: call edi
  loc_00E2EE33: mov ebx, eax
  loc_00E2EE35: lea ecx, var_3C
  loc_00E2EE38: push ecx
  loc_00E2EE39: push ebx
  loc_00E2EE3A: mov eax, [ebx]
  loc_00E2EE3C: call [eax+00000054h]
  loc_00E2EE3F: test eax, eax
  loc_00E2EE41: fnclex
  loc_00E2EE43: jge 00E2EE54h
  loc_00E2EE45: push 00000054h
  loc_00E2EE47: push 006DCBE8h
  loc_00E2EE4C: push ebx
  loc_00E2EE4D: push eax
  loc_00E2EE4E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EE54: lea ebx, var_40
  loc_00E2EE57: mov eax, var_3C
  loc_00E2EE5A: push ebx
  loc_00E2EE5B: mov ecx, 00000002h
  loc_00E2EE60: sub esp, 00000010h
  loc_00E2EE63: mov var_9C, ecx
  loc_00E2EE69: mov ebx, esp
  loc_00E2EE6B: mov var_94, 00000006h
  loc_00E2EE75: mov edx, [eax]
  loc_00E2EE77: push eax
  loc_00E2EE78: mov [ebx], ecx
  loc_00E2EE7A: mov ecx, var_98
  loc_00E2EE80: mov var_EC, eax
  loc_00E2EE86: mov [ebx+00000004h], ecx
  loc_00E2EE89: mov ecx, var_94
  loc_00E2EE8F: mov [ebx+00000008h], ecx
  loc_00E2EE92: mov ecx, var_90
  loc_00E2EE98: mov [ebx+0000000Ch], ecx
  loc_00E2EE9B: call [edx+00000028h]
  loc_00E2EE9E: test eax, eax
  loc_00E2EEA0: fnclex
  loc_00E2EEA2: jge 00E2EEB9h
  loc_00E2EEA4: mov edx, var_EC
  loc_00E2EEAA: push 00000028h
  loc_00E2EEAC: push 006E09E8h
  loc_00E2EEB1: push edx
  loc_00E2EEB2: push eax
  loc_00E2EEB3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EEB9: mov eax, [esi]
  loc_00E2EEBB: mov ebx, var_40
  loc_00E2EEBE: push esi
  loc_00E2EEBF: call [eax+000002FCh]
  loc_00E2EEC5: sub esp, 00000010h
  loc_00E2EEC8: mov var_64, eax
  loc_00E2EECB: mov edx, esp
  loc_00E2EECD: mov eax, 00000009h
  loc_00E2EED2: mov var_6C, eax
  loc_00E2EED5: mov ecx, [ebx]
  loc_00E2EED7: mov [edx], eax
  loc_00E2EED9: mov eax, var_68
  loc_00E2EEDC: push ebx
  loc_00E2EEDD: mov [edx+00000004h], eax
  loc_00E2EEE0: mov eax, var_64
  loc_00E2EEE3: mov [edx+00000008h], eax
  loc_00E2EEE6: mov eax, var_60
  loc_00E2EEE9: mov [edx+0000000Ch], eax
  loc_00E2EEEC: call [ecx+00000038h]
  loc_00E2EEEF: test eax, eax
  loc_00E2EEF1: fnclex
  loc_00E2EEF3: jge 00E2EF04h
  loc_00E2EEF5: push 00000038h
  loc_00E2EEF7: push 006E09F8h
  loc_00E2EEFC: push ebx
  loc_00E2EEFD: push eax
  loc_00E2EEFE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EF04: lea ecx, var_40
  loc_00E2EF07: lea edx, var_3C
  loc_00E2EF0A: push ecx
  loc_00E2EF0B: lea eax, var_38
  loc_00E2EF0E: push edx
  loc_00E2EF0F: lea ecx, var_34
  loc_00E2EF12: push eax
  loc_00E2EF13: push ecx
  loc_00E2EF14: push 00000004h
  loc_00E2EF16: call [00401048h] ; __vbaFreeObjList
  loc_00E2EF1C: lea edx, var_6C
  loc_00E2EF1F: lea eax, var_5C
  loc_00E2EF22: push edx
  loc_00E2EF23: push eax
  loc_00E2EF24: push 00000002h
  loc_00E2EF26: call [00401038h] ; __vbaFreeVarList
  loc_00E2EF2C: mov ecx, [esi]
  loc_00E2EF2E: add esp, 00000020h
  loc_00E2EF31: push 006DCBD8h
  loc_00E2EF36: push 00000000h
  loc_00E2EF38: push 00000018h
  loc_00E2EF3A: push esi
  loc_00E2EF3B: call [ecx+000004A8h]
  loc_00E2EF41: lea edx, var_38
  loc_00E2EF44: push eax
  loc_00E2EF45: push edx
  loc_00E2EF46: call edi
  loc_00E2EF48: push eax
  loc_00E2EF49: lea eax, var_5C
  loc_00E2EF4C: push eax
  loc_00E2EF4D: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2EF53: add esp, 00000010h
  loc_00E2EF56: push eax
  loc_00E2EF57: call [00401120h] ; __vbaCastObjVar
  loc_00E2EF5D: lea ecx, var_3C
  loc_00E2EF60: push eax
  loc_00E2EF61: push ecx
  loc_00E2EF62: call edi
  loc_00E2EF64: mov ebx, eax
  loc_00E2EF66: lea eax, var_40
  loc_00E2EF69: push eax
  loc_00E2EF6A: push ebx
  loc_00E2EF6B: mov edx, [ebx]
  loc_00E2EF6D: call [edx+00000054h]
  loc_00E2EF70: test eax, eax
  loc_00E2EF72: fnclex
  loc_00E2EF74: jge 00E2EF85h
  loc_00E2EF76: push 00000054h
  loc_00E2EF78: push 006DCBE8h
  loc_00E2EF7D: push ebx
  loc_00E2EF7E: push eax
  loc_00E2EF7F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EF85: lea ebx, var_44
  loc_00E2EF88: mov eax, var_40
  loc_00E2EF8B: push ebx
  loc_00E2EF8C: mov ecx, 00000002h
  loc_00E2EF91: sub esp, 00000010h
  loc_00E2EF94: mov var_9C, ecx
  loc_00E2EF9A: mov ebx, esp
  loc_00E2EF9C: mov var_94, 00000007h
  loc_00E2EFA6: mov edx, [eax]
  loc_00E2EFA8: push eax
  loc_00E2EFA9: mov [ebx], ecx
  loc_00E2EFAB: mov ecx, var_98
  loc_00E2EFB1: mov var_F4, eax
  loc_00E2EFB7: mov [ebx+00000004h], ecx
  loc_00E2EFBA: mov ecx, var_94
  loc_00E2EFC0: mov [ebx+00000008h], ecx
  loc_00E2EFC3: mov ecx, var_90
  loc_00E2EFC9: mov [ebx+0000000Ch], ecx
  loc_00E2EFCC: call [edx+00000028h]
  loc_00E2EFCF: test eax, eax
  loc_00E2EFD1: fnclex
  loc_00E2EFD3: jge 00E2EFEAh
  loc_00E2EFD5: mov edx, var_F4
  loc_00E2EFDB: push 00000028h
  loc_00E2EFDD: push 006E09E8h
  loc_00E2EFE2: push edx
  loc_00E2EFE3: push eax
  loc_00E2EFE4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2EFEA: mov eax, var_44
  loc_00E2EFED: mov ecx, [esi]
  loc_00E2EFEF: push esi
  loc_00E2EFF0: mov var_FC, eax
  loc_00E2EFF6: call [ecx+0000039Ch]
  loc_00E2EFFC: lea edx, var_34
  loc_00E2EFFF: push eax
  loc_00E2F000: push edx
  loc_00E2F001: call edi
  loc_00E2F003: mov ebx, eax
  loc_00E2F005: lea ecx, var_28
  loc_00E2F008: push ecx
  loc_00E2F009: push ebx
  loc_00E2F00A: mov eax, [ebx]
  loc_00E2F00C: call [eax+000000A0h]
  loc_00E2F012: test eax, eax
  loc_00E2F014: fnclex
  loc_00E2F016: jge 00E2F02Ah
  loc_00E2F018: push 000000A0h
  loc_00E2F01D: push 006DCB70h
  loc_00E2F022: push ebx
  loc_00E2F023: push eax
  loc_00E2F024: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F02A: sub esp, 00000010h
  loc_00E2F02D: mov eax, var_28
  loc_00E2F030: mov ebx, esp
  loc_00E2F032: mov ecx, 00000008h
  loc_00E2F037: mov var_6C, ecx
  loc_00E2F03A: mov edx, var_FC
  loc_00E2F040: mov [ebx], ecx
  loc_00E2F042: mov ecx, var_68
  loc_00E2F045: mov var_64, eax
  loc_00E2F048: mov var_28, 00000000h
  loc_00E2F04F: mov edx, [edx]
  loc_00E2F051: mov [ebx+00000004h], ecx
  loc_00E2F054: mov [ebx+00000008h], eax
  loc_00E2F057: mov eax, var_60
  loc_00E2F05A: mov [ebx+0000000Ch], eax
  loc_00E2F05D: mov ebx, var_FC
  loc_00E2F063: push ebx
  loc_00E2F064: call [edx+00000038h]
  loc_00E2F067: test eax, eax
  loc_00E2F069: fnclex
  loc_00E2F06B: jge 00E2F07Ch
  loc_00E2F06D: push 00000038h
  loc_00E2F06F: push 006E09F8h
  loc_00E2F074: push ebx
  loc_00E2F075: push eax
  loc_00E2F076: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F07C: lea ecx, var_44
  loc_00E2F07F: lea edx, var_40
  loc_00E2F082: push ecx
  loc_00E2F083: lea eax, var_3C
  loc_00E2F086: push edx
  loc_00E2F087: lea ecx, var_38
  loc_00E2F08A: push eax
  loc_00E2F08B: lea edx, var_34
  loc_00E2F08E: push ecx
  loc_00E2F08F: push edx
  loc_00E2F090: push 00000005h
  loc_00E2F092: call [00401048h] ; __vbaFreeObjList
  loc_00E2F098: lea eax, var_6C
  loc_00E2F09B: lea ecx, var_5C
  loc_00E2F09E: push eax
  loc_00E2F09F: push ecx
  loc_00E2F0A0: push 00000002h
  loc_00E2F0A2: call [00401038h] ; __vbaFreeVarList
  loc_00E2F0A8: mov edx, [esi]
  loc_00E2F0AA: add esp, 00000024h
  loc_00E2F0AD: push 006DCBD8h
  loc_00E2F0B2: push 00000000h
  loc_00E2F0B4: push 00000018h
  loc_00E2F0B6: push esi
  loc_00E2F0B7: call [edx+000004A8h]
  loc_00E2F0BD: push eax
  loc_00E2F0BE: lea eax, var_38
  loc_00E2F0C1: push eax
  loc_00E2F0C2: call edi
  loc_00E2F0C4: lea ecx, var_5C
  loc_00E2F0C7: push eax
  loc_00E2F0C8: push ecx
  loc_00E2F0C9: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2F0CF: add esp, 00000010h
  loc_00E2F0D2: push eax
  loc_00E2F0D3: call [00401120h] ; __vbaCastObjVar
  loc_00E2F0D9: lea edx, var_3C
  loc_00E2F0DC: push eax
  loc_00E2F0DD: push edx
  loc_00E2F0DE: call edi
  loc_00E2F0E0: mov ebx, eax
  loc_00E2F0E2: lea ecx, var_40
  loc_00E2F0E5: push ecx
  loc_00E2F0E6: push ebx
  loc_00E2F0E7: mov eax, [ebx]
  loc_00E2F0E9: call [eax+00000054h]
  loc_00E2F0EC: test eax, eax
  loc_00E2F0EE: fnclex
  loc_00E2F0F0: jge 00E2F101h
  loc_00E2F0F2: push 00000054h
  loc_00E2F0F4: push 006DCBE8h
  loc_00E2F0F9: push ebx
  loc_00E2F0FA: push eax
  loc_00E2F0FB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F101: lea ebx, var_44
  loc_00E2F104: mov eax, var_40
  loc_00E2F107: push ebx
  loc_00E2F108: mov ecx, 00000002h
  loc_00E2F10D: sub esp, 00000010h
  loc_00E2F110: mov var_9C, ecx
  loc_00E2F116: mov ebx, esp
  loc_00E2F118: mov var_94, 00000008h
  loc_00E2F122: mov edx, [eax]
  loc_00E2F124: push eax
  loc_00E2F125: mov [ebx], ecx
  loc_00E2F127: mov ecx, var_98
  loc_00E2F12D: mov var_F4, eax
  loc_00E2F133: mov [ebx+00000004h], ecx
  loc_00E2F136: mov ecx, var_94
  loc_00E2F13C: mov [ebx+00000008h], ecx
  loc_00E2F13F: mov ecx, var_90
  loc_00E2F145: mov [ebx+0000000Ch], ecx
  loc_00E2F148: call [edx+00000028h]
  loc_00E2F14B: test eax, eax
  loc_00E2F14D: fnclex
  loc_00E2F14F: jge 00E2F166h
  loc_00E2F151: mov edx, var_F4
  loc_00E2F157: push 00000028h
  loc_00E2F159: push 006E09E8h
  loc_00E2F15E: push edx
  loc_00E2F15F: push eax
  loc_00E2F160: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F166: mov eax, var_44
  loc_00E2F169: mov ecx, [esi]
  loc_00E2F16B: push esi
  loc_00E2F16C: mov var_FC, eax
  loc_00E2F172: call [ecx+00000368h]
  loc_00E2F178: lea edx, var_34
  loc_00E2F17B: push eax
  loc_00E2F17C: push edx
  loc_00E2F17D: call edi
  loc_00E2F17F: mov ebx, eax
  loc_00E2F181: lea ecx, var_28
  loc_00E2F184: push ecx
  loc_00E2F185: push ebx
  loc_00E2F186: mov eax, [ebx]
  loc_00E2F188: call [eax+00000050h]
  loc_00E2F18B: test eax, eax
  loc_00E2F18D: fnclex
  loc_00E2F18F: jge 00E2F1A0h
  loc_00E2F191: push 00000050h
  loc_00E2F193: push 006E0410h
  loc_00E2F198: push ebx
  loc_00E2F199: push eax
  loc_00E2F19A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F1A0: sub esp, 00000010h
  loc_00E2F1A3: mov eax, var_28
  loc_00E2F1A6: mov ebx, esp
  loc_00E2F1A8: mov ecx, 00000008h
  loc_00E2F1AD: mov var_6C, ecx
  loc_00E2F1B0: mov edx, var_FC
  loc_00E2F1B6: mov [ebx], ecx
  loc_00E2F1B8: mov ecx, var_68
  loc_00E2F1BB: mov var_64, eax
  loc_00E2F1BE: mov var_28, 00000000h
  loc_00E2F1C5: mov edx, [edx]
  loc_00E2F1C7: mov [ebx+00000004h], ecx
  loc_00E2F1CA: mov [ebx+00000008h], eax
  loc_00E2F1CD: mov eax, var_60
  loc_00E2F1D0: mov [ebx+0000000Ch], eax
  loc_00E2F1D3: mov ebx, var_FC
  loc_00E2F1D9: push ebx
  loc_00E2F1DA: call [edx+00000038h]
  loc_00E2F1DD: test eax, eax
  loc_00E2F1DF: fnclex
  loc_00E2F1E1: jge 00E2F1F2h
  loc_00E2F1E3: push 00000038h
  loc_00E2F1E5: push 006E09F8h
  loc_00E2F1EA: push ebx
  loc_00E2F1EB: push eax
  loc_00E2F1EC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F1F2: lea ecx, var_44
  loc_00E2F1F5: lea edx, var_40
  loc_00E2F1F8: push ecx
  loc_00E2F1F9: lea eax, var_3C
  loc_00E2F1FC: push edx
  loc_00E2F1FD: lea ecx, var_38
  loc_00E2F200: push eax
  loc_00E2F201: lea edx, var_34
  loc_00E2F204: push ecx
  loc_00E2F205: push edx
  loc_00E2F206: push 00000005h
  loc_00E2F208: call [00401048h] ; __vbaFreeObjList
  loc_00E2F20E: lea eax, var_6C
  loc_00E2F211: lea ecx, var_5C
  loc_00E2F214: push eax
  loc_00E2F215: push ecx
  loc_00E2F216: push 00000002h
  loc_00E2F218: call [00401038h] ; __vbaFreeVarList
  loc_00E2F21E: mov edx, [esi]
  loc_00E2F220: add esp, 00000024h
  loc_00E2F223: push esi
  loc_00E2F224: call [edx+0000037Ch]
  loc_00E2F22A: push eax
  loc_00E2F22B: lea eax, var_34
  loc_00E2F22E: push eax
  loc_00E2F22F: call edi
  loc_00E2F231: mov ebx, eax
  loc_00E2F233: lea edx, var_D0
  loc_00E2F239: push edx
  loc_00E2F23A: push ebx
  loc_00E2F23B: mov ecx, [ebx]
  loc_00E2F23D: call [ecx+00000098h]
  loc_00E2F243: test eax, eax
  loc_00E2F245: fnclex
  loc_00E2F247: jge 00E2F25Bh
  loc_00E2F249: push 00000098h
  loc_00E2F24E: push 006DCAD0h
  loc_00E2F253: push ebx
  loc_00E2F254: push eax
  loc_00E2F255: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F25B: xor ebx, ebx
  loc_00E2F25D: cmp var_D0, FFFFFFh
  loc_00E2F265: lea ecx, var_34
  loc_00E2F268: setz bl
  loc_00E2F26B: neg ebx
  loc_00E2F26D: call [00401254h] ; __vbaFreeObj
  loc_00E2F273: test bx, bx
  loc_00E2F276: jz 00E2F3F7h
  loc_00E2F27C: mov eax, [esi]
  loc_00E2F27E: push 006DCBD8h
  loc_00E2F283: push 00000000h
  loc_00E2F285: push 00000018h
  loc_00E2F287: push esi
  loc_00E2F288: call [eax+000004A8h]
  loc_00E2F28E: lea ecx, var_38
  loc_00E2F291: push eax
  loc_00E2F292: push ecx
  loc_00E2F293: call edi
  loc_00E2F295: lea edx, var_5C
  loc_00E2F298: push eax
  loc_00E2F299: push edx
  loc_00E2F29A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2F2A0: add esp, 00000010h
  loc_00E2F2A3: push eax
  loc_00E2F2A4: call [00401120h] ; __vbaCastObjVar
  loc_00E2F2AA: push eax
  loc_00E2F2AB: lea eax, var_3C
  loc_00E2F2AE: push eax
  loc_00E2F2AF: call edi
  loc_00E2F2B1: mov ebx, eax
  loc_00E2F2B3: lea edx, var_40
  loc_00E2F2B6: push edx
  loc_00E2F2B7: push ebx
  loc_00E2F2B8: mov ecx, [ebx]
  loc_00E2F2BA: call [ecx+00000054h]
  loc_00E2F2BD: test eax, eax
  loc_00E2F2BF: fnclex
  loc_00E2F2C1: jge 00E2F2D2h
  loc_00E2F2C3: push 00000054h
  loc_00E2F2C5: push 006DCBE8h
  loc_00E2F2CA: push ebx
  loc_00E2F2CB: push eax
  loc_00E2F2CC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F2D2: lea ebx, var_44
  loc_00E2F2D5: mov eax, var_40
  loc_00E2F2D8: push ebx
  loc_00E2F2D9: mov ecx, 00000002h
  loc_00E2F2DE: sub esp, 00000010h
  loc_00E2F2E1: mov var_9C, ecx
  loc_00E2F2E7: mov ebx, esp
  loc_00E2F2E9: mov var_94, 00000009h
  loc_00E2F2F3: mov edx, [eax]
  loc_00E2F2F5: push eax
  loc_00E2F2F6: mov [ebx], ecx
  loc_00E2F2F8: mov ecx, var_98
  loc_00E2F2FE: mov var_F4, eax
  loc_00E2F304: mov [ebx+00000004h], ecx
  loc_00E2F307: mov ecx, var_94
  loc_00E2F30D: mov [ebx+00000008h], ecx
  loc_00E2F310: mov ecx, var_90
  loc_00E2F316: mov [ebx+0000000Ch], ecx
  loc_00E2F319: call [edx+00000028h]
  loc_00E2F31C: test eax, eax
  loc_00E2F31E: fnclex
  loc_00E2F320: jge 00E2F337h
  loc_00E2F322: mov edx, var_F4
  loc_00E2F328: push 00000028h
  loc_00E2F32A: push 006E09E8h
  loc_00E2F32F: push edx
  loc_00E2F330: push eax
  loc_00E2F331: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F337: mov eax, var_44
  loc_00E2F33A: mov ecx, [esi]
  loc_00E2F33C: push esi
  loc_00E2F33D: mov var_FC, eax
  loc_00E2F343: call [ecx+00000388h]
  loc_00E2F349: lea edx, var_34
  loc_00E2F34C: push eax
  loc_00E2F34D: push edx
  loc_00E2F34E: call edi
  loc_00E2F350: mov ebx, eax
  loc_00E2F352: lea ecx, var_28
  loc_00E2F355: push ecx
  loc_00E2F356: push ebx
  loc_00E2F357: mov eax, [ebx]
  loc_00E2F359: call [eax+00000050h]
  loc_00E2F35C: test eax, eax
  loc_00E2F35E: fnclex
  loc_00E2F360: jge 00E2F371h
  loc_00E2F362: push 00000050h
  loc_00E2F364: push 006E0410h
  loc_00E2F369: push ebx
  loc_00E2F36A: push eax
  loc_00E2F36B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F371: sub esp, 00000010h
  loc_00E2F374: mov eax, var_28
  loc_00E2F377: mov ebx, esp
  loc_00E2F379: mov ecx, 00000008h
  loc_00E2F37E: mov var_6C, ecx
  loc_00E2F381: mov edx, var_FC
  loc_00E2F387: mov [ebx], ecx
  loc_00E2F389: mov ecx, var_68
  loc_00E2F38C: mov var_64, eax
  loc_00E2F38F: mov var_28, 00000000h
  loc_00E2F396: mov edx, [edx]
  loc_00E2F398: mov [ebx+00000004h], ecx
  loc_00E2F39B: mov [ebx+00000008h], eax
  loc_00E2F39E: mov eax, var_60
  loc_00E2F3A1: mov [ebx+0000000Ch], eax
  loc_00E2F3A4: mov ebx, var_FC
  loc_00E2F3AA: push ebx
  loc_00E2F3AB: call [edx+00000038h]
  loc_00E2F3AE: test eax, eax
  loc_00E2F3B0: fnclex
  loc_00E2F3B2: jge 00E2F3C3h
  loc_00E2F3B4: push 00000038h
  loc_00E2F3B6: push 006E09F8h
  loc_00E2F3BB: push ebx
  loc_00E2F3BC: push eax
  loc_00E2F3BD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F3C3: lea ecx, var_44
  loc_00E2F3C6: lea edx, var_40
  loc_00E2F3C9: push ecx
  loc_00E2F3CA: lea eax, var_3C
  loc_00E2F3CD: push edx
  loc_00E2F3CE: lea ecx, var_38
  loc_00E2F3D1: push eax
  loc_00E2F3D2: lea edx, var_34
  loc_00E2F3D5: push ecx
  loc_00E2F3D6: push edx
  loc_00E2F3D7: push 00000005h
  loc_00E2F3D9: call [00401048h] ; __vbaFreeObjList
  loc_00E2F3DF: lea eax, var_6C
  loc_00E2F3E2: lea ecx, var_5C
  loc_00E2F3E5: push eax
  loc_00E2F3E6: push ecx
  loc_00E2F3E7: push 00000002h
  loc_00E2F3E9: call [00401038h] ; __vbaFreeVarList
  loc_00E2F3EF: add esp, 00000024h
  loc_00E2F3F2: jmp 00E2F917h
  loc_00E2F3F7: mov edx, [esi]
  loc_00E2F3F9: push esi
  loc_00E2F3FA: call [edx+00000398h]
  loc_00E2F400: push eax
  loc_00E2F401: lea eax, var_34
  loc_00E2F404: push eax
  loc_00E2F405: call edi
  loc_00E2F407: mov ebx, eax
  loc_00E2F409: lea edx, var_D0
  loc_00E2F40F: push edx
  loc_00E2F410: push ebx
  loc_00E2F411: mov ecx, [ebx]
  loc_00E2F413: call [ecx+00000098h]
  loc_00E2F419: test eax, eax
  loc_00E2F41B: fnclex
  loc_00E2F41D: jge 00E2F431h
  loc_00E2F41F: push 00000098h
  loc_00E2F424: push 006DCAD0h
  loc_00E2F429: push ebx
  loc_00E2F42A: push eax
  loc_00E2F42B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F431: xor ebx, ebx
  loc_00E2F433: cmp var_D0, FFFFFFh
  loc_00E2F43B: lea ecx, var_34
  loc_00E2F43E: setz bl
  loc_00E2F441: neg ebx
  loc_00E2F443: call [00401254h] ; __vbaFreeObj
  loc_00E2F449: test bx, bx
  loc_00E2F44C: jz 00E2F917h
  loc_00E2F452: mov eax, [esi]
  loc_00E2F454: push esi
  loc_00E2F455: call [eax+000003B0h]
  loc_00E2F45B: lea ecx, var_34
  loc_00E2F45E: push eax
  loc_00E2F45F: push ecx
  loc_00E2F460: call edi
  loc_00E2F462: mov ebx, eax
  loc_00E2F464: lea eax, var_28
  loc_00E2F467: push eax
  loc_00E2F468: push ebx
  loc_00E2F469: mov edx, [ebx]
  loc_00E2F46B: call [edx+00000050h]
  loc_00E2F46E: test eax, eax
  loc_00E2F470: fnclex
  loc_00E2F472: jge 00E2F483h
  loc_00E2F474: push 00000050h
  loc_00E2F476: push 006E0410h
  loc_00E2F47B: push ebx
  loc_00E2F47C: push eax
  loc_00E2F47D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F483: mov ecx, var_28
  loc_00E2F486: push ecx
  loc_00E2F487: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2F48D: call [004010D8h] ; __vbaFpR8
  loc_00E2F493: fcomp real8 ptr [004022E8h]
  loc_00E2F499: fnstsw ax
  loc_00E2F49B: test ah, 41h
  loc_00E2F49E: jz 00E2F4A7h
  loc_00E2F4A0: mov ebx, 00000001h
  loc_00E2F4A5: jmp 00E2F4A9h
  loc_00E2F4A7: xor ebx, ebx
  loc_00E2F4A9: lea ecx, var_28
  loc_00E2F4AC: call [00401258h] ; __vbaFreeStr
  loc_00E2F4B2: lea ecx, var_34
  loc_00E2F4B5: call [00401254h] ; __vbaFreeObj
  loc_00E2F4BB: mov edx, [esi]
  loc_00E2F4BD: push esi
  loc_00E2F4BE: neg ebx
  loc_00E2F4C0: test bx, bx
  loc_00E2F4C3: jz 00E2F6FAh
  loc_00E2F4C9: call [edx+00000354h]
  loc_00E2F4CF: push eax
  loc_00E2F4D0: lea eax, var_38
  loc_00E2F4D3: push eax
  loc_00E2F4D4: call edi
  loc_00E2F4D6: mov ebx, eax
  loc_00E2F4D8: lea edx, var_2C
  loc_00E2F4DB: push edx
  loc_00E2F4DC: push ebx
  loc_00E2F4DD: mov ecx, [ebx]
  loc_00E2F4DF: call [ecx+00000050h]
  loc_00E2F4E2: test eax, eax
  loc_00E2F4E4: fnclex
  loc_00E2F4E6: jge 00E2F4F7h
  loc_00E2F4E8: push 00000050h
  loc_00E2F4EA: push 006E0410h
  loc_00E2F4EF: push ebx
  loc_00E2F4F0: push eax
  loc_00E2F4F1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F4F7: mov eax, var_2C
  loc_00E2F4FA: push eax
  loc_00E2F4FB: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2F501: mov ecx, [esi]
  loc_00E2F503: push esi
  loc_00E2F504: fstp real8 ptr var_D8
  loc_00E2F50A: call [ecx+000003B0h]
  loc_00E2F510: lea edx, var_3C
  loc_00E2F513: push eax
  loc_00E2F514: push edx
  loc_00E2F515: call edi
  loc_00E2F517: mov ebx, eax
  loc_00E2F519: lea ecx, var_30
  loc_00E2F51C: push ecx
  loc_00E2F51D: push ebx
  loc_00E2F51E: mov eax, [ebx]
  loc_00E2F520: call [eax+00000050h]
  loc_00E2F523: test eax, eax
  loc_00E2F525: fnclex
  loc_00E2F527: jge 00E2F538h
  loc_00E2F529: push 00000050h
  loc_00E2F52B: push 006E0410h
  loc_00E2F530: push ebx
  loc_00E2F531: push eax
  loc_00E2F532: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F538: mov edx, var_30
  loc_00E2F53B: push edx
  loc_00E2F53C: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2F542: mov eax, [esi]
  loc_00E2F544: push 006DCBD8h
  loc_00E2F549: fstp real8 ptr var_E0
  loc_00E2F54F: push 00000000h
  loc_00E2F551: push 00000018h
  loc_00E2F553: push esi
  loc_00E2F554: call [eax+000004A8h]
  loc_00E2F55A: lea ecx, var_40
  loc_00E2F55D: push eax
  loc_00E2F55E: push ecx
  loc_00E2F55F: call edi
  loc_00E2F561: lea edx, var_5C
  loc_00E2F564: push eax
  loc_00E2F565: push edx
  loc_00E2F566: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2F56C: add esp, 00000010h
  loc_00E2F56F: push eax
  loc_00E2F570: call [00401120h] ; __vbaCastObjVar
  loc_00E2F576: push eax
  loc_00E2F577: lea eax, var_44
  loc_00E2F57A: push eax
  loc_00E2F57B: call edi
  loc_00E2F57D: mov ebx, eax
  loc_00E2F57F: lea edx, var_48
  loc_00E2F582: push edx
  loc_00E2F583: push ebx
  loc_00E2F584: mov ecx, [ebx]
  loc_00E2F586: call [ecx+00000054h]
  loc_00E2F589: test eax, eax
  loc_00E2F58B: fnclex
  loc_00E2F58D: jge 00E2F59Eh
  loc_00E2F58F: push 00000054h
  loc_00E2F591: push 006DCBE8h
  loc_00E2F596: push ebx
  loc_00E2F597: push eax
  loc_00E2F598: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F59E: lea ebx, var_4C
  loc_00E2F5A1: mov eax, var_48
  loc_00E2F5A4: push ebx
  loc_00E2F5A5: mov ecx, 00000002h
  loc_00E2F5AA: sub esp, 00000010h
  loc_00E2F5AD: mov var_9C, ecx
  loc_00E2F5B3: mov ebx, esp
  loc_00E2F5B5: mov var_94, 00000009h
  loc_00E2F5BF: mov edx, [eax]
  loc_00E2F5C1: push eax
  loc_00E2F5C2: mov [ebx], ecx
  loc_00E2F5C4: mov ecx, var_98
  loc_00E2F5CA: mov var_104, eax
  loc_00E2F5D0: mov [ebx+00000004h], ecx
  loc_00E2F5D3: mov ecx, var_94
  loc_00E2F5D9: mov [ebx+00000008h], ecx
  loc_00E2F5DC: mov ecx, var_90
  loc_00E2F5E2: mov [ebx+0000000Ch], ecx
  loc_00E2F5E5: call [edx+00000028h]
  loc_00E2F5E8: test eax, eax
  loc_00E2F5EA: fnclex
  loc_00E2F5EC: jge 00E2F603h
  loc_00E2F5EE: mov edx, var_104
  loc_00E2F5F4: push 00000028h
  loc_00E2F5F6: push 006E09E8h
  loc_00E2F5FB: push edx
  loc_00E2F5FC: push eax
  loc_00E2F5FD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F603: mov eax, var_4C
  loc_00E2F606: mov ecx, [esi]
  loc_00E2F608: push esi
  loc_00E2F609: mov var_10C, eax
  loc_00E2F60F: call [ecx+00000378h]
  loc_00E2F615: lea edx, var_34
  loc_00E2F618: push eax
  loc_00E2F619: push edx
  loc_00E2F61A: call edi
  loc_00E2F61C: mov ebx, eax
  loc_00E2F61E: lea ecx, var_28
  loc_00E2F621: push ecx
  loc_00E2F622: push ebx
  loc_00E2F623: mov eax, [ebx]
  loc_00E2F625: call [eax+000000A0h]
  loc_00E2F62B: test eax, eax
  loc_00E2F62D: fnclex
  loc_00E2F62F: jge 00E2F643h
  loc_00E2F631: push 000000A0h
  loc_00E2F636: push 006DCB70h
  loc_00E2F63B: push ebx
  loc_00E2F63C: push eax
  loc_00E2F63D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F643: mov edx, var_28
  loc_00E2F646: push edx
  loc_00E2F647: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2F64D: fsub st0, real8 ptr var_D8
  loc_00E2F653: mov ebx, var_10C
  loc_00E2F659: fadd st0, real8 ptr var_E0
  loc_00E2F65F: mov ecx, [ebx]
  loc_00E2F661: fstp real8 ptr var_A4
  loc_00E2F667: fnstsw ax
  loc_00E2F669: test al, 0Dh
  loc_00E2F66B: jnz 00E3052Ch
  loc_00E2F671: sub esp, 00000010h
  loc_00E2F674: mov eax, 00000005h
  loc_00E2F679: mov edx, esp
  loc_00E2F67B: push ebx
  loc_00E2F67C: mov [edx], eax
  loc_00E2F67E: mov eax, var_A8
  loc_00E2F684: mov [edx+00000004h], eax
  loc_00E2F687: mov eax, var_A4
  loc_00E2F68D: mov [edx+00000008h], eax
  loc_00E2F690: mov eax, var_A0
  loc_00E2F696: mov [edx+0000000Ch], eax
  loc_00E2F699: call [ecx+00000038h]
  loc_00E2F69C: test eax, eax
  loc_00E2F69E: fnclex
  loc_00E2F6A0: jge 00E2F6B1h
  loc_00E2F6A2: push 00000038h
  loc_00E2F6A4: push 006E09F8h
  loc_00E2F6A9: push ebx
  loc_00E2F6AA: push eax
  loc_00E2F6AB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F6B1: lea ecx, var_30
  loc_00E2F6B4: lea edx, var_2C
  loc_00E2F6B7: push ecx
  loc_00E2F6B8: lea eax, var_28
  loc_00E2F6BB: push edx
  loc_00E2F6BC: push eax
  loc_00E2F6BD: push 00000003h
  loc_00E2F6BF: call [004011BCh] ; __vbaFreeStrList
  loc_00E2F6C5: lea ecx, var_4C
  loc_00E2F6C8: lea edx, var_48
  loc_00E2F6CB: push ecx
  loc_00E2F6CC: lea eax, var_44
  loc_00E2F6CF: push edx
  loc_00E2F6D0: lea ecx, var_40
  loc_00E2F6D3: push eax
  loc_00E2F6D4: lea edx, var_3C
  loc_00E2F6D7: push ecx
  loc_00E2F6D8: lea eax, var_38
  loc_00E2F6DB: push edx
  loc_00E2F6DC: lea ecx, var_34
  loc_00E2F6DF: push eax
  loc_00E2F6E0: push ecx
  loc_00E2F6E1: push 00000007h
  loc_00E2F6E3: call [00401048h] ; __vbaFreeObjList
  loc_00E2F6E9: add esp, 00000030h
  loc_00E2F6EC: lea ecx, var_5C
  loc_00E2F6EF: call [00401024h] ; __vbaFreeVar
  loc_00E2F6F5: jmp 00E2F917h
  loc_00E2F6FA: call [edx+00000388h]
  loc_00E2F700: push eax
  loc_00E2F701: lea eax, var_34
  loc_00E2F704: push eax
  loc_00E2F705: call edi
  loc_00E2F707: mov ebx, eax
  loc_00E2F709: lea edx, var_28
  loc_00E2F70C: push edx
  loc_00E2F70D: push ebx
  loc_00E2F70E: mov ecx, [ebx]
  loc_00E2F710: call [ecx+00000050h]
  loc_00E2F713: test eax, eax
  loc_00E2F715: fnclex
  loc_00E2F717: jge 00E2F728h
  loc_00E2F719: push 00000050h
  loc_00E2F71B: push 006E0410h
  loc_00E2F720: push ebx
  loc_00E2F721: push eax
  loc_00E2F722: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F728: mov eax, var_28
  loc_00E2F72B: push eax
  loc_00E2F72C: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2F732: call [004010D8h] ; __vbaFpR8
  loc_00E2F738: fcomp real8 ptr [004022E8h]
  loc_00E2F73E: fnstsw ax
  loc_00E2F740: test ah, 41h
  loc_00E2F743: jz 00E2F74Ch
  loc_00E2F745: mov ebx, 00000001h
  loc_00E2F74A: jmp 00E2F74Eh
  loc_00E2F74C: xor ebx, ebx
  loc_00E2F74E: lea ecx, var_28
  loc_00E2F751: call [00401258h] ; __vbaFreeStr
  loc_00E2F757: lea ecx, var_34
  loc_00E2F75A: call [00401254h] ; __vbaFreeObj
  loc_00E2F760: neg ebx
  loc_00E2F762: test bx, bx
  loc_00E2F765: jz 00E2F899h
  loc_00E2F76B: mov ecx, [esi]
  loc_00E2F76D: push 006DCBD8h
  loc_00E2F772: push 00000000h
  loc_00E2F774: push 00000018h
  loc_00E2F776: push esi
  loc_00E2F777: call [ecx+000004A8h]
  loc_00E2F77D: lea edx, var_34
  loc_00E2F780: push eax
  loc_00E2F781: push edx
  loc_00E2F782: call edi
  loc_00E2F784: push eax
  loc_00E2F785: lea eax, var_5C
  loc_00E2F788: push eax
  loc_00E2F789: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2F78F: add esp, 00000010h
  loc_00E2F792: push eax
  loc_00E2F793: call [00401120h] ; __vbaCastObjVar
  loc_00E2F799: lea ecx, var_38
  loc_00E2F79C: push eax
  loc_00E2F79D: push ecx
  loc_00E2F79E: call edi
  loc_00E2F7A0: mov ebx, eax
  loc_00E2F7A2: lea eax, var_3C
  loc_00E2F7A5: push eax
  loc_00E2F7A6: push ebx
  loc_00E2F7A7: mov edx, [ebx]
  loc_00E2F7A9: call [edx+00000054h]
  loc_00E2F7AC: test eax, eax
  loc_00E2F7AE: fnclex
  loc_00E2F7B0: jge 00E2F7C1h
  loc_00E2F7B2: push 00000054h
  loc_00E2F7B4: push 006DCBE8h
  loc_00E2F7B9: push ebx
  loc_00E2F7BA: push eax
  loc_00E2F7BB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F7C1: lea ebx, var_40
  loc_00E2F7C4: mov eax, var_3C
  loc_00E2F7C7: push ebx
  loc_00E2F7C8: mov ecx, 00000002h
  loc_00E2F7CD: sub esp, 00000010h
  loc_00E2F7D0: mov var_9C, ecx
  loc_00E2F7D6: mov ebx, esp
  loc_00E2F7D8: mov var_94, 00000009h
  loc_00E2F7E2: mov edx, [eax]
  loc_00E2F7E4: push eax
  loc_00E2F7E5: mov [ebx], ecx
  loc_00E2F7E7: mov ecx, var_98
  loc_00E2F7ED: mov var_EC, eax
  loc_00E2F7F3: mov [ebx+00000004h], ecx
  loc_00E2F7F6: mov ecx, var_94
  loc_00E2F7FC: mov [ebx+00000008h], ecx
  loc_00E2F7FF: mov ecx, var_90
  loc_00E2F805: mov [ebx+0000000Ch], ecx
  loc_00E2F808: call [edx+00000028h]
  loc_00E2F80B: test eax, eax
  loc_00E2F80D: fnclex
  loc_00E2F80F: jge 00E2F826h
  loc_00E2F811: mov edx, var_EC
  loc_00E2F817: push 00000028h
  loc_00E2F819: push 006E09E8h
  loc_00E2F81E: push edx
  loc_00E2F81F: push eax
  loc_00E2F820: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F826: sub esp, 00000010h
  loc_00E2F829: mov eax, var_40
  loc_00E2F82C: mov ebx, esp
  loc_00E2F82E: mov ecx, 00000002h
  loc_00E2F833: mov edx, [eax]
  loc_00E2F835: push eax
  loc_00E2F836: mov [ebx], ecx
  loc_00E2F838: mov ecx, var_A8
  loc_00E2F83E: mov var_F4, eax
  loc_00E2F844: mov [ebx+00000004h], ecx
  loc_00E2F847: xor ecx, ecx
  loc_00E2F849: mov [ebx+00000008h], ecx
  loc_00E2F84C: mov ecx, var_A0
  loc_00E2F852: mov [ebx+0000000Ch], ecx
  loc_00E2F855: call [edx+00000038h]
  loc_00E2F858: test eax, eax
  loc_00E2F85A: fnclex
  loc_00E2F85C: jge 00E2F873h
  loc_00E2F85E: mov edx, var_F4
  loc_00E2F864: push 00000038h
  loc_00E2F866: push 006E09F8h
  loc_00E2F86B: push edx
  loc_00E2F86C: push eax
  loc_00E2F86D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F873: lea eax, var_40
  loc_00E2F876: lea ecx, var_3C
  loc_00E2F879: push eax
  loc_00E2F87A: lea edx, var_38
  loc_00E2F87D: push ecx
  loc_00E2F87E: lea eax, var_34
  loc_00E2F881: push edx
  loc_00E2F882: push eax
  loc_00E2F883: push 00000004h
  loc_00E2F885: call [00401048h] ; __vbaFreeObjList
  loc_00E2F88B: add esp, 00000014h
  loc_00E2F88E: lea ecx, var_5C
  loc_00E2F891: call [00401024h] ; __vbaFreeVar
  loc_00E2F897: jmp 00E2F917h
  loc_00E2F899: mov ecx, 80020004h
  loc_00E2F89E: mov eax, 0000000Ah
  loc_00E2F8A3: mov var_84, ecx
  loc_00E2F8A9: mov var_74, ecx
  loc_00E2F8AC: mov var_64, ecx
  loc_00E2F8AF: lea edx, var_9C
  loc_00E2F8B5: lea ecx, var_5C
  loc_00E2F8B8: mov var_8C, eax
  loc_00E2F8BE: mov var_7C, eax
  loc_00E2F8C1: mov var_6C, eax
  loc_00E2F8C4: mov var_94, 006E1B58h ; "Masih ada Error di kembalian !"
  loc_00E2F8CE: mov var_9C, 00000008h
  loc_00E2F8D8: call [004011F0h] ; __vbaVarDup
  loc_00E2F8DE: lea ecx, var_8C
  loc_00E2F8E4: lea edx, var_7C
  loc_00E2F8E7: push ecx
  loc_00E2F8E8: lea eax, var_6C
  loc_00E2F8EB: push edx
  loc_00E2F8EC: push eax
  loc_00E2F8ED: lea ecx, var_5C
  loc_00E2F8F0: push 00000000h
  loc_00E2F8F2: push ecx
  loc_00E2F8F3: call [004010A8h] ; rtcMsgBox
  loc_00E2F8F9: lea edx, var_8C
  loc_00E2F8FF: lea eax, var_7C
  loc_00E2F902: push edx
  loc_00E2F903: lea ecx, var_6C
  loc_00E2F906: push eax
  loc_00E2F907: lea edx, var_5C
  loc_00E2F90A: push ecx
  loc_00E2F90B: push edx
  loc_00E2F90C: push 00000004h
  loc_00E2F90E: call [00401038h] ; __vbaFreeVarList
  loc_00E2F914: add esp, 00000014h
  loc_00E2F917: mov eax, [esi]
  loc_00E2F919: push 006DCBD8h
  loc_00E2F91E: push 00000000h
  loc_00E2F920: push 00000018h
  loc_00E2F922: push esi
  loc_00E2F923: call [eax+000004A8h]
  loc_00E2F929: lea ecx, var_38
  loc_00E2F92C: push eax
  loc_00E2F92D: push ecx
  loc_00E2F92E: call edi
  loc_00E2F930: lea edx, var_5C
  loc_00E2F933: push eax
  loc_00E2F934: push edx
  loc_00E2F935: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2F93B: add esp, 00000010h
  loc_00E2F93E: push eax
  loc_00E2F93F: call [00401120h] ; __vbaCastObjVar
  loc_00E2F945: push eax
  loc_00E2F946: lea eax, var_3C
  loc_00E2F949: push eax
  loc_00E2F94A: call edi
  loc_00E2F94C: mov ebx, eax
  loc_00E2F94E: lea edx, var_40
  loc_00E2F951: push edx
  loc_00E2F952: push ebx
  loc_00E2F953: mov ecx, [ebx]
  loc_00E2F955: call [ecx+00000054h]
  loc_00E2F958: test eax, eax
  loc_00E2F95A: fnclex
  loc_00E2F95C: jge 00E2F96Dh
  loc_00E2F95E: push 00000054h
  loc_00E2F960: push 006DCBE8h
  loc_00E2F965: push ebx
  loc_00E2F966: push eax
  loc_00E2F967: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F96D: lea ebx, var_44
  loc_00E2F970: mov eax, var_40
  loc_00E2F973: push ebx
  loc_00E2F974: mov ecx, 00000002h
  loc_00E2F979: sub esp, 00000010h
  loc_00E2F97C: mov var_9C, ecx
  loc_00E2F982: mov ebx, esp
  loc_00E2F984: mov var_94, 0000000Ah
  loc_00E2F98E: mov edx, [eax]
  loc_00E2F990: push eax
  loc_00E2F991: mov [ebx], ecx
  loc_00E2F993: mov ecx, var_98
  loc_00E2F999: mov var_F4, eax
  loc_00E2F99F: mov [ebx+00000004h], ecx
  loc_00E2F9A2: mov ecx, var_94
  loc_00E2F9A8: mov [ebx+00000008h], ecx
  loc_00E2F9AB: mov ecx, var_90
  loc_00E2F9B1: mov [ebx+0000000Ch], ecx
  loc_00E2F9B4: call [edx+00000028h]
  loc_00E2F9B7: test eax, eax
  loc_00E2F9B9: fnclex
  loc_00E2F9BB: jge 00E2F9D2h
  loc_00E2F9BD: mov edx, var_F4
  loc_00E2F9C3: push 00000028h
  loc_00E2F9C5: push 006E09E8h
  loc_00E2F9CA: push edx
  loc_00E2F9CB: push eax
  loc_00E2F9CC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2F9D2: mov eax, var_44
  loc_00E2F9D5: mov ecx, [esi]
  loc_00E2F9D7: push esi
  loc_00E2F9D8: mov var_FC, eax
  loc_00E2F9DE: call [ecx+00000354h]
  loc_00E2F9E4: lea edx, var_34
  loc_00E2F9E7: push eax
  loc_00E2F9E8: push edx
  loc_00E2F9E9: call edi
  loc_00E2F9EB: mov ebx, eax
  loc_00E2F9ED: lea ecx, var_28
  loc_00E2F9F0: push ecx
  loc_00E2F9F1: push ebx
  loc_00E2F9F2: mov eax, [ebx]
  loc_00E2F9F4: call [eax+00000050h]
  loc_00E2F9F7: test eax, eax
  loc_00E2F9F9: fnclex
  loc_00E2F9FB: jge 00E2FA0Ch
  loc_00E2F9FD: push 00000050h
  loc_00E2F9FF: push 006E0410h
  loc_00E2FA04: push ebx
  loc_00E2FA05: push eax
  loc_00E2FA06: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FA0C: sub esp, 00000010h
  loc_00E2FA0F: mov eax, var_28
  loc_00E2FA12: mov ebx, esp
  loc_00E2FA14: mov ecx, 00000008h
  loc_00E2FA19: mov var_6C, ecx
  loc_00E2FA1C: mov edx, var_FC
  loc_00E2FA22: mov [ebx], ecx
  loc_00E2FA24: mov ecx, var_68
  loc_00E2FA27: mov var_64, eax
  loc_00E2FA2A: mov var_28, 00000000h
  loc_00E2FA31: mov edx, [edx]
  loc_00E2FA33: mov [ebx+00000004h], ecx
  loc_00E2FA36: mov [ebx+00000008h], eax
  loc_00E2FA39: mov eax, var_60
  loc_00E2FA3C: mov [ebx+0000000Ch], eax
  loc_00E2FA3F: mov ebx, var_FC
  loc_00E2FA45: push ebx
  loc_00E2FA46: call [edx+00000038h]
  loc_00E2FA49: test eax, eax
  loc_00E2FA4B: fnclex
  loc_00E2FA4D: jge 00E2FA5Eh
  loc_00E2FA4F: push 00000038h
  loc_00E2FA51: push 006E09F8h
  loc_00E2FA56: push ebx
  loc_00E2FA57: push eax
  loc_00E2FA58: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FA5E: lea ecx, var_44
  loc_00E2FA61: lea edx, var_40
  loc_00E2FA64: push ecx
  loc_00E2FA65: lea eax, var_3C
  loc_00E2FA68: push edx
  loc_00E2FA69: lea ecx, var_38
  loc_00E2FA6C: push eax
  loc_00E2FA6D: lea edx, var_34
  loc_00E2FA70: push ecx
  loc_00E2FA71: push edx
  loc_00E2FA72: push 00000005h
  loc_00E2FA74: call [00401048h] ; __vbaFreeObjList
  loc_00E2FA7A: lea eax, var_6C
  loc_00E2FA7D: lea ecx, var_5C
  loc_00E2FA80: push eax
  loc_00E2FA81: push ecx
  loc_00E2FA82: push 00000002h
  loc_00E2FA84: call [00401038h] ; __vbaFreeVarList
  loc_00E2FA8A: mov edx, [esi]
  loc_00E2FA8C: add esp, 00000024h
  loc_00E2FA8F: push 006DCBD8h
  loc_00E2FA94: push 00000000h
  loc_00E2FA96: push 00000018h
  loc_00E2FA98: push esi
  loc_00E2FA99: call [edx+000004A8h]
  loc_00E2FA9F: push eax
  loc_00E2FAA0: lea eax, var_38
  loc_00E2FAA3: push eax
  loc_00E2FAA4: call edi
  loc_00E2FAA6: lea ecx, var_5C
  loc_00E2FAA9: push eax
  loc_00E2FAAA: push ecx
  loc_00E2FAAB: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2FAB1: add esp, 00000010h
  loc_00E2FAB4: push eax
  loc_00E2FAB5: call [00401120h] ; __vbaCastObjVar
  loc_00E2FABB: lea edx, var_3C
  loc_00E2FABE: push eax
  loc_00E2FABF: push edx
  loc_00E2FAC0: call edi
  loc_00E2FAC2: mov ebx, eax
  loc_00E2FAC4: lea ecx, var_40
  loc_00E2FAC7: push ecx
  loc_00E2FAC8: push ebx
  loc_00E2FAC9: mov eax, [ebx]
  loc_00E2FACB: call [eax+00000054h]
  loc_00E2FACE: test eax, eax
  loc_00E2FAD0: fnclex
  loc_00E2FAD2: jge 00E2FAE3h
  loc_00E2FAD4: push 00000054h
  loc_00E2FAD6: push 006DCBE8h
  loc_00E2FADB: push ebx
  loc_00E2FADC: push eax
  loc_00E2FADD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FAE3: lea ebx, var_44
  loc_00E2FAE6: mov eax, var_40
  loc_00E2FAE9: push ebx
  loc_00E2FAEA: mov ecx, 00000002h
  loc_00E2FAEF: sub esp, 00000010h
  loc_00E2FAF2: mov var_9C, ecx
  loc_00E2FAF8: mov ebx, esp
  loc_00E2FAFA: mov var_94, 0000000Bh
  loc_00E2FB04: mov edx, [eax]
  loc_00E2FB06: push eax
  loc_00E2FB07: mov [ebx], ecx
  loc_00E2FB09: mov ecx, var_98
  loc_00E2FB0F: mov var_F4, eax
  loc_00E2FB15: mov [ebx+00000004h], ecx
  loc_00E2FB18: mov ecx, var_94
  loc_00E2FB1E: mov [ebx+00000008h], ecx
  loc_00E2FB21: mov ecx, var_90
  loc_00E2FB27: mov [ebx+0000000Ch], ecx
  loc_00E2FB2A: call [edx+00000028h]
  loc_00E2FB2D: test eax, eax
  loc_00E2FB2F: fnclex
  loc_00E2FB31: jge 00E2FB48h
  loc_00E2FB33: mov edx, var_F4
  loc_00E2FB39: push 00000028h
  loc_00E2FB3B: push 006E09E8h
  loc_00E2FB40: push edx
  loc_00E2FB41: push eax
  loc_00E2FB42: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FB48: mov eax, var_44
  loc_00E2FB4B: mov ecx, [esi]
  loc_00E2FB4D: push esi
  loc_00E2FB4E: mov var_FC, eax
  loc_00E2FB54: call [ecx+00000304h]
  loc_00E2FB5A: lea edx, var_34
  loc_00E2FB5D: push eax
  loc_00E2FB5E: push edx
  loc_00E2FB5F: call edi
  loc_00E2FB61: mov ebx, eax
  loc_00E2FB63: lea ecx, var_28
  loc_00E2FB66: push ecx
  loc_00E2FB67: push ebx
  loc_00E2FB68: mov eax, [ebx]
  loc_00E2FB6A: call [eax+000000A0h]
  loc_00E2FB70: test eax, eax
  loc_00E2FB72: fnclex
  loc_00E2FB74: jge 00E2FB88h
  loc_00E2FB76: push 000000A0h
  loc_00E2FB7B: push 006DCB70h
  loc_00E2FB80: push ebx
  loc_00E2FB81: push eax
  loc_00E2FB82: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FB88: sub esp, 00000010h
  loc_00E2FB8B: mov eax, var_28
  loc_00E2FB8E: mov ebx, esp
  loc_00E2FB90: mov ecx, 00000008h
  loc_00E2FB95: mov var_6C, ecx
  loc_00E2FB98: mov edx, var_FC
  loc_00E2FB9E: mov [ebx], ecx
  loc_00E2FBA0: mov ecx, var_68
  loc_00E2FBA3: mov var_64, eax
  loc_00E2FBA6: mov var_28, 00000000h
  loc_00E2FBAD: mov edx, [edx]
  loc_00E2FBAF: mov [ebx+00000004h], ecx
  loc_00E2FBB2: mov [ebx+00000008h], eax
  loc_00E2FBB5: mov eax, var_60
  loc_00E2FBB8: mov [ebx+0000000Ch], eax
  loc_00E2FBBB: mov ebx, var_FC
  loc_00E2FBC1: push ebx
  loc_00E2FBC2: call [edx+00000038h]
  loc_00E2FBC5: test eax, eax
  loc_00E2FBC7: fnclex
  loc_00E2FBC9: jge 00E2FBDAh
  loc_00E2FBCB: push 00000038h
  loc_00E2FBCD: push 006E09F8h
  loc_00E2FBD2: push ebx
  loc_00E2FBD3: push eax
  loc_00E2FBD4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FBDA: lea ecx, var_44
  loc_00E2FBDD: lea edx, var_40
  loc_00E2FBE0: push ecx
  loc_00E2FBE1: lea eax, var_3C
  loc_00E2FBE4: push edx
  loc_00E2FBE5: lea ecx, var_38
  loc_00E2FBE8: push eax
  loc_00E2FBE9: lea edx, var_34
  loc_00E2FBEC: push ecx
  loc_00E2FBED: push edx
  loc_00E2FBEE: push 00000005h
  loc_00E2FBF0: call [00401048h] ; __vbaFreeObjList
  loc_00E2FBF6: lea eax, var_6C
  loc_00E2FBF9: lea ecx, var_5C
  loc_00E2FBFC: push eax
  loc_00E2FBFD: push ecx
  loc_00E2FBFE: push 00000002h
  loc_00E2FC00: call [00401038h] ; __vbaFreeVarList
  loc_00E2FC06: mov edx, [esi]
  loc_00E2FC08: add esp, 00000024h
  loc_00E2FC0B: push 006DCBD8h
  loc_00E2FC10: push 00000000h
  loc_00E2FC12: push 00000018h
  loc_00E2FC14: push esi
  loc_00E2FC15: call [edx+000004A8h]
  loc_00E2FC1B: push eax
  loc_00E2FC1C: lea eax, var_38
  loc_00E2FC1F: push eax
  loc_00E2FC20: call edi
  loc_00E2FC22: lea ecx, var_5C
  loc_00E2FC25: push eax
  loc_00E2FC26: push ecx
  loc_00E2FC27: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2FC2D: add esp, 00000010h
  loc_00E2FC30: push eax
  loc_00E2FC31: call [00401120h] ; __vbaCastObjVar
  loc_00E2FC37: lea edx, var_3C
  loc_00E2FC3A: push eax
  loc_00E2FC3B: push edx
  loc_00E2FC3C: call edi
  loc_00E2FC3E: mov ebx, eax
  loc_00E2FC40: lea ecx, var_40
  loc_00E2FC43: push ecx
  loc_00E2FC44: push ebx
  loc_00E2FC45: mov eax, [ebx]
  loc_00E2FC47: call [eax+00000054h]
  loc_00E2FC4A: test eax, eax
  loc_00E2FC4C: fnclex
  loc_00E2FC4E: jge 00E2FC5Fh
  loc_00E2FC50: push 00000054h
  loc_00E2FC52: push 006DCBE8h
  loc_00E2FC57: push ebx
  loc_00E2FC58: push eax
  loc_00E2FC59: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FC5F: lea ebx, var_44
  loc_00E2FC62: mov eax, var_40
  loc_00E2FC65: push ebx
  loc_00E2FC66: mov ecx, 00000002h
  loc_00E2FC6B: sub esp, 00000010h
  loc_00E2FC6E: mov var_9C, ecx
  loc_00E2FC74: mov ebx, esp
  loc_00E2FC76: mov var_94, 0000000Ch
  loc_00E2FC80: mov edx, [eax]
  loc_00E2FC82: push eax
  loc_00E2FC83: mov [ebx], ecx
  loc_00E2FC85: mov ecx, var_98
  loc_00E2FC8B: mov var_F4, eax
  loc_00E2FC91: mov [ebx+00000004h], ecx
  loc_00E2FC94: mov ecx, var_94
  loc_00E2FC9A: mov [ebx+00000008h], ecx
  loc_00E2FC9D: mov ecx, var_90
  loc_00E2FCA3: mov [ebx+0000000Ch], ecx
  loc_00E2FCA6: call [edx+00000028h]
  loc_00E2FCA9: test eax, eax
  loc_00E2FCAB: fnclex
  loc_00E2FCAD: jge 00E2FCC4h
  loc_00E2FCAF: mov edx, var_F4
  loc_00E2FCB5: push 00000028h
  loc_00E2FCB7: push 006E09E8h
  loc_00E2FCBC: push edx
  loc_00E2FCBD: push eax
  loc_00E2FCBE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FCC4: mov eax, var_44
  loc_00E2FCC7: mov ecx, [esi]
  loc_00E2FCC9: push esi
  loc_00E2FCCA: mov var_FC, eax
  loc_00E2FCD0: call [ecx+00000354h]
  loc_00E2FCD6: lea edx, var_34
  loc_00E2FCD9: push eax
  loc_00E2FCDA: push edx
  loc_00E2FCDB: call edi
  loc_00E2FCDD: mov ebx, eax
  loc_00E2FCDF: lea ecx, var_28
  loc_00E2FCE2: push ecx
  loc_00E2FCE3: push ebx
  loc_00E2FCE4: mov eax, [ebx]
  loc_00E2FCE6: call [eax+00000050h]
  loc_00E2FCE9: test eax, eax
  loc_00E2FCEB: fnclex
  loc_00E2FCED: jge 00E2FCFEh
  loc_00E2FCEF: push 00000050h
  loc_00E2FCF1: push 006E0410h
  loc_00E2FCF6: push ebx
  loc_00E2FCF7: push eax
  loc_00E2FCF8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FCFE: sub esp, 00000010h
  loc_00E2FD01: mov eax, var_28
  loc_00E2FD04: mov ebx, esp
  loc_00E2FD06: mov ecx, 00000008h
  loc_00E2FD0B: mov var_6C, ecx
  loc_00E2FD0E: mov edx, var_FC
  loc_00E2FD14: mov [ebx], ecx
  loc_00E2FD16: mov ecx, var_68
  loc_00E2FD19: mov var_64, eax
  loc_00E2FD1C: mov var_28, 00000000h
  loc_00E2FD23: mov edx, [edx]
  loc_00E2FD25: mov [ebx+00000004h], ecx
  loc_00E2FD28: mov [ebx+00000008h], eax
  loc_00E2FD2B: mov eax, var_60
  loc_00E2FD2E: mov [ebx+0000000Ch], eax
  loc_00E2FD31: mov ebx, var_FC
  loc_00E2FD37: push ebx
  loc_00E2FD38: call [edx+00000038h]
  loc_00E2FD3B: test eax, eax
  loc_00E2FD3D: fnclex
  loc_00E2FD3F: jge 00E2FD50h
  loc_00E2FD41: push 00000038h
  loc_00E2FD43: push 006E09F8h
  loc_00E2FD48: push ebx
  loc_00E2FD49: push eax
  loc_00E2FD4A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FD50: lea ecx, var_44
  loc_00E2FD53: lea edx, var_40
  loc_00E2FD56: push ecx
  loc_00E2FD57: lea eax, var_3C
  loc_00E2FD5A: push edx
  loc_00E2FD5B: lea ecx, var_38
  loc_00E2FD5E: push eax
  loc_00E2FD5F: lea edx, var_34
  loc_00E2FD62: push ecx
  loc_00E2FD63: push edx
  loc_00E2FD64: push 00000005h
  loc_00E2FD66: call [00401048h] ; __vbaFreeObjList
  loc_00E2FD6C: lea eax, var_6C
  loc_00E2FD6F: lea ecx, var_5C
  loc_00E2FD72: push eax
  loc_00E2FD73: push ecx
  loc_00E2FD74: push 00000002h
  loc_00E2FD76: call [00401038h] ; __vbaFreeVarList
  loc_00E2FD7C: mov edx, [esi]
  loc_00E2FD7E: add esp, 00000024h
  loc_00E2FD81: push esi
  loc_00E2FD82: call [edx+000003B0h]
  loc_00E2FD88: push eax
  loc_00E2FD89: lea eax, var_34
  loc_00E2FD8C: push eax
  loc_00E2FD8D: call edi
  loc_00E2FD8F: mov ebx, eax
  loc_00E2FD91: lea edx, var_28
  loc_00E2FD94: push edx
  loc_00E2FD95: push ebx
  loc_00E2FD96: mov ecx, [ebx]
  loc_00E2FD98: call [ecx+00000050h]
  loc_00E2FD9B: test eax, eax
  loc_00E2FD9D: fnclex
  loc_00E2FD9F: jge 00E2FDB0h
  loc_00E2FDA1: push 00000050h
  loc_00E2FDA3: push 006E0410h
  loc_00E2FDA8: push ebx
  loc_00E2FDA9: push eax
  loc_00E2FDAA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FDB0: mov eax, var_28
  loc_00E2FDB3: push eax
  loc_00E2FDB4: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2FDBA: call [004010D8h] ; __vbaFpR8
  loc_00E2FDC0: fcomp real8 ptr [004022E8h]
  loc_00E2FDC6: fnstsw ax
  loc_00E2FDC8: test ah, 01h
  loc_00E2FDCB: jz 00E2FDD4h
  loc_00E2FDCD: mov ebx, 00000001h
  loc_00E2FDD2: jmp 00E2FDD6h
  loc_00E2FDD4: xor ebx, ebx
  loc_00E2FDD6: lea ecx, var_28
  loc_00E2FDD9: call [00401258h] ; __vbaFreeStr
  loc_00E2FDDF: lea ecx, var_34
  loc_00E2FDE2: call [00401254h] ; __vbaFreeObj
  loc_00E2FDE8: mov ecx, [esi]
  loc_00E2FDEA: push 006DCBD8h
  loc_00E2FDEF: neg ebx
  loc_00E2FDF1: push 00000000h
  loc_00E2FDF3: push 00000018h
  loc_00E2FDF5: test bx, bx
  loc_00E2FDF8: push esi
  loc_00E2FDF9: jz 00E2FF24h
  loc_00E2FDFF: call [ecx+000004A8h]
  loc_00E2FE05: lea edx, var_34
  loc_00E2FE08: push eax
  loc_00E2FE09: push edx
  loc_00E2FE0A: call edi
  loc_00E2FE0C: push eax
  loc_00E2FE0D: lea eax, var_5C
  loc_00E2FE10: push eax
  loc_00E2FE11: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2FE17: add esp, 00000010h
  loc_00E2FE1A: push eax
  loc_00E2FE1B: call [00401120h] ; __vbaCastObjVar
  loc_00E2FE21: lea ecx, var_38
  loc_00E2FE24: push eax
  loc_00E2FE25: push ecx
  loc_00E2FE26: call edi
  loc_00E2FE28: mov ebx, eax
  loc_00E2FE2A: lea eax, var_3C
  loc_00E2FE2D: push eax
  loc_00E2FE2E: push ebx
  loc_00E2FE2F: mov edx, [ebx]
  loc_00E2FE31: call [edx+00000054h]
  loc_00E2FE34: test eax, eax
  loc_00E2FE36: fnclex
  loc_00E2FE38: jge 00E2FE49h
  loc_00E2FE3A: push 00000054h
  loc_00E2FE3C: push 006DCBE8h
  loc_00E2FE41: push ebx
  loc_00E2FE42: push eax
  loc_00E2FE43: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FE49: lea ebx, var_40
  loc_00E2FE4C: mov eax, var_3C
  loc_00E2FE4F: push ebx
  loc_00E2FE50: mov ecx, 00000002h
  loc_00E2FE55: sub esp, 00000010h
  loc_00E2FE58: mov var_9C, ecx
  loc_00E2FE5E: mov ebx, esp
  loc_00E2FE60: mov var_94, 0000000Dh
  loc_00E2FE6A: mov edx, [eax]
  loc_00E2FE6C: push eax
  loc_00E2FE6D: mov [ebx], ecx
  loc_00E2FE6F: mov ecx, var_98
  loc_00E2FE75: mov var_EC, eax
  loc_00E2FE7B: mov [ebx+00000004h], ecx
  loc_00E2FE7E: mov ecx, var_94
  loc_00E2FE84: mov [ebx+00000008h], ecx
  loc_00E2FE87: mov ecx, var_90
  loc_00E2FE8D: mov [ebx+0000000Ch], ecx
  loc_00E2FE90: call [edx+00000028h]
  loc_00E2FE93: test eax, eax
  loc_00E2FE95: fnclex
  loc_00E2FE97: jge 00E2FEAEh
  loc_00E2FE99: mov edx, var_EC
  loc_00E2FE9F: push 00000028h
  loc_00E2FEA1: push 006E09E8h
  loc_00E2FEA6: push edx
  loc_00E2FEA7: push eax
  loc_00E2FEA8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FEAE: sub esp, 00000010h
  loc_00E2FEB1: mov eax, var_40
  loc_00E2FEB4: mov ebx, esp
  loc_00E2FEB6: mov ecx, 00000002h
  loc_00E2FEBB: mov edx, [eax]
  loc_00E2FEBD: push eax
  loc_00E2FEBE: mov [ebx], ecx
  loc_00E2FEC0: mov ecx, var_A8
  loc_00E2FEC6: mov var_F4, eax
  loc_00E2FECC: mov [ebx+00000004h], ecx
  loc_00E2FECF: xor ecx, ecx
  loc_00E2FED1: mov [ebx+00000008h], ecx
  loc_00E2FED4: mov ecx, var_A0
  loc_00E2FEDA: mov [ebx+0000000Ch], ecx
  loc_00E2FEDD: call [edx+00000038h]
  loc_00E2FEE0: test eax, eax
  loc_00E2FEE2: fnclex
  loc_00E2FEE4: jge 00E2FEFBh
  loc_00E2FEE6: mov edx, var_F4
  loc_00E2FEEC: push 00000038h
  loc_00E2FEEE: push 006E09F8h
  loc_00E2FEF3: push edx
  loc_00E2FEF4: push eax
  loc_00E2FEF5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FEFB: lea eax, var_40
  loc_00E2FEFE: lea ecx, var_3C
  loc_00E2FF01: push eax
  loc_00E2FF02: lea edx, var_38
  loc_00E2FF05: push ecx
  loc_00E2FF06: lea eax, var_34
  loc_00E2FF09: push edx
  loc_00E2FF0A: push eax
  loc_00E2FF0B: push 00000004h
  loc_00E2FF0D: call [00401048h] ; __vbaFreeObjList
  loc_00E2FF13: add esp, 00000014h
  loc_00E2FF16: lea ecx, var_5C
  loc_00E2FF19: call [00401024h] ; __vbaFreeVar
  loc_00E2FF1F: jmp 00E3008Eh
  loc_00E2FF24: call [ecx+000004A8h]
  loc_00E2FF2A: lea edx, var_38
  loc_00E2FF2D: push eax
  loc_00E2FF2E: push edx
  loc_00E2FF2F: call edi
  loc_00E2FF31: push eax
  loc_00E2FF32: lea eax, var_5C
  loc_00E2FF35: push eax
  loc_00E2FF36: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2FF3C: add esp, 00000010h
  loc_00E2FF3F: push eax
  loc_00E2FF40: call [00401120h] ; __vbaCastObjVar
  loc_00E2FF46: lea ecx, var_3C
  loc_00E2FF49: push eax
  loc_00E2FF4A: push ecx
  loc_00E2FF4B: call edi
  loc_00E2FF4D: mov ebx, eax
  loc_00E2FF4F: lea eax, var_40
  loc_00E2FF52: push eax
  loc_00E2FF53: push ebx
  loc_00E2FF54: mov edx, [ebx]
  loc_00E2FF56: call [edx+00000054h]
  loc_00E2FF59: test eax, eax
  loc_00E2FF5B: fnclex
  loc_00E2FF5D: jge 00E2FF6Eh
  loc_00E2FF5F: push 00000054h
  loc_00E2FF61: push 006DCBE8h
  loc_00E2FF66: push ebx
  loc_00E2FF67: push eax
  loc_00E2FF68: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FF6E: lea ebx, var_44
  loc_00E2FF71: mov eax, var_40
  loc_00E2FF74: push ebx
  loc_00E2FF75: mov ecx, 00000002h
  loc_00E2FF7A: sub esp, 00000010h
  loc_00E2FF7D: mov var_9C, ecx
  loc_00E2FF83: mov ebx, esp
  loc_00E2FF85: mov var_94, 0000000Dh
  loc_00E2FF8F: mov edx, [eax]
  loc_00E2FF91: push eax
  loc_00E2FF92: mov [ebx], ecx
  loc_00E2FF94: mov ecx, var_98
  loc_00E2FF9A: mov var_F4, eax
  loc_00E2FFA0: mov [ebx+00000004h], ecx
  loc_00E2FFA3: mov ecx, var_94
  loc_00E2FFA9: mov [ebx+00000008h], ecx
  loc_00E2FFAC: mov ecx, var_90
  loc_00E2FFB2: mov [ebx+0000000Ch], ecx
  loc_00E2FFB5: call [edx+00000028h]
  loc_00E2FFB8: test eax, eax
  loc_00E2FFBA: fnclex
  loc_00E2FFBC: jge 00E2FFD3h
  loc_00E2FFBE: mov edx, var_F4
  loc_00E2FFC4: push 00000028h
  loc_00E2FFC6: push 006E09E8h
  loc_00E2FFCB: push edx
  loc_00E2FFCC: push eax
  loc_00E2FFCD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2FFD3: mov eax, var_44
  loc_00E2FFD6: mov ecx, [esi]
  loc_00E2FFD8: push esi
  loc_00E2FFD9: mov var_FC, eax
  loc_00E2FFDF: call [ecx+000003B0h]
  loc_00E2FFE5: lea edx, var_34
  loc_00E2FFE8: push eax
  loc_00E2FFE9: push edx
  loc_00E2FFEA: call edi
  loc_00E2FFEC: mov ebx, eax
  loc_00E2FFEE: lea ecx, var_28
  loc_00E2FFF1: push ecx
  loc_00E2FFF2: push ebx
  loc_00E2FFF3: mov eax, [ebx]
  loc_00E2FFF5: call [eax+00000050h]
  loc_00E2FFF8: test eax, eax
  loc_00E2FFFA: fnclex
  loc_00E2FFFC: jge 00E3000Dh
  loc_00E2FFFE: push 00000050h
  loc_00E30000: push 006E0410h
  loc_00E30005: push ebx
  loc_00E30006: push eax
  loc_00E30007: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3000D: sub esp, 00000010h
  loc_00E30010: mov eax, var_28
  loc_00E30013: mov ebx, esp
  loc_00E30015: mov ecx, 00000008h
  loc_00E3001A: mov var_6C, ecx
  loc_00E3001D: mov edx, var_FC
  loc_00E30023: mov [ebx], ecx
  loc_00E30025: mov ecx, var_68
  loc_00E30028: mov var_64, eax
  loc_00E3002B: mov var_28, 00000000h
  loc_00E30032: mov edx, [edx]
  loc_00E30034: mov [ebx+00000004h], ecx
  loc_00E30037: mov [ebx+00000008h], eax
  loc_00E3003A: mov eax, var_60
  loc_00E3003D: mov [ebx+0000000Ch], eax
  loc_00E30040: mov ebx, var_FC
  loc_00E30046: push ebx
  loc_00E30047: call [edx+00000038h]
  loc_00E3004A: test eax, eax
  loc_00E3004C: fnclex
  loc_00E3004E: jge 00E3005Fh
  loc_00E30050: push 00000038h
  loc_00E30052: push 006E09F8h
  loc_00E30057: push ebx
  loc_00E30058: push eax
  loc_00E30059: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3005F: lea ecx, var_44
  loc_00E30062: lea edx, var_40
  loc_00E30065: push ecx
  loc_00E30066: lea eax, var_3C
  loc_00E30069: push edx
  loc_00E3006A: lea ecx, var_38
  loc_00E3006D: push eax
  loc_00E3006E: lea edx, var_34
  loc_00E30071: push ecx
  loc_00E30072: push edx
  loc_00E30073: push 00000005h
  loc_00E30075: call [00401048h] ; __vbaFreeObjList
  loc_00E3007B: lea eax, var_6C
  loc_00E3007E: lea ecx, var_5C
  loc_00E30081: push eax
  loc_00E30082: push ecx
  loc_00E30083: push 00000002h
  loc_00E30085: call [00401038h] ; __vbaFreeVarList
  loc_00E3008B: add esp, 00000024h
  loc_00E3008E: sub esp, 00000010h
  loc_00E30091: mov ecx, 0000000Bh
  loc_00E30096: mov edx, esp
  loc_00E30098: mov var_9C, ecx
  loc_00E3009E: xor eax, eax
  loc_00E300A0: push 80010007h
  loc_00E300A5: mov [edx], ecx
  loc_00E300A7: mov ecx, var_98
  loc_00E300AD: mov var_94, eax
  loc_00E300B3: push esi
  loc_00E300B4: mov [edx+00000004h], ecx
  loc_00E300B7: mov ecx, [esi]
  loc_00E300B9: mov [edx+00000008h], eax
  loc_00E300BC: mov eax, var_90
  loc_00E300C2: mov [edx+0000000Ch], eax
  loc_00E300C5: call [ecx+00000400h]
  loc_00E300CB: lea edx, var_34
  loc_00E300CE: push eax
  loc_00E300CF: push edx
  loc_00E300D0: call edi
  loc_00E300D2: push eax
  loc_00E300D3: call [00401238h] ; __vbaLateIdSt
  loc_00E300D9: lea ecx, var_34
  loc_00E300DC: call [00401254h] ; __vbaFreeObj
  loc_00E300E2: mov eax, [esi]
  loc_00E300E4: push 006DCBD8h
  loc_00E300E9: push 00000000h
  loc_00E300EB: push 00000018h
  loc_00E300ED: push esi
  loc_00E300EE: call [eax+000004A8h]
  loc_00E300F4: lea ecx, var_38
  loc_00E300F7: push eax
  loc_00E300F8: push ecx
  loc_00E300F9: call edi
  loc_00E300FB: lea edx, var_5C
  loc_00E300FE: push eax
  loc_00E300FF: push edx
  loc_00E30100: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E30106: add esp, 00000010h
  loc_00E30109: push eax
  loc_00E3010A: call [00401120h] ; __vbaCastObjVar
  loc_00E30110: push eax
  loc_00E30111: lea eax, var_3C
  loc_00E30114: push eax
  loc_00E30115: call edi
  loc_00E30117: mov ebx, eax
  loc_00E30119: lea edx, var_40
  loc_00E3011C: push edx
  loc_00E3011D: push ebx
  loc_00E3011E: mov ecx, [ebx]
  loc_00E30120: call [ecx+00000054h]
  loc_00E30123: test eax, eax
  loc_00E30125: fnclex
  loc_00E30127: jge 00E30138h
  loc_00E30129: push 00000054h
  loc_00E3012B: push 006DCBE8h
  loc_00E30130: push ebx
  loc_00E30131: push eax
  loc_00E30132: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30138: lea ebx, var_44
  loc_00E3013B: mov eax, var_40
  loc_00E3013E: push ebx
  loc_00E3013F: mov ecx, 00000002h
  loc_00E30144: sub esp, 00000010h
  loc_00E30147: mov var_9C, ecx
  loc_00E3014D: mov ebx, esp
  loc_00E3014F: mov var_94, 0000000Eh
  loc_00E30159: mov edx, [eax]
  loc_00E3015B: push eax
  loc_00E3015C: mov [ebx], ecx
  loc_00E3015E: mov ecx, var_98
  loc_00E30164: mov var_F4, eax
  loc_00E3016A: mov [ebx+00000004h], ecx
  loc_00E3016D: mov ecx, var_94
  loc_00E30173: mov [ebx+00000008h], ecx
  loc_00E30176: mov ecx, var_90
  loc_00E3017C: mov [ebx+0000000Ch], ecx
  loc_00E3017F: call [edx+00000028h]
  loc_00E30182: test eax, eax
  loc_00E30184: fnclex
  loc_00E30186: jge 00E3019Dh
  loc_00E30188: mov edx, var_F4
  loc_00E3018E: push 00000028h
  loc_00E30190: push 006E09E8h
  loc_00E30195: push edx
  loc_00E30196: push eax
  loc_00E30197: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3019D: mov eax, var_44
  loc_00E301A0: mov ecx, [esi]
  loc_00E301A2: push esi
  loc_00E301A3: mov var_FC, eax
  loc_00E301A9: call [ecx+00000300h]
  loc_00E301AF: lea edx, var_34
  loc_00E301B2: push eax
  loc_00E301B3: push edx
  loc_00E301B4: call edi
  loc_00E301B6: mov ebx, eax
  loc_00E301B8: lea ecx, var_28
  loc_00E301BB: push ecx
  loc_00E301BC: push ebx
  loc_00E301BD: mov eax, [ebx]
  loc_00E301BF: call [eax+000000A0h]
  loc_00E301C5: test eax, eax
  loc_00E301C7: fnclex
  loc_00E301C9: jge 00E301DDh
  loc_00E301CB: push 000000A0h
  loc_00E301D0: push 006DCB70h
  loc_00E301D5: push ebx
  loc_00E301D6: push eax
  loc_00E301D7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E301DD: sub esp, 00000010h
  loc_00E301E0: mov eax, var_28
  loc_00E301E3: mov ebx, esp
  loc_00E301E5: mov ecx, 00000008h
  loc_00E301EA: mov var_6C, ecx
  loc_00E301ED: mov edx, var_FC
  loc_00E301F3: mov [ebx], ecx
  loc_00E301F5: mov ecx, var_68
  loc_00E301F8: mov var_64, eax
  loc_00E301FB: mov var_28, 00000000h
  loc_00E30202: mov edx, [edx]
  loc_00E30204: mov [ebx+00000004h], ecx
  loc_00E30207: mov [ebx+00000008h], eax
  loc_00E3020A: mov eax, var_60
  loc_00E3020D: mov [ebx+0000000Ch], eax
  loc_00E30210: mov ebx, var_FC
  loc_00E30216: push ebx
  loc_00E30217: call [edx+00000038h]
  loc_00E3021A: test eax, eax
  loc_00E3021C: fnclex
  loc_00E3021E: jge 00E3022Fh
  loc_00E30220: push 00000038h
  loc_00E30222: push 006E09F8h
  loc_00E30227: push ebx
  loc_00E30228: push eax
  loc_00E30229: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E3022F: lea ecx, var_44
  loc_00E30232: lea edx, var_40
  loc_00E30235: push ecx
  loc_00E30236: lea eax, var_3C
  loc_00E30239: push edx
  loc_00E3023A: lea ecx, var_38
  loc_00E3023D: push eax
  loc_00E3023E: lea edx, var_34
  loc_00E30241: push ecx
  loc_00E30242: push edx
  loc_00E30243: push 00000005h
  loc_00E30245: call [00401048h] ; __vbaFreeObjList
  loc_00E3024B: lea eax, var_6C
  loc_00E3024E: lea ecx, var_5C
  loc_00E30251: push eax
  loc_00E30252: push ecx
  loc_00E30253: push 00000002h
  loc_00E30255: call [00401038h] ; __vbaFreeVarList
  loc_00E3025B: mov edx, [esi]
  loc_00E3025D: add esp, 00000024h
  loc_00E30260: push 006DCBD8h
  loc_00E30265: push 00000000h
  loc_00E30267: push 00000018h
  loc_00E30269: push esi
  loc_00E3026A: call [edx+000004A8h]
  loc_00E30270: push eax
  loc_00E30271: lea eax, var_34
  loc_00E30274: push eax
  loc_00E30275: call edi
  loc_00E30277: lea ecx, var_5C
  loc_00E3027A: push eax
  loc_00E3027B: push ecx
  loc_00E3027C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E30282: add esp, 00000010h
  loc_00E30285: push eax
  loc_00E30286: call [00401120h] ; __vbaCastObjVar
  loc_00E3028C: lea edx, var_38
  loc_00E3028F: push eax
  loc_00E30290: push edx
  loc_00E30291: call edi
  loc_00E30293: sub esp, 00000010h
  loc_00E30296: mov ebx, eax
  loc_00E30298: mov edx, esp
  loc_00E3029A: mov eax, 0000000Ah
  loc_00E3029F: mov var_9C, eax
  loc_00E302A5: sub esp, 00000010h
  loc_00E302A8: mov [edx], eax
  loc_00E302AA: mov eax, var_A8
  loc_00E302B0: mov var_94, 80020004h
  loc_00E302BA: mov ecx, [ebx]
  loc_00E302BC: mov [edx+00000004h], eax
  loc_00E302BF: mov eax, 80020004h
  loc_00E302C4: mov [edx+00000008h], eax
  loc_00E302C7: mov eax, var_A0
  loc_00E302CD: mov [edx+0000000Ch], eax
  loc_00E302D0: mov eax, var_9C
  loc_00E302D6: mov edx, esp
  loc_00E302D8: push ebx
  loc_00E302D9: mov [edx], eax
  loc_00E302DB: mov eax, var_98
  loc_00E302E1: mov [edx+00000004h], eax
  loc_00E302E4: mov eax, var_94
  loc_00E302EA: mov [edx+00000008h], eax
  loc_00E302ED: mov eax, var_90
  loc_00E302F3: mov [edx+0000000Ch], eax
  loc_00E302F6: call [ecx+000000ACh]
  loc_00E302FC: test eax, eax
  loc_00E302FE: fnclex
  loc_00E30300: jge 00E30314h
  loc_00E30302: push 000000ACh
  loc_00E30307: push 006DCBE8h
  loc_00E3030C: push ebx
  loc_00E3030D: push eax
  loc_00E3030E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E30314: lea ecx, var_38
  loc_00E30317: lea edx, var_34
  loc_00E3031A: push ecx
  loc_00E3031B: push edx
  loc_00E3031C: push 00000002h
  loc_00E3031E: call [00401048h] ; __vbaFreeObjList
  loc_00E30324: add esp, 0000000Ch
  loc_00E30327: lea ecx, var_5C
  loc_00E3032A: call [00401024h] ; __vbaFreeVar
  loc_00E30330: mov eax, [esi]
  loc_00E30332: push 00000000h
  loc_00E30334: push 00000019h
  loc_00E30336: push esi
  loc_00E30337: call [eax+000004A8h]
  loc_00E3033D: lea ecx, var_34
  loc_00E30340: push eax
  loc_00E30341: push ecx
  loc_00E30342: call edi
  loc_00E30344: mov ebx, [00401030h] ; __vbaLateIdCall
  loc_00E3034A: push eax
  loc_00E3034B: call ebx
  loc_00E3034D: add esp, 0000000Ch
  loc_00E30350: lea ecx, var_34
  loc_00E30353: call [00401254h] ; __vbaFreeObj
  loc_00E30359: lea edx, var_9C
  loc_00E3035F: lea ecx, var_24
  loc_00E30362: mov var_94, 006E19B0h ; " select * from lc"
  loc_00E3036C: mov var_9C, 00000008h
  loc_00E30376: call [00401204h] ; __vbaVarCopy
  loc_00E3037C: lea edx, var_24
  loc_00E3037F: push edx
  loc_00E30380: call [00401230h] ; __vbaStrVarCopy
  loc_00E30386: sub esp, 00000010h
  loc_00E30389: mov ecx, 00000008h
  loc_00E3038E: mov edx, esp
  loc_00E30390: mov var_5C, ecx
  loc_00E30393: mov var_54, eax
  loc_00E30396: push 0000000Eh
  loc_00E30398: mov [edx], ecx
  loc_00E3039A: mov ecx, var_58
  loc_00E3039D: push esi
  loc_00E3039E: mov [edx+00000004h], ecx
  loc_00E303A1: mov ecx, [esi]
  loc_00E303A3: mov [edx+00000008h], eax
  loc_00E303A6: mov eax, var_50
  loc_00E303A9: mov [edx+0000000Ch], eax
  loc_00E303AC: call [ecx+000004A8h]
  loc_00E303B2: lea edx, var_34
  loc_00E303B5: push eax
  loc_00E303B6: push edx
  loc_00E303B7: call edi
  loc_00E303B9: push eax
  loc_00E303BA: call [00401238h] ; __vbaLateIdSt
  loc_00E303C0: lea ecx, var_34
  loc_00E303C3: call [00401254h] ; __vbaFreeObj
  loc_00E303C9: lea ecx, var_5C
  loc_00E303CC: call [00401024h] ; __vbaFreeVar
  loc_00E303D2: mov eax, [esi]
  loc_00E303D4: push 00000000h
  loc_00E303D6: push 00000019h
  loc_00E303D8: push esi
  loc_00E303D9: call [eax+000004A8h]
  loc_00E303DF: lea ecx, var_34
  loc_00E303E2: push eax
  loc_00E303E3: push ecx
  loc_00E303E4: call edi
  loc_00E303E6: push eax
  loc_00E303E7: call ebx
  loc_00E303E9: add esp, 0000000Ch
  loc_00E303EC: lea ecx, var_34
  loc_00E303EF: call [00401254h] ; __vbaFreeObj
  loc_00E303F5: mov edx, [esi]
  loc_00E303F7: push 006E05D8h
  loc_00E303FC: push esi
  loc_00E303FD: call [edx+000004A8h]
  loc_00E30403: push eax
  loc_00E30404: lea eax, var_34
  loc_00E30407: push eax
  loc_00E30408: call edi
  loc_00E3040A: push eax
  loc_00E3040B: call [00401224h] ; __vbaCastObj
  loc_00E30411: lea ecx, var_5C
  loc_00E30414: mov var_54, eax
  loc_00E30417: push ecx
  loc_00E30418: mov var_5C, 0000000Dh
  loc_00E3041F: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E30425: mov ecx, [eax]
  loc_00E30427: sub esp, 00000010h
  loc_00E3042A: mov edx, esp
  loc_00E3042C: push 00000000h
  loc_00E3042E: push 0000002Ah
  loc_00E30430: mov [edx], ecx
  loc_00E30432: mov ecx, [eax+00000004h]
  loc_00E30435: push esi
  loc_00E30436: mov [edx+00000004h], ecx
  loc_00E30439: mov ecx, [eax+00000008h]
  loc_00E3043C: mov eax, [eax+0000000Ch]
  loc_00E3043F: mov [edx+00000008h], ecx
  loc_00E30442: mov ecx, [esi]
  loc_00E30444: mov [edx+0000000Ch], eax
  loc_00E30447: call [ecx+000004ACh]
  loc_00E3044D: lea edx, var_38
  loc_00E30450: push eax
  loc_00E30451: push edx
  loc_00E30452: call edi
  loc_00E30454: push eax
  loc_00E30455: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E3045B: lea eax, var_38
  loc_00E3045E: lea ecx, var_34
  loc_00E30461: push eax
  loc_00E30462: push ecx
  loc_00E30463: push 00000002h
  loc_00E30465: call [00401048h] ; __vbaFreeObjList
  loc_00E3046B: add esp, 00000028h
  loc_00E3046E: lea ecx, var_5C
  loc_00E30471: call [00401024h] ; __vbaFreeVar
  loc_00E30477: mov edx, [esi]
  loc_00E30479: push 00000000h
  loc_00E3047B: push FFFFFDDAh
  loc_00E30480: push esi
  loc_00E30481: call [edx+000004ACh]
  loc_00E30487: push eax
  loc_00E30488: lea eax, var_34
  loc_00E3048B: push eax
  loc_00E3048C: call edi
  loc_00E3048E: push eax
  loc_00E3048F: call ebx
  loc_00E30491: add esp, 0000000Ch
  loc_00E30494: lea ecx, var_34
  loc_00E30497: call [00401254h] ; __vbaFreeObj
  loc_00E3049D: mov var_4, 00000000h
  loc_00E304A4: fwait
  loc_00E304A5: push 00E3050Dh
  loc_00E304AA: jmp 00E30503h
  loc_00E304AC: lea ecx, var_30
  loc_00E304AF: lea edx, var_2C
  loc_00E304B2: push ecx
  loc_00E304B3: lea eax, var_28
  loc_00E304B6: push edx
  loc_00E304B7: push eax
  loc_00E304B8: push 00000003h
  loc_00E304BA: call [004011BCh] ; __vbaFreeStrList
  loc_00E304C0: lea ecx, var_4C
  loc_00E304C3: lea edx, var_48
  loc_00E304C6: push ecx
  loc_00E304C7: lea eax, var_44
  loc_00E304CA: push edx
  loc_00E304CB: lea ecx, var_40
  loc_00E304CE: push eax
  loc_00E304CF: lea edx, var_3C
  loc_00E304D2: push ecx
  loc_00E304D3: lea eax, var_38
  loc_00E304D6: push edx
  loc_00E304D7: lea ecx, var_34
  loc_00E304DA: push eax
  loc_00E304DB: push ecx
  loc_00E304DC: push 00000007h
  loc_00E304DE: call [00401048h] ; __vbaFreeObjList
  loc_00E304E4: lea edx, var_8C
  loc_00E304EA: lea eax, var_7C
  loc_00E304ED: push edx
  loc_00E304EE: lea ecx, var_6C
  loc_00E304F1: push eax
  loc_00E304F2: lea edx, var_5C
  loc_00E304F5: push ecx
  loc_00E304F6: push edx
  loc_00E304F7: push 00000004h
  loc_00E304F9: call [00401038h] ; __vbaFreeVarList
  loc_00E304FF: add esp, 00000044h
  loc_00E30502: ret
  loc_00E30503: lea ecx, var_24
  loc_00E30506: call [00401024h] ; __vbaFreeVar
  loc_00E3050C: ret
  loc_00E3050D: mov eax, Me
  loc_00E30510: push eax
  loc_00E30511: mov ecx, [eax]
  loc_00E30513: call [ecx+00000008h]
  loc_00E30516: mov eax, var_4
  loc_00E30519: mov ecx, var_14
  loc_00E3051C: pop edi
  loc_00E3051D: pop esi
  loc_00E3051E: mov fs:[00000000h], ecx
  loc_00E30525: pop ebx
  loc_00E30526: mov esp, ebp
  loc_00E30528: pop ebp
  loc_00E30529: retn 0004h
End Sub

Private Sub simpan_UnknownEvent_9 'E26E10
  loc_00E26E10: push ebp
  loc_00E26E11: mov ebp, esp
  loc_00E26E13: sub esp, 0000000Ch
  loc_00E26E16: push 00402836h ; __vbaExceptHandler
  loc_00E26E1B: mov eax, fs:[00000000h]
  loc_00E26E21: push eax
  loc_00E26E22: mov fs:[00000000h], esp
  loc_00E26E29: sub esp, 00000180h
  loc_00E26E2F: push ebx
  loc_00E26E30: push esi
  loc_00E26E31: push edi
  loc_00E26E32: mov var_C, esp
  loc_00E26E35: mov var_8, 00402540h
  loc_00E26E3C: mov esi, Me
  loc_00E26E3F: mov eax, esi
  loc_00E26E41: and eax, 00000001h
  loc_00E26E44: mov var_4, eax
  loc_00E26E47: and esi, FFFFFFFEh
  loc_00E26E4A: push esi
  loc_00E26E4B: mov Me, esi
  loc_00E26E4E: mov ecx, [esi]
  loc_00E26E50: call [ecx+00000004h]
  loc_00E26E53: mov edx, [esi]
  loc_00E26E55: xor ebx, ebx
  loc_00E26E57: push esi
  loc_00E26E58: mov var_24, ebx
  loc_00E26E5B: mov var_28, ebx
  loc_00E26E5E: mov var_2C, ebx
  loc_00E26E61: mov var_30, ebx
  loc_00E26E64: mov var_34, ebx
  loc_00E26E67: mov var_38, ebx
  loc_00E26E6A: mov var_3C, ebx
  loc_00E26E6D: mov var_40, ebx
  loc_00E26E70: mov var_50, ebx
  loc_00E26E73: mov var_60, ebx
  loc_00E26E76: mov var_70, ebx
  loc_00E26E79: mov var_80, ebx
  loc_00E26E7C: mov var_90, ebx
  loc_00E26E82: mov var_A0, ebx
  loc_00E26E88: mov var_C4, ebx
  loc_00E26E8E: call [edx+00000430h]
  loc_00E26E94: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E26E9A: push eax
  loc_00E26E9B: lea eax, var_30
  loc_00E26E9E: push eax
  loc_00E26E9F: call edi
  loc_00E26EA1: mov ecx, [eax]
  loc_00E26EA3: lea edx, var_28
  loc_00E26EA6: push edx
  loc_00E26EA7: push eax
  loc_00E26EA8: mov var_D0, eax
  loc_00E26EAE: call [ecx+00000050h]
  loc_00E26EB1: cmp eax, ebx
  loc_00E26EB3: fnclex
  loc_00E26EB5: jge 00E26ED0h
  loc_00E26EB7: mov ecx, var_D0
  loc_00E26EBD: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E26EC3: push 00000050h
  loc_00E26EC5: push 006E0410h
  loc_00E26ECA: push ecx
  loc_00E26ECB: push eax
  loc_00E26ECC: call ebx
  loc_00E26ECE: jmp 00E26ED6h
  loc_00E26ED0: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E26ED6: mov edx, var_28
  loc_00E26ED9: push 006E1744h ; "select * from lc where no_daftar ='"
  loc_00E26EDE: push edx
  loc_00E26EDF: call [0040106Ch] ; __vbaStrCat
  loc_00E26EE5: mov edx, eax
  loc_00E26EE7: lea ecx, var_2C
  loc_00E26EEA: call [00401228h] ; __vbaStrMove
  loc_00E26EF0: push eax
  loc_00E26EF1: push 006DCB84h ; "'"
  loc_00E26EF6: call [0040106Ch] ; __vbaStrCat
  loc_00E26EFC: lea edx, var_50
  loc_00E26EFF: lea ecx, var_24
  loc_00E26F02: mov var_48, eax
  loc_00E26F05: mov var_50, 00000008h
  loc_00E26F0C: call [0040101Ch] ; __vbaVarMove
  loc_00E26F12: lea eax, var_2C
  loc_00E26F15: lea ecx, var_28
  loc_00E26F18: push eax
  loc_00E26F19: push ecx
  loc_00E26F1A: push 00000002h
  loc_00E26F1C: call [004011BCh] ; __vbaFreeStrList
  loc_00E26F22: add esp, 0000000Ch
  loc_00E26F25: lea ecx, var_30
  loc_00E26F28: call [00401254h] ; __vbaFreeObj
  loc_00E26F2E: lea edx, var_24
  loc_00E26F31: push edx
  loc_00E26F32: call [00401230h] ; __vbaStrVarCopy
  loc_00E26F38: sub esp, 00000010h
  loc_00E26F3B: mov ecx, 00000008h
  loc_00E26F40: mov edx, esp
  loc_00E26F42: mov var_50, ecx
  loc_00E26F45: mov var_48, eax
  loc_00E26F48: push 0000000Eh
  loc_00E26F4A: mov [edx], ecx
  loc_00E26F4C: mov ecx, var_4C
  loc_00E26F4F: push esi
  loc_00E26F50: mov [edx+00000004h], ecx
  loc_00E26F53: mov ecx, [esi]
  loc_00E26F55: mov [edx+00000008h], eax
  loc_00E26F58: mov eax, var_44
  loc_00E26F5B: mov [edx+0000000Ch], eax
  loc_00E26F5E: call [ecx+000004A8h]
  loc_00E26F64: lea edx, var_30
  loc_00E26F67: push eax
  loc_00E26F68: push edx
  loc_00E26F69: call edi
  loc_00E26F6B: push eax
  loc_00E26F6C: call [00401238h] ; __vbaLateIdSt
  loc_00E26F72: lea ecx, var_30
  loc_00E26F75: call [00401254h] ; __vbaFreeObj
  loc_00E26F7B: lea ecx, var_50
  loc_00E26F7E: call [00401024h] ; __vbaFreeVar
  loc_00E26F84: mov eax, [esi]
  loc_00E26F86: push 00000000h
  loc_00E26F88: push 00000019h
  loc_00E26F8A: push esi
  loc_00E26F8B: call [eax+000004A8h]
  loc_00E26F91: lea ecx, var_30
  loc_00E26F94: push eax
  loc_00E26F95: push ecx
  loc_00E26F96: call edi
  loc_00E26F98: push eax
  loc_00E26F99: call [00401030h] ; __vbaLateIdCall
  loc_00E26F9F: add esp, 0000000Ch
  loc_00E26FA2: lea ecx, var_30
  loc_00E26FA5: call [00401254h] ; __vbaFreeObj
  loc_00E26FAB: mov edx, [esi]
  loc_00E26FAD: push esi
  loc_00E26FAE: call [edx+00000378h]
  loc_00E26FB4: push eax
  loc_00E26FB5: lea eax, var_30
  loc_00E26FB8: push eax
  loc_00E26FB9: call edi
  loc_00E26FBB: mov ecx, [eax]
  loc_00E26FBD: lea edx, var_28
  loc_00E26FC0: push edx
  loc_00E26FC1: push eax
  loc_00E26FC2: mov var_D0, eax
  loc_00E26FC8: call [ecx+000000A0h]
  loc_00E26FCE: fnclex
  loc_00E26FD0: test eax, eax
  loc_00E26FD2: jge 00E26FE8h
  loc_00E26FD4: mov ecx, var_D0
  loc_00E26FDA: push 000000A0h
  loc_00E26FDF: push 006DCB70h
  loc_00E26FE4: push ecx
  loc_00E26FE5: push eax
  loc_00E26FE6: call ebx
  loc_00E26FE8: mov edx, var_28
  loc_00E26FEB: push edx
  loc_00E26FEC: push 006E0934h
  loc_00E26FF1: call [00401104h] ; __vbaStrCmp
  loc_00E26FF7: neg eax
  loc_00E26FF9: sbb eax, eax
  loc_00E26FFB: lea ecx, var_28
  loc_00E26FFE: inc eax
  loc_00E26FFF: neg eax
  loc_00E27001: mov var_D8, ax
  loc_00E27008: call [00401258h] ; __vbaFreeStr
  loc_00E2700E: lea ecx, var_30
  loc_00E27011: call [00401254h] ; __vbaFreeObj
  loc_00E27017: cmp var_D8, 0000h
  loc_00E2701F: jz 00E270F1h
  loc_00E27025: mov ecx, 80020004h
  loc_00E2702A: mov eax, 0000000Ah
  loc_00E2702F: mov var_78, ecx
  loc_00E27032: mov var_68, ecx
  loc_00E27035: lea edx, var_A0
  loc_00E2703B: lea ecx, var_60
  loc_00E2703E: mov var_80, eax
  loc_00E27041: mov var_70, eax
  loc_00E27044: mov var_98, 006E0B08h ; "Need To Do"
  loc_00E2704E: mov var_A0, 00000008h
  loc_00E27058: call [004011F0h] ; __vbaVarDup
  loc_00E2705E: lea edx, var_90
  loc_00E27064: lea ecx, var_50
  loc_00E27067: mov var_88, 006E1790h ; "Silahkan Input Jumlah Uang Bayar Terlebih dahulu !"
  loc_00E27071: mov var_90, 00000008h
  loc_00E2707B: call [004011F0h] ; __vbaVarDup
  loc_00E27081: lea eax, var_80
  loc_00E27084: lea ecx, var_70
  loc_00E27087: push eax
  loc_00E27088: lea edx, var_60
  loc_00E2708B: push ecx
  loc_00E2708C: push edx
  loc_00E2708D: lea eax, var_50
  loc_00E27090: push 00000040h
  loc_00E27092: push eax
  loc_00E27093: call [004010A8h] ; rtcMsgBox
  loc_00E27099: lea ecx, var_80
  loc_00E2709C: lea edx, var_70
  loc_00E2709F: push ecx
  loc_00E270A0: lea eax, var_60
  loc_00E270A3: push edx
  loc_00E270A4: lea ecx, var_50
  loc_00E270A7: push eax
  loc_00E270A8: push ecx
  loc_00E270A9: push 00000004h
  loc_00E270AB: call [00401038h] ; __vbaFreeVarList
  loc_00E270B1: mov edx, [esi]
  loc_00E270B3: add esp, 00000014h
  loc_00E270B6: push esi
  loc_00E270B7: call [edx+00000378h]
  loc_00E270BD: push eax
  loc_00E270BE: lea eax, var_30
  loc_00E270C1: push eax
  loc_00E270C2: call edi
  loc_00E270C4: mov esi, eax
  loc_00E270C6: push esi
  loc_00E270C7: mov ecx, [esi]
  loc_00E270C9: call [ecx+00000204h]
  loc_00E270CF: test eax, eax
  loc_00E270D1: fnclex
  loc_00E270D3: jge 00E270E3h
  loc_00E270D5: push 00000204h
  loc_00E270DA: push 006DCB70h
  loc_00E270DF: push esi
  loc_00E270E0: push eax
  loc_00E270E1: call ebx
  loc_00E270E3: lea ecx, var_30
  loc_00E270E6: call [00401254h] ; __vbaFreeObj
  loc_00E270EC: jmp 00E2C7E5h
  loc_00E270F1: mov edx, [esi]
  loc_00E270F3: push esi
  loc_00E270F4: call [edx+0000037Ch]
  loc_00E270FA: push eax
  loc_00E270FB: lea eax, var_30
  loc_00E270FE: push eax
  loc_00E270FF: call edi
  loc_00E27101: mov ecx, [eax]
  loc_00E27103: lea edx, var_C4
  loc_00E27109: push edx
  loc_00E2710A: push eax
  loc_00E2710B: mov var_D0, eax
  loc_00E27111: call [ecx+00000098h]
  loc_00E27117: test eax, eax
  loc_00E27119: fnclex
  loc_00E2711B: jge 00E27131h
  loc_00E2711D: mov ecx, var_D0
  loc_00E27123: push 00000098h
  loc_00E27128: push 006DCAD0h
  loc_00E2712D: push ecx
  loc_00E2712E: push eax
  loc_00E2712F: call ebx
  loc_00E27131: xor edx, edx
  loc_00E27133: cmp var_C4, FFFFFFh
  loc_00E2713B: lea ecx, var_30
  loc_00E2713E: setz dl
  loc_00E27141: neg edx
  loc_00E27143: mov var_D8, dx
  loc_00E2714A: call [00401254h] ; __vbaFreeObj
  loc_00E27150: cmp var_D8, 0000h
  loc_00E27158: jz 00E29CADh
  loc_00E2715E: mov eax, [esi]
  loc_00E27160: push 006DCBD8h
  loc_00E27165: push 00000000h
  loc_00E27167: push 00000018h
  loc_00E27169: push esi
  loc_00E2716A: call [eax+000004A8h]
  loc_00E27170: lea ecx, var_30
  loc_00E27173: push eax
  loc_00E27174: push ecx
  loc_00E27175: call edi
  loc_00E27177: lea edx, var_50
  loc_00E2717A: push eax
  loc_00E2717B: push edx
  loc_00E2717C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27182: add esp, 00000010h
  loc_00E27185: push eax
  loc_00E27186: call [00401120h] ; __vbaCastObjVar
  loc_00E2718C: push eax
  loc_00E2718D: lea eax, var_34
  loc_00E27190: push eax
  loc_00E27191: call edi
  loc_00E27193: mov ecx, [eax]
  loc_00E27195: lea edx, var_C4
  loc_00E2719B: push edx
  loc_00E2719C: push eax
  loc_00E2719D: mov var_D0, eax
  loc_00E271A3: call [ecx+00000050h]
  loc_00E271A6: test eax, eax
  loc_00E271A8: fnclex
  loc_00E271AA: jge 00E271BDh
  loc_00E271AC: mov ecx, var_D0
  loc_00E271B2: push 00000050h
  loc_00E271B4: push 006DCBE8h
  loc_00E271B9: push ecx
  loc_00E271BA: push eax
  loc_00E271BB: call ebx
  loc_00E271BD: mov dx, var_C4
  loc_00E271C4: lea eax, var_34
  loc_00E271C7: lea ecx, var_30
  loc_00E271CA: push eax
  loc_00E271CB: push ecx
  loc_00E271CC: push 00000002h
  loc_00E271CE: mov var_D8, dx
  loc_00E271D5: call [00401048h] ; __vbaFreeObjList
  loc_00E271DB: add esp, 0000000Ch
  loc_00E271DE: lea ecx, var_50
  loc_00E271E1: call [00401024h] ; __vbaFreeVar
  loc_00E271E7: cmp var_D8, 0000h
  loc_00E271EF: jz 00E2C7E5h
  loc_00E271F5: mov edx, [esi]
  loc_00E271F7: push 006DCBD8h
  loc_00E271FC: push 00000000h
  loc_00E271FE: push 00000018h
  loc_00E27200: push esi
  loc_00E27201: call [edx+000004A8h]
  loc_00E27207: push eax
  loc_00E27208: lea eax, var_30
  loc_00E2720B: push eax
  loc_00E2720C: call edi
  loc_00E2720E: lea ecx, var_50
  loc_00E27211: push eax
  loc_00E27212: push ecx
  loc_00E27213: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27219: add esp, 00000010h
  loc_00E2721C: push eax
  loc_00E2721D: call [00401120h] ; __vbaCastObjVar
  loc_00E27223: lea edx, var_34
  loc_00E27226: push eax
  loc_00E27227: push edx
  loc_00E27228: call edi
  loc_00E2722A: sub esp, 00000010h
  loc_00E2722D: mov ecx, 0000000Ah
  loc_00E27232: mov edx, esp
  loc_00E27234: mov var_A0, ecx
  loc_00E2723A: mov var_90, ecx
  loc_00E27240: mov var_98, 80020004h
  loc_00E2724A: mov [edx], ecx
  loc_00E2724C: mov ecx, var_9C
  loc_00E27252: sub esp, 00000010h
  loc_00E27255: mov var_88, 80020004h
  loc_00E2725F: mov [edx+00000004h], ecx
  loc_00E27262: mov ecx, var_98
  loc_00E27268: mov var_D0, eax
  loc_00E2726E: mov eax, [eax]
  loc_00E27270: mov [edx+00000008h], ecx
  loc_00E27273: mov ecx, var_94
  loc_00E27279: mov [edx+0000000Ch], ecx
  loc_00E2727C: mov ecx, var_90
  loc_00E27282: mov edx, esp
  loc_00E27284: mov [edx], ecx
  loc_00E27286: mov ecx, var_8C
  loc_00E2728C: mov [edx+00000004h], ecx
  loc_00E2728F: mov ecx, var_88
  loc_00E27295: mov [edx+00000008h], ecx
  loc_00E27298: mov ecx, var_84
  loc_00E2729E: mov [edx+0000000Ch], ecx
  loc_00E272A1: mov edx, var_D0
  loc_00E272A7: push edx
  loc_00E272A8: call [eax+00000078h]
  loc_00E272AB: test eax, eax
  loc_00E272AD: fnclex
  loc_00E272AF: jge 00E272C2h
  loc_00E272B1: mov ecx, var_D0
  loc_00E272B7: push 00000078h
  loc_00E272B9: push 006DCBE8h
  loc_00E272BE: push ecx
  loc_00E272BF: push eax
  loc_00E272C0: call ebx
  loc_00E272C2: lea edx, var_34
  loc_00E272C5: lea eax, var_30
  loc_00E272C8: push edx
  loc_00E272C9: push eax
  loc_00E272CA: push 00000002h
  loc_00E272CC: call [00401048h] ; __vbaFreeObjList
  loc_00E272D2: add esp, 0000000Ch
  loc_00E272D5: lea ecx, var_50
  loc_00E272D8: call [00401024h] ; __vbaFreeVar
  loc_00E272DE: mov ecx, [esi]
  loc_00E272E0: push 006DCBD8h
  loc_00E272E5: push 00000000h
  loc_00E272E7: push 00000018h
  loc_00E272E9: push esi
  loc_00E272EA: call [ecx+000004A8h]
  loc_00E272F0: lea edx, var_34
  loc_00E272F3: push eax
  loc_00E272F4: push edx
  loc_00E272F5: call edi
  loc_00E272F7: push eax
  loc_00E272F8: lea eax, var_50
  loc_00E272FB: push eax
  loc_00E272FC: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27302: add esp, 00000010h
  loc_00E27305: push eax
  loc_00E27306: call [00401120h] ; __vbaCastObjVar
  loc_00E2730C: lea ecx, var_38
  loc_00E2730F: push eax
  loc_00E27310: push ecx
  loc_00E27311: call edi
  loc_00E27313: mov edx, [eax]
  loc_00E27315: lea ecx, var_3C
  loc_00E27318: push ecx
  loc_00E27319: push eax
  loc_00E2731A: mov var_D8, eax
  loc_00E27320: call [edx+00000054h]
  loc_00E27323: test eax, eax
  loc_00E27325: fnclex
  loc_00E27327: jge 00E2733Ah
  loc_00E27329: mov edx, var_D8
  loc_00E2732F: push 00000054h
  loc_00E27331: push 006DCBE8h
  loc_00E27336: push edx
  loc_00E27337: push eax
  loc_00E27338: call ebx
  loc_00E2733A: lea edx, var_40
  loc_00E2733D: mov eax, 00000002h
  loc_00E27342: push edx
  loc_00E27343: mov ecx, var_3C
  loc_00E27346: sub esp, 00000010h
  loc_00E27349: mov var_90, eax
  loc_00E2734F: mov edx, esp
  loc_00E27351: mov var_88, 00000000h
  loc_00E2735B: mov var_E0, ecx
  loc_00E27361: mov ecx, [ecx]
  loc_00E27363: mov [edx], eax
  loc_00E27365: mov eax, var_8C
  loc_00E2736B: mov [edx+00000004h], eax
  loc_00E2736E: mov eax, var_88
  loc_00E27374: mov [edx+00000008h], eax
  loc_00E27377: mov eax, var_84
  loc_00E2737D: mov [edx+0000000Ch], eax
  loc_00E27380: mov edx, var_3C
  loc_00E27383: push edx
  loc_00E27384: call [ecx+00000028h]
  loc_00E27387: test eax, eax
  loc_00E27389: fnclex
  loc_00E2738B: jge 00E2739Eh
  loc_00E2738D: mov ecx, var_E0
  loc_00E27393: push 00000028h
  loc_00E27395: push 006E09E8h
  loc_00E2739A: push ecx
  loc_00E2739B: push eax
  loc_00E2739C: call ebx
  loc_00E2739E: mov edx, var_40
  loc_00E273A1: mov eax, [esi]
  loc_00E273A3: push esi
  loc_00E273A4: mov var_E8, edx
  loc_00E273AA: call [eax+00000430h]
  loc_00E273B0: lea ecx, var_30
  loc_00E273B3: push eax
  loc_00E273B4: push ecx
  loc_00E273B5: call edi
  loc_00E273B7: mov edx, [eax]
  loc_00E273B9: lea ecx, var_28
  loc_00E273BC: push ecx
  loc_00E273BD: push eax
  loc_00E273BE: mov var_D0, eax
  loc_00E273C4: call [edx+00000050h]
  loc_00E273C7: test eax, eax
  loc_00E273C9: fnclex
  loc_00E273CB: jge 00E273DEh
  loc_00E273CD: mov edx, var_D0
  loc_00E273D3: push 00000050h
  loc_00E273D5: push 006E0410h
  loc_00E273DA: push edx
  loc_00E273DB: push eax
  loc_00E273DC: call ebx
  loc_00E273DE: mov eax, var_28
  loc_00E273E1: mov ecx, var_E8
  loc_00E273E7: mov var_58, eax
  loc_00E273EA: mov eax, 00000008h
  loc_00E273EF: mov var_28, 00000000h
  loc_00E273F6: mov var_60, eax
  loc_00E273F9: mov edx, [ecx]
  loc_00E273FB: sub esp, 00000010h
  loc_00E273FE: mov ecx, esp
  loc_00E27400: mov [ecx], eax
  loc_00E27402: mov eax, var_5C
  loc_00E27405: mov [ecx+00000004h], eax
  loc_00E27408: mov eax, var_58
  loc_00E2740B: mov [ecx+00000008h], eax
  loc_00E2740E: mov eax, var_54
  loc_00E27411: mov [ecx+0000000Ch], eax
  loc_00E27414: mov ecx, var_E8
  loc_00E2741A: push ecx
  loc_00E2741B: call [edx+00000038h]
  loc_00E2741E: test eax, eax
  loc_00E27420: fnclex
  loc_00E27422: jge 00E27435h
  loc_00E27424: mov edx, var_E8
  loc_00E2742A: push 00000038h
  loc_00E2742C: push 006E09F8h
  loc_00E27431: push edx
  loc_00E27432: push eax
  loc_00E27433: call ebx
  loc_00E27435: lea eax, var_40
  loc_00E27438: lea ecx, var_3C
  loc_00E2743B: push eax
  loc_00E2743C: lea edx, var_38
  loc_00E2743F: push ecx
  loc_00E27440: lea eax, var_34
  loc_00E27443: push edx
  loc_00E27444: lea ecx, var_30
  loc_00E27447: push eax
  loc_00E27448: push ecx
  loc_00E27449: push 00000005h
  loc_00E2744B: call [00401048h] ; __vbaFreeObjList
  loc_00E27451: lea edx, var_60
  loc_00E27454: lea eax, var_50
  loc_00E27457: push edx
  loc_00E27458: push eax
  loc_00E27459: push 00000002h
  loc_00E2745B: call [00401038h] ; __vbaFreeVarList
  loc_00E27461: mov ecx, [esi]
  loc_00E27463: add esp, 00000024h
  loc_00E27466: push 006DCBD8h
  loc_00E2746B: push 00000000h
  loc_00E2746D: push 00000018h
  loc_00E2746F: push esi
  loc_00E27470: call [ecx+000004A8h]
  loc_00E27476: lea edx, var_34
  loc_00E27479: push eax
  loc_00E2747A: push edx
  loc_00E2747B: call edi
  loc_00E2747D: push eax
  loc_00E2747E: lea eax, var_50
  loc_00E27481: push eax
  loc_00E27482: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27488: add esp, 00000010h
  loc_00E2748B: push eax
  loc_00E2748C: call [00401120h] ; __vbaCastObjVar
  loc_00E27492: lea ecx, var_38
  loc_00E27495: push eax
  loc_00E27496: push ecx
  loc_00E27497: call edi
  loc_00E27499: mov edx, [eax]
  loc_00E2749B: lea ecx, var_3C
  loc_00E2749E: push ecx
  loc_00E2749F: push eax
  loc_00E274A0: mov var_D8, eax
  loc_00E274A6: call [edx+00000054h]
  loc_00E274A9: test eax, eax
  loc_00E274AB: fnclex
  loc_00E274AD: jge 00E274C0h
  loc_00E274AF: mov edx, var_D8
  loc_00E274B5: push 00000054h
  loc_00E274B7: push 006DCBE8h
  loc_00E274BC: push edx
  loc_00E274BD: push eax
  loc_00E274BE: call ebx
  loc_00E274C0: lea edx, var_40
  loc_00E274C3: mov eax, 00000002h
  loc_00E274C8: push edx
  loc_00E274C9: mov ecx, var_3C
  loc_00E274CC: sub esp, 00000010h
  loc_00E274CF: mov var_90, eax
  loc_00E274D5: mov edx, esp
  loc_00E274D7: mov var_88, 00000001h
  loc_00E274E1: mov var_E0, ecx
  loc_00E274E7: mov ecx, [ecx]
  loc_00E274E9: mov [edx], eax
  loc_00E274EB: mov eax, var_8C
  loc_00E274F1: mov [edx+00000004h], eax
  loc_00E274F4: mov eax, var_88
  loc_00E274FA: mov [edx+00000008h], eax
  loc_00E274FD: mov eax, var_84
  loc_00E27503: mov [edx+0000000Ch], eax
  loc_00E27506: mov edx, var_3C
  loc_00E27509: push edx
  loc_00E2750A: call [ecx+00000028h]
  loc_00E2750D: test eax, eax
  loc_00E2750F: fnclex
  loc_00E27511: jge 00E27524h
  loc_00E27513: mov ecx, var_E0
  loc_00E27519: push 00000028h
  loc_00E2751B: push 006E09E8h
  loc_00E27520: push ecx
  loc_00E27521: push eax
  loc_00E27522: call ebx
  loc_00E27524: mov edx, var_40
  loc_00E27527: mov eax, [esi]
  loc_00E27529: push esi
  loc_00E2752A: mov var_E8, edx
  loc_00E27530: call [eax+00000434h]
  loc_00E27536: lea ecx, var_30
  loc_00E27539: push eax
  loc_00E2753A: push ecx
  loc_00E2753B: call edi
  loc_00E2753D: mov edx, [eax]
  loc_00E2753F: lea ecx, var_28
  loc_00E27542: push ecx
  loc_00E27543: push eax
  loc_00E27544: mov var_D0, eax
  loc_00E2754A: call [edx+00000050h]
  loc_00E2754D: test eax, eax
  loc_00E2754F: fnclex
  loc_00E27551: jge 00E27564h
  loc_00E27553: mov edx, var_D0
  loc_00E27559: push 00000050h
  loc_00E2755B: push 006E0410h
  loc_00E27560: push edx
  loc_00E27561: push eax
  loc_00E27562: call ebx
  loc_00E27564: mov eax, var_28
  loc_00E27567: mov ecx, var_E8
  loc_00E2756D: mov var_58, eax
  loc_00E27570: mov eax, 00000008h
  loc_00E27575: mov var_28, 00000000h
  loc_00E2757C: mov var_60, eax
  loc_00E2757F: mov edx, [ecx]
  loc_00E27581: sub esp, 00000010h
  loc_00E27584: mov ecx, esp
  loc_00E27586: mov [ecx], eax
  loc_00E27588: mov eax, var_5C
  loc_00E2758B: mov [ecx+00000004h], eax
  loc_00E2758E: mov eax, var_58
  loc_00E27591: mov [ecx+00000008h], eax
  loc_00E27594: mov eax, var_54
  loc_00E27597: mov [ecx+0000000Ch], eax
  loc_00E2759A: mov ecx, var_E8
  loc_00E275A0: push ecx
  loc_00E275A1: call [edx+00000038h]
  loc_00E275A4: test eax, eax
  loc_00E275A6: fnclex
  loc_00E275A8: jge 00E275BBh
  loc_00E275AA: mov edx, var_E8
  loc_00E275B0: push 00000038h
  loc_00E275B2: push 006E09F8h
  loc_00E275B7: push edx
  loc_00E275B8: push eax
  loc_00E275B9: call ebx
  loc_00E275BB: lea eax, var_40
  loc_00E275BE: lea ecx, var_3C
  loc_00E275C1: push eax
  loc_00E275C2: lea edx, var_38
  loc_00E275C5: push ecx
  loc_00E275C6: lea eax, var_34
  loc_00E275C9: push edx
  loc_00E275CA: lea ecx, var_30
  loc_00E275CD: push eax
  loc_00E275CE: push ecx
  loc_00E275CF: push 00000005h
  loc_00E275D1: call [00401048h] ; __vbaFreeObjList
  loc_00E275D7: lea edx, var_60
  loc_00E275DA: lea eax, var_50
  loc_00E275DD: push edx
  loc_00E275DE: push eax
  loc_00E275DF: push 00000002h
  loc_00E275E1: call [00401038h] ; __vbaFreeVarList
  loc_00E275E7: mov ecx, [esi]
  loc_00E275E9: add esp, 00000024h
  loc_00E275EC: push 006DCBD8h
  loc_00E275F1: push 00000000h
  loc_00E275F3: push 00000018h
  loc_00E275F5: push esi
  loc_00E275F6: call [ecx+000004A8h]
  loc_00E275FC: lea edx, var_34
  loc_00E275FF: push eax
  loc_00E27600: push edx
  loc_00E27601: call edi
  loc_00E27603: push eax
  loc_00E27604: lea eax, var_50
  loc_00E27607: push eax
  loc_00E27608: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2760E: add esp, 00000010h
  loc_00E27611: push eax
  loc_00E27612: call [00401120h] ; __vbaCastObjVar
  loc_00E27618: lea ecx, var_38
  loc_00E2761B: push eax
  loc_00E2761C: push ecx
  loc_00E2761D: call edi
  loc_00E2761F: mov edx, [eax]
  loc_00E27621: lea ecx, var_3C
  loc_00E27624: push ecx
  loc_00E27625: push eax
  loc_00E27626: mov var_D8, eax
  loc_00E2762C: call [edx+00000054h]
  loc_00E2762F: test eax, eax
  loc_00E27631: fnclex
  loc_00E27633: jge 00E27646h
  loc_00E27635: mov edx, var_D8
  loc_00E2763B: push 00000054h
  loc_00E2763D: push 006DCBE8h
  loc_00E27642: push edx
  loc_00E27643: push eax
  loc_00E27644: call ebx
  loc_00E27646: lea edx, var_40
  loc_00E27649: mov eax, 00000002h
  loc_00E2764E: push edx
  loc_00E2764F: mov ecx, var_3C
  loc_00E27652: sub esp, 00000010h
  loc_00E27655: mov var_88, eax
  loc_00E2765B: mov edx, esp
  loc_00E2765D: mov var_90, eax
  loc_00E27663: mov var_E0, ecx
  loc_00E27669: mov ecx, [ecx]
  loc_00E2766B: mov [edx], eax
  loc_00E2766D: mov eax, var_8C
  loc_00E27673: mov [edx+00000004h], eax
  loc_00E27676: mov eax, var_88
  loc_00E2767C: mov [edx+00000008h], eax
  loc_00E2767F: mov eax, var_84
  loc_00E27685: mov [edx+0000000Ch], eax
  loc_00E27688: mov edx, var_3C
  loc_00E2768B: push edx
  loc_00E2768C: call [ecx+00000028h]
  loc_00E2768F: test eax, eax
  loc_00E27691: fnclex
  loc_00E27693: jge 00E276A6h
  loc_00E27695: mov ecx, var_E0
  loc_00E2769B: push 00000028h
  loc_00E2769D: push 006E09E8h
  loc_00E276A2: push ecx
  loc_00E276A3: push eax
  loc_00E276A4: call ebx
  loc_00E276A6: mov edx, var_40
  loc_00E276A9: mov eax, [esi]
  loc_00E276AB: push esi
  loc_00E276AC: mov var_E8, edx
  loc_00E276B2: call [eax+00000438h]
  loc_00E276B8: lea ecx, var_30
  loc_00E276BB: push eax
  loc_00E276BC: push ecx
  loc_00E276BD: call edi
  loc_00E276BF: mov edx, [eax]
  loc_00E276C1: lea ecx, var_28
  loc_00E276C4: push ecx
  loc_00E276C5: push eax
  loc_00E276C6: mov var_D0, eax
  loc_00E276CC: call [edx+00000050h]
  loc_00E276CF: test eax, eax
  loc_00E276D1: fnclex
  loc_00E276D3: jge 00E276E6h
  loc_00E276D5: mov edx, var_D0
  loc_00E276DB: push 00000050h
  loc_00E276DD: push 006E0410h
  loc_00E276E2: push edx
  loc_00E276E3: push eax
  loc_00E276E4: call ebx
  loc_00E276E6: mov eax, var_28
  loc_00E276E9: mov ecx, var_E8
  loc_00E276EF: mov var_58, eax
  loc_00E276F2: mov eax, 00000008h
  loc_00E276F7: mov var_28, 00000000h
  loc_00E276FE: mov var_60, eax
  loc_00E27701: mov edx, [ecx]
  loc_00E27703: sub esp, 00000010h
  loc_00E27706: mov ecx, esp
  loc_00E27708: mov [ecx], eax
  loc_00E2770A: mov eax, var_5C
  loc_00E2770D: mov [ecx+00000004h], eax
  loc_00E27710: mov eax, var_58
  loc_00E27713: mov [ecx+00000008h], eax
  loc_00E27716: mov eax, var_54
  loc_00E27719: mov [ecx+0000000Ch], eax
  loc_00E2771C: mov ecx, var_E8
  loc_00E27722: push ecx
  loc_00E27723: call [edx+00000038h]
  loc_00E27726: test eax, eax
  loc_00E27728: fnclex
  loc_00E2772A: jge 00E2773Dh
  loc_00E2772C: mov edx, var_E8
  loc_00E27732: push 00000038h
  loc_00E27734: push 006E09F8h
  loc_00E27739: push edx
  loc_00E2773A: push eax
  loc_00E2773B: call ebx
  loc_00E2773D: lea eax, var_40
  loc_00E27740: lea ecx, var_3C
  loc_00E27743: push eax
  loc_00E27744: lea edx, var_38
  loc_00E27747: push ecx
  loc_00E27748: lea eax, var_34
  loc_00E2774B: push edx
  loc_00E2774C: lea ecx, var_30
  loc_00E2774F: push eax
  loc_00E27750: push ecx
  loc_00E27751: push 00000005h
  loc_00E27753: call [00401048h] ; __vbaFreeObjList
  loc_00E27759: lea edx, var_60
  loc_00E2775C: lea eax, var_50
  loc_00E2775F: push edx
  loc_00E27760: push eax
  loc_00E27761: push 00000002h
  loc_00E27763: call [00401038h] ; __vbaFreeVarList
  loc_00E27769: mov ecx, [esi]
  loc_00E2776B: add esp, 00000024h
  loc_00E2776E: push 006DCBD8h
  loc_00E27773: push 00000000h
  loc_00E27775: push 00000018h
  loc_00E27777: push esi
  loc_00E27778: call [ecx+000004A8h]
  loc_00E2777E: lea edx, var_34
  loc_00E27781: push eax
  loc_00E27782: push edx
  loc_00E27783: call edi
  loc_00E27785: push eax
  loc_00E27786: lea eax, var_50
  loc_00E27789: push eax
  loc_00E2778A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27790: add esp, 00000010h
  loc_00E27793: push eax
  loc_00E27794: call [00401120h] ; __vbaCastObjVar
  loc_00E2779A: lea ecx, var_38
  loc_00E2779D: push eax
  loc_00E2779E: push ecx
  loc_00E2779F: call edi
  loc_00E277A1: mov edx, [eax]
  loc_00E277A3: lea ecx, var_3C
  loc_00E277A6: push ecx
  loc_00E277A7: push eax
  loc_00E277A8: mov var_D8, eax
  loc_00E277AE: call [edx+00000054h]
  loc_00E277B1: test eax, eax
  loc_00E277B3: fnclex
  loc_00E277B5: jge 00E277C8h
  loc_00E277B7: mov edx, var_D8
  loc_00E277BD: push 00000054h
  loc_00E277BF: push 006DCBE8h
  loc_00E277C4: push edx
  loc_00E277C5: push eax
  loc_00E277C6: call ebx
  loc_00E277C8: lea edx, var_40
  loc_00E277CB: mov eax, 00000002h
  loc_00E277D0: push edx
  loc_00E277D1: mov ecx, var_3C
  loc_00E277D4: sub esp, 00000010h
  loc_00E277D7: mov var_90, eax
  loc_00E277DD: mov edx, esp
  loc_00E277DF: mov var_88, 00000003h
  loc_00E277E9: mov var_E0, ecx
  loc_00E277EF: mov ecx, [ecx]
  loc_00E277F1: mov [edx], eax
  loc_00E277F3: mov eax, var_8C
  loc_00E277F9: mov [edx+00000004h], eax
  loc_00E277FC: mov eax, var_88
  loc_00E27802: mov [edx+00000008h], eax
  loc_00E27805: mov eax, var_84
  loc_00E2780B: mov [edx+0000000Ch], eax
  loc_00E2780E: mov edx, var_3C
  loc_00E27811: push edx
  loc_00E27812: call [ecx+00000028h]
  loc_00E27815: test eax, eax
  loc_00E27817: fnclex
  loc_00E27819: jge 00E2782Ch
  loc_00E2781B: mov ecx, var_E0
  loc_00E27821: push 00000028h
  loc_00E27823: push 006E09E8h
  loc_00E27828: push ecx
  loc_00E27829: push eax
  loc_00E2782A: call ebx
  loc_00E2782C: mov edx, var_40
  loc_00E2782F: mov eax, [esi]
  loc_00E27831: push esi
  loc_00E27832: mov var_E8, edx
  loc_00E27838: call [eax+0000044Ch]
  loc_00E2783E: lea ecx, var_30
  loc_00E27841: push eax
  loc_00E27842: push ecx
  loc_00E27843: call edi
  loc_00E27845: mov edx, [eax]
  loc_00E27847: lea ecx, var_28
  loc_00E2784A: push ecx
  loc_00E2784B: push eax
  loc_00E2784C: mov var_D0, eax
  loc_00E27852: call [edx+00000050h]
  loc_00E27855: test eax, eax
  loc_00E27857: fnclex
  loc_00E27859: jge 00E2786Ch
  loc_00E2785B: mov edx, var_D0
  loc_00E27861: push 00000050h
  loc_00E27863: push 006E0410h
  loc_00E27868: push edx
  loc_00E27869: push eax
  loc_00E2786A: call ebx
  loc_00E2786C: mov eax, var_28
  loc_00E2786F: mov ecx, var_E8
  loc_00E27875: mov var_58, eax
  loc_00E27878: mov eax, 00000008h
  loc_00E2787D: mov var_28, 00000000h
  loc_00E27884: mov var_60, eax
  loc_00E27887: mov edx, [ecx]
  loc_00E27889: sub esp, 00000010h
  loc_00E2788C: mov ecx, esp
  loc_00E2788E: mov [ecx], eax
  loc_00E27890: mov eax, var_5C
  loc_00E27893: mov [ecx+00000004h], eax
  loc_00E27896: mov eax, var_58
  loc_00E27899: mov [ecx+00000008h], eax
  loc_00E2789C: mov eax, var_54
  loc_00E2789F: mov [ecx+0000000Ch], eax
  loc_00E278A2: mov ecx, var_E8
  loc_00E278A8: push ecx
  loc_00E278A9: call [edx+00000038h]
  loc_00E278AC: test eax, eax
  loc_00E278AE: fnclex
  loc_00E278B0: jge 00E278C3h
  loc_00E278B2: mov edx, var_E8
  loc_00E278B8: push 00000038h
  loc_00E278BA: push 006E09F8h
  loc_00E278BF: push edx
  loc_00E278C0: push eax
  loc_00E278C1: call ebx
  loc_00E278C3: lea eax, var_40
  loc_00E278C6: lea ecx, var_3C
  loc_00E278C9: push eax
  loc_00E278CA: lea edx, var_38
  loc_00E278CD: push ecx
  loc_00E278CE: lea eax, var_34
  loc_00E278D1: push edx
  loc_00E278D2: lea ecx, var_30
  loc_00E278D5: push eax
  loc_00E278D6: push ecx
  loc_00E278D7: push 00000005h
  loc_00E278D9: call [00401048h] ; __vbaFreeObjList
  loc_00E278DF: lea edx, var_60
  loc_00E278E2: lea eax, var_50
  loc_00E278E5: push edx
  loc_00E278E6: push eax
  loc_00E278E7: push 00000002h
  loc_00E278E9: call [00401038h] ; __vbaFreeVarList
  loc_00E278EF: mov ecx, [esi]
  loc_00E278F1: add esp, 00000024h
  loc_00E278F4: push 006DCBD8h
  loc_00E278F9: push 00000000h
  loc_00E278FB: push 00000018h
  loc_00E278FD: push esi
  loc_00E278FE: call [ecx+000004A8h]
  loc_00E27904: lea edx, var_34
  loc_00E27907: push eax
  loc_00E27908: push edx
  loc_00E27909: call edi
  loc_00E2790B: push eax
  loc_00E2790C: lea eax, var_50
  loc_00E2790F: push eax
  loc_00E27910: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27916: add esp, 00000010h
  loc_00E27919: push eax
  loc_00E2791A: call [00401120h] ; __vbaCastObjVar
  loc_00E27920: lea ecx, var_38
  loc_00E27923: push eax
  loc_00E27924: push ecx
  loc_00E27925: call edi
  loc_00E27927: mov edx, [eax]
  loc_00E27929: lea ecx, var_3C
  loc_00E2792C: push ecx
  loc_00E2792D: push eax
  loc_00E2792E: mov var_D8, eax
  loc_00E27934: call [edx+00000054h]
  loc_00E27937: test eax, eax
  loc_00E27939: fnclex
  loc_00E2793B: jge 00E2794Eh
  loc_00E2793D: mov edx, var_D8
  loc_00E27943: push 00000054h
  loc_00E27945: push 006DCBE8h
  loc_00E2794A: push edx
  loc_00E2794B: push eax
  loc_00E2794C: call ebx
  loc_00E2794E: lea edx, var_40
  loc_00E27951: mov eax, 00000002h
  loc_00E27956: push edx
  loc_00E27957: mov ecx, var_3C
  loc_00E2795A: sub esp, 00000010h
  loc_00E2795D: mov var_90, eax
  loc_00E27963: mov edx, esp
  loc_00E27965: mov var_88, 0000000Eh
  loc_00E2796F: mov var_E0, ecx
  loc_00E27975: mov ecx, [ecx]
  loc_00E27977: mov [edx], eax
  loc_00E27979: mov eax, var_8C
  loc_00E2797F: mov [edx+00000004h], eax
  loc_00E27982: mov eax, var_88
  loc_00E27988: mov [edx+00000008h], eax
  loc_00E2798B: mov eax, var_84
  loc_00E27991: mov [edx+0000000Ch], eax
  loc_00E27994: mov edx, var_3C
  loc_00E27997: push edx
  loc_00E27998: call [ecx+00000028h]
  loc_00E2799B: test eax, eax
  loc_00E2799D: fnclex
  loc_00E2799F: jge 00E279B2h
  loc_00E279A1: mov ecx, var_E0
  loc_00E279A7: push 00000028h
  loc_00E279A9: push 006E09E8h
  loc_00E279AE: push ecx
  loc_00E279AF: push eax
  loc_00E279B0: call ebx
  loc_00E279B2: mov edx, var_40
  loc_00E279B5: mov eax, [esi]
  loc_00E279B7: push esi
  loc_00E279B8: mov var_E8, edx
  loc_00E279BE: call [eax+00000300h]
  loc_00E279C4: lea ecx, var_30
  loc_00E279C7: push eax
  loc_00E279C8: push ecx
  loc_00E279C9: call edi
  loc_00E279CB: mov edx, [eax]
  loc_00E279CD: lea ecx, var_28
  loc_00E279D0: push ecx
  loc_00E279D1: push eax
  loc_00E279D2: mov var_D0, eax
  loc_00E279D8: call [edx+000000A0h]
  loc_00E279DE: test eax, eax
  loc_00E279E0: fnclex
  loc_00E279E2: jge 00E279F8h
  loc_00E279E4: mov edx, var_D0
  loc_00E279EA: push 000000A0h
  loc_00E279EF: push 006DCB70h
  loc_00E279F4: push edx
  loc_00E279F5: push eax
  loc_00E279F6: call ebx
  loc_00E279F8: mov eax, var_28
  loc_00E279FB: mov ecx, var_E8
  loc_00E27A01: mov var_58, eax
  loc_00E27A04: mov eax, 00000008h
  loc_00E27A09: mov var_28, 00000000h
  loc_00E27A10: mov var_60, eax
  loc_00E27A13: mov edx, [ecx]
  loc_00E27A15: sub esp, 00000010h
  loc_00E27A18: mov ecx, esp
  loc_00E27A1A: mov [ecx], eax
  loc_00E27A1C: mov eax, var_5C
  loc_00E27A1F: mov [ecx+00000004h], eax
  loc_00E27A22: mov eax, var_58
  loc_00E27A25: mov [ecx+00000008h], eax
  loc_00E27A28: mov eax, var_54
  loc_00E27A2B: mov [ecx+0000000Ch], eax
  loc_00E27A2E: mov ecx, var_E8
  loc_00E27A34: push ecx
  loc_00E27A35: call [edx+00000038h]
  loc_00E27A38: test eax, eax
  loc_00E27A3A: fnclex
  loc_00E27A3C: jge 00E27A4Fh
  loc_00E27A3E: mov edx, var_E8
  loc_00E27A44: push 00000038h
  loc_00E27A46: push 006E09F8h
  loc_00E27A4B: push edx
  loc_00E27A4C: push eax
  loc_00E27A4D: call ebx
  loc_00E27A4F: lea eax, var_40
  loc_00E27A52: lea ecx, var_3C
  loc_00E27A55: push eax
  loc_00E27A56: lea edx, var_38
  loc_00E27A59: push ecx
  loc_00E27A5A: lea eax, var_34
  loc_00E27A5D: push edx
  loc_00E27A5E: lea ecx, var_30
  loc_00E27A61: push eax
  loc_00E27A62: push ecx
  loc_00E27A63: push 00000005h
  loc_00E27A65: call [00401048h] ; __vbaFreeObjList
  loc_00E27A6B: lea edx, var_60
  loc_00E27A6E: lea eax, var_50
  loc_00E27A71: push edx
  loc_00E27A72: push eax
  loc_00E27A73: push 00000002h
  loc_00E27A75: call [00401038h] ; __vbaFreeVarList
  loc_00E27A7B: mov ecx, [esi]
  loc_00E27A7D: add esp, 00000024h
  loc_00E27A80: push 006DCBD8h
  loc_00E27A85: push 00000000h
  loc_00E27A87: push 00000018h
  loc_00E27A89: push esi
  loc_00E27A8A: call [ecx+000004A8h]
  loc_00E27A90: lea edx, var_34
  loc_00E27A93: push eax
  loc_00E27A94: push edx
  loc_00E27A95: call edi
  loc_00E27A97: push eax
  loc_00E27A98: lea eax, var_50
  loc_00E27A9B: push eax
  loc_00E27A9C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27AA2: add esp, 00000010h
  loc_00E27AA5: push eax
  loc_00E27AA6: call [00401120h] ; __vbaCastObjVar
  loc_00E27AAC: lea ecx, var_38
  loc_00E27AAF: push eax
  loc_00E27AB0: push ecx
  loc_00E27AB1: call edi
  loc_00E27AB3: mov edx, [eax]
  loc_00E27AB5: lea ecx, var_3C
  loc_00E27AB8: push ecx
  loc_00E27AB9: push eax
  loc_00E27ABA: mov var_D8, eax
  loc_00E27AC0: call [edx+00000054h]
  loc_00E27AC3: test eax, eax
  loc_00E27AC5: fnclex
  loc_00E27AC7: jge 00E27ADAh
  loc_00E27AC9: mov edx, var_D8
  loc_00E27ACF: push 00000054h
  loc_00E27AD1: push 006DCBE8h
  loc_00E27AD6: push edx
  loc_00E27AD7: push eax
  loc_00E27AD8: call ebx
  loc_00E27ADA: lea edx, var_40
  loc_00E27ADD: mov eax, 00000002h
  loc_00E27AE2: push edx
  loc_00E27AE3: mov ecx, var_3C
  loc_00E27AE6: sub esp, 00000010h
  loc_00E27AE9: mov var_90, eax
  loc_00E27AEF: mov edx, esp
  loc_00E27AF1: mov var_88, 00000004h
  loc_00E27AFB: mov var_E0, ecx
  loc_00E27B01: mov ecx, [ecx]
  loc_00E27B03: mov [edx], eax
  loc_00E27B05: mov eax, var_8C
  loc_00E27B0B: mov [edx+00000004h], eax
  loc_00E27B0E: mov eax, var_88
  loc_00E27B14: mov [edx+00000008h], eax
  loc_00E27B17: mov eax, var_84
  loc_00E27B1D: mov [edx+0000000Ch], eax
  loc_00E27B20: mov edx, var_3C
  loc_00E27B23: push edx
  loc_00E27B24: call [ecx+00000028h]
  loc_00E27B27: test eax, eax
  loc_00E27B29: fnclex
  loc_00E27B2B: jge 00E27B3Eh
  loc_00E27B2D: mov ecx, var_E0
  loc_00E27B33: push 00000028h
  loc_00E27B35: push 006E09E8h
  loc_00E27B3A: push ecx
  loc_00E27B3B: push eax
  loc_00E27B3C: call ebx
  loc_00E27B3E: mov edx, var_40
  loc_00E27B41: mov eax, [esi]
  loc_00E27B43: push esi
  loc_00E27B44: mov var_E8, edx
  loc_00E27B4A: call [eax+00000450h]
  loc_00E27B50: lea ecx, var_30
  loc_00E27B53: push eax
  loc_00E27B54: push ecx
  loc_00E27B55: call edi
  loc_00E27B57: mov edx, [eax]
  loc_00E27B59: lea ecx, var_28
  loc_00E27B5C: push ecx
  loc_00E27B5D: push eax
  loc_00E27B5E: mov var_D0, eax
  loc_00E27B64: call [edx+00000050h]
  loc_00E27B67: test eax, eax
  loc_00E27B69: fnclex
  loc_00E27B6B: jge 00E27B7Eh
  loc_00E27B6D: mov edx, var_D0
  loc_00E27B73: push 00000050h
  loc_00E27B75: push 006E0410h
  loc_00E27B7A: push edx
  loc_00E27B7B: push eax
  loc_00E27B7C: call ebx
  loc_00E27B7E: mov eax, var_28
  loc_00E27B81: mov ecx, var_E8
  loc_00E27B87: mov var_58, eax
  loc_00E27B8A: mov eax, 00000008h
  loc_00E27B8F: mov var_28, 00000000h
  loc_00E27B96: mov var_60, eax
  loc_00E27B99: mov edx, [ecx]
  loc_00E27B9B: sub esp, 00000010h
  loc_00E27B9E: mov ecx, esp
  loc_00E27BA0: mov [ecx], eax
  loc_00E27BA2: mov eax, var_5C
  loc_00E27BA5: mov [ecx+00000004h], eax
  loc_00E27BA8: mov eax, var_58
  loc_00E27BAB: mov [ecx+00000008h], eax
  loc_00E27BAE: mov eax, var_54
  loc_00E27BB1: mov [ecx+0000000Ch], eax
  loc_00E27BB4: mov ecx, var_E8
  loc_00E27BBA: push ecx
  loc_00E27BBB: call [edx+00000038h]
  loc_00E27BBE: test eax, eax
  loc_00E27BC0: fnclex
  loc_00E27BC2: jge 00E27BD5h
  loc_00E27BC4: mov edx, var_E8
  loc_00E27BCA: push 00000038h
  loc_00E27BCC: push 006E09F8h
  loc_00E27BD1: push edx
  loc_00E27BD2: push eax
  loc_00E27BD3: call ebx
  loc_00E27BD5: lea eax, var_40
  loc_00E27BD8: lea ecx, var_3C
  loc_00E27BDB: push eax
  loc_00E27BDC: lea edx, var_38
  loc_00E27BDF: push ecx
  loc_00E27BE0: lea eax, var_34
  loc_00E27BE3: push edx
  loc_00E27BE4: lea ecx, var_30
  loc_00E27BE7: push eax
  loc_00E27BE8: push ecx
  loc_00E27BE9: push 00000005h
  loc_00E27BEB: call [00401048h] ; __vbaFreeObjList
  loc_00E27BF1: lea edx, var_60
  loc_00E27BF4: lea eax, var_50
  loc_00E27BF7: push edx
  loc_00E27BF8: push eax
  loc_00E27BF9: push 00000002h
  loc_00E27BFB: call [00401038h] ; __vbaFreeVarList
  loc_00E27C01: mov ecx, [esi]
  loc_00E27C03: add esp, 00000024h
  loc_00E27C06: push 006DCBD8h
  loc_00E27C0B: push 00000000h
  loc_00E27C0D: push 00000018h
  loc_00E27C0F: push esi
  loc_00E27C10: call [ecx+000004A8h]
  loc_00E27C16: lea edx, var_34
  loc_00E27C19: push eax
  loc_00E27C1A: push edx
  loc_00E27C1B: call edi
  loc_00E27C1D: push eax
  loc_00E27C1E: lea eax, var_50
  loc_00E27C21: push eax
  loc_00E27C22: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27C28: add esp, 00000010h
  loc_00E27C2B: push eax
  loc_00E27C2C: call [00401120h] ; __vbaCastObjVar
  loc_00E27C32: lea ecx, var_38
  loc_00E27C35: push eax
  loc_00E27C36: push ecx
  loc_00E27C37: call edi
  loc_00E27C39: mov edx, [eax]
  loc_00E27C3B: lea ecx, var_3C
  loc_00E27C3E: push ecx
  loc_00E27C3F: push eax
  loc_00E27C40: mov var_D8, eax
  loc_00E27C46: call [edx+00000054h]
  loc_00E27C49: test eax, eax
  loc_00E27C4B: fnclex
  loc_00E27C4D: jge 00E27C60h
  loc_00E27C4F: mov edx, var_D8
  loc_00E27C55: push 00000054h
  loc_00E27C57: push 006DCBE8h
  loc_00E27C5C: push edx
  loc_00E27C5D: push eax
  loc_00E27C5E: call ebx
  loc_00E27C60: lea edx, var_40
  loc_00E27C63: mov eax, 00000002h
  loc_00E27C68: push edx
  loc_00E27C69: mov ecx, var_3C
  loc_00E27C6C: sub esp, 00000010h
  loc_00E27C6F: mov var_90, eax
  loc_00E27C75: mov edx, esp
  loc_00E27C77: mov var_88, 00000005h
  loc_00E27C81: mov var_E0, ecx
  loc_00E27C87: mov ecx, [ecx]
  loc_00E27C89: mov [edx], eax
  loc_00E27C8B: mov eax, var_8C
  loc_00E27C91: mov [edx+00000004h], eax
  loc_00E27C94: mov eax, var_88
  loc_00E27C9A: mov [edx+00000008h], eax
  loc_00E27C9D: mov eax, var_84
  loc_00E27CA3: mov [edx+0000000Ch], eax
  loc_00E27CA6: mov edx, var_3C
  loc_00E27CA9: push edx
  loc_00E27CAA: call [ecx+00000028h]
  loc_00E27CAD: test eax, eax
  loc_00E27CAF: fnclex
  loc_00E27CB1: jge 00E27CC4h
  loc_00E27CB3: mov ecx, var_E0
  loc_00E27CB9: push 00000028h
  loc_00E27CBB: push 006E09E8h
  loc_00E27CC0: push ecx
  loc_00E27CC1: push eax
  loc_00E27CC2: call ebx
  loc_00E27CC4: mov edx, var_40
  loc_00E27CC7: mov eax, [esi]
  loc_00E27CC9: push esi
  loc_00E27CCA: mov var_E8, edx
  loc_00E27CD0: call [eax+00000424h]
  loc_00E27CD6: lea ecx, var_30
  loc_00E27CD9: push eax
  loc_00E27CDA: push ecx
  loc_00E27CDB: call edi
  loc_00E27CDD: mov edx, [eax]
  loc_00E27CDF: lea ecx, var_28
  loc_00E27CE2: push ecx
  loc_00E27CE3: push eax
  loc_00E27CE4: mov var_D0, eax
  loc_00E27CEA: call [edx+00000050h]
  loc_00E27CED: test eax, eax
  loc_00E27CEF: fnclex
  loc_00E27CF1: jge 00E27D04h
  loc_00E27CF3: mov edx, var_D0
  loc_00E27CF9: push 00000050h
  loc_00E27CFB: push 006E0410h
  loc_00E27D00: push edx
  loc_00E27D01: push eax
  loc_00E27D02: call ebx
  loc_00E27D04: mov eax, var_28
  loc_00E27D07: mov ecx, var_E8
  loc_00E27D0D: mov var_58, eax
  loc_00E27D10: mov eax, 00000008h
  loc_00E27D15: mov var_28, 00000000h
  loc_00E27D1C: mov var_60, eax
  loc_00E27D1F: mov edx, [ecx]
  loc_00E27D21: sub esp, 00000010h
  loc_00E27D24: mov ecx, esp
  loc_00E27D26: mov [ecx], eax
  loc_00E27D28: mov eax, var_5C
  loc_00E27D2B: mov [ecx+00000004h], eax
  loc_00E27D2E: mov eax, var_58
  loc_00E27D31: mov [ecx+00000008h], eax
  loc_00E27D34: mov eax, var_54
  loc_00E27D37: mov [ecx+0000000Ch], eax
  loc_00E27D3A: mov ecx, var_E8
  loc_00E27D40: push ecx
  loc_00E27D41: call [edx+00000038h]
  loc_00E27D44: test eax, eax
  loc_00E27D46: fnclex
  loc_00E27D48: jge 00E27D5Bh
  loc_00E27D4A: mov edx, var_E8
  loc_00E27D50: push 00000038h
  loc_00E27D52: push 006E09F8h
  loc_00E27D57: push edx
  loc_00E27D58: push eax
  loc_00E27D59: call ebx
  loc_00E27D5B: lea eax, var_40
  loc_00E27D5E: lea ecx, var_3C
  loc_00E27D61: push eax
  loc_00E27D62: lea edx, var_38
  loc_00E27D65: push ecx
  loc_00E27D66: lea eax, var_34
  loc_00E27D69: push edx
  loc_00E27D6A: lea ecx, var_30
  loc_00E27D6D: push eax
  loc_00E27D6E: push ecx
  loc_00E27D6F: push 00000005h
  loc_00E27D71: call [00401048h] ; __vbaFreeObjList
  loc_00E27D77: lea edx, var_60
  loc_00E27D7A: lea eax, var_50
  loc_00E27D7D: push edx
  loc_00E27D7E: push eax
  loc_00E27D7F: push 00000002h
  loc_00E27D81: call [00401038h] ; __vbaFreeVarList
  loc_00E27D87: mov ecx, [esi]
  loc_00E27D89: add esp, 00000024h
  loc_00E27D8C: push esi
  loc_00E27D8D: call [ecx+00000368h]
  loc_00E27D93: lea edx, var_34
  loc_00E27D96: push eax
  loc_00E27D97: push edx
  loc_00E27D98: call edi
  loc_00E27D9A: mov ecx, [eax]
  loc_00E27D9C: lea edx, var_2C
  loc_00E27D9F: push edx
  loc_00E27DA0: push eax
  loc_00E27DA1: mov var_D0, eax
  loc_00E27DA7: call [ecx+00000050h]
  loc_00E27DAA: test eax, eax
  loc_00E27DAC: fnclex
  loc_00E27DAE: jge 00E27DC1h
  loc_00E27DB0: mov ecx, var_D0
  loc_00E27DB6: push 00000050h
  loc_00E27DB8: push 006E0410h
  loc_00E27DBD: push ecx
  loc_00E27DBE: push eax
  loc_00E27DBF: call ebx
  loc_00E27DC1: mov edx, var_2C
  loc_00E27DC4: push edx
  loc_00E27DC5: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E27DCB: mov eax, [esi]
  loc_00E27DCD: push esi
  loc_00E27DCE: fstp real8 ptr var_CC
  loc_00E27DD4: call [eax+00000378h]
  loc_00E27DDA: lea ecx, var_30
  loc_00E27DDD: push eax
  loc_00E27DDE: push ecx
  loc_00E27DDF: call edi
  loc_00E27DE1: mov edx, [eax]
  loc_00E27DE3: lea ecx, var_28
  loc_00E27DE6: push ecx
  loc_00E27DE7: push eax
  loc_00E27DE8: mov var_D8, eax
  loc_00E27DEE: call [edx+000000A0h]
  loc_00E27DF4: test eax, eax
  loc_00E27DF6: fnclex
  loc_00E27DF8: jge 00E27E0Eh
  loc_00E27DFA: mov edx, var_D8
  loc_00E27E00: push 000000A0h
  loc_00E27E05: push 006DCB70h
  loc_00E27E0A: push edx
  loc_00E27E0B: push eax
  loc_00E27E0C: call ebx
  loc_00E27E0E: mov eax, var_28
  loc_00E27E11: push eax
  loc_00E27E12: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E27E18: call [004010D8h] ; __vbaFpR8
  loc_00E27E1E: fstp real8 ptr var_178
  loc_00E27E24: fld real8 ptr var_CC
  loc_00E27E2A: call [004010D8h] ; __vbaFpR8
  loc_00E27E30: fcomp real8 ptr var_178
  loc_00E27E36: mov var_17C, 00000001h
  loc_00E27E40: fnstsw ax
  loc_00E27E42: test ah, 01h
  loc_00E27E45: jnz 00E27E51h
  loc_00E27E47: mov var_17C, 00000000h
  loc_00E27E51: lea ecx, var_2C
  loc_00E27E54: lea edx, var_28
  loc_00E27E57: push ecx
  loc_00E27E58: push edx
  loc_00E27E59: push 00000002h
  loc_00E27E5B: call [004011BCh] ; __vbaFreeStrList
  loc_00E27E61: lea eax, var_34
  loc_00E27E64: lea ecx, var_30
  loc_00E27E67: push eax
  loc_00E27E68: push ecx
  loc_00E27E69: push 00000002h
  loc_00E27E6B: call [00401048h] ; __vbaFreeObjList
  loc_00E27E71: mov eax, var_17C
  loc_00E27E77: mov edx, [esi]
  loc_00E27E79: add esp, 00000018h
  loc_00E27E7C: neg eax
  loc_00E27E7E: push 006DCBD8h
  loc_00E27E83: push 00000000h
  loc_00E27E85: test ax, ax
  loc_00E27E88: push 00000018h
  loc_00E27E8A: push esi
  loc_00E27E8B: jz 00E28122h
  loc_00E27E91: call [edx+000004A8h]
  loc_00E27E97: push eax
  loc_00E27E98: lea eax, var_34
  loc_00E27E9B: push eax
  loc_00E27E9C: call edi
  loc_00E27E9E: lea ecx, var_50
  loc_00E27EA1: push eax
  loc_00E27EA2: push ecx
  loc_00E27EA3: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E27EA9: add esp, 00000010h
  loc_00E27EAC: push eax
  loc_00E27EAD: call [00401120h] ; __vbaCastObjVar
  loc_00E27EB3: lea edx, var_38
  loc_00E27EB6: push eax
  loc_00E27EB7: push edx
  loc_00E27EB8: call edi
  loc_00E27EBA: mov ecx, [eax]
  loc_00E27EBC: lea edx, var_3C
  loc_00E27EBF: push edx
  loc_00E27EC0: push eax
  loc_00E27EC1: mov var_D8, eax
  loc_00E27EC7: call [ecx+00000054h]
  loc_00E27ECA: test eax, eax
  loc_00E27ECC: fnclex
  loc_00E27ECE: jge 00E27EE1h
  loc_00E27ED0: mov ecx, var_D8
  loc_00E27ED6: push 00000054h
  loc_00E27ED8: push 006DCBE8h
  loc_00E27EDD: push ecx
  loc_00E27EDE: push eax
  loc_00E27EDF: call ebx
  loc_00E27EE1: mov ecx, var_3C
  loc_00E27EE4: mov eax, 00000002h
  loc_00E27EE9: mov var_88, 00000006h
  loc_00E27EF3: mov var_90, eax
  loc_00E27EF9: mov edx, [ecx]
  loc_00E27EFB: mov var_E0, ecx
  loc_00E27F01: lea ecx, var_40
  loc_00E27F04: push ecx
  loc_00E27F05: sub esp, 00000010h
  loc_00E27F08: mov ecx, esp
  loc_00E27F0A: mov [ecx], eax
  loc_00E27F0C: mov eax, var_8C
  loc_00E27F12: mov [ecx+00000004h], eax
  loc_00E27F15: mov eax, var_88
  loc_00E27F1B: mov [ecx+00000008h], eax
  loc_00E27F1E: mov eax, var_84
  loc_00E27F24: mov [ecx+0000000Ch], eax
  loc_00E27F27: mov ecx, var_3C
  loc_00E27F2A: push ecx
  loc_00E27F2B: call [edx+00000028h]
  loc_00E27F2E: test eax, eax
  loc_00E27F30: fnclex
  loc_00E27F32: jge 00E27F45h
  loc_00E27F34: mov edx, var_E0
  loc_00E27F3A: push 00000028h
  loc_00E27F3C: push 006E09E8h
  loc_00E27F41: push edx
  loc_00E27F42: push eax
  loc_00E27F43: call ebx
  loc_00E27F45: mov eax, var_40
  loc_00E27F48: mov ecx, [esi]
  loc_00E27F4A: push esi
  loc_00E27F4B: mov var_E8, eax
  loc_00E27F51: call [ecx+00000368h]
  loc_00E27F57: lea edx, var_30
  loc_00E27F5A: push eax
  loc_00E27F5B: push edx
  loc_00E27F5C: call edi
  loc_00E27F5E: mov ecx, [eax]
  loc_00E27F60: lea edx, var_28
  loc_00E27F63: push edx
  loc_00E27F64: push eax
  loc_00E27F65: mov var_D0, eax
  loc_00E27F6B: call [ecx+00000050h]
  loc_00E27F6E: test eax, eax
  loc_00E27F70: fnclex
  loc_00E27F72: jge 00E27F85h
  loc_00E27F74: mov ecx, var_D0
  loc_00E27F7A: push 00000050h
  loc_00E27F7C: push 006E0410h
  loc_00E27F81: push ecx
  loc_00E27F82: push eax
  loc_00E27F83: call ebx
  loc_00E27F85: mov eax, var_28
  loc_00E27F88: mov edx, var_E8
  loc_00E27F8E: mov var_58, eax
  loc_00E27F91: mov eax, 00000008h
  loc_00E27F96: mov var_28, 00000000h
  loc_00E27F9D: mov var_60, eax
  loc_00E27FA0: mov ecx, [edx]
  loc_00E27FA2: sub esp, 00000010h
  loc_00E27FA5: mov edx, esp
  loc_00E27FA7: mov [edx], eax
  loc_00E27FA9: mov eax, var_5C
  loc_00E27FAC: mov [edx+00000004h], eax
  loc_00E27FAF: mov eax, var_58
  loc_00E27FB2: mov [edx+00000008h], eax
  loc_00E27FB5: mov eax, var_54
  loc_00E27FB8: mov [edx+0000000Ch], eax
  loc_00E27FBB: mov edx, var_E8
  loc_00E27FC1: push edx
  loc_00E27FC2: call [ecx+00000038h]
  loc_00E27FC5: test eax, eax
  loc_00E27FC7: fnclex
  loc_00E27FC9: jge 00E27FDCh
  loc_00E27FCB: mov ecx, var_E8
  loc_00E27FD1: push 00000038h
  loc_00E27FD3: push 006E09F8h
  loc_00E27FD8: push ecx
  loc_00E27FD9: push eax
  loc_00E27FDA: call ebx
  loc_00E27FDC: lea edx, var_40
  loc_00E27FDF: lea eax, var_3C
  loc_00E27FE2: push edx
  loc_00E27FE3: lea ecx, var_38
  loc_00E27FE6: push eax
  loc_00E27FE7: lea edx, var_34
  loc_00E27FEA: push ecx
  loc_00E27FEB: lea eax, var_30
  loc_00E27FEE: push edx
  loc_00E27FEF: push eax
  loc_00E27FF0: push 00000005h
  loc_00E27FF2: call [00401048h] ; __vbaFreeObjList
  loc_00E27FF8: lea ecx, var_60
  loc_00E27FFB: lea edx, var_50
  loc_00E27FFE: push ecx
  loc_00E27FFF: push edx
  loc_00E28000: push 00000002h
  loc_00E28002: call [00401038h] ; __vbaFreeVarList
  loc_00E28008: mov eax, [esi]
  loc_00E2800A: add esp, 00000024h
  loc_00E2800D: push 006DCBD8h
  loc_00E28012: push 00000000h
  loc_00E28014: push 00000018h
  loc_00E28016: push esi
  loc_00E28017: call [eax+000004A8h]
  loc_00E2801D: lea ecx, var_30
  loc_00E28020: push eax
  loc_00E28021: push ecx
  loc_00E28022: call edi
  loc_00E28024: lea edx, var_50
  loc_00E28027: push eax
  loc_00E28028: push edx
  loc_00E28029: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2802F: add esp, 00000010h
  loc_00E28032: push eax
  loc_00E28033: call [00401120h] ; __vbaCastObjVar
  loc_00E28039: push eax
  loc_00E2803A: lea eax, var_34
  loc_00E2803D: push eax
  loc_00E2803E: call edi
  loc_00E28040: mov ecx, [eax]
  loc_00E28042: lea edx, var_38
  loc_00E28045: push edx
  loc_00E28046: push eax
  loc_00E28047: mov var_D0, eax
  loc_00E2804D: call [ecx+00000054h]
  loc_00E28050: test eax, eax
  loc_00E28052: fnclex
  loc_00E28054: jge 00E28067h
  loc_00E28056: mov ecx, var_D0
  loc_00E2805C: push 00000054h
  loc_00E2805E: push 006DCBE8h
  loc_00E28063: push ecx
  loc_00E28064: push eax
  loc_00E28065: call ebx
  loc_00E28067: mov ecx, var_38
  loc_00E2806A: mov eax, 00000002h
  loc_00E2806F: mov var_88, 0000000Fh
  loc_00E28079: mov var_90, eax
  loc_00E2807F: mov edx, [ecx]
  loc_00E28081: mov var_D8, ecx
  loc_00E28087: lea ecx, var_3C
  loc_00E2808A: push ecx
  loc_00E2808B: sub esp, 00000010h
  loc_00E2808E: mov ecx, esp
  loc_00E28090: mov [ecx], eax
  loc_00E28092: mov eax, var_8C
  loc_00E28098: mov [ecx+00000004h], eax
  loc_00E2809B: mov eax, var_88
  loc_00E280A1: mov [ecx+00000008h], eax
  loc_00E280A4: mov eax, var_84
  loc_00E280AA: mov [ecx+0000000Ch], eax
  loc_00E280AD: mov ecx, var_38
  loc_00E280B0: push ecx
  loc_00E280B1: call [edx+00000028h]
  loc_00E280B4: test eax, eax
  loc_00E280B6: fnclex
  loc_00E280B8: jge 00E280CBh
  loc_00E280BA: mov edx, var_D8
  loc_00E280C0: push 00000028h
  loc_00E280C2: push 006E09E8h
  loc_00E280C7: push edx
  loc_00E280C8: push eax
  loc_00E280C9: call ebx
  loc_00E280CB: mov eax, var_3C
  loc_00E280CE: mov ecx, [esi]
  loc_00E280D0: push esi
  loc_00E280D1: mov var_E0, eax
  loc_00E280D7: call [ecx+000002FCh]
  loc_00E280DD: mov edx, var_E0
  loc_00E280E3: mov var_58, eax
  loc_00E280E6: mov eax, 00000009h
  loc_00E280EB: sub esp, 00000010h
  loc_00E280EE: mov var_60, eax
  loc_00E280F1: mov ecx, [edx]
  loc_00E280F3: mov edx, esp
  loc_00E280F5: mov [edx], eax
  loc_00E280F7: mov eax, var_5C
  loc_00E280FA: mov [edx+00000004h], eax
  loc_00E280FD: mov eax, var_58
  loc_00E28100: mov [edx+00000008h], eax
  loc_00E28103: mov eax, var_54
  loc_00E28106: mov [edx+0000000Ch], eax
  loc_00E28109: mov edx, var_E0
  loc_00E2810F: push edx
  loc_00E28110: call [ecx+00000038h]
  loc_00E28113: test eax, eax
  loc_00E28115: fnclex
  loc_00E28117: jge 00E2837Fh
  loc_00E2811D: jmp 00E2836Eh
  loc_00E28122: call [edx+000004A8h]
  loc_00E28128: push eax
  loc_00E28129: lea eax, var_30
  loc_00E2812C: push eax
  loc_00E2812D: call edi
  loc_00E2812F: lea ecx, var_50
  loc_00E28132: push eax
  loc_00E28133: push ecx
  loc_00E28134: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2813A: add esp, 00000010h
  loc_00E2813D: push eax
  loc_00E2813E: call [00401120h] ; __vbaCastObjVar
  loc_00E28144: lea edx, var_34
  loc_00E28147: push eax
  loc_00E28148: push edx
  loc_00E28149: call edi
  loc_00E2814B: mov ecx, [eax]
  loc_00E2814D: lea edx, var_38
  loc_00E28150: push edx
  loc_00E28151: push eax
  loc_00E28152: mov var_D0, eax
  loc_00E28158: call [ecx+00000054h]
  loc_00E2815B: test eax, eax
  loc_00E2815D: fnclex
  loc_00E2815F: jge 00E28172h
  loc_00E28161: mov ecx, var_D0
  loc_00E28167: push 00000054h
  loc_00E28169: push 006DCBE8h
  loc_00E2816E: push ecx
  loc_00E2816F: push eax
  loc_00E28170: call ebx
  loc_00E28172: mov ecx, var_38
  loc_00E28175: mov eax, 00000002h
  loc_00E2817A: mov var_88, 00000006h
  loc_00E28184: mov var_90, eax
  loc_00E2818A: mov edx, [ecx]
  loc_00E2818C: mov var_D8, ecx
  loc_00E28192: lea ecx, var_3C
  loc_00E28195: push ecx
  loc_00E28196: sub esp, 00000010h
  loc_00E28199: mov ecx, esp
  loc_00E2819B: mov [ecx], eax
  loc_00E2819D: mov eax, var_8C
  loc_00E281A3: mov [ecx+00000004h], eax
  loc_00E281A6: mov eax, var_88
  loc_00E281AC: mov [ecx+00000008h], eax
  loc_00E281AF: mov eax, var_84
  loc_00E281B5: mov [ecx+0000000Ch], eax
  loc_00E281B8: mov ecx, var_38
  loc_00E281BB: push ecx
  loc_00E281BC: call [edx+00000028h]
  loc_00E281BF: test eax, eax
  loc_00E281C1: fnclex
  loc_00E281C3: jge 00E281D6h
  loc_00E281C5: mov edx, var_D8
  loc_00E281CB: push 00000028h
  loc_00E281CD: push 006E09E8h
  loc_00E281D2: push edx
  loc_00E281D3: push eax
  loc_00E281D4: call ebx
  loc_00E281D6: mov eax, var_3C
  loc_00E281D9: mov ecx, [esi]
  loc_00E281DB: push esi
  loc_00E281DC: mov var_E0, eax
  loc_00E281E2: call [ecx+000002FCh]
  loc_00E281E8: mov edx, var_E0
  loc_00E281EE: mov var_58, eax
  loc_00E281F1: mov eax, 00000009h
  loc_00E281F6: sub esp, 00000010h
  loc_00E281F9: mov var_60, eax
  loc_00E281FC: mov ecx, [edx]
  loc_00E281FE: mov edx, esp
  loc_00E28200: mov [edx], eax
  loc_00E28202: mov eax, var_5C
  loc_00E28205: mov [edx+00000004h], eax
  loc_00E28208: mov eax, var_58
  loc_00E2820B: mov [edx+00000008h], eax
  loc_00E2820E: mov eax, var_54
  loc_00E28211: mov [edx+0000000Ch], eax
  loc_00E28214: mov edx, var_E0
  loc_00E2821A: push edx
  loc_00E2821B: call [ecx+00000038h]
  loc_00E2821E: test eax, eax
  loc_00E28220: fnclex
  loc_00E28222: jge 00E28235h
  loc_00E28224: mov ecx, var_E0
  loc_00E2822A: push 00000038h
  loc_00E2822C: push 006E09F8h
  loc_00E28231: push ecx
  loc_00E28232: push eax
  loc_00E28233: call ebx
  loc_00E28235: lea edx, var_3C
  loc_00E28238: lea eax, var_38
  loc_00E2823B: push edx
  loc_00E2823C: lea ecx, var_34
  loc_00E2823F: push eax
  loc_00E28240: lea edx, var_30
  loc_00E28243: push ecx
  loc_00E28244: push edx
  loc_00E28245: push 00000004h
  loc_00E28247: call [00401048h] ; __vbaFreeObjList
  loc_00E2824D: lea eax, var_60
  loc_00E28250: lea ecx, var_50
  loc_00E28253: push eax
  loc_00E28254: push ecx
  loc_00E28255: push 00000002h
  loc_00E28257: call [00401038h] ; __vbaFreeVarList
  loc_00E2825D: mov edx, [esi]
  loc_00E2825F: add esp, 00000020h
  loc_00E28262: push 006DCBD8h
  loc_00E28267: push 00000000h
  loc_00E28269: push 00000018h
  loc_00E2826B: push esi
  loc_00E2826C: call [edx+000004A8h]
  loc_00E28272: push eax
  loc_00E28273: lea eax, var_30
  loc_00E28276: push eax
  loc_00E28277: call edi
  loc_00E28279: lea ecx, var_50
  loc_00E2827C: push eax
  loc_00E2827D: push ecx
  loc_00E2827E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E28284: add esp, 00000010h
  loc_00E28287: push eax
  loc_00E28288: call [00401120h] ; __vbaCastObjVar
  loc_00E2828E: lea edx, var_34
  loc_00E28291: push eax
  loc_00E28292: push edx
  loc_00E28293: call edi
  loc_00E28295: mov ecx, [eax]
  loc_00E28297: lea edx, var_38
  loc_00E2829A: push edx
  loc_00E2829B: push eax
  loc_00E2829C: mov var_D0, eax
  loc_00E282A2: call [ecx+00000054h]
  loc_00E282A5: test eax, eax
  loc_00E282A7: fnclex
  loc_00E282A9: jge 00E282BCh
  loc_00E282AB: mov ecx, var_D0
  loc_00E282B1: push 00000054h
  loc_00E282B3: push 006DCBE8h
  loc_00E282B8: push ecx
  loc_00E282B9: push eax
  loc_00E282BA: call ebx
  loc_00E282BC: mov ecx, var_38
  loc_00E282BF: mov eax, 00000002h
  loc_00E282C4: mov var_88, 0000000Fh
  loc_00E282CE: mov var_90, eax
  loc_00E282D4: mov edx, [ecx]
  loc_00E282D6: mov var_D8, ecx
  loc_00E282DC: lea ecx, var_3C
  loc_00E282DF: push ecx
  loc_00E282E0: sub esp, 00000010h
  loc_00E282E3: mov ecx, esp
  loc_00E282E5: mov [ecx], eax
  loc_00E282E7: mov eax, var_8C
  loc_00E282ED: mov [ecx+00000004h], eax
  loc_00E282F0: mov eax, var_88
  loc_00E282F6: mov [ecx+00000008h], eax
  loc_00E282F9: mov eax, var_84
  loc_00E282FF: mov [ecx+0000000Ch], eax
  loc_00E28302: mov ecx, var_38
  loc_00E28305: push ecx
  loc_00E28306: call [edx+00000028h]
  loc_00E28309: test eax, eax
  loc_00E2830B: fnclex
  loc_00E2830D: jge 00E28320h
  loc_00E2830F: mov edx, var_D8
  loc_00E28315: push 00000028h
  loc_00E28317: push 006E09E8h
  loc_00E2831C: push edx
  loc_00E2831D: push eax
  loc_00E2831E: call ebx
  loc_00E28320: mov eax, var_3C
  loc_00E28323: mov ecx, [esi]
  loc_00E28325: push esi
  loc_00E28326: mov var_E0, eax
  loc_00E2832C: call [ecx+000002FCh]
  loc_00E28332: mov edx, var_E0
  loc_00E28338: mov var_58, eax
  loc_00E2833B: mov eax, 00000009h
  loc_00E28340: sub esp, 00000010h
  loc_00E28343: mov var_60, eax
  loc_00E28346: mov ecx, [edx]
  loc_00E28348: mov edx, esp
  loc_00E2834A: mov [edx], eax
  loc_00E2834C: mov eax, var_5C
  loc_00E2834F: mov [edx+00000004h], eax
  loc_00E28352: mov eax, var_58
  loc_00E28355: mov [edx+00000008h], eax
  loc_00E28358: mov eax, var_54
  loc_00E2835B: mov [edx+0000000Ch], eax
  loc_00E2835E: mov edx, var_E0
  loc_00E28364: push edx
  loc_00E28365: call [ecx+00000038h]
  loc_00E28368: test eax, eax
  loc_00E2836A: fnclex
  loc_00E2836C: jge 00E2837Fh
  loc_00E2836E: mov ecx, var_E0
  loc_00E28374: push 00000038h
  loc_00E28376: push 006E09F8h
  loc_00E2837B: push ecx
  loc_00E2837C: push eax
  loc_00E2837D: call ebx
  loc_00E2837F: lea edx, var_3C
  loc_00E28382: lea eax, var_38
  loc_00E28385: push edx
  loc_00E28386: lea ecx, var_34
  loc_00E28389: push eax
  loc_00E2838A: lea edx, var_30
  loc_00E2838D: push ecx
  loc_00E2838E: push edx
  loc_00E2838F: push 00000004h
  loc_00E28391: call [00401048h] ; __vbaFreeObjList
  loc_00E28397: lea eax, var_60
  loc_00E2839A: lea ecx, var_50
  loc_00E2839D: push eax
  loc_00E2839E: push ecx
  loc_00E2839F: push 00000002h
  loc_00E283A1: call [00401038h] ; __vbaFreeVarList
  loc_00E283A7: mov edx, [esi]
  loc_00E283A9: add esp, 00000020h
  loc_00E283AC: push 006DCBD8h
  loc_00E283B1: push 00000000h
  loc_00E283B3: push 00000018h
  loc_00E283B5: push esi
  loc_00E283B6: call [edx+000004A8h]
  loc_00E283BC: push eax
  loc_00E283BD: lea eax, var_34
  loc_00E283C0: push eax
  loc_00E283C1: call edi
  loc_00E283C3: lea ecx, var_50
  loc_00E283C6: push eax
  loc_00E283C7: push ecx
  loc_00E283C8: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E283CE: add esp, 00000010h
  loc_00E283D1: push eax
  loc_00E283D2: call [00401120h] ; __vbaCastObjVar
  loc_00E283D8: lea edx, var_38
  loc_00E283DB: push eax
  loc_00E283DC: push edx
  loc_00E283DD: call edi
  loc_00E283DF: mov ecx, [eax]
  loc_00E283E1: lea edx, var_3C
  loc_00E283E4: push edx
  loc_00E283E5: push eax
  loc_00E283E6: mov var_D8, eax
  loc_00E283EC: call [ecx+00000054h]
  loc_00E283EF: test eax, eax
  loc_00E283F1: fnclex
  loc_00E283F3: jge 00E28406h
  loc_00E283F5: mov ecx, var_D8
  loc_00E283FB: push 00000054h
  loc_00E283FD: push 006DCBE8h
  loc_00E28402: push ecx
  loc_00E28403: push eax
  loc_00E28404: call ebx
  loc_00E28406: mov ecx, var_3C
  loc_00E28409: mov eax, 00000002h
  loc_00E2840E: mov var_88, 00000008h
  loc_00E28418: mov var_90, eax
  loc_00E2841E: mov edx, [ecx]
  loc_00E28420: mov var_E0, ecx
  loc_00E28426: lea ecx, var_40
  loc_00E28429: push ecx
  loc_00E2842A: sub esp, 00000010h
  loc_00E2842D: mov ecx, esp
  loc_00E2842F: mov [ecx], eax
  loc_00E28431: mov eax, var_8C
  loc_00E28437: mov [ecx+00000004h], eax
  loc_00E2843A: mov eax, var_88
  loc_00E28440: mov [ecx+00000008h], eax
  loc_00E28443: mov eax, var_84
  loc_00E28449: mov [ecx+0000000Ch], eax
  loc_00E2844C: mov ecx, var_3C
  loc_00E2844F: push ecx
  loc_00E28450: call [edx+00000028h]
  loc_00E28453: test eax, eax
  loc_00E28455: fnclex
  loc_00E28457: jge 00E2846Ah
  loc_00E28459: mov edx, var_E0
  loc_00E2845F: push 00000028h
  loc_00E28461: push 006E09E8h
  loc_00E28466: push edx
  loc_00E28467: push eax
  loc_00E28468: call ebx
  loc_00E2846A: mov eax, var_40
  loc_00E2846D: mov ecx, [esi]
  loc_00E2846F: push esi
  loc_00E28470: mov var_E8, eax
  loc_00E28476: call [ecx+00000368h]
  loc_00E2847C: lea edx, var_30
  loc_00E2847F: push eax
  loc_00E28480: push edx
  loc_00E28481: call edi
  loc_00E28483: mov ecx, [eax]
  loc_00E28485: lea edx, var_28
  loc_00E28488: push edx
  loc_00E28489: push eax
  loc_00E2848A: mov var_D0, eax
  loc_00E28490: call [ecx+00000050h]
  loc_00E28493: test eax, eax
  loc_00E28495: fnclex
  loc_00E28497: jge 00E284AAh
  loc_00E28499: mov ecx, var_D0
  loc_00E2849F: push 00000050h
  loc_00E284A1: push 006E0410h
  loc_00E284A6: push ecx
  loc_00E284A7: push eax
  loc_00E284A8: call ebx
  loc_00E284AA: mov eax, var_28
  loc_00E284AD: mov edx, var_E8
  loc_00E284B3: mov var_58, eax
  loc_00E284B6: mov eax, 00000008h
  loc_00E284BB: mov var_28, 00000000h
  loc_00E284C2: mov var_60, eax
  loc_00E284C5: mov ecx, [edx]
  loc_00E284C7: sub esp, 00000010h
  loc_00E284CA: mov edx, esp
  loc_00E284CC: mov [edx], eax
  loc_00E284CE: mov eax, var_5C
  loc_00E284D1: mov [edx+00000004h], eax
  loc_00E284D4: mov eax, var_58
  loc_00E284D7: mov [edx+00000008h], eax
  loc_00E284DA: mov eax, var_54
  loc_00E284DD: mov [edx+0000000Ch], eax
  loc_00E284E0: mov edx, var_E8
  loc_00E284E6: push edx
  loc_00E284E7: call [ecx+00000038h]
  loc_00E284EA: test eax, eax
  loc_00E284EC: fnclex
  loc_00E284EE: jge 00E28501h
  loc_00E284F0: mov ecx, var_E8
  loc_00E284F6: push 00000038h
  loc_00E284F8: push 006E09F8h
  loc_00E284FD: push ecx
  loc_00E284FE: push eax
  loc_00E284FF: call ebx
  loc_00E28501: lea edx, var_40
  loc_00E28504: lea eax, var_3C
  loc_00E28507: push edx
  loc_00E28508: lea ecx, var_38
  loc_00E2850B: push eax
  loc_00E2850C: lea edx, var_34
  loc_00E2850F: push ecx
  loc_00E28510: lea eax, var_30
  loc_00E28513: push edx
  loc_00E28514: push eax
  loc_00E28515: push 00000005h
  loc_00E28517: call [00401048h] ; __vbaFreeObjList
  loc_00E2851D: lea ecx, var_60
  loc_00E28520: lea edx, var_50
  loc_00E28523: push ecx
  loc_00E28524: push edx
  loc_00E28525: push 00000002h
  loc_00E28527: call [00401038h] ; __vbaFreeVarList
  loc_00E2852D: mov eax, [esi]
  loc_00E2852F: add esp, 00000024h
  loc_00E28532: push 006DCBD8h
  loc_00E28537: push 00000000h
  loc_00E28539: push 00000018h
  loc_00E2853B: push esi
  loc_00E2853C: call [eax+000004A8h]
  loc_00E28542: lea ecx, var_34
  loc_00E28545: push eax
  loc_00E28546: push ecx
  loc_00E28547: call edi
  loc_00E28549: lea edx, var_50
  loc_00E2854C: push eax
  loc_00E2854D: push edx
  loc_00E2854E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E28554: add esp, 00000010h
  loc_00E28557: push eax
  loc_00E28558: call [00401120h] ; __vbaCastObjVar
  loc_00E2855E: push eax
  loc_00E2855F: lea eax, var_38
  loc_00E28562: push eax
  loc_00E28563: call edi
  loc_00E28565: mov ecx, [eax]
  loc_00E28567: lea edx, var_3C
  loc_00E2856A: push edx
  loc_00E2856B: push eax
  loc_00E2856C: mov var_D8, eax
  loc_00E28572: call [ecx+00000054h]
  loc_00E28575: test eax, eax
  loc_00E28577: fnclex
  loc_00E28579: jge 00E2858Ch
  loc_00E2857B: mov ecx, var_D8
  loc_00E28581: push 00000054h
  loc_00E28583: push 006DCBE8h
  loc_00E28588: push ecx
  loc_00E28589: push eax
  loc_00E2858A: call ebx
  loc_00E2858C: mov ecx, var_3C
  loc_00E2858F: mov eax, 00000002h
  loc_00E28594: mov var_88, 0000000Bh
  loc_00E2859E: mov var_90, eax
  loc_00E285A4: mov edx, [ecx]
  loc_00E285A6: mov var_E0, ecx
  loc_00E285AC: lea ecx, var_40
  loc_00E285AF: push ecx
  loc_00E285B0: sub esp, 00000010h
  loc_00E285B3: mov ecx, esp
  loc_00E285B5: mov [ecx], eax
  loc_00E285B7: mov eax, var_8C
  loc_00E285BD: mov [ecx+00000004h], eax
  loc_00E285C0: mov eax, var_88
  loc_00E285C6: mov [ecx+00000008h], eax
  loc_00E285C9: mov eax, var_84
  loc_00E285CF: mov [ecx+0000000Ch], eax
  loc_00E285D2: mov ecx, var_3C
  loc_00E285D5: push ecx
  loc_00E285D6: call [edx+00000028h]
  loc_00E285D9: test eax, eax
  loc_00E285DB: fnclex
  loc_00E285DD: jge 00E285F0h
  loc_00E285DF: mov edx, var_E0
  loc_00E285E5: push 00000028h
  loc_00E285E7: push 006E09E8h
  loc_00E285EC: push edx
  loc_00E285ED: push eax
  loc_00E285EE: call ebx
  loc_00E285F0: mov eax, var_40
  loc_00E285F3: mov ecx, [esi]
  loc_00E285F5: push esi
  loc_00E285F6: mov var_E8, eax
  loc_00E285FC: call [ecx+00000304h]
  loc_00E28602: lea edx, var_30
  loc_00E28605: push eax
  loc_00E28606: push edx
  loc_00E28607: call edi
  loc_00E28609: mov ecx, [eax]
  loc_00E2860B: lea edx, var_28
  loc_00E2860E: push edx
  loc_00E2860F: push eax
  loc_00E28610: mov var_D0, eax
  loc_00E28616: call [ecx+000000A0h]
  loc_00E2861C: test eax, eax
  loc_00E2861E: fnclex
  loc_00E28620: jge 00E28636h
  loc_00E28622: mov ecx, var_D0
  loc_00E28628: push 000000A0h
  loc_00E2862D: push 006DCB70h
  loc_00E28632: push ecx
  loc_00E28633: push eax
  loc_00E28634: call ebx
  loc_00E28636: mov eax, var_28
  loc_00E28639: mov edx, var_E8
  loc_00E2863F: mov var_58, eax
  loc_00E28642: mov eax, 00000008h
  loc_00E28647: mov var_28, 00000000h
  loc_00E2864E: mov var_60, eax
  loc_00E28651: mov ecx, [edx]
  loc_00E28653: sub esp, 00000010h
  loc_00E28656: mov edx, esp
  loc_00E28658: mov [edx], eax
  loc_00E2865A: mov eax, var_5C
  loc_00E2865D: mov [edx+00000004h], eax
  loc_00E28660: mov eax, var_58
  loc_00E28663: mov [edx+00000008h], eax
  loc_00E28666: mov eax, var_54
  loc_00E28669: mov [edx+0000000Ch], eax
  loc_00E2866C: mov edx, var_E8
  loc_00E28672: push edx
  loc_00E28673: call [ecx+00000038h]
  loc_00E28676: test eax, eax
  loc_00E28678: fnclex
  loc_00E2867A: jge 00E2868Dh
  loc_00E2867C: mov ecx, var_E8
  loc_00E28682: push 00000038h
  loc_00E28684: push 006E09F8h
  loc_00E28689: push ecx
  loc_00E2868A: push eax
  loc_00E2868B: call ebx
  loc_00E2868D: lea edx, var_40
  loc_00E28690: lea eax, var_3C
  loc_00E28693: push edx
  loc_00E28694: lea ecx, var_38
  loc_00E28697: push eax
  loc_00E28698: lea edx, var_34
  loc_00E2869B: push ecx
  loc_00E2869C: lea eax, var_30
  loc_00E2869F: push edx
  loc_00E286A0: push eax
  loc_00E286A1: push 00000005h
  loc_00E286A3: call [00401048h] ; __vbaFreeObjList
  loc_00E286A9: lea ecx, var_60
  loc_00E286AC: lea edx, var_50
  loc_00E286AF: push ecx
  loc_00E286B0: push edx
  loc_00E286B1: push 00000002h
  loc_00E286B3: call [00401038h] ; __vbaFreeVarList
  loc_00E286B9: mov eax, [esi]
  loc_00E286BB: add esp, 00000024h
  loc_00E286BE: push 006DCBD8h
  loc_00E286C3: push 00000000h
  loc_00E286C5: push 00000018h
  loc_00E286C7: push esi
  loc_00E286C8: call [eax+000004A8h]
  loc_00E286CE: lea ecx, var_34
  loc_00E286D1: push eax
  loc_00E286D2: push ecx
  loc_00E286D3: call edi
  loc_00E286D5: lea edx, var_50
  loc_00E286D8: push eax
  loc_00E286D9: push edx
  loc_00E286DA: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E286E0: add esp, 00000010h
  loc_00E286E3: push eax
  loc_00E286E4: call [00401120h] ; __vbaCastObjVar
  loc_00E286EA: push eax
  loc_00E286EB: lea eax, var_38
  loc_00E286EE: push eax
  loc_00E286EF: call edi
  loc_00E286F1: mov ecx, [eax]
  loc_00E286F3: lea edx, var_3C
  loc_00E286F6: push edx
  loc_00E286F7: push eax
  loc_00E286F8: mov var_D8, eax
  loc_00E286FE: call [ecx+00000054h]
  loc_00E28701: test eax, eax
  loc_00E28703: fnclex
  loc_00E28705: jge 00E28718h
  loc_00E28707: mov ecx, var_D8
  loc_00E2870D: push 00000054h
  loc_00E2870F: push 006DCBE8h
  loc_00E28714: push ecx
  loc_00E28715: push eax
  loc_00E28716: call ebx
  loc_00E28718: mov ecx, var_3C
  loc_00E2871B: mov eax, 00000002h
  loc_00E28720: mov var_88, 0000000Ch
  loc_00E2872A: mov var_90, eax
  loc_00E28730: mov edx, [ecx]
  loc_00E28732: mov var_E0, ecx
  loc_00E28738: lea ecx, var_40
  loc_00E2873B: push ecx
  loc_00E2873C: sub esp, 00000010h
  loc_00E2873F: mov ecx, esp
  loc_00E28741: mov [ecx], eax
  loc_00E28743: mov eax, var_8C
  loc_00E28749: mov [ecx+00000004h], eax
  loc_00E2874C: mov eax, var_88
  loc_00E28752: mov [ecx+00000008h], eax
  loc_00E28755: mov eax, var_84
  loc_00E2875B: mov [ecx+0000000Ch], eax
  loc_00E2875E: mov ecx, var_3C
  loc_00E28761: push ecx
  loc_00E28762: call [edx+00000028h]
  loc_00E28765: test eax, eax
  loc_00E28767: fnclex
  loc_00E28769: jge 00E2877Ch
  loc_00E2876B: mov edx, var_E0
  loc_00E28771: push 00000028h
  loc_00E28773: push 006E09E8h
  loc_00E28778: push edx
  loc_00E28779: push eax
  loc_00E2877A: call ebx
  loc_00E2877C: mov eax, var_40
  loc_00E2877F: mov ecx, [esi]
  loc_00E28781: push esi
  loc_00E28782: mov var_E8, eax
  loc_00E28788: call [ecx+00000354h]
  loc_00E2878E: lea edx, var_30
  loc_00E28791: push eax
  loc_00E28792: push edx
  loc_00E28793: call edi
  loc_00E28795: mov ecx, [eax]
  loc_00E28797: lea edx, var_28
  loc_00E2879A: push edx
  loc_00E2879B: push eax
  loc_00E2879C: mov var_D0, eax
  loc_00E287A2: call [ecx+00000050h]
  loc_00E287A5: test eax, eax
  loc_00E287A7: fnclex
  loc_00E287A9: jge 00E287BCh
  loc_00E287AB: mov ecx, var_D0
  loc_00E287B1: push 00000050h
  loc_00E287B3: push 006E0410h
  loc_00E287B8: push ecx
  loc_00E287B9: push eax
  loc_00E287BA: call ebx
  loc_00E287BC: mov eax, var_28
  loc_00E287BF: mov edx, var_E8
  loc_00E287C5: mov var_58, eax
  loc_00E287C8: mov eax, 00000008h
  loc_00E287CD: mov var_28, 00000000h
  loc_00E287D4: mov var_60, eax
  loc_00E287D7: mov ecx, [edx]
  loc_00E287D9: sub esp, 00000010h
  loc_00E287DC: mov edx, esp
  loc_00E287DE: mov [edx], eax
  loc_00E287E0: mov eax, var_5C
  loc_00E287E3: mov [edx+00000004h], eax
  loc_00E287E6: mov eax, var_58
  loc_00E287E9: mov [edx+00000008h], eax
  loc_00E287EC: mov eax, var_54
  loc_00E287EF: mov [edx+0000000Ch], eax
  loc_00E287F2: mov edx, var_E8
  loc_00E287F8: push edx
  loc_00E287F9: call [ecx+00000038h]
  loc_00E287FC: test eax, eax
  loc_00E287FE: fnclex
  loc_00E28800: jge 00E28813h
  loc_00E28802: mov ecx, var_E8
  loc_00E28808: push 00000038h
  loc_00E2880A: push 006E09F8h
  loc_00E2880F: push ecx
  loc_00E28810: push eax
  loc_00E28811: call ebx
  loc_00E28813: lea edx, var_40
  loc_00E28816: lea eax, var_3C
  loc_00E28819: push edx
  loc_00E2881A: lea ecx, var_38
  loc_00E2881D: push eax
  loc_00E2881E: lea edx, var_34
  loc_00E28821: push ecx
  loc_00E28822: lea eax, var_30
  loc_00E28825: push edx
  loc_00E28826: push eax
  loc_00E28827: push 00000005h
  loc_00E28829: call [00401048h] ; __vbaFreeObjList
  loc_00E2882F: lea ecx, var_60
  loc_00E28832: lea edx, var_50
  loc_00E28835: push ecx
  loc_00E28836: push edx
  loc_00E28837: push 00000002h
  loc_00E28839: call [00401038h] ; __vbaFreeVarList
  loc_00E2883F: mov eax, [esi]
  loc_00E28841: add esp, 00000024h
  loc_00E28844: push 006DCBD8h
  loc_00E28849: push 00000000h
  loc_00E2884B: push 00000018h
  loc_00E2884D: push esi
  loc_00E2884E: call [eax+000004A8h]
  loc_00E28854: lea ecx, var_34
  loc_00E28857: push eax
  loc_00E28858: push ecx
  loc_00E28859: call edi
  loc_00E2885B: lea edx, var_50
  loc_00E2885E: push eax
  loc_00E2885F: push edx
  loc_00E28860: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E28866: add esp, 00000010h
  loc_00E28869: push eax
  loc_00E2886A: call [00401120h] ; __vbaCastObjVar
  loc_00E28870: push eax
  loc_00E28871: lea eax, var_38
  loc_00E28874: push eax
  loc_00E28875: call edi
  loc_00E28877: mov ecx, [eax]
  loc_00E28879: lea edx, var_3C
  loc_00E2887C: push edx
  loc_00E2887D: push eax
  loc_00E2887E: mov var_D8, eax
  loc_00E28884: call [ecx+00000054h]
  loc_00E28887: test eax, eax
  loc_00E28889: fnclex
  loc_00E2888B: jge 00E2889Eh
  loc_00E2888D: mov ecx, var_D8
  loc_00E28893: push 00000054h
  loc_00E28895: push 006DCBE8h
  loc_00E2889A: push ecx
  loc_00E2889B: push eax
  loc_00E2889C: call ebx
  loc_00E2889E: mov ecx, var_3C
  loc_00E288A1: mov eax, 00000002h
  loc_00E288A6: mov var_88, 0000000Dh
  loc_00E288B0: mov var_90, eax
  loc_00E288B6: mov edx, [ecx]
  loc_00E288B8: mov var_E0, ecx
  loc_00E288BE: lea ecx, var_40
  loc_00E288C1: push ecx
  loc_00E288C2: sub esp, 00000010h
  loc_00E288C5: mov ecx, esp
  loc_00E288C7: mov [ecx], eax
  loc_00E288C9: mov eax, var_8C
  loc_00E288CF: mov [ecx+00000004h], eax
  loc_00E288D2: mov eax, var_88
  loc_00E288D8: mov [ecx+00000008h], eax
  loc_00E288DB: mov eax, var_84
  loc_00E288E1: mov [ecx+0000000Ch], eax
  loc_00E288E4: mov ecx, var_3C
  loc_00E288E7: push ecx
  loc_00E288E8: call [edx+00000028h]
  loc_00E288EB: test eax, eax
  loc_00E288ED: fnclex
  loc_00E288EF: jge 00E28902h
  loc_00E288F1: mov edx, var_E0
  loc_00E288F7: push 00000028h
  loc_00E288F9: push 006E09E8h
  loc_00E288FE: push edx
  loc_00E288FF: push eax
  loc_00E28900: call ebx
  loc_00E28902: mov eax, var_40
  loc_00E28905: mov ecx, [esi]
  loc_00E28907: push esi
  loc_00E28908: mov var_E8, eax
  loc_00E2890E: call [ecx+000003B0h]
  loc_00E28914: lea edx, var_30
  loc_00E28917: push eax
  loc_00E28918: push edx
  loc_00E28919: call edi
  loc_00E2891B: mov ecx, [eax]
  loc_00E2891D: lea edx, var_28
  loc_00E28920: push edx
  loc_00E28921: push eax
  loc_00E28922: mov var_D0, eax
  loc_00E28928: call [ecx+00000050h]
  loc_00E2892B: test eax, eax
  loc_00E2892D: fnclex
  loc_00E2892F: jge 00E28942h
  loc_00E28931: mov ecx, var_D0
  loc_00E28937: push 00000050h
  loc_00E28939: push 006E0410h
  loc_00E2893E: push ecx
  loc_00E2893F: push eax
  loc_00E28940: call ebx
  loc_00E28942: mov eax, var_28
  loc_00E28945: mov edx, var_E8
  loc_00E2894B: mov var_58, eax
  loc_00E2894E: mov eax, 00000008h
  loc_00E28953: mov var_28, 00000000h
  loc_00E2895A: mov var_60, eax
  loc_00E2895D: mov ecx, [edx]
  loc_00E2895F: sub esp, 00000010h
  loc_00E28962: mov edx, esp
  loc_00E28964: mov [edx], eax
  loc_00E28966: mov eax, var_5C
  loc_00E28969: mov [edx+00000004h], eax
  loc_00E2896C: mov eax, var_58
  loc_00E2896F: mov [edx+00000008h], eax
  loc_00E28972: mov eax, var_54
  loc_00E28975: mov [edx+0000000Ch], eax
  loc_00E28978: mov edx, var_E8
  loc_00E2897E: push edx
  loc_00E2897F: call [ecx+00000038h]
  loc_00E28982: test eax, eax
  loc_00E28984: fnclex
  loc_00E28986: jge 00E28999h
  loc_00E28988: mov ecx, var_E8
  loc_00E2898E: push 00000038h
  loc_00E28990: push 006E09F8h
  loc_00E28995: push ecx
  loc_00E28996: push eax
  loc_00E28997: call ebx
  loc_00E28999: lea edx, var_40
  loc_00E2899C: lea eax, var_3C
  loc_00E2899F: push edx
  loc_00E289A0: lea ecx, var_38
  loc_00E289A3: push eax
  loc_00E289A4: lea edx, var_34
  loc_00E289A7: push ecx
  loc_00E289A8: lea eax, var_30
  loc_00E289AB: push edx
  loc_00E289AC: push eax
  loc_00E289AD: push 00000005h
  loc_00E289AF: call [00401048h] ; __vbaFreeObjList
  loc_00E289B5: lea ecx, var_60
  loc_00E289B8: lea edx, var_50
  loc_00E289BB: push ecx
  loc_00E289BC: push edx
  loc_00E289BD: push 00000002h
  loc_00E289BF: call [00401038h] ; __vbaFreeVarList
  loc_00E289C5: mov eax, [esi]
  loc_00E289C7: add esp, 00000024h
  loc_00E289CA: push esi
  loc_00E289CB: call [eax+0000037Ch]
  loc_00E289D1: lea ecx, var_30
  loc_00E289D4: push eax
  loc_00E289D5: push ecx
  loc_00E289D6: call edi
  loc_00E289D8: mov edx, [eax]
  loc_00E289DA: lea ecx, var_C4
  loc_00E289E0: push ecx
  loc_00E289E1: push eax
  loc_00E289E2: mov var_D0, eax
  loc_00E289E8: call [edx+00000098h]
  loc_00E289EE: test eax, eax
  loc_00E289F0: fnclex
  loc_00E289F2: jge 00E28A08h
  loc_00E289F4: mov edx, var_D0
  loc_00E289FA: push 00000098h
  loc_00E289FF: push 006DCAD0h
  loc_00E28A04: push edx
  loc_00E28A05: push eax
  loc_00E28A06: call ebx
  loc_00E28A08: xor eax, eax
  loc_00E28A0A: cmp var_C4, FFFFFFh
  loc_00E28A12: lea ecx, var_30
  loc_00E28A15: setz al
  loc_00E28A18: neg eax
  loc_00E28A1A: mov var_D8, ax
  loc_00E28A21: call [00401254h] ; __vbaFreeObj
  loc_00E28A27: cmp var_D8, 0000h
  loc_00E28A2F: jz 00E28BC0h
  loc_00E28A35: mov ecx, [esi]
  loc_00E28A37: push 006DCBD8h
  loc_00E28A3C: push 00000000h
  loc_00E28A3E: push 00000018h
  loc_00E28A40: push esi
  loc_00E28A41: call [ecx+000004A8h]
  loc_00E28A47: lea edx, var_34
  loc_00E28A4A: push eax
  loc_00E28A4B: push edx
  loc_00E28A4C: call edi
  loc_00E28A4E: push eax
  loc_00E28A4F: lea eax, var_50
  loc_00E28A52: push eax
  loc_00E28A53: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E28A59: add esp, 00000010h
  loc_00E28A5C: push eax
  loc_00E28A5D: call [00401120h] ; __vbaCastObjVar
  loc_00E28A63: lea ecx, var_38
  loc_00E28A66: push eax
  loc_00E28A67: push ecx
  loc_00E28A68: call edi
  loc_00E28A6A: mov edx, [eax]
  loc_00E28A6C: lea ecx, var_3C
  loc_00E28A6F: push ecx
  loc_00E28A70: push eax
  loc_00E28A71: mov var_D8, eax
  loc_00E28A77: call [edx+00000054h]
  loc_00E28A7A: test eax, eax
  loc_00E28A7C: fnclex
  loc_00E28A7E: jge 00E28A91h
  loc_00E28A80: mov edx, var_D8
  loc_00E28A86: push 00000054h
  loc_00E28A88: push 006DCBE8h
  loc_00E28A8D: push edx
  loc_00E28A8E: push eax
  loc_00E28A8F: call ebx
  loc_00E28A91: lea edx, var_40
  loc_00E28A94: mov eax, 00000002h
  loc_00E28A99: push edx
  loc_00E28A9A: mov ecx, var_3C
  loc_00E28A9D: sub esp, 00000010h
  loc_00E28AA0: mov var_90, eax
  loc_00E28AA6: mov edx, esp
  loc_00E28AA8: mov var_88, 00000009h
  loc_00E28AB2: mov var_E0, ecx
  loc_00E28AB8: mov ecx, [ecx]
  loc_00E28ABA: mov [edx], eax
  loc_00E28ABC: mov eax, var_8C
  loc_00E28AC2: mov [edx+00000004h], eax
  loc_00E28AC5: mov eax, var_88
  loc_00E28ACB: mov [edx+00000008h], eax
  loc_00E28ACE: mov eax, var_84
  loc_00E28AD4: mov [edx+0000000Ch], eax
  loc_00E28AD7: mov edx, var_3C
  loc_00E28ADA: push edx
  loc_00E28ADB: call [ecx+00000028h]
  loc_00E28ADE: test eax, eax
  loc_00E28AE0: fnclex
  loc_00E28AE2: jge 00E28AF5h
  loc_00E28AE4: mov ecx, var_E0
  loc_00E28AEA: push 00000028h
  loc_00E28AEC: push 006E09E8h
  loc_00E28AF1: push ecx
  loc_00E28AF2: push eax
  loc_00E28AF3: call ebx
  loc_00E28AF5: mov edx, var_40
  loc_00E28AF8: mov eax, [esi]
  loc_00E28AFA: push esi
  loc_00E28AFB: mov var_E8, edx
  loc_00E28B01: call [eax+00000388h]
  loc_00E28B07: lea ecx, var_30
  loc_00E28B0A: push eax
  loc_00E28B0B: push ecx
  loc_00E28B0C: call edi
  loc_00E28B0E: mov edx, [eax]
  loc_00E28B10: lea ecx, var_28
  loc_00E28B13: push ecx
  loc_00E28B14: push eax
  loc_00E28B15: mov var_D0, eax
  loc_00E28B1B: call [edx+00000050h]
  loc_00E28B1E: test eax, eax
  loc_00E28B20: fnclex
  loc_00E28B22: jge 00E28B35h
  loc_00E28B24: mov edx, var_D0
  loc_00E28B2A: push 00000050h
  loc_00E28B2C: push 006E0410h
  loc_00E28B31: push edx
  loc_00E28B32: push eax
  loc_00E28B33: call ebx
  loc_00E28B35: mov eax, var_28
  loc_00E28B38: mov ecx, var_E8
  loc_00E28B3E: mov var_58, eax
  loc_00E28B41: mov eax, 00000008h
  loc_00E28B46: mov var_28, 00000000h
  loc_00E28B4D: mov var_60, eax
  loc_00E28B50: mov edx, [ecx]
  loc_00E28B52: sub esp, 00000010h
  loc_00E28B55: mov ecx, esp
  loc_00E28B57: mov [ecx], eax
  loc_00E28B59: mov eax, var_5C
  loc_00E28B5C: mov [ecx+00000004h], eax
  loc_00E28B5F: mov eax, var_58
  loc_00E28B62: mov [ecx+00000008h], eax
  loc_00E28B65: mov eax, var_54
  loc_00E28B68: mov [ecx+0000000Ch], eax
  loc_00E28B6B: mov ecx, var_E8
  loc_00E28B71: push ecx
  loc_00E28B72: call [edx+00000038h]
  loc_00E28B75: test eax, eax
  loc_00E28B77: fnclex
  loc_00E28B79: jge 00E28B8Ch
  loc_00E28B7B: mov edx, var_E8
  loc_00E28B81: push 00000038h
  loc_00E28B83: push 006E09F8h
  loc_00E28B88: push edx
  loc_00E28B89: push eax
  loc_00E28B8A: call ebx
  loc_00E28B8C: lea eax, var_40
  loc_00E28B8F: lea ecx, var_3C
  loc_00E28B92: push eax
  loc_00E28B93: lea edx, var_38
  loc_00E28B96: push ecx
  loc_00E28B97: lea eax, var_34
  loc_00E28B9A: push edx
  loc_00E28B9B: lea ecx, var_30
  loc_00E28B9E: push eax
  loc_00E28B9F: push ecx
  loc_00E28BA0: push 00000005h
  loc_00E28BA2: call [00401048h] ; __vbaFreeObjList
  loc_00E28BA8: lea edx, var_60
  loc_00E28BAB: lea eax, var_50
  loc_00E28BAE: push edx
  loc_00E28BAF: push eax
  loc_00E28BB0: push 00000002h
  loc_00E28BB2: call [00401038h] ; __vbaFreeVarList
  loc_00E28BB8: add esp, 00000024h
  loc_00E28BBB: jmp 00E28F3Eh
  loc_00E28BC0: mov ecx, [esi]
  loc_00E28BC2: push esi
  loc_00E28BC3: call [ecx+0000037Ch]
  loc_00E28BC9: lea edx, var_30
  loc_00E28BCC: push eax
  loc_00E28BCD: push edx
  loc_00E28BCE: call edi
  loc_00E28BD0: mov ecx, [eax]
  loc_00E28BD2: lea edx, var_C4
  loc_00E28BD8: push edx
  loc_00E28BD9: push eax
  loc_00E28BDA: mov var_D0, eax
  loc_00E28BE0: call [ecx+00000098h]
  loc_00E28BE6: test eax, eax
  loc_00E28BE8: fnclex
  loc_00E28BEA: jge 00E28C00h
  loc_00E28BEC: mov ecx, var_D0
  loc_00E28BF2: push 00000098h
  loc_00E28BF7: push 006DCAD0h
  loc_00E28BFC: push ecx
  loc_00E28BFD: push eax
  loc_00E28BFE: call ebx
  loc_00E28C00: xor edx, edx
  loc_00E28C02: lea ecx, var_30
  loc_00E28C05: cmp var_C4, dx
  loc_00E28C0C: setz dl
  loc_00E28C0F: neg edx
  loc_00E28C11: mov var_D8, dx
  loc_00E28C18: call [00401254h] ; __vbaFreeObj
  loc_00E28C1E: cmp var_D8, 0000h
  loc_00E28C26: jz 00E28F3Eh
  loc_00E28C2C: mov eax, [esi]
  loc_00E28C2E: push esi
  loc_00E28C2F: call [eax+000003B0h]
  loc_00E28C35: lea ecx, var_30
  loc_00E28C38: push eax
  loc_00E28C39: push ecx
  loc_00E28C3A: call edi
  loc_00E28C3C: mov edx, [eax]
  loc_00E28C3E: lea ecx, var_28
  loc_00E28C41: push ecx
  loc_00E28C42: push eax
  loc_00E28C43: mov var_D0, eax
  loc_00E28C49: call [edx+00000050h]
  loc_00E28C4C: test eax, eax
  loc_00E28C4E: fnclex
  loc_00E28C50: jge 00E28C63h
  loc_00E28C52: mov edx, var_D0
  loc_00E28C58: push 00000050h
  loc_00E28C5A: push 006E0410h
  loc_00E28C5F: push edx
  loc_00E28C60: push eax
  loc_00E28C61: call ebx
  loc_00E28C63: mov eax, var_28
  loc_00E28C66: push eax
  loc_00E28C67: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E28C6D: call [004010D8h] ; __vbaFpR8
  loc_00E28C73: fcomp real8 ptr [004022E8h]
  loc_00E28C79: mov var_180, 00000001h
  loc_00E28C83: fnstsw ax
  loc_00E28C85: test ah, 41h
  loc_00E28C88: jnz 00E28C94h
  loc_00E28C8A: mov var_180, 00000000h
  loc_00E28C94: lea ecx, var_28
  loc_00E28C97: call [00401258h] ; __vbaFreeStr
  loc_00E28C9D: lea ecx, var_30
  loc_00E28CA0: call [00401254h] ; __vbaFreeObj
  loc_00E28CA6: mov eax, var_180
  loc_00E28CAC: mov ecx, [esi]
  loc_00E28CAE: neg eax
  loc_00E28CB0: push 006DCBD8h
  loc_00E28CB5: push 00000000h
  loc_00E28CB7: test ax, ax
  loc_00E28CBA: push 00000018h
  loc_00E28CBC: push esi
  loc_00E28CBD: jz 00E28E06h
  loc_00E28CC3: call [ecx+000004A8h]
  loc_00E28CC9: lea edx, var_34
  loc_00E28CCC: push eax
  loc_00E28CCD: push edx
  loc_00E28CCE: call edi
  loc_00E28CD0: push eax
  loc_00E28CD1: lea eax, var_50
  loc_00E28CD4: push eax
  loc_00E28CD5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E28CDB: add esp, 00000010h
  loc_00E28CDE: push eax
  loc_00E28CDF: call [00401120h] ; __vbaCastObjVar
  loc_00E28CE5: lea ecx, var_38
  loc_00E28CE8: push eax
  loc_00E28CE9: push ecx
  loc_00E28CEA: call edi
  loc_00E28CEC: mov edx, [eax]
  loc_00E28CEE: lea ecx, var_3C
  loc_00E28CF1: push ecx
  loc_00E28CF2: push eax
  loc_00E28CF3: mov var_D8, eax
  loc_00E28CF9: call [edx+00000054h]
  loc_00E28CFC: test eax, eax
  loc_00E28CFE: fnclex
  loc_00E28D00: jge 00E28D13h
  loc_00E28D02: mov edx, var_D8
  loc_00E28D08: push 00000054h
  loc_00E28D0A: push 006DCBE8h
  loc_00E28D0F: push edx
  loc_00E28D10: push eax
  loc_00E28D11: call ebx
  loc_00E28D13: lea edx, var_40
  loc_00E28D16: mov eax, 00000002h
  loc_00E28D1B: push edx
  loc_00E28D1C: mov ecx, var_3C
  loc_00E28D1F: sub esp, 00000010h
  loc_00E28D22: mov var_90, eax
  loc_00E28D28: mov edx, esp
  loc_00E28D2A: mov var_88, 00000009h
  loc_00E28D34: mov var_E0, ecx
  loc_00E28D3A: mov ecx, [ecx]
  loc_00E28D3C: mov [edx], eax
  loc_00E28D3E: mov eax, var_8C
  loc_00E28D44: mov [edx+00000004h], eax
  loc_00E28D47: mov eax, var_88
  loc_00E28D4D: mov [edx+00000008h], eax
  loc_00E28D50: mov eax, var_84
  loc_00E28D56: mov [edx+0000000Ch], eax
  loc_00E28D59: mov edx, var_3C
  loc_00E28D5C: push edx
  loc_00E28D5D: call [ecx+00000028h]
  loc_00E28D60: test eax, eax
  loc_00E28D62: fnclex
  loc_00E28D64: jge 00E28D77h
  loc_00E28D66: mov ecx, var_E0
  loc_00E28D6C: push 00000028h
  loc_00E28D6E: push 006E09E8h
  loc_00E28D73: push ecx
  loc_00E28D74: push eax
  loc_00E28D75: call ebx
  loc_00E28D77: mov edx, var_40
  loc_00E28D7A: mov eax, [esi]
  loc_00E28D7C: push esi
  loc_00E28D7D: mov var_E8, edx
  loc_00E28D83: call [eax+00000388h]
  loc_00E28D89: lea ecx, var_30
  loc_00E28D8C: push eax
  loc_00E28D8D: push ecx
  loc_00E28D8E: call edi
  loc_00E28D90: mov edx, [eax]
  loc_00E28D92: lea ecx, var_28
  loc_00E28D95: push ecx
  loc_00E28D96: push eax
  loc_00E28D97: mov var_D0, eax
  loc_00E28D9D: call [edx+00000050h]
  loc_00E28DA0: test eax, eax
  loc_00E28DA2: fnclex
  loc_00E28DA4: jge 00E28DB7h
  loc_00E28DA6: mov edx, var_D0
  loc_00E28DAC: push 00000050h
  loc_00E28DAE: push 006E0410h
  loc_00E28DB3: push edx
  loc_00E28DB4: push eax
  loc_00E28DB5: call ebx
  loc_00E28DB7: mov eax, var_28
  loc_00E28DBA: mov ecx, var_E8
  loc_00E28DC0: mov var_58, eax
  loc_00E28DC3: mov eax, 00000008h
  loc_00E28DC8: mov var_28, 00000000h
  loc_00E28DCF: mov var_60, eax
  loc_00E28DD2: mov edx, [ecx]
  loc_00E28DD4: sub esp, 00000010h
  loc_00E28DD7: mov ecx, esp
  loc_00E28DD9: mov [ecx], eax
  loc_00E28DDB: mov eax, var_5C
  loc_00E28DDE: mov [ecx+00000004h], eax
  loc_00E28DE1: mov eax, var_58
  loc_00E28DE4: mov [ecx+00000008h], eax
  loc_00E28DE7: mov eax, var_54
  loc_00E28DEA: mov [ecx+0000000Ch], eax
  loc_00E28DED: mov ecx, var_E8
  loc_00E28DF3: push ecx
  loc_00E28DF4: call [edx+00000038h]
  loc_00E28DF7: test eax, eax
  loc_00E28DF9: fnclex
  loc_00E28DFB: jge 00E28B8Ch
  loc_00E28E01: jmp 00E28B7Bh
  loc_00E28E06: call [ecx+000004A8h]
  loc_00E28E0C: lea edx, var_30
  loc_00E28E0F: push eax
  loc_00E28E10: push edx
  loc_00E28E11: call edi
  loc_00E28E13: push eax
  loc_00E28E14: lea eax, var_50
  loc_00E28E17: push eax
  loc_00E28E18: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E28E1E: add esp, 00000010h
  loc_00E28E21: push eax
  loc_00E28E22: call [00401120h] ; __vbaCastObjVar
  loc_00E28E28: lea ecx, var_34
  loc_00E28E2B: push eax
  loc_00E28E2C: push ecx
  loc_00E28E2D: call edi
  loc_00E28E2F: mov edx, [eax]
  loc_00E28E31: lea ecx, var_38
  loc_00E28E34: push ecx
  loc_00E28E35: push eax
  loc_00E28E36: mov var_D0, eax
  loc_00E28E3C: call [edx+00000054h]
  loc_00E28E3F: test eax, eax
  loc_00E28E41: fnclex
  loc_00E28E43: jge 00E28E56h
  loc_00E28E45: mov edx, var_D0
  loc_00E28E4B: push 00000054h
  loc_00E28E4D: push 006DCBE8h
  loc_00E28E52: push edx
  loc_00E28E53: push eax
  loc_00E28E54: call ebx
  loc_00E28E56: lea edx, var_3C
  loc_00E28E59: mov eax, 00000002h
  loc_00E28E5E: push edx
  loc_00E28E5F: mov ecx, var_38
  loc_00E28E62: sub esp, 00000010h
  loc_00E28E65: mov var_90, eax
  loc_00E28E6B: mov edx, esp
  loc_00E28E6D: mov var_88, 00000009h
  loc_00E28E77: mov var_D8, ecx
  loc_00E28E7D: mov ecx, [ecx]
  loc_00E28E7F: mov [edx], eax
  loc_00E28E81: mov eax, var_8C
  loc_00E28E87: mov [edx+00000004h], eax
  loc_00E28E8A: mov eax, var_88
  loc_00E28E90: mov [edx+00000008h], eax
  loc_00E28E93: mov eax, var_84
  loc_00E28E99: mov [edx+0000000Ch], eax
  loc_00E28E9C: mov edx, var_38
  loc_00E28E9F: push edx
  loc_00E28EA0: call [ecx+00000028h]
  loc_00E28EA3: test eax, eax
  loc_00E28EA5: fnclex
  loc_00E28EA7: jge 00E28EBAh
  loc_00E28EA9: mov ecx, var_D8
  loc_00E28EAF: push 00000028h
  loc_00E28EB1: push 006E09E8h
  loc_00E28EB6: push ecx
  loc_00E28EB7: push eax
  loc_00E28EB8: call ebx
  loc_00E28EBA: mov ecx, var_3C
  loc_00E28EBD: mov eax, 00000002h
  loc_00E28EC2: mov var_98, 00000000h
  loc_00E28ECC: mov var_A0, eax
  loc_00E28ED2: mov edx, [ecx]
  loc_00E28ED4: sub esp, 00000010h
  loc_00E28ED7: mov var_E0, ecx
  loc_00E28EDD: mov ecx, esp
  loc_00E28EDF: mov [ecx], eax
  loc_00E28EE1: mov eax, var_9C
  loc_00E28EE7: mov [ecx+00000004h], eax
  loc_00E28EEA: mov eax, var_98
  loc_00E28EF0: mov [ecx+00000008h], eax
  loc_00E28EF3: mov eax, var_94
  loc_00E28EF9: mov [ecx+0000000Ch], eax
  loc_00E28EFC: mov ecx, var_3C
  loc_00E28EFF: push ecx
  loc_00E28F00: call [edx+00000038h]
  loc_00E28F03: test eax, eax
  loc_00E28F05: fnclex
  loc_00E28F07: jge 00E28F1Ah
  loc_00E28F09: mov edx, var_E0
  loc_00E28F0F: push 00000038h
  loc_00E28F11: push 006E09F8h
  loc_00E28F16: push edx
  loc_00E28F17: push eax
  loc_00E28F18: call ebx
  loc_00E28F1A: lea eax, var_3C
  loc_00E28F1D: lea ecx, var_38
  loc_00E28F20: push eax
  loc_00E28F21: lea edx, var_34
  loc_00E28F24: push ecx
  loc_00E28F25: lea eax, var_30
  loc_00E28F28: push edx
  loc_00E28F29: push eax
  loc_00E28F2A: push 00000004h
  loc_00E28F2C: call [00401048h] ; __vbaFreeObjList
  loc_00E28F32: add esp, 00000014h
  loc_00E28F35: lea ecx, var_50
  loc_00E28F38: call [00401024h] ; __vbaFreeVar
  loc_00E28F3E: mov ecx, [esi]
  loc_00E28F40: push esi
  loc_00E28F41: call [ecx+00000398h]
  loc_00E28F47: lea edx, var_30
  loc_00E28F4A: push eax
  loc_00E28F4B: push edx
  loc_00E28F4C: call edi
  loc_00E28F4E: mov ecx, [eax]
  loc_00E28F50: lea edx, var_C4
  loc_00E28F56: push edx
  loc_00E28F57: push eax
  loc_00E28F58: mov var_D0, eax
  loc_00E28F5E: call [ecx+00000098h]
  loc_00E28F64: test eax, eax
  loc_00E28F66: fnclex
  loc_00E28F68: jge 00E28F7Eh
  loc_00E28F6A: mov ecx, var_D0
  loc_00E28F70: push 00000098h
  loc_00E28F75: push 006DCAD0h
  loc_00E28F7A: push ecx
  loc_00E28F7B: push eax
  loc_00E28F7C: call ebx
  loc_00E28F7E: xor edx, edx
  loc_00E28F80: cmp var_C4, FFFFFFh
  loc_00E28F88: lea ecx, var_30
  loc_00E28F8B: setz dl
  loc_00E28F8E: neg edx
  loc_00E28F90: mov var_D8, dx
  loc_00E28F97: call [00401254h] ; __vbaFreeObj
  loc_00E28F9D: cmp var_D8, 0000h
  loc_00E28FA5: mov eax, [esi]
  loc_00E28FA7: push esi
  loc_00E28FA8: jz 00E298A1h
  loc_00E28FAE: call [eax+000003B0h]
  loc_00E28FB4: lea ecx, var_30
  loc_00E28FB7: push eax
  loc_00E28FB8: push ecx
  loc_00E28FB9: call edi
  loc_00E28FBB: mov edx, [eax]
  loc_00E28FBD: lea ecx, var_28
  loc_00E28FC0: push ecx
  loc_00E28FC1: push eax
  loc_00E28FC2: mov var_D0, eax
  loc_00E28FC8: call [edx+00000050h]
  loc_00E28FCB: test eax, eax
  loc_00E28FCD: fnclex
  loc_00E28FCF: jge 00E28FE2h
  loc_00E28FD1: mov edx, var_D0
  loc_00E28FD7: push 00000050h
  loc_00E28FD9: push 006E0410h
  loc_00E28FDE: push edx
  loc_00E28FDF: push eax
  loc_00E28FE0: call ebx
  loc_00E28FE2: mov eax, var_28
  loc_00E28FE5: push eax
  loc_00E28FE6: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E28FEC: call [004010D8h] ; __vbaFpR8
  loc_00E28FF2: fcomp real8 ptr [004022E8h]
  loc_00E28FF8: mov var_184, 00000001h
  loc_00E29002: fnstsw ax
  loc_00E29004: test ah, 40h
  loc_00E29007: jnz 00E29013h
  loc_00E29009: mov var_184, 00000000h
  loc_00E29013: lea ecx, var_28
  loc_00E29016: call [00401258h] ; __vbaFreeStr
  loc_00E2901C: lea ecx, var_30
  loc_00E2901F: call [00401254h] ; __vbaFreeObj
  loc_00E29025: mov eax, var_184
  loc_00E2902B: neg eax
  loc_00E2902D: test ax, ax
  loc_00E29030: jz 00E29292h
  loc_00E29036: mov ecx, [esi]
  loc_00E29038: push 006DCBD8h
  loc_00E2903D: push 00000000h
  loc_00E2903F: push 00000018h
  loc_00E29041: push esi
  loc_00E29042: call [ecx+000004A8h]
  loc_00E29048: lea edx, var_30
  loc_00E2904B: push eax
  loc_00E2904C: push edx
  loc_00E2904D: call edi
  loc_00E2904F: push eax
  loc_00E29050: lea eax, var_50
  loc_00E29053: push eax
  loc_00E29054: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2905A: add esp, 00000010h
  loc_00E2905D: push eax
  loc_00E2905E: call [00401120h] ; __vbaCastObjVar
  loc_00E29064: lea ecx, var_34
  loc_00E29067: push eax
  loc_00E29068: push ecx
  loc_00E29069: call edi
  loc_00E2906B: mov edx, [eax]
  loc_00E2906D: lea ecx, var_38
  loc_00E29070: push ecx
  loc_00E29071: push eax
  loc_00E29072: mov var_D0, eax
  loc_00E29078: call [edx+00000054h]
  loc_00E2907B: test eax, eax
  loc_00E2907D: fnclex
  loc_00E2907F: jge 00E29092h
  loc_00E29081: mov edx, var_D0
  loc_00E29087: push 00000054h
  loc_00E29089: push 006DCBE8h
  loc_00E2908E: push edx
  loc_00E2908F: push eax
  loc_00E29090: call ebx
  loc_00E29092: lea edx, var_3C
  loc_00E29095: mov eax, 00000002h
  loc_00E2909A: push edx
  loc_00E2909B: mov ecx, var_38
  loc_00E2909E: sub esp, 00000010h
  loc_00E290A1: mov var_90, eax
  loc_00E290A7: mov edx, esp
  loc_00E290A9: mov var_88, 0000000Ah
  loc_00E290B3: mov var_D8, ecx
  loc_00E290B9: mov ecx, [ecx]
  loc_00E290BB: mov [edx], eax
  loc_00E290BD: mov eax, var_8C
  loc_00E290C3: mov [edx+00000004h], eax
  loc_00E290C6: mov eax, var_88
  loc_00E290CC: mov [edx+00000008h], eax
  loc_00E290CF: mov eax, var_84
  loc_00E290D5: mov [edx+0000000Ch], eax
  loc_00E290D8: mov edx, var_38
  loc_00E290DB: push edx
  loc_00E290DC: call [ecx+00000028h]
  loc_00E290DF: test eax, eax
  loc_00E290E1: fnclex
  loc_00E290E3: jge 00E290F6h
  loc_00E290E5: mov ecx, var_D8
  loc_00E290EB: push 00000028h
  loc_00E290ED: push 006E09E8h
  loc_00E290F2: push ecx
  loc_00E290F3: push eax
  loc_00E290F4: call ebx
  loc_00E290F6: mov ecx, var_3C
  loc_00E290F9: mov eax, 00000002h
  loc_00E290FE: mov var_98, 00000000h
  loc_00E29108: mov var_A0, eax
  loc_00E2910E: mov edx, [ecx]
  loc_00E29110: sub esp, 00000010h
  loc_00E29113: mov var_E0, ecx
  loc_00E29119: mov ecx, esp
  loc_00E2911B: mov [ecx], eax
  loc_00E2911D: mov eax, var_9C
  loc_00E29123: mov [ecx+00000004h], eax
  loc_00E29126: mov eax, var_98
  loc_00E2912C: mov [ecx+00000008h], eax
  loc_00E2912F: mov eax, var_94
  loc_00E29135: mov [ecx+0000000Ch], eax
  loc_00E29138: mov ecx, var_3C
  loc_00E2913B: push ecx
  loc_00E2913C: call [edx+00000038h]
  loc_00E2913F: test eax, eax
  loc_00E29141: fnclex
  loc_00E29143: jge 00E29156h
  loc_00E29145: mov edx, var_E0
  loc_00E2914B: push 00000038h
  loc_00E2914D: push 006E09F8h
  loc_00E29152: push edx
  loc_00E29153: push eax
  loc_00E29154: call ebx
  loc_00E29156: lea eax, var_3C
  loc_00E29159: lea ecx, var_38
  loc_00E2915C: push eax
  loc_00E2915D: lea edx, var_34
  loc_00E29160: push ecx
  loc_00E29161: lea eax, var_30
  loc_00E29164: push edx
  loc_00E29165: push eax
  loc_00E29166: push 00000004h
  loc_00E29168: call [00401048h] ; __vbaFreeObjList
  loc_00E2916E: add esp, 00000014h
  loc_00E29171: lea ecx, var_50
  loc_00E29174: call [00401024h] ; __vbaFreeVar
  loc_00E2917A: mov ecx, [esi]
  loc_00E2917C: push 006DCBD8h
  loc_00E29181: push 00000000h
  loc_00E29183: push 00000018h
  loc_00E29185: push esi
  loc_00E29186: call [ecx+000004A8h]
  loc_00E2918C: lea edx, var_30
  loc_00E2918F: push eax
  loc_00E29190: push edx
  loc_00E29191: call edi
  loc_00E29193: push eax
  loc_00E29194: lea eax, var_50
  loc_00E29197: push eax
  loc_00E29198: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2919E: add esp, 00000010h
  loc_00E291A1: push eax
  loc_00E291A2: call [00401120h] ; __vbaCastObjVar
  loc_00E291A8: lea ecx, var_34
  loc_00E291AB: push eax
  loc_00E291AC: push ecx
  loc_00E291AD: call edi
  loc_00E291AF: mov edx, [eax]
  loc_00E291B1: lea ecx, var_38
  loc_00E291B4: push ecx
  loc_00E291B5: push eax
  loc_00E291B6: mov var_D0, eax
  loc_00E291BC: call [edx+00000054h]
  loc_00E291BF: test eax, eax
  loc_00E291C1: fnclex
  loc_00E291C3: jge 00E291D6h
  loc_00E291C5: mov edx, var_D0
  loc_00E291CB: push 00000054h
  loc_00E291CD: push 006DCBE8h
  loc_00E291D2: push edx
  loc_00E291D3: push eax
  loc_00E291D4: call ebx
  loc_00E291D6: lea edx, var_3C
  loc_00E291D9: mov eax, 00000002h
  loc_00E291DE: push edx
  loc_00E291DF: mov ecx, var_38
  loc_00E291E2: sub esp, 00000010h
  loc_00E291E5: mov var_90, eax
  loc_00E291EB: mov edx, esp
  loc_00E291ED: mov var_88, 00000007h
  loc_00E291F7: mov var_D8, ecx
  loc_00E291FD: mov ecx, [ecx]
  loc_00E291FF: mov [edx], eax
  loc_00E29201: mov eax, var_8C
  loc_00E29207: mov [edx+00000004h], eax
  loc_00E2920A: mov eax, var_88
  loc_00E29210: mov [edx+00000008h], eax
  loc_00E29213: mov eax, var_84
  loc_00E29219: mov [edx+0000000Ch], eax
  loc_00E2921C: mov edx, var_38
  loc_00E2921F: push edx
  loc_00E29220: call [ecx+00000028h]
  loc_00E29223: test eax, eax
  loc_00E29225: fnclex
  loc_00E29227: jge 00E2923Ah
  loc_00E29229: mov ecx, var_D8
  loc_00E2922F: push 00000028h
  loc_00E29231: push 006E09E8h
  loc_00E29236: push ecx
  loc_00E29237: push eax
  loc_00E29238: call ebx
  loc_00E2923A: mov ecx, var_3C
  loc_00E2923D: mov eax, 00000002h
  loc_00E29242: mov var_98, 00000000h
  loc_00E2924C: mov var_A0, eax
  loc_00E29252: mov edx, [ecx]
  loc_00E29254: sub esp, 00000010h
  loc_00E29257: mov var_E0, ecx
  loc_00E2925D: mov ecx, esp
  loc_00E2925F: mov [ecx], eax
  loc_00E29261: mov eax, var_9C
  loc_00E29267: mov [ecx+00000004h], eax
  loc_00E2926A: mov eax, var_98
  loc_00E29270: mov [ecx+00000008h], eax
  loc_00E29273: mov eax, var_94
  loc_00E29279: mov [ecx+0000000Ch], eax
  loc_00E2927C: mov ecx, var_3C
  loc_00E2927F: push ecx
  loc_00E29280: call [edx+00000038h]
  loc_00E29283: test eax, eax
  loc_00E29285: fnclex
  loc_00E29287: jge 00E29B6Eh
  loc_00E2928D: jmp 00E29B5Dh
  loc_00E29292: mov ecx, [esi]
  loc_00E29294: push esi
  loc_00E29295: call [ecx+000003B0h]
  loc_00E2929B: lea edx, var_30
  loc_00E2929E: push eax
  loc_00E2929F: push edx
  loc_00E292A0: call edi
  loc_00E292A2: mov ecx, [eax]
  loc_00E292A4: lea edx, var_28
  loc_00E292A7: push edx
  loc_00E292A8: push eax
  loc_00E292A9: mov var_D0, eax
  loc_00E292AF: call [ecx+00000050h]
  loc_00E292B2: test eax, eax
  loc_00E292B4: fnclex
  loc_00E292B6: jge 00E292C9h
  loc_00E292B8: mov ecx, var_D0
  loc_00E292BE: push 00000050h
  loc_00E292C0: push 006E0410h
  loc_00E292C5: push ecx
  loc_00E292C6: push eax
  loc_00E292C7: call ebx
  loc_00E292C9: mov edx, var_28
  loc_00E292CC: push edx
  loc_00E292CD: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E292D3: call [004010D8h] ; __vbaFpR8
  loc_00E292D9: fcomp real8 ptr [004022E8h]
  loc_00E292DF: mov var_188, 00000001h
  loc_00E292E9: fnstsw ax
  loc_00E292EB: test ah, 01h
  loc_00E292EE: jnz 00E292FAh
  loc_00E292F0: mov var_188, 00000000h
  loc_00E292FA: lea ecx, var_28
  loc_00E292FD: call [00401258h] ; __vbaFreeStr
  loc_00E29303: lea ecx, var_30
  loc_00E29306: call [00401254h] ; __vbaFreeObj
  loc_00E2930C: mov eax, var_188
  loc_00E29312: push 006DCBD8h
  loc_00E29317: neg eax
  loc_00E29319: push 00000000h
  loc_00E2931B: push 00000018h
  loc_00E2931D: test ax, ax
  loc_00E29320: mov eax, [esi]
  loc_00E29322: push esi
  loc_00E29323: jz 00E29596h
  loc_00E29329: call [eax+000004A8h]
  loc_00E2932F: lea ecx, var_30
  loc_00E29332: push eax
  loc_00E29333: push ecx
  loc_00E29334: call edi
  loc_00E29336: lea edx, var_50
  loc_00E29339: push eax
  loc_00E2933A: push edx
  loc_00E2933B: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E29341: add esp, 00000010h
  loc_00E29344: push eax
  loc_00E29345: call [00401120h] ; __vbaCastObjVar
  loc_00E2934B: push eax
  loc_00E2934C: lea eax, var_34
  loc_00E2934F: push eax
  loc_00E29350: call edi
  loc_00E29352: mov ecx, [eax]
  loc_00E29354: lea edx, var_38
  loc_00E29357: push edx
  loc_00E29358: push eax
  loc_00E29359: mov var_D0, eax
  loc_00E2935F: call [ecx+00000054h]
  loc_00E29362: test eax, eax
  loc_00E29364: fnclex
  loc_00E29366: jge 00E29379h
  loc_00E29368: mov ecx, var_D0
  loc_00E2936E: push 00000054h
  loc_00E29370: push 006DCBE8h
  loc_00E29375: push ecx
  loc_00E29376: push eax
  loc_00E29377: call ebx
  loc_00E29379: mov ecx, var_38
  loc_00E2937C: mov eax, 00000002h
  loc_00E29381: mov var_88, 0000000Ah
  loc_00E2938B: mov var_90, eax
  loc_00E29391: mov edx, [ecx]
  loc_00E29393: mov var_D8, ecx
  loc_00E29399: lea ecx, var_3C
  loc_00E2939C: push ecx
  loc_00E2939D: sub esp, 00000010h
  loc_00E293A0: mov ecx, esp
  loc_00E293A2: mov [ecx], eax
  loc_00E293A4: mov eax, var_8C
  loc_00E293AA: mov [ecx+00000004h], eax
  loc_00E293AD: mov eax, var_88
  loc_00E293B3: mov [ecx+00000008h], eax
  loc_00E293B6: mov eax, var_84
  loc_00E293BC: mov [ecx+0000000Ch], eax
  loc_00E293BF: mov ecx, var_38
  loc_00E293C2: push ecx
  loc_00E293C3: call [edx+00000028h]
  loc_00E293C6: test eax, eax
  loc_00E293C8: fnclex
  loc_00E293CA: jge 00E293DDh
  loc_00E293CC: mov edx, var_D8
  loc_00E293D2: push 00000028h
  loc_00E293D4: push 006E09E8h
  loc_00E293D9: push edx
  loc_00E293DA: push eax
  loc_00E293DB: call ebx
  loc_00E293DD: sub esp, 00000010h
  loc_00E293E0: mov eax, 00000002h
  loc_00E293E5: mov edx, esp
  loc_00E293E7: mov ecx, var_3C
  loc_00E293EA: mov var_A0, eax
  loc_00E293F0: mov var_98, 00000000h
  loc_00E293FA: mov [edx], eax
  loc_00E293FC: mov eax, var_9C
  loc_00E29402: mov var_E0, ecx
  loc_00E29408: mov ecx, [ecx]
  loc_00E2940A: mov [edx+00000004h], eax
  loc_00E2940D: mov eax, var_98
  loc_00E29413: mov [edx+00000008h], eax
  loc_00E29416: mov eax, var_94
  loc_00E2941C: mov [edx+0000000Ch], eax
  loc_00E2941F: mov edx, var_3C
  loc_00E29422: push edx
  loc_00E29423: call [ecx+00000038h]
  loc_00E29426: test eax, eax
  loc_00E29428: fnclex
  loc_00E2942A: jge 00E2943Dh
  loc_00E2942C: mov ecx, var_E0
  loc_00E29432: push 00000038h
  loc_00E29434: push 006E09F8h
  loc_00E29439: push ecx
  loc_00E2943A: push eax
  loc_00E2943B: call ebx
  loc_00E2943D: lea edx, var_3C
  loc_00E29440: lea eax, var_38
  loc_00E29443: push edx
  loc_00E29444: lea ecx, var_34
  loc_00E29447: push eax
  loc_00E29448: lea edx, var_30
  loc_00E2944B: push ecx
  loc_00E2944C: push edx
  loc_00E2944D: push 00000004h
  loc_00E2944F: call [00401048h] ; __vbaFreeObjList
  loc_00E29455: add esp, 00000014h
  loc_00E29458: lea ecx, var_50
  loc_00E2945B: call [00401024h] ; __vbaFreeVar
  loc_00E29461: mov eax, [esi]
  loc_00E29463: push 006DCBD8h
  loc_00E29468: push 00000000h
  loc_00E2946A: push 00000018h
  loc_00E2946C: push esi
  loc_00E2946D: call [eax+000004A8h]
  loc_00E29473: lea ecx, var_30
  loc_00E29476: push eax
  loc_00E29477: push ecx
  loc_00E29478: call edi
  loc_00E2947A: lea edx, var_50
  loc_00E2947D: push eax
  loc_00E2947E: push edx
  loc_00E2947F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E29485: add esp, 00000010h
  loc_00E29488: push eax
  loc_00E29489: call [00401120h] ; __vbaCastObjVar
  loc_00E2948F: push eax
  loc_00E29490: lea eax, var_34
  loc_00E29493: push eax
  loc_00E29494: call edi
  loc_00E29496: mov ecx, [eax]
  loc_00E29498: lea edx, var_38
  loc_00E2949B: push edx
  loc_00E2949C: push eax
  loc_00E2949D: mov var_D0, eax
  loc_00E294A3: call [ecx+00000054h]
  loc_00E294A6: test eax, eax
  loc_00E294A8: fnclex
  loc_00E294AA: jge 00E294BDh
  loc_00E294AC: mov ecx, var_D0
  loc_00E294B2: push 00000054h
  loc_00E294B4: push 006DCBE8h
  loc_00E294B9: push ecx
  loc_00E294BA: push eax
  loc_00E294BB: call ebx
  loc_00E294BD: mov ecx, var_38
  loc_00E294C0: mov eax, 00000002h
  loc_00E294C5: mov var_88, 00000007h
  loc_00E294CF: mov var_90, eax
  loc_00E294D5: mov edx, [ecx]
  loc_00E294D7: mov var_D8, ecx
  loc_00E294DD: lea ecx, var_3C
  loc_00E294E0: push ecx
  loc_00E294E1: sub esp, 00000010h
  loc_00E294E4: mov ecx, esp
  loc_00E294E6: mov [ecx], eax
  loc_00E294E8: mov eax, var_8C
  loc_00E294EE: mov [ecx+00000004h], eax
  loc_00E294F1: mov eax, var_88
  loc_00E294F7: mov [ecx+00000008h], eax
  loc_00E294FA: mov eax, var_84
  loc_00E29500: mov [ecx+0000000Ch], eax
  loc_00E29503: mov ecx, var_38
  loc_00E29506: push ecx
  loc_00E29507: call [edx+00000028h]
  loc_00E2950A: test eax, eax
  loc_00E2950C: fnclex
  loc_00E2950E: jge 00E29521h
  loc_00E29510: mov edx, var_D8
  loc_00E29516: push 00000028h
  loc_00E29518: push 006E09E8h
  loc_00E2951D: push edx
  loc_00E2951E: push eax
  loc_00E2951F: call ebx
  loc_00E29521: sub esp, 00000010h
  loc_00E29524: mov eax, 00000002h
  loc_00E29529: mov edx, esp
  loc_00E2952B: mov ecx, var_3C
  loc_00E2952E: mov var_A0, eax
  loc_00E29534: mov var_98, 00000000h
  loc_00E2953E: mov [edx], eax
  loc_00E29540: mov eax, var_9C
  loc_00E29546: mov var_E0, ecx
  loc_00E2954C: mov ecx, [ecx]
  loc_00E2954E: mov [edx+00000004h], eax
  loc_00E29551: mov eax, var_98
  loc_00E29557: mov [edx+00000008h], eax
  loc_00E2955A: mov eax, var_94
  loc_00E29560: mov [edx+0000000Ch], eax
  loc_00E29563: mov edx, var_3C
  loc_00E29566: push edx
  loc_00E29567: call [ecx+00000038h]
  loc_00E2956A: test eax, eax
  loc_00E2956C: fnclex
  loc_00E2956E: jge 00E29581h
  loc_00E29570: mov ecx, var_E0
  loc_00E29576: push 00000038h
  loc_00E29578: push 006E09F8h
  loc_00E2957D: push ecx
  loc_00E2957E: push eax
  loc_00E2957F: call ebx
  loc_00E29581: lea edx, var_3C
  loc_00E29584: lea eax, var_38
  loc_00E29587: push edx
  loc_00E29588: lea ecx, var_34
  loc_00E2958B: push eax
  loc_00E2958C: lea edx, var_30
  loc_00E2958F: push ecx
  loc_00E29590: push edx
  loc_00E29591: jmp 00E29B7Eh
  loc_00E29596: call [eax+000004A8h]
  loc_00E2959C: lea ecx, var_34
  loc_00E2959F: push eax
  loc_00E295A0: push ecx
  loc_00E295A1: call edi
  loc_00E295A3: lea edx, var_50
  loc_00E295A6: push eax
  loc_00E295A7: push edx
  loc_00E295A8: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E295AE: add esp, 00000010h
  loc_00E295B1: push eax
  loc_00E295B2: call [00401120h] ; __vbaCastObjVar
  loc_00E295B8: push eax
  loc_00E295B9: lea eax, var_38
  loc_00E295BC: push eax
  loc_00E295BD: call edi
  loc_00E295BF: mov ecx, [eax]
  loc_00E295C1: lea edx, var_3C
  loc_00E295C4: push edx
  loc_00E295C5: push eax
  loc_00E295C6: mov var_D8, eax
  loc_00E295CC: call [ecx+00000054h]
  loc_00E295CF: test eax, eax
  loc_00E295D1: fnclex
  loc_00E295D3: jge 00E295E6h
  loc_00E295D5: mov ecx, var_D8
  loc_00E295DB: push 00000054h
  loc_00E295DD: push 006DCBE8h
  loc_00E295E2: push ecx
  loc_00E295E3: push eax
  loc_00E295E4: call ebx
  loc_00E295E6: mov ecx, var_3C
  loc_00E295E9: mov eax, 00000002h
  loc_00E295EE: mov var_88, 0000000Ah
  loc_00E295F8: mov var_90, eax
  loc_00E295FE: mov edx, [ecx]
  loc_00E29600: mov var_E0, ecx
  loc_00E29606: lea ecx, var_40
  loc_00E29609: push ecx
  loc_00E2960A: sub esp, 00000010h
  loc_00E2960D: mov ecx, esp
  loc_00E2960F: mov [ecx], eax
  loc_00E29611: mov eax, var_8C
  loc_00E29617: mov [ecx+00000004h], eax
  loc_00E2961A: mov eax, var_88
  loc_00E29620: mov [ecx+00000008h], eax
  loc_00E29623: mov eax, var_84
  loc_00E29629: mov [ecx+0000000Ch], eax
  loc_00E2962C: mov ecx, var_3C
  loc_00E2962F: push ecx
  loc_00E29630: call [edx+00000028h]
  loc_00E29633: test eax, eax
  loc_00E29635: fnclex
  loc_00E29637: jge 00E2964Ah
  loc_00E29639: mov edx, var_E0
  loc_00E2963F: push 00000028h
  loc_00E29641: push 006E09E8h
  loc_00E29646: push edx
  loc_00E29647: push eax
  loc_00E29648: call ebx
  loc_00E2964A: mov eax, var_40
  loc_00E2964D: mov ecx, [esi]
  loc_00E2964F: push esi
  loc_00E29650: mov var_E8, eax
  loc_00E29656: call [ecx+000003B0h]
  loc_00E2965C: lea edx, var_30
  loc_00E2965F: push eax
  loc_00E29660: push edx
  loc_00E29661: call edi
  loc_00E29663: mov ecx, [eax]
  loc_00E29665: lea edx, var_28
  loc_00E29668: push edx
  loc_00E29669: push eax
  loc_00E2966A: mov var_D0, eax
  loc_00E29670: call [ecx+00000050h]
  loc_00E29673: test eax, eax
  loc_00E29675: fnclex
  loc_00E29677: jge 00E2968Ah
  loc_00E29679: mov ecx, var_D0
  loc_00E2967F: push 00000050h
  loc_00E29681: push 006E0410h
  loc_00E29686: push ecx
  loc_00E29687: push eax
  loc_00E29688: call ebx
  loc_00E2968A: mov eax, var_28
  loc_00E2968D: mov edx, var_E8
  loc_00E29693: mov var_58, eax
  loc_00E29696: mov eax, 00000008h
  loc_00E2969B: mov var_28, 00000000h
  loc_00E296A2: mov var_60, eax
  loc_00E296A5: mov ecx, [edx]
  loc_00E296A7: sub esp, 00000010h
  loc_00E296AA: mov edx, esp
  loc_00E296AC: mov [edx], eax
  loc_00E296AE: mov eax, var_5C
  loc_00E296B1: mov [edx+00000004h], eax
  loc_00E296B4: mov eax, var_58
  loc_00E296B7: mov [edx+00000008h], eax
  loc_00E296BA: mov eax, var_54
  loc_00E296BD: mov [edx+0000000Ch], eax
  loc_00E296C0: mov edx, var_E8
  loc_00E296C6: push edx
  loc_00E296C7: call [ecx+00000038h]
  loc_00E296CA: test eax, eax
  loc_00E296CC: fnclex
  loc_00E296CE: jge 00E296E1h
  loc_00E296D0: mov ecx, var_E8
  loc_00E296D6: push 00000038h
  loc_00E296D8: push 006E09F8h
  loc_00E296DD: push ecx
  loc_00E296DE: push eax
  loc_00E296DF: call ebx
  loc_00E296E1: lea edx, var_40
  loc_00E296E4: lea eax, var_3C
  loc_00E296E7: push edx
  loc_00E296E8: lea ecx, var_38
  loc_00E296EB: push eax
  loc_00E296EC: lea edx, var_34
  loc_00E296EF: push ecx
  loc_00E296F0: lea eax, var_30
  loc_00E296F3: push edx
  loc_00E296F4: push eax
  loc_00E296F5: push 00000005h
  loc_00E296F7: call [00401048h] ; __vbaFreeObjList
  loc_00E296FD: lea ecx, var_60
  loc_00E29700: lea edx, var_50
  loc_00E29703: push ecx
  loc_00E29704: push edx
  loc_00E29705: push 00000002h
  loc_00E29707: call [00401038h] ; __vbaFreeVarList
  loc_00E2970D: mov eax, [esi]
  loc_00E2970F: add esp, 00000024h
  loc_00E29712: push 006DCBD8h
  loc_00E29717: push 00000000h
  loc_00E29719: push 00000018h
  loc_00E2971B: push esi
  loc_00E2971C: call [eax+000004A8h]
  loc_00E29722: lea ecx, var_34
  loc_00E29725: push eax
  loc_00E29726: push ecx
  loc_00E29727: call edi
  loc_00E29729: lea edx, var_50
  loc_00E2972C: push eax
  loc_00E2972D: push edx
  loc_00E2972E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E29734: add esp, 00000010h
  loc_00E29737: push eax
  loc_00E29738: call [00401120h] ; __vbaCastObjVar
  loc_00E2973E: push eax
  loc_00E2973F: lea eax, var_38
  loc_00E29742: push eax
  loc_00E29743: call edi
  loc_00E29745: mov ecx, [eax]
  loc_00E29747: lea edx, var_3C
  loc_00E2974A: push edx
  loc_00E2974B: push eax
  loc_00E2974C: mov var_D8, eax
  loc_00E29752: call [ecx+00000054h]
  loc_00E29755: test eax, eax
  loc_00E29757: fnclex
  loc_00E29759: jge 00E2976Ch
  loc_00E2975B: mov ecx, var_D8
  loc_00E29761: push 00000054h
  loc_00E29763: push 006DCBE8h
  loc_00E29768: push ecx
  loc_00E29769: push eax
  loc_00E2976A: call ebx
  loc_00E2976C: mov ecx, var_3C
  loc_00E2976F: mov eax, 00000002h
  loc_00E29774: mov var_88, 00000007h
  loc_00E2977E: mov var_90, eax
  loc_00E29784: mov edx, [ecx]
  loc_00E29786: mov var_E0, ecx
  loc_00E2978C: lea ecx, var_40
  loc_00E2978F: push ecx
  loc_00E29790: sub esp, 00000010h
  loc_00E29793: mov ecx, esp
  loc_00E29795: mov [ecx], eax
  loc_00E29797: mov eax, var_8C
  loc_00E2979D: mov [ecx+00000004h], eax
  loc_00E297A0: mov eax, var_88
  loc_00E297A6: mov [ecx+00000008h], eax
  loc_00E297A9: mov eax, var_84
  loc_00E297AF: mov [ecx+0000000Ch], eax
  loc_00E297B2: mov ecx, var_3C
  loc_00E297B5: push ecx
  loc_00E297B6: call [edx+00000028h]
  loc_00E297B9: test eax, eax
  loc_00E297BB: fnclex
  loc_00E297BD: jge 00E297D0h
  loc_00E297BF: mov edx, var_E0
  loc_00E297C5: push 00000028h
  loc_00E297C7: push 006E09E8h
  loc_00E297CC: push edx
  loc_00E297CD: push eax
  loc_00E297CE: call ebx
  loc_00E297D0: mov eax, var_40
  loc_00E297D3: mov ecx, [esi]
  loc_00E297D5: push esi
  loc_00E297D6: mov var_E8, eax
  loc_00E297DC: call [ecx+0000039Ch]
  loc_00E297E2: lea edx, var_30
  loc_00E297E5: push eax
  loc_00E297E6: push edx
  loc_00E297E7: call edi
  loc_00E297E9: mov ecx, [eax]
  loc_00E297EB: lea edx, var_28
  loc_00E297EE: push edx
  loc_00E297EF: push eax
  loc_00E297F0: mov var_D0, eax
  loc_00E297F6: call [ecx+000000A0h]
  loc_00E297FC: test eax, eax
  loc_00E297FE: fnclex
  loc_00E29800: jge 00E29816h
  loc_00E29802: mov ecx, var_D0
  loc_00E29808: push 000000A0h
  loc_00E2980D: push 006DCB70h
  loc_00E29812: push ecx
  loc_00E29813: push eax
  loc_00E29814: call ebx
  loc_00E29816: mov eax, var_28
  loc_00E29819: mov edx, var_E8
  loc_00E2981F: mov var_58, eax
  loc_00E29822: mov eax, 00000008h
  loc_00E29827: mov var_28, 00000000h
  loc_00E2982E: mov var_60, eax
  loc_00E29831: mov ecx, [edx]
  loc_00E29833: sub esp, 00000010h
  loc_00E29836: mov edx, esp
  loc_00E29838: mov [edx], eax
  loc_00E2983A: mov eax, var_5C
  loc_00E2983D: mov [edx+00000004h], eax
  loc_00E29840: mov eax, var_58
  loc_00E29843: mov [edx+00000008h], eax
  loc_00E29846: mov eax, var_54
  loc_00E29849: mov [edx+0000000Ch], eax
  loc_00E2984C: mov edx, var_E8
  loc_00E29852: push edx
  loc_00E29853: call [ecx+00000038h]
  loc_00E29856: test eax, eax
  loc_00E29858: fnclex
  loc_00E2985A: jge 00E2986Dh
  loc_00E2985C: mov ecx, var_E8
  loc_00E29862: push 00000038h
  loc_00E29864: push 006E09F8h
  loc_00E29869: push ecx
  loc_00E2986A: push eax
  loc_00E2986B: call ebx
  loc_00E2986D: lea edx, var_40
  loc_00E29870: lea eax, var_3C
  loc_00E29873: push edx
  loc_00E29874: lea ecx, var_38
  loc_00E29877: push eax
  loc_00E29878: lea edx, var_34
  loc_00E2987B: push ecx
  loc_00E2987C: lea eax, var_30
  loc_00E2987F: push edx
  loc_00E29880: push eax
  loc_00E29881: push 00000005h
  loc_00E29883: call [00401048h] ; __vbaFreeObjList
  loc_00E29889: lea ecx, var_60
  loc_00E2988C: lea edx, var_50
  loc_00E2988F: push ecx
  loc_00E29890: push edx
  loc_00E29891: push 00000002h
  loc_00E29893: call [00401038h] ; __vbaFreeVarList
  loc_00E29899: add esp, 00000024h
  loc_00E2989C: jmp 00E29B92h
  loc_00E298A1: call [eax+00000398h]
  loc_00E298A7: lea ecx, var_30
  loc_00E298AA: push eax
  loc_00E298AB: push ecx
  loc_00E298AC: call edi
  loc_00E298AE: mov edx, [eax]
  loc_00E298B0: lea ecx, var_C4
  loc_00E298B6: push ecx
  loc_00E298B7: push eax
  loc_00E298B8: mov var_D0, eax
  loc_00E298BE: call [edx+00000098h]
  loc_00E298C4: test eax, eax
  loc_00E298C6: fnclex
  loc_00E298C8: jge 00E298DEh
  loc_00E298CA: mov edx, var_D0
  loc_00E298D0: push 00000098h
  loc_00E298D5: push 006DCAD0h
  loc_00E298DA: push edx
  loc_00E298DB: push eax
  loc_00E298DC: call ebx
  loc_00E298DE: xor eax, eax
  loc_00E298E0: lea ecx, var_30
  loc_00E298E3: cmp var_C4, ax
  loc_00E298EA: setz al
  loc_00E298ED: neg eax
  loc_00E298EF: mov var_D8, ax
  loc_00E298F6: call [00401254h] ; __vbaFreeObj
  loc_00E298FC: cmp var_D8, 0000h
  loc_00E29904: jz 00E29B92h
  loc_00E2990A: mov ecx, [esi]
  loc_00E2990C: push 006DCBD8h
  loc_00E29911: push 00000000h
  loc_00E29913: push 00000018h
  loc_00E29915: push esi
  loc_00E29916: call [ecx+000004A8h]
  loc_00E2991C: lea edx, var_30
  loc_00E2991F: push eax
  loc_00E29920: push edx
  loc_00E29921: call edi
  loc_00E29923: push eax
  loc_00E29924: lea eax, var_50
  loc_00E29927: push eax
  loc_00E29928: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2992E: add esp, 00000010h
  loc_00E29931: push eax
  loc_00E29932: call [00401120h] ; __vbaCastObjVar
  loc_00E29938: lea ecx, var_34
  loc_00E2993B: push eax
  loc_00E2993C: push ecx
  loc_00E2993D: call edi
  loc_00E2993F: mov edx, [eax]
  loc_00E29941: lea ecx, var_38
  loc_00E29944: push ecx
  loc_00E29945: push eax
  loc_00E29946: mov var_D0, eax
  loc_00E2994C: call [edx+00000054h]
  loc_00E2994F: test eax, eax
  loc_00E29951: fnclex
  loc_00E29953: jge 00E29966h
  loc_00E29955: mov edx, var_D0
  loc_00E2995B: push 00000054h
  loc_00E2995D: push 006DCBE8h
  loc_00E29962: push edx
  loc_00E29963: push eax
  loc_00E29964: call ebx
  loc_00E29966: lea edx, var_3C
  loc_00E29969: mov eax, 00000002h
  loc_00E2996E: push edx
  loc_00E2996F: mov ecx, var_38
  loc_00E29972: sub esp, 00000010h
  loc_00E29975: mov var_90, eax
  loc_00E2997B: mov edx, esp
  loc_00E2997D: mov var_88, 0000000Ah
  loc_00E29987: mov var_D8, ecx
  loc_00E2998D: mov ecx, [ecx]
  loc_00E2998F: mov [edx], eax
  loc_00E29991: mov eax, var_8C
  loc_00E29997: mov [edx+00000004h], eax
  loc_00E2999A: mov eax, var_88
  loc_00E299A0: mov [edx+00000008h], eax
  loc_00E299A3: mov eax, var_84
  loc_00E299A9: mov [edx+0000000Ch], eax
  loc_00E299AC: mov edx, var_38
  loc_00E299AF: push edx
  loc_00E299B0: call [ecx+00000028h]
  loc_00E299B3: test eax, eax
  loc_00E299B5: fnclex
  loc_00E299B7: jge 00E299CAh
  loc_00E299B9: mov ecx, var_D8
  loc_00E299BF: push 00000028h
  loc_00E299C1: push 006E09E8h
  loc_00E299C6: push ecx
  loc_00E299C7: push eax
  loc_00E299C8: call ebx
  loc_00E299CA: mov ecx, var_3C
  loc_00E299CD: mov eax, 00000002h
  loc_00E299D2: mov var_98, 00000000h
  loc_00E299DC: mov var_A0, eax
  loc_00E299E2: mov edx, [ecx]
  loc_00E299E4: sub esp, 00000010h
  loc_00E299E7: mov var_E0, ecx
  loc_00E299ED: mov ecx, esp
  loc_00E299EF: mov [ecx], eax
  loc_00E299F1: mov eax, var_9C
  loc_00E299F7: mov [ecx+00000004h], eax
  loc_00E299FA: mov eax, var_98
  loc_00E29A00: mov [ecx+00000008h], eax
  loc_00E29A03: mov eax, var_94
  loc_00E29A09: mov [ecx+0000000Ch], eax
  loc_00E29A0C: mov ecx, var_3C
  loc_00E29A0F: push ecx
  loc_00E29A10: call [edx+00000038h]
  loc_00E29A13: test eax, eax
  loc_00E29A15: fnclex
  loc_00E29A17: jge 00E29A2Ah
  loc_00E29A19: mov edx, var_E0
  loc_00E29A1F: push 00000038h
  loc_00E29A21: push 006E09F8h
  loc_00E29A26: push edx
  loc_00E29A27: push eax
  loc_00E29A28: call ebx
  loc_00E29A2A: lea eax, var_3C
  loc_00E29A2D: lea ecx, var_38
  loc_00E29A30: push eax
  loc_00E29A31: lea edx, var_34
  loc_00E29A34: push ecx
  loc_00E29A35: lea eax, var_30
  loc_00E29A38: push edx
  loc_00E29A39: push eax
  loc_00E29A3A: push 00000004h
  loc_00E29A3C: call [00401048h] ; __vbaFreeObjList
  loc_00E29A42: add esp, 00000014h
  loc_00E29A45: lea ecx, var_50
  loc_00E29A48: call [00401024h] ; __vbaFreeVar
  loc_00E29A4E: mov ecx, [esi]
  loc_00E29A50: push 006DCBD8h
  loc_00E29A55: push 00000000h
  loc_00E29A57: push 00000018h
  loc_00E29A59: push esi
  loc_00E29A5A: call [ecx+000004A8h]
  loc_00E29A60: lea edx, var_30
  loc_00E29A63: push eax
  loc_00E29A64: push edx
  loc_00E29A65: call edi
  loc_00E29A67: push eax
  loc_00E29A68: lea eax, var_50
  loc_00E29A6B: push eax
  loc_00E29A6C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E29A72: add esp, 00000010h
  loc_00E29A75: push eax
  loc_00E29A76: call [00401120h] ; __vbaCastObjVar
  loc_00E29A7C: lea ecx, var_34
  loc_00E29A7F: push eax
  loc_00E29A80: push ecx
  loc_00E29A81: call edi
  loc_00E29A83: mov edx, [eax]
  loc_00E29A85: lea ecx, var_38
  loc_00E29A88: push ecx
  loc_00E29A89: push eax
  loc_00E29A8A: mov var_D0, eax
  loc_00E29A90: call [edx+00000054h]
  loc_00E29A93: test eax, eax
  loc_00E29A95: fnclex
  loc_00E29A97: jge 00E29AAAh
  loc_00E29A99: mov edx, var_D0
  loc_00E29A9F: push 00000054h
  loc_00E29AA1: push 006DCBE8h
  loc_00E29AA6: push edx
  loc_00E29AA7: push eax
  loc_00E29AA8: call ebx
  loc_00E29AAA: lea edx, var_3C
  loc_00E29AAD: mov eax, 00000002h
  loc_00E29AB2: push edx
  loc_00E29AB3: mov ecx, var_38
  loc_00E29AB6: sub esp, 00000010h
  loc_00E29AB9: mov var_90, eax
  loc_00E29ABF: mov edx, esp
  loc_00E29AC1: mov var_88, 00000007h
  loc_00E29ACB: mov var_D8, ecx
  loc_00E29AD1: mov ecx, [ecx]
  loc_00E29AD3: mov [edx], eax
  loc_00E29AD5: mov eax, var_8C
  loc_00E29ADB: mov [edx+00000004h], eax
  loc_00E29ADE: mov eax, var_88
  loc_00E29AE4: mov [edx+00000008h], eax
  loc_00E29AE7: mov eax, var_84
  loc_00E29AED: mov [edx+0000000Ch], eax
  loc_00E29AF0: mov edx, var_38
  loc_00E29AF3: push edx
  loc_00E29AF4: call [ecx+00000028h]
  loc_00E29AF7: test eax, eax
  loc_00E29AF9: fnclex
  loc_00E29AFB: jge 00E29B0Eh
  loc_00E29AFD: mov ecx, var_D8
  loc_00E29B03: push 00000028h
  loc_00E29B05: push 006E09E8h
  loc_00E29B0A: push ecx
  loc_00E29B0B: push eax
  loc_00E29B0C: call ebx
  loc_00E29B0E: mov ecx, var_3C
  loc_00E29B11: mov eax, 00000002h
  loc_00E29B16: mov var_98, 00000000h
  loc_00E29B20: mov var_A0, eax
  loc_00E29B26: mov edx, [ecx]
  loc_00E29B28: sub esp, 00000010h
  loc_00E29B2B: mov var_E0, ecx
  loc_00E29B31: mov ecx, esp
  loc_00E29B33: mov [ecx], eax
  loc_00E29B35: mov eax, var_9C
  loc_00E29B3B: mov [ecx+00000004h], eax
  loc_00E29B3E: mov eax, var_98
  loc_00E29B44: mov [ecx+00000008h], eax
  loc_00E29B47: mov eax, var_94
  loc_00E29B4D: mov [ecx+0000000Ch], eax
  loc_00E29B50: mov ecx, var_3C
  loc_00E29B53: push ecx
  loc_00E29B54: call [edx+00000038h]
  loc_00E29B57: test eax, eax
  loc_00E29B59: fnclex
  loc_00E29B5B: jge 00E29B6Eh
  loc_00E29B5D: mov edx, var_E0
  loc_00E29B63: push 00000038h
  loc_00E29B65: push 006E09F8h
  loc_00E29B6A: push edx
  loc_00E29B6B: push eax
  loc_00E29B6C: call ebx
  loc_00E29B6E: lea eax, var_3C
  loc_00E29B71: lea ecx, var_38
  loc_00E29B74: push eax
  loc_00E29B75: lea edx, var_34
  loc_00E29B78: push ecx
  loc_00E29B79: lea eax, var_30
  loc_00E29B7C: push edx
  loc_00E29B7D: push eax
  loc_00E29B7E: push 00000004h
  loc_00E29B80: call [00401048h] ; __vbaFreeObjList
  loc_00E29B86: add esp, 00000014h
  loc_00E29B89: lea ecx, var_50
  loc_00E29B8C: call [00401024h] ; __vbaFreeVar
  loc_00E29B92: mov ecx, [esi]
  loc_00E29B94: push 006DCBD8h
  loc_00E29B99: push 00000000h
  loc_00E29B9B: push 00000018h
  loc_00E29B9D: push esi
  loc_00E29B9E: call [ecx+000004A8h]
  loc_00E29BA4: lea edx, var_30
  loc_00E29BA7: push eax
  loc_00E29BA8: push edx
  loc_00E29BA9: call edi
  loc_00E29BAB: push eax
  loc_00E29BAC: lea eax, var_50
  loc_00E29BAF: push eax
  loc_00E29BB0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E29BB6: add esp, 00000010h
  loc_00E29BB9: push eax
  loc_00E29BBA: call [00401120h] ; __vbaCastObjVar
  loc_00E29BC0: lea ecx, var_34
  loc_00E29BC3: push eax
  loc_00E29BC4: push ecx
  loc_00E29BC5: call edi
  loc_00E29BC7: mov ecx, 0000000Ah
  loc_00E29BCC: mov var_98, 80020004h
  loc_00E29BD6: mov var_A0, ecx
  loc_00E29BDC: mov var_88, 80020004h
  loc_00E29BE6: mov var_90, ecx
  loc_00E29BEC: mov edx, [eax]
  loc_00E29BEE: sub esp, 00000010h
  loc_00E29BF1: mov var_D0, eax
  loc_00E29BF7: mov eax, esp
  loc_00E29BF9: sub esp, 00000010h
  loc_00E29BFC: mov [eax], ecx
  loc_00E29BFE: mov ecx, var_9C
  loc_00E29C04: mov [eax+00000004h], ecx
  loc_00E29C07: mov ecx, var_98
  loc_00E29C0D: mov [eax+00000008h], ecx
  loc_00E29C10: mov ecx, var_94
  loc_00E29C16: mov [eax+0000000Ch], ecx
  loc_00E29C19: mov ecx, var_90
  loc_00E29C1F: mov eax, esp
  loc_00E29C21: mov [eax], ecx
  loc_00E29C23: mov ecx, var_8C
  loc_00E29C29: mov [eax+00000004h], ecx
  loc_00E29C2C: mov ecx, var_88
  loc_00E29C32: mov [eax+00000008h], ecx
  loc_00E29C35: mov ecx, var_84
  loc_00E29C3B: mov [eax+0000000Ch], ecx
  loc_00E29C3E: mov eax, var_D0
  loc_00E29C44: push eax
  loc_00E29C45: call [edx+000000ACh]
  loc_00E29C4B: test eax, eax
  loc_00E29C4D: fnclex
  loc_00E29C4F: jge 00E29C65h
  loc_00E29C51: mov ecx, var_D0
  loc_00E29C57: push 000000ACh
  loc_00E29C5C: push 006DCBE8h
  loc_00E29C61: push ecx
  loc_00E29C62: push eax
  loc_00E29C63: call ebx
  loc_00E29C65: lea edx, var_34
  loc_00E29C68: lea eax, var_30
  loc_00E29C6B: push edx
  loc_00E29C6C: push eax
  loc_00E29C6D: push 00000002h
  loc_00E29C6F: call [00401048h] ; __vbaFreeObjList
  loc_00E29C75: add esp, 0000000Ch
  loc_00E29C78: lea ecx, var_50
  loc_00E29C7B: call [00401024h] ; __vbaFreeVar
  loc_00E29C81: mov ecx, [esi]
  loc_00E29C83: push 00000000h
  loc_00E29C85: push 00000019h
  loc_00E29C87: push esi
  loc_00E29C88: call [ecx+000004A8h]
  loc_00E29C8E: lea edx, var_30
  loc_00E29C91: push eax
  loc_00E29C92: push edx
  loc_00E29C93: call edi
  loc_00E29C95: push eax
  loc_00E29C96: call [00401030h] ; __vbaLateIdCall
  loc_00E29C9C: add esp, 0000000Ch
  loc_00E29C9F: lea ecx, var_30
  loc_00E29CA2: call [00401254h] ; __vbaFreeObj
  loc_00E29CA8: jmp 00E2C7E5h
  loc_00E29CAD: mov eax, [esi]
  loc_00E29CAF: push esi
  loc_00E29CB0: call [eax+00000398h]
  loc_00E29CB6: lea ecx, var_30
  loc_00E29CB9: push eax
  loc_00E29CBA: push ecx
  loc_00E29CBB: call edi
  loc_00E29CBD: mov edx, [eax]
  loc_00E29CBF: lea ecx, var_C4
  loc_00E29CC5: push ecx
  loc_00E29CC6: push eax
  loc_00E29CC7: mov var_D0, eax
  loc_00E29CCD: call [edx+00000098h]
  loc_00E29CD3: test eax, eax
  loc_00E29CD5: fnclex
  loc_00E29CD7: jge 00E29CEDh
  loc_00E29CD9: mov edx, var_D0
  loc_00E29CDF: push 00000098h
  loc_00E29CE4: push 006DCAD0h
  loc_00E29CE9: push edx
  loc_00E29CEA: push eax
  loc_00E29CEB: call ebx
  loc_00E29CED: xor eax, eax
  loc_00E29CEF: cmp var_C4, FFFFFFh
  loc_00E29CF7: lea ecx, var_30
  loc_00E29CFA: setz al
  loc_00E29CFD: neg eax
  loc_00E29CFF: mov var_D8, ax
  loc_00E29D06: call [00401254h] ; __vbaFreeObj
  loc_00E29D0C: cmp var_D8, 0000h
  loc_00E29D14: jz 00E2C7E5h
  loc_00E29D1A: mov ecx, [esi]
  loc_00E29D1C: push esi
  loc_00E29D1D: call [ecx+0000039Ch]
  loc_00E29D23: lea edx, var_30
  loc_00E29D26: push eax
  loc_00E29D27: push edx
  loc_00E29D28: call edi
  loc_00E29D2A: mov ecx, [eax]
  loc_00E29D2C: lea edx, var_28
  loc_00E29D2F: push edx
  loc_00E29D30: push eax
  loc_00E29D31: mov var_D0, eax
  loc_00E29D37: call [ecx+000000A0h]
  loc_00E29D3D: test eax, eax
  loc_00E29D3F: fnclex
  loc_00E29D41: jge 00E29D57h
  loc_00E29D43: mov ecx, var_D0
  loc_00E29D49: push 000000A0h
  loc_00E29D4E: push 006DCB70h
  loc_00E29D53: push ecx
  loc_00E29D54: push eax
  loc_00E29D55: call ebx
  loc_00E29D57: mov edx, var_28
  loc_00E29D5A: push edx
  loc_00E29D5B: push 006E1138h ; "-"
  loc_00E29D60: call [00401104h] ; __vbaStrCmp
  loc_00E29D66: neg eax
  loc_00E29D68: sbb eax, eax
  loc_00E29D6A: lea ecx, var_28
  loc_00E29D6D: inc eax
  loc_00E29D6E: neg eax
  loc_00E29D70: mov var_D8, ax
  loc_00E29D77: call [00401258h] ; __vbaFreeStr
  loc_00E29D7D: lea ecx, var_30
  loc_00E29D80: call [00401254h] ; __vbaFreeObj
  loc_00E29D86: cmp var_D8, 0000h
  loc_00E29D8E: jz 00E29E4Dh
  loc_00E29D94: mov ecx, 80020004h
  loc_00E29D99: mov eax, 0000000Ah
  loc_00E29D9E: mov var_78, ecx
  loc_00E29DA1: mov var_68, ecx
  loc_00E29DA4: lea edx, var_A0
  loc_00E29DAA: lea ecx, var_60
  loc_00E29DAD: mov var_80, eax
  loc_00E29DB0: mov var_70, eax
  loc_00E29DB3: mov var_98, 006E0B08h ; "Need To Do"
  loc_00E29DBD: mov var_A0, 00000008h
  loc_00E29DC7: call [004011F0h] ; __vbaVarDup
  loc_00E29DCD: lea edx, var_90
  loc_00E29DD3: lea ecx, var_50
  loc_00E29DD6: mov var_88, 006E17FCh ; "Silahkan Tentukan Cicilan Keberapa ?"
  loc_00E29DE0: mov var_90, 00000008h
  loc_00E29DEA: call [004011F0h] ; __vbaVarDup
  loc_00E29DF0: lea eax, var_80
  loc_00E29DF3: lea ecx, var_70
  loc_00E29DF6: push eax
  loc_00E29DF7: lea edx, var_60
  loc_00E29DFA: push ecx
  loc_00E29DFB: push edx
  loc_00E29DFC: lea eax, var_50
  loc_00E29DFF: push 00000040h
  loc_00E29E01: push eax
  loc_00E29E02: call [004010A8h] ; rtcMsgBox
  loc_00E29E08: lea ecx, var_80
  loc_00E29E0B: lea edx, var_70
  loc_00E29E0E: push ecx
  loc_00E29E0F: lea eax, var_60
  loc_00E29E12: push edx
  loc_00E29E13: lea ecx, var_50
  loc_00E29E16: push eax
  loc_00E29E17: push ecx
  loc_00E29E18: push 00000004h
  loc_00E29E1A: call [00401038h] ; __vbaFreeVarList
  loc_00E29E20: mov edx, [esi]
  loc_00E29E22: add esp, 00000014h
  loc_00E29E25: push esi
  loc_00E29E26: call [edx+0000039Ch]
  loc_00E29E2C: push eax
  loc_00E29E2D: lea eax, var_30
  loc_00E29E30: push eax
  loc_00E29E31: call edi
  loc_00E29E33: mov esi, eax
  loc_00E29E35: push esi
  loc_00E29E36: mov ecx, [esi]
  loc_00E29E38: call [ecx+00000204h]
  loc_00E29E3E: test eax, eax
  loc_00E29E40: fnclex
  loc_00E29E42: jge 00E270E3h
  loc_00E29E48: jmp 00E270D5h
  loc_00E29E4D: mov edx, [esi]
  loc_00E29E4F: push 00000000h
  loc_00E29E51: push 80010007h
  loc_00E29E56: push esi
  loc_00E29E57: call [edx+00000458h]
  loc_00E29E5D: push eax
  loc_00E29E5E: lea eax, var_30
  loc_00E29E61: push eax
  loc_00E29E62: call edi
  loc_00E29E64: lea ecx, var_50
  loc_00E29E67: push eax
  loc_00E29E68: push ecx
  loc_00E29E69: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E29E6F: add esp, 00000010h
  loc_00E29E72: push eax
  loc_00E29E73: call [004010CCh] ; __vbaBoolVar
  loc_00E29E79: xor edx, edx
  loc_00E29E7B: cmp ax, FFFFFFh
  loc_00E29E7F: setz dl
  loc_00E29E82: neg edx
  loc_00E29E84: lea ecx, var_30
  loc_00E29E87: mov var_D0, dx
  loc_00E29E8E: call [00401254h] ; __vbaFreeObj
  loc_00E29E94: lea ecx, var_50
  loc_00E29E97: call [00401024h] ; __vbaFreeVar
  loc_00E29E9D: cmp var_D0, 0000h
  loc_00E29EA5: jz 00E29EF2h
  loc_00E29EA7: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E29EAD: mov ecx, 80020004h
  loc_00E29EB2: mov var_78, ecx
  loc_00E29EB5: mov eax, 0000000Ah
  loc_00E29EBA: mov var_68, ecx
  loc_00E29EBD: mov edi, 00000008h
  loc_00E29EC2: lea edx, var_A0
  loc_00E29EC8: lea ecx, var_60
  loc_00E29ECB: mov var_80, eax
  loc_00E29ECE: mov var_70, eax
  loc_00E29ED1: mov var_98, 006E0B08h ; "Need To Do"
  loc_00E29EDB: mov var_A0, edi
  loc_00E29EE1: call __vbaVarDup
  loc_00E29EE3: mov var_88, 006E184Ch ; "Silahkan Tentukan tanggal transaksi !"
  loc_00E29EED: jmp 00E2C7A1h
  loc_00E29EF2: mov edx, [esi]
  loc_00E29EF4: push 006DCBD8h
  loc_00E29EF9: push 00000000h
  loc_00E29EFB: push 00000018h
  loc_00E29EFD: push esi
  loc_00E29EFE: call [edx+000004A8h]
  loc_00E29F04: push eax
  loc_00E29F05: lea eax, var_30
  loc_00E29F08: push eax
  loc_00E29F09: call edi
  loc_00E29F0B: lea ecx, var_50
  loc_00E29F0E: push eax
  loc_00E29F0F: push ecx
  loc_00E29F10: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E29F16: add esp, 00000010h
  loc_00E29F19: push eax
  loc_00E29F1A: call [00401120h] ; __vbaCastObjVar
  loc_00E29F20: lea edx, var_34
  loc_00E29F23: push eax
  loc_00E29F24: push edx
  loc_00E29F25: call edi
  loc_00E29F27: mov ecx, [eax]
  loc_00E29F29: lea edx, var_C4
  loc_00E29F2F: push edx
  loc_00E29F30: push eax
  loc_00E29F31: mov var_D0, eax
  loc_00E29F37: call [ecx+00000050h]
  loc_00E29F3A: test eax, eax
  loc_00E29F3C: fnclex
  loc_00E29F3E: jge 00E29F51h
  loc_00E29F40: mov ecx, var_D0
  loc_00E29F46: push 00000050h
  loc_00E29F48: push 006DCBE8h
  loc_00E29F4D: push ecx
  loc_00E29F4E: push eax
  loc_00E29F4F: call ebx
  loc_00E29F51: mov dx, var_C4
  loc_00E29F58: lea eax, var_34
  loc_00E29F5B: lea ecx, var_30
  loc_00E29F5E: push eax
  loc_00E29F5F: push ecx
  loc_00E29F60: push 00000002h
  loc_00E29F62: mov var_D8, dx
  loc_00E29F69: call [00401048h] ; __vbaFreeObjList
  loc_00E29F6F: add esp, 0000000Ch
  loc_00E29F72: lea ecx, var_50
  loc_00E29F75: call [00401024h] ; __vbaFreeVar
  loc_00E29F7B: cmp var_D8, 0000h
  loc_00E29F83: jz 00E2C75Bh
  loc_00E29F89: mov edx, [esi]
  loc_00E29F8B: push 006DCBD8h
  loc_00E29F90: push 00000000h
  loc_00E29F92: push 00000018h
  loc_00E29F94: push esi
  loc_00E29F95: call [edx+000004A8h]
  loc_00E29F9B: push eax
  loc_00E29F9C: lea eax, var_30
  loc_00E29F9F: push eax
  loc_00E29FA0: call edi
  loc_00E29FA2: lea ecx, var_50
  loc_00E29FA5: push eax
  loc_00E29FA6: push ecx
  loc_00E29FA7: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E29FAD: add esp, 00000010h
  loc_00E29FB0: push eax
  loc_00E29FB1: call [00401120h] ; __vbaCastObjVar
  loc_00E29FB7: lea edx, var_34
  loc_00E29FBA: push eax
  loc_00E29FBB: push edx
  loc_00E29FBC: call edi
  loc_00E29FBE: sub esp, 00000010h
  loc_00E29FC1: mov ecx, 0000000Ah
  loc_00E29FC6: mov edx, esp
  loc_00E29FC8: mov var_A0, ecx
  loc_00E29FCE: mov var_90, ecx
  loc_00E29FD4: mov var_98, 80020004h
  loc_00E29FDE: mov [edx], ecx
  loc_00E29FE0: mov ecx, var_9C
  loc_00E29FE6: sub esp, 00000010h
  loc_00E29FE9: mov var_88, 80020004h
  loc_00E29FF3: mov [edx+00000004h], ecx
  loc_00E29FF6: mov ecx, var_98
  loc_00E29FFC: mov var_D0, eax
  loc_00E2A002: mov eax, [eax]
  loc_00E2A004: mov [edx+00000008h], ecx
  loc_00E2A007: mov ecx, var_94
  loc_00E2A00D: mov [edx+0000000Ch], ecx
  loc_00E2A010: mov ecx, var_90
  loc_00E2A016: mov edx, esp
  loc_00E2A018: mov [edx], ecx
  loc_00E2A01A: mov ecx, var_8C
  loc_00E2A020: mov [edx+00000004h], ecx
  loc_00E2A023: mov ecx, var_88
  loc_00E2A029: mov [edx+00000008h], ecx
  loc_00E2A02C: mov ecx, var_84
  loc_00E2A032: mov [edx+0000000Ch], ecx
  loc_00E2A035: mov edx, var_D0
  loc_00E2A03B: push edx
  loc_00E2A03C: call [eax+00000078h]
  loc_00E2A03F: test eax, eax
  loc_00E2A041: fnclex
  loc_00E2A043: jge 00E2A056h
  loc_00E2A045: mov ecx, var_D0
  loc_00E2A04B: push 00000078h
  loc_00E2A04D: push 006DCBE8h
  loc_00E2A052: push ecx
  loc_00E2A053: push eax
  loc_00E2A054: call ebx
  loc_00E2A056: lea edx, var_34
  loc_00E2A059: lea eax, var_30
  loc_00E2A05C: push edx
  loc_00E2A05D: push eax
  loc_00E2A05E: push 00000002h
  loc_00E2A060: call [00401048h] ; __vbaFreeObjList
  loc_00E2A066: add esp, 0000000Ch
  loc_00E2A069: lea ecx, var_50
  loc_00E2A06C: call [00401024h] ; __vbaFreeVar
  loc_00E2A072: mov ecx, [esi]
  loc_00E2A074: push 006DCBD8h
  loc_00E2A079: push 00000000h
  loc_00E2A07B: push 00000018h
  loc_00E2A07D: push esi
  loc_00E2A07E: call [ecx+000004A8h]
  loc_00E2A084: lea edx, var_34
  loc_00E2A087: push eax
  loc_00E2A088: push edx
  loc_00E2A089: call edi
  loc_00E2A08B: push eax
  loc_00E2A08C: lea eax, var_50
  loc_00E2A08F: push eax
  loc_00E2A090: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2A096: add esp, 00000010h
  loc_00E2A099: push eax
  loc_00E2A09A: call [00401120h] ; __vbaCastObjVar
  loc_00E2A0A0: lea ecx, var_38
  loc_00E2A0A3: push eax
  loc_00E2A0A4: push ecx
  loc_00E2A0A5: call edi
  loc_00E2A0A7: mov edx, [eax]
  loc_00E2A0A9: lea ecx, var_3C
  loc_00E2A0AC: push ecx
  loc_00E2A0AD: push eax
  loc_00E2A0AE: mov var_D8, eax
  loc_00E2A0B4: call [edx+00000054h]
  loc_00E2A0B7: test eax, eax
  loc_00E2A0B9: fnclex
  loc_00E2A0BB: jge 00E2A0CEh
  loc_00E2A0BD: mov edx, var_D8
  loc_00E2A0C3: push 00000054h
  loc_00E2A0C5: push 006DCBE8h
  loc_00E2A0CA: push edx
  loc_00E2A0CB: push eax
  loc_00E2A0CC: call ebx
  loc_00E2A0CE: lea edx, var_40
  loc_00E2A0D1: mov eax, 00000002h
  loc_00E2A0D6: push edx
  loc_00E2A0D7: mov ecx, var_3C
  loc_00E2A0DA: sub esp, 00000010h
  loc_00E2A0DD: mov var_90, eax
  loc_00E2A0E3: mov edx, esp
  loc_00E2A0E5: mov var_88, 00000000h
  loc_00E2A0EF: mov var_E0, ecx
  loc_00E2A0F5: mov ecx, [ecx]
  loc_00E2A0F7: mov [edx], eax
  loc_00E2A0F9: mov eax, var_8C
  loc_00E2A0FF: mov [edx+00000004h], eax
  loc_00E2A102: mov eax, var_88
  loc_00E2A108: mov [edx+00000008h], eax
  loc_00E2A10B: mov eax, var_84
  loc_00E2A111: mov [edx+0000000Ch], eax
  loc_00E2A114: mov edx, var_3C
  loc_00E2A117: push edx
  loc_00E2A118: call [ecx+00000028h]
  loc_00E2A11B: test eax, eax
  loc_00E2A11D: fnclex
  loc_00E2A11F: jge 00E2A132h
  loc_00E2A121: mov ecx, var_E0
  loc_00E2A127: push 00000028h
  loc_00E2A129: push 006E09E8h
  loc_00E2A12E: push ecx
  loc_00E2A12F: push eax
  loc_00E2A130: call ebx
  loc_00E2A132: mov edx, var_40
  loc_00E2A135: mov eax, [esi]
  loc_00E2A137: push esi
  loc_00E2A138: mov var_E8, edx
  loc_00E2A13E: call [eax+00000430h]
  loc_00E2A144: lea ecx, var_30
  loc_00E2A147: push eax
  loc_00E2A148: push ecx
  loc_00E2A149: call edi
  loc_00E2A14B: mov edx, [eax]
  loc_00E2A14D: lea ecx, var_28
  loc_00E2A150: push ecx
  loc_00E2A151: push eax
  loc_00E2A152: mov var_D0, eax
  loc_00E2A158: call [edx+00000050h]
  loc_00E2A15B: test eax, eax
  loc_00E2A15D: fnclex
  loc_00E2A15F: jge 00E2A172h
  loc_00E2A161: mov edx, var_D0
  loc_00E2A167: push 00000050h
  loc_00E2A169: push 006E0410h
  loc_00E2A16E: push edx
  loc_00E2A16F: push eax
  loc_00E2A170: call ebx
  loc_00E2A172: mov eax, var_28
  loc_00E2A175: mov ecx, var_E8
  loc_00E2A17B: mov var_58, eax
  loc_00E2A17E: mov eax, 00000008h
  loc_00E2A183: mov var_28, 00000000h
  loc_00E2A18A: mov var_60, eax
  loc_00E2A18D: mov edx, [ecx]
  loc_00E2A18F: sub esp, 00000010h
  loc_00E2A192: mov ecx, esp
  loc_00E2A194: mov [ecx], eax
  loc_00E2A196: mov eax, var_5C
  loc_00E2A199: mov [ecx+00000004h], eax
  loc_00E2A19C: mov eax, var_58
  loc_00E2A19F: mov [ecx+00000008h], eax
  loc_00E2A1A2: mov eax, var_54
  loc_00E2A1A5: mov [ecx+0000000Ch], eax
  loc_00E2A1A8: mov ecx, var_E8
  loc_00E2A1AE: push ecx
  loc_00E2A1AF: call [edx+00000038h]
  loc_00E2A1B2: test eax, eax
  loc_00E2A1B4: fnclex
  loc_00E2A1B6: jge 00E2A1C9h
  loc_00E2A1B8: mov edx, var_E8
  loc_00E2A1BE: push 00000038h
  loc_00E2A1C0: push 006E09F8h
  loc_00E2A1C5: push edx
  loc_00E2A1C6: push eax
  loc_00E2A1C7: call ebx
  loc_00E2A1C9: lea eax, var_40
  loc_00E2A1CC: lea ecx, var_3C
  loc_00E2A1CF: push eax
  loc_00E2A1D0: lea edx, var_38
  loc_00E2A1D3: push ecx
  loc_00E2A1D4: lea eax, var_34
  loc_00E2A1D7: push edx
  loc_00E2A1D8: lea ecx, var_30
  loc_00E2A1DB: push eax
  loc_00E2A1DC: push ecx
  loc_00E2A1DD: push 00000005h
  loc_00E2A1DF: call [00401048h] ; __vbaFreeObjList
  loc_00E2A1E5: lea edx, var_60
  loc_00E2A1E8: lea eax, var_50
  loc_00E2A1EB: push edx
  loc_00E2A1EC: push eax
  loc_00E2A1ED: push 00000002h
  loc_00E2A1EF: call [00401038h] ; __vbaFreeVarList
  loc_00E2A1F5: mov ecx, [esi]
  loc_00E2A1F7: add esp, 00000024h
  loc_00E2A1FA: push 006DCBD8h
  loc_00E2A1FF: push 00000000h
  loc_00E2A201: push 00000018h
  loc_00E2A203: push esi
  loc_00E2A204: call [ecx+000004A8h]
  loc_00E2A20A: lea edx, var_34
  loc_00E2A20D: push eax
  loc_00E2A20E: push edx
  loc_00E2A20F: call edi
  loc_00E2A211: push eax
  loc_00E2A212: lea eax, var_50
  loc_00E2A215: push eax
  loc_00E2A216: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2A21C: add esp, 00000010h
  loc_00E2A21F: push eax
  loc_00E2A220: call [00401120h] ; __vbaCastObjVar
  loc_00E2A226: lea ecx, var_38
  loc_00E2A229: push eax
  loc_00E2A22A: push ecx
  loc_00E2A22B: call edi
  loc_00E2A22D: mov edx, [eax]
  loc_00E2A22F: lea ecx, var_3C
  loc_00E2A232: push ecx
  loc_00E2A233: push eax
  loc_00E2A234: mov var_D8, eax
  loc_00E2A23A: call [edx+00000054h]
  loc_00E2A23D: test eax, eax
  loc_00E2A23F: fnclex
  loc_00E2A241: jge 00E2A254h
  loc_00E2A243: mov edx, var_D8
  loc_00E2A249: push 00000054h
  loc_00E2A24B: push 006DCBE8h
  loc_00E2A250: push edx
  loc_00E2A251: push eax
  loc_00E2A252: call ebx
  loc_00E2A254: lea edx, var_40
  loc_00E2A257: mov eax, 00000002h
  loc_00E2A25C: push edx
  loc_00E2A25D: mov ecx, var_3C
  loc_00E2A260: sub esp, 00000010h
  loc_00E2A263: mov var_90, eax
  loc_00E2A269: mov edx, esp
  loc_00E2A26B: mov var_88, 00000001h
  loc_00E2A275: mov var_E0, ecx
  loc_00E2A27B: mov ecx, [ecx]
  loc_00E2A27D: mov [edx], eax
  loc_00E2A27F: mov eax, var_8C
  loc_00E2A285: mov [edx+00000004h], eax
  loc_00E2A288: mov eax, var_88
  loc_00E2A28E: mov [edx+00000008h], eax
  loc_00E2A291: mov eax, var_84
  loc_00E2A297: mov [edx+0000000Ch], eax
  loc_00E2A29A: mov edx, var_3C
  loc_00E2A29D: push edx
  loc_00E2A29E: call [ecx+00000028h]
  loc_00E2A2A1: test eax, eax
  loc_00E2A2A3: fnclex
  loc_00E2A2A5: jge 00E2A2B8h
  loc_00E2A2A7: mov ecx, var_E0
  loc_00E2A2AD: push 00000028h
  loc_00E2A2AF: push 006E09E8h
  loc_00E2A2B4: push ecx
  loc_00E2A2B5: push eax
  loc_00E2A2B6: call ebx
  loc_00E2A2B8: mov edx, var_40
  loc_00E2A2BB: mov eax, [esi]
  loc_00E2A2BD: push esi
  loc_00E2A2BE: mov var_E8, edx
  loc_00E2A2C4: call [eax+00000434h]
  loc_00E2A2CA: lea ecx, var_30
  loc_00E2A2CD: push eax
  loc_00E2A2CE: push ecx
  loc_00E2A2CF: call edi
  loc_00E2A2D1: mov edx, [eax]
  loc_00E2A2D3: lea ecx, var_28
  loc_00E2A2D6: push ecx
  loc_00E2A2D7: push eax
  loc_00E2A2D8: mov var_D0, eax
  loc_00E2A2DE: call [edx+00000050h]
  loc_00E2A2E1: test eax, eax
  loc_00E2A2E3: fnclex
  loc_00E2A2E5: jge 00E2A2F8h
  loc_00E2A2E7: mov edx, var_D0
  loc_00E2A2ED: push 00000050h
  loc_00E2A2EF: push 006E0410h
  loc_00E2A2F4: push edx
  loc_00E2A2F5: push eax
  loc_00E2A2F6: call ebx
  loc_00E2A2F8: mov eax, var_28
  loc_00E2A2FB: mov ecx, var_E8
  loc_00E2A301: mov var_58, eax
  loc_00E2A304: mov eax, 00000008h
  loc_00E2A309: mov var_28, 00000000h
  loc_00E2A310: mov var_60, eax
  loc_00E2A313: mov edx, [ecx]
  loc_00E2A315: sub esp, 00000010h
  loc_00E2A318: mov ecx, esp
  loc_00E2A31A: mov [ecx], eax
  loc_00E2A31C: mov eax, var_5C
  loc_00E2A31F: mov [ecx+00000004h], eax
  loc_00E2A322: mov eax, var_58
  loc_00E2A325: mov [ecx+00000008h], eax
  loc_00E2A328: mov eax, var_54
  loc_00E2A32B: mov [ecx+0000000Ch], eax
  loc_00E2A32E: mov ecx, var_E8
  loc_00E2A334: push ecx
  loc_00E2A335: call [edx+00000038h]
  loc_00E2A338: test eax, eax
  loc_00E2A33A: fnclex
  loc_00E2A33C: jge 00E2A34Fh
  loc_00E2A33E: mov edx, var_E8
  loc_00E2A344: push 00000038h
  loc_00E2A346: push 006E09F8h
  loc_00E2A34B: push edx
  loc_00E2A34C: push eax
  loc_00E2A34D: call ebx
  loc_00E2A34F: lea eax, var_40
  loc_00E2A352: lea ecx, var_3C
  loc_00E2A355: push eax
  loc_00E2A356: lea edx, var_38
  loc_00E2A359: push ecx
  loc_00E2A35A: lea eax, var_34
  loc_00E2A35D: push edx
  loc_00E2A35E: lea ecx, var_30
  loc_00E2A361: push eax
  loc_00E2A362: push ecx
  loc_00E2A363: push 00000005h
  loc_00E2A365: call [00401048h] ; __vbaFreeObjList
  loc_00E2A36B: lea edx, var_60
  loc_00E2A36E: lea eax, var_50
  loc_00E2A371: push edx
  loc_00E2A372: push eax
  loc_00E2A373: push 00000002h
  loc_00E2A375: call [00401038h] ; __vbaFreeVarList
  loc_00E2A37B: mov ecx, [esi]
  loc_00E2A37D: add esp, 00000024h
  loc_00E2A380: push 006DCBD8h
  loc_00E2A385: push 00000000h
  loc_00E2A387: push 00000018h
  loc_00E2A389: push esi
  loc_00E2A38A: call [ecx+000004A8h]
  loc_00E2A390: lea edx, var_34
  loc_00E2A393: push eax
  loc_00E2A394: push edx
  loc_00E2A395: call edi
  loc_00E2A397: push eax
  loc_00E2A398: lea eax, var_50
  loc_00E2A39B: push eax
  loc_00E2A39C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2A3A2: add esp, 00000010h
  loc_00E2A3A5: push eax
  loc_00E2A3A6: call [00401120h] ; __vbaCastObjVar
  loc_00E2A3AC: lea ecx, var_38
  loc_00E2A3AF: push eax
  loc_00E2A3B0: push ecx
  loc_00E2A3B1: call edi
  loc_00E2A3B3: mov edx, [eax]
  loc_00E2A3B5: lea ecx, var_3C
  loc_00E2A3B8: push ecx
  loc_00E2A3B9: push eax
  loc_00E2A3BA: mov var_D8, eax
  loc_00E2A3C0: call [edx+00000054h]
  loc_00E2A3C3: test eax, eax
  loc_00E2A3C5: fnclex
  loc_00E2A3C7: jge 00E2A3DAh
  loc_00E2A3C9: mov edx, var_D8
  loc_00E2A3CF: push 00000054h
  loc_00E2A3D1: push 006DCBE8h
  loc_00E2A3D6: push edx
  loc_00E2A3D7: push eax
  loc_00E2A3D8: call ebx
  loc_00E2A3DA: lea edx, var_40
  loc_00E2A3DD: mov eax, 00000002h
  loc_00E2A3E2: push edx
  loc_00E2A3E3: mov ecx, var_3C
  loc_00E2A3E6: sub esp, 00000010h
  loc_00E2A3E9: mov var_88, eax
  loc_00E2A3EF: mov edx, esp
  loc_00E2A3F1: mov var_90, eax
  loc_00E2A3F7: mov var_E0, ecx
  loc_00E2A3FD: mov ecx, [ecx]
  loc_00E2A3FF: mov [edx], eax
  loc_00E2A401: mov eax, var_8C
  loc_00E2A407: mov [edx+00000004h], eax
  loc_00E2A40A: mov eax, var_88
  loc_00E2A410: mov [edx+00000008h], eax
  loc_00E2A413: mov eax, var_84
  loc_00E2A419: mov [edx+0000000Ch], eax
  loc_00E2A41C: mov edx, var_3C
  loc_00E2A41F: push edx
  loc_00E2A420: call [ecx+00000028h]
  loc_00E2A423: test eax, eax
  loc_00E2A425: fnclex
  loc_00E2A427: jge 00E2A43Ah
  loc_00E2A429: mov ecx, var_E0
  loc_00E2A42F: push 00000028h
  loc_00E2A431: push 006E09E8h
  loc_00E2A436: push ecx
  loc_00E2A437: push eax
  loc_00E2A438: call ebx
  loc_00E2A43A: mov edx, var_40
  loc_00E2A43D: mov eax, [esi]
  loc_00E2A43F: push esi
  loc_00E2A440: mov var_E8, edx
  loc_00E2A446: call [eax+00000438h]
  loc_00E2A44C: lea ecx, var_30
  loc_00E2A44F: push eax
  loc_00E2A450: push ecx
  loc_00E2A451: call edi
  loc_00E2A453: mov edx, [eax]
  loc_00E2A455: lea ecx, var_28
  loc_00E2A458: push ecx
  loc_00E2A459: push eax
  loc_00E2A45A: mov var_D0, eax
  loc_00E2A460: call [edx+00000050h]
  loc_00E2A463: test eax, eax
  loc_00E2A465: fnclex
  loc_00E2A467: jge 00E2A47Ah
  loc_00E2A469: mov edx, var_D0
  loc_00E2A46F: push 00000050h
  loc_00E2A471: push 006E0410h
  loc_00E2A476: push edx
  loc_00E2A477: push eax
  loc_00E2A478: call ebx
  loc_00E2A47A: mov eax, var_28
  loc_00E2A47D: mov ecx, var_E8
  loc_00E2A483: mov var_58, eax
  loc_00E2A486: mov eax, 00000008h
  loc_00E2A48B: mov var_28, 00000000h
  loc_00E2A492: mov var_60, eax
  loc_00E2A495: mov edx, [ecx]
  loc_00E2A497: sub esp, 00000010h
  loc_00E2A49A: mov ecx, esp
  loc_00E2A49C: mov [ecx], eax
  loc_00E2A49E: mov eax, var_5C
  loc_00E2A4A1: mov [ecx+00000004h], eax
  loc_00E2A4A4: mov eax, var_58
  loc_00E2A4A7: mov [ecx+00000008h], eax
  loc_00E2A4AA: mov eax, var_54
  loc_00E2A4AD: mov [ecx+0000000Ch], eax
  loc_00E2A4B0: mov ecx, var_E8
  loc_00E2A4B6: push ecx
  loc_00E2A4B7: call [edx+00000038h]
  loc_00E2A4BA: test eax, eax
  loc_00E2A4BC: fnclex
  loc_00E2A4BE: jge 00E2A4D1h
  loc_00E2A4C0: mov edx, var_E8
  loc_00E2A4C6: push 00000038h
  loc_00E2A4C8: push 006E09F8h
  loc_00E2A4CD: push edx
  loc_00E2A4CE: push eax
  loc_00E2A4CF: call ebx
  loc_00E2A4D1: lea eax, var_40
  loc_00E2A4D4: lea ecx, var_3C
  loc_00E2A4D7: push eax
  loc_00E2A4D8: lea edx, var_38
  loc_00E2A4DB: push ecx
  loc_00E2A4DC: lea eax, var_34
  loc_00E2A4DF: push edx
  loc_00E2A4E0: lea ecx, var_30
  loc_00E2A4E3: push eax
  loc_00E2A4E4: push ecx
  loc_00E2A4E5: push 00000005h
  loc_00E2A4E7: call [00401048h] ; __vbaFreeObjList
  loc_00E2A4ED: lea edx, var_60
  loc_00E2A4F0: lea eax, var_50
  loc_00E2A4F3: push edx
  loc_00E2A4F4: push eax
  loc_00E2A4F5: push 00000002h
  loc_00E2A4F7: call [00401038h] ; __vbaFreeVarList
  loc_00E2A4FD: mov ecx, [esi]
  loc_00E2A4FF: add esp, 00000024h
  loc_00E2A502: push 006DCBD8h
  loc_00E2A507: push 00000000h
  loc_00E2A509: push 00000018h
  loc_00E2A50B: push esi
  loc_00E2A50C: call [ecx+000004A8h]
  loc_00E2A512: lea edx, var_34
  loc_00E2A515: push eax
  loc_00E2A516: push edx
  loc_00E2A517: call edi
  loc_00E2A519: push eax
  loc_00E2A51A: lea eax, var_50
  loc_00E2A51D: push eax
  loc_00E2A51E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2A524: add esp, 00000010h
  loc_00E2A527: push eax
  loc_00E2A528: call [00401120h] ; __vbaCastObjVar
  loc_00E2A52E: lea ecx, var_38
  loc_00E2A531: push eax
  loc_00E2A532: push ecx
  loc_00E2A533: call edi
  loc_00E2A535: mov edx, [eax]
  loc_00E2A537: lea ecx, var_3C
  loc_00E2A53A: push ecx
  loc_00E2A53B: push eax
  loc_00E2A53C: mov var_D8, eax
  loc_00E2A542: call [edx+00000054h]
  loc_00E2A545: test eax, eax
  loc_00E2A547: fnclex
  loc_00E2A549: jge 00E2A55Ch
  loc_00E2A54B: mov edx, var_D8
  loc_00E2A551: push 00000054h
  loc_00E2A553: push 006DCBE8h
  loc_00E2A558: push edx
  loc_00E2A559: push eax
  loc_00E2A55A: call ebx
  loc_00E2A55C: lea edx, var_40
  loc_00E2A55F: mov eax, 00000002h
  loc_00E2A564: push edx
  loc_00E2A565: mov ecx, var_3C
  loc_00E2A568: sub esp, 00000010h
  loc_00E2A56B: mov var_90, eax
  loc_00E2A571: mov edx, esp
  loc_00E2A573: mov var_88, 00000003h
  loc_00E2A57D: mov var_E0, ecx
  loc_00E2A583: mov ecx, [ecx]
  loc_00E2A585: mov [edx], eax
  loc_00E2A587: mov eax, var_8C
  loc_00E2A58D: mov [edx+00000004h], eax
  loc_00E2A590: mov eax, var_88
  loc_00E2A596: mov [edx+00000008h], eax
  loc_00E2A599: mov eax, var_84
  loc_00E2A59F: mov [edx+0000000Ch], eax
  loc_00E2A5A2: mov edx, var_3C
  loc_00E2A5A5: push edx
  loc_00E2A5A6: call [ecx+00000028h]
  loc_00E2A5A9: test eax, eax
  loc_00E2A5AB: fnclex
  loc_00E2A5AD: jge 00E2A5C0h
  loc_00E2A5AF: mov ecx, var_E0
  loc_00E2A5B5: push 00000028h
  loc_00E2A5B7: push 006E09E8h
  loc_00E2A5BC: push ecx
  loc_00E2A5BD: push eax
  loc_00E2A5BE: call ebx
  loc_00E2A5C0: mov edx, var_40
  loc_00E2A5C3: mov eax, [esi]
  loc_00E2A5C5: push esi
  loc_00E2A5C6: mov var_E8, edx
  loc_00E2A5CC: call [eax+0000044Ch]
  loc_00E2A5D2: lea ecx, var_30
  loc_00E2A5D5: push eax
  loc_00E2A5D6: push ecx
  loc_00E2A5D7: call edi
  loc_00E2A5D9: mov edx, [eax]
  loc_00E2A5DB: lea ecx, var_28
  loc_00E2A5DE: push ecx
  loc_00E2A5DF: push eax
  loc_00E2A5E0: mov var_D0, eax
  loc_00E2A5E6: call [edx+00000050h]
  loc_00E2A5E9: test eax, eax
  loc_00E2A5EB: fnclex
  loc_00E2A5ED: jge 00E2A600h
  loc_00E2A5EF: mov edx, var_D0
  loc_00E2A5F5: push 00000050h
  loc_00E2A5F7: push 006E0410h
  loc_00E2A5FC: push edx
  loc_00E2A5FD: push eax
  loc_00E2A5FE: call ebx
  loc_00E2A600: mov eax, var_28
  loc_00E2A603: mov ecx, var_E8
  loc_00E2A609: mov var_58, eax
  loc_00E2A60C: mov eax, 00000008h
  loc_00E2A611: mov var_28, 00000000h
  loc_00E2A618: mov var_60, eax
  loc_00E2A61B: mov edx, [ecx]
  loc_00E2A61D: sub esp, 00000010h
  loc_00E2A620: mov ecx, esp
  loc_00E2A622: mov [ecx], eax
  loc_00E2A624: mov eax, var_5C
  loc_00E2A627: mov [ecx+00000004h], eax
  loc_00E2A62A: mov eax, var_58
  loc_00E2A62D: mov [ecx+00000008h], eax
  loc_00E2A630: mov eax, var_54
  loc_00E2A633: mov [ecx+0000000Ch], eax
  loc_00E2A636: mov ecx, var_E8
  loc_00E2A63C: push ecx
  loc_00E2A63D: call [edx+00000038h]
  loc_00E2A640: test eax, eax
  loc_00E2A642: fnclex
  loc_00E2A644: jge 00E2A657h
  loc_00E2A646: mov edx, var_E8
  loc_00E2A64C: push 00000038h
  loc_00E2A64E: push 006E09F8h
  loc_00E2A653: push edx
  loc_00E2A654: push eax
  loc_00E2A655: call ebx
  loc_00E2A657: lea eax, var_40
  loc_00E2A65A: lea ecx, var_3C
  loc_00E2A65D: push eax
  loc_00E2A65E: lea edx, var_38
  loc_00E2A661: push ecx
  loc_00E2A662: lea eax, var_34
  loc_00E2A665: push edx
  loc_00E2A666: lea ecx, var_30
  loc_00E2A669: push eax
  loc_00E2A66A: push ecx
  loc_00E2A66B: push 00000005h
  loc_00E2A66D: call [00401048h] ; __vbaFreeObjList
  loc_00E2A673: lea edx, var_60
  loc_00E2A676: lea eax, var_50
  loc_00E2A679: push edx
  loc_00E2A67A: push eax
  loc_00E2A67B: push 00000002h
  loc_00E2A67D: call [00401038h] ; __vbaFreeVarList
  loc_00E2A683: mov ecx, [esi]
  loc_00E2A685: add esp, 00000024h
  loc_00E2A688: push 006DCBD8h
  loc_00E2A68D: push 00000000h
  loc_00E2A68F: push 00000018h
  loc_00E2A691: push esi
  loc_00E2A692: call [ecx+000004A8h]
  loc_00E2A698: lea edx, var_34
  loc_00E2A69B: push eax
  loc_00E2A69C: push edx
  loc_00E2A69D: call edi
  loc_00E2A69F: push eax
  loc_00E2A6A0: lea eax, var_50
  loc_00E2A6A3: push eax
  loc_00E2A6A4: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2A6AA: add esp, 00000010h
  loc_00E2A6AD: push eax
  loc_00E2A6AE: call [00401120h] ; __vbaCastObjVar
  loc_00E2A6B4: lea ecx, var_38
  loc_00E2A6B7: push eax
  loc_00E2A6B8: push ecx
  loc_00E2A6B9: call edi
  loc_00E2A6BB: mov edx, [eax]
  loc_00E2A6BD: lea ecx, var_3C
  loc_00E2A6C0: push ecx
  loc_00E2A6C1: push eax
  loc_00E2A6C2: mov var_D8, eax
  loc_00E2A6C8: call [edx+00000054h]
  loc_00E2A6CB: test eax, eax
  loc_00E2A6CD: fnclex
  loc_00E2A6CF: jge 00E2A6E2h
  loc_00E2A6D1: mov edx, var_D8
  loc_00E2A6D7: push 00000054h
  loc_00E2A6D9: push 006DCBE8h
  loc_00E2A6DE: push edx
  loc_00E2A6DF: push eax
  loc_00E2A6E0: call ebx
  loc_00E2A6E2: lea edx, var_40
  loc_00E2A6E5: mov eax, 00000002h
  loc_00E2A6EA: push edx
  loc_00E2A6EB: mov ecx, var_3C
  loc_00E2A6EE: sub esp, 00000010h
  loc_00E2A6F1: mov var_90, eax
  loc_00E2A6F7: mov edx, esp
  loc_00E2A6F9: mov var_88, 0000000Eh
  loc_00E2A703: mov var_E0, ecx
  loc_00E2A709: mov ecx, [ecx]
  loc_00E2A70B: mov [edx], eax
  loc_00E2A70D: mov eax, var_8C
  loc_00E2A713: mov [edx+00000004h], eax
  loc_00E2A716: mov eax, var_88
  loc_00E2A71C: mov [edx+00000008h], eax
  loc_00E2A71F: mov eax, var_84
  loc_00E2A725: mov [edx+0000000Ch], eax
  loc_00E2A728: mov edx, var_3C
  loc_00E2A72B: push edx
  loc_00E2A72C: call [ecx+00000028h]
  loc_00E2A72F: test eax, eax
  loc_00E2A731: fnclex
  loc_00E2A733: jge 00E2A746h
  loc_00E2A735: mov ecx, var_E0
  loc_00E2A73B: push 00000028h
  loc_00E2A73D: push 006E09E8h
  loc_00E2A742: push ecx
  loc_00E2A743: push eax
  loc_00E2A744: call ebx
  loc_00E2A746: mov edx, var_40
  loc_00E2A749: mov eax, [esi]
  loc_00E2A74B: push esi
  loc_00E2A74C: mov var_E8, edx
  loc_00E2A752: call [eax+00000300h]
  loc_00E2A758: lea ecx, var_30
  loc_00E2A75B: push eax
  loc_00E2A75C: push ecx
  loc_00E2A75D: call edi
  loc_00E2A75F: mov edx, [eax]
  loc_00E2A761: lea ecx, var_28
  loc_00E2A764: push ecx
  loc_00E2A765: push eax
  loc_00E2A766: mov var_D0, eax
  loc_00E2A76C: call [edx+000000A0h]
  loc_00E2A772: test eax, eax
  loc_00E2A774: fnclex
  loc_00E2A776: jge 00E2A78Ch
  loc_00E2A778: mov edx, var_D0
  loc_00E2A77E: push 000000A0h
  loc_00E2A783: push 006DCB70h
  loc_00E2A788: push edx
  loc_00E2A789: push eax
  loc_00E2A78A: call ebx
  loc_00E2A78C: mov eax, var_28
  loc_00E2A78F: mov ecx, var_E8
  loc_00E2A795: mov var_58, eax
  loc_00E2A798: mov eax, 00000008h
  loc_00E2A79D: mov var_28, 00000000h
  loc_00E2A7A4: mov var_60, eax
  loc_00E2A7A7: mov edx, [ecx]
  loc_00E2A7A9: sub esp, 00000010h
  loc_00E2A7AC: mov ecx, esp
  loc_00E2A7AE: mov [ecx], eax
  loc_00E2A7B0: mov eax, var_5C
  loc_00E2A7B3: mov [ecx+00000004h], eax
  loc_00E2A7B6: mov eax, var_58
  loc_00E2A7B9: mov [ecx+00000008h], eax
  loc_00E2A7BC: mov eax, var_54
  loc_00E2A7BF: mov [ecx+0000000Ch], eax
  loc_00E2A7C2: mov ecx, var_E8
  loc_00E2A7C8: push ecx
  loc_00E2A7C9: call [edx+00000038h]
  loc_00E2A7CC: test eax, eax
  loc_00E2A7CE: fnclex
  loc_00E2A7D0: jge 00E2A7E3h
  loc_00E2A7D2: mov edx, var_E8
  loc_00E2A7D8: push 00000038h
  loc_00E2A7DA: push 006E09F8h
  loc_00E2A7DF: push edx
  loc_00E2A7E0: push eax
  loc_00E2A7E1: call ebx
  loc_00E2A7E3: lea eax, var_40
  loc_00E2A7E6: lea ecx, var_3C
  loc_00E2A7E9: push eax
  loc_00E2A7EA: lea edx, var_38
  loc_00E2A7ED: push ecx
  loc_00E2A7EE: lea eax, var_34
  loc_00E2A7F1: push edx
  loc_00E2A7F2: lea ecx, var_30
  loc_00E2A7F5: push eax
  loc_00E2A7F6: push ecx
  loc_00E2A7F7: push 00000005h
  loc_00E2A7F9: call [00401048h] ; __vbaFreeObjList
  loc_00E2A7FF: lea edx, var_60
  loc_00E2A802: lea eax, var_50
  loc_00E2A805: push edx
  loc_00E2A806: push eax
  loc_00E2A807: push 00000002h
  loc_00E2A809: call [00401038h] ; __vbaFreeVarList
  loc_00E2A80F: mov ecx, [esi]
  loc_00E2A811: add esp, 00000024h
  loc_00E2A814: push 006DCBD8h
  loc_00E2A819: push 00000000h
  loc_00E2A81B: push 00000018h
  loc_00E2A81D: push esi
  loc_00E2A81E: call [ecx+000004A8h]
  loc_00E2A824: lea edx, var_34
  loc_00E2A827: push eax
  loc_00E2A828: push edx
  loc_00E2A829: call edi
  loc_00E2A82B: push eax
  loc_00E2A82C: lea eax, var_50
  loc_00E2A82F: push eax
  loc_00E2A830: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2A836: add esp, 00000010h
  loc_00E2A839: push eax
  loc_00E2A83A: call [00401120h] ; __vbaCastObjVar
  loc_00E2A840: lea ecx, var_38
  loc_00E2A843: push eax
  loc_00E2A844: push ecx
  loc_00E2A845: call edi
  loc_00E2A847: mov edx, [eax]
  loc_00E2A849: lea ecx, var_3C
  loc_00E2A84C: push ecx
  loc_00E2A84D: push eax
  loc_00E2A84E: mov var_D8, eax
  loc_00E2A854: call [edx+00000054h]
  loc_00E2A857: test eax, eax
  loc_00E2A859: fnclex
  loc_00E2A85B: jge 00E2A86Eh
  loc_00E2A85D: mov edx, var_D8
  loc_00E2A863: push 00000054h
  loc_00E2A865: push 006DCBE8h
  loc_00E2A86A: push edx
  loc_00E2A86B: push eax
  loc_00E2A86C: call ebx
  loc_00E2A86E: lea edx, var_40
  loc_00E2A871: mov eax, 00000002h
  loc_00E2A876: push edx
  loc_00E2A877: mov ecx, var_3C
  loc_00E2A87A: sub esp, 00000010h
  loc_00E2A87D: mov var_90, eax
  loc_00E2A883: mov edx, esp
  loc_00E2A885: mov var_88, 00000004h
  loc_00E2A88F: mov var_E0, ecx
  loc_00E2A895: mov ecx, [ecx]
  loc_00E2A897: mov [edx], eax
  loc_00E2A899: mov eax, var_8C
  loc_00E2A89F: mov [edx+00000004h], eax
  loc_00E2A8A2: mov eax, var_88
  loc_00E2A8A8: mov [edx+00000008h], eax
  loc_00E2A8AB: mov eax, var_84
  loc_00E2A8B1: mov [edx+0000000Ch], eax
  loc_00E2A8B4: mov edx, var_3C
  loc_00E2A8B7: push edx
  loc_00E2A8B8: call [ecx+00000028h]
  loc_00E2A8BB: test eax, eax
  loc_00E2A8BD: fnclex
  loc_00E2A8BF: jge 00E2A8D2h
  loc_00E2A8C1: mov ecx, var_E0
  loc_00E2A8C7: push 00000028h
  loc_00E2A8C9: push 006E09E8h
  loc_00E2A8CE: push ecx
  loc_00E2A8CF: push eax
  loc_00E2A8D0: call ebx
  loc_00E2A8D2: mov edx, var_40
  loc_00E2A8D5: mov eax, [esi]
  loc_00E2A8D7: push esi
  loc_00E2A8D8: mov var_E8, edx
  loc_00E2A8DE: call [eax+00000450h]
  loc_00E2A8E4: lea ecx, var_30
  loc_00E2A8E7: push eax
  loc_00E2A8E8: push ecx
  loc_00E2A8E9: call edi
  loc_00E2A8EB: mov edx, [eax]
  loc_00E2A8ED: lea ecx, var_28
  loc_00E2A8F0: push ecx
  loc_00E2A8F1: push eax
  loc_00E2A8F2: mov var_D0, eax
  loc_00E2A8F8: call [edx+00000050h]
  loc_00E2A8FB: test eax, eax
  loc_00E2A8FD: fnclex
  loc_00E2A8FF: jge 00E2A912h
  loc_00E2A901: mov edx, var_D0
  loc_00E2A907: push 00000050h
  loc_00E2A909: push 006E0410h
  loc_00E2A90E: push edx
  loc_00E2A90F: push eax
  loc_00E2A910: call ebx
  loc_00E2A912: mov eax, var_28
  loc_00E2A915: mov ecx, var_E8
  loc_00E2A91B: mov var_58, eax
  loc_00E2A91E: mov eax, 00000008h
  loc_00E2A923: mov var_28, 00000000h
  loc_00E2A92A: mov var_60, eax
  loc_00E2A92D: mov edx, [ecx]
  loc_00E2A92F: sub esp, 00000010h
  loc_00E2A932: mov ecx, esp
  loc_00E2A934: mov [ecx], eax
  loc_00E2A936: mov eax, var_5C
  loc_00E2A939: mov [ecx+00000004h], eax
  loc_00E2A93C: mov eax, var_58
  loc_00E2A93F: mov [ecx+00000008h], eax
  loc_00E2A942: mov eax, var_54
  loc_00E2A945: mov [ecx+0000000Ch], eax
  loc_00E2A948: mov ecx, var_E8
  loc_00E2A94E: push ecx
  loc_00E2A94F: call [edx+00000038h]
  loc_00E2A952: test eax, eax
  loc_00E2A954: fnclex
  loc_00E2A956: jge 00E2A969h
  loc_00E2A958: mov edx, var_E8
  loc_00E2A95E: push 00000038h
  loc_00E2A960: push 006E09F8h
  loc_00E2A965: push edx
  loc_00E2A966: push eax
  loc_00E2A967: call ebx
  loc_00E2A969: lea eax, var_40
  loc_00E2A96C: lea ecx, var_3C
  loc_00E2A96F: push eax
  loc_00E2A970: lea edx, var_38
  loc_00E2A973: push ecx
  loc_00E2A974: lea eax, var_34
  loc_00E2A977: push edx
  loc_00E2A978: lea ecx, var_30
  loc_00E2A97B: push eax
  loc_00E2A97C: push ecx
  loc_00E2A97D: push 00000005h
  loc_00E2A97F: call [00401048h] ; __vbaFreeObjList
  loc_00E2A985: lea edx, var_60
  loc_00E2A988: lea eax, var_50
  loc_00E2A98B: push edx
  loc_00E2A98C: push eax
  loc_00E2A98D: push 00000002h
  loc_00E2A98F: call [00401038h] ; __vbaFreeVarList
  loc_00E2A995: mov ecx, [esi]
  loc_00E2A997: add esp, 00000024h
  loc_00E2A99A: push 006DCBD8h
  loc_00E2A99F: push 00000000h
  loc_00E2A9A1: push 00000018h
  loc_00E2A9A3: push esi
  loc_00E2A9A4: call [ecx+000004A8h]
  loc_00E2A9AA: lea edx, var_34
  loc_00E2A9AD: push eax
  loc_00E2A9AE: push edx
  loc_00E2A9AF: call edi
  loc_00E2A9B1: push eax
  loc_00E2A9B2: lea eax, var_50
  loc_00E2A9B5: push eax
  loc_00E2A9B6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2A9BC: add esp, 00000010h
  loc_00E2A9BF: push eax
  loc_00E2A9C0: call [00401120h] ; __vbaCastObjVar
  loc_00E2A9C6: lea ecx, var_38
  loc_00E2A9C9: push eax
  loc_00E2A9CA: push ecx
  loc_00E2A9CB: call edi
  loc_00E2A9CD: mov edx, [eax]
  loc_00E2A9CF: lea ecx, var_3C
  loc_00E2A9D2: push ecx
  loc_00E2A9D3: push eax
  loc_00E2A9D4: mov var_D8, eax
  loc_00E2A9DA: call [edx+00000054h]
  loc_00E2A9DD: test eax, eax
  loc_00E2A9DF: fnclex
  loc_00E2A9E1: jge 00E2A9F4h
  loc_00E2A9E3: mov edx, var_D8
  loc_00E2A9E9: push 00000054h
  loc_00E2A9EB: push 006DCBE8h
  loc_00E2A9F0: push edx
  loc_00E2A9F1: push eax
  loc_00E2A9F2: call ebx
  loc_00E2A9F4: lea edx, var_40
  loc_00E2A9F7: mov eax, 00000002h
  loc_00E2A9FC: push edx
  loc_00E2A9FD: mov ecx, var_3C
  loc_00E2AA00: sub esp, 00000010h
  loc_00E2AA03: mov var_90, eax
  loc_00E2AA09: mov edx, esp
  loc_00E2AA0B: mov var_88, 00000005h
  loc_00E2AA15: mov var_E0, ecx
  loc_00E2AA1B: mov ecx, [ecx]
  loc_00E2AA1D: mov [edx], eax
  loc_00E2AA1F: mov eax, var_8C
  loc_00E2AA25: mov [edx+00000004h], eax
  loc_00E2AA28: mov eax, var_88
  loc_00E2AA2E: mov [edx+00000008h], eax
  loc_00E2AA31: mov eax, var_84
  loc_00E2AA37: mov [edx+0000000Ch], eax
  loc_00E2AA3A: mov edx, var_3C
  loc_00E2AA3D: push edx
  loc_00E2AA3E: call [ecx+00000028h]
  loc_00E2AA41: test eax, eax
  loc_00E2AA43: fnclex
  loc_00E2AA45: jge 00E2AA58h
  loc_00E2AA47: mov ecx, var_E0
  loc_00E2AA4D: push 00000028h
  loc_00E2AA4F: push 006E09E8h
  loc_00E2AA54: push ecx
  loc_00E2AA55: push eax
  loc_00E2AA56: call ebx
  loc_00E2AA58: mov edx, var_40
  loc_00E2AA5B: mov eax, [esi]
  loc_00E2AA5D: push esi
  loc_00E2AA5E: mov var_E8, edx
  loc_00E2AA64: call [eax+00000424h]
  loc_00E2AA6A: lea ecx, var_30
  loc_00E2AA6D: push eax
  loc_00E2AA6E: push ecx
  loc_00E2AA6F: call edi
  loc_00E2AA71: mov edx, [eax]
  loc_00E2AA73: lea ecx, var_28
  loc_00E2AA76: push ecx
  loc_00E2AA77: push eax
  loc_00E2AA78: mov var_D0, eax
  loc_00E2AA7E: call [edx+00000050h]
  loc_00E2AA81: test eax, eax
  loc_00E2AA83: fnclex
  loc_00E2AA85: jge 00E2AA98h
  loc_00E2AA87: mov edx, var_D0
  loc_00E2AA8D: push 00000050h
  loc_00E2AA8F: push 006E0410h
  loc_00E2AA94: push edx
  loc_00E2AA95: push eax
  loc_00E2AA96: call ebx
  loc_00E2AA98: mov eax, var_28
  loc_00E2AA9B: mov ecx, var_E8
  loc_00E2AAA1: mov var_58, eax
  loc_00E2AAA4: mov eax, 00000008h
  loc_00E2AAA9: mov var_28, 00000000h
  loc_00E2AAB0: mov var_60, eax
  loc_00E2AAB3: mov edx, [ecx]
  loc_00E2AAB5: sub esp, 00000010h
  loc_00E2AAB8: mov ecx, esp
  loc_00E2AABA: mov [ecx], eax
  loc_00E2AABC: mov eax, var_5C
  loc_00E2AABF: mov [ecx+00000004h], eax
  loc_00E2AAC2: mov eax, var_58
  loc_00E2AAC5: mov [ecx+00000008h], eax
  loc_00E2AAC8: mov eax, var_54
  loc_00E2AACB: mov [ecx+0000000Ch], eax
  loc_00E2AACE: mov ecx, var_E8
  loc_00E2AAD4: push ecx
  loc_00E2AAD5: call [edx+00000038h]
  loc_00E2AAD8: test eax, eax
  loc_00E2AADA: fnclex
  loc_00E2AADC: jge 00E2AAEFh
  loc_00E2AADE: mov edx, var_E8
  loc_00E2AAE4: push 00000038h
  loc_00E2AAE6: push 006E09F8h
  loc_00E2AAEB: push edx
  loc_00E2AAEC: push eax
  loc_00E2AAED: call ebx
  loc_00E2AAEF: lea eax, var_40
  loc_00E2AAF2: lea ecx, var_3C
  loc_00E2AAF5: push eax
  loc_00E2AAF6: lea edx, var_38
  loc_00E2AAF9: push ecx
  loc_00E2AAFA: lea eax, var_34
  loc_00E2AAFD: push edx
  loc_00E2AAFE: lea ecx, var_30
  loc_00E2AB01: push eax
  loc_00E2AB02: push ecx
  loc_00E2AB03: push 00000005h
  loc_00E2AB05: call [00401048h] ; __vbaFreeObjList
  loc_00E2AB0B: lea edx, var_60
  loc_00E2AB0E: lea eax, var_50
  loc_00E2AB11: push edx
  loc_00E2AB12: push eax
  loc_00E2AB13: push 00000002h
  loc_00E2AB15: call [00401038h] ; __vbaFreeVarList
  loc_00E2AB1B: mov ecx, [esi]
  loc_00E2AB1D: add esp, 00000024h
  loc_00E2AB20: push 006DCBD8h
  loc_00E2AB25: push 00000000h
  loc_00E2AB27: push 00000018h
  loc_00E2AB29: push esi
  loc_00E2AB2A: call [ecx+000004A8h]
  loc_00E2AB30: lea edx, var_34
  loc_00E2AB33: push eax
  loc_00E2AB34: push edx
  loc_00E2AB35: call edi
  loc_00E2AB37: push eax
  loc_00E2AB38: lea eax, var_50
  loc_00E2AB3B: push eax
  loc_00E2AB3C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2AB42: add esp, 00000010h
  loc_00E2AB45: push eax
  loc_00E2AB46: call [00401120h] ; __vbaCastObjVar
  loc_00E2AB4C: lea ecx, var_38
  loc_00E2AB4F: push eax
  loc_00E2AB50: push ecx
  loc_00E2AB51: call edi
  loc_00E2AB53: mov edx, [eax]
  loc_00E2AB55: lea ecx, var_3C
  loc_00E2AB58: push ecx
  loc_00E2AB59: push eax
  loc_00E2AB5A: mov var_D8, eax
  loc_00E2AB60: call [edx+00000054h]
  loc_00E2AB63: test eax, eax
  loc_00E2AB65: fnclex
  loc_00E2AB67: jge 00E2AB7Ah
  loc_00E2AB69: mov edx, var_D8
  loc_00E2AB6F: push 00000054h
  loc_00E2AB71: push 006DCBE8h
  loc_00E2AB76: push edx
  loc_00E2AB77: push eax
  loc_00E2AB78: call ebx
  loc_00E2AB7A: lea edx, var_40
  loc_00E2AB7D: mov eax, 00000002h
  loc_00E2AB82: push edx
  loc_00E2AB83: mov ecx, var_3C
  loc_00E2AB86: sub esp, 00000010h
  loc_00E2AB89: mov var_90, eax
  loc_00E2AB8F: mov edx, esp
  loc_00E2AB91: mov var_88, 0000000Fh
  loc_00E2AB9B: mov var_E0, ecx
  loc_00E2ABA1: mov ecx, [ecx]
  loc_00E2ABA3: mov [edx], eax
  loc_00E2ABA5: mov eax, var_8C
  loc_00E2ABAB: mov [edx+00000004h], eax
  loc_00E2ABAE: mov eax, var_88
  loc_00E2ABB4: mov [edx+00000008h], eax
  loc_00E2ABB7: mov eax, var_84
  loc_00E2ABBD: mov [edx+0000000Ch], eax
  loc_00E2ABC0: mov edx, var_3C
  loc_00E2ABC3: push edx
  loc_00E2ABC4: call [ecx+00000028h]
  loc_00E2ABC7: test eax, eax
  loc_00E2ABC9: fnclex
  loc_00E2ABCB: jge 00E2ABDEh
  loc_00E2ABCD: mov ecx, var_E0
  loc_00E2ABD3: push 00000028h
  loc_00E2ABD5: push 006E09E8h
  loc_00E2ABDA: push ecx
  loc_00E2ABDB: push eax
  loc_00E2ABDC: call ebx
  loc_00E2ABDE: mov edx, var_40
  loc_00E2ABE1: mov eax, [esi]
  loc_00E2ABE3: push esi
  loc_00E2ABE4: mov var_E8, edx
  loc_00E2ABEA: call [eax+000002FCh]
  loc_00E2ABF0: lea ecx, var_30
  loc_00E2ABF3: push eax
  loc_00E2ABF4: push ecx
  loc_00E2ABF5: call edi
  loc_00E2ABF7: mov edx, [eax]
  loc_00E2ABF9: lea ecx, var_28
  loc_00E2ABFC: push ecx
  loc_00E2ABFD: push eax
  loc_00E2ABFE: mov var_D0, eax
  loc_00E2AC04: call [edx+000000A0h]
  loc_00E2AC0A: test eax, eax
  loc_00E2AC0C: fnclex
  loc_00E2AC0E: jge 00E2AC24h
  loc_00E2AC10: mov edx, var_D0
  loc_00E2AC16: push 000000A0h
  loc_00E2AC1B: push 006DCB70h
  loc_00E2AC20: push edx
  loc_00E2AC21: push eax
  loc_00E2AC22: call ebx
  loc_00E2AC24: mov eax, var_28
  loc_00E2AC27: push eax
  loc_00E2AC28: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2AC2E: mov ecx, var_E8
  loc_00E2AC34: mov eax, 00000005h
  loc_00E2AC39: fstp real8 ptr var_98
  loc_00E2AC3F: mov var_A0, eax
  loc_00E2AC45: mov edx, [ecx]
  loc_00E2AC47: sub esp, 00000010h
  loc_00E2AC4A: mov ecx, esp
  loc_00E2AC4C: mov [ecx], eax
  loc_00E2AC4E: mov eax, var_9C
  loc_00E2AC54: mov [ecx+00000004h], eax
  loc_00E2AC57: mov eax, var_98
  loc_00E2AC5D: mov [ecx+00000008h], eax
  loc_00E2AC60: mov eax, var_94
  loc_00E2AC66: mov [ecx+0000000Ch], eax
  loc_00E2AC69: mov ecx, var_E8
  loc_00E2AC6F: push ecx
  loc_00E2AC70: call [edx+00000038h]
  loc_00E2AC73: test eax, eax
  loc_00E2AC75: fnclex
  loc_00E2AC77: jge 00E2AC8Ah
  loc_00E2AC79: mov edx, var_E8
  loc_00E2AC7F: push 00000038h
  loc_00E2AC81: push 006E09F8h
  loc_00E2AC86: push edx
  loc_00E2AC87: push eax
  loc_00E2AC88: call ebx
  loc_00E2AC8A: lea ecx, var_28
  loc_00E2AC8D: call [00401258h] ; __vbaFreeStr
  loc_00E2AC93: lea eax, var_40
  loc_00E2AC96: lea ecx, var_3C
  loc_00E2AC99: push eax
  loc_00E2AC9A: lea edx, var_38
  loc_00E2AC9D: push ecx
  loc_00E2AC9E: lea eax, var_34
  loc_00E2ACA1: push edx
  loc_00E2ACA2: lea ecx, var_30
  loc_00E2ACA5: push eax
  loc_00E2ACA6: push ecx
  loc_00E2ACA7: push 00000005h
  loc_00E2ACA9: call [00401048h] ; __vbaFreeObjList
  loc_00E2ACAF: add esp, 00000018h
  loc_00E2ACB2: lea ecx, var_50
  loc_00E2ACB5: call [00401024h] ; __vbaFreeVar
  loc_00E2ACBB: mov edx, [esi]
  loc_00E2ACBD: push 006DCBD8h
  loc_00E2ACC2: push 00000000h
  loc_00E2ACC4: push 00000018h
  loc_00E2ACC6: push esi
  loc_00E2ACC7: call [edx+000004A8h]
  loc_00E2ACCD: push eax
  loc_00E2ACCE: lea eax, var_34
  loc_00E2ACD1: push eax
  loc_00E2ACD2: call edi
  loc_00E2ACD4: lea ecx, var_50
  loc_00E2ACD7: push eax
  loc_00E2ACD8: push ecx
  loc_00E2ACD9: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2ACDF: add esp, 00000010h
  loc_00E2ACE2: push eax
  loc_00E2ACE3: call [00401120h] ; __vbaCastObjVar
  loc_00E2ACE9: lea edx, var_38
  loc_00E2ACEC: push eax
  loc_00E2ACED: push edx
  loc_00E2ACEE: call edi
  loc_00E2ACF0: mov ecx, [eax]
  loc_00E2ACF2: lea edx, var_3C
  loc_00E2ACF5: push edx
  loc_00E2ACF6: push eax
  loc_00E2ACF7: mov var_D8, eax
  loc_00E2ACFD: call [ecx+00000054h]
  loc_00E2AD00: test eax, eax
  loc_00E2AD02: fnclex
  loc_00E2AD04: jge 00E2AD17h
  loc_00E2AD06: mov ecx, var_D8
  loc_00E2AD0C: push 00000054h
  loc_00E2AD0E: push 006DCBE8h
  loc_00E2AD13: push ecx
  loc_00E2AD14: push eax
  loc_00E2AD15: call ebx
  loc_00E2AD17: mov ecx, var_3C
  loc_00E2AD1A: mov eax, 00000002h
  loc_00E2AD1F: mov var_88, 00000006h
  loc_00E2AD29: mov var_90, eax
  loc_00E2AD2F: mov edx, [ecx]
  loc_00E2AD31: mov var_E0, ecx
  loc_00E2AD37: lea ecx, var_40
  loc_00E2AD3A: push ecx
  loc_00E2AD3B: sub esp, 00000010h
  loc_00E2AD3E: mov ecx, esp
  loc_00E2AD40: mov [ecx], eax
  loc_00E2AD42: mov eax, var_8C
  loc_00E2AD48: mov [ecx+00000004h], eax
  loc_00E2AD4B: mov eax, var_88
  loc_00E2AD51: mov [ecx+00000008h], eax
  loc_00E2AD54: mov eax, var_84
  loc_00E2AD5A: mov [ecx+0000000Ch], eax
  loc_00E2AD5D: mov ecx, var_3C
  loc_00E2AD60: push ecx
  loc_00E2AD61: call [edx+00000028h]
  loc_00E2AD64: test eax, eax
  loc_00E2AD66: fnclex
  loc_00E2AD68: jge 00E2AD7Bh
  loc_00E2AD6A: mov edx, var_E0
  loc_00E2AD70: push 00000028h
  loc_00E2AD72: push 006E09E8h
  loc_00E2AD77: push edx
  loc_00E2AD78: push eax
  loc_00E2AD79: call ebx
  loc_00E2AD7B: mov eax, var_40
  loc_00E2AD7E: mov ecx, [esi]
  loc_00E2AD80: push esi
  loc_00E2AD81: mov var_E8, eax
  loc_00E2AD87: call [ecx+000002FCh]
  loc_00E2AD8D: lea edx, var_30
  loc_00E2AD90: push eax
  loc_00E2AD91: push edx
  loc_00E2AD92: call edi
  loc_00E2AD94: mov ecx, [eax]
  loc_00E2AD96: lea edx, var_28
  loc_00E2AD99: push edx
  loc_00E2AD9A: push eax
  loc_00E2AD9B: mov var_D0, eax
  loc_00E2ADA1: call [ecx+000000A0h]
  loc_00E2ADA7: test eax, eax
  loc_00E2ADA9: fnclex
  loc_00E2ADAB: jge 00E2ADC1h
  loc_00E2ADAD: mov ecx, var_D0
  loc_00E2ADB3: push 000000A0h
  loc_00E2ADB8: push 006DCB70h
  loc_00E2ADBD: push ecx
  loc_00E2ADBE: push eax
  loc_00E2ADBF: call ebx
  loc_00E2ADC1: mov edx, var_28
  loc_00E2ADC4: push edx
  loc_00E2ADC5: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2ADCB: mov ecx, var_E8
  loc_00E2ADD1: mov eax, 00000005h
  loc_00E2ADD6: fstp real8 ptr var_98
  loc_00E2ADDC: mov var_A0, eax
  loc_00E2ADE2: mov edx, [ecx]
  loc_00E2ADE4: sub esp, 00000010h
  loc_00E2ADE7: mov ecx, esp
  loc_00E2ADE9: mov [ecx], eax
  loc_00E2ADEB: mov eax, var_9C
  loc_00E2ADF1: mov [ecx+00000004h], eax
  loc_00E2ADF4: mov eax, var_98
  loc_00E2ADFA: mov [ecx+00000008h], eax
  loc_00E2ADFD: mov eax, var_94
  loc_00E2AE03: mov [ecx+0000000Ch], eax
  loc_00E2AE06: mov ecx, var_E8
  loc_00E2AE0C: push ecx
  loc_00E2AE0D: call [edx+00000038h]
  loc_00E2AE10: test eax, eax
  loc_00E2AE12: fnclex
  loc_00E2AE14: jge 00E2AE27h
  loc_00E2AE16: mov edx, var_E8
  loc_00E2AE1C: push 00000038h
  loc_00E2AE1E: push 006E09F8h
  loc_00E2AE23: push edx
  loc_00E2AE24: push eax
  loc_00E2AE25: call ebx
  loc_00E2AE27: lea ecx, var_28
  loc_00E2AE2A: call [00401258h] ; __vbaFreeStr
  loc_00E2AE30: lea eax, var_40
  loc_00E2AE33: lea ecx, var_3C
  loc_00E2AE36: push eax
  loc_00E2AE37: lea edx, var_38
  loc_00E2AE3A: push ecx
  loc_00E2AE3B: lea eax, var_34
  loc_00E2AE3E: push edx
  loc_00E2AE3F: lea ecx, var_30
  loc_00E2AE42: push eax
  loc_00E2AE43: push ecx
  loc_00E2AE44: push 00000005h
  loc_00E2AE46: call [00401048h] ; __vbaFreeObjList
  loc_00E2AE4C: add esp, 00000018h
  loc_00E2AE4F: lea ecx, var_50
  loc_00E2AE52: call [00401024h] ; __vbaFreeVar
  loc_00E2AE58: mov edx, [esi]
  loc_00E2AE5A: push 006DCBD8h
  loc_00E2AE5F: push 00000000h
  loc_00E2AE61: push 00000018h
  loc_00E2AE63: push esi
  loc_00E2AE64: call [edx+000004A8h]
  loc_00E2AE6A: push eax
  loc_00E2AE6B: lea eax, var_34
  loc_00E2AE6E: push eax
  loc_00E2AE6F: call edi
  loc_00E2AE71: lea ecx, var_50
  loc_00E2AE74: push eax
  loc_00E2AE75: push ecx
  loc_00E2AE76: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2AE7C: add esp, 00000010h
  loc_00E2AE7F: push eax
  loc_00E2AE80: call [00401120h] ; __vbaCastObjVar
  loc_00E2AE86: lea edx, var_38
  loc_00E2AE89: push eax
  loc_00E2AE8A: push edx
  loc_00E2AE8B: call edi
  loc_00E2AE8D: mov ecx, [eax]
  loc_00E2AE8F: lea edx, var_3C
  loc_00E2AE92: push edx
  loc_00E2AE93: push eax
  loc_00E2AE94: mov var_D8, eax
  loc_00E2AE9A: call [ecx+00000054h]
  loc_00E2AE9D: test eax, eax
  loc_00E2AE9F: fnclex
  loc_00E2AEA1: jge 00E2AEB4h
  loc_00E2AEA3: mov ecx, var_D8
  loc_00E2AEA9: push 00000054h
  loc_00E2AEAB: push 006DCBE8h
  loc_00E2AEB0: push ecx
  loc_00E2AEB1: push eax
  loc_00E2AEB2: call ebx
  loc_00E2AEB4: mov ecx, var_3C
  loc_00E2AEB7: mov eax, 00000002h
  loc_00E2AEBC: mov var_88, 00000008h
  loc_00E2AEC6: mov var_90, eax
  loc_00E2AECC: mov edx, [ecx]
  loc_00E2AECE: mov var_E0, ecx
  loc_00E2AED4: lea ecx, var_40
  loc_00E2AED7: push ecx
  loc_00E2AED8: sub esp, 00000010h
  loc_00E2AEDB: mov ecx, esp
  loc_00E2AEDD: mov [ecx], eax
  loc_00E2AEDF: mov eax, var_8C
  loc_00E2AEE5: mov [ecx+00000004h], eax
  loc_00E2AEE8: mov eax, var_88
  loc_00E2AEEE: mov [ecx+00000008h], eax
  loc_00E2AEF1: mov eax, var_84
  loc_00E2AEF7: mov [ecx+0000000Ch], eax
  loc_00E2AEFA: mov ecx, var_3C
  loc_00E2AEFD: push ecx
  loc_00E2AEFE: call [edx+00000028h]
  loc_00E2AF01: test eax, eax
  loc_00E2AF03: fnclex
  loc_00E2AF05: jge 00E2AF18h
  loc_00E2AF07: mov edx, var_E0
  loc_00E2AF0D: push 00000028h
  loc_00E2AF0F: push 006E09E8h
  loc_00E2AF14: push edx
  loc_00E2AF15: push eax
  loc_00E2AF16: call ebx
  loc_00E2AF18: mov eax, var_40
  loc_00E2AF1B: mov ecx, [esi]
  loc_00E2AF1D: push esi
  loc_00E2AF1E: mov var_E8, eax
  loc_00E2AF24: call [ecx+00000368h]
  loc_00E2AF2A: lea edx, var_30
  loc_00E2AF2D: push eax
  loc_00E2AF2E: push edx
  loc_00E2AF2F: call edi
  loc_00E2AF31: mov ecx, [eax]
  loc_00E2AF33: lea edx, var_28
  loc_00E2AF36: push edx
  loc_00E2AF37: push eax
  loc_00E2AF38: mov var_D0, eax
  loc_00E2AF3E: call [ecx+00000050h]
  loc_00E2AF41: test eax, eax
  loc_00E2AF43: fnclex
  loc_00E2AF45: jge 00E2AF58h
  loc_00E2AF47: mov ecx, var_D0
  loc_00E2AF4D: push 00000050h
  loc_00E2AF4F: push 006E0410h
  loc_00E2AF54: push ecx
  loc_00E2AF55: push eax
  loc_00E2AF56: call ebx
  loc_00E2AF58: mov eax, var_28
  loc_00E2AF5B: mov edx, var_E8
  loc_00E2AF61: mov var_58, eax
  loc_00E2AF64: mov eax, 00000008h
  loc_00E2AF69: mov var_28, 00000000h
  loc_00E2AF70: mov var_60, eax
  loc_00E2AF73: mov ecx, [edx]
  loc_00E2AF75: sub esp, 00000010h
  loc_00E2AF78: mov edx, esp
  loc_00E2AF7A: mov [edx], eax
  loc_00E2AF7C: mov eax, var_5C
  loc_00E2AF7F: mov [edx+00000004h], eax
  loc_00E2AF82: mov eax, var_58
  loc_00E2AF85: mov [edx+00000008h], eax
  loc_00E2AF88: mov eax, var_54
  loc_00E2AF8B: mov [edx+0000000Ch], eax
  loc_00E2AF8E: mov edx, var_E8
  loc_00E2AF94: push edx
  loc_00E2AF95: call [ecx+00000038h]
  loc_00E2AF98: test eax, eax
  loc_00E2AF9A: fnclex
  loc_00E2AF9C: jge 00E2AFAFh
  loc_00E2AF9E: mov ecx, var_E8
  loc_00E2AFA4: push 00000038h
  loc_00E2AFA6: push 006E09F8h
  loc_00E2AFAB: push ecx
  loc_00E2AFAC: push eax
  loc_00E2AFAD: call ebx
  loc_00E2AFAF: lea edx, var_40
  loc_00E2AFB2: lea eax, var_3C
  loc_00E2AFB5: push edx
  loc_00E2AFB6: lea ecx, var_38
  loc_00E2AFB9: push eax
  loc_00E2AFBA: lea edx, var_34
  loc_00E2AFBD: push ecx
  loc_00E2AFBE: lea eax, var_30
  loc_00E2AFC1: push edx
  loc_00E2AFC2: push eax
  loc_00E2AFC3: push 00000005h
  loc_00E2AFC5: call [00401048h] ; __vbaFreeObjList
  loc_00E2AFCB: lea ecx, var_60
  loc_00E2AFCE: lea edx, var_50
  loc_00E2AFD1: push ecx
  loc_00E2AFD2: push edx
  loc_00E2AFD3: push 00000002h
  loc_00E2AFD5: call [00401038h] ; __vbaFreeVarList
  loc_00E2AFDB: mov eax, [esi]
  loc_00E2AFDD: add esp, 00000024h
  loc_00E2AFE0: push 006DCBD8h
  loc_00E2AFE5: push 00000000h
  loc_00E2AFE7: push 00000018h
  loc_00E2AFE9: push esi
  loc_00E2AFEA: call [eax+000004A8h]
  loc_00E2AFF0: lea ecx, var_34
  loc_00E2AFF3: push eax
  loc_00E2AFF4: push ecx
  loc_00E2AFF5: call edi
  loc_00E2AFF7: lea edx, var_50
  loc_00E2AFFA: push eax
  loc_00E2AFFB: push edx
  loc_00E2AFFC: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2B002: add esp, 00000010h
  loc_00E2B005: push eax
  loc_00E2B006: call [00401120h] ; __vbaCastObjVar
  loc_00E2B00C: push eax
  loc_00E2B00D: lea eax, var_38
  loc_00E2B010: push eax
  loc_00E2B011: call edi
  loc_00E2B013: mov ecx, [eax]
  loc_00E2B015: lea edx, var_3C
  loc_00E2B018: push edx
  loc_00E2B019: push eax
  loc_00E2B01A: mov var_D8, eax
  loc_00E2B020: call [ecx+00000054h]
  loc_00E2B023: test eax, eax
  loc_00E2B025: fnclex
  loc_00E2B027: jge 00E2B03Ah
  loc_00E2B029: mov ecx, var_D8
  loc_00E2B02F: push 00000054h
  loc_00E2B031: push 006DCBE8h
  loc_00E2B036: push ecx
  loc_00E2B037: push eax
  loc_00E2B038: call ebx
  loc_00E2B03A: mov ecx, var_3C
  loc_00E2B03D: mov eax, 00000002h
  loc_00E2B042: mov var_88, 0000000Bh
  loc_00E2B04C: mov var_90, eax
  loc_00E2B052: mov edx, [ecx]
  loc_00E2B054: mov var_E0, ecx
  loc_00E2B05A: lea ecx, var_40
  loc_00E2B05D: push ecx
  loc_00E2B05E: sub esp, 00000010h
  loc_00E2B061: mov ecx, esp
  loc_00E2B063: mov [ecx], eax
  loc_00E2B065: mov eax, var_8C
  loc_00E2B06B: mov [ecx+00000004h], eax
  loc_00E2B06E: mov eax, var_88
  loc_00E2B074: mov [ecx+00000008h], eax
  loc_00E2B077: mov eax, var_84
  loc_00E2B07D: mov [ecx+0000000Ch], eax
  loc_00E2B080: mov ecx, var_3C
  loc_00E2B083: push ecx
  loc_00E2B084: call [edx+00000028h]
  loc_00E2B087: test eax, eax
  loc_00E2B089: fnclex
  loc_00E2B08B: jge 00E2B09Eh
  loc_00E2B08D: mov edx, var_E0
  loc_00E2B093: push 00000028h
  loc_00E2B095: push 006E09E8h
  loc_00E2B09A: push edx
  loc_00E2B09B: push eax
  loc_00E2B09C: call ebx
  loc_00E2B09E: mov eax, var_40
  loc_00E2B0A1: mov ecx, [esi]
  loc_00E2B0A3: push esi
  loc_00E2B0A4: mov var_E8, eax
  loc_00E2B0AA: call [ecx+00000304h]
  loc_00E2B0B0: lea edx, var_30
  loc_00E2B0B3: push eax
  loc_00E2B0B4: push edx
  loc_00E2B0B5: call edi
  loc_00E2B0B7: mov ecx, [eax]
  loc_00E2B0B9: lea edx, var_28
  loc_00E2B0BC: push edx
  loc_00E2B0BD: push eax
  loc_00E2B0BE: mov var_D0, eax
  loc_00E2B0C4: call [ecx+000000A0h]
  loc_00E2B0CA: test eax, eax
  loc_00E2B0CC: fnclex
  loc_00E2B0CE: jge 00E2B0E4h
  loc_00E2B0D0: mov ecx, var_D0
  loc_00E2B0D6: push 000000A0h
  loc_00E2B0DB: push 006DCB70h
  loc_00E2B0E0: push ecx
  loc_00E2B0E1: push eax
  loc_00E2B0E2: call ebx
  loc_00E2B0E4: mov eax, var_28
  loc_00E2B0E7: mov edx, var_E8
  loc_00E2B0ED: mov var_58, eax
  loc_00E2B0F0: mov eax, 00000008h
  loc_00E2B0F5: mov var_28, 00000000h
  loc_00E2B0FC: mov var_60, eax
  loc_00E2B0FF: mov ecx, [edx]
  loc_00E2B101: sub esp, 00000010h
  loc_00E2B104: mov edx, esp
  loc_00E2B106: mov [edx], eax
  loc_00E2B108: mov eax, var_5C
  loc_00E2B10B: mov [edx+00000004h], eax
  loc_00E2B10E: mov eax, var_58
  loc_00E2B111: mov [edx+00000008h], eax
  loc_00E2B114: mov eax, var_54
  loc_00E2B117: mov [edx+0000000Ch], eax
  loc_00E2B11A: mov edx, var_E8
  loc_00E2B120: push edx
  loc_00E2B121: call [ecx+00000038h]
  loc_00E2B124: test eax, eax
  loc_00E2B126: fnclex
  loc_00E2B128: jge 00E2B13Bh
  loc_00E2B12A: mov ecx, var_E8
  loc_00E2B130: push 00000038h
  loc_00E2B132: push 006E09F8h
  loc_00E2B137: push ecx
  loc_00E2B138: push eax
  loc_00E2B139: call ebx
  loc_00E2B13B: lea edx, var_40
  loc_00E2B13E: lea eax, var_3C
  loc_00E2B141: push edx
  loc_00E2B142: lea ecx, var_38
  loc_00E2B145: push eax
  loc_00E2B146: lea edx, var_34
  loc_00E2B149: push ecx
  loc_00E2B14A: lea eax, var_30
  loc_00E2B14D: push edx
  loc_00E2B14E: push eax
  loc_00E2B14F: push 00000005h
  loc_00E2B151: call [00401048h] ; __vbaFreeObjList
  loc_00E2B157: lea ecx, var_60
  loc_00E2B15A: lea edx, var_50
  loc_00E2B15D: push ecx
  loc_00E2B15E: push edx
  loc_00E2B15F: push 00000002h
  loc_00E2B161: call [00401038h] ; __vbaFreeVarList
  loc_00E2B167: mov eax, [esi]
  loc_00E2B169: add esp, 00000024h
  loc_00E2B16C: push 006DCBD8h
  loc_00E2B171: push 00000000h
  loc_00E2B173: push 00000018h
  loc_00E2B175: push esi
  loc_00E2B176: call [eax+000004A8h]
  loc_00E2B17C: lea ecx, var_34
  loc_00E2B17F: push eax
  loc_00E2B180: push ecx
  loc_00E2B181: call edi
  loc_00E2B183: lea edx, var_50
  loc_00E2B186: push eax
  loc_00E2B187: push edx
  loc_00E2B188: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2B18E: add esp, 00000010h
  loc_00E2B191: push eax
  loc_00E2B192: call [00401120h] ; __vbaCastObjVar
  loc_00E2B198: push eax
  loc_00E2B199: lea eax, var_38
  loc_00E2B19C: push eax
  loc_00E2B19D: call edi
  loc_00E2B19F: mov ecx, [eax]
  loc_00E2B1A1: lea edx, var_3C
  loc_00E2B1A4: push edx
  loc_00E2B1A5: push eax
  loc_00E2B1A6: mov var_D8, eax
  loc_00E2B1AC: call [ecx+00000054h]
  loc_00E2B1AF: test eax, eax
  loc_00E2B1B1: fnclex
  loc_00E2B1B3: jge 00E2B1C6h
  loc_00E2B1B5: mov ecx, var_D8
  loc_00E2B1BB: push 00000054h
  loc_00E2B1BD: push 006DCBE8h
  loc_00E2B1C2: push ecx
  loc_00E2B1C3: push eax
  loc_00E2B1C4: call ebx
  loc_00E2B1C6: mov ecx, var_3C
  loc_00E2B1C9: mov eax, 00000002h
  loc_00E2B1CE: mov var_88, 0000000Ch
  loc_00E2B1D8: mov var_90, eax
  loc_00E2B1DE: mov edx, [ecx]
  loc_00E2B1E0: mov var_E0, ecx
  loc_00E2B1E6: lea ecx, var_40
  loc_00E2B1E9: push ecx
  loc_00E2B1EA: sub esp, 00000010h
  loc_00E2B1ED: mov ecx, esp
  loc_00E2B1EF: mov [ecx], eax
  loc_00E2B1F1: mov eax, var_8C
  loc_00E2B1F7: mov [ecx+00000004h], eax
  loc_00E2B1FA: mov eax, var_88
  loc_00E2B200: mov [ecx+00000008h], eax
  loc_00E2B203: mov eax, var_84
  loc_00E2B209: mov [ecx+0000000Ch], eax
  loc_00E2B20C: mov ecx, var_3C
  loc_00E2B20F: push ecx
  loc_00E2B210: call [edx+00000028h]
  loc_00E2B213: test eax, eax
  loc_00E2B215: fnclex
  loc_00E2B217: jge 00E2B22Ah
  loc_00E2B219: mov edx, var_E0
  loc_00E2B21F: push 00000028h
  loc_00E2B221: push 006E09E8h
  loc_00E2B226: push edx
  loc_00E2B227: push eax
  loc_00E2B228: call ebx
  loc_00E2B22A: mov eax, var_40
  loc_00E2B22D: mov ecx, [esi]
  loc_00E2B22F: push esi
  loc_00E2B230: mov var_E8, eax
  loc_00E2B236: call [ecx+00000354h]
  loc_00E2B23C: lea edx, var_30
  loc_00E2B23F: push eax
  loc_00E2B240: push edx
  loc_00E2B241: call edi
  loc_00E2B243: mov ecx, [eax]
  loc_00E2B245: lea edx, var_28
  loc_00E2B248: push edx
  loc_00E2B249: push eax
  loc_00E2B24A: mov var_D0, eax
  loc_00E2B250: call [ecx+00000050h]
  loc_00E2B253: test eax, eax
  loc_00E2B255: fnclex
  loc_00E2B257: jge 00E2B26Ah
  loc_00E2B259: mov ecx, var_D0
  loc_00E2B25F: push 00000050h
  loc_00E2B261: push 006E0410h
  loc_00E2B266: push ecx
  loc_00E2B267: push eax
  loc_00E2B268: call ebx
  loc_00E2B26A: mov eax, var_28
  loc_00E2B26D: mov edx, var_E8
  loc_00E2B273: mov var_58, eax
  loc_00E2B276: mov eax, 00000008h
  loc_00E2B27B: mov var_28, 00000000h
  loc_00E2B282: mov var_60, eax
  loc_00E2B285: mov ecx, [edx]
  loc_00E2B287: sub esp, 00000010h
  loc_00E2B28A: mov edx, esp
  loc_00E2B28C: mov [edx], eax
  loc_00E2B28E: mov eax, var_5C
  loc_00E2B291: mov [edx+00000004h], eax
  loc_00E2B294: mov eax, var_58
  loc_00E2B297: mov [edx+00000008h], eax
  loc_00E2B29A: mov eax, var_54
  loc_00E2B29D: mov [edx+0000000Ch], eax
  loc_00E2B2A0: mov edx, var_E8
  loc_00E2B2A6: push edx
  loc_00E2B2A7: call [ecx+00000038h]
  loc_00E2B2AA: test eax, eax
  loc_00E2B2AC: fnclex
  loc_00E2B2AE: jge 00E2B2C1h
  loc_00E2B2B0: mov ecx, var_E8
  loc_00E2B2B6: push 00000038h
  loc_00E2B2B8: push 006E09F8h
  loc_00E2B2BD: push ecx
  loc_00E2B2BE: push eax
  loc_00E2B2BF: call ebx
  loc_00E2B2C1: lea edx, var_40
  loc_00E2B2C4: lea eax, var_3C
  loc_00E2B2C7: push edx
  loc_00E2B2C8: lea ecx, var_38
  loc_00E2B2CB: push eax
  loc_00E2B2CC: lea edx, var_34
  loc_00E2B2CF: push ecx
  loc_00E2B2D0: lea eax, var_30
  loc_00E2B2D3: push edx
  loc_00E2B2D4: push eax
  loc_00E2B2D5: push 00000005h
  loc_00E2B2D7: call [00401048h] ; __vbaFreeObjList
  loc_00E2B2DD: lea ecx, var_60
  loc_00E2B2E0: lea edx, var_50
  loc_00E2B2E3: push ecx
  loc_00E2B2E4: push edx
  loc_00E2B2E5: push 00000002h
  loc_00E2B2E7: call [00401038h] ; __vbaFreeVarList
  loc_00E2B2ED: mov eax, [esi]
  loc_00E2B2EF: add esp, 00000024h
  loc_00E2B2F2: push 006DCBD8h
  loc_00E2B2F7: push 00000000h
  loc_00E2B2F9: push 00000018h
  loc_00E2B2FB: push esi
  loc_00E2B2FC: call [eax+000004A8h]
  loc_00E2B302: lea ecx, var_34
  loc_00E2B305: push eax
  loc_00E2B306: push ecx
  loc_00E2B307: call edi
  loc_00E2B309: lea edx, var_50
  loc_00E2B30C: push eax
  loc_00E2B30D: push edx
  loc_00E2B30E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2B314: add esp, 00000010h
  loc_00E2B317: push eax
  loc_00E2B318: call [00401120h] ; __vbaCastObjVar
  loc_00E2B31E: push eax
  loc_00E2B31F: lea eax, var_38
  loc_00E2B322: push eax
  loc_00E2B323: call edi
  loc_00E2B325: mov ecx, [eax]
  loc_00E2B327: lea edx, var_3C
  loc_00E2B32A: push edx
  loc_00E2B32B: push eax
  loc_00E2B32C: mov var_D8, eax
  loc_00E2B332: call [ecx+00000054h]
  loc_00E2B335: test eax, eax
  loc_00E2B337: fnclex
  loc_00E2B339: jge 00E2B34Ch
  loc_00E2B33B: mov ecx, var_D8
  loc_00E2B341: push 00000054h
  loc_00E2B343: push 006DCBE8h
  loc_00E2B348: push ecx
  loc_00E2B349: push eax
  loc_00E2B34A: call ebx
  loc_00E2B34C: mov ecx, var_3C
  loc_00E2B34F: mov eax, 00000002h
  loc_00E2B354: mov var_88, 0000000Dh
  loc_00E2B35E: mov var_90, eax
  loc_00E2B364: mov edx, [ecx]
  loc_00E2B366: mov var_E0, ecx
  loc_00E2B36C: lea ecx, var_40
  loc_00E2B36F: push ecx
  loc_00E2B370: sub esp, 00000010h
  loc_00E2B373: mov ecx, esp
  loc_00E2B375: mov [ecx], eax
  loc_00E2B377: mov eax, var_8C
  loc_00E2B37D: mov [ecx+00000004h], eax
  loc_00E2B380: mov eax, var_88
  loc_00E2B386: mov [ecx+00000008h], eax
  loc_00E2B389: mov eax, var_84
  loc_00E2B38F: mov [ecx+0000000Ch], eax
  loc_00E2B392: mov ecx, var_3C
  loc_00E2B395: push ecx
  loc_00E2B396: call [edx+00000028h]
  loc_00E2B399: test eax, eax
  loc_00E2B39B: fnclex
  loc_00E2B39D: jge 00E2B3B0h
  loc_00E2B39F: mov edx, var_E0
  loc_00E2B3A5: push 00000028h
  loc_00E2B3A7: push 006E09E8h
  loc_00E2B3AC: push edx
  loc_00E2B3AD: push eax
  loc_00E2B3AE: call ebx
  loc_00E2B3B0: mov eax, var_40
  loc_00E2B3B3: mov ecx, [esi]
  loc_00E2B3B5: push esi
  loc_00E2B3B6: mov var_E8, eax
  loc_00E2B3BC: call [ecx+000003B0h]
  loc_00E2B3C2: lea edx, var_30
  loc_00E2B3C5: push eax
  loc_00E2B3C6: push edx
  loc_00E2B3C7: call edi
  loc_00E2B3C9: mov ecx, [eax]
  loc_00E2B3CB: lea edx, var_28
  loc_00E2B3CE: push edx
  loc_00E2B3CF: push eax
  loc_00E2B3D0: mov var_D0, eax
  loc_00E2B3D6: call [ecx+00000050h]
  loc_00E2B3D9: test eax, eax
  loc_00E2B3DB: fnclex
  loc_00E2B3DD: jge 00E2B3F0h
  loc_00E2B3DF: mov ecx, var_D0
  loc_00E2B3E5: push 00000050h
  loc_00E2B3E7: push 006E0410h
  loc_00E2B3EC: push ecx
  loc_00E2B3ED: push eax
  loc_00E2B3EE: call ebx
  loc_00E2B3F0: mov eax, var_28
  loc_00E2B3F3: mov edx, var_E8
  loc_00E2B3F9: mov var_58, eax
  loc_00E2B3FC: mov eax, 00000008h
  loc_00E2B401: mov var_28, 00000000h
  loc_00E2B408: mov var_60, eax
  loc_00E2B40B: mov ecx, [edx]
  loc_00E2B40D: sub esp, 00000010h
  loc_00E2B410: mov edx, esp
  loc_00E2B412: mov [edx], eax
  loc_00E2B414: mov eax, var_5C
  loc_00E2B417: mov [edx+00000004h], eax
  loc_00E2B41A: mov eax, var_58
  loc_00E2B41D: mov [edx+00000008h], eax
  loc_00E2B420: mov eax, var_54
  loc_00E2B423: mov [edx+0000000Ch], eax
  loc_00E2B426: mov edx, var_E8
  loc_00E2B42C: push edx
  loc_00E2B42D: call [ecx+00000038h]
  loc_00E2B430: test eax, eax
  loc_00E2B432: fnclex
  loc_00E2B434: jge 00E2B447h
  loc_00E2B436: mov ecx, var_E8
  loc_00E2B43C: push 00000038h
  loc_00E2B43E: push 006E09F8h
  loc_00E2B443: push ecx
  loc_00E2B444: push eax
  loc_00E2B445: call ebx
  loc_00E2B447: lea edx, var_40
  loc_00E2B44A: lea eax, var_3C
  loc_00E2B44D: push edx
  loc_00E2B44E: lea ecx, var_38
  loc_00E2B451: push eax
  loc_00E2B452: lea edx, var_34
  loc_00E2B455: push ecx
  loc_00E2B456: lea eax, var_30
  loc_00E2B459: push edx
  loc_00E2B45A: push eax
  loc_00E2B45B: push 00000005h
  loc_00E2B45D: call [00401048h] ; __vbaFreeObjList
  loc_00E2B463: lea ecx, var_60
  loc_00E2B466: lea edx, var_50
  loc_00E2B469: push ecx
  loc_00E2B46A: push edx
  loc_00E2B46B: push 00000002h
  loc_00E2B46D: call [00401038h] ; __vbaFreeVarList
  loc_00E2B473: mov eax, [esi]
  loc_00E2B475: add esp, 00000024h
  loc_00E2B478: push esi
  loc_00E2B479: call [eax+0000037Ch]
  loc_00E2B47F: lea ecx, var_30
  loc_00E2B482: push eax
  loc_00E2B483: push ecx
  loc_00E2B484: call edi
  loc_00E2B486: mov edx, [eax]
  loc_00E2B488: lea ecx, var_C4
  loc_00E2B48E: push ecx
  loc_00E2B48F: push eax
  loc_00E2B490: mov var_D0, eax
  loc_00E2B496: call [edx+00000098h]
  loc_00E2B49C: test eax, eax
  loc_00E2B49E: fnclex
  loc_00E2B4A0: jge 00E2B4B6h
  loc_00E2B4A2: mov edx, var_D0
  loc_00E2B4A8: push 00000098h
  loc_00E2B4AD: push 006DCAD0h
  loc_00E2B4B2: push edx
  loc_00E2B4B3: push eax
  loc_00E2B4B4: call ebx
  loc_00E2B4B6: xor eax, eax
  loc_00E2B4B8: cmp var_C4, FFFFFFh
  loc_00E2B4C0: lea ecx, var_30
  loc_00E2B4C3: setz al
  loc_00E2B4C6: neg eax
  loc_00E2B4C8: mov var_D8, ax
  loc_00E2B4CF: call [00401254h] ; __vbaFreeObj
  loc_00E2B4D5: cmp var_D8, 0000h
  loc_00E2B4DD: jz 00E2B66Eh
  loc_00E2B4E3: mov ecx, [esi]
  loc_00E2B4E5: push 006DCBD8h
  loc_00E2B4EA: push 00000000h
  loc_00E2B4EC: push 00000018h
  loc_00E2B4EE: push esi
  loc_00E2B4EF: call [ecx+000004A8h]
  loc_00E2B4F5: lea edx, var_34
  loc_00E2B4F8: push eax
  loc_00E2B4F9: push edx
  loc_00E2B4FA: call edi
  loc_00E2B4FC: push eax
  loc_00E2B4FD: lea eax, var_50
  loc_00E2B500: push eax
  loc_00E2B501: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2B507: add esp, 00000010h
  loc_00E2B50A: push eax
  loc_00E2B50B: call [00401120h] ; __vbaCastObjVar
  loc_00E2B511: lea ecx, var_38
  loc_00E2B514: push eax
  loc_00E2B515: push ecx
  loc_00E2B516: call edi
  loc_00E2B518: mov edx, [eax]
  loc_00E2B51A: lea ecx, var_3C
  loc_00E2B51D: push ecx
  loc_00E2B51E: push eax
  loc_00E2B51F: mov var_D8, eax
  loc_00E2B525: call [edx+00000054h]
  loc_00E2B528: test eax, eax
  loc_00E2B52A: fnclex
  loc_00E2B52C: jge 00E2B53Fh
  loc_00E2B52E: mov edx, var_D8
  loc_00E2B534: push 00000054h
  loc_00E2B536: push 006DCBE8h
  loc_00E2B53B: push edx
  loc_00E2B53C: push eax
  loc_00E2B53D: call ebx
  loc_00E2B53F: lea edx, var_40
  loc_00E2B542: mov eax, 00000002h
  loc_00E2B547: push edx
  loc_00E2B548: mov ecx, var_3C
  loc_00E2B54B: sub esp, 00000010h
  loc_00E2B54E: mov var_90, eax
  loc_00E2B554: mov edx, esp
  loc_00E2B556: mov var_88, 00000009h
  loc_00E2B560: mov var_E0, ecx
  loc_00E2B566: mov ecx, [ecx]
  loc_00E2B568: mov [edx], eax
  loc_00E2B56A: mov eax, var_8C
  loc_00E2B570: mov [edx+00000004h], eax
  loc_00E2B573: mov eax, var_88
  loc_00E2B579: mov [edx+00000008h], eax
  loc_00E2B57C: mov eax, var_84
  loc_00E2B582: mov [edx+0000000Ch], eax
  loc_00E2B585: mov edx, var_3C
  loc_00E2B588: push edx
  loc_00E2B589: call [ecx+00000028h]
  loc_00E2B58C: test eax, eax
  loc_00E2B58E: fnclex
  loc_00E2B590: jge 00E2B5A3h
  loc_00E2B592: mov ecx, var_E0
  loc_00E2B598: push 00000028h
  loc_00E2B59A: push 006E09E8h
  loc_00E2B59F: push ecx
  loc_00E2B5A0: push eax
  loc_00E2B5A1: call ebx
  loc_00E2B5A3: mov edx, var_40
  loc_00E2B5A6: mov eax, [esi]
  loc_00E2B5A8: push esi
  loc_00E2B5A9: mov var_E8, edx
  loc_00E2B5AF: call [eax+00000388h]
  loc_00E2B5B5: lea ecx, var_30
  loc_00E2B5B8: push eax
  loc_00E2B5B9: push ecx
  loc_00E2B5BA: call edi
  loc_00E2B5BC: mov edx, [eax]
  loc_00E2B5BE: lea ecx, var_28
  loc_00E2B5C1: push ecx
  loc_00E2B5C2: push eax
  loc_00E2B5C3: mov var_D0, eax
  loc_00E2B5C9: call [edx+00000050h]
  loc_00E2B5CC: test eax, eax
  loc_00E2B5CE: fnclex
  loc_00E2B5D0: jge 00E2B5E3h
  loc_00E2B5D2: mov edx, var_D0
  loc_00E2B5D8: push 00000050h
  loc_00E2B5DA: push 006E0410h
  loc_00E2B5DF: push edx
  loc_00E2B5E0: push eax
  loc_00E2B5E1: call ebx
  loc_00E2B5E3: mov eax, var_28
  loc_00E2B5E6: mov ecx, var_E8
  loc_00E2B5EC: mov var_58, eax
  loc_00E2B5EF: mov eax, 00000008h
  loc_00E2B5F4: mov var_28, 00000000h
  loc_00E2B5FB: mov var_60, eax
  loc_00E2B5FE: mov edx, [ecx]
  loc_00E2B600: sub esp, 00000010h
  loc_00E2B603: mov ecx, esp
  loc_00E2B605: mov [ecx], eax
  loc_00E2B607: mov eax, var_5C
  loc_00E2B60A: mov [ecx+00000004h], eax
  loc_00E2B60D: mov eax, var_58
  loc_00E2B610: mov [ecx+00000008h], eax
  loc_00E2B613: mov eax, var_54
  loc_00E2B616: mov [ecx+0000000Ch], eax
  loc_00E2B619: mov ecx, var_E8
  loc_00E2B61F: push ecx
  loc_00E2B620: call [edx+00000038h]
  loc_00E2B623: test eax, eax
  loc_00E2B625: fnclex
  loc_00E2B627: jge 00E2B63Ah
  loc_00E2B629: mov edx, var_E8
  loc_00E2B62F: push 00000038h
  loc_00E2B631: push 006E09F8h
  loc_00E2B636: push edx
  loc_00E2B637: push eax
  loc_00E2B638: call ebx
  loc_00E2B63A: lea eax, var_40
  loc_00E2B63D: lea ecx, var_3C
  loc_00E2B640: push eax
  loc_00E2B641: lea edx, var_38
  loc_00E2B644: push ecx
  loc_00E2B645: lea eax, var_34
  loc_00E2B648: push edx
  loc_00E2B649: lea ecx, var_30
  loc_00E2B64C: push eax
  loc_00E2B64D: push ecx
  loc_00E2B64E: push 00000005h
  loc_00E2B650: call [00401048h] ; __vbaFreeObjList
  loc_00E2B656: lea edx, var_60
  loc_00E2B659: lea eax, var_50
  loc_00E2B65C: push edx
  loc_00E2B65D: push eax
  loc_00E2B65E: push 00000002h
  loc_00E2B660: call [00401038h] ; __vbaFreeVarList
  loc_00E2B666: add esp, 00000024h
  loc_00E2B669: jmp 00E2B9ECh
  loc_00E2B66E: mov ecx, [esi]
  loc_00E2B670: push esi
  loc_00E2B671: call [ecx+0000037Ch]
  loc_00E2B677: lea edx, var_30
  loc_00E2B67A: push eax
  loc_00E2B67B: push edx
  loc_00E2B67C: call edi
  loc_00E2B67E: mov ecx, [eax]
  loc_00E2B680: lea edx, var_C4
  loc_00E2B686: push edx
  loc_00E2B687: push eax
  loc_00E2B688: mov var_D0, eax
  loc_00E2B68E: call [ecx+00000098h]
  loc_00E2B694: test eax, eax
  loc_00E2B696: fnclex
  loc_00E2B698: jge 00E2B6AEh
  loc_00E2B69A: mov ecx, var_D0
  loc_00E2B6A0: push 00000098h
  loc_00E2B6A5: push 006DCAD0h
  loc_00E2B6AA: push ecx
  loc_00E2B6AB: push eax
  loc_00E2B6AC: call ebx
  loc_00E2B6AE: xor edx, edx
  loc_00E2B6B0: lea ecx, var_30
  loc_00E2B6B3: cmp var_C4, dx
  loc_00E2B6BA: setz dl
  loc_00E2B6BD: neg edx
  loc_00E2B6BF: mov var_D8, dx
  loc_00E2B6C6: call [00401254h] ; __vbaFreeObj
  loc_00E2B6CC: cmp var_D8, 0000h
  loc_00E2B6D4: jz 00E2B9ECh
  loc_00E2B6DA: mov eax, [esi]
  loc_00E2B6DC: push esi
  loc_00E2B6DD: call [eax+000003B0h]
  loc_00E2B6E3: lea ecx, var_30
  loc_00E2B6E6: push eax
  loc_00E2B6E7: push ecx
  loc_00E2B6E8: call edi
  loc_00E2B6EA: mov edx, [eax]
  loc_00E2B6EC: lea ecx, var_28
  loc_00E2B6EF: push ecx
  loc_00E2B6F0: push eax
  loc_00E2B6F1: mov var_D0, eax
  loc_00E2B6F7: call [edx+00000050h]
  loc_00E2B6FA: test eax, eax
  loc_00E2B6FC: fnclex
  loc_00E2B6FE: jge 00E2B711h
  loc_00E2B700: mov edx, var_D0
  loc_00E2B706: push 00000050h
  loc_00E2B708: push 006E0410h
  loc_00E2B70D: push edx
  loc_00E2B70E: push eax
  loc_00E2B70F: call ebx
  loc_00E2B711: mov eax, var_28
  loc_00E2B714: push eax
  loc_00E2B715: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2B71B: call [004010D8h] ; __vbaFpR8
  loc_00E2B721: fcomp real8 ptr [004022E8h]
  loc_00E2B727: mov var_18C, 00000001h
  loc_00E2B731: fnstsw ax
  loc_00E2B733: test ah, 41h
  loc_00E2B736: jnz 00E2B742h
  loc_00E2B738: mov var_18C, 00000000h
  loc_00E2B742: lea ecx, var_28
  loc_00E2B745: call [00401258h] ; __vbaFreeStr
  loc_00E2B74B: lea ecx, var_30
  loc_00E2B74E: call [00401254h] ; __vbaFreeObj
  loc_00E2B754: mov eax, var_18C
  loc_00E2B75A: mov ecx, [esi]
  loc_00E2B75C: neg eax
  loc_00E2B75E: push 006DCBD8h
  loc_00E2B763: push 00000000h
  loc_00E2B765: test ax, ax
  loc_00E2B768: push 00000018h
  loc_00E2B76A: push esi
  loc_00E2B76B: jz 00E2B8B4h
  loc_00E2B771: call [ecx+000004A8h]
  loc_00E2B777: lea edx, var_34
  loc_00E2B77A: push eax
  loc_00E2B77B: push edx
  loc_00E2B77C: call edi
  loc_00E2B77E: push eax
  loc_00E2B77F: lea eax, var_50
  loc_00E2B782: push eax
  loc_00E2B783: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2B789: add esp, 00000010h
  loc_00E2B78C: push eax
  loc_00E2B78D: call [00401120h] ; __vbaCastObjVar
  loc_00E2B793: lea ecx, var_38
  loc_00E2B796: push eax
  loc_00E2B797: push ecx
  loc_00E2B798: call edi
  loc_00E2B79A: mov edx, [eax]
  loc_00E2B79C: lea ecx, var_3C
  loc_00E2B79F: push ecx
  loc_00E2B7A0: push eax
  loc_00E2B7A1: mov var_D8, eax
  loc_00E2B7A7: call [edx+00000054h]
  loc_00E2B7AA: test eax, eax
  loc_00E2B7AC: fnclex
  loc_00E2B7AE: jge 00E2B7C1h
  loc_00E2B7B0: mov edx, var_D8
  loc_00E2B7B6: push 00000054h
  loc_00E2B7B8: push 006DCBE8h
  loc_00E2B7BD: push edx
  loc_00E2B7BE: push eax
  loc_00E2B7BF: call ebx
  loc_00E2B7C1: lea edx, var_40
  loc_00E2B7C4: mov eax, 00000002h
  loc_00E2B7C9: push edx
  loc_00E2B7CA: mov ecx, var_3C
  loc_00E2B7CD: sub esp, 00000010h
  loc_00E2B7D0: mov var_90, eax
  loc_00E2B7D6: mov edx, esp
  loc_00E2B7D8: mov var_88, 00000009h
  loc_00E2B7E2: mov var_E0, ecx
  loc_00E2B7E8: mov ecx, [ecx]
  loc_00E2B7EA: mov [edx], eax
  loc_00E2B7EC: mov eax, var_8C
  loc_00E2B7F2: mov [edx+00000004h], eax
  loc_00E2B7F5: mov eax, var_88
  loc_00E2B7FB: mov [edx+00000008h], eax
  loc_00E2B7FE: mov eax, var_84
  loc_00E2B804: mov [edx+0000000Ch], eax
  loc_00E2B807: mov edx, var_3C
  loc_00E2B80A: push edx
  loc_00E2B80B: call [ecx+00000028h]
  loc_00E2B80E: test eax, eax
  loc_00E2B810: fnclex
  loc_00E2B812: jge 00E2B825h
  loc_00E2B814: mov ecx, var_E0
  loc_00E2B81A: push 00000028h
  loc_00E2B81C: push 006E09E8h
  loc_00E2B821: push ecx
  loc_00E2B822: push eax
  loc_00E2B823: call ebx
  loc_00E2B825: mov edx, var_40
  loc_00E2B828: mov eax, [esi]
  loc_00E2B82A: push esi
  loc_00E2B82B: mov var_E8, edx
  loc_00E2B831: call [eax+00000388h]
  loc_00E2B837: lea ecx, var_30
  loc_00E2B83A: push eax
  loc_00E2B83B: push ecx
  loc_00E2B83C: call edi
  loc_00E2B83E: mov edx, [eax]
  loc_00E2B840: lea ecx, var_28
  loc_00E2B843: push ecx
  loc_00E2B844: push eax
  loc_00E2B845: mov var_D0, eax
  loc_00E2B84B: call [edx+00000050h]
  loc_00E2B84E: test eax, eax
  loc_00E2B850: fnclex
  loc_00E2B852: jge 00E2B865h
  loc_00E2B854: mov edx, var_D0
  loc_00E2B85A: push 00000050h
  loc_00E2B85C: push 006E0410h
  loc_00E2B861: push edx
  loc_00E2B862: push eax
  loc_00E2B863: call ebx
  loc_00E2B865: mov eax, var_28
  loc_00E2B868: mov ecx, var_E8
  loc_00E2B86E: mov var_58, eax
  loc_00E2B871: mov eax, 00000008h
  loc_00E2B876: mov var_28, 00000000h
  loc_00E2B87D: mov var_60, eax
  loc_00E2B880: mov edx, [ecx]
  loc_00E2B882: sub esp, 00000010h
  loc_00E2B885: mov ecx, esp
  loc_00E2B887: mov [ecx], eax
  loc_00E2B889: mov eax, var_5C
  loc_00E2B88C: mov [ecx+00000004h], eax
  loc_00E2B88F: mov eax, var_58
  loc_00E2B892: mov [ecx+00000008h], eax
  loc_00E2B895: mov eax, var_54
  loc_00E2B898: mov [ecx+0000000Ch], eax
  loc_00E2B89B: mov ecx, var_E8
  loc_00E2B8A1: push ecx
  loc_00E2B8A2: call [edx+00000038h]
  loc_00E2B8A5: test eax, eax
  loc_00E2B8A7: fnclex
  loc_00E2B8A9: jge 00E2B63Ah
  loc_00E2B8AF: jmp 00E2B629h
  loc_00E2B8B4: call [ecx+000004A8h]
  loc_00E2B8BA: lea edx, var_30
  loc_00E2B8BD: push eax
  loc_00E2B8BE: push edx
  loc_00E2B8BF: call edi
  loc_00E2B8C1: push eax
  loc_00E2B8C2: lea eax, var_50
  loc_00E2B8C5: push eax
  loc_00E2B8C6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2B8CC: add esp, 00000010h
  loc_00E2B8CF: push eax
  loc_00E2B8D0: call [00401120h] ; __vbaCastObjVar
  loc_00E2B8D6: lea ecx, var_34
  loc_00E2B8D9: push eax
  loc_00E2B8DA: push ecx
  loc_00E2B8DB: call edi
  loc_00E2B8DD: mov edx, [eax]
  loc_00E2B8DF: lea ecx, var_38
  loc_00E2B8E2: push ecx
  loc_00E2B8E3: push eax
  loc_00E2B8E4: mov var_D0, eax
  loc_00E2B8EA: call [edx+00000054h]
  loc_00E2B8ED: test eax, eax
  loc_00E2B8EF: fnclex
  loc_00E2B8F1: jge 00E2B904h
  loc_00E2B8F3: mov edx, var_D0
  loc_00E2B8F9: push 00000054h
  loc_00E2B8FB: push 006DCBE8h
  loc_00E2B900: push edx
  loc_00E2B901: push eax
  loc_00E2B902: call ebx
  loc_00E2B904: lea edx, var_3C
  loc_00E2B907: mov eax, 00000002h
  loc_00E2B90C: push edx
  loc_00E2B90D: mov ecx, var_38
  loc_00E2B910: sub esp, 00000010h
  loc_00E2B913: mov var_90, eax
  loc_00E2B919: mov edx, esp
  loc_00E2B91B: mov var_88, 00000009h
  loc_00E2B925: mov var_D8, ecx
  loc_00E2B92B: mov ecx, [ecx]
  loc_00E2B92D: mov [edx], eax
  loc_00E2B92F: mov eax, var_8C
  loc_00E2B935: mov [edx+00000004h], eax
  loc_00E2B938: mov eax, var_88
  loc_00E2B93E: mov [edx+00000008h], eax
  loc_00E2B941: mov eax, var_84
  loc_00E2B947: mov [edx+0000000Ch], eax
  loc_00E2B94A: mov edx, var_38
  loc_00E2B94D: push edx
  loc_00E2B94E: call [ecx+00000028h]
  loc_00E2B951: test eax, eax
  loc_00E2B953: fnclex
  loc_00E2B955: jge 00E2B968h
  loc_00E2B957: mov ecx, var_D8
  loc_00E2B95D: push 00000028h
  loc_00E2B95F: push 006E09E8h
  loc_00E2B964: push ecx
  loc_00E2B965: push eax
  loc_00E2B966: call ebx
  loc_00E2B968: mov ecx, var_3C
  loc_00E2B96B: mov eax, 00000002h
  loc_00E2B970: mov var_98, 00000000h
  loc_00E2B97A: mov var_A0, eax
  loc_00E2B980: mov edx, [ecx]
  loc_00E2B982: sub esp, 00000010h
  loc_00E2B985: mov var_E0, ecx
  loc_00E2B98B: mov ecx, esp
  loc_00E2B98D: mov [ecx], eax
  loc_00E2B98F: mov eax, var_9C
  loc_00E2B995: mov [ecx+00000004h], eax
  loc_00E2B998: mov eax, var_98
  loc_00E2B99E: mov [ecx+00000008h], eax
  loc_00E2B9A1: mov eax, var_94
  loc_00E2B9A7: mov [ecx+0000000Ch], eax
  loc_00E2B9AA: mov ecx, var_3C
  loc_00E2B9AD: push ecx
  loc_00E2B9AE: call [edx+00000038h]
  loc_00E2B9B1: test eax, eax
  loc_00E2B9B3: fnclex
  loc_00E2B9B5: jge 00E2B9C8h
  loc_00E2B9B7: mov edx, var_E0
  loc_00E2B9BD: push 00000038h
  loc_00E2B9BF: push 006E09F8h
  loc_00E2B9C4: push edx
  loc_00E2B9C5: push eax
  loc_00E2B9C6: call ebx
  loc_00E2B9C8: lea eax, var_3C
  loc_00E2B9CB: lea ecx, var_38
  loc_00E2B9CE: push eax
  loc_00E2B9CF: lea edx, var_34
  loc_00E2B9D2: push ecx
  loc_00E2B9D3: lea eax, var_30
  loc_00E2B9D6: push edx
  loc_00E2B9D7: push eax
  loc_00E2B9D8: push 00000004h
  loc_00E2B9DA: call [00401048h] ; __vbaFreeObjList
  loc_00E2B9E0: add esp, 00000014h
  loc_00E2B9E3: lea ecx, var_50
  loc_00E2B9E6: call [00401024h] ; __vbaFreeVar
  loc_00E2B9EC: mov ecx, [esi]
  loc_00E2B9EE: push esi
  loc_00E2B9EF: call [ecx+00000398h]
  loc_00E2B9F5: lea edx, var_30
  loc_00E2B9F8: push eax
  loc_00E2B9F9: push edx
  loc_00E2B9FA: call edi
  loc_00E2B9FC: mov ecx, [eax]
  loc_00E2B9FE: lea edx, var_C4
  loc_00E2BA04: push edx
  loc_00E2BA05: push eax
  loc_00E2BA06: mov var_D0, eax
  loc_00E2BA0C: call [ecx+00000098h]
  loc_00E2BA12: test eax, eax
  loc_00E2BA14: fnclex
  loc_00E2BA16: jge 00E2BA2Ch
  loc_00E2BA18: mov ecx, var_D0
  loc_00E2BA1E: push 00000098h
  loc_00E2BA23: push 006DCAD0h
  loc_00E2BA28: push ecx
  loc_00E2BA29: push eax
  loc_00E2BA2A: call ebx
  loc_00E2BA2C: xor edx, edx
  loc_00E2BA2E: cmp var_C4, FFFFFFh
  loc_00E2BA36: lea ecx, var_30
  loc_00E2BA39: setz dl
  loc_00E2BA3C: neg edx
  loc_00E2BA3E: mov var_D8, dx
  loc_00E2BA45: call [00401254h] ; __vbaFreeObj
  loc_00E2BA4B: cmp var_D8, 0000h
  loc_00E2BA53: mov eax, [esi]
  loc_00E2BA55: push esi
  loc_00E2BA56: jz 00E2C34Fh
  loc_00E2BA5C: call [eax+000003B0h]
  loc_00E2BA62: lea ecx, var_30
  loc_00E2BA65: push eax
  loc_00E2BA66: push ecx
  loc_00E2BA67: call edi
  loc_00E2BA69: mov edx, [eax]
  loc_00E2BA6B: lea ecx, var_28
  loc_00E2BA6E: push ecx
  loc_00E2BA6F: push eax
  loc_00E2BA70: mov var_D0, eax
  loc_00E2BA76: call [edx+00000050h]
  loc_00E2BA79: test eax, eax
  loc_00E2BA7B: fnclex
  loc_00E2BA7D: jge 00E2BA90h
  loc_00E2BA7F: mov edx, var_D0
  loc_00E2BA85: push 00000050h
  loc_00E2BA87: push 006E0410h
  loc_00E2BA8C: push edx
  loc_00E2BA8D: push eax
  loc_00E2BA8E: call ebx
  loc_00E2BA90: mov eax, var_28
  loc_00E2BA93: push eax
  loc_00E2BA94: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2BA9A: call [004010D8h] ; __vbaFpR8
  loc_00E2BAA0: fcomp real8 ptr [004022E8h]
  loc_00E2BAA6: mov var_190, 00000001h
  loc_00E2BAB0: fnstsw ax
  loc_00E2BAB2: test ah, 40h
  loc_00E2BAB5: jnz 00E2BAC1h
  loc_00E2BAB7: mov var_190, 00000000h
  loc_00E2BAC1: lea ecx, var_28
  loc_00E2BAC4: call [00401258h] ; __vbaFreeStr
  loc_00E2BACA: lea ecx, var_30
  loc_00E2BACD: call [00401254h] ; __vbaFreeObj
  loc_00E2BAD3: mov eax, var_190
  loc_00E2BAD9: neg eax
  loc_00E2BADB: test ax, ax
  loc_00E2BADE: jz 00E2BD40h
  loc_00E2BAE4: mov ecx, [esi]
  loc_00E2BAE6: push 006DCBD8h
  loc_00E2BAEB: push 00000000h
  loc_00E2BAED: push 00000018h
  loc_00E2BAEF: push esi
  loc_00E2BAF0: call [ecx+000004A8h]
  loc_00E2BAF6: lea edx, var_30
  loc_00E2BAF9: push eax
  loc_00E2BAFA: push edx
  loc_00E2BAFB: call edi
  loc_00E2BAFD: push eax
  loc_00E2BAFE: lea eax, var_50
  loc_00E2BB01: push eax
  loc_00E2BB02: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2BB08: add esp, 00000010h
  loc_00E2BB0B: push eax
  loc_00E2BB0C: call [00401120h] ; __vbaCastObjVar
  loc_00E2BB12: lea ecx, var_34
  loc_00E2BB15: push eax
  loc_00E2BB16: push ecx
  loc_00E2BB17: call edi
  loc_00E2BB19: mov edx, [eax]
  loc_00E2BB1B: lea ecx, var_38
  loc_00E2BB1E: push ecx
  loc_00E2BB1F: push eax
  loc_00E2BB20: mov var_D0, eax
  loc_00E2BB26: call [edx+00000054h]
  loc_00E2BB29: test eax, eax
  loc_00E2BB2B: fnclex
  loc_00E2BB2D: jge 00E2BB40h
  loc_00E2BB2F: mov edx, var_D0
  loc_00E2BB35: push 00000054h
  loc_00E2BB37: push 006DCBE8h
  loc_00E2BB3C: push edx
  loc_00E2BB3D: push eax
  loc_00E2BB3E: call ebx
  loc_00E2BB40: lea edx, var_3C
  loc_00E2BB43: mov eax, 00000002h
  loc_00E2BB48: push edx
  loc_00E2BB49: mov ecx, var_38
  loc_00E2BB4C: sub esp, 00000010h
  loc_00E2BB4F: mov var_90, eax
  loc_00E2BB55: mov edx, esp
  loc_00E2BB57: mov var_88, 0000000Ah
  loc_00E2BB61: mov var_D8, ecx
  loc_00E2BB67: mov ecx, [ecx]
  loc_00E2BB69: mov [edx], eax
  loc_00E2BB6B: mov eax, var_8C
  loc_00E2BB71: mov [edx+00000004h], eax
  loc_00E2BB74: mov eax, var_88
  loc_00E2BB7A: mov [edx+00000008h], eax
  loc_00E2BB7D: mov eax, var_84
  loc_00E2BB83: mov [edx+0000000Ch], eax
  loc_00E2BB86: mov edx, var_38
  loc_00E2BB89: push edx
  loc_00E2BB8A: call [ecx+00000028h]
  loc_00E2BB8D: test eax, eax
  loc_00E2BB8F: fnclex
  loc_00E2BB91: jge 00E2BBA4h
  loc_00E2BB93: mov ecx, var_D8
  loc_00E2BB99: push 00000028h
  loc_00E2BB9B: push 006E09E8h
  loc_00E2BBA0: push ecx
  loc_00E2BBA1: push eax
  loc_00E2BBA2: call ebx
  loc_00E2BBA4: mov ecx, var_3C
  loc_00E2BBA7: mov eax, 00000002h
  loc_00E2BBAC: mov var_98, 00000000h
  loc_00E2BBB6: mov var_A0, eax
  loc_00E2BBBC: mov edx, [ecx]
  loc_00E2BBBE: sub esp, 00000010h
  loc_00E2BBC1: mov var_E0, ecx
  loc_00E2BBC7: mov ecx, esp
  loc_00E2BBC9: mov [ecx], eax
  loc_00E2BBCB: mov eax, var_9C
  loc_00E2BBD1: mov [ecx+00000004h], eax
  loc_00E2BBD4: mov eax, var_98
  loc_00E2BBDA: mov [ecx+00000008h], eax
  loc_00E2BBDD: mov eax, var_94
  loc_00E2BBE3: mov [ecx+0000000Ch], eax
  loc_00E2BBE6: mov ecx, var_3C
  loc_00E2BBE9: push ecx
  loc_00E2BBEA: call [edx+00000038h]
  loc_00E2BBED: test eax, eax
  loc_00E2BBEF: fnclex
  loc_00E2BBF1: jge 00E2BC04h
  loc_00E2BBF3: mov edx, var_E0
  loc_00E2BBF9: push 00000038h
  loc_00E2BBFB: push 006E09F8h
  loc_00E2BC00: push edx
  loc_00E2BC01: push eax
  loc_00E2BC02: call ebx
  loc_00E2BC04: lea eax, var_3C
  loc_00E2BC07: lea ecx, var_38
  loc_00E2BC0A: push eax
  loc_00E2BC0B: lea edx, var_34
  loc_00E2BC0E: push ecx
  loc_00E2BC0F: lea eax, var_30
  loc_00E2BC12: push edx
  loc_00E2BC13: push eax
  loc_00E2BC14: push 00000004h
  loc_00E2BC16: call [00401048h] ; __vbaFreeObjList
  loc_00E2BC1C: add esp, 00000014h
  loc_00E2BC1F: lea ecx, var_50
  loc_00E2BC22: call [00401024h] ; __vbaFreeVar
  loc_00E2BC28: mov ecx, [esi]
  loc_00E2BC2A: push 006DCBD8h
  loc_00E2BC2F: push 00000000h
  loc_00E2BC31: push 00000018h
  loc_00E2BC33: push esi
  loc_00E2BC34: call [ecx+000004A8h]
  loc_00E2BC3A: lea edx, var_30
  loc_00E2BC3D: push eax
  loc_00E2BC3E: push edx
  loc_00E2BC3F: call edi
  loc_00E2BC41: push eax
  loc_00E2BC42: lea eax, var_50
  loc_00E2BC45: push eax
  loc_00E2BC46: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2BC4C: add esp, 00000010h
  loc_00E2BC4F: push eax
  loc_00E2BC50: call [00401120h] ; __vbaCastObjVar
  loc_00E2BC56: lea ecx, var_34
  loc_00E2BC59: push eax
  loc_00E2BC5A: push ecx
  loc_00E2BC5B: call edi
  loc_00E2BC5D: mov edx, [eax]
  loc_00E2BC5F: lea ecx, var_38
  loc_00E2BC62: push ecx
  loc_00E2BC63: push eax
  loc_00E2BC64: mov var_D0, eax
  loc_00E2BC6A: call [edx+00000054h]
  loc_00E2BC6D: test eax, eax
  loc_00E2BC6F: fnclex
  loc_00E2BC71: jge 00E2BC84h
  loc_00E2BC73: mov edx, var_D0
  loc_00E2BC79: push 00000054h
  loc_00E2BC7B: push 006DCBE8h
  loc_00E2BC80: push edx
  loc_00E2BC81: push eax
  loc_00E2BC82: call ebx
  loc_00E2BC84: lea edx, var_3C
  loc_00E2BC87: mov eax, 00000002h
  loc_00E2BC8C: push edx
  loc_00E2BC8D: mov ecx, var_38
  loc_00E2BC90: sub esp, 00000010h
  loc_00E2BC93: mov var_90, eax
  loc_00E2BC99: mov edx, esp
  loc_00E2BC9B: mov var_88, 00000007h
  loc_00E2BCA5: mov var_D8, ecx
  loc_00E2BCAB: mov ecx, [ecx]
  loc_00E2BCAD: mov [edx], eax
  loc_00E2BCAF: mov eax, var_8C
  loc_00E2BCB5: mov [edx+00000004h], eax
  loc_00E2BCB8: mov eax, var_88
  loc_00E2BCBE: mov [edx+00000008h], eax
  loc_00E2BCC1: mov eax, var_84
  loc_00E2BCC7: mov [edx+0000000Ch], eax
  loc_00E2BCCA: mov edx, var_38
  loc_00E2BCCD: push edx
  loc_00E2BCCE: call [ecx+00000028h]
  loc_00E2BCD1: test eax, eax
  loc_00E2BCD3: fnclex
  loc_00E2BCD5: jge 00E2BCE8h
  loc_00E2BCD7: mov ecx, var_D8
  loc_00E2BCDD: push 00000028h
  loc_00E2BCDF: push 006E09E8h
  loc_00E2BCE4: push ecx
  loc_00E2BCE5: push eax
  loc_00E2BCE6: call ebx
  loc_00E2BCE8: mov ecx, var_3C
  loc_00E2BCEB: mov eax, 00000002h
  loc_00E2BCF0: mov var_98, 00000000h
  loc_00E2BCFA: mov var_A0, eax
  loc_00E2BD00: mov edx, [ecx]
  loc_00E2BD02: sub esp, 00000010h
  loc_00E2BD05: mov var_E0, ecx
  loc_00E2BD0B: mov ecx, esp
  loc_00E2BD0D: mov [ecx], eax
  loc_00E2BD0F: mov eax, var_9C
  loc_00E2BD15: mov [ecx+00000004h], eax
  loc_00E2BD18: mov eax, var_98
  loc_00E2BD1E: mov [ecx+00000008h], eax
  loc_00E2BD21: mov eax, var_94
  loc_00E2BD27: mov [ecx+0000000Ch], eax
  loc_00E2BD2A: mov ecx, var_3C
  loc_00E2BD2D: push ecx
  loc_00E2BD2E: call [edx+00000038h]
  loc_00E2BD31: test eax, eax
  loc_00E2BD33: fnclex
  loc_00E2BD35: jge 00E2C61Ch
  loc_00E2BD3B: jmp 00E2C60Bh
  loc_00E2BD40: mov ecx, [esi]
  loc_00E2BD42: push esi
  loc_00E2BD43: call [ecx+000003B0h]
  loc_00E2BD49: lea edx, var_30
  loc_00E2BD4C: push eax
  loc_00E2BD4D: push edx
  loc_00E2BD4E: call edi
  loc_00E2BD50: mov ecx, [eax]
  loc_00E2BD52: lea edx, var_28
  loc_00E2BD55: push edx
  loc_00E2BD56: push eax
  loc_00E2BD57: mov var_D0, eax
  loc_00E2BD5D: call [ecx+00000050h]
  loc_00E2BD60: test eax, eax
  loc_00E2BD62: fnclex
  loc_00E2BD64: jge 00E2BD77h
  loc_00E2BD66: mov ecx, var_D0
  loc_00E2BD6C: push 00000050h
  loc_00E2BD6E: push 006E0410h
  loc_00E2BD73: push ecx
  loc_00E2BD74: push eax
  loc_00E2BD75: call ebx
  loc_00E2BD77: mov edx, var_28
  loc_00E2BD7A: push edx
  loc_00E2BD7B: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2BD81: call [004010D8h] ; __vbaFpR8
  loc_00E2BD87: fcomp real8 ptr [004022E8h]
  loc_00E2BD8D: mov var_194, 00000001h
  loc_00E2BD97: fnstsw ax
  loc_00E2BD99: test ah, 01h
  loc_00E2BD9C: jnz 00E2BDA8h
  loc_00E2BD9E: mov var_194, 00000000h
  loc_00E2BDA8: lea ecx, var_28
  loc_00E2BDAB: call [00401258h] ; __vbaFreeStr
  loc_00E2BDB1: lea ecx, var_30
  loc_00E2BDB4: call [00401254h] ; __vbaFreeObj
  loc_00E2BDBA: mov eax, var_194
  loc_00E2BDC0: push 006DCBD8h
  loc_00E2BDC5: neg eax
  loc_00E2BDC7: push 00000000h
  loc_00E2BDC9: push 00000018h
  loc_00E2BDCB: test ax, ax
  loc_00E2BDCE: mov eax, [esi]
  loc_00E2BDD0: push esi
  loc_00E2BDD1: jz 00E2C044h
  loc_00E2BDD7: call [eax+000004A8h]
  loc_00E2BDDD: lea ecx, var_30
  loc_00E2BDE0: push eax
  loc_00E2BDE1: push ecx
  loc_00E2BDE2: call edi
  loc_00E2BDE4: lea edx, var_50
  loc_00E2BDE7: push eax
  loc_00E2BDE8: push edx
  loc_00E2BDE9: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2BDEF: add esp, 00000010h
  loc_00E2BDF2: push eax
  loc_00E2BDF3: call [00401120h] ; __vbaCastObjVar
  loc_00E2BDF9: push eax
  loc_00E2BDFA: lea eax, var_34
  loc_00E2BDFD: push eax
  loc_00E2BDFE: call edi
  loc_00E2BE00: mov ecx, [eax]
  loc_00E2BE02: lea edx, var_38
  loc_00E2BE05: push edx
  loc_00E2BE06: push eax
  loc_00E2BE07: mov var_D0, eax
  loc_00E2BE0D: call [ecx+00000054h]
  loc_00E2BE10: test eax, eax
  loc_00E2BE12: fnclex
  loc_00E2BE14: jge 00E2BE27h
  loc_00E2BE16: mov ecx, var_D0
  loc_00E2BE1C: push 00000054h
  loc_00E2BE1E: push 006DCBE8h
  loc_00E2BE23: push ecx
  loc_00E2BE24: push eax
  loc_00E2BE25: call ebx
  loc_00E2BE27: mov ecx, var_38
  loc_00E2BE2A: mov eax, 00000002h
  loc_00E2BE2F: mov var_88, 0000000Ah
  loc_00E2BE39: mov var_90, eax
  loc_00E2BE3F: mov edx, [ecx]
  loc_00E2BE41: mov var_D8, ecx
  loc_00E2BE47: lea ecx, var_3C
  loc_00E2BE4A: push ecx
  loc_00E2BE4B: sub esp, 00000010h
  loc_00E2BE4E: mov ecx, esp
  loc_00E2BE50: mov [ecx], eax
  loc_00E2BE52: mov eax, var_8C
  loc_00E2BE58: mov [ecx+00000004h], eax
  loc_00E2BE5B: mov eax, var_88
  loc_00E2BE61: mov [ecx+00000008h], eax
  loc_00E2BE64: mov eax, var_84
  loc_00E2BE6A: mov [ecx+0000000Ch], eax
  loc_00E2BE6D: mov ecx, var_38
  loc_00E2BE70: push ecx
  loc_00E2BE71: call [edx+00000028h]
  loc_00E2BE74: test eax, eax
  loc_00E2BE76: fnclex
  loc_00E2BE78: jge 00E2BE8Bh
  loc_00E2BE7A: mov edx, var_D8
  loc_00E2BE80: push 00000028h
  loc_00E2BE82: push 006E09E8h
  loc_00E2BE87: push edx
  loc_00E2BE88: push eax
  loc_00E2BE89: call ebx
  loc_00E2BE8B: sub esp, 00000010h
  loc_00E2BE8E: mov eax, 00000002h
  loc_00E2BE93: mov edx, esp
  loc_00E2BE95: mov ecx, var_3C
  loc_00E2BE98: mov var_A0, eax
  loc_00E2BE9E: mov var_98, 00000000h
  loc_00E2BEA8: mov [edx], eax
  loc_00E2BEAA: mov eax, var_9C
  loc_00E2BEB0: mov var_E0, ecx
  loc_00E2BEB6: mov ecx, [ecx]
  loc_00E2BEB8: mov [edx+00000004h], eax
  loc_00E2BEBB: mov eax, var_98
  loc_00E2BEC1: mov [edx+00000008h], eax
  loc_00E2BEC4: mov eax, var_94
  loc_00E2BECA: mov [edx+0000000Ch], eax
  loc_00E2BECD: mov edx, var_3C
  loc_00E2BED0: push edx
  loc_00E2BED1: call [ecx+00000038h]
  loc_00E2BED4: test eax, eax
  loc_00E2BED6: fnclex
  loc_00E2BED8: jge 00E2BEEBh
  loc_00E2BEDA: mov ecx, var_E0
  loc_00E2BEE0: push 00000038h
  loc_00E2BEE2: push 006E09F8h
  loc_00E2BEE7: push ecx
  loc_00E2BEE8: push eax
  loc_00E2BEE9: call ebx
  loc_00E2BEEB: lea edx, var_3C
  loc_00E2BEEE: lea eax, var_38
  loc_00E2BEF1: push edx
  loc_00E2BEF2: lea ecx, var_34
  loc_00E2BEF5: push eax
  loc_00E2BEF6: lea edx, var_30
  loc_00E2BEF9: push ecx
  loc_00E2BEFA: push edx
  loc_00E2BEFB: push 00000004h
  loc_00E2BEFD: call [00401048h] ; __vbaFreeObjList
  loc_00E2BF03: add esp, 00000014h
  loc_00E2BF06: lea ecx, var_50
  loc_00E2BF09: call [00401024h] ; __vbaFreeVar
  loc_00E2BF0F: mov eax, [esi]
  loc_00E2BF11: push 006DCBD8h
  loc_00E2BF16: push 00000000h
  loc_00E2BF18: push 00000018h
  loc_00E2BF1A: push esi
  loc_00E2BF1B: call [eax+000004A8h]
  loc_00E2BF21: lea ecx, var_30
  loc_00E2BF24: push eax
  loc_00E2BF25: push ecx
  loc_00E2BF26: call edi
  loc_00E2BF28: lea edx, var_50
  loc_00E2BF2B: push eax
  loc_00E2BF2C: push edx
  loc_00E2BF2D: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2BF33: add esp, 00000010h
  loc_00E2BF36: push eax
  loc_00E2BF37: call [00401120h] ; __vbaCastObjVar
  loc_00E2BF3D: push eax
  loc_00E2BF3E: lea eax, var_34
  loc_00E2BF41: push eax
  loc_00E2BF42: call edi
  loc_00E2BF44: mov ecx, [eax]
  loc_00E2BF46: lea edx, var_38
  loc_00E2BF49: push edx
  loc_00E2BF4A: push eax
  loc_00E2BF4B: mov var_D0, eax
  loc_00E2BF51: call [ecx+00000054h]
  loc_00E2BF54: test eax, eax
  loc_00E2BF56: fnclex
  loc_00E2BF58: jge 00E2BF6Bh
  loc_00E2BF5A: mov ecx, var_D0
  loc_00E2BF60: push 00000054h
  loc_00E2BF62: push 006DCBE8h
  loc_00E2BF67: push ecx
  loc_00E2BF68: push eax
  loc_00E2BF69: call ebx
  loc_00E2BF6B: mov ecx, var_38
  loc_00E2BF6E: mov eax, 00000002h
  loc_00E2BF73: mov var_88, 00000007h
  loc_00E2BF7D: mov var_90, eax
  loc_00E2BF83: mov edx, [ecx]
  loc_00E2BF85: mov var_D8, ecx
  loc_00E2BF8B: lea ecx, var_3C
  loc_00E2BF8E: push ecx
  loc_00E2BF8F: sub esp, 00000010h
  loc_00E2BF92: mov ecx, esp
  loc_00E2BF94: mov [ecx], eax
  loc_00E2BF96: mov eax, var_8C
  loc_00E2BF9C: mov [ecx+00000004h], eax
  loc_00E2BF9F: mov eax, var_88
  loc_00E2BFA5: mov [ecx+00000008h], eax
  loc_00E2BFA8: mov eax, var_84
  loc_00E2BFAE: mov [ecx+0000000Ch], eax
  loc_00E2BFB1: mov ecx, var_38
  loc_00E2BFB4: push ecx
  loc_00E2BFB5: call [edx+00000028h]
  loc_00E2BFB8: test eax, eax
  loc_00E2BFBA: fnclex
  loc_00E2BFBC: jge 00E2BFCFh
  loc_00E2BFBE: mov edx, var_D8
  loc_00E2BFC4: push 00000028h
  loc_00E2BFC6: push 006E09E8h
  loc_00E2BFCB: push edx
  loc_00E2BFCC: push eax
  loc_00E2BFCD: call ebx
  loc_00E2BFCF: sub esp, 00000010h
  loc_00E2BFD2: mov eax, 00000002h
  loc_00E2BFD7: mov edx, esp
  loc_00E2BFD9: mov ecx, var_3C
  loc_00E2BFDC: mov var_A0, eax
  loc_00E2BFE2: mov var_98, 00000000h
  loc_00E2BFEC: mov [edx], eax
  loc_00E2BFEE: mov eax, var_9C
  loc_00E2BFF4: mov var_E0, ecx
  loc_00E2BFFA: mov ecx, [ecx]
  loc_00E2BFFC: mov [edx+00000004h], eax
  loc_00E2BFFF: mov eax, var_98
  loc_00E2C005: mov [edx+00000008h], eax
  loc_00E2C008: mov eax, var_94
  loc_00E2C00E: mov [edx+0000000Ch], eax
  loc_00E2C011: mov edx, var_3C
  loc_00E2C014: push edx
  loc_00E2C015: call [ecx+00000038h]
  loc_00E2C018: test eax, eax
  loc_00E2C01A: fnclex
  loc_00E2C01C: jge 00E2C02Fh
  loc_00E2C01E: mov ecx, var_E0
  loc_00E2C024: push 00000038h
  loc_00E2C026: push 006E09F8h
  loc_00E2C02B: push ecx
  loc_00E2C02C: push eax
  loc_00E2C02D: call ebx
  loc_00E2C02F: lea edx, var_3C
  loc_00E2C032: lea eax, var_38
  loc_00E2C035: push edx
  loc_00E2C036: lea ecx, var_34
  loc_00E2C039: push eax
  loc_00E2C03A: lea edx, var_30
  loc_00E2C03D: push ecx
  loc_00E2C03E: push edx
  loc_00E2C03F: jmp 00E2C62Ch
  loc_00E2C044: call [eax+000004A8h]
  loc_00E2C04A: lea ecx, var_34
  loc_00E2C04D: push eax
  loc_00E2C04E: push ecx
  loc_00E2C04F: call edi
  loc_00E2C051: lea edx, var_50
  loc_00E2C054: push eax
  loc_00E2C055: push edx
  loc_00E2C056: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2C05C: add esp, 00000010h
  loc_00E2C05F: push eax
  loc_00E2C060: call [00401120h] ; __vbaCastObjVar
  loc_00E2C066: push eax
  loc_00E2C067: lea eax, var_38
  loc_00E2C06A: push eax
  loc_00E2C06B: call edi
  loc_00E2C06D: mov ecx, [eax]
  loc_00E2C06F: lea edx, var_3C
  loc_00E2C072: push edx
  loc_00E2C073: push eax
  loc_00E2C074: mov var_D8, eax
  loc_00E2C07A: call [ecx+00000054h]
  loc_00E2C07D: test eax, eax
  loc_00E2C07F: fnclex
  loc_00E2C081: jge 00E2C094h
  loc_00E2C083: mov ecx, var_D8
  loc_00E2C089: push 00000054h
  loc_00E2C08B: push 006DCBE8h
  loc_00E2C090: push ecx
  loc_00E2C091: push eax
  loc_00E2C092: call ebx
  loc_00E2C094: mov ecx, var_3C
  loc_00E2C097: mov eax, 00000002h
  loc_00E2C09C: mov var_88, 0000000Ah
  loc_00E2C0A6: mov var_90, eax
  loc_00E2C0AC: mov edx, [ecx]
  loc_00E2C0AE: mov var_E0, ecx
  loc_00E2C0B4: lea ecx, var_40
  loc_00E2C0B7: push ecx
  loc_00E2C0B8: sub esp, 00000010h
  loc_00E2C0BB: mov ecx, esp
  loc_00E2C0BD: mov [ecx], eax
  loc_00E2C0BF: mov eax, var_8C
  loc_00E2C0C5: mov [ecx+00000004h], eax
  loc_00E2C0C8: mov eax, var_88
  loc_00E2C0CE: mov [ecx+00000008h], eax
  loc_00E2C0D1: mov eax, var_84
  loc_00E2C0D7: mov [ecx+0000000Ch], eax
  loc_00E2C0DA: mov ecx, var_3C
  loc_00E2C0DD: push ecx
  loc_00E2C0DE: call [edx+00000028h]
  loc_00E2C0E1: test eax, eax
  loc_00E2C0E3: fnclex
  loc_00E2C0E5: jge 00E2C0F8h
  loc_00E2C0E7: mov edx, var_E0
  loc_00E2C0ED: push 00000028h
  loc_00E2C0EF: push 006E09E8h
  loc_00E2C0F4: push edx
  loc_00E2C0F5: push eax
  loc_00E2C0F6: call ebx
  loc_00E2C0F8: mov eax, var_40
  loc_00E2C0FB: mov ecx, [esi]
  loc_00E2C0FD: push esi
  loc_00E2C0FE: mov var_E8, eax
  loc_00E2C104: call [ecx+000003B0h]
  loc_00E2C10A: lea edx, var_30
  loc_00E2C10D: push eax
  loc_00E2C10E: push edx
  loc_00E2C10F: call edi
  loc_00E2C111: mov ecx, [eax]
  loc_00E2C113: lea edx, var_28
  loc_00E2C116: push edx
  loc_00E2C117: push eax
  loc_00E2C118: mov var_D0, eax
  loc_00E2C11E: call [ecx+00000050h]
  loc_00E2C121: test eax, eax
  loc_00E2C123: fnclex
  loc_00E2C125: jge 00E2C138h
  loc_00E2C127: mov ecx, var_D0
  loc_00E2C12D: push 00000050h
  loc_00E2C12F: push 006E0410h
  loc_00E2C134: push ecx
  loc_00E2C135: push eax
  loc_00E2C136: call ebx
  loc_00E2C138: mov eax, var_28
  loc_00E2C13B: mov edx, var_E8
  loc_00E2C141: mov var_58, eax
  loc_00E2C144: mov eax, 00000008h
  loc_00E2C149: mov var_28, 00000000h
  loc_00E2C150: mov var_60, eax
  loc_00E2C153: mov ecx, [edx]
  loc_00E2C155: sub esp, 00000010h
  loc_00E2C158: mov edx, esp
  loc_00E2C15A: mov [edx], eax
  loc_00E2C15C: mov eax, var_5C
  loc_00E2C15F: mov [edx+00000004h], eax
  loc_00E2C162: mov eax, var_58
  loc_00E2C165: mov [edx+00000008h], eax
  loc_00E2C168: mov eax, var_54
  loc_00E2C16B: mov [edx+0000000Ch], eax
  loc_00E2C16E: mov edx, var_E8
  loc_00E2C174: push edx
  loc_00E2C175: call [ecx+00000038h]
  loc_00E2C178: test eax, eax
  loc_00E2C17A: fnclex
  loc_00E2C17C: jge 00E2C18Fh
  loc_00E2C17E: mov ecx, var_E8
  loc_00E2C184: push 00000038h
  loc_00E2C186: push 006E09F8h
  loc_00E2C18B: push ecx
  loc_00E2C18C: push eax
  loc_00E2C18D: call ebx
  loc_00E2C18F: lea edx, var_40
  loc_00E2C192: lea eax, var_3C
  loc_00E2C195: push edx
  loc_00E2C196: lea ecx, var_38
  loc_00E2C199: push eax
  loc_00E2C19A: lea edx, var_34
  loc_00E2C19D: push ecx
  loc_00E2C19E: lea eax, var_30
  loc_00E2C1A1: push edx
  loc_00E2C1A2: push eax
  loc_00E2C1A3: push 00000005h
  loc_00E2C1A5: call [00401048h] ; __vbaFreeObjList
  loc_00E2C1AB: lea ecx, var_60
  loc_00E2C1AE: lea edx, var_50
  loc_00E2C1B1: push ecx
  loc_00E2C1B2: push edx
  loc_00E2C1B3: push 00000002h
  loc_00E2C1B5: call [00401038h] ; __vbaFreeVarList
  loc_00E2C1BB: mov eax, [esi]
  loc_00E2C1BD: add esp, 00000024h
  loc_00E2C1C0: push 006DCBD8h
  loc_00E2C1C5: push 00000000h
  loc_00E2C1C7: push 00000018h
  loc_00E2C1C9: push esi
  loc_00E2C1CA: call [eax+000004A8h]
  loc_00E2C1D0: lea ecx, var_34
  loc_00E2C1D3: push eax
  loc_00E2C1D4: push ecx
  loc_00E2C1D5: call edi
  loc_00E2C1D7: lea edx, var_50
  loc_00E2C1DA: push eax
  loc_00E2C1DB: push edx
  loc_00E2C1DC: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2C1E2: add esp, 00000010h
  loc_00E2C1E5: push eax
  loc_00E2C1E6: call [00401120h] ; __vbaCastObjVar
  loc_00E2C1EC: push eax
  loc_00E2C1ED: lea eax, var_38
  loc_00E2C1F0: push eax
  loc_00E2C1F1: call edi
  loc_00E2C1F3: mov ecx, [eax]
  loc_00E2C1F5: lea edx, var_3C
  loc_00E2C1F8: push edx
  loc_00E2C1F9: push eax
  loc_00E2C1FA: mov var_D8, eax
  loc_00E2C200: call [ecx+00000054h]
  loc_00E2C203: test eax, eax
  loc_00E2C205: fnclex
  loc_00E2C207: jge 00E2C21Ah
  loc_00E2C209: mov ecx, var_D8
  loc_00E2C20F: push 00000054h
  loc_00E2C211: push 006DCBE8h
  loc_00E2C216: push ecx
  loc_00E2C217: push eax
  loc_00E2C218: call ebx
  loc_00E2C21A: mov ecx, var_3C
  loc_00E2C21D: mov eax, 00000002h
  loc_00E2C222: mov var_88, 00000007h
  loc_00E2C22C: mov var_90, eax
  loc_00E2C232: mov edx, [ecx]
  loc_00E2C234: mov var_E0, ecx
  loc_00E2C23A: lea ecx, var_40
  loc_00E2C23D: push ecx
  loc_00E2C23E: sub esp, 00000010h
  loc_00E2C241: mov ecx, esp
  loc_00E2C243: mov [ecx], eax
  loc_00E2C245: mov eax, var_8C
  loc_00E2C24B: mov [ecx+00000004h], eax
  loc_00E2C24E: mov eax, var_88
  loc_00E2C254: mov [ecx+00000008h], eax
  loc_00E2C257: mov eax, var_84
  loc_00E2C25D: mov [ecx+0000000Ch], eax
  loc_00E2C260: mov ecx, var_3C
  loc_00E2C263: push ecx
  loc_00E2C264: call [edx+00000028h]
  loc_00E2C267: test eax, eax
  loc_00E2C269: fnclex
  loc_00E2C26B: jge 00E2C27Eh
  loc_00E2C26D: mov edx, var_E0
  loc_00E2C273: push 00000028h
  loc_00E2C275: push 006E09E8h
  loc_00E2C27A: push edx
  loc_00E2C27B: push eax
  loc_00E2C27C: call ebx
  loc_00E2C27E: mov eax, var_40
  loc_00E2C281: mov ecx, [esi]
  loc_00E2C283: push esi
  loc_00E2C284: mov var_E8, eax
  loc_00E2C28A: call [ecx+0000039Ch]
  loc_00E2C290: lea edx, var_30
  loc_00E2C293: push eax
  loc_00E2C294: push edx
  loc_00E2C295: call edi
  loc_00E2C297: mov ecx, [eax]
  loc_00E2C299: lea edx, var_28
  loc_00E2C29C: push edx
  loc_00E2C29D: push eax
  loc_00E2C29E: mov var_D0, eax
  loc_00E2C2A4: call [ecx+000000A0h]
  loc_00E2C2AA: test eax, eax
  loc_00E2C2AC: fnclex
  loc_00E2C2AE: jge 00E2C2C4h
  loc_00E2C2B0: mov ecx, var_D0
  loc_00E2C2B6: push 000000A0h
  loc_00E2C2BB: push 006DCB70h
  loc_00E2C2C0: push ecx
  loc_00E2C2C1: push eax
  loc_00E2C2C2: call ebx
  loc_00E2C2C4: mov eax, var_28
  loc_00E2C2C7: mov edx, var_E8
  loc_00E2C2CD: mov var_58, eax
  loc_00E2C2D0: mov eax, 00000008h
  loc_00E2C2D5: mov var_28, 00000000h
  loc_00E2C2DC: mov var_60, eax
  loc_00E2C2DF: mov ecx, [edx]
  loc_00E2C2E1: sub esp, 00000010h
  loc_00E2C2E4: mov edx, esp
  loc_00E2C2E6: mov [edx], eax
  loc_00E2C2E8: mov eax, var_5C
  loc_00E2C2EB: mov [edx+00000004h], eax
  loc_00E2C2EE: mov eax, var_58
  loc_00E2C2F1: mov [edx+00000008h], eax
  loc_00E2C2F4: mov eax, var_54
  loc_00E2C2F7: mov [edx+0000000Ch], eax
  loc_00E2C2FA: mov edx, var_E8
  loc_00E2C300: push edx
  loc_00E2C301: call [ecx+00000038h]
  loc_00E2C304: test eax, eax
  loc_00E2C306: fnclex
  loc_00E2C308: jge 00E2C31Bh
  loc_00E2C30A: mov ecx, var_E8
  loc_00E2C310: push 00000038h
  loc_00E2C312: push 006E09F8h
  loc_00E2C317: push ecx
  loc_00E2C318: push eax
  loc_00E2C319: call ebx
  loc_00E2C31B: lea edx, var_40
  loc_00E2C31E: lea eax, var_3C
  loc_00E2C321: push edx
  loc_00E2C322: lea ecx, var_38
  loc_00E2C325: push eax
  loc_00E2C326: lea edx, var_34
  loc_00E2C329: push ecx
  loc_00E2C32A: lea eax, var_30
  loc_00E2C32D: push edx
  loc_00E2C32E: push eax
  loc_00E2C32F: push 00000005h
  loc_00E2C331: call [00401048h] ; __vbaFreeObjList
  loc_00E2C337: lea ecx, var_60
  loc_00E2C33A: lea edx, var_50
  loc_00E2C33D: push ecx
  loc_00E2C33E: push edx
  loc_00E2C33F: push 00000002h
  loc_00E2C341: call [00401038h] ; __vbaFreeVarList
  loc_00E2C347: add esp, 00000024h
  loc_00E2C34A: jmp 00E2C640h
  loc_00E2C34F: call [eax+00000398h]
  loc_00E2C355: lea ecx, var_30
  loc_00E2C358: push eax
  loc_00E2C359: push ecx
  loc_00E2C35A: call edi
  loc_00E2C35C: mov edx, [eax]
  loc_00E2C35E: lea ecx, var_C4
  loc_00E2C364: push ecx
  loc_00E2C365: push eax
  loc_00E2C366: mov var_D0, eax
  loc_00E2C36C: call [edx+00000098h]
  loc_00E2C372: test eax, eax
  loc_00E2C374: fnclex
  loc_00E2C376: jge 00E2C38Ch
  loc_00E2C378: mov edx, var_D0
  loc_00E2C37E: push 00000098h
  loc_00E2C383: push 006DCAD0h
  loc_00E2C388: push edx
  loc_00E2C389: push eax
  loc_00E2C38A: call ebx
  loc_00E2C38C: xor eax, eax
  loc_00E2C38E: lea ecx, var_30
  loc_00E2C391: cmp var_C4, ax
  loc_00E2C398: setz al
  loc_00E2C39B: neg eax
  loc_00E2C39D: mov var_D8, ax
  loc_00E2C3A4: call [00401254h] ; __vbaFreeObj
  loc_00E2C3AA: cmp var_D8, 0000h
  loc_00E2C3B2: jz 00E2C640h
  loc_00E2C3B8: mov ecx, [esi]
  loc_00E2C3BA: push 006DCBD8h
  loc_00E2C3BF: push 00000000h
  loc_00E2C3C1: push 00000018h
  loc_00E2C3C3: push esi
  loc_00E2C3C4: call [ecx+000004A8h]
  loc_00E2C3CA: lea edx, var_30
  loc_00E2C3CD: push eax
  loc_00E2C3CE: push edx
  loc_00E2C3CF: call edi
  loc_00E2C3D1: push eax
  loc_00E2C3D2: lea eax, var_50
  loc_00E2C3D5: push eax
  loc_00E2C3D6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2C3DC: add esp, 00000010h
  loc_00E2C3DF: push eax
  loc_00E2C3E0: call [00401120h] ; __vbaCastObjVar
  loc_00E2C3E6: lea ecx, var_34
  loc_00E2C3E9: push eax
  loc_00E2C3EA: push ecx
  loc_00E2C3EB: call edi
  loc_00E2C3ED: mov edx, [eax]
  loc_00E2C3EF: lea ecx, var_38
  loc_00E2C3F2: push ecx
  loc_00E2C3F3: push eax
  loc_00E2C3F4: mov var_D0, eax
  loc_00E2C3FA: call [edx+00000054h]
  loc_00E2C3FD: test eax, eax
  loc_00E2C3FF: fnclex
  loc_00E2C401: jge 00E2C414h
  loc_00E2C403: mov edx, var_D0
  loc_00E2C409: push 00000054h
  loc_00E2C40B: push 006DCBE8h
  loc_00E2C410: push edx
  loc_00E2C411: push eax
  loc_00E2C412: call ebx
  loc_00E2C414: lea edx, var_3C
  loc_00E2C417: mov eax, 00000002h
  loc_00E2C41C: push edx
  loc_00E2C41D: mov ecx, var_38
  loc_00E2C420: sub esp, 00000010h
  loc_00E2C423: mov var_90, eax
  loc_00E2C429: mov edx, esp
  loc_00E2C42B: mov var_88, 0000000Ah
  loc_00E2C435: mov var_D8, ecx
  loc_00E2C43B: mov ecx, [ecx]
  loc_00E2C43D: mov [edx], eax
  loc_00E2C43F: mov eax, var_8C
  loc_00E2C445: mov [edx+00000004h], eax
  loc_00E2C448: mov eax, var_88
  loc_00E2C44E: mov [edx+00000008h], eax
  loc_00E2C451: mov eax, var_84
  loc_00E2C457: mov [edx+0000000Ch], eax
  loc_00E2C45A: mov edx, var_38
  loc_00E2C45D: push edx
  loc_00E2C45E: call [ecx+00000028h]
  loc_00E2C461: test eax, eax
  loc_00E2C463: fnclex
  loc_00E2C465: jge 00E2C478h
  loc_00E2C467: mov ecx, var_D8
  loc_00E2C46D: push 00000028h
  loc_00E2C46F: push 006E09E8h
  loc_00E2C474: push ecx
  loc_00E2C475: push eax
  loc_00E2C476: call ebx
  loc_00E2C478: mov ecx, var_3C
  loc_00E2C47B: mov eax, 00000002h
  loc_00E2C480: mov var_98, 00000000h
  loc_00E2C48A: mov var_A0, eax
  loc_00E2C490: mov edx, [ecx]
  loc_00E2C492: sub esp, 00000010h
  loc_00E2C495: mov var_E0, ecx
  loc_00E2C49B: mov ecx, esp
  loc_00E2C49D: mov [ecx], eax
  loc_00E2C49F: mov eax, var_9C
  loc_00E2C4A5: mov [ecx+00000004h], eax
  loc_00E2C4A8: mov eax, var_98
  loc_00E2C4AE: mov [ecx+00000008h], eax
  loc_00E2C4B1: mov eax, var_94
  loc_00E2C4B7: mov [ecx+0000000Ch], eax
  loc_00E2C4BA: mov ecx, var_3C
  loc_00E2C4BD: push ecx
  loc_00E2C4BE: call [edx+00000038h]
  loc_00E2C4C1: test eax, eax
  loc_00E2C4C3: fnclex
  loc_00E2C4C5: jge 00E2C4D8h
  loc_00E2C4C7: mov edx, var_E0
  loc_00E2C4CD: push 00000038h
  loc_00E2C4CF: push 006E09F8h
  loc_00E2C4D4: push edx
  loc_00E2C4D5: push eax
  loc_00E2C4D6: call ebx
  loc_00E2C4D8: lea eax, var_3C
  loc_00E2C4DB: lea ecx, var_38
  loc_00E2C4DE: push eax
  loc_00E2C4DF: lea edx, var_34
  loc_00E2C4E2: push ecx
  loc_00E2C4E3: lea eax, var_30
  loc_00E2C4E6: push edx
  loc_00E2C4E7: push eax
  loc_00E2C4E8: push 00000004h
  loc_00E2C4EA: call [00401048h] ; __vbaFreeObjList
  loc_00E2C4F0: add esp, 00000014h
  loc_00E2C4F3: lea ecx, var_50
  loc_00E2C4F6: call [00401024h] ; __vbaFreeVar
  loc_00E2C4FC: mov ecx, [esi]
  loc_00E2C4FE: push 006DCBD8h
  loc_00E2C503: push 00000000h
  loc_00E2C505: push 00000018h
  loc_00E2C507: push esi
  loc_00E2C508: call [ecx+000004A8h]
  loc_00E2C50E: lea edx, var_30
  loc_00E2C511: push eax
  loc_00E2C512: push edx
  loc_00E2C513: call edi
  loc_00E2C515: push eax
  loc_00E2C516: lea eax, var_50
  loc_00E2C519: push eax
  loc_00E2C51A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2C520: add esp, 00000010h
  loc_00E2C523: push eax
  loc_00E2C524: call [00401120h] ; __vbaCastObjVar
  loc_00E2C52A: lea ecx, var_34
  loc_00E2C52D: push eax
  loc_00E2C52E: push ecx
  loc_00E2C52F: call edi
  loc_00E2C531: mov edx, [eax]
  loc_00E2C533: lea ecx, var_38
  loc_00E2C536: push ecx
  loc_00E2C537: push eax
  loc_00E2C538: mov var_D0, eax
  loc_00E2C53E: call [edx+00000054h]
  loc_00E2C541: test eax, eax
  loc_00E2C543: fnclex
  loc_00E2C545: jge 00E2C558h
  loc_00E2C547: mov edx, var_D0
  loc_00E2C54D: push 00000054h
  loc_00E2C54F: push 006DCBE8h
  loc_00E2C554: push edx
  loc_00E2C555: push eax
  loc_00E2C556: call ebx
  loc_00E2C558: lea edx, var_3C
  loc_00E2C55B: mov eax, 00000002h
  loc_00E2C560: push edx
  loc_00E2C561: mov ecx, var_38
  loc_00E2C564: sub esp, 00000010h
  loc_00E2C567: mov var_90, eax
  loc_00E2C56D: mov edx, esp
  loc_00E2C56F: mov var_88, 00000007h
  loc_00E2C579: mov var_D8, ecx
  loc_00E2C57F: mov ecx, [ecx]
  loc_00E2C581: mov [edx], eax
  loc_00E2C583: mov eax, var_8C
  loc_00E2C589: mov [edx+00000004h], eax
  loc_00E2C58C: mov eax, var_88
  loc_00E2C592: mov [edx+00000008h], eax
  loc_00E2C595: mov eax, var_84
  loc_00E2C59B: mov [edx+0000000Ch], eax
  loc_00E2C59E: mov edx, var_38
  loc_00E2C5A1: push edx
  loc_00E2C5A2: call [ecx+00000028h]
  loc_00E2C5A5: test eax, eax
  loc_00E2C5A7: fnclex
  loc_00E2C5A9: jge 00E2C5BCh
  loc_00E2C5AB: mov ecx, var_D8
  loc_00E2C5B1: push 00000028h
  loc_00E2C5B3: push 006E09E8h
  loc_00E2C5B8: push ecx
  loc_00E2C5B9: push eax
  loc_00E2C5BA: call ebx
  loc_00E2C5BC: mov ecx, var_3C
  loc_00E2C5BF: mov eax, 00000002h
  loc_00E2C5C4: mov var_98, 00000000h
  loc_00E2C5CE: mov var_A0, eax
  loc_00E2C5D4: mov edx, [ecx]
  loc_00E2C5D6: sub esp, 00000010h
  loc_00E2C5D9: mov var_E0, ecx
  loc_00E2C5DF: mov ecx, esp
  loc_00E2C5E1: mov [ecx], eax
  loc_00E2C5E3: mov eax, var_9C
  loc_00E2C5E9: mov [ecx+00000004h], eax
  loc_00E2C5EC: mov eax, var_98
  loc_00E2C5F2: mov [ecx+00000008h], eax
  loc_00E2C5F5: mov eax, var_94
  loc_00E2C5FB: mov [ecx+0000000Ch], eax
  loc_00E2C5FE: mov ecx, var_3C
  loc_00E2C601: push ecx
  loc_00E2C602: call [edx+00000038h]
  loc_00E2C605: test eax, eax
  loc_00E2C607: fnclex
  loc_00E2C609: jge 00E2C61Ch
  loc_00E2C60B: mov edx, var_E0
  loc_00E2C611: push 00000038h
  loc_00E2C613: push 006E09F8h
  loc_00E2C618: push edx
  loc_00E2C619: push eax
  loc_00E2C61A: call ebx
  loc_00E2C61C: lea eax, var_3C
  loc_00E2C61F: lea ecx, var_38
  loc_00E2C622: push eax
  loc_00E2C623: lea edx, var_34
  loc_00E2C626: push ecx
  loc_00E2C627: lea eax, var_30
  loc_00E2C62A: push edx
  loc_00E2C62B: push eax
  loc_00E2C62C: push 00000004h
  loc_00E2C62E: call [00401048h] ; __vbaFreeObjList
  loc_00E2C634: add esp, 00000014h
  loc_00E2C637: lea ecx, var_50
  loc_00E2C63A: call [00401024h] ; __vbaFreeVar
  loc_00E2C640: mov ecx, [esi]
  loc_00E2C642: push 006DCBD8h
  loc_00E2C647: push 00000000h
  loc_00E2C649: push 00000018h
  loc_00E2C64B: push esi
  loc_00E2C64C: call [ecx+000004A8h]
  loc_00E2C652: lea edx, var_30
  loc_00E2C655: push eax
  loc_00E2C656: push edx
  loc_00E2C657: call edi
  loc_00E2C659: push eax
  loc_00E2C65A: lea eax, var_50
  loc_00E2C65D: push eax
  loc_00E2C65E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E2C664: add esp, 00000010h
  loc_00E2C667: push eax
  loc_00E2C668: call [00401120h] ; __vbaCastObjVar
  loc_00E2C66E: lea ecx, var_34
  loc_00E2C671: push eax
  loc_00E2C672: push ecx
  loc_00E2C673: call edi
  loc_00E2C675: mov ecx, 0000000Ah
  loc_00E2C67A: mov var_98, 80020004h
  loc_00E2C684: mov var_A0, ecx
  loc_00E2C68A: mov var_88, 80020004h
  loc_00E2C694: mov var_90, ecx
  loc_00E2C69A: mov edx, [eax]
  loc_00E2C69C: sub esp, 00000010h
  loc_00E2C69F: mov var_D0, eax
  loc_00E2C6A5: mov eax, esp
  loc_00E2C6A7: sub esp, 00000010h
  loc_00E2C6AA: mov [eax], ecx
  loc_00E2C6AC: mov ecx, var_9C
  loc_00E2C6B2: mov [eax+00000004h], ecx
  loc_00E2C6B5: mov ecx, var_98
  loc_00E2C6BB: mov [eax+00000008h], ecx
  loc_00E2C6BE: mov ecx, var_94
  loc_00E2C6C4: mov [eax+0000000Ch], ecx
  loc_00E2C6C7: mov ecx, var_90
  loc_00E2C6CD: mov eax, esp
  loc_00E2C6CF: mov [eax], ecx
  loc_00E2C6D1: mov ecx, var_8C
  loc_00E2C6D7: mov [eax+00000004h], ecx
  loc_00E2C6DA: mov ecx, var_88
  loc_00E2C6E0: mov [eax+00000008h], ecx
  loc_00E2C6E3: mov ecx, var_84
  loc_00E2C6E9: mov [eax+0000000Ch], ecx
  loc_00E2C6EC: mov eax, var_D0
  loc_00E2C6F2: push eax
  loc_00E2C6F3: call [edx+000000ACh]
  loc_00E2C6F9: test eax, eax
  loc_00E2C6FB: fnclex
  loc_00E2C6FD: jge 00E2C713h
  loc_00E2C6FF: mov ecx, var_D0
  loc_00E2C705: push 000000ACh
  loc_00E2C70A: push 006DCBE8h
  loc_00E2C70F: push ecx
  loc_00E2C710: push eax
  loc_00E2C711: call ebx
  loc_00E2C713: lea edx, var_34
  loc_00E2C716: lea eax, var_30
  loc_00E2C719: push edx
  loc_00E2C71A: push eax
  loc_00E2C71B: push 00000002h
  loc_00E2C71D: call [00401048h] ; __vbaFreeObjList
  loc_00E2C723: add esp, 0000000Ch
  loc_00E2C726: lea ecx, var_50
  loc_00E2C729: call [00401024h] ; __vbaFreeVar
  loc_00E2C72F: mov ecx, [esi]
  loc_00E2C731: push 00000000h
  loc_00E2C733: push 00000019h
  loc_00E2C735: push esi
  loc_00E2C736: call [ecx+000004A8h]
  loc_00E2C73C: lea edx, var_30
  loc_00E2C73F: push eax
  loc_00E2C740: push edx
  loc_00E2C741: call edi
  loc_00E2C743: push eax
  loc_00E2C744: call [00401030h] ; __vbaLateIdCall
  loc_00E2C74A: add esp, 0000000Ch
  loc_00E2C74D: lea ecx, var_30
  loc_00E2C750: call [00401254h] ; __vbaFreeObj
  loc_00E2C756: jmp 00E2C7E5h
  loc_00E2C75B: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E2C761: mov ecx, 80020004h
  loc_00E2C766: mov var_78, ecx
  loc_00E2C769: mov eax, 0000000Ah
  loc_00E2C76E: mov var_68, ecx
  loc_00E2C771: mov edi, 00000008h
  loc_00E2C776: lea edx, var_A0
  loc_00E2C77C: lea ecx, var_60
  loc_00E2C77F: mov var_80, eax
  loc_00E2C782: mov var_70, eax
  loc_00E2C785: mov var_98, 006E0C30h ; "INFO 001"
  loc_00E2C78F: mov var_A0, edi
  loc_00E2C795: call __vbaVarDup
  loc_00E2C797: mov var_88, 006E0BD4h ; "Data Sudah Ada !"
  loc_00E2C7A1: lea edx, var_90
  loc_00E2C7A7: lea ecx, var_50
  loc_00E2C7AA: mov var_90, edi
  loc_00E2C7B0: call __vbaVarDup
  loc_00E2C7B2: lea eax, var_80
  loc_00E2C7B5: lea ecx, var_70
  loc_00E2C7B8: push eax
  loc_00E2C7B9: lea edx, var_60
  loc_00E2C7BC: push ecx
  loc_00E2C7BD: push edx
  loc_00E2C7BE: lea eax, var_50
  loc_00E2C7C1: push 00000040h
  loc_00E2C7C3: push eax
  loc_00E2C7C4: call [004010A8h] ; rtcMsgBox
  loc_00E2C7CA: lea ecx, var_80
  loc_00E2C7CD: lea edx, var_70
  loc_00E2C7D0: push ecx
  loc_00E2C7D1: lea eax, var_60
  loc_00E2C7D4: push edx
  loc_00E2C7D5: lea ecx, var_50
  loc_00E2C7D8: push eax
  loc_00E2C7D9: push ecx
  loc_00E2C7DA: push 00000004h
  loc_00E2C7DC: call [00401038h] ; __vbaFreeVarList
  loc_00E2C7E2: add esp, 00000014h
  loc_00E2C7E5: mov var_4, 00000000h
  loc_00E2C7EC: fwait
  loc_00E2C7ED: push 00E2C846h
  loc_00E2C7F2: jmp 00E2C83Ch
  loc_00E2C7F4: lea edx, var_2C
  loc_00E2C7F7: lea eax, var_28
  loc_00E2C7FA: push edx
  loc_00E2C7FB: push eax
  loc_00E2C7FC: push 00000002h
  loc_00E2C7FE: call [004011BCh] ; __vbaFreeStrList
  loc_00E2C804: lea ecx, var_40
  loc_00E2C807: lea edx, var_3C
  loc_00E2C80A: push ecx
  loc_00E2C80B: lea eax, var_38
  loc_00E2C80E: push edx
  loc_00E2C80F: lea ecx, var_34
  loc_00E2C812: push eax
  loc_00E2C813: lea edx, var_30
  loc_00E2C816: push ecx
  loc_00E2C817: push edx
  loc_00E2C818: push 00000005h
  loc_00E2C81A: call [00401048h] ; __vbaFreeObjList
  loc_00E2C820: lea eax, var_80
  loc_00E2C823: lea ecx, var_70
  loc_00E2C826: push eax
  loc_00E2C827: lea edx, var_60
  loc_00E2C82A: push ecx
  loc_00E2C82B: lea eax, var_50
  loc_00E2C82E: push edx
  loc_00E2C82F: push eax
  loc_00E2C830: push 00000004h
  loc_00E2C832: call [00401038h] ; __vbaFreeVarList
  loc_00E2C838: add esp, 00000038h
  loc_00E2C83B: ret
  loc_00E2C83C: lea ecx, var_24
  loc_00E2C83F: call [00401024h] ; __vbaFreeVar
  loc_00E2C845: ret
  loc_00E2C846: mov eax, Me
  loc_00E2C849: push eax
  loc_00E2C84A: mov ecx, [eax]
  loc_00E2C84C: call [ecx+00000008h]
  loc_00E2C84F: mov eax, var_4
  loc_00E2C852: mov ecx, var_14
  loc_00E2C855: pop edi
  loc_00E2C856: pop esi
  loc_00E2C857: mov fs:[00000000h], ecx
  loc_00E2C85E: pop ebx
  loc_00E2C85F: mov esp, ebp
  loc_00E2C861: pop ebp
  loc_00E2C862: retn 0004h
End Sub

Private Sub metodes_UnknownEvent_9 'E26870
  loc_00E26870: push ebp
  loc_00E26871: mov ebp, esp
  loc_00E26873: sub esp, 0000000Ch
  loc_00E26876: push 00402836h ; __vbaExceptHandler
  loc_00E2687B: mov eax, fs:[00000000h]
  loc_00E26881: push eax
  loc_00E26882: mov fs:[00000000h], esp
  loc_00E26889: sub esp, 00000034h
  loc_00E2688C: push ebx
  loc_00E2688D: push esi
  loc_00E2688E: push edi
  loc_00E2688F: mov var_C, esp
  loc_00E26892: mov var_8, 00402520h
  loc_00E26899: mov esi, Me
  loc_00E2689C: mov eax, esi
  loc_00E2689E: and eax, 00000001h
  loc_00E268A1: mov var_4, eax
  loc_00E268A4: and esi, FFFFFFFEh
  loc_00E268A7: push esi
  loc_00E268A8: mov Me, esi
  loc_00E268AB: mov ecx, [esi]
  loc_00E268AD: call [ecx+00000004h]
  loc_00E268B0: mov edx, [esi]
  loc_00E268B2: push esi
  loc_00E268B3: mov var_18, 00000000h
  loc_00E268BA: call [edx+00000424h]
  loc_00E268C0: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E268C6: push eax
  loc_00E268C7: lea eax, var_18
  loc_00E268CA: push eax
  loc_00E268CB: call edi
  loc_00E268CD: mov ebx, eax
  loc_00E268CF: push 006E1728h ; "Lunas"
  loc_00E268D4: push ebx
  loc_00E268D5: mov ecx, [ebx]
  loc_00E268D7: call [ecx+00000054h]
  loc_00E268DA: test eax, eax
  loc_00E268DC: fnclex
  loc_00E268DE: jge 00E268EFh
  loc_00E268E0: push 00000054h
  loc_00E268E2: push 006E0410h
  loc_00E268E7: push ebx
  loc_00E268E8: push eax
  loc_00E268E9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E268EF: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E268F5: lea ecx, var_18
  loc_00E268F8: call ebx
  loc_00E268FA: mov edx, [esi]
  loc_00E268FC: push esi
  loc_00E268FD: call [edx+00000320h]
  loc_00E26903: push eax
  loc_00E26904: lea eax, var_18
  loc_00E26907: push eax
  loc_00E26908: call edi
  loc_00E2690A: mov ecx, [eax]
  loc_00E2690C: push 00000000h
  loc_00E2690E: push eax
  loc_00E2690F: mov var_3C, eax
  loc_00E26912: call [ecx+0000009Ch]
  loc_00E26918: test eax, eax
  loc_00E2691A: fnclex
  loc_00E2691C: jge 00E26933h
  loc_00E2691E: mov edx, var_3C
  loc_00E26921: push 0000009Ch
  loc_00E26926: push 006DCAD0h
  loc_00E2692B: push edx
  loc_00E2692C: push eax
  loc_00E2692D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E26933: lea ecx, var_18
  loc_00E26936: call ebx
  loc_00E26938: mov eax, [esi]
  loc_00E2693A: push esi
  loc_00E2693B: call [eax+0000034Ch]
  loc_00E26941: lea ecx, var_18
  loc_00E26944: push eax
  loc_00E26945: push ecx
  loc_00E26946: call edi
  loc_00E26948: mov edx, [eax]
  loc_00E2694A: push FFFFFFFFh
  loc_00E2694C: push eax
  loc_00E2694D: mov var_3C, eax
  loc_00E26950: call [edx+0000009Ch]
  loc_00E26956: test eax, eax
  loc_00E26958: fnclex
  loc_00E2695A: jge 00E26971h
  loc_00E2695C: mov ecx, var_3C
  loc_00E2695F: push 0000009Ch
  loc_00E26964: push 006DCAD0h
  loc_00E26969: push ecx
  loc_00E2696A: push eax
  loc_00E2696B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E26971: lea ecx, var_18
  loc_00E26974: call ebx
  loc_00E26976: mov edx, [esi]
  loc_00E26978: push esi
  loc_00E26979: call [edx+0000037Ch]
  loc_00E2697F: push eax
  loc_00E26980: lea eax, var_18
  loc_00E26983: push eax
  loc_00E26984: call edi
  loc_00E26986: mov ecx, [eax]
  loc_00E26988: push FFFFFFFFh
  loc_00E2698A: push eax
  loc_00E2698B: mov var_3C, eax
  loc_00E2698E: call [ecx+0000009Ch]
  loc_00E26994: test eax, eax
  loc_00E26996: fnclex
  loc_00E26998: jge 00E269AFh
  loc_00E2699A: mov edx, var_3C
  loc_00E2699D: push 0000009Ch
  loc_00E269A2: push 006DCAD0h
  loc_00E269A7: push edx
  loc_00E269A8: push eax
  loc_00E269A9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E269AF: lea ecx, var_18
  loc_00E269B2: call ebx
  loc_00E269B4: mov eax, [esi]
  loc_00E269B6: push esi
  loc_00E269B7: call [eax+00000398h]
  loc_00E269BD: lea ecx, var_18
  loc_00E269C0: push eax
  loc_00E269C1: push ecx
  loc_00E269C2: call edi
  loc_00E269C4: mov edx, [eax]
  loc_00E269C6: push 00000000h
  loc_00E269C8: push eax
  loc_00E269C9: mov var_3C, eax
  loc_00E269CC: call [edx+0000009Ch]
  loc_00E269D2: test eax, eax
  loc_00E269D4: fnclex
  loc_00E269D6: jge 00E269EDh
  loc_00E269D8: mov ecx, var_3C
  loc_00E269DB: push 0000009Ch
  loc_00E269E0: push 006DCAD0h
  loc_00E269E5: push ecx
  loc_00E269E6: push eax
  loc_00E269E7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E269ED: lea ecx, var_18
  loc_00E269F0: call ebx
  loc_00E269F2: mov edx, [esi]
  loc_00E269F4: push esi
  loc_00E269F5: call [edx+00000378h]
  loc_00E269FB: push eax
  loc_00E269FC: lea eax, var_18
  loc_00E269FF: push eax
  loc_00E26A00: call edi
  loc_00E26A02: mov ecx, [eax]
  loc_00E26A04: push eax
  loc_00E26A05: mov var_3C, eax
  loc_00E26A08: call [ecx+00000204h]
  loc_00E26A0E: test eax, eax
  loc_00E26A10: fnclex
  loc_00E26A12: jge 00E26A29h
  loc_00E26A14: mov edx, var_3C
  loc_00E26A17: push 00000204h
  loc_00E26A1C: push 006DCB70h
  loc_00E26A21: push edx
  loc_00E26A22: push eax
  loc_00E26A23: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E26A29: lea ecx, var_18
  loc_00E26A2C: call ebx
  loc_00E26A2E: sub esp, 00000010h
  loc_00E26A31: mov ecx, 0000000Bh
  loc_00E26A36: mov edx, esp
  loc_00E26A38: or eax, FFFFFFFFh
  loc_00E26A3B: push 8001000Dh
  loc_00E26A40: push esi
  loc_00E26A41: mov [edx], ecx
  loc_00E26A43: mov ecx, var_24
  loc_00E26A46: mov [edx+00000004h], ecx
  loc_00E26A49: mov ecx, [esi]
  loc_00E26A4B: mov [edx+00000008h], eax
  loc_00E26A4E: mov eax, var_1C
  loc_00E26A51: mov [edx+0000000Ch], eax
  loc_00E26A54: call [ecx+00000408h]
  loc_00E26A5A: lea edx, var_18
  loc_00E26A5D: push eax
  loc_00E26A5E: push edx
  loc_00E26A5F: call edi
  loc_00E26A61: push eax
  loc_00E26A62: call [00401238h] ; __vbaLateIdSt
  loc_00E26A68: lea ecx, var_18
  loc_00E26A6B: call ebx
  loc_00E26A6D: sub esp, 00000010h
  loc_00E26A70: mov ecx, 0000000Bh
  loc_00E26A75: mov edx, esp
  loc_00E26A77: or eax, FFFFFFFFh
  loc_00E26A7A: push 8001000Dh
  loc_00E26A7F: push esi
  loc_00E26A80: mov [edx], ecx
  loc_00E26A82: mov ecx, var_24
  loc_00E26A85: mov [edx+00000004h], ecx
  loc_00E26A88: mov ecx, [esi]
  loc_00E26A8A: mov [edx+00000008h], eax
  loc_00E26A8D: mov eax, var_1C
  loc_00E26A90: mov [edx+0000000Ch], eax
  loc_00E26A93: call [ecx+00000410h]
  loc_00E26A99: lea edx, var_18
  loc_00E26A9C: push eax
  loc_00E26A9D: push edx
  loc_00E26A9E: call edi
  loc_00E26AA0: push eax
  loc_00E26AA1: call [00401238h] ; __vbaLateIdSt
  loc_00E26AA7: lea ecx, var_18
  loc_00E26AAA: call ebx
  loc_00E26AAC: sub esp, 00000010h
  loc_00E26AAF: mov ecx, 0000000Bh
  loc_00E26AB4: mov edx, esp
  loc_00E26AB6: or eax, FFFFFFFFh
  loc_00E26AB9: push 8001000Dh
  loc_00E26ABE: push esi
  loc_00E26ABF: mov [edx], ecx
  loc_00E26AC1: mov ecx, var_24
  loc_00E26AC4: mov [edx+00000004h], ecx
  loc_00E26AC7: mov ecx, [esi]
  loc_00E26AC9: mov [edx+00000008h], eax
  loc_00E26ACC: mov eax, var_1C
  loc_00E26ACF: mov [edx+0000000Ch], eax
  loc_00E26AD2: call [ecx+00000404h]
  loc_00E26AD8: lea edx, var_18
  loc_00E26ADB: push eax
  loc_00E26ADC: push edx
  loc_00E26ADD: call edi
  loc_00E26ADF: push eax
  loc_00E26AE0: call [00401238h] ; __vbaLateIdSt
  loc_00E26AE6: lea ecx, var_18
  loc_00E26AE9: call ebx
  loc_00E26AEB: or eax, FFFFFFFFh
  loc_00E26AEE: sub esp, 00000010h
  loc_00E26AF1: mov edx, esp
  loc_00E26AF3: mov ecx, 0000000Bh
  loc_00E26AF8: push 8001000Dh
  loc_00E26AFD: push esi
  loc_00E26AFE: mov [edx], ecx
  loc_00E26B00: mov ecx, var_24
  loc_00E26B03: mov [edx+00000004h], ecx
  loc_00E26B06: mov ecx, [esi]
  loc_00E26B08: mov [edx+00000008h], eax
  loc_00E26B0B: mov eax, var_1C
  loc_00E26B0E: mov [edx+0000000Ch], eax
  loc_00E26B11: call [ecx+0000040Ch]
  loc_00E26B17: lea edx, var_18
  loc_00E26B1A: push eax
  loc_00E26B1B: push edx
  loc_00E26B1C: call edi
  loc_00E26B1E: push eax
  loc_00E26B1F: call [00401238h] ; __vbaLateIdSt
  loc_00E26B25: lea ecx, var_18
  loc_00E26B28: call ebx
  loc_00E26B2A: sub esp, 00000010h
  loc_00E26B2D: mov ecx, 0000000Bh
  loc_00E26B32: mov edx, esp
  loc_00E26B34: or eax, FFFFFFFFh
  loc_00E26B37: push 80010007h
  loc_00E26B3C: push esi
  loc_00E26B3D: mov [edx], ecx
  loc_00E26B3F: mov ecx, var_24
  loc_00E26B42: mov [edx+00000004h], ecx
  loc_00E26B45: mov ecx, [esi]
  loc_00E26B47: mov [edx+00000008h], eax
  loc_00E26B4A: mov eax, var_1C
  loc_00E26B4D: mov [edx+0000000Ch], eax
  loc_00E26B50: call [ecx+00000458h]
  loc_00E26B56: lea edx, var_18
  loc_00E26B59: push eax
  loc_00E26B5A: push edx
  loc_00E26B5B: call edi
  loc_00E26B5D: push eax
  loc_00E26B5E: call [00401238h] ; __vbaLateIdSt
  loc_00E26B64: lea ecx, var_18
  loc_00E26B67: call ebx
  loc_00E26B69: mov eax, [esi]
  loc_00E26B6B: push esi
  loc_00E26B6C: call [eax+00000378h]
  loc_00E26B72: lea ecx, var_18
  loc_00E26B75: push eax
  loc_00E26B76: push ecx
  loc_00E26B77: call edi
  loc_00E26B79: mov esi, eax
  loc_00E26B7B: push esi
  loc_00E26B7C: mov edx, [esi]
  loc_00E26B7E: call [edx+00000204h]
  loc_00E26B84: test eax, eax
  loc_00E26B86: fnclex
  loc_00E26B88: jge 00E26B9Ch
  loc_00E26B8A: push 00000204h
  loc_00E26B8F: push 006DCB70h
  loc_00E26B94: push esi
  loc_00E26B95: push eax
  loc_00E26B96: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E26B9C: lea ecx, var_18
  loc_00E26B9F: call ebx
  loc_00E26BA1: mov var_4, 00000000h
  loc_00E26BA8: push 00E26BBAh
  loc_00E26BAD: jmp 00E26BB9h
  loc_00E26BAF: lea ecx, var_18
  loc_00E26BB2: call [00401254h] ; __vbaFreeObj
  loc_00E26BB8: ret
  loc_00E26BB9: ret
  loc_00E26BBA: mov eax, Me
  loc_00E26BBD: push eax
  loc_00E26BBE: mov ecx, [eax]
  loc_00E26BC0: call [ecx+00000008h]
  loc_00E26BC3: mov eax, var_4
  loc_00E26BC6: mov ecx, var_14
  loc_00E26BC9: pop edi
  loc_00E26BCA: pop esi
  loc_00E26BCB: mov fs:[00000000h], ecx
  loc_00E26BD2: pop ebx
  loc_00E26BD3: mov esp, ebp
  loc_00E26BD5: pop ebp
  loc_00E26BD6: retn 0004h
End Sub

Private Sub jcbutton2_UnknownEvent_9 'E266E0
  loc_00E266E0: push ebp
  loc_00E266E1: mov ebp, esp
  loc_00E266E3: sub esp, 0000000Ch
  loc_00E266E6: push 00402836h ; __vbaExceptHandler
  loc_00E266EB: mov eax, fs:[00000000h]
  loc_00E266F1: push eax
  loc_00E266F2: mov fs:[00000000h], esp
  loc_00E266F9: sub esp, 00000034h
  loc_00E266FC: push ebx
  loc_00E266FD: push esi
  loc_00E266FE: push edi
  loc_00E266FF: mov var_C, esp
  loc_00E26702: mov var_8, 00402510h
  loc_00E26709: mov esi, Me
  loc_00E2670C: mov eax, esi
  loc_00E2670E: and eax, 00000001h
  loc_00E26711: mov var_4, eax
  loc_00E26714: and esi, FFFFFFFEh
  loc_00E26717: push esi
  loc_00E26718: mov Me, esi
  loc_00E2671B: mov ecx, [esi]
  loc_00E2671D: call [ecx+00000004h]
  loc_00E26720: mov edx, [esi]
  loc_00E26722: push esi
  loc_00E26723: mov var_18, 00000000h
  loc_00E2672A: call [edx+00000308h]
  loc_00E26730: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E26736: push eax
  loc_00E26737: lea eax, var_18
  loc_00E2673A: push eax
  loc_00E2673B: call ebx
  loc_00E2673D: mov edi, eax
  loc_00E2673F: push 00000000h
  loc_00E26741: push edi
  loc_00E26742: mov ecx, [edi]
  loc_00E26744: call [ecx+0000009Ch]
  loc_00E2674A: test eax, eax
  loc_00E2674C: fnclex
  loc_00E2674E: jge 00E26762h
  loc_00E26750: push 0000009Ch
  loc_00E26755: push 006DCAD0h
  loc_00E2675A: push edi
  loc_00E2675B: push eax
  loc_00E2675C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E26762: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E26768: lea ecx, var_18
  loc_00E2676B: call edi
  loc_00E2676D: mov edx, [esi]
  loc_00E2676F: push esi
  loc_00E26770: call [edx+0000034Ch]
  loc_00E26776: push eax
  loc_00E26777: lea eax, var_18
  loc_00E2677A: push eax
  loc_00E2677B: call ebx
  loc_00E2677D: mov ecx, [eax]
  loc_00E2677F: push FFFFFFFFh
  loc_00E26781: push eax
  loc_00E26782: mov var_3C, eax
  loc_00E26785: call [ecx+0000009Ch]
  loc_00E2678B: test eax, eax
  loc_00E2678D: fnclex
  loc_00E2678F: jge 00E267A6h
  loc_00E26791: mov edx, var_3C
  loc_00E26794: push 0000009Ch
  loc_00E26799: push 006DCAD0h
  loc_00E2679E: push edx
  loc_00E2679F: push eax
  loc_00E267A0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E267A6: lea ecx, var_18
  loc_00E267A9: call edi
  loc_00E267AB: sub esp, 00000010h
  loc_00E267AE: mov ecx, 0000000Bh
  loc_00E267B3: mov edx, esp
  loc_00E267B5: or eax, FFFFFFFFh
  loc_00E267B8: push 80010007h
  loc_00E267BD: push esi
  loc_00E267BE: mov [edx], ecx
  loc_00E267C0: mov ecx, var_24
  loc_00E267C3: mov [edx+00000004h], ecx
  loc_00E267C6: mov ecx, [esi]
  loc_00E267C8: mov [edx+00000008h], eax
  loc_00E267CB: mov eax, var_1C
  loc_00E267CE: mov [edx+0000000Ch], eax
  loc_00E267D1: call [ecx+000004ACh]
  loc_00E267D7: lea edx, var_18
  loc_00E267DA: push eax
  loc_00E267DB: push edx
  loc_00E267DC: call ebx
  loc_00E267DE: push eax
  loc_00E267DF: call [00401238h] ; __vbaLateIdSt
  loc_00E267E5: lea ecx, var_18
  loc_00E267E8: call edi
  loc_00E267EA: sub esp, 00000010h
  loc_00E267ED: mov ecx, 0000000Bh
  loc_00E267F2: mov edx, esp
  loc_00E267F4: or eax, FFFFFFFFh
  loc_00E267F7: push 8001000Dh
  loc_00E267FC: push esi
  loc_00E267FD: mov [edx], ecx
  loc_00E267FF: mov ecx, var_24
  loc_00E26802: mov [edx+00000004h], ecx
  loc_00E26805: mov ecx, [esi]
  loc_00E26807: mov [edx+00000008h], eax
  loc_00E2680A: mov eax, var_1C
  loc_00E2680D: mov [edx+0000000Ch], eax
  loc_00E26810: call [ecx+00000410h]
  loc_00E26816: lea edx, var_18
  loc_00E26819: push eax
  loc_00E2681A: push edx
  loc_00E2681B: call ebx
  loc_00E2681D: push eax
  loc_00E2681E: call [00401238h] ; __vbaLateIdSt
  loc_00E26824: lea ecx, var_18
  loc_00E26827: call edi
  loc_00E26829: mov var_4, 00000000h
  loc_00E26830: push 00E26842h
  loc_00E26835: jmp 00E26841h
  loc_00E26837: lea ecx, var_18
  loc_00E2683A: call [00401254h] ; __vbaFreeObj
  loc_00E26840: ret
  loc_00E26841: ret
  loc_00E26842: mov eax, Me
  loc_00E26845: push eax
  loc_00E26846: mov ecx, [eax]
  loc_00E26848: call [ecx+00000008h]
  loc_00E2684B: mov eax, var_4
  loc_00E2684E: mov ecx, var_14
  loc_00E26851: pop edi
  loc_00E26852: pop esi
  loc_00E26853: mov fs:[00000000h], ecx
  loc_00E2685A: pop ebx
  loc_00E2685B: mov esp, ebp
  loc_00E2685D: pop ebp
  loc_00E2685E: retn 0004h
End Sub

Private Sub tgl_UnknownEvent_9 'E2C870
  loc_00E2C870: push ebp
  loc_00E2C871: mov ebp, esp
  loc_00E2C873: sub esp, 0000000Ch
  loc_00E2C876: push 00402836h ; __vbaExceptHandler
  loc_00E2C87B: mov eax, fs:[00000000h]
  loc_00E2C881: push eax
  loc_00E2C882: mov fs:[00000000h], esp
  loc_00E2C889: sub esp, 00000048h
  loc_00E2C88C: push ebx
  loc_00E2C88D: push esi
  loc_00E2C88E: push edi
  loc_00E2C88F: mov var_C, esp
  loc_00E2C892: mov var_8, 00402550h
  loc_00E2C899: mov esi, Me
  loc_00E2C89C: mov eax, esi
  loc_00E2C89E: and eax, 00000001h
  loc_00E2C8A1: mov var_4, eax
  loc_00E2C8A4: and esi, FFFFFFFEh
  loc_00E2C8A7: push esi
  loc_00E2C8A8: mov Me, esi
  loc_00E2C8AB: mov ecx, [esi]
  loc_00E2C8AD: call [ecx+00000004h]
  loc_00E2C8B0: mov edx, [esi]
  loc_00E2C8B2: xor eax, eax
  loc_00E2C8B4: push esi
  loc_00E2C8B5: mov var_18, eax
  loc_00E2C8B8: mov var_1C, eax
  loc_00E2C8BB: mov var_2C, eax
  loc_00E2C8BE: call [edx+00000304h]
  loc_00E2C8C4: push eax
  loc_00E2C8C5: lea eax, var_1C
  loc_00E2C8C8: push eax
  loc_00E2C8C9: call [004010ACh] ; __vbaObjSet
  loc_00E2C8CF: lea ecx, var_2C
  loc_00E2C8D2: mov edi, eax
  loc_00E2C8D4: push ecx
  loc_00E2C8D5: call [004011D8h] ; rtcGetDateVar
  loc_00E2C8DB: mov ebx, [edi]
  loc_00E2C8DD: lea edx, var_2C
  loc_00E2C8E0: lea eax, var_18
  loc_00E2C8E3: push edx
  loc_00E2C8E4: push eax
  loc_00E2C8E5: call [00401180h] ; __vbaStrVarVal
  loc_00E2C8EB: push eax
  loc_00E2C8EC: push edi
  loc_00E2C8ED: call [ebx+000000A4h]
  loc_00E2C8F3: test eax, eax
  loc_00E2C8F5: fnclex
  loc_00E2C8F7: jge 00E2C90Bh
  loc_00E2C8F9: push 000000A4h
  loc_00E2C8FE: push 006DCB70h
  loc_00E2C903: push edi
  loc_00E2C904: push eax
  loc_00E2C905: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2C90B: lea ecx, var_18
  loc_00E2C90E: call [00401258h] ; __vbaFreeStr
  loc_00E2C914: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E2C91A: lea ecx, var_1C
  loc_00E2C91D: call edi
  loc_00E2C91F: lea ecx, var_2C
  loc_00E2C922: call [00401024h] ; __vbaFreeVar
  loc_00E2C928: sub esp, 00000010h
  loc_00E2C92B: mov ecx, 0000000Bh
  loc_00E2C930: mov edx, esp
  loc_00E2C932: xor eax, eax
  loc_00E2C934: push 80010007h
  loc_00E2C939: push esi
  loc_00E2C93A: mov [edx], ecx
  loc_00E2C93C: mov ecx, var_38
  loc_00E2C93F: mov [edx+00000004h], ecx
  loc_00E2C942: mov ecx, [esi]
  loc_00E2C944: mov [edx+00000008h], eax
  loc_00E2C947: mov eax, var_30
  loc_00E2C94A: mov [edx+0000000Ch], eax
  loc_00E2C94D: call [ecx+00000458h]
  loc_00E2C953: lea edx, var_1C
  loc_00E2C956: push eax
  loc_00E2C957: push edx
  loc_00E2C958: call [004010ACh] ; __vbaObjSet
  loc_00E2C95E: push eax
  loc_00E2C95F: call [00401238h] ; __vbaLateIdSt
  loc_00E2C965: lea ecx, var_1C
  loc_00E2C968: call edi
  loc_00E2C96A: mov var_4, 00000000h
  loc_00E2C971: push 00E2C995h
  loc_00E2C976: jmp 00E2C994h
  loc_00E2C978: lea ecx, var_18
  loc_00E2C97B: call [00401258h] ; __vbaFreeStr
  loc_00E2C981: lea ecx, var_1C
  loc_00E2C984: call [00401254h] ; __vbaFreeObj
  loc_00E2C98A: lea ecx, var_2C
  loc_00E2C98D: call [00401024h] ; __vbaFreeVar
  loc_00E2C993: ret
  loc_00E2C994: ret
  loc_00E2C995: mov eax, Me
  loc_00E2C998: push eax
  loc_00E2C999: mov ecx, [eax]
  loc_00E2C99B: call [ecx+00000008h]
  loc_00E2C99E: mov eax, var_4
  loc_00E2C9A1: mov ecx, var_14
  loc_00E2C9A4: pop edi
  loc_00E2C9A5: pop esi
  loc_00E2C9A6: mov fs:[00000000h], ecx
  loc_00E2C9AD: pop ebx
  loc_00E2C9AE: mov esp, ebp
  loc_00E2C9B0: pop ebp
  loc_00E2C9B1: retn 0004h
End Sub

Private Sub jcbutton1_UnknownEvent_9 'E26030
  loc_00E26030: push ebp
  loc_00E26031: mov ebp, esp
  loc_00E26033: sub esp, 0000000Ch
  loc_00E26036: push 00402836h ; __vbaExceptHandler
  loc_00E2603B: mov eax, fs:[00000000h]
  loc_00E26041: push eax
  loc_00E26042: mov fs:[00000000h], esp
  loc_00E26049: sub esp, 000000B4h
  loc_00E2604F: push ebx
  loc_00E26050: push esi
  loc_00E26051: push edi
  loc_00E26052: mov var_C, esp
  loc_00E26055: mov var_8, 00402500h
  loc_00E2605C: mov esi, Me
  loc_00E2605F: mov eax, esi
  loc_00E26061: and eax, 00000001h
  loc_00E26064: mov var_4, eax
  loc_00E26067: and esi, FFFFFFFEh
  loc_00E2606A: push esi
  loc_00E2606B: mov Me, esi
  loc_00E2606E: mov ecx, [esi]
  loc_00E26070: call [ecx+00000004h]
  loc_00E26073: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E26079: xor ebx, ebx
  loc_00E2607B: mov ecx, 0000000Ah
  loc_00E26080: mov var_58, ebx
  loc_00E26083: mov var_68, ebx
  loc_00E26086: mov eax, 80020004h
  loc_00E2608B: mov var_68, ecx
  loc_00E2608E: mov var_58, ecx
  loc_00E26091: mov var_88, ebx
  loc_00E26097: lea edx, var_88
  loc_00E2609D: lea ecx, var_48
  loc_00E260A0: mov var_24, ebx
  loc_00E260A3: mov var_28, ebx
  loc_00E260A6: mov var_38, ebx
  loc_00E260A9: mov var_48, ebx
  loc_00E260AC: mov var_78, ebx
  loc_00E260AF: mov var_B8, ebx
  loc_00E260B5: mov var_60, eax
  loc_00E260B8: mov var_50, eax
  loc_00E260BB: mov var_80, 006E1684h ; "Ask INFO"
  loc_00E260C2: mov var_88, 00000008h
  loc_00E260CC: call edi
  loc_00E260CE: lea edx, var_78
  loc_00E260D1: lea ecx, var_38
  loc_00E260D4: mov var_70, 006E162Ch ; "Apakah Sudah Pernah Mencicil Sebelumnya?"
  loc_00E260DB: mov var_78, 00000008h
  loc_00E260E2: call edi
  loc_00E260E4: lea edx, var_68
  loc_00E260E7: lea eax, var_58
  loc_00E260EA: push edx
  loc_00E260EB: lea ecx, var_48
  loc_00E260EE: push eax
  loc_00E260EF: push ecx
  loc_00E260F0: lea edx, var_38
  loc_00E260F3: push 00000044h
  loc_00E260F5: push edx
  loc_00E260F6: call [004010A8h] ; rtcMsgBox
  loc_00E260FC: lea edx, var_B8
  loc_00E26102: lea ecx, var_24
  loc_00E26105: mov var_B0, eax
  loc_00E2610B: mov var_B8, 00000003h
  loc_00E26115: call [0040101Ch] ; __vbaVarMove
  loc_00E2611B: lea eax, var_68
  loc_00E2611E: lea ecx, var_58
  loc_00E26121: push eax
  loc_00E26122: lea edx, var_48
  loc_00E26125: push ecx
  loc_00E26126: lea eax, var_38
  loc_00E26129: push edx
  loc_00E2612A: push eax
  loc_00E2612B: push 00000004h
  loc_00E2612D: call [00401038h] ; __vbaFreeVarList
  loc_00E26133: add esp, 00000014h
  loc_00E26136: lea ecx, var_24
  loc_00E26139: lea edx, var_78
  loc_00E2613C: mov var_70, 00000006h
  loc_00E26143: push ecx
  loc_00E26144: push edx
  loc_00E26145: mov var_78, 00008003h
  loc_00E2614C: call [00401108h] ; __vbaVarTstEq
  loc_00E26152: test ax, ax
  loc_00E26155: jz 00E262B2h
  loc_00E2615B: mov ecx, 0000000Ah
  loc_00E26160: mov eax, 80020004h
  loc_00E26165: mov var_68, ecx
  loc_00E26168: mov var_58, ecx
  loc_00E2616B: mov ebx, 00000008h
  loc_00E26170: lea edx, var_88
  loc_00E26176: lea ecx, var_48
  loc_00E26179: mov var_60, eax
  loc_00E2617C: mov var_50, eax
  loc_00E2617F: mov var_80, 006E16F0h ; "SMK RK BT PS"
  loc_00E26186: mov var_88, ebx
  loc_00E2618C: call edi
  loc_00E2618E: lea edx, var_78
  loc_00E26191: lea ecx, var_38
  loc_00E26194: mov var_70, 006E169Ch ; "Panggil Data sesuai dengan Kebutuhan !"
  loc_00E2619B: mov var_78, ebx
  loc_00E2619E: call edi
  loc_00E261A0: lea eax, var_68
  loc_00E261A3: lea ecx, var_58
  loc_00E261A6: push eax
  loc_00E261A7: lea edx, var_48
  loc_00E261AA: push ecx
  loc_00E261AB: push edx
  loc_00E261AC: lea eax, var_38
  loc_00E261AF: push 00000020h
  loc_00E261B1: push eax
  loc_00E261B2: call [004010A8h] ; rtcMsgBox
  loc_00E261B8: lea ecx, var_68
  loc_00E261BB: lea edx, var_58
  loc_00E261BE: push ecx
  loc_00E261BF: lea eax, var_48
  loc_00E261C2: push edx
  loc_00E261C3: lea ecx, var_38
  loc_00E261C6: push eax
  loc_00E261C7: push ecx
  loc_00E261C8: push 00000004h
  loc_00E261CA: call [00401038h] ; __vbaFreeVarList
  loc_00E261D0: add esp, 00000004h
  loc_00E261D3: mov ecx, 0000000Bh
  loc_00E261D8: mov edx, esp
  loc_00E261DA: mov var_78, ecx
  loc_00E261DD: or eax, FFFFFFFFh
  loc_00E261E0: push 80010007h
  loc_00E261E5: mov [edx], ecx
  loc_00E261E7: mov ecx, var_74
  loc_00E261EA: mov var_70, eax
  loc_00E261ED: push esi
  loc_00E261EE: mov [edx+00000004h], ecx
  loc_00E261F1: mov ecx, [esi]
  loc_00E261F3: mov [edx+00000008h], eax
  loc_00E261F6: mov eax, var_6C
  loc_00E261F9: mov [edx+0000000Ch], eax
  loc_00E261FC: call [ecx+000003A0h]
  loc_00E26202: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E26208: lea edx, var_28
  loc_00E2620B: push eax
  loc_00E2620C: push edx
  loc_00E2620D: call ebx
  loc_00E2620F: push eax
  loc_00E26210: call [00401238h] ; __vbaLateIdSt
  loc_00E26216: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E2621C: lea ecx, var_28
  loc_00E2621F: call edi
  loc_00E26221: sub esp, 00000010h
  loc_00E26224: mov ecx, 0000000Bh
  loc_00E26229: mov edx, esp
  loc_00E2622B: mov var_78, ecx
  loc_00E2622E: or eax, FFFFFFFFh
  loc_00E26231: push 80010007h
  loc_00E26236: mov [edx], ecx
  loc_00E26238: mov ecx, var_74
  loc_00E2623B: mov var_70, eax
  loc_00E2623E: push esi
  loc_00E2623F: mov [edx+00000004h], ecx
  loc_00E26242: mov ecx, [esi]
  loc_00E26244: mov [edx+00000008h], eax
  loc_00E26247: mov eax, var_6C
  loc_00E2624A: mov [edx+0000000Ch], eax
  loc_00E2624D: call [ecx+00000458h]
  loc_00E26253: push eax
  loc_00E26254: lea edx, var_28
  loc_00E26257: push edx
  loc_00E26258: call ebx
  loc_00E2625A: push eax
  loc_00E2625B: call [00401238h] ; __vbaLateIdSt
  loc_00E26261: lea ecx, var_28
  loc_00E26264: call edi
  loc_00E26266: sub esp, 00000010h
  loc_00E26269: mov ecx, 0000000Bh
  loc_00E2626E: mov edx, esp
  loc_00E26270: mov var_78, ecx
  loc_00E26273: or eax, FFFFFFFFh
  loc_00E26276: push 80010007h
  loc_00E2627B: mov [edx], ecx
  loc_00E2627D: mov ecx, var_74
  loc_00E26280: mov var_70, eax
  loc_00E26283: push esi
  loc_00E26284: mov [edx+00000004h], ecx
  loc_00E26287: mov ecx, [esi]
  loc_00E26289: mov [edx+00000008h], eax
  loc_00E2628C: mov eax, var_6C
  loc_00E2628F: mov [edx+0000000Ch], eax
  loc_00E26292: call [ecx+00000400h]
  loc_00E26298: lea edx, var_28
  loc_00E2629B: push eax
  loc_00E2629C: push edx
  loc_00E2629D: call ebx
  loc_00E2629F: push eax
  loc_00E262A0: call [00401238h] ; __vbaLateIdSt
  loc_00E262A6: lea ecx, var_28
  loc_00E262A9: call edi
  loc_00E262AB: xor eax, eax
  loc_00E262AD: jmp 00E2638Ch
  loc_00E262B2: mov edx, var_74
  loc_00E262B5: sub esp, 00000010h
  loc_00E262B8: mov ecx, esp
  loc_00E262BA: mov eax, 0000000Bh
  loc_00E262BF: mov var_78, eax
  loc_00E262C2: push 80010007h
  loc_00E262C7: mov [ecx], eax
  loc_00E262C9: mov eax, var_6C
  loc_00E262CC: push esi
  loc_00E262CD: mov var_70, ebx
  loc_00E262D0: mov [ecx+00000004h], edx
  loc_00E262D3: mov [ecx+00000008h], ebx
  loc_00E262D6: mov [ecx+0000000Ch], eax
  loc_00E262D9: mov ecx, [esi]
  loc_00E262DB: call [ecx+000003A0h]
  loc_00E262E1: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E262E7: lea edx, var_28
  loc_00E262EA: push eax
  loc_00E262EB: push edx
  loc_00E262EC: call ebx
  loc_00E262EE: push eax
  loc_00E262EF: call [00401238h] ; __vbaLateIdSt
  loc_00E262F5: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E262FB: lea ecx, var_28
  loc_00E262FE: call edi
  loc_00E26300: sub esp, 00000010h
  loc_00E26303: mov ecx, 0000000Bh
  loc_00E26308: mov edx, esp
  loc_00E2630A: mov var_78, ecx
  loc_00E2630D: or eax, FFFFFFFFh
  loc_00E26310: push 80010007h
  loc_00E26315: mov [edx], ecx
  loc_00E26317: mov ecx, var_74
  loc_00E2631A: mov var_70, eax
  loc_00E2631D: push esi
  loc_00E2631E: mov [edx+00000004h], ecx
  loc_00E26321: mov ecx, [esi]
  loc_00E26323: mov [edx+00000008h], eax
  loc_00E26326: mov eax, var_6C
  loc_00E26329: mov [edx+0000000Ch], eax
  loc_00E2632C: call [ecx+00000458h]
  loc_00E26332: lea edx, var_28
  loc_00E26335: push eax
  loc_00E26336: push edx
  loc_00E26337: call ebx
  loc_00E26339: push eax
  loc_00E2633A: call [00401238h] ; __vbaLateIdSt
  loc_00E26340: lea ecx, var_28
  loc_00E26343: call edi
  loc_00E26345: sub esp, 00000010h
  loc_00E26348: mov ecx, 0000000Bh
  loc_00E2634D: mov edx, esp
  loc_00E2634F: mov var_78, ecx
  loc_00E26352: xor eax, eax
  loc_00E26354: push 80010007h
  loc_00E26359: mov [edx], ecx
  loc_00E2635B: mov ecx, var_74
  loc_00E2635E: mov var_70, eax
  loc_00E26361: push esi
  loc_00E26362: mov [edx+00000004h], ecx
  loc_00E26365: mov ecx, [esi]
  loc_00E26367: mov [edx+00000008h], eax
  loc_00E2636A: mov eax, var_6C
  loc_00E2636D: mov [edx+0000000Ch], eax
  loc_00E26370: call [ecx+00000400h]
  loc_00E26376: lea edx, var_28
  loc_00E26379: push eax
  loc_00E2637A: push edx
  loc_00E2637B: call ebx
  loc_00E2637D: push eax
  loc_00E2637E: call [00401238h] ; __vbaLateIdSt
  loc_00E26384: lea ecx, var_28
  loc_00E26387: call edi
  loc_00E26389: or eax, FFFFFFFFh
  loc_00E2638C: sub esp, 00000010h
  loc_00E2638F: mov ecx, 0000000Bh
  loc_00E26394: mov edx, esp
  loc_00E26396: mov var_78, ecx
  loc_00E26399: mov var_70, eax
  loc_00E2639C: push 80010007h
  loc_00E263A1: mov [edx], ecx
  loc_00E263A3: mov ecx, var_74
  loc_00E263A6: push esi
  loc_00E263A7: mov [edx+00000004h], ecx
  loc_00E263AA: mov ecx, [esi]
  loc_00E263AC: mov [edx+00000008h], eax
  loc_00E263AF: mov eax, var_6C
  loc_00E263B2: mov [edx+0000000Ch], eax
  loc_00E263B5: call [ecx+00000408h]
  loc_00E263BB: lea edx, var_28
  loc_00E263BE: push eax
  loc_00E263BF: push edx
  loc_00E263C0: call ebx
  loc_00E263C2: push eax
  loc_00E263C3: call [00401238h] ; __vbaLateIdSt
  loc_00E263C9: lea ecx, var_28
  loc_00E263CC: call edi
  loc_00E263CE: mov eax, [esi]
  loc_00E263D0: push esi
  loc_00E263D1: call [eax+00000424h]
  loc_00E263D7: lea ecx, var_28
  loc_00E263DA: push eax
  loc_00E263DB: push ecx
  loc_00E263DC: call ebx
  loc_00E263DE: mov edx, [eax]
  loc_00E263E0: push 006E1710h ; "Mencicil"
  loc_00E263E5: push eax
  loc_00E263E6: mov var_BC, eax
  loc_00E263EC: call [edx+00000054h]
  loc_00E263EF: test eax, eax
  loc_00E263F1: fnclex
  loc_00E263F3: jge 00E2640Ah
  loc_00E263F5: mov ecx, var_BC
  loc_00E263FB: push 00000054h
  loc_00E263FD: push 006E0410h
  loc_00E26402: push ecx
  loc_00E26403: push eax
  loc_00E26404: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2640A: lea ecx, var_28
  loc_00E2640D: call edi
  loc_00E2640F: mov edx, [esi]
  loc_00E26411: push esi
  loc_00E26412: call [edx+00000320h]
  loc_00E26418: push eax
  loc_00E26419: lea eax, var_28
  loc_00E2641C: push eax
  loc_00E2641D: call ebx
  loc_00E2641F: mov ecx, [eax]
  loc_00E26421: push 00000000h
  loc_00E26423: push eax
  loc_00E26424: mov var_BC, eax
  loc_00E2642A: call [ecx+0000009Ch]
  loc_00E26430: test eax, eax
  loc_00E26432: fnclex
  loc_00E26434: jge 00E2644Eh
  loc_00E26436: mov edx, var_BC
  loc_00E2643C: push 0000009Ch
  loc_00E26441: push 006DCAD0h
  loc_00E26446: push edx
  loc_00E26447: push eax
  loc_00E26448: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2644E: lea ecx, var_28
  loc_00E26451: call edi
  loc_00E26453: mov eax, [esi]
  loc_00E26455: push esi
  loc_00E26456: call [eax+0000034Ch]
  loc_00E2645C: lea ecx, var_28
  loc_00E2645F: push eax
  loc_00E26460: push ecx
  loc_00E26461: call ebx
  loc_00E26463: mov edx, [eax]
  loc_00E26465: push FFFFFFFFh
  loc_00E26467: push eax
  loc_00E26468: mov var_BC, eax
  loc_00E2646E: call [edx+0000009Ch]
  loc_00E26474: test eax, eax
  loc_00E26476: fnclex
  loc_00E26478: jge 00E26492h
  loc_00E2647A: mov ecx, var_BC
  loc_00E26480: push 0000009Ch
  loc_00E26485: push 006DCAD0h
  loc_00E2648A: push ecx
  loc_00E2648B: push eax
  loc_00E2648C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E26492: lea ecx, var_28
  loc_00E26495: call edi
  loc_00E26497: mov edx, [esi]
  loc_00E26499: push esi
  loc_00E2649A: call [edx+0000037Ch]
  loc_00E264A0: push eax
  loc_00E264A1: lea eax, var_28
  loc_00E264A4: push eax
  loc_00E264A5: call ebx
  loc_00E264A7: mov ecx, [eax]
  loc_00E264A9: push 00000000h
  loc_00E264AB: push eax
  loc_00E264AC: mov var_BC, eax
  loc_00E264B2: call [ecx+0000009Ch]
  loc_00E264B8: test eax, eax
  loc_00E264BA: fnclex
  loc_00E264BC: jge 00E264D6h
  loc_00E264BE: mov edx, var_BC
  loc_00E264C4: push 0000009Ch
  loc_00E264C9: push 006DCAD0h
  loc_00E264CE: push edx
  loc_00E264CF: push eax
  loc_00E264D0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E264D6: lea ecx, var_28
  loc_00E264D9: call edi
  loc_00E264DB: mov eax, [esi]
  loc_00E264DD: push esi
  loc_00E264DE: call [eax+00000398h]
  loc_00E264E4: lea ecx, var_28
  loc_00E264E7: push eax
  loc_00E264E8: push ecx
  loc_00E264E9: call ebx
  loc_00E264EB: mov edx, [eax]
  loc_00E264ED: push FFFFFFFFh
  loc_00E264EF: push eax
  loc_00E264F0: mov var_BC, eax
  loc_00E264F6: call [edx+0000009Ch]
  loc_00E264FC: test eax, eax
  loc_00E264FE: fnclex
  loc_00E26500: jge 00E2651Ah
  loc_00E26502: mov ecx, var_BC
  loc_00E26508: push 0000009Ch
  loc_00E2650D: push 006DCAD0h
  loc_00E26512: push ecx
  loc_00E26513: push eax
  loc_00E26514: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2651A: lea ecx, var_28
  loc_00E2651D: call edi
  loc_00E2651F: mov edx, [esi]
  loc_00E26521: push esi
  loc_00E26522: call [edx+00000378h]
  loc_00E26528: push eax
  loc_00E26529: lea eax, var_28
  loc_00E2652C: push eax
  loc_00E2652D: call ebx
  loc_00E2652F: mov ecx, [eax]
  loc_00E26531: push eax
  loc_00E26532: mov var_BC, eax
  loc_00E26538: call [ecx+00000204h]
  loc_00E2653E: test eax, eax
  loc_00E26540: fnclex
  loc_00E26542: jge 00E2655Ch
  loc_00E26544: mov edx, var_BC
  loc_00E2654A: push 00000204h
  loc_00E2654F: push 006DCB70h
  loc_00E26554: push edx
  loc_00E26555: push eax
  loc_00E26556: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2655C: lea ecx, var_28
  loc_00E2655F: call edi
  loc_00E26561: sub esp, 00000010h
  loc_00E26564: mov ecx, 0000000Bh
  loc_00E26569: mov edx, esp
  loc_00E2656B: mov var_78, ecx
  loc_00E2656E: or eax, FFFFFFFFh
  loc_00E26571: push 8001000Dh
  loc_00E26576: mov [edx], ecx
  loc_00E26578: mov ecx, var_74
  loc_00E2657B: mov var_70, eax
  loc_00E2657E: push esi
  loc_00E2657F: mov [edx+00000004h], ecx
  loc_00E26582: mov ecx, [esi]
  loc_00E26584: mov [edx+00000008h], eax
  loc_00E26587: mov eax, var_6C
  loc_00E2658A: mov [edx+0000000Ch], eax
  loc_00E2658D: call [ecx+00000408h]
  loc_00E26593: lea edx, var_28
  loc_00E26596: push eax
  loc_00E26597: push edx
  loc_00E26598: call ebx
  loc_00E2659A: push eax
  loc_00E2659B: call [00401238h] ; __vbaLateIdSt
  loc_00E265A1: lea ecx, var_28
  loc_00E265A4: call edi
  loc_00E265A6: sub esp, 00000010h
  loc_00E265A9: mov ecx, 0000000Bh
  loc_00E265AE: mov edx, esp
  loc_00E265B0: mov var_78, ecx
  loc_00E265B3: or eax, FFFFFFFFh
  loc_00E265B6: push 8001000Dh
  loc_00E265BB: mov [edx], ecx
  loc_00E265BD: mov ecx, var_74
  loc_00E265C0: mov var_70, eax
  loc_00E265C3: push esi
  loc_00E265C4: mov [edx+00000004h], ecx
  loc_00E265C7: mov ecx, [esi]
  loc_00E265C9: mov [edx+00000008h], eax
  loc_00E265CC: mov eax, var_6C
  loc_00E265CF: mov [edx+0000000Ch], eax
  loc_00E265D2: call [ecx+00000410h]
  loc_00E265D8: lea edx, var_28
  loc_00E265DB: push eax
  loc_00E265DC: push edx
  loc_00E265DD: call ebx
  loc_00E265DF: push eax
  loc_00E265E0: call [00401238h] ; __vbaLateIdSt
  loc_00E265E6: lea ecx, var_28
  loc_00E265E9: call edi
  loc_00E265EB: sub esp, 00000010h
  loc_00E265EE: mov ecx, 0000000Bh
  loc_00E265F3: mov edx, esp
  loc_00E265F5: mov var_78, ecx
  loc_00E265F8: or eax, FFFFFFFFh
  loc_00E265FB: push 8001000Dh
  loc_00E26600: mov [edx], ecx
  loc_00E26602: mov ecx, var_74
  loc_00E26605: mov var_70, eax
  loc_00E26608: push esi
  loc_00E26609: mov [edx+00000004h], ecx
  loc_00E2660C: mov ecx, [esi]
  loc_00E2660E: mov [edx+00000008h], eax
  loc_00E26611: mov eax, var_6C
  loc_00E26614: mov [edx+0000000Ch], eax
  loc_00E26617: call [ecx+00000404h]
  loc_00E2661D: lea edx, var_28
  loc_00E26620: push eax
  loc_00E26621: push edx
  loc_00E26622: call ebx
  loc_00E26624: push eax
  loc_00E26625: call [00401238h] ; __vbaLateIdSt
  loc_00E2662B: lea ecx, var_28
  loc_00E2662E: call edi
  loc_00E26630: or eax, FFFFFFFFh
  loc_00E26633: sub esp, 00000010h
  loc_00E26636: mov edx, esp
  loc_00E26638: mov ecx, 0000000Bh
  loc_00E2663D: mov var_70, eax
  loc_00E26640: mov var_78, ecx
  loc_00E26643: mov [edx], ecx
  loc_00E26645: mov ecx, var_74
  loc_00E26648: push 8001000Dh
  loc_00E2664D: mov [edx+00000004h], ecx
  loc_00E26650: mov ecx, [esi]
  loc_00E26652: push esi
  loc_00E26653: mov [edx+00000008h], eax
  loc_00E26656: mov eax, var_6C
  loc_00E26659: mov [edx+0000000Ch], eax
  loc_00E2665C: call [ecx+0000040Ch]
  loc_00E26662: lea edx, var_28
  loc_00E26665: push eax
  loc_00E26666: push edx
  loc_00E26667: call ebx
  loc_00E26669: push eax
  loc_00E2666A: call [00401238h] ; __vbaLateIdSt
  loc_00E26670: lea ecx, var_28
  loc_00E26673: call edi
  loc_00E26675: mov var_4, 00000000h
  loc_00E2667C: push 00E266B2h
  loc_00E26681: jmp 00E266A8h
  loc_00E26683: lea ecx, var_28
  loc_00E26686: call [00401254h] ; __vbaFreeObj
  loc_00E2668C: lea eax, var_68
  loc_00E2668F: lea ecx, var_58
  loc_00E26692: push eax
  loc_00E26693: lea edx, var_48
  loc_00E26696: push ecx
  loc_00E26697: lea eax, var_38
  loc_00E2669A: push edx
  loc_00E2669B: push eax
  loc_00E2669C: push 00000004h
  loc_00E2669E: call [00401038h] ; __vbaFreeVarList
  loc_00E266A4: add esp, 00000014h
  loc_00E266A7: ret
  loc_00E266A8: lea ecx, var_24
  loc_00E266AB: call [00401024h] ; __vbaFreeVar
  loc_00E266B1: ret
  loc_00E266B2: mov eax, Me
  loc_00E266B5: push eax
  loc_00E266B6: mov ecx, [eax]
  loc_00E266B8: call [ecx+00000008h]
  loc_00E266BB: mov eax, var_4
  loc_00E266BE: mov ecx, var_14
  loc_00E266C1: pop edi
  loc_00E266C2: pop esi
  loc_00E266C3: mov fs:[00000000h], ecx
  loc_00E266CA: pop ebx
  loc_00E266CB: mov esp, ebp
  loc_00E266CD: pop ebp
  loc_00E266CE: retn 0004h
End Sub

Private Sub cari_UnknownEvent_9 'E250E0
  loc_00E250E0: push ebp
  loc_00E250E1: mov ebp, esp
  loc_00E250E3: sub esp, 0000000Ch
  loc_00E250E6: push 00402836h ; __vbaExceptHandler
  loc_00E250EB: mov eax, fs:[00000000h]
  loc_00E250F1: push eax
  loc_00E250F2: mov fs:[00000000h], esp
  loc_00E250F9: sub esp, 00000034h
  loc_00E250FC: push ebx
  loc_00E250FD: push esi
  loc_00E250FE: push edi
  loc_00E250FF: mov var_C, esp
  loc_00E25102: mov var_8, 004024A0h
  loc_00E25109: mov esi, Me
  loc_00E2510C: mov eax, esi
  loc_00E2510E: and eax, 00000001h
  loc_00E25111: mov var_4, eax
  loc_00E25114: and esi, FFFFFFFEh
  loc_00E25117: push esi
  loc_00E25118: mov Me, esi
  loc_00E2511B: mov ecx, [esi]
  loc_00E2511D: call [ecx+00000004h]
  loc_00E25120: mov edx, [esi]
  loc_00E25122: push esi
  loc_00E25123: mov var_18, 00000000h
  loc_00E2512A: call [edx+00000308h]
  loc_00E25130: push eax
  loc_00E25131: lea eax, var_18
  loc_00E25134: push eax
  loc_00E25135: call [004010ACh] ; __vbaObjSet
  loc_00E2513B: mov edi, eax
  loc_00E2513D: push FFFFFFFFh
  loc_00E2513F: push edi
  loc_00E25140: mov ecx, [edi]
  loc_00E25142: call [ecx+0000009Ch]
  loc_00E25148: test eax, eax
  loc_00E2514A: fnclex
  loc_00E2514C: jge 00E25160h
  loc_00E2514E: push 0000009Ch
  loc_00E25153: push 006DCAD0h
  loc_00E25158: push edi
  loc_00E25159: push eax
  loc_00E2515A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25160: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E25166: lea ecx, var_18
  loc_00E25169: call ebx
  loc_00E2516B: mov edx, [esi]
  loc_00E2516D: push esi
  loc_00E2516E: call [edx+0000034Ch]
  loc_00E25174: push eax
  loc_00E25175: lea eax, var_18
  loc_00E25178: push eax
  loc_00E25179: call [004010ACh] ; __vbaObjSet
  loc_00E2517F: mov edi, eax
  loc_00E25181: push 00000000h
  loc_00E25183: push edi
  loc_00E25184: mov ecx, [edi]
  loc_00E25186: call [ecx+0000009Ch]
  loc_00E2518C: test eax, eax
  loc_00E2518E: fnclex
  loc_00E25190: jge 00E251A4h
  loc_00E25192: push 0000009Ch
  loc_00E25197: push 006DCAD0h
  loc_00E2519C: push edi
  loc_00E2519D: push eax
  loc_00E2519E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E251A4: lea ecx, var_18
  loc_00E251A7: call ebx
  loc_00E251A9: sub esp, 00000010h
  loc_00E251AC: mov ecx, 0000000Bh
  loc_00E251B1: mov edx, esp
  loc_00E251B3: xor eax, eax
  loc_00E251B5: push 8001000Dh
  loc_00E251BA: push esi
  loc_00E251BB: mov [edx], ecx
  loc_00E251BD: mov ecx, var_24
  loc_00E251C0: mov [edx+00000004h], ecx
  loc_00E251C3: mov ecx, [esi]
  loc_00E251C5: mov [edx+00000008h], eax
  loc_00E251C8: mov eax, var_1C
  loc_00E251CB: mov [edx+0000000Ch], eax
  loc_00E251CE: call [ecx+00000410h]
  loc_00E251D4: lea edx, var_18
  loc_00E251D7: push eax
  loc_00E251D8: push edx
  loc_00E251D9: call [004010ACh] ; __vbaObjSet
  loc_00E251DF: push eax
  loc_00E251E0: call [00401238h] ; __vbaLateIdSt
  loc_00E251E6: lea ecx, var_18
  loc_00E251E9: call ebx
  loc_00E251EB: mov var_4, 00000000h
  loc_00E251F2: push 00E25204h
  loc_00E251F7: jmp 00E25203h
  loc_00E251F9: lea ecx, var_18
  loc_00E251FC: call [00401254h] ; __vbaFreeObj
  loc_00E25202: ret
  loc_00E25203: ret
  loc_00E25204: mov eax, Me
  loc_00E25207: push eax
  loc_00E25208: mov ecx, [eax]
  loc_00E2520A: call [ecx+00000008h]
  loc_00E2520D: mov eax, var_4
  loc_00E25210: mov ecx, var_14
  loc_00E25213: pop edi
  loc_00E25214: pop esi
  loc_00E25215: mov fs:[00000000h], ecx
  loc_00E2521C: pop ebx
  loc_00E2521D: mov esp, ebp
  loc_00E2521F: pop ebp
  loc_00E25220: retn 0004h
End Sub

Private Sub callx_UnknownEvent_9 'E24F50
  loc_00E24F50: push ebp
  loc_00E24F51: mov ebp, esp
  loc_00E24F53: sub esp, 0000000Ch
  loc_00E24F56: push 00402836h ; __vbaExceptHandler
  loc_00E24F5B: mov eax, fs:[00000000h]
  loc_00E24F61: push eax
  loc_00E24F62: mov fs:[00000000h], esp
  loc_00E24F69: sub esp, 00000034h
  loc_00E24F6C: push ebx
  loc_00E24F6D: push esi
  loc_00E24F6E: push edi
  loc_00E24F6F: mov var_C, esp
  loc_00E24F72: mov var_8, 00402490h
  loc_00E24F79: mov esi, Me
  loc_00E24F7C: mov eax, esi
  loc_00E24F7E: and eax, 00000001h
  loc_00E24F81: mov var_4, eax
  loc_00E24F84: and esi, FFFFFFFEh
  loc_00E24F87: push esi
  loc_00E24F88: mov Me, esi
  loc_00E24F8B: mov ecx, [esi]
  loc_00E24F8D: call [ecx+00000004h]
  loc_00E24F90: mov eax, [00E3D0ECh]
  loc_00E24F95: mov var_18, 00000000h
  loc_00E24F9C: test eax, eax
  loc_00E24F9E: jnz 00E24FB0h
  loc_00E24FA0: push 00E3D0ECh
  loc_00E24FA5: push 006CC808h
  loc_00E24FAA: call [004011A0h] ; __vbaNew2
  loc_00E24FB0: sub esp, 00000010h
  loc_00E24FB3: mov ecx, 0000000Ah
  loc_00E24FB8: mov ebx, esp
  loc_00E24FBA: mov var_28, ecx
  loc_00E24FBD: mov eax, 80020004h
  loc_00E24FC2: sub esp, 00000010h
  loc_00E24FC5: mov [ebx], ecx
  loc_00E24FC7: mov ecx, var_34
  loc_00E24FCA: mov var_20, eax
  loc_00E24FCD: mov edi, [00E3D0ECh]
  loc_00E24FD3: mov [ebx+00000004h], ecx
  loc_00E24FD6: mov ecx, esp
  loc_00E24FD8: mov edx, [edi]
  loc_00E24FDA: push edi
  loc_00E24FDB: mov [ebx+00000008h], eax
  loc_00E24FDE: mov eax, var_2C
  loc_00E24FE1: mov [ebx+0000000Ch], eax
  loc_00E24FE4: mov eax, var_28
  loc_00E24FE7: mov ebx, var_24
  loc_00E24FEA: mov [ecx], eax
  loc_00E24FEC: mov eax, var_20
  loc_00E24FEF: mov [ecx+00000004h], ebx
  loc_00E24FF2: mov [ecx+00000008h], eax
  loc_00E24FF5: mov eax, var_1C
  loc_00E24FF8: mov [ecx+0000000Ch], eax
  loc_00E24FFB: call [edx+000002B0h]
  loc_00E25001: test eax, eax
  loc_00E25003: fnclex
  loc_00E25005: jge 00E25019h
  loc_00E25007: push 000002B0h
  loc_00E2500C: push 006E00E8h
  loc_00E25011: push edi
  loc_00E25012: push eax
  loc_00E25013: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25019: sub esp, 00000010h
  loc_00E2501C: mov ecx, 0000000Bh
  loc_00E25021: mov edx, esp
  loc_00E25023: or eax, FFFFFFFFh
  loc_00E25026: push 80010007h
  loc_00E2502B: push esi
  loc_00E2502C: mov [edx], ecx
  loc_00E2502E: mov ecx, [esi]
  loc_00E25030: mov [edx+00000004h], ebx
  loc_00E25033: mov [edx+00000008h], eax
  loc_00E25036: mov eax, var_1C
  loc_00E25039: mov [edx+0000000Ch], eax
  loc_00E2503C: call [ecx+00000400h]
  loc_00E25042: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E25048: lea edx, var_18
  loc_00E2504B: push eax
  loc_00E2504C: push edx
  loc_00E2504D: call edi
  loc_00E2504F: push eax
  loc_00E25050: call [00401238h] ; __vbaLateIdSt
  loc_00E25056: lea ecx, var_18
  loc_00E25059: call [00401254h] ; __vbaFreeObj
  loc_00E2505F: sub esp, 00000010h
  loc_00E25062: mov ecx, 0000000Bh
  loc_00E25067: mov edx, esp
  loc_00E25069: xor eax, eax
  loc_00E2506B: push 80010007h
  loc_00E25070: push esi
  loc_00E25071: mov [edx], ecx
  loc_00E25073: mov ecx, [esi]
  loc_00E25075: mov [edx+00000004h], ebx
  loc_00E25078: mov [edx+00000008h], eax
  loc_00E2507B: mov eax, var_1C
  loc_00E2507E: mov [edx+0000000Ch], eax
  loc_00E25081: call [ecx+00000408h]
  loc_00E25087: lea edx, var_18
  loc_00E2508A: push eax
  loc_00E2508B: push edx
  loc_00E2508C: call edi
  loc_00E2508E: push eax
  loc_00E2508F: call [00401238h] ; __vbaLateIdSt
  loc_00E25095: lea ecx, var_18
  loc_00E25098: call [00401254h] ; __vbaFreeObj
  loc_00E2509E: mov var_4, 00000000h
  loc_00E250A5: push 00E250B7h
  loc_00E250AA: jmp 00E250B6h
  loc_00E250AC: lea ecx, var_18
  loc_00E250AF: call [00401254h] ; __vbaFreeObj
  loc_00E250B5: ret
  loc_00E250B6: ret
  loc_00E250B7: mov eax, Me
  loc_00E250BA: push eax
  loc_00E250BB: mov ecx, [eax]
  loc_00E250BD: call [ecx+00000008h]
  loc_00E250C0: mov eax, var_4
  loc_00E250C3: mov ecx, var_14
  loc_00E250C6: pop edi
  loc_00E250C7: pop esi
  loc_00E250C8: mov fs:[00000000h], ecx
  loc_00E250CF: pop ebx
  loc_00E250D0: mov esp, ebp
  loc_00E250D2: pop ebp
  loc_00E250D3: retn 0004h
End Sub

Private Sub Image7_Click() 'E25D30
  loc_00E25D30: push ebp
  loc_00E25D31: mov ebp, esp
  loc_00E25D33: sub esp, 0000000Ch
  loc_00E25D36: push 00402836h ; __vbaExceptHandler
  loc_00E25D3B: mov eax, fs:[00000000h]
  loc_00E25D41: push eax
  loc_00E25D42: mov fs:[00000000h], esp
  loc_00E25D49: sub esp, 00000034h
  loc_00E25D4C: push ebx
  loc_00E25D4D: push esi
  loc_00E25D4E: push edi
  loc_00E25D4F: mov var_C, esp
  loc_00E25D52: mov var_8, 004024F0h
  loc_00E25D59: mov esi, Me
  loc_00E25D5C: mov eax, esi
  loc_00E25D5E: and eax, 00000001h
  loc_00E25D61: mov var_4, eax
  loc_00E25D64: and esi, FFFFFFFEh
  loc_00E25D67: push esi
  loc_00E25D68: mov Me, esi
  loc_00E25D6B: mov ecx, [esi]
  loc_00E25D6D: call [ecx+00000004h]
  loc_00E25D70: mov edx, [esi]
  loc_00E25D72: push esi
  loc_00E25D73: mov var_18, 00000000h
  loc_00E25D7A: call [edx+0000034Ch]
  loc_00E25D80: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E25D86: push eax
  loc_00E25D87: lea eax, var_18
  loc_00E25D8A: push eax
  loc_00E25D8B: call ebx
  loc_00E25D8D: mov edi, eax
  loc_00E25D8F: push 00000000h
  loc_00E25D91: push edi
  loc_00E25D92: mov ecx, [edi]
  loc_00E25D94: call [ecx+0000009Ch]
  loc_00E25D9A: test eax, eax
  loc_00E25D9C: fnclex
  loc_00E25D9E: jge 00E25DB2h
  loc_00E25DA0: push 0000009Ch
  loc_00E25DA5: push 006DCAD0h
  loc_00E25DAA: push edi
  loc_00E25DAB: push eax
  loc_00E25DAC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25DB2: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E25DB8: lea ecx, var_18
  loc_00E25DBB: call edi
  loc_00E25DBD: mov edx, [esi]
  loc_00E25DBF: push esi
  loc_00E25DC0: call [edx+00000320h]
  loc_00E25DC6: push eax
  loc_00E25DC7: lea eax, var_18
  loc_00E25DCA: push eax
  loc_00E25DCB: call ebx
  loc_00E25DCD: mov ecx, [eax]
  loc_00E25DCF: push FFFFFFFFh
  loc_00E25DD1: push eax
  loc_00E25DD2: mov var_3C, eax
  loc_00E25DD5: call [ecx+0000009Ch]
  loc_00E25DDB: test eax, eax
  loc_00E25DDD: fnclex
  loc_00E25DDF: jge 00E25DF6h
  loc_00E25DE1: mov edx, var_3C
  loc_00E25DE4: push 0000009Ch
  loc_00E25DE9: push 006DCAD0h
  loc_00E25DEE: push edx
  loc_00E25DEF: push eax
  loc_00E25DF0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25DF6: lea ecx, var_18
  loc_00E25DF9: call edi
  loc_00E25DFB: mov eax, [esi]
  loc_00E25DFD: push esi
  loc_00E25DFE: call [eax+00000424h]
  loc_00E25E04: lea ecx, var_18
  loc_00E25E07: push eax
  loc_00E25E08: push ecx
  loc_00E25E09: call ebx
  loc_00E25E0B: mov edx, [eax]
  loc_00E25E0D: push 006E1614h ; "........"
  loc_00E25E12: push eax
  loc_00E25E13: mov var_3C, eax
  loc_00E25E16: call [edx+00000054h]
  loc_00E25E19: test eax, eax
  loc_00E25E1B: fnclex
  loc_00E25E1D: jge 00E25E31h
  loc_00E25E1F: mov ecx, var_3C
  loc_00E25E22: push 00000054h
  loc_00E25E24: push 006E0410h
  loc_00E25E29: push ecx
  loc_00E25E2A: push eax
  loc_00E25E2B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25E31: lea ecx, var_18
  loc_00E25E34: call edi
  loc_00E25E36: mov edx, [esi]
  loc_00E25E38: push esi
  loc_00E25E39: call [edx+00000378h]
  loc_00E25E3F: push eax
  loc_00E25E40: lea eax, var_18
  loc_00E25E43: push eax
  loc_00E25E44: call ebx
  loc_00E25E46: mov ecx, [eax]
  loc_00E25E48: push 006DCC80h
  loc_00E25E4D: push eax
  loc_00E25E4E: mov var_3C, eax
  loc_00E25E51: call [ecx+000000A4h]
  loc_00E25E57: test eax, eax
  loc_00E25E59: fnclex
  loc_00E25E5B: jge 00E25E72h
  loc_00E25E5D: mov edx, var_3C
  loc_00E25E60: push 000000A4h
  loc_00E25E65: push 006DCB70h
  loc_00E25E6A: push edx
  loc_00E25E6B: push eax
  loc_00E25E6C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25E72: lea ecx, var_18
  loc_00E25E75: call edi
  loc_00E25E77: mov eax, [esi]
  loc_00E25E79: push esi
  loc_00E25E7A: call [eax+00000388h]
  loc_00E25E80: lea ecx, var_18
  loc_00E25E83: push eax
  loc_00E25E84: push ecx
  loc_00E25E85: call ebx
  loc_00E25E87: mov edx, [eax]
  loc_00E25E89: push 006DCC80h
  loc_00E25E8E: push eax
  loc_00E25E8F: mov var_3C, eax
  loc_00E25E92: call [edx+00000054h]
  loc_00E25E95: test eax, eax
  loc_00E25E97: fnclex
  loc_00E25E99: jge 00E25EADh
  loc_00E25E9B: mov ecx, var_3C
  loc_00E25E9E: push 00000054h
  loc_00E25EA0: push 006E0410h
  loc_00E25EA5: push ecx
  loc_00E25EA6: push eax
  loc_00E25EA7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25EAD: lea ecx, var_18
  loc_00E25EB0: call edi
  loc_00E25EB2: mov edx, [esi]
  loc_00E25EB4: push esi
  loc_00E25EB5: call [edx+0000039Ch]
  loc_00E25EBB: push eax
  loc_00E25EBC: lea eax, var_18
  loc_00E25EBF: push eax
  loc_00E25EC0: call ebx
  loc_00E25EC2: mov ecx, [eax]
  loc_00E25EC4: push 006DCC80h
  loc_00E25EC9: push eax
  loc_00E25ECA: mov var_3C, eax
  loc_00E25ECD: call [ecx+000000A4h]
  loc_00E25ED3: test eax, eax
  loc_00E25ED5: fnclex
  loc_00E25ED7: jge 00E25EEEh
  loc_00E25ED9: mov edx, var_3C
  loc_00E25EDC: push 000000A4h
  loc_00E25EE1: push 006DCB70h
  loc_00E25EE6: push edx
  loc_00E25EE7: push eax
  loc_00E25EE8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25EEE: lea ecx, var_18
  loc_00E25EF1: call edi
  loc_00E25EF3: mov eax, [esi]
  loc_00E25EF5: push esi
  loc_00E25EF6: call [eax+000003B0h]
  loc_00E25EFC: lea ecx, var_18
  loc_00E25EFF: push eax
  loc_00E25F00: push ecx
  loc_00E25F01: call ebx
  loc_00E25F03: mov edx, [eax]
  loc_00E25F05: push 006DCC80h
  loc_00E25F0A: push eax
  loc_00E25F0B: mov var_3C, eax
  loc_00E25F0E: call [edx+00000054h]
  loc_00E25F11: test eax, eax
  loc_00E25F13: fnclex
  loc_00E25F15: jge 00E25F29h
  loc_00E25F17: mov ecx, var_3C
  loc_00E25F1A: push 00000054h
  loc_00E25F1C: push 006E0410h
  loc_00E25F21: push ecx
  loc_00E25F22: push eax
  loc_00E25F23: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25F29: lea ecx, var_18
  loc_00E25F2C: call edi
  loc_00E25F2E: sub esp, 00000010h
  loc_00E25F31: mov ecx, 0000000Bh
  loc_00E25F36: mov edx, esp
  loc_00E25F38: or eax, FFFFFFFFh
  loc_00E25F3B: push 80010007h
  loc_00E25F40: push esi
  loc_00E25F41: mov [edx], ecx
  loc_00E25F43: mov ecx, var_24
  loc_00E25F46: mov [edx+00000004h], ecx
  loc_00E25F49: mov ecx, [esi]
  loc_00E25F4B: mov [edx+00000008h], eax
  loc_00E25F4E: mov eax, var_1C
  loc_00E25F51: mov [edx+0000000Ch], eax
  loc_00E25F54: call [ecx+00000458h]
  loc_00E25F5A: lea edx, var_18
  loc_00E25F5D: push eax
  loc_00E25F5E: push edx
  loc_00E25F5F: call ebx
  loc_00E25F61: push eax
  loc_00E25F62: call [00401238h] ; __vbaLateIdSt
  loc_00E25F68: lea ecx, var_18
  loc_00E25F6B: call edi
  loc_00E25F6D: sub esp, 00000010h
  loc_00E25F70: mov ecx, 0000000Bh
  loc_00E25F75: mov edx, esp
  loc_00E25F77: xor eax, eax
  loc_00E25F79: push 80010007h
  loc_00E25F7E: push esi
  loc_00E25F7F: mov [edx], ecx
  loc_00E25F81: mov ecx, var_24
  loc_00E25F84: mov [edx+00000004h], ecx
  loc_00E25F87: mov ecx, [esi]
  loc_00E25F89: mov [edx+00000008h], eax
  loc_00E25F8C: mov eax, var_1C
  loc_00E25F8F: mov [edx+0000000Ch], eax
  loc_00E25F92: call [ecx+00000400h]
  loc_00E25F98: lea edx, var_18
  loc_00E25F9B: push eax
  loc_00E25F9C: push edx
  loc_00E25F9D: call ebx
  loc_00E25F9F: push eax
  loc_00E25FA0: call [00401238h] ; __vbaLateIdSt
  loc_00E25FA6: lea ecx, var_18
  loc_00E25FA9: call edi
  loc_00E25FAB: sub esp, 00000010h
  loc_00E25FAE: mov ecx, 0000000Bh
  loc_00E25FB3: mov edx, esp
  loc_00E25FB5: or eax, FFFFFFFFh
  loc_00E25FB8: push 80010007h
  loc_00E25FBD: push esi
  loc_00E25FBE: mov [edx], ecx
  loc_00E25FC0: mov ecx, var_24
  loc_00E25FC3: mov [edx+00000004h], ecx
  loc_00E25FC6: mov ecx, [esi]
  loc_00E25FC8: mov [edx+00000008h], eax
  loc_00E25FCB: mov eax, var_1C
  loc_00E25FCE: mov [edx+0000000Ch], eax
  loc_00E25FD1: call [ecx+00000408h]
  loc_00E25FD7: lea edx, var_18
  loc_00E25FDA: push eax
  loc_00E25FDB: push edx
  loc_00E25FDC: call ebx
  loc_00E25FDE: push eax
  loc_00E25FDF: call [00401238h] ; __vbaLateIdSt
  loc_00E25FE5: lea ecx, var_18
  loc_00E25FE8: call edi
  loc_00E25FEA: mov var_4, 00000000h
  loc_00E25FF1: push 00E26003h
  loc_00E25FF6: jmp 00E26002h
  loc_00E25FF8: lea ecx, var_18
  loc_00E25FFB: call [00401254h] ; __vbaFreeObj
  loc_00E26001: ret
  loc_00E26002: ret
  loc_00E26003: mov eax, Me
  loc_00E26006: push eax
  loc_00E26007: mov ecx, [eax]
  loc_00E26009: call [ecx+00000008h]
  loc_00E2600C: mov eax, var_4
  loc_00E2600F: mov ecx, var_14
  loc_00E26012: pop edi
  loc_00E26013: pop esi
  loc_00E26014: mov fs:[00000000h], ecx
  loc_00E2601B: pop ebx
  loc_00E2601C: mov esp, ebp
  loc_00E2601E: pop ebp
  loc_00E2601F: retn 0004h
End Sub

Private Sub refreshh_UnknownEvent_9 'E26BE0
  loc_00E26BE0: push ebp
  loc_00E26BE1: mov ebp, esp
  loc_00E26BE3: sub esp, 0000000Ch
  loc_00E26BE6: push 00402836h ; __vbaExceptHandler
  loc_00E26BEB: mov eax, fs:[00000000h]
  loc_00E26BF1: push eax
  loc_00E26BF2: mov fs:[00000000h], esp
  loc_00E26BF9: sub esp, 00000040h
  loc_00E26BFC: push ebx
  loc_00E26BFD: push esi
  loc_00E26BFE: push edi
  loc_00E26BFF: mov var_C, esp
  loc_00E26C02: mov var_8, 00402530h
  loc_00E26C09: mov esi, Me
  loc_00E26C0C: mov eax, esi
  loc_00E26C0E: and eax, 00000001h
  loc_00E26C11: mov var_4, eax
  loc_00E26C14: and esi, FFFFFFFEh
  loc_00E26C17: push esi
  loc_00E26C18: mov Me, esi
  loc_00E26C1B: mov ecx, [esi]
  loc_00E26C1D: call [ecx+00000004h]
  loc_00E26C20: sub esp, 00000010h
  loc_00E26C23: mov ecx, 00000008h
  loc_00E26C28: mov edx, esp
  loc_00E26C2A: xor eax, eax
  loc_00E26C2C: mov var_18, eax
  loc_00E26C2F: mov var_1C, eax
  loc_00E26C32: mov [edx], ecx
  loc_00E26C34: mov ecx, var_38
  loc_00E26C37: mov var_2C, eax
  loc_00E26C3A: mov eax, 006E1738h ; "lc"
  loc_00E26C3F: mov [edx+00000004h], ecx
  loc_00E26C42: mov ecx, [esi]
  loc_00E26C44: push 0000000Eh
  loc_00E26C46: push esi
  loc_00E26C47: mov [edx+00000008h], eax
  loc_00E26C4A: mov eax, var_30
  loc_00E26C4D: mov [edx+0000000Ch], eax
  loc_00E26C50: call [ecx+000004A8h]
  loc_00E26C56: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E26C5C: lea edx, var_18
  loc_00E26C5F: push eax
  loc_00E26C60: push edx
  loc_00E26C61: call edi
  loc_00E26C63: push eax
  loc_00E26C64: call [00401238h] ; __vbaLateIdSt
  loc_00E26C6A: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E26C70: lea ecx, var_18
  loc_00E26C73: call ebx
  loc_00E26C75: mov eax, [esi]
  loc_00E26C77: push 00000000h
  loc_00E26C79: push 00000019h
  loc_00E26C7B: push esi
  loc_00E26C7C: call [eax+000004A8h]
  loc_00E26C82: lea ecx, var_18
  loc_00E26C85: push eax
  loc_00E26C86: push ecx
  loc_00E26C87: call edi
  loc_00E26C89: push eax
  loc_00E26C8A: call [00401030h] ; __vbaLateIdCall
  loc_00E26C90: add esp, 0000000Ch
  loc_00E26C93: lea ecx, var_18
  loc_00E26C96: call ebx
  loc_00E26C98: mov edx, [esi]
  loc_00E26C9A: push 006E05D8h
  loc_00E26C9F: push esi
  loc_00E26CA0: call [edx+000004A8h]
  loc_00E26CA6: push eax
  loc_00E26CA7: lea eax, var_18
  loc_00E26CAA: push eax
  loc_00E26CAB: call edi
  loc_00E26CAD: push eax
  loc_00E26CAE: call [00401224h] ; __vbaCastObj
  loc_00E26CB4: lea ecx, var_2C
  loc_00E26CB7: mov var_24, eax
  loc_00E26CBA: push ecx
  loc_00E26CBB: mov var_2C, 0000000Dh
  loc_00E26CC2: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E26CC8: mov ecx, [eax]
  loc_00E26CCA: sub esp, 00000010h
  loc_00E26CCD: mov edx, esp
  loc_00E26CCF: mov [edx], ecx
  loc_00E26CD1: mov ecx, [eax+00000004h]
  loc_00E26CD4: mov [edx+00000004h], ecx
  loc_00E26CD7: mov ecx, [eax+00000008h]
  loc_00E26CDA: mov eax, [eax+0000000Ch]
  loc_00E26CDD: mov [edx+00000008h], ecx
  loc_00E26CE0: mov [edx+0000000Ch], eax
  loc_00E26CE3: mov ecx, [esi]
  loc_00E26CE5: push 00000000h
  loc_00E26CE7: push 0000002Ah
  loc_00E26CE9: push esi
  loc_00E26CEA: call [ecx+000004ACh]
  loc_00E26CF0: lea edx, var_1C
  loc_00E26CF3: push eax
  loc_00E26CF4: push edx
  loc_00E26CF5: call edi
  loc_00E26CF7: push eax
  loc_00E26CF8: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E26CFE: lea eax, var_1C
  loc_00E26D01: lea ecx, var_18
  loc_00E26D04: push eax
  loc_00E26D05: push ecx
  loc_00E26D06: push 00000002h
  loc_00E26D08: call [00401048h] ; __vbaFreeObjList
  loc_00E26D0E: add esp, 00000028h
  loc_00E26D11: lea ecx, var_2C
  loc_00E26D14: call [00401024h] ; __vbaFreeVar
  loc_00E26D1A: sub esp, 00000010h
  loc_00E26D1D: mov ecx, 0000000Bh
  loc_00E26D22: mov edx, esp
  loc_00E26D24: xor eax, eax
  loc_00E26D26: push 00000006h
  loc_00E26D28: push esi
  loc_00E26D29: mov [edx], ecx
  loc_00E26D2B: mov ecx, var_38
  loc_00E26D2E: mov [edx+00000004h], ecx
  loc_00E26D31: mov ecx, [esi]
  loc_00E26D33: mov [edx+00000008h], eax
  loc_00E26D36: mov eax, var_30
  loc_00E26D39: mov [edx+0000000Ch], eax
  loc_00E26D3C: call [ecx+000004ACh]
  loc_00E26D42: lea edx, var_18
  loc_00E26D45: push eax
  loc_00E26D46: push edx
  loc_00E26D47: call edi
  loc_00E26D49: push eax
  loc_00E26D4A: call [00401238h] ; __vbaLateIdSt
  loc_00E26D50: lea ecx, var_18
  loc_00E26D53: call ebx
  loc_00E26D55: sub esp, 00000010h
  loc_00E26D58: mov ecx, 0000000Bh
  loc_00E26D5D: mov edx, esp
  loc_00E26D5F: xor eax, eax
  loc_00E26D61: push 8001000Eh
  loc_00E26D66: push esi
  loc_00E26D67: mov [edx], ecx
  loc_00E26D69: mov ecx, var_38
  loc_00E26D6C: mov [edx+00000004h], ecx
  loc_00E26D6F: mov ecx, [esi]
  loc_00E26D71: mov [edx+00000008h], eax
  loc_00E26D74: mov eax, var_30
  loc_00E26D77: mov [edx+0000000Ch], eax
  loc_00E26D7A: call [ecx+000004ACh]
  loc_00E26D80: lea edx, var_18
  loc_00E26D83: push eax
  loc_00E26D84: push edx
  loc_00E26D85: call edi
  loc_00E26D87: push eax
  loc_00E26D88: call [00401238h] ; __vbaLateIdSt
  loc_00E26D8E: lea ecx, var_18
  loc_00E26D91: call ebx
  loc_00E26D93: mov eax, [esi]
  loc_00E26D95: push 00000000h
  loc_00E26D97: push FFFFFDDAh
  loc_00E26D9C: push esi
  loc_00E26D9D: call [eax+000004ACh]
  loc_00E26DA3: lea ecx, var_18
  loc_00E26DA6: push eax
  loc_00E26DA7: push ecx
  loc_00E26DA8: call edi
  loc_00E26DAA: push eax
  loc_00E26DAB: call [00401030h] ; __vbaLateIdCall
  loc_00E26DB1: add esp, 0000000Ch
  loc_00E26DB4: lea ecx, var_18
  loc_00E26DB7: call ebx
  loc_00E26DB9: mov var_4, 00000000h
  loc_00E26DC0: push 00E26DE5h
  loc_00E26DC5: jmp 00E26DE4h
  loc_00E26DC7: lea edx, var_1C
  loc_00E26DCA: lea eax, var_18
  loc_00E26DCD: push edx
  loc_00E26DCE: push eax
  loc_00E26DCF: push 00000002h
  loc_00E26DD1: call [00401048h] ; __vbaFreeObjList
  loc_00E26DD7: add esp, 0000000Ch
  loc_00E26DDA: lea ecx, var_2C
  loc_00E26DDD: call [00401024h] ; __vbaFreeVar
  loc_00E26DE3: ret
  loc_00E26DE4: ret
  loc_00E26DE5: mov eax, Me
  loc_00E26DE8: push eax
  loc_00E26DE9: mov ecx, [eax]
  loc_00E26DEB: call [ecx+00000008h]
  loc_00E26DEE: mov eax, var_4
  loc_00E26DF1: mov ecx, var_14
  loc_00E26DF4: pop edi
  loc_00E26DF5: pop esi
  loc_00E26DF6: mov fs:[00000000h], ecx
  loc_00E26DFD: pop ebx
  loc_00E26DFE: mov esp, ebp
  loc_00E26E00: pop ebp
  loc_00E26E01: retn 0004h
End Sub

Private Sub hapus_UnknownEvent_9 'E25BF0
  loc_00E25BF0: push ebp
  loc_00E25BF1: mov ebp, esp
  loc_00E25BF3: sub esp, 0000000Ch
  loc_00E25BF6: push 00402836h ; __vbaExceptHandler
  loc_00E25BFB: mov eax, fs:[00000000h]
  loc_00E25C01: push eax
  loc_00E25C02: mov fs:[00000000h], esp
  loc_00E25C09: sub esp, 00000028h
  loc_00E25C0C: push ebx
  loc_00E25C0D: push esi
  loc_00E25C0E: push edi
  loc_00E25C0F: mov var_C, esp
  loc_00E25C12: mov var_8, 004024E0h
  loc_00E25C19: mov esi, Me
  loc_00E25C1C: mov eax, esi
  loc_00E25C1E: and eax, 00000001h
  loc_00E25C21: mov var_4, eax
  loc_00E25C24: and esi, FFFFFFFEh
  loc_00E25C27: push esi
  loc_00E25C28: mov Me, esi
  loc_00E25C2B: mov ecx, [esi]
  loc_00E25C2D: call [ecx+00000004h]
  loc_00E25C30: mov edx, [esi]
  loc_00E25C32: xor eax, eax
  loc_00E25C34: push 006DCBD8h
  loc_00E25C39: push eax
  loc_00E25C3A: push 00000018h
  loc_00E25C3C: push esi
  loc_00E25C3D: mov var_18, eax
  loc_00E25C40: mov var_1C, eax
  loc_00E25C43: mov var_2C, eax
  loc_00E25C46: call [edx+000004A8h]
  loc_00E25C4C: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E25C52: push eax
  loc_00E25C53: lea eax, var_18
  loc_00E25C56: push eax
  loc_00E25C57: call ebx
  loc_00E25C59: lea ecx, var_2C
  loc_00E25C5C: push eax
  loc_00E25C5D: push ecx
  loc_00E25C5E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E25C64: add esp, 00000010h
  loc_00E25C67: push eax
  loc_00E25C68: call [00401120h] ; __vbaCastObjVar
  loc_00E25C6E: lea edx, var_1C
  loc_00E25C71: push eax
  loc_00E25C72: push edx
  loc_00E25C73: call ebx
  loc_00E25C75: mov edi, eax
  loc_00E25C77: push 00000001h
  loc_00E25C79: push edi
  loc_00E25C7A: mov eax, [edi]
  loc_00E25C7C: call [eax+00000084h]
  loc_00E25C82: test eax, eax
  loc_00E25C84: fnclex
  loc_00E25C86: jge 00E25C9Ah
  loc_00E25C88: push 00000084h
  loc_00E25C8D: push 006DCBE8h
  loc_00E25C92: push edi
  loc_00E25C93: push eax
  loc_00E25C94: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E25C9A: lea ecx, var_1C
  loc_00E25C9D: lea edx, var_18
  loc_00E25CA0: push ecx
  loc_00E25CA1: push edx
  loc_00E25CA2: push 00000002h
  loc_00E25CA4: call [00401048h] ; __vbaFreeObjList
  loc_00E25CAA: add esp, 0000000Ch
  loc_00E25CAD: lea ecx, var_2C
  loc_00E25CB0: call [00401024h] ; __vbaFreeVar
  loc_00E25CB6: mov eax, [esi]
  loc_00E25CB8: push 00000000h
  loc_00E25CBA: push FFFFFDDAh
  loc_00E25CBF: push esi
  loc_00E25CC0: call [eax+000004ACh]
  loc_00E25CC6: lea ecx, var_18
  loc_00E25CC9: push eax
  loc_00E25CCA: push ecx
  loc_00E25CCB: call ebx
  loc_00E25CCD: push eax
  loc_00E25CCE: call [00401030h] ; __vbaLateIdCall
  loc_00E25CD4: add esp, 0000000Ch
  loc_00E25CD7: lea ecx, var_18
  loc_00E25CDA: call [00401254h] ; __vbaFreeObj
  loc_00E25CE0: mov var_4, 00000000h
  loc_00E25CE7: push 00E25D0Ch
  loc_00E25CEC: jmp 00E25D0Bh
  loc_00E25CEE: lea edx, var_1C
  loc_00E25CF1: lea eax, var_18
  loc_00E25CF4: push edx
  loc_00E25CF5: push eax
  loc_00E25CF6: push 00000002h
  loc_00E25CF8: call [00401048h] ; __vbaFreeObjList
  loc_00E25CFE: add esp, 0000000Ch
  loc_00E25D01: lea ecx, var_2C
  loc_00E25D04: call [00401024h] ; __vbaFreeVar
  loc_00E25D0A: ret
  loc_00E25D0B: ret
  loc_00E25D0C: mov eax, Me
  loc_00E25D0F: push eax
  loc_00E25D10: mov ecx, [eax]
  loc_00E25D12: call [ecx+00000008h]
  loc_00E25D15: mov eax, var_4
  loc_00E25D18: mov ecx, var_14
  loc_00E25D1B: pop edi
  loc_00E25D1C: pop esi
  loc_00E25D1D: mov fs:[00000000h], ecx
  loc_00E25D24: pop ebx
  loc_00E25D25: mov esp, ebp
  loc_00E25D27: pop ebp
  loc_00E25D28: retn 0004h
End Sub

Private Sub Timer1_Timer() 'E2C9C0
  loc_00E2C9C0: push ebp
  loc_00E2C9C1: mov ebp, esp
  loc_00E2C9C3: sub esp, 0000000Ch
  loc_00E2C9C6: push 00402836h ; __vbaExceptHandler
  loc_00E2C9CB: mov eax, fs:[00000000h]
  loc_00E2C9D1: push eax
  loc_00E2C9D2: mov fs:[00000000h], esp
  loc_00E2C9D9: sub esp, 00000028h
  loc_00E2C9DC: push ebx
  loc_00E2C9DD: push esi
  loc_00E2C9DE: push edi
  loc_00E2C9DF: mov var_C, esp
  loc_00E2C9E2: mov var_8, 00402560h
  loc_00E2C9E9: mov esi, Me
  loc_00E2C9EC: mov eax, esi
  loc_00E2C9EE: and eax, 00000001h
  loc_00E2C9F1: mov var_4, eax
  loc_00E2C9F4: and esi, FFFFFFFEh
  loc_00E2C9F7: push esi
  loc_00E2C9F8: mov Me, esi
  loc_00E2C9FB: mov ecx, [esi]
  loc_00E2C9FD: call [ecx+00000004h]
  loc_00E2CA00: mov edx, [esi]
  loc_00E2CA02: xor eax, eax
  loc_00E2CA04: push esi
  loc_00E2CA05: mov var_18, eax
  loc_00E2CA08: mov var_1C, eax
  loc_00E2CA0B: mov var_2C, eax
  loc_00E2CA0E: call [edx+00000490h]
  loc_00E2CA14: push eax
  loc_00E2CA15: lea eax, var_1C
  loc_00E2CA18: push eax
  loc_00E2CA19: call [004010ACh] ; __vbaObjSet
  loc_00E2CA1F: lea ecx, var_2C
  loc_00E2CA22: mov edi, eax
  loc_00E2CA24: push ecx
  loc_00E2CA25: call [004011E8h] ; rtcGetTimeVar
  loc_00E2CA2B: mov ebx, [edi]
  loc_00E2CA2D: lea edx, var_2C
  loc_00E2CA30: lea eax, var_18
  loc_00E2CA33: push edx
  loc_00E2CA34: push eax
  loc_00E2CA35: call [00401180h] ; __vbaStrVarVal
  loc_00E2CA3B: push eax
  loc_00E2CA3C: push edi
  loc_00E2CA3D: call [ebx+00000054h]
  loc_00E2CA40: test eax, eax
  loc_00E2CA42: fnclex
  loc_00E2CA44: jge 00E2CA55h
  loc_00E2CA46: push 00000054h
  loc_00E2CA48: push 006E0410h
  loc_00E2CA4D: push edi
  loc_00E2CA4E: push eax
  loc_00E2CA4F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CA55: mov edi, [00401258h] ; __vbaFreeStr
  loc_00E2CA5B: lea ecx, var_18
  loc_00E2CA5E: call edi
  loc_00E2CA60: lea ecx, var_1C
  loc_00E2CA63: call [00401254h] ; __vbaFreeObj
  loc_00E2CA69: lea ecx, var_2C
  loc_00E2CA6C: call [00401024h] ; __vbaFreeVar
  loc_00E2CA72: mov ecx, [esi]
  loc_00E2CA74: push esi
  loc_00E2CA75: call [ecx+00000488h]
  loc_00E2CA7B: lea edx, var_1C
  loc_00E2CA7E: push eax
  loc_00E2CA7F: push edx
  loc_00E2CA80: call [004010ACh] ; __vbaObjSet
  loc_00E2CA86: mov esi, eax
  loc_00E2CA88: lea eax, var_2C
  loc_00E2CA8B: push eax
  loc_00E2CA8C: call [004011D8h] ; rtcGetDateVar
  loc_00E2CA92: mov ebx, [esi]
  loc_00E2CA94: lea ecx, var_2C
  loc_00E2CA97: lea edx, var_18
  loc_00E2CA9A: push ecx
  loc_00E2CA9B: push edx
  loc_00E2CA9C: call [00401180h] ; __vbaStrVarVal
  loc_00E2CAA2: push eax
  loc_00E2CAA3: push esi
  loc_00E2CAA4: call [ebx+00000054h]
  loc_00E2CAA7: test eax, eax
  loc_00E2CAA9: fnclex
  loc_00E2CAAB: jge 00E2CABCh
  loc_00E2CAAD: push 00000054h
  loc_00E2CAAF: push 006E0410h
  loc_00E2CAB4: push esi
  loc_00E2CAB5: push eax
  loc_00E2CAB6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E2CABC: lea ecx, var_18
  loc_00E2CABF: call edi
  loc_00E2CAC1: lea ecx, var_1C
  loc_00E2CAC4: call [00401254h] ; __vbaFreeObj
  loc_00E2CACA: lea ecx, var_2C
  loc_00E2CACD: call [00401024h] ; __vbaFreeVar
  loc_00E2CAD3: mov var_4, 00000000h
  loc_00E2CADA: push 00E2CAFEh
  loc_00E2CADF: jmp 00E2CAFDh
  loc_00E2CAE1: lea ecx, var_18
  loc_00E2CAE4: call [00401258h] ; __vbaFreeStr
  loc_00E2CAEA: lea ecx, var_1C
  loc_00E2CAED: call [00401254h] ; __vbaFreeObj
  loc_00E2CAF3: lea ecx, var_2C
  loc_00E2CAF6: call [00401024h] ; __vbaFreeVar
  loc_00E2CAFC: ret
  loc_00E2CAFD: ret
  loc_00E2CAFE: mov eax, Me
  loc_00E2CB01: push eax
  loc_00E2CB02: mov ecx, [eax]
  loc_00E2CB04: call [ecx+00000008h]
  loc_00E2CB07: mov eax, var_4
  loc_00E2CB0A: mov ecx, var_14
  loc_00E2CB0D: pop edi
  loc_00E2CB0E: pop esi
  loc_00E2CB0F: mov fs:[00000000h], ecx
  loc_00E2CB16: pop ebx
  loc_00E2CB17: mov esp, ebp
  loc_00E2CB19: pop ebp
  loc_00E2CB1A: retn 0004h
End Sub

Public Function Terbilang(Angka) 'E23C20
  loc_00E23C20: push ebp
  loc_00E23C21: mov ebp, esp
  loc_00E23C23: sub esp, 0000000Ch
  loc_00E23C26: push 00402836h ; __vbaExceptHandler
  loc_00E23C2B: mov eax, fs:[00000000h]
  loc_00E23C31: push eax
  loc_00E23C32: mov fs:[00000000h], esp
  loc_00E23C39: sub esp, 00000190h
  loc_00E23C3F: push ebx
  loc_00E23C40: push esi
  loc_00E23C41: push edi
  loc_00E23C42: mov var_C, esp
  loc_00E23C45: mov var_8, 00402470h
  loc_00E23C4C: xor esi, esi
  loc_00E23C4E: mov var_4, esi
  loc_00E23C51: mov eax, Me
  loc_00E23C54: push eax
  loc_00E23C55: mov ecx, [eax]
  loc_00E23C57: call [ecx+00000004h]
  loc_00E23C5A: mov edx, arg_10
  loc_00E23C5D: mov eax, Angka
  loc_00E23C60: mov var_28, esi
  loc_00E23C63: mov var_2C, esi
  loc_00E23C66: mov [edx], esi
  loc_00E23C68: mov ecx, [eax]
  loc_00E23C6A: push ecx
  loc_00E23C6B: mov var_34, esi
  loc_00E23C6E: mov var_44, esi
  loc_00E23C71: mov var_48, esi
  loc_00E23C74: mov var_4C, esi
  loc_00E23C77: mov var_50, esi
  loc_00E23C7A: mov var_60, esi
  loc_00E23C7D: mov var_70, esi
  loc_00E23C80: mov var_74, esi
  loc_00E23C83: mov var_78, esi
  loc_00E23C86: mov var_88, esi
  loc_00E23C8C: mov var_8C, esi
  loc_00E23C92: mov var_90, esi
  loc_00E23C98: mov var_A0, esi
  loc_00E23C9E: mov var_B0, esi
  loc_00E23CA4: mov var_C0, esi
  loc_00E23CAA: mov var_D0, esi
  loc_00E23CB0: mov var_E0, esi
  loc_00E23CB6: mov var_F0, esi
  loc_00E23CBC: mov var_100, esi
  loc_00E23CC2: mov var_110, esi
  loc_00E23CC8: mov var_120, esi
  loc_00E23CCE: mov var_130, esi
  loc_00E23CD4: mov var_140, esi
  loc_00E23CDA: mov var_150, esi
  loc_00E23CE0: mov var_160, esi
  loc_00E23CE6: mov var_170, esi
  loc_00E23CEC: mov var_184, esi
  loc_00E23CF2: call [0040102Ch] ; __vbaLenBstr
  loc_00E23CF8: mov ecx, eax
  loc_00E23CFA: call [00401110h] ; __vbaI2I4
  loc_00E23D00: mov esi, [004011D4h] ; __vbaVarCmpEq
  loc_00E23D06: mov var_18C, eax
  loc_00E23D0C: mov var_18, 00000001h
  loc_00E23D13: mov dx, var_18C
  loc_00E23D1A: cmp var_18, dx
  loc_00E23D1E: jg 00E23E61h
  loc_00E23D24: movsx edi, var_18
  loc_00E23D28: mov eax, Angka
  loc_00E23D2B: lea ecx, var_A0
  loc_00E23D31: mov var_128, eax
  loc_00E23D37: push ecx
  loc_00E23D38: lea edx, var_130
  loc_00E23D3E: push edi
  loc_00E23D3F: lea eax, var_B0
  loc_00E23D45: push edx
  loc_00E23D46: push eax
  loc_00E23D47: mov var_98, 00000001h
  loc_00E23D51: mov var_A0, 00000002h
  loc_00E23D5B: mov var_130, 00004008h
  loc_00E23D65: call [004010F0h] ; rtcMidCharVar
  loc_00E23D6B: lea ecx, var_B0
  loc_00E23D71: lea edx, var_150
  loc_00E23D77: push ecx
  loc_00E23D78: lea eax, var_C0
  loc_00E23D7E: push edx
  loc_00E23D7F: push eax
  loc_00E23D80: mov var_148, 006E0EB0h ; "."
  loc_00E23D8A: mov var_150, 00008008h
  loc_00E23D94: call __vbaVarCmpEq
  loc_00E23D96: lea ecx, var_D0
  loc_00E23D9C: push eax
  loc_00E23D9D: push ecx
  loc_00E23D9E: call [004011B8h] ; __vbaVarNot
  loc_00E23DA4: push eax
  loc_00E23DA5: call [004010DCh] ; __vbaBoolVarNull
  loc_00E23DAB: mov ebx, eax
  loc_00E23DAD: lea edx, var_B0
  loc_00E23DB3: lea eax, var_A0
  loc_00E23DB9: push edx
  loc_00E23DBA: push eax
  loc_00E23DBB: push 00000002h
  loc_00E23DBD: call [00401038h] ; __vbaFreeVarList
  loc_00E23DC3: add esp, 0000000Ch
  loc_00E23DC6: test bx, bx
  loc_00E23DC9: jz 00E23E4Ah
  loc_00E23DCB: mov ecx, Angka
  loc_00E23DCE: lea edx, var_A0
  loc_00E23DD4: mov var_128, ecx
  loc_00E23DDA: push edx
  loc_00E23DDB: lea eax, var_130
  loc_00E23DE1: push edi
  loc_00E23DE2: lea ecx, var_B0
  loc_00E23DE8: push eax
  loc_00E23DE9: push ecx
  loc_00E23DEA: mov var_98, 00000001h
  loc_00E23DF4: mov var_A0, 00000002h
  loc_00E23DFE: mov var_130, 00004008h
  loc_00E23E08: call [004010F0h] ; rtcMidCharVar
  loc_00E23E0E: lea edx, var_60
  loc_00E23E11: lea eax, var_B0
  loc_00E23E17: push edx
  loc_00E23E18: lea ecx, var_C0
  loc_00E23E1E: push eax
  loc_00E23E1F: push ecx
  loc_00E23E20: call [004011E0h] ; __vbaVarAdd
  loc_00E23E26: mov edx, eax
  loc_00E23E28: lea ecx, var_60
  loc_00E23E2B: call [0040101Ch] ; __vbaVarMove
  loc_00E23E31: lea edx, var_B0
  loc_00E23E37: lea eax, var_A0
  loc_00E23E3D: push edx
  loc_00E23E3E: push eax
  loc_00E23E3F: push 00000002h
  loc_00E23E41: call [00401038h] ; __vbaFreeVarList
  loc_00E23E47: add esp, 0000000Ch
  loc_00E23E4A: mov eax, 00000001h
  loc_00E23E4F: add ax, var_18
  loc_00E23E53: jo 00E24D51h
  loc_00E23E59: mov var_18, eax
  loc_00E23E5C: jmp 00E23D13h
  loc_00E23E61: mov edi, [004010D0h] ; rtcLeftTrimVar
  loc_00E23E67: lea ecx, var_60
  loc_00E23E6A: lea edx, var_A0
  loc_00E23E70: push ecx
  loc_00E23E71: push edx
  loc_00E23E72: call edi
  loc_00E23E74: lea eax, var_A0
  loc_00E23E7A: lea ecx, var_B0
  loc_00E23E80: push eax
  loc_00E23E81: push ecx
  loc_00E23E82: mov var_128, 00000000h
  loc_00E23E8C: mov var_130, 00008002h
  loc_00E23E96: call [00401080h] ; __vbaLenVar
  loc_00E23E9C: lea edx, var_130
  loc_00E23EA2: push eax
  loc_00E23EA3: push edx
  loc_00E23EA4: call [00401108h] ; __vbaVarTstEq
  loc_00E23EAA: lea ecx, var_A0
  loc_00E23EB0: mov ebx, eax
  loc_00E23EB2: call [00401024h] ; __vbaFreeVar
  loc_00E23EB8: test bx, bx
  loc_00E23EBB: jz 00E23EE5h
  loc_00E23EBD: lea edx, var_130
  loc_00E23EC3: lea ecx, var_44
  loc_00E23EC6: mov var_128, 006E0EB8h ; "Nol Rupiah"
  loc_00E23ED0: mov var_130, 00000008h
  loc_00E23EDA: call [00401204h] ; __vbaVarCopy
  loc_00E23EE0: jmp 00E24C65h
  loc_00E23EE5: lea edx, var_60
  loc_00E23EE8: lea ecx, var_A0
  loc_00E23EEE: call [00401204h] ; __vbaVarCopy
  loc_00E23EF4: lea eax, var_A0
  loc_00E23EFA: lea ecx, var_B0
  loc_00E23F00: push eax
  loc_00E23F01: push ecx
  loc_00E23F02: call [004010E4h] ; rtcRightTrimVar
  loc_00E23F08: lea edx, var_B0
  loc_00E23F0E: lea eax, var_C0
  loc_00E23F14: push edx
  loc_00E23F15: push eax
  loc_00E23F16: call edi
  loc_00E23F18: lea ecx, var_C0
  loc_00E23F1E: push ecx
  loc_00E23F1F: call [00401034h] ; __vbaStrVarMove
  loc_00E23F25: mov edx, eax
  loc_00E23F27: lea ecx, var_34
  loc_00E23F2A: call [00401228h] ; __vbaStrMove
  loc_00E23F30: mov ebx, [00401038h] ; __vbaFreeVarList
  loc_00E23F36: lea edx, var_C0
  loc_00E23F3C: lea eax, var_B0
  loc_00E23F42: push edx
  loc_00E23F43: lea ecx, var_A0
  loc_00E23F49: push eax
  loc_00E23F4A: push ecx
  loc_00E23F4B: push 00000003h
  loc_00E23F4D: call ebx
  loc_00E23F4F: mov edx, Angka
  loc_00E23F52: add esp, 00000010h
  loc_00E23F55: lea eax, var_A0
  loc_00E23F5B: mov var_128, edx
  loc_00E23F61: push eax
  loc_00E23F62: lea ecx, var_130
  loc_00E23F68: push 0000000Fh
  loc_00E23F6A: lea edx, var_B0
  loc_00E23F70: mov edi, 00000002h
  loc_00E23F75: push ecx
  loc_00E23F76: push edx
  loc_00E23F77: mov var_98, edi
  loc_00E23F7D: mov var_A0, edi
  loc_00E23F83: mov var_130, 00004008h
  loc_00E23F8D: call [004010F0h] ; rtcMidCharVar
  loc_00E23F93: lea eax, var_B0
  loc_00E23F99: push edi
  loc_00E23F9A: lea ecx, var_C0
  loc_00E23FA0: push eax
  loc_00E23FA1: push ecx
  loc_00E23FA2: call [0040122Ch] ; rtcRightCharVar
  loc_00E23FA8: lea edx, var_C0
  loc_00E23FAE: lea eax, var_90
  loc_00E23FB4: push edx
  loc_00E23FB5: push eax
  loc_00E23FB6: call [00401180h] ; __vbaStrVarVal
  loc_00E23FBC: push eax
  loc_00E23FBD: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E23FC3: call [00401200h] ; __vbaFpI2
  loc_00E23FC9: lea ecx, var_90
  loc_00E23FCF: mov edi, eax
  loc_00E23FD1: call [00401258h] ; __vbaFreeStr
  loc_00E23FD7: lea ecx, var_C0
  loc_00E23FDD: lea edx, var_B0
  loc_00E23FE3: push ecx
  loc_00E23FE4: lea eax, var_A0
  loc_00E23FEA: push edx
  loc_00E23FEB: push eax
  loc_00E23FEC: push 00000003h
  loc_00E23FEE: call ebx
  loc_00E23FF0: add esp, 00000010h
  loc_00E23FF3: test di, di
  loc_00E23FF6: jnz 00E24006h
  loc_00E23FF8: mov edx, 006DCC80h
  loc_00E23FFD: lea ecx, var_48
  loc_00E24000: call [004011B0h] ; __vbaStrCopy
  loc_00E24006: mov edi, [0040101Ch] ; __vbaVarMove
  loc_00E2400C: mov ebx, 00000002h
  loc_00E24011: lea edx, var_130
  loc_00E24017: lea ecx, var_88
  loc_00E2401D: mov var_128, 00000000h
  loc_00E24027: mov var_130, ebx
  loc_00E2402D: call edi
  loc_00E2402F: lea edx, var_130
  loc_00E24035: lea ecx, var_70
  loc_00E24038: mov var_128, 00000000h
  loc_00E24042: mov var_130, ebx
  loc_00E24048: call edi
  loc_00E2404A: mov edx, 006DCC80h
  loc_00E2404F: lea ecx, var_2C
  loc_00E24052: call [004011B0h] ; __vbaStrCopy
  loc_00E24058: mov edi, [00401118h] ; __vbaVarOr
  loc_00E2405E: mov ebx, 00008002h
  loc_00E24063: mov ecx, var_34
  loc_00E24066: push ecx
  loc_00E24067: call [0040102Ch] ; __vbaLenBstr
  loc_00E2406D: mov var_128, eax
  loc_00E24073: lea edx, var_88
  loc_00E24079: lea eax, var_130
  loc_00E2407F: push edx
  loc_00E24080: push eax
  loc_00E24081: mov var_130, 00008003h
  loc_00E2408B: call [004010D4h] ; __vbaVarTstLt
  loc_00E24091: test ax, ax
  loc_00E24094: jz 00E24AE1h
  loc_00E2409A: lea ecx, var_88
  loc_00E240A0: lea edx, var_130
  loc_00E240A6: push ecx
  loc_00E240A7: lea eax, var_A0
  loc_00E240AD: push edx
  loc_00E240AE: push eax
  loc_00E240AF: mov var_128, 00000001h
  loc_00E240B9: mov var_130, 00000002h
  loc_00E240C3: call [004011E0h] ; __vbaVarAdd
  loc_00E240C9: mov edx, eax
  loc_00E240CB: lea ecx, var_88
  loc_00E240D1: call [0040101Ch] ; __vbaVarMove
  loc_00E240D7: lea edx, var_A0
  loc_00E240DD: lea eax, var_88
  loc_00E240E3: lea ecx, var_34
  loc_00E240E6: push edx
  loc_00E240E7: push eax
  loc_00E240E8: mov var_98, 00000001h
  loc_00E240F2: mov var_A0, 00000002h
  loc_00E240FC: mov var_128, ecx
  loc_00E24102: mov var_130, 00004008h
  loc_00E2410C: call [004011D0h] ; __vbaI4Var
  loc_00E24112: lea ecx, var_130
  loc_00E24118: push eax
  loc_00E24119: lea edx, var_B0
  loc_00E2411F: push ecx
  loc_00E24120: push edx
  loc_00E24121: call [004010F0h] ; rtcMidCharVar
  loc_00E24127: lea eax, var_B0
  loc_00E2412D: push eax
  loc_00E2412E: call [00401034h] ; __vbaStrVarMove
  loc_00E24134: mov edx, eax
  loc_00E24136: lea ecx, var_8C
  loc_00E2413C: call [00401228h] ; __vbaStrMove
  loc_00E24142: lea ecx, var_B0
  loc_00E24148: lea edx, var_A0
  loc_00E2414E: push ecx
  loc_00E2414F: push edx
  loc_00E24150: push 00000002h
  loc_00E24152: call [00401038h] ; __vbaFreeVarList
  loc_00E24158: mov eax, var_8C
  loc_00E2415E: add esp, 0000000Ch
  loc_00E24161: push eax
  loc_00E24162: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E24168: fstp real8 ptr var_128
  loc_00E2416E: lea ecx, var_70
  loc_00E24171: lea edx, var_130
  loc_00E24177: push ecx
  loc_00E24178: lea eax, var_A0
  loc_00E2417E: push edx
  loc_00E2417F: push eax
  loc_00E24180: mov var_130, 00000005h
  loc_00E2418A: call [004011E0h] ; __vbaVarAdd
  loc_00E24190: mov edx, eax
  loc_00E24192: lea ecx, var_70
  loc_00E24195: call [0040101Ch] ; __vbaVarMove
  loc_00E2419B: mov ecx, var_34
  loc_00E2419E: push ecx
  loc_00E2419F: call [0040102Ch] ; __vbaLenBstr
  loc_00E241A5: mov var_128, eax
  loc_00E241AB: lea edx, var_130
  loc_00E241B1: lea eax, var_88
  loc_00E241B7: push edx
  loc_00E241B8: lea ecx, var_A0
  loc_00E241BE: push eax
  loc_00E241BF: push ecx
  loc_00E241C0: mov var_130, 00000003h
  loc_00E241CA: mov var_138, 00000001h
  loc_00E241D4: mov var_140, 00000002h
  loc_00E241DE: call [00401004h] ; __vbaVarSub
  loc_00E241E4: push eax
  loc_00E241E5: lea edx, var_140
  loc_00E241EB: lea eax, var_B0
  loc_00E241F1: push edx
  loc_00E241F2: push eax
  loc_00E241F3: call [004011E0h] ; __vbaVarAdd
  loc_00E241F9: mov edx, eax
  loc_00E241FB: lea ecx, var_28
  loc_00E241FE: call [0040101Ch] ; __vbaVarMove
  loc_00E24204: mov ecx, var_8C
  loc_00E2420A: push ecx
  loc_00E2420B: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E24211: fstp real8 ptr var_194
  loc_00E24217: mov eax, var_194
  loc_00E2421D: mov ecx, var_190
  loc_00E24223: test eax, eax
  loc_00E24225: jnz 00E2469Ch
  loc_00E2422B: cmp ecx, 3FF00000h
  loc_00E24231: jnz 00E2469Ch
  loc_00E24237: lea edx, var_28
  loc_00E2423A: lea eax, var_130
  loc_00E24240: push edx
  loc_00E24241: lea ecx, var_A0
  loc_00E24247: push eax
  loc_00E24248: push ecx
  loc_00E24249: mov var_128, 00000001h
  loc_00E24253: mov var_130, ebx
  loc_00E24259: mov var_138, 00000007h
  loc_00E24263: mov var_140, ebx
  loc_00E24269: mov var_148, 0000000Ah
  loc_00E24273: mov var_150, ebx
  loc_00E24279: mov var_158, 0000000Dh
  loc_00E24283: mov var_160, ebx
  loc_00E24289: call __vbaVarCmpEq
  loc_00E2428B: push eax
  loc_00E2428C: lea edx, var_28
  loc_00E2428F: lea eax, var_140
  loc_00E24295: push edx
  loc_00E24296: lea ecx, var_B0
  loc_00E2429C: push eax
  loc_00E2429D: push ecx
  loc_00E2429E: call __vbaVarCmpEq
  loc_00E242A0: lea edx, var_C0
  loc_00E242A6: push eax
  loc_00E242A7: push edx
  loc_00E242A8: call edi
  loc_00E242AA: push eax
  loc_00E242AB: lea eax, var_28
  loc_00E242AE: lea ecx, var_150
  loc_00E242B4: push eax
  loc_00E242B5: lea edx, var_D0
  loc_00E242BB: push ecx
  loc_00E242BC: push edx
  loc_00E242BD: call __vbaVarCmpEq
  loc_00E242BF: push eax
  loc_00E242C0: lea eax, var_E0
  loc_00E242C6: push eax
  loc_00E242C7: call edi
  loc_00E242C9: lea ecx, var_28
  loc_00E242CC: push eax
  loc_00E242CD: lea edx, var_160
  loc_00E242D3: push ecx
  loc_00E242D4: lea eax, var_F0
  loc_00E242DA: push edx
  loc_00E242DB: push eax
  loc_00E242DC: call __vbaVarCmpEq
  loc_00E242DE: lea ecx, var_100
  loc_00E242E4: push eax
  loc_00E242E5: push ecx
  loc_00E242E6: call edi
  loc_00E242E8: push eax
  loc_00E242E9: call [004010DCh] ; __vbaBoolVarNull
  loc_00E242EF: test ax, ax
  loc_00E242F2: jz 00E242FEh
  loc_00E242F4: mov edx, 006E0ED4h ; "Satu "
  loc_00E242F9: jmp 00E2473Eh
  loc_00E242FE: lea edx, var_28
  loc_00E24301: lea eax, var_130
  loc_00E24307: push edx
  loc_00E24308: push eax
  loc_00E24309: mov var_128, 00000004h
  loc_00E24313: mov var_130, ebx
  loc_00E24319: call [00401108h] ; __vbaVarTstEq
  loc_00E2431F: test ax, ax
  loc_00E24322: jz 00E24357h
  loc_00E24324: lea ecx, var_88
  loc_00E2432A: lea edx, var_130
  loc_00E24330: push ecx
  loc_00E24331: push edx
  loc_00E24332: mov var_128, 00000001h
  loc_00E2433C: mov var_130, ebx
  loc_00E24342: call [00401108h] ; __vbaVarTstEq
  loc_00E24348: test ax, ax
  loc_00E2434B: jz 00E242F4h
  loc_00E2434D: mov edx, 006E0EE4h ; "Se"
  loc_00E24352: jmp 00E2473Eh
  loc_00E24357: lea eax, var_28
  loc_00E2435A: lea ecx, var_130
  loc_00E24360: push eax
  loc_00E24361: lea edx, var_A0
  loc_00E24367: push ecx
  loc_00E24368: push edx
  loc_00E24369: mov var_128, 00000002h
  loc_00E24373: mov var_130, ebx
  loc_00E24379: mov var_138, 00000005h
  loc_00E24383: mov var_140, ebx
  loc_00E24389: mov var_148, 00000008h
  loc_00E24393: mov var_150, ebx
  loc_00E24399: mov var_158, 0000000Bh
  loc_00E243A3: mov var_160, ebx
  loc_00E243A9: mov var_168, 0000000Eh
  loc_00E243B3: mov var_170, ebx
  loc_00E243B9: call __vbaVarCmpEq
  loc_00E243BB: push eax
  loc_00E243BC: lea eax, var_28
  loc_00E243BF: lea ecx, var_140
  loc_00E243C5: push eax
  loc_00E243C6: lea edx, var_B0
  loc_00E243CC: push ecx
  loc_00E243CD: push edx
  loc_00E243CE: call __vbaVarCmpEq
  loc_00E243D0: push eax
  loc_00E243D1: lea eax, var_C0
  loc_00E243D7: push eax
  loc_00E243D8: call edi
  loc_00E243DA: lea ecx, var_28
  loc_00E243DD: push eax
  loc_00E243DE: lea edx, var_150
  loc_00E243E4: push ecx
  loc_00E243E5: lea eax, var_D0
  loc_00E243EB: push edx
  loc_00E243EC: push eax
  loc_00E243ED: call __vbaVarCmpEq
  loc_00E243EF: lea ecx, var_E0
  loc_00E243F5: push eax
  loc_00E243F6: push ecx
  loc_00E243F7: call edi
  loc_00E243F9: push eax
  loc_00E243FA: lea edx, var_28
  loc_00E243FD: lea eax, var_160
  loc_00E24403: push edx
  loc_00E24404: lea ecx, var_F0
  loc_00E2440A: push eax
  loc_00E2440B: push ecx
  loc_00E2440C: call __vbaVarCmpEq
  loc_00E2440E: lea edx, var_100
  loc_00E24414: push eax
  loc_00E24415: push edx
  loc_00E24416: call edi
  loc_00E24418: push eax
  loc_00E24419: lea eax, var_28
  loc_00E2441C: lea ecx, var_170
  loc_00E24422: push eax
  loc_00E24423: lea edx, var_110
  loc_00E24429: push ecx
  loc_00E2442A: push edx
  loc_00E2442B: call __vbaVarCmpEq
  loc_00E2442D: push eax
  loc_00E2442E: lea eax, var_120
  loc_00E24434: push eax
  loc_00E24435: call edi
  loc_00E24437: push eax
  loc_00E24438: call [004010DCh] ; __vbaBoolVarNull
  loc_00E2443E: test ax, ax
  loc_00E24441: jz 00E2434Dh
  loc_00E24447: lea ecx, var_88
  loc_00E2444D: lea edx, var_130
  loc_00E24453: push ecx
  loc_00E24454: lea eax, var_A0
  loc_00E2445A: push edx
  loc_00E2445B: push eax
  loc_00E2445C: mov var_128, 00000001h
  loc_00E24466: mov var_130, 00000002h
  loc_00E24470: call [004011E0h] ; __vbaVarAdd
  loc_00E24476: mov edx, eax
  loc_00E24478: lea ecx, var_88
  loc_00E2447E: call [0040101Ch] ; __vbaVarMove
  loc_00E24484: lea edx, var_A0
  loc_00E2448A: lea eax, var_88
  loc_00E24490: lea ecx, var_34
  loc_00E24493: push edx
  loc_00E24494: push eax
  loc_00E24495: mov var_98, 00000001h
  loc_00E2449F: mov var_A0, 00000002h
  loc_00E244A9: mov var_128, ecx
  loc_00E244AF: mov var_130, 00004008h
  loc_00E244B9: call [004011D0h] ; __vbaI4Var
  loc_00E244BF: lea ecx, var_130
  loc_00E244C5: push eax
  loc_00E244C6: lea edx, var_B0
  loc_00E244CC: push ecx
  loc_00E244CD: push edx
  loc_00E244CE: call [004010F0h] ; rtcMidCharVar
  loc_00E244D4: lea eax, var_B0
  loc_00E244DA: push eax
  loc_00E244DB: call [00401034h] ; __vbaStrVarMove
  loc_00E244E1: mov edx, eax
  loc_00E244E3: lea ecx, var_8C
  loc_00E244E9: call [00401228h] ; __vbaStrMove
  loc_00E244EF: lea ecx, var_B0
  loc_00E244F5: lea edx, var_A0
  loc_00E244FB: push ecx
  loc_00E244FC: push edx
  loc_00E244FD: push 00000002h
  loc_00E244FF: call [00401038h] ; __vbaFreeVarList
  loc_00E24505: mov eax, var_34
  loc_00E24508: add esp, 0000000Ch
  loc_00E2450B: push eax
  loc_00E2450C: call [0040102Ch] ; __vbaLenBstr
  loc_00E24512: lea ecx, var_130
  loc_00E24518: mov var_128, eax
  loc_00E2451E: lea edx, var_88
  loc_00E24524: push ecx
  loc_00E24525: lea eax, var_A0
  loc_00E2452B: push edx
  loc_00E2452C: push eax
  loc_00E2452D: mov var_130, 00000003h
  loc_00E24537: mov var_138, 00000001h
  loc_00E24541: mov var_140, 00000002h
  loc_00E2454B: call [00401004h] ; __vbaVarSub
  loc_00E24551: lea ecx, var_140
  loc_00E24557: push eax
  loc_00E24558: lea edx, var_B0
  loc_00E2455E: push ecx
  loc_00E2455F: push edx
  loc_00E24560: call [004011E0h] ; __vbaVarAdd
  loc_00E24566: mov edx, eax
  loc_00E24568: lea ecx, var_28
  loc_00E2456B: call [0040101Ch] ; __vbaVarMove
  loc_00E24571: mov edx, 006DCC80h
  loc_00E24576: lea ecx, var_50
  loc_00E24579: call [004011B0h] ; __vbaStrCopy
  loc_00E2457F: mov eax, var_8C
  loc_00E24585: push eax
  loc_00E24586: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E2458C: fstp real8 ptr var_19C
  loc_00E24592: fld real8 ptr var_19C
  loc_00E24598: fcomp real8 ptr [004022E8h]
  loc_00E2459E: fnstsw ax
  loc_00E245A0: test ah, 40h
  loc_00E245A3: jz 00E245AFh
  loc_00E245A5: mov edx, 006E0EF0h ; "Sepuluh "
  loc_00E245AA: jmp 00E2473Eh
  loc_00E245AF: mov ecx, var_19C
  loc_00E245B5: mov eax, var_198
  loc_00E245BB: test ecx, ecx
  loc_00E245BD: jnz 00E245D0h
  loc_00E245BF: cmp eax, 3FF00000h
  loc_00E245C4: jnz 00E245D0h
  loc_00E245C6: mov edx, 006E0F08h ; "Sebelas "
  loc_00E245CB: jmp 00E2473Eh
  loc_00E245D0: test ecx, ecx
  loc_00E245D2: jnz 00E24747h
  loc_00E245D8: cmp eax, 40000000h
  loc_00E245DD: jnz 00E245E9h
  loc_00E245DF: mov edx, 006E0F20h ; "Dua Belas "
  loc_00E245E4: jmp 00E2473Eh
  loc_00E245E9: test ecx, ecx
  loc_00E245EB: jnz 00E24747h
  loc_00E245F1: cmp eax, 40080000h
  loc_00E245F6: jnz 00E24602h
  loc_00E245F8: mov edx, 006E0F4Ch ; "Tiga Belas "
  loc_00E245FD: jmp 00E2473Eh
  loc_00E24602: test ecx, ecx
  loc_00E24604: jnz 00E24747h
  loc_00E2460A: cmp eax, 40100000h
  loc_00E2460F: jnz 00E2461Bh
  loc_00E24611: mov edx, 006E0F68h ; "Empat Belas "
  loc_00E24616: jmp 00E2473Eh
  loc_00E2461B: test ecx, ecx
  loc_00E2461D: jnz 00E24747h
  loc_00E24623: cmp eax, 40140000h
  loc_00E24628: jnz 00E24634h
  loc_00E2462A: mov edx, 006E0F88h ; "Lima Belas "
  loc_00E2462F: jmp 00E2473Eh
  loc_00E24634: test ecx, ecx
  loc_00E24636: jnz 00E24747h
  loc_00E2463C: cmp eax, 40180000h
  loc_00E24641: jnz 00E2464Dh
  loc_00E24643: mov edx, 006E0FA4h ; "Enam Belas "
  loc_00E24648: jmp 00E2473Eh
  loc_00E2464D: test ecx, ecx
  loc_00E2464F: jnz 00E24747h
  loc_00E24655: cmp eax, 401C0000h
  loc_00E2465A: jnz 00E24666h
  loc_00E2465C: mov edx, 006E0FC0h ; "Tujuh Belas "
  loc_00E24661: jmp 00E2473Eh
  loc_00E24666: test ecx, ecx
  loc_00E24668: jnz 00E24747h
  loc_00E2466E: cmp eax, 40200000h
  loc_00E24673: jnz 00E2467Fh
  loc_00E24675: mov edx, 006E0FE0h ; "Delapan Belas "
  loc_00E2467A: jmp 00E2473Eh
  loc_00E2467F: test ecx, ecx
  loc_00E24681: jnz 00E24747h
  loc_00E24687: cmp eax, 40220000h
  loc_00E2468C: jnz 00E24747h
  loc_00E24692: mov edx, 006E1004h ; "Sembilan Belas "
  loc_00E24697: jmp 00E2473Eh
  loc_00E2469C: test eax, eax
  loc_00E2469E: jnz 00E24739h
  loc_00E246A4: cmp ecx, 40000000h
  loc_00E246AA: jnz 00E246B6h
  loc_00E246AC: mov edx, 006E1028h ; "Dua "
  loc_00E246B1: jmp 00E2473Eh
  loc_00E246B6: test eax, eax
  loc_00E246B8: jnz 00E24739h
  loc_00E246BA: cmp ecx, 40080000h
  loc_00E246C0: jnz 00E246C9h
  loc_00E246C2: mov edx, 006E1038h ; "Tiga "
  loc_00E246C7: jmp 00E2473Eh
  loc_00E246C9: test eax, eax
  loc_00E246CB: jnz 00E24739h
  loc_00E246CD: cmp ecx, 40100000h
  loc_00E246D3: jnz 00E246DCh
  loc_00E246D5: mov edx, 006E1048h ; "Empat "
  loc_00E246DA: jmp 00E2473Eh
  loc_00E246DC: test eax, eax
  loc_00E246DE: jnz 00E24739h
  loc_00E246E0: cmp ecx, 40140000h
  loc_00E246E6: jnz 00E246EFh
  loc_00E246E8: mov edx, 006E105Ch ; "Lima "
  loc_00E246ED: jmp 00E2473Eh
  loc_00E246EF: test eax, eax
  loc_00E246F1: jnz 00E24739h
  loc_00E246F3: cmp ecx, 40180000h
  loc_00E246F9: jnz 00E24702h
  loc_00E246FB: mov edx, 006E106Ch ; "Enam "
  loc_00E24700: jmp 00E2473Eh
  loc_00E24702: test eax, eax
  loc_00E24704: jnz 00E24739h
  loc_00E24706: cmp ecx, 401C0000h
  loc_00E2470C: jnz 00E24715h
  loc_00E2470E: mov edx, 006E107Ch ; "Tujuh "
  loc_00E24713: jmp 00E2473Eh
  loc_00E24715: test eax, eax
  loc_00E24717: jnz 00E24739h
  loc_00E24719: cmp ecx, 40200000h
  loc_00E2471F: jnz 00E24728h
  loc_00E24721: mov edx, 006E1090h ; "Delapan "
  loc_00E24726: jmp 00E2473Eh
  loc_00E24728: test eax, eax
  loc_00E2472A: jnz 00E24739h
  loc_00E2472C: cmp ecx, 40220000h
  loc_00E24732: mov edx, 006E10A8h ; "Sembilan "
  loc_00E24737: jz 00E2473Eh
  loc_00E24739: mov edx, 006DCC80h
  loc_00E2473E: lea ecx, var_4C
  loc_00E24741: call [004011B0h] ; __vbaStrCopy
  loc_00E24747: mov ecx, var_8C
  loc_00E2474D: push ecx
  loc_00E2474E: call [0040125Ch] ; rtcR8ValFromBstr
  loc_00E24754: call [004010D8h] ; __vbaFpR8
  loc_00E2475A: fcomp real8 ptr [004022E8h]
  loc_00E24760: fnstsw ax
  loc_00E24762: test ah, 41h
  loc_00E24765: jnz 00E24952h
  loc_00E2476B: lea edx, var_28
  loc_00E2476E: lea eax, var_130
  loc_00E24774: push edx
  loc_00E24775: lea ecx, var_A0
  loc_00E2477B: push eax
  loc_00E2477C: push ecx
  loc_00E2477D: mov var_128, 00000002h
  loc_00E24787: mov var_130, ebx
  loc_00E2478D: mov var_138, 00000005h
  loc_00E24797: mov var_140, ebx
  loc_00E2479D: mov var_148, 00000008h
  loc_00E247A7: mov var_150, ebx
  loc_00E247AD: mov var_158, 0000000Bh
  loc_00E247B7: mov var_160, ebx
  loc_00E247BD: mov var_168, 0000000Eh
  loc_00E247C7: mov var_170, ebx
  loc_00E247CD: call __vbaVarCmpEq
  loc_00E247CF: push eax
  loc_00E247D0: lea edx, var_28
  loc_00E247D3: lea eax, var_140
  loc_00E247D9: push edx
  loc_00E247DA: lea ecx, var_B0
  loc_00E247E0: push eax
  loc_00E247E1: push ecx
  loc_00E247E2: call __vbaVarCmpEq
  loc_00E247E4: lea edx, var_C0
  loc_00E247EA: push eax
  loc_00E247EB: push edx
  loc_00E247EC: call edi
  loc_00E247EE: push eax
  loc_00E247EF: lea eax, var_28
  loc_00E247F2: lea ecx, var_150
  loc_00E247F8: push eax
  loc_00E247F9: lea edx, var_D0
  loc_00E247FF: push ecx
  loc_00E24800: push edx
  loc_00E24801: call __vbaVarCmpEq
  loc_00E24803: push eax
  loc_00E24804: lea eax, var_E0
  loc_00E2480A: push eax
  loc_00E2480B: call edi
  loc_00E2480D: lea ecx, var_28
  loc_00E24810: push eax
  loc_00E24811: lea edx, var_160
  loc_00E24817: push ecx
  loc_00E24818: lea eax, var_F0
  loc_00E2481E: push edx
  loc_00E2481F: push eax
  loc_00E24820: call __vbaVarCmpEq
  loc_00E24822: lea ecx, var_100
  loc_00E24828: push eax
  loc_00E24829: push ecx
  loc_00E2482A: call edi
  loc_00E2482C: push eax
  loc_00E2482D: lea edx, var_28
  loc_00E24830: lea eax, var_170
  loc_00E24836: push edx
  loc_00E24837: lea ecx, var_110
  loc_00E2483D: push eax
  loc_00E2483E: push ecx
  loc_00E2483F: call __vbaVarCmpEq
  loc_00E24841: lea edx, var_120
  loc_00E24847: push eax
  loc_00E24848: push edx
  loc_00E24849: call edi
  loc_00E2484B: push eax
  loc_00E2484C: call [004010DCh] ; __vbaBoolVarNull
  loc_00E24852: test ax, ax
  loc_00E24855: jz 00E24861h
  loc_00E24857: mov edx, 006E10C0h ; "Puluh "
  loc_00E2485C: jmp 00E24957h
  loc_00E24861: lea eax, var_28
  loc_00E24864: lea ecx, var_130
  loc_00E2486A: push eax
  loc_00E2486B: lea edx, var_A0
  loc_00E24871: push ecx
  loc_00E24872: push edx
  loc_00E24873: mov var_128, 00000003h
  loc_00E2487D: mov var_130, ebx
  loc_00E24883: mov var_138, 00000006h
  loc_00E2488D: mov var_140, ebx
  loc_00E24893: mov var_148, 00000009h
  loc_00E2489D: mov var_150, ebx
  loc_00E248A3: mov var_158, 0000000Ch
  loc_00E248AD: mov var_160, ebx
  loc_00E248B3: mov var_168, 0000000Fh
  loc_00E248BD: mov var_170, ebx
  loc_00E248C3: call __vbaVarCmpEq
  loc_00E248C5: push eax
  loc_00E248C6: lea eax, var_28
  loc_00E248C9: lea ecx, var_140
  loc_00E248CF: push eax
  loc_00E248D0: lea edx, var_B0
  loc_00E248D6: push ecx
  loc_00E248D7: push edx
  loc_00E248D8: call __vbaVarCmpEq
  loc_00E248DA: push eax
  loc_00E248DB: lea eax, var_C0
  loc_00E248E1: push eax
  loc_00E248E2: call edi
  loc_00E248E4: lea ecx, var_28
  loc_00E248E7: push eax
  loc_00E248E8: lea edx, var_150
  loc_00E248EE: push ecx
  loc_00E248EF: lea eax, var_D0
  loc_00E248F5: push edx
  loc_00E248F6: push eax
  loc_00E248F7: call __vbaVarCmpEq
  loc_00E248F9: lea ecx, var_E0
  loc_00E248FF: push eax
  loc_00E24900: push ecx
  loc_00E24901: call edi
  loc_00E24903: push eax
  loc_00E24904: lea edx, var_28
  loc_00E24907: lea eax, var_160
  loc_00E2490D: push edx
  loc_00E2490E: lea ecx, var_F0
  loc_00E24914: push eax
  loc_00E24915: push ecx
  loc_00E24916: call __vbaVarCmpEq
  loc_00E24918: lea edx, var_100
  loc_00E2491E: push eax
  loc_00E2491F: push edx
  loc_00E24920: call edi
  loc_00E24922: push eax
  loc_00E24923: lea eax, var_28
  loc_00E24926: lea ecx, var_170
  loc_00E2492C: push eax
  loc_00E2492D: lea edx, var_110
  loc_00E24933: push ecx
  loc_00E24934: push edx
  loc_00E24935: call __vbaVarCmpEq
  loc_00E24937: push eax
  loc_00E24938: lea eax, var_120
  loc_00E2493E: push eax
  loc_00E2493F: call edi
  loc_00E24941: push eax
  loc_00E24942: call [004010DCh] ; __vbaBoolVarNull
  loc_00E24948: test ax, ax
  loc_00E2494B: mov edx, 006E10D4h ; "Ratus "
  loc_00E24950: jnz 00E24957h
  loc_00E24952: mov edx, 006DCC80h
  loc_00E24957: lea ecx, var_50
  loc_00E2495A: call [004011B0h] ; __vbaStrCopy
  loc_00E24960: lea ecx, var_70
  loc_00E24963: lea edx, var_130
  loc_00E24969: push ecx
  loc_00E2496A: push edx
  loc_00E2496B: mov var_128, 00000000h
  loc_00E24975: mov var_130, ebx
  loc_00E2497B: call [00401008h] ; __vbaVarTstGt
  loc_00E24981: test ax, ax
  loc_00E24984: jz 00E24A9Eh
  loc_00E2498A: lea edx, var_28
  loc_00E2498D: lea ecx, var_184
  loc_00E24993: call [00401204h] ; __vbaVarCopy
  loc_00E24999: lea eax, var_184
  loc_00E2499F: lea ecx, var_130
  loc_00E249A5: push eax
  loc_00E249A6: push ecx
  loc_00E249A7: mov var_128, 00000004h
  loc_00E249B1: mov var_130, ebx
  loc_00E249B7: call [00401108h] ; __vbaVarTstEq
  loc_00E249BD: test ax, ax
  loc_00E249C0: jz 00E249D0h
  loc_00E249C2: mov edx, var_50
  loc_00E249C5: push edx
  loc_00E249C6: push 006E10E8h ; "Ribu "
  loc_00E249CB: jmp 00E24A6Ah
  loc_00E249D0: lea eax, var_184
  loc_00E249D6: lea ecx, var_130
  loc_00E249DC: push eax
  loc_00E249DD: push ecx
  loc_00E249DE: mov var_128, 00000007h
  loc_00E249E8: mov var_130, ebx
  loc_00E249EE: call [00401108h] ; __vbaVarTstEq
  loc_00E249F4: test ax, ax
  loc_00E249F7: jz 00E24A04h
  loc_00E249F9: mov edx, var_50
  loc_00E249FC: push edx
  loc_00E249FD: push 006E0F3Ch ; "Juta "
  loc_00E24A02: jmp 00E24A6Ah
  loc_00E24A04: lea eax, var_184
  loc_00E24A0A: lea ecx, var_130
  loc_00E24A10: push eax
  loc_00E24A11: push ecx
  loc_00E24A12: mov var_128, 0000000Ah
  loc_00E24A1C: mov var_130, ebx
  loc_00E24A22: call [00401108h] ; __vbaVarTstEq
  loc_00E24A28: test ax, ax
  loc_00E24A2B: jz 00E24A38h
  loc_00E24A2D: mov edx, var_50
  loc_00E24A30: push edx
  loc_00E24A31: push 006E10F8h ; "Milyar "
  loc_00E24A36: jmp 00E24A6Ah
  loc_00E24A38: lea eax, var_184
  loc_00E24A3E: lea ecx, var_130
  loc_00E24A44: push eax
  loc_00E24A45: push ecx
  loc_00E24A46: mov var_128, 0000000Dh
  loc_00E24A50: mov var_130, ebx
  loc_00E24A56: call [00401108h] ; __vbaVarTstEq
  loc_00E24A5C: test ax, ax
  loc_00E24A5F: jz 00E24A9Eh
  loc_00E24A61: mov edx, var_50
  loc_00E24A64: push edx
  loc_00E24A65: push 006E110Ch ; "Trilyun "
  loc_00E24A6A: call [0040106Ch] ; __vbaStrCat
  loc_00E24A70: mov edx, eax
  loc_00E24A72: lea ecx, var_50
  loc_00E24A75: call [00401228h] ; __vbaStrMove
  loc_00E24A7B: lea edx, var_130
  loc_00E24A81: lea ecx, var_70
  loc_00E24A84: mov var_128, 00000000h
  loc_00E24A8E: mov var_130, 00000002h
  loc_00E24A98: call [0040101Ch] ; __vbaVarMove
  loc_00E24A9E: mov eax, var_2C
  loc_00E24AA1: mov ecx, var_4C
  loc_00E24AA4: push eax
  loc_00E24AA5: push ecx
  loc_00E24AA6: call [0040106Ch] ; __vbaStrCat
  loc_00E24AAC: mov edx, eax
  loc_00E24AAE: lea ecx, var_90
  loc_00E24AB4: call [00401228h] ; __vbaStrMove
  loc_00E24ABA: mov edx, var_50
  loc_00E24ABD: push eax
  loc_00E24ABE: push edx
  loc_00E24ABF: call [0040106Ch] ; __vbaStrCat
  loc_00E24AC5: mov edx, eax
  loc_00E24AC7: lea ecx, var_2C
  loc_00E24ACA: call [00401228h] ; __vbaStrMove
  loc_00E24AD0: lea ecx, var_90
  loc_00E24AD6: call [00401258h] ; __vbaFreeStr
  loc_00E24ADC: jmp 00E24063h
  loc_00E24AE1: mov eax, var_2C
  loc_00E24AE4: mov ecx, var_48
  loc_00E24AE7: mov edi, [0040106Ch] ; __vbaStrCat
  loc_00E24AED: push eax
  loc_00E24AEE: push ecx
  loc_00E24AEF: call edi
  loc_00E24AF1: mov esi, [00401228h] ; __vbaStrMove
  loc_00E24AF7: mov edx, eax
  loc_00E24AF9: lea ecx, var_2C
  loc_00E24AFC: call __vbaStrMove
  loc_00E24AFE: mov edx, var_2C
  loc_00E24B01: push edx
  loc_00E24B02: push 006E1124h ; "Rupiah "
  loc_00E24B07: call edi
  loc_00E24B09: mov edx, eax
  loc_00E24B0B: lea ecx, var_78
  loc_00E24B0E: call __vbaStrMove
  loc_00E24B10: lea ecx, var_130
  loc_00E24B16: lea edx, var_A0
  loc_00E24B1C: lea eax, var_78
  loc_00E24B1F: mov ebx, 00004008h
  loc_00E24B24: push ecx
  loc_00E24B25: push edx
  loc_00E24B26: mov var_128, eax
  loc_00E24B2C: mov var_130, ebx
  loc_00E24B32: call [00401054h] ; rtcLowerCaseVar
  loc_00E24B38: mov edi, [00401034h] ; __vbaStrVarMove
  loc_00E24B3E: lea eax, var_A0
  loc_00E24B44: push eax
  loc_00E24B45: call edi
  loc_00E24B47: mov edx, eax
  loc_00E24B49: lea ecx, var_78
  loc_00E24B4C: call __vbaStrMove
  loc_00E24B4E: lea ecx, var_A0
  loc_00E24B54: call [00401024h] ; __vbaFreeVar
  loc_00E24B5A: lea edx, var_130
  loc_00E24B60: push 00000001h
  loc_00E24B62: lea eax, var_A0
  loc_00E24B68: lea ecx, var_78
  loc_00E24B6B: push edx
  loc_00E24B6C: push eax
  loc_00E24B6D: mov var_128, ecx
  loc_00E24B73: mov var_130, ebx
  loc_00E24B79: call [0040121Ch] ; rtcLeftCharVar
  loc_00E24B7F: lea ecx, var_A0
  loc_00E24B85: lea edx, var_B0
  loc_00E24B8B: push ecx
  loc_00E24B8C: push edx
  loc_00E24B8D: call [004010FCh] ; rtcUpperCaseVar
  loc_00E24B93: lea eax, var_B0
  loc_00E24B99: push eax
  loc_00E24B9A: call edi
  loc_00E24B9C: mov edx, eax
  loc_00E24B9E: lea ecx, var_74
  loc_00E24BA1: call __vbaStrMove
  loc_00E24BA3: mov ebx, [00401038h] ; __vbaFreeVarList
  loc_00E24BA9: lea ecx, var_B0
  loc_00E24BAF: lea edx, var_A0
  loc_00E24BB5: push ecx
  loc_00E24BB6: push edx
  loc_00E24BB7: push 00000002h
  loc_00E24BB9: call ebx
  loc_00E24BBB: mov ecx, var_78
  loc_00E24BBE: mov eax, var_74
  loc_00E24BC1: add esp, 0000000Ch
  loc_00E24BC4: mov var_148, eax
  loc_00E24BCA: mov var_150, 00000008h
  loc_00E24BD4: push ecx
  loc_00E24BD5: call [0040102Ch] ; __vbaLenBstr
  loc_00E24BDB: sub eax, 00000001h
  loc_00E24BDE: lea edx, var_78
  loc_00E24BE1: jo 00E24D51h
  loc_00E24BE7: mov var_98, eax
  loc_00E24BED: lea eax, var_A0
  loc_00E24BF3: push eax
  loc_00E24BF4: lea ecx, var_130
  loc_00E24BFA: push 00000002h
  loc_00E24BFC: mov var_A0, 00000003h
  loc_00E24C06: mov var_128, edx
  loc_00E24C0C: mov var_130, 00004008h
  loc_00E24C16: push ecx
  loc_00E24C17: lea edx, var_B0
  loc_00E24C1D: push edx
  loc_00E24C1E: call [004010F0h] ; rtcMidCharVar
  loc_00E24C24: lea eax, var_150
  loc_00E24C2A: lea ecx, var_B0
  loc_00E24C30: push eax
  loc_00E24C31: lea edx, var_C0
  loc_00E24C37: push ecx
  loc_00E24C38: push edx
  loc_00E24C39: call [00401184h] ; __vbaVarCat
  loc_00E24C3F: push eax
  loc_00E24C40: call edi
  loc_00E24C42: mov edx, eax
  loc_00E24C44: lea ecx, var_78
  loc_00E24C47: call __vbaStrMove
  loc_00E24C49: lea eax, var_C0
  loc_00E24C4F: lea ecx, var_B0
  loc_00E24C55: push eax
  loc_00E24C56: lea edx, var_A0
  loc_00E24C5C: push ecx
  loc_00E24C5D: push edx
  loc_00E24C5E: push 00000003h
  loc_00E24C60: call ebx
  loc_00E24C62: add esp, 00000010h
  loc_00E24C65: fwait
  loc_00E24C66: push 00E24D2Ah
  loc_00E24C6B: jmp 00E24CD3h
  loc_00E24C6D: test var_4, 04h
  loc_00E24C71: jz 00E24C7Ch
  loc_00E24C73: lea ecx, var_78
  loc_00E24C76: call [00401258h] ; __vbaFreeStr
  loc_00E24C7C: lea ecx, var_90
  loc_00E24C82: call [00401258h] ; __vbaFreeStr
  loc_00E24C88: lea eax, var_120
  loc_00E24C8E: lea ecx, var_110
  loc_00E24C94: push eax
  loc_00E24C95: lea edx, var_100
  loc_00E24C9B: push ecx
  loc_00E24C9C: lea eax, var_F0
  loc_00E24CA2: push edx
  loc_00E24CA3: lea ecx, var_E0
  loc_00E24CA9: push eax
  loc_00E24CAA: lea edx, var_D0
  loc_00E24CB0: push ecx
  loc_00E24CB1: lea eax, var_C0
  loc_00E24CB7: push edx
  loc_00E24CB8: lea ecx, var_B0
  loc_00E24CBE: push eax
  loc_00E24CBF: lea edx, var_A0
  loc_00E24CC5: push ecx
  loc_00E24CC6: push edx
  loc_00E24CC7: push 00000009h
  loc_00E24CC9: call [00401038h] ; __vbaFreeVarList
  loc_00E24CCF: add esp, 00000028h
  loc_00E24CD2: ret
  loc_00E24CD3: mov edi, [00401024h] ; __vbaFreeVar
  loc_00E24CD9: lea ecx, var_184
  loc_00E24CDF: call edi
  loc_00E24CE1: lea ecx, var_28
  loc_00E24CE4: call edi
  loc_00E24CE6: mov esi, [00401258h] ; __vbaFreeStr
  loc_00E24CEC: lea ecx, var_2C
  loc_00E24CEF: call __vbaFreeStr
  loc_00E24CF1: lea ecx, var_34
  loc_00E24CF4: call __vbaFreeStr
  loc_00E24CF6: lea ecx, var_44
  loc_00E24CF9: call edi
  loc_00E24CFB: lea ecx, var_48
  loc_00E24CFE: call __vbaFreeStr
  loc_00E24D00: lea ecx, var_4C
  loc_00E24D03: call __vbaFreeStr
  loc_00E24D05: lea ecx, var_50
  loc_00E24D08: call __vbaFreeStr
  loc_00E24D0A: lea ecx, var_60
  loc_00E24D0D: call edi
  loc_00E24D0F: lea ecx, var_70
  loc_00E24D12: call edi
  loc_00E24D14: lea ecx, var_74
  loc_00E24D17: call __vbaFreeStr
  loc_00E24D19: lea ecx, var_88
  loc_00E24D1F: call edi
  loc_00E24D21: lea ecx, var_8C
  loc_00E24D27: call __vbaFreeStr
  loc_00E24D29: ret
  loc_00E24D2A: mov eax, Me
  loc_00E24D2D: push eax
  loc_00E24D2E: mov ecx, [eax]
  loc_00E24D30: call [ecx+00000008h]
  loc_00E24D33: mov edx, arg_10
  loc_00E24D36: mov eax, var_78
  loc_00E24D39: mov [edx], eax
  loc_00E24D3B: mov eax, var_4
  loc_00E24D3E: mov ecx, var_14
  loc_00E24D41: pop edi
  loc_00E24D42: pop esi
  loc_00E24D43: mov fs:[00000000h], ecx
  loc_00E24D4A: pop ebx
  loc_00E24D4B: mov esp, ebp
  loc_00E24D4D: pop ebp
  loc_00E24D4E: retn 000Ch
End Function

Private Sub Proc_8_19_E24EB0(arg_C) 'E24EB0
  loc_00E24EB0: mov eax, [00E3D0ECh]
  loc_00E24EB5: sub esp, 00000010h
  loc_00E24EB8: test eax, eax
  loc_00E24EBA: push ebx
  loc_00E24EBB: push ebp
  loc_00E24EBC: push esi
  loc_00E24EBD: push edi
  loc_00E24EBE: jnz 00E24ED0h
  loc_00E24EC0: push 00E3D0ECh
  loc_00E24EC5: push 006CC808h
  loc_00E24ECA: call [004011A0h] ; __vbaNew2
  loc_00E24ED0: sub esp, 00000010h
  loc_00E24ED3: mov ecx, 0000000Ah
  loc_00E24ED8: mov ebp, esp
  loc_00E24EDA: mov edi, ecx
  loc_00E24EDC: mov eax, 80020004h
  loc_00E24EE1: sub esp, 00000010h
  loc_00E24EE4: mov [ebp], ecx
  loc_00E24EE7: mov ecx, arg_2C
  loc_00E24EEB: mov esi, [00E3D0ECh]
  loc_00E24EF1: mov edx, eax
  loc_00E24EF3: mov Proc_8_19_E24EB0, ecx
  loc_00E24EF6: mov ecx, esp
  loc_00E24EF8: mov ebx, [esi]
  loc_00E24EFA: push esi
  loc_00E24EFB: mov Me, eax
  loc_00E24EFE: mov eax, arg_38
  loc_00E24F02: mov arg_C, eax
  loc_00E24F05: mov eax, arg_30
  loc_00E24F09: mov [ecx], edi
  loc_00E24F0B: mov [ecx+00000004h], eax
  loc_00E24F0E: mov [ecx+00000008h], edx
  loc_00E24F11: mov edx, arg_38
  loc_00E24F15: mov [ecx+0000000Ch], edx
  loc_00E24F18: call [ebx+000002B0h]
  loc_00E24F1E: test eax, eax
  loc_00E24F20: fnclex
  loc_00E24F22: jge 00E24F36h
  loc_00E24F24: push 000002B0h
  loc_00E24F29: push 006E00E8h
  loc_00E24F2E: push esi
  loc_00E24F2F: push eax
  loc_00E24F30: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E24F36: pop edi
  loc_00E24F37: pop esi
  loc_00E24F38: pop ebp
  loc_00E24F39: xor eax, eax
  loc_00E24F3B: pop ebx
  loc_00E24F3C: add esp, 00000010h
  loc_00E24F3F: retn 0004h
End Sub
