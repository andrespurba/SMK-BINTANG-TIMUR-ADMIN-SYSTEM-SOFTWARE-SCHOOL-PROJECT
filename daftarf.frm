VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC0000F8754DA1}#2.0#0"; "C:\Windows\SysWow64\MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B14800A0C922E820}#6.0#0"; "C:\Windows\SysWow64\MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C600A0C90AEA82}#1.0#0"; "C:\Windows\SysWow64\MSDATGRD.OCX"
Begin VB.Form daftarf
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
  ClientWidth = 14670
  ClientHeight = 9705
  StartUpPosition = 2 'CenterScreen
  Begin VB.Frame Frame2
    Caption = "Frame2"
    BackColor = &H4000&
    Left = 30
    Top = 1890
    Width = 1875
    Height = 435
    TabIndex = 68
    BorderStyle = 0 'None
    Begin Project1.jcbutton tgl
      Left = 1500
      Top = 60
      Width = 315
      Height = 345
      TabIndex = 69
      OleObjectBlob = "daftarf.frx":0000
    End
    Begin VB.TextBox daftar
      BackColor = &HFFFFFF&
      Left = 60
      Top = 60
      Width = 1665
      Height = 345
      TabIndex = 70
      BorderStyle = 0 'None
      Alignment = 2 'Center
    End
  End
  Begin VB.Timer Timer3
    Interval = 100
    Left = 2160
    Top = 9210
  End
  Begin VB.ComboBox cmbjur
    BackColor = &H404040&
    ForeColor = &H80FF80&
    Left = 7350
    Top = 7320
    Width = 2235
    Height = 420
    Text = "-----Pilih"
    TabIndex = 66
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
    Caption = "Cari Disini"
    BackColor = &HE0E0E0&
    ForeColor = &H404040&
    Left = 10830
    Top = 6480
    Width = 3315
    Height = 1305
    TabIndex = 57
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
      Left = 1680
      Top = 240
      Width = 1575
      Height = 300
      TabIndex = 71
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
    Begin Project1.jcbutton jcbutton1
      Left = 3060
      Top = 0
      Width = 225
      Height = 225
      TabIndex = 62
      OleObjectBlob = "daftarf.frx":020C
    End
    Begin VB.OptionButton optnibu
      Caption = "Nama Ibu"
      BackColor = &HE0E0E0&
      ForeColor = &H404040&
      Left = 150
      Top = 510
      Width = 1455
      Height = 300
      TabIndex = 60
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
      TabIndex = 59
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
      TabIndex = 58
      BorderStyle = 0 'None
    End
  End
  Begin VB.Frame Frame1
    Caption = "Frame1"
    BackColor = &H4000&
    Left = 1980
    Top = 1890
    Width = 11055
    Height = 375
    TabIndex = 55
    BorderStyle = 0 'None
    Begin VB.Label Label54
      Caption = "Silahkan Input Data Calon Peseta Didik Secara lengkap."
      ForeColor = &H80FF80&
      Left = 270
      Top = 30
      Width = 7665
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
  End
  Begin VB.Timer Timer2
    Interval = 100
    Left = 1740
    Top = 9210
  End
  Begin VB.Timer Timer1
    Interval = 10
    Left = 1320
    Top = 9210
  End
  Begin VB.Frame fwali
    Caption = "Frame1"
    BackColor = &H404040&
    Left = 10980
    Top = 2430
    Width = 3165
    Height = 2985
    TabIndex = 45
    BorderStyle = 0 'None
    Begin VB.TextBox nawali
      BackColor = &HFFFFFF&
      Left = 1170
      Top = 510
      Width = 1605
      Height = 345
      TabIndex = 49
      BorderStyle = 0 'None
    End
    Begin VB.TextBox alamatwali
      BackColor = &HFFFFFF&
      Left = 1170
      Top = 1050
      Width = 1605
      Height = 345
      TabIndex = 48
      BorderStyle = 0 'None
    End
    Begin VB.TextBox nowali
      BackColor = &HFFFFFF&
      Left = 1170
      Top = 1560
      Width = 1605
      Height = 345
      TabIndex = 47
      BorderStyle = 0 'None
      MaxLength = 12
    End
    Begin VB.TextBox pwali
      BackColor = &HFFFFFF&
      Left = 1140
      Top = 2550
      Width = 1605
      Height = 345
      TabIndex = 46
      BorderStyle = 0 'None
    End
    Begin VB.Label Label24
      Caption = "Data Wali"
      BackColor = &H404040&
      ForeColor = &HFFFFFF&
      Left = 0
      Top = 0
      Width = 1995
      Height = 315
      TabIndex = 54
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
      Caption = "Nama"
      ForeColor = &H80FF80&
      Left = 30
      Top = 480
      Width = 1275
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
    Begin VB.Label Label26
      Caption = "Alamat"
      ForeColor = &H80FF80&
      Left = 60
      Top = 1020
      Width = 1185
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
    Begin VB.Label Label27
      Caption = "No.Telp"
      ForeColor = &H80FF80&
      Left = 60
      Top = 1530
      Width = 1275
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
    Begin VB.Label Label28
      Caption = "Pekerjaan Wali"
      ForeColor = &H80FF80&
      Left = 60
      Top = 2070
      Width = 1875
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
  End
  Begin VB.TextBox pibu
    BackColor = &HFFFFFF&
    Left = 8730
    Top = 6540
    Width = 1935
    Height = 345
    TabIndex = 43
    BorderStyle = 0 'None
  End
  Begin VB.TextBox payah
    BackColor = &HFFFFFF&
    Left = 8730
    Top = 5940
    Width = 1935
    Height = 345
    TabIndex = 40
    BorderStyle = 0 'None
  End
  Begin VB.TextBox nortu
    BackColor = &HFFFFFF&
    Left = 8730
    Top = 5010
    Width = 1935
    Height = 345
    TabIndex = 38
    BorderStyle = 0 'None
    MaxLength = 12
  End
  Begin VB.TextBox alamatortu
    BackColor = &HFFFFFF&
    Left = 8730
    Top = 4440
    Width = 1935
    Height = 345
    TabIndex = 36
    BorderStyle = 0 'None
  End
  Begin VB.TextBox nibu
    BackColor = &HFFFFFF&
    Left = 8730
    Top = 3840
    Width = 1935
    Height = 345
    TabIndex = 34
    BorderStyle = 0 'None
  End
  Begin VB.TextBox nayah
    BackColor = &HFFFFFF&
    Left = 8730
    Top = 3240
    Width = 1935
    Height = 345
    TabIndex = 31
    BorderStyle = 0 'None
  End
  Begin VB.ComboBox cmbagama
    BackColor = &H404040&
    ForeColor = &H80FF80&
    Left = 4200
    Top = 6300
    Width = 1935
    Height = 420
    Text = "-----Pilih"
    TabIndex = 30
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
  Begin VB.OptionButton optper
    Caption = "Perempuan"
    BackColor = &H404040&
    ForeColor = &HFFFFFF&
    Left = 5430
    Top = 3900
    Width = 1545
    Height = 330
    TabIndex = 27
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
  Begin VB.OptionButton optlaki
    Caption = "Laki-laki"
    BackColor = &H404040&
    ForeColor = &HFFFFFF&
    Left = 4110
    Top = 3870
    Width = 1455
    Height = 375
    TabIndex = 26
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
  Begin VB.TextBox alamat
    BackColor = &HFFFFFF&
    Left = 4200
    Top = 6930
    Width = 1935
    Height = 345
    TabIndex = 24
    BorderStyle = 0 'None
  End
  Begin MSComCtl2.DTPicker dttgl
    Left = 5100
    Top = 5700
    Width = 1035
    Height = 345
    TabIndex = 22
    OleObjectBlob = "daftarf.frx":042C
  End
  Begin VB.TextBox tempat
    BackColor = &HFFFFFF&
    Left = 4200
    Top = 5700
    Width = 885
    Height = 345
    TabIndex = 21
    BorderStyle = 0 'None
  End
  Begin VB.TextBox nisn
    BackColor = &HFFFFFF&
    Left = 4200
    Top = 5070
    Width = 1935
    Height = 345
    TabIndex = 19
    BorderStyle = 0 'None
    MaxLength = 12
  End
  Begin VB.TextBox asalsekolah
    BackColor = &HFFFFFF&
    Left = 4200
    Top = 4470
    Width = 1935
    Height = 345
    TabIndex = 16
    BorderStyle = 0 'None
  End
  Begin VB.TextBox nama
    BackColor = &HFFFFFF&
    Left = 4200
    Top = 3300
    Width = 1935
    Height = 345
    TabIndex = 13
    BorderStyle = 0 'None
  End
  Begin Project1.jcbutton pluss
    Left = 6180
    Top = 2760
    Width = 315
    Height = 345
    TabIndex = 11
    OleObjectBlob = "daftarf.frx":04EC
  End
  Begin VB.TextBox nomor
    BackColor = &HFFFFFF&
    Left = 4200
    Top = 2760
    Width = 1935
    Height = 345
    TabIndex = 10
    BorderStyle = 0 'None
  End
  Begin Project1.jcbutton simpan
    Left = 180
    Top = 2970
    Width = 1545
    Height = 705
    TabIndex = 4
    OleObjectBlob = "daftarf.frx":06F4
  End
  Begin MSAdodcLib.Adodc adopen
    Left = 180
    Top = 9210
    Width = 1200
    Height = 330
    Visible = 0   'False
    OleObjectBlob = "daftarf.frx":1150
  End
  Begin MSDataGridLib.DataGrid dgpen
    Left = 120
    Top = 8160
    Width = 14475
    Height = 1485
    TabIndex = 1
    OleObjectBlob = "daftarf.frx":148C
  End
  Begin Project1.jcbutton update
    Left = 180
    Top = 3780
    Width = 1545
    Height = 705
    TabIndex = 5
    OleObjectBlob = "daftarf.frx":1637
  End
  Begin Project1.jcbutton baru
    Left = 180
    Top = 4560
    Width = 1545
    Height = 705
    TabIndex = 6
    OleObjectBlob = "daftarf.frx":257F
  End
  Begin Project1.jcbutton hapus
    Left = 180
    Top = 5400
    Width = 1545
    Height = 705
    TabIndex = 7
    OleObjectBlob = "daftarf.frx":32F3
  End
  Begin Project1.jcbutton refreshh
    Left = 180
    Top = 6240
    Width = 1545
    Height = 735
    TabIndex = 9
    OleObjectBlob = "daftarf.frx":41CF
  End
  Begin Project1.jcbutton CARI
    Left = 12900
    Top = 7020
    Width = 1155
    Height = 735
    TabIndex = 61
    OleObjectBlob = "daftarf.frx":4FFB
  End
  Begin Project1.jcbutton jcbutton2
    Left = 4200
    Top = 4200
    Width = 675
    Height = 225
    TabIndex = 67
    OleObjectBlob = "daftarf.frx":6097
  End
  Begin VB.Shape sdaftar
    BorderColor = &HFF&
    Left = 30
    Top = 1890
    Width = 1905
    Height = 495
    BorderWidth = 2
  End
  Begin VB.Shape Shape13
    Left = 1920
    Top = 1890
    Width = 75
    Height = 5295
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Image back
    Picture = "daftarf.frx":62CB
    Left = 0
    Top = 0
    Width = 435
    Height = 450
    Stretch = -1  'True
  End
  Begin VB.Shape sback
    Left = 0
    Top = -30
    Width = 465
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &HFF&
    FillStyle = 0
  End
  Begin VB.Label Label1
    Caption = "Jurusan Yang Dipilih"
    ForeColor = &H4000&
    Left = 7320
    Top = 6930
    Width = 2445
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
  Begin VB.Label lbltgl
    Caption = "Labeltanggal"
    ForeColor = &H80FF80&
    Left = 12360
    Top = 90
    Width = 1275
    Height = 375
    TabIndex = 64
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
    Left = 13740
    Top = 90
    Width = 45
    Height = 285
    BorderStyle = 0 'Transparent
    FillColor = &HFF00&
    FillStyle = 0
  End
  Begin VB.Label lbljam
    Caption = "Labeljam"
    ForeColor = &H80FF80&
    Left = 13890
    Top = 90
    Width = 735
    Height = 375
    TabIndex = 63
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
  Begin VB.Shape sscari
    Left = 12750
    Top = 6900
    Width = 1455
    Height = 945
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Shape scari
    Left = 10770
    Top = 6420
    Width = 3435
    Height = 1425
    BorderStyle = 0 'Transparent
    FillColor = &HFF&
    FillStyle = 0
  End
  Begin VB.Shape Shape9
    Left = 13020
    Top = 2010
    Width = 225
    Height = 195
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &H80FFFF&
    FillStyle = 0
  End
  Begin VB.Shape ssimpan
    Left = 0
    Top = 2850
    Width = 1935
    Height = 945
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
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
  Begin VB.Image Image4
    Picture = "daftarf.frx":7D75
    Left = 13650
    Top = 630
    Width = 1305
    Height = 735
    Stretch = -1  'True
  End
  Begin VB.Image Image3
    Picture = "daftarf.frx":F24B
    Left = 150
    Top = 600
    Width = 915
    Height = 855
    Stretch = -1  'True
  End
  Begin VB.Label Label23
    Caption = "Ibu"
    ForeColor = &H80FF80&
    Left = 7350
    Top = 6540
    Width = 1815
    Height = 315
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
  Begin VB.Label Label22
    Caption = "Ayah"
    ForeColor = &H80FF80&
    Left = 7350
    Top = 5940
    Width = 1815
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
  Begin VB.Label Label21
    Caption = "Pekerjaan Orang Tua :"
    ForeColor = &H80FF80&
    Left = 7320
    Top = 5460
    Width = 2445
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
  Begin VB.Label Label20
    Caption = "No. Telp"
    ForeColor = &H80FF80&
    Left = 7320
    Top = 4980
    Width = 1215
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
  Begin VB.Label Label19
    Caption = "Alamat"
    ForeColor = &H80FF80&
    Left = 7320
    Top = 4410
    Width = 1215
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
  Begin VB.Label Label18
    Caption = "Ibu"
    ForeColor = &H80FF80&
    Left = 7350
    Top = 3840
    Width = 1815
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
  Begin VB.Label Label17
    Caption = "Ayah"
    ForeColor = &H80FF80&
    Left = 7350
    Top = 3240
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
  Begin VB.Label Label16
    Caption = "Nama Orang Tua :"
    ForeColor = &H80FF80&
    Left = 7350
    Top = 2760
    Width = 2085
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
  Begin VB.Label Label15
    Caption = "Data Orang Tua"
    BackColor = &H404040&
    ForeColor = &HFFFFFF&
    Left = 7350
    Top = 2370
    Width = 1995
    Height = 315
    TabIndex = 29
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
    Caption = "Data Siswa"
    BackColor = &H404040&
    ForeColor = &HFFFFFF&
    Left = 2250
    Top = 2370
    Width = 1245
    Height = 315
    TabIndex = 28
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
    Caption = "Alamat"
    ForeColor = &H80FF80&
    Left = 2250
    Top = 6900
    Width = 1815
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
  Begin VB.Label Label12
    Caption = "Agama"
    ForeColor = &H80FF80&
    Left = 2250
    Top = 6270
    Width = 1815
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
  Begin VB.Label Label11
    Caption = "Tempat,Tgl lahir"
    ForeColor = &H80FF80&
    Left = 2190
    Top = 5670
    Width = 1875
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
  Begin VB.Label Label10
    Caption = "NISN"
    ForeColor = &H80FF80&
    Left = 2250
    Top = 5070
    Width = 1815
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
  Begin VB.Label Label9
    Caption = "Asal Sekolah"
    ForeColor = &H80FF80&
    Left = 300
    Top = 2940
    Width = 1815
    Height = 315
    TabIndex = 17
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
    Caption = "Asal Sekolah"
    ForeColor = &H80FF80&
    Left = 2250
    Top = 4470
    Width = 1815
    Height = 315
    TabIndex = 15
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
    Caption = "Jenis Kelamin"
    ForeColor = &H80FF80&
    Left = 2250
    Top = 3870
    Width = 1815
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
  Begin VB.Label Label6
    Caption = "Nama"
    ForeColor = &H80FF80&
    Left = 2250
    Top = 3300
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
  Begin VB.Shape Shape14
    Left = -840
    Top = 7980
    Width = 15495
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Label Label5
    Caption = "Tombol Navigasi"
    ForeColor = &HFFFFFF&
    Left = 210
    Top = 2520
    Width = 1965
    Height = 315
    TabIndex = 8
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
  Begin VB.Label Label3
    Caption = "SMK Swasta RK Bintang Timur Pematang Siantar"
    ForeColor = &H4000&
    Left = 1290
    Top = 630
    Width = 3465
    Height = 675
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
  Begin VB.Shape Shape11
    Left = 13830
    Top = 2010
    Width = 225
    Height = 195
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Shape Shape10
    Left = 13440
    Top = 2010
    Width = 225
    Height = 195
    Shape = 3
    BorderStyle = 0 'Transparent
    FillColor = &H8080FF&
    FillStyle = 0
  End
  Begin VB.Shape Shape8
    Left = 1980
    Top = 2280
    Width = 12255
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Shape Shape7
    Left = 1950
    Top = 1890
    Width = 12285
    Height = 405
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Label Label2
    Caption = "NO. Daftar"
    ForeColor = &H80FF80&
    Left = 2250
    Top = 2730
    Width = 1815
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
    Left = -30
    Top = 2340
    Width = 1995
    Height = 4845
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Shape Shape4
    Left = 0
    Top = 8010
    Width = 14685
    Height = 1725
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Image Image2
    Picture = "daftarf.frx":0004252D
    Left = 12840
    Top = 630
    Width = 705
    Height = 675
    Stretch = -1  'True
  End
  Begin VB.Shape Shape12
    BorderColor = &HFFFFFF&
    Left = -60
    Top = 1530
    Width = 15015
    Height = 60
    BorderStyle = 3 'Dot
    FillColor = &HFFFFFF&
  End
  Begin VB.Label Label4
    Caption = "Register System"
    ForeColor = &H80FF80&
    Left = 6480
    Top = 60
    Width = 1815
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
  Begin VB.Shape Shape3
    Left = 0
    Top = 450
    Width = 14685
    Height = 75
    BorderStyle = 0 'Transparent
    FillColor = &H8000&
    FillStyle = 0
  End
  Begin VB.Shape Shape2
    Left = 0
    Top = 480
    Width = 14685
    Height = 1035
    BorderStyle = 0 'Transparent
    FillColor = &HFFFFFF&
    FillStyle = 0
  End
  Begin VB.Shape Shape1
    Left = 30
    Top = 30
    Width = 14685
    Height = 495
    BorderStyle = 0 'Transparent
    FillColor = &H4000&
    FillStyle = 0
  End
  Begin VB.Shape Shape16
    Left = 7230
    Top = 6930
    Width = 2505
    Height = 915
    BorderStyle = 0 'Transparent
    FillColor = &H80FF80&
    FillStyle = 0
  End
  Begin VB.Shape Shape5
    Left = 1950
    Top = 1980
    Width = 12255
    Height = 5865
    BorderStyle = 0 'Transparent
    FillColor = &H404040&
    FillStyle = 0
  End
  Begin VB.Image Image1
    Picture = "daftarf.frx":00048069
    Left = 30
    Top = 1410
    Width = 14625
    Height = 6855
    Stretch = -1  'True
  End
End

Attribute VB_Name = "daftarf"


Private Sub jcbutton2_UnknownEvent_9 'E096C0
  loc_00E096C0: push ebp
  loc_00E096C1: mov ebp, esp
  loc_00E096C3: sub esp, 0000000Ch
  loc_00E096C6: push 00402836h ; __vbaExceptHandler
  loc_00E096CB: mov eax, fs:[00000000h]
  loc_00E096D1: push eax
  loc_00E096D2: mov fs:[00000000h], esp
  loc_00E096D9: sub esp, 00000034h
  loc_00E096DC: push ebx
  loc_00E096DD: push esi
  loc_00E096DE: push edi
  loc_00E096DF: mov var_C, esp
  loc_00E096E2: mov var_8, 00402100h
  loc_00E096E9: mov esi, Me
  loc_00E096EC: mov eax, esi
  loc_00E096EE: and eax, 00000001h
  loc_00E096F1: mov var_4, eax
  loc_00E096F4: and esi, FFFFFFFEh
  loc_00E096F7: push esi
  loc_00E096F8: mov Me, esi
  loc_00E096FB: mov ecx, [esi]
  loc_00E096FD: call [ecx+00000004h]
  loc_00E09700: mov edx, [esi]
  loc_00E09702: push esi
  loc_00E09703: mov var_18, 00000000h
  loc_00E0970A: call [edx+0000037Ch]
  loc_00E09710: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E09716: push eax
  loc_00E09717: lea eax, var_18
  loc_00E0971A: push eax
  loc_00E0971B: call edi
  loc_00E0971D: mov ebx, eax
  loc_00E0971F: push FFFFFFFFh
  loc_00E09721: push ebx
  loc_00E09722: mov ecx, [ebx]
  loc_00E09724: call [ecx+0000009Ch]
  loc_00E0972A: test eax, eax
  loc_00E0972C: fnclex
  loc_00E0972E: jge 00E09742h
  loc_00E09730: push 0000009Ch
  loc_00E09735: push 006E03D4h
  loc_00E0973A: push ebx
  loc_00E0973B: push eax
  loc_00E0973C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09742: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E09748: lea ecx, var_18
  loc_00E0974B: call ebx
  loc_00E0974D: mov edx, [esi]
  loc_00E0974F: push esi
  loc_00E09750: call [edx+00000380h]
  loc_00E09756: push eax
  loc_00E09757: lea eax, var_18
  loc_00E0975A: push eax
  loc_00E0975B: call edi
  loc_00E0975D: mov edi, eax
  loc_00E0975F: push FFFFFFFFh
  loc_00E09761: push edi
  loc_00E09762: mov ecx, [edi]
  loc_00E09764: call [ecx+0000009Ch]
  loc_00E0976A: test eax, eax
  loc_00E0976C: fnclex
  loc_00E0976E: jge 00E09782h
  loc_00E09770: push 0000009Ch
  loc_00E09775: push 006E03D4h
  loc_00E0977A: push edi
  loc_00E0977B: push eax
  loc_00E0977C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09782: lea ecx, var_18
  loc_00E09785: call ebx
  loc_00E09787: mov edx, [esi]
  loc_00E09789: push esi
  loc_00E0978A: call [edx+00000380h]
  loc_00E09790: push eax
  loc_00E09791: lea eax, var_18
  loc_00E09794: push eax
  loc_00E09795: call [004010ACh] ; __vbaObjSet
  loc_00E0979B: mov edi, eax
  loc_00E0979D: push 00000000h
  loc_00E0979F: push edi
  loc_00E097A0: mov ecx, [edi]
  loc_00E097A2: call [ecx+000000E4h]
  loc_00E097A8: test eax, eax
  loc_00E097AA: fnclex
  loc_00E097AC: jge 00E097C0h
  loc_00E097AE: push 000000E4h
  loc_00E097B3: push 006E03D4h
  loc_00E097B8: push edi
  loc_00E097B9: push eax
  loc_00E097BA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E097C0: lea ecx, var_18
  loc_00E097C3: call ebx
  loc_00E097C5: mov edx, [esi]
  loc_00E097C7: push esi
  loc_00E097C8: call [edx+0000037Ch]
  loc_00E097CE: push eax
  loc_00E097CF: lea eax, var_18
  loc_00E097D2: push eax
  loc_00E097D3: call [004010ACh] ; __vbaObjSet
  loc_00E097D9: mov edi, eax
  loc_00E097DB: push 00000000h
  loc_00E097DD: push edi
  loc_00E097DE: mov ecx, [edi]
  loc_00E097E0: call [ecx+000000E4h]
  loc_00E097E6: test eax, eax
  loc_00E097E8: fnclex
  loc_00E097EA: jge 00E097FEh
  loc_00E097EC: push 000000E4h
  loc_00E097F1: push 006E03D4h
  loc_00E097F6: push edi
  loc_00E097F7: push eax
  loc_00E097F8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E097FE: lea ecx, var_18
  loc_00E09801: call ebx
  loc_00E09803: sub esp, 00000010h
  loc_00E09806: mov ecx, 0000000Bh
  loc_00E0980B: mov edx, esp
  loc_00E0980D: xor eax, eax
  loc_00E0980F: push 80010007h
  loc_00E09814: push esi
  loc_00E09815: mov [edx], ecx
  loc_00E09817: mov ecx, var_24
  loc_00E0981A: mov [edx+00000004h], ecx
  loc_00E0981D: mov ecx, [esi]
  loc_00E0981F: mov [edx+00000008h], eax
  loc_00E09822: mov eax, var_1C
  loc_00E09825: mov [edx+0000000Ch], eax
  loc_00E09828: call [ecx+000003B8h]
  loc_00E0982E: lea edx, var_18
  loc_00E09831: push eax
  loc_00E09832: push edx
  loc_00E09833: call [004010ACh] ; __vbaObjSet
  loc_00E09839: push eax
  loc_00E0983A: call [00401238h] ; __vbaLateIdSt
  loc_00E09840: lea ecx, var_18
  loc_00E09843: call ebx
  loc_00E09845: mov var_4, 00000000h
  loc_00E0984C: push 00E0985Eh
  loc_00E09851: jmp 00E0985Dh
  loc_00E09853: lea ecx, var_18
  loc_00E09856: call [00401254h] ; __vbaFreeObj
  loc_00E0985C: ret
  loc_00E0985D: ret
  loc_00E0985E: mov eax, Me
  loc_00E09861: push eax
  loc_00E09862: mov ecx, [eax]
  loc_00E09864: call [ecx+00000008h]
  loc_00E09867: mov eax, var_4
  loc_00E0986A: mov ecx, var_14
  loc_00E0986D: pop edi
  loc_00E0986E: pop esi
  loc_00E0986F: mov fs:[00000000h], ecx
  loc_00E09876: pop ebx
  loc_00E09877: mov esp, ebp
  loc_00E09879: pop ebp
  loc_00E0987A: retn 0004h
End Sub

Private Sub nomor_KeyPress(KeyAscii As Integer) 'E09A20
  loc_00E09A20: push ebp
  loc_00E09A21: mov ebp, esp
  loc_00E09A23: sub esp, 0000000Ch
  loc_00E09A26: push 00402836h ; __vbaExceptHandler
  loc_00E09A2B: mov eax, fs:[00000000h]
  loc_00E09A31: push eax
  loc_00E09A32: mov fs:[00000000h], esp
  loc_00E09A39: sub esp, 00000014h
  loc_00E09A3C: push ebx
  loc_00E09A3D: push esi
  loc_00E09A3E: push edi
  loc_00E09A3F: mov var_C, esp
  loc_00E09A42: mov var_8, 00402130h
  loc_00E09A49: mov esi, Me
  loc_00E09A4C: mov eax, esi
  loc_00E09A4E: and eax, 00000001h
  loc_00E09A51: mov var_4, eax
  loc_00E09A54: and esi, FFFFFFFEh
  loc_00E09A57: push esi
  loc_00E09A58: mov Me, esi
  loc_00E09A5B: mov ecx, [esi]
  loc_00E09A5D: call [ecx+00000004h]
  loc_00E09A60: mov edx, KeyAscii
  loc_00E09A63: xor edi, edi
  loc_00E09A65: mov var_18, edi
  loc_00E09A68: cmp [edx], 000Dh
  loc_00E09A6C: jnz 00E09AAEh
  loc_00E09A6E: mov eax, [esi]
  loc_00E09A70: push esi
  loc_00E09A71: call [eax+00000394h]
  loc_00E09A77: lea ecx, var_18
  loc_00E09A7A: push eax
  loc_00E09A7B: push ecx
  loc_00E09A7C: call [004010ACh] ; __vbaObjSet
  loc_00E09A82: mov esi, eax
  loc_00E09A84: push esi
  loc_00E09A85: mov edx, [esi]
  loc_00E09A87: call [edx+00000204h]
  loc_00E09A8D: cmp eax, edi
  loc_00E09A8F: fnclex
  loc_00E09A91: jge 00E09AA5h
  loc_00E09A93: push 00000204h
  loc_00E09A98: push 006DCB70h
  loc_00E09A9D: push esi
  loc_00E09A9E: push eax
  loc_00E09A9F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09AA5: lea ecx, var_18
  loc_00E09AA8: call [00401254h] ; __vbaFreeObj
  loc_00E09AAE: mov var_4, edi
  loc_00E09AB1: push 00E09AC3h
  loc_00E09AB6: jmp 00E09AC2h
  loc_00E09AB8: lea ecx, var_18
  loc_00E09ABB: call [00401254h] ; __vbaFreeObj
  loc_00E09AC1: ret
  loc_00E09AC2: ret
  loc_00E09AC3: mov eax, Me
  loc_00E09AC6: push eax
  loc_00E09AC7: mov ecx, [eax]
  loc_00E09AC9: call [ecx+00000008h]
  loc_00E09ACC: mov eax, var_4
  loc_00E09ACF: mov ecx, var_14
  loc_00E09AD2: pop edi
  loc_00E09AD3: pop esi
  loc_00E09AD4: mov fs:[00000000h], ecx
  loc_00E09ADB: pop ebx
  loc_00E09ADC: mov esp, ebp
  loc_00E09ADE: pop ebp
  loc_00E09ADF: retn 0008h
End Sub

Private Sub nama_KeyPress(KeyAscii As Integer) 'E09AF0
  loc_00E09AF0: push ebp
  loc_00E09AF1: mov ebp, esp
  loc_00E09AF3: sub esp, 0000000Ch
  loc_00E09AF6: push 00402836h ; __vbaExceptHandler
  loc_00E09AFB: mov eax, fs:[00000000h]
  loc_00E09B01: push eax
  loc_00E09B02: mov fs:[00000000h], esp
  loc_00E09B09: sub esp, 00000014h
  loc_00E09B0C: push ebx
  loc_00E09B0D: push esi
  loc_00E09B0E: push edi
  loc_00E09B0F: mov var_C, esp
  loc_00E09B12: mov var_8, 00402140h
  loc_00E09B19: mov esi, Me
  loc_00E09B1C: mov eax, esi
  loc_00E09B1E: and eax, 00000001h
  loc_00E09B21: mov var_4, eax
  loc_00E09B24: and esi, FFFFFFFEh
  loc_00E09B27: push esi
  loc_00E09B28: mov Me, esi
  loc_00E09B2B: mov ecx, [esi]
  loc_00E09B2D: call [ecx+00000004h]
  loc_00E09B30: mov edx, KeyAscii
  loc_00E09B33: xor edi, edi
  loc_00E09B35: mov var_18, edi
  loc_00E09B38: cmp [edx], 000Dh
  loc_00E09B3C: jnz 00E09B7Eh
  loc_00E09B3E: mov eax, [esi]
  loc_00E09B40: push esi
  loc_00E09B41: call [eax+00000390h]
  loc_00E09B47: lea ecx, var_18
  loc_00E09B4A: push eax
  loc_00E09B4B: push ecx
  loc_00E09B4C: call [004010ACh] ; __vbaObjSet
  loc_00E09B52: mov esi, eax
  loc_00E09B54: push esi
  loc_00E09B55: mov edx, [esi]
  loc_00E09B57: call [edx+00000204h]
  loc_00E09B5D: cmp eax, edi
  loc_00E09B5F: fnclex
  loc_00E09B61: jge 00E09B75h
  loc_00E09B63: push 00000204h
  loc_00E09B68: push 006DCB70h
  loc_00E09B6D: push esi
  loc_00E09B6E: push eax
  loc_00E09B6F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09B75: lea ecx, var_18
  loc_00E09B78: call [00401254h] ; __vbaFreeObj
  loc_00E09B7E: mov var_4, edi
  loc_00E09B81: push 00E09B93h
  loc_00E09B86: jmp 00E09B92h
  loc_00E09B88: lea ecx, var_18
  loc_00E09B8B: call [00401254h] ; __vbaFreeObj
  loc_00E09B91: ret
  loc_00E09B92: ret
  loc_00E09B93: mov eax, Me
  loc_00E09B96: push eax
  loc_00E09B97: mov ecx, [eax]
  loc_00E09B99: call [ecx+00000008h]
  loc_00E09B9C: mov eax, var_4
  loc_00E09B9F: mov ecx, var_14
  loc_00E09BA2: pop edi
  loc_00E09BA3: pop esi
  loc_00E09BA4: mov fs:[00000000h], ecx
  loc_00E09BAB: pop ebx
  loc_00E09BAC: mov esp, ebp
  loc_00E09BAE: pop ebp
  loc_00E09BAF: retn 0008h
End Sub

Private Sub asalsekolah_KeyPress(KeyAscii As Integer) 'E09BC0
  loc_00E09BC0: push ebp
  loc_00E09BC1: mov ebp, esp
  loc_00E09BC3: sub esp, 0000000Ch
  loc_00E09BC6: push 00402836h ; __vbaExceptHandler
  loc_00E09BCB: mov eax, fs:[00000000h]
  loc_00E09BD1: push eax
  loc_00E09BD2: mov fs:[00000000h], esp
  loc_00E09BD9: sub esp, 00000014h
  loc_00E09BDC: push ebx
  loc_00E09BDD: push esi
  loc_00E09BDE: push edi
  loc_00E09BDF: mov var_C, esp
  loc_00E09BE2: mov var_8, 00402150h
  loc_00E09BE9: mov esi, Me
  loc_00E09BEC: mov eax, esi
  loc_00E09BEE: and eax, 00000001h
  loc_00E09BF1: mov var_4, eax
  loc_00E09BF4: and esi, FFFFFFFEh
  loc_00E09BF7: push esi
  loc_00E09BF8: mov Me, esi
  loc_00E09BFB: mov ecx, [esi]
  loc_00E09BFD: call [ecx+00000004h]
  loc_00E09C00: mov edx, KeyAscii
  loc_00E09C03: xor edi, edi
  loc_00E09C05: mov var_18, edi
  loc_00E09C08: cmp [edx], 000Dh
  loc_00E09C0C: jnz 00E09C4Eh
  loc_00E09C0E: mov eax, [esi]
  loc_00E09C10: push esi
  loc_00E09C11: call [eax+0000038Ch]
  loc_00E09C17: lea ecx, var_18
  loc_00E09C1A: push eax
  loc_00E09C1B: push ecx
  loc_00E09C1C: call [004010ACh] ; __vbaObjSet
  loc_00E09C22: mov esi, eax
  loc_00E09C24: push esi
  loc_00E09C25: mov edx, [esi]
  loc_00E09C27: call [edx+00000204h]
  loc_00E09C2D: cmp eax, edi
  loc_00E09C2F: fnclex
  loc_00E09C31: jge 00E09C45h
  loc_00E09C33: push 00000204h
  loc_00E09C38: push 006DCB70h
  loc_00E09C3D: push esi
  loc_00E09C3E: push eax
  loc_00E09C3F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09C45: lea ecx, var_18
  loc_00E09C48: call [00401254h] ; __vbaFreeObj
  loc_00E09C4E: mov var_4, edi
  loc_00E09C51: push 00E09C63h
  loc_00E09C56: jmp 00E09C62h
  loc_00E09C58: lea ecx, var_18
  loc_00E09C5B: call [00401254h] ; __vbaFreeObj
  loc_00E09C61: ret
  loc_00E09C62: ret
  loc_00E09C63: mov eax, Me
  loc_00E09C66: push eax
  loc_00E09C67: mov ecx, [eax]
  loc_00E09C69: call [ecx+00000008h]
  loc_00E09C6C: mov eax, var_4
  loc_00E09C6F: mov ecx, var_14
  loc_00E09C72: pop edi
  loc_00E09C73: pop esi
  loc_00E09C74: mov fs:[00000000h], ecx
  loc_00E09C7B: pop ebx
  loc_00E09C7C: mov esp, ebp
  loc_00E09C7E: pop ebp
  loc_00E09C7F: retn 0008h
End Sub

Private Sub update_UnknownEvent_9 'E15A20
  loc_00E15A20: push ebp
  loc_00E15A21: mov ebp, esp
  loc_00E15A23: sub esp, 0000000Ch
  loc_00E15A26: push 00402836h ; __vbaExceptHandler
  loc_00E15A2B: mov eax, fs:[00000000h]
  loc_00E15A31: push eax
  loc_00E15A32: mov fs:[00000000h], esp
  loc_00E15A39: sub esp, 000000CCh
  loc_00E15A3F: push ebx
  loc_00E15A40: push esi
  loc_00E15A41: push edi
  loc_00E15A42: mov var_C, esp
  loc_00E15A45: mov var_8, 004022D8h
  loc_00E15A4C: mov esi, Me
  loc_00E15A4F: mov eax, esi
  loc_00E15A51: and eax, 00000001h
  loc_00E15A54: mov var_4, eax
  loc_00E15A57: and esi, FFFFFFFEh
  loc_00E15A5A: push esi
  loc_00E15A5B: mov Me, esi
  loc_00E15A5E: mov ecx, [esi]
  loc_00E15A60: call [ecx+00000004h]
  loc_00E15A63: mov edx, [esi]
  loc_00E15A65: xor ebx, ebx
  loc_00E15A67: push esi
  loc_00E15A68: mov var_18, ebx
  loc_00E15A6B: mov var_1C, ebx
  loc_00E15A6E: mov var_20, ebx
  loc_00E15A71: mov var_24, ebx
  loc_00E15A74: mov var_28, ebx
  loc_00E15A77: mov var_2C, ebx
  loc_00E15A7A: mov var_30, ebx
  loc_00E15A7D: mov var_34, ebx
  loc_00E15A80: mov var_44, ebx
  loc_00E15A83: mov var_54, ebx
  loc_00E15A86: mov var_78, ebx
  loc_00E15A89: call [edx+0000039Ch]
  loc_00E15A8F: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E15A95: push eax
  loc_00E15A96: lea eax, var_24
  loc_00E15A99: push eax
  loc_00E15A9A: call edi
  loc_00E15A9C: mov ecx, [eax]
  loc_00E15A9E: lea edx, var_18
  loc_00E15AA1: push edx
  loc_00E15AA2: push eax
  loc_00E15AA3: mov var_7C, eax
  loc_00E15AA6: call [ecx+000000A0h]
  loc_00E15AAC: cmp eax, ebx
  loc_00E15AAE: fnclex
  loc_00E15AB0: jge 00E15AC7h
  loc_00E15AB2: mov ecx, var_7C
  loc_00E15AB5: push 000000A0h
  loc_00E15ABA: push 006DCB70h
  loc_00E15ABF: push ecx
  loc_00E15AC0: push eax
  loc_00E15AC1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15AC7: mov edx, var_18
  loc_00E15ACA: push 006E0994h ; "select * from biodata where no_daftar ='"
  loc_00E15ACF: push edx
  loc_00E15AD0: lea ebx, [esi+00000034h]
  loc_00E15AD3: call [0040106Ch] ; __vbaStrCat
  loc_00E15AD9: mov edx, eax
  loc_00E15ADB: lea ecx, var_1C
  loc_00E15ADE: call [00401228h] ; __vbaStrMove
  loc_00E15AE4: push eax
  loc_00E15AE5: push 006DCB84h ; "'"
  loc_00E15AEA: call [0040106Ch] ; __vbaStrCat
  loc_00E15AF0: mov edx, eax
  loc_00E15AF2: lea ecx, var_20
  loc_00E15AF5: call [00401228h] ; __vbaStrMove
  loc_00E15AFB: mov edx, eax
  loc_00E15AFD: mov ecx, ebx
  loc_00E15AFF: call [004011B0h] ; __vbaStrCopy
  loc_00E15B05: lea eax, var_20
  loc_00E15B08: lea ecx, var_1C
  loc_00E15B0B: push eax
  loc_00E15B0C: lea edx, var_18
  loc_00E15B0F: push ecx
  loc_00E15B10: push edx
  loc_00E15B11: push 00000003h
  loc_00E15B13: call [004011BCh] ; __vbaFreeStrList
  loc_00E15B19: add esp, 00000010h
  loc_00E15B1C: lea ecx, var_24
  loc_00E15B1F: call [00401254h] ; __vbaFreeObj
  loc_00E15B25: mov edx, var_60
  loc_00E15B28: sub esp, 00000010h
  loc_00E15B2B: mov ecx, esp
  loc_00E15B2D: mov eax, 00004008h
  loc_00E15B32: push 0000000Eh
  loc_00E15B34: push esi
  loc_00E15B35: mov [ecx], eax
  loc_00E15B37: mov eax, var_58
  loc_00E15B3A: mov [ecx+00000004h], edx
  loc_00E15B3D: mov [ecx+00000008h], ebx
  loc_00E15B40: mov [ecx+0000000Ch], eax
  loc_00E15B43: mov ecx, [esi]
  loc_00E15B45: call [ecx+00000490h]
  loc_00E15B4B: lea edx, var_24
  loc_00E15B4E: push eax
  loc_00E15B4F: push edx
  loc_00E15B50: call edi
  loc_00E15B52: push eax
  loc_00E15B53: call [00401238h] ; __vbaLateIdSt
  loc_00E15B59: lea ecx, var_24
  loc_00E15B5C: call [00401254h] ; __vbaFreeObj
  loc_00E15B62: mov eax, [esi]
  loc_00E15B64: push 006DCBD8h
  loc_00E15B69: push 00000000h
  loc_00E15B6B: push 00000018h
  loc_00E15B6D: push esi
  loc_00E15B6E: call [eax+00000490h]
  loc_00E15B74: lea ecx, var_24
  loc_00E15B77: push eax
  loc_00E15B78: push ecx
  loc_00E15B79: call edi
  loc_00E15B7B: lea edx, var_44
  loc_00E15B7E: push eax
  loc_00E15B7F: push edx
  loc_00E15B80: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E15B86: add esp, 00000010h
  loc_00E15B89: push eax
  loc_00E15B8A: call [00401120h] ; __vbaCastObjVar
  loc_00E15B90: push eax
  loc_00E15B91: lea eax, var_28
  loc_00E15B94: push eax
  loc_00E15B95: call edi
  loc_00E15B97: sub esp, 00000010h
  loc_00E15B9A: mov ebx, eax
  loc_00E15B9C: mov edx, esp
  loc_00E15B9E: mov eax, 0000000Ah
  loc_00E15BA3: mov ecx, [ebx]
  loc_00E15BA5: mov var_64, eax
  loc_00E15BA8: mov [edx], eax
  loc_00E15BAA: mov eax, var_70
  loc_00E15BAD: mov var_5C, 80020004h
  loc_00E15BB4: mov [edx+00000004h], eax
  loc_00E15BB7: mov eax, 80020004h
  loc_00E15BBC: mov [edx+00000008h], eax
  loc_00E15BBF: mov eax, var_68
  loc_00E15BC2: sub esp, 00000010h
  loc_00E15BC5: mov [edx+0000000Ch], eax
  loc_00E15BC8: mov eax, var_64
  loc_00E15BCB: mov edx, esp
  loc_00E15BCD: push ebx
  loc_00E15BCE: mov [edx], eax
  loc_00E15BD0: mov eax, var_60
  loc_00E15BD3: mov [edx+00000004h], eax
  loc_00E15BD6: mov eax, var_5C
  loc_00E15BD9: mov [edx+00000008h], eax
  loc_00E15BDC: mov eax, var_58
  loc_00E15BDF: mov [edx+0000000Ch], eax
  loc_00E15BE2: call [ecx+000000ACh]
  loc_00E15BE8: test eax, eax
  loc_00E15BEA: fnclex
  loc_00E15BEC: jge 00E15C00h
  loc_00E15BEE: push 000000ACh
  loc_00E15BF3: push 006DCBE8h
  loc_00E15BF8: push ebx
  loc_00E15BF9: push eax
  loc_00E15BFA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15C00: lea ecx, var_28
  loc_00E15C03: lea edx, var_24
  loc_00E15C06: push ecx
  loc_00E15C07: push edx
  loc_00E15C08: push 00000002h
  loc_00E15C0A: call [00401048h] ; __vbaFreeObjList
  loc_00E15C10: add esp, 0000000Ch
  loc_00E15C13: lea ecx, var_44
  loc_00E15C16: call [00401024h] ; __vbaFreeVar
  loc_00E15C1C: mov eax, [esi]
  loc_00E15C1E: push 006DCBD8h
  loc_00E15C23: push 00000000h
  loc_00E15C25: push 00000018h
  loc_00E15C27: push esi
  loc_00E15C28: call [eax+00000490h]
  loc_00E15C2E: lea ecx, var_28
  loc_00E15C31: push eax
  loc_00E15C32: push ecx
  loc_00E15C33: call edi
  loc_00E15C35: lea edx, var_44
  loc_00E15C38: push eax
  loc_00E15C39: push edx
  loc_00E15C3A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E15C40: add esp, 00000010h
  loc_00E15C43: push eax
  loc_00E15C44: call [00401120h] ; __vbaCastObjVar
  loc_00E15C4A: push eax
  loc_00E15C4B: lea eax, var_2C
  loc_00E15C4E: push eax
  loc_00E15C4F: call edi
  loc_00E15C51: mov ebx, eax
  loc_00E15C53: lea edx, var_30
  loc_00E15C56: push edx
  loc_00E15C57: push ebx
  loc_00E15C58: mov ecx, [ebx]
  loc_00E15C5A: call [ecx+00000054h]
  loc_00E15C5D: test eax, eax
  loc_00E15C5F: fnclex
  loc_00E15C61: jge 00E15C72h
  loc_00E15C63: push 00000054h
  loc_00E15C65: push 006DCBE8h
  loc_00E15C6A: push ebx
  loc_00E15C6B: push eax
  loc_00E15C6C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15C72: lea ebx, var_34
  loc_00E15C75: mov eax, var_30
  loc_00E15C78: push ebx
  loc_00E15C79: mov ecx, 00000002h
  loc_00E15C7E: sub esp, 00000010h
  loc_00E15C81: mov edx, [eax]
  loc_00E15C83: mov ebx, esp
  loc_00E15C85: mov var_5C, 00000000h
  loc_00E15C8C: push eax
  loc_00E15C8D: mov var_8C, eax
  loc_00E15C93: mov [ebx], ecx
  loc_00E15C95: mov ecx, var_60
  loc_00E15C98: mov [ebx+00000004h], ecx
  loc_00E15C9B: mov ecx, var_5C
  loc_00E15C9E: mov [ebx+00000008h], ecx
  loc_00E15CA1: mov ecx, var_58
  loc_00E15CA4: mov [ebx+0000000Ch], ecx
  loc_00E15CA7: call [edx+00000028h]
  loc_00E15CAA: test eax, eax
  loc_00E15CAC: fnclex
  loc_00E15CAE: jge 00E15CC5h
  loc_00E15CB0: mov edx, var_8C
  loc_00E15CB6: push 00000028h
  loc_00E15CB8: push 006E09E8h
  loc_00E15CBD: push edx
  loc_00E15CBE: push eax
  loc_00E15CBF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15CC5: mov eax, var_34
  loc_00E15CC8: mov ecx, [esi]
  loc_00E15CCA: push esi
  loc_00E15CCB: mov var_94, eax
  loc_00E15CD1: call [ecx+0000039Ch]
  loc_00E15CD7: lea edx, var_24
  loc_00E15CDA: push eax
  loc_00E15CDB: push edx
  loc_00E15CDC: call edi
  loc_00E15CDE: mov ebx, eax
  loc_00E15CE0: lea ecx, var_18
  loc_00E15CE3: push ecx
  loc_00E15CE4: push ebx
  loc_00E15CE5: mov eax, [ebx]
  loc_00E15CE7: call [eax+000000A0h]
  loc_00E15CED: test eax, eax
  loc_00E15CEF: fnclex
  loc_00E15CF1: jge 00E15D05h
  loc_00E15CF3: push 000000A0h
  loc_00E15CF8: push 006DCB70h
  loc_00E15CFD: push ebx
  loc_00E15CFE: push eax
  loc_00E15CFF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15D05: sub esp, 00000010h
  loc_00E15D08: mov eax, var_18
  loc_00E15D0B: mov ebx, esp
  loc_00E15D0D: mov ecx, 00000008h
  loc_00E15D12: mov var_54, ecx
  loc_00E15D15: mov edx, var_94
  loc_00E15D1B: mov [ebx], ecx
  loc_00E15D1D: mov ecx, var_50
  loc_00E15D20: mov var_4C, eax
  loc_00E15D23: mov var_18, 00000000h
  loc_00E15D2A: mov edx, [edx]
  loc_00E15D2C: mov [ebx+00000004h], ecx
  loc_00E15D2F: mov [ebx+00000008h], eax
  loc_00E15D32: mov eax, var_48
  loc_00E15D35: mov [ebx+0000000Ch], eax
  loc_00E15D38: mov ebx, var_94
  loc_00E15D3E: push ebx
  loc_00E15D3F: call [edx+00000038h]
  loc_00E15D42: test eax, eax
  loc_00E15D44: fnclex
  loc_00E15D46: jge 00E15D57h
  loc_00E15D48: push 00000038h
  loc_00E15D4A: push 006E09F8h
  loc_00E15D4F: push ebx
  loc_00E15D50: push eax
  loc_00E15D51: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15D57: lea ecx, var_34
  loc_00E15D5A: lea edx, var_30
  loc_00E15D5D: push ecx
  loc_00E15D5E: lea eax, var_2C
  loc_00E15D61: push edx
  loc_00E15D62: lea ecx, var_28
  loc_00E15D65: push eax
  loc_00E15D66: lea edx, var_24
  loc_00E15D69: push ecx
  loc_00E15D6A: push edx
  loc_00E15D6B: push 00000005h
  loc_00E15D6D: call [00401048h] ; __vbaFreeObjList
  loc_00E15D73: lea eax, var_54
  loc_00E15D76: lea ecx, var_44
  loc_00E15D79: push eax
  loc_00E15D7A: push ecx
  loc_00E15D7B: push 00000002h
  loc_00E15D7D: call [00401038h] ; __vbaFreeVarList
  loc_00E15D83: mov edx, [esi]
  loc_00E15D85: add esp, 00000024h
  loc_00E15D88: push 006DCBD8h
  loc_00E15D8D: push 00000000h
  loc_00E15D8F: push 00000018h
  loc_00E15D91: push esi
  loc_00E15D92: call [edx+00000490h]
  loc_00E15D98: push eax
  loc_00E15D99: lea eax, var_28
  loc_00E15D9C: push eax
  loc_00E15D9D: call edi
  loc_00E15D9F: lea ecx, var_44
  loc_00E15DA2: push eax
  loc_00E15DA3: push ecx
  loc_00E15DA4: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E15DAA: add esp, 00000010h
  loc_00E15DAD: push eax
  loc_00E15DAE: call [00401120h] ; __vbaCastObjVar
  loc_00E15DB4: lea edx, var_2C
  loc_00E15DB7: push eax
  loc_00E15DB8: push edx
  loc_00E15DB9: call edi
  loc_00E15DBB: mov ebx, eax
  loc_00E15DBD: lea ecx, var_30
  loc_00E15DC0: push ecx
  loc_00E15DC1: push ebx
  loc_00E15DC2: mov eax, [ebx]
  loc_00E15DC4: call [eax+00000054h]
  loc_00E15DC7: test eax, eax
  loc_00E15DC9: fnclex
  loc_00E15DCB: jge 00E15DDCh
  loc_00E15DCD: push 00000054h
  loc_00E15DCF: push 006DCBE8h
  loc_00E15DD4: push ebx
  loc_00E15DD5: push eax
  loc_00E15DD6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15DDC: lea ebx, var_34
  loc_00E15DDF: mov eax, var_30
  loc_00E15DE2: push ebx
  loc_00E15DE3: mov ecx, 00000002h
  loc_00E15DE8: sub esp, 00000010h
  loc_00E15DEB: mov edx, [eax]
  loc_00E15DED: mov ebx, esp
  loc_00E15DEF: mov var_8C, eax
  loc_00E15DF5: push eax
  loc_00E15DF6: mov [ebx], ecx
  loc_00E15DF8: mov ecx, var_60
  loc_00E15DFB: mov [ebx+00000004h], ecx
  loc_00E15DFE: mov ecx, 00000001h
  loc_00E15E03: mov [ebx+00000008h], ecx
  loc_00E15E06: mov ecx, var_58
  loc_00E15E09: mov [ebx+0000000Ch], ecx
  loc_00E15E0C: call [edx+00000028h]
  loc_00E15E0F: test eax, eax
  loc_00E15E11: fnclex
  loc_00E15E13: jge 00E15E2Ah
  loc_00E15E15: mov edx, var_8C
  loc_00E15E1B: push 00000028h
  loc_00E15E1D: push 006E09E8h
  loc_00E15E22: push edx
  loc_00E15E23: push eax
  loc_00E15E24: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15E2A: mov eax, var_34
  loc_00E15E2D: mov ecx, [esi]
  loc_00E15E2F: push esi
  loc_00E15E30: mov var_94, eax
  loc_00E15E36: call [ecx+00000394h]
  loc_00E15E3C: lea edx, var_24
  loc_00E15E3F: push eax
  loc_00E15E40: push edx
  loc_00E15E41: call edi
  loc_00E15E43: mov ebx, eax
  loc_00E15E45: lea ecx, var_18
  loc_00E15E48: push ecx
  loc_00E15E49: push ebx
  loc_00E15E4A: mov eax, [ebx]
  loc_00E15E4C: call [eax+000000A0h]
  loc_00E15E52: test eax, eax
  loc_00E15E54: fnclex
  loc_00E15E56: jge 00E15E6Ah
  loc_00E15E58: push 000000A0h
  loc_00E15E5D: push 006DCB70h
  loc_00E15E62: push ebx
  loc_00E15E63: push eax
  loc_00E15E64: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15E6A: sub esp, 00000010h
  loc_00E15E6D: mov eax, var_18
  loc_00E15E70: mov ebx, esp
  loc_00E15E72: mov ecx, 00000008h
  loc_00E15E77: mov var_54, ecx
  loc_00E15E7A: mov edx, var_94
  loc_00E15E80: mov [ebx], ecx
  loc_00E15E82: mov ecx, var_50
  loc_00E15E85: mov var_4C, eax
  loc_00E15E88: mov var_18, 00000000h
  loc_00E15E8F: mov edx, [edx]
  loc_00E15E91: mov [ebx+00000004h], ecx
  loc_00E15E94: mov [ebx+00000008h], eax
  loc_00E15E97: mov eax, var_48
  loc_00E15E9A: mov [ebx+0000000Ch], eax
  loc_00E15E9D: mov ebx, var_94
  loc_00E15EA3: push ebx
  loc_00E15EA4: call [edx+00000038h]
  loc_00E15EA7: test eax, eax
  loc_00E15EA9: fnclex
  loc_00E15EAB: jge 00E15EBCh
  loc_00E15EAD: push 00000038h
  loc_00E15EAF: push 006E09F8h
  loc_00E15EB4: push ebx
  loc_00E15EB5: push eax
  loc_00E15EB6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15EBC: lea ecx, var_34
  loc_00E15EBF: lea edx, var_30
  loc_00E15EC2: push ecx
  loc_00E15EC3: lea eax, var_2C
  loc_00E15EC6: push edx
  loc_00E15EC7: lea ecx, var_28
  loc_00E15ECA: push eax
  loc_00E15ECB: lea edx, var_24
  loc_00E15ECE: push ecx
  loc_00E15ECF: push edx
  loc_00E15ED0: push 00000005h
  loc_00E15ED2: call [00401048h] ; __vbaFreeObjList
  loc_00E15ED8: lea eax, var_54
  loc_00E15EDB: lea ecx, var_44
  loc_00E15EDE: push eax
  loc_00E15EDF: push ecx
  loc_00E15EE0: push 00000002h
  loc_00E15EE2: call [00401038h] ; __vbaFreeVarList
  loc_00E15EE8: mov edx, [esi]
  loc_00E15EEA: add esp, 00000024h
  loc_00E15EED: push esi
  loc_00E15EEE: call [edx+00000380h]
  loc_00E15EF4: push eax
  loc_00E15EF5: lea eax, var_24
  loc_00E15EF8: push eax
  loc_00E15EF9: call edi
  loc_00E15EFB: mov ebx, eax
  loc_00E15EFD: lea edx, var_78
  loc_00E15F00: push edx
  loc_00E15F01: push ebx
  loc_00E15F02: mov ecx, [ebx]
  loc_00E15F04: call [ecx+000000E0h]
  loc_00E15F0A: test eax, eax
  loc_00E15F0C: fnclex
  loc_00E15F0E: jge 00E15F22h
  loc_00E15F10: push 000000E0h
  loc_00E15F15: push 006E03D4h
  loc_00E15F1A: push ebx
  loc_00E15F1B: push eax
  loc_00E15F1C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15F22: mov bx, var_78
  loc_00E15F26: lea ecx, var_24
  loc_00E15F29: call [00401254h] ; __vbaFreeObj
  loc_00E15F2F: push 006DCBD8h
  loc_00E15F34: push 00000000h
  loc_00E15F36: test bx, bx
  loc_00E15F39: push 00000018h
  loc_00E15F3B: jz 00E1601Ah
  loc_00E15F41: mov eax, [esi]
  loc_00E15F43: push esi
  loc_00E15F44: call [eax+00000490h]
  loc_00E15F4A: lea ecx, var_24
  loc_00E15F4D: push eax
  loc_00E15F4E: push ecx
  loc_00E15F4F: call edi
  loc_00E15F51: lea edx, var_44
  loc_00E15F54: push eax
  loc_00E15F55: push edx
  loc_00E15F56: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E15F5C: add esp, 00000010h
  loc_00E15F5F: push eax
  loc_00E15F60: call [00401120h] ; __vbaCastObjVar
  loc_00E15F66: push eax
  loc_00E15F67: lea eax, var_28
  loc_00E15F6A: push eax
  loc_00E15F6B: call edi
  loc_00E15F6D: mov ebx, eax
  loc_00E15F6F: lea edx, var_2C
  loc_00E15F72: push edx
  loc_00E15F73: push ebx
  loc_00E15F74: mov ecx, [ebx]
  loc_00E15F76: call [ecx+00000054h]
  loc_00E15F79: test eax, eax
  loc_00E15F7B: fnclex
  loc_00E15F7D: jge 00E15F8Eh
  loc_00E15F7F: push 00000054h
  loc_00E15F81: push 006DCBE8h
  loc_00E15F86: push ebx
  loc_00E15F87: push eax
  loc_00E15F88: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15F8E: lea ebx, var_30
  loc_00E15F91: mov eax, var_2C
  loc_00E15F94: push ebx
  loc_00E15F95: mov ecx, 00000002h
  loc_00E15F9A: sub esp, 00000010h
  loc_00E15F9D: mov edx, [eax]
  loc_00E15F9F: mov ebx, esp
  loc_00E15FA1: mov var_84, eax
  loc_00E15FA7: push eax
  loc_00E15FA8: mov [ebx], ecx
  loc_00E15FAA: mov ecx, var_60
  loc_00E15FAD: mov [ebx+00000004h], ecx
  loc_00E15FB0: mov ecx, 00000002h
  loc_00E15FB5: mov [ebx+00000008h], ecx
  loc_00E15FB8: mov ecx, var_58
  loc_00E15FBB: mov [ebx+0000000Ch], ecx
  loc_00E15FBE: call [edx+00000028h]
  loc_00E15FC1: test eax, eax
  loc_00E15FC3: fnclex
  loc_00E15FC5: jge 00E15FDCh
  loc_00E15FC7: mov edx, var_84
  loc_00E15FCD: push 00000028h
  loc_00E15FCF: push 006E09E8h
  loc_00E15FD4: push edx
  loc_00E15FD5: push eax
  loc_00E15FD6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15FDC: sub esp, 00000010h
  loc_00E15FDF: mov eax, var_30
  loc_00E15FE2: mov ebx, esp
  loc_00E15FE4: mov ecx, 00000008h
  loc_00E15FE9: mov edx, [eax]
  loc_00E15FEB: push eax
  loc_00E15FEC: mov [ebx], ecx
  loc_00E15FEE: mov ecx, var_70
  loc_00E15FF1: mov var_8C, eax
  loc_00E15FF7: mov [ebx+00000004h], ecx
  loc_00E15FFA: mov ecx, 006E0B24h ; "Laki - laki"
  loc_00E15FFF: mov [ebx+00000008h], ecx
  loc_00E16002: mov ecx, var_68
  loc_00E16005: mov [ebx+0000000Ch], ecx
  loc_00E16008: call [edx+00000038h]
  loc_00E1600B: test eax, eax
  loc_00E1600D: fnclex
  loc_00E1600F: jge 00E160FFh
  loc_00E16015: jmp 00E160EAh
  loc_00E1601A: mov ecx, [esi]
  loc_00E1601C: push esi
  loc_00E1601D: call [ecx+00000490h]
  loc_00E16023: lea edx, var_24
  loc_00E16026: push eax
  loc_00E16027: push edx
  loc_00E16028: call edi
  loc_00E1602A: push eax
  loc_00E1602B: lea eax, var_44
  loc_00E1602E: push eax
  loc_00E1602F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16035: add esp, 00000010h
  loc_00E16038: push eax
  loc_00E16039: call [00401120h] ; __vbaCastObjVar
  loc_00E1603F: lea ecx, var_28
  loc_00E16042: push eax
  loc_00E16043: push ecx
  loc_00E16044: call edi
  loc_00E16046: mov ebx, eax
  loc_00E16048: lea eax, var_2C
  loc_00E1604B: push eax
  loc_00E1604C: push ebx
  loc_00E1604D: mov edx, [ebx]
  loc_00E1604F: call [edx+00000054h]
  loc_00E16052: test eax, eax
  loc_00E16054: fnclex
  loc_00E16056: jge 00E16067h
  loc_00E16058: push 00000054h
  loc_00E1605A: push 006DCBE8h
  loc_00E1605F: push ebx
  loc_00E16060: push eax
  loc_00E16061: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16067: lea ebx, var_30
  loc_00E1606A: mov eax, var_2C
  loc_00E1606D: push ebx
  loc_00E1606E: mov ecx, 00000002h
  loc_00E16073: sub esp, 00000010h
  loc_00E16076: mov edx, [eax]
  loc_00E16078: mov ebx, esp
  loc_00E1607A: mov var_84, eax
  loc_00E16080: push eax
  loc_00E16081: mov [ebx], ecx
  loc_00E16083: mov ecx, var_60
  loc_00E16086: mov [ebx+00000004h], ecx
  loc_00E16089: mov ecx, 00000002h
  loc_00E1608E: mov [ebx+00000008h], ecx
  loc_00E16091: mov ecx, var_58
  loc_00E16094: mov [ebx+0000000Ch], ecx
  loc_00E16097: call [edx+00000028h]
  loc_00E1609A: test eax, eax
  loc_00E1609C: fnclex
  loc_00E1609E: jge 00E160B5h
  loc_00E160A0: mov edx, var_84
  loc_00E160A6: push 00000028h
  loc_00E160A8: push 006E09E8h
  loc_00E160AD: push edx
  loc_00E160AE: push eax
  loc_00E160AF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E160B5: sub esp, 00000010h
  loc_00E160B8: mov eax, var_30
  loc_00E160BB: mov ebx, esp
  loc_00E160BD: mov ecx, 00000008h
  loc_00E160C2: mov edx, [eax]
  loc_00E160C4: push eax
  loc_00E160C5: mov [ebx], ecx
  loc_00E160C7: mov ecx, var_70
  loc_00E160CA: mov var_8C, eax
  loc_00E160D0: mov [ebx+00000004h], ecx
  loc_00E160D3: mov ecx, 006E0B40h ; "Perempuan"
  loc_00E160D8: mov [ebx+00000008h], ecx
  loc_00E160DB: mov ecx, var_68
  loc_00E160DE: mov [ebx+0000000Ch], ecx
  loc_00E160E1: call [edx+00000038h]
  loc_00E160E4: test eax, eax
  loc_00E160E6: fnclex
  loc_00E160E8: jge 00E160FFh
  loc_00E160EA: mov edx, var_8C
  loc_00E160F0: push 00000038h
  loc_00E160F2: push 006E09F8h
  loc_00E160F7: push edx
  loc_00E160F8: push eax
  loc_00E160F9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E160FF: lea eax, var_30
  loc_00E16102: lea ecx, var_2C
  loc_00E16105: push eax
  loc_00E16106: lea edx, var_28
  loc_00E16109: push ecx
  loc_00E1610A: lea eax, var_24
  loc_00E1610D: push edx
  loc_00E1610E: push eax
  loc_00E1610F: push 00000004h
  loc_00E16111: call [00401048h] ; __vbaFreeObjList
  loc_00E16117: add esp, 00000014h
  loc_00E1611A: lea ecx, var_44
  loc_00E1611D: call [00401024h] ; __vbaFreeVar
  loc_00E16123: mov ecx, [esi]
  loc_00E16125: push 006DCBD8h
  loc_00E1612A: push 00000000h
  loc_00E1612C: push 00000018h
  loc_00E1612E: push esi
  loc_00E1612F: call [ecx+00000490h]
  loc_00E16135: lea edx, var_28
  loc_00E16138: push eax
  loc_00E16139: push edx
  loc_00E1613A: call edi
  loc_00E1613C: push eax
  loc_00E1613D: lea eax, var_44
  loc_00E16140: push eax
  loc_00E16141: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16147: add esp, 00000010h
  loc_00E1614A: push eax
  loc_00E1614B: call [00401120h] ; __vbaCastObjVar
  loc_00E16151: lea ecx, var_2C
  loc_00E16154: push eax
  loc_00E16155: push ecx
  loc_00E16156: call edi
  loc_00E16158: mov ebx, eax
  loc_00E1615A: lea eax, var_30
  loc_00E1615D: push eax
  loc_00E1615E: push ebx
  loc_00E1615F: mov edx, [ebx]
  loc_00E16161: call [edx+00000054h]
  loc_00E16164: test eax, eax
  loc_00E16166: fnclex
  loc_00E16168: jge 00E16179h
  loc_00E1616A: push 00000054h
  loc_00E1616C: push 006DCBE8h
  loc_00E16171: push ebx
  loc_00E16172: push eax
  loc_00E16173: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16179: lea ebx, var_34
  loc_00E1617C: mov eax, var_30
  loc_00E1617F: push ebx
  loc_00E16180: mov ecx, 00000002h
  loc_00E16185: sub esp, 00000010h
  loc_00E16188: mov edx, [eax]
  loc_00E1618A: mov ebx, esp
  loc_00E1618C: mov var_8C, eax
  loc_00E16192: push eax
  loc_00E16193: mov [ebx], ecx
  loc_00E16195: mov ecx, var_60
  loc_00E16198: mov [ebx+00000004h], ecx
  loc_00E1619B: mov ecx, 00000003h
  loc_00E161A0: mov [ebx+00000008h], ecx
  loc_00E161A3: mov ecx, var_58
  loc_00E161A6: mov [ebx+0000000Ch], ecx
  loc_00E161A9: call [edx+00000028h]
  loc_00E161AC: test eax, eax
  loc_00E161AE: fnclex
  loc_00E161B0: jge 00E161C7h
  loc_00E161B2: mov edx, var_8C
  loc_00E161B8: push 00000028h
  loc_00E161BA: push 006E09E8h
  loc_00E161BF: push edx
  loc_00E161C0: push eax
  loc_00E161C1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E161C7: mov eax, var_34
  loc_00E161CA: mov ecx, [esi]
  loc_00E161CC: push esi
  loc_00E161CD: mov var_94, eax
  loc_00E161D3: call [ecx+00000390h]
  loc_00E161D9: lea edx, var_24
  loc_00E161DC: push eax
  loc_00E161DD: push edx
  loc_00E161DE: call edi
  loc_00E161E0: mov ebx, eax
  loc_00E161E2: lea ecx, var_18
  loc_00E161E5: push ecx
  loc_00E161E6: push ebx
  loc_00E161E7: mov eax, [ebx]
  loc_00E161E9: call [eax+000000A0h]
  loc_00E161EF: test eax, eax
  loc_00E161F1: fnclex
  loc_00E161F3: jge 00E16207h
  loc_00E161F5: push 000000A0h
  loc_00E161FA: push 006DCB70h
  loc_00E161FF: push ebx
  loc_00E16200: push eax
  loc_00E16201: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16207: sub esp, 00000010h
  loc_00E1620A: mov eax, var_18
  loc_00E1620D: mov ebx, esp
  loc_00E1620F: mov ecx, 00000008h
  loc_00E16214: mov var_54, ecx
  loc_00E16217: mov edx, var_94
  loc_00E1621D: mov [ebx], ecx
  loc_00E1621F: mov ecx, var_50
  loc_00E16222: mov var_4C, eax
  loc_00E16225: mov var_18, 00000000h
  loc_00E1622C: mov edx, [edx]
  loc_00E1622E: mov [ebx+00000004h], ecx
  loc_00E16231: mov [ebx+00000008h], eax
  loc_00E16234: mov eax, var_48
  loc_00E16237: mov [ebx+0000000Ch], eax
  loc_00E1623A: mov ebx, var_94
  loc_00E16240: push ebx
  loc_00E16241: call [edx+00000038h]
  loc_00E16244: test eax, eax
  loc_00E16246: fnclex
  loc_00E16248: jge 00E16259h
  loc_00E1624A: push 00000038h
  loc_00E1624C: push 006E09F8h
  loc_00E16251: push ebx
  loc_00E16252: push eax
  loc_00E16253: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16259: lea ecx, var_34
  loc_00E1625C: lea edx, var_30
  loc_00E1625F: push ecx
  loc_00E16260: lea eax, var_2C
  loc_00E16263: push edx
  loc_00E16264: lea ecx, var_28
  loc_00E16267: push eax
  loc_00E16268: lea edx, var_24
  loc_00E1626B: push ecx
  loc_00E1626C: push edx
  loc_00E1626D: push 00000005h
  loc_00E1626F: call [00401048h] ; __vbaFreeObjList
  loc_00E16275: lea eax, var_54
  loc_00E16278: lea ecx, var_44
  loc_00E1627B: push eax
  loc_00E1627C: push ecx
  loc_00E1627D: push 00000002h
  loc_00E1627F: call [00401038h] ; __vbaFreeVarList
  loc_00E16285: mov edx, [esi]
  loc_00E16287: add esp, 00000024h
  loc_00E1628A: push 006DCBD8h
  loc_00E1628F: push 00000000h
  loc_00E16291: push 00000018h
  loc_00E16293: push esi
  loc_00E16294: call [edx+00000490h]
  loc_00E1629A: push eax
  loc_00E1629B: lea eax, var_28
  loc_00E1629E: push eax
  loc_00E1629F: call edi
  loc_00E162A1: lea ecx, var_44
  loc_00E162A4: push eax
  loc_00E162A5: push ecx
  loc_00E162A6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E162AC: add esp, 00000010h
  loc_00E162AF: push eax
  loc_00E162B0: call [00401120h] ; __vbaCastObjVar
  loc_00E162B6: lea edx, var_2C
  loc_00E162B9: push eax
  loc_00E162BA: push edx
  loc_00E162BB: call edi
  loc_00E162BD: mov ebx, eax
  loc_00E162BF: lea ecx, var_30
  loc_00E162C2: push ecx
  loc_00E162C3: push ebx
  loc_00E162C4: mov eax, [ebx]
  loc_00E162C6: call [eax+00000054h]
  loc_00E162C9: test eax, eax
  loc_00E162CB: fnclex
  loc_00E162CD: jge 00E162DEh
  loc_00E162CF: push 00000054h
  loc_00E162D1: push 006DCBE8h
  loc_00E162D6: push ebx
  loc_00E162D7: push eax
  loc_00E162D8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E162DE: lea ebx, var_34
  loc_00E162E1: mov eax, var_30
  loc_00E162E4: push ebx
  loc_00E162E5: mov ecx, 00000002h
  loc_00E162EA: sub esp, 00000010h
  loc_00E162ED: mov edx, [eax]
  loc_00E162EF: mov ebx, esp
  loc_00E162F1: mov var_8C, eax
  loc_00E162F7: push eax
  loc_00E162F8: mov [ebx], ecx
  loc_00E162FA: mov ecx, var_60
  loc_00E162FD: mov [ebx+00000004h], ecx
  loc_00E16300: mov ecx, 00000004h
  loc_00E16305: mov [ebx+00000008h], ecx
  loc_00E16308: mov ecx, var_58
  loc_00E1630B: mov [ebx+0000000Ch], ecx
  loc_00E1630E: call [edx+00000028h]
  loc_00E16311: test eax, eax
  loc_00E16313: fnclex
  loc_00E16315: jge 00E1632Ch
  loc_00E16317: mov edx, var_8C
  loc_00E1631D: push 00000028h
  loc_00E1631F: push 006E09E8h
  loc_00E16324: push edx
  loc_00E16325: push eax
  loc_00E16326: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1632C: mov eax, var_34
  loc_00E1632F: mov ecx, [esi]
  loc_00E16331: push esi
  loc_00E16332: mov var_94, eax
  loc_00E16338: call [ecx+0000038Ch]
  loc_00E1633E: lea edx, var_24
  loc_00E16341: push eax
  loc_00E16342: push edx
  loc_00E16343: call edi
  loc_00E16345: mov ebx, eax
  loc_00E16347: lea ecx, var_18
  loc_00E1634A: push ecx
  loc_00E1634B: push ebx
  loc_00E1634C: mov eax, [ebx]
  loc_00E1634E: call [eax+000000A0h]
  loc_00E16354: test eax, eax
  loc_00E16356: fnclex
  loc_00E16358: jge 00E1636Ch
  loc_00E1635A: push 000000A0h
  loc_00E1635F: push 006DCB70h
  loc_00E16364: push ebx
  loc_00E16365: push eax
  loc_00E16366: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1636C: sub esp, 00000010h
  loc_00E1636F: mov eax, var_18
  loc_00E16372: mov ebx, esp
  loc_00E16374: mov ecx, 00000008h
  loc_00E16379: mov var_54, ecx
  loc_00E1637C: mov edx, var_94
  loc_00E16382: mov [ebx], ecx
  loc_00E16384: mov ecx, var_50
  loc_00E16387: mov var_4C, eax
  loc_00E1638A: mov var_18, 00000000h
  loc_00E16391: mov edx, [edx]
  loc_00E16393: mov [ebx+00000004h], ecx
  loc_00E16396: mov [ebx+00000008h], eax
  loc_00E16399: mov eax, var_48
  loc_00E1639C: mov [ebx+0000000Ch], eax
  loc_00E1639F: mov ebx, var_94
  loc_00E163A5: push ebx
  loc_00E163A6: call [edx+00000038h]
  loc_00E163A9: test eax, eax
  loc_00E163AB: fnclex
  loc_00E163AD: jge 00E163BEh
  loc_00E163AF: push 00000038h
  loc_00E163B1: push 006E09F8h
  loc_00E163B6: push ebx
  loc_00E163B7: push eax
  loc_00E163B8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E163BE: lea ecx, var_34
  loc_00E163C1: lea edx, var_30
  loc_00E163C4: push ecx
  loc_00E163C5: lea eax, var_2C
  loc_00E163C8: push edx
  loc_00E163C9: lea ecx, var_28
  loc_00E163CC: push eax
  loc_00E163CD: lea edx, var_24
  loc_00E163D0: push ecx
  loc_00E163D1: push edx
  loc_00E163D2: push 00000005h
  loc_00E163D4: call [00401048h] ; __vbaFreeObjList
  loc_00E163DA: lea eax, var_54
  loc_00E163DD: lea ecx, var_44
  loc_00E163E0: push eax
  loc_00E163E1: push ecx
  loc_00E163E2: push 00000002h
  loc_00E163E4: call [00401038h] ; __vbaFreeVarList
  loc_00E163EA: mov edx, [esi]
  loc_00E163EC: add esp, 00000024h
  loc_00E163EF: push 006DCBD8h
  loc_00E163F4: push 00000000h
  loc_00E163F6: push 00000018h
  loc_00E163F8: push esi
  loc_00E163F9: call [edx+00000490h]
  loc_00E163FF: push eax
  loc_00E16400: lea eax, var_28
  loc_00E16403: push eax
  loc_00E16404: call edi
  loc_00E16406: lea ecx, var_44
  loc_00E16409: push eax
  loc_00E1640A: push ecx
  loc_00E1640B: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16411: add esp, 00000010h
  loc_00E16414: push eax
  loc_00E16415: call [00401120h] ; __vbaCastObjVar
  loc_00E1641B: lea edx, var_2C
  loc_00E1641E: push eax
  loc_00E1641F: push edx
  loc_00E16420: call edi
  loc_00E16422: mov ebx, eax
  loc_00E16424: lea ecx, var_30
  loc_00E16427: push ecx
  loc_00E16428: push ebx
  loc_00E16429: mov eax, [ebx]
  loc_00E1642B: call [eax+00000054h]
  loc_00E1642E: test eax, eax
  loc_00E16430: fnclex
  loc_00E16432: jge 00E16443h
  loc_00E16434: push 00000054h
  loc_00E16436: push 006DCBE8h
  loc_00E1643B: push ebx
  loc_00E1643C: push eax
  loc_00E1643D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16443: lea ebx, var_34
  loc_00E16446: mov eax, var_30
  loc_00E16449: push ebx
  loc_00E1644A: mov ecx, 00000002h
  loc_00E1644F: sub esp, 00000010h
  loc_00E16452: mov edx, [eax]
  loc_00E16454: mov ebx, esp
  loc_00E16456: mov var_8C, eax
  loc_00E1645C: push eax
  loc_00E1645D: mov [ebx], ecx
  loc_00E1645F: mov ecx, var_60
  loc_00E16462: mov [ebx+00000004h], ecx
  loc_00E16465: mov ecx, 00000005h
  loc_00E1646A: mov [ebx+00000008h], ecx
  loc_00E1646D: mov ecx, var_58
  loc_00E16470: mov [ebx+0000000Ch], ecx
  loc_00E16473: call [edx+00000028h]
  loc_00E16476: test eax, eax
  loc_00E16478: fnclex
  loc_00E1647A: jge 00E16491h
  loc_00E1647C: mov edx, var_8C
  loc_00E16482: push 00000028h
  loc_00E16484: push 006E09E8h
  loc_00E16489: push edx
  loc_00E1648A: push eax
  loc_00E1648B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16491: mov eax, var_34
  loc_00E16494: mov ecx, [esi]
  loc_00E16496: push esi
  loc_00E16497: mov var_94, eax
  loc_00E1649D: call [ecx+00000388h]
  loc_00E164A3: lea edx, var_24
  loc_00E164A6: push eax
  loc_00E164A7: push edx
  loc_00E164A8: call edi
  loc_00E164AA: mov ebx, eax
  loc_00E164AC: lea ecx, var_18
  loc_00E164AF: push ecx
  loc_00E164B0: push ebx
  loc_00E164B1: mov eax, [ebx]
  loc_00E164B3: call [eax+000000A0h]
  loc_00E164B9: test eax, eax
  loc_00E164BB: fnclex
  loc_00E164BD: jge 00E164D1h
  loc_00E164BF: push 000000A0h
  loc_00E164C4: push 006DCB70h
  loc_00E164C9: push ebx
  loc_00E164CA: push eax
  loc_00E164CB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E164D1: sub esp, 00000010h
  loc_00E164D4: mov eax, var_18
  loc_00E164D7: mov ebx, esp
  loc_00E164D9: mov ecx, 00000008h
  loc_00E164DE: mov var_54, ecx
  loc_00E164E1: mov edx, var_94
  loc_00E164E7: mov [ebx], ecx
  loc_00E164E9: mov ecx, var_50
  loc_00E164EC: mov var_4C, eax
  loc_00E164EF: mov var_18, 00000000h
  loc_00E164F6: mov edx, [edx]
  loc_00E164F8: mov [ebx+00000004h], ecx
  loc_00E164FB: mov [ebx+00000008h], eax
  loc_00E164FE: mov eax, var_48
  loc_00E16501: mov [ebx+0000000Ch], eax
  loc_00E16504: mov ebx, var_94
  loc_00E1650A: push ebx
  loc_00E1650B: call [edx+00000038h]
  loc_00E1650E: test eax, eax
  loc_00E16510: fnclex
  loc_00E16512: jge 00E16523h
  loc_00E16514: push 00000038h
  loc_00E16516: push 006E09F8h
  loc_00E1651B: push ebx
  loc_00E1651C: push eax
  loc_00E1651D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16523: lea ecx, var_34
  loc_00E16526: lea edx, var_30
  loc_00E16529: push ecx
  loc_00E1652A: lea eax, var_2C
  loc_00E1652D: push edx
  loc_00E1652E: lea ecx, var_28
  loc_00E16531: push eax
  loc_00E16532: lea edx, var_24
  loc_00E16535: push ecx
  loc_00E16536: push edx
  loc_00E16537: push 00000005h
  loc_00E16539: call [00401048h] ; __vbaFreeObjList
  loc_00E1653F: lea eax, var_54
  loc_00E16542: lea ecx, var_44
  loc_00E16545: push eax
  loc_00E16546: push ecx
  loc_00E16547: push 00000002h
  loc_00E16549: call [00401038h] ; __vbaFreeVarList
  loc_00E1654F: mov edx, [esi]
  loc_00E16551: add esp, 00000024h
  loc_00E16554: push 006DCBD8h
  loc_00E16559: push 00000000h
  loc_00E1655B: push 00000018h
  loc_00E1655D: push esi
  loc_00E1655E: call [edx+00000490h]
  loc_00E16564: push eax
  loc_00E16565: lea eax, var_28
  loc_00E16568: push eax
  loc_00E16569: call edi
  loc_00E1656B: lea ecx, var_44
  loc_00E1656E: push eax
  loc_00E1656F: push ecx
  loc_00E16570: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16576: add esp, 00000010h
  loc_00E16579: push eax
  loc_00E1657A: call [00401120h] ; __vbaCastObjVar
  loc_00E16580: lea edx, var_2C
  loc_00E16583: push eax
  loc_00E16584: push edx
  loc_00E16585: call edi
  loc_00E16587: mov ebx, eax
  loc_00E16589: lea ecx, var_30
  loc_00E1658C: push ecx
  loc_00E1658D: push ebx
  loc_00E1658E: mov eax, [ebx]
  loc_00E16590: call [eax+00000054h]
  loc_00E16593: test eax, eax
  loc_00E16595: fnclex
  loc_00E16597: jge 00E165A8h
  loc_00E16599: push 00000054h
  loc_00E1659B: push 006DCBE8h
  loc_00E165A0: push ebx
  loc_00E165A1: push eax
  loc_00E165A2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E165A8: lea ebx, var_34
  loc_00E165AB: mov eax, var_30
  loc_00E165AE: push ebx
  loc_00E165AF: mov ecx, 00000002h
  loc_00E165B4: sub esp, 00000010h
  loc_00E165B7: mov edx, [eax]
  loc_00E165B9: mov ebx, esp
  loc_00E165BB: mov var_84, eax
  loc_00E165C1: push eax
  loc_00E165C2: mov [ebx], ecx
  loc_00E165C4: mov ecx, var_60
  loc_00E165C7: mov [ebx+00000004h], ecx
  loc_00E165CA: mov ecx, 00000006h
  loc_00E165CF: mov [ebx+00000008h], ecx
  loc_00E165D2: mov ecx, var_58
  loc_00E165D5: mov [ebx+0000000Ch], ecx
  loc_00E165D8: call [edx+00000028h]
  loc_00E165DB: test eax, eax
  loc_00E165DD: fnclex
  loc_00E165DF: jge 00E165F6h
  loc_00E165E1: mov edx, var_84
  loc_00E165E7: push 00000028h
  loc_00E165E9: push 006E09E8h
  loc_00E165EE: push edx
  loc_00E165EF: push eax
  loc_00E165F0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E165F6: mov eax, var_34
  loc_00E165F9: push 00000000h
  loc_00E165FB: mov var_8C, eax
  loc_00E16601: push 00000014h
  loc_00E16603: mov ebx, [eax]
  loc_00E16605: mov eax, [esi]
  loc_00E16607: push esi
  loc_00E16608: call [eax+0000048Ch]
  loc_00E1660E: lea ecx, var_24
  loc_00E16611: push eax
  loc_00E16612: push ecx
  loc_00E16613: call edi
  loc_00E16615: lea edx, var_54
  loc_00E16618: push eax
  loc_00E16619: push edx
  loc_00E1661A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16620: mov edx, [eax]
  loc_00E16622: mov ecx, esp
  loc_00E16624: mov var_E0, ebx
  loc_00E1662A: mov ebx, var_8C
  loc_00E16630: mov [ecx], edx
  loc_00E16632: mov edx, [eax+00000004h]
  loc_00E16635: push ebx
  loc_00E16636: mov [ecx+00000004h], edx
  loc_00E16639: mov edx, [eax+00000008h]
  loc_00E1663C: mov eax, [eax+0000000Ch]
  loc_00E1663F: mov [ecx+00000008h], edx
  loc_00E16642: mov [ecx+0000000Ch], eax
  loc_00E16645: mov ecx, var_E0
  loc_00E1664B: call [ecx+00000038h]
  loc_00E1664E: test eax, eax
  loc_00E16650: fnclex
  loc_00E16652: jge 00E16663h
  loc_00E16654: push 00000038h
  loc_00E16656: push 006E09F8h
  loc_00E1665B: push ebx
  loc_00E1665C: push eax
  loc_00E1665D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16663: lea edx, var_34
  loc_00E16666: lea eax, var_30
  loc_00E16669: push edx
  loc_00E1666A: lea ecx, var_2C
  loc_00E1666D: push eax
  loc_00E1666E: lea edx, var_28
  loc_00E16671: push ecx
  loc_00E16672: lea eax, var_24
  loc_00E16675: push edx
  loc_00E16676: push eax
  loc_00E16677: push 00000005h
  loc_00E16679: call [00401048h] ; __vbaFreeObjList
  loc_00E1667F: lea ecx, var_54
  loc_00E16682: lea edx, var_44
  loc_00E16685: push ecx
  loc_00E16686: push edx
  loc_00E16687: push 00000002h
  loc_00E16689: call [00401038h] ; __vbaFreeVarList
  loc_00E1668F: mov eax, [esi]
  loc_00E16691: add esp, 00000024h
  loc_00E16694: push 006DCBD8h
  loc_00E16699: push 00000000h
  loc_00E1669B: push 00000018h
  loc_00E1669D: push esi
  loc_00E1669E: call [eax+00000490h]
  loc_00E166A4: lea ecx, var_28
  loc_00E166A7: push eax
  loc_00E166A8: push ecx
  loc_00E166A9: call edi
  loc_00E166AB: lea edx, var_44
  loc_00E166AE: push eax
  loc_00E166AF: push edx
  loc_00E166B0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E166B6: add esp, 00000010h
  loc_00E166B9: push eax
  loc_00E166BA: call [00401120h] ; __vbaCastObjVar
  loc_00E166C0: push eax
  loc_00E166C1: lea eax, var_2C
  loc_00E166C4: push eax
  loc_00E166C5: call edi
  loc_00E166C7: mov ebx, eax
  loc_00E166C9: lea edx, var_30
  loc_00E166CC: push edx
  loc_00E166CD: push ebx
  loc_00E166CE: mov ecx, [ebx]
  loc_00E166D0: call [ecx+00000054h]
  loc_00E166D3: test eax, eax
  loc_00E166D5: fnclex
  loc_00E166D7: jge 00E166E8h
  loc_00E166D9: push 00000054h
  loc_00E166DB: push 006DCBE8h
  loc_00E166E0: push ebx
  loc_00E166E1: push eax
  loc_00E166E2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E166E8: lea ebx, var_34
  loc_00E166EB: mov eax, var_30
  loc_00E166EE: push ebx
  loc_00E166EF: mov ecx, 00000002h
  loc_00E166F4: sub esp, 00000010h
  loc_00E166F7: mov edx, [eax]
  loc_00E166F9: mov ebx, esp
  loc_00E166FB: mov var_8C, eax
  loc_00E16701: push eax
  loc_00E16702: mov [ebx], ecx
  loc_00E16704: mov ecx, var_60
  loc_00E16707: mov [ebx+00000004h], ecx
  loc_00E1670A: mov ecx, 00000007h
  loc_00E1670F: mov [ebx+00000008h], ecx
  loc_00E16712: mov ecx, var_58
  loc_00E16715: mov [ebx+0000000Ch], ecx
  loc_00E16718: call [edx+00000028h]
  loc_00E1671B: test eax, eax
  loc_00E1671D: fnclex
  loc_00E1671F: jge 00E16736h
  loc_00E16721: mov edx, var_8C
  loc_00E16727: push 00000028h
  loc_00E16729: push 006E09E8h
  loc_00E1672E: push edx
  loc_00E1672F: push eax
  loc_00E16730: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16736: mov eax, var_34
  loc_00E16739: mov ecx, [esi]
  loc_00E1673B: push esi
  loc_00E1673C: mov var_94, eax
  loc_00E16742: call [ecx+00000378h]
  loc_00E16748: lea edx, var_24
  loc_00E1674B: push eax
  loc_00E1674C: push edx
  loc_00E1674D: call edi
  loc_00E1674F: mov ebx, eax
  loc_00E16751: lea ecx, var_18
  loc_00E16754: push ecx
  loc_00E16755: push ebx
  loc_00E16756: mov eax, [ebx]
  loc_00E16758: call [eax+000000A8h]
  loc_00E1675E: test eax, eax
  loc_00E16760: fnclex
  loc_00E16762: jge 00E16776h
  loc_00E16764: push 000000A8h
  loc_00E16769: push 006E0400h
  loc_00E1676E: push ebx
  loc_00E1676F: push eax
  loc_00E16770: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16776: sub esp, 00000010h
  loc_00E16779: mov eax, var_18
  loc_00E1677C: mov ebx, esp
  loc_00E1677E: mov ecx, 00000008h
  loc_00E16783: mov var_54, ecx
  loc_00E16786: mov edx, var_94
  loc_00E1678C: mov [ebx], ecx
  loc_00E1678E: mov ecx, var_50
  loc_00E16791: mov var_4C, eax
  loc_00E16794: mov var_18, 00000000h
  loc_00E1679B: mov edx, [edx]
  loc_00E1679D: mov [ebx+00000004h], ecx
  loc_00E167A0: mov [ebx+00000008h], eax
  loc_00E167A3: mov eax, var_48
  loc_00E167A6: mov [ebx+0000000Ch], eax
  loc_00E167A9: mov ebx, var_94
  loc_00E167AF: push ebx
  loc_00E167B0: call [edx+00000038h]
  loc_00E167B3: test eax, eax
  loc_00E167B5: fnclex
  loc_00E167B7: jge 00E167C8h
  loc_00E167B9: push 00000038h
  loc_00E167BB: push 006E09F8h
  loc_00E167C0: push ebx
  loc_00E167C1: push eax
  loc_00E167C2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E167C8: lea ecx, var_34
  loc_00E167CB: lea edx, var_30
  loc_00E167CE: push ecx
  loc_00E167CF: lea eax, var_2C
  loc_00E167D2: push edx
  loc_00E167D3: lea ecx, var_28
  loc_00E167D6: push eax
  loc_00E167D7: lea edx, var_24
  loc_00E167DA: push ecx
  loc_00E167DB: push edx
  loc_00E167DC: push 00000005h
  loc_00E167DE: call [00401048h] ; __vbaFreeObjList
  loc_00E167E4: lea eax, var_54
  loc_00E167E7: lea ecx, var_44
  loc_00E167EA: push eax
  loc_00E167EB: push ecx
  loc_00E167EC: push 00000002h
  loc_00E167EE: call [00401038h] ; __vbaFreeVarList
  loc_00E167F4: mov edx, [esi]
  loc_00E167F6: add esp, 00000024h
  loc_00E167F9: push 006DCBD8h
  loc_00E167FE: push 00000000h
  loc_00E16800: push 00000018h
  loc_00E16802: push esi
  loc_00E16803: call [edx+00000490h]
  loc_00E16809: push eax
  loc_00E1680A: lea eax, var_28
  loc_00E1680D: push eax
  loc_00E1680E: call edi
  loc_00E16810: lea ecx, var_44
  loc_00E16813: push eax
  loc_00E16814: push ecx
  loc_00E16815: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1681B: add esp, 00000010h
  loc_00E1681E: push eax
  loc_00E1681F: call [00401120h] ; __vbaCastObjVar
  loc_00E16825: lea edx, var_2C
  loc_00E16828: push eax
  loc_00E16829: push edx
  loc_00E1682A: call edi
  loc_00E1682C: mov ebx, eax
  loc_00E1682E: lea ecx, var_30
  loc_00E16831: push ecx
  loc_00E16832: push ebx
  loc_00E16833: mov eax, [ebx]
  loc_00E16835: call [eax+00000054h]
  loc_00E16838: test eax, eax
  loc_00E1683A: fnclex
  loc_00E1683C: jge 00E1684Dh
  loc_00E1683E: push 00000054h
  loc_00E16840: push 006DCBE8h
  loc_00E16845: push ebx
  loc_00E16846: push eax
  loc_00E16847: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1684D: lea ebx, var_34
  loc_00E16850: mov eax, var_30
  loc_00E16853: push ebx
  loc_00E16854: mov ecx, 00000002h
  loc_00E16859: sub esp, 00000010h
  loc_00E1685C: mov edx, [eax]
  loc_00E1685E: mov ebx, esp
  loc_00E16860: mov var_8C, eax
  loc_00E16866: push eax
  loc_00E16867: mov [ebx], ecx
  loc_00E16869: mov ecx, var_60
  loc_00E1686C: mov [ebx+00000004h], ecx
  loc_00E1686F: mov ecx, 00000008h
  loc_00E16874: mov [ebx+00000008h], ecx
  loc_00E16877: mov ecx, var_58
  loc_00E1687A: mov [ebx+0000000Ch], ecx
  loc_00E1687D: call [edx+00000028h]
  loc_00E16880: test eax, eax
  loc_00E16882: fnclex
  loc_00E16884: jge 00E1689Bh
  loc_00E16886: mov edx, var_8C
  loc_00E1688C: push 00000028h
  loc_00E1688E: push 006E09E8h
  loc_00E16893: push edx
  loc_00E16894: push eax
  loc_00E16895: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1689B: mov eax, var_34
  loc_00E1689E: mov ecx, [esi]
  loc_00E168A0: push esi
  loc_00E168A1: mov var_94, eax
  loc_00E168A7: call [ecx+00000384h]
  loc_00E168AD: lea edx, var_24
  loc_00E168B0: push eax
  loc_00E168B1: push edx
  loc_00E168B2: call edi
  loc_00E168B4: mov ebx, eax
  loc_00E168B6: lea ecx, var_18
  loc_00E168B9: push ecx
  loc_00E168BA: push ebx
  loc_00E168BB: mov eax, [ebx]
  loc_00E168BD: call [eax+000000A0h]
  loc_00E168C3: test eax, eax
  loc_00E168C5: fnclex
  loc_00E168C7: jge 00E168DBh
  loc_00E168C9: push 000000A0h
  loc_00E168CE: push 006DCB70h
  loc_00E168D3: push ebx
  loc_00E168D4: push eax
  loc_00E168D5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E168DB: sub esp, 00000010h
  loc_00E168DE: mov eax, var_18
  loc_00E168E1: mov ebx, esp
  loc_00E168E3: mov ecx, 00000008h
  loc_00E168E8: mov var_54, ecx
  loc_00E168EB: mov edx, var_94
  loc_00E168F1: mov [ebx], ecx
  loc_00E168F3: mov ecx, var_50
  loc_00E168F6: mov var_4C, eax
  loc_00E168F9: mov var_18, 00000000h
  loc_00E16900: mov edx, [edx]
  loc_00E16902: mov [ebx+00000004h], ecx
  loc_00E16905: mov [ebx+00000008h], eax
  loc_00E16908: mov eax, var_48
  loc_00E1690B: mov [ebx+0000000Ch], eax
  loc_00E1690E: mov ebx, var_94
  loc_00E16914: push ebx
  loc_00E16915: call [edx+00000038h]
  loc_00E16918: test eax, eax
  loc_00E1691A: fnclex
  loc_00E1691C: jge 00E1692Dh
  loc_00E1691E: push 00000038h
  loc_00E16920: push 006E09F8h
  loc_00E16925: push ebx
  loc_00E16926: push eax
  loc_00E16927: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1692D: lea ecx, var_34
  loc_00E16930: lea edx, var_30
  loc_00E16933: push ecx
  loc_00E16934: lea eax, var_2C
  loc_00E16937: push edx
  loc_00E16938: lea ecx, var_28
  loc_00E1693B: push eax
  loc_00E1693C: lea edx, var_24
  loc_00E1693F: push ecx
  loc_00E16940: push edx
  loc_00E16941: push 00000005h
  loc_00E16943: call [00401048h] ; __vbaFreeObjList
  loc_00E16949: lea eax, var_54
  loc_00E1694C: lea ecx, var_44
  loc_00E1694F: push eax
  loc_00E16950: push ecx
  loc_00E16951: push 00000002h
  loc_00E16953: call [00401038h] ; __vbaFreeVarList
  loc_00E16959: mov edx, [esi]
  loc_00E1695B: add esp, 00000024h
  loc_00E1695E: push 006DCBD8h
  loc_00E16963: push 00000000h
  loc_00E16965: push 00000018h
  loc_00E16967: push esi
  loc_00E16968: call [edx+00000490h]
  loc_00E1696E: push eax
  loc_00E1696F: lea eax, var_28
  loc_00E16972: push eax
  loc_00E16973: call edi
  loc_00E16975: lea ecx, var_44
  loc_00E16978: push eax
  loc_00E16979: push ecx
  loc_00E1697A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16980: add esp, 00000010h
  loc_00E16983: push eax
  loc_00E16984: call [00401120h] ; __vbaCastObjVar
  loc_00E1698A: lea edx, var_2C
  loc_00E1698D: push eax
  loc_00E1698E: push edx
  loc_00E1698F: call edi
  loc_00E16991: mov ebx, eax
  loc_00E16993: lea ecx, var_30
  loc_00E16996: push ecx
  loc_00E16997: push ebx
  loc_00E16998: mov eax, [ebx]
  loc_00E1699A: call [eax+00000054h]
  loc_00E1699D: test eax, eax
  loc_00E1699F: fnclex
  loc_00E169A1: jge 00E169B2h
  loc_00E169A3: push 00000054h
  loc_00E169A5: push 006DCBE8h
  loc_00E169AA: push ebx
  loc_00E169AB: push eax
  loc_00E169AC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E169B2: lea ebx, var_34
  loc_00E169B5: mov eax, var_30
  loc_00E169B8: push ebx
  loc_00E169B9: mov ecx, 00000002h
  loc_00E169BE: sub esp, 00000010h
  loc_00E169C1: mov edx, [eax]
  loc_00E169C3: mov ebx, esp
  loc_00E169C5: mov var_8C, eax
  loc_00E169CB: push eax
  loc_00E169CC: mov [ebx], ecx
  loc_00E169CE: mov ecx, var_60
  loc_00E169D1: mov [ebx+00000004h], ecx
  loc_00E169D4: mov ecx, 00000009h
  loc_00E169D9: mov [ebx+00000008h], ecx
  loc_00E169DC: mov ecx, var_58
  loc_00E169DF: mov [ebx+0000000Ch], ecx
  loc_00E169E2: call [edx+00000028h]
  loc_00E169E5: test eax, eax
  loc_00E169E7: fnclex
  loc_00E169E9: jge 00E16A00h
  loc_00E169EB: mov edx, var_8C
  loc_00E169F1: push 00000028h
  loc_00E169F3: push 006E09E8h
  loc_00E169F8: push edx
  loc_00E169F9: push eax
  loc_00E169FA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16A00: mov eax, var_34
  loc_00E16A03: mov ecx, [esi]
  loc_00E16A05: push esi
  loc_00E16A06: mov var_94, eax
  loc_00E16A0C: call [ecx+00000374h]
  loc_00E16A12: lea edx, var_24
  loc_00E16A15: push eax
  loc_00E16A16: push edx
  loc_00E16A17: call edi
  loc_00E16A19: mov ebx, eax
  loc_00E16A1B: lea ecx, var_18
  loc_00E16A1E: push ecx
  loc_00E16A1F: push ebx
  loc_00E16A20: mov eax, [ebx]
  loc_00E16A22: call [eax+000000A0h]
  loc_00E16A28: test eax, eax
  loc_00E16A2A: fnclex
  loc_00E16A2C: jge 00E16A40h
  loc_00E16A2E: push 000000A0h
  loc_00E16A33: push 006DCB70h
  loc_00E16A38: push ebx
  loc_00E16A39: push eax
  loc_00E16A3A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16A40: sub esp, 00000010h
  loc_00E16A43: mov eax, var_18
  loc_00E16A46: mov ebx, esp
  loc_00E16A48: mov ecx, 00000008h
  loc_00E16A4D: mov var_54, ecx
  loc_00E16A50: mov edx, var_94
  loc_00E16A56: mov [ebx], ecx
  loc_00E16A58: mov ecx, var_50
  loc_00E16A5B: mov var_4C, eax
  loc_00E16A5E: mov var_18, 00000000h
  loc_00E16A65: mov edx, [edx]
  loc_00E16A67: mov [ebx+00000004h], ecx
  loc_00E16A6A: mov [ebx+00000008h], eax
  loc_00E16A6D: mov eax, var_48
  loc_00E16A70: mov [ebx+0000000Ch], eax
  loc_00E16A73: mov ebx, var_94
  loc_00E16A79: push ebx
  loc_00E16A7A: call [edx+00000038h]
  loc_00E16A7D: test eax, eax
  loc_00E16A7F: fnclex
  loc_00E16A81: jge 00E16A92h
  loc_00E16A83: push 00000038h
  loc_00E16A85: push 006E09F8h
  loc_00E16A8A: push ebx
  loc_00E16A8B: push eax
  loc_00E16A8C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16A92: lea ecx, var_34
  loc_00E16A95: lea edx, var_30
  loc_00E16A98: push ecx
  loc_00E16A99: lea eax, var_2C
  loc_00E16A9C: push edx
  loc_00E16A9D: lea ecx, var_28
  loc_00E16AA0: push eax
  loc_00E16AA1: lea edx, var_24
  loc_00E16AA4: push ecx
  loc_00E16AA5: push edx
  loc_00E16AA6: push 00000005h
  loc_00E16AA8: call [00401048h] ; __vbaFreeObjList
  loc_00E16AAE: lea eax, var_54
  loc_00E16AB1: lea ecx, var_44
  loc_00E16AB4: push eax
  loc_00E16AB5: push ecx
  loc_00E16AB6: push 00000002h
  loc_00E16AB8: call [00401038h] ; __vbaFreeVarList
  loc_00E16ABE: mov edx, [esi]
  loc_00E16AC0: add esp, 00000024h
  loc_00E16AC3: push 006DCBD8h
  loc_00E16AC8: push 00000000h
  loc_00E16ACA: push 00000018h
  loc_00E16ACC: push esi
  loc_00E16ACD: call [edx+00000490h]
  loc_00E16AD3: push eax
  loc_00E16AD4: lea eax, var_28
  loc_00E16AD7: push eax
  loc_00E16AD8: call edi
  loc_00E16ADA: lea ecx, var_44
  loc_00E16ADD: push eax
  loc_00E16ADE: push ecx
  loc_00E16ADF: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16AE5: add esp, 00000010h
  loc_00E16AE8: push eax
  loc_00E16AE9: call [00401120h] ; __vbaCastObjVar
  loc_00E16AEF: lea edx, var_2C
  loc_00E16AF2: push eax
  loc_00E16AF3: push edx
  loc_00E16AF4: call edi
  loc_00E16AF6: mov ebx, eax
  loc_00E16AF8: lea ecx, var_30
  loc_00E16AFB: push ecx
  loc_00E16AFC: push ebx
  loc_00E16AFD: mov eax, [ebx]
  loc_00E16AFF: call [eax+00000054h]
  loc_00E16B02: test eax, eax
  loc_00E16B04: fnclex
  loc_00E16B06: jge 00E16B17h
  loc_00E16B08: push 00000054h
  loc_00E16B0A: push 006DCBE8h
  loc_00E16B0F: push ebx
  loc_00E16B10: push eax
  loc_00E16B11: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16B17: lea ebx, var_34
  loc_00E16B1A: mov eax, var_30
  loc_00E16B1D: push ebx
  loc_00E16B1E: mov ecx, 00000002h
  loc_00E16B23: sub esp, 00000010h
  loc_00E16B26: mov edx, [eax]
  loc_00E16B28: mov ebx, esp
  loc_00E16B2A: mov var_8C, eax
  loc_00E16B30: push eax
  loc_00E16B31: mov [ebx], ecx
  loc_00E16B33: mov ecx, var_60
  loc_00E16B36: mov [ebx+00000004h], ecx
  loc_00E16B39: mov ecx, 0000000Ah
  loc_00E16B3E: mov [ebx+00000008h], ecx
  loc_00E16B41: mov ecx, var_58
  loc_00E16B44: mov [ebx+0000000Ch], ecx
  loc_00E16B47: call [edx+00000028h]
  loc_00E16B4A: test eax, eax
  loc_00E16B4C: fnclex
  loc_00E16B4E: jge 00E16B65h
  loc_00E16B50: mov edx, var_8C
  loc_00E16B56: push 00000028h
  loc_00E16B58: push 006E09E8h
  loc_00E16B5D: push edx
  loc_00E16B5E: push eax
  loc_00E16B5F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16B65: mov eax, var_34
  loc_00E16B68: mov ecx, [esi]
  loc_00E16B6A: push esi
  loc_00E16B6B: mov var_94, eax
  loc_00E16B71: call [ecx+00000370h]
  loc_00E16B77: lea edx, var_24
  loc_00E16B7A: push eax
  loc_00E16B7B: push edx
  loc_00E16B7C: call edi
  loc_00E16B7E: mov ebx, eax
  loc_00E16B80: lea ecx, var_18
  loc_00E16B83: push ecx
  loc_00E16B84: push ebx
  loc_00E16B85: mov eax, [ebx]
  loc_00E16B87: call [eax+000000A0h]
  loc_00E16B8D: test eax, eax
  loc_00E16B8F: fnclex
  loc_00E16B91: jge 00E16BA5h
  loc_00E16B93: push 000000A0h
  loc_00E16B98: push 006DCB70h
  loc_00E16B9D: push ebx
  loc_00E16B9E: push eax
  loc_00E16B9F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16BA5: sub esp, 00000010h
  loc_00E16BA8: mov eax, var_18
  loc_00E16BAB: mov ebx, esp
  loc_00E16BAD: mov ecx, 00000008h
  loc_00E16BB2: mov var_54, ecx
  loc_00E16BB5: mov edx, var_94
  loc_00E16BBB: mov [ebx], ecx
  loc_00E16BBD: mov ecx, var_50
  loc_00E16BC0: mov var_4C, eax
  loc_00E16BC3: mov var_18, 00000000h
  loc_00E16BCA: mov edx, [edx]
  loc_00E16BCC: mov [ebx+00000004h], ecx
  loc_00E16BCF: mov [ebx+00000008h], eax
  loc_00E16BD2: mov eax, var_48
  loc_00E16BD5: mov [ebx+0000000Ch], eax
  loc_00E16BD8: mov ebx, var_94
  loc_00E16BDE: push ebx
  loc_00E16BDF: call [edx+00000038h]
  loc_00E16BE2: test eax, eax
  loc_00E16BE4: fnclex
  loc_00E16BE6: jge 00E16BF7h
  loc_00E16BE8: push 00000038h
  loc_00E16BEA: push 006E09F8h
  loc_00E16BEF: push ebx
  loc_00E16BF0: push eax
  loc_00E16BF1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16BF7: lea ecx, var_34
  loc_00E16BFA: lea edx, var_30
  loc_00E16BFD: push ecx
  loc_00E16BFE: lea eax, var_2C
  loc_00E16C01: push edx
  loc_00E16C02: lea ecx, var_28
  loc_00E16C05: push eax
  loc_00E16C06: lea edx, var_24
  loc_00E16C09: push ecx
  loc_00E16C0A: push edx
  loc_00E16C0B: push 00000005h
  loc_00E16C0D: call [00401048h] ; __vbaFreeObjList
  loc_00E16C13: lea eax, var_54
  loc_00E16C16: lea ecx, var_44
  loc_00E16C19: push eax
  loc_00E16C1A: push ecx
  loc_00E16C1B: push 00000002h
  loc_00E16C1D: call [00401038h] ; __vbaFreeVarList
  loc_00E16C23: mov edx, [esi]
  loc_00E16C25: add esp, 00000024h
  loc_00E16C28: push 006DCBD8h
  loc_00E16C2D: push 00000000h
  loc_00E16C2F: push 00000018h
  loc_00E16C31: push esi
  loc_00E16C32: call [edx+00000490h]
  loc_00E16C38: push eax
  loc_00E16C39: lea eax, var_28
  loc_00E16C3C: push eax
  loc_00E16C3D: call edi
  loc_00E16C3F: lea ecx, var_44
  loc_00E16C42: push eax
  loc_00E16C43: push ecx
  loc_00E16C44: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16C4A: add esp, 00000010h
  loc_00E16C4D: push eax
  loc_00E16C4E: call [00401120h] ; __vbaCastObjVar
  loc_00E16C54: lea edx, var_2C
  loc_00E16C57: push eax
  loc_00E16C58: push edx
  loc_00E16C59: call edi
  loc_00E16C5B: mov ebx, eax
  loc_00E16C5D: lea ecx, var_30
  loc_00E16C60: push ecx
  loc_00E16C61: push ebx
  loc_00E16C62: mov eax, [ebx]
  loc_00E16C64: call [eax+00000054h]
  loc_00E16C67: test eax, eax
  loc_00E16C69: fnclex
  loc_00E16C6B: jge 00E16C7Ch
  loc_00E16C6D: push 00000054h
  loc_00E16C6F: push 006DCBE8h
  loc_00E16C74: push ebx
  loc_00E16C75: push eax
  loc_00E16C76: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16C7C: lea ebx, var_34
  loc_00E16C7F: mov eax, var_30
  loc_00E16C82: push ebx
  loc_00E16C83: mov ecx, 00000002h
  loc_00E16C88: sub esp, 00000010h
  loc_00E16C8B: mov edx, [eax]
  loc_00E16C8D: mov ebx, esp
  loc_00E16C8F: mov var_8C, eax
  loc_00E16C95: push eax
  loc_00E16C96: mov [ebx], ecx
  loc_00E16C98: mov ecx, var_60
  loc_00E16C9B: mov [ebx+00000004h], ecx
  loc_00E16C9E: mov ecx, 0000000Bh
  loc_00E16CA3: mov [ebx+00000008h], ecx
  loc_00E16CA6: mov ecx, var_58
  loc_00E16CA9: mov [ebx+0000000Ch], ecx
  loc_00E16CAC: call [edx+00000028h]
  loc_00E16CAF: test eax, eax
  loc_00E16CB1: fnclex
  loc_00E16CB3: jge 00E16CCAh
  loc_00E16CB5: mov edx, var_8C
  loc_00E16CBB: push 00000028h
  loc_00E16CBD: push 006E09E8h
  loc_00E16CC2: push edx
  loc_00E16CC3: push eax
  loc_00E16CC4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16CCA: mov eax, var_34
  loc_00E16CCD: mov ecx, [esi]
  loc_00E16CCF: push esi
  loc_00E16CD0: mov var_94, eax
  loc_00E16CD6: call [ecx+0000036Ch]
  loc_00E16CDC: lea edx, var_24
  loc_00E16CDF: push eax
  loc_00E16CE0: push edx
  loc_00E16CE1: call edi
  loc_00E16CE3: mov ebx, eax
  loc_00E16CE5: lea ecx, var_18
  loc_00E16CE8: push ecx
  loc_00E16CE9: push ebx
  loc_00E16CEA: mov eax, [ebx]
  loc_00E16CEC: call [eax+000000A0h]
  loc_00E16CF2: test eax, eax
  loc_00E16CF4: fnclex
  loc_00E16CF6: jge 00E16D0Ah
  loc_00E16CF8: push 000000A0h
  loc_00E16CFD: push 006DCB70h
  loc_00E16D02: push ebx
  loc_00E16D03: push eax
  loc_00E16D04: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16D0A: sub esp, 00000010h
  loc_00E16D0D: mov eax, var_18
  loc_00E16D10: mov ebx, esp
  loc_00E16D12: mov ecx, 00000008h
  loc_00E16D17: mov var_54, ecx
  loc_00E16D1A: mov edx, var_94
  loc_00E16D20: mov [ebx], ecx
  loc_00E16D22: mov ecx, var_50
  loc_00E16D25: mov var_4C, eax
  loc_00E16D28: mov var_18, 00000000h
  loc_00E16D2F: mov edx, [edx]
  loc_00E16D31: mov [ebx+00000004h], ecx
  loc_00E16D34: mov [ebx+00000008h], eax
  loc_00E16D37: mov eax, var_48
  loc_00E16D3A: mov [ebx+0000000Ch], eax
  loc_00E16D3D: mov ebx, var_94
  loc_00E16D43: push ebx
  loc_00E16D44: call [edx+00000038h]
  loc_00E16D47: test eax, eax
  loc_00E16D49: fnclex
  loc_00E16D4B: jge 00E16D5Ch
  loc_00E16D4D: push 00000038h
  loc_00E16D4F: push 006E09F8h
  loc_00E16D54: push ebx
  loc_00E16D55: push eax
  loc_00E16D56: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16D5C: lea ecx, var_34
  loc_00E16D5F: lea edx, var_30
  loc_00E16D62: push ecx
  loc_00E16D63: lea eax, var_2C
  loc_00E16D66: push edx
  loc_00E16D67: lea ecx, var_28
  loc_00E16D6A: push eax
  loc_00E16D6B: lea edx, var_24
  loc_00E16D6E: push ecx
  loc_00E16D6F: push edx
  loc_00E16D70: push 00000005h
  loc_00E16D72: call [00401048h] ; __vbaFreeObjList
  loc_00E16D78: lea eax, var_54
  loc_00E16D7B: lea ecx, var_44
  loc_00E16D7E: push eax
  loc_00E16D7F: push ecx
  loc_00E16D80: push 00000002h
  loc_00E16D82: call [00401038h] ; __vbaFreeVarList
  loc_00E16D88: mov edx, [esi]
  loc_00E16D8A: add esp, 00000024h
  loc_00E16D8D: push 006DCBD8h
  loc_00E16D92: push 00000000h
  loc_00E16D94: push 00000018h
  loc_00E16D96: push esi
  loc_00E16D97: call [edx+00000490h]
  loc_00E16D9D: push eax
  loc_00E16D9E: lea eax, var_28
  loc_00E16DA1: push eax
  loc_00E16DA2: call edi
  loc_00E16DA4: lea ecx, var_44
  loc_00E16DA7: push eax
  loc_00E16DA8: push ecx
  loc_00E16DA9: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16DAF: add esp, 00000010h
  loc_00E16DB2: push eax
  loc_00E16DB3: call [00401120h] ; __vbaCastObjVar
  loc_00E16DB9: lea edx, var_2C
  loc_00E16DBC: push eax
  loc_00E16DBD: push edx
  loc_00E16DBE: call edi
  loc_00E16DC0: mov ebx, eax
  loc_00E16DC2: lea ecx, var_30
  loc_00E16DC5: push ecx
  loc_00E16DC6: push ebx
  loc_00E16DC7: mov eax, [ebx]
  loc_00E16DC9: call [eax+00000054h]
  loc_00E16DCC: test eax, eax
  loc_00E16DCE: fnclex
  loc_00E16DD0: jge 00E16DE1h
  loc_00E16DD2: push 00000054h
  loc_00E16DD4: push 006DCBE8h
  loc_00E16DD9: push ebx
  loc_00E16DDA: push eax
  loc_00E16DDB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16DE1: lea ebx, var_34
  loc_00E16DE4: mov eax, var_30
  loc_00E16DE7: push ebx
  loc_00E16DE8: mov ecx, 00000002h
  loc_00E16DED: sub esp, 00000010h
  loc_00E16DF0: mov edx, [eax]
  loc_00E16DF2: mov ebx, esp
  loc_00E16DF4: mov var_8C, eax
  loc_00E16DFA: push eax
  loc_00E16DFB: mov [ebx], ecx
  loc_00E16DFD: mov ecx, var_60
  loc_00E16E00: mov [ebx+00000004h], ecx
  loc_00E16E03: mov ecx, 0000000Ch
  loc_00E16E08: mov [ebx+00000008h], ecx
  loc_00E16E0B: mov ecx, var_58
  loc_00E16E0E: mov [ebx+0000000Ch], ecx
  loc_00E16E11: call [edx+00000028h]
  loc_00E16E14: test eax, eax
  loc_00E16E16: fnclex
  loc_00E16E18: jge 00E16E2Fh
  loc_00E16E1A: mov edx, var_8C
  loc_00E16E20: push 00000028h
  loc_00E16E22: push 006E09E8h
  loc_00E16E27: push edx
  loc_00E16E28: push eax
  loc_00E16E29: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16E2F: mov eax, var_34
  loc_00E16E32: mov ecx, [esi]
  loc_00E16E34: push esi
  loc_00E16E35: mov var_94, eax
  loc_00E16E3B: call [ecx+00000368h]
  loc_00E16E41: lea edx, var_24
  loc_00E16E44: push eax
  loc_00E16E45: push edx
  loc_00E16E46: call edi
  loc_00E16E48: mov ebx, eax
  loc_00E16E4A: lea ecx, var_18
  loc_00E16E4D: push ecx
  loc_00E16E4E: push ebx
  loc_00E16E4F: mov eax, [ebx]
  loc_00E16E51: call [eax+000000A0h]
  loc_00E16E57: test eax, eax
  loc_00E16E59: fnclex
  loc_00E16E5B: jge 00E16E6Fh
  loc_00E16E5D: push 000000A0h
  loc_00E16E62: push 006DCB70h
  loc_00E16E67: push ebx
  loc_00E16E68: push eax
  loc_00E16E69: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16E6F: sub esp, 00000010h
  loc_00E16E72: mov eax, var_18
  loc_00E16E75: mov ebx, esp
  loc_00E16E77: mov ecx, 00000008h
  loc_00E16E7C: mov var_54, ecx
  loc_00E16E7F: mov edx, var_94
  loc_00E16E85: mov [ebx], ecx
  loc_00E16E87: mov ecx, var_50
  loc_00E16E8A: mov var_4C, eax
  loc_00E16E8D: mov var_18, 00000000h
  loc_00E16E94: mov edx, [edx]
  loc_00E16E96: mov [ebx+00000004h], ecx
  loc_00E16E99: mov [ebx+00000008h], eax
  loc_00E16E9C: mov eax, var_48
  loc_00E16E9F: mov [ebx+0000000Ch], eax
  loc_00E16EA2: mov ebx, var_94
  loc_00E16EA8: push ebx
  loc_00E16EA9: call [edx+00000038h]
  loc_00E16EAC: test eax, eax
  loc_00E16EAE: fnclex
  loc_00E16EB0: jge 00E16EC1h
  loc_00E16EB2: push 00000038h
  loc_00E16EB4: push 006E09F8h
  loc_00E16EB9: push ebx
  loc_00E16EBA: push eax
  loc_00E16EBB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16EC1: lea ecx, var_34
  loc_00E16EC4: lea edx, var_30
  loc_00E16EC7: push ecx
  loc_00E16EC8: lea eax, var_2C
  loc_00E16ECB: push edx
  loc_00E16ECC: lea ecx, var_28
  loc_00E16ECF: push eax
  loc_00E16ED0: lea edx, var_24
  loc_00E16ED3: push ecx
  loc_00E16ED4: push edx
  loc_00E16ED5: push 00000005h
  loc_00E16ED7: call [00401048h] ; __vbaFreeObjList
  loc_00E16EDD: lea eax, var_54
  loc_00E16EE0: lea ecx, var_44
  loc_00E16EE3: push eax
  loc_00E16EE4: push ecx
  loc_00E16EE5: push 00000002h
  loc_00E16EE7: call [00401038h] ; __vbaFreeVarList
  loc_00E16EED: mov edx, [esi]
  loc_00E16EEF: add esp, 00000024h
  loc_00E16EF2: push 006DCBD8h
  loc_00E16EF7: push 00000000h
  loc_00E16EF9: push 00000018h
  loc_00E16EFB: push esi
  loc_00E16EFC: call [edx+00000490h]
  loc_00E16F02: push eax
  loc_00E16F03: lea eax, var_28
  loc_00E16F06: push eax
  loc_00E16F07: call edi
  loc_00E16F09: lea ecx, var_44
  loc_00E16F0C: push eax
  loc_00E16F0D: push ecx
  loc_00E16F0E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E16F14: add esp, 00000010h
  loc_00E16F17: push eax
  loc_00E16F18: call [00401120h] ; __vbaCastObjVar
  loc_00E16F1E: lea edx, var_2C
  loc_00E16F21: push eax
  loc_00E16F22: push edx
  loc_00E16F23: call edi
  loc_00E16F25: mov ebx, eax
  loc_00E16F27: lea ecx, var_30
  loc_00E16F2A: push ecx
  loc_00E16F2B: push ebx
  loc_00E16F2C: mov eax, [ebx]
  loc_00E16F2E: call [eax+00000054h]
  loc_00E16F31: test eax, eax
  loc_00E16F33: fnclex
  loc_00E16F35: jge 00E16F46h
  loc_00E16F37: push 00000054h
  loc_00E16F39: push 006DCBE8h
  loc_00E16F3E: push ebx
  loc_00E16F3F: push eax
  loc_00E16F40: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16F46: lea ebx, var_34
  loc_00E16F49: mov eax, var_30
  loc_00E16F4C: push ebx
  loc_00E16F4D: mov ecx, 00000002h
  loc_00E16F52: sub esp, 00000010h
  loc_00E16F55: mov edx, [eax]
  loc_00E16F57: mov ebx, esp
  loc_00E16F59: mov var_8C, eax
  loc_00E16F5F: push eax
  loc_00E16F60: mov [ebx], ecx
  loc_00E16F62: mov ecx, var_60
  loc_00E16F65: mov [ebx+00000004h], ecx
  loc_00E16F68: mov ecx, 0000000Dh
  loc_00E16F6D: mov [ebx+00000008h], ecx
  loc_00E16F70: mov ecx, var_58
  loc_00E16F73: mov [ebx+0000000Ch], ecx
  loc_00E16F76: call [edx+00000028h]
  loc_00E16F79: test eax, eax
  loc_00E16F7B: fnclex
  loc_00E16F7D: jge 00E16F94h
  loc_00E16F7F: mov edx, var_8C
  loc_00E16F85: push 00000028h
  loc_00E16F87: push 006E09E8h
  loc_00E16F8C: push edx
  loc_00E16F8D: push eax
  loc_00E16F8E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16F94: mov eax, var_34
  loc_00E16F97: mov ecx, [esi]
  loc_00E16F99: push esi
  loc_00E16F9A: mov var_94, eax
  loc_00E16FA0: call [ecx+00000364h]
  loc_00E16FA6: lea edx, var_24
  loc_00E16FA9: push eax
  loc_00E16FAA: push edx
  loc_00E16FAB: call edi
  loc_00E16FAD: mov ebx, eax
  loc_00E16FAF: lea ecx, var_18
  loc_00E16FB2: push ecx
  loc_00E16FB3: push ebx
  loc_00E16FB4: mov eax, [ebx]
  loc_00E16FB6: call [eax+000000A0h]
  loc_00E16FBC: test eax, eax
  loc_00E16FBE: fnclex
  loc_00E16FC0: jge 00E16FD4h
  loc_00E16FC2: push 000000A0h
  loc_00E16FC7: push 006DCB70h
  loc_00E16FCC: push ebx
  loc_00E16FCD: push eax
  loc_00E16FCE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E16FD4: sub esp, 00000010h
  loc_00E16FD7: mov eax, var_18
  loc_00E16FDA: mov ebx, esp
  loc_00E16FDC: mov ecx, 00000008h
  loc_00E16FE1: mov var_54, ecx
  loc_00E16FE4: mov edx, var_94
  loc_00E16FEA: mov [ebx], ecx
  loc_00E16FEC: mov ecx, var_50
  loc_00E16FEF: mov var_4C, eax
  loc_00E16FF2: mov var_18, 00000000h
  loc_00E16FF9: mov edx, [edx]
  loc_00E16FFB: mov [ebx+00000004h], ecx
  loc_00E16FFE: mov [ebx+00000008h], eax
  loc_00E17001: mov eax, var_48
  loc_00E17004: mov [ebx+0000000Ch], eax
  loc_00E17007: mov ebx, var_94
  loc_00E1700D: push ebx
  loc_00E1700E: call [edx+00000038h]
  loc_00E17011: test eax, eax
  loc_00E17013: fnclex
  loc_00E17015: jge 00E17026h
  loc_00E17017: push 00000038h
  loc_00E17019: push 006E09F8h
  loc_00E1701E: push ebx
  loc_00E1701F: push eax
  loc_00E17020: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E17026: lea ecx, var_34
  loc_00E17029: lea edx, var_30
  loc_00E1702C: push ecx
  loc_00E1702D: lea eax, var_2C
  loc_00E17030: push edx
  loc_00E17031: lea ecx, var_28
  loc_00E17034: push eax
  loc_00E17035: lea edx, var_24
  loc_00E17038: push ecx
  loc_00E17039: push edx
  loc_00E1703A: push 00000005h
  loc_00E1703C: call [00401048h] ; __vbaFreeObjList
  loc_00E17042: lea eax, var_54
  loc_00E17045: lea ecx, var_44
  loc_00E17048: push eax
  loc_00E17049: push ecx
  loc_00E1704A: push 00000002h
  loc_00E1704C: call [00401038h] ; __vbaFreeVarList
  loc_00E17052: mov edx, [esi]
  loc_00E17054: add esp, 00000024h
  loc_00E17057: push 006DCBD8h
  loc_00E1705C: push 00000000h
  loc_00E1705E: push 00000018h
  loc_00E17060: push esi
  loc_00E17061: call [edx+00000490h]
  loc_00E17067: push eax
  loc_00E17068: lea eax, var_28
  loc_00E1706B: push eax
  loc_00E1706C: call edi
  loc_00E1706E: lea ecx, var_44
  loc_00E17071: push eax
  loc_00E17072: push ecx
  loc_00E17073: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E17079: add esp, 00000010h
  loc_00E1707C: push eax
  loc_00E1707D: call [00401120h] ; __vbaCastObjVar
  loc_00E17083: lea edx, var_2C
  loc_00E17086: push eax
  loc_00E17087: push edx
  loc_00E17088: call edi
  loc_00E1708A: mov ebx, eax
  loc_00E1708C: lea ecx, var_30
  loc_00E1708F: push ecx
  loc_00E17090: push ebx
  loc_00E17091: mov eax, [ebx]
  loc_00E17093: call [eax+00000054h]
  loc_00E17096: test eax, eax
  loc_00E17098: fnclex
  loc_00E1709A: jge 00E170ABh
  loc_00E1709C: push 00000054h
  loc_00E1709E: push 006DCBE8h
  loc_00E170A3: push ebx
  loc_00E170A4: push eax
  loc_00E170A5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E170AB: lea ebx, var_34
  loc_00E170AE: mov eax, var_30
  loc_00E170B1: push ebx
  loc_00E170B2: mov ecx, 00000002h
  loc_00E170B7: sub esp, 00000010h
  loc_00E170BA: mov edx, [eax]
  loc_00E170BC: mov ebx, esp
  loc_00E170BE: mov var_8C, eax
  loc_00E170C4: push eax
  loc_00E170C5: mov [ebx], ecx
  loc_00E170C7: mov ecx, var_60
  loc_00E170CA: mov [ebx+00000004h], ecx
  loc_00E170CD: mov ecx, 0000000Eh
  loc_00E170D2: mov [ebx+00000008h], ecx
  loc_00E170D5: mov ecx, var_58
  loc_00E170D8: mov [ebx+0000000Ch], ecx
  loc_00E170DB: call [edx+00000028h]
  loc_00E170DE: test eax, eax
  loc_00E170E0: fnclex
  loc_00E170E2: jge 00E170F9h
  loc_00E170E4: mov edx, var_8C
  loc_00E170EA: push 00000028h
  loc_00E170EC: push 006E09E8h
  loc_00E170F1: push edx
  loc_00E170F2: push eax
  loc_00E170F3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E170F9: mov eax, var_34
  loc_00E170FC: mov ecx, [esi]
  loc_00E170FE: push esi
  loc_00E170FF: mov var_94, eax
  loc_00E17105: call [ecx+00000360h]
  loc_00E1710B: lea edx, var_24
  loc_00E1710E: push eax
  loc_00E1710F: push edx
  loc_00E17110: call edi
  loc_00E17112: mov ebx, eax
  loc_00E17114: lea ecx, var_18
  loc_00E17117: push ecx
  loc_00E17118: push ebx
  loc_00E17119: mov eax, [ebx]
  loc_00E1711B: call [eax+000000A0h]
  loc_00E17121: test eax, eax
  loc_00E17123: fnclex
  loc_00E17125: jge 00E17139h
  loc_00E17127: push 000000A0h
  loc_00E1712C: push 006DCB70h
  loc_00E17131: push ebx
  loc_00E17132: push eax
  loc_00E17133: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E17139: sub esp, 00000010h
  loc_00E1713C: mov eax, var_18
  loc_00E1713F: mov ebx, esp
  loc_00E17141: mov ecx, 00000008h
  loc_00E17146: mov var_54, ecx
  loc_00E17149: mov edx, var_94
  loc_00E1714F: mov [ebx], ecx
  loc_00E17151: mov ecx, var_50
  loc_00E17154: mov var_4C, eax
  loc_00E17157: mov var_18, 00000000h
  loc_00E1715E: mov edx, [edx]
  loc_00E17160: mov [ebx+00000004h], ecx
  loc_00E17163: mov [ebx+00000008h], eax
  loc_00E17166: mov eax, var_48
  loc_00E17169: mov [ebx+0000000Ch], eax
  loc_00E1716C: mov ebx, var_94
  loc_00E17172: push ebx
  loc_00E17173: call [edx+00000038h]
  loc_00E17176: test eax, eax
  loc_00E17178: fnclex
  loc_00E1717A: jge 00E1718Bh
  loc_00E1717C: push 00000038h
  loc_00E1717E: push 006E09F8h
  loc_00E17183: push ebx
  loc_00E17184: push eax
  loc_00E17185: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1718B: lea ecx, var_34
  loc_00E1718E: lea edx, var_30
  loc_00E17191: push ecx
  loc_00E17192: lea eax, var_2C
  loc_00E17195: push edx
  loc_00E17196: lea ecx, var_28
  loc_00E17199: push eax
  loc_00E1719A: lea edx, var_24
  loc_00E1719D: push ecx
  loc_00E1719E: push edx
  loc_00E1719F: push 00000005h
  loc_00E171A1: call [00401048h] ; __vbaFreeObjList
  loc_00E171A7: lea eax, var_54
  loc_00E171AA: lea ecx, var_44
  loc_00E171AD: push eax
  loc_00E171AE: push ecx
  loc_00E171AF: push 00000002h
  loc_00E171B1: call [00401038h] ; __vbaFreeVarList
  loc_00E171B7: mov edx, [esi]
  loc_00E171B9: add esp, 00000024h
  loc_00E171BC: push 006DCBD8h
  loc_00E171C1: push 00000000h
  loc_00E171C3: push 00000018h
  loc_00E171C5: push esi
  loc_00E171C6: call [edx+00000490h]
  loc_00E171CC: push eax
  loc_00E171CD: lea eax, var_28
  loc_00E171D0: push eax
  loc_00E171D1: call edi
  loc_00E171D3: lea ecx, var_44
  loc_00E171D6: push eax
  loc_00E171D7: push ecx
  loc_00E171D8: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E171DE: add esp, 00000010h
  loc_00E171E1: push eax
  loc_00E171E2: call [00401120h] ; __vbaCastObjVar
  loc_00E171E8: lea edx, var_2C
  loc_00E171EB: push eax
  loc_00E171EC: push edx
  loc_00E171ED: call edi
  loc_00E171EF: mov ebx, eax
  loc_00E171F1: lea ecx, var_30
  loc_00E171F4: push ecx
  loc_00E171F5: push ebx
  loc_00E171F6: mov eax, [ebx]
  loc_00E171F8: call [eax+00000054h]
  loc_00E171FB: test eax, eax
  loc_00E171FD: fnclex
  loc_00E171FF: jge 00E17210h
  loc_00E17201: push 00000054h
  loc_00E17203: push 006DCBE8h
  loc_00E17208: push ebx
  loc_00E17209: push eax
  loc_00E1720A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E17210: lea ebx, var_34
  loc_00E17213: mov eax, var_30
  loc_00E17216: push ebx
  loc_00E17217: mov ecx, 00000002h
  loc_00E1721C: sub esp, 00000010h
  loc_00E1721F: mov edx, [eax]
  loc_00E17221: mov ebx, esp
  loc_00E17223: mov var_8C, eax
  loc_00E17229: push eax
  loc_00E1722A: mov [ebx], ecx
  loc_00E1722C: mov ecx, var_60
  loc_00E1722F: mov [ebx+00000004h], ecx
  loc_00E17232: mov ecx, 00000013h
  loc_00E17237: mov [ebx+00000008h], ecx
  loc_00E1723A: mov ecx, var_58
  loc_00E1723D: mov [ebx+0000000Ch], ecx
  loc_00E17240: call [edx+00000028h]
  loc_00E17243: test eax, eax
  loc_00E17245: fnclex
  loc_00E17247: jge 00E1725Eh
  loc_00E17249: mov edx, var_8C
  loc_00E1724F: push 00000028h
  loc_00E17251: push 006E09E8h
  loc_00E17256: push edx
  loc_00E17257: push eax
  loc_00E17258: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1725E: mov eax, var_34
  loc_00E17261: mov ecx, [esi]
  loc_00E17263: push esi
  loc_00E17264: mov var_94, eax
  loc_00E1726A: call [ecx+0000030Ch]
  loc_00E17270: lea edx, var_24
  loc_00E17273: push eax
  loc_00E17274: push edx
  loc_00E17275: call edi
  loc_00E17277: mov ebx, eax
  loc_00E17279: lea ecx, var_18
  loc_00E1727C: push ecx
  loc_00E1727D: push ebx
  loc_00E1727E: mov eax, [ebx]
  loc_00E17280: call [eax+000000A8h]
  loc_00E17286: test eax, eax
  loc_00E17288: fnclex
  loc_00E1728A: jge 00E1729Eh
  loc_00E1728C: push 000000A8h
  loc_00E17291: push 006E0400h
  loc_00E17296: push ebx
  loc_00E17297: push eax
  loc_00E17298: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1729E: sub esp, 00000010h
  loc_00E172A1: mov eax, var_18
  loc_00E172A4: mov ebx, esp
  loc_00E172A6: mov ecx, 00000008h
  loc_00E172AB: mov var_54, ecx
  loc_00E172AE: mov edx, var_94
  loc_00E172B4: mov [ebx], ecx
  loc_00E172B6: mov ecx, var_50
  loc_00E172B9: mov var_4C, eax
  loc_00E172BC: mov var_18, 00000000h
  loc_00E172C3: mov edx, [edx]
  loc_00E172C5: mov [ebx+00000004h], ecx
  loc_00E172C8: mov [ebx+00000008h], eax
  loc_00E172CB: mov eax, var_48
  loc_00E172CE: mov [ebx+0000000Ch], eax
  loc_00E172D1: mov ebx, var_94
  loc_00E172D7: push ebx
  loc_00E172D8: call [edx+00000038h]
  loc_00E172DB: test eax, eax
  loc_00E172DD: fnclex
  loc_00E172DF: jge 00E172F0h
  loc_00E172E1: push 00000038h
  loc_00E172E3: push 006E09F8h
  loc_00E172E8: push ebx
  loc_00E172E9: push eax
  loc_00E172EA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E172F0: lea ecx, var_34
  loc_00E172F3: lea edx, var_30
  loc_00E172F6: push ecx
  loc_00E172F7: lea eax, var_2C
  loc_00E172FA: push edx
  loc_00E172FB: lea ecx, var_28
  loc_00E172FE: push eax
  loc_00E172FF: lea edx, var_24
  loc_00E17302: push ecx
  loc_00E17303: push edx
  loc_00E17304: push 00000005h
  loc_00E17306: call [00401048h] ; __vbaFreeObjList
  loc_00E1730C: lea eax, var_54
  loc_00E1730F: lea ecx, var_44
  loc_00E17312: push eax
  loc_00E17313: push ecx
  loc_00E17314: push 00000002h
  loc_00E17316: call [00401038h] ; __vbaFreeVarList
  loc_00E1731C: mov edx, [esi]
  loc_00E1731E: add esp, 00000024h
  loc_00E17321: push 006DCBD8h
  loc_00E17326: push 00000000h
  loc_00E17328: push 00000018h
  loc_00E1732A: push esi
  loc_00E1732B: call [edx+00000490h]
  loc_00E17331: push eax
  loc_00E17332: lea eax, var_28
  loc_00E17335: push eax
  loc_00E17336: call edi
  loc_00E17338: lea ecx, var_44
  loc_00E1733B: push eax
  loc_00E1733C: push ecx
  loc_00E1733D: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E17343: add esp, 00000010h
  loc_00E17346: push eax
  loc_00E17347: call [00401120h] ; __vbaCastObjVar
  loc_00E1734D: lea edx, var_2C
  loc_00E17350: push eax
  loc_00E17351: push edx
  loc_00E17352: call edi
  loc_00E17354: mov ebx, eax
  loc_00E17356: lea ecx, var_30
  loc_00E17359: push ecx
  loc_00E1735A: push ebx
  loc_00E1735B: mov eax, [ebx]
  loc_00E1735D: call [eax+00000054h]
  loc_00E17360: test eax, eax
  loc_00E17362: fnclex
  loc_00E17364: jge 00E17375h
  loc_00E17366: push 00000054h
  loc_00E17368: push 006DCBE8h
  loc_00E1736D: push ebx
  loc_00E1736E: push eax
  loc_00E1736F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E17375: lea ebx, var_34
  loc_00E17378: mov eax, var_30
  loc_00E1737B: push ebx
  loc_00E1737C: mov ecx, 00000002h
  loc_00E17381: sub esp, 00000010h
  loc_00E17384: mov edx, [eax]
  loc_00E17386: mov ebx, esp
  loc_00E17388: mov var_8C, eax
  loc_00E1738E: push eax
  loc_00E1738F: mov [ebx], ecx
  loc_00E17391: mov ecx, var_60
  loc_00E17394: mov [ebx+00000004h], ecx
  loc_00E17397: mov ecx, 00000014h
  loc_00E1739C: mov [ebx+00000008h], ecx
  loc_00E1739F: mov ecx, var_58
  loc_00E173A2: mov [ebx+0000000Ch], ecx
  loc_00E173A5: call [edx+00000028h]
  loc_00E173A8: test eax, eax
  loc_00E173AA: fnclex
  loc_00E173AC: jge 00E173C3h
  loc_00E173AE: mov edx, var_8C
  loc_00E173B4: push 00000028h
  loc_00E173B6: push 006E09E8h
  loc_00E173BB: push edx
  loc_00E173BC: push eax
  loc_00E173BD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E173C3: mov eax, var_34
  loc_00E173C6: mov ecx, [esi]
  loc_00E173C8: push esi
  loc_00E173C9: mov var_94, eax
  loc_00E173CF: call [ecx+00000304h]
  loc_00E173D5: lea edx, var_24
  loc_00E173D8: push eax
  loc_00E173D9: push edx
  loc_00E173DA: call edi
  loc_00E173DC: mov ebx, eax
  loc_00E173DE: lea ecx, var_18
  loc_00E173E1: push ecx
  loc_00E173E2: push ebx
  loc_00E173E3: mov eax, [ebx]
  loc_00E173E5: call [eax+000000A0h]
  loc_00E173EB: test eax, eax
  loc_00E173ED: fnclex
  loc_00E173EF: jge 00E17403h
  loc_00E173F1: push 000000A0h
  loc_00E173F6: push 006DCB70h
  loc_00E173FB: push ebx
  loc_00E173FC: push eax
  loc_00E173FD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E17403: sub esp, 00000010h
  loc_00E17406: mov eax, var_18
  loc_00E17409: mov ebx, esp
  loc_00E1740B: mov ecx, 00000008h
  loc_00E17410: mov var_54, ecx
  loc_00E17413: mov edx, var_94
  loc_00E17419: mov [ebx], ecx
  loc_00E1741B: mov ecx, var_50
  loc_00E1741E: mov var_4C, eax
  loc_00E17421: mov var_18, 00000000h
  loc_00E17428: mov edx, [edx]
  loc_00E1742A: mov [ebx+00000004h], ecx
  loc_00E1742D: mov [ebx+00000008h], eax
  loc_00E17430: mov eax, var_48
  loc_00E17433: mov [ebx+0000000Ch], eax
  loc_00E17436: mov ebx, var_94
  loc_00E1743C: push ebx
  loc_00E1743D: call [edx+00000038h]
  loc_00E17440: test eax, eax
  loc_00E17442: fnclex
  loc_00E17444: jge 00E17455h
  loc_00E17446: push 00000038h
  loc_00E17448: push 006E09F8h
  loc_00E1744D: push ebx
  loc_00E1744E: push eax
  loc_00E1744F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E17455: lea ecx, var_34
  loc_00E17458: lea edx, var_30
  loc_00E1745B: push ecx
  loc_00E1745C: lea eax, var_2C
  loc_00E1745F: push edx
  loc_00E17460: lea ecx, var_28
  loc_00E17463: push eax
  loc_00E17464: lea edx, var_24
  loc_00E17467: push ecx
  loc_00E17468: push edx
  loc_00E17469: push 00000005h
  loc_00E1746B: call [00401048h] ; __vbaFreeObjList
  loc_00E17471: lea eax, var_54
  loc_00E17474: lea ecx, var_44
  loc_00E17477: push eax
  loc_00E17478: push ecx
  loc_00E17479: push 00000002h
  loc_00E1747B: call [00401038h] ; __vbaFreeVarList
  loc_00E17481: mov edx, [esi]
  loc_00E17483: add esp, 00000024h
  loc_00E17486: push 006DCBD8h
  loc_00E1748B: push 00000000h
  loc_00E1748D: push 00000018h
  loc_00E1748F: push esi
  loc_00E17490: call [edx+00000490h]
  loc_00E17496: push eax
  loc_00E17497: lea eax, var_24
  loc_00E1749A: push eax
  loc_00E1749B: call edi
  loc_00E1749D: lea ecx, var_44
  loc_00E174A0: push eax
  loc_00E174A1: push ecx
  loc_00E174A2: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E174A8: add esp, 00000010h
  loc_00E174AB: push eax
  loc_00E174AC: call [00401120h] ; __vbaCastObjVar
  loc_00E174B2: lea edx, var_28
  loc_00E174B5: push eax
  loc_00E174B6: push edx
  loc_00E174B7: call edi
  loc_00E174B9: sub esp, 00000010h
  loc_00E174BC: mov ebx, eax
  loc_00E174BE: mov edx, esp
  loc_00E174C0: mov eax, 0000000Ah
  loc_00E174C5: mov var_64, eax
  loc_00E174C8: sub esp, 00000010h
  loc_00E174CB: mov [edx], eax
  loc_00E174CD: mov eax, var_70
  loc_00E174D0: mov ecx, [ebx]
  loc_00E174D2: mov var_5C, 80020004h
  loc_00E174D9: mov [edx+00000004h], eax
  loc_00E174DC: mov eax, 80020004h
  loc_00E174E1: mov [edx+00000008h], eax
  loc_00E174E4: mov eax, var_68
  loc_00E174E7: mov [edx+0000000Ch], eax
  loc_00E174EA: mov eax, var_64
  loc_00E174ED: mov edx, esp
  loc_00E174EF: push ebx
  loc_00E174F0: mov [edx], eax
  loc_00E174F2: mov eax, var_60
  loc_00E174F5: mov [edx+00000004h], eax
  loc_00E174F8: mov eax, var_5C
  loc_00E174FB: mov [edx+00000008h], eax
  loc_00E174FE: mov eax, var_58
  loc_00E17501: mov [edx+0000000Ch], eax
  loc_00E17504: call [ecx+000000ACh]
  loc_00E1750A: test eax, eax
  loc_00E1750C: fnclex
  loc_00E1750E: jge 00E17522h
  loc_00E17510: push 000000ACh
  loc_00E17515: push 006DCBE8h
  loc_00E1751A: push ebx
  loc_00E1751B: push eax
  loc_00E1751C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E17522: lea ecx, var_28
  loc_00E17525: lea edx, var_24
  loc_00E17528: push ecx
  loc_00E17529: push edx
  loc_00E1752A: push 00000002h
  loc_00E1752C: call [00401048h] ; __vbaFreeObjList
  loc_00E17532: add esp, 0000000Ch
  loc_00E17535: lea ecx, var_44
  loc_00E17538: call [00401024h] ; __vbaFreeVar
  loc_00E1753E: mov eax, [esi]
  loc_00E17540: push 00000000h
  loc_00E17542: push 00000019h
  loc_00E17544: push esi
  loc_00E17545: call [eax+00000490h]
  loc_00E1754B: lea ecx, var_24
  loc_00E1754E: push eax
  loc_00E1754F: push ecx
  loc_00E17550: call edi
  loc_00E17552: push eax
  loc_00E17553: call [00401030h] ; __vbaLateIdCall
  loc_00E17559: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E1755F: add esp, 0000000Ch
  loc_00E17562: lea ecx, var_24
  loc_00E17565: call ebx
  loc_00E17567: mov edx, 006E0E40h ; " select * from biodata"
  loc_00E1756C: lea ecx, [esi+00000034h]
  loc_00E1756F: call [004011B0h] ; __vbaStrCopy
  loc_00E17575: sub esp, 00000010h
  loc_00E17578: mov eax, 00004008h
  loc_00E1757D: mov edx, esp
  loc_00E1757F: mov ecx, var_58
  loc_00E17582: push 0000000Eh
  loc_00E17584: push esi
  loc_00E17585: mov [edx], eax
  loc_00E17587: mov eax, var_60
  loc_00E1758A: mov [edx+00000004h], eax
  loc_00E1758D: lea eax, [esi+00000034h]
  loc_00E17590: mov [edx+00000008h], eax
  loc_00E17593: mov [edx+0000000Ch], ecx
  loc_00E17596: mov edx, [esi]
  loc_00E17598: call [edx+00000490h]
  loc_00E1759E: push eax
  loc_00E1759F: lea eax, var_24
  loc_00E175A2: push eax
  loc_00E175A3: call edi
  loc_00E175A5: push eax
  loc_00E175A6: call [00401238h] ; __vbaLateIdSt
  loc_00E175AC: lea ecx, var_24
  loc_00E175AF: call ebx
  loc_00E175B1: mov ecx, [esi]
  loc_00E175B3: push 00000000h
  loc_00E175B5: push 00000019h
  loc_00E175B7: push esi
  loc_00E175B8: call [ecx+00000490h]
  loc_00E175BE: lea edx, var_24
  loc_00E175C1: push eax
  loc_00E175C2: push edx
  loc_00E175C3: call edi
  loc_00E175C5: push eax
  loc_00E175C6: call [00401030h] ; __vbaLateIdCall
  loc_00E175CC: add esp, 0000000Ch
  loc_00E175CF: lea ecx, var_24
  loc_00E175D2: call ebx
  loc_00E175D4: mov eax, [esi]
  loc_00E175D6: push 006E05D8h
  loc_00E175DB: push esi
  loc_00E175DC: call [eax+00000490h]
  loc_00E175E2: lea ecx, var_24
  loc_00E175E5: push eax
  loc_00E175E6: push ecx
  loc_00E175E7: call edi
  loc_00E175E9: push eax
  loc_00E175EA: call [00401224h] ; __vbaCastObj
  loc_00E175F0: lea edx, var_44
  loc_00E175F3: mov var_3C, eax
  loc_00E175F6: push edx
  loc_00E175F7: mov var_44, 0000000Dh
  loc_00E175FE: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E17604: mov edx, [eax]
  loc_00E17606: sub esp, 00000010h
  loc_00E17609: mov ecx, esp
  loc_00E1760B: mov [ecx], edx
  loc_00E1760D: mov edx, [eax+00000004h]
  loc_00E17610: push 00000000h
  loc_00E17612: push 0000002Ah
  loc_00E17614: mov [ecx+00000004h], edx
  loc_00E17617: mov edx, [eax+00000008h]
  loc_00E1761A: push esi
  loc_00E1761B: mov eax, [eax+0000000Ch]
  loc_00E1761E: mov [ecx+00000008h], edx
  loc_00E17621: mov [ecx+0000000Ch], eax
  loc_00E17624: mov ecx, [esi]
  loc_00E17626: call [ecx+00000494h]
  loc_00E1762C: lea edx, var_28
  loc_00E1762F: push eax
  loc_00E17630: push edx
  loc_00E17631: call edi
  loc_00E17633: push eax
  loc_00E17634: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E1763A: lea eax, var_28
  loc_00E1763D: lea ecx, var_24
  loc_00E17640: push eax
  loc_00E17641: push ecx
  loc_00E17642: push 00000002h
  loc_00E17644: call [00401048h] ; __vbaFreeObjList
  loc_00E1764A: add esp, 00000028h
  loc_00E1764D: lea ecx, var_44
  loc_00E17650: call [00401024h] ; __vbaFreeVar
  loc_00E17656: mov edx, [esi]
  loc_00E17658: push 00000000h
  loc_00E1765A: push FFFFFDDAh
  loc_00E1765F: push esi
  loc_00E17660: call [edx+00000494h]
  loc_00E17666: push eax
  loc_00E17667: lea eax, var_24
  loc_00E1766A: push eax
  loc_00E1766B: call edi
  loc_00E1766D: push eax
  loc_00E1766E: call [00401030h] ; __vbaLateIdCall
  loc_00E17674: add esp, 0000000Ch
  loc_00E17677: lea ecx, var_24
  loc_00E1767A: call ebx
  loc_00E1767C: mov var_4, 00000000h
  loc_00E17683: push 00E176CFh
  loc_00E17688: jmp 00E176CEh
  loc_00E1768A: lea ecx, var_20
  loc_00E1768D: lea edx, var_1C
  loc_00E17690: push ecx
  loc_00E17691: lea eax, var_18
  loc_00E17694: push edx
  loc_00E17695: push eax
  loc_00E17696: push 00000003h
  loc_00E17698: call [004011BCh] ; __vbaFreeStrList
  loc_00E1769E: lea ecx, var_34
  loc_00E176A1: lea edx, var_30
  loc_00E176A4: push ecx
  loc_00E176A5: lea eax, var_2C
  loc_00E176A8: push edx
  loc_00E176A9: lea ecx, var_28
  loc_00E176AC: push eax
  loc_00E176AD: lea edx, var_24
  loc_00E176B0: push ecx
  loc_00E176B1: push edx
  loc_00E176B2: push 00000005h
  loc_00E176B4: call [00401048h] ; __vbaFreeObjList
  loc_00E176BA: lea eax, var_54
  loc_00E176BD: lea ecx, var_44
  loc_00E176C0: push eax
  loc_00E176C1: push ecx
  loc_00E176C2: push 00000002h
  loc_00E176C4: call [00401038h] ; __vbaFreeVarList
  loc_00E176CA: add esp, 00000034h
  loc_00E176CD: ret
  loc_00E176CE: ret
  loc_00E176CF: mov eax, Me
  loc_00E176D2: push eax
  loc_00E176D3: mov edx, [eax]
  loc_00E176D5: call [edx+00000008h]
  loc_00E176D8: mov eax, var_4
  loc_00E176DB: mov ecx, var_14
  loc_00E176DE: pop edi
  loc_00E176DF: pop esi
  loc_00E176E0: mov fs:[00000000h], ecx
  loc_00E176E7: pop ebx
  loc_00E176E8: mov esp, ebp
  loc_00E176EA: pop ebp
  loc_00E176EB: retn 0004h
End Sub

Private Sub tempat_KeyPress(KeyAscii As Integer) 'E09E30
  loc_00E09E30: push ebp
  loc_00E09E31: mov ebp, esp
  loc_00E09E33: sub esp, 0000000Ch
  loc_00E09E36: push 00402836h ; __vbaExceptHandler
  loc_00E09E3B: mov eax, fs:[00000000h]
  loc_00E09E41: push eax
  loc_00E09E42: mov fs:[00000000h], esp
  loc_00E09E49: sub esp, 00000014h
  loc_00E09E4C: push ebx
  loc_00E09E4D: push esi
  loc_00E09E4E: push edi
  loc_00E09E4F: mov var_C, esp
  loc_00E09E52: mov var_8, 00402180h
  loc_00E09E59: mov esi, Me
  loc_00E09E5C: mov eax, esi
  loc_00E09E5E: and eax, 00000001h
  loc_00E09E61: mov var_4, eax
  loc_00E09E64: and esi, FFFFFFFEh
  loc_00E09E67: push esi
  loc_00E09E68: mov Me, esi
  loc_00E09E6B: mov ecx, [esi]
  loc_00E09E6D: call [ecx+00000004h]
  loc_00E09E70: mov edx, KeyAscii
  loc_00E09E73: xor edi, edi
  loc_00E09E75: mov var_18, edi
  loc_00E09E78: cmp [edx], 000Dh
  loc_00E09E7C: jnz 00E09EBEh
  loc_00E09E7E: mov eax, [esi]
  loc_00E09E80: push esi
  loc_00E09E81: call [eax+00000384h]
  loc_00E09E87: lea ecx, var_18
  loc_00E09E8A: push eax
  loc_00E09E8B: push ecx
  loc_00E09E8C: call [004010ACh] ; __vbaObjSet
  loc_00E09E92: mov esi, eax
  loc_00E09E94: push esi
  loc_00E09E95: mov edx, [esi]
  loc_00E09E97: call [edx+00000204h]
  loc_00E09E9D: cmp eax, edi
  loc_00E09E9F: fnclex
  loc_00E09EA1: jge 00E09EB5h
  loc_00E09EA3: push 00000204h
  loc_00E09EA8: push 006DCB70h
  loc_00E09EAD: push esi
  loc_00E09EAE: push eax
  loc_00E09EAF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09EB5: lea ecx, var_18
  loc_00E09EB8: call [00401254h] ; __vbaFreeObj
  loc_00E09EBE: mov var_4, edi
  loc_00E09EC1: push 00E09ED3h
  loc_00E09EC6: jmp 00E09ED2h
  loc_00E09EC8: lea ecx, var_18
  loc_00E09ECB: call [00401254h] ; __vbaFreeObj
  loc_00E09ED1: ret
  loc_00E09ED2: ret
  loc_00E09ED3: mov eax, Me
  loc_00E09ED6: push eax
  loc_00E09ED7: mov ecx, [eax]
  loc_00E09ED9: call [ecx+00000008h]
  loc_00E09EDC: mov eax, var_4
  loc_00E09EDF: mov ecx, var_14
  loc_00E09EE2: pop edi
  loc_00E09EE3: pop esi
  loc_00E09EE4: mov fs:[00000000h], ecx
  loc_00E09EEB: pop ebx
  loc_00E09EEC: mov esp, ebp
  loc_00E09EEE: pop ebp
  loc_00E09EEF: retn 0008h
End Sub

Private Sub optper_Click() 'E0AA00
  loc_00E0AA00: push ebp
  loc_00E0AA01: mov ebp, esp
  loc_00E0AA03: sub esp, 0000000Ch
  loc_00E0AA06: push 00402836h ; __vbaExceptHandler
  loc_00E0AA0B: mov eax, fs:[00000000h]
  loc_00E0AA11: push eax
  loc_00E0AA12: mov fs:[00000000h], esp
  loc_00E0AA19: sub esp, 00000034h
  loc_00E0AA1C: push ebx
  loc_00E0AA1D: push esi
  loc_00E0AA1E: push edi
  loc_00E0AA1F: mov var_C, esp
  loc_00E0AA22: mov var_8, 00402220h
  loc_00E0AA29: mov esi, Me
  loc_00E0AA2C: mov eax, esi
  loc_00E0AA2E: and eax, 00000001h
  loc_00E0AA31: mov var_4, eax
  loc_00E0AA34: and esi, FFFFFFFEh
  loc_00E0AA37: push esi
  loc_00E0AA38: mov Me, esi
  loc_00E0AA3B: mov ecx, [esi]
  loc_00E0AA3D: call [ecx+00000004h]
  loc_00E0AA40: mov edx, [esi]
  loc_00E0AA42: push esi
  loc_00E0AA43: mov var_18, 00000000h
  loc_00E0AA4A: call [edx+00000380h]
  loc_00E0AA50: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0AA56: push eax
  loc_00E0AA57: lea eax, var_18
  loc_00E0AA5A: push eax
  loc_00E0AA5B: call ebx
  loc_00E0AA5D: mov edi, eax
  loc_00E0AA5F: push 00000000h
  loc_00E0AA61: push edi
  loc_00E0AA62: mov ecx, [edi]
  loc_00E0AA64: call [ecx+0000009Ch]
  loc_00E0AA6A: test eax, eax
  loc_00E0AA6C: fnclex
  loc_00E0AA6E: jge 00E0AA82h
  loc_00E0AA70: push 0000009Ch
  loc_00E0AA75: push 006E03D4h
  loc_00E0AA7A: push edi
  loc_00E0AA7B: push eax
  loc_00E0AA7C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0AA82: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E0AA88: lea ecx, var_18
  loc_00E0AA8B: call edi
  loc_00E0AA8D: sub esp, 00000010h
  loc_00E0AA90: mov ecx, 0000000Bh
  loc_00E0AA95: mov edx, esp
  loc_00E0AA97: or eax, FFFFFFFFh
  loc_00E0AA9A: push 80010007h
  loc_00E0AA9F: push esi
  loc_00E0AAA0: mov [edx], ecx
  loc_00E0AAA2: mov ecx, var_24
  loc_00E0AAA5: mov [edx+00000004h], ecx
  loc_00E0AAA8: mov ecx, [esi]
  loc_00E0AAAA: mov [edx+00000008h], eax
  loc_00E0AAAD: mov eax, var_1C
  loc_00E0AAB0: mov [edx+0000000Ch], eax
  loc_00E0AAB3: call [ecx+000003B8h]
  loc_00E0AAB9: lea edx, var_18
  loc_00E0AABC: push eax
  loc_00E0AABD: push edx
  loc_00E0AABE: call ebx
  loc_00E0AAC0: push eax
  loc_00E0AAC1: call [00401238h] ; __vbaLateIdSt
  loc_00E0AAC7: lea ecx, var_18
  loc_00E0AACA: call edi
  loc_00E0AACC: mov var_4, 00000000h
  loc_00E0AAD3: push 00E0AAE5h
  loc_00E0AAD8: jmp 00E0AAE4h
  loc_00E0AADA: lea ecx, var_18
  loc_00E0AADD: call [00401254h] ; __vbaFreeObj
  loc_00E0AAE3: ret
  loc_00E0AAE4: ret
  loc_00E0AAE5: mov eax, Me
  loc_00E0AAE8: push eax
  loc_00E0AAE9: mov ecx, [eax]
  loc_00E0AAEB: call [ecx+00000008h]
  loc_00E0AAEE: mov eax, var_4
  loc_00E0AAF1: mov ecx, var_14
  loc_00E0AAF4: pop edi
  loc_00E0AAF5: pop esi
  loc_00E0AAF6: mov fs:[00000000h], ecx
  loc_00E0AAFD: pop ebx
  loc_00E0AAFE: mov esp, ebp
  loc_00E0AB00: pop ebp
  loc_00E0AB01: retn 0004h
End Sub

Private Sub alamat_KeyPress(KeyAscii As Integer) 'E09F00
  loc_00E09F00: push ebp
  loc_00E09F01: mov ebp, esp
  loc_00E09F03: sub esp, 0000000Ch
  loc_00E09F06: push 00402836h ; __vbaExceptHandler
  loc_00E09F0B: mov eax, fs:[00000000h]
  loc_00E09F11: push eax
  loc_00E09F12: mov fs:[00000000h], esp
  loc_00E09F19: sub esp, 00000020h
  loc_00E09F1C: push ebx
  loc_00E09F1D: push esi
  loc_00E09F1E: push edi
  loc_00E09F1F: mov var_C, esp
  loc_00E09F22: mov var_8, 00402190h
  loc_00E09F29: mov edi, Me
  loc_00E09F2C: mov eax, edi
  loc_00E09F2E: and eax, 00000001h
  loc_00E09F31: mov var_4, eax
  loc_00E09F34: and edi, FFFFFFFEh
  loc_00E09F37: push edi
  loc_00E09F38: mov Me, edi
  loc_00E09F3B: mov ecx, [edi]
  loc_00E09F3D: call [ecx+00000004h]
  loc_00E09F40: mov edx, KeyAscii
  loc_00E09F43: xor ebx, ebx
  loc_00E09F45: mov var_18, ebx
  loc_00E09F48: mov var_1C, ebx
  loc_00E09F4B: cmp [edx], 000Dh
  loc_00E09F4F: jnz 00E09F95h
  loc_00E09F51: mov eax, [edi]
  loc_00E09F53: push edi
  loc_00E09F54: call [eax+00000374h]
  loc_00E09F5A: lea ecx, var_1C
  loc_00E09F5D: push eax
  loc_00E09F5E: push ecx
  loc_00E09F5F: call [004010ACh] ; __vbaObjSet
  loc_00E09F65: mov esi, eax
  loc_00E09F67: push esi
  loc_00E09F68: mov edx, [esi]
  loc_00E09F6A: call [edx+00000204h]
  loc_00E09F70: cmp eax, ebx
  loc_00E09F72: fnclex
  loc_00E09F74: jge 00E09F88h
  loc_00E09F76: push 00000204h
  loc_00E09F7B: push 006DCB70h
  loc_00E09F80: push esi
  loc_00E09F81: push eax
  loc_00E09F82: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09F88: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E09F8E: lea ecx, var_1C
  loc_00E09F91: call ebx
  loc_00E09F93: jmp 00E09F9Bh
  loc_00E09F95: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E09F9B: mov eax, [edi]
  loc_00E09F9D: push edi
  loc_00E09F9E: call [eax+00000384h]
  loc_00E09FA4: lea ecx, var_1C
  loc_00E09FA7: push eax
  loc_00E09FA8: push ecx
  loc_00E09FA9: call [004010ACh] ; __vbaObjSet
  loc_00E09FAF: mov esi, eax
  loc_00E09FB1: lea eax, var_18
  loc_00E09FB4: push eax
  loc_00E09FB5: push esi
  loc_00E09FB6: mov edx, [esi]
  loc_00E09FB8: call [edx+000000A0h]
  loc_00E09FBE: test eax, eax
  loc_00E09FC0: fnclex
  loc_00E09FC2: jge 00E09FD6h
  loc_00E09FC4: push 000000A0h
  loc_00E09FC9: push 006DCB70h
  loc_00E09FCE: push esi
  loc_00E09FCF: push eax
  loc_00E09FD0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09FD6: mov ecx, var_18
  loc_00E09FD9: push ecx
  loc_00E09FDA: push 006E0708h ; "Keluar"
  loc_00E09FDF: call [00401104h] ; __vbaStrCmp
  loc_00E09FE5: mov esi, eax
  loc_00E09FE7: lea ecx, var_18
  loc_00E09FEA: neg esi
  loc_00E09FEC: sbb esi, esi
  loc_00E09FEE: inc esi
  loc_00E09FEF: neg esi
  loc_00E09FF1: call [00401258h] ; __vbaFreeStr
  loc_00E09FF7: lea ecx, var_1C
  loc_00E09FFA: call ebx
  loc_00E09FFC: test si, si
  loc_00E09FFF: jz 00E0A052h
  loc_00E0A001: mov eax, [00E3D9CCh]
  loc_00E0A006: test eax, eax
  loc_00E0A008: jnz 00E0A01Ah
  loc_00E0A00A: push 00E3D9CCh
  loc_00E0A00F: push 006DC960h
  loc_00E0A014: call [004011A0h] ; __vbaNew2
  loc_00E0A01A: mov esi, [00E3D9CCh]
  loc_00E0A020: lea eax, var_1C
  loc_00E0A023: push edi
  loc_00E0A024: push eax
  loc_00E0A025: mov edx, [esi]
  loc_00E0A027: mov var_34, edx
  loc_00E0A02A: call [004010B4h] ; __vbaObjSetAddref
  loc_00E0A030: mov ecx, var_34
  loc_00E0A033: push eax
  loc_00E0A034: push esi
  loc_00E0A035: call [ecx+00000010h]
  loc_00E0A038: test eax, eax
  loc_00E0A03A: fnclex
  loc_00E0A03C: jge 00E0A04Dh
  loc_00E0A03E: push 00000010h
  loc_00E0A040: push 006DC950h
  loc_00E0A045: push esi
  loc_00E0A046: push eax
  loc_00E0A047: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A04D: lea ecx, var_1C
  loc_00E0A050: call ebx
  loc_00E0A052: mov var_4, 00000000h
  loc_00E0A059: push 00E0A074h
  loc_00E0A05E: jmp 00E0A073h
  loc_00E0A060: lea ecx, var_18
  loc_00E0A063: call [00401258h] ; __vbaFreeStr
  loc_00E0A069: lea ecx, var_1C
  loc_00E0A06C: call [00401254h] ; __vbaFreeObj
  loc_00E0A072: ret
  loc_00E0A073: ret
  loc_00E0A074: mov eax, Me
  loc_00E0A077: push eax
  loc_00E0A078: mov edx, [eax]
  loc_00E0A07A: call [edx+00000008h]
  loc_00E0A07D: mov eax, var_4
  loc_00E0A080: mov ecx, var_14
  loc_00E0A083: pop edi
  loc_00E0A084: pop esi
  loc_00E0A085: mov fs:[00000000h], ecx
  loc_00E0A08C: pop ebx
  loc_00E0A08D: mov esp, ebp
  loc_00E0A08F: pop ebp
  loc_00E0A090: retn 0008h
End Sub

Private Sub nayah_KeyPress(KeyAscii As Integer) 'E0A0A0
  loc_00E0A0A0: push ebp
  loc_00E0A0A1: mov ebp, esp
  loc_00E0A0A3: sub esp, 0000000Ch
  loc_00E0A0A6: push 00402836h ; __vbaExceptHandler
  loc_00E0A0AB: mov eax, fs:[00000000h]
  loc_00E0A0B1: push eax
  loc_00E0A0B2: mov fs:[00000000h], esp
  loc_00E0A0B9: sub esp, 00000014h
  loc_00E0A0BC: push ebx
  loc_00E0A0BD: push esi
  loc_00E0A0BE: push edi
  loc_00E0A0BF: mov var_C, esp
  loc_00E0A0C2: mov var_8, 004021A0h
  loc_00E0A0C9: mov esi, Me
  loc_00E0A0CC: mov eax, esi
  loc_00E0A0CE: and eax, 00000001h
  loc_00E0A0D1: mov var_4, eax
  loc_00E0A0D4: and esi, FFFFFFFEh
  loc_00E0A0D7: push esi
  loc_00E0A0D8: mov Me, esi
  loc_00E0A0DB: mov ecx, [esi]
  loc_00E0A0DD: call [ecx+00000004h]
  loc_00E0A0E0: mov edx, KeyAscii
  loc_00E0A0E3: xor edi, edi
  loc_00E0A0E5: mov var_18, edi
  loc_00E0A0E8: cmp [edx], 000Dh
  loc_00E0A0EC: jnz 00E0A12Eh
  loc_00E0A0EE: mov eax, [esi]
  loc_00E0A0F0: push esi
  loc_00E0A0F1: call [eax+00000370h]
  loc_00E0A0F7: lea ecx, var_18
  loc_00E0A0FA: push eax
  loc_00E0A0FB: push ecx
  loc_00E0A0FC: call [004010ACh] ; __vbaObjSet
  loc_00E0A102: mov esi, eax
  loc_00E0A104: push esi
  loc_00E0A105: mov edx, [esi]
  loc_00E0A107: call [edx+00000204h]
  loc_00E0A10D: cmp eax, edi
  loc_00E0A10F: fnclex
  loc_00E0A111: jge 00E0A125h
  loc_00E0A113: push 00000204h
  loc_00E0A118: push 006DCB70h
  loc_00E0A11D: push esi
  loc_00E0A11E: push eax
  loc_00E0A11F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A125: lea ecx, var_18
  loc_00E0A128: call [00401254h] ; __vbaFreeObj
  loc_00E0A12E: mov var_4, edi
  loc_00E0A131: push 00E0A143h
  loc_00E0A136: jmp 00E0A142h
  loc_00E0A138: lea ecx, var_18
  loc_00E0A13B: call [00401254h] ; __vbaFreeObj
  loc_00E0A141: ret
  loc_00E0A142: ret
  loc_00E0A143: mov eax, Me
  loc_00E0A146: push eax
  loc_00E0A147: mov ecx, [eax]
  loc_00E0A149: call [ecx+00000008h]
  loc_00E0A14C: mov eax, var_4
  loc_00E0A14F: mov ecx, var_14
  loc_00E0A152: pop edi
  loc_00E0A153: pop esi
  loc_00E0A154: mov fs:[00000000h], ecx
  loc_00E0A15B: pop ebx
  loc_00E0A15C: mov esp, ebp
  loc_00E0A15E: pop ebp
  loc_00E0A15F: retn 0008h
End Sub

Private Sub cari_UnknownEvent_9 'E08290
  loc_00E08290: push ebp
  loc_00E08291: mov ebp, esp
  loc_00E08293: sub esp, 0000000Ch
  loc_00E08296: push 00402836h ; __vbaExceptHandler
  loc_00E0829B: mov eax, fs:[00000000h]
  loc_00E082A1: push eax
  loc_00E082A2: mov fs:[00000000h], esp
  loc_00E082A9: sub esp, 00000014h
  loc_00E082AC: push ebx
  loc_00E082AD: push esi
  loc_00E082AE: push edi
  loc_00E082AF: mov var_C, esp
  loc_00E082B2: mov var_8, 00402090h
  loc_00E082B9: mov esi, Me
  loc_00E082BC: mov eax, esi
  loc_00E082BE: and eax, 00000001h
  loc_00E082C1: mov var_4, eax
  loc_00E082C4: and esi, FFFFFFFEh
  loc_00E082C7: push esi
  loc_00E082C8: mov Me, esi
  loc_00E082CB: mov ecx, [esi]
  loc_00E082CD: call [ecx+00000004h]
  loc_00E082D0: mov edx, [esi]
  loc_00E082D2: push esi
  loc_00E082D3: mov var_18, 00000000h
  loc_00E082DA: call [edx+000003E0h]
  loc_00E082E0: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E082E6: push eax
  loc_00E082E7: lea eax, var_18
  loc_00E082EA: push eax
  loc_00E082EB: call ebx
  loc_00E082ED: mov edi, eax
  loc_00E082EF: push FFFFFFFFh
  loc_00E082F1: push edi
  loc_00E082F2: mov ecx, [edi]
  loc_00E082F4: call [ecx+0000008Ch]
  loc_00E082FA: test eax, eax
  loc_00E082FC: fnclex
  loc_00E082FE: jge 00E08312h
  loc_00E08300: push 0000008Ch
  loc_00E08305: push 006DCDA0h
  loc_00E0830A: push edi
  loc_00E0830B: push eax
  loc_00E0830C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08312: lea ecx, var_18
  loc_00E08315: call [00401254h] ; __vbaFreeObj
  loc_00E0831B: mov edx, [esi]
  loc_00E0831D: push esi
  loc_00E0831E: call [edx+00000310h]
  loc_00E08324: push eax
  loc_00E08325: lea eax, var_18
  loc_00E08328: push eax
  loc_00E08329: call ebx
  loc_00E0832B: mov edi, eax
  loc_00E0832D: push FFFFFFFFh
  loc_00E0832F: push edi
  loc_00E08330: mov ecx, [edi]
  loc_00E08332: call [ecx+0000009Ch]
  loc_00E08338: test eax, eax
  loc_00E0833A: fnclex
  loc_00E0833C: jge 00E08350h
  loc_00E0833E: push 0000009Ch
  loc_00E08343: push 006DCAD0h
  loc_00E08348: push edi
  loc_00E08349: push eax
  loc_00E0834A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08350: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E08356: lea ecx, var_18
  loc_00E08359: call edi
  loc_00E0835B: mov edx, [esi]
  loc_00E0835D: push esi
  loc_00E0835E: call [edx+000003DCh]
  loc_00E08364: push eax
  loc_00E08365: lea eax, var_18
  loc_00E08368: push eax
  loc_00E08369: call ebx
  loc_00E0836B: mov esi, eax
  loc_00E0836D: push 00000000h
  loc_00E0836F: push esi
  loc_00E08370: mov ecx, [esi]
  loc_00E08372: call [ecx+0000008Ch]
  loc_00E08378: test eax, eax
  loc_00E0837A: fnclex
  loc_00E0837C: jge 00E08390h
  loc_00E0837E: push 0000008Ch
  loc_00E08383: push 006DCDA0h
  loc_00E08388: push esi
  loc_00E08389: push eax
  loc_00E0838A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08390: lea ecx, var_18
  loc_00E08393: call edi
  loc_00E08395: mov var_4, 00000000h
  loc_00E0839C: push 00E083AEh
  loc_00E083A1: jmp 00E083ADh
  loc_00E083A3: lea ecx, var_18
  loc_00E083A6: call [00401254h] ; __vbaFreeObj
  loc_00E083AC: ret
  loc_00E083AD: ret
  loc_00E083AE: mov eax, Me
  loc_00E083B1: push eax
  loc_00E083B2: mov edx, [eax]
  loc_00E083B4: call [edx+00000008h]
  loc_00E083B7: mov eax, var_4
  loc_00E083BA: mov ecx, var_14
  loc_00E083BD: pop edi
  loc_00E083BE: pop esi
  loc_00E083BF: mov fs:[00000000h], ecx
  loc_00E083C6: pop ebx
  loc_00E083C7: mov esp, ebp
  loc_00E083C9: pop ebp
  loc_00E083CA: retn 0004h
End Sub

Private Sub nibu_KeyPress(KeyAscii As Integer) 'E0A170
  loc_00E0A170: push ebp
  loc_00E0A171: mov ebp, esp
  loc_00E0A173: sub esp, 0000000Ch
  loc_00E0A176: push 00402836h ; __vbaExceptHandler
  loc_00E0A17B: mov eax, fs:[00000000h]
  loc_00E0A181: push eax
  loc_00E0A182: mov fs:[00000000h], esp
  loc_00E0A189: sub esp, 000000C4h
  loc_00E0A18F: push ebx
  loc_00E0A190: push esi
  loc_00E0A191: push edi
  loc_00E0A192: mov var_C, esp
  loc_00E0A195: mov var_8, 004021B0h
  loc_00E0A19C: mov esi, Me
  loc_00E0A19F: mov eax, esi
  loc_00E0A1A1: and eax, 00000001h
  loc_00E0A1A4: mov var_4, eax
  loc_00E0A1A7: and esi, FFFFFFFEh
  loc_00E0A1AA: push esi
  loc_00E0A1AB: mov Me, esi
  loc_00E0A1AE: mov ecx, [esi]
  loc_00E0A1B0: call [ecx+00000004h]
  loc_00E0A1B3: mov edx, KeyAscii
  loc_00E0A1B6: xor eax, eax
  loc_00E0A1B8: mov var_24, eax
  loc_00E0A1BB: mov var_28, eax
  loc_00E0A1BE: cmp [edx], 000Dh
  loc_00E0A1C2: mov var_2C, eax
  loc_00E0A1C5: mov var_30, eax
  loc_00E0A1C8: mov var_40, eax
  loc_00E0A1CB: mov var_50, eax
  loc_00E0A1CE: mov var_60, eax
  loc_00E0A1D1: mov var_70, eax
  loc_00E0A1D4: mov var_80, eax
  loc_00E0A1D7: mov var_90, eax
  loc_00E0A1DD: mov var_C0, eax
  loc_00E0A1E3: jnz 00E0A3E8h
  loc_00E0A1E9: mov eax, [esi]
  loc_00E0A1EB: push esi
  loc_00E0A1EC: call [eax+00000368h]
  loc_00E0A1F2: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0A1F8: lea ecx, var_2C
  loc_00E0A1FB: push eax
  loc_00E0A1FC: push ecx
  loc_00E0A1FD: call ebx
  loc_00E0A1FF: mov edi, eax
  loc_00E0A201: push edi
  loc_00E0A202: mov edx, [edi]
  loc_00E0A204: call [edx+00000204h]
  loc_00E0A20A: test eax, eax
  loc_00E0A20C: fnclex
  loc_00E0A20E: jge 00E0A222h
  loc_00E0A210: push 00000204h
  loc_00E0A215: push 006DCB70h
  loc_00E0A21A: push edi
  loc_00E0A21B: push eax
  loc_00E0A21C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A222: lea ecx, var_2C
  loc_00E0A225: call [00401254h] ; __vbaFreeObj
  loc_00E0A22B: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E0A231: mov ecx, 80020004h
  loc_00E0A236: mov var_68, ecx
  loc_00E0A239: mov eax, 0000000Ah
  loc_00E0A23E: mov var_58, ecx
  loc_00E0A241: lea edx, var_90
  loc_00E0A247: lea ecx, var_50
  loc_00E0A24A: mov var_70, eax
  loc_00E0A24D: mov var_60, eax
  loc_00E0A250: mov var_88, 006DC934h ; "Alert"
  loc_00E0A25A: mov var_90, 00000008h
  loc_00E0A264: call edi
  loc_00E0A266: lea edx, var_80
  loc_00E0A269: lea ecx, var_40
  loc_00E0A26C: mov var_78, 006E071Ch ; "Apakah Calon Peserta Didik tinggal Bersama Orang Tua ?"
  loc_00E0A273: mov var_80, 00000008h
  loc_00E0A27A: call edi
  loc_00E0A27C: lea eax, var_70
  loc_00E0A27F: lea ecx, var_60
  loc_00E0A282: push eax
  loc_00E0A283: lea edx, var_50
  loc_00E0A286: push ecx
  loc_00E0A287: push edx
  loc_00E0A288: lea eax, var_40
  loc_00E0A28B: push 00000044h
  loc_00E0A28D: push eax
  loc_00E0A28E: call [004010A8h] ; rtcMsgBox
  loc_00E0A294: lea edx, var_C0
  loc_00E0A29A: lea ecx, var_24
  loc_00E0A29D: mov var_B8, eax
  loc_00E0A2A3: mov var_C0, 00000003h
  loc_00E0A2AD: call [0040101Ch] ; __vbaVarMove
  loc_00E0A2B3: lea ecx, var_70
  loc_00E0A2B6: lea edx, var_60
  loc_00E0A2B9: push ecx
  loc_00E0A2BA: lea eax, var_50
  loc_00E0A2BD: push edx
  loc_00E0A2BE: lea ecx, var_40
  loc_00E0A2C1: push eax
  loc_00E0A2C2: push ecx
  loc_00E0A2C3: push 00000004h
  loc_00E0A2C5: call [00401038h] ; __vbaFreeVarList
  loc_00E0A2CB: add esp, 00000014h
  loc_00E0A2CE: lea edx, var_24
  loc_00E0A2D1: lea eax, var_80
  loc_00E0A2D4: mov var_78, 00000006h
  loc_00E0A2DB: push edx
  loc_00E0A2DC: push eax
  loc_00E0A2DD: mov var_80, 00008003h
  loc_00E0A2E4: call [00401108h] ; __vbaVarTstEq
  loc_00E0A2EA: test ax, ax
  loc_00E0A2ED: jz 00E0A3AAh
  loc_00E0A2F3: mov ecx, [esi]
  loc_00E0A2F5: push esi
  loc_00E0A2F6: call [ecx+0000036Ch]
  loc_00E0A2FC: lea edx, var_30
  loc_00E0A2FF: push eax
  loc_00E0A300: push edx
  loc_00E0A301: call ebx
  loc_00E0A303: mov var_CC, eax
  loc_00E0A309: mov eax, [esi]
  loc_00E0A30B: push esi
  loc_00E0A30C: call [eax+00000384h]
  loc_00E0A312: lea ecx, var_2C
  loc_00E0A315: push eax
  loc_00E0A316: push ecx
  loc_00E0A317: call ebx
  loc_00E0A319: mov edi, eax
  loc_00E0A31B: lea eax, var_28
  loc_00E0A31E: push eax
  loc_00E0A31F: push edi
  loc_00E0A320: mov edx, [edi]
  loc_00E0A322: call [edx+000000A0h]
  loc_00E0A328: test eax, eax
  loc_00E0A32A: fnclex
  loc_00E0A32C: jge 00E0A340h
  loc_00E0A32E: push 000000A0h
  loc_00E0A333: push 006DCB70h
  loc_00E0A338: push edi
  loc_00E0A339: push eax
  loc_00E0A33A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A340: mov edi, var_CC
  loc_00E0A346: mov edx, var_28
  loc_00E0A349: push edx
  loc_00E0A34A: push edi
  loc_00E0A34B: mov ecx, [edi]
  loc_00E0A34D: call [ecx+000000A4h]
  loc_00E0A353: test eax, eax
  loc_00E0A355: fnclex
  loc_00E0A357: jge 00E0A36Bh
  loc_00E0A359: push 000000A4h
  loc_00E0A35E: push 006DCB70h
  loc_00E0A363: push edi
  loc_00E0A364: push eax
  loc_00E0A365: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A36B: lea ecx, var_28
  loc_00E0A36E: call [00401258h] ; __vbaFreeStr
  loc_00E0A374: lea eax, var_30
  loc_00E0A377: lea ecx, var_2C
  loc_00E0A37A: push eax
  loc_00E0A37B: push ecx
  loc_00E0A37C: push 00000002h
  loc_00E0A37E: call [00401048h] ; __vbaFreeObjList
  loc_00E0A384: mov edx, [esi]
  loc_00E0A386: add esp, 0000000Ch
  loc_00E0A389: push esi
  loc_00E0A38A: call [edx+00000368h]
  loc_00E0A390: push eax
  loc_00E0A391: lea eax, var_2C
  loc_00E0A394: push eax
  loc_00E0A395: call ebx
  loc_00E0A397: mov esi, eax
  loc_00E0A399: push esi
  loc_00E0A39A: mov ecx, [esi]
  loc_00E0A39C: call [ecx+00000204h]
  loc_00E0A3A2: test eax, eax
  loc_00E0A3A4: fnclex
  loc_00E0A3A6: jge 00E0A3DDh
  loc_00E0A3A8: jmp 00E0A3CBh
  loc_00E0A3AA: mov edx, [esi]
  loc_00E0A3AC: push esi
  loc_00E0A3AD: call [edx+0000036Ch]
  loc_00E0A3B3: push eax
  loc_00E0A3B4: lea eax, var_2C
  loc_00E0A3B7: push eax
  loc_00E0A3B8: call ebx
  loc_00E0A3BA: mov esi, eax
  loc_00E0A3BC: push esi
  loc_00E0A3BD: mov ecx, [esi]
  loc_00E0A3BF: call [ecx+00000204h]
  loc_00E0A3C5: test eax, eax
  loc_00E0A3C7: fnclex
  loc_00E0A3C9: jge 00E0A3DDh
  loc_00E0A3CB: push 00000204h
  loc_00E0A3D0: push 006DCB70h
  loc_00E0A3D5: push esi
  loc_00E0A3D6: push eax
  loc_00E0A3D7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A3DD: lea ecx, var_2C
  loc_00E0A3E0: call [00401254h] ; __vbaFreeObj
  loc_00E0A3E6: xor eax, eax
  loc_00E0A3E8: mov var_4, eax
  loc_00E0A3EB: push 00E0A431h
  loc_00E0A3F0: jmp 00E0A427h
  loc_00E0A3F2: lea ecx, var_28
  loc_00E0A3F5: call [00401258h] ; __vbaFreeStr
  loc_00E0A3FB: lea edx, var_30
  loc_00E0A3FE: lea eax, var_2C
  loc_00E0A401: push edx
  loc_00E0A402: push eax
  loc_00E0A403: push 00000002h
  loc_00E0A405: call [00401048h] ; __vbaFreeObjList
  loc_00E0A40B: lea ecx, var_70
  loc_00E0A40E: lea edx, var_60
  loc_00E0A411: push ecx
  loc_00E0A412: lea eax, var_50
  loc_00E0A415: push edx
  loc_00E0A416: lea ecx, var_40
  loc_00E0A419: push eax
  loc_00E0A41A: push ecx
  loc_00E0A41B: push 00000004h
  loc_00E0A41D: call [00401038h] ; __vbaFreeVarList
  loc_00E0A423: add esp, 00000020h
  loc_00E0A426: ret
  loc_00E0A427: lea ecx, var_24
  loc_00E0A42A: call [00401024h] ; __vbaFreeVar
  loc_00E0A430: ret
  loc_00E0A431: mov eax, Me
  loc_00E0A434: push eax
  loc_00E0A435: mov edx, [eax]
  loc_00E0A437: call [edx+00000008h]
  loc_00E0A43A: mov eax, var_4
  loc_00E0A43D: mov ecx, var_14
  loc_00E0A440: pop edi
  loc_00E0A441: pop esi
  loc_00E0A442: mov fs:[00000000h], ecx
  loc_00E0A449: pop ebx
  loc_00E0A44A: mov esp, ebp
  loc_00E0A44C: pop ebp
  loc_00E0A44D: retn 0008h
End Sub

Private Sub cmbjur_Click() 'E083D0
  loc_00E083D0: push ebp
  loc_00E083D1: mov ebp, esp
  loc_00E083D3: sub esp, 0000000Ch
  loc_00E083D6: push 00402836h ; __vbaExceptHandler
  loc_00E083DB: mov eax, fs:[00000000h]
  loc_00E083E1: push eax
  loc_00E083E2: mov fs:[00000000h], esp
  loc_00E083E9: sub esp, 000000ACh
  loc_00E083EF: push ebx
  loc_00E083F0: push esi
  loc_00E083F1: push edi
  loc_00E083F2: mov var_C, esp
  loc_00E083F5: mov var_8, 004020A0h
  loc_00E083FC: mov esi, Me
  loc_00E083FF: mov eax, esi
  loc_00E08401: and eax, 00000001h
  loc_00E08404: mov var_4, eax
  loc_00E08407: and esi, FFFFFFFEh
  loc_00E0840A: push esi
  loc_00E0840B: mov Me, esi
  loc_00E0840E: mov ecx, [esi]
  loc_00E08410: call [ecx+00000004h]
  loc_00E08413: mov edx, [esi]
  loc_00E08415: xor edi, edi
  loc_00E08417: push esi
  loc_00E08418: mov var_18, edi
  loc_00E0841B: mov var_1C, edi
  loc_00E0841E: mov var_20, edi
  loc_00E08421: mov var_30, edi
  loc_00E08424: mov var_40, edi
  loc_00E08427: mov var_50, edi
  loc_00E0842A: mov var_60, edi
  loc_00E0842D: mov var_70, edi
  loc_00E08430: mov var_80, edi
  loc_00E08433: mov var_A4, edi
  loc_00E08439: call [edx+0000030Ch]
  loc_00E0843F: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E08445: push eax
  loc_00E08446: lea eax, var_1C
  loc_00E08449: push eax
  loc_00E0844A: call ebx
  loc_00E0844C: mov ecx, [eax]
  loc_00E0844E: lea edx, var_18
  loc_00E08451: push edx
  loc_00E08452: push eax
  loc_00E08453: mov var_A8, eax
  loc_00E08459: call [ecx+000000A8h]
  loc_00E0845F: cmp eax, edi
  loc_00E08461: fnclex
  loc_00E08463: jge 00E0847Dh
  loc_00E08465: mov ecx, var_A8
  loc_00E0846B: push 000000A8h
  loc_00E08470: push 006E0400h
  loc_00E08475: push ecx
  loc_00E08476: push eax
  loc_00E08477: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0847D: mov edx, [esi]
  loc_00E0847F: push esi
  loc_00E08480: call [edx+00000380h]
  loc_00E08486: push eax
  loc_00E08487: lea eax, var_20
  loc_00E0848A: push eax
  loc_00E0848B: call ebx
  loc_00E0848D: mov ecx, [eax]
  loc_00E0848F: lea edx, var_A4
  loc_00E08495: push edx
  loc_00E08496: push eax
  loc_00E08497: mov var_B0, eax
  loc_00E0849D: call [ecx+000000E0h]
  loc_00E084A3: cmp eax, edi
  loc_00E084A5: fnclex
  loc_00E084A7: jge 00E084C1h
  loc_00E084A9: mov ecx, var_B0
  loc_00E084AF: push 000000E0h
  loc_00E084B4: push 006E03D4h
  loc_00E084B9: push ecx
  loc_00E084BA: push eax
  loc_00E084BB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E084C1: mov edx, var_18
  loc_00E084C4: push edx
  loc_00E084C5: push 006E049Ch ; "Tata Busana"
  loc_00E084CA: call [00401104h] ; __vbaStrCmp
  loc_00E084D0: mov edx, var_A4
  loc_00E084D6: lea ecx, var_18
  loc_00E084D9: neg eax
  loc_00E084DB: sbb eax, eax
  loc_00E084DD: inc eax
  loc_00E084DE: neg eax
  loc_00E084E0: and eax, edx
  loc_00E084E2: mov var_B8, eax
  loc_00E084E8: call [00401258h] ; __vbaFreeStr
  loc_00E084EE: lea eax, var_20
  loc_00E084F1: lea ecx, var_1C
  loc_00E084F4: push eax
  loc_00E084F5: push ecx
  loc_00E084F6: push 00000002h
  loc_00E084F8: call [00401048h] ; __vbaFreeObjList
  loc_00E084FE: add esp, 0000000Ch
  loc_00E08501: cmp var_B8, di
  loc_00E08508: jz 00E08646h
  loc_00E0850E: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E08514: mov ecx, 80020004h
  loc_00E08519: mov var_58, ecx
  loc_00E0851C: mov eax, 0000000Ah
  loc_00E08521: mov var_48, ecx
  loc_00E08524: lea edx, var_80
  loc_00E08527: lea ecx, var_40
  loc_00E0852A: mov var_60, eax
  loc_00E0852D: mov var_50, eax
  loc_00E08530: mov var_78, 006E0590h ; "ERROR 0011"
  loc_00E08537: mov var_80, 00000008h
  loc_00E0853E: call edi
  loc_00E08540: lea edx, var_70
  loc_00E08543: lea ecx, var_30
  loc_00E08546: mov var_68, 006E04B8h ; "Untuk Saat ini Tidak Ada Siswa Laki - laki untuk Jurusan Tata Busana, harap untuk menginput data ulang !"
  loc_00E0854D: mov var_70, 00000008h
  loc_00E08554: call edi
  loc_00E08556: lea edx, var_60
  loc_00E08559: lea eax, var_50
  loc_00E0855C: push edx
  loc_00E0855D: lea ecx, var_40
  loc_00E08560: push eax
  loc_00E08561: push ecx
  loc_00E08562: lea edx, var_30
  loc_00E08565: push 00000010h
  loc_00E08567: push edx
  loc_00E08568: call [004010A8h] ; rtcMsgBox
  loc_00E0856E: lea eax, var_60
  loc_00E08571: lea ecx, var_50
  loc_00E08574: push eax
  loc_00E08575: lea edx, var_40
  loc_00E08578: push ecx
  loc_00E08579: lea eax, var_30
  loc_00E0857C: push edx
  loc_00E0857D: push eax
  loc_00E0857E: push 00000004h
  loc_00E08580: call [00401038h] ; __vbaFreeVarList
  loc_00E08586: mov ecx, [esi]
  loc_00E08588: add esp, 00000014h
  loc_00E0858B: push esi
  loc_00E0858C: call [ecx+0000030Ch]
  loc_00E08592: lea edx, var_1C
  loc_00E08595: push eax
  loc_00E08596: push edx
  loc_00E08597: call ebx
  loc_00E08599: mov edi, eax
  loc_00E0859B: push 006E03E8h ; "-----Pilih"
  loc_00E085A0: push edi
  loc_00E085A1: mov eax, [edi]
  loc_00E085A3: call [eax+000000ACh]
  loc_00E085A9: test eax, eax
  loc_00E085AB: fnclex
  loc_00E085AD: jge 00E085C1h
  loc_00E085AF: push 000000ACh
  loc_00E085B4: push 006E0400h
  loc_00E085B9: push edi
  loc_00E085BA: push eax
  loc_00E085BB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E085C1: lea ecx, var_1C
  loc_00E085C4: call [00401254h] ; __vbaFreeObj
  loc_00E085CA: mov ecx, [esi]
  loc_00E085CC: push esi
  loc_00E085CD: call [ecx+0000037Ch]
  loc_00E085D3: lea edx, var_1C
  loc_00E085D6: push eax
  loc_00E085D7: push edx
  loc_00E085D8: call ebx
  loc_00E085DA: mov edi, eax
  loc_00E085DC: push FFFFFFFFh
  loc_00E085DE: push edi
  loc_00E085DF: mov eax, [edi]
  loc_00E085E1: call [eax+0000009Ch]
  loc_00E085E7: test eax, eax
  loc_00E085E9: fnclex
  loc_00E085EB: jge 00E085FFh
  loc_00E085ED: push 0000009Ch
  loc_00E085F2: push 006E03D4h
  loc_00E085F7: push edi
  loc_00E085F8: push eax
  loc_00E085F9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E085FF: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E08605: lea ecx, var_1C
  loc_00E08608: call edi
  loc_00E0860A: mov ecx, [esi]
  loc_00E0860C: push esi
  loc_00E0860D: call [ecx+00000380h]
  loc_00E08613: lea edx, var_1C
  loc_00E08616: push eax
  loc_00E08617: push edx
  loc_00E08618: call ebx
  loc_00E0861A: mov esi, eax
  loc_00E0861C: push FFFFFFFFh
  loc_00E0861E: push esi
  loc_00E0861F: mov eax, [esi]
  loc_00E08621: call [eax+0000009Ch]
  loc_00E08627: test eax, eax
  loc_00E08629: fnclex
  loc_00E0862B: jge 00E0863Fh
  loc_00E0862D: push 0000009Ch
  loc_00E08632: push 006E03D4h
  loc_00E08637: push esi
  loc_00E08638: push eax
  loc_00E08639: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0863F: lea ecx, var_1C
  loc_00E08642: call edi
  loc_00E08644: xor edi, edi
  loc_00E08646: mov var_4, edi
  loc_00E08649: push 00E08686h
  loc_00E0864E: jmp 00E08685h
  loc_00E08650: lea ecx, var_18
  loc_00E08653: call [00401258h] ; __vbaFreeStr
  loc_00E08659: lea ecx, var_20
  loc_00E0865C: lea edx, var_1C
  loc_00E0865F: push ecx
  loc_00E08660: push edx
  loc_00E08661: push 00000002h
  loc_00E08663: call [00401048h] ; __vbaFreeObjList
  loc_00E08669: lea eax, var_60
  loc_00E0866C: lea ecx, var_50
  loc_00E0866F: push eax
  loc_00E08670: lea edx, var_40
  loc_00E08673: push ecx
  loc_00E08674: lea eax, var_30
  loc_00E08677: push edx
  loc_00E08678: push eax
  loc_00E08679: push 00000004h
  loc_00E0867B: call [00401038h] ; __vbaFreeVarList
  loc_00E08681: add esp, 00000020h
  loc_00E08684: ret
  loc_00E08685: ret
  loc_00E08686: mov eax, Me
  loc_00E08689: push eax
  loc_00E0868A: mov ecx, [eax]
  loc_00E0868C: call [ecx+00000008h]
  loc_00E0868F: mov eax, var_4
  loc_00E08692: mov ecx, var_14
  loc_00E08695: pop edi
  loc_00E08696: pop esi
  loc_00E08697: mov fs:[00000000h], ecx
  loc_00E0869E: pop ebx
  loc_00E0869F: mov esp, ebp
  loc_00E086A1: pop ebp
  loc_00E086A2: retn 0004h
End Sub

Private Sub pibu_KeyPress(KeyAscii As Integer) 'E0AB10
  loc_00E0AB10: push ebp
  loc_00E0AB11: mov ebp, esp
  loc_00E0AB13: sub esp, 0000000Ch
  loc_00E0AB16: push 00402836h ; __vbaExceptHandler
  loc_00E0AB1B: mov eax, fs:[00000000h]
  loc_00E0AB21: push eax
  loc_00E0AB22: mov fs:[00000000h], esp
  loc_00E0AB29: sub esp, 000000B4h
  loc_00E0AB2F: push ebx
  loc_00E0AB30: push esi
  loc_00E0AB31: push edi
  loc_00E0AB32: mov var_C, esp
  loc_00E0AB35: mov var_8, 00402230h
  loc_00E0AB3C: mov esi, Me
  loc_00E0AB3F: mov eax, esi
  loc_00E0AB41: and eax, 00000001h
  loc_00E0AB44: mov var_4, eax
  loc_00E0AB47: and esi, FFFFFFFEh
  loc_00E0AB4A: push esi
  loc_00E0AB4B: mov Me, esi
  loc_00E0AB4E: mov ecx, [esi]
  loc_00E0AB50: call [ecx+00000004h]
  loc_00E0AB53: mov edx, KeyAscii
  loc_00E0AB56: xor eax, eax
  loc_00E0AB58: mov var_24, eax
  loc_00E0AB5B: mov var_28, eax
  loc_00E0AB5E: cmp [edx], 000Dh
  loc_00E0AB62: mov var_38, eax
  loc_00E0AB65: mov var_48, eax
  loc_00E0AB68: mov var_58, eax
  loc_00E0AB6B: mov var_68, eax
  loc_00E0AB6E: mov var_78, eax
  loc_00E0AB71: mov var_88, eax
  loc_00E0AB77: mov var_B8, eax
  loc_00E0AB7D: jnz 00E0ADDBh
  loc_00E0AB83: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E0AB89: mov ecx, 80020004h
  loc_00E0AB8E: mov var_60, ecx
  loc_00E0AB91: mov eax, 0000000Ah
  loc_00E0AB96: mov var_50, ecx
  loc_00E0AB99: mov ebx, 00000008h
  loc_00E0AB9E: lea edx, var_88
  loc_00E0ABA4: lea ecx, var_48
  loc_00E0ABA7: mov var_68, eax
  loc_00E0ABAA: mov var_58, eax
  loc_00E0ABAD: mov var_80, 006DC934h ; "Alert"
  loc_00E0ABB4: mov var_88, ebx
  loc_00E0ABBA: call edi
  loc_00E0ABBC: lea edx, var_78
  loc_00E0ABBF: lea ecx, var_38
  loc_00E0ABC2: mov var_70, 006E0858h ; "Apakah Calon Peserta Didik memilik Wali ?"
  loc_00E0ABC9: mov var_78, ebx
  loc_00E0ABCC: call edi
  loc_00E0ABCE: mov ebx, [004010A8h] ; rtcMsgBox
  loc_00E0ABD4: lea eax, var_68
  loc_00E0ABD7: lea ecx, var_58
  loc_00E0ABDA: push eax
  loc_00E0ABDB: lea edx, var_48
  loc_00E0ABDE: push ecx
  loc_00E0ABDF: push edx
  loc_00E0ABE0: lea eax, var_38
  loc_00E0ABE3: push 00000044h
  loc_00E0ABE5: push eax
  loc_00E0ABE6: call ebx
  loc_00E0ABE8: lea edx, var_B8
  loc_00E0ABEE: lea ecx, var_24
  loc_00E0ABF1: mov var_B0, eax
  loc_00E0ABF7: mov var_B8, 00000003h
  loc_00E0AC01: call [0040101Ch] ; __vbaVarMove
  loc_00E0AC07: lea ecx, var_68
  loc_00E0AC0A: lea edx, var_58
  loc_00E0AC0D: push ecx
  loc_00E0AC0E: lea eax, var_48
  loc_00E0AC11: push edx
  loc_00E0AC12: lea ecx, var_38
  loc_00E0AC15: push eax
  loc_00E0AC16: push ecx
  loc_00E0AC17: push 00000004h
  loc_00E0AC19: call [00401038h] ; __vbaFreeVarList
  loc_00E0AC1F: add esp, 00000014h
  loc_00E0AC22: lea edx, var_24
  loc_00E0AC25: lea eax, var_78
  loc_00E0AC28: mov var_70, 00000006h
  loc_00E0AC2F: push edx
  loc_00E0AC30: push eax
  loc_00E0AC31: mov var_78, 00008003h
  loc_00E0AC38: call [00401108h] ; __vbaVarTstEq
  loc_00E0AC3E: test ax, ax
  loc_00E0AC41: jz 00E0ACBDh
  loc_00E0AC43: mov ecx, [esi]
  loc_00E0AC45: push esi
  loc_00E0AC46: call [ecx+00000338h]
  loc_00E0AC4C: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0AC52: lea edx, var_28
  loc_00E0AC55: push eax
  loc_00E0AC56: push edx
  loc_00E0AC57: call ebx
  loc_00E0AC59: mov edi, eax
  loc_00E0AC5B: push FFFFFFFFh
  loc_00E0AC5D: push edi
  loc_00E0AC5E: mov eax, [edi]
  loc_00E0AC60: call [eax+0000009Ch]
  loc_00E0AC66: test eax, eax
  loc_00E0AC68: fnclex
  loc_00E0AC6A: jge 00E0AC7Eh
  loc_00E0AC6C: push 0000009Ch
  loc_00E0AC71: push 006DCAD0h
  loc_00E0AC76: push edi
  loc_00E0AC77: push eax
  loc_00E0AC78: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0AC7E: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E0AC84: lea ecx, var_28
  loc_00E0AC87: call edi
  loc_00E0AC89: mov ecx, [esi]
  loc_00E0AC8B: push esi
  loc_00E0AC8C: call [ecx+0000033Ch]
  loc_00E0AC92: lea edx, var_28
  loc_00E0AC95: push eax
  loc_00E0AC96: push edx
  loc_00E0AC97: call ebx
  loc_00E0AC99: mov esi, eax
  loc_00E0AC9B: push esi
  loc_00E0AC9C: mov eax, [esi]
  loc_00E0AC9E: call [eax+00000204h]
  loc_00E0ACA4: test eax, eax
  loc_00E0ACA6: fnclex
  loc_00E0ACA8: jge 00E0ADD4h
  loc_00E0ACAE: push 00000204h
  loc_00E0ACB3: push 006DCB70h
  loc_00E0ACB8: jmp 00E0ADCCh
  loc_00E0ACBD: lea ecx, var_24
  loc_00E0ACC0: lea edx, var_78
  loc_00E0ACC3: push ecx
  loc_00E0ACC4: push edx
  loc_00E0ACC5: mov var_70, 00000007h
  loc_00E0ACCC: mov var_78, 00008003h
  loc_00E0ACD3: call [00401108h] ; __vbaVarTstEq
  loc_00E0ACD9: test ax, ax
  loc_00E0ACDC: jz 00E0ADD9h
  loc_00E0ACE2: mov ecx, 80020004h
  loc_00E0ACE7: mov eax, 0000000Ah
  loc_00E0ACEC: mov var_60, ecx
  loc_00E0ACEF: mov var_50, ecx
  loc_00E0ACF2: lea edx, var_88
  loc_00E0ACF8: lea ecx, var_48
  loc_00E0ACFB: mov var_68, eax
  loc_00E0ACFE: mov var_58, eax
  loc_00E0AD01: mov var_80, 006E082Ch ; "All Right Reserved"
  loc_00E0AD08: mov var_88, 00000008h
  loc_00E0AD12: call edi
  loc_00E0AD14: lea edx, var_78
  loc_00E0AD17: lea ecx, var_38
  loc_00E0AD1A: mov var_70, 006E07D0h ; "Silahkan Simpan Data yang Telah Diinput ! "
  loc_00E0AD21: mov var_78, 00000008h
  loc_00E0AD28: call edi
  loc_00E0AD2A: lea eax, var_68
  loc_00E0AD2D: lea ecx, var_58
  loc_00E0AD30: push eax
  loc_00E0AD31: lea edx, var_48
  loc_00E0AD34: push ecx
  loc_00E0AD35: push edx
  loc_00E0AD36: lea eax, var_38
  loc_00E0AD39: push 00000054h
  loc_00E0AD3B: push eax
  loc_00E0AD3C: call ebx
  loc_00E0AD3E: lea ecx, var_68
  loc_00E0AD41: lea edx, var_58
  loc_00E0AD44: push ecx
  loc_00E0AD45: lea eax, var_48
  loc_00E0AD48: push edx
  loc_00E0AD49: lea ecx, var_38
  loc_00E0AD4C: push eax
  loc_00E0AD4D: push ecx
  loc_00E0AD4E: push 00000004h
  loc_00E0AD50: call [00401038h] ; __vbaFreeVarList
  loc_00E0AD56: mov edx, [esi]
  loc_00E0AD58: add esp, 00000014h
  loc_00E0AD5B: push esi
  loc_00E0AD5C: call [edx+000003E8h]
  loc_00E0AD62: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0AD68: push eax
  loc_00E0AD69: lea eax, var_28
  loc_00E0AD6C: push eax
  loc_00E0AD6D: call ebx
  loc_00E0AD6F: mov edi, eax
  loc_00E0AD71: push FFFFFFFFh
  loc_00E0AD73: push edi
  loc_00E0AD74: mov ecx, [edi]
  loc_00E0AD76: call [ecx+0000008Ch]
  loc_00E0AD7C: test eax, eax
  loc_00E0AD7E: fnclex
  loc_00E0AD80: jge 00E0AD94h
  loc_00E0AD82: push 0000008Ch
  loc_00E0AD87: push 006DCDA0h
  loc_00E0AD8C: push edi
  loc_00E0AD8D: push eax
  loc_00E0AD8E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0AD94: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E0AD9A: lea ecx, var_28
  loc_00E0AD9D: call edi
  loc_00E0AD9F: mov edx, [esi]
  loc_00E0ADA1: push esi
  loc_00E0ADA2: call [edx+00000338h]
  loc_00E0ADA8: push eax
  loc_00E0ADA9: lea eax, var_28
  loc_00E0ADAC: push eax
  loc_00E0ADAD: call ebx
  loc_00E0ADAF: mov esi, eax
  loc_00E0ADB1: push 00000000h
  loc_00E0ADB3: push esi
  loc_00E0ADB4: mov ecx, [esi]
  loc_00E0ADB6: call [ecx+0000009Ch]
  loc_00E0ADBC: test eax, eax
  loc_00E0ADBE: fnclex
  loc_00E0ADC0: jge 00E0ADD4h
  loc_00E0ADC2: push 0000009Ch
  loc_00E0ADC7: push 006DCAD0h
  loc_00E0ADCC: push esi
  loc_00E0ADCD: push eax
  loc_00E0ADCE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0ADD4: lea ecx, var_28
  loc_00E0ADD7: call edi
  loc_00E0ADD9: xor eax, eax
  loc_00E0ADDB: mov var_4, eax
  loc_00E0ADDE: push 00E0AE14h
  loc_00E0ADE3: jmp 00E0AE0Ah
  loc_00E0ADE5: lea ecx, var_28
  loc_00E0ADE8: call [00401254h] ; __vbaFreeObj
  loc_00E0ADEE: lea edx, var_68
  loc_00E0ADF1: lea eax, var_58
  loc_00E0ADF4: push edx
  loc_00E0ADF5: lea ecx, var_48
  loc_00E0ADF8: push eax
  loc_00E0ADF9: lea edx, var_38
  loc_00E0ADFC: push ecx
  loc_00E0ADFD: push edx
  loc_00E0ADFE: push 00000004h
  loc_00E0AE00: call [00401038h] ; __vbaFreeVarList
  loc_00E0AE06: add esp, 00000014h
  loc_00E0AE09: ret
  loc_00E0AE0A: lea ecx, var_24
  loc_00E0AE0D: call [00401024h] ; __vbaFreeVar
  loc_00E0AE13: ret
  loc_00E0AE14: mov eax, Me
  loc_00E0AE17: push eax
  loc_00E0AE18: mov ecx, [eax]
  loc_00E0AE1A: call [ecx+00000008h]
  loc_00E0AE1D: mov eax, var_4
  loc_00E0AE20: mov ecx, var_14
  loc_00E0AE23: pop edi
  loc_00E0AE24: pop esi
  loc_00E0AE25: mov fs:[00000000h], ecx
  loc_00E0AE2C: pop ebx
  loc_00E0AE2D: mov esp, ebp
  loc_00E0AE2F: pop ebp
  loc_00E0AE30: retn 0008h
End Sub

Private Sub pwali_KeyPress(KeyAscii As Integer) 'E0A770
  loc_00E0A770: push ebp
  loc_00E0A771: mov ebp, esp
  loc_00E0A773: sub esp, 0000000Ch
  loc_00E0A776: push 00402836h ; __vbaExceptHandler
  loc_00E0A77B: mov eax, fs:[00000000h]
  loc_00E0A781: push eax
  loc_00E0A782: mov fs:[00000000h], esp
  loc_00E0A789: sub esp, 00000094h
  loc_00E0A78F: push ebx
  loc_00E0A790: push esi
  loc_00E0A791: push edi
  loc_00E0A792: mov var_C, esp
  loc_00E0A795: mov var_8, 00402200h
  loc_00E0A79C: mov esi, Me
  loc_00E0A79F: mov eax, esi
  loc_00E0A7A1: and eax, 00000001h
  loc_00E0A7A4: mov var_4, eax
  loc_00E0A7A7: and esi, FFFFFFFEh
  loc_00E0A7AA: push esi
  loc_00E0A7AB: mov Me, esi
  loc_00E0A7AE: mov ecx, [esi]
  loc_00E0A7B0: call [ecx+00000004h]
  loc_00E0A7B3: mov edx, KeyAscii
  loc_00E0A7B6: xor eax, eax
  loc_00E0A7B8: mov var_18, eax
  loc_00E0A7BB: mov var_28, eax
  loc_00E0A7BE: cmp [edx], 000Dh
  loc_00E0A7C2: mov var_38, eax
  loc_00E0A7C5: mov var_48, eax
  loc_00E0A7C8: mov var_58, eax
  loc_00E0A7CB: mov var_68, eax
  loc_00E0A7CE: mov var_78, eax
  loc_00E0A7D1: jnz 00E0A893h
  loc_00E0A7D7: mov edi, [004011F0h] ; __vbaVarDup
  loc_00E0A7DD: mov ecx, 80020004h
  loc_00E0A7E2: mov var_50, ecx
  loc_00E0A7E5: mov eax, 0000000Ah
  loc_00E0A7EA: mov var_40, ecx
  loc_00E0A7ED: mov ebx, 00000008h
  loc_00E0A7F2: lea edx, var_78
  loc_00E0A7F5: lea ecx, var_38
  loc_00E0A7F8: mov var_58, eax
  loc_00E0A7FB: mov var_48, eax
  loc_00E0A7FE: mov var_70, 006E082Ch ; "All Right Reserved"
  loc_00E0A805: mov var_78, ebx
  loc_00E0A808: call edi
  loc_00E0A80A: lea edx, var_68
  loc_00E0A80D: lea ecx, var_28
  loc_00E0A810: mov var_60, 006E07D0h ; "Silahkan Simpan Data yang Telah Diinput ! "
  loc_00E0A817: mov var_68, ebx
  loc_00E0A81A: call edi
  loc_00E0A81C: lea eax, var_58
  loc_00E0A81F: lea ecx, var_48
  loc_00E0A822: push eax
  loc_00E0A823: lea edx, var_38
  loc_00E0A826: push ecx
  loc_00E0A827: push edx
  loc_00E0A828: lea eax, var_28
  loc_00E0A82B: push 00000044h
  loc_00E0A82D: push eax
  loc_00E0A82E: call [004010A8h] ; rtcMsgBox
  loc_00E0A834: lea ecx, var_58
  loc_00E0A837: lea edx, var_48
  loc_00E0A83A: push ecx
  loc_00E0A83B: lea eax, var_38
  loc_00E0A83E: push edx
  loc_00E0A83F: lea ecx, var_28
  loc_00E0A842: push eax
  loc_00E0A843: push ecx
  loc_00E0A844: push 00000004h
  loc_00E0A846: call [00401038h] ; __vbaFreeVarList
  loc_00E0A84C: mov edx, [esi]
  loc_00E0A84E: add esp, 00000014h
  loc_00E0A851: push esi
  loc_00E0A852: call [edx+000003E8h]
  loc_00E0A858: push eax
  loc_00E0A859: lea eax, var_18
  loc_00E0A85C: push eax
  loc_00E0A85D: call [004010ACh] ; __vbaObjSet
  loc_00E0A863: mov esi, eax
  loc_00E0A865: push FFFFFFFFh
  loc_00E0A867: push esi
  loc_00E0A868: mov ecx, [esi]
  loc_00E0A86A: call [ecx+0000008Ch]
  loc_00E0A870: test eax, eax
  loc_00E0A872: fnclex
  loc_00E0A874: jge 00E0A888h
  loc_00E0A876: push 0000008Ch
  loc_00E0A87B: push 006DCDA0h
  loc_00E0A880: push esi
  loc_00E0A881: push eax
  loc_00E0A882: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A888: lea ecx, var_18
  loc_00E0A88B: call [00401254h] ; __vbaFreeObj
  loc_00E0A891: xor eax, eax
  loc_00E0A893: mov var_4, eax
  loc_00E0A896: push 00E0A8C3h
  loc_00E0A89B: jmp 00E0A8C2h
  loc_00E0A89D: lea ecx, var_18
  loc_00E0A8A0: call [00401254h] ; __vbaFreeObj
  loc_00E0A8A6: lea edx, var_58
  loc_00E0A8A9: lea eax, var_48
  loc_00E0A8AC: push edx
  loc_00E0A8AD: lea ecx, var_38
  loc_00E0A8B0: push eax
  loc_00E0A8B1: lea edx, var_28
  loc_00E0A8B4: push ecx
  loc_00E0A8B5: push edx
  loc_00E0A8B6: push 00000004h
  loc_00E0A8B8: call [00401038h] ; __vbaFreeVarList
  loc_00E0A8BE: add esp, 00000014h
  loc_00E0A8C1: ret
  loc_00E0A8C2: ret
  loc_00E0A8C3: mov eax, Me
  loc_00E0A8C6: push eax
  loc_00E0A8C7: mov ecx, [eax]
  loc_00E0A8C9: call [ecx+00000008h]
  loc_00E0A8CC: mov eax, var_4
  loc_00E0A8CF: mov ecx, var_14
  loc_00E0A8D2: pop edi
  loc_00E0A8D3: pop esi
  loc_00E0A8D4: mov fs:[00000000h], ecx
  loc_00E0A8DB: pop ebx
  loc_00E0A8DC: mov esp, ebp
  loc_00E0A8DE: pop ebp
  loc_00E0A8DF: retn 0008h
End Sub

Private Sub simpan_UnknownEvent_9 'E0C3F0
  loc_00E0C3F0: push ebp
  loc_00E0C3F1: mov ebp, esp
  loc_00E0C3F3: sub esp, 0000000Ch
  loc_00E0C3F6: push 00402836h ; __vbaExceptHandler
  loc_00E0C3FB: mov eax, fs:[00000000h]
  loc_00E0C401: push eax
  loc_00E0C402: mov fs:[00000000h], esp
  loc_00E0C409: sub esp, 0000011Ch
  loc_00E0C40F: push ebx
  loc_00E0C410: push esi
  loc_00E0C411: push edi
  loc_00E0C412: mov var_C, esp
  loc_00E0C415: mov var_8, 00402260h
  loc_00E0C41C: mov esi, Me
  loc_00E0C41F: mov eax, esi
  loc_00E0C421: and eax, 00000001h
  loc_00E0C424: mov var_4, eax
  loc_00E0C427: and esi, FFFFFFFEh
  loc_00E0C42A: push esi
  loc_00E0C42B: mov Me, esi
  loc_00E0C42E: mov ecx, [esi]
  loc_00E0C430: call [ecx+00000004h]
  loc_00E0C433: mov edx, [esi]
  loc_00E0C435: xor ebx, ebx
  loc_00E0C437: push esi
  loc_00E0C438: mov var_18, ebx
  loc_00E0C43B: mov var_1C, ebx
  loc_00E0C43E: mov var_20, ebx
  loc_00E0C441: mov var_24, ebx
  loc_00E0C444: mov var_28, ebx
  loc_00E0C447: mov var_2C, ebx
  loc_00E0C44A: mov var_30, ebx
  loc_00E0C44D: mov var_34, ebx
  loc_00E0C450: mov var_44, ebx
  loc_00E0C453: mov var_54, ebx
  loc_00E0C456: mov var_64, ebx
  loc_00E0C459: mov var_74, ebx
  loc_00E0C45C: mov var_84, ebx
  loc_00E0C462: mov var_94, ebx
  loc_00E0C468: mov var_B8, ebx
  loc_00E0C46E: call [edx+0000039Ch]
  loc_00E0C474: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E0C47A: push eax
  loc_00E0C47B: lea eax, var_24
  loc_00E0C47E: push eax
  loc_00E0C47F: call edi
  loc_00E0C481: mov ecx, [eax]
  loc_00E0C483: lea edx, var_18
  loc_00E0C486: push edx
  loc_00E0C487: push eax
  loc_00E0C488: mov var_BC, eax
  loc_00E0C48E: call [ecx+000000A0h]
  loc_00E0C494: cmp eax, ebx
  loc_00E0C496: fnclex
  loc_00E0C498: jge 00E0C4B2h
  loc_00E0C49A: mov ecx, var_BC
  loc_00E0C4A0: push 000000A0h
  loc_00E0C4A5: push 006DCB70h
  loc_00E0C4AA: push ecx
  loc_00E0C4AB: push eax
  loc_00E0C4AC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0C4B2: mov edx, var_18
  loc_00E0C4B5: push 006E0994h ; "select * from biodata where no_daftar ='"
  loc_00E0C4BA: push edx
  loc_00E0C4BB: lea ebx, [esi+00000034h]
  loc_00E0C4BE: call [0040106Ch] ; __vbaStrCat
  loc_00E0C4C4: mov edx, eax
  loc_00E0C4C6: lea ecx, var_1C
  loc_00E0C4C9: call [00401228h] ; __vbaStrMove
  loc_00E0C4CF: push eax
  loc_00E0C4D0: push 006DCB84h ; "'"
  loc_00E0C4D5: call [0040106Ch] ; __vbaStrCat
  loc_00E0C4DB: mov edx, eax
  loc_00E0C4DD: lea ecx, var_20
  loc_00E0C4E0: call [00401228h] ; __vbaStrMove
  loc_00E0C4E6: mov edx, eax
  loc_00E0C4E8: mov ecx, ebx
  loc_00E0C4EA: call [004011B0h] ; __vbaStrCopy
  loc_00E0C4F0: lea eax, var_20
  loc_00E0C4F3: lea ecx, var_1C
  loc_00E0C4F6: push eax
  loc_00E0C4F7: lea edx, var_18
  loc_00E0C4FA: push ecx
  loc_00E0C4FB: push edx
  loc_00E0C4FC: push 00000003h
  loc_00E0C4FE: call [004011BCh] ; __vbaFreeStrList
  loc_00E0C504: add esp, 00000010h
  loc_00E0C507: lea ecx, var_24
  loc_00E0C50A: call [00401254h] ; __vbaFreeObj
  loc_00E0C510: mov edx, var_80
  loc_00E0C513: sub esp, 00000010h
  loc_00E0C516: mov ecx, esp
  loc_00E0C518: mov eax, 00004008h
  loc_00E0C51D: mov var_84, eax
  loc_00E0C523: push 0000000Eh
  loc_00E0C525: mov [ecx], eax
  loc_00E0C527: mov eax, var_78
  loc_00E0C52A: push esi
  loc_00E0C52B: mov var_7C, ebx
  loc_00E0C52E: mov [ecx+00000004h], edx
  loc_00E0C531: mov [ecx+00000008h], ebx
  loc_00E0C534: mov [ecx+0000000Ch], eax
  loc_00E0C537: mov ecx, [esi]
  loc_00E0C539: call [ecx+00000490h]
  loc_00E0C53F: lea edx, var_24
  loc_00E0C542: push eax
  loc_00E0C543: push edx
  loc_00E0C544: call edi
  loc_00E0C546: push eax
  loc_00E0C547: call [00401238h] ; __vbaLateIdSt
  loc_00E0C54D: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E0C553: lea ecx, var_24
  loc_00E0C556: call ebx
  loc_00E0C558: mov eax, [esi]
  loc_00E0C55A: push 00000000h
  loc_00E0C55C: push 00000019h
  loc_00E0C55E: push esi
  loc_00E0C55F: call [eax+00000490h]
  loc_00E0C565: lea ecx, var_24
  loc_00E0C568: push eax
  loc_00E0C569: push ecx
  loc_00E0C56A: call edi
  loc_00E0C56C: push eax
  loc_00E0C56D: call [00401030h] ; __vbaLateIdCall
  loc_00E0C573: add esp, 0000000Ch
  loc_00E0C576: lea ecx, var_24
  loc_00E0C579: call ebx
  loc_00E0C57B: mov edx, [esi]
  loc_00E0C57D: push esi
  loc_00E0C57E: call [edx+0000030Ch]
  loc_00E0C584: push eax
  loc_00E0C585: lea eax, var_24
  loc_00E0C588: push eax
  loc_00E0C589: call edi
  loc_00E0C58B: mov ebx, eax
  loc_00E0C58D: lea edx, var_18
  loc_00E0C590: push edx
  loc_00E0C591: push ebx
  loc_00E0C592: mov ecx, [ebx]
  loc_00E0C594: call [ecx+000000A8h]
  loc_00E0C59A: fnclex
  loc_00E0C59C: test eax, eax
  loc_00E0C59E: jge 00E0C5B2h
  loc_00E0C5A0: push 000000A8h
  loc_00E0C5A5: push 006E0400h
  loc_00E0C5AA: push ebx
  loc_00E0C5AB: push eax
  loc_00E0C5AC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0C5B2: mov eax, var_18
  loc_00E0C5B5: push eax
  loc_00E0C5B6: push 006E03E8h ; "-----Pilih"
  loc_00E0C5BB: call [00401104h] ; __vbaStrCmp
  loc_00E0C5C1: mov ebx, eax
  loc_00E0C5C3: lea ecx, var_18
  loc_00E0C5C6: neg ebx
  loc_00E0C5C8: sbb ebx, ebx
  loc_00E0C5CA: inc ebx
  loc_00E0C5CB: neg ebx
  loc_00E0C5CD: call [00401258h] ; __vbaFreeStr
  loc_00E0C5D3: lea ecx, var_24
  loc_00E0C5D6: call [00401254h] ; __vbaFreeObj
  loc_00E0C5DC: test bx, bx
  loc_00E0C5DF: jnz 00E0C704h
  loc_00E0C5E5: mov eax, [esi]
  loc_00E0C5E7: push esi
  loc_00E0C5E8: call [eax+00000378h]
  loc_00E0C5EE: lea ecx, var_24
  loc_00E0C5F1: push eax
  loc_00E0C5F2: push ecx
  loc_00E0C5F3: call edi
  loc_00E0C5F5: mov ebx, eax
  loc_00E0C5F7: lea eax, var_18
  loc_00E0C5FA: push eax
  loc_00E0C5FB: push ebx
  loc_00E0C5FC: mov edx, [ebx]
  loc_00E0C5FE: call [edx+000000A8h]
  loc_00E0C604: test eax, eax
  loc_00E0C606: fnclex
  loc_00E0C608: jge 00E0C61Ch
  loc_00E0C60A: push 000000A8h
  loc_00E0C60F: push 006E0400h
  loc_00E0C614: push ebx
  loc_00E0C615: push eax
  loc_00E0C616: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0C61C: mov ecx, var_18
  loc_00E0C61F: push ecx
  loc_00E0C620: push 006E03E8h ; "-----Pilih"
  loc_00E0C625: call [00401104h] ; __vbaStrCmp
  loc_00E0C62B: mov ebx, eax
  loc_00E0C62D: lea ecx, var_18
  loc_00E0C630: neg ebx
  loc_00E0C632: sbb ebx, ebx
  loc_00E0C634: inc ebx
  loc_00E0C635: neg ebx
  loc_00E0C637: call [00401258h] ; __vbaFreeStr
  loc_00E0C63D: lea ecx, var_24
  loc_00E0C640: call [00401254h] ; __vbaFreeObj
  loc_00E0C646: test bx, bx
  loc_00E0C649: jz 00E0C6B2h
  loc_00E0C64B: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E0C651: mov ecx, 80020004h
  loc_00E0C656: mov var_6C, ecx
  loc_00E0C659: mov eax, 0000000Ah
  loc_00E0C65E: mov var_5C, ecx
  loc_00E0C661: mov edi, 00000008h
  loc_00E0C666: lea edx, var_94
  loc_00E0C66C: lea ecx, var_54
  loc_00E0C66F: mov var_74, eax
  loc_00E0C672: mov var_64, eax
  loc_00E0C675: mov var_8C, 006E0B08h ; "Need To Do"
  loc_00E0C67F: mov var_94, edi
  loc_00E0C685: call __vbaVarDup
  loc_00E0C687: lea edx, var_84
  loc_00E0C68D: lea ecx, var_44
  loc_00E0C690: mov var_7C, 006E0A8Ch ; "Mohon Periksa Kembali Data Anda, Isi Seluruh Field data !!"
  loc_00E0C697: mov var_84, edi
  loc_00E0C69D: call __vbaVarDup
  loc_00E0C69F: lea edx, var_74
  loc_00E0C6A2: lea eax, var_64
  loc_00E0C6A5: push edx
  loc_00E0C6A6: lea ecx, var_54
  loc_00E0C6A9: push eax
  loc_00E0C6AA: push ecx
  loc_00E0C6AB: push 00000040h
  loc_00E0C6AD: jmp 00E0F3A0h
  loc_00E0C6B2: mov ecx, [esi]
  loc_00E0C6B4: push 00000000h
  loc_00E0C6B6: push 80010007h
  loc_00E0C6BB: push esi
  loc_00E0C6BC: call [ecx+000003B8h]
  loc_00E0C6C2: lea edx, var_24
  loc_00E0C6C5: push eax
  loc_00E0C6C6: push edx
  loc_00E0C6C7: call edi
  loc_00E0C6C9: push eax
  loc_00E0C6CA: lea eax, var_44
  loc_00E0C6CD: push eax
  loc_00E0C6CE: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0C6D4: add esp, 00000010h
  loc_00E0C6D7: push eax
  loc_00E0C6D8: call [004010CCh] ; __vbaBoolVar
  loc_00E0C6DE: mov bx, ax
  loc_00E0C6E1: lea ecx, var_24
  loc_00E0C6E4: neg bx
  loc_00E0C6E7: sbb ebx, ebx
  loc_00E0C6E9: inc ebx
  loc_00E0C6EA: neg ebx
  loc_00E0C6EC: call [00401254h] ; __vbaFreeObj
  loc_00E0C6F2: lea ecx, var_44
  loc_00E0C6F5: call [00401024h] ; __vbaFreeVar
  loc_00E0C6FB: test bx, bx
  loc_00E0C6FE: jz 00E0C785h
  loc_00E0C704: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E0C70A: mov ecx, 80020004h
  loc_00E0C70F: mov var_6C, ecx
  loc_00E0C712: mov eax, 0000000Ah
  loc_00E0C717: mov var_5C, ecx
  loc_00E0C71A: mov edi, 00000008h
  loc_00E0C71F: lea edx, var_94
  loc_00E0C725: lea ecx, var_54
  loc_00E0C728: mov var_74, eax
  loc_00E0C72B: mov var_64, eax
  loc_00E0C72E: mov var_8C, 006E0B08h ; "Need To Do"
  loc_00E0C738: mov var_94, edi
  loc_00E0C73E: call __vbaVarDup
  loc_00E0C740: lea edx, var_84
  loc_00E0C746: lea ecx, var_44
  loc_00E0C749: mov var_7C, 006E0A8Ch ; "Mohon Periksa Kembali Data Anda, Isi Seluruh Field data !!"
  loc_00E0C750: mov var_84, edi
  loc_00E0C756: call __vbaVarDup
  loc_00E0C758: lea ecx, var_74
  loc_00E0C75B: lea edx, var_64
  loc_00E0C75E: push ecx
  loc_00E0C75F: lea eax, var_54
  loc_00E0C762: push edx
  loc_00E0C763: push eax
  loc_00E0C764: lea ecx, var_44
  loc_00E0C767: push 00000040h
  loc_00E0C769: push ecx
  loc_00E0C76A: call [004010A8h] ; rtcMsgBox
  loc_00E0C770: lea edx, var_74
  loc_00E0C773: lea eax, var_64
  loc_00E0C776: push edx
  loc_00E0C777: lea ecx, var_54
  loc_00E0C77A: push eax
  loc_00E0C77B: lea edx, var_44
  loc_00E0C77E: push ecx
  loc_00E0C77F: push edx
  loc_00E0C780: jmp 00E0F3BAh
  loc_00E0C785: mov eax, [esi]
  loc_00E0C787: push 006DCBD8h
  loc_00E0C78C: push 00000000h
  loc_00E0C78E: push 00000018h
  loc_00E0C790: push esi
  loc_00E0C791: call [eax+00000490h]
  loc_00E0C797: lea ecx, var_24
  loc_00E0C79A: push eax
  loc_00E0C79B: push ecx
  loc_00E0C79C: call edi
  loc_00E0C79E: lea edx, var_44
  loc_00E0C7A1: push eax
  loc_00E0C7A2: push edx
  loc_00E0C7A3: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0C7A9: add esp, 00000010h
  loc_00E0C7AC: push eax
  loc_00E0C7AD: call [00401120h] ; __vbaCastObjVar
  loc_00E0C7B3: push eax
  loc_00E0C7B4: lea eax, var_28
  loc_00E0C7B7: push eax
  loc_00E0C7B8: call edi
  loc_00E0C7BA: mov ebx, eax
  loc_00E0C7BC: lea edx, var_B8
  loc_00E0C7C2: push edx
  loc_00E0C7C3: push ebx
  loc_00E0C7C4: mov ecx, [ebx]
  loc_00E0C7C6: call [ecx+00000050h]
  loc_00E0C7C9: test eax, eax
  loc_00E0C7CB: fnclex
  loc_00E0C7CD: jge 00E0C7DEh
  loc_00E0C7CF: push 00000050h
  loc_00E0C7D1: push 006DCBE8h
  loc_00E0C7D6: push ebx
  loc_00E0C7D7: push eax
  loc_00E0C7D8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0C7DE: mov bx, var_B8
  loc_00E0C7E5: lea eax, var_28
  loc_00E0C7E8: lea ecx, var_24
  loc_00E0C7EB: push eax
  loc_00E0C7EC: push ecx
  loc_00E0C7ED: push 00000002h
  loc_00E0C7EF: call [00401048h] ; __vbaFreeObjList
  loc_00E0C7F5: add esp, 0000000Ch
  loc_00E0C7F8: lea ecx, var_44
  loc_00E0C7FB: call [00401024h] ; __vbaFreeVar
  loc_00E0C801: test bx, bx
  loc_00E0C804: jz 00E0F356h
  loc_00E0C80A: mov edx, [esi]
  loc_00E0C80C: push 006DCBD8h
  loc_00E0C811: push 00000000h
  loc_00E0C813: push 00000018h
  loc_00E0C815: push esi
  loc_00E0C816: call [edx+00000490h]
  loc_00E0C81C: push eax
  loc_00E0C81D: lea eax, var_24
  loc_00E0C820: push eax
  loc_00E0C821: call edi
  loc_00E0C823: lea ecx, var_44
  loc_00E0C826: push eax
  loc_00E0C827: push ecx
  loc_00E0C828: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0C82E: add esp, 00000010h
  loc_00E0C831: push eax
  loc_00E0C832: call [00401120h] ; __vbaCastObjVar
  loc_00E0C838: lea edx, var_28
  loc_00E0C83B: push eax
  loc_00E0C83C: push edx
  loc_00E0C83D: call edi
  loc_00E0C83F: sub esp, 00000010h
  loc_00E0C842: mov ebx, eax
  loc_00E0C844: mov edx, esp
  loc_00E0C846: mov eax, 0000000Ah
  loc_00E0C84B: mov var_94, eax
  loc_00E0C851: mov var_84, eax
  loc_00E0C857: mov [edx], eax
  loc_00E0C859: mov eax, var_90
  loc_00E0C85F: mov ecx, 80020004h
  loc_00E0C864: sub esp, 00000010h
  loc_00E0C867: mov var_8C, ecx
  loc_00E0C86D: mov [edx+00000004h], eax
  loc_00E0C870: mov eax, var_8C
  loc_00E0C876: mov var_7C, ecx
  loc_00E0C879: mov [edx+00000008h], eax
  loc_00E0C87C: mov eax, var_88
  loc_00E0C882: mov ecx, [ebx]
  loc_00E0C884: mov [edx+0000000Ch], eax
  loc_00E0C887: mov eax, var_84
  loc_00E0C88D: mov edx, esp
  loc_00E0C88F: push ebx
  loc_00E0C890: mov [edx], eax
  loc_00E0C892: mov eax, var_80
  loc_00E0C895: mov [edx+00000004h], eax
  loc_00E0C898: mov eax, var_7C
  loc_00E0C89B: mov [edx+00000008h], eax
  loc_00E0C89E: mov eax, var_78
  loc_00E0C8A1: mov [edx+0000000Ch], eax
  loc_00E0C8A4: call [ecx+00000078h]
  loc_00E0C8A7: test eax, eax
  loc_00E0C8A9: fnclex
  loc_00E0C8AB: jge 00E0C8BCh
  loc_00E0C8AD: push 00000078h
  loc_00E0C8AF: push 006DCBE8h
  loc_00E0C8B4: push ebx
  loc_00E0C8B5: push eax
  loc_00E0C8B6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0C8BC: lea ecx, var_28
  loc_00E0C8BF: lea edx, var_24
  loc_00E0C8C2: push ecx
  loc_00E0C8C3: push edx
  loc_00E0C8C4: push 00000002h
  loc_00E0C8C6: call [00401048h] ; __vbaFreeObjList
  loc_00E0C8CC: add esp, 0000000Ch
  loc_00E0C8CF: lea ecx, var_44
  loc_00E0C8D2: call [00401024h] ; __vbaFreeVar
  loc_00E0C8D8: mov eax, [esi]
  loc_00E0C8DA: push 006DCBD8h
  loc_00E0C8DF: push 00000000h
  loc_00E0C8E1: push 00000018h
  loc_00E0C8E3: push esi
  loc_00E0C8E4: call [eax+00000490h]
  loc_00E0C8EA: lea ecx, var_28
  loc_00E0C8ED: push eax
  loc_00E0C8EE: push ecx
  loc_00E0C8EF: call edi
  loc_00E0C8F1: lea edx, var_44
  loc_00E0C8F4: push eax
  loc_00E0C8F5: push edx
  loc_00E0C8F6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0C8FC: add esp, 00000010h
  loc_00E0C8FF: push eax
  loc_00E0C900: call [00401120h] ; __vbaCastObjVar
  loc_00E0C906: push eax
  loc_00E0C907: lea eax, var_2C
  loc_00E0C90A: push eax
  loc_00E0C90B: call edi
  loc_00E0C90D: mov ebx, eax
  loc_00E0C90F: lea edx, var_30
  loc_00E0C912: push edx
  loc_00E0C913: push ebx
  loc_00E0C914: mov ecx, [ebx]
  loc_00E0C916: call [ecx+00000054h]
  loc_00E0C919: test eax, eax
  loc_00E0C91B: fnclex
  loc_00E0C91D: jge 00E0C92Eh
  loc_00E0C91F: push 00000054h
  loc_00E0C921: push 006DCBE8h
  loc_00E0C926: push ebx
  loc_00E0C927: push eax
  loc_00E0C928: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0C92E: lea ebx, var_34
  loc_00E0C931: mov eax, var_30
  loc_00E0C934: push ebx
  loc_00E0C935: mov ecx, 00000002h
  loc_00E0C93A: sub esp, 00000010h
  loc_00E0C93D: mov var_84, ecx
  loc_00E0C943: mov ebx, esp
  loc_00E0C945: mov var_7C, 00000000h
  loc_00E0C94C: mov edx, [eax]
  loc_00E0C94E: push eax
  loc_00E0C94F: mov [ebx], ecx
  loc_00E0C951: mov ecx, var_80
  loc_00E0C954: mov var_CC, eax
  loc_00E0C95A: mov [ebx+00000004h], ecx
  loc_00E0C95D: mov ecx, var_7C
  loc_00E0C960: mov [ebx+00000008h], ecx
  loc_00E0C963: mov ecx, var_78
  loc_00E0C966: mov [ebx+0000000Ch], ecx
  loc_00E0C969: call [edx+00000028h]
  loc_00E0C96C: test eax, eax
  loc_00E0C96E: fnclex
  loc_00E0C970: jge 00E0C987h
  loc_00E0C972: mov edx, var_CC
  loc_00E0C978: push 00000028h
  loc_00E0C97A: push 006E09E8h
  loc_00E0C97F: push edx
  loc_00E0C980: push eax
  loc_00E0C981: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0C987: mov eax, var_34
  loc_00E0C98A: mov ecx, [esi]
  loc_00E0C98C: push esi
  loc_00E0C98D: mov var_D4, eax
  loc_00E0C993: call [ecx+0000039Ch]
  loc_00E0C999: lea edx, var_24
  loc_00E0C99C: push eax
  loc_00E0C99D: push edx
  loc_00E0C99E: call edi
  loc_00E0C9A0: mov ebx, eax
  loc_00E0C9A2: lea ecx, var_18
  loc_00E0C9A5: push ecx
  loc_00E0C9A6: push ebx
  loc_00E0C9A7: mov eax, [ebx]
  loc_00E0C9A9: call [eax+000000A0h]
  loc_00E0C9AF: test eax, eax
  loc_00E0C9B1: fnclex
  loc_00E0C9B3: jge 00E0C9C7h
  loc_00E0C9B5: push 000000A0h
  loc_00E0C9BA: push 006DCB70h
  loc_00E0C9BF: push ebx
  loc_00E0C9C0: push eax
  loc_00E0C9C1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0C9C7: sub esp, 00000010h
  loc_00E0C9CA: mov eax, var_18
  loc_00E0C9CD: mov ebx, esp
  loc_00E0C9CF: mov ecx, 00000008h
  loc_00E0C9D4: mov var_54, ecx
  loc_00E0C9D7: mov edx, var_D4
  loc_00E0C9DD: mov [ebx], ecx
  loc_00E0C9DF: mov ecx, var_50
  loc_00E0C9E2: mov var_4C, eax
  loc_00E0C9E5: mov var_18, 00000000h
  loc_00E0C9EC: mov edx, [edx]
  loc_00E0C9EE: mov [ebx+00000004h], ecx
  loc_00E0C9F1: mov [ebx+00000008h], eax
  loc_00E0C9F4: mov eax, var_48
  loc_00E0C9F7: mov [ebx+0000000Ch], eax
  loc_00E0C9FA: mov ebx, var_D4
  loc_00E0CA00: push ebx
  loc_00E0CA01: call [edx+00000038h]
  loc_00E0CA04: test eax, eax
  loc_00E0CA06: fnclex
  loc_00E0CA08: jge 00E0CA19h
  loc_00E0CA0A: push 00000038h
  loc_00E0CA0C: push 006E09F8h
  loc_00E0CA11: push ebx
  loc_00E0CA12: push eax
  loc_00E0CA13: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CA19: lea ecx, var_34
  loc_00E0CA1C: lea edx, var_30
  loc_00E0CA1F: push ecx
  loc_00E0CA20: lea eax, var_2C
  loc_00E0CA23: push edx
  loc_00E0CA24: lea ecx, var_28
  loc_00E0CA27: push eax
  loc_00E0CA28: lea edx, var_24
  loc_00E0CA2B: push ecx
  loc_00E0CA2C: push edx
  loc_00E0CA2D: push 00000005h
  loc_00E0CA2F: call [00401048h] ; __vbaFreeObjList
  loc_00E0CA35: lea eax, var_54
  loc_00E0CA38: lea ecx, var_44
  loc_00E0CA3B: push eax
  loc_00E0CA3C: push ecx
  loc_00E0CA3D: push 00000002h
  loc_00E0CA3F: call [00401038h] ; __vbaFreeVarList
  loc_00E0CA45: mov edx, [esi]
  loc_00E0CA47: add esp, 00000024h
  loc_00E0CA4A: push 006DCBD8h
  loc_00E0CA4F: push 00000000h
  loc_00E0CA51: push 00000018h
  loc_00E0CA53: push esi
  loc_00E0CA54: call [edx+00000490h]
  loc_00E0CA5A: push eax
  loc_00E0CA5B: lea eax, var_28
  loc_00E0CA5E: push eax
  loc_00E0CA5F: call edi
  loc_00E0CA61: lea ecx, var_44
  loc_00E0CA64: push eax
  loc_00E0CA65: push ecx
  loc_00E0CA66: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0CA6C: add esp, 00000010h
  loc_00E0CA6F: push eax
  loc_00E0CA70: call [00401120h] ; __vbaCastObjVar
  loc_00E0CA76: lea edx, var_2C
  loc_00E0CA79: push eax
  loc_00E0CA7A: push edx
  loc_00E0CA7B: call edi
  loc_00E0CA7D: mov ebx, eax
  loc_00E0CA7F: lea ecx, var_30
  loc_00E0CA82: push ecx
  loc_00E0CA83: push ebx
  loc_00E0CA84: mov eax, [ebx]
  loc_00E0CA86: call [eax+00000054h]
  loc_00E0CA89: test eax, eax
  loc_00E0CA8B: fnclex
  loc_00E0CA8D: jge 00E0CA9Eh
  loc_00E0CA8F: push 00000054h
  loc_00E0CA91: push 006DCBE8h
  loc_00E0CA96: push ebx
  loc_00E0CA97: push eax
  loc_00E0CA98: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CA9E: lea ebx, var_34
  loc_00E0CAA1: mov eax, var_30
  loc_00E0CAA4: push ebx
  loc_00E0CAA5: mov ecx, 00000002h
  loc_00E0CAAA: sub esp, 00000010h
  loc_00E0CAAD: mov var_84, ecx
  loc_00E0CAB3: mov ebx, esp
  loc_00E0CAB5: mov var_7C, 00000001h
  loc_00E0CABC: mov edx, [eax]
  loc_00E0CABE: push eax
  loc_00E0CABF: mov [ebx], ecx
  loc_00E0CAC1: mov ecx, var_80
  loc_00E0CAC4: mov var_CC, eax
  loc_00E0CACA: mov [ebx+00000004h], ecx
  loc_00E0CACD: mov ecx, var_7C
  loc_00E0CAD0: mov [ebx+00000008h], ecx
  loc_00E0CAD3: mov ecx, var_78
  loc_00E0CAD6: mov [ebx+0000000Ch], ecx
  loc_00E0CAD9: call [edx+00000028h]
  loc_00E0CADC: test eax, eax
  loc_00E0CADE: fnclex
  loc_00E0CAE0: jge 00E0CAF7h
  loc_00E0CAE2: mov edx, var_CC
  loc_00E0CAE8: push 00000028h
  loc_00E0CAEA: push 006E09E8h
  loc_00E0CAEF: push edx
  loc_00E0CAF0: push eax
  loc_00E0CAF1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CAF7: mov eax, var_34
  loc_00E0CAFA: mov ecx, [esi]
  loc_00E0CAFC: push esi
  loc_00E0CAFD: mov var_D4, eax
  loc_00E0CB03: call [ecx+00000394h]
  loc_00E0CB09: lea edx, var_24
  loc_00E0CB0C: push eax
  loc_00E0CB0D: push edx
  loc_00E0CB0E: call edi
  loc_00E0CB10: mov ebx, eax
  loc_00E0CB12: lea ecx, var_18
  loc_00E0CB15: push ecx
  loc_00E0CB16: push ebx
  loc_00E0CB17: mov eax, [ebx]
  loc_00E0CB19: call [eax+000000A0h]
  loc_00E0CB1F: test eax, eax
  loc_00E0CB21: fnclex
  loc_00E0CB23: jge 00E0CB37h
  loc_00E0CB25: push 000000A0h
  loc_00E0CB2A: push 006DCB70h
  loc_00E0CB2F: push ebx
  loc_00E0CB30: push eax
  loc_00E0CB31: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CB37: sub esp, 00000010h
  loc_00E0CB3A: mov eax, var_18
  loc_00E0CB3D: mov ebx, esp
  loc_00E0CB3F: mov ecx, 00000008h
  loc_00E0CB44: mov var_54, ecx
  loc_00E0CB47: mov edx, var_D4
  loc_00E0CB4D: mov [ebx], ecx
  loc_00E0CB4F: mov ecx, var_50
  loc_00E0CB52: mov var_4C, eax
  loc_00E0CB55: mov var_18, 00000000h
  loc_00E0CB5C: mov edx, [edx]
  loc_00E0CB5E: mov [ebx+00000004h], ecx
  loc_00E0CB61: mov [ebx+00000008h], eax
  loc_00E0CB64: mov eax, var_48
  loc_00E0CB67: mov [ebx+0000000Ch], eax
  loc_00E0CB6A: mov ebx, var_D4
  loc_00E0CB70: push ebx
  loc_00E0CB71: call [edx+00000038h]
  loc_00E0CB74: test eax, eax
  loc_00E0CB76: fnclex
  loc_00E0CB78: jge 00E0CB89h
  loc_00E0CB7A: push 00000038h
  loc_00E0CB7C: push 006E09F8h
  loc_00E0CB81: push ebx
  loc_00E0CB82: push eax
  loc_00E0CB83: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CB89: lea ecx, var_34
  loc_00E0CB8C: lea edx, var_30
  loc_00E0CB8F: push ecx
  loc_00E0CB90: lea eax, var_2C
  loc_00E0CB93: push edx
  loc_00E0CB94: lea ecx, var_28
  loc_00E0CB97: push eax
  loc_00E0CB98: lea edx, var_24
  loc_00E0CB9B: push ecx
  loc_00E0CB9C: push edx
  loc_00E0CB9D: push 00000005h
  loc_00E0CB9F: call [00401048h] ; __vbaFreeObjList
  loc_00E0CBA5: lea eax, var_54
  loc_00E0CBA8: lea ecx, var_44
  loc_00E0CBAB: push eax
  loc_00E0CBAC: push ecx
  loc_00E0CBAD: push 00000002h
  loc_00E0CBAF: call [00401038h] ; __vbaFreeVarList
  loc_00E0CBB5: mov edx, [esi]
  loc_00E0CBB7: add esp, 00000024h
  loc_00E0CBBA: push esi
  loc_00E0CBBB: call [edx+00000380h]
  loc_00E0CBC1: push eax
  loc_00E0CBC2: lea eax, var_24
  loc_00E0CBC5: push eax
  loc_00E0CBC6: call edi
  loc_00E0CBC8: mov ebx, eax
  loc_00E0CBCA: lea edx, var_B8
  loc_00E0CBD0: push edx
  loc_00E0CBD1: push ebx
  loc_00E0CBD2: mov ecx, [ebx]
  loc_00E0CBD4: call [ecx+000000E0h]
  loc_00E0CBDA: test eax, eax
  loc_00E0CBDC: fnclex
  loc_00E0CBDE: jge 00E0CBF2h
  loc_00E0CBE0: push 000000E0h
  loc_00E0CBE5: push 006E03D4h
  loc_00E0CBEA: push ebx
  loc_00E0CBEB: push eax
  loc_00E0CBEC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CBF2: mov bx, var_B8
  loc_00E0CBF9: lea ecx, var_24
  loc_00E0CBFC: call [00401254h] ; __vbaFreeObj
  loc_00E0CC02: push 006DCBD8h
  loc_00E0CC07: push 00000000h
  loc_00E0CC09: test bx, bx
  loc_00E0CC0C: push 00000018h
  loc_00E0CC0E: jz 00E0CD0Bh
  loc_00E0CC14: mov eax, [esi]
  loc_00E0CC16: push esi
  loc_00E0CC17: call [eax+00000490h]
  loc_00E0CC1D: lea ecx, var_24
  loc_00E0CC20: push eax
  loc_00E0CC21: push ecx
  loc_00E0CC22: call edi
  loc_00E0CC24: lea edx, var_44
  loc_00E0CC27: push eax
  loc_00E0CC28: push edx
  loc_00E0CC29: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0CC2F: add esp, 00000010h
  loc_00E0CC32: push eax
  loc_00E0CC33: call [00401120h] ; __vbaCastObjVar
  loc_00E0CC39: push eax
  loc_00E0CC3A: lea eax, var_28
  loc_00E0CC3D: push eax
  loc_00E0CC3E: call edi
  loc_00E0CC40: mov ebx, eax
  loc_00E0CC42: lea edx, var_2C
  loc_00E0CC45: push edx
  loc_00E0CC46: push ebx
  loc_00E0CC47: mov ecx, [ebx]
  loc_00E0CC49: call [ecx+00000054h]
  loc_00E0CC4C: test eax, eax
  loc_00E0CC4E: fnclex
  loc_00E0CC50: jge 00E0CC61h
  loc_00E0CC52: push 00000054h
  loc_00E0CC54: push 006DCBE8h
  loc_00E0CC59: push ebx
  loc_00E0CC5A: push eax
  loc_00E0CC5B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CC61: lea ebx, var_30
  loc_00E0CC64: mov eax, var_2C
  loc_00E0CC67: push ebx
  loc_00E0CC68: mov ecx, 00000002h
  loc_00E0CC6D: sub esp, 00000010h
  loc_00E0CC70: mov var_7C, ecx
  loc_00E0CC73: mov ebx, esp
  loc_00E0CC75: mov var_84, ecx
  loc_00E0CC7B: mov edx, [eax]
  loc_00E0CC7D: push eax
  loc_00E0CC7E: mov [ebx], ecx
  loc_00E0CC80: mov ecx, var_80
  loc_00E0CC83: mov var_C4, eax
  loc_00E0CC89: mov [ebx+00000004h], ecx
  loc_00E0CC8C: mov ecx, var_7C
  loc_00E0CC8F: mov [ebx+00000008h], ecx
  loc_00E0CC92: mov ecx, var_78
  loc_00E0CC95: mov [ebx+0000000Ch], ecx
  loc_00E0CC98: call [edx+00000028h]
  loc_00E0CC9B: test eax, eax
  loc_00E0CC9D: fnclex
  loc_00E0CC9F: jge 00E0CCB6h
  loc_00E0CCA1: mov edx, var_C4
  loc_00E0CCA7: push 00000028h
  loc_00E0CCA9: push 006E09E8h
  loc_00E0CCAE: push edx
  loc_00E0CCAF: push eax
  loc_00E0CCB0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CCB6: sub esp, 00000010h
  loc_00E0CCB9: mov eax, var_30
  loc_00E0CCBC: mov ebx, esp
  loc_00E0CCBE: mov ecx, 00000008h
  loc_00E0CCC3: mov var_94, ecx
  loc_00E0CCC9: mov var_8C, 006E0B24h ; "Laki - laki"
  loc_00E0CCD3: mov edx, [eax]
  loc_00E0CCD5: mov [ebx], ecx
  loc_00E0CCD7: mov ecx, var_90
  loc_00E0CCDD: push eax
  loc_00E0CCDE: mov [ebx+00000004h], ecx
  loc_00E0CCE1: mov ecx, var_8C
  loc_00E0CCE7: mov var_CC, eax
  loc_00E0CCED: mov [ebx+00000008h], ecx
  loc_00E0CCF0: mov ecx, var_88
  loc_00E0CCF6: mov [ebx+0000000Ch], ecx
  loc_00E0CCF9: call [edx+00000038h]
  loc_00E0CCFC: test eax, eax
  loc_00E0CCFE: fnclex
  loc_00E0CD00: jge 00E0CE0Eh
  loc_00E0CD06: jmp 00E0CDF9h
  loc_00E0CD0B: mov ecx, [esi]
  loc_00E0CD0D: push esi
  loc_00E0CD0E: call [ecx+00000490h]
  loc_00E0CD14: lea edx, var_24
  loc_00E0CD17: push eax
  loc_00E0CD18: push edx
  loc_00E0CD19: call edi
  loc_00E0CD1B: push eax
  loc_00E0CD1C: lea eax, var_44
  loc_00E0CD1F: push eax
  loc_00E0CD20: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0CD26: add esp, 00000010h
  loc_00E0CD29: push eax
  loc_00E0CD2A: call [00401120h] ; __vbaCastObjVar
  loc_00E0CD30: lea ecx, var_28
  loc_00E0CD33: push eax
  loc_00E0CD34: push ecx
  loc_00E0CD35: call edi
  loc_00E0CD37: mov ebx, eax
  loc_00E0CD39: lea eax, var_2C
  loc_00E0CD3C: push eax
  loc_00E0CD3D: push ebx
  loc_00E0CD3E: mov edx, [ebx]
  loc_00E0CD40: call [edx+00000054h]
  loc_00E0CD43: test eax, eax
  loc_00E0CD45: fnclex
  loc_00E0CD47: jge 00E0CD58h
  loc_00E0CD49: push 00000054h
  loc_00E0CD4B: push 006DCBE8h
  loc_00E0CD50: push ebx
  loc_00E0CD51: push eax
  loc_00E0CD52: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CD58: lea ebx, var_30
  loc_00E0CD5B: mov eax, var_2C
  loc_00E0CD5E: push ebx
  loc_00E0CD5F: mov ecx, 00000002h
  loc_00E0CD64: sub esp, 00000010h
  loc_00E0CD67: mov var_7C, ecx
  loc_00E0CD6A: mov ebx, esp
  loc_00E0CD6C: mov var_84, ecx
  loc_00E0CD72: mov edx, [eax]
  loc_00E0CD74: push eax
  loc_00E0CD75: mov [ebx], ecx
  loc_00E0CD77: mov ecx, var_80
  loc_00E0CD7A: mov var_C4, eax
  loc_00E0CD80: mov [ebx+00000004h], ecx
  loc_00E0CD83: mov ecx, var_7C
  loc_00E0CD86: mov [ebx+00000008h], ecx
  loc_00E0CD89: mov ecx, var_78
  loc_00E0CD8C: mov [ebx+0000000Ch], ecx
  loc_00E0CD8F: call [edx+00000028h]
  loc_00E0CD92: test eax, eax
  loc_00E0CD94: fnclex
  loc_00E0CD96: jge 00E0CDADh
  loc_00E0CD98: mov edx, var_C4
  loc_00E0CD9E: push 00000028h
  loc_00E0CDA0: push 006E09E8h
  loc_00E0CDA5: push edx
  loc_00E0CDA6: push eax
  loc_00E0CDA7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CDAD: sub esp, 00000010h
  loc_00E0CDB0: mov eax, var_30
  loc_00E0CDB3: mov ebx, esp
  loc_00E0CDB5: mov ecx, 00000008h
  loc_00E0CDBA: mov var_94, ecx
  loc_00E0CDC0: mov var_8C, 006E0B40h ; "Perempuan"
  loc_00E0CDCA: mov edx, [eax]
  loc_00E0CDCC: mov [ebx], ecx
  loc_00E0CDCE: mov ecx, var_90
  loc_00E0CDD4: push eax
  loc_00E0CDD5: mov [ebx+00000004h], ecx
  loc_00E0CDD8: mov ecx, var_8C
  loc_00E0CDDE: mov var_CC, eax
  loc_00E0CDE4: mov [ebx+00000008h], ecx
  loc_00E0CDE7: mov ecx, var_88
  loc_00E0CDED: mov [ebx+0000000Ch], ecx
  loc_00E0CDF0: call [edx+00000038h]
  loc_00E0CDF3: test eax, eax
  loc_00E0CDF5: fnclex
  loc_00E0CDF7: jge 00E0CE0Eh
  loc_00E0CDF9: mov edx, var_CC
  loc_00E0CDFF: push 00000038h
  loc_00E0CE01: push 006E09F8h
  loc_00E0CE06: push edx
  loc_00E0CE07: push eax
  loc_00E0CE08: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CE0E: lea eax, var_30
  loc_00E0CE11: lea ecx, var_2C
  loc_00E0CE14: push eax
  loc_00E0CE15: lea edx, var_28
  loc_00E0CE18: push ecx
  loc_00E0CE19: lea eax, var_24
  loc_00E0CE1C: push edx
  loc_00E0CE1D: push eax
  loc_00E0CE1E: push 00000004h
  loc_00E0CE20: call [00401048h] ; __vbaFreeObjList
  loc_00E0CE26: add esp, 00000014h
  loc_00E0CE29: lea ecx, var_44
  loc_00E0CE2C: call [00401024h] ; __vbaFreeVar
  loc_00E0CE32: mov ecx, [esi]
  loc_00E0CE34: push 006DCBD8h
  loc_00E0CE39: push 00000000h
  loc_00E0CE3B: push 00000018h
  loc_00E0CE3D: push esi
  loc_00E0CE3E: call [ecx+00000490h]
  loc_00E0CE44: lea edx, var_28
  loc_00E0CE47: push eax
  loc_00E0CE48: push edx
  loc_00E0CE49: call edi
  loc_00E0CE4B: push eax
  loc_00E0CE4C: lea eax, var_44
  loc_00E0CE4F: push eax
  loc_00E0CE50: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0CE56: add esp, 00000010h
  loc_00E0CE59: push eax
  loc_00E0CE5A: call [00401120h] ; __vbaCastObjVar
  loc_00E0CE60: lea ecx, var_2C
  loc_00E0CE63: push eax
  loc_00E0CE64: push ecx
  loc_00E0CE65: call edi
  loc_00E0CE67: mov ebx, eax
  loc_00E0CE69: lea eax, var_30
  loc_00E0CE6C: push eax
  loc_00E0CE6D: push ebx
  loc_00E0CE6E: mov edx, [ebx]
  loc_00E0CE70: call [edx+00000054h]
  loc_00E0CE73: test eax, eax
  loc_00E0CE75: fnclex
  loc_00E0CE77: jge 00E0CE88h
  loc_00E0CE79: push 00000054h
  loc_00E0CE7B: push 006DCBE8h
  loc_00E0CE80: push ebx
  loc_00E0CE81: push eax
  loc_00E0CE82: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CE88: lea ebx, var_34
  loc_00E0CE8B: mov eax, var_30
  loc_00E0CE8E: push ebx
  loc_00E0CE8F: mov ecx, 00000002h
  loc_00E0CE94: sub esp, 00000010h
  loc_00E0CE97: mov var_84, ecx
  loc_00E0CE9D: mov ebx, esp
  loc_00E0CE9F: mov var_7C, 00000003h
  loc_00E0CEA6: mov edx, [eax]
  loc_00E0CEA8: push eax
  loc_00E0CEA9: mov [ebx], ecx
  loc_00E0CEAB: mov ecx, var_80
  loc_00E0CEAE: mov var_CC, eax
  loc_00E0CEB4: mov [ebx+00000004h], ecx
  loc_00E0CEB7: mov ecx, var_7C
  loc_00E0CEBA: mov [ebx+00000008h], ecx
  loc_00E0CEBD: mov ecx, var_78
  loc_00E0CEC0: mov [ebx+0000000Ch], ecx
  loc_00E0CEC3: call [edx+00000028h]
  loc_00E0CEC6: test eax, eax
  loc_00E0CEC8: fnclex
  loc_00E0CECA: jge 00E0CEE1h
  loc_00E0CECC: mov edx, var_CC
  loc_00E0CED2: push 00000028h
  loc_00E0CED4: push 006E09E8h
  loc_00E0CED9: push edx
  loc_00E0CEDA: push eax
  loc_00E0CEDB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CEE1: mov eax, var_34
  loc_00E0CEE4: mov ecx, [esi]
  loc_00E0CEE6: push esi
  loc_00E0CEE7: mov var_D4, eax
  loc_00E0CEED: call [ecx+00000390h]
  loc_00E0CEF3: lea edx, var_24
  loc_00E0CEF6: push eax
  loc_00E0CEF7: push edx
  loc_00E0CEF8: call edi
  loc_00E0CEFA: mov ebx, eax
  loc_00E0CEFC: lea ecx, var_18
  loc_00E0CEFF: push ecx
  loc_00E0CF00: push ebx
  loc_00E0CF01: mov eax, [ebx]
  loc_00E0CF03: call [eax+000000A0h]
  loc_00E0CF09: test eax, eax
  loc_00E0CF0B: fnclex
  loc_00E0CF0D: jge 00E0CF21h
  loc_00E0CF0F: push 000000A0h
  loc_00E0CF14: push 006DCB70h
  loc_00E0CF19: push ebx
  loc_00E0CF1A: push eax
  loc_00E0CF1B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CF21: sub esp, 00000010h
  loc_00E0CF24: mov eax, var_18
  loc_00E0CF27: mov ebx, esp
  loc_00E0CF29: mov ecx, 00000008h
  loc_00E0CF2E: mov var_54, ecx
  loc_00E0CF31: mov edx, var_D4
  loc_00E0CF37: mov [ebx], ecx
  loc_00E0CF39: mov ecx, var_50
  loc_00E0CF3C: mov var_4C, eax
  loc_00E0CF3F: mov var_18, 00000000h
  loc_00E0CF46: mov edx, [edx]
  loc_00E0CF48: mov [ebx+00000004h], ecx
  loc_00E0CF4B: mov [ebx+00000008h], eax
  loc_00E0CF4E: mov eax, var_48
  loc_00E0CF51: mov [ebx+0000000Ch], eax
  loc_00E0CF54: mov ebx, var_D4
  loc_00E0CF5A: push ebx
  loc_00E0CF5B: call [edx+00000038h]
  loc_00E0CF5E: test eax, eax
  loc_00E0CF60: fnclex
  loc_00E0CF62: jge 00E0CF73h
  loc_00E0CF64: push 00000038h
  loc_00E0CF66: push 006E09F8h
  loc_00E0CF6B: push ebx
  loc_00E0CF6C: push eax
  loc_00E0CF6D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CF73: lea ecx, var_34
  loc_00E0CF76: lea edx, var_30
  loc_00E0CF79: push ecx
  loc_00E0CF7A: lea eax, var_2C
  loc_00E0CF7D: push edx
  loc_00E0CF7E: lea ecx, var_28
  loc_00E0CF81: push eax
  loc_00E0CF82: lea edx, var_24
  loc_00E0CF85: push ecx
  loc_00E0CF86: push edx
  loc_00E0CF87: push 00000005h
  loc_00E0CF89: call [00401048h] ; __vbaFreeObjList
  loc_00E0CF8F: lea eax, var_54
  loc_00E0CF92: lea ecx, var_44
  loc_00E0CF95: push eax
  loc_00E0CF96: push ecx
  loc_00E0CF97: push 00000002h
  loc_00E0CF99: call [00401038h] ; __vbaFreeVarList
  loc_00E0CF9F: mov edx, [esi]
  loc_00E0CFA1: add esp, 00000024h
  loc_00E0CFA4: push 006DCBD8h
  loc_00E0CFA9: push 00000000h
  loc_00E0CFAB: push 00000018h
  loc_00E0CFAD: push esi
  loc_00E0CFAE: call [edx+00000490h]
  loc_00E0CFB4: push eax
  loc_00E0CFB5: lea eax, var_28
  loc_00E0CFB8: push eax
  loc_00E0CFB9: call edi
  loc_00E0CFBB: lea ecx, var_44
  loc_00E0CFBE: push eax
  loc_00E0CFBF: push ecx
  loc_00E0CFC0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0CFC6: add esp, 00000010h
  loc_00E0CFC9: push eax
  loc_00E0CFCA: call [00401120h] ; __vbaCastObjVar
  loc_00E0CFD0: lea edx, var_2C
  loc_00E0CFD3: push eax
  loc_00E0CFD4: push edx
  loc_00E0CFD5: call edi
  loc_00E0CFD7: mov ebx, eax
  loc_00E0CFD9: lea ecx, var_30
  loc_00E0CFDC: push ecx
  loc_00E0CFDD: push ebx
  loc_00E0CFDE: mov eax, [ebx]
  loc_00E0CFE0: call [eax+00000054h]
  loc_00E0CFE3: test eax, eax
  loc_00E0CFE5: fnclex
  loc_00E0CFE7: jge 00E0CFF8h
  loc_00E0CFE9: push 00000054h
  loc_00E0CFEB: push 006DCBE8h
  loc_00E0CFF0: push ebx
  loc_00E0CFF1: push eax
  loc_00E0CFF2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0CFF8: lea ebx, var_34
  loc_00E0CFFB: mov eax, var_30
  loc_00E0CFFE: push ebx
  loc_00E0CFFF: mov ecx, 00000002h
  loc_00E0D004: sub esp, 00000010h
  loc_00E0D007: mov var_84, ecx
  loc_00E0D00D: mov ebx, esp
  loc_00E0D00F: mov var_7C, 00000004h
  loc_00E0D016: mov edx, [eax]
  loc_00E0D018: push eax
  loc_00E0D019: mov [ebx], ecx
  loc_00E0D01B: mov ecx, var_80
  loc_00E0D01E: mov var_CC, eax
  loc_00E0D024: mov [ebx+00000004h], ecx
  loc_00E0D027: mov ecx, var_7C
  loc_00E0D02A: mov [ebx+00000008h], ecx
  loc_00E0D02D: mov ecx, var_78
  loc_00E0D030: mov [ebx+0000000Ch], ecx
  loc_00E0D033: call [edx+00000028h]
  loc_00E0D036: test eax, eax
  loc_00E0D038: fnclex
  loc_00E0D03A: jge 00E0D051h
  loc_00E0D03C: mov edx, var_CC
  loc_00E0D042: push 00000028h
  loc_00E0D044: push 006E09E8h
  loc_00E0D049: push edx
  loc_00E0D04A: push eax
  loc_00E0D04B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D051: mov eax, var_34
  loc_00E0D054: mov ecx, [esi]
  loc_00E0D056: push esi
  loc_00E0D057: mov var_D4, eax
  loc_00E0D05D: call [ecx+0000038Ch]
  loc_00E0D063: lea edx, var_24
  loc_00E0D066: push eax
  loc_00E0D067: push edx
  loc_00E0D068: call edi
  loc_00E0D06A: mov ebx, eax
  loc_00E0D06C: lea ecx, var_18
  loc_00E0D06F: push ecx
  loc_00E0D070: push ebx
  loc_00E0D071: mov eax, [ebx]
  loc_00E0D073: call [eax+000000A0h]
  loc_00E0D079: test eax, eax
  loc_00E0D07B: fnclex
  loc_00E0D07D: jge 00E0D091h
  loc_00E0D07F: push 000000A0h
  loc_00E0D084: push 006DCB70h
  loc_00E0D089: push ebx
  loc_00E0D08A: push eax
  loc_00E0D08B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D091: sub esp, 00000010h
  loc_00E0D094: mov eax, var_18
  loc_00E0D097: mov ebx, esp
  loc_00E0D099: mov ecx, 00000008h
  loc_00E0D09E: mov var_54, ecx
  loc_00E0D0A1: mov edx, var_D4
  loc_00E0D0A7: mov [ebx], ecx
  loc_00E0D0A9: mov ecx, var_50
  loc_00E0D0AC: mov var_4C, eax
  loc_00E0D0AF: mov var_18, 00000000h
  loc_00E0D0B6: mov edx, [edx]
  loc_00E0D0B8: mov [ebx+00000004h], ecx
  loc_00E0D0BB: mov [ebx+00000008h], eax
  loc_00E0D0BE: mov eax, var_48
  loc_00E0D0C1: mov [ebx+0000000Ch], eax
  loc_00E0D0C4: mov ebx, var_D4
  loc_00E0D0CA: push ebx
  loc_00E0D0CB: call [edx+00000038h]
  loc_00E0D0CE: test eax, eax
  loc_00E0D0D0: fnclex
  loc_00E0D0D2: jge 00E0D0E3h
  loc_00E0D0D4: push 00000038h
  loc_00E0D0D6: push 006E09F8h
  loc_00E0D0DB: push ebx
  loc_00E0D0DC: push eax
  loc_00E0D0DD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D0E3: lea ecx, var_34
  loc_00E0D0E6: lea edx, var_30
  loc_00E0D0E9: push ecx
  loc_00E0D0EA: lea eax, var_2C
  loc_00E0D0ED: push edx
  loc_00E0D0EE: lea ecx, var_28
  loc_00E0D0F1: push eax
  loc_00E0D0F2: lea edx, var_24
  loc_00E0D0F5: push ecx
  loc_00E0D0F6: push edx
  loc_00E0D0F7: push 00000005h
  loc_00E0D0F9: call [00401048h] ; __vbaFreeObjList
  loc_00E0D0FF: lea eax, var_54
  loc_00E0D102: lea ecx, var_44
  loc_00E0D105: push eax
  loc_00E0D106: push ecx
  loc_00E0D107: push 00000002h
  loc_00E0D109: call [00401038h] ; __vbaFreeVarList
  loc_00E0D10F: mov edx, [esi]
  loc_00E0D111: add esp, 00000024h
  loc_00E0D114: push 006DCBD8h
  loc_00E0D119: push 00000000h
  loc_00E0D11B: push 00000018h
  loc_00E0D11D: push esi
  loc_00E0D11E: call [edx+00000490h]
  loc_00E0D124: push eax
  loc_00E0D125: lea eax, var_28
  loc_00E0D128: push eax
  loc_00E0D129: call edi
  loc_00E0D12B: lea ecx, var_44
  loc_00E0D12E: push eax
  loc_00E0D12F: push ecx
  loc_00E0D130: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0D136: add esp, 00000010h
  loc_00E0D139: push eax
  loc_00E0D13A: call [00401120h] ; __vbaCastObjVar
  loc_00E0D140: lea edx, var_2C
  loc_00E0D143: push eax
  loc_00E0D144: push edx
  loc_00E0D145: call edi
  loc_00E0D147: mov ebx, eax
  loc_00E0D149: lea ecx, var_30
  loc_00E0D14C: push ecx
  loc_00E0D14D: push ebx
  loc_00E0D14E: mov eax, [ebx]
  loc_00E0D150: call [eax+00000054h]
  loc_00E0D153: test eax, eax
  loc_00E0D155: fnclex
  loc_00E0D157: jge 00E0D168h
  loc_00E0D159: push 00000054h
  loc_00E0D15B: push 006DCBE8h
  loc_00E0D160: push ebx
  loc_00E0D161: push eax
  loc_00E0D162: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D168: lea ebx, var_34
  loc_00E0D16B: mov eax, var_30
  loc_00E0D16E: push ebx
  loc_00E0D16F: mov ecx, 00000002h
  loc_00E0D174: sub esp, 00000010h
  loc_00E0D177: mov var_84, ecx
  loc_00E0D17D: mov ebx, esp
  loc_00E0D17F: mov var_7C, 00000005h
  loc_00E0D186: mov edx, [eax]
  loc_00E0D188: push eax
  loc_00E0D189: mov [ebx], ecx
  loc_00E0D18B: mov ecx, var_80
  loc_00E0D18E: mov var_CC, eax
  loc_00E0D194: mov [ebx+00000004h], ecx
  loc_00E0D197: mov ecx, var_7C
  loc_00E0D19A: mov [ebx+00000008h], ecx
  loc_00E0D19D: mov ecx, var_78
  loc_00E0D1A0: mov [ebx+0000000Ch], ecx
  loc_00E0D1A3: call [edx+00000028h]
  loc_00E0D1A6: test eax, eax
  loc_00E0D1A8: fnclex
  loc_00E0D1AA: jge 00E0D1C1h
  loc_00E0D1AC: mov edx, var_CC
  loc_00E0D1B2: push 00000028h
  loc_00E0D1B4: push 006E09E8h
  loc_00E0D1B9: push edx
  loc_00E0D1BA: push eax
  loc_00E0D1BB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D1C1: mov eax, var_34
  loc_00E0D1C4: mov ecx, [esi]
  loc_00E0D1C6: push esi
  loc_00E0D1C7: mov var_D4, eax
  loc_00E0D1CD: call [ecx+00000388h]
  loc_00E0D1D3: lea edx, var_24
  loc_00E0D1D6: push eax
  loc_00E0D1D7: push edx
  loc_00E0D1D8: call edi
  loc_00E0D1DA: mov ebx, eax
  loc_00E0D1DC: lea ecx, var_18
  loc_00E0D1DF: push ecx
  loc_00E0D1E0: push ebx
  loc_00E0D1E1: mov eax, [ebx]
  loc_00E0D1E3: call [eax+000000A0h]
  loc_00E0D1E9: test eax, eax
  loc_00E0D1EB: fnclex
  loc_00E0D1ED: jge 00E0D201h
  loc_00E0D1EF: push 000000A0h
  loc_00E0D1F4: push 006DCB70h
  loc_00E0D1F9: push ebx
  loc_00E0D1FA: push eax
  loc_00E0D1FB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D201: sub esp, 00000010h
  loc_00E0D204: mov eax, var_18
  loc_00E0D207: mov ebx, esp
  loc_00E0D209: mov ecx, 00000008h
  loc_00E0D20E: mov var_54, ecx
  loc_00E0D211: mov edx, var_D4
  loc_00E0D217: mov [ebx], ecx
  loc_00E0D219: mov ecx, var_50
  loc_00E0D21C: mov var_4C, eax
  loc_00E0D21F: mov var_18, 00000000h
  loc_00E0D226: mov edx, [edx]
  loc_00E0D228: mov [ebx+00000004h], ecx
  loc_00E0D22B: mov [ebx+00000008h], eax
  loc_00E0D22E: mov eax, var_48
  loc_00E0D231: mov [ebx+0000000Ch], eax
  loc_00E0D234: mov ebx, var_D4
  loc_00E0D23A: push ebx
  loc_00E0D23B: call [edx+00000038h]
  loc_00E0D23E: test eax, eax
  loc_00E0D240: fnclex
  loc_00E0D242: jge 00E0D253h
  loc_00E0D244: push 00000038h
  loc_00E0D246: push 006E09F8h
  loc_00E0D24B: push ebx
  loc_00E0D24C: push eax
  loc_00E0D24D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D253: lea ecx, var_34
  loc_00E0D256: lea edx, var_30
  loc_00E0D259: push ecx
  loc_00E0D25A: lea eax, var_2C
  loc_00E0D25D: push edx
  loc_00E0D25E: lea ecx, var_28
  loc_00E0D261: push eax
  loc_00E0D262: lea edx, var_24
  loc_00E0D265: push ecx
  loc_00E0D266: push edx
  loc_00E0D267: push 00000005h
  loc_00E0D269: call [00401048h] ; __vbaFreeObjList
  loc_00E0D26F: lea eax, var_54
  loc_00E0D272: lea ecx, var_44
  loc_00E0D275: push eax
  loc_00E0D276: push ecx
  loc_00E0D277: push 00000002h
  loc_00E0D279: call [00401038h] ; __vbaFreeVarList
  loc_00E0D27F: mov edx, [esi]
  loc_00E0D281: add esp, 00000024h
  loc_00E0D284: push 006DCBD8h
  loc_00E0D289: push 00000000h
  loc_00E0D28B: push 00000018h
  loc_00E0D28D: push esi
  loc_00E0D28E: call [edx+00000490h]
  loc_00E0D294: push eax
  loc_00E0D295: lea eax, var_28
  loc_00E0D298: push eax
  loc_00E0D299: call edi
  loc_00E0D29B: lea ecx, var_44
  loc_00E0D29E: push eax
  loc_00E0D29F: push ecx
  loc_00E0D2A0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0D2A6: add esp, 00000010h
  loc_00E0D2A9: push eax
  loc_00E0D2AA: call [00401120h] ; __vbaCastObjVar
  loc_00E0D2B0: lea edx, var_2C
  loc_00E0D2B3: push eax
  loc_00E0D2B4: push edx
  loc_00E0D2B5: call edi
  loc_00E0D2B7: mov ebx, eax
  loc_00E0D2B9: lea ecx, var_30
  loc_00E0D2BC: push ecx
  loc_00E0D2BD: push ebx
  loc_00E0D2BE: mov eax, [ebx]
  loc_00E0D2C0: call [eax+00000054h]
  loc_00E0D2C3: test eax, eax
  loc_00E0D2C5: fnclex
  loc_00E0D2C7: jge 00E0D2D8h
  loc_00E0D2C9: push 00000054h
  loc_00E0D2CB: push 006DCBE8h
  loc_00E0D2D0: push ebx
  loc_00E0D2D1: push eax
  loc_00E0D2D2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D2D8: lea ebx, var_34
  loc_00E0D2DB: mov eax, var_30
  loc_00E0D2DE: push ebx
  loc_00E0D2DF: mov ecx, 00000002h
  loc_00E0D2E4: sub esp, 00000010h
  loc_00E0D2E7: mov var_84, ecx
  loc_00E0D2ED: mov ebx, esp
  loc_00E0D2EF: mov var_7C, 00000006h
  loc_00E0D2F6: mov edx, [eax]
  loc_00E0D2F8: push eax
  loc_00E0D2F9: mov [ebx], ecx
  loc_00E0D2FB: mov ecx, var_80
  loc_00E0D2FE: mov var_C4, eax
  loc_00E0D304: mov [ebx+00000004h], ecx
  loc_00E0D307: mov ecx, var_7C
  loc_00E0D30A: mov [ebx+00000008h], ecx
  loc_00E0D30D: mov ecx, var_78
  loc_00E0D310: mov [ebx+0000000Ch], ecx
  loc_00E0D313: call [edx+00000028h]
  loc_00E0D316: test eax, eax
  loc_00E0D318: fnclex
  loc_00E0D31A: jge 00E0D331h
  loc_00E0D31C: mov edx, var_C4
  loc_00E0D322: push 00000028h
  loc_00E0D324: push 006E09E8h
  loc_00E0D329: push edx
  loc_00E0D32A: push eax
  loc_00E0D32B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D331: mov eax, var_34
  loc_00E0D334: push 00000000h
  loc_00E0D336: mov var_CC, eax
  loc_00E0D33C: push 00000014h
  loc_00E0D33E: mov ebx, [eax]
  loc_00E0D340: mov eax, [esi]
  loc_00E0D342: push esi
  loc_00E0D343: call [eax+0000048Ch]
  loc_00E0D349: lea ecx, var_24
  loc_00E0D34C: push eax
  loc_00E0D34D: push ecx
  loc_00E0D34E: call edi
  loc_00E0D350: lea edx, var_54
  loc_00E0D353: push eax
  loc_00E0D354: push edx
  loc_00E0D355: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0D35B: mov edx, [eax]
  loc_00E0D35D: mov ecx, esp
  loc_00E0D35F: mov var_130, ebx
  loc_00E0D365: mov ebx, var_CC
  loc_00E0D36B: mov [ecx], edx
  loc_00E0D36D: mov edx, [eax+00000004h]
  loc_00E0D370: push ebx
  loc_00E0D371: mov [ecx+00000004h], edx
  loc_00E0D374: mov edx, [eax+00000008h]
  loc_00E0D377: mov eax, [eax+0000000Ch]
  loc_00E0D37A: mov [ecx+00000008h], edx
  loc_00E0D37D: mov [ecx+0000000Ch], eax
  loc_00E0D380: mov ecx, var_130
  loc_00E0D386: call [ecx+00000038h]
  loc_00E0D389: test eax, eax
  loc_00E0D38B: fnclex
  loc_00E0D38D: jge 00E0D3A2h
  loc_00E0D38F: push 00000038h
  loc_00E0D391: push 006E09F8h
  loc_00E0D396: push ebx
  loc_00E0D397: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D39D: push eax
  loc_00E0D39E: call ebx
  loc_00E0D3A0: jmp 00E0D3A8h
  loc_00E0D3A2: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0D3A8: lea edx, var_34
  loc_00E0D3AB: lea eax, var_30
  loc_00E0D3AE: push edx
  loc_00E0D3AF: lea ecx, var_2C
  loc_00E0D3B2: push eax
  loc_00E0D3B3: lea edx, var_28
  loc_00E0D3B6: push ecx
  loc_00E0D3B7: lea eax, var_24
  loc_00E0D3BA: push edx
  loc_00E0D3BB: push eax
  loc_00E0D3BC: push 00000005h
  loc_00E0D3BE: call [00401048h] ; __vbaFreeObjList
  loc_00E0D3C4: lea ecx, var_54
  loc_00E0D3C7: lea edx, var_44
  loc_00E0D3CA: push ecx
  loc_00E0D3CB: push edx
  loc_00E0D3CC: push 00000002h
  loc_00E0D3CE: call [00401038h] ; __vbaFreeVarList
  loc_00E0D3D4: mov eax, [esi]
  loc_00E0D3D6: add esp, 00000024h
  loc_00E0D3D9: push 006DCBD8h
  loc_00E0D3DE: push 00000000h
  loc_00E0D3E0: push 00000018h
  loc_00E0D3E2: push esi
  loc_00E0D3E3: call [eax+00000490h]
  loc_00E0D3E9: lea ecx, var_28
  loc_00E0D3EC: push eax
  loc_00E0D3ED: push ecx
  loc_00E0D3EE: call edi
  loc_00E0D3F0: lea edx, var_44
  loc_00E0D3F3: push eax
  loc_00E0D3F4: push edx
  loc_00E0D3F5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0D3FB: add esp, 00000010h
  loc_00E0D3FE: push eax
  loc_00E0D3FF: call [00401120h] ; __vbaCastObjVar
  loc_00E0D405: push eax
  loc_00E0D406: lea eax, var_2C
  loc_00E0D409: push eax
  loc_00E0D40A: call edi
  loc_00E0D40C: mov ecx, [eax]
  loc_00E0D40E: lea edx, var_30
  loc_00E0D411: push edx
  loc_00E0D412: push eax
  loc_00E0D413: mov var_C4, eax
  loc_00E0D419: call [ecx+00000054h]
  loc_00E0D41C: test eax, eax
  loc_00E0D41E: fnclex
  loc_00E0D420: jge 00E0D433h
  loc_00E0D422: mov ecx, var_C4
  loc_00E0D428: push 00000054h
  loc_00E0D42A: push 006DCBE8h
  loc_00E0D42F: push ecx
  loc_00E0D430: push eax
  loc_00E0D431: call ebx
  loc_00E0D433: mov ecx, var_30
  loc_00E0D436: mov eax, 00000002h
  loc_00E0D43B: mov var_7C, 00000007h
  loc_00E0D442: mov var_84, eax
  loc_00E0D448: mov edx, [ecx]
  loc_00E0D44A: mov var_CC, ecx
  loc_00E0D450: lea ecx, var_34
  loc_00E0D453: push ecx
  loc_00E0D454: sub esp, 00000010h
  loc_00E0D457: mov ecx, esp
  loc_00E0D459: mov [ecx], eax
  loc_00E0D45B: mov eax, var_80
  loc_00E0D45E: mov [ecx+00000004h], eax
  loc_00E0D461: mov eax, var_7C
  loc_00E0D464: mov [ecx+00000008h], eax
  loc_00E0D467: mov eax, var_78
  loc_00E0D46A: mov [ecx+0000000Ch], eax
  loc_00E0D46D: mov ecx, var_30
  loc_00E0D470: push ecx
  loc_00E0D471: call [edx+00000028h]
  loc_00E0D474: test eax, eax
  loc_00E0D476: fnclex
  loc_00E0D478: jge 00E0D48Bh
  loc_00E0D47A: mov edx, var_CC
  loc_00E0D480: push 00000028h
  loc_00E0D482: push 006E09E8h
  loc_00E0D487: push edx
  loc_00E0D488: push eax
  loc_00E0D489: call ebx
  loc_00E0D48B: mov eax, var_34
  loc_00E0D48E: mov ecx, [esi]
  loc_00E0D490: push esi
  loc_00E0D491: mov var_D4, eax
  loc_00E0D497: call [ecx+00000378h]
  loc_00E0D49D: lea edx, var_24
  loc_00E0D4A0: push eax
  loc_00E0D4A1: push edx
  loc_00E0D4A2: call edi
  loc_00E0D4A4: mov ecx, [eax]
  loc_00E0D4A6: lea edx, var_18
  loc_00E0D4A9: push edx
  loc_00E0D4AA: push eax
  loc_00E0D4AB: mov var_BC, eax
  loc_00E0D4B1: call [ecx+000000A8h]
  loc_00E0D4B7: test eax, eax
  loc_00E0D4B9: fnclex
  loc_00E0D4BB: jge 00E0D4D1h
  loc_00E0D4BD: mov ecx, var_BC
  loc_00E0D4C3: push 000000A8h
  loc_00E0D4C8: push 006E0400h
  loc_00E0D4CD: push ecx
  loc_00E0D4CE: push eax
  loc_00E0D4CF: call ebx
  loc_00E0D4D1: mov eax, var_18
  loc_00E0D4D4: mov edx, var_D4
  loc_00E0D4DA: mov var_4C, eax
  loc_00E0D4DD: mov eax, 00000008h
  loc_00E0D4E2: mov var_18, 00000000h
  loc_00E0D4E9: mov var_54, eax
  loc_00E0D4EC: mov ecx, [edx]
  loc_00E0D4EE: sub esp, 00000010h
  loc_00E0D4F1: mov edx, esp
  loc_00E0D4F3: mov [edx], eax
  loc_00E0D4F5: mov eax, var_50
  loc_00E0D4F8: mov [edx+00000004h], eax
  loc_00E0D4FB: mov eax, var_4C
  loc_00E0D4FE: mov [edx+00000008h], eax
  loc_00E0D501: mov eax, var_48
  loc_00E0D504: mov [edx+0000000Ch], eax
  loc_00E0D507: mov edx, var_D4
  loc_00E0D50D: push edx
  loc_00E0D50E: call [ecx+00000038h]
  loc_00E0D511: test eax, eax
  loc_00E0D513: fnclex
  loc_00E0D515: jge 00E0D528h
  loc_00E0D517: mov ecx, var_D4
  loc_00E0D51D: push 00000038h
  loc_00E0D51F: push 006E09F8h
  loc_00E0D524: push ecx
  loc_00E0D525: push eax
  loc_00E0D526: call ebx
  loc_00E0D528: lea edx, var_34
  loc_00E0D52B: lea eax, var_30
  loc_00E0D52E: push edx
  loc_00E0D52F: lea ecx, var_2C
  loc_00E0D532: push eax
  loc_00E0D533: lea edx, var_28
  loc_00E0D536: push ecx
  loc_00E0D537: lea eax, var_24
  loc_00E0D53A: push edx
  loc_00E0D53B: push eax
  loc_00E0D53C: push 00000005h
  loc_00E0D53E: call [00401048h] ; __vbaFreeObjList
  loc_00E0D544: lea ecx, var_54
  loc_00E0D547: lea edx, var_44
  loc_00E0D54A: push ecx
  loc_00E0D54B: push edx
  loc_00E0D54C: push 00000002h
  loc_00E0D54E: call [00401038h] ; __vbaFreeVarList
  loc_00E0D554: mov eax, [esi]
  loc_00E0D556: add esp, 00000024h
  loc_00E0D559: push 006DCBD8h
  loc_00E0D55E: push 00000000h
  loc_00E0D560: push 00000018h
  loc_00E0D562: push esi
  loc_00E0D563: call [eax+00000490h]
  loc_00E0D569: lea ecx, var_28
  loc_00E0D56C: push eax
  loc_00E0D56D: push ecx
  loc_00E0D56E: call edi
  loc_00E0D570: lea edx, var_44
  loc_00E0D573: push eax
  loc_00E0D574: push edx
  loc_00E0D575: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0D57B: add esp, 00000010h
  loc_00E0D57E: push eax
  loc_00E0D57F: call [00401120h] ; __vbaCastObjVar
  loc_00E0D585: push eax
  loc_00E0D586: lea eax, var_2C
  loc_00E0D589: push eax
  loc_00E0D58A: call edi
  loc_00E0D58C: mov ecx, [eax]
  loc_00E0D58E: lea edx, var_30
  loc_00E0D591: push edx
  loc_00E0D592: push eax
  loc_00E0D593: mov var_C4, eax
  loc_00E0D599: call [ecx+00000054h]
  loc_00E0D59C: test eax, eax
  loc_00E0D59E: fnclex
  loc_00E0D5A0: jge 00E0D5B3h
  loc_00E0D5A2: mov ecx, var_C4
  loc_00E0D5A8: push 00000054h
  loc_00E0D5AA: push 006DCBE8h
  loc_00E0D5AF: push ecx
  loc_00E0D5B0: push eax
  loc_00E0D5B1: call ebx
  loc_00E0D5B3: mov ecx, var_30
  loc_00E0D5B6: mov eax, 00000002h
  loc_00E0D5BB: mov var_7C, 00000008h
  loc_00E0D5C2: mov var_84, eax
  loc_00E0D5C8: mov edx, [ecx]
  loc_00E0D5CA: mov var_CC, ecx
  loc_00E0D5D0: lea ecx, var_34
  loc_00E0D5D3: push ecx
  loc_00E0D5D4: sub esp, 00000010h
  loc_00E0D5D7: mov ecx, esp
  loc_00E0D5D9: mov [ecx], eax
  loc_00E0D5DB: mov eax, var_80
  loc_00E0D5DE: mov [ecx+00000004h], eax
  loc_00E0D5E1: mov eax, var_7C
  loc_00E0D5E4: mov [ecx+00000008h], eax
  loc_00E0D5E7: mov eax, var_78
  loc_00E0D5EA: mov [ecx+0000000Ch], eax
  loc_00E0D5ED: mov ecx, var_30
  loc_00E0D5F0: push ecx
  loc_00E0D5F1: call [edx+00000028h]
  loc_00E0D5F4: test eax, eax
  loc_00E0D5F6: fnclex
  loc_00E0D5F8: jge 00E0D60Bh
  loc_00E0D5FA: mov edx, var_CC
  loc_00E0D600: push 00000028h
  loc_00E0D602: push 006E09E8h
  loc_00E0D607: push edx
  loc_00E0D608: push eax
  loc_00E0D609: call ebx
  loc_00E0D60B: mov eax, var_34
  loc_00E0D60E: mov ecx, [esi]
  loc_00E0D610: push esi
  loc_00E0D611: mov var_D4, eax
  loc_00E0D617: call [ecx+00000384h]
  loc_00E0D61D: lea edx, var_24
  loc_00E0D620: push eax
  loc_00E0D621: push edx
  loc_00E0D622: call edi
  loc_00E0D624: mov ecx, [eax]
  loc_00E0D626: lea edx, var_18
  loc_00E0D629: push edx
  loc_00E0D62A: push eax
  loc_00E0D62B: mov var_BC, eax
  loc_00E0D631: call [ecx+000000A0h]
  loc_00E0D637: test eax, eax
  loc_00E0D639: fnclex
  loc_00E0D63B: jge 00E0D651h
  loc_00E0D63D: mov ecx, var_BC
  loc_00E0D643: push 000000A0h
  loc_00E0D648: push 006DCB70h
  loc_00E0D64D: push ecx
  loc_00E0D64E: push eax
  loc_00E0D64F: call ebx
  loc_00E0D651: mov eax, var_18
  loc_00E0D654: mov edx, var_D4
  loc_00E0D65A: mov var_4C, eax
  loc_00E0D65D: mov eax, 00000008h
  loc_00E0D662: mov var_18, 00000000h
  loc_00E0D669: mov var_54, eax
  loc_00E0D66C: mov ecx, [edx]
  loc_00E0D66E: sub esp, 00000010h
  loc_00E0D671: mov edx, esp
  loc_00E0D673: mov [edx], eax
  loc_00E0D675: mov eax, var_50
  loc_00E0D678: mov [edx+00000004h], eax
  loc_00E0D67B: mov eax, var_4C
  loc_00E0D67E: mov [edx+00000008h], eax
  loc_00E0D681: mov eax, var_48
  loc_00E0D684: mov [edx+0000000Ch], eax
  loc_00E0D687: mov edx, var_D4
  loc_00E0D68D: push edx
  loc_00E0D68E: call [ecx+00000038h]
  loc_00E0D691: test eax, eax
  loc_00E0D693: fnclex
  loc_00E0D695: jge 00E0D6A8h
  loc_00E0D697: mov ecx, var_D4
  loc_00E0D69D: push 00000038h
  loc_00E0D69F: push 006E09F8h
  loc_00E0D6A4: push ecx
  loc_00E0D6A5: push eax
  loc_00E0D6A6: call ebx
  loc_00E0D6A8: lea edx, var_34
  loc_00E0D6AB: lea eax, var_30
  loc_00E0D6AE: push edx
  loc_00E0D6AF: lea ecx, var_2C
  loc_00E0D6B2: push eax
  loc_00E0D6B3: lea edx, var_28
  loc_00E0D6B6: push ecx
  loc_00E0D6B7: lea eax, var_24
  loc_00E0D6BA: push edx
  loc_00E0D6BB: push eax
  loc_00E0D6BC: push 00000005h
  loc_00E0D6BE: call [00401048h] ; __vbaFreeObjList
  loc_00E0D6C4: lea ecx, var_54
  loc_00E0D6C7: lea edx, var_44
  loc_00E0D6CA: push ecx
  loc_00E0D6CB: push edx
  loc_00E0D6CC: push 00000002h
  loc_00E0D6CE: call [00401038h] ; __vbaFreeVarList
  loc_00E0D6D4: mov eax, [esi]
  loc_00E0D6D6: add esp, 00000024h
  loc_00E0D6D9: push 006DCBD8h
  loc_00E0D6DE: push 00000000h
  loc_00E0D6E0: push 00000018h
  loc_00E0D6E2: push esi
  loc_00E0D6E3: call [eax+00000490h]
  loc_00E0D6E9: lea ecx, var_28
  loc_00E0D6EC: push eax
  loc_00E0D6ED: push ecx
  loc_00E0D6EE: call edi
  loc_00E0D6F0: lea edx, var_44
  loc_00E0D6F3: push eax
  loc_00E0D6F4: push edx
  loc_00E0D6F5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0D6FB: add esp, 00000010h
  loc_00E0D6FE: push eax
  loc_00E0D6FF: call [00401120h] ; __vbaCastObjVar
  loc_00E0D705: push eax
  loc_00E0D706: lea eax, var_2C
  loc_00E0D709: push eax
  loc_00E0D70A: call edi
  loc_00E0D70C: mov ecx, [eax]
  loc_00E0D70E: lea edx, var_30
  loc_00E0D711: push edx
  loc_00E0D712: push eax
  loc_00E0D713: mov var_C4, eax
  loc_00E0D719: call [ecx+00000054h]
  loc_00E0D71C: test eax, eax
  loc_00E0D71E: fnclex
  loc_00E0D720: jge 00E0D733h
  loc_00E0D722: mov ecx, var_C4
  loc_00E0D728: push 00000054h
  loc_00E0D72A: push 006DCBE8h
  loc_00E0D72F: push ecx
  loc_00E0D730: push eax
  loc_00E0D731: call ebx
  loc_00E0D733: mov ecx, var_30
  loc_00E0D736: mov eax, 00000002h
  loc_00E0D73B: mov var_7C, 00000009h
  loc_00E0D742: mov var_84, eax
  loc_00E0D748: mov edx, [ecx]
  loc_00E0D74A: mov var_CC, ecx
  loc_00E0D750: lea ecx, var_34
  loc_00E0D753: push ecx
  loc_00E0D754: sub esp, 00000010h
  loc_00E0D757: mov ecx, esp
  loc_00E0D759: mov [ecx], eax
  loc_00E0D75B: mov eax, var_80
  loc_00E0D75E: mov [ecx+00000004h], eax
  loc_00E0D761: mov eax, var_7C
  loc_00E0D764: mov [ecx+00000008h], eax
  loc_00E0D767: mov eax, var_78
  loc_00E0D76A: mov [ecx+0000000Ch], eax
  loc_00E0D76D: mov ecx, var_30
  loc_00E0D770: push ecx
  loc_00E0D771: call [edx+00000028h]
  loc_00E0D774: test eax, eax
  loc_00E0D776: fnclex
  loc_00E0D778: jge 00E0D78Bh
  loc_00E0D77A: mov edx, var_CC
  loc_00E0D780: push 00000028h
  loc_00E0D782: push 006E09E8h
  loc_00E0D787: push edx
  loc_00E0D788: push eax
  loc_00E0D789: call ebx
  loc_00E0D78B: mov eax, var_34
  loc_00E0D78E: mov ecx, [esi]
  loc_00E0D790: push esi
  loc_00E0D791: mov var_D4, eax
  loc_00E0D797: call [ecx+00000374h]
  loc_00E0D79D: lea edx, var_24
  loc_00E0D7A0: push eax
  loc_00E0D7A1: push edx
  loc_00E0D7A2: call edi
  loc_00E0D7A4: mov ecx, [eax]
  loc_00E0D7A6: lea edx, var_18
  loc_00E0D7A9: push edx
  loc_00E0D7AA: push eax
  loc_00E0D7AB: mov var_BC, eax
  loc_00E0D7B1: call [ecx+000000A0h]
  loc_00E0D7B7: test eax, eax
  loc_00E0D7B9: fnclex
  loc_00E0D7BB: jge 00E0D7D1h
  loc_00E0D7BD: mov ecx, var_BC
  loc_00E0D7C3: push 000000A0h
  loc_00E0D7C8: push 006DCB70h
  loc_00E0D7CD: push ecx
  loc_00E0D7CE: push eax
  loc_00E0D7CF: call ebx
  loc_00E0D7D1: mov eax, var_18
  loc_00E0D7D4: mov edx, var_D4
  loc_00E0D7DA: mov var_4C, eax
  loc_00E0D7DD: mov eax, 00000008h
  loc_00E0D7E2: mov var_18, 00000000h
  loc_00E0D7E9: mov var_54, eax
  loc_00E0D7EC: mov ecx, [edx]
  loc_00E0D7EE: sub esp, 00000010h
  loc_00E0D7F1: mov edx, esp
  loc_00E0D7F3: mov [edx], eax
  loc_00E0D7F5: mov eax, var_50
  loc_00E0D7F8: mov [edx+00000004h], eax
  loc_00E0D7FB: mov eax, var_4C
  loc_00E0D7FE: mov [edx+00000008h], eax
  loc_00E0D801: mov eax, var_48
  loc_00E0D804: mov [edx+0000000Ch], eax
  loc_00E0D807: mov edx, var_D4
  loc_00E0D80D: push edx
  loc_00E0D80E: call [ecx+00000038h]
  loc_00E0D811: test eax, eax
  loc_00E0D813: fnclex
  loc_00E0D815: jge 00E0D828h
  loc_00E0D817: mov ecx, var_D4
  loc_00E0D81D: push 00000038h
  loc_00E0D81F: push 006E09F8h
  loc_00E0D824: push ecx
  loc_00E0D825: push eax
  loc_00E0D826: call ebx
  loc_00E0D828: lea edx, var_34
  loc_00E0D82B: lea eax, var_30
  loc_00E0D82E: push edx
  loc_00E0D82F: lea ecx, var_2C
  loc_00E0D832: push eax
  loc_00E0D833: lea edx, var_28
  loc_00E0D836: push ecx
  loc_00E0D837: lea eax, var_24
  loc_00E0D83A: push edx
  loc_00E0D83B: push eax
  loc_00E0D83C: push 00000005h
  loc_00E0D83E: call [00401048h] ; __vbaFreeObjList
  loc_00E0D844: lea ecx, var_54
  loc_00E0D847: lea edx, var_44
  loc_00E0D84A: push ecx
  loc_00E0D84B: push edx
  loc_00E0D84C: push 00000002h
  loc_00E0D84E: call [00401038h] ; __vbaFreeVarList
  loc_00E0D854: mov eax, [esi]
  loc_00E0D856: add esp, 00000024h
  loc_00E0D859: push 006DCBD8h
  loc_00E0D85E: push 00000000h
  loc_00E0D860: push 00000018h
  loc_00E0D862: push esi
  loc_00E0D863: call [eax+00000490h]
  loc_00E0D869: lea ecx, var_28
  loc_00E0D86C: push eax
  loc_00E0D86D: push ecx
  loc_00E0D86E: call edi
  loc_00E0D870: lea edx, var_44
  loc_00E0D873: push eax
  loc_00E0D874: push edx
  loc_00E0D875: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0D87B: add esp, 00000010h
  loc_00E0D87E: push eax
  loc_00E0D87F: call [00401120h] ; __vbaCastObjVar
  loc_00E0D885: push eax
  loc_00E0D886: lea eax, var_2C
  loc_00E0D889: push eax
  loc_00E0D88A: call edi
  loc_00E0D88C: mov ecx, [eax]
  loc_00E0D88E: lea edx, var_30
  loc_00E0D891: push edx
  loc_00E0D892: push eax
  loc_00E0D893: mov var_C4, eax
  loc_00E0D899: call [ecx+00000054h]
  loc_00E0D89C: test eax, eax
  loc_00E0D89E: fnclex
  loc_00E0D8A0: jge 00E0D8B3h
  loc_00E0D8A2: mov ecx, var_C4
  loc_00E0D8A8: push 00000054h
  loc_00E0D8AA: push 006DCBE8h
  loc_00E0D8AF: push ecx
  loc_00E0D8B0: push eax
  loc_00E0D8B1: call ebx
  loc_00E0D8B3: mov ecx, var_30
  loc_00E0D8B6: mov eax, 00000002h
  loc_00E0D8BB: mov var_7C, 0000000Ah
  loc_00E0D8C2: mov var_84, eax
  loc_00E0D8C8: mov edx, [ecx]
  loc_00E0D8CA: mov var_CC, ecx
  loc_00E0D8D0: lea ecx, var_34
  loc_00E0D8D3: push ecx
  loc_00E0D8D4: sub esp, 00000010h
  loc_00E0D8D7: mov ecx, esp
  loc_00E0D8D9: mov [ecx], eax
  loc_00E0D8DB: mov eax, var_80
  loc_00E0D8DE: mov [ecx+00000004h], eax
  loc_00E0D8E1: mov eax, var_7C
  loc_00E0D8E4: mov [ecx+00000008h], eax
  loc_00E0D8E7: mov eax, var_78
  loc_00E0D8EA: mov [ecx+0000000Ch], eax
  loc_00E0D8ED: mov ecx, var_30
  loc_00E0D8F0: push ecx
  loc_00E0D8F1: call [edx+00000028h]
  loc_00E0D8F4: test eax, eax
  loc_00E0D8F6: fnclex
  loc_00E0D8F8: jge 00E0D90Bh
  loc_00E0D8FA: mov edx, var_CC
  loc_00E0D900: push 00000028h
  loc_00E0D902: push 006E09E8h
  loc_00E0D907: push edx
  loc_00E0D908: push eax
  loc_00E0D909: call ebx
  loc_00E0D90B: mov eax, var_34
  loc_00E0D90E: mov ecx, [esi]
  loc_00E0D910: push esi
  loc_00E0D911: mov var_D4, eax
  loc_00E0D917: call [ecx+00000370h]
  loc_00E0D91D: lea edx, var_24
  loc_00E0D920: push eax
  loc_00E0D921: push edx
  loc_00E0D922: call edi
  loc_00E0D924: mov ecx, [eax]
  loc_00E0D926: lea edx, var_18
  loc_00E0D929: push edx
  loc_00E0D92A: push eax
  loc_00E0D92B: mov var_BC, eax
  loc_00E0D931: call [ecx+000000A0h]
  loc_00E0D937: test eax, eax
  loc_00E0D939: fnclex
  loc_00E0D93B: jge 00E0D951h
  loc_00E0D93D: mov ecx, var_BC
  loc_00E0D943: push 000000A0h
  loc_00E0D948: push 006DCB70h
  loc_00E0D94D: push ecx
  loc_00E0D94E: push eax
  loc_00E0D94F: call ebx
  loc_00E0D951: mov eax, var_18
  loc_00E0D954: mov edx, var_D4
  loc_00E0D95A: mov var_4C, eax
  loc_00E0D95D: mov eax, 00000008h
  loc_00E0D962: mov var_18, 00000000h
  loc_00E0D969: mov var_54, eax
  loc_00E0D96C: mov ecx, [edx]
  loc_00E0D96E: sub esp, 00000010h
  loc_00E0D971: mov edx, esp
  loc_00E0D973: mov [edx], eax
  loc_00E0D975: mov eax, var_50
  loc_00E0D978: mov [edx+00000004h], eax
  loc_00E0D97B: mov eax, var_4C
  loc_00E0D97E: mov [edx+00000008h], eax
  loc_00E0D981: mov eax, var_48
  loc_00E0D984: mov [edx+0000000Ch], eax
  loc_00E0D987: mov edx, var_D4
  loc_00E0D98D: push edx
  loc_00E0D98E: call [ecx+00000038h]
  loc_00E0D991: test eax, eax
  loc_00E0D993: fnclex
  loc_00E0D995: jge 00E0D9A8h
  loc_00E0D997: mov ecx, var_D4
  loc_00E0D99D: push 00000038h
  loc_00E0D99F: push 006E09F8h
  loc_00E0D9A4: push ecx
  loc_00E0D9A5: push eax
  loc_00E0D9A6: call ebx
  loc_00E0D9A8: lea edx, var_34
  loc_00E0D9AB: lea eax, var_30
  loc_00E0D9AE: push edx
  loc_00E0D9AF: lea ecx, var_2C
  loc_00E0D9B2: push eax
  loc_00E0D9B3: lea edx, var_28
  loc_00E0D9B6: push ecx
  loc_00E0D9B7: lea eax, var_24
  loc_00E0D9BA: push edx
  loc_00E0D9BB: push eax
  loc_00E0D9BC: push 00000005h
  loc_00E0D9BE: call [00401048h] ; __vbaFreeObjList
  loc_00E0D9C4: lea ecx, var_54
  loc_00E0D9C7: lea edx, var_44
  loc_00E0D9CA: push ecx
  loc_00E0D9CB: push edx
  loc_00E0D9CC: push 00000002h
  loc_00E0D9CE: call [00401038h] ; __vbaFreeVarList
  loc_00E0D9D4: mov eax, [esi]
  loc_00E0D9D6: add esp, 00000024h
  loc_00E0D9D9: push 006DCBD8h
  loc_00E0D9DE: push 00000000h
  loc_00E0D9E0: push 00000018h
  loc_00E0D9E2: push esi
  loc_00E0D9E3: call [eax+00000490h]
  loc_00E0D9E9: lea ecx, var_28
  loc_00E0D9EC: push eax
  loc_00E0D9ED: push ecx
  loc_00E0D9EE: call edi
  loc_00E0D9F0: lea edx, var_44
  loc_00E0D9F3: push eax
  loc_00E0D9F4: push edx
  loc_00E0D9F5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0D9FB: add esp, 00000010h
  loc_00E0D9FE: push eax
  loc_00E0D9FF: call [00401120h] ; __vbaCastObjVar
  loc_00E0DA05: push eax
  loc_00E0DA06: lea eax, var_2C
  loc_00E0DA09: push eax
  loc_00E0DA0A: call edi
  loc_00E0DA0C: mov ecx, [eax]
  loc_00E0DA0E: lea edx, var_30
  loc_00E0DA11: push edx
  loc_00E0DA12: push eax
  loc_00E0DA13: mov var_C4, eax
  loc_00E0DA19: call [ecx+00000054h]
  loc_00E0DA1C: test eax, eax
  loc_00E0DA1E: fnclex
  loc_00E0DA20: jge 00E0DA33h
  loc_00E0DA22: mov ecx, var_C4
  loc_00E0DA28: push 00000054h
  loc_00E0DA2A: push 006DCBE8h
  loc_00E0DA2F: push ecx
  loc_00E0DA30: push eax
  loc_00E0DA31: call ebx
  loc_00E0DA33: mov ecx, var_30
  loc_00E0DA36: mov eax, 00000002h
  loc_00E0DA3B: mov var_7C, 0000000Bh
  loc_00E0DA42: mov var_84, eax
  loc_00E0DA48: mov edx, [ecx]
  loc_00E0DA4A: mov var_CC, ecx
  loc_00E0DA50: lea ecx, var_34
  loc_00E0DA53: push ecx
  loc_00E0DA54: sub esp, 00000010h
  loc_00E0DA57: mov ecx, esp
  loc_00E0DA59: mov [ecx], eax
  loc_00E0DA5B: mov eax, var_80
  loc_00E0DA5E: mov [ecx+00000004h], eax
  loc_00E0DA61: mov eax, var_7C
  loc_00E0DA64: mov [ecx+00000008h], eax
  loc_00E0DA67: mov eax, var_78
  loc_00E0DA6A: mov [ecx+0000000Ch], eax
  loc_00E0DA6D: mov ecx, var_30
  loc_00E0DA70: push ecx
  loc_00E0DA71: call [edx+00000028h]
  loc_00E0DA74: test eax, eax
  loc_00E0DA76: fnclex
  loc_00E0DA78: jge 00E0DA8Bh
  loc_00E0DA7A: mov edx, var_CC
  loc_00E0DA80: push 00000028h
  loc_00E0DA82: push 006E09E8h
  loc_00E0DA87: push edx
  loc_00E0DA88: push eax
  loc_00E0DA89: call ebx
  loc_00E0DA8B: mov eax, var_34
  loc_00E0DA8E: mov ecx, [esi]
  loc_00E0DA90: push esi
  loc_00E0DA91: mov var_D4, eax
  loc_00E0DA97: call [ecx+0000036Ch]
  loc_00E0DA9D: lea edx, var_24
  loc_00E0DAA0: push eax
  loc_00E0DAA1: push edx
  loc_00E0DAA2: call edi
  loc_00E0DAA4: mov ecx, [eax]
  loc_00E0DAA6: lea edx, var_18
  loc_00E0DAA9: push edx
  loc_00E0DAAA: push eax
  loc_00E0DAAB: mov var_BC, eax
  loc_00E0DAB1: call [ecx+000000A0h]
  loc_00E0DAB7: test eax, eax
  loc_00E0DAB9: fnclex
  loc_00E0DABB: jge 00E0DAD1h
  loc_00E0DABD: mov ecx, var_BC
  loc_00E0DAC3: push 000000A0h
  loc_00E0DAC8: push 006DCB70h
  loc_00E0DACD: push ecx
  loc_00E0DACE: push eax
  loc_00E0DACF: call ebx
  loc_00E0DAD1: mov eax, var_18
  loc_00E0DAD4: mov edx, var_D4
  loc_00E0DADA: mov var_4C, eax
  loc_00E0DADD: mov eax, 00000008h
  loc_00E0DAE2: mov var_18, 00000000h
  loc_00E0DAE9: mov var_54, eax
  loc_00E0DAEC: mov ecx, [edx]
  loc_00E0DAEE: sub esp, 00000010h
  loc_00E0DAF1: mov edx, esp
  loc_00E0DAF3: mov [edx], eax
  loc_00E0DAF5: mov eax, var_50
  loc_00E0DAF8: mov [edx+00000004h], eax
  loc_00E0DAFB: mov eax, var_4C
  loc_00E0DAFE: mov [edx+00000008h], eax
  loc_00E0DB01: mov eax, var_48
  loc_00E0DB04: mov [edx+0000000Ch], eax
  loc_00E0DB07: mov edx, var_D4
  loc_00E0DB0D: push edx
  loc_00E0DB0E: call [ecx+00000038h]
  loc_00E0DB11: test eax, eax
  loc_00E0DB13: fnclex
  loc_00E0DB15: jge 00E0DB28h
  loc_00E0DB17: mov ecx, var_D4
  loc_00E0DB1D: push 00000038h
  loc_00E0DB1F: push 006E09F8h
  loc_00E0DB24: push ecx
  loc_00E0DB25: push eax
  loc_00E0DB26: call ebx
  loc_00E0DB28: lea edx, var_34
  loc_00E0DB2B: lea eax, var_30
  loc_00E0DB2E: push edx
  loc_00E0DB2F: lea ecx, var_2C
  loc_00E0DB32: push eax
  loc_00E0DB33: lea edx, var_28
  loc_00E0DB36: push ecx
  loc_00E0DB37: lea eax, var_24
  loc_00E0DB3A: push edx
  loc_00E0DB3B: push eax
  loc_00E0DB3C: push 00000005h
  loc_00E0DB3E: call [00401048h] ; __vbaFreeObjList
  loc_00E0DB44: lea ecx, var_54
  loc_00E0DB47: lea edx, var_44
  loc_00E0DB4A: push ecx
  loc_00E0DB4B: push edx
  loc_00E0DB4C: push 00000002h
  loc_00E0DB4E: call [00401038h] ; __vbaFreeVarList
  loc_00E0DB54: mov eax, [esi]
  loc_00E0DB56: add esp, 00000024h
  loc_00E0DB59: push 006DCBD8h
  loc_00E0DB5E: push 00000000h
  loc_00E0DB60: push 00000018h
  loc_00E0DB62: push esi
  loc_00E0DB63: call [eax+00000490h]
  loc_00E0DB69: lea ecx, var_28
  loc_00E0DB6C: push eax
  loc_00E0DB6D: push ecx
  loc_00E0DB6E: call edi
  loc_00E0DB70: lea edx, var_44
  loc_00E0DB73: push eax
  loc_00E0DB74: push edx
  loc_00E0DB75: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0DB7B: add esp, 00000010h
  loc_00E0DB7E: push eax
  loc_00E0DB7F: call [00401120h] ; __vbaCastObjVar
  loc_00E0DB85: push eax
  loc_00E0DB86: lea eax, var_2C
  loc_00E0DB89: push eax
  loc_00E0DB8A: call edi
  loc_00E0DB8C: mov ecx, [eax]
  loc_00E0DB8E: lea edx, var_30
  loc_00E0DB91: push edx
  loc_00E0DB92: push eax
  loc_00E0DB93: mov var_C4, eax
  loc_00E0DB99: call [ecx+00000054h]
  loc_00E0DB9C: test eax, eax
  loc_00E0DB9E: fnclex
  loc_00E0DBA0: jge 00E0DBB3h
  loc_00E0DBA2: mov ecx, var_C4
  loc_00E0DBA8: push 00000054h
  loc_00E0DBAA: push 006DCBE8h
  loc_00E0DBAF: push ecx
  loc_00E0DBB0: push eax
  loc_00E0DBB1: call ebx
  loc_00E0DBB3: mov ecx, var_30
  loc_00E0DBB6: mov eax, 00000002h
  loc_00E0DBBB: mov var_7C, 0000000Ch
  loc_00E0DBC2: mov var_84, eax
  loc_00E0DBC8: mov edx, [ecx]
  loc_00E0DBCA: mov var_CC, ecx
  loc_00E0DBD0: lea ecx, var_34
  loc_00E0DBD3: push ecx
  loc_00E0DBD4: sub esp, 00000010h
  loc_00E0DBD7: mov ecx, esp
  loc_00E0DBD9: mov [ecx], eax
  loc_00E0DBDB: mov eax, var_80
  loc_00E0DBDE: mov [ecx+00000004h], eax
  loc_00E0DBE1: mov eax, var_7C
  loc_00E0DBE4: mov [ecx+00000008h], eax
  loc_00E0DBE7: mov eax, var_78
  loc_00E0DBEA: mov [ecx+0000000Ch], eax
  loc_00E0DBED: mov ecx, var_30
  loc_00E0DBF0: push ecx
  loc_00E0DBF1: call [edx+00000028h]
  loc_00E0DBF4: test eax, eax
  loc_00E0DBF6: fnclex
  loc_00E0DBF8: jge 00E0DC0Bh
  loc_00E0DBFA: mov edx, var_CC
  loc_00E0DC00: push 00000028h
  loc_00E0DC02: push 006E09E8h
  loc_00E0DC07: push edx
  loc_00E0DC08: push eax
  loc_00E0DC09: call ebx
  loc_00E0DC0B: mov eax, var_34
  loc_00E0DC0E: mov ecx, [esi]
  loc_00E0DC10: push esi
  loc_00E0DC11: mov var_D4, eax
  loc_00E0DC17: call [ecx+00000368h]
  loc_00E0DC1D: lea edx, var_24
  loc_00E0DC20: push eax
  loc_00E0DC21: push edx
  loc_00E0DC22: call edi
  loc_00E0DC24: mov ecx, [eax]
  loc_00E0DC26: lea edx, var_18
  loc_00E0DC29: push edx
  loc_00E0DC2A: push eax
  loc_00E0DC2B: mov var_BC, eax
  loc_00E0DC31: call [ecx+000000A0h]
  loc_00E0DC37: test eax, eax
  loc_00E0DC39: fnclex
  loc_00E0DC3B: jge 00E0DC51h
  loc_00E0DC3D: mov ecx, var_BC
  loc_00E0DC43: push 000000A0h
  loc_00E0DC48: push 006DCB70h
  loc_00E0DC4D: push ecx
  loc_00E0DC4E: push eax
  loc_00E0DC4F: call ebx
  loc_00E0DC51: mov eax, var_18
  loc_00E0DC54: mov edx, var_D4
  loc_00E0DC5A: mov var_4C, eax
  loc_00E0DC5D: mov eax, 00000008h
  loc_00E0DC62: mov var_18, 00000000h
  loc_00E0DC69: mov var_54, eax
  loc_00E0DC6C: mov ecx, [edx]
  loc_00E0DC6E: sub esp, 00000010h
  loc_00E0DC71: mov edx, esp
  loc_00E0DC73: mov [edx], eax
  loc_00E0DC75: mov eax, var_50
  loc_00E0DC78: mov [edx+00000004h], eax
  loc_00E0DC7B: mov eax, var_4C
  loc_00E0DC7E: mov [edx+00000008h], eax
  loc_00E0DC81: mov eax, var_48
  loc_00E0DC84: mov [edx+0000000Ch], eax
  loc_00E0DC87: mov edx, var_D4
  loc_00E0DC8D: push edx
  loc_00E0DC8E: call [ecx+00000038h]
  loc_00E0DC91: test eax, eax
  loc_00E0DC93: fnclex
  loc_00E0DC95: jge 00E0DCA8h
  loc_00E0DC97: mov ecx, var_D4
  loc_00E0DC9D: push 00000038h
  loc_00E0DC9F: push 006E09F8h
  loc_00E0DCA4: push ecx
  loc_00E0DCA5: push eax
  loc_00E0DCA6: call ebx
  loc_00E0DCA8: lea edx, var_34
  loc_00E0DCAB: lea eax, var_30
  loc_00E0DCAE: push edx
  loc_00E0DCAF: lea ecx, var_2C
  loc_00E0DCB2: push eax
  loc_00E0DCB3: lea edx, var_28
  loc_00E0DCB6: push ecx
  loc_00E0DCB7: lea eax, var_24
  loc_00E0DCBA: push edx
  loc_00E0DCBB: push eax
  loc_00E0DCBC: push 00000005h
  loc_00E0DCBE: call [00401048h] ; __vbaFreeObjList
  loc_00E0DCC4: lea ecx, var_54
  loc_00E0DCC7: lea edx, var_44
  loc_00E0DCCA: push ecx
  loc_00E0DCCB: push edx
  loc_00E0DCCC: push 00000002h
  loc_00E0DCCE: call [00401038h] ; __vbaFreeVarList
  loc_00E0DCD4: mov eax, [esi]
  loc_00E0DCD6: add esp, 00000024h
  loc_00E0DCD9: push 006DCBD8h
  loc_00E0DCDE: push 00000000h
  loc_00E0DCE0: push 00000018h
  loc_00E0DCE2: push esi
  loc_00E0DCE3: call [eax+00000490h]
  loc_00E0DCE9: lea ecx, var_28
  loc_00E0DCEC: push eax
  loc_00E0DCED: push ecx
  loc_00E0DCEE: call edi
  loc_00E0DCF0: lea edx, var_44
  loc_00E0DCF3: push eax
  loc_00E0DCF4: push edx
  loc_00E0DCF5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0DCFB: add esp, 00000010h
  loc_00E0DCFE: push eax
  loc_00E0DCFF: call [00401120h] ; __vbaCastObjVar
  loc_00E0DD05: push eax
  loc_00E0DD06: lea eax, var_2C
  loc_00E0DD09: push eax
  loc_00E0DD0A: call edi
  loc_00E0DD0C: mov ecx, [eax]
  loc_00E0DD0E: lea edx, var_30
  loc_00E0DD11: push edx
  loc_00E0DD12: push eax
  loc_00E0DD13: mov var_C4, eax
  loc_00E0DD19: call [ecx+00000054h]
  loc_00E0DD1C: test eax, eax
  loc_00E0DD1E: fnclex
  loc_00E0DD20: jge 00E0DD33h
  loc_00E0DD22: mov ecx, var_C4
  loc_00E0DD28: push 00000054h
  loc_00E0DD2A: push 006DCBE8h
  loc_00E0DD2F: push ecx
  loc_00E0DD30: push eax
  loc_00E0DD31: call ebx
  loc_00E0DD33: mov ecx, var_30
  loc_00E0DD36: mov eax, 00000002h
  loc_00E0DD3B: mov var_7C, 0000000Dh
  loc_00E0DD42: mov var_84, eax
  loc_00E0DD48: mov edx, [ecx]
  loc_00E0DD4A: mov var_CC, ecx
  loc_00E0DD50: lea ecx, var_34
  loc_00E0DD53: push ecx
  loc_00E0DD54: sub esp, 00000010h
  loc_00E0DD57: mov ecx, esp
  loc_00E0DD59: mov [ecx], eax
  loc_00E0DD5B: mov eax, var_80
  loc_00E0DD5E: mov [ecx+00000004h], eax
  loc_00E0DD61: mov eax, var_7C
  loc_00E0DD64: mov [ecx+00000008h], eax
  loc_00E0DD67: mov eax, var_78
  loc_00E0DD6A: mov [ecx+0000000Ch], eax
  loc_00E0DD6D: mov ecx, var_30
  loc_00E0DD70: push ecx
  loc_00E0DD71: call [edx+00000028h]
  loc_00E0DD74: test eax, eax
  loc_00E0DD76: fnclex
  loc_00E0DD78: jge 00E0DD8Bh
  loc_00E0DD7A: mov edx, var_CC
  loc_00E0DD80: push 00000028h
  loc_00E0DD82: push 006E09E8h
  loc_00E0DD87: push edx
  loc_00E0DD88: push eax
  loc_00E0DD89: call ebx
  loc_00E0DD8B: mov eax, var_34
  loc_00E0DD8E: mov ecx, [esi]
  loc_00E0DD90: push esi
  loc_00E0DD91: mov var_D4, eax
  loc_00E0DD97: call [ecx+00000364h]
  loc_00E0DD9D: lea edx, var_24
  loc_00E0DDA0: push eax
  loc_00E0DDA1: push edx
  loc_00E0DDA2: call edi
  loc_00E0DDA4: mov ecx, [eax]
  loc_00E0DDA6: lea edx, var_18
  loc_00E0DDA9: push edx
  loc_00E0DDAA: push eax
  loc_00E0DDAB: mov var_BC, eax
  loc_00E0DDB1: call [ecx+000000A0h]
  loc_00E0DDB7: test eax, eax
  loc_00E0DDB9: fnclex
  loc_00E0DDBB: jge 00E0DDD1h
  loc_00E0DDBD: mov ecx, var_BC
  loc_00E0DDC3: push 000000A0h
  loc_00E0DDC8: push 006DCB70h
  loc_00E0DDCD: push ecx
  loc_00E0DDCE: push eax
  loc_00E0DDCF: call ebx
  loc_00E0DDD1: mov eax, var_18
  loc_00E0DDD4: mov edx, var_D4
  loc_00E0DDDA: mov var_4C, eax
  loc_00E0DDDD: mov eax, 00000008h
  loc_00E0DDE2: mov var_18, 00000000h
  loc_00E0DDE9: mov var_54, eax
  loc_00E0DDEC: mov ecx, [edx]
  loc_00E0DDEE: sub esp, 00000010h
  loc_00E0DDF1: mov edx, esp
  loc_00E0DDF3: mov [edx], eax
  loc_00E0DDF5: mov eax, var_50
  loc_00E0DDF8: mov [edx+00000004h], eax
  loc_00E0DDFB: mov eax, var_4C
  loc_00E0DDFE: mov [edx+00000008h], eax
  loc_00E0DE01: mov eax, var_48
  loc_00E0DE04: mov [edx+0000000Ch], eax
  loc_00E0DE07: mov edx, var_D4
  loc_00E0DE0D: push edx
  loc_00E0DE0E: call [ecx+00000038h]
  loc_00E0DE11: test eax, eax
  loc_00E0DE13: fnclex
  loc_00E0DE15: jge 00E0DE28h
  loc_00E0DE17: mov ecx, var_D4
  loc_00E0DE1D: push 00000038h
  loc_00E0DE1F: push 006E09F8h
  loc_00E0DE24: push ecx
  loc_00E0DE25: push eax
  loc_00E0DE26: call ebx
  loc_00E0DE28: lea edx, var_34
  loc_00E0DE2B: lea eax, var_30
  loc_00E0DE2E: push edx
  loc_00E0DE2F: lea ecx, var_2C
  loc_00E0DE32: push eax
  loc_00E0DE33: lea edx, var_28
  loc_00E0DE36: push ecx
  loc_00E0DE37: lea eax, var_24
  loc_00E0DE3A: push edx
  loc_00E0DE3B: push eax
  loc_00E0DE3C: push 00000005h
  loc_00E0DE3E: call [00401048h] ; __vbaFreeObjList
  loc_00E0DE44: lea ecx, var_54
  loc_00E0DE47: lea edx, var_44
  loc_00E0DE4A: push ecx
  loc_00E0DE4B: push edx
  loc_00E0DE4C: push 00000002h
  loc_00E0DE4E: call [00401038h] ; __vbaFreeVarList
  loc_00E0DE54: mov eax, [esi]
  loc_00E0DE56: add esp, 00000024h
  loc_00E0DE59: push 006DCBD8h
  loc_00E0DE5E: push 00000000h
  loc_00E0DE60: push 00000018h
  loc_00E0DE62: push esi
  loc_00E0DE63: call [eax+00000490h]
  loc_00E0DE69: lea ecx, var_28
  loc_00E0DE6C: push eax
  loc_00E0DE6D: push ecx
  loc_00E0DE6E: call edi
  loc_00E0DE70: lea edx, var_44
  loc_00E0DE73: push eax
  loc_00E0DE74: push edx
  loc_00E0DE75: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0DE7B: add esp, 00000010h
  loc_00E0DE7E: push eax
  loc_00E0DE7F: call [00401120h] ; __vbaCastObjVar
  loc_00E0DE85: push eax
  loc_00E0DE86: lea eax, var_2C
  loc_00E0DE89: push eax
  loc_00E0DE8A: call edi
  loc_00E0DE8C: mov ecx, [eax]
  loc_00E0DE8E: lea edx, var_30
  loc_00E0DE91: push edx
  loc_00E0DE92: push eax
  loc_00E0DE93: mov var_C4, eax
  loc_00E0DE99: call [ecx+00000054h]
  loc_00E0DE9C: test eax, eax
  loc_00E0DE9E: fnclex
  loc_00E0DEA0: jge 00E0DEB3h
  loc_00E0DEA2: mov ecx, var_C4
  loc_00E0DEA8: push 00000054h
  loc_00E0DEAA: push 006DCBE8h
  loc_00E0DEAF: push ecx
  loc_00E0DEB0: push eax
  loc_00E0DEB1: call ebx
  loc_00E0DEB3: mov ecx, var_30
  loc_00E0DEB6: mov eax, 00000002h
  loc_00E0DEBB: mov var_7C, 0000000Eh
  loc_00E0DEC2: mov var_84, eax
  loc_00E0DEC8: mov edx, [ecx]
  loc_00E0DECA: mov var_CC, ecx
  loc_00E0DED0: lea ecx, var_34
  loc_00E0DED3: push ecx
  loc_00E0DED4: sub esp, 00000010h
  loc_00E0DED7: mov ecx, esp
  loc_00E0DED9: mov [ecx], eax
  loc_00E0DEDB: mov eax, var_80
  loc_00E0DEDE: mov [ecx+00000004h], eax
  loc_00E0DEE1: mov eax, var_7C
  loc_00E0DEE4: mov [ecx+00000008h], eax
  loc_00E0DEE7: mov eax, var_78
  loc_00E0DEEA: mov [ecx+0000000Ch], eax
  loc_00E0DEED: mov ecx, var_30
  loc_00E0DEF0: push ecx
  loc_00E0DEF1: call [edx+00000028h]
  loc_00E0DEF4: test eax, eax
  loc_00E0DEF6: fnclex
  loc_00E0DEF8: jge 00E0DF0Bh
  loc_00E0DEFA: mov edx, var_CC
  loc_00E0DF00: push 00000028h
  loc_00E0DF02: push 006E09E8h
  loc_00E0DF07: push edx
  loc_00E0DF08: push eax
  loc_00E0DF09: call ebx
  loc_00E0DF0B: mov eax, var_34
  loc_00E0DF0E: mov ecx, [esi]
  loc_00E0DF10: push esi
  loc_00E0DF11: mov var_D4, eax
  loc_00E0DF17: call [ecx+00000360h]
  loc_00E0DF1D: lea edx, var_24
  loc_00E0DF20: push eax
  loc_00E0DF21: push edx
  loc_00E0DF22: call edi
  loc_00E0DF24: mov ecx, [eax]
  loc_00E0DF26: lea edx, var_18
  loc_00E0DF29: push edx
  loc_00E0DF2A: push eax
  loc_00E0DF2B: mov var_BC, eax
  loc_00E0DF31: call [ecx+000000A0h]
  loc_00E0DF37: test eax, eax
  loc_00E0DF39: fnclex
  loc_00E0DF3B: jge 00E0DF51h
  loc_00E0DF3D: mov ecx, var_BC
  loc_00E0DF43: push 000000A0h
  loc_00E0DF48: push 006DCB70h
  loc_00E0DF4D: push ecx
  loc_00E0DF4E: push eax
  loc_00E0DF4F: call ebx
  loc_00E0DF51: mov eax, var_18
  loc_00E0DF54: mov edx, var_D4
  loc_00E0DF5A: mov var_4C, eax
  loc_00E0DF5D: mov eax, 00000008h
  loc_00E0DF62: mov var_18, 00000000h
  loc_00E0DF69: mov var_54, eax
  loc_00E0DF6C: mov ecx, [edx]
  loc_00E0DF6E: sub esp, 00000010h
  loc_00E0DF71: mov edx, esp
  loc_00E0DF73: mov [edx], eax
  loc_00E0DF75: mov eax, var_50
  loc_00E0DF78: mov [edx+00000004h], eax
  loc_00E0DF7B: mov eax, var_4C
  loc_00E0DF7E: mov [edx+00000008h], eax
  loc_00E0DF81: mov eax, var_48
  loc_00E0DF84: mov [edx+0000000Ch], eax
  loc_00E0DF87: mov edx, var_D4
  loc_00E0DF8D: push edx
  loc_00E0DF8E: call [ecx+00000038h]
  loc_00E0DF91: test eax, eax
  loc_00E0DF93: fnclex
  loc_00E0DF95: jge 00E0DFA8h
  loc_00E0DF97: mov ecx, var_D4
  loc_00E0DF9D: push 00000038h
  loc_00E0DF9F: push 006E09F8h
  loc_00E0DFA4: push ecx
  loc_00E0DFA5: push eax
  loc_00E0DFA6: call ebx
  loc_00E0DFA8: lea edx, var_34
  loc_00E0DFAB: lea eax, var_30
  loc_00E0DFAE: push edx
  loc_00E0DFAF: lea ecx, var_2C
  loc_00E0DFB2: push eax
  loc_00E0DFB3: lea edx, var_28
  loc_00E0DFB6: push ecx
  loc_00E0DFB7: lea eax, var_24
  loc_00E0DFBA: push edx
  loc_00E0DFBB: push eax
  loc_00E0DFBC: push 00000005h
  loc_00E0DFBE: call [00401048h] ; __vbaFreeObjList
  loc_00E0DFC4: lea ecx, var_54
  loc_00E0DFC7: lea edx, var_44
  loc_00E0DFCA: push ecx
  loc_00E0DFCB: push edx
  loc_00E0DFCC: push 00000002h
  loc_00E0DFCE: call [00401038h] ; __vbaFreeVarList
  loc_00E0DFD4: mov eax, [esi]
  loc_00E0DFD6: add esp, 00000024h
  loc_00E0DFD9: push 006DCBD8h
  loc_00E0DFDE: push 00000000h
  loc_00E0DFE0: push 00000018h
  loc_00E0DFE2: push esi
  loc_00E0DFE3: call [eax+00000490h]
  loc_00E0DFE9: lea ecx, var_28
  loc_00E0DFEC: push eax
  loc_00E0DFED: push ecx
  loc_00E0DFEE: call edi
  loc_00E0DFF0: lea edx, var_44
  loc_00E0DFF3: push eax
  loc_00E0DFF4: push edx
  loc_00E0DFF5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0DFFB: add esp, 00000010h
  loc_00E0DFFE: push eax
  loc_00E0DFFF: call [00401120h] ; __vbaCastObjVar
  loc_00E0E005: push eax
  loc_00E0E006: lea eax, var_2C
  loc_00E0E009: push eax
  loc_00E0E00A: call edi
  loc_00E0E00C: mov ecx, [eax]
  loc_00E0E00E: lea edx, var_30
  loc_00E0E011: push edx
  loc_00E0E012: push eax
  loc_00E0E013: mov var_C4, eax
  loc_00E0E019: call [ecx+00000054h]
  loc_00E0E01C: test eax, eax
  loc_00E0E01E: fnclex
  loc_00E0E020: jge 00E0E033h
  loc_00E0E022: mov ecx, var_C4
  loc_00E0E028: push 00000054h
  loc_00E0E02A: push 006DCBE8h
  loc_00E0E02F: push ecx
  loc_00E0E030: push eax
  loc_00E0E031: call ebx
  loc_00E0E033: mov ecx, var_30
  loc_00E0E036: mov eax, 00000002h
  loc_00E0E03B: mov var_7C, 00000014h
  loc_00E0E042: mov var_84, eax
  loc_00E0E048: mov edx, [ecx]
  loc_00E0E04A: mov var_CC, ecx
  loc_00E0E050: lea ecx, var_34
  loc_00E0E053: push ecx
  loc_00E0E054: sub esp, 00000010h
  loc_00E0E057: mov ecx, esp
  loc_00E0E059: mov [ecx], eax
  loc_00E0E05B: mov eax, var_80
  loc_00E0E05E: mov [ecx+00000004h], eax
  loc_00E0E061: mov eax, var_7C
  loc_00E0E064: mov [ecx+00000008h], eax
  loc_00E0E067: mov eax, var_78
  loc_00E0E06A: mov [ecx+0000000Ch], eax
  loc_00E0E06D: mov ecx, var_30
  loc_00E0E070: push ecx
  loc_00E0E071: call [edx+00000028h]
  loc_00E0E074: test eax, eax
  loc_00E0E076: fnclex
  loc_00E0E078: jge 00E0E08Bh
  loc_00E0E07A: mov edx, var_CC
  loc_00E0E080: push 00000028h
  loc_00E0E082: push 006E09E8h
  loc_00E0E087: push edx
  loc_00E0E088: push eax
  loc_00E0E089: call ebx
  loc_00E0E08B: mov eax, var_34
  loc_00E0E08E: mov ecx, [esi]
  loc_00E0E090: push esi
  loc_00E0E091: mov var_D4, eax
  loc_00E0E097: call [ecx+00000304h]
  loc_00E0E09D: lea edx, var_24
  loc_00E0E0A0: push eax
  loc_00E0E0A1: push edx
  loc_00E0E0A2: call edi
  loc_00E0E0A4: mov ecx, [eax]
  loc_00E0E0A6: lea edx, var_18
  loc_00E0E0A9: push edx
  loc_00E0E0AA: push eax
  loc_00E0E0AB: mov var_BC, eax
  loc_00E0E0B1: call [ecx+000000A0h]
  loc_00E0E0B7: test eax, eax
  loc_00E0E0B9: fnclex
  loc_00E0E0BB: jge 00E0E0D1h
  loc_00E0E0BD: mov ecx, var_BC
  loc_00E0E0C3: push 000000A0h
  loc_00E0E0C8: push 006DCB70h
  loc_00E0E0CD: push ecx
  loc_00E0E0CE: push eax
  loc_00E0E0CF: call ebx
  loc_00E0E0D1: mov eax, var_18
  loc_00E0E0D4: mov edx, var_D4
  loc_00E0E0DA: mov var_4C, eax
  loc_00E0E0DD: mov eax, 00000008h
  loc_00E0E0E2: mov var_18, 00000000h
  loc_00E0E0E9: mov var_54, eax
  loc_00E0E0EC: mov ecx, [edx]
  loc_00E0E0EE: sub esp, 00000010h
  loc_00E0E0F1: mov edx, esp
  loc_00E0E0F3: mov [edx], eax
  loc_00E0E0F5: mov eax, var_50
  loc_00E0E0F8: mov [edx+00000004h], eax
  loc_00E0E0FB: mov eax, var_4C
  loc_00E0E0FE: mov [edx+00000008h], eax
  loc_00E0E101: mov eax, var_48
  loc_00E0E104: mov [edx+0000000Ch], eax
  loc_00E0E107: mov edx, var_D4
  loc_00E0E10D: push edx
  loc_00E0E10E: call [ecx+00000038h]
  loc_00E0E111: test eax, eax
  loc_00E0E113: fnclex
  loc_00E0E115: jge 00E0E128h
  loc_00E0E117: mov ecx, var_D4
  loc_00E0E11D: push 00000038h
  loc_00E0E11F: push 006E09F8h
  loc_00E0E124: push ecx
  loc_00E0E125: push eax
  loc_00E0E126: call ebx
  loc_00E0E128: lea edx, var_34
  loc_00E0E12B: lea eax, var_30
  loc_00E0E12E: push edx
  loc_00E0E12F: lea ecx, var_2C
  loc_00E0E132: push eax
  loc_00E0E133: lea edx, var_28
  loc_00E0E136: push ecx
  loc_00E0E137: lea eax, var_24
  loc_00E0E13A: push edx
  loc_00E0E13B: push eax
  loc_00E0E13C: push 00000005h
  loc_00E0E13E: call [00401048h] ; __vbaFreeObjList
  loc_00E0E144: lea ecx, var_54
  loc_00E0E147: lea edx, var_44
  loc_00E0E14A: push ecx
  loc_00E0E14B: push edx
  loc_00E0E14C: push 00000002h
  loc_00E0E14E: call [00401038h] ; __vbaFreeVarList
  loc_00E0E154: mov eax, [esi]
  loc_00E0E156: add esp, 00000024h
  loc_00E0E159: push 006DCBD8h
  loc_00E0E15E: push 00000000h
  loc_00E0E160: push 00000018h
  loc_00E0E162: push esi
  loc_00E0E163: call [eax+00000490h]
  loc_00E0E169: lea ecx, var_28
  loc_00E0E16C: push eax
  loc_00E0E16D: push ecx
  loc_00E0E16E: call edi
  loc_00E0E170: lea edx, var_44
  loc_00E0E173: push eax
  loc_00E0E174: push edx
  loc_00E0E175: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0E17B: add esp, 00000010h
  loc_00E0E17E: push eax
  loc_00E0E17F: call [00401120h] ; __vbaCastObjVar
  loc_00E0E185: push eax
  loc_00E0E186: lea eax, var_2C
  loc_00E0E189: push eax
  loc_00E0E18A: call edi
  loc_00E0E18C: mov ecx, [eax]
  loc_00E0E18E: lea edx, var_30
  loc_00E0E191: push edx
  loc_00E0E192: push eax
  loc_00E0E193: mov var_C4, eax
  loc_00E0E199: call [ecx+00000054h]
  loc_00E0E19C: test eax, eax
  loc_00E0E19E: fnclex
  loc_00E0E1A0: jge 00E0E1B3h
  loc_00E0E1A2: mov ecx, var_C4
  loc_00E0E1A8: push 00000054h
  loc_00E0E1AA: push 006DCBE8h
  loc_00E0E1AF: push ecx
  loc_00E0E1B0: push eax
  loc_00E0E1B1: call ebx
  loc_00E0E1B3: mov ecx, var_30
  loc_00E0E1B6: mov eax, 00000002h
  loc_00E0E1BB: mov var_7C, 00000013h
  loc_00E0E1C2: mov var_84, eax
  loc_00E0E1C8: mov edx, [ecx]
  loc_00E0E1CA: mov var_CC, ecx
  loc_00E0E1D0: lea ecx, var_34
  loc_00E0E1D3: push ecx
  loc_00E0E1D4: sub esp, 00000010h
  loc_00E0E1D7: mov ecx, esp
  loc_00E0E1D9: mov [ecx], eax
  loc_00E0E1DB: mov eax, var_80
  loc_00E0E1DE: mov [ecx+00000004h], eax
  loc_00E0E1E1: mov eax, var_7C
  loc_00E0E1E4: mov [ecx+00000008h], eax
  loc_00E0E1E7: mov eax, var_78
  loc_00E0E1EA: mov [ecx+0000000Ch], eax
  loc_00E0E1ED: mov ecx, var_30
  loc_00E0E1F0: push ecx
  loc_00E0E1F1: call [edx+00000028h]
  loc_00E0E1F4: test eax, eax
  loc_00E0E1F6: fnclex
  loc_00E0E1F8: jge 00E0E20Bh
  loc_00E0E1FA: mov edx, var_CC
  loc_00E0E200: push 00000028h
  loc_00E0E202: push 006E09E8h
  loc_00E0E207: push edx
  loc_00E0E208: push eax
  loc_00E0E209: call ebx
  loc_00E0E20B: mov eax, var_34
  loc_00E0E20E: mov ecx, [esi]
  loc_00E0E210: push esi
  loc_00E0E211: mov var_D4, eax
  loc_00E0E217: call [ecx+0000030Ch]
  loc_00E0E21D: lea edx, var_24
  loc_00E0E220: push eax
  loc_00E0E221: push edx
  loc_00E0E222: call edi
  loc_00E0E224: mov ecx, [eax]
  loc_00E0E226: lea edx, var_18
  loc_00E0E229: push edx
  loc_00E0E22A: push eax
  loc_00E0E22B: mov var_BC, eax
  loc_00E0E231: call [ecx+000000A8h]
  loc_00E0E237: test eax, eax
  loc_00E0E239: fnclex
  loc_00E0E23B: jge 00E0E251h
  loc_00E0E23D: mov ecx, var_BC
  loc_00E0E243: push 000000A8h
  loc_00E0E248: push 006E0400h
  loc_00E0E24D: push ecx
  loc_00E0E24E: push eax
  loc_00E0E24F: call ebx
  loc_00E0E251: mov eax, var_18
  loc_00E0E254: mov edx, var_D4
  loc_00E0E25A: mov var_4C, eax
  loc_00E0E25D: mov eax, 00000008h
  loc_00E0E262: mov var_18, 00000000h
  loc_00E0E269: mov var_54, eax
  loc_00E0E26C: mov ecx, [edx]
  loc_00E0E26E: sub esp, 00000010h
  loc_00E0E271: mov edx, esp
  loc_00E0E273: mov [edx], eax
  loc_00E0E275: mov eax, var_50
  loc_00E0E278: mov [edx+00000004h], eax
  loc_00E0E27B: mov eax, var_4C
  loc_00E0E27E: mov [edx+00000008h], eax
  loc_00E0E281: mov eax, var_48
  loc_00E0E284: mov [edx+0000000Ch], eax
  loc_00E0E287: mov edx, var_D4
  loc_00E0E28D: push edx
  loc_00E0E28E: call [ecx+00000038h]
  loc_00E0E291: test eax, eax
  loc_00E0E293: fnclex
  loc_00E0E295: jge 00E0E2A8h
  loc_00E0E297: mov ecx, var_D4
  loc_00E0E29D: push 00000038h
  loc_00E0E29F: push 006E09F8h
  loc_00E0E2A4: push ecx
  loc_00E0E2A5: push eax
  loc_00E0E2A6: call ebx
  loc_00E0E2A8: lea edx, var_34
  loc_00E0E2AB: lea eax, var_30
  loc_00E0E2AE: push edx
  loc_00E0E2AF: lea ecx, var_2C
  loc_00E0E2B2: push eax
  loc_00E0E2B3: lea edx, var_28
  loc_00E0E2B6: push ecx
  loc_00E0E2B7: lea eax, var_24
  loc_00E0E2BA: push edx
  loc_00E0E2BB: push eax
  loc_00E0E2BC: push 00000005h
  loc_00E0E2BE: call [00401048h] ; __vbaFreeObjList
  loc_00E0E2C4: lea ecx, var_54
  loc_00E0E2C7: lea edx, var_44
  loc_00E0E2CA: push ecx
  loc_00E0E2CB: push edx
  loc_00E0E2CC: push 00000002h
  loc_00E0E2CE: call [00401038h] ; __vbaFreeVarList
  loc_00E0E2D4: mov eax, [esi]
  loc_00E0E2D6: add esp, 00000024h
  loc_00E0E2D9: push esi
  loc_00E0E2DA: call [eax+0000032Ch]
  loc_00E0E2E0: lea ecx, var_24
  loc_00E0E2E3: push eax
  loc_00E0E2E4: push ecx
  loc_00E0E2E5: call edi
  loc_00E0E2E7: mov edx, [eax]
  loc_00E0E2E9: push 008080FFh
  loc_00E0E2EE: push eax
  loc_00E0E2EF: mov var_BC, eax
  loc_00E0E2F5: call [edx+0000006Ch]
  loc_00E0E2F8: test eax, eax
  loc_00E0E2FA: fnclex
  loc_00E0E2FC: jge 00E0E30Fh
  loc_00E0E2FE: mov ecx, var_BC
  loc_00E0E304: push 0000006Ch
  loc_00E0E306: push 006E0410h
  loc_00E0E30B: push ecx
  loc_00E0E30C: push eax
  loc_00E0E30D: call ebx
  loc_00E0E30F: lea ecx, var_24
  loc_00E0E312: call [00401254h] ; __vbaFreeObj
  loc_00E0E318: mov edx, [esi]
  loc_00E0E31A: push esi
  loc_00E0E31B: call [edx+0000032Ch]
  loc_00E0E321: push eax
  loc_00E0E322: lea eax, var_24
  loc_00E0E325: push eax
  loc_00E0E326: call edi
  loc_00E0E328: mov ecx, [eax]
  loc_00E0E32A: push 006E0B58h ; "Sekarang Silahkan Input Data Perincian Biaya Di Menu"
  loc_00E0E32F: push eax
  loc_00E0E330: mov var_BC, eax
  loc_00E0E336: call [ecx+00000054h]
  loc_00E0E339: test eax, eax
  loc_00E0E33B: fnclex
  loc_00E0E33D: jge 00E0E350h
  loc_00E0E33F: mov edx, var_BC
  loc_00E0E345: push 00000054h
  loc_00E0E347: push 006E0410h
  loc_00E0E34C: push edx
  loc_00E0E34D: push eax
  loc_00E0E34E: call ebx
  loc_00E0E350: lea ecx, var_24
  loc_00E0E353: call [00401254h] ; __vbaFreeObj
  loc_00E0E359: mov eax, [esi]
  loc_00E0E35B: push esi
  loc_00E0E35C: call [eax+000003C8h]
  loc_00E0E362: lea ecx, var_24
  loc_00E0E365: push eax
  loc_00E0E366: push ecx
  loc_00E0E367: call edi
  loc_00E0E369: mov edx, [eax]
  loc_00E0E36B: push FFFFFFFFh
  loc_00E0E36D: push eax
  loc_00E0E36E: mov var_BC, eax
  loc_00E0E374: call [edx+0000008Ch]
  loc_00E0E37A: test eax, eax
  loc_00E0E37C: fnclex
  loc_00E0E37E: jge 00E0E394h
  loc_00E0E380: mov ecx, var_BC
  loc_00E0E386: push 0000008Ch
  loc_00E0E38B: push 006DCDA0h
  loc_00E0E390: push ecx
  loc_00E0E391: push eax
  loc_00E0E392: call ebx
  loc_00E0E394: lea ecx, var_24
  loc_00E0E397: call [00401254h] ; __vbaFreeObj
  loc_00E0E39D: mov edx, [esi]
  loc_00E0E39F: push esi
  loc_00E0E3A0: call [edx+00000330h]
  loc_00E0E3A6: push eax
  loc_00E0E3A7: lea eax, var_24
  loc_00E0E3AA: push eax
  loc_00E0E3AB: call edi
  loc_00E0E3AD: mov ecx, [eax]
  loc_00E0E3AF: push 00000000h
  loc_00E0E3B1: push eax
  loc_00E0E3B2: mov var_BC, eax
  loc_00E0E3B8: call [ecx+0000005Ch]
  loc_00E0E3BB: test eax, eax
  loc_00E0E3BD: fnclex
  loc_00E0E3BF: jge 00E0E3D2h
  loc_00E0E3C1: mov edx, var_BC
  loc_00E0E3C7: push 0000005Ch
  loc_00E0E3C9: push 006DCAE0h
  loc_00E0E3CE: push edx
  loc_00E0E3CF: push eax
  loc_00E0E3D0: call ebx
  loc_00E0E3D2: lea ecx, var_24
  loc_00E0E3D5: call [00401254h] ; __vbaFreeObj
  loc_00E0E3DB: mov eax, [esi]
  loc_00E0E3DD: push esi
  loc_00E0E3DE: call [eax+0000032Ch]
  loc_00E0E3E4: lea ecx, var_24
  loc_00E0E3E7: push eax
  loc_00E0E3E8: push ecx
  loc_00E0E3E9: call edi
  loc_00E0E3EB: mov edx, [eax]
  loc_00E0E3ED: push 43870000h
  loc_00E0E3F2: push eax
  loc_00E0E3F3: mov var_BC, eax
  loc_00E0E3F9: call [edx+00000074h]
  loc_00E0E3FC: test eax, eax
  loc_00E0E3FE: fnclex
  loc_00E0E400: jge 00E0E413h
  loc_00E0E402: mov ecx, var_BC
  loc_00E0E408: push 00000074h
  loc_00E0E40A: push 006E0410h
  loc_00E0E40F: push ecx
  loc_00E0E410: push eax
  loc_00E0E411: call ebx
  loc_00E0E413: lea ecx, var_24
  loc_00E0E416: call [00401254h] ; __vbaFreeObj
  loc_00E0E41C: mov edx, [esi]
  loc_00E0E41E: push esi
  loc_00E0E41F: call [edx+00000308h]
  loc_00E0E425: push eax
  loc_00E0E426: lea eax, var_24
  loc_00E0E429: push eax
  loc_00E0E42A: call edi
  loc_00E0E42C: mov ecx, [eax]
  loc_00E0E42E: push FFFFFFFFh
  loc_00E0E430: push eax
  loc_00E0E431: mov var_BC, eax
  loc_00E0E437: call [ecx+0000005Ch]
  loc_00E0E43A: test eax, eax
  loc_00E0E43C: fnclex
  loc_00E0E43E: jge 00E0E451h
  loc_00E0E440: mov edx, var_BC
  loc_00E0E446: push 0000005Ch
  loc_00E0E448: push 006DCAE0h
  loc_00E0E44D: push edx
  loc_00E0E44E: push eax
  loc_00E0E44F: call ebx
  loc_00E0E451: lea ecx, var_24
  loc_00E0E454: call [00401254h] ; __vbaFreeObj
  loc_00E0E45A: sub esp, 00000010h
  loc_00E0E45D: mov ecx, 0000000Bh
  loc_00E0E462: mov edx, esp
  loc_00E0E464: mov var_84, ecx
  loc_00E0E46A: xor eax, eax
  loc_00E0E46C: push 8001000Dh
  loc_00E0E471: mov [edx], ecx
  loc_00E0E473: mov ecx, var_80
  loc_00E0E476: mov var_7C, eax
  loc_00E0E479: push esi
  loc_00E0E47A: mov [edx+00000004h], ecx
  loc_00E0E47D: mov ecx, [esi]
  loc_00E0E47F: mov [edx+00000008h], eax
  loc_00E0E482: mov eax, var_78
  loc_00E0E485: mov [edx+0000000Ch], eax
  loc_00E0E488: call [ecx+000003A0h]
  loc_00E0E48E: lea edx, var_24
  loc_00E0E491: push eax
  loc_00E0E492: push edx
  loc_00E0E493: call edi
  loc_00E0E495: push eax
  loc_00E0E496: call [00401238h] ; __vbaLateIdSt
  loc_00E0E49C: lea ecx, var_24
  loc_00E0E49F: call [00401254h] ; __vbaFreeObj
  loc_00E0E4A5: mov eax, [esi]
  loc_00E0E4A7: push esi
  loc_00E0E4A8: call [eax+00000338h]
  loc_00E0E4AE: lea ecx, var_24
  loc_00E0E4B1: push eax
  loc_00E0E4B2: push ecx
  loc_00E0E4B3: call edi
  loc_00E0E4B5: mov edx, [eax]
  loc_00E0E4B7: lea ecx, var_B8
  loc_00E0E4BD: push ecx
  loc_00E0E4BE: push eax
  loc_00E0E4BF: mov var_BC, eax
  loc_00E0E4C5: call [edx+00000098h]
  loc_00E0E4CB: test eax, eax
  loc_00E0E4CD: fnclex
  loc_00E0E4CF: jge 00E0E4E5h
  loc_00E0E4D1: mov edx, var_BC
  loc_00E0E4D7: push 00000098h
  loc_00E0E4DC: push 006DCAD0h
  loc_00E0E4E1: push edx
  loc_00E0E4E2: push eax
  loc_00E0E4E3: call ebx
  loc_00E0E4E5: xor eax, eax
  loc_00E0E4E7: cmp var_B8, FFFFFFh
  loc_00E0E4EF: lea ecx, var_24
  loc_00E0E4F2: setz al
  loc_00E0E4F5: neg eax
  loc_00E0E4F7: mov var_C4, ax
  loc_00E0E4FE: call [00401254h] ; __vbaFreeObj
  loc_00E0E504: cmp var_C4, 0000h
  loc_00E0E50C: jz 00E0EC86h
  loc_00E0E512: mov ecx, [esi]
  loc_00E0E514: push 006DCBD8h
  loc_00E0E519: push 00000000h
  loc_00E0E51B: push 00000018h
  loc_00E0E51D: push esi
  loc_00E0E51E: call [ecx+00000490h]
  loc_00E0E524: lea edx, var_28
  loc_00E0E527: push eax
  loc_00E0E528: push edx
  loc_00E0E529: call edi
  loc_00E0E52B: push eax
  loc_00E0E52C: lea eax, var_44
  loc_00E0E52F: push eax
  loc_00E0E530: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0E536: add esp, 00000010h
  loc_00E0E539: push eax
  loc_00E0E53A: call [00401120h] ; __vbaCastObjVar
  loc_00E0E540: lea ecx, var_2C
  loc_00E0E543: push eax
  loc_00E0E544: push ecx
  loc_00E0E545: call edi
  loc_00E0E547: mov edx, [eax]
  loc_00E0E549: lea ecx, var_30
  loc_00E0E54C: push ecx
  loc_00E0E54D: push eax
  loc_00E0E54E: mov var_C4, eax
  loc_00E0E554: call [edx+00000054h]
  loc_00E0E557: test eax, eax
  loc_00E0E559: fnclex
  loc_00E0E55B: jge 00E0E56Eh
  loc_00E0E55D: mov edx, var_C4
  loc_00E0E563: push 00000054h
  loc_00E0E565: push 006DCBE8h
  loc_00E0E56A: push edx
  loc_00E0E56B: push eax
  loc_00E0E56C: call ebx
  loc_00E0E56E: lea edx, var_34
  loc_00E0E571: mov eax, 00000002h
  loc_00E0E576: push edx
  loc_00E0E577: mov ecx, var_30
  loc_00E0E57A: sub esp, 00000010h
  loc_00E0E57D: mov var_84, eax
  loc_00E0E583: mov edx, esp
  loc_00E0E585: mov var_7C, 0000000Fh
  loc_00E0E58C: mov var_CC, ecx
  loc_00E0E592: mov ecx, [ecx]
  loc_00E0E594: mov [edx], eax
  loc_00E0E596: mov eax, var_80
  loc_00E0E599: mov [edx+00000004h], eax
  loc_00E0E59C: mov eax, var_7C
  loc_00E0E59F: mov [edx+00000008h], eax
  loc_00E0E5A2: mov eax, var_78
  loc_00E0E5A5: mov [edx+0000000Ch], eax
  loc_00E0E5A8: mov edx, var_30
  loc_00E0E5AB: push edx
  loc_00E0E5AC: call [ecx+00000028h]
  loc_00E0E5AF: test eax, eax
  loc_00E0E5B1: fnclex
  loc_00E0E5B3: jge 00E0E5C6h
  loc_00E0E5B5: mov ecx, var_CC
  loc_00E0E5BB: push 00000028h
  loc_00E0E5BD: push 006E09E8h
  loc_00E0E5C2: push ecx
  loc_00E0E5C3: push eax
  loc_00E0E5C4: call ebx
  loc_00E0E5C6: mov edx, var_34
  loc_00E0E5C9: mov eax, [esi]
  loc_00E0E5CB: push esi
  loc_00E0E5CC: mov var_D4, edx
  loc_00E0E5D2: call [eax+0000033Ch]
  loc_00E0E5D8: lea ecx, var_24
  loc_00E0E5DB: push eax
  loc_00E0E5DC: push ecx
  loc_00E0E5DD: call edi
  loc_00E0E5DF: mov edx, [eax]
  loc_00E0E5E1: lea ecx, var_18
  loc_00E0E5E4: push ecx
  loc_00E0E5E5: push eax
  loc_00E0E5E6: mov var_BC, eax
  loc_00E0E5EC: call [edx+000000A0h]
  loc_00E0E5F2: test eax, eax
  loc_00E0E5F4: fnclex
  loc_00E0E5F6: jge 00E0E60Ch
  loc_00E0E5F8: mov edx, var_BC
  loc_00E0E5FE: push 000000A0h
  loc_00E0E603: push 006DCB70h
  loc_00E0E608: push edx
  loc_00E0E609: push eax
  loc_00E0E60A: call ebx
  loc_00E0E60C: mov eax, var_18
  loc_00E0E60F: mov ecx, var_D4
  loc_00E0E615: mov var_4C, eax
  loc_00E0E618: mov eax, 00000008h
  loc_00E0E61D: mov var_18, 00000000h
  loc_00E0E624: mov var_54, eax
  loc_00E0E627: mov edx, [ecx]
  loc_00E0E629: sub esp, 00000010h
  loc_00E0E62C: mov ecx, esp
  loc_00E0E62E: mov [ecx], eax
  loc_00E0E630: mov eax, var_50
  loc_00E0E633: mov [ecx+00000004h], eax
  loc_00E0E636: mov eax, var_4C
  loc_00E0E639: mov [ecx+00000008h], eax
  loc_00E0E63C: mov eax, var_48
  loc_00E0E63F: mov [ecx+0000000Ch], eax
  loc_00E0E642: mov ecx, var_D4
  loc_00E0E648: push ecx
  loc_00E0E649: call [edx+00000038h]
  loc_00E0E64C: test eax, eax
  loc_00E0E64E: fnclex
  loc_00E0E650: jge 00E0E663h
  loc_00E0E652: mov edx, var_D4
  loc_00E0E658: push 00000038h
  loc_00E0E65A: push 006E09F8h
  loc_00E0E65F: push edx
  loc_00E0E660: push eax
  loc_00E0E661: call ebx
  loc_00E0E663: lea eax, var_34
  loc_00E0E666: lea ecx, var_30
  loc_00E0E669: push eax
  loc_00E0E66A: lea edx, var_2C
  loc_00E0E66D: push ecx
  loc_00E0E66E: lea eax, var_28
  loc_00E0E671: push edx
  loc_00E0E672: lea ecx, var_24
  loc_00E0E675: push eax
  loc_00E0E676: push ecx
  loc_00E0E677: push 00000005h
  loc_00E0E679: call [00401048h] ; __vbaFreeObjList
  loc_00E0E67F: lea edx, var_54
  loc_00E0E682: lea eax, var_44
  loc_00E0E685: push edx
  loc_00E0E686: push eax
  loc_00E0E687: push 00000002h
  loc_00E0E689: call [00401038h] ; __vbaFreeVarList
  loc_00E0E68F: mov ecx, [esi]
  loc_00E0E691: add esp, 00000024h
  loc_00E0E694: push 006DCBD8h
  loc_00E0E699: push 00000000h
  loc_00E0E69B: push 00000018h
  loc_00E0E69D: push esi
  loc_00E0E69E: call [ecx+00000490h]
  loc_00E0E6A4: lea edx, var_28
  loc_00E0E6A7: push eax
  loc_00E0E6A8: push edx
  loc_00E0E6A9: call edi
  loc_00E0E6AB: push eax
  loc_00E0E6AC: lea eax, var_44
  loc_00E0E6AF: push eax
  loc_00E0E6B0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0E6B6: add esp, 00000010h
  loc_00E0E6B9: push eax
  loc_00E0E6BA: call [00401120h] ; __vbaCastObjVar
  loc_00E0E6C0: lea ecx, var_2C
  loc_00E0E6C3: push eax
  loc_00E0E6C4: push ecx
  loc_00E0E6C5: call edi
  loc_00E0E6C7: mov edx, [eax]
  loc_00E0E6C9: lea ecx, var_30
  loc_00E0E6CC: push ecx
  loc_00E0E6CD: push eax
  loc_00E0E6CE: mov var_C4, eax
  loc_00E0E6D4: call [edx+00000054h]
  loc_00E0E6D7: test eax, eax
  loc_00E0E6D9: fnclex
  loc_00E0E6DB: jge 00E0E6EEh
  loc_00E0E6DD: mov edx, var_C4
  loc_00E0E6E3: push 00000054h
  loc_00E0E6E5: push 006DCBE8h
  loc_00E0E6EA: push edx
  loc_00E0E6EB: push eax
  loc_00E0E6EC: call ebx
  loc_00E0E6EE: lea edx, var_34
  loc_00E0E6F1: mov eax, 00000002h
  loc_00E0E6F6: push edx
  loc_00E0E6F7: mov ecx, var_30
  loc_00E0E6FA: sub esp, 00000010h
  loc_00E0E6FD: mov var_84, eax
  loc_00E0E703: mov edx, esp
  loc_00E0E705: mov var_7C, 00000010h
  loc_00E0E70C: mov var_CC, ecx
  loc_00E0E712: mov ecx, [ecx]
  loc_00E0E714: mov [edx], eax
  loc_00E0E716: mov eax, var_80
  loc_00E0E719: mov [edx+00000004h], eax
  loc_00E0E71C: mov eax, var_7C
  loc_00E0E71F: mov [edx+00000008h], eax
  loc_00E0E722: mov eax, var_78
  loc_00E0E725: mov [edx+0000000Ch], eax
  loc_00E0E728: mov edx, var_30
  loc_00E0E72B: push edx
  loc_00E0E72C: call [ecx+00000028h]
  loc_00E0E72F: test eax, eax
  loc_00E0E731: fnclex
  loc_00E0E733: jge 00E0E746h
  loc_00E0E735: mov ecx, var_CC
  loc_00E0E73B: push 00000028h
  loc_00E0E73D: push 006E09E8h
  loc_00E0E742: push ecx
  loc_00E0E743: push eax
  loc_00E0E744: call ebx
  loc_00E0E746: mov edx, var_34
  loc_00E0E749: mov eax, [esi]
  loc_00E0E74B: push esi
  loc_00E0E74C: mov var_D4, edx
  loc_00E0E752: call [eax+00000340h]
  loc_00E0E758: lea ecx, var_24
  loc_00E0E75B: push eax
  loc_00E0E75C: push ecx
  loc_00E0E75D: call edi
  loc_00E0E75F: mov edx, [eax]
  loc_00E0E761: lea ecx, var_18
  loc_00E0E764: push ecx
  loc_00E0E765: push eax
  loc_00E0E766: mov var_BC, eax
  loc_00E0E76C: call [edx+000000A0h]
  loc_00E0E772: test eax, eax
  loc_00E0E774: fnclex
  loc_00E0E776: jge 00E0E78Ch
  loc_00E0E778: mov edx, var_BC
  loc_00E0E77E: push 000000A0h
  loc_00E0E783: push 006DCB70h
  loc_00E0E788: push edx
  loc_00E0E789: push eax
  loc_00E0E78A: call ebx
  loc_00E0E78C: mov eax, var_18
  loc_00E0E78F: mov ecx, var_D4
  loc_00E0E795: mov var_4C, eax
  loc_00E0E798: mov eax, 00000008h
  loc_00E0E79D: mov var_18, 00000000h
  loc_00E0E7A4: mov var_54, eax
  loc_00E0E7A7: mov edx, [ecx]
  loc_00E0E7A9: sub esp, 00000010h
  loc_00E0E7AC: mov ecx, esp
  loc_00E0E7AE: mov [ecx], eax
  loc_00E0E7B0: mov eax, var_50
  loc_00E0E7B3: mov [ecx+00000004h], eax
  loc_00E0E7B6: mov eax, var_4C
  loc_00E0E7B9: mov [ecx+00000008h], eax
  loc_00E0E7BC: mov eax, var_48
  loc_00E0E7BF: mov [ecx+0000000Ch], eax
  loc_00E0E7C2: mov ecx, var_D4
  loc_00E0E7C8: push ecx
  loc_00E0E7C9: call [edx+00000038h]
  loc_00E0E7CC: test eax, eax
  loc_00E0E7CE: fnclex
  loc_00E0E7D0: jge 00E0E7E3h
  loc_00E0E7D2: mov edx, var_D4
  loc_00E0E7D8: push 00000038h
  loc_00E0E7DA: push 006E09F8h
  loc_00E0E7DF: push edx
  loc_00E0E7E0: push eax
  loc_00E0E7E1: call ebx
  loc_00E0E7E3: lea eax, var_34
  loc_00E0E7E6: lea ecx, var_30
  loc_00E0E7E9: push eax
  loc_00E0E7EA: lea edx, var_2C
  loc_00E0E7ED: push ecx
  loc_00E0E7EE: lea eax, var_28
  loc_00E0E7F1: push edx
  loc_00E0E7F2: lea ecx, var_24
  loc_00E0E7F5: push eax
  loc_00E0E7F6: push ecx
  loc_00E0E7F7: push 00000005h
  loc_00E0E7F9: call [00401048h] ; __vbaFreeObjList
  loc_00E0E7FF: lea edx, var_54
  loc_00E0E802: lea eax, var_44
  loc_00E0E805: push edx
  loc_00E0E806: push eax
  loc_00E0E807: push 00000002h
  loc_00E0E809: call [00401038h] ; __vbaFreeVarList
  loc_00E0E80F: mov ecx, [esi]
  loc_00E0E811: add esp, 00000024h
  loc_00E0E814: push 006DCBD8h
  loc_00E0E819: push 00000000h
  loc_00E0E81B: push 00000018h
  loc_00E0E81D: push esi
  loc_00E0E81E: call [ecx+00000490h]
  loc_00E0E824: lea edx, var_28
  loc_00E0E827: push eax
  loc_00E0E828: push edx
  loc_00E0E829: call edi
  loc_00E0E82B: push eax
  loc_00E0E82C: lea eax, var_44
  loc_00E0E82F: push eax
  loc_00E0E830: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0E836: add esp, 00000010h
  loc_00E0E839: push eax
  loc_00E0E83A: call [00401120h] ; __vbaCastObjVar
  loc_00E0E840: lea ecx, var_2C
  loc_00E0E843: push eax
  loc_00E0E844: push ecx
  loc_00E0E845: call edi
  loc_00E0E847: mov edx, [eax]
  loc_00E0E849: lea ecx, var_30
  loc_00E0E84C: push ecx
  loc_00E0E84D: push eax
  loc_00E0E84E: mov var_C4, eax
  loc_00E0E854: call [edx+00000054h]
  loc_00E0E857: test eax, eax
  loc_00E0E859: fnclex
  loc_00E0E85B: jge 00E0E86Eh
  loc_00E0E85D: mov edx, var_C4
  loc_00E0E863: push 00000054h
  loc_00E0E865: push 006DCBE8h
  loc_00E0E86A: push edx
  loc_00E0E86B: push eax
  loc_00E0E86C: call ebx
  loc_00E0E86E: lea edx, var_34
  loc_00E0E871: mov eax, 00000002h
  loc_00E0E876: push edx
  loc_00E0E877: mov ecx, var_30
  loc_00E0E87A: sub esp, 00000010h
  loc_00E0E87D: mov var_84, eax
  loc_00E0E883: mov edx, esp
  loc_00E0E885: mov var_7C, 00000011h
  loc_00E0E88C: mov var_CC, ecx
  loc_00E0E892: mov ecx, [ecx]
  loc_00E0E894: mov [edx], eax
  loc_00E0E896: mov eax, var_80
  loc_00E0E899: mov [edx+00000004h], eax
  loc_00E0E89C: mov eax, var_7C
  loc_00E0E89F: mov [edx+00000008h], eax
  loc_00E0E8A2: mov eax, var_78
  loc_00E0E8A5: mov [edx+0000000Ch], eax
  loc_00E0E8A8: mov edx, var_30
  loc_00E0E8AB: push edx
  loc_00E0E8AC: call [ecx+00000028h]
  loc_00E0E8AF: test eax, eax
  loc_00E0E8B1: fnclex
  loc_00E0E8B3: jge 00E0E8C6h
  loc_00E0E8B5: mov ecx, var_CC
  loc_00E0E8BB: push 00000028h
  loc_00E0E8BD: push 006E09E8h
  loc_00E0E8C2: push ecx
  loc_00E0E8C3: push eax
  loc_00E0E8C4: call ebx
  loc_00E0E8C6: mov edx, var_34
  loc_00E0E8C9: mov eax, [esi]
  loc_00E0E8CB: push esi
  loc_00E0E8CC: mov var_D4, edx
  loc_00E0E8D2: call [eax+00000344h]
  loc_00E0E8D8: lea ecx, var_24
  loc_00E0E8DB: push eax
  loc_00E0E8DC: push ecx
  loc_00E0E8DD: call edi
  loc_00E0E8DF: mov edx, [eax]
  loc_00E0E8E1: lea ecx, var_18
  loc_00E0E8E4: push ecx
  loc_00E0E8E5: push eax
  loc_00E0E8E6: mov var_BC, eax
  loc_00E0E8EC: call [edx+000000A0h]
  loc_00E0E8F2: test eax, eax
  loc_00E0E8F4: fnclex
  loc_00E0E8F6: jge 00E0E90Ch
  loc_00E0E8F8: mov edx, var_BC
  loc_00E0E8FE: push 000000A0h
  loc_00E0E903: push 006DCB70h
  loc_00E0E908: push edx
  loc_00E0E909: push eax
  loc_00E0E90A: call ebx
  loc_00E0E90C: mov eax, var_18
  loc_00E0E90F: mov ecx, var_D4
  loc_00E0E915: mov var_4C, eax
  loc_00E0E918: mov eax, 00000008h
  loc_00E0E91D: mov var_18, 00000000h
  loc_00E0E924: mov var_54, eax
  loc_00E0E927: mov edx, [ecx]
  loc_00E0E929: sub esp, 00000010h
  loc_00E0E92C: mov ecx, esp
  loc_00E0E92E: mov [ecx], eax
  loc_00E0E930: mov eax, var_50
  loc_00E0E933: mov [ecx+00000004h], eax
  loc_00E0E936: mov eax, var_4C
  loc_00E0E939: mov [ecx+00000008h], eax
  loc_00E0E93C: mov eax, var_48
  loc_00E0E93F: mov [ecx+0000000Ch], eax
  loc_00E0E942: mov ecx, var_D4
  loc_00E0E948: push ecx
  loc_00E0E949: call [edx+00000038h]
  loc_00E0E94C: test eax, eax
  loc_00E0E94E: fnclex
  loc_00E0E950: jge 00E0E963h
  loc_00E0E952: mov edx, var_D4
  loc_00E0E958: push 00000038h
  loc_00E0E95A: push 006E09F8h
  loc_00E0E95F: push edx
  loc_00E0E960: push eax
  loc_00E0E961: call ebx
  loc_00E0E963: lea eax, var_34
  loc_00E0E966: lea ecx, var_30
  loc_00E0E969: push eax
  loc_00E0E96A: lea edx, var_2C
  loc_00E0E96D: push ecx
  loc_00E0E96E: lea eax, var_28
  loc_00E0E971: push edx
  loc_00E0E972: lea ecx, var_24
  loc_00E0E975: push eax
  loc_00E0E976: push ecx
  loc_00E0E977: push 00000005h
  loc_00E0E979: call [00401048h] ; __vbaFreeObjList
  loc_00E0E97F: lea edx, var_54
  loc_00E0E982: lea eax, var_44
  loc_00E0E985: push edx
  loc_00E0E986: push eax
  loc_00E0E987: push 00000002h
  loc_00E0E989: call [00401038h] ; __vbaFreeVarList
  loc_00E0E98F: mov ecx, [esi]
  loc_00E0E991: add esp, 00000024h
  loc_00E0E994: push 006DCBD8h
  loc_00E0E999: push 00000000h
  loc_00E0E99B: push 00000018h
  loc_00E0E99D: push esi
  loc_00E0E99E: call [ecx+00000490h]
  loc_00E0E9A4: lea edx, var_28
  loc_00E0E9A7: push eax
  loc_00E0E9A8: push edx
  loc_00E0E9A9: call edi
  loc_00E0E9AB: push eax
  loc_00E0E9AC: lea eax, var_44
  loc_00E0E9AF: push eax
  loc_00E0E9B0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0E9B6: add esp, 00000010h
  loc_00E0E9B9: push eax
  loc_00E0E9BA: call [00401120h] ; __vbaCastObjVar
  loc_00E0E9C0: lea ecx, var_2C
  loc_00E0E9C3: push eax
  loc_00E0E9C4: push ecx
  loc_00E0E9C5: call edi
  loc_00E0E9C7: mov edx, [eax]
  loc_00E0E9C9: lea ecx, var_30
  loc_00E0E9CC: push ecx
  loc_00E0E9CD: push eax
  loc_00E0E9CE: mov var_C4, eax
  loc_00E0E9D4: call [edx+00000054h]
  loc_00E0E9D7: test eax, eax
  loc_00E0E9D9: fnclex
  loc_00E0E9DB: jge 00E0E9EEh
  loc_00E0E9DD: mov edx, var_C4
  loc_00E0E9E3: push 00000054h
  loc_00E0E9E5: push 006DCBE8h
  loc_00E0E9EA: push edx
  loc_00E0E9EB: push eax
  loc_00E0E9EC: call ebx
  loc_00E0E9EE: lea edx, var_34
  loc_00E0E9F1: mov eax, 00000002h
  loc_00E0E9F6: push edx
  loc_00E0E9F7: mov ecx, var_30
  loc_00E0E9FA: sub esp, 00000010h
  loc_00E0E9FD: mov var_84, eax
  loc_00E0EA03: mov edx, esp
  loc_00E0EA05: mov var_7C, 00000012h
  loc_00E0EA0C: mov var_CC, ecx
  loc_00E0EA12: mov ecx, [ecx]
  loc_00E0EA14: mov [edx], eax
  loc_00E0EA16: mov eax, var_80
  loc_00E0EA19: mov [edx+00000004h], eax
  loc_00E0EA1C: mov eax, var_7C
  loc_00E0EA1F: mov [edx+00000008h], eax
  loc_00E0EA22: mov eax, var_78
  loc_00E0EA25: mov [edx+0000000Ch], eax
  loc_00E0EA28: mov edx, var_30
  loc_00E0EA2B: push edx
  loc_00E0EA2C: call [ecx+00000028h]
  loc_00E0EA2F: test eax, eax
  loc_00E0EA31: fnclex
  loc_00E0EA33: jge 00E0EA46h
  loc_00E0EA35: mov ecx, var_CC
  loc_00E0EA3B: push 00000028h
  loc_00E0EA3D: push 006E09E8h
  loc_00E0EA42: push ecx
  loc_00E0EA43: push eax
  loc_00E0EA44: call ebx
  loc_00E0EA46: mov edx, var_34
  loc_00E0EA49: mov eax, [esi]
  loc_00E0EA4B: push esi
  loc_00E0EA4C: mov var_D4, edx
  loc_00E0EA52: call [eax+00000348h]
  loc_00E0EA58: lea ecx, var_24
  loc_00E0EA5B: push eax
  loc_00E0EA5C: push ecx
  loc_00E0EA5D: call edi
  loc_00E0EA5F: mov edx, [eax]
  loc_00E0EA61: lea ecx, var_18
  loc_00E0EA64: push ecx
  loc_00E0EA65: push eax
  loc_00E0EA66: mov var_BC, eax
  loc_00E0EA6C: call [edx+000000A0h]
  loc_00E0EA72: test eax, eax
  loc_00E0EA74: fnclex
  loc_00E0EA76: jge 00E0EA8Ch
  loc_00E0EA78: mov edx, var_BC
  loc_00E0EA7E: push 000000A0h
  loc_00E0EA83: push 006DCB70h
  loc_00E0EA88: push edx
  loc_00E0EA89: push eax
  loc_00E0EA8A: call ebx
  loc_00E0EA8C: mov eax, var_18
  loc_00E0EA8F: mov ecx, var_D4
  loc_00E0EA95: mov var_4C, eax
  loc_00E0EA98: mov eax, 00000008h
  loc_00E0EA9D: mov var_18, 00000000h
  loc_00E0EAA4: mov var_54, eax
  loc_00E0EAA7: mov edx, [ecx]
  loc_00E0EAA9: sub esp, 00000010h
  loc_00E0EAAC: mov ecx, esp
  loc_00E0EAAE: mov [ecx], eax
  loc_00E0EAB0: mov eax, var_50
  loc_00E0EAB3: mov [ecx+00000004h], eax
  loc_00E0EAB6: mov eax, var_4C
  loc_00E0EAB9: mov [ecx+00000008h], eax
  loc_00E0EABC: mov eax, var_48
  loc_00E0EABF: mov [ecx+0000000Ch], eax
  loc_00E0EAC2: mov ecx, var_D4
  loc_00E0EAC8: push ecx
  loc_00E0EAC9: call [edx+00000038h]
  loc_00E0EACC: test eax, eax
  loc_00E0EACE: fnclex
  loc_00E0EAD0: jge 00E0EAE3h
  loc_00E0EAD2: mov edx, var_D4
  loc_00E0EAD8: push 00000038h
  loc_00E0EADA: push 006E09F8h
  loc_00E0EADF: push edx
  loc_00E0EAE0: push eax
  loc_00E0EAE1: call ebx
  loc_00E0EAE3: lea eax, var_34
  loc_00E0EAE6: lea ecx, var_30
  loc_00E0EAE9: push eax
  loc_00E0EAEA: lea edx, var_2C
  loc_00E0EAED: push ecx
  loc_00E0EAEE: lea eax, var_28
  loc_00E0EAF1: push edx
  loc_00E0EAF2: lea ecx, var_24
  loc_00E0EAF5: push eax
  loc_00E0EAF6: push ecx
  loc_00E0EAF7: push 00000005h
  loc_00E0EAF9: call [00401048h] ; __vbaFreeObjList
  loc_00E0EAFF: lea edx, var_54
  loc_00E0EB02: lea eax, var_44
  loc_00E0EB05: push edx
  loc_00E0EB06: push eax
  loc_00E0EB07: push 00000002h
  loc_00E0EB09: call [00401038h] ; __vbaFreeVarList
  loc_00E0EB0F: mov ecx, [esi]
  loc_00E0EB11: add esp, 00000024h
  loc_00E0EB14: push 006DCBD8h
  loc_00E0EB19: push 00000000h
  loc_00E0EB1B: push 00000018h
  loc_00E0EB1D: push esi
  loc_00E0EB1E: call [ecx+00000490h]
  loc_00E0EB24: lea edx, var_24
  loc_00E0EB27: push eax
  loc_00E0EB28: push edx
  loc_00E0EB29: call edi
  loc_00E0EB2B: push eax
  loc_00E0EB2C: lea eax, var_44
  loc_00E0EB2F: push eax
  loc_00E0EB30: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0EB36: add esp, 00000010h
  loc_00E0EB39: push eax
  loc_00E0EB3A: call [00401120h] ; __vbaCastObjVar
  loc_00E0EB40: lea ecx, var_28
  loc_00E0EB43: push eax
  loc_00E0EB44: push ecx
  loc_00E0EB45: call edi
  loc_00E0EB47: mov ecx, 80020004h
  loc_00E0EB4C: sub esp, 00000010h
  loc_00E0EB4F: mov var_8C, ecx
  loc_00E0EB55: mov var_7C, ecx
  loc_00E0EB58: mov ebx, eax
  loc_00E0EB5A: mov ecx, esp
  loc_00E0EB5C: mov eax, 0000000Ah
  loc_00E0EB61: sub esp, 00000010h
  loc_00E0EB64: mov var_94, eax
  loc_00E0EB6A: mov var_84, eax
  loc_00E0EB70: mov [ecx], eax
  loc_00E0EB72: mov eax, var_90
  loc_00E0EB78: mov edx, [ebx]
  loc_00E0EB7A: mov [ecx+00000004h], eax
  loc_00E0EB7D: mov eax, var_8C
  loc_00E0EB83: mov [ecx+00000008h], eax
  loc_00E0EB86: mov eax, var_88
  loc_00E0EB8C: mov [ecx+0000000Ch], eax
  loc_00E0EB8F: mov eax, var_84
  loc_00E0EB95: mov ecx, esp
  loc_00E0EB97: push ebx
  loc_00E0EB98: mov [ecx], eax
  loc_00E0EB9A: mov eax, var_80
  loc_00E0EB9D: mov [ecx+00000004h], eax
  loc_00E0EBA0: mov eax, var_7C
  loc_00E0EBA3: mov [ecx+00000008h], eax
  loc_00E0EBA6: mov eax, var_78
  loc_00E0EBA9: mov [ecx+0000000Ch], eax
  loc_00E0EBAC: call [edx+000000ACh]
  loc_00E0EBB2: test eax, eax
  loc_00E0EBB4: fnclex
  loc_00E0EBB6: jge 00E0EBCAh
  loc_00E0EBB8: push 000000ACh
  loc_00E0EBBD: push 006DCBE8h
  loc_00E0EBC2: push ebx
  loc_00E0EBC3: push eax
  loc_00E0EBC4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0EBCA: lea ecx, var_28
  loc_00E0EBCD: lea edx, var_24
  loc_00E0EBD0: push ecx
  loc_00E0EBD1: push edx
  loc_00E0EBD2: push 00000002h
  loc_00E0EBD4: call [00401048h] ; __vbaFreeObjList
  loc_00E0EBDA: add esp, 0000000Ch
  loc_00E0EBDD: lea ecx, var_44
  loc_00E0EBE0: call [00401024h] ; __vbaFreeVar
  loc_00E0EBE6: mov eax, [esi]
  loc_00E0EBE8: push 00000000h
  loc_00E0EBEA: push 00000019h
  loc_00E0EBEC: push esi
  loc_00E0EBED: call [eax+00000490h]
  loc_00E0EBF3: lea ecx, var_24
  loc_00E0EBF6: push eax
  loc_00E0EBF7: push ecx
  loc_00E0EBF8: call edi
  loc_00E0EBFA: push eax
  loc_00E0EBFB: call [00401030h] ; __vbaLateIdCall
  loc_00E0EC01: add esp, 0000000Ch
  loc_00E0EC04: lea ecx, var_24
  loc_00E0EC07: call [00401254h] ; __vbaFreeObj
  loc_00E0EC0D: mov edx, [esi]
  loc_00E0EC0F: push esi
  loc_00E0EC10: call [edx+00000334h]
  loc_00E0EC16: push eax
  loc_00E0EC17: lea eax, var_24
  loc_00E0EC1A: push eax
  loc_00E0EC1B: call edi
  loc_00E0EC1D: mov ebx, eax
  loc_00E0EC1F: push FFFFFFFFh
  loc_00E0EC21: push ebx
  loc_00E0EC22: mov ecx, [ebx]
  loc_00E0EC24: call [ecx+0000005Ch]
  loc_00E0EC27: test eax, eax
  loc_00E0EC29: fnclex
  loc_00E0EC2B: jge 00E0EC3Ch
  loc_00E0EC2D: push 0000005Ch
  loc_00E0EC2F: push 006DCAE0h
  loc_00E0EC34: push ebx
  loc_00E0EC35: push eax
  loc_00E0EC36: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0EC3C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E0EC42: lea ecx, var_24
  loc_00E0EC45: call ebx
  loc_00E0EC47: mov edx, [esi]
  loc_00E0EC49: push esi
  loc_00E0EC4A: call [edx+000003E8h]
  loc_00E0EC50: push eax
  loc_00E0EC51: lea eax, var_24
  loc_00E0EC54: push eax
  loc_00E0EC55: call edi
  loc_00E0EC57: mov esi, eax
  loc_00E0EC59: push 00000000h
  loc_00E0EC5B: push esi
  loc_00E0EC5C: mov ecx, [esi]
  loc_00E0EC5E: call [ecx+0000008Ch]
  loc_00E0EC64: test eax, eax
  loc_00E0EC66: fnclex
  loc_00E0EC68: jge 00E0EC7Ch
  loc_00E0EC6A: push 0000008Ch
  loc_00E0EC6F: push 006DCDA0h
  loc_00E0EC74: push esi
  loc_00E0EC75: push eax
  loc_00E0EC76: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0EC7C: lea ecx, var_24
  loc_00E0EC7F: call ebx
  loc_00E0EC81: jmp 00E0F3C5h
  loc_00E0EC86: mov edx, [esi]
  loc_00E0EC88: push esi
  loc_00E0EC89: call [edx+00000338h]
  loc_00E0EC8F: push eax
  loc_00E0EC90: lea eax, var_24
  loc_00E0EC93: push eax
  loc_00E0EC94: call edi
  loc_00E0EC96: mov ecx, [eax]
  loc_00E0EC98: lea edx, var_B8
  loc_00E0EC9E: push edx
  loc_00E0EC9F: push eax
  loc_00E0ECA0: mov var_BC, eax
  loc_00E0ECA6: call [ecx+00000098h]
  loc_00E0ECAC: test eax, eax
  loc_00E0ECAE: fnclex
  loc_00E0ECB0: jge 00E0ECC6h
  loc_00E0ECB2: mov ecx, var_BC
  loc_00E0ECB8: push 00000098h
  loc_00E0ECBD: push 006DCAD0h
  loc_00E0ECC2: push ecx
  loc_00E0ECC3: push eax
  loc_00E0ECC4: call ebx
  loc_00E0ECC6: xor edx, edx
  loc_00E0ECC8: lea ecx, var_24
  loc_00E0ECCB: cmp var_B8, dx
  loc_00E0ECD2: setz dl
  loc_00E0ECD5: neg edx
  loc_00E0ECD7: mov var_C4, dx
  loc_00E0ECDE: call [00401254h] ; __vbaFreeObj
  loc_00E0ECE4: cmp var_C4, 0000h
  loc_00E0ECEC: jz 00E0F3C5h
  loc_00E0ECF2: mov eax, [esi]
  loc_00E0ECF4: push 006DCBD8h
  loc_00E0ECF9: push 00000000h
  loc_00E0ECFB: push 00000018h
  loc_00E0ECFD: push esi
  loc_00E0ECFE: call [eax+00000490h]
  loc_00E0ED04: lea ecx, var_24
  loc_00E0ED07: push eax
  loc_00E0ED08: push ecx
  loc_00E0ED09: call edi
  loc_00E0ED0B: lea edx, var_44
  loc_00E0ED0E: push eax
  loc_00E0ED0F: push edx
  loc_00E0ED10: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0ED16: add esp, 00000010h
  loc_00E0ED19: push eax
  loc_00E0ED1A: call [00401120h] ; __vbaCastObjVar
  loc_00E0ED20: push eax
  loc_00E0ED21: lea eax, var_28
  loc_00E0ED24: push eax
  loc_00E0ED25: call edi
  loc_00E0ED27: mov ecx, [eax]
  loc_00E0ED29: lea edx, var_2C
  loc_00E0ED2C: push edx
  loc_00E0ED2D: push eax
  loc_00E0ED2E: mov var_BC, eax
  loc_00E0ED34: call [ecx+00000054h]
  loc_00E0ED37: test eax, eax
  loc_00E0ED39: fnclex
  loc_00E0ED3B: jge 00E0ED4Eh
  loc_00E0ED3D: mov ecx, var_BC
  loc_00E0ED43: push 00000054h
  loc_00E0ED45: push 006DCBE8h
  loc_00E0ED4A: push ecx
  loc_00E0ED4B: push eax
  loc_00E0ED4C: call ebx
  loc_00E0ED4E: mov ecx, var_2C
  loc_00E0ED51: mov eax, 00000002h
  loc_00E0ED56: mov var_7C, 0000000Fh
  loc_00E0ED5D: mov var_84, eax
  loc_00E0ED63: mov edx, [ecx]
  loc_00E0ED65: mov var_C4, ecx
  loc_00E0ED6B: lea ecx, var_30
  loc_00E0ED6E: push ecx
  loc_00E0ED6F: sub esp, 00000010h
  loc_00E0ED72: mov ecx, esp
  loc_00E0ED74: mov [ecx], eax
  loc_00E0ED76: mov eax, var_80
  loc_00E0ED79: mov [ecx+00000004h], eax
  loc_00E0ED7C: mov eax, var_7C
  loc_00E0ED7F: mov [ecx+00000008h], eax
  loc_00E0ED82: mov eax, var_78
  loc_00E0ED85: mov [ecx+0000000Ch], eax
  loc_00E0ED88: mov ecx, var_2C
  loc_00E0ED8B: push ecx
  loc_00E0ED8C: call [edx+00000028h]
  loc_00E0ED8F: test eax, eax
  loc_00E0ED91: fnclex
  loc_00E0ED93: jge 00E0EDA6h
  loc_00E0ED95: mov edx, var_C4
  loc_00E0ED9B: push 00000028h
  loc_00E0ED9D: push 006E09E8h
  loc_00E0EDA2: push edx
  loc_00E0EDA3: push eax
  loc_00E0EDA4: call ebx
  loc_00E0EDA6: sub esp, 00000010h
  loc_00E0EDA9: mov eax, 00000008h
  loc_00E0EDAE: mov edx, esp
  loc_00E0EDB0: mov ecx, var_30
  loc_00E0EDB3: mov var_94, eax
  loc_00E0EDB9: mov var_8C, 006E0BC8h ; " - "
  loc_00E0EDC3: mov [edx], eax
  loc_00E0EDC5: mov eax, var_90
  loc_00E0EDCB: mov var_CC, ecx
  loc_00E0EDD1: mov ecx, [ecx]
  loc_00E0EDD3: mov [edx+00000004h], eax
  loc_00E0EDD6: mov eax, var_8C
  loc_00E0EDDC: mov [edx+00000008h], eax
  loc_00E0EDDF: mov eax, var_88
  loc_00E0EDE5: mov [edx+0000000Ch], eax
  loc_00E0EDE8: mov edx, var_30
  loc_00E0EDEB: push edx
  loc_00E0EDEC: call [ecx+00000038h]
  loc_00E0EDEF: test eax, eax
  loc_00E0EDF1: fnclex
  loc_00E0EDF3: jge 00E0EE06h
  loc_00E0EDF5: mov ecx, var_CC
  loc_00E0EDFB: push 00000038h
  loc_00E0EDFD: push 006E09F8h
  loc_00E0EE02: push ecx
  loc_00E0EE03: push eax
  loc_00E0EE04: call ebx
  loc_00E0EE06: lea edx, var_30
  loc_00E0EE09: lea eax, var_2C
  loc_00E0EE0C: push edx
  loc_00E0EE0D: lea ecx, var_28
  loc_00E0EE10: push eax
  loc_00E0EE11: lea edx, var_24
  loc_00E0EE14: push ecx
  loc_00E0EE15: push edx
  loc_00E0EE16: push 00000004h
  loc_00E0EE18: call [00401048h] ; __vbaFreeObjList
  loc_00E0EE1E: add esp, 00000014h
  loc_00E0EE21: lea ecx, var_44
  loc_00E0EE24: call [00401024h] ; __vbaFreeVar
  loc_00E0EE2A: mov eax, [esi]
  loc_00E0EE2C: push 006DCBD8h
  loc_00E0EE31: push 00000000h
  loc_00E0EE33: push 00000018h
  loc_00E0EE35: push esi
  loc_00E0EE36: call [eax+00000490h]
  loc_00E0EE3C: lea ecx, var_24
  loc_00E0EE3F: push eax
  loc_00E0EE40: push ecx
  loc_00E0EE41: call edi
  loc_00E0EE43: lea edx, var_44
  loc_00E0EE46: push eax
  loc_00E0EE47: push edx
  loc_00E0EE48: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0EE4E: add esp, 00000010h
  loc_00E0EE51: push eax
  loc_00E0EE52: call [00401120h] ; __vbaCastObjVar
  loc_00E0EE58: push eax
  loc_00E0EE59: lea eax, var_28
  loc_00E0EE5C: push eax
  loc_00E0EE5D: call edi
  loc_00E0EE5F: mov ecx, [eax]
  loc_00E0EE61: lea edx, var_2C
  loc_00E0EE64: push edx
  loc_00E0EE65: push eax
  loc_00E0EE66: mov var_BC, eax
  loc_00E0EE6C: call [ecx+00000054h]
  loc_00E0EE6F: test eax, eax
  loc_00E0EE71: fnclex
  loc_00E0EE73: jge 00E0EE86h
  loc_00E0EE75: mov ecx, var_BC
  loc_00E0EE7B: push 00000054h
  loc_00E0EE7D: push 006DCBE8h
  loc_00E0EE82: push ecx
  loc_00E0EE83: push eax
  loc_00E0EE84: call ebx
  loc_00E0EE86: mov ecx, var_2C
  loc_00E0EE89: mov eax, 00000002h
  loc_00E0EE8E: mov var_7C, 00000010h
  loc_00E0EE95: mov var_84, eax
  loc_00E0EE9B: mov edx, [ecx]
  loc_00E0EE9D: mov var_C4, ecx
  loc_00E0EEA3: lea ecx, var_30
  loc_00E0EEA6: push ecx
  loc_00E0EEA7: sub esp, 00000010h
  loc_00E0EEAA: mov ecx, esp
  loc_00E0EEAC: mov [ecx], eax
  loc_00E0EEAE: mov eax, var_80
  loc_00E0EEB1: mov [ecx+00000004h], eax
  loc_00E0EEB4: mov eax, var_7C
  loc_00E0EEB7: mov [ecx+00000008h], eax
  loc_00E0EEBA: mov eax, var_78
  loc_00E0EEBD: mov [ecx+0000000Ch], eax
  loc_00E0EEC0: mov ecx, var_2C
  loc_00E0EEC3: push ecx
  loc_00E0EEC4: call [edx+00000028h]
  loc_00E0EEC7: test eax, eax
  loc_00E0EEC9: fnclex
  loc_00E0EECB: jge 00E0EEDEh
  loc_00E0EECD: mov edx, var_C4
  loc_00E0EED3: push 00000028h
  loc_00E0EED5: push 006E09E8h
  loc_00E0EEDA: push edx
  loc_00E0EEDB: push eax
  loc_00E0EEDC: call ebx
  loc_00E0EEDE: sub esp, 00000010h
  loc_00E0EEE1: mov eax, 00000008h
  loc_00E0EEE6: mov edx, esp
  loc_00E0EEE8: mov ecx, var_30
  loc_00E0EEEB: mov var_94, eax
  loc_00E0EEF1: mov var_8C, 006E0BC8h ; " - "
  loc_00E0EEFB: mov [edx], eax
  loc_00E0EEFD: mov eax, var_90
  loc_00E0EF03: mov var_CC, ecx
  loc_00E0EF09: mov ecx, [ecx]
  loc_00E0EF0B: mov [edx+00000004h], eax
  loc_00E0EF0E: mov eax, var_8C
  loc_00E0EF14: mov [edx+00000008h], eax
  loc_00E0EF17: mov eax, var_88
  loc_00E0EF1D: mov [edx+0000000Ch], eax
  loc_00E0EF20: mov edx, var_30
  loc_00E0EF23: push edx
  loc_00E0EF24: call [ecx+00000038h]
  loc_00E0EF27: test eax, eax
  loc_00E0EF29: fnclex
  loc_00E0EF2B: jge 00E0EF3Eh
  loc_00E0EF2D: mov ecx, var_CC
  loc_00E0EF33: push 00000038h
  loc_00E0EF35: push 006E09F8h
  loc_00E0EF3A: push ecx
  loc_00E0EF3B: push eax
  loc_00E0EF3C: call ebx
  loc_00E0EF3E: lea edx, var_30
  loc_00E0EF41: lea eax, var_2C
  loc_00E0EF44: push edx
  loc_00E0EF45: lea ecx, var_28
  loc_00E0EF48: push eax
  loc_00E0EF49: lea edx, var_24
  loc_00E0EF4C: push ecx
  loc_00E0EF4D: push edx
  loc_00E0EF4E: push 00000004h
  loc_00E0EF50: call [00401048h] ; __vbaFreeObjList
  loc_00E0EF56: add esp, 00000014h
  loc_00E0EF59: lea ecx, var_44
  loc_00E0EF5C: call [00401024h] ; __vbaFreeVar
  loc_00E0EF62: mov eax, [esi]
  loc_00E0EF64: push 006DCBD8h
  loc_00E0EF69: push 00000000h
  loc_00E0EF6B: push 00000018h
  loc_00E0EF6D: push esi
  loc_00E0EF6E: call [eax+00000490h]
  loc_00E0EF74: lea ecx, var_24
  loc_00E0EF77: push eax
  loc_00E0EF78: push ecx
  loc_00E0EF79: call edi
  loc_00E0EF7B: lea edx, var_44
  loc_00E0EF7E: push eax
  loc_00E0EF7F: push edx
  loc_00E0EF80: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0EF86: add esp, 00000010h
  loc_00E0EF89: push eax
  loc_00E0EF8A: call [00401120h] ; __vbaCastObjVar
  loc_00E0EF90: push eax
  loc_00E0EF91: lea eax, var_28
  loc_00E0EF94: push eax
  loc_00E0EF95: call edi
  loc_00E0EF97: mov ecx, [eax]
  loc_00E0EF99: lea edx, var_2C
  loc_00E0EF9C: push edx
  loc_00E0EF9D: push eax
  loc_00E0EF9E: mov var_BC, eax
  loc_00E0EFA4: call [ecx+00000054h]
  loc_00E0EFA7: test eax, eax
  loc_00E0EFA9: fnclex
  loc_00E0EFAB: jge 00E0EFBEh
  loc_00E0EFAD: mov ecx, var_BC
  loc_00E0EFB3: push 00000054h
  loc_00E0EFB5: push 006DCBE8h
  loc_00E0EFBA: push ecx
  loc_00E0EFBB: push eax
  loc_00E0EFBC: call ebx
  loc_00E0EFBE: mov ecx, var_2C
  loc_00E0EFC1: mov eax, 00000002h
  loc_00E0EFC6: mov var_7C, 00000011h
  loc_00E0EFCD: mov var_84, eax
  loc_00E0EFD3: mov edx, [ecx]
  loc_00E0EFD5: mov var_C4, ecx
  loc_00E0EFDB: lea ecx, var_30
  loc_00E0EFDE: push ecx
  loc_00E0EFDF: sub esp, 00000010h
  loc_00E0EFE2: mov ecx, esp
  loc_00E0EFE4: mov [ecx], eax
  loc_00E0EFE6: mov eax, var_80
  loc_00E0EFE9: mov [ecx+00000004h], eax
  loc_00E0EFEC: mov eax, var_7C
  loc_00E0EFEF: mov [ecx+00000008h], eax
  loc_00E0EFF2: mov eax, var_78
  loc_00E0EFF5: mov [ecx+0000000Ch], eax
  loc_00E0EFF8: mov ecx, var_2C
  loc_00E0EFFB: push ecx
  loc_00E0EFFC: call [edx+00000028h]
  loc_00E0EFFF: test eax, eax
  loc_00E0F001: fnclex
  loc_00E0F003: jge 00E0F016h
  loc_00E0F005: mov edx, var_C4
  loc_00E0F00B: push 00000028h
  loc_00E0F00D: push 006E09E8h
  loc_00E0F012: push edx
  loc_00E0F013: push eax
  loc_00E0F014: call ebx
  loc_00E0F016: sub esp, 00000010h
  loc_00E0F019: mov eax, 00000008h
  loc_00E0F01E: mov edx, esp
  loc_00E0F020: mov ecx, var_30
  loc_00E0F023: mov var_94, eax
  loc_00E0F029: mov var_8C, 006E0BC8h ; " - "
  loc_00E0F033: mov [edx], eax
  loc_00E0F035: mov eax, var_90
  loc_00E0F03B: mov var_CC, ecx
  loc_00E0F041: mov ecx, [ecx]
  loc_00E0F043: mov [edx+00000004h], eax
  loc_00E0F046: mov eax, var_8C
  loc_00E0F04C: mov [edx+00000008h], eax
  loc_00E0F04F: mov eax, var_88
  loc_00E0F055: mov [edx+0000000Ch], eax
  loc_00E0F058: mov edx, var_30
  loc_00E0F05B: push edx
  loc_00E0F05C: call [ecx+00000038h]
  loc_00E0F05F: test eax, eax
  loc_00E0F061: fnclex
  loc_00E0F063: jge 00E0F076h
  loc_00E0F065: mov ecx, var_CC
  loc_00E0F06B: push 00000038h
  loc_00E0F06D: push 006E09F8h
  loc_00E0F072: push ecx
  loc_00E0F073: push eax
  loc_00E0F074: call ebx
  loc_00E0F076: lea edx, var_30
  loc_00E0F079: lea eax, var_2C
  loc_00E0F07C: push edx
  loc_00E0F07D: lea ecx, var_28
  loc_00E0F080: push eax
  loc_00E0F081: lea edx, var_24
  loc_00E0F084: push ecx
  loc_00E0F085: push edx
  loc_00E0F086: push 00000004h
  loc_00E0F088: call [00401048h] ; __vbaFreeObjList
  loc_00E0F08E: add esp, 00000014h
  loc_00E0F091: lea ecx, var_44
  loc_00E0F094: call [00401024h] ; __vbaFreeVar
  loc_00E0F09A: mov eax, [esi]
  loc_00E0F09C: push 006DCBD8h
  loc_00E0F0A1: push 00000000h
  loc_00E0F0A3: push 00000018h
  loc_00E0F0A5: push esi
  loc_00E0F0A6: call [eax+00000490h]
  loc_00E0F0AC: lea ecx, var_24
  loc_00E0F0AF: push eax
  loc_00E0F0B0: push ecx
  loc_00E0F0B1: call edi
  loc_00E0F0B3: lea edx, var_44
  loc_00E0F0B6: push eax
  loc_00E0F0B7: push edx
  loc_00E0F0B8: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0F0BE: add esp, 00000010h
  loc_00E0F0C1: push eax
  loc_00E0F0C2: call [00401120h] ; __vbaCastObjVar
  loc_00E0F0C8: push eax
  loc_00E0F0C9: lea eax, var_28
  loc_00E0F0CC: push eax
  loc_00E0F0CD: call edi
  loc_00E0F0CF: mov ecx, [eax]
  loc_00E0F0D1: lea edx, var_2C
  loc_00E0F0D4: push edx
  loc_00E0F0D5: push eax
  loc_00E0F0D6: mov var_BC, eax
  loc_00E0F0DC: call [ecx+00000054h]
  loc_00E0F0DF: test eax, eax
  loc_00E0F0E1: fnclex
  loc_00E0F0E3: jge 00E0F0F6h
  loc_00E0F0E5: mov ecx, var_BC
  loc_00E0F0EB: push 00000054h
  loc_00E0F0ED: push 006DCBE8h
  loc_00E0F0F2: push ecx
  loc_00E0F0F3: push eax
  loc_00E0F0F4: call ebx
  loc_00E0F0F6: mov ecx, var_2C
  loc_00E0F0F9: mov eax, 00000002h
  loc_00E0F0FE: mov var_7C, 00000012h
  loc_00E0F105: mov var_84, eax
  loc_00E0F10B: mov edx, [ecx]
  loc_00E0F10D: mov var_C4, ecx
  loc_00E0F113: lea ecx, var_30
  loc_00E0F116: push ecx
  loc_00E0F117: sub esp, 00000010h
  loc_00E0F11A: mov ecx, esp
  loc_00E0F11C: mov [ecx], eax
  loc_00E0F11E: mov eax, var_80
  loc_00E0F121: mov [ecx+00000004h], eax
  loc_00E0F124: mov eax, var_7C
  loc_00E0F127: mov [ecx+00000008h], eax
  loc_00E0F12A: mov eax, var_78
  loc_00E0F12D: mov [ecx+0000000Ch], eax
  loc_00E0F130: mov ecx, var_2C
  loc_00E0F133: push ecx
  loc_00E0F134: call [edx+00000028h]
  loc_00E0F137: test eax, eax
  loc_00E0F139: fnclex
  loc_00E0F13B: jge 00E0F14Eh
  loc_00E0F13D: mov edx, var_C4
  loc_00E0F143: push 00000028h
  loc_00E0F145: push 006E09E8h
  loc_00E0F14A: push edx
  loc_00E0F14B: push eax
  loc_00E0F14C: call ebx
  loc_00E0F14E: sub esp, 00000010h
  loc_00E0F151: mov eax, 00000008h
  loc_00E0F156: mov edx, esp
  loc_00E0F158: mov ecx, var_30
  loc_00E0F15B: mov var_94, eax
  loc_00E0F161: mov var_8C, 006E0BC8h ; " - "
  loc_00E0F16B: mov [edx], eax
  loc_00E0F16D: mov eax, var_90
  loc_00E0F173: mov var_CC, ecx
  loc_00E0F179: mov ecx, [ecx]
  loc_00E0F17B: mov [edx+00000004h], eax
  loc_00E0F17E: mov eax, var_8C
  loc_00E0F184: mov [edx+00000008h], eax
  loc_00E0F187: mov eax, var_88
  loc_00E0F18D: mov [edx+0000000Ch], eax
  loc_00E0F190: mov edx, var_30
  loc_00E0F193: push edx
  loc_00E0F194: call [ecx+00000038h]
  loc_00E0F197: test eax, eax
  loc_00E0F199: fnclex
  loc_00E0F19B: jge 00E0F1AEh
  loc_00E0F19D: mov ecx, var_CC
  loc_00E0F1A3: push 00000038h
  loc_00E0F1A5: push 006E09F8h
  loc_00E0F1AA: push ecx
  loc_00E0F1AB: push eax
  loc_00E0F1AC: call ebx
  loc_00E0F1AE: lea edx, var_30
  loc_00E0F1B1: lea eax, var_2C
  loc_00E0F1B4: push edx
  loc_00E0F1B5: lea ecx, var_28
  loc_00E0F1B8: push eax
  loc_00E0F1B9: lea edx, var_24
  loc_00E0F1BC: push ecx
  loc_00E0F1BD: push edx
  loc_00E0F1BE: push 00000004h
  loc_00E0F1C0: call [00401048h] ; __vbaFreeObjList
  loc_00E0F1C6: add esp, 00000014h
  loc_00E0F1C9: lea ecx, var_44
  loc_00E0F1CC: call [00401024h] ; __vbaFreeVar
  loc_00E0F1D2: mov eax, [esi]
  loc_00E0F1D4: push 006DCBD8h
  loc_00E0F1D9: push 00000000h
  loc_00E0F1DB: push 00000018h
  loc_00E0F1DD: push esi
  loc_00E0F1DE: call [eax+00000490h]
  loc_00E0F1E4: lea ecx, var_24
  loc_00E0F1E7: push eax
  loc_00E0F1E8: push ecx
  loc_00E0F1E9: call edi
  loc_00E0F1EB: lea edx, var_44
  loc_00E0F1EE: push eax
  loc_00E0F1EF: push edx
  loc_00E0F1F0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0F1F6: add esp, 00000010h
  loc_00E0F1F9: push eax
  loc_00E0F1FA: call [00401120h] ; __vbaCastObjVar
  loc_00E0F200: push eax
  loc_00E0F201: lea eax, var_28
  loc_00E0F204: push eax
  loc_00E0F205: call edi
  loc_00E0F207: mov ecx, 0000000Ah
  loc_00E0F20C: mov var_8C, 80020004h
  loc_00E0F216: mov var_94, ecx
  loc_00E0F21C: mov var_7C, 80020004h
  loc_00E0F223: mov var_84, ecx
  loc_00E0F229: mov edx, [eax]
  loc_00E0F22B: sub esp, 00000010h
  loc_00E0F22E: mov var_BC, eax
  loc_00E0F234: mov eax, esp
  loc_00E0F236: sub esp, 00000010h
  loc_00E0F239: mov [eax], ecx
  loc_00E0F23B: mov ecx, var_90
  loc_00E0F241: mov [eax+00000004h], ecx
  loc_00E0F244: mov ecx, var_8C
  loc_00E0F24A: mov [eax+00000008h], ecx
  loc_00E0F24D: mov ecx, var_88
  loc_00E0F253: mov [eax+0000000Ch], ecx
  loc_00E0F256: mov ecx, var_84
  loc_00E0F25C: mov eax, esp
  loc_00E0F25E: mov [eax], ecx
  loc_00E0F260: mov ecx, var_80
  loc_00E0F263: mov [eax+00000004h], ecx
  loc_00E0F266: mov ecx, var_7C
  loc_00E0F269: mov [eax+00000008h], ecx
  loc_00E0F26C: mov ecx, var_78
  loc_00E0F26F: mov [eax+0000000Ch], ecx
  loc_00E0F272: mov eax, var_BC
  loc_00E0F278: push eax
  loc_00E0F279: call [edx+000000ACh]
  loc_00E0F27F: test eax, eax
  loc_00E0F281: fnclex
  loc_00E0F283: jge 00E0F299h
  loc_00E0F285: mov ecx, var_BC
  loc_00E0F28B: push 000000ACh
  loc_00E0F290: push 006DCBE8h
  loc_00E0F295: push ecx
  loc_00E0F296: push eax
  loc_00E0F297: call ebx
  loc_00E0F299: lea edx, var_28
  loc_00E0F29C: lea eax, var_24
  loc_00E0F29F: push edx
  loc_00E0F2A0: push eax
  loc_00E0F2A1: push 00000002h
  loc_00E0F2A3: call [00401048h] ; __vbaFreeObjList
  loc_00E0F2A9: add esp, 0000000Ch
  loc_00E0F2AC: lea ecx, var_44
  loc_00E0F2AF: call [00401024h] ; __vbaFreeVar
  loc_00E0F2B5: mov ecx, [esi]
  loc_00E0F2B7: push 00000000h
  loc_00E0F2B9: push 00000019h
  loc_00E0F2BB: push esi
  loc_00E0F2BC: call [ecx+00000490h]
  loc_00E0F2C2: lea edx, var_24
  loc_00E0F2C5: push eax
  loc_00E0F2C6: push edx
  loc_00E0F2C7: call edi
  loc_00E0F2C9: push eax
  loc_00E0F2CA: call [00401030h] ; __vbaLateIdCall
  loc_00E0F2D0: add esp, 0000000Ch
  loc_00E0F2D3: lea ecx, var_24
  loc_00E0F2D6: call [00401254h] ; __vbaFreeObj
  loc_00E0F2DC: mov eax, [esi]
  loc_00E0F2DE: push esi
  loc_00E0F2DF: call [eax+00000334h]
  loc_00E0F2E5: lea ecx, var_24
  loc_00E0F2E8: push eax
  loc_00E0F2E9: push ecx
  loc_00E0F2EA: call edi
  loc_00E0F2EC: mov edx, [eax]
  loc_00E0F2EE: push FFFFFFFFh
  loc_00E0F2F0: push eax
  loc_00E0F2F1: mov var_BC, eax
  loc_00E0F2F7: call [edx+0000005Ch]
  loc_00E0F2FA: test eax, eax
  loc_00E0F2FC: fnclex
  loc_00E0F2FE: jge 00E0F311h
  loc_00E0F300: mov ecx, var_BC
  loc_00E0F306: push 0000005Ch
  loc_00E0F308: push 006DCAE0h
  loc_00E0F30D: push ecx
  loc_00E0F30E: push eax
  loc_00E0F30F: call ebx
  loc_00E0F311: lea ecx, var_24
  loc_00E0F314: call [00401254h] ; __vbaFreeObj
  loc_00E0F31A: mov edx, [esi]
  loc_00E0F31C: push esi
  loc_00E0F31D: call [edx+000003E8h]
  loc_00E0F323: push eax
  loc_00E0F324: lea eax, var_24
  loc_00E0F327: push eax
  loc_00E0F328: call edi
  loc_00E0F32A: mov esi, eax
  loc_00E0F32C: push 00000000h
  loc_00E0F32E: push esi
  loc_00E0F32F: mov ecx, [esi]
  loc_00E0F331: call [ecx+0000008Ch]
  loc_00E0F337: test eax, eax
  loc_00E0F339: fnclex
  loc_00E0F33B: jge 00E0F34Bh
  loc_00E0F33D: push 0000008Ch
  loc_00E0F342: push 006DCDA0h
  loc_00E0F347: push esi
  loc_00E0F348: push eax
  loc_00E0F349: call ebx
  loc_00E0F34B: lea ecx, var_24
  loc_00E0F34E: call [00401254h] ; __vbaFreeObj
  loc_00E0F354: jmp 00E0F3C5h
  loc_00E0F356: mov ecx, 80020004h
  loc_00E0F35B: mov eax, 0000000Ah
  loc_00E0F360: mov var_6C, ecx
  loc_00E0F363: mov var_5C, ecx
  loc_00E0F366: mov var_4C, ecx
  loc_00E0F369: lea edx, var_84
  loc_00E0F36F: lea ecx, var_44
  loc_00E0F372: mov var_74, eax
  loc_00E0F375: mov var_64, eax
  loc_00E0F378: mov var_54, eax
  loc_00E0F37B: mov var_7C, 006E0BD4h ; "Data Sudah Ada !"
  loc_00E0F382: mov var_84, 00000008h
  loc_00E0F38C: call [004011F0h] ; __vbaVarDup
  loc_00E0F392: lea edx, var_74
  loc_00E0F395: lea eax, var_64
  loc_00E0F398: push edx
  loc_00E0F399: lea ecx, var_54
  loc_00E0F39C: push eax
  loc_00E0F39D: push ecx
  loc_00E0F39E: push 00000000h
  loc_00E0F3A0: lea edx, var_44
  loc_00E0F3A3: push edx
  loc_00E0F3A4: call [004010A8h] ; rtcMsgBox
  loc_00E0F3AA: lea eax, var_74
  loc_00E0F3AD: lea ecx, var_64
  loc_00E0F3B0: push eax
  loc_00E0F3B1: lea edx, var_54
  loc_00E0F3B4: push ecx
  loc_00E0F3B5: lea eax, var_44
  loc_00E0F3B8: push edx
  loc_00E0F3B9: push eax
  loc_00E0F3BA: push 00000004h
  loc_00E0F3BC: call [00401038h] ; __vbaFreeVarList
  loc_00E0F3C2: add esp, 00000014h
  loc_00E0F3C5: mov var_4, 00000000h
  loc_00E0F3CC: fwait
  loc_00E0F3CD: push 00E0F421h
  loc_00E0F3D2: jmp 00E0F420h
  loc_00E0F3D4: lea ecx, var_20
  loc_00E0F3D7: lea edx, var_1C
  loc_00E0F3DA: push ecx
  loc_00E0F3DB: lea eax, var_18
  loc_00E0F3DE: push edx
  loc_00E0F3DF: push eax
  loc_00E0F3E0: push 00000003h
  loc_00E0F3E2: call [004011BCh] ; __vbaFreeStrList
  loc_00E0F3E8: lea ecx, var_34
  loc_00E0F3EB: lea edx, var_30
  loc_00E0F3EE: push ecx
  loc_00E0F3EF: lea eax, var_2C
  loc_00E0F3F2: push edx
  loc_00E0F3F3: lea ecx, var_28
  loc_00E0F3F6: push eax
  loc_00E0F3F7: lea edx, var_24
  loc_00E0F3FA: push ecx
  loc_00E0F3FB: push edx
  loc_00E0F3FC: push 00000005h
  loc_00E0F3FE: call [00401048h] ; __vbaFreeObjList
  loc_00E0F404: lea eax, var_74
  loc_00E0F407: lea ecx, var_64
  loc_00E0F40A: push eax
  loc_00E0F40B: lea edx, var_54
  loc_00E0F40E: push ecx
  loc_00E0F40F: lea eax, var_44
  loc_00E0F412: push edx
  loc_00E0F413: push eax
  loc_00E0F414: push 00000004h
  loc_00E0F416: call [00401038h] ; __vbaFreeVarList
  loc_00E0F41C: add esp, 0000003Ch
  loc_00E0F41F: ret
  loc_00E0F420: ret
  loc_00E0F421: mov eax, Me
  loc_00E0F424: push eax
  loc_00E0F425: mov ecx, [eax]
  loc_00E0F427: call [ecx+00000008h]
  loc_00E0F42A: mov eax, var_4
  loc_00E0F42D: mov ecx, var_14
  loc_00E0F430: pop edi
  loc_00E0F431: pop esi
  loc_00E0F432: mov fs:[00000000h], ecx
  loc_00E0F439: pop ebx
  loc_00E0F43A: mov esp, ebp
  loc_00E0F43C: pop ebp
  loc_00E0F43D: retn 0004h
End Sub

Private Sub nowali_KeyPress(KeyAscii As Integer) 'E09D60
  loc_00E09D60: push ebp
  loc_00E09D61: mov ebp, esp
  loc_00E09D63: sub esp, 0000000Ch
  loc_00E09D66: push 00402836h ; __vbaExceptHandler
  loc_00E09D6B: mov eax, fs:[00000000h]
  loc_00E09D71: push eax
  loc_00E09D72: mov fs:[00000000h], esp
  loc_00E09D79: sub esp, 00000014h
  loc_00E09D7C: push ebx
  loc_00E09D7D: push esi
  loc_00E09D7E: push edi
  loc_00E09D7F: mov var_C, esp
  loc_00E09D82: mov var_8, 00402170h
  loc_00E09D89: mov esi, Me
  loc_00E09D8C: mov eax, esi
  loc_00E09D8E: and eax, 00000001h
  loc_00E09D91: mov var_4, eax
  loc_00E09D94: and esi, FFFFFFFEh
  loc_00E09D97: push esi
  loc_00E09D98: mov Me, esi
  loc_00E09D9B: mov ecx, [esi]
  loc_00E09D9D: call [ecx+00000004h]
  loc_00E09DA0: mov edx, KeyAscii
  loc_00E09DA3: xor edi, edi
  loc_00E09DA5: mov var_18, edi
  loc_00E09DA8: cmp [edx], 000Dh
  loc_00E09DAC: jnz 00E09DEEh
  loc_00E09DAE: mov eax, [esi]
  loc_00E09DB0: push esi
  loc_00E09DB1: call [eax+00000348h]
  loc_00E09DB7: lea ecx, var_18
  loc_00E09DBA: push eax
  loc_00E09DBB: push ecx
  loc_00E09DBC: call [004010ACh] ; __vbaObjSet
  loc_00E09DC2: mov esi, eax
  loc_00E09DC4: push esi
  loc_00E09DC5: mov edx, [esi]
  loc_00E09DC7: call [edx+00000204h]
  loc_00E09DCD: cmp eax, edi
  loc_00E09DCF: fnclex
  loc_00E09DD1: jge 00E09DE5h
  loc_00E09DD3: push 00000204h
  loc_00E09DD8: push 006DCB70h
  loc_00E09DDD: push esi
  loc_00E09DDE: push eax
  loc_00E09DDF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09DE5: lea ecx, var_18
  loc_00E09DE8: call [00401254h] ; __vbaFreeObj
  loc_00E09DEE: mov var_4, edi
  loc_00E09DF1: push 00E09E03h
  loc_00E09DF6: jmp 00E09E02h
  loc_00E09DF8: lea ecx, var_18
  loc_00E09DFB: call [00401254h] ; __vbaFreeObj
  loc_00E09E01: ret
  loc_00E09E02: ret
  loc_00E09E03: mov eax, Me
  loc_00E09E06: push eax
  loc_00E09E07: mov ecx, [eax]
  loc_00E09E09: call [ecx+00000008h]
  loc_00E09E0C: mov eax, var_4
  loc_00E09E0F: mov ecx, var_14
  loc_00E09E12: pop edi
  loc_00E09E13: pop esi
  loc_00E09E14: mov fs:[00000000h], ecx
  loc_00E09E1B: pop ebx
  loc_00E09E1C: mov esp, ebp
  loc_00E09E1E: pop ebp
  loc_00E09E1F: retn 0008h
End Sub

Private Sub nisn_KeyPress(KeyAscii As Integer) 'E09C90
  loc_00E09C90: push ebp
  loc_00E09C91: mov ebp, esp
  loc_00E09C93: sub esp, 0000000Ch
  loc_00E09C96: push 00402836h ; __vbaExceptHandler
  loc_00E09C9B: mov eax, fs:[00000000h]
  loc_00E09CA1: push eax
  loc_00E09CA2: mov fs:[00000000h], esp
  loc_00E09CA9: sub esp, 00000014h
  loc_00E09CAC: push ebx
  loc_00E09CAD: push esi
  loc_00E09CAE: push edi
  loc_00E09CAF: mov var_C, esp
  loc_00E09CB2: mov var_8, 00402160h
  loc_00E09CB9: mov esi, Me
  loc_00E09CBC: mov eax, esi
  loc_00E09CBE: and eax, 00000001h
  loc_00E09CC1: mov var_4, eax
  loc_00E09CC4: and esi, FFFFFFFEh
  loc_00E09CC7: push esi
  loc_00E09CC8: mov Me, esi
  loc_00E09CCB: mov ecx, [esi]
  loc_00E09CCD: call [ecx+00000004h]
  loc_00E09CD0: mov edx, KeyAscii
  loc_00E09CD3: xor edi, edi
  loc_00E09CD5: mov var_18, edi
  loc_00E09CD8: cmp [edx], 000Dh
  loc_00E09CDC: jnz 00E09D1Eh
  loc_00E09CDE: mov eax, [esi]
  loc_00E09CE0: push esi
  loc_00E09CE1: call [eax+00000388h]
  loc_00E09CE7: lea ecx, var_18
  loc_00E09CEA: push eax
  loc_00E09CEB: push ecx
  loc_00E09CEC: call [004010ACh] ; __vbaObjSet
  loc_00E09CF2: mov esi, eax
  loc_00E09CF4: push esi
  loc_00E09CF5: mov edx, [esi]
  loc_00E09CF7: call [edx+00000204h]
  loc_00E09CFD: cmp eax, edi
  loc_00E09CFF: fnclex
  loc_00E09D01: jge 00E09D15h
  loc_00E09D03: push 00000204h
  loc_00E09D08: push 006DCB70h
  loc_00E09D0D: push esi
  loc_00E09D0E: push eax
  loc_00E09D0F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09D15: lea ecx, var_18
  loc_00E09D18: call [00401254h] ; __vbaFreeObj
  loc_00E09D1E: mov var_4, edi
  loc_00E09D21: push 00E09D33h
  loc_00E09D26: jmp 00E09D32h
  loc_00E09D28: lea ecx, var_18
  loc_00E09D2B: call [00401254h] ; __vbaFreeObj
  loc_00E09D31: ret
  loc_00E09D32: ret
  loc_00E09D33: mov eax, Me
  loc_00E09D36: push eax
  loc_00E09D37: mov ecx, [eax]
  loc_00E09D39: call [ecx+00000008h]
  loc_00E09D3C: mov eax, var_4
  loc_00E09D3F: mov ecx, var_14
  loc_00E09D42: pop edi
  loc_00E09D43: pop esi
  loc_00E09D44: mov fs:[00000000h], ecx
  loc_00E09D4B: pop ebx
  loc_00E09D4C: mov esp, ebp
  loc_00E09D4E: pop ebp
  loc_00E09D4F: retn 0008h
End Sub

Private Sub payah_KeyPress(KeyAscii As Integer) 'E0A520
  loc_00E0A520: push ebp
  loc_00E0A521: mov ebp, esp
  loc_00E0A523: sub esp, 0000000Ch
  loc_00E0A526: push 00402836h ; __vbaExceptHandler
  loc_00E0A52B: mov eax, fs:[00000000h]
  loc_00E0A531: push eax
  loc_00E0A532: mov fs:[00000000h], esp
  loc_00E0A539: sub esp, 00000014h
  loc_00E0A53C: push ebx
  loc_00E0A53D: push esi
  loc_00E0A53E: push edi
  loc_00E0A53F: mov var_C, esp
  loc_00E0A542: mov var_8, 004021D0h
  loc_00E0A549: mov esi, Me
  loc_00E0A54C: mov eax, esi
  loc_00E0A54E: and eax, 00000001h
  loc_00E0A551: mov var_4, eax
  loc_00E0A554: and esi, FFFFFFFEh
  loc_00E0A557: push esi
  loc_00E0A558: mov Me, esi
  loc_00E0A55B: mov ecx, [esi]
  loc_00E0A55D: call [ecx+00000004h]
  loc_00E0A560: mov edx, KeyAscii
  loc_00E0A563: xor edi, edi
  loc_00E0A565: mov var_18, edi
  loc_00E0A568: cmp [edx], 000Dh
  loc_00E0A56C: jnz 00E0A5AEh
  loc_00E0A56E: mov eax, [esi]
  loc_00E0A570: push esi
  loc_00E0A571: call [eax+00000360h]
  loc_00E0A577: lea ecx, var_18
  loc_00E0A57A: push eax
  loc_00E0A57B: push ecx
  loc_00E0A57C: call [004010ACh] ; __vbaObjSet
  loc_00E0A582: mov esi, eax
  loc_00E0A584: push esi
  loc_00E0A585: mov edx, [esi]
  loc_00E0A587: call [edx+00000204h]
  loc_00E0A58D: cmp eax, edi
  loc_00E0A58F: fnclex
  loc_00E0A591: jge 00E0A5A5h
  loc_00E0A593: push 00000204h
  loc_00E0A598: push 006DCB70h
  loc_00E0A59D: push esi
  loc_00E0A59E: push eax
  loc_00E0A59F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A5A5: lea ecx, var_18
  loc_00E0A5A8: call [00401254h] ; __vbaFreeObj
  loc_00E0A5AE: mov var_4, edi
  loc_00E0A5B1: push 00E0A5C3h
  loc_00E0A5B6: jmp 00E0A5C2h
  loc_00E0A5B8: lea ecx, var_18
  loc_00E0A5BB: call [00401254h] ; __vbaFreeObj
  loc_00E0A5C1: ret
  loc_00E0A5C2: ret
  loc_00E0A5C3: mov eax, Me
  loc_00E0A5C6: push eax
  loc_00E0A5C7: mov ecx, [eax]
  loc_00E0A5C9: call [ecx+00000008h]
  loc_00E0A5CC: mov eax, var_4
  loc_00E0A5CF: mov ecx, var_14
  loc_00E0A5D2: pop edi
  loc_00E0A5D3: pop esi
  loc_00E0A5D4: mov fs:[00000000h], ecx
  loc_00E0A5DB: pop ebx
  loc_00E0A5DC: mov esp, ebp
  loc_00E0A5DE: pop ebp
  loc_00E0A5DF: retn 0008h
End Sub

Private Sub refreshh_UnknownEvent_9 'E0C1C0
  loc_00E0C1C0: push ebp
  loc_00E0C1C1: mov ebp, esp
  loc_00E0C1C3: sub esp, 0000000Ch
  loc_00E0C1C6: push 00402836h ; __vbaExceptHandler
  loc_00E0C1CB: mov eax, fs:[00000000h]
  loc_00E0C1D1: push eax
  loc_00E0C1D2: mov fs:[00000000h], esp
  loc_00E0C1D9: sub esp, 00000040h
  loc_00E0C1DC: push ebx
  loc_00E0C1DD: push esi
  loc_00E0C1DE: push edi
  loc_00E0C1DF: mov var_C, esp
  loc_00E0C1E2: mov var_8, 00402250h
  loc_00E0C1E9: mov esi, Me
  loc_00E0C1EC: mov eax, esi
  loc_00E0C1EE: and eax, 00000001h
  loc_00E0C1F1: mov var_4, eax
  loc_00E0C1F4: and esi, FFFFFFFEh
  loc_00E0C1F7: push esi
  loc_00E0C1F8: mov Me, esi
  loc_00E0C1FB: mov ecx, [esi]
  loc_00E0C1FD: call [ecx+00000004h]
  loc_00E0C200: sub esp, 00000010h
  loc_00E0C203: mov ecx, 00000008h
  loc_00E0C208: mov edx, esp
  loc_00E0C20A: xor eax, eax
  loc_00E0C20C: mov var_18, eax
  loc_00E0C20F: mov var_1C, eax
  loc_00E0C212: mov [edx], ecx
  loc_00E0C214: mov ecx, var_38
  loc_00E0C217: mov var_2C, eax
  loc_00E0C21A: mov eax, 006E0A78h ; "biodata"
  loc_00E0C21F: mov [edx+00000004h], ecx
  loc_00E0C222: mov ecx, [esi]
  loc_00E0C224: push 0000000Eh
  loc_00E0C226: push esi
  loc_00E0C227: mov [edx+00000008h], eax
  loc_00E0C22A: mov eax, var_30
  loc_00E0C22D: mov [edx+0000000Ch], eax
  loc_00E0C230: call [ecx+00000490h]
  loc_00E0C236: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E0C23C: lea edx, var_18
  loc_00E0C23F: push eax
  loc_00E0C240: push edx
  loc_00E0C241: call edi
  loc_00E0C243: push eax
  loc_00E0C244: call [00401238h] ; __vbaLateIdSt
  loc_00E0C24A: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E0C250: lea ecx, var_18
  loc_00E0C253: call ebx
  loc_00E0C255: mov eax, [esi]
  loc_00E0C257: push 00000000h
  loc_00E0C259: push 00000019h
  loc_00E0C25B: push esi
  loc_00E0C25C: call [eax+00000490h]
  loc_00E0C262: lea ecx, var_18
  loc_00E0C265: push eax
  loc_00E0C266: push ecx
  loc_00E0C267: call edi
  loc_00E0C269: push eax
  loc_00E0C26A: call [00401030h] ; __vbaLateIdCall
  loc_00E0C270: add esp, 0000000Ch
  loc_00E0C273: lea ecx, var_18
  loc_00E0C276: call ebx
  loc_00E0C278: mov edx, [esi]
  loc_00E0C27A: push 006E05D8h
  loc_00E0C27F: push esi
  loc_00E0C280: call [edx+00000490h]
  loc_00E0C286: push eax
  loc_00E0C287: lea eax, var_18
  loc_00E0C28A: push eax
  loc_00E0C28B: call edi
  loc_00E0C28D: push eax
  loc_00E0C28E: call [00401224h] ; __vbaCastObj
  loc_00E0C294: lea ecx, var_2C
  loc_00E0C297: mov var_24, eax
  loc_00E0C29A: push ecx
  loc_00E0C29B: mov var_2C, 0000000Dh
  loc_00E0C2A2: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E0C2A8: mov ecx, [eax]
  loc_00E0C2AA: sub esp, 00000010h
  loc_00E0C2AD: mov edx, esp
  loc_00E0C2AF: mov [edx], ecx
  loc_00E0C2B1: mov ecx, [eax+00000004h]
  loc_00E0C2B4: mov [edx+00000004h], ecx
  loc_00E0C2B7: mov ecx, [eax+00000008h]
  loc_00E0C2BA: mov eax, [eax+0000000Ch]
  loc_00E0C2BD: mov [edx+00000008h], ecx
  loc_00E0C2C0: mov [edx+0000000Ch], eax
  loc_00E0C2C3: mov ecx, [esi]
  loc_00E0C2C5: push 00000000h
  loc_00E0C2C7: push 0000002Ah
  loc_00E0C2C9: push esi
  loc_00E0C2CA: call [ecx+00000494h]
  loc_00E0C2D0: lea edx, var_1C
  loc_00E0C2D3: push eax
  loc_00E0C2D4: push edx
  loc_00E0C2D5: call edi
  loc_00E0C2D7: push eax
  loc_00E0C2D8: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E0C2DE: lea eax, var_1C
  loc_00E0C2E1: lea ecx, var_18
  loc_00E0C2E4: push eax
  loc_00E0C2E5: push ecx
  loc_00E0C2E6: push 00000002h
  loc_00E0C2E8: call [00401048h] ; __vbaFreeObjList
  loc_00E0C2EE: add esp, 00000028h
  loc_00E0C2F1: lea ecx, var_2C
  loc_00E0C2F4: call [00401024h] ; __vbaFreeVar
  loc_00E0C2FA: sub esp, 00000010h
  loc_00E0C2FD: mov ecx, 0000000Bh
  loc_00E0C302: mov edx, esp
  loc_00E0C304: xor eax, eax
  loc_00E0C306: push 00000006h
  loc_00E0C308: push esi
  loc_00E0C309: mov [edx], ecx
  loc_00E0C30B: mov ecx, var_38
  loc_00E0C30E: mov [edx+00000004h], ecx
  loc_00E0C311: mov ecx, [esi]
  loc_00E0C313: mov [edx+00000008h], eax
  loc_00E0C316: mov eax, var_30
  loc_00E0C319: mov [edx+0000000Ch], eax
  loc_00E0C31C: call [ecx+00000494h]
  loc_00E0C322: lea edx, var_18
  loc_00E0C325: push eax
  loc_00E0C326: push edx
  loc_00E0C327: call edi
  loc_00E0C329: push eax
  loc_00E0C32A: call [00401238h] ; __vbaLateIdSt
  loc_00E0C330: lea ecx, var_18
  loc_00E0C333: call ebx
  loc_00E0C335: sub esp, 00000010h
  loc_00E0C338: mov ecx, 0000000Bh
  loc_00E0C33D: mov edx, esp
  loc_00E0C33F: xor eax, eax
  loc_00E0C341: push 8001000Eh
  loc_00E0C346: push esi
  loc_00E0C347: mov [edx], ecx
  loc_00E0C349: mov ecx, var_38
  loc_00E0C34C: mov [edx+00000004h], ecx
  loc_00E0C34F: mov ecx, [esi]
  loc_00E0C351: mov [edx+00000008h], eax
  loc_00E0C354: mov eax, var_30
  loc_00E0C357: mov [edx+0000000Ch], eax
  loc_00E0C35A: call [ecx+00000494h]
  loc_00E0C360: lea edx, var_18
  loc_00E0C363: push eax
  loc_00E0C364: push edx
  loc_00E0C365: call edi
  loc_00E0C367: push eax
  loc_00E0C368: call [00401238h] ; __vbaLateIdSt
  loc_00E0C36E: lea ecx, var_18
  loc_00E0C371: call ebx
  loc_00E0C373: mov eax, [esi]
  loc_00E0C375: push 00000000h
  loc_00E0C377: push FFFFFDDAh
  loc_00E0C37C: push esi
  loc_00E0C37D: call [eax+00000494h]
  loc_00E0C383: lea ecx, var_18
  loc_00E0C386: push eax
  loc_00E0C387: push ecx
  loc_00E0C388: call edi
  loc_00E0C38A: push eax
  loc_00E0C38B: call [00401030h] ; __vbaLateIdCall
  loc_00E0C391: add esp, 0000000Ch
  loc_00E0C394: lea ecx, var_18
  loc_00E0C397: call ebx
  loc_00E0C399: mov var_4, 00000000h
  loc_00E0C3A0: push 00E0C3C5h
  loc_00E0C3A5: jmp 00E0C3C4h
  loc_00E0C3A7: lea edx, var_1C
  loc_00E0C3AA: lea eax, var_18
  loc_00E0C3AD: push edx
  loc_00E0C3AE: push eax
  loc_00E0C3AF: push 00000002h
  loc_00E0C3B1: call [00401048h] ; __vbaFreeObjList
  loc_00E0C3B7: add esp, 0000000Ch
  loc_00E0C3BA: lea ecx, var_2C
  loc_00E0C3BD: call [00401024h] ; __vbaFreeVar
  loc_00E0C3C3: ret
  loc_00E0C3C4: ret
  loc_00E0C3C5: mov eax, Me
  loc_00E0C3C8: push eax
  loc_00E0C3C9: mov ecx, [eax]
  loc_00E0C3CB: call [ecx+00000008h]
  loc_00E0C3CE: mov eax, var_4
  loc_00E0C3D1: mov ecx, var_14
  loc_00E0C3D4: pop edi
  loc_00E0C3D5: pop esi
  loc_00E0C3D6: mov fs:[00000000h], ecx
  loc_00E0C3DD: pop ebx
  loc_00E0C3DE: mov esp, ebp
  loc_00E0C3E0: pop ebp
  loc_00E0C3E1: retn 0004h
End Sub

Private Sub hapus_UnknownEvent_9 'E09340
  loc_00E09340: push ebp
  loc_00E09341: mov ebp, esp
  loc_00E09343: sub esp, 0000000Ch
  loc_00E09346: push 00402836h ; __vbaExceptHandler
  loc_00E0934B: mov eax, fs:[00000000h]
  loc_00E09351: push eax
  loc_00E09352: mov fs:[00000000h], esp
  loc_00E09359: sub esp, 00000028h
  loc_00E0935C: push ebx
  loc_00E0935D: push esi
  loc_00E0935E: push edi
  loc_00E0935F: mov var_C, esp
  loc_00E09362: mov var_8, 004020E0h
  loc_00E09369: mov esi, Me
  loc_00E0936C: mov eax, esi
  loc_00E0936E: and eax, 00000001h
  loc_00E09371: mov var_4, eax
  loc_00E09374: and esi, FFFFFFFEh
  loc_00E09377: push esi
  loc_00E09378: mov Me, esi
  loc_00E0937B: mov ecx, [esi]
  loc_00E0937D: call [ecx+00000004h]
  loc_00E09380: mov edx, [esi]
  loc_00E09382: xor eax, eax
  loc_00E09384: push 006DCBD8h
  loc_00E09389: push eax
  loc_00E0938A: push 00000018h
  loc_00E0938C: push esi
  loc_00E0938D: mov var_18, eax
  loc_00E09390: mov var_1C, eax
  loc_00E09393: mov var_2C, eax
  loc_00E09396: call [edx+00000490h]
  loc_00E0939C: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E093A2: push eax
  loc_00E093A3: lea eax, var_18
  loc_00E093A6: push eax
  loc_00E093A7: call ebx
  loc_00E093A9: lea ecx, var_2C
  loc_00E093AC: push eax
  loc_00E093AD: push ecx
  loc_00E093AE: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E093B4: add esp, 00000010h
  loc_00E093B7: push eax
  loc_00E093B8: call [00401120h] ; __vbaCastObjVar
  loc_00E093BE: lea edx, var_1C
  loc_00E093C1: push eax
  loc_00E093C2: push edx
  loc_00E093C3: call ebx
  loc_00E093C5: mov edi, eax
  loc_00E093C7: push 00000001h
  loc_00E093C9: push edi
  loc_00E093CA: mov eax, [edi]
  loc_00E093CC: call [eax+00000084h]
  loc_00E093D2: test eax, eax
  loc_00E093D4: fnclex
  loc_00E093D6: jge 00E093EAh
  loc_00E093D8: push 00000084h
  loc_00E093DD: push 006DCBE8h
  loc_00E093E2: push edi
  loc_00E093E3: push eax
  loc_00E093E4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E093EA: lea ecx, var_1C
  loc_00E093ED: lea edx, var_18
  loc_00E093F0: push ecx
  loc_00E093F1: push edx
  loc_00E093F2: push 00000002h
  loc_00E093F4: call [00401048h] ; __vbaFreeObjList
  loc_00E093FA: add esp, 0000000Ch
  loc_00E093FD: lea ecx, var_2C
  loc_00E09400: call [00401024h] ; __vbaFreeVar
  loc_00E09406: mov eax, [esi]
  loc_00E09408: push 00000000h
  loc_00E0940A: push FFFFFDDAh
  loc_00E0940F: push esi
  loc_00E09410: call [eax+00000494h]
  loc_00E09416: lea ecx, var_18
  loc_00E09419: push eax
  loc_00E0941A: push ecx
  loc_00E0941B: call ebx
  loc_00E0941D: push eax
  loc_00E0941E: call [00401030h] ; __vbaLateIdCall
  loc_00E09424: add esp, 0000000Ch
  loc_00E09427: lea ecx, var_18
  loc_00E0942A: call [00401254h] ; __vbaFreeObj
  loc_00E09430: mov var_4, 00000000h
  loc_00E09437: push 00E0945Ch
  loc_00E0943C: jmp 00E0945Bh
  loc_00E0943E: lea edx, var_1C
  loc_00E09441: lea eax, var_18
  loc_00E09444: push edx
  loc_00E09445: push eax
  loc_00E09446: push 00000002h
  loc_00E09448: call [00401048h] ; __vbaFreeObjList
  loc_00E0944E: add esp, 0000000Ch
  loc_00E09451: lea ecx, var_2C
  loc_00E09454: call [00401024h] ; __vbaFreeVar
  loc_00E0945A: ret
  loc_00E0945B: ret
  loc_00E0945C: mov eax, Me
  loc_00E0945F: push eax
  loc_00E09460: mov ecx, [eax]
  loc_00E09462: call [ecx+00000008h]
  loc_00E09465: mov eax, var_4
  loc_00E09468: mov ecx, var_14
  loc_00E0946B: pop edi
  loc_00E0946C: pop esi
  loc_00E0946D: mov fs:[00000000h], ecx
  loc_00E09474: pop ebx
  loc_00E09475: mov esp, ebp
  loc_00E09477: pop ebp
  loc_00E09478: retn 0004h
End Sub

Private Sub alamatortu_KeyPress(KeyAscii As Integer) 'E07640
  loc_00E07640: push ebp
  loc_00E07641: mov ebp, esp
  loc_00E07643: sub esp, 0000000Ch
  loc_00E07646: push 00402836h ; __vbaExceptHandler
  loc_00E0764B: mov eax, fs:[00000000h]
  loc_00E07651: push eax
  loc_00E07652: mov fs:[00000000h], esp
  loc_00E07659: sub esp, 00000014h
  loc_00E0765C: push ebx
  loc_00E0765D: push esi
  loc_00E0765E: push edi
  loc_00E0765F: mov var_C, esp
  loc_00E07662: mov var_8, 00402060h
  loc_00E07669: mov esi, Me
  loc_00E0766C: mov eax, esi
  loc_00E0766E: and eax, 00000001h
  loc_00E07671: mov var_4, eax
  loc_00E07674: and esi, FFFFFFFEh
  loc_00E07677: push esi
  loc_00E07678: mov Me, esi
  loc_00E0767B: mov ecx, [esi]
  loc_00E0767D: call [ecx+00000004h]
  loc_00E07680: mov edx, KeyAscii
  loc_00E07683: xor edi, edi
  loc_00E07685: mov var_18, edi
  loc_00E07688: cmp [edx], 000Dh
  loc_00E0768C: jnz 00E076CEh
  loc_00E0768E: mov eax, [esi]
  loc_00E07690: push esi
  loc_00E07691: call [eax+00000368h]
  loc_00E07697: lea ecx, var_18
  loc_00E0769A: push eax
  loc_00E0769B: push ecx
  loc_00E0769C: call [004010ACh] ; __vbaObjSet
  loc_00E076A2: mov esi, eax
  loc_00E076A4: push esi
  loc_00E076A5: mov edx, [esi]
  loc_00E076A7: call [edx+00000204h]
  loc_00E076AD: cmp eax, edi
  loc_00E076AF: fnclex
  loc_00E076B1: jge 00E076C5h
  loc_00E076B3: push 00000204h
  loc_00E076B8: push 006DCB70h
  loc_00E076BD: push esi
  loc_00E076BE: push eax
  loc_00E076BF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E076C5: lea ecx, var_18
  loc_00E076C8: call [00401254h] ; __vbaFreeObj
  loc_00E076CE: mov var_4, edi
  loc_00E076D1: push 00E076E3h
  loc_00E076D6: jmp 00E076E2h
  loc_00E076D8: lea ecx, var_18
  loc_00E076DB: call [00401254h] ; __vbaFreeObj
  loc_00E076E1: ret
  loc_00E076E2: ret
  loc_00E076E3: mov eax, Me
  loc_00E076E6: push eax
  loc_00E076E7: mov ecx, [eax]
  loc_00E076E9: call [ecx+00000008h]
  loc_00E076EC: mov eax, var_4
  loc_00E076EF: mov ecx, var_14
  loc_00E076F2: pop edi
  loc_00E076F3: pop esi
  loc_00E076F4: mov fs:[00000000h], ecx
  loc_00E076FB: pop ebx
  loc_00E076FC: mov esp, ebp
  loc_00E076FE: pop ebp
  loc_00E076FF: retn 0008h
End Sub

Private Sub tgl_UnknownEvent_9 'E0F440
  loc_00E0F440: push ebp
  loc_00E0F441: mov ebp, esp
  loc_00E0F443: sub esp, 0000000Ch
  loc_00E0F446: push 00402836h ; __vbaExceptHandler
  loc_00E0F44B: mov eax, fs:[00000000h]
  loc_00E0F451: push eax
  loc_00E0F452: mov fs:[00000000h], esp
  loc_00E0F459: sub esp, 00000058h
  loc_00E0F45C: push ebx
  loc_00E0F45D: push esi
  loc_00E0F45E: push edi
  loc_00E0F45F: mov var_C, esp
  loc_00E0F462: mov var_8, 00402270h
  loc_00E0F469: mov esi, Me
  loc_00E0F46C: mov eax, esi
  loc_00E0F46E: and eax, 00000001h
  loc_00E0F471: mov var_4, eax
  loc_00E0F474: and esi, FFFFFFFEh
  loc_00E0F477: push esi
  loc_00E0F478: mov Me, esi
  loc_00E0F47B: mov ecx, [esi]
  loc_00E0F47D: call [ecx+00000004h]
  loc_00E0F480: mov edx, [esi]
  loc_00E0F482: xor eax, eax
  loc_00E0F484: push esi
  loc_00E0F485: mov var_18, eax
  loc_00E0F488: mov var_1C, eax
  loc_00E0F48B: mov var_20, eax
  loc_00E0F48E: mov var_30, eax
  loc_00E0F491: call [edx+0000032Ch]
  loc_00E0F497: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E0F49D: push eax
  loc_00E0F49E: lea eax, var_1C
  loc_00E0F4A1: push eax
  loc_00E0F4A2: call edi
  loc_00E0F4A4: mov ebx, eax
  loc_00E0F4A6: push 0080FF80h
  loc_00E0F4AB: push ebx
  loc_00E0F4AC: mov ecx, [ebx]
  loc_00E0F4AE: call [ecx+0000006Ch]
  loc_00E0F4B1: test eax, eax
  loc_00E0F4B3: fnclex
  loc_00E0F4B5: jge 00E0F4C6h
  loc_00E0F4B7: push 0000006Ch
  loc_00E0F4B9: push 006E0410h
  loc_00E0F4BE: push ebx
  loc_00E0F4BF: push eax
  loc_00E0F4C0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F4C6: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E0F4CC: lea ecx, var_1C
  loc_00E0F4CF: call ebx
  loc_00E0F4D1: mov edx, [esi]
  loc_00E0F4D3: push esi
  loc_00E0F4D4: call [edx+0000032Ch]
  loc_00E0F4DA: push eax
  loc_00E0F4DB: lea eax, var_1C
  loc_00E0F4DE: push eax
  loc_00E0F4DF: call edi
  loc_00E0F4E1: mov ecx, [eax]
  loc_00E0F4E3: push 006E0424h ; "Silahkan Input Data Calon Peserta Didik Secara Lengkap !"
  loc_00E0F4E8: push eax
  loc_00E0F4E9: mov var_54, eax
  loc_00E0F4EC: call [ecx+00000054h]
  loc_00E0F4EF: test eax, eax
  loc_00E0F4F1: fnclex
  loc_00E0F4F3: jge 00E0F507h
  loc_00E0F4F5: mov edx, var_54
  loc_00E0F4F8: push 00000054h
  loc_00E0F4FA: push 006E0410h
  loc_00E0F4FF: push edx
  loc_00E0F500: push eax
  loc_00E0F501: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F507: lea ecx, var_1C
  loc_00E0F50A: call ebx
  loc_00E0F50C: mov eax, [esi]
  loc_00E0F50E: push esi
  loc_00E0F50F: call [eax+000003BCh]
  loc_00E0F515: lea ecx, var_1C
  loc_00E0F518: push eax
  loc_00E0F519: push ecx
  loc_00E0F51A: call edi
  loc_00E0F51C: mov edx, [eax]
  loc_00E0F51E: push 00000000h
  loc_00E0F520: push eax
  loc_00E0F521: mov var_54, eax
  loc_00E0F524: call [edx+0000008Ch]
  loc_00E0F52A: test eax, eax
  loc_00E0F52C: fnclex
  loc_00E0F52E: jge 00E0F545h
  loc_00E0F530: mov ecx, var_54
  loc_00E0F533: push 0000008Ch
  loc_00E0F538: push 006DCDA0h
  loc_00E0F53D: push ecx
  loc_00E0F53E: push eax
  loc_00E0F53F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F545: lea ecx, var_1C
  loc_00E0F548: call ebx
  loc_00E0F54A: mov edx, [esi]
  loc_00E0F54C: push esi
  loc_00E0F54D: call [edx+00000330h]
  loc_00E0F553: push eax
  loc_00E0F554: lea eax, var_1C
  loc_00E0F557: push eax
  loc_00E0F558: call edi
  loc_00E0F55A: mov ecx, [eax]
  loc_00E0F55C: push FFFFFFFFh
  loc_00E0F55E: push eax
  loc_00E0F55F: mov var_54, eax
  loc_00E0F562: call [ecx+0000005Ch]
  loc_00E0F565: test eax, eax
  loc_00E0F567: fnclex
  loc_00E0F569: jge 00E0F57Dh
  loc_00E0F56B: mov edx, var_54
  loc_00E0F56E: push 0000005Ch
  loc_00E0F570: push 006DCAE0h
  loc_00E0F575: push edx
  loc_00E0F576: push eax
  loc_00E0F577: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F57D: lea ecx, var_1C
  loc_00E0F580: call ebx
  loc_00E0F582: sub esp, 00000010h
  loc_00E0F585: mov ecx, 0000000Bh
  loc_00E0F58A: mov edx, esp
  loc_00E0F58C: xor eax, eax
  loc_00E0F58E: push 80010007h
  loc_00E0F593: push esi
  loc_00E0F594: mov [edx], ecx
  loc_00E0F596: mov ecx, var_3C
  loc_00E0F599: mov [edx+00000004h], ecx
  loc_00E0F59C: mov ecx, [esi]
  loc_00E0F59E: mov [edx+00000008h], eax
  loc_00E0F5A1: mov eax, var_34
  loc_00E0F5A4: mov [edx+0000000Ch], eax
  loc_00E0F5A7: call [ecx+00000300h]
  loc_00E0F5AD: lea edx, var_1C
  loc_00E0F5B0: push eax
  loc_00E0F5B1: push edx
  loc_00E0F5B2: call edi
  loc_00E0F5B4: push eax
  loc_00E0F5B5: call [00401238h] ; __vbaLateIdSt
  loc_00E0F5BB: lea ecx, var_1C
  loc_00E0F5BE: call ebx
  loc_00E0F5C0: mov eax, [esi]
  loc_00E0F5C2: push esi
  loc_00E0F5C3: call [eax+00000304h]
  loc_00E0F5C9: lea ecx, var_1C
  loc_00E0F5CC: push eax
  loc_00E0F5CD: push ecx
  loc_00E0F5CE: call edi
  loc_00E0F5D0: lea edx, var_30
  loc_00E0F5D3: mov ebx, eax
  loc_00E0F5D5: push edx
  loc_00E0F5D6: mov var_54, ebx
  loc_00E0F5D9: call [004011D8h] ; rtcGetDateVar
  loc_00E0F5DF: mov ebx, [ebx]
  loc_00E0F5E1: lea eax, var_30
  loc_00E0F5E4: lea ecx, var_18
  loc_00E0F5E7: push eax
  loc_00E0F5E8: push ecx
  loc_00E0F5E9: call [00401180h] ; __vbaStrVarVal
  loc_00E0F5EF: mov edx, ebx
  loc_00E0F5F1: mov ebx, var_54
  loc_00E0F5F4: push eax
  loc_00E0F5F5: push ebx
  loc_00E0F5F6: call [edx+000000A4h]
  loc_00E0F5FC: test eax, eax
  loc_00E0F5FE: fnclex
  loc_00E0F600: jge 00E0F614h
  loc_00E0F602: push 000000A4h
  loc_00E0F607: push 006DCB70h
  loc_00E0F60C: push ebx
  loc_00E0F60D: push eax
  loc_00E0F60E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F614: lea ecx, var_18
  loc_00E0F617: call [00401258h] ; __vbaFreeStr
  loc_00E0F61D: lea ecx, var_1C
  loc_00E0F620: call [00401254h] ; __vbaFreeObj
  loc_00E0F626: lea ecx, var_30
  loc_00E0F629: call [00401024h] ; __vbaFreeVar
  loc_00E0F62F: mov eax, [esi]
  loc_00E0F631: push esi
  loc_00E0F632: call [eax+0000039Ch]
  loc_00E0F638: lea ecx, var_1C
  loc_00E0F63B: push eax
  loc_00E0F63C: push ecx
  loc_00E0F63D: call edi
  loc_00E0F63F: mov ebx, eax
  loc_00E0F641: push ebx
  loc_00E0F642: mov edx, [ebx]
  loc_00E0F644: call [edx+00000204h]
  loc_00E0F64A: test eax, eax
  loc_00E0F64C: fnclex
  loc_00E0F64E: jge 00E0F662h
  loc_00E0F650: push 00000204h
  loc_00E0F655: push 006DCB70h
  loc_00E0F65A: push ebx
  loc_00E0F65B: push eax
  loc_00E0F65C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F662: lea ecx, var_1C
  loc_00E0F665: call [00401254h] ; __vbaFreeObj
  loc_00E0F66B: mov eax, [00E3D0D8h]
  loc_00E0F670: test eax, eax
  loc_00E0F672: jnz 00E0F684h
  loc_00E0F674: push 00E3D0D8h
  loc_00E0F679: push 006D300Ch
  loc_00E0F67E: call [004011A0h] ; __vbaNew2
  loc_00E0F684: mov ebx, [00E3D0D8h]
  loc_00E0F68A: push ebx
  loc_00E0F68B: mov eax, [ebx]
  loc_00E0F68D: call [eax+000002B4h]
  loc_00E0F693: test eax, eax
  loc_00E0F695: fnclex
  loc_00E0F697: jge 00E0F6ABh
  loc_00E0F699: push 000002B4h
  loc_00E0F69E: push 006E0014h
  loc_00E0F6A3: push ebx
  loc_00E0F6A4: push eax
  loc_00E0F6A5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F6AB: mov eax, [00E3D0D8h]
  loc_00E0F6B0: test eax, eax
  loc_00E0F6B2: jnz 00E0F6C9h
  loc_00E0F6B4: push 00E3D0D8h
  loc_00E0F6B9: push 006D300Ch
  loc_00E0F6BE: call [004011A0h] ; __vbaNew2
  loc_00E0F6C4: mov eax, [00E3D0D8h]
  loc_00E0F6C9: mov ecx, [eax]
  loc_00E0F6CB: push eax
  loc_00E0F6CC: call [ecx+00000304h]
  loc_00E0F6D2: lea edx, var_20
  loc_00E0F6D5: push eax
  loc_00E0F6D6: push edx
  loc_00E0F6D7: call edi
  loc_00E0F6D9: mov ebx, eax
  loc_00E0F6DB: mov eax, [esi]
  loc_00E0F6DD: push esi
  loc_00E0F6DE: call [eax+00000304h]
  loc_00E0F6E4: lea ecx, var_1C
  loc_00E0F6E7: push eax
  loc_00E0F6E8: push ecx
  loc_00E0F6E9: call edi
  loc_00E0F6EB: mov esi, eax
  loc_00E0F6ED: lea eax, var_18
  loc_00E0F6F0: push eax
  loc_00E0F6F1: push esi
  loc_00E0F6F2: mov edx, [esi]
  loc_00E0F6F4: call [edx+000000A0h]
  loc_00E0F6FA: test eax, eax
  loc_00E0F6FC: fnclex
  loc_00E0F6FE: jge 00E0F712h
  loc_00E0F700: push 000000A0h
  loc_00E0F705: push 006DCB70h
  loc_00E0F70A: push esi
  loc_00E0F70B: push eax
  loc_00E0F70C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F712: mov edx, var_18
  loc_00E0F715: mov ecx, [ebx]
  loc_00E0F717: push edx
  loc_00E0F718: push ebx
  loc_00E0F719: call [ecx+000000A4h]
  loc_00E0F71F: test eax, eax
  loc_00E0F721: fnclex
  loc_00E0F723: jge 00E0F737h
  loc_00E0F725: push 000000A4h
  loc_00E0F72A: push 006DCB70h
  loc_00E0F72F: push ebx
  loc_00E0F730: push eax
  loc_00E0F731: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F737: lea ecx, var_18
  loc_00E0F73A: call [00401258h] ; __vbaFreeStr
  loc_00E0F740: lea eax, var_20
  loc_00E0F743: lea ecx, var_1C
  loc_00E0F746: push eax
  loc_00E0F747: push ecx
  loc_00E0F748: push 00000002h
  loc_00E0F74A: call [00401048h] ; __vbaFreeObjList
  loc_00E0F750: add esp, 0000000Ch
  loc_00E0F753: mov var_4, 00000000h
  loc_00E0F75A: push 00E0F788h
  loc_00E0F75F: jmp 00E0F787h
  loc_00E0F761: lea ecx, var_18
  loc_00E0F764: call [00401258h] ; __vbaFreeStr
  loc_00E0F76A: lea edx, var_20
  loc_00E0F76D: lea eax, var_1C
  loc_00E0F770: push edx
  loc_00E0F771: push eax
  loc_00E0F772: push 00000002h
  loc_00E0F774: call [00401048h] ; __vbaFreeObjList
  loc_00E0F77A: add esp, 0000000Ch
  loc_00E0F77D: lea ecx, var_30
  loc_00E0F780: call [00401024h] ; __vbaFreeVar
  loc_00E0F786: ret
  loc_00E0F787: ret
  loc_00E0F788: mov eax, Me
  loc_00E0F78B: push eax
  loc_00E0F78C: mov ecx, [eax]
  loc_00E0F78E: call [ecx+00000008h]
  loc_00E0F791: mov eax, var_4
  loc_00E0F794: mov ecx, var_14
  loc_00E0F797: pop edi
  loc_00E0F798: pop esi
  loc_00E0F799: mov fs:[00000000h], ecx
  loc_00E0F7A0: pop ebx
  loc_00E0F7A1: mov esp, ebp
  loc_00E0F7A3: pop ebp
  loc_00E0F7A4: retn 0004h
End Sub

Private Sub alamatwali_KeyPress(KeyAscii As Integer) 'E09950
  loc_00E09950: push ebp
  loc_00E09951: mov ebp, esp
  loc_00E09953: sub esp, 0000000Ch
  loc_00E09956: push 00402836h ; __vbaExceptHandler
  loc_00E0995B: mov eax, fs:[00000000h]
  loc_00E09961: push eax
  loc_00E09962: mov fs:[00000000h], esp
  loc_00E09969: sub esp, 00000014h
  loc_00E0996C: push ebx
  loc_00E0996D: push esi
  loc_00E0996E: push edi
  loc_00E0996F: mov var_C, esp
  loc_00E09972: mov var_8, 00402120h
  loc_00E09979: mov esi, Me
  loc_00E0997C: mov eax, esi
  loc_00E0997E: and eax, 00000001h
  loc_00E09981: mov var_4, eax
  loc_00E09984: and esi, FFFFFFFEh
  loc_00E09987: push esi
  loc_00E09988: mov Me, esi
  loc_00E0998B: mov ecx, [esi]
  loc_00E0998D: call [ecx+00000004h]
  loc_00E09990: mov edx, KeyAscii
  loc_00E09993: xor edi, edi
  loc_00E09995: mov var_18, edi
  loc_00E09998: cmp [edx], 000Dh
  loc_00E0999C: jnz 00E099DEh
  loc_00E0999E: mov eax, [esi]
  loc_00E099A0: push esi
  loc_00E099A1: call [eax+00000344h]
  loc_00E099A7: lea ecx, var_18
  loc_00E099AA: push eax
  loc_00E099AB: push ecx
  loc_00E099AC: call [004010ACh] ; __vbaObjSet
  loc_00E099B2: mov esi, eax
  loc_00E099B4: push esi
  loc_00E099B5: mov edx, [esi]
  loc_00E099B7: call [edx+00000204h]
  loc_00E099BD: cmp eax, edi
  loc_00E099BF: fnclex
  loc_00E099C1: jge 00E099D5h
  loc_00E099C3: push 00000204h
  loc_00E099C8: push 006DCB70h
  loc_00E099CD: push esi
  loc_00E099CE: push eax
  loc_00E099CF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E099D5: lea ecx, var_18
  loc_00E099D8: call [00401254h] ; __vbaFreeObj
  loc_00E099DE: mov var_4, edi
  loc_00E099E1: push 00E099F3h
  loc_00E099E6: jmp 00E099F2h
  loc_00E099E8: lea ecx, var_18
  loc_00E099EB: call [00401254h] ; __vbaFreeObj
  loc_00E099F1: ret
  loc_00E099F2: ret
  loc_00E099F3: mov eax, Me
  loc_00E099F6: push eax
  loc_00E099F7: mov ecx, [eax]
  loc_00E099F9: call [ecx+00000008h]
  loc_00E099FC: mov eax, var_4
  loc_00E099FF: mov ecx, var_14
  loc_00E09A02: pop edi
  loc_00E09A03: pop esi
  loc_00E09A04: mov fs:[00000000h], ecx
  loc_00E09A0B: pop ebx
  loc_00E09A0C: mov esp, ebp
  loc_00E09A0E: pop ebp
  loc_00E09A0F: retn 0008h
End Sub

Private Sub optlaki_Click() 'E0A8F0
  loc_00E0A8F0: push ebp
  loc_00E0A8F1: mov ebp, esp
  loc_00E0A8F3: sub esp, 0000000Ch
  loc_00E0A8F6: push 00402836h ; __vbaExceptHandler
  loc_00E0A8FB: mov eax, fs:[00000000h]
  loc_00E0A901: push eax
  loc_00E0A902: mov fs:[00000000h], esp
  loc_00E0A909: sub esp, 00000034h
  loc_00E0A90C: push ebx
  loc_00E0A90D: push esi
  loc_00E0A90E: push edi
  loc_00E0A90F: mov var_C, esp
  loc_00E0A912: mov var_8, 00402210h
  loc_00E0A919: mov esi, Me
  loc_00E0A91C: mov eax, esi
  loc_00E0A91E: and eax, 00000001h
  loc_00E0A921: mov var_4, eax
  loc_00E0A924: and esi, FFFFFFFEh
  loc_00E0A927: push esi
  loc_00E0A928: mov Me, esi
  loc_00E0A92B: mov ecx, [esi]
  loc_00E0A92D: call [ecx+00000004h]
  loc_00E0A930: mov edx, [esi]
  loc_00E0A932: push esi
  loc_00E0A933: mov var_18, 00000000h
  loc_00E0A93A: call [edx+0000037Ch]
  loc_00E0A940: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0A946: push eax
  loc_00E0A947: lea eax, var_18
  loc_00E0A94A: push eax
  loc_00E0A94B: call ebx
  loc_00E0A94D: mov edi, eax
  loc_00E0A94F: push 00000000h
  loc_00E0A951: push edi
  loc_00E0A952: mov ecx, [edi]
  loc_00E0A954: call [ecx+0000009Ch]
  loc_00E0A95A: test eax, eax
  loc_00E0A95C: fnclex
  loc_00E0A95E: jge 00E0A972h
  loc_00E0A960: push 0000009Ch
  loc_00E0A965: push 006E03D4h
  loc_00E0A96A: push edi
  loc_00E0A96B: push eax
  loc_00E0A96C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A972: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E0A978: lea ecx, var_18
  loc_00E0A97B: call edi
  loc_00E0A97D: sub esp, 00000010h
  loc_00E0A980: mov ecx, 0000000Bh
  loc_00E0A985: mov edx, esp
  loc_00E0A987: or eax, FFFFFFFFh
  loc_00E0A98A: push 80010007h
  loc_00E0A98F: push esi
  loc_00E0A990: mov [edx], ecx
  loc_00E0A992: mov ecx, var_24
  loc_00E0A995: mov [edx+00000004h], ecx
  loc_00E0A998: mov ecx, [esi]
  loc_00E0A99A: mov [edx+00000008h], eax
  loc_00E0A99D: mov eax, var_1C
  loc_00E0A9A0: mov [edx+0000000Ch], eax
  loc_00E0A9A3: call [ecx+000003B8h]
  loc_00E0A9A9: lea edx, var_18
  loc_00E0A9AC: push eax
  loc_00E0A9AD: push edx
  loc_00E0A9AE: call ebx
  loc_00E0A9B0: push eax
  loc_00E0A9B1: call [00401238h] ; __vbaLateIdSt
  loc_00E0A9B7: lea ecx, var_18
  loc_00E0A9BA: call edi
  loc_00E0A9BC: mov var_4, 00000000h
  loc_00E0A9C3: push 00E0A9D5h
  loc_00E0A9C8: jmp 00E0A9D4h
  loc_00E0A9CA: lea ecx, var_18
  loc_00E0A9CD: call [00401254h] ; __vbaFreeObj
  loc_00E0A9D3: ret
  loc_00E0A9D4: ret
  loc_00E0A9D5: mov eax, Me
  loc_00E0A9D8: push eax
  loc_00E0A9D9: mov ecx, [eax]
  loc_00E0A9DB: call [ecx+00000008h]
  loc_00E0A9DE: mov eax, var_4
  loc_00E0A9E1: mov ecx, var_14
  loc_00E0A9E4: pop edi
  loc_00E0A9E5: pop esi
  loc_00E0A9E6: mov fs:[00000000h], ecx
  loc_00E0A9ED: pop ebx
  loc_00E0A9EE: mov esp, ebp
  loc_00E0A9F0: pop ebp
  loc_00E0A9F1: retn 0004h
End Sub

Private Sub back_Click() 'E07710
  loc_00E07710: push ebp
  loc_00E07711: mov ebp, esp
  loc_00E07713: sub esp, 0000000Ch
  loc_00E07716: push 00402836h ; __vbaExceptHandler
  loc_00E0771B: mov eax, fs:[00000000h]
  loc_00E07721: push eax
  loc_00E07722: mov fs:[00000000h], esp
  loc_00E07729: sub esp, 00000038h
  loc_00E0772C: push ebx
  loc_00E0772D: push esi
  loc_00E0772E: push edi
  loc_00E0772F: mov var_C, esp
  loc_00E07732: mov var_8, 00402070h
  loc_00E07739: mov eax, Me
  loc_00E0773C: mov ecx, eax
  loc_00E0773E: and ecx, 00000001h
  loc_00E07741: mov var_4, ecx
  loc_00E07744: and al, FEh
  loc_00E07746: push eax
  loc_00E07747: mov Me, eax
  loc_00E0774A: mov edx, [eax]
  loc_00E0774C: call [edx+00000004h]
  loc_00E0774F: mov eax, [00E3D024h]
  loc_00E07754: or edi, FFFFFFFFh
  loc_00E07757: test eax, eax
  loc_00E07759: mov var_18, 00000000h
  loc_00E07760: mov esi, 0000000Bh
  loc_00E07765: jnz 00E0777Ch
  loc_00E07767: push 00E3D024h
  loc_00E0776C: push 006CE120h
  loc_00E07771: call [004011A0h] ; __vbaNew2
  loc_00E07777: mov eax, [00E3D024h]
  loc_00E0777C: sub esp, 00000010h
  loc_00E0777F: mov edx, [eax]
  loc_00E07781: mov ecx, esp
  loc_00E07783: push 80010007h
  loc_00E07788: push eax
  loc_00E07789: mov [ecx], esi
  loc_00E0778B: mov esi, var_24
  loc_00E0778E: mov [ecx+00000004h], esi
  loc_00E07791: mov [ecx+00000008h], edi
  loc_00E07794: mov edi, var_1C
  loc_00E07797: mov [ecx+0000000Ch], edi
  loc_00E0779A: call [edx+00000318h]
  loc_00E077A0: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E077A6: push eax
  loc_00E077A7: lea eax, var_18
  loc_00E077AA: push eax
  loc_00E077AB: call ebx
  loc_00E077AD: push eax
  loc_00E077AE: call [00401238h] ; __vbaLateIdSt
  loc_00E077B4: lea ecx, var_18
  loc_00E077B7: call [00401254h] ; __vbaFreeObj
  loc_00E077BD: mov eax, [00E3D024h]
  loc_00E077C2: test eax, eax
  loc_00E077C4: jnz 00E077DBh
  loc_00E077C6: push 00E3D024h
  loc_00E077CB: push 006CE120h
  loc_00E077D0: call [004011A0h] ; __vbaNew2
  loc_00E077D6: mov eax, [00E3D024h]
  loc_00E077DB: sub esp, 00000010h
  loc_00E077DE: mov ecx, 0000000Bh
  loc_00E077E3: mov edx, esp
  loc_00E077E5: push 80010007h
  loc_00E077EA: push eax
  loc_00E077EB: mov [edx], ecx
  loc_00E077ED: or ecx, FFFFFFFFh
  loc_00E077F0: mov [edx+00000004h], esi
  loc_00E077F3: mov [edx+00000008h], ecx
  loc_00E077F6: mov ecx, [eax]
  loc_00E077F8: mov [edx+0000000Ch], edi
  loc_00E077FB: call [ecx+0000031Ch]
  loc_00E07801: lea edx, var_18
  loc_00E07804: push eax
  loc_00E07805: push edx
  loc_00E07806: call ebx
  loc_00E07808: push eax
  loc_00E07809: call [00401238h] ; __vbaLateIdSt
  loc_00E0780F: lea ecx, var_18
  loc_00E07812: call [00401254h] ; __vbaFreeObj
  loc_00E07818: mov eax, [00E3D024h]
  loc_00E0781D: mov var_20, FFFFFFFFh
  loc_00E07824: test eax, eax
  loc_00E07826: jnz 00E0783Dh
  loc_00E07828: push 00E3D024h
  loc_00E0782D: push 006CE120h
  loc_00E07832: call [004011A0h] ; __vbaNew2
  loc_00E07838: mov eax, [00E3D024h]
  loc_00E0783D: sub esp, 00000010h
  loc_00E07840: mov ecx, 0000000Bh
  loc_00E07845: mov edx, esp
  loc_00E07847: push 80010007h
  loc_00E0784C: push eax
  loc_00E0784D: mov [edx], ecx
  loc_00E0784F: mov ecx, var_20
  loc_00E07852: mov [edx+00000004h], esi
  loc_00E07855: mov [edx+00000008h], ecx
  loc_00E07858: mov [edx+0000000Ch], edi
  loc_00E0785B: mov edx, [eax]
  loc_00E0785D: call [edx+00000320h]
  loc_00E07863: push eax
  loc_00E07864: lea eax, var_18
  loc_00E07867: push eax
  loc_00E07868: call ebx
  loc_00E0786A: push eax
  loc_00E0786B: call [00401238h] ; __vbaLateIdSt
  loc_00E07871: lea ecx, var_18
  loc_00E07874: call [00401254h] ; __vbaFreeObj
  loc_00E0787A: mov eax, [00E3D024h]
  loc_00E0787F: test eax, eax
  loc_00E07881: jnz 00E07898h
  loc_00E07883: push 00E3D024h
  loc_00E07888: push 006CE120h
  loc_00E0788D: call [004011A0h] ; __vbaNew2
  loc_00E07893: mov eax, [00E3D024h]
  loc_00E07898: sub esp, 00000010h
  loc_00E0789B: mov ecx, 00000008h
  loc_00E078A0: mov edx, esp
  loc_00E078A2: push FFFFFDFAh
  loc_00E078A7: push eax
  loc_00E078A8: mov [edx], ecx
  loc_00E078AA: mov ecx, 006DCDB4h ; " A D M I N "
  loc_00E078AF: mov [edx+00000004h], esi
  loc_00E078B2: mov [edx+00000008h], ecx
  loc_00E078B5: mov ecx, [eax]
  loc_00E078B7: mov [edx+0000000Ch], edi
  loc_00E078BA: call [ecx+00000324h]
  loc_00E078C0: lea edx, var_18
  loc_00E078C3: push eax
  loc_00E078C4: push edx
  loc_00E078C5: call ebx
  loc_00E078C7: push eax
  loc_00E078C8: call [00401238h] ; __vbaLateIdSt
  loc_00E078CE: lea ecx, var_18
  loc_00E078D1: call [00401254h] ; __vbaFreeObj
  loc_00E078D7: mov eax, [00E3D024h]
  loc_00E078DC: mov var_28, 0000000Bh
  loc_00E078E3: test eax, eax
  loc_00E078E5: jnz 00E078FCh
  loc_00E078E7: push 00E3D024h
  loc_00E078EC: push 006CE120h
  loc_00E078F1: call [004011A0h] ; __vbaNew2
  loc_00E078F7: mov eax, [00E3D024h]
  loc_00E078FC: mov ecx, var_28
  loc_00E078FF: sub esp, 00000010h
  loc_00E07902: mov edx, esp
  loc_00E07904: push 8001000Dh
  loc_00E07909: push eax
  loc_00E0790A: mov [edx], ecx
  loc_00E0790C: xor ecx, ecx
  loc_00E0790E: mov [edx+00000004h], esi
  loc_00E07911: mov [edx+00000008h], ecx
  loc_00E07914: mov [edx+0000000Ch], edi
  loc_00E07917: mov edx, [eax]
  loc_00E07919: call [edx+00000324h]
  loc_00E0791F: push eax
  loc_00E07920: lea eax, var_18
  loc_00E07923: push eax
  loc_00E07924: call ebx
  loc_00E07926: push eax
  loc_00E07927: call [00401238h] ; __vbaLateIdSt
  loc_00E0792D: lea ecx, var_18
  loc_00E07930: call [00401254h] ; __vbaFreeObj
  loc_00E07936: mov eax, [00E3D024h]
  loc_00E0793B: test eax, eax
  loc_00E0793D: jnz 00E07954h
  loc_00E0793F: push 00E3D024h
  loc_00E07944: push 006CE120h
  loc_00E07949: call [004011A0h] ; __vbaNew2
  loc_00E0794F: mov eax, [00E3D024h]
  loc_00E07954: sub esp, 00000010h
  loc_00E07957: mov ecx, 00000003h
  loc_00E0795C: mov edx, esp
  loc_00E0795E: push FFFFFE0Bh
  loc_00E07963: push eax
  loc_00E07964: mov [edx], ecx
  loc_00E07966: mov ecx, 00404040h
  loc_00E0796B: mov [edx+00000004h], esi
  loc_00E0796E: mov [edx+00000008h], ecx
  loc_00E07971: mov ecx, [eax]
  loc_00E07973: mov [edx+0000000Ch], edi
  loc_00E07976: call [ecx+00000324h]
  loc_00E0797C: lea edx, var_18
  loc_00E0797F: push eax
  loc_00E07980: push edx
  loc_00E07981: call ebx
  loc_00E07983: push eax
  loc_00E07984: call [00401238h] ; __vbaLateIdSt
  loc_00E0798A: lea ecx, var_18
  loc_00E0798D: call [00401254h] ; __vbaFreeObj
  loc_00E07993: mov eax, [00E3D024h]
  loc_00E07998: mov var_28, 00000003h
  loc_00E0799F: test eax, eax
  loc_00E079A1: jnz 00E079B8h
  loc_00E079A3: push 00E3D024h
  loc_00E079A8: push 006CE120h
  loc_00E079AD: call [004011A0h] ; __vbaNew2
  loc_00E079B3: mov eax, [00E3D024h]
  loc_00E079B8: mov ecx, var_28
  loc_00E079BB: sub esp, 00000010h
  loc_00E079BE: mov edx, esp
  loc_00E079C0: push FFFFFDFFh
  loc_00E079C5: push eax
  loc_00E079C6: mov [edx], ecx
  loc_00E079C8: mov ecx, 00E0E0E0h
  loc_00E079CD: mov [edx+00000004h], esi
  loc_00E079D0: mov [edx+00000008h], ecx
  loc_00E079D3: mov [edx+0000000Ch], edi
  loc_00E079D6: mov edx, [eax]
  loc_00E079D8: call [edx+00000324h]
  loc_00E079DE: push eax
  loc_00E079DF: lea eax, var_18
  loc_00E079E2: push eax
  loc_00E079E3: call ebx
  loc_00E079E5: push eax
  loc_00E079E6: call [00401238h] ; __vbaLateIdSt
  loc_00E079EC: lea ecx, var_18
  loc_00E079EF: call [00401254h] ; __vbaFreeObj
  loc_00E079F5: mov eax, [00E3D9CCh]
  loc_00E079FA: test eax, eax
  loc_00E079FC: jnz 00E07A0Eh
  loc_00E079FE: push 00E3D9CCh
  loc_00E07A03: push 006DC960h
  loc_00E07A08: call [004011A0h] ; __vbaNew2
  loc_00E07A0E: mov ecx, Me
  loc_00E07A11: mov eax, [00E3D9CCh]
  loc_00E07A16: lea edx, var_18
  loc_00E07A19: push ecx
  loc_00E07A1A: mov ebx, [eax]
  loc_00E07A1C: push edx
  loc_00E07A1D: mov var_3C, eax
  loc_00E07A20: call [004010B4h] ; __vbaObjSetAddref
  loc_00E07A26: mov var_4C, ebx
  loc_00E07A29: mov ebx, var_3C
  loc_00E07A2C: push eax
  loc_00E07A2D: mov eax, var_4C
  loc_00E07A30: push ebx
  loc_00E07A31: call [eax+00000010h]
  loc_00E07A34: test eax, eax
  loc_00E07A36: fnclex
  loc_00E07A38: jge 00E07A49h
  loc_00E07A3A: push 00000010h
  loc_00E07A3C: push 006DC950h
  loc_00E07A41: push ebx
  loc_00E07A42: push eax
  loc_00E07A43: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07A49: lea ecx, var_18
  loc_00E07A4C: call [00401254h] ; __vbaFreeObj
  loc_00E07A52: mov eax, [00E3D024h]
  loc_00E07A57: test eax, eax
  loc_00E07A59: jnz 00E07A6Bh
  loc_00E07A5B: push 00E3D024h
  loc_00E07A60: push 006CE120h
  loc_00E07A65: call [004011A0h] ; __vbaNew2
  loc_00E07A6B: sub esp, 00000010h
  loc_00E07A6E: mov eax, 0000000Ah
  loc_00E07A73: mov edx, esp
  loc_00E07A75: mov var_28, eax
  loc_00E07A78: sub esp, 00000010h
  loc_00E07A7B: mov ebx, [00E3D024h]
  loc_00E07A81: mov [edx], eax
  loc_00E07A83: mov eax, var_34
  loc_00E07A86: mov ecx, [ebx]
  loc_00E07A88: mov var_20, 80020004h
  loc_00E07A8F: mov [edx+00000004h], eax
  loc_00E07A92: mov eax, 80020004h
  loc_00E07A97: mov [edx+00000008h], eax
  loc_00E07A9A: mov eax, var_2C
  loc_00E07A9D: mov [edx+0000000Ch], eax
  loc_00E07AA0: mov eax, var_28
  loc_00E07AA3: mov edx, esp
  loc_00E07AA5: push ebx
  loc_00E07AA6: mov [edx], eax
  loc_00E07AA8: mov eax, var_20
  loc_00E07AAB: mov [edx+00000004h], esi
  loc_00E07AAE: mov [edx+00000008h], eax
  loc_00E07AB1: mov [edx+0000000Ch], edi
  loc_00E07AB4: call [ecx+000002B0h]
  loc_00E07ABA: test eax, eax
  loc_00E07ABC: fnclex
  loc_00E07ABE: jge 00E07AD2h
  loc_00E07AC0: push 000002B0h
  loc_00E07AC5: push 006DC970h
  loc_00E07ACA: push ebx
  loc_00E07ACB: push eax
  loc_00E07ACC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07AD2: mov eax, Me
  loc_00E07AD5: push eax
  loc_00E07AD6: mov ecx, [eax]
  loc_00E07AD8: call [ecx+000003C8h]
  loc_00E07ADE: lea edx, var_18
  loc_00E07AE1: push eax
  loc_00E07AE2: push edx
  loc_00E07AE3: call [004010ACh] ; __vbaObjSet
  loc_00E07AE9: mov esi, eax
  loc_00E07AEB: push 00000000h
  loc_00E07AED: push esi
  loc_00E07AEE: mov eax, [esi]
  loc_00E07AF0: call [eax+0000008Ch]
  loc_00E07AF6: test eax, eax
  loc_00E07AF8: fnclex
  loc_00E07AFA: jge 00E07B0Eh
  loc_00E07AFC: push 0000008Ch
  loc_00E07B01: push 006DCDA0h
  loc_00E07B06: push esi
  loc_00E07B07: push eax
  loc_00E07B08: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07B0E: lea ecx, var_18
  loc_00E07B11: call [00401254h] ; __vbaFreeObj
  loc_00E07B17: mov var_4, 00000000h
  loc_00E07B1E: push 00E07B30h
  loc_00E07B23: jmp 00E07B2Fh
  loc_00E07B25: lea ecx, var_18
  loc_00E07B28: call [00401254h] ; __vbaFreeObj
  loc_00E07B2E: ret
  loc_00E07B2F: ret
  loc_00E07B30: mov eax, Me
  loc_00E07B33: push eax
  loc_00E07B34: mov ecx, [eax]
  loc_00E07B36: call [ecx+00000008h]
  loc_00E07B39: mov eax, var_4
  loc_00E07B3C: mov ecx, var_14
  loc_00E07B3F: pop edi
  loc_00E07B40: pop esi
  loc_00E07B41: mov fs:[00000000h], ecx
  loc_00E07B48: pop ebx
  loc_00E07B49: mov esp, ebp
  loc_00E07B4B: pop ebp
  loc_00E07B4C: retn 0004h
End Sub

Private Sub fcari_Click() 'E086B0
  loc_00E086B0: push ebp
  loc_00E086B1: mov ebp, esp
  loc_00E086B3: sub esp, 0000000Ch
  loc_00E086B6: push 00402836h ; __vbaExceptHandler
  loc_00E086BB: mov eax, fs:[00000000h]
  loc_00E086C1: push eax
  loc_00E086C2: mov fs:[00000000h], esp
  loc_00E086C9: sub esp, 00000034h
  loc_00E086CC: push ebx
  loc_00E086CD: push esi
  loc_00E086CE: push edi
  loc_00E086CF: mov var_C, esp
  loc_00E086D2: mov var_8, 004020B0h
  loc_00E086D9: mov esi, Me
  loc_00E086DC: mov eax, esi
  loc_00E086DE: and eax, 00000001h
  loc_00E086E1: mov var_4, eax
  loc_00E086E4: and esi, FFFFFFFEh
  loc_00E086E7: push esi
  loc_00E086E8: mov Me, esi
  loc_00E086EB: mov ecx, [esi]
  loc_00E086ED: call [ecx+00000004h]
  loc_00E086F0: mov edx, [esi]
  loc_00E086F2: push esi
  loc_00E086F3: mov var_18, 00000000h
  loc_00E086FA: call [edx+000003E0h]
  loc_00E08700: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E08706: push eax
  loc_00E08707: lea eax, var_18
  loc_00E0870A: push eax
  loc_00E0870B: call ebx
  loc_00E0870D: mov edi, eax
  loc_00E0870F: push 00000000h
  loc_00E08711: push edi
  loc_00E08712: mov ecx, [edi]
  loc_00E08714: call [ecx+0000008Ch]
  loc_00E0871A: test eax, eax
  loc_00E0871C: fnclex
  loc_00E0871E: jge 00E08732h
  loc_00E08720: push 0000008Ch
  loc_00E08725: push 006DCDA0h
  loc_00E0872A: push edi
  loc_00E0872B: push eax
  loc_00E0872C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08732: lea ecx, var_18
  loc_00E08735: call [00401254h] ; __vbaFreeObj
  loc_00E0873B: mov edx, [esi]
  loc_00E0873D: push esi
  loc_00E0873E: call [edx+000003DCh]
  loc_00E08744: push eax
  loc_00E08745: lea eax, var_18
  loc_00E08748: push eax
  loc_00E08749: call ebx
  loc_00E0874B: mov edi, eax
  loc_00E0874D: push 00000000h
  loc_00E0874F: push edi
  loc_00E08750: mov ecx, [edi]
  loc_00E08752: call [ecx+0000008Ch]
  loc_00E08758: test eax, eax
  loc_00E0875A: fnclex
  loc_00E0875C: jge 00E08770h
  loc_00E0875E: push 0000008Ch
  loc_00E08763: push 006DCDA0h
  loc_00E08768: push edi
  loc_00E08769: push eax
  loc_00E0876A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08770: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E08776: lea ecx, var_18
  loc_00E08779: call edi
  loc_00E0877B: sub esp, 00000010h
  loc_00E0877E: mov ecx, 0000000Bh
  loc_00E08783: mov edx, esp
  loc_00E08785: xor eax, eax
  loc_00E08787: push 80010007h
  loc_00E0878C: push esi
  loc_00E0878D: mov [edx], ecx
  loc_00E0878F: mov ecx, var_24
  loc_00E08792: mov [edx+00000004h], ecx
  loc_00E08795: mov ecx, [esi]
  loc_00E08797: mov [edx+00000008h], eax
  loc_00E0879A: mov eax, var_1C
  loc_00E0879D: mov [edx+0000000Ch], eax
  loc_00E087A0: call [ecx+000003B4h]
  loc_00E087A6: lea edx, var_18
  loc_00E087A9: push eax
  loc_00E087AA: push edx
  loc_00E087AB: call ebx
  loc_00E087AD: push eax
  loc_00E087AE: call [00401238h] ; __vbaLateIdSt
  loc_00E087B4: lea ecx, var_18
  loc_00E087B7: call edi
  loc_00E087B9: mov eax, [esi]
  loc_00E087BB: push esi
  loc_00E087BC: call [eax+00000330h]
  loc_00E087C2: lea ecx, var_18
  loc_00E087C5: push eax
  loc_00E087C6: push ecx
  loc_00E087C7: call ebx
  loc_00E087C9: mov esi, eax
  loc_00E087CB: push FFFFFFFFh
  loc_00E087CD: push esi
  loc_00E087CE: mov edx, [esi]
  loc_00E087D0: call [edx+0000005Ch]
  loc_00E087D3: test eax, eax
  loc_00E087D5: fnclex
  loc_00E087D7: jge 00E087E8h
  loc_00E087D9: push 0000005Ch
  loc_00E087DB: push 006DCAE0h
  loc_00E087E0: push esi
  loc_00E087E1: push eax
  loc_00E087E2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E087E8: lea ecx, var_18
  loc_00E087EB: call edi
  loc_00E087ED: mov var_4, 00000000h
  loc_00E087F4: push 00E08806h
  loc_00E087F9: jmp 00E08805h
  loc_00E087FB: lea ecx, var_18
  loc_00E087FE: call [00401254h] ; __vbaFreeObj
  loc_00E08804: ret
  loc_00E08805: ret
  loc_00E08806: mov eax, Me
  loc_00E08809: push eax
  loc_00E0880A: mov ecx, [eax]
  loc_00E0880C: call [ecx+00000008h]
  loc_00E0880F: mov eax, var_4
  loc_00E08812: mov ecx, var_14
  loc_00E08815: pop edi
  loc_00E08816: pop esi
  loc_00E08817: mov fs:[00000000h], ecx
  loc_00E0881E: pop ebx
  loc_00E0881F: mov esp, ebp
  loc_00E08821: pop ebp
  loc_00E08822: retn 0004h
End Sub

Private Sub baru_UnknownEvent_9 'E07B50
  loc_00E07B50: push ebp
  loc_00E07B51: mov ebp, esp
  loc_00E07B53: sub esp, 0000000Ch
  loc_00E07B56: push 00402836h ; __vbaExceptHandler
  loc_00E07B5B: mov eax, fs:[00000000h]
  loc_00E07B61: push eax
  loc_00E07B62: mov fs:[00000000h], esp
  loc_00E07B69: sub esp, 00000034h
  loc_00E07B6C: push ebx
  loc_00E07B6D: push esi
  loc_00E07B6E: push edi
  loc_00E07B6F: mov var_C, esp
  loc_00E07B72: mov var_8, 00402080h
  loc_00E07B79: mov esi, Me
  loc_00E07B7C: mov eax, esi
  loc_00E07B7E: and eax, 00000001h
  loc_00E07B81: mov var_4, eax
  loc_00E07B84: and esi, FFFFFFFEh
  loc_00E07B87: push esi
  loc_00E07B88: mov Me, esi
  loc_00E07B8B: mov ecx, [esi]
  loc_00E07B8D: call [ecx+00000004h]
  loc_00E07B90: sub esp, 00000010h
  loc_00E07B93: mov ecx, 0000000Bh
  loc_00E07B98: mov edx, esp
  loc_00E07B9A: xor eax, eax
  loc_00E07B9C: mov var_18, eax
  loc_00E07B9F: push 80010007h
  loc_00E07BA4: mov [edx], ecx
  loc_00E07BA6: mov ecx, var_24
  loc_00E07BA9: push esi
  loc_00E07BAA: mov [edx+00000004h], ecx
  loc_00E07BAD: mov ecx, [esi]
  loc_00E07BAF: mov [edx+00000008h], eax
  loc_00E07BB2: mov eax, var_1C
  loc_00E07BB5: mov [edx+0000000Ch], eax
  loc_00E07BB8: call [ecx+000003B8h]
  loc_00E07BBE: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E07BC4: lea edx, var_18
  loc_00E07BC7: push eax
  loc_00E07BC8: push edx
  loc_00E07BC9: call edi
  loc_00E07BCB: push eax
  loc_00E07BCC: call [00401238h] ; __vbaLateIdSt
  loc_00E07BD2: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E07BD8: lea ecx, var_18
  loc_00E07BDB: call ebx
  loc_00E07BDD: mov eax, [esi]
  loc_00E07BDF: push esi
  loc_00E07BE0: call [eax+0000039Ch]
  loc_00E07BE6: lea ecx, var_18
  loc_00E07BE9: push eax
  loc_00E07BEA: push ecx
  loc_00E07BEB: call edi
  loc_00E07BED: mov edx, [eax]
  loc_00E07BEF: push 006E03C8h ; "00001"
  loc_00E07BF4: push eax
  loc_00E07BF5: mov var_3C, eax
  loc_00E07BF8: call [edx+000000A4h]
  loc_00E07BFE: test eax, eax
  loc_00E07C00: fnclex
  loc_00E07C02: jge 00E07C19h
  loc_00E07C04: mov ecx, var_3C
  loc_00E07C07: push 000000A4h
  loc_00E07C0C: push 006DCB70h
  loc_00E07C11: push ecx
  loc_00E07C12: push eax
  loc_00E07C13: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07C19: lea ecx, var_18
  loc_00E07C1C: call ebx
  loc_00E07C1E: mov edx, [esi]
  loc_00E07C20: push esi
  loc_00E07C21: call [edx+00000394h]
  loc_00E07C27: push eax
  loc_00E07C28: lea eax, var_18
  loc_00E07C2B: push eax
  loc_00E07C2C: call edi
  loc_00E07C2E: mov ecx, [eax]
  loc_00E07C30: push 006DCC80h
  loc_00E07C35: push eax
  loc_00E07C36: mov var_3C, eax
  loc_00E07C39: call [ecx+000000A4h]
  loc_00E07C3F: test eax, eax
  loc_00E07C41: fnclex
  loc_00E07C43: jge 00E07C5Ah
  loc_00E07C45: mov edx, var_3C
  loc_00E07C48: push 000000A4h
  loc_00E07C4D: push 006DCB70h
  loc_00E07C52: push edx
  loc_00E07C53: push eax
  loc_00E07C54: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07C5A: lea ecx, var_18
  loc_00E07C5D: call ebx
  loc_00E07C5F: mov eax, [esi]
  loc_00E07C61: push esi
  loc_00E07C62: call [eax+00000380h]
  loc_00E07C68: lea ecx, var_18
  loc_00E07C6B: push eax
  loc_00E07C6C: push ecx
  loc_00E07C6D: call edi
  loc_00E07C6F: mov edx, [eax]
  loc_00E07C71: push FFFFFFFFh
  loc_00E07C73: push eax
  loc_00E07C74: mov var_3C, eax
  loc_00E07C77: call [edx+0000009Ch]
  loc_00E07C7D: test eax, eax
  loc_00E07C7F: fnclex
  loc_00E07C81: jge 00E07C98h
  loc_00E07C83: mov ecx, var_3C
  loc_00E07C86: push 0000009Ch
  loc_00E07C8B: push 006E03D4h
  loc_00E07C90: push ecx
  loc_00E07C91: push eax
  loc_00E07C92: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07C98: lea ecx, var_18
  loc_00E07C9B: call ebx
  loc_00E07C9D: mov edx, [esi]
  loc_00E07C9F: push esi
  loc_00E07CA0: call [edx+0000037Ch]
  loc_00E07CA6: push eax
  loc_00E07CA7: lea eax, var_18
  loc_00E07CAA: push eax
  loc_00E07CAB: call edi
  loc_00E07CAD: mov ecx, [eax]
  loc_00E07CAF: push FFFFFFFFh
  loc_00E07CB1: push eax
  loc_00E07CB2: mov var_3C, eax
  loc_00E07CB5: call [ecx+0000009Ch]
  loc_00E07CBB: test eax, eax
  loc_00E07CBD: fnclex
  loc_00E07CBF: jge 00E07CD6h
  loc_00E07CC1: mov edx, var_3C
  loc_00E07CC4: push 0000009Ch
  loc_00E07CC9: push 006E03D4h
  loc_00E07CCE: push edx
  loc_00E07CCF: push eax
  loc_00E07CD0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07CD6: lea ecx, var_18
  loc_00E07CD9: call ebx
  loc_00E07CDB: mov eax, [esi]
  loc_00E07CDD: push esi
  loc_00E07CDE: call [eax+00000390h]
  loc_00E07CE4: lea ecx, var_18
  loc_00E07CE7: push eax
  loc_00E07CE8: push ecx
  loc_00E07CE9: call edi
  loc_00E07CEB: mov edx, [eax]
  loc_00E07CED: push 006DCC80h
  loc_00E07CF2: push eax
  loc_00E07CF3: mov var_3C, eax
  loc_00E07CF6: call [edx+000000A4h]
  loc_00E07CFC: test eax, eax
  loc_00E07CFE: fnclex
  loc_00E07D00: jge 00E07D17h
  loc_00E07D02: mov ecx, var_3C
  loc_00E07D05: push 000000A4h
  loc_00E07D0A: push 006DCB70h
  loc_00E07D0F: push ecx
  loc_00E07D10: push eax
  loc_00E07D11: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07D17: lea ecx, var_18
  loc_00E07D1A: call ebx
  loc_00E07D1C: mov edx, [esi]
  loc_00E07D1E: push esi
  loc_00E07D1F: call [edx+0000038Ch]
  loc_00E07D25: push eax
  loc_00E07D26: lea eax, var_18
  loc_00E07D29: push eax
  loc_00E07D2A: call edi
  loc_00E07D2C: mov ecx, [eax]
  loc_00E07D2E: push 006DCC80h
  loc_00E07D33: push eax
  loc_00E07D34: mov var_3C, eax
  loc_00E07D37: call [ecx+000000A4h]
  loc_00E07D3D: test eax, eax
  loc_00E07D3F: fnclex
  loc_00E07D41: jge 00E07D58h
  loc_00E07D43: mov edx, var_3C
  loc_00E07D46: push 000000A4h
  loc_00E07D4B: push 006DCB70h
  loc_00E07D50: push edx
  loc_00E07D51: push eax
  loc_00E07D52: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07D58: lea ecx, var_18
  loc_00E07D5B: call ebx
  loc_00E07D5D: mov eax, [esi]
  loc_00E07D5F: push esi
  loc_00E07D60: call [eax+00000388h]
  loc_00E07D66: lea ecx, var_18
  loc_00E07D69: push eax
  loc_00E07D6A: push ecx
  loc_00E07D6B: call edi
  loc_00E07D6D: mov edx, [eax]
  loc_00E07D6F: push 006DCC80h
  loc_00E07D74: push eax
  loc_00E07D75: mov var_3C, eax
  loc_00E07D78: call [edx+000000A4h]
  loc_00E07D7E: test eax, eax
  loc_00E07D80: fnclex
  loc_00E07D82: jge 00E07D99h
  loc_00E07D84: mov ecx, var_3C
  loc_00E07D87: push 000000A4h
  loc_00E07D8C: push 006DCB70h
  loc_00E07D91: push ecx
  loc_00E07D92: push eax
  loc_00E07D93: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07D99: lea ecx, var_18
  loc_00E07D9C: call ebx
  loc_00E07D9E: sub esp, 00000010h
  loc_00E07DA1: mov ecx, 00000002h
  loc_00E07DA6: mov edx, esp
  loc_00E07DA8: xor eax, eax
  loc_00E07DAA: push 00000014h
  loc_00E07DAC: push esi
  loc_00E07DAD: mov [edx], ecx
  loc_00E07DAF: mov ecx, var_24
  loc_00E07DB2: mov [edx+00000004h], ecx
  loc_00E07DB5: mov ecx, [esi]
  loc_00E07DB7: mov [edx+00000008h], eax
  loc_00E07DBA: mov eax, var_1C
  loc_00E07DBD: mov [edx+0000000Ch], eax
  loc_00E07DC0: call [ecx+0000048Ch]
  loc_00E07DC6: lea edx, var_18
  loc_00E07DC9: push eax
  loc_00E07DCA: push edx
  loc_00E07DCB: call edi
  loc_00E07DCD: push eax
  loc_00E07DCE: call [00401238h] ; __vbaLateIdSt
  loc_00E07DD4: lea ecx, var_18
  loc_00E07DD7: call ebx
  loc_00E07DD9: mov eax, [esi]
  loc_00E07DDB: push esi
  loc_00E07DDC: call [eax+00000378h]
  loc_00E07DE2: lea ecx, var_18
  loc_00E07DE5: push eax
  loc_00E07DE6: push ecx
  loc_00E07DE7: call edi
  loc_00E07DE9: mov edx, [eax]
  loc_00E07DEB: push 006E03E8h ; "-----Pilih"
  loc_00E07DF0: push eax
  loc_00E07DF1: mov var_3C, eax
  loc_00E07DF4: call [edx+000000ACh]
  loc_00E07DFA: test eax, eax
  loc_00E07DFC: fnclex
  loc_00E07DFE: jge 00E07E15h
  loc_00E07E00: mov ecx, var_3C
  loc_00E07E03: push 000000ACh
  loc_00E07E08: push 006E0400h
  loc_00E07E0D: push ecx
  loc_00E07E0E: push eax
  loc_00E07E0F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07E15: lea ecx, var_18
  loc_00E07E18: call ebx
  loc_00E07E1A: mov edx, [esi]
  loc_00E07E1C: push esi
  loc_00E07E1D: call [edx+00000384h]
  loc_00E07E23: push eax
  loc_00E07E24: lea eax, var_18
  loc_00E07E27: push eax
  loc_00E07E28: call edi
  loc_00E07E2A: mov ecx, [eax]
  loc_00E07E2C: push 006DCC80h
  loc_00E07E31: push eax
  loc_00E07E32: mov var_3C, eax
  loc_00E07E35: call [ecx+000000A4h]
  loc_00E07E3B: test eax, eax
  loc_00E07E3D: fnclex
  loc_00E07E3F: jge 00E07E56h
  loc_00E07E41: mov edx, var_3C
  loc_00E07E44: push 000000A4h
  loc_00E07E49: push 006DCB70h
  loc_00E07E4E: push edx
  loc_00E07E4F: push eax
  loc_00E07E50: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07E56: lea ecx, var_18
  loc_00E07E59: call ebx
  loc_00E07E5B: mov eax, [esi]
  loc_00E07E5D: push esi
  loc_00E07E5E: call [eax+00000374h]
  loc_00E07E64: lea ecx, var_18
  loc_00E07E67: push eax
  loc_00E07E68: push ecx
  loc_00E07E69: call edi
  loc_00E07E6B: mov edx, [eax]
  loc_00E07E6D: push 006DCC80h
  loc_00E07E72: push eax
  loc_00E07E73: mov var_3C, eax
  loc_00E07E76: call [edx+000000A4h]
  loc_00E07E7C: test eax, eax
  loc_00E07E7E: fnclex
  loc_00E07E80: jge 00E07E97h
  loc_00E07E82: mov ecx, var_3C
  loc_00E07E85: push 000000A4h
  loc_00E07E8A: push 006DCB70h
  loc_00E07E8F: push ecx
  loc_00E07E90: push eax
  loc_00E07E91: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07E97: lea ecx, var_18
  loc_00E07E9A: call ebx
  loc_00E07E9C: mov edx, [esi]
  loc_00E07E9E: push esi
  loc_00E07E9F: call [edx+00000370h]
  loc_00E07EA5: push eax
  loc_00E07EA6: lea eax, var_18
  loc_00E07EA9: push eax
  loc_00E07EAA: call edi
  loc_00E07EAC: mov ecx, [eax]
  loc_00E07EAE: push 006DCC80h
  loc_00E07EB3: push eax
  loc_00E07EB4: mov var_3C, eax
  loc_00E07EB7: call [ecx+000000A4h]
  loc_00E07EBD: test eax, eax
  loc_00E07EBF: fnclex
  loc_00E07EC1: jge 00E07ED8h
  loc_00E07EC3: mov edx, var_3C
  loc_00E07EC6: push 000000A4h
  loc_00E07ECB: push 006DCB70h
  loc_00E07ED0: push edx
  loc_00E07ED1: push eax
  loc_00E07ED2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07ED8: lea ecx, var_18
  loc_00E07EDB: call ebx
  loc_00E07EDD: mov eax, [esi]
  loc_00E07EDF: push esi
  loc_00E07EE0: call [eax+0000036Ch]
  loc_00E07EE6: lea ecx, var_18
  loc_00E07EE9: push eax
  loc_00E07EEA: push ecx
  loc_00E07EEB: call edi
  loc_00E07EED: mov edx, [eax]
  loc_00E07EEF: push 006DCC80h
  loc_00E07EF4: push eax
  loc_00E07EF5: mov var_3C, eax
  loc_00E07EF8: call [edx+000000A4h]
  loc_00E07EFE: test eax, eax
  loc_00E07F00: fnclex
  loc_00E07F02: jge 00E07F19h
  loc_00E07F04: mov ecx, var_3C
  loc_00E07F07: push 000000A4h
  loc_00E07F0C: push 006DCB70h
  loc_00E07F11: push ecx
  loc_00E07F12: push eax
  loc_00E07F13: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07F19: lea ecx, var_18
  loc_00E07F1C: call ebx
  loc_00E07F1E: mov edx, [esi]
  loc_00E07F20: push esi
  loc_00E07F21: call [edx+0000030Ch]
  loc_00E07F27: push eax
  loc_00E07F28: lea eax, var_18
  loc_00E07F2B: push eax
  loc_00E07F2C: call edi
  loc_00E07F2E: mov ecx, [eax]
  loc_00E07F30: push 006E03E8h ; "-----Pilih"
  loc_00E07F35: push eax
  loc_00E07F36: mov var_3C, eax
  loc_00E07F39: call [ecx+000000ACh]
  loc_00E07F3F: test eax, eax
  loc_00E07F41: fnclex
  loc_00E07F43: jge 00E07F5Ah
  loc_00E07F45: mov edx, var_3C
  loc_00E07F48: push 000000ACh
  loc_00E07F4D: push 006E0400h
  loc_00E07F52: push edx
  loc_00E07F53: push eax
  loc_00E07F54: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07F5A: lea ecx, var_18
  loc_00E07F5D: call ebx
  loc_00E07F5F: mov eax, [esi]
  loc_00E07F61: push esi
  loc_00E07F62: call [eax+00000368h]
  loc_00E07F68: lea ecx, var_18
  loc_00E07F6B: push eax
  loc_00E07F6C: push ecx
  loc_00E07F6D: call edi
  loc_00E07F6F: mov edx, [eax]
  loc_00E07F71: push 006DCC80h
  loc_00E07F76: push eax
  loc_00E07F77: mov var_3C, eax
  loc_00E07F7A: call [edx+000000A4h]
  loc_00E07F80: test eax, eax
  loc_00E07F82: fnclex
  loc_00E07F84: jge 00E07F9Bh
  loc_00E07F86: mov ecx, var_3C
  loc_00E07F89: push 000000A4h
  loc_00E07F8E: push 006DCB70h
  loc_00E07F93: push ecx
  loc_00E07F94: push eax
  loc_00E07F95: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07F9B: lea ecx, var_18
  loc_00E07F9E: call ebx
  loc_00E07FA0: mov edx, [esi]
  loc_00E07FA2: push esi
  loc_00E07FA3: call [edx+00000364h]
  loc_00E07FA9: push eax
  loc_00E07FAA: lea eax, var_18
  loc_00E07FAD: push eax
  loc_00E07FAE: call edi
  loc_00E07FB0: mov ecx, [eax]
  loc_00E07FB2: push 006DCC80h
  loc_00E07FB7: push eax
  loc_00E07FB8: mov var_3C, eax
  loc_00E07FBB: call [ecx+000000A4h]
  loc_00E07FC1: test eax, eax
  loc_00E07FC3: fnclex
  loc_00E07FC5: jge 00E07FDCh
  loc_00E07FC7: mov edx, var_3C
  loc_00E07FCA: push 000000A4h
  loc_00E07FCF: push 006DCB70h
  loc_00E07FD4: push edx
  loc_00E07FD5: push eax
  loc_00E07FD6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E07FDC: lea ecx, var_18
  loc_00E07FDF: call ebx
  loc_00E07FE1: mov eax, [esi]
  loc_00E07FE3: push esi
  loc_00E07FE4: call [eax+00000360h]
  loc_00E07FEA: lea ecx, var_18
  loc_00E07FED: push eax
  loc_00E07FEE: push ecx
  loc_00E07FEF: call edi
  loc_00E07FF1: mov edx, [eax]
  loc_00E07FF3: push 006DCC80h
  loc_00E07FF8: push eax
  loc_00E07FF9: mov var_3C, eax
  loc_00E07FFC: call [edx+000000A4h]
  loc_00E08002: test eax, eax
  loc_00E08004: fnclex
  loc_00E08006: jge 00E0801Dh
  loc_00E08008: mov ecx, var_3C
  loc_00E0800B: push 000000A4h
  loc_00E08010: push 006DCB70h
  loc_00E08015: push ecx
  loc_00E08016: push eax
  loc_00E08017: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0801D: lea ecx, var_18
  loc_00E08020: call ebx
  loc_00E08022: mov edx, [esi]
  loc_00E08024: push esi
  loc_00E08025: call [edx+0000033Ch]
  loc_00E0802B: push eax
  loc_00E0802C: lea eax, var_18
  loc_00E0802F: push eax
  loc_00E08030: call edi
  loc_00E08032: mov ecx, [eax]
  loc_00E08034: push 006DCC80h
  loc_00E08039: push eax
  loc_00E0803A: mov var_3C, eax
  loc_00E0803D: call [ecx+000000A4h]
  loc_00E08043: test eax, eax
  loc_00E08045: fnclex
  loc_00E08047: jge 00E0805Eh
  loc_00E08049: mov edx, var_3C
  loc_00E0804C: push 000000A4h
  loc_00E08051: push 006DCB70h
  loc_00E08056: push edx
  loc_00E08057: push eax
  loc_00E08058: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0805E: lea ecx, var_18
  loc_00E08061: call ebx
  loc_00E08063: mov eax, [esi]
  loc_00E08065: push esi
  loc_00E08066: call [eax+00000340h]
  loc_00E0806C: lea ecx, var_18
  loc_00E0806F: push eax
  loc_00E08070: push ecx
  loc_00E08071: call edi
  loc_00E08073: mov edx, [eax]
  loc_00E08075: push 006DCC80h
  loc_00E0807A: push eax
  loc_00E0807B: mov var_3C, eax
  loc_00E0807E: call [edx+000000A4h]
  loc_00E08084: test eax, eax
  loc_00E08086: fnclex
  loc_00E08088: jge 00E0809Fh
  loc_00E0808A: mov ecx, var_3C
  loc_00E0808D: push 000000A4h
  loc_00E08092: push 006DCB70h
  loc_00E08097: push ecx
  loc_00E08098: push eax
  loc_00E08099: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0809F: lea ecx, var_18
  loc_00E080A2: call ebx
  loc_00E080A4: mov edx, [esi]
  loc_00E080A6: push esi
  loc_00E080A7: call [edx+00000344h]
  loc_00E080AD: push eax
  loc_00E080AE: lea eax, var_18
  loc_00E080B1: push eax
  loc_00E080B2: call edi
  loc_00E080B4: mov ecx, [eax]
  loc_00E080B6: push 006DCC80h
  loc_00E080BB: push eax
  loc_00E080BC: mov var_3C, eax
  loc_00E080BF: call [ecx+000000A4h]
  loc_00E080C5: test eax, eax
  loc_00E080C7: fnclex
  loc_00E080C9: jge 00E080E0h
  loc_00E080CB: mov edx, var_3C
  loc_00E080CE: push 000000A4h
  loc_00E080D3: push 006DCB70h
  loc_00E080D8: push edx
  loc_00E080D9: push eax
  loc_00E080DA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E080E0: lea ecx, var_18
  loc_00E080E3: call ebx
  loc_00E080E5: mov eax, [esi]
  loc_00E080E7: push esi
  loc_00E080E8: call [eax+00000348h]
  loc_00E080EE: lea ecx, var_18
  loc_00E080F1: push eax
  loc_00E080F2: push ecx
  loc_00E080F3: call edi
  loc_00E080F5: mov edx, [eax]
  loc_00E080F7: push 006DCC80h
  loc_00E080FC: push eax
  loc_00E080FD: mov var_3C, eax
  loc_00E08100: call [edx+000000A4h]
  loc_00E08106: test eax, eax
  loc_00E08108: fnclex
  loc_00E0810A: jge 00E08121h
  loc_00E0810C: mov ecx, var_3C
  loc_00E0810F: push 000000A4h
  loc_00E08114: push 006DCB70h
  loc_00E08119: push ecx
  loc_00E0811A: push eax
  loc_00E0811B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08121: lea ecx, var_18
  loc_00E08124: call ebx
  loc_00E08126: mov edx, [esi]
  loc_00E08128: push esi
  loc_00E08129: call [edx+000003C8h]
  loc_00E0812F: push eax
  loc_00E08130: lea eax, var_18
  loc_00E08133: push eax
  loc_00E08134: call edi
  loc_00E08136: mov ecx, [eax]
  loc_00E08138: push 00000000h
  loc_00E0813A: push eax
  loc_00E0813B: mov var_3C, eax
  loc_00E0813E: call [ecx+0000008Ch]
  loc_00E08144: test eax, eax
  loc_00E08146: fnclex
  loc_00E08148: jge 00E0815Fh
  loc_00E0814A: mov edx, var_3C
  loc_00E0814D: push 0000008Ch
  loc_00E08152: push 006DCDA0h
  loc_00E08157: push edx
  loc_00E08158: push eax
  loc_00E08159: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0815F: lea ecx, var_18
  loc_00E08162: call ebx
  loc_00E08164: mov eax, [esi]
  loc_00E08166: push esi
  loc_00E08167: call [eax+0000032Ch]
  loc_00E0816D: lea ecx, var_18
  loc_00E08170: push eax
  loc_00E08171: push ecx
  loc_00E08172: call edi
  loc_00E08174: mov edx, [eax]
  loc_00E08176: push 0080FF80h
  loc_00E0817B: push eax
  loc_00E0817C: mov var_3C, eax
  loc_00E0817F: call [edx+0000006Ch]
  loc_00E08182: test eax, eax
  loc_00E08184: fnclex
  loc_00E08186: jge 00E0819Ah
  loc_00E08188: mov ecx, var_3C
  loc_00E0818B: push 0000006Ch
  loc_00E0818D: push 006E0410h
  loc_00E08192: push ecx
  loc_00E08193: push eax
  loc_00E08194: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0819A: lea ecx, var_18
  loc_00E0819D: call ebx
  loc_00E0819F: mov edx, [esi]
  loc_00E081A1: push esi
  loc_00E081A2: call [edx+0000032Ch]
  loc_00E081A8: push eax
  loc_00E081A9: lea eax, var_18
  loc_00E081AC: push eax
  loc_00E081AD: call edi
  loc_00E081AF: mov ecx, [eax]
  loc_00E081B1: push 006E0424h ; "Silahkan Input Data Calon Peserta Didik Secara Lengkap !"
  loc_00E081B6: push eax
  loc_00E081B7: mov var_3C, eax
  loc_00E081BA: call [ecx+00000054h]
  loc_00E081BD: test eax, eax
  loc_00E081BF: fnclex
  loc_00E081C1: jge 00E081D5h
  loc_00E081C3: mov edx, var_3C
  loc_00E081C6: push 00000054h
  loc_00E081C8: push 006E0410h
  loc_00E081CD: push edx
  loc_00E081CE: push eax
  loc_00E081CF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E081D5: lea ecx, var_18
  loc_00E081D8: call ebx
  loc_00E081DA: mov eax, [esi]
  loc_00E081DC: push esi
  loc_00E081DD: call [eax+00000330h]
  loc_00E081E3: lea ecx, var_18
  loc_00E081E6: push eax
  loc_00E081E7: push ecx
  loc_00E081E8: call edi
  loc_00E081EA: mov edx, [eax]
  loc_00E081EC: push FFFFFFFFh
  loc_00E081EE: push eax
  loc_00E081EF: mov var_3C, eax
  loc_00E081F2: call [edx+0000005Ch]
  loc_00E081F5: test eax, eax
  loc_00E081F7: fnclex
  loc_00E081F9: jge 00E0820Dh
  loc_00E081FB: mov ecx, var_3C
  loc_00E081FE: push 0000005Ch
  loc_00E08200: push 006DCAE0h
  loc_00E08205: push ecx
  loc_00E08206: push eax
  loc_00E08207: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0820D: lea ecx, var_18
  loc_00E08210: call ebx
  loc_00E08212: sub esp, 00000010h
  loc_00E08215: mov ecx, 0000000Bh
  loc_00E0821A: mov edx, esp
  loc_00E0821C: or eax, FFFFFFFFh
  loc_00E0821F: push 8001000Dh
  loc_00E08224: push esi
  loc_00E08225: mov [edx], ecx
  loc_00E08227: mov ecx, var_24
  loc_00E0822A: mov [edx+00000004h], ecx
  loc_00E0822D: mov ecx, [esi]
  loc_00E0822F: mov [edx+00000008h], eax
  loc_00E08232: mov eax, var_1C
  loc_00E08235: mov [edx+0000000Ch], eax
  loc_00E08238: call [ecx+000003A0h]
  loc_00E0823E: lea edx, var_18
  loc_00E08241: push eax
  loc_00E08242: push edx
  loc_00E08243: call edi
  loc_00E08245: push eax
  loc_00E08246: call [00401238h] ; __vbaLateIdSt
  loc_00E0824C: lea ecx, var_18
  loc_00E0824F: call ebx
  loc_00E08251: mov var_4, 00000000h
  loc_00E08258: push 00E0826Ah
  loc_00E0825D: jmp 00E08269h
  loc_00E0825F: lea ecx, var_18
  loc_00E08262: call [00401254h] ; __vbaFreeObj
  loc_00E08268: ret
  loc_00E08269: ret
  loc_00E0826A: mov eax, Me
  loc_00E0826D: push eax
  loc_00E0826E: mov ecx, [eax]
  loc_00E08270: call [ecx+00000008h]
  loc_00E08273: mov eax, var_4
  loc_00E08276: mov ecx, var_14
  loc_00E08279: pop edi
  loc_00E0827A: pop esi
  loc_00E0827B: mov fs:[00000000h], ecx
  loc_00E08282: pop ebx
  loc_00E08283: mov esp, ebp
  loc_00E08285: pop ebp
  loc_00E08286: retn 0004h
End Sub

Private Sub nawali_KeyPress(KeyAscii As Integer) 'E09880
  loc_00E09880: push ebp
  loc_00E09881: mov ebp, esp
  loc_00E09883: sub esp, 0000000Ch
  loc_00E09886: push 00402836h ; __vbaExceptHandler
  loc_00E0988B: mov eax, fs:[00000000h]
  loc_00E09891: push eax
  loc_00E09892: mov fs:[00000000h], esp
  loc_00E09899: sub esp, 00000014h
  loc_00E0989C: push ebx
  loc_00E0989D: push esi
  loc_00E0989E: push edi
  loc_00E0989F: mov var_C, esp
  loc_00E098A2: mov var_8, 00402110h
  loc_00E098A9: mov esi, Me
  loc_00E098AC: mov eax, esi
  loc_00E098AE: and eax, 00000001h
  loc_00E098B1: mov var_4, eax
  loc_00E098B4: and esi, FFFFFFFEh
  loc_00E098B7: push esi
  loc_00E098B8: mov Me, esi
  loc_00E098BB: mov ecx, [esi]
  loc_00E098BD: call [ecx+00000004h]
  loc_00E098C0: mov edx, KeyAscii
  loc_00E098C3: xor edi, edi
  loc_00E098C5: mov var_18, edi
  loc_00E098C8: cmp [edx], 000Dh
  loc_00E098CC: jnz 00E0990Eh
  loc_00E098CE: mov eax, [esi]
  loc_00E098D0: push esi
  loc_00E098D1: call [eax+00000340h]
  loc_00E098D7: lea ecx, var_18
  loc_00E098DA: push eax
  loc_00E098DB: push ecx
  loc_00E098DC: call [004010ACh] ; __vbaObjSet
  loc_00E098E2: mov esi, eax
  loc_00E098E4: push esi
  loc_00E098E5: mov edx, [esi]
  loc_00E098E7: call [edx+00000204h]
  loc_00E098ED: cmp eax, edi
  loc_00E098EF: fnclex
  loc_00E098F1: jge 00E09905h
  loc_00E098F3: push 00000204h
  loc_00E098F8: push 006DCB70h
  loc_00E098FD: push esi
  loc_00E098FE: push eax
  loc_00E098FF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09905: lea ecx, var_18
  loc_00E09908: call [00401254h] ; __vbaFreeObj
  loc_00E0990E: mov var_4, edi
  loc_00E09911: push 00E09923h
  loc_00E09916: jmp 00E09922h
  loc_00E09918: lea ecx, var_18
  loc_00E0991B: call [00401254h] ; __vbaFreeObj
  loc_00E09921: ret
  loc_00E09922: ret
  loc_00E09923: mov eax, Me
  loc_00E09926: push eax
  loc_00E09927: mov ecx, [eax]
  loc_00E09929: call [ecx+00000008h]
  loc_00E0992C: mov eax, var_4
  loc_00E0992F: mov ecx, var_14
  loc_00E09932: pop edi
  loc_00E09933: pop esi
  loc_00E09934: mov fs:[00000000h], ecx
  loc_00E0993B: pop ebx
  loc_00E0993C: mov esp, ebp
  loc_00E0993E: pop ebp
  loc_00E0993F: retn 0008h
End Sub

Private Sub optno_Click() 'E0A6B0
  loc_00E0A6B0: push ebp
  loc_00E0A6B1: mov ebp, esp
  loc_00E0A6B3: sub esp, 0000000Ch
  loc_00E0A6B6: push 00402836h ; __vbaExceptHandler
  loc_00E0A6BB: mov eax, fs:[00000000h]
  loc_00E0A6C1: push eax
  loc_00E0A6C2: mov fs:[00000000h], esp
  loc_00E0A6C9: sub esp, 00000014h
  loc_00E0A6CC: push ebx
  loc_00E0A6CD: push esi
  loc_00E0A6CE: push edi
  loc_00E0A6CF: mov var_C, esp
  loc_00E0A6D2: mov var_8, 004021F0h
  loc_00E0A6D9: mov esi, Me
  loc_00E0A6DC: mov eax, esi
  loc_00E0A6DE: and eax, 00000001h
  loc_00E0A6E1: mov var_4, eax
  loc_00E0A6E4: and esi, FFFFFFFEh
  loc_00E0A6E7: push esi
  loc_00E0A6E8: mov Me, esi
  loc_00E0A6EB: mov ecx, [esi]
  loc_00E0A6ED: call [ecx+00000004h]
  loc_00E0A6F0: mov edx, [esi]
  loc_00E0A6F2: xor edi, edi
  loc_00E0A6F4: push esi
  loc_00E0A6F5: mov var_18, edi
  loc_00E0A6F8: call [edx+000003E0h]
  loc_00E0A6FE: push eax
  loc_00E0A6FF: lea eax, var_18
  loc_00E0A702: push eax
  loc_00E0A703: call [004010ACh] ; __vbaObjSet
  loc_00E0A709: mov esi, eax
  loc_00E0A70B: push edi
  loc_00E0A70C: push esi
  loc_00E0A70D: mov ecx, [esi]
  loc_00E0A70F: call [ecx+0000008Ch]
  loc_00E0A715: cmp eax, edi
  loc_00E0A717: fnclex
  loc_00E0A719: jge 00E0A72Dh
  loc_00E0A71B: push 0000008Ch
  loc_00E0A720: push 006DCDA0h
  loc_00E0A725: push esi
  loc_00E0A726: push eax
  loc_00E0A727: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A72D: lea ecx, var_18
  loc_00E0A730: call [00401254h] ; __vbaFreeObj
  loc_00E0A736: mov var_4, edi
  loc_00E0A739: push 00E0A74Bh
  loc_00E0A73E: jmp 00E0A74Ah
  loc_00E0A740: lea ecx, var_18
  loc_00E0A743: call [00401254h] ; __vbaFreeObj
  loc_00E0A749: ret
  loc_00E0A74A: ret
  loc_00E0A74B: mov eax, Me
  loc_00E0A74E: push eax
  loc_00E0A74F: mov edx, [eax]
  loc_00E0A751: call [edx+00000008h]
  loc_00E0A754: mov eax, var_4
  loc_00E0A757: mov ecx, var_14
  loc_00E0A75A: pop edi
  loc_00E0A75B: pop esi
  loc_00E0A75C: mov fs:[00000000h], ecx
  loc_00E0A763: pop ebx
  loc_00E0A764: mov esp, ebp
  loc_00E0A766: pop ebp
  loc_00E0A767: retn 0004h
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer) 'E100C0
  loc_00E100C0: push ebp
  loc_00E100C1: mov ebp, esp
  loc_00E100C3: sub esp, 0000000Ch
  loc_00E100C6: push 00402836h ; __vbaExceptHandler
  loc_00E100CB: mov eax, fs:[00000000h]
  loc_00E100D1: push eax
  loc_00E100D2: mov fs:[00000000h], esp
  loc_00E100D9: sub esp, 00000194h
  loc_00E100DF: push ebx
  loc_00E100E0: push esi
  loc_00E100E1: push edi
  loc_00E100E2: mov var_C, esp
  loc_00E100E5: mov var_8, 004022C8h
  loc_00E100EC: mov esi, Me
  loc_00E100EF: mov eax, esi
  loc_00E100F1: and eax, 00000001h
  loc_00E100F4: mov var_4, eax
  loc_00E100F7: and esi, FFFFFFFEh
  loc_00E100FA: push esi
  loc_00E100FB: mov Me, esi
  loc_00E100FE: mov ecx, [esi]
  loc_00E10100: call [ecx+00000004h]
  loc_00E10103: mov edx, [esi]
  loc_00E10105: xor ebx, ebx
  loc_00E10107: push esi
  loc_00E10108: mov var_18, ebx
  loc_00E1010B: mov var_1C, ebx
  loc_00E1010E: mov var_20, ebx
  loc_00E10111: mov var_24, ebx
  loc_00E10114: mov var_28, ebx
  loc_00E10117: mov var_2C, ebx
  loc_00E1011A: mov var_30, ebx
  loc_00E1011D: mov var_34, ebx
  loc_00E10120: mov var_44, ebx
  loc_00E10123: mov var_54, ebx
  loc_00E10126: mov var_64, ebx
  loc_00E10129: mov var_74, ebx
  loc_00E1012C: mov var_84, ebx
  loc_00E10132: mov var_94, ebx
  loc_00E10138: mov var_B8, ebx
  loc_00E1013E: call [edx+00000320h]
  loc_00E10144: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E1014A: push eax
  loc_00E1014B: lea eax, var_24
  loc_00E1014E: push eax
  loc_00E1014F: call edi
  loc_00E10151: mov ecx, [eax]
  loc_00E10153: lea edx, var_B8
  loc_00E10159: push edx
  loc_00E1015A: push eax
  loc_00E1015B: mov var_BC, eax
  loc_00E10161: call [ecx+000000E0h]
  loc_00E10167: cmp eax, ebx
  loc_00E10169: fnclex
  loc_00E1016B: jge 00E10185h
  loc_00E1016D: mov ecx, var_BC
  loc_00E10173: push 000000E0h
  loc_00E10178: push 006E03D4h
  loc_00E1017D: push ecx
  loc_00E1017E: push eax
  loc_00E1017F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10185: mov dx, var_B8
  loc_00E1018C: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E10192: lea ecx, var_24
  loc_00E10195: mov var_C4, dx
  loc_00E1019C: call ebx
  loc_00E1019E: cmp var_C4, 0000h
  loc_00E101A6: jz 00E1201Ch
  loc_00E101AC: mov eax, [esi]
  loc_00E101AE: push esi
  loc_00E101AF: call [eax+00000324h]
  loc_00E101B5: lea ecx, var_24
  loc_00E101B8: push eax
  loc_00E101B9: push ecx
  loc_00E101BA: call edi
  loc_00E101BC: mov edx, [eax]
  loc_00E101BE: lea ecx, var_18
  loc_00E101C1: push ecx
  loc_00E101C2: push eax
  loc_00E101C3: mov var_BC, eax
  loc_00E101C9: call [edx+000000A0h]
  loc_00E101CF: test eax, eax
  loc_00E101D1: fnclex
  loc_00E101D3: jge 00E101EDh
  loc_00E101D5: mov edx, var_BC
  loc_00E101DB: push 000000A0h
  loc_00E101E0: push 006DCB70h
  loc_00E101E5: push edx
  loc_00E101E6: push eax
  loc_00E101E7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E101ED: mov eax, var_18
  loc_00E101F0: push 006E0C48h ; "select * from biodata where no_daftar like '" & Chr(37)
  loc_00E101F5: push eax
  loc_00E101F6: call [0040106Ch] ; __vbaStrCat
  loc_00E101FC: mov edx, eax
  loc_00E101FE: lea ecx, var_1C
  loc_00E10201: call [00401228h] ; __vbaStrMove
  loc_00E10207: push eax
  loc_00E10208: push 006E0CA8h ; Chr(37) & "' order by no_daftar asc"
  loc_00E1020D: call [0040106Ch] ; __vbaStrCat
  loc_00E10213: mov edx, eax
  loc_00E10215: lea ecx, var_20
  loc_00E10218: call [00401228h] ; __vbaStrMove
  loc_00E1021E: mov edx, eax
  loc_00E10220: lea ecx, [esi+00000034h]
  loc_00E10223: call [004011B0h] ; __vbaStrCopy
  loc_00E10229: lea ecx, var_20
  loc_00E1022C: lea edx, var_1C
  loc_00E1022F: push ecx
  loc_00E10230: lea eax, var_18
  loc_00E10233: push edx
  loc_00E10234: push eax
  loc_00E10235: push 00000003h
  loc_00E10237: call [004011BCh] ; __vbaFreeStrList
  loc_00E1023D: add esp, 00000010h
  loc_00E10240: lea ecx, var_24
  loc_00E10243: call ebx
  loc_00E10245: sub esp, 00000010h
  loc_00E10248: mov ecx, 00004008h
  loc_00E1024D: mov edx, esp
  loc_00E1024F: mov var_84, ecx
  loc_00E10255: lea eax, [esi+00000034h]
  loc_00E10258: push 0000000Eh
  loc_00E1025A: mov [edx], ecx
  loc_00E1025C: mov ecx, var_80
  loc_00E1025F: mov var_7C, eax
  loc_00E10262: push esi
  loc_00E10263: mov [edx+00000004h], ecx
  loc_00E10266: mov ecx, [esi]
  loc_00E10268: mov [edx+00000008h], eax
  loc_00E1026B: mov eax, var_78
  loc_00E1026E: mov [edx+0000000Ch], eax
  loc_00E10271: call [ecx+00000490h]
  loc_00E10277: lea edx, var_24
  loc_00E1027A: push eax
  loc_00E1027B: push edx
  loc_00E1027C: call edi
  loc_00E1027E: push eax
  loc_00E1027F: call [00401238h] ; __vbaLateIdSt
  loc_00E10285: lea ecx, var_24
  loc_00E10288: call ebx
  loc_00E1028A: mov eax, [esi]
  loc_00E1028C: push 00000000h
  loc_00E1028E: push 00000019h
  loc_00E10290: push esi
  loc_00E10291: call [eax+00000490h]
  loc_00E10297: lea ecx, var_24
  loc_00E1029A: push eax
  loc_00E1029B: push ecx
  loc_00E1029C: call edi
  loc_00E1029E: push eax
  loc_00E1029F: call [00401030h] ; __vbaLateIdCall
  loc_00E102A5: add esp, 0000000Ch
  loc_00E102A8: lea ecx, var_24
  loc_00E102AB: call ebx
  loc_00E102AD: mov edx, [esi]
  loc_00E102AF: push 006DCBD8h
  loc_00E102B4: push 00000000h
  loc_00E102B6: push 00000018h
  loc_00E102B8: push esi
  loc_00E102B9: call [edx+00000490h]
  loc_00E102BF: push eax
  loc_00E102C0: lea eax, var_24
  loc_00E102C3: push eax
  loc_00E102C4: call edi
  loc_00E102C6: lea ecx, var_44
  loc_00E102C9: push eax
  loc_00E102CA: push ecx
  loc_00E102CB: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E102D1: add esp, 00000010h
  loc_00E102D4: push eax
  loc_00E102D5: call [00401120h] ; __vbaCastObjVar
  loc_00E102DB: lea edx, var_28
  loc_00E102DE: push eax
  loc_00E102DF: push edx
  loc_00E102E0: call edi
  loc_00E102E2: mov ebx, eax
  loc_00E102E4: lea ecx, var_B8
  loc_00E102EA: push ecx
  loc_00E102EB: push ebx
  loc_00E102EC: mov eax, [ebx]
  loc_00E102EE: call [eax+00000050h]
  loc_00E102F1: test eax, eax
  loc_00E102F3: fnclex
  loc_00E102F5: jge 00E10306h
  loc_00E102F7: push 00000050h
  loc_00E102F9: push 006DCBE8h
  loc_00E102FE: push ebx
  loc_00E102FF: push eax
  loc_00E10300: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10306: mov bx, var_B8
  loc_00E1030D: lea edx, var_28
  loc_00E10310: lea eax, var_24
  loc_00E10313: push edx
  loc_00E10314: push eax
  loc_00E10315: push 00000002h
  loc_00E10317: call [00401048h] ; __vbaFreeObjList
  loc_00E1031D: add esp, 0000000Ch
  loc_00E10320: lea ecx, var_44
  loc_00E10323: call [00401024h] ; __vbaFreeVar
  loc_00E10329: test bx, bx
  loc_00E1032C: jnz 00E12207h
  loc_00E10332: mov eax, [esi]
  loc_00E10334: push esi
  loc_00E10335: call [eax+0000039Ch]
  loc_00E1033B: lea ecx, var_34
  loc_00E1033E: push eax
  loc_00E1033F: push ecx
  loc_00E10340: call edi
  loc_00E10342: mov edx, [esi]
  loc_00E10344: push 006DCBD8h
  loc_00E10349: push 00000000h
  loc_00E1034B: push 00000018h
  loc_00E1034D: push esi
  loc_00E1034E: mov var_D4, eax
  loc_00E10354: call [edx+00000490h]
  loc_00E1035A: push eax
  loc_00E1035B: lea eax, var_24
  loc_00E1035E: push eax
  loc_00E1035F: call edi
  loc_00E10361: lea ecx, var_44
  loc_00E10364: push eax
  loc_00E10365: push ecx
  loc_00E10366: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1036C: add esp, 00000010h
  loc_00E1036F: push eax
  loc_00E10370: call [00401120h] ; __vbaCastObjVar
  loc_00E10376: lea edx, var_28
  loc_00E10379: push eax
  loc_00E1037A: push edx
  loc_00E1037B: call edi
  loc_00E1037D: mov ebx, eax
  loc_00E1037F: lea ecx, var_2C
  loc_00E10382: push ecx
  loc_00E10383: push ebx
  loc_00E10384: mov eax, [ebx]
  loc_00E10386: call [eax+00000054h]
  loc_00E10389: test eax, eax
  loc_00E1038B: fnclex
  loc_00E1038D: jge 00E1039Eh
  loc_00E1038F: push 00000054h
  loc_00E10391: push 006DCBE8h
  loc_00E10396: push ebx
  loc_00E10397: push eax
  loc_00E10398: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1039E: lea ebx, var_30
  loc_00E103A1: mov eax, var_2C
  loc_00E103A4: push ebx
  loc_00E103A5: mov ecx, 00000002h
  loc_00E103AA: sub esp, 00000010h
  loc_00E103AD: mov var_84, ecx
  loc_00E103B3: mov ebx, esp
  loc_00E103B5: mov var_7C, 00000000h
  loc_00E103BC: mov edx, [eax]
  loc_00E103BE: push eax
  loc_00E103BF: mov [ebx], ecx
  loc_00E103C1: mov ecx, var_80
  loc_00E103C4: mov var_C4, eax
  loc_00E103CA: mov [ebx+00000004h], ecx
  loc_00E103CD: mov ecx, var_7C
  loc_00E103D0: mov [ebx+00000008h], ecx
  loc_00E103D3: mov ecx, var_78
  loc_00E103D6: mov [ebx+0000000Ch], ecx
  loc_00E103D9: call [edx+00000028h]
  loc_00E103DC: test eax, eax
  loc_00E103DE: fnclex
  loc_00E103E0: jge 00E103F7h
  loc_00E103E2: mov edx, var_C4
  loc_00E103E8: push 00000028h
  loc_00E103EA: push 006E09E8h
  loc_00E103EF: push edx
  loc_00E103F0: push eax
  loc_00E103F1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E103F7: mov eax, var_30
  loc_00E103FA: lea edx, var_54
  loc_00E103FD: push edx
  loc_00E103FE: push eax
  loc_00E103FF: mov ecx, [eax]
  loc_00E10401: mov ebx, eax
  loc_00E10403: call [ecx+00000034h]
  loc_00E10406: test eax, eax
  loc_00E10408: fnclex
  loc_00E1040A: jge 00E1041Bh
  loc_00E1040C: push 00000034h
  loc_00E1040E: push 006E09F8h
  loc_00E10413: push ebx
  loc_00E10414: push eax
  loc_00E10415: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1041B: mov eax, var_D4
  loc_00E10421: lea ecx, var_54
  loc_00E10424: push ecx
  loc_00E10425: mov ebx, [eax]
  loc_00E10427: call [00401034h] ; __vbaStrVarMove
  loc_00E1042D: mov edx, eax
  loc_00E1042F: lea ecx, var_18
  loc_00E10432: call [00401228h] ; __vbaStrMove
  loc_00E10438: mov edx, ebx
  loc_00E1043A: mov ebx, var_D4
  loc_00E10440: push eax
  loc_00E10441: push ebx
  loc_00E10442: call [edx+000000A4h]
  loc_00E10448: test eax, eax
  loc_00E1044A: fnclex
  loc_00E1044C: jge 00E10460h
  loc_00E1044E: push 000000A4h
  loc_00E10453: push 006DCB70h
  loc_00E10458: push ebx
  loc_00E10459: push eax
  loc_00E1045A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10460: lea ecx, var_18
  loc_00E10463: call [00401258h] ; __vbaFreeStr
  loc_00E10469: lea eax, var_34
  loc_00E1046C: lea ecx, var_30
  loc_00E1046F: push eax
  loc_00E10470: lea edx, var_2C
  loc_00E10473: push ecx
  loc_00E10474: lea eax, var_28
  loc_00E10477: push edx
  loc_00E10478: lea ecx, var_24
  loc_00E1047B: push eax
  loc_00E1047C: push ecx
  loc_00E1047D: push 00000005h
  loc_00E1047F: call [00401048h] ; __vbaFreeObjList
  loc_00E10485: lea edx, var_54
  loc_00E10488: lea eax, var_44
  loc_00E1048B: push edx
  loc_00E1048C: push eax
  loc_00E1048D: push 00000002h
  loc_00E1048F: call [00401038h] ; __vbaFreeVarList
  loc_00E10495: mov ecx, [esi]
  loc_00E10497: add esp, 00000024h
  loc_00E1049A: push esi
  loc_00E1049B: call [ecx+00000394h]
  loc_00E104A1: lea edx, var_34
  loc_00E104A4: push eax
  loc_00E104A5: push edx
  loc_00E104A6: call edi
  loc_00E104A8: push 006DCBD8h
  loc_00E104AD: mov var_D4, eax
  loc_00E104B3: mov eax, [esi]
  loc_00E104B5: push 00000000h
  loc_00E104B7: push 00000018h
  loc_00E104B9: push esi
  loc_00E104BA: call [eax+00000490h]
  loc_00E104C0: lea ecx, var_24
  loc_00E104C3: push eax
  loc_00E104C4: push ecx
  loc_00E104C5: call edi
  loc_00E104C7: lea edx, var_44
  loc_00E104CA: push eax
  loc_00E104CB: push edx
  loc_00E104CC: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E104D2: add esp, 00000010h
  loc_00E104D5: push eax
  loc_00E104D6: call [00401120h] ; __vbaCastObjVar
  loc_00E104DC: push eax
  loc_00E104DD: lea eax, var_28
  loc_00E104E0: push eax
  loc_00E104E1: call edi
  loc_00E104E3: mov ebx, eax
  loc_00E104E5: lea edx, var_2C
  loc_00E104E8: push edx
  loc_00E104E9: push ebx
  loc_00E104EA: mov ecx, [ebx]
  loc_00E104EC: call [ecx+00000054h]
  loc_00E104EF: test eax, eax
  loc_00E104F1: fnclex
  loc_00E104F3: jge 00E10504h
  loc_00E104F5: push 00000054h
  loc_00E104F7: push 006DCBE8h
  loc_00E104FC: push ebx
  loc_00E104FD: push eax
  loc_00E104FE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10504: lea ebx, var_30
  loc_00E10507: mov eax, var_2C
  loc_00E1050A: push ebx
  loc_00E1050B: mov ecx, 00000002h
  loc_00E10510: sub esp, 00000010h
  loc_00E10513: mov var_84, ecx
  loc_00E10519: mov ebx, esp
  loc_00E1051B: mov var_7C, 00000001h
  loc_00E10522: mov edx, [eax]
  loc_00E10524: push eax
  loc_00E10525: mov [ebx], ecx
  loc_00E10527: mov ecx, var_80
  loc_00E1052A: mov var_C4, eax
  loc_00E10530: mov [ebx+00000004h], ecx
  loc_00E10533: mov ecx, var_7C
  loc_00E10536: mov [ebx+00000008h], ecx
  loc_00E10539: mov ecx, var_78
  loc_00E1053C: mov [ebx+0000000Ch], ecx
  loc_00E1053F: call [edx+00000028h]
  loc_00E10542: test eax, eax
  loc_00E10544: fnclex
  loc_00E10546: jge 00E1055Dh
  loc_00E10548: mov edx, var_C4
  loc_00E1054E: push 00000028h
  loc_00E10550: push 006E09E8h
  loc_00E10555: push edx
  loc_00E10556: push eax
  loc_00E10557: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1055D: mov eax, var_30
  loc_00E10560: lea edx, var_54
  loc_00E10563: push edx
  loc_00E10564: push eax
  loc_00E10565: mov ecx, [eax]
  loc_00E10567: mov ebx, eax
  loc_00E10569: call [ecx+00000034h]
  loc_00E1056C: test eax, eax
  loc_00E1056E: fnclex
  loc_00E10570: jge 00E10581h
  loc_00E10572: push 00000034h
  loc_00E10574: push 006E09F8h
  loc_00E10579: push ebx
  loc_00E1057A: push eax
  loc_00E1057B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10581: mov eax, var_D4
  loc_00E10587: lea ecx, var_54
  loc_00E1058A: push ecx
  loc_00E1058B: mov ebx, [eax]
  loc_00E1058D: call [00401034h] ; __vbaStrVarMove
  loc_00E10593: mov edx, eax
  loc_00E10595: lea ecx, var_18
  loc_00E10598: call [00401228h] ; __vbaStrMove
  loc_00E1059E: mov edx, ebx
  loc_00E105A0: mov ebx, var_D4
  loc_00E105A6: push eax
  loc_00E105A7: push ebx
  loc_00E105A8: call [edx+000000A4h]
  loc_00E105AE: test eax, eax
  loc_00E105B0: fnclex
  loc_00E105B2: jge 00E105C6h
  loc_00E105B4: push 000000A4h
  loc_00E105B9: push 006DCB70h
  loc_00E105BE: push ebx
  loc_00E105BF: push eax
  loc_00E105C0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E105C6: lea ecx, var_18
  loc_00E105C9: call [00401258h] ; __vbaFreeStr
  loc_00E105CF: lea eax, var_34
  loc_00E105D2: lea ecx, var_30
  loc_00E105D5: push eax
  loc_00E105D6: lea edx, var_2C
  loc_00E105D9: push ecx
  loc_00E105DA: lea eax, var_28
  loc_00E105DD: push edx
  loc_00E105DE: lea ecx, var_24
  loc_00E105E1: push eax
  loc_00E105E2: push ecx
  loc_00E105E3: push 00000005h
  loc_00E105E5: call [00401048h] ; __vbaFreeObjList
  loc_00E105EB: lea edx, var_54
  loc_00E105EE: lea eax, var_44
  loc_00E105F1: push edx
  loc_00E105F2: push eax
  loc_00E105F3: push 00000002h
  loc_00E105F5: call [00401038h] ; __vbaFreeVarList
  loc_00E105FB: mov ecx, [esi]
  loc_00E105FD: add esp, 00000024h
  loc_00E10600: push 006DCBD8h
  loc_00E10605: push 00000000h
  loc_00E10607: push 00000018h
  loc_00E10609: push esi
  loc_00E1060A: call [ecx+00000490h]
  loc_00E10610: lea edx, var_24
  loc_00E10613: push eax
  loc_00E10614: push edx
  loc_00E10615: call edi
  loc_00E10617: push eax
  loc_00E10618: lea eax, var_44
  loc_00E1061B: push eax
  loc_00E1061C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E10622: add esp, 00000010h
  loc_00E10625: push eax
  loc_00E10626: call [00401120h] ; __vbaCastObjVar
  loc_00E1062C: lea ecx, var_28
  loc_00E1062F: push eax
  loc_00E10630: push ecx
  loc_00E10631: call edi
  loc_00E10633: mov ebx, eax
  loc_00E10635: lea eax, var_2C
  loc_00E10638: push eax
  loc_00E10639: push ebx
  loc_00E1063A: mov edx, [ebx]
  loc_00E1063C: call [edx+00000054h]
  loc_00E1063F: test eax, eax
  loc_00E10641: fnclex
  loc_00E10643: jge 00E10654h
  loc_00E10645: push 00000054h
  loc_00E10647: push 006DCBE8h
  loc_00E1064C: push ebx
  loc_00E1064D: push eax
  loc_00E1064E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10654: lea ebx, var_30
  loc_00E10657: mov eax, var_2C
  loc_00E1065A: push ebx
  loc_00E1065B: mov ecx, 00000002h
  loc_00E10660: sub esp, 00000010h
  loc_00E10663: mov var_7C, ecx
  loc_00E10666: mov ebx, esp
  loc_00E10668: mov var_84, ecx
  loc_00E1066E: mov edx, [eax]
  loc_00E10670: push eax
  loc_00E10671: mov [ebx], ecx
  loc_00E10673: mov ecx, var_80
  loc_00E10676: mov var_C4, eax
  loc_00E1067C: mov [ebx+00000004h], ecx
  loc_00E1067F: mov ecx, var_7C
  loc_00E10682: mov [ebx+00000008h], ecx
  loc_00E10685: mov ecx, var_78
  loc_00E10688: mov [ebx+0000000Ch], ecx
  loc_00E1068B: call [edx+00000028h]
  loc_00E1068E: test eax, eax
  loc_00E10690: fnclex
  loc_00E10692: jge 00E106A9h
  loc_00E10694: mov edx, var_C4
  loc_00E1069A: push 00000028h
  loc_00E1069C: push 006E09E8h
  loc_00E106A1: push edx
  loc_00E106A2: push eax
  loc_00E106A3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E106A9: mov eax, var_30
  loc_00E106AC: lea edx, var_54
  loc_00E106AF: push edx
  loc_00E106B0: push eax
  loc_00E106B1: mov ecx, [eax]
  loc_00E106B3: mov ebx, eax
  loc_00E106B5: call [ecx+00000034h]
  loc_00E106B8: test eax, eax
  loc_00E106BA: fnclex
  loc_00E106BC: jge 00E106CDh
  loc_00E106BE: push 00000034h
  loc_00E106C0: push 006E09F8h
  loc_00E106C5: push ebx
  loc_00E106C6: push eax
  loc_00E106C7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E106CD: lea eax, var_54
  loc_00E106D0: lea ecx, var_94
  loc_00E106D6: push eax
  loc_00E106D7: push ecx
  loc_00E106D8: mov var_8C, 006E0B24h ; "Laki - laki"
  loc_00E106E2: mov var_94, 00008008h
  loc_00E106EC: call [00401108h] ; __vbaVarTstEq
  loc_00E106F2: mov bx, ax
  loc_00E106F5: lea edx, var_30
  loc_00E106F8: lea eax, var_2C
  loc_00E106FB: push edx
  loc_00E106FC: lea ecx, var_28
  loc_00E106FF: push eax
  loc_00E10700: lea edx, var_24
  loc_00E10703: push ecx
  loc_00E10704: push edx
  loc_00E10705: push 00000004h
  loc_00E10707: call [00401048h] ; __vbaFreeObjList
  loc_00E1070D: lea eax, var_54
  loc_00E10710: lea ecx, var_44
  loc_00E10713: push eax
  loc_00E10714: push ecx
  loc_00E10715: push 00000002h
  loc_00E10717: call [00401038h] ; __vbaFreeVarList
  loc_00E1071D: add esp, 00000020h
  loc_00E10720: test bx, bx
  loc_00E10723: jz 00E107D1h
  loc_00E10729: mov edx, [esi]
  loc_00E1072B: push esi
  loc_00E1072C: call [edx+00000380h]
  loc_00E10732: push eax
  loc_00E10733: lea eax, var_24
  loc_00E10736: push eax
  loc_00E10737: call edi
  loc_00E10739: mov ebx, eax
  loc_00E1073B: push FFFFFFFFh
  loc_00E1073D: push ebx
  loc_00E1073E: mov ecx, [ebx]
  loc_00E10740: call [ecx+000000E4h]
  loc_00E10746: test eax, eax
  loc_00E10748: fnclex
  loc_00E1074A: jge 00E1075Eh
  loc_00E1074C: push 000000E4h
  loc_00E10751: push 006E03D4h
  loc_00E10756: push ebx
  loc_00E10757: push eax
  loc_00E10758: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1075E: lea ecx, var_24
  loc_00E10761: call [00401254h] ; __vbaFreeObj
  loc_00E10767: mov edx, [esi]
  loc_00E10769: push esi
  loc_00E1076A: call [edx+0000037Ch]
  loc_00E10770: push eax
  loc_00E10771: lea eax, var_24
  loc_00E10774: push eax
  loc_00E10775: call edi
  loc_00E10777: mov ebx, eax
  loc_00E10779: push FFFFFFFFh
  loc_00E1077B: push ebx
  loc_00E1077C: mov ecx, [ebx]
  loc_00E1077E: call [ecx+0000009Ch]
  loc_00E10784: test eax, eax
  loc_00E10786: fnclex
  loc_00E10788: jge 00E1079Ch
  loc_00E1078A: push 0000009Ch
  loc_00E1078F: push 006E03D4h
  loc_00E10794: push ebx
  loc_00E10795: push eax
  loc_00E10796: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1079C: lea ecx, var_24
  loc_00E1079F: call [00401254h] ; __vbaFreeObj
  loc_00E107A5: mov edx, [esi]
  loc_00E107A7: push esi
  loc_00E107A8: call [edx+00000380h]
  loc_00E107AE: push eax
  loc_00E107AF: lea eax, var_24
  loc_00E107B2: push eax
  loc_00E107B3: call edi
  loc_00E107B5: mov ebx, eax
  loc_00E107B7: push FFFFFFFFh
  loc_00E107B9: push ebx
  loc_00E107BA: mov ecx, [ebx]
  loc_00E107BC: call [ecx+0000009Ch]
  loc_00E107C2: test eax, eax
  loc_00E107C4: fnclex
  loc_00E107C6: jge 00E109ADh
  loc_00E107CC: jmp 00E1099Bh
  loc_00E107D1: mov edx, [esi]
  loc_00E107D3: push 006DCBD8h
  loc_00E107D8: push 00000000h
  loc_00E107DA: push 00000018h
  loc_00E107DC: push esi
  loc_00E107DD: call [edx+00000490h]
  loc_00E107E3: push eax
  loc_00E107E4: lea eax, var_24
  loc_00E107E7: push eax
  loc_00E107E8: call edi
  loc_00E107EA: lea ecx, var_44
  loc_00E107ED: push eax
  loc_00E107EE: push ecx
  loc_00E107EF: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E107F5: add esp, 00000010h
  loc_00E107F8: push eax
  loc_00E107F9: call [00401120h] ; __vbaCastObjVar
  loc_00E107FF: lea edx, var_28
  loc_00E10802: push eax
  loc_00E10803: push edx
  loc_00E10804: call edi
  loc_00E10806: mov ebx, eax
  loc_00E10808: lea ecx, var_2C
  loc_00E1080B: push ecx
  loc_00E1080C: push ebx
  loc_00E1080D: mov eax, [ebx]
  loc_00E1080F: call [eax+00000054h]
  loc_00E10812: test eax, eax
  loc_00E10814: fnclex
  loc_00E10816: jge 00E10827h
  loc_00E10818: push 00000054h
  loc_00E1081A: push 006DCBE8h
  loc_00E1081F: push ebx
  loc_00E10820: push eax
  loc_00E10821: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10827: lea ebx, var_30
  loc_00E1082A: mov eax, var_2C
  loc_00E1082D: push ebx
  loc_00E1082E: mov ecx, 00000002h
  loc_00E10833: sub esp, 00000010h
  loc_00E10836: mov var_7C, ecx
  loc_00E10839: mov ebx, esp
  loc_00E1083B: mov var_84, ecx
  loc_00E10841: mov edx, [eax]
  loc_00E10843: push eax
  loc_00E10844: mov [ebx], ecx
  loc_00E10846: mov ecx, var_80
  loc_00E10849: mov var_C4, eax
  loc_00E1084F: mov [ebx+00000004h], ecx
  loc_00E10852: mov ecx, var_7C
  loc_00E10855: mov [ebx+00000008h], ecx
  loc_00E10858: mov ecx, var_78
  loc_00E1085B: mov [ebx+0000000Ch], ecx
  loc_00E1085E: call [edx+00000028h]
  loc_00E10861: test eax, eax
  loc_00E10863: fnclex
  loc_00E10865: jge 00E1087Ch
  loc_00E10867: mov edx, var_C4
  loc_00E1086D: push 00000028h
  loc_00E1086F: push 006E09E8h
  loc_00E10874: push edx
  loc_00E10875: push eax
  loc_00E10876: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1087C: mov eax, var_30
  loc_00E1087F: lea edx, var_54
  loc_00E10882: push edx
  loc_00E10883: push eax
  loc_00E10884: mov ecx, [eax]
  loc_00E10886: mov ebx, eax
  loc_00E10888: call [ecx+00000034h]
  loc_00E1088B: test eax, eax
  loc_00E1088D: fnclex
  loc_00E1088F: jge 00E108A0h
  loc_00E10891: push 00000034h
  loc_00E10893: push 006E09F8h
  loc_00E10898: push ebx
  loc_00E10899: push eax
  loc_00E1089A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E108A0: lea eax, var_54
  loc_00E108A3: lea ecx, var_94
  loc_00E108A9: push eax
  loc_00E108AA: push ecx
  loc_00E108AB: mov var_8C, 006E0B40h ; "Perempuan"
  loc_00E108B5: mov var_94, 00008008h
  loc_00E108BF: call [00401108h] ; __vbaVarTstEq
  loc_00E108C5: mov bx, ax
  loc_00E108C8: lea edx, var_30
  loc_00E108CB: lea eax, var_2C
  loc_00E108CE: push edx
  loc_00E108CF: lea ecx, var_28
  loc_00E108D2: push eax
  loc_00E108D3: lea edx, var_24
  loc_00E108D6: push ecx
  loc_00E108D7: push edx
  loc_00E108D8: push 00000004h
  loc_00E108DA: call [00401048h] ; __vbaFreeObjList
  loc_00E108E0: lea eax, var_54
  loc_00E108E3: lea ecx, var_44
  loc_00E108E6: push eax
  loc_00E108E7: push ecx
  loc_00E108E8: push 00000002h
  loc_00E108EA: call [00401038h] ; __vbaFreeVarList
  loc_00E108F0: add esp, 00000020h
  loc_00E108F3: test bx, bx
  loc_00E108F6: jz 00E109BBh
  loc_00E108FC: mov edx, [esi]
  loc_00E108FE: push esi
  loc_00E108FF: call [edx+0000037Ch]
  loc_00E10905: push eax
  loc_00E10906: lea eax, var_24
  loc_00E10909: push eax
  loc_00E1090A: call edi
  loc_00E1090C: mov ebx, eax
  loc_00E1090E: push FFFFFFFFh
  loc_00E10910: push ebx
  loc_00E10911: mov ecx, [ebx]
  loc_00E10913: call [ecx+000000E4h]
  loc_00E10919: test eax, eax
  loc_00E1091B: fnclex
  loc_00E1091D: jge 00E10931h
  loc_00E1091F: push 000000E4h
  loc_00E10924: push 006E03D4h
  loc_00E10929: push ebx
  loc_00E1092A: push eax
  loc_00E1092B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10931: lea ecx, var_24
  loc_00E10934: call [00401254h] ; __vbaFreeObj
  loc_00E1093A: mov edx, [esi]
  loc_00E1093C: push esi
  loc_00E1093D: call [edx+0000037Ch]
  loc_00E10943: push eax
  loc_00E10944: lea eax, var_24
  loc_00E10947: push eax
  loc_00E10948: call edi
  loc_00E1094A: mov ebx, eax
  loc_00E1094C: push FFFFFFFFh
  loc_00E1094E: push ebx
  loc_00E1094F: mov ecx, [ebx]
  loc_00E10951: call [ecx+0000009Ch]
  loc_00E10957: test eax, eax
  loc_00E10959: fnclex
  loc_00E1095B: jge 00E1096Fh
  loc_00E1095D: push 0000009Ch
  loc_00E10962: push 006E03D4h
  loc_00E10967: push ebx
  loc_00E10968: push eax
  loc_00E10969: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1096F: lea ecx, var_24
  loc_00E10972: call [00401254h] ; __vbaFreeObj
  loc_00E10978: mov edx, [esi]
  loc_00E1097A: push esi
  loc_00E1097B: call [edx+00000380h]
  loc_00E10981: push eax
  loc_00E10982: lea eax, var_24
  loc_00E10985: push eax
  loc_00E10986: call edi
  loc_00E10988: mov ebx, eax
  loc_00E1098A: push FFFFFFFFh
  loc_00E1098C: push ebx
  loc_00E1098D: mov ecx, [ebx]
  loc_00E1098F: call [ecx+0000009Ch]
  loc_00E10995: test eax, eax
  loc_00E10997: fnclex
  loc_00E10999: jge 00E109ADh
  loc_00E1099B: push 0000009Ch
  loc_00E109A0: push 006E03D4h
  loc_00E109A5: push ebx
  loc_00E109A6: push eax
  loc_00E109A7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E109AD: lea ecx, var_24
  loc_00E109B0: call [00401254h] ; __vbaFreeObj
  loc_00E109B6: jmp 00E10A45h
  loc_00E109BB: mov ebx, [004011F0h] ; __vbaVarDup
  loc_00E109C1: mov ecx, 0000000Ah
  loc_00E109C6: mov eax, 80020004h
  loc_00E109CB: mov var_74, ecx
  loc_00E109CE: mov var_64, ecx
  loc_00E109D1: lea edx, var_94
  loc_00E109D7: lea ecx, var_54
  loc_00E109DA: mov var_6C, eax
  loc_00E109DD: mov var_5C, eax
  loc_00E109E0: mov var_8C, 006E07B8h ; "ERROR 001"
  loc_00E109EA: mov var_94, 00000008h
  loc_00E109F4: call ebx
  loc_00E109F6: lea edx, var_84
  loc_00E109FC: lea ecx, var_44
  loc_00E109FF: mov var_7C, 006E0790h ; "Ada Yang Error !"
  loc_00E10A06: mov var_84, 00000008h
  loc_00E10A10: call ebx
  loc_00E10A12: lea edx, var_74
  loc_00E10A15: lea eax, var_64
  loc_00E10A18: push edx
  loc_00E10A19: lea ecx, var_54
  loc_00E10A1C: push eax
  loc_00E10A1D: push ecx
  loc_00E10A1E: lea edx, var_44
  loc_00E10A21: push 00000040h
  loc_00E10A23: push edx
  loc_00E10A24: call [004010A8h] ; rtcMsgBox
  loc_00E10A2A: lea eax, var_74
  loc_00E10A2D: lea ecx, var_64
  loc_00E10A30: push eax
  loc_00E10A31: lea edx, var_54
  loc_00E10A34: push ecx
  loc_00E10A35: lea eax, var_44
  loc_00E10A38: push edx
  loc_00E10A39: push eax
  loc_00E10A3A: push 00000004h
  loc_00E10A3C: call [00401038h] ; __vbaFreeVarList
  loc_00E10A42: add esp, 00000014h
  loc_00E10A45: mov ecx, [esi]
  loc_00E10A47: push esi
  loc_00E10A48: call [ecx+00000390h]
  loc_00E10A4E: lea edx, var_34
  loc_00E10A51: push eax
  loc_00E10A52: push edx
  loc_00E10A53: call edi
  loc_00E10A55: push 006DCBD8h
  loc_00E10A5A: mov var_D4, eax
  loc_00E10A60: mov eax, [esi]
  loc_00E10A62: push 00000000h
  loc_00E10A64: push 00000018h
  loc_00E10A66: push esi
  loc_00E10A67: call [eax+00000490h]
  loc_00E10A6D: lea ecx, var_24
  loc_00E10A70: push eax
  loc_00E10A71: push ecx
  loc_00E10A72: call edi
  loc_00E10A74: lea edx, var_44
  loc_00E10A77: push eax
  loc_00E10A78: push edx
  loc_00E10A79: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E10A7F: add esp, 00000010h
  loc_00E10A82: push eax
  loc_00E10A83: call [00401120h] ; __vbaCastObjVar
  loc_00E10A89: push eax
  loc_00E10A8A: lea eax, var_28
  loc_00E10A8D: push eax
  loc_00E10A8E: call edi
  loc_00E10A90: mov ebx, eax
  loc_00E10A92: lea edx, var_2C
  loc_00E10A95: push edx
  loc_00E10A96: push ebx
  loc_00E10A97: mov ecx, [ebx]
  loc_00E10A99: call [ecx+00000054h]
  loc_00E10A9C: test eax, eax
  loc_00E10A9E: fnclex
  loc_00E10AA0: jge 00E10AB1h
  loc_00E10AA2: push 00000054h
  loc_00E10AA4: push 006DCBE8h
  loc_00E10AA9: push ebx
  loc_00E10AAA: push eax
  loc_00E10AAB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10AB1: lea ebx, var_30
  loc_00E10AB4: mov eax, var_2C
  loc_00E10AB7: push ebx
  loc_00E10AB8: mov ecx, 00000002h
  loc_00E10ABD: sub esp, 00000010h
  loc_00E10AC0: mov var_84, ecx
  loc_00E10AC6: mov ebx, esp
  loc_00E10AC8: mov var_7C, 00000003h
  loc_00E10ACF: mov edx, [eax]
  loc_00E10AD1: push eax
  loc_00E10AD2: mov [ebx], ecx
  loc_00E10AD4: mov ecx, var_80
  loc_00E10AD7: mov var_C4, eax
  loc_00E10ADD: mov [ebx+00000004h], ecx
  loc_00E10AE0: mov ecx, var_7C
  loc_00E10AE3: mov [ebx+00000008h], ecx
  loc_00E10AE6: mov ecx, var_78
  loc_00E10AE9: mov [ebx+0000000Ch], ecx
  loc_00E10AEC: call [edx+00000028h]
  loc_00E10AEF: test eax, eax
  loc_00E10AF1: fnclex
  loc_00E10AF3: jge 00E10B0Ah
  loc_00E10AF5: mov edx, var_C4
  loc_00E10AFB: push 00000028h
  loc_00E10AFD: push 006E09E8h
  loc_00E10B02: push edx
  loc_00E10B03: push eax
  loc_00E10B04: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10B0A: mov eax, var_30
  loc_00E10B0D: lea edx, var_54
  loc_00E10B10: push edx
  loc_00E10B11: push eax
  loc_00E10B12: mov ecx, [eax]
  loc_00E10B14: mov ebx, eax
  loc_00E10B16: call [ecx+00000034h]
  loc_00E10B19: test eax, eax
  loc_00E10B1B: fnclex
  loc_00E10B1D: jge 00E10B2Eh
  loc_00E10B1F: push 00000034h
  loc_00E10B21: push 006E09F8h
  loc_00E10B26: push ebx
  loc_00E10B27: push eax
  loc_00E10B28: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10B2E: mov eax, var_D4
  loc_00E10B34: lea ecx, var_54
  loc_00E10B37: push ecx
  loc_00E10B38: mov ebx, [eax]
  loc_00E10B3A: call [00401034h] ; __vbaStrVarMove
  loc_00E10B40: mov edx, eax
  loc_00E10B42: lea ecx, var_18
  loc_00E10B45: call [00401228h] ; __vbaStrMove
  loc_00E10B4B: mov edx, ebx
  loc_00E10B4D: mov ebx, var_D4
  loc_00E10B53: push eax
  loc_00E10B54: push ebx
  loc_00E10B55: call [edx+000000A4h]
  loc_00E10B5B: test eax, eax
  loc_00E10B5D: fnclex
  loc_00E10B5F: jge 00E10B73h
  loc_00E10B61: push 000000A4h
  loc_00E10B66: push 006DCB70h
  loc_00E10B6B: push ebx
  loc_00E10B6C: push eax
  loc_00E10B6D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10B73: lea ecx, var_18
  loc_00E10B76: call [00401258h] ; __vbaFreeStr
  loc_00E10B7C: lea eax, var_34
  loc_00E10B7F: lea ecx, var_30
  loc_00E10B82: push eax
  loc_00E10B83: lea edx, var_2C
  loc_00E10B86: push ecx
  loc_00E10B87: lea eax, var_28
  loc_00E10B8A: push edx
  loc_00E10B8B: lea ecx, var_24
  loc_00E10B8E: push eax
  loc_00E10B8F: push ecx
  loc_00E10B90: push 00000005h
  loc_00E10B92: call [00401048h] ; __vbaFreeObjList
  loc_00E10B98: lea edx, var_54
  loc_00E10B9B: lea eax, var_44
  loc_00E10B9E: push edx
  loc_00E10B9F: push eax
  loc_00E10BA0: push 00000002h
  loc_00E10BA2: call [00401038h] ; __vbaFreeVarList
  loc_00E10BA8: mov ecx, [esi]
  loc_00E10BAA: add esp, 00000024h
  loc_00E10BAD: push esi
  loc_00E10BAE: call [ecx+0000038Ch]
  loc_00E10BB4: lea edx, var_34
  loc_00E10BB7: push eax
  loc_00E10BB8: push edx
  loc_00E10BB9: call edi
  loc_00E10BBB: push 006DCBD8h
  loc_00E10BC0: mov var_D4, eax
  loc_00E10BC6: mov eax, [esi]
  loc_00E10BC8: push 00000000h
  loc_00E10BCA: push 00000018h
  loc_00E10BCC: push esi
  loc_00E10BCD: call [eax+00000490h]
  loc_00E10BD3: lea ecx, var_24
  loc_00E10BD6: push eax
  loc_00E10BD7: push ecx
  loc_00E10BD8: call edi
  loc_00E10BDA: lea edx, var_44
  loc_00E10BDD: push eax
  loc_00E10BDE: push edx
  loc_00E10BDF: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E10BE5: add esp, 00000010h
  loc_00E10BE8: push eax
  loc_00E10BE9: call [00401120h] ; __vbaCastObjVar
  loc_00E10BEF: push eax
  loc_00E10BF0: lea eax, var_28
  loc_00E10BF3: push eax
  loc_00E10BF4: call edi
  loc_00E10BF6: mov ebx, eax
  loc_00E10BF8: lea edx, var_2C
  loc_00E10BFB: push edx
  loc_00E10BFC: push ebx
  loc_00E10BFD: mov ecx, [ebx]
  loc_00E10BFF: call [ecx+00000054h]
  loc_00E10C02: test eax, eax
  loc_00E10C04: fnclex
  loc_00E10C06: jge 00E10C17h
  loc_00E10C08: push 00000054h
  loc_00E10C0A: push 006DCBE8h
  loc_00E10C0F: push ebx
  loc_00E10C10: push eax
  loc_00E10C11: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10C17: lea ebx, var_30
  loc_00E10C1A: mov eax, var_2C
  loc_00E10C1D: push ebx
  loc_00E10C1E: mov ecx, 00000002h
  loc_00E10C23: sub esp, 00000010h
  loc_00E10C26: mov var_84, ecx
  loc_00E10C2C: mov ebx, esp
  loc_00E10C2E: mov var_7C, 00000004h
  loc_00E10C35: mov edx, [eax]
  loc_00E10C37: push eax
  loc_00E10C38: mov [ebx], ecx
  loc_00E10C3A: mov ecx, var_80
  loc_00E10C3D: mov var_C4, eax
  loc_00E10C43: mov [ebx+00000004h], ecx
  loc_00E10C46: mov ecx, var_7C
  loc_00E10C49: mov [ebx+00000008h], ecx
  loc_00E10C4C: mov ecx, var_78
  loc_00E10C4F: mov [ebx+0000000Ch], ecx
  loc_00E10C52: call [edx+00000028h]
  loc_00E10C55: test eax, eax
  loc_00E10C57: fnclex
  loc_00E10C59: jge 00E10C70h
  loc_00E10C5B: mov edx, var_C4
  loc_00E10C61: push 00000028h
  loc_00E10C63: push 006E09E8h
  loc_00E10C68: push edx
  loc_00E10C69: push eax
  loc_00E10C6A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10C70: mov eax, var_30
  loc_00E10C73: lea edx, var_54
  loc_00E10C76: push edx
  loc_00E10C77: push eax
  loc_00E10C78: mov ecx, [eax]
  loc_00E10C7A: mov ebx, eax
  loc_00E10C7C: call [ecx+00000034h]
  loc_00E10C7F: test eax, eax
  loc_00E10C81: fnclex
  loc_00E10C83: jge 00E10C94h
  loc_00E10C85: push 00000034h
  loc_00E10C87: push 006E09F8h
  loc_00E10C8C: push ebx
  loc_00E10C8D: push eax
  loc_00E10C8E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10C94: mov eax, var_D4
  loc_00E10C9A: lea ecx, var_54
  loc_00E10C9D: push ecx
  loc_00E10C9E: mov ebx, [eax]
  loc_00E10CA0: call [00401034h] ; __vbaStrVarMove
  loc_00E10CA6: mov edx, eax
  loc_00E10CA8: lea ecx, var_18
  loc_00E10CAB: call [00401228h] ; __vbaStrMove
  loc_00E10CB1: mov edx, ebx
  loc_00E10CB3: mov ebx, var_D4
  loc_00E10CB9: push eax
  loc_00E10CBA: push ebx
  loc_00E10CBB: call [edx+000000A4h]
  loc_00E10CC1: test eax, eax
  loc_00E10CC3: fnclex
  loc_00E10CC5: jge 00E10CD9h
  loc_00E10CC7: push 000000A4h
  loc_00E10CCC: push 006DCB70h
  loc_00E10CD1: push ebx
  loc_00E10CD2: push eax
  loc_00E10CD3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10CD9: lea ecx, var_18
  loc_00E10CDC: call [00401258h] ; __vbaFreeStr
  loc_00E10CE2: lea eax, var_34
  loc_00E10CE5: lea ecx, var_30
  loc_00E10CE8: push eax
  loc_00E10CE9: lea edx, var_2C
  loc_00E10CEC: push ecx
  loc_00E10CED: lea eax, var_28
  loc_00E10CF0: push edx
  loc_00E10CF1: lea ecx, var_24
  loc_00E10CF4: push eax
  loc_00E10CF5: push ecx
  loc_00E10CF6: push 00000005h
  loc_00E10CF8: call [00401048h] ; __vbaFreeObjList
  loc_00E10CFE: lea edx, var_54
  loc_00E10D01: lea eax, var_44
  loc_00E10D04: push edx
  loc_00E10D05: push eax
  loc_00E10D06: push 00000002h
  loc_00E10D08: call [00401038h] ; __vbaFreeVarList
  loc_00E10D0E: mov ecx, [esi]
  loc_00E10D10: add esp, 00000024h
  loc_00E10D13: push esi
  loc_00E10D14: call [ecx+00000388h]
  loc_00E10D1A: lea edx, var_34
  loc_00E10D1D: push eax
  loc_00E10D1E: push edx
  loc_00E10D1F: call edi
  loc_00E10D21: push 006DCBD8h
  loc_00E10D26: mov var_D4, eax
  loc_00E10D2C: mov eax, [esi]
  loc_00E10D2E: push 00000000h
  loc_00E10D30: push 00000018h
  loc_00E10D32: push esi
  loc_00E10D33: call [eax+00000490h]
  loc_00E10D39: lea ecx, var_24
  loc_00E10D3C: push eax
  loc_00E10D3D: push ecx
  loc_00E10D3E: call edi
  loc_00E10D40: lea edx, var_44
  loc_00E10D43: push eax
  loc_00E10D44: push edx
  loc_00E10D45: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E10D4B: add esp, 00000010h
  loc_00E10D4E: push eax
  loc_00E10D4F: call [00401120h] ; __vbaCastObjVar
  loc_00E10D55: push eax
  loc_00E10D56: lea eax, var_28
  loc_00E10D59: push eax
  loc_00E10D5A: call edi
  loc_00E10D5C: mov ebx, eax
  loc_00E10D5E: lea edx, var_2C
  loc_00E10D61: push edx
  loc_00E10D62: push ebx
  loc_00E10D63: mov ecx, [ebx]
  loc_00E10D65: call [ecx+00000054h]
  loc_00E10D68: test eax, eax
  loc_00E10D6A: fnclex
  loc_00E10D6C: jge 00E10D7Dh
  loc_00E10D6E: push 00000054h
  loc_00E10D70: push 006DCBE8h
  loc_00E10D75: push ebx
  loc_00E10D76: push eax
  loc_00E10D77: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10D7D: lea ebx, var_30
  loc_00E10D80: mov eax, var_2C
  loc_00E10D83: push ebx
  loc_00E10D84: mov ecx, 00000002h
  loc_00E10D89: sub esp, 00000010h
  loc_00E10D8C: mov var_84, ecx
  loc_00E10D92: mov ebx, esp
  loc_00E10D94: mov var_7C, 00000005h
  loc_00E10D9B: mov edx, [eax]
  loc_00E10D9D: push eax
  loc_00E10D9E: mov [ebx], ecx
  loc_00E10DA0: mov ecx, var_80
  loc_00E10DA3: mov var_C4, eax
  loc_00E10DA9: mov [ebx+00000004h], ecx
  loc_00E10DAC: mov ecx, var_7C
  loc_00E10DAF: mov [ebx+00000008h], ecx
  loc_00E10DB2: mov ecx, var_78
  loc_00E10DB5: mov [ebx+0000000Ch], ecx
  loc_00E10DB8: call [edx+00000028h]
  loc_00E10DBB: test eax, eax
  loc_00E10DBD: fnclex
  loc_00E10DBF: jge 00E10DD6h
  loc_00E10DC1: mov edx, var_C4
  loc_00E10DC7: push 00000028h
  loc_00E10DC9: push 006E09E8h
  loc_00E10DCE: push edx
  loc_00E10DCF: push eax
  loc_00E10DD0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10DD6: mov eax, var_30
  loc_00E10DD9: lea edx, var_54
  loc_00E10DDC: push edx
  loc_00E10DDD: push eax
  loc_00E10DDE: mov ecx, [eax]
  loc_00E10DE0: mov ebx, eax
  loc_00E10DE2: call [ecx+00000034h]
  loc_00E10DE5: test eax, eax
  loc_00E10DE7: fnclex
  loc_00E10DE9: jge 00E10DFAh
  loc_00E10DEB: push 00000034h
  loc_00E10DED: push 006E09F8h
  loc_00E10DF2: push ebx
  loc_00E10DF3: push eax
  loc_00E10DF4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10DFA: mov eax, var_D4
  loc_00E10E00: lea ecx, var_54
  loc_00E10E03: push ecx
  loc_00E10E04: mov ebx, [eax]
  loc_00E10E06: call [00401034h] ; __vbaStrVarMove
  loc_00E10E0C: mov edx, eax
  loc_00E10E0E: lea ecx, var_18
  loc_00E10E11: call [00401228h] ; __vbaStrMove
  loc_00E10E17: mov edx, ebx
  loc_00E10E19: mov ebx, var_D4
  loc_00E10E1F: push eax
  loc_00E10E20: push ebx
  loc_00E10E21: call [edx+000000A4h]
  loc_00E10E27: test eax, eax
  loc_00E10E29: fnclex
  loc_00E10E2B: jge 00E10E3Fh
  loc_00E10E2D: push 000000A4h
  loc_00E10E32: push 006DCB70h
  loc_00E10E37: push ebx
  loc_00E10E38: push eax
  loc_00E10E39: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10E3F: lea ecx, var_18
  loc_00E10E42: call [00401258h] ; __vbaFreeStr
  loc_00E10E48: lea eax, var_34
  loc_00E10E4B: lea ecx, var_30
  loc_00E10E4E: push eax
  loc_00E10E4F: lea edx, var_2C
  loc_00E10E52: push ecx
  loc_00E10E53: lea eax, var_28
  loc_00E10E56: push edx
  loc_00E10E57: lea ecx, var_24
  loc_00E10E5A: push eax
  loc_00E10E5B: push ecx
  loc_00E10E5C: push 00000005h
  loc_00E10E5E: call [00401048h] ; __vbaFreeObjList
  loc_00E10E64: lea edx, var_54
  loc_00E10E67: lea eax, var_44
  loc_00E10E6A: push edx
  loc_00E10E6B: push eax
  loc_00E10E6C: push 00000002h
  loc_00E10E6E: call [00401038h] ; __vbaFreeVarList
  loc_00E10E74: mov ecx, [esi]
  loc_00E10E76: add esp, 00000024h
  loc_00E10E79: push 006DCBD8h
  loc_00E10E7E: push 00000000h
  loc_00E10E80: push 00000018h
  loc_00E10E82: push esi
  loc_00E10E83: call [ecx+00000490h]
  loc_00E10E89: lea edx, var_24
  loc_00E10E8C: push eax
  loc_00E10E8D: push edx
  loc_00E10E8E: call edi
  loc_00E10E90: push eax
  loc_00E10E91: lea eax, var_44
  loc_00E10E94: push eax
  loc_00E10E95: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E10E9B: add esp, 00000010h
  loc_00E10E9E: push eax
  loc_00E10E9F: call [00401120h] ; __vbaCastObjVar
  loc_00E10EA5: lea ecx, var_28
  loc_00E10EA8: push eax
  loc_00E10EA9: push ecx
  loc_00E10EAA: call edi
  loc_00E10EAC: mov ebx, eax
  loc_00E10EAE: lea eax, var_2C
  loc_00E10EB1: push eax
  loc_00E10EB2: push ebx
  loc_00E10EB3: mov edx, [ebx]
  loc_00E10EB5: call [edx+00000054h]
  loc_00E10EB8: test eax, eax
  loc_00E10EBA: fnclex
  loc_00E10EBC: jge 00E10ECDh
  loc_00E10EBE: push 00000054h
  loc_00E10EC0: push 006DCBE8h
  loc_00E10EC5: push ebx
  loc_00E10EC6: push eax
  loc_00E10EC7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10ECD: lea ebx, var_30
  loc_00E10ED0: mov eax, var_2C
  loc_00E10ED3: push ebx
  loc_00E10ED4: mov ecx, 00000002h
  loc_00E10ED9: sub esp, 00000010h
  loc_00E10EDC: mov var_84, ecx
  loc_00E10EE2: mov ebx, esp
  loc_00E10EE4: mov var_7C, 00000006h
  loc_00E10EEB: mov edx, [eax]
  loc_00E10EED: push eax
  loc_00E10EEE: mov [ebx], ecx
  loc_00E10EF0: mov ecx, var_80
  loc_00E10EF3: mov var_C4, eax
  loc_00E10EF9: mov [ebx+00000004h], ecx
  loc_00E10EFC: mov ecx, var_7C
  loc_00E10EFF: mov [ebx+00000008h], ecx
  loc_00E10F02: mov ecx, var_78
  loc_00E10F05: mov [ebx+0000000Ch], ecx
  loc_00E10F08: call [edx+00000028h]
  loc_00E10F0B: test eax, eax
  loc_00E10F0D: fnclex
  loc_00E10F0F: jge 00E10F26h
  loc_00E10F11: mov edx, var_C4
  loc_00E10F17: push 00000028h
  loc_00E10F19: push 006E09E8h
  loc_00E10F1E: push edx
  loc_00E10F1F: push eax
  loc_00E10F20: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E10F26: sub esp, 00000010h
  loc_00E10F29: mov eax, var_30
  loc_00E10F2C: mov edx, esp
  loc_00E10F2E: mov ecx, 00000009h
  loc_00E10F33: mov var_54, ecx
  loc_00E10F36: mov var_4C, eax
  loc_00E10F39: mov [edx], ecx
  loc_00E10F3B: mov ecx, var_50
  loc_00E10F3E: push 00000014h
  loc_00E10F40: push esi
  loc_00E10F41: mov [edx+00000004h], ecx
  loc_00E10F44: mov ecx, [esi]
  loc_00E10F46: mov var_30, 00000000h
  loc_00E10F4D: mov [edx+00000008h], eax
  loc_00E10F50: mov eax, var_48
  loc_00E10F53: mov [edx+0000000Ch], eax
  loc_00E10F56: call [ecx+0000048Ch]
  loc_00E10F5C: lea edx, var_34
  loc_00E10F5F: push eax
  loc_00E10F60: push edx
  loc_00E10F61: call edi
  loc_00E10F63: push eax
  loc_00E10F64: call [00401238h] ; __vbaLateIdSt
  loc_00E10F6A: lea eax, var_34
  loc_00E10F6D: lea ecx, var_2C
  loc_00E10F70: push eax
  loc_00E10F71: lea edx, var_28
  loc_00E10F74: push ecx
  loc_00E10F75: lea eax, var_24
  loc_00E10F78: push edx
  loc_00E10F79: push eax
  loc_00E10F7A: push 00000004h
  loc_00E10F7C: call [00401048h] ; __vbaFreeObjList
  loc_00E10F82: lea ecx, var_54
  loc_00E10F85: lea edx, var_44
  loc_00E10F88: push ecx
  loc_00E10F89: push edx
  loc_00E10F8A: push 00000002h
  loc_00E10F8C: call [00401038h] ; __vbaFreeVarList
  loc_00E10F92: mov eax, [esi]
  loc_00E10F94: add esp, 00000020h
  loc_00E10F97: push esi
  loc_00E10F98: call [eax+00000378h]
  loc_00E10F9E: lea ecx, var_34
  loc_00E10FA1: push eax
  loc_00E10FA2: push ecx
  loc_00E10FA3: call edi
  loc_00E10FA5: mov edx, [esi]
  loc_00E10FA7: push 006DCBD8h
  loc_00E10FAC: push 00000000h
  loc_00E10FAE: push 00000018h
  loc_00E10FB0: push esi
  loc_00E10FB1: mov var_D4, eax
  loc_00E10FB7: call [edx+00000490h]
  loc_00E10FBD: push eax
  loc_00E10FBE: lea eax, var_24
  loc_00E10FC1: push eax
  loc_00E10FC2: call edi
  loc_00E10FC4: lea ecx, var_44
  loc_00E10FC7: push eax
  loc_00E10FC8: push ecx
  loc_00E10FC9: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E10FCF: add esp, 00000010h
  loc_00E10FD2: push eax
  loc_00E10FD3: call [00401120h] ; __vbaCastObjVar
  loc_00E10FD9: lea edx, var_28
  loc_00E10FDC: push eax
  loc_00E10FDD: push edx
  loc_00E10FDE: call edi
  loc_00E10FE0: mov ebx, eax
  loc_00E10FE2: lea ecx, var_2C
  loc_00E10FE5: push ecx
  loc_00E10FE6: push ebx
  loc_00E10FE7: mov eax, [ebx]
  loc_00E10FE9: call [eax+00000054h]
  loc_00E10FEC: test eax, eax
  loc_00E10FEE: fnclex
  loc_00E10FF0: jge 00E11001h
  loc_00E10FF2: push 00000054h
  loc_00E10FF4: push 006DCBE8h
  loc_00E10FF9: push ebx
  loc_00E10FFA: push eax
  loc_00E10FFB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11001: lea ebx, var_30
  loc_00E11004: mov eax, var_2C
  loc_00E11007: push ebx
  loc_00E11008: mov ecx, 00000002h
  loc_00E1100D: sub esp, 00000010h
  loc_00E11010: mov var_84, ecx
  loc_00E11016: mov ebx, esp
  loc_00E11018: mov var_7C, 00000007h
  loc_00E1101F: mov edx, [eax]
  loc_00E11021: push eax
  loc_00E11022: mov [ebx], ecx
  loc_00E11024: mov ecx, var_80
  loc_00E11027: mov var_C4, eax
  loc_00E1102D: mov [ebx+00000004h], ecx
  loc_00E11030: mov ecx, var_7C
  loc_00E11033: mov [ebx+00000008h], ecx
  loc_00E11036: mov ecx, var_78
  loc_00E11039: mov [ebx+0000000Ch], ecx
  loc_00E1103C: call [edx+00000028h]
  loc_00E1103F: test eax, eax
  loc_00E11041: fnclex
  loc_00E11043: jge 00E1105Ah
  loc_00E11045: mov edx, var_C4
  loc_00E1104B: push 00000028h
  loc_00E1104D: push 006E09E8h
  loc_00E11052: push edx
  loc_00E11053: push eax
  loc_00E11054: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1105A: mov eax, var_30
  loc_00E1105D: lea edx, var_54
  loc_00E11060: push edx
  loc_00E11061: push eax
  loc_00E11062: mov ecx, [eax]
  loc_00E11064: mov ebx, eax
  loc_00E11066: call [ecx+00000034h]
  loc_00E11069: test eax, eax
  loc_00E1106B: fnclex
  loc_00E1106D: jge 00E1107Eh
  loc_00E1106F: push 00000034h
  loc_00E11071: push 006E09F8h
  loc_00E11076: push ebx
  loc_00E11077: push eax
  loc_00E11078: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1107E: mov eax, var_D4
  loc_00E11084: lea ecx, var_54
  loc_00E11087: push ecx
  loc_00E11088: mov ebx, [eax]
  loc_00E1108A: call [00401034h] ; __vbaStrVarMove
  loc_00E11090: mov edx, eax
  loc_00E11092: lea ecx, var_18
  loc_00E11095: call [00401228h] ; __vbaStrMove
  loc_00E1109B: mov edx, ebx
  loc_00E1109D: mov ebx, var_D4
  loc_00E110A3: push eax
  loc_00E110A4: push ebx
  loc_00E110A5: call [edx+000000ACh]
  loc_00E110AB: test eax, eax
  loc_00E110AD: fnclex
  loc_00E110AF: jge 00E110C3h
  loc_00E110B1: push 000000ACh
  loc_00E110B6: push 006E0400h
  loc_00E110BB: push ebx
  loc_00E110BC: push eax
  loc_00E110BD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E110C3: lea ecx, var_18
  loc_00E110C6: call [00401258h] ; __vbaFreeStr
  loc_00E110CC: lea eax, var_34
  loc_00E110CF: lea ecx, var_30
  loc_00E110D2: push eax
  loc_00E110D3: lea edx, var_2C
  loc_00E110D6: push ecx
  loc_00E110D7: lea eax, var_28
  loc_00E110DA: push edx
  loc_00E110DB: lea ecx, var_24
  loc_00E110DE: push eax
  loc_00E110DF: push ecx
  loc_00E110E0: push 00000005h
  loc_00E110E2: call [00401048h] ; __vbaFreeObjList
  loc_00E110E8: lea edx, var_54
  loc_00E110EB: lea eax, var_44
  loc_00E110EE: push edx
  loc_00E110EF: push eax
  loc_00E110F0: push 00000002h
  loc_00E110F2: call [00401038h] ; __vbaFreeVarList
  loc_00E110F8: mov ecx, [esi]
  loc_00E110FA: add esp, 00000024h
  loc_00E110FD: push esi
  loc_00E110FE: call [ecx+00000384h]
  loc_00E11104: lea edx, var_34
  loc_00E11107: push eax
  loc_00E11108: push edx
  loc_00E11109: call edi
  loc_00E1110B: push 006DCBD8h
  loc_00E11110: mov var_D4, eax
  loc_00E11116: mov eax, [esi]
  loc_00E11118: push 00000000h
  loc_00E1111A: push 00000018h
  loc_00E1111C: push esi
  loc_00E1111D: call [eax+00000490h]
  loc_00E11123: lea ecx, var_24
  loc_00E11126: push eax
  loc_00E11127: push ecx
  loc_00E11128: call edi
  loc_00E1112A: lea edx, var_44
  loc_00E1112D: push eax
  loc_00E1112E: push edx
  loc_00E1112F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11135: add esp, 00000010h
  loc_00E11138: push eax
  loc_00E11139: call [00401120h] ; __vbaCastObjVar
  loc_00E1113F: push eax
  loc_00E11140: lea eax, var_28
  loc_00E11143: push eax
  loc_00E11144: call edi
  loc_00E11146: mov ebx, eax
  loc_00E11148: lea edx, var_2C
  loc_00E1114B: push edx
  loc_00E1114C: push ebx
  loc_00E1114D: mov ecx, [ebx]
  loc_00E1114F: call [ecx+00000054h]
  loc_00E11152: test eax, eax
  loc_00E11154: fnclex
  loc_00E11156: jge 00E11167h
  loc_00E11158: push 00000054h
  loc_00E1115A: push 006DCBE8h
  loc_00E1115F: push ebx
  loc_00E11160: push eax
  loc_00E11161: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11167: lea ebx, var_30
  loc_00E1116A: mov eax, var_2C
  loc_00E1116D: push ebx
  loc_00E1116E: mov ecx, 00000002h
  loc_00E11173: sub esp, 00000010h
  loc_00E11176: mov var_84, ecx
  loc_00E1117C: mov ebx, esp
  loc_00E1117E: mov var_7C, 00000008h
  loc_00E11185: mov edx, [eax]
  loc_00E11187: push eax
  loc_00E11188: mov [ebx], ecx
  loc_00E1118A: mov ecx, var_80
  loc_00E1118D: mov var_C4, eax
  loc_00E11193: mov [ebx+00000004h], ecx
  loc_00E11196: mov ecx, var_7C
  loc_00E11199: mov [ebx+00000008h], ecx
  loc_00E1119C: mov ecx, var_78
  loc_00E1119F: mov [ebx+0000000Ch], ecx
  loc_00E111A2: call [edx+00000028h]
  loc_00E111A5: test eax, eax
  loc_00E111A7: fnclex
  loc_00E111A9: jge 00E111C0h
  loc_00E111AB: mov edx, var_C4
  loc_00E111B1: push 00000028h
  loc_00E111B3: push 006E09E8h
  loc_00E111B8: push edx
  loc_00E111B9: push eax
  loc_00E111BA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E111C0: mov eax, var_30
  loc_00E111C3: lea edx, var_54
  loc_00E111C6: push edx
  loc_00E111C7: push eax
  loc_00E111C8: mov ecx, [eax]
  loc_00E111CA: mov ebx, eax
  loc_00E111CC: call [ecx+00000034h]
  loc_00E111CF: test eax, eax
  loc_00E111D1: fnclex
  loc_00E111D3: jge 00E111E4h
  loc_00E111D5: push 00000034h
  loc_00E111D7: push 006E09F8h
  loc_00E111DC: push ebx
  loc_00E111DD: push eax
  loc_00E111DE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E111E4: mov eax, var_D4
  loc_00E111EA: lea ecx, var_54
  loc_00E111ED: push ecx
  loc_00E111EE: mov ebx, [eax]
  loc_00E111F0: call [00401034h] ; __vbaStrVarMove
  loc_00E111F6: mov edx, eax
  loc_00E111F8: lea ecx, var_18
  loc_00E111FB: call [00401228h] ; __vbaStrMove
  loc_00E11201: mov edx, ebx
  loc_00E11203: mov ebx, var_D4
  loc_00E11209: push eax
  loc_00E1120A: push ebx
  loc_00E1120B: call [edx+000000A4h]
  loc_00E11211: test eax, eax
  loc_00E11213: fnclex
  loc_00E11215: jge 00E11229h
  loc_00E11217: push 000000A4h
  loc_00E1121C: push 006DCB70h
  loc_00E11221: push ebx
  loc_00E11222: push eax
  loc_00E11223: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11229: lea ecx, var_18
  loc_00E1122C: call [00401258h] ; __vbaFreeStr
  loc_00E11232: lea eax, var_34
  loc_00E11235: lea ecx, var_30
  loc_00E11238: push eax
  loc_00E11239: lea edx, var_2C
  loc_00E1123C: push ecx
  loc_00E1123D: lea eax, var_28
  loc_00E11240: push edx
  loc_00E11241: lea ecx, var_24
  loc_00E11244: push eax
  loc_00E11245: push ecx
  loc_00E11246: push 00000005h
  loc_00E11248: call [00401048h] ; __vbaFreeObjList
  loc_00E1124E: lea edx, var_54
  loc_00E11251: lea eax, var_44
  loc_00E11254: push edx
  loc_00E11255: push eax
  loc_00E11256: push 00000002h
  loc_00E11258: call [00401038h] ; __vbaFreeVarList
  loc_00E1125E: mov ecx, [esi]
  loc_00E11260: add esp, 00000024h
  loc_00E11263: push esi
  loc_00E11264: call [ecx+00000374h]
  loc_00E1126A: lea edx, var_34
  loc_00E1126D: push eax
  loc_00E1126E: push edx
  loc_00E1126F: call edi
  loc_00E11271: push 006DCBD8h
  loc_00E11276: mov var_D4, eax
  loc_00E1127C: mov eax, [esi]
  loc_00E1127E: push 00000000h
  loc_00E11280: push 00000018h
  loc_00E11282: push esi
  loc_00E11283: call [eax+00000490h]
  loc_00E11289: lea ecx, var_24
  loc_00E1128C: push eax
  loc_00E1128D: push ecx
  loc_00E1128E: call edi
  loc_00E11290: lea edx, var_44
  loc_00E11293: push eax
  loc_00E11294: push edx
  loc_00E11295: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1129B: add esp, 00000010h
  loc_00E1129E: push eax
  loc_00E1129F: call [00401120h] ; __vbaCastObjVar
  loc_00E112A5: push eax
  loc_00E112A6: lea eax, var_28
  loc_00E112A9: push eax
  loc_00E112AA: call edi
  loc_00E112AC: mov ebx, eax
  loc_00E112AE: lea edx, var_2C
  loc_00E112B1: push edx
  loc_00E112B2: push ebx
  loc_00E112B3: mov ecx, [ebx]
  loc_00E112B5: call [ecx+00000054h]
  loc_00E112B8: test eax, eax
  loc_00E112BA: fnclex
  loc_00E112BC: jge 00E112CDh
  loc_00E112BE: push 00000054h
  loc_00E112C0: push 006DCBE8h
  loc_00E112C5: push ebx
  loc_00E112C6: push eax
  loc_00E112C7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E112CD: lea ebx, var_30
  loc_00E112D0: mov eax, var_2C
  loc_00E112D3: push ebx
  loc_00E112D4: mov ecx, 00000002h
  loc_00E112D9: sub esp, 00000010h
  loc_00E112DC: mov var_84, ecx
  loc_00E112E2: mov ebx, esp
  loc_00E112E4: mov var_7C, 00000009h
  loc_00E112EB: mov edx, [eax]
  loc_00E112ED: push eax
  loc_00E112EE: mov [ebx], ecx
  loc_00E112F0: mov ecx, var_80
  loc_00E112F3: mov var_C4, eax
  loc_00E112F9: mov [ebx+00000004h], ecx
  loc_00E112FC: mov ecx, var_7C
  loc_00E112FF: mov [ebx+00000008h], ecx
  loc_00E11302: mov ecx, var_78
  loc_00E11305: mov [ebx+0000000Ch], ecx
  loc_00E11308: call [edx+00000028h]
  loc_00E1130B: test eax, eax
  loc_00E1130D: fnclex
  loc_00E1130F: jge 00E11326h
  loc_00E11311: mov edx, var_C4
  loc_00E11317: push 00000028h
  loc_00E11319: push 006E09E8h
  loc_00E1131E: push edx
  loc_00E1131F: push eax
  loc_00E11320: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11326: mov eax, var_30
  loc_00E11329: lea edx, var_54
  loc_00E1132C: push edx
  loc_00E1132D: push eax
  loc_00E1132E: mov ecx, [eax]
  loc_00E11330: mov ebx, eax
  loc_00E11332: call [ecx+00000034h]
  loc_00E11335: test eax, eax
  loc_00E11337: fnclex
  loc_00E11339: jge 00E1134Ah
  loc_00E1133B: push 00000034h
  loc_00E1133D: push 006E09F8h
  loc_00E11342: push ebx
  loc_00E11343: push eax
  loc_00E11344: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1134A: mov eax, var_D4
  loc_00E11350: lea ecx, var_54
  loc_00E11353: push ecx
  loc_00E11354: mov ebx, [eax]
  loc_00E11356: call [00401034h] ; __vbaStrVarMove
  loc_00E1135C: mov edx, eax
  loc_00E1135E: lea ecx, var_18
  loc_00E11361: call [00401228h] ; __vbaStrMove
  loc_00E11367: mov edx, ebx
  loc_00E11369: mov ebx, var_D4
  loc_00E1136F: push eax
  loc_00E11370: push ebx
  loc_00E11371: call [edx+000000A4h]
  loc_00E11377: test eax, eax
  loc_00E11379: fnclex
  loc_00E1137B: jge 00E1138Fh
  loc_00E1137D: push 000000A4h
  loc_00E11382: push 006DCB70h
  loc_00E11387: push ebx
  loc_00E11388: push eax
  loc_00E11389: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1138F: lea ecx, var_18
  loc_00E11392: call [00401258h] ; __vbaFreeStr
  loc_00E11398: lea eax, var_34
  loc_00E1139B: lea ecx, var_30
  loc_00E1139E: push eax
  loc_00E1139F: lea edx, var_2C
  loc_00E113A2: push ecx
  loc_00E113A3: lea eax, var_28
  loc_00E113A6: push edx
  loc_00E113A7: lea ecx, var_24
  loc_00E113AA: push eax
  loc_00E113AB: push ecx
  loc_00E113AC: push 00000005h
  loc_00E113AE: call [00401048h] ; __vbaFreeObjList
  loc_00E113B4: lea edx, var_54
  loc_00E113B7: lea eax, var_44
  loc_00E113BA: push edx
  loc_00E113BB: push eax
  loc_00E113BC: push 00000002h
  loc_00E113BE: call [00401038h] ; __vbaFreeVarList
  loc_00E113C4: mov ecx, [esi]
  loc_00E113C6: add esp, 00000024h
  loc_00E113C9: push esi
  loc_00E113CA: call [ecx+00000370h]
  loc_00E113D0: lea edx, var_34
  loc_00E113D3: push eax
  loc_00E113D4: push edx
  loc_00E113D5: call edi
  loc_00E113D7: push 006DCBD8h
  loc_00E113DC: mov var_D4, eax
  loc_00E113E2: mov eax, [esi]
  loc_00E113E4: push 00000000h
  loc_00E113E6: push 00000018h
  loc_00E113E8: push esi
  loc_00E113E9: call [eax+00000490h]
  loc_00E113EF: lea ecx, var_24
  loc_00E113F2: push eax
  loc_00E113F3: push ecx
  loc_00E113F4: call edi
  loc_00E113F6: lea edx, var_44
  loc_00E113F9: push eax
  loc_00E113FA: push edx
  loc_00E113FB: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11401: add esp, 00000010h
  loc_00E11404: push eax
  loc_00E11405: call [00401120h] ; __vbaCastObjVar
  loc_00E1140B: push eax
  loc_00E1140C: lea eax, var_28
  loc_00E1140F: push eax
  loc_00E11410: call edi
  loc_00E11412: mov ebx, eax
  loc_00E11414: lea edx, var_2C
  loc_00E11417: push edx
  loc_00E11418: push ebx
  loc_00E11419: mov ecx, [ebx]
  loc_00E1141B: call [ecx+00000054h]
  loc_00E1141E: test eax, eax
  loc_00E11420: fnclex
  loc_00E11422: jge 00E11433h
  loc_00E11424: push 00000054h
  loc_00E11426: push 006DCBE8h
  loc_00E1142B: push ebx
  loc_00E1142C: push eax
  loc_00E1142D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11433: lea ebx, var_30
  loc_00E11436: mov eax, var_2C
  loc_00E11439: push ebx
  loc_00E1143A: mov ecx, 00000002h
  loc_00E1143F: sub esp, 00000010h
  loc_00E11442: mov var_84, ecx
  loc_00E11448: mov ebx, esp
  loc_00E1144A: mov var_7C, 0000000Ah
  loc_00E11451: mov edx, [eax]
  loc_00E11453: push eax
  loc_00E11454: mov [ebx], ecx
  loc_00E11456: mov ecx, var_80
  loc_00E11459: mov var_C4, eax
  loc_00E1145F: mov [ebx+00000004h], ecx
  loc_00E11462: mov ecx, var_7C
  loc_00E11465: mov [ebx+00000008h], ecx
  loc_00E11468: mov ecx, var_78
  loc_00E1146B: mov [ebx+0000000Ch], ecx
  loc_00E1146E: call [edx+00000028h]
  loc_00E11471: test eax, eax
  loc_00E11473: fnclex
  loc_00E11475: jge 00E1148Ch
  loc_00E11477: mov edx, var_C4
  loc_00E1147D: push 00000028h
  loc_00E1147F: push 006E09E8h
  loc_00E11484: push edx
  loc_00E11485: push eax
  loc_00E11486: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1148C: mov eax, var_30
  loc_00E1148F: lea edx, var_54
  loc_00E11492: push edx
  loc_00E11493: push eax
  loc_00E11494: mov ecx, [eax]
  loc_00E11496: mov ebx, eax
  loc_00E11498: call [ecx+00000034h]
  loc_00E1149B: test eax, eax
  loc_00E1149D: fnclex
  loc_00E1149F: jge 00E114B0h
  loc_00E114A1: push 00000034h
  loc_00E114A3: push 006E09F8h
  loc_00E114A8: push ebx
  loc_00E114A9: push eax
  loc_00E114AA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E114B0: mov eax, var_D4
  loc_00E114B6: lea ecx, var_54
  loc_00E114B9: push ecx
  loc_00E114BA: mov ebx, [eax]
  loc_00E114BC: call [00401034h] ; __vbaStrVarMove
  loc_00E114C2: mov edx, eax
  loc_00E114C4: lea ecx, var_18
  loc_00E114C7: call [00401228h] ; __vbaStrMove
  loc_00E114CD: mov edx, ebx
  loc_00E114CF: mov ebx, var_D4
  loc_00E114D5: push eax
  loc_00E114D6: push ebx
  loc_00E114D7: call [edx+000000A4h]
  loc_00E114DD: test eax, eax
  loc_00E114DF: fnclex
  loc_00E114E1: jge 00E114F5h
  loc_00E114E3: push 000000A4h
  loc_00E114E8: push 006DCB70h
  loc_00E114ED: push ebx
  loc_00E114EE: push eax
  loc_00E114EF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E114F5: lea ecx, var_18
  loc_00E114F8: call [00401258h] ; __vbaFreeStr
  loc_00E114FE: lea eax, var_34
  loc_00E11501: lea ecx, var_30
  loc_00E11504: push eax
  loc_00E11505: lea edx, var_2C
  loc_00E11508: push ecx
  loc_00E11509: lea eax, var_28
  loc_00E1150C: push edx
  loc_00E1150D: lea ecx, var_24
  loc_00E11510: push eax
  loc_00E11511: push ecx
  loc_00E11512: push 00000005h
  loc_00E11514: call [00401048h] ; __vbaFreeObjList
  loc_00E1151A: lea edx, var_54
  loc_00E1151D: lea eax, var_44
  loc_00E11520: push edx
  loc_00E11521: push eax
  loc_00E11522: push 00000002h
  loc_00E11524: call [00401038h] ; __vbaFreeVarList
  loc_00E1152A: mov ecx, [esi]
  loc_00E1152C: add esp, 00000024h
  loc_00E1152F: push esi
  loc_00E11530: call [ecx+0000036Ch]
  loc_00E11536: lea edx, var_34
  loc_00E11539: push eax
  loc_00E1153A: push edx
  loc_00E1153B: call edi
  loc_00E1153D: push 006DCBD8h
  loc_00E11542: mov var_D4, eax
  loc_00E11548: mov eax, [esi]
  loc_00E1154A: push 00000000h
  loc_00E1154C: push 00000018h
  loc_00E1154E: push esi
  loc_00E1154F: call [eax+00000490h]
  loc_00E11555: lea ecx, var_24
  loc_00E11558: push eax
  loc_00E11559: push ecx
  loc_00E1155A: call edi
  loc_00E1155C: lea edx, var_44
  loc_00E1155F: push eax
  loc_00E11560: push edx
  loc_00E11561: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11567: add esp, 00000010h
  loc_00E1156A: push eax
  loc_00E1156B: call [00401120h] ; __vbaCastObjVar
  loc_00E11571: push eax
  loc_00E11572: lea eax, var_28
  loc_00E11575: push eax
  loc_00E11576: call edi
  loc_00E11578: mov ebx, eax
  loc_00E1157A: lea edx, var_2C
  loc_00E1157D: push edx
  loc_00E1157E: push ebx
  loc_00E1157F: mov ecx, [ebx]
  loc_00E11581: call [ecx+00000054h]
  loc_00E11584: test eax, eax
  loc_00E11586: fnclex
  loc_00E11588: jge 00E11599h
  loc_00E1158A: push 00000054h
  loc_00E1158C: push 006DCBE8h
  loc_00E11591: push ebx
  loc_00E11592: push eax
  loc_00E11593: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11599: lea ebx, var_30
  loc_00E1159C: mov eax, var_2C
  loc_00E1159F: push ebx
  loc_00E115A0: mov ecx, 00000002h
  loc_00E115A5: sub esp, 00000010h
  loc_00E115A8: mov var_84, ecx
  loc_00E115AE: mov ebx, esp
  loc_00E115B0: mov var_7C, 0000000Bh
  loc_00E115B7: mov edx, [eax]
  loc_00E115B9: push eax
  loc_00E115BA: mov [ebx], ecx
  loc_00E115BC: mov ecx, var_80
  loc_00E115BF: mov var_C4, eax
  loc_00E115C5: mov [ebx+00000004h], ecx
  loc_00E115C8: mov ecx, var_7C
  loc_00E115CB: mov [ebx+00000008h], ecx
  loc_00E115CE: mov ecx, var_78
  loc_00E115D1: mov [ebx+0000000Ch], ecx
  loc_00E115D4: call [edx+00000028h]
  loc_00E115D7: test eax, eax
  loc_00E115D9: fnclex
  loc_00E115DB: jge 00E115F2h
  loc_00E115DD: mov edx, var_C4
  loc_00E115E3: push 00000028h
  loc_00E115E5: push 006E09E8h
  loc_00E115EA: push edx
  loc_00E115EB: push eax
  loc_00E115EC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E115F2: mov eax, var_30
  loc_00E115F5: lea edx, var_54
  loc_00E115F8: push edx
  loc_00E115F9: push eax
  loc_00E115FA: mov ecx, [eax]
  loc_00E115FC: mov ebx, eax
  loc_00E115FE: call [ecx+00000034h]
  loc_00E11601: test eax, eax
  loc_00E11603: fnclex
  loc_00E11605: jge 00E11616h
  loc_00E11607: push 00000034h
  loc_00E11609: push 006E09F8h
  loc_00E1160E: push ebx
  loc_00E1160F: push eax
  loc_00E11610: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11616: mov eax, var_D4
  loc_00E1161C: lea ecx, var_54
  loc_00E1161F: push ecx
  loc_00E11620: mov ebx, [eax]
  loc_00E11622: call [00401034h] ; __vbaStrVarMove
  loc_00E11628: mov edx, eax
  loc_00E1162A: lea ecx, var_18
  loc_00E1162D: call [00401228h] ; __vbaStrMove
  loc_00E11633: mov edx, ebx
  loc_00E11635: mov ebx, var_D4
  loc_00E1163B: push eax
  loc_00E1163C: push ebx
  loc_00E1163D: call [edx+000000A4h]
  loc_00E11643: test eax, eax
  loc_00E11645: fnclex
  loc_00E11647: jge 00E1165Bh
  loc_00E11649: push 000000A4h
  loc_00E1164E: push 006DCB70h
  loc_00E11653: push ebx
  loc_00E11654: push eax
  loc_00E11655: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1165B: lea ecx, var_18
  loc_00E1165E: call [00401258h] ; __vbaFreeStr
  loc_00E11664: lea eax, var_34
  loc_00E11667: lea ecx, var_30
  loc_00E1166A: push eax
  loc_00E1166B: lea edx, var_2C
  loc_00E1166E: push ecx
  loc_00E1166F: lea eax, var_28
  loc_00E11672: push edx
  loc_00E11673: lea ecx, var_24
  loc_00E11676: push eax
  loc_00E11677: push ecx
  loc_00E11678: push 00000005h
  loc_00E1167A: call [00401048h] ; __vbaFreeObjList
  loc_00E11680: lea edx, var_54
  loc_00E11683: lea eax, var_44
  loc_00E11686: push edx
  loc_00E11687: push eax
  loc_00E11688: push 00000002h
  loc_00E1168A: call [00401038h] ; __vbaFreeVarList
  loc_00E11690: mov ecx, [esi]
  loc_00E11692: add esp, 00000024h
  loc_00E11695: push esi
  loc_00E11696: call [ecx+00000368h]
  loc_00E1169C: lea edx, var_34
  loc_00E1169F: push eax
  loc_00E116A0: push edx
  loc_00E116A1: call edi
  loc_00E116A3: push 006DCBD8h
  loc_00E116A8: mov var_D4, eax
  loc_00E116AE: mov eax, [esi]
  loc_00E116B0: push 00000000h
  loc_00E116B2: push 00000018h
  loc_00E116B4: push esi
  loc_00E116B5: call [eax+00000490h]
  loc_00E116BB: lea ecx, var_24
  loc_00E116BE: push eax
  loc_00E116BF: push ecx
  loc_00E116C0: call edi
  loc_00E116C2: lea edx, var_44
  loc_00E116C5: push eax
  loc_00E116C6: push edx
  loc_00E116C7: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E116CD: add esp, 00000010h
  loc_00E116D0: push eax
  loc_00E116D1: call [00401120h] ; __vbaCastObjVar
  loc_00E116D7: push eax
  loc_00E116D8: lea eax, var_28
  loc_00E116DB: push eax
  loc_00E116DC: call edi
  loc_00E116DE: mov ebx, eax
  loc_00E116E0: lea edx, var_2C
  loc_00E116E3: push edx
  loc_00E116E4: push ebx
  loc_00E116E5: mov ecx, [ebx]
  loc_00E116E7: call [ecx+00000054h]
  loc_00E116EA: test eax, eax
  loc_00E116EC: fnclex
  loc_00E116EE: jge 00E116FFh
  loc_00E116F0: push 00000054h
  loc_00E116F2: push 006DCBE8h
  loc_00E116F7: push ebx
  loc_00E116F8: push eax
  loc_00E116F9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E116FF: lea ebx, var_30
  loc_00E11702: mov eax, var_2C
  loc_00E11705: push ebx
  loc_00E11706: mov ecx, 00000002h
  loc_00E1170B: sub esp, 00000010h
  loc_00E1170E: mov var_84, ecx
  loc_00E11714: mov ebx, esp
  loc_00E11716: mov var_7C, 0000000Ch
  loc_00E1171D: mov edx, [eax]
  loc_00E1171F: push eax
  loc_00E11720: mov [ebx], ecx
  loc_00E11722: mov ecx, var_80
  loc_00E11725: mov var_C4, eax
  loc_00E1172B: mov [ebx+00000004h], ecx
  loc_00E1172E: mov ecx, var_7C
  loc_00E11731: mov [ebx+00000008h], ecx
  loc_00E11734: mov ecx, var_78
  loc_00E11737: mov [ebx+0000000Ch], ecx
  loc_00E1173A: call [edx+00000028h]
  loc_00E1173D: test eax, eax
  loc_00E1173F: fnclex
  loc_00E11741: jge 00E11758h
  loc_00E11743: mov edx, var_C4
  loc_00E11749: push 00000028h
  loc_00E1174B: push 006E09E8h
  loc_00E11750: push edx
  loc_00E11751: push eax
  loc_00E11752: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11758: mov eax, var_30
  loc_00E1175B: lea edx, var_54
  loc_00E1175E: push edx
  loc_00E1175F: push eax
  loc_00E11760: mov ecx, [eax]
  loc_00E11762: mov ebx, eax
  loc_00E11764: call [ecx+00000034h]
  loc_00E11767: test eax, eax
  loc_00E11769: fnclex
  loc_00E1176B: jge 00E1177Ch
  loc_00E1176D: push 00000034h
  loc_00E1176F: push 006E09F8h
  loc_00E11774: push ebx
  loc_00E11775: push eax
  loc_00E11776: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1177C: mov eax, var_D4
  loc_00E11782: lea ecx, var_54
  loc_00E11785: push ecx
  loc_00E11786: mov ebx, [eax]
  loc_00E11788: call [00401034h] ; __vbaStrVarMove
  loc_00E1178E: mov edx, eax
  loc_00E11790: lea ecx, var_18
  loc_00E11793: call [00401228h] ; __vbaStrMove
  loc_00E11799: mov edx, ebx
  loc_00E1179B: mov ebx, var_D4
  loc_00E117A1: push eax
  loc_00E117A2: push ebx
  loc_00E117A3: call [edx+000000A4h]
  loc_00E117A9: test eax, eax
  loc_00E117AB: fnclex
  loc_00E117AD: jge 00E117C1h
  loc_00E117AF: push 000000A4h
  loc_00E117B4: push 006DCB70h
  loc_00E117B9: push ebx
  loc_00E117BA: push eax
  loc_00E117BB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E117C1: lea ecx, var_18
  loc_00E117C4: call [00401258h] ; __vbaFreeStr
  loc_00E117CA: lea eax, var_34
  loc_00E117CD: lea ecx, var_30
  loc_00E117D0: push eax
  loc_00E117D1: lea edx, var_2C
  loc_00E117D4: push ecx
  loc_00E117D5: lea eax, var_28
  loc_00E117D8: push edx
  loc_00E117D9: lea ecx, var_24
  loc_00E117DC: push eax
  loc_00E117DD: push ecx
  loc_00E117DE: push 00000005h
  loc_00E117E0: call [00401048h] ; __vbaFreeObjList
  loc_00E117E6: lea edx, var_54
  loc_00E117E9: lea eax, var_44
  loc_00E117EC: push edx
  loc_00E117ED: push eax
  loc_00E117EE: push 00000002h
  loc_00E117F0: call [00401038h] ; __vbaFreeVarList
  loc_00E117F6: mov ecx, [esi]
  loc_00E117F8: add esp, 00000024h
  loc_00E117FB: push esi
  loc_00E117FC: call [ecx+00000364h]
  loc_00E11802: lea edx, var_34
  loc_00E11805: push eax
  loc_00E11806: push edx
  loc_00E11807: call edi
  loc_00E11809: push 006DCBD8h
  loc_00E1180E: mov var_D4, eax
  loc_00E11814: mov eax, [esi]
  loc_00E11816: push 00000000h
  loc_00E11818: push 00000018h
  loc_00E1181A: push esi
  loc_00E1181B: call [eax+00000490h]
  loc_00E11821: lea ecx, var_24
  loc_00E11824: push eax
  loc_00E11825: push ecx
  loc_00E11826: call edi
  loc_00E11828: lea edx, var_44
  loc_00E1182B: push eax
  loc_00E1182C: push edx
  loc_00E1182D: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11833: add esp, 00000010h
  loc_00E11836: push eax
  loc_00E11837: call [00401120h] ; __vbaCastObjVar
  loc_00E1183D: push eax
  loc_00E1183E: lea eax, var_28
  loc_00E11841: push eax
  loc_00E11842: call edi
  loc_00E11844: mov ebx, eax
  loc_00E11846: lea edx, var_2C
  loc_00E11849: push edx
  loc_00E1184A: push ebx
  loc_00E1184B: mov ecx, [ebx]
  loc_00E1184D: call [ecx+00000054h]
  loc_00E11850: test eax, eax
  loc_00E11852: fnclex
  loc_00E11854: jge 00E11865h
  loc_00E11856: push 00000054h
  loc_00E11858: push 006DCBE8h
  loc_00E1185D: push ebx
  loc_00E1185E: push eax
  loc_00E1185F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11865: lea ebx, var_30
  loc_00E11868: mov eax, var_2C
  loc_00E1186B: push ebx
  loc_00E1186C: mov ecx, 00000002h
  loc_00E11871: sub esp, 00000010h
  loc_00E11874: mov var_84, ecx
  loc_00E1187A: mov ebx, esp
  loc_00E1187C: mov var_7C, 0000000Dh
  loc_00E11883: mov edx, [eax]
  loc_00E11885: push eax
  loc_00E11886: mov [ebx], ecx
  loc_00E11888: mov ecx, var_80
  loc_00E1188B: mov var_C4, eax
  loc_00E11891: mov [ebx+00000004h], ecx
  loc_00E11894: mov ecx, var_7C
  loc_00E11897: mov [ebx+00000008h], ecx
  loc_00E1189A: mov ecx, var_78
  loc_00E1189D: mov [ebx+0000000Ch], ecx
  loc_00E118A0: call [edx+00000028h]
  loc_00E118A3: test eax, eax
  loc_00E118A5: fnclex
  loc_00E118A7: jge 00E118BEh
  loc_00E118A9: mov edx, var_C4
  loc_00E118AF: push 00000028h
  loc_00E118B1: push 006E09E8h
  loc_00E118B6: push edx
  loc_00E118B7: push eax
  loc_00E118B8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E118BE: mov eax, var_30
  loc_00E118C1: lea edx, var_54
  loc_00E118C4: push edx
  loc_00E118C5: push eax
  loc_00E118C6: mov ecx, [eax]
  loc_00E118C8: mov ebx, eax
  loc_00E118CA: call [ecx+00000034h]
  loc_00E118CD: test eax, eax
  loc_00E118CF: fnclex
  loc_00E118D1: jge 00E118E2h
  loc_00E118D3: push 00000034h
  loc_00E118D5: push 006E09F8h
  loc_00E118DA: push ebx
  loc_00E118DB: push eax
  loc_00E118DC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E118E2: mov eax, var_D4
  loc_00E118E8: lea ecx, var_54
  loc_00E118EB: push ecx
  loc_00E118EC: mov ebx, [eax]
  loc_00E118EE: call [00401034h] ; __vbaStrVarMove
  loc_00E118F4: mov edx, eax
  loc_00E118F6: lea ecx, var_18
  loc_00E118F9: call [00401228h] ; __vbaStrMove
  loc_00E118FF: mov edx, ebx
  loc_00E11901: mov ebx, var_D4
  loc_00E11907: push eax
  loc_00E11908: push ebx
  loc_00E11909: call [edx+000000A4h]
  loc_00E1190F: test eax, eax
  loc_00E11911: fnclex
  loc_00E11913: jge 00E11927h
  loc_00E11915: push 000000A4h
  loc_00E1191A: push 006DCB70h
  loc_00E1191F: push ebx
  loc_00E11920: push eax
  loc_00E11921: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11927: lea ecx, var_18
  loc_00E1192A: call [00401258h] ; __vbaFreeStr
  loc_00E11930: lea eax, var_34
  loc_00E11933: lea ecx, var_30
  loc_00E11936: push eax
  loc_00E11937: lea edx, var_2C
  loc_00E1193A: push ecx
  loc_00E1193B: lea eax, var_28
  loc_00E1193E: push edx
  loc_00E1193F: lea ecx, var_24
  loc_00E11942: push eax
  loc_00E11943: push ecx
  loc_00E11944: push 00000005h
  loc_00E11946: call [00401048h] ; __vbaFreeObjList
  loc_00E1194C: lea edx, var_54
  loc_00E1194F: lea eax, var_44
  loc_00E11952: push edx
  loc_00E11953: push eax
  loc_00E11954: push 00000002h
  loc_00E11956: call [00401038h] ; __vbaFreeVarList
  loc_00E1195C: mov ecx, [esi]
  loc_00E1195E: add esp, 00000024h
  loc_00E11961: push esi
  loc_00E11962: call [ecx+00000360h]
  loc_00E11968: lea edx, var_34
  loc_00E1196B: push eax
  loc_00E1196C: push edx
  loc_00E1196D: call edi
  loc_00E1196F: push 006DCBD8h
  loc_00E11974: mov var_D4, eax
  loc_00E1197A: mov eax, [esi]
  loc_00E1197C: push 00000000h
  loc_00E1197E: push 00000018h
  loc_00E11980: push esi
  loc_00E11981: call [eax+00000490h]
  loc_00E11987: lea ecx, var_24
  loc_00E1198A: push eax
  loc_00E1198B: push ecx
  loc_00E1198C: call edi
  loc_00E1198E: lea edx, var_44
  loc_00E11991: push eax
  loc_00E11992: push edx
  loc_00E11993: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11999: add esp, 00000010h
  loc_00E1199C: push eax
  loc_00E1199D: call [00401120h] ; __vbaCastObjVar
  loc_00E119A3: push eax
  loc_00E119A4: lea eax, var_28
  loc_00E119A7: push eax
  loc_00E119A8: call edi
  loc_00E119AA: mov ebx, eax
  loc_00E119AC: lea edx, var_2C
  loc_00E119AF: push edx
  loc_00E119B0: push ebx
  loc_00E119B1: mov ecx, [ebx]
  loc_00E119B3: call [ecx+00000054h]
  loc_00E119B6: test eax, eax
  loc_00E119B8: fnclex
  loc_00E119BA: jge 00E119CBh
  loc_00E119BC: push 00000054h
  loc_00E119BE: push 006DCBE8h
  loc_00E119C3: push ebx
  loc_00E119C4: push eax
  loc_00E119C5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E119CB: lea ebx, var_30
  loc_00E119CE: mov eax, var_2C
  loc_00E119D1: push ebx
  loc_00E119D2: mov ecx, 00000002h
  loc_00E119D7: sub esp, 00000010h
  loc_00E119DA: mov var_84, ecx
  loc_00E119E0: mov ebx, esp
  loc_00E119E2: mov var_7C, 0000000Eh
  loc_00E119E9: mov edx, [eax]
  loc_00E119EB: push eax
  loc_00E119EC: mov [ebx], ecx
  loc_00E119EE: mov ecx, var_80
  loc_00E119F1: mov var_C4, eax
  loc_00E119F7: mov [ebx+00000004h], ecx
  loc_00E119FA: mov ecx, var_7C
  loc_00E119FD: mov [ebx+00000008h], ecx
  loc_00E11A00: mov ecx, var_78
  loc_00E11A03: mov [ebx+0000000Ch], ecx
  loc_00E11A06: call [edx+00000028h]
  loc_00E11A09: test eax, eax
  loc_00E11A0B: fnclex
  loc_00E11A0D: jge 00E11A24h
  loc_00E11A0F: mov edx, var_C4
  loc_00E11A15: push 00000028h
  loc_00E11A17: push 006E09E8h
  loc_00E11A1C: push edx
  loc_00E11A1D: push eax
  loc_00E11A1E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11A24: mov eax, var_30
  loc_00E11A27: lea edx, var_54
  loc_00E11A2A: push edx
  loc_00E11A2B: push eax
  loc_00E11A2C: mov ecx, [eax]
  loc_00E11A2E: mov ebx, eax
  loc_00E11A30: call [ecx+00000034h]
  loc_00E11A33: test eax, eax
  loc_00E11A35: fnclex
  loc_00E11A37: jge 00E11A48h
  loc_00E11A39: push 00000034h
  loc_00E11A3B: push 006E09F8h
  loc_00E11A40: push ebx
  loc_00E11A41: push eax
  loc_00E11A42: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11A48: mov eax, var_D4
  loc_00E11A4E: lea ecx, var_54
  loc_00E11A51: push ecx
  loc_00E11A52: mov ebx, [eax]
  loc_00E11A54: call [00401034h] ; __vbaStrVarMove
  loc_00E11A5A: mov edx, eax
  loc_00E11A5C: lea ecx, var_18
  loc_00E11A5F: call [00401228h] ; __vbaStrMove
  loc_00E11A65: mov edx, ebx
  loc_00E11A67: mov ebx, var_D4
  loc_00E11A6D: push eax
  loc_00E11A6E: push ebx
  loc_00E11A6F: call [edx+000000A4h]
  loc_00E11A75: test eax, eax
  loc_00E11A77: fnclex
  loc_00E11A79: jge 00E11A8Dh
  loc_00E11A7B: push 000000A4h
  loc_00E11A80: push 006DCB70h
  loc_00E11A85: push ebx
  loc_00E11A86: push eax
  loc_00E11A87: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11A8D: lea ecx, var_18
  loc_00E11A90: call [00401258h] ; __vbaFreeStr
  loc_00E11A96: lea eax, var_34
  loc_00E11A99: lea ecx, var_30
  loc_00E11A9C: push eax
  loc_00E11A9D: lea edx, var_2C
  loc_00E11AA0: push ecx
  loc_00E11AA1: lea eax, var_28
  loc_00E11AA4: push edx
  loc_00E11AA5: lea ecx, var_24
  loc_00E11AA8: push eax
  loc_00E11AA9: push ecx
  loc_00E11AAA: push 00000005h
  loc_00E11AAC: call [00401048h] ; __vbaFreeObjList
  loc_00E11AB2: lea edx, var_54
  loc_00E11AB5: lea eax, var_44
  loc_00E11AB8: push edx
  loc_00E11AB9: push eax
  loc_00E11ABA: push 00000002h
  loc_00E11ABC: call [00401038h] ; __vbaFreeVarList
  loc_00E11AC2: mov ecx, [esi]
  loc_00E11AC4: add esp, 00000024h
  loc_00E11AC7: push esi
  loc_00E11AC8: call [ecx+0000033Ch]
  loc_00E11ACE: lea edx, var_34
  loc_00E11AD1: push eax
  loc_00E11AD2: push edx
  loc_00E11AD3: call edi
  loc_00E11AD5: push 006DCBD8h
  loc_00E11ADA: mov var_D4, eax
  loc_00E11AE0: mov eax, [esi]
  loc_00E11AE2: push 00000000h
  loc_00E11AE4: push 00000018h
  loc_00E11AE6: push esi
  loc_00E11AE7: call [eax+00000490h]
  loc_00E11AED: lea ecx, var_24
  loc_00E11AF0: push eax
  loc_00E11AF1: push ecx
  loc_00E11AF2: call edi
  loc_00E11AF4: lea edx, var_44
  loc_00E11AF7: push eax
  loc_00E11AF8: push edx
  loc_00E11AF9: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11AFF: add esp, 00000010h
  loc_00E11B02: push eax
  loc_00E11B03: call [00401120h] ; __vbaCastObjVar
  loc_00E11B09: push eax
  loc_00E11B0A: lea eax, var_28
  loc_00E11B0D: push eax
  loc_00E11B0E: call edi
  loc_00E11B10: mov ebx, eax
  loc_00E11B12: lea edx, var_2C
  loc_00E11B15: push edx
  loc_00E11B16: push ebx
  loc_00E11B17: mov ecx, [ebx]
  loc_00E11B19: call [ecx+00000054h]
  loc_00E11B1C: test eax, eax
  loc_00E11B1E: fnclex
  loc_00E11B20: jge 00E11B31h
  loc_00E11B22: push 00000054h
  loc_00E11B24: push 006DCBE8h
  loc_00E11B29: push ebx
  loc_00E11B2A: push eax
  loc_00E11B2B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11B31: lea ebx, var_30
  loc_00E11B34: mov eax, var_2C
  loc_00E11B37: push ebx
  loc_00E11B38: mov ecx, 00000002h
  loc_00E11B3D: sub esp, 00000010h
  loc_00E11B40: mov var_84, ecx
  loc_00E11B46: mov ebx, esp
  loc_00E11B48: mov var_7C, 0000000Fh
  loc_00E11B4F: mov edx, [eax]
  loc_00E11B51: push eax
  loc_00E11B52: mov [ebx], ecx
  loc_00E11B54: mov ecx, var_80
  loc_00E11B57: mov var_C4, eax
  loc_00E11B5D: mov [ebx+00000004h], ecx
  loc_00E11B60: mov ecx, var_7C
  loc_00E11B63: mov [ebx+00000008h], ecx
  loc_00E11B66: mov ecx, var_78
  loc_00E11B69: mov [ebx+0000000Ch], ecx
  loc_00E11B6C: call [edx+00000028h]
  loc_00E11B6F: test eax, eax
  loc_00E11B71: fnclex
  loc_00E11B73: jge 00E11B8Ah
  loc_00E11B75: mov edx, var_C4
  loc_00E11B7B: push 00000028h
  loc_00E11B7D: push 006E09E8h
  loc_00E11B82: push edx
  loc_00E11B83: push eax
  loc_00E11B84: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11B8A: mov eax, var_30
  loc_00E11B8D: lea edx, var_54
  loc_00E11B90: push edx
  loc_00E11B91: push eax
  loc_00E11B92: mov ecx, [eax]
  loc_00E11B94: mov ebx, eax
  loc_00E11B96: call [ecx+00000034h]
  loc_00E11B99: test eax, eax
  loc_00E11B9B: fnclex
  loc_00E11B9D: jge 00E11BAEh
  loc_00E11B9F: push 00000034h
  loc_00E11BA1: push 006E09F8h
  loc_00E11BA6: push ebx
  loc_00E11BA7: push eax
  loc_00E11BA8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11BAE: mov eax, var_D4
  loc_00E11BB4: lea ecx, var_54
  loc_00E11BB7: push ecx
  loc_00E11BB8: mov ebx, [eax]
  loc_00E11BBA: call [00401034h] ; __vbaStrVarMove
  loc_00E11BC0: mov edx, eax
  loc_00E11BC2: lea ecx, var_18
  loc_00E11BC5: call [00401228h] ; __vbaStrMove
  loc_00E11BCB: mov edx, ebx
  loc_00E11BCD: mov ebx, var_D4
  loc_00E11BD3: push eax
  loc_00E11BD4: push ebx
  loc_00E11BD5: call [edx+000000A4h]
  loc_00E11BDB: test eax, eax
  loc_00E11BDD: fnclex
  loc_00E11BDF: jge 00E11BF3h
  loc_00E11BE1: push 000000A4h
  loc_00E11BE6: push 006DCB70h
  loc_00E11BEB: push ebx
  loc_00E11BEC: push eax
  loc_00E11BED: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11BF3: lea ecx, var_18
  loc_00E11BF6: call [00401258h] ; __vbaFreeStr
  loc_00E11BFC: lea eax, var_34
  loc_00E11BFF: lea ecx, var_30
  loc_00E11C02: push eax
  loc_00E11C03: lea edx, var_2C
  loc_00E11C06: push ecx
  loc_00E11C07: lea eax, var_28
  loc_00E11C0A: push edx
  loc_00E11C0B: lea ecx, var_24
  loc_00E11C0E: push eax
  loc_00E11C0F: push ecx
  loc_00E11C10: push 00000005h
  loc_00E11C12: call [00401048h] ; __vbaFreeObjList
  loc_00E11C18: lea edx, var_54
  loc_00E11C1B: lea eax, var_44
  loc_00E11C1E: push edx
  loc_00E11C1F: push eax
  loc_00E11C20: push 00000002h
  loc_00E11C22: call [00401038h] ; __vbaFreeVarList
  loc_00E11C28: mov ecx, [esi]
  loc_00E11C2A: add esp, 00000024h
  loc_00E11C2D: push esi
  loc_00E11C2E: call [ecx+00000340h]
  loc_00E11C34: lea edx, var_34
  loc_00E11C37: push eax
  loc_00E11C38: push edx
  loc_00E11C39: call edi
  loc_00E11C3B: push 006DCBD8h
  loc_00E11C40: mov var_D4, eax
  loc_00E11C46: mov eax, [esi]
  loc_00E11C48: push 00000000h
  loc_00E11C4A: push 00000018h
  loc_00E11C4C: push esi
  loc_00E11C4D: call [eax+00000490h]
  loc_00E11C53: lea ecx, var_24
  loc_00E11C56: push eax
  loc_00E11C57: push ecx
  loc_00E11C58: call edi
  loc_00E11C5A: lea edx, var_44
  loc_00E11C5D: push eax
  loc_00E11C5E: push edx
  loc_00E11C5F: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11C65: add esp, 00000010h
  loc_00E11C68: push eax
  loc_00E11C69: call [00401120h] ; __vbaCastObjVar
  loc_00E11C6F: push eax
  loc_00E11C70: lea eax, var_28
  loc_00E11C73: push eax
  loc_00E11C74: call edi
  loc_00E11C76: mov ebx, eax
  loc_00E11C78: lea edx, var_2C
  loc_00E11C7B: push edx
  loc_00E11C7C: push ebx
  loc_00E11C7D: mov ecx, [ebx]
  loc_00E11C7F: call [ecx+00000054h]
  loc_00E11C82: test eax, eax
  loc_00E11C84: fnclex
  loc_00E11C86: jge 00E11C97h
  loc_00E11C88: push 00000054h
  loc_00E11C8A: push 006DCBE8h
  loc_00E11C8F: push ebx
  loc_00E11C90: push eax
  loc_00E11C91: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11C97: lea ebx, var_30
  loc_00E11C9A: mov eax, var_2C
  loc_00E11C9D: push ebx
  loc_00E11C9E: mov ecx, 00000002h
  loc_00E11CA3: sub esp, 00000010h
  loc_00E11CA6: mov var_84, ecx
  loc_00E11CAC: mov ebx, esp
  loc_00E11CAE: mov var_7C, 00000010h
  loc_00E11CB5: mov edx, [eax]
  loc_00E11CB7: push eax
  loc_00E11CB8: mov [ebx], ecx
  loc_00E11CBA: mov ecx, var_80
  loc_00E11CBD: mov var_C4, eax
  loc_00E11CC3: mov [ebx+00000004h], ecx
  loc_00E11CC6: mov ecx, var_7C
  loc_00E11CC9: mov [ebx+00000008h], ecx
  loc_00E11CCC: mov ecx, var_78
  loc_00E11CCF: mov [ebx+0000000Ch], ecx
  loc_00E11CD2: call [edx+00000028h]
  loc_00E11CD5: test eax, eax
  loc_00E11CD7: fnclex
  loc_00E11CD9: jge 00E11CF0h
  loc_00E11CDB: mov edx, var_C4
  loc_00E11CE1: push 00000028h
  loc_00E11CE3: push 006E09E8h
  loc_00E11CE8: push edx
  loc_00E11CE9: push eax
  loc_00E11CEA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11CF0: mov eax, var_30
  loc_00E11CF3: lea edx, var_54
  loc_00E11CF6: push edx
  loc_00E11CF7: push eax
  loc_00E11CF8: mov ecx, [eax]
  loc_00E11CFA: mov ebx, eax
  loc_00E11CFC: call [ecx+00000034h]
  loc_00E11CFF: test eax, eax
  loc_00E11D01: fnclex
  loc_00E11D03: jge 00E11D14h
  loc_00E11D05: push 00000034h
  loc_00E11D07: push 006E09F8h
  loc_00E11D0C: push ebx
  loc_00E11D0D: push eax
  loc_00E11D0E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11D14: mov eax, var_D4
  loc_00E11D1A: lea ecx, var_54
  loc_00E11D1D: push ecx
  loc_00E11D1E: mov ebx, [eax]
  loc_00E11D20: call [00401034h] ; __vbaStrVarMove
  loc_00E11D26: mov edx, eax
  loc_00E11D28: lea ecx, var_18
  loc_00E11D2B: call [00401228h] ; __vbaStrMove
  loc_00E11D31: mov edx, ebx
  loc_00E11D33: mov ebx, var_D4
  loc_00E11D39: push eax
  loc_00E11D3A: push ebx
  loc_00E11D3B: call [edx+000000A4h]
  loc_00E11D41: test eax, eax
  loc_00E11D43: fnclex
  loc_00E11D45: jge 00E11D59h
  loc_00E11D47: push 000000A4h
  loc_00E11D4C: push 006DCB70h
  loc_00E11D51: push ebx
  loc_00E11D52: push eax
  loc_00E11D53: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11D59: lea ecx, var_18
  loc_00E11D5C: call [00401258h] ; __vbaFreeStr
  loc_00E11D62: lea eax, var_34
  loc_00E11D65: lea ecx, var_30
  loc_00E11D68: push eax
  loc_00E11D69: lea edx, var_2C
  loc_00E11D6C: push ecx
  loc_00E11D6D: lea eax, var_28
  loc_00E11D70: push edx
  loc_00E11D71: lea ecx, var_24
  loc_00E11D74: push eax
  loc_00E11D75: push ecx
  loc_00E11D76: push 00000005h
  loc_00E11D78: call [00401048h] ; __vbaFreeObjList
  loc_00E11D7E: lea edx, var_54
  loc_00E11D81: lea eax, var_44
  loc_00E11D84: push edx
  loc_00E11D85: push eax
  loc_00E11D86: push 00000002h
  loc_00E11D88: call [00401038h] ; __vbaFreeVarList
  loc_00E11D8E: mov ecx, [esi]
  loc_00E11D90: add esp, 00000024h
  loc_00E11D93: push esi
  loc_00E11D94: call [ecx+00000344h]
  loc_00E11D9A: lea edx, var_34
  loc_00E11D9D: push eax
  loc_00E11D9E: push edx
  loc_00E11D9F: call edi
  loc_00E11DA1: push 006DCBD8h
  loc_00E11DA6: mov var_D4, eax
  loc_00E11DAC: mov eax, [esi]
  loc_00E11DAE: push 00000000h
  loc_00E11DB0: push 00000018h
  loc_00E11DB2: push esi
  loc_00E11DB3: call [eax+00000490h]
  loc_00E11DB9: lea ecx, var_24
  loc_00E11DBC: push eax
  loc_00E11DBD: push ecx
  loc_00E11DBE: call edi
  loc_00E11DC0: lea edx, var_44
  loc_00E11DC3: push eax
  loc_00E11DC4: push edx
  loc_00E11DC5: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11DCB: add esp, 00000010h
  loc_00E11DCE: push eax
  loc_00E11DCF: call [00401120h] ; __vbaCastObjVar
  loc_00E11DD5: push eax
  loc_00E11DD6: lea eax, var_28
  loc_00E11DD9: push eax
  loc_00E11DDA: call edi
  loc_00E11DDC: mov ebx, eax
  loc_00E11DDE: lea edx, var_2C
  loc_00E11DE1: push edx
  loc_00E11DE2: push ebx
  loc_00E11DE3: mov ecx, [ebx]
  loc_00E11DE5: call [ecx+00000054h]
  loc_00E11DE8: test eax, eax
  loc_00E11DEA: fnclex
  loc_00E11DEC: jge 00E11DFDh
  loc_00E11DEE: push 00000054h
  loc_00E11DF0: push 006DCBE8h
  loc_00E11DF5: push ebx
  loc_00E11DF6: push eax
  loc_00E11DF7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11DFD: lea ebx, var_30
  loc_00E11E00: mov eax, var_2C
  loc_00E11E03: push ebx
  loc_00E11E04: mov ecx, 00000002h
  loc_00E11E09: sub esp, 00000010h
  loc_00E11E0C: mov var_84, ecx
  loc_00E11E12: mov ebx, esp
  loc_00E11E14: mov var_7C, 00000011h
  loc_00E11E1B: mov edx, [eax]
  loc_00E11E1D: push eax
  loc_00E11E1E: mov [ebx], ecx
  loc_00E11E20: mov ecx, var_80
  loc_00E11E23: mov var_C4, eax
  loc_00E11E29: mov [ebx+00000004h], ecx
  loc_00E11E2C: mov ecx, var_7C
  loc_00E11E2F: mov [ebx+00000008h], ecx
  loc_00E11E32: mov ecx, var_78
  loc_00E11E35: mov [ebx+0000000Ch], ecx
  loc_00E11E38: call [edx+00000028h]
  loc_00E11E3B: test eax, eax
  loc_00E11E3D: fnclex
  loc_00E11E3F: jge 00E11E56h
  loc_00E11E41: mov edx, var_C4
  loc_00E11E47: push 00000028h
  loc_00E11E49: push 006E09E8h
  loc_00E11E4E: push edx
  loc_00E11E4F: push eax
  loc_00E11E50: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11E56: mov eax, var_30
  loc_00E11E59: lea edx, var_54
  loc_00E11E5C: push edx
  loc_00E11E5D: push eax
  loc_00E11E5E: mov ecx, [eax]
  loc_00E11E60: mov ebx, eax
  loc_00E11E62: call [ecx+00000034h]
  loc_00E11E65: test eax, eax
  loc_00E11E67: fnclex
  loc_00E11E69: jge 00E11E7Ah
  loc_00E11E6B: push 00000034h
  loc_00E11E6D: push 006E09F8h
  loc_00E11E72: push ebx
  loc_00E11E73: push eax
  loc_00E11E74: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11E7A: mov eax, var_D4
  loc_00E11E80: lea ecx, var_54
  loc_00E11E83: push ecx
  loc_00E11E84: mov ebx, [eax]
  loc_00E11E86: call [00401034h] ; __vbaStrVarMove
  loc_00E11E8C: mov edx, eax
  loc_00E11E8E: lea ecx, var_18
  loc_00E11E91: call [00401228h] ; __vbaStrMove
  loc_00E11E97: mov edx, ebx
  loc_00E11E99: mov ebx, var_D4
  loc_00E11E9F: push eax
  loc_00E11EA0: push ebx
  loc_00E11EA1: call [edx+000000A4h]
  loc_00E11EA7: test eax, eax
  loc_00E11EA9: fnclex
  loc_00E11EAB: jge 00E11EBFh
  loc_00E11EAD: push 000000A4h
  loc_00E11EB2: push 006DCB70h
  loc_00E11EB7: push ebx
  loc_00E11EB8: push eax
  loc_00E11EB9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11EBF: lea ecx, var_18
  loc_00E11EC2: call [00401258h] ; __vbaFreeStr
  loc_00E11EC8: lea eax, var_34
  loc_00E11ECB: lea ecx, var_30
  loc_00E11ECE: push eax
  loc_00E11ECF: lea edx, var_2C
  loc_00E11ED2: push ecx
  loc_00E11ED3: lea eax, var_28
  loc_00E11ED6: push edx
  loc_00E11ED7: lea ecx, var_24
  loc_00E11EDA: push eax
  loc_00E11EDB: push ecx
  loc_00E11EDC: push 00000005h
  loc_00E11EDE: call [00401048h] ; __vbaFreeObjList
  loc_00E11EE4: lea edx, var_54
  loc_00E11EE7: lea eax, var_44
  loc_00E11EEA: push edx
  loc_00E11EEB: push eax
  loc_00E11EEC: push 00000002h
  loc_00E11EEE: call [00401038h] ; __vbaFreeVarList
  loc_00E11EF4: mov ecx, [esi]
  loc_00E11EF6: add esp, 00000024h
  loc_00E11EF9: push esi
  loc_00E11EFA: call [ecx+00000348h]
  loc_00E11F00: lea edx, var_34
  loc_00E11F03: push eax
  loc_00E11F04: push edx
  loc_00E11F05: call edi
  loc_00E11F07: push 006DCBD8h
  loc_00E11F0C: mov var_D4, eax
  loc_00E11F12: mov eax, [esi]
  loc_00E11F14: push 00000000h
  loc_00E11F16: push 00000018h
  loc_00E11F18: push esi
  loc_00E11F19: call [eax+00000490h]
  loc_00E11F1F: lea ecx, var_24
  loc_00E11F22: push eax
  loc_00E11F23: push ecx
  loc_00E11F24: call edi
  loc_00E11F26: lea edx, var_44
  loc_00E11F29: push eax
  loc_00E11F2A: push edx
  loc_00E11F2B: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E11F31: add esp, 00000010h
  loc_00E11F34: push eax
  loc_00E11F35: call [00401120h] ; __vbaCastObjVar
  loc_00E11F3B: push eax
  loc_00E11F3C: lea eax, var_28
  loc_00E11F3F: push eax
  loc_00E11F40: call edi
  loc_00E11F42: mov ebx, eax
  loc_00E11F44: lea edx, var_2C
  loc_00E11F47: push edx
  loc_00E11F48: push ebx
  loc_00E11F49: mov ecx, [ebx]
  loc_00E11F4B: call [ecx+00000054h]
  loc_00E11F4E: test eax, eax
  loc_00E11F50: fnclex
  loc_00E11F52: jge 00E11F63h
  loc_00E11F54: push 00000054h
  loc_00E11F56: push 006DCBE8h
  loc_00E11F5B: push ebx
  loc_00E11F5C: push eax
  loc_00E11F5D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11F63: lea ebx, var_30
  loc_00E11F66: mov eax, var_2C
  loc_00E11F69: push ebx
  loc_00E11F6A: mov ecx, 00000002h
  loc_00E11F6F: sub esp, 00000010h
  loc_00E11F72: mov var_84, ecx
  loc_00E11F78: mov ebx, esp
  loc_00E11F7A: mov var_7C, 00000012h
  loc_00E11F81: mov edx, [eax]
  loc_00E11F83: push eax
  loc_00E11F84: mov [ebx], ecx
  loc_00E11F86: mov ecx, var_80
  loc_00E11F89: mov var_C4, eax
  loc_00E11F8F: mov [ebx+00000004h], ecx
  loc_00E11F92: mov ecx, var_7C
  loc_00E11F95: mov [ebx+00000008h], ecx
  loc_00E11F98: mov ecx, var_78
  loc_00E11F9B: mov [ebx+0000000Ch], ecx
  loc_00E11F9E: call [edx+00000028h]
  loc_00E11FA1: test eax, eax
  loc_00E11FA3: fnclex
  loc_00E11FA5: jge 00E11FBCh
  loc_00E11FA7: mov edx, var_C4
  loc_00E11FAD: push 00000028h
  loc_00E11FAF: push 006E09E8h
  loc_00E11FB4: push edx
  loc_00E11FB5: push eax
  loc_00E11FB6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11FBC: mov eax, var_30
  loc_00E11FBF: lea edx, var_54
  loc_00E11FC2: push edx
  loc_00E11FC3: push eax
  loc_00E11FC4: mov ecx, [eax]
  loc_00E11FC6: mov ebx, eax
  loc_00E11FC8: call [ecx+00000034h]
  loc_00E11FCB: test eax, eax
  loc_00E11FCD: fnclex
  loc_00E11FCF: jge 00E11FE0h
  loc_00E11FD1: push 00000034h
  loc_00E11FD3: push 006E09F8h
  loc_00E11FD8: push ebx
  loc_00E11FD9: push eax
  loc_00E11FDA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E11FE0: mov eax, var_D4
  loc_00E11FE6: lea ecx, var_54
  loc_00E11FE9: push ecx
  loc_00E11FEA: mov ebx, [eax]
  loc_00E11FEC: call [00401034h] ; __vbaStrVarMove
  loc_00E11FF2: mov edx, eax
  loc_00E11FF4: lea ecx, var_18
  loc_00E11FF7: call [00401228h] ; __vbaStrMove
  loc_00E11FFD: mov edx, ebx
  loc_00E11FFF: mov ebx, var_D4
  loc_00E12005: push eax
  loc_00E12006: push ebx
  loc_00E12007: call [edx+000000A4h]
  loc_00E1200D: test eax, eax
  loc_00E1200F: fnclex
  loc_00E12011: jge 00E139EEh
  loc_00E12017: jmp 00E139DCh
  loc_00E1201C: mov ecx, [esi]
  loc_00E1201E: push esi
  loc_00E1201F: call [ecx+0000031Ch]
  loc_00E12025: lea edx, var_24
  loc_00E12028: push eax
  loc_00E12029: push edx
  loc_00E1202A: call edi
  loc_00E1202C: mov ecx, [eax]
  loc_00E1202E: lea edx, var_B8
  loc_00E12034: push edx
  loc_00E12035: push eax
  loc_00E12036: mov var_BC, eax
  loc_00E1203C: call [ecx+000000E0h]
  loc_00E12042: test eax, eax
  loc_00E12044: fnclex
  loc_00E12046: jge 00E12060h
  loc_00E12048: mov ecx, var_BC
  loc_00E1204E: push 000000E0h
  loc_00E12053: push 006E03D4h
  loc_00E12058: push ecx
  loc_00E12059: push eax
  loc_00E1205A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12060: mov dx, var_B8
  loc_00E12067: lea ecx, var_24
  loc_00E1206A: mov var_C4, dx
  loc_00E12071: call ebx
  loc_00E12073: cmp var_C4, 0000h
  loc_00E1207B: jz 00E13A52h
  loc_00E12081: mov eax, [esi]
  loc_00E12083: push esi
  loc_00E12084: call [eax+00000324h]
  loc_00E1208A: lea ecx, var_24
  loc_00E1208D: push eax
  loc_00E1208E: push ecx
  loc_00E1208F: call edi
  loc_00E12091: mov edx, [eax]
  loc_00E12093: lea ecx, var_18
  loc_00E12096: push ecx
  loc_00E12097: push eax
  loc_00E12098: mov var_BC, eax
  loc_00E1209E: call [edx+000000A0h]
  loc_00E120A4: test eax, eax
  loc_00E120A6: fnclex
  loc_00E120A8: jge 00E120C2h
  loc_00E120AA: mov edx, var_BC
  loc_00E120B0: push 000000A0h
  loc_00E120B5: push 006DCB70h
  loc_00E120BA: push edx
  loc_00E120BB: push eax
  loc_00E120BC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E120C2: mov eax, var_18
  loc_00E120C5: push 006E0D20h ; "select * from biodata where nama_ibu like '" & Chr(37)
  loc_00E120CA: push eax
  loc_00E120CB: call [0040106Ch] ; __vbaStrCat
  loc_00E120D1: mov edx, eax
  loc_00E120D3: lea ecx, var_1C
  loc_00E120D6: call [00401228h] ; __vbaStrMove
  loc_00E120DC: push eax
  loc_00E120DD: push 006E0D80h ; Chr(37) & "' order by nama_ibu asc"
  loc_00E120E2: call [0040106Ch] ; __vbaStrCat
  loc_00E120E8: mov edx, eax
  loc_00E120EA: lea ecx, var_20
  loc_00E120ED: call [00401228h] ; __vbaStrMove
  loc_00E120F3: mov edx, eax
  loc_00E120F5: lea ecx, [esi+00000034h]
  loc_00E120F8: call [004011B0h] ; __vbaStrCopy
  loc_00E120FE: lea ecx, var_20
  loc_00E12101: lea edx, var_1C
  loc_00E12104: push ecx
  loc_00E12105: lea eax, var_18
  loc_00E12108: push edx
  loc_00E12109: push eax
  loc_00E1210A: push 00000003h
  loc_00E1210C: call [004011BCh] ; __vbaFreeStrList
  loc_00E12112: add esp, 00000010h
  loc_00E12115: lea ecx, var_24
  loc_00E12118: call ebx
  loc_00E1211A: sub esp, 00000010h
  loc_00E1211D: mov ecx, 00004008h
  loc_00E12122: mov edx, esp
  loc_00E12124: mov var_84, ecx
  loc_00E1212A: lea eax, [esi+00000034h]
  loc_00E1212D: push 0000000Eh
  loc_00E1212F: mov [edx], ecx
  loc_00E12131: mov ecx, var_80
  loc_00E12134: mov var_7C, eax
  loc_00E12137: push esi
  loc_00E12138: mov [edx+00000004h], ecx
  loc_00E1213B: mov ecx, [esi]
  loc_00E1213D: mov [edx+00000008h], eax
  loc_00E12140: mov eax, var_78
  loc_00E12143: mov [edx+0000000Ch], eax
  loc_00E12146: call [ecx+00000490h]
  loc_00E1214C: lea edx, var_24
  loc_00E1214F: push eax
  loc_00E12150: push edx
  loc_00E12151: call edi
  loc_00E12153: push eax
  loc_00E12154: call [00401238h] ; __vbaLateIdSt
  loc_00E1215A: lea ecx, var_24
  loc_00E1215D: call ebx
  loc_00E1215F: mov eax, [esi]
  loc_00E12161: push 00000000h
  loc_00E12163: push 00000019h
  loc_00E12165: push esi
  loc_00E12166: call [eax+00000490h]
  loc_00E1216C: lea ecx, var_24
  loc_00E1216F: push eax
  loc_00E12170: push ecx
  loc_00E12171: call edi
  loc_00E12173: push eax
  loc_00E12174: call [00401030h] ; __vbaLateIdCall
  loc_00E1217A: add esp, 0000000Ch
  loc_00E1217D: lea ecx, var_24
  loc_00E12180: call ebx
  loc_00E12182: mov edx, [esi]
  loc_00E12184: push 006DCBD8h
  loc_00E12189: push 00000000h
  loc_00E1218B: push 00000018h
  loc_00E1218D: push esi
  loc_00E1218E: call [edx+00000490h]
  loc_00E12194: push eax
  loc_00E12195: lea eax, var_24
  loc_00E12198: push eax
  loc_00E12199: call edi
  loc_00E1219B: lea ecx, var_44
  loc_00E1219E: push eax
  loc_00E1219F: push ecx
  loc_00E121A0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E121A6: add esp, 00000010h
  loc_00E121A9: push eax
  loc_00E121AA: call [00401120h] ; __vbaCastObjVar
  loc_00E121B0: lea edx, var_28
  loc_00E121B3: push eax
  loc_00E121B4: push edx
  loc_00E121B5: call edi
  loc_00E121B7: mov ebx, eax
  loc_00E121B9: lea ecx, var_B8
  loc_00E121BF: push ecx
  loc_00E121C0: push ebx
  loc_00E121C1: mov eax, [ebx]
  loc_00E121C3: call [eax+00000050h]
  loc_00E121C6: test eax, eax
  loc_00E121C8: fnclex
  loc_00E121CA: jge 00E121DBh
  loc_00E121CC: push 00000050h
  loc_00E121CE: push 006DCBE8h
  loc_00E121D3: push ebx
  loc_00E121D4: push eax
  loc_00E121D5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E121DB: mov bx, var_B8
  loc_00E121E2: lea edx, var_28
  loc_00E121E5: lea eax, var_24
  loc_00E121E8: push edx
  loc_00E121E9: push eax
  loc_00E121EA: push 00000002h
  loc_00E121EC: call [00401048h] ; __vbaFreeObjList
  loc_00E121F2: add esp, 0000000Ch
  loc_00E121F5: lea ecx, var_44
  loc_00E121F8: call [00401024h] ; __vbaFreeVar
  loc_00E121FE: test bx, bx
  loc_00E12201: jz 00E12293h
  loc_00E12207: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E1220D: mov ecx, 0000000Ah
  loc_00E12212: mov eax, 80020004h
  loc_00E12217: mov var_74, ecx
  loc_00E1221A: mov var_64, ecx
  loc_00E1221D: mov edi, 00000008h
  loc_00E12222: lea edx, var_94
  loc_00E12228: lea ecx, var_54
  loc_00E1222B: mov var_6C, eax
  loc_00E1222E: mov var_5C, eax
  loc_00E12231: mov var_8C, 006E0D04h ; "Not Found !"
  loc_00E1223B: mov var_94, edi
  loc_00E12241: call __vbaVarDup
  loc_00E12243: lea edx, var_84
  loc_00E12249: lea ecx, var_44
  loc_00E1224C: mov var_7C, 006E0CE0h ; "Data Tidak Ada"
  loc_00E12253: mov var_84, edi
  loc_00E12259: call __vbaVarDup
  loc_00E1225B: lea ecx, var_74
  loc_00E1225E: lea edx, var_64
  loc_00E12261: push ecx
  loc_00E12262: lea eax, var_54
  loc_00E12265: push edx
  loc_00E12266: push eax
  loc_00E12267: lea ecx, var_44
  loc_00E1226A: push 00000010h
  loc_00E1226C: push ecx
  loc_00E1226D: call [004010A8h] ; rtcMsgBox
  loc_00E12273: lea edx, var_74
  loc_00E12276: lea eax, var_64
  loc_00E12279: push edx
  loc_00E1227A: lea ecx, var_54
  loc_00E1227D: push eax
  loc_00E1227E: lea edx, var_44
  loc_00E12281: push ecx
  loc_00E12282: push edx
  loc_00E12283: push 00000004h
  loc_00E12285: call [00401038h] ; __vbaFreeVarList
  loc_00E1228B: add esp, 00000014h
  loc_00E1228E: jmp 00E159A6h
  loc_00E12293: mov eax, [esi]
  loc_00E12295: push esi
  loc_00E12296: call [eax+0000039Ch]
  loc_00E1229C: lea ecx, var_34
  loc_00E1229F: push eax
  loc_00E122A0: push ecx
  loc_00E122A1: call edi
  loc_00E122A3: mov edx, [esi]
  loc_00E122A5: push 006DCBD8h
  loc_00E122AA: push 00000000h
  loc_00E122AC: push 00000018h
  loc_00E122AE: push esi
  loc_00E122AF: mov var_D4, eax
  loc_00E122B5: call [edx+00000490h]
  loc_00E122BB: push eax
  loc_00E122BC: lea eax, var_24
  loc_00E122BF: push eax
  loc_00E122C0: call edi
  loc_00E122C2: lea ecx, var_44
  loc_00E122C5: push eax
  loc_00E122C6: push ecx
  loc_00E122C7: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E122CD: add esp, 00000010h
  loc_00E122D0: push eax
  loc_00E122D1: call [00401120h] ; __vbaCastObjVar
  loc_00E122D7: lea edx, var_28
  loc_00E122DA: push eax
  loc_00E122DB: push edx
  loc_00E122DC: call edi
  loc_00E122DE: mov ebx, eax
  loc_00E122E0: lea ecx, var_2C
  loc_00E122E3: push ecx
  loc_00E122E4: push ebx
  loc_00E122E5: mov eax, [ebx]
  loc_00E122E7: call [eax+00000054h]
  loc_00E122EA: test eax, eax
  loc_00E122EC: fnclex
  loc_00E122EE: jge 00E122FFh
  loc_00E122F0: push 00000054h
  loc_00E122F2: push 006DCBE8h
  loc_00E122F7: push ebx
  loc_00E122F8: push eax
  loc_00E122F9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E122FF: lea ebx, var_30
  loc_00E12302: mov eax, var_2C
  loc_00E12305: push ebx
  loc_00E12306: mov ecx, 00000002h
  loc_00E1230B: sub esp, 00000010h
  loc_00E1230E: mov var_84, ecx
  loc_00E12314: mov ebx, esp
  loc_00E12316: mov var_7C, 00000000h
  loc_00E1231D: mov edx, [eax]
  loc_00E1231F: push eax
  loc_00E12320: mov [ebx], ecx
  loc_00E12322: mov ecx, var_80
  loc_00E12325: mov var_C4, eax
  loc_00E1232B: mov [ebx+00000004h], ecx
  loc_00E1232E: mov ecx, var_7C
  loc_00E12331: mov [ebx+00000008h], ecx
  loc_00E12334: mov ecx, var_78
  loc_00E12337: mov [ebx+0000000Ch], ecx
  loc_00E1233A: call [edx+00000028h]
  loc_00E1233D: test eax, eax
  loc_00E1233F: fnclex
  loc_00E12341: jge 00E12358h
  loc_00E12343: mov edx, var_C4
  loc_00E12349: push 00000028h
  loc_00E1234B: push 006E09E8h
  loc_00E12350: push edx
  loc_00E12351: push eax
  loc_00E12352: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12358: mov eax, var_30
  loc_00E1235B: lea edx, var_54
  loc_00E1235E: push edx
  loc_00E1235F: push eax
  loc_00E12360: mov ecx, [eax]
  loc_00E12362: mov ebx, eax
  loc_00E12364: call [ecx+00000034h]
  loc_00E12367: test eax, eax
  loc_00E12369: fnclex
  loc_00E1236B: jge 00E1237Ch
  loc_00E1236D: push 00000034h
  loc_00E1236F: push 006E09F8h
  loc_00E12374: push ebx
  loc_00E12375: push eax
  loc_00E12376: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1237C: mov eax, var_D4
  loc_00E12382: lea ecx, var_54
  loc_00E12385: push ecx
  loc_00E12386: mov ebx, [eax]
  loc_00E12388: call [00401034h] ; __vbaStrVarMove
  loc_00E1238E: mov edx, eax
  loc_00E12390: lea ecx, var_18
  loc_00E12393: call [00401228h] ; __vbaStrMove
  loc_00E12399: mov edx, ebx
  loc_00E1239B: mov ebx, var_D4
  loc_00E123A1: push eax
  loc_00E123A2: push ebx
  loc_00E123A3: call [edx+000000A4h]
  loc_00E123A9: test eax, eax
  loc_00E123AB: fnclex
  loc_00E123AD: jge 00E123C1h
  loc_00E123AF: push 000000A4h
  loc_00E123B4: push 006DCB70h
  loc_00E123B9: push ebx
  loc_00E123BA: push eax
  loc_00E123BB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E123C1: lea ecx, var_18
  loc_00E123C4: call [00401258h] ; __vbaFreeStr
  loc_00E123CA: lea eax, var_34
  loc_00E123CD: lea ecx, var_30
  loc_00E123D0: push eax
  loc_00E123D1: lea edx, var_2C
  loc_00E123D4: push ecx
  loc_00E123D5: lea eax, var_28
  loc_00E123D8: push edx
  loc_00E123D9: lea ecx, var_24
  loc_00E123DC: push eax
  loc_00E123DD: push ecx
  loc_00E123DE: push 00000005h
  loc_00E123E0: call [00401048h] ; __vbaFreeObjList
  loc_00E123E6: lea edx, var_54
  loc_00E123E9: lea eax, var_44
  loc_00E123EC: push edx
  loc_00E123ED: push eax
  loc_00E123EE: push 00000002h
  loc_00E123F0: call [00401038h] ; __vbaFreeVarList
  loc_00E123F6: mov ecx, [esi]
  loc_00E123F8: add esp, 00000024h
  loc_00E123FB: push esi
  loc_00E123FC: call [ecx+00000394h]
  loc_00E12402: lea edx, var_34
  loc_00E12405: push eax
  loc_00E12406: push edx
  loc_00E12407: call edi
  loc_00E12409: push 006DCBD8h
  loc_00E1240E: mov var_D4, eax
  loc_00E12414: mov eax, [esi]
  loc_00E12416: push 00000000h
  loc_00E12418: push 00000018h
  loc_00E1241A: push esi
  loc_00E1241B: call [eax+00000490h]
  loc_00E12421: lea ecx, var_24
  loc_00E12424: push eax
  loc_00E12425: push ecx
  loc_00E12426: call edi
  loc_00E12428: lea edx, var_44
  loc_00E1242B: push eax
  loc_00E1242C: push edx
  loc_00E1242D: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E12433: add esp, 00000010h
  loc_00E12436: push eax
  loc_00E12437: call [00401120h] ; __vbaCastObjVar
  loc_00E1243D: push eax
  loc_00E1243E: lea eax, var_28
  loc_00E12441: push eax
  loc_00E12442: call edi
  loc_00E12444: mov ebx, eax
  loc_00E12446: lea edx, var_2C
  loc_00E12449: push edx
  loc_00E1244A: push ebx
  loc_00E1244B: mov ecx, [ebx]
  loc_00E1244D: call [ecx+00000054h]
  loc_00E12450: test eax, eax
  loc_00E12452: fnclex
  loc_00E12454: jge 00E12465h
  loc_00E12456: push 00000054h
  loc_00E12458: push 006DCBE8h
  loc_00E1245D: push ebx
  loc_00E1245E: push eax
  loc_00E1245F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12465: lea ebx, var_30
  loc_00E12468: mov eax, var_2C
  loc_00E1246B: push ebx
  loc_00E1246C: mov ecx, 00000002h
  loc_00E12471: sub esp, 00000010h
  loc_00E12474: mov var_84, ecx
  loc_00E1247A: mov ebx, esp
  loc_00E1247C: mov var_7C, 00000001h
  loc_00E12483: mov edx, [eax]
  loc_00E12485: push eax
  loc_00E12486: mov [ebx], ecx
  loc_00E12488: mov ecx, var_80
  loc_00E1248B: mov var_C4, eax
  loc_00E12491: mov [ebx+00000004h], ecx
  loc_00E12494: mov ecx, var_7C
  loc_00E12497: mov [ebx+00000008h], ecx
  loc_00E1249A: mov ecx, var_78
  loc_00E1249D: mov [ebx+0000000Ch], ecx
  loc_00E124A0: call [edx+00000028h]
  loc_00E124A3: test eax, eax
  loc_00E124A5: fnclex
  loc_00E124A7: jge 00E124BEh
  loc_00E124A9: mov edx, var_C4
  loc_00E124AF: push 00000028h
  loc_00E124B1: push 006E09E8h
  loc_00E124B6: push edx
  loc_00E124B7: push eax
  loc_00E124B8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E124BE: mov eax, var_30
  loc_00E124C1: lea edx, var_54
  loc_00E124C4: push edx
  loc_00E124C5: push eax
  loc_00E124C6: mov ecx, [eax]
  loc_00E124C8: mov ebx, eax
  loc_00E124CA: call [ecx+00000034h]
  loc_00E124CD: test eax, eax
  loc_00E124CF: fnclex
  loc_00E124D1: jge 00E124E2h
  loc_00E124D3: push 00000034h
  loc_00E124D5: push 006E09F8h
  loc_00E124DA: push ebx
  loc_00E124DB: push eax
  loc_00E124DC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E124E2: mov eax, var_D4
  loc_00E124E8: lea ecx, var_54
  loc_00E124EB: push ecx
  loc_00E124EC: mov ebx, [eax]
  loc_00E124EE: call [00401034h] ; __vbaStrVarMove
  loc_00E124F4: mov edx, eax
  loc_00E124F6: lea ecx, var_18
  loc_00E124F9: call [00401228h] ; __vbaStrMove
  loc_00E124FF: mov edx, ebx
  loc_00E12501: mov ebx, var_D4
  loc_00E12507: push eax
  loc_00E12508: push ebx
  loc_00E12509: call [edx+000000A4h]
  loc_00E1250F: test eax, eax
  loc_00E12511: fnclex
  loc_00E12513: jge 00E12527h
  loc_00E12515: push 000000A4h
  loc_00E1251A: push 006DCB70h
  loc_00E1251F: push ebx
  loc_00E12520: push eax
  loc_00E12521: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12527: lea ecx, var_18
  loc_00E1252A: call [00401258h] ; __vbaFreeStr
  loc_00E12530: lea eax, var_34
  loc_00E12533: lea ecx, var_30
  loc_00E12536: push eax
  loc_00E12537: lea edx, var_2C
  loc_00E1253A: push ecx
  loc_00E1253B: lea eax, var_28
  loc_00E1253E: push edx
  loc_00E1253F: lea ecx, var_24
  loc_00E12542: push eax
  loc_00E12543: push ecx
  loc_00E12544: push 00000005h
  loc_00E12546: call [00401048h] ; __vbaFreeObjList
  loc_00E1254C: lea edx, var_54
  loc_00E1254F: lea eax, var_44
  loc_00E12552: push edx
  loc_00E12553: push eax
  loc_00E12554: push 00000002h
  loc_00E12556: call [00401038h] ; __vbaFreeVarList
  loc_00E1255C: mov ecx, [esi]
  loc_00E1255E: add esp, 00000024h
  loc_00E12561: push 006DCBD8h
  loc_00E12566: push 00000000h
  loc_00E12568: push 00000018h
  loc_00E1256A: push esi
  loc_00E1256B: call [ecx+00000490h]
  loc_00E12571: lea edx, var_24
  loc_00E12574: push eax
  loc_00E12575: push edx
  loc_00E12576: call edi
  loc_00E12578: push eax
  loc_00E12579: lea eax, var_44
  loc_00E1257C: push eax
  loc_00E1257D: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E12583: add esp, 00000010h
  loc_00E12586: push eax
  loc_00E12587: call [00401120h] ; __vbaCastObjVar
  loc_00E1258D: lea ecx, var_28
  loc_00E12590: push eax
  loc_00E12591: push ecx
  loc_00E12592: call edi
  loc_00E12594: mov ebx, eax
  loc_00E12596: lea eax, var_2C
  loc_00E12599: push eax
  loc_00E1259A: push ebx
  loc_00E1259B: mov edx, [ebx]
  loc_00E1259D: call [edx+00000054h]
  loc_00E125A0: test eax, eax
  loc_00E125A2: fnclex
  loc_00E125A4: jge 00E125B5h
  loc_00E125A6: push 00000054h
  loc_00E125A8: push 006DCBE8h
  loc_00E125AD: push ebx
  loc_00E125AE: push eax
  loc_00E125AF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E125B5: lea ebx, var_30
  loc_00E125B8: mov eax, var_2C
  loc_00E125BB: push ebx
  loc_00E125BC: mov ecx, 00000002h
  loc_00E125C1: sub esp, 00000010h
  loc_00E125C4: mov var_7C, ecx
  loc_00E125C7: mov ebx, esp
  loc_00E125C9: mov var_84, ecx
  loc_00E125CF: mov edx, [eax]
  loc_00E125D1: push eax
  loc_00E125D2: mov [ebx], ecx
  loc_00E125D4: mov ecx, var_80
  loc_00E125D7: mov var_C4, eax
  loc_00E125DD: mov [ebx+00000004h], ecx
  loc_00E125E0: mov ecx, var_7C
  loc_00E125E3: mov [ebx+00000008h], ecx
  loc_00E125E6: mov ecx, var_78
  loc_00E125E9: mov [ebx+0000000Ch], ecx
  loc_00E125EC: call [edx+00000028h]
  loc_00E125EF: test eax, eax
  loc_00E125F1: fnclex
  loc_00E125F3: jge 00E1260Ah
  loc_00E125F5: mov edx, var_C4
  loc_00E125FB: push 00000028h
  loc_00E125FD: push 006E09E8h
  loc_00E12602: push edx
  loc_00E12603: push eax
  loc_00E12604: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1260A: mov eax, var_30
  loc_00E1260D: lea edx, var_54
  loc_00E12610: push edx
  loc_00E12611: push eax
  loc_00E12612: mov ecx, [eax]
  loc_00E12614: mov ebx, eax
  loc_00E12616: call [ecx+00000034h]
  loc_00E12619: test eax, eax
  loc_00E1261B: fnclex
  loc_00E1261D: jge 00E1262Eh
  loc_00E1261F: push 00000034h
  loc_00E12621: push 006E09F8h
  loc_00E12626: push ebx
  loc_00E12627: push eax
  loc_00E12628: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1262E: lea eax, var_54
  loc_00E12631: lea ecx, var_94
  loc_00E12637: push eax
  loc_00E12638: push ecx
  loc_00E12639: mov var_8C, 006E0B24h ; "Laki - laki"
  loc_00E12643: mov var_94, 00008008h
  loc_00E1264D: call [00401108h] ; __vbaVarTstEq
  loc_00E12653: mov bx, ax
  loc_00E12656: lea edx, var_30
  loc_00E12659: lea eax, var_2C
  loc_00E1265C: push edx
  loc_00E1265D: lea ecx, var_28
  loc_00E12660: push eax
  loc_00E12661: lea edx, var_24
  loc_00E12664: push ecx
  loc_00E12665: push edx
  loc_00E12666: push 00000004h
  loc_00E12668: call [00401048h] ; __vbaFreeObjList
  loc_00E1266E: lea eax, var_54
  loc_00E12671: lea ecx, var_44
  loc_00E12674: push eax
  loc_00E12675: push ecx
  loc_00E12676: push 00000002h
  loc_00E12678: call [00401038h] ; __vbaFreeVarList
  loc_00E1267E: add esp, 00000020h
  loc_00E12681: test bx, bx
  loc_00E12684: jz 00E12732h
  loc_00E1268A: mov edx, [esi]
  loc_00E1268C: push esi
  loc_00E1268D: call [edx+00000380h]
  loc_00E12693: push eax
  loc_00E12694: lea eax, var_24
  loc_00E12697: push eax
  loc_00E12698: call edi
  loc_00E1269A: mov ebx, eax
  loc_00E1269C: push FFFFFFFFh
  loc_00E1269E: push ebx
  loc_00E1269F: mov ecx, [ebx]
  loc_00E126A1: call [ecx+000000E4h]
  loc_00E126A7: test eax, eax
  loc_00E126A9: fnclex
  loc_00E126AB: jge 00E126BFh
  loc_00E126AD: push 000000E4h
  loc_00E126B2: push 006E03D4h
  loc_00E126B7: push ebx
  loc_00E126B8: push eax
  loc_00E126B9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E126BF: lea ecx, var_24
  loc_00E126C2: call [00401254h] ; __vbaFreeObj
  loc_00E126C8: mov edx, [esi]
  loc_00E126CA: push esi
  loc_00E126CB: call [edx+0000037Ch]
  loc_00E126D1: push eax
  loc_00E126D2: lea eax, var_24
  loc_00E126D5: push eax
  loc_00E126D6: call edi
  loc_00E126D8: mov ebx, eax
  loc_00E126DA: push FFFFFFFFh
  loc_00E126DC: push ebx
  loc_00E126DD: mov ecx, [ebx]
  loc_00E126DF: call [ecx+0000009Ch]
  loc_00E126E5: test eax, eax
  loc_00E126E7: fnclex
  loc_00E126E9: jge 00E126FDh
  loc_00E126EB: push 0000009Ch
  loc_00E126F0: push 006E03D4h
  loc_00E126F5: push ebx
  loc_00E126F6: push eax
  loc_00E126F7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E126FD: lea ecx, var_24
  loc_00E12700: call [00401254h] ; __vbaFreeObj
  loc_00E12706: mov edx, [esi]
  loc_00E12708: push esi
  loc_00E12709: call [edx+00000380h]
  loc_00E1270F: push eax
  loc_00E12710: lea eax, var_24
  loc_00E12713: push eax
  loc_00E12714: call edi
  loc_00E12716: mov ebx, eax
  loc_00E12718: push FFFFFFFFh
  loc_00E1271A: push ebx
  loc_00E1271B: mov ecx, [ebx]
  loc_00E1271D: call [ecx+0000009Ch]
  loc_00E12723: test eax, eax
  loc_00E12725: fnclex
  loc_00E12727: jge 00E1290Eh
  loc_00E1272D: jmp 00E128FCh
  loc_00E12732: mov edx, [esi]
  loc_00E12734: push 006DCBD8h
  loc_00E12739: push 00000000h
  loc_00E1273B: push 00000018h
  loc_00E1273D: push esi
  loc_00E1273E: call [edx+00000490h]
  loc_00E12744: push eax
  loc_00E12745: lea eax, var_24
  loc_00E12748: push eax
  loc_00E12749: call edi
  loc_00E1274B: lea ecx, var_44
  loc_00E1274E: push eax
  loc_00E1274F: push ecx
  loc_00E12750: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E12756: add esp, 00000010h
  loc_00E12759: push eax
  loc_00E1275A: call [00401120h] ; __vbaCastObjVar
  loc_00E12760: lea edx, var_28
  loc_00E12763: push eax
  loc_00E12764: push edx
  loc_00E12765: call edi
  loc_00E12767: mov ebx, eax
  loc_00E12769: lea ecx, var_2C
  loc_00E1276C: push ecx
  loc_00E1276D: push ebx
  loc_00E1276E: mov eax, [ebx]
  loc_00E12770: call [eax+00000054h]
  loc_00E12773: test eax, eax
  loc_00E12775: fnclex
  loc_00E12777: jge 00E12788h
  loc_00E12779: push 00000054h
  loc_00E1277B: push 006DCBE8h
  loc_00E12780: push ebx
  loc_00E12781: push eax
  loc_00E12782: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12788: lea ebx, var_30
  loc_00E1278B: mov eax, var_2C
  loc_00E1278E: push ebx
  loc_00E1278F: mov ecx, 00000002h
  loc_00E12794: sub esp, 00000010h
  loc_00E12797: mov var_7C, ecx
  loc_00E1279A: mov ebx, esp
  loc_00E1279C: mov var_84, ecx
  loc_00E127A2: mov edx, [eax]
  loc_00E127A4: push eax
  loc_00E127A5: mov [ebx], ecx
  loc_00E127A7: mov ecx, var_80
  loc_00E127AA: mov var_C4, eax
  loc_00E127B0: mov [ebx+00000004h], ecx
  loc_00E127B3: mov ecx, var_7C
  loc_00E127B6: mov [ebx+00000008h], ecx
  loc_00E127B9: mov ecx, var_78
  loc_00E127BC: mov [ebx+0000000Ch], ecx
  loc_00E127BF: call [edx+00000028h]
  loc_00E127C2: test eax, eax
  loc_00E127C4: fnclex
  loc_00E127C6: jge 00E127DDh
  loc_00E127C8: mov edx, var_C4
  loc_00E127CE: push 00000028h
  loc_00E127D0: push 006E09E8h
  loc_00E127D5: push edx
  loc_00E127D6: push eax
  loc_00E127D7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E127DD: mov eax, var_30
  loc_00E127E0: lea edx, var_54
  loc_00E127E3: push edx
  loc_00E127E4: push eax
  loc_00E127E5: mov ecx, [eax]
  loc_00E127E7: mov ebx, eax
  loc_00E127E9: call [ecx+00000034h]
  loc_00E127EC: test eax, eax
  loc_00E127EE: fnclex
  loc_00E127F0: jge 00E12801h
  loc_00E127F2: push 00000034h
  loc_00E127F4: push 006E09F8h
  loc_00E127F9: push ebx
  loc_00E127FA: push eax
  loc_00E127FB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12801: lea eax, var_54
  loc_00E12804: lea ecx, var_94
  loc_00E1280A: push eax
  loc_00E1280B: push ecx
  loc_00E1280C: mov var_8C, 006E0B40h ; "Perempuan"
  loc_00E12816: mov var_94, 00008008h
  loc_00E12820: call [00401108h] ; __vbaVarTstEq
  loc_00E12826: mov bx, ax
  loc_00E12829: lea edx, var_30
  loc_00E1282C: lea eax, var_2C
  loc_00E1282F: push edx
  loc_00E12830: lea ecx, var_28
  loc_00E12833: push eax
  loc_00E12834: lea edx, var_24
  loc_00E12837: push ecx
  loc_00E12838: push edx
  loc_00E12839: push 00000004h
  loc_00E1283B: call [00401048h] ; __vbaFreeObjList
  loc_00E12841: lea eax, var_54
  loc_00E12844: lea ecx, var_44
  loc_00E12847: push eax
  loc_00E12848: push ecx
  loc_00E12849: push 00000002h
  loc_00E1284B: call [00401038h] ; __vbaFreeVarList
  loc_00E12851: add esp, 00000020h
  loc_00E12854: test bx, bx
  loc_00E12857: jz 00E1291Ch
  loc_00E1285D: mov edx, [esi]
  loc_00E1285F: push esi
  loc_00E12860: call [edx+0000037Ch]
  loc_00E12866: push eax
  loc_00E12867: lea eax, var_24
  loc_00E1286A: push eax
  loc_00E1286B: call edi
  loc_00E1286D: mov ebx, eax
  loc_00E1286F: push FFFFFFFFh
  loc_00E12871: push ebx
  loc_00E12872: mov ecx, [ebx]
  loc_00E12874: call [ecx+000000E4h]
  loc_00E1287A: test eax, eax
  loc_00E1287C: fnclex
  loc_00E1287E: jge 00E12892h
  loc_00E12880: push 000000E4h
  loc_00E12885: push 006E03D4h
  loc_00E1288A: push ebx
  loc_00E1288B: push eax
  loc_00E1288C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12892: lea ecx, var_24
  loc_00E12895: call [00401254h] ; __vbaFreeObj
  loc_00E1289B: mov edx, [esi]
  loc_00E1289D: push esi
  loc_00E1289E: call [edx+0000037Ch]
  loc_00E128A4: push eax
  loc_00E128A5: lea eax, var_24
  loc_00E128A8: push eax
  loc_00E128A9: call edi
  loc_00E128AB: mov ebx, eax
  loc_00E128AD: push FFFFFFFFh
  loc_00E128AF: push ebx
  loc_00E128B0: mov ecx, [ebx]
  loc_00E128B2: call [ecx+0000009Ch]
  loc_00E128B8: test eax, eax
  loc_00E128BA: fnclex
  loc_00E128BC: jge 00E128D0h
  loc_00E128BE: push 0000009Ch
  loc_00E128C3: push 006E03D4h
  loc_00E128C8: push ebx
  loc_00E128C9: push eax
  loc_00E128CA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E128D0: lea ecx, var_24
  loc_00E128D3: call [00401254h] ; __vbaFreeObj
  loc_00E128D9: mov edx, [esi]
  loc_00E128DB: push esi
  loc_00E128DC: call [edx+00000380h]
  loc_00E128E2: push eax
  loc_00E128E3: lea eax, var_24
  loc_00E128E6: push eax
  loc_00E128E7: call edi
  loc_00E128E9: mov ebx, eax
  loc_00E128EB: push FFFFFFFFh
  loc_00E128ED: push ebx
  loc_00E128EE: mov ecx, [ebx]
  loc_00E128F0: call [ecx+0000009Ch]
  loc_00E128F6: test eax, eax
  loc_00E128F8: fnclex
  loc_00E128FA: jge 00E1290Eh
  loc_00E128FC: push 0000009Ch
  loc_00E12901: push 006E03D4h
  loc_00E12906: push ebx
  loc_00E12907: push eax
  loc_00E12908: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1290E: lea ecx, var_24
  loc_00E12911: call [00401254h] ; __vbaFreeObj
  loc_00E12917: jmp 00E129A6h
  loc_00E1291C: mov ebx, [004011F0h] ; __vbaVarDup
  loc_00E12922: mov ecx, 0000000Ah
  loc_00E12927: mov eax, 80020004h
  loc_00E1292C: mov var_74, ecx
  loc_00E1292F: mov var_64, ecx
  loc_00E12932: lea edx, var_94
  loc_00E12938: lea ecx, var_54
  loc_00E1293B: mov var_6C, eax
  loc_00E1293E: mov var_5C, eax
  loc_00E12941: mov var_8C, 006E07B8h ; "ERROR 001"
  loc_00E1294B: mov var_94, 00000008h
  loc_00E12955: call ebx
  loc_00E12957: lea edx, var_84
  loc_00E1295D: lea ecx, var_44
  loc_00E12960: mov var_7C, 006E0790h ; "Ada Yang Error !"
  loc_00E12967: mov var_84, 00000008h
  loc_00E12971: call ebx
  loc_00E12973: lea edx, var_74
  loc_00E12976: lea eax, var_64
  loc_00E12979: push edx
  loc_00E1297A: lea ecx, var_54
  loc_00E1297D: push eax
  loc_00E1297E: push ecx
  loc_00E1297F: lea edx, var_44
  loc_00E12982: push 00000040h
  loc_00E12984: push edx
  loc_00E12985: call [004010A8h] ; rtcMsgBox
  loc_00E1298B: lea eax, var_74
  loc_00E1298E: lea ecx, var_64
  loc_00E12991: push eax
  loc_00E12992: lea edx, var_54
  loc_00E12995: push ecx
  loc_00E12996: lea eax, var_44
  loc_00E12999: push edx
  loc_00E1299A: push eax
  loc_00E1299B: push 00000004h
  loc_00E1299D: call [00401038h] ; __vbaFreeVarList
  loc_00E129A3: add esp, 00000014h
  loc_00E129A6: mov ecx, [esi]
  loc_00E129A8: push esi
  loc_00E129A9: call [ecx+00000390h]
  loc_00E129AF: lea edx, var_34
  loc_00E129B2: push eax
  loc_00E129B3: push edx
  loc_00E129B4: call edi
  loc_00E129B6: push 006DCBD8h
  loc_00E129BB: mov var_D4, eax
  loc_00E129C1: mov eax, [esi]
  loc_00E129C3: push 00000000h
  loc_00E129C5: push 00000018h
  loc_00E129C7: push esi
  loc_00E129C8: call [eax+00000490h]
  loc_00E129CE: lea ecx, var_24
  loc_00E129D1: push eax
  loc_00E129D2: push ecx
  loc_00E129D3: call edi
  loc_00E129D5: lea edx, var_44
  loc_00E129D8: push eax
  loc_00E129D9: push edx
  loc_00E129DA: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E129E0: add esp, 00000010h
  loc_00E129E3: push eax
  loc_00E129E4: call [00401120h] ; __vbaCastObjVar
  loc_00E129EA: push eax
  loc_00E129EB: lea eax, var_28
  loc_00E129EE: push eax
  loc_00E129EF: call edi
  loc_00E129F1: mov ebx, eax
  loc_00E129F3: lea edx, var_2C
  loc_00E129F6: push edx
  loc_00E129F7: push ebx
  loc_00E129F8: mov ecx, [ebx]
  loc_00E129FA: call [ecx+00000054h]
  loc_00E129FD: test eax, eax
  loc_00E129FF: fnclex
  loc_00E12A01: jge 00E12A12h
  loc_00E12A03: push 00000054h
  loc_00E12A05: push 006DCBE8h
  loc_00E12A0A: push ebx
  loc_00E12A0B: push eax
  loc_00E12A0C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12A12: lea ebx, var_30
  loc_00E12A15: mov eax, var_2C
  loc_00E12A18: push ebx
  loc_00E12A19: mov ecx, 00000002h
  loc_00E12A1E: sub esp, 00000010h
  loc_00E12A21: mov var_84, ecx
  loc_00E12A27: mov ebx, esp
  loc_00E12A29: mov var_7C, 00000003h
  loc_00E12A30: mov edx, [eax]
  loc_00E12A32: push eax
  loc_00E12A33: mov [ebx], ecx
  loc_00E12A35: mov ecx, var_80
  loc_00E12A38: mov var_C4, eax
  loc_00E12A3E: mov [ebx+00000004h], ecx
  loc_00E12A41: mov ecx, var_7C
  loc_00E12A44: mov [ebx+00000008h], ecx
  loc_00E12A47: mov ecx, var_78
  loc_00E12A4A: mov [ebx+0000000Ch], ecx
  loc_00E12A4D: call [edx+00000028h]
  loc_00E12A50: test eax, eax
  loc_00E12A52: fnclex
  loc_00E12A54: jge 00E12A6Bh
  loc_00E12A56: mov edx, var_C4
  loc_00E12A5C: push 00000028h
  loc_00E12A5E: push 006E09E8h
  loc_00E12A63: push edx
  loc_00E12A64: push eax
  loc_00E12A65: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12A6B: mov eax, var_30
  loc_00E12A6E: lea edx, var_54
  loc_00E12A71: push edx
  loc_00E12A72: push eax
  loc_00E12A73: mov ecx, [eax]
  loc_00E12A75: mov ebx, eax
  loc_00E12A77: call [ecx+00000034h]
  loc_00E12A7A: test eax, eax
  loc_00E12A7C: fnclex
  loc_00E12A7E: jge 00E12A8Fh
  loc_00E12A80: push 00000034h
  loc_00E12A82: push 006E09F8h
  loc_00E12A87: push ebx
  loc_00E12A88: push eax
  loc_00E12A89: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12A8F: mov eax, var_D4
  loc_00E12A95: lea ecx, var_54
  loc_00E12A98: push ecx
  loc_00E12A99: mov ebx, [eax]
  loc_00E12A9B: call [00401034h] ; __vbaStrVarMove
  loc_00E12AA1: mov edx, eax
  loc_00E12AA3: lea ecx, var_18
  loc_00E12AA6: call [00401228h] ; __vbaStrMove
  loc_00E12AAC: mov edx, ebx
  loc_00E12AAE: mov ebx, var_D4
  loc_00E12AB4: push eax
  loc_00E12AB5: push ebx
  loc_00E12AB6: call [edx+000000A4h]
  loc_00E12ABC: test eax, eax
  loc_00E12ABE: fnclex
  loc_00E12AC0: jge 00E12AD4h
  loc_00E12AC2: push 000000A4h
  loc_00E12AC7: push 006DCB70h
  loc_00E12ACC: push ebx
  loc_00E12ACD: push eax
  loc_00E12ACE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12AD4: lea ecx, var_18
  loc_00E12AD7: call [00401258h] ; __vbaFreeStr
  loc_00E12ADD: lea eax, var_34
  loc_00E12AE0: lea ecx, var_30
  loc_00E12AE3: push eax
  loc_00E12AE4: lea edx, var_2C
  loc_00E12AE7: push ecx
  loc_00E12AE8: lea eax, var_28
  loc_00E12AEB: push edx
  loc_00E12AEC: lea ecx, var_24
  loc_00E12AEF: push eax
  loc_00E12AF0: push ecx
  loc_00E12AF1: push 00000005h
  loc_00E12AF3: call [00401048h] ; __vbaFreeObjList
  loc_00E12AF9: lea edx, var_54
  loc_00E12AFC: lea eax, var_44
  loc_00E12AFF: push edx
  loc_00E12B00: push eax
  loc_00E12B01: push 00000002h
  loc_00E12B03: call [00401038h] ; __vbaFreeVarList
  loc_00E12B09: mov ecx, [esi]
  loc_00E12B0B: add esp, 00000024h
  loc_00E12B0E: push esi
  loc_00E12B0F: call [ecx+0000038Ch]
  loc_00E12B15: lea edx, var_34
  loc_00E12B18: push eax
  loc_00E12B19: push edx
  loc_00E12B1A: call edi
  loc_00E12B1C: push 006DCBD8h
  loc_00E12B21: mov var_D4, eax
  loc_00E12B27: mov eax, [esi]
  loc_00E12B29: push 00000000h
  loc_00E12B2B: push 00000018h
  loc_00E12B2D: push esi
  loc_00E12B2E: call [eax+00000490h]
  loc_00E12B34: lea ecx, var_24
  loc_00E12B37: push eax
  loc_00E12B38: push ecx
  loc_00E12B39: call edi
  loc_00E12B3B: lea edx, var_44
  loc_00E12B3E: push eax
  loc_00E12B3F: push edx
  loc_00E12B40: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E12B46: add esp, 00000010h
  loc_00E12B49: push eax
  loc_00E12B4A: call [00401120h] ; __vbaCastObjVar
  loc_00E12B50: push eax
  loc_00E12B51: lea eax, var_28
  loc_00E12B54: push eax
  loc_00E12B55: call edi
  loc_00E12B57: mov ebx, eax
  loc_00E12B59: lea edx, var_2C
  loc_00E12B5C: push edx
  loc_00E12B5D: push ebx
  loc_00E12B5E: mov ecx, [ebx]
  loc_00E12B60: call [ecx+00000054h]
  loc_00E12B63: test eax, eax
  loc_00E12B65: fnclex
  loc_00E12B67: jge 00E12B78h
  loc_00E12B69: push 00000054h
  loc_00E12B6B: push 006DCBE8h
  loc_00E12B70: push ebx
  loc_00E12B71: push eax
  loc_00E12B72: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12B78: lea ebx, var_30
  loc_00E12B7B: mov eax, var_2C
  loc_00E12B7E: push ebx
  loc_00E12B7F: mov ecx, 00000002h
  loc_00E12B84: sub esp, 00000010h
  loc_00E12B87: mov var_84, ecx
  loc_00E12B8D: mov ebx, esp
  loc_00E12B8F: mov var_7C, 00000004h
  loc_00E12B96: mov edx, [eax]
  loc_00E12B98: push eax
  loc_00E12B99: mov [ebx], ecx
  loc_00E12B9B: mov ecx, var_80
  loc_00E12B9E: mov var_C4, eax
  loc_00E12BA4: mov [ebx+00000004h], ecx
  loc_00E12BA7: mov ecx, var_7C
  loc_00E12BAA: mov [ebx+00000008h], ecx
  loc_00E12BAD: mov ecx, var_78
  loc_00E12BB0: mov [ebx+0000000Ch], ecx
  loc_00E12BB3: call [edx+00000028h]
  loc_00E12BB6: test eax, eax
  loc_00E12BB8: fnclex
  loc_00E12BBA: jge 00E12BD1h
  loc_00E12BBC: mov edx, var_C4
  loc_00E12BC2: push 00000028h
  loc_00E12BC4: push 006E09E8h
  loc_00E12BC9: push edx
  loc_00E12BCA: push eax
  loc_00E12BCB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12BD1: mov eax, var_30
  loc_00E12BD4: lea edx, var_54
  loc_00E12BD7: push edx
  loc_00E12BD8: push eax
  loc_00E12BD9: mov ecx, [eax]
  loc_00E12BDB: mov ebx, eax
  loc_00E12BDD: call [ecx+00000034h]
  loc_00E12BE0: test eax, eax
  loc_00E12BE2: fnclex
  loc_00E12BE4: jge 00E12BF5h
  loc_00E12BE6: push 00000034h
  loc_00E12BE8: push 006E09F8h
  loc_00E12BED: push ebx
  loc_00E12BEE: push eax
  loc_00E12BEF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12BF5: mov eax, var_D4
  loc_00E12BFB: lea ecx, var_54
  loc_00E12BFE: push ecx
  loc_00E12BFF: mov ebx, [eax]
  loc_00E12C01: call [00401034h] ; __vbaStrVarMove
  loc_00E12C07: mov edx, eax
  loc_00E12C09: lea ecx, var_18
  loc_00E12C0C: call [00401228h] ; __vbaStrMove
  loc_00E12C12: mov edx, ebx
  loc_00E12C14: mov ebx, var_D4
  loc_00E12C1A: push eax
  loc_00E12C1B: push ebx
  loc_00E12C1C: call [edx+000000A4h]
  loc_00E12C22: test eax, eax
  loc_00E12C24: fnclex
  loc_00E12C26: jge 00E12C3Ah
  loc_00E12C28: push 000000A4h
  loc_00E12C2D: push 006DCB70h
  loc_00E12C32: push ebx
  loc_00E12C33: push eax
  loc_00E12C34: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12C3A: lea ecx, var_18
  loc_00E12C3D: call [00401258h] ; __vbaFreeStr
  loc_00E12C43: lea eax, var_34
  loc_00E12C46: lea ecx, var_30
  loc_00E12C49: push eax
  loc_00E12C4A: lea edx, var_2C
  loc_00E12C4D: push ecx
  loc_00E12C4E: lea eax, var_28
  loc_00E12C51: push edx
  loc_00E12C52: lea ecx, var_24
  loc_00E12C55: push eax
  loc_00E12C56: push ecx
  loc_00E12C57: push 00000005h
  loc_00E12C59: call [00401048h] ; __vbaFreeObjList
  loc_00E12C5F: lea edx, var_54
  loc_00E12C62: lea eax, var_44
  loc_00E12C65: push edx
  loc_00E12C66: push eax
  loc_00E12C67: push 00000002h
  loc_00E12C69: call [00401038h] ; __vbaFreeVarList
  loc_00E12C6F: mov ecx, [esi]
  loc_00E12C71: add esp, 00000024h
  loc_00E12C74: push esi
  loc_00E12C75: call [ecx+00000388h]
  loc_00E12C7B: lea edx, var_34
  loc_00E12C7E: push eax
  loc_00E12C7F: push edx
  loc_00E12C80: call edi
  loc_00E12C82: push 006DCBD8h
  loc_00E12C87: mov var_D4, eax
  loc_00E12C8D: mov eax, [esi]
  loc_00E12C8F: push 00000000h
  loc_00E12C91: push 00000018h
  loc_00E12C93: push esi
  loc_00E12C94: call [eax+00000490h]
  loc_00E12C9A: lea ecx, var_24
  loc_00E12C9D: push eax
  loc_00E12C9E: push ecx
  loc_00E12C9F: call edi
  loc_00E12CA1: lea edx, var_44
  loc_00E12CA4: push eax
  loc_00E12CA5: push edx
  loc_00E12CA6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E12CAC: add esp, 00000010h
  loc_00E12CAF: push eax
  loc_00E12CB0: call [00401120h] ; __vbaCastObjVar
  loc_00E12CB6: push eax
  loc_00E12CB7: lea eax, var_28
  loc_00E12CBA: push eax
  loc_00E12CBB: call edi
  loc_00E12CBD: mov ebx, eax
  loc_00E12CBF: lea edx, var_2C
  loc_00E12CC2: push edx
  loc_00E12CC3: push ebx
  loc_00E12CC4: mov ecx, [ebx]
  loc_00E12CC6: call [ecx+00000054h]
  loc_00E12CC9: test eax, eax
  loc_00E12CCB: fnclex
  loc_00E12CCD: jge 00E12CDEh
  loc_00E12CCF: push 00000054h
  loc_00E12CD1: push 006DCBE8h
  loc_00E12CD6: push ebx
  loc_00E12CD7: push eax
  loc_00E12CD8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12CDE: lea ebx, var_30
  loc_00E12CE1: mov eax, var_2C
  loc_00E12CE4: push ebx
  loc_00E12CE5: mov ecx, 00000002h
  loc_00E12CEA: sub esp, 00000010h
  loc_00E12CED: mov var_84, ecx
  loc_00E12CF3: mov ebx, esp
  loc_00E12CF5: mov var_7C, 00000005h
  loc_00E12CFC: mov edx, [eax]
  loc_00E12CFE: push eax
  loc_00E12CFF: mov [ebx], ecx
  loc_00E12D01: mov ecx, var_80
  loc_00E12D04: mov var_C4, eax
  loc_00E12D0A: mov [ebx+00000004h], ecx
  loc_00E12D0D: mov ecx, var_7C
  loc_00E12D10: mov [ebx+00000008h], ecx
  loc_00E12D13: mov ecx, var_78
  loc_00E12D16: mov [ebx+0000000Ch], ecx
  loc_00E12D19: call [edx+00000028h]
  loc_00E12D1C: test eax, eax
  loc_00E12D1E: fnclex
  loc_00E12D20: jge 00E12D37h
  loc_00E12D22: mov edx, var_C4
  loc_00E12D28: push 00000028h
  loc_00E12D2A: push 006E09E8h
  loc_00E12D2F: push edx
  loc_00E12D30: push eax
  loc_00E12D31: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12D37: mov eax, var_30
  loc_00E12D3A: lea edx, var_54
  loc_00E12D3D: push edx
  loc_00E12D3E: push eax
  loc_00E12D3F: mov ecx, [eax]
  loc_00E12D41: mov ebx, eax
  loc_00E12D43: call [ecx+00000034h]
  loc_00E12D46: test eax, eax
  loc_00E12D48: fnclex
  loc_00E12D4A: jge 00E12D5Bh
  loc_00E12D4C: push 00000034h
  loc_00E12D4E: push 006E09F8h
  loc_00E12D53: push ebx
  loc_00E12D54: push eax
  loc_00E12D55: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12D5B: mov eax, var_D4
  loc_00E12D61: lea ecx, var_54
  loc_00E12D64: push ecx
  loc_00E12D65: mov ebx, [eax]
  loc_00E12D67: call [00401034h] ; __vbaStrVarMove
  loc_00E12D6D: mov edx, eax
  loc_00E12D6F: lea ecx, var_18
  loc_00E12D72: call [00401228h] ; __vbaStrMove
  loc_00E12D78: mov edx, ebx
  loc_00E12D7A: mov ebx, var_D4
  loc_00E12D80: push eax
  loc_00E12D81: push ebx
  loc_00E12D82: call [edx+000000A4h]
  loc_00E12D88: test eax, eax
  loc_00E12D8A: fnclex
  loc_00E12D8C: jge 00E12DA0h
  loc_00E12D8E: push 000000A4h
  loc_00E12D93: push 006DCB70h
  loc_00E12D98: push ebx
  loc_00E12D99: push eax
  loc_00E12D9A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12DA0: lea ecx, var_18
  loc_00E12DA3: call [00401258h] ; __vbaFreeStr
  loc_00E12DA9: lea eax, var_34
  loc_00E12DAC: lea ecx, var_30
  loc_00E12DAF: push eax
  loc_00E12DB0: lea edx, var_2C
  loc_00E12DB3: push ecx
  loc_00E12DB4: lea eax, var_28
  loc_00E12DB7: push edx
  loc_00E12DB8: lea ecx, var_24
  loc_00E12DBB: push eax
  loc_00E12DBC: push ecx
  loc_00E12DBD: push 00000005h
  loc_00E12DBF: call [00401048h] ; __vbaFreeObjList
  loc_00E12DC5: lea edx, var_54
  loc_00E12DC8: lea eax, var_44
  loc_00E12DCB: push edx
  loc_00E12DCC: push eax
  loc_00E12DCD: push 00000002h
  loc_00E12DCF: call [00401038h] ; __vbaFreeVarList
  loc_00E12DD5: mov ecx, [esi]
  loc_00E12DD7: add esp, 00000024h
  loc_00E12DDA: push 006DCBD8h
  loc_00E12DDF: push 00000000h
  loc_00E12DE1: push 00000018h
  loc_00E12DE3: push esi
  loc_00E12DE4: call [ecx+00000490h]
  loc_00E12DEA: lea edx, var_24
  loc_00E12DED: push eax
  loc_00E12DEE: push edx
  loc_00E12DEF: call edi
  loc_00E12DF1: push eax
  loc_00E12DF2: lea eax, var_44
  loc_00E12DF5: push eax
  loc_00E12DF6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E12DFC: add esp, 00000010h
  loc_00E12DFF: push eax
  loc_00E12E00: call [00401120h] ; __vbaCastObjVar
  loc_00E12E06: lea ecx, var_28
  loc_00E12E09: push eax
  loc_00E12E0A: push ecx
  loc_00E12E0B: call edi
  loc_00E12E0D: mov ebx, eax
  loc_00E12E0F: lea eax, var_2C
  loc_00E12E12: push eax
  loc_00E12E13: push ebx
  loc_00E12E14: mov edx, [ebx]
  loc_00E12E16: call [edx+00000054h]
  loc_00E12E19: test eax, eax
  loc_00E12E1B: fnclex
  loc_00E12E1D: jge 00E12E2Eh
  loc_00E12E1F: push 00000054h
  loc_00E12E21: push 006DCBE8h
  loc_00E12E26: push ebx
  loc_00E12E27: push eax
  loc_00E12E28: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12E2E: lea ebx, var_30
  loc_00E12E31: mov eax, var_2C
  loc_00E12E34: push ebx
  loc_00E12E35: mov ecx, 00000002h
  loc_00E12E3A: sub esp, 00000010h
  loc_00E12E3D: mov var_84, ecx
  loc_00E12E43: mov ebx, esp
  loc_00E12E45: mov var_7C, 00000006h
  loc_00E12E4C: mov edx, [eax]
  loc_00E12E4E: push eax
  loc_00E12E4F: mov [ebx], ecx
  loc_00E12E51: mov ecx, var_80
  loc_00E12E54: mov var_C4, eax
  loc_00E12E5A: mov [ebx+00000004h], ecx
  loc_00E12E5D: mov ecx, var_7C
  loc_00E12E60: mov [ebx+00000008h], ecx
  loc_00E12E63: mov ecx, var_78
  loc_00E12E66: mov [ebx+0000000Ch], ecx
  loc_00E12E69: call [edx+00000028h]
  loc_00E12E6C: test eax, eax
  loc_00E12E6E: fnclex
  loc_00E12E70: jge 00E12E87h
  loc_00E12E72: mov edx, var_C4
  loc_00E12E78: push 00000028h
  loc_00E12E7A: push 006E09E8h
  loc_00E12E7F: push edx
  loc_00E12E80: push eax
  loc_00E12E81: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12E87: sub esp, 00000010h
  loc_00E12E8A: mov eax, var_30
  loc_00E12E8D: mov edx, esp
  loc_00E12E8F: mov ecx, 00000009h
  loc_00E12E94: mov var_54, ecx
  loc_00E12E97: mov var_4C, eax
  loc_00E12E9A: mov [edx], ecx
  loc_00E12E9C: mov ecx, var_50
  loc_00E12E9F: push 00000014h
  loc_00E12EA1: push esi
  loc_00E12EA2: mov [edx+00000004h], ecx
  loc_00E12EA5: mov ecx, [esi]
  loc_00E12EA7: mov var_30, 00000000h
  loc_00E12EAE: mov [edx+00000008h], eax
  loc_00E12EB1: mov eax, var_48
  loc_00E12EB4: mov [edx+0000000Ch], eax
  loc_00E12EB7: call [ecx+0000048Ch]
  loc_00E12EBD: lea edx, var_34
  loc_00E12EC0: push eax
  loc_00E12EC1: push edx
  loc_00E12EC2: call edi
  loc_00E12EC4: push eax
  loc_00E12EC5: call [00401238h] ; __vbaLateIdSt
  loc_00E12ECB: lea eax, var_34
  loc_00E12ECE: lea ecx, var_2C
  loc_00E12ED1: push eax
  loc_00E12ED2: lea edx, var_28
  loc_00E12ED5: push ecx
  loc_00E12ED6: lea eax, var_24
  loc_00E12ED9: push edx
  loc_00E12EDA: push eax
  loc_00E12EDB: push 00000004h
  loc_00E12EDD: call [00401048h] ; __vbaFreeObjList
  loc_00E12EE3: lea ecx, var_54
  loc_00E12EE6: lea edx, var_44
  loc_00E12EE9: push ecx
  loc_00E12EEA: push edx
  loc_00E12EEB: push 00000002h
  loc_00E12EED: call [00401038h] ; __vbaFreeVarList
  loc_00E12EF3: mov eax, [esi]
  loc_00E12EF5: add esp, 00000020h
  loc_00E12EF8: push esi
  loc_00E12EF9: call [eax+00000378h]
  loc_00E12EFF: lea ecx, var_34
  loc_00E12F02: push eax
  loc_00E12F03: push ecx
  loc_00E12F04: call edi
  loc_00E12F06: mov edx, [esi]
  loc_00E12F08: push 006DCBD8h
  loc_00E12F0D: push 00000000h
  loc_00E12F0F: push 00000018h
  loc_00E12F11: push esi
  loc_00E12F12: mov var_D4, eax
  loc_00E12F18: call [edx+00000490h]
  loc_00E12F1E: push eax
  loc_00E12F1F: lea eax, var_24
  loc_00E12F22: push eax
  loc_00E12F23: call edi
  loc_00E12F25: lea ecx, var_44
  loc_00E12F28: push eax
  loc_00E12F29: push ecx
  loc_00E12F2A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E12F30: add esp, 00000010h
  loc_00E12F33: push eax
  loc_00E12F34: call [00401120h] ; __vbaCastObjVar
  loc_00E12F3A: lea edx, var_28
  loc_00E12F3D: push eax
  loc_00E12F3E: push edx
  loc_00E12F3F: call edi
  loc_00E12F41: mov ebx, eax
  loc_00E12F43: lea ecx, var_2C
  loc_00E12F46: push ecx
  loc_00E12F47: push ebx
  loc_00E12F48: mov eax, [ebx]
  loc_00E12F4A: call [eax+00000054h]
  loc_00E12F4D: test eax, eax
  loc_00E12F4F: fnclex
  loc_00E12F51: jge 00E12F62h
  loc_00E12F53: push 00000054h
  loc_00E12F55: push 006DCBE8h
  loc_00E12F5A: push ebx
  loc_00E12F5B: push eax
  loc_00E12F5C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12F62: lea ebx, var_30
  loc_00E12F65: mov eax, var_2C
  loc_00E12F68: push ebx
  loc_00E12F69: mov ecx, 00000002h
  loc_00E12F6E: sub esp, 00000010h
  loc_00E12F71: mov var_84, ecx
  loc_00E12F77: mov ebx, esp
  loc_00E12F79: mov var_7C, 00000007h
  loc_00E12F80: mov edx, [eax]
  loc_00E12F82: push eax
  loc_00E12F83: mov [ebx], ecx
  loc_00E12F85: mov ecx, var_80
  loc_00E12F88: mov var_C4, eax
  loc_00E12F8E: mov [ebx+00000004h], ecx
  loc_00E12F91: mov ecx, var_7C
  loc_00E12F94: mov [ebx+00000008h], ecx
  loc_00E12F97: mov ecx, var_78
  loc_00E12F9A: mov [ebx+0000000Ch], ecx
  loc_00E12F9D: call [edx+00000028h]
  loc_00E12FA0: test eax, eax
  loc_00E12FA2: fnclex
  loc_00E12FA4: jge 00E12FBBh
  loc_00E12FA6: mov edx, var_C4
  loc_00E12FAC: push 00000028h
  loc_00E12FAE: push 006E09E8h
  loc_00E12FB3: push edx
  loc_00E12FB4: push eax
  loc_00E12FB5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12FBB: mov eax, var_30
  loc_00E12FBE: lea edx, var_54
  loc_00E12FC1: push edx
  loc_00E12FC2: push eax
  loc_00E12FC3: mov ecx, [eax]
  loc_00E12FC5: mov ebx, eax
  loc_00E12FC7: call [ecx+00000034h]
  loc_00E12FCA: test eax, eax
  loc_00E12FCC: fnclex
  loc_00E12FCE: jge 00E12FDFh
  loc_00E12FD0: push 00000034h
  loc_00E12FD2: push 006E09F8h
  loc_00E12FD7: push ebx
  loc_00E12FD8: push eax
  loc_00E12FD9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E12FDF: mov eax, var_D4
  loc_00E12FE5: lea ecx, var_54
  loc_00E12FE8: push ecx
  loc_00E12FE9: mov ebx, [eax]
  loc_00E12FEB: call [00401034h] ; __vbaStrVarMove
  loc_00E12FF1: mov edx, eax
  loc_00E12FF3: lea ecx, var_18
  loc_00E12FF6: call [00401228h] ; __vbaStrMove
  loc_00E12FFC: mov edx, ebx
  loc_00E12FFE: mov ebx, var_D4
  loc_00E13004: push eax
  loc_00E13005: push ebx
  loc_00E13006: call [edx+000000ACh]
  loc_00E1300C: test eax, eax
  loc_00E1300E: fnclex
  loc_00E13010: jge 00E13024h
  loc_00E13012: push 000000ACh
  loc_00E13017: push 006E0400h
  loc_00E1301C: push ebx
  loc_00E1301D: push eax
  loc_00E1301E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13024: lea ecx, var_18
  loc_00E13027: call [00401258h] ; __vbaFreeStr
  loc_00E1302D: lea eax, var_34
  loc_00E13030: lea ecx, var_30
  loc_00E13033: push eax
  loc_00E13034: lea edx, var_2C
  loc_00E13037: push ecx
  loc_00E13038: lea eax, var_28
  loc_00E1303B: push edx
  loc_00E1303C: lea ecx, var_24
  loc_00E1303F: push eax
  loc_00E13040: push ecx
  loc_00E13041: push 00000005h
  loc_00E13043: call [00401048h] ; __vbaFreeObjList
  loc_00E13049: lea edx, var_54
  loc_00E1304C: lea eax, var_44
  loc_00E1304F: push edx
  loc_00E13050: push eax
  loc_00E13051: push 00000002h
  loc_00E13053: call [00401038h] ; __vbaFreeVarList
  loc_00E13059: mov ecx, [esi]
  loc_00E1305B: add esp, 00000024h
  loc_00E1305E: push esi
  loc_00E1305F: call [ecx+00000384h]
  loc_00E13065: lea edx, var_34
  loc_00E13068: push eax
  loc_00E13069: push edx
  loc_00E1306A: call edi
  loc_00E1306C: push 006DCBD8h
  loc_00E13071: mov var_D4, eax
  loc_00E13077: mov eax, [esi]
  loc_00E13079: push 00000000h
  loc_00E1307B: push 00000018h
  loc_00E1307D: push esi
  loc_00E1307E: call [eax+00000490h]
  loc_00E13084: lea ecx, var_24
  loc_00E13087: push eax
  loc_00E13088: push ecx
  loc_00E13089: call edi
  loc_00E1308B: lea edx, var_44
  loc_00E1308E: push eax
  loc_00E1308F: push edx
  loc_00E13090: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E13096: add esp, 00000010h
  loc_00E13099: push eax
  loc_00E1309A: call [00401120h] ; __vbaCastObjVar
  loc_00E130A0: push eax
  loc_00E130A1: lea eax, var_28
  loc_00E130A4: push eax
  loc_00E130A5: call edi
  loc_00E130A7: mov ebx, eax
  loc_00E130A9: lea edx, var_2C
  loc_00E130AC: push edx
  loc_00E130AD: push ebx
  loc_00E130AE: mov ecx, [ebx]
  loc_00E130B0: call [ecx+00000054h]
  loc_00E130B3: test eax, eax
  loc_00E130B5: fnclex
  loc_00E130B7: jge 00E130C8h
  loc_00E130B9: push 00000054h
  loc_00E130BB: push 006DCBE8h
  loc_00E130C0: push ebx
  loc_00E130C1: push eax
  loc_00E130C2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E130C8: lea ebx, var_30
  loc_00E130CB: mov eax, var_2C
  loc_00E130CE: push ebx
  loc_00E130CF: mov ecx, 00000002h
  loc_00E130D4: sub esp, 00000010h
  loc_00E130D7: mov var_84, ecx
  loc_00E130DD: mov ebx, esp
  loc_00E130DF: mov var_7C, 00000008h
  loc_00E130E6: mov edx, [eax]
  loc_00E130E8: push eax
  loc_00E130E9: mov [ebx], ecx
  loc_00E130EB: mov ecx, var_80
  loc_00E130EE: mov var_C4, eax
  loc_00E130F4: mov [ebx+00000004h], ecx
  loc_00E130F7: mov ecx, var_7C
  loc_00E130FA: mov [ebx+00000008h], ecx
  loc_00E130FD: mov ecx, var_78
  loc_00E13100: mov [ebx+0000000Ch], ecx
  loc_00E13103: call [edx+00000028h]
  loc_00E13106: test eax, eax
  loc_00E13108: fnclex
  loc_00E1310A: jge 00E13121h
  loc_00E1310C: mov edx, var_C4
  loc_00E13112: push 00000028h
  loc_00E13114: push 006E09E8h
  loc_00E13119: push edx
  loc_00E1311A: push eax
  loc_00E1311B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13121: mov eax, var_30
  loc_00E13124: lea edx, var_54
  loc_00E13127: push edx
  loc_00E13128: push eax
  loc_00E13129: mov ecx, [eax]
  loc_00E1312B: mov ebx, eax
  loc_00E1312D: call [ecx+00000034h]
  loc_00E13130: test eax, eax
  loc_00E13132: fnclex
  loc_00E13134: jge 00E13145h
  loc_00E13136: push 00000034h
  loc_00E13138: push 006E09F8h
  loc_00E1313D: push ebx
  loc_00E1313E: push eax
  loc_00E1313F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13145: mov eax, var_D4
  loc_00E1314B: lea ecx, var_54
  loc_00E1314E: push ecx
  loc_00E1314F: mov ebx, [eax]
  loc_00E13151: call [00401034h] ; __vbaStrVarMove
  loc_00E13157: mov edx, eax
  loc_00E13159: lea ecx, var_18
  loc_00E1315C: call [00401228h] ; __vbaStrMove
  loc_00E13162: mov edx, ebx
  loc_00E13164: mov ebx, var_D4
  loc_00E1316A: push eax
  loc_00E1316B: push ebx
  loc_00E1316C: call [edx+000000A4h]
  loc_00E13172: test eax, eax
  loc_00E13174: fnclex
  loc_00E13176: jge 00E1318Ah
  loc_00E13178: push 000000A4h
  loc_00E1317D: push 006DCB70h
  loc_00E13182: push ebx
  loc_00E13183: push eax
  loc_00E13184: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1318A: lea ecx, var_18
  loc_00E1318D: call [00401258h] ; __vbaFreeStr
  loc_00E13193: lea eax, var_34
  loc_00E13196: lea ecx, var_30
  loc_00E13199: push eax
  loc_00E1319A: lea edx, var_2C
  loc_00E1319D: push ecx
  loc_00E1319E: lea eax, var_28
  loc_00E131A1: push edx
  loc_00E131A2: lea ecx, var_24
  loc_00E131A5: push eax
  loc_00E131A6: push ecx
  loc_00E131A7: push 00000005h
  loc_00E131A9: call [00401048h] ; __vbaFreeObjList
  loc_00E131AF: lea edx, var_54
  loc_00E131B2: lea eax, var_44
  loc_00E131B5: push edx
  loc_00E131B6: push eax
  loc_00E131B7: push 00000002h
  loc_00E131B9: call [00401038h] ; __vbaFreeVarList
  loc_00E131BF: mov ecx, [esi]
  loc_00E131C1: add esp, 00000024h
  loc_00E131C4: push esi
  loc_00E131C5: call [ecx+00000374h]
  loc_00E131CB: lea edx, var_34
  loc_00E131CE: push eax
  loc_00E131CF: push edx
  loc_00E131D0: call edi
  loc_00E131D2: push 006DCBD8h
  loc_00E131D7: mov var_D4, eax
  loc_00E131DD: mov eax, [esi]
  loc_00E131DF: push 00000000h
  loc_00E131E1: push 00000018h
  loc_00E131E3: push esi
  loc_00E131E4: call [eax+00000490h]
  loc_00E131EA: lea ecx, var_24
  loc_00E131ED: push eax
  loc_00E131EE: push ecx
  loc_00E131EF: call edi
  loc_00E131F1: lea edx, var_44
  loc_00E131F4: push eax
  loc_00E131F5: push edx
  loc_00E131F6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E131FC: add esp, 00000010h
  loc_00E131FF: push eax
  loc_00E13200: call [00401120h] ; __vbaCastObjVar
  loc_00E13206: push eax
  loc_00E13207: lea eax, var_28
  loc_00E1320A: push eax
  loc_00E1320B: call edi
  loc_00E1320D: mov ebx, eax
  loc_00E1320F: lea edx, var_2C
  loc_00E13212: push edx
  loc_00E13213: push ebx
  loc_00E13214: mov ecx, [ebx]
  loc_00E13216: call [ecx+00000054h]
  loc_00E13219: test eax, eax
  loc_00E1321B: fnclex
  loc_00E1321D: jge 00E1322Eh
  loc_00E1321F: push 00000054h
  loc_00E13221: push 006DCBE8h
  loc_00E13226: push ebx
  loc_00E13227: push eax
  loc_00E13228: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1322E: lea ebx, var_30
  loc_00E13231: mov eax, var_2C
  loc_00E13234: push ebx
  loc_00E13235: mov ecx, 00000002h
  loc_00E1323A: sub esp, 00000010h
  loc_00E1323D: mov var_84, ecx
  loc_00E13243: mov ebx, esp
  loc_00E13245: mov var_7C, 00000009h
  loc_00E1324C: mov edx, [eax]
  loc_00E1324E: push eax
  loc_00E1324F: mov [ebx], ecx
  loc_00E13251: mov ecx, var_80
  loc_00E13254: mov var_C4, eax
  loc_00E1325A: mov [ebx+00000004h], ecx
  loc_00E1325D: mov ecx, var_7C
  loc_00E13260: mov [ebx+00000008h], ecx
  loc_00E13263: mov ecx, var_78
  loc_00E13266: mov [ebx+0000000Ch], ecx
  loc_00E13269: call [edx+00000028h]
  loc_00E1326C: test eax, eax
  loc_00E1326E: fnclex
  loc_00E13270: jge 00E13287h
  loc_00E13272: mov edx, var_C4
  loc_00E13278: push 00000028h
  loc_00E1327A: push 006E09E8h
  loc_00E1327F: push edx
  loc_00E13280: push eax
  loc_00E13281: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13287: mov eax, var_30
  loc_00E1328A: lea edx, var_54
  loc_00E1328D: push edx
  loc_00E1328E: push eax
  loc_00E1328F: mov ecx, [eax]
  loc_00E13291: mov ebx, eax
  loc_00E13293: call [ecx+00000034h]
  loc_00E13296: test eax, eax
  loc_00E13298: fnclex
  loc_00E1329A: jge 00E132ABh
  loc_00E1329C: push 00000034h
  loc_00E1329E: push 006E09F8h
  loc_00E132A3: push ebx
  loc_00E132A4: push eax
  loc_00E132A5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E132AB: mov eax, var_D4
  loc_00E132B1: lea ecx, var_54
  loc_00E132B4: push ecx
  loc_00E132B5: mov ebx, [eax]
  loc_00E132B7: call [00401034h] ; __vbaStrVarMove
  loc_00E132BD: mov edx, eax
  loc_00E132BF: lea ecx, var_18
  loc_00E132C2: call [00401228h] ; __vbaStrMove
  loc_00E132C8: mov edx, ebx
  loc_00E132CA: mov ebx, var_D4
  loc_00E132D0: push eax
  loc_00E132D1: push ebx
  loc_00E132D2: call [edx+000000A4h]
  loc_00E132D8: test eax, eax
  loc_00E132DA: fnclex
  loc_00E132DC: jge 00E132F0h
  loc_00E132DE: push 000000A4h
  loc_00E132E3: push 006DCB70h
  loc_00E132E8: push ebx
  loc_00E132E9: push eax
  loc_00E132EA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E132F0: lea ecx, var_18
  loc_00E132F3: call [00401258h] ; __vbaFreeStr
  loc_00E132F9: lea eax, var_34
  loc_00E132FC: lea ecx, var_30
  loc_00E132FF: push eax
  loc_00E13300: lea edx, var_2C
  loc_00E13303: push ecx
  loc_00E13304: lea eax, var_28
  loc_00E13307: push edx
  loc_00E13308: lea ecx, var_24
  loc_00E1330B: push eax
  loc_00E1330C: push ecx
  loc_00E1330D: push 00000005h
  loc_00E1330F: call [00401048h] ; __vbaFreeObjList
  loc_00E13315: lea edx, var_54
  loc_00E13318: lea eax, var_44
  loc_00E1331B: push edx
  loc_00E1331C: push eax
  loc_00E1331D: push 00000002h
  loc_00E1331F: call [00401038h] ; __vbaFreeVarList
  loc_00E13325: mov ecx, [esi]
  loc_00E13327: add esp, 00000024h
  loc_00E1332A: push esi
  loc_00E1332B: call [ecx+00000370h]
  loc_00E13331: lea edx, var_34
  loc_00E13334: push eax
  loc_00E13335: push edx
  loc_00E13336: call edi
  loc_00E13338: push 006DCBD8h
  loc_00E1333D: mov var_D4, eax
  loc_00E13343: mov eax, [esi]
  loc_00E13345: push 00000000h
  loc_00E13347: push 00000018h
  loc_00E13349: push esi
  loc_00E1334A: call [eax+00000490h]
  loc_00E13350: lea ecx, var_24
  loc_00E13353: push eax
  loc_00E13354: push ecx
  loc_00E13355: call edi
  loc_00E13357: lea edx, var_44
  loc_00E1335A: push eax
  loc_00E1335B: push edx
  loc_00E1335C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E13362: add esp, 00000010h
  loc_00E13365: push eax
  loc_00E13366: call [00401120h] ; __vbaCastObjVar
  loc_00E1336C: push eax
  loc_00E1336D: lea eax, var_28
  loc_00E13370: push eax
  loc_00E13371: call edi
  loc_00E13373: mov ebx, eax
  loc_00E13375: lea edx, var_2C
  loc_00E13378: push edx
  loc_00E13379: push ebx
  loc_00E1337A: mov ecx, [ebx]
  loc_00E1337C: call [ecx+00000054h]
  loc_00E1337F: test eax, eax
  loc_00E13381: fnclex
  loc_00E13383: jge 00E13394h
  loc_00E13385: push 00000054h
  loc_00E13387: push 006DCBE8h
  loc_00E1338C: push ebx
  loc_00E1338D: push eax
  loc_00E1338E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13394: lea ebx, var_30
  loc_00E13397: mov eax, var_2C
  loc_00E1339A: push ebx
  loc_00E1339B: mov ecx, 00000002h
  loc_00E133A0: sub esp, 00000010h
  loc_00E133A3: mov var_84, ecx
  loc_00E133A9: mov ebx, esp
  loc_00E133AB: mov var_7C, 0000000Ah
  loc_00E133B2: mov edx, [eax]
  loc_00E133B4: push eax
  loc_00E133B5: mov [ebx], ecx
  loc_00E133B7: mov ecx, var_80
  loc_00E133BA: mov var_C4, eax
  loc_00E133C0: mov [ebx+00000004h], ecx
  loc_00E133C3: mov ecx, var_7C
  loc_00E133C6: mov [ebx+00000008h], ecx
  loc_00E133C9: mov ecx, var_78
  loc_00E133CC: mov [ebx+0000000Ch], ecx
  loc_00E133CF: call [edx+00000028h]
  loc_00E133D2: test eax, eax
  loc_00E133D4: fnclex
  loc_00E133D6: jge 00E133EDh
  loc_00E133D8: mov edx, var_C4
  loc_00E133DE: push 00000028h
  loc_00E133E0: push 006E09E8h
  loc_00E133E5: push edx
  loc_00E133E6: push eax
  loc_00E133E7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E133ED: mov eax, var_30
  loc_00E133F0: lea edx, var_54
  loc_00E133F3: push edx
  loc_00E133F4: push eax
  loc_00E133F5: mov ecx, [eax]
  loc_00E133F7: mov ebx, eax
  loc_00E133F9: call [ecx+00000034h]
  loc_00E133FC: test eax, eax
  loc_00E133FE: fnclex
  loc_00E13400: jge 00E13411h
  loc_00E13402: push 00000034h
  loc_00E13404: push 006E09F8h
  loc_00E13409: push ebx
  loc_00E1340A: push eax
  loc_00E1340B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13411: mov eax, var_D4
  loc_00E13417: lea ecx, var_54
  loc_00E1341A: push ecx
  loc_00E1341B: mov ebx, [eax]
  loc_00E1341D: call [00401034h] ; __vbaStrVarMove
  loc_00E13423: mov edx, eax
  loc_00E13425: lea ecx, var_18
  loc_00E13428: call [00401228h] ; __vbaStrMove
  loc_00E1342E: mov edx, ebx
  loc_00E13430: mov ebx, var_D4
  loc_00E13436: push eax
  loc_00E13437: push ebx
  loc_00E13438: call [edx+000000A4h]
  loc_00E1343E: test eax, eax
  loc_00E13440: fnclex
  loc_00E13442: jge 00E13456h
  loc_00E13444: push 000000A4h
  loc_00E13449: push 006DCB70h
  loc_00E1344E: push ebx
  loc_00E1344F: push eax
  loc_00E13450: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13456: lea ecx, var_18
  loc_00E13459: call [00401258h] ; __vbaFreeStr
  loc_00E1345F: lea eax, var_34
  loc_00E13462: lea ecx, var_30
  loc_00E13465: push eax
  loc_00E13466: lea edx, var_2C
  loc_00E13469: push ecx
  loc_00E1346A: lea eax, var_28
  loc_00E1346D: push edx
  loc_00E1346E: lea ecx, var_24
  loc_00E13471: push eax
  loc_00E13472: push ecx
  loc_00E13473: push 00000005h
  loc_00E13475: call [00401048h] ; __vbaFreeObjList
  loc_00E1347B: lea edx, var_54
  loc_00E1347E: lea eax, var_44
  loc_00E13481: push edx
  loc_00E13482: push eax
  loc_00E13483: push 00000002h
  loc_00E13485: call [00401038h] ; __vbaFreeVarList
  loc_00E1348B: mov ecx, [esi]
  loc_00E1348D: add esp, 00000024h
  loc_00E13490: push esi
  loc_00E13491: call [ecx+0000036Ch]
  loc_00E13497: lea edx, var_34
  loc_00E1349A: push eax
  loc_00E1349B: push edx
  loc_00E1349C: call edi
  loc_00E1349E: push 006DCBD8h
  loc_00E134A3: mov var_D4, eax
  loc_00E134A9: mov eax, [esi]
  loc_00E134AB: push 00000000h
  loc_00E134AD: push 00000018h
  loc_00E134AF: push esi
  loc_00E134B0: call [eax+00000490h]
  loc_00E134B6: lea ecx, var_24
  loc_00E134B9: push eax
  loc_00E134BA: push ecx
  loc_00E134BB: call edi
  loc_00E134BD: lea edx, var_44
  loc_00E134C0: push eax
  loc_00E134C1: push edx
  loc_00E134C2: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E134C8: add esp, 00000010h
  loc_00E134CB: push eax
  loc_00E134CC: call [00401120h] ; __vbaCastObjVar
  loc_00E134D2: push eax
  loc_00E134D3: lea eax, var_28
  loc_00E134D6: push eax
  loc_00E134D7: call edi
  loc_00E134D9: mov ebx, eax
  loc_00E134DB: lea edx, var_2C
  loc_00E134DE: push edx
  loc_00E134DF: push ebx
  loc_00E134E0: mov ecx, [ebx]
  loc_00E134E2: call [ecx+00000054h]
  loc_00E134E5: test eax, eax
  loc_00E134E7: fnclex
  loc_00E134E9: jge 00E134FAh
  loc_00E134EB: push 00000054h
  loc_00E134ED: push 006DCBE8h
  loc_00E134F2: push ebx
  loc_00E134F3: push eax
  loc_00E134F4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E134FA: lea ebx, var_30
  loc_00E134FD: mov eax, var_2C
  loc_00E13500: push ebx
  loc_00E13501: mov ecx, 00000002h
  loc_00E13506: sub esp, 00000010h
  loc_00E13509: mov var_84, ecx
  loc_00E1350F: mov ebx, esp
  loc_00E13511: mov var_7C, 0000000Bh
  loc_00E13518: mov edx, [eax]
  loc_00E1351A: push eax
  loc_00E1351B: mov [ebx], ecx
  loc_00E1351D: mov ecx, var_80
  loc_00E13520: mov var_C4, eax
  loc_00E13526: mov [ebx+00000004h], ecx
  loc_00E13529: mov ecx, var_7C
  loc_00E1352C: mov [ebx+00000008h], ecx
  loc_00E1352F: mov ecx, var_78
  loc_00E13532: mov [ebx+0000000Ch], ecx
  loc_00E13535: call [edx+00000028h]
  loc_00E13538: test eax, eax
  loc_00E1353A: fnclex
  loc_00E1353C: jge 00E13553h
  loc_00E1353E: mov edx, var_C4
  loc_00E13544: push 00000028h
  loc_00E13546: push 006E09E8h
  loc_00E1354B: push edx
  loc_00E1354C: push eax
  loc_00E1354D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13553: mov eax, var_30
  loc_00E13556: lea edx, var_54
  loc_00E13559: push edx
  loc_00E1355A: push eax
  loc_00E1355B: mov ecx, [eax]
  loc_00E1355D: mov ebx, eax
  loc_00E1355F: call [ecx+00000034h]
  loc_00E13562: test eax, eax
  loc_00E13564: fnclex
  loc_00E13566: jge 00E13577h
  loc_00E13568: push 00000034h
  loc_00E1356A: push 006E09F8h
  loc_00E1356F: push ebx
  loc_00E13570: push eax
  loc_00E13571: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13577: mov eax, var_D4
  loc_00E1357D: lea ecx, var_54
  loc_00E13580: push ecx
  loc_00E13581: mov ebx, [eax]
  loc_00E13583: call [00401034h] ; __vbaStrVarMove
  loc_00E13589: mov edx, eax
  loc_00E1358B: lea ecx, var_18
  loc_00E1358E: call [00401228h] ; __vbaStrMove
  loc_00E13594: mov edx, ebx
  loc_00E13596: mov ebx, var_D4
  loc_00E1359C: push eax
  loc_00E1359D: push ebx
  loc_00E1359E: call [edx+000000A4h]
  loc_00E135A4: test eax, eax
  loc_00E135A6: fnclex
  loc_00E135A8: jge 00E135BCh
  loc_00E135AA: push 000000A4h
  loc_00E135AF: push 006DCB70h
  loc_00E135B4: push ebx
  loc_00E135B5: push eax
  loc_00E135B6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E135BC: lea ecx, var_18
  loc_00E135BF: call [00401258h] ; __vbaFreeStr
  loc_00E135C5: lea eax, var_34
  loc_00E135C8: lea ecx, var_30
  loc_00E135CB: push eax
  loc_00E135CC: lea edx, var_2C
  loc_00E135CF: push ecx
  loc_00E135D0: lea eax, var_28
  loc_00E135D3: push edx
  loc_00E135D4: lea ecx, var_24
  loc_00E135D7: push eax
  loc_00E135D8: push ecx
  loc_00E135D9: push 00000005h
  loc_00E135DB: call [00401048h] ; __vbaFreeObjList
  loc_00E135E1: lea edx, var_54
  loc_00E135E4: lea eax, var_44
  loc_00E135E7: push edx
  loc_00E135E8: push eax
  loc_00E135E9: push 00000002h
  loc_00E135EB: call [00401038h] ; __vbaFreeVarList
  loc_00E135F1: mov ecx, [esi]
  loc_00E135F3: add esp, 00000024h
  loc_00E135F6: push esi
  loc_00E135F7: call [ecx+00000368h]
  loc_00E135FD: lea edx, var_34
  loc_00E13600: push eax
  loc_00E13601: push edx
  loc_00E13602: call edi
  loc_00E13604: push 006DCBD8h
  loc_00E13609: mov var_D4, eax
  loc_00E1360F: mov eax, [esi]
  loc_00E13611: push 00000000h
  loc_00E13613: push 00000018h
  loc_00E13615: push esi
  loc_00E13616: call [eax+00000490h]
  loc_00E1361C: lea ecx, var_24
  loc_00E1361F: push eax
  loc_00E13620: push ecx
  loc_00E13621: call edi
  loc_00E13623: lea edx, var_44
  loc_00E13626: push eax
  loc_00E13627: push edx
  loc_00E13628: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1362E: add esp, 00000010h
  loc_00E13631: push eax
  loc_00E13632: call [00401120h] ; __vbaCastObjVar
  loc_00E13638: push eax
  loc_00E13639: lea eax, var_28
  loc_00E1363C: push eax
  loc_00E1363D: call edi
  loc_00E1363F: mov ebx, eax
  loc_00E13641: lea edx, var_2C
  loc_00E13644: push edx
  loc_00E13645: push ebx
  loc_00E13646: mov ecx, [ebx]
  loc_00E13648: call [ecx+00000054h]
  loc_00E1364B: test eax, eax
  loc_00E1364D: fnclex
  loc_00E1364F: jge 00E13660h
  loc_00E13651: push 00000054h
  loc_00E13653: push 006DCBE8h
  loc_00E13658: push ebx
  loc_00E13659: push eax
  loc_00E1365A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13660: lea ebx, var_30
  loc_00E13663: mov eax, var_2C
  loc_00E13666: push ebx
  loc_00E13667: mov ecx, 00000002h
  loc_00E1366C: sub esp, 00000010h
  loc_00E1366F: mov var_84, ecx
  loc_00E13675: mov ebx, esp
  loc_00E13677: mov var_7C, 0000000Ch
  loc_00E1367E: mov edx, [eax]
  loc_00E13680: push eax
  loc_00E13681: mov [ebx], ecx
  loc_00E13683: mov ecx, var_80
  loc_00E13686: mov var_C4, eax
  loc_00E1368C: mov [ebx+00000004h], ecx
  loc_00E1368F: mov ecx, var_7C
  loc_00E13692: mov [ebx+00000008h], ecx
  loc_00E13695: mov ecx, var_78
  loc_00E13698: mov [ebx+0000000Ch], ecx
  loc_00E1369B: call [edx+00000028h]
  loc_00E1369E: test eax, eax
  loc_00E136A0: fnclex
  loc_00E136A2: jge 00E136B9h
  loc_00E136A4: mov edx, var_C4
  loc_00E136AA: push 00000028h
  loc_00E136AC: push 006E09E8h
  loc_00E136B1: push edx
  loc_00E136B2: push eax
  loc_00E136B3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E136B9: mov eax, var_30
  loc_00E136BC: lea edx, var_54
  loc_00E136BF: push edx
  loc_00E136C0: push eax
  loc_00E136C1: mov ecx, [eax]
  loc_00E136C3: mov ebx, eax
  loc_00E136C5: call [ecx+00000034h]
  loc_00E136C8: test eax, eax
  loc_00E136CA: fnclex
  loc_00E136CC: jge 00E136DDh
  loc_00E136CE: push 00000034h
  loc_00E136D0: push 006E09F8h
  loc_00E136D5: push ebx
  loc_00E136D6: push eax
  loc_00E136D7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E136DD: mov eax, var_D4
  loc_00E136E3: lea ecx, var_54
  loc_00E136E6: push ecx
  loc_00E136E7: mov ebx, [eax]
  loc_00E136E9: call [00401034h] ; __vbaStrVarMove
  loc_00E136EF: mov edx, eax
  loc_00E136F1: lea ecx, var_18
  loc_00E136F4: call [00401228h] ; __vbaStrMove
  loc_00E136FA: mov edx, ebx
  loc_00E136FC: mov ebx, var_D4
  loc_00E13702: push eax
  loc_00E13703: push ebx
  loc_00E13704: call [edx+000000A4h]
  loc_00E1370A: test eax, eax
  loc_00E1370C: fnclex
  loc_00E1370E: jge 00E13722h
  loc_00E13710: push 000000A4h
  loc_00E13715: push 006DCB70h
  loc_00E1371A: push ebx
  loc_00E1371B: push eax
  loc_00E1371C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13722: lea ecx, var_18
  loc_00E13725: call [00401258h] ; __vbaFreeStr
  loc_00E1372B: lea eax, var_34
  loc_00E1372E: lea ecx, var_30
  loc_00E13731: push eax
  loc_00E13732: lea edx, var_2C
  loc_00E13735: push ecx
  loc_00E13736: lea eax, var_28
  loc_00E13739: push edx
  loc_00E1373A: lea ecx, var_24
  loc_00E1373D: push eax
  loc_00E1373E: push ecx
  loc_00E1373F: push 00000005h
  loc_00E13741: call [00401048h] ; __vbaFreeObjList
  loc_00E13747: lea edx, var_54
  loc_00E1374A: lea eax, var_44
  loc_00E1374D: push edx
  loc_00E1374E: push eax
  loc_00E1374F: push 00000002h
  loc_00E13751: call [00401038h] ; __vbaFreeVarList
  loc_00E13757: mov ecx, [esi]
  loc_00E13759: add esp, 00000024h
  loc_00E1375C: push esi
  loc_00E1375D: call [ecx+00000364h]
  loc_00E13763: lea edx, var_34
  loc_00E13766: push eax
  loc_00E13767: push edx
  loc_00E13768: call edi
  loc_00E1376A: push 006DCBD8h
  loc_00E1376F: mov var_D4, eax
  loc_00E13775: mov eax, [esi]
  loc_00E13777: push 00000000h
  loc_00E13779: push 00000018h
  loc_00E1377B: push esi
  loc_00E1377C: call [eax+00000490h]
  loc_00E13782: lea ecx, var_24
  loc_00E13785: push eax
  loc_00E13786: push ecx
  loc_00E13787: call edi
  loc_00E13789: lea edx, var_44
  loc_00E1378C: push eax
  loc_00E1378D: push edx
  loc_00E1378E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E13794: add esp, 00000010h
  loc_00E13797: push eax
  loc_00E13798: call [00401120h] ; __vbaCastObjVar
  loc_00E1379E: push eax
  loc_00E1379F: lea eax, var_28
  loc_00E137A2: push eax
  loc_00E137A3: call edi
  loc_00E137A5: mov ebx, eax
  loc_00E137A7: lea edx, var_2C
  loc_00E137AA: push edx
  loc_00E137AB: push ebx
  loc_00E137AC: mov ecx, [ebx]
  loc_00E137AE: call [ecx+00000054h]
  loc_00E137B1: test eax, eax
  loc_00E137B3: fnclex
  loc_00E137B5: jge 00E137C6h
  loc_00E137B7: push 00000054h
  loc_00E137B9: push 006DCBE8h
  loc_00E137BE: push ebx
  loc_00E137BF: push eax
  loc_00E137C0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E137C6: lea ebx, var_30
  loc_00E137C9: mov eax, var_2C
  loc_00E137CC: push ebx
  loc_00E137CD: mov ecx, 00000002h
  loc_00E137D2: sub esp, 00000010h
  loc_00E137D5: mov var_84, ecx
  loc_00E137DB: mov ebx, esp
  loc_00E137DD: mov var_7C, 0000000Dh
  loc_00E137E4: mov edx, [eax]
  loc_00E137E6: push eax
  loc_00E137E7: mov [ebx], ecx
  loc_00E137E9: mov ecx, var_80
  loc_00E137EC: mov var_C4, eax
  loc_00E137F2: mov [ebx+00000004h], ecx
  loc_00E137F5: mov ecx, var_7C
  loc_00E137F8: mov [ebx+00000008h], ecx
  loc_00E137FB: mov ecx, var_78
  loc_00E137FE: mov [ebx+0000000Ch], ecx
  loc_00E13801: call [edx+00000028h]
  loc_00E13804: test eax, eax
  loc_00E13806: fnclex
  loc_00E13808: jge 00E1381Fh
  loc_00E1380A: mov edx, var_C4
  loc_00E13810: push 00000028h
  loc_00E13812: push 006E09E8h
  loc_00E13817: push edx
  loc_00E13818: push eax
  loc_00E13819: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1381F: mov eax, var_30
  loc_00E13822: lea edx, var_54
  loc_00E13825: push edx
  loc_00E13826: push eax
  loc_00E13827: mov ecx, [eax]
  loc_00E13829: mov ebx, eax
  loc_00E1382B: call [ecx+00000034h]
  loc_00E1382E: test eax, eax
  loc_00E13830: fnclex
  loc_00E13832: jge 00E13843h
  loc_00E13834: push 00000034h
  loc_00E13836: push 006E09F8h
  loc_00E1383B: push ebx
  loc_00E1383C: push eax
  loc_00E1383D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13843: mov eax, var_D4
  loc_00E13849: lea ecx, var_54
  loc_00E1384C: push ecx
  loc_00E1384D: mov ebx, [eax]
  loc_00E1384F: call [00401034h] ; __vbaStrVarMove
  loc_00E13855: mov edx, eax
  loc_00E13857: lea ecx, var_18
  loc_00E1385A: call [00401228h] ; __vbaStrMove
  loc_00E13860: mov edx, ebx
  loc_00E13862: mov ebx, var_D4
  loc_00E13868: push eax
  loc_00E13869: push ebx
  loc_00E1386A: call [edx+000000A4h]
  loc_00E13870: test eax, eax
  loc_00E13872: fnclex
  loc_00E13874: jge 00E13888h
  loc_00E13876: push 000000A4h
  loc_00E1387B: push 006DCB70h
  loc_00E13880: push ebx
  loc_00E13881: push eax
  loc_00E13882: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13888: lea ecx, var_18
  loc_00E1388B: call [00401258h] ; __vbaFreeStr
  loc_00E13891: lea eax, var_34
  loc_00E13894: lea ecx, var_30
  loc_00E13897: push eax
  loc_00E13898: lea edx, var_2C
  loc_00E1389B: push ecx
  loc_00E1389C: lea eax, var_28
  loc_00E1389F: push edx
  loc_00E138A0: lea ecx, var_24
  loc_00E138A3: push eax
  loc_00E138A4: push ecx
  loc_00E138A5: push 00000005h
  loc_00E138A7: call [00401048h] ; __vbaFreeObjList
  loc_00E138AD: lea edx, var_54
  loc_00E138B0: lea eax, var_44
  loc_00E138B3: push edx
  loc_00E138B4: push eax
  loc_00E138B5: push 00000002h
  loc_00E138B7: call [00401038h] ; __vbaFreeVarList
  loc_00E138BD: mov ecx, [esi]
  loc_00E138BF: add esp, 00000024h
  loc_00E138C2: push esi
  loc_00E138C3: call [ecx+00000360h]
  loc_00E138C9: lea edx, var_34
  loc_00E138CC: push eax
  loc_00E138CD: push edx
  loc_00E138CE: call edi
  loc_00E138D0: push 006DCBD8h
  loc_00E138D5: mov var_D4, eax
  loc_00E138DB: mov eax, [esi]
  loc_00E138DD: push 00000000h
  loc_00E138DF: push 00000018h
  loc_00E138E1: push esi
  loc_00E138E2: call [eax+00000490h]
  loc_00E138E8: lea ecx, var_24
  loc_00E138EB: push eax
  loc_00E138EC: push ecx
  loc_00E138ED: call edi
  loc_00E138EF: lea edx, var_44
  loc_00E138F2: push eax
  loc_00E138F3: push edx
  loc_00E138F4: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E138FA: add esp, 00000010h
  loc_00E138FD: push eax
  loc_00E138FE: call [00401120h] ; __vbaCastObjVar
  loc_00E13904: push eax
  loc_00E13905: lea eax, var_28
  loc_00E13908: push eax
  loc_00E13909: call edi
  loc_00E1390B: mov ebx, eax
  loc_00E1390D: lea edx, var_2C
  loc_00E13910: push edx
  loc_00E13911: push ebx
  loc_00E13912: mov ecx, [ebx]
  loc_00E13914: call [ecx+00000054h]
  loc_00E13917: test eax, eax
  loc_00E13919: fnclex
  loc_00E1391B: jge 00E1392Ch
  loc_00E1391D: push 00000054h
  loc_00E1391F: push 006DCBE8h
  loc_00E13924: push ebx
  loc_00E13925: push eax
  loc_00E13926: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1392C: lea ebx, var_30
  loc_00E1392F: mov eax, var_2C
  loc_00E13932: push ebx
  loc_00E13933: mov ecx, 00000002h
  loc_00E13938: sub esp, 00000010h
  loc_00E1393B: mov var_84, ecx
  loc_00E13941: mov ebx, esp
  loc_00E13943: mov var_7C, 0000000Eh
  loc_00E1394A: mov edx, [eax]
  loc_00E1394C: push eax
  loc_00E1394D: mov [ebx], ecx
  loc_00E1394F: mov ecx, var_80
  loc_00E13952: mov var_C4, eax
  loc_00E13958: mov [ebx+00000004h], ecx
  loc_00E1395B: mov ecx, var_7C
  loc_00E1395E: mov [ebx+00000008h], ecx
  loc_00E13961: mov ecx, var_78
  loc_00E13964: mov [ebx+0000000Ch], ecx
  loc_00E13967: call [edx+00000028h]
  loc_00E1396A: test eax, eax
  loc_00E1396C: fnclex
  loc_00E1396E: jge 00E13985h
  loc_00E13970: mov edx, var_C4
  loc_00E13976: push 00000028h
  loc_00E13978: push 006E09E8h
  loc_00E1397D: push edx
  loc_00E1397E: push eax
  loc_00E1397F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13985: mov eax, var_30
  loc_00E13988: lea edx, var_54
  loc_00E1398B: push edx
  loc_00E1398C: push eax
  loc_00E1398D: mov ecx, [eax]
  loc_00E1398F: mov ebx, eax
  loc_00E13991: call [ecx+00000034h]
  loc_00E13994: test eax, eax
  loc_00E13996: fnclex
  loc_00E13998: jge 00E139A9h
  loc_00E1399A: push 00000034h
  loc_00E1399C: push 006E09F8h
  loc_00E139A1: push ebx
  loc_00E139A2: push eax
  loc_00E139A3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E139A9: mov eax, var_D4
  loc_00E139AF: lea ecx, var_54
  loc_00E139B2: push ecx
  loc_00E139B3: mov ebx, [eax]
  loc_00E139B5: call [00401034h] ; __vbaStrVarMove
  loc_00E139BB: mov edx, eax
  loc_00E139BD: lea ecx, var_18
  loc_00E139C0: call [00401228h] ; __vbaStrMove
  loc_00E139C6: mov edx, ebx
  loc_00E139C8: mov ebx, var_D4
  loc_00E139CE: push eax
  loc_00E139CF: push ebx
  loc_00E139D0: call [edx+000000A4h]
  loc_00E139D6: test eax, eax
  loc_00E139D8: fnclex
  loc_00E139DA: jge 00E139EEh
  loc_00E139DC: push 000000A4h
  loc_00E139E1: push 006DCB70h
  loc_00E139E6: push ebx
  loc_00E139E7: push eax
  loc_00E139E8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E139EE: lea ecx, var_18
  loc_00E139F1: call [00401258h] ; __vbaFreeStr
  loc_00E139F7: lea eax, var_34
  loc_00E139FA: lea ecx, var_30
  loc_00E139FD: push eax
  loc_00E139FE: lea edx, var_2C
  loc_00E13A01: push ecx
  loc_00E13A02: lea eax, var_28
  loc_00E13A05: push edx
  loc_00E13A06: lea ecx, var_24
  loc_00E13A09: push eax
  loc_00E13A0A: push ecx
  loc_00E13A0B: push 00000005h
  loc_00E13A0D: call [00401048h] ; __vbaFreeObjList
  loc_00E13A13: lea edx, var_54
  loc_00E13A16: lea eax, var_44
  loc_00E13A19: push edx
  loc_00E13A1A: push eax
  loc_00E13A1B: push 00000002h
  loc_00E13A1D: call [00401038h] ; __vbaFreeVarList
  loc_00E13A23: mov ecx, [esi]
  loc_00E13A25: add esp, 00000024h
  loc_00E13A28: push esi
  loc_00E13A29: call [ecx+00000338h]
  loc_00E13A2F: lea edx, var_24
  loc_00E13A32: push eax
  loc_00E13A33: push edx
  loc_00E13A34: call edi
  loc_00E13A36: mov esi, eax
  loc_00E13A38: push FFFFFFFFh
  loc_00E13A3A: push esi
  loc_00E13A3B: mov eax, [esi]
  loc_00E13A3D: call [eax+0000009Ch]
  loc_00E13A43: test eax, eax
  loc_00E13A45: fnclex
  loc_00E13A47: jge 00E1599Dh
  loc_00E13A4D: jmp 00E1598Bh
  loc_00E13A52: mov ecx, [esi]
  loc_00E13A54: push esi
  loc_00E13A55: call [ecx+00000314h]
  loc_00E13A5B: lea edx, var_24
  loc_00E13A5E: push eax
  loc_00E13A5F: push edx
  loc_00E13A60: call edi
  loc_00E13A62: mov ecx, [eax]
  loc_00E13A64: lea edx, var_B8
  loc_00E13A6A: push edx
  loc_00E13A6B: push eax
  loc_00E13A6C: mov var_BC, eax
  loc_00E13A72: call [ecx+000000E0h]
  loc_00E13A78: test eax, eax
  loc_00E13A7A: fnclex
  loc_00E13A7C: jge 00E13A96h
  loc_00E13A7E: mov ecx, var_BC
  loc_00E13A84: push 000000E0h
  loc_00E13A89: push 006E03D4h
  loc_00E13A8E: push ecx
  loc_00E13A8F: push eax
  loc_00E13A90: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13A96: mov dx, var_B8
  loc_00E13A9D: lea ecx, var_24
  loc_00E13AA0: mov var_C4, dx
  loc_00E13AA7: call ebx
  loc_00E13AA9: cmp var_C4, 0000h
  loc_00E13AB1: jz 00E159A6h
  loc_00E13AB7: mov eax, [esi]
  loc_00E13AB9: push esi
  loc_00E13ABA: call [eax+00000324h]
  loc_00E13AC0: lea ecx, var_24
  loc_00E13AC3: push eax
  loc_00E13AC4: push ecx
  loc_00E13AC5: call edi
  loc_00E13AC7: mov edx, [eax]
  loc_00E13AC9: lea ecx, var_18
  loc_00E13ACC: push ecx
  loc_00E13ACD: push eax
  loc_00E13ACE: mov var_BC, eax
  loc_00E13AD4: call [edx+000000A0h]
  loc_00E13ADA: test eax, eax
  loc_00E13ADC: fnclex
  loc_00E13ADE: jge 00E13AF8h
  loc_00E13AE0: mov edx, var_BC
  loc_00E13AE6: push 000000A0h
  loc_00E13AEB: push 006DCB70h
  loc_00E13AF0: push edx
  loc_00E13AF1: push eax
  loc_00E13AF2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13AF8: mov eax, var_18
  loc_00E13AFB: push 006E0DB8h ; "select * from biodata where nama like '" & Chr(37)
  loc_00E13B00: push eax
  loc_00E13B01: call [0040106Ch] ; __vbaStrCat
  loc_00E13B07: mov edx, eax
  loc_00E13B09: lea ecx, var_1C
  loc_00E13B0C: call [00401228h] ; __vbaStrMove
  loc_00E13B12: push eax
  loc_00E13B13: push 006E0E10h ; Chr(37) & "' order by nama asc"
  loc_00E13B18: call [0040106Ch] ; __vbaStrCat
  loc_00E13B1E: mov edx, eax
  loc_00E13B20: lea ecx, var_20
  loc_00E13B23: call [00401228h] ; __vbaStrMove
  loc_00E13B29: mov edx, eax
  loc_00E13B2B: lea ecx, [esi+00000034h]
  loc_00E13B2E: call [004011B0h] ; __vbaStrCopy
  loc_00E13B34: lea ecx, var_20
  loc_00E13B37: lea edx, var_1C
  loc_00E13B3A: push ecx
  loc_00E13B3B: lea eax, var_18
  loc_00E13B3E: push edx
  loc_00E13B3F: push eax
  loc_00E13B40: push 00000003h
  loc_00E13B42: call [004011BCh] ; __vbaFreeStrList
  loc_00E13B48: add esp, 00000010h
  loc_00E13B4B: lea ecx, var_24
  loc_00E13B4E: call ebx
  loc_00E13B50: sub esp, 00000010h
  loc_00E13B53: mov ecx, 00004008h
  loc_00E13B58: mov edx, esp
  loc_00E13B5A: mov var_84, ecx
  loc_00E13B60: lea eax, [esi+00000034h]
  loc_00E13B63: push 0000000Eh
  loc_00E13B65: mov [edx], ecx
  loc_00E13B67: mov ecx, var_80
  loc_00E13B6A: mov var_7C, eax
  loc_00E13B6D: push esi
  loc_00E13B6E: mov [edx+00000004h], ecx
  loc_00E13B71: mov ecx, [esi]
  loc_00E13B73: mov [edx+00000008h], eax
  loc_00E13B76: mov eax, var_78
  loc_00E13B79: mov [edx+0000000Ch], eax
  loc_00E13B7C: call [ecx+00000490h]
  loc_00E13B82: lea edx, var_24
  loc_00E13B85: push eax
  loc_00E13B86: push edx
  loc_00E13B87: call edi
  loc_00E13B89: push eax
  loc_00E13B8A: call [00401238h] ; __vbaLateIdSt
  loc_00E13B90: lea ecx, var_24
  loc_00E13B93: call ebx
  loc_00E13B95: mov eax, [esi]
  loc_00E13B97: push 00000000h
  loc_00E13B99: push 00000019h
  loc_00E13B9B: push esi
  loc_00E13B9C: call [eax+00000490h]
  loc_00E13BA2: lea ecx, var_24
  loc_00E13BA5: push eax
  loc_00E13BA6: push ecx
  loc_00E13BA7: call edi
  loc_00E13BA9: push eax
  loc_00E13BAA: call [00401030h] ; __vbaLateIdCall
  loc_00E13BB0: add esp, 0000000Ch
  loc_00E13BB3: lea ecx, var_24
  loc_00E13BB6: call ebx
  loc_00E13BB8: mov edx, [esi]
  loc_00E13BBA: push 006DCBD8h
  loc_00E13BBF: push 00000000h
  loc_00E13BC1: push 00000018h
  loc_00E13BC3: push esi
  loc_00E13BC4: call [edx+00000490h]
  loc_00E13BCA: push eax
  loc_00E13BCB: lea eax, var_24
  loc_00E13BCE: push eax
  loc_00E13BCF: call edi
  loc_00E13BD1: lea ecx, var_44
  loc_00E13BD4: push eax
  loc_00E13BD5: push ecx
  loc_00E13BD6: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E13BDC: add esp, 00000010h
  loc_00E13BDF: push eax
  loc_00E13BE0: call [00401120h] ; __vbaCastObjVar
  loc_00E13BE6: lea edx, var_28
  loc_00E13BE9: push eax
  loc_00E13BEA: push edx
  loc_00E13BEB: call edi
  loc_00E13BED: mov ebx, eax
  loc_00E13BEF: lea ecx, var_B8
  loc_00E13BF5: push ecx
  loc_00E13BF6: push ebx
  loc_00E13BF7: mov eax, [ebx]
  loc_00E13BF9: call [eax+00000050h]
  loc_00E13BFC: test eax, eax
  loc_00E13BFE: fnclex
  loc_00E13C00: jge 00E13C11h
  loc_00E13C02: push 00000050h
  loc_00E13C04: push 006DCBE8h
  loc_00E13C09: push ebx
  loc_00E13C0A: push eax
  loc_00E13C0B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13C11: mov bx, var_B8
  loc_00E13C18: lea edx, var_28
  loc_00E13C1B: lea eax, var_24
  loc_00E13C1E: push edx
  loc_00E13C1F: push eax
  loc_00E13C20: push 00000002h
  loc_00E13C22: call [00401048h] ; __vbaFreeObjList
  loc_00E13C28: add esp, 0000000Ch
  loc_00E13C2B: lea ecx, var_44
  loc_00E13C2E: call [00401024h] ; __vbaFreeVar
  loc_00E13C34: test bx, bx
  loc_00E13C37: jnz 00E12207h
  loc_00E13C3D: mov eax, [esi]
  loc_00E13C3F: push esi
  loc_00E13C40: call [eax+0000039Ch]
  loc_00E13C46: lea ecx, var_34
  loc_00E13C49: push eax
  loc_00E13C4A: push ecx
  loc_00E13C4B: call edi
  loc_00E13C4D: mov edx, [esi]
  loc_00E13C4F: push 006DCBD8h
  loc_00E13C54: push 00000000h
  loc_00E13C56: push 00000018h
  loc_00E13C58: push esi
  loc_00E13C59: mov var_D4, eax
  loc_00E13C5F: call [edx+00000490h]
  loc_00E13C65: push eax
  loc_00E13C66: lea eax, var_24
  loc_00E13C69: push eax
  loc_00E13C6A: call edi
  loc_00E13C6C: lea ecx, var_44
  loc_00E13C6F: push eax
  loc_00E13C70: push ecx
  loc_00E13C71: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E13C77: add esp, 00000010h
  loc_00E13C7A: push eax
  loc_00E13C7B: call [00401120h] ; __vbaCastObjVar
  loc_00E13C81: lea edx, var_28
  loc_00E13C84: push eax
  loc_00E13C85: push edx
  loc_00E13C86: call edi
  loc_00E13C88: mov ebx, eax
  loc_00E13C8A: lea ecx, var_2C
  loc_00E13C8D: push ecx
  loc_00E13C8E: push ebx
  loc_00E13C8F: mov eax, [ebx]
  loc_00E13C91: call [eax+00000054h]
  loc_00E13C94: test eax, eax
  loc_00E13C96: fnclex
  loc_00E13C98: jge 00E13CA9h
  loc_00E13C9A: push 00000054h
  loc_00E13C9C: push 006DCBE8h
  loc_00E13CA1: push ebx
  loc_00E13CA2: push eax
  loc_00E13CA3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13CA9: lea ebx, var_30
  loc_00E13CAC: mov eax, var_2C
  loc_00E13CAF: push ebx
  loc_00E13CB0: mov ecx, 00000002h
  loc_00E13CB5: sub esp, 00000010h
  loc_00E13CB8: mov var_84, ecx
  loc_00E13CBE: mov ebx, esp
  loc_00E13CC0: mov var_7C, 00000000h
  loc_00E13CC7: mov edx, [eax]
  loc_00E13CC9: push eax
  loc_00E13CCA: mov [ebx], ecx
  loc_00E13CCC: mov ecx, var_80
  loc_00E13CCF: mov var_C4, eax
  loc_00E13CD5: mov [ebx+00000004h], ecx
  loc_00E13CD8: mov ecx, var_7C
  loc_00E13CDB: mov [ebx+00000008h], ecx
  loc_00E13CDE: mov ecx, var_78
  loc_00E13CE1: mov [ebx+0000000Ch], ecx
  loc_00E13CE4: call [edx+00000028h]
  loc_00E13CE7: test eax, eax
  loc_00E13CE9: fnclex
  loc_00E13CEB: jge 00E13D02h
  loc_00E13CED: mov edx, var_C4
  loc_00E13CF3: push 00000028h
  loc_00E13CF5: push 006E09E8h
  loc_00E13CFA: push edx
  loc_00E13CFB: push eax
  loc_00E13CFC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13D02: mov eax, var_30
  loc_00E13D05: lea edx, var_54
  loc_00E13D08: push edx
  loc_00E13D09: push eax
  loc_00E13D0A: mov ecx, [eax]
  loc_00E13D0C: mov ebx, eax
  loc_00E13D0E: call [ecx+00000034h]
  loc_00E13D11: test eax, eax
  loc_00E13D13: fnclex
  loc_00E13D15: jge 00E13D26h
  loc_00E13D17: push 00000034h
  loc_00E13D19: push 006E09F8h
  loc_00E13D1E: push ebx
  loc_00E13D1F: push eax
  loc_00E13D20: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13D26: mov eax, var_D4
  loc_00E13D2C: lea ecx, var_54
  loc_00E13D2F: push ecx
  loc_00E13D30: mov ebx, [eax]
  loc_00E13D32: call [00401034h] ; __vbaStrVarMove
  loc_00E13D38: mov edx, eax
  loc_00E13D3A: lea ecx, var_18
  loc_00E13D3D: call [00401228h] ; __vbaStrMove
  loc_00E13D43: mov edx, ebx
  loc_00E13D45: mov ebx, var_D4
  loc_00E13D4B: push eax
  loc_00E13D4C: push ebx
  loc_00E13D4D: call [edx+000000A4h]
  loc_00E13D53: test eax, eax
  loc_00E13D55: fnclex
  loc_00E13D57: jge 00E13D6Bh
  loc_00E13D59: push 000000A4h
  loc_00E13D5E: push 006DCB70h
  loc_00E13D63: push ebx
  loc_00E13D64: push eax
  loc_00E13D65: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13D6B: lea ecx, var_18
  loc_00E13D6E: call [00401258h] ; __vbaFreeStr
  loc_00E13D74: lea eax, var_34
  loc_00E13D77: lea ecx, var_30
  loc_00E13D7A: push eax
  loc_00E13D7B: lea edx, var_2C
  loc_00E13D7E: push ecx
  loc_00E13D7F: lea eax, var_28
  loc_00E13D82: push edx
  loc_00E13D83: lea ecx, var_24
  loc_00E13D86: push eax
  loc_00E13D87: push ecx
  loc_00E13D88: push 00000005h
  loc_00E13D8A: call [00401048h] ; __vbaFreeObjList
  loc_00E13D90: lea edx, var_54
  loc_00E13D93: lea eax, var_44
  loc_00E13D96: push edx
  loc_00E13D97: push eax
  loc_00E13D98: push 00000002h
  loc_00E13D9A: call [00401038h] ; __vbaFreeVarList
  loc_00E13DA0: mov ecx, [esi]
  loc_00E13DA2: add esp, 00000024h
  loc_00E13DA5: push esi
  loc_00E13DA6: call [ecx+00000394h]
  loc_00E13DAC: lea edx, var_34
  loc_00E13DAF: push eax
  loc_00E13DB0: push edx
  loc_00E13DB1: call edi
  loc_00E13DB3: push 006DCBD8h
  loc_00E13DB8: mov var_D4, eax
  loc_00E13DBE: mov eax, [esi]
  loc_00E13DC0: push 00000000h
  loc_00E13DC2: push 00000018h
  loc_00E13DC4: push esi
  loc_00E13DC5: call [eax+00000490h]
  loc_00E13DCB: lea ecx, var_24
  loc_00E13DCE: push eax
  loc_00E13DCF: push ecx
  loc_00E13DD0: call edi
  loc_00E13DD2: lea edx, var_44
  loc_00E13DD5: push eax
  loc_00E13DD6: push edx
  loc_00E13DD7: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E13DDD: add esp, 00000010h
  loc_00E13DE0: push eax
  loc_00E13DE1: call [00401120h] ; __vbaCastObjVar
  loc_00E13DE7: push eax
  loc_00E13DE8: lea eax, var_28
  loc_00E13DEB: push eax
  loc_00E13DEC: call edi
  loc_00E13DEE: mov ebx, eax
  loc_00E13DF0: lea edx, var_2C
  loc_00E13DF3: push edx
  loc_00E13DF4: push ebx
  loc_00E13DF5: mov ecx, [ebx]
  loc_00E13DF7: call [ecx+00000054h]
  loc_00E13DFA: test eax, eax
  loc_00E13DFC: fnclex
  loc_00E13DFE: jge 00E13E0Fh
  loc_00E13E00: push 00000054h
  loc_00E13E02: push 006DCBE8h
  loc_00E13E07: push ebx
  loc_00E13E08: push eax
  loc_00E13E09: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13E0F: lea ebx, var_30
  loc_00E13E12: mov eax, var_2C
  loc_00E13E15: push ebx
  loc_00E13E16: mov ecx, 00000002h
  loc_00E13E1B: sub esp, 00000010h
  loc_00E13E1E: mov var_84, ecx
  loc_00E13E24: mov ebx, esp
  loc_00E13E26: mov var_7C, 00000001h
  loc_00E13E2D: mov edx, [eax]
  loc_00E13E2F: push eax
  loc_00E13E30: mov [ebx], ecx
  loc_00E13E32: mov ecx, var_80
  loc_00E13E35: mov var_C4, eax
  loc_00E13E3B: mov [ebx+00000004h], ecx
  loc_00E13E3E: mov ecx, var_7C
  loc_00E13E41: mov [ebx+00000008h], ecx
  loc_00E13E44: mov ecx, var_78
  loc_00E13E47: mov [ebx+0000000Ch], ecx
  loc_00E13E4A: call [edx+00000028h]
  loc_00E13E4D: test eax, eax
  loc_00E13E4F: fnclex
  loc_00E13E51: jge 00E13E68h
  loc_00E13E53: mov edx, var_C4
  loc_00E13E59: push 00000028h
  loc_00E13E5B: push 006E09E8h
  loc_00E13E60: push edx
  loc_00E13E61: push eax
  loc_00E13E62: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13E68: mov eax, var_30
  loc_00E13E6B: lea edx, var_54
  loc_00E13E6E: push edx
  loc_00E13E6F: push eax
  loc_00E13E70: mov ecx, [eax]
  loc_00E13E72: mov ebx, eax
  loc_00E13E74: call [ecx+00000034h]
  loc_00E13E77: test eax, eax
  loc_00E13E79: fnclex
  loc_00E13E7B: jge 00E13E8Ch
  loc_00E13E7D: push 00000034h
  loc_00E13E7F: push 006E09F8h
  loc_00E13E84: push ebx
  loc_00E13E85: push eax
  loc_00E13E86: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13E8C: mov eax, var_D4
  loc_00E13E92: lea ecx, var_54
  loc_00E13E95: push ecx
  loc_00E13E96: mov ebx, [eax]
  loc_00E13E98: call [00401034h] ; __vbaStrVarMove
  loc_00E13E9E: mov edx, eax
  loc_00E13EA0: lea ecx, var_18
  loc_00E13EA3: call [00401228h] ; __vbaStrMove
  loc_00E13EA9: mov edx, ebx
  loc_00E13EAB: mov ebx, var_D4
  loc_00E13EB1: push eax
  loc_00E13EB2: push ebx
  loc_00E13EB3: call [edx+000000A4h]
  loc_00E13EB9: test eax, eax
  loc_00E13EBB: fnclex
  loc_00E13EBD: jge 00E13ED1h
  loc_00E13EBF: push 000000A4h
  loc_00E13EC4: push 006DCB70h
  loc_00E13EC9: push ebx
  loc_00E13ECA: push eax
  loc_00E13ECB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13ED1: lea ecx, var_18
  loc_00E13ED4: call [00401258h] ; __vbaFreeStr
  loc_00E13EDA: lea eax, var_34
  loc_00E13EDD: lea ecx, var_30
  loc_00E13EE0: push eax
  loc_00E13EE1: lea edx, var_2C
  loc_00E13EE4: push ecx
  loc_00E13EE5: lea eax, var_28
  loc_00E13EE8: push edx
  loc_00E13EE9: lea ecx, var_24
  loc_00E13EEC: push eax
  loc_00E13EED: push ecx
  loc_00E13EEE: push 00000005h
  loc_00E13EF0: call [00401048h] ; __vbaFreeObjList
  loc_00E13EF6: lea edx, var_54
  loc_00E13EF9: lea eax, var_44
  loc_00E13EFC: push edx
  loc_00E13EFD: push eax
  loc_00E13EFE: push 00000002h
  loc_00E13F00: call [00401038h] ; __vbaFreeVarList
  loc_00E13F06: mov ecx, [esi]
  loc_00E13F08: add esp, 00000024h
  loc_00E13F0B: push 006DCBD8h
  loc_00E13F10: push 00000000h
  loc_00E13F12: push 00000018h
  loc_00E13F14: push esi
  loc_00E13F15: call [ecx+00000490h]
  loc_00E13F1B: lea edx, var_24
  loc_00E13F1E: push eax
  loc_00E13F1F: push edx
  loc_00E13F20: call edi
  loc_00E13F22: push eax
  loc_00E13F23: lea eax, var_44
  loc_00E13F26: push eax
  loc_00E13F27: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E13F2D: add esp, 00000010h
  loc_00E13F30: push eax
  loc_00E13F31: call [00401120h] ; __vbaCastObjVar
  loc_00E13F37: lea ecx, var_28
  loc_00E13F3A: push eax
  loc_00E13F3B: push ecx
  loc_00E13F3C: call edi
  loc_00E13F3E: mov ebx, eax
  loc_00E13F40: lea eax, var_2C
  loc_00E13F43: push eax
  loc_00E13F44: push ebx
  loc_00E13F45: mov edx, [ebx]
  loc_00E13F47: call [edx+00000054h]
  loc_00E13F4A: test eax, eax
  loc_00E13F4C: fnclex
  loc_00E13F4E: jge 00E13F5Fh
  loc_00E13F50: push 00000054h
  loc_00E13F52: push 006DCBE8h
  loc_00E13F57: push ebx
  loc_00E13F58: push eax
  loc_00E13F59: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13F5F: lea ebx, var_30
  loc_00E13F62: mov eax, var_2C
  loc_00E13F65: push ebx
  loc_00E13F66: mov ecx, 00000002h
  loc_00E13F6B: sub esp, 00000010h
  loc_00E13F6E: mov var_7C, ecx
  loc_00E13F71: mov ebx, esp
  loc_00E13F73: mov var_84, ecx
  loc_00E13F79: mov edx, [eax]
  loc_00E13F7B: push eax
  loc_00E13F7C: mov [ebx], ecx
  loc_00E13F7E: mov ecx, var_80
  loc_00E13F81: mov var_C4, eax
  loc_00E13F87: mov [ebx+00000004h], ecx
  loc_00E13F8A: mov ecx, var_7C
  loc_00E13F8D: mov [ebx+00000008h], ecx
  loc_00E13F90: mov ecx, var_78
  loc_00E13F93: mov [ebx+0000000Ch], ecx
  loc_00E13F96: call [edx+00000028h]
  loc_00E13F99: test eax, eax
  loc_00E13F9B: fnclex
  loc_00E13F9D: jge 00E13FB4h
  loc_00E13F9F: mov edx, var_C4
  loc_00E13FA5: push 00000028h
  loc_00E13FA7: push 006E09E8h
  loc_00E13FAC: push edx
  loc_00E13FAD: push eax
  loc_00E13FAE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13FB4: mov eax, var_30
  loc_00E13FB7: lea edx, var_54
  loc_00E13FBA: push edx
  loc_00E13FBB: push eax
  loc_00E13FBC: mov ecx, [eax]
  loc_00E13FBE: mov ebx, eax
  loc_00E13FC0: call [ecx+00000034h]
  loc_00E13FC3: test eax, eax
  loc_00E13FC5: fnclex
  loc_00E13FC7: jge 00E13FD8h
  loc_00E13FC9: push 00000034h
  loc_00E13FCB: push 006E09F8h
  loc_00E13FD0: push ebx
  loc_00E13FD1: push eax
  loc_00E13FD2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E13FD8: lea eax, var_54
  loc_00E13FDB: lea ecx, var_94
  loc_00E13FE1: push eax
  loc_00E13FE2: push ecx
  loc_00E13FE3: mov var_8C, 006E0B24h ; "Laki - laki"
  loc_00E13FED: mov var_94, 00008008h
  loc_00E13FF7: call [00401108h] ; __vbaVarTstEq
  loc_00E13FFD: mov bx, ax
  loc_00E14000: lea edx, var_30
  loc_00E14003: lea eax, var_2C
  loc_00E14006: push edx
  loc_00E14007: lea ecx, var_28
  loc_00E1400A: push eax
  loc_00E1400B: lea edx, var_24
  loc_00E1400E: push ecx
  loc_00E1400F: push edx
  loc_00E14010: push 00000004h
  loc_00E14012: call [00401048h] ; __vbaFreeObjList
  loc_00E14018: lea eax, var_54
  loc_00E1401B: lea ecx, var_44
  loc_00E1401E: push eax
  loc_00E1401F: push ecx
  loc_00E14020: push 00000002h
  loc_00E14022: call [00401038h] ; __vbaFreeVarList
  loc_00E14028: add esp, 00000020h
  loc_00E1402B: test bx, bx
  loc_00E1402E: jz 00E140DCh
  loc_00E14034: mov edx, [esi]
  loc_00E14036: push esi
  loc_00E14037: call [edx+00000380h]
  loc_00E1403D: push eax
  loc_00E1403E: lea eax, var_24
  loc_00E14041: push eax
  loc_00E14042: call edi
  loc_00E14044: mov ebx, eax
  loc_00E14046: push FFFFFFFFh
  loc_00E14048: push ebx
  loc_00E14049: mov ecx, [ebx]
  loc_00E1404B: call [ecx+000000E4h]
  loc_00E14051: test eax, eax
  loc_00E14053: fnclex
  loc_00E14055: jge 00E14069h
  loc_00E14057: push 000000E4h
  loc_00E1405C: push 006E03D4h
  loc_00E14061: push ebx
  loc_00E14062: push eax
  loc_00E14063: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14069: lea ecx, var_24
  loc_00E1406C: call [00401254h] ; __vbaFreeObj
  loc_00E14072: mov edx, [esi]
  loc_00E14074: push esi
  loc_00E14075: call [edx+0000037Ch]
  loc_00E1407B: push eax
  loc_00E1407C: lea eax, var_24
  loc_00E1407F: push eax
  loc_00E14080: call edi
  loc_00E14082: mov ebx, eax
  loc_00E14084: push FFFFFFFFh
  loc_00E14086: push ebx
  loc_00E14087: mov ecx, [ebx]
  loc_00E14089: call [ecx+0000009Ch]
  loc_00E1408F: test eax, eax
  loc_00E14091: fnclex
  loc_00E14093: jge 00E140A7h
  loc_00E14095: push 0000009Ch
  loc_00E1409A: push 006E03D4h
  loc_00E1409F: push ebx
  loc_00E140A0: push eax
  loc_00E140A1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E140A7: lea ecx, var_24
  loc_00E140AA: call [00401254h] ; __vbaFreeObj
  loc_00E140B0: mov edx, [esi]
  loc_00E140B2: push esi
  loc_00E140B3: call [edx+00000380h]
  loc_00E140B9: push eax
  loc_00E140BA: lea eax, var_24
  loc_00E140BD: push eax
  loc_00E140BE: call edi
  loc_00E140C0: mov ebx, eax
  loc_00E140C2: push FFFFFFFFh
  loc_00E140C4: push ebx
  loc_00E140C5: mov ecx, [ebx]
  loc_00E140C7: call [ecx+0000009Ch]
  loc_00E140CD: test eax, eax
  loc_00E140CF: fnclex
  loc_00E140D1: jge 00E142B8h
  loc_00E140D7: jmp 00E142A6h
  loc_00E140DC: mov edx, [esi]
  loc_00E140DE: push 006DCBD8h
  loc_00E140E3: push 00000000h
  loc_00E140E5: push 00000018h
  loc_00E140E7: push esi
  loc_00E140E8: call [edx+00000490h]
  loc_00E140EE: push eax
  loc_00E140EF: lea eax, var_24
  loc_00E140F2: push eax
  loc_00E140F3: call edi
  loc_00E140F5: lea ecx, var_44
  loc_00E140F8: push eax
  loc_00E140F9: push ecx
  loc_00E140FA: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E14100: add esp, 00000010h
  loc_00E14103: push eax
  loc_00E14104: call [00401120h] ; __vbaCastObjVar
  loc_00E1410A: lea edx, var_28
  loc_00E1410D: push eax
  loc_00E1410E: push edx
  loc_00E1410F: call edi
  loc_00E14111: mov ebx, eax
  loc_00E14113: lea ecx, var_2C
  loc_00E14116: push ecx
  loc_00E14117: push ebx
  loc_00E14118: mov eax, [ebx]
  loc_00E1411A: call [eax+00000054h]
  loc_00E1411D: test eax, eax
  loc_00E1411F: fnclex
  loc_00E14121: jge 00E14132h
  loc_00E14123: push 00000054h
  loc_00E14125: push 006DCBE8h
  loc_00E1412A: push ebx
  loc_00E1412B: push eax
  loc_00E1412C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14132: lea ebx, var_30
  loc_00E14135: mov eax, var_2C
  loc_00E14138: push ebx
  loc_00E14139: mov ecx, 00000002h
  loc_00E1413E: sub esp, 00000010h
  loc_00E14141: mov var_7C, ecx
  loc_00E14144: mov ebx, esp
  loc_00E14146: mov var_84, ecx
  loc_00E1414C: mov edx, [eax]
  loc_00E1414E: push eax
  loc_00E1414F: mov [ebx], ecx
  loc_00E14151: mov ecx, var_80
  loc_00E14154: mov var_C4, eax
  loc_00E1415A: mov [ebx+00000004h], ecx
  loc_00E1415D: mov ecx, var_7C
  loc_00E14160: mov [ebx+00000008h], ecx
  loc_00E14163: mov ecx, var_78
  loc_00E14166: mov [ebx+0000000Ch], ecx
  loc_00E14169: call [edx+00000028h]
  loc_00E1416C: test eax, eax
  loc_00E1416E: fnclex
  loc_00E14170: jge 00E14187h
  loc_00E14172: mov edx, var_C4
  loc_00E14178: push 00000028h
  loc_00E1417A: push 006E09E8h
  loc_00E1417F: push edx
  loc_00E14180: push eax
  loc_00E14181: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14187: mov eax, var_30
  loc_00E1418A: lea edx, var_54
  loc_00E1418D: push edx
  loc_00E1418E: push eax
  loc_00E1418F: mov ecx, [eax]
  loc_00E14191: mov ebx, eax
  loc_00E14193: call [ecx+00000034h]
  loc_00E14196: test eax, eax
  loc_00E14198: fnclex
  loc_00E1419A: jge 00E141ABh
  loc_00E1419C: push 00000034h
  loc_00E1419E: push 006E09F8h
  loc_00E141A3: push ebx
  loc_00E141A4: push eax
  loc_00E141A5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E141AB: lea eax, var_54
  loc_00E141AE: lea ecx, var_94
  loc_00E141B4: push eax
  loc_00E141B5: push ecx
  loc_00E141B6: mov var_8C, 006E0B40h ; "Perempuan"
  loc_00E141C0: mov var_94, 00008008h
  loc_00E141CA: call [00401108h] ; __vbaVarTstEq
  loc_00E141D0: mov bx, ax
  loc_00E141D3: lea edx, var_30
  loc_00E141D6: lea eax, var_2C
  loc_00E141D9: push edx
  loc_00E141DA: lea ecx, var_28
  loc_00E141DD: push eax
  loc_00E141DE: lea edx, var_24
  loc_00E141E1: push ecx
  loc_00E141E2: push edx
  loc_00E141E3: push 00000004h
  loc_00E141E5: call [00401048h] ; __vbaFreeObjList
  loc_00E141EB: lea eax, var_54
  loc_00E141EE: lea ecx, var_44
  loc_00E141F1: push eax
  loc_00E141F2: push ecx
  loc_00E141F3: push 00000002h
  loc_00E141F5: call [00401038h] ; __vbaFreeVarList
  loc_00E141FB: add esp, 00000020h
  loc_00E141FE: test bx, bx
  loc_00E14201: jz 00E142C6h
  loc_00E14207: mov edx, [esi]
  loc_00E14209: push esi
  loc_00E1420A: call [edx+0000037Ch]
  loc_00E14210: push eax
  loc_00E14211: lea eax, var_24
  loc_00E14214: push eax
  loc_00E14215: call edi
  loc_00E14217: mov ebx, eax
  loc_00E14219: push FFFFFFFFh
  loc_00E1421B: push ebx
  loc_00E1421C: mov ecx, [ebx]
  loc_00E1421E: call [ecx+000000E4h]
  loc_00E14224: test eax, eax
  loc_00E14226: fnclex
  loc_00E14228: jge 00E1423Ch
  loc_00E1422A: push 000000E4h
  loc_00E1422F: push 006E03D4h
  loc_00E14234: push ebx
  loc_00E14235: push eax
  loc_00E14236: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1423C: lea ecx, var_24
  loc_00E1423F: call [00401254h] ; __vbaFreeObj
  loc_00E14245: mov edx, [esi]
  loc_00E14247: push esi
  loc_00E14248: call [edx+0000037Ch]
  loc_00E1424E: push eax
  loc_00E1424F: lea eax, var_24
  loc_00E14252: push eax
  loc_00E14253: call edi
  loc_00E14255: mov ebx, eax
  loc_00E14257: push FFFFFFFFh
  loc_00E14259: push ebx
  loc_00E1425A: mov ecx, [ebx]
  loc_00E1425C: call [ecx+0000009Ch]
  loc_00E14262: test eax, eax
  loc_00E14264: fnclex
  loc_00E14266: jge 00E1427Ah
  loc_00E14268: push 0000009Ch
  loc_00E1426D: push 006E03D4h
  loc_00E14272: push ebx
  loc_00E14273: push eax
  loc_00E14274: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1427A: lea ecx, var_24
  loc_00E1427D: call [00401254h] ; __vbaFreeObj
  loc_00E14283: mov edx, [esi]
  loc_00E14285: push esi
  loc_00E14286: call [edx+00000380h]
  loc_00E1428C: push eax
  loc_00E1428D: lea eax, var_24
  loc_00E14290: push eax
  loc_00E14291: call edi
  loc_00E14293: mov ebx, eax
  loc_00E14295: push FFFFFFFFh
  loc_00E14297: push ebx
  loc_00E14298: mov ecx, [ebx]
  loc_00E1429A: call [ecx+0000009Ch]
  loc_00E142A0: test eax, eax
  loc_00E142A2: fnclex
  loc_00E142A4: jge 00E142B8h
  loc_00E142A6: push 0000009Ch
  loc_00E142AB: push 006E03D4h
  loc_00E142B0: push ebx
  loc_00E142B1: push eax
  loc_00E142B2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E142B8: lea ecx, var_24
  loc_00E142BB: call [00401254h] ; __vbaFreeObj
  loc_00E142C1: jmp 00E14350h
  loc_00E142C6: mov ebx, [004011F0h] ; __vbaVarDup
  loc_00E142CC: mov ecx, 0000000Ah
  loc_00E142D1: mov eax, 80020004h
  loc_00E142D6: mov var_74, ecx
  loc_00E142D9: mov var_64, ecx
  loc_00E142DC: lea edx, var_94
  loc_00E142E2: lea ecx, var_54
  loc_00E142E5: mov var_6C, eax
  loc_00E142E8: mov var_5C, eax
  loc_00E142EB: mov var_8C, 006E07B8h ; "ERROR 001"
  loc_00E142F5: mov var_94, 00000008h
  loc_00E142FF: call ebx
  loc_00E14301: lea edx, var_84
  loc_00E14307: lea ecx, var_44
  loc_00E1430A: mov var_7C, 006E0790h ; "Ada Yang Error !"
  loc_00E14311: mov var_84, 00000008h
  loc_00E1431B: call ebx
  loc_00E1431D: lea edx, var_74
  loc_00E14320: lea eax, var_64
  loc_00E14323: push edx
  loc_00E14324: lea ecx, var_54
  loc_00E14327: push eax
  loc_00E14328: push ecx
  loc_00E14329: lea edx, var_44
  loc_00E1432C: push 00000040h
  loc_00E1432E: push edx
  loc_00E1432F: call [004010A8h] ; rtcMsgBox
  loc_00E14335: lea eax, var_74
  loc_00E14338: lea ecx, var_64
  loc_00E1433B: push eax
  loc_00E1433C: lea edx, var_54
  loc_00E1433F: push ecx
  loc_00E14340: lea eax, var_44
  loc_00E14343: push edx
  loc_00E14344: push eax
  loc_00E14345: push 00000004h
  loc_00E14347: call [00401038h] ; __vbaFreeVarList
  loc_00E1434D: add esp, 00000014h
  loc_00E14350: mov ecx, [esi]
  loc_00E14352: push esi
  loc_00E14353: call [ecx+00000390h]
  loc_00E14359: lea edx, var_34
  loc_00E1435C: push eax
  loc_00E1435D: push edx
  loc_00E1435E: call edi
  loc_00E14360: push 006DCBD8h
  loc_00E14365: mov var_D4, eax
  loc_00E1436B: mov eax, [esi]
  loc_00E1436D: push 00000000h
  loc_00E1436F: push 00000018h
  loc_00E14371: push esi
  loc_00E14372: call [eax+00000490h]
  loc_00E14378: lea ecx, var_24
  loc_00E1437B: push eax
  loc_00E1437C: push ecx
  loc_00E1437D: call edi
  loc_00E1437F: lea edx, var_44
  loc_00E14382: push eax
  loc_00E14383: push edx
  loc_00E14384: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1438A: add esp, 00000010h
  loc_00E1438D: push eax
  loc_00E1438E: call [00401120h] ; __vbaCastObjVar
  loc_00E14394: push eax
  loc_00E14395: lea eax, var_28
  loc_00E14398: push eax
  loc_00E14399: call edi
  loc_00E1439B: mov ebx, eax
  loc_00E1439D: lea edx, var_2C
  loc_00E143A0: push edx
  loc_00E143A1: push ebx
  loc_00E143A2: mov ecx, [ebx]
  loc_00E143A4: call [ecx+00000054h]
  loc_00E143A7: test eax, eax
  loc_00E143A9: fnclex
  loc_00E143AB: jge 00E143BCh
  loc_00E143AD: push 00000054h
  loc_00E143AF: push 006DCBE8h
  loc_00E143B4: push ebx
  loc_00E143B5: push eax
  loc_00E143B6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E143BC: lea ebx, var_30
  loc_00E143BF: mov eax, var_2C
  loc_00E143C2: push ebx
  loc_00E143C3: mov ecx, 00000002h
  loc_00E143C8: sub esp, 00000010h
  loc_00E143CB: mov var_84, ecx
  loc_00E143D1: mov ebx, esp
  loc_00E143D3: mov var_7C, 00000003h
  loc_00E143DA: mov edx, [eax]
  loc_00E143DC: push eax
  loc_00E143DD: mov [ebx], ecx
  loc_00E143DF: mov ecx, var_80
  loc_00E143E2: mov var_C4, eax
  loc_00E143E8: mov [ebx+00000004h], ecx
  loc_00E143EB: mov ecx, var_7C
  loc_00E143EE: mov [ebx+00000008h], ecx
  loc_00E143F1: mov ecx, var_78
  loc_00E143F4: mov [ebx+0000000Ch], ecx
  loc_00E143F7: call [edx+00000028h]
  loc_00E143FA: test eax, eax
  loc_00E143FC: fnclex
  loc_00E143FE: jge 00E14415h
  loc_00E14400: mov edx, var_C4
  loc_00E14406: push 00000028h
  loc_00E14408: push 006E09E8h
  loc_00E1440D: push edx
  loc_00E1440E: push eax
  loc_00E1440F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14415: mov eax, var_30
  loc_00E14418: lea edx, var_54
  loc_00E1441B: push edx
  loc_00E1441C: push eax
  loc_00E1441D: mov ecx, [eax]
  loc_00E1441F: mov ebx, eax
  loc_00E14421: call [ecx+00000034h]
  loc_00E14424: test eax, eax
  loc_00E14426: fnclex
  loc_00E14428: jge 00E14439h
  loc_00E1442A: push 00000034h
  loc_00E1442C: push 006E09F8h
  loc_00E14431: push ebx
  loc_00E14432: push eax
  loc_00E14433: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14439: mov eax, var_D4
  loc_00E1443F: lea ecx, var_54
  loc_00E14442: push ecx
  loc_00E14443: mov ebx, [eax]
  loc_00E14445: call [00401034h] ; __vbaStrVarMove
  loc_00E1444B: mov edx, eax
  loc_00E1444D: lea ecx, var_18
  loc_00E14450: call [00401228h] ; __vbaStrMove
  loc_00E14456: mov edx, ebx
  loc_00E14458: mov ebx, var_D4
  loc_00E1445E: push eax
  loc_00E1445F: push ebx
  loc_00E14460: call [edx+000000A4h]
  loc_00E14466: test eax, eax
  loc_00E14468: fnclex
  loc_00E1446A: jge 00E1447Eh
  loc_00E1446C: push 000000A4h
  loc_00E14471: push 006DCB70h
  loc_00E14476: push ebx
  loc_00E14477: push eax
  loc_00E14478: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1447E: lea ecx, var_18
  loc_00E14481: call [00401258h] ; __vbaFreeStr
  loc_00E14487: lea eax, var_34
  loc_00E1448A: lea ecx, var_30
  loc_00E1448D: push eax
  loc_00E1448E: lea edx, var_2C
  loc_00E14491: push ecx
  loc_00E14492: lea eax, var_28
  loc_00E14495: push edx
  loc_00E14496: lea ecx, var_24
  loc_00E14499: push eax
  loc_00E1449A: push ecx
  loc_00E1449B: push 00000005h
  loc_00E1449D: call [00401048h] ; __vbaFreeObjList
  loc_00E144A3: lea edx, var_54
  loc_00E144A6: lea eax, var_44
  loc_00E144A9: push edx
  loc_00E144AA: push eax
  loc_00E144AB: push 00000002h
  loc_00E144AD: call [00401038h] ; __vbaFreeVarList
  loc_00E144B3: mov ecx, [esi]
  loc_00E144B5: add esp, 00000024h
  loc_00E144B8: push esi
  loc_00E144B9: call [ecx+0000038Ch]
  loc_00E144BF: lea edx, var_34
  loc_00E144C2: push eax
  loc_00E144C3: push edx
  loc_00E144C4: call edi
  loc_00E144C6: push 006DCBD8h
  loc_00E144CB: mov var_D4, eax
  loc_00E144D1: mov eax, [esi]
  loc_00E144D3: push 00000000h
  loc_00E144D5: push 00000018h
  loc_00E144D7: push esi
  loc_00E144D8: call [eax+00000490h]
  loc_00E144DE: lea ecx, var_24
  loc_00E144E1: push eax
  loc_00E144E2: push ecx
  loc_00E144E3: call edi
  loc_00E144E5: lea edx, var_44
  loc_00E144E8: push eax
  loc_00E144E9: push edx
  loc_00E144EA: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E144F0: add esp, 00000010h
  loc_00E144F3: push eax
  loc_00E144F4: call [00401120h] ; __vbaCastObjVar
  loc_00E144FA: push eax
  loc_00E144FB: lea eax, var_28
  loc_00E144FE: push eax
  loc_00E144FF: call edi
  loc_00E14501: mov ebx, eax
  loc_00E14503: lea edx, var_2C
  loc_00E14506: push edx
  loc_00E14507: push ebx
  loc_00E14508: mov ecx, [ebx]
  loc_00E1450A: call [ecx+00000054h]
  loc_00E1450D: test eax, eax
  loc_00E1450F: fnclex
  loc_00E14511: jge 00E14522h
  loc_00E14513: push 00000054h
  loc_00E14515: push 006DCBE8h
  loc_00E1451A: push ebx
  loc_00E1451B: push eax
  loc_00E1451C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14522: lea ebx, var_30
  loc_00E14525: mov eax, var_2C
  loc_00E14528: push ebx
  loc_00E14529: mov ecx, 00000002h
  loc_00E1452E: sub esp, 00000010h
  loc_00E14531: mov var_84, ecx
  loc_00E14537: mov ebx, esp
  loc_00E14539: mov var_7C, 00000004h
  loc_00E14540: mov edx, [eax]
  loc_00E14542: push eax
  loc_00E14543: mov [ebx], ecx
  loc_00E14545: mov ecx, var_80
  loc_00E14548: mov var_C4, eax
  loc_00E1454E: mov [ebx+00000004h], ecx
  loc_00E14551: mov ecx, var_7C
  loc_00E14554: mov [ebx+00000008h], ecx
  loc_00E14557: mov ecx, var_78
  loc_00E1455A: mov [ebx+0000000Ch], ecx
  loc_00E1455D: call [edx+00000028h]
  loc_00E14560: test eax, eax
  loc_00E14562: fnclex
  loc_00E14564: jge 00E1457Bh
  loc_00E14566: mov edx, var_C4
  loc_00E1456C: push 00000028h
  loc_00E1456E: push 006E09E8h
  loc_00E14573: push edx
  loc_00E14574: push eax
  loc_00E14575: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1457B: mov eax, var_30
  loc_00E1457E: lea edx, var_54
  loc_00E14581: push edx
  loc_00E14582: push eax
  loc_00E14583: mov ecx, [eax]
  loc_00E14585: mov ebx, eax
  loc_00E14587: call [ecx+00000034h]
  loc_00E1458A: test eax, eax
  loc_00E1458C: fnclex
  loc_00E1458E: jge 00E1459Fh
  loc_00E14590: push 00000034h
  loc_00E14592: push 006E09F8h
  loc_00E14597: push ebx
  loc_00E14598: push eax
  loc_00E14599: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1459F: mov eax, var_D4
  loc_00E145A5: lea ecx, var_54
  loc_00E145A8: push ecx
  loc_00E145A9: mov ebx, [eax]
  loc_00E145AB: call [00401034h] ; __vbaStrVarMove
  loc_00E145B1: mov edx, eax
  loc_00E145B3: lea ecx, var_18
  loc_00E145B6: call [00401228h] ; __vbaStrMove
  loc_00E145BC: mov edx, ebx
  loc_00E145BE: mov ebx, var_D4
  loc_00E145C4: push eax
  loc_00E145C5: push ebx
  loc_00E145C6: call [edx+000000A4h]
  loc_00E145CC: test eax, eax
  loc_00E145CE: fnclex
  loc_00E145D0: jge 00E145E4h
  loc_00E145D2: push 000000A4h
  loc_00E145D7: push 006DCB70h
  loc_00E145DC: push ebx
  loc_00E145DD: push eax
  loc_00E145DE: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E145E4: lea ecx, var_18
  loc_00E145E7: call [00401258h] ; __vbaFreeStr
  loc_00E145ED: lea eax, var_34
  loc_00E145F0: lea ecx, var_30
  loc_00E145F3: push eax
  loc_00E145F4: lea edx, var_2C
  loc_00E145F7: push ecx
  loc_00E145F8: lea eax, var_28
  loc_00E145FB: push edx
  loc_00E145FC: lea ecx, var_24
  loc_00E145FF: push eax
  loc_00E14600: push ecx
  loc_00E14601: push 00000005h
  loc_00E14603: call [00401048h] ; __vbaFreeObjList
  loc_00E14609: lea edx, var_54
  loc_00E1460C: lea eax, var_44
  loc_00E1460F: push edx
  loc_00E14610: push eax
  loc_00E14611: push 00000002h
  loc_00E14613: call [00401038h] ; __vbaFreeVarList
  loc_00E14619: mov ecx, [esi]
  loc_00E1461B: add esp, 00000024h
  loc_00E1461E: push esi
  loc_00E1461F: call [ecx+00000388h]
  loc_00E14625: lea edx, var_34
  loc_00E14628: push eax
  loc_00E14629: push edx
  loc_00E1462A: call edi
  loc_00E1462C: push 006DCBD8h
  loc_00E14631: mov var_D4, eax
  loc_00E14637: mov eax, [esi]
  loc_00E14639: push 00000000h
  loc_00E1463B: push 00000018h
  loc_00E1463D: push esi
  loc_00E1463E: call [eax+00000490h]
  loc_00E14644: lea ecx, var_24
  loc_00E14647: push eax
  loc_00E14648: push ecx
  loc_00E14649: call edi
  loc_00E1464B: lea edx, var_44
  loc_00E1464E: push eax
  loc_00E1464F: push edx
  loc_00E14650: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E14656: add esp, 00000010h
  loc_00E14659: push eax
  loc_00E1465A: call [00401120h] ; __vbaCastObjVar
  loc_00E14660: push eax
  loc_00E14661: lea eax, var_28
  loc_00E14664: push eax
  loc_00E14665: call edi
  loc_00E14667: mov ebx, eax
  loc_00E14669: lea edx, var_2C
  loc_00E1466C: push edx
  loc_00E1466D: push ebx
  loc_00E1466E: mov ecx, [ebx]
  loc_00E14670: call [ecx+00000054h]
  loc_00E14673: test eax, eax
  loc_00E14675: fnclex
  loc_00E14677: jge 00E14688h
  loc_00E14679: push 00000054h
  loc_00E1467B: push 006DCBE8h
  loc_00E14680: push ebx
  loc_00E14681: push eax
  loc_00E14682: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14688: lea ebx, var_30
  loc_00E1468B: mov eax, var_2C
  loc_00E1468E: push ebx
  loc_00E1468F: mov ecx, 00000002h
  loc_00E14694: sub esp, 00000010h
  loc_00E14697: mov var_84, ecx
  loc_00E1469D: mov ebx, esp
  loc_00E1469F: mov var_7C, 00000005h
  loc_00E146A6: mov edx, [eax]
  loc_00E146A8: push eax
  loc_00E146A9: mov [ebx], ecx
  loc_00E146AB: mov ecx, var_80
  loc_00E146AE: mov var_C4, eax
  loc_00E146B4: mov [ebx+00000004h], ecx
  loc_00E146B7: mov ecx, var_7C
  loc_00E146BA: mov [ebx+00000008h], ecx
  loc_00E146BD: mov ecx, var_78
  loc_00E146C0: mov [ebx+0000000Ch], ecx
  loc_00E146C3: call [edx+00000028h]
  loc_00E146C6: test eax, eax
  loc_00E146C8: fnclex
  loc_00E146CA: jge 00E146E1h
  loc_00E146CC: mov edx, var_C4
  loc_00E146D2: push 00000028h
  loc_00E146D4: push 006E09E8h
  loc_00E146D9: push edx
  loc_00E146DA: push eax
  loc_00E146DB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E146E1: mov eax, var_30
  loc_00E146E4: lea edx, var_54
  loc_00E146E7: push edx
  loc_00E146E8: push eax
  loc_00E146E9: mov ecx, [eax]
  loc_00E146EB: mov ebx, eax
  loc_00E146ED: call [ecx+00000034h]
  loc_00E146F0: test eax, eax
  loc_00E146F2: fnclex
  loc_00E146F4: jge 00E14705h
  loc_00E146F6: push 00000034h
  loc_00E146F8: push 006E09F8h
  loc_00E146FD: push ebx
  loc_00E146FE: push eax
  loc_00E146FF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14705: mov eax, var_D4
  loc_00E1470B: lea ecx, var_54
  loc_00E1470E: push ecx
  loc_00E1470F: mov ebx, [eax]
  loc_00E14711: call [00401034h] ; __vbaStrVarMove
  loc_00E14717: mov edx, eax
  loc_00E14719: lea ecx, var_18
  loc_00E1471C: call [00401228h] ; __vbaStrMove
  loc_00E14722: mov edx, ebx
  loc_00E14724: mov ebx, var_D4
  loc_00E1472A: push eax
  loc_00E1472B: push ebx
  loc_00E1472C: call [edx+000000A4h]
  loc_00E14732: test eax, eax
  loc_00E14734: fnclex
  loc_00E14736: jge 00E1474Ah
  loc_00E14738: push 000000A4h
  loc_00E1473D: push 006DCB70h
  loc_00E14742: push ebx
  loc_00E14743: push eax
  loc_00E14744: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1474A: lea ecx, var_18
  loc_00E1474D: call [00401258h] ; __vbaFreeStr
  loc_00E14753: lea eax, var_34
  loc_00E14756: lea ecx, var_30
  loc_00E14759: push eax
  loc_00E1475A: lea edx, var_2C
  loc_00E1475D: push ecx
  loc_00E1475E: lea eax, var_28
  loc_00E14761: push edx
  loc_00E14762: lea ecx, var_24
  loc_00E14765: push eax
  loc_00E14766: push ecx
  loc_00E14767: push 00000005h
  loc_00E14769: call [00401048h] ; __vbaFreeObjList
  loc_00E1476F: lea edx, var_54
  loc_00E14772: lea eax, var_44
  loc_00E14775: push edx
  loc_00E14776: push eax
  loc_00E14777: push 00000002h
  loc_00E14779: call [00401038h] ; __vbaFreeVarList
  loc_00E1477F: mov ecx, [esi]
  loc_00E14781: add esp, 00000024h
  loc_00E14784: push 006DCBD8h
  loc_00E14789: push 00000000h
  loc_00E1478B: push 00000018h
  loc_00E1478D: push esi
  loc_00E1478E: call [ecx+00000490h]
  loc_00E14794: lea edx, var_24
  loc_00E14797: push eax
  loc_00E14798: push edx
  loc_00E14799: call edi
  loc_00E1479B: push eax
  loc_00E1479C: lea eax, var_44
  loc_00E1479F: push eax
  loc_00E147A0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E147A6: add esp, 00000010h
  loc_00E147A9: push eax
  loc_00E147AA: call [00401120h] ; __vbaCastObjVar
  loc_00E147B0: lea ecx, var_28
  loc_00E147B3: push eax
  loc_00E147B4: push ecx
  loc_00E147B5: call edi
  loc_00E147B7: mov ebx, eax
  loc_00E147B9: lea eax, var_2C
  loc_00E147BC: push eax
  loc_00E147BD: push ebx
  loc_00E147BE: mov edx, [ebx]
  loc_00E147C0: call [edx+00000054h]
  loc_00E147C3: test eax, eax
  loc_00E147C5: fnclex
  loc_00E147C7: jge 00E147D8h
  loc_00E147C9: push 00000054h
  loc_00E147CB: push 006DCBE8h
  loc_00E147D0: push ebx
  loc_00E147D1: push eax
  loc_00E147D2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E147D8: lea ebx, var_30
  loc_00E147DB: mov eax, var_2C
  loc_00E147DE: push ebx
  loc_00E147DF: mov ecx, 00000002h
  loc_00E147E4: sub esp, 00000010h
  loc_00E147E7: mov var_84, ecx
  loc_00E147ED: mov ebx, esp
  loc_00E147EF: mov var_7C, 00000006h
  loc_00E147F6: mov edx, [eax]
  loc_00E147F8: push eax
  loc_00E147F9: mov [ebx], ecx
  loc_00E147FB: mov ecx, var_80
  loc_00E147FE: mov var_C4, eax
  loc_00E14804: mov [ebx+00000004h], ecx
  loc_00E14807: mov ecx, var_7C
  loc_00E1480A: mov [ebx+00000008h], ecx
  loc_00E1480D: mov ecx, var_78
  loc_00E14810: mov [ebx+0000000Ch], ecx
  loc_00E14813: call [edx+00000028h]
  loc_00E14816: test eax, eax
  loc_00E14818: fnclex
  loc_00E1481A: jge 00E14831h
  loc_00E1481C: mov edx, var_C4
  loc_00E14822: push 00000028h
  loc_00E14824: push 006E09E8h
  loc_00E14829: push edx
  loc_00E1482A: push eax
  loc_00E1482B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14831: sub esp, 00000010h
  loc_00E14834: mov eax, var_30
  loc_00E14837: mov edx, esp
  loc_00E14839: mov ecx, 00000009h
  loc_00E1483E: mov var_54, ecx
  loc_00E14841: mov var_4C, eax
  loc_00E14844: mov [edx], ecx
  loc_00E14846: mov ecx, var_50
  loc_00E14849: push 00000014h
  loc_00E1484B: push esi
  loc_00E1484C: mov [edx+00000004h], ecx
  loc_00E1484F: mov ecx, [esi]
  loc_00E14851: mov var_30, 00000000h
  loc_00E14858: mov [edx+00000008h], eax
  loc_00E1485B: mov eax, var_48
  loc_00E1485E: mov [edx+0000000Ch], eax
  loc_00E14861: call [ecx+0000048Ch]
  loc_00E14867: lea edx, var_34
  loc_00E1486A: push eax
  loc_00E1486B: push edx
  loc_00E1486C: call edi
  loc_00E1486E: push eax
  loc_00E1486F: call [00401238h] ; __vbaLateIdSt
  loc_00E14875: lea eax, var_34
  loc_00E14878: lea ecx, var_2C
  loc_00E1487B: push eax
  loc_00E1487C: lea edx, var_28
  loc_00E1487F: push ecx
  loc_00E14880: lea eax, var_24
  loc_00E14883: push edx
  loc_00E14884: push eax
  loc_00E14885: push 00000004h
  loc_00E14887: call [00401048h] ; __vbaFreeObjList
  loc_00E1488D: lea ecx, var_54
  loc_00E14890: lea edx, var_44
  loc_00E14893: push ecx
  loc_00E14894: push edx
  loc_00E14895: push 00000002h
  loc_00E14897: call [00401038h] ; __vbaFreeVarList
  loc_00E1489D: mov eax, [esi]
  loc_00E1489F: add esp, 00000020h
  loc_00E148A2: push esi
  loc_00E148A3: call [eax+00000378h]
  loc_00E148A9: lea ecx, var_34
  loc_00E148AC: push eax
  loc_00E148AD: push ecx
  loc_00E148AE: call edi
  loc_00E148B0: mov edx, [esi]
  loc_00E148B2: push 006DCBD8h
  loc_00E148B7: push 00000000h
  loc_00E148B9: push 00000018h
  loc_00E148BB: push esi
  loc_00E148BC: mov var_D4, eax
  loc_00E148C2: call [edx+00000490h]
  loc_00E148C8: push eax
  loc_00E148C9: lea eax, var_24
  loc_00E148CC: push eax
  loc_00E148CD: call edi
  loc_00E148CF: lea ecx, var_44
  loc_00E148D2: push eax
  loc_00E148D3: push ecx
  loc_00E148D4: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E148DA: add esp, 00000010h
  loc_00E148DD: push eax
  loc_00E148DE: call [00401120h] ; __vbaCastObjVar
  loc_00E148E4: lea edx, var_28
  loc_00E148E7: push eax
  loc_00E148E8: push edx
  loc_00E148E9: call edi
  loc_00E148EB: mov ebx, eax
  loc_00E148ED: lea ecx, var_2C
  loc_00E148F0: push ecx
  loc_00E148F1: push ebx
  loc_00E148F2: mov eax, [ebx]
  loc_00E148F4: call [eax+00000054h]
  loc_00E148F7: test eax, eax
  loc_00E148F9: fnclex
  loc_00E148FB: jge 00E1490Ch
  loc_00E148FD: push 00000054h
  loc_00E148FF: push 006DCBE8h
  loc_00E14904: push ebx
  loc_00E14905: push eax
  loc_00E14906: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1490C: lea ebx, var_30
  loc_00E1490F: mov eax, var_2C
  loc_00E14912: push ebx
  loc_00E14913: mov ecx, 00000002h
  loc_00E14918: sub esp, 00000010h
  loc_00E1491B: mov var_84, ecx
  loc_00E14921: mov ebx, esp
  loc_00E14923: mov var_7C, 00000007h
  loc_00E1492A: mov edx, [eax]
  loc_00E1492C: push eax
  loc_00E1492D: mov [ebx], ecx
  loc_00E1492F: mov ecx, var_80
  loc_00E14932: mov var_C4, eax
  loc_00E14938: mov [ebx+00000004h], ecx
  loc_00E1493B: mov ecx, var_7C
  loc_00E1493E: mov [ebx+00000008h], ecx
  loc_00E14941: mov ecx, var_78
  loc_00E14944: mov [ebx+0000000Ch], ecx
  loc_00E14947: call [edx+00000028h]
  loc_00E1494A: test eax, eax
  loc_00E1494C: fnclex
  loc_00E1494E: jge 00E14965h
  loc_00E14950: mov edx, var_C4
  loc_00E14956: push 00000028h
  loc_00E14958: push 006E09E8h
  loc_00E1495D: push edx
  loc_00E1495E: push eax
  loc_00E1495F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14965: mov eax, var_30
  loc_00E14968: lea edx, var_54
  loc_00E1496B: push edx
  loc_00E1496C: push eax
  loc_00E1496D: mov ecx, [eax]
  loc_00E1496F: mov ebx, eax
  loc_00E14971: call [ecx+00000034h]
  loc_00E14974: test eax, eax
  loc_00E14976: fnclex
  loc_00E14978: jge 00E14989h
  loc_00E1497A: push 00000034h
  loc_00E1497C: push 006E09F8h
  loc_00E14981: push ebx
  loc_00E14982: push eax
  loc_00E14983: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14989: mov eax, var_D4
  loc_00E1498F: lea ecx, var_54
  loc_00E14992: push ecx
  loc_00E14993: mov ebx, [eax]
  loc_00E14995: call [00401034h] ; __vbaStrVarMove
  loc_00E1499B: mov edx, eax
  loc_00E1499D: lea ecx, var_18
  loc_00E149A0: call [00401228h] ; __vbaStrMove
  loc_00E149A6: mov edx, ebx
  loc_00E149A8: mov ebx, var_D4
  loc_00E149AE: push eax
  loc_00E149AF: push ebx
  loc_00E149B0: call [edx+000000ACh]
  loc_00E149B6: test eax, eax
  loc_00E149B8: fnclex
  loc_00E149BA: jge 00E149CEh
  loc_00E149BC: push 000000ACh
  loc_00E149C1: push 006E0400h
  loc_00E149C6: push ebx
  loc_00E149C7: push eax
  loc_00E149C8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E149CE: lea ecx, var_18
  loc_00E149D1: call [00401258h] ; __vbaFreeStr
  loc_00E149D7: lea eax, var_34
  loc_00E149DA: lea ecx, var_30
  loc_00E149DD: push eax
  loc_00E149DE: lea edx, var_2C
  loc_00E149E1: push ecx
  loc_00E149E2: lea eax, var_28
  loc_00E149E5: push edx
  loc_00E149E6: lea ecx, var_24
  loc_00E149E9: push eax
  loc_00E149EA: push ecx
  loc_00E149EB: push 00000005h
  loc_00E149ED: call [00401048h] ; __vbaFreeObjList
  loc_00E149F3: lea edx, var_54
  loc_00E149F6: lea eax, var_44
  loc_00E149F9: push edx
  loc_00E149FA: push eax
  loc_00E149FB: push 00000002h
  loc_00E149FD: call [00401038h] ; __vbaFreeVarList
  loc_00E14A03: mov ecx, [esi]
  loc_00E14A05: add esp, 00000024h
  loc_00E14A08: push esi
  loc_00E14A09: call [ecx+00000384h]
  loc_00E14A0F: lea edx, var_34
  loc_00E14A12: push eax
  loc_00E14A13: push edx
  loc_00E14A14: call edi
  loc_00E14A16: push 006DCBD8h
  loc_00E14A1B: mov var_D4, eax
  loc_00E14A21: mov eax, [esi]
  loc_00E14A23: push 00000000h
  loc_00E14A25: push 00000018h
  loc_00E14A27: push esi
  loc_00E14A28: call [eax+00000490h]
  loc_00E14A2E: lea ecx, var_24
  loc_00E14A31: push eax
  loc_00E14A32: push ecx
  loc_00E14A33: call edi
  loc_00E14A35: lea edx, var_44
  loc_00E14A38: push eax
  loc_00E14A39: push edx
  loc_00E14A3A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E14A40: add esp, 00000010h
  loc_00E14A43: push eax
  loc_00E14A44: call [00401120h] ; __vbaCastObjVar
  loc_00E14A4A: push eax
  loc_00E14A4B: lea eax, var_28
  loc_00E14A4E: push eax
  loc_00E14A4F: call edi
  loc_00E14A51: mov ebx, eax
  loc_00E14A53: lea edx, var_2C
  loc_00E14A56: push edx
  loc_00E14A57: push ebx
  loc_00E14A58: mov ecx, [ebx]
  loc_00E14A5A: call [ecx+00000054h]
  loc_00E14A5D: test eax, eax
  loc_00E14A5F: fnclex
  loc_00E14A61: jge 00E14A72h
  loc_00E14A63: push 00000054h
  loc_00E14A65: push 006DCBE8h
  loc_00E14A6A: push ebx
  loc_00E14A6B: push eax
  loc_00E14A6C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14A72: lea ebx, var_30
  loc_00E14A75: mov eax, var_2C
  loc_00E14A78: push ebx
  loc_00E14A79: mov ecx, 00000002h
  loc_00E14A7E: sub esp, 00000010h
  loc_00E14A81: mov var_84, ecx
  loc_00E14A87: mov ebx, esp
  loc_00E14A89: mov var_7C, 00000008h
  loc_00E14A90: mov edx, [eax]
  loc_00E14A92: push eax
  loc_00E14A93: mov [ebx], ecx
  loc_00E14A95: mov ecx, var_80
  loc_00E14A98: mov var_C4, eax
  loc_00E14A9E: mov [ebx+00000004h], ecx
  loc_00E14AA1: mov ecx, var_7C
  loc_00E14AA4: mov [ebx+00000008h], ecx
  loc_00E14AA7: mov ecx, var_78
  loc_00E14AAA: mov [ebx+0000000Ch], ecx
  loc_00E14AAD: call [edx+00000028h]
  loc_00E14AB0: test eax, eax
  loc_00E14AB2: fnclex
  loc_00E14AB4: jge 00E14ACBh
  loc_00E14AB6: mov edx, var_C4
  loc_00E14ABC: push 00000028h
  loc_00E14ABE: push 006E09E8h
  loc_00E14AC3: push edx
  loc_00E14AC4: push eax
  loc_00E14AC5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14ACB: mov eax, var_30
  loc_00E14ACE: lea edx, var_54
  loc_00E14AD1: push edx
  loc_00E14AD2: push eax
  loc_00E14AD3: mov ecx, [eax]
  loc_00E14AD5: mov ebx, eax
  loc_00E14AD7: call [ecx+00000034h]
  loc_00E14ADA: test eax, eax
  loc_00E14ADC: fnclex
  loc_00E14ADE: jge 00E14AEFh
  loc_00E14AE0: push 00000034h
  loc_00E14AE2: push 006E09F8h
  loc_00E14AE7: push ebx
  loc_00E14AE8: push eax
  loc_00E14AE9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14AEF: mov eax, var_D4
  loc_00E14AF5: lea ecx, var_54
  loc_00E14AF8: push ecx
  loc_00E14AF9: mov ebx, [eax]
  loc_00E14AFB: call [00401034h] ; __vbaStrVarMove
  loc_00E14B01: mov edx, eax
  loc_00E14B03: lea ecx, var_18
  loc_00E14B06: call [00401228h] ; __vbaStrMove
  loc_00E14B0C: mov edx, ebx
  loc_00E14B0E: mov ebx, var_D4
  loc_00E14B14: push eax
  loc_00E14B15: push ebx
  loc_00E14B16: call [edx+000000A4h]
  loc_00E14B1C: test eax, eax
  loc_00E14B1E: fnclex
  loc_00E14B20: jge 00E14B34h
  loc_00E14B22: push 000000A4h
  loc_00E14B27: push 006DCB70h
  loc_00E14B2C: push ebx
  loc_00E14B2D: push eax
  loc_00E14B2E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14B34: lea ecx, var_18
  loc_00E14B37: call [00401258h] ; __vbaFreeStr
  loc_00E14B3D: lea eax, var_34
  loc_00E14B40: lea ecx, var_30
  loc_00E14B43: push eax
  loc_00E14B44: lea edx, var_2C
  loc_00E14B47: push ecx
  loc_00E14B48: lea eax, var_28
  loc_00E14B4B: push edx
  loc_00E14B4C: lea ecx, var_24
  loc_00E14B4F: push eax
  loc_00E14B50: push ecx
  loc_00E14B51: push 00000005h
  loc_00E14B53: call [00401048h] ; __vbaFreeObjList
  loc_00E14B59: lea edx, var_54
  loc_00E14B5C: lea eax, var_44
  loc_00E14B5F: push edx
  loc_00E14B60: push eax
  loc_00E14B61: push 00000002h
  loc_00E14B63: call [00401038h] ; __vbaFreeVarList
  loc_00E14B69: mov ecx, [esi]
  loc_00E14B6B: add esp, 00000024h
  loc_00E14B6E: push esi
  loc_00E14B6F: call [ecx+00000374h]
  loc_00E14B75: lea edx, var_34
  loc_00E14B78: push eax
  loc_00E14B79: push edx
  loc_00E14B7A: call edi
  loc_00E14B7C: push 006DCBD8h
  loc_00E14B81: mov var_D4, eax
  loc_00E14B87: mov eax, [esi]
  loc_00E14B89: push 00000000h
  loc_00E14B8B: push 00000018h
  loc_00E14B8D: push esi
  loc_00E14B8E: call [eax+00000490h]
  loc_00E14B94: lea ecx, var_24
  loc_00E14B97: push eax
  loc_00E14B98: push ecx
  loc_00E14B99: call edi
  loc_00E14B9B: lea edx, var_44
  loc_00E14B9E: push eax
  loc_00E14B9F: push edx
  loc_00E14BA0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E14BA6: add esp, 00000010h
  loc_00E14BA9: push eax
  loc_00E14BAA: call [00401120h] ; __vbaCastObjVar
  loc_00E14BB0: push eax
  loc_00E14BB1: lea eax, var_28
  loc_00E14BB4: push eax
  loc_00E14BB5: call edi
  loc_00E14BB7: mov ebx, eax
  loc_00E14BB9: lea edx, var_2C
  loc_00E14BBC: push edx
  loc_00E14BBD: push ebx
  loc_00E14BBE: mov ecx, [ebx]
  loc_00E14BC0: call [ecx+00000054h]
  loc_00E14BC3: test eax, eax
  loc_00E14BC5: fnclex
  loc_00E14BC7: jge 00E14BD8h
  loc_00E14BC9: push 00000054h
  loc_00E14BCB: push 006DCBE8h
  loc_00E14BD0: push ebx
  loc_00E14BD1: push eax
  loc_00E14BD2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14BD8: lea ebx, var_30
  loc_00E14BDB: mov eax, var_2C
  loc_00E14BDE: push ebx
  loc_00E14BDF: mov ecx, 00000002h
  loc_00E14BE4: sub esp, 00000010h
  loc_00E14BE7: mov var_84, ecx
  loc_00E14BED: mov ebx, esp
  loc_00E14BEF: mov var_7C, 00000009h
  loc_00E14BF6: mov edx, [eax]
  loc_00E14BF8: push eax
  loc_00E14BF9: mov [ebx], ecx
  loc_00E14BFB: mov ecx, var_80
  loc_00E14BFE: mov var_C4, eax
  loc_00E14C04: mov [ebx+00000004h], ecx
  loc_00E14C07: mov ecx, var_7C
  loc_00E14C0A: mov [ebx+00000008h], ecx
  loc_00E14C0D: mov ecx, var_78
  loc_00E14C10: mov [ebx+0000000Ch], ecx
  loc_00E14C13: call [edx+00000028h]
  loc_00E14C16: test eax, eax
  loc_00E14C18: fnclex
  loc_00E14C1A: jge 00E14C31h
  loc_00E14C1C: mov edx, var_C4
  loc_00E14C22: push 00000028h
  loc_00E14C24: push 006E09E8h
  loc_00E14C29: push edx
  loc_00E14C2A: push eax
  loc_00E14C2B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14C31: mov eax, var_30
  loc_00E14C34: lea edx, var_54
  loc_00E14C37: push edx
  loc_00E14C38: push eax
  loc_00E14C39: mov ecx, [eax]
  loc_00E14C3B: mov ebx, eax
  loc_00E14C3D: call [ecx+00000034h]
  loc_00E14C40: test eax, eax
  loc_00E14C42: fnclex
  loc_00E14C44: jge 00E14C55h
  loc_00E14C46: push 00000034h
  loc_00E14C48: push 006E09F8h
  loc_00E14C4D: push ebx
  loc_00E14C4E: push eax
  loc_00E14C4F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14C55: mov eax, var_D4
  loc_00E14C5B: lea ecx, var_54
  loc_00E14C5E: push ecx
  loc_00E14C5F: mov ebx, [eax]
  loc_00E14C61: call [00401034h] ; __vbaStrVarMove
  loc_00E14C67: mov edx, eax
  loc_00E14C69: lea ecx, var_18
  loc_00E14C6C: call [00401228h] ; __vbaStrMove
  loc_00E14C72: mov edx, ebx
  loc_00E14C74: mov ebx, var_D4
  loc_00E14C7A: push eax
  loc_00E14C7B: push ebx
  loc_00E14C7C: call [edx+000000A4h]
  loc_00E14C82: test eax, eax
  loc_00E14C84: fnclex
  loc_00E14C86: jge 00E14C9Ah
  loc_00E14C88: push 000000A4h
  loc_00E14C8D: push 006DCB70h
  loc_00E14C92: push ebx
  loc_00E14C93: push eax
  loc_00E14C94: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14C9A: lea ecx, var_18
  loc_00E14C9D: call [00401258h] ; __vbaFreeStr
  loc_00E14CA3: lea eax, var_34
  loc_00E14CA6: lea ecx, var_30
  loc_00E14CA9: push eax
  loc_00E14CAA: lea edx, var_2C
  loc_00E14CAD: push ecx
  loc_00E14CAE: lea eax, var_28
  loc_00E14CB1: push edx
  loc_00E14CB2: lea ecx, var_24
  loc_00E14CB5: push eax
  loc_00E14CB6: push ecx
  loc_00E14CB7: push 00000005h
  loc_00E14CB9: call [00401048h] ; __vbaFreeObjList
  loc_00E14CBF: lea edx, var_54
  loc_00E14CC2: lea eax, var_44
  loc_00E14CC5: push edx
  loc_00E14CC6: push eax
  loc_00E14CC7: push 00000002h
  loc_00E14CC9: call [00401038h] ; __vbaFreeVarList
  loc_00E14CCF: mov ecx, [esi]
  loc_00E14CD1: add esp, 00000024h
  loc_00E14CD4: push esi
  loc_00E14CD5: call [ecx+00000370h]
  loc_00E14CDB: lea edx, var_34
  loc_00E14CDE: push eax
  loc_00E14CDF: push edx
  loc_00E14CE0: call edi
  loc_00E14CE2: push 006DCBD8h
  loc_00E14CE7: mov var_D4, eax
  loc_00E14CED: mov eax, [esi]
  loc_00E14CEF: push 00000000h
  loc_00E14CF1: push 00000018h
  loc_00E14CF3: push esi
  loc_00E14CF4: call [eax+00000490h]
  loc_00E14CFA: lea ecx, var_24
  loc_00E14CFD: push eax
  loc_00E14CFE: push ecx
  loc_00E14CFF: call edi
  loc_00E14D01: lea edx, var_44
  loc_00E14D04: push eax
  loc_00E14D05: push edx
  loc_00E14D06: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E14D0C: add esp, 00000010h
  loc_00E14D0F: push eax
  loc_00E14D10: call [00401120h] ; __vbaCastObjVar
  loc_00E14D16: push eax
  loc_00E14D17: lea eax, var_28
  loc_00E14D1A: push eax
  loc_00E14D1B: call edi
  loc_00E14D1D: mov ebx, eax
  loc_00E14D1F: lea edx, var_2C
  loc_00E14D22: push edx
  loc_00E14D23: push ebx
  loc_00E14D24: mov ecx, [ebx]
  loc_00E14D26: call [ecx+00000054h]
  loc_00E14D29: test eax, eax
  loc_00E14D2B: fnclex
  loc_00E14D2D: jge 00E14D3Eh
  loc_00E14D2F: push 00000054h
  loc_00E14D31: push 006DCBE8h
  loc_00E14D36: push ebx
  loc_00E14D37: push eax
  loc_00E14D38: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14D3E: lea ebx, var_30
  loc_00E14D41: mov eax, var_2C
  loc_00E14D44: push ebx
  loc_00E14D45: mov ecx, 00000002h
  loc_00E14D4A: sub esp, 00000010h
  loc_00E14D4D: mov var_84, ecx
  loc_00E14D53: mov ebx, esp
  loc_00E14D55: mov var_7C, 0000000Ah
  loc_00E14D5C: mov edx, [eax]
  loc_00E14D5E: push eax
  loc_00E14D5F: mov [ebx], ecx
  loc_00E14D61: mov ecx, var_80
  loc_00E14D64: mov var_C4, eax
  loc_00E14D6A: mov [ebx+00000004h], ecx
  loc_00E14D6D: mov ecx, var_7C
  loc_00E14D70: mov [ebx+00000008h], ecx
  loc_00E14D73: mov ecx, var_78
  loc_00E14D76: mov [ebx+0000000Ch], ecx
  loc_00E14D79: call [edx+00000028h]
  loc_00E14D7C: test eax, eax
  loc_00E14D7E: fnclex
  loc_00E14D80: jge 00E14D97h
  loc_00E14D82: mov edx, var_C4
  loc_00E14D88: push 00000028h
  loc_00E14D8A: push 006E09E8h
  loc_00E14D8F: push edx
  loc_00E14D90: push eax
  loc_00E14D91: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14D97: mov eax, var_30
  loc_00E14D9A: lea edx, var_54
  loc_00E14D9D: push edx
  loc_00E14D9E: push eax
  loc_00E14D9F: mov ecx, [eax]
  loc_00E14DA1: mov ebx, eax
  loc_00E14DA3: call [ecx+00000034h]
  loc_00E14DA6: test eax, eax
  loc_00E14DA8: fnclex
  loc_00E14DAA: jge 00E14DBBh
  loc_00E14DAC: push 00000034h
  loc_00E14DAE: push 006E09F8h
  loc_00E14DB3: push ebx
  loc_00E14DB4: push eax
  loc_00E14DB5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14DBB: mov eax, var_D4
  loc_00E14DC1: lea ecx, var_54
  loc_00E14DC4: push ecx
  loc_00E14DC5: mov ebx, [eax]
  loc_00E14DC7: call [00401034h] ; __vbaStrVarMove
  loc_00E14DCD: mov edx, eax
  loc_00E14DCF: lea ecx, var_18
  loc_00E14DD2: call [00401228h] ; __vbaStrMove
  loc_00E14DD8: mov edx, ebx
  loc_00E14DDA: mov ebx, var_D4
  loc_00E14DE0: push eax
  loc_00E14DE1: push ebx
  loc_00E14DE2: call [edx+000000A4h]
  loc_00E14DE8: test eax, eax
  loc_00E14DEA: fnclex
  loc_00E14DEC: jge 00E14E00h
  loc_00E14DEE: push 000000A4h
  loc_00E14DF3: push 006DCB70h
  loc_00E14DF8: push ebx
  loc_00E14DF9: push eax
  loc_00E14DFA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14E00: lea ecx, var_18
  loc_00E14E03: call [00401258h] ; __vbaFreeStr
  loc_00E14E09: lea eax, var_34
  loc_00E14E0C: lea ecx, var_30
  loc_00E14E0F: push eax
  loc_00E14E10: lea edx, var_2C
  loc_00E14E13: push ecx
  loc_00E14E14: lea eax, var_28
  loc_00E14E17: push edx
  loc_00E14E18: lea ecx, var_24
  loc_00E14E1B: push eax
  loc_00E14E1C: push ecx
  loc_00E14E1D: push 00000005h
  loc_00E14E1F: call [00401048h] ; __vbaFreeObjList
  loc_00E14E25: lea edx, var_54
  loc_00E14E28: lea eax, var_44
  loc_00E14E2B: push edx
  loc_00E14E2C: push eax
  loc_00E14E2D: push 00000002h
  loc_00E14E2F: call [00401038h] ; __vbaFreeVarList
  loc_00E14E35: mov ecx, [esi]
  loc_00E14E37: add esp, 00000024h
  loc_00E14E3A: push esi
  loc_00E14E3B: call [ecx+0000036Ch]
  loc_00E14E41: lea edx, var_34
  loc_00E14E44: push eax
  loc_00E14E45: push edx
  loc_00E14E46: call edi
  loc_00E14E48: push 006DCBD8h
  loc_00E14E4D: mov var_D4, eax
  loc_00E14E53: mov eax, [esi]
  loc_00E14E55: push 00000000h
  loc_00E14E57: push 00000018h
  loc_00E14E59: push esi
  loc_00E14E5A: call [eax+00000490h]
  loc_00E14E60: lea ecx, var_24
  loc_00E14E63: push eax
  loc_00E14E64: push ecx
  loc_00E14E65: call edi
  loc_00E14E67: lea edx, var_44
  loc_00E14E6A: push eax
  loc_00E14E6B: push edx
  loc_00E14E6C: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E14E72: add esp, 00000010h
  loc_00E14E75: push eax
  loc_00E14E76: call [00401120h] ; __vbaCastObjVar
  loc_00E14E7C: push eax
  loc_00E14E7D: lea eax, var_28
  loc_00E14E80: push eax
  loc_00E14E81: call edi
  loc_00E14E83: mov ebx, eax
  loc_00E14E85: lea edx, var_2C
  loc_00E14E88: push edx
  loc_00E14E89: push ebx
  loc_00E14E8A: mov ecx, [ebx]
  loc_00E14E8C: call [ecx+00000054h]
  loc_00E14E8F: test eax, eax
  loc_00E14E91: fnclex
  loc_00E14E93: jge 00E14EA4h
  loc_00E14E95: push 00000054h
  loc_00E14E97: push 006DCBE8h
  loc_00E14E9C: push ebx
  loc_00E14E9D: push eax
  loc_00E14E9E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14EA4: lea ebx, var_30
  loc_00E14EA7: mov eax, var_2C
  loc_00E14EAA: push ebx
  loc_00E14EAB: mov ecx, 00000002h
  loc_00E14EB0: sub esp, 00000010h
  loc_00E14EB3: mov var_84, ecx
  loc_00E14EB9: mov ebx, esp
  loc_00E14EBB: mov var_7C, 0000000Bh
  loc_00E14EC2: mov edx, [eax]
  loc_00E14EC4: push eax
  loc_00E14EC5: mov [ebx], ecx
  loc_00E14EC7: mov ecx, var_80
  loc_00E14ECA: mov var_C4, eax
  loc_00E14ED0: mov [ebx+00000004h], ecx
  loc_00E14ED3: mov ecx, var_7C
  loc_00E14ED6: mov [ebx+00000008h], ecx
  loc_00E14ED9: mov ecx, var_78
  loc_00E14EDC: mov [ebx+0000000Ch], ecx
  loc_00E14EDF: call [edx+00000028h]
  loc_00E14EE2: test eax, eax
  loc_00E14EE4: fnclex
  loc_00E14EE6: jge 00E14EFDh
  loc_00E14EE8: mov edx, var_C4
  loc_00E14EEE: push 00000028h
  loc_00E14EF0: push 006E09E8h
  loc_00E14EF5: push edx
  loc_00E14EF6: push eax
  loc_00E14EF7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14EFD: mov eax, var_30
  loc_00E14F00: lea edx, var_54
  loc_00E14F03: push edx
  loc_00E14F04: push eax
  loc_00E14F05: mov ecx, [eax]
  loc_00E14F07: mov ebx, eax
  loc_00E14F09: call [ecx+00000034h]
  loc_00E14F0C: test eax, eax
  loc_00E14F0E: fnclex
  loc_00E14F10: jge 00E14F21h
  loc_00E14F12: push 00000034h
  loc_00E14F14: push 006E09F8h
  loc_00E14F19: push ebx
  loc_00E14F1A: push eax
  loc_00E14F1B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14F21: mov eax, var_D4
  loc_00E14F27: lea ecx, var_54
  loc_00E14F2A: push ecx
  loc_00E14F2B: mov ebx, [eax]
  loc_00E14F2D: call [00401034h] ; __vbaStrVarMove
  loc_00E14F33: mov edx, eax
  loc_00E14F35: lea ecx, var_18
  loc_00E14F38: call [00401228h] ; __vbaStrMove
  loc_00E14F3E: mov edx, ebx
  loc_00E14F40: mov ebx, var_D4
  loc_00E14F46: push eax
  loc_00E14F47: push ebx
  loc_00E14F48: call [edx+000000A4h]
  loc_00E14F4E: test eax, eax
  loc_00E14F50: fnclex
  loc_00E14F52: jge 00E14F66h
  loc_00E14F54: push 000000A4h
  loc_00E14F59: push 006DCB70h
  loc_00E14F5E: push ebx
  loc_00E14F5F: push eax
  loc_00E14F60: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E14F66: lea ecx, var_18
  loc_00E14F69: call [00401258h] ; __vbaFreeStr
  loc_00E14F6F: lea eax, var_34
  loc_00E14F72: lea ecx, var_30
  loc_00E14F75: push eax
  loc_00E14F76: lea edx, var_2C
  loc_00E14F79: push ecx
  loc_00E14F7A: lea eax, var_28
  loc_00E14F7D: push edx
  loc_00E14F7E: lea ecx, var_24
  loc_00E14F81: push eax
  loc_00E14F82: push ecx
  loc_00E14F83: push 00000005h
  loc_00E14F85: call [00401048h] ; __vbaFreeObjList
  loc_00E14F8B: lea edx, var_54
  loc_00E14F8E: lea eax, var_44
  loc_00E14F91: push edx
  loc_00E14F92: push eax
  loc_00E14F93: push 00000002h
  loc_00E14F95: call [00401038h] ; __vbaFreeVarList
  loc_00E14F9B: mov ecx, [esi]
  loc_00E14F9D: add esp, 00000024h
  loc_00E14FA0: push esi
  loc_00E14FA1: call [ecx+00000368h]
  loc_00E14FA7: lea edx, var_34
  loc_00E14FAA: push eax
  loc_00E14FAB: push edx
  loc_00E14FAC: call edi
  loc_00E14FAE: push 006DCBD8h
  loc_00E14FB3: mov var_D4, eax
  loc_00E14FB9: mov eax, [esi]
  loc_00E14FBB: push 00000000h
  loc_00E14FBD: push 00000018h
  loc_00E14FBF: push esi
  loc_00E14FC0: call [eax+00000490h]
  loc_00E14FC6: lea ecx, var_24
  loc_00E14FC9: push eax
  loc_00E14FCA: push ecx
  loc_00E14FCB: call edi
  loc_00E14FCD: lea edx, var_44
  loc_00E14FD0: push eax
  loc_00E14FD1: push edx
  loc_00E14FD2: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E14FD8: add esp, 00000010h
  loc_00E14FDB: push eax
  loc_00E14FDC: call [00401120h] ; __vbaCastObjVar
  loc_00E14FE2: push eax
  loc_00E14FE3: lea eax, var_28
  loc_00E14FE6: push eax
  loc_00E14FE7: call edi
  loc_00E14FE9: mov ebx, eax
  loc_00E14FEB: lea edx, var_2C
  loc_00E14FEE: push edx
  loc_00E14FEF: push ebx
  loc_00E14FF0: mov ecx, [ebx]
  loc_00E14FF2: call [ecx+00000054h]
  loc_00E14FF5: test eax, eax
  loc_00E14FF7: fnclex
  loc_00E14FF9: jge 00E1500Ah
  loc_00E14FFB: push 00000054h
  loc_00E14FFD: push 006DCBE8h
  loc_00E15002: push ebx
  loc_00E15003: push eax
  loc_00E15004: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1500A: lea ebx, var_30
  loc_00E1500D: mov eax, var_2C
  loc_00E15010: push ebx
  loc_00E15011: mov ecx, 00000002h
  loc_00E15016: sub esp, 00000010h
  loc_00E15019: mov var_84, ecx
  loc_00E1501F: mov ebx, esp
  loc_00E15021: mov var_7C, 0000000Ch
  loc_00E15028: mov edx, [eax]
  loc_00E1502A: push eax
  loc_00E1502B: mov [ebx], ecx
  loc_00E1502D: mov ecx, var_80
  loc_00E15030: mov var_C4, eax
  loc_00E15036: mov [ebx+00000004h], ecx
  loc_00E15039: mov ecx, var_7C
  loc_00E1503C: mov [ebx+00000008h], ecx
  loc_00E1503F: mov ecx, var_78
  loc_00E15042: mov [ebx+0000000Ch], ecx
  loc_00E15045: call [edx+00000028h]
  loc_00E15048: test eax, eax
  loc_00E1504A: fnclex
  loc_00E1504C: jge 00E15063h
  loc_00E1504E: mov edx, var_C4
  loc_00E15054: push 00000028h
  loc_00E15056: push 006E09E8h
  loc_00E1505B: push edx
  loc_00E1505C: push eax
  loc_00E1505D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15063: mov eax, var_30
  loc_00E15066: lea edx, var_54
  loc_00E15069: push edx
  loc_00E1506A: push eax
  loc_00E1506B: mov ecx, [eax]
  loc_00E1506D: mov ebx, eax
  loc_00E1506F: call [ecx+00000034h]
  loc_00E15072: test eax, eax
  loc_00E15074: fnclex
  loc_00E15076: jge 00E15087h
  loc_00E15078: push 00000034h
  loc_00E1507A: push 006E09F8h
  loc_00E1507F: push ebx
  loc_00E15080: push eax
  loc_00E15081: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15087: mov eax, var_D4
  loc_00E1508D: lea ecx, var_54
  loc_00E15090: push ecx
  loc_00E15091: mov ebx, [eax]
  loc_00E15093: call [00401034h] ; __vbaStrVarMove
  loc_00E15099: mov edx, eax
  loc_00E1509B: lea ecx, var_18
  loc_00E1509E: call [00401228h] ; __vbaStrMove
  loc_00E150A4: mov edx, ebx
  loc_00E150A6: mov ebx, var_D4
  loc_00E150AC: push eax
  loc_00E150AD: push ebx
  loc_00E150AE: call [edx+000000A4h]
  loc_00E150B4: test eax, eax
  loc_00E150B6: fnclex
  loc_00E150B8: jge 00E150CCh
  loc_00E150BA: push 000000A4h
  loc_00E150BF: push 006DCB70h
  loc_00E150C4: push ebx
  loc_00E150C5: push eax
  loc_00E150C6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E150CC: lea ecx, var_18
  loc_00E150CF: call [00401258h] ; __vbaFreeStr
  loc_00E150D5: lea eax, var_34
  loc_00E150D8: lea ecx, var_30
  loc_00E150DB: push eax
  loc_00E150DC: lea edx, var_2C
  loc_00E150DF: push ecx
  loc_00E150E0: lea eax, var_28
  loc_00E150E3: push edx
  loc_00E150E4: lea ecx, var_24
  loc_00E150E7: push eax
  loc_00E150E8: push ecx
  loc_00E150E9: push 00000005h
  loc_00E150EB: call [00401048h] ; __vbaFreeObjList
  loc_00E150F1: lea edx, var_54
  loc_00E150F4: lea eax, var_44
  loc_00E150F7: push edx
  loc_00E150F8: push eax
  loc_00E150F9: push 00000002h
  loc_00E150FB: call [00401038h] ; __vbaFreeVarList
  loc_00E15101: mov ecx, [esi]
  loc_00E15103: add esp, 00000024h
  loc_00E15106: push esi
  loc_00E15107: call [ecx+00000364h]
  loc_00E1510D: lea edx, var_34
  loc_00E15110: push eax
  loc_00E15111: push edx
  loc_00E15112: call edi
  loc_00E15114: push 006DCBD8h
  loc_00E15119: mov var_D4, eax
  loc_00E1511F: mov eax, [esi]
  loc_00E15121: push 00000000h
  loc_00E15123: push 00000018h
  loc_00E15125: push esi
  loc_00E15126: call [eax+00000490h]
  loc_00E1512C: lea ecx, var_24
  loc_00E1512F: push eax
  loc_00E15130: push ecx
  loc_00E15131: call edi
  loc_00E15133: lea edx, var_44
  loc_00E15136: push eax
  loc_00E15137: push edx
  loc_00E15138: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1513E: add esp, 00000010h
  loc_00E15141: push eax
  loc_00E15142: call [00401120h] ; __vbaCastObjVar
  loc_00E15148: push eax
  loc_00E15149: lea eax, var_28
  loc_00E1514C: push eax
  loc_00E1514D: call edi
  loc_00E1514F: mov ebx, eax
  loc_00E15151: lea edx, var_2C
  loc_00E15154: push edx
  loc_00E15155: push ebx
  loc_00E15156: mov ecx, [ebx]
  loc_00E15158: call [ecx+00000054h]
  loc_00E1515B: test eax, eax
  loc_00E1515D: fnclex
  loc_00E1515F: jge 00E15170h
  loc_00E15161: push 00000054h
  loc_00E15163: push 006DCBE8h
  loc_00E15168: push ebx
  loc_00E15169: push eax
  loc_00E1516A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15170: lea ebx, var_30
  loc_00E15173: mov eax, var_2C
  loc_00E15176: push ebx
  loc_00E15177: mov ecx, 00000002h
  loc_00E1517C: sub esp, 00000010h
  loc_00E1517F: mov var_84, ecx
  loc_00E15185: mov ebx, esp
  loc_00E15187: mov var_7C, 0000000Dh
  loc_00E1518E: mov edx, [eax]
  loc_00E15190: push eax
  loc_00E15191: mov [ebx], ecx
  loc_00E15193: mov ecx, var_80
  loc_00E15196: mov var_C4, eax
  loc_00E1519C: mov [ebx+00000004h], ecx
  loc_00E1519F: mov ecx, var_7C
  loc_00E151A2: mov [ebx+00000008h], ecx
  loc_00E151A5: mov ecx, var_78
  loc_00E151A8: mov [ebx+0000000Ch], ecx
  loc_00E151AB: call [edx+00000028h]
  loc_00E151AE: test eax, eax
  loc_00E151B0: fnclex
  loc_00E151B2: jge 00E151C9h
  loc_00E151B4: mov edx, var_C4
  loc_00E151BA: push 00000028h
  loc_00E151BC: push 006E09E8h
  loc_00E151C1: push edx
  loc_00E151C2: push eax
  loc_00E151C3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E151C9: mov eax, var_30
  loc_00E151CC: lea edx, var_54
  loc_00E151CF: push edx
  loc_00E151D0: push eax
  loc_00E151D1: mov ecx, [eax]
  loc_00E151D3: mov ebx, eax
  loc_00E151D5: call [ecx+00000034h]
  loc_00E151D8: test eax, eax
  loc_00E151DA: fnclex
  loc_00E151DC: jge 00E151EDh
  loc_00E151DE: push 00000034h
  loc_00E151E0: push 006E09F8h
  loc_00E151E5: push ebx
  loc_00E151E6: push eax
  loc_00E151E7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E151ED: mov eax, var_D4
  loc_00E151F3: lea ecx, var_54
  loc_00E151F6: push ecx
  loc_00E151F7: mov ebx, [eax]
  loc_00E151F9: call [00401034h] ; __vbaStrVarMove
  loc_00E151FF: mov edx, eax
  loc_00E15201: lea ecx, var_18
  loc_00E15204: call [00401228h] ; __vbaStrMove
  loc_00E1520A: mov edx, ebx
  loc_00E1520C: mov ebx, var_D4
  loc_00E15212: push eax
  loc_00E15213: push ebx
  loc_00E15214: call [edx+000000A4h]
  loc_00E1521A: test eax, eax
  loc_00E1521C: fnclex
  loc_00E1521E: jge 00E15232h
  loc_00E15220: push 000000A4h
  loc_00E15225: push 006DCB70h
  loc_00E1522A: push ebx
  loc_00E1522B: push eax
  loc_00E1522C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15232: lea ecx, var_18
  loc_00E15235: call [00401258h] ; __vbaFreeStr
  loc_00E1523B: lea eax, var_34
  loc_00E1523E: lea ecx, var_30
  loc_00E15241: push eax
  loc_00E15242: lea edx, var_2C
  loc_00E15245: push ecx
  loc_00E15246: lea eax, var_28
  loc_00E15249: push edx
  loc_00E1524A: lea ecx, var_24
  loc_00E1524D: push eax
  loc_00E1524E: push ecx
  loc_00E1524F: push 00000005h
  loc_00E15251: call [00401048h] ; __vbaFreeObjList
  loc_00E15257: lea edx, var_54
  loc_00E1525A: lea eax, var_44
  loc_00E1525D: push edx
  loc_00E1525E: push eax
  loc_00E1525F: push 00000002h
  loc_00E15261: call [00401038h] ; __vbaFreeVarList
  loc_00E15267: mov ecx, [esi]
  loc_00E15269: add esp, 00000024h
  loc_00E1526C: push esi
  loc_00E1526D: call [ecx+00000360h]
  loc_00E15273: lea edx, var_34
  loc_00E15276: push eax
  loc_00E15277: push edx
  loc_00E15278: call edi
  loc_00E1527A: push 006DCBD8h
  loc_00E1527F: mov var_D4, eax
  loc_00E15285: mov eax, [esi]
  loc_00E15287: push 00000000h
  loc_00E15289: push 00000018h
  loc_00E1528B: push esi
  loc_00E1528C: call [eax+00000490h]
  loc_00E15292: lea ecx, var_24
  loc_00E15295: push eax
  loc_00E15296: push ecx
  loc_00E15297: call edi
  loc_00E15299: lea edx, var_44
  loc_00E1529C: push eax
  loc_00E1529D: push edx
  loc_00E1529E: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E152A4: add esp, 00000010h
  loc_00E152A7: push eax
  loc_00E152A8: call [00401120h] ; __vbaCastObjVar
  loc_00E152AE: push eax
  loc_00E152AF: lea eax, var_28
  loc_00E152B2: push eax
  loc_00E152B3: call edi
  loc_00E152B5: mov ebx, eax
  loc_00E152B7: lea edx, var_2C
  loc_00E152BA: push edx
  loc_00E152BB: push ebx
  loc_00E152BC: mov ecx, [ebx]
  loc_00E152BE: call [ecx+00000054h]
  loc_00E152C1: test eax, eax
  loc_00E152C3: fnclex
  loc_00E152C5: jge 00E152D6h
  loc_00E152C7: push 00000054h
  loc_00E152C9: push 006DCBE8h
  loc_00E152CE: push ebx
  loc_00E152CF: push eax
  loc_00E152D0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E152D6: lea ebx, var_30
  loc_00E152D9: mov eax, var_2C
  loc_00E152DC: push ebx
  loc_00E152DD: mov ecx, 00000002h
  loc_00E152E2: sub esp, 00000010h
  loc_00E152E5: mov var_84, ecx
  loc_00E152EB: mov ebx, esp
  loc_00E152ED: mov var_7C, 0000000Eh
  loc_00E152F4: mov edx, [eax]
  loc_00E152F6: push eax
  loc_00E152F7: mov [ebx], ecx
  loc_00E152F9: mov ecx, var_80
  loc_00E152FC: mov var_C4, eax
  loc_00E15302: mov [ebx+00000004h], ecx
  loc_00E15305: mov ecx, var_7C
  loc_00E15308: mov [ebx+00000008h], ecx
  loc_00E1530B: mov ecx, var_78
  loc_00E1530E: mov [ebx+0000000Ch], ecx
  loc_00E15311: call [edx+00000028h]
  loc_00E15314: test eax, eax
  loc_00E15316: fnclex
  loc_00E15318: jge 00E1532Fh
  loc_00E1531A: mov edx, var_C4
  loc_00E15320: push 00000028h
  loc_00E15322: push 006E09E8h
  loc_00E15327: push edx
  loc_00E15328: push eax
  loc_00E15329: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1532F: mov eax, var_30
  loc_00E15332: lea edx, var_54
  loc_00E15335: push edx
  loc_00E15336: push eax
  loc_00E15337: mov ecx, [eax]
  loc_00E15339: mov ebx, eax
  loc_00E1533B: call [ecx+00000034h]
  loc_00E1533E: test eax, eax
  loc_00E15340: fnclex
  loc_00E15342: jge 00E15353h
  loc_00E15344: push 00000034h
  loc_00E15346: push 006E09F8h
  loc_00E1534B: push ebx
  loc_00E1534C: push eax
  loc_00E1534D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15353: mov eax, var_D4
  loc_00E15359: lea ecx, var_54
  loc_00E1535C: push ecx
  loc_00E1535D: mov ebx, [eax]
  loc_00E1535F: call [00401034h] ; __vbaStrVarMove
  loc_00E15365: mov edx, eax
  loc_00E15367: lea ecx, var_18
  loc_00E1536A: call [00401228h] ; __vbaStrMove
  loc_00E15370: mov edx, ebx
  loc_00E15372: mov ebx, var_D4
  loc_00E15378: push eax
  loc_00E15379: push ebx
  loc_00E1537A: call [edx+000000A4h]
  loc_00E15380: test eax, eax
  loc_00E15382: fnclex
  loc_00E15384: jge 00E15398h
  loc_00E15386: push 000000A4h
  loc_00E1538B: push 006DCB70h
  loc_00E15390: push ebx
  loc_00E15391: push eax
  loc_00E15392: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15398: lea ecx, var_18
  loc_00E1539B: call [00401258h] ; __vbaFreeStr
  loc_00E153A1: lea eax, var_34
  loc_00E153A4: lea ecx, var_30
  loc_00E153A7: push eax
  loc_00E153A8: lea edx, var_2C
  loc_00E153AB: push ecx
  loc_00E153AC: lea eax, var_28
  loc_00E153AF: push edx
  loc_00E153B0: lea ecx, var_24
  loc_00E153B3: push eax
  loc_00E153B4: push ecx
  loc_00E153B5: push 00000005h
  loc_00E153B7: call [00401048h] ; __vbaFreeObjList
  loc_00E153BD: lea edx, var_54
  loc_00E153C0: lea eax, var_44
  loc_00E153C3: push edx
  loc_00E153C4: push eax
  loc_00E153C5: push 00000002h
  loc_00E153C7: call [00401038h] ; __vbaFreeVarList
  loc_00E153CD: mov ecx, [esi]
  loc_00E153CF: add esp, 00000024h
  loc_00E153D2: push esi
  loc_00E153D3: call [ecx+0000033Ch]
  loc_00E153D9: lea edx, var_34
  loc_00E153DC: push eax
  loc_00E153DD: push edx
  loc_00E153DE: call edi
  loc_00E153E0: push 006DCBD8h
  loc_00E153E5: mov var_D4, eax
  loc_00E153EB: mov eax, [esi]
  loc_00E153ED: push 00000000h
  loc_00E153EF: push 00000018h
  loc_00E153F1: push esi
  loc_00E153F2: call [eax+00000490h]
  loc_00E153F8: lea ecx, var_24
  loc_00E153FB: push eax
  loc_00E153FC: push ecx
  loc_00E153FD: call edi
  loc_00E153FF: lea edx, var_44
  loc_00E15402: push eax
  loc_00E15403: push edx
  loc_00E15404: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1540A: add esp, 00000010h
  loc_00E1540D: push eax
  loc_00E1540E: call [00401120h] ; __vbaCastObjVar
  loc_00E15414: push eax
  loc_00E15415: lea eax, var_28
  loc_00E15418: push eax
  loc_00E15419: call edi
  loc_00E1541B: mov ebx, eax
  loc_00E1541D: lea edx, var_2C
  loc_00E15420: push edx
  loc_00E15421: push ebx
  loc_00E15422: mov ecx, [ebx]
  loc_00E15424: call [ecx+00000054h]
  loc_00E15427: test eax, eax
  loc_00E15429: fnclex
  loc_00E1542B: jge 00E1543Ch
  loc_00E1542D: push 00000054h
  loc_00E1542F: push 006DCBE8h
  loc_00E15434: push ebx
  loc_00E15435: push eax
  loc_00E15436: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1543C: lea ebx, var_30
  loc_00E1543F: mov eax, var_2C
  loc_00E15442: push ebx
  loc_00E15443: mov ecx, 00000002h
  loc_00E15448: sub esp, 00000010h
  loc_00E1544B: mov var_84, ecx
  loc_00E15451: mov ebx, esp
  loc_00E15453: mov var_7C, 0000000Fh
  loc_00E1545A: mov edx, [eax]
  loc_00E1545C: push eax
  loc_00E1545D: mov [ebx], ecx
  loc_00E1545F: mov ecx, var_80
  loc_00E15462: mov var_C4, eax
  loc_00E15468: mov [ebx+00000004h], ecx
  loc_00E1546B: mov ecx, var_7C
  loc_00E1546E: mov [ebx+00000008h], ecx
  loc_00E15471: mov ecx, var_78
  loc_00E15474: mov [ebx+0000000Ch], ecx
  loc_00E15477: call [edx+00000028h]
  loc_00E1547A: test eax, eax
  loc_00E1547C: fnclex
  loc_00E1547E: jge 00E15495h
  loc_00E15480: mov edx, var_C4
  loc_00E15486: push 00000028h
  loc_00E15488: push 006E09E8h
  loc_00E1548D: push edx
  loc_00E1548E: push eax
  loc_00E1548F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15495: mov eax, var_30
  loc_00E15498: lea edx, var_54
  loc_00E1549B: push edx
  loc_00E1549C: push eax
  loc_00E1549D: mov ecx, [eax]
  loc_00E1549F: mov ebx, eax
  loc_00E154A1: call [ecx+00000034h]
  loc_00E154A4: test eax, eax
  loc_00E154A6: fnclex
  loc_00E154A8: jge 00E154B9h
  loc_00E154AA: push 00000034h
  loc_00E154AC: push 006E09F8h
  loc_00E154B1: push ebx
  loc_00E154B2: push eax
  loc_00E154B3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E154B9: mov eax, var_D4
  loc_00E154BF: lea ecx, var_54
  loc_00E154C2: push ecx
  loc_00E154C3: mov ebx, [eax]
  loc_00E154C5: call [00401034h] ; __vbaStrVarMove
  loc_00E154CB: mov edx, eax
  loc_00E154CD: lea ecx, var_18
  loc_00E154D0: call [00401228h] ; __vbaStrMove
  loc_00E154D6: mov edx, ebx
  loc_00E154D8: mov ebx, var_D4
  loc_00E154DE: push eax
  loc_00E154DF: push ebx
  loc_00E154E0: call [edx+000000A4h]
  loc_00E154E6: test eax, eax
  loc_00E154E8: fnclex
  loc_00E154EA: jge 00E154FEh
  loc_00E154EC: push 000000A4h
  loc_00E154F1: push 006DCB70h
  loc_00E154F6: push ebx
  loc_00E154F7: push eax
  loc_00E154F8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E154FE: lea ecx, var_18
  loc_00E15501: call [00401258h] ; __vbaFreeStr
  loc_00E15507: lea eax, var_34
  loc_00E1550A: lea ecx, var_30
  loc_00E1550D: push eax
  loc_00E1550E: lea edx, var_2C
  loc_00E15511: push ecx
  loc_00E15512: lea eax, var_28
  loc_00E15515: push edx
  loc_00E15516: lea ecx, var_24
  loc_00E15519: push eax
  loc_00E1551A: push ecx
  loc_00E1551B: push 00000005h
  loc_00E1551D: call [00401048h] ; __vbaFreeObjList
  loc_00E15523: lea edx, var_54
  loc_00E15526: lea eax, var_44
  loc_00E15529: push edx
  loc_00E1552A: push eax
  loc_00E1552B: push 00000002h
  loc_00E1552D: call [00401038h] ; __vbaFreeVarList
  loc_00E15533: mov ecx, [esi]
  loc_00E15535: add esp, 00000024h
  loc_00E15538: push esi
  loc_00E15539: call [ecx+00000340h]
  loc_00E1553F: lea edx, var_34
  loc_00E15542: push eax
  loc_00E15543: push edx
  loc_00E15544: call edi
  loc_00E15546: push 006DCBD8h
  loc_00E1554B: mov var_D4, eax
  loc_00E15551: mov eax, [esi]
  loc_00E15553: push 00000000h
  loc_00E15555: push 00000018h
  loc_00E15557: push esi
  loc_00E15558: call [eax+00000490h]
  loc_00E1555E: lea ecx, var_24
  loc_00E15561: push eax
  loc_00E15562: push ecx
  loc_00E15563: call edi
  loc_00E15565: lea edx, var_44
  loc_00E15568: push eax
  loc_00E15569: push edx
  loc_00E1556A: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E15570: add esp, 00000010h
  loc_00E15573: push eax
  loc_00E15574: call [00401120h] ; __vbaCastObjVar
  loc_00E1557A: push eax
  loc_00E1557B: lea eax, var_28
  loc_00E1557E: push eax
  loc_00E1557F: call edi
  loc_00E15581: mov ebx, eax
  loc_00E15583: lea edx, var_2C
  loc_00E15586: push edx
  loc_00E15587: push ebx
  loc_00E15588: mov ecx, [ebx]
  loc_00E1558A: call [ecx+00000054h]
  loc_00E1558D: test eax, eax
  loc_00E1558F: fnclex
  loc_00E15591: jge 00E155A2h
  loc_00E15593: push 00000054h
  loc_00E15595: push 006DCBE8h
  loc_00E1559A: push ebx
  loc_00E1559B: push eax
  loc_00E1559C: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E155A2: lea ebx, var_30
  loc_00E155A5: mov eax, var_2C
  loc_00E155A8: push ebx
  loc_00E155A9: mov ecx, 00000002h
  loc_00E155AE: sub esp, 00000010h
  loc_00E155B1: mov var_84, ecx
  loc_00E155B7: mov ebx, esp
  loc_00E155B9: mov var_7C, 00000010h
  loc_00E155C0: mov edx, [eax]
  loc_00E155C2: push eax
  loc_00E155C3: mov [ebx], ecx
  loc_00E155C5: mov ecx, var_80
  loc_00E155C8: mov var_C4, eax
  loc_00E155CE: mov [ebx+00000004h], ecx
  loc_00E155D1: mov ecx, var_7C
  loc_00E155D4: mov [ebx+00000008h], ecx
  loc_00E155D7: mov ecx, var_78
  loc_00E155DA: mov [ebx+0000000Ch], ecx
  loc_00E155DD: call [edx+00000028h]
  loc_00E155E0: test eax, eax
  loc_00E155E2: fnclex
  loc_00E155E4: jge 00E155FBh
  loc_00E155E6: mov edx, var_C4
  loc_00E155EC: push 00000028h
  loc_00E155EE: push 006E09E8h
  loc_00E155F3: push edx
  loc_00E155F4: push eax
  loc_00E155F5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E155FB: mov eax, var_30
  loc_00E155FE: lea edx, var_54
  loc_00E15601: push edx
  loc_00E15602: push eax
  loc_00E15603: mov ecx, [eax]
  loc_00E15605: mov ebx, eax
  loc_00E15607: call [ecx+00000034h]
  loc_00E1560A: test eax, eax
  loc_00E1560C: fnclex
  loc_00E1560E: jge 00E1561Fh
  loc_00E15610: push 00000034h
  loc_00E15612: push 006E09F8h
  loc_00E15617: push ebx
  loc_00E15618: push eax
  loc_00E15619: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1561F: mov eax, var_D4
  loc_00E15625: lea ecx, var_54
  loc_00E15628: push ecx
  loc_00E15629: mov ebx, [eax]
  loc_00E1562B: call [00401034h] ; __vbaStrVarMove
  loc_00E15631: mov edx, eax
  loc_00E15633: lea ecx, var_18
  loc_00E15636: call [00401228h] ; __vbaStrMove
  loc_00E1563C: mov edx, ebx
  loc_00E1563E: mov ebx, var_D4
  loc_00E15644: push eax
  loc_00E15645: push ebx
  loc_00E15646: call [edx+000000A4h]
  loc_00E1564C: test eax, eax
  loc_00E1564E: fnclex
  loc_00E15650: jge 00E15664h
  loc_00E15652: push 000000A4h
  loc_00E15657: push 006DCB70h
  loc_00E1565C: push ebx
  loc_00E1565D: push eax
  loc_00E1565E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15664: lea ecx, var_18
  loc_00E15667: call [00401258h] ; __vbaFreeStr
  loc_00E1566D: lea eax, var_34
  loc_00E15670: lea ecx, var_30
  loc_00E15673: push eax
  loc_00E15674: lea edx, var_2C
  loc_00E15677: push ecx
  loc_00E15678: lea eax, var_28
  loc_00E1567B: push edx
  loc_00E1567C: lea ecx, var_24
  loc_00E1567F: push eax
  loc_00E15680: push ecx
  loc_00E15681: push 00000005h
  loc_00E15683: call [00401048h] ; __vbaFreeObjList
  loc_00E15689: lea edx, var_54
  loc_00E1568C: lea eax, var_44
  loc_00E1568F: push edx
  loc_00E15690: push eax
  loc_00E15691: push 00000002h
  loc_00E15693: call [00401038h] ; __vbaFreeVarList
  loc_00E15699: mov ecx, [esi]
  loc_00E1569B: add esp, 00000024h
  loc_00E1569E: push esi
  loc_00E1569F: call [ecx+00000344h]
  loc_00E156A5: lea edx, var_34
  loc_00E156A8: push eax
  loc_00E156A9: push edx
  loc_00E156AA: call edi
  loc_00E156AC: push 006DCBD8h
  loc_00E156B1: mov var_D4, eax
  loc_00E156B7: mov eax, [esi]
  loc_00E156B9: push 00000000h
  loc_00E156BB: push 00000018h
  loc_00E156BD: push esi
  loc_00E156BE: call [eax+00000490h]
  loc_00E156C4: lea ecx, var_24
  loc_00E156C7: push eax
  loc_00E156C8: push ecx
  loc_00E156C9: call edi
  loc_00E156CB: lea edx, var_44
  loc_00E156CE: push eax
  loc_00E156CF: push edx
  loc_00E156D0: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E156D6: add esp, 00000010h
  loc_00E156D9: push eax
  loc_00E156DA: call [00401120h] ; __vbaCastObjVar
  loc_00E156E0: push eax
  loc_00E156E1: lea eax, var_28
  loc_00E156E4: push eax
  loc_00E156E5: call edi
  loc_00E156E7: mov ebx, eax
  loc_00E156E9: lea edx, var_2C
  loc_00E156EC: push edx
  loc_00E156ED: push ebx
  loc_00E156EE: mov ecx, [ebx]
  loc_00E156F0: call [ecx+00000054h]
  loc_00E156F3: test eax, eax
  loc_00E156F5: fnclex
  loc_00E156F7: jge 00E15708h
  loc_00E156F9: push 00000054h
  loc_00E156FB: push 006DCBE8h
  loc_00E15700: push ebx
  loc_00E15701: push eax
  loc_00E15702: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15708: lea ebx, var_30
  loc_00E1570B: mov eax, var_2C
  loc_00E1570E: push ebx
  loc_00E1570F: mov ecx, 00000002h
  loc_00E15714: sub esp, 00000010h
  loc_00E15717: mov var_84, ecx
  loc_00E1571D: mov ebx, esp
  loc_00E1571F: mov var_7C, 00000011h
  loc_00E15726: mov edx, [eax]
  loc_00E15728: push eax
  loc_00E15729: mov [ebx], ecx
  loc_00E1572B: mov ecx, var_80
  loc_00E1572E: mov var_C4, eax
  loc_00E15734: mov [ebx+00000004h], ecx
  loc_00E15737: mov ecx, var_7C
  loc_00E1573A: mov [ebx+00000008h], ecx
  loc_00E1573D: mov ecx, var_78
  loc_00E15740: mov [ebx+0000000Ch], ecx
  loc_00E15743: call [edx+00000028h]
  loc_00E15746: test eax, eax
  loc_00E15748: fnclex
  loc_00E1574A: jge 00E15761h
  loc_00E1574C: mov edx, var_C4
  loc_00E15752: push 00000028h
  loc_00E15754: push 006E09E8h
  loc_00E15759: push edx
  loc_00E1575A: push eax
  loc_00E1575B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15761: mov eax, var_30
  loc_00E15764: lea edx, var_54
  loc_00E15767: push edx
  loc_00E15768: push eax
  loc_00E15769: mov ecx, [eax]
  loc_00E1576B: mov ebx, eax
  loc_00E1576D: call [ecx+00000034h]
  loc_00E15770: test eax, eax
  loc_00E15772: fnclex
  loc_00E15774: jge 00E15785h
  loc_00E15776: push 00000034h
  loc_00E15778: push 006E09F8h
  loc_00E1577D: push ebx
  loc_00E1577E: push eax
  loc_00E1577F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15785: mov eax, var_D4
  loc_00E1578B: lea ecx, var_54
  loc_00E1578E: push ecx
  loc_00E1578F: mov ebx, [eax]
  loc_00E15791: call [00401034h] ; __vbaStrVarMove
  loc_00E15797: mov edx, eax
  loc_00E15799: lea ecx, var_18
  loc_00E1579C: call [00401228h] ; __vbaStrMove
  loc_00E157A2: mov edx, ebx
  loc_00E157A4: mov ebx, var_D4
  loc_00E157AA: push eax
  loc_00E157AB: push ebx
  loc_00E157AC: call [edx+000000A4h]
  loc_00E157B2: test eax, eax
  loc_00E157B4: fnclex
  loc_00E157B6: jge 00E157CAh
  loc_00E157B8: push 000000A4h
  loc_00E157BD: push 006DCB70h
  loc_00E157C2: push ebx
  loc_00E157C3: push eax
  loc_00E157C4: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E157CA: lea ecx, var_18
  loc_00E157CD: call [00401258h] ; __vbaFreeStr
  loc_00E157D3: lea eax, var_34
  loc_00E157D6: lea ecx, var_30
  loc_00E157D9: push eax
  loc_00E157DA: lea edx, var_2C
  loc_00E157DD: push ecx
  loc_00E157DE: lea eax, var_28
  loc_00E157E1: push edx
  loc_00E157E2: lea ecx, var_24
  loc_00E157E5: push eax
  loc_00E157E6: push ecx
  loc_00E157E7: push 00000005h
  loc_00E157E9: call [00401048h] ; __vbaFreeObjList
  loc_00E157EF: lea edx, var_54
  loc_00E157F2: lea eax, var_44
  loc_00E157F5: push edx
  loc_00E157F6: push eax
  loc_00E157F7: push 00000002h
  loc_00E157F9: call [00401038h] ; __vbaFreeVarList
  loc_00E157FF: mov ecx, [esi]
  loc_00E15801: add esp, 00000024h
  loc_00E15804: push esi
  loc_00E15805: call [ecx+00000348h]
  loc_00E1580B: lea edx, var_34
  loc_00E1580E: push eax
  loc_00E1580F: push edx
  loc_00E15810: call edi
  loc_00E15812: push 006DCBD8h
  loc_00E15817: mov var_D4, eax
  loc_00E1581D: mov eax, [esi]
  loc_00E1581F: push 00000000h
  loc_00E15821: push 00000018h
  loc_00E15823: push esi
  loc_00E15824: call [eax+00000490h]
  loc_00E1582A: lea ecx, var_24
  loc_00E1582D: push eax
  loc_00E1582E: push ecx
  loc_00E1582F: call edi
  loc_00E15831: lea edx, var_44
  loc_00E15834: push eax
  loc_00E15835: push edx
  loc_00E15836: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E1583C: add esp, 00000010h
  loc_00E1583F: push eax
  loc_00E15840: call [00401120h] ; __vbaCastObjVar
  loc_00E15846: push eax
  loc_00E15847: lea eax, var_28
  loc_00E1584A: push eax
  loc_00E1584B: call edi
  loc_00E1584D: mov ebx, eax
  loc_00E1584F: lea edx, var_2C
  loc_00E15852: push edx
  loc_00E15853: push ebx
  loc_00E15854: mov ecx, [ebx]
  loc_00E15856: call [ecx+00000054h]
  loc_00E15859: test eax, eax
  loc_00E1585B: fnclex
  loc_00E1585D: jge 00E1586Eh
  loc_00E1585F: push 00000054h
  loc_00E15861: push 006DCBE8h
  loc_00E15866: push ebx
  loc_00E15867: push eax
  loc_00E15868: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1586E: lea ebx, var_30
  loc_00E15871: mov eax, var_2C
  loc_00E15874: push ebx
  loc_00E15875: mov ecx, 00000002h
  loc_00E1587A: sub esp, 00000010h
  loc_00E1587D: mov var_84, ecx
  loc_00E15883: mov ebx, esp
  loc_00E15885: mov var_7C, 00000012h
  loc_00E1588C: mov edx, [eax]
  loc_00E1588E: push eax
  loc_00E1588F: mov [ebx], ecx
  loc_00E15891: mov ecx, var_80
  loc_00E15894: mov var_C4, eax
  loc_00E1589A: mov [ebx+00000004h], ecx
  loc_00E1589D: mov ecx, var_7C
  loc_00E158A0: mov [ebx+00000008h], ecx
  loc_00E158A3: mov ecx, var_78
  loc_00E158A6: mov [ebx+0000000Ch], ecx
  loc_00E158A9: call [edx+00000028h]
  loc_00E158AC: test eax, eax
  loc_00E158AE: fnclex
  loc_00E158B0: jge 00E158C7h
  loc_00E158B2: mov edx, var_C4
  loc_00E158B8: push 00000028h
  loc_00E158BA: push 006E09E8h
  loc_00E158BF: push edx
  loc_00E158C0: push eax
  loc_00E158C1: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E158C7: mov eax, var_30
  loc_00E158CA: lea edx, var_54
  loc_00E158CD: push edx
  loc_00E158CE: push eax
  loc_00E158CF: mov ecx, [eax]
  loc_00E158D1: mov ebx, eax
  loc_00E158D3: call [ecx+00000034h]
  loc_00E158D6: test eax, eax
  loc_00E158D8: fnclex
  loc_00E158DA: jge 00E158EBh
  loc_00E158DC: push 00000034h
  loc_00E158DE: push 006E09F8h
  loc_00E158E3: push ebx
  loc_00E158E4: push eax
  loc_00E158E5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E158EB: mov eax, var_D4
  loc_00E158F1: lea ecx, var_54
  loc_00E158F4: push ecx
  loc_00E158F5: mov ebx, [eax]
  loc_00E158F7: call [00401034h] ; __vbaStrVarMove
  loc_00E158FD: mov edx, eax
  loc_00E158FF: lea ecx, var_18
  loc_00E15902: call [00401228h] ; __vbaStrMove
  loc_00E15908: mov edx, ebx
  loc_00E1590A: mov ebx, var_D4
  loc_00E15910: push eax
  loc_00E15911: push ebx
  loc_00E15912: call [edx+000000A4h]
  loc_00E15918: test eax, eax
  loc_00E1591A: fnclex
  loc_00E1591C: jge 00E15930h
  loc_00E1591E: push 000000A4h
  loc_00E15923: push 006DCB70h
  loc_00E15928: push ebx
  loc_00E15929: push eax
  loc_00E1592A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E15930: lea ecx, var_18
  loc_00E15933: call [00401258h] ; __vbaFreeStr
  loc_00E15939: lea eax, var_34
  loc_00E1593C: lea ecx, var_30
  loc_00E1593F: push eax
  loc_00E15940: lea edx, var_2C
  loc_00E15943: push ecx
  loc_00E15944: lea eax, var_28
  loc_00E15947: push edx
  loc_00E15948: lea ecx, var_24
  loc_00E1594B: push eax
  loc_00E1594C: push ecx
  loc_00E1594D: push 00000005h
  loc_00E1594F: call [00401048h] ; __vbaFreeObjList
  loc_00E15955: lea edx, var_54
  loc_00E15958: lea eax, var_44
  loc_00E1595B: push edx
  loc_00E1595C: push eax
  loc_00E1595D: push 00000002h
  loc_00E1595F: call [00401038h] ; __vbaFreeVarList
  loc_00E15965: mov ecx, [esi]
  loc_00E15967: add esp, 00000024h
  loc_00E1596A: push esi
  loc_00E1596B: call [ecx+00000338h]
  loc_00E15971: lea edx, var_24
  loc_00E15974: push eax
  loc_00E15975: push edx
  loc_00E15976: call edi
  loc_00E15978: mov esi, eax
  loc_00E1597A: push FFFFFFFFh
  loc_00E1597C: push esi
  loc_00E1597D: mov eax, [esi]
  loc_00E1597F: call [eax+0000009Ch]
  loc_00E15985: test eax, eax
  loc_00E15987: fnclex
  loc_00E15989: jge 00E1599Dh
  loc_00E1598B: push 0000009Ch
  loc_00E15990: push 006DCAD0h
  loc_00E15995: push esi
  loc_00E15996: push eax
  loc_00E15997: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1599D: lea ecx, var_24
  loc_00E159A0: call [00401254h] ; __vbaFreeObj
  loc_00E159A6: mov var_4, 00000000h
  loc_00E159AD: push 00E15A01h
  loc_00E159B2: jmp 00E15A00h
  loc_00E159B4: lea ecx, var_20
  loc_00E159B7: lea edx, var_1C
  loc_00E159BA: push ecx
  loc_00E159BB: lea eax, var_18
  loc_00E159BE: push edx
  loc_00E159BF: push eax
  loc_00E159C0: push 00000003h
  loc_00E159C2: call [004011BCh] ; __vbaFreeStrList
  loc_00E159C8: lea ecx, var_34
  loc_00E159CB: lea edx, var_30
  loc_00E159CE: push ecx
  loc_00E159CF: lea eax, var_2C
  loc_00E159D2: push edx
  loc_00E159D3: lea ecx, var_28
  loc_00E159D6: push eax
  loc_00E159D7: lea edx, var_24
  loc_00E159DA: push ecx
  loc_00E159DB: push edx
  loc_00E159DC: push 00000005h
  loc_00E159DE: call [00401048h] ; __vbaFreeObjList
  loc_00E159E4: lea eax, var_74
  loc_00E159E7: lea ecx, var_64
  loc_00E159EA: push eax
  loc_00E159EB: lea edx, var_54
  loc_00E159EE: push ecx
  loc_00E159EF: lea eax, var_44
  loc_00E159F2: push edx
  loc_00E159F3: push eax
  loc_00E159F4: push 00000004h
  loc_00E159F6: call [00401038h] ; __vbaFreeVarList
  loc_00E159FC: add esp, 0000003Ch
  loc_00E159FF: ret
  loc_00E15A00: ret
  loc_00E15A01: mov eax, Me
  loc_00E15A04: push eax
  loc_00E15A05: mov ecx, [eax]
  loc_00E15A07: call [ecx+00000008h]
  loc_00E15A0A: mov eax, var_4
  loc_00E15A0D: mov ecx, var_14
  loc_00E15A10: pop edi
  loc_00E15A11: pop esi
  loc_00E15A12: mov fs:[00000000h], ecx
  loc_00E15A19: pop ebx
  loc_00E15A1A: mov esp, ebp
  loc_00E15A1C: pop ebp
  loc_00E15A1D: retn 0008h
End Sub

Private Sub txtcari_Click() 'E10000
  loc_00E10000: push ebp
  loc_00E10001: mov ebp, esp
  loc_00E10003: sub esp, 0000000Ch
  loc_00E10006: push 00402836h ; __vbaExceptHandler
  loc_00E1000B: mov eax, fs:[00000000h]
  loc_00E10011: push eax
  loc_00E10012: mov fs:[00000000h], esp
  loc_00E10019: sub esp, 00000014h
  loc_00E1001C: push ebx
  loc_00E1001D: push esi
  loc_00E1001E: push edi
  loc_00E1001F: mov var_C, esp
  loc_00E10022: mov var_8, 004022B8h
  loc_00E10029: mov esi, Me
  loc_00E1002C: mov eax, esi
  loc_00E1002E: and eax, 00000001h
  loc_00E10031: mov var_4, eax
  loc_00E10034: and esi, FFFFFFFEh
  loc_00E10037: push esi
  loc_00E10038: mov Me, esi
  loc_00E1003B: mov ecx, [esi]
  loc_00E1003D: call [ecx+00000004h]
  loc_00E10040: mov edx, [esi]
  loc_00E10042: xor edi, edi
  loc_00E10044: push esi
  loc_00E10045: mov var_18, edi
  loc_00E10048: call [edx+000003E0h]
  loc_00E1004E: push eax
  loc_00E1004F: lea eax, var_18
  loc_00E10052: push eax
  loc_00E10053: call [004010ACh] ; __vbaObjSet
  loc_00E10059: mov esi, eax
  loc_00E1005B: push edi
  loc_00E1005C: push esi
  loc_00E1005D: mov ecx, [esi]
  loc_00E1005F: call [ecx+0000008Ch]
  loc_00E10065: cmp eax, edi
  loc_00E10067: fnclex
  loc_00E10069: jge 00E1007Dh
  loc_00E1006B: push 0000008Ch
  loc_00E10070: push 006DCDA0h
  loc_00E10075: push esi
  loc_00E10076: push eax
  loc_00E10077: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E1007D: lea ecx, var_18
  loc_00E10080: call [00401254h] ; __vbaFreeObj
  loc_00E10086: mov var_4, edi
  loc_00E10089: push 00E1009Bh
  loc_00E1008E: jmp 00E1009Ah
  loc_00E10090: lea ecx, var_18
  loc_00E10093: call [00401254h] ; __vbaFreeObj
  loc_00E10099: ret
  loc_00E1009A: ret
  loc_00E1009B: mov eax, Me
  loc_00E1009E: push eax
  loc_00E1009F: mov edx, [eax]
  loc_00E100A1: call [edx+00000008h]
  loc_00E100A4: mov eax, var_4
  loc_00E100A7: mov ecx, var_14
  loc_00E100AA: pop edi
  loc_00E100AB: pop esi
  loc_00E100AC: mov fs:[00000000h], ecx
  loc_00E100B3: pop ebx
  loc_00E100B4: mov esp, ebp
  loc_00E100B6: pop ebp
  loc_00E100B7: retn 0004h
End Sub

Private Sub optnibu_Click() 'E0A5F0
  loc_00E0A5F0: push ebp
  loc_00E0A5F1: mov ebp, esp
  loc_00E0A5F3: sub esp, 0000000Ch
  loc_00E0A5F6: push 00402836h ; __vbaExceptHandler
  loc_00E0A5FB: mov eax, fs:[00000000h]
  loc_00E0A601: push eax
  loc_00E0A602: mov fs:[00000000h], esp
  loc_00E0A609: sub esp, 00000014h
  loc_00E0A60C: push ebx
  loc_00E0A60D: push esi
  loc_00E0A60E: push edi
  loc_00E0A60F: mov var_C, esp
  loc_00E0A612: mov var_8, 004021E0h
  loc_00E0A619: mov esi, Me
  loc_00E0A61C: mov eax, esi
  loc_00E0A61E: and eax, 00000001h
  loc_00E0A621: mov var_4, eax
  loc_00E0A624: and esi, FFFFFFFEh
  loc_00E0A627: push esi
  loc_00E0A628: mov Me, esi
  loc_00E0A62B: mov ecx, [esi]
  loc_00E0A62D: call [ecx+00000004h]
  loc_00E0A630: mov edx, [esi]
  loc_00E0A632: xor edi, edi
  loc_00E0A634: push esi
  loc_00E0A635: mov var_18, edi
  loc_00E0A638: call [edx+000003E0h]
  loc_00E0A63E: push eax
  loc_00E0A63F: lea eax, var_18
  loc_00E0A642: push eax
  loc_00E0A643: call [004010ACh] ; __vbaObjSet
  loc_00E0A649: mov esi, eax
  loc_00E0A64B: push edi
  loc_00E0A64C: push esi
  loc_00E0A64D: mov ecx, [esi]
  loc_00E0A64F: call [ecx+0000008Ch]
  loc_00E0A655: cmp eax, edi
  loc_00E0A657: fnclex
  loc_00E0A659: jge 00E0A66Dh
  loc_00E0A65B: push 0000008Ch
  loc_00E0A660: push 006DCDA0h
  loc_00E0A665: push esi
  loc_00E0A666: push eax
  loc_00E0A667: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A66D: lea ecx, var_18
  loc_00E0A670: call [00401254h] ; __vbaFreeObj
  loc_00E0A676: mov var_4, edi
  loc_00E0A679: push 00E0A68Bh
  loc_00E0A67E: jmp 00E0A68Ah
  loc_00E0A680: lea ecx, var_18
  loc_00E0A683: call [00401254h] ; __vbaFreeObj
  loc_00E0A689: ret
  loc_00E0A68A: ret
  loc_00E0A68B: mov eax, Me
  loc_00E0A68E: push eax
  loc_00E0A68F: mov edx, [eax]
  loc_00E0A691: call [edx+00000008h]
  loc_00E0A694: mov eax, var_4
  loc_00E0A697: mov ecx, var_14
  loc_00E0A69A: pop edi
  loc_00E0A69B: pop esi
  loc_00E0A69C: mov fs:[00000000h], ecx
  loc_00E0A6A3: pop ebx
  loc_00E0A6A4: mov esp, ebp
  loc_00E0A6A6: pop ebp
  loc_00E0A6A7: retn 0004h
End Sub

Private Sub pluss_UnknownEvent_9 'E0AE40
  loc_00E0AE40: push ebp
  loc_00E0AE41: mov ebp, esp
  loc_00E0AE43: sub esp, 0000000Ch
  loc_00E0AE46: push 00402836h ; __vbaExceptHandler
  loc_00E0AE4B: mov eax, fs:[00000000h]
  loc_00E0AE51: push eax
  loc_00E0AE52: mov fs:[00000000h], esp
  loc_00E0AE59: sub esp, 00000130h
  loc_00E0AE5F: push ebx
  loc_00E0AE60: push esi
  loc_00E0AE61: push edi
  loc_00E0AE62: mov var_C, esp
  loc_00E0AE65: mov var_8, 00402240h
  loc_00E0AE6C: mov esi, Me
  loc_00E0AE6F: mov eax, esi
  loc_00E0AE71: and eax, 00000001h
  loc_00E0AE74: mov var_4, eax
  loc_00E0AE77: and esi, FFFFFFFEh
  loc_00E0AE7A: push esi
  loc_00E0AE7B: mov Me, esi
  loc_00E0AE7E: mov ecx, [esi]
  loc_00E0AE80: call [ecx+00000004h]
  loc_00E0AE83: mov edx, [esi]
  loc_00E0AE85: xor edi, edi
  loc_00E0AE87: push esi
  loc_00E0AE88: mov var_24, edi
  loc_00E0AE8B: mov var_2C, edi
  loc_00E0AE8E: mov var_30, edi
  loc_00E0AE91: mov var_34, edi
  loc_00E0AE94: mov var_38, edi
  loc_00E0AE97: mov var_3C, edi
  loc_00E0AE9A: mov var_40, edi
  loc_00E0AE9D: mov var_44, edi
  loc_00E0AEA0: mov var_48, edi
  loc_00E0AEA3: mov var_58, edi
  loc_00E0AEA6: mov var_68, edi
  loc_00E0AEA9: mov var_78, edi
  loc_00E0AEAC: mov var_88, edi
  loc_00E0AEB2: mov var_98, edi
  loc_00E0AEB8: mov var_A8, edi
  loc_00E0AEBE: mov var_D8, edi
  loc_00E0AEC4: mov var_DC, edi
  loc_00E0AECA: call [edx+0000039Ch]
  loc_00E0AED0: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0AED6: push eax
  loc_00E0AED7: lea eax, var_38
  loc_00E0AEDA: push eax
  loc_00E0AEDB: call ebx
  loc_00E0AEDD: mov ecx, [eax]
  loc_00E0AEDF: lea edx, var_2C
  loc_00E0AEE2: push edx
  loc_00E0AEE3: push eax
  loc_00E0AEE4: mov var_E0, eax
  loc_00E0AEEA: call [ecx+000000A0h]
  loc_00E0AEF0: cmp eax, edi
  loc_00E0AEF2: fnclex
  loc_00E0AEF4: jge 00E0AF0Eh
  loc_00E0AEF6: mov ecx, var_E0
  loc_00E0AEFC: push 000000A0h
  loc_00E0AF01: push 006DCB70h
  loc_00E0AF06: push ecx
  loc_00E0AF07: push eax
  loc_00E0AF08: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0AF0E: mov eax, var_2C
  loc_00E0AF11: lea edx, var_68
  loc_00E0AF14: mov var_50, eax
  loc_00E0AF17: push edx
  loc_00E0AF18: lea eax, var_58
  loc_00E0AF1B: push 00000001h
  loc_00E0AF1D: lea ecx, var_78
  loc_00E0AF20: push eax
  loc_00E0AF21: push ecx
  loc_00E0AF22: mov var_60, 00000001h
  loc_00E0AF29: mov var_68, 00000002h
  loc_00E0AF30: mov var_2C, edi
  loc_00E0AF33: mov var_58, 00000008h
  loc_00E0AF3A: call [004010F0h] ; rtcMidCharVar
  loc_00E0AF40: lea edx, var_78
  loc_00E0AF43: lea eax, var_A8
  loc_00E0AF49: push edx
  loc_00E0AF4A: push eax
  loc_00E0AF4B: mov var_A0, edi
  loc_00E0AF51: mov var_A8, 00008002h
  loc_00E0AF5B: call [004011CCh] ; __vbaVarTstNe
  loc_00E0AF61: mov edi, [00401254h] ; __vbaFreeObj
  loc_00E0AF67: lea ecx, var_38
  loc_00E0AF6A: mov var_E8, ax
  loc_00E0AF71: call edi
  loc_00E0AF73: lea ecx, var_78
  loc_00E0AF76: lea edx, var_68
  loc_00E0AF79: push ecx
  loc_00E0AF7A: lea eax, var_58
  loc_00E0AF7D: push edx
  loc_00E0AF7E: push eax
  loc_00E0AF7F: push 00000003h
  loc_00E0AF81: call [00401038h] ; __vbaFreeVarList
  loc_00E0AF87: mov ecx, [esi]
  loc_00E0AF89: add esp, 00000010h
  loc_00E0AF8C: cmp var_E8, 0000h
  loc_00E0AF94: push esi
  loc_00E0AF95: jz 00E0B1A2h
  loc_00E0AF9B: call [ecx+0000039Ch]
  loc_00E0AFA1: lea edx, var_38
  loc_00E0AFA4: push eax
  loc_00E0AFA5: push edx
  loc_00E0AFA6: call ebx
  loc_00E0AFA8: mov ecx, [eax]
  loc_00E0AFAA: lea edx, var_2C
  loc_00E0AFAD: push edx
  loc_00E0AFAE: push eax
  loc_00E0AFAF: mov var_E0, eax
  loc_00E0AFB5: call [ecx+000000A0h]
  loc_00E0AFBB: test eax, eax
  loc_00E0AFBD: fnclex
  loc_00E0AFBF: jge 00E0AFD9h
  loc_00E0AFC1: mov ecx, var_E0
  loc_00E0AFC7: push 000000A0h
  loc_00E0AFCC: push 006DCB70h
  loc_00E0AFD1: push ecx
  loc_00E0AFD2: push eax
  loc_00E0AFD3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0AFD9: mov eax, var_2C
  loc_00E0AFDC: lea edx, var_58
  loc_00E0AFDF: mov var_50, eax
  loc_00E0AFE2: push 00000005h
  loc_00E0AFE4: lea eax, var_68
  loc_00E0AFE7: push edx
  loc_00E0AFE8: push eax
  loc_00E0AFE9: mov var_2C, 00000000h
  loc_00E0AFF0: mov var_58, 00000008h
  loc_00E0AFF7: call [0040122Ch] ; rtcRightCharVar
  loc_00E0AFFD: lea ecx, var_68
  loc_00E0B000: lea edx, var_98
  loc_00E0B006: push ecx
  loc_00E0B007: lea eax, var_78
  loc_00E0B00A: push edx
  loc_00E0B00B: push eax
  loc_00E0B00C: mov var_90, 00000001h
  loc_00E0B016: mov var_98, 00000002h
  loc_00E0B020: call [004011E0h] ; __vbaVarAdd
  loc_00E0B026: push eax
  loc_00E0B027: call [0040118Ch] ; __vbaI2Var
  loc_00E0B02D: lea ecx, var_38
  loc_00E0B030: mov var_28, eax
  loc_00E0B033: call edi
  loc_00E0B035: lea ecx, var_78
  loc_00E0B038: lea edx, var_68
  loc_00E0B03B: push ecx
  loc_00E0B03C: lea eax, var_58
  loc_00E0B03F: push edx
  loc_00E0B040: push eax
  loc_00E0B041: push 00000003h
  loc_00E0B043: call [00401038h] ; __vbaFreeVarList
  loc_00E0B049: movsx ecx, var_28
  loc_00E0B04D: add esp, 00000010h
  loc_00E0B050: cmp ecx, 000186A0h
  loc_00E0B056: jnz 00E0B135h
  loc_00E0B05C: mov ecx, 0000000Ah
  loc_00E0B061: mov eax, 80020004h
  loc_00E0B066: mov var_88, ecx
  loc_00E0B06C: mov var_78, ecx
  loc_00E0B06F: lea edx, var_A8
  loc_00E0B075: lea ecx, var_68
  loc_00E0B078: mov var_80, eax
  loc_00E0B07B: mov var_70, eax
  loc_00E0B07E: mov var_A0, 006E090Ch ; "Informasi"
  loc_00E0B088: mov var_A8, 00000008h
  loc_00E0B092: call [004011F0h] ; __vbaVarDup
  loc_00E0B098: lea edx, var_98
  loc_00E0B09E: lea ecx, var_58
  loc_00E0B0A1: mov var_90, 006E08B0h ; "NILAI SUDAH MELEBIHI 19999!!! MERESET NILAI"
  loc_00E0B0AB: mov var_98, 00000008h
  loc_00E0B0B5: call [004011F0h] ; __vbaVarDup
  loc_00E0B0BB: lea edx, var_88
  loc_00E0B0C1: lea eax, var_78
  loc_00E0B0C4: push edx
  loc_00E0B0C5: lea ecx, var_68
  loc_00E0B0C8: push eax
  loc_00E0B0C9: push ecx
  loc_00E0B0CA: lea edx, var_58
  loc_00E0B0CD: push 00000040h
  loc_00E0B0CF: push edx
  loc_00E0B0D0: call [004010A8h] ; rtcMsgBox
  loc_00E0B0D6: lea edx, var_D8
  loc_00E0B0DC: lea ecx, var_24
  loc_00E0B0DF: mov var_D0, eax
  loc_00E0B0E5: mov var_D8, 00000003h
  loc_00E0B0EF: call [0040101Ch] ; __vbaVarMove
  loc_00E0B0F5: lea eax, var_88
  loc_00E0B0FB: lea ecx, var_78
  loc_00E0B0FE: push eax
  loc_00E0B0FF: lea edx, var_68
  loc_00E0B102: push ecx
  loc_00E0B103: lea eax, var_58
  loc_00E0B106: push edx
  loc_00E0B107: push eax
  loc_00E0B108: push 00000004h
  loc_00E0B10A: call [00401038h] ; __vbaFreeVarList
  loc_00E0B110: mov ecx, [esi]
  loc_00E0B112: add esp, 00000014h
  loc_00E0B115: push esi
  loc_00E0B116: call [ecx+0000039Ch]
  loc_00E0B11C: lea edx, var_38
  loc_00E0B11F: push eax
  loc_00E0B120: push edx
  loc_00E0B121: call ebx
  loc_00E0B123: mov ecx, [eax]
  loc_00E0B125: mov var_E0, eax
  loc_00E0B12B: push 006E03C8h ; "00001"
  loc_00E0B130: jmp 00E0B995h
  loc_00E0B135: mov eax, [esi]
  loc_00E0B137: push esi
  loc_00E0B138: call [eax+0000039Ch]
  loc_00E0B13E: lea ecx, var_38
  loc_00E0B141: push eax
  loc_00E0B142: push ecx
  loc_00E0B143: call ebx
  loc_00E0B145: mov edx, var_28
  loc_00E0B148: mov ebx, [eax]
  loc_00E0B14A: push edx
  loc_00E0B14B: mov var_E0, eax
  loc_00E0B151: call [0040100Ch] ; __vbaStrI2
  loc_00E0B157: mov edx, eax
  loc_00E0B159: lea ecx, var_2C
  loc_00E0B15C: call [00401228h] ; __vbaStrMove
  loc_00E0B162: mov var_134, ebx
  loc_00E0B168: mov ebx, var_E0
  loc_00E0B16E: push eax
  loc_00E0B16F: mov eax, var_134
  loc_00E0B175: push ebx
  loc_00E0B176: call [eax+000000A4h]
  loc_00E0B17C: test eax, eax
  loc_00E0B17E: fnclex
  loc_00E0B180: jge 00E0B194h
  loc_00E0B182: push 000000A4h
  loc_00E0B187: push 006DCB70h
  loc_00E0B18C: push ebx
  loc_00E0B18D: push eax
  loc_00E0B18E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B194: lea ecx, var_2C
  loc_00E0B197: call [00401258h] ; __vbaFreeStr
  loc_00E0B19D: jmp 00E0BA34h
  loc_00E0B1A2: call [ecx+0000039Ch]
  loc_00E0B1A8: lea edx, var_38
  loc_00E0B1AB: push eax
  loc_00E0B1AC: push edx
  loc_00E0B1AD: call ebx
  loc_00E0B1AF: mov ecx, [eax]
  loc_00E0B1B1: lea edx, var_2C
  loc_00E0B1B4: push edx
  loc_00E0B1B5: push eax
  loc_00E0B1B6: mov var_E0, eax
  loc_00E0B1BC: call [ecx+000000A0h]
  loc_00E0B1C2: test eax, eax
  loc_00E0B1C4: fnclex
  loc_00E0B1C6: jge 00E0B1E0h
  loc_00E0B1C8: mov ecx, var_E0
  loc_00E0B1CE: push 000000A0h
  loc_00E0B1D3: push 006DCB70h
  loc_00E0B1D8: push ecx
  loc_00E0B1D9: push eax
  loc_00E0B1DA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B1E0: mov eax, var_2C
  loc_00E0B1E3: lea edx, var_68
  loc_00E0B1E6: mov var_50, eax
  loc_00E0B1E9: push edx
  loc_00E0B1EA: lea eax, var_58
  loc_00E0B1ED: push 00000002h
  loc_00E0B1EF: lea ecx, var_78
  loc_00E0B1F2: push eax
  loc_00E0B1F3: push ecx
  loc_00E0B1F4: mov var_60, 00000001h
  loc_00E0B1FB: mov var_68, 00000002h
  loc_00E0B202: mov var_2C, 00000000h
  loc_00E0B209: mov var_58, 00000008h
  loc_00E0B210: call [004010F0h] ; rtcMidCharVar
  loc_00E0B216: lea edx, var_78
  loc_00E0B219: lea eax, var_A8
  loc_00E0B21F: push edx
  loc_00E0B220: push eax
  loc_00E0B221: mov var_A0, 00000000h
  loc_00E0B22B: mov var_A8, 00008002h
  loc_00E0B235: call [004011CCh] ; __vbaVarTstNe
  loc_00E0B23B: lea ecx, var_38
  loc_00E0B23E: mov var_E8, ax
  loc_00E0B245: call edi
  loc_00E0B247: lea ecx, var_78
  loc_00E0B24A: lea edx, var_68
  loc_00E0B24D: push ecx
  loc_00E0B24E: lea eax, var_58
  loc_00E0B251: push edx
  loc_00E0B252: push eax
  loc_00E0B253: push 00000003h
  loc_00E0B255: call [00401038h] ; __vbaFreeVarList
  loc_00E0B25B: add esp, 00000010h
  loc_00E0B25E: cmp var_E8, 0000h
  loc_00E0B266: jz 00E0B3A7h
  loc_00E0B26C: mov ecx, [esi]
  loc_00E0B26E: push esi
  loc_00E0B26F: call [ecx+0000039Ch]
  loc_00E0B275: lea edx, var_38
  loc_00E0B278: push eax
  loc_00E0B279: push edx
  loc_00E0B27A: call ebx
  loc_00E0B27C: mov ecx, [eax]
  loc_00E0B27E: lea edx, var_2C
  loc_00E0B281: push edx
  loc_00E0B282: push eax
  loc_00E0B283: mov var_E0, eax
  loc_00E0B289: call [ecx+000000A0h]
  loc_00E0B28F: test eax, eax
  loc_00E0B291: fnclex
  loc_00E0B293: jge 00E0B2ADh
  loc_00E0B295: mov ecx, var_E0
  loc_00E0B29B: push 000000A0h
  loc_00E0B2A0: push 006DCB70h
  loc_00E0B2A5: push ecx
  loc_00E0B2A6: push eax
  loc_00E0B2A7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B2AD: mov eax, var_2C
  loc_00E0B2B0: lea edx, var_58
  loc_00E0B2B3: mov var_50, eax
  loc_00E0B2B6: push 00000004h
  loc_00E0B2B8: lea eax, var_68
  loc_00E0B2BB: push edx
  loc_00E0B2BC: push eax
  loc_00E0B2BD: mov var_2C, 00000000h
  loc_00E0B2C4: mov var_58, 00000008h
  loc_00E0B2CB: call [0040122Ch] ; rtcRightCharVar
  loc_00E0B2D1: lea ecx, var_68
  loc_00E0B2D4: lea edx, var_98
  loc_00E0B2DA: push ecx
  loc_00E0B2DB: lea eax, var_78
  loc_00E0B2DE: push edx
  loc_00E0B2DF: push eax
  loc_00E0B2E0: mov var_90, 00000001h
  loc_00E0B2EA: mov var_98, 00000002h
  loc_00E0B2F4: call [004011E0h] ; __vbaVarAdd
  loc_00E0B2FA: push eax
  loc_00E0B2FB: call [0040118Ch] ; __vbaI2Var
  loc_00E0B301: lea ecx, var_38
  loc_00E0B304: mov var_28, eax
  loc_00E0B307: call edi
  loc_00E0B309: lea ecx, var_78
  loc_00E0B30C: lea edx, var_68
  loc_00E0B30F: push ecx
  loc_00E0B310: lea eax, var_58
  loc_00E0B313: push edx
  loc_00E0B314: push eax
  loc_00E0B315: push 00000003h
  loc_00E0B317: call [00401038h] ; __vbaFreeVarList
  loc_00E0B31D: add esp, 00000010h
  loc_00E0B320: cmp var_28, 2710h
  loc_00E0B326: jnz 00E0B34Ah
  loc_00E0B328: mov ecx, [esi]
  loc_00E0B32A: push esi
  loc_00E0B32B: call [ecx+0000039Ch]
  loc_00E0B331: lea edx, var_38
  loc_00E0B334: push eax
  loc_00E0B335: push edx
  loc_00E0B336: call ebx
  loc_00E0B338: mov ecx, [eax]
  loc_00E0B33A: mov var_E0, eax
  loc_00E0B340: push 006E0924h ; "10000"
  loc_00E0B345: jmp 00E0B995h
  loc_00E0B34A: mov eax, [esi]
  loc_00E0B34C: push esi
  loc_00E0B34D: call [eax+0000039Ch]
  loc_00E0B353: lea ecx, var_38
  loc_00E0B356: push eax
  loc_00E0B357: push ecx
  loc_00E0B358: call ebx
  loc_00E0B35A: mov edx, var_28
  loc_00E0B35D: mov ebx, [eax]
  loc_00E0B35F: push 006E0934h
  loc_00E0B364: push edx
  loc_00E0B365: mov var_E0, eax
  loc_00E0B36B: call [0040100Ch] ; __vbaStrI2
  loc_00E0B371: mov edx, eax
  loc_00E0B373: lea ecx, var_2C
  loc_00E0B376: call [00401228h] ; __vbaStrMove
  loc_00E0B37C: push eax
  loc_00E0B37D: call [0040106Ch] ; __vbaStrCat
  loc_00E0B383: mov edx, eax
  loc_00E0B385: lea ecx, var_30
  loc_00E0B388: call [00401228h] ; __vbaStrMove
  loc_00E0B38E: mov var_138, ebx
  loc_00E0B394: mov ebx, var_E0
  loc_00E0B39A: push eax
  loc_00E0B39B: mov eax, var_138
  loc_00E0B3A1: push ebx
  loc_00E0B3A2: jmp 00E0B7DBh
  loc_00E0B3A7: mov eax, [esi]
  loc_00E0B3A9: push esi
  loc_00E0B3AA: call [eax+0000039Ch]
  loc_00E0B3B0: lea ecx, var_38
  loc_00E0B3B3: push eax
  loc_00E0B3B4: push ecx
  loc_00E0B3B5: call ebx
  loc_00E0B3B7: mov edx, [eax]
  loc_00E0B3B9: lea ecx, var_2C
  loc_00E0B3BC: push ecx
  loc_00E0B3BD: push eax
  loc_00E0B3BE: mov var_E0, eax
  loc_00E0B3C4: call [edx+000000A0h]
  loc_00E0B3CA: test eax, eax
  loc_00E0B3CC: fnclex
  loc_00E0B3CE: jge 00E0B3E8h
  loc_00E0B3D0: mov edx, var_E0
  loc_00E0B3D6: push 000000A0h
  loc_00E0B3DB: push 006DCB70h
  loc_00E0B3E0: push edx
  loc_00E0B3E1: push eax
  loc_00E0B3E2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B3E8: mov eax, var_2C
  loc_00E0B3EB: lea ecx, var_58
  loc_00E0B3EE: mov var_50, eax
  loc_00E0B3F1: lea eax, var_68
  loc_00E0B3F4: push eax
  loc_00E0B3F5: push 00000003h
  loc_00E0B3F7: lea edx, var_78
  loc_00E0B3FA: push ecx
  loc_00E0B3FB: push edx
  loc_00E0B3FC: mov var_60, 00000001h
  loc_00E0B403: mov var_68, 00000002h
  loc_00E0B40A: mov var_2C, 00000000h
  loc_00E0B411: mov var_58, 00000008h
  loc_00E0B418: call [004010F0h] ; rtcMidCharVar
  loc_00E0B41E: lea eax, var_78
  loc_00E0B421: lea ecx, var_A8
  loc_00E0B427: push eax
  loc_00E0B428: push ecx
  loc_00E0B429: mov var_A0, 00000000h
  loc_00E0B433: mov var_A8, 00008002h
  loc_00E0B43D: call [004011CCh] ; __vbaVarTstNe
  loc_00E0B443: lea ecx, var_38
  loc_00E0B446: mov var_E8, ax
  loc_00E0B44D: call edi
  loc_00E0B44F: lea edx, var_78
  loc_00E0B452: lea eax, var_68
  loc_00E0B455: push edx
  loc_00E0B456: lea ecx, var_58
  loc_00E0B459: push eax
  loc_00E0B45A: push ecx
  loc_00E0B45B: push 00000003h
  loc_00E0B45D: call [00401038h] ; __vbaFreeVarList
  loc_00E0B463: add esp, 00000010h
  loc_00E0B466: cmp var_E8, 0000h
  loc_00E0B46E: jz 00E0B5AFh
  loc_00E0B474: mov edx, [esi]
  loc_00E0B476: push esi
  loc_00E0B477: call [edx+0000039Ch]
  loc_00E0B47D: push eax
  loc_00E0B47E: lea eax, var_38
  loc_00E0B481: push eax
  loc_00E0B482: call ebx
  loc_00E0B484: mov ecx, [eax]
  loc_00E0B486: lea edx, var_2C
  loc_00E0B489: push edx
  loc_00E0B48A: push eax
  loc_00E0B48B: mov var_E0, eax
  loc_00E0B491: call [ecx+000000A0h]
  loc_00E0B497: test eax, eax
  loc_00E0B499: fnclex
  loc_00E0B49B: jge 00E0B4B5h
  loc_00E0B49D: mov ecx, var_E0
  loc_00E0B4A3: push 000000A0h
  loc_00E0B4A8: push 006DCB70h
  loc_00E0B4AD: push ecx
  loc_00E0B4AE: push eax
  loc_00E0B4AF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B4B5: mov eax, var_2C
  loc_00E0B4B8: lea edx, var_58
  loc_00E0B4BB: mov var_50, eax
  loc_00E0B4BE: push 00000003h
  loc_00E0B4C0: lea eax, var_68
  loc_00E0B4C3: push edx
  loc_00E0B4C4: push eax
  loc_00E0B4C5: mov var_2C, 00000000h
  loc_00E0B4CC: mov var_58, 00000008h
  loc_00E0B4D3: call [0040122Ch] ; rtcRightCharVar
  loc_00E0B4D9: lea ecx, var_68
  loc_00E0B4DC: lea edx, var_98
  loc_00E0B4E2: push ecx
  loc_00E0B4E3: lea eax, var_78
  loc_00E0B4E6: push edx
  loc_00E0B4E7: push eax
  loc_00E0B4E8: mov var_90, 00000001h
  loc_00E0B4F2: mov var_98, 00000002h
  loc_00E0B4FC: call [004011E0h] ; __vbaVarAdd
  loc_00E0B502: push eax
  loc_00E0B503: call [0040118Ch] ; __vbaI2Var
  loc_00E0B509: lea ecx, var_38
  loc_00E0B50C: mov var_28, eax
  loc_00E0B50F: call edi
  loc_00E0B511: lea ecx, var_78
  loc_00E0B514: lea edx, var_68
  loc_00E0B517: push ecx
  loc_00E0B518: lea eax, var_58
  loc_00E0B51B: push edx
  loc_00E0B51C: push eax
  loc_00E0B51D: push 00000003h
  loc_00E0B51F: call [00401038h] ; __vbaFreeVarList
  loc_00E0B525: add esp, 00000010h
  loc_00E0B528: cmp var_28, 03E8h
  loc_00E0B52E: jnz 00E0B552h
  loc_00E0B530: mov ecx, [esi]
  loc_00E0B532: push esi
  loc_00E0B533: call [ecx+0000039Ch]
  loc_00E0B539: lea edx, var_38
  loc_00E0B53C: push eax
  loc_00E0B53D: push edx
  loc_00E0B53E: call ebx
  loc_00E0B540: mov ecx, [eax]
  loc_00E0B542: mov var_E0, eax
  loc_00E0B548: push 006E093Ch ; "01000"
  loc_00E0B54D: jmp 00E0B995h
  loc_00E0B552: mov eax, [esi]
  loc_00E0B554: push esi
  loc_00E0B555: call [eax+0000039Ch]
  loc_00E0B55B: lea ecx, var_38
  loc_00E0B55E: push eax
  loc_00E0B55F: push ecx
  loc_00E0B560: call ebx
  loc_00E0B562: mov edx, var_28
  loc_00E0B565: mov ebx, [eax]
  loc_00E0B567: push 006E094Ch ; "00"
  loc_00E0B56C: push edx
  loc_00E0B56D: mov var_E0, eax
  loc_00E0B573: call [0040100Ch] ; __vbaStrI2
  loc_00E0B579: mov edx, eax
  loc_00E0B57B: lea ecx, var_2C
  loc_00E0B57E: call [00401228h] ; __vbaStrMove
  loc_00E0B584: push eax
  loc_00E0B585: call [0040106Ch] ; __vbaStrCat
  loc_00E0B58B: mov edx, eax
  loc_00E0B58D: lea ecx, var_30
  loc_00E0B590: call [00401228h] ; __vbaStrMove
  loc_00E0B596: mov var_13C, ebx
  loc_00E0B59C: mov ebx, var_E0
  loc_00E0B5A2: push eax
  loc_00E0B5A3: mov eax, var_13C
  loc_00E0B5A9: push ebx
  loc_00E0B5AA: jmp 00E0B7DBh
  loc_00E0B5AF: mov eax, [esi]
  loc_00E0B5B1: push esi
  loc_00E0B5B2: call [eax+0000039Ch]
  loc_00E0B5B8: lea ecx, var_38
  loc_00E0B5BB: push eax
  loc_00E0B5BC: push ecx
  loc_00E0B5BD: call ebx
  loc_00E0B5BF: mov edx, [eax]
  loc_00E0B5C1: lea ecx, var_2C
  loc_00E0B5C4: push ecx
  loc_00E0B5C5: push eax
  loc_00E0B5C6: mov var_E0, eax
  loc_00E0B5CC: call [edx+000000A0h]
  loc_00E0B5D2: test eax, eax
  loc_00E0B5D4: fnclex
  loc_00E0B5D6: jge 00E0B5F0h
  loc_00E0B5D8: mov edx, var_E0
  loc_00E0B5DE: push 000000A0h
  loc_00E0B5E3: push 006DCB70h
  loc_00E0B5E8: push edx
  loc_00E0B5E9: push eax
  loc_00E0B5EA: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B5F0: mov eax, var_2C
  loc_00E0B5F3: lea ecx, var_58
  loc_00E0B5F6: mov var_50, eax
  loc_00E0B5F9: lea eax, var_68
  loc_00E0B5FC: push eax
  loc_00E0B5FD: push 00000004h
  loc_00E0B5FF: lea edx, var_78
  loc_00E0B602: push ecx
  loc_00E0B603: push edx
  loc_00E0B604: mov var_60, 00000001h
  loc_00E0B60B: mov var_68, 00000002h
  loc_00E0B612: mov var_2C, 00000000h
  loc_00E0B619: mov var_58, 00000008h
  loc_00E0B620: call [004010F0h] ; rtcMidCharVar
  loc_00E0B626: lea eax, var_78
  loc_00E0B629: lea ecx, var_A8
  loc_00E0B62F: push eax
  loc_00E0B630: push ecx
  loc_00E0B631: mov var_A0, 00000000h
  loc_00E0B63B: mov var_A8, 00008002h
  loc_00E0B645: call [004011CCh] ; __vbaVarTstNe
  loc_00E0B64B: lea ecx, var_38
  loc_00E0B64E: mov var_E8, ax
  loc_00E0B655: call edi
  loc_00E0B657: lea edx, var_78
  loc_00E0B65A: lea eax, var_68
  loc_00E0B65D: push edx
  loc_00E0B65E: lea ecx, var_58
  loc_00E0B661: push eax
  loc_00E0B662: push ecx
  loc_00E0B663: push 00000003h
  loc_00E0B665: call [00401038h] ; __vbaFreeVarList
  loc_00E0B66B: add esp, 00000010h
  loc_00E0B66E: cmp var_E8, 0000h
  loc_00E0B676: jz 00E0B7F0h
  loc_00E0B67C: mov edx, [esi]
  loc_00E0B67E: push esi
  loc_00E0B67F: call [edx+0000039Ch]
  loc_00E0B685: push eax
  loc_00E0B686: lea eax, var_38
  loc_00E0B689: push eax
  loc_00E0B68A: call ebx
  loc_00E0B68C: mov ecx, [eax]
  loc_00E0B68E: lea edx, var_2C
  loc_00E0B691: push edx
  loc_00E0B692: push eax
  loc_00E0B693: mov var_E0, eax
  loc_00E0B699: call [ecx+000000A0h]
  loc_00E0B69F: test eax, eax
  loc_00E0B6A1: fnclex
  loc_00E0B6A3: jge 00E0B6BDh
  loc_00E0B6A5: mov ecx, var_E0
  loc_00E0B6AB: push 000000A0h
  loc_00E0B6B0: push 006DCB70h
  loc_00E0B6B5: push ecx
  loc_00E0B6B6: push eax
  loc_00E0B6B7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B6BD: mov eax, var_2C
  loc_00E0B6C0: lea edx, var_58
  loc_00E0B6C3: mov var_50, eax
  loc_00E0B6C6: push 00000002h
  loc_00E0B6C8: lea eax, var_68
  loc_00E0B6CB: push edx
  loc_00E0B6CC: push eax
  loc_00E0B6CD: mov var_2C, 00000000h
  loc_00E0B6D4: mov var_58, 00000008h
  loc_00E0B6DB: call [0040122Ch] ; rtcRightCharVar
  loc_00E0B6E1: lea ecx, var_68
  loc_00E0B6E4: lea edx, var_98
  loc_00E0B6EA: push ecx
  loc_00E0B6EB: lea eax, var_78
  loc_00E0B6EE: push edx
  loc_00E0B6EF: push eax
  loc_00E0B6F0: mov var_90, 00000001h
  loc_00E0B6FA: mov var_98, 00000002h
  loc_00E0B704: call [004011E0h] ; __vbaVarAdd
  loc_00E0B70A: push eax
  loc_00E0B70B: call [0040118Ch] ; __vbaI2Var
  loc_00E0B711: lea ecx, var_38
  loc_00E0B714: mov var_28, eax
  loc_00E0B717: call edi
  loc_00E0B719: lea ecx, var_78
  loc_00E0B71C: lea edx, var_68
  loc_00E0B71F: push ecx
  loc_00E0B720: lea eax, var_58
  loc_00E0B723: push edx
  loc_00E0B724: push eax
  loc_00E0B725: push 00000003h
  loc_00E0B727: call [00401038h] ; __vbaFreeVarList
  loc_00E0B72D: add esp, 00000010h
  loc_00E0B730: cmp var_28, 0064h
  loc_00E0B735: jnz 00E0B783h
  loc_00E0B737: mov ecx, [esi]
  loc_00E0B739: push esi
  loc_00E0B73A: call [ecx+0000039Ch]
  loc_00E0B740: lea edx, var_38
  loc_00E0B743: push eax
  loc_00E0B744: push edx
  loc_00E0B745: call ebx
  loc_00E0B747: mov ecx, [eax]
  loc_00E0B749: push 006E0958h ; "00100"
  loc_00E0B74E: push eax
  loc_00E0B74F: mov var_E0, eax
  loc_00E0B755: call [ecx+000000A4h]
  loc_00E0B75B: test eax, eax
  loc_00E0B75D: fnclex
  loc_00E0B75F: jge 00E0B779h
  loc_00E0B761: mov edx, var_E0
  loc_00E0B767: push 000000A4h
  loc_00E0B76C: push 006DCB70h
  loc_00E0B771: push edx
  loc_00E0B772: push eax
  loc_00E0B773: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B779: lea ecx, var_38
  loc_00E0B77C: call edi
  loc_00E0B77E: jmp 00E0BA3Fh
  loc_00E0B783: mov eax, [esi]
  loc_00E0B785: push esi
  loc_00E0B786: call [eax+0000039Ch]
  loc_00E0B78C: lea ecx, var_38
  loc_00E0B78F: push eax
  loc_00E0B790: push ecx
  loc_00E0B791: call ebx
  loc_00E0B793: mov edx, var_28
  loc_00E0B796: mov ebx, [eax]
  loc_00E0B798: push 006E0968h ; "000"
  loc_00E0B79D: push edx
  loc_00E0B79E: mov var_E0, eax
  loc_00E0B7A4: call [0040100Ch] ; __vbaStrI2
  loc_00E0B7AA: mov edx, eax
  loc_00E0B7AC: lea ecx, var_2C
  loc_00E0B7AF: call [00401228h] ; __vbaStrMove
  loc_00E0B7B5: push eax
  loc_00E0B7B6: call [0040106Ch] ; __vbaStrCat
  loc_00E0B7BC: mov edx, eax
  loc_00E0B7BE: lea ecx, var_30
  loc_00E0B7C1: call [00401228h] ; __vbaStrMove
  loc_00E0B7C7: mov var_140, ebx
  loc_00E0B7CD: mov ebx, var_E0
  loc_00E0B7D3: push eax
  loc_00E0B7D4: mov eax, var_140
  loc_00E0B7DA: push ebx
  loc_00E0B7DB: call [eax+000000A4h]
  loc_00E0B7E1: test eax, eax
  loc_00E0B7E3: fnclex
  loc_00E0B7E5: jge 00E0BA21h
  loc_00E0B7EB: jmp 00E0BA0Fh
  loc_00E0B7F0: mov eax, [esi]
  loc_00E0B7F2: push esi
  loc_00E0B7F3: call [eax+0000039Ch]
  loc_00E0B7F9: lea ecx, var_38
  loc_00E0B7FC: push eax
  loc_00E0B7FD: push ecx
  loc_00E0B7FE: call ebx
  loc_00E0B800: mov edx, [eax]
  loc_00E0B802: lea ecx, var_2C
  loc_00E0B805: push ecx
  loc_00E0B806: push eax
  loc_00E0B807: mov var_E0, eax
  loc_00E0B80D: call [edx+000000A0h]
  loc_00E0B813: test eax, eax
  loc_00E0B815: fnclex
  loc_00E0B817: jge 00E0B831h
  loc_00E0B819: mov edx, var_E0
  loc_00E0B81F: push 000000A0h
  loc_00E0B824: push 006DCB70h
  loc_00E0B829: push edx
  loc_00E0B82A: push eax
  loc_00E0B82B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B831: mov eax, var_2C
  loc_00E0B834: lea ecx, var_58
  loc_00E0B837: mov var_50, eax
  loc_00E0B83A: lea eax, var_68
  loc_00E0B83D: push eax
  loc_00E0B83E: push 00000005h
  loc_00E0B840: lea edx, var_78
  loc_00E0B843: push ecx
  loc_00E0B844: push edx
  loc_00E0B845: mov var_60, 00000001h
  loc_00E0B84C: mov var_68, 00000002h
  loc_00E0B853: mov var_2C, 00000000h
  loc_00E0B85A: mov var_58, 00000008h
  loc_00E0B861: call [004010F0h] ; rtcMidCharVar
  loc_00E0B867: lea eax, var_78
  loc_00E0B86A: lea ecx, var_A8
  loc_00E0B870: push eax
  loc_00E0B871: push ecx
  loc_00E0B872: mov var_A0, 00000000h
  loc_00E0B87C: mov var_A8, 00008002h
  loc_00E0B886: call [004011CCh] ; __vbaVarTstNe
  loc_00E0B88C: lea ecx, var_38
  loc_00E0B88F: mov var_E8, ax
  loc_00E0B896: call edi
  loc_00E0B898: lea edx, var_78
  loc_00E0B89B: lea eax, var_68
  loc_00E0B89E: push edx
  loc_00E0B89F: lea ecx, var_58
  loc_00E0B8A2: push eax
  loc_00E0B8A3: push ecx
  loc_00E0B8A4: push 00000003h
  loc_00E0B8A6: call [00401038h] ; __vbaFreeVarList
  loc_00E0B8AC: add esp, 00000010h
  loc_00E0B8AF: cmp var_E8, 0000h
  loc_00E0B8B7: jz 00E0BA3Fh
  loc_00E0B8BD: mov edx, [esi]
  loc_00E0B8BF: push esi
  loc_00E0B8C0: call [edx+0000039Ch]
  loc_00E0B8C6: push eax
  loc_00E0B8C7: lea eax, var_38
  loc_00E0B8CA: push eax
  loc_00E0B8CB: call ebx
  loc_00E0B8CD: mov ecx, [eax]
  loc_00E0B8CF: lea edx, var_2C
  loc_00E0B8D2: push edx
  loc_00E0B8D3: push eax
  loc_00E0B8D4: mov var_E0, eax
  loc_00E0B8DA: call [ecx+000000A0h]
  loc_00E0B8E0: test eax, eax
  loc_00E0B8E2: fnclex
  loc_00E0B8E4: jge 00E0B8FEh
  loc_00E0B8E6: mov ecx, var_E0
  loc_00E0B8EC: push 000000A0h
  loc_00E0B8F1: push 006DCB70h
  loc_00E0B8F6: push ecx
  loc_00E0B8F7: push eax
  loc_00E0B8F8: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0B8FE: mov eax, var_2C
  loc_00E0B901: lea edx, var_58
  loc_00E0B904: mov var_50, eax
  loc_00E0B907: push 00000001h
  loc_00E0B909: lea eax, var_68
  loc_00E0B90C: push edx
  loc_00E0B90D: push eax
  loc_00E0B90E: mov var_2C, 00000000h
  loc_00E0B915: mov var_58, 00000008h
  loc_00E0B91C: call [0040122Ch] ; rtcRightCharVar
  loc_00E0B922: lea ecx, var_68
  loc_00E0B925: lea edx, var_98
  loc_00E0B92B: push ecx
  loc_00E0B92C: lea eax, var_78
  loc_00E0B92F: push edx
  loc_00E0B930: push eax
  loc_00E0B931: mov var_90, 00000001h
  loc_00E0B93B: mov var_98, 00000002h
  loc_00E0B945: call [004011E0h] ; __vbaVarAdd
  loc_00E0B94B: push eax
  loc_00E0B94C: call [0040118Ch] ; __vbaI2Var
  loc_00E0B952: lea ecx, var_38
  loc_00E0B955: mov var_28, eax
  loc_00E0B958: call edi
  loc_00E0B95A: lea ecx, var_78
  loc_00E0B95D: lea edx, var_68
  loc_00E0B960: push ecx
  loc_00E0B961: lea eax, var_58
  loc_00E0B964: push edx
  loc_00E0B965: push eax
  loc_00E0B966: push 00000003h
  loc_00E0B968: call [00401038h] ; __vbaFreeVarList
  loc_00E0B96E: add esp, 00000010h
  loc_00E0B971: cmp var_28, 000Ah
  loc_00E0B976: jnz 00E0B9ABh
  loc_00E0B978: mov ecx, [esi]
  loc_00E0B97A: push esi
  loc_00E0B97B: call [ecx+0000039Ch]
  loc_00E0B981: lea edx, var_38
  loc_00E0B984: push eax
  loc_00E0B985: push edx
  loc_00E0B986: call ebx
  loc_00E0B988: mov ecx, [eax]
  loc_00E0B98A: mov var_E0, eax
  loc_00E0B990: push 006E0974h ; "00010"
  loc_00E0B995: push eax
  loc_00E0B996: call [ecx+000000A4h]
  loc_00E0B99C: test eax, eax
  loc_00E0B99E: fnclex
  loc_00E0B9A0: jge 00E0B779h
  loc_00E0B9A6: jmp 00E0B761h
  loc_00E0B9AB: mov eax, [esi]
  loc_00E0B9AD: push esi
  loc_00E0B9AE: call [eax+0000039Ch]
  loc_00E0B9B4: lea ecx, var_38
  loc_00E0B9B7: push eax
  loc_00E0B9B8: push ecx
  loc_00E0B9B9: call ebx
  loc_00E0B9BB: mov edx, var_28
  loc_00E0B9BE: mov ebx, [eax]
  loc_00E0B9C0: push 006E0984h ; "0000"
  loc_00E0B9C5: push edx
  loc_00E0B9C6: mov var_E0, eax
  loc_00E0B9CC: call [0040100Ch] ; __vbaStrI2
  loc_00E0B9D2: mov edx, eax
  loc_00E0B9D4: lea ecx, var_2C
  loc_00E0B9D7: call [00401228h] ; __vbaStrMove
  loc_00E0B9DD: push eax
  loc_00E0B9DE: call [0040106Ch] ; __vbaStrCat
  loc_00E0B9E4: mov edx, eax
  loc_00E0B9E6: lea ecx, var_30
  loc_00E0B9E9: call [00401228h] ; __vbaStrMove
  loc_00E0B9EF: mov var_144, ebx
  loc_00E0B9F5: mov ebx, var_E0
  loc_00E0B9FB: push eax
  loc_00E0B9FC: mov eax, var_144
  loc_00E0BA02: push ebx
  loc_00E0BA03: call [eax+000000A4h]
  loc_00E0BA09: test eax, eax
  loc_00E0BA0B: fnclex
  loc_00E0BA0D: jge 00E0BA21h
  loc_00E0BA0F: push 000000A4h
  loc_00E0BA14: push 006DCB70h
  loc_00E0BA19: push ebx
  loc_00E0BA1A: push eax
  loc_00E0BA1B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0BA21: lea ecx, var_30
  loc_00E0BA24: lea edx, var_2C
  loc_00E0BA27: push ecx
  loc_00E0BA28: push edx
  loc_00E0BA29: push 00000002h
  loc_00E0BA2B: call [004011BCh] ; __vbaFreeStrList
  loc_00E0BA31: add esp, 0000000Ch
  loc_00E0BA34: lea ecx, var_38
  loc_00E0BA37: call edi
  loc_00E0BA39: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0BA3F: mov eax, [esi]
  loc_00E0BA41: push esi
  loc_00E0BA42: call [eax+0000039Ch]
  loc_00E0BA48: lea ecx, var_38
  loc_00E0BA4B: push eax
  loc_00E0BA4C: push ecx
  loc_00E0BA4D: call ebx
  loc_00E0BA4F: mov edx, [eax]
  loc_00E0BA51: lea ecx, var_2C
  loc_00E0BA54: push ecx
  loc_00E0BA55: push eax
  loc_00E0BA56: mov var_E0, eax
  loc_00E0BA5C: call [edx+000000A0h]
  loc_00E0BA62: test eax, eax
  loc_00E0BA64: fnclex
  loc_00E0BA66: jge 00E0BA80h
  loc_00E0BA68: mov edx, var_E0
  loc_00E0BA6E: push 000000A0h
  loc_00E0BA73: push 006DCB70h
  loc_00E0BA78: push edx
  loc_00E0BA79: push eax
  loc_00E0BA7A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0BA80: mov eax, var_2C
  loc_00E0BA83: push 006E0994h ; "select * from biodata where no_daftar ='"
  loc_00E0BA88: push eax
  loc_00E0BA89: call [0040106Ch] ; __vbaStrCat
  loc_00E0BA8F: mov edx, eax
  loc_00E0BA91: lea ecx, var_30
  loc_00E0BA94: call [00401228h] ; __vbaStrMove
  loc_00E0BA9A: push eax
  loc_00E0BA9B: push 006DCB84h ; "'"
  loc_00E0BAA0: call [0040106Ch] ; __vbaStrCat
  loc_00E0BAA6: mov edx, eax
  loc_00E0BAA8: lea ecx, var_34
  loc_00E0BAAB: call [00401228h] ; __vbaStrMove
  loc_00E0BAB1: mov edx, eax
  loc_00E0BAB3: lea ecx, [esi+00000034h]
  loc_00E0BAB6: call [004011B0h] ; __vbaStrCopy
  loc_00E0BABC: lea ecx, var_34
  loc_00E0BABF: lea edx, var_30
  loc_00E0BAC2: push ecx
  loc_00E0BAC3: lea eax, var_2C
  loc_00E0BAC6: push edx
  loc_00E0BAC7: push eax
  loc_00E0BAC8: push 00000003h
  loc_00E0BACA: call [004011BCh] ; __vbaFreeStrList
  loc_00E0BAD0: add esp, 00000010h
  loc_00E0BAD3: lea ecx, var_38
  loc_00E0BAD6: call edi
  loc_00E0BAD8: sub esp, 00000010h
  loc_00E0BADB: mov ecx, 00004008h
  loc_00E0BAE0: mov edx, esp
  loc_00E0BAE2: mov var_98, ecx
  loc_00E0BAE8: lea eax, [esi+00000034h]
  loc_00E0BAEB: push 0000000Eh
  loc_00E0BAED: mov [edx], ecx
  loc_00E0BAEF: mov ecx, var_94
  loc_00E0BAF5: mov var_90, eax
  loc_00E0BAFB: push esi
  loc_00E0BAFC: mov [edx+00000004h], ecx
  loc_00E0BAFF: mov ecx, [esi]
  loc_00E0BB01: mov [edx+00000008h], eax
  loc_00E0BB04: mov eax, var_8C
  loc_00E0BB0A: mov [edx+0000000Ch], eax
  loc_00E0BB0D: call [ecx+00000490h]
  loc_00E0BB13: lea edx, var_38
  loc_00E0BB16: push eax
  loc_00E0BB17: push edx
  loc_00E0BB18: call ebx
  loc_00E0BB1A: push eax
  loc_00E0BB1B: call [00401238h] ; __vbaLateIdSt
  loc_00E0BB21: lea ecx, var_38
  loc_00E0BB24: call edi
  loc_00E0BB26: mov eax, [esi]
  loc_00E0BB28: push 00000000h
  loc_00E0BB2A: push 00000019h
  loc_00E0BB2C: push esi
  loc_00E0BB2D: call [eax+00000490h]
  loc_00E0BB33: lea ecx, var_38
  loc_00E0BB36: push eax
  loc_00E0BB37: push ecx
  loc_00E0BB38: call ebx
  loc_00E0BB3A: push eax
  loc_00E0BB3B: call [00401030h] ; __vbaLateIdCall
  loc_00E0BB41: add esp, 0000000Ch
  loc_00E0BB44: lea ecx, var_38
  loc_00E0BB47: call edi
  loc_00E0BB49: mov edx, [esi]
  loc_00E0BB4B: push 006DCBD8h
  loc_00E0BB50: push 00000000h
  loc_00E0BB52: push 00000018h
  loc_00E0BB54: push esi
  loc_00E0BB55: call [edx+00000490h]
  loc_00E0BB5B: push eax
  loc_00E0BB5C: lea eax, var_38
  loc_00E0BB5F: push eax
  loc_00E0BB60: call ebx
  loc_00E0BB62: lea ecx, var_58
  loc_00E0BB65: push eax
  loc_00E0BB66: push ecx
  loc_00E0BB67: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0BB6D: add esp, 00000010h
  loc_00E0BB70: push eax
  loc_00E0BB71: call [00401120h] ; __vbaCastObjVar
  loc_00E0BB77: lea edx, var_3C
  loc_00E0BB7A: push eax
  loc_00E0BB7B: push edx
  loc_00E0BB7C: call ebx
  loc_00E0BB7E: mov ecx, [eax]
  loc_00E0BB80: lea edx, var_DC
  loc_00E0BB86: push edx
  loc_00E0BB87: push eax
  loc_00E0BB88: mov var_E0, eax
  loc_00E0BB8E: call [ecx+00000050h]
  loc_00E0BB91: test eax, eax
  loc_00E0BB93: fnclex
  loc_00E0BB95: jge 00E0BBACh
  loc_00E0BB97: mov ecx, var_E0
  loc_00E0BB9D: push 00000050h
  loc_00E0BB9F: push 006DCBE8h
  loc_00E0BBA4: push ecx
  loc_00E0BBA5: push eax
  loc_00E0BBA6: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0BBAC: mov dx, var_DC
  loc_00E0BBB3: lea eax, var_3C
  loc_00E0BBB6: lea ecx, var_38
  loc_00E0BBB9: push eax
  loc_00E0BBBA: push ecx
  loc_00E0BBBB: push 00000002h
  loc_00E0BBBD: mov var_E8, dx
  loc_00E0BBC4: call [00401048h] ; __vbaFreeObjList
  loc_00E0BBCA: add esp, 0000000Ch
  loc_00E0BBCD: lea ecx, var_58
  loc_00E0BBD0: call [00401024h] ; __vbaFreeVar
  loc_00E0BBD6: cmp var_E8, 0000h
  loc_00E0BBDE: jz 00E0BF65h
  loc_00E0BBE4: mov edx, [esi]
  loc_00E0BBE6: push 006DCBD8h
  loc_00E0BBEB: push 00000000h
  loc_00E0BBED: push 00000018h
  loc_00E0BBEF: push esi
  loc_00E0BBF0: call [edx+00000490h]
  loc_00E0BBF6: push eax
  loc_00E0BBF7: lea eax, var_38
  loc_00E0BBFA: push eax
  loc_00E0BBFB: call ebx
  loc_00E0BBFD: lea ecx, var_58
  loc_00E0BC00: push eax
  loc_00E0BC01: push ecx
  loc_00E0BC02: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0BC08: add esp, 00000010h
  loc_00E0BC0B: push eax
  loc_00E0BC0C: call [00401120h] ; __vbaCastObjVar
  loc_00E0BC12: lea edx, var_3C
  loc_00E0BC15: push eax
  loc_00E0BC16: push edx
  loc_00E0BC17: call ebx
  loc_00E0BC19: sub esp, 00000010h
  loc_00E0BC1C: mov ecx, 0000000Ah
  loc_00E0BC21: mov edx, esp
  loc_00E0BC23: mov var_A8, ecx
  loc_00E0BC29: mov var_98, ecx
  loc_00E0BC2F: mov var_A0, 80020004h
  loc_00E0BC39: mov [edx], ecx
  loc_00E0BC3B: mov ecx, var_A4
  loc_00E0BC41: sub esp, 00000010h
  loc_00E0BC44: mov var_90, 80020004h
  loc_00E0BC4E: mov [edx+00000004h], ecx
  loc_00E0BC51: mov ecx, var_A0
  loc_00E0BC57: mov var_E0, eax
  loc_00E0BC5D: mov eax, [eax]
  loc_00E0BC5F: mov [edx+00000008h], ecx
  loc_00E0BC62: mov ecx, var_9C
  loc_00E0BC68: mov [edx+0000000Ch], ecx
  loc_00E0BC6B: mov ecx, var_98
  loc_00E0BC71: mov edx, esp
  loc_00E0BC73: mov [edx], ecx
  loc_00E0BC75: mov ecx, var_94
  loc_00E0BC7B: mov [edx+00000004h], ecx
  loc_00E0BC7E: mov ecx, var_90
  loc_00E0BC84: mov [edx+00000008h], ecx
  loc_00E0BC87: mov ecx, var_8C
  loc_00E0BC8D: mov [edx+0000000Ch], ecx
  loc_00E0BC90: mov edx, var_E0
  loc_00E0BC96: push edx
  loc_00E0BC97: call [eax+00000078h]
  loc_00E0BC9A: test eax, eax
  loc_00E0BC9C: fnclex
  loc_00E0BC9E: jge 00E0BCB5h
  loc_00E0BCA0: mov ecx, var_E0
  loc_00E0BCA6: push 00000078h
  loc_00E0BCA8: push 006DCBE8h
  loc_00E0BCAD: push ecx
  loc_00E0BCAE: push eax
  loc_00E0BCAF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0BCB5: lea edx, var_3C
  loc_00E0BCB8: lea eax, var_38
  loc_00E0BCBB: push edx
  loc_00E0BCBC: push eax
  loc_00E0BCBD: push 00000002h
  loc_00E0BCBF: call [00401048h] ; __vbaFreeObjList
  loc_00E0BCC5: add esp, 0000000Ch
  loc_00E0BCC8: lea ecx, var_58
  loc_00E0BCCB: call [00401024h] ; __vbaFreeVar
  loc_00E0BCD1: mov ecx, [esi]
  loc_00E0BCD3: push 006DCBD8h
  loc_00E0BCD8: push 00000000h
  loc_00E0BCDA: push 00000018h
  loc_00E0BCDC: push esi
  loc_00E0BCDD: call [ecx+00000490h]
  loc_00E0BCE3: lea edx, var_3C
  loc_00E0BCE6: push eax
  loc_00E0BCE7: push edx
  loc_00E0BCE8: call ebx
  loc_00E0BCEA: push eax
  loc_00E0BCEB: lea eax, var_58
  loc_00E0BCEE: push eax
  loc_00E0BCEF: call [0040112Ch] ; __vbaLateIdCallLd
  loc_00E0BCF5: add esp, 00000010h
  loc_00E0BCF8: push eax
  loc_00E0BCF9: call [00401120h] ; __vbaCastObjVar
  loc_00E0BCFF: lea ecx, var_40
  loc_00E0BD02: push eax
  loc_00E0BD03: push ecx
  loc_00E0BD04: call ebx
  loc_00E0BD06: mov edx, [eax]
  loc_00E0BD08: lea ecx, var_44
  loc_00E0BD0B: push ecx
  loc_00E0BD0C: push eax
  loc_00E0BD0D: mov var_E8, eax
  loc_00E0BD13: call [edx+00000054h]
  loc_00E0BD16: test eax, eax
  loc_00E0BD18: fnclex
  loc_00E0BD1A: jge 00E0BD31h
  loc_00E0BD1C: mov edx, var_E8
  loc_00E0BD22: push 00000054h
  loc_00E0BD24: push 006DCBE8h
  loc_00E0BD29: push edx
  loc_00E0BD2A: push eax
  loc_00E0BD2B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0BD31: lea edx, var_48
  loc_00E0BD34: mov eax, 00000002h
  loc_00E0BD39: push edx
  loc_00E0BD3A: mov ecx, var_44
  loc_00E0BD3D: sub esp, 00000010h
  loc_00E0BD40: mov var_98, eax
  loc_00E0BD46: mov edx, esp
  loc_00E0BD48: mov var_90, 00000000h
  loc_00E0BD52: mov var_F0, ecx
  loc_00E0BD58: mov ecx, [ecx]
  loc_00E0BD5A: mov [edx], eax
  loc_00E0BD5C: mov eax, var_94
  loc_00E0BD62: mov [edx+00000004h], eax
  loc_00E0BD65: mov eax, var_90
  loc_00E0BD6B: mov [edx+00000008h], eax
  loc_00E0BD6E: mov eax, var_8C
  loc_00E0BD74: mov [edx+0000000Ch], eax
  loc_00E0BD77: mov edx, var_44
  loc_00E0BD7A: push edx
  loc_00E0BD7B: call [ecx+00000028h]
  loc_00E0BD7E: test eax, eax
  loc_00E0BD80: fnclex
  loc_00E0BD82: jge 00E0BD99h
  loc_00E0BD84: mov ecx, var_F0
  loc_00E0BD8A: push 00000028h
  loc_00E0BD8C: push 006E09E8h
  loc_00E0BD91: push ecx
  loc_00E0BD92: push eax
  loc_00E0BD93: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0BD99: mov edx, var_48
  loc_00E0BD9C: mov eax, [esi]
  loc_00E0BD9E: push esi
  loc_00E0BD9F: mov var_F8, edx
  loc_00E0BDA5: call [eax+0000039Ch]
  loc_00E0BDAB: lea ecx, var_38
  loc_00E0BDAE: push eax
  loc_00E0BDAF: push ecx
  loc_00E0BDB0: call ebx
  loc_00E0BDB2: mov edx, [eax]
  loc_00E0BDB4: lea ecx, var_2C
  loc_00E0BDB7: push ecx
  loc_00E0BDB8: push eax
  loc_00E0BDB9: mov var_E0, eax
  loc_00E0BDBF: call [edx+000000A0h]
  loc_00E0BDC5: test eax, eax
  loc_00E0BDC7: fnclex
  loc_00E0BDC9: jge 00E0BDE3h
  loc_00E0BDCB: mov edx, var_E0
  loc_00E0BDD1: push 000000A0h
  loc_00E0BDD6: push 006DCB70h
  loc_00E0BDDB: push edx
  loc_00E0BDDC: push eax
  loc_00E0BDDD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0BDE3: mov eax, var_2C
  loc_00E0BDE6: mov ecx, var_F8
  loc_00E0BDEC: mov var_60, eax
  loc_00E0BDEF: mov eax, 00000008h
  loc_00E0BDF4: mov var_2C, 00000000h
  loc_00E0BDFB: mov var_68, eax
  loc_00E0BDFE: mov edx, [ecx]
  loc_00E0BE00: sub esp, 00000010h
  loc_00E0BE03: mov ecx, esp
  loc_00E0BE05: mov [ecx], eax
  loc_00E0BE07: mov eax, var_64
  loc_00E0BE0A: mov [ecx+00000004h], eax
  loc_00E0BE0D: mov eax, var_60
  loc_00E0BE10: mov [ecx+00000008h], eax
  loc_00E0BE13: mov eax, var_5C
  loc_00E0BE16: mov [ecx+0000000Ch], eax
  loc_00E0BE19: mov ecx, var_F8
  loc_00E0BE1F: push ecx
  loc_00E0BE20: call [edx+00000038h]
  loc_00E0BE23: test eax, eax
  loc_00E0BE25: fnclex
  loc_00E0BE27: jge 00E0BE3Eh
  loc_00E0BE29: mov edx, var_F8
  loc_00E0BE2F: push 00000038h
  loc_00E0BE31: push 006E09F8h
  loc_00E0BE36: push edx
  loc_00E0BE37: push eax
  loc_00E0BE38: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0BE3E: lea eax, var_48
  loc_00E0BE41: lea ecx, var_44
  loc_00E0BE44: push eax
  loc_00E0BE45: lea edx, var_40
  loc_00E0BE48: push ecx
  loc_00E0BE49: lea eax, var_3C
  loc_00E0BE4C: push edx
  loc_00E0BE4D: lea ecx, var_38
  loc_00E0BE50: push eax
  loc_00E0BE51: push ecx
  loc_00E0BE52: push 00000005h
  loc_00E0BE54: call [00401048h] ; __vbaFreeObjList
  loc_00E0BE5A: lea edx, var_68
  loc_00E0BE5D: lea eax, var_58
  loc_00E0BE60: push edx
  loc_00E0BE61: push eax
  loc_00E0BE62: push 00000002h
  loc_00E0BE64: call [00401038h] ; __vbaFreeVarList
  loc_00E0BE6A: add esp, 00000014h
  loc_00E0BE6D: mov ecx, 0000000Bh
  loc_00E0BE72: mov edx, esp
  loc_00E0BE74: mov var_98, ecx
  loc_00E0BE7A: or eax, FFFFFFFFh
  loc_00E0BE7D: push 8001000Dh
  loc_00E0BE82: mov [edx], ecx
  loc_00E0BE84: mov ecx, var_94
  loc_00E0BE8A: mov var_90, eax
  loc_00E0BE90: push esi
  loc_00E0BE91: mov [edx+00000004h], ecx
  loc_00E0BE94: mov ecx, [esi]
  loc_00E0BE96: mov [edx+00000008h], eax
  loc_00E0BE99: mov eax, var_8C
  loc_00E0BE9F: mov [edx+0000000Ch], eax
  loc_00E0BEA2: call [ecx+000003A0h]
  loc_00E0BEA8: lea edx, var_38
  loc_00E0BEAB: push eax
  loc_00E0BEAC: push edx
  loc_00E0BEAD: call ebx
  loc_00E0BEAF: push eax
  loc_00E0BEB0: call [00401238h] ; __vbaLateIdSt
  loc_00E0BEB6: lea ecx, var_38
  loc_00E0BEB9: call edi
  loc_00E0BEBB: sub esp, 00000010h
  loc_00E0BEBE: mov ecx, 0000000Bh
  loc_00E0BEC3: mov edx, esp
  loc_00E0BEC5: mov var_98, ecx
  loc_00E0BECB: or eax, FFFFFFFFh
  loc_00E0BECE: push 8001000Dh
  loc_00E0BED3: mov [edx], ecx
  loc_00E0BED5: mov ecx, var_94
  loc_00E0BEDB: mov var_90, eax
  loc_00E0BEE1: push esi
  loc_00E0BEE2: mov [edx+00000004h], ecx
  loc_00E0BEE5: mov ecx, [esi]
  loc_00E0BEE7: mov [edx+00000008h], eax
  loc_00E0BEEA: mov eax, var_8C
  loc_00E0BEF0: mov [edx+0000000Ch], eax
  loc_00E0BEF3: call [ecx+000003A4h]
  loc_00E0BEF9: lea edx, var_38
  loc_00E0BEFC: push eax
  loc_00E0BEFD: push edx
  loc_00E0BEFE: call ebx
  loc_00E0BF00: push eax
  loc_00E0BF01: call [00401238h] ; __vbaLateIdSt
  loc_00E0BF07: lea ecx, var_38
  loc_00E0BF0A: call edi
  loc_00E0BF0C: or eax, FFFFFFFFh
  loc_00E0BF0F: sub esp, 00000010h
  loc_00E0BF12: mov edx, esp
  loc_00E0BF14: mov ecx, 0000000Bh
  loc_00E0BF19: mov var_98, ecx
  loc_00E0BF1F: mov var_90, eax
  loc_00E0BF25: mov [edx], ecx
  loc_00E0BF27: mov ecx, var_94
  loc_00E0BF2D: push 8001000Dh
  loc_00E0BF32: push esi
  loc_00E0BF33: mov [edx+00000004h], ecx
  loc_00E0BF36: mov ecx, [esi]
  loc_00E0BF38: mov [edx+00000008h], eax
  loc_00E0BF3B: mov eax, var_8C
  loc_00E0BF41: mov [edx+0000000Ch], eax
  loc_00E0BF44: call [ecx+000003A8h]
  loc_00E0BF4A: lea edx, var_38
  loc_00E0BF4D: push eax
  loc_00E0BF4E: push edx
  loc_00E0BF4F: call ebx
  loc_00E0BF51: push eax
  loc_00E0BF52: call [00401238h] ; __vbaLateIdSt
  loc_00E0BF58: lea ecx, var_38
  loc_00E0BF5B: call edi
  loc_00E0BF5D: or eax, FFFFFFFFh
  loc_00E0BF60: jmp 00E0C0ECh
  loc_00E0BF65: mov ecx, 80020004h
  loc_00E0BF6A: mov eax, 0000000Ah
  loc_00E0BF6F: mov var_80, ecx
  loc_00E0BF72: mov var_70, ecx
  loc_00E0BF75: lea edx, var_A8
  loc_00E0BF7B: lea ecx, var_68
  loc_00E0BF7E: mov var_88, eax
  loc_00E0BF84: mov var_78, eax
  loc_00E0BF87: mov var_A0, 006E0A60h ; "Attention"
  loc_00E0BF91: mov var_A8, 00000008h
  loc_00E0BF9B: call [004011F0h] ; __vbaVarDup
  loc_00E0BFA1: lea edx, var_98
  loc_00E0BFA7: lea ecx, var_58
  loc_00E0BFAA: mov var_90, 006E0A0Ch ; "Nomor Daftar Sudah Di input sebelumnya"
  loc_00E0BFB4: mov var_98, 00000008h
  loc_00E0BFBE: call [004011F0h] ; __vbaVarDup
  loc_00E0BFC4: lea eax, var_88
  loc_00E0BFCA: lea ecx, var_78
  loc_00E0BFCD: push eax
  loc_00E0BFCE: lea edx, var_68
  loc_00E0BFD1: push ecx
  loc_00E0BFD2: push edx
  loc_00E0BFD3: lea eax, var_58
  loc_00E0BFD6: push 00000010h
  loc_00E0BFD8: push eax
  loc_00E0BFD9: call [004010A8h] ; rtcMsgBox
  loc_00E0BFDF: lea ecx, var_88
  loc_00E0BFE5: lea edx, var_78
  loc_00E0BFE8: push ecx
  loc_00E0BFE9: lea eax, var_68
  loc_00E0BFEC: push edx
  loc_00E0BFED: lea ecx, var_58
  loc_00E0BFF0: push eax
  loc_00E0BFF1: push ecx
  loc_00E0BFF2: push 00000004h
  loc_00E0BFF4: call [00401038h] ; __vbaFreeVarList
  loc_00E0BFFA: add esp, 00000004h
  loc_00E0BFFD: mov ecx, 0000000Bh
  loc_00E0C002: mov edx, esp
  loc_00E0C004: mov var_98, ecx
  loc_00E0C00A: xor eax, eax
  loc_00E0C00C: push 8001000Dh
  loc_00E0C011: mov [edx], ecx
  loc_00E0C013: mov ecx, var_94
  loc_00E0C019: mov var_90, eax
  loc_00E0C01F: push esi
  loc_00E0C020: mov [edx+00000004h], ecx
  loc_00E0C023: mov ecx, [esi]
  loc_00E0C025: mov [edx+00000008h], eax
  loc_00E0C028: mov eax, var_8C
  loc_00E0C02E: mov [edx+0000000Ch], eax
  loc_00E0C031: call [ecx+000003A0h]
  loc_00E0C037: lea edx, var_38
  loc_00E0C03A: push eax
  loc_00E0C03B: push edx
  loc_00E0C03C: call ebx
  loc_00E0C03E: push eax
  loc_00E0C03F: call [00401238h] ; __vbaLateIdSt
  loc_00E0C045: lea ecx, var_38
  loc_00E0C048: call edi
  loc_00E0C04A: sub esp, 00000010h
  loc_00E0C04D: mov ecx, 0000000Bh
  loc_00E0C052: mov edx, esp
  loc_00E0C054: mov var_98, ecx
  loc_00E0C05A: xor eax, eax
  loc_00E0C05C: push 8001000Dh
  loc_00E0C061: mov [edx], ecx
  loc_00E0C063: mov ecx, var_94
  loc_00E0C069: mov var_90, eax
  loc_00E0C06F: push esi
  loc_00E0C070: mov [edx+00000004h], ecx
  loc_00E0C073: mov ecx, [esi]
  loc_00E0C075: mov [edx+00000008h], eax
  loc_00E0C078: mov eax, var_8C
  loc_00E0C07E: mov [edx+0000000Ch], eax
  loc_00E0C081: call [ecx+000003A4h]
  loc_00E0C087: lea edx, var_38
  loc_00E0C08A: push eax
  loc_00E0C08B: push edx
  loc_00E0C08C: call ebx
  loc_00E0C08E: push eax
  loc_00E0C08F: call [00401238h] ; __vbaLateIdSt
  loc_00E0C095: lea ecx, var_38
  loc_00E0C098: call edi
  loc_00E0C09A: sub esp, 00000010h
  loc_00E0C09D: mov ecx, 0000000Bh
  loc_00E0C0A2: mov edx, esp
  loc_00E0C0A4: mov var_98, ecx
  loc_00E0C0AA: xor eax, eax
  loc_00E0C0AC: push 8001000Dh
  loc_00E0C0B1: mov [edx], ecx
  loc_00E0C0B3: mov ecx, var_94
  loc_00E0C0B9: mov var_90, eax
  loc_00E0C0BF: push esi
  loc_00E0C0C0: mov [edx+00000004h], ecx
  loc_00E0C0C3: mov ecx, [esi]
  loc_00E0C0C5: mov [edx+00000008h], eax
  loc_00E0C0C8: mov eax, var_8C
  loc_00E0C0CE: mov [edx+0000000Ch], eax
  loc_00E0C0D1: call [ecx+000003A8h]
  loc_00E0C0D7: lea edx, var_38
  loc_00E0C0DA: push eax
  loc_00E0C0DB: push edx
  loc_00E0C0DC: call ebx
  loc_00E0C0DE: push eax
  loc_00E0C0DF: call [00401238h] ; __vbaLateIdSt
  loc_00E0C0E5: lea ecx, var_38
  loc_00E0C0E8: call edi
  loc_00E0C0EA: xor eax, eax
  loc_00E0C0EC: sub esp, 00000010h
  loc_00E0C0EF: mov ecx, 0000000Bh
  loc_00E0C0F4: mov edx, esp
  loc_00E0C0F6: mov var_98, ecx
  loc_00E0C0FC: mov var_90, eax
  loc_00E0C102: push 8001000Dh
  loc_00E0C107: mov [edx], ecx
  loc_00E0C109: mov ecx, var_94
  loc_00E0C10F: push esi
  loc_00E0C110: mov [edx+00000004h], ecx
  loc_00E0C113: mov ecx, [esi]
  loc_00E0C115: mov [edx+00000008h], eax
  loc_00E0C118: mov eax, var_8C
  loc_00E0C11E: mov [edx+0000000Ch], eax
  loc_00E0C121: call [ecx+000003ACh]
  loc_00E0C127: lea edx, var_38
  loc_00E0C12A: push eax
  loc_00E0C12B: push edx
  loc_00E0C12C: call ebx
  loc_00E0C12E: push eax
  loc_00E0C12F: call [00401238h] ; __vbaLateIdSt
  loc_00E0C135: lea ecx, var_38
  loc_00E0C138: call edi
  loc_00E0C13A: mov var_4, 00000000h
  loc_00E0C141: push 00E0C1A1h
  loc_00E0C146: jmp 00E0C197h
  loc_00E0C148: lea eax, var_34
  loc_00E0C14B: lea ecx, var_30
  loc_00E0C14E: push eax
  loc_00E0C14F: lea edx, var_2C
  loc_00E0C152: push ecx
  loc_00E0C153: push edx
  loc_00E0C154: push 00000003h
  loc_00E0C156: call [004011BCh] ; __vbaFreeStrList
  loc_00E0C15C: lea eax, var_48
  loc_00E0C15F: lea ecx, var_44
  loc_00E0C162: push eax
  loc_00E0C163: lea edx, var_40
  loc_00E0C166: push ecx
  loc_00E0C167: lea eax, var_3C
  loc_00E0C16A: push edx
  loc_00E0C16B: lea ecx, var_38
  loc_00E0C16E: push eax
  loc_00E0C16F: push ecx
  loc_00E0C170: push 00000005h
  loc_00E0C172: call [00401048h] ; __vbaFreeObjList
  loc_00E0C178: lea edx, var_88
  loc_00E0C17E: lea eax, var_78
  loc_00E0C181: push edx
  loc_00E0C182: lea ecx, var_68
  loc_00E0C185: push eax
  loc_00E0C186: lea edx, var_58
  loc_00E0C189: push ecx
  loc_00E0C18A: push edx
  loc_00E0C18B: push 00000004h
  loc_00E0C18D: call [00401038h] ; __vbaFreeVarList
  loc_00E0C193: add esp, 0000003Ch
  loc_00E0C196: ret
  loc_00E0C197: lea ecx, var_24
  loc_00E0C19A: call [00401024h] ; __vbaFreeVar
  loc_00E0C1A0: ret
  loc_00E0C1A1: mov eax, Me
  loc_00E0C1A4: push eax
  loc_00E0C1A5: mov ecx, [eax]
  loc_00E0C1A7: call [ecx+00000008h]
  loc_00E0C1AA: mov eax, var_4
  loc_00E0C1AD: mov ecx, var_14
  loc_00E0C1B0: pop edi
  loc_00E0C1B1: pop esi
  loc_00E0C1B2: mov fs:[00000000h], ecx
  loc_00E0C1B9: pop ebx
  loc_00E0C1BA: mov esp, ebp
  loc_00E0C1BC: pop ebp
  loc_00E0C1BD: retn 0004h
End Sub

Private Sub Timer1_Timer() 'E0F7B0
  loc_00E0F7B0: push ebp
  loc_00E0F7B1: mov ebp, esp
  loc_00E0F7B3: sub esp, 0000000Ch
  loc_00E0F7B6: push 00402836h ; __vbaExceptHandler
  loc_00E0F7BB: mov eax, fs:[00000000h]
  loc_00E0F7C1: push eax
  loc_00E0F7C2: mov fs:[00000000h], esp
  loc_00E0F7C9: sub esp, 000000B0h
  loc_00E0F7CF: push ebx
  loc_00E0F7D0: push esi
  loc_00E0F7D1: push edi
  loc_00E0F7D2: mov var_C, esp
  loc_00E0F7D5: mov var_8, 00402280h
  loc_00E0F7DC: mov esi, Me
  loc_00E0F7DF: mov eax, esi
  loc_00E0F7E1: and eax, 00000001h
  loc_00E0F7E4: mov var_4, eax
  loc_00E0F7E7: and esi, FFFFFFFEh
  loc_00E0F7EA: push esi
  loc_00E0F7EB: mov Me, esi
  loc_00E0F7EE: mov ecx, [esi]
  loc_00E0F7F0: call [ecx+00000004h]
  loc_00E0F7F3: xor ebx, ebx
  loc_00E0F7F5: lea edi, [esi+00000038h]
  loc_00E0F7F8: mov var_70, ebx
  loc_00E0F7FB: lea edx, var_70
  loc_00E0F7FE: mov ecx, edi
  loc_00E0F800: mov var_18, ebx
  loc_00E0F803: mov var_1C, ebx
  loc_00E0F806: mov var_20, ebx
  loc_00E0F809: mov var_30, ebx
  loc_00E0F80C: mov var_40, ebx
  loc_00E0F80F: mov var_50, ebx
  loc_00E0F812: mov var_60, ebx
  loc_00E0F815: mov var_80, ebx
  loc_00E0F818: mov var_A4, ebx
  loc_00E0F81E: mov var_68, ebx
  loc_00E0F821: mov var_70, 00000002h
  loc_00E0F828: call [0040101Ch] ; __vbaVarMove
  loc_00E0F82E: lea eax, [esi+00000048h]
  loc_00E0F831: lea edx, var_70
  loc_00E0F834: push eax
  loc_00E0F835: push edx
  loc_00E0F836: mov var_68, 00000064h
  loc_00E0F83D: mov var_70, 00008002h
  loc_00E0F844: mov var_C0, eax
  loc_00E0F84A: call [00401088h] ; __vbaVarTstLe
  loc_00E0F850: test ax, ax
  loc_00E0F853: jz 00E0FBF9h
  loc_00E0F859: lea eax, var_30
  loc_00E0F85C: mov var_28, 80020004h
  loc_00E0F863: push eax
  loc_00E0F864: mov var_30, 0000000Ah
  loc_00E0F86B: call [004010A0h] ; rtcRandomize
  loc_00E0F871: mov ebx, [00401024h] ; __vbaFreeVar
  loc_00E0F877: lea ecx, var_30
  loc_00E0F87A: call ebx
  loc_00E0F87C: lea ecx, var_30
  loc_00E0F87F: mov var_28, 80020004h
  loc_00E0F886: push ecx
  loc_00E0F887: mov var_30, 0000000Ah
  loc_00E0F88E: call [0040109Ch] ; rtcRandomNext
  loc_00E0F894: fstp real4 ptr var_A4
  loc_00E0F89A: fld real4 ptr var_A4
  loc_00E0F8A0: fmul st0, real4 ptr [004012E4h]
  loc_00E0F8A6: lea ecx, var_30
  loc_00E0F8A9: fadd st0, real4 ptr [004012E0h]
  loc_00E0F8AF: fstp real8 ptr [esi+00000058h]
  loc_00E0F8B2: fnstsw ax
  loc_00E0F8B4: test al, 0Dh
  loc_00E0F8B6: jnz 00E0FC59h
  loc_00E0F8BC: call ebx
  loc_00E0F8BE: fld real8 ptr [esi+00000058h]
  loc_00E0F8C1: cmp [00E3D000h], 00000000h
  loc_00E0F8C8: jnz 00E0F8D2h
  loc_00E0F8CA: fdiv st0, real8 ptr [004012D8h]
  loc_00E0F8D0: jmp 00E0F8E3h
  loc_00E0F8D2: push [004012DCh]
  loc_00E0F8D8: push [004012D8h]
  loc_00E0F8DE: call 00402854h ; _adj_fdiv_m64
  loc_00E0F8E3: lea edx, var_70
  loc_00E0F8E6: mov ecx, edi
  loc_00E0F8E8: mov var_70, 00000005h
  loc_00E0F8EF: fstp real8 ptr var_68
  loc_00E0F8F2: fnstsw ax
  loc_00E0F8F4: test al, 0Dh
  loc_00E0F8F6: jnz 00E0FC59h
  loc_00E0F8FC: call [0040101Ch] ; __vbaVarMove
  loc_00E0F902: push 00000000h
  loc_00E0F904: lea edx, var_30
  loc_00E0F907: push edi
  loc_00E0F908: push edx
  loc_00E0F909: call [00401170h] ; rtcRound
  loc_00E0F90F: lea edx, var_30
  loc_00E0F912: mov ecx, edi
  loc_00E0F914: call [0040101Ch] ; __vbaVarMove
  loc_00E0F91A: lea ecx, var_30
  loc_00E0F91D: call ebx
  loc_00E0F91F: mov eax, var_C0
  loc_00E0F925: lea ecx, var_30
  loc_00E0F928: push eax
  loc_00E0F929: push edi
  loc_00E0F92A: push ecx
  loc_00E0F92B: call [004011E0h] ; __vbaVarAdd
  loc_00E0F931: mov edx, eax
  loc_00E0F933: mov ecx, edi
  loc_00E0F935: call [0040101Ch] ; __vbaVarMove
  loc_00E0F93B: lea ecx, var_30
  loc_00E0F93E: call ebx
  loc_00E0F940: lea edx, var_30
  loc_00E0F943: push edi
  loc_00E0F944: push edx
  loc_00E0F945: call [004011FCh] ; rtcVarStrFromVar
  loc_00E0F94B: lea ecx, var_30
  loc_00E0F94E: lea eax, [esi+00000060h]
  loc_00E0F951: push ecx
  loc_00E0F952: mov var_C4, eax
  loc_00E0F958: call [00401034h] ; __vbaStrVarMove
  loc_00E0F95E: mov edx, eax
  loc_00E0F960: lea ecx, var_18
  loc_00E0F963: call [00401228h] ; __vbaStrMove
  loc_00E0F969: mov ecx, var_C4
  loc_00E0F96F: mov edx, eax
  loc_00E0F971: call [004011B0h] ; __vbaStrCopy
  loc_00E0F977: lea ecx, var_18
  loc_00E0F97A: call [00401258h] ; __vbaFreeStr
  loc_00E0F980: lea ecx, var_30
  loc_00E0F983: call ebx
  loc_00E0F985: lea edx, var_70
  loc_00E0F988: push edi
  loc_00E0F989: push edx
  loc_00E0F98A: mov var_68, 00000064h
  loc_00E0F991: mov var_70, 00008002h
  loc_00E0F998: call [00401210h] ; __vbaVarTstGe
  loc_00E0F99E: test ax, ax
  loc_00E0F9A1: mov eax, [esi]
  loc_00E0F9A3: push esi
  loc_00E0F9A4: jz 00E0FB38h
  loc_00E0F9AA: call [eax+000003ECh]
  loc_00E0F9B0: lea ecx, var_1C
  loc_00E0F9B3: push eax
  loc_00E0F9B4: push ecx
  loc_00E0F9B5: call [004010ACh] ; __vbaObjSet
  loc_00E0F9BB: mov edi, eax
  loc_00E0F9BD: push 46653800h
  loc_00E0F9C2: push edi
  loc_00E0F9C3: mov edx, [edi]
  loc_00E0F9C5: call [edx+0000007Ch]
  loc_00E0F9C8: test eax, eax
  loc_00E0F9CA: fnclex
  loc_00E0F9CC: jge 00E0F9DDh
  loc_00E0F9CE: push 0000007Ch
  loc_00E0F9D0: push 006DCDA0h
  loc_00E0F9D5: push edi
  loc_00E0F9D6: push eax
  loc_00E0F9D7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0F9DD: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E0F9E3: lea ecx, var_1C
  loc_00E0F9E6: call ebx
  loc_00E0F9E8: mov eax, [esi]
  loc_00E0F9EA: push esi
  loc_00E0F9EB: call [eax+000003ECh]
  loc_00E0F9F1: lea ecx, var_1C
  loc_00E0F9F4: push eax
  loc_00E0F9F5: push ecx
  loc_00E0F9F6: call [004010ACh] ; __vbaObjSet
  loc_00E0F9FC: mov edi, eax
  loc_00E0F9FE: lea eax, var_A4
  loc_00E0FA04: push eax
  loc_00E0FA05: push edi
  loc_00E0FA06: mov edx, [edi]
  loc_00E0FA08: call [edx+00000078h]
  loc_00E0FA0B: test eax, eax
  loc_00E0FA0D: fnclex
  loc_00E0FA0F: jge 00E0FA20h
  loc_00E0FA11: push 00000078h
  loc_00E0FA13: push 006DCDA0h
  loc_00E0FA18: push edi
  loc_00E0FA19: push eax
  loc_00E0FA1A: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FA20: cmp var_A4, 46653800h
  loc_00E0FA2A: jnz 00E0FA33h
  loc_00E0FA2C: mov eax, 00000001h
  loc_00E0FA31: jmp 00E0FA35h
  loc_00E0FA33: xor eax, eax
  loc_00E0FA35: neg eax
  loc_00E0FA37: lea ecx, var_1C
  loc_00E0FA3A: mov di, ax
  loc_00E0FA3D: call ebx
  loc_00E0FA3F: test di, di
  loc_00E0FA42: jz 00E0FBF7h
  loc_00E0FA48: mov ecx, [esi]
  loc_00E0FA4A: push esi
  loc_00E0FA4B: call [ecx+000003ECh]
  loc_00E0FA51: lea edx, var_1C
  loc_00E0FA54: push eax
  loc_00E0FA55: push edx
  loc_00E0FA56: call [004010ACh] ; __vbaObjSet
  loc_00E0FA5C: mov edi, eax
  loc_00E0FA5E: push 3F800000h
  loc_00E0FA63: push edi
  loc_00E0FA64: mov eax, [edi]
  loc_00E0FA66: call [eax+0000007Ch]
  loc_00E0FA69: test eax, eax
  loc_00E0FA6B: fnclex
  loc_00E0FA6D: jge 00E0FA7Eh
  loc_00E0FA6F: push 0000007Ch
  loc_00E0FA71: push 006DCDA0h
  loc_00E0FA76: push edi
  loc_00E0FA77: push eax
  loc_00E0FA78: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FA7E: lea ecx, var_1C
  loc_00E0FA81: call ebx
  loc_00E0FA83: mov ecx, [esi]
  loc_00E0FA85: push esi
  loc_00E0FA86: call [ecx+00000334h]
  loc_00E0FA8C: lea edx, var_1C
  loc_00E0FA8F: push eax
  loc_00E0FA90: push edx
  loc_00E0FA91: call [004010ACh] ; __vbaObjSet
  loc_00E0FA97: mov esi, eax
  loc_00E0FA99: push 00000000h
  loc_00E0FA9B: push esi
  loc_00E0FA9C: mov eax, [esi]
  loc_00E0FA9E: call [eax+0000005Ch]
  loc_00E0FAA1: test eax, eax
  loc_00E0FAA3: fnclex
  loc_00E0FAA5: jge 00E0FAB6h
  loc_00E0FAA7: push 0000005Ch
  loc_00E0FAA9: push 006DCAE0h
  loc_00E0FAAE: push esi
  loc_00E0FAAF: push eax
  loc_00E0FAB0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FAB6: lea ecx, var_1C
  loc_00E0FAB9: call ebx
  loc_00E0FABB: mov esi, [004011F0h] ; __vbaVarDup
  loc_00E0FAC1: mov ecx, 80020004h
  loc_00E0FAC6: mov var_58, ecx
  loc_00E0FAC9: mov eax, 0000000Ah
  loc_00E0FACE: mov var_48, ecx
  loc_00E0FAD1: mov edi, 00000008h
  loc_00E0FAD6: lea edx, var_80
  loc_00E0FAD9: lea ecx, var_40
  loc_00E0FADC: mov var_60, eax
  loc_00E0FADF: mov var_50, eax
  loc_00E0FAE2: mov var_78, 006E0C30h ; "INFO 001"
  loc_00E0FAE9: mov var_80, edi
  loc_00E0FAEC: call __vbaVarDup
  loc_00E0FAEE: lea edx, var_70
  loc_00E0FAF1: lea ecx, var_30
  loc_00E0FAF4: mov var_68, 006E0BFCh ; "Data Telah Tersimpan !"
  loc_00E0FAFB: mov var_70, edi
  loc_00E0FAFE: call __vbaVarDup
  loc_00E0FB00: lea ecx, var_60
  loc_00E0FB03: lea edx, var_50
  loc_00E0FB06: push ecx
  loc_00E0FB07: lea eax, var_40
  loc_00E0FB0A: push edx
  loc_00E0FB0B: push eax
  loc_00E0FB0C: lea ecx, var_30
  loc_00E0FB0F: push 00000040h
  loc_00E0FB11: push ecx
  loc_00E0FB12: call [004010A8h] ; rtcMsgBox
  loc_00E0FB18: lea edx, var_60
  loc_00E0FB1B: lea eax, var_50
  loc_00E0FB1E: push edx
  loc_00E0FB1F: lea ecx, var_40
  loc_00E0FB22: push eax
  loc_00E0FB23: lea edx, var_30
  loc_00E0FB26: push ecx
  loc_00E0FB27: push edx
  loc_00E0FB28: push 00000004h
  loc_00E0FB2A: call [00401038h] ; __vbaFreeVarList
  loc_00E0FB30: add esp, 00000014h
  loc_00E0FB33: jmp 00E0FBF7h
  loc_00E0FB38: call [eax+000003ECh]
  loc_00E0FB3E: mov ebx, [004010ACh] ; __vbaObjSet
  loc_00E0FB44: lea ecx, var_20
  loc_00E0FB47: push eax
  loc_00E0FB48: push ecx
  loc_00E0FB49: call ebx
  loc_00E0FB4B: mov edx, [esi]
  loc_00E0FB4D: push esi
  loc_00E0FB4E: mov edi, eax
  loc_00E0FB50: call [edx+000003ECh]
  loc_00E0FB56: push eax
  loc_00E0FB57: lea eax, var_1C
  loc_00E0FB5A: push eax
  loc_00E0FB5B: call ebx
  loc_00E0FB5D: mov esi, eax
  loc_00E0FB5F: lea edx, var_A4
  loc_00E0FB65: push edx
  loc_00E0FB66: push esi
  loc_00E0FB67: mov ecx, [esi]
  loc_00E0FB69: call [ecx+00000078h]
  loc_00E0FB6C: test eax, eax
  loc_00E0FB6E: fnclex
  loc_00E0FB70: jge 00E0FB81h
  loc_00E0FB72: push 00000078h
  loc_00E0FB74: push 006DCDA0h
  loc_00E0FB79: push esi
  loc_00E0FB7A: push eax
  loc_00E0FB7B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FB81: fld real4 ptr var_A4
  loc_00E0FB87: fadd st0, real4 ptr [004012D0h]
  loc_00E0FB8D: mov ecx, [edi]
  loc_00E0FB8F: push ecx
  loc_00E0FB90: fnstsw ax
  loc_00E0FB92: test al, 0Dh
  loc_00E0FB94: jnz 00E0FC59h
  loc_00E0FB9A: fstp real4 ptr [esp]
  loc_00E0FB9D: push edi
  loc_00E0FB9E: call [ecx+0000007Ch]
  loc_00E0FBA1: test eax, eax
  loc_00E0FBA3: fnclex
  loc_00E0FBA5: jge 00E0FBB6h
  loc_00E0FBA7: push 0000007Ch
  loc_00E0FBA9: push 006DCDA0h
  loc_00E0FBAE: push edi
  loc_00E0FBAF: push eax
  loc_00E0FBB0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FBB6: lea edx, var_20
  loc_00E0FBB9: lea eax, var_1C
  loc_00E0FBBC: push edx
  loc_00E0FBBD: push eax
  loc_00E0FBBE: push 00000002h
  loc_00E0FBC0: call [00401048h] ; __vbaFreeObjList
  loc_00E0FBC6: mov ecx, var_C4
  loc_00E0FBCC: add esp, 0000000Ch
  loc_00E0FBCF: mov edx, [ecx]
  loc_00E0FBD1: push edx
  loc_00E0FBD2: call [004011A4h] ; __vbaR8Str
  loc_00E0FBD8: call [0040124Ch] ; __vbaFPInt
  loc_00E0FBDE: mov ecx, var_C0
  loc_00E0FBE4: lea edx, var_70
  loc_00E0FBE7: fstp real8 ptr var_68
  loc_00E0FBEA: mov var_70, 00000005h
  loc_00E0FBF1: call [0040101Ch] ; __vbaVarMove
  loc_00E0FBF7: xor ebx, ebx
  loc_00E0FBF9: mov var_4, ebx
  loc_00E0FBFC: fwait
  loc_00E0FBFD: push 00E0FC3Ah
  loc_00E0FC02: jmp 00E0FC39h
  loc_00E0FC04: lea ecx, var_18
  loc_00E0FC07: call [00401258h] ; __vbaFreeStr
  loc_00E0FC0D: lea eax, var_20
  loc_00E0FC10: lea ecx, var_1C
  loc_00E0FC13: push eax
  loc_00E0FC14: push ecx
  loc_00E0FC15: push 00000002h
  loc_00E0FC17: call [00401048h] ; __vbaFreeObjList
  loc_00E0FC1D: lea edx, var_60
  loc_00E0FC20: lea eax, var_50
  loc_00E0FC23: push edx
  loc_00E0FC24: lea ecx, var_40
  loc_00E0FC27: push eax
  loc_00E0FC28: lea edx, var_30
  loc_00E0FC2B: push ecx
  loc_00E0FC2C: push edx
  loc_00E0FC2D: push 00000004h
  loc_00E0FC2F: call [00401038h] ; __vbaFreeVarList
  loc_00E0FC35: add esp, 00000020h
  loc_00E0FC38: ret
  loc_00E0FC39: ret
  loc_00E0FC3A: mov eax, Me
  loc_00E0FC3D: push eax
  loc_00E0FC3E: mov ecx, [eax]
  loc_00E0FC40: call [ecx+00000008h]
  loc_00E0FC43: mov eax, var_4
  loc_00E0FC46: mov ecx, var_14
  loc_00E0FC49: pop edi
  loc_00E0FC4A: pop esi
  loc_00E0FC4B: mov fs:[00000000h], ecx
  loc_00E0FC52: pop ebx
  loc_00E0FC53: mov esp, ebp
  loc_00E0FC55: pop ebp
  loc_00E0FC56: retn 0004h
End Sub

Private Sub Timer2_Timer() 'E0FC60
  loc_00E0FC60: push ebp
  loc_00E0FC61: mov ebp, esp
  loc_00E0FC63: sub esp, 0000000Ch
  loc_00E0FC66: push 00402836h ; __vbaExceptHandler
  loc_00E0FC6B: mov eax, fs:[00000000h]
  loc_00E0FC71: push eax
  loc_00E0FC72: mov fs:[00000000h], esp
  loc_00E0FC79: sub esp, 0000002Ch
  loc_00E0FC7C: push ebx
  loc_00E0FC7D: push esi
  loc_00E0FC7E: push edi
  loc_00E0FC7F: mov var_C, esp
  loc_00E0FC82: mov var_8, 00402298h
  loc_00E0FC89: mov esi, Me
  loc_00E0FC8C: mov eax, esi
  loc_00E0FC8E: and eax, 00000001h
  loc_00E0FC91: mov var_4, eax
  loc_00E0FC94: and esi, FFFFFFFEh
  loc_00E0FC97: push esi
  loc_00E0FC98: mov Me, esi
  loc_00E0FC9B: mov ecx, [esi]
  loc_00E0FC9D: call [ecx+00000004h]
  loc_00E0FCA0: mov edx, [esi]
  loc_00E0FCA2: xor ebx, ebx
  loc_00E0FCA4: push esi
  loc_00E0FCA5: mov var_18, ebx
  loc_00E0FCA8: mov var_1C, ebx
  loc_00E0FCAB: mov var_20, ebx
  loc_00E0FCAE: mov var_24, ebx
  loc_00E0FCB1: call [edx+0000032Ch]
  loc_00E0FCB7: push eax
  loc_00E0FCB8: lea eax, var_1C
  loc_00E0FCBB: push eax
  loc_00E0FCBC: call [004010ACh] ; __vbaObjSet
  loc_00E0FCC2: mov edi, eax
  loc_00E0FCC4: lea edx, var_24
  loc_00E0FCC7: push edx
  loc_00E0FCC8: push edi
  loc_00E0FCC9: mov ecx, [edi]
  loc_00E0FCCB: call [ecx+00000080h]
  loc_00E0FCD1: cmp eax, ebx
  loc_00E0FCD3: fnclex
  loc_00E0FCD5: jge 00E0FCEDh
  loc_00E0FCD7: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FCDD: push 00000080h
  loc_00E0FCE2: push 006E0410h
  loc_00E0FCE7: push edi
  loc_00E0FCE8: push eax
  loc_00E0FCE9: call ebx
  loc_00E0FCEB: jmp 00E0FCF3h
  loc_00E0FCED: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FCF3: mov eax, [esi]
  loc_00E0FCF5: push esi
  loc_00E0FCF6: call [eax+0000032Ch]
  loc_00E0FCFC: lea ecx, var_18
  loc_00E0FCFF: push eax
  loc_00E0FD00: push ecx
  loc_00E0FD01: call [004010ACh] ; __vbaObjSet
  loc_00E0FD07: mov edi, eax
  loc_00E0FD09: lea eax, var_20
  loc_00E0FD0C: push eax
  loc_00E0FD0D: push edi
  loc_00E0FD0E: mov edx, [edi]
  loc_00E0FD10: call [edx+00000070h]
  loc_00E0FD13: test eax, eax
  loc_00E0FD15: fnclex
  loc_00E0FD17: jge 00E0FD24h
  loc_00E0FD19: push 00000070h
  loc_00E0FD1B: push 006E0410h
  loc_00E0FD20: push edi
  loc_00E0FD21: push eax
  loc_00E0FD22: call ebx
  loc_00E0FD24: fld real4 ptr var_24
  loc_00E0FD27: fadd st0, real4 ptr var_20
  loc_00E0FD2A: fnstsw ax
  loc_00E0FD2C: test al, 0Dh
  loc_00E0FD2E: jnz 00E0FE8Eh
  loc_00E0FD34: call [004010C0h] ; __vbaFpR4
  loc_00E0FD3A: fcomp real4 ptr [00401888h]
  loc_00E0FD40: fnstsw ax
  loc_00E0FD42: test ah, 41h
  loc_00E0FD45: jz 00E0FD4Eh
  loc_00E0FD47: mov eax, 00000001h
  loc_00E0FD4C: jmp 00E0FD50h
  loc_00E0FD4E: xor eax, eax
  loc_00E0FD50: lea ecx, var_1C
  loc_00E0FD53: lea edx, var_18
  loc_00E0FD56: push ecx
  loc_00E0FD57: push edx
  loc_00E0FD58: neg eax
  loc_00E0FD5A: push 00000002h
  loc_00E0FD5C: mov edi, eax
  loc_00E0FD5E: call [00401048h] ; __vbaFreeObjList
  loc_00E0FD64: add esp, 0000000Ch
  loc_00E0FD67: test di, di
  loc_00E0FD6A: jz 00E0FDC7h
  loc_00E0FD6C: mov eax, [esi]
  loc_00E0FD6E: push esi
  loc_00E0FD6F: call [eax+0000032Ch]
  loc_00E0FD75: lea ecx, var_18
  loc_00E0FD78: push eax
  loc_00E0FD79: push ecx
  loc_00E0FD7A: call [004010ACh] ; __vbaObjSet
  loc_00E0FD80: mov edx, [esi]
  loc_00E0FD82: mov edi, eax
  loc_00E0FD84: lea eax, var_20
  loc_00E0FD87: push eax
  loc_00E0FD88: push esi
  loc_00E0FD89: call [edx+00000080h]
  loc_00E0FD8F: test eax, eax
  loc_00E0FD91: fnclex
  loc_00E0FD93: jge 00E0FDA3h
  loc_00E0FD95: push 00000080h
  loc_00E0FD9A: push 006DFCB8h
  loc_00E0FD9F: push esi
  loc_00E0FDA0: push eax
  loc_00E0FDA1: call ebx
  loc_00E0FDA3: mov edx, var_20
  loc_00E0FDA6: mov ecx, [edi]
  loc_00E0FDA8: push edx
  loc_00E0FDA9: push edi
  loc_00E0FDAA: call [ecx+00000074h]
  loc_00E0FDAD: test eax, eax
  loc_00E0FDAF: fnclex
  loc_00E0FDB1: jge 00E0FDBEh
  loc_00E0FDB3: push 00000074h
  loc_00E0FDB5: push 006E0410h
  loc_00E0FDBA: push edi
  loc_00E0FDBB: push eax
  loc_00E0FDBC: call ebx
  loc_00E0FDBE: lea ecx, var_18
  loc_00E0FDC1: call [00401254h] ; __vbaFreeObj
  loc_00E0FDC7: mov eax, [esi]
  loc_00E0FDC9: push esi
  loc_00E0FDCA: call [eax+0000032Ch]
  loc_00E0FDD0: lea ecx, var_1C
  loc_00E0FDD3: push eax
  loc_00E0FDD4: push ecx
  loc_00E0FDD5: call [004010ACh] ; __vbaObjSet
  loc_00E0FDDB: mov edx, [esi]
  loc_00E0FDDD: push esi
  loc_00E0FDDE: mov edi, eax
  loc_00E0FDE0: call [edx+0000032Ch]
  loc_00E0FDE6: push eax
  loc_00E0FDE7: lea eax, var_18
  loc_00E0FDEA: push eax
  loc_00E0FDEB: call [004010ACh] ; __vbaObjSet
  loc_00E0FDF1: mov esi, eax
  loc_00E0FDF3: lea edx, var_20
  loc_00E0FDF6: push edx
  loc_00E0FDF7: push esi
  loc_00E0FDF8: mov ecx, [esi]
  loc_00E0FDFA: call [ecx+00000070h]
  loc_00E0FDFD: test eax, eax
  loc_00E0FDFF: fnclex
  loc_00E0FE01: jge 00E0FE0Eh
  loc_00E0FE03: push 00000070h
  loc_00E0FE05: push 006E0410h
  loc_00E0FE0A: push esi
  loc_00E0FE0B: push eax
  loc_00E0FE0C: call ebx
  loc_00E0FE0E: fld real4 ptr var_20
  loc_00E0FE11: fsub st0, real4 ptr [00402290h]
  loc_00E0FE17: mov ecx, [edi]
  loc_00E0FE19: push ecx
  loc_00E0FE1A: fnstsw ax
  loc_00E0FE1C: test al, 0Dh
  loc_00E0FE1E: jnz 00E0FE8Eh
  loc_00E0FE20: fstp real4 ptr [esp]
  loc_00E0FE23: push edi
  loc_00E0FE24: call [ecx+00000074h]
  loc_00E0FE27: test eax, eax
  loc_00E0FE29: fnclex
  loc_00E0FE2B: jge 00E0FE38h
  loc_00E0FE2D: push 00000074h
  loc_00E0FE2F: push 006E0410h
  loc_00E0FE34: push edi
  loc_00E0FE35: push eax
  loc_00E0FE36: call ebx
  loc_00E0FE38: lea edx, var_1C
  loc_00E0FE3B: lea eax, var_18
  loc_00E0FE3E: push edx
  loc_00E0FE3F: push eax
  loc_00E0FE40: push 00000002h
  loc_00E0FE42: call [00401048h] ; __vbaFreeObjList
  loc_00E0FE48: add esp, 0000000Ch
  loc_00E0FE4B: mov var_4, 00000000h
  loc_00E0FE52: fwait
  loc_00E0FE53: push 00E0FE6Fh
  loc_00E0FE58: jmp 00E0FE6Eh
  loc_00E0FE5A: lea ecx, var_1C
  loc_00E0FE5D: lea edx, var_18
  loc_00E0FE60: push ecx
  loc_00E0FE61: push edx
  loc_00E0FE62: push 00000002h
  loc_00E0FE64: call [00401048h] ; __vbaFreeObjList
  loc_00E0FE6A: add esp, 0000000Ch
  loc_00E0FE6D: ret
  loc_00E0FE6E: ret
  loc_00E0FE6F: mov eax, Me
  loc_00E0FE72: push eax
  loc_00E0FE73: mov ecx, [eax]
  loc_00E0FE75: call [ecx+00000008h]
  loc_00E0FE78: mov eax, var_4
  loc_00E0FE7B: mov ecx, var_14
  loc_00E0FE7E: pop edi
  loc_00E0FE7F: pop esi
  loc_00E0FE80: mov fs:[00000000h], ecx
  loc_00E0FE87: pop ebx
  loc_00E0FE88: mov esp, ebp
  loc_00E0FE8A: pop ebp
  loc_00E0FE8B: retn 0004h
End Sub

Private Sub Timer3_Timer() 'E0FEA0
  loc_00E0FEA0: push ebp
  loc_00E0FEA1: mov ebp, esp
  loc_00E0FEA3: sub esp, 0000000Ch
  loc_00E0FEA6: push 00402836h ; __vbaExceptHandler
  loc_00E0FEAB: mov eax, fs:[00000000h]
  loc_00E0FEB1: push eax
  loc_00E0FEB2: mov fs:[00000000h], esp
  loc_00E0FEB9: sub esp, 00000028h
  loc_00E0FEBC: push ebx
  loc_00E0FEBD: push esi
  loc_00E0FEBE: push edi
  loc_00E0FEBF: mov var_C, esp
  loc_00E0FEC2: mov var_8, 004022A8h
  loc_00E0FEC9: mov esi, Me
  loc_00E0FECC: mov eax, esi
  loc_00E0FECE: and eax, 00000001h
  loc_00E0FED1: mov var_4, eax
  loc_00E0FED4: and esi, FFFFFFFEh
  loc_00E0FED7: push esi
  loc_00E0FED8: mov Me, esi
  loc_00E0FEDB: mov ecx, [esi]
  loc_00E0FEDD: call [ecx+00000004h]
  loc_00E0FEE0: mov edx, [esi]
  loc_00E0FEE2: xor eax, eax
  loc_00E0FEE4: push esi
  loc_00E0FEE5: mov var_18, eax
  loc_00E0FEE8: mov var_1C, eax
  loc_00E0FEEB: mov var_2C, eax
  loc_00E0FEEE: call [edx+000003D8h]
  loc_00E0FEF4: push eax
  loc_00E0FEF5: lea eax, var_1C
  loc_00E0FEF8: push eax
  loc_00E0FEF9: call [004010ACh] ; __vbaObjSet
  loc_00E0FEFF: lea ecx, var_2C
  loc_00E0FF02: mov edi, eax
  loc_00E0FF04: push ecx
  loc_00E0FF05: call [004011E8h] ; rtcGetTimeVar
  loc_00E0FF0B: mov ebx, [edi]
  loc_00E0FF0D: lea edx, var_2C
  loc_00E0FF10: lea eax, var_18
  loc_00E0FF13: push edx
  loc_00E0FF14: push eax
  loc_00E0FF15: call [00401180h] ; __vbaStrVarVal
  loc_00E0FF1B: push eax
  loc_00E0FF1C: push edi
  loc_00E0FF1D: call [ebx+00000054h]
  loc_00E0FF20: test eax, eax
  loc_00E0FF22: fnclex
  loc_00E0FF24: jge 00E0FF35h
  loc_00E0FF26: push 00000054h
  loc_00E0FF28: push 006E0410h
  loc_00E0FF2D: push edi
  loc_00E0FF2E: push eax
  loc_00E0FF2F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FF35: mov edi, [00401258h] ; __vbaFreeStr
  loc_00E0FF3B: lea ecx, var_18
  loc_00E0FF3E: call edi
  loc_00E0FF40: lea ecx, var_1C
  loc_00E0FF43: call [00401254h] ; __vbaFreeObj
  loc_00E0FF49: lea ecx, var_2C
  loc_00E0FF4C: call [00401024h] ; __vbaFreeVar
  loc_00E0FF52: mov ecx, [esi]
  loc_00E0FF54: push esi
  loc_00E0FF55: call [ecx+000003D0h]
  loc_00E0FF5B: lea edx, var_1C
  loc_00E0FF5E: push eax
  loc_00E0FF5F: push edx
  loc_00E0FF60: call [004010ACh] ; __vbaObjSet
  loc_00E0FF66: mov esi, eax
  loc_00E0FF68: lea eax, var_2C
  loc_00E0FF6B: push eax
  loc_00E0FF6C: call [004011D8h] ; rtcGetDateVar
  loc_00E0FF72: mov ebx, [esi]
  loc_00E0FF74: lea ecx, var_2C
  loc_00E0FF77: lea edx, var_18
  loc_00E0FF7A: push ecx
  loc_00E0FF7B: push edx
  loc_00E0FF7C: call [00401180h] ; __vbaStrVarVal
  loc_00E0FF82: push eax
  loc_00E0FF83: push esi
  loc_00E0FF84: call [ebx+00000054h]
  loc_00E0FF87: test eax, eax
  loc_00E0FF89: fnclex
  loc_00E0FF8B: jge 00E0FF9Ch
  loc_00E0FF8D: push 00000054h
  loc_00E0FF8F: push 006E0410h
  loc_00E0FF94: push esi
  loc_00E0FF95: push eax
  loc_00E0FF96: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0FF9C: lea ecx, var_18
  loc_00E0FF9F: call edi
  loc_00E0FFA1: lea ecx, var_1C
  loc_00E0FFA4: call [00401254h] ; __vbaFreeObj
  loc_00E0FFAA: lea ecx, var_2C
  loc_00E0FFAD: call [00401024h] ; __vbaFreeVar
  loc_00E0FFB3: mov var_4, 00000000h
  loc_00E0FFBA: push 00E0FFDEh
  loc_00E0FFBF: jmp 00E0FFDDh
  loc_00E0FFC1: lea ecx, var_18
  loc_00E0FFC4: call [00401258h] ; __vbaFreeStr
  loc_00E0FFCA: lea ecx, var_1C
  loc_00E0FFCD: call [00401254h] ; __vbaFreeObj
  loc_00E0FFD3: lea ecx, var_2C
  loc_00E0FFD6: call [00401024h] ; __vbaFreeVar
  loc_00E0FFDC: ret
  loc_00E0FFDD: ret
  loc_00E0FFDE: mov eax, Me
  loc_00E0FFE1: push eax
  loc_00E0FFE2: mov ecx, [eax]
  loc_00E0FFE4: call [ecx+00000008h]
  loc_00E0FFE7: mov eax, var_4
  loc_00E0FFEA: mov ecx, var_14
  loc_00E0FFED: pop edi
  loc_00E0FFEE: pop esi
  loc_00E0FFEF: mov fs:[00000000h], ecx
  loc_00E0FFF6: pop ebx
  loc_00E0FFF7: mov esp, ebp
  loc_00E0FFF9: pop ebp
  loc_00E0FFFA: retn 0004h
End Sub

Private Sub Form_Load() 'E08A80
  loc_00E08A80: push ebp
  loc_00E08A81: mov ebp, esp
  loc_00E08A83: sub esp, 0000000Ch
  loc_00E08A86: push 00402836h ; __vbaExceptHandler
  loc_00E08A8B: mov eax, fs:[00000000h]
  loc_00E08A91: push eax
  loc_00E08A92: mov fs:[00000000h], esp
  loc_00E08A99: sub esp, 00000070h
  loc_00E08A9C: push ebx
  loc_00E08A9D: push esi
  loc_00E08A9E: push edi
  loc_00E08A9F: mov var_C, esp
  loc_00E08AA2: mov var_8, 004020D0h
  loc_00E08AA9: mov esi, Me
  loc_00E08AAC: mov eax, esi
  loc_00E08AAE: and eax, 00000001h
  loc_00E08AB1: mov var_4, eax
  loc_00E08AB4: and esi, FFFFFFFEh
  loc_00E08AB7: push esi
  loc_00E08AB8: mov Me, esi
  loc_00E08ABB: mov ecx, [esi]
  loc_00E08ABD: call [ecx+00000004h]
  loc_00E08AC0: mov edx, [esi]
  loc_00E08AC2: xor eax, eax
  loc_00E08AC4: push esi
  loc_00E08AC5: mov var_18, eax
  loc_00E08AC8: mov var_1C, eax
  loc_00E08ACB: mov var_20, eax
  loc_00E08ACE: mov var_30, eax
  loc_00E08AD1: call [edx+00000330h]
  loc_00E08AD7: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E08ADD: push eax
  loc_00E08ADE: lea eax, var_1C
  loc_00E08AE1: push eax
  loc_00E08AE2: call edi
  loc_00E08AE4: mov ebx, eax
  loc_00E08AE6: push 00000000h
  loc_00E08AE8: push ebx
  loc_00E08AE9: mov ecx, [ebx]
  loc_00E08AEB: call [ecx+0000005Ch]
  loc_00E08AEE: test eax, eax
  loc_00E08AF0: fnclex
  loc_00E08AF2: jge 00E08B03h
  loc_00E08AF4: push 0000005Ch
  loc_00E08AF6: push 006DCAE0h
  loc_00E08AFB: push ebx
  loc_00E08AFC: push eax
  loc_00E08AFD: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08B03: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E08B09: lea ecx, var_1C
  loc_00E08B0C: call ebx
  loc_00E08B0E: mov edx, [esi]
  loc_00E08B10: push esi
  loc_00E08B11: call [edx+0000032Ch]
  loc_00E08B17: push eax
  loc_00E08B18: lea eax, var_1C
  loc_00E08B1B: push eax
  loc_00E08B1C: call edi
  loc_00E08B1E: mov ecx, [eax]
  loc_00E08B20: push 43870000h
  loc_00E08B25: push eax
  loc_00E08B26: mov var_54, eax
  loc_00E08B29: call [ecx+00000074h]
  loc_00E08B2C: test eax, eax
  loc_00E08B2E: fnclex
  loc_00E08B30: jge 00E08B44h
  loc_00E08B32: mov edx, var_54
  loc_00E08B35: push 00000074h
  loc_00E08B37: push 006E0410h
  loc_00E08B3C: push edx
  loc_00E08B3D: push eax
  loc_00E08B3E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08B44: lea ecx, var_1C
  loc_00E08B47: call ebx
  loc_00E08B49: sub esp, 00000010h
  loc_00E08B4C: mov ecx, 0000000Bh
  loc_00E08B51: mov edx, esp
  loc_00E08B53: xor eax, eax
  loc_00E08B55: push 80010007h
  loc_00E08B5A: push esi
  loc_00E08B5B: mov [edx], ecx
  loc_00E08B5D: mov ecx, var_3C
  loc_00E08B60: mov [edx+00000004h], ecx
  loc_00E08B63: mov ecx, [esi]
  loc_00E08B65: mov [edx+00000008h], eax
  loc_00E08B68: mov eax, var_34
  loc_00E08B6B: mov [edx+0000000Ch], eax
  loc_00E08B6E: call [ecx+000003B8h]
  loc_00E08B74: lea edx, var_1C
  loc_00E08B77: push eax
  loc_00E08B78: push edx
  loc_00E08B79: call edi
  loc_00E08B7B: push eax
  loc_00E08B7C: call [00401238h] ; __vbaLateIdSt
  loc_00E08B82: lea ecx, var_1C
  loc_00E08B85: call ebx
  loc_00E08B87: lea ecx, [esi+00000034h]
  loc_00E08B8A: mov edx, 006E05ACh ; "select * from biodata"
  loc_00E08B8F: call [004011B0h] ; __vbaStrCopy
  loc_00E08B95: mov edx, var_3C
  loc_00E08B98: sub esp, 00000010h
  loc_00E08B9B: mov ecx, esp
  loc_00E08B9D: mov eax, 00004008h
  loc_00E08BA2: push 0000000Eh
  loc_00E08BA4: push esi
  loc_00E08BA5: mov [ecx], eax
  loc_00E08BA7: lea eax, [esi+00000034h]
  loc_00E08BAA: mov [ecx+00000004h], edx
  loc_00E08BAD: mov [ecx+00000008h], eax
  loc_00E08BB0: mov eax, var_34
  loc_00E08BB3: mov [ecx+0000000Ch], eax
  loc_00E08BB6: mov ecx, [esi]
  loc_00E08BB8: call [ecx+00000490h]
  loc_00E08BBE: lea edx, var_1C
  loc_00E08BC1: push eax
  loc_00E08BC2: push edx
  loc_00E08BC3: call edi
  loc_00E08BC5: push eax
  loc_00E08BC6: call [00401238h] ; __vbaLateIdSt
  loc_00E08BCC: lea ecx, var_1C
  loc_00E08BCF: call ebx
  loc_00E08BD1: mov eax, [esi]
  loc_00E08BD3: push 00000000h
  loc_00E08BD5: push 00000019h
  loc_00E08BD7: push esi
  loc_00E08BD8: call [eax+00000490h]
  loc_00E08BDE: lea ecx, var_1C
  loc_00E08BE1: push eax
  loc_00E08BE2: push ecx
  loc_00E08BE3: call edi
  loc_00E08BE5: push eax
  loc_00E08BE6: call [00401030h] ; __vbaLateIdCall
  loc_00E08BEC: add esp, 0000000Ch
  loc_00E08BEF: lea ecx, var_1C
  loc_00E08BF2: call ebx
  loc_00E08BF4: mov edx, [esi]
  loc_00E08BF6: push 006E05D8h
  loc_00E08BFB: push esi
  loc_00E08BFC: call [edx+00000490h]
  loc_00E08C02: push eax
  loc_00E08C03: lea eax, var_1C
  loc_00E08C06: push eax
  loc_00E08C07: call edi
  loc_00E08C09: push eax
  loc_00E08C0A: call [00401224h] ; __vbaCastObj
  loc_00E08C10: lea ecx, var_30
  loc_00E08C13: mov var_28, eax
  loc_00E08C16: push ecx
  loc_00E08C17: mov var_30, 0000000Dh
  loc_00E08C1E: call [004011F8h] ; __vbaVerifyVarObj
  loc_00E08C24: mov ecx, [eax]
  loc_00E08C26: sub esp, 00000010h
  loc_00E08C29: mov edx, esp
  loc_00E08C2B: mov [edx], ecx
  loc_00E08C2D: mov ecx, [eax+00000004h]
  loc_00E08C30: push 00000000h
  loc_00E08C32: push 0000002Ah
  loc_00E08C34: mov [edx+00000004h], ecx
  loc_00E08C37: mov ecx, [eax+00000008h]
  loc_00E08C3A: push esi
  loc_00E08C3B: mov eax, [eax+0000000Ch]
  loc_00E08C3E: mov [edx+00000008h], ecx
  loc_00E08C41: mov ecx, [esi]
  loc_00E08C43: mov [edx+0000000Ch], eax
  loc_00E08C46: call [ecx+00000494h]
  loc_00E08C4C: lea edx, var_20
  loc_00E08C4F: push eax
  loc_00E08C50: push edx
  loc_00E08C51: call edi
  loc_00E08C53: push eax
  loc_00E08C54: call [0040116Ch] ; __vbaLateIdStAd
  loc_00E08C5A: lea eax, var_20
  loc_00E08C5D: lea ecx, var_1C
  loc_00E08C60: push eax
  loc_00E08C61: push ecx
  loc_00E08C62: push 00000002h
  loc_00E08C64: call [00401048h] ; __vbaFreeObjList
  loc_00E08C6A: add esp, 00000028h
  loc_00E08C6D: lea ecx, var_30
  loc_00E08C70: call [00401024h] ; __vbaFreeVar
  loc_00E08C76: mov edx, [esi]
  loc_00E08C78: push 00000000h
  loc_00E08C7A: push FFFFFDDAh
  loc_00E08C7F: push esi
  loc_00E08C80: call [edx+00000494h]
  loc_00E08C86: push eax
  loc_00E08C87: lea eax, var_1C
  loc_00E08C8A: push eax
  loc_00E08C8B: call edi
  loc_00E08C8D: push eax
  loc_00E08C8E: call [00401030h] ; __vbaLateIdCall
  loc_00E08C94: add esp, 0000000Ch
  loc_00E08C97: lea ecx, var_1C
  loc_00E08C9A: call ebx
  loc_00E08C9C: mov ecx, [esi]
  loc_00E08C9E: push esi
  loc_00E08C9F: call [ecx+00000378h]
  loc_00E08CA5: lea edx, var_1C
  loc_00E08CA8: push eax
  loc_00E08CA9: push edx
  loc_00E08CAA: call edi
  loc_00E08CAC: mov edx, [eax]
  loc_00E08CAE: sub esp, 00000010h
  loc_00E08CB1: mov var_64, edx
  loc_00E08CB4: mov edx, esp
  loc_00E08CB6: mov ecx, 0000000Ah
  loc_00E08CBB: mov var_54, eax
  loc_00E08CBE: mov eax, var_3C
  loc_00E08CC1: mov [edx], ecx
  loc_00E08CC3: mov ecx, var_34
  loc_00E08CC6: push 006E05ECh ; "Islam"
  loc_00E08CCB: mov [edx+00000004h], eax
  loc_00E08CCE: mov eax, 80020004h
  loc_00E08CD3: mov [edx+00000008h], eax
  loc_00E08CD6: mov eax, var_64
  loc_00E08CD9: mov [edx+0000000Ch], ecx
  loc_00E08CDC: mov edx, var_54
  loc_00E08CDF: push edx
  loc_00E08CE0: call [eax+000001ECh]
  loc_00E08CE6: test eax, eax
  loc_00E08CE8: fnclex
  loc_00E08CEA: jge 00E08D01h
  loc_00E08CEC: mov ecx, var_54
  loc_00E08CEF: push 000001ECh
  loc_00E08CF4: push 006E0400h
  loc_00E08CF9: push ecx
  loc_00E08CFA: push eax
  loc_00E08CFB: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08D01: lea ecx, var_1C
  loc_00E08D04: call ebx
  loc_00E08D06: mov edx, [esi]
  loc_00E08D08: push esi
  loc_00E08D09: call [edx+00000378h]
  loc_00E08D0F: push eax
  loc_00E08D10: lea eax, var_1C
  loc_00E08D13: push eax
  loc_00E08D14: call edi
  loc_00E08D16: mov edx, [eax]
  loc_00E08D18: sub esp, 00000010h
  loc_00E08D1B: mov var_68, edx
  loc_00E08D1E: mov edx, esp
  loc_00E08D20: mov ecx, 0000000Ah
  loc_00E08D25: mov var_54, eax
  loc_00E08D28: mov eax, var_3C
  loc_00E08D2B: mov [edx], ecx
  loc_00E08D2D: mov ecx, var_34
  loc_00E08D30: push 006E05FCh ; "Katholik"
  loc_00E08D35: mov [edx+00000004h], eax
  loc_00E08D38: mov eax, 80020004h
  loc_00E08D3D: mov [edx+00000008h], eax
  loc_00E08D40: mov eax, var_68
  loc_00E08D43: mov [edx+0000000Ch], ecx
  loc_00E08D46: mov edx, var_54
  loc_00E08D49: push edx
  loc_00E08D4A: call [eax+000001ECh]
  loc_00E08D50: test eax, eax
  loc_00E08D52: fnclex
  loc_00E08D54: jge 00E08D6Bh
  loc_00E08D56: mov ecx, var_54
  loc_00E08D59: push 000001ECh
  loc_00E08D5E: push 006E0400h
  loc_00E08D63: push ecx
  loc_00E08D64: push eax
  loc_00E08D65: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08D6B: lea ecx, var_1C
  loc_00E08D6E: call ebx
  loc_00E08D70: mov edx, [esi]
  loc_00E08D72: push esi
  loc_00E08D73: call [edx+00000378h]
  loc_00E08D79: push eax
  loc_00E08D7A: lea eax, var_1C
  loc_00E08D7D: push eax
  loc_00E08D7E: call edi
  loc_00E08D80: mov edx, [eax]
  loc_00E08D82: sub esp, 00000010h
  loc_00E08D85: mov var_6C, edx
  loc_00E08D88: mov edx, esp
  loc_00E08D8A: mov ecx, 0000000Ah
  loc_00E08D8F: mov var_54, eax
  loc_00E08D92: mov eax, var_3C
  loc_00E08D95: mov [edx], ecx
  loc_00E08D97: mov ecx, var_34
  loc_00E08D9A: push 006E0614h ; "Protestant"
  loc_00E08D9F: mov [edx+00000004h], eax
  loc_00E08DA2: mov eax, 80020004h
  loc_00E08DA7: mov [edx+00000008h], eax
  loc_00E08DAA: mov eax, var_6C
  loc_00E08DAD: mov [edx+0000000Ch], ecx
  loc_00E08DB0: mov edx, var_54
  loc_00E08DB3: push edx
  loc_00E08DB4: call [eax+000001ECh]
  loc_00E08DBA: test eax, eax
  loc_00E08DBC: fnclex
  loc_00E08DBE: jge 00E08DD5h
  loc_00E08DC0: mov ecx, var_54
  loc_00E08DC3: push 000001ECh
  loc_00E08DC8: push 006E0400h
  loc_00E08DCD: push ecx
  loc_00E08DCE: push eax
  loc_00E08DCF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08DD5: lea ecx, var_1C
  loc_00E08DD8: call ebx
  loc_00E08DDA: mov edx, [esi]
  loc_00E08DDC: push esi
  loc_00E08DDD: call [edx+00000378h]
  loc_00E08DE3: push eax
  loc_00E08DE4: lea eax, var_1C
  loc_00E08DE7: push eax
  loc_00E08DE8: call edi
  loc_00E08DEA: mov edx, [eax]
  loc_00E08DEC: sub esp, 00000010h
  loc_00E08DEF: mov var_70, edx
  loc_00E08DF2: mov edx, esp
  loc_00E08DF4: mov ecx, 0000000Ah
  loc_00E08DF9: mov var_54, eax
  loc_00E08DFC: mov eax, var_3C
  loc_00E08DFF: mov [edx], ecx
  loc_00E08E01: mov ecx, var_34
  loc_00E08E04: push 006E0630h ; "Hindu-Buddha"
  loc_00E08E09: mov [edx+00000004h], eax
  loc_00E08E0C: mov eax, 80020004h
  loc_00E08E11: mov [edx+00000008h], eax
  loc_00E08E14: mov eax, var_70
  loc_00E08E17: mov [edx+0000000Ch], ecx
  loc_00E08E1A: mov edx, var_54
  loc_00E08E1D: push edx
  loc_00E08E1E: call [eax+000001ECh]
  loc_00E08E24: test eax, eax
  loc_00E08E26: fnclex
  loc_00E08E28: jge 00E08E3Fh
  loc_00E08E2A: mov ecx, var_54
  loc_00E08E2D: push 000001ECh
  loc_00E08E32: push 006E0400h
  loc_00E08E37: push ecx
  loc_00E08E38: push eax
  loc_00E08E39: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08E3F: lea ecx, var_1C
  loc_00E08E42: call ebx
  loc_00E08E44: mov edx, [esi]
  loc_00E08E46: push esi
  loc_00E08E47: call [edx+00000378h]
  loc_00E08E4D: push eax
  loc_00E08E4E: lea eax, var_1C
  loc_00E08E51: push eax
  loc_00E08E52: call edi
  loc_00E08E54: mov edx, [eax]
  loc_00E08E56: sub esp, 00000010h
  loc_00E08E59: mov var_74, edx
  loc_00E08E5C: mov edx, esp
  loc_00E08E5E: mov ecx, 0000000Ah
  loc_00E08E63: mov var_54, eax
  loc_00E08E66: mov eax, var_3C
  loc_00E08E69: mov [edx], ecx
  loc_00E08E6B: mov ecx, var_34
  loc_00E08E6E: push 006E0650h ; "Khong Hu Chu"
  loc_00E08E73: mov [edx+00000004h], eax
  loc_00E08E76: mov eax, 80020004h
  loc_00E08E7B: mov [edx+00000008h], eax
  loc_00E08E7E: mov eax, var_74
  loc_00E08E81: mov [edx+0000000Ch], ecx
  loc_00E08E84: mov edx, var_54
  loc_00E08E87: push edx
  loc_00E08E88: call [eax+000001ECh]
  loc_00E08E8E: test eax, eax
  loc_00E08E90: fnclex
  loc_00E08E92: jge 00E08EA9h
  loc_00E08E94: mov ecx, var_54
  loc_00E08E97: push 000001ECh
  loc_00E08E9C: push 006E0400h
  loc_00E08EA1: push ecx
  loc_00E08EA2: push eax
  loc_00E08EA3: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08EA9: lea ecx, var_1C
  loc_00E08EAC: call ebx
  loc_00E08EAE: mov edx, [esi]
  loc_00E08EB0: push esi
  loc_00E08EB1: call [edx+0000030Ch]
  loc_00E08EB7: push eax
  loc_00E08EB8: lea eax, var_1C
  loc_00E08EBB: push eax
  loc_00E08EBC: call edi
  loc_00E08EBE: mov edx, [eax]
  loc_00E08EC0: sub esp, 00000010h
  loc_00E08EC3: mov var_78, edx
  loc_00E08EC6: mov edx, esp
  loc_00E08EC8: mov ecx, 0000000Ah
  loc_00E08ECD: mov var_54, eax
  loc_00E08ED0: mov eax, var_3C
  loc_00E08ED3: mov [edx], ecx
  loc_00E08ED5: mov ecx, var_34
  loc_00E08ED8: push 006E0670h ; "Rekayasa Perangkat Lunak"
  loc_00E08EDD: mov [edx+00000004h], eax
  loc_00E08EE0: mov eax, 80020004h
  loc_00E08EE5: mov [edx+00000008h], eax
  loc_00E08EE8: mov eax, var_78
  loc_00E08EEB: mov [edx+0000000Ch], ecx
  loc_00E08EEE: mov edx, var_54
  loc_00E08EF1: push edx
  loc_00E08EF2: call [eax+000001ECh]
  loc_00E08EF8: test eax, eax
  loc_00E08EFA: fnclex
  loc_00E08EFC: jge 00E08F13h
  loc_00E08EFE: mov ecx, var_54
  loc_00E08F01: push 000001ECh
  loc_00E08F06: push 006E0400h
  loc_00E08F0B: push ecx
  loc_00E08F0C: push eax
  loc_00E08F0D: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08F13: lea ecx, var_1C
  loc_00E08F16: call ebx
  loc_00E08F18: mov edx, [esi]
  loc_00E08F1A: push esi
  loc_00E08F1B: call [edx+0000030Ch]
  loc_00E08F21: push eax
  loc_00E08F22: lea eax, var_1C
  loc_00E08F25: push eax
  loc_00E08F26: call edi
  loc_00E08F28: mov edx, [eax]
  loc_00E08F2A: sub esp, 00000010h
  loc_00E08F2D: mov var_7C, edx
  loc_00E08F30: mov edx, esp
  loc_00E08F32: mov ecx, 0000000Ah
  loc_00E08F37: mov var_54, eax
  loc_00E08F3A: mov eax, var_3C
  loc_00E08F3D: mov [edx], ecx
  loc_00E08F3F: mov ecx, var_34
  loc_00E08F42: push 006E049Ch ; "Tata Busana"
  loc_00E08F47: mov [edx+00000004h], eax
  loc_00E08F4A: mov eax, 80020004h
  loc_00E08F4F: mov [edx+00000008h], eax
  loc_00E08F52: mov eax, var_7C
  loc_00E08F55: mov [edx+0000000Ch], ecx
  loc_00E08F58: mov edx, var_54
  loc_00E08F5B: push edx
  loc_00E08F5C: call [eax+000001ECh]
  loc_00E08F62: test eax, eax
  loc_00E08F64: fnclex
  loc_00E08F66: jge 00E08F7Dh
  loc_00E08F68: mov ecx, var_54
  loc_00E08F6B: push 000001ECh
  loc_00E08F70: push 006E0400h
  loc_00E08F75: push ecx
  loc_00E08F76: push eax
  loc_00E08F77: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08F7D: lea ecx, var_1C
  loc_00E08F80: call ebx
  loc_00E08F82: mov edx, [esi]
  loc_00E08F84: push esi
  loc_00E08F85: call [edx+0000030Ch]
  loc_00E08F8B: push eax
  loc_00E08F8C: lea eax, var_1C
  loc_00E08F8F: push eax
  loc_00E08F90: call edi
  loc_00E08F92: mov edx, [eax]
  loc_00E08F94: sub esp, 00000010h
  loc_00E08F97: mov var_54, eax
  loc_00E08F9A: mov eax, esp
  loc_00E08F9C: mov ecx, 0000000Ah
  loc_00E08FA1: mov var_38, 80020004h
  loc_00E08FA8: mov [eax], ecx
  loc_00E08FAA: mov ecx, var_3C
  loc_00E08FAD: push 006E06A8h ; "Farmasi"
  loc_00E08FB2: mov [eax+00000004h], ecx
  loc_00E08FB5: mov ecx, var_38
  loc_00E08FB8: mov [eax+00000008h], ecx
  loc_00E08FBB: mov ecx, var_34
  loc_00E08FBE: mov [eax+0000000Ch], ecx
  loc_00E08FC1: mov eax, var_54
  loc_00E08FC4: push eax
  loc_00E08FC5: call [edx+000001ECh]
  loc_00E08FCB: test eax, eax
  loc_00E08FCD: fnclex
  loc_00E08FCF: jge 00E08FE6h
  loc_00E08FD1: mov ecx, var_54
  loc_00E08FD4: push 000001ECh
  loc_00E08FD9: push 006E0400h
  loc_00E08FDE: push ecx
  loc_00E08FDF: push eax
  loc_00E08FE0: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08FE6: lea ecx, var_1C
  loc_00E08FE9: call ebx
  loc_00E08FEB: mov edx, [esi]
  loc_00E08FED: push esi
  loc_00E08FEE: call [edx+00000338h]
  loc_00E08FF4: push eax
  loc_00E08FF5: lea eax, var_1C
  loc_00E08FF8: push eax
  loc_00E08FF9: call edi
  loc_00E08FFB: mov ecx, [eax]
  loc_00E08FFD: push 00000000h
  loc_00E08FFF: push eax
  loc_00E09000: mov var_54, eax
  loc_00E09003: call [ecx+0000009Ch]
  loc_00E09009: test eax, eax
  loc_00E0900B: fnclex
  loc_00E0900D: jge 00E09024h
  loc_00E0900F: mov edx, var_54
  loc_00E09012: push 0000009Ch
  loc_00E09017: push 006DCAD0h
  loc_00E0901C: push edx
  loc_00E0901D: push eax
  loc_00E0901E: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09024: lea ecx, var_1C
  loc_00E09027: call ebx
  loc_00E09029: mov eax, [esi]
  loc_00E0902B: push esi
  loc_00E0902C: call [eax+00000334h]
  loc_00E09032: lea ecx, var_1C
  loc_00E09035: push eax
  loc_00E09036: push ecx
  loc_00E09037: call edi
  loc_00E09039: mov edx, [eax]
  loc_00E0903B: push 00000000h
  loc_00E0903D: push eax
  loc_00E0903E: mov var_54, eax
  loc_00E09041: call [edx+0000005Ch]
  loc_00E09044: test eax, eax
  loc_00E09046: fnclex
  loc_00E09048: jge 00E0905Ch
  loc_00E0904A: mov ecx, var_54
  loc_00E0904D: push 0000005Ch
  loc_00E0904F: push 006DCAE0h
  loc_00E09054: push ecx
  loc_00E09055: push eax
  loc_00E09056: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0905C: lea ecx, var_1C
  loc_00E0905F: call ebx
  loc_00E09061: mov edx, [esi]
  loc_00E09063: push esi
  loc_00E09064: call [edx+000003E8h]
  loc_00E0906A: push eax
  loc_00E0906B: lea eax, var_1C
  loc_00E0906E: push eax
  loc_00E0906F: call edi
  loc_00E09071: mov ecx, [eax]
  loc_00E09073: push 00000000h
  loc_00E09075: push eax
  loc_00E09076: mov var_54, eax
  loc_00E09079: call [ecx+0000008Ch]
  loc_00E0907F: test eax, eax
  loc_00E09081: fnclex
  loc_00E09083: jge 00E0909Ah
  loc_00E09085: mov edx, var_54
  loc_00E09088: push 0000008Ch
  loc_00E0908D: push 006DCDA0h
  loc_00E09092: push edx
  loc_00E09093: push eax
  loc_00E09094: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0909A: lea ecx, var_1C
  loc_00E0909D: call ebx
  loc_00E0909F: mov eax, [esi]
  loc_00E090A1: push esi
  loc_00E090A2: call [eax+0000039Ch]
  loc_00E090A8: lea ecx, var_1C
  loc_00E090AB: push eax
  loc_00E090AC: push ecx
  loc_00E090AD: call edi
  loc_00E090AF: mov edx, [eax]
  loc_00E090B1: push 006E03C8h ; "00001"
  loc_00E090B6: push eax
  loc_00E090B7: mov var_54, eax
  loc_00E090BA: call [edx+000000A4h]
  loc_00E090C0: test eax, eax
  loc_00E090C2: fnclex
  loc_00E090C4: jge 00E090DBh
  loc_00E090C6: mov ecx, var_54
  loc_00E090C9: push 000000A4h
  loc_00E090CE: push 006DCB70h
  loc_00E090D3: push ecx
  loc_00E090D4: push eax
  loc_00E090D5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E090DB: lea ecx, var_1C
  loc_00E090DE: call ebx
  loc_00E090E0: mov edx, [esi]
  loc_00E090E2: push esi
  loc_00E090E3: call [edx+000003E0h]
  loc_00E090E9: push eax
  loc_00E090EA: lea eax, var_1C
  loc_00E090ED: push eax
  loc_00E090EE: call edi
  loc_00E090F0: mov ecx, [eax]
  loc_00E090F2: push 00000000h
  loc_00E090F4: push eax
  loc_00E090F5: mov var_54, eax
  loc_00E090F8: call [ecx+0000008Ch]
  loc_00E090FE: test eax, eax
  loc_00E09100: fnclex
  loc_00E09102: jge 00E09119h
  loc_00E09104: mov edx, var_54
  loc_00E09107: push 0000008Ch
  loc_00E0910C: push 006DCDA0h
  loc_00E09111: push edx
  loc_00E09112: push eax
  loc_00E09113: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09119: lea ecx, var_1C
  loc_00E0911C: call ebx
  loc_00E0911E: mov eax, [esi]
  loc_00E09120: push esi
  loc_00E09121: call [eax+00000310h]
  loc_00E09127: lea ecx, var_1C
  loc_00E0912A: push eax
  loc_00E0912B: push ecx
  loc_00E0912C: call edi
  loc_00E0912E: mov edx, [eax]
  loc_00E09130: push 00000000h
  loc_00E09132: push eax
  loc_00E09133: mov var_54, eax
  loc_00E09136: call [edx+0000009Ch]
  loc_00E0913C: test eax, eax
  loc_00E0913E: fnclex
  loc_00E09140: jge 00E09157h
  loc_00E09142: mov ecx, var_54
  loc_00E09145: push 0000009Ch
  loc_00E0914A: push 006DCAD0h
  loc_00E0914F: push ecx
  loc_00E09150: push eax
  loc_00E09151: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09157: lea ecx, var_1C
  loc_00E0915A: call ebx
  loc_00E0915C: mov edx, [esi]
  loc_00E0915E: push esi
  loc_00E0915F: call [edx+000003C8h]
  loc_00E09165: push eax
  loc_00E09166: lea eax, var_1C
  loc_00E09169: push eax
  loc_00E0916A: call edi
  loc_00E0916C: mov ecx, [eax]
  loc_00E0916E: push 00000000h
  loc_00E09170: push eax
  loc_00E09171: mov var_54, eax
  loc_00E09174: call [ecx+0000008Ch]
  loc_00E0917A: test eax, eax
  loc_00E0917C: fnclex
  loc_00E0917E: jge 00E09195h
  loc_00E09180: mov edx, var_54
  loc_00E09183: push 0000008Ch
  loc_00E09188: push 006DCDA0h
  loc_00E0918D: push edx
  loc_00E0918E: push eax
  loc_00E0918F: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09195: lea ecx, var_1C
  loc_00E09198: call ebx
  loc_00E0919A: mov eax, [esi]
  loc_00E0919C: push esi
  loc_00E0919D: call [eax+000003D0h]
  loc_00E091A3: lea ecx, var_1C
  loc_00E091A6: push eax
  loc_00E091A7: push ecx
  loc_00E091A8: call edi
  loc_00E091AA: lea edx, var_30
  loc_00E091AD: mov ebx, eax
  loc_00E091AF: push edx
  loc_00E091B0: mov var_54, ebx
  loc_00E091B3: call [004011D8h] ; rtcGetDateVar
  loc_00E091B9: mov ebx, [ebx]
  loc_00E091BB: lea eax, var_30
  loc_00E091BE: lea ecx, var_18
  loc_00E091C1: push eax
  loc_00E091C2: push ecx
  loc_00E091C3: call [00401180h] ; __vbaStrVarVal
  loc_00E091C9: mov edx, ebx
  loc_00E091CB: mov ebx, var_54
  loc_00E091CE: push eax
  loc_00E091CF: push ebx
  loc_00E091D0: call [edx+00000054h]
  loc_00E091D3: test eax, eax
  loc_00E091D5: fnclex
  loc_00E091D7: jge 00E091E8h
  loc_00E091D9: push 00000054h
  loc_00E091DB: push 006E0410h
  loc_00E091E0: push ebx
  loc_00E091E1: push eax
  loc_00E091E2: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E091E8: lea ecx, var_18
  loc_00E091EB: call [00401258h] ; __vbaFreeStr
  loc_00E091F1: lea ecx, var_1C
  loc_00E091F4: call [00401254h] ; __vbaFreeObj
  loc_00E091FA: lea ecx, var_30
  loc_00E091FD: call [00401024h] ; __vbaFreeVar
  loc_00E09203: mov eax, [esi]
  loc_00E09205: push esi
  loc_00E09206: call [eax+000003D8h]
  loc_00E0920C: lea ecx, var_1C
  loc_00E0920F: push eax
  loc_00E09210: push ecx
  loc_00E09211: call edi
  loc_00E09213: lea edx, var_30
  loc_00E09216: mov ebx, eax
  loc_00E09218: push edx
  loc_00E09219: mov var_54, ebx
  loc_00E0921C: call [004011E8h] ; rtcGetTimeVar
  loc_00E09222: mov ebx, [ebx]
  loc_00E09224: lea eax, var_30
  loc_00E09227: lea ecx, var_18
  loc_00E0922A: push eax
  loc_00E0922B: push ecx
  loc_00E0922C: call [00401180h] ; __vbaStrVarVal
  loc_00E09232: mov edx, ebx
  loc_00E09234: mov ebx, var_54
  loc_00E09237: push eax
  loc_00E09238: push ebx
  loc_00E09239: call [edx+00000054h]
  loc_00E0923C: test eax, eax
  loc_00E0923E: fnclex
  loc_00E09240: jge 00E09251h
  loc_00E09242: push 00000054h
  loc_00E09244: push 006E0410h
  loc_00E09249: push ebx
  loc_00E0924A: push eax
  loc_00E0924B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09251: lea ecx, var_18
  loc_00E09254: call [00401258h] ; __vbaFreeStr
  loc_00E0925A: lea ecx, var_1C
  loc_00E0925D: call [00401254h] ; __vbaFreeObj
  loc_00E09263: lea ecx, var_30
  loc_00E09266: call [00401024h] ; __vbaFreeVar
  loc_00E0926C: mov eax, [esi]
  loc_00E0926E: push esi
  loc_00E0926F: call [eax+0000032Ch]
  loc_00E09275: lea ecx, var_1C
  loc_00E09278: push eax
  loc_00E09279: push ecx
  loc_00E0927A: call edi
  loc_00E0927C: mov ebx, eax
  loc_00E0927E: push 006E06BCh ; "Isi Tanggal Daftar Terlebih Dahulu"
  loc_00E09283: push ebx
  loc_00E09284: mov edx, [ebx]
  loc_00E09286: call [edx+00000054h]
  loc_00E09289: test eax, eax
  loc_00E0928B: fnclex
  loc_00E0928D: jge 00E0929Eh
  loc_00E0928F: push 00000054h
  loc_00E09291: push 006E0410h
  loc_00E09296: push ebx
  loc_00E09297: push eax
  loc_00E09298: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0929E: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E092A4: lea ecx, var_1C
  loc_00E092A7: call ebx
  loc_00E092A9: mov eax, [esi]
  loc_00E092AB: push esi
  loc_00E092AC: call [eax+0000032Ch]
  loc_00E092B2: lea ecx, var_1C
  loc_00E092B5: push eax
  loc_00E092B6: push ecx
  loc_00E092B7: call edi
  loc_00E092B9: mov esi, eax
  loc_00E092BB: push 008080FFh
  loc_00E092C0: push esi
  loc_00E092C1: mov edx, [esi]
  loc_00E092C3: call [edx+0000006Ch]
  loc_00E092C6: test eax, eax
  loc_00E092C8: fnclex
  loc_00E092CA: jge 00E092DBh
  loc_00E092CC: push 0000006Ch
  loc_00E092CE: push 006E0410h
  loc_00E092D3: push esi
  loc_00E092D4: push eax
  loc_00E092D5: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E092DB: lea ecx, var_1C
  loc_00E092DE: call ebx
  loc_00E092E0: mov var_4, 00000000h
  loc_00E092E7: fwait
  loc_00E092E8: push 00E09316h
  loc_00E092ED: jmp 00E09315h
  loc_00E092EF: lea ecx, var_18
  loc_00E092F2: call [00401258h] ; __vbaFreeStr
  loc_00E092F8: lea eax, var_20
  loc_00E092FB: lea ecx, var_1C
  loc_00E092FE: push eax
  loc_00E092FF: push ecx
  loc_00E09300: push 00000002h
  loc_00E09302: call [00401048h] ; __vbaFreeObjList
  loc_00E09308: add esp, 0000000Ch
  loc_00E0930B: lea ecx, var_30
  loc_00E0930E: call [00401024h] ; __vbaFreeVar
  loc_00E09314: ret
  loc_00E09315: ret
  loc_00E09316: mov eax, Me
  loc_00E09319: push eax
  loc_00E0931A: mov edx, [eax]
  loc_00E0931C: call [edx+00000008h]
  loc_00E0931F: mov eax, var_4
  loc_00E09322: mov ecx, var_14
  loc_00E09325: pop edi
  loc_00E09326: pop esi
  loc_00E09327: mov fs:[00000000h], ecx
  loc_00E0932E: pop ebx
  loc_00E0932F: mov esp, ebp
  loc_00E09331: pop ebp
  loc_00E09332: retn 0004h
End Sub

Private Sub Form_Unload(Cancel As Integer) 'E08830
  loc_00E08830: push ebp
  loc_00E08831: mov ebp, esp
  loc_00E08833: sub esp, 0000000Ch
  loc_00E08836: push 00402836h ; __vbaExceptHandler
  loc_00E0883B: mov eax, fs:[00000000h]
  loc_00E08841: push eax
  loc_00E08842: mov fs:[00000000h], esp
  loc_00E08849: sub esp, 0000005Ch
  loc_00E0884C: push ebx
  loc_00E0884D: push esi
  loc_00E0884E: push edi
  loc_00E0884F: mov var_C, esp
  loc_00E08852: mov var_8, 004020C0h
  loc_00E08859: mov esi, Me
  loc_00E0885C: mov eax, esi
  loc_00E0885E: and eax, 00000001h
  loc_00E08861: mov var_4, eax
  loc_00E08864: and esi, FFFFFFFEh
  loc_00E08867: push esi
  loc_00E08868: mov Me, esi
  loc_00E0886B: mov ecx, [esi]
  loc_00E0886D: call [ecx+00000004h]
  loc_00E08870: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08876: xor eax, eax
  loc_00E08878: mov var_18, eax
  loc_00E0887B: mov var_4C, eax
  loc_00E0887E: mov var_50, eax
  loc_00E08881: mov edx, [esi]
  loc_00E08883: lea eax, var_4C
  loc_00E08886: push eax
  loc_00E08887: push esi
  loc_00E08888: call [edx+00000070h]
  loc_00E0888B: test eax, eax
  loc_00E0888D: fnclex
  loc_00E0888F: jge 00E0889Ch
  loc_00E08891: push 00000070h
  loc_00E08893: push 006DFCB8h
  loc_00E08898: push esi
  loc_00E08899: push eax
  loc_00E0889A: call ebx
  loc_00E0889C: fld real4 ptr var_4C
  loc_00E0889F: fadd st0, real4 ptr [00401298h]
  loc_00E088A5: mov ecx, [esi]
  loc_00E088A7: push ecx
  loc_00E088A8: fnstsw ax
  loc_00E088AA: test al, 0Dh
  loc_00E088AC: jnz 00E08A70h
  loc_00E088B2: fstp real4 ptr [esp]
  loc_00E088B5: push esi
  loc_00E088B6: call [ecx+00000074h]
  loc_00E088B9: test eax, eax
  loc_00E088BB: fnclex
  loc_00E088BD: jge 00E088CAh
  loc_00E088BF: push 00000074h
  loc_00E088C1: push 006DFCB8h
  loc_00E088C6: push esi
  loc_00E088C7: push eax
  loc_00E088C8: call ebx
  loc_00E088CA: mov edx, [esi]
  loc_00E088CC: lea eax, var_4C
  loc_00E088CF: push eax
  loc_00E088D0: push esi
  loc_00E088D1: call [edx+00000070h]
  loc_00E088D4: test eax, eax
  loc_00E088D6: fnclex
  loc_00E088D8: jge 00E088E5h
  loc_00E088DA: push 00000070h
  loc_00E088DC: push 006DFCB8h
  loc_00E088E1: push esi
  loc_00E088E2: push eax
  loc_00E088E3: call ebx
  loc_00E088E5: mov ecx, [esi]
  loc_00E088E7: lea edx, var_50
  loc_00E088EA: push edx
  loc_00E088EB: push esi
  loc_00E088EC: call [ecx+00000078h]
  loc_00E088EF: test eax, eax
  loc_00E088F1: fnclex
  loc_00E088F3: jge 00E08900h
  loc_00E088F5: push 00000078h
  loc_00E088F7: push 006DFCB8h
  loc_00E088FC: push esi
  loc_00E088FD: push eax
  loc_00E088FE: call ebx
  loc_00E08900: sub esp, 00000010h
  loc_00E08903: mov ecx, 0000000Ah
  loc_00E08908: mov ebx, esp
  loc_00E0890A: mov eax, 80020004h
  loc_00E0890F: mov edx, eax
  loc_00E08911: sub esp, 00000010h
  loc_00E08914: mov [ebx], ecx
  loc_00E08916: mov ecx, var_44
  loc_00E08919: mov edi, [esi]
  loc_00E0891B: mov [ebx+00000004h], ecx
  loc_00E0891E: mov ecx, esp
  loc_00E08920: sub esp, 00000010h
  loc_00E08923: mov [ebx+00000008h], eax
  loc_00E08926: mov eax, var_3C
  loc_00E08929: mov [ebx+0000000Ch], eax
  loc_00E0892C: mov eax, 0000000Ah
  loc_00E08931: mov [ecx], eax
  loc_00E08933: mov eax, var_34
  loc_00E08936: mov [ecx+00000004h], eax
  loc_00E08939: mov eax, 00000004h
  loc_00E0893E: mov [ecx+00000008h], edx
  loc_00E08941: mov edx, var_2C
  loc_00E08944: mov [ecx+0000000Ch], edx
  loc_00E08947: mov edx, var_24
  loc_00E0894A: mov ecx, esp
  loc_00E0894C: mov [ecx], eax
  loc_00E0894E: mov eax, var_50
  loc_00E08951: mov [ecx+00000004h], edx
  loc_00E08954: mov [ecx+00000008h], eax
  loc_00E08957: mov eax, var_1C
  loc_00E0895A: mov [ecx+0000000Ch], eax
  loc_00E0895D: mov ecx, var_4C
  loc_00E08960: push ecx
  loc_00E08961: push esi
  loc_00E08962: call [edi+000002A4h]
  loc_00E08968: test eax, eax
  loc_00E0896A: fnclex
  loc_00E0896C: jge 00E08984h
  loc_00E0896E: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E08974: push 000002A4h
  loc_00E08979: push 006DFCB8h
  loc_00E0897E: push esi
  loc_00E0897F: push eax
  loc_00E08980: call ebx
  loc_00E08982: jmp 00E0898Ah
  loc_00E08984: mov ebx, [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0898A: call [004010BCh] ; rtcDoEvents
  loc_00E08990: mov edx, [esi]
  loc_00E08992: lea eax, var_4C
  loc_00E08995: push eax
  loc_00E08996: push esi
  loc_00E08997: call [edx+00000070h]
  loc_00E0899A: test eax, eax
  loc_00E0899C: fnclex
  loc_00E0899E: jge 00E089ABh
  loc_00E089A0: push 00000070h
  loc_00E089A2: push 006DFCB8h
  loc_00E089A7: push esi
  loc_00E089A8: push eax
  loc_00E089A9: call ebx
  loc_00E089AB: mov eax, [00E3D9CCh]
  loc_00E089B0: test eax, eax
  loc_00E089B2: jnz 00E089C4h
  loc_00E089B4: push 00E3D9CCh
  loc_00E089B9: push 006DC960h
  loc_00E089BE: call [004011A0h] ; __vbaNew2
  loc_00E089C4: mov edi, [00E3D9CCh]
  loc_00E089CA: lea edx, var_18
  loc_00E089CD: push edx
  loc_00E089CE: push edi
  loc_00E089CF: mov ecx, [edi]
  loc_00E089D1: call [ecx+00000018h]
  loc_00E089D4: test eax, eax
  loc_00E089D6: fnclex
  loc_00E089D8: jge 00E089E5h
  loc_00E089DA: push 00000018h
  loc_00E089DC: push 006DC950h
  loc_00E089E1: push edi
  loc_00E089E2: push eax
  loc_00E089E3: call ebx
  loc_00E089E5: mov eax, var_18
  loc_00E089E8: lea edx, var_50
  loc_00E089EB: push edx
  loc_00E089EC: push eax
  loc_00E089ED: mov ecx, [eax]
  loc_00E089EF: mov edi, eax
  loc_00E089F1: call [ecx+00000098h]
  loc_00E089F7: test eax, eax
  loc_00E089F9: fnclex
  loc_00E089FB: jge 00E08A0Bh
  loc_00E089FD: push 00000098h
  loc_00E08A02: push 006DCAF0h
  loc_00E08A07: push edi
  loc_00E08A08: push eax
  loc_00E08A09: call ebx
  loc_00E08A0B: fld real4 ptr var_4C
  loc_00E08A0E: fcomp real4 ptr var_50
  loc_00E08A11: fnstsw ax
  loc_00E08A13: test ah, 41h
  loc_00E08A16: jz 00E08A1Fh
  loc_00E08A18: mov eax, 00000001h
  loc_00E08A1D: jmp 00E08A21h
  loc_00E08A1F: xor eax, eax
  loc_00E08A21: neg eax
  loc_00E08A23: lea ecx, var_18
  loc_00E08A26: mov edi, eax
  loc_00E08A28: call [00401254h] ; __vbaFreeObj
  loc_00E08A2E: test di, di
  loc_00E08A31: jnz 00E08881h
  loc_00E08A37: mov var_4, 00000000h
  loc_00E08A3E: fwait
  loc_00E08A3F: push 00E08A51h
  loc_00E08A44: jmp 00E08A50h
  loc_00E08A46: lea ecx, var_18
  loc_00E08A49: call [00401254h] ; __vbaFreeObj
  loc_00E08A4F: ret
  loc_00E08A50: ret
  loc_00E08A51: mov eax, Me
  loc_00E08A54: push eax
  loc_00E08A55: mov ecx, [eax]
  loc_00E08A57: call [ecx+00000008h]
  loc_00E08A5A: mov eax, var_4
  loc_00E08A5D: mov ecx, var_14
  loc_00E08A60: pop edi
  loc_00E08A61: pop esi
  loc_00E08A62: mov fs:[00000000h], ecx
  loc_00E08A69: pop ebx
  loc_00E08A6A: mov esp, ebp
  loc_00E08A6C: pop ebp
  loc_00E08A6D: retn 0008h
End Sub

Private Sub jcbutton1_UnknownEvent_9 'E09480
  loc_00E09480: push ebp
  loc_00E09481: mov ebp, esp
  loc_00E09483: sub esp, 0000000Ch
  loc_00E09486: push 00402836h ; __vbaExceptHandler
  loc_00E0948B: mov eax, fs:[00000000h]
  loc_00E09491: push eax
  loc_00E09492: mov fs:[00000000h], esp
  loc_00E09499: sub esp, 00000034h
  loc_00E0949C: push ebx
  loc_00E0949D: push esi
  loc_00E0949E: push edi
  loc_00E0949F: mov var_C, esp
  loc_00E094A2: mov var_8, 004020F0h
  loc_00E094A9: mov esi, Me
  loc_00E094AC: mov eax, esi
  loc_00E094AE: and eax, 00000001h
  loc_00E094B1: mov var_4, eax
  loc_00E094B4: and esi, FFFFFFFEh
  loc_00E094B7: push esi
  loc_00E094B8: mov Me, esi
  loc_00E094BB: mov ecx, [esi]
  loc_00E094BD: call [ecx+00000004h]
  loc_00E094C0: mov edx, [esi]
  loc_00E094C2: push esi
  loc_00E094C3: mov var_18, 00000000h
  loc_00E094CA: call [edx+00000310h]
  loc_00E094D0: mov edi, [004010ACh] ; __vbaObjSet
  loc_00E094D6: push eax
  loc_00E094D7: lea eax, var_18
  loc_00E094DA: push eax
  loc_00E094DB: call edi
  loc_00E094DD: mov ebx, eax
  loc_00E094DF: push 00000000h
  loc_00E094E1: push ebx
  loc_00E094E2: mov ecx, [ebx]
  loc_00E094E4: call [ecx+0000009Ch]
  loc_00E094EA: test eax, eax
  loc_00E094EC: fnclex
  loc_00E094EE: jge 00E09502h
  loc_00E094F0: push 0000009Ch
  loc_00E094F5: push 006DCAD0h
  loc_00E094FA: push ebx
  loc_00E094FB: push eax
  loc_00E094FC: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09502: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E09508: lea ecx, var_18
  loc_00E0950B: call ebx
  loc_00E0950D: sub esp, 00000010h
  loc_00E09510: mov ecx, 0000000Bh
  loc_00E09515: mov edx, esp
  loc_00E09517: or eax, FFFFFFFFh
  loc_00E0951A: push 80010007h
  loc_00E0951F: push esi
  loc_00E09520: mov [edx], ecx
  loc_00E09522: mov ecx, var_24
  loc_00E09525: mov [edx+00000004h], ecx
  loc_00E09528: mov ecx, [esi]
  loc_00E0952A: mov [edx+00000008h], eax
  loc_00E0952D: mov eax, var_1C
  loc_00E09530: mov [edx+0000000Ch], eax
  loc_00E09533: call [ecx+000003B4h]
  loc_00E09539: lea edx, var_18
  loc_00E0953C: push eax
  loc_00E0953D: push edx
  loc_00E0953E: call edi
  loc_00E09540: push eax
  loc_00E09541: call [00401238h] ; __vbaLateIdSt
  loc_00E09547: lea ecx, var_18
  loc_00E0954A: call ebx
  loc_00E0954C: mov eax, [esi]
  loc_00E0954E: push esi
  loc_00E0954F: call [eax+000003DCh]
  loc_00E09555: lea ecx, var_18
  loc_00E09558: push eax
  loc_00E09559: push ecx
  loc_00E0955A: call edi
  loc_00E0955C: mov ebx, eax
  loc_00E0955E: push FFFFFFFFh
  loc_00E09560: push ebx
  loc_00E09561: mov edx, [ebx]
  loc_00E09563: call [edx+0000008Ch]
  loc_00E09569: test eax, eax
  loc_00E0956B: fnclex
  loc_00E0956D: jge 00E09581h
  loc_00E0956F: push 0000008Ch
  loc_00E09574: push 006DCDA0h
  loc_00E09579: push ebx
  loc_00E0957A: push eax
  loc_00E0957B: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E09581: lea ecx, var_18
  loc_00E09584: call [00401254h] ; __vbaFreeObj
  loc_00E0958A: mov eax, [esi]
  loc_00E0958C: push esi
  loc_00E0958D: call [eax+000003E0h]
  loc_00E09593: lea ecx, var_18
  loc_00E09596: push eax
  loc_00E09597: push ecx
  loc_00E09598: call edi
  loc_00E0959A: mov ebx, eax
  loc_00E0959C: push 00000000h
  loc_00E0959E: push ebx
  loc_00E0959F: mov edx, [ebx]
  loc_00E095A1: call [edx+0000008Ch]
  loc_00E095A7: test eax, eax
  loc_00E095A9: fnclex
  loc_00E095AB: jge 00E095BFh
  loc_00E095AD: push 0000008Ch
  loc_00E095B2: push 006DCDA0h
  loc_00E095B7: push ebx
  loc_00E095B8: push eax
  loc_00E095B9: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E095BF: lea ecx, var_18
  loc_00E095C2: call [00401254h] ; __vbaFreeObj
  loc_00E095C8: mov eax, [esi]
  loc_00E095CA: push esi
  loc_00E095CB: call [eax+00000320h]
  loc_00E095D1: lea ecx, var_18
  loc_00E095D4: push eax
  loc_00E095D5: push ecx
  loc_00E095D6: call edi
  loc_00E095D8: mov ebx, eax
  loc_00E095DA: push 00000000h
  loc_00E095DC: push ebx
  loc_00E095DD: mov edx, [ebx]
  loc_00E095DF: call [edx+000000E4h]
  loc_00E095E5: test eax, eax
  loc_00E095E7: fnclex
  loc_00E095E9: jge 00E095FDh
  loc_00E095EB: push 000000E4h
  loc_00E095F0: push 006E03D4h
  loc_00E095F5: push ebx
  loc_00E095F6: push eax
  loc_00E095F7: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E095FD: lea ecx, var_18
  loc_00E09600: call [00401254h] ; __vbaFreeObj
  loc_00E09606: mov eax, [esi]
  loc_00E09608: push esi
  loc_00E09609: call [eax+0000031Ch]
  loc_00E0960F: lea ecx, var_18
  loc_00E09612: push eax
  loc_00E09613: push ecx
  loc_00E09614: call edi
  loc_00E09616: mov ebx, eax
  loc_00E09618: push 00000000h
  loc_00E0961A: push ebx
  loc_00E0961B: mov edx, [ebx]
  loc_00E0961D: call [edx+000000E4h]
  loc_00E09623: test eax, eax
  loc_00E09625: fnclex
  loc_00E09627: jge 00E0963Bh
  loc_00E09629: push 000000E4h
  loc_00E0962E: push 006E03D4h
  loc_00E09633: push ebx
  loc_00E09634: push eax
  loc_00E09635: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0963B: mov ebx, [00401254h] ; __vbaFreeObj
  loc_00E09641: lea ecx, var_18
  loc_00E09644: call ebx
  loc_00E09646: mov eax, [esi]
  loc_00E09648: push esi
  loc_00E09649: call [eax+00000324h]
  loc_00E0964F: lea ecx, var_18
  loc_00E09652: push eax
  loc_00E09653: push ecx
  loc_00E09654: call edi
  loc_00E09656: mov esi, eax
  loc_00E09658: push 006DCC80h
  loc_00E0965D: push esi
  loc_00E0965E: mov edx, [esi]
  loc_00E09660: call [edx+000000A4h]
  loc_00E09666: test eax, eax
  loc_00E09668: fnclex
  loc_00E0966A: jge 00E0967Eh
  loc_00E0966C: push 000000A4h
  loc_00E09671: push 006DCB70h
  loc_00E09676: push esi
  loc_00E09677: push eax
  loc_00E09678: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0967E: lea ecx, var_18
  loc_00E09681: call ebx
  loc_00E09683: mov var_4, 00000000h
  loc_00E0968A: push 00E0969Ch
  loc_00E0968F: jmp 00E0969Bh
  loc_00E09691: lea ecx, var_18
  loc_00E09694: call [00401254h] ; __vbaFreeObj
  loc_00E0969A: ret
  loc_00E0969B: ret
  loc_00E0969C: mov eax, Me
  loc_00E0969F: push eax
  loc_00E096A0: mov ecx, [eax]
  loc_00E096A2: call [ecx+00000008h]
  loc_00E096A5: mov eax, var_4
  loc_00E096A8: mov ecx, var_14
  loc_00E096AB: pop edi
  loc_00E096AC: pop esi
  loc_00E096AD: mov fs:[00000000h], ecx
  loc_00E096B4: pop ebx
  loc_00E096B5: mov esp, ebp
  loc_00E096B7: pop ebp
  loc_00E096B8: retn 0004h
End Sub

Private Sub nortu_KeyPress(KeyAscii As Integer) 'E0A450
  loc_00E0A450: push ebp
  loc_00E0A451: mov ebp, esp
  loc_00E0A453: sub esp, 0000000Ch
  loc_00E0A456: push 00402836h ; __vbaExceptHandler
  loc_00E0A45B: mov eax, fs:[00000000h]
  loc_00E0A461: push eax
  loc_00E0A462: mov fs:[00000000h], esp
  loc_00E0A469: sub esp, 00000014h
  loc_00E0A46C: push ebx
  loc_00E0A46D: push esi
  loc_00E0A46E: push edi
  loc_00E0A46F: mov var_C, esp
  loc_00E0A472: mov var_8, 004021C0h
  loc_00E0A479: mov esi, Me
  loc_00E0A47C: mov eax, esi
  loc_00E0A47E: and eax, 00000001h
  loc_00E0A481: mov var_4, eax
  loc_00E0A484: and esi, FFFFFFFEh
  loc_00E0A487: push esi
  loc_00E0A488: mov Me, esi
  loc_00E0A48B: mov ecx, [esi]
  loc_00E0A48D: call [ecx+00000004h]
  loc_00E0A490: mov edx, KeyAscii
  loc_00E0A493: xor edi, edi
  loc_00E0A495: mov var_18, edi
  loc_00E0A498: cmp [edx], 000Dh
  loc_00E0A49C: jnz 00E0A4DEh
  loc_00E0A49E: mov eax, [esi]
  loc_00E0A4A0: push esi
  loc_00E0A4A1: call [eax+00000364h]
  loc_00E0A4A7: lea ecx, var_18
  loc_00E0A4AA: push eax
  loc_00E0A4AB: push ecx
  loc_00E0A4AC: call [004010ACh] ; __vbaObjSet
  loc_00E0A4B2: mov esi, eax
  loc_00E0A4B4: push esi
  loc_00E0A4B5: mov edx, [esi]
  loc_00E0A4B7: call [edx+00000204h]
  loc_00E0A4BD: cmp eax, edi
  loc_00E0A4BF: fnclex
  loc_00E0A4C1: jge 00E0A4D5h
  loc_00E0A4C3: push 00000204h
  loc_00E0A4C8: push 006DCB70h
  loc_00E0A4CD: push esi
  loc_00E0A4CE: push eax
  loc_00E0A4CF: call [0040107Ch] ; __vbaHresultCheckObj
  loc_00E0A4D5: lea ecx, var_18
  loc_00E0A4D8: call [00401254h] ; __vbaFreeObj
  loc_00E0A4DE: mov var_4, edi
  loc_00E0A4E1: push 00E0A4F3h
  loc_00E0A4E6: jmp 00E0A4F2h
  loc_00E0A4E8: lea ecx, var_18
  loc_00E0A4EB: call [00401254h] ; __vbaFreeObj
  loc_00E0A4F1: ret
  loc_00E0A4F2: ret
  loc_00E0A4F3: mov eax, Me
  loc_00E0A4F6: push eax
  loc_00E0A4F7: mov ecx, [eax]
  loc_00E0A4F9: call [ecx+00000008h]
  loc_00E0A4FC: mov eax, var_4
  loc_00E0A4FF: mov ecx, var_14
  loc_00E0A502: pop edi
  loc_00E0A503: pop esi
  loc_00E0A504: mov fs:[00000000h], ecx
  loc_00E0A50B: pop ebx
  loc_00E0A50C: mov esp, ebp
  loc_00E0A50E: pop ebp
  loc_00E0A50F: retn 0008h
End Sub
