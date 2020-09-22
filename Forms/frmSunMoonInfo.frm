VERSION 5.00
Begin VB.Form frmSunMoonInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SunMoonInfo"
   ClientHeight    =   8292
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   12144
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   10.2
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSunMoonInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   691
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1012
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      Caption         =   "Sho&w Moon"
      Height          =   372
      Index           =   1
      Left            =   5040
      TabIndex        =   82
      Top             =   7800
      Width           =   1332
   End
   Begin VB.PictureBox picMoonInfo 
      BorderStyle     =   0  'None
      Height          =   7212
      Left            =   0
      ScaleHeight     =   601
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   541
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   6492
      Begin VB.ListBox lstMoonThisYear 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800080&
         Height          =   1620
         Left            =   1440
         TabIndex        =   62
         Top             =   5520
         Visible         =   0   'False
         Width           =   4932
      End
      Begin VB.ListBox lstMoonThisMonth 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800080&
         Height          =   1620
         Left            =   1440
         TabIndex        =   61
         Top             =   5520
         Width           =   4932
      End
      Begin VB.CommandButton cmdPhases 
         Caption         =   "Previous &Phases"
         Height          =   372
         Left            =   120
         TabIndex        =   48
         Top             =   2880
         Width           =   1932
      End
      Begin VB.TextBox txtMoonPhases 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800080&
         Height          =   972
         Left            =   2160
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2280
         Width           =   4212
      End
      Begin VB.PictureBox picMoon 
         BackColor       =   &H00000000&
         Height          =   2040
         Index           =   0
         Left            =   4320
         ScaleHeight     =   166
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   166
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   120
         Width           =   2040
         Begin VB.PictureBox picMoon 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   1092
            Index           =   1
            Left            =   240
            ScaleHeight     =   91
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   91
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   240
            Width           =   1092
         End
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "T&his month:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   30
         Left            =   120
         TabIndex        =   60
         Tag             =   "False"
         Top             =   5520
         Width           =   1212
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   10
         Left            =   2160
         TabIndex        =   59
         Top             =   5040
         Width           =   4212
      End
      Begin VB.Line linSunMoonInfo 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         Index           =   2
         X1              =   345
         X2              =   345
         Y1              =   10
         Y2              =   180
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Cycle:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   29
         Left            =   120
         TabIndex        =   58
         Top             =   5040
         Width           =   1812
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   9
         Left            =   2160
         TabIndex        =   57
         Top             =   4680
         Width           =   4212
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Age:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   28
         Left            =   120
         TabIndex        =   56
         Top             =   4680
         Width           =   1812
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Angle:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   27
         Left            =   120
         TabIndex        =   54
         Top             =   4320
         Width           =   1812
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Illumination:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   26
         Left            =   120
         TabIndex        =   52
         Top             =   3960
         Width           =   1812
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Date / Time:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   25
         Left            =   120
         TabIndex        =   50
         Top             =   3600
         Width           =   1812
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Position"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   24
         Left            =   2160
         TabIndex        =   49
         Top             =   3360
         Width           =   4212
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   8
         Left            =   2160
         TabIndex        =   55
         Top             =   4320
         Width           =   4212
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   7
         Left            =   2160
         TabIndex        =   53
         Top             =   3960
         Width           =   4212
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   6
         Left            =   2160
         TabIndex        =   51
         Top             =   3600
         Width           =   4212
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Phases:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   23
         Left            =   120
         TabIndex        =   46
         Top             =   2280
         Width           =   1812
      End
      Begin VB.Image imgMoonState 
         Height          =   384
         Index           =   2
         Left            =   1560
         Picture         =   "frmSunMoonInfo.frx":08CA
         Top             =   1140
         Width           =   384
      End
      Begin VB.Image imgMoonState 
         Height          =   384
         Index           =   1
         Left            =   1560
         Picture         =   "frmSunMoonInfo.frx":1594
         Top             =   780
         Width           =   384
      End
      Begin VB.Image imgMoonState 
         Height          =   384
         Index           =   0
         Left            =   1560
         Picture         =   "frmSunMoonInfo.frx":225E
         Top             =   420
         Width           =   384
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   5
         Left            =   2160
         TabIndex        =   43
         Top             =   1920
         Width           =   1812
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   4
         Left            =   2160
         TabIndex        =   41
         Top             =   1560
         Width           =   1812
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   3
         Left            =   2160
         TabIndex        =   39
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   2
         Left            =   2160
         TabIndex        =   37
         Top             =   840
         Width           =   1812
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   1
         Left            =   2160
         TabIndex        =   35
         Top             =   480
         Width           =   1812
      End
      Begin VB.Label lblMoonDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800080&
         Height          =   252
         Index           =   0
         Left            =   2160
         TabIndex        =   33
         Top             =   120
         Width           =   1812
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Age:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   22
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Width           =   1332
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moon Status:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   21
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   1332
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moonset:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   20
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moontransit:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   19
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Moonrise:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   18
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   17
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   1332
      End
   End
   Begin VB.CheckBox chkTimeInUTC 
      Caption         =   "Show time in &UTC"
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   120
      TabIndex        =   80
      Top             =   7860
      Width           =   2292
   End
   Begin VB.Timer tmrPosition 
      Interval        =   1000
      Left            =   2640
      Top             =   7800
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show Year"
      Height          =   372
      Index           =   0
      Left            =   3600
      TabIndex        =   81
      Top             =   7800
      Width           =   1332
   End
   Begin VB.TextBox txtTwilights 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   500
      Index           =   2
      Left            =   4440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1932
   End
   Begin VB.TextBox txtTwilights 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   500
      Index           =   1
      Left            =   4440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1932
   End
   Begin VB.TextBox txtTwilights 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   500
      Index           =   0
      Left            =   4440
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   840
      Width           =   1932
   End
   Begin VB.CheckBox chkSaveCityInfo 
      Caption         =   "&Keep changes in file: CityInfo.dat"
      ForeColor       =   &H00C00000&
      Height          =   312
      Left            =   6840
      TabIndex        =   75
      Top             =   7320
      Width           =   3612
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "&Quit"
      Height          =   372
      Index           =   0
      Left            =   11160
      TabIndex        =   79
      Top             =   7800
      Width           =   852
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "&Delete City"
      Height          =   372
      Index           =   3
      Left            =   9720
      TabIndex        =   78
      Top             =   7800
      Width           =   1332
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "&Change City"
      Height          =   372
      Index           =   2
      Left            =   8280
      TabIndex        =   77
      Top             =   7800
      Width           =   1332
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "&Add City"
      Height          =   372
      Index           =   1
      Left            =   6840
      TabIndex        =   76
      Top             =   7800
      Width           =   1332
   End
   Begin VB.TextBox txtCityInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   288
      Index           =   3
      Left            =   8160
      TabIndex        =   74
      Top             =   6840
      Width           =   3852
   End
   Begin VB.TextBox txtCityInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   288
      Index           =   2
      Left            =   8160
      TabIndex        =   72
      Top             =   6360
      Width           =   3852
   End
   Begin VB.TextBox txtCityInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   288
      Index           =   1
      Left            =   8160
      TabIndex        =   70
      Top             =   5880
      Width           =   3852
   End
   Begin VB.TextBox txtCityInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   288
      Index           =   0
      Left            =   8160
      TabIndex        =   68
      Top             =   5400
      Width           =   3852
   End
   Begin VB.ListBox lstCityInfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   3900
      Left            =   6840
      TabIndex        =   64
      Top             =   480
      Width           =   5172
   End
   Begin VB.ListBox lstSunThisYear 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   2760
      Left            =   1440
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   4932
   End
   Begin VB.ListBox lstSunThisMonth 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800080&
      Height          =   2760
      Left            =   1440
      TabIndex        =   29
      Top             =   4920
      Width           =   4932
   End
   Begin VB.ComboBox cmbCities 
      BackColor       =   &H80000009&
      ForeColor       =   &H00800080&
      Height          =   324
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   4932
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sun Elevation:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   15
      Left            =   120
      TabIndex        =   85
      Top             =   4440
      Width           =   1812
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   10
      Left            =   2160
      TabIndex        =   84
      Top             =   4440
      Width           =   4212
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sun Zenith:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   14
      Left            =   120
      TabIndex        =   83
      Top             =   4080
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sun Azimuth:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Top             =   3720
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Date / Time:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3360
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sun Position"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   11
      Left            =   2160
      TabIndex        =   22
      Top             =   3120
      Width           =   4212
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   9
      Left            =   2160
      TabIndex        =   27
      Top             =   4080
      Width           =   4212
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   8
      Left            =   2160
      TabIndex        =   25
      Top             =   3720
      Width           =   4212
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   7
      Left            =   2160
      TabIndex        =   23
      Top             =   3360
      Width           =   4212
   End
   Begin VB.Line linSunMoonInfo 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   350
      X2              =   350
      Y1              =   50
      Y2              =   251
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Astronomical Twilight:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   10
      Left            =   4440
      TabIndex        =   20
      Top             =   2280
      Width           =   1932
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Nautical Twilight:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   9
      Left            =   4440
      TabIndex        =   18
      Top             =   1440
      Width           =   1932
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Civil Twilight:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   8
      Left            =   4440
      TabIndex        =   16
      Top             =   600
      Width           =   1932
   End
   Begin VB.Image imgSunState 
      Height          =   384
      Index           =   3
      Left            =   1560
      Picture         =   "frmSunMoonInfo.frx":2F28
      Top             =   1980
      Width           =   384
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   6
      Left            =   2160
      TabIndex        =   15
      Top             =   2760
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sun Declination:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Time&Zone:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   99
      Left            =   6840
      TabIndex        =   73
      Top             =   6864
      Width           =   1212
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "L&ongitude:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   98
      Left            =   6840
      TabIndex        =   71
      Top             =   6384
      Width           =   1212
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "&Latitude:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   97
      Left            =   6840
      TabIndex        =   69
      Top             =   5904
      Width           =   1212
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "City na&me:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   96
      Left            =   6840
      TabIndex        =   67
      Top             =   5424
      Width           =   1212
   End
   Begin VB.Line linSunMoonInfo 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   550
      X2              =   550
      Y1              =   10
      Y2              =   681
   End
   Begin VB.Label lblCityCount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Left            =   11040
      TabIndex        =   66
      Top             =   120
      Width           =   972
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of cities:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   95
      Left            =   9240
      TabIndex        =   65
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "I&nfo for all cities:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   94
      Left            =   6840
      TabIndex        =   63
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   5
      Left            =   2160
      TabIndex        =   13
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   4
      Left            =   2160
      TabIndex        =   11
      Top             =   2040
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "T&his month:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   16
      Left            =   120
      TabIndex        =   28
      Tag             =   "False"
      Top             =   4920
      Width           =   1212
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Equation of Time:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1812
   End
   Begin VB.Image imgSunState 
      Height          =   384
      Index           =   2
      Left            =   1560
      Picture         =   "frmSunMoonInfo.frx":37F2
      Top             =   1620
      Width           =   384
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Time of Sun:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1332
   End
   Begin VB.Image imgSunState 
      Height          =   384
      Index           =   1
      Left            =   1560
      Picture         =   "frmSunMoonInfo.frx":40BC
      Top             =   1260
      Width           =   384
   End
   Begin VB.Image imgSunState 
      Height          =   384
      Index           =   0
      Left            =   1560
      Picture         =   "frmSunMoonInfo.frx":4986
      Top             =   900
      Width           =   384
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Suntransit:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "C&ity:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   160
      Width           =   1212
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   3
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1812
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sunset:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1332
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sunrise:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label lblSunDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800080&
      Height          =   252
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   1812
   End
End
Attribute VB_Name = "frmSunMoonInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Class
Private WithEvents SunMoonInfo As clsSunMoonInfo
Attribute SunMoonInfo.VB_VarHelpID = -1

' Private Enumerations
Private Enum RoundTypes
   RoundUp
   RoundDown
   RoundNearest
End Enum

Private Enum TimeTypes
   Seconds
   Minutes
   MinutesSeconds
End Enum

' Private Variables
Private FillListBox   As Boolean
Private PrevousPhases As Boolean
Private ResultListBox As ListBox
Private CityIndex     As Long
Private FocusObject   As Object
Private CityName      As String
Private DateFormat    As String

' Private API's
Private Declare Function GetLocaleInfo Lib "Kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetSystemDefaultLCID Lib "Kernel32" () As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As String, lParam As Any) As Long

Private Function CheckInput() As String

Dim intItem    As Integer
Dim strCaption As String
Dim strInput   As String

   For intItem = 0 To 3
      If txtCityInfo.Item(intItem).Text = "" Then
         If strInput <> "" Then strInput = strInput & vbCrLf
         
         strCaption = lblInfo.Item(11 + intItem).Caption
         strInput = strInput & Left(strCaption, Len(strCaption) - 1) & " is empty!"
      End If
   Next 'intItem
   
   CheckInput = strInput

End Function

Private Function GenerateCycleText(ByVal IsDate As Date, ByVal Phase As Integer, ByVal MoonIllumination As Double) As String

Dim strCycleText As String

   With SunMoonInfo
      If .MoonDescription = WaxingGibbous Then
         strCycleText = "Waxing Gibbous"
         
      ElseIf .MoonDescription = WaningGibbous Then
         strCycleText = "Waning Gibbous"
         
      ElseIf .MoonDescription = WaningCrescent Then
         strCycleText = "Waning Crescent"
         
      ElseIf .MoonDescription = WaxingCrescent Then
         strCycleText = "Waxing Crescent"
      End If
      
      If Phase < 90 Then
         .DateCalculate = IsDate
         
         If Format(MoonIllumination, "###") = "" Then
            strCycleText = strCycleText & " 100% of New"
            
         Else
            strCycleText = strCycleText & " " & Format(MoonIllumination, "###") & "% of Full"
         End If
         
         .DateCalculate = Now
         
      Else
         Phase = (100 - MoonIllumination)
         
         If Format(GetPercentage(Phase), "###") = "" Then
            strCycleText = strCycleText & " 100% of Full"
            
         Else
            strCycleText = strCycleText & " " & Format(GetPercentage(Phase), "###") & "% of New"
         End If
      End If
   End With
   
   GenerateCycleText = strCycleText & " Moon"

End Function

Private Function GetDateFormat() As String

Const LOCALE_IDATE As Long = &H21
Const LOCALE_SDATE As Long = &H1D

Dim lngLocaleID    As Long
Dim strBuffer      As String * 256
Dim strDateFormat  As String
Dim strSeparator   As String

   lngLocaleID = GetSystemDefaultLCID
   
   If GetLocaleInfo(lngLocaleID, LOCALE_SDATE, strBuffer, Len(strBuffer)) Then
      strSeparator = Left(strBuffer, 1)
      
   Else
      strSeparator = "-"
   End If
   
   If GetLocaleInfo(lngLocaleID, LOCALE_IDATE, strBuffer, Len(strBuffer)) Then
      Select Case Val(Left(strBuffer, 1))
         Case 0
            strDateFormat = "mm_dd_yyyy"
            
         Case 1
            strDateFormat = "dd_mm_yyyy"
            
         Case 2
            strDateFormat = "yyyy_mm_dd"
      End Select
      
   Else
      strDateFormat = "dd_mm_yyyy"
   End If
   
   GetDateFormat = Replace(strDateFormat, "_", strSeparator)

End Function

Private Function GetListIndex(ByVal hWnd As Long, ByVal City As String) As Integer

Const LB_FINDSTRING As Long = &H18F

   GetListIndex = SendMessage(hWnd, LB_FINDSTRING, 0, ByVal "Name:" & vbTab & vbTab & City)

End Function

Private Function GetMoonInfo(ByVal MoonInfo As MoonStatus) As String

Dim strInfo As String

   If MoonInfo = None Then
      strInfo = "None"
      
   ElseIf MoonInfo = AboveHorizon Then
      strInfo = "Above horizon"
      
   ElseIf MoonInfo = BelowHorizon Then
      strInfo = "Below horizon"
      
   ElseIf MoonInfo = Up Then
      strInfo = "Up"
      
   ElseIf MoonInfo = Down Then
      strInfo = "Down"
      
   ElseIf MoonInfo = Rising Then
      strInfo = "Rising"
      
   ElseIf MoonInfo = Setting Then
      strInfo = "Setting"
      
   ElseIf MoonInfo = UpAllDay Then
      strInfo = "Up all day"
      
   ElseIf MoonInfo = DownAllDay Then
      strInfo = "Down all day"
   End If
   
   GetMoonInfo = strInfo

End Function

Private Function GetPercentage(ByVal Percentage As Integer) As Single

Dim sngPercentage As Single

   On Local Error GoTo ExitFunction
   
   If Percentage = 0 Then
      sngPercentage = 100
      
   ElseIf Percentage > 100 Then
      sngPercentage = 0
      
   Else
      sngPercentage = (Percentage * 100) / 100
   End If
   
ExitFunction:
   On Local Error GoTo 0
   GetPercentage = sngPercentage

End Function

Private Function QuickSort(ByRef DataArray() As String, ByVal LowerBound As Long, ByVal UpperBound As Long)

Dim intItem   As Integer
Dim lngLower  As Long
Dim lngUpper  As Long
Dim strBuffer As String
Dim strMiddle As String

   If UpperBound <= LowerBound Then Exit Function
   
   lngLower = LowerBound
   lngUpper = UpperBound
   strMiddle = DataArray((LowerBound + UpperBound) \ 2, 0)
   
   Do While (lngLower <= lngUpper)
      Do While DataArray(lngLower, 0) < strMiddle
         lngLower = lngLower + 1
         
         If lngLower = UpperBound Then Exit Do
      Loop
      
      Do While strMiddle < DataArray(lngUpper, 0)
         lngUpper = lngUpper - 1
         
         If lngUpper = LowerBound Then Exit Do
      Loop
      
      If lngLower <= lngUpper Then
         For intItem = 0 To 1
            strBuffer = DataArray(lngLower, intItem)
            DataArray(lngLower, intItem) = DataArray(lngUpper, intItem)
            DataArray(lngUpper, intItem) = strBuffer
         Next 'intItem
         
         lngLower = lngLower + 1
         lngUpper = lngUpper - 1
      End If
   Loop
   
   If LowerBound < lngUpper Then QuickSort DataArray(), LowerBound, lngUpper
   If lngLower < UpperBound Then QuickSort DataArray(), lngLower, UpperBound

End Function

Private Function RoundTime(ByVal IsTime As Date, ByVal TimeType As TimeTypes, Optional ByVal RoundType As RoundTypes = RoundNearest, Optional ByVal IntoQuarters As Boolean) As String

Dim intHour     As Integer
Dim intMinute   As Integer
Dim intMultiply As Integer
Dim intRound    As Integer
Dim intSecond   As Integer

   IsTime = Format(IsTime, "hh:mm:ss")
   intHour = Hour(IsTime)
   intMinute = Minute(IsTime)
   intSecond = Second(IsTime)
   
   If IntoQuarters Then
      If TimeType <> Minutes Then
         If RoundType = RoundUp Then
            Select Case intSecond
               Case Is > 45
                  intSecond = 0
                  intMinute = intMinute + 1
                  
               Case Is > 30
                  intSecond = 45
                  
               Case Is > 15
                  intSecond = 30
                  
               Case Else
                  intSecond = 15
            End Select
            
         ElseIf RoundType = RoundDown Then
            Select Case intSecond
               Case Is < 15
                  intSecond = 0
                  
               Case Is < 30
                  intSecond = 15
                  
               Case Is < 45
                  intSecond = 30
                  
               Case Else
                  intSecond = 45
            End Select
            
         ' RoundType = RoundNearest
         Else
            intMultiply = intSecond \ 15
            intSecond = 15 And ((intSecond - intMultiply * 15) > 7)
            intSecond = intSecond + intMultiply * 15
            
            If intSecond > 59 Then
               intSecond = intSecond - 60
               intMinute = intMinute + 1
            End If
         End If
      End If
      
      If TimeType <> Seconds Then
         If RoundType = RoundUp Then
            Select Case intMinute
               Case Is > 45
                  intMinute = 0
                  intHour = intHour + 1
                  
               Case Is > 30
                  intMinute = 45
                  
               Case Is > 15
                  intMinute = 30
                  
               Case Else
                  intMinute = 15
            End Select
            
         ElseIf RoundType = RoundDown Then
            Select Case intMinute
               Case Is < 15
                  intMinute = 0
                  
               Case Is < 30
                  intMinute = 15
                  
               Case Is < 45
                  intMinute = 30
                  
               Case Else
                  intMinute = 45
            End Select
            
         ' RoundType = RoundNearest
         Else
            intMultiply = intMinute \ 15
            intMinute = 15 And ((intMinute - intMultiply * 15) > 7)
            intMinute = intMinute + intMultiply * 15
            
            If intMinute > 59 Then
               intMinute = intMinute - 60
               intHour = intHour + 1
            End If
         End If
      End If
      
   Else
      If TimeType <> Minutes Then
         intRound = Round((intSecond / 0.6) / 100 + 0.005, 3)
         
         If RoundType = RoundUp Then
            If intRound Then
               intMinute = intMinute + 1
               intSecond = 0
            End If
            
         ElseIf RoundType = RoundDown Then
            If intRound = 0 Then intSecond = 0
            
         ' RoundType = RoundNearest
         ElseIf intRound Then
            intMinute = intMinute + 1
            intSecond = 0
            
         Else
            intSecond = 0
         End If
      End If
      
      If TimeType <> Seconds Then
         intRound = Round((intMinute / 0.6) / 100 + 0.005, 3)
         
         If RoundType = RoundUp Then
            If intRound Then
               intHour = intHour + 1
               intMinute = 0
            End If
            
         ElseIf RoundType = RoundDown Then
            If intRound = 0 Then intMinute = 0
            
         ' RoundType = RoundNearest
         ElseIf intRound Then
            intHour = intHour + 1
            intMinute = 0
            
         Else
            intMinute = 0
         End If
      End If
   End If
   
   RoundTime = Format(TimeSerial(intHour, intMinute, intSecond), "hh:mm:ss")

End Function

Private Sub CalculateSunMoon()

   FillListBox = False
   
   With SunMoonInfo
      .TimeInUTC = CBool(chkTimeInUTC.Value)
      .DateCalculate = Now
      
      Call .SunCalculate
      Call .MoonPhases(PrevousPhases)
      Call .MoonCalculate
   End With

End Sub

Private Sub FillCaption(ByVal Index As Integer, ByVal ShowYear As Boolean)

Dim strButton As String
Dim strLabel  As String

   If ShowYear Then
      strButton = "Month"
      strLabel = "year"
      
   Else
      strButton = "Year"
      strLabel = "month"
   End If
   
   With lblInfo.Item(Index)
      .Caption = "T&his " & strLabel & ":"
      .Tag = ShowYear
   End With
   
   cmdShow.Item(0).Caption = "&Show " & strButton

End Sub

Private Sub FillBox(ByRef Box As ListBox, ByVal FillSun As Boolean, ByVal Year As Integer, ByVal Month As Integer, ByVal MaxDays As Integer)

Dim intDay As Integer

   FillListBox = True
   Set ResultListBox = Box
   
   With SunMoonInfo
      For intDay = 1 To MaxDays
         .DateCalculate = Format(DateSerial(Year, Month, intDay), "dd/mm/yyyy")
         
         If FillSun Then
            Call .SunCalculate
            
         Else
            Call .MoonCalculate
         End If
      Next 'intDay
      
      .DateCalculate = Now
   End With
   
   FillListBox = False
   Set ResultListBox = Nothing

End Sub

Private Sub FillThisMonth(ByVal IsDate As Date, ByVal FillSun As Boolean)

Dim intDay(1)   As Integer
Dim intMonth(1) As Integer
Dim intYear(1)  As Integer
Dim lstBox      As ListBox

   If FillSun Then
      Set lstBox = lstSunThisMonth
      
   Else
      Set lstBox = lstMoonThisMonth
   End If
   
   lstBox.Clear
   intYear(0) = Year(IsDate)
   intYear(1) = intYear(0)
   intMonth(0) = Month(IsDate)
   intMonth(1) = intMonth(0) + 1
   
   If intMonth(1) > 12 Then
      intMonth(1) = 1
      intYear(1) = intYear(1) + 1
   End If
   
   intDay(1) = DateDiff("d", DateSerial(intYear(0), intMonth(0), 0), DateSerial(intYear(1), intMonth(1), 0))
   
   Call FillBox(lstBox, FillSun, intYear(0), intMonth(0), intDay(1))
   Call SetTopIndex(lstBox, IsDate)
   
   Set lstBox = Nothing
   Erase intDay, intMonth, intYear

End Sub

Private Sub FillThisYear(ByVal IsDate As Date, ByVal FillSun As Boolean)

Dim intDay(1)   As Integer
Dim intMonth(1) As Integer
Dim intYear(1)  As Integer
Dim lstBox      As ListBox

   If FillSun Then
      Set lstBox = lstSunThisYear
      
   Else
      Set lstBox = lstMoonThisYear
   End If
   
   lstBox.Clear
   intYear(0) = Year(IsDate)
   intYear(1) = intYear(0)
   
   For intMonth(0) = 1 To 12
      intMonth(1) = intMonth(0) + 1
      
      If intMonth(1) > 12 Then
         intMonth(1) = 1
         intYear(1) = intYear(1) + 1
      End If
      
      intDay(1) = DateDiff("d", DateSerial(intYear(0), intMonth(0), 0), DateSerial(intYear(1), intMonth(1), 0))
      
      Call FillBox(lstBox, FillSun, intYear(0), intMonth(0), intDay(1))
   Next 'intMonth(0)
   
   Call SetTopIndex(lstBox, IsDate)
   
   Set lstBox = Nothing
   Erase intDay, intMonth, intYear

End Sub

Private Sub RefreshCityInfo()

Dim lngCount As Long

   With SunMoonInfo
      cmbCities.Clear
      lstCityInfo.Clear
      
      For lngCount = 0 To .CityCount - 1
         cmbCities.AddItem .CityName(lngCount)
      Next 'lngCount
      
      If lngCount Then
         FillListBox = True
         .CityGet , True
         lstCityInfo.RemoveItem lstCityInfo.ListCount - 1
         cmbCities.ListIndex = .CityIndex(CityName)
         
         With lstCityInfo
            .ListIndex = GetListIndex(.hWnd, CityName)
            
            If .ListIndex > -1 Then .TopIndex = .ListIndex
         End With
         
      Else
         For lngCount = 0 To 5
            lblSunDate.Item(lngCount).Caption = ""
            
            If lngCount < 4 Then txtCityInfo.Item(lngCount).Text = ""
         Next 'lngCount
         
         lstSunThisMonth.Clear
         lstSunThisYear.Clear
      End If
      
      lblCityCount.Caption = .CityCount
   End With

End Sub

Private Sub SetTopIndex(ByRef Box As ListBox, ByVal IsDate As Date)

Dim intCount As Integer
Dim strDate  As String

   With Box
      .RemoveItem .ListCount - 1
      strDate = Format(IsDate, DateFormat)
      
      For intCount = 0 To .ListCount - 1
         If InStr(.List(intCount), strDate) Then
            .ListIndex = intCount
            .TopIndex = intCount
            Exit For
         End If
      Next 'intCount
   End With

End Sub

Private Sub chkSaveCityInfo_LostFocus()

   Set FocusObject = chkSaveCityInfo

End Sub

Private Sub chkTimeInUTC_Click()

   Call CalculateSunMoon

End Sub

Private Sub chkTimeInUTC_LostFocus()

   Set FocusObject = chkTimeInUTC

End Sub

Private Sub cmbCities_Click()

   With SunMoonInfo
      If .CitySet(cmbCities.Text) Then
         With lstCityInfo
            .ListIndex = GetListIndex(.hWnd, cmbCities.Text)
            
            If .ListIndex > -1 Then .TopIndex = .ListIndex
         End With
         
         FillListBox = False
         CityName = cmbCities.Text
         .CityGet CityName
         
         Call CalculateSunMoon
      End If
   End With

End Sub

Private Sub cmbCities_LostFocus()

   Set FocusObject = cmbCities

End Sub

Private Sub cmdChoose_Click(Index As Integer)

Dim blnSave    As Boolean
Dim strProcess As String
Dim varResult  As Variant

   blnSave = CBool(chkSaveCityInfo.Value)
   
   Select Case Index
      Case 0
         Unload Me
         Exit Sub
         
      Case 1
         With txtCityInfo
            varResult = CheckInput
            
            If varResult = Empty Then varResult = SunMoonInfo.CityAdd(.Item(0).Text, .Item(1).Text, .Item(2).Text, CDbl(SunMoonInfo.ValidateValue(.Item(3).Text)), blnSave)
         End With
         
         If IsNumeric(varResult) Then
            CityIndex = varResult
            CityName = SunMoonInfo.CityName(CityIndex)
            
         Else
            strProcess = "Add"
         End If
         
      Case 2
         With txtCityInfo
            If SunMoonInfo.CityCount = 0 Then
               varResult = "Error - No more cities available!"
               
            ElseIf CityIndex > -1 Then
               varResult = CheckInput
               
               If varResult = Empty Then varResult = SunMoonInfo.CityChange(CityIndex, .Item(0).Text, .Item(1).Text, .Item(2).Text, CDbl(SunMoonInfo.ValidateValue(.Item(3).Text)), blnSave)
               
            Else
               varResult = "Error - No city selected!"
            End If
         End With
         
         If IsNumeric(varResult) Then
            CityName = SunMoonInfo.CityName(CityIndex)
            
         Else
            strProcess = "Change"
         End If
         
      Case 3
         ' Delete selected city
         varResult = SunMoonInfo.CityDelete(CityIndex, blnSave)
         
         If varResult = True Then
            CityName = ""
            CityIndex = -1
            
         Else
            strProcess = "Delete"
         End If
   End Select
   
   If strProcess = "" Then
      Call RefreshCityInfo
      
   Else
      MsgBox varResult, vbOKOnly + vbExclamation, "CityInfo - " & strProcess
   End If

End Sub

Private Sub cmdChoose_LostFocus(Index As Integer)

   Set FocusObject = cmdChoose.Item(Index)

End Sub

Private Sub cmdPhases_Click()

   With cmdPhases
      If PrevousPhases Then
         .Caption = "Previous &Phases"
         
      Else
         .Caption = "Next &Phases"
      End If
   End With
   
   PrevousPhases = Not PrevousPhases
   
   Call SunMoonInfo.MoonPhases(PrevousPhases)

End Sub

Private Sub cmdPhases_LostFocus()

   Set FocusObject = cmdPhases

End Sub

Private Sub cmdShow_Click(Index As Integer)

Dim blnShowYear As Boolean

   If Index Then
      If picMoonInfo.Visible Then
         cmdShow.Item(1).Caption = "Sho&w Moon"
         blnShowYear = CBool(lblInfo.Item(16).Tag)
         
         Call FillCaption(16, blnShowYear)
         
      Else
         cmdShow.Item(1).Caption = "Sho&w Sun"
         blnShowYear = CBool(lblInfo.Item(30).Tag)
         
         Call FillCaption(30, blnShowYear)
      End If
      
      lblInfo.Item(16).Visible = Not lblInfo.Item(16).Visible
      picMoonInfo.Visible = Not picMoonInfo.Visible
      
   ElseIf picMoonInfo.Visible Then
      If lstMoonThisMonth.Visible Then
         lstMoonThisYear.Visible = True
         lstMoonThisMonth.Visible = False
         
         Call FillCaption(30, True)
        
      Else
         lstMoonThisMonth.Visible = True
         lstMoonThisYear.Visible = False
         
         Call FillCaption(30, False)
      End If
      
   Else
      If lstSunThisMonth.Visible Then
         lstSunThisYear.Visible = True
         lstSunThisMonth.Visible = False
         
         Call FillCaption(16, True)
         
      Else
         lstSunThisMonth.Visible = True
         lstSunThisYear.Visible = False
         
         Call FillCaption(16, False)
      End If
   End If

End Sub

Private Sub cmdShow_LostFocus(Index As Integer)

   Set FocusObject = cmdShow.Item(Index)

End Sub

Private Sub lstCityInfo_DblClick()

   With lstCityInfo
      FillListBox = False
      SunMoonInfo.CityGet .ItemData(.ListIndex)
      .ListIndex = GetListIndex(.hWnd, txtCityInfo.Item(0).Text)
      
      If .ListIndex > -1 Then .TopIndex = .ListIndex
   End With

End Sub

Private Sub Form_Load()

   Set SunMoonInfo = New clsSunMoonInfo
   CityName = "Amsterdam"
   chkSaveCityInfo.ToolTipText = "Map: " & SunMoonInfo.MapCityInfo
   DateFormat = GetDateFormat
   Show
   DoEvents
   
   Call RefreshCityInfo

End Sub

Private Sub lstMoonThisMonth_LostFocus()

   Set FocusObject = lstMoonThisMonth

End Sub

Private Sub lstMoonThisYear_LostFocus()

   Set FocusObject = lstMoonThisYear

End Sub

Private Sub lstSunThisMonth_LostFocus()

   Set FocusObject = lstSunThisMonth

End Sub

Private Sub lstSunThisYear_LostFocus()

   Set FocusObject = lstSunThisYear

End Sub

Private Sub picMoon_GotFocus(Index As Integer)

   FocusObject.SetFocus

End Sub

Private Sub SunMoonInfo_ResultCityInfo(Name As String, Index As Long, Latitude As String, Longitude As String, TimeZone As Double)

   If FillListBox Then
      With lstCityInfo
         .AddItem "Name:" & vbTab & vbTab & Name
         .ItemData(.NewIndex) = Index
         .AddItem "Index:" & vbTab & vbTab & Index
         .ItemData(.NewIndex) = Index
         .AddItem "Latitude:" & vbTab & vbTab & Latitude
         .ItemData(.NewIndex) = Index
         .AddItem "Longitude:" & vbTab & Longitude
         .ItemData(.NewIndex) = Index
         .AddItem "TimeZone:" & vbTab & TimeZone
         .ItemData(.NewIndex) = Index
         .AddItem ""
      End With
      
   Else
      With txtCityInfo
         CityIndex = Index
         .Item(0).Text = Name
         .Item(1).Text = Latitude
         .Item(2).Text = Longitude
         .Item(3).Text = TimeZone
      End With
   End If

End Sub

Private Sub SunMoonInfo_ResultMoonCalculate(IsDate As Date, Moonrise As Date, Moontransit As Date, Moonset As Date, MoonsetInfo As MoonStatus, MoonriseInfo As MoonStatus, MoonStatus As MoonStatus, MoonAge As Double)

Dim strMoonrise   As String
Dim strMoonset    As String
Dim strMoonStatus As String

   If MoonriseInfo = Done Then
      If FillListBox Then
         strMoonrise = Format(Moonrise, "dd-mm-yyyy - hh:mm:ss")
         
      Else
         strMoonrise = Format(RoundTime(Format(Moonrise, "hh:mm:ss"), Seconds), "hh:mm")
      End If
      
   Else
      strMoonrise = GetMoonInfo(MoonriseInfo)
   End If
   
   If MoonsetInfo = Done Then
      If FillListBox Then
         strMoonset = Format(Moonset, "dd-mm-yyyy - hh:mm:ss")
         
      Else
         strMoonset = Format(RoundTime(Format(Moonset, "hh:mm:ss"), Seconds), "hh:mm")
      End If
      
   Else
      strMoonset = GetMoonInfo(MoonsetInfo)
   End If
   
   strMoonStatus = GetMoonInfo(MoonStatus)
   
   If FillListBox Then
      With ResultListBox
         .AddItem "Moonrise:" & vbTab & vbTab & strMoonrise
         .AddItem "Moontransit:" & vbTab & vbTab & Format(Moontransit, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Moonset:" & vbTab & vbTab & vbTab & strMoonset
         
         Call SunMoonInfo.MoonPhases(False)
         
         .AddItem "Moon Status:" & vbTab & vbTab & strMoonStatus
         .AddItem "Moon Age:" & vbTab & vbTab & Fix(MoonAge)
      End With
      
      Call SunMoonInfo.MoonPosition(DateValue(IsDate) & " " & TimeValue("12:00:00"))
      
   Else
      With lblMoonDate
         .Item(0).Caption = Format(IsDate, DateFormat)
         .Item(1).Caption = strMoonrise
         .Item(2).Caption = Format(Moontransit, "hh:mm:ss")
         .Item(3).Caption = strMoonset
         .Item(4).Caption = strMoonStatus
         .Item(5).Caption = Fix(MoonAge)
      End With
      
      Call tmrPosition_Timer
      Call FillThisMonth(Now, False)
      Call FillThisYear(Now, False)
   End If

End Sub

Private Sub SunMoonInfo_ResultMoonPhases(IsDate As Date, NewMoon As Date, FirstQuarter As Date, FullMoon As Date, LastQuarter As Date, PreviousPhases As Boolean)

Dim intBegin       As Integer
Dim intEnd         As Integer
Dim intPhases      As Integer
Dim intStep        As Integer
Dim strPhase(3, 1) As String
Dim strText        As String

   strPhase(0, 0) = CDbl(NewMoon)
   strPhase(0, 1) = "New Moon:"
   strPhase(1, 0) = CDbl(FirstQuarter)
   strPhase(1, 1) = "First Quarter:"
   strPhase(2, 0) = CDbl(FullMoon)
   strPhase(2, 1) = "Full Moon:"
   strPhase(3, 0) = CDbl(LastQuarter)
   strPhase(3, 1) = "Last Quarter:"
   
   If PreviousPhases Then
      strText = "Previous "
      intBegin = 3
      intEnd = 0
      intStep = -1
      
   Else
      strText = "Next "
      intBegin = 0
      intEnd = 3
      intStep = 1
   End If
   
   QuickSort strPhase(), 0, 3
   
   If FillListBox Then
      ResultListBox.AddItem strText & strPhase(0, 1) & vbTab & vbTab & Format(strPhase(0, 0), "dd-mm-yyyy - hh:mm:ss")
      
   Else
      lblInfo.Item(23).Caption = strText & "Phases:"
      
      With txtMoonPhases
         .Text = ""
         
         For intPhases = intBegin To intEnd Step intStep
            .Text = .Text & "  " & strPhase(intPhases, 1) & vbTab & "     " & Format(strPhase(intPhases, 0), "dd-mm-yyyy - hh:mm:ss") & vbCrLf
         Next 'intPhases
      End With
   End If
   
   Erase strPhase

End Sub

Private Sub SunMoonInfo_ResultMoonPosition(IsDate As Date, MoonIllumination As Double, MoonAngle As Double, MoonAge As Double)

Dim intPhase As Integer
Dim strMoon  As String

   intPhase = MoonAngle / 2
   strMoon = GenerateCycleText(IsDate, intPhase, MoonIllumination)
   
   If FillListBox Then
      With ResultListBox
         .AddItem "Moon Cycle:    " & strMoon
         .AddItem ""
      End With
      
   Else
      If intPhase > 179 Then intPhase = intPhase - 179
      
      With picMoon.Item(1)
         .Picture = LoadResPicture(intPhase, vbResBitmap)
         .Top = (picMoon.Item(0).ScaleHeight - .ScaleHeight) / 2
         .Left = (picMoon.Item(0).ScaleWidth - .ScaleWidth) / 2
         
         If SunMoonInfo.Hemisphere = South Then
            .PaintPicture .Picture, .ScaleWidth, 0, -.ScaleWidth, .ScaleHeight, 0, 0, .ScaleWidth, .ScaleHeight, vbSrcCopy
         End If
      End With
      
      With lblMoonDate
         .Item(6).Caption = Format(IsDate, DateFormat) & " / " & Format(IsDate, "hh:mm:ss")
         .Item(7).Caption = MoonIllumination
         .Item(8).Caption = MoonAngle
         .Item(9).Caption = MoonAge
         .Item(10).Caption = strMoon
      End With
   End If

End Sub

Private Sub SunMoonInfo_ResultSunCalculate(IsDate As Date, Sunrise As Date, Suntransit As Date, Sunset As Date, SunTime As Date, CivilTwilightBegin As Date, CivilTwilightEnd As Date, NauticalTwilightBegin As Date, NauticalTwilightEnd As Date, AstronomicalTwilightBegin As Date, AstronomicalTwilightEnd As Date, EquationOfTime As Double, SunDeclination As Double)

Dim strTwilights As String

   If FillListBox Then
      With ResultListBox
         .AddItem "Sunrise:" & vbTab & vbTab & vbTab & Format(Sunrise, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Suntransit:" & vbTab & vbTab & Format(Suntransit, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Sunset:" & vbTab & vbTab & vbTab & Format(Sunset, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Time of Sun:" & vbTab & vbTab & Format(SunTime, "hh:mm:ss")
         .AddItem "Civil twilight begin:" & vbTab & vbTab & Format(CivilTwilightBegin, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Civil twilight end:" & vbTab & vbTab & Format(CivilTwilightEnd, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Nautical twilight begin:" & vbTab & Format(NauticalTwilightBegin, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Nautical twilight end:" & vbTab & Format(NauticalTwilightEnd, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Astronomical twilight begin:" & vbTab & Format(AstronomicalTwilightBegin, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Astronomical twilight end:" & vbTab & Format(AstronomicalTwilightEnd, "dd-mm-yyyy - hh:mm:ss")
         .AddItem "Equation of Time:" & vbTab & vbTab & EquationOfTime
         .AddItem "Sun Declination:" & vbTab & vbTab & SunDeclination
         .AddItem ""
      End With
      
   Else
      With lblSunDate
         .Item(0).Caption = Format(IsDate, DateFormat)
         .Item(1).Caption = Format(RoundTime(Format(Sunrise, "hh:mm:ss"), Seconds), "hh:mm")
         .Item(2).Caption = Format(Suntransit, "hh:mm:ss")
         .Item(3).Caption = Format(RoundTime(Format(Sunset, "hh:mm:ss"), Seconds), "hh:mm")
         .Item(4).Caption = Format(SunTime, "hh:mm:ss")
         .Item(5).Caption = Round(EquationOfTime, 4)
         .Item(6).Caption = Round(SunDeclination, 4)
      End With
      
      With txtTwilights
         strTwilights = "Begin:" & vbTab & Format(RoundTime(Format(CivilTwilightBegin, "hh:mm:ss"), Seconds), "hh:mm") & vbCrLf
         strTwilights = strTwilights & "End:" & vbTab & Format(RoundTime(Format(CivilTwilightEnd, "hh:mm:ss"), Seconds), "hh:mm")
         .Item(0).Text = strTwilights
         strTwilights = "Begin:" & vbTab & Format(RoundTime(Format(NauticalTwilightBegin, "hh:mm:ss"), Seconds), "hh:mm") & vbCrLf
         strTwilights = strTwilights & "End:" & vbTab & Format(RoundTime(Format(NauticalTwilightEnd, "hh:mm:ss"), Seconds), "hh:mm")
         .Item(1).Text = strTwilights
         strTwilights = "Begin:" & vbTab & Format(RoundTime(Format(AstronomicalTwilightBegin, "hh:mm:ss"), Seconds), "hh:mm") & vbCrLf
         strTwilights = strTwilights & "End:" & vbTab & Format(RoundTime(Format(AstronomicalTwilightEnd, "hh:mm:ss"), Seconds), "hh:mm")
         .Item(2).Text = strTwilights
      End With
      
      Call tmrPosition_Timer
      Call FillThisMonth(Now, True)
      Call FillThisYear(Now, True)
   End If

End Sub

Private Sub SunMoonInfo_ResultSunPosition(IsDate As Date, SunAzimuth As Double, SunZenith As Double, SunElevation As Double)

   With lblSunDate
      .Item(7).Caption = Format(IsDate, DateFormat) & " / " & Format(IsDate, "hh:mm:ss")
      .Item(8).Caption = SunAzimuth
      .Item(9).Caption = SunZenith
      .Item(10).Caption = SunElevation
   End With

End Sub

Private Sub tmrPosition_Timer()

   Call SunMoonInfo.SunPosition(Now)
   Call SunMoonInfo.MoonPosition(Now)

End Sub

Private Sub txtCityInfo_LostFocus(Index As Integer)

   Set FocusObject = txtCityInfo.Item(Index)

End Sub

Private Sub txtMoonPhases_GotFocus()

   FocusObject.SetFocus

End Sub

Private Sub txtTwilights_GotFocus(Index As Integer)

   FocusObject.SetFocus

End Sub
