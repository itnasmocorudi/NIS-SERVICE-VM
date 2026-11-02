VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtAkun 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Setting Kode Jurnal"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14130
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   14130
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7785
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13732
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   706
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483626
      TabCaption(0)   =   "-"
      TabPicture(0)   =   "frmtAkun.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "."
      TabPicture(1)   =   "frmtAkun.frx":0754
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5(53)"
      Tab(1).Control(1)=   "Label5(52)"
      Tab(1).Control(2)=   "Label5(51)"
      Tab(1).Control(3)=   "Label9(25)"
      Tab(1).Control(4)=   "Label6(19)"
      Tab(1).Control(5)=   "Label12(25)"
      Tab(1).Control(6)=   "Label3(31)"
      Tab(1).Control(7)=   "Label5(50)"
      Tab(1).Control(8)=   "Shape1(83)"
      Tab(1).Control(9)=   "Shape1(82)"
      Tab(1).Control(10)=   "Shape1(81)"
      Tab(1).Control(11)=   "Shape1(46)"
      Tab(1).Control(12)=   "Shape1(45)"
      Tab(1).Control(13)=   "Shape1(44)"
      Tab(1).Control(14)=   "Label5(24)"
      Tab(1).Control(15)=   "Label3(16)"
      Tab(1).Control(16)=   "Label12(16)"
      Tab(1).Control(17)=   "Label6(8)"
      Tab(1).Control(18)=   "Label9(15)"
      Tab(1).Control(19)=   "Label5(23)"
      Tab(1).Control(20)=   "txt(63)"
      Tab(1).Control(21)=   "txtx(63)"
      Tab(1).Control(22)=   "txt(64)"
      Tab(1).Control(23)=   "txtx(64)"
      Tab(1).Control(24)=   "txt(65)"
      Tab(1).Control(25)=   "txtx(65)"
      Tab(1).Control(26)=   "txtx(62)"
      Tab(1).Control(27)=   "txt(62)"
      Tab(1).Control(28)=   "txtx(61)"
      Tab(1).Control(29)=   "txt(61)"
      Tab(1).Control(30)=   "txtx(60)"
      Tab(1).Control(31)=   "txt(60)"
      Tab(1).ControlCount=   32
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   60
         Left            =   -71760
         MaxLength       =   15
         TabIndex        =   57
         Top             =   1320
         Width           =   1130
      End
      Begin VB.TextBox txtx 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         IMEMode         =   3  'DISABLE
         Index           =   60
         Left            =   -70515
         MaxLength       =   40
         TabIndex        =   56
         Top             =   1320
         Width           =   6000
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   61
         Left            =   -71760
         MaxLength       =   15
         TabIndex        =   55
         Top             =   1680
         Width           =   1130
      End
      Begin VB.TextBox txtx 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         IMEMode         =   3  'DISABLE
         Index           =   61
         Left            =   -70515
         MaxLength       =   40
         TabIndex        =   54
         Top             =   1680
         Width           =   6000
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   62
         Left            =   -71760
         MaxLength       =   15
         TabIndex        =   53
         Top             =   2040
         Width           =   1130
      End
      Begin VB.TextBox txtx 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         IMEMode         =   3  'DISABLE
         Index           =   62
         Left            =   -70515
         MaxLength       =   40
         TabIndex        =   52
         Top             =   2040
         Width           =   6000
      End
      Begin VB.TextBox txtx 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         IMEMode         =   3  'DISABLE
         Index           =   65
         Left            =   -70515
         MaxLength       =   40
         TabIndex        =   6
         Top             =   3720
         Width           =   6000
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   65
         Left            =   -71760
         MaxLength       =   15
         TabIndex        =   5
         Top             =   3720
         Width           =   1130
      End
      Begin VB.TextBox txtx 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         IMEMode         =   3  'DISABLE
         Index           =   64
         Left            =   -70515
         MaxLength       =   40
         TabIndex        =   4
         Top             =   3360
         Width           =   6000
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   64
         Left            =   -71760
         MaxLength       =   15
         TabIndex        =   3
         Top             =   3360
         Width           =   1130
      End
      Begin VB.TextBox txtx 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         IMEMode         =   3  'DISABLE
         Index           =   63
         Left            =   -70515
         MaxLength       =   40
         TabIndex        =   2
         Top             =   3000
         Width           =   6000
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   63
         Left            =   -71760
         MaxLength       =   15
         TabIndex        =   1
         Top             =   3000
         Width           =   1130
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   6855
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   12091
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "DYNA"
         TabPicture(0)   =   "frmtAkun.frx":0CFC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Line4(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSTab4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "NON DYNA"
         TabPicture(1)   =   "frmtAkun.frx":0D18
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Line4(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "SSTab5"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin TabDlg.SSTab SSTab4 
            Height          =   6255
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   11033
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "EXTERNAL (Pihak Ketiga)"
            TabPicture(0)   =   "frmtAkun.frx":0D34
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Line4(2)"
            Tab(0).Control(1)=   "Frame3(0)"
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "INTERNAL (Pemakaian Sendiri)"
            TabPicture(1)   =   "frmtAkun.frx":0D50
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Line4(4)"
            Tab(1).Control(1)=   "Line4(3)"
            Tab(1).Control(2)=   "Label5(1)"
            Tab(1).Control(3)=   "Label5(5)"
            Tab(1).Control(4)=   "Label7(1)"
            Tab(1).Control(5)=   "Label9(1)"
            Tab(1).Control(6)=   "Label5(6)"
            Tab(1).Control(7)=   "Label5(7)"
            Tab(1).Control(8)=   "Label9(2)"
            Tab(1).Control(9)=   "Label5(8)"
            Tab(1).Control(10)=   "Label5(9)"
            Tab(1).Control(11)=   "Label7(2)"
            Tab(1).Control(12)=   "Label3(0)"
            Tab(1).Control(13)=   "Label6(1)"
            Tab(1).Control(14)=   "Shape1(1)"
            Tab(1).Control(15)=   "Shape1(6)"
            Tab(1).Control(16)=   "Shape1(10)"
            Tab(1).Control(17)=   "Shape1(13)"
            Tab(1).Control(18)=   "Label12(4)"
            Tab(1).Control(19)=   "Label3(3)"
            Tab(1).Control(20)=   "Shape1(14)"
            Tab(1).Control(21)=   "Shape1(15)"
            Tab(1).Control(22)=   "Shape1(16)"
            Tab(1).Control(23)=   "Shape1(17)"
            Tab(1).Control(24)=   "Shape1(18)"
            Tab(1).Control(25)=   "Label12(5)"
            Tab(1).Control(26)=   "Label8(0)"
            Tab(1).Control(27)=   "Label3(5)"
            Tab(1).Control(28)=   "Label6(3)"
            Tab(1).Control(29)=   "Label12(6)"
            Tab(1).Control(30)=   "Label12(7)"
            Tab(1).Control(31)=   "Shape1(19)"
            Tab(1).Control(32)=   "txt(10)"
            Tab(1).Control(33)=   "txt(11)"
            Tab(1).Control(34)=   "txt(12)"
            Tab(1).Control(35)=   "txt(13)"
            Tab(1).Control(36)=   "txt(14)"
            Tab(1).Control(37)=   "txt(15)"
            Tab(1).Control(38)=   "txt(16)"
            Tab(1).Control(39)=   "txt(17)"
            Tab(1).Control(40)=   "txt(18)"
            Tab(1).Control(41)=   "txt(19)"
            Tab(1).Control(42)=   "txtx(10)"
            Tab(1).Control(43)=   "txtx(11)"
            Tab(1).Control(44)=   "txtx(12)"
            Tab(1).Control(45)=   "txtx(13)"
            Tab(1).Control(46)=   "txtx(14)"
            Tab(1).Control(47)=   "txtx(15)"
            Tab(1).Control(48)=   "txtx(16)"
            Tab(1).Control(49)=   "txtx(17)"
            Tab(1).Control(50)=   "txtx(18)"
            Tab(1).Control(51)=   "txtx(19)"
            Tab(1).ControlCount=   52
            TabCaption(2)   =   "AFILIASI (PPn)"
            TabPicture(2)   =   "frmtAkun.frx":0D6C
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Line4(5)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "Label5(10)"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "Label5(11)"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "Label7(3)"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "Label9(3)"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "Label5(12)"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "Label5(13)"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).Control(7)=   "Label9(4)"
            Tab(2).Control(7).Enabled=   0   'False
            Tab(2).Control(8)=   "Label5(14)"
            Tab(2).Control(8).Enabled=   0   'False
            Tab(2).Control(9)=   "Label5(15)"
            Tab(2).Control(9).Enabled=   0   'False
            Tab(2).Control(10)=   "Label7(5)"
            Tab(2).Control(10).Enabled=   0   'False
            Tab(2).Control(11)=   "Label3(6)"
            Tab(2).Control(11).Enabled=   0   'False
            Tab(2).Control(12)=   "Label6(4)"
            Tab(2).Control(12).Enabled=   0   'False
            Tab(2).Control(13)=   "Shape1(20)"
            Tab(2).Control(13).Enabled=   0   'False
            Tab(2).Control(14)=   "Shape1(21)"
            Tab(2).Control(14).Enabled=   0   'False
            Tab(2).Control(15)=   "Shape1(22)"
            Tab(2).Control(15).Enabled=   0   'False
            Tab(2).Control(16)=   "Shape1(23)"
            Tab(2).Control(16).Enabled=   0   'False
            Tab(2).Control(17)=   "Label12(8)"
            Tab(2).Control(17).Enabled=   0   'False
            Tab(2).Control(18)=   "Label3(7)"
            Tab(2).Control(18).Enabled=   0   'False
            Tab(2).Control(19)=   "Shape1(24)"
            Tab(2).Control(19).Enabled=   0   'False
            Tab(2).Control(20)=   "Shape1(25)"
            Tab(2).Control(20).Enabled=   0   'False
            Tab(2).Control(21)=   "Shape1(26)"
            Tab(2).Control(21).Enabled=   0   'False
            Tab(2).Control(22)=   "Shape1(27)"
            Tab(2).Control(22).Enabled=   0   'False
            Tab(2).Control(23)=   "Shape1(28)"
            Tab(2).Control(23).Enabled=   0   'False
            Tab(2).Control(24)=   "Label12(9)"
            Tab(2).Control(24).Enabled=   0   'False
            Tab(2).Control(25)=   "Label8(2)"
            Tab(2).Control(25).Enabled=   0   'False
            Tab(2).Control(26)=   "Label3(8)"
            Tab(2).Control(26).Enabled=   0   'False
            Tab(2).Control(27)=   "Label6(5)"
            Tab(2).Control(27).Enabled=   0   'False
            Tab(2).Control(28)=   "Label12(10)"
            Tab(2).Control(28).Enabled=   0   'False
            Tab(2).Control(29)=   "Label12(11)"
            Tab(2).Control(29).Enabled=   0   'False
            Tab(2).Control(30)=   "Shape1(29)"
            Tab(2).Control(30).Enabled=   0   'False
            Tab(2).Control(31)=   "txt(20)"
            Tab(2).Control(31).Enabled=   0   'False
            Tab(2).Control(32)=   "txt(21)"
            Tab(2).Control(32).Enabled=   0   'False
            Tab(2).Control(33)=   "txt(22)"
            Tab(2).Control(33).Enabled=   0   'False
            Tab(2).Control(34)=   "txt(23)"
            Tab(2).Control(34).Enabled=   0   'False
            Tab(2).Control(35)=   "txt(24)"
            Tab(2).Control(35).Enabled=   0   'False
            Tab(2).Control(36)=   "txt(25)"
            Tab(2).Control(36).Enabled=   0   'False
            Tab(2).Control(37)=   "txt(26)"
            Tab(2).Control(37).Enabled=   0   'False
            Tab(2).Control(38)=   "txt(27)"
            Tab(2).Control(38).Enabled=   0   'False
            Tab(2).Control(39)=   "txt(28)"
            Tab(2).Control(39).Enabled=   0   'False
            Tab(2).Control(40)=   "txt(29)"
            Tab(2).Control(40).Enabled=   0   'False
            Tab(2).Control(41)=   "txtx(20)"
            Tab(2).Control(41).Enabled=   0   'False
            Tab(2).Control(42)=   "txtx(21)"
            Tab(2).Control(42).Enabled=   0   'False
            Tab(2).Control(43)=   "txtx(22)"
            Tab(2).Control(43).Enabled=   0   'False
            Tab(2).Control(44)=   "txtx(23)"
            Tab(2).Control(44).Enabled=   0   'False
            Tab(2).Control(45)=   "txtx(24)"
            Tab(2).Control(45).Enabled=   0   'False
            Tab(2).Control(46)=   "txtx(25)"
            Tab(2).Control(46).Enabled=   0   'False
            Tab(2).Control(47)=   "txtx(26)"
            Tab(2).Control(47).Enabled=   0   'False
            Tab(2).Control(48)=   "txtx(27)"
            Tab(2).Control(48).Enabled=   0   'False
            Tab(2).Control(49)=   "txtx(28)"
            Tab(2).Control(49).Enabled=   0   'False
            Tab(2).Control(50)=   "txtx(29)"
            Tab(2).Control(50).Enabled=   0   'False
            Tab(2).ControlCount=   51
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   29
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   131
               Top             =   4080
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   28
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   130
               Top             =   3750
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   27
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   129
               Top             =   3240
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   26
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   128
               Top             =   2910
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   25
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   127
               Top             =   2580
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   24
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   126
               Top             =   2250
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   23
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   125
               Top             =   1920
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   22
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   124
               Top             =   1590
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   21
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   123
               Top             =   1260
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   20
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   122
               Top             =   840
               Width           =   6000
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   29
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   121
               Top             =   4080
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   28
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   120
               Top             =   3750
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   27
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   119
               Top             =   3240
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   26
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   118
               Top             =   2910
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   25
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   117
               Top             =   2580
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   24
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   116
               Top             =   2250
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   23
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   115
               Top             =   1920
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   22
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   114
               Top             =   1590
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   21
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   113
               Top             =   1260
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   20
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   112
               Top             =   840
               Width           =   1130
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   19
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   91
               Top             =   4080
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   18
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   90
               Top             =   3750
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   17
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   89
               Top             =   3240
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   16
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   88
               Top             =   2910
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   15
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   87
               Top             =   2580
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   14
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   86
               Top             =   2250
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   13
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   85
               Top             =   1920
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   12
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   84
               Top             =   1590
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   11
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   83
               Top             =   1260
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   10
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   82
               Top             =   840
               Width           =   6000
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   19
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   81
               Top             =   4080
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   18
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   80
               Top             =   3750
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   17
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   79
               Top             =   3240
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   16
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   78
               Top             =   2910
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   15
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   77
               Top             =   2580
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   14
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   76
               Top             =   2250
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   13
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   75
               Top             =   1920
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   12
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   74
               Top             =   1590
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   11
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   73
               Top             =   1260
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   10
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   72
               Top             =   840
               Width           =   1130
            End
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5175
               Index           =   0
               Left            =   -74400
               TabIndex        =   9
               Top             =   840
               Width           =   9975
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   0
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   29
                  Top             =   0
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   1
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   28
                  Top             =   420
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   2
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   27
                  Top             =   750
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   3
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   26
                  Top             =   1080
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   4
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   25
                  Top             =   1410
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   5
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   24
                  Top             =   1740
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   6
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   23
                  Top             =   2070
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   7
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   22
                  Top             =   2400
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   8
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   21
                  Top             =   2910
                  Width           =   1130
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   9
                  Left            =   2490
                  MaxLength       =   15
                  TabIndex        =   20
                  Top             =   3240
                  Width           =   1130
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   0
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   19
                  Top             =   0
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   1
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   18
                  Top             =   420
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   2
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   17
                  Top             =   750
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   3
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   16
                  Top             =   1080
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   4
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   15
                  Top             =   1410
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   5
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   14
                  Top             =   1740
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   6
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   13
                  Top             =   2070
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   7
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   12
                  Top             =   2400
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   8
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   11
                  Top             =   2910
                  Width           =   6000
               End
               Begin VB.TextBox txtx 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   310
                  IMEMode         =   3  'DISABLE
                  Index           =   9
                  Left            =   3735
                  MaxLength       =   40
                  TabIndex        =   10
                  Top             =   3240
                  Width           =   6000
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Piutang Service"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   3
                  Left            =   0
                  TabIndex        =   50
                  Top             =   0
                  Width           =   1530
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Penj. Olie"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   2
                  Left            =   240
                  TabIndex        =   49
                  Top             =   1440
                  Width           =   930
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Penj. Accesories"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   48
                  Top             =   2010
                  Width           =   1620
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Penj. Sublet"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   47
                  Top             =   750
                  Width           =   1170
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Penj. Jasa"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   240
                  TabIndex        =   46
                  Top             =   420
                  Width           =   990
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Penj. Part"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   4
                  Left            =   240
                  TabIndex        =   45
                  Top             =   1080
                  Width           =   960
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "BP In Proses"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   17
                  Left            =   240
                  TabIndex        =   44
                  Top             =   3255
                  Width           =   1230
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "HPP Jasa Sublet"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   26
                  Left            =   0
                  TabIndex        =   43
                  Top             =   2940
                  Width           =   1545
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "PPN"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   27
                  Left            =   240
                  TabIndex        =   42
                  Top             =   1720
                  Width           =   375
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Materai"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   4
                  Left            =   240
                  TabIndex        =   41
                  Top             =   2415
                  Width           =   735
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   4
                  Left            =   2400
                  TabIndex        =   40
                  Top             =   2910
                  Width           =   60
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   39
                  Top             =   3225
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   11
                  Left            =   3630
                  Top             =   2910
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   12
                  Left            =   3630
                  Top             =   3240
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   9
                  Left            =   3630
                  Top             =   2070
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   7
                  Left            =   3630
                  Top             =   1740
                  Width           =   60
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   38
                  Top             =   1740
                  Width           =   60
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   37
                  Top             =   2070
                  Width           =   45
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   5
                  Left            =   3630
                  Top             =   1410
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   4
                  Left            =   3630
                  Top             =   1080
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   3
                  Left            =   3630
                  Top             =   750
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   2
                  Left            =   3630
                  Top             =   420
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   8
                  Left            =   3630
                  Top             =   0
                  Width           =   60
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   36
                  Top             =   750
                  Width           =   60
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   35
                  Top             =   1080
                  Width           =   60
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   34
                  Top             =   1410
                  Width           =   60
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   2400
                  TabIndex        =   33
                  Top             =   420
                  Width           =   60
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   0
                  Left            =   2400
                  TabIndex        =   32
                  Top             =   0
                  Width           =   60
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   3
                  Left            =   2400
                  TabIndex        =   31
                  Top             =   2400
                  Width           =   60
               End
               Begin VB.Shape Shape1 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H000080FF&
                  Height          =   315
                  Index           =   0
                  Left            =   3630
                  Top             =   2400
                  Width           =   60
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "DYNA - EXTERNAL (Pihak Ketiga)"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Index           =   44
                  Left            =   5640
                  TabIndex        =   30
                  Top             =   4560
                  Width           =   4140
               End
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   29
               Left            =   4230
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   3000
               TabIndex        =   151
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   3000
               TabIndex        =   150
               Top             =   840
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   3000
               TabIndex        =   149
               Top             =   1260
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   3000
               TabIndex        =   148
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   3000
               TabIndex        =   147
               Top             =   1920
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   3000
               TabIndex        =   146
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   28
               Left            =   4230
               Top             =   840
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   27
               Left            =   4230
               Top             =   1260
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   26
               Left            =   4230
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   25
               Left            =   4230
               Top             =   1920
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   24
               Left            =   4230
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   3000
               TabIndex        =   145
               Top             =   2910
               Width           =   45
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   3000
               TabIndex        =   144
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   23
               Left            =   4230
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   22
               Left            =   4230
               Top             =   2910
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   21
               Left            =   4230
               Top             =   4080
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   20
               Left            =   4230
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   3000
               TabIndex        =   143
               Top             =   4065
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   3000
               TabIndex        =   142
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Materai"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   840
               TabIndex        =   141
               Top             =   3255
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PPN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   840
               TabIndex        =   140
               Top             =   2565
               Width           =   375
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HPP Jasa Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   600
               TabIndex        =   139
               Top             =   3780
               Width           =   1545
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BP In Proses"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   840
               TabIndex        =   138
               Top             =   4095
               Width           =   1230
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Part"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   840
               TabIndex        =   137
               Top             =   1920
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Jasa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   840
               TabIndex        =   136
               Top             =   1260
               Width           =   990
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   840
               TabIndex        =   135
               Top             =   1590
               Width           =   1170
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Accesories"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   840
               TabIndex        =   134
               Top             =   2850
               Width           =   1620
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Olie"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   840
               TabIndex        =   133
               Top             =   2280
               Width           =   930
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piutang Service"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   600
               TabIndex        =   132
               Top             =   840
               Width           =   1530
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   19
               Left            =   -70770
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   -72000
               TabIndex        =   111
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   -72000
               TabIndex        =   110
               Top             =   840
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   -72000
               TabIndex        =   109
               Top             =   1260
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   -72000
               TabIndex        =   108
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   -72000
               TabIndex        =   107
               Top             =   1920
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   -72000
               TabIndex        =   106
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   18
               Left            =   -70770
               Top             =   840
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   17
               Left            =   -70770
               Top             =   1260
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   16
               Left            =   -70770
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   15
               Left            =   -70770
               Top             =   1920
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   14
               Left            =   -70770
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   -72000
               TabIndex        =   105
               Top             =   2910
               Width           =   45
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   -72000
               TabIndex        =   104
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   13
               Left            =   -70770
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   10
               Left            =   -70770
               Top             =   2910
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   6
               Left            =   -70770
               Top             =   4080
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   1
               Left            =   -70770
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   -72000
               TabIndex        =   103
               Top             =   4065
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   -72000
               TabIndex        =   102
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Materai"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   -74160
               TabIndex        =   101
               Top             =   3255
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PPN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   -74160
               TabIndex        =   100
               Top             =   2565
               Width           =   375
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HPP Jasa Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   -74400
               TabIndex        =   99
               Top             =   3780
               Width           =   1545
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BP In Proses"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   -74160
               TabIndex        =   98
               Top             =   4095
               Width           =   1230
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Part"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   -74160
               TabIndex        =   97
               Top             =   1920
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Jasa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   -74160
               TabIndex        =   96
               Top             =   1260
               Width           =   990
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   -74160
               TabIndex        =   95
               Top             =   1590
               Width           =   1170
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Accesories"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   -74160
               TabIndex        =   94
               Top             =   2850
               Width           =   1620
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Olie"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   -74160
               TabIndex        =   93
               Top             =   2280
               Width           =   930
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piutang Service"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   -74400
               TabIndex        =   92
               Top             =   840
               Width           =   1530
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   2
               X1              =   -71400
               X2              =   -63960
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   3
               X1              =   -75000
               X2              =   -71280
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   4
               X1              =   -67680
               X2              =   -63240
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   5
               X1              =   0
               X2              =   7320
               Y1              =   360
               Y2              =   360
            End
         End
         Begin TabDlg.SSTab SSTab5 
            Height          =   6015
            Left            =   -74880
            TabIndex        =   51
            Top             =   480
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   10610
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "EXTERNAL (Pihak Ketiga)"
            TabPicture(0)   =   "frmtAkun.frx":0D88
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Line4(6)"
            Tab(0).Control(1)=   "Label5(16)"
            Tab(0).Control(2)=   "Label5(17)"
            Tab(0).Control(3)=   "Label7(6)"
            Tab(0).Control(4)=   "Label9(5)"
            Tab(0).Control(5)=   "Label5(18)"
            Tab(0).Control(6)=   "Label5(19)"
            Tab(0).Control(7)=   "Label9(6)"
            Tab(0).Control(8)=   "Label5(20)"
            Tab(0).Control(9)=   "Label5(21)"
            Tab(0).Control(10)=   "Label7(7)"
            Tab(0).Control(11)=   "Label3(9)"
            Tab(0).Control(12)=   "Label6(6)"
            Tab(0).Control(13)=   "Shape1(30)"
            Tab(0).Control(14)=   "Shape1(31)"
            Tab(0).Control(15)=   "Shape1(32)"
            Tab(0).Control(16)=   "Shape1(33)"
            Tab(0).Control(17)=   "Label12(12)"
            Tab(0).Control(18)=   "Label3(10)"
            Tab(0).Control(19)=   "Shape1(34)"
            Tab(0).Control(20)=   "Shape1(35)"
            Tab(0).Control(21)=   "Shape1(36)"
            Tab(0).Control(22)=   "Shape1(37)"
            Tab(0).Control(23)=   "Shape1(38)"
            Tab(0).Control(24)=   "Label12(13)"
            Tab(0).Control(25)=   "Label8(3)"
            Tab(0).Control(26)=   "Label3(11)"
            Tab(0).Control(27)=   "Label6(7)"
            Tab(0).Control(28)=   "Label12(14)"
            Tab(0).Control(29)=   "Label12(15)"
            Tab(0).Control(30)=   "Shape1(39)"
            Tab(0).Control(31)=   "txt(30)"
            Tab(0).Control(32)=   "txt(31)"
            Tab(0).Control(33)=   "txt(32)"
            Tab(0).Control(34)=   "txt(33)"
            Tab(0).Control(35)=   "txt(34)"
            Tab(0).Control(36)=   "txt(35)"
            Tab(0).Control(37)=   "txt(36)"
            Tab(0).Control(38)=   "txt(37)"
            Tab(0).Control(39)=   "txt(38)"
            Tab(0).Control(40)=   "txt(39)"
            Tab(0).Control(41)=   "txtx(30)"
            Tab(0).Control(42)=   "txtx(31)"
            Tab(0).Control(43)=   "txtx(32)"
            Tab(0).Control(44)=   "txtx(33)"
            Tab(0).Control(45)=   "txtx(34)"
            Tab(0).Control(46)=   "txtx(35)"
            Tab(0).Control(47)=   "txtx(36)"
            Tab(0).Control(48)=   "txtx(37)"
            Tab(0).Control(49)=   "txtx(38)"
            Tab(0).Control(50)=   "txtx(39)"
            Tab(0).ControlCount=   51
            TabCaption(1)   =   "INTERNAL (Pemakaian Sendiri)"
            TabPicture(1)   =   "frmtAkun.frx":0DA4
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Line4(9)"
            Tab(1).Control(1)=   "Line4(8)"
            Tab(1).Control(2)=   "Label5(22)"
            Tab(1).Control(3)=   "Label5(25)"
            Tab(1).Control(4)=   "Label7(8)"
            Tab(1).Control(5)=   "Label9(7)"
            Tab(1).Control(6)=   "Label5(28)"
            Tab(1).Control(7)=   "Label5(29)"
            Tab(1).Control(8)=   "Label9(8)"
            Tab(1).Control(9)=   "Label5(30)"
            Tab(1).Control(10)=   "Label5(31)"
            Tab(1).Control(11)=   "Label7(9)"
            Tab(1).Control(12)=   "Label3(12)"
            Tab(1).Control(13)=   "Label6(9)"
            Tab(1).Control(14)=   "Shape1(40)"
            Tab(1).Control(15)=   "Shape1(41)"
            Tab(1).Control(16)=   "Shape1(42)"
            Tab(1).Control(17)=   "Shape1(43)"
            Tab(1).Control(18)=   "Label12(17)"
            Tab(1).Control(19)=   "Label3(13)"
            Tab(1).Control(20)=   "Shape1(47)"
            Tab(1).Control(21)=   "Shape1(48)"
            Tab(1).Control(22)=   "Shape1(49)"
            Tab(1).Control(23)=   "Shape1(50)"
            Tab(1).Control(24)=   "Shape1(51)"
            Tab(1).Control(25)=   "Label12(18)"
            Tab(1).Control(26)=   "Label8(4)"
            Tab(1).Control(27)=   "Label3(14)"
            Tab(1).Control(28)=   "Label6(10)"
            Tab(1).Control(29)=   "Label12(19)"
            Tab(1).Control(30)=   "Label12(20)"
            Tab(1).Control(31)=   "Shape1(52)"
            Tab(1).Control(32)=   "txt(40)"
            Tab(1).Control(33)=   "txt(41)"
            Tab(1).Control(34)=   "txt(42)"
            Tab(1).Control(35)=   "txt(43)"
            Tab(1).Control(36)=   "txt(44)"
            Tab(1).Control(37)=   "txt(45)"
            Tab(1).Control(38)=   "txt(46)"
            Tab(1).Control(39)=   "txt(47)"
            Tab(1).Control(40)=   "txt(48)"
            Tab(1).Control(41)=   "txt(49)"
            Tab(1).Control(42)=   "txtx(40)"
            Tab(1).Control(43)=   "txtx(41)"
            Tab(1).Control(44)=   "txtx(42)"
            Tab(1).Control(45)=   "txtx(43)"
            Tab(1).Control(46)=   "txtx(44)"
            Tab(1).Control(47)=   "txtx(45)"
            Tab(1).Control(48)=   "txtx(46)"
            Tab(1).Control(49)=   "txtx(47)"
            Tab(1).Control(50)=   "txtx(48)"
            Tab(1).Control(51)=   "txtx(49)"
            Tab(1).ControlCount=   52
            TabCaption(2)   =   "AFILIASI (PPn)"
            TabPicture(2)   =   "frmtAkun.frx":0DC0
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Line4(7)"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "Label5(32)"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "Label5(33)"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "Label7(10)"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "Label9(9)"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "Label5(34)"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "Label5(35)"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).Control(7)=   "Label9(10)"
            Tab(2).Control(7).Enabled=   0   'False
            Tab(2).Control(8)=   "Label5(36)"
            Tab(2).Control(8).Enabled=   0   'False
            Tab(2).Control(9)=   "Label5(37)"
            Tab(2).Control(9).Enabled=   0   'False
            Tab(2).Control(10)=   "Label7(11)"
            Tab(2).Control(10).Enabled=   0   'False
            Tab(2).Control(11)=   "Label3(15)"
            Tab(2).Control(11).Enabled=   0   'False
            Tab(2).Control(12)=   "Label6(11)"
            Tab(2).Control(12).Enabled=   0   'False
            Tab(2).Control(13)=   "Shape1(53)"
            Tab(2).Control(13).Enabled=   0   'False
            Tab(2).Control(14)=   "Shape1(54)"
            Tab(2).Control(14).Enabled=   0   'False
            Tab(2).Control(15)=   "Shape1(55)"
            Tab(2).Control(15).Enabled=   0   'False
            Tab(2).Control(16)=   "Shape1(56)"
            Tab(2).Control(16).Enabled=   0   'False
            Tab(2).Control(17)=   "Label12(21)"
            Tab(2).Control(17).Enabled=   0   'False
            Tab(2).Control(18)=   "Label3(17)"
            Tab(2).Control(18).Enabled=   0   'False
            Tab(2).Control(19)=   "Shape1(57)"
            Tab(2).Control(19).Enabled=   0   'False
            Tab(2).Control(20)=   "Shape1(58)"
            Tab(2).Control(20).Enabled=   0   'False
            Tab(2).Control(21)=   "Shape1(59)"
            Tab(2).Control(21).Enabled=   0   'False
            Tab(2).Control(22)=   "Shape1(60)"
            Tab(2).Control(22).Enabled=   0   'False
            Tab(2).Control(23)=   "Shape1(61)"
            Tab(2).Control(23).Enabled=   0   'False
            Tab(2).Control(24)=   "Label12(22)"
            Tab(2).Control(24).Enabled=   0   'False
            Tab(2).Control(25)=   "Label8(5)"
            Tab(2).Control(25).Enabled=   0   'False
            Tab(2).Control(26)=   "Label3(18)"
            Tab(2).Control(26).Enabled=   0   'False
            Tab(2).Control(27)=   "Label6(12)"
            Tab(2).Control(27).Enabled=   0   'False
            Tab(2).Control(28)=   "Label12(23)"
            Tab(2).Control(28).Enabled=   0   'False
            Tab(2).Control(29)=   "Label12(24)"
            Tab(2).Control(29).Enabled=   0   'False
            Tab(2).Control(30)=   "Shape1(62)"
            Tab(2).Control(30).Enabled=   0   'False
            Tab(2).Control(31)=   "txt(50)"
            Tab(2).Control(31).Enabled=   0   'False
            Tab(2).Control(32)=   "txt(51)"
            Tab(2).Control(32).Enabled=   0   'False
            Tab(2).Control(33)=   "txt(52)"
            Tab(2).Control(33).Enabled=   0   'False
            Tab(2).Control(34)=   "txt(53)"
            Tab(2).Control(34).Enabled=   0   'False
            Tab(2).Control(35)=   "txt(54)"
            Tab(2).Control(35).Enabled=   0   'False
            Tab(2).Control(36)=   "txt(55)"
            Tab(2).Control(36).Enabled=   0   'False
            Tab(2).Control(37)=   "txt(56)"
            Tab(2).Control(37).Enabled=   0   'False
            Tab(2).Control(38)=   "txt(57)"
            Tab(2).Control(38).Enabled=   0   'False
            Tab(2).Control(39)=   "txt(58)"
            Tab(2).Control(39).Enabled=   0   'False
            Tab(2).Control(40)=   "txt(59)"
            Tab(2).Control(40).Enabled=   0   'False
            Tab(2).Control(41)=   "txtx(50)"
            Tab(2).Control(41).Enabled=   0   'False
            Tab(2).Control(42)=   "txtx(51)"
            Tab(2).Control(42).Enabled=   0   'False
            Tab(2).Control(43)=   "txtx(52)"
            Tab(2).Control(43).Enabled=   0   'False
            Tab(2).Control(44)=   "txtx(53)"
            Tab(2).Control(44).Enabled=   0   'False
            Tab(2).Control(45)=   "txtx(54)"
            Tab(2).Control(45).Enabled=   0   'False
            Tab(2).Control(46)=   "txtx(55)"
            Tab(2).Control(46).Enabled=   0   'False
            Tab(2).Control(47)=   "txtx(56)"
            Tab(2).Control(47).Enabled=   0   'False
            Tab(2).Control(48)=   "txtx(57)"
            Tab(2).Control(48).Enabled=   0   'False
            Tab(2).Control(49)=   "txtx(58)"
            Tab(2).Control(49).Enabled=   0   'False
            Tab(2).Control(50)=   "txtx(59)"
            Tab(2).Control(50).Enabled=   0   'False
            Tab(2).ControlCount=   51
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   59
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   251
               Top             =   4080
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   58
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   250
               Top             =   3750
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   57
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   249
               Top             =   3240
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   56
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   248
               Top             =   2910
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   55
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   247
               Top             =   2580
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   54
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   246
               Top             =   2250
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   53
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   245
               Top             =   1920
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   52
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   244
               Top             =   1590
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   51
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   243
               Top             =   1260
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   50
               Left            =   4335
               MaxLength       =   40
               TabIndex        =   242
               Top             =   840
               Width           =   6000
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   59
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   241
               Top             =   4080
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   58
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   240
               Top             =   3750
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   57
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   239
               Top             =   3240
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   56
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   238
               Top             =   2910
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   55
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   237
               Top             =   2580
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   54
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   236
               Top             =   2250
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   53
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   235
               Top             =   1920
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   52
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   234
               Top             =   1590
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   51
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   233
               Top             =   1260
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   50
               Left            =   3090
               MaxLength       =   15
               TabIndex        =   232
               Top             =   840
               Width           =   1130
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   49
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   211
               Top             =   4080
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   48
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   210
               Top             =   3750
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   47
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   209
               Top             =   3240
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   46
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   208
               Top             =   2910
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   45
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   207
               Top             =   2580
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   44
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   206
               Top             =   2250
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   43
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   205
               Top             =   1920
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   42
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   204
               Top             =   1590
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   41
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   203
               Top             =   1260
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   40
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   202
               Top             =   840
               Width           =   6000
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   49
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   201
               Top             =   4080
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   48
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   200
               Top             =   3750
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   47
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   199
               Top             =   3240
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   46
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   198
               Top             =   2910
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   45
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   197
               Top             =   2580
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   44
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   196
               Top             =   2250
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   43
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   195
               Top             =   1920
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   42
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   194
               Top             =   1590
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   41
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   193
               Top             =   1260
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   40
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   192
               Top             =   840
               Width           =   1130
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   39
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   171
               Top             =   4080
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   38
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   170
               Top             =   3750
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   37
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   169
               Top             =   3240
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   36
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   168
               Top             =   2910
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   35
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   167
               Top             =   2580
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   34
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   166
               Top             =   2250
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   33
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   165
               Top             =   1920
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   32
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   164
               Top             =   1590
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   31
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   163
               Top             =   1260
               Width           =   6000
            End
            Begin VB.TextBox txtx 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   310
               IMEMode         =   3  'DISABLE
               Index           =   30
               Left            =   -70665
               MaxLength       =   40
               TabIndex        =   162
               Top             =   840
               Width           =   6000
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   39
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   161
               Top             =   4080
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   38
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   160
               Top             =   3750
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   37
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   159
               Top             =   3240
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   36
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   158
               Top             =   2910
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   35
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   157
               Top             =   2580
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   34
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   156
               Top             =   2250
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   33
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   155
               Top             =   1920
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   32
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   154
               Top             =   1590
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   31
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   153
               Top             =   1260
               Width           =   1130
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   30
               Left            =   -71910
               MaxLength       =   15
               TabIndex        =   152
               Top             =   840
               Width           =   1130
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   62
               Left            =   4230
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   24
               Left            =   3000
               TabIndex        =   271
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   23
               Left            =   3000
               TabIndex        =   270
               Top             =   840
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   3000
               TabIndex        =   269
               Top             =   1260
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   3000
               TabIndex        =   268
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   3000
               TabIndex        =   267
               Top             =   1920
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   3000
               TabIndex        =   266
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   61
               Left            =   4230
               Top             =   840
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   60
               Left            =   4230
               Top             =   1260
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   59
               Left            =   4230
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   58
               Left            =   4230
               Top             =   1920
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   57
               Left            =   4230
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   3000
               TabIndex        =   265
               Top             =   2910
               Width           =   45
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   3000
               TabIndex        =   264
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   56
               Left            =   4230
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   55
               Left            =   4230
               Top             =   2910
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   54
               Left            =   4230
               Top             =   4080
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   53
               Left            =   4230
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   3000
               TabIndex        =   263
               Top             =   4065
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   3000
               TabIndex        =   262
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Materai"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   840
               TabIndex        =   261
               Top             =   3255
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PPN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   37
               Left            =   840
               TabIndex        =   260
               Top             =   2565
               Width           =   375
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HPP Jasa Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   36
               Left            =   600
               TabIndex        =   259
               Top             =   3780
               Width           =   1545
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BP In Proses"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   840
               TabIndex        =   258
               Top             =   4095
               Width           =   1230
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Part"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   35
               Left            =   840
               TabIndex        =   257
               Top             =   1920
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Jasa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   34
               Left            =   840
               TabIndex        =   256
               Top             =   1260
               Width           =   990
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   840
               TabIndex        =   255
               Top             =   1590
               Width           =   1170
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Accesories"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   840
               TabIndex        =   254
               Top             =   2850
               Width           =   1620
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Olie"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   33
               Left            =   840
               TabIndex        =   253
               Top             =   2280
               Width           =   930
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piutang Service"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   32
               Left            =   600
               TabIndex        =   252
               Top             =   840
               Width           =   1530
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   52
               Left            =   -70770
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   -72000
               TabIndex        =   231
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   -72000
               TabIndex        =   230
               Top             =   840
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   -72000
               TabIndex        =   229
               Top             =   1260
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   -72000
               TabIndex        =   228
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   -72000
               TabIndex        =   227
               Top             =   1920
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   -72000
               TabIndex        =   226
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   51
               Left            =   -70770
               Top             =   840
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   50
               Left            =   -70770
               Top             =   1260
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   49
               Left            =   -70770
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   48
               Left            =   -70770
               Top             =   1920
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   47
               Left            =   -70770
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   -72000
               TabIndex        =   225
               Top             =   2910
               Width           =   45
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   -72000
               TabIndex        =   224
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   43
               Left            =   -70770
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   42
               Left            =   -70770
               Top             =   2910
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   41
               Left            =   -70770
               Top             =   4080
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   40
               Left            =   -70770
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   -72000
               TabIndex        =   223
               Top             =   4065
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   -72000
               TabIndex        =   222
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Materai"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   -74160
               TabIndex        =   221
               Top             =   3255
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PPN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   31
               Left            =   -74160
               TabIndex        =   220
               Top             =   2565
               Width           =   375
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HPP Jasa Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   30
               Left            =   -74400
               TabIndex        =   219
               Top             =   3780
               Width           =   1545
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BP In Proses"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   -74160
               TabIndex        =   218
               Top             =   4095
               Width           =   1230
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Part"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   29
               Left            =   -74160
               TabIndex        =   217
               Top             =   1920
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Jasa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   28
               Left            =   -74160
               TabIndex        =   216
               Top             =   1260
               Width           =   990
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   -74160
               TabIndex        =   215
               Top             =   1590
               Width           =   1170
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Accesories"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   -74160
               TabIndex        =   214
               Top             =   2850
               Width           =   1620
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Olie"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   25
               Left            =   -74160
               TabIndex        =   213
               Top             =   2280
               Width           =   930
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piutang Service"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   -74400
               TabIndex        =   212
               Top             =   840
               Width           =   1530
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   39
               Left            =   -70770
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   -72000
               TabIndex        =   191
               Top             =   3240
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   -72000
               TabIndex        =   190
               Top             =   840
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   -72000
               TabIndex        =   189
               Top             =   1260
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   -72000
               TabIndex        =   188
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   -72000
               TabIndex        =   187
               Top             =   1920
               Width           =   60
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   -72000
               TabIndex        =   186
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   38
               Left            =   -70770
               Top             =   840
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   37
               Left            =   -70770
               Top             =   1260
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   36
               Left            =   -70770
               Top             =   1590
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   35
               Left            =   -70770
               Top             =   1920
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   34
               Left            =   -70770
               Top             =   2250
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   -72000
               TabIndex        =   185
               Top             =   2910
               Width           =   45
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   -72000
               TabIndex        =   184
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   33
               Left            =   -70770
               Top             =   2580
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   32
               Left            =   -70770
               Top             =   2910
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   31
               Left            =   -70770
               Top             =   4080
               Width           =   60
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H0080C0FF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H000080FF&
               Height          =   315
               Index           =   30
               Left            =   -70770
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   -72000
               TabIndex        =   183
               Top             =   4065
               Width           =   60
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   -72000
               TabIndex        =   182
               Top             =   3750
               Width           =   60
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Materai"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   -74160
               TabIndex        =   181
               Top             =   3255
               Width           =   735
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PPN"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   -74160
               TabIndex        =   180
               Top             =   2565
               Width           =   375
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HPP Jasa Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   -74400
               TabIndex        =   179
               Top             =   3780
               Width           =   1545
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BP In Proses"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   -74160
               TabIndex        =   178
               Top             =   4095
               Width           =   1230
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Part"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   -74160
               TabIndex        =   177
               Top             =   1920
               Width           =   960
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Jasa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   -74160
               TabIndex        =   176
               Top             =   1260
               Width           =   990
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Sublet"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   -74160
               TabIndex        =   175
               Top             =   1590
               Width           =   1170
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Accesories"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   -74160
               TabIndex        =   174
               Top             =   2850
               Width           =   1620
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penj. Olie"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   -74160
               TabIndex        =   173
               Top             =   2280
               Width           =   930
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Piutang Service"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   -74400
               TabIndex        =   172
               Top             =   840
               Width           =   1530
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   6
               X1              =   -71400
               X2              =   -63960
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   7
               X1              =   0
               X2              =   7320
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   8
               X1              =   -78720
               X2              =   -71280
               Y1              =   360
               Y2              =   360
            End
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   9
               X1              =   -67680
               X2              =   -60240
               Y1              =   360
               Y2              =   360
            End
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000080FF&
            BorderWidth     =   2
            Index           =   0
            X1              =   5640
            X2              =   11280
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000080FF&
            BorderWidth     =   2
            Index           =   1
            X1              =   -75000
            X2              =   -69360
            Y1              =   360
            Y2              =   360
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         X1              =   2880
         X2              =   11520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemakaian Sublet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   -74280
         TabIndex        =   71
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPn Sublet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   -74280
         TabIndex        =   70
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   -71925
         TabIndex        =   69
         Top             =   2040
         Width           =   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   -71925
         TabIndex        =   68
         Top             =   1320
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   -71925
         TabIndex        =   67
         Top             =   1680
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hutang Sublet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   -74040
         TabIndex        =   66
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         Height          =   315
         Index           =   44
         Left            =   -70620
         Top             =   1320
         Width           =   60
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         Height          =   315
         Index           =   45
         Left            =   -70620
         Top             =   1680
         Width           =   60
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         Height          =   315
         Index           =   46
         Left            =   -70620
         Top             =   2040
         Width           =   60
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         Height          =   315
         Index           =   81
         Left            =   -70620
         Top             =   3720
         Width           =   60
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         Height          =   315
         Index           =   82
         Left            =   -70620
         Top             =   3360
         Width           =   60
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000080FF&
         Height          =   315
         Index           =   83
         Left            =   -70620
         Top             =   3000
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hutang Sublet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   50
         Left            =   -74040
         TabIndex        =   65
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   -71925
         TabIndex        =   64
         Top             =   3360
         Width           =   60
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   -71925
         TabIndex        =   63
         Top             =   3000
         Width           =   60
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   -71925
         TabIndex        =   62
         Top             =   3720
         Width           =   60
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPn Sublet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   -74280
         TabIndex        =   61
         Top             =   3360
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemakaian Sublet"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   51
         Left            =   -74280
         TabIndex        =   60
         Top             =   3000
         Width           =   1755
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D Y N A ---------------------------------------------------------------------------------------------------"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   52
         Left            =   -74280
         TabIndex        =   59
         Top             =   960
         Width           =   9675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N O N   D Y N A -------------------------------------------------------------------------------------------"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   53
         Left            =   -74280
         TabIndex        =   58
         Top             =   2640
         Width           =   9660
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   0
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin ucXPButton.XPButton cmdSave 
      Height          =   495
      Left            =   9480
      TabIndex        =   272
      Top             =   8280
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&s"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmtAkun.frx":0DDC
      PICN            =   "frmtAkun.frx":0DF8
      PICH            =   "frmtAkun.frx":130B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdExit 
      Height          =   495
      Left            =   10680
      TabIndex        =   273
      Top             =   8280
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&x"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmtAkun.frx":18AD
      PICN            =   "frmtAkun.frx":18C9
      PICH            =   "frmtAkun.frx":1D9D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9735
      Left            =   0
      Picture         =   "frmtAkun.frx":22E8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmtAkun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
'''    Dim aa As String
'''    Dim xHasil As String
'''    xHasil = ""
'''    If xUpdate Then
'''        xHasil = "UPDATE " + xTabel + " SET "
'''        i = 0
'''        Do While gsSql1(i) <> ""
'''            If i <> 0 Then
'''                xHasil = xHasil + ","
'''            End If
'''
'''            aa = UCase(gsSql1(i))
'''            Select Case aa
'''            Case "PPN", "DISC"
'''                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
'''            Case Else
'''                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
'''            End Select
'''            i = i + 1
'''        Loop
'''    Else
'''        xHasil = "INSERT INTO " + xTabel + " ("
'''        i = 0
'''        Do While gsSql1(i) <> ""
'''            If i <> 0 Then
'''                xHasil = xHasil + ","
'''            End If
'''            xHasil = xHasil + gsSql1(i)
'''            i = i + 1
'''        Loop
'''        xHasil = xHasil + ") VALUES ("
'''
'''        i = 0
'''        Do While gsSql1(i) <> ""
'''            If i <> 0 Then
'''                xHasil = xHasil + ","
'''            End If
'''
'''            aa = UCase(gsSql1(i))
'''            Select Case aa
'''            Case "PPN", "DISC"
'''                xHasil = xHasil + gsSql2(i)
'''            Case Else
'''                xHasil = xHasil + "'" + gsSql2(i) + "'"
'''            End Select
'''            i = i + 1
'''        Loop
'''        xHasil = xHasil + ")"
'''    End If
'''    GensSql = xHasil
'''End Function
'''
'''Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    FormDrag Me
'''End Sub
'''
'''Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    FormDrag Me
'''End Sub
'''
'''Private Sub cmdExit_Click()
'''    Unload Me
'''End Sub
'''
'''Private Sub cmdSimpan_Click()
'''Dim rsSearch As ADODB.Recordset
'''
'''    Call ClearsSql
'''    gsSql1(0) = "FS_KD_LABOUR_INT"
'''    gsSql2(0) = Text(0)
'''    gsSql1(1) = "FS_KD_METERAI_INT"
'''    gsSql2(1) = Text(5)
'''    gsSql1(2) = "FS_KD_PPN_INT"
'''    gsSql2(2) = Text(10)
'''    gsSql1(3) = "FS_KD_PIUTANG_INT"
'''    gsSql2(3) = Text(6)
'''    gsSql1(4) = "FS_KD_LABOUR_EXT"
'''    gsSql2(4) = Text(2)
'''    gsSql1(5) = "FS_KD_METERAI_EXT"
'''    gsSql2(5) = Text(4)
'''    gsSql1(6) = "FS_KD_PPN_EXT"
'''    gsSql2(6) = Text(14)
'''    gsSql1(7) = "FS_KD_PIUTANG_EXT"
'''    gsSql2(7) = Text(12)
'''    gsSql1(8) = "FS_KD_LABOUR_AFFNONPPN"
'''    gsSql2(8) = Text(58)
'''    gsSql1(9) = "FS_KD_METERAI_AFFNONPPN"
'''    gsSql2(9) = Text(56)
'''    gsSql1(10) = "FS_KD_PIUTANG_AFFNONPPN"
'''    gsSql2(10) = Text(57)
'''    gsSql1(11) = "FS_KD_PPN_AFFNONPPN"
'''    gsSql2(11) = Text(62)
'''    gsSql1(12) = "FS_KD_LABOUR_AFFPPN"
'''    gsSql2(12) = Text(48)
'''    gsSql1(13) = "FS_KD_METERAI_AFFPPN"
'''    gsSql2(13) = Text(49)
'''    gsSql1(14) = "FS_KD_PIUTANG_AFFPPN"
'''    gsSql2(14) = Text(52)
'''    gsSql1(15) = "FS_KD_PPN_AFFPPN"
'''    gsSql2(15) = Text(54)
'''
'''    ' RETUR
'''    gsSql1(16) = "FS_KD_RLABOUR_INT"
'''    gsSql2(16) = Text(26)
'''    gsSql1(17) = "FS_KD_RMETERAI_INT"
'''    gsSql2(17) = Text(24)
'''    gsSql1(18) = "FS_KD_RPIUTANG_INT"
'''    gsSql2(18) = Text(25)
'''    gsSql1(19) = "FS_KD_RPPN_INT"
'''    gsSql2(19) = Text(30)
'''    gsSql1(20) = "FS_KD_RLABOUR_EXT"
'''    gsSql2(20) = Text(16)
'''    gsSql1(21) = "FS_KD_RMETERAI_EXT"
'''    gsSql2(21) = Text(17)
'''    gsSql1(22) = "FS_KD_RPIUTANG_EXT"
'''    gsSql2(22) = Text(20)
'''    gsSql1(23) = "FS_KD_RPPN_EXT"
'''    gsSql2(23) = Text(22)
'''
'''    gsSql1(24) = "FS_KD_RLABOUR_AFFNONPPN"
'''    gsSql2(24) = Text(34)
'''    gsSql1(25) = "FS_KD_RMETERAI_AFFNONPPN"
'''    gsSql2(25) = Text(32)
'''    gsSql1(26) = "FS_KD_RPIUTANG_AFFNONPPN"
'''    gsSql2(26) = Text(33)
'''    gsSql1(27) = "FS_KD_RPPN_AFFNONPPN"
'''    gsSql2(27) = Text(38)
'''    gsSql1(28) = "FS_KD_RLABOUR_AFFPPN"
'''    gsSql2(28) = Text(40)
'''    gsSql1(29) = "FS_KD_RMETERAI_AFFPPN"
'''    gsSql2(29) = Text(41)
'''    gsSql1(30) = "FS_KD_RPIUTANG_AFFPPN"
'''    gsSql2(30) = Text(44)
'''    gsSql1(31) = "FS_KD_RPPN_AFFPPN"
'''    gsSql2(31) = Text(46)
'''
'''    gsSql1(32) = "FS_KD_PDP"
'''    gsSql2(32) = Text(65)
'''    gsSql1(33) = "FS_KD_HUTANG"
'''    gsSql2(33) = Text(68)
'''
'''    gsSql1(34) = "FS_KD_SUBLET_INT"
'''    gsSql2(34) = Text(70)
'''    gsSql1(35) = "FS_KD_SUBLET_EXT"
'''    gsSql2(35) = Text(72)
'''    gsSql1(36) = "FS_KD_SUBLET_AFFNONPPN"
'''    gsSql2(36) = Text(74)
'''    gsSql1(37) = "FS_KD_SUBLET_AFFPPN"
'''    gsSql2(37) = Text(76)
'''    gsSql1(38) = "FS_KD_RSUBLET_INT"
'''    gsSql2(38) = Text(78)
'''    gsSql1(39) = "FS_KD_RSUBLET_EXT"
'''    gsSql2(39) = Text(80)
'''    gsSql1(40) = "FS_KD_RSUBLET_AFFNONPPN"
'''    gsSql2(40) = Text(82)
'''    gsSql1(41) = "FS_KD_RSUBLET_AFFPPN"
'''    gsSql2(41) = Text(84)
'''
'''    gsSql1(42) = "FS_KD_HPP_EXT"
'''    gsSql2(42) = Text(64)
'''    gsSql1(43) = "FS_KD_HPP_INT"
'''    gsSql2(43) = Text(88)
'''    gsSql1(44) = "FS_KD_HPP_AFFNONPPN"
'''    gsSql2(44) = Text(92)
'''    gsSql1(45) = "FS_KD_HPP_AFFPPN"
'''    gsSql2(45) = Text(94)
'''
'''    gsSql1(46) = "FS_KD_RHPP_EXT"
'''    gsSql2(46) = Text(86)
'''    gsSql1(47) = "FS_KD_RHPP_INT"
'''    gsSql2(47) = Text(90)
'''    gsSql1(48) = "FS_KD_RHPP_AFFNONPPN"
'''    gsSql2(48) = Text(98)
'''    gsSql1(49) = "FS_KD_RHPP_AFFPPN"
'''    gsSql2(49) = Text(96)
'''
'''    gsSql1(50) = "FS_KD_PPN"
'''    gsSql2(50) = Text(100)
'''
'''    Set rsSearch = objO.FetchData("Select * from T_AKUN")
'''    If Not rsSearch.BOF Then
'''        sSQL = GensSql("T_AKUN", True)
'''    Else
'''        sSQL = GensSql("T_AKUN", False)
'''    End If
'''    Call objO.UpdateData(sSQL)
'''
'''    gsSql1(51) = "userEntry"
'''    gsSql2(51) = gSA
'''    gsSql1(52) = "tglEntry"
'''    gsSql2(52) = Format(isnow, "MM/DD/YYYY HH:NN:SS")
'''    gsSql1(53) = "clientName"
'''    gsSql2(53) = ComptName
'''
'''    Call objO.UpdateData(sSQL)
'''    Text(0).SetFocus
'''End Sub
'''
'''Private Sub Form_Load()
'''Dim rsSearch As ADODB.Recordset
'''
'''    Call hapusForm
'''
'''    Set rsSearch = objO.FetchData("Select * from T_AKUN")
'''    If Not rsSearch.BOF Then
'''        Text(0) = rsSearch("FS_KD_LABOUR_INT")
'''        Text(70) = rsSearch("FS_KD_SUBLET_INT")
'''        Text(5) = rsSearch("FS_KD_METERAI_INT")
'''        Text(10) = rsSearch("FS_KD_PPN_INT")
'''        Text(6) = rsSearch("FS_KD_PIUTANG_INT")
'''        Text(88) = rsSearch("FS_KD_HPP_INT")
'''
'''        Text(2) = rsSearch("FS_KD_LABOUR_EXT")
'''        Text(72) = rsSearch("FS_KD_SUBLET_EXT")
'''        Text(4) = rsSearch("FS_KD_METERAI_EXT")
'''        Text(14) = rsSearch("FS_KD_PPN_EXT")
'''        Text(12) = rsSearch("FS_KD_PIUTANG_EXT")
'''        Text(64) = rsSearch("FS_KD_HPP_EXT")
'''
'''        Text(58) = rsSearch("FS_KD_LABOUR_AFFNONPPN")
'''        Text(74) = rsSearch("FS_KD_SUBLET_AFFNONPPN")
'''        Text(56) = rsSearch("FS_KD_METERAI_AFFNONPPN")
'''        Text(62) = rsSearch("FS_KD_PPN_AFFNONPPN")
'''        Text(57) = rsSearch("FS_KD_PIUTANG_AFFNONPPN")
'''        Text(92) = rsSearch("FS_KD_HPP_AFFNONPPN")
'''
'''        Text(48) = rsSearch("FS_KD_LABOUR_AFFPPN")
'''        Text(76) = rsSearch("FS_KD_SUBLET_AFFPPN")
'''        Text(49) = rsSearch("FS_KD_METERAI_AFFPPN")
'''        Text(54) = rsSearch("FS_KD_PPN_AFFPPN")
'''        Text(52) = rsSearch("FS_KD_PIUTANG_AFFPPN")
'''        Text(94) = rsSearch("FS_KD_HPP_AFFPPN")
'''
'''
'''        ''' RETUR
'''        Text(26) = rsSearch("FS_KD_RLABOUR_INT")
'''        Text(78) = rsSearch("FS_KD_RSUBLET_INT")
'''        Text(24) = rsSearch("FS_KD_RMETERAI_INT")
'''        Text(25) = rsSearch("FS_KD_RPIUTANG_INT")
'''        Text(30) = rsSearch("FS_KD_RPPN_INT")
'''        Text(90) = rsSearch("FS_KD_RHPP_INT")
'''
'''        Text(16) = rsSearch("FS_KD_RLABOUR_EXT")
'''        Text(80) = rsSearch("FS_KD_RSUBLET_EXT")
'''        Text(17) = rsSearch("FS_KD_RMETERAI_EXT")
'''        Text(20) = rsSearch("FS_KD_RPIUTANG_EXT")
'''        Text(22) = rsSearch("FS_KD_RPPN_EXT")
'''        Text(86) = rsSearch("FS_KD_RHPP_EXT")
'''
'''        Text(34) = rsSearch("FS_KD_RLABOUR_AFFNONPPN")
'''        Text(82) = rsSearch("FS_KD_RSUBLET_AFFNONPPN")
'''        Text(32) = rsSearch("FS_KD_RMETERAI_AFFNONPPN")
'''        Text(38) = rsSearch("FS_KD_RPPN_AFFNONPPN")
'''        Text(33) = rsSearch("FS_KD_RPIUTANG_AFFNONPPN")
'''        Text(98) = rsSearch("FS_KD_RHPP_AFFNONPPN")
'''
'''        Text(40) = rsSearch("FS_KD_RLABOUR_AFFPPN")
'''        Text(84) = rsSearch("FS_KD_RSUBLET_AFFPPN")
'''        Text(41) = rsSearch("FS_KD_RMETERAI_AFFPPN")
'''        Text(46) = rsSearch("FS_KD_RPPN_AFFPPN")
'''        Text(44) = rsSearch("FS_KD_RPIUTANG_AFFPPN")
'''        Text(96) = rsSearch("FS_KD_RHPP_AFFPPN")
'''
'''        Text(65) = rsSearch("FS_KD_PDP")
'''        Text(68) = rsSearch("FS_KD_HUTANG")
'''        Text(100) = rsSearch("FS_KD_PPN")
'''    End If
'''
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(0)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(1) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(5)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(7) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(10)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(11) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(6)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(9) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(2)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(3) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(4)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(8) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(14)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(15) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(12)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(13) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(58)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(59) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(56)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(60) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(62)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(63) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(57)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(61) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(48)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(50) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(49)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(51) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(54)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(55) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(52)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(53) = rsSearch("namagl")
'''        End If
'''
'''        ''RETUR
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(26)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(27) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(24)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(28) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(30)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(31) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(25)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(29) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(16)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(18) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(17)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(19) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(22)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(23) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(20)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(21) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(34)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(35) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(32)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(36) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(38)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(39) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(33)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(37) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(40)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(42) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(41)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(43) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(46)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(47) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(44)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(45) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(65)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(67) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(68)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(69) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(70)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(71) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(72)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(73) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(74)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(75) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(76)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(77) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(78)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(79) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(80)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(81) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(82)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(83) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(84)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(85) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(64)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(66) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(88)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(89) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(92)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(93) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(94)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(95) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(86)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(87) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(90)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(91) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(98)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(99) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(96)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(97) = rsSearch("namagl")
'''        End If
'''
'''        Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(100)) & "'")
'''        If Not rsSearch.BOF Then
'''            Text(101) = rsSearch("namagl")
'''        End If
'''End Sub
'''
'''Private Sub Form_Resize()
'''    asxPowerBanner1.Height = Me.Height - 500
'''End Sub
'''
'''Private Sub Text_GotFocus(Index As Integer)
'''    Dim DiTab As Boolean
'''    DiTab = False
'''    If Text(Index).BackColor = &HC0E0FF Then
'''        DiTab = True
'''    Else
'''        If gMainKeyCari <> "" Then
'''            Text(Index) = gMainKeyCari
'''            gMainKeyCari = ""
'''            DiTab = True
'''        End If
'''    End If
'''    If DiTab Then
'''        SendKeys ("{TAB}")
'''    Else
'''        SendKeys ("{HOME}+{END}")
'''    End If
'''End Sub
'''
'''Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'''    Call CekTombol(KeyCode)
'''    If KeyCode = vbKeyF12 Then
'''        Select Case Index
'''            Case 0, 2, 4, 5, 6, 10, 12, 14, 16, 17, 20, 22, 24, 25, 26, 30, 58, 56, 62, 57, 48, 49, 54, 52, 34, 32, 38, 33, 40, 41, 46, 44, 64, 65, 68, 70, 72, 74, 76, 78, 80, 82, 84, 88, 92, 94, 86, 90, 98, 96
'''                gMainKeyCari = Text(Index)
'''                gSearchOption = "KODEJNRL"
'''                frmCariAcess.Show (vbModal)
'''        End Select
'''    End If
'''End Sub
'''
'''Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
'''    If KeyAscii = 39 Then
'''        KeyAscii = 0
'''        SendKeys ("´")
'''    End If
'''    If KeyAscii = 34 Then
'''        KeyAscii = 0
'''        SendKeys ("”")
'''    End If
'''End Sub
'''
'''Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
'''Dim rsSearch As ADODB.Recordset
'''
'''
'''    Select Case Index
'''        Case 0, 2, 4, 5, 6, 10, 12, 14, 16, 17, 20, 22, 24, 25, 26, 30, 58, 56, 62, 57, 48, 49, 54, 52, 34, 32, 38, 33, 40, 41, 46, 44, 64, 65, 68, 70, 72, 74, 76, 78, 80, 82, 84, 88, 92, 94, 86, 90, 98, 96, 100
'''            If Trim(Text(Index)) <> "" Then
'''
'''                    Set rsSearch = objO.FetchDataACC("select kodegl, namagl from glmst where kodegl = '" & Trim(Text(Index)) & "'")
'''                    If rsSearch.BOF Then
'''                        Cancel = True
'''                        gMainKeyCari = Text(Index)
'''                        gSearchOption = "KODEJNRL"
'''                        frmCariAcess.Show (vbModal)
'''                    Else
'''                        Select Case Index
'''                            Case 0
'''                                Text(1) = rsSearch("namagl")
'''                            Case 2
'''                                Text(3) = rsSearch("namagl")
'''                            Case 4
'''                                Text(8) = rsSearch("namagl")
'''                            Case 5
'''                                Text(7) = rsSearch("namagl")
'''                            Case 6
'''                                Text(9) = rsSearch("namagl")
'''                            Case 10
'''                                Text(11) = rsSearch("namagl")
'''                            Case 12
'''                                Text(13) = rsSearch("namagl")
'''                            Case 14
'''                                Text(15) = rsSearch("namagl")
'''                            Case 16
'''                                Text(18) = rsSearch("namagl")
'''                            Case 17
'''                                Text(19) = rsSearch("namagl")
'''                            Case 20
'''                                Text(21) = rsSearch("namagl")
'''                            Case 22
'''                                Text(23) = rsSearch("namagl")
'''                            Case 24
'''                                Text(28) = rsSearch("namagl")
'''                            Case 25
'''                                Text(29) = rsSearch("namagl")
'''                            Case 26
'''                                Text(27) = rsSearch("namagl")
'''                            Case 30
'''                                Text(31) = rsSearch("namagl")
'''                            Case 58
'''                                Text(59) = rsSearch("namagl")
'''                            Case 56
'''                                Text(60) = rsSearch("namagl")
'''                            Case 62
'''                                Text(63) = rsSearch("namagl")
'''                            Case 57
'''                                Text(61) = rsSearch("namagl")
'''                            Case 48
'''                                Text(50) = rsSearch("namagl")
'''                            Case 49
'''                                Text(51) = rsSearch("namagl")
'''                            Case 54
'''                                Text(55) = rsSearch("namagl")
'''                            Case 52
'''                                Text(53) = rsSearch("namagl")
'''                            Case 34
'''                                Text(35) = rsSearch("namagl")
'''                            Case 32
'''                                Text(36) = rsSearch("namagl")
'''                            Case 38
'''                                Text(39) = rsSearch("namagl")
'''                            Case 33
'''                                Text(37) = rsSearch("namagl")
'''                            Case 40
'''                                Text(42) = rsSearch("namagl")
'''                            Case 41
'''                                Text(43) = rsSearch("namagl")
'''                            Case 46
'''                                Text(47) = rsSearch("namagl")
'''                            Case 44
'''                                Text(45) = rsSearch("namagl")
'''                            Case 58
'''                                Text(59) = rsSearch("namagl")
'''                            Case 56
'''                                Text(60) = rsSearch("namagl")
'''                            Case 62
'''                                Text(63) = rsSearch("namagl")
'''                            Case 57
'''                                Text(61) = rsSearch("namagl")
'''                            Case 48
'''                                Text(50) = rsSearch("namagl")
'''                            Case 49
'''                                Text(51) = rsSearch("namagl")
'''                            Case 54
'''                                Text(55) = rsSearch("namagl")
'''                            Case 52
'''                                Text(53) = rsSearch("namagl")
'''                            Case 34
'''                                Text(35) = rsSearch("namagl")
'''                            Case 32
'''                                Text(36) = rsSearch("namagl")
'''                            Case 38
'''                                Text(39) = rsSearch("namagl")
'''                            Case 33
'''                                Text(37) = rsSearch("namagl")
'''                            Case 40
'''                                Text(42) = rsSearch("namagl")
'''                            Case 41
'''                                Text(43) = rsSearch("namagl")
'''                            Case 46
'''                                Text(47) = rsSearch("namagl")
'''                            Case 44
'''                                Text(45) = rsSearch("namagl")
'''
'''                            Case 64
'''                                Text(66) = rsSearch("namagl")
'''                            Case 65
'''                                Text(67) = rsSearch("namagl")
'''                            Case 68
'''                                Text(69) = rsSearch("namagl")
'''
'''                            Case 70
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 72
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 74
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 76
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 78
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 80
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 82
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 84
'''                                Text(Index + 1) = rsSearch("namagl")
'''
'''                            Case 88
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 92
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 94
'''                                Text(Index + 1) = rsSearch("namagl")
'''
'''                            Case 86
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 90
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 98
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 96
'''                                Text(Index + 1) = rsSearch("namagl")
'''                            Case 100
'''                                Text(Index + 1) = rsSearch("namagl")
'''                        End Select
'''                    End If
'''
'''            Else
'''                Select Case Index
'''                    Case 0
'''                        Text(1) = ""
'''                    Case 2
'''                        Text(3) = ""
'''                    Case 4
'''                        Text(8) = ""
'''                    Case 5
'''                        Text(7) = ""
'''                    Case 6
'''                        Text(9) = ""
'''                    Case 10
'''                        Text(11) = ""
'''                    Case 12
'''                        Text(13) = ""
'''                    Case 14
'''                        Text(15) = ""
'''                    Case 16
'''                        Text(18) = ""
'''                    Case 17
'''                        Text(19) = ""
'''                    Case 20
'''                        Text(21) = ""
'''                    Case 22
'''                        Text(23) = ""
'''                    Case 24
'''                        Text(28) = ""
'''                    Case 25
'''                        Text(29) = ""
'''                    Case 26
'''                        Text(27) = ""
'''                    Case 30
'''                        Text(31) = ""
'''                    Case 58
'''                        Text(59) = ""
'''                    Case 56
'''                        Text(60) = ""
'''                    Case 62
'''                        Text(63) = ""
'''                    Case 57
'''                        Text(61) = ""
'''                    Case 48
'''                        Text(50) = ""
'''                    Case 49
'''                        Text(51) = ""
'''                    Case 54
'''                        Text(55) = ""
'''                    Case 52
'''                        Text(53) = ""
'''                    Case 34
'''                        Text(35) = ""
'''                    Case 32
'''                        Text(36) = ""
'''                    Case 38
'''                        Text(39) = ""
'''                    Case 33
'''                        Text(37) = ""
'''                    Case 40
'''                        Text(42) = ""
'''                    Case 41
'''                        Text(43) = ""
'''                    Case 46
'''                        Text(47) = ""
'''                    Case 44
'''                        Text(45) = ""
'''
'''                    Case 64
'''                        Text(66) = ""
'''                    Case 65
'''                        Text(67) = ""
'''                    Case 68
'''                        Text(69) = ""
'''                    Case 70
'''                        Text(Index + 1) = ""
'''                    Case 72
'''                        Text(Index + 1) = ""
'''                    Case 74
'''                        Text(Index + 1) = ""
'''                    Case 76
'''                        Text(Index + 1) = ""
'''                    Case 78
'''                        Text(Index + 1) = ""
'''                    Case 80
'''                        Text(Index + 1) = ""
'''                    Case 82
'''                        Text(Index + 1) = ""
'''                    Case 84
'''                        Text(Index + 1) = ""
'''
'''                    Case 88
'''                        Text(Index + 1) = ""
'''                    Case 92
'''                        Text(Index + 1) = ""
'''                    Case 94
'''                        Text(Index + 1) = ""
'''
'''                    Case 86
'''                        Text(Index + 1) = ""
'''                    Case 90
'''                        Text(Index + 1) = ""
'''                    Case 98
'''                        Text(Index + 1) = ""
'''                    Case 96
'''                        Text(Index + 1) = ""
'''                    Case 100
'''                        Text(Index + 1) = ""
'''                End Select
'''            End If
'''    End Select
'''
'''End Sub
'''
'''Private Sub hapusForm()
'''    Dim i As Integer
'''
'''    For i = 0 To 101
'''        Text(i) = ""
'''    Next
'''End Sub


'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private rsB As ADODB.Recordset
Private IsNew As Boolean



Private Sub Form_Load()
    Call ViewMasterAkun
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
        Dim C As Integer
        If MsgBox("Simpan Master Akun  ?", vbQuestion + vbYesNo, "Master Akun") = vbYes Then
        
            For C = 0 To 65
                objO.nID = C
                objO.nTxtAcc = txt(C).Text
            Next C
            
            For C = 66 To 69
                objO.nID = C
                objO.nTxtAcc = ""
            Next C
            
            Call objO.CreateNew_AkunNew(0, IsNew)
            Call objO.CreateNew_AkunNew(1, IsNew)
            Call objO.CreateNew_AkunNew(2, IsNew)
            Call objO.CreateNew_AkunNew(3, IsNew)
            Call objO.CreateNew_AkunNew(4, IsNew)
            Call objO.CreateNew_AkunNew(5, IsNew)
            Call objO.CreateNew_AkunNew(6, IsNew)
        End If
End Sub


Private Sub ViewMasterAkun()
    Dim E As Integer
    Dim C As Integer
    Set rsB = objO.Get_AkunNew
    If rsB.RecordCount = 0 Then
        IsNew = True
        Exit Sub
    Else
        IsNew = False
    End If
    Dim nAkun As String
    nAkun = ""
    E = 0
    rsB.MoveFirst
    Do Until rsB.EOF
        For C = 1 To 10
            If (C + (E * 10) - 1) < 66 Then
                txt(C + (E * 10) - 1).Text = rsB.Fields(C).value
                If rsB.Fields(C).value <> "" Then
                    If nAkun = "" Then
                        nAkun = "'" & rsB.Fields(C).value & "'"
                    Else
                        nAkun = nAkun & "," & "'" & rsB.Fields(C).value & "'"
                    End If
                End If
            End If
        Next C
        rsB.MoveNext
        E = E + 1
    Loop

    Dim rsN As ADODB.Recordset
    Set rsN = objO.Get_KodeAkun(" Where kodegl in (" & nAkun & ")")
    For C = 0 To 65
        If txt(C).Text <> "" Then
            rsN.Find "kodegl = '" & Trim(txt(C).Text) & "'", , adSearchForward, 1
            If Not rsN.EOF Then
                txtx(C).Text = rsN.Fields(1).value
            End If
        End If
    Next C
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        gMainKeyCari = txt(Index)
        gSearchOption = "KODEJNRL"
        frmCariAcess.Show (vbModal)
        If gMainKeyCari <> "" Then
            txt(Index).Text = gMainKeyCari
            txtx(Index).Text = objO.FetchData_String_Acc("select NamaGL from GLMSt WITH(NOLOCK) Where kodeGL = '" & gMainKeyCari & "'")
        End If
    End If
End Sub

Private Sub txtx_GotFocus(Index As Integer)
    Sendkeys "{tab}"
End Sub

