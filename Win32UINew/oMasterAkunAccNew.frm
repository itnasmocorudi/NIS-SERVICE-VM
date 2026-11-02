VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form oMasterAkunAccNew 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "oMasterAkunAcc"
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   Icon            =   "oMasterAkunAccNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "oMasterAkunAccNew.frx":29C12
   ScaleHeight     =   10500
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   10000
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin TabDlg.SSTab SSTab1 
         Height          =   9225
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   16272
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   706
         ShowFocusRect   =   0   'False
         BackColor       =   -2147483626
         TabCaption(0)   =   "-"
         TabPicture(0)   =   "oMasterAkunAccNew.frx":2BD78
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Line2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSTab3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "."
         TabPicture(1)   =   "oMasterAkunAccNew.frx":2C4CC
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label5(23)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label9(15)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Label6(8)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label12(16)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label3(16)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label5(24)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Shape1(44)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Shape1(45)"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Shape1(46)"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "Shape1(81)"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "Shape1(82)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "Shape1(83)"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "Label5(50)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "Label3(31)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "Label12(25)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "Label6(19)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "Label9(25)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "Label5(51)"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "Label5(52)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "Label5(53)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "txt(78)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "txtx(78)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "txt(79)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "txtx(79)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "txt(80)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).Control(25)=   "txtx(80)"
         Tab(1).Control(25).Enabled=   0   'False
         Tab(1).Control(26)=   "txtx(83)"
         Tab(1).Control(26).Enabled=   0   'False
         Tab(1).Control(27)=   "txt(83)"
         Tab(1).Control(27).Enabled=   0   'False
         Tab(1).Control(28)=   "txtx(82)"
         Tab(1).Control(28).Enabled=   0   'False
         Tab(1).Control(29)=   "txt(82)"
         Tab(1).Control(29).Enabled=   0   'False
         Tab(1).Control(30)=   "txtx(81)"
         Tab(1).Control(30).Enabled=   0   'False
         Tab(1).Control(31)=   "txt(81)"
         Tab(1).Control(31).Enabled=   0   'False
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
            Index           =   81
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   348
            Top             =   3000
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
            Index           =   81
            Left            =   4485
            MaxLength       =   40
            TabIndex        =   347
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
            Index           =   82
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   346
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
            Index           =   82
            Left            =   4485
            MaxLength       =   40
            TabIndex        =   345
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
            Index           =   83
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   344
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
            Index           =   83
            Left            =   4485
            MaxLength       =   40
            TabIndex        =   343
            Top             =   3720
            Width           =   6000
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   6855
            Left            =   -74880
            TabIndex        =   16
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
            TabPicture(0)   =   "oMasterAkunAccNew.frx":2CA74
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Line4(0)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSTab4"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "NON DYNA"
            TabPicture(1)   =   "oMasterAkunAccNew.frx":2CA90
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "SSTab5"
            Tab(1).Control(1)=   "Line4(1)"
            Tab(1).ControlCount=   2
            Begin TabDlg.SSTab SSTab4 
               Height          =   6255
               Left            =   120
               TabIndex        =   17
               Top             =   480
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   11033
               _Version        =   393216
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
               TabPicture(0)   =   "oMasterAkunAccNew.frx":2CAAC
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Line4(2)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Frame3(0)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "INTERNAL (Pemakaian Sendiri)"
               TabPicture(1)   =   "oMasterAkunAccNew.frx":2CAC8
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Line4(3)"
               Tab(1).Control(1)=   "Line4(4)"
               Tab(1).Control(2)=   "Frame3(1)"
               Tab(1).ControlCount=   3
               TabCaption(2)   =   "AFILIASI (PPn)"
               TabPicture(2)   =   "oMasterAkunAccNew.frx":2CAE4
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Frame3(2)"
               Tab(2).Control(1)=   "Line4(5)"
               Tab(2).ControlCount=   2
               Begin VB.Frame Frame3 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   5175
                  Index           =   2
                  Left            =   -74520
                  TabIndex        =   125
                  Top             =   720
                  Width           =   9975
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   151
                     Top             =   3960
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   150
                     Top             =   3630
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   149
                     Top             =   3300
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   148
                     Top             =   2970
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
                     Left            =   3720
                     MaxLength       =   40
                     TabIndex        =   147
                     Top             =   2640
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   146
                     Top             =   2310
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   145
                     Top             =   1980
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   144
                     Top             =   1650
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   143
                     Top             =   1320
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
                     Index           =   29
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   142
                     Top             =   990
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   141
                     Top             =   660
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   140
                     Top             =   330
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   139
                     Top             =   0
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
                     Index           =   38
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   138
                     Top             =   3960
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   137
                     Top             =   3630
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   136
                     Top             =   3300
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   135
                     Top             =   2970
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   134
                     Top             =   2640
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   133
                     Top             =   2310
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   132
                     Top             =   1980
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   131
                     Top             =   1650
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   130
                     Top             =   1320
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
                     Index           =   29
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   129
                     Top             =   990
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   128
                     Top             =   660
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   127
                     Top             =   330
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   126
                     Top             =   0
                     Width           =   1130
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "DYNA - AFILIASI (PPn)"
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
                     Index           =   46
                     Left            =   6840
                     TabIndex        =   339
                     Top             =   4560
                     Width           =   2850
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   38
                     Left            =   3630
                     Top             =   3300
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   37
                     Left            =   3630
                     Top             =   2970
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   36
                     Left            =   3630
                     Top             =   2640
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
                     Left            =   2400
                     TabIndex        =   177
                     Top             =   3300
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
                     Left            =   2400
                     TabIndex        =   176
                     Top             =   2640
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
                     Left            =   2400
                     TabIndex        =   175
                     Top             =   2970
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
                     Left            =   2400
                     TabIndex        =   174
                     Top             =   330
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
                     Left            =   2400
                     TabIndex        =   173
                     Top             =   0
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
                     Left            =   2400
                     TabIndex        =   172
                     Top             =   660
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
                     Left            =   2400
                     TabIndex        =   171
                     Top             =   1650
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
                     Left            =   2400
                     TabIndex        =   170
                     Top             =   1320
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
                     Left            =   2400
                     TabIndex        =   169
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   35
                     Left            =   3630
                     Top             =   0
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   34
                     Left            =   3630
                     Top             =   330
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   33
                     Left            =   3630
                     Top             =   660
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   32
                     Left            =   3630
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   31
                     Left            =   3630
                     Top             =   1320
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   30
                     Left            =   3630
                     Top             =   1650
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
                     Left            =   2400
                     TabIndex        =   168
                     Top             =   2310
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
                     Left            =   2400
                     TabIndex        =   167
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   29
                     Left            =   3630
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   28
                     Left            =   3630
                     Top             =   2310
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   27
                     Left            =   3630
                     Top             =   3960
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   26
                     Left            =   3630
                     Top             =   3630
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
                     Left            =   2400
                     TabIndex        =   166
                     Top             =   3940
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
                     Left            =   2400
                     TabIndex        =   165
                     Top             =   3630
                     Width           =   60
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "HPP Jasa BP"
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
                     Left            =   0
                     TabIndex        =   164
                     Top             =   3000
                     Width           =   1170
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
                     Left            =   240
                     TabIndex        =   163
                     Top             =   2660
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
                     Index           =   16
                     Left            =   240
                     TabIndex        =   162
                     Top             =   2340
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
                     Index           =   15
                     Left            =   0
                     TabIndex        =   161
                     Top             =   3660
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
                     Index           =   9
                     Left            =   240
                     TabIndex        =   160
                     Top             =   3980
                     Width           =   1230
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Hutang Pemborong"
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
                     Left            =   240
                     TabIndex        =   159
                     Top             =   3315
                     Width           =   1860
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
                     Index           =   14
                     Left            =   240
                     TabIndex        =   158
                     Top             =   1320
                     Width           =   960
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Penj. Jasa BP"
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
                     Left            =   240
                     TabIndex        =   157
                     Top             =   660
                     Width           =   1290
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
                     Left            =   240
                     TabIndex        =   156
                     Top             =   990
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
                     Left            =   240
                     TabIndex        =   155
                     Top             =   2010
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
                     Index           =   12
                     Left            =   240
                     TabIndex        =   154
                     Top             =   1680
                     Width           =   930
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Asuransi"
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
                     Left            =   0
                     TabIndex        =   153
                     Top             =   330
                     Width           =   1650
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Umum"
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
                     Left            =   0
                     TabIndex        =   152
                     Top             =   0
                     Width           =   1410
                  End
               End
               Begin VB.Frame Frame3 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   5175
                  Index           =   1
                  Left            =   -74520
                  TabIndex        =   72
                  Top             =   720
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
                     Index           =   13
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   98
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
                     Index           =   14
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   97
                     Top             =   330
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   96
                     Top             =   660
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   95
                     Top             =   990
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   94
                     Top             =   1320
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   93
                     Top             =   1650
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
                     Index           =   19
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   92
                     Top             =   1980
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   91
                     Top             =   2310
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   90
                     Top             =   2640
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   89
                     Top             =   2970
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   88
                     Top             =   3300
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   87
                     Top             =   3630
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   86
                     Top             =   3960
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
                     Index           =   13
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   85
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
                     Index           =   14
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   84
                     Top             =   330
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   83
                     Top             =   660
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   82
                     Top             =   990
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   81
                     Top             =   1320
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   80
                     Top             =   1650
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
                     Index           =   19
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   79
                     Top             =   1980
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   78
                     Top             =   2310
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   77
                     Top             =   2640
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   76
                     Top             =   2970
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   75
                     Top             =   3300
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   74
                     Top             =   3630
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   73
                     Top             =   3960
                     Width           =   6000
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "DYNA - INTERNAL (Pemakaian Sendiri)"
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
                     Index           =   45
                     Left            =   4920
                     TabIndex        =   338
                     Top             =   4560
                     Width           =   6105
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Umum"
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
                     Left            =   0
                     TabIndex        =   124
                     Top             =   0
                     Width           =   1410
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Asuransi"
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
                     Left            =   0
                     TabIndex        =   123
                     Top             =   330
                     Width           =   1650
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
                     Index           =   9
                     Left            =   240
                     TabIndex        =   122
                     Top             =   1680
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
                     Index           =   2
                     Left            =   240
                     TabIndex        =   121
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
                     Index           =   4
                     Left            =   240
                     TabIndex        =   120
                     Top             =   990
                     Width           =   1170
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Penj. Jasa BP"
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
                     Left            =   240
                     TabIndex        =   119
                     Top             =   660
                     Width           =   1290
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
                     Left            =   240
                     TabIndex        =   118
                     Top             =   1320
                     Width           =   960
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Hutang Pemborong"
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
                     Left            =   240
                     TabIndex        =   117
                     Top             =   3315
                     Width           =   1860
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
                     Left            =   240
                     TabIndex        =   116
                     Top             =   3980
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
                     Index           =   6
                     Left            =   0
                     TabIndex        =   115
                     Top             =   3660
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
                     Index           =   5
                     Left            =   240
                     TabIndex        =   114
                     Top             =   2340
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
                     Index           =   1
                     Left            =   240
                     TabIndex        =   113
                     Top             =   2660
                     Width           =   735
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "HPP Jasa BP"
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
                     Left            =   0
                     TabIndex        =   112
                     Top             =   3000
                     Width           =   1170
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
                     Left            =   2400
                     TabIndex        =   111
                     Top             =   3630
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
                     Left            =   2400
                     TabIndex        =   110
                     Top             =   3940
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   25
                     Left            =   3630
                     Top             =   3630
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   24
                     Left            =   3630
                     Top             =   3960
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   23
                     Left            =   3630
                     Top             =   2310
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   22
                     Left            =   3630
                     Top             =   1980
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
                     Left            =   2400
                     TabIndex        =   109
                     Top             =   1980
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
                     Left            =   2400
                     TabIndex        =   108
                     Top             =   2310
                     Width           =   45
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   21
                     Left            =   3630
                     Top             =   1650
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   20
                     Left            =   3630
                     Top             =   1320
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   19
                     Left            =   3630
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   18
                     Left            =   3630
                     Top             =   660
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   17
                     Left            =   3630
                     Top             =   330
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   16
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
                     Index           =   6
                     Left            =   2400
                     TabIndex        =   107
                     Top             =   990
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
                     Left            =   2400
                     TabIndex        =   106
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
                     Index           =   7
                     Left            =   2400
                     TabIndex        =   105
                     Top             =   1650
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
                     Left            =   2400
                     TabIndex        =   104
                     Top             =   660
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
                     Left            =   2400
                     TabIndex        =   103
                     Top             =   0
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
                     Left            =   2400
                     TabIndex        =   102
                     Top             =   330
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
                     Left            =   2400
                     TabIndex        =   101
                     Top             =   2970
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
                     Index           =   4
                     Left            =   2400
                     TabIndex        =   100
                     Top             =   2640
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
                     Left            =   2400
                     TabIndex        =   99
                     Top             =   3300
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   15
                     Left            =   3630
                     Top             =   2640
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   14
                     Left            =   3630
                     Top             =   2970
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   13
                     Left            =   3630
                     Top             =   3300
                     Width           =   60
                  End
               End
               Begin VB.Frame Frame3 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   5175
                  Index           =   0
                  Left            =   480
                  TabIndex        =   19
                  Top             =   720
                  Width           =   9975
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   35
                     Top             =   3960
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   33
                     Top             =   3630
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   59
                     Top             =   3300
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
                     TabIndex        =   57
                     Top             =   2970
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
                     TabIndex        =   55
                     Top             =   2640
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
                     TabIndex        =   40
                     Top             =   2310
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
                     TabIndex        =   42
                     Top             =   1980
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
                     TabIndex        =   44
                     Top             =   1650
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
                     TabIndex        =   46
                     Top             =   1320
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
                     TabIndex        =   48
                     Top             =   990
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
                     TabIndex        =   50
                     Top             =   660
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
                     TabIndex        =   52
                     Top             =   330
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
                     Index           =   0
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   54
                     Top             =   0
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
                     Index           =   12
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   36
                     Top             =   3960
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   34
                     Top             =   3630
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   60
                     Top             =   3300
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
                     TabIndex        =   58
                     Top             =   2970
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
                     TabIndex        =   56
                     Top             =   2640
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
                     TabIndex        =   39
                     Top             =   2310
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
                     TabIndex        =   41
                     Top             =   1980
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
                     TabIndex        =   43
                     Top             =   1650
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
                     TabIndex        =   45
                     Top             =   1320
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
                     TabIndex        =   47
                     Top             =   990
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
                     TabIndex        =   49
                     Top             =   660
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
                     TabIndex        =   51
                     Top             =   330
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
                     Index           =   0
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   53
                     Top             =   0
                     Width           =   1130
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
                     TabIndex        =   337
                     Top             =   4560
                     Width           =   4140
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   10
                     Left            =   3630
                     Top             =   3300
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   6
                     Left            =   3630
                     Top             =   2970
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   0
                     Left            =   3630
                     Top             =   2640
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
                     Left            =   2400
                     TabIndex        =   71
                     Top             =   3300
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
                     TabIndex        =   70
                     Top             =   2640
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
                     Left            =   2400
                     TabIndex        =   69
                     Top             =   2970
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
                     Left            =   2400
                     TabIndex        =   68
                     Top             =   330
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
                     TabIndex        =   67
                     Top             =   0
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
                     TabIndex        =   66
                     Top             =   660
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
                     TabIndex        =   65
                     Top             =   1650
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
                     TabIndex        =   64
                     Top             =   1320
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
                     TabIndex        =   63
                     Top             =   990
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
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   1
                     Left            =   3630
                     Top             =   330
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   2
                     Left            =   3630
                     Top             =   660
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   3
                     Left            =   3630
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   4
                     Left            =   3630
                     Top             =   1320
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   5
                     Left            =   3630
                     Top             =   1650
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
                     TabIndex        =   62
                     Top             =   2310
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
                     Index           =   2
                     Left            =   2400
                     TabIndex        =   61
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   7
                     Left            =   3630
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   9
                     Left            =   3630
                     Top             =   2310
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   12
                     Left            =   3630
                     Top             =   3960
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   11
                     Left            =   3630
                     Top             =   3630
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
                     TabIndex        =   38
                     Top             =   3940
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
                     Index           =   4
                     Left            =   2400
                     TabIndex        =   37
                     Top             =   3630
                     Width           =   60
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "HPP Jasa BP"
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
                     Left            =   0
                     TabIndex        =   32
                     Top             =   3000
                     Width           =   1170
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
                     TabIndex        =   31
                     Top             =   2660
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
                     Index           =   27
                     Left            =   240
                     TabIndex        =   30
                     Top             =   2340
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
                     Index           =   26
                     Left            =   0
                     TabIndex        =   29
                     Top             =   3660
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
                     Index           =   17
                     Left            =   240
                     TabIndex        =   28
                     Top             =   3980
                     Width           =   1230
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Hutang Pemborong"
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
                     Left            =   240
                     TabIndex        =   27
                     Top             =   3315
                     Width           =   1860
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
                     TabIndex        =   26
                     Top             =   1320
                     Width           =   960
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Penj. Jasa BP"
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
                     TabIndex        =   25
                     Top             =   660
                     Width           =   1290
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
                     TabIndex        =   24
                     Top             =   990
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
                     Index           =   0
                     Left            =   240
                     TabIndex        =   23
                     Top             =   2010
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
                     Index           =   2
                     Left            =   240
                     TabIndex        =   22
                     Top             =   1680
                     Width           =   930
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Asuransi"
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
                     Left            =   0
                     TabIndex        =   21
                     Top             =   330
                     Width           =   1650
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Umum"
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
                     TabIndex        =   20
                     Top             =   0
                     Width           =   1410
                  End
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H000080FF&
                  BorderWidth     =   2
                  Index           =   5
                  X1              =   -75000
                  X2              =   -67680
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
                  Index           =   3
                  X1              =   -75000
                  X2              =   -71280
                  Y1              =   360
                  Y2              =   360
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H000080FF&
                  BorderWidth     =   2
                  Index           =   2
                  X1              =   3600
                  X2              =   11040
                  Y1              =   360
                  Y2              =   360
               End
            End
            Begin TabDlg.SSTab SSTab5 
               Height          =   6015
               Left            =   -74880
               TabIndex        =   18
               Top             =   480
               Width           =   11055
               _ExtentX        =   19500
               _ExtentY        =   10610
               _Version        =   393216
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
               TabPicture(0)   =   "oMasterAkunAccNew.frx":2CB00
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Line4(6)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Frame3(3)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "INTERNAL (Pemakaian Sendiri)"
               TabPicture(1)   =   "oMasterAkunAccNew.frx":2CB1C
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Line4(8)"
               Tab(1).Control(1)=   "Line4(9)"
               Tab(1).Control(2)=   "Frame3(4)"
               Tab(1).ControlCount=   3
               TabCaption(2)   =   "AFILIASI (PPn)"
               TabPicture(2)   =   "oMasterAkunAccNew.frx":2CB38
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Frame3(5)"
               Tab(2).Control(1)=   "Line4(7)"
               Tab(2).ControlCount=   2
               Begin VB.Frame Frame3 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   5175
                  Index           =   5
                  Left            =   -74520
                  TabIndex        =   284
                  Top             =   720
                  Width           =   9975
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
                     Index           =   77
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   310
                     Top             =   3960
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
                     Index           =   76
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   309
                     Top             =   3630
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
                     Index           =   75
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   308
                     Top             =   3300
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
                     Index           =   74
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   307
                     Top             =   2970
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
                     Index           =   73
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   306
                     Top             =   2640
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
                     Index           =   72
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   305
                     Top             =   2310
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
                     Index           =   71
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   304
                     Top             =   1980
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
                     Index           =   70
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   303
                     Top             =   1650
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
                     Index           =   69
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   302
                     Top             =   1320
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
                     Index           =   68
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   301
                     Top             =   990
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
                     Index           =   67
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   300
                     Top             =   660
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
                     Index           =   66
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   299
                     Top             =   330
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   298
                     Top             =   0
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
                     Index           =   77
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   297
                     Top             =   3960
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
                     Index           =   76
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   296
                     Top             =   3630
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
                     Index           =   75
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   295
                     Top             =   3300
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
                     Index           =   74
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   294
                     Top             =   2970
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
                     Index           =   73
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   293
                     Top             =   2640
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
                     Index           =   72
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   292
                     Top             =   2310
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
                     Index           =   71
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   291
                     Top             =   1980
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
                     Index           =   70
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   290
                     Top             =   1650
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
                     Index           =   69
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   289
                     Top             =   1320
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
                     Index           =   68
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   288
                     Top             =   990
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
                     Index           =   67
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   287
                     Top             =   660
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
                     Index           =   66
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   286
                     Top             =   330
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
                     Index           =   65
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   285
                     Top             =   0
                     Width           =   1130
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "NON DYNA - AFILIASI (PPn)"
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
                     Index           =   49
                     Left            =   6240
                     TabIndex        =   342
                     Top             =   4560
                     Width           =   3480
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   80
                     Left            =   3630
                     Top             =   3300
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   79
                     Left            =   3630
                     Top             =   2970
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   78
                     Left            =   3630
                     Top             =   2640
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
                     Index           =   18
                     Left            =   2400
                     TabIndex        =   336
                     Top             =   3300
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
                     Left            =   2400
                     TabIndex        =   335
                     Top             =   2640
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
                     Index           =   30
                     Left            =   2400
                     TabIndex        =   334
                     Top             =   2970
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
                     Index           =   29
                     Left            =   2400
                     TabIndex        =   333
                     Top             =   330
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
                     Left            =   2400
                     TabIndex        =   332
                     Top             =   0
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
                     Index           =   17
                     Left            =   2400
                     TabIndex        =   331
                     Top             =   660
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
                     Index           =   28
                     Left            =   2400
                     TabIndex        =   330
                     Top             =   1650
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
                     Left            =   2400
                     TabIndex        =   329
                     Top             =   1320
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
                     Left            =   2400
                     TabIndex        =   328
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   77
                     Left            =   3630
                     Top             =   0
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   76
                     Left            =   3630
                     Top             =   330
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   75
                     Left            =   3630
                     Top             =   660
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   74
                     Left            =   3630
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   73
                     Left            =   3630
                     Top             =   1320
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   72
                     Left            =   3630
                     Top             =   1650
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
                     Index           =   27
                     Left            =   2400
                     TabIndex        =   327
                     Top             =   2310
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
                     Left            =   2400
                     TabIndex        =   326
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   71
                     Left            =   3630
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   70
                     Left            =   3630
                     Top             =   2310
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   69
                     Left            =   3630
                     Top             =   3960
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   68
                     Left            =   3630
                     Top             =   3630
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
                     Index           =   16
                     Left            =   2400
                     TabIndex        =   325
                     Top             =   3940
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
                     Index           =   26
                     Left            =   2400
                     TabIndex        =   324
                     Top             =   3630
                     Width           =   60
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "HPP Jasa BP"
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
                     Index           =   43
                     Left            =   0
                     TabIndex        =   323
                     Top             =   3000
                     Width           =   1170
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
                     Left            =   240
                     TabIndex        =   322
                     Top             =   2660
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
                     Index           =   42
                     Left            =   240
                     TabIndex        =   321
                     Top             =   2340
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
                     Index           =   41
                     Left            =   0
                     TabIndex        =   320
                     Top             =   3660
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
                     Index           =   24
                     Left            =   240
                     TabIndex        =   319
                     Top             =   3980
                     Width           =   1230
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Hutang Pemborong"
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
                     Left            =   240
                     TabIndex        =   318
                     Top             =   3315
                     Width           =   1860
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
                     Index           =   40
                     Left            =   240
                     TabIndex        =   317
                     Top             =   1320
                     Width           =   960
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Penj. Jasa BP"
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
                     Index           =   39
                     Left            =   240
                     TabIndex        =   316
                     Top             =   660
                     Width           =   1290
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
                     Index           =   22
                     Left            =   240
                     TabIndex        =   315
                     Top             =   990
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
                     Left            =   240
                     TabIndex        =   314
                     Top             =   2010
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
                     Index           =   38
                     Left            =   240
                     TabIndex        =   313
                     Top             =   1680
                     Width           =   930
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Asuransi"
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
                     Left            =   0
                     TabIndex        =   312
                     Top             =   330
                     Width           =   1650
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Umum"
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
                     Left            =   0
                     TabIndex        =   311
                     Top             =   0
                     Width           =   1410
                  End
               End
               Begin VB.Frame Frame3 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   5175
                  Index           =   4
                  Left            =   -74520
                  TabIndex        =   231
                  Top             =   720
                  Width           =   9975
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   257
                     Top             =   3960
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
                     Index           =   63
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   256
                     Top             =   3630
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
                     Index           =   62
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   255
                     Top             =   3300
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
                     Index           =   61
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   254
                     Top             =   2970
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
                     Index           =   60
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   253
                     Top             =   2640
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
                     Index           =   59
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   252
                     Top             =   2310
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   251
                     Top             =   1980
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   250
                     Top             =   1650
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   249
                     Top             =   1320
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   248
                     Top             =   990
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   247
                     Top             =   660
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   246
                     Top             =   330
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   245
                     Top             =   0
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   244
                     Top             =   3960
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
                     Index           =   63
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   243
                     Top             =   3630
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
                     Index           =   62
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   242
                     Top             =   3300
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
                     Index           =   61
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   241
                     Top             =   2970
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
                     Index           =   60
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   240
                     Top             =   2640
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
                     Index           =   59
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   239
                     Top             =   2310
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   238
                     Top             =   1980
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   237
                     Top             =   1650
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   236
                     Top             =   1320
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   235
                     Top             =   990
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   234
                     Top             =   660
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   233
                     Top             =   330
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   232
                     Top             =   0
                     Width           =   1130
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "NON DYNA - INTERNAL (Pemakaian Sendiri)"
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
                     Index           =   48
                     Left            =   4320
                     TabIndex        =   341
                     Top             =   4560
                     Width           =   5415
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   67
                     Left            =   3630
                     Top             =   3300
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   66
                     Left            =   3630
                     Top             =   2970
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   65
                     Left            =   3630
                     Top             =   2640
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
                     Index           =   15
                     Left            =   2400
                     TabIndex        =   283
                     Top             =   3300
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
                     Left            =   2400
                     TabIndex        =   282
                     Top             =   2640
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
                     Index           =   25
                     Left            =   2400
                     TabIndex        =   281
                     Top             =   2970
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
                     Index           =   24
                     Left            =   2400
                     TabIndex        =   280
                     Top             =   330
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
                     Left            =   2400
                     TabIndex        =   279
                     Top             =   0
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
                     Index           =   14
                     Left            =   2400
                     TabIndex        =   278
                     Top             =   660
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
                     Index           =   23
                     Left            =   2400
                     TabIndex        =   277
                     Top             =   1650
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
                     Left            =   2400
                     TabIndex        =   276
                     Top             =   1320
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
                     Left            =   2400
                     TabIndex        =   275
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   64
                     Left            =   3630
                     Top             =   0
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   63
                     Left            =   3630
                     Top             =   330
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   62
                     Left            =   3630
                     Top             =   660
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   61
                     Left            =   3630
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   60
                     Left            =   3630
                     Top             =   1320
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   59
                     Left            =   3630
                     Top             =   1650
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
                     Index           =   22
                     Left            =   2400
                     TabIndex        =   274
                     Top             =   2310
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
                     Left            =   2400
                     TabIndex        =   273
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   58
                     Left            =   3630
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   57
                     Left            =   3630
                     Top             =   2310
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   56
                     Left            =   3630
                     Top             =   3960
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   55
                     Left            =   3630
                     Top             =   3630
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
                     Index           =   13
                     Left            =   2400
                     TabIndex        =   272
                     Top             =   3940
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
                     Index           =   21
                     Left            =   2400
                     TabIndex        =   271
                     Top             =   3630
                     Width           =   60
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "HPP Jasa BP"
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
                     Left            =   0
                     TabIndex        =   270
                     Top             =   3000
                     Width           =   1170
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
                     Left            =   240
                     TabIndex        =   269
                     Top             =   2660
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
                     Index           =   35
                     Left            =   240
                     TabIndex        =   268
                     Top             =   2340
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
                     Index           =   34
                     Left            =   0
                     TabIndex        =   267
                     Top             =   3660
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
                     Index           =   20
                     Left            =   240
                     TabIndex        =   266
                     Top             =   3980
                     Width           =   1230
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Hutang Pemborong"
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
                     Left            =   240
                     TabIndex        =   265
                     Top             =   3315
                     Width           =   1860
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
                     Index           =   33
                     Left            =   240
                     TabIndex        =   264
                     Top             =   1320
                     Width           =   960
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Penj. Jasa BP"
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
                     Left            =   240
                     TabIndex        =   263
                     Top             =   660
                     Width           =   1290
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
                     Index           =   18
                     Left            =   240
                     TabIndex        =   262
                     Top             =   990
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
                     Left            =   240
                     TabIndex        =   261
                     Top             =   2010
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
                     Index           =   31
                     Left            =   240
                     TabIndex        =   260
                     Top             =   1680
                     Width           =   930
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Asuransi"
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
                     Left            =   0
                     TabIndex        =   259
                     Top             =   330
                     Width           =   1650
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Umum"
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
                     Left            =   0
                     TabIndex        =   258
                     Top             =   0
                     Width           =   1410
                  End
               End
               Begin VB.Frame Frame3 
                  BorderStyle     =   0  'None
                  Caption         =   "Frame3"
                  Height          =   5175
                  Index           =   3
                  Left            =   480
                  TabIndex        =   178
                  Top             =   720
                  Width           =   9975
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   204
                     Top             =   3960
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   203
                     Top             =   3630
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
                     Index           =   49
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   202
                     Top             =   3300
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   201
                     Top             =   2970
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   200
                     Top             =   2640
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   199
                     Top             =   2310
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   198
                     Top             =   1980
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   197
                     Top             =   1650
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   196
                     Top             =   1320
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   195
                     Top             =   990
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   194
                     Top             =   660
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
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   193
                     Top             =   330
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
                     Index           =   39
                     Left            =   3735
                     MaxLength       =   40
                     TabIndex        =   192
                     Top             =   0
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
                     Index           =   51
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   191
                     Top             =   3960
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   190
                     Top             =   3630
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
                     Index           =   49
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   189
                     Top             =   3300
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   188
                     Top             =   2970
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   187
                     Top             =   2640
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   186
                     Top             =   2310
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   185
                     Top             =   1980
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   184
                     Top             =   1650
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   183
                     Top             =   1320
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   182
                     Top             =   990
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   181
                     Top             =   660
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
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   180
                     Top             =   330
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
                     Index           =   39
                     Left            =   2490
                     MaxLength       =   15
                     TabIndex        =   179
                     Top             =   0
                     Width           =   1130
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "NON DYNA - EXTERNAL (Pihak Ketiga)"
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
                     Index           =   47
                     Left            =   4920
                     TabIndex        =   340
                     Top             =   4560
                     Width           =   4770
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   54
                     Left            =   3630
                     Top             =   3300
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   53
                     Left            =   3630
                     Top             =   2970
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   52
                     Left            =   3630
                     Top             =   2640
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
                     Left            =   2400
                     TabIndex        =   230
                     Top             =   3300
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
                     Left            =   2400
                     TabIndex        =   229
                     Top             =   2640
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
                     Index           =   20
                     Left            =   2400
                     TabIndex        =   228
                     Top             =   2970
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
                     Index           =   19
                     Left            =   2400
                     TabIndex        =   227
                     Top             =   330
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
                     Left            =   2400
                     TabIndex        =   226
                     Top             =   0
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
                     Left            =   2400
                     TabIndex        =   225
                     Top             =   660
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
                     Left            =   2400
                     TabIndex        =   224
                     Top             =   1650
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
                     Left            =   2400
                     TabIndex        =   223
                     Top             =   1320
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
                     Left            =   2400
                     TabIndex        =   222
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   51
                     Left            =   3630
                     Top             =   0
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   50
                     Left            =   3630
                     Top             =   330
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   49
                     Left            =   3630
                     Top             =   660
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   48
                     Left            =   3630
                     Top             =   990
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   47
                     Left            =   3630
                     Top             =   1320
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   43
                     Left            =   3630
                     Top             =   1650
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
                     Left            =   2400
                     TabIndex        =   221
                     Top             =   2310
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
                     Left            =   2400
                     TabIndex        =   220
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   42
                     Left            =   3630
                     Top             =   1980
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   41
                     Left            =   3630
                     Top             =   2310
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   40
                     Left            =   3630
                     Top             =   3960
                     Width           =   60
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H0080C0FF&
                     BackStyle       =   1  'Opaque
                     BorderColor     =   &H000080FF&
                     Height          =   315
                     Index           =   39
                     Left            =   3630
                     Top             =   3630
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
                     Left            =   2400
                     TabIndex        =   219
                     Top             =   3940
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
                     Left            =   2400
                     TabIndex        =   218
                     Top             =   3630
                     Width           =   60
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "HPP Jasa BP"
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
                     Left            =   0
                     TabIndex        =   217
                     Top             =   3000
                     Width           =   1170
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
                     Left            =   240
                     TabIndex        =   216
                     Top             =   2660
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
                     Index           =   25
                     Left            =   240
                     TabIndex        =   215
                     Top             =   2340
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
                     Index           =   22
                     Left            =   0
                     TabIndex        =   214
                     Top             =   3660
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
                     Index           =   13
                     Left            =   240
                     TabIndex        =   213
                     Top             =   3980
                     Width           =   1230
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Hutang Pemborong"
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
                     Left            =   240
                     TabIndex        =   212
                     Top             =   3315
                     Width           =   1860
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
                     Index           =   21
                     Left            =   240
                     TabIndex        =   211
                     Top             =   1320
                     Width           =   960
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Penj. Jasa BP"
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
                     Left            =   240
                     TabIndex        =   210
                     Top             =   660
                     Width           =   1290
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
                     Index           =   11
                     Left            =   240
                     TabIndex        =   209
                     Top             =   990
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
                     Left            =   240
                     TabIndex        =   208
                     Top             =   2010
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
                     Index           =   19
                     Left            =   240
                     TabIndex        =   207
                     Top             =   1680
                     Width           =   930
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Asuransi"
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
                     Left            =   0
                     TabIndex        =   206
                     Top             =   330
                     Width           =   1650
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Piutang Umum"
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
                     Left            =   0
                     TabIndex        =   205
                     Top             =   0
                     Width           =   1410
                  End
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
                  Index           =   7
                  X1              =   -75000
                  X2              =   -67680
                  Y1              =   360
                  Y2              =   360
               End
               Begin VB.Line Line4 
                  BorderColor     =   &H000080FF&
                  BorderWidth     =   2
                  Index           =   6
                  X1              =   3600
                  X2              =   11040
                  Y1              =   360
                  Y2              =   360
               End
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
            Begin VB.Line Line4 
               BorderColor     =   &H000080FF&
               BorderWidth     =   2
               Index           =   0
               X1              =   5640
               X2              =   11280
               Y1              =   360
               Y2              =   360
            End
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
            Index           =   80
            Left            =   4485
            MaxLength       =   40
            TabIndex        =   15
            Top             =   2040
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
            Index           =   80
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   14
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
            Index           =   79
            Left            =   4485
            MaxLength       =   40
            TabIndex        =   13
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
            Index           =   79
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   12
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
            Index           =   78
            Left            =   4485
            MaxLength       =   40
            TabIndex        =   11
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
            Index           =   78
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   10
            Top             =   1320
            Width           =   1130
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
            Left            =   720
            TabIndex        =   356
            Top             =   2640
            Width           =   9660
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
            Left            =   720
            TabIndex        =   355
            Top             =   960
            Width           =   9675
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
            Left            =   720
            TabIndex        =   354
            Top             =   3000
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
            Index           =   25
            Left            =   720
            TabIndex        =   353
            Top             =   3360
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
            Index           =   19
            Left            =   3075
            TabIndex        =   352
            Top             =   3720
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
            Left            =   3075
            TabIndex        =   351
            Top             =   3000
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
            Index           =   31
            Left            =   3075
            TabIndex        =   350
            Top             =   3360
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
            Left            =   960
            TabIndex        =   349
            Top             =   3720
            Width           =   1365
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   83
            Left            =   4380
            Top             =   3000
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   82
            Left            =   4380
            Top             =   3360
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   81
            Left            =   4380
            Top             =   3720
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   46
            Left            =   4380
            Top             =   2040
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   45
            Left            =   4380
            Top             =   1680
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   44
            Left            =   4380
            Top             =   1320
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
            Left            =   960
            TabIndex        =   9
            Top             =   2040
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
            Index           =   16
            Left            =   3075
            TabIndex        =   8
            Top             =   1680
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
            Left            =   3075
            TabIndex        =   7
            Top             =   1320
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
            Index           =   8
            Left            =   3075
            TabIndex        =   6
            Top             =   2040
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
            Index           =   15
            Left            =   720
            TabIndex        =   5
            Top             =   1680
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
            Index           =   23
            Left            =   720
            TabIndex        =   4
            Top             =   1320
            Width           =   1755
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000080FF&
            BorderWidth     =   2
            X1              =   -72120
            X2              =   -63480
            Y1              =   480
            Y2              =   480
         End
      End
      Begin ucXPButton.XPButton cmdExit 
         Height          =   495
         Left            =   10320
         TabIndex        =   1
         Top             =   9360
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
         MICON           =   "oMasterAkunAccNew.frx":2CB54
         PICN            =   "oMasterAkunAccNew.frx":2CB70
         PICH            =   "oMasterAkunAccNew.frx":2D044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSave 
         Height          =   495
         Left            =   5160
         TabIndex        =   2
         Top             =   9360
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
         MICON           =   "oMasterAkunAccNew.frx":2D58F
         PICN            =   "oMasterAkunAccNew.frx":2D5AB
         PICH            =   "oMasterAkunAccNew.frx":2DABE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   11640
         Y1              =   9960
         Y2              =   9960
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   11640
         Y1              =   9240
         Y2              =   9240
      End
   End
End
Attribute VB_Name = "oMasterAkunAccNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private rsB As ADODB.Recordset
Private IsNew As Boolean


Private Sub Form_Load()
    Me.Caption = TitleBP
    Move 3240, objA.MyTop
    Me.Width = objA.MyLeft
    Me.Height = fHeight - 260
    Frame1.Left = (((objA.MyLeft - 480) - Val(Frame1.Width)) / 2) + 495
    SSTab1.TabCaption(0) = ""
    SSTab1.TabCaption(1) = ""
    Call ViewMasterAkun
End Sub

Private Sub cmdExit_Click()
    rsMy.Find "NamaForm = 'MASTER AKUN ACC'", , adSearchForward, 1
    If Not rsMy.EOF Then rsMy.Delete adAffectCurrent
    Call zMainForm.ViewMyActiveForm
    If rsMy.RecordCount > 0 Then
        rsMy.MoveLast
        Call ShowMyForm(rsMy.Fields(0).value)
    Else
        zMainForm.PicBG.Visible = False
        zMainForm.FrSc.Visible = False
    End If
    Unload Me
End Sub

Private Sub cmdSave_Click()
    With msg
        .PassingMsg = "Simpan Master Akun ?"
        .PassingTipe = 1
        .Show vbModal
        If .GetMsg = True Then
            For c = 0 To 83
                objA.nID = c
                objA.nTxtAcc = txt(c).Text
            Next c
            
            For c = 84 To 90
                objA.nID = c
                objA.nTxtAcc = ""
            Next c
            
            Call objA.CreateNew_AkunNew(0, IsNew)
            Call objA.CreateNew_AkunNew(1, IsNew)
            Call objA.CreateNew_AkunNew(2, IsNew)
            Call objA.CreateNew_AkunNew(3, IsNew)
            Call objA.CreateNew_AkunNew(4, IsNew)
            Call objA.CreateNew_AkunNew(5, IsNew)
            Call objA.CreateNew_AkunNew(6, IsNew)
        End If
    End With
End Sub


Private Sub ViewMasterAkun()
    Set rsB = objA.Get_AkunNew
    If rsB.RecordCount = 0 Then
        IsNew = True
        Exit Sub
    Else
        IsNew = False
    End If
    Dim nAkun As String
    nAkun = ""
    e = 0
    rsB.MoveFirst
    Do Until rsB.EOF
        For c = 1 To 13
            If (c + (e * 13) - 1) < 84 Then
                txt(c + (e * 13) - 1).Text = rsB.Fields(c).value
                If rsB.Fields(c).value <> "" Then
                    If nAkun = "" Then
                        nAkun = "'" & rsB.Fields(c).value & "'"
                    Else
                        nAkun = nAkun & "," & "'" & rsB.Fields(c).value & "'"
                    End If
                End If
            End If
        Next c
        rsB.MoveNext
        e = e + 1
    Loop

    Dim rsN As ADODB.Recordset
    Set rsN = objA.Get_KodeAkun(" Where kodegl in (" & nAkun & ")")
    For c = 0 To 83
        If txt(c).Text <> "" Then
            rsN.Find "kodegl = '" & Trim(txt(c).Text) & "'", , adSearchForward, 1
            If Not rsN.EOF Then
                txtx(c).Text = rsN.Fields(1).value
            End If
        End If
    Next c
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        ft.ShowForm = "GL"
        ft.Show vbModal
        If ft.IsPass = True Then
            Dim rsGL As ADODB.Recordset
            Set rsGL = ft.MyRs
            txt(Index).Text = rsGL.Fields(0).value
            txtx(Index).Text = rsGL.Fields(1).value
        End If
    End If
End Sub

Private Sub txtx_GotFocus(Index As Integer)
    SendKeys "{tab}"
End Sub
