VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmSIDMasterPelangganNonMerge 
   BackColor       =   &H80000016&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14970
   Icon            =   "frmSIDMasterPelangganNonMerge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSIDMasterPelangganNonMerge.frx":29C12
   ScaleHeight     =   10500
   ScaleWidth      =   14970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   9855
      Left            =   600
      TabIndex        =   214
      Top             =   -240
      Visible         =   0   'False
      Width           =   11415
      Begin VB.Frame frameNotif 
         Height          =   1455
         Left            =   3120
         TabIndex        =   229
         Top             =   3120
         Visible         =   0   'False
         Width           =   6375
         Begin ucXPButton.XPButton cmdNotifProses 
            Height          =   495
            Left            =   4920
            TabIndex        =   230
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&a"
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
            MICON           =   "frmSIDMasterPelangganNonMerge.frx":2BEE2
            PICN            =   "frmSIDMasterPelangganNonMerge.frx":2BEFE
            PICH            =   "frmSIDMasterPelangganNonMerge.frx":3291B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdNotifUbah 
            Height          =   495
            Left            =   3720
            TabIndex        =   231
            Top             =   720
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&a"
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
            MICON           =   "frmSIDMasterPelangganNonMerge.frx":39538
            PICN            =   "frmSIDMasterPelangganNonMerge.frx":39554
            PICH            =   "frmSIDMasterPelangganNonMerge.frx":3FAB6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label10 
            Caption         =   "Notifikasi ID Suggest telah dipilih /  Buat ID Baru"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   232
            Top             =   240
            Width           =   6135
         End
      End
      Begin ARProgBarCtrl.ARProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   228
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         Value           =   0
         ForeColor       =   16744576
         BackColor       =   -2147483626
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6600
         Top             =   5760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtUploadKTP 
         BackColor       =   &H00C0E0FF&
         Height          =   405
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   222
         Top             =   4920
         Width           =   7455
      End
      Begin ucXPButton.XPButton cmdUploadKTP 
         Height          =   495
         Left            =   120
         TabIndex        =   215
         Top             =   4920
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":462D9
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":462F5
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":4D7FD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdProses 
         Height          =   495
         Left            =   9960
         TabIndex        =   216
         Top             =   4920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":557A5
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":557C1
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":5C1DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid GridFind 
         Height          =   2895
         Left            =   120
         TabIndex        =   217
         Top             =   600
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   13
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pilih"
         Columns(0).DataField=   "pilih"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KodeLgn"
         Columns(1).DataField=   "KodeLgn"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No KTP"
         Columns(2).DataField=   "NoKTP"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nama 1"
         Columns(3).DataField=   "NamaCust1"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nama 2"
         Columns(4).DataField=   "NamaCust2"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nama 3"
         Columns(5).DataField=   "NamaCust3"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TypeCustomer"
         Columns(6).DataField=   "TypeCustomer"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "HP1"
         Columns(7).DataField=   "HP1"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "HP2"
         Columns(8).DataField=   "HP2"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Tempat Lahir"
         Columns(9).DataField=   "TempatLahir"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Tgl Lahir"
         Columns(10).DataField=   "TglLahir"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Kode Pos"
         Columns(11).DataField=   "KodePos"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Propinsi"
         Columns(12).DataField=   "Propinsi"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Kota"
         Columns(13).DataField=   "Kota"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Kecamatan"
         Columns(14).DataField=   "Kecamatan"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Kelurahan"
         Columns(15).DataField=   "Kelurahan"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Alamat"
         Columns(16).DataField=   "Alamat"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Score"
         Columns(17).DataField=   "Score"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "AddScore"
         Columns(18).DataField=   "AddScore"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "TotalScore"
         Columns(19).DataField=   "TotalScore"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).DataField=   ""
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).DataField=   ""
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   22
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=22"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=3281"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=3281"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(53)=   "Column(13).Width=3281"
         Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
         Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
         Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
         Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
         Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
         Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
         Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(69)=   "Column(17).Width=106"
         Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(17).AllowSizing=0"
         Splits(0)._ColumnProps(72)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(73)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(74)=   "Column(18).Width=106"
         Splits(0)._ColumnProps(75)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(76)=   "Column(18).AllowSizing=0"
         Splits(0)._ColumnProps(77)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(78)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(79)=   "Column(19).Width=106"
         Splits(0)._ColumnProps(80)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(19).AllowSizing=0"
         Splits(0)._ColumnProps(82)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(83)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(84)=   "Column(20).Width=3281"
         Splits(0)._ColumnProps(85)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(86)=   "Column(20)._WidthInPix=3175"
         Splits(0)._ColumnProps(87)=   "Column(20).Visible=0"
         Splits(0)._ColumnProps(88)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(89)=   "Column(21).Width=3281"
         Splits(0)._ColumnProps(90)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(91)=   "Column(21)._WidthInPix=3175"
         Splits(0)._ColumnProps(92)=   "Column(21).Visible=0"
         Splits(0)._ColumnProps(93)=   "Column(21).Order=22"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
         DeadAreaBackColor=   14933984
         RowDividerColor =   14933984
         RowSubDividerColor=   14933984
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=862,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=859,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=860,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=861,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=870,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=867,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=868,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=869,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=878,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=875,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=876,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=877,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=886,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=883,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=884,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=885,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=894,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=891,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=892,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=893,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=62,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=66,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=70,.parent=13"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=67,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=68,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=69,.parent=17"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=74,.parent=13"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=71,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=72,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=73,.parent=17"
         _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=78,.parent=13"
         _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=75,.parent=14"
         _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=76,.parent=15"
         _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=77,.parent=17"
         _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=82,.parent=13"
         _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=79,.parent=14"
         _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=80,.parent=15"
         _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=81,.parent=17"
         _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=86,.parent=13"
         _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=14"
         _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=15"
         _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=17"
         _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=91,.parent=13"
         _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=88,.parent=14"
         _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=89,.parent=15"
         _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=90,.parent=17"
         _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=95,.parent=13"
         _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=92,.parent=14"
         _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=93,.parent=15"
         _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=94,.parent=17"
         _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=99,.parent=13"
         _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=96,.parent=14"
         _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=97,.parent=15"
         _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=98,.parent=17"
         _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=103,.parent=13"
         _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=100,.parent=14"
         _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=101,.parent=15"
         _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=102,.parent=17"
         _StyleDefs(124) =   "Named:id=33:Normal"
         _StyleDefs(125) =   ":id=33,.parent=0"
         _StyleDefs(126) =   "Named:id=34:Heading"
         _StyleDefs(127) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(128) =   ":id=34,.wraptext=-1"
         _StyleDefs(129) =   "Named:id=35:Footing"
         _StyleDefs(130) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(131) =   "Named:id=36:Selected"
         _StyleDefs(132) =   ":id=36,.parent=33,.bgcolor=&H80FF&,.fgcolor=&H8000000E&"
         _StyleDefs(133) =   "Named:id=37:Caption"
         _StyleDefs(134) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(135) =   "Named:id=38:HighlightRow"
         _StyleDefs(136) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(137) =   "Named:id=39:EvenRow"
         _StyleDefs(138) =   ":id=39,.parent=33"
         _StyleDefs(139) =   "Named:id=40:OddRow"
         _StyleDefs(140) =   ":id=40,.parent=33"
         _StyleDefs(141) =   "Named:id=41:RecordSelector"
         _StyleDefs(142) =   ":id=41,.parent=34"
         _StyleDefs(143) =   "Named:id=42:FilterBar"
         _StyleDefs(144) =   ":id=42,.parent=33"
         _StyleDefs(145) =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(146) =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(147) =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(148) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         _StyleDefs(149) =   "Named:id=87:SudahDitagih"
         _StyleDefs(150) =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(151) =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(152) =   ":id=87,.fontname=Tahoma"
         _StyleDefs(153) =   "Named:id=514:Kuning"
         _StyleDefs(154) =   ":id=514,.parent=87,.bgcolor=&HFFFF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(155) =   ":id=514,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(156) =   ":id=514,.fontname=Tahoma"
         _StyleDefs(157) =   "Named:id=515:Biru"
         _StyleDefs(158) =   ":id=515,.parent=514,.bgcolor=&HFF8080&"
         _StyleDefs(159) =   "Named:id=516:Merah"
         _StyleDefs(160) =   ":id=516,.parent=515,.bgcolor=&H88A0F7&,.fgcolor=&H0&"
         _StyleDefs(161) =   "Named:id=517:Hijau"
         _StyleDefs(162) =   ":id=517,.parent=516,.bgcolor=&HFF00&"
         _StyleDefs(163) =   "Named:id=518:Putih"
         _StyleDefs(164) =   ":id=518,.parent=517,.bgcolor=&HFFFFFF&"
      End
      Begin TrueOleDBGrid70.TDBGrid gridInput 
         Height          =   855
         Left            =   120
         TabIndex        =   219
         Top             =   4080
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   1508
         _LayoutType     =   4
         _RowHeight      =   13
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pilih"
         Columns(0).DataField=   "pilih"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KodeLgn"
         Columns(1).DataField=   "KodeLgn"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No KTP"
         Columns(2).DataField=   "NoKTP"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nama 1"
         Columns(3).DataField=   "NamaCust1"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nama 2"
         Columns(4).DataField=   "NamaCust2"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nama 3"
         Columns(5).DataField=   "NamaCust3"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TypeCustomer"
         Columns(6).DataField=   "TypeCustomer"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "HP1"
         Columns(7).DataField=   "HP1"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "HP2"
         Columns(8).DataField=   "HP2"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Tempat Lahir"
         Columns(9).DataField=   "TempatLahir"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Tgl Lahir"
         Columns(10).DataField=   "TglLahir"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Kode Pos"
         Columns(11).DataField=   "KodePos"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Propinsi"
         Columns(12).DataField=   "Propinsi"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Kota"
         Columns(13).DataField=   "Kota"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Kecamatan"
         Columns(14).DataField=   "Kecamatan"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Kelurahan"
         Columns(15).DataField=   "Kelurahan"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Alamat"
         Columns(16).DataField=   "Alamat"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Score"
         Columns(17).DataField=   "Score"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "AddScore"
         Columns(18).DataField=   "AddScore"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "TotalScore"
         Columns(19).DataField=   "TotalScore"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).DataField=   ""
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).DataField=   ""
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   22
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=22"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=3281"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=3281"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(53)=   "Column(13).Width=3281"
         Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
         Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
         Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
         Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
         Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
         Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
         Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(69)=   "Column(17).Width=3281"
         Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=3175"
         Splits(0)._ColumnProps(72)=   "Column(17).AllowSizing=0"
         Splits(0)._ColumnProps(73)=   "Column(17).Visible=0"
         Splits(0)._ColumnProps(74)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(75)=   "Column(18).Width=106"
         Splits(0)._ColumnProps(76)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(77)=   "Column(18).AllowSizing=0"
         Splits(0)._ColumnProps(78)=   "Column(18).Visible=0"
         Splits(0)._ColumnProps(79)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(80)=   "Column(19).Width=106"
         Splits(0)._ColumnProps(81)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(82)=   "Column(19).AllowSizing=0"
         Splits(0)._ColumnProps(83)=   "Column(19).Visible=0"
         Splits(0)._ColumnProps(84)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(85)=   "Column(20).Width=3281"
         Splits(0)._ColumnProps(86)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(20)._WidthInPix=3175"
         Splits(0)._ColumnProps(88)=   "Column(20).Visible=0"
         Splits(0)._ColumnProps(89)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(90)=   "Column(21).Width=3281"
         Splits(0)._ColumnProps(91)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(92)=   "Column(21)._WidthInPix=3175"
         Splits(0)._ColumnProps(93)=   "Column(21).Order=22"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
         DeadAreaBackColor=   14933984
         RowDividerColor =   14933984
         RowSubDividerColor=   14933984
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=862,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=859,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=860,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=861,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=870,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=867,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=868,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=869,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=878,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=875,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=876,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=877,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=886,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=883,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=884,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=885,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=894,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=891,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=892,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=893,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=62,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=66,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=70,.parent=13"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=67,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=68,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=69,.parent=17"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=74,.parent=13"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=71,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=72,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=73,.parent=17"
         _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=78,.parent=13"
         _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=75,.parent=14"
         _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=76,.parent=15"
         _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=77,.parent=17"
         _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=82,.parent=13"
         _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=79,.parent=14"
         _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=80,.parent=15"
         _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=81,.parent=17"
         _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=86,.parent=13"
         _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=14"
         _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=15"
         _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=17"
         _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=91,.parent=13"
         _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=88,.parent=14"
         _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=89,.parent=15"
         _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=90,.parent=17"
         _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=95,.parent=13"
         _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=92,.parent=14"
         _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=93,.parent=15"
         _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=94,.parent=17"
         _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=99,.parent=13"
         _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=96,.parent=14"
         _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=97,.parent=15"
         _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=98,.parent=17"
         _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=103,.parent=13"
         _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=100,.parent=14"
         _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=101,.parent=15"
         _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=102,.parent=17"
         _StyleDefs(124) =   "Named:id=33:Normal"
         _StyleDefs(125) =   ":id=33,.parent=0"
         _StyleDefs(126) =   "Named:id=34:Heading"
         _StyleDefs(127) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(128) =   ":id=34,.wraptext=-1"
         _StyleDefs(129) =   "Named:id=35:Footing"
         _StyleDefs(130) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(131) =   "Named:id=36:Selected"
         _StyleDefs(132) =   ":id=36,.parent=33,.bgcolor=&H80FF&,.fgcolor=&H8000000E&"
         _StyleDefs(133) =   "Named:id=37:Caption"
         _StyleDefs(134) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(135) =   "Named:id=38:HighlightRow"
         _StyleDefs(136) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(137) =   "Named:id=39:EvenRow"
         _StyleDefs(138) =   ":id=39,.parent=33"
         _StyleDefs(139) =   "Named:id=40:OddRow"
         _StyleDefs(140) =   ":id=40,.parent=33"
         _StyleDefs(141) =   "Named:id=41:RecordSelector"
         _StyleDefs(142) =   ":id=41,.parent=34"
         _StyleDefs(143) =   "Named:id=42:FilterBar"
         _StyleDefs(144) =   ":id=42,.parent=33"
         _StyleDefs(145) =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(146) =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(147) =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(148) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         _StyleDefs(149) =   "Named:id=87:SudahDitagih"
         _StyleDefs(150) =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(151) =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(152) =   ":id=87,.fontname=Tahoma"
         _StyleDefs(153) =   "Named:id=514:Kuning"
         _StyleDefs(154) =   ":id=514,.parent=87,.bgcolor=&HFFFF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(155) =   ":id=514,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(156) =   ":id=514,.fontname=Tahoma"
         _StyleDefs(157) =   "Named:id=515:Biru"
         _StyleDefs(158) =   ":id=515,.parent=514,.bgcolor=&HFF8080&"
         _StyleDefs(159) =   "Named:id=516:Merah"
         _StyleDefs(160) =   ":id=516,.parent=515,.bgcolor=&H88A0F7&,.fgcolor=&H0&"
         _StyleDefs(161) =   "Named:id=517:Hijau"
         _StyleDefs(162) =   ":id=517,.parent=516,.bgcolor=&HFF00&"
         _StyleDefs(163) =   "Named:id=518:Putih"
         _StyleDefs(164) =   ":id=518,.parent=517,.bgcolor=&HFFFFFF&"
      End
      Begin ucXPButton.XPButton cmdClear 
         Height          =   495
         Left            =   9960
         TabIndex        =   221
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":62DFB
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":62E17
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":696AF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton XPButton1 
         Height          =   495
         Left            =   8760
         TabIndex        =   223
         Top             =   120
         Visible         =   0   'False
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":70404
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":70420
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":708F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblLoading 
         Height          =   255
         Left            =   4560
         TabIndex        =   226
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label LblJumlah 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data :"
         Height          =   195
         Left            =   120
         TabIndex        =   220
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: MASTER PELANGGAN ::."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   4080
         TabIndex        =   218
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   3735
         Left            =   120
         Stretch         =   -1  'True
         Top             =   5520
         Width           =   6015
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   3600
         Width           =   11280
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   480
      TabIndex        =   8
      Top             =   0
      Width           =   13695
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         ForeColor       =   &H80000008&
         Height          =   8850
         Left            =   120
         ScaleHeight     =   8820
         ScaleWidth      =   11370
         TabIndex        =   10
         Top             =   375
         Width           =   11400
         Begin VB.PictureBox PicMove 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   20000
            Left            =   15
            ScaleHeight     =   19995
            ScaleWidth      =   11295
            TabIndex        =   11
            Top             =   105
            Width           =   11295
            Begin VB.TextBox txtUploadKTPSusulan 
               BackColor       =   &H00C0E0FF&
               Height          =   405
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   224
               Top             =   4680
               Visible         =   0   'False
               Width           =   7455
            End
            Begin ucXPButton.XPButton cmdUploadSusulan 
               Height          =   495
               Left            =   0
               TabIndex        =   225
               Top             =   4680
               Visible         =   0   'False
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   873
               BTYPE           =   6
               TX              =   "&a"
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
               MICON           =   "frmSIDMasterPelangganNonMerge.frx":70E3F
               PICN            =   "frmSIDMasterPelangganNonMerge.frx":70E5B
               PICH            =   "frmSIDMasterPelangganNonMerge.frx":78363
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Frame FrDataPemakai 
               BackColor       =   &H80000016&
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   4440
               Left            =   20
               TabIndex        =   121
               Top             =   12120
               Width           =   11190
               Begin VB.Frame FrAlmS 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000016&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1335
                  Left            =   4800
                  TabIndex        =   179
                  Top             =   1515
                  Visible         =   0   'False
                  Width           =   2730
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
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
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   183
                     Top             =   1005
                     Width           =   2715
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
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
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   182
                     Top             =   675
                     Width           =   2715
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
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
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   181
                     Top             =   345
                     Width           =   2715
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
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
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   180
                     Top             =   15
                     Width           =   2715
                  End
               End
               Begin VB.ComboBox cboOlahS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   61
                  Top             =   3255
                  Width           =   2205
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   74
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   63
                  Top             =   3930
                  Width           =   2190
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   73
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   62
                  Top             =   3600
                  Width           =   2190
               End
               Begin VB.CheckBox ChSurat 
                  BackColor       =   &H80000016&
                  Caption         =   "Data Pemakai sama dengan Data Pelanggan"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00202020&
                  Height          =   255
                  Left            =   75
                  TabIndex        =   49
                  Top             =   240
                  Width           =   5655
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   49
                  Left            =   7875
                  MaxLength       =   25
                  TabIndex        =   51
                  Top             =   555
                  Width           =   3075
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   48
                  Left            =   1755
                  MaxLength       =   50
                  TabIndex        =   50
                  Top             =   555
                  Width           =   3510
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   52
                  Left            =   7875
                  MaxLength       =   150
                  TabIndex        =   55
                  Top             =   885
                  Width           =   3075
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   50
                  Left            =   2700
                  MaxLength       =   30
                  TabIndex        =   53
                  Top             =   885
                  Width           =   2560
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   53
                  Left            =   1230
                  MaxLength       =   50
                  TabIndex        =   59
                  Top             =   2565
                  Width           =   2190
               End
               Begin VB.TextBox txt 
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
                  Height          =   810
                  IMEMode         =   3  'DISABLE
                  Index           =   54
                  Left            =   4800
                  MaxLength       =   150
                  MultiLine       =   -1  'True
                  TabIndex        =   68
                  Top             =   2850
                  Width           =   2985
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   51
                  Left            =   5280
                  MaxLength       =   30
                  TabIndex        =   54
                  Top             =   885
                  Width           =   2580
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   55
                  Left            =   4800
                  MaxLength       =   3
                  TabIndex        =   69
                  Top             =   3675
                  Width           =   570
               End
               Begin VB.ComboBox cboTitleS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1755
                  Style           =   2  'Dropdown List
                  TabIndex        =   52
                  Top             =   885
                  Width           =   930
               End
               Begin VB.ComboBox cboAgamaS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   57
                  Top             =   1905
                  Width           =   2205
               End
               Begin VB.ComboBox cboPekS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   58
                  Top             =   2235
                  Width           =   2205
               End
               Begin VB.ComboBox cboHobiS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   60
                  Top             =   2910
                  Width           =   2205
               End
               Begin VB.ComboBox cboProvS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   64
                  Top             =   1575
                  Width           =   2700
               End
               Begin VB.ComboBox cboKotaS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   65
                  Top             =   1890
                  Width           =   2700
               End
               Begin VB.ComboBox cboKecS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   66
                  Top             =   2205
                  Width           =   2700
               End
               Begin VB.ComboBox cboKelS 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   67
                  Top             =   2520
                  Width           =   2700
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
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
                  Height          =   320
                  Index           =   56
                  Left            =   5580
                  MaxLength       =   3
                  TabIndex        =   70
                  Top             =   3675
                  Width           =   570
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
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
                  Height          =   320
                  Index           =   57
                  Left            =   4800
                  MaxLength       =   10
                  TabIndex        =   71
                  Top             =   4005
                  Width           =   1350
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
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
                  Height          =   320
                  Index           =   58
                  Left            =   8760
                  MaxLength       =   4
                  TabIndex        =   72
                  Top             =   1575
                  Width           =   645
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   61
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   75
                  Top             =   2235
                  Width           =   2085
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   59
                  Left            =   9420
                  MaxLength       =   10
                  TabIndex        =   73
                  Top             =   1575
                  Width           =   1425
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   60
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   74
                  Top             =   1905
                  Width           =   2085
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   62
                  Left            =   8760
                  MaxLength       =   100
                  TabIndex        =   76
                  Top             =   2565
                  Width           =   2085
               End
               Begin TDBDate6Ctl.TDBDate DtLahirS 
                  Height          =   315
                  Left            =   1230
                  TabIndex        =   56
                  Top             =   1575
                  Width           =   1530
                  _Version        =   65536
                  _ExtentX        =   2699
                  _ExtentY        =   556
                  Calendar        =   "frmSIDMasterPelangganNonMerge.frx":8030B
                  Caption         =   "frmSIDMasterPelangganNonMerge.frx":80437
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmSIDMasterPelangganNonMerge.frx":80493
                  Keys            =   "frmSIDMasterPelangganNonMerge.frx":804B1
                  Spin            =   "frmSIDMasterPelangganNonMerge.frx":8050F
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   12640511
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  CursorPosition  =   0
                  DataProperty    =   0
                  DisplayFormat   =   "dd/mm/yyyy"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  FirstMonth      =   4
                  ForeColor       =   0
                  Format          =   "dd/mm/yyyy"
                  HighlightText   =   0
                  IMEMode         =   3
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxDate         =   2958465
                  MinDate         =   2
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  PromptChar      =   "_"
                  ReadOnly        =   0
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "__/__/____"
                  ValidateMode    =   0
                  ValueVT         =   6815745
                  Value           =   39183
                  CenturyMode     =   0
               End
               Begin ucXPButton.XPButton cmdBackS 
                  Height          =   345
                  Left            =   7530
                  TabIndex        =   184
                  Top             =   1530
                  Visible         =   0   'False
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  BTYPE           =   3
                  TX              =   ""
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
                  MICON           =   "frmSIDMasterPelangganNonMerge.frx":80537
                  PICN            =   "frmSIDMasterPelangganNonMerge.frx":80553
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
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
                  Left            =   1125
                  TabIndex        =   178
                  Top             =   3990
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Preferensi"
                  BeginProperty Font 
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
                  Left            =   75
                  TabIndex        =   177
                  Top             =   3990
                  Width           =   1020
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
                  Left            =   1125
                  TabIndex        =   176
                  Top             =   3660
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
                  Left            =   1125
                  TabIndex        =   175
                  Top             =   3315
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suku"
                  BeginProperty Font 
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
                  Left            =   75
                  TabIndex        =   174
                  Top             =   3660
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Olahraga"
                  BeginProperty Font 
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
                  Left            =   75
                  TabIndex        =   173
                  Top             =   3315
                  Width           =   885
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
                  Left            =   1635
                  TabIndex        =   163
                  Top             =   570
                  Width           =   60
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No. KTP"
                  BeginProperty Font 
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
                  Left            =   75
                  TabIndex        =   162
                  Top             =   570
                  Width           =   735
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   " Data Pemakai"
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
                  Left            =   0
                  TabIndex        =   161
                  Top             =   0
                  Width           =   1410
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
                  Left            =   7750
                  TabIndex        =   160
                  Top             =   570
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
                  Left            =   1635
                  TabIndex        =   159
                  Top             =   900
                  Width           =   60
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama Customer"
                  BeginProperty Font 
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
                  Left            =   75
                  TabIndex        =   158
                  Top             =   900
                  Width           =   1545
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "NPWP"
                  BeginProperty Font 
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
                  Left            =   7020
                  TabIndex        =   157
                  Top             =   570
                  Width           =   555
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tgl Lahir"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   56
                  Left            =   75
                  TabIndex        =   156
                  Top             =   1605
                  Width           =   855
               End
               Begin VB.Label Label13 
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
                  Index           =   42
                  Left            =   1125
                  TabIndex        =   155
                  Top             =   1590
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Agama"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   55
                  Left            =   75
                  TabIndex        =   154
                  Top             =   1905
                  Width           =   675
               End
               Begin VB.Label Label13 
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
                  Index           =   41
                  Left            =   1125
                  TabIndex        =   153
                  Top             =   1905
                  Width           =   60
               End
               Begin VB.Line Line7 
                  BorderColor     =   &H000080FF&
                  BorderWidth     =   2
                  X1              =   1635
                  X2              =   10950
                  Y1              =   120
                  Y2              =   120
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
                  Left            =   1125
                  TabIndex        =   152
                  Top             =   2910
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hobi"
                  BeginProperty Font 
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
                  Left            =   75
                  TabIndex        =   151
                  Top             =   2925
                  Width           =   435
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
                  Left            =   1125
                  TabIndex        =   150
                  Top             =   2250
                  Width           =   60
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pekerjaan"
                  BeginProperty Font 
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
                  Left            =   75
                  TabIndex        =   149
                  Top             =   2250
                  Width           =   1005
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Data Diri"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   225
                  Index           =   54
                  Left            =   0
                  TabIndex        =   148
                  Top             =   1275
                  Width           =   3480
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Alamat"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   53
                  Left            =   3495
                  TabIndex        =   147
                  Top             =   1275
                  Width           =   4365
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Telepon"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   52
                  Left            =   7875
                  TabIndex        =   146
                  Top             =   1275
                  Width           =   3075
               End
               Begin VB.Shape Shape9 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   0
                  Top             =   1500
                  Width           =   3480
               End
               Begin VB.Shape Shape8 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   3495
                  Top             =   1500
                  Width           =   4365
               End
               Begin VB.Label Label13 
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
                  Index           =   40
                  Left            =   4700
                  TabIndex        =   145
                  Top             =   1605
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Provinsi"
                  BeginProperty Font 
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
                  Left            =   3550
                  TabIndex        =   144
                  Top             =   1620
                  Width           =   795
               End
               Begin VB.Label Label13 
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
                  Index           =   39
                  Left            =   4700
                  TabIndex        =   143
                  Top             =   1920
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kab/Kota"
                  BeginProperty Font 
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
                  Left            =   3550
                  TabIndex        =   142
                  Top             =   1935
                  Width           =   915
               End
               Begin VB.Label Label13 
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
                  Index           =   38
                  Left            =   4700
                  TabIndex        =   141
                  Top             =   2235
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kecamatan"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   49
                  Left            =   3550
                  TabIndex        =   140
                  Top             =   2235
                  Width           =   1080
               End
               Begin VB.Label Label13 
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
                  Index           =   37
                  Left            =   4700
                  TabIndex        =   139
                  Top             =   2565
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kelurahan"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   48
                  Left            =   3550
                  TabIndex        =   138
                  Top             =   2565
                  Width           =   990
               End
               Begin VB.Label Label13 
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
                  Index           =   36
                  Left            =   4700
                  TabIndex        =   137
                  Top             =   2835
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jalan /"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   47
                  Left            =   3550
                  TabIndex        =   136
                  Top             =   2835
                  Width           =   690
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kompleks;"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   46
                  Left            =   3550
                  TabIndex        =   135
                  Top             =   3075
                  Width           =   1005
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   45
                  Left            =   3550
                  TabIndex        =   134
                  Top             =   3285
                  Width           =   255
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":           /"
                  BeginProperty Font 
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
                  Left            =   4700
                  TabIndex        =   133
                  Top             =   3690
                  Width           =   840
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "RT/RW"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Index           =   44
                  Left            =   3550
                  TabIndex        =   132
                  Top             =   3690
                  Width           =   660
               End
               Begin VB.Label Label13 
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
                  Index           =   34
                  Left            =   4695
                  TabIndex        =   131
                  Top             =   4035
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kode Pos"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   130
                  Top             =   4035
                  Width           =   885
               End
               Begin VB.Shape Shape7 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   7875
                  Top             =   1500
                  Width           =   3075
               End
               Begin VB.Label Label13 
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
                  Index           =   33
                  Left            =   8650
                  TabIndex        =   129
                  Top             =   1620
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "PSTN"
                  BeginProperty Font 
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
                  Left            =   7940
                  TabIndex        =   128
                  Top             =   1620
                  Width           =   495
               End
               Begin VB.Label Label13 
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
                  Index           =   32
                  Left            =   8650
                  TabIndex        =   127
                  Top             =   1935
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "HP 1"
                  BeginProperty Font 
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
                  Left            =   7940
                  TabIndex        =   126
                  Top             =   1935
                  Width           =   435
               End
               Begin VB.Label Label13 
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
                  Left            =   8650
                  TabIndex        =   125
                  Top             =   2265
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "HP 2"
                  BeginProperty Font 
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
                  Left            =   7940
                  TabIndex        =   124
                  Top             =   2265
                  Width           =   435
               End
               Begin VB.Label Label13 
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
                  Left            =   8650
                  TabIndex        =   123
                  Top             =   2595
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-mail"
                  BeginProperty Font 
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
                  Left            =   7940
                  TabIndex        =   122
                  Top             =   2595
                  Width           =   630
               End
            End
            Begin VB.Frame FrDtPlgan 
               BackColor       =   &H80000016&
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   4860
               Left            =   30
               TabIndex        =   12
               Top             =   -45
               Width           =   11190
               Begin MSMask.MaskEdBox txtNpwp 
                  Height          =   330
                  Left            =   8085
                  TabIndex        =   213
                  Top             =   810
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   582
                  _Version        =   393216
                  MaxLength       =   21
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "###.###.###.#-###.###"
                  PromptChar      =   "_"
               End
               Begin VB.ComboBox cboKewarganegaraan 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  ItemData        =   "frmSIDMasterPelangganNonMerge.frx":8075E
                  Left            =   1965
                  List            =   "frmSIDMasterPelangganNonMerge.frx":80760
                  Style           =   2  'Dropdown List
                  TabIndex        =   205
                  Top             =   495
                  Width           =   1125
               End
               Begin MSMask.MaskEdBox MaskEdBox9 
                  Height          =   270
                  Left            =   6180
                  TabIndex        =   204
                  Top             =   1845
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin VB.TextBox txt 
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
                  Height          =   320
                  IMEMode         =   3  'DISABLE
                  Index           =   23
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   23
                  Top             =   2880
                  Width           =   2075
               End
               Begin MSMask.MaskEdBox MaskEdBox8 
                  Height          =   270
                  Left            =   3345
                  TabIndex        =   203
                  Top             =   3600
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
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
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   25
                  Top             =   3570
                  Width           =   2075
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
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   24
                  Top             =   3225
                  Width           =   2075
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
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   21
                  Top             =   2145
                  Width           =   2075
               End
               Begin VB.Frame FrAlm 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000016&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   1335
                  Left            =   4800
                  TabIndex        =   165
                  Top             =   2145
                  Visible         =   0   'False
                  Width           =   3045
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
                     BackColor       =   &H0080FFFF&
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
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   36
                     Top             =   1005
                     Width           =   2950
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
                     BackColor       =   &H0080FFFF&
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
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   35
                     Top             =   675
                     Width           =   2950
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
                     BackColor       =   &H0080FFFF&
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
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   34
                     Top             =   345
                     Width           =   2950
                  End
                  Begin VB.TextBox txt 
                     Appearance      =   0  'Flat
                     BackColor       =   &H0080FFFF&
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
                     Left            =   0
                     MaxLength       =   100
                     TabIndex        =   33
                     Top             =   15
                     Width           =   2950
                  End
                  Begin MSMask.MaskEdBox MaskEdBox3 
                     Height          =   270
                     Left            =   2970
                     TabIndex        =   196
                     Top             =   30
                     Width           =   60
                     _ExtentX        =   106
                     _ExtentY        =   476
                     _Version        =   393216
                     Appearance      =   0
                     Enabled         =   0   'False
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox MaskEdBox1 
                     Height          =   270
                     Left            =   2970
                     TabIndex        =   197
                     Top             =   360
                     Width           =   60
                     _ExtentX        =   106
                     _ExtentY        =   476
                     _Version        =   393216
                     Appearance      =   0
                     Enabled         =   0   'False
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox MaskEdBox2 
                     Height          =   270
                     Left            =   2970
                     TabIndex        =   198
                     Top             =   705
                     Width           =   60
                     _ExtentX        =   106
                     _ExtentY        =   476
                     _Version        =   393216
                     Appearance      =   0
                     Enabled         =   0   'False
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox MaskEdBox4 
                     Height          =   270
                     Left            =   2970
                     TabIndex        =   199
                     Top             =   1035
                     Width           =   60
                     _ExtentX        =   106
                     _ExtentY        =   476
                     _Version        =   393216
                     Appearance      =   0
                     Enabled         =   0   'False
                     PromptChar      =   "_"
                  End
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
                  Left            =   8760
                  MaxLength       =   15
                  TabIndex        =   193
                  Top             =   4110
                  Width           =   2085
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
                  Left            =   8760
                  MaxLength       =   30
                  TabIndex        =   190
                  Top             =   3765
                  Width           =   2085
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
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   186
                  Top             =   3150
                  Width           =   2085
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
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   27
                  Top             =   4245
                  Width           =   2075
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
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   26
                  Top             =   3915
                  Width           =   2075
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   795
                  IMEMode         =   3  'DISABLE
                  Index           =   8
                  Left            =   4800
                  MaxLength       =   150
                  MultiLine       =   -1  'True
                  TabIndex        =   41
                  Top             =   3480
                  Width           =   2985
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
                  Left            =   4800
                  MaxLength       =   3
                  TabIndex        =   42
                  Top             =   4290
                  Width           =   570
               End
               Begin VB.ComboBox cboKec 
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   2190
                  Width           =   2700
               End
               Begin VB.ComboBox cboKel 
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   2505
                  Width           =   2700
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
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
                  Index           =   10
                  Left            =   5580
                  MaxLength       =   3
                  TabIndex        =   43
                  Top             =   4290
                  Width           =   570
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Index           =   11
                  Left            =   4800
                  MaxLength       =   6
                  TabIndex        =   32
                  Top             =   1830
                  Width           =   1350
               End
               Begin VB.ComboBox cboProv 
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   1875
                  Visible         =   0   'False
                  Width           =   2700
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   1245
                  MaxLength       =   50
                  TabIndex        =   22
                  Top             =   2535
                  Width           =   2075
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   8085
                  MaxLength       =   150
                  TabIndex        =   19
                  Top             =   1170
                  Width           =   3075
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   5490
                  MaxLength       =   30
                  TabIndex        =   18
                  Top             =   1170
                  Width           =   2580
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   2910
                  MaxLength       =   30
                  TabIndex        =   17
                  Top             =   1170
                  Width           =   2560
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
                  Left            =   8085
                  MaxLength       =   20
                  TabIndex        =   15
                  Top             =   825
                  Width           =   3075
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   1965
                  MaxLength       =   16
                  TabIndex        =   14
                  Top             =   825
                  Width           =   3510
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
                  Left            =   8760
                  MaxLength       =   100
                  TabIndex        =   48
                  Top             =   2820
                  Width           =   2085
               End
               Begin VB.TextBox txt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
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
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   46
                  Top             =   2160
                  Width           =   2085
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
                  Left            =   9420
                  MaxLength       =   10
                  TabIndex        =   45
                  Top             =   1830
                  Width           =   1425
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
                  Left            =   8760
                  MaxLength       =   20
                  TabIndex        =   47
                  Top             =   2490
                  Width           =   2085
               End
               Begin VB.TextBox txt 
                  Alignment       =   2  'Center
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
                  Index           =   12
                  Left            =   8760
                  MaxLength       =   4
                  TabIndex        =   44
                  Top             =   1830
                  Width           =   645
               End
               Begin VB.ComboBox cboTitle 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1965
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   1170
                  Width           =   930
               End
               Begin VB.ComboBox cboTipe 
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1965
                  Style           =   2  'Dropdown List
                  TabIndex        =   13
                  Top             =   165
                  Width           =   2490
               End
               Begin TDBDate6Ctl.TDBDate DtLahir 
                  Height          =   315
                  Left            =   1245
                  TabIndex        =   20
                  Top             =   1815
                  Width           =   1530
                  _Version        =   65536
                  _ExtentX        =   2699
                  _ExtentY        =   556
                  Calendar        =   "frmSIDMasterPelangganNonMerge.frx":80762
                  Caption         =   "frmSIDMasterPelangganNonMerge.frx":8088E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "frmSIDMasterPelangganNonMerge.frx":808EA
                  Keys            =   "frmSIDMasterPelangganNonMerge.frx":80908
                  Spin            =   "frmSIDMasterPelangganNonMerge.frx":80966
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  CursorPosition  =   0
                  DataProperty    =   0
                  DisplayFormat   =   "dd/mm/yyyy"
                  EditMode        =   0
                  Enabled         =   -1
                  ErrorBeep       =   0
                  FirstMonth      =   4
                  ForeColor       =   0
                  Format          =   "dd/mm/yyyy"
                  HighlightText   =   0
                  IMEMode         =   3
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxDate         =   2958465
                  MinDate         =   2
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  PromptChar      =   "_"
                  ReadOnly        =   0
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "__/__/____"
                  ValidateMode    =   0
                  ValueVT         =   6815745
                  Value           =   39183
                  CenturyMode     =   0
               End
               Begin ucXPButton.XPButton cmdBack 
                  Height          =   345
                  Left            =   7455
                  TabIndex        =   166
                  Top             =   3735
                  Visible         =   0   'False
                  Width           =   300
                  _ExtentX        =   529
                  _ExtentY        =   609
                  BTYPE           =   3
                  TX              =   ""
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
                  MICON           =   "frmSIDMasterPelangganNonMerge.frx":8098E
                  PICN            =   "frmSIDMasterPelangganNonMerge.frx":809AA
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSMask.MaskEdBox MaskEdBox5 
                  Height          =   270
                  Left            =   3360
                  TabIndex        =   200
                  Top             =   2160
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MaskEdBox6 
                  Height          =   270
                  Left            =   3360
                  TabIndex        =   201
                  Top             =   2565
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MaskEdBox7 
                  Height          =   270
                  Left            =   3345
                  TabIndex        =   202
                  Top             =   3255
                  Width           =   60
                  _ExtentX        =   106
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin VB.ComboBox cboAgama 
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
                  Height          =   315
                  Left            =   1245
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   2130
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.ComboBox cboPek 
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1245
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   2220
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.ComboBox cboHobi 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1200
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   2250
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.ComboBox cboOlah 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1230
                  Style           =   2  'Dropdown List
                  TabIndex        =   31
                  Top             =   2235
                  Visible         =   0   'False
                  Width           =   2205
               End
               Begin VB.ComboBox cboKota 
                  BackColor       =   &H0080FFFF&
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   4800
                  Style           =   2  'Dropdown List
                  TabIndex        =   38
                  Top             =   1875
                  Visible         =   0   'False
                  Width           =   2700
               End
               Begin VB.Frame Frame2 
                  BackColor       =   &H80000016&
                  Height          =   540
                  Left            =   7245
                  TabIndex        =   209
                  Top             =   225
                  Width           =   3885
                  Begin VB.TextBox txt 
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
                     Height          =   320
                     Index           =   1
                     Left            =   2040
                     MaxLength       =   12
                     TabIndex        =   210
                     Top             =   135
                     Width           =   1740
                  End
                  Begin VB.Label Label23 
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
                     Left            =   1935
                     TabIndex        =   212
                     Top             =   195
                     Width           =   60
                  End
                  Begin VB.Label Label24 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Kode Customer"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   120
                     TabIndex        =   211
                     Top             =   195
                     Width           =   1470
                  End
               End
               Begin VB.Label Label2 
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
                  Left            =   1860
                  TabIndex        =   208
                  Top             =   540
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
                  Left            =   1650
                  TabIndex        =   207
                  Top             =   900
                  Width           =   60
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kewarganegaraan"
                  BeginProperty Font 
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
                  Left            =   45
                  TabIndex        =   206
                  Top             =   555
                  Width           =   1770
               End
               Begin VB.Label Label13 
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
                  Left            =   8640
                  TabIndex        =   195
                  Top             =   4170
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Telp"
                  BeginProperty Font 
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
                  Left            =   7935
                  TabIndex        =   194
                  Top             =   4140
                  Width           =   420
               End
               Begin VB.Label Label13 
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
                  Left            =   8640
                  TabIndex        =   192
                  Top             =   3795
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama"
                  BeginProperty Font 
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
                  Left            =   7935
                  TabIndex        =   191
                  Top             =   3795
                  Width           =   555
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Data PIC"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   4
                  Left            =   7860
                  TabIndex        =   189
                  Top             =   3495
                  Width           =   3075
               End
               Begin VB.Label Label13 
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
                  Left            =   8640
                  TabIndex        =   188
                  Top             =   3180
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No Fax"
                  BeginProperty Font 
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
                  Left            =   7935
                  TabIndex        =   187
                  Top             =   3180
                  Width           =   675
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
                  Left            =   1125
                  TabIndex        =   172
                  Top             =   4290
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Preferensi"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   171
                  Top             =   4305
                  Width           =   1020
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
                  Left            =   1125
                  TabIndex        =   170
                  Top             =   3975
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
                  Left            =   1125
                  TabIndex        =   169
                  Top             =   3630
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Suku"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   168
                  Top             =   3975
                  Width           =   480
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Olahraga"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   167
                  Top             =   3630
                  Width           =   885
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "E-mail"
                  BeginProperty Font 
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
                  Left            =   7935
                  TabIndex        =   120
                  Top             =   2850
                  Width           =   630
               End
               Begin VB.Label Label13 
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
                  Left            =   8655
                  TabIndex        =   119
                  Top             =   2850
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "HP 2"
                  BeginProperty Font 
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
                  Left            =   7935
                  TabIndex        =   118
                  Top             =   2520
                  Width           =   435
               End
               Begin VB.Label Label13 
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
                  Left            =   8655
                  TabIndex        =   117
                  Top             =   2520
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "HP 1"
                  BeginProperty Font 
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
                  Left            =   7935
                  TabIndex        =   116
                  Top             =   2190
                  Width           =   435
               End
               Begin VB.Label Label13 
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
                  Left            =   8655
                  TabIndex        =   115
                  Top             =   2190
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "PSTN"
                  BeginProperty Font 
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
                  Left            =   7935
                  TabIndex        =   114
                  Top             =   1875
                  Width           =   495
               End
               Begin VB.Label Label13 
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
                  Left            =   8655
                  TabIndex        =   113
                  Top             =   1875
                  Width           =   60
               End
               Begin VB.Shape Shape4 
                  BackColor       =   &H80000016&
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   7875
                  Top             =   1785
                  Width           =   3075
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kode Pos"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   112
                  Top             =   1875
                  Width           =   885
               End
               Begin VB.Label Label13 
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
                  Left            =   4680
                  TabIndex        =   111
                  Top             =   1920
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "RT/RW"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   110
                  Top             =   4320
                  Width           =   660
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":           /"
                  BeginProperty Font 
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
                  Left            =   4695
                  TabIndex        =   109
                  Top             =   4320
                  Width           =   840
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "No"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   108
                  Top             =   3915
                  Width           =   255
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kompleks;"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   107
                  Top             =   3705
                  Width           =   1005
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Jalan /"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   106
                  Top             =   3465
                  Width           =   690
               End
               Begin VB.Label Label13 
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
                  Left            =   4695
                  TabIndex        =   105
                  Top             =   3465
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kelurahan"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   104
                  Top             =   3180
                  Width           =   990
               End
               Begin VB.Label Label13 
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
                  Left            =   4695
                  TabIndex        =   103
                  Top             =   3180
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kecamatan"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   102
                  Top             =   2850
                  Width           =   1080
               End
               Begin VB.Label Label13 
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
                  Left            =   4695
                  TabIndex        =   101
                  Top             =   2880
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Kab/Kota"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   100
                  Top             =   2550
                  Width           =   915
               End
               Begin VB.Label Label13 
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
                  Left            =   4695
                  TabIndex        =   99
                  Top             =   2535
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Provinsi"
                  BeginProperty Font 
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
                  Left            =   3555
                  TabIndex        =   98
                  Top             =   2235
                  Width           =   795
               End
               Begin VB.Label Label13 
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
                  Left            =   4695
                  TabIndex        =   97
                  Top             =   2220
                  Width           =   60
               End
               Begin VB.Shape Shape3 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   3495
                  Top             =   1785
                  Width           =   4365
               End
               Begin VB.Shape Shape12 
                  BorderColor     =   &H000080FF&
                  Height          =   2910
                  Left            =   0
                  Top             =   1785
                  Width           =   3480
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Telepon"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   12
                  Left            =   7875
                  TabIndex        =   96
                  Top             =   1515
                  Width           =   3075
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Alamat"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000B&
                  Height          =   225
                  Index           =   11
                  Left            =   3495
                  TabIndex        =   95
                  Top             =   1515
                  Width           =   4365
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  Caption         =   "Data Diri"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   225
                  Index           =   0
                  Left            =   0
                  TabIndex        =   94
                  Top             =   1515
                  Width           =   3480
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Pekerjaan"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   93
                  Top             =   2565
                  Width           =   1005
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
                  Left            =   1125
                  TabIndex        =   92
                  Top             =   2565
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Hobi"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   91
                  Top             =   3285
                  Width           =   435
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
                  Left            =   1125
                  TabIndex        =   90
                  Top             =   3285
                  Width           =   60
               End
               Begin VB.Line Line1 
                  BorderColor     =   &H000080FF&
                  BorderWidth     =   2
                  X1              =   1635
                  X2              =   11190
                  Y1              =   120
                  Y2              =   120
               End
               Begin VB.Label Label13 
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
                  Left            =   1125
                  TabIndex        =   89
                  Top             =   2205
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Agama"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   88
                  Top             =   2205
                  Width           =   675
               End
               Begin VB.Label Label13 
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
                  Left            =   1125
                  TabIndex        =   87
                  Top             =   1845
                  Width           =   60
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tgl Lahir"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   86
                  Top             =   1860
                  Width           =   855
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "NPWP"
                  BeginProperty Font 
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
                  Left            =   6540
                  TabIndex        =   85
                  Top             =   870
                  Width           =   555
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nama Customer"
                  BeginProperty Font 
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
                  Left            =   60
                  TabIndex        =   84
                  Top             =   1230
                  Width           =   1545
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
                  Left            =   1860
                  TabIndex        =   83
                  Top             =   1230
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
                  Left            =   7965
                  TabIndex        =   82
                  Top             =   870
                  Width           =   60
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipe Customer"
                  BeginProperty Font 
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
                  Left            =   420
                  TabIndex        =   81
                  Top             =   225
                  Width           =   1410
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   " Data Pelanggan"
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
                  Left            =   0
                  TabIndex        =   80
                  Top             =   0
                  Width           =   1590
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Passport / KITAS"
                  BeginProperty Font 
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
                  Left            =   45
                  TabIndex        =   79
                  Top             =   885
                  Width           =   1680
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
                  Left            =   1860
                  TabIndex        =   78
                  Top             =   885
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
                  Left            =   1845
                  TabIndex        =   77
                  Top             =   225
                  Width           =   60
               End
            End
            Begin VB.Label lblFotoKTPCustomer 
               BackColor       =   &H80000016&
               Caption         =   "FOTO KTP CUSTOMER"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   0
               TabIndex        =   227
               Top             =   5160
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.Image Image3 
               Height          =   3135
               Left            =   0
               Stretch         =   -1  'True
               Top             =   5520
               Visible         =   0   'False
               Width           =   5055
            End
            Begin VB.Image Image2 
               Height          =   3495
               Left            =   5160
               Stretch         =   -1  'True
               Top             =   5160
               Visible         =   0   'False
               Width           =   5895
            End
         End
      End
      Begin ucXPButton.XPButton cmdAdd 
         Height          =   495
         Left            =   6240
         TabIndex        =   0
         Top             =   9120
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":80BB5
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":80BD1
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":8101D
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
         Left            =   9840
         TabIndex        =   7
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":814E5
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":81501
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":819D5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdCreate 
         Height          =   495
         Left            =   1455
         TabIndex        =   2
         Top             =   9045
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&r"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":81F20
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":81F3C
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":82471
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdOpen 
         Height          =   495
         Left            =   5025
         TabIndex        =   6
         Top             =   9375
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&o"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":82A64
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":82A80
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":82F41
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
         Left            =   1440
         TabIndex        =   3
         Top             =   9360
         Visible         =   0   'False
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":83497
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":834B3
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":839C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdCancel 
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   9360
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&c"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":83F68
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":83F84
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":844C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdDelete 
         Height          =   495
         Left            =   3840
         TabIndex        =   5
         Top             =   9360
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&d"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":84ABF
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":84ADB
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":85063
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdEdit 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   9360
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&e"
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":8569F
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":856BB
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":85B66
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   8520
         TabIndex        =   185
         Top             =   9360
         Width           =   1170
         _ExtentX        =   2064
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
         MICON           =   "frmSIDMasterPelangganNonMerge.frx":8608B
         PICN            =   "frmSIDMasterPelangganNonMerge.frx":860A7
         PICH            =   "frmSIDMasterPelangganNonMerge.frx":865BA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: MASTER PELANGGAN ::."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   4380
         TabIndex        =   164
         Top             =   60
         Width           =   2775
      End
      Begin VB.Label lblJml 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10320
         TabIndex        =   9
         Top             =   60
         Width           =   1020
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11400
         Y1              =   9240
         Y2              =   9240
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11400
         Y1              =   9960
         Y2              =   9960
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11400
      End
   End
End
Attribute VB_Name = "frmSIDMasterPelangganNonMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private rsFind  As ADODB.Recordset
Private rsMove  As ADODB.Recordset
Private rsBy    As New ADODB.Recordset
Private onFind  As Boolean
Private onF12   As Boolean
Private xWAC    As Boolean
Private passBy  As String
Private nValid  As String
Private xAs()   As String
Private kliPr   As Boolean
Private kliKo   As Boolean
Private kliKc   As Boolean
Private kliKl   As Boolean
Private kliPrp  As Boolean
Private kliKop  As Boolean
Private kliKcp  As Boolean
Private kliKlp  As Boolean
Private kliPrs  As Boolean
Private kliKos  As Boolean
Private kliKcs  As Boolean
Private kliKls  As Boolean

Private nDpn As String
Private nTgh As String
Private nBlk As String
Private kodeLgnLama As String
Private isEditData As Boolean
Private isLoad As Boolean
Private noRangka As String
Private isBolehArmst As Boolean
'20180313 REV'LONCAT JIKA KODELGN SUDAH ADA D ACC
Dim iKodelgn As Integer
Private rsTemp As ADODB.Recordset
Private rsInput As ADODB.Recordset
Private isTambahNew As Boolean

'' 20181011REV1 | REV | PK M.WASISTO | PENAMBAHAN SCORING SINGLE ID
Private isWajibKTP As Boolean
Private iProgress As Double
Dim cols As TrueOleDBGrid70.Columns

Private filterRsTemp As String

Public Sub isEdit(Optional value As Boolean = False)
    isEditData = value
End Sub

Public Property Let setKodeLgnLama(value As String)
    kodeLgnLama = NulSID(value)
End Property

Public Property Let setNoRangka(value As String)
    noRangka = NulSID(value)
End Property

Public Property Let MyRsBy(value As ADODB.Recordset)
    Set rsBy = value
End Property

Public Property Let MyPassBy(value As String)
    passBy = value
End Property

Public Property Get rsMyMove() As ADODB.Recordset
    Set rsMyMove = rsMove
End Property

Public Property Let WAC(value As Boolean)
    xWAC = value
End Property

Public Property Let Asura(value As String)
    xAs() = Split(value, ";")
End Property

Private Sub ARProgressBar1_GotFocus()

End Sub

Private Sub cboAgamaS_GotFocus()
    If cboAgamaS.BackColor = fColor Then
        cboAgamaS.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboAgamaS.Locked = False
    End If
End Sub

Private Sub cboKewarganegaraan_Change()
    If UCase(cboTipe.text) = "PERORANGAN" Then
        If UCase(cboKewarganegaraan.text) = "WNI" Then
            Label5(0).Caption = "No. KTP"
        ElseIf UCase(cboKewarganegaraan.text) = "WNA" Then
            Label5(0).Caption = "Passport / KITAS"
        End If
    End If
End Sub

Private Sub cboKewarganegaraan_Click()
    Call cboKewarganegaraan_Change
End Sub

Private Sub cboOlah_GotFocus()
    If cboOlah.BackColor = fColor Then
        cboOlah.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboOlah.Locked = False
    End If
End Sub

Private Sub cmdClear_Click()
    Call prosesScoringTerakhir
End Sub

Private Function simpanImageKTPSusulan() As Boolean
    Dim fotoStream      As New ADODB.Stream
    Dim readStream      As ADODB.Stream
    Dim rsUploadKTP     As New ADODB.Recordset
    
   On Error GoTo simpanImageKTPSusulan_Error
    sSQL = "SELECT TOP 1 * FROM UploadKTP WHERE IDCUST = '" & NulSID(txt(1)) & "' ORDER BY TGLENTRY DESC "
    If rsUploadKTP.State Then rsUploadKTP.Close
    If fotoStream.State Then fotoStream.Close
    rsUploadKTP.Open sSQL, SQLConSID, adOpenKeyset, adLockOptimistic
     
        With fotoStream
            .Type = adTypeBinary
            .Open
            If CommonDialog1.FileName <> "" Then
            .LoadFromFile CommonDialog1.FileName
             
            End If
            If rsUploadKTP.EOF Then
                rsUploadKTP.AddNew
                rsUploadKTP!idcust = NulSID(txt(1))
                rsUploadKTP!delIndex = NulSID(gKode & Left(McDelIndex(7), 8))
                rsUploadKTP!App = NulSID("SV")
                rsUploadKTP!dealer = NulSID(Right(gDb, 2))
                rsUploadKTP!userentry = NulSID(gSA)
                rsUploadKTP!PIC = .Read
                rsUploadKTP!tglentry = isnow
                rsUploadKTP.Update
            Else
                rsUploadKTP!PIC = .Read
                rsUploadKTP!UpdateKe = CDbl(rsUploadKTP!UpdateKe) + 1
                rsUploadKTP!UserUpdate = NulSID(gSA)
                rsUploadKTP!TglUpdate = isnow
                rsUploadKTP.Update
            End If
        End With
        Set fotoStream = Nothing
        simpanImageKTPSusulan = True
   On Error GoTo 0
   Exit Function

simpanImageKTPSusulan_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    simpanImageKTPSusulan = False
    varErlError = Erl
    varErrNumber = err.Number
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure simpanImageKTPSusulan of Form frmSIDMasterPelangganNonMerge")
End Function

Private Function simpanImageKTP() As Boolean
    Dim fotoStream      As New ADODB.Stream
    Dim readStream      As ADODB.Stream
    Dim rsUploadKTP     As New ADODB.Recordset
   On Error GoTo simpanImageKTP_Error
    
    sSQL = "SELECT TOP 1 * FROM UploadKTP "
    If rsUploadKTP.State Then rsUploadKTP.Close
    If fotoStream.State Then fotoStream.Close
    rsUploadKTP.Open sSQL, SQLConSID, adOpenKeyset, adLockOptimistic
     
        With fotoStream
            .Type = adTypeBinary
            .Open
            If CommonDialog1.FileName <> "" Then
            .LoadFromFile CommonDialog1.FileName
             
            End If
            rsUploadKTP.AddNew
            rsUploadKTP!idcust = NulSID(txt(1))
            rsUploadKTP!delIndex = NulSID(gKode & Left(McDelIndex(7), 8))
            rsUploadKTP!App = NulSID("SV")
            rsUploadKTP!dealer = NulSID(Right(gDb, 2))
            rsUploadKTP!userentry = NulSID(gSA)
            rsUploadKTP!PIC = .Read
            rsUploadKTP!tglentry = isnow
            rsUploadKTP.Update
             
        End With
        Set fotoStream = Nothing
        simpanImageKTP = True
   On Error GoTo 0
   Exit Function
simpanImageKTP_Error:
    simpanImageKTP = False
    MsgBox "Error " & Erl & " (" & err.Description & ") in procedure simpanImageKTP of Form frmSIDMasterPelangganNonMerge"
End Function

 
Private Sub cmdNotifProses_Click()
    With frmSIDCustomer
        If .Visible Then
            .text(2) = Trim(gBantuan)
            Unload Me
            If frmSIDSearchCust.Visible Then
                Unload frmSIDSearchCust
            End If
            
            Exit Sub
        End If
    End With
    
    With frmBooking
        If .Visible Then
            .text(2) = Trim(gBantuan)
            Unload Me
            If frmSIDSearchCust.Visible Then
                Unload frmSIDSearchCust
            End If
           
            Exit Sub
        End If
    End With
    
    Unload Me

   On Error GoTo 0
   Exit Sub
End Sub

Private Sub cmdNotifUbah_Click()
   Frame3.Visible = False
    isEditData = True
    txt(1).text = gBantuan
    noRangka = ""
End Sub

Private Sub cmdProses_Click()
          Dim kodeLgn As String
          Dim boleh As Boolean
          Dim lengkapi As String
          Dim isBegin As Boolean
          Dim vStatus As String
          Dim hasilKTP As Boolean
          Dim idPilih As String
1        On Error GoTo cmdProses_Click_Error
          Dim isPilihID As Boolean
2         isPilihID = False
3         isWajibKTP = False
4         idPilih = ""
5         If rsTemp.RecordCount > 0 Then
6             isWajibKTP = True
7         End If
8         If rsTemp.RecordCount > 0 Then
9             filterRsTemp = rsTemp.filter
10            rsTemp.MoveFirst
11            rsTemp.filter = " pilih = true"
12            If rsTemp.RecordCount > 0 Then
13                isPilihID = True
14                idPilih = NulSID(rsTemp!kodeLgn)
15            End If
16            rsTemp.filter = filterRsTemp
              
17        End If
18        If isPilihID = False Then
19            If rsInput!Pilih Then
20                isPilihID = True
21            End If
22        End If
          
23        If isPilihID = False Then
24            MsgBox "Silahkan pilih data yang akan disimpan.", vbInformation
25            Exit Sub
26        End If
27        If isTambahNew Then
doCariID:
28            sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE idcust = '" & NulSID(txt(1)) & "' "
29            Set rsCari = objO.FetchData(sSQL)
30            If Not rsCari.EOF Then
      '            MsgBox "Id Cust sudah ada atas nama " & NulSID(NulSID(rsCari!NamaCust1) & " " & NulSID(rsCari!namacust2) & " " & NulSID(rsCari!namacust3)) & vbCrLf & "Silahkan lakukan proses simpan ulang.", vbInformation
31                txt(1).text = LastID(cboTipe.text)
       '           Exit Sub
32                GoTo doCariID
33            End If
34            hasilKTP = False
35            If NulSID(txtUploadKTP.text) <> "" Then
36                hasilKTP = simpanImageKTP
37                If hasilKTP = False Then
38                    MsgBox "Gagal Upload KTP", vbInformation
39                    Exit Sub
40                End If
41            Else
42                If isWajibKTP = True Then
43                    MsgBox "Upload KTP wajib diupload sebelum VALIDASI ORDER", vbInformation
44                End If
45            End If
46            If NulSID(txtNpwp) = "" Or Val(NulSID(txtNpwp)) = 0 Then
47                txtNpwp.text = "000.000.000.0-000.000"
48            ElseIf InStr(1, txtNpwp.text, "_") Then
49                MsgBox "No NPWP belum lengkap.", vbInformation
                  'txt(1).Text = LastID(cboTipe.Text)
50                Exit Sub
51            End If
                
              Dim namaLengkap As String
              Dim Alamat As String
              Dim provinsi As String
              Dim Kota As String
              Dim kec As String
              Dim kel As String
52            provinsi = "": Kota = "": kec = "": kel = ""
              
53            Alamat = NulSID(txt(8))
                  
                  'If cmdBack.Visible = True Then
54            Alamat = Alamat & " Kel." & NulSID(txt(66)) & " Kec." & NulSID(txt(65)) & " " & NulSID(txt(64)) & " Prov." & NulSID(txt(63))
55            provinsi = UCase(NulSID(txt(63)))
56            Kota = UCase(NulSID(txt(64)))
57            kec = UCase(NulSID(txt(65)))
58            kel = UCase(NulSID(txt(66)))
                      
       
59            namaLengkap = NulSID(NulSID(txt(4)) & " " & NulSID(txt(5)) & " " & NulSID(txt(6)))
                  
60            If kodeLgnLama <> NulSID(txt(1)) And kodeLgnLama <> "" Then
61                kodeLgn = kodeLgnLama
62            Else
63                kodeLgn = NulSID(txt(1))
64            End If
                  
65            If kodeLgnLama = "" Then
66                kodeLgnLama = NulSID(txt(1))
67            End If
              
              '20/12/2012  Cust Non Axtif
              Dim rsNon As ADODB.Recordset
68            Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) where KodeLgn = '" & NulSID(kodeLgn) & "' and nonaktifflag  = 1")
69            If rsNon.RecordCount > 0 Then
70                MsgBox "Kode Langgan : " & txt(1) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut", vbCritical + vbOKOnly, "Master Pelanggan Non Aktif"
71                pelangganAda = False
72                Exit Sub
73            End If
              
74            isBegin = False
75            Call objHT.OpenTransactionAll
76            isBegin = True
77            If UCase(cboTipe.text) = UCase("Pemerintahan") Then
78                vStatus = "P"
79            Else
80                vStatus = "S"
81            End If
                  
82            Set rsSearch = objO.FetchDataACC("select KodeLgn from ARMST WITH(NOLOCK) where KodeLgn = '" & NulSID(kodeLgn) & "'")
83            If rsSearch.EOF Then
                  '' insert data armst
84                sSQL1 = "INSERT INTO [LOGACC]..ArmstLog (KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                  "        StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, TglEntry,FPGabungan,NonPPn,Afiliasi ,DBSumber,Keterangan,ClientName,KodeUserChange,TglChange,TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak, SIUP, No_Mou ) "
85                sSQL1 = sSQL1 & " VALUES ( '" & NulSID(kodeLgn) & "', " & _
                  "                  '" & s2d(NulSID(namaLengkap)) & "', " & _
                  "                  '" & s2d(Left(NulSID(kodeLgn), 3)) & "', " & _
                  "                  '" & s2d(NulSID(txt(8))) & "', " & _
                  "                  '" & s2d(NulSID(txt(66)) + ", " + NulSID(txt(65))) & "', " & _
                  "                  '" & s2d(NulSID(txt(64)) + ", " + NulSID(txt(63))) & "', " & _
                  "                  '" & s2d(NulSID(txt(14))) & "', " & _
                  "                  '" & s2d(NulSID(txt(12)) & NulSID(txt(13))) & "', " & _
                  "                  '" & s2d(NulSID(txtNpwp)) & "', " & _
                  "                  'D0', 0, " & _
                  "                  '" & s2d(NulSID(txt(18))) & "', 'GRP', " & _
                  "                  '" & s2d(vStatus) & "', '', " & _
                  "                  " & "0" & ", 'Customer " & cboTipe.text & "', 0, " & _
                  "                  '" & gSA & "', GETDATE(),0,0,0, " & _
                  "                  '" & Trim(gDb) & "','Data Baru','" & CompName & "' ,'" & gSA & "',Getdate(), " & _
                  "                  '" & "D0" & "', '" & "D0" & "', '" & "D0" & "', '" & "D0" & "', " & _
                  "                  " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & _
                  "                  " & 0 & ", '" & 0 & "', '', '', '' ) "   'A.Rev1711291-SingleId

86                sSQL = "INSERT INTO ARMST(KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2,Alamat3" _
                       & " ,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry,HubungDengan,Telepon1)" _
                      & " VALUES ('" & NulSID(kodeLgn) & "','" & NulSID(namaLengkap) & "','" & Left(NulSID(kodeLgn), 3) & "'" _
                      & ",'" & s2d(NulSID(txt(8))) & "','" & s2d(NulSID(txt(66)) + ", " + NulSID(txt(65))) & "','" & s2d(NulSID(txt(64)) + ", " + NulSID(txt(63))) & "','" & Replace(Replace(NulSID(txt(12) & txt(13)) & "," & NulSID(txt(14)) & "," & NulSID(txt(15)), ">", ""), " ", "") & "','" & NulSID(txt(17)) & "'" _
                      & ",'" & NulSID(txtNpwp) & "','D0','D0','D0','D0','D0','" & vStatus & "','" & gSA & "','" & Format(isnow, "MM/DD/YYYY") & "','" & NulSID(txt(18)) & "','" & NulSID(txt(19)) & "')"
                  
87                 Call objHT.InsertData_ACC(sSQL1)
88                 Call objHT.InsertData_ACC(sSQL)
                    
89            Else
90                If isBolehArmst = True Then
91                    sSQL1 = "INSERT INTO [LOGACC]..ArmstLog (KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                                                        "StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, TglEntry,FPGabungan,NonPPn,Afiliasi ,No_MOU,DBSumber,Keterangan,ClientName, " & _
                                                                        "KodeUserChange,TglChange,TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak,SIUP ) " & _
                                        "SELECT KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                                                        "StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, GETDATE(),FPGabungan,NonPPn,Afiliasi,No_MOU, '" & Trim(gDb) & "','Sebelum Diedit','" & CompName & "' ,'" & gSA & "'," & _
                                        "Getdate(),TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak,SIUP from Armst WHERE KodeLgn='" & NulSID(kodeLgn) & "' "

                       
92                    sSQL = "UPDATE ARMST SET NamaLgn = '" & NulSID(namaLengkap) & "'," _
                           & " KodeWil = '" & Left(NulSID(kodeLgn), 3) & "',Alamat1 = '" & s2d(NulSID(txt(8))) & "',Alamat2 = '" & s2d(NulSID(txt(66)) + ", " + NulSID(txt(65))) & "',Alamat3 = '" & s2d(NulSID(txt(64)) + ", " + NulSID(txt(63))) & "',TglEntry = '" & Format(isnow, "MM/DD/YYYY") & "'," _
                          & " Telepon = '" & Replace(Replace(NulSID(txt(12) & txt(13)) & "," & NulSID(txt(14)) & "," & NulSID(txt(15)), ">", ""), " ", "") & "',Npwp = '" & NulSID(txtNpwp) & "'," _
                          & " TOPGRP = 'D0',StatusPersh = '" & vStatus & "', KodeUser='" & gSA & "',Nofax = '" & NulSID(txt(17)) & "',HubungDengan = '" & NulSID(txt(18)) & "',Telepon1 = '" & NulSID(txt(19)) & "'" _
                          & " where KodeLgn = '" & NulSID(kodeLgn) & "'"
93                    sSQL = "UPDATE ARMST SET NamaLgn = '" & NulSID(namaLengkap) & "'," _
                           & " KodeWil = '" & Left(NulSID(kodeLgn), 3) & "',Alamat1 = '" & s2d(NulSID(txt(8))) & "',Alamat2 = '" & s2d(NulSID(txt(66)) + ", " + NulSID(txt(65))) & "',Alamat3 = '" & s2d(NulSID(txt(64)) + ", " + NulSID(txt(63))) & "',TglEntry = '" & Format(isnow, "MM/DD/YYYY") & "'," _
                          & " Telepon = '" & Replace(Replace(NulSID(txt(12) & txt(13)) & "," & NulSID(txt(14)) & "," & NulSID(txt(15)), ">", ""), " ", "") & "',Npwp = '" & NulSID(txtNpwp) & "'," _
                          & " StatusPersh = '" & vStatus & "', KodeUser='" & gSA & "',Nofax = '" & NulSID(txt(17)) & "',HubungDengan = '" & NulSID(txt(18)) & "',Telepon1 = '" & NulSID(txt(19)) & "'" _
                          & " where KodeLgn = '" & NulSID(kodeLgn) & "'"
94                     Call objHT.InsertData_ACC(sSQL1)
95                     Call objHT.InsertData_ACC(sSQL)
                        
96                End If
97            End If
              
               
98             Set rsSearch = objHT.FetchDataBengkel("select IDCust from SINGLEIDORI..NEWCUSTOMER WITH(NOLOCK) where IDCust = '" & NulSID(txt(1)) & "'")
99             If rsSearch.EOF Then
100                sSQL = "INSERT INTO SINGLEIDORI..NEWCUSTOMER (IDCust,TypeCustomer,NoKTP,noNPWP,Title,NamaCust1,NamaCust2,NamaCust3,tglLahir,Agama,Pekerjaan,Propinsi,Kota,Kecamatan,Kelurahan,KodePos,alamat,RT,RW,IdTelp,telp,HP1,HP2,Email,Hobi,Olahraga,Suku,Preferensi,FPGab,NonPPn,TopVhc,BatasVhc,TopGrp,BatasGrp,TopBp,BatasBp,TopPart,BatasPart,TglEntry,UserEntry,NoFax,namaPIC,telpPIC,kewarganegaraan,pinbb)"
101                sSQL = sSQL & " VALUES ('" & NulSID(txt(1)) & "','" & UCase(NulSID(cboTipe.text)) & "','" & UCase(NulSID(txt(2))) & "','" & NulSID(txtNpwp) & "','" & UCase(NulSID(cboTitle.text)) & "','" & UCase(NulSID(txt(4))) & "','" & UCase(NulSID(txt(5))) & "','" & UCase(NulSID(txt(6))) & "','" & Format(DtLahir.value, "mm/dd/yyyy") & "','" & UCase(txt(20).text) & "','" & IIf(UCase(NulSID(txt(7))) = "LAINNYA", UCase(NulSID(txt(23))), UCase(NulSID(txt(7)))) & "','" & provinsi & "','" & Kota & "','" & kec & "','" & kel & "',"
102                sSQL = sSQL & " '" & UCase(NulSID(txt(11))) & "','" & UCase(NulSID(txt(8))) & "','" & UCase(NulSID(txt(9))) & "','" & UCase(NulSID(txt(10))) & "','" & UCase(NulSID(txt(12))) & "','" & UCase(NulSID(txt(13))) & "','" & UCase(NulSID(txt(14))) & "','" & UCase(NulSID(txt(15))) & "','" & UCase(NulSID(txt(16))) & "',"
103                sSQL = sSQL & " '" & UCase(NulSID(txt(21).text)) & "','" & UCase(NulSID(txt(22).text)) & "','" & UCase(NulSID(txt(71))) & "','" & UCase(NulSID(txt(72))) & "',0,0,'D0',0,'D0',0,'D0',0,'D0',0,GETDATE(),'" & gSA & "','" & NulSID(txt(17)) & "','" & NulSID(txt(18)) & "','" & NulSID(txt(19)) & "','" & NulSID(cboKewarganegaraan.text) & "','" & IIf(isWajibKTP, "SCORING", "") & "')"
104            Else
105                sSQL = "UPDATE SINGLEIDORI..NEWCUSTOMER SET NoKTP = '" & UCase(NulSID(txt(2))) & "',noNPWP = '" & UCase(NulSID(txtNpwp)) & "',Title = '" & UCase(cboTitle.text) & "',NamaCust1='" & UCase(NulSID(txt(4))) & "',NamaCust2='" & UCase(NulSID(txt(5))) & "',NamaCust3='" & UCase(NulSID(txt(6))) & "',tglLahir='" & Format(DtLahir.value, "mm/dd/yyyy") & "',Agama='" & UCase(txt(20).text) & "',Pekerjaan='" & IIf(UCase(NulSID(txt(7))) = "LAINNYA", UCase(NulSID(txt(23))), UCase(NulSID(txt(7)))) & "',Propinsi='" & provinsi & "',Kota='" & Kota & "',Kecamatan='" & kec & "',"
106                sSQL = sSQL & " Kelurahan='" & kel & "',KodePos='" & UCase(NulSID(txt(11))) & "',alamat='" & UCase(NulSID(txt(8))) & "',RT='" & UCase(NulSID(txt(9))) & "',RW='" & UCase(NulSID(txt(10))) & "',IdTelp='" & UCase(NulSID(txt(12))) & "',telp='" & UCase(NulSID(txt(13))) & "',HP1='" & UCase(NulSID(txt(14))) & "',HP2='" & UCase(NulSID(txt(15))) & "',Email='" & UCase(NulSID(txt(16))) & "',Hobi='" & UCase(NulSID(txt(21).text)) & "',Olahraga='" & UCase(NulSID(txt(22).text)) & "',Suku='" & UCase(NulSID(txt(71))) & "',Preferensi='" & UCase(NulSID(txt(72))) & "',NoFax = '" & NulSID(txt(17)) & "',namaPIC= '" & NulSID(txt(18)) & "',telpPIC= '" & NulSID(txt(19)) & "',kewarganegaraan = '" & NulSID(cboKewarganegaraan.text) & "',UserUpdate = '" & NulSID(gSA) & "',TglUpdate = GETDATE(),DealerUpdate = '" & Right(gDb, 2) & "', AppUpdate = 'SVC' WHERE IDCust = '" & NulSID(txt(1)) & "' "
107            End If
               
108            Call objHT.InsertData_Bengkel(sSQL)
               
109            Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM SINGLEIDORI..Bridging WHERE idcust = '" & NulSID(txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' ")
110            If rsSearch.EOF Then
111                sSQL = "INSERT INTO SINGLEIDORI..Bridging(KodeDealer,kodeLgn,IDCust,TglEntry,UserEntry) VALUES ('" & Right(gDb, 2) & "','" & kodeLgnLama & "','" & NulSID(txt(1)) & "',getdate(),'" & gSA & "')"
112                Call objHT.InsertData_Bengkel(sSQL)
113            End If
114            Call objHT.CloseTransactionAll
115            gBantuan = Trim(txt(1))
      '            Call hapusForm
      '            Text(57) = ""
                  
                  
116       Else
              
117           If MsgBox("Dengan memilih proses berarti anda menggunakan data customer yang di suggest oleh sistem, jika ada perubahan data silahkan edit data tersebut", vbOKCancel + vbQuestion, "Konfirmasi") = vBok Then
                  ''20190129
118               If NulSID(idPilih) = "" Then
119                   MsgBox "Tidak ada ID Customer yang dipilih", vbInformation
120                   Exit Sub
121               Else
122                   gBantuan = NulSID(idPilih) 'NulSID(rsTemp!kodeLgn)
123                   Unload Me
124               End If
125           Else
126               Exit Sub
127           End If
              
128       End If
          
129       If isTambahNew Then
130           MsgBox "Data berhasil disimpan", vbInformation
              ''20190201 | REV | DITAMBAHKAN PILIHAN PROSES ATAU EDIT
              '' ---------------------------------------------------------------
131           cmdProses.Enabled = False
132           frameNotif.Visible = True
133       Else
134           With frmSIDCustomer
135               If .Visible Then
136                   .text(2) = Trim(gBantuan)
137                   Unload Me
138               End If
139           End With
              
140           With frmBooking
141               If .Visible Then
142                   .text(2) = Trim(gBantuan)
143                   Unload Me
144               End If
145           End With
              
146           Unload Me
147       End If
          
          

148      On Error GoTo 0
149      Exit Sub

cmdProses_Click_Error:
150       If isBegin Then
151           Call objHT.CloseTransactionAll_Err
152       End If
153       If hasilKTP = True Then
154           sSQL = "DELETE FROM SINGLEIDORI..UploadKTP WHERE IDCUST = '" & kodeLgnLama & "'"
155           Call objO.InsertData(sSQL)
156       End If
157       MsgBox "Error " & Erl & " (" & err.Description & ") in procedure cmdProses_Click of Form frmSIDMasterPelangganNonMerge"
End Sub

Private Sub cmdSimpan_Click()
         
         
         Dim kodeLgn As String
1        On Error GoTo cmdSimpan_Click_Error
          Dim boleh As Boolean
          Dim lengkapi As String
          Dim isBegin As Boolean
          Dim vStatus As String
clickUlang:
2         boleh = True
3         lengkapi = ""
          
4         If UCase(cboTipe.text) <> "PERORANGAN" And NulSID(cboTipe.text) <> "" Then
5             If NulSID(txt(18)) = "" Then
6                 boleh = False
7                 lengkapi = lengkapi & " Nama PIC,"
8             End If
9             If NulSID(txt(19)) = "" Then
10                boleh = False
11                lengkapi = lengkapi & " Telp PIC,"
12            End If
13        End If
          
14        If NulSID(txt(1)) = "" Then
15            boleh = False
16            lengkapi = lengkapi & " Kode Customer,"
17        End If
          
          ''NAMA LENGKAP
18        If NulSID(txt(4) & txt(5) & txt(6)) = "" Then
19            boleh = False
20            lengkapi = lengkapi & " Nama Customer,"
21        End If
          
          ''EMAIL
      '    If nulsid(txt(16)) = "" Then
      '        boleh = False
      '        lengkapi = lengkapi & " Email,"
      '    End If
          
          ''ALAMAT (provinsi, kab/kota, kecamatan, kelurahan, jalan)
      '    If cmdBack.Visible = False Then
      '        If cboProv.Text = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Provinsi,"
      '        End If
      '
      '        If cboKota.Text = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Kota,"
      '        End If
      '
      '        If cboKec.Text = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Kecamatan,"
      '        End If
      '
      '        If cboKel.Text = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Kelurahan,"
      '        End If
      '
      '        If nulsid(txt(8)) = "" Then
      '            boleh = False
      '            lengkapi = lengkapi & " Jalan/Kompleks,"
      '        End If
      '    Else
22            If NulSID(txt(63)) = "" Then
23                boleh = False
24                lengkapi = lengkapi & " Provinsi,"
25            End If
              
26            If NulSID(txt(64)) = "" Then
27                boleh = False
28                lengkapi = lengkapi & " Kota,"
29            End If
              
30            If NulSID(txt(65)) = "" Then
31                boleh = False
32                lengkapi = lengkapi & " Kecamatan,"
33            End If
              
34            If NulSID(txt(66)) = "" Then
35                boleh = False
36                lengkapi = lengkapi & " Kelurahan,"
37            End If
              
38            If NulSID(txt(8)) = "" Then
39                boleh = False
40                lengkapi = lengkapi & " Jalan/Kompleks,"
41            End If
      '    End If
          
42        If NulSID(txt(2)) = "" Or Val(NulSID(txt(2))) = "0" Then
              '' 20180612 || REVISI jika kitas boleh diisi huruf jika kosong baru ditampilkan pesan
43            If Label5(0).Caption = "Passport / KITAS" And NulSID(txt(2)) = "" Then
                  'MsgBox Label5(0).Caption & " wajib diisi sebelum VALIDASI ORDER", vbInformation
44            Else
45                If Label5(0).Caption <> "Passport / KITAS" Then
                      'MsgBox Label5(0).Caption & " wajib diisi sebelum VALIDASI ORDER", vbInformation
46                End If
47                If Label5(0).Caption = "No. KTP" Then
                     ' Txt(2) = "0000000000000000"
48                End If
49            End If
50        End If
          
51        If Log.singleIDScoring = 1 And Label5(0).Caption = "No. KTP" And NulSID(txt(2)) <> "" Then
52            If Len(NulSID(txt(2))) < 16 Then
53                boleh = False
54                lengkapi = lengkapi & " No KTP harus 16 karakter,"
55            End If
56            If NulSID(txt(2)) = "" Then
57                boleh = False
58                lengkapi = lengkapi & " No KTP harus diisi,"
59            End If
              
60            If Log.lockKTPNoHPScoring = 1 Then
61                If digitsama(NulSID(txt(2)), 5) >= 6 Then
62                    boleh = False
63                    lengkapi = lengkapi & " 6 digit pada No KTP mengandung kata yang sama,"
64                End If
65            End If
66        End If
          
67        If Log.singleIDScoring = 1 And NulSID(cboTipe.text) = "PERORANGAN" Then
68            If NulSID(DtLahir.value) = "" Then
69                boleh = False
70                lengkapi = lengkapi & "Tanggal lahir harus diisi, "
71            ElseIf DtLahir.value = "01/01/1900" Then
72                boleh = False
73                lengkapi = lengkapi & "Tanggal lahir harus diisi, "
74            ElseIf DtLahir.value = "01/01/3000" Then
75                boleh = False
76                lengkapi = lengkapi & "Tanggal lahir harus diisi, "
77            End If
78        End If
          
79        If Log.singleIDScoring = 1 Then
80            If Len(NulSID(txt(14))) < 8 Then
81                boleh = False
82                lengkapi = lengkapi & "No HP minimal 8 karakter, "
83            End If
              
84            If Log.lockKTPNoHPScoring = 1 Then
85                If digitsama(NulSID(txt(14)), 3) >= 4 Then
86                    boleh = False
87                    lengkapi = lengkapi & " 4 digit pada No HP mengandung kata yang sama,"
88                End If
89            End If
90        End If
          
91        If NulSID(txt(14)) = "" And NulSID(txt(15)) = "" Then
92            boleh = False
93            lengkapi = lengkapi & " No HP,"
94        End If
          
95        If NulSID(DtLahir.value) <> "" Then
96            If DateDiff("D", Format(Now, "dd/mm/yyyy"), Format(DtLahir.value, "dd/mm/yyyy")) > 0 Then
97                boleh = False
98                lengkapi = lengkapi & " Tanggal Lahir melebihi tanggal hari ini,"
99            End If
100       End If
          
101       If boleh = True Then
              'If Get_ConnectionAcces(NisDb, NisUSER, NisPass, NisKomp) Then
                  Dim rsCari As ADODB.Recordset
102               sSQL = "SELECT * FROM [SINGLEIDORI]..TOPOLOGYDAERAH WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(63)) & "' and KOTA = '" & NulSID(txt(64)) & "' and KECAMATAN = '" & NulSID(txt(65)) & "' and KELURAHAN = '" & NulSID(txt(66)) & "' and KODEPOS = '" & NulSID(txt(11)) & "' "
103               Set rsCari = objO.FetchData(sSQL)
104               If rsCari.EOF Then
105                   MsgBox "PROPINSI / KOTA / KECAMATAN / KELURAHAN / KODEPOS tidak cocok", vbInformation
106                   Exit Sub
107               End If
                  '' 20180927DEV  ' SCORING SINGLE ID
108               If UCase(cboTipe.text) = "PERORANGAN" And Log.singleIDScoring = 1 And isEditData = False Then
109                    isWajibKTP = False
110                    Call prosesScoring
111                    Exit Sub
112               End If
                  
113               If Not (NulSID(txt(2)) = "" Or Val(NulSID(txt(2))) = "0") Then
114                   sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE idcust = '" & NulSID(txt(1)) & "' "
115                   Set rsCari = objO.FetchData(sSQL)
116                   If Not rsCari.EOF Then
117                       If NulSID(txt(2)) <> NulSID(rsCari!noKTP) Then
118                           sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NOKTP = '" & NulSID(txt(2)) & "' "
119                           Set rsCari = objO.FetchData(sSQL)
120                           If Not rsCari.EOF Then
                                  'MsgBox "No KTP sudah ada, tidak boleh sama.", vbInformation
                                  '' 20190208 || edit no ktp bebas di UR tidak ada (M. Wasisto)
                                  'MsgBox Label5(0) & " sudah ada, tidak boleh sama.", vbInformation
                                  'Exit Sub
121                           End If
122                       End If
123                   End If
124               End If
                  
                  ''20180130 KELOLOSAN KODELGN LAMA TIDAK TERISI DI BRIDGING
                  ''------------------------------------------------------------------------
125               If kodeLgnLama = "" Then
126                   sSQL = "SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE idcust = '" & NulSID(txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' "
127                   Set rsCari = objO.FetchData(sSQL)
128                   If Not rsCari.EOF Then
129                       kodeLgnLama = NulSID(rsCari!kodeLgn)
130                   Else
131                       kodeLgnLama = NulSID(txt(1))
132                   End If
133               Else
134                   sSQL = "SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE kodelgn = '" & NulSID(kodeLgnLama) & "' and idcust = '" & NulSID(txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' "
135                   Set rsCari = objO.FetchData(sSQL)
136                   If rsCari.EOF Then
                          ''jika tidak ada kodelgn lama dan idcust , cek di bridging untuk kodelgn baru
137                       sSQL = "SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE idcust = '" & NulSID(txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' "
138                       Set rsCari = objO.FetchData(sSQL)
139                       If Not rsCari.EOF Then
140                           kodeLgnLama = NulSID(rsCari!kodeLgn)
141                       Else
142                           kodeLgnLama = NulSID(txt(1))
143                       End If
144                   End If
145               End If
                  
146               If kodeLgnLama <> NulSID(txt(1)) And kodeLgnLama <> "" Then
147                   kodeLgn = kodeLgnLama
148               Else
149                   kodeLgn = NulSID(txt(1))
150               End If
                  
151               If isEditData = False Then
152                   sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE idcust = '" & NulSID(txt(1)) & "' "
153                   Set rsCari = objO.FetchData(sSQL)
154                   If Not rsCari.EOF Then
155                       MsgBox "Id Cust sudah ada atas nama " & NulSID(NulSID(rsCari!namacust1) & " " & NulSID(rsCari!namacust2) & " " & NulSID(rsCari!namacust3)) & vbCrLf & "Silahkan lakukan proses simpan ulang.", vbInformation
156                       txt(1).text = LastID(cboTipe.text)
157                       Exit Sub
158                   End If
                      
159                   Set rsCari = objO.FetchDataACC("SELECT NamaLgn FROM armst WITH(NOLOCK) where KodeLgn = '" & NulSID(txt(1)) & "'")
160                   If rsCari.RecordCount > 0 Then
                          'MsgBox "Id Cust sudah ada atas nama " & NulSID(NulSID(rsCari!Namalgn)) & vbCrLf & "Silahkan lakukan proses simpan ulang.", vbInformation
                          '20180313 REV'LONCAT JIKA KODELGN SUDAH ADA D ACC
161                       iKodelgn = iKodelgn + 1
162                       txt(1).text = LastID(cboTipe.text, iKodelgn)
                          'Exit Sub
163                       GoTo clickUlang
                          
164                   End If
                      
                      ''MENGECEK NO KTP / NO NPWP
165                   If UCase(cboTipe.text) = "PERORANGAN" Then
                          ''KTP TIDAK BOLEH KEMBAR
166                       If NulSID(txt(2)) = "" Or Val(NulSID(txt(2))) = 0 Then
167                           txt(2) = ""
168                       Else
169                           sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NoKTP = '" & NulSID(txt(2)) & "' "
170                           Set rsCari = objO.FetchData(sSQL)
171                           If Not rsCari.EOF Then
                                  'MsgBox "No KTP sudah ada atas nama " & NulSID(NulSID(rsCari!NamaCust1) & " " & NulSID(rsCari!NamaCust2) & " " & NulSID(rsCari!NamaCust3)) & vbCrLf & "Silahkan diisi dengan data lain.", vbInformation
172                               MsgBox Label5(0) & " sudah ada atas nama " & NulSID(NulSID(rsCari!namacust1) & " " & NulSID(rsCari!namacust2) & " " & NulSID(rsCari!namacust3)) & vbCrLf & "Silahkan diisi dengan data lain.", vbInformation
                                  'txt(1).Text = LastID(cboTipe.Text)
173                               Exit Sub
174                           End If
175                       End If
176                   Else
                          ''NPWP TIDAK BOLEH KEMBAR
177                       If NulSID(txtNpwp) = "" Or Val(NulSID(txtNpwp)) = 0 Then
178                           txtNpwp.text = "000.000.000.0-000.000"
179                       Else
180                           sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE noNPWP = '" & NulSID(txtNpwp) & "' "
181                           Set rsCari = objO.FetchData(sSQL)
182                           If Not rsCari.EOF Then
183                               MsgBox "No NPWP sudah ada atas nama " & NulSID(NulSID(rsCari!namacust1) & " " & NulSID(rsCari!namacust2) & " " & NulSID(rsCari!namacust3)) & vbCrLf & "Silahkan diisi dengan data lain.", vbInformation
                                  'txt(1).Text = LastID(cboTipe.Text)
184                               Exit Sub
185                           End If
186                       End If
                          
187                   End If
188               Else
189                   If cmdUploadSusulan.Visible And NulSID(cboTipe) = "PERORANGAN" Then
190                       If NulSID(txtUploadKTPSusulan.text) <> "" Then
191                           If simpanImageKTPSusulan = False Then
192                               Exit Sub
193                           End If
194                       End If
195                   End If
196               End If
                  
197               If NulSID(txtNpwp) = "" Or Val(NulSID(txtNpwp)) = 0 Then
198                   txtNpwp.text = "000.000.000.0-000.000"
199               ElseIf InStr(1, txtNpwp.text, "_") Then
200                   MsgBox "No NPWP belum lengkap.", vbInformation
                      'txt(1).Text = LastID(cboTipe.Text)
201                   Exit Sub
202               End If
                    
                  Dim namaLengkap As String
                  Dim Alamat As String
                  Dim provinsi As String
                  Dim Kota As String
                  Dim kec As String
                  Dim kel As String
203               provinsi = "": Kota = "": kec = "": kel = ""
                  
204               Alamat = NulSID(txt(8))
                  
                  'If cmdBack.Visible = True Then
205                   Alamat = Alamat & " Kel." & NulSID(txt(66)) & " Kec." & NulSID(txt(65)) & " " & NulSID(txt(64)) & " Prov." & NulSID(txt(63))
206                   provinsi = UCase(NulSID(txt(63)))
207                   Kota = UCase(NulSID(txt(64)))
208                   kec = UCase(NulSID(txt(65)))
209                   kel = UCase(NulSID(txt(66)))
                      
      '            Else
      '                Alamat = Alamat & " Kel." & nulsid(cboKel.Text) & " Kec." & nulsid(cboKec.Text) & " " & nulsid(cboKota.Text) & " Prov." & nulsid(cboProv.Text)
      '                provinsi = UCase(cboProv.Text)
      '                Kota = UCase(cboKota.Text)
      '                kec = UCase(cboKec.Text)
      '                kel = UCase(cboKel.Text)
      '            End If
                  
210               namaLengkap = NulSID(NulSID(txt(4)) & " " & NulSID(txt(5)) & " " & NulSID(txt(6)))
                  
                  
                  
211               If kodeLgnLama = "" Then
212                   kodeLgnLama = NulSID(txt(1))
213               End If
                  
                  '20/12/2012  Cust Non Axtif
                  Dim rsNon As ADODB.Recordset
214               Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) where KodeLgn = '" & NulSID(kodeLgn) & "' and nonaktifflag  = 1")
215               If rsNon.RecordCount > 0 Then
216                   MsgBox "Kode Langgan : " & txt(1) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut", vbCritical + vbOKOnly, "Master Pelanggan Non Aktif"
217                   pelangganAda = False
218                   Exit Sub
219               End If
                  
                  ''sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NamaCust1 = '" & nulsid(txt(4)) & "' and NamaCust2 = '" & nulsid(txt(5)) & "' and NamaCust3 = '" & nulsid(txt(6)) & "' and Pekerjaan = '" & UCase(nulsid(txt(7))) & "' and Alamat = '" & nulsid(txt(8)) & "' "
                  ''20171227 || 7.  Saat save, untuk pengecekan data duplikasi yang sebelumnya menggunakan Nama+Alamat+Pekerjaan, untuk Alamat, nama jalan/kompleks/no tidak dicek
                  Dim namaSama As String
220               If NulSID(txt(4)) = "" Then
221                   MsgBox "Nama Customer 1 harus diisi terlebih dahulu", vbInformation
222                   Exit Sub
223               End If
                  ''20180130
224               namaSama = "(namacust1+NamaCust2+NamaCust3 LIKE '%" & IIf(NulSID(txt(4)) <> "", NulSID(txt(4)) & "%", "") & IIf(NulSID(txt(5)) <> "", NulSID(txt(5)) & "%", "") & IIf(NulSID(txt(6)) <> "", NulSID(txt(6)) & "%", "") & "%')"
                  
                  'sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NamaCust1 = '" & NulSID(Txt(4)) & "' and NamaCust2 = '" & NulSID(Txt(5)) & "' and NamaCust3 = '" & NulSID(Txt(6)) & "' and Pekerjaan = '" & UCase(NulSID(Txt(7))) & "' and Propinsi = '" & NulSID(Txt(63)) & "' and Kota = '" & NulSID(Txt(64)) & "' and Kecamatan = '" & NulSID(Txt(65)) & "' and Kelurahan = '" & NulSID(Txt(66)) & "'"
225               sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE " & namaSama & " and Pekerjaan = '" & UCase(NulSID(txt(7))) & "' and Propinsi = '" & NulSID(txt(63)) & "' and Kota = '" & NulSID(txt(64)) & "' and Kecamatan = '" & NulSID(txt(65)) & "' and Kelurahan = '" & NulSID(txt(66)) & "'"
                  
226               Set rsNon = objO.FetchData(sSQL)
227               If rsNon.RecordCount > 0 And isEditData = False Then
228                   MsgBox "Nama Customer, Alamat dan Pekerjaan sudah ada dengan ID Pelanggan " & rsNon!idcust & vbCrLf & "Silahkan ganti dengan data yang lain.", vbInformation
                      'pelangganAda = False
229                   Exit Sub
230               End If
                  
231               If NulSID(txt(2)) <> "" And Val(NulSID(txt(2))) <> 0 And isEditData = True Then
232                   sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE NoKTP = '" & NulSID(txt(2)) & "' and idcust <> '" & NulSID(txt(1)) & "' "
233                   Set rsNon = objO.FetchData(sSQL)
234                   If rsNon.RecordCount > 0 Then
                          'MsgBox "No KTP " & NulSID(txt(2)) & " sudah ada." & vbCrLf & "Silahkan ganti dengan data yang lain.", vbInformation
                          '' 20190208 || edit no ktp bebas di UR tidak ada (M. Wasisto)
235                       If Log.singleIDScoring = 0 Then
236                           MsgBox Label5(0) & " " & NulSID(txt(2)) & " sudah ada." & vbCrLf & "Silahkan ganti dengan data yang lain.", vbInformation
237                           pelangganAda = False
238                           Exit Sub
239                       End If
240                   End If
241               End If
                  
      '            If Left(Trim(Text(57)), 3) = "AFL" Then
      '                MsgBox "Tidak diperkenankan membuat Kode Pelanggan Afiliasi.", vbCritical + vbOKOnly, "Master Pelanggan Salah"
      '                pelangganAda = False
      '                Exit Sub
      '            End If
242               isBegin = False
243               Call objHT.OpenTransactionAll
244               isBegin = True
245               If UCase(cboTipe.text) = UCase("Pemerintahan") Then
246                   vStatus = "P"
247               Else
248                   vStatus = "S"
249               End If
                  
250              Set rsSearch = objO.FetchDataACC("select KodeLgn from ARMST WITH(NOLOCK) where KodeLgn = '" & NulSID(kodeLgn) & "'")
251              If rsSearch.EOF Then
                     '' insert data armst
252                  sSQL1 = "INSERT INTO [LOGACC]..ArmstLog (KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                     "        StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, TglEntry,FPGabungan,NonPPn,Afiliasi ,DBSumber,Keterangan,ClientName,KodeUserChange,TglChange,TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak, SIUP, No_Mou ) "
253                  sSQL1 = sSQL1 & " VALUES ( '" & NulSID(kodeLgn) & "', " & _
                     "                  '" & s2d(NulSID(namaLengkap)) & "', " & _
                     "                  '" & s2d(Left(NulSID(kodeLgn), 3)) & "', " & _
                     "                  '" & s2d(NulSID(txt(8))) & "', " & _
                     "                  '" & s2d(NulSID(txt(66)) + ", " + NulSID(txt(65))) & "', " & _
                     "                  '" & s2d(NulSID(txt(64)) + ", " + NulSID(txt(63))) & "', " & _
                     "                  '" & s2d(NulSID(txt(14))) & "', " & _
                     "                  '" & s2d(NulSID(txt(12)) & NulSID(txt(13))) & "', " & _
                     "                  '" & s2d(NulSID(txtNpwp)) & "', " & _
                     "                  'D0', 0, " & _
                     "                  '" & s2d(NulSID(txt(18))) & "', 'GRP', " & _
                     "                  '" & s2d(vStatus) & "', '', " & _
                     "                  " & "0" & ", 'Customer " & cboTipe.text & "', 0, " & _
                     "                  '" & gSA & "', GETDATE(),0,0,0, " & _
                     "                  '" & Trim(gDb) & "','Data Baru','" & CompName & "' ,'" & gSA & "',Getdate(), " & _
                     "                  '" & "D0" & "', '" & "D0" & "', '" & "D0" & "', '" & "D0" & "', " & _
                     "                  " & 0 & ", " & 0 & ", " & 0 & ", " & 0 & ", " & _
                     "                  " & 0 & ", '" & 0 & "', '', '', '' ) "   'A.Rev1711291-SingleId

254                  sSQL = "INSERT INTO ARMST(KodeLgn,NamaLgn,KodeWil,Alamat1,Alamat2,Alamat3" _
                          & " ,Telepon,Nofax,Npwp,SyaratBayar,TOPPart, TOPGRP, TOPBP, TOPVHC,StatusPersh,KodeUser,TglEntry,HubungDengan,Telepon1)" _
                         & " VALUES ('" & NulSID(kodeLgn) & "','" & NulSID(namaLengkap) & "','" & Left(NulSID(kodeLgn), 3) & "'" _
                         & ",'" & s2d(NulSID(txt(8))) & "','" & s2d(NulSID(txt(66)) + ", " + NulSID(txt(65))) & "','" & s2d(NulSID(txt(64)) + ", " + NulSID(txt(63))) & "','" & Replace(Replace(NulSID(txt(12) & txt(13)) & "," & NulSID(txt(14)) & "," & NulSID(txt(15)), ">", ""), " ", "") & "','" & NulSID(txt(17)) & "'" _
                         & ",'" & NulSID(txtNpwp) & "','D0','D0','D0','D0','D0','" & vStatus & "','" & gSA & "','" & Format(isnow, "MM/DD/YYYY") & "','" & NulSID(txt(18)) & "','" & NulSID(txt(19)) & "')"
                     
255                   Call objHT.InsertData_ACC(sSQL1)
256                   Call objHT.InsertData_ACC(sSQL)
                       
257              Else
258                  If isBolehArmst = True Then
259                      sSQL1 = "INSERT INTO [LOGACC]..ArmstLog (KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                                                           "StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, TglEntry,FPGabungan,NonPPn,Afiliasi ,No_MOU,DBSumber,Keterangan,ClientName, " & _
                                                                           "KodeUserChange,TglChange,TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak,SIUP ) " & _
                                           "SELECT KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Nofax, Npwp, SyaratBayar, BatasKredit, HubungDengan, KodeSales, " & _
                                                                           "StatusPersh, KodeGl, NonAktifFlag, Memo,StatusMobil, KodeUser, GETDATE(),FPGabungan,NonPPn,Afiliasi,No_MOU, '" & Trim(gDb) & "','Sebelum Diedit','" & CompName & "' ,'" & gSA & "'," & _
                                           "Getdate(),TOPPart,TOPGRP,TOPBP,TOPVhc,BatasPart,BatasGRP,BatasBP,BatasVhc,PartshopFlag,Typecust,FakturPajak,SIUP from Armst WHERE KodeLgn='" & NulSID(kodeLgn) & "' "
         
                          
260                      sSQL = "UPDATE ARMST SET NamaLgn = '" & NulSID(namaLengkap) & "'," _
                              & " KodeWil = '" & Left(NulSID(kodeLgn), 3) & "',Alamat1 = '" & s2d(NulSID(txt(8))) & "',Alamat2 = '" & s2d(NulSID(txt(66)) + ", " + NulSID(txt(65))) & "',Alamat3 = '" & s2d(NulSID(txt(64)) + ", " + NulSID(txt(63))) & "',TglEntry = '" & Format(isnow, "MM/DD/YYYY") & "'," _
                             & " Telepon = '" & Replace(Replace(NulSID(txt(12) & txt(13)) & "," & NulSID(txt(14)) & "," & NulSID(txt(15)), ">", ""), " ", "") & "',Npwp = '" & NulSID(txtNpwp) & "'," _
                             & " TOPGRP = 'D0',StatusPersh = '" & vStatus & "', KodeUser='" & gSA & "',Nofax = '" & NulSID(txt(17)) & "',HubungDengan = '" & NulSID(txt(18)) & "',Telepon1 = '" & NulSID(txt(19)) & "'" _
                             & " where KodeLgn = '" & NulSID(kodeLgn) & "'"
261                      sSQL = "UPDATE ARMST SET NamaLgn = '" & NulSID(namaLengkap) & "'," _
                              & " KodeWil = '" & Left(NulSID(kodeLgn), 3) & "',Alamat1 = '" & s2d(NulSID(txt(8))) & "',Alamat2 = '" & s2d(NulSID(txt(66)) + ", " + NulSID(txt(65))) & "',Alamat3 = '" & s2d(NulSID(txt(64)) + ", " + NulSID(txt(63))) & "',TglEntry = '" & Format(isnow, "MM/DD/YYYY") & "'," _
                             & " Telepon = '" & Replace(Replace(NulSID(txt(12) & txt(13)) & "," & NulSID(txt(14)) & "," & NulSID(txt(15)), ">", ""), " ", "") & "',Npwp = '" & NulSID(txtNpwp) & "'," _
                             & " StatusPersh = '" & vStatus & "', KodeUser='" & gSA & "',Nofax = '" & NulSID(txt(17)) & "',HubungDengan = '" & NulSID(txt(18)) & "',Telepon1 = '" & NulSID(txt(19)) & "'" _
                             & " where KodeLgn = '" & NulSID(kodeLgn) & "'"
262                       Call objHT.InsertData_ACC(sSQL1)
263                       Call objHT.InsertData_ACC(sSQL)
                           
264                  End If
265              End If
                 
                  
266               Set rsSearch = objHT.FetchDataBengkel("select IDCust from SINGLEIDORI..NEWCUSTOMER WITH(NOLOCK) where IDCust = '" & NulSID(txt(1)) & "'")
267               If rsSearch.EOF Then
268                   sSQL = "INSERT INTO SINGLEIDORI..NEWCUSTOMER (IDCust,TypeCustomer,NoKTP,noNPWP,Title,NamaCust1,NamaCust2,NamaCust3,tglLahir,Agama,Pekerjaan,Propinsi,Kota,Kecamatan,Kelurahan,KodePos,alamat,RT,RW,IdTelp,telp,HP1,HP2,Email,Hobi,Olahraga,Suku,Preferensi,FPGab,NonPPn,TopVhc,BatasVhc,TopGrp,BatasGrp,TopBp,BatasBp,TopPart,BatasPart,TglEntry,UserEntry,NoFax,namaPIC,telpPIC,kewarganegaraan)"
269                   sSQL = sSQL & " VALUES ('" & NulSID(txt(1)) & "','" & UCase(NulSID(cboTipe.text)) & "','" & UCase(NulSID(txt(2))) & "','" & NulSID(txtNpwp) & "','" & UCase(NulSID(cboTitle.text)) & "','" & UCase(NulSID(txt(4))) & "','" & UCase(NulSID(txt(5))) & "','" & UCase(NulSID(txt(6))) & "','" & Format(DtLahir.value, "mm/dd/yyyy") & "','" & UCase(txt(20).text) & "','" & IIf(UCase(NulSID(txt(7))) = "LAINNYA", UCase(NulSID(txt(23))), UCase(NulSID(txt(7)))) & "','" & provinsi & "','" & Kota & "','" & kec & "','" & kel & "',"
270                   sSQL = sSQL & " '" & UCase(NulSID(txt(11))) & "','" & UCase(NulSID(txt(8))) & "','" & UCase(NulSID(txt(9))) & "','" & UCase(NulSID(txt(10))) & "','" & UCase(NulSID(txt(12))) & "','" & UCase(NulSID(txt(13))) & "','" & UCase(NulSID(txt(14))) & "','" & UCase(NulSID(txt(15))) & "','" & UCase(NulSID(txt(16))) & "',"
271                   sSQL = sSQL & " '" & UCase(NulSID(txt(21).text)) & "','" & UCase(NulSID(txt(22).text)) & "','" & UCase(NulSID(txt(71))) & "','" & UCase(NulSID(txt(72))) & "',0,0,'D0',0,'D0',0,'D0',0,'D0',0,GETDATE(),'" & gSA & "','" & NulSID(txt(17)) & "','" & NulSID(txt(18)) & "','" & NulSID(txt(19)) & "','" & NulSID(cboKewarganegaraan.text) & "')"
272               Else
273                   sSQL = "UPDATE SINGLEIDORI..NEWCUSTOMER SET NoKTP = '" & UCase(NulSID(txt(2))) & "',noNPWP = '" & UCase(NulSID(txtNpwp)) & "',Title = '" & UCase(cboTitle.text) & "',NamaCust1='" & UCase(NulSID(txt(4))) & "',NamaCust2='" & UCase(NulSID(txt(5))) & "',NamaCust3='" & UCase(NulSID(txt(6))) & "',tglLahir='" & Format(DtLahir.value, "mm/dd/yyyy") & "',Agama='" & UCase(txt(20).text) & "',Pekerjaan='" & IIf(UCase(NulSID(txt(7))) = "LAINNYA", UCase(NulSID(txt(23))), UCase(NulSID(txt(7)))) & "',Propinsi='" & provinsi & "',Kota='" & Kota & "',Kecamatan='" & kec & "',"
274                   sSQL = sSQL & " Kelurahan='" & kel & "',KodePos='" & UCase(NulSID(txt(11))) & "',alamat='" & UCase(NulSID(txt(8))) & "',RT='" & UCase(NulSID(txt(9))) & "',RW='" & UCase(NulSID(txt(10))) & "',IdTelp='" & UCase(NulSID(txt(12))) & "',telp='" & UCase(NulSID(txt(13))) & "',HP1='" & UCase(NulSID(txt(14))) & "',HP2='" & UCase(NulSID(txt(15))) & "',Email='" & UCase(NulSID(txt(16))) & "',Hobi='" & UCase(NulSID(txt(21).text)) & "',Olahraga='" & UCase(NulSID(txt(22).text)) & "',Suku='" & UCase(NulSID(txt(71))) & "',Preferensi='" & UCase(NulSID(txt(72))) & "',NoFax = '" & NulSID(txt(17)) & "',namaPIC= '" & NulSID(txt(18)) & "',telpPIC= '" & NulSID(txt(19)) & "',kewarganegaraan = '" & NulSID(cboKewarganegaraan.text) & "',UserUpdate = '" & NulSID(gSA) & "',TglUpdate = GETDATE(),DealerUpdate = '" & Right(gDb, 2) & "', AppUpdate = 'SVC' WHERE IDCust = '" & NulSID(txt(1)) & "' "
275               End If
                  
276               Call objHT.InsertData_Bengkel(sSQL)
                  
277               Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM SINGLEIDORI..Bridging WHERE idcust = '" & NulSID(txt(1)) & "' and kodedealer = '" & Right(gDb, 2) & "' ")
278               If rsSearch.EOF Then
279                   sSQL = "INSERT INTO SINGLEIDORI..Bridging(KodeDealer,kodeLgn,IDCust,TglEntry,UserEntry) VALUES ('" & Right(gDb, 2) & "','" & kodeLgnLama & "','" & NulSID(txt(1)) & "',getdate(),'" & gSA & "')"
280                   Call objHT.InsertData_Bengkel(sSQL)
281               End If
                  
                  
                  
282               Call objHT.CloseTransactionAll
                  
283               gBantuan = Trim(txt(1))
      '            Call hapusForm
      '            Text(57) = ""
                  
284               With frmSIDCustomer
285                   If .Visible Then
286                       .text(2) = Trim(gBantuan)
287                       Unload Me
288                   End If
289               End With
                  
290               With frmBooking
291                   If .Visible Then
292                       .text(2) = Trim(gBantuan)
293                       Unload Me
294                   End If
295               End With
296               MsgBox "Data berhasil disimpan", vbInformation
297               Unload Me
              'End If
298       Else
299           MsgBox "Silahkan lengkapi data berikut:" & vbCrLf & lengkapi, vbInformation
300           Exit Sub
301       End If

302      On Error GoTo 0
303      Exit Sub
          

304      On Error GoTo 0
305      Exit Sub

cmdSimpan_Click_Error:
306       If isBegin Then
307           Call objHT.CloseTransactionAll_Err
308       End If
309       MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmSIDMasterPelangganNonMerge"
End Sub

Private Sub cmdUploadKTP_Click()
    Dim PictBmp As String
    CommonDialog1.FileName = ""
    CommonDialog1.filter = "Pictures (*.jpg;*.jpeg;)|*.jpg;*.jpeg;"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.ShowOpen
    PictBmp = CommonDialog1.FileName
    txtUploadKTP.text = PictBmp
    If PictBmp = "" Then Exit Sub
    If FileLen(PictBmp) > 500000 Then
        MsgBox "Ukuran foto terlalu besar, Ukuran foto maksimal 500 KB"
        Close #1
        Exit Sub
    End If
    If CommonDialog1.FileName = "" Then Exit Sub
     
    Image1.Picture = LoadPicture(CommonDialog1.FileName)
 

End Sub

Private Sub cmdUploadSusulan_Click()
    Dim PictBmp As String
    CommonDialog1.FileName = ""
    CommonDialog1.filter = "Pictures (*.jpg;*.jpeg;)|*.jpg;*.jpeg;"
    CommonDialog1.Flags = cdlOFNHideReadOnly
    CommonDialog1.ShowOpen
    PictBmp = CommonDialog1.FileName
    txtUploadKTPSusulan.text = PictBmp
    If PictBmp = "" Then Exit Sub
    If FileLen(PictBmp) > 500000 Then
        MsgBox "Ukuran foto terlalu besar, Ukuran foto maksimal 500 KB"
        Close #1
        Exit Sub
    End If
    If CommonDialog1.FileName = "" Then Exit Sub
     
    Image2.Picture = LoadPicture(CommonDialog1.FileName)
 
End Sub

Private Sub DTLahir_GotFocus()
    If DtLahir.BackColor = fColor Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Activate()
    If isEditData Then
        
        
        txt(4).BackColor = &H80FFFF
        txt(5).BackColor = &H80FFFF
        txt(6).BackColor = &H80FFFF
        
        txt(14).BackColor = &H80FFFF
        
        
        cboTipe.Enabled = False
        txt(1).Enabled = False
        If NulSID(txt(1).text) <> "" And isLoad = False Then
            Call callPanggilDataNewCustomer
        End If
    End If
End Sub

Private Sub Form_Load()
1         isLoad = False
2         isBolehArmst = True
3         Me.Caption = "" 'TitleBP
4         Move 0, objA.MyTop + 370
5         Me.Width = objA.MyLeft
         ' Me.Height = fHeight - 260
6        gMainKeyCari = ""
          'Frame1.Left = (((objA.MyLeft - 480) - Val(Frame1.Width)) / 2) + 495
7         Call SetControlForm("load")
          'Call ViewPassing
8         Call Load_Cbo
9         FrAlm.Visible = True
          '20180313 REV'LONCAT JIKA KODELGN SUDAH ADA D ACC
10        iKodelgn = 1
11        If isEditData Then
              
              
12            txt(4).BackColor = &H80FFFF '&HC0E0FF
13            txt(5).BackColor = &H80FFFF ' &HC0E0FF
14            txt(6).BackColor = &H80FFFF ' &HC0E0FF
              
15            cboTipe.Enabled = False
16            txt(1).Enabled = False
17            If NulSID(txt(1).text) <> "" And isLoad = False Then
18                Call callPanggilDataNewCustomer
19            End If
20        End If
          'isLoad = True
End Sub

Private Sub callPanggilDataNewCustomer()
          Dim rsCust As ADODB.Recordset
          Dim rsHelp As ADODB.Recordset
1         If NulSID(noRangka) <> "" Then
2             sSQL = "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE no_chassis = '" & NulSID(noRangka) & "'"
3             Set rsCust = objO.FetchData(sSQL)
4             If Not rsCust.EOF Then
5                 kodeLgnLama = ""
6                 kodeLgnLama = NulSID(rsCust!FS_KD_PLG)
7                 sSQL = "SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE KodeDealer = '" & Right(gDb, 2) & "' AND IDCUST= '" & NulSID(txt(1).text) & "' AND KODELGN = '" & NulSID(kodeLgnLama) & "'"
8                 Set rsCust = objO.FetchData(sSQL)
9                 If rsCust.EOF Then
10                    isBolehArmst = False
11                End If
12            Else
13                isBolehArmst = False
14            End If
15        End If
          
16        sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST= '" & NulSID(txt(1).text) & "'"
17        Set rsCust = objO.FetchData(sSQL)
18        If Not rsCust.EOF Then
19            txt(1) = NulSID(rsCust!idcust)
20            txt(2) = NulSID(rsCust!noKTP)
21            If NulSID(rsCust!noNPWP) <> "" Then
                  'txtNpwp.Mask = ""
22                txtNpwp.text = setNpwpSID(Replace(Replace(NulSID(rsCust!noNPWP), ".", ""), "-", ""))
                  'txtNpwp.Mask = "00.000.000.0-000.000" '"##.###.###.#-###.###"
      '        Else
      '        txtNpwp.Text = ""
23            Else
24                txtNpwp.text = "000.000.000.0-000.000"
25            End If
26            txt(4) = NulSID(rsCust!namacust1)
27            txt(5) = NulSID(rsCust!namacust2)
28            txt(6) = NulSID(rsCust!namacust3)
              
29            txt(4).Enabled = False
30            txt(5).Enabled = False
31            txt(6).Enabled = False
              
32            DtLahir.value = NulSID(rsCust!tglLahir)
              
33            Set rsHelp = objO.FetchData("SELECT DISTINCT agama FROM [SingleIdOri]..agama WITH (NOLOCK) WHERE agama = '" & NulSID(rsCust!Agama) & "'")
34            If Not rsHelp.EOF Then
                  'cboAgama.Text = nulsid(rsCust!Agama)
35                txt(20).text = NulSID(rsCust!Agama)
36            End If
37            cboKewarganegaraan.Clear
38            cboKewarganegaraan.AddItem ("WNI")
39            cboKewarganegaraan.AddItem ("WNA")
              
40            If NulSID(rsCust!kewarganegaraan) <> "" Then
41                cboKewarganegaraan.text = NulSID(rsCust!kewarganegaraan)
42            End If
43            cboTipe.text = NulSID(rsCust!TypeCustomer)
44            Call cboTipe_Change
              
45            If UCase(cboTipe.text) <> "PERORANGAN" And NulSID(cboTipe.text) <> "" Then
46                txt(18).Enabled = True
47                txt(19).Enabled = True
                  
                  ''NPWP 16 Digit
48                txtNpwp.Enabled = True
49                txt(2).Enabled = False
                  
50            Else
                  ''NPWP 16 Digit
51                txtNpwp.Enabled = False
52                txt(2).Enabled = True
                  
53                txt(18).text = ""
54                txt(19).text = ""
55                txt(18).Enabled = False
56                txt(19).Enabled = False
57                txt(18).BackColor = fColor
58                txt(19).BackColor = fColor
59            End If
              
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE PROPINSI = '" & nulsid(rsCust!Propinsi) & "'")
      '        If Not rsHelp.EOF Then
      '            cboProv.text = UCase(nulsid(rsCust!Propinsi))
      '            Call cboProv_Change
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE kota = '" & nulsid(rsCust!Kota) & "'")
      '            If Not rsHelp.EOF Then
      '                cboKota.text = UCase(nulsid(rsCust!Kota))
      '                Call cboKota_Change
      '            End If
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE kota = '" & nulsid(rsCust!Kota) & "' and KECAMATAN = '" & nulsid(rsCust!Kecamatan) & "'")
      '            If Not rsHelp.EOF Then
      '                cboKec.text = UCase(nulsid(rsCust!Kecamatan))
      '                Call cboKec_Change
      '            End If
      '
      '            Set rsHelp = objO.FetchData("SELECT DISTINCT PROPINSI FROM [SingleIdOri]..TopologyDaerah WHERE KECAMATAN = '" & nulsid(rsCust!Kecamatan) & "' and kelurahan = '" & nulsid(rsCust!kelurahan) & "'")
      '            If Not rsHelp.EOF Then
      '                cboKel.text = (nulsid(rsCust!kelurahan))
      '            Call cboKel_Change
      '            End If
      '
      '        Else
      '            'Call hideComboProv(False)]
      '            FrAlm.Visible = False
60                cmdBack.Visible = False
61                txt(63) = UCase(NulSID(rsCust!Propinsi))
62                txt(64) = UCase(NulSID(rsCust!Kota))
63                txt(65) = UCase(NulSID(rsCust!kecamatan))
64                txt(66) = UCase(NulSID(rsCust!kelurahan))
      '        End If
               
65            txt(8) = NulSID(rsCust!Alamat)
66            txt(9) = Format(NulSID(rsCust!rt), "0##") 'nulsid(rsCust!RT)
67            txt(10) = Format(NulSID(rsCust!rw), "0##")
68            txt(11) = NulSID(rsCust!KodePos)
              
69            txt(12) = NulSID(rsCust!IdTelp)
70            txt(13) = NulSID(rsCust!telp)
71            txt(14) = NulSID(rsCust!hp1)
72            txt(15) = NulSID(rsCust!hp2)
73            txt(16) = NulSID(rsCust!email)
              
74            Set rsHelp = objO.FetchData("SELECT DISTINCT Pekerjaan FROM [SingleIdOri]..Pekerjaan WITH (NOLOCK) WHERE Pekerjaan = '" & NulSID(rsCust!pekerjaan) & "'")
75            If Not rsHelp.EOF Then
76                txt(7) = UCase(NulSID(rsCust!pekerjaan))

77            Else
78                txt(7) = "LAINNYA"
                 ' Call cboPek_Change
                 ' txt(7) = UCase(nulsid(rsCust!Pekerjaan))
79                txt(23) = UCase(NulSID(rsCust!pekerjaan))
80                txt(23).BackColor = &HFFFFFF
81            End If
              
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT hobi FROM [SingleIdOri]..hobi WHERE hobi = '" & nulsid(rsCust!Hobi) & "'")
      '        If Not rsHelp.EOF Then
      '            cboHobi.Text = UCase(nulsid(rsCust!Hobi))
      '        End If
              
82            txt(21) = UCase(NulSID(rsCust!hobi))
83            txt(22) = UCase(NulSID(rsCust!Olahraga))
              
      '        Set rsHelp = objO.FetchData("SELECT DISTINCT olahraga FROM [SingleIdOri]..olahraga WHERE olahraga = '" & nulsid(rsCust!olahraga) & "'")
      '        If Not rsHelp.EOF Then
      '            cboOlah.Text = UCase(nulsid(rsCust!olahraga))
      '        End If

84            Set rsHelp = objO.FetchData("SELECT DISTINCT Title FROM [SingleIdOri]..Title WITH (NOLOCK) WHERE Title = '" & NulSID(rsCust!Title) & "'")
85            If Not rsHelp.EOF Then
86                cboTitle.text = UCase(NulSID(rsCust!Title))
87            End If
              
88            txt(71) = UCase(NulSID(rsCust!Suku))
89            txt(72) = UCase(NulSID(rsCust!Preferensi))
              
90            txt(17) = UCase(NulSID(rsCust!NoFax))
91            txt(18) = UCase(NulSID(rsCust!NamaPIC))
92            txt(19) = UCase(NulSID(rsCust!TelpPIC))
              
              '' 20181011REV1 | REV | PK M.WASISTO | PENAMBAHAN SCORING SINGLE ID
93            If Log.singleIDScoring Then
94                isWajibKTP = False
95                If NulSID(rsCust!PinBB) = "SCORING" Then
96                    isWajibKTP = True
97                End If
98            End If
              ''------------------------------------------------------------------------------------------
              '' 20181011REV1 | REV | PK M.WASISTO | PENAMBAHAN SCORING SINGLE ID                        -
              ''------------------------------------------------------------------------------------------
              ''  - EDIT CUSTOMER
              ''      = JIKA ADA OUTSTANDING WO ATAS USER TERSEBUT MAKA BISA EDIT KTP + UPLOAD KTP
              ''      = JIKA TIDAK ADA OUTSTANDING WO MAKA HANYA BISA EDIT KTP
              ''------------------------------------------------------------------------------------------
99            If Log.singleIDScoring = 1 And UCase(cboTipe.text) = "PERORANGAN" Then
100               lblFotoKTPCustomer.Visible = True
101               Image3.Visible = True
                  
                  '' =====================================================
                  '' 20181016 | menampilkan foto ktp jika sudah diupload
                  '' =====================================================
102               sSQL = "SELECT TOP 1 * FROM [singleidori]..UploadKTP WITH (NOLOCK) WHERE IDCUST = '" & NulSID(txt(1).text) & "' ORDER BY TGLENTRY DESC "
103               Set rsHelp = objO.FetchData(sSQL)
104               If Not rsHelp.EOF Then
                      Dim fotoStream      As New ADODB.Stream
105                   Set fotoStream = New ADODB.Stream
106                   With fotoStream
107                       .Type = adTypeBinary
108                       .Open
109                       .Write rsHelp.Fields("pic").value 'rsTemp.Fields("pic").VALUE
                      
110                       .SaveToFile App.Path & "\fotoKTP.bmp", adSaveCreateOverWrite
                          
111                       Image3.Picture = LoadPicture(App.Path & "\fotoKTP.bmp")
112                   End With
113               Else
114                   Image3.Picture = Nothing
115               End If
                  
116               sSQL = "SELECT a.NOINV,a.no_wo,c.*"
117               sSQL = sSQL & " FROM [SingleIDOri]..NewCustomer c WITH (NOLOCK)"
118               sSQL = sSQL & " inner join [SingleIDOri]..Bridging b WITH (NOLOCK) on c.IDCust = b.IDCust and b.KodeDealer = '" & Right(gDb, 2) & "'"
119               sSQL = sSQL & " left join JOBDTA a WITH (NOLOCK) on a.FS_KD_PLG = b.KodeLgn"
120               sSQL = sSQL & " where c.IDCust = '" & NulSID(txt(1).text) & "' and ((a.NOINV = '' and a.FS_LOG_DEL = '')"
121               sSQL = sSQL & " or c.PinBB = 'SCORING')"
122               Set rsHelp = objO.FetchData(sSQL)
123               If Not rsHelp.EOF Then
124                  cmdUploadSusulan.Visible = True
125                  txtUploadKTPSusulan.Visible = True
126                  Image2.Visible = True
127               Else
128                  cmdUploadSusulan.Visible = False
129                  txtUploadKTPSusulan.Visible = False
130                  Image2.Visible = False
131               End If
132           End If
133            isLoad = True
134       End If
End Sub

Private Sub SetControlForm(nSET As String)
On Error GoTo ErrNull
    With Me
    Select Case nSET
        Case "load"
            '.txt(1).BackColor = &HC0E0FF
            .txt(1).text = ""
            .txt(1).Enabled = False
            For i = 2 To 72
                '.txt(i).BackColor = tColor
                .txt(i).text = ""
            Next i
            
            .cboAgama.Clear
            .cboPek.Clear
            .cboHobi.Clear
            .cboOlah.Clear
            
            .cboProv.Clear
            .cboKota.Clear
            .cboKec.Clear
            .cboKel.Clear
            
            '.cboProvP.Clear
'            .cboKotaP.Clear
'            .cboKecP.Clear
'            .cboKelP.Clear
            
            .cboAgamaS.Clear
            .cboPekS.Clear
            .cboHobiS.Clear
            .cboOlahS.Clear
            
            .cboProvS.Clear
            .cboKotaS.Clear
            .cboKecS.Clear
            .cboKelS.Clear
            
            .cboTitle.BackColor = fColor
            .cboTipe.BackColor = fColor
            
            .cboAgama.BackColor = fColor
            .cboPek.BackColor = fColor
            .cboHobi.BackColor = fColor
            .cboOlah.BackColor = fColor
            
            .cboProv.BackColor = fColor
            .cboKota.BackColor = fColor
            .cboKec.BackColor = fColor
            .cboKel.BackColor = fColor
            
'            .cboProvP.BackColor = fColor
'            .cboKotaP.BackColor = fColor
'            .cboKecP.BackColor = fColor
'            .cboKelP.BackColor = fColor
            
            .cboTitleS.BackColor = fColor
           ' txt(7).BackColor = fColor
            .cboAgamaS.BackColor = fColor
            .cboPekS.BackColor = fColor
            .cboHobiS.BackColor = fColor
            
            .cboProvS.BackColor = fColor
            .cboKotaS.BackColor = fColor
            .cboKecS.BackColor = fColor
            .cboKelS.BackColor = fColor
            .DtLahir.BackColor = fColor
'            .DtDdate.BackColor = fColor
'            .DtPolish0.BackColor = fColor
'            .DtPolish1.BackColor = fColor
            .DtLahirS.BackColor = fColor
            .ChSurat.ForeColor = fColorz
'            .chPolish.ForeColor = fColorz
'            .chPolish.value = 0
            .ChSurat.value = 0
            .DtLahir.text = "__/__/____"
'            .DtDdate.text = "__/__/____"
'            .DtPolish0.text = "__/__/____"
'            .DtPolish1.text = "__/__/____"
            .DtLahirS.text = "__/__/____"
            .cmdSave.Visible = True
            .cmdCreate.Visible = False
            .cmdAdd.Visible = True
            .cmdAdd.Enabled = True
            .cmdCreate.Enabled = False
            .cmdEdit.Enabled = False
            .cmdSave.Enabled = False
            .cmdDelete.Enabled = False
            .cmdCancel.Enabled = False
            .cmdOpen.Enabled = False
            .cmdExit.Enabled = True
            
        Case "load_update"
            .txt(1).BackColor = tColor
            For i = 2 To 43
                .txt(i).BackColor = fColor
            Next i
'            .DtDdate.BackColor = fColor
            'SID .DtDOTam.BackColor = fColor
'            .DtPolish0.BackColor = fColor
'            .DtPolish1.BackColor = fColor
'            .chPolish.ForeColor = fColorz
            .ChSurat.ForeColor = fColorz
            'SID .optTipe(0).ForeColor = fColorz
            'SID .optTipe(1).ForeColor = fColorz
            
            .cmdSave.Visible = True
            .cmdCreate.Visible = False
            .cmdEdit.Visible = True
            .cmdAdd.Visible = False
            .cmdCreate.Enabled = False
            .cmdEdit.Enabled = True
            .cmdSave.Enabled = False
            .cmdDelete.Enabled = True
            .cmdCancel.Enabled = True
            .cmdOpen.Enabled = True
            .cmdExit.Enabled = True
            
        Case "add"
            Call Load_Cbo
            For i = 1 To 74
                txt(i).BackColor = tColorx              'Kuning
                Select Case i
                    Case 2, 3, 15, 16, 30, 37, 38, 46, 47, 48, 49, 53, 61, 62, 71, 72, 73, 74
                        txt(i).BackColor = tColor       'Putih
                    Case 1, 7, 11, 22, 23, 25, 29, 31, 32, 33, 34, 35, 36, 42, 57
                        txt(i).BackColor = fColor       'Orange
                End Select
                txt(i).text = ""
            Next i
            
            If txt(31).text <> "" Then
'                .DtPolish0.BackColor = tColorx
'                .DtPolish0.text = "__/__/____"
'                .DtPolish1.BackColor = tColorx
'                .DtPolish1.text = "__/__/____"
            Else
'                .DtPolish0.BackColor = tColor
'                .DtPolish0.text = "__/__/____"
'                .DtPolish1.BackColor = tColor
'                .DtPolish1.text = "__/__/____"
            End If
            
            .cboTitle.BackColor = tColor
            .cboTipe.BackColor = tColor
            .cboAgama.BackColor = tColor
            .cboPek.BackColor = tColor
            .cboHobi.BackColor = tColor
            .cboProv.BackColor = tColor
            .cboKota.BackColor = tColor
            .cboKec.BackColor = tColor
            .cboKel.BackColor = tColor
'            .cboProvP.BackColor = tColor
'            .cboKotaP.BackColor = tColor
'            .cboKecP.BackColor = tColor
'            .cboKelP.BackColor = tColor
            .cboTitleS.BackColor = tColor
            .cboAgamaS.BackColor = tColor
            .cboPekS.BackColor = tColor
            .cboHobiS.BackColor = tColor
            .cboProvS.BackColor = tColor
            .cboKotaS.BackColor = tColor
            .cboKecS.BackColor = tColor
            .cboKelS.BackColor = tColor
            .DtLahir.BackColor = tColor
'            .DtDdate.BackColor = fColor
'            .DtPolish0.BackColor = tColor
'            .DtPolish1.BackColor = tColor
            .DtLahirS.BackColor = tColor
            .ChSurat.ForeColor = tColorz
'            .chPolish.ForeColor = tColorz
'            .chPolish.value = 0
            .ChSurat.value = 0
            .DtLahir.text = "__/__/____"
'            .DtDdate.text = "__/__/____"
'            .DtPolish0.text = "__/__/____"
'            .DtPolish1.text = "__/__/____"
            .DtLahirS.text = "__/__/____"
            
            .cmdSave.Visible = False
            .cmdCreate.Visible = True
            .cmdAdd.Enabled = False
            .cmdCreate.Enabled = True
            .cmdEdit.Enabled = False
            .cmdSave.Enabled = False
            .cmdCancel.Enabled = True
            .cmdDelete.Enabled = False
            .cmdOpen.Enabled = False
            .cmdExit.Enabled = False
            .cboTipe.SetFocus
            
        Case "edit"
            txt(1).BackColor = fColor
            If xWAC = True Then txt(1).BackColor = tColor
            For i = 2 To 4
                txt(i).BackColor = tColorx
            Next i
            For i = 5 To 6
                txt(i).BackColor = tColor
            Next i
            For i = 7 To 14
                txt(i).BackColor = tColorx
                If i = 9 Then txt(9).BackColor = tColor
            Next i
            For i = 17 To 21
                If i <> 18 Then txt(i).BackColor = tColorx
            Next i
            If Left(UCase(txt(14).text), 4) = "ZOTH" Then
                txt(16).BackColor = tColor
            Else
                txt(16).BackColor = fColor
            End If
            txt(19).BackColor = fColor
            txt(23).BackColor = tColor
            If txt(24).text <> "" Then
                For i = 31 To 33
                    txt(i).BackColor = tColorx
                Next i
                For i = 34 To 35
                    txt(i).BackColor = tColor
                Next i
                For i = 36 To 43
                    txt(i).BackColor = tColorx
                    If i = 40 Or i = 41 Then txt(i).BackColor = tColor
                Next i
'                .DtPolish0.BackColor = tColorx
'                .DtPolish1.BackColor = tColorx
            Else
                For i = 31 To 36
                    txt(i).BackColor = tColor
                Next i
                For i = 37 To 43
                    txt(i).BackColor = tColorx
                    If i = 40 Or i = 41 Then txt(i).BackColor = tColor
                Next i
'                .DtPolish0.BackColor = tColor
'                .DtPolish1.BackColor = tColor
            End If
            
'            .DtDdate.BackColor = tColor
            'SID .DtDOTam.BackColor = tColor
'            .chPolish.ForeColor = tColorz
            .ChSurat.ForeColor = tColorz
            'SID .optTipe(0).ForeColor = tColorz
            'SID .optTipe(1).ForeColor = tColorz
            
            .cmdSave.Visible = True
            .cmdCreate.Visible = False
            .cmdAdd.Enabled = False
            .cmdEdit.Enabled = False
            .cmdCreate.Enabled = False
            .cmdSave.Enabled = True
            .cmdCancel.Enabled = True
            .cmdDelete.Enabled = False
            .cmdOpen.Enabled = False
            .cmdExit.Enabled = False
    End Select
    End With
'txt(43).SetFocus
ErrNull:
End Sub

Private Sub Load_Cbo()

    cboKewarganegaraan.Clear
    cboKewarganegaraan.AddItem ("WNI")
    cboKewarganegaraan.AddItem ("WNA")
    Dim rsA As New ADODB.Recordset
    '===================================================
    '   TIPE
    '===================================================
    Set rsA = objSID.FetchDataSID("Select * from TipeCustomer With (NoLock) Order By Kode")
    With cboTipe
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Tipe")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   AGAMA
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Upper(Agama) as Agama from Agama With (NoLock)")
    With cboAgama
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Agama")
                rsA.MoveNext
            Loop
        End If
    End With
    With cboAgamaS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Agama")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   PEKERJAAN
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Upper(Pekerjaan) as Pekerjaan from Pekerjaan With (NoLock) Order By Pekerjaan")
    With cboPek
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Pekerjaan")
                rsA.MoveNext
            Loop
        End If
    End With
    With cboPekS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Pekerjaan")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   HOBI
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Upper(Hobi) as Hobi from Hobi With (NoLock) Order By Hobi")
    With cboHobi
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Hobi")
                rsA.MoveNext
            Loop
        End If
    End With
    With cboHobiS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Hobi")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   OLAHRAGA
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Upper(Olahraga) as Olahraga from Olahraga With (NoLock) Order By Olahraga")
    With cboOlah
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Olahraga")
                rsA.MoveNext
            Loop
        End If
    End With
    With cboOlahS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Olahraga")
                rsA.MoveNext
            Loop
        End If
    End With
    '===================================================
    '   PROVINSI
    '===================================================
    Set rsA = Nothing
    Set rsA = objSID.FetchDataSID("Select Distinct UPPER(Propinsi) AS Propinsi from TopologyDaerah With (NoLock) Order By Propinsi")
    With cboProv
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Propinsi")
                rsA.MoveNext
            Loop
        End If
        .AddItem "LAINNYA"
    End With

    With cboProvS
        .Clear
        If rsA.RecordCount > 0 Then
            rsA.MoveFirst
            Do Until rsA.EOF
                .AddItem rsA("Propinsi")
                rsA.MoveNext
            Loop
        End If
        .AddItem "LAINNYA"
    End With
    
'    With cboProvP
'        .Clear
'        If rsA.RecordCount > 0 Then
'            rsA.MoveFirst
'            Do Until rsA.EOF
'                .AddItem rsA("Propinsi")
'                rsA.MoveNext
'            Loop
'        End If
'        .AddItem "LAINNYA"
'    End With
End Sub

Private Sub cboAgama_Click()
    cboPek.SetFocus
End Sub

Private Sub cboAgama_GotFocus()
    If cboAgama.BackColor = fColor Then
        cboAgama.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboAgama.Locked = False
    End If
End Sub

Private Sub cboPek_Click()
    Call cboPek_Change
End Sub

Private Sub cboPek_Change()
    If UCase(cboPek.text) = "LAINNYA" Then
        txt(7).BackColor = tColor
        txt(7).SetFocus
    Else
        txt(7).BackColor = fColor
        Sendkeys ("{TAB}")
    End If
End Sub

Private Sub cboHobi_Click()
    Sendkeys ("{TAB}")
End Sub

Private Sub cboProv_Click()
    If cboProv.Tag = "" Then
        kliPr = True
        Call cboProv_Change
        cboKel.Clear
        cboKec.Clear
    End If
End Sub

Private Sub cboProv_Change()
    If kliPr = True Then
        If cboProv.text <> "" Then
            Dim rsX As New ADODB.Recordset, z As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProv.text)) & "' Order By Kota")
            With cboKota
                .Clear: z = 0: cboKota.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If z = 0 Then cboKota.Tag = rsX("Kota")
                        .AddItem rsX("Kota")
                        rsX.MoveNext
                        z = z + 1
                    Loop
                    kliKo = False
                    cboKota.text = cboKota.Tag
                    cboKota.Tag = ""
                    cboKota.SetFocus
                End If
            End With
        Else
            cboKota.Tag = "": cboKota.Clear
        End If
        If cboProv.text = "LAINNYA" Then
            FrAlm.Visible = True
            cmdBack.Visible = True
            txt(63).SetFocus
        Else
            FrAlm.Visible = False
            cmdBack.Visible = False
        End If
        kliPr = False
    End If
End Sub

Private Sub cboKota_Click()
    If cboKota.Tag = "" Then
        kliKo = True
        Call cboKota_Change
    End If
End Sub

Private Sub cboKota_Change()
    If kliKo = True Then
        If cboKota.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKota.text)) & "' Order By Kecamatan")
            With cboKec
                .Clear: X = 0: cboKota.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If X = 0 Then cboKec.Tag = rsX("Kecamatan")
                        .AddItem rsX("Kecamatan")
                        rsX.MoveNext
                        X = X + 1
                    Loop
                    cboKec.text = cboKec.Tag
                    cboKec.Tag = ""
                    cboKec.SetFocus
                End If
            End With
        Else
            cboKec.Tag = "": cboKec.Clear
        End If
    kliKo = False
    End If
End Sub

Private Sub cboKec_Click()
    If cboKec.Tag = "" Then
        kliKc = True
        Call cboKec_Change
    End If
End Sub

Private Sub cboKec_Change()
    If kliKc = True Then
        If cboKec.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKec.text)) & "' Order By Kelurahan")
            With cboKel
                .Clear: X = 0: cboKel.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If X = 0 Then cboKel.Tag = rsX("Kelurahan")
                        .AddItem rsX("Kelurahan")
                        rsX.MoveNext
                        X = X + 1
                    Loop
                    cboKel.text = cboKel.Tag
                    cboKel.Tag = ""
                    cboKel.SetFocus
                End If
            End With
        Else
            cboKel.Tag = "": cboKel.Clear
        End If
    kliKc = False
    End If
End Sub

Private Sub cboKel_Click()
    If cboKel.Tag = "" Then
        kliKl = True
        Call cboKel_Change
    End If
End Sub

Private Sub cboKel_Change()
    If kliKl = True Then
        If cboKel.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct IsNull(KodePos,'')as KodePos From TopologyDaerah Where UPPER(Kelurahan) = '" & Trim(UCase(cboKel.text)) & "'")
            If rsX.RecordCount > 0 Then
                rsX.MoveFirst
                txt(11).text = rsX("KodePos")
                txt(8).SetFocus
            Else
                txt(11).text = ""
                txt(8).SetFocus
            End If
        Else
            txt(11).text = ""
        End If
        kliKl = False
    End If
End Sub

Private Sub cboPek_GotFocus()
    If cboPek.BackColor = fColor Then
        cboPek.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboPek.Locked = False
    End If
End Sub

Private Sub cboHobi_GotFocus()
    If cboHobi.BackColor = fColor Then
        cboHobi.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboHobi.Locked = False
    End If
End Sub

Private Sub cboTipe_GotFocus()
    If cboTipe.BackColor = fColor Then
        cboTipe.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboTipe.Locked = False
    End If
End Sub

Private Sub cboTitle_GotFocus()
    If cboTitle.BackColor = fColor Then
        cboTitle.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboTitle.Locked = False
    End If
End Sub

Private Sub cboProv_GotFocus()
    If cboProv.BackColor = fColor Then
        cboProv.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboProv.Locked = False
    End If
End Sub

Private Sub cboKota_GotFocus()
    If cboKota.BackColor = fColor Then
        cboKota.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboKota.Locked = False
    End If
End Sub

Private Sub cboKec_GotFocus()
    If cboKec.BackColor = fColor Then
        cboKec.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboKec.Locked = False
    End If
End Sub

Private Sub cboKel_GotFocus()
    If cboKel.BackColor = fColor Then
        cboKel.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboKel.Locked = False
    End If
End Sub
'
'Private Sub cboProvP_Click()
'    If cboProvP.Tag = "" Then
'        kliPrp = True
'        Call cboProvP_Change
'        cboKelP.Clear
'        cboKecP.Clear
'    End If
'End Sub
'
'Private Sub cboProvP_Change()
'    If kliPrp = True Then
'        If cboProvP.Text <> "" Then
'            Dim rsX As New ADODB.Recordset, z As Integer
'            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvP.Text)) & "' Order By Kota")
'            With cboKotaP
'                .Clear: z = 0: cboKotaP.Tag = ""
'                If rsX.RecordCount > 0 Then
'                    rsX.MoveFirst
'                    Do Until rsX.EOF
'                        If z = 0 Then cboKotaP.Tag = rsX("Kota")
'                        .AddItem rsX("Kota")
'                        rsX.MoveNext
'                        z = z + 1
'                    Loop
'                    kliKop = False
'                    cboKotaP.Text = cboKotaP.Tag
'                    cboKotaP.Tag = ""
'                    cboKotaP.SetFocus
'                End If
'            End With
'        Else
'            cboKotaP.Tag = "": cboKotaP.Clear
'        End If
'        If cboProvP.Text = "LAINNYA" Then
'            FrAlmP.Visible = True
'            cmdBackP.Visible = True
'            txt(67).SetFocus
'        Else
'            FrAlmP.Visible = False
'            cmdBackP.Visible = False
'        End If
'        kliPrp = False
'    End If
'End Sub
'
'Private Sub cboKotaP_Click()
'    If cboKotaP.Tag = "" Then
'        kliKop = True
'        Call cboKotaP_Change
'    End If
'End Sub
'
'Private Sub cboKotaP_Change()
'    If kliKop = True Then
'        If cboKotaP.Text <> "" Then
'            Dim rsX As New ADODB.Recordset, X As Integer
'            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKotaP.Text)) & "' Order By Kecamatan")
'            With cboKecP
'                .Clear: X = 0: cboKotaP.Tag = ""
'                If rsX.RecordCount > 0 Then
'                    rsX.MoveFirst
'                    Do Until rsX.EOF
'                        If X = 0 Then cboKecP.Tag = rsX("Kecamatan")
'                        .AddItem rsX("Kecamatan")
'                        rsX.MoveNext
'                        X = X + 1
'                    Loop
'                    cboKecP.Text = cboKecP.Tag
'                    cboKecP.Tag = ""
'                    cboKecP.SetFocus
'                End If
'            End With
'        Else
'            cboKecP.Tag = "": cboKecP.Clear
'        End If
'    kliKop = False
'    End If
'End Sub
'
'Private Sub cboKecP_Click()
'    If cboKecP.Tag = "" Then
'        kliKcp = True
'        Call cboKecP_Change
'    End If
'End Sub

'Private Sub cboKecP_Change()
'    If kliKcp = True Then
'        If cboKecP.Text <> "" Then
'            Dim rsX As New ADODB.Recordset, X As Integer
'            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecP.Text)) & "' Order By Kelurahan")
'            With cboKelP
'                .Clear: X = 0: cboKelP.Tag = ""
'                If rsX.RecordCount > 0 Then
'                    rsX.MoveFirst
'                    Do Until rsX.EOF
'                        If X = 0 Then cboKelP.Tag = rsX("Kelurahan")
'                        .AddItem rsX("Kelurahan")
'                        rsX.MoveNext
'                        X = X + 1
'                    Loop
'                    cboKelP.Text = cboKelP.Tag
'                    cboKelP.Tag = ""
'                    cboKelP.SetFocus
'                End If
'            End With
'        Else
'            cboKelP.Tag = "": cboKelP.Clear
'        End If
'    kliKcp = False
'    End If
'End Sub

'Private Sub cboKelP_Click()
'    If cboKelP.Tag = "" Then
'        kliKlp = True
'        Call cboKelP_Change
'    End If
'End Sub
'
'Private Sub cboKelP_Change()
'    If kliKlp = True Then
'        If cboKelP.Text <> "" Then
'            Dim rsX As New ADODB.Recordset, X As Integer
'            Set rsX = objSID.FetchDataSID("Select Distinct IsNull(KodePos,'')as KodePos From TopologyDaerah Where UPPER(Kelurahan) = '" & Trim(UCase(cboKelP.Text)) & "'")
'            If rsX.RecordCount > 0 Then
'                rsX.MoveFirst
'                txt(42).Text = rsX("KodePos")
'                txt(39).SetFocus
'            Else
'                txt(42).Text = ""
'                txt(39).SetFocus
'            End If
'        Else
'            txt(42).Text = ""
'        End If
'        kliKlp = False
'    End If
'End Sub
'
'Private Sub cboProvP_GotFocus()
'    If cboProvP.BackColor = fColor Then
'        cboProvP.Locked = True
'        SendKeys ("{TAB}")
'    Else
'        cboProvP.Locked = False
'    End If
'End Sub
'
'Private Sub cboKotaP_GotFocus()
'    If cboKotaP.BackColor = fColor Then
'        cboKotaP.Locked = True
'        SendKeys ("{TAB}")
'    Else
'        cboKotaP.Locked = False
'    End If
'End Sub
'
'Private Sub cboKecP_GotFocus()
'    If cboKecP.BackColor = fColor Then
'        cboKecP.Locked = True
'        SendKeys ("{TAB}")
'    Else
'        cboKecP.Locked = False
'    End If
'End Sub
'
'Private Sub cboKelP_GotFocus()
'    If cboKelP.BackColor = fColor Then
'        cboKelP.Locked = True
'        SendKeys ("{TAB}")
'    Else
'        cboKelP.Locked = False
'    End If
'End Sub

Private Sub cboHobiS_Click()
    Sendkeys ("{TAB}")
End Sub

Private Sub cboProvS_Click()
    If cboProvS.Tag = "" Then
        kliPrs = True
        Call cboProvS_Change
        cboKelS.Clear
        cboKecS.Clear
    End If
End Sub

Private Sub cboProvS_Change()
    If kliPrs = True Then
        If cboProvS.text <> "" Then
            Dim rsX As New ADODB.Recordset, z As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvS.text)) & "' Order By Kota")
            With cboKotaS
                .Clear: z = 0: cboKotaS.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If z = 0 Then cboKotaS.Tag = rsX("Kota")
                        .AddItem rsX("Kota")
                        rsX.MoveNext
                        z = z + 1
                    Loop
                    kliKos = False
                    cboKotaS.text = cboKotaS.Tag
                    cboKotaS.Tag = ""
                    cboKotaS.SetFocus
                End If
            End With
        Else
            cboKotaS.Tag = "": cboKotaS.Clear
        End If
        If cboProvS.text = "LAINNYA" Then
            FrAlmS.Visible = True
            cmdBackS.Visible = True
            txt(67).SetFocus
        Else
            FrAlmS.Visible = False
            cmdBackS.Visible = False
        End If
        kliPrp = False
    End If
End Sub

Private Sub cboKotaS_Click()
    If cboKotaS.Tag = "" Then
        kliKos = True
        Call cboKotaS_Change
    End If
End Sub

Private Sub cboKotaS_Change()
    If kliKos = True Then
        If cboKotaS.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKotaS.text)) & "' Order By Kecamatan")
            With cboKecS
                .Clear: X = 0: cboKotaS.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If X = 0 Then cboKecS.Tag = rsX("Kecamatan")
                        .AddItem rsX("Kecamatan")
                        rsX.MoveNext
                        X = X + 1
                    Loop
                    cboKecS.text = cboKecS.Tag
                    cboKecS.Tag = ""
                    cboKecS.SetFocus
                End If
            End With
        Else
            cboKecS.Tag = "": cboKecS.Clear
        End If
    kliKos = False
    End If
End Sub

Private Sub cboKecS_Click()
    If cboKecS.Tag = "" Then
        kliKcs = True
        Call cboKecS_Change
    End If
End Sub

Private Sub cboKecS_Change()
    If kliKcs = True Then
        If cboKecS.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecS.text)) & "' Order By Kelurahan")
            With cboKelS
                .Clear: X = 0: cboKelS.Tag = ""
                If rsX.RecordCount > 0 Then
                    rsX.MoveFirst
                    Do Until rsX.EOF
                        If X = 0 Then cboKelS.Tag = rsX("Kelurahan")
                        .AddItem rsX("Kelurahan")
                        rsX.MoveNext
                        X = X + 1
                    Loop
                    cboKelS.text = cboKelS.Tag
                    cboKelS.Tag = ""
                    cboKelS.SetFocus
                End If
            End With
        Else
            cboKelS.Tag = "": cboKelS.Clear
        End If
    kliKcs = False
    End If
End Sub

Private Sub cboKelS_Click()
    If cboKelS.Tag = "" Then
        kliKls = True
        Call cboKelS_Change
    End If
End Sub

Private Sub cboKelS_Change()
    If kliKls = True Then
        If cboKelS.text <> "" Then
            Dim rsX As New ADODB.Recordset, X As Integer
            Set rsX = objSID.FetchDataSID("Select Distinct IsNull(KodePos,'')as KodePos From TopologyDaerah Where UPPER(Kelurahan) = '" & Trim(UCase(cboKelS.text)) & "'")
            If rsX.RecordCount > 0 Then
                rsX.MoveFirst
                txt(57).text = rsX("KodePos")
                txt(54).SetFocus
            Else
                txt(57).text = ""
                txt(54).SetFocus
            End If
        Else
            txt(57).text = ""
        End If
        kliKls = False
    End If
End Sub
'
'Private Sub LoadCbo_KotaP()
'    Dim rsX As New ADODB.Recordset, z As Integer
'    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvP.Text)) & "' Order By Kota")
'    With cboKotaP
'        .Clear: z = 0: cboKotaP.Tag = ""
'        If rsX.RecordCount > 0 Then
'            rsX.MoveFirst
'            Do Until rsX.EOF
'                If z = 0 Then cboKotaP.Tag = rsX("Kota")
'                .AddItem rsX("Kota")
'                rsX.MoveNext
'                z = z + 1
'            Loop
'        End If
'    End With
'End Sub
'
'Private Sub LoadCbo_KecP()
'    Dim rsX As New ADODB.Recordset, z As Integer
'    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKotaP.Text)) & "' Order By Kecamatan")
'    With cboKecP
'        .Clear: z = 0: cboKecP.Tag = ""
'        If rsX.RecordCount > 0 Then
'            rsX.MoveFirst
'            Do Until rsX.EOF
'                If z = 0 Then cboKecP.Tag = rsX("Kecamatan")
'                .AddItem rsX("Kecamatan")
'                rsX.MoveNext
'                z = z + 1
'            Loop
'        End If
'    End With
'End Sub
'
'Private Sub LoadCbo_KelP()
'    Dim rsX As New ADODB.Recordset, z As Integer
'    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecP.Text)) & "' Order By Kelurahan")
'    With cboKelP
'        .Clear: z = 0: cboKelP.Tag = ""
'        If rsX.RecordCount > 0 Then
'            rsX.MoveFirst
'            Do Until rsX.EOF
'                If z = 0 Then cboKelP.Tag = rsX("Kelurahan")
'                .AddItem rsX("Kelurahan")
'                rsX.MoveNext
'                z = z + 1
'            Loop
'        End If
'    End With
'End Sub

Private Sub LoadCbo_KotaS()
    Dim rsX As New ADODB.Recordset, z As Integer
    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kota,'')) as Kota From TopologyDaerah Where UPPER(Propinsi) = '" & Trim(UCase(cboProvS.text)) & "' Order By Kota")
    With cboKotaS
        .Clear: z = 0: cboKotaS.Tag = ""
        If rsX.RecordCount > 0 Then
            rsX.MoveFirst
            Do Until rsX.EOF
                If z = 0 Then cboKotaS.Tag = rsX("Kota")
                .AddItem rsX("Kota")
                rsX.MoveNext
                z = z + 1
            Loop
        End If
    End With
End Sub

Private Sub LoadCbo_KecS()
    Dim rsX As New ADODB.Recordset, z As Integer
    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kecamatan,'')) as Kecamatan From TopologyDaerah Where UPPER(Kota) = '" & Trim(UCase(cboKotaS.text)) & "' Order By Kecamatan")
    With cboKecS
        .Clear: z = 0: cboKecS.Tag = ""
        If rsX.RecordCount > 0 Then
            rsX.MoveFirst
            Do Until rsX.EOF
                If z = 0 Then cboKecS.Tag = rsX("Kecamatan")
                .AddItem rsX("Kecamatan")
                rsX.MoveNext
                z = z + 1
            Loop
        End If
    End With
End Sub

Private Sub LoadCbo_KelS()
    Dim rsX As New ADODB.Recordset, z As Integer
    Set rsX = objSID.FetchDataSID("Select Distinct UPPER(IsNull(Kelurahan,'')) as Kelurahan From TopologyDaerah Where UPPER(Kecamatan) = '" & Trim(UCase(cboKecS.text)) & "' Order By Kelurahan")
    With cboKelS
        .Clear: z = 0: cboKelS.Tag = ""
        If rsX.RecordCount > 0 Then
            rsX.MoveFirst
            Do Until rsX.EOF
                If z = 0 Then cboKelS.Tag = rsX("Kelurahan")
                .AddItem rsX("Kelurahan")
                rsX.MoveNext
                z = z + 1
            Loop
        End If
    End With
End Sub

Private Sub cboTitleS_GotFocus()
    If cboTitleS.BackColor = fColor Then
        cboTitleS.Locked = True
        Sendkeys ("{TAB}")
    Else
        cboTitleS.Locked = False
    End If
End Sub

Private Sub cmdBack_Click()
    FrAlm.Visible = False
    cmdBack.Visible = False
    Call Load_Cbo
End Sub

'
'Private Sub isSavingData()
'    If IsValid = True Then
'        If objP.isExist_ItemKodeLgn(txt(1).Text, txt(19).Text) = False Then
'            '-----------------------------------------------------------------------------------------------------------------
'            '   INSERT CUSTDAT
'            '-----------------------------------------------------------------------------------------------------------------
''            Call objP.AddNew_MasterPelanggan(txt(1).text, txt(2).text, txtnpwp.text, txt(4).text, txt(5).text, _
''                                         txt(6).text, txt(7).text, txt(8).text, txt(9).text, IIf(optTipe(0).value = True, 1, 0), _
''                                         (txt(10).text & "-" & txt(11).text & (IIf(txt(12) <> "", "-" & txt(12).text, ""))), _
''                                         txt(13).text, txt(14).text, IIf(Left(txt(14).text, 4) = "ZOTH", txt(16).text, ""), txt(17).text, txt(19).text, txt(20).text, _
''                                         (IIf(DtDOTam.text = "__/__/____", "01/01/3000", DtDOTam.value)), _
''                                         (IIf(DtDdate.text = "__/__/____", "01/01/3000", DtDdate.value)), _
''                                         txt(21).text, txt(23).text, chPolish.value, txt(31).text, txt(32).text, _
''                                         txt(33).text, txt(34).text, txt(35).text, txt(36).text, _
''                                         (IIf(DtPolish0.text = "__/__/____", "01/01/3000", DtPolish0.value)), _
''                                         (IIf(DtPolish1.text = "__/__/____", "01/01/3000", DtPolish1.value)), _
''                                         ChSurat.value, txt(37).text, txt(38).text, txt(39).text, _
''                                         txt(40).text, txt(41).text, txt(42).text, txt(43).text)
'        Else
'            '-----------------------------------------------------------------------------------------------------------------
'            '   UPDATE CUSTDAT
'            '-----------------------------------------------------------------------------------------------------------------
'            Call objP.Update_MasterPelanggan(txt(1).Text, txt(2).Text, txtnpwp.Text, txt(4).Text, txt(5).Text, _
'                                         txt(6).Text, txt(7).Text, txt(8).Text, txt(9).Text, IIf(optTipe(0).value = True, 1, 0), _
'                                         (txt(10).Text & "-" & txt(11).Text & (IIf(txt(12) <> "", "-" & txt(12).Text, ""))), _
'                                         txt(13).Text, txt(14).Text, IIf(Left(txt(14).Text, 4) = "ZOTH", txt(16).Text, ""), txt(17).Text, txt(19).Text, txt(20).Text, _
'                                         (IIf(DtDOTam.Text = "__/__/____", "01/01/3000", DtDOTam.value)), _
'                                         (IIf(DtDdate.Text = "__/__/____", "01/01/3000", DtDdate.value)), _
'                                         txt(21).Text, txt(23).Text, chPolish.value, txt(31).Text, txt(32).Text, _
'                                         txt(33).Text, txt(34).Text, txt(35).Text, txt(36).Text, _
'                                         (IIf(DtPolish0.Text = "__/__/____", "01/01/3000", DtPolish0.value)), _
'                                         (IIf(DtPolish1.Text = "__/__/____", "01/01/3000", DtPolish1.value)), _
'                                         ChSurat.value, txt(37).Text, txt(38).Text, txt(39).Text, _
'                                         txt(40).Text, txt(41).Text, txt(42).Text, txt(43).Text)
'            '-----------------------------------------------------------------------------------------------------------------
'            '   UPDATE JOBDTA
'            '-----------------------------------------------------------------------------------------------------------------
'            Call objWO.Update_DataJobDTA(txt(2).Text, txtnpwp.Text, txt(4).Text, txt(5).Text, txt(6).Text, txt(7).Text, txt(8).Text, _
'                                        (txt(10).Text & "-" & txt(11).Text & (IIf(txt(12) <> "", "-" & txt(12).Text, ""))), _
'                                        txt(13).Text, txt(14).Text, txt(17).Text, txt(20).Text, txt(19).Text)
'        End If
'        ft.ShowForm = "ARMST"
'        For C = 1 To 9
'            ft.DataPelanggan = txt(C).Text
'        Next C
'        ft.DataPelanggan = IIf(optTipe(0).value = True, "0", "1")
'        ft.Show vbModal
'        Call SetControlForm("load_update")
'    Else
'        With msg
'            .PassingMsg = "Data Master Pelanggan Belum Lengkap !" & Chr(13) & nValid
'            .passingTipe = 2
'            .Show vbModal
'        End With
'    End If
'End Sub

Private Sub BindToControl_MasterAcc(rs As ADODB.Recordset)
    For i = 1 To 9
        txt(i).text = ""
    Next i
    For i = 0 To 3
        txt(i + 1).text = NulSID(rs.Fields(i).value)
    Next i
    txt(6).text = IIf(Left(Replace(Replace(NulSID(rs.Fields(4).value), "-", ""), " ", ""), 2) = "08", "", Replace(Replace(NulSID(rs.Fields(4).value), "-", ""), " ", ""))
    txt(7).text = IIf(Left(Replace(Replace(NulSID(rs.Fields(5).value), "-", ""), " ", ""), 2) <> "08", "", Replace(Replace(NulSID(rs.Fields(5).value), "-", ""), " ", ""))
    txt(8).text = IIf(Left(Replace(Replace(NulSID(rs.Fields(5).value), "-", ""), " ", ""), 2) = "08", "", Replace(Replace(NulSID(rs.Fields(5).value), "-", ""), " ", ""))
    txt(10).text = ""
    txt(11).text = ""
    txt(12).text = ""
    For i = 13 To 20
        txt(i).text = ""
    Next i
    DtDOTam.text = "__/__/____"
    DtDdate.text = "__/__/____"
    For i = 21 To 29
        txt(i).text = ""
    Next i
    chPolish.value = 0
    For i = 31 To 36
        txt(i).text = ""
    Next i
    DtDOTam.text = "__/__/____"
    DtDdate.text = "__/__/____"
    ChSurat.value = 0
    For i = 37 To 43
        txt(i).text = ""
    Next i
End Sub

Private Sub BindToControl_Model(rs As ADODB.Recordset)
    txt(21).text = rs.Fields(0).value
    txt(22).text = rs.Fields(1).value
    txt(23).text = rs.Fields(2).value
End Sub

Private Sub BindToControl_Dealer(rs As ADODB.Recordset)
    txt(28).text = rs.Fields(0).value
    txt(29).text = rs.Fields(1).value
End Sub

Private Sub BindToControl_Warna(rs As ADODB.Recordset)
    txt(24).text = rs.Fields(0).value
    txt(25).text = rs.Fields(1).value
End Sub

Private Sub BindToControl_Asuransi(rs As ADODB.Recordset)
    txt(30).text = NulSID(rs.Fields(0).value)
    txt(31).text = NulSID(rs.Fields(1).value)
    txt(32).text = NulSID(rs.Fields(2).value)
    txt(33).text = NulSID(rs.Fields(3).value)
    txt(34).text = NulSID(rs.Fields(4).value)
    txt(35).text = NulSID(rs.Fields(5).value)
    txt(36).text = NulSID(rs.Fields(6).value)
End Sub

Private Sub BindToControl_Master(rs As ADODB.Recordset, isJustCust As Boolean)
    For i = 0 To 8
        txt(i + 1).text = rs.Fields(i).value
    Next i
    'SID optTipe(0).value = IIf(CInt(rs.Fields(9).value) = -1, True, False)
    'SID optTipe(1).value = IIf(CInt(rs.Fields(9).value) = 0, True, False)
    If isJustCust = False Then
        Call pecahNoPol(Trim(rs.Fields(10).value), nDpn, nTgh, nBlk)
        txt(10).text = nDpn
        txt(11).text = nTgh
        txt(12).text = nBlk
        For i = 13 To 20
            txt(i).text = rs.Fields(i - 2).value
        Next i
        If Left(UCase(txt(21).text), 4) = "ZOTH" Then
            txt(23).text = rs.Fields(47).value
        End If
        'SID DtDOTam.Text = IIf(Trim(rs.Fields(19).value) = "01/01/3000", "__/__/____", rs.Fields(19).value)
        'SID DtDdate.Text = IIf(Trim(rs.Fields(20).value) = "01/01/3000", "__/__/____", rs.Fields(20).value)
        For i = 21 To 29
            txt(i).text = NulSID(rs.Fields(i).value)
        Next i
        chPolish.value = IIf(rs.Fields(30).value = True, 1, 0)
        For i = 31 To 36
            txt(i).text = rs.Fields(i).value
        Next i
        DtPolish0.text = IIf(Trim(rs.Fields(37).value) = "01/01/3000", "__/__/____", rs.Fields(37).value)
        DtPolish1.text = IIf(Trim(rs.Fields(38).value) = "01/01/3000", "__/__/____", rs.Fields(38).value)
        ChSurat.value = IIf(rs.Fields(39).value = True, 1, 0)
        For i = 37 To 43
            txt(i).text = rs.Fields(i + 3).value
        Next i
    Else
        txt(10).text = ""
        txt(11).text = ""
        txt(12).text = ""
        For i = 13 To 20
            txt(i).text = ""
        Next i
        'SID DtDOTam.Text = "__/__/____"
        'SID DtDdate.Text = "__/__/____"
        For i = 21 To 29
            txt(i).text = ""
        Next i
        chPolish.value = 0
        For i = 31 To 36
            txt(i).text = ""
        Next i
        'SID DtDOTam.Text = "__/__/____"
        DtDdate.text = "__/__/____"
        ChSurat.value = 0
        For i = 37 To 43
            txt(i).text = ""
        Next i
    End If
End Sub

'Private Sub chPolish_Click()
'    If chPolish.value = 1 Then
'        If txt(4).Text = "" Then Exit Sub
'        For i = 37 To 47
'            If i = 37 Then
'                txt(i).Text = txt(4).Text & " " & txt(5).Text & IIf(txt(6).Text <> "", " " & txt(6).Text, "")
'            Else
'                txt(i).Text = txt(i - 31).Text
'            End If
'            txt(i).BackColor = fColor
'            If i = 38 Then
'                txt(i).Text = ""
'                txt(i).BackColor = tColor
'            End If
'        Next i
'        cboProvP.BackColor = fColor
'        cboKotaP.BackColor = fColor
'        cboKecP.BackColor = fColor
'        cboKelP.BackColor = fColor
'        If cboProv.Text = "LAINNYA" Or FrAlm.Visible = True Then
'            FrAlmP.Visible = True
'            cmdBackP.Visible = True
'            cmdBackP.Enabled = False
'            For i = 67 To 70
'                txt(i).Text = txt(i - 4).Text
'            Next i
'        Else
'            FrAlmP.Visible = False
'            cmdBackP.Visible = False
'
'            If cboProv.Text = "" Then GoTo GoKota
'            cboProvP.Tag = "ari"
'            cboProvP.Text = cboProv.Text
'            Call LoadCbo_KotaP
'            cboProvP.Tag = ""
'GoKota:
'            If cboKota.Text = "" Then GoTo GoKec
'            cboKotaP.Tag = "ari"
'            cboKotaP.Text = cboKota.Text
'            Call LoadCbo_KecP
'            cboKotaP.Tag = ""
'GoKec:
'            If cboKec.Text = "" Then GoTo GoKel
'            cboKecP.Tag = "ari"
'            cboKecP.Text = cboKec.Text
'            Call LoadCbo_KelP
'            cboKecP.Tag = ""
'GoKel:
'            If cboKel.Text = "" Then GoTo GoEnd
'            cboKelP.Tag = "ari"
'            cboKelP.Text = cboKel.Text
'            cboKelP.Tag = ""
'GoEnd:
'        End If
'        cboProvP.BackColor = fColor
'        cboKotaP.BackColor = fColor
'        cboKecP.BackColor = fColor
'        cboKelP.BackColor = fColor
'        txt(38).SetFocus
'    Else
'        For i = 37 To 47
'            txt(i).Text = ""
'            txt(i).BackColor = tColor
'            If i = 39 Then txt(i).BackColor = tColorx
'            If i = 40 Then txt(i).BackColor = tColorx
'            If i = 41 Then txt(i).BackColor = tColorx
'            If i = 43 Then txt(i).BackColor = tColorx
'            If i = 44 Then txt(i).BackColor = tColorx
'            If i = 45 Then txt(i).BackColor = tColorx
'            If i = 42 Then txt(i).BackColor = fColor
'        Next i
'        For i = 67 To 70
'            txt(i).Text = ""
'            txt(i).BackColor = tColor
'        Next i
'        cboProvP.BackColor = tColor
'        cboKotaP.BackColor = tColor
'        cboKecP.BackColor = tColor
'        cboKelP.BackColor = tColor
'        FrAlmP.Visible = False
'        cmdBackP.Visible = False
'        cboKotaP.Clear
'        cboKecP.Clear
'        cboKelP.Clear
'    End If
'End Sub

Private Sub chSurat_Click()
    If ChSurat.value = 1 Then
        For i = 37 To 43
            txt(i).text = txt(i - 35).text
        Next i
    Else
        For i = 37 To 43
            txt(i).text = ""
        Next i
    End If
End Sub

Private Sub GridFind_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        Dim k As String
        Dim filter As String
        k = NulSID(rsTemp!kodeLgn)
        filter = filterRsTemp 'rsTemp.filter
        
        If Not rsTemp.EOF Then
            'rsKodeLgn.
            rsTemp.MoveFirst
'            Do While Not rsTemp.EOF
'                rsTemp!Pilih = False
'                rsTemp.MoveNext
'            Loop
            rsTemp.filter = "Pilih=" & True
            If Not rsTemp.EOF Then
                
                rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                    rsTemp!Pilih = False
                    rsTemp.MoveNext
                Loop
            End If
            rsTemp.filter = ""
            
            rsTemp.filter = "kodelgn=" & k
            If Not rsTemp.EOF Then
                rsTemp!Pilih = True
            
            End If
            rsTemp.filter = ""
            rsTemp.filter = filter
        End If
        
        'rsInput!Pilih = False
        gridInput.Columns(0).value = False
        If CBool(gridInput.Columns(0).value) Then
            cmdUploadKTP.Enabled = True
            isTambahNew = True
        Else
            isTambahNew = False
            txtUploadKTP.text = ""
            cmdUploadKTP.Enabled = False
        End If
        gridInput.Refresh
        GridFind.Refresh
    End If
End Sub
  
Private Sub GridFind_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex = 0 Then
    
    Else
        Cancel = True
    End If
End Sub

Private Sub GridFind_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'    If ColIndex = 0 Then
'
'    Else
'        GridFind.Columns(ColIndex).value = OldValue
'    End If
End Sub

Private Sub gridInput_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        Dim filter As String
        If rsTemp.RecordCount > 0 Then
            filter = filterRsTemp ' rsTemp.filter
            rsTemp.MoveFirst
'            Do While Not rsTemp.EOF
'                rsTemp!Pilih = False
'                rsTemp.MoveNext
'            Loop
            rsTemp.filter = "Pilih=" & True
            If Not rsTemp.EOF Then
                rsTemp!Pilih = False
            End If
            rsTemp.filter = ""
            rsTemp.filter = filter
            rsTemp.MoveFirst
            
            Set cols = GridFind.Columns
            For Each col In cols
                col.FilterText = ""
            Next
             
        End If
        If CBool(gridInput.Columns(0).value) Then
            cmdUploadKTP.Enabled = True
            isTambahNew = True
        Else
            isTambahNew = False
            txtUploadKTP.text = ""
            cmdUploadKTP.Enabled = False
        End If
    End If
End Sub
 
Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 63, 64, 65, 66, 11
                
                If gMainKeyCari = "" Then gMainKeyCari = txt(Index)
                gSearchOption = "TOPOLOGY"
                If frmCari.Visible Then
                    frmCari.Hide
                End If
                frmCari.show (vbModal)
            Case 20
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDAGAMA"
                frmCari.show (vbModal)
            Case 7
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDPEKERJAAN"
                frmCari.show (vbModal)
            Case 21
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDHOBI"
                frmCari.show (vbModal)
            Case 22
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDOLAHRAGA"
                frmCari.show (vbModal)
            Case 1          'Custdat
                onFind = True
                onF12 = True
                ft.ShowForm = "SID"
                ft.show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    If cmdCreate.Enabled = True Then
                        Call BindToControl_Master(rsFind, True)
                        cmdEdit.Enabled = False
                        cmdDelete.Enabled = False
                        cmdOpen.Enabled = False
                    Else
                        Call BindToControl_Master(rsFind, False)
                        cmdEdit.Enabled = True
                        cmdDelete.Enabled = True
                        cmdOpen.Enabled = True
                    End If
                    cmdAdd.Visible = False
                    cmdCancel.Enabled = True
                Else
                    If ft.IsActi = "CLOSE" Then
                        txt(1).SetFocus
                    ElseIf ft.IsActi = "ADD" Then
                        Call CmdAdd_Click
                    End If
                End If
                onF12 = False
            Case 21          '" & DB_bengkel & "..Model
                onFind = True
                ft.ShowForm = "Model"
                ft.show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    Call BindToControl_Model(rsFind)
                    txt(24).SetFocus
                Else
                    txt(21).SetFocus
                End If
            Case 24          '" & DB_spk & "..InWarnaExt
                onFind = True
                ft.ShowForm = "Warna"
                ft.show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    Call BindToControl_Warna(rsFind)
                    txt(26).SetFocus
                Else
                    txt(24).SetFocus
                End If
            Case 28         '" & DB_bengkel & "..T_Dealer
                onFind = True
                ft.ShowForm = "Dealer"
                ft.show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    Call BindToControl_Dealer(rsFind)
                    txt(30).SetFocus
                Else
                    txt(28).SetFocus
                End If
            Case 30         'ACC11..ARMST(ASS)
                onFind = True
                ft.ShowForm = "AsuransiBp"
                ft.show vbModal
                If ft.IsPass = True Then
                    Set rsFind = ft.MyRs
                    Call BindToControl_Asuransi(rsFind)
                    txt(37).SetFocus
                Else
                    txt(30).SetFocus
                End If
        End Select
        onFind = False
    ElseIf KeyCode = 13 Then
        If Index = 11 Then
            If NulSID(txt(Index)) <> "" Then
                Dim rsTopology As ADODB.Recordset
                Set rsTopology = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TOPOLOGYDAERAH WITH (NOLOCK) WHERE KODEPOS = '" & NulSID(txt(11)) & "'")
                If Not rsTopology.EOF Then
                    rsTopology.MoveFirst
                    txt(63) = NulSID(rsTopology!Propinsi)
                    txt(64) = NulSID(rsTopology!Kota) 'nulsid(DGrid.Columns(1))
                    txt(65) = NulSID(rsTopology!kecamatan) 'nulsid(DGrid.Columns(2))
                    txt(66) = NulSID(rsTopology!kelurahan) 'nulsid(DGrid.Columns(3))
                    txt(11) = NulSID(rsTopology!KodePos)
                Else
                    MsgBox "Data kodepos tidak terdaftar.", vbInformation
                    txt(63) = ""
                    txt(64) = "" 'nulsid(DGrid.Columns(1))
                    txt(65) = "" 'nulsid(DGrid.Columns(2))
                    txt(66) = "" 'nulsid(DGrid.Columns(3))
                    txt(11) = "" 'nulsid(DGrid.Columns(4))
                End If
            End If
        Else
            Call txt_Validate(Index, False)
        End If
    Else
        Select Case Index
            Case 2
                If cboTipe.text = "" Then
                    MsgBox "Tipe Customer Belum dipilih!"
                    cboTipe.SetFocus
                End If
'            Case 63, 64, 65, 66, 20, 7, 21, 22
'                KeyCode = 0
                 
        End Select
        
    End If
End Sub

'=====================================================================================
'============================= SETTING CONTROL =======================================
'=====================================================================================
Private Sub txt_GotFocus(Index As Integer)
    'Label1(3).Caption = Index: DoEvents
    Dim DiTab As Boolean
    DiTab = False
    If txt(Index).BackColor = fColor Then
        txt(Index).Locked = True
        Sendkeys ("{TAB}")
    Else
        txt(Index).Locked = False
        If Index = 9 Then
            If txt(8).text = "" Then txt(8).SetFocus
        End If
    End If
    If gMainKeyCari <> "" Then
        txt(Index) = UCase(NulSID(gMainKeyCari))
        gMainKeyCari = ""
        DiTab = True
        Sendkeys ("{TAB}")
        'txt(Index).SetFocus
        'SendKeys ("{TAB}")
    End If
    
    If Index = 7 Then
        If NulSID(UCase(txt(Index))) = "LAINNYA" Then
            'txt(23).SetFocus
            txt(23).BackColor = &HFFFFFF
        End If
    End If
'    If DiTab Then
'        SendKeys ("{TAB}")
'    Else
'        SendKeys ("{HOME}+{END}")
'    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = 19 Then
        Select Case KeyAscii
            Case 65 To 90
                'OK
            Case 97 To 122
                'OK
            Case 48 To 57
                'OK
            Case 8
                'OK
            Case 13
                'OK
            Case Else
                KeyAscii = 0
        End Select
    End If
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        If onFind = False Then
            If Index = 1 Or Index = 24 Or Index = 21 Or Index = 23 Then
                Select Case Index
                    Case 1
                        If txt(1).text = "" Then Exit Sub
                        If onF12 = True Then Exit Sub
                        If cmdSave.Enabled = True Then Exit Sub
                        Dim rsExist As ADODB.Recordset
                        Set rsExist = objP.FetchListAllPelanggan("Where KodeLg = '" & txt(1).text & "'")
                        If rsExist.RecordCount = 0 Then
                            Dim rsExistAcc As ADODB.Recordset
                            Set rsExistAcc = objP.FetchPelangganACC("Where a.KodeLgn = '" & txt(1).text & "'")
                            If rsExistAcc.RecordCount > 0 Then
                                With msg
                                    .PassingMsg = "Data Master Pelanggan Kode : " & txt(1).text & " Sudah terdaftar di Accounting, Tampilkan Data ?"
                                    .passingTipe = 1
                                    .show vbModal
                                    If .GetMsg = True Then
                                        If cmdAdd.Enabled = True Or cmdEdit.Enabled = True Then
                                            If cmdCreate.Enabled = True Then
                                                Call BindToControl_MasterAcc(rsExistAcc)
                                                cmdEdit.Enabled = False
                                                cmdDelete.Enabled = False
                                                cmdOpen.Enabled = False
                                            Else
                                                Call BindToControl_MasterAcc(rsExistAcc)
                                                cmdOpen.Enabled = True
                                                cmdEdit.Enabled = True
                                                cmdDelete.Enabled = True
                                            End If
                                        Else
                                            Call BindToControl_MasterAcc(rsExistAcc)
                                            cmdEdit.Enabled = False
                                            cmdDelete.Enabled = False
                                            cmdOpen.Enabled = False
                                        End If
                                        cmdAdd.Enabled = False
                                        cmdAdd.Visible = False
                                        cmdCancel.Enabled = True
                                    Else
                                        txt(1).text = ""
                                        If cmdCreate.Enabled = True Or cmdSave.Enabled = True Then Exit Sub
                                        Call SetControlForm("load")
                                    End If
                                End With
                            Else
                                'If cmdCreate.Enabled = True Or cmdSave.Enabled = True Then Exit Sub
                                With msg
                                    .PassingMsg = "Data Master Pelanggan Baru ?"
                                    .passingTipe = 1
                                    .show vbModal
                                    If .GetMsg = True Then
                                        Dim nkode As String
                                        nkode = txt(1).text
                                        Call SetControlForm("add")
                                        txt(1).text = nkode
                                    Else
                                        txt(1).text = ""
                                    End If
                                End With
                            End If
                        Else
                            With msg
                                .PassingMsg = "Data Master Pelanggan Kode : " & txt(1).text & " Sudah terdaftar, Tampilkan Data ?"
                                .passingTipe = 1
                                .show vbModal
                                If .GetMsg = True Then
                                    If cmdAdd.Enabled = True Or cmdEdit.Enabled = True Then
                                        If cmdCreate.Enabled = True Then
                                            Call BindToControl_Master(rsExist, True)
                                            cmdEdit.Enabled = False
                                            cmdDelete.Enabled = False
                                            cmdOpen.Enabled = False
                                        Else
                                            Call BindToControl_Master(rsExist, False)
                                            cmdOpen.Enabled = True
                                            cmdEdit.Enabled = True
                                            cmdDelete.Enabled = True
                                        End If
                                    Else
                                        Call BindToControl_Master(rsExist, True)
                                        cmdEdit.Enabled = False
                                        cmdDelete.Enabled = False
                                        cmdOpen.Enabled = False
                                    End If
                                    cmdAdd.Enabled = False
                                    cmdAdd.Visible = False
                                    cmdCancel.Enabled = True
                                Else
                                    txt(1).text = ""
                                    If cmdCreate.Enabled = True Or cmdSave.Enabled = True Then Exit Sub
                                    Call SetControlForm("load")
                                End If
                            End With
                            'cmdOpen.SetFocus
                        End If
'                    Case 21
'                        'Dim objM As New clsModel
'                        Set rsFind = objM.FetchListNamaModel(" Where MDL = '" & txt(21).Text & "'")
'                        If rsFind.RecordCount > 0 Then
'                            Call BindToControl_Model(rsFind)
'                        Else
'                            txt(21).Text = ""
'                            txt(22).Text = ""
'                            txt(23).Text = ""
'                        End If
'                    Case 24
'                        Dim objW As New clsWarna
'                        Set rsFind = objW.FetchListWarna(" Where KodeWarna = '" & txt(24).Text & "'")
'                        If rsFind.RecordCount > 0 Then
'                            Call BindToControl_Warna(rsFind)
'                        Else
'                            txt(24).Text = ""
'                            txt(25).Text = ""
'                        End If
'                    Case 28
'                        Set rsFind = objD.FetchListNamaDealer(" Where FS_KD_Dealer = '" & txt(21).Text & "'")
'                        If rsFind.RecordCount > 0 Then
'                            Call BindToControl_Dealer(rsFind)
'                        Else
'                            txt(28).Text = ""
'                            txt(29).Text = ""
'                        End If
'                    Case 30
'                        Dim objA As New clsAsuransi
'                        Set rsFind = objA.FetchListAsuransiBp(" Where Kode = '" & txt(30).Text & "'")
'                        If rsFind.RecordCount > 0 Then
'                            Call BindToControl_Asuransi(rsFind)
'                        Else
'                            txt(30).Text = ""
'                            txt(31).Text = ""
'                            txt(32).Text = ""
'                            txt(33).Text = ""
'                            txt(34).Text = ""
'                            txt(35).Text = ""
'                            txt(36).Text = ""
'                        End If
                End Select
            Else
                If Index = 66 Then
                    txt(8).SetFocus
                End If
            End If
        End If
    Else
        If txt(Index).BackColor = &HC0E0FF Then
            KeyAscii = 0
            Exit Sub
        End If
        If KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            Select Case Index
                Case 3, 9, 10, 12, 13, 14, 15, 19, 40, 41, 43, 44, 45, 46, 55, 56, 58, 59, 60, 61
                    If KeyAscii <= 57 And KeyAscii >= 48 Then
                        KeyAscii = KeyAscii
                    ElseIf KeyAscii = 46 Then
                        KeyAscii = KeyAscii
                    Else
                        KeyAscii = 0
                    End If
                Case 2
                    If Label5(0).Caption <> "Passport / KITAS" Then
                        If KeyAscii <= 57 And KeyAscii >= 48 Then
                            KeyAscii = KeyAscii
                        ElseIf KeyAscii = 46 Then
                            'KeyAscii = KeyAscii
                            KeyAscii = 0
                        Else
                            KeyAscii = 0
                        End If
                    Else
                        If KeyAscii <= 57 And KeyAscii >= 48 Then
                            KeyAscii = KeyAscii
                        ElseIf KeyAscii = 46 Then
                            KeyAscii = KeyAscii
                        ElseIf KeyAscii <= 90 And KeyAscii >= 65 Then
                            KeyAscii = KeyAscii
                        ElseIf KeyAscii <= 122 And KeyAscii >= 97 Then
                            KeyAscii = KeyAscii
                        Else
                            KeyAscii = 0
                        End If
                    End If
'                Case 63, 64, 65, 66, 20, 7, 21, 22
'                    KeyAscii = 0
                Case Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
        End If
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
     
    If gMainKeyCari <> "" Then
        txt(Index) = UCase(NulSID(gMainKeyCari))
        gMainKeyCari = ""
        'SendKeys ("{TAB}")
    End If
End Sub

'Private Sub DtDotam_GotFocus()
'    If DtDOTam.BackColor = fColor Then SendKeys "{TAB}"
'End Sub

'Private Sub DtDdate_GotFocus()
'    If DtDdate.BackColor = fColor Then SendKeys "{TAB}"
'End Sub
'
'Private Sub DtPolish0_GotFocus()
'    If DtPolish0.BackColor = fColor Then SendKeys "{TAB}"
'End Sub
'
'Private Sub DtPolish1_GotFocus()
'    If DtPolish1.BackColor = fColor Then SendKeys "{TAB}"
'End Sub

'Private Sub optTipe_GotFocus(Index As Integer)
'    If optTipe(Index).ForeColor = fColorz Then SendKeys "{TAB}"
'End Sub

'Private Sub chPolish_GotFocus()
'    If chPolish.ForeColor = fColorz Then SendKeys "{TAB}"
'End Sub

Private Sub chSurat_GotFocus()
    If ChSurat.ForeColor = fColorz Then Sendkeys "{TAB}"
End Sub
'
'Private Sub DtDotam_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub
'
'Private Sub DtDdate_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub
'
'Private Sub DtPolish0_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

'Private Sub DtPolish1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub
'
'Private Sub optTipe_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

'Private Sub chPolish_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{TAB}"
'End Sub

Private Sub chSurat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub cboTipe_Click()
    Call cboTipe_Change
End Sub

Private Sub cboTipe_Change()
    If cboTipe.text <> "" Then
        Dim rsZ As New ADODB.Recordset
        Set rsZ = objSID.FetchDataSID("Select b.Title From TipeCustomer a Inner Join Title b On a.Kode=b.Tipe Where Upper(a.Tipe) = '" & UCase(cboTipe.text) & "'")
        With cboTitle
            .Clear
            If rsZ.RecordCount > 0 Then
                rsZ.MoveFirst
                Do Until rsZ.EOF
                    .AddItem rsZ("Title")
                    rsZ.MoveNext
                Loop
            End If
        End With
        DtLahir.BackColor = &HFFFFFF
        If UCase(cboTipe.text) = "PERUSAHAAN UMUM" Then
            Label5(0).Caption = "No. TDP"
            txtNpwp.BackColor = &HFFFFFF ' &H80FFFF
            txt(2).BackColor = &HFFFFFF
        ElseIf UCase(cboTipe.text) = "PERORANGAN" Then
            Label5(0).Caption = "No. KTP"
            txt(2).BackColor = &HFFFFFF ' &H80FFFF
            txtNpwp.BackColor = &HFFFFFF
            DtLahir.BackColor = &H80FFFF
        Else
            Label5(0).Caption = "No. TDP"
            txtNpwp.BackColor = &HFFFFFF ' &H80FFFF
            txt(2).BackColor = &HFFFFFF
        End If
        If isEditData = False Then
            If UCase(cboTipe.text) = "PERUSAHAAN UMUM" Then
                Label5(0).Caption = "No. TDP"
                txt(1).text = LastID(cboTipe.text)
            ElseIf UCase(cboTipe.text) = "PERORANGAN" Then
                Label5(0).Caption = "No. KTP"
                txt(1).text = LastID(cboTipe.text)
            Else
                Label5(0).Caption = "No. TDP"
                txt(1).text = LastID(cboTipe.text)
            End If
            
            If UCase(cboTipe.text) <> "PERORANGAN" And NulSID(cboTipe.text) <> "" Then
                txt(18).Enabled = True
                txt(19).Enabled = True
                txt(18).BackColor = &H80FFFF
                txt(19).BackColor = &H80FFFF
                
            Else
                txt(18).text = ""
                txt(19).text = ""
                txt(18).Enabled = False
                txt(19).Enabled = False
                txt(18).BackColor = &HC0E0FF
                txt(19).BackColor = &HC0E0FF
                
            End If
        End If
        txt(2).SetFocus
        
    End If
End Sub

Private Sub cboTitle_Click()
    Call cboTitle_Change
End Sub

Private Sub cboTitle_Change()
    If cboTipe.text = "" Then
        MsgBox "Tipe Customer belum dipilih!"
        cboTipe.SetFocus
    End If
    If txt(4).Enabled Then txt(4).SetFocus
End Sub

'Private Sub ViewPassing()
'    If passBy = "" Then Exit Sub
'    Select Case passBy
'        Case "BOK"
'            If rsBy.Fields(6).value = "" Then
'                Call SetControlForm("add")
'            Else
'                Call SetControlForm("edit")
'            End If
'            For C = 1 To 8
'                txt(C).Text = rsBy.Fields(C + 5).value
'            Next C
'            Call pecahNoPol(Trim(rsBy.Fields(5).value), nDpn, nTgh, nBlk)
'            txt(10).Text = nDpn
'            txt(11).Text = nTgh
'            txt(12).Text = nBlk
'            txt(13).Text = Trim(rsBy.Fields(14).value)
'            txt(21).Text = Trim(rsBy.Fields(15).value)
'            Call txt_KeyPress(21, 13)
'            txt(24).Text = Trim(rsBy.Fields(16).value)
'            Call txt_KeyPress(24, 13)
'            txt(19).Text = Trim(rsBy.Fields(17).value)
'            txt(20).Text = Trim(rsBy.Fields(18).value)
''            If Val(Right(Trim(rsBy.Fields(19).value), 4)) < 1901 Then      'SID
''                DtDOTam.Text = "__/__/____"
''            Else
''                DtDOTam.Text = IIf(Trim(rsBy.Fields(19).value) = "01/01/3000", "__/__/____", rsBy.Fields(19).value)
''            End If
'            DtDdate.Text = IIf(Trim(rsBy.Fields(20).value) = "01/01/3000", "__/__/____", rsBy.Fields(20).value)
'            txt(28).Text = Trim(rsBy.Fields(21).value)
'            Call txt_KeyPress(28, 13)
'            txt(23).Text = Trim(rsBy.Fields(22).value)
'            Call txt_KeyPress(23, 13)
'            If txt(1).Text <> "" And txt(19).Text <> "" Then
'                Dim rsFindIn As ADODB.Recordset
'                Set rsFindIn = objP.FetchListAllPelanggan(" Where c.KodeLg = '" & txt(1).Text & "' And c.NoRangka = '" & txt(19).Text & "'")
'                If rsFindIn.RecordCount > 0 Then
'                    Call cmdCancel_Click
'                    Call BindToControl_Master(rsFindIn, False)
'                    cmdEdit.Enabled = True
'                    cmdDelete.Enabled = True
'                    cmdOpen.Enabled = True
'                    cmdAdd.Visible = False
'                    cmdCancel.Enabled = False
'                End If
'            End If
'        Case "NEWBOK"
'            On Error GoTo ErNul
'            If rsBy.Fields(0).value = "" Then
'                Call SetControlForm("add")
'            Else
'                Call SetControlForm("edit")
'            End If
'            For C = 1 To 9
'                txt(C).Text = rsBy.Fields(C - 1).value
'            Next C
'            'optTipe(CInt(IIf(rsBy.Fields(9).value = "0", 1, 0))).value = True  'SID
'            Call pecahNoPol(Trim(rsBy.Fields(10).value), nDpn, nTgh, nBlk)
'            txt(10).Text = nDpn
'            txt(11).Text = nTgh
'            txt(12).Text = nBlk
'            txt(13).Text = Trim(rsBy.Fields(11).value)
'            txt(14).Text = Trim(rsBy.Fields(12).value)
'            Call txt_KeyPress(14, 13)
'            txt(24).Text = Trim(rsBy.Fields(15).value)
'            Call txt_KeyPress(17, 13)
'            txt(19).Text = Trim(rsBy.Fields(17).value)
'            txt(20).Text = Trim(rsBy.Fields(18).value)
''SID         If Val(Right(Trim(rsBy.Fields(19).value), 4)) < 1901 Then
''                DtDOTam.Text = "__/__/____"
''            Else
''                DtDOTam.Text = IIf(Trim(rsBy.Fields(19).value) = "01/01/3000", "__/__/____", rsBy.Fields(19).value)
''            End If
'            DtDdate.Text = IIf(Trim(rsBy.Fields(20).value) = "01/01/3000", "__/__/____", rsBy.Fields(20).value)
'            txt(21).Text = Trim(rsBy.Fields(21).value)
'            Call txt_KeyPress(21, 13)
'            txt(23).Text = Trim(rsBy.Fields(23).value)
'            Call txt_KeyPress(23, 13)
'            For C = 31 To 36
'                txt(C).Text = Trim(rsBy.Fields(C).value)
'            Next C
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(37).value) = "01/01/3000", "__/__/____", rsBy.Fields(37).value)
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(38).value) = "01/01/3000", "__/__/____", rsBy.Fields(38).value)
'            For C = 37 To 43
'                txt(C).Text = Trim(rsBy.Fields(C + 3).value)
'            Next C
'            If txt(1).Text <> "" And txt(19).Text <> "" Then
'                Dim rsFindInx As ADODB.Recordset
'                Set rsFindInx = objP.FetchListAllPelanggan(" Where c.KodeLg = '" & txt(1).Text & "' And c.NoRangka = '" & txt(19).Text & "'")
'                If rsFindInx.RecordCount > 0 Then
'                    Call cmdCancel_Click
'                    Call BindToControl_Master(rsFindInx, False)
'                    cmdEdit.Enabled = True
'                    cmdDelete.Enabled = True
'                    cmdOpen.Enabled = True
'                    cmdAdd.Visible = False
'                    cmdCancel.Enabled = False
'                End If
'            End If
'        Case "EST"
'            On Error GoTo ErNul
'            If rsBy.Fields(5).value = "" Then
'                Call SetControlForm("add")
'            Else
'                Call SetControlForm("edit")
'            End If
'            For C = 1 To 9
'                txt(C).Text = rsBy.Fields(C + 4).value
'            Next C
'            If rsBy.Fields(4).value <> "" Then
'                Call pecahNoPol(Trim(rsBy.Fields(4).value), nDpn, nTgh, nBlk)
'                txt(10).Text = nDpn
'                txt(11).Text = nTgh
'                txt(12).Text = nBlk
'            Else
'                txt(10).Text = ""
'                txt(11).Text = ""
'                txt(12).Text = ""
'            End If
'            txt(13).Text = Trim(rsBy.Fields(14).value)
'            txt(14).Text = Trim(rsBy.Fields(15).value)
'            Call txt_KeyPress(14, 13)
'            txt(17).Text = Trim(rsBy.Fields(16).value)
'            Call txt_KeyPress(17, 13)
'            txt(19).Text = Trim(rsBy.Fields(17).value)
'            txt(20).Text = Trim(rsBy.Fields(18).value)
''SID         If Val(Right(Trim(rsBy.Fields(19).value), 4)) < 1901 Then
''                DtDOTam.Text = "__/__/____"
''            Else
''                DtDOTam.Text = IIf(Trim(rsBy.Fields(19).value) = "01/01/3000", "__/__/____", rsBy.Fields(19).value)
''            End If
'            DtDdate.Text = IIf(Trim(rsBy.Fields(20).value) = "01/01/3000", "__/__/____", rsBy.Fields(20).value)
'            txt(21).Text = Trim(rsBy.Fields(21).value)
'            Call txt_KeyPress(21, 13)
'            txt(23).Text = Trim(rsBy.Fields(22).value)
'            Call txt_KeyPress(23, 13)
'            For C = 31 To 36
'                txt(C).Text = Trim(rsBy.Fields(C - 8).value)
'            Next C
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(29).value) = "01/01/3000", "__/__/____", rsBy.Fields(29).value)
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(30).value) = "01/01/3000", "__/__/____", rsBy.Fields(30).value)
'            If txt(1).Text <> "" And txt(19).Text <> "" Then
'                Dim rsExist As ADODB.Recordset
'                Set rsExist = objP.FetchListAllPelanggan("Where C.KodeLg = '" & txt(1).Text & "' And NoRangka = '" & txt(19).Text & "'")
'                If rsExist.RecordCount > 0 Then
'                    Call BindToControl_Master(rsExist, False)
'                End If
'            End If
'        Case "WO"
'        Case "NewEST"
'
'            If rsBy.EOF = True Then
'                Call SetControlForm("add")
'                Exit Sub
'            Else
'                txt(1).Text = rsBy.Fields(0).value
'                If rsBy.Fields(0).value = "" Then
'                    Call SetControlForm("add")
'                Else
'                    Call SetControlForm("edit")
'                End If
'            End If
'            For C = 1 To 8
'                txt(C + 1).Text = rsBy.Fields(C).value
'            Next C
'            If rsBy.Fields(9).value = False Then
'                'SID Me.optTipe(1).value = True
'            Else
'                'SID Me.optTipe(0).value = True
'            End If
'            If rsBy.Fields(10).value <> "" Then
'                Call pecahNoPol(Trim(rsBy.Fields(10).value), nDpn, nTgh, nBlk)
'                txt(10).Text = nDpn
'                txt(11).Text = nTgh
'                txt(12).Text = nBlk
'            Else
'                txt(10).Text = ""
'                txt(11).Text = ""
'                txt(12).Text = ""
'            End If
'            txt(13).Text = Trim(rsBy.Fields(11).value)
'            txt(14).Text = Trim(rsBy.Fields(12).value)
'            Call txt_KeyPress(14, 13)
'            txt(17).Text = Trim(rsBy.Fields(15).value)
'            Call txt_KeyPress(17, 13)
'            txt(19).Text = Trim(rsBy.Fields(17).value)
'            txt(20).Text = Trim(rsBy.Fields(18).value)
''SID         If Val(Right(Trim(rsBy.Fields(19).value), 4)) < 1901 Then
''                DtDOTam.Text = "__/__/____"
''            Else
''                DtDOTam.Text = IIf(Trim(rsBy.Fields(19).value) = "01/01/3000", "__/__/____", rsBy.Fields(19).value)
''            End If
'            DtDdate.Text = IIf(Trim(rsBy.Fields(20).value) = "01/01/3000", "__/__/____", rsBy.Fields(20).value)
'            txt(21).Text = Trim(rsBy.Fields(21).value)
'            Call txt_KeyPress(21, 13)
'            txt(23).Text = Trim(rsBy.Fields(23).value)
'            Call txt_KeyPress(23, 13)
'            chPolish.value = IIf(rsBy.Fields(30).value = True, 1, 0)
'            For C = 31 To 36
'                txt(C).Text = Trim(rsBy.Fields(C).value)
'            Next C
'            DtPolish0.Text = IIf(Trim(rsBy.Fields(37).value) = "01/01/3000", "__/__/____", rsBy.Fields(37).value)
'            DtPolish1.Text = IIf(Trim(rsBy.Fields(38).value) = "01/01/3000", "__/__/____", rsBy.Fields(38).value)
'            ChSurat.value = IIf(rsBy.Fields(39).value = True, 1, 0)
'            For C = 37 To 43
'                txt(C).Text = Trim(rsBy.Fields(C + 3).value)
'            Next C
'        Case "NEWBOKWAC"
'            On Error GoTo ErNul
'            'rsBy.Sort = "[" & rsBy.Fields(0).value & "]" & " Asc"
'            rsBy.Find "Pilih = 'TRUE'", , adSearchForward, 1
'            If rsBy.RecordCount = 0 Then Exit Sub
'            If rsBy.Fields(0).value = False Then
'                Call SetControlForm("add")
'            Else
'                Call SetControlForm("edit")
'            End If
'
'            txt(1).Text = ""
'            txt(2).Text = nulsid(rsBy("Nama"), "")
'            txtnpwp.Text = nulsid(rsBy("Alamat"), "")
'            txt(4).Text = nulsid(rsBy("Kota"), "")
'            txt(5).Text = ""
'            txt(6).Text = nulsid(rsBy("Telp"), "")
'            txt(7).Text = ""
'            txt(8).Text = ""
'            txt(9).Text = nulsid(rsBy("Email"), "")
'            'SID optTipe(0).value = False
'            'SID optTipe(1).value = False
'            Call pecahNoPol(Trim(rsBy("NoPolisi")), nDpn, nTgh, nBlk)
'            txt(10).Text = nDpn
'            txt(11).Text = nTgh
'            txt(12).Text = nBlk
'            txt(13).Text = Trim(nulsid(rsBy("ThMobil"), ""))
'            txt(14).Text = Trim(nulsid(rsBy("FullModel"), ""))
'            Call txt_KeyPress(14, 13)
'            txt(17).Text = ""
'            'Call txt_KeyPress(17, 13)
'            txt(18).Text = rsBy("Color")
'            txt(19).Text = Trim(nulsid(rsBy("NoRangka"), ""))
'            txt(20).Text = Trim(nulsid(rsBy("NoMesin"), ""))
'            'SID DtDOTam.Text = "__/__/____"
'            DtDdate.Text = "__/__/____"
'            txt(28).Text = Trim(nulsid(rsBy("KodeDealerWAC"), ""))
'            Call txt_KeyPress(28, 13)
'            'txt(23).Text = Trim(nulsid(rsBy("KodeAssBP"), ""))
''            Call txt_KeyPress(23, 13)
''            txt(31).Text = Trim(nulsid(rsBy("NamaPo"), ""))
''            txt(32).Text = Trim(nulsid(rsBy("AlamatPo"), ""))
''            txt(33).Text = Trim(nulsid(rsBy("KotaPo"), ""))
''            txt(34).Text = ""
''            txt(35).Text = Trim(nulsid(rsBy("TelpPo"), ""))
''            txt(36).Text = Trim(nulsid(rsBy("NomorPolish"), ""))
''            DtPolish0.Text = IIf(Trim(nulsid(rsBy("JangkaPolishAwal"), "01/01/3000")) = "01/01/3000", "__/__/____", rsBy("JangkaPolishAwal"))
''            DtPolish1.Text = IIf(Trim(nulsid(rsBy("JangkaPolishAkhir"), "01/01/3000")) = "01/01/3000", "__/__/____", rsBy("JangkaPolishAkhir"))
'            If xAs(0) <> "" Then
'                chPolish.value = 1
'            Else
'                chPolish.value = 0
'            End If
'            txt(23).Text = xAs(0)
'            Call txt_KeyPress(23, 13)
'            txt(31).Text = xAs(1)
'            txt(32).Text = xAs(2)
'            txt(33).Text = xAs(3)
'            txt(34).Text = ""
'            txt(35).Text = xAs(4)
'            txt(36).Text = xAs(5)
'            DtPolish0.Text = xAs(6)
'            DtPolish1.Text = xAs(7)
'            For C = 37 To 43
'                txt(C).Text = ""
'            Next C
'            If txt(1).Text <> "" And txt(19).Text <> "" Then
'                Dim rsFindIny As ADODB.Recordset
'                Set rsFindIny = objP.FetchListAllPelanggan(" Where c.KodeLg = '" & txt(1).Text & "' And c.NoRangka = '" & txt(19).Text & "'")
'                If rsFindIny.RecordCount > 0 Then
'                    Call cmdCancel_Click
'                    Call BindToControl_Master(rsFindIny, False)
'                    cmdEdit.Enabled = True
'                    cmdDelete.Enabled = True
'                    cmdOpen.Enabled = True
'                    cmdAdd.Visible = False
'                    cmdCancel.Enabled = False
'                End If
'            End If
'    End Select
'    On Error Resume Next
'    If xWAC = False Then
'        If Left(UCase(txt(21).Text), 4) = "ZOTH" Then
'            txt(23).Text = rsBy.Fields(47).value
'        End If
'    End If
'ErNul:
'End Sub
'
'Private Sub VSc_Change()
'    PicMove.Top = -(VSc.value)
'End Sub
'
'Private Sub VSc_Scroll()
'    PicMove.Top = -(VSc.value)
'End Sub
'
'Private Sub cmdSave_Click()
'    '-----------------------------------------------------------------------------------------------------------------
'    '   CEK DI ARMST KODE & NAMA
'    '-----------------------------------------------------------------------------------------------------------------
'    Dim rsCPL As ADODB.Recordset
'    Set rsCPL = objP.FetchCekPelangganACC(Trim(txt(1).Text))
'    If rsCPL.RecordCount > 0 Then
'        If Trim(rsCPL("NamaLgn")) = Trim(txt(2).Text) Then
'            Call isSavingData
'        Else
'            With msg
'                .PassingMsg = "Data Master Pelanggan Kode : " & txt(1).Text & "'" & Chr(13) & "Sudah Terdaftar Atas Nama : " & rsCPL("NamaLgn") & Chr(13) & "Silahkan ganti Kode Pelanggan!"
'                .passingTipe = 2
'                .Show vbModal
'            End With
'            Exit Sub
'        End If
'    Else
'        Call isSavingData
'    End If
'End Sub

Private Sub CmdAdd_Click()
    Call SetControlForm("add")
End Sub
'
'Private Sub cmdCreate_Click()
'    Dim rsRangka As ADODB.Recordset
'    '-------------------------------------------------------------
'    '   1. Cek NoRangka di CustDat
'    '-------------------------------------------------------------
'    Set rsRangka = objP.FetchExistRangka(txt(19).Text)
'    If rsRangka.RecordCount > 0 Then
'        With msg
'            .PassingMsg = "Nomor Rangka '" & txt(19).Text & "'" & Chr(13) & Chr(13) & "SUDAH TERDAFTAR !"
'            .passingTipe = 2
'            .Show vbModal
'        End With
'        Exit Sub
'    End If
'
'    Dim rsLgn As ADODB.Recordset
'    '-------------------------------------------------------------
'    '   Cek KodeLgn di ARMST
'    '-------------------------------------------------------------
'    Set rsLgn = objP.FetchPelangganACC(" Where KodeLgn = '" & txt(1).Text & "'")
'    If rsLgn.RecordCount > 0 Then
'    '-------------------------------------------------------------
'    '   2. Terdaftar di ARMST
'    '-------------------------------------------------------------
'        With msg
'            .PassingMsg = "'" & txt(1).Text & "' TERDAFTAR DI ACC!" & Chr(13) & "Nama : " & rsLgn.Fields(1).value & Chr(13) & "Alamat : " & rsLgn.Fields(2).value
'            .passingTipe = 2
'            .Show vbModal
'        End With
'        If Trim(rsLgn.Fields(1).value) = Trim(txt(2).Text) Then
'        '-------------------------------------------------------------
'        '   3. Ada | Jika nama sama, Cek di Custdat (Kode & Nama)
'        '-------------------------------------------------------------
'            Dim rsCRK As ADODB.Recordset
'            Set rsCRK = objP.FetchCekPelanggan(Trim(txt(1).Text), txt(2).Text)
'            If rsCRK.RecordCount > 0 Then
'                If Trim(rsCRK("NoRangka")) <> Trim(txt(19).Text) Then
'                '-------------------------------------------------------------
'                '  4. Ada | Jika No Rangka tidak sesuai, buat kode baru
'                '-------------------------------------------------------------
'                    With msg
'                        .PassingMsg = "Tambahkan No Rangka : " & Trim(txt(19).Text) & Chr(13) & " pada KodeLgn : " & Trim(txt(1).Text) & "?"
''''                        .PassingMsg = "No Rangka Tidak Sesuai : " & rsCRK("NoRangka") & Chr(13) & " Ganti Kode Pelanggan Baru"
'                        .passingTipe = 1
'                        .Show vbModal
'                        If .GetMsg = False Then
'                            Exit Sub
'                        Else
'                            Call isSavingData
'                        End If
'                    End With
'                Else
'                '-------------------------------------------------------------
'                '   4. Jika No Rangka sesuai, update data
'                '-------------------------------------------------------------
'                    With msg
'                        .PassingMsg = "'" & txt(1).Text & "' SUDAH TERDAFTAR !" & Chr(13) & Chr(13) & "YAKIN UPDATE DATA ???"
'                        .passingTipe = 1
'                        .Show vbModal
'                        If .GetMsg = False Then
'                            Exit Sub
'                        Else
'                            Call isSavingData
'                        End If
'                    End With
'                End If
'            Else
'            '-------------------------------------------------------------
'            '  Simpan Data
'            '-------------------------------------------------------------
'                Call isSavingData
'            End If
'        Else
'        '-------------------------------------------------------------
'        '   3. Gnti Kode Baru
'        '-------------------------------------------------------------
'            With msg
'                .PassingMsg = "'" & txt(1).Text & "' tidak sesuai ACC! (" & rsLgn("NamaLgn") & ")" & Chr(13) & " Ganti Kode Pelanggan Baru"
'                .passingTipe = 2
'                .Show vbModal
'            End With
'            txt(1).Text = ""
'            txt(1).SetFocus
'            Exit Sub
'        End If
'    Else
'    '-------------------------------------------------------------
'    '   2. Tidak terdaftar di ARMST
'    '-------------------------------------------------------------
'        Call isSavingData
'    End If
'End Sub
'
'Private Sub cmdEdit_Click()
'    If txt(19).Text = "" Then
'        With msg
'            .PassingMsg = "Nomor Rangka masih kosong, lakukan pembuatan (add) data pelanggan ?"
'            .passingTipe = 3
'            .Show vbModal
'        End With
'        Exit Sub
'    End If
'    Call SetControlForm("edit")
'End Sub
'
'Private Sub CmdDelete_Click()
'    With msg
'        .PassingMsg = "Hapus Master Pelanggan : " & txt(1).Text & ", No Rangka : " & txt(19).Text & " ?"
'        .passingTipe = 1
'        .Show vbModal
'        If .GetMsg = True Then
'            Call objP.Delete_ItemMasterPelanggan(txt(19).Text)
'            Call SetControlForm("load")
'        End If
'    End With
'End Sub
'
'Private Sub cmdOpen_Click()
'    '---------------------------------------------------------------------------------------------------------------------
'    '   02 SEPTEMBER 2015 |   REV. CEK ARMST & CUSTDAT (ANTISIPASI DOBEL KODELGN) |   ARIKDLG
'    '---------------------------------------------------------------------------------------------------------------------
'    '-----------------------------------------------------------------------------------------------------------------
'    '   CEK DI ARMST KODE & NAMA
'    '-----------------------------------------------------------------------------------------------------------------
'    Dim rsCPL As ADODB.Recordset
'    Set rsCPL = objP.FetchCekPelangganACC(Trim(txt(1).Text))
'    If rsCPL.RecordCount > 0 Then
'        If Trim(rsCPL("NamaLgn")) = Trim(txt(2).Text) Then
'            Set rsMove = objP.FetchListAllPelanggan(" Where KodeLg = '" & txt(1).Text & "' And NoRangka ='" & txt(19).Text & "'")
'            Unload Me
'        Else
'            With msg
'                .PassingMsg = "'" & txt(1).Text & "'" & Chr(13) & "Terdaftar Atas Nama : " & rsCPL("NamaLgn") & Chr(13) & "Silahkan ganti Kode Pelanggan!"
'                .passingTipe = 2
'                .Show vbModal
'            End With
'            txt(1).Text = ""
'            txt(1).SetFocus
'            Exit Sub
'        End If
'    End If
'End Sub

Private Sub cmdCancel_Click()
    Call SetControlForm("load")
End Sub

Private Sub cmdExit_Click()
   
    Unload Me
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txt(Index).BackColor = fColor Then
        txt(Index).SelStart = 0
        txt(Index).SelLength = Len(txt(Index).text)
        Clipboard.SetText txt(Index).text
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txt(Index).BackColor = fColor Then
        txt(Index).SelStart = 0
        txt(Index).SelLength = Len(txt(Index).text)
        Clipboard.SetText txt(Index).text
    End If
End Sub

'Private Sub Zom_Click()
'    onFind = True
'    onF12 = True
'    ft.ShowForm = "PLGN"
'    ft.Show vbModal
'    txt(1).SetFocus
'    onF12 = False
'End Sub

Private Function isValid() As Boolean
    isValid = True
    If txt(15) = "" Then txt(14) = ""
    If txt(18) = "" Then txt(17) = ""
    If txt(22) = "" Then txt(21) = ""
    If txt(24) = "" Then txt(23) = ""
        
    If txt(1).text = "" Or txt(2).text = "" Or txtNpwp.text = "" Or txt(4).text = "" Then
        isValid = False
        nValid = "[Nama/Alamat]"
        Exit Function
    End If
    If txt(7).text = "" And txt(8).text = "" Then
        isValid = False
        nValid = "[Telp Pelanggan]"
        Exit Function
    End If
    If txt(11).text = "" Or txt(13).text = "" _
    Or txt(21).text = "" Or txt(24).text = "" Or txt(26).text = "" Or txt(27).text = "" _
    Or txt(21).text = "" Then
        isValid = False
        nValid = "[Data Kendaraan]"
        Exit Function
    End If
    If txt(37).text = "" Or txt(38).text = "" Or txt(39).text = "" Then
        isValid = False
        nValid = "[Data Surat]"
        Exit Function
    End If
    If txt(23).text <> "" Then
        If txt(31).text = "" Or txt(32).text = "" Or txt(33).text = "" Or txt(36).text = "" _
        Or DtPolish0.text = "__/__/____" Or DtPolish1.text = "__/__/____" Then
            isValid = False
            nValid = "[Data Polis]"
            Exit Function
        End If
        If DtPolish0.value >= DtPolish1.value Then
            isValid = False
            nValid = "[Tanggal Polis]"
            Exit Function
        End If
    End If
    If txt(42).text = "" And txt(43).text = "" Then
        isValid = False
        nValid = "[Telp Surat]"
        Exit Function
    End If
    If Len(txt(19).text) < 9 Then
        isValid = False
        nValid = "[No Rangka Tidak Valid]"
        Exit Function
    End If
    Dim rsWx As ADODB.Recordset
    Set rsWx = objWl.FetchItemWilayah(Left(txt(1).text, 3))
    If rsWx.RecordCount = 0 Then
        isValid = False
        nValid = "[Kode Wilayah pada KODE LANGGAN, tidak terdaftar]"
        Exit Function
    End If
End Function

'20180313 REV'LONCAT JIKA KODELGN SUDAH ADA D ACC
Private Function LastID(tipe As String, Optional nilai As Integer = 1) As String
    Dim rsI As New ADODB.Recordset
    LastID = ""
    Set rsI = objSID.FetchLastID(tipe)
    If rsI.RecordCount > 0 Then
        rsI.MoveFirst
        LastID = rsI("Kode") & Format(rsI("LastUrut") + nilai, "0000000")
    End If
End Function


Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Dim rsCari As ADODB.Recordset
    
    If Index = 2 Then
 
    ElseIf Index = 20 Then
        If NulSID(txt(Index)) <> "" Then
            Dim rsAgama As ADODB.Recordset
            Set rsAgama = objO.FetchData("SELECT * FROM [SINGLEIDORI]..agama WITH (NOLOCK) WHERE agama = '" & NulSID(txt(Index)) & "'")
            If rsAgama.EOF Then
                txt(Index) = ""
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDAGAMA"
                frmCari.show (vbModal)
            End If
        End If
    ElseIf Index = 7 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..PEKERJAAN WITH (NOLOCK) WHERE pekerjaan = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                txt(Index) = ""
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDPEKERJAAN"
                frmCari.show (vbModal)
            End If
            If NulSID(UCase(txt(Index))) <> "LAINNYA" Then
                txt(23).BackColor = &HC0E0FF
                txt(23) = ""
            End If
        End If
    ElseIf Index = 21 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..HOBI WITH (NOLOCK) WHERE HOBI = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                txt(Index) = ""
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDHOBI"
                frmCari.show (vbModal)
            End If
        End If
    ElseIf Index = 22 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..OLAHRAGA WITH (NOLOCK) WHERE OLAHRAGA = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                txt(Index) = ""
                gMainKeyCari = txt(Index)
                gSearchOption = "SIDOLAHRAGA"
                frmCari.show (vbModal)
            End If
        End If
    ElseIf Index = 63 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
                
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 64 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(63)) & "' and KOTA = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 65 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(63)) & "' and KOTA = '" & NulSID(txt(64)) & "' and KECAMATAN = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 66 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(63)) & "' and KOTA = '" & NulSID(txt(64)) & "' and KECAMATAN = '" & NulSID(txt(65)) & "' and KELURAHAN = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 9 Or Index = 10 Then
        txt(Index) = Format(NulSID(txt(Index)), "0##")
    End If
End Sub


Private Sub prosesScoring()
         Dim vScore     As Double
         Dim vPembagi   As Double
         Dim vNilai     As Double
         Dim sFilter4   As String

1        On Error GoTo prosesScoring_Error
2        Frame3.Visible = True
3        frameNotif.Visible = False
4        Set rsFind = New ADODB.Recordset
5        CreateTblFind
6        CreateTblInput
7        isTambahNew = False
8        Set GridFind.DataSource = Nothing
         
9        With rsInput
10        .AddNew
11        !kodeLgn = NulSID(txt(1))
12        !tglentry = NulSID(isnow)
13        !userentry = NulSID(gSA)
14        !noKTP = NulSID(txt(2))
15        !namacust1 = NulSID(txt(4))
16        !namacust2 = NulSID(txt(5))
17        !namacust3 = NulSID(txt(6))
18        !TypeCustomer = NulSID(cboTipe.text)
19        !hp1 = NulSID(txt(14))
20        !hp2 = NulSID(txt(15))
21        !TempatLahir = NulSID("")
22        !tglLahir = IIf(NulSID(DtLahir.value) = "", "01/01/1900", NulSID(DtLahir.value)) 'rsFind!tglLahir
23        !Alamat = NulSID(txt(8))
24        !Propinsi = NulSID(txt(63))
25        !Kota = NulSID(txt(64))
26        !kecamatan = NulSID(txt(65))
27        !kelurahan = NulSID(txt(66))
28        !KodePos = NulSID(txt(11))
29        !score = 0
30        !addscore = 0
31        !TotalScore = 0
32        !Pilih = False
33        .Update
34       End With
         
35        If NulSID(txt(2)) = "" Or Val(NulSID(txt(2))) = "0" Then
36            If Label5(0).Caption = "No. KTP" Then
37                txt(2) = "0000000000000000"
38            End If
39        End If
40       If ProfilScoring.SearchingKTP = True And NulSID(txt(2)) <> "0000000000000000" Then
41           sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
42           sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
43           sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER "
44           sSQL = sSQL + "WHERE NoKTP = '" & NulSID(txt(2)) & "' "
45           sSQL = sSQL + "AND NoKTP not in (Select Kata from Exeption where Kolom = 'NoKTP') "
46           sSQL = sSQL + "ORDER BY KODELGN"
47           rsFind.Open sSQL, SQLConSID, adOpenKeyset
48           If rsFind.RecordCount <> 0 Then
49              isWajibKTP = True
                'lblJumlah = "Jumlah Data Dengan Keyword No. KTP: " & FormatNumber(rsFind.RecordCount, 0)
50              LblJumlah = "Jumlah ID Customer yang Identik : " & FormatNumber(rsFind.RecordCount, 0)
                
                'MsgBox "Pencarian dengan No KTP menemukan hasil.", vbInformation
51              MsgBox "Menemukan ID Customer Identik.", vbInformation
52                iProgress = 0
53                lblLoading.Visible = True
54                lblLoading.Caption = ""
55                ProgressBar1.Visible = True
56                ProgressBar1.Max = rsFind.RecordCount
                  
57                rsFind.MoveFirst
58                Do While Not rsFind.EOF
59                    rsTemp.AddNew
60                    rsTemp!kodeLgn = rsFind!kodeLgn
61                    rsTemp!tglentry = rsFind!tglentry
62                    rsTemp!userentry = Nul(rsFind!userentry)
63                    rsTemp!noKTP = Nul(rsFind!noKTP)
64                    rsTemp!namacust1 = Nul(rsFind!namacust1)
65                    rsTemp!namacust2 = Nul(rsFind!namacust2)
66                    rsTemp!namacust3 = Nul(rsFind!namacust3)
67                    rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
68                    rsTemp!hp1 = Nul(rsFind!hp1)
69                    rsTemp!hp2 = Nul(rsFind!hp2)
70                    rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
71                    rsTemp!tglLahir = rsFind!tglLahir
72                    rsTemp!Alamat = Nul(rsFind!Alamat)
73                    rsTemp!Propinsi = Nul(rsFind!Propinsi)
74                    rsTemp!Kota = Nul(rsFind!Kota)
75                    rsTemp!kecamatan = Nul(rsFind!kecamatan)
76                    rsTemp!kelurahan = Nul(rsFind!kelurahan)
77                    rsTemp!KodePos = Nul(rsFind!KodePos)
78                    rsTemp!score = 0
79                    rsTemp!addscore = 0
80                    rsTemp!TotalScore = 0
81                    rsTemp!Pilih = False
82                    rsTemp.Update
83                    rsFind.MoveNext
                      
84                    iProgress = iProgress + 1
85                    ProgressBar1.value = iProgress
86                    DoEvents
                      
87                    lblLoading.Caption = "Data " & iProgress & " dari " & rsFind.RecordCount
                      
88                Loop
                  
89                rsTemp.MoveFirst
90                rsTemp!Pilih = True
                  
91                filterRsTemp = ""
92                Set GridFind.DataSource = rsTemp
                  
93                lblLoading.Visible = False
94                lblLoading.Caption = ""
95                ProgressBar1.Visible = False
96                cmdClear.Enabled = False
97                Exit Sub
98           End If
99       End If

CariNama:
100      If NulSID(txt(4)) <> "" And NulSID(txt(5)) = "" And NulSID(txt(6)) = "" Then
              '' 20190115
              'Call prosesScoringTerakhir
101           GoTo CariNoHP
102           Exit Sub
103      End If
104      If ProfilScoring.SearchingNama = True Then
105          Set rsFind = New ADODB.Recordset
106          sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
107          sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
108          sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER "
109          sSQL = sSQL + "WHERE NamaCust1 = '" & NulSID(txt(4)) & "' AND NamaCust2 = '" & NulSID(txt(5)) & "' "
110          sSQL = sSQL + "ORDER BY KODELGN"
111          rsFind.Open sSQL, SQLConSID, adOpenKeyset
112          If rsFind.RecordCount <> 0 Then
113             isWajibKTP = True
                'lblJumlah = "Jumlah Data Dengan Keyword Nama Identik: " & FormatNumber(rsFind.RecordCount, 0)
114             LblJumlah = "Jumlah ID Customer yang Identik : " & FormatNumber(rsFind.RecordCount, 0)
      '          If MsgBox("Pencarian dengan Nama Identik menemukan hasil, pergunakan data ini?", vbYesNo + vbExclamation) = vbNo Then
      '             GoTo ScoringNama1Khusus
      '          Else
                   'MsgBox "Pencarian dengan Nama Identik menemukan hasil.", vbInformation
115                MsgBox "Menemukan ID Customer Identik.", vbInformation
116                rsFind.MoveFirst
117                iProgress = 0
118                lblLoading.Visible = True
119                lblLoading.Caption = ""
120                ProgressBar1.Visible = True
121                ProgressBar1.Max = rsFind.RecordCount
                  
122                Do While Not rsFind.EOF
123                   rsTemp.AddNew
124                   rsTemp!kodeLgn = rsFind!kodeLgn
125                   rsTemp!tglentry = rsFind!tglentry
126                   rsTemp!userentry = Nul(rsFind!userentry)
127                   rsTemp!noKTP = Nul(rsFind!noKTP)
128                   rsTemp!namacust1 = Nul(rsFind!namacust1)
129                   rsTemp!namacust2 = Nul(rsFind!namacust2)
130                   rsTemp!namacust3 = Nul(rsFind!namacust3)
131                   rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
132                   rsTemp!hp1 = Nul(rsFind!hp1)
133                   rsTemp!hp2 = Nul(rsFind!hp2)
134                   rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
135                   rsTemp!tglLahir = rsFind!tglLahir
136                   rsTemp!Alamat = Nul(rsFind!Alamat)
137                   rsTemp!Propinsi = Nul(rsFind!Propinsi)
138                   rsTemp!Kota = Nul(rsFind!Kota)
139                   rsTemp!kecamatan = Nul(rsFind!kecamatan)
140                   rsTemp!kelurahan = Nul(rsFind!kelurahan)
141                   rsTemp!KodePos = Nul(rsFind!KodePos)
142                   rsTemp!score = 0
143                   rsTemp!addscore = 0
144                   rsTemp!TotalScore = 0
145                   rsTemp!Pilih = False
146                   rsTemp.Update
147                   rsFind.MoveNext
                      
148                   iProgress = iProgress + 1
149                   ProgressBar1.value = iProgress
150                   DoEvents
                      
151                   lblLoading.Caption = "Data " & iProgress & " dari " & rsFind.RecordCount
                      
152                Loop
                   
153                rsTemp.MoveFirst
154                rsTemp!Pilih = True
155                filterRsTemp = ""
                   
156                Set GridFind.DataSource = rsTemp
                   
157                lblLoading.Visible = False
158                lblLoading.Caption = ""
159                ProgressBar1.Visible = False
160                cmdClear.Enabled = True
161                Exit Sub
162                Exit Sub
      '          End If
163          End If
164      End If
         
CariNoHP:
165      If ProfilScoring.SearchingHP = True And NulSID(txt(14)) <> "" Then
166          Set rsFind = New ADODB.Recordset
167          sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
168          sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
169          sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER "
170          sSQL = sSQL + "WHERE HP1 = '" & NulSID(txt(14)) & "' "
171          sSQL = sSQL + "AND HP1 not in (Select Kata from Exeption where Kolom = 'NoHP') "
172          sSQL = sSQL + "ORDER BY KODELGN"
173          rsFind.Open sSQL, SQLConSID, adOpenKeyset
174          If rsFind.RecordCount <> 0 Then
175            isWajibKTP = True
176             LblJumlah = "Jumlah Data Dengan Keyword No. HP: " & FormatNumber(rsFind.RecordCount, 0)
               ' If MsgBox("Pencarian dengan No HP menemukan hasil, pergunakan data ini?", vbYesNo + vbExclamation) = vbNo Then
               '    GoTo CariNama
               ' Else
               'MsgBox "Pencarian dengan No HP menemukan hasil.", vbInformation
177            MsgBox "Menemukan ID Customer Identik.", vbInformation
178                iProgress = 0
179                lblLoading.Visible = True
180                lblLoading.Caption = ""
181                ProgressBar1.Visible = True
182                ProgressBar1.Max = rsFind.RecordCount
                   
183                rsFind.MoveFirst
184                Do While Not rsFind.EOF
185                   rsTemp.AddNew
186                   rsTemp!kodeLgn = rsFind!kodeLgn
187                   rsTemp!tglentry = rsFind!tglentry
188                   rsTemp!userentry = Nul(rsFind!userentry)
189                   rsTemp!noKTP = Nul(rsFind!noKTP)
190                   rsTemp!namacust1 = Nul(rsFind!namacust1)
191                   rsTemp!namacust2 = Nul(rsFind!namacust2)
192                   rsTemp!namacust3 = Nul(rsFind!namacust3)
193                   rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
194                   rsTemp!hp1 = Nul(rsFind!hp1)
195                   rsTemp!hp2 = Nul(rsFind!hp2)
196                   rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
197                   rsTemp!tglLahir = rsFind!tglLahir
198                   rsTemp!Alamat = Nul(rsFind!Alamat)
199                   rsTemp!Propinsi = Nul(rsFind!Propinsi)
200                   rsTemp!Kota = Nul(rsFind!Kota)
201                   rsTemp!kecamatan = Nul(rsFind!kecamatan)
202                   rsTemp!kelurahan = Nul(rsFind!kelurahan)
203                   rsTemp!KodePos = Nul(rsFind!KodePos)
204                   rsTemp!score = 0
205                   rsTemp!addscore = 0
206                   rsTemp!TotalScore = 0
207                   rsTemp!Pilih = False
208                   rsTemp.Update
209                   rsFind.MoveNext
                       
                      
210                   iProgress = iProgress + 1
211                   ProgressBar1.value = iProgress
212                   DoEvents
                      
213                   lblLoading.Caption = "Data " & iProgress & " dari " & rsFind.RecordCount
                      
214                Loop
215                rsTemp.MoveFirst
216                rsTemp!Pilih = True
217                filterRsTemp = ""
218                Set GridFind.DataSource = rsTemp
                   
219                lblLoading.Visible = False
220                lblLoading.Caption = ""
221                ProgressBar1.Visible = False
222                cmdClear.Enabled = True
223                Exit Sub
                'End If
224          End If
225      End If
226      If rsTemp.RecordCount = 0 Then
227           prosesScoringTerakhir
228      End If

229      On Error GoTo 0
230      Exit Sub

prosesScoring_Error:

231       MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure prosesScoring of Form frmSIDMasterPelangganNonMerge"
End Sub
 
Private Sub prosesScoringTerakhir()
             
1        On Error GoTo prosesScoringTerakhir_Error
2         cmdClear.Enabled = False
3         Set GridFind.DataSource = Nothing
ScoringNama1Khusus:
4        If ProfilScoring.ScoringNamaKhusus = True Then
5           If NulSID(txt(4)) <> "" And NulSID(txt(5)) = "" And NulSID(txt(6)) = "" Then
6               Set rsFind = New ADODB.Recordset
7               sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
8               sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
9               sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER WITH (NOLOCK) "
10              sSQL = sSQL + "WHERE Left(NamaCust1," & ProfilScoring.LockNama1Khusus & ") = '" & Left(NulSID(txt(4)), ProfilScoring.LockNama1Khusus) & "' "
11              sSQL = sSQL + "ORDER BY KODELGN"
12              rsFind.Open sSQL, SQLConSID, adOpenKeyset
13              Call CreateTblFind
14              DoEvents
15              If rsFind.RecordCount <> 0 Then
16                 rsFind.MoveFirst
17                 iProgress = 0
18                 lblLoading.Visible = True
19                 lblLoading.Caption = ""
20                 ProgressBar1.Visible = True
21                 ProgressBar1.Max = rsFind.RecordCount
              
22                 DoEvents
23                 lblLoading.Caption = "Data " & iProgress & " dari " & rsFind.RecordCount
                      
24                 Do While Not rsFind.EOF
25                    vScore = 0
26                    vPembagi = 0
27                    vNilai = 0
28                    rsTemp.AddNew
29                    rsTemp!kodeLgn = rsFind!kodeLgn
30                    rsTemp!tglentry = rsFind!tglentry
31                    rsTemp!userentry = Nul(rsFind!userentry)
32                    rsTemp!noKTP = Nul(rsFind!noKTP)
33                    rsTemp!namacust1 = Nul(rsFind!namacust1)
34                    rsTemp!namacust2 = Nul(rsFind!namacust2)
35                    rsTemp!namacust3 = Nul(rsFind!namacust3)
36                    rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
37                    rsTemp!hp1 = Nul(rsFind!hp1)
38                    rsTemp!hp2 = Nul(rsFind!hp2)
39                    rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
40                    rsTemp!tglLahir = rsFind!tglLahir
41                    rsTemp!Alamat = Nul(rsFind!Alamat)
42                    rsTemp!Propinsi = Nul(rsFind!Propinsi)
43                    rsTemp!Kota = Nul(rsFind!Kota)
44                    rsTemp!kecamatan = Nul(rsFind!kecamatan)
45                    rsTemp!kelurahan = Nul(rsFind!kelurahan)
46                    rsTemp!KodePos = Nul(rsFind!KodePos)

47                    If NulSID(txt(4)) <> "" And Nul(rsFind!namacust1) <> "" Then
48                        vPembagi = (Len(rsFind!namacust1) + Len(NulSID(txt(4)))) / 2
49                        For X = 1 To Len(NulSID(txt(4)))
50                            vNilai = vNilai + IIf(Mid$(LCase(NulSID(txt(4))), X, 1) = Mid$(LCase(rsFind!namacust1), X, 1), 1, 0)
51                        Next X
52                        vScore = vNilai / vPembagi
53                    End If
54                    rsTemp!score = vScore
55                    rsTemp!addscore = 0
56                    rsTemp!TotalScore = 0
57                    If (ProfilScoring.ScoringNoHP = True) And (Nul(Nul(rsFind!hp1)) = NulSID(txt(14))) Then rsTemp!addscore = rsTemp!addscore + 0.25
58                    If (ProfilScoring.ScoringTglLahir = True) And (rsFind!tglLahir = DtLahir.value) Then rsTemp!addscore = rsTemp!addscore + 0.25
59                    If (ProfilScoring.ScoringKelurahan = True) And (NulSID(rsFind!kelurahan) = NulSID(txt(66))) Then rsTemp!addscore = rsTemp!addscore + 0.25
60                    If (ProfilScoring.ScoringKecamatan = True) And (NulSID(rsFind!kecamatan) = NulSID(txt(65))) Then rsTemp!addscore = rsTemp!addscore + 0.25
61                    rsTemp!TotalScore = rsTemp!score + rsTemp!addscore
62                    rsTemp!Pilih = False
63                    rsTemp.Update
64                    rsFind.MoveNext
                      
65                    lblLoading.Caption = " " & Round((iProgress * 100) / rsFind.RecordCount, 2) & " %" '"Data " & iProgress & " dari " & rsFind.RecordCount
66                    DoEvents
67                    iProgress = iProgress + 1
68                    ProgressBar1.value = iProgress
69                    DoEvents
70                 Loop
                   
71                 lblLoading.Visible = False
72                 lblLoading.Caption = ""
73                 ProgressBar1.Visible = False
74              End If
75              rsTemp.filter = "addScore > 0 and score > " & k2t(ProfilScoring.nScoringNamaKhusus)
                'rsTemp.filter = "Score > " & k2t(ProfilScoring.nScoringNamaKhusus)
76              filterRsTemp = rsTemp.filter
77              rsTemp.Sort = "TotalScore desc, Score desc, AddScore desc"
78              LblJumlah = "Jumlah Data Dengan  Nama 1 Khusus: " & FormatNumber(rsTemp.RecordCount, 0)
79              If rsTemp.RecordCount <> 0 Then
80                  isWajibKTP = True
                    
      '              If MsgBox("Pencarian dengan Nama 1 Khusus menemukan hasil, pergunakan data ini?", vbYesNo + vbExclamation) = vbNo Then
      '                 GoTo ScoringNama2
      '              Else
      '                 Exit Sub
      '              End If
                  'MsgBox "Pencarian dengan Nama 1 Khusus menemukan hasil " & rsTemp.RecordCount, vbInformation
81                MsgBox "Menemukan ID Customer Identik.", vbInformation
                  
82                rsTemp.MoveFirst
83                rsTemp!Pilih = True
84                Set GridFind.DataSource = rsTemp
85                Exit Sub
                    
86              End If
87              Exit Sub
88          End If
89       End If

ScoringNama2:
90      If ProfilScoring.ScoringNamaSimultan = True Then
91              Set rsFind = New ADODB.Recordset
92              sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
93              sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
94              sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER WITH (NOLOCK) "
95              sSQL = sSQL + "WHERE Left(NamaCust1," & ProfilScoring.LockNama1 & ") = '" & Left(NulSID(txt(4)), ProfilScoring.LockNama1) & "' "
96              sSQL = sSQL + "AND Left(NamaCust2," & ProfilScoring.LockNama2 & ") = '" & Left(NulSID(txt(5)), ProfilScoring.LockNama2) & "' "
97              sSQL = sSQL + "ORDER BY KODELGN"
98              rsFind.Open sSQL, SQLConSID, adOpenKeyset
                
99              DoEvents
100             Call CreateTblFind
101             iProgress = 0
102             lblLoading.Visible = True
103             lblLoading.Caption = ""
104             ProgressBar1.Visible = True
105             ProgressBar1.Max = rsFind.RecordCount
106             If rsFind.RecordCount <> 0 Then
107                rsFind.MoveFirst
108                Do While Not rsFind.EOF
109                   vScore = 0
110                   vPembagi = 0
111                   vNilai = 0
112                   rsTemp.AddNew
113                   rsTemp!kodeLgn = rsFind!kodeLgn
114                   rsTemp!tglentry = rsFind!tglentry
115                   rsTemp!userentry = Nul(rsFind!userentry)
116                   rsTemp!noKTP = Nul(rsFind!noKTP)
117                   rsTemp!namacust1 = Nul(rsFind!namacust1)
118                   rsTemp!namacust2 = Nul(rsFind!namacust2)
119                   rsTemp!namacust3 = Nul(rsFind!namacust3)
120                   rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
121                   rsTemp!hp1 = Nul(rsFind!hp1)
122                   rsTemp!hp2 = Nul(rsFind!hp2)
123                   rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
124                   rsTemp!tglLahir = rsFind!tglLahir
125                   rsTemp!Alamat = Nul(rsFind!Alamat)
126                   rsTemp!Propinsi = Nul(rsFind!Propinsi)
127                   rsTemp!Kota = Nul(rsFind!Kota)
128                   rsTemp!kecamatan = Nul(rsFind!kecamatan)
129                   rsTemp!kelurahan = Nul(rsFind!kelurahan)
130                   rsTemp!KodePos = Nul(rsFind!KodePos)

131                   If NulSID(txt(5)) <> "" And Nul(rsFind!namacust2) <> "" Then
132                       vPembagi = (Len(rsFind!namacust2) + Len(NulSID(txt(5)))) / 2
133                       For X = 1 To Len(NulSID(txt(5)))
134                           vNilai = vNilai + IIf(Mid$(LCase(NulSID(txt(5))), X, 1) = Mid$(LCase(rsFind!namacust2), X, 1), 1, 0)
135                       Next X
136                       vScore = vNilai / vPembagi
137                   End If
138                   rsTemp!score = vScore
139                   rsTemp!addscore = 0
140                   rsTemp!TotalScore = 0
141                   If (ProfilScoring.ScoringNoHP = True) And (Nul(Nul(rsFind!hp1)) = NulSID(txt(14))) Then rsTemp!addscore = rsTemp!addscore + 0.25
142                   If (ProfilScoring.ScoringTglLahir = True) And (rsFind!tglLahir = DtLahir.value) Then rsTemp!addscore = rsTemp!addscore + 0.25
143                   If (ProfilScoring.ScoringKelurahan = True) And (NulSID(rsFind!kelurahan) = NulSID(txt(66))) Then rsTemp!addscore = rsTemp!addscore + 0.25
144                   If (ProfilScoring.ScoringKecamatan = True) And (NulSID(rsFind!kecamatan) = NulSID(txt(65))) Then rsTemp!addscore = rsTemp!addscore + 0.25
145                   rsTemp!TotalScore = rsTemp!score + rsTemp!addscore
146                   rsTemp!Pilih = False
147                   rsTemp.Update
148                   rsFind.MoveNext
                      
149                   iProgress = iProgress + 1
150                   ProgressBar1.value = iProgress
151                   DoEvents
152                Loop
153                lblLoading.Visible = False
154                lblLoading.Caption = ""
155                ProgressBar1.Visible = False
156             End If
157             rsTemp.filter = "addScore > 0 and Score > " & k2t(ProfilScoring.nScoringNama2)
158             filterRsTemp = rsTemp.filter
159             rsTemp.Sort = "TotalScore desc, Score desc, AddScore desc"
160             LblJumlah = "Jumlah Data Dengan  Nama Simultan: " & FormatNumber(rsTemp.RecordCount, 0)
161             If rsTemp.RecordCount <> 0 Then
162                 isWajibKTP = True
                    'MsgBox "Filter dengan batas Nama 2 " & rsTemp.filter & " menghasilkan " & rsTemp.RecordCount & " data", vbInformation, "Scoring Nama 2"
163                 MsgBox "Menemukan ID Customer Identik.", vbInformation
                    
164                 rsTemp.MoveFirst
165                 rsTemp!Pilih = True
166             Else
                    'MsgBox "Filter dengan batas Nama 2 " & rsTemp.filter & " tidak menghasilkan data!", vbExclamation, "Scoring Nama 2"
167                 MsgBox "Tidak menemukan ID Customer Identik.", vbInformation
168             End If
169             Set GridFind.DataSource = rsTemp
170     End If

171      On Error GoTo 0
172      Exit Sub

prosesScoringTerakhir_Error:

173       MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure prosesScoringTerakhir of Form frmSIDMasterPelangganNonMerge"
       
       
End Sub

Private Sub CreateTblFind(Optional Edit As Boolean = False)
On Error GoTo vb_error
 Set rsTemp = New ADODB.Recordset
 If rsTemp.State Then rsTemp.Close
 With rsTemp
  .Fields.Append "Pilih", adBoolean
  .Fields.Append "Score", adDouble
  .Fields.Append "AddScore", adDouble
  .Fields.Append "TotalScore", adDouble
  .Fields.Append "KODELGN", adVarChar, 13
  .Fields.Append "TglEntry", adDate
  .Fields.Append "UserEntry", adVarChar, 50
  .Fields.Append "NoKTP", adVarChar, 30
  .Fields.Append "NamaCust1", adVarChar, 300
  .Fields.Append "NamaCust2", adVarChar, 300
  .Fields.Append "NamaCust3", adVarChar, 300
  .Fields.Append "TypeCustomer", adVarChar, 50
  .Fields.Append "HP1", adVarChar, 30
  .Fields.Append "HP2", adVarChar, 30
  .Fields.Append "TempatLahir", adVarChar, 100
  .Fields.Append "TglLahir", adDate
  .Fields.Append "Alamat", adVarChar, 300
  .Fields.Append "Propinsi", adVarChar, 100
  .Fields.Append "Kota", adVarChar, 100
  .Fields.Append "Kecamatan", adVarChar, 100
  .Fields.Append "Kelurahan", adVarChar, 100
  .Fields.Append "KodePos", adVarChar, 30
  .Open
'  Set GridFind.DataSource = rsTemp
  End With
Exit Sub
vb_error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CreateTblFind"), vbExclamation, "Err Number : " & Erl
End Sub

Private Sub CreateTblInput(Optional Edit As Boolean = False)
On Error GoTo vb_error
 Set rsInput = New ADODB.Recordset
 With rsInput
  .Fields.Append "Pilih", adBoolean
  .Fields.Append "Score", adDouble
  .Fields.Append "AddScore", adDouble
  .Fields.Append "TotalScore", adDouble
  .Fields.Append "KODELGN", adVarChar, 13
  .Fields.Append "TglEntry", adDate
  .Fields.Append "UserEntry", adVarChar, 50
  .Fields.Append "NoKTP", adVarChar, 30
  .Fields.Append "NamaCust1", adVarChar, 300
  .Fields.Append "NamaCust2", adVarChar, 300
  .Fields.Append "NamaCust3", adVarChar, 300
  .Fields.Append "TypeCustomer", adVarChar, 50
  .Fields.Append "HP1", adVarChar, 30
  .Fields.Append "HP2", adVarChar, 30
  .Fields.Append "TempatLahir", adVarChar, 100
  .Fields.Append "TglLahir", adDate
  .Fields.Append "Alamat", adVarChar, 300
  .Fields.Append "Propinsi", adVarChar, 100
  .Fields.Append "Kota", adVarChar, 100
  .Fields.Append "Kecamatan", adVarChar, 100
  .Fields.Append "Kelurahan", adVarChar, 100
  .Fields.Append "KodePos", adVarChar, 30
  .Open
  Set gridInput.DataSource = rsInput
  End With
Exit Sub
vb_error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CreateTblFind"), vbExclamation, "Err Number : " & Erl
End Sub

Private Sub XPButton1_Click()
    Frame3.Visible = False
End Sub

Private Sub GridFind_FilterChange()
    On Error GoTo errHandler
    Set cols = GridFind.Columns
    Dim C As Integer
    Dim filter As String
    filter = filterRsTemp 'rsTemp.filter
    
    C = GridFind.col
    GridFind.HoldFields
    If filter = "" Then
        rsTemp.filter = getFilter()
    Else
        If getFilter() = "" Then
            rsTemp.filter = filter
        Else
            rsTemp.filter = filter & " AND " & getFilter()
        End If
        
    End If
    
    GridFind.col = C
    GridFind.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        
        End If
    Next col
    getFilter = tmp
End Function

Function InstrCount(StringToSearch As String, _
           StringToFind As String) As Long

    If Len(StringToFind) Then
        InstrCount = UBound(Split(StringToSearch, StringToFind))
    End If
End Function

Public Function tBaca(dat As String) As String
    Dim z As Integer
    tBaca = ""
    If dat = "" Then Exit Function
    z = Len(Trim(dat))
    For C = 1 To z
        If IsNumeric(Mid(dat, C, 1)) = True Then
            If Mid(dat, C, 1) = "." Or Mid(dat, C, 1) = "," Or Mid(dat, C, 1) = "-" Then
                tBaca = Mid(dat, C, 1)
                Exit For
            End If
        Else
            tBaca = Mid(dat, C, 1)
            Exit For
        End If
    Next
End Function

Public Function digitsama(dat As String, jum As Integer) As Double
    Dim z As Integer, cari As String, hasil As Integer
    digitsama = 0: cari = "": hasil = 1
    If dat = "" Then Exit Function
    z = Len(Trim(dat))
    For C = 1 To z
        If cari = "" Then
            cari = Mid(dat, C, 1)
        Else
            If cari = Mid(dat, C, 1) Then
                If C = 1 Then
                    GoTo lanjut
                Else
                    hasil = hasil + 1
                    If hasil > jum Then
                        digitsama = hasil
                        GoTo finish
                    End If
                End If
            Else
                hasil = 1
                cari = Mid(dat, C, 1)
            End If
        End If
lanjut:
    Next
finish:
    If hasil > jum Then digitsama = hasil
End Function


