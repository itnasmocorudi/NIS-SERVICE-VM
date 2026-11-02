VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form FrmTwc_ViewCabang 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameTombol 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   14895
      Begin ucXPButton.XPButton cmdClose 
         Height          =   495
         Left            =   13560
         TabIndex        =   1
         Top             =   0
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
         MICON           =   "FrmTwc_ViewCabang.frx":0000
         PICN            =   "FrmTwc_ViewCabang.frx":001C
         PICH            =   "FrmTwc_ViewCabang.frx":04F0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdExport 
         Height          =   495
         Left            =   12240
         TabIndex        =   2
         Top             =   0
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
         MICON           =   "FrmTwc_ViewCabang.frx":0A3B
         PICN            =   "FrmTwc_ViewCabang.frx":0A57
         PICH            =   "FrmTwc_ViewCabang.frx":0FEC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdView 
         Height          =   495
         Left            =   11040
         TabIndex        =   3
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&v"
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
         MICON           =   "FrmTwc_ViewCabang.frx":1627
         PICN            =   "FrmTwc_ViewCabang.frx":1643
         PICH            =   "FrmTwc_ViewCabang.frx":1B05
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "2760"
      Height          =   10425
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14895
      Begin TrueOleDBGrid70.TDBGrid GridRekap 
         Height          =   6520
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   14400
         _ExtentX        =   25400
         _ExtentY        =   11509
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No. PTC"
         Columns(0).DataField=   "Nobukti"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No. Urut"
         Columns(1).DataField=   "Nobaris"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No. W. O."
         Columns(2).DataField=   "Nowo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Claim No."
         Columns(3).DataField=   "ClaimNo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "No. Rangka"
         Columns(4).DataField=   "NoRangka"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Ope. No."
         Columns(5).DataField=   "OpeNo"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "O. F. P."
         Columns(6).DataField=   "Ofp"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Kodelgn"
         Columns(7).DataField=   "Kodelgn"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Customer"
         Columns(8).DataField=   "Namalgn"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Tipe Mobil"
         Columns(9).DataField=   "Fmdl"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Warna"
         Columns(10).DataField=   "Color"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "No. Polisi"
         Columns(11).DataField=   "No_Polisi"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "No. Invoice"
         Columns(12).DataField=   "Noinvoice"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Tgl. Invoice"
         Columns(13).DataField=   "Invdate"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Hour"
         Columns(14).DataField=   "Hour"
         Columns(14).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "Jasa"
         Columns(15).DataField=   "Jasa"
         Columns(15).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "Sublet"
         Columns(16).DataField=   "Sublet"
         Columns(16).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Part"
         Columns(17).DataField=   "Part"
         Columns(17).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).Caption=   "Sub Total"
         Columns(18).DataField=   "Subtotal"
         Columns(18).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).Caption=   "Pajak"
         Columns(19).DataField=   "Pajak"
         Columns(19).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).Caption=   "Lain-Lain"
         Columns(20).DataField=   "Lain"
         Columns(20).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).Caption=   "Tot. Pengajuan"
         Columns(21).DataField=   "Total"
         Columns(21).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(22)._VlistStyle=   0
         Columns(22)._MaxComboItems=   5
         Columns(22).Caption=   "Disetujui"
         Columns(22).DataField=   "Acc"
         Columns(22).NumberFormat=   "True/False"
         Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(23)._VlistStyle=   0
         Columns(23)._MaxComboItems=   5
         Columns(23).Caption=   "Tot. Disetujui"
         Columns(23).DataField=   "TotAcc"
         Columns(23).NumberFormat=   "#,##0.#0;(#,##0.#0)"
         Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(24)._VlistStyle=   0
         Columns(24)._MaxComboItems=   5
         Columns(24).Caption=   "Keterangan"
         Columns(24).DataField=   "Keterangan"
         Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(25)._VlistStyle=   0
         Columns(25)._MaxComboItems=   5
         Columns(25).Caption=   "Proses Finance"
         Columns(25).DataField=   "Finance"
         Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   26
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).AllowSizing=   -1  'True
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   8421376
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=26"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4233"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4154"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(0)._FootDivider=0"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=1323"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1244"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=1693"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1614"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1693"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1614"
         Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(23)=   "Column(4).Width=4180"
         Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=4101"
         Splits(0)._ColumnProps(26)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=1693"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=1614"
         Splits(0)._ColumnProps(32)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._ColStyle=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(35)=   "Column(6).Width=1773"
         Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=1693"
         Splits(0)._ColumnProps(38)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(40)=   "Column(7).Width=1826"
         Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=1746"
         Splits(0)._ColumnProps(43)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(45)=   "Column(8).Width=3413"
         Splits(0)._ColumnProps(46)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(8)._WidthInPix=3334"
         Splits(0)._ColumnProps(48)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(50)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(53)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(55)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(56)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(58)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(60)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(61)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(62)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(63)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(64)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(65)=   "Column(12).Width=1931"
         Splits(0)._ColumnProps(66)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(12)._WidthInPix=1852"
         Splits(0)._ColumnProps(68)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(69)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(70)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(71)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(72)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(73)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(74)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(75)=   "Column(14).Width=926"
         Splits(0)._ColumnProps(76)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(77)=   "Column(14)._WidthInPix=847"
         Splits(0)._ColumnProps(78)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(79)=   "Column(14)._ColStyle=2"
         Splits(0)._ColumnProps(80)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(81)=   "Column(15).Width=2646"
         Splits(0)._ColumnProps(82)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(15)._WidthInPix=2566"
         Splits(0)._ColumnProps(84)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(85)=   "Column(15)._ColStyle=2"
         Splits(0)._ColumnProps(86)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(87)=   "Column(16).Width=2646"
         Splits(0)._ColumnProps(88)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(89)=   "Column(16)._WidthInPix=2566"
         Splits(0)._ColumnProps(90)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(91)=   "Column(16)._ColStyle=2"
         Splits(0)._ColumnProps(92)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(93)=   "Column(17).Width=2646"
         Splits(0)._ColumnProps(94)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(95)=   "Column(17)._WidthInPix=2566"
         Splits(0)._ColumnProps(96)=   "Column(17)._EditAlways=0"
         Splits(0)._ColumnProps(97)=   "Column(17)._ColStyle=2"
         Splits(0)._ColumnProps(98)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(99)=   "Column(18).Width=2646"
         Splits(0)._ColumnProps(100)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(101)=   "Column(18)._WidthInPix=2566"
         Splits(0)._ColumnProps(102)=   "Column(18)._EditAlways=0"
         Splits(0)._ColumnProps(103)=   "Column(18)._ColStyle=2"
         Splits(0)._ColumnProps(104)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(105)=   "Column(19).Width=2646"
         Splits(0)._ColumnProps(106)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(107)=   "Column(19)._WidthInPix=2566"
         Splits(0)._ColumnProps(108)=   "Column(19)._EditAlways=0"
         Splits(0)._ColumnProps(109)=   "Column(19)._ColStyle=2"
         Splits(0)._ColumnProps(110)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(111)=   "Column(20).Width=2646"
         Splits(0)._ColumnProps(112)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(113)=   "Column(20)._WidthInPix=2566"
         Splits(0)._ColumnProps(114)=   "Column(20)._EditAlways=0"
         Splits(0)._ColumnProps(115)=   "Column(20)._ColStyle=2"
         Splits(0)._ColumnProps(116)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(117)=   "Column(21).Width=2646"
         Splits(0)._ColumnProps(118)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(119)=   "Column(21)._WidthInPix=2566"
         Splits(0)._ColumnProps(120)=   "Column(21)._EditAlways=0"
         Splits(0)._ColumnProps(121)=   "Column(21)._ColStyle=2"
         Splits(0)._ColumnProps(122)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(123)=   "Column(22).Width=1640"
         Splits(0)._ColumnProps(124)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(125)=   "Column(22)._WidthInPix=1561"
         Splits(0)._ColumnProps(126)=   "Column(22)._EditAlways=0"
         Splits(0)._ColumnProps(127)=   "Column(22)._ColStyle=1"
         Splits(0)._ColumnProps(128)=   "Column(22).Order=23"
         Splits(0)._ColumnProps(129)=   "Column(23).Width=2646"
         Splits(0)._ColumnProps(130)=   "Column(23).DividerColor=0"
         Splits(0)._ColumnProps(131)=   "Column(23)._WidthInPix=2566"
         Splits(0)._ColumnProps(132)=   "Column(23)._EditAlways=0"
         Splits(0)._ColumnProps(133)=   "Column(23)._ColStyle=2"
         Splits(0)._ColumnProps(134)=   "Column(23).Order=24"
         Splits(0)._ColumnProps(135)=   "Column(24).Width=9737"
         Splits(0)._ColumnProps(136)=   "Column(24).DividerColor=0"
         Splits(0)._ColumnProps(137)=   "Column(24)._WidthInPix=9657"
         Splits(0)._ColumnProps(138)=   "Column(24)._EditAlways=0"
         Splits(0)._ColumnProps(139)=   "Column(24)._ColStyle=0"
         Splits(0)._ColumnProps(140)=   "Column(24).Order=25"
         Splits(0)._ColumnProps(141)=   "Column(25).Width=2725"
         Splits(0)._ColumnProps(142)=   "Column(25).DividerColor=0"
         Splits(0)._ColumnProps(143)=   "Column(25)._WidthInPix=2646"
         Splits(0)._ColumnProps(144)=   "Column(25)._EditAlways=0"
         Splits(0)._ColumnProps(145)=   "Column(25).Order=26"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Arial"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         ColumnFooters   =   -1  'True
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         WrapCellPointer =   -1  'True
         MultipleLines   =   0
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Arial"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.borderColor=&H80000013&"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.fgcolor=&H0&"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80FFFF&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&H80000005&,.fgcolor=&H0&,.bold=0"
         _StyleDefs(26)  =   ":id=13,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(27)  =   ":id=13,.fontname=Arial"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=900,.italic=0"
         _StyleDefs(30)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(31)  =   ":id=14,.fontname=Arial"
         _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.borderColor=&H80FF80&,.bold=0"
         _StyleDefs(35)  =   ":id=18,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(36)  =   ":id=18,.fontname=MS Sans Serif"
         _StyleDefs(37)  =   "Splits(0).EditorStyle:id=17,.parent=7,.namedParent=33,.bgcolor=&HFFFFFF&,.bold=0"
         _StyleDefs(38)  =   ":id=17,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(39)  =   ":id=17,.fontname=Arial"
         _StyleDefs(40)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.namedParent=38"
         _StyleDefs(41)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF&"
         _StyleDefs(42)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.namedParent=35,.bgcolor=&H80FFFF&"
         _StyleDefs(43)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(44)  =   "Splits(0).FilterBarStyle:id=24,.parent=12,.bgcolor=&HFFFFFF&,.fgcolor=&H0&"
         _StyleDefs(45)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.locked=0"
         _StyleDefs(46)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14,.bold=-1,.fontsize=900"
         _StyleDefs(47)  =   ":id=51,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(48)  =   ":id=51,.fontname=Arial"
         _StyleDefs(49)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15,.bgcolor=&HFFFFFF&"
         _StyleDefs(50)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(51)  =   "Splits(0).Columns(1).Style:id=130,.parent=13"
         _StyleDefs(52)  =   "Splits(0).Columns(1).HeadingStyle:id=127,.parent=14"
         _StyleDefs(53)  =   "Splits(0).Columns(1).FooterStyle:id=128,.parent=15"
         _StyleDefs(54)  =   "Splits(0).Columns(1).EditorStyle:id=129,.parent=17"
         _StyleDefs(55)  =   "Splits(0).Columns(2).Style:id=118,.parent=13"
         _StyleDefs(56)  =   "Splits(0).Columns(2).HeadingStyle:id=115,.parent=14"
         _StyleDefs(57)  =   "Splits(0).Columns(2).FooterStyle:id=116,.parent=15"
         _StyleDefs(58)  =   "Splits(0).Columns(2).EditorStyle:id=117,.parent=17"
         _StyleDefs(59)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=0"
         _StyleDefs(60)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14,.bold=-1,.fontsize=900"
         _StyleDefs(61)  =   ":id=63,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(62)  =   ":id=63,.fontname=Arial"
         _StyleDefs(63)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.bgcolor=&HFFFFFF&"
         _StyleDefs(64)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(66)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14,.bold=-1,.fontsize=900"
         _StyleDefs(67)  =   ":id=29,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(68)  =   ":id=29,.fontname=Arial"
         _StyleDefs(69)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15,.bgcolor=&HFFFFFF&"
         _StyleDefs(70)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(71)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
         _StyleDefs(72)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14,.bold=-1,.fontsize=900"
         _StyleDefs(73)  =   ":id=55,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(74)  =   ":id=55,.fontname=Arial"
         _StyleDefs(75)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15,.bgcolor=&HFFFFFF&"
         _StyleDefs(76)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(78)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(7).Style:id=102,.parent=13"
         _StyleDefs(82)  =   "Splits(0).Columns(7).HeadingStyle:id=99,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(7).FooterStyle:id=100,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(7).EditorStyle:id=101,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(8).Style:id=106,.parent=13"
         _StyleDefs(86)  =   "Splits(0).Columns(8).HeadingStyle:id=103,.parent=14"
         _StyleDefs(87)  =   "Splits(0).Columns(8).FooterStyle:id=104,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(8).EditorStyle:id=105,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(9).Style:id=110,.parent=13"
         _StyleDefs(90)  =   "Splits(0).Columns(9).HeadingStyle:id=107,.parent=14"
         _StyleDefs(91)  =   "Splits(0).Columns(9).FooterStyle:id=108,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(9).EditorStyle:id=109,.parent=17"
         _StyleDefs(93)  =   "Splits(0).Columns(10).Style:id=122,.parent=13"
         _StyleDefs(94)  =   "Splits(0).Columns(10).HeadingStyle:id=119,.parent=14"
         _StyleDefs(95)  =   "Splits(0).Columns(10).FooterStyle:id=120,.parent=15"
         _StyleDefs(96)  =   "Splits(0).Columns(10).EditorStyle:id=121,.parent=17"
         _StyleDefs(97)  =   "Splits(0).Columns(11).Style:id=126,.parent=13"
         _StyleDefs(98)  =   "Splits(0).Columns(11).HeadingStyle:id=123,.parent=14"
         _StyleDefs(99)  =   "Splits(0).Columns(11).FooterStyle:id=124,.parent=15"
         _StyleDefs(100) =   "Splits(0).Columns(11).EditorStyle:id=125,.parent=17"
         _StyleDefs(101) =   "Splits(0).Columns(12).Style:id=114,.parent=13"
         _StyleDefs(102) =   "Splits(0).Columns(12).HeadingStyle:id=111,.parent=14"
         _StyleDefs(103) =   "Splits(0).Columns(12).FooterStyle:id=112,.parent=15"
         _StyleDefs(104) =   "Splits(0).Columns(12).EditorStyle:id=113,.parent=17"
         _StyleDefs(105) =   "Splits(0).Columns(13).Style:id=134,.parent=13"
         _StyleDefs(106) =   "Splits(0).Columns(13).HeadingStyle:id=131,.parent=14"
         _StyleDefs(107) =   "Splits(0).Columns(13).FooterStyle:id=132,.parent=15"
         _StyleDefs(108) =   "Splits(0).Columns(13).EditorStyle:id=133,.parent=17"
         _StyleDefs(109) =   "Splits(0).Columns(14).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(110) =   "Splits(0).Columns(14).HeadingStyle:id=43,.parent=14"
         _StyleDefs(111) =   "Splits(0).Columns(14).FooterStyle:id=44,.parent=15"
         _StyleDefs(112) =   "Splits(0).Columns(14).EditorStyle:id=45,.parent=17"
         _StyleDefs(113) =   "Splits(0).Columns(15).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(114) =   "Splits(0).Columns(15).HeadingStyle:id=47,.parent=14"
         _StyleDefs(115) =   "Splits(0).Columns(15).FooterStyle:id=48,.parent=15"
         _StyleDefs(116) =   "Splits(0).Columns(15).EditorStyle:id=49,.parent=17"
         _StyleDefs(117) =   "Splits(0).Columns(16).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(118) =   "Splits(0).Columns(16).HeadingStyle:id=59,.parent=14"
         _StyleDefs(119) =   "Splits(0).Columns(16).FooterStyle:id=60,.parent=15"
         _StyleDefs(120) =   "Splits(0).Columns(16).EditorStyle:id=61,.parent=17"
         _StyleDefs(121) =   "Splits(0).Columns(17).Style:id=70,.parent=13,.alignment=1"
         _StyleDefs(122) =   "Splits(0).Columns(17).HeadingStyle:id=67,.parent=14"
         _StyleDefs(123) =   "Splits(0).Columns(17).FooterStyle:id=68,.parent=15"
         _StyleDefs(124) =   "Splits(0).Columns(17).EditorStyle:id=69,.parent=17"
         _StyleDefs(125) =   "Splits(0).Columns(18).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(126) =   "Splits(0).Columns(18).HeadingStyle:id=71,.parent=14"
         _StyleDefs(127) =   "Splits(0).Columns(18).FooterStyle:id=72,.parent=15"
         _StyleDefs(128) =   "Splits(0).Columns(18).EditorStyle:id=73,.parent=17"
         _StyleDefs(129) =   "Splits(0).Columns(19).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(130) =   "Splits(0).Columns(19).HeadingStyle:id=75,.parent=14"
         _StyleDefs(131) =   "Splits(0).Columns(19).FooterStyle:id=76,.parent=15"
         _StyleDefs(132) =   "Splits(0).Columns(19).EditorStyle:id=77,.parent=17"
         _StyleDefs(133) =   "Splits(0).Columns(20).Style:id=82,.parent=13,.alignment=1"
         _StyleDefs(134) =   "Splits(0).Columns(20).HeadingStyle:id=79,.parent=14"
         _StyleDefs(135) =   "Splits(0).Columns(20).FooterStyle:id=80,.parent=15"
         _StyleDefs(136) =   "Splits(0).Columns(20).EditorStyle:id=81,.parent=17"
         _StyleDefs(137) =   "Splits(0).Columns(21).Style:id=86,.parent=13,.alignment=1"
         _StyleDefs(138) =   "Splits(0).Columns(21).HeadingStyle:id=83,.parent=14"
         _StyleDefs(139) =   "Splits(0).Columns(21).FooterStyle:id=84,.parent=15"
         _StyleDefs(140) =   "Splits(0).Columns(21).EditorStyle:id=85,.parent=17"
         _StyleDefs(141) =   "Splits(0).Columns(22).Style:id=90,.parent=13,.alignment=2"
         _StyleDefs(142) =   "Splits(0).Columns(22).HeadingStyle:id=87,.parent=14"
         _StyleDefs(143) =   "Splits(0).Columns(22).FooterStyle:id=88,.parent=15"
         _StyleDefs(144) =   "Splits(0).Columns(22).EditorStyle:id=89,.parent=17"
         _StyleDefs(145) =   "Splits(0).Columns(23).Style:id=94,.parent=13,.alignment=1"
         _StyleDefs(146) =   "Splits(0).Columns(23).HeadingStyle:id=91,.parent=14"
         _StyleDefs(147) =   "Splits(0).Columns(23).FooterStyle:id=92,.parent=15"
         _StyleDefs(148) =   "Splits(0).Columns(23).EditorStyle:id=93,.parent=17"
         _StyleDefs(149) =   "Splits(0).Columns(24).Style:id=98,.parent=13,.alignment=0"
         _StyleDefs(150) =   "Splits(0).Columns(24).HeadingStyle:id=95,.parent=14"
         _StyleDefs(151) =   "Splits(0).Columns(24).FooterStyle:id=96,.parent=15"
         _StyleDefs(152) =   "Splits(0).Columns(24).EditorStyle:id=97,.parent=17"
         _StyleDefs(153) =   "Splits(0).Columns(25).Style:id=138,.parent=13"
         _StyleDefs(154) =   "Splits(0).Columns(25).HeadingStyle:id=135,.parent=14"
         _StyleDefs(155) =   "Splits(0).Columns(25).FooterStyle:id=136,.parent=15"
         _StyleDefs(156) =   "Splits(0).Columns(25).EditorStyle:id=137,.parent=17"
         _StyleDefs(157) =   "Named:id=33:Normal"
         _StyleDefs(158) =   ":id=33,.parent=0,.bgcolor=&HFF80&,.fgcolor=&HFFFFFF&,.borderColor=&H800040&"
         _StyleDefs(159) =   "Named:id=34:Heading"
         _StyleDefs(160) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(161) =   ":id=34,.wraptext=-1"
         _StyleDefs(162) =   "Named:id=35:Footing"
         _StyleDefs(163) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(164) =   "Named:id=36:Selected"
         _StyleDefs(165) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(166) =   ":id=36,.borderColor=&H80000013&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(167) =   ":id=36,.strikethrough=0,.charset=0"
         _StyleDefs(168) =   ":id=36,.fontname=MS Sans Serif"
         _StyleDefs(169) =   "Named:id=37:Caption"
         _StyleDefs(170) =   ":id=37,.parent=34,.alignment=2,.bgcolor=&H80000009&"
         _StyleDefs(171) =   "Named:id=38:HighlightRow"
         _StyleDefs(172) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&HFFFFFF&,.borderColor=&H800040&"
         _StyleDefs(173) =   "Named:id=39:EvenRow"
         _StyleDefs(174) =   ":id=39,.parent=33"
         _StyleDefs(175) =   "Named:id=40:OddRow"
         _StyleDefs(176) =   ":id=40,.parent=33,.bgcolor=&H4000&"
         _StyleDefs(177) =   "Named:id=41:RecordSelector"
         _StyleDefs(178) =   ":id=41,.parent=34"
         _StyleDefs(179) =   "Named:id=42:FilterBar"
         _StyleDefs(180) =   ":id=42,.parent=33,.bgcolor=&HFF0000&"
      End
      Begin VB.TextBox txtKodelgn 
         Appearance      =   0  'Flat
         Height          =   333
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2300
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   11000
         TabIndex        =   21
         Top             =   8600
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtPpn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00BCE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   630
            Width           =   2655
         End
         Begin VB.TextBox txtDpp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   990
            TabIndex        =   23
            Top             =   315
            Width           =   2655
         End
         Begin VB.TextBox txtDppPpn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   990
            TabIndex        =   22
            Top             =   945
            Width           =   2655
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DPP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   28
            Top             =   315
            Width           =   1005
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H00BCE4D7&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PPN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   27
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   26
            Top             =   945
            Width           =   1005
         End
         Begin VB.Label Label15 
            Appearance      =   0  'Flat
            BackColor       =   &H003F3F3F&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Perincian Claim Disetujui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   3645
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   4200
         TabIndex        =   8
         Top             =   8600
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox txtTotalAcc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00BCE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   630
            Width           =   2655
         End
         Begin VB.TextBox txtQtyAju 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   2160
            TabIndex        =   15
            Top             =   315
            Width           =   975
         End
         Begin VB.TextBox txtQtyAcc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00BCE4D7&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   2160
            TabIndex        =   14
            Top             =   630
            Width           =   975
         End
         Begin VB.TextBox txtQtyTolak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   2160
            TabIndex        =   13
            Top             =   945
            Width           =   975
         End
         Begin VB.TextBox txtTotalAju 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   3120
            TabIndex        =   12
            Top             =   315
            Width           =   2655
         End
         Begin VB.TextBox txtTotalTolak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   333
            Left            =   3120
            TabIndex        =   11
            Top             =   945
            Width           =   2655
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H003F3F3F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   333
            Left            =   3120
            TabIndex        =   10
            Text            =   "Nominal "
            Top             =   0
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H003F3F3F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   333
            Left            =   2160
            TabIndex        =   9
            Text            =   "Jumlah "
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Claim Diajukan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   20
            Top             =   315
            Width           =   2205
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00BCE4D7&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Claim Disetujui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   19
            Top             =   630
            Width           =   2205
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Claim Ditolak"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   0
            TabIndex        =   18
            Top             =   945
            Width           =   2205
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H003F3F3F&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Summary Claim"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   2205
         End
      End
      Begin VB.TextBox TxtNamaCabang 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   333
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2300
         Width           =   3615
      End
      Begin VB.TextBox cboKodeLangganan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   333
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   10440
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "Ms. Excel|*.xls"
      End
      Begin AIFCmp1.asxPowerBanner LabeLKet 
         Height          =   105
         Left            =   120
         Top             =   2820
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   185
         StartColor      =   701524
         EndColor        =   701524
         FormatString    =   ""
         Orientation     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBDate6Ctl.TDBDate DtTglAwal 
         Height          =   330
         Left            =   4320
         TabIndex        =   30
         Top             =   1575
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   587
         Calendar        =   "FrmTwc_ViewCabang.frx":1F30
         Caption         =   "FrmTwc_ViewCabang.frx":205C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmTwc_ViewCabang.frx":20B8
         Keys            =   "FrmTwc_ViewCabang.frx":20D6
         Spin            =   "FrmTwc_ViewCabang.frx":2134
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
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
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "11/04/2007"
         ValidateMode    =   0
         ValueVT         =   6815751
         Value           =   39183
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate DtTglAkhir 
         Height          =   330
         Left            =   6480
         TabIndex        =   31
         Top             =   1575
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   587
         Calendar        =   "FrmTwc_ViewCabang.frx":215C
         Caption         =   "FrmTwc_ViewCabang.frx":2288
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmTwc_ViewCabang.frx":22E4
         Keys            =   "FrmTwc_ViewCabang.frx":2302
         Spin            =   "FrmTwc_ViewCabang.frx":2360
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
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
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   "_"
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "11/04/2007"
         ValidateMode    =   0
         ValueVT         =   6815751
         Value           =   39183
         CenturyMode     =   0
      End
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image Image1 
         Height          =   400
         Left            =   120
         Picture         =   "FrmTwc_ViewCabang.frx":2388
         Stretch         =   -1  'True
         Top             =   840
         Width           =   14655
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Finance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4320
         TabIndex        =   42
         Top             =   2055
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   120
         Picture         =   "FrmTwc_ViewCabang.frx":664D
         Stretch         =   -1  'True
         Top             =   10020
         Width           =   14655
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "View Proses Warranty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   11400
         TabIndex        =   41
         Top             =   925
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perincian Warranty Claim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   40
         Top             =   3075
         Width           =   2625
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Cabang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2055
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Cabang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s / d"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5880
         TabIndex        =   37
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4320
         TabIndex        =   36
         Top             =   1365
         Width           =   840
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H003F3F3F&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " No. PTC HO Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         TabIndex        =   35
         Top             =   8595
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.Label txtPtc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   240
         TabIndex        =   34
         Top             =   8910
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H003F3F3F&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Status Proses HO Finance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   240
         TabIndex        =   33
         Top             =   9225
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.Label txtFinance 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   240
         TabIndex        =   32
         Top             =   9540
         Visible         =   0   'False
         Width           =   2925
      End
   End
End
Attribute VB_Name = "FrmTwc_ViewCabang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsTemp              As New ADODB.Recordset
Dim NegativeStyle       As New TrueOleDBGrid70.Style
Dim cols                As TrueOleDBGrid70.Columns
Dim bCekData            As Boolean

Private Sub SetAdodc()
    Dim rsDe As ADODB.Recordset
    Set rsDe = objO.FetchData("Select FS_KD_Dealer, FS_NM_Dealer FROm T_Profil WITH (NOLOCK)")
    cboKodeLangganan.Text = rsDe.Fields(0).value
    TxtNamaCabang.Text = rsDe.Fields(1).value
End Sub
'
'Private Sub CboKodeLangganan_ClickData()
'CreateTbl
'If cboKodeLangganan <> "" Then
'   Set rsFind = New ADODB.Recordset
'   rsFind.Open "Select Kodedealer, Namadealer from Dealer where Kd_Twc_Dealer = '" & cboKodeLangganan & "' ", DbConP, adOpenKeyset
'   If rsFind.RecordCount <> 0 Then
'    TxtNamaCabang = rsFind!namadealer
'    txtKodelgn = rsFind!KodeDealer
'   Else
'    TxtNamaCabang = ""
'    txtKodelgn = ""
'   End If
'End If
'
'End Sub

Private Sub cmdExport_Click()
On Error GoTo Err_Export
Dim nmFile As String
CDialog.filter = "*.csv"
CDialog.FileName = "Rincian PTC." & cboKodeLangganan
CDialog.CancelError = True
CDialog.ShowSave
    nmFile = CDialog.FileName & ".csv"
    GridRekap.ExportToDelimitedFile nmFile, , ";", , , True
    If MsgBox("File " & nmFile & " telah di Export." + vbCrLf + "Apakah mau dibuka sekarang?", vbQuestion + vbYesNo, "Export File") = vbYes Then
       Shell "Explorer " & nmFile, vbMaximizedFocus
    End If
Err_Export:


End Sub

Private Sub cmdView_Click()
1     bCekData = False
2     createTbl
3     Dim rsFind As ADODB.Recordset
4     sSQL = "SELECT L.NOBUKTI, L.NOBARIS, "
5     sSQL = sSQL + "L.NOWO, L.CLAIMNO, L.NORANGKA, L.OPENO, L.OFP, "
6     sSQL = sSQL + "L.HOUR, L.JASA, L.SUBLET, L.PART, L.SUBTOTAL, L.PAJAK, L.LAIN, L.TOTAL, "
7     sSQL = sSQL + "T.Fs_kd_plg Kodelgn, T.Customer Namalgn, T.No_Polisi, T.Fmdl, T.Color, L.NOINVOICE, "
8     sSQL = sSQL + "CASE WHEN L.ACCFLAG=1 THEN 'YA' ELSE 'TIDAK' END AS ACC , L.TOTACC, L.KETERANGAN, "
9     sSQL = sSQL + "CASE WHEN L.ACCFLAG=0 THEN 'TIDAK DIPROSES' ELSE (CASE WHEN (H.FinanceFLAG=1 AND Isnull(H.Userbayar,'')<>'') THEN 'DIBAYAR' WHEN (H.FinanceFLAG=1 AND Isnull(H.Userbayar,'')='') THEN 'PENAGIHAN' ELSE 'PENGAJUAN' END) END AS Finance, "
10    sSQL = sSQL + "H.QtyAcc, H.QtyTolak, H.JlhAcc, H.JlhTolak, H.JlhDpp, H.JlhPpn, T.Invdate "
11    sSQL = sSQL + "FROM PTCLIST L WITH (NOLOCK) "
12    sSQL = sSQL + "LEFT JOIN PTCHDR H WITH (NOLOCK) ON L.NOBUKTI=H.NOBUKTI "
13    sSQL = sSQL + "LEFT JOIN T_INVOICENEW T WITH (NOLOCK) ON H.DBCABANG=T.CABANG AND L.NOWO=T.No_wo AND L.NORANGKA=T.No_Chassis AND L.NOINVOICE=T.NOINV "
14    sSQL = sSQL + "WHERE (isnull(T.tpsvc,'') = 'w') and (isnull(T.Fs_No_Retur,'') = '') and (Kodesvc = '" & cboKodeLangganan & "') and (Tglproses Between '" & Format(DtTglAwal, "mm/dd/yyyy") & "' and '" & Format(DtTglAkhir, "mm/dd/yyyy") & " 23:59:59') "
15    sSQL = sSQL + "Order By L.NOBUKTI, L.NOBARIS "
16      Set rsFind = objO.FetchDataHO(sSQL)
17        With rsFind
18        If .RecordCount > 0 Then
19           .MoveFirst
20           Do Until .EOF
21            rsTemp.AddNew
22            rsTemp!Nobukti = Nul(.Fields(0).value)
23            rsTemp!noBaris = Nul(.Fields(1).value)
24            rsTemp!noWO = Nul(.Fields(2).value)
25            rsTemp!Claimno = Nul(.Fields(3).value)
26            rsTemp!noRangka = Nul(.Fields(4).value)
27            rsTemp!Openo = Nul(.Fields(5).value)
28            rsTemp!OFP = Nul(.Fields(6).value)
29            rsTemp!Hour = Nul(.Fields(7).value, 0)
30            rsTemp!Jasa = Nul(.Fields(8).value, 0)
31            rsTemp!SUBLET = Nul(.Fields(9).value, 0)
32            rsTemp!Part = Nul(.Fields(10).value, 0)
33            rsTemp!SUBTOTAL = Nul(.Fields(11).value, 0)
34            rsTemp!Pajak = Nul(.Fields(12).value, 0)
35            rsTemp!lain = Nul(.Fields(13).value, 0)
36            rsTemp!total = Nul(.Fields(14).value, 0)
37            rsTemp!Kodelgn = Nul(.Fields(15).value)
38            rsTemp!Namalgn = Nul(.Fields(16).value, 0)
39            rsTemp!NO_POLISI = Nul(.Fields(17).value, 0)
40            rsTemp!Fmdl = Nul(.Fields(18).value, 0)
41            rsTemp!Color = Nul(.Fields(19).value, 0)
42            rsTemp!NoInvoice = Nul(.Fields(20).value, 0)
43            rsTemp!InvDate = Nul(.Fields(31).value)
44            rsTemp!Acc = Nul(.Fields(21).value)
45            rsTemp!TotAcc = Nul(.Fields(22).value, 0)
46            rsTemp!keterangan = Nul(.Fields(23).value)
47            rsTemp!Finance = Nul(.Fields(24).value)
48            rsTemp!Qtyacc = Nul(.Fields(25).value, 0)
49            rsTemp!Qtytolak = Nul(.Fields(26).value, 0)
50            rsTemp!jlhacc = Nul(.Fields(27).value, 0)
51            rsTemp!jlhtolak = Nul(.Fields(28).value, 0)
52            rsTemp!jlhdpp = Nul(.Fields(29).value, 0)
53            rsTemp!jlhppn = Nul(.Fields(30).value, 0)

'65            nTotalAju = nTotalAju + rsTemp!Total
'66            GridRekap.Columns(19).FooterText = FormatNumber(nTotalAju, 2, , vbTrue)
'67            txtTotalAju = FormatNumber(nTotalAju, 2, , vbTrue)

54            rsTemp.Update
55            .MoveNext
56           Loop
57           bCekData = True
58        End If
59        End With
End Sub

Private Sub Form_Load()
Move 0, 0
For i = 0 To 19
    GridRekap.Splits(0).Columns(i).Locked = True
Next i

DtTglAwal = Date
DtTglAkhir = Date

SetAdodc
createTbl

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub createTbl(Optional Edit As Boolean = False)
1     On Error GoTo vb_error
2      If rsTemp.State Then rsTemp.Close
3      Set rsTemp = Nothing
4      With rsTemp
5       .Fields.Append "Nobukti", adVarChar, 25
6       .Fields.Append "Nobaris", adSingle
7       .Fields.Append "Nowo", adVarChar, 20
8       .Fields.Append "Claimno", adVarChar, 20
9       .Fields.Append "Norangka", adVarChar, 30
10      .Fields.Append "Openo", adVarChar, 20
11      .Fields.Append "OFP", adVarChar, 20
12      .Fields.Append "Hour", adCurrency
13      .Fields.Append "Jasa", adCurrency
14      .Fields.Append "Sublet", adCurrency
15      .Fields.Append "Part", adCurrency
16      .Fields.Append "Subtotal", adCurrency
17      .Fields.Append "Pajak", adCurrency
18      .Fields.Append "Lain", adCurrency
19      .Fields.Append "Total", adCurrency
20      .Fields.Append "Kodelgn", adVarChar, 20
21      .Fields.Append "Namalgn", adVarChar, 20
22      .Fields.Append "No_Polisi", adVarChar, 20
23      .Fields.Append "Fmdl", adVarChar, 100
24      .Fields.Append "Color", adVarChar, 50
25      .Fields.Append "Noinvoice", adVarChar, 20
26      .Fields.Append "Invdate", adDate
27      .Fields.Append "Acc", adVarChar, 10
28      .Fields.Append "TotAcc", adCurrency
29      .Fields.Append "Keterangan", adVarChar, 1000
30      .Fields.Append "Finance", adVarChar, 25
31      .Fields.Append "QtyAcc", adSingle
32      .Fields.Append "QtyTolak", adSingle
33      .Fields.Append "JlhAcc", adCurrency
34      .Fields.Append "JlhTolak", adCurrency
35      .Fields.Append "JlhDpp", adCurrency
36      .Fields.Append "JlhPpn", adCurrency
37      .Open
      '28       AJumlah = 0
      '29       BJumlah = 0
      '30       AJlAcc = 0
      '31       AJlTlk = 0
38      Set GridRekap.DataSource = rsTemp
39      End With
40      sNamacabang = TxtNamaCabang
41      sKodelgn = txtKodelgn
        'Call KosongText(Me)
42      TxtNamaCabang = sNamacabang
43      txtKodelgn = sKodelgn
44      Text1 = "Jumlah "
45      Text2 = "Nominal "
46      txtPtc = ""
47      txtFinance = ""
48    Exit Sub
vb_error:
49        MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CreateTbl"), vbExclamation, "Err Number : " & Erl
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
   'If Me.Width < (MenuUtama.Width - MenuUtama.panelkiri.Width) Then Me.Width = (MenuUtama.Width - MenuUtama.panelkiri.Width) - 200
   'Frame1.Left = (Me.Width - Frame1.Width) / 2
   'FrameTombol.Left = (Me.Width - FrameTombol.Width) / 2
End If
End Sub

Private Sub GridRekap_FilterChange()
      Dim C As Integer
2     Set cols = GridRekap.Columns
3     C = GridRekap.col
4     GridRekap.HoldFields
6     rsTemp.filter = getFilter()
7       If rsTemp.EOF Then
8          MsgBox "Data yang dimaksud tidak dapat ditemukan !", vbInformation
9          For Each col In GridRekap.Columns
10            col.FilterText = ""
11         Next col
12         rsTemp.filter = ""
13         Set GridRekap.DataSource = rsTemp
14         GridRekap.ReOpen
'15         Call HitungTotaL("='" & KodeGudang.Text & "'")
16         GridRekap.col = C
'17         LblRecord = "Jumlah Line Item: " & rsTemp.RecordCount
18         Exit Sub
19      End If
'20    Call HitungTotaL("='" & KodeGudang.Text & "'")
21    GridRekap.col = C
'22    LblRecord = "Jumlah Line Item: " & rsTemp.RecordCount
23    GridRekap.EditActive = True
24    Exit Sub

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
            tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
        End If
    Next col
    getFilter = tmp

End Function

Private Sub GridRekap_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If bCekData = False Then Exit Sub
With rsTemp
 If Not rsTemp.EOF Then
    txtQtyAju = !Qtyacc + !Qtytolak
    txtTotalAju = FormatNumber(!jlhacc + !jlhtolak, 2, , vbTrue)
    txtQtyAcc = !Qtyacc
    txtTotalAcc = FormatNumber(!jlhacc, 2, , vbTrue)
    txtQtyTolak = !Qtytolak
    txtTotalTolak = FormatNumber(!jlhtolak, 2, , vbTrue)
    txtDpp = FormatNumber(!jlhdpp, 2, , vbTrue)
    txtPpn = FormatNumber(!jlhppn, 2, , vbTrue)
    txtDppPpn = FormatNumber(!jlhdpp + !jlhppn, 2, , vbTrue)
    txtPtc = " " & !Nobukti
    txtFinance = " " & !Finance
 End If
End With
End Sub


