VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MsComCtl.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FormReminder1000 
   BorderStyle     =   0  'None
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "FormReminder1000.frx":0000
   ScaleHeight     =   8700
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScrollFUB 
      Height          =   8655
      LargeChange     =   300
      Left            =   11730
      Max             =   6500
      SmallChange     =   350
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   480
      ScaleHeight     =   8655
      ScaleWidth      =   12255
      TabIndex        =   1
      Top             =   0
      Width           =   12255
      Begin TabDlg.SSTab SSFUB 
         Height          =   16005
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   28231
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "  "
         TabPicture(0)   =   "FormReminder1000.frx":EF16
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Shape4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Shape2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label52"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label55"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label50"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Shape6"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblJml"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Shape1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Line22"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Line21"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Line16"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Line20"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Line19"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lblKet"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Shape3"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Line7"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Line6"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "ImgView"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Line4"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Line3"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Line2"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Label2"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Line1"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Label1"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "lblPart"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "ShapeXLS"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Line10"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "lbLen"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Label51"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Label54(0)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Shape8"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "Line13"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "Label53"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "Label56"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "Line14"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "Shape9"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "Label57"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "Label59"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "Line23"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "Label60"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "Label69"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "lblSave"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "Label70"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "Line25"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "FrameGRP"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "sService"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "MAPIMessages1"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "MAPISession1"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "Frame2"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "DTSend"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "DtRemain0"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "GridXlsMrs"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "ComDialogXls"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "DtRemain"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "ImageList"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "cmdExit"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).Control(57)=   "cmdTop"
         Tab(0).Control(57).Enabled=   0   'False
         Tab(0).Control(58)=   "Frame4"
         Tab(0).Control(58).Enabled=   0   'False
         Tab(0).Control(59)=   "Frame10"
         Tab(0).Control(59).Enabled=   0   'False
         Tab(0).Control(60)=   "Frame9"
         Tab(0).Control(60).Enabled=   0   'False
         Tab(0).Control(61)=   "Frame1"
         Tab(0).Control(61).Enabled=   0   'False
         Tab(0).Control(62)=   "cmdProcess"
         Tab(0).Control(62).Enabled=   0   'False
         Tab(0).Control(63)=   "cboType"
         Tab(0).Control(63).Enabled=   0   'False
         Tab(0).Control(64)=   "cmdXlsMRS"
         Tab(0).Control(64).Enabled=   0   'False
         Tab(0).Control(65)=   "CmdXls"
         Tab(0).Control(65).Enabled=   0   'False
         Tab(0).Control(66)=   "onS"
         Tab(0).Control(66).Enabled=   0   'False
         Tab(0).Control(67)=   "onC"
         Tab(0).Control(67).Enabled=   0   'False
         Tab(0).Control(68)=   "cmdTemplate"
         Tab(0).Control(68).Enabled=   0   'False
         Tab(0).Control(69)=   "txtPesan"
         Tab(0).Control(69).Enabled=   0   'False
         Tab(0).Control(70)=   "fot"
         Tab(0).Control(70).Enabled=   0   'False
         Tab(0).Control(71)=   "cmdSMS"
         Tab(0).Control(71).Enabled=   0   'False
         Tab(0).Control(72)=   "chCDMA"
         Tab(0).Control(72).Enabled=   0   'False
         Tab(0).Control(73)=   "chGSM"
         Tab(0).Control(73).Enabled=   0   'False
         Tab(0).Control(74)=   "cmdSave"
         Tab(0).Control(74).Enabled=   0   'False
         Tab(0).Control(75)=   "GridRemainder"
         Tab(0).Control(75).Enabled=   0   'False
         Tab(0).Control(76)=   "on1K"
         Tab(0).Control(76).Enabled=   0   'False
         Tab(0).Control(77)=   "cmd1K"
         Tab(0).Control(77).Enabled=   0   'False
         Tab(0).Control(78)=   "onD"
         Tab(0).Control(78).Enabled=   0   'False
         Tab(0).ControlCount=   79
         Begin VB.CheckBox onD 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7200
            TabIndex        =   210
            Top             =   5000
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.CommandButton cmd1K 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   142
            Top             =   4990
            Width           =   495
         End
         Begin VB.CheckBox on1K 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Service 1.000 KM"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   141
            Top             =   5000
            Width           =   1815
         End
         Begin TrueOleDBGrid70.TDBGrid GridRemainder 
            Height          =   2895
            Left            =   120
            TabIndex        =   86
            Top             =   2040
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   5106
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "No SPK"
            Columns(0).DataField=   "NoSPK"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Tgl DO"
            Columns(1).DataField=   "Tgl_DO"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Status Data"
            Columns(2).DataField=   "statusdata"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tgl CR1"
            Columns(3).DataField=   "Tgl_CR1"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Km CR1"
            Columns(4).DataField=   "Km_CR1"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tgl CR2"
            Columns(5).DataField=   "Tgl_CR2"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Km CR2"
            Columns(6).DataField=   "Km_CR2"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nama Customer"
            Columns(7).DataField=   "Nama_Cust"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Alamat Customer"
            Columns(8).DataField=   "Alamat_Cust"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Kota"
            Columns(9).DataField=   "kota_Cust"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "KodePos"
            Columns(10).DataField=   "kodePos_cust"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "No Polisi"
            Columns(11).DataField=   "No_Polisi"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "No Engine"
            Columns(12).DataField=   "No_Engine"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "No Chassis"
            Columns(13).DataField=   "NoChassis"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Telpon"
            Columns(14).DataField=   "Telpon"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "GSM"
            Columns(15).DataField=   "GSM"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "CDMA"
            Columns(16).DataField=   "CDMA"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "Kode Type"
            Columns(17).DataField=   "Kode_Type"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).Caption=   "Nama Type"
            Columns(18).DataField=   "Nama_Type"
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(19)._VlistStyle=   0
            Columns(19)._MaxComboItems=   5
            Columns(19).Caption=   "Kode Warna"
            Columns(19).DataField=   "Kode_Warna"
            Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(20)._VlistStyle=   0
            Columns(20)._MaxComboItems=   5
            Columns(20).Caption=   "Nama Warna"
            Columns(20).DataField=   "Nama_Warna"
            Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(21)._VlistStyle=   0
            Columns(21)._MaxComboItems=   5
            Columns(21).Caption=   "Wiraniaga"
            Columns(21).DataField=   "Wiraniaga"
            Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(22)._VlistStyle=   0
            Columns(22)._MaxComboItems=   5
            Columns(22).Caption=   "Nama STNK"
            Columns(22).DataField=   "Nama_STNK"
            Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(23)._VlistStyle=   0
            Columns(23)._MaxComboItems=   5
            Columns(23).Caption=   "Alamat STNK"
            Columns(23).DataField=   "Alamat_STNK"
            Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(24)._VlistStyle=   0
            Columns(24)._MaxComboItems=   5
            Columns(24).Caption=   "Kota STNK"
            Columns(24).DataField=   "Kota_STNK"
            Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(25)._VlistStyle=   0
            Columns(25)._MaxComboItems=   5
            Columns(25).Caption=   "KodePos STNK"
            Columns(25).DataField=   "KodePos_STNK"
            Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(26)._VlistStyle=   0
            Columns(26)._MaxComboItems=   5
            Columns(26).Caption=   "Tempat Lahir"
            Columns(26).DataField=   "Tempat_Lahir"
            Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(27)._VlistStyle=   0
            Columns(27)._MaxComboItems=   5
            Columns(27).Caption=   "Tanggal Lahir"
            Columns(27).DataField=   "Tanggal_Lahir"
            Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(28)._VlistStyle=   0
            Columns(28)._MaxComboItems=   5
            Columns(28).Caption=   "Agama"
            Columns(28).DataField=   "Agama"
            Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(29)._VlistStyle=   0
            Columns(29)._MaxComboItems=   5
            Columns(29).Caption=   "Pekerjaan"
            Columns(29).DataField=   "Pekerjaan"
            Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(30)._VlistStyle=   0
            Columns(30)._MaxComboItems=   5
            Columns(30).Caption=   "Hoby"
            Columns(30).DataField=   "Hoby"
            Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(31)._VlistStyle=   0
            Columns(31)._MaxComboItems=   5
            Columns(31).Caption=   "Tgl Svc"
            Columns(31).DataField=   "Tgl_Service"
            Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(32)._VlistStyle=   4
            Columns(32)._MaxComboItems=   5
            Columns(32).Caption=   "Select"
            Columns(32).DataField=   ""
            Columns(32).NumberFormat=   "Yes/No"
            Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(33)._VlistStyle=   0
            Columns(33)._MaxComboItems=   5
            Columns(33).Caption=   "No_Chassis"
            Columns(33).DataField=   "Nomor_Rangka"
            Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(34)._VlistStyle=   0
            Columns(34)._MaxComboItems=   5
            Columns(34).Caption=   "Kode_Lgn"
            Columns(34).DataField=   "Kode_Lgn"
            Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   35
            Splits(0)._UserFlags=   0
            Splits(0).AllowSizing=   -1  'True
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=35"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2646"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2540"
            Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(10)=   "Column(2).Width=3281"
            Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=3175"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=2646"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2540"
            Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1667"
            Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(24)=   "Column(5).Width=2646"
            Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2540"
            Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(29)=   "Column(6).Width=1773"
            Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=1667"
            Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(34)=   "Column(7).Width=5477"
            Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=5371"
            Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(38)=   "Column(8).Width=7064"
            Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=6959"
            Splits(0)._ColumnProps(41)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(42)=   "Column(9).Width=3281"
            Splits(0)._ColumnProps(43)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(9)._WidthInPix=3175"
            Splits(0)._ColumnProps(45)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(46)=   "Column(10).Width=2196"
            Splits(0)._ColumnProps(47)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(10)._WidthInPix=2090"
            Splits(0)._ColumnProps(49)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(50)=   "Column(11).Width=2646"
            Splits(0)._ColumnProps(51)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(11)._WidthInPix=2540"
            Splits(0)._ColumnProps(53)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(54)=   "Column(12).Width=3281"
            Splits(0)._ColumnProps(55)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(56)=   "Column(12)._WidthInPix=3175"
            Splits(0)._ColumnProps(57)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(58)=   "Column(13).Width=3889"
            Splits(0)._ColumnProps(59)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(60)=   "Column(13)._WidthInPix=3784"
            Splits(0)._ColumnProps(61)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(62)=   "Column(14).Width=3281"
            Splits(0)._ColumnProps(63)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(14)._WidthInPix=3175"
            Splits(0)._ColumnProps(65)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(66)=   "Column(15).Width=3281"
            Splits(0)._ColumnProps(67)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(68)=   "Column(15)._WidthInPix=3175"
            Splits(0)._ColumnProps(69)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(70)=   "Column(16).Width=3281"
            Splits(0)._ColumnProps(71)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(72)=   "Column(16)._WidthInPix=3175"
            Splits(0)._ColumnProps(73)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(74)=   "Column(17).Width=4419"
            Splits(0)._ColumnProps(75)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(76)=   "Column(17)._WidthInPix=4313"
            Splits(0)._ColumnProps(77)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(78)=   "Column(18).Width=3969"
            Splits(0)._ColumnProps(79)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(80)=   "Column(18)._WidthInPix=3863"
            Splits(0)._ColumnProps(81)=   "Column(18).Order=19"
            Splits(0)._ColumnProps(82)=   "Column(19).Width=1773"
            Splits(0)._ColumnProps(83)=   "Column(19).DividerColor=0"
            Splits(0)._ColumnProps(84)=   "Column(19)._WidthInPix=1667"
            Splits(0)._ColumnProps(85)=   "Column(19).Order=20"
            Splits(0)._ColumnProps(86)=   "Column(20).Width=3519"
            Splits(0)._ColumnProps(87)=   "Column(20).DividerColor=0"
            Splits(0)._ColumnProps(88)=   "Column(20)._WidthInPix=3413"
            Splits(0)._ColumnProps(89)=   "Column(20).Order=21"
            Splits(0)._ColumnProps(90)=   "Column(21).Width=5292"
            Splits(0)._ColumnProps(91)=   "Column(21).DividerColor=0"
            Splits(0)._ColumnProps(92)=   "Column(21)._WidthInPix=5186"
            Splits(0)._ColumnProps(93)=   "Column(21).Order=22"
            Splits(0)._ColumnProps(94)=   "Column(22).Width=5292"
            Splits(0)._ColumnProps(95)=   "Column(22).DividerColor=0"
            Splits(0)._ColumnProps(96)=   "Column(22)._WidthInPix=5186"
            Splits(0)._ColumnProps(97)=   "Column(22).Order=23"
            Splits(0)._ColumnProps(98)=   "Column(23).Width=9181"
            Splits(0)._ColumnProps(99)=   "Column(23).DividerColor=0"
            Splits(0)._ColumnProps(100)=   "Column(23)._WidthInPix=9075"
            Splits(0)._ColumnProps(101)=   "Column(23).Order=24"
            Splits(0)._ColumnProps(102)=   "Column(24).Width=3281"
            Splits(0)._ColumnProps(103)=   "Column(24).DividerColor=0"
            Splits(0)._ColumnProps(104)=   "Column(24)._WidthInPix=3175"
            Splits(0)._ColumnProps(105)=   "Column(24).Order=25"
            Splits(0)._ColumnProps(106)=   "Column(25).Width=3281"
            Splits(0)._ColumnProps(107)=   "Column(25).DividerColor=0"
            Splits(0)._ColumnProps(108)=   "Column(25)._WidthInPix=3175"
            Splits(0)._ColumnProps(109)=   "Column(25).Order=26"
            Splits(0)._ColumnProps(110)=   "Column(26).Width=3281"
            Splits(0)._ColumnProps(111)=   "Column(26).DividerColor=0"
            Splits(0)._ColumnProps(112)=   "Column(26)._WidthInPix=3175"
            Splits(0)._ColumnProps(113)=   "Column(26).Order=27"
            Splits(0)._ColumnProps(114)=   "Column(27).Width=3281"
            Splits(0)._ColumnProps(115)=   "Column(27).DividerColor=0"
            Splits(0)._ColumnProps(116)=   "Column(27)._WidthInPix=3175"
            Splits(0)._ColumnProps(117)=   "Column(27).Order=28"
            Splits(0)._ColumnProps(118)=   "Column(28).Width=3281"
            Splits(0)._ColumnProps(119)=   "Column(28).DividerColor=0"
            Splits(0)._ColumnProps(120)=   "Column(28)._WidthInPix=3175"
            Splits(0)._ColumnProps(121)=   "Column(28).Order=29"
            Splits(0)._ColumnProps(122)=   "Column(29).Width=3519"
            Splits(0)._ColumnProps(123)=   "Column(29).DividerColor=0"
            Splits(0)._ColumnProps(124)=   "Column(29)._WidthInPix=3413"
            Splits(0)._ColumnProps(125)=   "Column(29).Order=30"
            Splits(0)._ColumnProps(126)=   "Column(30).Width=3519"
            Splits(0)._ColumnProps(127)=   "Column(30).DividerColor=0"
            Splits(0)._ColumnProps(128)=   "Column(30)._WidthInPix=3413"
            Splits(0)._ColumnProps(129)=   "Column(30).Order=31"
            Splits(0)._ColumnProps(130)=   "Column(31).Width=3281"
            Splits(0)._ColumnProps(131)=   "Column(31).DividerColor=0"
            Splits(0)._ColumnProps(132)=   "Column(31)._WidthInPix=3175"
            Splits(0)._ColumnProps(133)=   "Column(31).Order=32"
            Splits(0)._ColumnProps(134)=   "Column(32).Width=450"
            Splits(0)._ColumnProps(135)=   "Column(32).DividerColor=0"
            Splits(0)._ColumnProps(136)=   "Column(32)._WidthInPix=344"
            Splits(0)._ColumnProps(137)=   "Column(32).Visible=0"
            Splits(0)._ColumnProps(138)=   "Column(32).Order=33"
            Splits(0)._ColumnProps(139)=   "Column(33).Width=4048"
            Splits(0)._ColumnProps(140)=   "Column(33).DividerColor=0"
            Splits(0)._ColumnProps(141)=   "Column(33)._WidthInPix=3942"
            Splits(0)._ColumnProps(142)=   "Column(33).Visible=0"
            Splits(0)._ColumnProps(143)=   "Column(33).Order=34"
            Splits(0)._ColumnProps(144)=   "Column(34).Width=2831"
            Splits(0)._ColumnProps(145)=   "Column(34).DividerColor=0"
            Splits(0)._ColumnProps(146)=   "Column(34)._WidthInPix=2725"
            Splits(0)._ColumnProps(147)=   "Column(34).Visible=0"
            Splits(0)._ColumnProps(148)=   "Column(34).Order=35"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            Appearance      =   0
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            WrapCellPointer =   -1  'True
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   12577002
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
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
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=82,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=174,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=171,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=172,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=173,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=106,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=103,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=104,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=105,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=102,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=98,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=95,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=96,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=97,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=94,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=91,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=92,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=93,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=58,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=55,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=56,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=57,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=28,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=25,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=26,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=27,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=54,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=51,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=52,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=53,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=50,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=47,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=48,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=49,.parent=17"
            _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=46,.parent=13"
            _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=43,.parent=14"
            _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=44,.parent=15"
            _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=45,.parent=17"
            _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=32,.parent=13"
            _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=29,.parent=14"
            _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=30,.parent=15"
            _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=31,.parent=17"
            _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=90,.parent=13"
            _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=87,.parent=14"
            _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=88,.parent=15"
            _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=89,.parent=17"
            _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=110,.parent=13"
            _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=107,.parent=14"
            _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=108,.parent=15"
            _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=109,.parent=17"
            _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=114,.parent=13"
            _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
            _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
            _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
            _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=118,.parent=13"
            _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=115,.parent=14"
            _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=116,.parent=15"
            _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=117,.parent=17"
            _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=122,.parent=13"
            _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=14"
            _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=15"
            _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=17"
            _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=126,.parent=13"
            _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=123,.parent=14"
            _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=124,.parent=15"
            _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=125,.parent=17"
            _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=130,.parent=13"
            _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=127,.parent=14"
            _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=128,.parent=15"
            _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=129,.parent=17"
            _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=134,.parent=13"
            _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=131,.parent=14"
            _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=132,.parent=15"
            _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=133,.parent=17"
            _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=138,.parent=13"
            _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=135,.parent=14"
            _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=136,.parent=15"
            _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=137,.parent=17"
            _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=142,.parent=13"
            _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=139,.parent=14"
            _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=140,.parent=15"
            _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=141,.parent=17"
            _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=146,.parent=13"
            _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=143,.parent=14"
            _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=144,.parent=15"
            _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=145,.parent=17"
            _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=150,.parent=13"
            _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=14"
            _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=15"
            _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=17"
            _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=154,.parent=13"
            _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=151,.parent=14"
            _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=152,.parent=15"
            _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=153,.parent=17"
            _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=158,.parent=13"
            _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=155,.parent=14"
            _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=156,.parent=15"
            _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=157,.parent=17"
            _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=170,.parent=13"
            _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=167,.parent=14"
            _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=168,.parent=15"
            _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=169,.parent=17"
            _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=70,.parent=13"
            _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=67,.parent=14"
            _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=68,.parent=15"
            _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=69,.parent=17"
            _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=162,.parent=13"
            _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=159,.parent=14"
            _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=160,.parent=15"
            _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=161,.parent=17"
            _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=166,.parent=13"
            _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=163,.parent=14"
            _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=164,.parent=15"
            _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=165,.parent=17"
            _StyleDefs(176) =   "Named:id=33:Normal"
            _StyleDefs(177) =   ":id=33,.parent=0"
            _StyleDefs(178) =   "Named:id=34:Heading"
            _StyleDefs(179) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(180) =   ":id=34,.wraptext=-1"
            _StyleDefs(181) =   "Named:id=35:Footing"
            _StyleDefs(182) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(183) =   "Named:id=36:Selected"
            _StyleDefs(184) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(185) =   "Named:id=37:Caption"
            _StyleDefs(186) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(187) =   "Named:id=38:HighlightRow"
            _StyleDefs(188) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(189) =   "Named:id=39:EvenRow"
            _StyleDefs(190) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(191) =   "Named:id=40:OddRow"
            _StyleDefs(192) =   ":id=40,.parent=33"
            _StyleDefs(193) =   "Named:id=41:RecordSelector"
            _StyleDefs(194) =   ":id=41,.parent=34"
            _StyleDefs(195) =   "Named:id=42:FilterBar"
            _StyleDefs(196) =   ":id=42,.parent=33"
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   735
            Left            =   9360
            Picture         =   "FormReminder1000.frx":EF32
            Style           =   1  'Graphical
            TabIndex        =   132
            Top             =   3420
            Width           =   735
         End
         Begin VB.CheckBox chGSM 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   127
            Top             =   3480
            Width           =   2535
         End
         Begin VB.CheckBox chCDMA 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6000
            TabIndex        =   126
            Top             =   3840
            Width           =   2775
         End
         Begin VB.CommandButton cmdSMS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SEND"
            Height          =   735
            Left            =   10260
            Picture         =   "FormReminder1000.frx":F5FE
            Style           =   1  'Graphical
            TabIndex        =   125
            Top             =   3420
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox fot 
            Caption         =   "Footer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   121
            Top             =   3060
            Width           =   3255
         End
         Begin VB.TextBox txtPesan 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   5160
            MaxLength       =   1000
            MultiLine       =   -1  'True
            TabIndex        =   120
            Top             =   2160
            Width           =   4935
         End
         Begin VB.CommandButton cmdTemplate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Template"
            Height          =   732
            Left            =   10260
            Picture         =   "FormReminder1000.frx":F980
            Style           =   1  'Graphical
            TabIndex        =   119
            Top             =   2220
            Width           =   855
         End
         Begin VB.CheckBox onC 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Call By Phone"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9240
            TabIndex        =   118
            Top             =   5000
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox onS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8040
            TabIndex        =   117
            Top             =   5000
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CommandButton CmdXls 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Grid"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9480
            Picture         =   "FormReminder1000.frx":FFD5
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmdXlsMRS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MRS"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8760
            Picture         =   "FormReminder1000.frx":10451
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cboType 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   600
            Width           =   2055
         End
         Begin VB.CommandButton cmdProcess 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pr&ocess"
            Height          =   732
            Left            =   5400
            Picture         =   "FormReminder1000.frx":108CD
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   480
            Width           =   732
         End
         Begin VB.Frame Frame1 
            Caption         =   " BPK "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   120
            TabIndex        =   43
            Top             =   10560
            Width           =   4215
            Begin TDBDate6Ctl.TDBDate DtDO 
               Height          =   330
               Left            =   1680
               TabIndex        =   44
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormReminder1000.frx":10F3A
               Caption         =   "FormReminder1000.frx":11066
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormReminder1000.frx":110D2
               Keys            =   "FormReminder1000.frx":110F0
               Spin            =   "FormReminder1000.frx":1114E
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
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
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate DtDEC 
               Height          =   330
               Left            =   1680
               TabIndex        =   45
               Top             =   720
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormReminder1000.frx":11176
               Caption         =   "FormReminder1000.frx":112A2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormReminder1000.frx":1130E
               Keys            =   "FormReminder1000.frx":1132C
               Spin            =   "FormReminder1000.frx":1138A
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
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
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl DEC"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   49
               Top             =   720
               Width           =   765
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl DO / BPK"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   48
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   47
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   46
               Top             =   735
               Width           =   75
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   " CR 1 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   4320
            TabIndex        =   38
            Top             =   10560
            Width           =   3495
            Begin VB.TextBox txtKMCR1 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   15
               TabIndex        =   39
               Top             =   720
               Width           =   1575
            End
            Begin TDBDate6Ctl.TDBDate DTCR1 
               Height          =   330
               Left            =   1680
               TabIndex        =   40
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormReminder1000.frx":113B2
               Caption         =   "FormReminder1000.frx":114DE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormReminder1000.frx":1154A
               Keys            =   "FormReminder1000.frx":11568
               Spin            =   "FormReminder1000.frx":115C6
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
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
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Validasi     :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   240
               TabIndex        =   42
               Top             =   360
               Width           =   1245
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angka KM       :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   270
               TabIndex        =   41
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   " CR 2 "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   7800
            TabIndex        =   31
            Top             =   10560
            Width           =   3495
            Begin VB.TextBox txtKMCR2 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   15
               TabIndex        =   32
               Top             =   720
               Width           =   1575
            End
            Begin TDBDate6Ctl.TDBDate DTCR2 
               Height          =   330
               Left            =   1680
               TabIndex        =   33
               Top             =   360
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calendar        =   "FormReminder1000.frx":115EE
               Caption         =   "FormReminder1000.frx":1171A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormReminder1000.frx":11786
               Keys            =   "FormReminder1000.frx":117A4
               Spin            =   "FormReminder1000.frx":11802
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
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
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1440
               TabIndex        =   37
               Top             =   735
               Width           =   75
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1440
               TabIndex        =   36
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angka KM"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   35
               Top             =   720
               Width           =   960
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Validasi"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   34
               Top             =   360
               Width           =   1110
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   " FAKTUR DATA STNK "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   2055
            Left            =   120
            TabIndex        =   5
            Top             =   11760
            Width           =   11175
            Begin VB.TextBox txtNamaSTNK 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   50
               TabIndex        =   13
               Top             =   360
               Width           =   4455
            End
            Begin VB.TextBox txtAlamatSTNK 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   1680
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   12
               Top             =   720
               Width           =   4455
            End
            Begin VB.TextBox txtKotaSTNK 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   11
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtKodePosSTNK 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4920
               MaxLength       =   30
               TabIndex        =   10
               Top             =   1440
               Width           =   1215
            End
            Begin VB.TextBox txtTempatLahir 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7800
               MaxLength       =   15
               TabIndex        =   9
               Top             =   360
               Width           =   1665
            End
            Begin VB.TextBox txtPekerjaan 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7800
               MaxLength       =   15
               TabIndex        =   8
               Top             =   1080
               Width           =   3135
            End
            Begin VB.TextBox txtHoby 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7800
               MaxLength       =   15
               TabIndex        =   7
               Top             =   1440
               Width           =   3135
            End
            Begin VB.TextBox txtAgama 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7800
               MaxLength       =   15
               TabIndex        =   6
               Top             =   720
               Width           =   1665
            End
            Begin TDBDate6Ctl.TDBDate DTLahir 
               Height          =   330
               Left            =   9480
               TabIndex        =   14
               Top             =   360
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   582
               Calendar        =   "FormReminder1000.frx":1182A
               Caption         =   "FormReminder1000.frx":11956
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormReminder1000.frx":119C2
               Keys            =   "FormReminder1000.frx":119E0
               Spin            =   "FormReminder1000.frx":11A3E
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
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
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama STNK"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   30
               Top             =   360
               Width           =   1080
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alamat "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   29
               Top             =   720
               Width           =   750
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kota"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   28
               Top             =   1440
               Width           =   465
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   27
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   26
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Pos   :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3720
               TabIndex        =   25
               Top             =   1440
               Width           =   1140
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   24
               Top             =   1440
               Width           =   75
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pekerjaan"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6480
               TabIndex        =   23
               Top             =   1080
               Width           =   930
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Agama"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6480
               TabIndex        =   22
               Top             =   705
               Width           =   645
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TTL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6480
               TabIndex        =   21
               Top             =   360
               Width           =   345
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   8280
               TabIndex        =   20
               Top             =   1200
               Width           =   75
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   7680
               TabIndex        =   19
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   7680
               TabIndex        =   18
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label66 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hoby"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6480
               TabIndex        =   17
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   7680
               TabIndex        =   16
               Top             =   1440
               Width           =   75
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   7680
               TabIndex        =   15
               Top             =   1080
               Width           =   75
            End
         End
         Begin VB.CommandButton cmdTop 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Back To Top"
            Height          =   735
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FormReminder1000.frx":11A66
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   14280
            Width           =   11175
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   10440
            Picture         =   "FormReminder1000.frx":11DF1
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   480
            Width           =   735
         End
         Begin MSComctlLib.ImageList ImageList 
            Left            =   7920
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   145
            ImageHeight     =   78
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReminder1000.frx":12581
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReminder1000.frx":13A98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReminder1000.frx":15793
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate DtRemain 
            Height          =   315
            Left            =   3600
            TabIndex        =   85
            Top             =   270
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "FormReminder1000.frx":16E83
            Caption         =   "FormReminder1000.frx":16FAF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FormReminder1000.frx":1701B
            Keys            =   "FormReminder1000.frx":17039
            Spin            =   "FormReminder1000.frx":17097
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
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   2118189057
            Value           =   2.12482692446619E-314
            CenturyMode     =   0
         End
         Begin MSComDlg.CommonDialog ComDialogXls 
            Left            =   8280
            Top             =   840
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin TrueOleDBGrid70.TDBGrid GridXlsMrs 
            Height          =   495
            Left            =   9600
            TabIndex        =   110
            Top             =   600
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   873
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nomor"
            Columns(0).DataField=   "Nomor"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nama Customer"
            Columns(1).DataField=   "Nama_Customer"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tgl DO"
            Columns(2).DataField=   "Tgl_DO"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tgl CR1"
            Columns(3).DataField=   "Tgl_CR1"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Km CR1"
            Columns(4).DataField=   "Km_CR1"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tgl CR2"
            Columns(5).DataField=   "Tgl_CR2"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Km CR2"
            Columns(6).DataField=   "Km_CR2"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "No Polisi"
            Columns(7).DataField=   "No_Polisi"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "No Chassis"
            Columns(8).DataField=   "No_Rangka"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Telpon"
            Columns(9).DataField=   "Telp_Rumah"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "GSM"
            Columns(10).DataField=   "Telp_GSM"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "CDMA"
            Columns(11).DataField=   "Telp_CDMA"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Tgl ETA"
            Columns(12).DataField=   "Tgl_Eta"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "Tgl SMS"
            Columns(13).DataField=   "Tgl_SMS"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Tgl Service"
            Columns(14).DataField=   "Tgl_Service"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Tgl Follow"
            Columns(15).DataField=   "Tgl_Follow"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "Komentar"
            Columns(16).DataField=   "Komentar"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "Keterangan"
            Columns(17).DataField=   "Keterangan"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).Caption=   "Wiraniaga"
            Columns(18).DataField=   "Wiraniaga"
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   19
            Splits(0)._UserFlags=   0
            Splits(0).Locked=   -1  'True
            Splits(0).AllowSizing=   -1  'True
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColSelect=   0   'False
            Splits(0).FetchRowStyle=   -1  'True
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=19"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2646"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2540"
            Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=1"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=2646"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2540"
            Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1667"
            Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(24)=   "Column(5).Width=2646"
            Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2540"
            Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(29)=   "Column(6).Width=1773"
            Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=1667"
            Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(34)=   "Column(7).Width=2646"
            Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=2540"
            Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(38)=   "Column(8).Width=3889"
            Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=3784"
            Splits(0)._ColumnProps(41)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(42)=   "Column(9).Width=3281"
            Splits(0)._ColumnProps(43)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(44)=   "Column(9)._WidthInPix=3175"
            Splits(0)._ColumnProps(45)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(46)=   "Column(10).Width=3281"
            Splits(0)._ColumnProps(47)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(48)=   "Column(10)._WidthInPix=3175"
            Splits(0)._ColumnProps(49)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(50)=   "Column(11).Width=3281"
            Splits(0)._ColumnProps(51)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(11)._WidthInPix=3175"
            Splits(0)._ColumnProps(53)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(54)=   "Column(12).Width=3519"
            Splits(0)._ColumnProps(55)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(56)=   "Column(12)._WidthInPix=3413"
            Splits(0)._ColumnProps(57)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(58)=   "Column(13).Width=1773"
            Splits(0)._ColumnProps(59)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(60)=   "Column(13)._WidthInPix=1667"
            Splits(0)._ColumnProps(61)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(62)=   "Column(14).Width=3519"
            Splits(0)._ColumnProps(63)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(14)._WidthInPix=3413"
            Splits(0)._ColumnProps(65)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(66)=   "Column(15).Width=3281"
            Splits(0)._ColumnProps(67)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(68)=   "Column(15)._WidthInPix=3175"
            Splits(0)._ColumnProps(69)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(70)=   "Column(16).Width=3281"
            Splits(0)._ColumnProps(71)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(72)=   "Column(16)._WidthInPix=3175"
            Splits(0)._ColumnProps(73)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(74)=   "Column(17).Width=3281"
            Splits(0)._ColumnProps(75)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(76)=   "Column(17)._WidthInPix=3175"
            Splits(0)._ColumnProps(77)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(78)=   "Column(18).Width=5292"
            Splits(0)._ColumnProps(79)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(80)=   "Column(18)._WidthInPix=5186"
            Splits(0)._ColumnProps(81)=   "Column(18).Order=19"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            AllowArrows     =   0   'False
            WrapCellPointer =   -1  'True
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
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
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=82,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=106,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=103,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=104,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=105,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=102,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=50,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=46,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=43,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=44,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=45,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=29,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=30,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=31,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=110,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=107,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=108,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=109,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=114,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=111,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=112,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=113,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=118,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=115,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=116,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=117,.parent=17"
            _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=66,.parent=13"
            _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=63,.parent=14"
            _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=64,.parent=15"
            _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=65,.parent=17"
            _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=62,.parent=13"
            _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=59,.parent=14"
            _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=60,.parent=15"
            _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=61,.parent=17"
            _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=28,.parent=13"
            _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=25,.parent=14"
            _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=26,.parent=15"
            _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=27,.parent=17"
            _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=122,.parent=13"
            _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=119,.parent=14"
            _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=120,.parent=15"
            _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=121,.parent=17"
            _StyleDefs(112) =   "Named:id=33:Normal"
            _StyleDefs(113) =   ":id=33,.parent=0"
            _StyleDefs(114) =   "Named:id=34:Heading"
            _StyleDefs(115) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(116) =   ":id=34,.wraptext=-1"
            _StyleDefs(117) =   "Named:id=35:Footing"
            _StyleDefs(118) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(119) =   "Named:id=36:Selected"
            _StyleDefs(120) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(121) =   "Named:id=37:Caption"
            _StyleDefs(122) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(123) =   "Named:id=38:HighlightRow"
            _StyleDefs(124) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(125) =   "Named:id=39:EvenRow"
            _StyleDefs(126) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(127) =   "Named:id=40:OddRow"
            _StyleDefs(128) =   ":id=40,.parent=33"
            _StyleDefs(129) =   "Named:id=41:RecordSelector"
            _StyleDefs(130) =   ":id=41,.parent=34"
            _StyleDefs(131) =   "Named:id=42:FilterBar"
            _StyleDefs(132) =   ":id=42,.parent=33"
         End
         Begin TDBDate6Ctl.TDBDate DtRemain0 
            Height          =   315
            Left            =   1800
            TabIndex        =   114
            Top             =   270
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "FormReminder1000.frx":170BF
            Caption         =   "FormReminder1000.frx":171EB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FormReminder1000.frx":17257
            Keys            =   "FormReminder1000.frx":17275
            Spin            =   "FormReminder1000.frx":172D3
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
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   2118189057
            Value           =   2.12482692446619E-314
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate DTSend 
            Height          =   330
            Left            =   7200
            TabIndex        =   133
            Top             =   4280
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   582
            Calendar        =   "FormReminder1000.frx":172FB
            Caption         =   "FormReminder1000.frx":17427
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FormReminder1000.frx":17493
            Keys            =   "FormReminder1000.frx":174B1
            Spin            =   "FormReminder1000.frx":1750F
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd/mm/yyyy"
            EditMode        =   0
            Enabled         =   0
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
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   2118189057
            Value           =   2.12482692446619E-314
            CenturyMode     =   0
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   120
            TabIndex        =   94
            Top             =   9360
            Width           =   11175
            Begin VB.TextBox txtNamaType 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   98
               Top             =   720
               Width           =   3855
            End
            Begin VB.TextBox txtNamaWarna 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7440
               MaxLength       =   100
               TabIndex        =   97
               Top             =   720
               Width           =   3495
            End
            Begin VB.TextBox txtKodeWarna 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7440
               MaxLength       =   100
               TabIndex        =   96
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtKodeType 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1680
               MaxLength       =   100
               TabIndex        =   95
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Warna"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6000
               TabIndex        =   108
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   7320
               TabIndex        =   107
               Top             =   735
               Width           =   75
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   106
               Top             =   720
               Width           =   75
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   105
               Top             =   720
               Width           =   1050
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   1560
               TabIndex        =   102
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   7320
               TabIndex        =   101
               Top             =   375
               Width           =   75
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   100
               Top             =   360
               Width           =   1020
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Warna"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6000
               TabIndex        =   99
               Top             =   360
               Width           =   1125
            End
         End
         Begin MSMAPI.MAPISession MAPISession1 
            Left            =   7560
            Top             =   360
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DownloadMail    =   -1  'True
            LogonUI         =   -1  'True
            NewSession      =   0   'False
         End
         Begin MSMAPI.MAPIMessages MAPIMessages1 
            Left            =   7680
            Top             =   600
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            AddressEditFieldCount=   1
            AddressModifiable=   0   'False
            AddressResolveUI=   0   'False
            FetchSorted     =   0   'False
            FetchUnreadOnly =   0   'False
         End
         Begin TabDlg.SSTab sService 
            Height          =   2415
            Left            =   120
            TabIndex        =   143
            Top             =   5280
            Visible         =   0   'False
            Width           =   11145
            _ExtentX        =   19659
            _ExtentY        =   4260
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   5
            TabHeight       =   706
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   ".:: Customer ::."
            TabPicture(0)   =   "FormReminder1000.frx":17537
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Line32"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Line34"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Line35"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Line36"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label82"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label83"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label84"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label85"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label86"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label87"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "Label88"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "Label89"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "Label92"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "Label93"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "Label94"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "Label95"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "Label96"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "Label97"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "Label100"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "Label101"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "Label103"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "Line37"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "Line38"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "Label104"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "Label43(2)"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "Label3"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "Label5"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "nEngine"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "nTglSvc"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "nkmSvc"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).Control(30)=   "nAlamat"
            Tab(0).Control(30).Enabled=   0   'False
            Tab(0).Control(31)=   "nNama"
            Tab(0).Control(31).Enabled=   0   'False
            Tab(0).Control(32)=   "nNoPol"
            Tab(0).Control(32).Enabled=   0   'False
            Tab(0).Control(33)=   "nChassis"
            Tab(0).Control(33).Enabled=   0   'False
            Tab(0).Control(34)=   "nGSM"
            Tab(0).Control(34).Enabled=   0   'False
            Tab(0).Control(35)=   "nPSTN"
            Tab(0).Control(35).Enabled=   0   'False
            Tab(0).Control(36)=   "nCDMA"
            Tab(0).Control(36).Enabled=   0   'False
            Tab(0).Control(37)=   "txtNoWO(0)"
            Tab(0).Control(37).Enabled=   0   'False
            Tab(0).Control(38)=   "txtNoInv(0)"
            Tab(0).Control(38).Enabled=   0   'False
            Tab(0).Control(39)=   "txtEmail"
            Tab(0).Control(39).Enabled=   0   'False
            Tab(0).ControlCount=   40
            TabCaption(1)   =   ".:: Job Order ::."
            TabPicture(1)   =   "FormReminder1000.frx":17553
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label43(3)"
            Tab(1).Control(1)=   "Line39"
            Tab(1).Control(2)=   "Line40"
            Tab(1).Control(3)=   "Line41"
            Tab(1).Control(4)=   "Line42"
            Tab(1).Control(5)=   "Line43"
            Tab(1).Control(6)=   "Line44"
            Tab(1).Control(7)=   "tRte(7)"
            Tab(1).Control(8)=   "tRte(6)"
            Tab(1).Control(9)=   "tRte(5)"
            Tab(1).Control(10)=   "tRte(4)"
            Tab(1).Control(11)=   "tRep(7)"
            Tab(1).Control(12)=   "tRep(6)"
            Tab(1).Control(13)=   "tRep(5)"
            Tab(1).Control(14)=   "tRep(4)"
            Tab(1).Control(15)=   "tRte(3)"
            Tab(1).Control(16)=   "tRte(2)"
            Tab(1).Control(17)=   "tRte(1)"
            Tab(1).Control(18)=   "tRte(0)"
            Tab(1).Control(19)=   "tRep(3)"
            Tab(1).Control(20)=   "tRep(2)"
            Tab(1).Control(21)=   "tRep(1)"
            Tab(1).Control(22)=   "tRep(0)"
            Tab(1).Control(23)=   "txtNoInv(2)"
            Tab(1).Control(24)=   "txtNoWO(2)"
            Tab(1).ControlCount=   25
            Begin VB.TextBox txtEmail 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   330
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   213
               Top             =   2040
               Width           =   3255
            End
            Begin VB.TextBox txtNoWO 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   -65880
               MaxLength       =   30
               TabIndex        =   207
               Top             =   60
               Width           =   1575
            End
            Begin VB.TextBox txtNoInv 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   -68640
               MaxLength       =   30
               TabIndex        =   206
               Top             =   60
               Width           =   1575
            End
            Begin VB.TextBox txtNoInv 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   6360
               MaxLength       =   30
               TabIndex        =   173
               Top             =   60
               Width           =   1575
            End
            Begin VB.TextBox txtNoWO 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   9120
               MaxLength       =   30
               TabIndex        =   172
               Top             =   60
               Width           =   1575
            End
            Begin VB.TextBox txtNoInv 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   -67920
               MaxLength       =   30
               TabIndex        =   171
               Top             =   0
               Width           =   1575
            End
            Begin VB.TextBox txtNoWO 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   -65160
               MaxLength       =   30
               TabIndex        =   170
               Top             =   0
               Width           =   1575
            End
            Begin VB.TextBox nCDMA 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9120
               MaxLength       =   25
               TabIndex        =   169
               Top             =   1800
               Width           =   1815
            End
            Begin VB.TextBox nPSTN 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   330
               Left            =   9120
               MaxLength       =   40
               TabIndex        =   168
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox nGSM 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9120
               MaxLength       =   25
               TabIndex        =   167
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox nChassis 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1440
               MaxLength       =   25
               TabIndex        =   166
               Top             =   1680
               Width           =   3255
            End
            Begin VB.TextBox nNoPol 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6360
               MaxLength       =   15
               TabIndex        =   165
               Top             =   2040
               Width           =   1575
            End
            Begin VB.TextBox nNama 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1440
               MaxLength       =   100
               TabIndex        =   164
               Top             =   600
               Width           =   6495
            End
            Begin VB.TextBox nAlamat 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1440
               MaxLength       =   200
               TabIndex        =   163
               Top             =   960
               Width           =   6495
            End
            Begin VB.TextBox nkmSvc 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6360
               MaxLength       =   25
               TabIndex        =   162
               Top             =   1680
               Width           =   1575
            End
            Begin VB.TextBox nTglSvc 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   330
               Left            =   6360
               MaxLength       =   40
               TabIndex        =   161
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox nEngine 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1440
               MaxLength       =   25
               TabIndex        =   160
               Top             =   1320
               Width           =   3255
            End
            Begin VB.TextBox tRep 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   -74760
               MaxLength       =   100
               TabIndex        =   159
               Top             =   720
               Width           =   4670
            End
            Begin VB.TextBox tRep 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   -74760
               MaxLength       =   100
               TabIndex        =   158
               Top             =   1080
               Width           =   4670
            End
            Begin VB.TextBox tRep 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   -74760
               MaxLength       =   100
               TabIndex        =   157
               Top             =   1440
               Width           =   4670
            End
            Begin VB.TextBox tRep 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   -74760
               MaxLength       =   100
               TabIndex        =   156
               Top             =   1800
               Width           =   4670
            End
            Begin VB.TextBox tRte 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   0
               Left            =   -70080
               MaxLength       =   25
               TabIndex        =   155
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox tRte 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   1
               Left            =   -70080
               MaxLength       =   25
               TabIndex        =   154
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox tRte 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   2
               Left            =   -70080
               MaxLength       =   25
               TabIndex        =   153
               Top             =   1440
               Width           =   495
            End
            Begin VB.TextBox tRte 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   3
               Left            =   -70080
               MaxLength       =   25
               TabIndex        =   152
               Top             =   1800
               Width           =   495
            End
            Begin VB.TextBox tRep 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   -69240
               MaxLength       =   100
               TabIndex        =   151
               Top             =   720
               Width           =   4670
            End
            Begin VB.TextBox tRep 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   -69240
               MaxLength       =   100
               TabIndex        =   150
               Top             =   1080
               Width           =   4670
            End
            Begin VB.TextBox tRep 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   6
               Left            =   -69240
               MaxLength       =   100
               TabIndex        =   149
               Top             =   1440
               Width           =   4670
            End
            Begin VB.TextBox tRep 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   -69240
               MaxLength       =   100
               TabIndex        =   148
               Top             =   1800
               Width           =   4670
            End
            Begin VB.TextBox tRte 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   4
               Left            =   -64560
               MaxLength       =   25
               TabIndex        =   147
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox tRte 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   -64560
               MaxLength       =   25
               TabIndex        =   146
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox tRte 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   6
               Left            =   -64560
               MaxLength       =   25
               TabIndex        =   145
               Top             =   1440
               Width           =   495
            End
            Begin VB.TextBox tRte 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   -64560
               MaxLength       =   25
               TabIndex        =   144
               Top             =   1800
               Width           =   495
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email Cust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   215
               Top             =   2040
               Width           =   900
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               TabIndex        =   214
               Top             =   2040
               Width           =   75
            End
            Begin VB.Line Line44 
               BorderColor     =   &H000000FF&
               X1              =   -70680
               X2              =   -70680
               Y1              =   480
               Y2              =   240
            End
            Begin VB.Line Line43 
               BorderColor     =   &H000000FF&
               X1              =   -70680
               X2              =   -61320
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Line Line42 
               BorderColor     =   &H000000FF&
               X1              =   -70680
               X2              =   -70920
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line41 
               BorderColor     =   &H000000FF&
               X1              =   -72480
               X2              =   -72720
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line40 
               BorderColor     =   &H000000FF&
               X1              =   -72720
               X2              =   -72720
               Y1              =   240
               Y2              =   480
            End
            Begin VB.Line Line39 
               BorderColor     =   &H000000FF&
               X1              =   -72720
               X2              =   -75120
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Invoice  :                                              No WO  :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   -69720
               TabIndex        =   208
               Top             =   120
               Width           =   3795
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Invoice  :                                              No WO  :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   5040
               TabIndex        =   205
               Top             =   120
               Width           =   3795
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   -65280
               TabIndex        =   204
               Top             =   30
               Width           =   75
            End
            Begin VB.Label Label152 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   -68040
               TabIndex        =   203
               Top             =   30
               Width           =   75
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Invoice"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   -69000
               TabIndex        =   202
               Top             =   30
               Width           =   915
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No WO"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   -66000
               TabIndex        =   201
               Top             =   30
               Width           =   645
            End
            Begin VB.Label Label104 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ".:: TELPON ::."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   8160
               TabIndex        =   200
               Top             =   675
               Width           =   2775
            End
            Begin VB.Line Line38 
               X1              =   8160
               X2              =   10920
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Line Line37 
               X1              =   8160
               X2              =   10920
               Y1              =   960
               Y2              =   960
            End
            Begin VB.Label Label103 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CDMA       :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8160
               TabIndex        =   199
               Top             =   1800
               Width           =   870
            End
            Begin VB.Label Label101 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PSTN        :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8160
               TabIndex        =   198
               Top             =   1080
               Width           =   870
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GSM          :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8145
               TabIndex        =   197
               Top             =   1440
               Width           =   885
            End
            Begin VB.Label Label97 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Chassis"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   196
               Top             =   1680
               Width           =   1065
            End
            Begin VB.Label Label96 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               TabIndex        =   195
               Top             =   1680
               Width           =   75
            End
            Begin VB.Label Label95 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               TabIndex        =   194
               Top             =   960
               Width           =   75
            End
            Begin VB.Label Label94 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               TabIndex        =   193
               Top             =   600
               Width           =   75
            End
            Begin VB.Label Label93 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alamat Cust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   192
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label Label92 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama  Cust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   191
               Top             =   600
               Width           =   1110
            End
            Begin VB.Label Label89 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nomor Polisi"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   4920
               TabIndex        =   190
               Top             =   2040
               Width           =   1245
            End
            Begin VB.Label Label88 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6240
               TabIndex        =   189
               Top             =   2040
               Width           =   75
            End
            Begin VB.Label Label87 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6240
               TabIndex        =   188
               Top             =   1680
               Width           =   75
            End
            Begin VB.Label Label86 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6240
               TabIndex        =   187
               Top             =   1320
               Width           =   75
            End
            Begin VB.Label Label85 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Km Service"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   4920
               TabIndex        =   186
               Top             =   1680
               Width           =   1065
            End
            Begin VB.Label Label84 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Service"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   4920
               TabIndex        =   185
               Top             =   1320
               Width           =   1065
            End
            Begin VB.Label Label83 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Engine"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   184
               Top             =   1320
               Width           =   990
            End
            Begin VB.Label Label82 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1320
               TabIndex        =   183
               Top             =   1320
               Width           =   75
            End
            Begin VB.Line Line36 
               BorderColor     =   &H000000FF&
               X1              =   2160
               X2              =   11520
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Line Line35 
               BorderColor     =   &H000000FF&
               X1              =   2160
               X2              =   2160
               Y1              =   480
               Y2              =   240
            End
            Begin VB.Line Line34 
               BorderColor     =   &H000000FF&
               X1              =   2160
               X2              =   1920
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line33 
               BorderColor     =   &H000000FF&
               X1              =   -840
               X2              =   -1080
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line32 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   120
               Y1              =   240
               Y2              =   480
            End
            Begin VB.Line Line31 
               BorderColor     =   &H000000FF&
               X1              =   -70560
               X2              =   -61200
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Line Line30 
               BorderColor     =   &H000000FF&
               X1              =   -70560
               X2              =   -70560
               Y1              =   480
               Y2              =   240
            End
            Begin VB.Line Line29 
               BorderColor     =   &H000000FF&
               X1              =   -70560
               X2              =   -70800
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line28 
               BorderColor     =   &H000000FF&
               X1              =   -72360
               X2              =   -72600
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line27 
               BorderColor     =   &H000000FF&
               X1              =   -72600
               X2              =   -72600
               Y1              =   240
               Y2              =   480
            End
            Begin VB.Line Line26 
               BorderColor     =   &H000000FF&
               X1              =   -72600
               X2              =   -75000
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label81 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "1."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -74760
               TabIndex        =   182
               Top             =   840
               Width           =   150
            End
            Begin VB.Label Label80 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "2."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -74760
               TabIndex        =   181
               Top             =   1200
               Width           =   150
            End
            Begin VB.Label Label79 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "3."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -74760
               TabIndex        =   180
               Top             =   1560
               Width           =   150
            End
            Begin VB.Label Label78 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "4."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -74760
               TabIndex        =   179
               Top             =   1920
               Width           =   150
            End
            Begin VB.Label Label77 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "6."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -69000
               TabIndex        =   178
               Top             =   1200
               Width           =   150
            End
            Begin VB.Label Label76 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "7."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -69000
               TabIndex        =   177
               Top             =   1560
               Width           =   150
            End
            Begin VB.Label Label75 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "8."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -69000
               TabIndex        =   176
               Top             =   1920
               Width           =   150
            End
            Begin VB.Label Label72 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "5."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -69000
               TabIndex        =   175
               Top             =   840
               Width           =   150
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   $"FormReminder1000.frx":1756F
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   -74880
               TabIndex        =   174
               Top             =   540
               Width           =   11175
            End
            Begin VB.Line Line24 
               BorderStyle     =   3  'Dot
               X1              =   -74880
               X2              =   -63600
               Y1              =   770
               Y2              =   770
            End
         End
         Begin VB.Frame FrameGRP 
            Caption         =   ".:: Data Surat Menyurat ::."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   3375
            Left            =   120
            TabIndex        =   50
            Top             =   5280
            Width           =   11205
            Begin VB.CheckBox chDM 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Direct Mail"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6480
               TabIndex        =   211
               Top             =   2520
               Width           =   1695
            End
            Begin VB.CommandButton CmdOpen 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Edit Cust"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   3960
               Picture         =   "FormReminder1000.frx":17648
               Style           =   1  'Graphical
               TabIndex        =   136
               Top             =   2520
               Width           =   975
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H80000005&
               Caption         =   " Wiraniaga : "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   0
               TabIndex        =   103
               Top             =   2480
               Width           =   3735
               Begin VB.TextBox txtwiraniaga 
                  Appearance      =   0  'Flat
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   120
                  MaxLength       =   50
                  TabIndex        =   104
                  Top             =   360
                  Width           =   3525
               End
            End
            Begin VB.TextBox txtCDMA 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9240
               MaxLength       =   25
               TabIndex        =   63
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox txtTelp 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   330
               Left            =   9240
               MaxLength       =   40
               TabIndex        =   62
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox txtGSM 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9240
               MaxLength       =   25
               TabIndex        =   61
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtNoChassis 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               MaxLength       =   25
               TabIndex        =   60
               Top             =   1800
               Width           =   2655
            End
            Begin VB.TextBox txtNoPol 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6360
               MaxLength       =   15
               TabIndex        =   59
               Top             =   1800
               Width           =   1575
            End
            Begin VB.TextBox txtNama 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               MaxLength       =   100
               TabIndex        =   58
               Top             =   360
               Width           =   6015
            End
            Begin VB.TextBox txtAlamat 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               MaxLength       =   200
               TabIndex        =   57
               Top             =   720
               Width           =   6015
            End
            Begin VB.TextBox txtEmailAdr 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               MaxLength       =   25
               TabIndex        =   56
               Top             =   1440
               Width           =   6015
            End
            Begin VB.CommandButton cmdCall 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&CALL"
               Height          =   735
               Left            =   5200
               Picture         =   "FormReminder1000.frx":17C41
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   2520
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CheckBox chSMS 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SMS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6480
               TabIndex        =   54
               Top             =   2760
               Width           =   1695
            End
            Begin VB.CheckBox chCall 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Call By Phone"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   6480
               TabIndex        =   53
               Top             =   3000
               Width           =   1695
            End
            Begin VB.TextBox txtKota 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               MaxLength       =   25
               TabIndex        =   52
               Top             =   1080
               Width           =   2655
            End
            Begin VB.TextBox txtPos 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6360
               MaxLength       =   25
               TabIndex        =   51
               Top             =   1080
               Width           =   1575
            End
            Begin TDBDate6Ctl.TDBDate txtRem 
               Height          =   330
               Left            =   9240
               TabIndex        =   115
               Top             =   1800
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               Calendar        =   "FormReminder1000.frx":17F61
               Caption         =   "FormReminder1000.frx":1808D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormReminder1000.frx":180F9
               Keys            =   "FormReminder1000.frx":18117
               Spin            =   "FormReminder1000.frx":18175
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd/mm/yyyy"
               EditMode        =   0
               Enabled         =   0
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
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   2118189057
               Value           =   2.12482692446619E-314
               CenturyMode     =   0
            End
            Begin VB.CommandButton cmdDM 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Create DM"
               Height          =   735
               Left            =   5200
               Picture         =   "FormReminder1000.frx":1819D
               Style           =   1  'Graphical
               TabIndex        =   209
               Top             =   2520
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label nDM 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ".  .  .  .  .  .  .  .  ."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   8400
               TabIndex        =   212
               Top             =   2520
               Width           =   1500
            End
            Begin VB.Shape Shape10 
               BorderColor     =   &H00808080&
               FillColor       =   &H00FFFFFF&
               Height          =   855
               Left            =   3840
               Shape           =   4  'Rounded Rectangle
               Top             =   2460
               Width           =   1215
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   9120
               TabIndex        =   113
               Top             =   1800
               Width           =   75
            End
            Begin VB.Label Label48 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tgl Svc"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   8280
               TabIndex        =   112
               Top             =   1800
               Width           =   615
            End
            Begin VB.Line Line9 
               BorderColor     =   &H000000FF&
               X1              =   8280
               X2              =   11040
               Y1              =   1710
               Y2              =   1710
            End
            Begin VB.Line Line11 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   11160
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line12 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   11160
               Y1              =   2205
               Y2              =   2205
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ".:: TELPON ::."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   8280
               TabIndex        =   82
               Top             =   265
               Width           =   2775
            End
            Begin VB.Line Line15 
               BorderColor     =   &H000000FF&
               X1              =   8280
               X2              =   11040
               Y1              =   520
               Y2              =   520
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CDMA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   8280
               TabIndex        =   81
               Top             =   1320
               Width           =   600
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   9120
               TabIndex        =   80
               Top             =   1320
               Width           =   75
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PSTN"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   8280
               TabIndex        =   79
               Top             =   600
               Width           =   480
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GSM"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   8280
               TabIndex        =   78
               Top             =   960
               Width           =   435
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   9120
               TabIndex        =   77
               Top             =   600
               Width           =   75
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   9120
               TabIndex        =   76
               Top             =   960
               Width           =   75
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Chassis"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   75
               Top             =   1800
               Width           =   1065
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alamat Cust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   74
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama  Cust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   73
               Top             =   360
               Width           =   1110
            End
            Begin VB.Label Label91 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "No Polisi"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5280
               TabIndex        =   72
               Top             =   1800
               Width           =   855
            End
            Begin VB.Label Label90 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6240
               TabIndex        =   71
               Top             =   1800
               Width           =   75
            End
            Begin VB.Label lblMerk 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2520
               TabIndex        =   70
               Top             =   360
               Width           =   75
            End
            Begin VB.Label Label73 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email Cust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   240
               TabIndex        =   69
               Top             =   1440
               Width           =   900
            End
            Begin VB.Line Line17 
               BorderColor     =   &H000000FF&
               X1              =   120
               X2              =   11160
               Y1              =   2280
               Y2              =   2280
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kota"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   240
               TabIndex        =   68
               Top             =   1080
               Width           =   465
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   6240
               TabIndex        =   67
               Top             =   1080
               Width           =   75
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Pos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5280
               TabIndex        =   66
               Top             =   1080
               Width           =   900
            End
            Begin VB.Label nSMS 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ".  .  .  .  .  .  .  .  ."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   8400
               TabIndex        =   65
               Top             =   2790
               Width           =   1500
            End
            Begin VB.Label nCall 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ".  .  .  .  .  .  .  .  ."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   8400
               TabIndex        =   64
               Top             =   3030
               Width           =   1500
            End
            Begin VB.Shape Shape7 
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               Height          =   840
               Left            =   6360
               Top             =   2460
               Width           =   4695
            End
            Begin VB.Shape Shape5 
               BorderColor     =   &H00808080&
               FillColor       =   &H00FFFFFF&
               Height          =   855
               Left            =   5090
               Shape           =   4  'Rounded Rectangle
               Top             =   2460
               Visible         =   0   'False
               Width           =   1200
            End
         End
         Begin VB.Line Line25 
            BorderColor     =   &H00808080&
            X1              =   5160
            X2              =   11160
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Data  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   9000
            TabIndex        =   140
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label lblSave 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save Item : 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   9360
            TabIndex        =   139
            Top             =   4320
            Width           =   1725
         End
         Begin VB.Label Label69 
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1695
            TabIndex        =   138
            Top             =   655
            Width           =   150
         End
         Begin VB.Label Label60 
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   137
            Top             =   280
            Width           =   150
         End
         Begin VB.Line Line23 
            BorderColor     =   &H00808080&
            X1              =   5160
            X2              =   11160
            Y1              =   4845
            Y2              =   4845
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Kirim SMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5160
            TabIndex        =   135
            Top             =   4320
            Width           =   1485
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6840
            TabIndex        =   134
            Top             =   4320
            Width           =   75
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   855
            Left            =   10200
            Top             =   3360
            Width           =   975
         End
         Begin VB.Line Line14 
            BorderColor     =   &H00808080&
            X1              =   5160
            X2              =   10080
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   131
            Top             =   3480
            Width           =   75
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5880
            TabIndex        =   129
            Top             =   3840
            Width           =   75
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00808080&
            X1              =   5160
            X2              =   10080
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00808080&
            Height          =   855
            Left            =   10200
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   124
            Top             =   1920
            Width           =   75
         End
         Begin VB.Label Label51 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Isi Pesan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5160
            TabIndex        =   123
            Top             =   1920
            Width           =   840
         End
         Begin VB.Label lbLen 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Char :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8760
            TabIndex        =   122
            Top             =   3060
            Width           =   1335
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000FF&
            X1              =   5040
            X2              =   5040
            Y1              =   2040
            Y2              =   4800
         End
         Begin VB.Shape ShapeXLS 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   9360
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblPart 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "DATA KENDARAAN - STNK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   92
            Top             =   8955
            Width           =   10695
         End
         Begin VB.Label Label1 
            Caption         =   "Tgl Reminder                                        s/d"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   90
            Top             =   300
            Width           =   3600
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   5160
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label Label2 
            Caption         =   "Type Reminder"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   89
            Top             =   675
            Width           =   2415
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   255
            X2              =   5160
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   255
            X2              =   5160
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   5145
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Image ImgView 
            Height          =   1095
            Left            =   6360
            Top             =   240
            Width           =   2295
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            X1              =   480
            X2              =   4095
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            X1              =   1320
            X2              =   4080
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   5295
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblKet 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   88
            Top             =   1035
            Width           =   4695
         End
         Begin VB.Line Line19 
            BorderColor     =   &H000000FF&
            X1              =   1440
            X2              =   3960
            Y1              =   9060
            Y2              =   9060
         End
         Begin VB.Line Line20 
            BorderColor     =   &H000000FF&
            X1              =   720
            X2              =   3960
            Y1              =   9165
            Y2              =   9165
         End
         Begin VB.Line Line16 
            BorderColor     =   &H000000FF&
            X1              =   2400
            X2              =   11400
            Y1              =   14130
            Y2              =   14130
         End
         Begin VB.Line Line21 
            BorderColor     =   &H000000FF&
            X1              =   1320
            X2              =   10320
            Y1              =   14025
            Y2              =   14025
         End
         Begin VB.Line Line22 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   9120
            Y1              =   13920
            Y2              =   13920
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   10320
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblJml 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10320
            TabIndex        =   87
            Top             =   1560
            Width           =   855
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0FF&
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   8880
            Width           =   11220
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "View List Reminder :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5160
            TabIndex        =   116
            Top             =   4995
            Width           =   1980
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GSM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5160
            TabIndex        =   130
            Top             =   3480
            Width           =   1035
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CDMA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5160
            TabIndex        =   128
            Top             =   3840
            Width           =   1080
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "LIST REMINDER 1000 KM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   91
            Top             =   1560
            Width           =   10335
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   495
            Left            =   135
            Shape           =   4  'Rounded Rectangle
            Top             =   1440
            Width           =   11120
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   570
            Left            =   120
            Top             =   4920
            Width           =   10980
         End
      End
   End
   Begin TrueOleDBGrid70.TDBGrid GridXLS 
      Height          =   2055
      Left            =   12840
      TabIndex        =   111
      Top             =   2280
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   3625
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "FormReminder1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private OnLoad As Boolean
Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String

Private objRem As New ClsRemainder
Private rsRem As ADODB.Recordset
'.:: RecordSet Temporary ::.
Private rsTem As ADODB.Recordset
Private rsGrid As New ADODB.Recordset
'.:: Item Temporary ::.
Private rsSvc As ADODB.Recordset
Private TglNext1 As Date
Private TglNext As Date
Private valDM As Integer
Private valSM As Integer
Private valCl As Integer
Private NoSPK As String
Private rsTempXLS As New ADODB.Recordset
Private rsTempMRS As New ADODB.Recordset
Private rsList As New ADODB.Recordset

Private Remainder As String
Private ft As New FormTemplateSMS

Private Harga As Double       '==> #RP
Private waktu As String       '==> #JMs
Private NewId As Double

Private rs1K As ADODB.Recordset
Private rs1 As ADODB.Recordset

Private rsStatusData As ADODB.Recordset

Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property

Private Sub SetPicture()
   cmdDM.Visible = False
   cmdSMS.Visible = False
   CmdCall.Visible = False
   Shape5.Visible = False
   Select Case cboType.Text
      Case "Direct Mail"
         ImgView.Picture = ImageList.ListImages(1).Picture
         lblKet.Caption = "Ket  :  H-14 (Next Service)" ' = " & DateAdd("d", 14, Me.DtRemain.value) & ")"
         TglNext = DateAdd("d", 14, Me.DtRemain.value)
         cmdDM.Visible = True
         cmdXlsMRS.Visible = False
         ShapeXLS.Left = 9360
         ShapeXLS.Width = 975
      Case "SMS"
         ImgView.Picture = ImageList.ListImages(3).Picture
         lblKet.Caption = "Ket  :  H-7 (Next Service)" ' -> " & DateAdd("d", 7, Me.DtRemain.value) & ")"
         TglNext = DateAdd("d", 7, Me.DtRemain.value)
         cmdSMS.Visible = True
         cmdXlsMRS.Visible = True
         ShapeXLS.Left = 8640
         ShapeXLS.Width = 1695
      Case "Call By Phone"
         ImgView.Picture = ImageList.ListImages(2).Picture
         lblKet.Caption = "Ket  :  H-3 (Next Service)" ' -> " & DateAdd("d", 3, Me.DtRemain.value) & ")"
         TglNext = DateAdd("d", 3, Me.DtRemain.value)
         CmdCall.Visible = True
         Shape5.Visible = True
         cmdXlsMRS.Visible = False
         ShapeXLS.Left = 9360
         ShapeXLS.Width = 975
      Case Else
         ImgView.Picture = Nothing
         lblKet.Caption = "Keterangan : - "
         TglNext = Null
         cmdXlsMRS.Visible = False
         ShapeXLS.Left = 9360
         ShapeXLS.Width = 975
   End Select
   lblJml.Caption = "0"
End Sub

Private Sub cboType_Change()
   Call cmdProcess_Click
End Sub

Private Sub cboType_Click()
   Call SetPicture
    Dim nC As Integer
    If Left(cboType.Text, 1) = "S" Then
        For nC = 0 To 30
            GridRemainder.Columns(nC).Visible = False
        Next nC
        For nC = 31 To 33
            GridRemainder.Columns(nC).Visible = True
        Next nC
        GridRemainder.Width = 4875
        cmdTemplate.Visible = True
        cmdSave.Visible = True
    'ElseIf Left(cboType.text, 1) = "D" Then
        
    Else
        For nC = 0 To 30
            GridRemainder.Columns(nC).Visible = True
        Next nC
        For nC = 31 To 33
            GridRemainder.Columns(nC).Visible = False
        Next nC
        GridRemainder.Width = 11115
        cmdTemplate.Visible = False
        cmdSave.Visible = False
    End If
   Call cmdProcess_Click
End Sub

Private Sub chCall_Click()
   Call ListRemainderValue
End Sub

Private Sub chDM_Click()
    Call ListRemainderValue
End Sub

Private Sub chSMS_Click()
   Call ListRemainderValue
End Sub

Private Sub cmdDM_Click()
    On Error GoTo err
    If rsList.RecordCount = 0 Then Exit Sub
    
'    If MsgBox("Via Email ?", vbYesNo + vbDefaultButton1, "Email") = vbYes Then
'        If MAPISession1.SessionID <> 0 Then MAPISession1.SignOff
'        MAPISession1.SignOn
'        MAPIMessages1.Compose
'        MAPIMessages1.RecipIndex = 1
'        MAPIMessages1.RecipDisplayName = txtEmail
'        MAPIMessages1.MsgSubject = "Pemberitahuan Reminder Service Pertama " & support.kopOrder(1)
'        MAPIMessages1.MsgNoteText = ""
'        MAPIMessages1.ResolveName
'        MAPIMessages1.Send
'    Else
        GoTo viaDM
'    End If
    Exit Sub

viaDM:
    Dim f As New FormTransparan
    With f
       .xRsRem = rsList
       .xIsForm = "DM"
       .xSBI_1K = True
       .show (vbModal)
    End With
    Call cmdProcess_Click
    Exit Sub
    
err:

End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOpen_Click()
   Dim f As New FormTransparan
   With f
      .xChassis = Trim(txtNoChassis.Text)
      .xIsForm = "CUST_EDIT"
'      .Top = 2250
'      .Left = Me.Left
      .isMe = True
      .show (vbModal)
   End With
   Call cmdProcess_Click
End Sub

Private Sub cmdSave_Click()
If GridRemainder.Columns(31).value <> -1 Then
   MsgBox "Silahkan pilih item customer dahulu !", vbCritical + vbOKOnly, "MRS"
   Exit Sub
End If
If txtPesan.Text = "" Then
   MsgBox "Silahkan isi pesan sms anda !" & Chr(13) & "Atau pilih template SMS terlebih dahulu", vbCritical + vbOKOnly, "MRS"
   Exit Sub
End If
If chGSM.value = 0 And chCDMA.value = 0 Then
   MsgBox "Silahkan pilih no telp terlebih dahulu", vbCritical + vbOKOnly, "MRS"
   Exit Sub
End If
    Dim nPesanSMS As String
    nPesanSMS = Trim(Me.txtPesan.Text) & IIf(fot.value = 1, " " & fot.Caption, "")
    
    objRem.RemoveItemSMS (Me.txtNoChassis.Text)
    With objRem
        If Me.chGSM.value = 1 Then
            .NomorRangka = Trim(Me.txtNoChassis.Text)
            .NoGSM = Trim(Replace(Me.chGSM.Caption, "-", ""))
            .NoCDMA = Trim("")
            .TglSendNow = Me.DTSend.value
            .PesanSMS = nPesanSMS
            .AddNewItemSMS_Temp
        End If
        If Me.chCDMA.value = 1 Then
            .NomorRangka = Trim(Me.txtNoChassis.Text)
            .NoGSM = Trim("")
            .NoCDMA = Trim(Replace(Me.chCDMA.Caption, "-", ""))
            .TglSendNow = Me.DTSend.value
            .PesanSMS = nPesanSMS
            .AddNewItemSMS_Temp
        End If
    End With
    Call ShowItemSave_Count
End Sub

Private Sub ShowItemSave_Count()
    Dim rsSave As ADODB.Recordset
    Set rsSave = objRem.GetDataSMS
    lblSave.Caption = "Save Item : " & rsSave.RecordCount
End Sub

Private Sub cmdTemplate_Click()
   ft.OnLoaded = True
   ft.show vbModal
   Remainder = ft.Remainder
   txtPesan.Text = Remainder
   Call PhoneChange
End Sub

Private Sub cmdTop_Click()
   Me.VScrollFUB.value = 1
End Sub

Private Sub CmdXls_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsList.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Reminder 1000 Km Kosong !"
5             Beep
6             Er.show (vbModal)
7             Exit Sub
8         End If
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".XLS"
          
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
16        sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
17        sSQL = sSQL & " values ('" & gSA & "','FormReminder1000',getdate(),'" & IIf(IsNull(DtRemain0), "1900-01-01", Format(DtRemain0, "yyyy-mm-dd")) & "','" & Format(DtRemain, "yyyy-mm-dd") & "','cmdXls - " & cboType.Text & "','" & rsList.RecordCount & "','" & onVal & "','" & ComptName & "') "
18        Call objO.InsertData(sSQL)
          
19        GridRemainder.Columns(7).Visible = False
20        GridRemainder.Columns(22).Visible = False
21         GridRemainder.ExportToFile onVal, True
22        GridRemainder.Columns(7).Visible = True
23        GridRemainder.Columns(22).Visible = True
24                Er.HeadCap = "Export to excel !"
25                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
26                Beep
27                Er.show (vbModal)
           'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
28         Exit Sub
Err_Export:
29          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
30          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub Form_Load()
   '' 20200512DEV
   If Log.disableExportMra = 1 Then
        If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
            cmdXlsMRS.Enabled = True
            CmdXls.Enabled = True
        Else
            cmdXlsMRS.Enabled = False
            CmdXls.Enabled = False
        End If
   End If
   If OnLoad = False Then Exit Sub
   Call ListAllCombo
   Call LoadForm
End Sub

Private Sub ListAllCombo()
   With Me
      .cboType.Clear
      .cboType.AddItem "Direct Mail"
      .cboType.AddItem "SMS"
      .cboType.AddItem "Call By Phone"
   End With
End Sub

Private Sub LoadForm()
   Me.DtRemain.value = Date
   Me.fot.Caption = Log.LogNas
   Me.fot.Caption = gPerush
   Call objRem.RemoveAllItemSMS
   lblSave.Caption = "Save Item : 0"
   Call NullRecord
End Sub

Private Sub cmdProcess_Click()

   If cboType.Text = "" Then
      'MsgBox "Silahkan Isi Type Remainder dan Kelipatan KM Terlebih Dahulu !", vbInformation + vbOKOnly
      Exit Sub
   End If
   Call ListDataRemainder
   Call ShowItemSave_Count
End Sub


Private Sub ListDataRemainder()
      Dim NextList As Boolean
      'On Error Resume Next
1        Call SetPicture
         'If TglNext < "01/01/1900" Then
2        If IsNull(DtRemain0) = True And IsNull(DtRemain) = True Then
3           Call errTglNextService
4           Exit Sub
5        End If
6        If DtRemain0 > "01/01/1900" And IsNull(DtRemain) = True Then
7           Call errTglNextService
8           Exit Sub
9        End If
10       If DtRemain0 > DtRemain Then
11          Call errTglNextService
12          Exit Sub
13       End If
         
14       Set GridRemainder.DataSource = Nothing
          Dim nRem As Integer
15        If IsNull(DtRemain0) = True And DtRemain.value > "01/01/1900" Then
16            nRem = 1
17            Select Case cboType.Text
                 Case "SMS"
18                  lblKet.Caption = "Ket  :  H-7 (Next) 01/01/1900 - " & DateAdd("d", 7, Me.DtRemain.value)
19               Case "Call By Phone"
20                  lblKet.Caption = "Ket  :  H-3 (Next) 01/01/1900 - " & DateAdd("d", 3, Me.DtRemain.value)
21               Case "Direct Mail"
22                  lblKet.Caption = "Ket  :  H-14 (Next) 01/01/1900 - " & DateAdd("d", 14, Me.DtRemain.value)
23            End Select
24        ElseIf DtRemain0.value > "01/01/1900" And DtRemain.value > "01/01/1900" Then
25            nRem = 0
26            Select Case cboType.Text
                 Case "SMS"
27                  lblKet.Caption = "Ket  :  H-7 (Next) " & DateAdd("d", 7, Me.DtRemain0.value) & " - " & DateAdd("d", 7, Me.DtRemain.value)
28               Case "Call By Phone"
29                  lblKet.Caption = "Ket  :  H-3 (Next) " & DateAdd("d", 3, Me.DtRemain0.value) & " - " & DateAdd("d", 3, Me.DtRemain.value)
30               Case "Direct Mail"
31                  lblKet.Caption = "Ket  :  H-14 (Next) " & DateAdd("d", 14, Me.DtRemain0.value) & " - " & DateAdd("d", 14, Me.DtRemain.value)
32            End Select
33        End If
          
          Dim tNext0 As Date
          Dim tNext1 As Date
          Dim tRem0 As Date
          Dim tRem1 As Date
34            tRem0 = "01/01/1900"
35            tRem1 = "01/01/1900"
              'tRem1 = Me.DtRemain.value
36            Select Case cboType.Text
                 Case "SMS"
37                    If IsNull(DtRemain0) = True Then
38                        tNext0 = "01/01/1900"
                          'tRem0 = "01/01/1900"
39                    Else
40                        tNext0 = DateAdd("d", 7, Me.DtRemain0.value)
                          'tRem0 = Me.DtRemain0.value
41                    End If
42                    tNext1 = DateAdd("d", 7, Me.DtRemain.value)
43               Case "Call By Phone"
44                    If IsNull(DtRemain0) = True Then
45                        tNext0 = "01/01/1900"
                          'tRem0 = "01/01/1900"
46                    Else
47                        tNext0 = DateAdd("d", 3, Me.DtRemain0.value)
48                        tRem0 = Me.DtRemain0.value
49                    End If
50                    tNext1 = DateAdd("d", 3, Me.DtRemain.value)
51                    tRem1 = Me.DtRemain.value
52               Case "Direct Mail"
53                    If IsNull(DtRemain0) = True Then
54                        tNext0 = "01/01/1900"
                          'tRem0 = "01/01/1900"
55                    Else
56                        tNext0 = DateAdd("d", 14, Me.DtRemain0.value)
                          'tRem0 = Me.DtRemain0.value
57                    End If
58                    tNext1 = DateAdd("d", 14, Me.DtRemain.value)
                      'tRem1 = Me.DtRemain.value
59            End Select
          
60       Set rsRem = objRem.FetchDataRemainder1000(tNext0, tNext1, nRem, tRem0, tRem1)
61       If rsRem.RecordCount = 0 Then
62             Call NullRecord
         
63       Else
                  '' 20200512DEV
64                If Log.disableExportMra = 1 Then
65                    If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
66                        cmdXlsMRS.Enabled = True
67                        CmdXls.Enabled = True
68                    Else
69                        cmdXlsMRS.Enabled = False
70                        CmdXls.Enabled = False
71                    End If
72                Else
73                    cmdXlsMRS.Enabled = True
74                    CmdXls.Enabled = True
75                End If
      '            CmdXls.Enabled = True
      '            cmdXlsMRS.Enabled = True
76                cmdTemplate.Enabled = True
77                chGSM.Enabled = True
78                chCDMA.Enabled = True
79                cmdSave.Enabled = True
80                cmdSMS.Enabled = True
81                cmdOpen.Enabled = True
                          
82                CreateMainList True
83                rsRem.MoveFirst
                  Dim isValid As Boolean
                  Dim noPTSN, noHP1, noHP2 As String
                  
84                Do Until rsRem.EOF
85                    NextList = False
86                    isValid = True
                      Dim xNoSPK As String
87                    xNoSPK = rsRem.Fields(0).value
88                    If onS.value = 1 Then
89                        If AddToListRemainderSMS(xNoSPK) = True Then NextList = True
90                    End If
91                    If onC.value = 1 Then
92                        If AddToListRemainderCall(xNoSPK) = True Then NextList = True
93                    End If
94                    If onD.value = 1 Then
95                        If AddToListRemainderDM(xNoSPK) = True Then NextList = True
96                    End If
97                    If NextList = False Then
98                        sSQL = "select a.NoSPK, a.Dealer, a.NoRangka, UserId, TglProses, a.tipetelp, StatusNoTelp, UserUpdateNIK, TglUpdateNik, NoTelpRevisi"
99                        sSQL = sSQL & " from HistLogInvalidOnCall a with (nolock)"
100                       sSQL = sSQL & " left join LogInvalidTelp b with (nolock) on a.NoSPK = b.NoSPK and a.NoRangka = b.NoRangka and a.Dealer = b.Dealer and a.tipetelp = b.tipetelp and a.TipeCustomer = b.TipeCustomer"
101                       sSQL = sSQL & " where a.TipeCustomer = 'Surat Menyurat' and a.NoSPK = '" & NulSID(rsRem!spk_no) & "' and a.NoRangka = '" & NulSID(rsRem!noChassis) & "'"
                          
102                       sSQL = "select c.spk_no, LEFT(c.spk_no, 4) Dealer, c.fixcr1, D.NoRangka, UserId, TglProses, a.TipeCustomer, a.tipetelp, b.StatusNoTelp, UserUpdateNIK, TglUpdateNik, NoTelpRevisi"
103                       sSQL = sSQL & " from AFI c with (nolock)"
104                       sSQL = sSQL & " inner join SPK D with (nolock) on c.SPK_No = d.SPK_No"
105                       sSQL = sSQL & " left join HistLogInvalidOnCall a with (nolock) on c.SPK_No = a.NoSPK and a.TipeCustomer = 'Surat Menyurat'"
106                       sSQL = sSQL & " left join LogInvalidTelp b with (nolock) on a.NoSPK = b.NoSPK and a.NoRangka = b.NoRangka and a.Dealer = b.Dealer and a.tipetelp = b.tipetelp and a.TipeCustomer = b.TipeCustomer"
107                       sSQL = sSQL & " where c.spk_no = '" & NulSID(rsRem!spk_no) & "' and d.NoRangka = '" & NulSID(rsRem!noChassis) & "'"
108                       Set rsStatusData = objO.FetchDataSPK(sSQL)
                          
109                       noPTSN = NulSID(rsRem!Telpon)
110                       noHP1 = NulSID(rsRem!GSM)
111                       noHP2 = NulSID(rsRem!CDMA)
112                       isValid = False
113                       If rsStatusData.RecordCount > 0 Then
114                           rsStatusData.MoveFirst
115                           Do While Not rsStatusData.EOF
116                               If (NulSID(rsStatusData!StatusNoTelp) = False Or NulSID(rsStatusData!StatusNoTelp) = "") And NulSID(rsStatusData!fixcr1) = True Then
117                                   isValid = True
118                                   If NulSID(rsStatusData!tipetelp) = "PSTN" Then
119                                       noPTSN = NulSID(rsStatusData!NoTelpRevisi)
120                                   End If
121                                   If NulSID(rsStatusData!tipetelp) = "HP1" Then
122                                       noHP1 = NulSID(rsStatusData!NoTelpRevisi)
123                                   End If
124                                   If NulSID(rsStatusData!tipetelp) = "HP2" Then
125                                       noHP2 = NulSID(rsStatusData!NoTelpRevisi)
126                                   End If
127                               Else
128                                   isValid = False
129                               End If
130                               rsStatusData.MoveNext
131                           Loop
132                       Else
133                           isValid = False
134                           noPTSN = NulSID(rsRem!Telpon)
135                           noHP1 = NulSID(rsRem!GSM)
136                           noHP2 = NulSID(rsRem!CDMA)
137                       End If
138                       rsList.AddNew
139                       rsList!NoSPK = rsRem!spk_no
140                       rsList!Tgl_DO = CStr(Format(rsRem!TglDo, "DD/MM/YYYY"))
141                       rsList!Tgl_CR1 = CStr(Format(rsRem!TglCR1, "DD/MM/YYYY"))
142                       rsList!Km_CR1 = rsRem!KmCR1
143                       rsList!Tgl_CR2 = CStr(Format(rsRem!TglCR2, "DD/MM/YYYY"))
144                       rsList!Km_CR2 = NulA(rsRem!KmCR2)
145                       rsList!Nama_Cust = rsRem!UseName
146                       rsList!Alamat_Cust = rsRem!UseAddr
147                       rsList!Kota_Cust = rsRem!UseCity
148                       rsList!KodePos_Cust = rsRem!UsePos
149                       rsList!NO_POLISI = Nul(rsRem!NoPolisi)
150                       rsList!No_Engine = rsRem!NoEngine
151                       rsList!noChassis = rsRem!noChassis
152                       rsList!Telpon = noPTSN 'Nul(rsRem!Telpon)
153                       rsList!GSM = noHP1 'Nul(rsRem!GSM)
154                       rsList!CDMA = noHP2 'Nul(rsRem!CDMA)
155                       rsList!Kode_Type = rsRem!Kodetype
156                       rsList!Nama_Type = rsRem!Namatype
157                       rsList!Kode_Warna = rsRem!KodeWarna
158                       rsList!Nama_Warna = rsRem!Namawarna
159                       rsList!WIRANIAGA = rsRem!WIRANIAGA
160                       rsList!Nama_stnk = rsRem!namastnk
161                       rsList!Alamat_STNK = rsRem!AlamatSTNK
162                       rsList!Kota_STNK = rsRem!Kotastnk
163                       rsList!KodePos_STNK = rsRem!KodePosSTNK
164                       rsList!Tempat_Lahir = rsRem!TempatLahir
165                       rsList!Tanggal_Lahir = IIf(IsNull(rsRem!TanggalLahir) = True, "", Format(rsRem!TanggalLahir, "DD/MM/YYYY"))
166                       rsList!Agama = rsRem!Agama
167                       rsList!pekerjaan = rsRem!pekerjaan
168                       rsList!Hoby = rsRem!Hoby
169                       rsList!Tgl_Service = CStr(Format(rsRem!Tgl1, "DD/MM/YYYY"))
170                       rsList!Select = 0
171                       rsList!Nomor_Rangka = rsRem!NO_CHASSIS
172                       rsList!Kode_Lgn = rsRem!Kode_Lgn
173                       rsList!Email = ""
174                       rsList!statusdata = IIf(isValid, "Valid", "Proses Revisi")
175                   End If
176                   rsRem.MoveNext
177               Loop
178               rsRem.MoveFirst
179               lblJml.Caption = rsList.RecordCount
180               Call BindToControl(rsList)
181               Call objRem.RemoveAllItemSMS
182      End If
         'Set Me.GridRemainder.DataSource = rsRem
         'lblJml.Caption = rsRem.RecordCount
         'Call BindToControl(rsRem)
End Sub

Private Sub NullRecord()

         NoSPK = ""
         valSM = 0
         valCl = 0
         valDM = 0
         chSMS.value = valSM
         chCall.value = valCl
         chDM.value = valDM
         nDM.Caption = ".  .  .  .  .  .  .  .  ."
         nSMS.Caption = ".  .  .  .  .  .  .  .  ."
         nCall.Caption = ".  .  .  .  .  .  .  .  ."
    With Me
       .CmdXls.Enabled = False
       .cmdXlsMRS.Enabled = False
       .cmdTemplate.Enabled = False
       .chGSM.Enabled = False
       .chCDMA.Enabled = False
       .cmdSave.Enabled = False
       .cmdSMS.Enabled = False
       .DTSend.Text = "__/__/____"
       .cmdOpen.Enabled = False

        .DtDO.Text = "__/__/____"
        .DtDEC.Text = "__/__/____"
        .DTCR1.Text = "__/__/____"
        .txtKMCR1.Text = ""
        .DTCR2.Text = "__/__/____"
        .txtKMCR2.Text = ""
        .txtNama.Text = ""
        .txtAlamat.Text = ""
        .txtKota.Text = ""
        .txtPos.Text = ""
        .txtNoPol.Text = ""
'        .txtNoEngine.Text = ""
        .txtEmailAdr.Text = ""
        .txtNoChassis.Text = ""
        .txtTelp.Text = ""
        .txtGSM.Text = ""
        .txtCDMA.Text = ""
        .txtKodeType.Text = ""
        .txtNamaType.Text = ""
        .txtKodeWarna.Text = ""
        .txtNamaWarna.Text = ""
        .txtwiraniaga.Text = ""
        .txtNamaSTNK.Text = ""
        .txtAlamatSTNK.Text = ""
        .txtKotaSTNK.Text = ""
        .txtKodePosSTNK.Text = ""
        .txtTempatLahir.Text = ""
        .DtLahir.Text = "__/__/____"
        .txtAgama.Text = ""
        .txtPekerjaan.Text = ""
        .txtHoby.Text = ""
        .txtRem.Text = ""
    End With
End Sub

'Private Sub CreateTblMRS(Optional Edit As Boolean = False)
'    On Error GoTo Vb_Error
'    If rsTempMRS.State Then rsTempMRS.Close
'    With rsTempMRS
'        .Fields.Append "Select", adBoolean
'        .Fields.Append "NoRangka", adVarChar, 30
'        .Fields.Append "KodeLgn", adVarChar, 15
'        .Open
'        Set GridMRS.DataSource = rsTempMRS
'    End With
'    Exit Sub
'Vb_Error:
'    Err.Raise Err.Number, Err.Source, Err.Description
'End Sub

Private Sub errTglNextService()
        Dim Er As New ErrInfo
        Er.HeadCap = "Tanggal Reminder !"
        Er.InfoCap = "Tanggal Remainder tidak valid !"
        Beep
        Er.show (vbModal)
      'MsgBox "Tanggal Remainder tidak valid !", vbExclamation, "Remainder"
End Sub

Private Sub n1()
With Me
         .txtKodePosSTNK.Text = ""
         .txtTempatLahir.Text = ""
         .txtGSM.Text = ""
         .txtCDMA.Text = ""
         .txtKodeType.Text = ""
         .txtNamaType.Text = ""
         .txtKodeWarna.Text = ""
         .txtNamaWarna.Text = ""
         .txtwiraniaga.Text = ""
         .DtLahir.Text = "__/__/____"
         .txtAgama.Text = ""
         .txtPekerjaan.Text = ""
         .txtHoby.Text = ""
         NoSPK = ""
         .txtKotaSTNK.Text = ""
         valSM = 0
         valCl = 0
         .chSMS.value = valSM
         .chCall.value = valCl
         .nSMS.Caption = ".  .  .  .  .  .  .  .  ."
         .nCall.Caption = ".  .  .  .  .  .  .  .  ."
         .txtTelp.Text = ""
         .txtNamaSTNK.Text = ""
         .txtAlamatSTNK.Text = ""
         .txtAlamat.Text = ""
         .txtKota.Text = ""
         .txtPos.Text = ""
         .txtNoPol.Text = ""
'         .txtNoEngine.Text = ""
         .txtEmailAdr.Text = ""
         .txtNoChassis.Text = ""
End With
End Sub

Private Sub BindToControl(rs As ADODB.Recordset)
On Error Resume Next
   With Me
    NoSPK = rs.Fields(0).value
    Call ListRemainder
    
    Set .DtDO.DataSource = rs
        DtDO.DataField = IIf(IsNull(rs.Fields(1).Name) = True, "", rs.Fields(1).Name)
    Set .DtDEC.DataSource = rs
        DtDEC.DataField = IIf(IsNull(rs.Fields(1).Name) = True, "", rs.Fields(1).Name)
    Set .DTCR1.DataSource = rs
        DTCR1.DataField = IIf(IsNull(rs.Fields(2).Name) = True, "", rs.Fields(2).Name)
    Set .txtKMCR1.DataSource = rs
        txtKMCR1.DataField = IIf(IsNull(rs.Fields(3).Name) = True, "", rs.Fields(3).Name)
    Set .DTCR2.DataSource = rs
        DTCR2.DataField = IIf(IsNull(rs.Fields(4).Name) = True, "", rs.Fields(4).Name)
    Set .txtKMCR2.DataSource = rs
        txtKMCR2.DataField = IIf(IsNull(rs.Fields(5).Name) = True, "", rs.Fields(5).Name)
    Set .txtNama.DataSource = rs
        txtNama.DataField = IIf(IsNull(rs.Fields(6).Name) = True, "", rs.Fields(6).Name)
    Set .txtAlamat.DataSource = rs
        txtAlamat.DataField = IIf(IsNull(rs.Fields(7).Name) = True, "", rs.Fields(7).Name)
    Set .txtKota.DataSource = rs
        txtKota.DataField = IIf(IsNull(rs.Fields(8).Name) = True, "", rs.Fields(8).Name)
    Set .txtPos.DataSource = rs
        txtPos.DataField = IIf(IsNull(rs.Fields(9).Name) = True, "", rs.Fields(9).Name)
    Set .txtNoPol.DataSource = rs
        txtNoPol.DataField = IIf(IsNull(rs.Fields(10).Name) = True, "", rs.Fields(10).Name)
'    Set .txtNoEngine.DataSource = rs
'        txtNoEngine.DataField = IIf(IsNull(rs.Fields(11).Name) = True, "", rs.Fields(11).Name)
'    Set .txtEmailAdr.DataSource = rs
'        txtEmailAdr.DataField
    Set .txtNoChassis.DataSource = rs
        txtNoChassis.DataField = IIf(IsNull(rs.Fields(12).Name) = True, "", rs.Fields(12).Name)
    Set .txtTelp.DataSource = rs
        txtTelp.DataField = IIf(IsNull(rs.Fields(13).Name) = True, "", rs.Fields(13).Name)
    Set .txtGSM.DataSource = rs
        txtGSM.DataField = IIf(IsNull(rs.Fields(14).Name) = True, "", rs.Fields(14).Name)
    Set .txtCDMA.DataSource = rs
        txtCDMA.DataField = IIf(IsNull(rs.Fields(15).Name) = True, "", rs.Fields(15).Name)
    Set .txtKodeType.DataSource = rs
        txtKodeType.DataField = IIf(IsNull(rs.Fields(16).Name) = True, "", rs.Fields(16).Name)
    Set .txtNamaType.DataSource = rs
        txtNamaType.DataField = IIf(IsNull(rs.Fields(17).Name) = True, "", rs.Fields(17).Name)
    Set .txtKodeWarna.DataSource = rs
        txtKodeWarna.DataField = IIf(IsNull(rs.Fields(18).Name) = True, "", rs.Fields(18).Name)
    Set .txtNamaWarna.DataSource = rs
        txtNamaWarna.DataField = IIf(IsNull(rs.Fields(19).Name) = True, "", rs.Fields(19).Name)
    Set .txtwiraniaga.DataSource = rs
        txtwiraniaga.DataField = IIf(IsNull(rs.Fields(20).Name) = True, "", rs.Fields(20).Name)
    Set .txtNamaSTNK.DataSource = rs
        txtNamaSTNK.DataField = IIf(IsNull(rs.Fields(21).Name) = True, "", rs.Fields(21).Name)
    Set .txtAlamatSTNK.DataSource = rs
        txtAlamatSTNK.DataField = IIf(IsNull(rs.Fields(22).Name) = True, "", rs.Fields(22).Name)
    Set .txtKotaSTNK.DataSource = rs
        txtKotaSTNK.DataField = IIf(IsNull(rs.Fields(23).Name) = True, "", rs.Fields(23).Name)
    Set .txtKodePosSTNK.DataSource = rs
        txtKodePosSTNK.DataField = IIf(IsNull(rs.Fields(24).Name) = True, "", rs.Fields(24).Name)
    Set .txtTempatLahir.DataSource = rs
        txtTempatLahir.DataField = IIf(IsNull(rs.Fields(25).Name) = True, "", rs.Fields(25).Name)
    Set .DtLahir.DataSource = rs
        DtLahir.DataField = IIf(IsNull(rs.Fields(26).Name) = True, "", rs.Fields(26).Name)
    Set .txtAgama.DataSource = rs
        txtAgama.DataField = IIf(IsNull(rs.Fields(27).Name) = True, "", rs.Fields(27).Name)
    Set .txtPekerjaan.DataSource = rs
        txtPekerjaan.DataField = IIf(IsNull(rs.Fields(28).Name) = True, "", rs.Fields(28).Name)
    Set .txtHoby.DataSource = rs
        txtHoby.DataField = IIf(IsNull(rs.Fields(29).Name) = True, "", rs.Fields(29).Name)
    Set .txtRem.DataSource = rs
        txtRem.DataField = IIf(IsNull(rs.Fields(30).Name) = True, "", rs.Fields(30).Name)
   End With
   Call PhoneChange
End Sub

Private Sub ListRemainder()
On Error Resume Next
   Dim rsListItem As ADODB.Recordset
   Set rsListItem = objRem.FetchListRemainder1000(Trim(NoSPK), 0)
   With Me
      valSM = 0
      valCl = 0
      valDM = 0
      .chSMS.value = valSM
      .chCall.value = valCl
      .chDM.value = valDM
      .nDM.Caption = ".  .  .  .  .  .  .  .  ."
      .nSMS.Caption = ".  .  .  .  .  .  .  .  ."
      .nCall.Caption = ".  .  .  .  .  .  .  .  ."
      If rsListItem.RecordCount <> 0 Then
         rsListItem.MoveFirst
         Do Until rsListItem.EOF
            If rsListItem.Fields(0).value = "DM" Then
               valDM = 1
               .chDM.value = valDM
               .nDM.Caption = Trim(rsListItem.Fields(1).value)
            ElseIf rsListItem.Fields(0).value = "1K" Then
               valSM = 1
               .chSMS.value = valSM
               .nSMS.Caption = Trim(rsListItem.Fields(1).value)
            End If
            rsListItem.MoveNext
         Loop
      End If
   End With
   Dim rsCall As ADODB.Recordset
   Set rsCall = objRem.FetchListRemainder1000(NoSPK, 2)
   If rsCall.RecordCount <> 0 Then
      valCl = 1
      chCall.value = valCl
      nCall.Caption = Trim(rsCall.Fields(0).value) & "  - Call : " & rsCall.RecordCount
   End If
End Sub

Private Function AddToListRemainderDM(tNoSPK As String) As Boolean
On Error Resume Next
   AddToListRemainderDM = False
   Dim rsListItem As ADODB.Recordset
   Set rsListItem = objRem.FetchListRemainder1000(Trim(tNoSPK), 3)
      If rsListItem.RecordCount <> 0 Then
         rsListItem.MoveFirst
         Do Until rsListItem.EOF
            If rsListItem.Fields(0).value = "DM" Then
                AddToListRemainderDM = True
            End If
            rsListItem.MoveNext
         Loop
      End If
End Function

Private Function AddToListRemainderSMS(tNoSPK As String) As Boolean
On Error Resume Next
   AddToListRemainderSMS = False
   Dim rsListItem As ADODB.Recordset
   Set rsListItem = objRem.FetchListRemainder1000(Trim(tNoSPK), 1)
      If rsListItem.RecordCount <> 0 Then
         rsListItem.MoveFirst
         Do Until rsListItem.EOF
            If rsListItem.Fields(0).value = "1K" Then
                AddToListRemainderSMS = True
            End If
            rsListItem.MoveNext
         Loop
      End If
End Function

Private Function AddToListRemainderCall(tNoSPK As String) As Boolean
    AddToListRemainderCall = False
    Dim rsCall As ADODB.Recordset
    Set rsCall = objRem.FetchListRemainder1000(tNoSPK, 2)
    If rsCall.RecordCount <> 0 Then
        AddToListRemainderCall = True
    End If
End Function

Private Sub ListRemainderValue()
   chDM.value = valDM
   chSMS.value = valSM
   chCall.value = valCl
End Sub

Private Sub GridRemainder_FilterChange()
On Error GoTo err
   Set cols = GridRemainder.Columns
   Dim C As Integer

   C = GridRemainder.col
   GridRemainder.HoldFields
   rsList.filter = getFilter()
   If rsList.EOF = True Then
      For Each col In GridRemainder.Columns
         col.FilterText = ""
      Next col
      rsList.filter = ""
   End If
   GridRemainder.col = C
   GridRemainder.EditActive = True
   lblJml.Caption = rsList.RecordCount
   Exit Sub
err:
   For Each col In cols
      col.FilterText = ""
   Next col
   Resume Next
End Sub

Private Function getFilter() As String
   On Error GoTo err

On Error GoTo err
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

   Exit Function
err:
   Resume Next
End Function

Private Sub GridRemainder_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 14 Then rsList.Sort = GridRemainder.Columns(ColIndex).DataField & " " & Urut
4        GridRemainder.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        lblJml.Caption = rsList.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
End Sub

Private Sub GridRemainder_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo rsKosong
    sService.Visible = False
   If rsList.RecordCount = 0 Then Exit Sub
   Call BindToControl(rsList)
    Set rs1K = objRem.FetchItem_1K(rsList.Fields(12).value)
    If rs1K.RecordCount = 0 Then
        on1K.value = 0
        on1K.BackColor = &HFFFFFF
    Else
        rs1K.MoveFirst
        Do Until rs1K.EOF
            If rs1K.Fields(0).value = "R001" Then
                on1K.value = 1
                on1K.BackColor = vbRed
                Exit Do
            Else
                on1K.value = 0
                on1K.BackColor = &HFFFFFF
            End If
            rs1K.MoveNext
        Loop
    End If
rsKosong:

End Sub

Private Sub Label52_Click()
    If chCDMA.Enabled = True Then
        chCDMA.value = IIf(chCDMA.value = 0, 1, 0)
    End If
End Sub

Private Sub Label55_Click()
    If chGSM.Enabled = True Then
        chGSM.value = IIf(chGSM.value = 0, 1, 0)
    End If
End Sub

Private Sub PhoneChange()
On Error Resume Next
    Me.chGSM.Caption = txtGSM.Text
    If txtGSM.Text = "" Or Len(txtGSM.Text) < 6 Then
        Me.chGSM.value = 0
        Me.chGSM.Enabled = False
    Else
        Me.chGSM.Enabled = True
    End If
    Me.chCDMA.Caption = txtCDMA.Text
    If txtCDMA.Text = "" Or Len(txtCDMA.Text) < 6 Then
        Me.chCDMA.value = 0
        Me.chCDMA.Enabled = False
    Else
        Me.chCDMA.Enabled = True
    End If
    
   'Dim hrg As String
   'hrg = "Jasa : Rp." & txtHargaJasa.text & ",-  Part/Oli : Rp." & txtHargaPart.text & ",-"
   'Harga = txtHargaJasa + txtHargaPart
   'waktu = cboJam.text & ":" & cboMenit.text
   Dim MyText As String
   Dim IsText As String
   IsText = Remainder
   MyText = Replace(IsText, "#NM", txtNama.Text)
   MyText = Replace(MyText, "#AL", txtAlamat.Text)
   MyText = Replace(MyText, "#NP", txtNoPol.Text)
   MyText = Replace(MyText, "#TG", txtRem.Text)
   MyText = Replace(MyText, "#KM", "1.000")
   'MyText = Replace(MyText, "#RP", hrg)
   'MyText = Replace(MyText, "#JM", waktu)
   Dim i As Integer
   For i = 0 To 5
      MyText = Replace(MyText, "  ", " ")
   Next i
   txtPesan.Text = MyText
   
   fot.value = 0
   chGSM.value = 0
   chCDMA.value = 0
   DTSend.value = isnow
End Sub


Private Sub cmd1K_Click()
    If on1K.value = 1 Then
        If sService.Visible = False Then
            Set rs1 = objRem.GetData1K(rsList.Fields(12).value)
            If rs1.RecordCount <> 0 Then
               With Me
                Set .txtNoInv(0).DataSource = rs1
                    txtNoInv(0).DataField = rs1.Fields(0).Name
                Set .txtNoInv(1).DataSource = rs1
                    txtNoInv(1).DataField = rs1.Fields(0).Name
                Set .txtNoWO(0).DataSource = rs1
                    txtNoWO(0).DataField = rs1.Fields(1).Name
                Set .txtNoWO(1).DataSource = rs1
                    txtNoWO(1).DataField = rs1.Fields(1).Name
                Set .nTglSvc.DataSource = rs1
                    nTglSvc.DataField = rs1.Fields(2).Name
                Set .nkmSvc.DataSource = rs1
                    nkmSvc.DataField = rs1.Fields(3).Name
                Set .nNama.DataSource = rs1
                    nNama.DataField = rs1.Fields(4).Name
                Set .nAlamat.DataSource = rs1
                    nAlamat.DataField = rs1.Fields(5).Name
                Set .nNoPol.DataSource = rs1
                    nNoPol.DataField = rs1.Fields(6).Name
                Set .nEngine.DataSource = rs1
                    nEngine.DataField = rs1.Fields(7).Name
                Set .nChassis.DataSource = rs1
                    nChassis.DataField = rs1.Fields(8).Name
                Set .nPSTN.DataSource = rs1
                    nPSTN.DataField = rs1.Fields(9).Name
                Set .nGSM.DataSource = rs1
                    nGSM.DataField = rs1.Fields(10).Name
                Set .nCDMA.DataSource = rs1
                    nCDMA.DataField = rs1.Fields(11).Name
                Set .txtEmail.DataSource = rs1
                    txtEmail.DataField = rs1.Fields(12).Name
                    
                Dim nR As Integer
                For nR = 0 To 7
                    Set .tRep(nR).DataSource = rs1
                        tRep(nR).DataField = rs1.Fields(nR + 12).Name
                Next nR
                For nR = 0 To 7
                    Set .tRte(nR).DataSource = rs1
                        tRte(nR).DataField = rs1.Fields(nR + 20).Name
                Next nR
               End With
            End If
            sService.Visible = True
        Else
            sService.Visible = False
        End If
    End If
End Sub

Private Sub txtPesan_Change()
    lbLen.Caption = "Char : " & Len(txtPesan.Text)
End Sub

Private Sub VScrollFUB_Change()
    Me.SSFUB.Top = -(VScrollFUB.value + 200)
End Sub
Private Sub VScrollFUB_Scroll()
    Me.SSFUB.Top = -(VScrollFUB.value + 200)
End Sub

Private Sub cmdSMS_Click()
   If rsList.RecordCount = 0 Then Exit Sub
'   If chSMS.value = 1 Then
'            Dim Er As New ErrInfo
'            Er.HeadCap = "Reminder - SMS !"
'            Er.InfoCap = "Reminder SMS sudah pernah di proses !"
'
'            Beep
'            Er.Show (vbModal)
'      'MsgBox "Remainder - SMS sudah pernah di proses !", vbExclamation, "SMS Remainder"
'      Exit Sub
'   End If
'''   Dim f As New FormTransparan
'''   With f
'''      .xRsRem = rsList
'''      .Left = Me.Left
'''      .Top = 2250
'''      .xIsForm = "SMS1000"
'''      .Show (vbModal)
'''   End With
'''   Call cmdProcess_Click

Dim ListSMS As String
Call TambahSMS
    Dim rsInTem As ADODB.Recordset
    Set rsInTem = objRem.GetDataSMS

    If rsInTem.RecordCount = 0 Then
        MsgBox "Data SMS Masih Kosong !", vbInformation + vbOKOnly, "SMS"
        Exit Sub
    End If
    Dim objPul As New ClsRemainder
    rsInTem.MoveFirst
    Do Until rsInTem.EOF
        With objRem
            .NoSMS = NewId
            ListSMS = ListSMS & rsInTem.Fields(0).value
            .pesan = Trim(rsInTem.Fields(4).value)
            If rsInTem.Fields(1).value <> "" Then
                .Tujuan = Trim(rsInTem.Fields(1).value)
                ListSMS = ListSMS & " -  No.HP : " & rsInTem.Fields(1).value
                Call objRem.AddNew_LogRemainder("1K", rsInTem.Fields(0).value, objRem.MyNoSPK(rsInTem.Fields(0).value), TglNext, 1000, rsInTem.Fields(4).value, rsInTem.Fields(1).value, 0, 0, 0, 0, 0)
            End If
            If rsInTem.Fields(2).value <> "" Then
                .Tujuan = Trim(rsInTem.Fields(2).value)
                ListSMS = ListSMS & " -  No.HP : " & rsInTem.Fields(2).value
                Call objRem.AddNew_LogRemainder("1K", rsInTem.Fields(0).value, objRem.MyNoSPK(rsInTem.Fields(0).value), TglNext, 1000, rsInTem.Fields(4).value, rsInTem.Fields(2).value, 0, 0, 0, 0, 0)
            End If
            .waktu = isnow
            .AddNewItemSMS
        End With

        Call objPul.UpdatePulsa
        rsInTem.MoveNext
        NewId = Val(NewId) + 1
        ListSMS = ListSMS & Chr(13)
    Loop
    MsgBox "Data List SMS : " & Chr(13) & "----------------------------------------------------------------------" & Chr(13) & ListSMS & Chr(13) & "----------------------------------------------------------------------"
    MsgBox "Data SMS telah dikirim !"

    objRem.RemoveAllItemSMS
    Call cmdProcess_Click
End Sub

Private Sub TambahSMS()
    Dim rsSMS As ADODB.Recordset
    Set rsSMS = objRem.GetDataOutbox
    
    If rsSMS.RecordCount = 0 Then
        NewId = "1"
    Else
        rsSMS.MoveLast
        NewId = Val(rsSMS.Fields(0).value) + 1
    End If
End Sub

Private Sub cmdCall_Click()
   If rsList.RecordCount = 0 Then Exit Sub
   Dim f As New FormTransparan
   With f
      .xRsRem = rsList
      .xRsAll = rsRem
      .xIsForm = "CALL1000"
      .xIsTgl = DtRemain.value
      .xNoPol = txtNoPol.Text
'      .Top = 2250
'      .Left = Me.Left
      .show (vbModal)
   End With
   Call cmdProcess_Click
End Sub

Private Sub cmdXlsMRS_Click()
1     On Error GoTo Keluar
          Dim i As Integer
          Dim urutan As String
2         urutan = 1
3         Call CreateTblXLS(True)
4         rsList.MoveFirst
5         Do Until rsList.EOF
6                     rsTempXLS.AddNew
7                     rsTempXLS!nomor = urutan
8                     rsTempXLS!Nama_Customer = rsList!Nama_Cust
9                     rsTempXLS!Tgl_DO = Trim(Format(rsList!Tgl_DO, "DD/MM/YYYY"))
10                    rsTempXLS!Tgl_CR1 = Trim(Format(rsList!Tgl_CR1, "DD/MM/YYYY"))
11                    rsTempXLS!Km_CR1 = Trim(rsList!Km_CR1)
12                    rsTempXLS!Tgl_CR2 = Trim(Format(rsList!Tgl_CR2, "DD/MM/YYYY"))
13                    rsTempXLS!Km_CR2 = Trim(rsList!Km_CR2)
14                    rsTempXLS!NO_POLISI = Nul(rsList!NO_POLISI)
15                    rsTempXLS!No_Rangka = rsList!noChassis
16                    rsTempXLS!Telp_Rumah = rsList!Telpon
17                    rsTempXLS!Telp_GSM = rsList!GSM
18                    rsTempXLS!Telp_CDMA = rsList!CDMA
19                    rsTempXLS!Tgl_ETA = Trim(Format(rsList!Tgl_Service, "DD/MM/YYYY"))
20                    rsTempXLS!Tgl_SMS = Format(isnow, "DD/MM/YYYY")
21                    rsTempXLS!Tgl_Servis = ""
22                    rsTempXLS!Tgl_Follow_Up = ""
23                    rsTempXLS!Komentar = ""
24                    rsTempXLS!keterangan = ""
25                    rsTempXLS!Nama_Sales = rsList!WIRANIAGA
                  
26                    rsTempXLS.Update
                      
                      
27                urutan = urutan + 1
28                rsList.MoveNext
29        Loop
          Dim Er As New ErrInfo
30        If rsTempXLS.RecordCount = 0 Then
31            Er.HeadCap = "Export To Excel !"
32            Er.InfoCap = "List Reminder 1000 Km Kosong !"
33            Beep
34            Er.show (vbModal)
35            Exit Sub
36        End If
37    On Error GoTo Err_Export
          Dim MyPath As String
          Dim onVal As String
38        MyPath = ""
          Dim f As New FormExplorer
39        f.show (vbModal)
40        MyPath = f.PathSaving
41        If MyPath = "" Then Exit Sub
42        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
43        If f.FileName = "" Then Exit Sub
44        onVal = MyPath & f.FileName & ".XLS"
45        Set Me.GridXLS.DataSource = rsTempXLS
          
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
46        sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
47        sSQL = sSQL & " values ('" & gSA & "','FormReminder1000',getdate(),'" & IIf(IsNull(DtRemain0), "1900-01-01", Format(DtRemain0, "yyyy-mm-dd")) & "','" & Format(DtRemain, "yyyy-mm-dd") & "','cmdXlsMRS - " & cboType.Text & "','" & rsTempXLS.RecordCount & "','" & onVal & "','" & ComptName & "') "
48        Call objO.InsertData(sSQL)
          
49         GridXLS.ExportToFile onVal, True
50                Er.HeadCap = "Export to excel !"
51                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
52                Beep
53                Er.show (vbModal)
           'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
54         Exit Sub
Err_Export:
55          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
56          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub


Private Sub CreateTblXLS(Optional Edit As Boolean = False)
    On Error GoTo vb_error
    If rsTempXLS.State Then rsTempXLS.Close
    Set rsTempXLS = Nothing
    With rsTempXLS
        .Fields.Append "Nomor", adVarChar, 9
        .Fields.Append "Nama_Customer", adVarChar, 200
        .Fields.Append "Tgl_DO", adVarChar, 50
        .Fields.Append "Tgl_CR1", adVarChar, 50
        .Fields.Append "KM_CR1", adVarChar, 10
        .Fields.Append "Tgl_CR2", adVarChar, 50
        .Fields.Append "KM_CR2", adVarChar, 10
        .Fields.Append "No_Polisi", adVarChar, 25
        .Fields.Append "No_Rangka", adVarChar, 30
        .Fields.Append "Telp_Rumah", adVarChar, 15
        .Fields.Append "Telp_GSM", adVarChar, 15
        .Fields.Append "Telp_CDMA", adVarChar, 15
        .Fields.Append "Tgl_ETA", adVarChar, 50
        .Fields.Append "Tgl_SMS", adVarChar, 50
        .Fields.Append "Tgl_Servis", adVarChar, 50
        .Fields.Append "Tgl_Follow_Up", adVarChar, 50
        .Fields.Append "Komentar", adVarChar, 50
        .Fields.Append "Keterangan", adVarChar, 50
        .Fields.Append "Nama_Sales", adVarChar, 100
        .Open
    End With
    Me.GridXlsMrs.DataSource = rsTempXLS
    Exit Sub
vb_error:
    err.Raise err.Number, err.Source, err.Description
End Sub


Private Sub CreateMainList(Optional Edit As Boolean = False)
1         On Error GoTo vb_error
2         If rsList.State Then rsList.Close
3         Set rsList = Nothing
4         With rsList
5             .Fields.Append "NoSPK", adVarChar, 15
6             .Fields.Append "Tgl_DO", adVarChar, 50
7             .Fields.Append "Tgl_CR1", adVarChar, 50
8             .Fields.Append "KM_CR1", adVarChar, 10
9             .Fields.Append "Tgl_CR2", adVarChar, 50
10            .Fields.Append "KM_CR2", adVarChar, 10
11            .Fields.Append "Nama_Cust", adVarChar, 100
12            .Fields.Append "Alamat_Cust", adVarChar, 200
13            .Fields.Append "Kota_Cust", adVarChar, 50
14            .Fields.Append "KodePos_Cust", adVarChar, 10
15            .Fields.Append "No_Polisi", adVarChar, 25
16            .Fields.Append "No_Engine", adVarChar, 25
17            .Fields.Append "NoChassis", adVarChar, 30    '12
18            .Fields.Append "Telpon", adVarChar, 15
19            .Fields.Append "GSM", adVarChar, 15
20            .Fields.Append "CDMA", adVarChar, 15
21            .Fields.Append "Kode_Type", adVarChar, 50
22            .Fields.Append "Nama_Type", adVarChar, 50
23            .Fields.Append "Kode_Warna", adVarChar, 15
24            .Fields.Append "Nama_Warna", adVarChar, 50
25            .Fields.Append "Wiraniaga", adVarChar, 50
26            .Fields.Append "Nama_STNK", adVarChar, 100
27            .Fields.Append "Alamat_STNK", adVarChar, 200
28            .Fields.Append "Kota_STNK", adVarChar, 50
29            .Fields.Append "KodePos_STNK", adVarChar, 10
30            .Fields.Append "Tempat_Lahir", adVarChar, 50
31            .Fields.Append "Tanggal_Lahir", adVarChar, 50
32            .Fields.Append "Agama", adVarChar, 25
33            .Fields.Append "Pekerjaan", adVarChar, 50
34            .Fields.Append "Hoby", adVarChar, 25
35            .Fields.Append "Tgl_Service", adVarChar, 50
36            .Fields.Append "Select", adBoolean
37            .Fields.Append "Nomor_Rangka", adVarChar, 50
38            .Fields.Append "Kode_Lgn", adVarChar, 25
39            .Fields.Append "email", adVarChar, 50
40            .Fields.Append "statusdata", adVarChar, 100
41            .Open
42        End With

43        Me.GridRemainder.DataSource = rsList
44        Exit Sub
vb_error:
45        err.Raise err.Number, err.Source, err.Description
End Sub


