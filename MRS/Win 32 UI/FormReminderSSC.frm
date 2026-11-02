VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormReminderSSC 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormReminderSSC.frx":0000
   ScaleHeight     =   8355
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdHistory 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&History"
      Height          =   735
      Left            =   9120
      Picture         =   "FormReminderSSC.frx":E9B6
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   7560
      Width           =   855
   End
   Begin VB.CommandButton CmdXls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "X&LS"
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
      Left            =   10200
      Picture         =   "FormReminderSSC.frx":F023
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtModel 
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
      Left            =   5040
      MaxLength       =   100
      TabIndex        =   51
      Top             =   5640
      Width           =   4455
   End
   Begin VB.TextBox txtWarranty 
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   50
      Top             =   6000
      Width           =   7215
   End
   Begin VB.TextBox txtMDL 
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   49
      Top             =   5640
      Width           =   2655
   End
   Begin VB.TextBox txtId_Warranty 
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   48
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox txtKeterangan 
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   47
      Top             =   6360
      Width           =   7215
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   9360
      Top             =   6840
   End
   Begin VB.CommandButton cmdSMS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Send SMS"
      Height          =   735
      Left            =   6360
      Picture         =   "FormReminderSSC.frx":F49F
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7560
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
      Left            =   840
      TabIndex        =   39
      Top             =   7560
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
      Left            =   840
      TabIndex        =   38
      Top             =   7920
      Width           =   1695
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
      Left            =   11160
      Picture         =   "FormReminderSSC.frx":F821
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   7560
      Width           =   735
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
      Left            =   9840
      MaxLength       =   100
      TabIndex        =   14
      Top             =   4395
      Width           =   1815
   End
   Begin VB.TextBox txtTelpon 
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
      Left            =   9840
      MaxLength       =   100
      TabIndex        =   13
      Top             =   3675
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
      Left            =   9840
      MaxLength       =   100
      TabIndex        =   12
      Top             =   4035
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   11
      Top             =   4755
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
      Left            =   6720
      MaxLength       =   100
      TabIndex        =   10
      Top             =   4755
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   9
      Top             =   3315
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   8
      Top             =   3675
      Width           =   6015
   End
   Begin VB.TextBox txtNoEngine 
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   7
      Top             =   4395
      Width           =   2655
   End
   Begin VB.CommandButton cmdCall 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Call"
      Height          =   735
      Left            =   7320
      Picture         =   "FormReminderSSC.frx":FFB1
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   975
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   5
      Top             =   4035
      Width           =   2655
   End
   Begin VB.TextBox txtKodePos 
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
      Left            =   6720
      MaxLength       =   100
      TabIndex        =   4
      Top             =   4035
      Width           =   1575
   End
   Begin TrueOleDBGrid70.TDBGrid GridSSC 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nama_Customer"
      Columns(0).DataField=   "Nama_Customer"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Alamat"
      Columns(1).DataField=   "Alamat"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Kota"
      Columns(2).DataField=   "Kota"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "KodePos"
      Columns(3).DataField=   "KodePos"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "NoPolisi"
      Columns(4).DataField=   "NoPolisi"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "NoEngine"
      Columns(5).DataField=   "NoEngine"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "NoChassis"
      Columns(6).DataField=   "NoChassis"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "PSTN"
      Columns(7).DataField=   "PSTN"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "GSM"
      Columns(8).DataField=   "GSM"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "CDMA"
      Columns(9).DataField=   "CDMA"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Id_Warranty"
      Columns(10).DataField=   "Id_Warranty"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Kode_Model"
      Columns(11).DataField=   "Kode_Model"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Nama_Model"
      Columns(12).DataField=   "Nama_Model"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Warranty"
      Columns(13).DataField=   "Warranty"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Keterangan"
      Columns(14).DataField=   "Keterangan"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Tgl_Mulai"
      Columns(15).DataField=   "Tgl_Mulai"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Tgl_Selesai"
      Columns(16).DataField=   "Tgl_Selesai"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   17
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=17"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4842"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4736"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=5741"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5636"
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
      Splits(0)._ColumnProps(25)=   "Column(6).Width=3969"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3863"
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
      Splits(0)._ColumnProps(49)=   "Column(12).Width=3969"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3863"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=4419"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=4313"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=6165"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=6059"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=110,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=107,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=108,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=109,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=106,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=102,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=98,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=95,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=96,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=97,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=94,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=90,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=87,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=88,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=89,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=86,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=83,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=84,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=85,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=70,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=66,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=63,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=64,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=65,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=62,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=59,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=60,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=61,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=58,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=55,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=56,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=57,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=54,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=51,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=52,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=53,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=50,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=47,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=48,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=49,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=46,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=43,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=44,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=45,.parent=17"
      _StyleDefs(104) =   "Named:id=33:Normal"
      _StyleDefs(105) =   ":id=33,.parent=0"
      _StyleDefs(106) =   "Named:id=34:Heading"
      _StyleDefs(107) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(108) =   ":id=34,.wraptext=-1"
      _StyleDefs(109) =   "Named:id=35:Footing"
      _StyleDefs(110) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(111) =   "Named:id=36:Selected"
      _StyleDefs(112) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(113) =   "Named:id=37:Caption"
      _StyleDefs(114) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(115) =   "Named:id=38:HighlightRow"
      _StyleDefs(116) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(117) =   "Named:id=39:EvenRow"
      _StyleDefs(118) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(119) =   "Named:id=40:OddRow"
      _StyleDefs(120) =   ":id=40,.parent=33"
      _StyleDefs(121) =   "Named:id=41:RecordSelector"
      _StyleDefs(122) =   ":id=41,.parent=34"
      _StyleDefs(123) =   "Named:id=42:FilterBar"
      _StyleDefs(124) =   ":id=42,.parent=33"
   End
   Begin TDBDate6Ctl.TDBDate DtMulai 
      Height          =   315
      Left            =   3480
      TabIndex        =   44
      Top             =   6840
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      Calendar        =   "FormReminderSSC.frx":102D1
      Caption         =   "FormReminderSSC.frx":103FD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormReminderSSC.frx":10469
      Keys            =   "FormReminderSSC.frx":10487
      Spin            =   "FormReminderSSC.frx":104E5
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
   Begin TDBDate6Ctl.TDBDate DtSelesai 
      Height          =   315
      Left            =   5640
      TabIndex        =   45
      Top             =   6840
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   556
      Calendar        =   "FormReminderSSC.frx":1050D
      Caption         =   "FormReminderSSC.frx":10639
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormReminderSSC.frx":106A5
      Keys            =   "FormReminderSSC.frx":106C3
      Spin            =   "FormReminderSSC.frx":10721
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
   Begin MSComDlg.CommonDialog ComDialogXls 
      Left            =   11520
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   9000
      Shape           =   4  'Rounded Rectangle
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   10080
      Shape           =   4  'Rounded Rectangle
      Top             =   7440
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label23 
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
      Left            =   2160
      TabIndex        =   59
      Top             =   6000
      Width           =   75
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warranty"
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
      Left            =   600
      TabIndex        =   58
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Id Warranty"
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
      Left            =   600
      TabIndex        =   57
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Left            =   600
      TabIndex        =   56
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label19 
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
      Left            =   2160
      TabIndex        =   55
      Top             =   5280
      Width           =   75
   End
   Begin VB.Label Label18 
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
      Left            =   2160
      TabIndex        =   54
      Top             =   5640
      Width           =   75
   End
   Begin VB.Label Label2 
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
      Left            =   2160
      TabIndex        =   53
      Top             =   6360
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
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
      Left            =   600
      TabIndex        =   52
      Top             =   6360
      Width           =   1110
   End
   Begin VB.Label lblBerlaku 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Berlaku Selamanya"
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
      Left            =   7560
      TabIndex        =   46
      Top             =   6870
      Width           =   1800
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   7440
      Width           =   945
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5280
      X2              =   5520
      Y1              =   6975
      Y2              =   6975
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3240
      TabIndex        =   43
      Top             =   6870
      Width           =   75
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
      Left            =   2880
      TabIndex        =   42
      Top             =   7590
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
      Left            =   2880
      TabIndex        =   41
      Top             =   7920
      Width           =   1500
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   3240
      Y2              =   3240
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
      Left            =   8880
      TabIndex        =   36
      Top             =   3315
      Width           =   2775
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000000FF&
      X1              =   8880
      X2              =   11640
      Y1              =   3555
      Y2              =   3555
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
      Left            =   8880
      TabIndex        =   35
      Top             =   4395
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
      Left            =   9720
      TabIndex        =   34
      Top             =   4395
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
      Left            =   8880
      TabIndex        =   33
      Top             =   3675
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
      Left            =   8880
      TabIndex        =   32
      Top             =   4035
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
      Left            =   9720
      TabIndex        =   31
      Top             =   3675
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
      Left            =   9720
      TabIndex        =   30
      Top             =   4035
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
      Left            =   600
      TabIndex        =   29
      Top             =   4755
      Width           =   1065
   End
   Begin VB.Label Label6 
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
      Left            =   2160
      TabIndex        =   28
      Top             =   4755
      Width           =   75
   End
   Begin VB.Label Label5 
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
      Left            =   2160
      TabIndex        =   27
      Top             =   3675
      Width           =   75
   End
   Begin VB.Label Label15 
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
      Left            =   2160
      TabIndex        =   26
      Top             =   3315
      Width           =   75
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
      Left            =   600
      TabIndex        =   25
      Top             =   3675
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
      Left            =   600
      TabIndex        =   24
      Top             =   3315
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
      Left            =   5640
      TabIndex        =   23
      Top             =   4755
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
      Left            =   6600
      TabIndex        =   22
      Top             =   4755
      Width           =   75
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masa Berlaku Waranty SSC"
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
      Left            =   600
      TabIndex        =   21
      Top             =   6870
      Width           =   2550
   End
   Begin VB.Label Label73 
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
      Left            =   600
      TabIndex        =   20
      Top             =   4395
      Width           =   990
   End
   Begin VB.Label Label74 
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
      Left            =   2160
      TabIndex        =   19
      Top             =   4395
      Width           =   75
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   7245
      Y2              =   7245
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   7440
      Width           =   2145
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   840
      Left            =   720
      Top             =   7440
      Width           =   5295
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
      Left            =   2160
      TabIndex        =   18
      Top             =   4035
      Width           =   75
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
      Left            =   600
      TabIndex        =   17
      Top             =   4040
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
      Left            =   6600
      TabIndex        =   16
      Top             =   4035
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
      Left            =   5640
      TabIndex        =   15
      Top             =   4035
      Width           =   900
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIST REMINDER SSC"
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
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   9255
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   735
      X2              =   4455
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1800
      X2              =   4455
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2775
      X2              =   4455
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblJml 
      Alignment       =   1  'Right Justify
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
      Left            =   11040
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Data  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   495
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11685
   End
End
Attribute VB_Name = "FormReminderSSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private OnLoad As Boolean
Private rsW As ADODB.Recordset
Private rsC As ADODB.Recordset
Private objc As New ClsSSC
Private rsSSC As New ADODB.Recordset
Private rsR As ADODB.Recordset
Private rsP As ADODB.Recordset

Private NoRangka1 As Double
Private NoRangka2 As Double
Private idRangka As String

Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String

Private valSM As Integer
Private valCl As Integer

Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property

Private Sub chCall_Click()
   chCall.value = valCl
End Sub

Private Sub chSMS_Click()
   chSMS.value = valSM
End Sub

Private Sub cmdCall_Click()
   Dim f As New ReminderSscCall
   f.rsItemCall = rsSSC
'   f.Left = Me.Left
'   f.Top = 2250
   f.show (vbModal)
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdHistory_Click()
    Dim ft As New FormTransparan
    ft.xChassis = txtNoChassis
    ft.xIsForm = "HIST"
    ft.show (vbModal)
End Sub

Private Sub cmdSMS_Click()
   If rsSSC.RecordCount = 0 Then Exit Sub
   Dim f As New FormTransparan
   With f
      .xRsRem = rsSSC
      .xIsForm = "SMS_SSC"
      '.Top = 2250
      '.Left = Me.Left
      .show (vbModal)
   End With
   Call ViewRecords
End Sub

Private Sub CmdXls_Click()
1     On Error GoTo Err_Export
         Dim Er As New ErrInfo
2        If rsSSC.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data SSC Kosong !"
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
17        sSQL = sSQL & " values ('" & gSA & "','" & Me.Name & "',getdate(),'1900-01-01',getdate(),'','" & rsSSC.RecordCount & "','" & onVal & "','" & ComptName & "') "
18        Call objO.InsertData(sSQL)
          
19         GridSSC.Columns(1).Visible = False
20         GridSSC.ExportToFile onVal, True
21         GridSSC.Columns(1).Visible = True
22                Er.HeadCap = "Export to excel !"
23                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
24                Beep
25                Er.show (vbModal)
                  'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
26         Exit Sub
Err_Export:
27          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
28          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub Timer1_Timer()
   Dim R, G, B
   R = 255 * Rnd
   G = 255 * Rnd
   B = 255 * Rnd
   Me.lblBerlaku.ForeColor = RGB(R, G, B)
End Sub

Private Sub Form_Load()
   '' 20200512DEV
   If Log.disableExportMra = 1 Then
        If gTipeUser = "KABENG" Or gTipeUser = "ASSKABENG" Or gTipeUser = "ADMIN" Or gTipeUser = "HO" Then
            'cmdCSV.Enabled = True
            CmdXls.Enabled = True
        Else
            'cmdCSV.Enabled = False
            CmdXls.Enabled = False
        End If
   End If
   Call ListDataSSC
End Sub

Private Sub ListDataSSC()
1        Set rsW = objc.FetchDataHoWarranty
2        If rsW.RecordCount = 0 Then
            'MsgBox "Tidak ada data warranty SSC !", vbInformation + vbOKOnly, "Reminder SSC"
3           Unload Me
4           Exit Sub
5        End If
6        CreateTblSSC True
         Dim n As Integer
7        rsW.MoveFirst
8        Do Until rsW.EOF
9           n = Val(Len(Trim(rsW.Fields(3).value))) - 6
10          idRangka = Mid(Trim(rsW.Fields(3).value), 1, n)
11          NoRangka1 = Mid(Trim(rsW.Fields(3).value), n + 1, 6)
12          NoRangka2 = IIf((Len(Trim(rsW.Fields(4).value)) > 15), (Mid(Trim(rsW.Fields(4).value), n + 1, 6)), (Mid(Trim(rsW.Fields(3).value), n + 1, 6)))
            
13          Set rsC = objc.FetchDataCustdatWarranty(idRangka, n, NoRangka1, NoRangka2)
14          If rsC.RecordCount <> 0 Then
15             rsC.MoveFirst
16             Do Until rsC.EOF
17                If Val(Mid(Trim(rsC.Fields(6).value), n + 1, 6)) >= NoRangka1 And Val(Mid(Trim(rsC.Fields(6).value), n + 1, 6)) <= NoRangka2 Then
18                   rsSSC.AddNew
19                   rsSSC!Nama_Customer = Trim(NullText(rsC!CUSTOMER))
20                   rsSSC!Alamat = Left(Trim(NullText(rsC!cust_adres)), 100)
21                   rsSSC!Kota = Trim(NullText(rsC!Kota))
22                   rsSSC!KodePos = Trim(NullText(rsC!KodePos))
23                   rsSSC!NoPolisi = Trim(NullText(rsC!NO_POLISI))
24                   rsSSC!NoEngine = Trim(NullText(rsC!No_Engine))
25                   rsSSC!noChassis = Trim(NullText(rsC!No_Chassis))
26                   rsSSC!PSTN = Trim(NullText(rsC!Telpon))
27                   rsSSC!GSM = Trim(NullText(rsC!GSM))
28                   rsSSC!CDMA = Trim(NullText(rsC!CDMA))
29                   rsSSC!Id_warranty = Trim(NullText(rsW!Id_warranty))
30                   rsSSC!Kode_Model = Trim(NullText(rsW!MDL))
31                   rsSSC!Nama_Model = Trim(NullText(rsW!model))
32                   rsSSC!warranty = Trim(NullText(rsW!warranty))
33                   rsSSC!keterangan = Trim(NullText(rsW!keterangan))
34                   rsSSC!Tgl_Mulai = Format(CDate(rsW!TglMulai), "DD/MM/YYYY")
35                   rsSSC!Tgl_Selesai = Format(CDate(rsW!TglSelesai), "DD/MM/YYYY")
36                   rsSSC.Update
37                End If
38                rsC.MoveNext
39             Loop
40          End If
41          rsW.MoveNext
42       Loop
43       If rsSSC.RecordCount <> 0 Then
44          lblJml.Caption = rsSSC.RecordCount
45          rsSSC.MoveFirst
46          Call BindToControlSSC(rsSSC)
47          Me.cmdCall.Enabled = True
48          Me.cmdSMS.Enabled = True
49       Else
50          lblJml.Caption = "0"
51          Me.cmdCall.Enabled = False
52          Me.cmdSMS.Enabled = False
53       End If
End Sub


Private Sub CreateTblSSC(Optional Edit As Boolean = False)
1         On Error GoTo vb_error
2         If rsSSC.State Then rsSSC.Close
3         Set rsSSC = Nothing
4         With rsSSC
5             .Fields.Append "Nama_Customer", adVarChar, 50
6             .Fields.Append "Alamat", adVarChar, 100
7             .Fields.Append "Kota", adVarChar, 50
8             .Fields.Append "KodePos", adVarChar, 10
9             .Fields.Append "NoPolisi", adVarChar, 12
10            .Fields.Append "NoEngine", adVarChar, 15
11            .Fields.Append "NoChassis", adVarChar, 25
12            .Fields.Append "PSTN", adVarChar, 30
13            .Fields.Append "GSM", adVarChar, 20
14            .Fields.Append "CDMA", adVarChar, 20
15            .Fields.Append "Id_Warranty", adVarChar, 10
16            .Fields.Append "Kode_Model", adVarChar, 25
17            .Fields.Append "Nama_Model", adVarChar, 50
18            .Fields.Append "Warranty", adVarChar, 25
19            .Fields.Append "Keterangan", adVarChar, 30
20            .Fields.Append "Tgl_Mulai", adDate
21            .Fields.Append "Tgl_Selesai", adDate
22            .Open
23          Set GridSSC.DataSource = rsSSC
24        End With

25        Exit Sub
vb_error:
26        err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub BindToControlSSC(rs As ADODB.Recordset)
   With Me
      .txtNama.Text = rs.Fields(0).value
      .txtAlamat.Text = rs.Fields(1).value
      .txtKota.Text = rs.Fields(2).value
      .txtKodePos.Text = rs.Fields(3).value
      .txtNoPol.Text = rs.Fields(4).value
      .txtNoEngine.Text = rs.Fields(5).value
      .txtNoChassis.Text = rs.Fields(6).value
      .txtTelpon.Text = rs.Fields(7).value
      .txtGSM.Text = rs.Fields(8).value
      .txtCDMA.Text = rs.Fields(9).value
      .txtId_Warranty.Text = rs.Fields(10).value
      .txtMDL.Text = rs.Fields(11).value
      .txtModel.Text = rs.Fields(12).value
      .txtWarranty.Text = rs.Fields(13).value
      .txtKeterangan.Text = rs.Fields(14).value
      .dtMulai.value = Format(rs.Fields(15).value, "DD/MM/YYYY")
      .dtSelesai.value = Format(rs.Fields(16).value, "DD/MM/YYYY")
   End With
   
   If Me.dtMulai.value = Me.dtSelesai.value Then
      Me.lblBerlaku.Caption = "Lifetime Warranty"
   Else
      Dim hari As Double
      hari = DateDiff("m", dtMulai.value, dtSelesai.value)
      Me.lblBerlaku.Caption = "Berlaku " & hari & " Bulan"
   End If
   
   Call ViewRecords
End Sub

Private Sub ViewRecords()
   Set rsR = objc.FetchDataReminderSSC(txtNoChassis.Text, txtId_Warranty.Text, False)
   Dim nS As Integer
   nS = 0
   nSMS.Caption = ". . . . . . . . ."
   valSM = 0
   chSMS.value = valSM
   If rsR.RecordCount <> 0 Then
      rsR.MoveFirst
      Do Until rsR.EOF
         If rsR.Fields(1).value = "SC" Then
            nS = nS + 1
            nSMS.Caption = Trim(rsR.Fields(0).value) & "  SMS - " & nS
            valSM = 1
         End If
         rsR.MoveNext
      Loop
   End If
   Me.chSMS.value = valSM

   Set rsP = objc.FetchDataReminderSSC(txtNoChassis.Text, txtId_Warranty.Text, True)
   Dim nC As Integer
   nC = 0
   nCall.Caption = ". . . . . . . . ."
   valCl = 0
   chCall.value = valCl
   If rsP.RecordCount <> 0 Then
      rsP.MoveFirst
      Do Until rsP.EOF
         If rsP.Fields(1).value = "CL" Then
            nC = nC + 1
            nCall.Caption = Trim(rsP.Fields(0).value) & "  Call - " & nC
            valCl = 1
         End If
         rsP.MoveNext
      Loop
   End If
   Me.chCall.value = valCl
End Sub

Private Sub GridSSC_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If rsSSC.RecordCount <> 0 Then
      Call BindToControlSSC(rsSSC)
   End If
End Sub


Private Sub GridSSC_FilterChange()
On Error GoTo err
   Set cols = Me.GridSSC.Columns
   Dim C As Integer
   
   C = GridSSC.col
   GridSSC.HoldFields
   rsSSC.filter = getFilter()
   If rsSSC.EOF = True Then
      For Each col In GridSSC.Columns
         col.FilterText = ""
      Next col
      rsSSC.filter = ""
   End If
   GridSSC.col = C
   GridSSC.EditActive = True
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
          tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub GridSSC_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 18 Then rsSSC.Sort = GridSSC.Columns(ColIndex).DataField & " " & Urut
4        GridSSC.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form FormSMS_Inbox")

End Sub


