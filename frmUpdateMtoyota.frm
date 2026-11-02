VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmUpdateMtoyota 
   Caption         =   ".: Update Data MToyota :."
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text 
      Height          =   350
      Index           =   39
      Left            =   2760
      TabIndex        =   132
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmUpdateMtoyota.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   9015
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   13455
      Begin TabDlg.SSTab SSTab1 
         Height          =   6495
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   11456
         _Version        =   393216
         Tab             =   1
         TabHeight       =   520
         BackColor       =   14737632
         TabCaption(0)   =   "DATA INFORMASI CUSTOMER"
         TabPicture(0)   =   "frmUpdateMtoyota.frx":8B3C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "DATA INFORMASI KENDARAAN"
         TabPicture(1)   =   "frmUpdateMtoyota.frx":8B58
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "DATA KELENGKAPAN LAIN"
         TabPicture(2)   =   "frmUpdateMtoyota.frx":8B74
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame4"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame4 
            Caption         =   "Kelengkapan Lain"
            Height          =   6015
            Left            =   -74880
            TabIndex        =   116
            Top             =   360
            Width           =   12975
            Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
               Height          =   2295
               Left            =   120
               TabIndex        =   117
               Top             =   240
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   4048
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Insurance Provider"
               Columns(0).DataField=   "insurance_provider"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Insurance Provider Other"
               Columns(1).DataField=   "insurance_provider_other"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Insurance Policy Number"
               Columns(2).DataField=   "insurance_policy_number"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   953
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   15790320
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=3969"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3889"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=3969"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3889"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               Caption         =   "ToyotaID Insurance Coverage"
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   15790320
               RowDividerColor =   15790320
               RowSubDividerColor=   15790320
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
               _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(48)  =   "Named:id=33:Normal"
               _StyleDefs(49)  =   ":id=33,.parent=0"
               _StyleDefs(50)  =   "Named:id=34:Heading"
               _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(52)  =   ":id=34,.wraptext=-1"
               _StyleDefs(53)  =   "Named:id=35:Footing"
               _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(55)  =   "Named:id=36:Selected"
               _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(57)  =   "Named:id=37:Caption"
               _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(59)  =   "Named:id=38:HighlightRow"
               _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=39:EvenRow"
               _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(63)  =   "Named:id=40:OddRow"
               _StyleDefs(64)  =   ":id=40,.parent=33"
               _StyleDefs(65)  =   "Named:id=41:RecordSelector"
               _StyleDefs(66)  =   ":id=41,.parent=34"
               _StyleDefs(67)  =   "Named:id=42:FilterBar"
               _StyleDefs(68)  =   ":id=42,.parent=33"
            End
            Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
               Height          =   2775
               Left            =   120
               TabIndex        =   118
               Top             =   2640
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   4895
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Location Preference"
               Columns(0).DataField=   "location_preference"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   1
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   953
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   15790320
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=1"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               Caption         =   "ToyotaID Outlet Location"
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   15790320
               RowDividerColor =   15790320
               RowSubDividerColor=   15790320
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
               _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.locked=-1"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
               _StyleDefs(40)  =   "Named:id=33:Normal"
               _StyleDefs(41)  =   ":id=33,.parent=0"
               _StyleDefs(42)  =   "Named:id=34:Heading"
               _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(44)  =   ":id=34,.wraptext=-1"
               _StyleDefs(45)  =   "Named:id=35:Footing"
               _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(47)  =   "Named:id=36:Selected"
               _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(49)  =   "Named:id=37:Caption"
               _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(51)  =   "Named:id=38:HighlightRow"
               _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(53)  =   "Named:id=39:EvenRow"
               _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(55)  =   "Named:id=40:OddRow"
               _StyleDefs(56)  =   ":id=40,.parent=33"
               _StyleDefs(57)  =   "Named:id=41:RecordSelector"
               _StyleDefs(58)  =   ":id=41,.parent=34"
               _StyleDefs(59)  =   "Named:id=42:FilterBar"
               _StyleDefs(60)  =   ":id=42,.parent=33"
            End
            Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
               Height          =   2775
               Left            =   6480
               TabIndex        =   119
               Top             =   2640
               Width           =   6375
               _ExtentX        =   11245
               _ExtentY        =   4895
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Outlet Preference"
               Columns(0).DataField=   "outlet_preference"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   1
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   953
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   15790320
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=1"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               Caption         =   "ToyotaID Outlet Preference"
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   15790320
               RowDividerColor =   15790320
               RowSubDividerColor=   15790320
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
               _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.locked=-1"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
               _StyleDefs(40)  =   "Named:id=33:Normal"
               _StyleDefs(41)  =   ":id=33,.parent=0"
               _StyleDefs(42)  =   "Named:id=34:Heading"
               _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(44)  =   ":id=34,.wraptext=-1"
               _StyleDefs(45)  =   "Named:id=35:Footing"
               _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(47)  =   "Named:id=36:Selected"
               _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(49)  =   "Named:id=37:Caption"
               _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(51)  =   "Named:id=38:HighlightRow"
               _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(53)  =   "Named:id=39:EvenRow"
               _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(55)  =   "Named:id=40:OddRow"
               _StyleDefs(56)  =   ":id=40,.parent=33"
               _StyleDefs(57)  =   "Named:id=41:RecordSelector"
               _StyleDefs(58)  =   ":id=41,.parent=34"
               _StyleDefs(59)  =   "Named:id=42:FilterBar"
               _StyleDefs(60)  =   ":id=42,.parent=33"
            End
            Begin TrueOleDBGrid70.TDBGrid TDBGrid4 
               Height          =   2295
               Left            =   6480
               TabIndex        =   142
               Top             =   240
               Width           =   6375
               _ExtentX        =   11245
               _ExtentY        =   4048
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Insurance Policy Number"
               Columns(0).DataField=   "insurance_policy_number"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Insurance Type"
               Columns(1).DataField=   "insurance_type"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Insurance Coverage"
               Columns(2).DataField=   "insurance_coverage"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   3
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   953
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   15790320
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=3"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
               Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
               Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               Caption         =   "ToyotaID Insurance Coverage Policies"
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   15790320
               RowDividerColor =   15790320
               RowSubDividerColor=   15790320
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
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
               _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(48)  =   "Named:id=33:Normal"
               _StyleDefs(49)  =   ":id=33,.parent=0"
               _StyleDefs(50)  =   "Named:id=34:Heading"
               _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(52)  =   ":id=34,.wraptext=-1"
               _StyleDefs(53)  =   "Named:id=35:Footing"
               _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(55)  =   "Named:id=36:Selected"
               _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(57)  =   "Named:id=37:Caption"
               _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(59)  =   "Named:id=38:HighlightRow"
               _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(61)  =   "Named:id=39:EvenRow"
               _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(63)  =   "Named:id=40:OddRow"
               _StyleDefs(64)  =   ":id=40,.parent=33"
               _StyleDefs(65)  =   "Named:id=41:RecordSelector"
               _StyleDefs(66)  =   ":id=41,.parent=34"
               _StyleDefs(67)  =   "Named:id=42:FilterBar"
               _StyleDefs(68)  =   ":id=42,.parent=33"
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Data Informasi Kendaraan"
            Height          =   6015
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   12975
            Begin TDBDate6Ctl.TDBDate TDBDate4 
               Height          =   345
               Left            =   8880
               TabIndex        =   145
               Top             =   2760
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   609
               Calendar        =   "frmUpdateMtoyota.frx":8B90
               Caption         =   "frmUpdateMtoyota.frx":8CBC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmUpdateMtoyota.frx":8D28
               Keys            =   "frmUpdateMtoyota.frx":8D46
               Spin            =   "frmUpdateMtoyota.frx":8DA4
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
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
               Text            =   "30-10-2025"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   45960
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate TDBDate3 
               Height          =   345
               Left            =   8880
               TabIndex        =   144
               Top             =   2400
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   609
               Calendar        =   "frmUpdateMtoyota.frx":8DCC
               Caption         =   "frmUpdateMtoyota.frx":8EF8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmUpdateMtoyota.frx":8F64
               Keys            =   "frmUpdateMtoyota.frx":8F82
               Spin            =   "frmUpdateMtoyota.frx":8FE0
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
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
               Text            =   "30-10-2025"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   45960
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate TDBDate2 
               Height          =   345
               Left            =   8880
               TabIndex        =   143
               Top             =   960
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   609
               Calendar        =   "frmUpdateMtoyota.frx":9008
               Caption         =   "frmUpdateMtoyota.frx":9134
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmUpdateMtoyota.frx":91A0
               Keys            =   "frmUpdateMtoyota.frx":91BE
               Spin            =   "frmUpdateMtoyota.frx":921C
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
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
               Text            =   "30-10-2025"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   45960
               CenturyMode     =   0
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   42
               Left            =   8880
               MaxLength       =   250
               TabIndex        =   140
               Top             =   3480
               Width           =   2895
            End
            Begin VB.ComboBox Combo9 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":9244
               Left            =   8880
               List            =   "frmUpdateMtoyota.frx":924E
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   3120
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   41
               Left            =   8880
               TabIndex        =   136
               Top             =   2760
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   40
               Left            =   8880
               TabIndex        =   134
               Top             =   2400
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox Text 
               BackColor       =   &H00C0E0FF&
               Enabled         =   0   'False
               Height          =   350
               Index           =   38
               Left            =   8880
               TabIndex        =   130
               Top             =   2040
               Width           =   2895
            End
            Begin VB.TextBox Text 
               BackColor       =   &H00C0E0FF&
               Enabled         =   0   'False
               Height          =   350
               Index           =   37
               Left            =   8880
               TabIndex        =   128
               Top             =   1680
               Width           =   2895
            End
            Begin VB.ComboBox Combo8 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":9267
               Left            =   9240
               List            =   "frmUpdateMtoyota.frx":9271
               Style           =   2  'Dropdown List
               TabIndex        =   126
               Top             =   5400
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   36
               Left            =   9240
               TabIndex        =   124
               Top             =   5040
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   35
               Left            =   9240
               TabIndex        =   122
               Top             =   4680
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   34
               Left            =   9240
               TabIndex        =   120
               Top             =   4320
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   33
               Left            =   8880
               MaxLength       =   256
               TabIndex        =   114
               Top             =   1320
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   32
               Left            =   8880
               TabIndex        =   112
               Top             =   960
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   31
               Left            =   8880
               MaxLength       =   64
               TabIndex        =   110
               Top             =   600
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   30
               Left            =   8880
               MaxLength       =   32
               TabIndex        =   108
               Top             =   240
               Width           =   2895
            End
            Begin VB.CheckBox ckMainVhc 
               Height          =   255
               Left            =   3000
               TabIndex        =   106
               Top             =   4560
               Width           =   375
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":928A
               Left            =   3000
               List            =   "frmUpdateMtoyota.frx":9294
               Style           =   2  'Dropdown List
               TabIndex        =   104
               Top             =   4200
               Width           =   2895
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":92AD
               Left            =   3000
               List            =   "frmUpdateMtoyota.frx":92B7
               Style           =   2  'Dropdown List
               TabIndex        =   102
               Top             =   3840
               Width           =   2895
            End
            Begin VB.ComboBox Combo5 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":92D0
               Left            =   3000
               List            =   "frmUpdateMtoyota.frx":92DA
               Style           =   2  'Dropdown List
               TabIndex        =   100
               Top             =   3480
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   29
               Left            =   3000
               MaxLength       =   20
               TabIndex        =   98
               Top             =   3120
               Width           =   2895
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":92F3
               Left            =   3000
               List            =   "frmUpdateMtoyota.frx":92FD
               Style           =   2  'Dropdown List
               TabIndex        =   96
               Top             =   960
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   28
               Left            =   3000
               MaxLength       =   64
               TabIndex        =   94
               Top             =   2760
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   27
               Left            =   3000
               MaxLength       =   128
               TabIndex        =   92
               Top             =   2400
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   26
               Left            =   3000
               MaxLength       =   16
               TabIndex        =   90
               Top             =   2040
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   25
               Left            =   3000
               MaxLength       =   64
               TabIndex        =   88
               Top             =   1680
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   24
               Left            =   3000
               TabIndex        =   86
               Top             =   1320
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   23
               Left            =   3000
               MaxLength       =   16
               TabIndex        =   84
               Top             =   600
               Width           =   2895
            End
            Begin VB.TextBox Text 
               BackColor       =   &H00C0E0FF&
               Enabled         =   0   'False
               Height          =   350
               Index           =   22
               Left            =   3000
               TabIndex        =   82
               Top             =   240
               Width           =   2895
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Reasons For Change Others"
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
               Index           =   66
               Left            =   6120
               TabIndex        =   141
               Top             =   3525
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Reasons For Change"
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
               Index           =   65
               Left            =   6000
               TabIndex        =   139
               Top             =   3150
               Width           =   2745
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Valid To"
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
               Index           =   64
               Left            =   6120
               TabIndex        =   137
               Top             =   2805
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Valid From"
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
               Index           =   63
               Left            =   6120
               TabIndex        =   135
               Top             =   2445
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Insurance Start Datetime"
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
               Index           =   61
               Left            =   6120
               TabIndex        =   131
               Top             =   2085
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Insurance End Datetime"
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
               Index           =   60
               Left            =   6120
               TabIndex        =   129
               Top             =   1725
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Insurance Type"
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
               Index           =   59
               Left            =   6240
               TabIndex        =   127
               Top             =   5430
               Visible         =   0   'False
               Width           =   2745
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Insurance Policy Number"
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
               Index           =   58
               Left            =   6480
               TabIndex        =   125
               Top             =   5085
               Visible         =   0   'False
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Insurance Provider Other"
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
               Index           =   57
               Left            =   6480
               TabIndex        =   123
               Top             =   4725
               Visible         =   0   'False
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Insurance Provider"
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
               Index           =   56
               Left            =   6480
               TabIndex        =   121
               Top             =   4365
               Visible         =   0   'False
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Stnk Address"
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
               Index           =   55
               Left            =   6120
               TabIndex        =   115
               Top             =   1365
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Stnk Expiry Date"
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
               Index           =   54
               Left            =   6120
               TabIndex        =   113
               Top             =   1005
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Stnk Name"
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
               Index           =   53
               Left            =   6120
               TabIndex        =   111
               Top             =   645
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Stnk Number"
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
               Index           =   52
               Left            =   6120
               TabIndex        =   109
               Top             =   285
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Main Vehicle Flag"
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
               Index           =   51
               Left            =   240
               TabIndex        =   107
               Top             =   4560
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Interest To Tintouch"
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
               Index           =   50
               Left            =   120
               TabIndex        =   105
               Top             =   4230
               Width           =   2745
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Interest To Tcare"
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
               Index           =   49
               Left            =   360
               TabIndex        =   103
               Top             =   3870
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Category"
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
               Index           =   48
               Left            =   360
               TabIndex        =   101
               Top             =   3510
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Miles Travelled"
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
               Index           =   47
               Left            =   240
               TabIndex        =   99
               Top             =   3165
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Primary User"
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
               Index           =   46
               Left            =   360
               TabIndex        =   97
               Top             =   990
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Color"
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
               Index           =   45
               Left            =   240
               TabIndex        =   95
               Top             =   2805
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Variant"
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
               Index           =   44
               Left            =   240
               TabIndex        =   93
               Top             =   2445
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Model"
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
               Index           =   43
               Left            =   240
               TabIndex        =   91
               Top             =   2085
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Color Code"
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
               Index           =   42
               Left            =   240
               TabIndex        =   89
               Top             =   1725
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Katashiki Suffix"
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
               Index           =   41
               Left            =   240
               TabIndex        =   87
               Top             =   1365
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Police Number"
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
               Index           =   40
               Left            =   240
               TabIndex        =   85
               Top             =   645
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "VIN"
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
               Index           =   31
               Left            =   240
               TabIndex        =   83
               Top             =   285
               Width           =   2505
            End
         End
         Begin VB.Frame Frame2 
            Height          =   5895
            Left            =   -74880
            TabIndex        =   10
            Top             =   360
            Width           =   12975
            Begin VB.CheckBox ckPrimary 
               Height          =   255
               Left            =   8640
               TabIndex        =   80
               Top             =   5520
               Width           =   375
            End
            Begin VB.TextBox Text 
               Height          =   585
               Index           =   21
               Left            =   8640
               MaxLength       =   256
               MultiLine       =   -1  'True
               TabIndex        =   79
               Top             =   4920
               Width           =   4215
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   20
               Left            =   8640
               MaxLength       =   5
               TabIndex        =   78
               Top             =   4560
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   19
               Left            =   8640
               MaxLength       =   64
               TabIndex        =   77
               Top             =   4200
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   18
               Left            =   8640
               MaxLength       =   64
               TabIndex        =   76
               Top             =   3840
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   17
               Left            =   8640
               MaxLength       =   64
               TabIndex        =   75
               Top             =   3480
               Width           =   2895
            End
            Begin MSMask.MaskEdBox txtTime 
               Height          =   350
               Index           =   0
               Left            =   8640
               TabIndex        =   71
               Top             =   360
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   609
               _Version        =   393216
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   "_"
            End
            Begin VB.CheckBox ckNews 
               Height          =   255
               Left            =   8640
               TabIndex        =   70
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   13
               Left            =   8640
               MaxLength       =   32
               TabIndex        =   69
               Top             =   1080
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   14
               Left            =   8640
               MaxLength       =   156
               TabIndex        =   68
               Top             =   2400
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   15
               Left            =   8640
               MaxLength       =   156
               TabIndex        =   67
               Top             =   2760
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   16
               Left            =   8640
               MaxLength       =   64
               TabIndex        =   66
               Top             =   3120
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   9
               Left            =   2760
               MaxLength       =   64
               TabIndex        =   65
               Top             =   3720
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   10
               Left            =   2760
               MaxLength       =   150
               TabIndex        =   64
               Top             =   4440
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   11
               Left            =   2760
               MaxLength       =   150
               TabIndex        =   63
               Top             =   4800
               Width           =   2895
            End
            Begin VB.TextBox Text 
               BackColor       =   &H00C0E0FF&
               Enabled         =   0   'False
               Height          =   350
               Index           =   12
               Left            =   2760
               TabIndex        =   62
               Top             =   5160
               Width           =   2895
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":9316
               Left            =   2760
               List            =   "frmUpdateMtoyota.frx":9320
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   5520
               Width           =   2895
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":9339
               Left            =   2760
               List            =   "frmUpdateMtoyota.frx":9346
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   4080
               Width           =   2895
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "frmUpdateMtoyota.frx":935F
               Left            =   2760
               List            =   "frmUpdateMtoyota.frx":936C
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   3000
               Width           =   2895
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   345
               Left            =   2760
               TabIndex        =   43
               Top             =   2640
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   609
               Calendar        =   "frmUpdateMtoyota.frx":9386
               Caption         =   "frmUpdateMtoyota.frx":94B2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmUpdateMtoyota.frx":951E
               Keys            =   "frmUpdateMtoyota.frx":953C
               Spin            =   "frmUpdateMtoyota.frx":959A
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
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
               Text            =   "30-10-2025"
               ValidateMode    =   0
               ValueVT         =   7
               Value           =   45960
               CenturyMode     =   0
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   8
               Left            =   2760
               MaxLength       =   16
               TabIndex        =   26
               Top             =   3360
               Width           =   2895
            End
            Begin VB.TextBox Text 
               BackColor       =   &H00C0E0FF&
               Enabled         =   0   'False
               Height          =   350
               Index           =   7
               Left            =   2760
               TabIndex        =   24
               Top             =   2280
               Width           =   2895
            End
            Begin VB.TextBox Text 
               BackColor       =   &H00C0E0FF&
               Enabled         =   0   'False
               Height          =   350
               Index           =   6
               Left            =   2760
               TabIndex        =   22
               Top             =   1920
               Width           =   2895
            End
            Begin VB.TextBox Text 
               BackColor       =   &H00C0E0FF&
               Enabled         =   0   'False
               Height          =   350
               Index           =   5
               Left            =   2760
               TabIndex        =   20
               Top             =   1560
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   4
               Left            =   2760
               MaxLength       =   64
               TabIndex        =   18
               Top             =   1200
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   3
               Left            =   2760
               MaxLength       =   64
               TabIndex        =   16
               Top             =   840
               Width           =   2895
            End
            Begin VB.TextBox Text 
               Height          =   350
               Index           =   2
               Left            =   2760
               MaxLength       =   64
               TabIndex        =   14
               Top             =   480
               Width           =   2895
            End
            Begin VB.TextBox Text 
               BackColor       =   &H00C0E0FF&
               Enabled         =   0   'False
               Height          =   350
               Index           =   1
               Left            =   2760
               TabIndex        =   12
               Top             =   120
               Width           =   2895
            End
            Begin MSMask.MaskEdBox txtTime 
               Height          =   345
               Index           =   1
               Left            =   8640
               TabIndex        =   72
               Top             =   720
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   609
               _Version        =   393216
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   "_"
            End
            Begin ucXPButton.XPButton cmdUploadNPWP 
               Height          =   495
               Left            =   8640
               TabIndex        =   73
               Top             =   1440
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   873
               BTYPE           =   2
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
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmUpdateMtoyota.frx":95C2
               PICN            =   "frmUpdateMtoyota.frx":95DE
               PICH            =   "frmUpdateMtoyota.frx":B258
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ucXPButton.XPButton cmdHapus 
               Height          =   495
               Left            =   8640
               TabIndex        =   74
               Top             =   1920
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   873
               BTYPE           =   2
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
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmUpdateMtoyota.frx":CFA0
               PICN            =   "frmUpdateMtoyota.frx":CFBC
               PICH            =   "frmUpdateMtoyota.frx":106FA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "News Promotion Event Flag"
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
               Index           =   39
               Left            =   5880
               TabIndex        =   61
               Top             =   120
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Prefered Contact Time Start"
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
               Index           =   38
               Left            =   5880
               TabIndex        =   60
               Top             =   480
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Prefered Contact Time End"
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
               Index           =   37
               Left            =   5880
               TabIndex        =   59
               Top             =   840
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "NPWP Number"
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
               Index           =   36
               Left            =   5880
               TabIndex        =   58
               Top             =   1200
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "NPWP File"
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
               Index           =   35
               Left            =   5880
               TabIndex        =   57
               Top             =   1560
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "KTP File"
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
               Index           =   34
               Left            =   5880
               TabIndex        =   56
               Top             =   2040
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Address Label"
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
               Index           =   33
               Left            =   5880
               TabIndex        =   55
               Top             =   2450
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Residence Address"
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
               Index           =   32
               Left            =   5880
               TabIndex        =   54
               Top             =   2808
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Province"
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
               Index           =   30
               Left            =   5880
               TabIndex        =   53
               Top             =   3168
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "City"
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
               Index           =   29
               Left            =   5880
               TabIndex        =   52
               Top             =   3528
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "District"
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
               Index           =   28
               Left            =   5880
               TabIndex        =   51
               Top             =   3888
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Sub District"
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
               Index           =   27
               Left            =   5880
               TabIndex        =   50
               Top             =   4248
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Postal Code"
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
               Index           =   26
               Left            =   5880
               TabIndex        =   49
               Top             =   4608
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Detail Address"
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
               Index           =   25
               Left            =   5880
               TabIndex        =   48
               Top             =   4968
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Is Primary"
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
               Index           =   24
               Left            =   5880
               TabIndex        =   47
               Top             =   5520
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Category"
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
               Index           =   23
               Left            =   120
               TabIndex        =   42
               Top             =   5520
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Toyota Single ID Status"
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
               Index           =   22
               Left            =   120
               TabIndex        =   41
               Top             =   5160
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Company Business Sector"
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
               Index           =   21
               Left            =   120
               TabIndex        =   40
               Top             =   4800
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Company Name"
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
               Index           =   20
               Left            =   120
               TabIndex        =   39
               Top             =   4440
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Gender"
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
               Index           =   19
               Left            =   120
               TabIndex        =   38
               Top             =   4080
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Occupation"
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
               Index           =   18
               Left            =   120
               TabIndex        =   37
               Top             =   3720
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "KTP Number"
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
               Index           =   17
               Left            =   120
               TabIndex        =   36
               Top             =   3408
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Verification Channel"
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
               Index           =   16
               Left            =   120
               TabIndex        =   35
               Top             =   3030
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date"
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
               Index           =   7
               Left            =   120
               TabIndex        =   25
               Top             =   2685
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
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
               Index           =   6
               Left            =   120
               TabIndex        =   23
               Top             =   2325
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Phone Number Secondary"
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
               Index           =   5
               Left            =   120
               TabIndex        =   21
               Top             =   1965
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Phone Number"
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
               Index           =   4
               Left            =   120
               TabIndex        =   19
               Top             =   1605
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nick Name"
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
               Index           =   3
               Left            =   120
               TabIndex        =   17
               Top             =   1245
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
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
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Top             =   885
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
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
               Index           =   1
               Left            =   120
               TabIndex        =   13
               Top             =   525
               Width           =   2505
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "One Account ID"
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
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   165
               Width           =   2505
            End
         End
      End
      Begin VB.TextBox Text 
         Height          =   495
         Index           =   0
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
      Begin ARProgBarCtrl.ARProgressBar ProgBar 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   -240
         Visible         =   0   'False
         Width           =   13125
         _ExtentX        =   23151
         _ExtentY        =   1085
         Value           =   0
         ForeColor       =   52479
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionForeColor=   128
         IniColor        =   52479
      End
      Begin ucXPButton.XPButton cmdView 
         Height          =   495
         Left            =   4800
         TabIndex        =   3
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   2
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUpdateMtoyota.frx":13FDE
         PICN            =   "frmUpdateMtoyota.frx":13FFA
         PICH            =   "frmUpdateMtoyota.frx":17D5B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   12120
         TabIndex        =   5
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUpdateMtoyota.frx":1BC92
         PICN            =   "frmUpdateMtoyota.frx":1BCAE
         PICH            =   "frmUpdateMtoyota.frx":225E1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdBatal 
         Height          =   495
         Left            =   10920
         TabIndex        =   6
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUpdateMtoyota.frx":29298
         PICN            =   "frmUpdateMtoyota.frx":292B4
         PICH            =   "frmUpdateMtoyota.frx":2F8AD
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
         Left            =   9720
         TabIndex        =   7
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUpdateMtoyota.frx":3615C
         PICN            =   "frmUpdateMtoyota.frx":36178
         PICH            =   "frmUpdateMtoyota.frx":3C97B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   495
         Left            =   8520
         TabIndex        =   146
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUpdateMtoyota.frx":43521
         PICN            =   "frmUpdateMtoyota.frx":4353D
         PICH            =   "frmUpdateMtoyota.frx":49CE0
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
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   12600
         X2              =   120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka"
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
         Left            =   315
         TabIndex        =   1
         Top             =   360
         Width           =   1305
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stnk Address"
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
      Index           =   62
      Left            =   0
      TabIndex        =   133
      Top             =   45
      Width           =   2505
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Account ID"
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
      Index           =   15
      Left            =   960
      TabIndex        =   34
      Top             =   7365
      Width           =   2505
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Account ID"
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
      Index           =   14
      Left            =   960
      TabIndex        =   33
      Top             =   7005
      Width           =   2505
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Account ID"
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
      Index           =   13
      Left            =   960
      TabIndex        =   32
      Top             =   6645
      Width           =   2505
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Account ID"
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
      Index           =   12
      Left            =   960
      TabIndex        =   31
      Top             =   6285
      Width           =   2505
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Account ID"
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
      Index           =   11
      Left            =   960
      TabIndex        =   30
      Top             =   5925
      Width           =   2505
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Account ID"
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
      Index           =   10
      Left            =   960
      TabIndex        =   29
      Top             =   5565
      Width           =   2505
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Account ID"
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
      Index           =   9
      Left            =   960
      TabIndex        =   28
      Top             =   5205
      Width           =   2505
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "One Account ID"
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
      Index           =   8
      Left            =   960
      TabIndex        =   27
      Top             =   4845
      Width           =   2505
   End
End
Attribute VB_Name = "frmUpdateMtoyota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmUpdateMtoyota
' DateTime  : 19/09/2018
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------


Dim rsGrid      As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsTemp1, rsTemp2, rsTemp3, rsTemp4, rsTemp5  As New ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset
Dim cols        As TrueOleDBGrid70.Columns
Dim Urut        As String
Dim isTambah    As Boolean
Dim isBegin     As Boolean
Dim isLoad      As Boolean
Private isEdit  As Boolean

Private dppJasa, opexLembur, ketersediaan, totalBobot, iProgress As Double

Private Sub ber_sih()
    For i = 1 To 42
        Text(i) = ""
    Next
    TDBDate1.value = "01-01-1900"
    TDBDate2.value = "01-01-1900"
    TDBDate3.value = "01-01-1900"
    TDBDate4.value = "01-01-1900"
    txtTime(0) = "00:00"
    txtTime(1) = "00:00"
    ckNews.value = Unchecked
    ckPrimary.value = Unchecked
    ckMainVhc.value = Unchecked
    Combo1.Clear
    Combo1.AddItem ("SMS")
    Combo1.AddItem ("WHATSAPP")
    Combo1.AddItem ("EMAIL")
    
    Combo2.Clear
    Combo2.AddItem ("MALE")
    Combo2.AddItem ("FEMALE")
    Combo2.AddItem ("OTHER")
    
    Combo3.Clear
    Combo3.AddItem ("INDIVIDUAL")
    Combo3.AddItem ("COMPANY")
    
    Combo4.Clear
    Combo4.AddItem ("MASTER")
    Combo4.AddItem ("MEMBER")
    
    Combo5.Clear
    Combo5.AddItem ("RETAIL")
    Combo5.AddItem ("FLEET")
    
    Combo6.Clear
    Combo6.AddItem ("EXT")
    Combo6.AddItem ("DONT_UNDERSTAND")
    Combo6.AddItem ("SCHEDULE")
    Combo6.AddItem ("BAD_EXPERIENCE")
    Combo6.AddItem ("DONT_CARE")
    Combo6.AddItem ("FREE_SERVICE")
    Combo6.AddItem ("STD_OTHERS")
    Combo6.AddItem ("LIMITED_USAGE")
    Combo6.AddItem ("LIMIT_SERVICE")
    Combo6.AddItem ("DOUBT")
    Combo6.AddItem ("NON_OFFICIAL")
    Combo6.AddItem ("NOT_INT_OTHERS")

    Combo7.Clear
    Combo7.AddItem ("ACTV_SOON")
    Combo7.AddItem ("DONT_NEED")
    Combo7.AddItem ("DATA_CONCERN")
    Combo7.AddItem ("NOT_FAMILIAR")
    Combo7.AddItem ("LIMITED_INTERNET")
    Combo7.AddItem ("RELIABILITY_CONCERN")
    Combo7.AddItem ("AUTOMATIC_CHARGE")
    Combo7.AddItem ("COMPLICATED")
    Combo7.AddItem ("LACK_SUPPORT")
    Combo7.AddItem ("NOT_INT_OTHERS")
    
    
    Combo8.Clear
    Combo8.AddItem ("TLO")
    Combo8.AddItem ("ALL_RISK")
    Combo8.AddItem ("COMPREHENSIVE")
    Combo8.AddItem ("COMBINATION")
    
    Combo9.Clear
    Combo9.AddItem ("TRANSFER_TO_FAMILY/FRIENDS")
    Combo9.AddItem ("SOLD_TO_ANOTHER_PARTY")
    Combo9.AddItem ("OTHERS")
    
    TDBGrid1.DataSource = Nothing
    TDBGrid2.DataSource = Nothing
    TDBGrid3.DataSource = Nothing
    TDBGrid4.DataSource = Nothing
    isLoad = False
    isEdit = True
    
    Text(1).Enabled = False
    Text(5).Enabled = False
    Text(6).Enabled = False
    Text(7).Enabled = False
    Text(22).Enabled = False
    
End Sub


Private Sub cmdSimpan_Click()
1         If isEdit = True Then
2             If MsgBox("Apakah yakin mengubah data Customer MToyota", vbYesNo, "Konfirmasi Perubahan Data Mtoyota") = vbYes Then
3                 sSQL = "UPDATE ReplBengkel99..p2_ToyotaID SET first_name = '" & NulSID(Text(2)) & "',last_name = '" & NulSID(Text(3)) & "', nick_name = '" & NulSID(Text(4)) & "',phone_number = '" & NulSID(Text(5)) & "',email = '" & NulSID(Text(7)) & "', birth_date = '" & Format(TDBDate1, "yyyy-mm-dd") & "',"
4                 sSQL = sSQL & " verification_channel = '" & NulSID(Combo1.Text) & "',ktp_number = '" & NulSID(Text(8)) & "',occupation = '" & NulSID(Text(9)) & "',gender = '" & NulSID(Combo2.Text) & "',address_label = '" & NulSID(Text(14)) & "',residence_address = '" & NulSID(Text(15)) & "',"
5                 sSQL = sSQL & " province = '" & NulSID(Text(16)) & "',city = '" & NulSID(Text(17)) & "',district = '" & NulSID(Text(18)) & "',subdistrict = '" & NulSID(Text(19)) & "',postal_code = '" & NulSID(Text(20)) & "',detail_address = '" & NulSID(Text(21)) & "',"
6                 sSQL = sSQL & " toyota_single_ID_status = '" & NulSID(Text(12)) & "',customer_category = '" & NulSID(Combo3.Text) & "', primary_user = '" & NulSID(Combo4.Text) & "',police_number = '" & NulSID(Text(23)) & "',model = '" & NulSID(Text(26)) & "',"
7                 sSQL = sSQL & " variant = '" & NulSID(Text(27)) & "',color = '" & NulSID(Text(28)) & "',stnk_number = '" & NulSID(Text(30)) & "',stnk_name = '" & NulSID(Text(31)) & "',stnk_expiry_date = '" & Format(TDBDate2, "yyyy-mm-dd") & "',"
8                 sSQL = sSQL & " stnk_address = '" & NulSID(Text(33)) & "',vehicle_category = '" & NulSID(Combo5.Text) & "', valid_from= '" & Format(TDBDate3, "yyyy-mm-dd") & "',valid_to= '" & Format(TDBDate4, "yyyy-mm-dd") & "',"
9                 sSQL = sSQL & " company_name = '" & NulSID(Text(10)) & "',company_business_sector = '" & NulSID(Text(11)) & "',news_promotion_event_notification_flag = '" & IIf(ckNews.value = vbChecked, 1, 0) & "',"
10                sSQL = sSQL & " preferred_contact_time_start = '" & txtTime(0) & "',preferred_contact_time_end = '" & txtTime(1) & "',npwp_number = '" & NulSID(Text(13)) & "',katashiki_suffix = '" & NulSID(Text(24)) & "',"
11                sSQL = sSQL & " color_code = '" & NulSID(Text(25)) & "',miles_travelled = '" & NulSID(Text(29)) & "',customer_interest_to_tcare = '" & NulSID(Combo6.Text) & "',customer_interest_to_tintouch = '" & NulSID(Combo7.Text) & "',"
12                sSQL = sSQL & " reasons_for_change= '" & NulSID(Combo9.Text) & "',reasons_for_change_others = '" & NulSID(Text(42)) & "' where one_account_ID = '" & NulSID(Text(1)) & "' and vin = '" & NulSID(Text(22)) & "'"
13                Call objO.InsertData(sSQL)
                  
14                sSQL = "insert into ReplBengkel99..kirimDataDidx (tipe,nobookingapps,nobooking,status,proses,event_no,tglentry,KODEDEALER) values"
15                sSQL = sSQL & " ('TOYOTA_ID','" & NulSID(Text(1)) & "','','COMPLETED',0,'' ,getdate(),'" & Right(gDb, 2) & "')"
          
16                Call objO.InsertData(sSQL)
17                MsgBox "Data berhasil disimpan.", vbInformation
18                Call cmdView_Click
19            End If
20        Else
21            If objO.FetchData_Count("select * from ReplBengkel99..p2_ToyotaID where vin = '" & NulSID(Text(22)) & "'") > 0 Then
22                sSQL = "UPDATE ReplBengkel99..p2_ToyotaID SET first_name = '" & NulSID(Text(2)) & "',last_name = '" & NulSID(Text(3)) & "', nick_name = '" & NulSID(Text(4)) & "',phone_number = '" & NulSID(Text(5)) & "',email = '" & NulSID(Text(7)) & "', birth_date = '" & Format(TDBDate1, "yyyy-mm-dd") & "',"
23                sSQL = sSQL & " verification_channel = '" & NulSID(Combo1.Text) & "',ktp_number = '" & NulSID(Text(8)) & "',occupation = '" & NulSID(Text(9)) & "',gender = '" & NulSID(Combo2.Text) & "',address_label = '" & NulSID(Text(14)) & "',residence_address = '" & NulSID(Text(15)) & "',"
24                sSQL = sSQL & " province = '" & NulSID(Text(16)) & "',city = '" & NulSID(Text(17)) & "',district = '" & NulSID(Text(18)) & "',subdistrict = '" & NulSID(Text(19)) & "',postal_code = '" & NulSID(Text(20)) & "',detail_address = '" & NulSID(Text(21)) & "',"
25                sSQL = sSQL & " toyota_single_ID_status = '" & NulSID(Text(12)) & "',customer_category = '" & NulSID(Combo3.Text) & "', primary_user = '" & NulSID(Combo4.Text) & "',police_number = '" & NulSID(Text(23)) & "',model = '" & NulSID(Text(26)) & "',"
26                sSQL = sSQL & " variant = '" & NulSID(Text(27)) & "',color = '" & NulSID(Text(28)) & "',stnk_number = '" & NulSID(Text(30)) & "',stnk_name = '" & NulSID(Text(31)) & "',stnk_expiry_date = '" & Format(TDBDate2, "yyyy-mm-dd") & "',"
27                sSQL = sSQL & " stnk_address = '" & NulSID(Text(33)) & "',vehicle_category = '" & NulSID(Combo5.Text) & "', valid_from= '" & Format(TDBDate3, "yyyy-mm-dd") & "',valid_to= '" & Format(TDBDate4, "yyyy-mm-dd") & "',"
28                sSQL = sSQL & " company_name = '" & NulSID(Text(10)) & "',company_business_sector = '" & NulSID(Text(11)) & "',news_promotion_event_notification_flag = '" & IIf(ckNews.value = vbChecked, 1, 0) & "',"
29                sSQL = sSQL & " preferred_contact_time_start = '" & txtTime(0) & "',preferred_contact_time_end = '" & txtTime(1) & "',npwp_number = '" & NulSID(Text(13)) & "',katashiki_suffix = '" & NulSID(Text(24)) & "',"
30                sSQL = sSQL & " color_code = '" & NulSID(Text(25)) & "',miles_travelled = '" & NulSID(Text(29)) & "',customer_interest_to_tcare = '" & NulSID(Combo6.Text) & "',customer_interest_to_tintouch = '" & NulSID(Combo7.Text) & "',"
31                sSQL = sSQL & " reasons_for_change= '" & NulSID(Combo9.Text) & "',reasons_for_change_others = '" & NulSID(Text(42)) & "' where vin = '" & NulSID(Text(22)) & "'"
                  
32                Call objO.InsertData(sSQL)
33            Else
34                sSQL = "INSERT INTO ReplBengkel99..p2_ToyotaID (event_ID, [timestamp],one_account_ID,dealer_customer_ID,first_name,last_name,phone_number,email,birth_date,verification_channel,ktp_number,occupation,"
35                sSQL = sSQL & " gender,address_label,residence_address,province,city,district,subdistrict,postal_code,detail_address,registration_channel,registration_datetime,consent_given,"
36                sSQL = sSQL & " consent_given_at,consent_given_during,toyota_single_ID_status,customer_category,primary_user,vin,police_number,katashiki_suffix,color_code,model,variant,"
37                sSQL = sSQL & " color,stnk_number,stnk_name,stnk_expiry_date,stnk_address,vehicle_category,first_name_tokenize,last_name_tokenize,phone_number_tokenize,email_tokenize,"
38                sSQL = sSQL & " nick_name,phone_number_secondary,company_name,company_business_sector,news_promotion_event_notification_flag,preferred_contact_time_start,preferred_contact_time_end,"
39                sSQL = sSQL & " npwp_number,npwp_file,ktp_file,is_primary,miles_travelled,customer_interest_to_tcare,customer_interest_to_tintouch,main_vehicle_flag,stnk_file,invoice_file,"
40                sSQL = sSQL & " insurance_end_datetime , insurance_start_datetime, valid_from, valid_to, reasons_for_change, reasons_for_change_others) VALUES "
41                sSQL = sSQL & " (NEWID(), getdate(),'','" & Left(NulSID(Text(2)), 5) & "_" & NulSID(Text(5)) & "','" & NulSID(Text(2)) & "','" & NulSID(Text(3)) & "','" & NulSID(Text(5)) & "','" & NulSID(Text(7)) & "','" & Format(TDBDate1, "yyyy-mm-dd") & "','" & NulSID(Combo1.Text) & "','" & NulSID(Text(8)) & "','" & NulSID(Text(9)) & "',"
42                sSQL = sSQL & " '" & NulSID(Combo2.Text) & "','" & NulSID(Text(14)) & "','" & NulSID(Text(15)) & "','" & NulSID(Text(16)) & "','" & NulSID(Text(17)) & "','" & NulSID(Text(18)) & "','" & NulSID(Text(19)) & "','" & NulSID(Text(20)) & "','" & NulSID(Text(21)) & "','DEALER_SYSTEM',getdate(),'1',"
43                sSQL = sSQL & " getdate(),'ADD_VEHICLE','ACTIVE','" & NulSID(Combo3.Text) & "','" & NulSID(Combo4.Text) & "','" & NulSID(Text(22)) & "','" & NulSID(Text(23)) & "','" & NulSID(Text(24)) & "','" & NulSID(Text(25)) & "','" & NulSID(Text(26)) & "','" & NulSID(Text(27)) & "',"
44                sSQL = sSQL & " '" & NulSID(Text(28)) & "','" & NulSID(Text(30)) & "','" & NulSID(Text(31)) & "','" & Format(TDBDate2, "yyyy-mm-dd") & "','" & NulSID(Text(33)) & "','" & NulSID(Combo5.Text) & "',NULL,NULL,NULL,NULL,"
45                sSQL = sSQL & " NULL,NULL,NULL,NULL,NULL,NULL,NULL,"
46                sSQL = sSQL & " NULL,NULL,NULL,'1',NULL,NULL,NULL,'1',NULL,NULL,"
47                sSQL = sSQL & " NULL , NULL, NULL, NULL, NULL, NULL)"
48                Call objO.InsertData(sSQL)
                  
49                sSQL = "insert into p2_ToyotaID_customer_type (event_ID, one_account_ID, customer_type, vin) VALUES (newid(),'', 'BUYER', '" & NulSID(Text(22)) & "')"

50                Call objO.InsertData(sSQL)
51                sSQL = "insert into p2_ToyotaID_customer_type (event_ID, one_account_ID, customer_type, vin) VALUES (newid(),'', 'OWNER', '" & NulSID(Text(22)) & "')"

52                Call objO.InsertData(sSQL)
53            End If
              
54            sSQL = "insert into ReplBengkel99..kirimDataDidx (tipe,nobookingapps,nobooking, noSPK,status,proses,event_no,tglentry,KODEDEALER) values"
55            sSQL = sSQL & " ('TOYOTAID','" & NulSID(Text(22)) & "','" & NulSID(Text(22)) & "','" & NulSID(Text(22)) & "','SERVICE',0,'' ,getdate(),'" & Right(gDb, 2) & "')"

56            Call objO.InsertData(sSQL)
57            If objO.FetchData_Count("select * from ReplBengkel99..p2_oneAccess where phone_number = '" & NulSID(Text(5)) & "'") = 0 Then
58                sSQL = "INSERT INTO ReplBengkel99..p2_oneAccess (event_ID,[timestamp],one_account_ID,dealer_customer_ID,first_name,last_name,phone_number,email,verification_channel,ktp_number,occupation,gender,registration_channel,"
59                sSQL = sSQL & " registration_datetime,consent_given,consent_given_at,consent_given_during,address_label,residence_address,province,city,district,subdistrict,postal_code,detail_address,"
60                sSQL = sSQL & " first_name_tokenize,last_name_tokenize,email_tokenize,phone_number_tokenize,npwp_number,phone_number_secondary,is_primary,birth_date_char,birth_date) VALUES "
61                sSQL = sSQL & " (NEWID(), getdate(),'','" & Left(NulSID(Text(2)), 5) & "_" & NulSID(Text(5)) & "','" & NulSID(Text(2)) & "','" & NulSID(Text(3)) & "','" & NulSID(Text(5)) & "','" & NulSID(Text(7)) & "','" & NulSID(Combo1.Text) & "','" & NulSID(Text(8)) & "','" & NulSID(Text(9)) & "','" & NulSID(Combo2.Text) & "','DEALER_SYSTEM',"
62                sSQL = sSQL & " getdate(),1,getdate(),'REGISTRATION','" & NulSID(Text(14)) & "','" & NulSID(Text(15)) & "','" & NulSID(Text(16)) & "','" & NulSID(Text(17)) & "','" & NulSID(Text(18)) & "','" & NulSID(Text(19)) & "','" & NulSID(Text(20)) & "','" & NulSID(Text(21)) & "',"
63                sSQL = sSQL & " NULL,NULL,NULL,NULL,NULL,NULL,1,NULL,NULL)"
64                Call objO.InsertData(sSQL)
                  
65                sSQL = "insert into ReplBengkel99..kirimdatadidx (tipe, nobookingapps, nobooking, noSPK, Kodedealer, status, proses, event_no, tglentry) VALUES ('INQUIRY', '" & NulSID(Text(5)) & "', '" & NulSID(Text(5)) & "', '" & NulSID(Text(5)) & "', '" & Right(gDb, 2) & "', 'SERVICE', 0, '', getdate() )"
66                Call objO.InsertData(sSQL)
67            End If
68            Text(0) = NulSID(Text(22))
              
69            MsgBox "Data berhasil disimpan.", vbInformation
70            Call cmdView_Click
                  
71        End If
End Sub

Private Sub refreshData()
1         If NulSID(Text(0)) <> "" Then
2             sSQL = "select * from ReplBengkel99..p2_ToyotaID where vin = '" & NulSID(Text(0)) & "'"
3             Set rsGrid = objO.FetchData(sSQL)
4             If Not rsGrid.EOF Then
5                 rsGrid.MoveFirst
6                 Text(1) = NulSID(rsGrid!one_account_ID)
7                 Text(2) = NulSID(rsGrid!first_name)
8                 Text(3) = NulSID(rsGrid!last_name)
9                 Text(4) = NulSID(rsGrid!nick_name)
10                Text(5) = NulSID(rsGrid!phone_number)
11                Text(6) = NulSID(rsGrid!phone_number_secondary)
12                Text(7) = NulSID(rsGrid!email)
13                TDBDate1.value = Format(NulSID(rsGrid!birth_date), "dd-mm-yyyy")
14                TDBDate2.value = Format(NulSID(rsGrid!stnk_expiry_date), "dd-mm-yyyy")
15                TDBDate3.value = Format(NulSID(rsGrid!valid_from), "dd-mm-yyyy")
16                TDBDate4.value = Format(NulSID(rsGrid!valid_to), "dd-mm-yyyy")
17                If NulSID(rsGrid!verification_channel) <> "" Then
18                    Combo1.Text = NulSID(rsGrid!verification_channel)
19                End If
20                Text(8) = NulSID(rsGrid!ktp_number)
21                Text(9) = NulSID(rsGrid!occupation)
22                If NulSID(rsGrid!gender) <> "" Then
23                    Combo2.Text = NulSID(rsGrid!gender)
24                End If
25                Text(10) = NulSID(rsGrid!company_name)
26                Text(11) = NulSID(rsGrid!company_business_sector)
27                Text(12) = NulSID(rsGrid!toyota_single_ID_status)
28                If NulSID(rsGrid!customer_category) <> "" Then
29                    Combo3.Text = NulSID(rsGrid!customer_category)
30                End If
                  
31                If NulSID(rsGrid!news_promotion_event_notification_flag) = "1" Then
32                    ckNews.value = vbChecked
33                Else
34                    ckNews.value = vbUnchecked
35                End If
36                txtTime(0) = IIf(NulSID(rsGrid!preferred_contact_time_start) <> "", Left(Format(NulSID(rsGrid!preferred_contact_time_start), "hh:mm"), 5), "00:00")
37                txtTime(1) = IIf(NulSID(rsGrid!preferred_contact_time_end) <> "", Left(Format(NulSID(rsGrid!preferred_contact_time_end), "hh:mm"), 5), "00:00")
38                Text(13) = NulSID(rsGrid!npwp_number)
39                Text(14) = NulSID(rsGrid!address_label)
40                Text(15) = NulSID(rsGrid!residence_address)
41                Text(16) = NulSID(rsGrid!province)
42                Text(17) = NulSID(rsGrid!city)
43                Text(18) = NulSID(rsGrid!district)
44                Text(19) = NulSID(rsGrid!subdistrict)
45                Text(20) = NulSID(rsGrid!postal_code)
46                Text(21) = NulSID(rsGrid!detail_address)
47                If NulSID(rsGrid!is_primary) = "1" Then
48                    ckPrimary.value = vbChecked
49                Else
50                    ckPrimary.value = vbUnchecked
51                End If
                  
                  ''-- KENDARAAN
                  
52                Text(22) = NulSID(rsGrid!vin)
53                Text(23) = NulSID(rsGrid!police_number)
54                If NulSID(rsGrid!primary_user) <> "" Then
55                    Combo4.Text = NulSID(rsGrid!primary_user)
56                End If
57                Text(24) = NulSID(rsGrid!katashiki_suffix)
58                Text(25) = NulSID(rsGrid!Color_code)
59                Text(26) = NulSID(rsGrid!model)
60                Text(27) = NulSID(rsGrid!Variant)
61                Text(28) = NulSID(rsGrid!Color)
62                Text(29) = NulSID(rsGrid!miles_travelled)
63                If NulSID(rsGrid!vehicle_category) <> "" Then
64                    Combo5.Text = NulSID(rsGrid!vehicle_category)
65                End If
66                If NulSID(rsGrid!customer_interest_to_tcare) <> "" Then
67                    Combo6.Text = NulSID(rsGrid!customer_interest_to_tcare)
68                End If
69                If NulSID(rsGrid!customer_interest_to_tintouch) <> "" Then
70                    Combo7.Text = NulSID(rsGrid!customer_interest_to_tintouch)
71                End If
72                If NulSID(rsGrid!main_vehicle_flag) = "1" Then
73                    ckMainVhc.value = vbChecked
74                Else
75                    ckPrimary.value = vbUnchecked
76                End If
                  
77                Text(30) = NulSID(rsGrid!stnk_number)
78                Text(31) = NulSID(rsGrid!stnk_name)
79                Text(32) = NulSID(rsGrid!stnk_expiry_date)
80                Text(33) = NulSID(rsGrid!stnk_address)
                  
      '            Text(34) = NulSID(rsGrid!insurance_provider)
      '            Text(35) = NulSID(rsGrid!insurance_provider_other)
      '            Text(36) = NulSID(rsGrid!insurance_policy_number)
      '            If NulSID(rsGrid!insurance_type) <> "" Then
      '                Combo8.Text = NulSID(rsGrid!insurance_type)
      '            End If
81                Text(37) = NulSID(rsGrid!insurance_end_datetime)
82                Text(38) = NulSID(rsGrid!insurance_start_datetime)
83                Text(40) = NulSID(rsGrid!valid_from)
84                Text(41) = NulSID(rsGrid!valid_to)
85                If NulSID(rsGrid!reasons_for_change) <> "" Then
86                    Combo9.Text = NulSID(rsGrid!reasons_for_change)
87                End If
88                Text(42) = NulSID(rsGrid!reasons_for_change_others)
                  ''-- AKHIR KENDARAAN
                  
89                Set rsTemp1 = objO.FetchData("select * FROM ReplBengkel99..p2_ToyotaID_insurance_coverage where one_account_id = '" & NulSID(Text(1)) & "'")
90                TDBGrid1.DataSource = rsTemp1
                  
91                Set rsTemp2 = objO.FetchData("select * FROM ReplBengkel99..p2_ToyotaID_insurance_coverage_policies where one_account_id = '" & NulSID(Text(1)) & "'")
92                TDBGrid4.DataSource = rsTemp2
                  
93                Set rsTemp3 = objO.FetchData("select * FROM ReplBengkel99..p2_ToyotaID_outlet_location where one_account_id = '" & NulSID(Text(1)) & "'")
94                TDBGrid2.DataSource = rsTemp3
                  
95                Set rsTemp4 = objO.FetchData("select * FROM ReplBengkel99..p2_ToyotaID_outlet_preference where one_account_id = '" & NulSID(Text(1)) & "'")
96                TDBGrid3.DataSource = rsTemp4
97            End If
98        End If
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub


Private Sub cmdTambah_Click()
    Call ber_sih
    
    isEdit = False
    
    Text(1).Enabled = True
    Text(5).Enabled = True
    Text(6).Enabled = True
    Text(7).Enabled = True
    Text(22).Enabled = True
    
End Sub

Private Sub cmdView_Click()
    Call ber_sih
    Call refreshData
    
End Sub

Private Sub Form_Activate()
Text(0).SetFocus
End Sub



Private Sub Text_LostFocus(Index As Integer)
    If Index = 22 And isEdit = False Then
        sSQL = "select a.KodeKategori as katashikiSuffix,  b.Kode_Kategory as Model,a.NamaBarang Variant, a.KodeWarna Color_code, c.NamaEnglish Color  "
        sSQL = sSQL & " from SPK00..inmst a WITH (NOLOCK) "
        sSQL = sSQL & " Inner join SPK00..inkategory b WITH (NOLOCK) ON a.Kodekategori = b.Kodekategory"
        sSQL = sSQL & " Inner join SPK00..InWarnaExt c WITH (NOLOCK) on a.KodeWarna = c.KodeWarna"
        sSQL = sSQL & " where a.KodeItem = '" & NulSID(Text(22)) & "'"
        
        Set rsTemp5 = objO.FetchDataHO(sSQL)
        If Not rsTemp5.EOF Then
            rsTemp5.MoveFirst
            Text(24).Text = NulSID(rsTemp5!katashikiSuffix)
            Text(26).Text = NulSID(rsTemp5!model)
            Text(27).Text = NulSID(rsTemp5!Variant)
            Text(25).Text = NulSID(rsTemp5!Color_code)
            Text(28).Text = NulSID(rsTemp5!Color)
            
        End If
    End If
End Sub
