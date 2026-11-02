VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "arprogbar.ocx"
Begin VB.Form frmHOEvent 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Event HO"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   13350
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   9480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   8415
      Left            =   480
      TabIndex        =   19
      Top             =   480
      Width           =   12495
      Begin VB.Frame frmeModel 
         BackColor       =   &H0068ECE6&
         Caption         =   $"frmHOEvent.frx":0000
         Height          =   3615
         Left            =   3240
         TabIndex        =   66
         Top             =   4320
         Visible         =   0   'False
         Width           =   9135
         Begin TrueOleDBGrid70.TDBGrid GridModel 
            Height          =   2895
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   5106
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "GROUP"
            Columns(0).DataField=   "GRPMDL"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "MODEL"
            Columns(1).DataField=   "KODEMDL"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "KETERANGAN"
            Columns(2).DataField=   "NAMAMDL"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   4
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "OK"
            Columns(3).DataField=   "IKUT"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   6876390
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=112,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.locked=-1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
         Begin AIFCmp1.asxPowerButton cmdTutup 
            Height          =   360
            Left            =   120
            TabIndex        =   68
            Top             =   3240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "&Tutup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
      End
      Begin VB.Frame frameKriteria 
         BackColor       =   &H0068ECE6&
         Caption         =   "                                        KRITERIA KHUSUS                                          "
         Height          =   4335
         Left            =   6960
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Reminder MRS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   76
            Top             =   3240
            Width           =   1575
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Program Khusus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   71
            Top             =   2665
            Width           =   1575
         End
         Begin VB.CheckBox chModel 
            BackColor       =   &H0068ECE6&
            Caption         =   "Semua Model"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   70
            Top             =   3570
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Model Kendaraan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   69
            Top             =   3570
            Width           =   1575
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Delivery Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   65
            Top             =   2330
            Width           =   1575
         End
         Begin VB.Frame framBook 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1800
            TabIndex        =   50
            Top             =   360
            Width           =   3135
            Begin VB.OptionButton opBook 
               BackColor       =   &H0068ECE6&
               Caption         =   "Booking"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   53
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton opBook 
               BackColor       =   &H0068ECE6&
               Caption         =   "Non Booking"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   52
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opBook 
               BackColor       =   &H0068ECE6&
               Caption         =   "All"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   51
               Top             =   0
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tipe Order"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tahun Kendaraan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   48
            Top             =   685
            Width           =   1575
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "KM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   1015
            Width           =   1215
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Service Terakhir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   46
            Top             =   1345
            Width           =   1575
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tgl Lahir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   45
            Top             =   1660
            Width           =   1215
         End
         Begin VB.CheckBox chKriteria 
            BackColor       =   &H0068ECE6&
            Caption         =   "Hobby"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   44
            Top             =   2015
            Width           =   1215
         End
         Begin AIFCmp1.asxPowerButton cmdTutup1 
            Height          =   360
            Left            =   0
            TabIndex        =   54
            Top             =   3960
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "&Tutup"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   1800
            TabIndex        =   5
            Top             =   1000
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   3120
            TabIndex        =   6
            Top             =   1000
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   9
            Left            =   1800
            TabIndex        =   3
            Top             =   670
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">9999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   10
            Left            =   3120
            TabIndex        =   4
            Top             =   670
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">9999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   13
            Left            =   1800
            TabIndex        =   7
            Top             =   1330
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99999"
            PromptChar      =   ">"
         End
         Begin MSComCtl2.DTPicker tglLahir1 
            Height          =   255
            Left            =   1800
            TabIndex        =   55
            Top             =   1660
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   115474433
            CurrentDate     =   39853
         End
         Begin MSComCtl2.DTPicker tglLahir2 
            Height          =   255
            Left            =   3600
            TabIndex        =   56
            Top             =   1660
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   115474433
            CurrentDate     =   39853
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   1800
            TabIndex        =   9
            Top             =   2000
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   270
            Left            =   4215
            TabIndex        =   57
            Top             =   2007
            Visible         =   0   'False
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker Ddate1 
            Height          =   255
            Left            =   1800
            TabIndex        =   62
            Top             =   2330
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   115474433
            CurrentDate     =   39853
         End
         Begin MSComCtl2.DTPicker Ddate2 
            Height          =   255
            Left            =   3600
            TabIndex        =   63
            Top             =   2330
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   115474433
            CurrentDate     =   39853
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   6
            Left            =   1800
            TabIndex        =   2
            Top             =   2650
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   270
            Left            =   4215
            TabIndex        =   72
            Top             =   2657
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   7
            Left            =   1800
            TabIndex        =   73
            Top             =   2970
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   3120
            TabIndex        =   8
            Top             =   1330
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">99999"
            PromptChar      =   ">"
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   75
            Top             =   1375
            Width           =   225
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3240
            TabIndex        =   64
            Top             =   2360
            Width           =   225
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   61
            Top             =   1045
            Width           =   225
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   60
            Top             =   715
            Width           =   225
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hari"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4080
            TabIndex        =   59
            Top             =   1375
            Width           =   285
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3255
            TabIndex        =   58
            Top             =   1690
            Width           =   225
         End
      End
      Begin VB.CheckBox chKrti 
         BackColor       =   &H0068ECE6&
         Caption         =   "Kriteria Khusus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   42
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chStatus 
         BackColor       =   &H0068ECE6&
         Caption         =   "Status Aktif"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   40
         Top             =   570
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chAktif 
         BackColor       =   &H0068ECE6&
         Caption         =   "Hanya Display Event yang masih aktif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chAllDealer 
         BackColor       =   &H0068ECE6&
         Caption         =   "Semua Dealer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0068ECE6&
         Caption         =   " Tipe : "
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   6975
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtRumus 
            Height          =   285
            Left            =   2520
            MaxLength       =   200
            TabIndex        =   41
            Top             =   0
            Visible         =   0   'False
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker RangeTgl1 
            Height          =   255
            Left            =   1200
            TabIndex        =   14
            Top             =   630
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   115605505
            CurrentDate     =   39853
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H0068ECE6&
            Caption         =   "Text"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   855
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tanggal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   630
            Width           =   975
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   13
            Top             =   315
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSComCtl2.DTPicker RangeTgl2 
            Height          =   255
            Left            =   3000
            TabIndex        =   15
            Top             =   630
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   115605505
            CurrentDate     =   39853
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   16
            Top             =   960
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   1200
            TabIndex        =   17
            Top             =   1200
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   50
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2715
            TabIndex        =   38
            Top             =   660
            Width           =   225
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kriteria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   37
            Top             =   360
            Width           =   510
         End
         Begin VB.Shape Shape1 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   3000
            Top             =   315
            Width           =   60
         End
      End
      Begin VB.Frame frmeDealer 
         BackColor       =   &H0068ECE6&
         Caption         =   $"frmHOEvent.frx":00B4
         Height          =   3615
         Left            =   1200
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   6615
         Begin TrueOleDBGrid70.TDBGrid GridDealer 
            Height          =   3015
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   5318
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE DEALER"
            Columns(0).DataField=   "KODEDEALER"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NAMA DEALER"
            Columns(1).DataField=   "NAMADEALER"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   4
            Columns(2)._MaxComboItems=   5
            Columns(2).DataField=   "TERLIBAT"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
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
            Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   6876390
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H808080&"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF00&"
            _StyleDefs(21)  =   ":id=9,.borderColor=&H80000001&"
            _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(49)  =   "Named:id=33:Normal"
            _StyleDefs(50)  =   ":id=33,.parent=0"
            _StyleDefs(51)  =   "Named:id=34:Heading"
            _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   ":id=34,.wraptext=-1"
            _StyleDefs(54)  =   "Named:id=35:Footing"
            _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   "Named:id=36:Selected"
            _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(58)  =   "Named:id=37:Caption"
            _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(60)  =   "Named:id=38:HighlightRow"
            _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(62)  =   "Named:id=39:EvenRow"
            _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(64)  =   "Named:id=40:OddRow"
            _StyleDefs(65)  =   ":id=40,.parent=33"
            _StyleDefs(66)  =   "Named:id=41:RecordSelector"
            _StyleDefs(67)  =   ":id=41,.parent=34"
            _StyleDefs(68)  =   "Named:id=42:FilterBar"
            _StyleDefs(69)  =   ":id=42,.parent=33"
         End
         Begin AIFCmp1.asxPowerButton cmdTutupDealer 
            Height          =   360
            Left            =   120
            TabIndex        =   22
            Top             =   3240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Tut&up"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
      End
      Begin MSComCtl2.DTPicker TglBerlaku1 
         Height          =   255
         Left            =   1215
         TabIndex        =   10
         Top             =   1455
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   115671041
         CurrentDate     =   39853
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1215
         TabIndex        =   0
         Top             =   855
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin TrueOleDBGrid70.TDBGrid GridEvent 
         Height          =   6015
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   10610
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   6876390
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
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFF80&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
         _StyleDefs(61)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(62)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(63)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(64)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1215
         TabIndex        =   24
         Top             =   555
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSComCtl2.DTPicker TglBerlaku2 
         Height          =   255
         Left            =   3015
         TabIndex        =   11
         Top             =   1455
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   115998721
         CurrentDate     =   39853
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   1215
         TabIndex        =   1
         Top             =   1155
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   29
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Event"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berlaku "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   495
         TabIndex        =   27
         Top             =   1485
         Width           =   570
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2670
         TabIndex        =   26
         Top             =   1485
         Width           =   225
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Event"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   825
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   5295
         Top             =   870
         Width           =   60
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   8835
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   15584
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "    Daftar Program"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   4194304
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12540
      _ExtentX        =   22119
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   8400
      TabIndex        =   18
      Top             =   9000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Tambah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdUbah 
      Height          =   360
      Left            =   9360
      TabIndex        =   30
      Top             =   9000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Ubah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   10320
      TabIndex        =   31
      Top             =   9000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdKeluar 
      Height          =   360
      Left            =   11280
      TabIndex        =   32
      Top             =   9000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdCetak 
      Height          =   360
      Left            =   7080
      TabIndex        =   33
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Cetak"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton tblDeactivation 
      Height          =   360
      Left            =   480
      TabIndex        =   74
      Top             =   9000
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Deactivation"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   5520
      TabIndex        =   77
      Top             =   9000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   12648447
      Caption         =   "Import Excell"
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
   Begin ARProgBarCtrl.ARProgressBar ProgBar 
      Height          =   375
      Left            =   1680
      TabIndex        =   78
      Top             =   9000
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   661
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
End
Attribute VB_Name = "frmHOEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 18 MEI 2011
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim newRecord As Boolean
Dim rsGrid As New ADODB.Recordset
Dim rsGrid1 As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsDealer As New ADODB.Recordset
Dim rsGridMdl As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns
Dim support As New ClsForm
Dim tipeBY As String

Public Property Let passingTipe(VALUES As String)
    tipeBY = VALUES
End Property

Public Property Let passingFormula(VALUES As String)
    txtRumus = VALUES
End Property

Private Sub createRS()
    
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    With rsDealer
        .Fields.Append "KODEDEALER", adVarChar, 15
        .Fields.Append "NAMADEALER", adVarChar, 25
        .Fields.Append "TERLIBAT", adBoolean
        .Open
    End With
    
    Set GridDealer.DataSource = rsDealer
End Sub

Private Sub hapusForm()
    
    For i = 0 To 14
        Select Case i
            Case 9, 10, 11, 12, 13, 8
                text(i) = 0
            Case 7
            Case Else
                text(i) = ""
        End Select
    Next i
    TglBerlaku1 = Format(isnow, "DD/MM/YYYY")
    TglBerlaku2 = Format(isnow, "DD/MM/YYYY")
    RangeTgl1 = Format(isnow, "DD/MM/YYYY")
    RangeTgl2 = Format(isnow, "DD/MM/YYYY")
    tglLahir1 = Format(isnow, "DD/MM/YYYY")
    tglLahir2 = Format(isnow, "DD/MM/YYYY")
    Ddate1 = Format(isnow, "DD/MM/YYYY")
    Ddate2 = Format(isnow, "DD/MM/YYYY")
    frmeModel.Visible = False
    frmeDealer.Visible = False
    TglMulai = isnow   'Now()
    TglSelesai = isnow 'Now()
    chAllDealer = Unchecked
    chModel = Unchecked
    chPart = Unchecked
    chInv2 = Unchecked
    tglLahir1.Enabled = False
    tglLahir2.Enabled = False
    Ddate1.Enabled = False
    Ddate2.Enabled = False
    
    For i = 0 To 9
        chKriteria(i).value = Unchecked
    Next
    chKrti.value = Unchecked
    
    chAktif.value = Checked
    'If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
End Sub

Private Sub bukaTutup(ids As Integer)
    
    Select Case ids
        Case 0
            For i = 1 To 5
                text(i).Enabled = True
            Next i
            TglBerlaku1.Enabled = True
            TglBerlaku2.Enabled = True
            RangeTgl1.Enabled = True
            RangeTgl2.Enabled = True
            chAllDealer.Enabled = True
            opTipe(0).Enabled = True
            opTipe(1).Enabled = True
            chKrti.Enabled = True
        Case 1
            For i = 1 To 5
                text(i).Enabled = False
            Next i
            TglBerlaku1.Enabled = False
            TglBerlaku2.Enabled = False
            RangeTgl1.Enabled = False
            RangeTgl2.Enabled = False
            chAllDealer.Enabled = False
            opTipe(0).Enabled = False
            opTipe(1).Enabled = False
            chKrti.Enabled = False
            frameKriteria.Visible = False
    End Select
    
End Sub

Private Sub chAktif_Click()
    Call refreshGrid
End Sub

Private Sub chAllDealer_Click()

    If chAllDealer.Enabled Then
        If cmdSimpan.Caption = "&Simpan" Then
            If chAllDealer = Unchecked Then
                frmeDealer.Visible = True
                If rsDealer.State = adStateClosed Then
                    Call isiDealer
                End If
                frmeDealer.Visible = True
                GridDealer.SetFocus
            ElseIf chAllDealer = Checked Then
                frmeDealer.Visible = False
            End If
        End If
    End If
End Sub

Private Sub chKriteria_Click(Index As Integer)
    If chKriteria(Index) = Checked Then
        Select Case Index
            Case 0
                framBook.Enabled = True
                opBook(0).Enabled = True
                opBook(1).Enabled = True
                opBook(2).Enabled = True
            Case 1
                text(9).Enabled = True
                text(10).Enabled = True
            Case 2
                text(11).Enabled = True
                text(12).Enabled = True
            Case 3
                text(13).Enabled = True
                text(8).Enabled = True
            Case 4
                tglLahir1.Enabled = True
                tglLahir2.Enabled = True
            Case 5
                text(14).Enabled = True
            Case 6
                Ddate1.Enabled = True
                Ddate2.Enabled = True
            Case 7
                chModel.Enabled = True
            Case 8
                text(6).Enabled = True
            Case 9
                ' nothing
        End Select
    ElseIf chKriteria(Index) = Unchecked Then
        Select Case Index
            Case 0
                framBook.Enabled = False
                opBook(0).Enabled = False
                opBook(1).Enabled = False
                opBook(2).Enabled = False
                opBook(2).value = True
            Case 1
                text(9).Enabled = False
                text(10).Enabled = False
                text(9) = 0
                text(10) = 0
            Case 2
                text(11).Enabled = False
                text(12).Enabled = False
                text(11) = 0
                text(12) = 0
            Case 3
                text(13).Enabled = False
                text(8).Enabled = False
                text(13) = 0
                text(8) = 0
            Case 4
                tglLahir1.Enabled = False
                tglLahir2.Enabled = False
                tglLahir1 = Format(isnow, "DD/MM/YYYY")
                tglLahir2 = Format(isnow, "DD/MM/YYYY")
            Case 5
                text(14).Enabled = False
                text(14) = ""
            Case 6
                Ddate1.Enabled = False
                Ddate2.Enabled = False
            Case 7
                chModel.Enabled = False
            Case 8
                text(6).Enabled = False
            Case 9
                ' nothing
        End Select
    End If
End Sub

Private Sub chKrti_Click()
    If chKrti.Enabled Then
    If chKrti = Checked Then
        frameKriteria.Visible = True
        
        framBook.Enabled = IIf(chKriteria(0) = Checked, True, False)
        opBook(0).Enabled = IIf(chKriteria(0) = Checked, True, False)
        opBook(1).Enabled = IIf(chKriteria(0) = Checked, True, False)
        opBook(2).Enabled = IIf(chKriteria(0) = Checked, True, False)
        text(9).Enabled = IIf(chKriteria(1) = Checked, True, False)
        text(10).Enabled = IIf(chKriteria(1) = Checked, True, False)
        text(11).Enabled = IIf(chKriteria(2) = Checked, True, False)
        text(12).Enabled = IIf(chKriteria(2) = Checked, True, False)
        text(13).Enabled = IIf(chKriteria(3) = Checked, True, False)
        text(8).Enabled = IIf(chKriteria(3) = Checked, True, False)
        tglLahir1.Enabled = IIf(chKriteria(4) = Checked, True, False)
        tglLahir2.Enabled = IIf(chKriteria(4) = Checked, True, False)
        text(14).Enabled = IIf(chKriteria(5) = Checked, True, False)
        Ddate1.Enabled = IIf(chKriteria(6) = Checked, True, False)
        Ddate2.Enabled = IIf(chKriteria(6) = Checked, True, False)
        chModel.Enabled = IIf(chKriteria(7) = Checked, True, False)
        text(6).Enabled = IIf(chKriteria(8) = Checked, True, False)
    Else
        frameKriteria.Visible = False
    End If
    End If
End Sub

Private Sub chModel_Click()
    If cmdSimpan.Caption = "&Simpan" Then
        If chModel.value = Checked Then
            frmeModel.Visible = False
        ElseIf chModel.value = Unchecked Then
            frmeModel.Visible = True
            Call isiModel
        End If
    End If
End Sub

Private Sub cmdHapus_Click()
    If Trim(text(0)) <> "" Then
        If MsgBox("Yakin akan menghapus Event ?", vbYesNo + vbQuestion, "Hapus Event") = vbYes Then
            If Log.LogHOConn = 1 Then
                If ConHO = True Then
                    Call objO.UpdateDataHO("UPDATE HO_EVENT SET FS_LOG_DEL = '" & Trim(gSA) & "', FS_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "' WHERE NOEVENT = '" & Trim(text(0)) & "'")
                    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('EVENT', '" & Trim(text(0)) & "','D')")
                Else
                    MsgBox "Koneksi HO Terputus !", vbInformation + vbOKOnly
                End If
            End If
        End If
    End If
    Call refreshGrid
End Sub

Private Sub cmdImport_Click()
      Dim ErrBegin As Boolean, konter As Long
1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim Jml As Integer
2         konter = 1
3         progBar.Min = 0
4         CommonDialog1.FileName = ""
5         CommonDialog1.InitDir = App.Path
6         CommonDialog1.filter = "Exel File (*.xls)|*.xls"
7         CommonDialog1.ShowOpen
8         If CommonDialog1.FileName <> "" Then
9            If CnExel.State Then CnExel.Close
10           CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
11           GoSub Proses
12        Else
13          Exit Sub
14        End If
15     Exit Sub
       
Proses:
16    ErrBegin = True
         Dim VarNmFile As String
17       Call objHT.OpenTransactionHO
         Dim detailProgram As String
         Dim idEvent As String
         Dim lastNamaEvent As String
18       If RcExcel.State Then RcExcel.Close
19       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
20       If RcExcel.RecordCount > 0 Then
21          progBar.Visible = True
22          progBar.Max = RcExcel.RecordCount
23          RcExcel.MoveFirst
24          idEvent = ""
25          lastNamaEvent = ""
26          Do Until RcExcel.EOF
27             DoEvents
               
28             progBar.value = konter
29             detailProgram = ""
30             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
31             If lastNamaEvent <> Nul(Trim(RcExcel.Fields(0).value)) Then
32                idEvent = isiNoUrut
33                Call objHT.InsertData_Ho_Only("INSERT INTO TEMPKIRIM VALUES ('EVENT', '" & Trim(idEvent) & "','S')")
                  
34             End If
35             lastNamaEvent = Nul(Trim(RcExcel.Fields(0).value))
36             If Nul(Trim(RcExcel.Fields(1).value)) = "" Then
37                Call objHT.CloseTransactionHO_err
38                progBar.Visible = False
39                MsgBox "Keterangan " & Trim(RcExcel.Fields(1).value) & " baris ke-" & konter & " salah/kosong !", vbCritical, "Error"
40                Exit Sub
41             End If
42             If Nul(Trim(RcExcel.Fields(2).value)) = "" Then
43                Call objHT.CloseTransactionHO_err
44                progBar.Visible = False
45                MsgBox "Tanggal Awal Event " & Trim(RcExcel.Fields(2).value) & " baris ke-" & konter & " salah/kosong !", vbCritical, "Error"
46                Exit Sub
47             End If
48             If Nul(Trim(RcExcel.Fields(3).value)) = "" Then
49                Call objHT.CloseTransactionHO_err
50                progBar.Visible = False
51                MsgBox "Tanggal Awal Event " & Trim(RcExcel.Fields(3).value) & " baris ke-" & konter & " salah/kosong !", vbCritical, "Error"
52                Exit Sub
53             End If
54             If DateDiff("d", Format(Trim(RcExcel.Fields(2).value), "MM/DD/YYYY"), Format(Trim(RcExcel.Fields(3).value), "MM/DD/YYYY")) < 0 Then
55                Call objHT.CloseTransactionHO_err
56                progBar.Visible = False
57                MsgBox "Tanggal range event baris ke-" & konter & " salah !", vbCritical, "Error"
58                Exit Sub
59             End If
60             If Nul(Trim(RcExcel.Fields(4).value)) = "" And Nul(Trim(RcExcel.Fields(5).value)) = "" Then
61                Call objHT.CloseTransactionHO_err
62                progBar.Visible = False
63                MsgBox "Dealer pada baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
64                Exit Sub
65             End If
66             If Nul(Trim(RcExcel.Fields(6).value)) <> "" Then
67                If UCase(Nul(Trim(RcExcel.Fields(6).value))) = "BOOKING" Or UCase(Nul(Trim(RcExcel.Fields(6).value))) = "NON BOOKING" Or UCase(Nul(Trim(RcExcel.Fields(6).value))) = "ALL" Then
68              Else
69                    Call objHT.CloseTransactionHO_err
70                    progBar.Visible = False
71                    MsgBox "Tipe Order pada baris ke-" & konter & " salah format !", vbCritical, "Error"
72                    Exit Sub
73                End If
74             End If
75             If Nul(Trim(RcExcel.Fields(7).value)) <> "" Then
76                If IsNumeric(Nul(Trim(RcExcel.Fields(7).value))) = False Then
77                    Call objHT.CloseTransactionHO_err
78                    progBar.Visible = False
79                    MsgBox "Tahun kendaraan awal pada baris ke-" & konter & " salah format !", vbCritical, "Error"
80                    Exit Sub
81                End If
82             End If
83             If Nul(Trim(RcExcel.Fields(8).value)) <> "" Then
84                If IsNumeric(Nul(Trim(RcExcel.Fields(8).value))) = False Then
85                    Call objHT.CloseTransactionHO_err
86                    progBar.Visible = False
87                    MsgBox "Tahun kendaraan akhir pada baris ke-" & konter & " salah format !", vbCritical, "Error"
88                    Exit Sub
89                End If
90             End If
               
               '-- VALIDASI KM KENDARAAN --
91             If Nul(Trim(RcExcel.Fields(9).value)) <> "" Then
92                If IsNumeric(Nul(Trim(RcExcel.Fields(9).value))) = False Then
93                    Call objHT.CloseTransactionHO_err
94                    progBar.Visible = False
95                    MsgBox "KM kendaraan awal pada baris ke-" & konter & " salah format !", vbCritical, "Error"
96                    Exit Sub
97                End If
98             End If
99             If Nul(Trim(RcExcel.Fields(10).value)) <> "" Then
100               If IsNumeric(Nul(Trim(RcExcel.Fields(10).value))) = False Then
101                   Call objHT.CloseTransactionHO_err
102                   progBar.Visible = False
103                   MsgBox "KM kendaraan akhir pada baris ke-" & konter & " salah format !", vbCritical, "Error"
104                   Exit Sub
105               End If
106            End If
               '-- AKHIR VALIDASI KM KENDARAAN --
               
               '-- VALIDASI SERVICE TERAKHIR --
107            If Nul(Trim(RcExcel.Fields(11).value)) <> "" Then
108               If IsNumeric(Nul(Trim(RcExcel.Fields(11).value))) = False Then
109                   Call objHT.CloseTransactionHO_err
110                   progBar.Visible = False
111                   MsgBox "Service terakhir kendaraan awal pada baris ke-" & konter & " salah format !", vbCritical, "Error"
112                   Exit Sub
113               End If
114            End If
115            If Nul(Trim(RcExcel.Fields(12).value)) <> "" Then
116               If IsNumeric(Nul(Trim(RcExcel.Fields(12).value))) = False Then
117                   Call objHT.CloseTransactionHO_err
118                   progBar.Visible = False
119                   MsgBox "Service terakhir kendaraan akhir pada baris ke-" & konter & " salah format !", vbCritical, "Error"
120                   Exit Sub
121               End If
122            End If
               '-- AKHIR VALIDASI SERVICE TERAKHIR --
               
               '-- VALIDASI DATA TANGGAL LAHIR --
123            If Nul(Trim(RcExcel.Fields(13).value)) <> "" Then
124               If IsDate(Nul(Trim(RcExcel.Fields(13).value))) = False Then
125                   Call objHT.CloseTransactionHO_err
126                   progBar.Visible = False
127                   MsgBox "Tanggal Lahir Awal Event " & Trim(RcExcel.Fields(13).value) & " baris ke-" & konter & " salah/kosong !", vbCritical, "Error"
128                   Exit Sub
129               End If
130            End If
131            If Nul(Trim(RcExcel.Fields(14).value)) <> "" Then
132               If IsDate(Nul(Trim(RcExcel.Fields(14).value))) = False Then
133                   Call objHT.CloseTransactionHO_err
134                   progBar.Visible = False
135                   MsgBox "Tanggal Awal Event " & Trim(RcExcel.Fields(14).value) & " baris ke-" & konter & " salah/kosong !", vbCritical, "Error"
136                   Exit Sub
137               End If
138            End If
139            If Nul(Trim(RcExcel.Fields(13).value)) <> "" And Nul(Trim(RcExcel.Fields(14).value)) <> "" Then
140               If DateDiff("d", Format(Trim(RcExcel.Fields(13).value), "MM/DD/YYYY"), Format(Trim(RcExcel.Fields(14).value), "MM/DD/YYYY")) < 0 Then
141                  Call objHT.CloseTransactionHO_err
142                  progBar.Visible = False
143                  MsgBox "Range tanggal lahir baris ke-" & konter & " salah !", vbCritical, "Error"
144                  Exit Sub
145               End If
146            End If
               '-- AKHIR VALIDASI TANGGAL LAHIR --
               
               '-- VALIDASI DATA DELIV DATE --
147            If Nul(Trim(RcExcel.Fields(16).value)) <> "" Then
148               If IsDate(Nul(Trim(RcExcel.Fields(16).value))) = False Then
149                   Call objHT.CloseTransactionHO_err
150                   progBar.Visible = False
151                   MsgBox "Deliv date Awal " & Trim(RcExcel.Fields(16).value) & " baris ke-" & konter & " salah/kosong !", vbCritical, "Error"
152                   Exit Sub
153               End If
154            End If
155            If Nul(Trim(RcExcel.Fields(17).value)) <> "" Then
156               If IsDate(Nul(Trim(RcExcel.Fields(17).value))) = False Then
157                   Call objHT.CloseTransactionHO_err
158                   progBar.Visible = False
159                   MsgBox "Deliv date akhir " & Trim(RcExcel.Fields(17).value) & " baris ke-" & konter & " salah/kosong !", vbCritical, "Error"
160                   Exit Sub
161               End If
162            End If
163            If Nul(Trim(RcExcel.Fields(16).value)) <> "" And Nul(Trim(RcExcel.Fields(17).value)) <> "" Then
164               If DateDiff("d", Format(Trim(RcExcel.Fields(16).value), "MM/DD/YYYY"), Format(Trim(RcExcel.Fields(17).value), "MM/DD/YYYY")) < 0 Then
165                  Call objHT.CloseTransactionHO_err
166                  progBar.Visible = False
167                  MsgBox "Range deliv date baris ke-" & konter & " salah !", vbCritical, "Error"
168                  Exit Sub
169               End If
170            End If
               '-- AKHIR VALIDASI DELIV DATE --
               
               '-- VALIDASI DATA PROGRAM KHUSUS
171            If Nul(Trim(RcExcel.Fields(18).value)) <> "" Then
172               Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_program A WITH(NOLOCK) LEFT JOIN HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Nul(Trim(RcExcel.Fields(18).value)) & "' and FS_LOG_DEL = ''")
173               If rsTemp.EOF Then
174                   Call objHT.CloseTransactionHO_err
175                   progBar.Visible = False
176                   MsgBox "Kode program baris ke-" & konter & " salah / tidak ditemukan !", vbCritical, "Error"
177                   Exit Sub
178               Else
179                   detailProgram = rsTemp("PROGRAM")
180               End If
181            End If
               
               '-- AKHIR VALIDASI PROGRAM KHUSUS --
               
182            If Nul(Trim(RcExcel.Fields(20).value)) = "" And Nul(Trim(RcExcel.Fields(21).value)) = "" Then
183               Call objHT.CloseTransactionHO_err
184               progBar.Visible = False
185               MsgBox "Model pada baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
186               Exit Sub
187            End If
               
188               If Log.LogHOConn = 1 Then
189                   If ConHO = True Then
                          'Call saveLogMenu(gSA, ComptName, gVersi, "frmHOEvent", "cmdSimpan", 0)
190                       erTranAll = False
191                       onTR = True
                         
192                       Call ClearsSql
                          
193                       gsSql1(0) = "NOEVENT"
194                       gsSql2(0) = idEvent
195                       gsSql1(1) = "NAMAEVENT"
196                       gsSql2(1) = Nul(Trim(RcExcel.Fields(0).value))
197                       gsSql1(2) = "TEXTEVENT"
198                       gsSql2(2) = Nul(Trim(RcExcel.Fields(1).value))
199                       gsSql1(3) = "BERLAKU1"
200                       gsSql2(3) = Format(Nul(Trim(RcExcel.Fields(2).value)), "MM/DD/YYYY")
201                       gsSql1(4) = "BERLAKU2"
202                       gsSql2(4) = Format(Nul(Trim(RcExcel.Fields(3).value)), "MM/DD/YYYY")
203                       gsSql1(5) = "AKTIF"
204                       gsSql2(5) = 1
205                       gsSql1(6) = "DEALERBERLAKU"
206                       gsSql2(6) = Nul(Trim(RcExcel.Fields(4).value))
207                       gsSql1(7) = "FS_LOG_ADD"
208                       gsSql2(7) = gSA
209                       gsSql1(8) = "FS_TGL_ADD"
210                       gsSql2(8) = Format(isnow, "MM/DD/YYYY")
                          
211                       gsSql1(9) = "tipeOrder"
212                       gsSql2(9) = Nul(Trim(RcExcel.Fields(6).value))
213                       gsSql1(10) = "thVehicle1"
214                       gsSql2(10) = Nul(Trim(RcExcel.Fields(7).value))
215                       gsSql1(11) = "thVehicle2"
216                       gsSql2(11) = Nul(Trim(RcExcel.Fields(8).value))
217                       gsSql1(12) = "KM1"
218                       gsSql2(12) = CDbl(Nul(Trim(RcExcel.Fields(9).value)))
219                       gsSql1(13) = "KM2"
220                       gsSql2(13) = CDbl(Nul(Trim(RcExcel.Fields(10).value)))
221                       gsSql1(14) = "lastService1"
222                       gsSql2(14) = Nul(Trim(RcExcel.Fields(11).value))
223                       gsSql1(15) = "lastService2"
224                       gsSql2(15) = Nul(Trim(RcExcel.Fields(12).value))
225                       gsSql1(16) = "birthDay1"
226                       gsSql2(16) = IIf(Nul(Trim(RcExcel.Fields(13).value)) = "", "01/01/1900", Format(Nul(Trim(RcExcel.Fields(13).value)), "MM/DD/YYYY"))
227                       gsSql1(17) = "birthDay2"
228                       gsSql2(17) = IIf(Nul(Trim(RcExcel.Fields(14).value)) = "", "01/01/1900", Format(Nul(Trim(RcExcel.Fields(14).value)), "MM/DD/YYYY"))
229                       gsSql1(18) = "Hobby"
230                       gsSql2(18) = Nul(Trim(RcExcel.Fields(15).value))
231                       gsSql1(19) = "DelivDay1"
232                       gsSql2(19) = IIf(Nul(Trim(RcExcel.Fields(16).value)) = "", "01/01/1900", Format(Nul(Trim(RcExcel.Fields(16).value)), "MM/DD/YYYY"))
233                       gsSql1(20) = "DelivDay2"
234                       gsSql2(20) = IIf(Nul(Trim(RcExcel.Fields(17).value)) = "", "01/01/1900", Format(Nul(Trim(RcExcel.Fields(17).value)), "MM/DD/YYYY"))
235                       gsSql1(21) = "modelVehichle"
236                       gsSql2(21) = Nul(Trim(RcExcel.Fields(20).value))
237                       gsSql1(22) = "KodeProgram"
238                       gsSql2(22) = Nul(Trim(RcExcel.Fields(18).value))
239                       gsSql1(23) = "Direminder"
240                       gsSql2(23) = IIf(Nul(Trim(RcExcel.Fields(19).value)) = "", "0", "1")
                  
241                       If rsTemp.State Then rsTemp.Close
242                       Set rsTemp = support.searchEvent(idEvent)   'oConnHO.Execute("SELECT * FROM HO_EVENT WHERE NOEVENT = '" & Text(0) & "' and FS_LOG_DEL = ''")
243                       If rsTemp.EOF Then
244                           sSQL = GensSql2("HO_EVENT", False)
245                           Call objHT.UpdateData_Ho_Only(sSQL)
246                       End If
                         
247                       If rsTemp.State <> adStateClosed Then rsTemp.Close
                          
                          ' SIMPAN MODEL
248                       If UCase(Nul(Trim(RcExcel.Fields(20).value))) = "V" Then
249                           Set rsTemp = objO.FetchDataHO("SELECT * FROM MODEL WITH (NOLOCK) WHERE MDL LIKE '%" & Nul(Trim(RcExcel.Fields(21).value)) & "%'")
250                           If Not rsTemp.EOF Then
251                               rsTemp.MoveFirst
252                               Call objHT.DeleteData_Ho_Only("DELETE HO_EVTMODEL WHERE NOEVENT = '" & idEvent & "' AND GRPMDL = '" & rsTemp("MDL") & "'")
253                               Call objHT.InsertData_Ho_Only("INSERT INTO HO_EVTMODEL (noEvent,GRPMDL,MDL,[DESC]) values ('" & idEvent & "','" & rsTemp("MDL") & "','" & rsTemp("MODEL") & "','" & rsTemp("DESC") & "')")
254                           End If
255                       End If
                      
                          ' SIMPAN DEALER
256                       If UCase(Nul(Trim(RcExcel.Fields(4).value))) = "V" Then
257                           Set rsTemp = objO.FetchDataHO("select * from t_dealer where internal = 1 and fs_kd_dealer = '%" & Nul(Trim(RcExcel.Fields(5).value)) & "%'")
258                           If Not rsTemp.EOF Then
259                               rsTemp.MoveFirst
260                               Call objHT.DeleteData_Ho_Only("DELETE HO_EVTDEALER WHERE NOEVENT = '" & idEvent & "' AND kddealer = '" & rsTemp("fs_kd_dealer") & "'")
261                               Call objHT.InsertData_Ho_Only("INSERT INTO HO_EVTDEALER (noEvent,kdDealer,namaDealer) values ('" & idEvent & "','" & rsTemp("fs_kd_dealer") & "','" & rsTemp("fs_nm_dealer") & "')")
262                           End If
263                       End If
                   
264                   End If
265               End If
266            RcExcel.MoveNext
267            konter = konter + 1
268         Loop
269      End If
270      RcExcel.Close
271      CnExel.Close
272      Call objHT.CloseTransactionHO
273      MsgBox "Proses Import telah selesai " & konter - 1 & " data!", vbInformation
274      Call refreshGrid
275      progBar.Visible = False
      '   GridAFI.Refresh
276      Return
277      Exit Sub
vb_error:
278       Call objHT.CloseTransactionHO_err
          
279      If ErrBegin = True Then
           ' DbCon.RollbackTrans
280      End If
281      progBar.Visible = False
282      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")
End Sub

Private Sub CmdKeluar_Click()
    If rsGrid.State <> adStateClosed Then rsGrid.Close
    If rsGrid1.State <> adStateClosed Then rsGrid1.Close
    If rsTemp.State <> adStateClosed Then rsTemp.Close
    If rsDealer.State <> adStateClosed Then rsDealer.Close
    If rsGridMdl.State <> adStateClosed Then rsGridMdl.Close
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
          
1         On Error GoTo vbError
2         If cmdSimpan.Caption = "&Tambah" Then
3             Call hapusForm
4             Call bukaTutup(0)
5             cmdSimpan.Caption = "&Simpan"
6             cmdUbah.Caption = "&Batal"
7             newRecord = True
8         ElseIf cmdSimpan.Caption = "&Simpan" Then
9             boleh = True
10            pesan = "Lengkapi Data Berikut "
              
11            If text(1) = "" Then
12                boleh = False
13                pesan = pesan + "Nama Event. " & vbCr
14            End If
15            If text(2) = "" Then
16                boleh = False
17                pesan = pesan + "Keterangan. " & vbCr
18            End If

19            For i = 0 To 9
20                Select Case i
                      Case 1
21                        If chKriteria(i).value = Checked Then
22                            If Trim(text(9)) = "" Or Trim(text(10)) = "" Or Trim(text(9)) = "0" Or Trim(text(10)) = "0" Or Trim(text(9)) > Trim(text(10)) Then
23                                boleh = False
24                                pesan = pesan + "Kriteria Tahun Kendaraan. " & vbCr
25                            End If
26                        End If
27                    Case 2
28                        If chKriteria(i).value = Checked Then
29                            If Trim(text(11)) = "" Or Trim(text(12)) = "" Or Trim(text(11)) = "0" Or Trim(text(12)) = "0" Or Trim(text(11)) > Trim(text(12)) Then
30                                boleh = False
31                                pesan = pesan + "Kriteria KM Kendaraan. " & vbCr
32                            End If
33                        End If
34                    Case 3
35                        If chKriteria(i).value = Checked Then
36                            If Trim(text(13)) = "" Or Trim(text(13)) = "0" Or Trim(text(8)) = "" Or Trim(text(8)) = "0" Then
37                                boleh = False
38                                pesan = pesan + "Kriteria Service Terakhir. " & vbCr
39                            End If
40                        End If
41                    Case 4
42                        If chKriteria(i).value = Checked Then
43                            If tglLahir1 = tglLahir2 Or (tglLahir1 = Format(isnow, "mm/dd/yyyy") And tglLahir2 = Format(isnow, "mm/dd/yyyy")) Or (tglLahir1 > tglLahir2) Then
44                                boleh = False
45                                pesan = pesan + "Kriteria Tanggal Lahir Customer. " & vbCr
46                            End If
47                        End If
48                    Case 5
49                        If chKriteria(i).value = Checked Then
50                            If Trim(text(14)) = "" Then
51                                boleh = False
52                                pesan = pesan + "Kriteria hobby Customer. " & vbCr
53                            End If
54                        End If
55                    Case 6
56                        If chKriteria(i).value = Checked Then
57                            If Ddate1 = Ddate2 Or (Ddate1 = Format(isnow, "mm/dd/yyyy") And Ddate2 = Format(isnow, "mm/dd/yyyy")) Then
58                                boleh = False
59                                pesan = pesan + "Kriteria Tanggal Delivery. " & vbCr
60                            End If
61                        End If
62                    Case 8
63                        If chKriteria(i).value = Checked Then
64                            If Trim(text(6)) = "" Then
65                                boleh = False
66                                pesan = pesan + "Kode Program. " & vbCr
67                            End If
68                        End If
69                End Select
70            Next i
              
71            If boleh Then
72                If Log.LogHOConn = 1 Then
73                    If ConHO = True Or Log.RPLHistori = 1 Then
74                        Call saveLogMenu(gSA, ComptName, gVersi, "frmHOEvent", "cmdSimpan", 0)

75                        erTranAll = False
76                        onTR = True
77                        Call objHT.OpenTransactionHO
                  
78                        Call ClearsSql
79                        If Trim(text(0)) = "" Then
80                            text(0) = isiNoUrut
81                        End If
82                        gsSql1(0) = "NOEVENT"
83                        gsSql2(0) = text(0)
84                        gsSql1(1) = "NAMAEVENT"
85                        gsSql2(1) = text(1)
86                        gsSql1(2) = "TEXTEVENT"
87                        gsSql2(2) = text(2)
88                        gsSql1(3) = "BERLAKU1"
89                        gsSql2(3) = Format(TglBerlaku1, "MM/DD/YYYY")
90                        gsSql1(4) = "BERLAKU2"
91                        gsSql2(4) = Format(TglBerlaku2, "MM/DD/YYYY")
92                        gsSql1(5) = "AKTIF"
93                        gsSql2(5) = IIf(chStatus.value, 1, 0)
94                        gsSql1(6) = "DEALERBERLAKU"
95                        gsSql2(6) = IIf(chAllDealer.value, "A", "V")
96                        gsSql1(7) = "FS_LOG_ADD"
97                        gsSql2(7) = gSA
98                        gsSql1(8) = "FS_TGL_ADD"
99                        gsSql2(8) = Format(isnow, "MM/DD/YYYY")
                          
100                       gsSql1(9) = "tipeOrder"
101                       gsSql2(9) = IIf(chKriteria(0).value = Unchecked, "", IIf(opBook(0), "Booking", IIf(opBook(1), "Non Booking", "All")))
102                       gsSql1(10) = "thVehicle1"
103                       gsSql2(10) = Trim(text(9))
104                       gsSql1(11) = "thVehicle2"
105                       gsSql2(11) = Trim(text(10))
106                       gsSql1(12) = "KM1"
107                       gsSql2(12) = CDbl(text(11))
108                       gsSql1(13) = "KM2"
109                       gsSql2(13) = CDbl(text(12))
110                       gsSql1(14) = "lastService1"
111                       gsSql2(14) = Trim(text(13))
112                       gsSql1(15) = "lastService2"
113                       gsSql2(15) = Trim(text(8))
114                       gsSql1(16) = "birthDay1"
115                       gsSql2(16) = IIf(chKriteria(4).value = Unchecked, "01/01/1900", Format(tglLahir1, "MM/DD/YYYY"))
116                       gsSql1(17) = "birthDay2"
117                       gsSql2(17) = IIf(chKriteria(4).value = Unchecked, "01/01/1900", Format(tglLahir2, "MM/DD/YYYY"))
118                       gsSql1(18) = "Hobby"
119                       gsSql2(18) = Trim(text(14))
120                       gsSql1(19) = "DelivDay1"
121                       gsSql2(19) = IIf(chKriteria(6).value = Unchecked, "01/01/1900", Format(Ddate1, "MM/DD/YYYY"))
122                       gsSql1(20) = "DelivDay2"
123                       gsSql2(20) = IIf(chKriteria(6).value = Unchecked, "01/01/1900", Format(Ddate2, "MM/DD/YYYY"))
124                       gsSql1(21) = "modelVehichle"
125                       gsSql2(21) = IIf(chKriteria(7).value = Unchecked, "", IIf(chModel.value = Checked, "A", "V"))
126                       gsSql1(22) = "KodeProgram"
127                       gsSql2(22) = Trim(text(6))
128                       gsSql1(23) = "Direminder"
129                       gsSql2(23) = IIf(chKriteria(9).value = Unchecked, "0", "1")
                  
                  
130                       If rsTemp.State Then rsTemp.Close
131                       Set rsTemp = support.searchEvent(Trim(text(0)))   'oConnHO.Execute("SELECT * FROM HO_EVENT WHERE NOEVENT = '" & Text(0) & "' and FS_LOG_DEL = ''")
132                       If Not rsTemp.EOF Then
                              ' DATA SUDAH ADA
133                           If newRecord Then
134                               MsgBox "DATA SUDAH ADA. Jika ingin merubah pilih UBAH", vbCritical, "DUPLIKASI DATA"
135                           Else
136                               temp = MsgBox("Yakin akan merubah Data ?", vbYesNo, "DUPLIKASI DATA")
137                               If temp = vbYes Then
138                                   sSQL = GensSql2("HO_EVENT", True) + " where NOEVENT = '" & Trim(text(0)) & "' and FS_LOG_DEL = ''"
139                               End If
140                           End If
141                       Else
                              ' DATA BARU
142                           sSQL = GensSql2("HO_EVENT", False)
143                       End If
144                       Call objHT.UpdateData_Ho_Only(sSQL)
145                       If rsTemp.State <> adStateClosed Then rsTemp.Close
                          
                          ' SIMPAN MODEL
146                       If chKriteria(7).value = Checked Then
147                           If rsGridMdl.State Then
148                               rsGridMdl.MoveFirst
149                               Call objHT.DeleteData_Ho_Only("DELETE HO_EVTMODEL WHERE NOEVENT = '" & text(0) & "'")
150                               Do While Not rsGridMdl.EOF
151                                   If rsGridMdl("IKUT") Then
152                                       Call objHT.InsertData_Ho_Only("INSERT INTO HO_EVTMODEL (noEvent,GRPMDL,MDL,[DESC]) values ('" & text(0) & "','" & rsGridMdl("GRPMDL") & "','" & rsGridMdl("KODEMDL") & "','" & rsGridMdl("NAMAMDL") & "')")
153                                   End If
154                                   rsGridMdl.MoveNext
155                               Loop
156                           End If
157                       End If
                      
                          ' SIMPAN DEALER
158                       If chAllDealer.value Then
159                           Set rsTemp = objHT.FetchDataHO("SELECT fs_kd_dealer, fs_nm_dealer, ISNULL(NOEVENT,'') AS KODE FROM t_dealer a WITH(NOLOCK) LEFT JOIN HO_EVTDEALER b WITH(NOLOCK) ON a.fs_kd_dealer = b.KDDEALER and A.FS_LOG_DEL = '' AND b.NOEVENT = '" & text(0) & "' " _
                                        & "WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
160                           If Not rsTemp.EOF Then
161                               rsTemp.MoveFirst
162                               Call objHT.DeleteData_Ho_Only("DELETE HO_EVTDEALER WHERE NOEVENT = '" & text(0) & "'")
163                               Do While Not rsTemp.EOF
164                                    Call objHT.InsertData_Ho_Only("INSERT INTO HO_EVTDEALER (noEvent,kdDealer,namaDealer) VALUES ('" & text(0) & "', '" & rsTemp("fs_kd_dealer") & "', '" & rsTemp("fs_nm_dealer") & "')")
                                       
165                                   rsTemp.MoveNext
166                               Loop
167                           End If
168                       Else
169                           If rsDealer.State Then
170                               rsDealer.MoveFirst
171                               Call objHT.DeleteData_Ho_Only("DELETE HO_EVTDEALER WHERE NOEVENT = '" & text(0) & "'")
172                               Do While Not rsDealer.EOF
173                                   If rsDealer("TERLIBAT") Then
174                                       Call objHT.InsertData_Ho_Only("INSERT INTO HO_EVTDEALER (noEvent,kdDealer,namaDealer) VALUES ('" & text(0) & "', '" & rsDealer("KODEDEALER") & "', '" & rsDealer("NAMADEALER") & "')")
175                                   End If
176                                   rsDealer.MoveNext
177                               Loop
178                           End If
179                       End If
                          
                          
                          
                          
180                       If Trim(text(0)) <> "" Then
181                           Call objHT.InsertData_Ho_Only("INSERT INTO TEMPKIRIM VALUES ('EVENT', '" & Trim(text(0)) & "','S')")
182                       End If
183                   End If
184               End If
                  
185               If erTranAll = False Then
186                   Call objHT.CloseTransactionHO
187                   onTR = False
188                   Call saveLogMenu(gSA, ComptName, gVersi, "frmHOEvent", "cmdSimpan", 1)
189                   MsgBox "Proses berhasil disimpan.", vbInformation, "Berhasil Simpan"
                      'Exit Sub
190               Else
191                   Call objHT.CloseTransactionHO_err
192                   onTR = False
193                   Call saveLogMenu(gSA, ComptName, gVersi, "frmOrder", "tblSimpan", 2)
194                   MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
195                   Exit Sub
196               End If
                  
197               Call bukaTutup(1)
198               Call hapusForm
199               Call refreshGrid
                  
200               cmdSimpan.Caption = "&Tambah"
201               cmdUbah.Caption = "&Ubah"
202           Else
203               MsgBox pesan, vbCritical, "KELENGKAPAN"
204               Exit Sub
205           End If
206       End If
207       Exit Sub

vbError:
208       Call objHT.CloseTransactionHO_err
209       onTR = False
210       Call saveLogMenu(gSA, ComptName, gVersi, "frmOrder", "tblSimpan", 2)
211       MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
212       Exit Sub
End Sub

Private Sub cmdTutup_Click()
    frmeModel.Visible = False
End Sub

Private Sub cmdTutup1_Click()
    frameKriteria.Visible = False
End Sub

Private Sub cmdTutupDealer_Click()
    frmeDealer.Visible = False
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call bukaTutup(0)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call bukaTutup(1)
        Call hapusForm
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
        Call refreshGrid
    End If
End Sub

Private Sub Form_Load()
    Call hapusForm
    Call bukaTutup(1)
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridEvent.Width = Frame1.Width - 200
    GridEvent.Height = Frame1.Height - chAllDealer.Top - chAllDealer.Height - 300
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    progBar.Top = cmdKeluar.Top
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    cmdImport.Left = cmdCetak.Left - cmdImport.Width - 100
    progBar.Left = cmdImport.Left - progBar.Width - 100
    
    tblDeactivation.Left = Frame1.Left
    tblDeactivation.Top = Frame1.Top + Frame1.Height + 100
    
End Sub

Private Sub GridEvent_FilterChange()
    On Error GoTo errHandler
    Set cols = GridEvent.Columns
    Dim C As Integer
    
    C = GridEvent.col
    GridEvent.HoldFields
    rsGrid.filter = getFilter()
    GridEvent.col = C
    GridEvent.EditActive = True
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

Private Sub GridEvent_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsGrid1.State <> adStateClosed Then
        If Not rsGrid1.EOF Then
            text(0) = Trim(rsGrid1.Fields(0))
            text(1) = Trim(rsGrid1.Fields(1))
            text(2) = Trim(rsGrid1.Fields(2))
            For i = 0 To 9
                Select Case i
                    Case 0
                        chKriteria(i).value = IIf(Trim(rsGrid1.Fields(3)) = "", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            If Trim(rsGrid1.Fields(3)) = "Booking" Then
                                opBook(0).value = True
                            ElseIf Trim(rsGrid1.Fields(3)) = "Non Booking" Then
                                opBook(1).value = True
                            Else
                                opBook(2).value = True
                            End If
                        End If
                    Case 1
                        chKriteria(i).value = IIf(Trim(rsGrid1.Fields(4)) = "0" Or Trim(rsGrid1.Fields(5)) = "0", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            text(9) = CDbl(rsGrid1.Fields(4))
                            text(10) = CDbl(rsGrid1.Fields(5))
                        End If
                    Case 2
                        chKriteria(i).value = IIf(Trim(rsGrid1.Fields(6)) = "0" Or Trim(rsGrid1.Fields(7)) = "0", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            text(11) = CDbl(rsGrid1.Fields(6))
                            text(12) = CDbl(rsGrid1.Fields(7))
                        End If
                    Case 3
                        chKriteria(i).value = IIf(Trim(rsGrid1.Fields(8)) = "0" Or Trim(rsGrid1.Fields(9)) = "0", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            text(13) = CDbl(rsGrid1.Fields(8))
                            text(8) = CDbl(rsGrid1.Fields(9))
                        End If
                    Case 4
                        chKriteria(i).value = IIf(Format(rsGrid1.Fields(10), "MM/DD/YYYY") = "01/01/1900" Or Format(rsGrid1.Fields(11), "MM/DD/YYYY") = "01/01/1900", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            tglLahir1 = Format(rsGrid1.Fields(10), "MM/DD/YYYY")
                            tglLahir2 = Format(rsGrid1.Fields(11), "MM/DD/YYYY")
                        End If
                    Case 5
                        chKriteria(i).value = IIf(Trim(rsGrid1.Fields(12)) = "", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            text(14) = Trim(rsGrid1.Fields(12))
                        End If
                    Case 6
                        chKriteria(i).value = IIf(Format(rsGrid1.Fields(13), "MM/DD/YYYY") = "01/01/1900" Or Format(rsGrid1.Fields(14), "MM/DD/YYYY") = "01/01/1900", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            Ddate1 = Format(rsGrid1.Fields(13), "MM/DD/YYYY")
                            Ddate2 = Format(rsGrid1.Fields(14), "MM/DD/YYYY")
                        End If
                    Case 7
                        chKriteria(i).value = IIf(Trim(rsGrid1.Fields(15)) = "", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            If Trim(rsGrid1.Fields(15)) = "A" Then
                                chModel.value = Checked
                            End If
                        End If
                    Case 8
                        chKriteria(i).value = IIf(Trim(rsGrid1.Fields(20)) = "", Unchecked, Checked)
                        If chKriteria(i).value = Checked Then
                            text(6) = Trim(rsGrid1.Fields(19))
                        End If
                    Case 9
                        chKriteria(i).value = IIf(Trim(rsGrid1.Fields(21)) = "0", Unchecked, Checked)
                End Select
                If chKriteria(i).value = Checked Then
                    chKrti.value = Checked
                End If
            Next
            
            chAllDealer = IIf(Trim(rsGrid1.Fields(16)) = "V", Unchecked, Checked)
            TglBerlaku1 = Format(rsGrid1.Fields(17), "DD/MM/YYYY")
            TglBerlaku2 = Format(rsGrid1.Fields(18), "DD/MM/YYYY")
            chStatus.value = IIf(Trim(rsGrid1.Fields(19)) = "Non Aktif", Unchecked, Checked)
            If chStatus.value = Checked And Trim(text(0)) <> "" Then
                tblDeactivation.Visible = True
            End If
        End If
    End If
End Sub

Private Sub opTipe_GotFocus(Index As Integer)
    If opTipe(Index) Then
        Select Case Index
            Case 0
                RangeTgl1.Enabled = True
                RangeTgl2.Enabled = True
                text(3).Enabled = False
                text(4).Enabled = False
            Case 1
                RangeTgl1.Enabled = False
                RangeTgl2.Enabled = False
                text(3).Enabled = True
                text(4).Enabled = True
        End Select
    End If
End Sub

Private Sub opTipe_Validate(Index As Integer, Cancel As Boolean)
    If opTipe(Index) Then
        Select Case Index
            Case 0
                RangeTgl1.Enabled = True
                RangeTgl2.Enabled = True
                text(3).Enabled = False
                text(4).Enabled = False
            Case 1
                RangeTgl1.Enabled = False
                RangeTgl2.Enabled = False
                text(3).Enabled = True
                text(4).Enabled = True
        End Select
    End If
End Sub

Private Sub tblDeactivation_Click()
    If MsgBox("Me-Non aktif-kan Program?", vbYesNo + vbDefaultButton1, "Deactivation") = vbYes Then
        Call objO.UpdateDataHO("UPDATE HO_EVENT SET fs_Log_DeAktifasi = '" & Trim(gSA) & "', fs_Tgl_DeAktifasi = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', AKTIF = 0 WHERE NOEVENT = '" & Trim(text(0)) & "'")
        Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('EVENT', '" & Trim(text(0)) & "','S')")
        Call refreshGrid
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
        
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 1
                gMainKeyCari = text(Index)
                gSearchOption = "EVENT"
                frmCariHO.Show (vbModal)
            Case 5
                gMainKeyCari = text(Index)
                gSearchOption = "KRITERIA"
                frmCariHO.Show (vbModal)
            Case 6
                ' PROGRAM
                gMainKeyCari = text(Index)
                gSearchOption = "PROGRAM"
                frmCariHO.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
        If rsTmp.State Then rsTmp.Close
        Set rsTmp = objO.FetchDataHO("SELECT fs_kd_dealer, fs_nm_dealer, ISNULL(NOEVENT,'') AS KODE FROM t_dealer a WITH(NOLOCK) LEFT JOIN HO_EVTDEALER b WITH(NOLOCK) ON a.fs_kd_dealer = b.KDDEALER and A.FS_LOG_DEL = '' AND b.NOEVENT = '" & text(0) & "' " _
                                  & "WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
        If Not rsTmp.EOF Then
            Call createRS
            Do While Not rsTmp.EOF
                rsDealer.AddNew
                rsDealer!KodeDealer = rsTmp("FS_KD_DEALER")
                rsDealer!namaDealer = rsTmp("FS_NM_DEALER")
                If rsTmp("KODE") <> "" Then
                    rsDealer!terlibat = True
                Else
                    rsDealer!terlibat = False
                End If
                rsDealer.Update
                rsTmp.MoveNext
            Loop
            
            GridDealer.DataSource = rsDealer
            GridDealer.EvenRowStyle.BackColor = vbYellow
            GridDealer.OddRowStyle.BackColor = vbWhite
            GridDealer.AlternatingRowStyle = True
        End If
End Sub

Private Sub refreshGrid()
    GridEvent.ClearFields
        If rsGrid1.State Then rsGrid1.Close
        Set rsGrid1 = support.displayEvent(IIf(chAktif.value = Checked, 1, 0))
        If Not rsGrid1.EOF Then
            GridEvent.DataSource = rsGrid1
            Call HeaderColom
        End If
End Sub

Private Sub HeaderColom()
    With GridEvent
        .Columns(0).Caption = "No"
        .Columns(0).Width = 1200
        .Columns(0).DataField = "NO"
        .Columns(0).Locked = True
        .Columns(1).Caption = "Event"
        .Columns(1).Width = 1500
        .Columns(1).DataField = "EVENT"
        .Columns(1).Locked = True
        .Columns(2).Caption = "Keterangan"
        .Columns(2).Width = 3000
        .Columns(2).DataField = "KETERANGAN"
        .Columns(2).Locked = True
        .Columns(3).Caption = "Tipe Booking"
        .Columns(3).Width = 1000
        .Columns(3).DataField = "Tipe_Booking"
        .Columns(3).Locked = True
        .Columns(4).Caption = "Th Mulai"
        .Columns(4).Width = 800
        .Columns(4).DataField = "Tahun_Mulai"
        .Columns(4).Locked = True
        .Columns(5).Caption = "Th Akhir"
        .Columns(5).Width = 800
        .Columns(5).DataField = "Tahun_Akhir"
        .Columns(5).Locked = True
        .Columns(6).Caption = "KM Mulai"
        .Columns(6).Width = 900
        .Columns(6).DataField = "KM_Mulai"
        .Columns(6).Locked = True
        .Columns(7).Caption = "KM Akhir"
        .Columns(7).Width = 900
        .Columns(7).DataField = "KM_Akhir"
        .Columns(7).Locked = True
        .Columns(8).Caption = "Service Terakhir"
        .Columns(8).Width = 1000
        .Columns(8).DataField = "Lastservice1"
        .Columns(8).Locked = True
        .Columns(9).Caption = "Service Terakhir"
        .Columns(9).Width = 1000
        .Columns(9).DataField = "Lastservice2"
        .Columns(9).Locked = True
        .Columns(10).Caption = "Ultah Mulai"
        .Columns(10).Width = 1500
        .Columns(10).DataField = "Ultah_Mulai"
        .Columns(10).Locked = True
        .Columns(11).Caption = "Ultah Akhir"
        .Columns(11).Width = 1500
        .Columns(11).DataField = "Ultah_Selesai"
        .Columns(11).Locked = True
        .Columns(12).Caption = "Hobby"
        .Columns(12).Width = 2000
        .Columns(12).DataField = "Hobby"
        .Columns(12).Locked = True
        .Columns(13).Caption = "Delivery Mulai"
        .Columns(13).Width = 1500
        .Columns(13).DataField = "Delivery_Mulai"
        .Columns(13).Locked = True
        .Columns(14).Caption = "Delivery Akhir"
        .Columns(14).Width = 1500
        .Columns(14).DataField = "Delivery_Akhir"
        .Columns(14).Locked = True
        .Columns(15).Caption = "Model"
        .Columns(15).Width = 800
        .Columns(15).DataField = "Model"
        .Columns(15).Locked = True
        .Columns(16).Caption = "Dealer"
        .Columns(16).Width = 800
        .Columns(16).DataField = "Dealer"
        .Columns(16).Locked = True
        .Columns(17).Caption = "Mulai Berlaku"
        .Columns(17).Width = 1500
        .Columns(17).DataField = "Tanggal_Mulai"
        .Columns(17).Locked = True
        .Columns(18).Caption = "Akhir Berlaku"
        .Columns(18).Width = 1500
        .Columns(18).DataField = "Tanggal_Selesai"
        .Columns(18).Locked = True
        .Columns(19).Caption = "Status"
        .Columns(19).Width = 800
        .Columns(19).DataField = "Status_Event"
        .Columns(19).Locked = True
        .Columns(20).Caption = "Add"
        .Columns(20).Width = 800
        .Columns(20).DataField = "User_Add"
        .Columns(20).Locked = True
        .Columns(21).Caption = "Deactivation"
        .Columns(21).Width = 800
        .Columns(21).DataField = "Deactivation"
        .Columns(21).Locked = True
        .Columns(22).Caption = "Kode Program"
        .Columns(22).Width = 1200
        .Columns(22).DataField = "Kode"
        .Columns(22).Locked = True
        .Columns(23).Caption = "Reminder"
        .Columns(23).Width = 1000
        .Columns(23).DataField = "direminder"
        .Columns(23).Locked = True
    End With
End Sub

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 3500
    GridDealer.Columns(2).Width = 1000
    GridDealer.Columns(2).Alignment = dbgCenter
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As New ADODB.Recordset
    
     
    Select Case Index
        Case 1
            If text(Index) <> "" Then
                    Set rsSearch = objO.FetchDataHO("Select * from HO_EVENT WITH(NOLOCK) where NAMAEVENT = '" & Trim(text(Index)) & "' and FS_LOG_DEL = ''")
                    If Not rsSearch.EOF Then
                        text(0) = Trim(rsSearch("NOEVENT"))
                        text(2) = Trim(rsSearch("TEXTEVENT"))
                        TglBerlaku1 = Format(rsSearch("BERLAKU1"), "DD/MM/YYYY")
                        TglBerlaku2 = Format(rsSearch("BERLAKU2"), "DD/MM/YYYY")
                        
                        For i = 0 To 9
                            Select Case i
                                Case 0
                                    chKriteria(i).value = IIf(Trim(rsSearch("tipeOrder")) = "", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        If Trim(rsSearch("tipeOrder")) = "Booking" Then
                                            opBook(0).value = True
                                        ElseIf Trim(rsSearch("tipeOrder")) = "Non Booking" Then
                                            opBook(1).value = True
                                        Else
                                            opBook(2).value = True
                                        End If
                                    End If
                                Case 1
                                    chKriteria(i).value = IIf(Trim(rsSearch("thVehicle1")) = "0" Or Trim(rsSearch("thVehicle2")) = "0", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        text(9) = CDbl(Trim(rsSearch("thVehicle1")))
                                        text(10) = CDbl(Trim(rsSearch("thVehicle2")))
                                    End If
                                Case 2
                                    chKriteria(i).value = IIf(Trim(rsSearch("km1")) = "0" Or Trim(rsSearch("km2")) = "0", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        text(11) = CDbl(rsSearch("km1"))
                                        text(12) = CDbl(rsSearch("km2"))
                                    End If
                                Case 3
                                    chKriteria(i).value = IIf(Trim(rsSearch("lastService1")) = "0" Or Trim(rsSearch("lastService2")) = "0", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        text(13) = CDbl(rsSearch("lastService1"))
                                        text(8) = CDbl(rsSearch("lastService2"))
                                    End If
                                Case 4
                                    chKriteria(i).value = IIf(Format(rsSearch("birthday1"), "MM/DD/YYYY") = "01/01/1900" Or Format(rsSearch("birthday2"), "MM/DD/YYYY") = "01/01/1900", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        tglLahir1 = Format(rsSearch("birthday1"), "MM/DD/YYYY")
                                        tglLahir2 = Format(rsSearch("birthday2"), "MM/DD/YYYY")
                                    End If
                                Case 5
                                    chKriteria(i).value = IIf(Trim(rsSearch("hobby")) = "", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        text(14) = Trim(rsSearch("hobby"))
                                    End If
                                Case 6
                                    chKriteria(i).value = IIf(Format(rsSearch("delivday1"), "MM/DD/YYYY") = "01/01/1900" Or Format(rsGrid1.Fields(13), "MM/DD/YYYY") = "01/01/1900", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        Ddate1 = Format(rsSearch("delivday1"), "MM/DD/YYYY")
                                        Ddate2 = Format(rsSearch("delivday2"), "MM/DD/YYYY")
                                    End If
                                Case 7
                                    chKriteria(i).value = IIf(Trim(rsSearch("modelVehicle")) = "", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        If Trim(rsSearch("modelVehicle")) = "A" Then
                                            chModel.value = Checked
                                        End If
                                    End If
                                Case 8
                                    chKriteria(i).value = IIf(Trim(rsSearch("kodeProgram")) = "", Unchecked, Checked)
                                    If chKriteria(i).value = Checked Then
                                        text(6) = Trim(rsSearch("KodeProgram"))
                                    End If
                                Case 9
                                    chKriteria(i).value = IIf(Trim(rsSearch("diReminder")) = "0", Unchecked, Checked)
                            End Select
                            If chKriteria(i).value = Checked Then
                                chKrti.value = Checked
                            End If
                        Next
                        
                        chAllDealer = IIf(Trim(rsGrid1.Fields(15)) = "V", Unchecked, Checked)
                        chStatus.value = IIf(Trim(rsGrid1.Fields(18)) = "Non Aktif", Unchecked, Checked)
                        If chStatus.value = Checked And Trim(text(0)) <> "" Then
                            tblDeactivation.Visible = True
                        End If
                    End If
            End If
        Case 5
            If text(Index) <> "" Then
                    Set rsSearch = objO.FetchDataHO("Select * from KRITERIA_EVENT WITH(NOLOCK) where KRITERIA = '" & Trim(text(Index)) & "'")
                    If Not rsSearch.EOF Then
                        opTipe(0).Enabled = False
                        opTipe(1).Enabled = False
                        RangeTgl1.Enabled = False
                        RangeTgl2.Enabled = False
                        text(3).Enabled = False
                        text(4).Enabled = False
                        If Trim(tipeBY) = "TANGGAL" Then
                            opTipe(0).Enabled = True
                            opTipe(0).value = True
                            opTipe(0).SetFocus
                            opTipe(0).Enabled = False
                            RangeTgl1.Enabled = True
                            RangeTgl2.Enabled = True
                        ElseIf Trim(tipeBY) = "TEXT" Then
                            opTipe(1).Enabled = True
                            opTipe(1).value = True
                            opTipe(1).SetFocus
                            opTipe(1).Enabled = False
                            text(3).Enabled = True
                            text(4).Enabled = True
                        End If
                    Else
                    End If
            End If
        Case 6
            If text(Index) <> "" Then
                Set rsSearch = objO.FetchDataHO("SELECT * FROM HO_program A WITH(NOLOCK) LEFT JOIN HO_PROCBG B WITH(NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE A.id_Program = '" & Trim(text(Index)) & "' and FS_LOG_DEL = ''")
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "PROGRAM"
                    frmCariHO.Show (vbModal)
                Else
                    text(7) = rsSearch("PROGRAM")
                End If
            End If
    End Select
     
    Set rsSearch = Nothing
End Sub

Private Function isiNoUrut() As String
        Set rsTemp = objO.FetchDataHO("SELECT NOEVENT FROM HO_EVENT WITH(NOLOCK) WHERE left(NOEVENT,2) = '" & Right(year(isnow), 2) & "' AND SUBSTRING(NOEVENT,3,2) = '" & Format(Month(isnow), "0#") & "' ORDER BY NOEVENT DESC")
        If Not rsTemp.EOF Then
            isiNoUrut = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(CDbl(Right(rsTemp("NOEVENT"), 4)) + 1, "0###")
        Else
            isiNoUrut = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(1, "0###")
        End If
End Function

Private Sub isiModel()
    Set rsTemp = support.searchVehicleModel
    If Not rsTemp.EOF Then
        Call createTbl
        Do While Not rsTemp.EOF
            rsGridMdl.AddNew
            rsGridMdl!grpMDL = Trim(rsTemp("FS_GRP_MDL"))
            rsGridMdl!kodemdl = Trim(rsTemp("MDL"))
            rsGridMdl!NAMAMDL = Trim(rsTemp("DESC"))
            rsGridMdl!ikut = IIf(rsTemp("PRO") = "", False, True)
            rsGridMdl.Update
            rsTemp.MoveNext
        Loop
        
        rsGridMdl.MoveFirst
        GridModel.Columns(0).Width = 1200
        GridModel.Columns(1).Width = 2000
        GridModel.Columns(2).Width = 4000
        GridModel.Columns(3).Alignment = dbgCenter
        GridModel.Columns(3).Width = 800
    End If
    If rsTemp.State <> adStateClosed Then rsTemp.Close
End Sub

Private Sub createTbl()
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    With rsGridMdl
        .Fields.Append "GRPMDL", adVarChar, 20
        .Fields.Append "KODEMDL", adVarChar, 20
        .Fields.Append "NAMAMDL", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridModel.DataSource = rsGridMdl
    GridModel.AlternatingRowStyle = True
End Sub

Private Sub GridModel_FilterChange()
    On Error GoTo errHandler
    Set cols = GridModel.Columns
    Dim C As Integer
    
    C = GridModel.col
    GridModel.HoldFields
    rsGridMdl.filter = getFilter()
    GridModel.col = C
    GridModel.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub GridModel_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 3 Then
        If rsGridMdl.State Then
            rsGridMdl.MoveFirst
            Do While Not rsGridMdl.EOF
                rsGridMdl!ikut = IIf(rsGridMdl!ikut = True, False, True)
                rsGridMdl.MoveNext
            Loop
            rsGridMdl.MoveFirst
        End If
    End If
End Sub
