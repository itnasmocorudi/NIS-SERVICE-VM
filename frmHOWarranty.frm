VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmHOWarranty 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DAFTAR WARRANTY"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20220
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   6975
      Left            =   480
      TabIndex        =   16
      Top             =   480
      Width           =   11895
      Begin VB.Frame frmeJob 
         BackColor       =   &H0068ECE6&
         Caption         =   $"frmHOWarranty.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5655
         Left            =   1440
         TabIndex        =   45
         Top             =   600
         Width           =   10335
         Begin VB.CheckBox chJob 
            BackColor       =   &H0068ECE6&
            Caption         =   "Job"
            Height          =   255
            Left            =   2880
            TabIndex        =   83
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox chPart 
            BackColor       =   &H0068ECE6&
            Caption         =   "Part"
            Height          =   255
            Left            =   2880
            TabIndex        =   82
            Top             =   1315
            Value           =   1  'Checked
            Width           =   615
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0068ECE6&
            Height          =   1695
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   2655
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Satuan"
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
               Index           =   17
               Left            =   1440
               TabIndex        =   81
               Top             =   600
               Width           =   855
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Flag"
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
               Index           =   16
               Left            =   1440
               TabIndex        =   78
               Top             =   1080
               Width           =   735
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Qty"
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
               Index           =   15
               Left            =   1440
               TabIndex        =   77
               Top             =   840
               Width           =   735
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Nama Part"
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
               Index           =   14
               Left            =   1440
               TabIndex        =   76
               Top             =   360
               Width           =   1095
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Item Part"
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
               Index           =   13
               Left            =   1440
               TabIndex        =   75
               Top             =   120
               Width           =   975
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Rate"
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
               Index           =   12
               Left            =   120
               TabIndex        =   74
               Top             =   1320
               Width           =   1215
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Operasi"
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
               Index           =   11
               Left            =   120
               TabIndex        =   73
               Top             =   1080
               Width           =   975
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Kode Operasi"
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
               Index           =   10
               Left            =   120
               TabIndex        =   72
               Top             =   840
               Width           =   1335
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Job"
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
               TabIndex        =   71
               Top             =   600
               Width           =   1215
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Model"
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
               TabIndex        =   70
               Top             =   360
               Width           =   1215
            End
            Begin VB.CheckBox chPilih 
               BackColor       =   &H0068ECE6&
               Caption         =   "Kode Model"
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
               TabIndex        =   68
               Top             =   120
               Width           =   1215
            End
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   13
            Left            =   3570
            TabIndex        =   37
            Top             =   360
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   5520
            TabIndex        =   46
            Top             =   360
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
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
         Begin TrueOleDBGrid70.TDBGrid GridJob 
            Height          =   3495
            Left            =   120
            TabIndex        =   47
            Top             =   2040
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   6165
            _LayoutType     =   4
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
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).DataField=   ""
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).DataField=   ""
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).DataField=   ""
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).DataField=   ""
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).DataField=   ""
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14933984
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
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
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFF80&"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=46,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=28,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=32,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=17"
            _StyleDefs(80)  =   "Named:id=33:Normal"
            _StyleDefs(81)  =   ":id=33,.parent=0"
            _StyleDefs(82)  =   "Named:id=34:Heading"
            _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(84)  =   ":id=34,.wraptext=-1"
            _StyleDefs(85)  =   "Named:id=35:Footing"
            _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(87)  =   "Named:id=36:Selected"
            _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=37:Caption"
            _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(91)  =   "Named:id=38:HighlightRow"
            _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=39:EvenRow"
            _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
            _StyleDefs(95)  =   "Named:id=40:OddRow"
            _StyleDefs(96)  =   ":id=40,.parent=33"
            _StyleDefs(97)  =   "Named:id=41:RecordSelector"
            _StyleDefs(98)  =   ":id=41,.parent=34"
            _StyleDefs(99)  =   "Named:id=42:FilterBar"
            _StyleDefs(100) =   ":id=42,.parent=33"
            _StyleDefs(101) =   "Named:id=131:TipeTanpaTipe"
            _StyleDefs(102) =   ":id=131,.parent=33,.fgcolor=&HFF&"
            _StyleDefs(103) =   "Named:id=132:TigaDigitKosong"
            _StyleDefs(104) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         End
         Begin AIFCmp1.asxPowerButton addJob 
            Height          =   360
            Left            =   8760
            TabIndex        =   44
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
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
         Begin AIFCmp1.asxPowerButton editJob 
            Height          =   360
            Left            =   8760
            TabIndex        =   48
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
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
         Begin AIFCmp1.asxPowerButton delJob 
            Height          =   360
            Left            =   8760
            TabIndex        =   49
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
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
         Begin AIFCmp1.asxPowerButton closeJob 
            Height          =   360
            Left            =   8760
            TabIndex        =   50
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
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
            Index           =   16
            Left            =   5520
            TabIndex        =   39
            Top             =   675
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   15
            Left            =   3570
            TabIndex        =   38
            Top             =   675
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
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
            Mask            =   ">CCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   17
            Left            =   3570
            TabIndex        =   40
            Top             =   1305
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   18
            Left            =   5520
            TabIndex        =   41
            Top             =   1305
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   19
            Left            =   3570
            TabIndex        =   42
            Top             =   1605
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   20
            Left            =   7920
            TabIndex        =   54
            Top             =   675
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   5
            Left            =   5520
            TabIndex        =   69
            Top             =   1005
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   6
            Left            =   6240
            TabIndex        =   43
            Top             =   1605
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">CCCCC"
            PromptChar      =   ">"
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Satuan"
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
            Left            =   5520
            TabIndex        =   80
            Top             =   1680
            Width           =   510
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
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
            Left            =   3270
            TabIndex        =   53
            Top             =   1650
            Width           =   270
         End
         Begin VB.Shape Shape6 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   4320
            Top             =   690
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   7800
            Top             =   675
            Width           =   60
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Operasi"
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
            Left            =   4500
            TabIndex        =   52
            Top             =   720
            Width           =   960
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   5400
            Top             =   360
            Width           =   60
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
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
            Left            =   3120
            TabIndex        =   51
            Top             =   405
            Width           =   420
         End
      End
      Begin VB.Frame frmeUnit 
         BackColor       =   &H0068ECE6&
         Caption         =   $"frmHOWarranty.frx":009E
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   1200
         TabIndex        =   28
         Top             =   1080
         Width           =   10335
         Begin VB.ComboBox cboCountry1 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1080
            Width           =   855
         End
         Begin VB.ComboBox cboCountry 
            Height          =   315
            Left            =   2250
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   720
            Width           =   855
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   7
            Left            =   2250
            TabIndex        =   8
            Top             =   360
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   4200
            TabIndex        =   29
            Top             =   360
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
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
            Index           =   10
            Left            =   4080
            TabIndex        =   11
            Top             =   720
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   9
            Left            =   3120
            TabIndex        =   10
            Top             =   720
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   ">CCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   4080
            TabIndex        =   14
            Top             =   1080
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   3120
            TabIndex        =   13
            Top             =   1080
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   ">CCCCCC"
            PromptChar      =   ">"
         End
         Begin TrueOleDBGrid70.TDBGrid GridUnit 
            Height          =   3255
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   5741
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
         Begin AIFCmp1.asxPowerButton cmdAddUnit 
            Height          =   360
            Left            =   8520
            TabIndex        =   15
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
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
         Begin AIFCmp1.asxPowerButton cmdUbahUnit 
            Height          =   360
            Left            =   8520
            TabIndex        =   34
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
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
         Begin AIFCmp1.asxPowerButton cmdDelUnit 
            Height          =   360
            Left            =   8520
            TabIndex        =   35
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
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
         Begin AIFCmp1.asxPowerButton cmdTutup 
            Height          =   360
            Left            =   8520
            TabIndex        =   36
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
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
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Rangka"
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
            Left            =   1320
            TabIndex        =   32
            Top             =   1125
            Width           =   780
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   4080
            Top             =   360
            Width           =   60
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model"
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
            Left            =   1680
            TabIndex        =   31
            Top             =   405
            Width           =   420
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Rangka"
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
            Left            =   1320
            TabIndex        =   30
            Top             =   765
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0068ECE6&
         Height          =   1935
         Left            =   6600
         TabIndex        =   59
         Top             =   120
         Width           =   1575
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tgl TI"
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
            TabIndex        =   66
            Top             =   1560
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "No TI"
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
            TabIndex        =   65
            Top             =   1320
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tgl Selesai"
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
            TabIndex        =   64
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tgl Mulai"
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
            TabIndex        =   63
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
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
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Warranty"
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
            TabIndex        =   61
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Kode "
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
            TabIndex        =   60
            Top             =   120
            Width           =   855
         End
      End
      Begin MSComCtl2.DTPicker TglMulai 
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   1150
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   96600065
         CurrentDate     =   39853
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   0
         Top             =   550
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin TrueOleDBGrid70.TDBGrid GridWarranty 
         Height          =   4695
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   8281
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
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   1305
         _ExtentX        =   2302
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
      Begin MSComCtl2.DTPicker TglSelesai 
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   1150
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   96600065
         CurrentDate     =   39853
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   1
         Top             =   850
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin AIFCmp1.asxPowerButton cmdUnit 
         Height          =   360
         Left            =   10560
         TabIndex        =   6
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "&Isi Unit"
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
      Begin AIFCmp1.asxPowerButton cmdJob 
         Height          =   360
         Left            =   10560
         TabIndex        =   7
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "&Job dan Part"
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
         Index           =   3
         Left            =   2400
         TabIndex        =   55
         Top             =   1450
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   3435
         TabIndex        =   4
         Top             =   1450
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSComCtl2.DTPicker TglTI 
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   96600065
         CurrentDate     =   39853
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Left            =   1680
         TabIndex        =   58
         Top             =   1800
         Width           =   570
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3300
         TabIndex        =   57
         Top             =   1485
         Width           =   60
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No TI"
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
         Left            =   1860
         TabIndex        =   56
         Top             =   1500
         Width           =   390
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
         Left            =   1410
         TabIndex        =   23
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Warranty"
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
         Left            =   1110
         TabIndex        =   22
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   1185
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
         Left            =   3855
         TabIndex        =   20
         Top             =   1185
         Width           =   225
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Warranty"
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
         Left            =   1155
         TabIndex        =   19
         Top             =   285
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   5280
         Top             =   557
         Width           =   60
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7875
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13891
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "    Daftar Warranty"
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
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   7800
      TabIndex        =   24
      Top             =   7560
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
      Left            =   8760
      TabIndex        =   25
      Top             =   7560
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
      Left            =   9720
      TabIndex        =   26
      Top             =   7560
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
      Left            =   10680
      TabIndex        =   27
      Top             =   7560
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
      Left            =   6480
      TabIndex        =   79
      Top             =   7560
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
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   4920
      TabIndex        =   84
      Top             =   7560
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
      Left            =   1080
      TabIndex        =   85
      Top             =   7560
      Visible         =   0   'False
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
Attribute VB_Name = "frmHOWarranty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 09 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsGrid As New ADODB.Recordset
Dim rsGridUnit As New ADODB.Recordset
Dim rsGridJob As New ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim newRecord As Boolean
Dim pesan As String
Dim cols As TrueOleDBGrid70.Columns

Private Sub ber_sih(idx As Integer)
    If idx = 1 Then
        For i = 0 To 4
            Text(i) = ""
        Next
        Text(3) = "TI" & year(isnow)
    ElseIf idx = 2 Then
        For i = 7 To 12
            Text(i) = ""
        Next
    ElseIf idx = 3 Then
        For i = 13 To 20
            Select Case i
                Case 19, 20
                    Text(i) = 0
                Case Else
                    Text(i) = ""
            End Select
        Next
        Text(6) = ""
    End If
End Sub

Private Sub buka_Tutup(idx As Integer, Optional Flg As Boolean = True)
    If idx = 1 Then
        Text(1).Enabled = Flg
        Text(2).Enabled = Flg
        Text(4).Enabled = Flg
        Text(3).Enabled = Flg
        TglMulai.Enabled = Flg
        TglSelesai.Enabled = Flg
        TglTI.Enabled = Flg
        cmdUnit.Enabled = Flg
        cmdJob.Enabled = Flg
    ElseIf idx = 2 Then
        Text(7).Enabled = Flg
        Text(9).Enabled = Flg
        Text(11).Enabled = Flg
        Text(10).Enabled = Flg
        Text(12).Enabled = Flg
        cboCountry.Enabled = Flg
        cboCountry1.Enabled = Flg
    ElseIf idx = 3 Then
        Text(13).Enabled = Flg
        Text(15).Enabled = Flg
        Text(16).Enabled = Flg
        Text(17).Enabled = Flg
        Text(18).Enabled = Flg
        Text(19).Enabled = Flg
        Text(6).Enabled = Flg
    End If
End Sub

Private Function isiIdExcel() As String
    Dim noUrut As Integer
    Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_Warranty WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY id_Warranty DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("id_Warranty"), Len(rsTemp("id_Warranty")) - 4))
        noUrut = noUrut + 1
        isiIdExcel = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        isiIdExcel = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(noUrut, "00000#")
    End If
End Function

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_Warranty WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY id_Warranty DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("id_Warranty"), Len(rsTemp("id_Warranty")) - 4))
        noUrut = noUrut + 1
        Text(0) = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        Text(0) = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(noUrut, "00000#")
    End If
End Sub

Private Sub isiCombo()
    Set rsTemp = objO.FetchDataHO("SELECT * FROM T_COUNTRY WITH (NOLOCK) ORDER BY kode")
    If Not rsTemp.EOF Then
        cboCountry.Clear
        cboCountry1.Clear
        Do While Not rsTemp.EOF
            cboCountry.AddItem rsTemp("KODE")
            cboCountry1.AddItem rsTemp("KODE")
            rsTemp.MoveNext
        Loop
    End If
End Sub

Private Sub refreshGrid(idx As Integer)
    If idx = 1 Then
        GridWarranty.ClearFields
        Set rsGrid = objO.FetchDataHO("SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI, NOTI, TGLTI FROM HO_Warranty WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY id_warranty")
        If Not rsGrid.EOF Then
            GridWarranty.DataSource = rsGrid
            Call HeaderColom(idx)
        End If
    ElseIf idx = 2 Then
        GridUnit.ClearFields
        Set rsGridUnit = objO.FetchDataHO("SELECT ID_WARRANTY, MDL, MODEL, NO_RANGKA1, NO_RANGKA2 FROM tempHO_wardtl WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "' ORDER BY MDl")
        If Not rsGridUnit.EOF Then
            GridUnit.DataSource = rsGridUnit
            Call HeaderColom(idx)
        Else
            GridUnit.DataSource = rsGridUnit
            Call HeaderColom(idx)
        End If
    ElseIf idx = 3 Then
        GridJob.ClearFields
        Set rsGridJob = objO.FetchDataHO("SELECT MDL,MODEL,JOB,FS_ITEM,JOBKET,RATE,ITEMPART,NAMAPART,QTY,FLAG,SATPART FROM tempHO_WARJOBPART WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "' ORDER BY mdl, flag, job, fs_item, itempart")
        If Not rsGridJob.EOF Then
            GridJob.DataSource = rsGridJob
            Call HeaderColom(idx)
        End If
    End If
End Sub

Private Sub HeaderColom(idz As Integer)
    If idz = 1 Then
        GridWarranty.Columns(0).Caption = "ID"
        GridWarranty.Columns(0).Width = IIf(chPilih(0) = Checked, 1200, 0)
        GridWarranty.Columns(1).Caption = "WARRANTY"
        GridWarranty.Columns(1).Width = IIf(chPilih(1) = Checked, 2500, 0)
        GridWarranty.Columns(2).Caption = "KETERANGAN"
        GridWarranty.Columns(2).Width = IIf(chPilih(2) = Checked, 3000, 0)
        GridWarranty.Columns(3).Caption = "TANGGAL MULAI"
        GridWarranty.Columns(3).Width = IIf(chPilih(3) = Checked, 1500, 0)
        GridWarranty.Columns(4).Caption = "TANGGAL SELESAI"
        GridWarranty.Columns(4).Width = IIf(chPilih(4) = Checked, 1500, 0)
        GridWarranty.Columns(5).Caption = "NO TI"
        GridWarranty.Columns(5).Width = IIf(chPilih(5) = Checked, 1200, 0)
        GridWarranty.Columns(6).Caption = "TANGGAL TI"
        GridWarranty.Columns(6).Width = IIf(chPilih(6) = Checked, 1200, 0)
    ElseIf idz = 2 Then
        GridUnit.Columns(0).Caption = "ID"
        GridUnit.Columns(0).Width = 1200
        GridUnit.Columns(1).Caption = "MDL"
        GridUnit.Columns(1).Width = 2000
        GridUnit.Columns(2).Caption = "MODEL"
        GridUnit.Columns(2).Width = 2500
        GridUnit.Columns(3).Caption = "NO RANGKA AWAL"
        GridUnit.Columns(3).Width = 2500
        GridUnit.Columns(4).Caption = "NO RANGKA AKHIR"
        GridUnit.Columns(4).Width = 2500
    ElseIf idz = 3 Then
        GridJob.Columns(0).Caption = "MDL"
        GridJob.Columns(0).Width = IIf(chPilih(7) = Checked, 2000, 0)
        GridJob.Columns(1).Caption = "MODEL"
        GridJob.Columns(1).Width = IIf(chPilih(8) = Checked, 2500, 0)
        GridJob.Columns(2).Caption = "JOB"
        GridJob.Columns(2).Width = IIf(chPilih(9) = Checked, 1000, 0)
        GridJob.Columns(3).Caption = "KODE OPERASI"
        GridJob.Columns(3).Width = IIf(chPilih(10) = Checked, 2500, 0)
        GridJob.Columns(4).Caption = "OPERASI"
        GridJob.Columns(4).Width = IIf(chPilih(11) = Checked, 3000, 0)
        GridJob.Columns(5).Caption = "RATE"
        GridJob.Columns(5).Width = IIf(chPilih(12) = Checked, 1200, 0)
        GridJob.Columns(6).Caption = "ITEMPART"
        GridJob.Columns(6).Width = IIf(chPilih(13) = Checked, 2000, 0)
        GridJob.Columns(7).Caption = "NAMAPART"
        GridJob.Columns(7).Width = IIf(chPilih(14) = Checked, 2500, 0)
        GridJob.Columns(8).Caption = "QTY"
        GridJob.Columns(8).Width = IIf(chPilih(15) = Checked, 1200, 0)
        GridJob.Columns(9).Caption = "FLAG"
        GridJob.Columns(9).Width = IIf(chPilih(16) = Checked, 800, 0)
        GridJob.Columns(10).Caption = "SATPART"
        GridJob.Columns(10).Width = IIf(chPilih(17) = Checked, 800, 0)
    End If
End Sub

Private Sub addJob_Click()
Dim boleh As Boolean

    If addJob.Caption = "&Tambah" Then
        Call ber_sih(3)
        Call buka_Tutup(3, True)
        Call refreshGrid(3)
        Text(13).BackColor = &H80000005
        If chJob = Checked Then
            Text(15) = "TWC"
            Text(16).BackColor = &H80000005
        ElseIf chPart = Checked Then
            Text(17).BackColor = &H80000005
            Text(18).BackColor = &H80000005
            Text(19).BackColor = &H80000005
            Text(6).BackColor = &H80000005
        End If
        
        Text(13).SetFocus
        addJob.Caption = "&Simpan"
        editJob.Caption = "&Batal"
    ElseIf addJob.Caption = "&Simpan" Then
        
        boleh = True
        pesan = "Lengkapi data "
        If Text(13) = "" Then
            boleh = False
            pesan = pesan + "Model. "
        End If
        If chJob = Checked Then
            If Text(15) = "" Then
                boleh = False
                pesan = pesan + "Group. "
            End If
            If Text(16) = "" Then
                boleh = False
                pesan = pesan + "Job Order. "
            End If
        ElseIf chPart = Checked Then
            If Text(17) = "" Then
                boleh = False
                pesan = pesan + "Kode Part. "
            End If
            If Text(18) = "" Then
                boleh = False
                pesan = pesan + "Nama Part. "
            End If
            If Text(19) = "" Then
                boleh = False
                pesan = pesan + "Qty. "
            End If
            If Text(6) = "" Then
                boleh = False
                pesan = pesan + "Satuan. "
            End If
        End If
        
        If boleh Then
            ClearsSql
            gsSql1(0) = "ID_WARRANTY"
            gsSql2(0) = Text(0)
            gsSql1(1) = "MDL"
            gsSql2(1) = Text(13)
            gsSql1(2) = "MODEL"
            gsSql2(2) = Text(14)
            If chJob.value = Checked Then
                gsSql1(3) = "JOB"
                gsSql2(3) = Text(15)
                gsSql1(4) = "FS_ITEM"
                gsSql2(4) = Text(16)
                gsSql1(5) = "JOBKET"
                gsSql2(5) = Text(5)
                gsSql1(6) = "RATE"
                gsSql2(6) = Replace(Text(20), ",", ".")
                gsSql1(7) = "FLAG"
                gsSql2(7) = "J"
                gsSql1(8) = "KETERANGAN"
                gsSql2(8) = Text(5)
                
                gsSql1(9) = "ITEMPART"
                gsSql2(9) = ""
                gsSql1(10) = "NAMAPART"
                gsSql2(10) = ""
                gsSql1(11) = "QTY"
                gsSql2(11) = 0
                gsSql1(12) = "SATPART"
                gsSql2(12) = ""

            ElseIf chPart.value = Checked Then
                gsSql1(3) = "JOB"
                gsSql2(3) = ""
                gsSql1(4) = "FS_ITEM"
                gsSql2(4) = ""
                gsSql1(5) = "JOBKET"
                gsSql2(5) = ""
                gsSql1(6) = "RATE"
                gsSql2(6) = 0
                gsSql1(7) = "KETERANGAN"
                gsSql2(7) = ""
            
                gsSql1(8) = "ITEMPART"
                gsSql2(8) = Text(17)
                gsSql1(9) = "NAMAPART"
                gsSql2(9) = Text(18)
                gsSql1(10) = "QTY"
                gsSql2(10) = CDbl(Text(19))
                gsSql1(11) = "FLAG"
                gsSql2(11) = "P"
                gsSql1(12) = "SATPART"
                gsSql2(12) = Trim(Text(6))
            End If
            
            If rsTemp.State Then rsTemp.Close
            If chJob.value = Checked Then
                Set rsTemp = objO.FetchDataHO("SELECT * FROM tempHO_warjobpart WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "' AND mdl = '" & Text(13) & "' AND FS_ITEM = '" & Text(16) & "'")
            ElseIf chPart.value = Checked Then
                Set rsTemp = objO.FetchDataHO("SELECT * FROM tempHO_warjobpart WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "' AND mdl = '" & Text(13) & "' AND itempart = '" & Text(17) & "'")
            End If
            
            If Not rsTemp.EOF Then
                    If chJob.value = Checked Then
                        ' UPDATE
                        sSQL = GensSql2("tempHO_WARJOBPART", True) + "WHERE id_warranty = '" & Text(0) & "' AND mdl = '" & Text(13) & "' AND job = '" & Text(15) & "'"
                    Else
                        sSQL = GensSql2("tempHO_WARJOBPART", True) + "WHERE id_warranty = '" & Text(0) & "' AND mdl = '" & Text(13) & "' AND itempart = '" & Text(17) & "'"
                    End If
            Else
                ' INSERT
                sSQL = GensSql2("tempHO_WARJOBPART", False)
            End If
            Call objO.UpdateDataHO(sSQL)
            
            Call ber_sih(3)
            Call buka_Tutup(3, False)
            Call refreshGrid(3)
            
            addJob.Caption = "&Tambah"
            editJob.Caption = "&Ubah"
        Else
            MsgBox pesan, vbCritical, "KELENGKAPAN"
            Exit Sub
        End If
    End If
End Sub

Private Sub cboCountry_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text(10).SetFocus
    End If
End Sub

Private Sub cboCountry1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text(12).SetFocus
    End If
End Sub

Private Sub chJob_Click()
    If chJob.value = Checked Then
        chPart.value = Unchecked
        Text(17).BackColor = &HC0E0FF
        Text(18).BackColor = &HC0E0FF
        Text(19).BackColor = &HC0E0FF
        Text(6).BackColor = &HC0E0FF
    
        If addJob.Caption = "&Simpan" Then
            Text(15).BackColor = &H80000005
            Text(16).BackColor = &H80000005
        Else
            Text(15).BackColor = &HC0E0FF
            Text(16).BackColor = &HC0E0FF
        End If
    End If
End Sub

Private Sub chPart_Click()
    If chPart.value = Checked Then
        chJob.value = Unchecked
        If addJob.Caption = "&Simpan" Then
            Text(17).BackColor = &H80000005
            Text(18).BackColor = &H80000005
            Text(19).BackColor = &H80000005
            Text(6).BackColor = &H80000005
        Else
            Text(17).BackColor = &HC0E0FF
            Text(18).BackColor = &HC0E0FF
            Text(19).BackColor = &HC0E0FF
            Text(6).BackColor = &HC0E0FF
        End If
        Text(15).BackColor = &HC0E0FF
        Text(16).BackColor = &HC0E0FF
    End If
End Sub

Private Sub chPilih_Click(Index As Integer)
    If Not rsGrid.EOF Then
    Select Case Index
        ' UNTUK GRID WARRANTY
        Case 0
            GridWarranty.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1200, 0)
        Case 1
            GridWarranty.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2500, 0)
        Case 2
            GridWarranty.Columns(Index).Width = IIf(chPilih(Index) = Checked, 3000, 0)
        Case 3
            GridWarranty.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1500, 0)
        Case 4
            GridWarranty.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1500, 0)
        Case 5
            GridWarranty.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1200, 0)
        Case 6
            GridWarranty.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1200, 0)
        ' UNTUK GRID JOB & PART
        Case 7
            GridJob.Columns(0).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 8
            GridJob.Columns(1).Width = IIf(chPilih(Index) = Checked, 2500, 0)
        Case 9
            GridJob.Columns(2).Width = IIf(chPilih(Index) = Checked, 1000, 0)
        Case 10
            GridJob.Columns(3).Width = IIf(chPilih(Index) = Checked, 2500, 0)
        Case 11
            GridJob.Columns(4).Width = IIf(chPilih(Index) = Checked, 3000, 0)
        Case 12
            GridJob.Columns(5).Width = IIf(chPilih(Index) = Checked, 1200, 0)
        Case 13
            GridJob.Columns(6).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 14
            GridJob.Columns(7).Width = IIf(chPilih(Index) = Checked, 2500, 0)
        Case 15
            GridJob.Columns(8).Width = IIf(chPilih(Index) = Checked, 1200, 0)
        Case 16
            GridJob.Columns(9).Width = IIf(chPilih(Index) = Checked, 800, 0)
        Case 17
            GridJob.Columns(10).Width = IIf(chPilih(Index) = Checked, 800, 0)
    End Select
    End If
End Sub

Private Sub closeJob_Click()
    addJob.Caption = "&Tambah"
    editJob.Caption = "&Ubah"
    frmeJob.Visible = False
End Sub

Private Sub cmdAddUnit_Click()
Dim boleh As Boolean
Dim pesan As String

    If cmdAddUnit.Caption = "&Tambah" Then
        Call ber_sih(2)
        Call buka_Tutup(2)
        
        Text(7).SetFocus
        cmdAddUnit.Caption = "&Simpan"
        cmdUbahUnit.Caption = "&Batal"
        newRecord = True
    ElseIf cmdAddUnit.Caption = "&Simpan" Then
    
        boleh = True
        pesan = "Lengkapi Data Berikut "
        
        If Text(7) = "" Then
            boleh = False
            pesan = pesan + "Model. "
        End If
        If Text(10) = "" Then
            boleh = False
            pesan = pesan + "No. Rangka "
        End If
        If Text(12) = "" Then
            boleh = False
            pesan = pesan + "No. Rangka "
        End If
        
        If boleh Then
            Call ClearsSql
            gsSql1(0) = "ID_WARRANTY"
            gsSql2(0) = Text(0)
            gsSql1(1) = "MDL"
            gsSql2(1) = Text(7)
            gsSql1(2) = "MODEL"
            gsSql2(2) = Text(8)
            'SIMPAN UNTUK BEBERAPA NO RANGKA
            gsSql1(3) = "NO_RANGKA1"
            gsSql2(3) = cboCountry & Replace(Trim(Text(9)), ">", "") & Replace(Trim(Text(10)), ">", "")
            gsSql1(4) = "NO_RANGKA2"
            If Text(12) = "" Then
                gsSql2(4) = ""
            Else
                gsSql2(4) = cboCountry & Replace(Trim(Text(11)), ">", "") & Replace(Trim(Text(12)), ">", "")
            End If
        
            ' SIMPAN DATA
            Set rsTemp = objO.FetchDataHO("SELECT * FROM tempHO_wardtl WITH (NOLOCK) WHERE id_Warranty = '" & Text(0) & "' AND mdl = '" & Text(7) & "' AND no_Rangka1 < '" & cboCountry & Replace(Text(9), " > ", "") & Replace(Text(10), " > ", "") & "' AND no_Rangka2 > '" & cboCountry & Replace(Text(9), " > ", "") & Replace(Text(10), " > ", "") & "'")
            If Not rsTemp.EOF Then
                ' DATA SUDAH ADA
                If newRecord Then
                    MsgBox "DATA SUDAH ADA. Jika ingin merubah pilih UBAH", vbCritical, "DUPLIKASI DATA"
                Else
                    temp = MsgBox("Yakin akan merubah Data ?", vbYesNo, "DUPLIKASI DATA")
                    If temp = vbYes Then
                        sSQL = GensSql2("tempHO_wardtl", True) + " where id_Warranty = '" & Text(0) & "'"
                    End If
                End If
            Else
                ' DATA BARU
                sSQL = GensSql2("tempHO_wardtl", False)
            End If
             Call objO.UpdateDataHO(sSQL)
            
            Call ber_sih(2)
            Call buka_Tutup(2, False)
            Call refreshGrid(2)
            
            cmdAddUnit.Caption = "&Tambah"
            cmdUbahUnit.Caption = "&Ubah"
        Else
            MsgBox pesan, vbCritical, "KELENGKAPAN"
            Exit Sub
        End If
    End If

End Sub

Private Sub cmdDelUnit_Click()
    nrk1 = cboCountry & Replace(Text(9), ">", "") & Replace(Text(10), ">", "")
    If Text(12) = "" Then
        nrk2 = ""
    Else
        nrk2 = cboCountry1 & Replace(Text(11), ">", "") & Replace(Text(12), ">", "")
    End If
    Set rsTemp = objO.FetchDataHO("SELECT * FROM tempHO_wardtl WITH (NOLOCK) WHERE id_Warranty = '" & Text(0) & "' AND mdl = '" & Text(7) & "' AND no_rangka1 = '" & nrk1 & "' AND no_rangka2 = '" & nrk2 & "'")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin Akan menghapus data ?", vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.DeleteDataHO("DELETE tempHO_wardtl WHERE id_Warranty = '" & Text(0) & "' AND mdl = '" & Text(7) & "' AND no_rangka1 = '" & nrk1 & "' AND no_rangka2 = '" & nrk2 & "'")
        End If
    Else
        MsgBox "Pilih data yang akan dihapus"
        Exit Sub
    End If
    Call ber_sih(2)
    Call refreshGrid(2)
End Sub

Private Sub cmdHapus_Click()
    Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_WARRANTY WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "' and FS_LOG_DEL = ''")
    If Not rsTemp.EOF Then
        temp = MsgBox("Takin akan menghapus data Warranty : " & Text(1), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE HO_WARRANTY SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE id_warranty = '" & Text(0) & "' and FS_LOG_DEL = ''")
            Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('WARRANTY', '" & Text(0) & "','D')")
            Call ber_sih(1)
            Call refreshGrid(1)
        End If
    Else
        MsgBox "Pilih data yang akan dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdImport_Click()
         Dim ErrBegin As Boolean, konter As Long
       

1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim rsTempExcel As ADODB.Recordset
         Dim Jml As Integer
         Dim tempKeterangan As String
         Dim idExcel As String
         
        ' Dim konter As Integer
         
2       konter = 1
3       ProgBar.Min = 0
4       CommonDialog1.FileName = ""
5       CommonDialog1.InitDir = App.Path
6       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
7       CommonDialog1.ShowOpen
8       If CommonDialog1.FileName <> "" Then
9          If CnExel.State Then CnExel.Close
10         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;"""
11         GoSub Proses
12     Else
13        Exit Sub
14     End If
15     Exit Sub
       
Proses:
16    ErrBegin = True
         Dim VarNmFile As String
         Dim namaDealer As String
         Dim InsertData As Boolean
17       Call objHT.OpenTransactionHO
18        gBantuan = "SBE"
19       If RcExcel.State Then RcExcel.Close
20       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
21       If RcExcel.RecordCount > 0 Then
22          ProgBar.Visible = True
23          ProgBar.Max = RcExcel.RecordCount
24          RcExcel.MoveFirst
25          tempKeterangan = ""
26          idExcel = ""
27          Do Until RcExcel.EOF
28             DoEvents
               
29             ProgBar.value = konter
30             model = ""
31                InsertData = False
32                If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
33                If Trim(RcExcel.Fields(4).value) = "" Then
34                    Call objHT.CloseTransactionHO_err
35                    ProgBar.Visible = False
36                    MsgBox "No TI baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
37                    Exit Sub
38                End If
                  
39                Call ClearsSql
                  
                  
40                If idExcel = "" Then
41                    idExcel = isiIdExcel
42                    InsertData = True
43                ElseIf tempKeterangan <> Trim(RcExcel.Fields(0).value) & Trim(RcExcel.Fields(1).value) & Format(Trim(RcExcel.Fields(2).value), "mm/dd/yyyy") & Format(Trim(RcExcel.Fields(3).value), "mm/dd/yyyy") Then
44                    idExcel = isiIdExcel
45                    InsertData = True
46                End If
                  
47                If tempKeterangan = "" Then
48                    tempKeterangan = Trim(RcExcel.Fields(0).value) & Trim(RcExcel.Fields(1).value) & Format(Trim(RcExcel.Fields(2).value), "mm/dd/yyyy") & Format(Trim(RcExcel.Fields(3).value), "mm/dd/yyyy")
49                ElseIf tempKeterangan <> Trim(RcExcel.Fields(0).value) & Trim(RcExcel.Fields(1).value) & Format(Trim(RcExcel.Fields(2).value), "mm/dd/yyyy") & Format(Trim(RcExcel.Fields(3).value), "mm/dd/yyyy") Then
50                    tempKeterangan = Trim(RcExcel.Fields(0).value) & Trim(RcExcel.Fields(1).value) & Format(Trim(RcExcel.Fields(2).value), "mm/dd/yyyy") & Format(Trim(RcExcel.Fields(3).value), "mm/dd/yyyy")
51                End If
52                If InsertData Then
53                 gsSql1(0) = "ID_WARRANTY"
54                 gsSql2(0) = idExcel
55                 gsSql1(1) = "WARRANTY"
56                 gsSql2(1) = Trim(RcExcel.Fields(0).value)
57                 gsSql1(2) = "KETERANGAN"
58                 gsSql2(2) = Trim(RcExcel.Fields(1).value)
59                 gsSql1(3) = "TGLMULAI"
60                 gsSql2(3) = IIf(Trim(RcExcel.Fields(2).value) = "", "", Format(Trim(RcExcel.Fields(2).value), "mm/dd/yyyy"))
61                 gsSql1(4) = "TGLSELESAI"
62                 gsSql2(4) = IIf(Trim(RcExcel.Fields(2).value) = "", "", Format(Trim(RcExcel.Fields(3).value), "mm/dd/yyyy"))
63                 gsSql1(5) = "NOTI"
64                 gsSql2(5) = IIf(Nul(Trim(RcExcel.Fields(4).value)) = "", "", Trim(RcExcel.Fields(4).value))
65                 gsSql1(6) = "TGLTI"
66                 gsSql2(6) = IIf(Nul(Trim(RcExcel.Fields(5).value)) = "", "", Format(Trim(RcExcel.Fields(5).value), "mm/dd/yyyy"))

                   
                   ' SIMPAN DATA
67                 sSQL = GensSql2("HO_WARRANTY", False)
68                 Call objHT.InsertData_Ho(sSQL)
69                 Call objHT.InsertData_Ho("INSERT INTO TEMPKIRIM VALUES ('WARRANTY', '" & gsSql2(0) & "', 'S')")
                  
70                End If
                  
                 
                  '-------------------------------------------------------------------
                  '- 2015-04-01 | RUDI | SIMPAN DETAIL UNIT KENDARAAN MODEL WARRANTY -
                  '-------------------------------------------------------------------
                  
71                If Trim(RcExcel.Fields(6).value) <> "" And Trim(RcExcel.Fields(7).value) <> "" And Trim(RcExcel.Fields(8).value) <> "" And Trim(RcExcel.Fields(9).value) <> "" Then
                      'JIKA ADA DATA MAKA SIMPAN DATA
72                    Call objHT.InsertData_Ho("INSERT INTO ho_warDtl (id_Warranty,MDL,MODEL,No_Rangka1,No_Rangka2) VALUES ('" & idExcel & "','" & Trim(RcExcel.Fields(6).value) & "','" & Trim(RcExcel.Fields(7).value) & "','" & Trim(RcExcel.Fields(8).value) & "','" & Trim(RcExcel.Fields(9).value) & "')")
73                End If
                  
                  '--------------------------------------------------------------------
                  '- 2015-04-01 | RUDI | SIMPAN DETAIL JOB & PART                     -
                  '--------------------------------------------------------------------
                  
74                If Trim(RcExcel.Fields(6).value) <> "" And Trim(RcExcel.Fields(7).value) <> "" And Trim(RcExcel.Fields(10).value) <> "" And Trim(RcExcel.Fields(11).value) <> "" And Trim(RcExcel.Fields(13).value) <> "" Then
                      'lakukan insert data job
75                    Call objHT.InsertData_Ho("INSERT INTO ho_WarJobPart (ID_WARRANTY,MDL,MODEL,JOB,FS_ITEM,JOBKET,RATE,ITEMPART,NAMAPART,SATPART,QTY,FLAG,KETERANGAN) VALUES ('" & idExcel & "','" & Trim(RcExcel.Fields(6).value) & "','" & Trim(RcExcel.Fields(7).value) & "','" & Trim(RcExcel.Fields(10).value) & "','" & Trim(RcExcel.Fields(11).value) & "','" & Trim(RcExcel.Fields(12).value) & "'," & Replace(Trim(RcExcel.Fields(13).value), ",", ".") & ",'','','',0,'J','')")
                      
76                ElseIf Trim(RcExcel.Fields(6).value) <> "" And Trim(RcExcel.Fields(7).value) <> "" And Trim(RcExcel.Fields(14).value) <> "" And Trim(RcExcel.Fields(15).value) <> "" And Trim(RcExcel.Fields(17).value) <> "" Then
                      'lakukan insert data part
77                    Call objHT.InsertData_Ho("INSERT INTO ho_WarJobPart (ID_WARRANTY,MDL,MODEL,JOB,FS_ITEM,JOBKET,RATE,ITEMPART,NAMAPART,SATPART,QTY,FLAG,KETERANGAN) VALUES ('" & idExcel & "','" & Trim(RcExcel.Fields(6).value) & "','" & Trim(RcExcel.Fields(7).value) & "','" & Trim(RcExcel.Fields(10).value) & "','" & Trim(RcExcel.Fields(11).value) & "','" & Trim(RcExcel.Fields(12).value) & "'," & Replace(Trim(RcExcel.Fields(13).value), ",", ".") & ",'" & Trim(RcExcel.Fields(14).value) & "','" & Trim(RcExcel.Fields(15).value) & "','" & Trim(RcExcel.Fields(16).value) & "'," & CInt(Trim(RcExcel.Fields(17).value)) & ",'P','" & Trim(RcExcel.Fields(19).value) & "')")
78                End If
                  
79             RcExcel.MoveNext
80             konter = konter + 1
81          Loop
82       End If
83       RcExcel.Close
84       CnExel.Close
85       ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
86        Call objHT.CloseTransactionHO
87        Call refreshGrid(1)
88       MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
      '   GridAFI.Refresh
89       Return
90       Exit Sub
vb_error:
         
91       If ErrBegin = True Then
           ' DbCon.RollbackTrans
92       End If
93       ProgBar.Visible = False
         'MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")
         
         '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
94       varErlError = Erl
95       varErrNumber = err.Number
         ''varErlError,varErrNumber
96       Call objHT.CloseTransactionHO_err
         
97       MsgBox ErrMessage(varErlError, varErrNumber, "Procedure : " & Me.Name & ".CmdImport_Click")
       
End Sub

Private Sub cmdJob_Click()
    Call ber_sih(3)
    Call refreshGrid(3)
    chJob.value = Checked
    Text(13).BackColor = &HC0E0FF
    Text(15).BackColor = &HC0E0FF
    Text(16).BackColor = &HC0E0FF
    Text(17).BackColor = &HC0E0FF
    Text(18).BackColor = &HC0E0FF
    Text(19).BackColor = &HC0E0FF
    Text(6).BackColor = &HC0E0FF
    Call buka_Tutup(3, False)
    
    frmeJob.Top = cmdJob.Top + cmdJob.Height  '1680
    frmeJob.Left = cmdJob.Left + cmdJob.Width - frmeJob.Width  '1680
    frmeJob.Visible = True
End Sub

Private Sub CmdKeluar_Click()
    Call objO.DeleteDataHO("DELETE tempho_warDtl where id_warranty = '" & Trim(Text(0)) & "'")
    Call objO.DeleteDataHO("DElETE tempho_WarJobPart WHERE id_warranty = '" & Trim(Text(0)) & "'")
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String

1         If cmdSimpan.Caption = "&Tambah" Then
2             newRecord = True
3             Call ber_sih(1)
4             Call buka_Tutup(1)
              
5             Text(1).SetFocus
6             cmdSimpan.Caption = "&Simpan"
7             cmdUbah.Caption = "&Ubah"
8             newRecord = True
9             Call isiId
10        ElseIf cmdSimpan.Caption = "&Simpan" Then
              
11            boleh = True
12            pesan = "Lengkapi Data Berikut "
              
13            If Text(1) = "" Then
14                boleh = False
15                pesan = pesan + "Warranty. "
16            End If
17            If Text(4) = "" Then
18                boleh = False
19                pesan = pesan + "No. TI. "
20            End If
              
21            If boleh Then
                  
22                Call ClearsSql
                  
23                gsSql1(0) = "ID_WARRANTY"
24                gsSql2(0) = Text(0)
25                gsSql1(1) = "WARRANTY"
26                gsSql2(1) = Text(1)
27                gsSql1(2) = "KETERANGAN"
28                gsSql2(2) = Text(2)
29                gsSql1(3) = "TGLMULAI"
30                gsSql2(3) = Format(TglMulai, "mm/dd/yyyy")
31                gsSql1(4) = "TGLSELESAI"
32                gsSql2(4) = Format(TglSelesai, "mm/dd/yyyy")
33                gsSql1(5) = "NOTI"
34                gsSql2(5) = Text(3) & "-" & Text(4)
35                gsSql1(6) = "TGLTI"
36                gsSql2(6) = Format(TglTI, "mm/dd/yyyy")
                  
                  ' SIMPAN DATA
37                If rsTemp.State Then rsTemp.Close
38                Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_WARRANTY WITH (NOLOCK) WHERE id_Warranty = '" & Text(0) & "' and FS_LOG_DEL = ''")
39                If Not rsTemp.EOF Then
40                        temp = MsgBox("Yakin akan merubah Data ?", vbYesNo, "DUPLIKASI DATA")
41                        If temp = vbYes Then
42                            sSQL = GensSql2("HO_WARRANTY", True) + " where id_Warranty = '" & Text(0) & "' and FS_LOG_DEL = ''"
43                        End If
44                Else
                      ' DATA BARU
45                    sSQL = GensSql2("HO_WARRANTY", False)
46                End If
47                Call objO.UpdateDataHO(sSQL)
                  
                  ' SIMPAN DETAIL
48                Set rsTemp = objO.FetchDataHO("SELECT * FROM tempho_warDtl WITH (NOLOCK) WHERE id_warranty =  '" & Trim(Text(0)) & "'")
49                If Not rsTemp.EOF Then
50                    Call objO.DeleteDataHO("DELETE ho_warDtl WHERE id_warranty = '" & Trim(Text(0)) & "'")
51                    Call objO.InsertDataHO("INSERT INTO ho_warDtl (id_Warranty,MDL,MODEL,No_Rangka1,No_Rangka2) SELECT ID_WARRANTY, MDL, MODEL, NO_RANGKA1, NO_RANGKA2 FROM tempho_warDtl WITH (NOLOCK) WHERE id_warranty =  '" & Trim(Text(0)) & "' ORDER BY mdl, no_rangka1")
52                    Call objO.DeleteDataHO("DELETE tempho_warDtl where id_warranty = '" & Trim(Text(0)) & "'")
53                Else
54                    Call objO.DeleteDataHO("DELETE ho_warDtl WHERE id_warranty = '" & Trim(Text(0)) & "'")
55                End If
56                rsTemp.Close
                  
                  ' SIMPAN DETAIL JOB & PART
57                Set rsTemp = objO.FetchDataHO("SELECT * FROM tempho_WarJobPart WITH (NOLOCK) WHERE id_warranty =  '" & Trim(Text(0)) & "'")
58                If Not rsTemp.EOF Then
59                    Call objO.DeleteDataHO("DELETE ho_WarJobPart WHERE id_warranty = '" & Trim(Text(0)) & "'")
60                    Call objO.InsertDataHO("INSERT INTO ho_WarJobPart (ID_WARRANTY,MDL,MODEL,JOB,FS_ITEM,JOBKET,RATE,ITEMPART,NAMAPART,SATPART,QTY,FLAG,KETERANGAN) SELECT ID_WARRANTY,MDL,MODEL,JOB,FS_ITEM,JOBKET,RATE,ITEMPART,NAMAPART,SATPART,QTY,FLAG,KETERANGAN FROM tempho_WarJobPart WITH (NOLOCK) WHERE id_warranty = '" & Trim(Text(0)) & "' ORDER BY mdl, flag, job, fs_item, itempart")
61                    Call objO.DeleteDataHO("DElETE tempho_WarJobPart WHERE id_warranty = '" & Trim(Text(0)) & "'")
62                Else
63                    Call objO.DeleteDataHO("DELETE ho_WarJobPart WHERE id_warranty = '" & Trim(Text(0)) & "'")
64                End If
65                rsTemp.Close
                  
66                Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('WARRANTY', '" & Text(0) & "', 'S')")
67                Call ber_sih(1)
68                Call buka_Tutup(1, False)
69                Call refreshGrid(1)
                  
70                cmdSimpan.Caption = "&Tambah"
71                cmdUbah.Caption = "&Ubah"
72            Else
73                MsgBox pesan, vbCritical, "KELENGKAPAN"
74                Exit Sub
75            End If
76        End If
End Sub

Private Sub cmdTutup_Click()
    frmeUnit.Visible = False
    cmdAddUnit.Caption = "&Tambah"
    cmdUbahUnit.Caption = "&Ubah"
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        newRecord = False
        Call buka_Tutup(1, True)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(1, False)
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
    End If
End Sub

Private Sub cmdUbahUnit_Click()
    If cmdUbahUnit.Caption = "&Ubah" Then
        Call buka_Tutup(2, True)
        cmdUbahUnit.Caption = "&Batal"
        cmdAddUnit.Caption = "&Simpan"
    ElseIf cmdUbahUnit.Caption = "&Batal" Then
        Call ber_sih(2)
        Call buka_Tutup(2, False)
        cmdUbahUnit.Caption = "&Ubah"
        cmdAddUnit.Caption = "&Tambah"
    End If
End Sub

Private Sub cmdUnit_Click()
    Call ber_sih(2)
    Call refreshGrid(2)
    Call isiCombo
    frmeUnit.Top = cmdUnit.Top + cmdUnit.Height  '1680
    frmeUnit.Left = cmdUnit.Left + cmdUnit.Width - frmeUnit.Width  '1680
    frmeUnit.Visible = True
End Sub

Private Sub delJob_Click()
'    If boleh Then
        If chJob.value = Checked Then
            Set rsTemp = objO.FetchDataHO("SELECT * FROM tempHO_warjobpart WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "' AND mdl = '" & Text(13) & "' AND job = '" & Text(15) & "' AND fs_item = '" & Text(16) & "'")
            sSQL = "DELETE tempHO_WARJOBPART WHERE id_warranty = '" & Text(0) & "' AND mdl = '" & Text(13) & "' AND job = '" & Text(15) & "' AND fs_item = '" & Text(16) & "'"
        Else
            Set rsTemp = objO.FetchDataHO("SELECT * FROM tempHO_warjobpart WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "' AND mdl = '" & Text(13) & "' AND itempart = '" & Text(17) & "'")
            sSQL = "DELETE tempHO_WARJOBPART WHERE id_warranty = '" & Text(0) & "' AND mdl = '" & Text(13) & "' AND itempart = '" & Text(17) & "'"
        End If
        If Not rsTemp.EOF Then
            temp = MsgBox("Akan menghapus data ?", vbYesNo, "HAPUS DATA")
            If temp = vbYes Then
                Call objO.UpdateDataHO(sSQL)
            End If
        End If
        Call ber_sih(3)
        Call refreshGrid(3)
'    End If
End Sub

Private Sub editJob_Click()
    If editJob.Caption = "&Ubah" Then
        Call buka_Tutup(3, True)
        
        If chJob = Checked Then
            Text(15) = "TWC"
            Text(16).BackColor = &H80000005
        Else
            Text(17).BackColor = &H80000005
            Text(18).BackColor = &H80000005
            Text(19).BackColor = &H80000005
            Text(6).BackColor = &H80000005
        End If
        editJob.Caption = "&Batal"
        addJob.Caption = "&Simpan"
    ElseIf editJob.Caption = "&Batal" Then
        Call ber_sih(3)
        Call buka_Tutup(3, False)
        Text(15).BackColor = &HC0E0FF
        Text(16).BackColor = &HC0E0FF
        Text(17).BackColor = &HC0E0FF
        Text(18).BackColor = &HC0E0FF
        Text(19).BackColor = &HC0E0FF
        Text(6).BackColor = &HC0E0FF
        editJob.Caption = "&Ubah"
        addJob.Caption = "&Tambah"
    End If
End Sub

Private Sub Form_Load()
    Call ber_sih(1)
    Call ber_sih(2)
    Call ber_sih(3)
    Call buka_Tutup(1, False)
    Call buka_Tutup(2, False)
    Call buka_Tutup(3, False)
    Call refreshGrid(1)
    For i = 0 To 17
        chPilih(i) = Checked
    Next
    
    TglMulai = Format(isnow, "dd/mm/yyyy")
    TglSelesai = Format(isnow, "dd/mm/yyyy")
    TglTI = Format(isnow, "dd/mm/yyyy")
    frmeUnit.Visible = False
    frmeJob.Visible = False
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridWarranty.Width = Frame1.Width - 200
    GridWarranty.Height = Frame1.Height - Text(5).Top - Text(5).Height - 1000
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    ProgBar.Top = cmdKeluar.Top
        
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    cmdImport.Left = cmdCetak.Left - cmdImport.Width - 100
    ProgBar.Left = cmdImport.Left - ProgBar.Width - 100
    
End Sub

Private Sub GridJob_FilterChange()
    On Error GoTo errHandler
    Set cols = GridJob.Columns
    Dim C As Integer
    
    C = GridJob.col
    GridJob.HoldFields
    rsGridJob.filter = getFilter()
    GridJob.col = C
    GridJob.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub GridJob_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If addJob.Caption = "&Tambah" Then
        If Not rsGridJob.EOF Then
            Text(13) = rsGridJob.Fields(0)      'GridJob.Columns(0)
            Text(14) = rsGridJob.Fields(1)      'GridJob.Columns(1)
            Text(15) = rsGridJob.Fields(2)      'GridJob.Columns(2)
            Text(16) = rsGridJob.Fields(3)      'GridJob.Columns(3)
            Text(5) = rsGridJob.Fields(4)       'GridJob.Columns(4)
            Text(20) = rsGridJob.Fields(5)      'GridJob.Columns(5)
            Text(17) = rsGridJob.Fields(6)      'GridJob.Columns(6)
            Text(18) = rsGridJob.Fields(7)      'GridJob.Columns(7)
            Text(19) = rsGridJob.Fields(8)      'GridJob.Columns(8)
            Text(6) = rsGridJob.Fields(10)      'GridJob.Columns(10)
            If Trim(rsGridJob.Fields(9)) = "J" Then
                chJob.value = Checked
                chPart.value = Unchecked
                Text(15).BackColor = &H80000005
                Text(16).BackColor = &H80000005
                
                Text(17).BackColor = &HC0E0FF
                Text(18).BackColor = &HC0E0FF
                Text(19).BackColor = &HC0E0FF
                Text(6).BackColor = &HC0E0FF
            ElseIf Trim(rsGridJob.Fields(9)) = "P" Then
                chPart.value = Checked
                chJob.value = Unchecked
                Text(15).BackColor = &HC0E0FF
                Text(16).BackColor = &HC0E0FF
                
                Text(17).BackColor = &H80000005
                Text(18).BackColor = &H80000005
                Text(19).BackColor = &H80000005
                Text(6).BackColor = &H80000005
            End If
        End If
    End If
End Sub

Private Sub GridUnit_Click()
    'GridUnit_RowColChange
End Sub

Private Sub GridUnit_FilterChange()
    On Error GoTo errHandler
    Set cols = GridUnit.Columns
    Dim C As Integer
    
    C = GridUnit.col
    GridUnit.HoldFields
    rsGridUnit.filter = getFilter()
    GridUnit.col = C
    GridUnit.EditActive = True
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

Private Sub GridUnit_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGridUnit.EOF Then
        Text(7) = rsGridUnit.Fields(1)      'GridUnit.Columns(1)
        Text(8) = rsGridUnit.Fields(2)      'GridUnit.Columns(2)
        cboCountry = Left(rsGridUnit.Fields(3), 3)
        Text(9) = Mid(rsGridUnit.Fields(3), 4, 6)
        Text(10) = Right(rsGridUnit.Fields(3), Len(rsGridUnit.Fields(3)) - 9)
        If rsGridUnit.Fields(4) = "" Then
            Text(11) = ""
            Text(12) = ""
        Else
            cboCountry1 = Left(rsGridUnit.Fields(4), 3)
            Text(11) = Mid(rsGridUnit.Fields(4), 4, 6)
            Text(12) = Right(rsGridUnit.Fields(4), Len(rsGridUnit.Fields(4)) - 9)
        End If
    End If
End Sub

Private Sub GridWarranty_FilterChange()
    On Error GoTo errHandler
    Set cols = GridWarranty.Columns
    Dim C As Integer
    
    C = GridWarranty.col
    GridWarranty.HoldFields
    rsGrid.filter = getFilter()
    GridWarranty.col = C
    GridWarranty.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub GridWarranty_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF Then
        Text(0) = rsGrid.Fields(0)      'GridWarranty.Columns(0)
        Text(1) = rsGrid.Fields(1)      'GridWarranty.Columns(1)
        Text(2) = rsGrid.Fields(2)      'GridWarranty.Columns(2)
        TglMulai = rsGrid.Fields(3)     'GridWarranty.Columns(3)
        TglSelesai = rsGrid.Fields(4)   'GridWarranty.Columns(4)
        Text(3) = Left(rsGrid.Fields(5), 6)
        If rsGrid.Fields(5) = "" Then
            Text(4) = ""
        Else
            Text(4) = Right(rsGrid.Fields(5), Len(rsGrid.Fields(5)) - 7)
        End If
        
        Tgl = rsGrid.Fields(6)
        Call objO.InsertDataHO("INSERT INTO tempho_warDtl SELECT ID_WARRANTY, MDL, MODEL, NO_RANGKA1, NO_RANGKA2 FROM ho_warDtl WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM tempho_wardtl WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "') AND id_warranty = '" & Text(0) & "' ORDER BY mdl, no_rangka1")
        Call objO.InsertDataHO("INSERT INTO tempho_warjobpart SELECT ID_WARRANTY,MDL,MODEL,JOB,FS_ITEM,JOBKET,RATE,ITEMPART,NAMAPART,SATPART,QTY,FLAG,KETERANGAN " _
                 & "FROM ho_warjobpart WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM tempho_warjobpart WITH (NOLOCK) WHERE id_warranty = '" & Text(0) & "') AND " _
                 & "id_warranty = '" & Text(0) & "' ORDER BY mdl, job, fs_item, itempart")
                 
        Call refreshGrid(2)
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
        
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index) = gMainKeyCari
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
                gMainKeyCari = Text(Index)
                gSearchOption = "HOWRT"
                frmCariHO.show (vbModal)
            Case 7, 13
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCariHO.show (vbModal)
            Case 15
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCariHO.show (vbModal)
            Case 16
                gMainKeyCari = Text(Index)
                gSearchOption = "JOBLIST"
                gBantuan = Trim(Text(13))
                gBantuan2 = Trim(Text(15))
                frmCariHO.show (vbModal)
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

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 1
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("Select * from HO_warranty WITH (NOLOCK) where warranty ='" & Text(Index) & "' and FS_LOG_DEL = ''")
                If Not oRs.EOF Then
                    Text(0) = oRs("ID_WARRANTY")
                    Text(2) = oRs("KETERANGAN")
                    TglMulai = Format(oRs("TGLMULAI"), "dd/mm/yyyy")
                    TglSelesai = Format(oRs("TGLSELESAI"), "dd/mm/yyyy")
                    Text(3) = Left(oRs("NOTI"), 6)
                    Text(4) = Right(oRs("NOTI"), Len(oRs("NOTI")) - 6)
                End If
                oRs.Close
                Call refreshGrid(2)
            End If
        Case 4
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("SELECT * FROM HO_WARRANTY WITH (NOLOCK) WHERE NOTI = '" & Text(3) & "-" & Text(Index) & "' AND ID_WARRANTY <> '" & Text(0) & "' and FS_LOG_DEL = ''")
                If Not oRs.EOF Then
                    MsgBox "No TI Sudah dipakai, Gunakan nomor lain !!!", vbCritical, "DUPLIKASI"
                    Text(Index) = ""
                    Text(Index).SetFocus
                    Exit Sub
                End If
                oRs.Close
            End If
        Case 7
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("Select * from model WITH (NOLOCK) where mdl = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCariHO.show (vbModal)
                Else
                    Text(8) = oRs("DESC")
                    Text(9) = oRs("NO_CHASSIS")
                    Text(11) = oRs("NO_CHASSIS")
                End If
                oRs.Close
            End If
        Case 13
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("SELECT * FROM MODEL WITH (NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCariHO.show (vbModal)
                    Text(Index).SetFocus
                    oRs.Close
                    Exit Sub
                Else
                    Text(14) = oRs("desc")
                End If
                oRs.Close
            End If
        Case 15
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("SELECT * FROM job WITH (NOLOCK) WHERE job = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCariHO.show (vbModal)
                End If
                oRs.Close
            End If
        Case 16
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("select * from HO_FLATE WITH (NOLOCK) where job = '" & Trim(Text(15)) & "' AND mdl = '" & Text(13) & "' AND fs_item = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOBLIST"
                    gBantuan = Trim(Text(15))
                    gBantuan2 = Trim(Text(13))
                    frmCariHO.show (vbModal)
                Else
                    Text(20) = CDbl(oRs("RATE"))
                    Text(5) = oRs("KETERANGAN")
                End If
                oRs.Close
            End If
    End Select
End Sub

