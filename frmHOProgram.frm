VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmHOProgram 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DAFTAR PROGRAM"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   12795
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   10080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   8535
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   12255
      Begin VB.CheckBox ckNotifikasi 
         BackColor       =   &H0068ECE6&
         Caption         =   "Payment Notifications"
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
         Left            =   3000
         TabIndex        =   79
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2175
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
         Left            =   6840
         TabIndex        =   76
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame frameKriteria 
         BackColor       =   &H0068ECE6&
         Caption         =   "                                    KRITERIA KHUSUS                                                  "
         Height          =   3015
         Left            =   13080
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   5655
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
            TabIndex        =   75
            Top             =   2160
            Visible         =   0   'False
            Width           =   1215
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
            TabIndex        =   74
            Top             =   1800
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
            TabIndex        =   73
            Top             =   1440
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
            TabIndex        =   72
            Top             =   1080
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
            TabIndex        =   71
            Top             =   720
            Width           =   1575
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
            TabIndex        =   70
            Top             =   360
            Width           =   1215
         End
         Begin VB.Frame framBook 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   2040
            TabIndex        =   53
            Top             =   360
            Width           =   3255
            Begin VB.OptionButton opBook 
               BackColor       =   &H0068ECE6&
               Caption         =   "All"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   56
               Top             =   0
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton opBook 
               BackColor       =   &H0068ECE6&
               Caption         =   "Non Booking"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   55
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opBook 
               BackColor       =   &H0068ECE6&
               Caption         =   "Booking"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   54
               Top             =   0
               Width           =   1095
            End
         End
         Begin AIFCmp1.asxPowerButton cmdTutup1 
            Height          =   360
            Left            =   0
            TabIndex        =   52
            Top             =   2640
            Width           =   5655
            _ExtentX        =   9975
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
            Left            =   2040
            TabIndex        =   57
            Top             =   1080
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
            Left            =   3360
            TabIndex        =   58
            Top             =   1080
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
            Left            =   2040
            TabIndex        =   60
            Top             =   720
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
            Left            =   3360
            TabIndex        =   61
            Top             =   720
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
            Left            =   2040
            TabIndex        =   63
            Top             =   1440
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
            Left            =   2040
            TabIndex        =   65
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   107544577
            CurrentDate     =   39853
         End
         Begin MSComCtl2.DTPicker tglLahir2 
            Height          =   255
            Left            =   3840
            TabIndex        =   66
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   107544577
            CurrentDate     =   39853
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   2040
            TabIndex        =   68
            Top             =   2115
            Visible         =   0   'False
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
            Left            =   4455
            TabIndex        =   69
            Top             =   2130
            Visible         =   0   'False
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
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
            Left            =   3495
            TabIndex        =   67
            Top             =   1830
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
            Left            =   2940
            TabIndex        =   64
            Top             =   1485
            Width           =   285
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
            Left            =   3000
            TabIndex        =   62
            Top             =   765
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
            Left            =   3000
            TabIndex        =   59
            Top             =   1125
            Width           =   225
         End
      End
      Begin VB.Frame frmeDealer 
         BackColor       =   &H0068ECE6&
         Caption         =   $"frmHOProgram.frx":0000
         Height          =   3615
         Left            =   7800
         TabIndex        =   44
         Top             =   1920
         Width           =   6615
         Begin TrueOleDBGrid70.TDBGrid GridDealer 
            Height          =   3015
            Left            =   120
            TabIndex        =   45
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
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
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
            _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
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
            TabIndex        =   46
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
      Begin VB.Frame frmeModel 
         BackColor       =   &H0068ECE6&
         Caption         =   $"frmHOProgram.frx":009C
         Height          =   3615
         Left            =   6240
         TabIndex        =   29
         Top             =   1680
         Width           =   9135
         Begin TrueOleDBGrid70.TDBGrid GridModel 
            Height          =   2895
            Left            =   120
            TabIndex        =   31
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
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
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
            TabIndex        =   32
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
      Begin VB.CheckBox chInv2 
         BackColor       =   &H0068ECE6&
         Caption         =   "Invoice Ke Dealer Penjual"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   1770
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0068ECE6&
         Caption         =   "                          DEALER                             "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6840
         TabIndex        =   47
         Top             =   720
         Width           =   3975
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
            Left            =   960
            TabIndex        =   11
            Top             =   960
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   7
            Left            =   930
            TabIndex        =   10
            Top             =   240
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
            Mask            =   ">CCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   930
            TabIndex        =   48
            Top             =   600
            Width           =   2865
            _ExtentX        =   5054
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
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berlaku"
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
            Left            =   255
            TabIndex        =   50
            Top             =   990
            Width           =   525
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dealer Pendaftar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -120
            TabIndex        =   49
            Top             =   285
            Width           =   960
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   1800
            Top             =   240
            Width           =   60
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0068ECE6&
         Height          =   975
         Left            =   6840
         TabIndex        =   33
         Top             =   2040
         Width           =   3975
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Disc Part"
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
            Left            =   2640
            TabIndex        =   42
            Top             =   600
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Kode Part"
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
            Left            =   2640
            TabIndex        =   41
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Status Part"
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
            Left            =   2640
            TabIndex        =   40
            Top             =   120
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Disc Jasa"
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
            Left            =   1440
            TabIndex        =   39
            Top             =   600
            Value           =   1  'Checked
            Width           =   1095
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
            Left            =   1440
            TabIndex        =   38
            Top             =   360
            Value           =   1  'Checked
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
            Left            =   1440
            TabIndex        =   37
            Top             =   120
            Value           =   1  'Checked
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
            TabIndex        =   36
            Top             =   600
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Program"
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
            TabIndex        =   35
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Kode"
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
            TabIndex        =   34
            Top             =   120
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.CheckBox chPart 
         BackColor       =   &H0068ECE6&
         Caption         =   "Part Khusus"
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
         Left            =   1680
         TabIndex        =   7
         Top             =   2055
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   8
         Top             =   2040
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   "999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   2865
         _ExtentX        =   5054
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
         Left            =   1680
         TabIndex        =   4
         Top             =   1400
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker TglMulai 
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   1100
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   111214593
         CurrentDate     =   39853
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   0
         Top             =   500
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
      Begin TrueOleDBGrid70.TDBGrid GridProgram 
         Height          =   5295
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9340
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
         Left            =   1680
         TabIndex        =   17
         Top             =   200
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
         Left            =   3480
         TabIndex        =   3
         Top             =   1100
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   111214593
         CurrentDate     =   39853
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   1
         Top             =   795
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
         Index           =   3
         Left            =   1680
         TabIndex        =   5
         Top             =   1755
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   "999,99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   9
         Top             =   2355
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Part"
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
         Left            =   525
         TabIndex        =   27
         Top             =   2085
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   2085
         Width           =   165
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   285
         TabIndex        =   30
         Top             =   1425
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3480
         Top             =   2355
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4560
         Top             =   507
         Width           =   60
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   2400
         TabIndex        =   25
         Top             =   1800
         Width           =   165
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Jasa"
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
         TabIndex        =   24
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Program"
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
         Left            =   525
         TabIndex        =   23
         Top             =   245
         Width           =   1005
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
         Left            =   3135
         TabIndex        =   22
         Top             =   1125
         Width           =   225
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
         Left            =   960
         TabIndex        =   21
         Top             =   1125
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Program"
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
         Left            =   480
         TabIndex        =   19
         Top             =   545
         Width           =   1050
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
         Left            =   690
         TabIndex        =   18
         Top             =   840
         Width           =   840
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
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   8160
      TabIndex        =   12
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
      Left            =   9120
      TabIndex        =   13
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
      Left            =   10080
      TabIndex        =   14
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
      Left            =   11040
      TabIndex        =   20
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
      Left            =   6840
      TabIndex        =   43
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
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   5280
      TabIndex        =   77
      Top             =   9000
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
      Left            =   1440
      TabIndex        =   78
      Top             =   9000
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
   Begin AIFCmp1.asxPowerButton cmdExport 
      Height          =   360
      Left            =   0
      TabIndex        =   80
      Top             =   9000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      BackColor       =   12648447
      Caption         =   "Export Excel"
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
End
Attribute VB_Name = "frmHOProgram"
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
Dim rsGridMdl As New ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim rsDealer As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns

Private Sub ber_sih()
    For i = 0 To 14
        Select Case i
            Case 3, 4, 9, 10, 11, 12, 13
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    
    frmeModel.Visible = False
    frmeDealer.Visible = False
    TglMulai = isnow   'Now()
    TglSelesai = isnow 'Now()
    chAllDealer = Unchecked
    chModel = Unchecked
    chPart = Unchecked
    chInv2 = Unchecked
    ckNotifikasi = Unchecked
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    For i = 1 To 8
        Text(i).Enabled = flg1
    Next
    TglMulai.Enabled = flg1
    TglSelesai.Enabled = flg1
    chPart.Enabled = flg1
    chModel.Enabled = flg1
    chAllDealer.Enabled = flg1
    chInv2.Enabled = flg1
    If Not flg1 Then
        frmeDealer.Visible = flg1
        frmeModel.Visible = flg1
    End If
    opBook(0).Enabled = Flg
    opBook(1).Enabled = Flg
    opBook(2).Enabled = Flg
    Text(9).Enabled = Flg
    Text(10).Enabled = Flg
    Text(11).Enabled = Flg
    Text(12).Enabled = Flg
    Text(13).Enabled = Flg
    tglLahir1.Enabled = Flg
    tglLahir2.Enabled = Flg
    Text(14).Enabled = Flg
End Sub

Private Function isiIdExcell() As String
    Dim noUrut As Integer
    Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_PROGRAM WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY id_program DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("id_program"), Len(rsTemp("id_program")) - 4))
        noUrut = noUrut + 1
        isiIdExcell = Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        isiIdExcell = Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    End If
End Function

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_PROGRAM WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY id_program DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("id_program"), Len(rsTemp("id_program")) - 4))
        noUrut = noUrut + 1
        Text(0) = Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        Text(0) = Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    End If
End Sub

Private Sub createTbl()
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    With rsGridMdl
        .Fields.Append "GRPMDL", adVarChar, 15
        .Fields.Append "KODEMDL", adVarChar, 20
        .Fields.Append "NAMAMDL", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridModel.DataSource = rsGridMdl
    GridModel.AlternatingRowStyle = True
End Sub

Private Sub isiModel()
'''    'If rsTemp.State Then rsTemp.Close
'''    Set rsTemp = objO.FetchDataHO("SELECT a.fs_grp_mdl, a.mdl, a.[desc], ISNULL(id_program, '') as PRO FROM model a LEFT JOIN ho_PROmdl b ON a.MDL = b.MDL ORDER BY a.fs_grp_mdl, a.mdl")
'''    If Not rsTemp.EOF Then
'''        Call createTbl
'''        Do While Not rsTemp.EOF
'''            rsGridMdl.AddNew
'''            rsGridMdl!GRPMDL = Trim(rsTemp("FS_GRP_MDL"))
'''            rsGridMdl!kodemdl = Trim(rsTemp("MDL"))
'''            rsGridMdl!NAMAMDL = Trim(rsTemp("DESC"))
'''            rsGridMdl!ikut = IIf(rsTemp("PRO") = "", False, True)
'''            rsGridMdl.Update
'''            rsTemp.MoveNext
'''        Loop
'''
'''        rsGridMdl.MoveFirst
'''        GridModel.Columns(0).Width = 1200
'''        GridModel.Columns(1).Width = 2000
'''        GridModel.Columns(2).Width = 4000
'''        GridModel.Columns(3).Alignment = dbgCenter
'''        GridModel.Columns(3).Width = 800
'''    End If
    
    'If rsTemp.State Then rsTemp.Close
    Set rsTemp = objO.FetchDataHO("SELECT DISTINCT a.MODEL, a.mdl, a.[desc], ISNULL(id_program, '') as PRO FROM model a WITH (NOLOCK) LEFT JOIN ho_PROmdl b WITH (NOLOCK) ON a.MDL = b.MDL AND B.ID_PROGRAM = '" & Trim(Text(0)) & "' ORDER BY a.MODEL, a.mdl")
    If Not rsTemp.EOF Then
        Call createTbl
        Do While Not rsTemp.EOF
            rsGridMdl.AddNew
            rsGridMdl!grpMDL = Trim(rsTemp("MODEL"))
            rsGridMdl!kodeMDL = Trim(rsTemp("MDL"))
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
End Sub

Private Sub refreshGrid()
    GridProgram.ClearFields
        Set rsGrid = objO.FetchDataHO("SELECT id_program as KODE, PROGRAM, KETERANGAN, TglMulai as [MULAI BERLAKU], tglSelesai as [SELESAI BERLAKU], discJasa as [DISC JASA], flagpart as [PART BERLAKU], kodePART as [KODE PART], discPart as [DISC PART], dealerdaft AS [DEALER PENDAFTAR], NAMADEALER AS [NAMA DEALER], FLAGDEALER AS [DEALER BERLAKU], FLAGMODEL AS MODEL, DOUBLEINV AS INVOICE,flagNotifWA_VA FROM ho_program WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY id_program")
        If Not rsGrid.EOF Then
            GridProgram.DataSource = rsGrid
            Call HeaderColom
        End If
End Sub

Private Sub HeaderColom()
    With GridProgram
        .Columns(0).Caption = "Kode"
        .Columns(0).Width = IIf(chPilih(0) = Checked, 1400, 0)
        .Columns(1).Caption = "Program"
        .Columns(1).Width = IIf(chPilih(1) = Checked, 3500, 0)
        .Columns(2).Caption = "Keterangan"
        .Columns(2).Width = IIf(chPilih(2) = Checked, 5000, 0)
        .Columns(3).Caption = "Mulai Berlaku"
        .Columns(3).Width = IIf(chPilih(3) = Checked, 1400, 0)
        .Columns(4).Caption = "Selesai Berlaku"
        .Columns(4).Width = IIf(chPilih(4) = Checked, 1400, 0)
        .Columns(5).Caption = "Jasa"
        .Columns(5).Width = IIf(chPilih(5) = Checked, 800, 0)
        .Columns(6).Caption = "Part Berlaku"
        .Columns(6).Width = IIf(chPilih(6) = Checked, 300, 0)
        .Columns(7).Caption = "Kode Part"
        .Columns(7).Width = IIf(chPilih(7) = Checked, 2000, 0)
        .Columns(8).Caption = "Part"
        .Columns(8).Width = IIf(chPilih(8) = Checked, 800, 0)
        .Columns(9).Caption = "Kode Dealer"
        .Columns(9).Width = 800
        .Columns(10).Caption = "Nama Dealer"
        .Columns(10).Width = 3000
        .Columns(11).Caption = "Dealer"
        .Columns(11).Width = 300
        .Columns(12).Caption = "Model"
        .Columns(12).Width = 300
        .Columns(13).Caption = "Invoice"
        .Columns(13).Width = 300
        .Columns(14).Caption = "Notif WA VA"
        .Columns(14).Width = 300
    End With
End Sub

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

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 3500
    GridDealer.Columns(2).Width = 1000
    GridDealer.Columns(2).Alignment = dbgCenter
End Sub

Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
    Set rsTmp = objO.FetchDataHO("SELECT fs_kd_dealer, fs_nm_dealer, ISNULL(ID_PROGRAM,'') AS KODE FROM t_dealer a WITH (NOLOCK) " _
                            & "LEFT JOIN HO_PROCBG b ON a.fs_kd_dealer = b.kodedealer and A.FS_LOG_DEL = '' AND b.ID_PROGRAM = '" & Text(0) & "' WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
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


Private Sub chAllDealer_Click()

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
End Sub

Private Sub chAllDealer_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub chInv2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
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
                Text(9).Enabled = True
                Text(10).Enabled = True
            Case 2
                Text(11).Enabled = True
                Text(12).Enabled = True
            Case 3
                Text(13).Enabled = True
            Case 4
                tglLahir1.Enabled = True
                tglLahir2.Enabled = True
            Case 5
                Text(14).Enabled = True
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
                Text(9).Enabled = False
                Text(10).Enabled = False
                Text(9) = 0
                Text(10) = 0
            Case 2
                Text(11).Enabled = False
                Text(12).Enabled = False
                Text(11) = 0
                Text(12) = 0
            Case 3
                Text(13).Enabled = False
                Text(13) = 0
            Case 4
                tglLahir1.Enabled = False
                tglLahir2.Enabled = False
                tglLahir1 = Format(isnow, "DD/MM/YYYY")
                tglLahir2 = Format(isnow, "DD/MM/YYYY")
            Case 5
                Text(14).Enabled = False
                Text(14) = ""
        End Select
    End If
End Sub

Private Sub chKrti_Click()
    If chKrti = Checked Then
        frameKriteria.Visible = True
    Else
        frameKriteria.Visible = False
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

Private Sub chModel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub chPart_Click()
    If chPart.value = Checked Then
        Text(5).Visible = True
        Text(6).Visible = True
        Shape1.Visible = True
    Else
        Text(5).Visible = False
        Text(6).Visible = False
        Shape1.Visible = False
    End If
End Sub

Private Sub chPart_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub chPilih_Click(Index As Integer)
    Select Case Index
        Case 0
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1500, 0)
        Case 1
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2500, 0)
        Case 2
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 3000, 0)
        Case 3
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 4
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 5
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 800, 0)
        Case 6
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 800, 0)
        Case 7
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 8
            GridProgram.Columns(Index).Width = IIf(chPilih(Index) = Checked, 800, 0)
    End Select
End Sub

Private Sub cmdCetak_Click()
    frmHOProgramCtk.show
End Sub

Private Sub cmdHapus_Click()
    'If rsTemp.State Then rsTemp.Close
    Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_program WITH (NOLOCK) WHERE id_program = '" & Text(0) & "' and FS_LOG_DEL = ''")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus Program : " & Text(1), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE ho_program SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE id_program = '" & Text(0) & "' and FS_LOG_DEL = ''")
            Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('PROGRAM', '" & Text(0) & "','D')")
        End If
    Else
        MsgBox "Pilih data yang akan dihapus"
        Exit Sub
    End If
    Call ber_sih
    Call refreshGrid
End Sub

Private Sub cmdImport_Click()
        Dim ErrBegin As Boolean, konter As Long
1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim rsTempExcel As ADODB.Recordset
         Dim Jml As Integer
        ' Dim konter As Integer
         Dim idProgram As String
         Dim namaProgram As String
         Dim isInsert As Boolean
2       konter = 0
3       ProgBar.Min = 0
4       CommonDialog1.FileName = ""
5       CommonDialog1.InitDir = App.Path
6       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
7       CommonDialog1.ShowOpen
8       If CommonDialog1.FileName <> "" Then
9          If CnExel.State Then CnExel.Close
10         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
11         GoSub Proses
12     Else
13        Exit Sub
14     End If
15     Exit Sub
       
Proses:
16    ErrBegin = True
         Dim VarNmFile As String
         Dim namaDealer As String
         Dim namaDealerCbg As String
         Dim model As String
         Dim grpMDL As String
         Dim desc As String
         
17       Call objHT.OpenTransactionHO
18        gBantuan = "SBE"
19       If RcExcel.State Then RcExcel.Close
20       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
21       If RcExcel.RecordCount > 0 Then
22          ProgBar.Visible = True
23          ProgBar.Max = RcExcel.RecordCount
24          RcExcel.MoveFirst
25          idProgram = ""
26          namaProgram = ""
            
27          Do Until RcExcel.EOF
28             DoEvents
29             konter = konter + 1
30             ProgBar.value = konter
               
31             namaDealer = ""
32             model = ""
33             grpMDL = ""
34             desc = ""
35             isInsert = False
               
36             If Nul(Trim(RcExcel.Fields(0).value)) = "" Then Exit Do
37             If namaProgram <> Nul(Trim(RcExcel.Fields(0).value)) Then
38                idProgram = isiIdExcell
39                namaProgram = Nul(Trim(RcExcel.Fields(0).value))
40                isInsert = True
41             End If
42            If Nul(Trim(RcExcel.Fields(0).value)) = "" Then
43                Call objHT.CloseTransactionHO_err
44                ProgBar.Visible = False
45                MsgBox "Nama Program baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
46                Exit Sub
47            End If
48            If Nul(Trim(RcExcel.Fields(1).value)) = "" Then
49                Call objHT.CloseTransactionHO_err
50                ProgBar.Visible = False
51                MsgBox "Keterangan baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
52                Exit Sub
53            End If
54            If Nul(Trim(RcExcel.Fields(11).value)) = "" Then
55                If Nul(Trim(RcExcel.Fields(12).value)) = "" Then
56                    Call objHT.CloseTransactionHO_err
57                    ProgBar.Visible = False
58                    MsgBox "Dealer Pendaftar. Minimal HO. baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
59                    Exit Sub
60                End If
61            Else
62                Set rsTempExcel = objO.FetchDataHO("select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1 AND FS_KD_DEALER = '" & Trim(RcExcel.Fields(11).value) & "'")
63                If rsTempExcel.RecordCount > 0 Then
64                    namaDealer = Trim(rsTempExcel("FS_NM_DEALER"))
65                Else
66                    Call objHT.CloseTransactionHO_err
67                    ProgBar.Visible = False
68                    MsgBox "Kode dealer " & Trim(RcExcel.Fields(11).value) & " baris ke-" & konter & " salah / tidak ada data !", vbCritical, "Error"
69                    Exit Sub
70                End If
71            End If
72            If Nul(Trim(RcExcel.Fields(4).value)) = "" Then
73                If Nul(Trim(RcExcel.Fields(5).value)) = "" Then
74                    Call objHT.CloseTransactionHO_err
75                    ProgBar.Visible = False
76                    MsgBox "MODEL baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
77                    Exit Sub
78                Else
79                    Set rsTempExcel = objO.FetchDataHO("SELECT * FROM MODEL WHERE FS_LOG_DEL = '' AND MDL = '" & Trim(RcExcel.Fields(5).value) & "'")
80                    If rsTempExcel.RecordCount > 0 Then
81                        model = Trim(rsTempExcel("mdl"))
82                        desc = Trim(rsTempExcel("desc"))
83                        grpMDL = Trim(rsTempExcel("MODEL"))
84                    Else
                          'jika tidak ada data
85                        Call objHT.CloseTransactionHO_err
86                        ProgBar.Visible = False
87                        MsgBox "MODEL baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
88                        Exit Sub
89                    End If
90                End If
91            Else
                  
92            End If
93            If Nul(RcExcel.Fields(12).value) = "" Then
94                Call objHT.CloseTransactionHO_err
95                ProgBar.Visible = False
96                MsgBox "Kode cabang kolom ke 13 baris ke-" & konter & vbCrLf & "Isikan 1 untuk semua dealer, 0 untuk hanya dealer pendaftar, atau isikan kode dealer yang lain.", vbCritical, "Error"
97                Exit Sub
98            Else
99                If Nul(RcExcel.Fields(12).value) <> "1" Or Nul(RcExcel.Fields(12).value) <> "0" Then
100                   Set rsTempExcel = objO.FetchDataHO("select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1 AND FS_KD_DEALER = '" & Trim(RcExcel.Fields(12).value) & "'")
101                   If rsTempExcel.RecordCount > 0 Then
102                       namaDealerCbg = Trim(rsTempExcel("FS_NM_DEALER"))
103                   Else
104                       Call objHT.CloseTransactionHO_err
105                       ProgBar.Visible = False
106                       MsgBox "Kode dealer " & Trim(RcExcel.Fields(12).value) & " baris ke-" & konter & " tidak terdaftar !", vbCritical, "Error"
107                       Exit Sub
108                   End If
109               End If
110           End If
              
              
111               Call ClearsSql
112               gsSql1(0) = "ID_PROGRAM"
113               gsSql2(0) = idProgram
114               gsSql1(1) = "PROGRAM"
115               gsSql2(1) = Trim(RcExcel.Fields(0).value)
116               gsSql1(2) = "KETERANGAN"
117               gsSql2(2) = Trim(RcExcel.Fields(1).value)
118               gsSql1(3) = "TGLMULAI"
119               gsSql2(3) = Format(Trim(RcExcel.Fields(2).value), "mm/dd/yyyy")
120               gsSql1(4) = "TGLSELESAI"
121               gsSql2(4) = Format(Trim(RcExcel.Fields(3).value), "mm/dd/yyyy")
122               gsSql1(5) = "DISCJASA"
123               gsSql2(5) = Trim(RcExcel.Fields(6).value)
124               gsSql1(6) = "DISCPART"
125               gsSql2(6) = Trim(RcExcel.Fields(9).value)
126               gsSql1(7) = "DEALERDAFT"
127               gsSql2(7) = Trim(RcExcel.Fields(11).value)
128               gsSql1(8) = "NAMADEALER"
129               gsSql2(8) = namaDealer
130               gsSql1(9) = "FLAGDEALER"
131               gsSql2(9) = IIf(Trim(RcExcel.Fields(12).value) <> "", "A", "V")
132               gsSql1(10) = "DOUBLEINV"
133               gsSql2(10) = IIf(Trim(RcExcel.Fields(7).value) <> "", 1, 0)
134               gsSql1(11) = "FLAGMODEL"
135               gsSql2(11) = IIf(Trim(RcExcel.Fields(4).value) <> "", "A", "V")
136               gsSql1(12) = "FLAGPART"
137               If Trim(RcExcel.Fields(7).value) <> "" Then
138                   gsSql2(12) = "A"
139               Else
140                   gsSql2(12) = "K"
141                   gsSql1(13) = "KODEPART"
142                   gsSql2(13) = Trim(RcExcel.Fields(10).value)
143               End If
                  
                      ' DATA BARU
144               If isInsert Then
145                   sSQL = GensSql2("HO_PROGRAM", False)
146                   Call objHT.InsertData_Ho(sSQL)
147               End If
                  ' SIMPAN MODEL
148               If Nul(Trim(RcExcel.Fields(4).value)) = "" Then
149                   Call objHT.InsertData_Ho("DELETE ho_ProMdl WHERE id_program = '" & gsSql2(0) & "' and mdl = '" & model & "'")
150                   Call objHT.InsertData_Ho("INSERT INTO ho_promdl (ID_PROGRAM,GRPMDL,MDL,[DESC]) values ('" & gsSql2(0) & "','" & grpMDL & "','" & model & "','" & desc & "')")
151               Else
152                   Set rsTempExcel = objO.FetchDataHO("SELECT * FROM MODEL WHERE FS_LOG_DEL = ''")
153                   If rsTempExcel.RecordCount > 0 Then
154                       rsTempExcel.MoveFirst
155                       Call objHT.InsertData_Ho("DELETE ho_ProMdl WHERE id_program = '" & gsSql2(0) & "'")
156                       Do While Not rsTempExcel.EOF
157                           Call objHT.InsertData_Ho("INSERT INTO ho_promdl (ID_PROGRAM,GRPMDL,MDL,[DESC]) values ('" & gsSql2(0) & "','" & rsTempExcel("model") & "','" & rsTempExcel("mdl") & "','" & rsTempExcel("desc") & "')")
158                           rsTempExcel.MoveNext
159                       Loop
160                   End If
161               End If
                  
                  ' SIMPAN DEALER
162               If Nul(Trim(RcExcel.Fields(12).value)) = "1" Then
                      ' jika semua dealer insert
163                   Set rsTempExcel = objO.FetchDataHO("select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1")
164                   If rsTempExcel.RecordCount > 0 Then
165                       rsTempExcel.MoveFirst
166                       Call objHT.InsertData_Ho("DELETE HO_PROCBG WHERE ID_PROGRAM = '" & gsSql2(0) & "'")
167                       Do While Not rsTempExcel.EOF
168                           Call objHT.InsertData_Ho("INSERT INTO HO_PROCBG (ID_PROGRAM,KODEDEALER,NAMADEALER) VALUES ('" & gsSql2(0) & "', '" & Trim(rsTempExcel("FS_KD_DEALER")) & "', '" & Trim(rsTempExcel("FS_nm_DEALER")) & "')")
169                           rsTempExcel.MoveNext
170                       Loop
171                   End If
172               Else
173                   If Nul(Trim(RcExcel.Fields(12).value)) = "0" Then
174                       Call objHT.InsertData_Ho("DELETE HO_PROCBG WHERE ID_PROGRAM = '" & gsSql2(0) & "' and kodedealer = '" & Trim(RcExcel.Fields(11).value) & "'")
175                       Call objHT.InsertData_Ho("INSERT INTO HO_PROCBG (ID_PROGRAM,KODEDEALER,NAMADEALER) VALUES ('" & gsSql2(0) & "', '" & Trim(RcExcel.Fields(11).value) & "', '" & namaDealer & "')")
176                   Else
177                       Call objHT.InsertData_Ho("DELETE HO_PROCBG WHERE ID_PROGRAM = '" & gsSql2(0) & "' and kodedealer = '" & Trim(RcExcel.Fields(12).value) & "'")
178                       Call objHT.InsertData_Ho("INSERT INTO HO_PROCBG  (ID_PROGRAM,KODEDEALER,NAMADEALER) VALUES ('" & gsSql2(0) & "', '" & Trim(RcExcel.Fields(12).value) & "', '" & namaDealerCbg & "')")
179                   End If
                      
180               End If
                   
181               Call objHT.InsertData_Ho("INSERT INTO TEMPKIRIM VALUES ('PROGRAM', '" & gsSql2(0) & "','S')")
                  
182            RcExcel.MoveNext
183         Loop
184      End If
185      RcExcel.Close
186      CnExel.Close
187      ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
188       Call objHT.CloseTransactionHO
189       Call refreshGrid
190      MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
      '   GridAFI.Refresh
191      Return
192      Exit Sub
vb_error:
193      If ErrBegin = True Then
           ' DbCon.RollbackTrans
194      End If
195      ProgBar.Visible = False
196      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String

1         If cmdSimpan.Caption = "&Tambah" Then
2             Call ber_sih
3             Call buka_Tutup(False, True)
              
4             Text(1).SetFocus
5             cmdSimpan.Caption = "&Simpan"
6             cmdUbah.Caption = "&Batal"
7             newRecord = True
8         ElseIf cmdSimpan.Caption = "&Simpan" Then
          
9             boleh = True
10            pesan = "Lengkapi Data Berikut "
              
11            If Text(1) = "" Then
12                boleh = False
13                pesan = pesan + "Nama Program. "
14            End If
15            If Text(2) = "" Then
16                boleh = False
17                pesan = pesan + "Keterangan. "
18            End If
19            If chPart = Checked Then
20                If Text(5) = "" Then
21                    boleh = False
22                    pesan = pesan + "Part. "
23                End If
24            End If
25            If Text(7) = "" Then
26                boleh = False
27                pesan = pesan + "Dealer Pendaftar. Minimal HO. "
28            End If
29            If chAllDealer = Unchecked Then
30                If rsDealer.State <> adStateClosed Then
31                    rsDealer.MoveFirst
32                    rsDealer.Find "TERLIBAT = '" & True & "'", adSearchForward, 1
33                    If rsDealer.EOF Then
34                        boleh = False
35                        pesan = pesan + "Dealer belum dipilih. "
36                    End If
37                Else
38                    boleh = False
39                    pesan = pesan + "Dealer belum dipilih. "
40                End If
41            End If
              
42            If boleh Then
43                If Text(0) = "" Then
44                    Call isiId
45                End If
                  
46                Call ClearsSql
47                gsSql1(0) = "ID_PROGRAM"
48                gsSql2(0) = Text(0)
49                gsSql1(1) = "PROGRAM"
50                gsSql2(1) = Text(1)
51                gsSql1(2) = "KETERANGAN"
52                gsSql2(2) = Text(2)
53                gsSql1(3) = "TGLMULAI"
54                gsSql2(3) = Format(TglMulai, "mm/dd/yyyy")
55                gsSql1(4) = "TGLSELESAI"
56                gsSql2(4) = Format(TglSelesai, "mm/dd/yyyy")
57                gsSql1(5) = "DISCJASA"
58                gsSql2(5) = Text(3)
59                gsSql1(6) = "DISCPART"
60                gsSql2(6) = Text(4)
61                gsSql1(7) = "DEALERDAFT"
62                gsSql2(7) = Text(7)
63                gsSql1(8) = "NAMADEALER"
64                gsSql2(8) = Text(8)
65                gsSql1(9) = "FLAGDEALER"
66                gsSql2(9) = IIf(chAllDealer = Checked, "A", "V")
67                gsSql1(10) = "DOUBLEINV"
68                gsSql2(10) = IIf(chInv2 = Checked, 1, 0)
69                gsSql1(11) = "FLAGMODEL"
70                gsSql2(11) = IIf(chModel = Checked, "A", "V")
71                gsSql1(12) = "FLAGPART"
72                If chPart.value = Unchecked Then
73                    gsSql2(12) = "A"
74                    gsSql1(13) = "KODEPART"
75                    gsSql2(13) = ""
76                    gsSql1(14) = "NamaPart"
77                    gsSql2(14) = ""
78                    gsSql1(15) = "fs_log_del"
79                    gsSql2(15) = ""
80                    gsSql1(16) = "fd_tgl_del"
81                    gsSql2(16) = "01/01/1900"
                      
82                Else
83                    gsSql2(12) = "K"
84                    gsSql1(13) = "KODEPART"
85                    gsSql2(13) = Text(5)
86                    gsSql1(14) = "NamaPart"
87                    gsSql2(14) = ""
88                    gsSql1(15) = "fs_log_del"
89                    gsSql2(15) = ""
90                    gsSql1(16) = "fd_tgl_del"
91                    gsSql2(16) = "01/01/1900"
92                End If
                  
93                gsSql1(17) = "flagNotifWA_VA"
94                gsSql2(17) = IIf(ckNotifikasi = Checked, "1", "0")
          
                  ' SIMPAN DATA
95                Set rsTemp = objO.FetchDataHO("SELECT * FROM Ho_Program WITH (NOLOCK) WHERE id_program = '" & Text(0) & "' and FS_LOG_DEL = ''")
96                If Not rsTemp.EOF Then
                      ' DATA SUDAH ADA
97                    If newRecord Then
98                        MsgBox "DATA SUDAH ADA. Jika ingin merubah pilih UBAH", vbCritical, "DUPLIKASI DATA"
99                    Else
100                       temp = MsgBox("Yakin akan merubah Data ?", vbYesNo, "DUPLIKASI DATA")
101                       If temp = vbYes Then
102                           sSQL = GensSql2("HO_PROGRAM", True) + " where id_program = '" & Text(0) & "' and FS_LOG_DEL = ''"
103                       End If
104                   End If
105               Else
                      ' DATA BARU
106                   sSQL = GensSql2("HO_PROGRAM", False)
107               End If
108               Call objO.UpdateDataHO(sSQL)
                  
                  ' SIMPAN MODEL
109               If rsGridMdl.State Then
110                   rsGridMdl.MoveFirst
111                   Call objO.DeleteDataHO("DELETE ho_ProMdl WHERE id_program = '" & Text(0) & "'")
112                   Do While Not rsGridMdl.EOF
113                       If rsGridMdl("IKUT") Then
114                           Call objO.InsertDataHO("INSERT INTO ho_promdl (ID_PROGRAM,GRPMDL,MDL,[DESC]) values ('" & Text(0) & "','" & rsGridMdl("GRPMDL") & "','" & rsGridMdl("KODEMDL") & "','" & rsGridMdl("NAMAMDL") & "')")
115                       End If
116                       rsGridMdl.MoveNext
117                   Loop
118               End If
                  
                  ' SIMPAN DEALER
119               If rsDealer.State Then
120                   rsDealer.MoveFirst
121                   Call objO.DeleteDataHO("DELETE HO_PROCBG WHERE ID_PROGRAM = '" & Text(0) & "'")
122                   Do While Not rsDealer.EOF
123                       If rsDealer("TERLIBAT") Then
124                           Call objO.InsertDataHO("INSERT INTO HO_PROCBG (ID_PROGRAM,KODEDEALER,NAMADEALER) VALUES ('" & Text(0) & "', '" & rsDealer("KODEDEALER") & "', '" & rsDealer("NAMADEALER") & "')")
125                       End If
126                       rsDealer.MoveNext
127                   Loop
128               End If
                  
129               Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('PROGRAM', '" & Text(0) & "','S')")
                  
130               Call ber_sih
131               Call buka_Tutup(False, False)
132               Call refreshGrid
                  
133               cmdSimpan.Caption = "&Tambah"
134               cmdUbah.Caption = "&Ubah"
135           Else
136               MsgBox pesan, vbCritical, "KELENGKAPAN"
137               Exit Sub
138           End If
139       End If
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
        Call buka_Tutup(False, True)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
    End If
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridProgram.Width = Frame1.Width - 200
    GridProgram.Height = Frame1.Height - Text(6).Top - Text(6).Height - 300
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    cmdExport.Top = cmdKeluar.Top
    ProgBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    cmdImport.Left = cmdCetak.Left - cmdImport.Width - 100
    cmdExport.Left = cmdImport.Left - cmdExport.Width - 100
    ProgBar.Left = cmdImport.Left - ProgBar.Width - 100
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

Private Sub GridProgram_FilterChange()
    On Error GoTo errHandler
    Set cols = GridProgram.Columns
    Dim C As Integer
    
    C = GridProgram.col
    GridProgram.HoldFields
    rsGrid.filter = getFilter()
    GridProgram.col = C
    GridProgram.EditActive = True
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

Private Sub GridProgram_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF Then
        Text(0) = rsGrid.Fields(0)      'GridProgram.Columns(0)
        Text(1) = rsGrid.Fields(1)      'GridProgram.Columns(1)
        Text(2) = rsGrid.Fields(2)      'GridProgram.Columns(2)
        TglMulai = Format(rsGrid.Fields(3), "dd/mm/yyyy")
        TglSelesai = Format(rsGrid.Fields(4), "dd/mm/yyyy")
        Text(3) = rsGrid.Fields(5)      'GridProgram.Columns(5)
        If Trim(rsGrid.Fields(6)) = "A" Then
            chPart.value = Unchecked
        Else
            chPart.value = Checked
            Text(5) = rsGrid.Fields(6)  'GridProgram.Columns(7)
        End If
        Text(4) = rsGrid.Fields(8)      'GridProgram.Columns(8)
        Text(7) = rsGrid.Fields(9)      'GridProgram.Columns(9)
        Text(8) = rsGrid.Fields(10)     'GridProgram.Columns(10)
        chAllDealer = IIf(rsGrid.Fields(11) = "A", Checked, Unchecked)
        chModel = IIf(rsGrid.Fields(12) = "A", Checked, Unchecked)
        chInv2 = IIf(rsGrid.Fields(13) = 1, Checked, Unchecked)
        ckNotifikasi = IIf(rsGrid.Fields(14) = 1, Checked, Unchecked)
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
                gSearchOption = "HOPRG"
                frmCariHO.show (vbModal)
            Case 5
                gMainKeyCari = Text(Index)
                gSearchOption = "PART"
                frmCariAcess.show (vbModal)
            Case 7
                gMainKeyCari = Text(Index)
                gBantuan = "INTERNAL"
                gSearchOption = "DEALER"
                frmCariHO.show (vbModal)
            Case 14
                gMainKeyCari = Text(Index)
                gSearchOption = "HOBI"
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

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 3, 4
            If Text(3) = "" Then Text(3) = 0
            If Text(Index) > 100 Then
                Text(Index) = 100
                Text(Index).SelLength = Len(Text(Index))
                Text(Index).SetFocus
            ElseIf Text(Index) < 0 Then
                Text(Index) = 0
                Text(Index).SelLength = Len(Text(Index))
                Text(Index).SetFocus
            End If
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 1
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("Select * from HO_PROGRAM WITH (NOLOCK) where id_program ='" & Text(0) & "' and FS_LOG_DEL = ''")
                If Not oRs.EOF Then
                    Text(1) = oRs("program")
                    Text(2) = oRs("keterangan")
                    Text(3) = oRs("discJasa")
                    Text(4) = oRs("discPart")
                    TglMulai = Format(oRs("TGLMULAI"), "DD/MM/YYYY")
                    TglSelesai = Format(oRs("TglSelesai"), "DD/MM/YYYY")
                    If Trim(oRs("FLAGPART") = "A") Then
                        chPart.value = Unchecked
                        Text(5).Visible = False
                        Text(6).Visible = False
                        Shape1.Visible = False
                    Else
                        chPart.value = Checked
                        Text(5) = oRs("KODEPART")
                        Text(5).Visible = True
                        Text(6).Visible = True
                        Shape1.Visible = True
                    End If
                End If
                oRs.Close
            End If
        Case 5
            If Text(Index) <> "" Then
                    Set oRs = objO.FetchDataACC("select * From inmst WITH (NOLOCK) where kodeItem = '" & Text(Index) & "'")
                    If oRs.BOF Then
                        Cancel = True
                        gMainKeyCari = Text(Index)
                        gSearchOption = "PART"
                        frmCariAcess.show (vbModal)
                    Else
                        Text(6) = oRs("NamaBarang")
                    End If
                    oRs.Close
                
            Else
                Text(Index + 1) = ""
            End If
        Case 7
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("SELECT * FROM T_DEALER WITH (NOLOCK) WHERE FS_KD_DEALER = '" & Text(Index) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gBantuan = "INTERNAL"
                    gSearchOption = "DEALER"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = oRs("FS_NM_DEALER")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
    End Select
End Sub


Private Sub cmdExport_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsGrid.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Reminder Service Kosong !"
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
16       ' GridMOU.Columns(7).Visible = False
17         GridProgram.ExportToFile onVal, True
18       ' GridMOU.Columns(7).Visible = True
19                Er.HeadCap = "Export to excel !"
20                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
21                Beep
22                Er.show (vbModal)
23         Exit Sub
Err_Export:
24          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
25          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub
