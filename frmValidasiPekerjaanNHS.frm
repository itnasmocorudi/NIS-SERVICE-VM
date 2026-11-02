VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmValidasiPekerjaanNHS 
   Caption         =   ".: Validasi Pekerjaan NHS  :."
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   16320
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   8775
      Left            =   480
      TabIndex        =   28
      Top             =   -120
      Visible         =   0   'False
      Width           =   12495
      Begin VB.Frame Frame4 
         Caption         =   "Detail Order"
         Height          =   975
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   12255
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   3
            Left            =   960
            TabIndex        =   30
            Top             =   240
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   4
            Left            =   960
            TabIndex        =   31
            Top             =   525
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   0
            Left            =   3840
            TabIndex        =   34
            Top             =   240
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   1
            Left            =   3840
            TabIndex        =   35
            Top             =   525
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   2
            Left            =   8160
            TabIndex        =   38
            Top             =   240
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   5
            Left            =   8160
            TabIndex        =   39
            Top             =   525
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Polisi"
            Height          =   195
            Left            =   7455
            TabIndex        =   41
            Top             =   570
            Width           =   585
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
            Height          =   195
            Left            =   7440
            TabIndex        =   40
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teknisi"
            Height          =   195
            Left            =   3240
            TabIndex        =   37
            Top             =   570
            Width           =   480
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model "
            Height          =   195
            Left            =   3345
            TabIndex        =   36
            Top             =   240
            Width           =   465
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Order"
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   570
            Width           =   660
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID Req"
            Height          =   195
            Left            =   435
            TabIndex        =   32
            Top             =   240
            Width           =   495
         End
      End
      Begin TrueOleDBGrid70.TDBGrid gridJobSebelum 
         Height          =   2175
         Left            =   240
         TabIndex        =   42
         Top             =   1320
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Program"
         Columns(0).DataField=   "Program"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama Program"
         Columns(1).DataField=   "NamaProgram"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Job Order"
         Columns(2).DataField=   "Joborder"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Rate"
         Columns(3).DataField=   "Rate"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
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
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
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
         Caption         =   "Job Order Sebelum"
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&H47ACF8&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
         _StyleDefs(73)  =   "Named:id=83:EditRow"
         _StyleDefs(74)  =   ":id=83,.parent=33,.locked=-1"
      End
      Begin TrueOleDBGrid70.TDBGrid gridJobReq 
         Height          =   2175
         Left            =   6600
         TabIndex        =   43
         Top             =   1320
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pilih"
         Columns(0).DataField=   "pilih"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Program"
         Columns(1).DataField=   "Program"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama Program"
         Columns(2).DataField=   "namaprogram"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Job Order"
         Columns(3).DataField=   "Joborder"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Rate"
         Columns(4).DataField=   "Rate"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=3545"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=3466"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8708"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=2302"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2223"
         Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8708"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
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
         Caption         =   "Job Order Permintaan"
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&H47ACF8&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
         _StyleDefs(77)  =   "Named:id=83:EditRow"
         _StyleDefs(78)  =   ":id=83,.parent=33,.locked=-1"
      End
      Begin TrueOleDBGrid70.TDBGrid gridPartSebelum 
         Height          =   2175
         Left            =   240
         TabIndex        =   44
         Top             =   3600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode Item"
         Columns(0).DataField=   "Kodeitem"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama Part"
         Columns(1).DataField=   "namapart"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Qty"
         Columns(2).DataField=   "Qty"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
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
         Caption         =   "Part Order Sebelum"
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
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
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&H47ACF8&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
         _StyleDefs(69)  =   "Named:id=83:EditRow"
         _StyleDefs(70)  =   ":id=83,.parent=33,.locked=-1"
      End
      Begin TrueOleDBGrid70.TDBGrid gridPartReq 
         Height          =   2175
         Left            =   6600
         TabIndex        =   45
         Top             =   3600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pilih"
         Columns(0).DataField=   "pilih"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Kode Item"
         Columns(1).DataField=   "kodeitem"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama Part"
         Columns(2).DataField=   "namapart"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Qty"
         Columns(3).DataField=   "qty"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3519"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3440"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8708"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8708"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
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
         Caption         =   "Part Order Permintaan"
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=74,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
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
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&H47ACF8&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
         _StyleDefs(73)  =   "Named:id=83:EditRow"
         _StyleDefs(74)  =   ":id=83,.parent=33,.locked=-1"
      End
      Begin TrueOleDBGrid70.TDBGrid gridSubletSebelum 
         Height          =   2175
         Left            =   240
         TabIndex        =   46
         Top             =   5880
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode Promo"
         Columns(0).DataField=   "KdPromo"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama Promo"
         Columns(1).DataField=   "NamaPromo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Kode Sublet"
         Columns(2).DataField=   "kdsublet"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nama Sublet"
         Columns(3).DataField=   "nama"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Qty"
         Columns(4).DataField=   "Qty"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
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
         Caption         =   "Sublet Order Sebelum"
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&H47ACF8&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
         _StyleDefs(77)  =   "Named:id=83:EditRow"
         _StyleDefs(78)  =   ":id=83,.parent=33,.locked=-1"
      End
      Begin TrueOleDBGrid70.TDBGrid gridSubletReq 
         Height          =   2175
         Left            =   6600
         TabIndex        =   47
         Top             =   5880
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pilih"
         Columns(0).DataField=   "pilih"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Kode Promo"
         Columns(1).DataField=   "kdpromo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama Promo"
         Columns(2).DataField=   "NamaPromo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Kode Sublet"
         Columns(3).DataField=   "kdsublet"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nama Sublet"
         Columns(4).DataField=   "Nama"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Qty"
         Columns(5).DataField=   "Qty"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2117"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2037"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8708"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8708"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(25)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
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
         Caption         =   "Sublet Order Permintaan"
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=78,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0"
         _StyleDefs(62)  =   "Named:id=34:Heading"
         _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(64)  =   ":id=34,.wraptext=-1"
         _StyleDefs(65)  =   "Named:id=35:Footing"
         _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(67)  =   "Named:id=36:Selected"
         _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=37:Caption"
         _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(71)  =   "Named:id=38:HighlightRow"
         _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=39:EvenRow"
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&H47ACF8&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
         _StyleDefs(81)  =   "Named:id=83:EditRow"
         _StyleDefs(82)  =   ":id=83,.parent=33,.locked=-1"
      End
      Begin ucXPButton.XPButton cmdSimpanDetail 
         Height          =   495
         Left            =   9960
         TabIndex        =   48
         Top             =   8160
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
         MICON           =   "frmValidasiPekerjaanNHS.frx":0000
         PICN            =   "frmValidasiPekerjaanNHS.frx":001C
         PICH            =   "frmValidasiPekerjaanNHS.frx":681F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdBatalDetail 
         Height          =   495
         Left            =   11160
         TabIndex        =   49
         Top             =   8160
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
         MICON           =   "frmValidasiPekerjaanNHS.frx":D3C5
         PICN            =   "frmValidasiPekerjaanNHS.frx":D3E1
         PICH            =   "frmValidasiPekerjaanNHS.frx":139DA
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
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   13560
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   12975
      Begin TrueOleDBGrid70.TDBDropDown TDBDropDown2 
         Bindings        =   "frmValidasiPekerjaanNHS.frx":1A289
         Height          =   1335
         Left            =   240
         TabIndex        =   22
         Top             =   5280
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2355
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode Petugas"
         Columns(0).DataField=   "KODE"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nik Petugas"
         Columns(1).DataField=   "nik"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No Polisi"
         Columns(2).DataField=   "kodeunit"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         AllowRowSizing  =   0   'False
         Appearance      =   1
         BorderStyle     =   1
         ColumnHeaders   =   -1  'True
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   ""
         DataField       =   "SHIFT"
         IntegralHeight  =   0   'False
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   0   'False
         DataMember      =   ""
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   14215660
         ValueTranslate  =   0   'False
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=180,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
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
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   2880
         TabIndex        =   6
         Top             =   3360
         Visible         =   0   'False
         Width           =   7215
         Begin VB.TextBox txtAlasan 
            Height          =   975
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   480
            Width           =   6975
         End
         Begin ucXPButton.XPButton cmdSimpanPerubahan 
            Height          =   495
            Left            =   4680
            TabIndex        =   9
            Top             =   1560
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
            MICON           =   "frmValidasiPekerjaanNHS.frx":1A29E
            PICN            =   "frmValidasiPekerjaanNHS.frx":1A2BA
            PICH            =   "frmValidasiPekerjaanNHS.frx":20ABD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdBatalKonfirmasi 
            Height          =   495
            Left            =   5880
            TabIndex        =   27
            Top             =   1560
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
            MICON           =   "frmValidasiPekerjaanNHS.frx":27663
            PICN            =   "frmValidasiPekerjaanNHS.frx":2767F
            PICH            =   "frmValidasiPekerjaanNHS.frx":2DC78
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblParam 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblParam 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblParam 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Alasan"
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
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   585
         End
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   10920
         TabIndex        =   1
         Top             =   7440
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
         MICON           =   "frmValidasiPekerjaanNHS.frx":34527
         PICN            =   "frmValidasiPekerjaanNHS.frx":34543
         PICH            =   "frmValidasiPekerjaanNHS.frx":3AE76
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdLihatDetail 
         Height          =   495
         Left            =   8640
         TabIndex        =   2
         Top             =   7440
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   " "
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
         MICON           =   "frmValidasiPekerjaanNHS.frx":41B2D
         PICN            =   "frmValidasiPekerjaanNHS.frx":41B49
         PICH            =   "frmValidasiPekerjaanNHS.frx":461F3
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
         Left            =   2640
         TabIndex        =   3
         Top             =   7440
         Visible         =   0   'False
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
         MICON           =   "frmValidasiPekerjaanNHS.frx":4B103
         PICN            =   "frmValidasiPekerjaanNHS.frx":4B11F
         PICH            =   "frmValidasiPekerjaanNHS.frx":51681
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc AdoAs 
         Height          =   330
         Index           =   0
         Left            =   10680
         Top             =   480
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc AdoAs 
         Height          =   330
         Index           =   1
         Left            =   8640
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   5040
         TabIndex        =   10
         Top             =   7440
         Visible         =   0   'False
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
         MICON           =   "frmValidasiPekerjaanNHS.frx":57EA4
         PICN            =   "frmValidasiPekerjaanNHS.frx":57EC0
         PICH            =   "frmValidasiPekerjaanNHS.frx":5E6C3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBDate6Ctl.TDBDate Tgl1 
         Height          =   330
         Left            =   1800
         TabIndex        =   14
         Top             =   120
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmValidasiPekerjaanNHS.frx":65269
         Caption         =   "frmValidasiPekerjaanNHS.frx":65395
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmValidasiPekerjaanNHS.frx":65401
         Keys            =   "frmValidasiPekerjaanNHS.frx":6541F
         Spin            =   "frmValidasiPekerjaanNHS.frx":6547D
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
      Begin TDBDate6Ctl.TDBDate Tgl2 
         Height          =   330
         Left            =   3720
         TabIndex        =   15
         Top             =   120
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmValidasiPekerjaanNHS.frx":654A5
         Caption         =   "frmValidasiPekerjaanNHS.frx":655D1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmValidasiPekerjaanNHS.frx":6563D
         Keys            =   "frmValidasiPekerjaanNHS.frx":6565B
         Spin            =   "frmValidasiPekerjaanNHS.frx":656B9
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
      Begin TDBDate6Ctl.TDBDate Tgl3 
         Height          =   330
         Left            =   1800
         TabIndex        =   16
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmValidasiPekerjaanNHS.frx":656E1
         Caption         =   "frmValidasiPekerjaanNHS.frx":6580D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmValidasiPekerjaanNHS.frx":65879
         Keys            =   "frmValidasiPekerjaanNHS.frx":65897
         Spin            =   "frmValidasiPekerjaanNHS.frx":658F5
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
      Begin TDBDate6Ctl.TDBDate Tgl4 
         Height          =   330
         Left            =   3720
         TabIndex        =   17
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmValidasiPekerjaanNHS.frx":6591D
         Caption         =   "frmValidasiPekerjaanNHS.frx":65A49
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmValidasiPekerjaanNHS.frx":65AB5
         Keys            =   "frmValidasiPekerjaanNHS.frx":65AD3
         Spin            =   "frmValidasiPekerjaanNHS.frx":65B31
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
      Begin ucXPButton.XPButton cmdRefresh 
         Height          =   495
         Left            =   5520
         TabIndex        =   20
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
         MICON           =   "frmValidasiPekerjaanNHS.frx":65B59
         PICN            =   "frmValidasiPekerjaanNHS.frx":65B75
         PICH            =   "frmValidasiPekerjaanNHS.frx":69153
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKirimSuratTugas 
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   7920
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
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
         MICON           =   "frmValidasiPekerjaanNHS.frx":6C86C
         PICN            =   "frmValidasiPekerjaanNHS.frx":6C888
         PICH            =   "frmValidasiPekerjaanNHS.frx":70F32
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid gridNHS 
         Height          =   6375
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   11245
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pilih"
         Columns(0).DataField=   "pilih"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "ID Req"
         Columns(1).DataField=   "idreq"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "User"
         Columns(2).DataField=   "User"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tanggal Req"
         Columns(3).DataField=   "ReqBook"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Petugas NHS"
         Columns(4).DataField=   "User"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Nama Customer"
         Columns(5).DataField=   "customer"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "No Order"
         Columns(6).DataField=   "noWO"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "No Booking"
         Columns(7).DataField=   "nofak"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Model Kendaraan"
         Columns(8).DataField=   "model"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Kode Model"
         Columns(9).DataField=   "mdl"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "No Polisi"
         Columns(10).DataField=   "NoPolisi"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1693"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8708"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2302"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2223"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8708"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=3519"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3440"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8708"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2117"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2037"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8708"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=8708"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=8708"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(46)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(9)._ColStyle=8708"
         Splits(0)._ColumnProps(50)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(51)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(52)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(54)=   "Column(10)._ColStyle=8708"
         Splits(0)._ColumnProps(55)=   "Column(10).Order=11"
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
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   0
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14,.alignment=2"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14,.alignment=2"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.alignment=2"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14,.alignment=2"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14,.alignment=2"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14,.alignment=2"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.locked=-1"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14,.alignment=2"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.locked=-1"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14,.alignment=2"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
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
         _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&H47ACF8&"
         _StyleDefs(95)  =   "Named:id=40:OddRow"
         _StyleDefs(96)  =   ":id=40,.parent=33"
         _StyleDefs(97)  =   "Named:id=41:RecordSelector"
         _StyleDefs(98)  =   ":id=41,.parent=34"
         _StyleDefs(99)  =   "Named:id=42:FilterBar"
         _StyleDefs(100) =   ":id=42,.parent=33"
         _StyleDefs(101) =   "Named:id=83:EditRow"
         _StyleDefs(102) =   ":id=83,.parent=33,.locked=-1"
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   3480
         TabIndex        =   19
         Top             =   195
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   3480
         TabIndex        =   18
         Top             =   555
         Width           =   60
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal WO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   615
         TabIndex        =   5
         Top             =   555
         Width           =   1005
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Req"
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
         TabIndex        =   4
         Top             =   255
         Width           =   1305
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmValidasiPekerjaanNHS.frx":75E42
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmValidasiPekerjaanNHS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmValidasiPekerjaanNHS
' DateTime  : 05/01/2021 15:13
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------


Dim rsGrid      As New ADODB.Recordset
Dim rsJobReq    As New ADODB.Recordset
Dim rsPartReq   As New ADODB.Recordset
Dim rsSubletReq As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset
Dim cols        As TrueOleDBGrid70.Columns
Dim Urut        As String
Dim isTambah    As Boolean
Dim isHapus     As Boolean
Dim isBegin     As Boolean
Dim Report      As Object

Private Sub createTableGrid()
   On Error GoTo createTableGrid_Error
         If rsGrid.State Then rsGrid.Close
          Set rsGrid = Nothing
          With rsGrid
              .Fields.Append "Pilih", adBoolean
              .Fields.Append "User", adVarChar, 20
              .Fields.Append "ReqBook", adDate
              .Fields.Append "Customer", adVarChar, 200
              .Fields.Append "idreq", adVarChar, 200
              .Fields.Append "noWO", adVarChar, 200
              .Fields.Append "Model", adVarChar, 100
              .Fields.Append "NoPolisi", adVarChar, 15
              .Fields.Append "Nofak", adVarChar, 30
              .Fields.Append "Mdl", adVarChar, 30
             .Open
          End With
   On Error GoTo 0
   Exit Sub
createTableGrid_Error:
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure createTableGrid of Form frmValidasiPekerjaanNHS"
End Sub

Private Sub createTableJobReq()
   On Error GoTo createTableGrid_Error
         If rsJobReq.State Then rsJobReq.Close
          Set rsJobReq = Nothing
          With rsJobReq
              .Fields.Append "Pilih", adBoolean
              .Fields.Append "Program", adVarChar, 50
              .Fields.Append "NamaProgram", adVarChar, 100
              .Fields.Append "Joborder", adVarChar, 100
              .Fields.Append "rate", adDouble
             .Open
          End With
   On Error GoTo 0
   Exit Sub
createTableGrid_Error:
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure createTableGrid of Form frmValidasiPekerjaanNHS"
End Sub

Private Sub createTablePartReq()
   On Error GoTo createTableGrid_Error
         If rsPartReq.State Then rsPartReq.Close
          Set rsPartReq = Nothing
          With rsPartReq
              .Fields.Append "Pilih", adBoolean
              .Fields.Append "kodeitem", adVarChar, 50
              .Fields.Append "namapart", adVarChar, 100
              .Fields.Append "qty", adDouble
             .Open
          End With
   On Error GoTo 0
   Exit Sub
createTableGrid_Error:
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure createTableGrid of Form frmValidasiPekerjaanNHS"
End Sub

Private Sub createTableSubletReq()
   On Error GoTo createTableGrid_Error
         If rsSubletReq.State Then rsSubletReq.Close
          Set rsSubletReq = Nothing
          With rsSubletReq
              .Fields.Append "Pilih", adBoolean
              .Fields.Append "kdpromo", adVarChar, 50
              .Fields.Append "kdsublet", adVarChar, 50
              .Fields.Append "nama", adVarChar, 100
              .Fields.Append "namapromo", adVarChar, 100
              .Fields.Append "qty", adDouble
             .Open
          End With
   On Error GoTo 0
   Exit Sub
createTableGrid_Error:
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure createTableGrid of Form frmValidasiPekerjaanNHS"
End Sub

Private Sub ber_sih()
    Tgl1.value = isnow
    Tgl2.value = isnow

    Tgl3.value = ""
    Tgl4.value = ""
    
    With AdoAs(0)
        .ConnectionString = adoCon
        .CommandType = adCmdText
        .RecordSource = "SELECT FS_KD_PETUGAS AS KODE,NIK,b.NoPol as kodeunit FROM PlottingPetugasNHS a WITH (NOLOCK) inner join MstUnitNHS b with (nolock) on a.KodeUnit = b.KodeUnit and a.KodeCabang = b.KodeCabang WHERE a.KodeCabang = '" & gDb & "'"
        .Refresh
    End With
    TDBDropDown2.DataSource = AdoAs(0)
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)

End Sub
 

Private Sub cmdBatalDetail_Click()
    Frame3.Visible = False
    Call cmdRefresh_Click
End Sub

Private Sub cmdBatalKonfirmasi_Click()
    If MsgBox("Apakah yakin membatalkan proses?", vbYesNo) = vbYes Then
        Frame2.Visible = False
        cmdSimpan.Enabled = True
        Call cmdRefresh_Click
    End If
End Sub
 
Private Sub cmdEdit_Click()

    Me.MousePointer = vbHourglass
    isTambah = False
    isHapus = False
    Call buka_Tutup(True)
    Me.MousePointer = vbDefault
End Sub

 

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdKirimSuratTugas_Click()
          
1        On Error GoTo cmdKirimSuratTugas_Click_Error
          Dim isBegin     As Boolean
          Dim boleh       As Boolean
          Dim pesan       As String
          Dim allModel    As Byte
          Dim allDealer   As Byte
          Dim KodeUnit    As String
          Dim nik         As String
          Dim FS_KD_PETUGAS   As String
         ' Dim pesan       As String
          Dim rsPetugas   As ADODB.Recordset
2         boleh = True
3         isBegin = False
          
4         If rsGrid.State = adStateOpen Then
5             If rsGrid.RecordCount > 0 Then
6                 rsGrid.MoveFirst
7                 rsGrid.filter = "pilih = true"
8                 If Not rsGrid.EOF Then
                       
9                     rsGrid.filter = ""
10                Else
11                    rsGrid.filter = ""
12                    GoTo exitsimpan
13                End If
14            Else
15                GoTo exitsimpan
16            End If
17        Else
exitsimpan:
18            MsgBox "Pilih data yang akan diproses", vbInformation
19            Exit Sub
20        End If
          
21        If boleh Then
22            Call objHT.OpenTransactionBengkel
23            isBegin = True
24            Do While Not rsGrid.EOF
25                sSQL = "SELECT FS_KD_PETUGAS AS KODE,NIK,b.kodeunit as kodeunit FROM PlottingPetugasNHS a WITH (NOLOCK) inner join MstUnitNHS b with (nolock) on a.KodeUnit = b.KodeUnit and a.KodeCabang = b.KodeCabang WHERE a.KodeCabang = '" & gDb & "' and FS_KD_PETUGAS = '" & NulSID(rsGrid!Teknisi) & "'"
26                Set rsPetugas = objHT.FetchDataBengkel(sSQL)
27                If Not rsPetugas.EOF Then
                  
28                    sSQL = "INSERT INTO LogPlottingbookingNHS (NOFAK,MDL,NO_POLISI,NO_ENGINE,KodeUnit,NIK,FS_KD_PETUGAS,UserEntry,TglEntry,cetak,UserCetak,TglCetak,kirim,TglKirim,UserKirim,[Status],TglStatus)"
29                    sSQL = sSQL & " SELECT NOFAK,MDL,NO_POLISI,NO_ENGINE,KodeUnit,NIK,FS_KD_PETUGAS,UserEntry,TglEntry,cetak,UserCetak,TglCetak,kirim,TglKirim,UserKirim,[Status],TglStatus FROM PlottingbookingNHS WITH (NOLOCK) WHERE NOFAK = '" & NulSID(rsGrid!NoFak) & "' and NO_POLISI = '" & NulSID(rsGrid!NoPolisi) & "'"
30                    Call objHT.InsertData_Bengkel(sSQL)
                      
31                    sSQL = "UPDATE PlottingbookingNHS SET kirim = '1',tglKirim = getdate(),UserKirim = '" & gSA & "' WHERE NOFAK = '" & NulSID(rsGrid!NoFak) & "' and NO_POLISI = '" & NulSID(rsGrid!NoPolisi) & "'"
32                    Call objHT.InsertData_Bengkel(sSQL)
33                End If
34                rsGrid.MoveNext
35            Loop
36            pesan = "Data berhasil diubah."
37            rsGrid.filter = ""

38            Call objHT.CloseTransactionBengkel
39            MsgBox pesan, vbInformation
40            cmdSimpan.Enabled = True
41            Call cmdRefresh_Click
42        Else
43            MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
44        End If
45       On Error GoTo 0
46       Exit Sub

cmdKirimSuratTugas_Click_Error:

47        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdKirimSuratTugas_Click of Form frmValidasiPekerjaanNHS"


End Sub

Private Sub cmdLihatDetail_Click()
              Dim rsData As ADODB.Recordset
1         If rsGrid.State = adStateOpen Then
2             If rsGrid.RecordCount > 0 Then
3                 rsGrid.MoveFirst
4                 rsGrid.filter = "pilih = true"
5                 If Not rsGrid.EOF Then
6                     If rsGrid.RecordCount > 1 Then
7                         MsgBox "Pilih salah satu data yang akan dicetak.", vbInformation
8                         Exit Sub
9                     Else
                          
10                    End If
11                Else
12                    rsGrid.filter = ""
13                    GoTo exitsimpan
14                End If
15            Else
16                GoTo exitsimpan
17            End If
18        Else
exitsimpan:
19            MsgBox "Pilih data yang akan diproses", vbInformation
20            Exit Sub
21        End If
          
22        Text(3) = NulSID(rsGrid!idreq)
23        Text(4) = NulSID(rsGrid!noWO)
24        Text(0) = NulSID(rsGrid!model)
25        Text(1) = NulSID(rsGrid!User)
26        Text(2) = NulSID(rsGrid!customer)
27        Text(5) = NulSID(rsGrid!NoPolisi)
          
          Dim joborderSebelum As ADODB.Recordset
28        sSQL = "SELECT a.ID_PROGRAM AS PROGRAM,case when ISNULL(b.PROGRAM,'') <> '' then ISNULL(b.PROGRAM,'') else ISNULL(c.namaFP,'') end namaprogram , a.REP AS JOBORDER, a.RTE AS RATE  "
29        sSQL = sSQL & " FROM jobdetail a WITH (NOLOCK)"
30        sSQL = sSQL & " left join [ReplBengkel99]..HO_PROGRAM b with (nolock) on a.id_program = b.id_Program and b.FS_LOG_DEL = ''"
31        sSQL = sSQL & " left join [ReplBengkel99]..HO_FIXPRICE c with (nolock) on a.id_program = c.id_FP and c.FS_LOG_DEL = ''"
32        sSQL = sSQL & " WHERE a.NO_WO = '" & NulSID(rsGrid!noWO) & "'"
33        Set joborderSebelum = objO.FetchData(sSQL)
34        Set gridJobSebelum.DataSource = joborderSebelum
          
35        sSQL = "SELECT a.PROGRAM,case when ISNULL(b.PROGRAM,'') <> '' then ISNULL(b.PROGRAM,'') else ISNULL(c.namaFP,'') end namaprogram , a.REP AS JOBORDER, a.RTE AS RATE  "
36        sSQL = sSQL & " FROM ReqUbahJobWONHS a WITH (NOLOCK)"
37        sSQL = sSQL & " left join [ReplBengkel99]..HO_PROGRAM b with (nolock) on a.program = b.id_Program and b.FS_LOG_DEL = ''"
38        sSQL = sSQL & " left join [ReplBengkel99]..HO_FIXPRICE c with (nolock) on a.program = c.id_FP and c.FS_LOG_DEL = ''"
39        sSQL = sSQL & " WHERE a.idreq = '" & NulSID(rsGrid!idreq) & "'"
40        Set rsTemp = objO.FetchData(sSQL)
41        Call createTableJobReq
42        If Not rsTemp.EOF Then
43            rsTemp.MoveFirst
44            Do While Not rsTemp.EOF
45                With rsJobReq
46                    .AddNew
47                    !Pilih = True
48                    !program = NulSID(rsTemp!program)
49                    !namaProgram = NulSID(rsTemp!namaProgram)
50                    !jobOrder = NulSID(rsTemp!jobOrder)
51                    !Rate = NulSID(rsTemp!Rate)
52                    .Update
53                End With
54                rsTemp.MoveNext
55            Loop
56        End If
57        Set gridJobReq.DataSource = rsJobReq
          
          Dim partSebelum As ADODB.Recordset
58        sSQL = "SELECT KODE_ITEM kodeitem,NAMA_ITEM namapart,QTY  FROM JOBPART WITH (NOLOCK) WHERE NO_WO = '" & NulSID(rsGrid!noWO) & "'"
59        Set partSebelum = objO.FetchData(sSQL)
60        Set gridPartSebelum.DataSource = partSebelum
          
61        sSQL = "SELECT kode_item, nama_item, qty  FROM ReqUbahPARTWONHS WITH (NOLOCK) WHERE idreq = '" & NulSID(rsGrid!idreq) & "'"
62        Set rsTemp = objO.FetchData(sSQL)
63            Call createTablePartReq
64        If Not rsTemp.EOF Then
65            rsTemp.MoveFirst
66            Do While Not rsTemp.EOF
67                With rsPartReq
68                    .AddNew
69                    !Pilih = True
70                    !KodeItem = NulSID(rsTemp!KODE_ITEM)
71                    !namaPart = NulSID(rsTemp!NAMA_ITEM)
72                    !qty = NulSID(rsTemp!qty)
73                    .Update
74                End With
75                rsTemp.MoveNext
76            Loop
77        End If
78        Set gridPartReq.DataSource = rsPartReq
          
          
          Dim subletSebelum As ADODB.Recordset
79        sSQL = "SELECT a.KODEITEM kdsublet,a.NAMABARANG nama,a.qty,a.KDPROMO,isnull(b.NAMAPROMO,'') as NAMAPROMO"
80        sSQL = sSQL & " FROM JOBSUBLET a WITH (NOLOCK)"
81        sSQL = sSQL & " left join [ReplBengkel99]..MSTSUBLETPROMO b with (nolock) on a.KDPROMO = b.KDPROMO and b.FS_LOG_DEL = ''"
82        sSQL = sSQL & " WHERE a.NO_WO = '" & NulSID(rsGrid!noWO) & "'"
83        Set subletSebelum = objO.FetchData(sSQL)
84        Set gridSubletSebelum.DataSource = subletSebelum
          
85        sSQL = "SELECT a.KODEITEM kdsublet,a.NAMABARANG Nama,a.qty,a.KDPROMO,isnull(b.NAMAPROMO,'') as NAMAPROMO "
86        sSQL = sSQL & " FROM ReqUbahsubletWONHS a WITH (NOLOCK)"
87        sSQL = sSQL & " left join [ReplBengkel99]..MSTSUBLETPROMO b with (nolock) on a.KDPROMO = b.KDPROMO and b.FS_LOG_DEL = ''"
88        sSQL = sSQL & " WHERE a.idreq = '" & NulSID(rsGrid!idreq) & "'"
89        Set rsTemp = objO.FetchData(sSQL)
90        Call createTableSubletReq
91        If Not rsTemp.EOF Then
92            rsTemp.MoveFirst
93            Do While Not rsTemp.EOF
94                With rsSubletReq
95                    .AddNew
96                    !Pilih = True
97                    !KDSUBLET = NulSID(rsTemp!KDSUBLET)
98                    !KDPROMO = NulSID(rsTemp!KDPROMO)
99                    !Nama = NulSID(rsTemp!Nama)
100                   !NAMAPROMO = NulSID(rsTemp!NAMAPROMO)
101                   !qty = NulSID(rsTemp!qty)
102                   .Update
103               End With
104               rsTemp.MoveNext
105           Loop
106       End If
107       Set gridSubletReq.DataSource = rsSubletReq
108       Frame3.Visible = True
109       Screen.MousePointer = vbDefault
110       rsGrid.filter = ""
111       Exit Sub
End Sub

Private Sub cmdRefresh_Click()
    Dim kriteria As String
    kriteria = ""
    If Tgl1.value <> "" And Tgl2.value <> "" Then
        kriteria = kriteria & " AND (a.tglentry between '" & Format(Tgl1.value, "yyyy-mm-dd") & " 00:00' and '" & Format(Tgl2.value, "yyyy-mm-dd") & " 23:59') "
    End If
    If Tgl3.value <> "" And Tgl4.value <> "" Then
        kriteria = kriteria & " AND (b.[date] between '" & Format(Tgl3.value, "yyyy-mm-dd") & " 00:00' and '" & Format(Tgl4.value, "yyyy-mm-dd") & " 23:59') "
    End If
    If kriteria <> "" Then
        Call refreshGrid(kriteria)
    Else
        MsgBox "Kriteria tanggal harus diisi", vbInformation
        Exit Sub
    End If
End Sub

Private Sub cmdSettingPrinter_Click()
Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdSimpan_Click()

1        On Error GoTo cmdSimpan_Click_Error
          Dim isBegin     As Boolean
          Dim boleh       As Boolean
          Dim pesan       As String
          Dim allModel    As Byte
          Dim allDealer   As Byte
          Dim KodeUnit    As String
          Dim nik         As String
          Dim FS_KD_PETUGAS   As String
         ' Dim pesan       As String
          Dim rsPetugas   As ADODB.Recordset
          Dim rsCekData   As ADODB.Recordset
          
2         boleh = True
3         isBegin = False
          
4         If rsGrid.State = adStateOpen Then
5             If rsGrid.RecordCount > 0 Then
6                 rsGrid.MoveFirst
7                 rsGrid.filter = "pilih = true"
8                 If Not rsGrid.EOF Then
9                     Do While Not rsGrid.EOF
                      
10                        If NulSID(rsGrid!Teknisi) = "" Then
11                            boleh = False
12                            pesan = "Petugas untuk no polisi " & NulSID(rsGrid!NoPolisi) & " harus diisi" & vbCrLf
13                        ElseIf objO.FetchData_String("SELECT FS_KD_PETUGAS AS KODE,NIK,b.NoPol as kodeunit FROM PlottingPetugasNHS a WITH (NOLOCK) inner join MstUnitNHS b with (nolock) on a.KodeUnit = b.KodeUnit and a.KodeCabang = b.KodeCabang WHERE a.KodeCabang = '" & gDb & "' and FS_KD_PETUGAS = '" & NulSID(rsGrid!Teknisi) & "'") = "" Then
14                            boleh = False
15                            pesan = "Petugas " & NulSID(rsGrid!Teknisi) & " tidak terdaftar" & vbCrLf
16                        End If
                          
                          ''penambahan pengecekan untuk part harus terdaftar dalam inmst
                          ''----------------------------------------------------------------------
17                        sSQL = "SELECT KODEITEM FROM BOOK_PART WITH (NOLOCK) WHERE NOFAK = '" & NulSID(rsGrid!NoFak) & "' "
18                        Set rsCekData = objO.FetchData(sSQL)
19                        If Not rsCekData.EOF Then
20                            rsCekData.MoveFirst
21                            Do While Not rsCekData.EOF
22                                sSQL = "SELECT KODEITEM FROM INMST WITH (NOLOCK) WHERE KODEITEM = '" & NulSID(rsCekData!KodeItem) & "' AND kodegdg = '03'"
23                                If NulSID(rsCekData!KodeItem) = "" Then
24                                    boleh = False
25                                    pesan = "Kode item part di no booking " & NulSID(rsGrid!NoFak) & " tidak terdaftar" & vbCrLf
                                      
26                                ElseIf objO.FetchData_String_Acc(sSQL) = "" Then
27                                    boleh = False
28                                    pesan = "Kode item part di no booking " & NulSID(rsGrid!NoFak) & " tidak terdaftar" & vbCrLf
29                                End If
30                                rsCekData.MoveNext
31                            Loop
32                        End If
                          
                          ''penambahan pengecekan untuk sublet harus terdaftar dalam sentralisasi sublet
                          ''----------------------------------------------------------------------
33                        sSQL = "SELECT KODEITEM FROM BOOK_SUBLET WITH (NOLOCK) WHERE NOFAK = '" & NulSID(rsGrid!NoFak) & "' "
34                        Set rsCekData = objO.FetchData(sSQL)
35                        If Not rsCekData.EOF Then
36                            rsCekData.MoveFirst
37                            Do While Not rsCekData.EOF
38                                sSQL = "SELECT kdsublet FROM [" & dbReplikasiHO & "]..mstsublet WITH (NOLOCK) WHERE fs_log_del = '' and kdsublet = '" & NulSID(rsCekData!KodeItem) & "' "
39                                If NulSID(rsCekData!KodeItem) = "" Then
40                                    boleh = False
41                                    pesan = "Kode Sublet di no booking " & NulSID(rsGrid!NoFak) & " tidak terdaftar" & vbCrLf
                                      
42                                ElseIf objO.FetchData_String(sSQL) = "" Then
43                                    boleh = False
44                                    pesan = "Kode Sublet di no booking " & NulSID(rsGrid!NoFak) & " tidak terdaftar" & vbCrLf
45                                End If
46                                rsCekData.MoveNext
47                            Loop
48                        End If
49                        rsGrid.MoveNext
50                    Loop
51                    rsGrid.filter = ""
52                Else
53                    rsGrid.filter = ""
54                    GoTo exitsimpan
55                End If
56            Else
57                GoTo exitsimpan
58            End If
59        Else
exitsimpan:
60            MsgBox "Pilih data yang akan diproses", vbInformation
61            Exit Sub
62        End If
       

63        If boleh Then
64            Call objHT.OpenTransactionBengkel
65            isBegin = True
66            erTranAll = False
67            rsGrid.filter = "pilih = true and isedit = false"
68            If Not rsGrid.EOF Then
69                Do While Not rsGrid.EOF
                      ''insert data
                      
70                    sSQL = "SELECT FS_KD_PETUGAS AS KODE,NIK,b.kodeunit as kodeunit FROM PlottingPetugasNHS a WITH (NOLOCK) inner join MstUnitNHS b with (nolock) on a.KodeUnit = b.KodeUnit and a.KodeCabang = b.KodeCabang WHERE a.KodeCabang = '" & gDb & "' and FS_KD_PETUGAS = '" & NulSID(rsGrid!Teknisi) & "'"
71                    Set rsPetugas = objHT.FetchDataBengkel(sSQL)
72                    If Not rsPetugas.EOF Then
73                        sSQL = "insert into PlottingBookingNHS (NOFAK,MDL,NO_POLISI,NO_ENGINE,KodeUnit,"
74                        sSQL = sSQL & " NIK,FS_KD_PETUGAS,UserEntry,TglEntry,cetak,UserCetak,TglCetak,kirim,"
75                        sSQL = sSQL & " TglKirim,UserKirim,[Status],TglStatus)"
76                        sSQL = sSQL & " values ('" & NulSID(rsGrid!NoFak) & "','" & NulSID(rsGrid!MDL) & "','" & NulSID(rsGrid!NoPolisi) & "','" & NulSID(rsGrid!NoEngine) & "','" & NulSID(rsPetugas!KodeUnit) & "','" & NulSID(rsPetugas!nik) & "','" & NulSID(rsPetugas!kode) & "','" & gSA & "',"
77                        sSQL = sSQL & " getdate(),'0','','1900-01-01','0','1900-01-01','','','1900-01-01')"
78                        Call objHT.InsertData_Bengkel(sSQL)
79                    End If
80                    rsGrid.MoveNext
81                Loop
82            End If
83            rsGrid.filter = ""
84            rsGrid.filter = "pilih = true and isedit = true"
85            If Not rsGrid.EOF Then
                  '' commit atas data tambah
86                Call objHT.CloseTransactionBengkel
87                Frame2.Visible = True
88                txtAlasan.Text = ""
89                cmdSimpan.Enabled = False
90                Exit Sub
91            End If

92            If erTranAll = False Then
93                Call objHT.CloseTransactionBengkel
94                MsgBox "Berhasil menyimpan data.", vbInformation
95                Call ber_sih
96                Call buka_Tutup(False)
97                Call cmdRefresh_Click ''refreshGrid(
98            Else
errorRollback:
99                Call objHT.CloseTransactionBengkel_err
100               MsgBox "Gagal menyimpan data.", vbInformation
101               Exit Sub
102           End If
103       Else
104           MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
105       End If

106      On Error GoTo 0
107      Exit Sub

cmdSimpan_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
108       varErlError = Erl
109       varErrNumber = err.Number
110       If isBegin = True Then
111           Call objHT.CloseTransactionBengkel_err
112       End If
113       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmValidasiPekerjaanNHS")
End Sub

Private Sub cmdSimpanDetail_Click()
          Dim isBegin As Boolean
          Dim noRangka As String
          Dim rsJobDetail   As ADODB.Recordset
          
1        On Error GoTo cmdSimpanDetail_Click_Error

2         isBegin = False
3         If MsgBox("Apakah yakin untuk validasi data?", vbYesNo) = vbYes Then
4       Call objHT.OpenTransactionBengkel
5       isBegin = True
6       sSQL = "SELECT NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & NulSID(Text(4)) & "' AND FS_LOG_DEL = ''"
7       noRangka = objHT.FetchDataBengkel_String(sSQL)
8       sSQL = "UPDATE ReqUbahWONHS SET [status] = '1' WHERE idreq = '" & NulSID(Text(3)) & "' "
9       Call objHT.InsertData_Bengkel(sSQL)
        
10      If Not rsJobReq.EOF Then
11          rsJobReq.MoveFirst
12          rsJobReq.filter = "pilih = true"
13          If Not rsJobReq.EOF Then
14              rsJobReq.MoveFirst
15              Do While Not rsJobReq.EOF
16                  sSQL = "INSERT INTO jobdetail (no_wo,no_chassis,Date,klpkerja,rep,rte,actrte,addJob,mekanik,grp,OFFn,disc,Time,VALID,FREEC,isFP,id_program,Price)"
17                  sSQL = sSQL & " SELECT '" & NulSID(Text(4)) & "','" & noRangka & "','" & Format(isnow, "YYYY-MM-DD") & "',klpkerja,rep,rte,actrte,0,'','',OFFN,0,'','','','',PROGRAM,0 FROM ReqUbahJobWONHS WHERE idreq = '" & NulSID(Text(3)) & "' and REP = '" & NulSID(rsJobReq!jobOrder) & "' "
18                  Call objHT.InsertData_Bengkel(sSQL)
19                  sSQL = "UPDATE ReqUbahJobWONHS SET disetujui = '1',userApprove = '" & gSA & "',TglApprove = getdate() WHERE idreq = '" & NulSID(Text(3)) & "' and REP = '" & NulSID(rsJobReq!jobOrder) & "' "
20                  Call objHT.InsertData_Bengkel(sSQL)
21                  rsJobReq.MoveNext
22              Loop
23          End If
24          rsJobReq.filter = ""
25          rsJobReq.filter = "pilih = false"
26          If Not rsJobReq.EOF Then
27              rsJobReq.MoveFirst
28              Do While Not rsJobReq.EOF
29                  sSQL = "UPDATE ReqUbahJobWONHS SET disetujui = '2',userApprove = '" & gSA & "',TglApprove = getdate() WHERE idreq = '" & NulSID(Text(3)) & "' and REP = '" & NulSID(rsJobReq!jobOrder) & "' "
30                  Call objHT.InsertData_Bengkel(sSQL)
31                  rsJobReq.MoveNext
32              Loop
33          End If
34          rsJobReq.filter = ""
35      End If
        
        ''proses update jasa ke jobdta
36      sSQL = "SELECT * FROM jobdetail WITH (NOLOCK) WHERE no_wo = '" & NulSID(Text(4)) & "' AND no_chassis = '" & noRangka & "' "
37      Set rsJobDetail = objHT.FetchDataBengkel(sSQL)
38      If Not rsJobDetail.EOF Then
39          rsJobDetail.MoveFirst
40          Call ClearsSql
41          i = 1
42          kolom = 0
43          total = 0
44          Do While Not rsJobDetail.EOF
45                gsSql1(kolom) = "REP" & i
46                gsSql2(kolom) = NulSID(rsJobDetail!rep)
47                kolom = kolom + 1
48                gsSql1(kolom) = "RTE" & i
49                gsSql2(kolom) = k2t(NulSID(rsJobDetail!rte))
50                kolom = kolom + 1
                  
                   
51                gsSql1(kolom) = "ACTRATE" & i
52                gsSql2(kolom) = Replace(CDbl(NulSID(rsJobDetail!actrte)), ",", ".")
53                kolom = kolom + 1
                  
                  ' GROUP JOB
54                gsSql1(kolom) = "KLPKERJA" & i
55                gsSql2(kolom) = NulSID(rsJobDetail!klpkerja)
56                kolom = kolom + 1
                  
                  ' NILAI HARGA JOB
57                gsSql1(kolom) = "OFF" & i
58                gsSql2(kolom) = hitungLabour(NulSID(rsJobDetail!rte), objHT.FetchDataBengkel_String("SELECT MODEL FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & NulSID(Text(4)) & "' AND FS_LOG_DEL = ''"))  '' NulSID(rsJobDetail!OFFn)
59                total = total + gsSql2(kolom)
60                kolom = kolom + 1
                  
                  ' MECH YG MENGERJAKAN PER JOB
      '            gsSql1(5 + i) = "MEKANIK" & i
      '            gsSql2(5 + i) = dGrid.TextMatrix(i, 6)
      '
      '            Set rsHlp = objO.FetchData("SELECT [GROUP] FROM MEKANIK WITH (NOLOCK) WHERE MECH = '" & Trim(gsSql2(69 + i)) & "'")
      '            If Not rsHlp.EOF Then
      '                gsSql1(6 + i) = "GRP" & i
      '                gsSql2(6 + i) = rsHlp("GROUP")
      '            Else
      '                gsSql1(6 + i) = "GRP" & i
      '                gsSql2(6 + i) = ""
      '            End If
61              i = i + 1
62              rsJobDetail.MoveNext
63          Loop
64          gsSql1(kolom) = "LABOR"
65          gsSql2(kolom) = total
66          sSQL = GensSql2("JOBDTA", True) + " ,TOTAL = (TOTAL - LABOR)+" & total & "  WHERE no_chassis = '" & NulSID(noRangka) & "' AND no_WO = '" & Trim(Text(4)) & "'"
67          Call objHT.InsertData_Bengkel(sSQL)
68      End If
69      If Not rsPartReq.EOF Then
70          rsPartReq.MoveFirst
71          rsPartReq.filter = "pilih = true"
72          If Not rsPartReq.EOF Then
73              rsPartReq.MoveFirst
74              Do While Not rsPartReq.EOF
75                  sSQL = "INSERT INTO JOBPART (NO_WO,NO_CHASSIS,KODE_ITEM,NAMA_ITEM,SAT_ITEM,QTY)"
76                  sSQL = sSQL & " SELECT '" & NulSID(Text(4)) & "','" & noRangka & "',KODE_ITEM,NAMA_ITEM,SAT_ITEM,QTY FROM ReqUbahPARTWONHS WHERE idreq = '" & NulSID(Text(3)) & "' and KODE_ITEM = '" & NulSID(rsPartReq!KodeItem) & "' "
77                  Call objHT.InsertData_Bengkel(sSQL)
78                  sSQL = "UPDATE ReqUbahPARTWONHS SET disetujui = '1',userApprove = '" & gSA & "',TglApprove = getdate() WHERE idreq = '" & NulSID(Text(3)) & "' and kode_item = '" & NulSID(rsPartReq!KodeItem) & "' "
79                  Call objHT.InsertData_Bengkel(sSQL)
80                  rsPartReq.MoveNext
81              Loop
82          End If
83          rsPartReq.filter = ""
84          rsPartReq.filter = "pilih = false"
85          If Not rsPartReq.EOF Then
86              rsPartReq.MoveFirst
87              Do While Not rsJobReq.EOF
88                  sSQL = "UPDATE ReqUbahPARTWONHS SET disetujui = '2',userApprove = '" & gSA & "',TglApprove = getdate() WHERE idreq = '" & NulSID(Text(3)) & "' and kode_item = '" & NulSID(rsPartReq!KodeItem) & "' "
89                  Call objHT.InsertData_Bengkel(sSQL)
90                  rsPartReq.MoveNext
91              Loop
92          End If
93          rsPartReq.filter = ""
94      End If
        
        
95      If Not rsSubletReq.EOF Then
96          rsSubletReq.MoveFirst
97          rsSubletReq.filter = "pilih = true"
98          If Not rsSubletReq.EOF Then
99              rsSubletReq.MoveFirst
100             Do While Not rsSubletReq.EOF
101                 sSQL = "INSERT INTO JOBSUBLET (NO_WO,NO_CHASSIS,KodeItem,NAMABARANG,KODESATUAN,qty,Price,KDPROMO)"
102                 sSQL = sSQL & " SELECT '" & NulSID(Text(4)) & "','" & noRangka & "',KodeItem,NAMABARANG,KODESATUAN,qty,hARGa,KDPROMO FROM ReqUbahsubletWONHS WHERE idreq = '" & NulSID(Text(3)) & "' and kodeitem = '" & NulSID(rsSubletReq!KDSUBLET) & "' "
103                 Call objHT.InsertData_Bengkel(sSQL)
104                 sSQL = "UPDATE ReqUbahsubletWONHS SET disetujui = '1',userApprove = '" & gSA & "',TglApprove = getdate() WHERE idreq = '" & NulSID(Text(3)) & "' and kodeitem = '" & NulSID(rsSubletReq!KDSUBLET) & "' "
105                 Call objHT.InsertData_Bengkel(sSQL)
106                 rsSubletReq.MoveNext
107             Loop
108         End If
109         rsSubletReq.filter = ""
110         rsSubletReq.filter = "pilih = fALse"
111         If Not rsSubletReq.EOF Then
112             rsSubletReq.MoveFirst
113             Do While Not rsSubletReq.EOF
114                 sSQL = "UPDATE ReqUbahsubletWONHS SET disetujui = '2',userApprove = '" & gSA & "',TglApprove = getdate() WHERE idreq = '" & NulSID(Text(3)) & "' and kodeitem = '" & NulSID(rsSubletReq!KDSUBLET) & "' "
115                 Call objHT.InsertData_Bengkel(sSQL)
116                 rsSubletReq.MoveNext
117             Loop
118         End If
119         rsSubletReq.filter = ""
120     End If
121     Call objHT.CloseTransactionBengkel
122     MsgBox "Data berhasil disimpan.", vbInformation
123     Frame3.Visible = False
124     Call cmdRefresh_Click
        
125       End If

126      On Error GoTo 0
127      Exit Sub

cmdSimpanDetail_Click_Error:
128       If isBegin Then
129     Call objHT.CloseTransactionBengkel_err
130       End If
131       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpanDetail_Click of Form frmValidasiPekerjaanNHS"


End Sub

Private Sub cmdSimpanPerubahan_Click()

1        On Error GoTo cmdSimpanPerubahan_Click_Error
          Dim isBegin     As Boolean
          Dim boleh       As Boolean
          Dim pesan       As String
          Dim rsCek       As ADODB.Recordset

          Dim KodeUnit    As String
          Dim nik         As String
          Dim FS_KD_PETUGAS   As String
          Dim rsPetugas   As ADODB.Recordset

2         boleh = True
3         isBegin = False

4         If NulSID(txtAlasan.Text) = "" Then
5             boleh = False
6             pesan = pesan & " - Alasan harus diisi." & vbCrLf
7         ElseIf Len(NulSID(txtAlasan.Text)) < 5 Then
8             boleh = False
9             pesan = pesan & " - Alasan minimal 5 karakter." & vbCrLf
10        End If

11        If boleh Then
12            Call objHT.OpenTransactionBengkel
13            isBegin = True
14            Do While Not rsGrid.EOF
15                sSQL = "SELECT FS_KD_PETUGAS AS KODE,NIK,b.kodeunit as kodeunit FROM PlottingPetugasNHS a WITH (NOLOCK) inner join MstUnitNHS b with (nolock) on a.KodeUnit = b.KodeUnit and a.KodeCabang = b.KodeCabang WHERE a.KodeCabang = '" & gDb & "' and FS_KD_PETUGAS = '" & NulSID(rsGrid!Teknisi) & "'"
16                Set rsPetugas = objHT.FetchDataBengkel(sSQL)
17                If Not rsPetugas.EOF Then
                  
18                    sSQL = "INSERT INTO LogPlottingbookingNHS (NOFAK,MDL,NO_POLISI,NO_ENGINE,KodeUnit,NIK,FS_KD_PETUGAS,UserEntry,TglEntry,cetak,UserCetak,TglCetak,kirim,TglKirim,UserKirim,[Status],TglStatus,UserEdit, TglEdit,AlasanEdit)"
19                    sSQL = sSQL & " SELECT NOFAK,MDL,NO_POLISI,NO_ENGINE,KodeUnit,NIK,FS_KD_PETUGAS,UserEntry,TglEntry,cetak,UserCetak,TglCetak,kirim,TglKirim,UserKirim,[Status],TglStatus,'" & gSA & "',getdate(),'" & NulSID(txtAlasan.Text) & "' FROM PlottingbookingNHS WITH (NOLOCK) WHERE NOFAK = '" & NulSID(rsGrid!NoFak) & "' and NO_POLISI = '" & NulSID(rsGrid!NoPolisi) & "'"
20                    Call objHT.InsertData_Bengkel(sSQL)
                      
21                    sSQL = "UPDATE PlottingbookingNHS SET KodeUnit = '" & NulSID(rsPetugas!KodeUnit) & "',NIK = '" & NulSID(rsPetugas!nik) & "',FS_KD_PETUGAS = '" & NulSID(rsPetugas!kode) & "' WHERE NOFAK = '" & NulSID(rsGrid!NoFak) & "' and NO_POLISI = '" & NulSID(rsGrid!NoPolisi) & "'"
22                    Call objHT.InsertData_Bengkel(sSQL)
23                End If
24                rsGrid.MoveNext
25            Loop
26            pesan = "Data berhasil diubah."
27            rsGrid.filter = ""

28            Call objHT.CloseTransactionBengkel
29            MsgBox pesan, vbInformation
30            Frame2.Visible = False
31            cmdSimpan.Enabled = True
32            Call cmdRefresh_Click
33        Else
34            MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
35        End If
36       On Error GoTo 0
37       Exit Sub

cmdSimpanPerubahan_Click_Error:

38        If isBegin Then
39            Call objHT.CloseTransactionBengkel_err
40        End If
41        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpanPerubahan_Click of Form frmValidasiPekerjaanNHS"


End Sub

 

Private Sub cmdTutup_Click()

    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False)
   '' Call refreshGrid("")
End Sub


Private Sub gridnhs_FilterChange()
On Error GoTo errHandler
    Set cols = gridNHS.Columns
    Dim C As Integer

    C = gridNHS.col
    gridNHS.HoldFields
    rsGrid.filter = getFilter()
    gridNHS.col = C
    gridNHS.EditActive = True
    Exit Sub
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub gridnhs_HeadClick(ByVal ColIndex As Integer)

1        On Error GoTo gridnhs_HeadClick_Error
2         Urut = IIf(Urut = "DESC", "ASC", "DESC")
3         rsGrid.Sort = gridNHS.Columns(ColIndex).DataField & " " & Urut
4         gridNHS.SetFocus
5         Sendkeys "{Right " & Val(ColIndex) & "}"

6        On Error GoTo 0
7        Exit Sub

gridnhs_HeadClick_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
8         varErlError = Erl
9         varErrNumber = err.Number
10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure gridnhs_HeadClick of Form frmValidasiPekerjaanNHS")
End Sub

Private Sub gridnhs_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo gridnhs_RowColChange_Error
    If Not rsGrid.EOF Then
        sSQL = "SELECT KodeUnit + ' - ' + NoPol AS UNIT,KetModel,KodeUnit FROM mstunitnhs WITH (NOLOCK) where kodeunit = '" & NulSID(rsGrid!KodeUnit) & "'"
        DcKodeNhs.value = objO.FetchData_String(sSQL)
        sSQL = "SELECT FS_NM_PETUGAS + ' | ' + nik as petugas,FS_KD_PETUGAS,nik FROM T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA' and ISNULL(nik,'') = '" & NulSID(rsGrid!nik) & "'"
        DCPetugas.value = objO.FetchData_String(sSQL)
    End If
   On Error GoTo 0
   Exit Sub
gridnhs_RowColChange_Error:
End Sub

Private Sub refreshGrid(kriteria As String)
    sSQL = "SELECT a.idreq,a.nowo,a.mdl,a.no_polisi,b.NOFAK,a.customer,a.kodeunit,a.fs_kd_petugas,a.tglentry,b.fmdl"
    sSQL = sSQL & " FROM ReqUbahWONHS a WITH (NOLOCK)"
    sSQL = sSQL & " inner join JOBDTA b with (nolock) on a.nowo = b.NO_WO and b.FS_LOG_DEL = ''"
    sSQL = sSQL & " WHERE b.noinv = '' and b.JOB = 3 and a.[status] = 0 " & kriteria
    Set rsTemp = objO.FetchData(sSQL)
    Call createTableGrid
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            With rsGrid
                .AddNew
                !Pilih = False
                !User = NulSID(rsTemp!FS_KD_PETUGAS)
                !ReqBook = NulSID(rsTemp!tglentry)
                !customer = NulSID(rsTemp!customer)
                !idreq = NulSID(rsTemp!idreq)
                !noWO = NulSID(rsTemp!noWO)
                !model = NulSID(rsTemp!Fmdl)
                !NoPolisi = NulSID(rsTemp!NO_POLISI)
                !NoFak = NulSID(rsTemp!NoFak)
                !MDL = NulSID(rsTemp!MDL)
                .Update
            End With
            rsTemp.MoveNext
        Loop
    End If
    Set gridNHS.DataSource = rsGrid
End Sub
 
Private Sub isiId()
'    Dim noUrut As Integer
'    Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUBLET WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY RIGHT(KDSUBLET,10) DESC")
'    If Not rsTemp.EOF Then
'        noUrut = CDbl(Right(rsTemp("KDSUBLET"), 4))
'        noUrut = noUrut + 1
'        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
'    Else
'        noUrut = 1
'        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
'    End If
End Sub

  
Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
   Dim col As TrueOleDBGrid70.Columns

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
  
