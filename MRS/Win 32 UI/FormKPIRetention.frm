VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "arprogbar.ocx"
Begin VB.Form FormKPIRetention 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   12750
   ShowInTaskbar   =   0   'False
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
      Left            =   8940
      Picture         =   "FormKPIRetention.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pr&ocess"
      Height          =   732
      Left            =   7860
      Picture         =   "FormKPIRetention.frx":0790
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7320
      Width           =   732
   End
   Begin ucXPButton.XPButton thnMinus 
      Height          =   495
      Left            =   885
      TabIndex        =   1
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormKPIRetention.frx":0DFD
      PICN            =   "FormKPIRetention.frx":0E19
      PICH            =   "FormKPIRetention.frx":1024
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton blnMinus 
      Height          =   495
      Left            =   8655
      TabIndex        =   2
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormKPIRetention.frx":122F
      PICN            =   "FormKPIRetention.frx":124B
      PICH            =   "FormKPIRetention.frx":1456
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton thnPlus 
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormKPIRetention.frx":1661
      PICN            =   "FormKPIRetention.frx":167D
      PICH            =   "FormKPIRetention.frx":1874
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton blnPlus 
      Height          =   495
      Left            =   11010
      TabIndex        =   4
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormKPIRetention.frx":1A6B
      PICN            =   "FormKPIRetention.frx":1A87
      PICH            =   "FormKPIRetention.frx":1C7E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5880
      Left            =   690
      TabIndex        =   7
      Top             =   1215
      Width           =   9495
      Begin TrueDBGrid70.TDBGrid gridNotValid 
         Height          =   1935
         Left            =   105
         TabIndex        =   30
         Top             =   3825
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3413
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Wiraniaga"
         Columns(0).DataField=   "NamaSales"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Supervisor"
         Columns(1).DataField=   "Nama_Supervisor"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Jumlah Data"
         Columns(2).DataField=   "total"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KodeSales"
         Columns(3).DataField=   "KodeSales"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Kode_Supervisor"
         Columns(4).DataField=   "Kode_Supervisor"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   953
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4419"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4339"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=4419"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4339"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1693"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         DeadAreaBackColor=   15790320
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin VB.Label KPI_5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8310
         TabIndex        =   29
         Top             =   2685
         Width           =   1005
      End
      Begin VB.Label KPI_6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8310
         TabIndex        =   28
         Top             =   2970
         Width           =   1005
      End
      Begin VB.Label KPI_4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8325
         TabIndex        =   27
         Top             =   2025
         Width           =   1005
      End
      Begin VB.Label KPI_2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8325
         TabIndex        =   26
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label KPI_3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8325
         TabIndex        =   25
         Top             =   1740
         Width           =   1005
      End
      Begin VB.Label KPI_1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8325
         TabIndex        =   24
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOT VALID - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -30
         TabIndex        =   23
         Top             =   3015
         Width           =   1545
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah total data tidak Valid hasil Reminder MRA dari data yang telah di transfer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1620
         TabIndex        =   22
         Top             =   3000
         Width           =   6195
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VALID - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -30
         TabIndex        =   21
         Top             =   2715
         Width           =   1545
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Reminder MRA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   90
         TabIndex        =   20
         Top             =   2355
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah total data Valid hasil Reminder MRA dari data yang telah di transfer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1620
         TabIndex        =   19
         Top             =   2715
         Width           =   6195
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah data pelanggan yang unitnya telah delivery ditransfer ke after sales melalui Clock Off PDS Man"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   18
         Top             =   720
         Width           =   7935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Transfer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   420
         Width           =   2895
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOT YET PDFU - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -60
         TabIndex        =   16
         Top             =   2040
         Width           =   1545
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah total data yang belum di PDFU dari data yang telah di transfer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1590
         TabIndex        =   15
         Top             =   2025
         Width           =   6195
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NOT VALID - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -60
         TabIndex        =   14
         Top             =   1755
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah total data tidak Valid hasil PDFU dari data yang telah di transfer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1590
         TabIndex        =   13
         Top             =   1740
         Width           =   6195
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "VALID - "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -60
         TabIndex        =   12
         Top             =   1455
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ".:: SALES DATABASE KPI DATA TRANSFER ::."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   255
         TabIndex        =   10
         Top             =   105
         Width           =   3930
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   ".:: Reminder MRA DATA TIDAK VALID ::."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   285
         TabIndex        =   11
         Top             =   3405
         Width           =   3855
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   9315
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   375
         Left            =   45
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   9315
      End
      Begin VB.Label ket1 
         BackStyle       =   0  'Transparent
         Caption         =   "PDFU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   1095
         Width           =   2895
      End
      Begin VB.Label ket1a 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah total data Valid hasil PDFU dari data yang telah di transfer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1590
         TabIndex        =   8
         Top             =   1455
         Width           =   6195
      End
   End
   Begin ARProgBarCtrl.ARProgressBar ProBar 
      Height          =   375
      Left            =   1425
      TabIndex        =   31
      Top             =   7680
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      Value           =   50
      ForeColor       =   65280
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   8655
      Left            =   30
      Picture         =   "FormKPIRetention.frx":1E75
      Top             =   75
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   8820
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   975
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7740
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   975
   End
   Begin VB.Shape ShapeBar 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1305
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label LabBar 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Collecting KPI : ... "
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
      Left            =   1425
      TabIndex        =   34
      Top             =   7320
      Visible         =   0   'False
      Width           =   6120
   End
   Begin VB.Label nBulan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AGUSTUS"
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
      Left            =   9255
      TabIndex        =   6
      Top             =   780
      Width           =   1695
   End
   Begin VB.Label nTahun 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Periode  :  2009"
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
      Left            =   1545
      TabIndex        =   5
      Top             =   735
      Width           =   2295
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   435
      Left            =   1365
      Shape           =   4  'Rounded Rectangle
      Top             =   630
      Width           =   2595
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   135
      X2              =   11775
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   435
      Left            =   9135
      Shape           =   4  'Rounded Rectangle
      Top             =   690
      Width           =   1875
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KPI RETENTION"
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
      TabIndex        =   0
      Top             =   180
      Width           =   9975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   3960
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   1440
      X2              =   3960
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   3960
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   510
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   11715
   End
End
Attribute VB_Name = "FormKPIRetention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OnLoad As Boolean

Private rs1 As ADODB.Recordset
Private rs2 As ADODB.Recordset
Private rs3 As ADODB.Recordset

Private rsKPI As New ADODB.Recordset

Private Bulan(12) As String
Private Bln As Integer
Private Thn As Integer
Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdProcess_Click()
    On Error Resume Next
    Dim nilai As Double
    cmdProcess.Enabled = False
    
    LabBar.Visible = True
    ProBar.Visible = True
    ShapeBar.Visible = True
    
    ProBar.Min = 0
    ProBar.Max = 7
    ProBar.value = 0
    ProBar.Caption = "0 of 7 ... Collection KPI"
    DoEvents
    
    ''2017-09-29 || mengambil data total transfer
    
    sSQL = "select nospk from [pds" & Right(gDb, 2) & "]..JadwalDriver where month(TglSampaiAct) = " & Bln & " and YEAR(tglsampaiact) = " & Thn
    Set rs1 = objO.FetchData(sSQL)
    nilai = 0
    If Not rs1.EOF Then
        nilai = rs1.RecordCount
    End If
    KPI_1.Caption = nilai
    ProBar.value = 1
    ProBar.Caption = "1 of 7 ... Total Data Transfer"
    DoEvents
    
    sSQL = "select COUNT(okcr1) total from t_transferCr1 where OkCr1 = 1 and nospk in (select nospk from [pds" & Right(gDb, 2) & "]..JadwalDriver where month(TglSampaiAct) = " & Bln & " and YEAR(tglsampaiact) = " & Thn & ")"
    Set rs1 = objO.FetchDataSPK(sSQL)
    nilai = 0
    If Not rs1.EOF Then
        nilai = Nul(rs1!total)
    End If
    KPI_2.Caption = nilai
    ProBar.value = 2
    ProBar.Caption = "2 of 7 ... Total Data Transfer PDFU Valid"
    DoEvents
    
    sSQL = "select COUNT(okcr1) total from t_transferCr1 where OkCr1 = 0 and nospk in (select nospk from [pds" & Right(gDb, 2) & "]..JadwalDriver where month(TglSampaiAct) = " & Bln & " and YEAR(tglsampaiact) = " & Thn & ")"
    Set rs1 = objO.FetchDataSPK(sSQL)
    nilai = 0
    If Not rs1.EOF Then
        nilai = Nul(rs1!total)
    End If
    KPI_3.Caption = nilai
    ProBar.value = 3
    ProBar.Caption = "3 of 7 ... Total Data Transfer PDFU Not Valid"
    DoEvents
    
    KPI_4.Caption = CDbl(KPI_1.Caption) - (CDbl(KPI_2.Caption) + CDbl(KPI_3.Caption))
    ProBar.value = 4
    ProBar.Caption = "4 of 7 ... Total Data Transfer Not Yet PDFU Not Valid"
    DoEvents
    
    sSQL = "select COUNT(NoRangka) total"
    sSQL = sSQL & " From M_ListReminderCall"
    sSQL = sSQL & " where Pesan_Telp not like '%Nomor tidak valid%' and NoRangka in ("
    sSQL = sSQL & " select (b.NoRangka)"
    sSQL = sSQL & " from [pds" & Right(gDb, 2) & "]..JadwalDriver a"
    sSQL = sSQL & " inner join [SPK" & Right(gDb, 2) & "]..SPK b on a.nospk = b.SPK_No"
    sSQL = sSQL & " Where Month(tglsampaiact) = " & Bln & " And year(tglsampaiact) = " & Thn
    sSQL = sSQL & " )"
    Set rs1 = objO.FetchData(sSQL)
    nilai = 0
    If Not rs1.EOF Then
        nilai = rs1!total
    End If
    KPI_5.Caption = nilai
    ProBar.value = 5
    ProBar.Caption = "5 of 7 ... Total Reminder MRA Valid"
    DoEvents
     
    sSQL = "select COUNT(NoRangka) total"
    sSQL = sSQL & " From M_ListReminderCall"
    sSQL = sSQL & " where Pesan_Telp like '%Nomor tidak valid%' and NoRangka in ("
    sSQL = sSQL & " select (b.NoRangka)"
    sSQL = sSQL & " from [pds" & Right(gDb, 2) & "]..JadwalDriver a"
    sSQL = sSQL & " inner join [SPK" & Right(gDb, 2) & "]..SPK b on a.nospk = b.SPK_No"
    sSQL = sSQL & " Where Month(tglsampaiact) = " & Bln & " And year(tglsampaiact) = " & Thn
    sSQL = sSQL & " )"
    Set rs1 = objO.FetchData(sSQL)
    nilai = 0
    If Not rs1.EOF Then
        nilai = rs1!total
    End If
    KPI_6.Caption = nilai
    ProBar.value = 6
    ProBar.Caption = "6 of 7 ... Total Reminder MRA Not Valid"
    DoEvents
    
    ProBar.value = 7
    ProBar.Caption = "7 of 7 ... Collect Data Reminder MRA Not Valid"
    DoEvents
    
    sSQL = "select count(a.NoRangka) total ,b.NamaSales ,c.Nama_Supervisor ,b.KodeSales,b.Kode_Supervisor"
    sSQL = sSQL & " from [SPK" & Right(gDb, 2) & "]..spk a"
    sSQL = sSQL & " left join [SPK" & Right(gDb, 2) & "]..ArSales b on a.Kode_Sales = b.KodeSales"
    sSQL = sSQL & " left join [SPK" & Right(gDb, 2) & "]..ArSupervisor c on b.Kode_Supervisor = c.Kode_Supervisor"
    sSQL = sSQL & " Where SPK_Status = 1"
    sSQL = sSQL & " and NoRangka in ("
    sSQL = sSQL & " select (NoRangka) total"
    sSQL = sSQL & " From M_ListReminderCall"
    sSQL = sSQL & " where Pesan_Telp like '%Nomor tidak valid%' and NoRangka in ("
    sSQL = sSQL & " select (b.NoRangka)"
    sSQL = sSQL & " from [pds" & Right(gDb, 2) & "]..JadwalDriver a"
    sSQL = sSQL & " inner join [SPK" & Right(gDb, 2) & "]..SPK b on a.nospk = b.SPK_No"
    sSQL = sSQL & " Where Month(tglsampaiact) = " & Bln & " And year(tglsampaiact) = " & Thn
    sSQL = sSQL & " ))"
    sSQL = sSQL & " group by b.NamaSales ,c.Nama_Supervisor ,b.KodeSales,b.Kode_Supervisor"
    sSQL = sSQL & " order by b.Kode_Supervisor"

    Set rs1 = objO.FetchData(sSQL)
    gridNotValid.DataSource = rs1
    gridNotValid.Refresh
    
    LabBar.Visible = False
    ProBar.Visible = False
    ShapeBar.Visible = False
    cmdProcess.Enabled = True
End Sub

Private Sub Form_Load()
    Call FillAllCombo
    Bln = Month(isnow) 'Month(Now)
    Thn = year(isnow)  'Year(Now)
    nBulan.Caption = Bulan(Bln)
    nTahun.Caption = "Periode  :  " & Thn
End Sub

Private Sub FillAllCombo()
   Bulan(0) = "TOTAL"
   Bulan(1) = "JANUARI"
   Bulan(2) = "FEBRUARI"
   Bulan(3) = "MARET"
   Bulan(4) = "APRIL"
   Bulan(5) = "MEI"
   Bulan(6) = "JUNI"
   Bulan(7) = "JULI"
   Bulan(8) = "AGUSTUS"
   Bulan(9) = "SEPTEMBER"
   Bulan(10) = "OKTOBER"
   Bulan(11) = "NOVEMBER"
   Bulan(12) = "DESEMBER"
End Sub
  

Private Sub thnPlus_Click()
    Thn = Thn + 1
    nTahun.Caption = "Periode  :  " & Thn
    'Call cmdProcess_Click
End Sub

Private Sub thnMinus_Click()
    Thn = Thn - 1
    If Thn < 2009 Then
        Thn = 2009
    Else
        nTahun.Caption = "Periode  :  " & Thn
        'Call cmdProcess_Click
    End If
End Sub

Private Sub blnMinus_Click()
    Bln = Bln - 1
    If Bln = -1 Then Bln = 12
    nBulan.Caption = Bulan(Bln)
    'Call cmdProcess_Click
End Sub

Private Sub blnPlus_Click()
    Bln = Bln + 1
    If Bln = 13 Then Bln = 0
    nBulan.Caption = Bulan(Bln)
    'Call cmdProcess_Click
End Sub
