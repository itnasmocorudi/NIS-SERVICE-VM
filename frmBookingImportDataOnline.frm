VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmBookingImportDataOnline 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "hari"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin TrueOleDBGrid70.TDBGrid ListImport 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5106
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
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
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
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
      Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
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
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
   End
   Begin AIFCmp1.asxPowerButton tblImport 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "IMPORT DATA ONLINE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton tblSimpan 
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "SIMPAN"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton tblBatal 
      Height          =   495
      Left            =   8520
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "BATAL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin VB.Label lblOT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "DATA BOOKING ONLINE BELUM TRANSFER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1365
      Width           =   9735
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   9840
      X2              =   9240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   9840
      X2              =   9000
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   9840
      X2              =   8880
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORT DATA BOOKING ONLINE HO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmBookingImportDataOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsImport As ADODB.Recordset
Dim rsTempImport As New ADODB.Recordset

Private Function createImport() As Boolean
    On Error GoTo vb_error
    
    If rsTempImport.State Then rsTempImport.Close
    Set rsTempImport = Nothing
    With rsTempImport
        .Fields.Append "No", adInteger
        .Fields.Append "Import", adBoolean
        .Fields.Append "Id", adInteger
        .Fields.Append "NoPolisi", adVarChar, 25
        .Fields.Append "NamaCustomer", adVarChar, 50
        .Fields.Append "JenisKendaraan", adVarChar, 50
        .Fields.Append "TanggalBooking", adVarChar, 25
        .Fields.Append "JamBooking", adVarChar, 25
        .Fields.Append "NoTel", adVarChar, 25
        .Fields.Append "Email", adVarChar, 100
        .Fields.Append "Kota", adVarChar, 25
        ' ==================================================================================
        ' 14 Feb 2017 | THANIA | update tambahan keluhan, job order & model kendaraan
        ' ==================================================================================
        .Fields.Append "jenis_Service", adVarChar, 50
        .Fields.Append "Service_lain", adVarChar, 50
        .Fields.Append "keluhan", adVarChar, 200
        ' ==================================================================================
        
        '' =================================================================================
        '' 2018-04-23 penambahan nobookingapps dan kodepromoapp
        '' =================================================================================
        .Fields.Append "nobookingapps", adVarChar, 100
        .Fields.Append "kodepromoapp", adVarChar, 100
        ' ==================================================================================
        ' 04-05-2018 update kode booking cars world
        ' ----------------------------------------------------------------------------------
        .Fields.Append "noBookCarsWorld", adVarChar, 100
        
        .Open
    End With
    If rsTempImport.State Then createImport = True
    ListImport.DataSource = rsTempImport
    Call headColumns
    Exit Function
    
vb_error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function


Private Sub headColumns()
    With ListImport
        .Columns(0).Caption = "No"
        .Columns(0).DataField = "No"
        .Columns(0).Width = 600
        
        .Columns(1).Caption = "Import"
        .Columns(1).DataField = "Import"
        .Columns(1).Width = 500
        
        .Columns(3).Caption = "No Polisi"
        .Columns(3).DataField = "NoPolisi"
        .Columns(3).Width = 1000
        
        .Columns(2).Caption = "ID Booking"
        .Columns(2).DataField = "ID"
        .Columns(2).Width = 1000
        
        .Columns(4).Caption = "Nama Customer"
        .Columns(4).DataField = "NamaCustomer"
        .Columns(4).Width = 2000
        
        .Columns(5).Caption = "Jenis Kendaraan"
        .Columns(5).DataField = "JenisKendaraan"
        .Columns(5).Width = 2000
        
        .Columns(6).Caption = "Tanggal Booking"
        .Columns(6).DataField = "TanggalBooking"
        .Columns(6).Width = 1000
        
        .Columns(7).Caption = "Jam Booking"
        .Columns(7).DataField = "JamBooking"
        .Columns(7).Width = 1000
        
        .Columns(8).DataField = "NoTel"
        .Columns(8).Width = 1000
        .Columns(8).Visible = False
        
        .Columns(9).Visible = False
        .Columns(9).DataField = "Email"
        .Columns(9).Width = 1000
        
        .Columns(10).Visible = False
        .Columns(10).DataField = "Kota"
        .Columns(10).Width = 1000
    End With
End Sub

Private Sub refreshGrid()
          Dim i As Integer
1         If Log.RPLHistori = 0 Then
2             Set rsImport = objO.FetchDataHO("select Notel,Email,Kota,convert(varchar,tanggal_booking,103) tanggalbooking,convert(varchar,jam_booking,108) jambooking,id,nopol as NoPolisi,nama_pelanggan as NamaCustomer,jenis_mobil as JenisKendaraan, jenis_service, service_lain, keluhan,nobookingapps,kodepromoapp,noBookCarsWorld from bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
3         Else
4             If Log.flagOffBookingLocal = 1 Then
5                 Set rsImport = objO.FetchDataHO("select Notel,Email,Kota,convert(varchar,tanggal_booking,103) tanggalbooking,convert(varchar,jam_booking,108) jambooking,id,nopol as NoPolisi,nama_pelanggan as NamaCustomer,jenis_mobil as JenisKendaraan, jenis_service, service_lain, keluhan,nobookingapps,kodepromoapp,noBookCarsWorld from bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
6             Else
7                 Set rsImport = objO.FetchData("select Notel,Email,Kota,convert(varchar,tanggal_booking,103) tanggalbooking,convert(varchar,jam_booking,108) jambooking,id,nopol as NoPolisi,nama_pelanggan as NamaCustomer,jenis_mobil as JenisKendaraan, jenis_service, service_lain, keluhan,nobookingapps,kodepromoapp,noBookCarsWorld from [" & dbReplikasiHO & "]..bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "' AND FTRANSFER = 0")
8             End If
9         End If
          
10        If rsImport.RecordCount > 0 Then
11            If createImport Then
12                i = 1
13                rsImport.MoveFirst
14                Do While Not rsImport.EOF
15                    With rsTempImport
16                        .AddNew
17                        !No = i
18                        !Import = True
19                        !id = rsImport("ID")
20                        !NoPolisi = Trim(rsImport("NoPolisi"))
21                        !namacustomer = Trim(rsImport("NamaCustomer"))
22                        !JenisKendaraan = Trim(rsImport("JenisKendaraan"))
23                        !TanggalBooking = Trim(rsImport("TanggalBooking"))
24                        !JamBooking = Trim(rsImport("JamBooking"))
25                        !Notel = Trim(rsImport("Notel"))
26                        !email = Trim(rsImport("Email"))
27                        !Kota = Trim(rsImport("Kota"))
                          ' ==================================================================================
                          ' 14 Feb 2017 | THANIA | update tambahan keluhan, job order & model kendaraan
                          ' ==================================================================================
28                        !jenis_Service = Trim(rsImport("Jenis_service"))
29                        !Service_lain = Trim(rsImport("service_lain"))
30                        !keluhan = Trim(rsImport("Keluhan"))
                          ' ==================================================================================
                          
                          ' ===============================================================================
                          ' 2018-04-23 penambahan nobookingapps dan kodepromoapp
                          ' ===============================================================================
31                        !noBookingApps = NulSID(rsImport("nobookingapps"))
32                        !kodePromoApp = NulSID(rsImport("kodepromoapp"))
                          ' ==================================================================================
                          ' 04-05-2018 update kode booking cars world
                          ' ----------------------------------------------------------------------------------
33                        !noBookCarsWorld = NulSID(rsImport!noBookCarsWorld)
34                        .Update
35                        .MoveNext
36                    End With
37                    i = i + 1
38                    rsImport.MoveNext
39                Loop
40            End If
41        Else
42            If createImport Then
43                TblSimpan.Enabled = False
44                MsgBox "Tidak ada data yang ditampilkan.", vbInformation
45            End If
46            Exit Sub
47        End If
48        ListImport.DataSource = rsTempImport
49        Call headColumns
50        TblSimpan.Enabled = True
End Sub


Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Move 150, 800, Me.Width, Me.Height
End Sub

Private Sub Form_Load()
    If Get_ConnectionProfil_Bengkel99 = False And Log.RPLHistori = 0 Then
        MsgBox "Gagal Koneksi Ke HO!", vbCritical
    Else
        Call refreshGrid
    End If
End Sub

Private Sub tblBatal_Click()
    Unload Me
    frmBookingMenuOnline.asxPowerButton3_Click
    
End Sub

Private Sub tblImport_Click()

1         If Get_ConnectionProfil_Bengkel99 = False And Log.RPLHistori = 0 Then
2             MsgBox "Gagal Koneksi Ke HO!", vbCritical
3         Else
4             Call refreshGrid
              Dim rsTransfer As ADODB.Recordset
              'Set rsTransfer = objO.FetchData("SELECT ID FROM BOOKINGONLINE WHERE fTransferHO = 0 AND FPROSES = 2")
5             Set rsTransfer = objO.FetchData("SELECT ID FROM BOOKINGONLINE WHERE fTransferHO = 0")
6             If rsTransfer.RecordCount > 0 Then
7                 rsTransfer.MoveFirst
8                 If Log.RPLHistori = 0 Then
9                     Call objHT.OpenTransactionHO
10                    erTranAll = False
11                    Do While Not rsTransfer.EOF
12                        Call objHT.InsertData_Ho("UPDATE BOOKINGONLINE SET FPROSES = 2 WHERE ID = '" & rsTransfer("ID") & "'")
13                        rsTransfer.MoveNext
14                    Loop
15                    If erTranAll = False Then
16                        Call objHT.CloseTransactionHO
17                    Else
18                        Call objHT.CloseTransactionHO_err
19                    End If
20                Else
21                    If Log.flagOffBookingLocal = 1 Then
22                        Call objHT.OpenTransactionHO
23                    Else
24                        Call objHT.OpenTransactionBengkel
25                    End If
26                    erTranAll = False
                      Dim i As String
                     ' rsTransfer.Requery
27                    Do While Not rsTransfer.EOF
28                        i = rsTransfer!id
29                        sSQL = ""
      '                    sSQL = "UPDATE [" & dbReplikasiHO & "]..BOOKINGONLINE SET FPROSES = 2 WHERE ID = '" & NulA(rsTransfer("ID")) & "'"
30                        If Log.flagOffBookingLocal = 1 Then
31                            sSQL = "UPDATE BENGKEL99.DBO.BOOKINGONLINE SET FPROSES = 2 WHERE ID = " & i
32                            Call objHT.InsertData_Ho(sSQL)

33                        Else
34                            sSQL = "UPDATE [" & dbReplikasiHO & "]..BOOKINGONLINE SET FPROSES = 2 WHERE ID = " & i
                              'sSQL = sSQL & i
35                            Call objHT.InsertData_Bengkel(sSQL)
36                        End If
37                        rsTransfer.MoveNext
38                    Loop
39                    If erTranAll = False Then
40                        If Log.flagOffBookingLocal = 1 Then
41                            Call objHT.CloseTransactionHO
42                        Else
43                            Call objHT.CloseTransactionBengkel
44                        End If
45                    Else
46                        If Log.flagOffBookingLocal = 1 Then
47                            Call objHT.CloseTransactionHO_err
48                        Else
49                            Call objHT.CloseTransactionBengkel_err
50                        End If
51                    End If
52                End If
53            End If
54        End If
End Sub

Private Sub tblSimpan_Click()
1         If MsgBox("Apakah yakin mengeksport data?", vbYesNo) = vbYes Then
2             If rsTempImport.RecordCount > 0 Then
                  Dim strSql As String
3                 rsTempImport.MoveFirst
4                 erTranAll = False
5                 If Log.RPLHistori = 0 Then
6                     Call objHT.OpenTransactionImportHO
7                 Else
8                     If Log.flagOffBookingLocal = 1 Then
9                         Call objHT.OpenTransactionImportHO
10                    Else
11                        Call objHT.OpenTransactionBengkel
12                    End If
                      
13                End If
14                Do While Not rsTempImport.EOF
15                    If rsTempImport("Import") Then
                          
                          ' strSql = "INSERT INTO [BookingOnline]([id],[nopol], [nama_pelanggan], [notel], [email], [jenis_mobil], [dealer], [kota], [tanggal_booking], [jam_booking],[userEntry], [tglEntry],[HasilFU],[noFak]) " & _
                          '          "VALUES ('" & rsTempImport("Id") & "','" & rsTempImport("NoPolisi") & "','" & Trim(rsTempImport("NamaCustomer")) & "','" & Trim(rsTempImport("Notel")) & "','" & Trim(rsTempImport("Email")) & "','" & Trim(rsTempImport("JenisKendaraan")) & "','" & Log.KodeDealer & "','" & Trim(rsTempImport("Kota")) & "','" & Format(Trim(rsTempImport("TanggalBooking")), "MM/DD/YYYY") & "','" & Trim(rsTempImport("JamBooking")) & "','" & gSA & "',getdate(),'','')"
                          
                          ' ==================================================================================
                          ' 14 Feb 2017 | THANIA | update tambahan keluhan, job order & model kendaraan
                          ' ==================================================================================
                          
                          ''23-04-2018 || ditambahkan nobookingapp dan kodeprogram
                          ''04-05-2018 update kode booking cars world
16                        strSql = "INSERT INTO [BookingOnline]([id],[nopol], [nama_pelanggan], [notel], [email], [jenis_mobil], [dealer], [kota], [tanggal_booking], [jam_booking],[userEntry], [tglEntry],[HasilFU],[noFak], [jenis_service], [service_lain],[keluhan],[nobookingapps],[kodepromoapp],[noBookCarsWorld]) " & _
                                   "VALUES ('" & rsTempImport("Id") & "','" & rsTempImport("NoPolisi") & "','" & Left(Trim(rsTempImport("NamaCustomer")), 50) & "','" & Trim(rsTempImport("Notel")) & "','" & Left(Trim(rsTempImport("Email")), 32) & "','" & Trim(rsTempImport("JenisKendaraan")) & "','" & Log.KodeDealer & "','" & Trim(rsTempImport("Kota")) & "','" & Format(Trim(rsTempImport("TanggalBooking")), "MM/DD/YYYY") & "','" & Trim(rsTempImport("JamBooking")) & "','" & gSA & "',getdate(),'','' " & _
                                   ",'" & Trim(rsTempImport("jenis_service")) & "','" & Trim(rsTempImport("service_lain")) & "','" & Trim(rsTempImport("keluhan")) & "','" & Trim(rsTempImport("nobookingapps")) & "','" & Trim(rsTempImport("kodepromoapp")) & "','" & NulSID(rsTempImport("noBookCarsWorld")) & "')"
                                   
17                        Call objHT.InsertData_Bengkel(strSql)
                          
18                        If Log.RPLHistori = 0 Then
19                            Call objHT.InsertDataImport_Ho("UPDATE BENGKEL99.DBO.BOOKINGONLINE SET fTransfer = 1 WHERE id = '" & rsTempImport("ID") & "'")
20                        Else
21                            If Log.flagOffBookingLocal = 1 Then
22                                Call objHT.InsertDataImport_Ho("UPDATE BENGKEL99.DBO.BOOKINGONLINE SET fTransfer = 1 WHERE id = '" & rsTempImport("ID") & "'")
23                            Else
24                                Call objHT.InsertData_Bengkel("UPDATE [" & dbReplikasiHO & "]..BOOKINGONLINE SET fTransfer = 1 WHERE id = '" & rsTempImport("ID") & "'")
25                            End If
26                        End If
27                    End If
28                    rsTempImport.MoveNext
29                Loop
30                If erTranAll = False Then
                      'Call objHT.CloseTransactionImportHO
31                    If Log.RPLHistori = 0 Then
32                        Call objHT.CloseTransactionImportHO
33                    Else
34                        If Log.flagOffBookingLocal = 1 Then
35                            Call objHT.CloseTransactionImportHO
36                        Else
37                            Call objHT.CloseTransactionBengkel
38                        End If
39                    End If
40                    MsgBox "Data berhasil ditambahkan", vbInformation
41                    Call createImport
42                    Exit Sub
43                End If
44            End If
45        End If
End Sub
