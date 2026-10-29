VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmWarranty 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WARRANTY"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   13770
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmWarranty.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   -120
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6060
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   11115
      Begin TrueOleDBGrid70.TDBGrid TData1 
         Height          =   4575
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8070
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
         Columns(5)._VlistStyle=   4
         Columns(5)._MaxComboItems=   5
         Columns(5).ValueItems(0)._DefaultItem=   0
         Columns(5).ValueItems(0).Value=   ""
         Columns(5).ValueItems(0).Value.vt=   8
         Columns(5).ValueItems(0).DisplayValue=   ""
         Columns(5).ValueItems(0).DisplayValue.vt=   8
         Columns(5).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(5).ValueItems.Count=   1
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=84,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0,.bgcolor=&H80000005&"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   1455
         TabIndex        =   1
         Top             =   540
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   1455
         TabIndex        =   2
         Top             =   240
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   1455
         TabIndex        =   3
         Top             =   840
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   3
         Left            =   8880
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
         Height          =   270
         Index           =   4
         Left            =   10080
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         Height          =   6015
         Left            =   0
         Top             =   0
         Width           =   11055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   195
         Left            =   8160
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   585
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka"
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   285
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   885
         Width           =   615
      End
   End
   Begin ucXPButton.XPButton cmdExit 
      Height          =   495
      Left            =   10440
      TabIndex        =   14
      Top             =   6240
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
      MICON           =   "frmWarranty.frx":2133
      PICN            =   "frmWarranty.frx":214F
      PICH            =   "frmWarranty.frx":2623
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
      BorderColor     =   &H000080FF&
      X1              =   11640
      X2              =   600
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label BrsPart 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12000
      TabIndex        =   9
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Brs 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12000
      TabIndex        =   8
      Top             =   480
      Width           =   645
   End
End
Attribute VB_Name = "frmWarranty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsTemp As ADODB.Recordset
Dim rsRec As New ADODB.Recordset
Dim History As New XArrayDB
Dim baris As Integer
Dim berhasil As Boolean

Private Sub BeriItem(noRgk As String, modL As String, nKol As Integer)
          
1         Call HeaderCol
2         kolom = CDbl(nKol)
          
      '    Set rsTemp = objO.FetchData("SELECT a.*, d.no_rangka FROM ho_warranty A WITH(NOLOCK) INNER JOIN ho_wardtl B WITH(NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' LEFT OUTER JOIN HISTORY C WITH(NOLOCK) ON A.ID_WARRANTY = C.KODETIPE AND C.FS_LOG_DEL = '' " _
      '         & "LEFT JOIN tempwar D WITH(NOLOCK) on a.id_warranty = d.id_warranty " _
      '         & "WHERE MDL = '" & modL & "' AND no_rangka1 <= '" & noRgk & "' AND no_Rangka2 >= '" & noRgk & "' AND ISNULL(C.KODETIPE,'') = '' " _
      '         & "ORDER BY A.id_warranty,mdl,no_Rangka1")
          
          ''2017-07-13 menghilangkan  model pada select
          
      '    If Log.RPLHistori = 0 Then
      '        Set rsTemp = objO.FetchData("SELECT a.*, d.no_rangka FROM ho_warranty A WITH(NOLOCK) INNER JOIN ho_wardtl B WITH(NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' LEFT OUTER JOIN HISTORY C WITH(NOLOCK) ON A.ID_WARRANTY = C.KODETIPE AND C.FS_LOG_DEL = '' " _
      '             & "LEFT JOIN tempwar D WITH(NOLOCK) on a.id_warranty = d.id_warranty " _
      '             & "WHERE no_rangka1 <= '" & noRgk & "' AND no_Rangka2 >= '" & noRgk & "' AND ISNULL(C.KODETIPE,'') = '' " _
      '             & "ORDER BY A.id_warranty,mdl,no_Rangka1")
      '    Else
      '        Set rsTemp = objO.FetchData("SELECT a.*, d.no_rangka FROM [" & dbReplikasiHO & "]..ho_warranty A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..ho_wardtl B WITH(NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' LEFT OUTER JOIN HISTORY C WITH(NOLOCK) ON A.ID_WARRANTY = C.KODETIPE AND C.FS_LOG_DEL = '' " _
      '             & "LEFT JOIN tempwar D WITH(NOLOCK) on a.id_warranty = d.id_warranty " _
      '             & "WHERE no_rangka1 <= '" & noRgk & "' AND no_Rangka2 >= '" & noRgk & "' AND ISNULL(C.KODETIPE,'') = '' " _
      '             & "ORDER BY A.id_warranty,mdl,no_Rangka1")
      '
      '    End If

3         If Log.RPLHistori = 0 Then
4             Set rsTemp = objO.FetchData("SELECT distinct a.*, d.no_rangka FROM ho_warranty A WITH(NOLOCK) INNER JOIN ho_wardtl B WITH(NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' LEFT OUTER JOIN HISTORY C WITH(NOLOCK) ON A.ID_WARRANTY = C.KODETIPE AND C.FS_LOG_DEL = '' " _
                   & "LEFT JOIN tempwar D WITH(NOLOCK) on a.id_warranty = d.id_warranty " _
                   & "WHERE no_rangka1 <= '" & noRgk & "' AND no_Rangka2 >= '" & noRgk & "' AND ISNULL(C.KODETIPE,'') = '' " _
                   & "ORDER BY A.id_warranty ")
5         Else
              '' ----------------------------------------------------
              '' 20180920
              '' ----------------------------------------------------
6             If Log.lockSCC = 1 Then
7                 Set rsTemp = objO.FetchData("SELECT distinct a.*, d.no_rangka FROM [" & dbReplikasiHO & "]..ho_warranty A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..ho_wardtl B WITH(NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' LEFT OUTER JOIN [" & dbReplikasiHO & "]..HISTORYnew C WITH(NOLOCK) ON C.NO_CHASSIS = '" & NulSID(noRgk) & "' AND A.ID_WARRANTY = C.KODETIPE AND C.FS_LOG_DEL = '' " _
                   & "LEFT JOIN tempwar D WITH(NOLOCK) on a.id_warranty = d.id_warranty " _
                   & "WHERE a.id_warranty not in (SELECT ID_WARRANTY FROM TEMPWAR WHERE NO_RANGKA = '" & noRgk & "') AND no_rangka1 <= '" & noRgk & "' AND no_Rangka2 >= '" & noRgk & "' AND ISNULL(C.KODETIPE,'') = '' " _
                   & "ORDER BY A.id_warranty ")
8             Else
9                 Set rsTemp = objO.FetchData("SELECT distinct a.*, d.no_rangka FROM [" & dbReplikasiHO & "]..ho_warranty A WITH(NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..ho_wardtl B WITH(NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' LEFT OUTER JOIN [" & dbReplikasiHO & "]..HISTORYnew C WITH(NOLOCK) ON C.NO_CHASSIS = '" & NulSID(noRgk) & "' AND A.ID_WARRANTY = C.KODETIPE AND C.FS_LOG_DEL = '' " _
                   & "LEFT JOIN tempwar D WITH(NOLOCK) on a.id_warranty = d.id_warranty " _
                   & "WHERE no_rangka1 <= '" & noRgk & "' AND no_Rangka2 >= '" & noRgk & "' AND ISNULL(C.KODETIPE,'') = '' " _
                   & "ORDER BY A.id_warranty ")
10            End If
              
              
11        End If
12        If Not rsTemp.EOF Then
              ' ITEM WARRANTY ADA
13            baris = CDbl(rsTemp.RecordCount)
              
14            History.Clear
15            History.ReDim 0, baris, 0, kolom
16            Call createTbl
              
17            i = 0
18            Do While Not rsTemp.EOF
                      ' WARRANTY YG BELUM PERNAH DIBUAT ORDER
19                    History(i, 0) = Nul(rsTemp("ID_WARRANTY"))      'rsRec("ID_WARRANTY")
20                    History(i, 1) = Nul(rsTemp("WARRANTY"))
21                    History(i, 2) = Nul(rsTemp("KETERANGAN"))
                      'History(i, 3) = Nul(rsTemp("TIPEWAR"))
22                    History(i, 4) = TglDDMMYYYY(rsTemp("TGLMULAI"))
23                    History(i, 5) = TglDDMMYYYY(rsTemp("TGLSELESAI"))
                      'History(i, 6) = rsTemp("DISCJASA") & " %"
                      'History(i, 7) = rsTemp("DISCPART") & " %"
                      
24                    rsRec.AddNew
25                    rsRec!id = Nul(rsTemp("ID_WARRANTY"))
26                    rsRec!warranty = Nul(rsTemp("WARRANTY"))
27                    rsRec!keterangan = Nul(rsTemp("KETERANGAN"))
                      'rsRec!tipe = Nul(rsTemp("TIPEWAR"))
28                    rsRec!TglMulai = TglDDMMYYYY(rsTemp("TGLMULAI"))
29                    rsRec!TglSelesai = TglDDMMYYYY(rsTemp("TGLSELESAI"))
                      'rsRec!DISCJASA = rsTemp("DISCJASA") & " %"
                      'rsRec!discpart = rsTemp("DISCPART") & " %"
30                    If rsTemp("NO_RANGKA") <> "" Then
31                        rsRec!Proses = True
32                    Else
33                        rsRec!Proses = False
34                    End If
35                    rsRec.Update
                      
36                i = i + 1
37                rsTemp.MoveNext
38            Loop
              
39        End If
End Sub

Private Sub HeaderCol()
    TData1.Columns(0).Caption = "ID"
    TData1.Columns(0).DataField = "ID"
    TData1.Columns(0).Width = 1200
    TData1.Columns(1).Caption = "WARRANTY"
    TData1.Columns(1).DataField = "WARRANTY"
    TData1.Columns(1).Width = 2000
    TData1.Columns(2).Caption = "KETERANGAN"
    TData1.Columns(2).DataField = "KETERANGAN"
    TData1.Columns(2).Width = 2000
    TData1.Columns(3).Caption = "TANGGAL MULAI"
    TData1.Columns(3).DataField = "TGLMULAI"
    TData1.Columns(3).Width = 1500
    TData1.Columns(4).Caption = "TANGGAL SELESAI"
    TData1.Columns(4).DataField = "TGLSELESAI"
    TData1.Columns(4).Width = 1500
    TData1.Columns(5).Caption = "PROSES"
    TData1.Columns(5).DataField = "PROSES"
    TData1.Columns(5).Width = 800
    TData1.Columns(5).Alignment = dbgCenter
End Sub

Private Sub cmdExit_Click()
    Call cekTanda
    Unload Me
End Sub

Private Sub cekTanda()
Dim rsHlp As New ADODB.Recordset
    
    If rsRec.State Then
        rsRec.MoveFirst
        Do While Not rsRec.EOF
            If rsRec("PROSES") Then
                Set rsHlp = objO.FetchData("SELECT * FROM TEMPWAR WITH(NOLOCK) WHERE NO_RANGKA = '" & Trim(text(0)) & "' AND ID_WARRANTY = '" & Trim(rsRec("ID")) & "'")
                If rsHlp.EOF Then
                    Call objO.InsertData("delete tempwar WHERE NO_RANGKA = '" & Trim(text(0)) & "'")
                    Call objO.InsertData("INSERT INTO TEMPWAR SELECT '" & Trim(text(0)) & "','" & Trim(rsRec("ID")) & "' WHERE NOT EXISTS " _
                                 & "(SELECT * FROM TEMPWAR WITH(NOLOCK) WHERE no_rangka = '" & Trim(text(0)) & "' AND id_warranty = '" & Trim(rsRec("ID")) & "')")
                    Call inputJobOrder
                    Call cekBerhasilInput(berhasil)
                End If
            End If
            rsRec.MoveNext
        Loop
    End If
End Sub

Private Sub cekBerhasilInput(berhasilInput As Boolean)
      Dim ntRate As Double
          
1         If Not berhasilInput Then
              
2             If Log.RPLHistori = 0 Then
3                 Set rsTemp = objO.FetchData("SELECT * FROM tempwar a WITH(NOLOCK) LEFT JOIN ho_warjobpart b WITH(NOLOCK) ON a.id_warranty = b.id_warranty INNER JOIN HO_WARRANTY C WITH(NOLOCK) ON A.ID_WARRANTY = C.ID_WARRANTY WHERE no_rangka = '" & Trim(text(0)) & "'")
4             Else
5                 Set rsTemp = objO.FetchData("SELECT * FROM tempwar a WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..ho_warjobpart b WITH(NOLOCK) ON a.id_warranty = b.id_warranty INNER JOIN [" & dbReplikasiHO & "]..HO_WARRANTY C WITH(NOLOCK) ON A.ID_WARRANTY = C.ID_WARRANTY WHERE no_rangka = '" & Trim(text(0)) & "'")
6             End If
7             If Not rsTemp.EOF Then
8                 Do While Not rsTemp.EOF
9                     Select Case gTag
                          Case "BOOKING"
10                            With frmBooking
11                                Select Case rsTemp("FLAG")
                                      Case "J"
12                                        If cekDouble(.dGrid, Trim(rsTemp("FS_ITEM")), Brs, 1) Then
13                                            For i = gBantuan To Brs
14                                                .dGrid.TextMatrix(i, 0) = .dGrid.TextMatrix(i + 1, 0)
15                                                .dGrid.TextMatrix(i, 1) = .dGrid.TextMatrix(i + 1, 1)
16                                                .dGrid.TextMatrix(i, 2) = .dGrid.TextMatrix(i + 1, 2)
17                                                .dGrid.TextMatrix(i, 3) = .dGrid.TextMatrix(i + 1, 3)
18                                                .dGrid.TextMatrix(i, 4) = .dGrid.TextMatrix(i + 1, 4)
19                                            Next
20                                            Brs = Brs - 1
21                                            myChassis = Trim(text(0))
22                                            .text(39) = Format(CDbl(.text(39)) - hitungLabour(CDbl(rsTemp("RATE")), .text(11)), "0#,#")
23                                        End If
24                                    Case "P"
25                                        If cekDouble(.partGrid, Trim(rsTemp("ITEMPART")), BrsPart, 2) Then
26                                            If .partGrid.TextMatrix(gBantuan, 6) Then
27                                                .text(55) = Format(CDbl(.text(55)) - CDbl(.partGrid.TextMatrix(gBantuan, 5)), "#0,0")
28                                            Else
29                                                .text(40) = Format(CDbl(.text(40)) - CDbl(.partGrid.TextMatrix(gBantuan, 5)), "#0,0")
30                                            End If
31                                            For i = gBantuan To BrsPart
32                                                .partGrid.TextMatrix(i, 0) = .partGrid.TextMatrix(i + 1, 0)
33                                                .partGrid.TextMatrix(i, 1) = .partGrid.TextMatrix(i + 1, 1)
34                                                .partGrid.TextMatrix(i, 2) = .partGrid.TextMatrix(i + 1, 2)
35                                                .partGrid.TextMatrix(i, 3) = .partGrid.TextMatrix(i + 1, 3)
36                                                .partGrid.TextMatrix(i, 4) = .partGrid.TextMatrix(i + 1, 4)
37                                                .partGrid.TextMatrix(i, 5) = .partGrid.TextMatrix(i + 1, 5)
38                                                .partGrid.TextMatrix(i, 6) = .partGrid.TextMatrix(i + 1, 6)
39                                            Next
40                                            BrsPart = BrsPart - 1
41                                        End If
42                                End Select
43                                gBantuan = Brs
44                                gBantuan2 = BrsPart
45                                .text(70) = CDbl(ntRate)
46                                .text(53) = Format(CDbl(.text(39)) + CDbl(.text(40)) + CDbl(.text(52)) + CDbl(.text(55)), "#0,#")
47                            End With
48                        Case "ORDER"
49                            With frmOrder
50                                Select Case rsTemp("FLAG")
                                      Case "J"
51                                        If cekDouble(.dGrid, Trim(rsTemp("FS_ITEM")), Brs, 1) Then
52                                            For i = gBantuan To Brs
53                                                .dGrid.TextMatrix(i, 0) = .dGrid.TextMatrix(i + 1, 0)
54                                                .dGrid.TextMatrix(i, 1) = .dGrid.TextMatrix(i + 1, 1)
55                                                .dGrid.TextMatrix(i, 2) = .dGrid.TextMatrix(i + 1, 2)
56                                                .dGrid.TextMatrix(i, 3) = .dGrid.TextMatrix(i + 1, 3)
57                                                .dGrid.TextMatrix(i, 4) = .dGrid.TextMatrix(i + 1, 4)
58                                                .dGrid.TextMatrix(i, 5) = .dGrid.TextMatrix(i + 1, 5)
59                                                .dGrid.TextMatrix(i, 6) = .dGrid.TextMatrix(i + 1, 6)
60                                            Next
61                                            myChassis = Trim(text(0))
62                                            .text(56) = Format(CDbl(.text(56)) + hitungLabour(CDbl(rsTemp("RATE")), .text(11)), "0#,#")
63                                            Brs = Brs - 1
64                                        End If
65                                    Case "P"
66                                        Call .hapusPart(rsTemp("ITEMPART"), rsTemp("NAMAPART"), 0)
67                                End Select
68                                .text(61) = Format(CDbl(.text(56)) + CDbl(.text(57)) + CDbl(.text(58)) + CDbl(.text(59)) + CDbl(.text(60)) + CDbl(.text(77)), "#0,#")
69                                .text(21) = ""
70                                .text(22) = ""
71                            End With
72                    End Select
73                    rsTemp.MoveNext
74                Loop
75            End If
76            Call objO.DeleteData("DELETE TEMPWAR WHERE no_rangka = '" & Trim(text(0)) & "'")
77        End If
End Sub

Private Sub inputJobOrder()
      Dim nMech As String
      Dim temp As String
      Dim hrg As Double
      Dim oliPart As Boolean
      Dim ntRate As Double

      '### FRANCHISE DITUTUP, KARENA PART DARI DEPO HO, BUKAN DARI ACCOUNTING ###
          
1         berhasil = True
          
2         If Log.RPLHistori = 0 Then
3             Set rsTemp = objO.FetchData("SELECT * FROM tempwar a WITH(NOLOCK) LEFT JOIN ho_warjobpart b WITH(NOLOCK) ON a.id_warranty = b.id_warranty INNER JOIN HO_WARRANTY C WITH(NOLOCK) ON A.ID_WARRANTY = C.ID_WARRANTY WHERE no_rangka = '" & Trim(text(0)) & "'")
4         Else
5             Set rsTemp = objO.FetchData("SELECT * FROM tempwar a WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..ho_warjobpart b WITH(NOLOCK) ON a.id_warranty = b.id_warranty INNER JOIN [" & dbReplikasiHO & "]..HO_WARRANTY C WITH(NOLOCK) ON A.ID_WARRANTY = C.ID_WARRANTY WHERE no_rangka = '" & Trim(text(0)) & "'")
6         End If
7         If Not rsTemp.EOF Then
8             ntRate = 0
9             Do While Not rsTemp.EOF
              
10                If gTag = "BOOKING" Then
11                    With frmBooking
12                        If Trim(rsTemp("FLAG")) = "J" Then          ' JOB ORDER
ulangi:
13                            nMech = InputBox("Input Mekanik yang mengerjakan !" & vbCr & "Job : " & rsTemp("FS_ITEM") & pesan, "MEKANIK")
14                            nMech = UCase(nMech)
15                            If nMech = "" Then
16                                temp = MsgBox("Membatalkan WARRANTY ?", vbYesNo, "VALIDASI MEKANIK")
17                                If temp = vbNo Then
18                                    GoTo ulangi
19                                Else
20                                    berhasil = False
21                                    Exit Sub
22                                End If
23                            Else
24                                If Not cekMekanik(nMech, CDbl(rsTemp("RATE")), Format(text(3), "DD/MM/YYYY"), Format(text(4), "HH:NN")) Then
25                                    pesan = gBantuan3
26                                    GoTo ulangi
27                                Else
28                                    If Not cekDouble(.dGrid, Trim(rsTemp("FS_ITEM")), Brs, 1) Then
29                                        .dGrid.TextMatrix(Brs.Caption, 1) = rsTemp("FS_ITEM")
30                                        .dGrid.TextMatrix(Brs.Caption, 2) = CDbl(rsTemp("RATE"))
31                                        .dGrid.TextMatrix(Brs.Caption, 3) = False
32                                        .dGrid.TextMatrix(Brs.Caption, 4) = "TWC"
33                                        myChassis = Trim(text(0))
34                                        .text(39) = Format(CDbl(.text(39)) + hitungLabour(CDbl(rsTemp("RATE")), .text(11), True), "0#,#")
35                                        ntRate = ntRate + CDbl(rsTemp("RATE"))
36                                        Brs = Brs + 1
37                                        .text(64) = nMech
38                                    End If
39                                End If
40                            End If
41                        ElseIf Trim(rsTemp("FLAG")) = "P" Then      ' PART
42                            hrg = 0
43                            oliPart = False
44                            If Not cekDouble(.partGrid, Trim(rsTemp("ITEMPART")), BrsPart, 2) Then
45                                .partGrid.TextMatrix(BrsPart.Caption, 1) = rsTemp("ITEMPART")
46                                .partGrid.TextMatrix(BrsPart.Caption, 2) = rsTemp("NAMAPART")
47                                .partGrid.TextMatrix(BrsPart.Caption, 3) = rsTemp("SATPART")
48                                .partGrid.TextMatrix(BrsPart.Caption, 4) = CDbl(rsTemp("QTY"))
49                                .partGrid.TextMatrix(BrsPart.Caption, 5) = CDbl(hrg)
50                                .partGrid.TextMatrix(BrsPart.Caption, 6) = oliPart
                                  
51                                If oliPart Then
52                                    .text(55) = Format(CDbl(.text(55)) + CDbl(hrg), "#0,0")
53                                Else
54                                    .text(40) = Format(CDbl(.text(40)) + CDbl(hrg), "#0,0")
55                                End If
56                                BrsPart = BrsPart + 1
57                            End If
58                        End If
59                        gBantuan = Brs
60                        gBantuan2 = BrsPart
61                        .text(71) = rsTemp("ID_WARRANTY")
62                        .text(27) = rsTemp("WARRANTY")
63                        .text(70) = CDbl(ntRate)
64                        .text(53) = Format(CDbl(.text(39)) + CDbl(.text(40)) + CDbl(.text(52)) + CDbl(.text(55)), "#0,#")
65                    End With
66                ElseIf gTag = "ORDER" Then
67                    With frmOrder
68                        If Trim(rsTemp("FLAG")) = "J" Then          ' JOB ORDER
ULANGI2:
69                            nMech = InputBox("Input Mekanik yang mengerjakan !" & vbCr & "Job : " & rsTemp("FS_ITEM") & " , " & rsTemp("JOBKET") & pesan, "MEKANIK")
70                            nMech = UCase(nMech)
71                            If nMech = "" Then
72                                temp = MsgBox("Membatalkan WARRANTY ?", vbYesNo, "VALIDASI MEKANIK")
73                                If temp = vbNo Then
74                                    GoTo ULANGI2
75                                Else
76                                    berhasil = False
77                                    Exit Sub
78                                End If
79                            Else
80                                If Not cekMekanik(nMech, CDbl(rsTemp("RATE")), Format(text(3), "DD/MM/YYYY"), Format(text(4), "HH:NN")) Then
81                                    pesan = gBantuan3
82                                    GoTo ULANGI2
83                                Else
84                                    If Not cekDouble(.dGrid, Trim(rsTemp("FS_ITEM")), Brs, 1) Then
85                                        .dGrid.TextMatrix(Brs.Caption, 1) = rsTemp("FS_ITEM")
86                                        .dGrid.TextMatrix(Brs.Caption, 2) = CDbl(rsTemp("RATE"))
87                                        .dGrid.TextMatrix(Brs.Caption, 3) = False
88                                        .dGrid.TextMatrix(Brs.Caption, 4) = "TWC"
89                                        myChassis = Trim(text(0))
90                                        .dGrid.TextMatrix(Brs.Caption, 5) = hitungLabour(CDbl(rsTemp("RATE")), .text(11), True)
91                                        .dGrid.TextMatrix(Brs.Caption, 6) = nMech
92                                        .text(56) = Format(CDbl(.text(56)) + hitungLabour(CDbl(rsTemp("RATE")), .text(11), True), "0#,#")
93                                        .text(61) = Format(CDbl(.text(56)) + CDbl(.text(57)) + CDbl(.text(58)) + CDbl(.text(59)) + CDbl(.text(60)) + CDbl(.text(77)), "#0,#")
94                                        ntRate = ntRate + CDbl(rsTemp("RATE"))
95                                        .text(50) = nMech
96                                        Brs = Brs + 1
97                                    End If
98                                End If
99                            End If
100                       ElseIf Trim(rsTemp("FLAG")) = "P" Then      ' PART
101                           hrg = 0
102                           oliPart = False
103                           Call .isiPart(rsTemp("ITEMPART"), rsTemp("NAMAPART"), rsTemp("SATPART"), rsTemp("QTY"), oliPart, "", 0)
104                       End If
105                       .text(21) = rsTemp("ID_WARRANTY")
106                       .text(129) = NulSID(rsTemp!warranty)
                          '.Text(22) = rsTemp("WARRANTY")
107                       .text(23) = rsTemp("WARRANTY")
108                       .text(96) = CDbl(ntRate)
109                   End With
110               End If
111           rsTemp.MoveNext
112           Loop
113       End If
End Sub

Private Sub Form_Load()

    TData1.EvenRowStyle.BackColor = &HC0FFC0
    TData1.OddRowStyle.BackColor = &HFFC0C0
    TData1.AlternatingRowStyle = True
    
    For i = 0 To 2
        text(i).Enabled = False
    Next
    
    If gTag = "BOOKING" Then
        Call BeriItem(frmBooking.text(41), frmBooking.text(10), 8)
    ElseIf gTag = "ORDER" Then
        Call BeriItem(frmOrder.text(15), frmOrder.text(10), 8)
    End If
End Sub

Private Sub createTbl()
1         On Error GoTo vb_error
          
2         If rsRec.State Then rsRec.Close
3         Set rsRec = Nothing
4         With rsRec
5             .Fields.Append "ID", adVarChar, 10
6             .Fields.Append "WARRANTY", adVarChar, 25
7             .Fields.Append "KETERANGAN", adVarChar, 30
8             .Fields.Append "TGLMULAI", adDate
9             .Fields.Append "TGLSELESAI", adDate
10            .Fields.Append "PROSES", adBoolean
11            .Open
12        End With
13        Set TData1.DataSource = rsRec
14        Exit Sub
          
vb_error:
15       MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Sub

Private Sub TData1_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 5 Then
        Dim k As String
        k = rsRec("ID")
        
        If Not rsRec.EOF Then
            rsRec.MoveFirst
            Do While Not rsRec.EOF
                rsRec!Proses = False
                rsRec.MoveNext
            Loop
            rsRec.filter = "id=" & k
            If Not rsRec.EOF Then
                rsRec!Proses = True
            End If
            rsRec.filter = ""
        End If
    End If
End Sub

