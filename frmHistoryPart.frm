VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmHistoryPart 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Part History"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   5580
      Left            =   450
      TabIndex        =   4
      Top             =   345
      Width           =   11115
      Begin TrueDBGrid70.TDBGrid TData1 
         Height          =   4605
         Left            =   30
         TabIndex        =   3
         Top             =   945
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8123
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tanggal"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No Invoice"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No. WO"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Total Part"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Kode Part"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Penggantian Part"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Qty"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Harga"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1720"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1879"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1799"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=1799"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=1720"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1693"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2566"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2487"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=6244"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=6165"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=661"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=582"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=1852"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=1773"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
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
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H0&,.fgcolor=&HFFFFFF&"
         _StyleDefs(9)   =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=25,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=68,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=27,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=59,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=67,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=69,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=70,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=54,.parent=25"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=26"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=27"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=59"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=20,.parent=25"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=26"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=27"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=59"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=16,.parent=25"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=13,.parent=26"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=14,.parent=27"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=15,.parent=59"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=66,.parent=25"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=26"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=27"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=59"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=74,.parent=25"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=26"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=27"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=59"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=78,.parent=25"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=75,.parent=26"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=76,.parent=27"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=77,.parent=59"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=24,.parent=25"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=21,.parent=26"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=22,.parent=27"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=23,.parent=59"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=32,.parent=25"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=26"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=27"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=59"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   2055
         TabIndex        =   2
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
         Index           =   1
         Left            =   2055
         TabIndex        =   1
         Top             =   240
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "No Polisi"
         Height          =   195
         Left            =   1380
         TabIndex        =   6
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Customer"
         Height          =   195
         Left            =   1335
         TabIndex        =   5
         Top             =   585
         Width           =   660
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   6630
      Left            =   0
      Top             =   -225
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   11695
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Part History"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   30
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9195
      TabIndex        =   0
      Top             =   6000
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmHistoryPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim History As New XArrayDB
Const Kolom = 7
Dim Baris As Integer

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub BeriItem()
    Dim OldInv As String
    
    Baris = 0
    sSQL = "select count(B.NoInv) as JML" _
         & " from history as A inner join t_part as B on A.NO_WO = B.NO_WO " _
         & " where NO_POLISI= '" & gNoPolisi & "' and B.NOINV <> ''"
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        Baris = oRs("JML")
    End If
    oRs.Close
    
    History.Clear
    History.ReDim 0, Baris, 0, Kolom
    i = 0
    
    OldInv = "LAMA"
    sSQL = "select [Date], A.NoInv, A.no_wo, PART, FRONT, KODEITEM, Keterangan, QTY, HARGA, OLI" _
         & " from history as A inner join t_part as B on A.NO_WO = B.NO_WO " _
         & " where NO_POLISI= '" & gNoPolisi & "' and B.NOINV <> '' order by [DATE] DESC, KODEITEM"
   oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        Do While Not oRs.EOF
            If OldInv <> Trim(oRs("NOINV")) Then
                History(i, 0) = Trim(oRs("DATE"))
                History(i, 1) = Trim(oRs("NOINV"))
                History(i, 2) = Trim(oRs("NO_WO"))
                History(i, 3) = Format(Trim(oRs("PART")), "#,##0.00")
                History(i, 4) = Trim(oRs("KODEITEM"))
                History(i, 5) = Trim(oRs("Keterangan"))
                History(i, 6) = Trim(oRs("QTY"))
                History(i, 7) = Format(Trim(oRs("HARGA")), "#,##0.00")
            Else
                History(i, 0) = ""
                History(i, 1) = ""
                History(i, 2) = ""
                History(i, 3) = ""
                History(i, 4) = Trim(oRs("KODEITEM"))
                History(i, 5) = Trim(oRs("Keterangan"))
                History(i, 6) = Trim(oRs("QTY"))
                History(i, 7) = Format(Trim(oRs("HARGA")), "#,##0.00")
            End If
            
            OldInv = Trim(oRs("NOINV"))
            oRs.MoveNext
            i = i + 1
        Loop
    End If
    oRs.Close
    TData1.Array = History
    TData1.ReBind
End Sub

Private Sub cmdExit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call hapusForm
    TData1.EvenRowStyle.BackColor = &HC0FFC0   ' &H80FFFF
    TData1.OddRowStyle.BackColor = &HFFC0C0  ' &HC0C000     ' &HC0FFFF
    TData1.AlternatingRowStyle = True
    Call BeriItem
    Text(1) = gNoPolisi
    Text(0) = gCust
End Sub

Private Sub TData1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub hapusForm()
    Dim i As Integer
    
    For i = 0 To 1
        Text(i) = ""
    Next
End Sub
