VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCariSPK 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc DbConn 
      Height          =   375
      Left            =   2400
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
   Begin MSMask.MaskEdBox txtcari 
      Height          =   315
      Left            =   405
      TabIndex        =   0
      Top             =   420
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   556
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
      Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
      PromptChar      =   ">"
   End
   Begin VB.ComboBox IndexBy 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4065
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   2190
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   405
      Top             =   60
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxSysButton BtnTutup 
      Height          =   285
      Left            =   5955
      Top             =   60
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      ButtonStyle     =   0
      ButtonFlags     =   16384
   End
   Begin TrueOleDBGrid70.TDBGrid dGrid 
      Height          =   5775
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   10186
      _LayoutType     =   0
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
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
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
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   6615
      Left            =   0
      Picture         =   "frmCariSPK.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmCariSPK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strConn As String
Dim strConn1 As String
Dim rsGrid As New ADODB.Recordset

Private Sub ChangeForm(ByVal Lebar As Integer)
    Me.Width = Lebar
    dGrid.Width = Lebar - 550
    txtcari.Width = Lebar - 2775
    asxBars1.Width = Lebar - 900
    IndexBy.Left = Lebar - 2325
    btnTutup.Left = Lebar - 450
End Sub

Private Sub btnTutup_Click()
    Unload Me
End Sub

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If dGrid.row = 0 Then
                txtcari.SetFocus
            End If
        Case vbKeyReturn
            If Not dGrid.row = -1 Then
                gMainKeyCari = dGrid.Columns(0)
            Else
                gMainKeyCari = ""
            End If
            Unload Me
        Case vbKeyEscape
            gMainKeyCari = ""
            Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If LTrim(RTrim(gMainKeyCari)) <> "" Then
        IndexBy.ListIndex = 0
        txtcari = gMainKeyCari
        txtcari.SelStart = Len(txtcari)
        gMainKeyCari = ""
        Call txtCari_Change
    Else
        IndexBy.ListIndex = 0
    End If
    txtcari.SetFocus
End Sub

Private Sub HeadColum(xcolum As Integer, xLebar As Double, xField As String, xCaption As String, xIndex As Boolean)
    dGrid.Columns(xcolum).Width = (dGrid.Width - 500) * xLebar
    dGrid.Columns(xcolum).DataField = xField
    If Left(xField, 2) = "FD" Then
        dGrid.Columns(xcolum).NumberFormat = "DD/MM/YYYY"
    End If
    dGrid.Columns(xcolum).Caption = xCaption
End Sub

Private Sub panggilHead()
    Select Case UCase(gSearchOption)
        Case "WARNA"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.45, "NAMAIND", "Bahasa Indonesia", True)
            Call HeadColum(1, 0.5, "NAMAENGLISH", "Bahasa Inggris", True)
        Case "SALES"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "KodeSales", "Kode Sales", True)
            Call HeadColum(1, 0.6, "NamaSales", "Nama Sales", True)
        Case "SABP"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "KodeSales", "NIK SA", True)
            Call HeadColum(1, 0.6, "NamaSales", "Nama SA BP", True)
    End Select
End Sub

Private Sub panggilCombo()
    IndexBy.Clear
    Select Case UCase(gSearchOption)
        Case "WARNA"
            IndexBy.AddItem ("Bahasa Indonesia")
            IndexBy.AddItem ("Bahasa Inggris")
        Case "SALES"
            IndexBy.AddItem ("Kode Sales")
            IndexBy.AddItem ("Nama Sales")
        Case "SABP"
            IndexBy.AddItem ("NIK SA")
            IndexBy.AddItem ("Nama SA BP")
    End Select
    IndexBy.ListIndex = 0
End Sub

Private Sub Form_Load()
'Ubah
    IndexBy_Click
    Call panggilCombo
    Call CenterForm(Me, True)
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height
End Sub

Private Sub IndexBy_Click()
        Select Case UCase(gSearchOption)
            Case "WARNA"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT DISTINCT NAMAIND, NAMAENGLISH FROM INKATEGORYMSTCLR A WITH(NOLOCK) INNER JOIN INWARNAEXT B WITH(NOLOCK) ON A.KODEWARNA = B.KODEWARNA INNER JOIN INKATEGORYMST C WITH(NOLOCK) ON A.KODE_KATEGORY = C.KODE_KATEGORY INNER JOIN INKATEGORY D WITH(NOLOCK) ON C.KODE_KATEGORY = D.KODE_KATEGORY WHERE KODEKATEGORY = '" & gBantuan & "'"
                        strConn1 = " ORDER BY NAMAIND"
                    Case 1
                        strConn = "SELECT DISTINCT NAMAIND, NAMAENGLISH FROM INKATEGORYMSTCLR A WITH(NOLOCK) INNER JOIN INWARNAEXT B WITH(NOLOCK) ON A.KODEWARNA = B.KODEWARNA INNER JOIN INKATEGORYMST C WITH(NOLOCK) ON A.KODE_KATEGORY = C.KODE_KATEGORY INNER JOIN INKATEGORY D WITH(NOLOCK) ON C.KODE_KATEGORY = D.KODE_KATEGORY WHERE KODEKATEGORY = '" & gBantuan & "'"
                        strConn1 = " ORDER BY NAMAENGLISH"
                End Select
            Case "SALES"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select KodeSales,NamaSales from ArSales where TdkAktif = 0"
                        strConn1 = " ORDER BY KodeSales"
                    Case 1
                        strConn = "select KodeSales,NamaSales from ArSales where TdkAktif = 0"
                        strConn1 = " ORDER BY NamaSales"
                End Select
            Case "SABP"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select nik as KodeSales,[User_name] as NamaSales from [replbp99]..[user] where nik <> '' and dealer = '" & gBantuan & "'"
                        strConn1 = " ORDER BY KodeSales"
                    Case 1
                        strConn = "select nik as KodeSales,[User_name] as NamaSales from [replbp99]..[user] where nik <> '' and dealer = '" & gBantuan & "'"
                        strConn1 = " ORDER BY NamaSales"
                End Select
        End Select
        refresh_Grid (strConn & " " & strConn1)
        If Me.Visible Then
            txtcari.SetFocus
        End If
End Sub

Private Sub IndexBy_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = 38 Or KeyCode = 40 Then
        KeyCode = 0
    End If
End Sub

Private Sub txtCari_Change()
    If dGrid.Columns(IndexBy.ListIndex).DataField <> "" Then
        txtcari.SelStart = Len(txtcari)
        If InStr(UCase(strConn), "WHERE") > 0 Then
            Select Case dGrid.Columns(IndexBy.ListIndex).DataField
                Case "MECH"
                    refresh_Grid (strConn & " AND A.[" & dGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & strConn1)
                Case Else
                    refresh_Grid (strConn & " AND [" & dGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & strConn1)
            End Select
        Else
            refresh_Grid (strConn & " WHERE [" & dGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & strConn1)
        End If
    End If
End Sub

Private Sub txtCari_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            dGrid.SetFocus
        Case vbKeyDown
            dGrid.SetFocus
        Case vbKeyReturn
            If Not dGrid.row = -1 Then
                gMainKeyCari = dGrid.Columns(0)
            Else
                gMainKeyCari = ""
            End If
            Unload Me
        Case vbKeyEscape
            gMainKeyCari = ""
            Unload Me
    End Select
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub refresh_Grid(StrCon As String)
    Set rsGrid = objO.FetchDataSPK(StrCon)
    If Not rsGrid.EOF Then
        dGrid.DataSource = rsGrid
        Call panggilHead
    End If
End Sub

