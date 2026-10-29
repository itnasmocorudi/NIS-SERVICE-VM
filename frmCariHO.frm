VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCariHO 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox IndexBy 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4065
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   2190
   End
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
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   6690
      Left            =   0
      Top             =   0
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   11800
      StartColor      =   8454143
      EndColor        =   16777215
      FormatString    =   "    Pencarian"
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
      Left            =   405
      Top             =   60
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   503
      BackColor       =   16777215
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
End
Attribute VB_Name = "frmCariHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      Dim strConn As String
Dim strConn1 As String
Dim rsGrid As New ADODB.Recordset
'Dim SQLCONBengkel3 As String

Private Sub ChangeForm(ByVal Lebar As Integer)
    Me.Width = Lebar
    dGrid.Width = Lebar - 550
    txtcari.Width = Lebar - 2775
    asxBars1.Width = Lebar - 900
    IndexBy.Left = Lebar - 2325
    BtnTutup.Left = Lebar - 450
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
                If gSearchOption = "KRITERIA" Then
                    frmHOEvent.passingTipe = Trim(dGrid.Columns(1))
                    frmHOEvent.passingFormula = Trim(dGrid.Columns(2))
                ElseIf gSearchOption = "SUPPLIERSUBLET" Then
                    frmHOSublet.Text(2).Text = NulSID(dGrid.Columns(1))
                ElseIf gSearchOption = "HOSUBLET" Then
                    If gBantuan = "SubletPacket" Then
                        frmHOSubletPacket.Text(2).Text = NulSID(dGrid.Columns(1))
                    ElseIf gBantuan = "InsentifSublet" Then
                    
                    Else
                        frmHOPromoSublet.Text(2).Text = NulSID(dGrid.Columns(1))
                        frmHOPromoSublet.hargaBeli.Enabled = True
                        frmHOPromoSublet.hargaBeli.value = NulSID(dGrid.Columns(3))
                        frmHOPromoSublet.hargaBeli.Enabled = False
                    End If
                End If
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
        Case "SALES"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "KodeSales", "Kode Sales", True)
            Call HeadColum(1, 0.6, "NamaSales", "Nama Sales", True)
        Case "MDL"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.4, "MDL", "MDL", True)
            Call HeadColum(1, 0.7, "DESC", "Keterangan", True)
            Call HeadColum(2, 0.1, "MODEL", "Model", True)
            Call HeadColum(3, 0.33, "FS_GRP_MDL", "Group Model", True)
        Case "GRPMDL"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.4, "MDL", "MDL", True)
        Case "JENIS_KERJA"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.2, "JOB", "Kode", True)
            Call HeadColum(1, 0.78, "REMARK", "Keterangan", True)
        Case "GROUP_MODEL"
            Call ChangeForm(4000)
            Call HeadColum(0, 0.98, "FS_GRP_MDL", "Group Model", True)
        Case "JOBLIST"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.78, "FS_ITEM", "Item", True)
        Case "JOB_ORDER"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.7, "FS_ITEM", "Job Order", True)
            Call HeadColum(1, 0.9, "KETERANGAN", "Keterangan", True)
        Case "JOB_ORDER_FOP"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.7, "FS_ITEM", "Job Order", True)
            Call HeadColum(1, 0.9, "KETERANGAN", "Keterangan", True)
        Case "KELAS_MOBIL"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "CLASSCODE", "Tipe", True)
            Call HeadColum(1, 0.5, "CLASSNAME", "Keterangan", True)
        Case "MOU"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.1, "NO_MOU", "No MOU", True)
            Call HeadColum(1, 0.2, "PERUSAHAAN", "Nama Perusahaan", True)
            Call HeadColum(2, 0.2, "PIC", "Contact Person", True)
            Call HeadColum(3, 0.3, "ALAMAT", "Alamat", True)
            Call HeadColum(4, 0.1, "KOTA", "Kota", True)
        Case "PT_MOU"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.3, "NPWP", "NPWP", True)
            Call HeadColum(1, 0.3, "PERUSAHAAN", "Nama Perusahaan", True)
            Call HeadColum(2, 0.2, "PIC", "Contact Person", True)
            Call HeadColum(3, 0.35, "ALAMAT", "Alamat", True)
            Call HeadColum(4, 0.2, "KOTA", "Kota", True)
        Case "PROGRAM"
            Call ChangeForm(12000)
            Call HeadColum(0, 0.1, "ID_PROGRAM", "Kode", True)
            Call HeadColum(1, 0.3, "PROGRAM", "Nama Program", True)
            Call HeadColum(2, 0.6, "KETERANGAN", "Keterangan", True)
            Call HeadColum(3, 0.1, "TGLMULAI", "Tanggal Mulai", True)
            Call HeadColum(4, 0.1, "TGLSELESAI", "Tanggal Selesai", True)
        Case "HOPRG"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "PROGRAM", "Nama Program", True)
            Call HeadColum(1, 0.1, "ID_PROGRAM", "Kode", True)
            Call HeadColum(2, 0.2, "KETERANGAN", "Keterangan", True)
            Call HeadColum(3, 0.2, "TGLMULAI", "Tanggal Mulai", False)
            Call HeadColum(4, 0.2, "TGLSELESAI", "Tanggal Selesai", False)
        Case "WARRANTY"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.1, "ID_WARRANTY", "Kode", True)
            Call HeadColum(1, 0.2, "WARRANTY", "Nama Warranty", True)
            Call HeadColum(2, 0.2, "KETERANGAN", "Keterangan", True)
            Call HeadColum(3, 0.2, "TGLMULAI", "Tanggal Mulai", True)
            Call HeadColum(4, 0.2, "TGLSELESAI", "Tanggal Selesai", True)
        Case "HOWRT"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "WARRANTY", "Nama Warranty", True)
            Call HeadColum(1, 0.1, "ID_WARRANTY", "Kode", True)
            Call HeadColum(2, 0.2, "KETERANGAN", "Keterangan", True)
            Call HeadColum(3, 0.2, "TGLMULAI", "Tanggal Mulai", False)
            Call HeadColum(4, 0.2, "TGLSELESAI", "Tanggal Selesai", False)
        Case "DEALER"
            Call ChangeForm(7000)
            Call HeadColum(0, 0.34, "FS_KD_DEALER", "Kode", True)
            Call HeadColum(1, 0.64, "FS_NM_DEALER", "Nama Dealer", True)
        Case "CITY"
            Call ChangeForm(4000)
            Call HeadColum(0, 0.6, "KOTA", "Kota", True)
        Case "PARTHO"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.3, "ITEM_PART", "Kode Part", True)
            Call HeadColum(1, 0.7, "NAMA_PART", "Nama Part", True)
        Case "COUNTRY"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "KODE", "Kode Negara", True)
            Call HeadColum(1, 0.6, "KETERANGAN", "Negara", True)
        Case "EVENT"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "NOEVENT", "Kode", True)
            Call HeadColum(1, 0.6, "NAMAEVENT", "Event", True)
            Call HeadColum(2, 0.3, "BERLAKU1", "Berlaku Mulai", True)
            Call HeadColum(3, 0.3, "BERLAKU2", "Berlaku Selesai", True)
        Case "KRITERIA"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "KRITERIA", "Kriteria", True)
            Call HeadColum(1, 0.6, "TIPE", "Tipe", True)
            Call HeadColum(2, 1.6, "RUMUS", "Rumus", True)
        Case "HOBI"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.97, "FS_NM_HOBI", "Nama Hobi", True)
        Case "SID"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.1, "IDCUST", "ID CUST", True)
            Call HeadColum(1, 0.3, "NAMA", "Nama Lengkap", True)
            Call HeadColum(2, 0.4, "ALAMAT", "Alamat", True)
            Call HeadColum(3, 0.1, "KOTA", "Kota", True)
        Case "SUPPLIERSUBLET"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.2, "KDSUPP", "KD SUPPLIER", True)
            Call HeadColum(1, 0.3, "SUPPLIER", "NAMA SUPPLIER", True)
            Call HeadColum(2, 0.4, "ALAMAT", "ALAMAT", True)
            Call HeadColum(3, 0.1, "KOTA", "KOTA", True)
        Case "HOSUBLET"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.2, "KDSUBLET", "KD SUBLET", True)
            Call HeadColum(1, 0.4, "NAMASUBLET", "NAMA SUBLET", True)
            Call HeadColum(2, 0.2, "KDSUPP", "KD SUPPLIER", True)
            Call HeadColum(3, 0.2, "HARGABELI", "HARGA BELI", True)
        Case "DATAKARYAWAN"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.2, "NIK", "NIK", True)
            Call HeadColum(1, 0.6, "NAMA", "NAMA KARYAWAN", True)
            Call HeadColum(2, 0.2, "POSISI", "POSISI", True)
        Case "DATAPOSISI"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.9, "POSISI", "POSISI", True)
        Case "DATATIPEMP"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.3, "KdTipe", "KODE", True)
            Call HeadColum(1, 0.7, "Tipe", "POSISI", True)
        Case "DATAMSTKPI"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.3, "KdKPI", "KODE", True)
            Call HeadColum(1, 0.7, "NamaKPI", "NAMAKPI", True)
        Case "PARTDEPO"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.3, "kodeitem", "Kode Part", True)
            Call HeadColum(1, 0.7, "namabarang", "Nama Part", True)
    End Select
End Sub

Private Sub panggilCombo()
    IndexBy.Clear
    Select Case UCase(gSearchOption)
        Case "SALES"
            IndexBy.AddItem ("Kode Sales")
            IndexBy.AddItem ("Nama Sales")
        Case "MDL"
            IndexBy.AddItem ("MDL")
            IndexBy.AddItem ("Keterangan")
            IndexBy.AddItem ("Model")
            IndexBy.AddItem ("Group Model")
        Case "GRPMDL"
            IndexBy.AddItem ("MDL")
        Case "JENIS_KERJA"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Keterangan")
        Case "GROUP_MODEL"
            IndexBy.AddItem ("Group Model")
        Case "JOBLIST"
            IndexBy.AddItem ("Item")
        Case "JOB_ORDER"
            IndexBy.AddItem ("Job Order")
            IndexBy.AddItem ("Keterangan")
        Case "JOB_ORDER_FOP"
            IndexBy.AddItem ("Job Order")
            IndexBy.AddItem ("Keterangan")
        Case "KELAS_MOBIL"
            IndexBy.AddItem ("Tipe")
            IndexBy.AddItem ("Keterangan")
        Case "MOU"
            IndexBy.AddItem ("No MOU")
            IndexBy.AddItem ("Nama Perusahaan")
            IndexBy.AddItem ("Contact Person")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
        Case "PT_MOU"
            IndexBy.AddItem ("NPWP")
            IndexBy.AddItem ("Nama Perusahaan")
            IndexBy.AddItem ("Contact Person")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
        Case "PROGRAM"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Nama Program")
            IndexBy.AddItem ("Keterangan")
            IndexBy.AddItem ("Tanggal Mulai")
            IndexBy.AddItem ("Tanggal Selesai")
        Case "HOPRG"
            IndexBy.AddItem ("Nama Program")
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Keterangan")
            IndexBy.AddItem ("Tanggal Mulai")
            IndexBy.AddItem ("Tanggal Selesai")
        Case "WARRANTY"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Nama Warranty")
            IndexBy.AddItem ("Keterangan")
            IndexBy.AddItem ("Tanggal Mulai")
            IndexBy.AddItem ("Tanggal Selesai")
        Case "HOWRT"
            IndexBy.AddItem ("Nama Warranty")
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Keterangan")
            IndexBy.AddItem ("Tanggal Mulai")
            IndexBy.AddItem ("Tanggal Selesai")
        Case "DEALER"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Nama Dealer")
        Case "CITY"
            IndexBy.AddItem ("Kota")
        Case "PARTHO"
            IndexBy.AddItem ("Kode Item")
            IndexBy.AddItem ("Nama Item")
        Case "COUNTRY"
            IndexBy.AddItem ("Kode Negara")
            IndexBy.AddItem ("Negara")
        Case "EVENT"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Event")
            IndexBy.AddItem ("Berlaku Mulai")
            IndexBy.AddItem ("Berlaku Selesai")
        Case "KRITERIA"
            IndexBy.AddItem ("Kriteria")
            IndexBy.AddItem ("Tipe")
            IndexBy.AddItem ("Rumus")
        Case "HOBI"
            IndexBy.AddItem ("Nama Hobi")
        Case "SID"
            IndexBy.AddItem ("ID CUST")
            IndexBy.AddItem ("Nama Lengkap")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
        Case "SUPPLIERSUBLET"
            IndexBy.AddItem ("Kd Supplier")
            IndexBy.AddItem ("Nama Supplier")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
        Case "HOSUBLET"
            IndexBy.AddItem ("Kd Sublet")
            IndexBy.AddItem ("Nama Sublet")
            IndexBy.AddItem ("Kd Supplier")
            IndexBy.AddItem ("Harga Beli")
        Case "DATAKARYAWAN"
            IndexBy.AddItem ("Nik")
            IndexBy.AddItem ("Nama Karyawan")
            IndexBy.AddItem ("Posisi")
        Case "DATAPOSISI"
            IndexBy.AddItem ("Posisi")
        Case "DATATIPEMP"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Posisi")
        Case "DATAMSTKPI"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Nama KPI")
        Case "PARTDEPO"
            IndexBy.AddItem ("Kode Item")
            IndexBy.AddItem ("Nama Item")
            
    End Select
    IndexBy.ListIndex = 0
End Sub

Private Sub Form_Load()
    On Error GoTo Salah
        
    IndexBy_Click
'    call panggilHead
    Call panggilCombo
    Call CenterForm(Me, True)
    Exit Sub
    
Salah:
    MsgBox "Connect ke AAS Gagal benahi connection melalui profil!"
End Sub

Private Sub IndexBy_Click()
Dim conOk As Boolean

    On Error GoTo Salahlagi
    conOk = False
    If oConnHO.State = adStateClosed Then
        conOk = Get_ConnectionHO(HODb, HOUser, HOPass, HOKomp)
    Else
        conOk = True
    End If
    If conOk Then
        Select Case UCase(gSearchOption)
            Case "SALES"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select KodeSales,NamaSales from Stars99..ArSales where TdkAktif = 0 and KodeCabang = '" & Log.kodeCabang & "'"
                        strConn1 = " ORDER BY KodeSales"
                    Case 1
                        strConn = "select KodeSales,NamaSales from Stars99..ArSales where TdkAktif = 0 and KodeCabang = '" & Log.kodeCabang & "'"
                        strConn1 = " ORDER BY NamaSales"
                End Select
            Case "MDL"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select MDL, [DESC], MODEL, FS_GRP_MDL from model WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " order by MDL"
                    Case 1
                        strConn = "select MDL, [DESC], MODEL, FS_GRP_MDL from model WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " order by [DESC]"
                    Case 2
                        strConn = "select MDL, [DESC], MODEL, FS_GRP_MDL from model WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " order by MODEL"
                    Case 3
                        strConn = "select MDL, [DESC], MODEL, FS_GRP_MDL from model WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " order by FS_GRP_MDL"
                End Select
            Case "GRPMDL"
                strConn = "select DISTINCT LEFT(MDL,CHARINDEX('-',MDL)-1) as MDL from model WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                strConn1 = " order by LEFT(MDL,CHARINDEX('-',MDL)-1)"
            Case "JENIS_KERJA"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select JOB,REMARK from job WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " order by JOB"
                    Case 1
                        strConn = "select JOB,REMARK from job WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " order by REMARK"
                End Select
            Case "GROUP_MODEL"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select distinct FS_GRP_MDL from MODEL WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " order by FS_GRP_MDL"
                End Select
            Case "JOBLIST"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select distinct FS_ITEM from HO_FLATE WITH(NOLOCK) where MDL = '" & gBantuan & "' AND JOB = '" & gBantuan2 & "' AND FS_LOG_DEL = ''"
                        strConn1 = " order by FS_ITEM"
                    Case 1
                        strConn = "select distinct FS_ITEM from HO_FLATE WITH(NOLOCK) where MDL = '" & gBantuan & "' AND JOB = '" & gBantuan2 & "' AND FS_LOG_DEL = ''"
                        strConn1 = " order by FS_JNS_KERJA"
                End Select
            Case "JOB_ORDER"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT FS_ITEM, KETERANGAN FROM JOBORDER WITH(NOLOCK) WHERE JOB = '" & gBantuan & "' AND FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY FS_ITEM"
                End Select
            Case "JOB_ORDER_FOP"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT FS_ITEM, KETERANGAN FROM ho_fop_job WITH(NOLOCK) WHERE JOB = '" & gBantuan & "' AND FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY FS_ITEM"
                End Select
            Case "KELAS_MOBIL"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT CLASSCODE, CLASSNAME FROM CARCLASS WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY CLASSCODE"
                    Case 1
                        strConn = "SELECT CLASSCODE, CLASSNAME FROM CARCLASS WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY CLASSNAME"
                End Select
            Case "MOU"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY NO_MOU"
                    Case 1
                        strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY PERUSAHAAN"
                    Case 2
                        strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY PIC"
                    Case 3
                        strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY ALAMAT"
                    Case 4
                        strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY KOTA"
                End Select
            Case "PT_MOU"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT DISTINCT NPWP, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY NPWP"
                    Case 1
                        strConn = "SELECT DISTINCT NPWP, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY PERUSAHAAN"
                    Case 2
                        strConn = "SELECT DISTINCT NPWP, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY PIC"
                    Case 3
                        strConn = "SELECT DISTINCT NPWP, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY ALAMAT"
                    Case 4
                        strConn = "SELECT DISTINCT NPWP, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY KOTA"
                End Select
            Case "PROGRAM"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY ID_PROGRAM"
                    Case 1
                        strConn = "SELECT ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY PROGRAM"
                    Case 2
                        strConn = "SELECT ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY KETERANGAN"
                    Case 3
                        strConn = "SELECT ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY TGLMULAI"
                    Case 4
                        strConn = "SELECT ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY TGLSELESAI"
                End Select
            Case "HOPRG"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT PROGRAM, ID_PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY ID_PROGRAM"
                    Case 1
                        strConn = "SELECT PROGRAM, ID_PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY PROGRAM"
                    Case 2
                        strConn = "SELECT PROGRAM, ID_PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY KETERANGAN"
                End Select
            Case "WARRANTY"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY ID_WARRANTY"
                    Case 1
                        strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY WARRANTY"
                    Case 2
                        strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY KETERANGAN"
                    Case 3
                        strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY TGLMULAI"
                    Case 4
                        strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY TGLSELESAI"
                End Select
            Case "HOWRT"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT WARRANTY, ID_WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY ID_WARRANTY"
                    Case 1
                        strConn = "SELECT WARRANTY, ID_WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY WARRANTY"
                    Case 2
                        strConn = "SELECT WARRANTY, ID_WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                        strConn1 = " ORDER BY KETERANGAN"
                End Select
            Case "DEALER"
                If gBantuan = "INTERNAL" Then
                    Select Case IndexBy.ListIndex
                        Case 0, -1
                            strConn = "select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1"
                            strConn1 = " order by FS_KD_DEALER"
                        Case 1
                            strConn = "select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1"
                            strConn1 = " order by FS_NM_DEALER"
                    End Select
                ElseIf gBantuan = "ALL" Then
                    Select Case IndexBy.ListIndex
                        Case 0, -1
                            strConn = "select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                            strConn1 = " order by FS_KD_DEALER"
                        Case 1
                            strConn = "select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = ''"
                            strConn1 = " order by FS_NM_DEALER"
                    End Select
                End If
            Case "CITY"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT KOTA FROM T_CITY WITH(NOLOCK)"
                        strConn1 = " ORDER BY KOTA"
                End Select
            Case "PARTHO"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT DISTINCT ITEM_PART, NAMA_PART FROM HO_PARTPACK WITH(NOLOCK) "
                        strConn1 = " ORDER BY ITEM_PART"
                End Select
            Case "COUNTRY"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT KODE, KETERANGAN FROM T_COUNTRY WITH(NOLOCK) "
                        strConn1 = " ORDER BY KODE"
                    Case 1
                        strConn = "SELECT KODE, KETERANGAN FROM T_COUNTRY WITH(NOLOCK) "
                        strConn1 = " ORDER BY KODE"
                    Case 2
                        strConn = "SELECT KODE, KETERANGAN FROM T_COUNTRY WITH(NOLOCK) "
                        strConn1 = " ORDER BY KETERANGAN"
                End Select
            Case "EVENT"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT NOEVENT, NAMAEVENT, BERLAKU1, BERLAKU2 FROM HO_EVENT WITH(NOLOCK) "
                        strConn1 = " ORDER BY NOEVENT"
                    Case 1
                        strConn = "SELECT NOEVENT, NAMAEVENT, BERLAKU1, BERLAKU2 FROM HO_EVENT WITH(NOLOCK) "
                        strConn1 = " ORDER BY NOEVENT"
                    Case 2
                        strConn = "SELECT NOEVENT, NAMAEVENT, BERLAKU1, BERLAKU2 FROM HO_EVENT WITH(NOLOCK) "
                        strConn1 = " ORDER BY NAMAEVENT"
                    Case 3
                        strConn = "SELECT NOEVENT, NAMAEVENT, BERLAKU1, BERLAKU2 FROM HO_EVENT WITH(NOLOCK) "
                        strConn1 = " ORDER BY BERLAKU1"
                    Case 4
                        strConn = "SELECT NOEVENT, NAMAEVENT, BERLAKU1, BERLAKU2 FROM HO_EVENT WITH(NOLOCK) "
                        strConn1 = " ORDER BY BERLAKU2"
                End Select
            Case "KRITERIA"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT KRITERIA, TIPE, RUMUS FROM KRITERIA_EVENT WITH(NOLOCK) "
                        strConn1 = " ORDER BY KRITERIA"
                    Case 1
                        strConn = "SELECT KRITERIA, TIPE, RUMUS FROM KRITERIA_EVENT WITH(NOLOCK) "
                        strConn1 = " ORDER BY KRITERIA"
                    Case 2
                        strConn = "SELECT KRITERIA, TIPE, RUMUS FROM KRITERIA_EVENT WITH(NOLOCK) "
                        strConn1 = " ORDER BY TIPE"
                End Select
            Case "HOBI"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select FS_NM_HOBI from T_HOBI WITH(NOLOCK) "
                        strConn1 = " order by FS_NM_HOBI"
                End Select
            Case "SID"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select top 100 idcust,NamaCust1 + ' ' + NamaCust2 + ' ' + namacust3 as nama,alamat,kota from [singleidori]..NewCustomer a WITH (nolock) "
                        strConn1 = " order by idcust"
                    Case 1
                        strConn = "select top 100 idcust,NamaCust1 + ' ' + NamaCust2 + ' ' + namacust3 as nama,alamat,kota from [singleidori]..NewCustomer a WITH (nolock) "
                        strConn1 = " order by nama"
                    Case 2
                        strConn = "select top 100 idcust,NamaCust1 + ' ' + NamaCust2 + ' ' + namacust3 as nama,alamat,kota from [singleidori]..NewCustomer a WITH (nolock) "
                        strConn1 = " order by alamat"
                    Case 3
                        strConn = "select top 100 idcust,NamaCust1 + ' ' + NamaCust2 + ' ' + namacust3 as nama,alamat,kota from [singleidori]..NewCustomer A WITH (nolock) "
                        strConn1 = " order by kota"
                End Select
            Case "SUPPLIERSUBLET"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select * From (SELECT KDSUPP,SUPPLIER,ALAMAT1 + ' ' + ALAMAT2 AS ALAMAT, KOTA FROM MSTSUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '') a WHERE 1=1 "
                        strConn1 = " order by KDSUPP"
                    Case 1
                        strConn = "select * From (SELECT KDSUPP,SUPPLIER,ALAMAT1 + ' ' + ALAMAT2 AS ALAMAT, KOTA FROM MSTSUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '') a WHERE 1=1 "
                        strConn1 = " order by SUPPLIER"
                    Case 2
                        strConn = "select * From (SELECT KDSUPP,SUPPLIER,ALAMAT1 + ' ' + ALAMAT2 AS ALAMAT, KOTA FROM MSTSUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '') a WHERE 1=1 "
                        strConn1 = " order by ALAMAT"
                    Case 3
                        strConn = "select * From (SELECT KDSUPP,SUPPLIER,ALAMAT1 + ' ' + ALAMAT2 AS ALAMAT, KOTA FROM MSTSUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '') A WHERE 1=1 "
                        strConn1 = " order by KOTA"
                End Select
            Case "HOSUBLET"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select * From (SELECT KDSUBLET,NAMASUBLET,KDSUPP,HARGABELI FROM MSTSUBLET WITH (NOLOCK) WHERE FS_LOG_DEL = '') a WHERE 1=1 "
                        strConn1 = " order by KDSUBLET"
                    Case 1
                        strConn = "select * From (SELECT KDSUBLET,NAMASUBLET,KDSUPP,HARGABELI FROM MSTSUBLET WITH (NOLOCK) WHERE FS_LOG_DEL = '') a WHERE 1=1 "
                        strConn1 = " order by NAMASUBLET"
                    Case 2
                        strConn = "select * From (SELECT KDSUBLET,NAMASUBLET,KDSUPP,HARGABELI FROM MSTSUBLET WITH (NOLOCK) WHERE FS_LOG_DEL = '') a WHERE 1=1 "
                        strConn1 = " order by KDSUPP"
                    Case 3
                        strConn = "select * From (SELECT KDSUBLET,NAMASUBLET,KDSUPP,HARGABELI FROM MSTSUBLET WITH (NOLOCK) WHERE FS_LOG_DEL = '') A WHERE 1=1 "
                        strConn1 = " order by HARGABELI"
                End Select
            Case "DATAKARYAWAN"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select * From (Select NIK, Nama, Posisi from fkaryawan.dbo.datakaryawan a inner join ProfilSPK.dbo.DealerBD b on a.Dealer = b.kodeHRD where divisi='service' and b.kdNR = '" & gKdCbg & "' and ISNULL([tanggal keluar],'') = '' and Posisi not like '%BP%') a WHERE 1=1 "
                        strConn1 = " order by NIK"
                    Case 1
                        strConn = "select * From (Select NIK, Nama, Posisi from fkaryawan.dbo.datakaryawan a inner join ProfilSPK.dbo.DealerBD b on a.Dealer = b.kodeHRD where divisi='service' and b.kdNR = '" & gKdCbg & "' and ISNULL([tanggal keluar],'') = '' and Posisi not like '%BP%') a WHERE 1=1 "
                        strConn1 = " order by Nama"
                    Case 2
                        strConn = "select * From (Select NIK, Nama, Posisi from fkaryawan.dbo.datakaryawan a inner join ProfilSPK.dbo.DealerBD b on a.Dealer = b.kodeHRD where divisi='service' and b.kdNR = '" & gKdCbg & "' and ISNULL([tanggal keluar],'') = '' and Posisi not like '%BP%') a WHERE 1=1 "
                        strConn1 = " order by Posisi"
                End Select
            Case "DATAPOSISI"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select * From (select distinct posisi from fkaryawan.dbo.DataKaryawan where Divisi='service' or posisi like '%ADH%' or posisi like '%Branch%') a WHERE 1=1 "
                        strConn1 = " order by posisi"
                End Select
            Case "DATATIPEMP"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select * From (select kdtipe,tipe from insentif..GRTipeMP where ISNULL(UserDel,'') = '') a WHERE 1=1 "
                        strConn1 = " order by kdtipe"
                    Case 1
                        strConn = "select * From (select kdtipe,tipe from insentif..GRTipeMP where ISNULL(UserDel,'') = '') a WHERE 1=1 "
                        strConn1 = " order by tipe"
                End Select
            Case "DATAMSTKPI"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "select * From (select KdKPI,NamaKPI from insentif..GRMstKPI where ISNULL(UserDel,'') = '') a WHERE 1=1 "
                        strConn1 = " order by KdKPI"
                    Case 1
                        strConn = "select * From (select KdKPI,NamaKPI from insentif..GRMstKPI where ISNULL(UserDel,'') = '') a WHERE 1=1 "
                        strConn1 = " order by NamaKPI"
                End Select
            Case "PARTDEPO"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT * FROM (select distinct KodeItem,NamaBarang from " & gAlamatDepo & "[depo00-" & Format(Now(), "yyyy") & Format(Now(), "mm") & "].dbo.inmst) a WHERE 1=1 "
                        strConn1 = " ORDER BY KodeItem"
                    Case 1
                        strConn = "SELECT * FROM (select distinct KodeItem,NamaBarang from " & gAlamatDepo & "[depo00-" & Format(Now(), "yyyy") & Format(Now(), "mm") & "].dbo.inmst) a WHERE 1=1 "
                        strConn1 = " ORDER BY NamaBarang"
                End Select
            Case "SUBLETHO"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn = "SELECT * FROM (select KDSUBLET,NAMASUBLET from MSTSUBLET where FS_LOG_DEL = '') a WHERE 1=1 "
                        strConn1 = " ORDER BY KDSUBLET"
                    Case 1
                        strConn = "SELECT * FROM (select KDSUBLET,NAMASUBLET from MSTSUBLET where FS_LOG_DEL = '') a WHERE 1=1 "
                        strConn1 = " ORDER BY NAMASUBLET"
                End Select
            End Select
        refresh_Grid (strConn & " " & strConn1)
        If Me.Visible Then
            txtcari.SetFocus
        End If
    End If
    Exit Sub
    
Salahlagi:
    MsgBox "Error pengambilan data!"
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
        If UCase(gSearchOption) = "SID" And dGrid.Columns(IndexBy.ListIndex).DataField = "NAMA" Then
            If InStr(UCase(strConn), "WHERE") > 0 Then
                refresh_Grid (strConn & " AND NamaCust1 + ' ' + NamaCust2 + ' ' + namacust3 LIKE '%" & Trim(txtcari) & "%' " & strConn1)
            Else
                refresh_Grid (strConn & " WHERE NamaCust1 + ' ' + NamaCust2 + ' ' + namacust3 LIKE '%" & Trim(txtcari) & "%' " & strConn1)
            End If
        Else
            If InStr(UCase(strConn), "WHERE") > 0 Then
                refresh_Grid (strConn & " AND [" & dGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & strConn1)
            Else
                refresh_Grid (strConn & " WHERE [" & dGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & strConn1)
            End If
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

    If rsGrid.State Then rsGrid.Close
    Set rsGrid = objO.FetchDataHO(StrCon)
    If Not rsGrid.EOF Then
        dGrid.DataSource = rsGrid
        Call panggilHead
    End If
End Sub


