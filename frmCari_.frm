VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCari 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   6285
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
      Picture         =   "frmCari.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmCari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strConn As String
Dim strConn1 As String
Dim rsGrid As New ADODB.Recordset

Private Sub ChangeForm(ByVal Lebar As Integer)
    Me.Width = Lebar
    DGrid.Width = Lebar - 550
    txtcari.Width = Lebar - 2775
    asxBars1.Width = Lebar - 900
    IndexBy.Left = Lebar - 2325
    BtnTutup.Left = Lebar - 450
End Sub

Private Sub BtnTutup_Click()
    Unload Me
End Sub

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If DGrid.row = 0 Then
                txtcari.SetFocus
            End If
        Case vbKeyReturn
            If Not DGrid.row = -1 Then
                gMainKeyCari = DGrid.Columns(0)
                If gBantuan = "Monitoring MOU" Then
                    gMainKeyCari = DGrid.Columns(1)
                End If
                If gSearchOption = "ESTPOLISI" Or gSearchOption = "ESTRANGKA" Then
                    frmEstimasi.passingXX = DGrid.Columns(1)
                ElseIf gSearchOption = "ANTRIAN" Then
                    frmOrder.passingType = DGrid.Columns(1)
                    frmOrder.passingIDAntrian = DGrid.Columns(2)
                    frmOrder.passingTipeAntrian = DGrid.Columns(3)
                ElseIf gSearchOption = "NPWP" Then
                    frmOrder.passingType = DGrid.Columns(1)
                ElseIf gSearchOption = "DLRSHIFT" Then
                    If frmTShift.Visible Then
                        frmTShift.Text(0) = Trim(DGrid.Columns(0))
                    End If
                ElseIf gSearchOption = "CUSTOMER3" Then
                    If frmInfoNoPolisi.Visible Then
                        frmInfoNoPolisi.passingXX = Trim(DGrid.Columns(0))
                    End If
                ElseIf gSearchOption = "BOOK" Then
                    If DGrid.Columns(1) <> "" Then
                        frmBooking.Text(1) = DGrid.Columns(1)
                    End If
                ElseIf gSearchOption = "INVULANG" Then
                     ''2016-05-04 || jika ada -a / -b dihilangkan
                     If InStr(gMainKeyCari, "-") > 0 Then
                        gMainKeyCari = Left(gMainKeyCari, InStr(gMainKeyCari, "-") - 1)
                     End If
                End If
            Else
                gMainKeyCari = ""
            End If
            If gSearchOption = "INV2" Then
                Me.Hide
            Else
                Unload Me
            End If
            Exit Sub
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
    DGrid.Columns(xcolum).Width = (DGrid.Width - 500) * xLebar
    DGrid.Columns(xcolum).DataField = xField
    If Left(xField, 2) = "FD" Then
        DGrid.Columns(xcolum).NumberFormat = "DD/MM/YYYY"
    End If
    DGrid.Columns(xcolum).Caption = xCaption
    If xIndex = False Then
        DGrid.Columns(xcolum).Visible = False
    End If
End Sub

Private Sub panggilHead()
    Select Case UCase(gSearchOption)
        Case "OLI"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.2, "FS_KD_OLI", "Kode", True)
            Call HeadColum(1, 0.2, "TIPE", "Jenis", True)
            Call HeadColum(2, 0.2, "GENUINE", "Genuine", False)
        Case "AKUN"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "FS_TIPE", "Tipe", True)
            Call HeadColum(1, 0.58, "FS_TIPE_KET", "Keterangan Transaksi", True)
            Call HeadColum(2, 0.2, "KODEJURNAL", "Kode Jurnal", True)
        Case "HOBI"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.97, "FS_NM_HOBI", "Nama Hobi", True)
        Case "CUSTOMER"
            Call ChangeForm(11900)
            Call HeadColum(0, 0.21, "NO_CHASSIS", "No Chassis", True)
            Call HeadColum(1, 0.1, "NO_POLISI", "No Polisi", True)
            Call HeadColum(2, 0.26, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.27, "CUST_ADRES", "Alamat", True)
            Call HeadColum(4, 0.15, "NO_ENGINE", "No Engine", True)
        Case "CUSTOMER3"
            Call ChangeForm(11900)
            Call HeadColum(0, 0.1, "NO_POLISI", "No Polisi", True)
            Call HeadColum(1, 0.21, "NO_CHASSIS", "No Chassis", True)
            Call HeadColum(2, 0.26, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.27, "CUST_ADRES", "Alamat", True)
            Call HeadColum(4, 0.15, "NO_ENGINE", "No Engine", True)
        Case "BOOK"
            Call ChangeForm(10000)
            Call HeadColum(0, 0.14, "NO_POLISI", "No Polisi", True)
            Call HeadColum(1, 0.15, "NOFAK", "NOFAK", True)
            Call HeadColum(2, 0.15, "DATE_BOOK", "Tanggal Pesan", True)
            Call HeadColum(3, 0.15, "DATE", "Tanggal Booking", True)
            Call HeadColum(4, 0.08, "DELIV", "Status", True)
            Call HeadColum(5, 0.4, "CUSTOMER", "Customer", True)
            Call HeadColum(6, 0.34, "CUST_ADRES", "Alamat", True)
            Call HeadColum(7, 0.13, "KOTA", "Kota", True)
        Case "ORDER"
            Call ChangeForm(11000)
            Call HeadColum(0, 0.09, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.1, "DATE", "Tanggal", True)
            Call HeadColum(2, 0.1, "FRONT", "SA", True)
            Call HeadColum(3, 0.1, "NO_POLISI", "No Polisi", True)
            Call HeadColum(4, 0.1, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(5, 0.4, "CUSTOMER", "Customer", True)
            Call HeadColum(6, 0.5, "CUST_ADRES", "Alamat", True)
            Call HeadColum(7, 0.2, "NO_CHASSIS", "No Chasis", True)
        Case "METERAI"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.68, "FN_BATAS", "Batas Minimal", True)
            Call HeadColum(1, 0.3, "FN_NILAI", "Nilai", True)
        Case "HISTORY"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "NO_POLISI", "No Polisi", True)
            Call HeadColum(1, 0.68, "CUSTOMER", "Customer", True)
        Case "PROGRESS"
            Call ChangeForm(7000)
            Call HeadColum(0, 0.23, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.41, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.23, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.11, "RPR_COND", "Kondisi", True)
        Case "REG"
            Call ChangeForm(7000)
            Call HeadColum(0, 0.23, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.41, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.23, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.11, "RPR_COND", "Kondisi", True)
        Case "DEALER"
            Call ChangeForm(7000)
            Call HeadColum(0, 0.34, "FS_KD_DEALER", "Kode", True)
            Call HeadColum(1, 0.64, "FS_NM_DEALER", "Nama Dealer", True)
        Case "FORMATSMS"
            Call ChangeForm(7000)
            Call HeadColum(0, 0.34, "typeSMS", "Tipe SMS", True)
            Call HeadColum(1, 0.64, "pesan1", "Pesan", True)
        Case "REGISTBATAL"
            Call ChangeForm(7000)
            Call HeadColum(0, 0.23, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.41, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.23, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.11, "RPR_COND", "Kondisi", True)
        Case "BOOKBATAL"
            Call ChangeForm(10000)
            Call HeadColum(0, 0.14, "NO_POLISI", "No Polisi", True)
            Call HeadColum(1, 0.24, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.34, "CUST_ADRES", "Alamat", True)
            Call HeadColum(3, 0.13, "KOTA", "Kota", True)
            Call HeadColum(4, 0.13, "FS_KD_DEALER", "Dealer", True)
        Case "ORDERBATAL"
            Call ChangeForm(11000)
            Call HeadColum(0, 0.09, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.1, "DATE", "Tanggal", True)
            Call HeadColum(2, 0.1, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.1, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(4, 0.4, "CUSTOMER", "Customer", True)
            Call HeadColum(5, 0.5, "CUST_ADRES", "Alamat", True)
            Call HeadColum(6, 0.2, "NO_CHASSIS", "No Chasis", True)
        Case "LIHATORDER"
            Call ChangeForm(10000)
            Call HeadColum(0, 0.14, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.24, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.34, "CUST_ADRES", "Alamat", True)
            Call HeadColum(3, 0.13, "KOTA", "Kota", True)
            Call HeadColum(4, 0.13, "FS_KD_DEALER", "Dealer", True)
        Case "INV"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.15, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.15, "NOINV", "No Invoice", True)
            Call HeadColum(2, 0.15, "TANGGAL", "Tanggal", True)
            Call HeadColum(3, 0.15, "FRONT", "SA", True)
            Call HeadColum(4, 0.15, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(5, 0.38, "CUSTOMER", "Customer", True)
            Call HeadColum(6, 0.15, "NO_POLISI", "No Polisi", True)
        Case "INV2"
            Call ChangeForm(14000)
            Call HeadColum(0, 0.08, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.08, "TANGGAL", "Tanggal", True)
            Call HeadColum(2, 0.09, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.08, "NOINV", "No Invoice", True)
            Call HeadColum(4, 0.05, "FRONT", "SA", True)
            Call HeadColum(5, 0.08, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(6, 0.28, "CUSTOMER", "Customer", True)
            Call HeadColum(7, 0.15, "NO_CHASSIS", "No Rangka", True)
            Call HeadColum(8, 0.04, "RPR_COND", "Closing", False)
        Case "INVCTKULANG"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.15, "NOINV", "No Invoice", True)
            Call HeadColum(1, 0.15, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(2, 0.38, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.15, "NO_POLISI", "No Polisi", True)
            Call HeadColum(4, 0.15, "NO_WO", "No WO", True)
        Case "INVULANG"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.15, "NOINV", "No Invoice", True)
            Call HeadColum(1, 0.15, "NO_WO", "No WO", True)
            Call HeadColum(2, 0.15, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(3, 0.38, "CUSTOMER", "Customer", True)
            Call HeadColum(4, 0.15, "NO_POLISI", "No Polisi", True)
        Case "INVBATAL"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.16, "NOINV", "No Invoice", True)
            Call HeadColum(1, 0.34, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.16, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.16, "NO_WO", "No WO", True)
            Call HeadColum(4, 0.16, "FS_NO_RETUR", "No Retur", True)
        Case "INVEDIT"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "NOINV", "No Invoice", True)
            Call HeadColum(1, 0.38, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.2, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.2, "NO_WO", "No WO", True)
        Case "EDITRATE"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.38, "NoInv", "No Invoice", True)
            Call HeadColum(2, 0.2, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.2, "NO_POLISI", "No Polisi", True)
        Case "INVCATAT"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "NOINV", "No Invoice", True)
            Call HeadColum(1, 0.38, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.2, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.2, "NO_WO", "No WO", True)
        Case "INVBATALCATAT"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "NOINV", "No Invoice", True)
            Call HeadColum(1, 0.38, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.2, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.2, "NO_WO", "No WO", True)
        Case "INVBAYAR"
            Call ChangeForm(7000)
            Call HeadColum(0, 0.25, "NOINV", "No Invoice", True)
            Call HeadColum(1, 0.48, "CUSTOMER", "Customer", True)
            Call HeadColum(2, 0.25, "NO_POLISI", "No Polisi", True)
        Case "MEKANIK"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "MECH", "Mech", True)
            Call HeadColum(1, 0.38, "NAME", "Nama", True)
            Call HeadColum(2, 0.2, "GROUP", "Group", True)
            Call HeadColum(3, 0.2, "LEVMECH", "Level", True)
        Case "GROUP_MEKANIK"
            Call ChangeForm(7500)
            Call HeadColum(0, 0.2, "GROUP", "Group", True)
            Call HeadColum(1, 0.58, "NAME", "Nama", True)
            Call HeadColum(2, 0.2, "MECH", "Mech", True)
        Case "FOREMAN"
            Call ChangeForm(7500)
            Call HeadColum(0, 0.2, "FS_KD_PETUGAS", "Inisial", True)
            Call HeadColum(1, 0.58, "FS_NM_PETUGAS", "Nama", True)
        Case "MDL"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.4, "MDL", "MDL", True)
            Call HeadColum(1, 0.9, "DESC", "Keterangan", True)
            Call HeadColum(2, 0.3, "MODEL", "Model", True)
            Call HeadColum(3, 0.33, "FS_GRP_MDL", "Group Model", True)
        Case "JENIS_KERJA"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.2, "JOB", "Kode", True)
            Call HeadColum(1, 0.78, "REMARK", "Keterangan", True)
        Case "PETUGAS"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.2, "FS_KD_PETUGAS", "Kode", True)
            Call HeadColum(1, 0.78, "FS_NM_PETUGAS", "Nama Petugas", True)
        Case "GROUP_MODEL"
            Call ChangeForm(4000)
            Call HeadColum(0, 0.98, "FS_GRP_MDL", "Group Model", True)
        Case "PARTOLI"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "NOINV", "No Invoice", True)
            Call HeadColum(1, 0.3, "NO_WO", "No WO", True)
            Call HeadColum(2, 0.38, "NO_POLISI", "No Polisi", True)
        Case "JOBLIST2"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.8, "FS_ITEM", "Item", True)
            Call HeadColum(1, 0.15, "RATE", "Rate", True)
            Call HeadColum(2, 0.5, "KETERANGAN", "Keterangan", True)
    
        ' ,;; ==========00000ooooo00000========== TAMBAHAN FUNGSI ==========00000ooooo00000========== ;;,
        ' ,;; ==========00000ooooo00000========== TAMBAHAN FUNGSI ==========00000ooooo00000========== ;;,
    
        Case "CUSTOMER2"
            Call ChangeForm(11900)
            Call HeadColum(0, 0.21, "NO_CHASSIS", "No Chassis", True)
            Call HeadColum(1, 0.1, "NO_POLISI", "No Polisi", True)
            Call HeadColum(2, 0.26, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.27, "CUST_ADRES", "Alamat", True)
            Call HeadColum(4, 0.15, "NO_ENGINE", "No Engine", True)
        Case "JOB_ORDER"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.7, "FS_ITEM", "Job Order", True)
    '            Call HeadColum(1, 0.6, "KETERANGAN", "Keterangan", True)
        Case "KELAS_MOBIL"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "CLASSCODE", "Tipe", True)
            Call HeadColum(1, 0.5, "CLASSNAME", "Keterangan", True)
        Case "AGAMA"
            Call ChangeForm(4000)
            Call HeadColum(0, 0.6, "AGAMA", "Agama", True)
    '        Case "CITY"
    '            Call ChangeForm(4000)
    '            Call HeadColum(0, 0.6, "KOTA", "Kota", True)
        Case "LEVMEKANIK"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.1, "LEVELMECH", "Level", True)
            Call HeadColum(1, 0.6, "KETERANGAN", "Keterangan", True)
        Case "SUBLET"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.6, "NAMABARANG", "Sublet", True)
        Case "POSISI_MOBIL"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.6, "posisiAwal", "Posisi", True)
        Case "TPSVC"
            Call ChangeForm(4000)
            Call HeadColum(0, 0.8, "TPSERVICE", "Tipe", True)
        Case "TP_ORD"
            Call ChangeForm(4000)
            Call HeadColum(0, 0.8, "TPORDER", "Tipe", True)
        Case "PROSES"
            Call ChangeForm(4000)
            Call HeadColum(0, 0.8, "PROSES", "Tipe", True)
        Case "RETURN"
            Call ChangeForm(4000)
            Call HeadColum(0, 0.8, "RETUR", "Tipe", True)
        Case "KONDISI"
            Call ChangeForm(4000)
            Select Case gBantuan
                Case 1
                    Call HeadColum(0, 0.8, "WAKTU", "Waktu", True)
                Case 2
                    Call HeadColum(0, 0.8, "FREKUENSI", "Frekuensi", True)
                Case 3
                    Call HeadColum(0, 0.8, "TEMPAT", "Tempat", True)
                Case 4
                    Call HeadColum(0, 0.8, "WARN", "Warn", True)
                Case 5
                    Call HeadColum(0, 0.8, "MULAI", "Mulai", True)
                Case 6
                    Call HeadColum(0, 0.8, "SELESAI", "Selesai", True)
            End Select
        Case "BOOKING"
            Call ChangeForm(12000)
            Call HeadColum(0, 0.1, "NOFAK", "No Booking", True)
            Call HeadColum(1, 0.1, "DATE", "Tanggal", True)
            Call HeadColum(2, 0.1, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.08, "DELIV", "Status", True)
            Call HeadColum(4, 0.4, "CUSTOMER", "Customer", True)
            Call HeadColum(5, 0.8, "CUST_ADRES", "Alamat", True)
            Call HeadColum(6, 0.2, "KOTA", "Kota", True)
            Call HeadColum(7, 0.3, "FMDL", "Model", True)
            
        Case "BTLBOOK"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "NOFAK", "No Booking", True)
            Call HeadColum(1, 0.2, "NO_POLISI", "No Polisi", True)
            Call HeadColum(2, 0.5, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.8, "CUST_ADRES", "Alamat", True)
            Call HeadColum(4, 0.3, "KOTA", "Kota", True)
            Call HeadColum(5, 0.3, "FMDL", "Model", True)
            Call HeadColum(6, 0.2, "DELIV", "Status", True)

            Case "PROGRAM"
                Call ChangeForm(8000)
                Call HeadColum(0, 0.2, "ID_PROGRAM", "Kode", True)
                Call HeadColum(1, 0.4, "PROGRAM", "Nama Program", True)
                Call HeadColum(2, 0.5, "KETERANGAN", "Keterangan", True)
                Call HeadColum(3, 0.2, "TGLMULAI", "Tanggal Mulai", True)
                Call HeadColum(4, 0.2, "TGLSELESAI", "Tanggal Selesai", True)
            Case "PROGRAM_ALL"
                Call ChangeForm(8000)
                Call HeadColum(0, 0.2, "ID_PROGRAM", "Kode", True)
                Call HeadColum(1, 0.4, "PROGRAM", "Nama Program", True)
                Call HeadColum(2, 0.5, "KETERANGAN", "Keterangan", True)
                Call HeadColum(3, 0.2, "TGLMULAI", "Tanggal Mulai", True)
                Call HeadColum(4, 0.2, "TGLSELESAI", "Tanggal Selesai", True)
            Case "WARRANTY"
                Call ChangeForm(8000)
                Call HeadColum(0, 0.1, "ID_WARRANTY", "Kode", True)
                Call HeadColum(1, 0.2, "WARRANTY", "Nama Warranty", True)
                Call HeadColum(2, 0.2, "KETERANGAN", "Keterangan", True)
                Call HeadColum(3, 0.2, "TGLMULAI", "Tanggal Mulai", True)
                Call HeadColum(4, 0.2, "TGLSELESAI", "Tanggal Selesai", True)
        Case "ASURANSI"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.4, "ASURANSI", "Nama Asuransi", True)
            Call HeadColum(1, 0.2, "PIC", "Contact Person", True)
            Call HeadColum(2, 0.3, "ALAMAT", "Alamat", True)
            Call HeadColum(3, 0.2, "KOTA", "Kota", True)
        Case "ESTPOLISI"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "NO_POLISI", "No. Polisi", True)
            Call HeadColum(1, 0.4, "NOEST", "No. Estimasi", True)
            Call HeadColum(2, 0.4, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.8, "CUST_ADRES", "Alamat", True)
            Call HeadColum(4, 0.1, "FMDL", "Model", True)
            Call HeadColum(5, 0.1, "STATUS", "Status Estimasi", True)
        Case "ESTRANGKA"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "NO_RANGKA", "No. Rangka", True)
            Call HeadColum(1, 0.4, "NOEST", "No. Estimasi", True)
            Call HeadColum(2, 0.4, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.8, "CUST_ADRES", "Alamat", True)
            Call HeadColum(4, 0.1, "FMDL", "Model", True)
            Call HeadColum(5, 0.1, "STATUS", "Status Estimasi", True)
        Case "BTLREG"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.15, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.15, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(2, 0.38, "CUSTOMER", "Customer", True)
            Call HeadColum(3, 0.15, "NO_POLISI", "No Polisi", True)
            Call HeadColum(4, 0.15, "NOINV", "No Invoice", True)
        Case "SA"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.2, "FS_KD_PETUGAS", "Kode", True)
            Call HeadColum(1, 0.4, "FS_NM_PETUGAS", "Petugas", True)
        Case "WARNA"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.2, "NAMAIND", "Bahasa Indonesia", True)
            Call HeadColum(1, 0.4, "NAMAENGLISH", "Bahasa Inggris", True)
        Case "DEALERINV"
            Call ChangeForm(8000)
            Call HeadColum(0, 0.2, "FS_KD_DEALER", "Kode Dealer", True)
            Call HeadColum(1, 0.6, "FS_NM_DEALER", "Nama Dealer", True)
        Case "TIPEPETUGAS"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.4, "TIPE", "Tipe Petugas", True)
        Case "COUNTRY"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "KODE", "Kode Negara", True)
            Call HeadColum(1, 0.6, "KETERANGAN", "Negara", True)
        Case "STATUSAKUN"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "TIPE", "Tipe", True)
            Call HeadColum(1, 0.6, "KETERANGAN", "Keterangan", True)
        Case "VALIDSUBLET"
            Call ChangeForm(11000)
            Call HeadColum(0, 0.09, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.1, "DATE", "Tanggal", True)
            Call HeadColum(2, 0.1, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.1, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(4, 0.4, "CUSTOMER", "Customer", True)
            Call HeadColum(5, 0.5, "CUST_ADRES", "Alamat", True)
            Call HeadColum(6, 0.2, "NO_CHASSIS", "No Chasis", True)
        Case "RETURSUBLET"
            Call ChangeForm(11000)
            Call HeadColum(0, 0.09, "NO_WO", "No WO", True)
            Call HeadColum(1, 0.1, "DATE", "Tanggal", True)
            Call HeadColum(2, 0.1, "NO_POLISI", "No Polisi", True)
            Call HeadColum(3, 0.1, "FS_KD_PLG", "Kode Plg", True)
            Call HeadColum(4, 0.4, "CUSTOMER", "Customer", True)
            Call HeadColum(5, 0.5, "CUST_ADRES", "Alamat", True)
            Call HeadColum(6, 0.2, "NO_CHASSIS", "No Chasis", True)
        Case "STALL"
            Call ChangeForm(5000)
            Call HeadColum(0, 0.1, "NOSTALL", "No Stall", True)
        Case "GRPMDL"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.4, "MDL", "MDL", True)
        Case "ALASANDTG"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.4, "ALASAN", "Alasan Kedatangan", True)
        Case "NPWP"
            Call ChangeForm(15000)
            Call HeadColum(0, 0.2, "NPWP", "NPWP", True)
            Call HeadColum(1, 0.4, "NAMAFP", "Nama Customer", True)
            Call HeadColum(2, 0.5, "ALAMATFP1", "Alamat Customer", True)
            Call HeadColum(3, 0.3, "ALAMATFP2", "Kota Customer", True)
        Case "TPPRINTER"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "TPPRINT", "Tipe Printer", True)
        Case "ANTRIAN"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "QUEUE", "Nomor", True)
            Call HeadColum(1, 0.3, "TIPE", "Tipe", True)
            Call HeadColum(2, 0.3, "ID", "ID", False)
            Call HeadColum(3, 0.3, "NAMA", "NAMA", False)
        Case "TIPEESTMASI"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.5, "KETERANGAN", "Tipe", True)
        Case "DLRSHIFT"
            Call ChangeForm(6000)
            Call HeadColum(0, 0.3, "SHIFT", "Shift", True)
            Call HeadColum(1, 0.3, "JAMMULAI", "Jam Mulai", True)
            Call HeadColum(2, 0.3, "JAMSELESAI", "Jam Selesai", True)
        Case "MOU"
            Call ChangeForm(9000)
            Call HeadColum(0, 0.1, "NO_MOU", "No MOU", True)
            Call HeadColum(1, 0.2, "PERUSAHAAN", "Nama Perusahaan", True)
            Call HeadColum(2, 0.2, "PIC", "Contact Person", True)
            Call HeadColum(3, 0.3, "ALAMAT", "Alamat", True)
            Call HeadColum(4, 0.1, "KOTA", "Kota", True)

    End Select
End Sub

Private Sub panggilCombo()
    IndexBy.Clear
    Select Case UCase(gSearchOption)
        Case "OLI"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Jenis")
        Case "AKUN"
            IndexBy.AddItem ("Tipe")
            IndexBy.AddItem ("Keterangan Transaksi")
            IndexBy.AddItem ("Kode Jurnal")
        Case "HOBI"
            IndexBy.AddItem ("Nama Hobi")
        Case "CUSTOMER"
            IndexBy.AddItem ("No Chassis")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("No Engine")
        Case "CUSTOMER3"
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("No Chassis")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("No Engine")
        Case "BOOK"
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("NOFAK")
            IndexBy.AddItem ("Tanggal Pesan")
            IndexBy.AddItem ("Tanggal Booking")
            IndexBy.AddItem ("Status")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
        Case "ORDER"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Tanggal")
            IndexBy.AddItem ("SA")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("No Chassis")
        Case "METERAI"
            IndexBy.AddItem ("Batas Minimal")
            IndexBy.AddItem ("Nilai")
        Case "HISTORY"
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Customer")
        Case "PROGRESS"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Kondisi")
        Case "REG"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Kondisi")
        Case "DEALER"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Nama Dealer")
        Case "FORMATSMS"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Format Sms")
        Case "REGISTBATAL"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Kondisi")
        Case "BOOKBATAL"
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
            IndexBy.AddItem ("Dealer")
        Case "ORDERBATAL"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Tanggal")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("No Chassis")
        Case "LIHATORDER"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
            IndexBy.AddItem ("Dealer")
        Case "INV"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("Tanggal")
            IndexBy.AddItem ("SA")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
        Case "INV2"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Tanggal")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("SA")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Rangka")
            IndexBy.AddItem ("Closing")
        Case "INVCTKULANG"
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("No WO")
        Case "INVULANG"
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
        Case "INVBATAL"
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("No Retur")
        Case "INVEDIT"
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("No WO")
        Case "EDITRATE"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
        Case "INVCATAT"
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("No WO")
        Case "INVBATALCATAT"
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("No WO")
        Case "INVBAYAR"
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
        Case "MEKANIK"
            IndexBy.AddItem ("Mech")
            IndexBy.AddItem ("Nama")
            IndexBy.AddItem ("Group")
            IndexBy.AddItem ("Level")
        Case "GROUP_MEKANIK"
            IndexBy.AddItem ("Group")
            IndexBy.AddItem ("Nama")
            IndexBy.AddItem ("Mech")
        Case "FOREMAN"
            IndexBy.AddItem ("Inisial")
            IndexBy.AddItem ("Nama")
        Case "MDL"
            IndexBy.AddItem ("MDL")
            IndexBy.AddItem ("Keterangan")
            IndexBy.AddItem ("Model")
            IndexBy.AddItem ("Group Model")
        Case "JENIS_KERJA"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Keterangan")
        Case "PETUGAS"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Nama Petugas")
        Case "GROUP_MODEL"
            IndexBy.AddItem ("Group Model")
        Case "PARTOLI"
            IndexBy.AddItem ("No Invoice")
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("No Polisi")
        Case "JOBLIST2"
            IndexBy.AddItem ("Item")
            IndexBy.AddItem ("Rate")
            IndexBy.AddItem ("Keterangan")

    
        ' ,;; ==========00000ooooo00000========== TAMBAHAN FUNGSI ==========00000ooooo00000========== ;;,
        ' ,;; ==========00000ooooo00000========== TAMBAHAN FUNGSI ==========00000ooooo00000========== ;;,
    
        Case "CUSTOMER2"
            IndexBy.AddItem ("No Chassis")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("No Engine")
        Case "JOB_ORDER"
            IndexBy.AddItem ("Job Order")
        Case "KELAS_MOBIL"
            IndexBy.AddItem ("Tipe")
            IndexBy.AddItem ("Keterangan")
        Case "AGAMA"
            IndexBy.AddItem ("Agama")
        Case "LEVMEKANIK"
            IndexBy.AddItem ("Level")
            IndexBy.AddItem ("Keterangan")
        Case "SUBLET"
            IndexBy.AddItem ("Sublet")
        Case "POSISI_MOBIL"
            IndexBy.AddItem ("Posisi")
        Case "TPSVC"
            IndexBy.AddItem ("Tipe")
        Case "TP_ORD"
            IndexBy.AddItem ("Tipe")
        Case "PROSES"
            IndexBy.AddItem ("Tipe")
        Case "RETURN"
            IndexBy.AddItem ("Tipe")
        Case "KONDISI"
            Select Case gBantuan
                Case 1
                    IndexBy.AddItem ("Waktu")
                Case 2
                    IndexBy.AddItem ("Frekuensi")
                Case 3
                    IndexBy.AddItem ("Tempat")
                Case 4
                    IndexBy.AddItem ("Warn")
                Case 5
                    IndexBy.AddItem ("Mulai")
                Case 6
                    IndexBy.AddItem ("Selesai")
            End Select
        Case "BOOKING"
            IndexBy.AddItem ("No Booking")
            IndexBy.AddItem ("Tanggal")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Status")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
            IndexBy.AddItem ("Model")
        Case "BTLBOOK"
            IndexBy.AddItem ("No Booking")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
            IndexBy.AddItem ("Model")
            IndexBy.AddItem ("Status")
            Case "PROGRAM"
                IndexBy.AddItem ("Kode")
                IndexBy.AddItem ("Nama Program")
                IndexBy.AddItem ("Keterangan")
                IndexBy.AddItem ("Tanggal Mulai")
                IndexBy.AddItem ("Tanggal Selesai")
            Case "PROGRAM_ALL"
                IndexBy.AddItem ("Kode")
                IndexBy.AddItem ("Nama Program")
                IndexBy.AddItem ("Keterangan")
                IndexBy.AddItem ("Tanggal Mulai")
                IndexBy.AddItem ("Tanggal Selesai")
            Case "WARRANTY"
                IndexBy.AddItem ("Kode")
                IndexBy.AddItem ("Nama Warranty")
                IndexBy.AddItem ("Keterangan")
                IndexBy.AddItem ("Tanggal Mulai")
                IndexBy.AddItem ("Tanggal Selesai")
        Case "ASURANSI"
            IndexBy.AddItem ("Nama Asuransi")
            IndexBy.AddItem ("Contact Person")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
        Case "ESTPOLISI"
            IndexBy.AddItem ("No. Polisi")
            IndexBy.AddItem ("No. Estimasi")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Model")
            IndexBy.AddItem ("Status Estimasi")
        Case "ESTRANGKA"
            IndexBy.AddItem ("No. Rangka")
            IndexBy.AddItem ("No. Estimasi")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Model")
            IndexBy.AddItem ("Status Estimasi")
        Case "BTLREG"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("No Invoice")
        Case "SA"
            IndexBy.AddItem ("Kode")
            IndexBy.AddItem ("Petugas")
        Case "WARNA"
            IndexBy.AddItem ("Bahasa Indonesia")
            IndexBy.AddItem ("Bahasa Inggris")
        Case "DEALERINV"
            IndexBy.AddItem ("Kode Dealer")
            IndexBy.AddItem ("Nama Dealer")
        Case "TIPEPETUGAS"
            IndexBy.AddItem ("Tipe Petugas")
        Case "COUNTRY"
            IndexBy.AddItem ("Kode Negara")
            IndexBy.AddItem ("Negara")
        Case "STATUSAKUN"
            IndexBy.AddItem ("Tipe")
            IndexBy.AddItem ("Keterangan")
        Case "VALIDSUBLET"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Tanggal")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("No Chassis")
        Case "RETURSUBLET"
            IndexBy.AddItem ("No WO")
            IndexBy.AddItem ("Tanggal")
            IndexBy.AddItem ("No Polisi")
            IndexBy.AddItem ("Kode Plg")
            IndexBy.AddItem ("Customer")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("No Chassis")
        Case "STALL"
            IndexBy.AddItem ("No Stall")
        Case "GRPMDL"
            IndexBy.AddItem ("MDL")
        Case "ALASANDTG"
            IndexBy.AddItem ("Alasan Kedatangan")
        Case "NPWP"
            IndexBy.AddItem ("NPWP")
            IndexBy.AddItem ("Nama Customer")
            IndexBy.AddItem ("Alamat Customer")
            IndexBy.AddItem ("Kota Customer")
        Case "TPPRINTER"
            IndexBy.AddItem ("TPPRINT")
        Case "ANTRIAN"
            IndexBy.AddItem ("QUEUE")
            IndexBy.AddItem ("TIPE")
            IndexBy.AddItem ("ID")
            IndexBy.AddItem ("NAMA")
        Case "TIPEESTMASI"
            IndexBy.AddItem ("KETERANGAN")
        Case "DLRSHIFT"
            IndexBy.AddItem ("SHIFT")
            IndexBy.AddItem ("JAMMULAI")
            IndexBy.AddItem ("JAMSELESAI")
        Case "MOU"
            IndexBy.AddItem ("No MOU")
            IndexBy.AddItem ("Nama Perusahaan")
            IndexBy.AddItem ("Contact Person")
            IndexBy.AddItem ("Alamat")
            IndexBy.AddItem ("Kota")
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
        Case "OLI"
            strConn = "select top 50 FS_KD_OLI, CASE TIPE WHEN 'P' THEN 'PART' ELSE 'OLI' END AS TIPE, CASE GENUINE WHEN '1' THEN 'GENUINE' ELSE 'NON GENUINE' END AS GENUINE from T_OLI WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by FS_KD_OLI"
                Case 1
                    strConn1 = " order by TIPE"
            End Select
        Case "AKUN"
            strConn = "select top 50 FS_TIPE, FS_TIPE_KET, KODEJURNAL from T_AKUN WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by FS_TIPE"
                Case 1
                    strConn1 = " order by FS_TIPE_KET"
                Case 2
                    strConn1 = " order by KODEJURNAL"
            End Select
        Case "HOBI"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "select top 50 FS_NM_HOBI from T_HOBI WITH (NOLOCK)"
                    strConn1 = " order by FS_NM_HOBI"
            End Select
        Case "CUSTOMER"
            strConn = "select top 50 NO_CHASSIS, NO_POLISI, CUSTOMER, CUST_ADRES, NO_ENGINE from custdat WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_CHASSIS"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by CUST_ADRES"
                Case 3
                    strConn1 = " order by NO_ENGINE"
                Case 4
                    strConn1 = " order by NO_POLISI"
            End Select
        Case "CUSTOMER3"
            strConn = "select top 50 NO_POLISI, NO_CHASSIS , CUSTOMER, CUST_ADRES, NO_ENGINE from custdat WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_POLISI"
                Case 1
                    strConn1 = " order by NO_CHASSIS"
                Case 2
                    strConn1 = " order by CUSTOMER"
                Case 3
                    strConn1 = " order by CUST_ADRES"
                Case 4
                    strConn1 = " order by NO_ENGINE"
            End Select
        Case "BOOK"
            strConn = "select Top 50 AA.* from (SELECT DISTINCT NOFAK, NO_POLISI, DATE_BOOK, DATE, DELIV, CUSTOMER, CUST_ADRES, KOTA from book WITH (NOLOCK) WHERE (DELIV = 'BOK' OR LEFT(DELIV,3) = 'RES')) AA "
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = "order by AA.NO_POLISI"
                Case 1
                    strConn1 = "order by AA.NOFAK"
                Case 2
                    strConn1 = "order by AA.DATE_BOOK"
                Case 3
                    strConn1 = "order by AA.DATE DESC"
                Case 4
                    strConn1 = "order by AA.DELIV"
                Case 5
                    strConn1 = "order by AA.CUSTOMER"
                Case 6
                    strConn1 = "order by AA.CUST_ADRES"
                Case 7
                    strConn1 = "order by AA.KOTA"
            End Select
        Case "ORDER"
            strConn = "select top 50 NO_WO, [DATE], FRONT, NO_CHASSIS, NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta WITH (NOLOCK) where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH (NOLOCK) WHERE USERDEL = '')" & gBantuan3
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " Order By NO_WO DESC"
                Case 1
                    strConn1 = " Order By [DATE] DESC"
                Case 2
                    strConn1 = " Order By FRONT"
                Case 3
                    strConn1 = " Order By FS_KD_PLG"
                Case 4
                    strConn1 = " Order By CUSTOMER"
                Case 5
                    strConn1 = " Order By CUST_ADRES"
                Case 6
                    strConn1 = " Order By NO_CHASSIS"
                Case 7
                    strConn1 = " Order By NO_POLISI"
            End Select
        Case "HISTORY"
            strConn = "select top 50 AA.* from (Select distinct NO_POLISI, CUSTOMER from HISTORY WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND CUSTOMER <> '') AA "
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by AA.NO_POLISI"
                Case 1
                    strConn1 = " order by AA.CUSTOMER"
            End Select
        Case "REGISTBATAL"
            strConn = "select top 50 AA.* from(select NO_WO, CUSTOMER, NO_POLISI, RPR_COND from jobdta WITH (NOLOCK) where DELIV<>'BTL' AND RPR_COND = 'K') AA "
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by AA.NO_WO DESC"
                Case 1
                    strConn1 = " order by AA.CUSTOMER"
                Case 2
                    strConn1 = " order by AA.NO_POLISI"
                Case 3
                    strConn1 = " order by AA.RPR_COND"
            End Select
        Case "PROGRESS"
            strConn = "select top 50 NO_WO, CUSTOMER, NO_POLISI, RPR_COND from jobdta WITH (NOLOCK) where DELIV<>'BTL' AND RPR_COND <> 'K'"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_WO DESC"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by NO_POLISI"
                Case 3
                    strConn1 = " order by RPR_COND"
            End Select
        Case "REG"
            strConn = "select Top 50 NO_WO, CUSTOMER, NO_POLISI, RPR_COND from jobdta WITH (NOLOCK) where DELIV<>'BTL' AND NOINV <>'' AND RPR_COND <> 'K'"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_WO DESC"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by NO_POLISI"
                Case 3
                    strConn1 = " order by RPR_COND"
            End Select
        Case "FORMATSMS"
            strConn = "select top 50 TypeSMS,Pesan1 from M_Template"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn1 = " order by TypeSMS"
                    Case 1
                        strConn1 = " order by TypeSMS"
                End Select
        Case "DEALER"
            If gBantuan = "INTERNAL" Then
                strConn = "select Top 50 FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn1 = " order by FS_KD_DEALER"
                    Case 1
                        strConn1 = " order by FS_NM_DEALER"
                End Select
            ElseIf gBantuan = "ALL" Then
                strConn = "select top 50 FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn1 = " order by FS_KD_DEALER"
                    Case 1
                        strConn1 = " order by FS_NM_DEALER"
                End Select
            End If
        Case "BOOKBATAL"
            strConn = "select Top 50 NO_POLISI, CUSTOMER, CUST_ADRES, KOTA, FS_KD_DEALER from book WITH (NOLOCK) where DELIV='BOK'"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_POLISI"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by CUST_ADRES"
                Case 3
                    strConn1 = " order by KOTA"
                Case 4
                    strConn1 = " order by FS_KD_DEALER"
            End Select
        Case "METERAI"
            strConn = "select FN_BATAS, FN_NILAI from t_meterai WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by FN_BATAS"
                Case 1
                    strConn1 = " order by FN_NILAI"
            End Select
        Case "ORDERBATAL"
            strConn = "select top 50 NO_WO, [DATE], NO_CHASSIS, NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta WITH (NOLOCK) where DELIV <> 'BTL' AND RPR_COND = '' AND NOINV = '' AND NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH (NOLOCK) WHERE USERDEL = '')"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " Order By NO_WO DESC"
                Case 1
                    strConn1 = " Order By [DATE] DESC"
                Case 2
                    strConn1 = " Order By FS_KD_PLG"
                Case 3
                    strConn1 = " Order By CUSTOMER"
                Case 4
                    strConn1 = " Order By CUST_ADRES"
                Case 5
                    strConn1 = " Order By NO_CHASSIS"
                Case 6
                    strConn1 = " Order By NO_POLISI"
            End Select
        Case "LIHATORDER"
            strConn = "select top 50 NO_WO, CUSTOMER, CUST_ADRES, KOTA, FS_KD_DEALER from jobdta WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " Order By NO_WO DESC"
                Case 1
                    strConn1 = " Order By CUSTOMER"
                Case 2
                    strConn1 = " Order By CUST_ADRES"
                Case 3
                    strConn1 = " Order By KOTA"
                Case 4
                    strConn1 = " Order By FS_KD_DEALER"
            End Select
        Case "INV"
            strConn = "select top 50 NO_WO, [DATE] AS TANGGAL, FRONT, FS_KD_PLG, CUSTOMER, NO_POLISI, NOINV from jobdta WITH (NOLOCK) where (DELIV='PRC' or DELIV='OKE') AND NO_WO IN (SELECT NO_WO FROM VALIDORDER WITH (NOLOCK) WHERE USERDEL = '') " & gBantuan & ""
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_WO DESC"
                Case 1
                    strConn1 = " order by NOINV DESC"
                Case 2
                    strConn1 = " order by [DATE] DESC"
                Case 3
                    strConn1 = " order by FRONT"
                Case 4
                    strConn1 = " order by FS_KD_PLG"
                Case 5
                    strConn1 = " order by CUSTOMER"
                Case 6
                    strConn1 = " order by NO_POLISI"
            End Select
        Case "INV2"
            strConn = "select TOP 50 NO_WO, [DATE] AS TANGGAL, FRONT, FS_KD_PLG, CUSTOMER, NO_POLISI, NOINV, NO_CHASSIS, RPR_COND from jobdta WITH (NOLOCK) where NO_WO <> '' " & gBantuan & ""
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_WO DESC"
                Case 1
                    strConn1 = " order by [DATE] DESC"
                Case 2
                    strConn1 = " order by NO_POLISI, [DATE] DESC "
                Case 3
                    strConn1 = " order by NOINV DESC"
                Case 4
                    strConn1 = " order by FRONT, [DATE] DESC"
                Case 5
                    strConn1 = " order by FS_KD_PLG, [DATE] DESC"
                Case 6
                    strConn1 = " order by CUSTOMER, [DATE] DESC"
                Case 7
                    strConn1 = " order by NO_CHASSIS, [DATE] DESC"
            End Select
        Case "INVCTKULANG"
            strConn = "select top 50 NOINV, FS_KD_PLG, CUSTOMER, NO_POLISI, NO_WO from T_INVOICE WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NOINV DESC"
                Case 1
                    strConn1 = " order by FS_KD_PLG"
                Case 2
                    strConn1 = " order by CUSTOMER"
                Case 3
                    strConn1 = " order by NO_POLISI"
                Case 4
                    strConn1 = " order by NO_WO DESC"
            End Select
        Case "INVULANG"
            strConn = "select top 50 NO_WO, FS_KD_PLG, CUSTOMER, NO_POLISI, NOINV from T_INVOICE WITH (NOLOCK) where DELIV='OKE'"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NOINV DESC"
                Case 1
                    strConn1 = " order by NOINV DESC, NO_WO DESC"
                Case 2
                    strConn1 = " order by NOINV DESC, FS_KD_PLG"
                Case 3
                    strConn1 = " order by NOINV DESC, CUSTOMER"
                Case 4
                    strConn1 = " order by NOINV DESC, NO_POLISI"
            End Select
        Case "INVBATAL"
            strConn = "select top 50 NOINV, CUSTOMER, NO_POLISI, NO_WO, FS_NO_RETUR from T_INVOICE WITH (NOLOCK) where DELIV <> 'BTL' and RPR_COND <> 'K'"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NOINV DESC"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by NO_POLISI"
                Case 3
                    strConn1 = " order by NO_WO DESC"
                Case 4
                    strConn1 = " order by FS_NO_RETUR"
            End Select
        Case "INVEDIT"
            strConn = "select top 50 NOINV, CUSTOMER, NO_POLISI, NO_WO from T_INVOICE WITH (NOLOCK) where FS_NO_RETUR = ''"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NOINV DESC"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by NO_POLISI"
                Case 3
                    strConn1 = " order by NO_WO DESC"
            End Select
        Case "EDITRATE"
            strConn = "select top 50 NO_WO, NOINV, CUSTOMER, NO_POLISI from JOBDTA WITH (NOLOCK) where DELIV <> 'BTL'"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_WO DESC"
                Case 1
                    strConn1 = " order by NOINV DESC"
                Case 2
                    strConn1 = " order by CUSTOMER"
                Case 3
                    strConn1 = " order by NO_POLISI"
            End Select
        Case "INVCATAT"
            strConn = "select top 50 NOINV, CUSTOMER, NO_POLISI, NO_WO from T_INVOICE WITH (NOLOCK) where FS_NO_RETUR = '' and FS_KREDIT=''"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NOINV DESC"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by NO_POLISI"
                Case 3
                    strConn1 = " order by NO_WO DESC"
            End Select
        Case "INVBATALCATAT"
            strConn = "select top 50 NOINV, CUSTOMER, NO_POLISI, NO_WO from T_INVOICE WITH (NOLOCK) where FS_NO_RETUR = '' and FS_KREDIT='1'"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NOINV DESC"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by NO_POLISI"
                Case 3
                    strConn1 = " order by NO_WO DESC"
            End Select
        Case "INVBAYAR"
            strConn = "Select top 50 NOINV, CUSTOMER, NO_POLISI from T_INVOICE WITH (NOLOCK) where INVDATE='" & TglYYYYMMDD(gTglInv) & "' and FS_NO_RETUR = ''"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NOINV DESC"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by NO_POLISI"
            End Select
        Case "MEKANIK"
            If gBantuan = "BOOKING" Then
                strConn = "Select top 50 AA.* from (select DISTINCT A.MECH,[NAME],[GROUP], LEVMECH from mekanik A WITH (NOLOCK) LEFT JOIN T_MEKANIK B WITH (NOLOCK) ON A.MECH = B.MECH AND hari = '" & gBantuan2 & "' AND TANGGAL = '" & gBantuan3 & "' LEFT JOIN T_STALL C WITH (NOLOCK) ON C.HARI = '" & gBantuan2 & "' WHERE DATEDIFF(n,WAKTUMULAI,WAKTUSELESAI) - ISNULL(ALOKASI,0) > 0 ) AA "
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn1 = " order by AA.MECH"
                    Case 1
                        strConn1 = " order by AA.[NAME]"
                    Case 2
                        strConn1 = " order by AA.[GROUP]"
                End Select
            ElseIf gBantuan = "ALL" Then
            strConn = "Select top 50 AA.* FROm (Select DISTINCT MECH,[NAME],[GROUP], LEVMECH from mekanik A " & gBantuan4 & ") AA "
                    
                Select Case IndexBy.ListIndex
                    
                    Case 0, -1
                        strConn1 = " order by AA.MECH"
                    Case 1
                        strConn1 = " order by AA.[NAME]"
                    Case 2
                        strConn1 = " order by AA.[GROUP]"
                End Select
            ElseIf gBantuan = "ORDER" Then
                strConn = "select Top 50 AA.* from (Select DISTINCT A.MECH,[NAME],[GROUP], LEVMECH from mekanik A WITH (NOLOCK) LEFT JOIN T_MEKANIK B WITH (NOLOCK) ON A.MECH = B.MECH where hari = '" & gBantuan2 & "' AND TANGGAL = '" & gBantuan3 & "' AND (MENITKERJA-ALOKASI) > 0 AND MASUK = 1 AND JAMMULAI <= '" & Format(isnow, "HH:NN") & "' AND CASE WHEN JAMSELESAI  < JAMMULAI THEN CAST(LEFT(JAMSELESAI, CHARINDEX(':', JAMSELESAI)-1) + 24 AS varchar) + ':' + CAST(RIGHT(JAMSELESAI, CHARINDEX(':', JAMSELESAI)-1) AS varchar) ELSE JAMSELESAI END >= '" & Format(isnow, "HH:NN") & "') AA "
                Select Case IndexBy.ListIndex
                    Case 0, -1
                        strConn1 = " order by AA.MECH"
                    Case 1
                        strConn1 = " order by AA.[NAME]"
                    Case 2
                        strConn1 = " order by AA.[GROUP]"
                End Select
            End If
        Case "GROUP_MEKANIK"
            strConn = "select [GROUP],[NAME], MECH from mekanik WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by [GROUP]"
                Case 1
                    strConn1 = " order by [NAME]"
                Case 2
                    strConn1 = " order by MECH"
            End Select
        Case "FOREMAN"
            strConn = "select Top 50 AA.* from (Select DISTINCT FS_KD_PETUGAS, FS_NM_PETUGAS from T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'FOREMAN') AA "
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by AA.FS_KD_PETUGAS"
                Case 1
                    strConn1 = " order by AA.FS_NM_PETUGAS"
            End Select
        Case "MDL"
            strConn = "select Top 50 MDL, [DESC], MODEL, FS_GRP_MDL from model WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by MDL"
                Case 1
                    strConn1 = " order by [DESC]"
                Case 2
                    strConn1 = " order by MODEL"
                Case 3
                    strConn1 = " order by FS_GRP_MDL"
            End Select
        Case "JENIS_KERJA"
            strConn = "Select top 50 AA.* FROm (select DISTINCT JOB,REMARK from HO_FLATE WITH (NOLOCK) WHERE FS_LOG_DEL = '') AA "
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by AA.JOB"
                Case 1
                    strConn1 = " order by AA.REMARK"
            End Select
        Case "PETUGAS"
            strConn = "select FS_KD_PETUGAS, FS_NM_PETUGAS from t_petugas WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by FS_KD_PETUGAS"
                Case 1
                    strConn1 = " order by FS_NM_PETUGAS"
            End Select
        Case "PARTOLI"
            strConn = "select NOINV, NO_WO, NO_POLISI from T_INVOICE WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NOINV DESC"
                Case 1
                    strConn1 = " order by NO_WO DESC"
                Case 2
                    strConn1 = " order by NO_POLISI"
            End Select
        Case "JOBLIST2"
                strConn = "Select Top 50 AA.* from (select distinct FS_ITEM, RATE, KETERANGAN from HO_FLATE WITH (NOLOCK) where MDL = '" & gBantuan & "' AND JOB = '" & gBantuan2 & "' AND FS_LOG_DEL = '') AA "
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by AA.FS_ITEM"
                Case 2
                    strConn1 = " order by AA.KETERANGAN"
            End Select
        Case "GROUP_MODEL"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "select distinct FS_GRP_MDL from MODEL WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " order by FS_GRP_MDL"
            End Select
            
        ' ,;; ==========00000ooooo00000========== TAMBAHAN FUNGSI ==========00000ooooo00000========== ;;, '
        ' ,;; ==========00000ooooo00000========== TAMBAHAN FUNGSI ==========00000ooooo00000========== ;;, '
        Case "CUSTOMER2"
            strConn = "select top 50 NO_CHASSIS, NO_POLISI, CUSTOMER, CUST_ADRES, NO_ENGINE from CUSTDAT WITH (NOLOCK)"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " order by NO_CHASSIS"
                Case 1
                    strConn1 = " order by CUSTOMER"
                Case 2
                    strConn1 = " order by CUST_ADRES"
                Case 3
                    strConn1 = " order by NO_ENGINE"
                Case 4
                    strConn1 = " order by NO_POLISI"
            End Select
        Case "JOB_ORDER"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT FS_ITEM, KETERANGAN FROM HO_FLATE WITH (NOLOCK) WHERE JOB = '" & gBantuan & "' AND FS_LOG_DEL = ''"
                    strConn1 = " ORDER BY FS_ITEM, KETERANGAN"
            End Select
        Case "KELAS_MOBIL"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT CLASSCODE, CLASSNAME FROM CARCLASS WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " ORDER BY CLASSCODE"
                Case 1
                    strConn = "SELECT CLASSCODE, CLASSNAME FROM CARCLASS WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " ORDER BY CLASSNAME"
            End Select
        Case "AGAMA"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT AGAMA FROM T_AGAMA WITH (NOLOCK)"
                    strConn1 = " ORDER BY AGAMA"
            End Select
        Case "LEVMEKANIK"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT LEVELMECH, KETERANGAN FROM T_MECHLEV WITH (NOLOCK)"
                    strConn1 = " ORDER BY LEVELMECH"
            End Select
        Case "SUBLET"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT NAMABARANG FROM JOBSUBLET WITH (NOLOCK)"
                    strConn1 = " ORDER BY NAMABARANG"
            End Select
        Case "POSISI_MOBIL"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT posisiAWAL FROM AntriCuci WITH (NOLOCK)"
                    strConn1 = " ORDER BY posisiAWAL"
            End Select
        Case "TPSVC"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT TPSERVICE FROM T_REPO WITH (NOLOCK) WHERE TPSERVICE <> ''"
                    strConn1 = " ORDER BY TPSERVICE"
            End Select
        Case "TP_ORD"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT TPORDER FROM T_REPO WITH (NOLOCK) WHERE TPORDER <> ''"
                    strConn1 = " ORDER BY TPORDER"
            End Select
        Case "PROSES"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT PROSES FROM T_REPO WITH (NOLOCK) WHERE PROSES <> ''"
                    strConn1 = " ORDER BY PROSES"
            End Select
        Case "RETURN"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT RETUR FROM T_REPO WITH (NOLOCK) WHERE RETUR <> ''"
                    strConn1 = " ORDER BY RETUR DESC"
            End Select
        Case "KONDISI"
            Select Case gBantuan
                Case 1
                    Select Case IndexBy.ListIndex
                        Case 0, -1
                            strConn = "SELECT WAKTU FROM T_KONDISI WITH (NOLOCK)"
                            strConn1 = " ORDER BY ID"
                    End Select
                Case 2
                    Select Case IndexBy.ListIndex
                        Case 0, -1
                            strConn = "SELECT FREKUENSI FROM T_KONDISI WITH (NOLOCK)"
                            strConn1 = " ORDER BY ID"
                    End Select
                Case 3
                    Select Case IndexBy.ListIndex
                        Case 0, -1
                            strConn = "SELECT TEMPAT FROM T_KONDISI WITH (NOLOCK)"
                            strConn1 = " ORDER BY ID"
                    End Select
                Case 4
                    Select Case IndexBy.ListIndex
                        Case 0, -1
                            strConn = "SELECT WARN FROM T_KONDISI WITH (NOLOCK)"
                            strConn1 = " ORDER BY ID"
                    End Select
                Case 5
                    Select Case IndexBy.ListIndex
                        Case 0, -1
                            strConn = "SELECT MULAI FROM T_KONDISI WITH (NOLOCK)"
                            strConn1 = " ORDER BY ID"
                    End Select
                Case 6
                    Select Case IndexBy.ListIndex
                        Case 0, -1
                            strConn = "SELECT SELESAI FROM T_KONDISI WITH (NOLOCK)"
                            strConn1 = " ORDER BY ID"
                    End Select
            End Select
        Case "BOOKING"
            strConn = "SELECT top 50 AA.* from (Select DISTINCT NOFAK, [DATE], NO_POLISI, DELIV, CUSTOMER, CUST_ADRES, KOTA, FMDL FROM BOOK WITH (NOLOCK) WHERE [DATE] >= '" & Format(isnow, "MM/DD/YYYY") & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE') AA "
            Select Case IndexBy.ListIndex
                '(DELIV = 'BOK' OR LEFT(DELIV,3) = 'RES')
                Case 0, -1
                    strConn1 = " ORDER BY AA.NOFAK DESC"
                Case 1
                    strConn1 = " ORDER BY AA.[DATE] DESC"
                Case 2
                    strConn1 = " ORDER BY AA.NO_POLISI"
                Case 3
                    strConn1 = " ORDER BY AA.CUSTOMER"
                Case 4
                    strConn1 = " ORDER BY AA.CUST_ADRES"
                Case 5
                    strConn1 = " ORDER BY AA.KOTA"
                Case 6
                    strConn1 = " ORDER BY AA.FMDL"
            End Select
        Case "BTLBOOK"
            strConn = "SELECT top 50 AA.* from (Select DISTINCT NOFAK, NO_POLISI, CUSTOMER, CUST_ADRES, KOTA, FMDL, DELIV FROM BOOK WITH (NOLOCK) WHERE DELIV <> 'OKE') AA "
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = " ORDER BY AA.NOFAK DESC"
                Case 1
                    strConn1 = " ORDER BY AA.NO_POLISI"
                Case 2
                    strConn1 = " ORDER BY AA.CUSTOMER"
                Case 3
                    strConn1 = " ORDER BY AA.CUST_ADRES"
                Case 4
                    strConn1 = " ORDER BY AA.KOTA"
                Case 5
                    strConn1 = " ORDER BY AA.FMDL"
            End Select
        Case "PROGRAM"
            strConn = "SELECT top 50 AA.* from (Select DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI) AND FS_LOG_DEL = '' " & gBantuan & " ) AA "
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn1 = "ORDER BY AA.ID_PROGRAM"
                Case 1
                    strConn1 = "ORDER BY AA.PROGRAM"
                Case 2
                    strConn1 = "ORDER BY AA.KETERANGAN"
                Case 3
                    strConn1 = "ORDER BY AA.TGLMULAI"
                Case 4
                    strConn1 = "ORDER BY AA.TGLSELESAI"
            End Select
        Case "PROGRAM_ALL"
            'ANTZTEIN SAMPAI DISINI ..............
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE isnull(A.id_program,'') <> ''  " & gBantuan & " "
                    'strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A LEFT JOIN HO_PROCBG B ON A.ID_PROGRAM = B.ID_PROGRAM WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)  " & gBantuan & " "
                    strConn1 = "ORDER BY A.ID_PROGRAM"
                Case 1
                    strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE isnull(A.id_program,'') <> ''  " & gBantuan & " "
                    'strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A LEFT JOIN HO_PROCBG B ON A.ID_PROGRAM = B.ID_PROGRAM WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)  " & gBantuan & " "
                    strConn1 = "ORDER BY PROGRAM"
                Case 2
                    strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE isnull(A.id_program,'') <> ''  " & gBantuan & " "
                    'strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A LEFT JOIN HO_PROCBG B ON A.ID_PROGRAM = B.ID_PROGRAM WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)  " & gBantuan & " "
                    strConn1 = "ORDER BY KETERANGAN"
                Case 3
                    strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE isnull(A.id_program,'') <> ''  " & gBantuan & " "
                    'strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A LEFT JOIN HO_PROCBG B ON A.ID_PROGRAM = B.ID_PROGRAM WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)  " & gBantuan & " "
                    strConn1 = "ORDER BY TGLMULAI"
                Case 4
                    strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE isnull(A.id_program,'') <> ''  " & gBantuan & " "
                    'strConn = "SELECT DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A LEFT JOIN HO_PROCBG B ON A.ID_PROGRAM = B.ID_PROGRAM WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)  " & gBantuan & " "
                    strConn1 = "ORDER BY TGLSELESAI"
            End Select

        Case "WARRANTY"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH (NOLOCK) WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI) "
                    strConn1 = " ORDER BY ID_WARRANTY"
                Case 1
                    strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH (NOLOCK) WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI) "
                    strConn1 = " ORDER BY WARRANTY"
                Case 2
                    strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH (NOLOCK) WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI) "
                    strConn1 = " ORDER BY KETERANGAN"
                Case 3
                    strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH (NOLOCK) WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI) "
                    strConn1 = " ORDER BY TGLMULAI"
                Case 4
                    strConn = "SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_WARRANTY WITH (NOLOCK) WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI) "
                    strConn1 = " ORDER BY TGLSELESAI"
            End Select
        Case "ASURANSI"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT ASURANSI, PIC, ALAMAT, KOTA FROM T_ASURANSI WITH (NOLOCK)"
                    strConn1 = " ORDER BY ASURANSI"
                Case 1
                    strConn = "SELECT DISTINCT ASURANSI, PIC, ALAMAT, KOTA FROM T_ASURANSI WITH (NOLOCK)"
                    strConn1 = " ORDER BY PIC"
                Case 2
                    strConn = "SELECT DISTINCT ASURANSI, PIC, ALAMAT, KOTA FROM T_ASURANSI WITH (NOLOCK)"
                    strConn1 = " ORDER BY ALAMAT"
                Case 3
                    strConn = "SELECT DISTINCT ASURANSI, PIC, ALAMAT, KOTA FROM T_ASURANSI WITH (NOLOCK)"
                    strConn1 = " ORDER BY KOTA"
            End Select
        Case "ESTPOLISI"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT NO_POLISI, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY NO_POLISI"
                Case 1
                    strConn = "SELECT DISTINCT NO_POLISI, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY NOEST"
                Case 2
                    strConn = "SELECT DISTINCT NO_POLISI, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY CUSTOMER"
                Case 3
                    strConn = "SELECT DISTINCT NO_POLISI, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY CUST_ADRES"
                Case 4
                    strConn = "SELECT DISTINCT NO_POLISI, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY FMDL"
                Case 5
                    strConn = "SELECT DISTINCT NO_POLISI, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY STATUS"
            End Select
        Case "ESTRANGKA"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT NO_RANGKA, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY NO_RANGKA"
                Case 1
                    strConn = "SELECT DISTINCT NO_RANGKA, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY NOEST"
                Case 2
                    strConn = "SELECT DISTINCT NO_RANGKA, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY CUSTOMER"
                Case 3
                    strConn = "SELECT DISTINCT NO_RANGKA, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY CUST_ADRES"
                Case 4
                    strConn = "SELECT DISTINCT NO_RANGKA, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY FMDL"
                Case 5
                    strConn = "SELECT DISTINCT NO_RANGKA, NOEST, CUSTOMER, CUST_ADRES, FMDL, COLOR, STATUS FROM ESTIMASI WITH (NOLOCK)"
                    strConn1 = " ORDER BY STATUS"
            End Select
        Case "BTLREG"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "select NO_WO, FS_KD_PLG, CUSTOMER, NO_POLISI, NOINV from jobdta WITH (NOLOCK) where DELIV <> 'BTL' and RPR_COND = 'K'"
                    strConn1 = " order by NO_WO DESC"
                Case 1
                    strConn = "select NO_WO, FS_KD_PLG, CUSTOMER, NO_POLISI, NOINV from jobdta WITH (NOLOCK) where DELIV <> 'BTL' and RPR_COND = 'K'"
                    strConn1 = " order by FS_KD_PLG"
                Case 2
                    strConn = "select NO_WO, FS_KD_PLG, CUSTOMER, NO_POLISI, NOINV from jobdta WITH (NOLOCK) where DELIV <> 'BTL' and RPR_COND = 'K'"
                    strConn1 = " order by CUSTOMER"
                Case 3
                    strConn = "select NO_WO, FS_KD_PLG, CUSTOMER, NO_POLISI, NOINV from jobdta WITH (NOLOCK) where DELIV <> 'BTL' and RPR_COND = 'K'"
                    strConn1 = " order by NO_POLISI"
                Case 4
                    strConn = "select NO_WO, FS_KD_PLG, CUSTOMER, NO_POLISI, NOINV from jobdta WITH (NOLOCK) where DELIV <> 'BTL' and RPR_COND = 'K'"
                    strConn1 = " order by NOINV DESC"
            End Select
        Case "SA"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT FS_KD_PETUGAS, FS_NM_PETUGAS FROM T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA'"
                    strConn1 = " ORDER BY FS_KD_PETUGAS"
                Case 1
                    strConn = "SELECT FS_KD_PETUGAS, FS_NM_PETUGAS FROM T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA'"
                    strConn1 = " ORDER BY FS_NM_PETUGAS"
            End Select
        Case "DEALERINV"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER FROM T_INVOICE A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS " _
                            & "WHERE A.FS_LOG_DEL = '' AND RIGHT(FS_RC,3) = '001' AND INVDATE BETWEEN '" & Format(gBantuan, "MM/DD/YYYY") & "' AND '" & Format(gBantuan2, "MM/DD/YYYY") & "'"
                    strConn1 = " ORDER BY B.FS_KD_DEALER"
                Case 1
                    strConn = "SELECT B.FS_KD_DEALER, B.FS_NM_DEALER FROM T_INVOICE A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS " _
                            & "WHERE A.FS_LOG_DEL = '' AND RIGHT(FS_RC,3) = '001' AND INVDATE BETWEEN '" & Format(gBantuan, "MM/DD/YYYY") & "' AND '" & Format(gBantuan2, "MM/DD/YYYY") & "'"
                    strConn1 = " ORDER BY B.FS_NM_DEALER"
            End Select
        Case "TIPEPETUGAS"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT TIPE FROM TIPEPETUGAS WITH (NOLOCK) "
                    strConn1 = ""
                Case 1
                    strConn = "SELECT TIPE FROM TIPEPETUGAS WITH (NOLOCK) "
                    strConn1 = " ORDER BY TIPE"
            End Select
        Case "COUNTRY"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT KODE, KETERANGAN FROM T_COUNTRY WITH (NOLOCK) "
                    strConn1 = " ORDER BY KODE"
                Case 1
                    strConn = "SELECT KODE, KETERANGAN FROM T_COUNTRY WITH (NOLOCK) "
                    strConn1 = " ORDER BY KODE"
                Case 2
                    strConn = "SELECT KODE, KETERANGAN FROM T_COUNTRY WITH (NOLOCK) "
                    strConn1 = " ORDER BY KETERANGAN"
            End Select
        Case "STATUSAKUN"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT TIPE, KETERANGAN FROM T_STATUSAKUN WITH (NOLOCK) "
                    strConn1 = " ORDER BY TIPE"
                Case 1
                    strConn = "SELECT TIPE, KETERANGAN FROM T_STATUSAKUN WITH (NOLOCK) "
                    strConn1 = " ORDER BY TIPE"
                Case 2
                    strConn = "SELECT TIPE, KETERANGAN FROM T_STATUSAKUN WITH (NOLOCK) "
                    strConn1 = " ORDER BY KETERANGAN"
            End Select
        Case "VALIDSUBLET"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND VALID = 0"
                    strConn1 = " Order By A.NO_WO DESC"
                Case 1
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND VALID = 0"
                    strConn1 = " Order By [DATE] DESC"
                Case 2
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND VALID = 0"
                    strConn1 = " Order By FS_KD_PLG"
                Case 3
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND VALID = 0"
                    strConn1 = " Order By CUSTOMER"
                Case 4
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND VALID = 0"
                    strConn1 = " Order By CUST_ADRES"
                Case 5
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND VALID = 0"
                    strConn1 = " Order By A.NO_CHASSIS"
                Case 6
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND VALID = 0"
                    strConn1 = " Order By A.NO_POLISI"
            End Select
        Case "RETURSUBLET"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WHERE USERDEL = '') AND VALID = 1"
                    strConn1 = " Order By A.NO_WO DESC"
                Case 1
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WHERE USERDEL = '') AND VALID = 1"
                    strConn1 = " Order By [DATE] DESC"
                Case 2
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WHERE USERDEL = '') AND VALID = 1"
                    strConn1 = " Order By FS_KD_PLG"
                Case 3
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WHERE USERDEL = '') AND VALID = 1"
                    strConn1 = " Order By CUSTOMER"
                Case 4
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WHERE USERDEL = '') AND VALID = 1"
                    strConn1 = " Order By CUST_ADRES"
                Case 5
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WHERE USERDEL = '') AND VALID = 1"
                    strConn1 = " Order By A.NO_CHASSIS"
                Case 6
                    strConn = "select DISTINCT A.NO_WO, [DATE], A.NO_CHASSIS, A.NO_POLISI, FS_KD_PLG, CUSTOMER, CUST_ADRES from jobdta A WITH (NOLOCK) INNER JOIN JOBSUBLET B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO where DELIV <> 'BTL' AND RPR_COND <> 'K' AND NOINV = '' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WHERE USERDEL = '') AND VALID = 1"
                    strConn1 = " Order By A.NO_POLISI"
            End Select
        Case "STALL"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "select DISTINCT NOSTALL FROM STALL WITH (NOLOCK) "
                    strConn1 = " order by CAST(SUBSTRING(NOSTALL,CASE WHEN LEN(NOSTALL) >7 THEN 7 ELSE 4 END,2) AS INTEGER)"
                End Select
        Case "GRPMDL"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "select DISTINCT LEFT(MDL,CHARINDEX('-',MDL)-1) as MDL from model WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " order by LEFT(MDL,CHARINDEX('-',MDL)-1)"
            End Select
        Case "ALASANDTG"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "select DISTINCT ALASAN from T_ALASAN WITH (NOLOCK) "
                    strConn1 = " order by ALASAN"
            End Select
        Case "NPWP"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT NPWP, NAMAFP, ALAMATFP1, ALAMATFP2 FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) "
                    strConn1 = " ORDER BY NPWP DESC"
                Case 1
                    strConn = "SELECT DISTINCT NPWP, NAMAFP, ALAMATFP1, ALAMATFP2 FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) "
                    strConn1 = " ORDER BY NAMAFP,NPWP DESC"
                Case 2
                    strConn = "SELECT DISTINCT NPWP, NAMAFP, ALAMATFP1, ALAMATFP2 FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) "
                    strConn1 = " ORDER BY ALAMATFP1,NPWP DESC"
                Case 3
                    strConn = "SELECT DISTINCT NPWP, NAMAFP, ALAMATFP1, ALAMATFP2 FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH (NOLOCK) "
                    strConn1 = " ORDER BY ALAMATFP2,NPWP DESC"
            End Select
        Case "TPPRINTER"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT TPPRINT FROM TIPEPRINTER WITH (NOLOCK) "
                    strConn1 = " ORDER BY TPPRINT"
            End Select
        Case "ANTRIAN"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    'strConn = "SELECT DISTINCT QUEUE, TIPE FROM " & Log.LogQueue & "..PANGGILAN WHERE SA = '" & Trim(gBantuan) & "' AND NOPOL = '' AND ID NOT IN (SELECT IDCALL FROM LISTANTRIAN) AND LEN(TIPE) = 3 AND CONVERT(varchar(10), TGL_PANGGIL, 101) = '" & gBantuan2 & "' AND CONVERT(varchar(8), TGL_PANGGIL, 8) <= '" & gBantuan3 & "'"
                    
                    'UPDATE GANTI ANTRIAN DENGAN EM
                    
'                    strConn = "SELECT DISTINCT QUEUE, TIPE FROM " & Log.LogQueue & "..PANGGILAN WHERE SA = '" & Trim(gBantuan) & "' AND NOPOL = '' AND ID NOT IN (SELECT IDCALL FROM LISTANTRIAN) AND TIPE IN ('BOK','NON','EM') AND CONVERT(varchar(10), TGL_PANGGIL, 101) = '" & gBantuan2 & "' AND CONVERT(varchar(8), TGL_PANGGIL, 8) <= '" & gBantuan3 & "'"
'                    strConn1 = " ORDER BY TIPE, QUEUE DESC"

'                    strConn = "SELECT DISTINCT QUEUE, TIPE,ID FROM " & Log.LogQueue & "..PANGGILAN WITH (NOLOCK) WHERE SA = '" & Trim(gBantuan) & "' AND NOPOL = '' AND ID NOT IN (SELECT IDCALL FROM LISTANTRIAN WITH (NOLOCK) WHERE  tanggal = '" & gBantuan2 & "') AND TIPE IN ('BOK','NON','EM') AND CONVERT(varchar(10), TGL_PANGGIL, 101) = '" & gBantuan2 & "'" ' AND CONVERT(varchar(8), TGL_PANGGIL, 8) <= '" & gBantuan3 & "'"
'                    strConn1 = " ORDER BY TIPE, QUEUE DESC"
                    
                    ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                    '' 2017-06-10
                    ''-----------------------------------------------------------
                    strConn = "SELECT DISTINCT a.QUEUE, a.TIPE as NAMA, b.SUBMENU as TIPE, a.ID FROM " & Log.LogQueue & "..PANGGILAN a WITH (NOLOCK)INNER JOIN " & Log.LogQueue & "..MENU b WITH (NOLOCK) ON b.Tipe = a.Tipe and b.Status = 1 and b.MENU = 'GRP' WHERE SA = '" & Trim(gBantuan) & "' AND NOPOL = '' AND a.ID NOT IN (SELECT IDCALL FROM LISTANTRIAN WITH (NOLOCK) WHERE  tanggal = '" & gBantuan2 & "') AND CONVERT(varchar(10), TGL_PANGGIL, 101) = '" & gBantuan2 & "'  "
                    strConn1 = " ORDER BY TIPE, QUEUE DESC"
                    ''
            End Select
        Case "TIPEESTMASI"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT KETERANGAN FROM T_HITESTIMASI WITH (NOLOCK) "
                    strConn1 = " ORDER BY KETERANGAN"
            End Select
        Case "DLRSHIFT"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT DISTINCT SHIFT, JAMMULAI, JAMSELESAI FROM DEALERSHIFT WITH (NOLOCK) " & gBantuan
                    strConn1 = " ORDER BY SHIFT "
            End Select
        Case "MOU"
            Select Case IndexBy.ListIndex
                Case 0, -1
                    strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " ORDER BY NO_MOU"
                Case 1
                    strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " ORDER BY PERUSAHAAN"
                Case 2
                    strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " ORDER BY PIC"
                Case 3
                    strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " ORDER BY ALAMAT"
                Case 4
                    strConn = "SELECT NO_MOU, PERUSAHAAN, PIC, ALAMAT, KOTA FROM MOUHDR WITH (NOLOCK) WHERE FS_LOG_DEL = ''"
                    strConn1 = " ORDER BY KOTA"
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
    If DGrid.Columns(IndexBy.ListIndex).DataField <> "" Then
        txtcari.SelStart = Len(txtcari)
        If InStr(UCase(strConn), "WHERE") > 0 Then
        Dim strConnAsli As String
        strConnAsli = strConn
        If InStr(1, strConn, "AA.") > 0 And InStr(1, Right(strConn, 3), "AA") > 0 Then
           strConn = Mid(strConn, 1, Len(strConn) - 5)
        End If
        
            Select Case DGrid.Columns(IndexBy.ListIndex).DataField
                Case "MECH"
                   Call refresh_Grid(strConn & " AND A.[" & DGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                Case "FS_KD_DEALER"
                    If gSearchOption = "DEALERINV" Then
                       Call refresh_Grid(strConn & " AND B.[" & DGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                    Else
                       Call refresh_Grid(strConn & " AND [" & DGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                    End If
                Case "NO_WO"
                    If gSearchOption = "VALIDSUBLET" Or gSearchOption = "RETURSUBLET" Then
                        Call refresh_Grid(strConn & " AND A.[" & DGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                    Else
                        Call refresh_Grid(strConn & " AND [" & DGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                    End If
                Case "TANGGAL"
                    Call refresh_Grid(strConn & " AND [DATE] LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                Case "ID_PROGRAM"
                    refresh_Grid (strConn & " AND A.ID_PROGRAM LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                Case "NO_POLISI"
                    Select Case gSearchOption
                        Case "VALIDSUBLET", "RETURSUBLET"
                            Call refresh_Grid(strConn & " AND A.NO_POLISI LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                        Case Else
                            Call refresh_Grid(strConn & " AND NO_POLISI LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
                    End Select
                Case Else
                    Call refresh_Grid(strConn & " AND [" & DGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
            End Select
        Else
            Call refresh_Grid(strConn & " WHERE [" & DGrid.Columns(IndexBy.ListIndex).DataField & "] LIKE '%" & Trim(txtcari) & "%' " & IIf(InStr(1, strConnAsli, "AA.") > 0, ") AA ", "") & strConn1)
        End If
    End If
End Sub

Private Sub txtCari_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            DGrid.SetFocus
        Case vbKeyDown
            DGrid.SetFocus
        Case vbKeyReturn
            If Not DGrid.row = -1 Then
                gMainKeyCari = DGrid.Columns(0)
                If gSearchOption = "ESTPOLISI" Or gSearchOption = "ESTRANGKA" Then
                    frmEstimasi.passingXX = DGrid.Columns(1)
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

Private Sub txtcari_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub refresh_Grid(StrCon As String)

    If rsGrid.State Then rsGrid.Close
    Set rsGrid = objO.FetchData(StrCon)
    If Not rsGrid.EOF Then
        DGrid.DataSource = rsGrid
        Call panggilHead
    End If
End Sub
