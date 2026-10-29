VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.ocx"
Begin VB.Form frmHOTT1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Master T1"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   5760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   16744576
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   5175
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   11295
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CC"
         PromptChar      =   ">"
      End
      Begin TrueOleDBGrid70.TDBGrid GridJobOrder 
         Height          =   3735
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6588
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H77EAE1&"
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
         Index           =   1
         Left            =   1920
         TabIndex        =   10
         Top             =   600
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   1935
         TabIndex        =   12
         Top             =   960
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESC : "
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
         Left            =   1200
         TabIndex        =   13
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TERM : "
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
         Left            =   1140
         TabIndex        =   11
         Top             =   660
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID : "
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
         Left            =   1410
         TabIndex        =   6
         Top             =   300
         Width           =   330
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   7080
      TabIndex        =   1
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   65280
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
      Left            =   8040
      TabIndex        =   3
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   65535
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
      Left            =   9000
      TabIndex        =   4
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   255
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
      Left            =   9960
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12632256
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
   Begin ARProgBarCtrl.ARProgressBar ProgBar 
      Height          =   255
      Left            =   375
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   450
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
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   6195
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   10927
      StartColor      =   49152
      EndColor        =   12648447
      FormatString    =   "    Table Master T1"
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
End
Attribute VB_Name = "frmHOTT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOTT1
' DateTime  : 27/07/2015 11:18
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------

Dim rsGrid As New ADODB.Recordset       ' Recordset Grid
Dim rsTemp As ADODB.Recordset           ' Recordset Proses
Dim cols As TrueOleDBGrid70.Columns
Dim boleh As Boolean
Dim pesan As String
Dim gAtur As Boolean
'Dim sTampung As String

Private Sub ber_sih()
    For i = 0 To 2
        Text(i) = ""
    Next
End Sub

Private Sub buka_Tutup(Optional Ket As Boolean = True, Optional stat As Boolean = True)
    Text(0).Enabled = Ket
    Text(1).Enabled = stat
    Text(2).Enabled = stat
End Sub

Private Sub refreshGrid()
    GridJobOrder.ClearFields
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = objO.FetchDataHO("SELECT * FROM MASTER_T1 WITH(NOLOCK) WHERE ISHAPUS = '0' ORDER BY ID DESC ")
    If Not rsGrid.EOF Then
        GridJobOrder.DataSource = rsGrid
        GridJobOrder.Columns(0).Width = 800
        GridJobOrder.Columns(1).Width = 6000
        GridJobOrder.Columns(2).Width = 3500
        GridJobOrder.Columns(3).Width = 500
    End If
End Sub

Private Sub cboKM_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub chSBE_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub cmdCetak_Click()
    frmHOTT1Ctk.Show
End Sub

Private Sub cmdHapus_Click()
    If Trim(Text(0)) <> "" Then
        temp = MsgBox("Yakin akan menghapus data ID = " & Trim(Text(0)) & " ?", vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE master_t1 SET ishapus = '1', tglupdate = getdate() WHERE id = '" & Trim(Text(0)) & "' and ishapus = '0'")
            Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('M_T1','" & Trim(Text(0)) & "','D')")
            Call ber_sih
            Call refreshGrid
        End If
    Else
        MsgBox "Pilih Data yang akan dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdImport_Click()
      Dim ErrBegin As Boolean, konter As Long
1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim Jml As Integer
2       konter = 0
3       ProgBar.Min = 0
4       CommonDialog1.FileName = ""
5       CommonDialog1.InitDir = App.Path
6       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
7       CommonDialog1.ShowOpen
8       If CommonDialog1.FileName <> "" Then
9          If CnExel.State Then CnExel.Close
10         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;"""
11         GoSub ProsesImport
12     Else
13        Exit Sub
14     End If
15     Exit Sub
       
ProsesImport:
16    ErrBegin = True
         Dim VarNmFile As String
17       Call objHT.OpenTransactionHO
18       If RcExcel.State Then RcExcel.Close
19       RcExcel.Open "Select * from [TB_M_T1$]", CnExel, adOpenKeyset
20       If RcExcel.RecordCount > 0 Then
21          ProgBar.Visible = True
22          ProgBar.Max = RcExcel.RecordCount
23          RcExcel.MoveFirst
24          Do Until RcExcel.EOF
25             DoEvents
26             konter = konter + 1
27             ProgBar.value = konter
28             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
      '         If objO.FetchDataHO_Count("select id from master_t1 WITH(NOLOCK) WHERE ishapus = '0' and id = '" & Trim(RcExcel.Fields(0).value) & "'") > 0 Then
      '            Call objHT.CloseTransactionHO_err
      '            ProgBar.Visible = False
      '            MsgBox "ID " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " sudah ada!" & vbCrLf & "Proses dibatalkan", vbCritical, "Error"
      '            Exit Sub
      '         End If
               
29             If Nul(RcExcel.Fields(1).value) = "" Then
30                Call objHT.CloseTransactionHO_err
31                ProgBar.Visible = False
32                MsgBox "TERM baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
33                Exit Sub
34             End If
               
35             If Nul(RcExcel.Fields(2).value) = "" Then
36                Call objHT.CloseTransactionHO_err
37                ProgBar.Visible = False
38                MsgBox "DESC baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
39                Exit Sub
40             End If
               
41             Set rsTemp = objO.FetchDataHO("select id from master_t1 WITH(NOLOCK) WHERE ishapus = '0' and id = '" & Trim(RcExcel.Fields(0).value) & "'")
42            If Not rsTemp.EOF Then
              'UPDATE DATA
                  
43                Call objHT.UpdateData_Ho("UPDATE MASTER_T1 SET id = '" & Nul(RcExcel.Fields(0).value) & "',term = '" & Nul(RcExcel.Fields(1).value) & "',[desc] = '" & Nul(RcExcel.Fields(2).value) & "',tglupdate = getdate() WHERE ID = '" & Nul(RcExcel.Fields(0).value) & "' AND ISHAPUS = '0'")
44                Call objHT.UpdateData_Ho("INSERT INTO TEMPKIRIM VALUES ('M_T1','" & Nul(RcExcel.Fields(0).value) & "','S')")
45            Else
              ' INPUT DATA BARU
46                Call objHT.UpdateData_Ho("INSERT INTO MASTER_T1 VALUES ('" & Nul(RcExcel.Fields(0).value) & "','" & Nul(RcExcel.Fields(1).value) & "','" & Nul(RcExcel.Fields(2).value) & "',getdate(),getdate(),'0')")
47                Call objHT.UpdateData_Ho("INSERT INTO TEMPKIRIM VALUES ('M_T1','" & Nul(RcExcel.Fields(0).value) & "','S')")
48            End If
49             RcExcel.MoveNext
50          Loop
51       End If
52       RcExcel.Close
53       CnExel.Close
      '   ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
54        Call objHT.CloseTransactionHO
55       MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
56       Call refreshGrid
57       ProgBar.Visible = False
      '   GridAFI.Refresh
58       Return
59       Exit Sub
vb_error:
60       If ErrBegin = True Then
           ' DbCon.RollbackTrans
61       End If
62       ProgBar.Visible = False
63       MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
Dim jobOrd As String

    If cmdSimpan.Caption = "&Tambah" Then
        Call ber_sih
        Call buka_Tutup
        Text(0).SetFocus
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
        Me.Tag = "baru"
    ElseIf cmdSimpan.Caption = "&Simpan" Then
        boleh = True
        
        ' CHECK DATA KOSONG
        pesan = ""
        
        If Trim(Text(0)) = "" Then
            boleh = False
            pesan = pesan + "ID. "
        End If
        
        If Trim(Text(1)) = "" Then
            boleh = False
            pesan = pesan + "Term. "
        End If
        
        If Trim(Text(2)) = "" Then
            boleh = False
            pesan = pesan + "Desc. "
        End If
        
        ' JIKA DATA VALID
        If boleh Then
           
            ' PROSES SIMPAN
            Set rsTemp = objO.FetchDataHO("SELECT * FROM master_t1 WITH (NOLOCK) WHERE id = '" & Trim(Text(0)) & "' and ishapus = '0'")
            If Not rsTemp.EOF Then
                'UPDATE DATA
                Call objO.UpdateDataHO("UPDATE MASTER_T1 SET id = '" & Trim(Text(0)) & "',term = '" & Trim(Text(1)) & "',[desc] = '" & Trim(Text(2)) & "',tglupdate = getdate() WHERE ID = '" & Trim(Text(0)) & "' AND ISHAPUS = '0'")
                Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('M_T1','" & Trim(Text(0)) & "','S')")
            Else
                ' INPUT DATA BARU
                Call objO.UpdateDataHO("INSERT INTO MASTER_T1 VALUES ('" & Trim(Text(0)) & "','" & Trim(Text(1)) & "','" & Trim(Text(2)) & "',getdate(),getdate(),'0')")
                Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('M_T1','" & Trim(Text(0)) & "','S')")
            End If
        
            Call ber_sih
            Call buka_Tutup(False, False)
            Call refreshGrid
            cmdSimpan.Caption = "&Tambah"
            cmdUbah.Caption = "&Ubah"
        Else
            MsgBox pesan, vbExclamation, "KELENGKAPAN"
        End If
    End If
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call buka_Tutup(False, True)
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
        Me.Tag = "ubah"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdSimpan.Caption = "&Tambah"
        cmdUbah.Caption = "&Ubah"
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
    GridJobOrder.Width = Frame1.Width - 200
    GridJobOrder.Height = Frame1.Height - Text(2).Top - Text(2).Height - 200
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    ProgBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdImport.Left = cmdSimpan.Left - cmdImport.Width - 100
    ProgBar.Left = cmdImport.Left - ProgBar.Width - 100
End Sub

Private Sub GridJobOrder_FilterChange()
    On Error GoTo errHandler
    Set cols = GridJobOrder.Columns
    Dim C As Integer
    
    C = GridJobOrder.col
    GridJobOrder.HoldFields
    rsGrid.filter = getFilter()
    GridJobOrder.col = C
    GridJobOrder.EditActive = True
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

Private Sub GridJobOrder_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsGrid.State <> adStateClosed Then
        If Not rsGrid.EOF Then
            'If GridJobOrder.row > 0 Then
            Text(1) = rsGrid.Fields(1)
            Text(0) = rsGrid.Fields(0) 'GridJobOrder.Columns(0)
            Text(2) = rsGrid.Fields(2)   'gridJobOrder.Columns(2)
            'chSBE.value = IIf(rsGrid.Fields(3) = "Y", Checked, Unchecked)   'gridJobOrder.Columns(3)
            
            
        End If
    End If
End Sub

Private Sub text_GotFocus(Index As Integer)
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
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCariHO.Show (vbModal)
            Case 1
                If Text(0) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = Text(0)
                    frmCariHO.Show (vbModal)
                Else
                    MsgBox "Isi Tipe Job Order", vbCritical, "KELENGKAPAN"
                    Text(0).SetFocus
                    Exit Sub
                End If
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
    If Index = 0 Then
        If Me.Tag = "baru" Then
            Set rsTemp = objO.FetchDataHO("SELECT ID FROM MASTER_T1 WHERE ISHAPUS = '0' AND ID = '" & Trim(Text(0)) & "'")
            If Not rsTemp.EOF Then
                MsgBox "Id = " & Trim(Text(0)) & " sudah ada silahkan ganti dengan yang lain.", vbInformation
                Text(0) = ""
                Exit Sub
            End If
        End If
    End If
End Sub
