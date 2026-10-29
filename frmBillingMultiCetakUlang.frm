VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmBillingMultiCetakUlang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "FLATE RATE"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   13695
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
      Height          =   270
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   9600
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   3090
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12240
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   6615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   11655
      Begin TrueOleDBGrid70.TDBGrid GridFlate 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6800
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO INVOICE"
         Columns(0).DataField=   "noinv"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "TGL IMPORT"
         Columns(1).DataField=   "tglimport"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
         _StyleDefs(65)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(66)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(67)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(68)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7395
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13044
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "    Multiple Cetak Invoice"
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
      Left            =   7200
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
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
      Left            =   8160
      TabIndex        =   3
      Top             =   7200
      Visible         =   0   'False
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
      Left            =   9120
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
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
      Left            =   10080
      TabIndex        =   5
      Top             =   7200
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
      Left            =   7080
      TabIndex        =   6
      Top             =   7200
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
   Begin AIFCmp1.asxPowerButton cmdDisplay 
      Height          =   360
      Left            =   480
      TabIndex        =   7
      Top             =   7200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Refresh Data"
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
      Left            =   5625
      TabIndex        =   8
      Top             =   7200
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
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   7200
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
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
End
Attribute VB_Name = "frmBillingMultiCetakUlang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmBillingMultiCetakUlang
' DateTime  : 04/06/2024 09:02
' Author    : HO-VM-RUDI
' Purpose   :
'---------------------------------------------------------------------------------------

 
Dim rsGrid As New ADODB.Recordset
Dim rsCetakNewBilling As New ADODB.Recordset

Dim rsTemp As ADODB.Recordset
Dim newRecord As Boolean
Dim pesan As String
Dim boleh As Boolean
Dim SSQL2 As String
Dim statusSBE As String
Dim cols As TrueOleDBGrid70.Columns
 
Dim nilaiPPnInv     As Double
Dim nilaiDppInv     As Double
Dim nilaiTotalInv   As Double
Private Sub refreshGrid()
Dim nCol As Integer
Dim rsRec As New ADODB.Recordset
        
    ' PAKAI RECORDSET
   On Error GoTo refreshGrid_Error

    If rsRec.State Then rsRec.Close
    Set rsRec = objO.FetchData("SELECT noinv,tglentry FROM cetakMultipleInvoice WITH (NOLOCK) WHERE iscetak = 0 ORDER BY noinv")
      
        Call createRS
    If Not rsRec.EOF Then
        Do While Not rsRec.EOF
            rsGrid.AddNew
            rsGrid!noInv = Trim(rsRec("noinv"))
            rsGrid!tglimport = Format(rsRec("tglentry"), "dd-mm-yyyy")
            rsGrid.Update
            rsRec.MoveNext
        Loop
    Else
        Set rsGrid = Nothing
    End If
    GridFlate.DataSource = rsGrid

   On Error GoTo 0
   Exit Sub

refreshGrid_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form frmBillingMultiCetakUlang")
End Sub

Private Sub createRS()
   On Error GoTo createRS_Error

    On Error GoTo vb_error
    
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = Nothing
    With rsGrid
        .Fields.Append "noinv", adVarChar, 50
        .Fields.Append "tglImport", adDate
        .Open
    End With
    Set GridFlate.DataSource = rsGrid
    Exit Sub
    
vb_error:
   MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateRs"), vbCritical

   On Error GoTo 0
   Exit Sub

createRS_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRS of Form frmBillingMultiCetakUlang")
End Sub
 
 

Private Sub cmdCetak_Click()
    If Not rsGrid.EOF Then
        rsGrid.MoveFirst
        Do While Not rsGrid.EOF
        
            ''Call cetakInv1New(rsGrid!noINv, 1)
            Dim f As New frmBillingMultiCetakUlangCetak
            f.noInv = NulSID(rsGrid!noInv)
            f.show vbModal
            
            objO.InsertData ("UPDATE cetakMultipleInvoice SET tglcetak = GETDATE(),iscetak = 1 WHERE noinv = '" & NulSID(rsGrid!noInv) & "'")
            rsGrid.MoveNext
        Loop
        refreshGrid
    Else
        MsgBox "Tidak ada data yang ditampilkan.", vbInformation
    End If
End Sub

Private Sub cmdDisplay_Click()
    Call refreshGrid
End Sub
 
Private Sub cmdImport_Click()
          
1         On Error GoTo cmdImport_Click_Error
          Dim ErrBegin As Boolean, konter As Long
          Dim CnExel As New ADODB.Connection
          Dim RcExcel As New ADODB.Recordset
          Dim Jml As Integer
2         konter = 0
3         ProgBar.Min = 0
4         CommonDialog1.FileName = ""
5         CommonDialog1.InitDir = App.Path
6         CommonDialog1.filter = "Exel File (*.xls)|*.xls"
7         CommonDialog1.ShowOpen
8         If CommonDialog1.FileName <> "" Then
9             If CnExel.State Then CnExel.Close
10            CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
11            GoSub ProsesImport
12        Else
13            Exit Sub
14        End If
15        Exit Sub
          
ProsesImport:
16        ErrBegin = True
          Dim VarNmFile As String
          Dim job As String
          Dim namaModel As String
          Dim rsExs As ADODB.Recordset
17        Call objHT.OpenTransactionBengkel
          
18        If RcExcel.State Then RcExcel.Close
19        RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
20        If RcExcel.RecordCount > 0 Then
21            ProgBar.Visible = True
22            ProgBar.Max = RcExcel.RecordCount
23            RcExcel.MoveFirst
24            Do Until RcExcel.EOF
25                DoEvents
26                konter = konter + 1
27                ProgBar.value = konter
28                job = ""
29                namaModel = ""
                  
30                If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
                   
31                sSQL = "SELECT * FROM T_INVOICE WITH (NOLOCK) WHERE NOINV = '" & Trim(RcExcel.Fields(0).value) & "' and FS_LOG_DEL = ''"
32                Set rsTemp = objO.FetchData(sSQL)
33                If rsTemp.EOF Then
34                    Call objHT.CloseTransactionBengkel_err
35                    ProgBar.Visible = False
36                    MsgBox "NOINV " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " tidak ditemukan!", vbCritical, "Error"
37                    Exit Sub
                   
38                End If
                    
39                Call objHT.InsertData_Bengkel("INSERT INTO cetakMultipleInvoice (noinv,tglentry,tglcetak,iscetak) VALUES ('" & Trim(RcExcel.Fields(0).value) & "',GETDATE(),'1900-01-01','0')")
                       
41                RcExcel.MoveNext
42            Loop
43        End If
44        RcExcel.Close
45        CnExel.Close
          '   ProgBar.Visible = False
          '  ' IsiGrid
          '   DbCon.CommitTrans
          '   ErrBegin = False
46        Call objHT.CloseTransactionBengkel
47        MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
48        Call refreshGrid
49        ProgBar.Visible = False
50        On Error GoTo 0
51        Exit Sub

cmdImport_Click_Error:
52    ProgBar.Visible = False
53        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdImport_Click of Form frmBillingMultiCetakUlang")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub
 

Private Sub cmdSettingPrinter_Click()
Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()

    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub Form_Load()
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridFlate.Width = Frame1.Width - 200
    GridFlate.Height = Frame1.Height
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    cmdImport.Left = cmdCetak.Left - cmdImport.Width - 100
    cmdDisplay.Top = cmdSimpan.Top
End Sub

Private Sub GridFlate_FilterChange()
1        On Error GoTo GridFlate_FilterChange_Error

2         Set cols = GridFlate.Columns
          Dim C As Integer
          
3         C = GridFlate.col
4         GridFlate.HoldFields
5         rsGrid.filter = getFilter()
6         GridFlate.col = C
7         GridFlate.EditActive = True

8        On Error GoTo 0
9        Exit Sub

GridFlate_FilterChange_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GridFlate_FilterChange of Form frmBillingMultiCetakUlang")
End Sub

Private Function getFilter() As String
      Dim tmp As String
      Dim n As Integer
1        On Error GoTo getFilter_Error

2         For Each col In cols
3             If Trim(col.FilterText) <> "" Then
4                 n = n + 1
5                 If n > 1 Then
6                     tmp = tmp & " AND "
7                 End If
8                 tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
9             End If
10        Next col
11        getFilter = tmp

12       On Error GoTo 0
13       Exit Function

getFilter_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form frmBillingMultiCetakUlang")
End Function
 


Private Sub cetakInv1New(noInv As String, No As Integer)
    Dim rsRepo As ADODB.Recordset
    Dim rsAddJob As ADODB.Recordset
    Dim rsProfil As ADODB.Recordset
    Dim rsRinciSBE As ADODB.Recordset
    Dim rpt As New rtpNewInvoice
    Dim hlp1 As String
    Dim hlp2 As String
    Dim hlp3 As String
    Dim hlp4 As String
    Dim hlp6 As String
    Dim gcat As String
    Dim infoBill As String
    Dim no_wo As String
    Dim nilaiMAT As Double
    Dim tglInv  As String
    
    
    Dim onlySublet      As Boolean
    Dim isSBE           As Boolean
    Dim isFixPrice      As Boolean
    Dim discFixPrice    As Double
    Dim total           As Double
    
    If Log.notifInvNCM = 1 Then
        rpt.boxNCM.Suppress = True
        rpt.txtNCM1.Suppress = True
        rpt.txtNCM2.Suppress = True
        
        If isJualNRM = True Then
            Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NCM'")
            If Not rsTemp.EOF Then
            
                rpt.boxNCM.Suppress = False
                rpt.txtNCM1.Suppress = False
                rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                
                rpt.txtNCM2.Suppress = False
                rpt.txtNCM2.SetText Trim(rsTemp!Text2)
             
            End If
        Else
            Set rsTemp = objO.FetchData("SELECT * FROM notifInvoice WITH (NOLOCK) WHERE tipe = 'NonNCM'")
            If Not rsTemp.EOF Then
            
                rpt.boxNCM.Suppress = False
                rpt.txtNCM1.Suppress = False
                rpt.txtNCM1.SetText Trim(rsTemp!Text1)
                
                rpt.txtNCM2.Suppress = False
                rpt.txtNCM2.SetText Trim(rsTemp!Text2)
            End If
        End If
    End If
    nilaiMAT = 0
    createTableNewBilling
    sSQL = "select a.NOINV, a.NO_WO,a.FS_KD_PLG,a.CUSTOMER,a.CUST_ADRES,"
    sSQL = sSQL & " a.NO_POLISI,a.NO_CHASSIS,a.FMDL,a.NO_ENGINE,a.NO_WO,a.FRONT,A.STATUS,A.MECH,A.INVDATE,A.INVTIME,"
    sSQL = sSQL & " A.REP1,A.REP2,A.REP3,A.REP4,A.REP5,A.REP6,A.REP7,A.REP8,"
    sSQL = sSQL & " A.RTE1,A.RTE2,A.RTE3,A.RTE4,A.RTE5,A.RTE6,A.RTE7,A.RTE8,"
    sSQL = sSQL & " A.OFF1,A.OFF2,A.OFF3,A.OFF4,A.OFF5,A.OFF6,A.OFF7,A.OFF8,"
    sSQL = sSQL & " A.DISC1,A.DISC2,A.DISC3,A.DISC4,A.DISC5,A.DISC6,A.DISC7,A.DISC8,"
    sSQL = sSQL & " PART_G,PART_NG,MTRL,OIL_G,OIL_NG,MET,SBLT,LABOR,BODY,TOTAL-(FN_PAJAK+met) AS DPP,FN_PAJAK,TOTAL,FS_RC,DATE,PPn"
    sSQL = sSQL & " from T_INVOICE a with (nolock) "
    sSQL = sSQL & " where NOINV = '" & noInv & "'"
    sSQL = sSQL & " order by INVDATE desc"
    Set rsRepo = objO.FetchData(sSQL)
    
    sSQL = "SELECT A.* FROM HIST_JOBADD A WITH(NOLOCK) INNER JOIN T_INVOICE B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_INV = B.NOINV AND A.NO_CHASSIS = B.NO_CHASSIS AND A.FS_LOG_DEL = B.FS_LOG_DEL  WHERE   A.FS_LOG_DEL = '' AND B.NOINV LIKE '" & Trim(noInv) & "%'"
    Set rsAddJob = objO.FetchData(sSQL)
    
    sSQL = "select FS_NM_PERUSH,FS_ALM_PERUSH,'No Telpon : ' + FS_TELP + ', Fax : ' + FS_FAX AS FS_TELP,'Email : ' + FS_EMAIL as FS_EMAIL,'NPWP : ' + FS_NPWP + ' Tanggal : ' as FS_NPWP,'PKP : ' + FS_PKP as FS_PKP,bankname1,bankname2,bankname3,bankno1,bankno2,bankno3 from t_profil"
    Set rsProfil = objO.FetchData(sSQL)
    If rsProfil.EOF Then
        MsgBox "Gagal memuat data profil dealer", vbInformation
        Exit Sub
    End If
    If Not rsRepo.EOF Then

        nilaiPPnInv = NulSID(rsRepo("PPn"))
        nilaiDppInv = 100
        nilaiTotalInv = nilaiDppInv + nilaiPPnInv
        nilaiMAT = rsRepo!met
        rsRepo.MoveFirst
        isSBE = False
        no_wo = NulSID(rsRepo!no_wo)
        tglInv = TglDDMMMMYYYY(rsRepo!InvDate, 1) & " " & rsRepo!INVTIME
        If NulSID(rsRepo!FS_RC) <> "" Then
            isSBE = True
        End If
        onlySublet = True
        Do While Not rsRepo.EOF
            For i = 1 To 8
                hlp1 = "REP" & i
                hlp2 = "RTE" & i
                hlp3 = "OFF" & i
                hlp4 = "DISC" & i
                
                '' -------------------------------------------------
                '' 20200316 || cetakan invoice untuk fix price ditampilkan disc
                isFixPrice = False
                discFixPrice = 0
                If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(rsRepo!Date, "DD/MM/YYYY")) >= 0 Then
                    If objO.FetchData_String("select no_wo from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsRepo(hlp1)) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'") <> "" Then
                        isFixPrice = True
                        discFixPrice = objO.FetchData_Double("select disc from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsRepo(hlp1)) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'")
                    End If
                End If
                
                If Trim(rsRepo(hlp1)) <> "" Then
                    With rsCetakNewBilling
                        .AddNew
                        !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                        !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                        !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                        !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                        !alamatperusahaanpajak = NulSID("")
                        !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                        !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                        !noInv = NulSID(rsRepo!noInv)
                        !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                        !customer = NulSID(rsRepo!customer)
                        !customeralamat = NulSID(rsRepo!cust_adres)
                        !NoPolisi = NulSID(rsRepo!NO_POLISI)
                        !model = NulSID(rsRepo!Fmdl)
                        !noRangka = NulSID(rsRepo!NO_CHASSIS)
                        !noMesin = NulSID(rsRepo!No_Engine)
                        !noOrder = NulSID(rsRepo!no_wo)
                        !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                        !tipecustomer = NulSID(rsRepo!Status)
                        !mekanik = NulSID(rsRepo!Mech)
                        !billing = NulSID(rsRepo!FRONT) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                        !job = Left(IIf(InStr(NulSID(rsRepo(hlp1)), ".000") > 0, " ", "") & NulSID(rsRepo(hlp1)), 50)
                        !jobdetail = NulSID("")
                        !jobRate = k2t(NulSID(rsRepo(hlp2)))
                        !jobharga = NulSID(rsRepo(hlp3))
                        '!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
                        If isFixPrice = True Then
                            !jobdisc = NulSID(discFixPrice & " %")
                        Else
                            !jobdisc = NulSID(rsRepo(hlp4) & " %")
                        End If
                        !jobsubtotal = k2t(NulSID((rsRepo(hlp3) * (100 - NulSID(rsRepo(hlp4)))) / 100)) ''NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
                        !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                        !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                        !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                        !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                        !PPNtotal = NulA(rsRepo!fn_pajak)
                        !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                        .Update
                    End With
                    onlySublet = False
                    If isSBE Then
                        Dim nilaisbe As Integer
                        nilaisbe = CInt(Right(NulSID(rsRepo!FS_RC), 3))
                        If InStr(NulSID(rsRepo(hlp1)), nilaisbe & ".000") > 0 Then
                            Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe, String(3, "0")) & "' order by Status")
                            If Not rsRinciSBE.EOF Then
                                rsRinciSBE.MoveFirst
                                Do While Not rsRinciSBE.EOF
                                    With rsCetakNewBilling
                                        .AddNew
                                        !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                                        !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                                        !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                                        !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                                        !alamatperusahaanpajak = NulSID("")
                                        !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                                        !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                                        !noInv = NulSID(rsRepo!noInv)
                                        !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                                        !customer = NulSID(rsRepo!customer)
                                        !customeralamat = NulSID(rsRepo!cust_adres)
                                        !NoPolisi = NulSID(rsRepo!NO_POLISI)
                                        !model = NulSID(rsRepo!Fmdl)
                                        !noRangka = NulSID(rsRepo!NO_CHASSIS)
                                        !noMesin = NulSID(rsRepo!No_Engine)
                                        !noOrder = NulSID(rsRepo!no_wo)
                                        !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                                        !tipecustomer = NulSID(rsRepo!Status)
                                        !mekanik = NulSID(rsRepo!Mech)
                                        !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                                        !job = Left(" " & NulSID(rsRepo(hlp1)), 50)
                                        !jobdetail = NulSID(rsRinciSBE!Detail)
                                        !jobRate = k2t(NulSID(rsRepo(hlp2)))
                                        !jobharga = NulSID(rsRepo(hlp3))
                                        ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsRepo(hlp4) & " %"), "") 'NulSID(rsRepo(hlp4) & " %")
                                        If isFixPrice = True Then
                                            !jobdisc = NulSID(discFixPrice & " %")
                                        Else
                                            !jobdisc = NulSID(rsRepo(hlp4) & " %")
                                        End If
                                        !jobsubtotal = k2t(NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100)))
                                        !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                                        !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                                        !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                                        !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                                        !PPNtotal = NulA(rsRepo!fn_pajak)
                                        !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                                        .Update
                                    End With
                                    rsRinciSBE.MoveNext
                                Loop
                            End If
                        End If
                    End If
                End If
            Next
            
            If onlySublet Then
                With rsCetakNewBilling
                    .AddNew
                    !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                    !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                    !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                    !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                    !alamatperusahaanpajak = NulSID("")
                    !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                    !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                    !noInv = NulSID(rsRepo!noInv)
                    !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                    !customer = NulSID(rsRepo!customer)
                    !customeralamat = NulSID(rsRepo!cust_adres)
                    !NoPolisi = NulSID(rsRepo!NO_POLISI)
                    !model = NulSID(rsRepo!Fmdl)
                    !noRangka = NulSID(rsRepo!NO_CHASSIS)
                    !noMesin = NulSID(rsRepo!No_Engine)
                    !noOrder = NulSID(rsRepo!no_wo)
                    !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                    !tipecustomer = NulSID(rsRepo!Status)
                    !mekanik = NulSID(rsRepo!Mech)
                    !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                    !job = NulSID("Order Sublet")
                    !jobdetail = NulSID("")
                    !jobRate = 0 'NulSID(rsRepo(hlp2))
                    !jobharga = 0 'NulSID(rsRepo(hlp3))
                    !jobdisc = 0 'NulSID(rsRepo(hlp4) & " %")
                    !jobsubtotal = 0 ' NulSID(rsRepo(hlp3) * ((100 - NulSID(rsRepo(hlp4))) / 100))
                    !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                    !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                    !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                    !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                    !PPNtotal = NulA(rsRepo!fn_pajak)
                    !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                    .Update
                End With
            End If
            rsRepo.MoveNext
        Loop
        
        If Not rsAddJob.EOF Then
            rsAddJob.MoveFirst
            rsRepo.MoveFirst
            Do While Not rsAddJob.EOF
                '' -------------------------------------------------
                '' 20200316 || cetakan invoice untuk fix price ditampilkan disc
                isFixPrice = False
                discFixPrice = 0
                If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(rsRepo!Date, "DD/MM/YYYY")) >= 0 Then
                    If objO.FetchData_String("select no_wo from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsAddJob!rep) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'") <> "" Then
                        isFixPrice = True
                        discFixPrice = objO.FetchData_Double("select disc from t_invoicejobdetail where fs_log_del = '' and isFP = 1 and rep = '" & NulSID(rsAddJob!rep) & "' and no_wo = '" & NulSID(rsRepo!no_wo) & "'")
                    End If
                End If
                With rsCetakNewBilling
                    .AddNew
                    !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                    !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                    !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                    !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                    !alamatperusahaanpajak = NulSID("")
                    !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                    !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                    !noInv = NulSID(rsRepo!noInv)
                    !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                    !customer = NulSID(rsRepo!customer)
                    !customeralamat = NulSID(rsRepo!cust_adres)
                    !NoPolisi = NulSID(rsRepo!NO_POLISI)
                    !model = NulSID(rsRepo!Fmdl)
                    !noRangka = NulSID(rsRepo!NO_CHASSIS)
                    !noMesin = NulSID(rsRepo!No_Engine)
                    !noOrder = NulSID(rsRepo!no_wo)
                    !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                    !tipecustomer = NulSID(rsRepo!Status)
                    !mekanik = NulSID(rsRepo!Mech)
                    !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                    !job = Left("Add. " & NulSID(rsAddJob!rep), 50)
                    !jobdetail = NulSID("")
                    !jobRate = k2t(NulSID(rsAddJob!rte))
                    !jobharga = NulSID(rsAddJob!offn)
                    ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "") 'NulSID(rsAddJob!disc) & " %"
                    If isFixPrice = True Then
                        !jobdisc = IIf(chCtkDisc.value = False, NulSID(discFixPrice & " %"), "")
                    Else
                        !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "")
                    End If
                    !jobsubtotal = Round(NulSID(rsAddJob!offn * ((100 - NulSID(rsAddJob!disc)) / 100)))
                    !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''nilaiDppInv / nilaiTotalInv,0)  ''10 / 11, 0)
                    !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                    !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                    !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                    !PPNtotal = NulA(rsRepo!fn_pajak)
                    !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                    .Update
                End With
                If isSBE Then
                    Dim nilaisbe2 As Integer
                    nilaisbe2 = CInt(Right(NulSID(rsRepo!FS_RC), 3))
                    If InStr(NulSID(rsRepo(hlp1)), nilaisbe2 & ".000") > 0 Then
                        Set rsRinciSBE = objO.FetchData("select rinciansb as detail from rinciansbe where tipesb = 'R" & Format(nilaisbe2, String(3, "0")) & "' order by Status")
                        If Not rsRinciSBE.EOF Then
                            rsRinciSBE.MoveFirst
                            Do While Not rsRinciSBE.EOF
                                With rsCetakNewBilling
                                    .AddNew
                                    !namaperusahaan = NulSID(rsProfil!FS_NM_PERUSH)
                                    !alamatperusahaan = NulSID(rsProfil!FS_ALM_PERUSH)
                                    !teleponperusahaan = NulSID(rsProfil!FS_TELP)
                                    !emailperusahaan = NulSID(rsProfil!FS_EMAIL)
                                    !alamatperusahaanpajak = NulSID("")
                                    !npwpperusahaan = NulSID(rsProfil!FS_NPWP)
                                    !pkpperusahaan = NulSID(rsProfil!FS_PKP)
                                    !noInv = NulSID(rsRepo!noInv)
                                    !kodeLgn = objO.FetchData_String("SELECT idcust FROM [SINGLEIDORI].dbo.bridging WITH (NOLOCK) WHERE kodelgn = '" & NulSID(rsRepo!FS_KD_PLG) & "' AND kodedealer = '" & Right(gDb, 2) & "' ")
                                    !customer = NulSID(rsRepo!customer)
                                    !customeralamat = NulSID(rsRepo!cust_adres)
                                    !NoPolisi = NulSID(rsRepo!NO_POLISI)
                                    !model = NulSID(rsRepo!Fmdl)
                                    !noRangka = NulSID(rsRepo!NO_CHASSIS)
                                    !noMesin = NulSID(rsRepo!No_Engine)
                                    !noOrder = NulSID(rsRepo!no_wo)
                                    !namaSA = objO.FetchData_String("select FS_NM_PETUGAS from T_PETUGAS with (nolock) where FS_KD_PETUGAS = '" & NulSID(rsRepo!FRONT) & "'")
                                    !tipecustomer = NulSID(rsRepo!Status)
                                    !mekanik = NulSID(rsRepo!Mech)
                                    !billing = NulSID(rsRepo!FRONT) ''NulSID(rsRepo!InvDate) ''Format(NulSID(rsRepo!InvDate), "dd-mm-yyyy")
                                    !job = Left(" " & "Add. " & NulSID(rsAddJob!rep), 50)
                                    !jobdetail = NulSID(rsRinciSBE!Detail)
                                    !jobRate = k2t(NulSID(rsAddJob!rte))
                                    !jobharga = NulSID(rsAddJob!offn)
                                    ''!jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "") 'NulSID(rsAddJob!disc) & " %"
                                    If isFixPrice = True Then
                                        !jobdisc = IIf(chCtkDisc.value = False, NulSID(discFixPrice & " %"), "")
                                    Else
                                        !jobdisc = IIf(chCtkDisc.value = False, NulSID(rsAddJob!disc) & " %", "")
                                    End If
                                    !jobsubtotal = k2t(NulSID(rsAddJob!offn * ((100 - NulSID(rsAddJob!disc)) / 100)))
                                    !dppJasa = Round(NulA(rsRepo!labor) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                                    !dpppart = Round((NulA(rsRepo!PART_G) + NulA(rsRepo!PART_NG) + NulA(rsRepo!MTRL) + NulA(rsRepo!OIL_G) + NulA(rsRepo!OIL_NG)) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                                    !dppsublet = Round(NulA(rsRepo!sblt) * nilaiDppInv / nilaiTotalInv, 0) ''10 / 11, 0)
                                    !dpptotal = NulA(!dppJasa + !dpppart + !dppsublet) ''NulA(rsRepo!dpp)
                                    !PPNtotal = NulA(rsRepo!fn_pajak)
                                    !grandTotal = NulA(!dpptotal + !PPNtotal + nilaiMAT) ''NulA(rsRepo!total)
                                    .Update
                                End With
                                rsRinciSBE.MoveNext
                            Loop
                        End If
                    End If
                End If
                rsAddJob.MoveNext
            Loop
        End If
        gcat = ""
        If nRC Then
            Dim rsHlp As ADODB.Recordset
            Dim nBulan As String
            Dim nSbe As String
            nSbe = ""
            Set rsHlp = objO.FetchData("SELECT CASE WHEN JT_TGL < JT_BLN THEN JT_TGL ELSE JT_BLN END AS JT_TGL, SBENEXT FROM M_ESTIMASI WITH (NOLOCK) WHERE NORANGKA = '" & Trim(rsRepo!NO_CHASSIS) & "' AND TGLSVCNOW = '" & Format(rsRepo!Date, "MM/DD/YYYY") & "'")
            If Not rsHlp.EOF Then
                nBulan = GetNamaBulan(Month(rsHlp("JT_TGL"))) & " " & year(rsHlp("JT_TGL"))
                nSbe = Trim(rsHlp("SBENEXT"))
            End If
            rsHlp.Close
            If Trim(nSbe) <> "" Then
                'gcat = IIf(nHeavy, "Kembali untuk service berkala setelah kelipatan 5.000 KM atau bulan " & nBulan, "Kembali untuk service berkala setelah kelipatan 10.000 KM atau bulan " & nBulan)
                ''gcat = "Kembali untuk service berkala " & Format(CDbl(nSbe), "#,#0") & " KM atau bulan " & nBulan
                gcat = "Kembali untuk service berkala bulan " & nBulan & " atau " & Format(CDbl(nSbe), "#,#0") & " KM "
                gcat = gcat + IIf(gcat <> "", " (mana yang tercapai lebih dahulu).", "")
            Else
               '' Call objO.InsertData("INSERT INTO ESTIMASISALAH (norangka, tglservice, kmsvc, sbesvc, kelipatan, noinv, nowo,customer, alamat,nopolisi, telpon, gsm,cdma) VALUES " & _
                               "('" & Trim(Text(17)) & "','" & Format(Text(43), "MM/DD/YYYY") & "'," & CDbl(Text(37)) & ",'" & Trim(Text(33)) & "','" & IIf(nHeavy, 5000, 10000) & "', " & _
                               "'" & Trim(Text(20)) & "','" & Trim(Trim(Text(0))) & "','" & Trim(Text(4)) & "','" & Trim(txt5) & "','" & Trim(Text(16)) & "','" & Trim(Text(7)) & "','" & Trim(Text(8)) & "','" & Trim(Text(9)) & "')")
            End If
        End If
        
        rpt.txtTerbilang.SetText "Terbilang (Rupiah) " & Trim(Angka2Huruf(rsCetakNewBilling!grandTotal, 1))
        ''20240829 revisi menutup kembali service
        rpt.txtNextService.SetText "" 'gcat
        
        '' inspection ====================
        If Log.CR7 = 1 Then
            Dim rsInspection As ADODB.Recordset
            Dim total1 As Double
            Dim total2 As Double
            Dim total3 As Double
            Dim total4 As Double
            
            sSQL = "select sp1,sp2,sp3,sp4,estn1,estn2,estn3,estn4,estj1,estj2,estj3,estj4,estp1,estp2,estp3,estp4,ests1,ests2,ests3,ests4 from INSPECTION with (nolock) where NoWo = '" & no_wo & "' order by tglentry desc"
            Set rsInspection = objO.FetchData(sSQL)
            If Not rsInspection.EOF Then
                rsInspection.MoveFirst
                rpt.txtSaran1.SetText NulSID(rsInspection!sp1)
                rpt.txtSaran2.SetText NulSID(rsInspection!sp2)
                rpt.txtSaran3.SetText NulSID(rsInspection!sp3)
                rpt.txtSaran4.SetText NulSID(rsInspection!sp4)
                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtWaktuSVC1.SetText Format(NulSID(rsInspection!estn1), "dd-mm-yyyy")
                Else
                    rpt.txtWaktuSVC1.SetText ""
                End If
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtWaktuSVC2.SetText Format(NulSID(rsInspection!estn2), "dd-mm-yyyy")
                Else
                    rpt.txtWaktuSVC2.SetText ""
                End If
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtWaktuSVC3.SetText Format(NulSID(rsInspection!estn3), "dd-mm-yyyy")
                Else
                    rpt.txtWaktuSVC3.SetText ""
                End If
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtWaktuSVC4.SetText Format(NulSID(rsInspection!estn4), "dd-mm-yyyy")
                Else
                    rpt.txtWaktuSVC4.SetText ""
                End If
                                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtEstJasa1.SetText Format(NulSID(rsInspection!estj1), "#,0")
                Else
                    rpt.txtEstJasa1.SetText ""
                End If
                total1 = total1 + NulSID(rsInspection!estj1)
                
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtEstJasa2.SetText Format(NulSID(rsInspection!estj2), "#,0")
                Else
                    rpt.txtEstJasa2.SetText ""
                End If
                total2 = total2 + NulSID(rsInspection!estj2)
                
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtEstJasa3.SetText Format(NulSID(rsInspection!estj3), "#,0")
                Else
                    rpt.txtEstJasa3.SetText ""
                End If
                total3 = total3 + NulSID(rsInspection!estj3)
                
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtEstJasa4.SetText Format(NulSID(rsInspection!estj4), "#,0")
                Else
                    rpt.txtEstJasa4.SetText ""
                End If
                total4 = total4 + NulSID(rsInspection!estj4)
                                                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtEstPart1.SetText Format(NulSID(rsInspection!estp1), "#,0")
                Else
                    rpt.txtEstPart1.SetText ""
                End If
                total1 = total1 + NulSID(rsInspection!estp1)
                                                
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtEstPart2.SetText Format(NulSID(rsInspection!estp2), "#,0")
                Else
                    rpt.txtEstPart2.SetText ""
                End If
                total2 = total2 + NulSID(rsInspection!estp2)
                                                
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtEstPart3.SetText Format(NulSID(rsInspection!estp3), "#,0")
                Else
                    rpt.txtEstPart3.SetText ""
                End If
                total3 = total3 + NulSID(rsInspection!estp3)
                                                
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtEstPart4.SetText Format(NulSID(rsInspection!estp4), "#,0")
                Else
                    rpt.txtEstPart4.SetText ""
                End If
                total4 = total4 + NulSID(rsInspection!estp4)
                                
                                                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtEstSublet1.SetText Format(NulSID(rsInspection!ests1), "#,0")
                Else
                    rpt.txtEstSublet1.SetText ""
                End If
                total1 = total1 + NulSID(rsInspection!ests1)
                                                
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtEstSublet2.SetText Format(NulSID(rsInspection!ests2), "#,0")
                Else
                    rpt.txtEstSublet2.SetText ""
                End If
                total2 = total2 + NulSID(rsInspection!ests2)
                                                
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtEstSublet3.SetText Format(NulSID(rsInspection!ests3), "#,0")
                Else
                    rpt.txtEstSublet3.SetText ""
                End If
                total3 = total3 + NulSID(rsInspection!ests3)
                                                
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtEstSublet4.SetText Format(NulSID(rsInspection!ests4), "#,0")
                Else
                    rpt.txtEstSublet4.SetText ""
                End If
                total4 = total4 + NulSID(rsInspection!ests4)
                                
                If NulSID(rsInspection!sp1) <> "" Then
                    rpt.txtEstTotal1.SetText Format(NulSID(total1), "#,0")
                Else
                    rpt.txtEstTotal1.SetText ""
                End If
                If NulSID(rsInspection!sp2) <> "" Then
                    rpt.txtEstTotal2.SetText Format(NulSID(total2), "#,0")
                Else
                    rpt.txtEstTotal2.SetText ""
                End If
                If NulSID(rsInspection!sp3) <> "" Then
                    rpt.txtEstTotal3.SetText Format(NulSID(total3), "#,0")
                Else
                    rpt.txtEstTotal3.SetText ""
                End If
                If NulSID(rsInspection!sp4) <> "" Then
                    rpt.txtEstTotal4.SetText Format(NulSID(total4), "#,0")
                Else
                    rpt.txtEstTotal4.SetText ""
                End If
                
            End If
        End If
        '' akhir inspection ==============
                         
        rpt.txtBankNama1.SetText NulSID(rsProfil!bankname1)
        rpt.txtBankNama2.SetText NulSID(rsProfil!bankname2)
        rpt.txtBankNama3.SetText NulSID(rsProfil!bankname3)
        
        'rpt.txtBankNo1.SetText NulSID(rsProfil!bankno1)
        'rpt.txtBankNo2.SetText NulSID(rsProfil!bankno2)
        'rpt.txtBankNo3.SetText NulSID(rsProfil!bankno3)
        
        rpt.txtInfoMaterai.SetText ""
        rpt.txtLunasMaterai.SetText ""
        
        If Log.fakturMaterai = 1 Then
            Set rsTemp = objO.FetchData("SELECT * FROM t_invoiceMaterai WITH (NOLOCK) WHERE NOINV = '" & noInv & "'")
            If Not rsTemp.EOF Then
                rpt.txtInfoMaterai.SetText "Ijin Pembubuhan Materai No. " & NulSID(rsTemp!NOIJIN)
                rpt.txtLunasMaterai.SetText "BEA MATERAI LUNAS " & NulSID(rsTemp!nilai)
            End If
        End If
        rsCetakNewBilling.MoveFirst
        
        rpt.Text57.SetText NulSID(objO.FetchData_String("SELECT FS_KA_BENGKEL FROM T_PROFIL"))
        rpt.Text8.SetText NulSID(objO.FetchData_String("SELECT FS_KOTA FROM T_PROFIL"))
        
        rpt.txtMaterai.SetText IIf(nilaiMAT = 0, "0,00", Format(NulSID(nilaiMAT), "#,###.00")) ''Format(NulSID(nilaiMAT), "#,###.00")
        rpt.txtTglBilling.SetText tglInv
        rpt.Text41.SetText "PPn " & nilaiPPnInv & " %"
        rpt.Text55.SetText "PPn " & nilaiPPnInv & " %"
        rpt.database.SetDataSource rsCetakNewBilling
        If No = 1 Then
            With CRViewer1
                .ReportSource = rpt
                .ViewReport
                .Zoom 100
                .Top = 0
                .Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 20
                .Width = Frame1.Width + 100
                .Height = asxPowerBanner1.Height
                
                cmdTutup.Top = .Top + 430
                cmdTutup.Left = .Left
                cmdTutup.Visible = True
                cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 5
                cmdSettingPrinter.Visible = True
                cmdSettingPrinter.Top = cmdTutup.Top
                .Visible = True
            End With
            DoEvents
            cmdTutup.Visible = True
        
        End If
        Screen.MousePointer = vbDefault
        Exit Sub
        
    Else
        MsgBox "Data invoice tidak ditemukan", vbInformation
    End If
End Sub



Private Sub createTableNewBilling()
1        On Error GoTo createTableNewBilling_Error
2         If rsCetakNewBilling.State Then rsCetakNewBilling.Close
3         Set rsCetakNewBilling = Nothing
          
4         With rsCetakNewBilling
5             .Fields.Append "namaperusahaan", adVarChar, 200
6             .Fields.Append "alamatperusahaan", adVarChar, 200
7             .Fields.Append "teleponperusahaan", adVarChar, 200
8             .Fields.Append "emailperusahaan", adVarChar, 200
9             .Fields.Append "alamatperusahaanpajak", adVarChar, 200
10            .Fields.Append "npwpperusahaan", adVarChar, 200
11            .Fields.Append "pkpperusahaan", adVarChar, 200
12            .Fields.Append "noinv", adVarChar, 20
13            .Fields.Append "kodelgn", adVarChar, 30
14            .Fields.Append "customer", adVarChar, 100
15            .Fields.Append "customeralamat", adVarChar, 200
16            .Fields.Append "nopolisi", adVarChar, 12
17            .Fields.Append "model", adVarChar, 50
18            .Fields.Append "norangka", adVarChar, 25
19            .Fields.Append "nomesin", adVarChar, 15
20            .Fields.Append "noorder", adVarChar, 40
21            .Fields.Append "namasa", adVarChar, 40
22            .Fields.Append "tipecustomer", adVarChar, 40
23            .Fields.Append "mekanik", adVarChar, 40
24            .Fields.Append "printed", adVarChar, 40
              ''.Fields.Append "billing", adDate
              ''20191108 REQ MB THANIA BILLING DIBUAT INISIAL FRONT
25            .Fields.Append "billing", adVarChar, 10
26            .Fields.Append "job", adVarChar, 50
27            .Fields.Append "jobdetail", adVarChar, 100
28            .Fields.Append "jobrate", adDouble
29            .Fields.Append "jobharga", adDouble
30            .Fields.Append "jobdisc", adVarChar, 10
31            .Fields.Append "jobsubtotal", adDouble
32            .Fields.Append "dppJasa", adDouble
33            .Fields.Append "dppPart", adDouble
34            .Fields.Append "dppSublet", adDouble
35            .Fields.Append "dppTotal", adDouble
36            .Fields.Append "ppnTotal", adDouble
37            .Fields.Append "grandTotal", adDouble
38            .Open
39        End With

40       On Error GoTo 0
41       Exit Sub
createTableNewBilling_Error:
42        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createTableNewBilling of Form frmBilling"
End Sub
