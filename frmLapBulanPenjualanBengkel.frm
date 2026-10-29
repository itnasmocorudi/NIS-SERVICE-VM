VERSION 5.00
Object = "{698E14D0-8B82-11D1-8B57-00A0C98CD92B}#1.0#0"; "arviewer.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLapBulanPenjualanBengkel 
   Caption         =   "Laporan Bulanan Penjualan Bengkel (Detail)"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   8535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "&Export Text"
      Height          =   390
      Left            =   4155
      TabIndex        =   12
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblConvert 
      Caption         =   "&Simpan Text"
      Height          =   390
      Left            =   2265
      TabIndex        =   11
      Top             =   7065
      Width           =   1890
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmLapBulanPenjualanBengkel.frx":0000
      Left            =   1560
      List            =   "frmLapBulanPenjualanBengkel.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   735
   End
   Begin VB.CommandButton TblLayar 
      Caption         =   "Cetak &Layar"
      Height          =   390
      Left            =   6045
      TabIndex        =   3
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblCetak 
      Caption         =   "Cetak &Printer"
      Height          =   390
      Left            =   7935
      TabIndex        =   4
      Top             =   7065
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9825
      TabIndex        =   5
      Top             =   7065
      Width           =   1890
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   150
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7560
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13335
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Laporan Bulanan Penjualan Bengkel (Detail)"
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
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5490
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewerCtl.ARViewer ARViewer1 
      Height          =   3375
      Left            =   615
      TabIndex        =   13
      Top             =   3180
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   5953
      SectionData     =   "frmLapBulanPenjualanBengkel.frx":0020
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   885
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "s/d Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   510
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2835
      TabIndex        =   8
      Top             =   510
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
      Height          =   255
      Left            =   2835
      TabIndex        =   7
      Top             =   180
      Width           =   1425
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   180
      Width           =   960
   End
End
Attribute VB_Name = "frmLapBulanPenjualanBengkel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HapusForm()
    Text(0) = ">>->>->>>>"
End Sub

Private Sub cboStatus_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombolCombo(KeyCode)
End Sub

Private Sub Form_Load()
    Call HapusForm
End Sub

Private Sub Cetak(Layar As Byte)
    Dim rpt As New RptLapBulanPenjualanBengkel
    Dim vNamaTabel As String
    
    If Not (CekTgl(Text(0)) And CekTgl(Text(1))) And cboStatus.Text = "" Then
        MsgBox "Masukkan Status, Tgl Awal dan Akhir !"
    Else
        gTglTemp = Text(0) & " s/d " & Text(1)
        gKetTemp = "(" & cboStatus.Text & ")"
        gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
        
        vNamaTabel = "Bulan_DetilPenjualan" & DriveSerialNumber("c:")
        
        Call BuatTabelInv(vNamaTabel)
        
        sSQL = "insert into [" & vNamaTabel & "] select * from T_INVOICE " _
             & " where invdate>='" & TglYYYYMMDD(Text(0)) & "' and invdate<='" & TglYYYYMMDD(Text(1)) & "'" _
             & " union all " _
             & " select " & NmFieldRetur & " from T_INVOICE where convert(char(10),FD_RETUR,120) >='" & TglYYYYMMDD(Text(0)) & "'" _
             & " and convert(char(10),FD_RETUR,120) <='" & TglYYYYMMDD(Text(1)) & "'"
        oConn.Execute sSQL
        
        Select Case cboStatus.ListIndex
            Case 0
                sSQL = "SELECT P1.* FROM (SELECT FN_RETURN, INVDATE, count(NOINV) as JML_INV, sum(FN_JML_ORDER) as JML_ORD, sum(LABOR) as LABOR, sum(SBLT) as SBLT, sum(PART) as PART, sum(OIL) as OIL, sum(MTRL) as MTRL, sum(TOTAL-MET) as SUB_TTL," _
                     & " sum(MET) as MET, sum(TOTAL) as GTOTAL, sum(TOTAL-MET-FN_PAJAK) as NETTO, sum(FN_PAJAK) as PAJAK, SUM(CASE P.FS_KREDIT WHEN '1' THEN P.TOTAL ELSE 0 END) AS KREDIT," _
                     & " SUM(CASE P.FS_KREDIT WHEN '' THEN P.TOTAL ELSE 0 END) AS CASH, SUM(CASE P.KLPKERJA1 WHEN 'BOR' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'BOR' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'BOR' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'BOR' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'BOR' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'BOR' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'BOR' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'BOR' THEN P.OFF8 ELSE 0 END) AS JML_BOR, SUM(CASE P.KLPKERJA1 WHEN 'RTJ' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'RTJ' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'RTJ' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'RTJ' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'RTJ' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'RTJ' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'RTJ' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'RTJ' THEN P.OFF8 ELSE 0 END) AS JML_ELC, SUM(CASE P.KLPKERJA1 WHEN 'ERC' THEN P.OFF1 ELSE 0 END) +"
                
                sSQL = sSQL + " SUM(CASE P.KLPKERJA2 WHEN 'ERC' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'ERC' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'ERC' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'ERC' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'ERC' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'ERC' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'ERC' THEN P.OFF8 ELSE 0 END) AS JML_ERC, SUM(CASE P.KLPKERJA1 WHEN 'GRP' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'GRP' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'GRP' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'GRP' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'GRP' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'GRP' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'GRP' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'GRP' THEN P.OFF8 ELSE 0 END) AS JML_GRP, SUM(CASE P.KLPKERJA1 WHEN 'IRC' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'IRC' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'IRC' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'IRC' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'IRC' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'IRC' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'IRC' THEN P.OFF7 ELSE 0 END) +"
                
                sSQL = sSQL + " SUM(CASE P.KLPKERJA8 WHEN 'IRC' THEN P.OFF8 ELSE 0 END) AS JML_IRC, SUM(CASE P.KLPKERJA1 WHEN 'LUB' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'LUB' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'LUB' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'LUB' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'LUB' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'LUB' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'LUB' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'LUB' THEN P.OFF8 ELSE 0 END) AS JML_LUB, SUM(CASE P.KLPKERJA1 WHEN 'TWC' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'TWC' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'TWC' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'TWC' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'TWC' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'TWC' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'TWC' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'TWC' THEN P.OFF8 ELSE 0 END) AS JML_TWC, SUM(CASE P.KLPKERJA1 WHEN 'WAS' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'WAS' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'WAS' THEN P.OFF3 ELSE 0 END) +"
                
               sSQL = sSQL + " SUM(CASE P.KLPKERJA4 WHEN 'WAS' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'WAS' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'WAS' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'WAS' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'WAS' THEN P.OFF8 ELSE 0 END) AS JML_WAS, "
                     
                sSQL = sSQL + " SUM(CASE P.KLPKERJA1 WHEN 'PDS' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'PDS' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'PDS' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'PDS' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'PDS' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'PDS' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'PDS' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'PDS' THEN P.OFF8 ELSE 0 END) AS JML_PDI, " _

                sSQL = sSQL + " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%PDS%' THEN 1 ELSE 0 END) AS T_PDI, " _
                     
                sSQL = sSQL + " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%BOR%' THEN 1 ELSE 0 END) AS T_BOR, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%RTJ%' THEN 1 ELSE 0 END) AS T_ELC, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%ERC%' THEN 1 ELSE 0 END) AS T_ERC, " _
                     & " SUM(CASE P.KLPKERJA1 WHEN 'GRP' THEN 1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'GRP' THEN 1 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'GRP' THEN 1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'GRP' THEN 1 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'GRP' THEN 1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'GRP' THEN 1 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'GRP' THEN 1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'GRP' THEN 1 ELSE 0 END) AS T_GRP, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%IRC%' THEN 1 ELSE 0 END) AS T_IRC, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%LUB%' THEN 1 ELSE 0 END) AS T_LUB, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%TWC%' THEN 1 ELSE 0 END) AS T_TWC, "
                     
                sSQL = sSQL + " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%WAS%' THEN 1 ELSE 0 END) AS T_WAS" _
                     & " FROM [" & vNamaTabel & "] AS P" _
                     & " WHERE invdate>='" & TglYYYYMMDD(Text(0)) & "' and invdate<='" & TglYYYYMMDD(Text(1)) & "'" _
                     & " GROUP BY P.FN_RETURN, P.INVDATE) AS P1 ORDER BY FN_RETURN, INVDATE"
            Case Else
                sSQL = "SELECT P1.* FROM (SELECT FN_RETURN, INVDATE, count(NOINV) as JML_INV, sum(FN_JML_ORDER) as JML_ORD, sum(LABOR) as LABOR, sum(SBLT) as SBLT, sum(PART) as PART, sum(OIL) as OIL, sum(MTRL) as MTRL, sum(TOTAL-MET) as SUB_TTL," _
                     & " sum(MET) as MET, sum(TOTAL) as GTOTAL, sum(TOTAL-MET-FN_PAJAK) as NETTO, sum(FN_PAJAK) as PAJAK, SUM(CASE P.FS_KREDIT WHEN '1' THEN P.TOTAL ELSE 0 END) AS KREDIT," _
                     & " SUM(CASE P.FS_KREDIT WHEN '' THEN P.TOTAL ELSE 0 END) AS CASH, SUM(CASE P.KLPKERJA1 WHEN 'BOR' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'BOR' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'BOR' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'BOR' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'BOR' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'BOR' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'BOR' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'BOR' THEN P.OFF8 ELSE 0 END) AS JML_BOR, SUM(CASE P.KLPKERJA1 WHEN 'RTJ' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'RTJ' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'RTJ' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'RTJ' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'RTJ' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'RTJ' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'RTJ' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'RTJ' THEN P.OFF8 ELSE 0 END) AS JML_ELC, SUM(CASE P.KLPKERJA1 WHEN 'ERC' THEN P.OFF1 ELSE 0 END) +"
                
                sSQL = sSQL + " SUM(CASE P.KLPKERJA2 WHEN 'ERC' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'ERC' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'ERC' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'ERC' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'ERC' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'ERC' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'ERC' THEN P.OFF8 ELSE 0 END) AS JML_ERC, SUM(CASE P.KLPKERJA1 WHEN 'GRP' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'GRP' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'GRP' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'GRP' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'GRP' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'GRP' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'GRP' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'GRP' THEN P.OFF8 ELSE 0 END) AS JML_GRP, SUM(CASE P.KLPKERJA1 WHEN 'IRC' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'IRC' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'IRC' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'IRC' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'IRC' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'IRC' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'IRC' THEN P.OFF7 ELSE 0 END) +"
                
                sSQL = sSQL + " SUM(CASE P.KLPKERJA8 WHEN 'IRC' THEN P.OFF8 ELSE 0 END) AS JML_IRC, SUM(CASE P.KLPKERJA1 WHEN 'LUB' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'LUB' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'LUB' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'LUB' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'LUB' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'LUB' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'LUB' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'LUB' THEN P.OFF8 ELSE 0 END) AS JML_LUB, SUM(CASE P.KLPKERJA1 WHEN 'TWC' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'TWC' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'TWC' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'TWC' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'TWC' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'TWC' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'TWC' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'TWC' THEN P.OFF8 ELSE 0 END) AS JML_TWC, SUM(CASE P.KLPKERJA1 WHEN 'WAS' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'WAS' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'WAS' THEN P.OFF3 ELSE 0 END) +"
                
               sSQL = sSQL + " SUM(CASE P.KLPKERJA4 WHEN 'WAS' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'WAS' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'WAS' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'WAS' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'WAS' THEN P.OFF8 ELSE 0 END) AS JML_WAS, "
                     
                sSQL = sSQL + " SUM(CASE P.KLPKERJA1 WHEN 'PDS' THEN P.OFF1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'PDS' THEN P.OFF2 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'PDS' THEN P.OFF3 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'PDS' THEN P.OFF4 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'PDS' THEN P.OFF5 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'PDS' THEN P.OFF6 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'PDS' THEN P.OFF7 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'PDS' THEN P.OFF8 ELSE 0 END) AS JML_PDI, " _

                sSQL = sSQL + " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%PDS%' THEN 1 ELSE 0 END) AS T_PDI, " _
                     
                sSQL = sSQL + " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%BOR%' THEN 1 ELSE 0 END) AS T_BOR, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%RTJ%' THEN 1 ELSE 0 END) AS T_ELC, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%ERC%' THEN 1 ELSE 0 END) AS T_ERC, " _
                     & " SUM(CASE P.KLPKERJA1 WHEN 'GRP' THEN 1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA2 WHEN 'GRP' THEN 1 ELSE 0 END) + SUM(CASE P.KLPKERJA3 WHEN 'GRP' THEN 1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA4 WHEN 'GRP' THEN 1 ELSE 0 END) + SUM(CASE P.KLPKERJA5 WHEN 'GRP' THEN 1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA6 WHEN 'GRP' THEN 1 ELSE 0 END) + SUM(CASE P.KLPKERJA7 WHEN 'GRP' THEN 1 ELSE 0 END) +" _
                     & " SUM(CASE P.KLPKERJA8 WHEN 'GRP' THEN 1 ELSE 0 END) AS T_GRP, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%IRC%' THEN 1 ELSE 0 END) AS T_IRC, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%LUB%' THEN 1 ELSE 0 END) AS T_LUB, " _
                     & " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%TWC%' THEN 1 ELSE 0 END) AS T_TWC, "
                     
                sSQL = sSQL + " SUM(CASE WHEN P.KLPKERJA1 + P.KLPKERJA2 + P.KLPKERJA3 + P.KLPKERJA4 + P.KLPKERJA5 " _
                     & " + P.KLPKERJA6 + P.KLPKERJA7 + P.KLPKERJA8 LIKE '%WAS%' THEN 1 ELSE 0 END) AS T_WAS" _
                     & " FROM [" & vNamaTabel & "] AS P" _
                     & " WHERE invdate>='" & TglYYYYMMDD(Text(0)) & "' and invdate<='" & TglYYYYMMDD(Text(1)) & "' and STATUS='" & cboStatus.Text & "'" _
                     & " GROUP BY P.FN_RETURN, P.INVDATE) AS P1 ORDER BY FN_RETURN, INVDATE"
        End Select
        rpt.DataControl1.ConnectionString = gReportConn
        rpt.DataControl1.Source = sSQL
'        rpt.Printer.PaperSize = 14
'        rpt.Printer.Orientation = 2
        rpt.Printer.PaperSize = 39
        rpt.Printer.Orientation = ddOPortrait
        Select Case Layar
            Case 1
                rpt.Show
            Case 2
                ARViewer1.ReportSource = rpt
            Case Else
                rpt.PrintReport False
                Unload rpt
        End Select
    End If
End Sub

Private Sub TblCetak_Click()
    Call Cetak(0)
End Sub

Private Sub TblConvert_Click()
    Dim txt As New ActiveReportsTextExport.ARExportText

    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.ShowSave
    
    txt.FileName = CommonDialog1.FileName
    txt.Export ARViewer1.Pages
End Sub

Private Sub Command2_Click()
    Call Cetak(2)
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblLayar_Click()
    Call Cetak(1)
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

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0, 1
            If CekTgl(Text(Index)) <> True Then
                Text(Index) = gTanggal
            End If
            If TglYYYYMMDD(Text(0)) > TglYYYYMMDD(Text(1)) Then
                MsgBox "Tanggal awal lebih besar dari tanggal akhir !"
            End If
    End Select
End Sub

