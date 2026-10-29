VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmMonitoringOTD 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Monitoring OTD "
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   12165
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10610
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
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
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.namedParent=39"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.namedParent=40"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
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
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&H6F9DFF&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
      _StyleDefs(61)  =   "Named:id=29:KuningHijau"
      _StyleDefs(62)  =   ":id=29,.parent=42,.bgcolor=&HFFFF&,.fgcolor=&H60BF&,.bold=-1,.fontsize=825"
      _StyleDefs(63)  =   ":id=29,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(64)  =   ":id=29,.fontname=MS Sans Serif"
      _StyleDefs(65)  =   "Named:id=30:PutihHijau"
      _StyleDefs(66)  =   ":id=30,.parent=29,.bgcolor=&HFFFFFF&,.fgcolor=&H60BF&"
      _StyleDefs(67)  =   "Named:id=31:MerahHijau"
      _StyleDefs(68)  =   ":id=31,.parent=30,.bgcolor=&HFF&,.fgcolor=&H60BF&"
      _StyleDefs(69)  =   "Named:id=32:KuningHitam"
      _StyleDefs(70)  =   ":id=32,.parent=29,.fgcolor=&H646464&,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(71)  =   ":id=32,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(72)  =   ":id=32,.fontname=MS Sans Serif"
      _StyleDefs(73)  =   "Named:id=43:PutihHitam"
      _StyleDefs(74)  =   ":id=43,.parent=30,.fgcolor=&H646464&,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(75)  =   ":id=43,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(76)  =   ":id=43,.fontname=MS Sans Serif"
      _StyleDefs(77)  =   "Named:id=44:MerahHitam"
      _StyleDefs(78)  =   ":id=44,.parent=31,.fgcolor=&H646464&,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(79)  =   ":id=44,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(80)  =   ":id=44,.fontname=MS Sans Serif"
   End
End
Attribute VB_Name = "frmMonitoringOTD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As New ADODB.Recordset
Dim rsTemp2 As New ADODB.Recordset
Dim rsTemp3 As New ADODB.Recordset

Dim rsGrid As New ADODB.Recordset
Dim nCount As Integer
Dim cols As TrueOleDBGrid70.Columns
Dim tglAwal As String
Dim tglAkhir As String
Public monitoringSA As String
Public tanggalMonitoring As String



Private Sub headColumns()
'[DATE], CUSTOMER, CUST_ADRES, TELPFU, JAMFU, KETFU, A.MDL + ' | ' + A.FMDL, A.COLOR, A.NO_CHASSIS, A.NO_WO, A.NO_POLISI, KM

   On Error GoTo headColumns_Error

    With Grid
        .Columns(0).Caption = "NO WO"
        .Columns(0).DataField = "NO_WO"
        .Columns(0).Width = 1200
        .Columns(1).Caption = "NO POLISI"
        .Columns(1).DataField = "NO_POLISI"
        .Columns(1).Width = 1200
        .Columns(2).Caption = "MECH START"
        .Columns(2).DataField = "MEK"
        .Columns(2).Width = 2000
        
        .Columns(3).Caption = "LAST PROSES"
        .Columns(3).DataField = "posisi"
        .Columns(3).Width = 2500
        
        .Columns(4).Caption = "OTD"
        .Columns(4).DataField = "otd"
        .Columns(4).Width = 1200
        
        .Columns(5).Caption = "PROMISE"
        .Columns(5).DataField = "rubah"
        .Columns(5).Width = 1500
        
        .Columns(6).Visible = False
        .Columns(7).Visible = False
        .Columns(8).Visible = False
        .Columns(9).Visible = False
    End With

   On Error GoTo 0
   Exit Sub

headColumns_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumns of Form frm3DFU")
End Sub

Private Sub refreshGrid()
1        On Error GoTo refreshGrid_Error

2         Grid.ClearFields
3         If rsGrid.State Then rsGrid.Close
      '    sSQL = "select l.NO_WO,a.NO_POLISI ,"
      '    sSQL = sSQL & " (select top 1 TANGGAL + START from tempProgress with (nolock) where NOORDER = l.NO_WO order by TANGGAL,START) MEK,"
      '    sSQL = sSQL & " TIMESERAH, (select top 1 rubah from JOBSerah where No_Wo = l.NO_WO order by rubah desc) rubah,"
      '    sSQL = sSQL & " isnull((select top 1 id from [Queue" & Right(gDb, 2) & "]..Panggilan where NoPol = a.NO_POLISI),0) panggilan, l.timeserah otd, "
      '    sSQL = sSQL & " DATEDIFF(MINUTE,GETDATE(), l.dateserah + l.timeserah) selisihOTD, l.dateserah + l.timeserah janjiserah"
      '    sSQL = sSQL & " , case s.flag when 'N' then 'Unit belum dikerjakan' when 'P' then 'Unit dihentikan' when 'S' then 'Unit sedang diproses' when 'G' then 'Unit Antri Cuci' when 'H' then 'Waiting Inspection' when 'A' then 'Sedang Proses Cuci' when 'D' then 'Tunggu Proses Drying' when 'R' then 'Sedang Proses Drying' when 'Q' then 'Proses Inspection' when 'F' then 'Selesai Dikerjakan' else '' end posisi "
      '    sSQL = sSQL & " from JOBDTL l with (nolock)"
      '    sSQL = sSQL & " inner join JOBDTA a with (nolock) on l.NO_WO = a.NO_WO and l.NO_CHASSIS = a.NO_CHASSIS"
      '    sSQL = sSQL & " left join Stall s with (nolock) on REPLACE(SUBSTRING(s.KETERANGAN, CHARINDEX('|',s.KETERANGAN)+1,8),'|','') = A.NO_WO "
      '    sSQL = sSQL & " where FRONT = '" & monitoringSA & "' and DATESERAH = '" & Format(tanggalMonitoring, "mm/dd/yyyy") & "' and a.DELIV <> 'BTL' and a.RPR_COND <> 'K'"
          
4         sSQL = "select distinct l.NO_WO,a.NO_POLISI ,"
5         sSQL = sSQL & " (select top 1 TANGGAL + START from tempProgress with (nolock) where NOORDER = l.NO_WO order by TANGGAL,START) MEK,"
6         sSQL = sSQL & " TIMESERAH, (select top 1 rubah from JOBSerah where No_Wo = l.NO_WO order by rubah desc) rubah,"
7         sSQL = sSQL & " isnull((select top 1 1 from [Queue" & Right(gDb, 2) & "]..FinishService where TglFinish = t.INVDATE and JamFinish = t.INVTIME and NoPol = a.NO_POLISI),0) panggilan, l.timeserah otd, "
8         sSQL = sSQL & " case when ISNULL(t.NOINV,'') <> '' then isnull((select top 1 DATEDIFF(MINUTE,TglFinish + JamFinish, l.dateserah + l.timeserah) from [Queue" & Right(gDb, 2) & "]..FinishService where TglFinish = t.INVDATE and JamFinish = t.INVTIME and NoPol = a.NO_POLISI),0) else DATEDIFF(MINUTE,GETDATE(), l.dateserah + l.timeserah) end selisihOTD, l.dateserah + l.timeserah janjiserah"
9         sSQL = sSQL & " , case s.flag when 'N' then 'Unit belum dikerjakan' when 'P' then 'Unit dihentikan' when 'S' then 'Unit sedang diproses' when 'G' then 'Unit Antri Cuci' when 'H' then 'Waiting Inspection' when 'A' then 'Sedang Proses Cuci' when 'D' then 'Tunggu Proses Drying' when 'R' then 'Sedang Proses Drying' when 'Q' then 'Proses Inspection' when 'F' then 'Selesai Dikerjakan' else '' end posisi "
10        sSQL = sSQL & " from JOBDTL l with (nolock)"
11        sSQL = sSQL & " inner join JOBDTA a with (nolock) on l.NO_WO = a.NO_WO and l.NO_CHASSIS = a.NO_CHASSIS"
12        sSQL = sSQL & " left join T_INVOICE t with (nolock) on a.NO_WO = t.NO_WO and a.NO_CHASSIS = t.NO_CHASSIS and t.DELIV <> 'btl'"
13        sSQL = sSQL & " left join Stall s with (nolock) on REPLACE(SUBSTRING(s.KETERANGAN, CHARINDEX('|',s.KETERANGAN)+1,8),'|','') = A.NO_WO "
14        sSQL = sSQL & " where a.FRONT = '" & monitoringSA & "' and DATESERAH = '" & Format(tanggalMonitoring, "mm/dd/yyyy") & "' and a.DELIV <> 'BTL' and a.RPR_COND <> 'K'"
          
          
          
      '    Set rsGrid = objO.FetchData("SELECT [DATE], C.usename, C.useaddr, TELPFU, JAMFU, KETFU, ANCUST, TELPANCUST, A.MDL + ' | ' + A.FMDL AS MODEL, A.COLOR, A.NO_CHASSIS, A.NO_WO, A.NO_POLISI, KM, C.TELPON+' , '+C.GSM+' , '+C.CDMA+' , '+C.USEPSTN+' , '+C.USEGSM+' , '+C.USECDMA AS TELPON FROM JOBDTA A WITH (NOLOCK) " '                                    & "INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO " '                                    & "LEFT JOIN CUSTDAT C WITH (NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS AND A.NO_POLISI = C.NO_POLISI " '                                    & "WHERE TGLFU <= '" & Format(Tgl, "MM/DD/YYYY") & "' AND RPR_COND <> '' AND DELIV <> 'BTL' AND FRONT = '" & gSA & "' AND YEAR(TGLFU) > 1900 " '                                    & "AND A.NO_WO NOT IN (SELECT DISTINCT NO_WO FROM DTLFU WITH (NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS AND NO_WO = A.NO_WO AND (SAMBUNG = 'Y' OR COUNTER = 3) ) AND A.FS_LOG_DEL = '' ORDER BY TGLFU DESC, JAMFU ASC")

15       Set rsGrid = objO.FetchData(sSQL)

16        If Not rsGrid.EOF Then
             ' jmlBlmFollow.Caption = "Unit Belum Follow Up : " & rsGrid.RecordCount
17            Grid.DataSource = rsGrid
18            Grid.FetchRowStyle = True
19            Call headColumns
           
20        End If
       

21       On Error GoTo 0
22       Exit Sub

refreshGrid_Error:

23        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form frm3DFU")

End Sub


Private Sub Form_Load()
   If monitoringSA = "" Then
      monitoringSA = gSA
   End If
   If tanggalMonitoring = "" Then
      tanggalMonitoring = Format(isnow, "dd/mm/yyyy")
   End If
   
   Me.Caption = "Monitoring OTD : " & monitoringSA & " Tanggal : " & tanggalMonitoring
   Call refreshGrid
End Sub

Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
   Dim rs As New ADODB.Recordset
    Set rs = rsGrid.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    'RowStyle = "putih"
    If rs("Panggilan") > 0 Then
         ''20160722 font coklat bold
         ''background warna
         ''0-30 kuning
         ''> 30 putih
         ''< 0 merah
         ''jika sudah dipanggil -: background hijau
         If Trim(CDbl(rs!SELISIHOTD)) >= 0 And Trim(CDbl(rs!SELISIHOTD)) <= 30 Then
            RowStyle = "KuningHijau"
         ElseIf Trim(CDbl(rs!SELISIHOTD)) > 30 Then
            RowStyle = "PutihHijau"
         Else
            RowStyle = "MerahHijau"
         End If
    Else
         ''fonthitam
         ''jika belum dipanggil -: background putih
         If Trim(CDbl(rs!SELISIHOTD)) >= 0 And Trim(CDbl(rs!SELISIHOTD)) <= 30 Then
            RowStyle = "KuningHitam"
         ElseIf Trim(CDbl(rs!SELISIHOTD)) > 30 Then
            RowStyle = "PutihHitam"
         Else
            RowStyle = "MerahHitam"
         End If
    End If
End Sub
