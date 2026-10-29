VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form FormKPIHarian 
   BackColor       =   &H00BFE8EA&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   13380
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1050
   End
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
      Height          =   300
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1395
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   7095
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
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali Ke Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin TDBDate6Ctl.TDBDate DtDOAwal 
      Height          =   330
      Left            =   1905
      TabIndex        =   0
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "FormKPIHarian.frx":0000
      Caption         =   "FormKPIHarian.frx":012C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormKPIHarian.frx":0198
      Keys            =   "FormKPIHarian.frx":01B6
      Spin            =   "FormKPIHarian.frx":0214
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   2118189057
      Value           =   2.12482692446619E-314
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate DtDOAkhir 
      Height          =   330
      Left            =   3945
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "FormKPIHarian.frx":023C
      Caption         =   "FormKPIHarian.frx":0368
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormKPIHarian.frx":03D4
      Keys            =   "FormKPIHarian.frx":03F2
      Spin            =   "FormKPIHarian.frx":0450
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   2118189057
      Value           =   2.12482692446619E-314
      CenturyMode     =   0
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   0
      Picture         =   "FormKPIHarian.frx":0478
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KPI Harian"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   165
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   2880
      X2              =   5880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   5880
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   8
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3465
      TabIndex        =   7
      Top             =   645
      Width           =   345
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "FormKPIHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 01 Oktober 2010
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsRepo As New ADODB.Recordset
Dim boleh As Boolean
Dim vNamaTabel As String
Dim rpt As Object

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
Dim pesan As String
Dim temp As String
Dim nI As Integer
Dim namaTanggal As String
Dim nKPIa As Integer
Dim nKPIb As Integer
Dim nKPIc As Integer
Dim tKPIa(13) As Integer
Dim tKPIb(13) As Integer
Dim tKPIc(13) As Integer
Dim nHasil As String
Dim nUrut As Integer
Dim nKeterangan As String
Dim rsTemp As ADODB.Recordset

    
    'On Error Resume Next
    boleh = True
    pesan = ""
    If DateDiff("d", Format(DtDOAwal, "DD/MM/YYYY"), Format(DtDOAkhir, "DD/MM/YYYY")) > 31 Or DateDiff("d", Format(DtDOAwal, "DD/MM/YYYY"), Format(DtDOAkhir, "DD/MM/YYYY")) < 0 Then
        pesan = pesan + "Range Tanggal tidak boleh lebih dari 31 Hari. "
        boleh = False
    End If
    
    If boleh Then
        sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
             & " drop table [tempdb]..[" & vNamaTabel & "]"
        Call objO.UpdateData(sSQL)
        
        temp = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ([URUT][NUMERIC](9) NOT NULL DEFAULT 0, [KETERANGAN][varchar](100) NOT NULL DEFAULT '', [BOLD][char](1) NOT NULL DEFAULT '0' "
        For i = 0 To DateDiff("d", Format(DtDOAwal, "DD/MM/YYYY"), Format(DtDOAkhir, "DD/MM/YYYY"))
            temp = temp + ", [" & "A" & Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY"))) & "][varchar](10) NOT NULL DEFAULT '0'"
        Next
        temp = temp + ", [TOTAL][varchar](10) NOT NULL DEFAULT '0') ON [PRIMARY]"
        Call objO.UpdateData(temp)
        
        For i = 1 To 12
            tKPIa(i) = 0
            tKPIb(i) = 0
        Next
        For i = 0 To DateDiff("d", Format(DtDOAwal, "DD/MM/YYYY"), Format(DtDOAkhir, "DD/MM/YYYY"))
            Set rsTemp = objO.FetchData("EXEC KPIHARIAN '" & Format(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")), "MM/DD/YYYY") & "'")
            If Not rsTemp.EOF Then
                nI = 1
                nKPIa = 0
                nKPIb = 0
                nKPIc = 0
                namaTanggal = "A" & Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
                Do While nI <= 45 '45
                    Select Case nI
                        ' ************************************************* '
                        ''' GROUP KPI 1
                        ' ************************************************* '
                        Case 1
                            nUrut = 1
                            nKeterangan = "SALES DATABASE"
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2, True)
                            End If
                        Case 2
                            nUrut = 3
                            nKeterangan = "     Jumlah data pelanggan yang unitnya telah delivery (Pd)"
                            rsTemp.Find "KETERANGAN = 'KPI_1A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(1) = tKPIa(1) + nKPIa
                        Case 3
                            nUrut = 4
                            nKeterangan = "     Jumlah data pelanggan valid bulan ini dari yang unitnya telah delivery (Pv)"
                            rsTemp.Find "KETERANGAN = 'KPI_1B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(1) = tKPIb(1) + nKPIb
                        Case 4
                            nUrut = 2
                            nKeterangan = "1. KPI 1 : Database Valid  (=Pv/Pd)"
                            If nKPIa <> 0 Then
                                nHasil = (Round(CDbl(nKPIb / nKPIa) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        Case 5
                            nUrut = 5
                            nKeterangan = "     Jumlah carry over pelanggan valid dari yg unitnya telah delivery"
                            rsTemp.Find "KETERANGAN = 'KPI_1C'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIc = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIc = 0
                            End If
                            tKPIc(1) = tKPIc(1) + nKPIc
                        ' ************************************************* '
                        ''' GROUP KPI 2
                        ' ************************************************* '
                        Case 6
                            nUrut = 7
                            nKeterangan = "     Jumlah data pelanggan yang unitnya telah delivery (Pd)"
                            rsTemp.Find "KETERANGAN = 'KPI_2A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(2) = tKPIa(2) + nKPIa
                        Case 7
                            nUrut = 8
                            nKeterangan = "     Jumlah data pelanggan yang unitnya telah delivery ditransfer ke after sales (Pt)"
                            rsTemp.Find "KETERANGAN = 'KPI_2B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(2) = tKPIb(2) + nKPIb
                        Case 8
                            nUrut = 6
                            nKeterangan = "2. KPI 2 : Database Transfer (=Pt/Pd)"
                            If nKPIa <> 0 Then
                                nHasil = (Round(CDbl(nKPIb / nKPIa) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        Case 9
                            nUrut = 9
                            nKeterangan = "     Jumlah carry over data pelanggan yang unitnya telah delivery ditransfer ke after sales"
                            rsTemp.Find "KETERANGAN = 'KPI_2C'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIc = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIc = 0
                            End If
                            tKPIc(2) = tKPIc(2) + nKPIc
                        ' ************************************************* '
                        ''' GROUP KPI 3 SALES
                        ' ************************************************* '
                        Case 10
                            nUrut = 10
                            nKeterangan = "1K SERVICE INVITATION / CR2"
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2, True)
                            End If
                        Case 11
                            nUrut = 12
                            nKeterangan = "     Jumlah pelanggan yang seharusnya diingatkan servis 1.000 (Pc)"
                            rsTemp.Find "KETERANGAN = 'KPI_31A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(3) = tKPIa(3) + nKPIa
                        Case 12
                            nUrut = 13
                            nKeterangan = "     Jumlah pelanggan yang aktual diingatkan servis 1.000 (Pa)"
                            rsTemp.Find "KETERANGAN = 'KPI_31B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(3) = tKPIb(3) + nKPIb
                        Case 13
                            nUrut = 11
                            nKeterangan = "1. KPI 3 : 1K Sales (=Pa/Pc)"
                            If nKPIa <> 0 Then
                                nHasil = (Round(CDbl(nKPIb / nKPIa) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 4
                        ' ************************************************* '
                        Case 14
                            nUrut = 14
                            nKeterangan = "MRS PROCESS"
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2, True)
                            End If
                        Case 15
                            nUrut = 16
                            nKeterangan = "     Jumlah Direct Mail Diterima Pelanggan (Dt)"
                            rsTemp.Find "KETERANGAN = 'KPI_4A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(4) = tKPIa(4) + nKPIa
                        Case 16
                            nUrut = 17
                            nKeterangan = "     Jumlah Direct Mail Dikirim (Dk)"
                            rsTemp.Find "KETERANGAN = 'KPI_4B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(4) = tKPIb(4) + nKPIb
                        Case 17
                            nUrut = 15
                            nKeterangan = "1. KPI 4 : Akurasi Alamat Pelanggan (=Dt/Dk)"
                            If nKPIb <> 0 Then
                                nHasil = (Round(CDbl(nKPIa / nKPIb) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 5
                        ' ************************************************* '
                        Case 18
                            nUrut = 19
                            nKeterangan = "     Jumlah Telepon Terhubung (1) (Tt)"
                            rsTemp.Find "KETERANGAN = 'KPI_5A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(5) = tKPIa(5) + nKPIa
                        Case 19
                            nUrut = 20
                            nKeterangan = "     Jumlah Pelanggan Yang Ditelepon (Tk)"
                            rsTemp.Find "KETERANGAN = 'KPI_5B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(5) = tKPIb(5) + nKPIb
                        Case 20
                            nUrut = 18
                            nKeterangan = "2. KPI 5 : Akurasi Nomor Telepon Pelanggan (=Tt/Tk)"
                            If nKPIb <> 0 Then
                                nHasil = (Round(CDbl(nKPIa / nKPIb) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 6
                        ' ************************************************* '
                        Case 21
                            nUrut = 22
                            nKeterangan = "     Jumlah Pelanggan Yang Harus Diundang (P1=Aktif + Pasif + Baru)"
                            rsTemp.Find "KETERANGAN = 'KPI_6A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(6) = tKPIa(6) + nKPIa
                        Case 22
                            nUrut = 23
                            nKeterangan = "     Jumlah Pelanggan Yang Diundang (P2= DM + SMS + E + Tlp)"
                            rsTemp.Find "KETERANGAN = 'KPI_6B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(6) = tKPIb(6) + nKPIb
                        Case 23
                            nUrut = 21
                            nKeterangan = "3. KPI 6 : Persentase Pelanggan Yang Diundang (=P2/P1)"
                            If nKPIa <> 0 Then
                                nHasil = (Round(CDbl(nKPIb / nKPIa) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        
                        ' ************************************************* '
                        ''' GROUP KPI 6 MRS
                        ' ************************************************* '
                        Case 24
                            nUrut = 25
                            nKeterangan = "     Jumlah pelanggan yang seharusnya diundang servis 1.000 (Pc)"
                            rsTemp.Find "KETERANGAN = 'KPI_62A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(7) = tKPIa(7) + nKPIa
                        Case 25
                            nUrut = 26
                            nKeterangan = "     Jumlah pelanggan yang aktual diundang servis 1.000 (Pa)"
                            rsTemp.Find "KETERANGAN = 'KPI_62B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(7) = tKPIb(7) + nKPIb
                        Case 26
                            nUrut = 24
                            nKeterangan = "4. KPI 6 : 1K Activity (=Pa/Pc)"
                            If nKPIa <> 0 Then
                                nHasil = (Round(CDbl(nKPIb / nKPIa) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 7
                        ' ************************************************* '
                        Case 27
                            nUrut = 28
                            nKeterangan = "     Jumlah Pelanggan Yang Ditelepon (Tk)"
                            rsTemp.Find "KETERANGAN = 'KPI_7A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(8) = tKPIa(8) + nKPIa
                        Case 28
                            nUrut = 29
                            nKeterangan = "     Jumlah Pelanggan Yang Diundang (P2)"
                            rsTemp.Find "KETERANGAN = 'KPI_7B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(8) = tKPIb(8) + nKPIb
                        Case 29
                            nUrut = 27
                            nKeterangan = "5. KPI 7 : Persentase Pelanggan Yang Ditelepon (=Tk/P2)"
                            If nKPIb <> 0 Then
                                nHasil = (Round(CDbl(nKPIa / nKPIb) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 8
                        ' ************************************************* '
                        Case 30
                            nUrut = 31
                            nKeterangan = "     Jumlah Pelanggan Yang Appoinment (2) (A1)"
                            rsTemp.Find "KETERANGAN = 'KPI_8A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(9) = tKPIa(9) + nKPIa
                        Case 31
                            nUrut = 32
                            nKeterangan = "     Jumlah Telepon Terhubung (Tt)"
                            rsTemp.Find "KETERANGAN = 'KPI_8B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(9) = tKPIb(9) + nKPIb
                        Case 32
                            nUrut = 30
                            nKeterangan = "6. KPI 8 : Persentase Pelanggan Yang Appoinment (=A1/Tt)"
                            If nKPIb <> 0 Then
                                nHasil = (Round(CDbl(nKPIa / nKPIb) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 9
                        ' ************************************************* '
                        Case 33
                            nUrut = 34
                            nKeterangan = "     Jumlah Pelanggan Yang Datang Setelah Diundang (3) (Pd)"
                            rsTemp.Find "KETERANGAN = 'KPI_9A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(10) = tKPIa(10) + nKPIa
                        Case 34
                            nUrut = 35
                            nKeterangan = "     Jumlah Pelanggan Yang Diundang (P2)"
                            rsTemp.Find "KETERANGAN = 'KPI_9B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(10) = tKPIb(10) + nKPIb
                        Case 35
                            nUrut = 33
                            nKeterangan = "7. KPI 9 : Persentase Pelanggan Yang Datang Setelah Diundang (=Pd/P2)"
                            If nKPIb <> 0 Then
                                nHasil = (Round(CDbl(nKPIa / nKPIb) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 10
                        ' ************************************************* '
                        Case 36
                            nUrut = 36
                            nKeterangan = "MRS RESULT"
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2, True)
                            End If
                        Case 37
                            nUrut = 38
                            nKeterangan = "     Jumlah app. Karena aktifitas MRS (A1)"
                            rsTemp.Find "KETERANGAN = 'KPI_10A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(11) = tKPIa(11) + nKPIa
                        Case 38
                            nUrut = 39
                            nKeterangan = "     Jumlah total app. (A2)"
                            rsTemp.Find "KETERANGAN = 'KPI_10B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(11) = tKPIb(11) + nKPIb
                        Case 39
                            nUrut = 37
                            nKeterangan = "1. KPI 10 : Kontribusi MRS terhadap appointment. (=A1/A2)"
                            If nKPIb <> 0 Then
                                nHasil = (Round(CDbl(nKPIa / nKPIb) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 11
                        ' ************************************************* '
                        Case 40
                            nUrut = 41
                            nKeterangan = "     Jumlah unit SBE karena aktifitas MRS (Pd)"
                            rsTemp.Find "KETERANGAN = 'KPI_11A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(12) = tKPIa(12) + nKPIa
                        Case 41
                            nUrut = 42
                            nKeterangan = "     Jumlah total unit SBE (Pt)"
                            rsTemp.Find "KETERANGAN = 'KPI_11B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(12) = tKPIb(12) + nKPIb
                        Case 42
                            nUrut = 40
                            nKeterangan = "2. KPI 11 : Kontribusi MRS terhadap SBE (=Pd/Pt)"
                            If nKPIb <> 0 Then
                                nHasil = (Round(CDbl(nKPIa / nKPIb) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                        ' ************************************************* '
                        ''' GROUP KPI 12
                        ' ************************************************* '
                        Case 43
                            nUrut = 44
                            nKeterangan = "     Jumlah unit yang telah dikirimkan kepada pelanggan (N-1) (Uk)"
                            rsTemp.Find "KETERANGAN = 'KPI_12A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIa = nKPIa + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIa = nKPIa + 0
                            End If
                            tKPIa(13) = tKPIa(13) + nKPIa
                        Case 44
                            nUrut = 45
                            nKeterangan = "     Jumlah pelanggan yang servis 1.000 dari penjualan unit N-1 (Us)"
                            rsTemp.Find "KETERANGAN = 'KPI_12B'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nKPIb = nKPIb + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, 0, 2)
                                End If
                                nKPIb = nKPIb + 0
                            End If
                            tKPIb(13) = tKPIb(13) + nKPIb
                        Case 45
                            nUrut = 43
                            nKeterangan = "3. KPI 12 : SBI Retention (4) (=Us/Uk) "
                            If nKPIa <> 0 Then
                                nHasil = (Round(CDbl(nKPIb / nKPIa) * 100, 2)) & " %"
                            Else
                                nHasil = "0 %"
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                            Else
                                sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                            End If
                            nKPIa = 0
                            nKPIb = 0
                    End Select
                    nI = nI + 1
                    Call objO.UpdateData(sSQL)
                Loop
            End If
        Next
        ''' UNTUK NILAI TOTAL
        namaTanggal = "TOTAL"
        nI = 1
        Do While nI <= 45 '45
            Select Case nI
                ' ************************************************* '
                ''' GROUP KPI 1
                ' ************************************************* '
                Case 1
                    nUrut = 1
                    nKeterangan = "SALES DATABASE"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2, True)
                    End If
                Case 2
                    nUrut = 3
                    nKeterangan = "     Jumlah data pelanggan yang unitnya telah delivery (Pd)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(1)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(1)), 2)
                    End If
                Case 3
                    nUrut = 4
                    nKeterangan = "     Jumlah data pelanggan valid bulan ini dari yang unitnya telah delivery (Pv)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(1)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(1)), 2)
                    End If
                Case 4
                    nUrut = 2
                    nKeterangan = "1. KPI 1 : Database Valid  (=Pv/Pd)"
                    If tKPIa(1) <> 0 Then
                        nHasil = (Round(CDbl(tKPIb(1) / tKPIa(1)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                Case 5
                    nUrut = 5
                    nKeterangan = "     Jumlah carry over pelanggan valid dari yg unitnya telah delivery"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIc(1)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIc(1)), 2)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 2
                ' ************************************************* '
                Case 6
                    nUrut = 7
                    nKeterangan = "     Jumlah data pelanggan yang unitnya telah delivery (Pd)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(2)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(2)), 2)
                    End If
                Case 7
                    nUrut = 8
                    nKeterangan = "     Jumlah data pelanggan yang unitnya telah delivery ditransfer ke after sales (Pt)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(2)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(2)), 2)
                    End If
                Case 8
                    nUrut = 6
                    nKeterangan = "2. KPI 2 : Database Transfer (=Pt/Pd)"
                    If tKPIa(2) <> 0 Then
                        nHasil = (Round(CDbl(tKPIb(2) / tKPIa(2)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                Case 9
                    nUrut = 9
                    nKeterangan = "     Jumlah carry over data pelanggan yang unitnya telah delivery ditransfer ke after sales"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIc(2)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIc(2)), 2)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 3 SALES AA
                ' ************************************************* '
                Case 10
                    nUrut = 10
                    nKeterangan = "1K SERVICE INVITATION / CR2"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2, True)
                    End If
                Case 11
                    nUrut = 12
                    nKeterangan = "     Jumlah pelanggan yang seharusnya diingatkan servis 1.000 (Pc)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(3)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(3)), 2)
                    End If
                Case 12
                    nUrut = 13
                    nKeterangan = "     Jumlah pelanggan yang aktual diingatkan servis 1.000 (Pa)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(3)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(3)), 2)
                    End If
                Case 13
                    nUrut = 11
                    nKeterangan = "1. KPI 3 : 1K Sales (=Pa/Pc)"
                    If tKPIa(3) <> 0 Then
                        nHasil = (Round(CDbl(tKPIb(3) / tKPIa(3)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 4
                ' ************************************************* '
                Case 14
                    nUrut = 14
                    nKeterangan = "MRS PROCESS"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2, True)
                    End If
                Case 15
                    nUrut = 16
                    nKeterangan = "     Jumlah Direct Mail Diterima Pelanggan (Dt)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(4)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(4)), 2)
                    End If
                Case 16
                    nUrut = 17
                    nKeterangan = "     Jumlah Direct Mail Dikirim (Dk)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(4)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(4)), 2)
                    End If
                Case 17
                    nUrut = 15
                    nKeterangan = "1. KPI 4 : Akurasi Alamat Pelanggan (=Dt/Dk)"
                    If tKPIb(4) <> 0 Then
                        nHasil = (Round(CDbl(tKPIa(4) / tKPIb(4)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 5
                ' ************************************************* '
                Case 18
                    nUrut = 19
                    nKeterangan = "     Jumlah Telepon Terhubung (1) (Tt)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(5)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(5)), 2)
                    End If
                Case 19
                    nUrut = 20
                    nKeterangan = "     Jumlah Pelanggan Yang Ditelepon (Tk)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(5)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(5)), 2)
                    End If
                Case 20
                    nUrut = 18
                    nKeterangan = "2. KPI 5 : Akurasi Nomor Telepon Pelanggan (=Tt/Tk)"
                    If tKPIb(5) <> 0 Then
                        nHasil = (Round(CDbl(tKPIa(5) / tKPIb(5)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 6
                ' ************************************************* '
                Case 21
                    nUrut = 22
                    nKeterangan = "     Jumlah Pelanggan Yang Harus Diundang (P1=Aktif + Pasif + Baru)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(6)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(6)), 2)
                    End If
                Case 22
                    nUrut = 23
                    nKeterangan = "     Jumlah Pelanggan Yang Diundang (P2= DM + SMS + E + Tlp)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(6)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(6)), 2)
                    End If
                Case 23
                    nUrut = 21
                    nKeterangan = "3. KPI 6 : Persentase Pelanggan Yang Diundang (=P2/P1)"
                    If tKPIa(6) <> 0 Then
                        nHasil = (Round(CDbl(tKPIb(6) / tKPIa(6)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                
                ' ************************************************* '
                ''' GROUP KPI 6 MRS
                ' ************************************************* '
                Case 24
                    nUrut = 25
                    nKeterangan = "     Jumlah pelanggan yang seharusnya diundang servis 1.000 (Pc)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(7)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(7)), 2)
                    End If
                Case 25
                    nUrut = 26
                    nKeterangan = "     Jumlah pelanggan yang aktual diundang servis 1.000 (Pa)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(7)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(7)), 2)
                    End If
                Case 26
                    nUrut = 24
                    nKeterangan = "4. KPI 6 : 1K Activity (=Pa/Pc)"
                    If tKPIa(7) <> 0 Then
                        nHasil = (Round(CDbl(tKPIb(7) / tKPIa(7)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 7
                ' ************************************************* '
                Case 27
                    nUrut = 28
                    nKeterangan = "     Jumlah Pelanggan Yang Ditelepon (Tk)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(8)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(8)), 2)
                    End If
                Case 28
                    nUrut = 29
                    nKeterangan = "     Jumlah Pelanggan Yang Diundang (P2)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(8)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(8)), 2)
                    End If
                Case 29
                    nUrut = 27
                    nKeterangan = "5. KPI 7 : Persentase Pelanggan Yang Ditelepon (=Tk/P2)"
                    If tKPIb(8) <> 0 Then
                        nHasil = (Round(CDbl(tKPIa(8) / tKPIb(8)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 8
                ' ************************************************* '
                Case 30
                    nUrut = 31
                    nKeterangan = "     Jumlah Pelanggan Yang Appoinment (2) (A1)"
                        If cekInsert(nUrut, nKeterangan) Then
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(9)), 1)
                        Else
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(9)), 2)
                        End If
                Case 31
                    nUrut = 32
                    nKeterangan = "     Jumlah Telepon Terhubung (Tt)"
                        If cekInsert(nUrut, nKeterangan) Then
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(9)), 1)
                        Else
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(9)), 2)
                        End If
                Case 32
                    nUrut = 30
                    nKeterangan = "6. KPI 8 : Persentase Pelanggan Yang Appoinment (=A1/Tt)"
                    If tKPIb(9) <> 0 Then
                        nHasil = (Round(CDbl(tKPIa(9) / tKPIb(9)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 9
                ' ************************************************* '
                Case 33
                    nUrut = 34
                    nKeterangan = "     Jumlah Pelanggan Yang Datang Setelah Diundang (3) (Pd)"
                        If cekInsert(nUrut, nKeterangan) Then
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(10)), 1)
                        Else
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(10)), 2)
                        End If
                Case 34
                    nUrut = 35
                    nKeterangan = "     Jumlah Pelanggan Yang Diundang (P2)"
                        If cekInsert(nUrut, nKeterangan) Then
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(10)), 1)
                        Else
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(10)), 2)
                        End If
                Case 35
                    nUrut = 33
                    nKeterangan = "7. KPI 9 : Persentase Pelanggan Yang Datang Setelah Diundang (=Pd/P2)"
                    If tKPIb(10) <> 0 Then
                        nHasil = (Round(CDbl(tKPIa(10) / tKPIb(10)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 10
                ' ************************************************* '
                Case 36
                    nUrut = 36
                    nKeterangan = "MRS RESULT"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2, True)
                    End If
                Case 37
                    nUrut = 38
                    nKeterangan = "     Jumlah app. Karena aktifitas MRS (A1)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(11)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(11)), 2)
                    End If
                Case 38
                    nUrut = 39
                    nKeterangan = "     Jumlah total app. (A2)"
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(11)), 1)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(11)), 2)
                    End If
                Case 39
                    nUrut = 37
                    nKeterangan = "1. KPI 10 : Kontribusi MRS terhadap appointment. (=A1/A2)"
                    If tKPIb(11) <> 0 Then
                        nHasil = (Round(CDbl(tKPIa(11) / tKPIb(11)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 11
                ' ************************************************* '
                Case 40
                    nUrut = 41
                    nKeterangan = "     Jumlah unit SBE karena aktifitas MRS (Pd)"
                        If cekInsert(nUrut, nKeterangan) Then
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(12)), 1)
                        Else
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(12)), 2)
                        End If
                Case 41
                    nUrut = 42
                    nKeterangan = "     Jumlah total unit SBE (Pt)"
                        If cekInsert(nUrut, nKeterangan) Then
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(12)), 1)
                        Else
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(12)), 2)
                        End If
                Case 42
                    nUrut = 40
                    nKeterangan = "2. KPI 11 : Kontribusi MRS terhadap SBE (=Pd/Pt)"
                    If tKPIb(12) <> 0 Then
                        nHasil = (Round(CDbl(tKPIa(12) / tKPIb(12)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
                ' ************************************************* '
                ''' GROUP KPI 12
                ' ************************************************* '
                Case 43
                    nUrut = 44
                    nKeterangan = "     Jumlah unit yang telah dikirimkan kepada pelanggan (N-1) (Uk)"
                        If cekInsert(nUrut, nKeterangan) Then
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(13)), 1)
                        Else
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIa(13)), 2)
                        End If
                Case 44
                    nUrut = 45
                    nKeterangan = "     Jumlah pelanggan yang servis 1.000 dari penjualan unit N-1 (Us)"
                        If cekInsert(nUrut, nKeterangan) Then
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(13)), 1)
                        Else
                            sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, CStr(tKPIb(13)), 2)
                        End If
                Case 45
                    nUrut = 43
                    nKeterangan = "3. KPI 12 : SBI Retention (4) (=Us/Uk) "
                    If tKPIa(13) <> 0 Then
                        nHasil = (Round(CDbl(tKPIb(13) / tKPIa(13)) * 100, 2)) & " %"
                    Else
                        nHasil = "0 %"
                    End If
                    If cekInsert(nUrut, nKeterangan) Then
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 1, True)
                    Else
                        sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, nHasil, 2, True)
                    End If
            End Select
            nI = nI + 1
            Call objO.UpdateData(sSQL)
        Loop
        For j = 1 To 36
            nUrut = j
            Select Case j
                Case 1
                    nKeterangan = "SALES DATABASE"
                Case 10
                    nKeterangan = "1K SERVICE INVITATION / CR2"
                Case 14
                    nKeterangan = "MRS PROCESS"
                Case 36
                    nKeterangan = "MRS RESULT"
                Case Else
                    nKeterangan = ""
            End Select
            If Trim(nKeterangan) <> "" Then
                For i = 0 To DateDiff("d", Format(DtDOAwal, "DD/MM/YYYY"), Format(DtDOAkhir, "DD/MM/YYYY"))
                    namaTanggal = "A" & Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
                    sSQL = insertTabel(namaTanggal, nUrut, nKeterangan, "", 2)
                    Call objO.UpdateData(sSQL)
                Next
            End If
        Next
        Call Cetak
    Else
        MsgBox pesan, vbInformation + vbOKOnly, "KPI Harian"
    End If
End Sub

Private Sub cmdSettingPrinter_Click()
    rpt.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub Form_Activate()
    vNamaTabel = "RepoKPIHarian" & DriveSerialNumber("c:")
End Sub

Private Sub Form_Load()
    DtDOAwal = Format(isnow, "DD/MM/YYYY")
    DtDOAkhir = Format(isnow, "DD/MM/YYYY")
End Sub

Private Sub Cetak()
Set rpt = New RptKPIHarian

    sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY URUT"
    Set rsRepo = objO.FetchData(sSQL)
    rpt.Database.SetDataSource rsRepo
    rpt.txtPerusahaan.SetText gPerush
    rpt.txtAlamat.SetText gAlamat
    rpt.txtTanggal.SetText "dari Tanggal " & TglDDMMMMYYYY(Format(DtDOAwal, "DD/MM/YYYY"), 1) & "    ~    " & TglDDMMMMYYYY(Format(DtDOAkhir, "DD/MM/YYYY"), 1)
    '' CLEAR
    rpt.Text1.SetText ""
    rpt.Text2.SetText ""
    rpt.Text3.SetText ""
    rpt.Text4.SetText ""
    rpt.Text5.SetText ""
    rpt.Text6.SetText ""
    rpt.Text7.SetText ""
    rpt.Text8.SetText ""
    rpt.Text9.SetText ""
    rpt.Text10.SetText ""
    rpt.Text11.SetText ""
    rpt.Text12.SetText ""
    rpt.Text13.SetText ""
    rpt.Text14.SetText ""
    rpt.Text15.SetText ""
    rpt.Text16.SetText ""
    rpt.Text17.SetText ""
    rpt.Text18.SetText ""
    rpt.Text19.SetText ""
    rpt.Text20.SetText ""
    rpt.Text21.SetText ""
    rpt.Text22.SetText ""
    rpt.Text23.SetText ""
    rpt.Text24.SetText ""
    rpt.Text25.SetText ""
    rpt.Text26.SetText ""
    rpt.Text27.SetText ""
    rpt.Text28.SetText ""
    rpt.Text29.SetText ""
    rpt.Text30.SetText ""
    rpt.Text31.SetText ""
    rpt.Text32.SetText ""
    
    For i = 0 To DateDiff("d", Format(DtDOAwal, "DD/MM/YYYY"), Format(DtDOAkhir, "DD/MM/YYYY"))
        Select Case i + 1
            Case 1
                rpt.Text1.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 2
                rpt.Text2.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 3
                rpt.Text3.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 4
                rpt.Text4.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 5
                rpt.Text5.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 6
                rpt.Text6.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 7
                rpt.Text7.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 8
                rpt.Text8.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 9
                rpt.Text9.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 10
                rpt.Text10.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 11
                rpt.Text11.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 12
                rpt.Text12.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 13
                rpt.Text13.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 14
                rpt.Text14.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 15
                rpt.Text15.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 16
                rpt.Text16.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 17
                rpt.Text17.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 18
                rpt.Text18.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 19
                rpt.Text19.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 20
                rpt.Text20.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 21
                rpt.Text21.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 22
                rpt.Text22.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 23
                rpt.Text23.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 24
                rpt.Text24.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 25
                rpt.Text25.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 26
                rpt.Text26.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 27
                rpt.Text27.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 28
                rpt.Text28.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 29
                rpt.Text29.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 30
                rpt.Text30.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
            Case 31
                rpt.Text31.SetText Day(DateAdd("d", i, Format(DtDOAwal, "DD/MM/YYYY")))
        End Select
    Next
    j = DateDiff("d", Format(DtDOAwal, "DD/MM/YYYY"), Format(DtDOAkhir, "DD/MM/YYYY")) + 2
    Select Case j
        Case 1
            rpt.Text1.SetText "Total"
        Case 2
            rpt.Text2.SetText "Total"
        Case 3
            rpt.Text3.SetText "Total"
        Case 4
            rpt.Text4.SetText "Total"
        Case 5
            rpt.Text5.SetText "Total"
        Case 6
            rpt.Text6.SetText "Total"
        Case 7
            rpt.Text7.SetText "Total"
        Case 8
            rpt.Text8.SetText "Total"
        Case 9
            rpt.Text9.SetText "Total"
        Case 10
            rpt.Text10.SetText "Total"
        Case 11
            rpt.Text11.SetText "Total"
        Case 12
            rpt.Text12.SetText "Total"
        Case 13
            rpt.Text13.SetText "Total"
        Case 14
            rpt.Text14.SetText "Total"
        Case 15
            rpt.Text15.SetText "Total"
        Case 16
            rpt.Text16.SetText "Total"
        Case 17
            rpt.Text17.SetText "Total"
        Case 18
            rpt.Text18.SetText "Total"
        Case 19
            rpt.Text19.SetText "Total"
        Case 20
            rpt.Text20.SetText "Total"
        Case 21
            rpt.Text21.SetText "Total"
        Case 22
            rpt.Text22.SetText "Total"
        Case 23
            rpt.Text23.SetText "Total"
        Case 24
            rpt.Text24.SetText "Total"
        Case 25
            rpt.Text25.SetText "Total"
        Case 26
            rpt.Text26.SetText "Total"
        Case 27
            rpt.Text27.SetText "Total"
        Case 28
            rpt.Text28.SetText "Total"
        Case 29
            rpt.Text29.SetText "Total"
        Case 30
            rpt.Text30.SetText "Total"
        Case 31
            rpt.Text31.SetText "Total"
    End Select
    
    
    With CRViewer1
        .ReportSource = rpt
        .ViewReport
        .Zoom 100
        .Top = 0
        .Left = Image1.Left + Image1.Width + 20
        .Width = Me.Width - 500
        .Height = Image1.Height
        
        cmdTutup.Top = .Top + 430
        cmdTutup.Left = .Left
        cmdTutup.Visible = True
        
        cmdSettingPrinter.Top = cmdTutup.Top
        cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 20
        cmdSettingPrinter.Visible = True
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Function cekInsert(nID As Integer, nKet As String) As Boolean
Dim rsHlp As ADODB.Recordset

    cekInsert = False
     
    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE urut = '" & nID & "' AND KETERANGAN = '" & nKet & "'")
    If rsHlp.EOF Then
        cekInsert = True
    End If
    rsHlp.Close
    
End Function

Private Function insertTabel(nBulan As String, nBaris As Integer, nKet As String, nJml As String, nkode As Integer, Optional diBold As Boolean = False) As String
    
    If nkode = 1 Then
        insertTabel = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (URUT, KETERANGAN, BOLD, " & nBulan & ")  VALUES (" & nBaris & ",'" & nKet & "','" & IIf(diBold, 1, 0) & "','" & Replace(nJml, ",", ".") & "')"
    ElseIf nkode = 2 Then
        insertTabel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & nBulan & " = '" & Replace(nJml, ",", ".") & "' WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    ElseIf nkode = 3 Then
        insertTabel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & nBulan & " = " & nBulan & " + '" & Replace(nJml, ",", ".") & "' WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    End If
End Function

Private Sub Form_Resize()
    Image1.Height = Me.Height - 50
End Sub
