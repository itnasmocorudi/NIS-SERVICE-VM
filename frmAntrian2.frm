VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmAntrian2 
   BackColor       =   &H00BFE8EA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANTRIAN"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   12735
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame framCetak 
      BackColor       =   &H00BFE8EA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   5895
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Antrian Ambil Kend :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Visible         =   0   'False
         Width           =   5610
      End
      Begin VB.Image cmdCetak 
         Height          =   1560
         Index           =   2
         Left            =   3840
         Picture         =   "frmAntrian2.frx":0000
         Top             =   840
         Width           =   1665
      End
      Begin VB.Image cmdCetak 
         Height          =   1560
         Index           =   1
         Left            =   2040
         Picture         =   "frmAntrian2.frx":16A2
         Top             =   840
         Width           =   1665
      End
      Begin VB.Image cmdCetak 
         Height          =   1560
         Index           =   0
         Left            =   240
         Picture         =   "frmAntrian2.frx":2DCD
         Top             =   840
         Width           =   1665
      End
      Begin VB.Label lb1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2890
         Width           =   5415
      End
      Begin VB.Label lb2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   250
         TabIndex        =   8
         Top             =   2865
         Width           =   5415
      End
      Begin VB.Image ImgOff 
         Height          =   1560
         Index           =   2
         Left            =   3840
         Picture         =   "frmAntrian2.frx":42B8
         Top             =   840
         Width           =   1665
      End
      Begin VB.Image ImgOff 
         Height          =   1560
         Index           =   1
         Left            =   2040
         Picture         =   "frmAntrian2.frx":581C
         Top             =   840
         Width           =   1665
      End
      Begin VB.Image ImgOff 
         Height          =   1560
         Index           =   0
         Left            =   240
         Picture         =   "frmAntrian2.frx":6E8B
         Top             =   840
         Width           =   1665
      End
      Begin VB.Image Img 
         Height          =   3285
         Left            =   0
         Picture         =   "frmAntrian2.frx":82B0
         Top             =   0
         Width           =   5820
      End
   End
   Begin VB.Frame framPanggil 
      BackColor       =   &H00BFE8EA&
      Caption         =   "                             ^ Panggil Antrian ^                               "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   5775
      Begin VB.CommandButton cmdPanggil 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Panggil Booking >>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3000
         Width           =   1890
      End
      Begin VB.CommandButton cmdUlang 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Panggil Ulang"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3000
         Width           =   1410
      End
      Begin VB.CommandButton cmdPanggil 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Panggil Non Booking >>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   2250
      End
      Begin VB.Label jamMekanik 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Kerja Mekanik belum disetting !!!"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   405
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Visible         =   0   'False
         Width           =   5250
      End
      Begin VB.Label lblAmbil 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Antrian Ambil Kend :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   5520
      End
      Begin VB.Label noPol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblNonBook 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Antrian Non Booking :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   5535
      End
      Begin VB.Label lblBook 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Antrian Booking :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   5535
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6975
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmAntrian2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : -
'=======================================

Dim rsTemp As New ADODB.Recordset
Dim rsAntri As New ADODB.Recordset
Dim antrianNo As Integer
Dim antBook As Integer
Dim antNon As Integer
Dim antAmb As Integer

'Private Function cetakAntrian(No As Integer, noP As String)
'    With Printer
'        .FontName = "Trebuchet MS"
'
'        '### NAMA PERUSAHAAN ###
'        .Font.Size = 14
'        .Font.Bold = True
'        Printer.CurrentX = (.ScaleWidth - .TextWidth(gPerush)) / 2
'        Printer.Print UCase(Trim(gPerush))
'
'        '### ALAMAT PERUSAHAAN ###
'        .Font.Bold = False
'        .Font.Size = 11
'        Printer.CurrentX = (.ScaleWidth - .TextWidth(gAlamat)) / 2
'        Printer.Print UCase(Trim(gAlamat))
'
'        Printer.CurrentX = (.ScaleWidth - .TextWidth(gCabang) - .TextWidth(gKota)) / 2
'        Printer.Print UCase(Trim(gCabang)) & " - " & UCase(Trim(gKota))
'
'        Printer.CurrentX = (.ScaleWidth - .TextWidth(gTelp) - .TextWidth("Telp. ")) / 2
'        Printer.Print "Telp. " & UCase(Trim(gTelp))
'
'        Printer.CurrentX = (.ScaleWidth - .TextWidth("========================")) / 2
'        Printer.Print "========================"
'
'        Printer.Print Space(10)
'        Printer.Print Space(10)
'
'        .Font.Size = 16
'        Printer.Print "No. Antrian : " & Format(No, "####0")
'        .Font.Size = 11
'        'Printer.Print "No. Polisi  : " & UCase(Trim(noP))
'        If noP = "A" Then
'            Printer.Print "AMBIL KENDARAAN"
'        ElseIf noP = "B" Then
'            Printer.Print "BOOKING"
'        ElseIf noP = "N" Then
'            Printer.Print "NON BOOKING"
'        End If
'
'        Printer.Print Space(10)
'        Printer.Print Space(10)
'        Printer.Print Space(10)
'
'        Printer.Print "Tanggal. " & Format(gTglServer, "dd/mm/yyyy") & vbTab & vbTab & vbTab & "Jam. " & Format(gTglServer, "HH:NN:SS")
'        Printer.Print Space(10)
'        Printer.EndDoc
'    End With
'End Function

Private Sub Cetak(Layar As Byte, No As Integer, noP As String)
    Dim rpt As New RptAntrian2
    
'    gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
'    sSQL = "select date,no_polisi,customer,(CASE noinv WHEN '' THEN 0 ELSE total END) as TOTAL,front,RPR_COND from JOBDTA " _
'         & " where date = '" & TglYYYYMMDD(Text(0)) & "'" _
'         & " order by date,no_polisi"
'    rpt.DataControl1.ConnectionString = gReportConn
'    rpt.DataControl1.Source = sSQL
    
    rpt.Field1 = "" 'UCase(Trim(gCabang))
    rpt.Field2 = UCase(Trim(gPerush))
    rpt.Field3 = UCase(Trim(gAlamat)) & " - " & UCase(Trim(gKotaCbg))
    rpt.Field8 = "Telp. " & UCase(Trim(gTelp))
    rpt.Field4 = Format(No, "####0")
    
    If noP = "A" Then
        rpt.Field5 = "AMBIL KENDARAAN"
    ElseIf noP = "B" Then
        rpt.Field5 = "BOOKING"
    ElseIf noP = "N" Then
        rpt.Field5 = "NON BOOKING"
    End If
    rpt.Field6 = Format(gTglServer, "dd/mm/yyyy")
    rpt.Field7 = Format(gTglServer, "HH:NN:SS")
    
    rpt.Printer.Orientation = ddOPortrait
    rpt.Printer.PaperSize = 14
    Select Case Layar
        Case 1
            rpt.Show
        Case 2
            ARViewer1.ReportSource = rpt
        Case Else
            rpt.Show
            rpt.PrintReport False
            Unload rpt
    End Select
End Sub

Private Sub cekJamKerja()
Dim SQL As String

    SQL = "SELECT * FROM T_MEKANIK WHERE TANGGAL = '" & Format(IsNow, "DD/MM/YYYY") & "' AND MASUK = 1"
    Dim rs As New ADODB.Recordset
    If rs.State Then rs.Close
    Set rs = objO.FetchData(SQL)
    If rs.EOF Then
        jamMekanik.Visible = True
        If jamMekanik.ForeColor = &HFF& Then
             jamMekanik.ForeColor = &H80FF&
        ElseIf jamMekanik.ForeColor = &H80FF& Then
             jamMekanik.ForeColor = &HFFFF&
        ElseIf jamMekanik.ForeColor = &HFFFF& Then
             jamMekanik.ForeColor = &HFF&
        End If
    Else
        jamMekanik.Visible = False
    End If
    rs.Close
End Sub

Private Sub hitungAntrian()
    antBook = 0
    antNon = 0
    antAmb = 0
    If rsAntri.State Then rsAntri.Close
    Set rsAntri = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE statusPanggil <> 'Y' ORDER BY StatusCust, noAntrian")
    If Not rsAntri.EOF Then
        Do While Not rsAntri.EOF
            Select Case rsAntri("StatusCust")
                Case "A"
                    antAmb = antAmb + 1
                Case "B"
                    antBook = antBook + 1
                Case "N"
                    antNon = antNon + 1
            End Select
            rsAntri.MoveNext
        Loop
    End If
    
    lblAmbil = "Ambil Kendaraan : " & antAmb
    lblBook = "Antrian Booking : " & antBook
    lblNonBook = "Antrian Non Booking : " & antNon
    If antAmb = 0 And antNon = 0 Then
        cmdPanggil(1).Enabled = False
    Else
        cmdPanggil(1).Enabled = True
        If antBook = 0 Then
            cmdPanggil(0).Enabled = False
        Else
            cmdPanggil(0).Enabled = True
        End If
    End If
End Sub

Private Sub cmdCetak_Click(Index As Integer)
    
    ClearsSql
    Select Case Index
        Case 0  'Booking
            If rsTemp.State Then rsTemp.Close
            Set rsTemp = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE StatusCust = 'B' AND jam between '" & Format(gTglServer, "mm/dd/yyyy") & " 00:00:00' AND '" & Format(gTglServer, "mm/dd/yyyy") & " 23:59:59' ORDER BY noAntrian DESC")
            If Not rsTemp.EOF Then
                antrianNo = CDbl(rsTemp("noAntrian")) + 1
            Else
                antrianNo = 1
            End If
            gsSql1(0) = "NOANTRIAN"
            gsSql2(0) = antrianNo
            gsSql1(1) = "STATUSCUST"
            gsSql2(1) = "B"
        Case 1  'Non Booking
            If rsTemp.State Then rsTemp.Close
            Set rsTemp = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE StatusCust = 'N' AND jam between '" & Format(gTglServer, "mm/dd/yyyy") & " 00:00:00' AND '" & Format(gTglServer, "mm/dd/yyyy") & " 23:59:59' ORDER BY noAntrian DESC")
            If Not rsTemp.EOF Then
                antrianNo = CDbl(rsTemp("noAntrian")) + 1
            Else
                antrianNo = 1
            End If
            gsSql1(0) = "NOANTRIAN"
            gsSql2(0) = antrianNo
            gsSql1(1) = "STATUSCUST"
            gsSql2(1) = "N"
            
        Case 2  'Ambil Kendaraan
            If rsTemp.State Then rsTemp.Close
            Set rsTemp = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE StatusCust = 'A' AND jam between '" & Format(gTglServer, "mm/dd/yyyy") & " 00:00:00' AND '" & Format(gTglServer, "mm/dd/yyyy") & " 23:59:59' ORDER BY noAntrian DESC")
            If Not rsTemp.EOF Then
                antrianNo = CDbl(rsTemp("noAntrian")) + 1
            Else
                antrianNo = 1
            End If
            gsSql1(0) = "NOANTRIAN"
            gsSql2(0) = antrianNo
            gsSql1(1) = "STATUSCUST"
            gsSql2(1) = "A"
    End Select
    gsSql1(2) = "NOPOL"
    gsSql2(2) = ""
    gsSql1(3) = "JAM"
    gsSql2(3) = Format(gTglServer, "MM/DD/YYYY HH:NN:SS")
    gsSql1(4) = "STATUSPANGGIL"
    gsSql2(4) = ""
    gsSql1(5) = "KETERANGAN"
    gsSql2(5) = ""
    
    sSQL = GensSql2("T_ANTRIAN", False)
    oConn.Execute sSQL
    
    'Call cetakAntrian(antrianNo, gsSql2(1))
    Call Cetak(0, antrianNo, gsSql2(1))
End Sub

Private Sub cmdCetak_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdPanggil_Click(Index As Integer)

    If rsTemp.State Then rsTemp.Close
    Select Case Index
        Case 0
            Set rsTemp = oConn.Execute("select * from t_antrian where statusPanggil <> 'Y' AND statusCust = 'B' order by statusCust, noantrian")
        Case 1
            Set rsTemp = oConn.Execute("select * from t_antrian where statusPanggil <> 'Y' AND statusCust <> 'B' order by statusCust, noantrian")
    End Select
    If Not rsTemp.EOF Then
        
        gNoAntri = rsTemp("NOANTRIAN")
        gTpAntri = rsTemp("STATUSCUST")
        oConn.Execute ("UPDATE T_ANTRIAN SET statusPanggil = 'Y', jamPanggil = GETDATE(), loket = '" & gLoket & "' WHERE noAntrian = '" & Trim(rsTemp!noAntrian) & "' AND statusCust = '" & Trim(rsTemp!StatusCust) & "'")
        Call PanggilCust("AwalAntrian.wav", rsTemp("Noantrian"), "AkhirAntrian.wav", True, gLoket, rsTemp("statusCust"))
        cmdUlang.Enabled = True
        
        With frmOrder
            If Trim(gTpAntri) <> "A" Then
                If Trim(gTpAntri) = "B" Then
                    .cmdPilih.Caption = "No. Booking"
                    .cmdPilih.Picture = .ImageList1.ListImages(1).Picture
                ElseIf Trim(gTpAntri) = "N" Then
                    .cmdPilih.Caption = "No. Rangka"
                    .cmdPilih.Picture = .ImageList1.ListImages(2).Picture
                End If
                .Show
            Else
                frmOrderValidasi.Show
                frmOrderValidasi.WindowState = vbMaximized
            End If
        End With
    End If
End Sub

Private Sub cmdUlang_Click()
Dim rsSearch As New ADODB.Recordset
Dim bolehPanggil As Boolean

    If gNoAntri <> 0 Then
    
        bolehPanggil = True
        If rsSearch.State Then rsSearch.Close
        Set rsSearch = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE NOANTRIAN = '" & Trim(gNoAntri) & "' AND STATUSCUST = '" & Trim(gTpAntri) & "'")
        If Not rsSearch.EOF Then
            If Trim(rsSearch("NO_WO")) <> "" Then
                bolehPanggil = False
            End If
        Else
            bolehPanggil = False
        End If
        If rsSearch.State <> adStateClosed Then rsSearch.Close
        
        If bolehPanggil Then
            Call PanggilCust("AwalAntrian.wav", str(gNoAntri), "AkhirAntrian.wav", True, gLoket, gTpAntri)
            oConn.Execute ("UPDATE T_ANTRIAN SET statusPanggil = 'Y', jamPanggil = GETDATE(), loket = '" & gLoket & "' WHERE noAntrian = '" & Trim(gNoAntri) & "' AND statusCust = '" & Trim(gTpAntri) & "'")
            With frmOrder
                If Trim(gTpAntri) <> "A" Then
                    If Trim(gTpAntri) = "B" Then
                        .cmdPilih.Caption = "No. Booking"
                        .cmdPilih.Picture = .ImageList1.ListImages(1).Picture
                    ElseIf Trim(gTpAntri) = "N" Then
                        .cmdPilih.Caption = "No. Rangka"
                        .cmdPilih.Picture = .ImageList1.ListImages(2).Picture
                    End If
                    .Show
                Else
                    frmOrderValidasi.Show
                    frmOrderValidasi.WindowState = vbMaximized
                End If
            End With
        End If
    End If
End Sub

Private Sub Form_Load()
    Call SetPicOff
    If Trim(antrian) = "Cetak" Then
        framPanggil.Visible = False
        framCetak.Visible = True
        framCetak.Top = 120
        framCetak.Left = 120
        Me.Width = framCetak.Width + 240
        Me.Height = framCetak.Height + 700
        ' LIHAT LAGI
        Label1.Caption = IIf(Weekday(gTglServer) = 1, 6, Weekday(gTglServer) - 2) & WeekdayName(IIf(Weekday(gTglServer) = 1, 8, Weekday(gTglServer)) - 1)
        'WeekdayName(IIf(Weekday(gTglServer) = 1, Weekday(gTglServer) + 6, Weekday(gTglServer) - 1)) & " - " & Format(gTglServer, "DD MMMM YYYY")
        'cmdCetak(0).SetFocus
    ElseIf Trim(antrian) = "Panggil" Then
        cmdUlang.Enabled = False
        framCetak.Visible = False
        framPanggil.Visible = True
        framPanggil.Top = 120
        framPanggil.Left = 120
        Me.Width = framPanggil.Width + 240
        Me.Height = framPanggil.Height + 700
        'cmdPanggil.SetFocus
        Call hitungAntrian
    End If
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    'If nC10 > 300 Then
        Call hitungAntrian
        Call cekJamKerja
    'End If
End Sub


'.:: =================================================================== ::.
'.:: =================================================================== ::.
'.:: =================================================================== ::.
Private Sub SetPicOff()
   cmdCetak(0).Visible = False
   cmdCetak(1).Visible = False
   cmdCetak(2).Visible = False
End Sub

Private Sub cmdCetak_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
      Case 0
         lb1.Caption = "Antrian Customer Booking"
         lb2.Caption = "Antrian Customer Booking"
      Case 1
         lb1.Caption = "Antrian Customer Non Booking"
         lb2.Caption = "Antrian Customer Non Booking"
      Case 2
         lb1.Caption = "Antrian Customer Ambil Mobil"
         lb2.Caption = "Antrian Customer Ambil Mobil"
   End Select
End Sub
Private Sub ImgOff_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetPicOff
   cmdCetak(Index).Visible = True
End Sub
Private Sub Img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.lb1.Caption = ".:: Selamat Datang ::."
   Me.lb2.Caption = ".:: Selamat Datang ::."
End Sub

