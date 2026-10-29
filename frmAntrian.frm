VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmAntrian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Antrian"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10410
   Begin VB.Frame framLoket 
      Caption         =   "      ^ Loket Antrian ^     "
      BeginProperty Font 
         Name            =   "MingLiU"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   3615
      Begin MSMask.MaskEdBox Text 
         Height          =   435
         Index           =   1
         Left            =   2025
         TabIndex        =   3
         Top             =   480
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1170
         TabIndex        =   11
         Top             =   480
         Width           =   675
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6975
      Left            =   3840
      TabIndex        =   9
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   6720
   End
   Begin VB.Frame framPanggil 
      Caption         =   "     ^ Panggil Antrian ^    "
      BeginProperty Font 
         Name            =   "MingLiU"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3615
      Begin VB.CommandButton cmdPanggil 
         Caption         =   "Panggil Berikut >>"
         Height          =   390
         Left            =   847
         TabIndex        =   12
         Top             =   2160
         Width           =   1890
      End
      Begin VB.Label lblBook 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Antrian Booking :"
         BeginProperty Font 
            Name            =   "PMingLiU"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2010
      End
      Begin VB.Label lblNonBook 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Antrian Non Booking :"
         BeginProperty Font 
            Name            =   "PMingLiU"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2535
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
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame framCetak 
      Caption         =   "      ^ Cetak Antrian ^     "
      BeginProperty Font 
         Name            =   "MingLiU"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         Height          =   390
         Left            =   840
         TabIndex        =   2
         Top             =   1560
         Width           =   1890
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   435
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Top             =   563
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   767
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CC-CCCC-CCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmAntrian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAntri As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsTemp2 As New ADODB.Recordset
Dim rsAntriLama As New ADODB.Recordset
Dim Report As Report

Private Sub cmdCetak_Click()
    Dim antriBook As Integer
    Dim antriNonBook As Integer
    Dim noAntri As Integer
    Dim ket As String
    
    If notBlank(Text(0)) Then
    'Text(0) = Replace(Text(0), ">", "")
    noAntri = 0
    antriBook = 0
    antriNonBook = 0
        If rsAntriLama.State Then rsAntriLama.Close
        Set rsAntriLama = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE noPol = '" & Trim(Replace(Text(0), ">", "")) & "'")
        If rsAntriLama.RecordCount = 0 Then
            If rsTemp.State Then rsTemp.Close
            Set rsTemp = oConn.Execute("SELECT * FROM book WHERE no_polisi = '" & Trim(Replace(Text(0), ">", "")) & "' AND (deliv = 'BOK' OR deliv = 'RES')")
            If rsTemp.RecordCount > 0 Then
                ' ambil no urut booking
                If rsTemp2.State Then rsTemp2.Close
                Set rsTemp2 = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE statusCust = 'B' ORDER BY noAntrian DESC")
                If rsTemp2.RecordCount > 0 Then
                    antriBook = CDbl(rsTemp2!noAntrian) + 1
                Else
                    antriBook = 1
                End If
                
                ' kendaraan customer booking
                If rsTemp!Time < Format(Now(), "HH:MM") Then
                    ket = "Konsumen Booking Terlambat " & Format(Format(Now(), "HH:MM") - rsTemp!Time, "HH:MM") & " menit"
                Else
                    ket = ""
                End If
                oConn.Execute "INSERT INTO T_ANTRIAN VALUES ('" & CDbl(antriBook) & "','B','" & Trim(Replace(Replace(Text(0), ">", ""), " ", "")) & "','" & Format(Trim(Now()), "mm/dd/yyyy HH:MM:SS") & "','','" & Trim(ket) & "') "
                MsgBox "Selamat Datang" & vbCr & "Anda Terdaftar sebagai Pelanggan Booking" & vbCr & vbCr & "Silahkan Menunggu di ruang tunggu"
            Else
                ' ambil no urut non booking
                If rsTemp2.State Then rsTemp2.Close
                Set rsTemp2 = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE statusCust = 'N' ORDER BY noAntrian DESC")
                If rsTemp2.RecordCount > 0 Then
                    antriNonBook = CDbl(rsTemp2!noAntrian) + 1
                Else
                    antriNonBook = 1
                End If
            
                ' kendaraan customer non booking
                oConn.Execute "INSERT INTO T_ANTRIAN VALUES ('" & CDbl(antriNonBook) & "','N','" & Trim(Replace(Replace(Text(0), ">", ""), " ", "")) & "','" & Format(Trim(Now()), "mm/dd/yyyy HH:MM:SS") & "','','" & Trim(ket) & "') "
                MsgBox "Selamat Datang" & vbCr & "Silahkan Menunggu di ruang tunggu"
            End If
            noAntri = IIf(antriBook > 0, antriBook, antriNonBook)
        
            Call PrintNoAntrian2(noAntri, Replace(Replace(Text(0), ">", ""), " ", ""))
            Text(0).Mask = ""
            Text(0) = ""
            Text(0).Mask = "CC-CCCC-CCC"
            Text(0).SelLength = Len(Text(0))
            Text(0).SetFocus
        Else
            MsgBox "Sudah dalam Antrian, Tidak Perlu mengambil nomor Antrian"
            Exit Sub
        End If
    End If
End Sub

Private Function AntrianPanggil(tipe As String) As String
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE StatusCust = '" & Trim(tipe) & "' and StatusPanggil <> 'Y' ORDER BY noAntrian")
    If rsTemp.RecordCount > 0 Then
        ' panggil dengan suara
        'MsgBox "Pemilik kendaraan dengan No Polisi " & Trim(rsTemp!noPol) & " diharap menuju loket " & Trim(gLoket)
        Call PanggilCust("AwalAntrian.wav", Trim(rsTemp!noPol), "AkhirAntrian.wav", Trim(gLoket))
        
        oConn.Execute "UPDATE T_ANTRIAN SET statusPanggil = 'Y' WHERE noPol = '" & Trim(rsTemp!noPol) & "' and statusPanggil <> 'Y'"
        AntrianPanggil = Trim(rsTemp!noPol) & vbCr & "Keterangan : " & Trim(rsTemp!Keterangan)
    End If
End Function

Private Sub cmdPanggil_Click()
    Dim nopl As String
    
    If framPanggil.Visible Then
        n = frmOrder.Text(25).Mask
        frmOrder.Text(25).Mask = ""
        If CDbl(Right(lblBook, (Len(lblBook) - 18))) > 0 Then
            nopl = AntrianPanggil("B")
            frmOrder.cmdPilih.Caption = "No. Booking"
            frmOrder.Text(25) = noPanggil(Left(nopl, InStr(nopl, vbCr) - 1))
        ElseIf CDbl(Right(lblNonBook, (Len(lblNonBook) - 22))) > 0 Then
            nopl = AntrianPanggil("N")
            frmOrder.cmdPilih.Caption = "No. Frame"
            frmOrder.Text(25) = noPanggil(Left(nopl, InStr(nopl, vbCr) - 1))
        Else
            MsgBox "Antrian Kosong"
        End If
        noPol.Caption = "No. Polisi " & nopl
        
        frmOrder.Text(25).Mask = n
        If frmOrder.Text(25) <> "" Then
            frmOrder.Show
        Else
            frmtCustomer.Show
            frmtPlgn.Show
        End If
    End If
End Sub

Private Function noPanggil(nPol As String) As String
    Dim rsT As New ADODB.Recordset
    Dim rsRg As New ADODB.Recordset
    
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = oConn.Execute("SELECT * FROM book WHERE no_polisi = '" & Trim(nPol) & "' AND deliv = 'BOK'")
    If Not rsTemp.EOF Then
        ' CUSTOMER BOOKING
        noPanggil = Trim(rsTemp!noFak)
    Else
        If rsT.State Then rsT.Close
        oSQL = "SELECT * FROM custdat where no_polisi = '" & Trim(nPol) & "'"
        Set rsT = oConn.Execute(oSQL)
        If Not rsT.EOF Then
            ' CUSTOMER NON BOOKING
            noPanggil = Trim(rsT!no_chassis)
            If rsRg.State Then rsRg.Close
            Set rsRg = oConn.Execute("SELECT * FROM jobdta WHERE no_polisi = '" & Trim(nPol) & "' AND RPR_COND <> 'K' ORDER BY [date] DESC")
            If Not rsRg.EOF Then
                ' INVOICE LAMA BELUM DITUTUP
                MsgBox "WO lama Belum Registrasi !!!"
                n = frmRegistrasi.Text(0).Mask
                frmRegistrasi.Text(0) = Trim(rsRg!No_WO)
                frmRegistrasi.Text(0).Mask = n
                frmRegistrasi.Show
            End If
        Else
            ' CUSTOMER BARU
            MsgBox "Customer Baru, tambah Pelanggan"
            noPanggil = ""
        End If
    End If
End Function

Private Sub Form_Load()
    If Trim(antrian) = "Cetak" Then
        framPanggil.Visible = False
        framLoket.Visible = False
        framCetak.Visible = True
        framCetak.Top = 120
        framCetak.Left = 120
        Me.Width = framCetak.Width + 240
        Me.Height = framCetak.Height + 700
        'Text(0) = ""
    ElseIf Trim(antrian) = "Panggil" Then
        framCetak.Visible = False
        framLoket.Visible = True
        framLoket.Top = 120
        framLoket.Left = 120
        Me.Width = framLoket.Width + 240
        Me.Height = framLoket.Height + 700
        
        Call hitungAntrian
    End If
    Timer1.Enabled = True
End Sub

Private Sub hitungAntrian()
    Dim antriBook As Integer
    Dim antriNonBook As Integer
    
    antriBook = 0
    antriNonBook = 0
    If rsAntri.State Then rsAntri.Close
    Set rsAntri = oConn.Execute("SELECT * FROM T_ANTRIAN WHERE statusPanggil <> 'Y' ORDER BY noAntrian")
    If rsAntri.RecordCount > 0 Then
        Do While Not rsAntri.EOF
            If Trim(rsAntri!StatusCust) = "B" Then
                ' customer booking
                antriBook = antriBook + 1
            ElseIf Trim(rsAntri!StatusCust) = "N" Then
                ' customer non booking
                antriNonBook = antriNonBook + 1
            End If
            rsAntri.MoveNext
        Loop
    End If
    lblBook = "Antrian Booking : " & antriBook
    lblNonBook = "Antrian Non Booking : " & antriNonBook
    Layar.StatusBar1.Panels(7).Text = "  " & lblBook & " / " & lblNonBook
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If (Index = 1 And KeyAscii = 13) Then
        gLoket = Replace(Trim(Text(1)), ">", "")
        framLoket.Visible = False
        framPanggil.Top = 120
        framPanggil.Left = 120
        framPanggil.Caption = "^ Antrian ^o0o^ Loket  " & gLoket & "^"
        Me.Width = framPanggil.Width + 240
        Me.Height = framPanggil.Height + 700
        framPanggil.Visible = True
        Exit Sub
    ElseIf (Index = 0 And KeyAscii = 13) Then
        Enter KeyAscii
    End If
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub Timer1_Timer()
    Call hitungAntrian
End Sub

Private Function PrintNoAntrian2(No As Integer, noP As String)
    With Printer
        .FontName = "Trebuchet MS"
        
        '### NAMA PERUSAHAAN ###
        .Font.Size = 14
        .Font.Bold = True
        Printer.CurrentX = (.ScaleWidth - .TextWidth(gPerush)) / 2
        Printer.Print UCase(Trim(gPerush))
        
        '### ALAMAT PERUSAHAAN ###
        .Font.Bold = False
        .Font.Size = 11
        Printer.CurrentX = (.ScaleWidth - .TextWidth(gAlamat)) / 2
        Printer.Print UCase(Trim(gAlamat))
        
        Printer.CurrentX = (.ScaleWidth - .TextWidth(gCabang) - .TextWidth(gKota)) / 2
        Printer.Print UCase(Trim(gCabang)) & " - " & UCase(Trim(gKota))
        
        Printer.CurrentX = (.ScaleWidth - .TextWidth(gTelp) - .TextWidth("Telp. ")) / 2
        Printer.Print "Telp. " & UCase(Trim(gTelp))
        
        Printer.CurrentX = (.ScaleWidth - .TextWidth("========================")) / 2
        Printer.Print "========================"
        
        Printer.Print Space(10)
        Printer.Print Space(10)
        
        .Font.Size = 16
        Printer.Print "No. Antrian : " & Format(No, "####0")
        .Font.Size = 11
        Printer.Print "No. Polisi  : " & UCase(Trim(noP))
        
        Printer.Print Space(10)
        Printer.Print Space(10)
        Printer.Print Space(10)
        
        Printer.Print "Tanggal. " & Format(Now(), "dd/mm/yyyy") & vbTab & vbTab & vbTab & "Jam. " & Format(Now(), "HH:MM:SS")
        Printer.Print Space(10)
    End With
End Function

Private Function PrintNoAntrian(No As Integer, noP As String)
    Dim frm As Form
    Dim rsRepo As New ADODB.Recordset
    
    Set Report = New RptAntrian
    Set frm = frmAntrian
    
    'Me.Height = CRViewer1.Height + 200
    'Me.Width = framCetak.Width + CRViewer1.Width + 400
    
    On Error GoTo vbError
    With Report
        .Text1.SetText Trim(gPerush)
        .Text2.SetText Trim(gAlamat)
        .Text3.SetText Trim(gKota)
        .Text5.SetText Format(No, "####0")
        .Text7.SetText Trim(noP)
    End With
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .PrintReport
    End With
    Exit Function
    
vbError:
    MsgBox ""
    Exit Function
End Function
