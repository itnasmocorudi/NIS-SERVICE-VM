VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmRepoHOBulanan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   12585
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
      Height          =   270
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   600
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   7575
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton TblSimpan 
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4440
      TabIndex        =   5
      Top             =   2760
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6360
      TabIndex        =   6
      Top             =   2760
      Width           =   1890
   End
   Begin VB.ComboBox cboBulan1 
      Height          =   315
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox cboBulan2 
      Height          =   315
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   8400
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   14817
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "   Laporan Bulanan "
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
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   270
      Left            =   4515
      TabIndex        =   11
      Top             =   1320
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   6
      Left            =   1605
      TabIndex        =   4
      Top             =   1320
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   7
      Left            =   1605
      TabIndex        =   12
      Top             =   1680
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Penjual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   13
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   945
      TabIndex        =   10
      Top             =   630
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   945
      TabIndex        =   9
      Top             =   990
      Width           =   570
   End
End
Attribute VB_Name = "frmRepoHOBulanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim vNamaTabel As String

Private Sub hapusForm()
    For i = 0 To 1
        text(i) = 0
    Next
    
    cboBulan1.Clear
    cboBulan2.Clear
    For i = 1 To 12
        cboBulan1.AddItem Format(i, "00") & " | " & MonthName(i)
        cboBulan2.AddItem Format(i, "00") & " | " & MonthName(i)
    Next
    cboBulan1 = Format(Month(isnow), "00") & " | " & MonthName(Month(isnow))
    cboBulan2 = Format(Month(isnow), "00") & " | " & MonthName(Month(isnow))
    text(0) = Year(isnow)
    text(1) = Year(isnow)
End Sub

Private Function cekInsert(nID As Integer, nKet As String) As Boolean
Dim rsHlp As ADODB.Recordset

    cekInsert = False
    Set rsHlp = objO.FetchDataHO("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE urut = '" & nID & "' AND KETERANGAN = '" & nKet & "'")
    If rsHlp.EOF Then
        cekInsert = True
    End If
    
End Function

Private Function insertTabel(nBulan As String, nBaris As Integer, nKet As String, nJml As Double, nkode As Integer) As String
    
    If nkode = 1 Then
        insertTabel = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (URUT, KETERANGAN, " & nBulan & ")  VALUES (" & nBaris & ",'" & nKet & "'," & CDbl(nJml) & ")"
    ElseIf nkode = 2 Then
        insertTabel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & nBulan & " = " & CDbl(nJml) & " WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    End If
End Function

Private Sub cboBulan1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cboBulan2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cmdTutup_Click()
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateDataHO(sSQL)
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Activate()
    vNamaTabel = "RepoHOBulanan" & DriveSerialNumber("c:")
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblSimpan_Click()
Dim pesan As String
Dim temp As String
Dim boleh As Boolean
Dim nI As Integer
Dim nTotal As Double
Dim nSubTotal As Double
Dim namaBulan As String
Dim nUrut As Integer
Dim nKeterangan As String
Dim nDealer As String
    
    boleh = True
    If text(0) = "" Then
        pesan = pesan + "Tahun Awal. "
        boleh = False
    End If
    If text(1) = "" Then
        pesan = pesan + "Tahun Akhir. "
        boleh = False
    End If
    If CDbl(text(0)) <> CDbl(text(1)) Then
        pesan = pesan + "Tahun harus sama. "
        boleh = False
    End If
    If cboBulan1 = "" Then
        pesan = pesan + "Bulan Awal. "
        boleh = False
    End If
    If cboBulan2 = "" Then
        pesan = pesan + "Bulan Akhir. "
        boleh = False
    End If
    If Trim(text(6)) = "" Then
        nDealer = "NR "
    Else
        nDealer = Trim(text(6))
    End If
    
    If boleh Then
        sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
             & " drop table [tempdb]..[" & vNamaTabel & "]"
        Call objO.UpdateDataHO(sSQL)
        
        temp = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ([URUT][NUMERIC](9) NOT NULL DEFAULT 0, [KETERANGAN][varchar](100) NOT NULL DEFAULT '' "
        For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
            temp = temp + ", [" & GetNamaBulan(i) & "][money] NOT NULL DEFAULT 0"
        Next
        temp = temp + ") ON [PRIMARY]"
        Call objO.UpdateDataHO(temp)
        
        For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
            aConn.Open (SQLConHO)
            Set rsTemp = aConn.Execute("EXEC REPOBULANAN " & i & ", " & text(0) & ", '" & nDealer & "'")
            aConn.Close
            If Not rsTemp.EOF Then
                nTotal = 0
                nSubTotal = 0
                nI = 1
                namaBulan = GetNamaBulan(i)
                Do While nI <= 33
                    Select Case nI
                        Case 1
                            nUrut = 1
                            nKeterangan = "10K saja "
                            rsTemp.Find "KETERANGAN = '10'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 2
                            nUrut = 2
                            nKeterangan = "20K saja "
                            rsTemp.Find "KETERANGAN = '20'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 3
                            nUrut = 3
                            nKeterangan = "30K saja "
                            rsTemp.Find "KETERANGAN = '30'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 4
                            nUrut = 4
                            nKeterangan = "40K saja "
                            rsTemp.Find "KETERANGAN = '40'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 5
                            nUrut = 5
                            nKeterangan = "50K saja "
                            rsTemp.Find "KETERANGAN = '50'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 6
                            nUrut = 6
                            nKeterangan = "60K saja "
                            rsTemp.Find "KETERANGAN = '60'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 7
                            nUrut = 7
                            nKeterangan = "70K saja "
                            rsTemp.Find "KETERANGAN = '70'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 8
                            nUrut = 8
                            nKeterangan = "80K saja "
                            rsTemp.Find "KETERANGAN = '80'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 9
                            nUrut = 9
                            nKeterangan = "90K saja "
                            rsTemp.Find "KETERANGAN = '90'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 10
                            nUrut = 10
                            nKeterangan = "100K saja "
                            rsTemp.Find "KETERANGAN = '100'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 11
                            nUrut = 11
                            nKeterangan = "Kelipatan 5K dibawah 100K "
                            rsTemp.Find "KETERANGAN = 'Kelipatan 5K'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 12
                            nUrut = 12
                            nKeterangan = "SBE lainnya diatas 100K "
                            rsTemp.Find "KETERANGAN = 'Lainnya Diatas 100K'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 13
                            ' TOTAL SBE
                            nUrut = 13
                            nKeterangan = "  TOTAL Servis Berkala Eksternal "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nSubTotal = nSubTotal + nTotal
                            nTotal = 0
                        Case 14
                            nUrut = 14
                            nKeterangan = "Unit Ganti Oli Mesin "
                            rsTemp.Find "KETERANGAN = 'Lainnya Ganti Oli'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 15
                            nUrut = 15
                            nKeterangan = "General Repair Lainnya "
                            rsTemp.Find "KETERANGAN = 'Lainnya General Repair'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 16
                            ' SUB TOTAL GRP
                            nUrut = 16
                            nKeterangan = "  SUBTOTAL General Repair "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nSubTotal = nSubTotal + nTotal
                            nTotal = 0
                        Case 17
                            ' TOTAL CPUS
                            nUrut = 17
                            nKeterangan = "  TOTAL CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 2)
                            End If
                            nTotal = 0
                        Case 18
                            nUrut = 18
                            nKeterangan = "1K Penjualan Outlet Sama "
                            rsTemp.Find "KETERANGAN = 'SBI 1 dealer'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                        Case 19
                            nUrut = 19
                            nKeterangan = "1K Total "
                            rsTemp.Find "KETERANGAN = 'SBI beda dealer'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 20
                            nUrut = 20
                            nKeterangan = "Warranty "
                            rsTemp.Find "KETERANGAN = 'nilai WARRANTY'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 21
                            nUrut = 21
                            nKeterangan = "Return Job "
                            rsTemp.Find "KETERANGAN = 'Return Job'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 22
                            nUrut = 22
                            nKeterangan = "Pre Delivery Servis "
                            rsTemp.Find "KETERANGAN = 'PDS'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 23
                            ' SUBTOTAL NONCPUS
                            nUrut = 23
                            nKeterangan = "  SUBTOTAL Non CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nSubTotal = nSubTotal + nTotal
                            nTotal = 0
                        Case 24
                            nUrut = 24
                            nKeterangan = "Jumlah Unit THS "
                            rsTemp.Find "KETERANGAN = 'THS'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                        Case 25
                            ' JUMLAH TOTAL UNIT
                            nUrut = 25
                            nKeterangan = "  JUMLAH TOTAL UNIT "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 2)
                            End If
                            nSubTotal = 0
                            nTotal = 0
                        Case 26
                            nUrut = 26
                            nKeterangan = "Jumlah Pelanggan Booking yang Booking pada hari yang sama "
                            rsTemp.Find "KETERANGAN = 'Booking Hari Sama'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                        Case 27
                            nUrut = 27
                            nKeterangan = "Labor "
                            rsTemp.Find "KETERANGAN = 'Revenue Jasa'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 28
                            nUrut = 28
                            nKeterangan = "Labor Non Ppn "
                            rsTemp.Find "KETERANGAN = 'Revenue Jasa Non ppn'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 29
                            nUrut = 29
                            nKeterangan = "Part Toyota & Oli TGMO "
                            rsTemp.Find "KETERANGAN = 'Revenue Part dan Oli TGMO'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 30
                            nUrut = 30
                            nKeterangan = "Oli / Bahan Non Toyota "
                            rsTemp.Find "KETERANGAN = 'Revenue Oli dan Bahan Non Toyota'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 31
                            nUrut = 31
                            nKeterangan = "Pekerjaan Luar "
                            rsTemp.Find "KETERANGAN = 'Revenue Pekerjaan Luar'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                        Case 32
                            ' SUB TOTAL REVENUE CPUS
                            nUrut = 32
                            nKeterangan = "  SUBTOTAL Revenue CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nSubTotal = 0
                            nTotal = 0
                        Case 33
                            nUrut = 33
                            nKeterangan = "Total Servis Berkala Eks Revenue "
                            rsTemp.Find "KETERANGAN = 'Revenue SBE'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                        Case 34
                            nUrut = 34
                            nKeterangan = "Jam Tersedia (Biasa & Lembur) "
                            rsTemp.Find "Keterangan = 'Jam Tersedia'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                        Case 35
                            nUrut = 35
                            nKeterangan = "Jam Terpakai "
                            rsTemp.Find "Keterangan = 'Jam Terpakai'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                        Case 36
                            nUrut = 36
                            nKeterangan = "Jam Terjual "
                            rsTemp.Find "Keterangan = 'Jam Terjual'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                    End Select
                    nI = nI + 1
                    Call objO.UpdateData(sSQL)
                Loop
            End If
            rsTemp.Close
        Next
        
        Call Cetak
    Else
        MsgBox pesan, vbCritical, "KELENGKAPAN"
        Exit Sub
    End If
End Sub

Private Sub Cetak()
Dim oRs As Object
Dim Report As New RptRepoHOBulanan
Screen.MousePointer = vbHourglass
    

    sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY urut"
    Set oRs = objO.FetchDataHO(sSQL)
    Report.Database.SetDataSource oRs
    
    Report.txtPerusahaan.SetText IIf(Trim(text(6)) <> "", Trim(text(7)), "ALL DEALER") 'gPerush
    Report.txtTahun.SetText CDbl(text(0))
    Report.txtTanggal.SetText Mid(cboBulan1, InStr(cboBulan1, "| ") + 2, Len(cboBulan1) - InStr(cboBulan1, "| ") + 2) & CDbl(text(0)) & " - " & Mid(cboBulan2, InStr(cboBulan2, "| ") + 2, Len(cboBulan2) - InStr(cboBulan2, "| ") + 2) & CDbl(text(1))
    
    Call bersihBulan(Report)
    j = CDbl(Left(cboBulan1, InStr(cboBulan1, " |") - 1))
    For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
        Select Case i - (j - 1)
            Case 1
                Report.txtBln1.SetText GetNamaBulan(i)
            Case 2
                Report.txtBln2.SetText GetNamaBulan(i)
            Case 3
                Report.txtBln3.SetText GetNamaBulan(i)
            Case 4
                Report.txtBln4.SetText GetNamaBulan(i)
            Case 5
                Report.txtBln5.SetText GetNamaBulan(i)
            Case 6
                Report.txtBln6.SetText GetNamaBulan(i)
            Case 7
                Report.txtBln7.SetText GetNamaBulan(i)
            Case 8
                Report.txtBln8.SetText GetNamaBulan(i)
            Case 9
                Report.txtBln9.SetText GetNamaBulan(i)
            Case 10
                Report.txtBln10.SetText GetNamaBulan(i)
            Case 11
                Report.txtBln11.SetText GetNamaBulan(i)
            Case 12
                Report.txtBln12.SetText GetNamaBulan(i)
        End Select
    Next
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub bersihBulan(rpt As Object)
    
    rpt.txtBln1.SetText ""
    rpt.txtBln2.SetText ""
    rpt.txtBln3.SetText ""
    rpt.txtBln4.SetText ""
    rpt.txtBln5.SetText ""
    rpt.txtBln6.SetText ""
    rpt.txtBln7.SetText ""
    rpt.txtBln8.SetText ""
    rpt.txtBln9.SetText ""
    rpt.txtBln10.SetText ""
    rpt.txtBln11.SetText ""
    rpt.txtBln12.SetText ""
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
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
            Case 6
                gMainKeyCari = text(Index)
                gBantuan = "INTERNAL"
                gSearchOption = "DEALER"
                frmCariHO.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 6
            If text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & text(Index) & "' AND INTERNAL = 1")
                If oRs.EOF Then
                    gMainKeyCari = text(Index)
                    gBantuan = "INTERNAL"
                    gSearchOption = "DEALER"
                    frmCariHO.Show (vbModal)
                Else
                    text(Index + 1) = oRs("FS_NM_DEALER")
                End If
            End If
    End Select
End Sub
