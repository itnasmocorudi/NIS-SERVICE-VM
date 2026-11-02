VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmMapping 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Mapping Customer"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   12135
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8505
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   540
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrintOut 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13950
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   540
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   10830
      TabIndex        =   19
      Top             =   570
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H000000FF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   540
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7020
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   540
      Width           =   1440
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   9720
      TabIndex        =   22
      Top             =   555
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      BTYPE           =   8
      TX              =   "Printer :"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRepoMapping.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   8880
      TabIndex        =   7
      Top             =   120
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00BFE8EA&
      Height          =   1935
      Left            =   480
      TabIndex        =   9
      Top             =   600
      Width           =   6735
      Begin VB.ComboBox thMobil 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3500
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox thMobil 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00BFE8EA&
         Height          =   1215
         Left            =   2040
         TabIndex        =   12
         Top             =   600
         Width           =   4575
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00BFE8EA&
            Caption         =   "1 Model"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   120
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00BFE8EA&
            Caption         =   "Beberapa Model"
            Height          =   255
            Index           =   1
            Left            =   1470
            TabIndex        =   2
            Top             =   120
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox4 
            Height          =   270
            Left            =   2070
            TabIndex        =   13
            Top             =   480
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
            Index           =   8
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   2190
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3345
         TabIndex        =   15
         Top             =   225
         Width           =   105
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Model :"
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
         Left            =   885
         TabIndex        =   11
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Kendaraan :"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.CommandButton cmdCetak 
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   0
      Picture         =   "frmRepoMapping.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mapping Customer"
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
      Left            =   600
      TabIndex        =   8
      Top             =   180
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   4080
      X2              =   7080
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   3840
      X2              =   7080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As Object

Private Sub isiCombo()
          Dim rsSearch As ADODB.Recordset
          
1        On Error GoTo isiCombo_Error

2         For i = 0 To 1
3             thMobil(i).Clear
4             Set rsSearch = objO.FetchData("SELECT DISTINCT FS_TH_MOBIL FROM CUSTDAT WITH(NOLOCK) WHERE FS_TH_MOBIL > '1900' ORDER BY FS_TH_MOBIL")
5             If Not rsSearch.EOF Then
6                 Do While Not rsSearch.EOF
7                     thMobil(i).AddItem Trim(rsSearch("FS_TH_MOBIL"))
8                     rsSearch.MoveNext
9                 Loop
10            End If
               
11            Set rsSearch = Nothing
12        Next i

13       On Error GoTo 0
14       Exit Sub

isiCombo_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiCombo of Form frmMapping")
End Sub

Private Sub hapusForm()
    Call isiCombo
    opVehicle(0).value = True
    Text(7) = ""
    Text(8) = ""
    Text(14) = ""
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub

Private Sub cmdCetak_Click()
          Dim rsRepo As ADODB.Recordset
          Dim Txt As String
1        On Error GoTo cmdCetak_Click_Error

2         Set Report = New RptMapping1
          
3         Txt = ""
4         If Trim(thMobil(0)) <> "" And Trim(thMobil(1)) <> "" Then
5             Txt = " AND FS_TH_MOBIL >= '" & Trim(thMobil(0)) & "' AND FS_TH_MOBIL <= '" & Trim(thMobil(1)) & "'"
6         Else
7             Txt = ""
8         End If
9         If Trim(Text(7)) <> "" Or Trim(Text(14)) <> "" Then
10            If opVehicle(0) Then
11                Txt = Txt + " AND MDL = '" & Trim(Text(7)) & "'"
12            ElseIf opVehicle(1) Then
13                Txt = Txt + " AND FMDL LIKE '%" & Trim(Text(14)) & "%'"
14            End If
15        End If
          
      '''    sSQL = "SELECT DISTINCT NO_CHASSIS, NO_POLISI, FS_TH_MOBIL, MDL, FMDL, KELIP,  CASE WHEN FS_KETERANGAN <> '1.000' THEN 'SBE ' ELSE 'SBI ' END + FS_KETERANGAN AS SERVICE, " _
      '''         & "(SELECT TOP 1 FS_RC FROM HISTORY WHERE NO_CHASSIS = A.NO_CHASSIS AND FS_RC <> '' ORDER BY FS_RC DESC) AS FS_RC, " _
      '''         & "(SELECT TOP 1 DATE FROM HISTORY WHERE NO_CHASSIS = A.NO_CHASSIS AND FS_RC <> '' ORDER BY FS_RC DESC) AS DATE, " _
      '''         & "(SELECT TOP 1 FS_KD_DEALER FROM HISTORY WHERE NO_CHASSIS = A.NO_CHASSIS AND FS_RC <> '' ORDER BY FS_RC DESC) AS FS_KD_DEALER, " _
      '''         & "(SELECT TOP 1 FS_NM_DEALER FROM HISTORY WHERE NO_CHASSIS = A.NO_CHASSIS AND FS_RC <> '' ORDER BY FS_RC DESC) AS FS_NM_DEALER " _
      '''         & "FROM CUSTDAT A LEFT JOIN T_REGULAR B ON B.FS_RC = (SELECT TOP 1 FS_RC FROM HISTORY WHERE NO_CHASSIS = A.NO_CHASSIS AND FS_RC <> '' ORDER BY FS_RC DESC) " _
      '''         & "WHERE ISNULL((SELECT TOP 1 FS_RC FROM HISTORY WHERE NO_CHASSIS = A.NO_CHASSIS AND FS_RC <> '' ORDER BY FS_RC DESC),'') <> '' " & txt & " ORDER BY NO_CHASSIS"

16        sSQL = "SELECT DISTINCT Z.NO_CHASSIS, Z.NO_POLISI, FS_TH_MOBIL, MDL, FMDL, " & _
                 "(SELECT TOP 1 FS_KETERANGAN  FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '10K'  AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) + '  ' + " & _
                 "CONVERT(varchar(12), (SELECT TOP 1 DATE FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '10K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC), 103) + '  ' + " & _
                 "(SELECT TOP 1 FS_NM_DEALER FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '10K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) AS '10K', " & _
                 "(SELECT TOP 1 FS_KETERANGAN  FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '20K'  AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) + '  ' + " & _
                 "CONVERT(varchar(12), (SELECT TOP 1 DATE FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '20K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC), 103) + '  ' + " & _
                 "(SELECT TOP 1 FS_NM_DEALER FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '20K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) AS '20K' , " & _
                 "(SELECT TOP 1 FS_KETERANGAN  FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '40K'  AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) + '  ' + " & _
                 "CONVERT(varchar(12), (SELECT TOP 1 DATE FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '40K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC), 103) + '  ' + " & _
                 "(SELECT TOP 1 FS_NM_DEALER FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '40K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) AS '40K', " & _
                 "(SELECT TOP 1 FS_KETERANGAN  FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '80K'  AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) + '  ' + " & _
                 "CONVERT(varchar(12), (SELECT TOP 1 DATE FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '80K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC), 103) + '  ' + " & _
                 "(SELECT TOP 1 FS_NM_DEALER FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '80K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) AS '80K', " & _
                 "(SELECT TOP 1 FS_KETERANGAN  FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '100K'  AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) + '  ' + " & _
                 "CONVERT(varchar(12), (SELECT TOP 1 DATE FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '100K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC), 103) + '  ' + " & _
                 "(SELECT TOP 1 FS_NM_DEALER FROM HISTORY A WITH(NOLOCK) INNER JOIN T_REGULAR B WITH(NOLOCK) ON A.FS_RC = B.FS_RC WHERE KELIP = '100K' AND A.NO_CHASSIS = Z.NO_CHASSIS ORDER BY DATE DESC) AS '100K' " & _
                 "FROM CUSTDAT Z WITH(NOLOCK) WHERE Z.NO_CHASSIS <> '' " & Txt & " ORDER BY Z.NO_CHASSIS"
           
17        If Log.RPLHistori = 1 Then
18            sSQL = UCase(sSQL)
19            sSQL = Replace(sSQL, "T_REGULAR", "[" & dbReplikasiHO & "]..T_REGULAR")
20        End If
          
21        Set rsRepo = objO.FetchData(sSQL)
22        Report.database.SetDataSource rsRepo
23        Report.txtPerusahaan.SetText gPerush
24        Report.txtAlamat.SetText gAlamat
25        Report.txtTanggal.SetText "Tanggal : " & TglDDMMMMYYYY(Format(isnow, "DD/MM/YYYY"), 1)

26        With CRViewer1
27            .ReportSource = Report
28            .ViewReport
29            .Zoom 100
30            .Top = 0
31            .Left = Image1.Left + Image1.Width + 20
32            .Width = Me.Width - 500
33            .Height = Image1.Height - 500
              
34            cmdTutup.Top = .Top + 430
35            cmdTutup.Left = .Left
36            cmdTutup.Visible = True
              
37            cmdSettingPrinter.Top = cmdTutup.Top
38            cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 20
39            cmdSettingPrinter.Visible = True
40            .Visible = True
41        End With
42        cmdTutup.Visible = True
43        Screen.MousePointer = vbDefault
44        Call DisableExport
45        Call setLokasiTombol
46       On Error GoTo 0
47       Exit Sub

cmdCetak_Click_Error:

48        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetak_Click of Form frmMapping")
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
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub Form_Resize()
    Image1.Height = Me.Height - 50
End Sub

Private Sub opVehicle_Click(Index As Integer)
    If opVehicle(0) Then
        Text(7).Visible = True
        Text(8).Visible = True
        MaskEdBox4.Visible = True
        Text(14).Visible = False
    Else
        Text(7).Visible = False
        Text(8).Visible = False
        MaskEdBox4.Visible = False
        Text(14).Visible = True
        Text(14).Left = Text(7).Left
        Text(14).SelLength = Len(Text(14))
        Text(14).SetFocus
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 7
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Dim rsSearch As ADODB.Recordset
     
    Select Case Index
        Case 7
            If Text(Index) <> "" Then
            
                
                If Log.RPLHistori = 0 Then
                
                    Set rsSearch = objO.FetchData("SELECT * FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                Else
                    Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                End If
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCari.Show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("DESC")
                End If
                 
            Else
                Text(Index + 1) = ""
            End If
    End Select
    Set rsSearch = Nothing
End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 445
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 445
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 445
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 445
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 445
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdPrintOut.Top = CRViewer1.Top + 445
    cmdPrintOut.Left = cboPrinter.Left + cboPrinter.Width + 30
End Sub

Private Sub DisableExport()
    Dim objPrinter As Printer
    Set objPrinter = Printer
    If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
        cboPrinter = ""
    Else
        cboPrinter = objPrinter.DeviceName
    End If
    cboPrinter.Clear
    
    For Each pr In Printers
        If InStr(UCase(pr.DeviceName), "PDF") = 0 Then cboPrinter.AddItem pr.DeviceName
    Next pr
    If Log.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        CRViewer1.EnablePrintButton = True
        cmdExport.Visible = False
        cboPrinter.Visible = False
        cmdPrintOut.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdPrintOut_Click()
    For Each pr In Printers
        If UCase(pr.DeviceName) = UCase(cboPrinter) Then
            GoTo Cetak
        End If
    Next pr
    MsgBox "Printer tidak ditemukan Pilih Printer dahulu", vbInformation, "Proses Cetak Gagal"
    cboPrinter.SetFocus
    Exit Sub
    
Cetak:
    Dim Psize As Integer
    Dim POr As Integer
    
    Psize = Report.PaperSize
    POr = Report.PaperOrientation
    
    Set Printer = Printers(cboPrinter.ListIndex)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    Report.PaperOrientation = POr
    Report.PaperSize = Psize
    Report.PrintOut
    
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Salah
    Dim NamaFile As String
    NamaFile = ""
    
    Dim f As New FormExplorer
    f.Show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    
    If Log.PdfOnly = 1 Then
         
        NamaFile = MyPath & f.FileName
        
        If UCase(Right(Trim(f.FileName), 5)) = ".XLSX" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 5)
        If UCase(Right(Trim(f.FileName), 4)) = ".XLS" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 4)
         
        
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
        Report.ExportOptions.RTFExportAllPages = True
        Report.ExportOptions.FormatType = crEFTExactRichText
        Report.Export False
        MsgBox NamaFile & ".Doc" & "  ... Data sudah di Export !", vbInformation
        Exit Sub
        With msg
            .PassingMsg = NamaFile & ".Doc" & "  ... Data sudah di Export !"
            .passingTipe = 2
            .Show vbModal
        End With
        Exit Sub
    ElseIf onExpo = 0 Then
        Report.Export
    End If
    Exit Sub
Salah:
    If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
    MsgBox "Err.Line : " & Erl & "Err.Numb : " & err.Number & "Err.Desc. : " & err.Description & vbCr & "Procedure : " & Me.Name & " export"
Keluar:
End Sub



