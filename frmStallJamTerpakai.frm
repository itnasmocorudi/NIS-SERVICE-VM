VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmStallJamTerpakai 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Jam Terpakai"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   15645
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
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Frame framSearch 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Display"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton searchClose 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Tutup"
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin TDBDate6Ctl.TDBDate DtTanggal 
         Height          =   330
         Left            =   2295
         TabIndex        =   5
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmStallJamTerpakai.frx":0000
         Caption         =   "frmStallJamTerpakai.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmStallJamTerpakai.frx":0198
         Keys            =   "frmStallJamTerpakai.frx":01B6
         Spin            =   "frmStallJamTerpakai.frx":0214
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
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Display"
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
         Left            =   600
         TabIndex        =   6
         Top             =   525
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Ganti Hari"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid GridWaktu 
      Height          =   8625
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   15214
      _Version        =   393216
      Cols            =   6
      HighLight       =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid GridWaktu 
      Height          =   8625
      Index           =   1
      Left            =   7800
      TabIndex        =   8
      Top             =   0
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   15214
      _Version        =   393216
      Cols            =   6
      HighLight       =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStallJamTerpakai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTotal As Integer
Dim nLimit As Integer

Private Sub cmdDisplay_Click()
    Call TampilWaktu(Format(DtTanggal, "DD/MM/YYYY"))
    framSearch.Visible = False
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub CmdSearch_Click()
    framSearch.Visible = True
End Sub

Private Sub Form_Load()
    nLimit = 30
    Call TampilWaktu(Format(isnow, "DD/MM/YYYY"))
    framSearch.Visible = False
End Sub

Private Sub searchClose_Click()
    framSearch.Visible = False
End Sub

Private Sub TampilWaktu(Tggl As Date)
      Dim rsWaktu As ADODB.Recordset
      Dim nStall As String
      Dim i As Integer

1         For i = 0 To 1
2             GridWaktu(i).TextMatrix(0, 0) = "..."
3             GridWaktu(i).TextMatrix(0, 1) = "STALL"
              
              'FUNGSI MENGHASILKAN WAKTU TERSEDIA BERDASARKAN TGL TTN (HARI H)
4             If DateDiff("d", isnow, Tggl) >= 2 Then
5                 If year(Tggl) <= 2017 Then
6                     Call objO.UpdateData("EXEC sp_lihatStall2017 '" & Format(Tggl, "mm/dd/yyyy") & "'")
7                 Else
8                     Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl, "mm/dd/yyyy") & "'")
9                 End If
10                GridWaktu(i).TextMatrix(0, 2) = "H - 2"
11                GridWaktu(i).TextMatrix(0, 3) = "H - 1"
12                GridWaktu(i).TextMatrix(0, 4) = "  H  "
13                GridWaktu(i).TextMatrix(0, 5) = "H + 1"
14            ElseIf DateDiff("d", isnow, Tggl) = 1 Then
15                If year(Tggl) <= 2017 Then
16                    Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 1, "mm/dd/yyyy") & "'")
17                Else
18                    Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 1, "mm/dd/yyyy") & "'")
19                End If
                  
20                GridWaktu(i).TextMatrix(0, 2) = "H - 1"
21                GridWaktu(i).TextMatrix(0, 3) = "  H  "
22                GridWaktu(i).TextMatrix(0, 4) = "H + 1"
23                GridWaktu(i).TextMatrix(0, 5) = "H + 2"
24            ElseIf DateDiff("d", isnow, Tggl) <= 0 Then
25                If year(Tggl) <= 2017 Then
26                    Call objO.UpdateData("EXEC sp_lihatStall2017 '" & Format(Tggl + 2, "mm/dd/yyyy") & "'")
27                Else
28                    Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 2, "mm/dd/yyyy") & "'")
29                End If
                  
30                GridWaktu(i).TextMatrix(0, 2) = "  H  "
31                GridWaktu(i).TextMatrix(0, 3) = "H + 1"
32                GridWaktu(i).TextMatrix(0, 4) = "H + 2"
33                GridWaktu(i).TextMatrix(0, 5) = "H + 3"
34            End If
35        Next
          
          'FUNSGI UNTUK MENAMPILKAN WAKTU TERSEDIA H-2 SAMPAI H+1
          ''Set rsWaktu = objO.FetchData("SELECT * FROM STALLWKF WITH(NOLOCK) ORDER BY LEFT(NOSTALL,1), CASE LEFT(NOSTALL,1) WHEN 'S' THEN CAST(SUBSTRING(NOSTALL, 7, LEN(NOSTALL)-6) AS INTEGER) ELSE CAST(SUBSTRING(NOSTALL, 5, LEN(NOSTALL)-4) AS INTEGER) END, [H-2], [H-1], H, [H+1]")
          '' ----------------------------------------------------------------------
          '' 20181113 | REV | YANG DIAMBIL NOSTALL TANPA STALL FINAL DAN SUBLET
          '' ----------------------------------------------------------------------
36        Set rsWaktu = objO.FetchData("SELECT * FROM STALLWKF WITH(NOLOCK) where NoStall not in ('STALL FINL','STALL SUBL') ORDER BY LEFT(NOSTALL,1), CASE LEFT(NOSTALL,1) WHEN 'S' THEN CAST(SUBSTRING(NOSTALL, 7, LEN(NOSTALL)-6) AS INTEGER) ELSE CAST(SUBSTRING(NOSTALL, 5, LEN(NOSTALL)-4) AS INTEGER) END, [H-2], [H-1], H, [H+1]")
37        If Not rsWaktu.EOF Then
38            i = 1
39            nTotal = rsWaktu.RecordCount
              
40            GridWaktu(0).Rows = IIf(nTotal > nLimit, Val(nTotal / 2), nTotal) + 5
41            GridWaktu(1).Rows = IIf(nTotal > nLimit, Val(nTotal / 2), 0) + 5
42            GridWaktu(0).ColWidth(0) = 200
43            GridWaktu(0).ColWidth(1) = 1000
44            GridWaktu(0).ColWidth(2) = 1500
45            GridWaktu(0).ColWidth(3) = 1500
46            GridWaktu(0).ColWidth(4) = 1500
47            GridWaktu(0).ColWidth(5) = 1500
              
48            GridWaktu(1).ColWidth(0) = 200
49            GridWaktu(1).ColWidth(1) = 1000
50            GridWaktu(1).ColWidth(2) = 1500
51            GridWaktu(1).ColWidth(3) = 1500
52            GridWaktu(1).ColWidth(4) = 1500
53            GridWaktu(1).ColWidth(5) = 1500
54            nStall = ""
              
              
55            Do While Not rsWaktu.EOF
56                If i <= IIf(nTotal > nLimit, Val(nTotal / 2), nTotal) Then
57                    If Trim(rsWaktu("NOSTALL")) = nStall Then
58                        GridWaktu(0).TextMatrix(i, 1) = ""
59                        GridWaktu(0).TextMatrix(i, 2) = Nul(Trim(rsWaktu("H-2")), "")
60                        GridWaktu(0).TextMatrix(i, 3) = Nul(Trim(rsWaktu("H-1")), "")
61                        GridWaktu(0).TextMatrix(i, 4) = Nul(Trim(rsWaktu("H")), "")
62                        GridWaktu(0).TextMatrix(i, 5) = Nul(Trim(rsWaktu("H+1")))
63                    Else
64                        GridWaktu(0).TextMatrix(i, 0) = "*"
65                        GridWaktu(0).TextMatrix(i, 1) = Trim(rsWaktu("NOSTALL"))
66                        GridWaktu(0).TextMatrix(i, 2) = Nul(Trim(rsWaktu("H-2")), "")
67                        GridWaktu(0).TextMatrix(i, 3) = Nul(Trim(rsWaktu("H-1")), "")
68                        GridWaktu(0).TextMatrix(i, 4) = Nul(Trim(rsWaktu("H")), "")
69                        GridWaktu(0).TextMatrix(i, 5) = Nul(Trim(rsWaktu("H+1")))
70                        nStall = Trim(rsWaktu("NOSTALL"))
71                    End If
72                Else
73                    If Trim(rsWaktu("NOSTALL")) = nStall Then
74                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 1) = ""
75                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 2) = Nul(Trim(rsWaktu("H-2")), "")
76                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 3) = Nul(Trim(rsWaktu("H-1")), "")
77                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 4) = Nul(Trim(rsWaktu("H")), "")
78                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 5) = Nul(Trim(rsWaktu("H+1")))
79                    Else
80                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 0) = "*"
81                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 1) = Trim(rsWaktu("NOSTALL"))
82                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 2) = Nul(Trim(rsWaktu("H-2")), "")
83                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 3) = Nul(Trim(rsWaktu("H-1")), "")
84                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 4) = Nul(Trim(rsWaktu("H")), "")
85                        GridWaktu(1).TextMatrix(i + 1 - Val(nTotal / 2), 5) = Nul(Trim(rsWaktu("H+1")))
86                        nStall = Trim(rsWaktu("NOSTALL"))
87                    End If
88                End If
89                i = i + 1
90                rsWaktu.MoveNext
91            Loop
92        End If
End Sub
