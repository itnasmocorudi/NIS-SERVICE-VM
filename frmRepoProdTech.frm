VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmRepoProdTech 
   Caption         =   "Laporan Produktivitas Teknisi"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   8400
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   11
      Top             =   4800
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
      EnableRefreshButton=   -1  'True
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
      Left            =   4080
      TabIndex        =   8
      Top             =   3960
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
      Left            =   6000
      TabIndex        =   9
      Top             =   3960
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   480
      TabIndex        =   12
      Top             =   480
      Width           =   7455
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   2
         Top             =   840
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2400
         TabIndex        =   13
         Top             =   840
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
         Index           =   2
         Left            =   1530
         TabIndex        =   3
         Top             =   1200
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   2400
         TabIndex        =   14
         Top             =   1200
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
         Index           =   1
         Left            =   2520
         TabIndex        =   15
         Top             =   840
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   16
         Top             =   1200
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   1530
         TabIndex        =   4
         Top             =   1560
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   2400
         TabIndex        =   17
         Top             =   1560
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
         Left            =   1530
         TabIndex        =   5
         Top             =   1920
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
         Left            =   3480
         TabIndex        =   18
         Top             =   1920
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
         Index           =   5
         Left            =   2520
         TabIndex        =   19
         Top             =   1560
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1530
         TabIndex        =   20
         Top             =   2280
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
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   51380225
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   1530
         TabIndex        =   6
         Top             =   2640
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   3375
         TabIndex        =   27
         Top             =   2640
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   3375
         TabIndex        =   29
         Top             =   3000
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
         Index           =   9
         Left            =   1530
         TabIndex        =   7
         Top             =   3000
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proses"
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
         Left            =   960
         TabIndex        =   28
         Top             =   2685
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   7440
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3240
         TabIndex        =   26
         Top             =   240
         Width           =   60
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
         Left            =   870
         TabIndex        =   25
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Kendaraan"
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
         Left            =   195
         TabIndex        =   24
         Top             =   1965
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SA"
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
         Left            =   1245
         TabIndex        =   23
         Top             =   1605
         Width           =   195
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mechanic"
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
         Left            =   780
         TabIndex        =   22
         Top             =   1245
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Foreman"
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
         Left            =   810
         TabIndex        =   21
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return Job"
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
         Left            =   645
         TabIndex        =   30
         Top             =   3000
         Width           =   795
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   4800
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   8467
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Laporan Produktivitas Mekanik"
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
End
Attribute VB_Name = "frmRepoProdTech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 25 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : -
'=======================================

Dim rsTemp As New ADODB.Recordset
Dim Report As New RptRepoJmlUnit
Dim vNamaTabel As String

Private Sub hapusForm()
    For i = 0 To 8
        Text(i) = ""
    Next
    Tgl1 = gTglServer
    Tgl2 = gTglServer
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As New ADODB.Recordset

    createTemp = False
    sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
         & " drop table [dbo].[" & vNamaTabel & "]"
    oConn.Execute sSQL

    sSQL = "CREATE TABLE [dbo].[" & vNamaTabel & "] ( " _
         & "[TANGGAL][DATETIME] NOT NULL DEFAULT '01/01/3000', " _
         & "[SA][CHAR](5) NOT NULL DEFAULT '', " _
         & "[FOREMAN][CHAR](5) NOT NULL DEFAULT '', " _
         & "[MEKANIK][CHAR](5) NOT NULL DEFAULT '', " _
         & "[RATEJ][DECIMAL](18,2) NOT NULL DEFAULT 0, " _
         & "[ACTRATE][DECIMAL](18,2) NOT NULL DEFAULT 0, " _
         & "[FN_RETURN][NUMERIC](18) NOT NULL DEFAULT 0, " _
         & "[PROSES][VARCHAR](10) NOT NULL DEFAULT '', " _
         & "[STALL][VARCHAR](10) NOT NULL DEFAULT '', " _
         & "[TERSEDIA][DECIMAL](18,2)) ON [PRIMARY]"
    oConn.Execute sSQL
    
    If rsHlp.State Then rsHlp.Close
    Set rsHlp = oConn.Execute("select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1")
    If Not rsHlp.EOF Then
        createTemp = True
    End If
End Function

Private Sub IsiTempRepo(ket As String, vNamaTabel As String, txt As String)
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim rsHlp As New ADODB.Recordset
Dim btkTabel As Boolean
Dim nGroup As String
    
    If ket <> "" Then
        ket = "AND " & ket
    End If
    ket = ket & " ORDER BY E.[DATE], NO_WO "
    
'    If rsTemp.State Then rsTemp.Close
'    Set rsTemp = oConn.Execute("SELECT [DATE], FRONT, MEKANIK1, MEKANIK2, MEKANIK3, MEKANIK4, MEKANIK5, MEKANIK6, MEKANIK7, MEKANIK8, " _
'               & "RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, ACTRATE1, ACTRATE2, ACTRATE3, ACTRATE4, ACTRATE5, ACTRATE6, ACTRATE7, ACTRATE8, " _
'               & "FN_RETURN, (CASE RPR_COND WHEN '' THEN 'WAITING' WHEN 'P' THEN 'PROCESS' WHEN 'S' THEN 'FINISIH' WHEN 'K' THEN 'CLOSING' END) AS PROSES, NOSTALL FROM JOBDTA A LEFT JOIN STALL B ON REPLACE(SUBSTRING(B.KETERANGAN, CHARINDEX('|',B.KETERANGAN)+1,8),'|','') = A.NO_WO " _
'               & "AND B.TANGGAL = A.[DATE] WHERE DELIV <> 'BTL' AND [DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' " & ket & "")
'    If Not rsTemp.EOF Then
'        btkTabel = createTemp(vNamaTabel)
'        Do While Not rsTemp.EOF
'            For i = 1 To 8
'                hlp = "MEKANIK" & i
'                If rsTemp(hlp) <> "" Then
'                    hlp2 = "RTE" & i
'                    hlp3 = "ACTRATE" & i
'                    If rsHlp.State Then rsHlp.Close
'                    Set rsHlp = oConn.Execute("SELECT [GROUP] FROM MEKANIK WHERE MECH = '" & rsTemp(hlp) & "'")
'                    If Not rsHlp.EOF Then
'                        nGroup = rsHlp("GROUP")
'                    End If
'
'                    If rsHlp.State Then rsHlp.Close
'                    Set rsHlp = oConn.Execute("SELECT MENITKERJA FROM T_MEKANIK WHERE MECH = '" & rsTemp(hlp) & "' AND TANGGAL = '" & rsTemp("DATE") & "'")
'                    If Not rsHlp.EOF Then
'                        nMenit = Round(rsHlp("MENITKERJA") / 60, 2)
'                    Else
'                        nMenit = 0
'                    End If
'
'                    If btkTabel Then
'                        oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & Format(TglYYYYMMDD(rsTemp("DATE")), "mm/dd/yyyy") & "', '" & rsTemp("FRONT") & "', '" & nGroup & "', '" & rsTemp(hlp) & "', " _
'                                     & "'" & Replace(rsTemp(hlp2), ",", ".") & "', '" & Replace(rsTemp(hlp3), ",", ".") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("PROSES") & "', '" & rsTemp("NOSTALL") & "', '" & nMenit & "')")
'                    End If
'                End If
'            Next
'            rsTemp.MoveNext
'        Loop

    Dim rteJMech As Double
    If rsTemp.State Then rsTemp.Close
''''    Set rsTemp = oConn.Execute("SELECT NO_WO, [DATE], FRONT, B.MECH AS MEKANIK, B.RATE AS ACTRATE, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, FN_RETURN, " _
''''                             & "(CASE RPR_COND WHEN '' THEN 'WAITING' WHEN 'P' THEN 'PROCESS' WHEN 'S' THEN 'FINISIH' WHEN 'K' THEN 'CLOSING' END) AS PROSES, NOSTALL " _
''''                             & "FROM JOBDTA A INNER JOIN ( " _
''''                             & "SELECT LEFT(KETERANGAN,CHARINDEX(' |', KETERANGAN)) AS NO_POL, MECH, SUM(RATE) AS RATE  FROM HELPPROGRESS " _
''''                             & "GROUP BY MECH, LEFT(KETERANGAN,CHARINDEX(' |', KETERANGAN))) B ON A.NO_POLISI = B.NO_POL " _
''''                             & "LEFT JOIN STALL C ON REPLACE(SUBSTRING(C.KETERANGAN, CHARINDEX('|',C.KETERANGAN)+1,8),'|','') = A.NO_WO " _
''''                             & "ORDER BY [DATE], NO_WO")
    Set rsTemp = oConn.Execute("SELECT NO_WO, [DATE], FRONT, D.SELISIHTOT AS ACTRATE, E.MECH AS MEKANIK, D.SELISIH AS ACTMECH, RTE1+RTE2+RTE3+RTE4+RTE5+RTE6+RTE7+RTE8 AS RATEJ " _
                             & ", FN_RETURN,(CASE RPR_COND WHEN '' THEN 'WAITING' WHEN 'P' THEN 'PROCESS' WHEN 'S' THEN 'FINISIH' WHEN 'K' THEN 'CLOSING' END) AS PROSES, NOSTALL " _
                             & "FROM JOBDTA E INNER JOIN ( " _
                             & "SELECT A.NOORDER, SUM(DATEDIFF(n, A.START,A.PAUSE)) AS SELISIHTOT, B.MECH, B.SELISIH FROM TEMPPROGRESS A LEFT JOIN ( " _
                             & "SELECT NOORDER, MECH, SUM(DATEDIFF(n, START,PAUSE)) AS SELISIH FROM TEMPPROGRESS GROUP BY NOORDER, MECH) B ON A.NOORDER = B.NOORDER " _
                             & "GROUP BY A.NOORDER, B.MECH, B.SELISIH) D ON E.NO_WO = D.NOORDER AND D.SELISIHTOT > 0 " _
                             & "LEFT JOIN STALL C ON REPLACE(SUBSTRING(C.KETERANGAN, CHARINDEX('|',C.KETERANGAN)+1,8),'|','') = E.NO_WO " _
                             & "WHERE DELIV <> 'BTL' AND [DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' " & ket & "")
    If Not rsTemp.EOF Then
        btkTabel = createTemp(vNamaTabel)
        Do While Not rsTemp.EOF
            rteJMech = Round((CDbl(rsTemp("ACTMECH")) / CDbl(rsTemp("ACTRATE"))) * CDbl(rsTemp("RATEJ")), 2)
            If rsHlp.State Then rsHlp.Close
            Set rsHlp = oConn.Execute("SELECT [GROUP] FROM MEKANIK WHERE MECH = '" & rsTemp("MEKANIK") & "'")
            If Not rsHlp.EOF Then
                nGroup = rsHlp("GROUP")
            End If

            If rsHlp.State Then rsHlp.Close
            Set rsHlp = oConn.Execute("SELECT MENITKERJA FROM T_MEKANIK WHERE MECH = '" & rsTemp("MEKANIK") & "' AND TANGGAL = '" & rsTemp("DATE") & "'")
            If Not rsHlp.EOF Then
                nMenit = Round(rsHlp("MENITKERJA") / 60, 2)
            Else
                nMenit = 0
            End If

            If btkTabel Then
                oConn.Execute ("INSERT INTO [" & vNamaTabel & "] VALUES ('" & Format(TglYYYYMMDD(rsTemp("DATE")), "mm/dd/yyyy") & "', '" & rsTemp("FRONT") & "', '" & nGroup & "', '" & rsTemp("MEKANIK") & "', " _
                             & "'" & Replace(CDbl(rteJMech), ",", ".") & "', '" & Replace(Round(CDbl(rsTemp("ACTMECH")) / 60, 2), ",", ".") & "', '" & rsTemp("FN_RETURN") & "', '" & rsTemp("PROSES") & "', '" & rsTemp("NOSTALL") & "', '" & Replace(nMenit, ",", ".") & "')")
            End If
            rsTemp.MoveNext
        Loop
        Call Cetak(vNamaTabel, txt)
    Else
        MsgBox "Tidak ada data !!"
        Exit Sub
    End If
End Sub

Private Sub Cetak(vNamaTabel As String, txt As String)
Dim oRs As Object
Dim Report As New RptRepoProdTech
Screen.MousePointer = vbHourglass
    
    If txt <> "" Then
        txt = " WHERE " & txt
    End If
    '
    sSQL = "SELECT TANGGAL, SA, FOREMAN, B.[NAME], MEKANIK, SUM(ISNULL(RATEJ,0)) AS RATEJ, ISNULL(ACTRATE,0) AS ACTRATE, FN_RETURN, PROSES, STALL, TERSEDIA, ISNULL([NAME], '') AS NAMA " _
         & ", ROUND(SUM((RATEJ/ACTRATE)*100),0) AS EFISIENSI, ROUND(SUM((RATEJ/TERSEDIA)*100),0) AS OVERALL, ROUND(SUM((ACTRATE/TERSEDIA)*100),0) AS UTILISASI " _
         & "FROM " & vNamaTabel & " A LEFT JOIN MEKANIK B ON A.FOREMAN = B.MECH " & txt & " " _
         & "GROUP BY TANGGAL, SA, FOREMAN, B.[NAME], MEKANIK, FN_RETURN, PROSES, STALL, TERSEDIA, [NAME], ACTRATE " _
         & "ORDER BY TANGGAL"
    Set oRs = oConn.Execute(sSQL)
    Report.Database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
   
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdTutup_Click()
    sSQL = "DELETE [" & vNamaTabel & "]"
    oConn.Execute sSQL
    CRViewer1.Visible = False
    cmdTutup.Visible = False
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

Private Sub Form_Load()
    Call hapusForm
    vNamaTabel = "ProdTech" & DriveSerialNumber("c:")
End Sub

Private Sub TblKeluar_Click()
    sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
         & " drop table [dbo].[" & vNamaTabel & "]"
    oConn.Execute sSQL
    Unload Me
End Sub

Private Sub TblSimpan_Click()
Dim boleh As Boolean
Dim TxtTemp As String
Dim txt As String
    
    boleh = False
    TxtTemp = ""
    txt = ""
    If Text(0) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "FOREMAN = '" & Text(0) & "' "
    End If
    If Text(2) <> "" Then
        boleh = True
        If txt <> "" Then
            txt = txt & "AND "
        End If
        txt = txt & "MEKANIK = '" & Text(2) & "' "
    End If
    If Text(4) <> "" Then
        boleh = True
        If TxtTemp <> "" Then
            TxtTemp = TxtTemp & "AND "
        End If
        TxtTemp = TxtTemp & "FRONT = '" & Text(4) & "' "
    End If
    If Text(6) <> "" Then
        boleh = True
        If TxtTemp <> "" Then
            TxtTemp = TxtTemp & "AND "
        End If
        TxtTemp = TxtTemp & "MDL = '" & Text(6) & "' "
    End If
    If Text(8) <> "" Then
        boleh = True
        If TxtTemp <> "" Then
            TxtTemp = TxtTemp & "AND "
        End If
        If Text(8) = "PROSES" Then
            TxtTemp = TxtTemp & "RPR_COND = 'P'"
        ElseIf Text(8) = "WAITING" Then
            TxtTemp = TxtTemp & "RPR_COND = ''"
        ElseIf Text(8) = "FINISH" Then
            TxtTemp = TxtTemp & "RPR_COND = 'S'"
        ElseIf Text(8) = "CLOSING" Then
            TxtTemp = TxtTemp & "RPR_COND = 'K'"
        End If
    End If
    If Text(9) <> "" Then
        boleh = True
        If TxtTemp <> "" Then
            TxtTemp = TxtTemp & "AND "
        End If
        If Text(9) = "RETURN JOB" Then
            TxtTemp = TxtTemp & "FN_RETURN = '1'"
        ElseIf Text(9) = "NON RETURN JOB" Then
            TxtTemp = TxtTemp & "FN_RETURN = '0'"
        End If
    End If
    
    Call IsiTempRepo(TxtTemp, vNamaTabel, txt)
End Sub

Private Sub text_GotFocus(Index As Integer)
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
'    Text(Index).SelLength = Len(Text(Index))
'    Text(Index).SetFocus
End Sub

Private Sub text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "GROUP_MEKANIK"
                frmCari.Show (vbModal)
            Case 2
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ALL"
                gBantuan4 = ""
                frmCari.Show (vbModal)
            Case 4
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.Show (vbModal)
            Case 6
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.Show (vbModal)
            Case 8
                gMainKeyCari = Text(Index)
                gSearchOption = "PROSES"
                frmCari.Show (vbModal)
            Case 9
                gMainKeyCari = Text(Index)
                gSearchOption = "RETURN"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
'    If Text(Index) <> "" Then
        Select Case Index
            Case 0
                If Text(Index) <> "" Then
                    sSQL = "SELECT * FROM MEKANIK WHERE [GROUP] = '" & Text(Index) & "' AND MECH = '" & Text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "GROUP_MEKANIK"
                        frmCari.Show (vbModal)
                    Else
                        Text(Index + 1) = oRs("NAME")
                    End If
                    oRs.Close
                Else
                    Text(Index + 1) = ""
                End If
            Case 2
                If Text(Index) <> "" Then
                    sSQL = "SELECT * FROM MEKANIK WHERE MECH = '" & Text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "MEKANIK"
                        gBantuan = "ALL"
                        gBantuan4 = ""
                        frmCari.Show (vbModal)
                    Else
                        Text(Index + 1) = oRs("NAME")
                    End If
                    oRs.Close
                Else
                    Text(Index + 1) = ""
                End If
            Case 4
                If Text(Index) <> "" Then
                    sSQL = "SELECT * FROM T_PETUGAS WHERE FS_KD_PETUGAS = '" & Text(Index) & "' AND TIPE = 'SA'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "SA"
                        frmCari.Show (vbModal)
                    Else
                        Text(Index + 1) = oRs("FS_NM_PETUGAS")
                    End If
                    oRs.Close
                Else
                    Text(Index + 1) = ""
                End If
            Case 6
                If Text(Index) <> "" Then
                    sSQL = "SELECT * FROM MODEL WHERE MDL = '" & Text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "MDL"
                        frmCari.Show (vbModal)
                    Else
                        Text(Index + 1) = oRs("DESC")
                    End If
                    oRs.Close
                Else
                    Text(Index + 1) = ""
                End If
            Case 8
                If Text(Index) <> "" Then
                    sSQL = "SELECT PROSES FROM T_REPO WHERE PROSES = '" & Text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "PROSES"
                        frmCari.Show (vbModal)
                    End If
                    oRs.Close
                End If
            Case 9
                If Text(Index) <> "" Then
                    sSQL = "SELECT RETUR FROM T_REPO WHERE RETUR = '" & Text(Index) & "'"
                    oRs.Open sSQL, oConn, adOpenStatic
                    If oRs.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "RETURN"
                        frmCari.Show (vbModal)
                    End If
                    oRs.Close
                End If
        End Select
'    End If
End Sub
