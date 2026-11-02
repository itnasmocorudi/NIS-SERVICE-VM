VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoInvRinci 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Rincian Laporan Revenue"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ucXPButton.XPButton XPButton1 
      Height          =   270
      Left            =   4350
      TabIndex        =   25
      Top             =   5730
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   476
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
      MICON           =   "frmRepoInvRincian.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5730
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
      Left            =   1590
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5730
      Width           =   1440
   End
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
      Left            =   3075
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5730
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCetak 
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
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5730
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5400
      TabIndex        =   20
      Top             =   5760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   840
      TabIndex        =   0
      Top             =   5985
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
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1890
   End
   Begin VB.CommandButton TblSimpan 
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
      Height          =   390
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1890
   End
   Begin VB.CommandButton tblRefresh 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Refresh"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   7695
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1605
         TabIndex        =   12
         Top             =   1440
         Width           =   3855
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Beberapa Model"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   14
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton opVehicle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "1 Model"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.OptionButton opTipe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rinci Tipe Job Order"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton opTipe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rinci Tipe Kendaraan"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97386497
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97386497
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1575
         TabIndex        =   15
         Top             =   1785
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
         Left            =   3525
         TabIndex        =   16
         Top             =   1785
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
         Left            =   1575
         TabIndex        =   17
         Top             =   2145
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
         Left            =   1965
         TabIndex        =   18
         Top             =   2145
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
         Left            =   240
         TabIndex        =   19
         Top             =   1470
         Width           =   1245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   6960
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
         Left            =   2640
         TabIndex        =   11
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
         Left            =   270
         TabIndex        =   10
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe"
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
         Left            =   540
         TabIndex        =   9
         Top             =   720
         Width           =   300
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8175
      Left            =   0
      Picture         =   "frmRepoInvRincian.frx":001C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmRepoInvRinci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objBilling As New BClsBilling
Dim Report As Object
Dim vNamaTabel1 As String
Private txtTemp As String
'------------------------------------
Dim U10K As Integer
Dim U20K As Integer
Dim U30K As Integer
Dim U40K As Integer
Dim U50K As Integer
Dim U60K As Integer
Dim U70K As Integer
Dim U80K As Integer
Dim U90K As Integer
Dim U100K As Integer
Dim U5K As Integer
Dim U101K As Integer
Dim UPDS As Integer
Dim USBI As Integer
Dim UGRP As Integer
Dim UTWC As Integer
Dim ULUB As Integer
Dim URTJ As Integer
Dim UTHS As Integer
Dim USBE As Integer

Private Sub clearTemp()
    U10K = "0"
    U20K = "0"
    U30K = "0"
    U40K = "0"
    U50K = "0"
    U60K = "0"
    U70K = "0"
    U80K = "0"
    U90K = "0"
    U100K = "0"
    U5K = "0"
    U101K = "0"
    UPDS = "0"
    USBI = "0"
    UGRP = "0"
    UTWC = "0"
    ULUB = "0"
    URTJ = "0"
    UTHS = "0"
End Sub


Private Sub IsiRinciKM(vNamaTabel As String, txt As String)
      Dim rsTemp As ADODB.Recordset
          Dim tStr
          Dim tStr2
          Dim tStr3
          
1         If Log.RPLHistori = 0 Then
2             Set rsTemp = objO.FetchData("SELECT A.*, CASE FS_RC WHEN '' THEN '0' ELSE FS_RC END AS FS_RC, ISNULL(B.DOUBLEINV,'0') AS NDUOINV, CASE WHEN (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7+KLPKERJA8,'LUB',''))= 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7+KLPKERJA8 <> '') THEN 1 ELSE 0 END AS ISLUB FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN HO_PROGRAM B WITH(NOLOCK) ON A.KODETIPE = B.ID_PROGRAM WHERE INVDATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND DELIV <> 'BTL' AND A.FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B' " & txtTemp)
3         Else
4             Set rsTemp = objO.FetchData("SELECT A.*, CASE FS_RC WHEN '' THEN '0' ELSE FS_RC END AS FS_RC, ISNULL(B.DOUBLEINV,'0') AS NDUOINV, CASE WHEN (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7+KLPKERJA8,'LUB',''))= 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7+KLPKERJA8 <> '') THEN 1 ELSE 0 END AS ISLUB FROM T_INVOICE A WITH(NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM B WITH(NOLOCK) ON A.KODETIPE = B.ID_PROGRAM AND B.FS_LOG_DEL = '' WHERE INVDATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' AND DELIV <> 'BTL' AND A.FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B' " & txtTemp)
5         End If
6         If Not rsTemp Then
7             Call BuatTabelRepoInv(vNamaTabel)
8             If createTemp(vNamaTabel) Then
9                 Do While Not rsTemp.EOF
                      '#########################################################'
                      '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                      '#########################################################'

10                    If ((Format(rsTemp("INVDATE"), "MM/DD/YYYY") >= Format(Tgl1, "MM/DD/YYYY")) And (Format(rsTemp("INVDATE"), "MM/DD/YYYY") <= Format(Tgl2, "MM/DD/YYYY"))) Then
11                        Call clearTemp
12                        U10K = IIf(rsTemp("FS_RC") = "R010" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
13                        U20K = IIf(rsTemp("FS_RC") = "R020" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
14                        U30K = IIf(rsTemp("FS_RC") = "R030" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
15                        U40K = IIf(rsTemp("FS_RC") = "R040" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
16                        U50K = IIf(rsTemp("FS_RC") = "R050" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
17                        U60K = IIf(rsTemp("FS_RC") = "R060" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
18                        U70K = IIf(rsTemp("FS_RC") = "R070" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
19                        U80K = IIf(rsTemp("FS_RC") = "R080" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
20                        U90K = IIf(rsTemp("FS_RC") = "R090" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
21                        U100K = IIf(rsTemp("FS_RC") = "R100" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
22                        U5K = IIf(((CDbl(Right(rsTemp("FS_RC"), 3)) / 5) Mod 2 = 1 And CDbl(Right(rsTemp("FS_RC"), 3)) < 100) And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
23                        U101K = IIf(CDbl(Right(rsTemp("FS_RC"), 3)) > 100 And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, 1, 0)
24                        UPDS = IIf(rsTemp("FS_RC") = "0" And rsTemp("PDS") = 1 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 0, 1, 0)
25                        USBI = IIf(rsTemp("FS_RC") = "R001" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 0, 1, 0)
26                        UGRP = IIf(CDbl(Right(rsTemp("FS_RC"), 3)) = 0 And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 0, 1, 0)
27                        UTWC = IIf(rsTemp("FS_RC") = "0" And rsTemp("PDS") = 0 And rsTemp("TPSVC") = "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 0, 1, 0)
28                        ULUB = IIf(rsTemp("FS_RC") = "0" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 1, 1, 0)
29                        URTJ = IIf(rsTemp("FN_RETURN") = 1, 1, 0)
30                        UTHS = IIf(rsTemp("THS") = 1, 1, 0)
31                        USBE = IIf(CDbl(Right(rsTemp("FS_RC"), 3)) > 1 And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 1, 1, 0)
                              
32                        sSQL = ""
33                        sSQL = "INSERT INTO " & vNamaTabel & " VALUES ('" & Trim(rsTemp("FRONT")) & "','" & Trim(rsTemp("NOINV")) & "','" & Trim(rsTemp("NO_WO")) & "','" & Trim(rsTemp("NO_POLISI")) & "','" & Trim(rsTemp("NO_CHASSIS")) & "','" & CDbl(URTJ) & "', " & _
                                 "'" & CDbl(UTHS) & "','" & CDbl(USBI) & "','" & CDbl(USBE) & "', '" & CDbl(UGRP) & "', " & _
                                 "'" & CDbl(UTWC) & "','" & CDbl(UPDS) & "', '" & CDbl(ULUB) & "','" & Trim(rsTemp("STATUS")) & "','" & Trim(rsTemp("CUSTOMER")) & "'," & _
                                 "'" & Trim(rsTemp("NOFAKTUR")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', "
34                        sSQL = sSQL + "'" & CDbl(U10K) & "','" & CDbl(U20K) & "','" & CDbl(U30K) & "','" & CDbl(U40K) & "', '" & CDbl(U50K) & "','" & CDbl(U60K) & "','" & CDbl(U70K) & "','" & CDbl(U80K) & "', " & _
                                 "'" & CDbl(U90K) & "','" & CDbl(U100K) & "', '" & CDbl(U5K) & "', '" & CDbl(U101K) & "')"
35                        Call objO.UpdateData(sSQL)
36                    End If
                                      
37                    If Trim(rsTemp("FS_NO_RETUR")) <> "" Then
38                        Call clearTemp
39                        U10K = IIf(rsTemp("FS_RC") = "R010" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
40                        U20K = IIf(rsTemp("FS_RC") = "R020" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
41                        U30K = IIf(rsTemp("FS_RC") = "R030" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
42                        U40K = IIf(rsTemp("FS_RC") = "R040" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
43                        U50K = IIf(rsTemp("FS_RC") = "R050" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
44                        U60K = IIf(rsTemp("FS_RC") = "R060" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
45                        U70K = IIf(rsTemp("FS_RC") = "R070" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
46                        U80K = IIf(rsTemp("FS_RC") = "R080" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
47                        U90K = IIf(rsTemp("FS_RC") = "R090" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
48                        U100K = IIf(rsTemp("FS_RC") = "R100" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
49                        U5K = IIf(((CDbl(Right(rsTemp("FS_RC"), 3)) / 5) Mod 2 = 1 And CDbl(Right(rsTemp("FS_RC"), 3)) < 100) And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
50                        U101K = IIf(CDbl(Right(rsTemp("FS_RC"), 3)) > 100 And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0, -1, 0)
51                        UPDS = IIf(rsTemp("FS_RC") = "0" And rsTemp("PDS") = 1 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 0, -1, 0)
52                        USBI = IIf(rsTemp("FS_RC") = "R001" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 0, -1, 0)
53                        UGRP = IIf(CDbl(Right(rsTemp("FS_RC"), 3)) = 0 And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 0, -1, 0)
54                        UTWC = IIf(rsTemp("FS_RC") = "0" And rsTemp("PDS") = 0 And rsTemp("TPSVC") = "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 0, -1, 0)
55                        ULUB = IIf(rsTemp("FS_RC") = "0" And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 1, -1, 0)
56                        URTJ = IIf(rsTemp("FN_RETURN") = 1, -1, 0)
57                        UTHS = IIf(rsTemp("THS") = 1, -1, 0)
58                        USBE = IIf(CDbl(Right(rsTemp("FS_RC"), 3)) > 1 And rsTemp("PDS") = 0 And rsTemp("TPSVC") <> "W" And rsTemp("FN_RETURN") = 0 And rsTemp("ISLUB") = 1, -1, 0)
                      
59                        sSQL = ""
60                        sSQL = "INSERT INTO " & vNamaTabel & " VALUES ('" & Trim(rsTemp("FRONT")) & "','" & Trim(rsTemp("NOINV")) & "','" & Trim(rsTemp("NO_WO")) & "','" & Trim(rsTemp("NO_POLISI")) & "','" & Trim(rsTemp("NO_CHASSIS")) & "','" & CDbl(URTJ) & "', " & _
                                 "'" & CDbl(UTHS) & "','" & CDbl(USBI) & "','" & CDbl(USBE) & "', '" & CDbl(UGRP) & "', " & _
                                 "'" & CDbl(UTWC) & "','" & CDbl(UPDS) & "', '" & CDbl(ULUB) & "','" & Trim(rsTemp("STATUS")) & "','" & Trim(rsTemp("CUSTOMER")) & "'," & _
                                 "'" & Trim(rsTemp("NOFAKTUR")) & "','" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY") & "', "
61                        sSQL = sSQL + "'" & CDbl(U10K) & "','" & CDbl(U20K) & "','" & CDbl(U30K) & "','" & CDbl(U40K) & "', '" & CDbl(U50K) & "','" & CDbl(U60K) & "','" & CDbl(U70K) & "','" & CDbl(U80K) & "', " & _
                                 "'" & CDbl(U90K) & "','" & CDbl(U100K) & "', '" & CDbl(U5K) & "', '" & CDbl(U101K) & "')"
62                        Call objO.UpdateData(sSQL)
63                    End If
64                    rsTemp.MoveNext
65                Loop
66            End If
67            Call Cetak(vNamaTabel, txt)
68        End If
69        If rsTemp.State <> adStateClosed Then rsTemp.Close
End Sub

Private Sub Cetak(vNamaTabel As String, txt As String)
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New rptRepoInvRinci
    Set oRs = objO.FetchData("SELECT SA, INVDATE, NOINV, NO_WO, NO_POLISI, CUSTOMER, [10K], [20K], [30K], [40K], [50K], [60K], [70K], [80K], [90K], [100K], [X5K], [U100K], PDS, SBI, GRP, TWC, LUB, RTJ, THS FROM RINCIKM WITH(NOLOCK) ORDER BY INVDATE")
    
    Report.database.SetDataSource oRs
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
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub BuatTabelRepoInv(vNamaTabel As String)
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)

    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
         & "    [SA][char](5) NOT NULL DEFAULT '', " _
         & "    [NOINV][varchar](15) NOT NULL DEFAULT '', " _
         & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NO_POLISI][varchar](12) NOT NULL DEFAULT '', " _
         & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "    [RTJ][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [THS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [GRP][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TWC][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PDS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [LUB][decimal](18,2) NOT NULL DEFAULT 0, " _
         
    sSQL = sSQL + "[TIPEINV][varchar](20) NOT NULL DEFAULT '', " _
         & "    [CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
         & "    [NOFAKTUR][varchar](30) NOT NULL DEFAULT ''," _
         & "    [INVDATE][datetime] NOT NULL DEFAULT '01/01/3000', " _
         & "    [10K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [20K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [30K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [40K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [50K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [60K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [70K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [80K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [90K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [100K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [x5K][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [U100K][decimal](18,2) NOT NULL DEFAULT 0) ON [PRIMARY]"
    Call objO.UpdateData(sSQL)
End Sub

Private Sub cmdSettingPrinter_Click()
    If opTipe(0).value Then
        Report.PrinterSetup Me.hWnd
        CRViewer1.Refresh
    End If
End Sub

Private Sub cmdTutup_Click()
    sSQL = "DELETE [" & vNamaTabel1 & "]"
    Call objO.UpdateData(sSQL)
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

Private Sub Form_Load()
    Call bersih
End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 345
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdCetak.Top = CRViewer1.Top + 345
    cmdCetak.Left = cboPrinter.Left + cboPrinter.Width + 30
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
        cmdCetak.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdCetak_Click()
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
    f.show (vbModal)
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
            .show vbModal
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

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call bersih
End Sub

Private Sub bersih()
    Tgl1 = Format(isnow, "DD/MM/YYYY")
    Tgl2 = Format(isnow, "DD/MM/YYYY")
    opVehicle(0).value = True
    Text(7).Text = ""
    Text(8).Text = ""
    Text(14).Text = ""
    txtTemp = ""
End Sub

Private Sub tblSimpan_Click()
txtTemp = ""
    If opTipe(0).value Then

        If Text(7) <> "" Or Text(14) <> "" Then
            If opVehicle(0) Then
                txtTemp = txtTemp & "AND A.MDL = '" & Trim(Text(7)) & "' "
            Else
                txtTemp = txtTemp & "AND A.FMDL LIKE '%" & Trim(Text(14)) & "%' "
            End If
        End If
        vNamaTabel1 = "RinciKM"
        Call IsiRinciKM(vNamaTabel1, "")
        Call DisableExport
        Call setLokasiTombol
    End If
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset

    createTemp = False
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepoInv(vNamaTabel)
        createTemp = True
    End If
End Function

Private Function statusDealer(codeStatus As Integer)
    Select Case codeStatus
        Case 1
            statusDealer = "AFFNONPPN"
        Case 2
            statusDealer = "AFFPPN"
        Case 3
            statusDealer = "EXT"
    End Select
End Function

Private Function cekDuo(hrg As Double, nINV As String, Optional stat As Integer)
    If Right(Trim(nINV), 2) = "-B" Then
        cekDuo = 0
        If stat = 1 Then
            cekDuo = CDbl(hrg)
        End If
    Else
        cekDuo = CDbl(hrg)
        If stat = 2 Then
            cekDuo = 0
        End If
    End If
End Function

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    asxPowerBanner1.Height = Me.Height - 500
End Sub

