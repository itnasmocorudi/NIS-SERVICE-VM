VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOPerbaikanNoRangkaTWC 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "PERBAIKAN NO RANGKA"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7800
      TabIndex        =   18
      Top             =   4800
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   8438015
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdGanti 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5880
      TabIndex        =   22
      Top             =   4800
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Ganti No Rangka"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   11775
      Begin MSAdodcLib.Adodc AdcKode 
         Height          =   330
         Left            =   8760
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DCKodeGroup 
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   720
         Width           =   3195
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   5644
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "- Pilih Nama Dealer -"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   11400
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   2
            Left            =   1440
            TabIndex        =   4
            Top             =   480
            Width           =   3700
            _ExtentX        =   6535
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   3
            Left            =   1440
            TabIndex        =   5
            Top             =   840
            Width           =   3700
            _ExtentX        =   6535
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   4
            Left            =   1440
            TabIndex        =   6
            Top             =   1200
            Width           =   3700
            _ExtentX        =   6535
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   465
            Index           =   7
            Left            =   8400
            TabIndex        =   7
            Top             =   480
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   820
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99999999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   6
            Left            =   1440
            TabIndex        =   8
            Top             =   1920
            Width           =   3700
            _ExtentX        =   6535
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   5
            Left            =   1440
            TabIndex        =   24
            Top             =   1560
            Width           =   3700
            _ExtentX        =   6535
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   30
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No WO"
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
            Left            =   750
            TabIndex        =   25
            Top             =   1612
            Width           =   630
         End
         Begin VB.Line Line8 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   3840
            Y1              =   195
            Y2              =   195
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   3840
            Y1              =   75
            Y2              =   75
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Detail TWC "
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
            Left            =   3840
            TabIndex        =   15
            Top             =   0
            Width           =   4665
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   8520
            X2              =   12600
            Y1              =   75
            Y2              =   75
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   8520
            X2              =   12600
            Y1              =   195
            Y2              =   195
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   12855
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Bukti"
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
            TabIndex        =   14
            Top             =   532
            Width           =   780
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Claim"
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
            Left            =   585
            TabIndex        =   13
            Top             =   892
            Width           =   795
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No TWC"
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
            Left            =   660
            TabIndex        =   12
            Top             =   1252
            Width           =   720
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   7695
            TabIndex        =   11
            Top             =   592
            Width           =   480
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Rangka"
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
            Left            =   360
            TabIndex        =   10
            Top             =   1972
            Width           =   1020
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   3960
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   3960
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Customer && Kendaraan"
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
            Height          =   300
            Left            =   4320
            TabIndex        =   9
            Top             =   45
            Width           =   3675
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   8640
            X2              =   12720
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   8640
            X2              =   12720
            Y1              =   240
            Y2              =   240
         End
      End
      Begin MSMask.MaskEdBox text 
         Height          =   345
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Top             =   1200
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   345
         Index           =   1
         Left            =   2520
         TabIndex        =   20
         Top             =   4440
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERBAIKAN NOMOR RANGKA TWC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   2775
         TabIndex        =   23
         Top             =   120
         Width           =   4860
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BorderColor     =   &H80000002&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   15
         Width           =   11760
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Rangka Benar"
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
         Left            =   405
         TabIndex        =   21
         Top             =   4440
         Width           =   2025
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Rangka Salah"
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
         Left            =   480
         TabIndex        =   17
         Top             =   1200
         Width           =   1980
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Dealer"
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
         Left            =   1200
         TabIndex        =   16
         Top             =   780
         Width           =   1215
      End
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmHOPerbaikanNoRangkaTWC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dealer As String
Private dBase As String
Private koneksi As String
Private usernameDealer As String
Private passwordDealer As String
Private usernameHO As String
Private passwordHO As String
Private usernameGlobal As String
Private passwordGlobal As String
Private isDoubleCustDat As Boolean
Private isDoubleCustDatHO As Boolean

Private isDealerNewEnkripsi As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGanti_Click()
          Dim rsTWC As ADODB.Recordset
1         If Len(Trim(text(0))) = 0 Then
2             MsgBox "No Rangka yang salah harus diisi!", vbCritical, "Informasi"
3             Exit Sub
4         ElseIf Len(Trim(text(1))) = 0 Then
5             MsgBox "No Rangka yang salah harus diisi!", vbCritical, "Informasi"
6             Exit Sub
7         End If
8         sSQL = "select top 1 * FROM [10.10.10.180].[replbengkel99].dbo.ptclist where Norangka = '" & NulSID(text(0)) & "'"
9         Set rsTWC = objO.FetchDataHO(sSQL)
10        If Not rsTWC.EOF Then
11            MsgBox "Nomor Rangka sudah divalidasi TWC", vbInformation
12            Exit Sub
13        End If
          'Call hapusForm
14        If PencarianNoRangka = False Then Exit Sub
          
15        If MsgBox("Apakah yakin mengganti No Rangka " & Trim(text(0)) & " dengan " & Trim(text(1)) & " ?", vbYesNo, "Konfirmasi Penggantian No Rangka") = vbYes Then
              'lakukan pergantian data
16            On Error GoTo vbError
              'koneksi = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer
17            erTranAll = False
              'perbaikan kecabang
             ' Call objHT.OpenTransactionPerbaikanDataHO(koneksi)
18            Call objHT.OpenTransactionHO
              
19            sSQL = "update [10.10.10.180].[" & dbReplikasiHO & "].dbo.Ptcfile set norangka = '" & NulSID(text(1)) & "' where Norangka = '" & NulSID(text(0)) & "'"
              
20            Call objHT.UpdateData_Ho(sSQL)
               
21            Call objHT.UpdateData_Ho("INSERT INTO PROFILBENGKEL..LOGUPDATE VALUES ('UpdNokaTWC','','" & Trim(text(0)) & "','" & gSA & "',GETDATE(),'" & Trim(text(1)) & "','" & Trim(text(0)) & "')")
              
22            If erTranAll = False Then
23                Call objHT.CloseTransactionHO
                  'Call saveLogMenu(gSA, ComptName, gVersi, "frmHOPerbaikanNoRangkaTWC", "gantiNoRangka", 1)
24                MsgBox "No Rangka berhasil diubah!", vbInformation, "Informasi Berhasil Ganti No Rangka"
25                Call hapusForm
26                text(0) = ""
27                Call FetchCombo
28                Exit Sub
29            Else
30                Call objHT.CloseTransactionHO_err
                 ' Call saveLogMenu(gSA, ComptName, gVersi, "frmHOPerbaikanNoRangkaTWC", "gantiNoRangka", 2)
31                MsgBox "Terjadi kesalahan saat mengganti No Rangka!", vbInformation, "Gagal Simpan"
32                Exit Sub
33            End If
vbError:
34            Call objHT.CloseTransactionHO_err
              'Call saveLogMenu(gSA, ComptName, gVersi, "CloseTransactionPerbaikanDataHO_err", "gantiNoRangka", 2)
35            MsgBox "Terjadi kesalahan saat mengganti No Rangka!", vbInformation, "Gagal Simpan"
36            Exit Sub
37        End If
End Sub

Private Sub DCKodeGroup_Change()
    'dealer = ip dealer
    'dbase  = database dealer
    isDealerNewEnkripsi = False
    FS_KD_DEALER = DCKodeGroup.Columns(0).value
    dealer = DCKodeGroup.Columns(2).value
    dBase = DCKodeGroup.Columns(3).value
    
    If DCKodeGroup.Columns(6).value = "1" Then isDealerNewEnkripsi = True
            
    If isDealerNewEnkripsi Then
        usernameDealer = Dekripsi(DCKodeGroup.Columns(4).value, KataKunci(3))
        passwordDealer = Dekripsi(DCKodeGroup.Columns(5).value, KataKunci(5))
    Else
        usernameDealer = DekripsiGamaLama(DCKodeGroup.Columns(4).value, KataKunciGamaLama(3))
        passwordDealer = DekripsiGamaLama(DCKodeGroup.Columns(5).value, KataKunciGamaLama(5))
    End If
End Sub

Private Sub DCKodeGroup_Click()
    Call DCKodeGroup_Change
End Sub

Private Sub Form_Load()
    Call FetchCombo
End Sub

Private Sub FetchCombo()
    AdcKode.ConnectionString = SQLConHO
    AdcKode.RecordSource = "select FS_KD_DEALER, FS_NM_DEALER,VPNCABANG,DBASECBG,USERCBG,PASSCBG,NEWENKRIP from t_dealer WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1"
    DCKodeGroup.DataSourceList = AdcKode
    DCKodeGroup.DataFieldList = "FS_NM_DEALER"
    DCKodeGroup.Columns(0).Width = 1200
    DCKodeGroup.Columns(1).Width = 4000
    DCKodeGroup.Columns(0).Caption = "Kode Dealer"
    DCKodeGroup.Columns(1).Caption = "Nama Dealer"
    DCKodeGroup.Columns(2).Visible = False
    DCKodeGroup.Columns(3).Visible = False
    DCKodeGroup.Columns(4).Visible = False
    DCKodeGroup.Columns(5).Visible = False
    DCKodeGroup.Columns(6).Visible = False
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call text_Validate(Index, True)
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If dealer <> "" Then
                If text(Index) <> "" Then
                    Call hapusForm
                    Call PencarianNoRangkaSalah
                Else
                    MsgBox "No Rangka harus diisi!", vbCritical, "Informasi"
                End If
            Else
                MsgBox "Dealer harus dipilih!", vbCritical, "Informasi"
            End If
    End Select
End Sub

Private Sub PencarianNoRangkaSalah()
          Dim noRangkaSalah As String
          Dim noRangkaBenar As String
          
1         On Error GoTo errHandler
          
2         noRangkaSalah = Trim(text(0))
          'noRangkaBenar = Trim(text(1))
          
          Dim rsCustdatSalah As New ADODB.Recordset
          Dim rsCustdatBenar As New ADODB.Recordset
          Dim rsCustdatSalahHO As New ADODB.Recordset
          Dim rsCustdatBenarHO As New ADODB.Recordset
           
              
3         isDoubleCustDat = False
4         isDoubleCustDatHO = False
          Dim rs As New ADODB.Recordset
          Dim rsHO As New ADODB.Recordset
           
          'mengambil data dealer
      '    rsCustdatSalah.Open "SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaSalah) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
      '    'rsCustdatBenar.Open "SELECT * FROM CUSTDAT WHERE NO_CHASSIS = '" & Trim(noRangkaBenar) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dbase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
      '
      '    'mengambil data HO
      '    Set rsCustdatSalahHO = objO.FetchDataHO("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangkaSalah) & "'")
          
          '' 20190528 MENGAMBIL DATA PTCFILE 10.180
          '' ---------------------------------------------------
          
5         sSQL = "select a.nobukti,a.claimno,a.twcno,a.total,a.nowo,a.norangka"
6         sSQL = sSQL & " from [10.10.10.180].replbengkel99.dbo.ptcfile a"
7         sSQL = sSQL & " where a.norangka = '" & NulSID(noRangkaSalah) & "' and a.dealer = '" & Replace(NulSID(DCKodeGroup.Columns(0).value), " ", "") & "'"
8         Set rsCustdatSalah = objO.FetchDataHO(sSQL)
          
9        If Not rsCustdatSalah.EOF Then
          
10            text(2) = Left(NulSID(rsCustdatSalah("nobukti")), 60)
11            text(3) = Left(NulSID(rsCustdatSalah("claimno")), 60)
12            text(4) = Left(NulSID(rsCustdatSalah("twcno")), 60)
13            text(5) = Left(NulSID(rsCustdatSalah("nowo")), 60)
14            text(6) = Left(NulSID(rsCustdatSalah("norangka")), 60)
15            text(7) = Left(NulSID(rsCustdatSalah("total")), 60)
              
16            Label8.Visible = True
17            text(1).Visible = True
18            cmdGanti.Visible = True
19            text(1).SetFocus
              
20        Else
21            MsgBox "Tidak ada No Rangka yang dicari", vbOKOnly, "Informasi"
22        End If
23        Exit Sub
errHandler:
24        MsgBox "Terjadi kesalahan dalam mengambil data!", vbCritical, "Informasi" 'err.Source & ":" & vbCrLf & err.Description
End Sub

Private Function PencarianNoRangka() As Boolean
          Dim noRangkaSalah As String
          Dim noRangkaBenar As String
          
1         On Error GoTo errHandler
2         PencarianNoRangka = True
3         noRangkaSalah = Trim(text(0))
4         noRangkaBenar = Trim(text(1))
          
          Dim rsCustdatSalah As New ADODB.Recordset
          Dim rsCustdatBenar As New ADODB.Recordset
          Dim rsCustdatSalahHO As New ADODB.Recordset
          Dim rsCustdatBenarHO As New ADODB.Recordset
          
5         Set rsCustdatBenar = objO.FetchDataHO("SELECT TOP 10 * FROM T_INVOICEnew WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(noRangkaBenar) & "'")
6         If Not rsCustdatBenar.EOF Then
7             PencarianNoRangka = True
8         Else
9             MsgBox "No Rangka " & noRangkaBenar & " tidak terdaftar.", vbCritical
10            PencarianNoRangka = False
11        End If
12        Exit Function
errHandler:
13        MsgBox "Terjadi kesalahan dalam mengambil data!", vbCritical, "Informasi" 'err.Source & ":" & vbCrLf & err.Description
14        PencarianNoRangka = False
End Function

Private Sub hapusForm()
    For i = 2 To 7
        text(i) = ""
    Next
    text(1) = ""
    Label8.Visible = False
    text(1).Visible = False
    cmdGanti.Visible = False
End Sub
