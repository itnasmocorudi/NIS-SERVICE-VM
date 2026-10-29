VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmRepoHOWar 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   12555
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   720
      TabIndex        =   13
      Top             =   5280
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   480
      TabIndex        =   14
      Top             =   480
      Width           =   12015
      Begin VB.Frame frmeSSC 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   7000
         TabIndex        =   33
         Top             =   2280
         Visible         =   0   'False
         Width           =   4575
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   885
            TabIndex        =   9
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   270
            Left            =   2160
            TabIndex        =   34
            Top             =   0
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
            Left            =   885
            TabIndex        =   35
            Top             =   360
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Index           =   10
            Left            =   885
            TabIndex        =   37
            Top             =   720
            Width           =   3555
            _ExtentX        =   6271
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
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe SSC"
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
            Left            =   45
            TabIndex        =   36
            Top             =   45
            Width           =   630
         End
      End
      Begin VB.Frame frmeWarranty 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Top             =   2280
         Width           =   2775
         Begin VB.ComboBox cboWar 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Warranty"
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
            Left            =   0
            TabIndex        =   32
            Top             =   60
            Width           =   690
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   7920
         TabIndex        =   30
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton opLap 
            BackColor       =   &H00FFFFFF&
            Caption         =   "SSC"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   3
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton opLap 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Warranty"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   7890
         TabIndex        =   6
         Top             =   840
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
      Begin VB.OptionButton opDealer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dealer Penjual"
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   28
         Top             =   840
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   4
         Top             =   840
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
      Begin VB.OptionButton opDealer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dealer Service"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1530
         TabIndex        =   0
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   354287617
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   4440
         TabIndex        =   15
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
         TabIndex        =   5
         Top             =   1560
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   4440
         TabIndex        =   16
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
         Index           =   3
         Left            =   1530
         TabIndex        =   17
         Top             =   1920
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   10800
         TabIndex        =   18
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
         Index           =   1
         Left            =   1530
         TabIndex        =   19
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
         Index           =   5
         Left            =   7890
         TabIndex        =   20
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
         Index           =   6
         Left            =   7890
         TabIndex        =   7
         Top             =   1560
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   10800
         TabIndex        =   21
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
         Index           =   7
         Left            =   7890
         TabIndex        =   22
         Top             =   1920
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
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   354287617
         CurrentDate     =   39852
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Laporan"
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
         Left            =   6870
         TabIndex        =   29
         Top             =   240
         Width           =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         X1              =   0
         X2              =   12000
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
         Left            =   3000
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
         Left            =   750
         TabIndex        =   25
         Top             =   240
         Width           =   570
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
         Left            =   6765
         TabIndex        =   24
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Service"
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
         Left            =   285
         TabIndex        =   23
         Top             =   1605
         Width           =   1035
      End
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
      Left            =   10440
      TabIndex        =   11
      Top             =   3960
      Width           =   1890
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
      Left            =   8520
      TabIndex        =   10
      Top             =   3960
      Width           =   1890
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
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
      FormatString    =   "Laporan Warranty"
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
Attribute VB_Name = "frmRepoHOWar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vNamaTabel As String

Private Sub hapusForm()
    For i = 0 To 10
        text(i) = ""
    Next
    Tgl1 = Format(isnow, "DD/MM/YYYY")
    Tgl2 = Format(isnow, "DD/MM/YYYY")
    
    cboWar.Clear
    cboWar.AddItem ("Other")
    cboWar.AddItem ("All")
    frmeWarranty.Visible = True
    
    opLap(0).value = True
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Load()
    Call hapusForm
    opDealer(0) = True
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

Private Sub opDealer_Click(Index As Integer)
    
    Select Case Index
        Case 0
            text(0).Enabled = True
            text(1).Enabled = True
            text(2).Enabled = True
            text(3).Enabled = True
            
            text(4).Enabled = False
            text(5).Enabled = False
            text(6).Enabled = False
            text(7).Enabled = False
        Case 1
            text(0).Enabled = False
            text(1).Enabled = False
            text(2).Enabled = False
            text(3).Enabled = False
            
            text(4).Enabled = True
            text(5).Enabled = True
            text(6).Enabled = True
            text(7).Enabled = True
    End Select
End Sub

Private Sub opLap_Click(Index As Integer)
    If opLap(Index) Then
        Select Case Index
            Case 0
                frmeWarranty.Visible = True
                frmeSSC.Visible = False
            Case 1
                frmeWarranty.Visible = False
                frmeSSC.Visible = True
        End Select
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim boleh As Boolean
Dim pesan As String
Dim Txt As String
    
    boleh = True
    If opDealer(0) Then
        If text(0) = "" Then
            boleh = False
            pesan = pesan + "Dealer Service. "
        End If
    ElseIf opDealer(1) Then
        If text(4) = "" Then
            boleh = False
            pesan = pesan + "Dealer Penjual. "
        End If
    End If
    
    If opLap(0) Then
        If cboWar = "" Then
            boleh = False
            pesan = pesan + "Jenis Warranty. "
        End If
    ElseIf opLap(1) Then
        If Trim(text(8)) = "" Then
            boleh = False
            pesan = pesan + "Jenis SSC. "
        End If
    End If
    
    If boleh Then
        If opDealer(0) Then
            If text(2) <> "" Then
                Txt = "A.FS_KD_DEALER BETWEEN '" & text(0) & "' AND '" & text(2) & "'"
            Else
                Txt = "A.FS_KD_DEALER = '" & text(0) & "'"
            End If
        ElseIf opDealer(1) Then
            If text(6) <> "" Then
                Txt = "B.FS_KD_DEALER BETWEEN '" & text(4) & "' AND '" & text(6) & "'"
            Else
                Txt = "B.FS_KD_DEALER = '" & text(4) & "'"
            End If
        End If
        
        If frmeWarranty.Visible Then
            Dim txt1 As String
            Select Case cboWar
                Case "SSC"
                    txt1 = "KODETIPE <> ''"
                Case "Other"
                    txt1 = "KODETIPE = ''"
                Case "All"
                    txt1 = ""
            End Select
            If Txt <> "" Then
                If txt1 <> "" Then
                    Txt = Txt & " AND " & txt1
                End If
            End If
            vNamaTabel = "HOWAR" & DriveSerialNumber("c:")
            Call IsiTempRepo(vNamaTabel, Txt)
        ElseIf frmeSSC.Visible Then
            If Txt <> "" Then
                Txt = Txt & " AND KODETIPE = '" & Trim(text(8)) & "'"
            End If
            vNamaTabel = "HOSSC" & DriveSerialNumber("c:")
            Call IsiTempRepo1(vNamaTabel, Txt)
        End If
    Else
        MsgBox pesan, vbCritical, "KELENGKAPAN"
        Exit Sub
    End If
End Sub

Private Sub IsiTempRepo1(nTabel As String, ntxt As String)
Dim rsHlp As ADODB.Recordset
Dim rsHlp2 As ADODB.Recordset
Dim hlp As String
Dim hlp2 As String

    Set rsHlp = objO.FetchDataHO("SELECT DISTINCT A.NO_CHASSIS, A.NO_ENGINE, MDL, NO_WO, NOINV, DDATE, B.FS_NM_DEALER AS [DEALER_PENJUAL], A.NO_POLISI, [DATE], KODETIPE, A.FS_NM_DEALER AS [DEALER_REPAIR], " _
                              & "REP1,REP2,REP3,REP4,REP5,REP6,REP7, RTE1 , RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, C.PRICE AS HRGJASA, A.DISC, memoPart, memoSublet " _
                              & "FROM HISTORY A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN CARCLASS C WITH(NOLOCK) ON B.MODEL = C.CLASSCODE AND C.FS_LOG_DEL = ''" _
                              & "WHERE A.FS_LOG_DEL = '' AND TIPEWO = 'W' AND " & ntxt & " ORDER BY A.NO_CHASSIS")
    If Not rsHlp.EOF Then
        Do While Not rsHlp.EOF
            If rsHlp("MEMOSUBLET") <> "" Then
                Call potongSUBLET(rsHlp("NO_WO"), rsHlp("NOINV"), rsHlp("NO_CHASSIS"), rsHlp("MEMOSUBLET"))
            End If
            If rsHlp("MEMOPART") <> "" Then
                Call potongPart(rsHlp("NO_WO"), rsHlp("NOINV"), rsHlp("NO_CHASSIS"), rsHlp("MEMOPART"))
            End If
            
            For i = 1 To 7
                hlp = "REP" & i
                If rsHlp(hlp) <> "" Then
                    hlp2 = "RTE" & i    '"ACTRATE" & i
                    If createTemp(vNamaTabel, 4) Then
                        
                        ' CEK KELENGKAPAN PENGGANTIAN PART
                        Dim nLengkap As String
                        Set rsHlp2 = objO.FetchData("SELECT DISTINCT (CASE (Q.QTY - T.QTY) WHEN  0 THEN 'OK' ELSE 'NOK' END) AS STATUS " _
                                                 & "FROM [" & vNamaTabel & "-PART] Q WITH(NOLOCK) LEFT JOIN HO_WARJOBPART T WITH(NOLOCK) ON SUBSTRING(Q.KODEITEM, 3, LEN(Q.KODEITEM)-2) = T.ITEMPART " _
                                                 & "GROUP BY (CASE (Q.QTY - T.QTY) WHEN  0 THEN 'OK' ELSE 'NOK' END)")
                        If Not rsHlp2.EOF Then
                            If rsHlp2.RecordCount > 1 Or rsHlp2.RecordCount = 0 Then
                                ' TIDAK LENGKAP
                                nLengkap = "PERIKSA"
                            ElseIf rsHlp2.RecordCount = 1 Then
                                If rsHlp2("STATUS") = "OK" Then
                                    ' LENGKAP
                                    nLengkap = "GANTI"
                                End If
                            End If
                        End If
                        
                        Set rsHlp2 = objO.FetchDataHO("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "'")
                        If Not rsHlp2.EOF Then
                            Call objO.UpdateDataHO("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET NO_ENGINE = '" & rsHlp("NO_ENGINE") & "', MDL = '" & rsHlp("MDL") & "' ,DDATE = '" & Format(rsHlp("DDATE"), "MM/DD/YYYY") & "',DEALERPENJUAL = '" & rsHlp("DEALER_PENJUAL") & "',NO_POLISI = '" & rsHlp("NO_POLISI") & "',[DATE] = '" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "', " _
                                           & "KODETIPE = '" & rsHlp("KODETIPE") & "',DEALERREPAIR = '" & rsHlp("DEALER_REPAIR") & "',REPS = '" & rsHlp(hlp) & "',RTE = " & rsHlp(hlp2) & ", DISC = '" & CDbl(rsHlp("DISC")) & "', HRGJASA = " & CDbl(rsHlp("HRGJASA")) & ", PART = '" & nLengkap & "' WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "'")
                        Else
                            'NO_CHASSIS,NO_ENGINE,MDL,NO_WO,NOINV,DDATE,DEALERPENJUAL,NO_POLISI,[DATE],KODETIPE,DEALERREPAIR,REPS,RTE,PART,SUBLET
                            Call objO.InsertDataHO("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsHlp("NO_CHASSIS") & "', '" & rsHlp("NO_ENGINE") & "', '" & rsHlp("MDL") & "', '" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "', '" & Format(rsHlp("DDATE"), "MM/DD/YYYY") & "', '" & Trim(rsHlp("DEALER_PENJUAL")) & "', " _
                                           & "'" & rsHlp("NO_POLISI") & "', '" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "', '" & Trim(rsHlp("KODETIPE")) & "', '" & Trim(rsHlp("DEALER_REPAIR")) & "', '" & rsHlp(hlp) & "', '" & rsHlp(hlp2) & "', '" & CDbl(rsHlp("DISC")) & "', " & CDbl(rsHlp("HRGJASA")) & ", '" & Trim(nLengkap) & "')")
                        End If
                    End If
                End If
            Next
            rsHlp.MoveNext
        Loop
        
        sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "]"
        Call Cetak2(sSQL)
    Else
        MsgBox "Tidak ada Data !"
        Exit Sub
    End If
End Sub

Private Sub Cetak2(Txt As String)
Dim oRs As Object   'New ADODB.Recordset  'Object
Dim Report As New RptHOSSC
Screen.MousePointer = vbHourglass
    
    Set oRs = objO.FetchDataHO(Txt)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtTanggal.SetText "Tanggal " & Format(Tgl1, "DD MMMM YYYY") & " - " & Format(Tgl2, "DD MMMM YYYY")
        
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
        cmdTutup.Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub IsiTempRepo(nTabel As String, ntxt As String)
Dim rsHlp As ADODB.Recordset
Dim rsHlp2 As ADODB.Recordset
Dim hlp As String
Dim hlp2 As String
Dim SSQL2 As String
Dim SSQL3 As String
    
    Set rsHlp = objO.FetchDataHO("SELECT DISTINCT A.NO_CHASSIS, NO_WO, NOINV, DDATE, B.FS_NM_DEALER AS [DEALER_PENJUAL], A.NO_POLISI, [DATE], KODETIPE, A.FS_NM_DEALER AS [DEALER_REPAIR], " _
                              & "REP1,REP2,REP3,REP4,REP5,REP6,REP7, RTE1 , RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, C.PRICE AS HRGJASA, A.DISC, memoPart, memoSublet " _
                              & "FROM HISTORY A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN CARCLASS C WITH(NOLOCK) ON B.MODEL = C.CLASSCODE AND C.FS_LOG_DEL = ''" _
                              & "WHERE A.FS_LOG_DEL = '' AND TIPEWO = 'W' AND " & ntxt & " ORDER BY A.NO_CHASSIS")
    If Not rsHlp.EOF Then
        Do While Not rsHlp.EOF
            If createTemp(vNamaTabel & "-SUBLET", 3) Then
                If rsHlp("MEMOSUBLET") <> "" Then
                    Call potongSUBLET(rsHlp("NO_WO"), rsHlp("NOINV"), rsHlp("NO_CHASSIS"), rsHlp("MEMOSUBLET"))
                End If
            End If
            
            If createTemp(vNamaTabel & "-PART", 2) Then
                If rsHlp("MEMOPART") <> "" Then
                    Call potongPart(rsHlp("NO_WO"), rsHlp("NOINV"), rsHlp("NO_CHASSIS"), rsHlp("MEMOPART"))
                End If
            End If
            For i = 1 To 7
                hlp = "REP" & i
                If rsHlp(hlp) <> "" Then
                    hlp2 = "RTE" & i    '"ACTRATE" & i
                    If createTemp(vNamaTabel, 1) Then
                        Set rsHlp2 = objO.FetchDataHO("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "'")
                        If Not rsHlp2.EOF Then
                            Call objO.UpdateDataHO("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET DDATE = '" & Format(rsHlp("DDATE"), "MM/DD/YYYY") & "',DEALERPENJUAL = '" & rsHlp("DEALER_PENJUAL") & "',NO_POLISI = '" & rsHlp("NO_POLISI") & "',[DATE] = '" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "', " _
                                           & "KODETIPE = '" & rsHlp("KODETIPE") & "',DEALERREPAIR = '" & rsHlp("DEALER_REPAIR") & "',REPS = '" & rsHlp(hlp) & "',RTE = " & rsHlp(hlp2) & ", DISC = '" & CDbl(rsHlp("DISC")) & "', HRGJASA = " & CDbl(rsHlp("HRGJASA")) & " WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "'")
                        Else
                            Call objO.InsertDataHO("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsHlp("NO_CHASSIS") & "', '" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "', '" & Format(rsHlp("DDATE"), "MM/DD/YYYY") & "', '" & Trim(rsHlp("DEALER_PENJUAL")) & "', " _
                                           & "'" & rsHlp("NO_POLISI") & "', '" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "', '" & Trim(rsHlp("KODETIPE")) & "', '" & Trim(rsHlp("DEALER_REPAIR")) & "', '" & rsHlp(hlp) & "', '" & rsHlp(hlp2) & "', '" & CDbl(rsHlp("DISC")) & "', " & CDbl(rsHlp("HRGJASA")) & ")")
                        End If
                    End If
                End If
            Next
            rsHlp.MoveNext
        Loop
        
        sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "]"
        SSQL2 = "SELECT * FROM [" & vNamaTabel & "-PART]"
        SSQL3 = "SELECT * FROM [" & vNamaTabel & "-SUBLET]"
        Call Cetak(sSQL, SSQL2, SSQL3)
    Else
        MsgBox "Tidak ada Data !"
        Exit Sub
    End If
End Sub

Private Sub Cetak(Txt As String, txt2 As String, txt3 As String)
Dim oRs As Object   'New ADODB.Recordset  'Object
Dim jHlp As Object  'New ADODB.Recordset
Dim pHlp As Object  'New ADODB.Recordset
Dim Report As New RptHOWar
Dim crxSubReportJob As CRAXDRT.Report
Dim crxSubReportPart As CRAXDRT.Report
Screen.MousePointer = vbHourglass
    
    Set oRs = objO.FetchDataHO(Txt)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtTanggal.SetText Format(Tgl1, "DD MMM YYYY") & " - " & Format(Tgl2, "DD MMM YYYY")
    Report.txtDealer1.SetText IIf(Trim(text(1)) <> "", text(1), text(5))
    Report.txtDealer2.SetText IIf(Trim(text(1)) <> "", text(3), text(7))
    
    Set jHlp = objO.FetchDataHO(txt2)
    Set crxSubReportJob = Report.Subreport3.OpenSubreport
    crxSubReportJob.database.SetDataSource jHlp
    
    Set pHlp = objO.FetchDataHO(txt3)
    Set crxSubReportPart = Report.Subreport4.OpenSubreport
    crxSubReportPart.database.SetDataSource pHlp
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
        cmdTutup.Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub potongSUBLET(nWO As String, nINV As String, nRgk As String, mmSUBLET As String)
Dim tHlp As String
Dim tHlp2 As String
Dim tHlp3 As String
Dim tmSUBLET As String
Dim lNow As Integer
Dim lNow2 As Integer
Dim i As Integer
Dim nPO, nw, kdItem, nmItem, noInv, nFak, nChasis, kdSat, nNota, nKet, nFlag As String
Dim qty, Harga As Double
Dim rsHlp3 As ADODB.Recordset
    
    lNow = Len(mmSUBLET)
    tmSUBLET = mmSUBLET
    If createTemp(vNamaTabel & "-SUBLET", 3) Then
        Do While lNow > 0
            ' PISAH TANDA |
            tHlp = Left(tmSUBLET, InStr(tmSUBLET, "|"))
            lNow = lNow - Len(tHlp)
            
            lNow2 = Len(tHlp)
            i = 13
            tHlp3 = tHlp
            Do While lNow2 > 0
                If i > 1 Then
                    tHlp2 = Left(Trim(tHlp3), InStr(Trim(tHlp3), "^"))
                Else
                    tHlp2 = tHlp3
                End If
                lNow2 = lNow2 - Len(tHlp2)
                Select Case i
                    Case 13
                        nPO = Replace(Trim(tHlp2), "^", "")
                    Case 12
                        nw = Replace(Trim(tHlp2), "^", "")
                    Case 11
                        noInv = Trim(Replace(tHlp2, "^", ""))
                    Case 10
                        nChasis = Replace(Trim(tHlp2), "^", "")
                    Case 9
                        nFak = Replace(Trim(tHlp2), "^", "")
                    Case 8
                        kdItem = Replace(Trim(tHlp2), "^", "")
                    Case 7
                        nmItem = Replace(Trim(tHlp2), "^", "")
                    Case 6
                        kdSat = Replace(Trim(tHlp2), "^", "")
                    Case 5
                        qty = CDbl(Replace(tHlp2, "^", ""))
                    Case 4
                        Harga = CDbl(Replace(tHlp2, "^", ""))
                    Case 3
                        nNota = Replace(Trim(tHlp2), "^", "")
                    Case 2
                        nKet = Replace(Trim(tHlp2), "^", "")
                    Case 1
                        nFlag = Replace(Trim(tHlp2), "|", "")
                End Select
                
                i = i - 1
                tHlp3 = Mid(Trim(tHlp3), Len(tHlp2) + 1, Len(tHlp3) - Len(tHlp2))
            Loop
            Set rsHlp3 = objO.FetchDataHO("SELECT * FROM [" & vNamaTabel & "-SUBLET] WHERE NO_CHASSIS = '" & nRgk & "' AND NO_WO = '" & nWO & "' AND NOINV = '" & nINV & "' AND NOPO = '" & nPO & "' AND NAMABARANG = '" & nmItem & "'")
            If Not rsHlp3.EOF Then
                Call objO.UpdateDataHO("UPDATE [" & vNamaTabel & "-SUBLET] SET KODEITEM = '" & kdItem & "', NAMABARANG = '" & nmItem & "', KODESATUAN = '" & kdSat & "', QTY = '" & qty & "', HARGA = " & Harga & ", KETERANGAN = '" & nKet & "', FLAG = '" & nFlag & "' WHERE NO_CHASSIS = '" & nRgk & "' AND NO_WO = '" & nWO & "' AND NOINV = '" & nINV & "' AND NOPO = '" & nPO & "' AND NAMABARANG = '" & nmItem & "'")
            Else
                Call objO.InsertDataHO("INSERT INTO [" & vNamaTabel & "-SUBLET] VALUES ('" & nPO & "', '" & nw & "', '" & noInv & "', '" & nChasis & "', '" & nFak & "', '" & kdItem & "','" & nmItem & "', '" & kdSat & "','" & qty & "', " & Harga & ", '" & nNota & "', '" & nKet & "', '" & nFlag & "')")
            End If
            tmSUBLET = Mid(Trim(tmSUBLET), Len(tHlp) + 1, Len(tmSUBLET) - Len(tHlp))
        Loop
    End If
End Sub

Private Sub potongPart(nWO As String, nINV As String, nRgk As String, mmPART As String)
Dim tHlp As String
Dim tHlp2 As String
Dim tHlp3 As String
Dim tmPART As String
Dim lNow As Integer
Dim lNow2 As Integer
Dim i As Integer
Dim nw, kdItem, nmItem, noInv As String
Dim qty, Harga As Double
Dim oli, disc As Integer
Dim rsHlp3 As ADODB.Recordset
    
    lNow = Len(mmPART)
    tmPART = mmPART
    If createTemp(vNamaTabel & "-PART", 2) Then
        Do While lNow > 0
            ' PISAH TANDA |
            tHlp = Left(tmPART, InStr(tmPART, "|"))
            lNow = lNow - Len(tHlp)
            
            lNow2 = Len(tHlp)
            i = 8
            tHlp3 = tHlp
    '        tHlp2 = Left(tHlp3, InStr(tHlp3, " "))
            Do While lNow2 > 0
                If i > 1 Then
                    tHlp2 = Left(Trim(tHlp3), InStr(Trim(tHlp3), "^"))
                Else
                    tHlp2 = tHlp3
                End If
                lNow2 = lNow2 - Len(tHlp2)
                Select Case i
                    Case 8
                        nw = Replace(Trim(tHlp2), "^", "")
                    Case 7
                        kdItem = Replace(Trim(tHlp2), "^", "")
                    Case 6
                        nmItem = Replace(Trim(tHlp2), "^", "")
                    Case 5
                        qty = CDbl(Replace(tHlp2, "^", ""))
                    Case 4
                        Harga = CDbl(Replace(tHlp2, "^", ""))
                    Case 3
                        oli = IIf(Replace(tHlp2, "^", ""), 1, 0)
                    Case 2
                        noInv = Trim(Replace(tHlp2, "^", ""))
                    Case 1
                        disc = CDbl(Replace(tHlp2, "|", ""))
                End Select

                i = i - 1
                tHlp3 = Mid(Trim(tHlp3), Len(tHlp2) + 1, Len(tHlp3) - Len(tHlp2))
            Loop
            'NO_WO, NOINV, KODEITEM, KETERANGAN, QTY, HARGA, OLI, DISC
            Set rsHlp3 = objO.FetchDataHO("SELECT * FROM [" & vNamaTabel & "-PART] WHERE NO_CHASSIS = '" & nRgk & "' AND NO_WO = '" & nWO & "' AND NOINV = '" & nINV & "' AND KODEITEM = '" & kdItem & "'")
            If Not rsHlp3.EOF Then
                Call objO.UpdateDataHO("UPDATE [" & vNamaTabel & "-PART] SET KETERANGAN = '" & nmItem & "', QTY = '" & qty & "', HARGA = " & Harga & ", OLI = '" & oli & "', DISC = '" & disc & "' WHERE NO_WO = '" & nWO & "' AND NOINV = '" & nINV & "' AND KODEITEM = '" & kdItem & "'")
            Else
                Call objO.InsertDataHO("INSERT INTO [" & vNamaTabel & "-PART] VALUES ('" & nRgk & "', '" & nw & "','" & noInv & "', '" & kdItem & "','" & nmItem & "','" & qty & "', " & Harga & ", '" & oli & "', '" & disc & "')")
            End If
            tmPART = Mid(Trim(tmPART), Len(tHlp) + 1, Len(tmPART) - Len(tHlp))
        Loop
    End If
End Sub

Private Function createTemp(vNamaTabel As String, Index As Integer) As Boolean
Dim rsHlp As ADODB.Recordset

    createTemp = False
    Select Case Index
        Case 1      ' TABEL UTAMA
            Set rsHlp = objO.FetchDataHO("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
            If Not rsHlp.EOF Then
                createTemp = True
            Else
                Call BuatTabelRepoWarHO(vNamaTabel)
                createTemp = True
            End If
        Case 4
            Set rsHlp = objO.FetchDataHO("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
            If Not rsHlp.EOF Then
                createTemp = True
            Else
                Call BuatTabelRepoSSCHO(vNamaTabel)
                createTemp = True
            End If
        Case 2      ' TABEL PART
            Set rsHlp = objO.FetchDataHO("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
            If Not rsHlp.EOF Then
                createTemp = True
                Call objO.DeleteDataHO("DELETE [TEMPDB]..[" & vNamaTabel & "]")
            Else
                Call BuatTabelPartWarHO(vNamaTabel)
                createTemp = True
            End If
        Case 3      ' TABEL SUBLET
            Set rsHlp = objO.FetchDataHO("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
            If Not rsHlp.EOF Then
                createTemp = True
                Call objO.DeleteDataHO("DELETE [TEMPDB]..[" & vNamaTabel & "]")
            Else
                Call BuatTabelSubletWarHO(vNamaTabel)
                createTemp = True
            End If
    End Select
End Function

Private Sub BuatTabelPartWarHO(vNamaTabel As String)

    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateDataHO(sSQL)
        
    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ( " _
         & "        [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "        [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "        [NOINV][varchar](10) NOT NULL DEFAULT '', " _
         & "        [KODEITEM][varchar](25) NOT NULL DEFAULT '', " _
         & "        [KETERANGAN][varchar](50) NOT NULL DEFAULT '', " _
         & "        [QTY][numeric](9) NOT NULL DEFAULT 0, " _
         & "        [HARGA][money] NOT NULL DEFAULT 0, " _
         & "        [OLI][char](1) NOT NULL DEFAULT 0, " _
         & "        [DISC][decimal](18,2) NOT NULL DEFAULT 0 " _
         & ") ON [PRIMARY]"
    Call objO.UpdateDataHO(sSQL)
End Sub

Private Sub BuatTabelSubletWarHO(vNamaTabel As String)

    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateDataHO(sSQL)
    
    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ( " _
         & "    [NOPO][varchar](20) NOT NULL DEFAULT '', " _
         & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NOINV][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "    [NOFAK][varchar](10) NOT NULL DEFAULT '', " _
         & "    [KODEITEM][varchar](15) NOT NULL DEFAULT '', " _
         & "    [NAMABARANG][varchar](40) NOT NULL DEFAULT '', " _
         & "    [KODESATUAN][varchar](6) NOT NULL DEFAULT '', " _
         & "    [QTY][numeric](9) NOT NULL DEFAULT 0, " _
         & "    [HARGA][money] NOT NULL DEFAULT 0, " _
         & "    [NONOTA][varchar](10) NOT NULL DEFAULT '', " _
         & "    [KETERANGAN][varchar](30) NOT NULL DEFAULT '', " _
         & "    [FLAG][char](1) NOT NULL DEFAULT '' " _
         & ") ON [PRIMARY]"
    Call objO.UpdateDataHO(sSQL)
End Sub

Private Sub BuatTabelRepoSSCHO(vNamaTabel As String)
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateDataHO(sSQL)
    
    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
         & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "    [NO_ENGINE][varchar](15) NOT NULL DEFAULT '', " _
         & "    [MDL][varchar](25) NOT NULL DEFAULT '', " _
         & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NOINV][varchar](10) NOT NULL DEFAULT '', " _
         & "    [DDATE][datetime] NOT NULL DEFAULT '01/01/1900', " _
         & "    [DEALERPENJUAL][varchar](30) NOT NULL DEFAULT '', " _
         & "    [NO_POLISI][varchar](15) NOT NULL DEFAULT '', " _
         & "    [DATE][datetime] NOT NULL DEFAULT '01/01/3000', " _
         & "    [KODETIPE][char](10) NOT NULL DEFAULT '', " _
         & "    [DEALERREPAIR][varchar](30) NOT NULL DEFAULT '', " _
         & "    [REPS][varchar](50) NOT NULL DEFAULT '', " _
         & "    [RTE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [DISC][numeric](9) NOT NULL DEFAULT 0, " _
         & "    [HRGJASA][money] NOT NULL DEFAULT 0, " _
         & "    [PART][varchar](15) NOT NULL DEFAULT '') ON [PRIMARY]"
    Call objO.UpdateDataHO(sSQL)
End Sub

Private Sub BuatTabelRepoWarHO(vNamaTabel As String)
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateDataHO(sSQL)
    
    sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
         & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NOINV][varchar](10) NOT NULL DEFAULT '', " _
         & "    [DDATE][datetime] NOT NULL DEFAULT '01/01/1900', " _
         & "    [DEALERPENJUAL][varchar](30) NOT NULL DEFAULT '', " _
         & "    [NO_POLISI][varchar](15) NOT NULL DEFAULT '', " _
         & "    [DATE][datetime] NOT NULL DEFAULT '01/01/3000', " _
         & "    [KODETIPE][char](1) NOT NULL DEFAULT '', " _
         & "    [DEALERREPAIR][varchar](30) NOT NULL DEFAULT '', " _
         & "    [REPS][varchar](50) NOT NULL DEFAULT '', " _
         & "    [RTE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [DISC][numeric](9) NOT NULL DEFAULT 0, [HRGJASA][money] NOT NULL DEFAULT 0) ON [PRIMARY]"
    Call objO.UpdateDataHO(sSQL)
End Sub

Private Sub Text_GotFocus(Index As Integer)
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0, 2
                gMainKeyCari = text(Index)
                gBantuan = "INTERNAL"
                gSearchOption = "DEALER"
                frmCariHO.Show (vbModal)
            Case 4, 6
                gMainKeyCari = text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCariHO.Show (vbModal)
            Case 8
                gMainKeyCari = text(Index)
                gSearchOption = "WARRANTY"
                frmCariHO.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
1         Select Case Index
              Case 0, 2
2                 If text(Index) <> "" Then
3                     Set oRs = objO.FetchDataHO("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & text(Index) & "' AND INTERNAL = 1")
4                     If oRs.EOF Then
5                         gMainKeyCari = text(Index)
6                         gBantuan = "INTERNAL"
7                         gSearchOption = "DEALER"
8                         frmCariHO.Show (vbModal)
9                     Else
10                        text(Index + 1) = oRs("FS_NM_DEALER")
11                    End If
12                End If
13            Case 4, 6
14                If text(Index) <> "" Then
15                    Set oRs = objO.FetchDataHO("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & text(Index) & "'") ' AND INTERNAL = 0")
16                    If oRs.EOF Then
17                        gMainKeyCari = text(Index)
18                        gBantuan = "ALL"
19                        gSearchOption = "DEALER"
20                        frmCariHO.Show (vbModal)
21                    Else
22                        text(Index + 1) = oRs("FS_NM_DEALER")
23                    End If
24                End If
25            Case 8
26                If text(Index) <> "" Then
                      
27                    If Log.RPLHistori = 0 Then
28                        Set oRs = objO.FetchDataHO("SELECT * FROM HO_WARRANTY WITH(NOLOCK) WHERE ID_WARRANTY = '" & text(Index) & "' AND FS_LOG_DEL = ''")
29                    Else
30                        Set oRs = objO.FetchDataHO("SELECT * FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH(NOLOCK) WHERE ID_WARRANTY = '" & text(Index) & "' AND FS_LOG_DEL = ''")
31                    End If
32                    If oRs.EOF Then
33                        gMainKeyCari = text(Index)
34                        gSearchOption = "WARRANTY"
35                        frmCariHO.Show (vbModal)
36                    Else
37                        text(9) = oRs("WARRANTY")
38                        text(10) = oRs("KETERANGAN")
39                    End If
40                End If
41        End Select
End Sub
