VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form RemainderDM 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "RemainderDM.frx":0000
   ScaleHeight     =   7560
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPesan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   2160
      MaxLength       =   999
      MultiLine       =   -1  'True
      TabIndex        =   43
      Top             =   4560
      Width           =   7455
   End
   Begin VB.CommandButton cmdTemplate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "RemainderDM.frx":D943
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtTelpon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      MaxLength       =   70
      TabIndex        =   37
      Top             =   1560
      Width           =   7455
   End
   Begin VB.TextBox txtNoPol 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      MaxLength       =   70
      TabIndex        =   33
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtSBEBulan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      MaxLength       =   70
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtSBEKm 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      MaxLength       =   70
      TabIndex        =   17
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ComboBox cboJam 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3960
      Width           =   735
   End
   Begin VB.ComboBox cboMenit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtAlamat2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      MaxLength       =   70
      TabIndex        =   4
      Top             =   1200
      Width           =   7455
   End
   Begin VB.TextBox txtAlamat1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      MaxLength       =   70
      TabIndex        =   3
      Top             =   840
      Width           =   7455
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      MaxLength       =   70
      TabIndex        =   2
      Top             =   480
      Width           =   7455
   End
   Begin VB.CommandButton cmdPrint1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   732
      Left            =   7680
      Picture         =   "RemainderDM.frx":DF3C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   732
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      Picture         =   "RemainderDM.frx":E4C2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6360
      Width           =   735
   End
   Begin TDBNumber6Ctl.TDBNumber txtHargaJasa 
      Height          =   330
      Left            =   7320
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   582
      Calculator      =   "RemainderDM.frx":EC52
      Caption         =   "RemainderDM.frx":EC72
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "RemainderDM.frx":ECDE
      Keys            =   "RemainderDM.frx":ECFC
      Spin            =   "RemainderDM.frx":ED46
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1397882885
      MinValueVT      =   1112342533
   End
   Begin TDBNumber6Ctl.TDBNumber txtHargaPart 
      Height          =   330
      Left            =   7320
      TabIndex        =   27
      Top             =   2760
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   582
      Calculator      =   "RemainderDM.frx":ED6E
      Caption         =   "RemainderDM.frx":ED8E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "RemainderDM.frx":EDFA
      Keys            =   "RemainderDM.frx":EE18
      Spin            =   "RemainderDM.frx":EE62
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1397882885
      MinValueVT      =   1112342533
   End
   Begin TDBNumber6Ctl.TDBNumber txtNextKm 
      Height          =   330
      Left            =   2160
      TabIndex        =   36
      Top             =   2760
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   582
      Calculator      =   "RemainderDM.frx":EE8A
      Caption         =   "RemainderDM.frx":EEAA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "RemainderDM.frx":EF16
      Keys            =   "RemainderDM.frx":EF34
      Spin            =   "RemainderDM.frx":EF7E
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0)"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   94896129
      Value           =   0
      MaxValueVT      =   1397882885
      MinValueVT      =   1112342533
   End
   Begin TDBNumber6Ctl.TDBNumber txtChemical 
      Height          =   330
      Left            =   7320
      TabIndex        =   44
      Top             =   3120
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   582
      Calculator      =   "RemainderDM.frx":EFA6
      Caption         =   "RemainderDM.frx":EFC6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "RemainderDM.frx":F032
      Keys            =   "RemainderDM.frx":F050
      Spin            =   "RemainderDM.frx":F09A
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1397882885
      MinValueVT      =   1112342533
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chemical"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5880
      TabIndex        =   46
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   45
      Top             =   3120
      Width           =   315
   End
   Begin VB.Shape ShFalse 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   975
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   9600
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   41
      Top             =   4200
      Width           =   75
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Isi Pesan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   40
      Top             =   4560
      Width           =   780
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   39
      Top             =   1560
      Width           =   75
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telepon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   38
      Top             =   1560
      Width           =   675
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Polisi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   35
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   34
      Top             =   2400
      Width           =   75
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   4320
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Service Kendaraan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   600
      TabIndex        =   32
      Top             =   2040
      Width           =   2280
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   5880
      X2              =   9480
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estimasi Waktu Pengerjaan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5880
      TabIndex        =   31
      Top             =   3600
      Width           =   2625
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   30
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   29
      Top             =   2760
      Width           =   75
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parts/Oli"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   28
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   26
      Top             =   2400
      Width           =   315
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6840
      TabIndex        =   25
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jasa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   24
      Top             =   2400
      Width           =   390
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   5880
      X2              =   9480
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   23
      Top             =   3480
      Width           =   525
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   21
      Top             =   3120
      Width           =   300
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   20
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SBE "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   19
      Top             =   3120
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   18
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next Service"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   16
      Top             =   2760
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   15
      Top             =   2760
      Width           =   75
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   9600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   14
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   13
      Top             =   840
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   11
      Top             =   480
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estimasi Harga  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu  :                       Jam                   Menit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5880
      TabIndex        =   9
      Top             =   3960
      Width           =   3420
   End
   Begin VB.Shape Shape5 
      Height          =   345
      Left            =   6825
      Top             =   3960
      Width           =   765
   End
   Begin VB.Shape Shape6 
      Height          =   345
      Left            =   8025
      Top             =   3960
      Width           =   765
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tujuan DM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   1035
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   600
      X2              =   9840
      Y1              =   7380
      Y2              =   7380
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      Height          =   975
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000000&
      Height          =   975
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   975
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   9600
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   9720
      X2              =   9720
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   480
      X2              =   9720
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   480
      X2              =   9720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   480
      X2              =   9720
      Y1              =   50
      Y2              =   50
   End
End
Attribute VB_Name = "RemainderDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private TglNext As Date
Private TglLast As Date
Private noRangka As String
Private LastNoWo As String
Private objR As New ClsRemainder
Private rsDM As New ADODB.Recordset
Private rsRem As ADODB.Recordset
Private dr1 As ReportDM
Private sbi_1K As Boolean
Private support As New clsFormDealer
Private ft As New FormTemplateSMS
Private Remainder As String
Private cetakKM   As String

Public Property Let xRsRem(value As ADODB.Recordset)
   Set rsRem = value
End Property

Public Property Let xSBI(value As Boolean)
   sbi_1K = value
End Property

Public Property Let dataKM(value As String)
   cetakKM = value
End Property

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdTemplate_Click()
    ft.OnLoaded = True
    ft.Show vbModal
    Remainder = ft.Remainder
    txtPesan.Text = Remainder
End Sub

Private Sub Form_Load()
   Call BindToControl
   Call ListComboWaktu
   Call SetEstimasiService
End Sub

Private Sub BindToControl()
   Me.txtNama.Text = Trim(NullText(rsRem.Fields(6).value))
   Me.txtAlamat1.Text = Trim(NullText(rsRem.Fields(7).value))
   Me.txtAlamat2.Text = Trim(NullText(rsRem.Fields(8).value & " " & rsRem.Fields(9).value))
   Me.txtNoPol.Text = Trim(NullText(rsRem.Fields(10).value))
   Me.txtNextKm = Val(rsRem.Fields(5).value)
   
   If Trim(NullText(rsRem.Fields(15).value)) <> "" Then
        Me.txtTelpon = Trim(NullText(rsRem.Fields(14).value)) & "; " & Trim(NullText(rsRem.Fields(15).value))
   Else
        Me.txtTelpon = Trim(NullText(rsRem.Fields(14).value))
   End If
        If Nul(rsRem.Fields(4).value) = "" Then
            TglNext = DateAdd("d", 30, rsRem.Fields(2).value)
        Else
            TglNext = Format(rsRem.Fields(4).value, "dd/mm/yyyy")
        End If
      noRangka = Trim(NullText(rsRem.Fields(12).value))
      LastNoWo = IIf(sbi_1K, Trim(NullText(rsRem.Fields(0).value)), Trim(NullText(rsRem.Fields(1).value)))
      TglLast = Trim(NullText(rsRem.Fields(2).value))
End Sub

Private Sub ListComboWaktu()
   With Me
      .cboJam.Clear
      Dim i As Integer
      For i = 0 To 5
         .cboJam.AddItem i
      Next i
      .cboMenit.Clear
      For i = 0 To 50 Step 5
         .cboMenit.AddItem i
      Next i
      .cboJam.Text = 0
      .cboMenit.Text = 0
   End With
End Sub

Private Sub SetEstimasiService()
Dim obj As New ClsRemainder
Dim rsA As ADODB.Recordset
Dim rsB As ADODB.Recordset
Dim rsC As ADODB.Recordset
Dim rsD As ADODB.Recordset
Dim rsH As ADODB.Recordset
Dim HrgPart As Double
Dim Rate As Double
Dim Price As Double
Dim MDL As String
Dim FsKet As String
Dim jam As Integer

On Error Resume Next
    Set rsA = obj.FetchHargaJasa(noRangka, True, sbi_1K)
    Price = CDbl(rsA.Fields(0).value)
    MDL = Trim(NullText(rsA.Fields(1).value))
    
    '    Set rsB = obj.FetchHargaJasa(noRangka, False, sbi_1K)
'    FsKet = Trim(NullText(rsB.Fields(0).value))
    FsKet = cetakKM
    txtSBEKm.Text = FsKet
    ''2016-09-22 || ambil km dari form reminder service, 10.000 / 20.000
    
        
        
    Set rsC = obj.FetchRate(MDL, FsKet)
    Rate = CDbl(rsC.Fields(0).value)
    Me.txtHargaJasa = Price * Rate
        
    Set rsD = obj.FetchItemPart(MDL, FsKet)
    rsD.MoveFirst
    Do Until rsD.EOF
       HrgPart = Val(HrgPart) + (HargaJualPart(rsD.Fields(0).value) * rsD.Fields(1).value)
       rsD.MoveNext
    Loop
    Me.txtHargaPart = HrgPart
    
    Me.cboJam = Left(Rate, 1)
    Me.cboMenit = (Rate * 60) Mod 60
    
    If sbi_1K Then
         txtSBEBulan.Text = "1.0"
    Else
         Set rsH = obj.FetchHeavyUse(noRangka)
         If rsH.Fields(0).value = "1" Then
            txtSBEBulan.Text = "3.0"
         ElseIf rsH.Fields(0).value = "0" Then
            txtSBEBulan.Text = "6.0"
         End If
    End If
End Sub

Private Sub cmdPrint1_Click()
Dim boleh As Boolean

    boleh = False
    If Trim(txtPesan) <> "" Then
        boleh = True
    ElseIf Trim(txtPesan) = "" Then
        If MsgBox("Cetak DM tanpa Catatan ?", vbYesNo + vbDefaultButton1, "Cetak DM") = vbYes Then
            boleh = True
        End If
    End If
    
   If MsgBox("Cetak DM untuk customer : " & txtNama.Text & " :  " & txtNoPol.Text, vbQuestion + vbYesNo, "Cetak Direct Mail !") = vbYes And boleh Then
      Dim pesan As String
      pesan = txtNama.Text & " - " & txtAlamat1.Text & " " & txtAlamat2.Text
      Call objR.AddNew_LogRemainder("DM", noRangka, LastNoWo, TglNext, txtNextKm, txtNama.Text, txtAlamat1.Text & " " & txtAlamat2.Text, 0, txtHargaJasa, txtHargaPart, IIf(cboJam.Text = "", 0, Val(cboJam.Text)), IIf(cboMenit.Text = "", 0, Val(cboMenit.Text)))
      
      Set dr1 = New ReportDM
      CreateTblDM True
      rsDM.AddNew
      rsDM!noRangka = Trim(rsRem!noChassis)
      rsDM.Update
              
        dr1.txtNama.SetText Me.txtNama.Text
        dr1.txtAlamat1.SetText Me.txtAlamat1.Text
        dr1.txtAlamat2.SetText Me.txtAlamat2.Text
        dr1.txtNextKm.SetText Me.txtSBEKm.Text
        dr1.txtNextBln.SetText Me.txtSBEBulan.Text
        dr1.txtJasa.SetText IIf(txtSBEKm = "1.000", "-- FREE --", Me.txtHargaJasa.Text)
        dr1.txtPart.SetText IIf(txtSBEKm = "1.000", "", Me.txtHargaPart.Text)
        dr1.txtEtaJam.SetText Me.cboJam.Text & " : " & Me.cboMenit.Text
        dr1.txtNoPol.SetText Me.txtNoPol.Text
        dr1.txtTelp.SetText Me.txtTelpon.Text
        dr1.txtPerush.SetText support.kopOrder(1)
        dr1.txtAlamatPerush.SetText support.kopOrder(4)
        dr1.txtAlamatPerush2.SetText support.kopOrder(5)
        dr1.txtTelpPerush.SetText support.kopOrder(6)
        dr1.txtTelpPerush2.SetText support.kopOrder(3)
        dr1.txtCatatan.SetText Me.txtPesan.Text
        'tambahan chemical
        dr1.txtChemical.SetText txtChemical.Text
        
        
        With FrmctkOrder.CRViewer1
            .ReportSource = dr1
            .ViewReport
            .Zoom 50 '100
            .Visible = True
        End With
        FrmctkOrder.cmdTutup.Visible = True
        Screen.MousePointer = vbDefault
        FrmctkOrder.WindowState = vbMaximized
        FrmctkOrder.Show vbModal
      
   End If
End Sub

Private Sub CreateTblDM(Optional Edit As Boolean = False)
1         On Error GoTo vb_error
2         If rsDM.State Then rsDM.Close
3         Set rsDM = Nothing
4         With rsDM
5             .Fields.Append "NoRangka", adVarChar, 30
6             .Open
7         End With
8         Exit Sub
vb_error:
9         err.Raise err.Number, err.Source, err.Description
End Sub
