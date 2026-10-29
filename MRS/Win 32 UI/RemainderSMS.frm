VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form RemainderSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: SMS ::."
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RemainderSMS.frx":0000
   ScaleHeight     =   4920
   ScaleWidth      =   9750
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox fot 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Footer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   34
      Top             =   1500
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   600
      TabIndex        =   22
      Top             =   1920
      Width           =   4215
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Creat&e"
         Height          =   735
         Left            =   3240
         Picture         =   "RemainderSMS.frx":DA26
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H80000000&
         Height          =   975
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   975
      End
      Begin VB.Shape Shape8 
         Height          =   1455
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "#NM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "#AL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "#NP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "#TG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "#KM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   ": Nama Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   ": Alamat Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   ": Nomor Polisi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   ": Tanggal Next Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   ": KM Next Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   23
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clea&r"
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
      Left            =   8640
      Picture         =   "RemainderSMS.frx":E073
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   480
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
      Left            =   8740
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
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
      Left            =   7580
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Send To : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1800
      TabIndex        =   8
      Top             =   3600
      Width           =   3015
      Begin VB.CheckBox chGSM 
         Caption         =   "GSM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox chCDMA 
         Caption         =   "CDMA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbGSM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "085640528855"
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
         Left            =   1320
         TabIndex        =   12
         Top             =   315
         Width           =   1440
      End
      Begin VB.Label lbCDMA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "088888888888"
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
         Left            =   1320
         TabIndex        =   11
         Top             =   675
         Width           =   1440
      End
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
      Picture         =   "RemainderSMS.frx":E68C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Send"
      Height          =   732
      Left            =   5040
      Picture         =   "RemainderSMS.frx":EE1C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdTemplate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Template"
      Height          =   732
      Left            =   720
      Picture         =   "RemainderSMS.frx":F489
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtPesan1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   600
      MaxLength       =   70
      TabIndex        =   2
      Top             =   370
      Width           =   7695
   End
   Begin VB.TextBox txtPesan2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   600
      MaxLength       =   70
      TabIndex        =   1
      Top             =   720
      Width           =   7695
   End
   Begin VB.TextBox txtPesan3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   600
      MaxLength       =   70
      TabIndex        =   0
      Top             =   1065
      Width           =   7695
   End
   Begin TDBNumber6Ctl.TDBNumber txtHargaJasa 
      Height          =   330
      Left            =   7560
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   582
      Calculator      =   "RemainderSMS.frx":FADE
      Caption         =   "RemainderSMS.frx":FAFE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "RemainderSMS.frx":FB6A
      Keys            =   "RemainderSMS.frx":FB88
      Spin            =   "RemainderSMS.frx":FBD2
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
      Left            =   7560
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   582
      Calculator      =   "RemainderSMS.frx":FBFA
      Caption         =   "RemainderSMS.frx":FC1A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "RemainderSMS.frx":FC86
      Keys            =   "RemainderSMS.frx":FCA4
      Spin            =   "RemainderSMS.frx":FCEE
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
   Begin VB.Line Line9 
      X1              =   8520
      X2              =   9480
      Y1              =   1400
      Y2              =   1400
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Est Harga Part    :   Rp."
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
      Left            =   5280
      TabIndex        =   21
      Top             =   2540
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000000&
      Height          =   975
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   975
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   8280
      X2              =   600
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lbLen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Char :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   1520
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      Height          =   380
      Left            =   8725
      Top             =   2865
      Visible         =   0   'False
      Width           =   760
   End
   Begin VB.Shape Shape5 
      Height          =   380
      Left            =   7565
      Top             =   2865
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   9720
      X2              =   9720
      Y1              =   0
      Y2              =   4880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estimasi Waktu  :                       Jam              Menit"
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
      Left            =   5280
      TabIndex        =   15
      Top             =   2925
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Est Harga Jasa   :   Rp."
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
      Left            =   5280
      TabIndex        =   14
      Top             =   2175
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   9600
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000000&
      Height          =   975
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000000&
      Height          =   975
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000000&
      Height          =   975
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label42 
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
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label54 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   75
   End
   Begin VB.Shape Shape1 
      Height          =   1050
      Left            =   585
      Top             =   360
      Width           =   7740
   End
   Begin VB.Line Line8 
      BorderStyle     =   4  'Dash-Dot
      X1              =   600
      X2              =   8280
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line7 
      BorderStyle     =   4  'Dash-Dot
      X1              =   600
      X2              =   8280
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   480
      X2              =   9720
      Y1              =   45
      Y2              =   45
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   480
      X2              =   9720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   480
      X2              =   9720
      Y1              =   4860
      Y2              =   4860
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   480
      X2              =   9720
      Y1              =   4905
      Y2              =   4905
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   35
      Top             =   1440
      Width           =   7695
   End
End
Attribute VB_Name = "RemainderSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private Nama        As String        '==> #NM
Private Alamat      As String        '==> #AL
Private NoPolisi    As String        '==> #NP
Private TglNext     As Date          '==> #TG
Private KmNext      As Double        '==> #KM
Private GSM         As String
Private CDMA        As String
Private Harga       As Double        '==> #RP
Private Waktu       As String        '==> #JM
Private noRangka    As String
Private LastNoWo    As String

Private Remainder   As String
Private objR        As New ClsRemainder

Private rsRem       As ADODB.Recordset
Private isSSC       As Integer

Private nLen        As Integer
Private nSMS        As String
Private objPul      As New ClsRemainder

Public Property Let xRsRem(value As ADODB.Recordset)
   Set rsRem = value
End Property

Public Property Let IsLoadSSC(value As Integer)
   isSSC = value
End Property


Private Sub cmdCreate_Click()
On Error Resume Next
   Dim hrg As String
   hrg = "Jasa : Rp." & txtHargaJasa.Text & ",-  Part/Oli : Rp." & txtHargaPart.Text & ",-"
   Harga = txtHargaJasa + txtHargaPart
   Waktu = cboJam.Text & ":" & cboMenit.Text
   Dim MyText As String
   Dim IsText As String
   IsText = Me.txtPesan1.Text & " " & Me.txtPesan2.Text & " " & Me.txtPesan3.Text
   MyText = Replace(IsText, "#NM", Nama)
   MyText = Replace(MyText, "#AL", Alamat)
   MyText = Replace(MyText, "#NP", NoPolisi)
   MyText = Replace(MyText, "#TG", TglNext)
   MyText = Replace(MyText, "#KM", KmNext)
   MyText = Replace(MyText, "#RP", hrg)
   MyText = Replace(MyText, "#JM", Waktu)
   Dim i As Integer
   For i = 0 To 5
      MyText = Replace(MyText, "  ", " ")
   Next i
   With Me
      .txtPesan1.Text = Left(MyText, 70)
      .txtPesan2.Text = Mid(MyText, 71, 140)
      .txtPesan3.Text = Mid(MyText, 141, 210)
   End With
End Sub

Private Sub cmdClear_Click()
   With Me
      .txtPesan1.Text = ""
      .txtPesan2.Text = ""
      .txtPesan3.Text = ""
      .fot.value = 0
   End With
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSend_Click()
   Call cmdCreate_Click
   Dim Er As New ErrInfo
   If nLen > 160 Then
        Er.HeadCap = "Character SMS"
        Er.InfoCap = "Format SMS maksimal 160 character !"
        Beep
        Er.Show (vbModal)
      Exit Sub
   End If
   Dim isPesan As String
   Dim isTelp As String
   isPesan = txtPesan1.Text & " " & txtPesan2.Text & " " & txtPesan3.Text
   If isPesan = "" Then
        Er.HeadCap = "SMS !"
        Er.InfoCap = "Pesan Masih Kosong !"
        Beep
        Er.Show (vbModal)
        Exit Sub
   End If
   If chGSM.value = 0 And chCDMA.value = 0 Then
        Er.HeadCap = "SMS !"
        Er.InfoCap = "Nomor Telp Tujuan Belum Dipilih !"
        Beep
        Er.Show (vbModal)
      Exit Sub
   End If
   If chGSM.value = 1 Then
      isTelp = lbGSM.Caption
   ElseIf chCDMA.value = 1 Then
      isTelp = lbCDMA.Caption
   End If
   Dim milix As String
   If isSSC = 0 Then
      Call objR.AddNew_LogRemainder("SM", noRangka, LastNoWo, TglNext, KmNext, isPesan, isTelp, 0, txtHargaJasa.Text, txtHargaPart.Text, cboJam.Text, cboMenit.Text)
      milix = "SBE NEXT"
   ElseIf isSSC = 1 Then
      Call objR.AddNew_LogRemainder("SC", noRangka, LastNoWo, TglNext, 0, isPesan, isTelp, 0, 0, 0, cboJam.Text, cboMenit.Text)
      milix = "SSC"
   ElseIf isSSC = 2 Then
      Call objR.AddNew_LogRemainder("1K", noRangka, LastNoWo, TglNext, 1000, isPesan, isTelp, 0, 0, 0, cboJam.Text, cboMenit.Text)
      milix = "SBE 1000"
   End If
   
    Dim NewId As Double
    Dim rsSMS As ADODB.Recordset
    Set rsSMS = objR.GetDataOutbox
    
    If rsSMS.RecordCount = 0 Then
        NewId = "1"
    Else
        rsSMS.MoveLast
        NewId = Val(rsSMS.Fields(0).value) + 1
    End If
    
   If chGSM.value = 1 Then
        With objR
            .NoSMS = NewId
            .pesan = isPesan
            .Tujuan = lbGSM.Caption
            .Waktu = isnow
            .AddNewItemSMS
        End With
        NewId = NewId + 1
    End If
   If chCDMA.value = 1 Then
        With objR
            .NoSMS = NewId
            .pesan = isPesan
            .Tujuan = lbCDMA.Caption
            .Waktu = isnow
            .AddNewItemSMS
        End With
   End If
        
    Call objPul.UpdatePulsa
        Er.HeadCap = "SMS !"
        Er.InfoCap = "Pesan telah dikirim !"
        Beep
        Er.Show (vbModal)
   Unload Me
End Sub

Private Sub cmdTemplate_Click()
   Dim ft As New FormTemplateSMS
   ft.OnLoaded = True
   ft.Show vbModal, Me
   Remainder = ft.Remainder
   nSMS = Remainder
   
   With Me
      .txtPesan1.Text = Left(Remainder, 70)
      .txtPesan2.Text = Mid(Remainder, 71, 140)
      .txtPesan3.Text = Mid(Remainder, 141, 210)
   End With
End Sub

Private Sub Form_Load()
   Call BindToControl
   If GSM <> "" Then
      lbGSM.Caption = GSM
      chGSM.Enabled = True
   Else
      lbGSM.Caption = "-"
      chGSM.Enabled = False
   End If
   If CDMA <> "" Then
      lbCDMA.Caption = CDMA
      chCDMA.Enabled = True
   Else
      lbCDMA.Caption = "-"
      chCDMA.Enabled = False
   End If
   Call ListComboWaktu
   Call Eta
   
   Remainder = nSMS
   Me.fot.Caption = Log.LogNas
   With Me
      .txtPesan1.Text = Left(Remainder, 70)
      .txtPesan2.Text = Mid(Remainder, 71, 140)
      .txtPesan3.Text = Mid(Remainder, 141, 210)
   End With
End Sub

Private Sub BindToControl()
   If isSSC = 0 Then
      Nama = Trim(NullText(rsRem.Fields(6).value))
      Alamat = Trim(NullText(rsRem.Fields(7).value & " " & rsRem.Fields(8).value & " " & rsRem.Fields(9).value))
      NoPolisi = Trim(NullText(rsRem.Fields(10).value))
      TglNext = Format(rsRem.Fields(4).value, "dd/mm/yyyy")
      KmNext = Val(rsRem.Fields(5).value)
      GSM = Trim(NullText(rsRem.Fields(14).value))
      CDMA = Trim(NullText(rsRem.Fields(15).value))
      noRangka = Trim(NullText(rsRem.Fields(12).value))
      LastNoWo = Trim(NullText(rsRem.Fields(1).value))
      Frame1.Visible = True
   ElseIf isSSC = 1 Then
      Nama = Trim(NullText(rsRem.Fields(0).value))
      Alamat = Trim(NullText(rsRem.Fields(1).value & " " & rsRem.Fields(2).value & " " & rsRem.Fields(3).value))
      NoPolisi = Trim(NullText(rsRem.Fields(4).value))
      TglNext = IIf(Format(rsRem.Fields(16).value, "dd/mm/yyyy") = Format(rsRem.Fields(15).value, "dd/mm/yyyy"), 1 / 1 / 9999, Format(rsRem.Fields(16).value, "dd/mm/yyyy"))
      KmNext = 0
      GSM = Trim(NullText(rsRem.Fields(8).value))
      CDMA = Trim(NullText(rsRem.Fields(9).value))
      noRangka = Trim(NullText(rsRem.Fields(6).value))
      LastNoWo = Trim(NullText(rsRem.Fields(10).value))
      Frame1.Visible = True
   ElseIf isSSC = 2 Then
      Nama = Trim(NullText(rsRem.Fields(6).value))
      Alamat = Trim(NullText(rsRem.Fields(7).value & " " & rsRem.Fields(8).value & " " & rsRem.Fields(9).value))
      NoPolisi = Trim(NullText(rsRem.Fields(10).value))
      TglNext = Format(rsRem.Fields(30).value, "dd/mm/yyyy")
      KmNext = 1000
      GSM = Trim(NullText(rsRem.Fields(14).value))
      CDMA = Trim(NullText(rsRem.Fields(15).value))
      noRangka = Trim(NullText(rsRem.Fields(12).value))
      LastNoWo = Trim(NullText(rsRem.Fields(0).value))
      Frame1.Visible = True
   End If
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

Private Sub Eta()
   Dim rsEta As ADODB.Recordset
   Set rsEta = objR.FetchEta(LastNoWo, "DM")
   If rsEta.RecordCount = 0 Then Exit Sub
   With Me
      .txtHargaJasa = rsEta.Fields(0).value
      .txtHargaPart = rsEta.Fields(1).value
      .cboJam.Text = rsEta.Fields(2).value
      .cboMenit.Text = rsEta.Fields(3).value
   End With
End Sub

Private Sub lbGSM_Click()
   If Me.chGSM.Enabled = True And Me.lbGSM.Caption <> "-" Then
      If Me.chGSM.value = 1 Then
         chGSM.value = 0
      Else
         chGSM.value = 1
      End If
   End If
End Sub

Private Sub lbCDMA_Click()
   If Me.chCDMA.Enabled = True And Me.lbCDMA.Caption <> "-" Then
      If Me.chCDMA.value = 1 Then
         chCDMA.value = 0
      Else
         chCDMA.value = 1
      End If
   End If
End Sub

Private Sub txtPesan1_Change()
   If Len(txtPesan1.Text) > 59 Then txtPesan2.SetFocus
   Call LenChr
End Sub
Private Sub txtPesan2_Change()
   If Len(txtPesan2.Text) > 59 Then txtPesan3.SetFocus
   Call LenChr
End Sub
Private Sub txtPesan3_Change()
   Call LenChr
End Sub
Private Sub fot_Click()
   Call LenChr
End Sub

Private Sub LenChr()
   nLen = Val(Len(txtPesan1.Text)) + Val(Len(txtPesan2.Text)) + Val(Len(txtPesan3.Text))
   If fot.value = 1 Then nLen = (nLen + Len(fot.Caption)) - 7
   lbLen.Caption = "Char : " & nLen
End Sub

Private Sub txtPesan2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If Len(txtPesan2.Text) = 0 Then
            txtPesan1.SetFocus
            SendKeys ("{END}")
        End If
    End If
End Sub
Private Sub txtPesan3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        If Len(txtPesan3.Text) = 0 Then
            txtPesan2.SetFocus
            SendKeys ("{END}")
        End If
    End If
End Sub
