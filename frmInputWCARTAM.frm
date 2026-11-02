VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmInputWCARTAM 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Input WCAR TAM"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14850
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmInputWCARTAM.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   28
      Top             =   0
      Width           =   450
   End
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
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   14640
      TabIndex        =   8
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
   Begin VB.CheckBox chAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tampilkan belum input Claim dan WCAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   960
      Width           =   4815
   End
   Begin VB.Frame frmeKonfirm 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   4680
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   5120
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   0
         Top             =   2400
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   2
         Left            =   1440
         TabIndex        =   1
         Top             =   2760
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   503
         _Version        =   393216
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
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   1440
         TabIndex        =   26
         Top             =   3240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&s"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInputWCARTAM.frx":8C6C
         PICN            =   "frmInputWCARTAM.frx":8C88
         PICH            =   "frmInputWCARTAM.frx":919B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdBatal 
         Height          =   495
         Left            =   2640
         TabIndex        =   27
         Top             =   3240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   2
         TX              =   "&b"
         ENAB            =   -1  'True
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInputWCARTAM.frx":973D
         PICN            =   "frmInputWCARTAM.frx":9759
         PICH            =   "frmInputWCARTAM.frx":9C8E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   29
         Top             =   480
         Width           =   3435
         _ExtentX        =   6059
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
         Left            =   1440
         TabIndex        =   31
         Top             =   840
         Width           =   3435
         _ExtentX        =   6059
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
         Left            =   1440
         TabIndex        =   33
         Top             =   1200
         Width           =   3435
         _ExtentX        =   6059
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
         Left            =   1440
         TabIndex        =   35
         Top             =   1560
         Width           =   3435
         _ExtentX        =   6059
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
         Left            =   1440
         TabIndex        =   37
         Top             =   1920
         Width           =   3435
         _ExtentX        =   6059
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   6
         Left            =   585
         TabIndex        =   38
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   4
         Left            =   495
         TabIndex        =   36
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   3
         Left            =   690
         TabIndex        =   34
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Inv :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   2
         Left            =   795
         TabIndex        =   32
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   1
         Left            =   765
         TabIndex        =   30
         Top             =   480
         Width           =   615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         Index           =   1
         X1              =   0
         X2              =   5400
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   5160
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         Index           =   0
         X1              =   -240
         X2              =   5160
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: DATA PENAGIH ::."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1560
         TabIndex        =   14
         Top             =   60
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WCAR :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   570
         TabIndex        =   4
         Top             =   2760
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Claim :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   3
         Top             =   2400
         Width           =   720
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   5115
      End
   End
   Begin TrueOleDBGrid70.TDBGrid gridTagih 
      Height          =   6495
      Left            =   480
      TabIndex        =   15
      Top             =   1920
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   11456
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8).NumberFormat=   "FormatText Event"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9).NumberFormat=   "FormatText Event"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=3281"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   14933984
      RowDividerColor =   14933984
      RowSubDividerColor=   14933984
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=86,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=83,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=84,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=85,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=32,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=17"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(91)  =   "Named:id=38:HighlightRow"
      _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=39:EvenRow"
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
      _StyleDefs(101) =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(102) =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(103) =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(104) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(105) =   "Named:id=87:SudahDitagih"
      _StyleDefs(106) =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(107) =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(108) =   ":id=87,.fontname=Tahoma"
   End
   Begin TDBDate6Ctl.TDBDate Tgl1 
      Height          =   330
      Left            =   2280
      TabIndex        =   17
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInputWCARTAM.frx":A245
      Caption         =   "frmInputWCARTAM.frx":A371
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInputWCARTAM.frx":A3DD
      Keys            =   "frmInputWCARTAM.frx":A3FB
      Spin            =   "frmInputWCARTAM.frx":A459
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
   Begin TDBDate6Ctl.TDBDate Tgl2 
      Height          =   330
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInputWCARTAM.frx":A481
      Caption         =   "frmInputWCARTAM.frx":A5AD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInputWCARTAM.frx":A619
      Keys            =   "frmInputWCARTAM.frx":A637
      Spin            =   "frmInputWCARTAM.frx":A695
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
   Begin TDBDate6Ctl.TDBDate Tgl3 
      Height          =   330
      Left            =   2280
      TabIndex        =   21
      Top             =   480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInputWCARTAM.frx":A6BD
      Caption         =   "frmInputWCARTAM.frx":A7E9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInputWCARTAM.frx":A855
      Keys            =   "frmInputWCARTAM.frx":A873
      Spin            =   "frmInputWCARTAM.frx":A8D1
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
   Begin TDBDate6Ctl.TDBDate Tgl4 
      Height          =   330
      Left            =   4320
      TabIndex        =   22
      Top             =   480
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmInputWCARTAM.frx":A8F9
      Caption         =   "frmInputWCARTAM.frx":AA25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInputWCARTAM.frx":AA91
      Keys            =   "frmInputWCARTAM.frx":AAAF
      Spin            =   "frmInputWCARTAM.frx":AB0D
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
   Begin ucXPButton.XPButton cmdRef 
      Height          =   495
      Left            =   5880
      TabIndex        =   16
      Top             =   120
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&r"
      ENAB            =   -1  'True
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInputWCARTAM.frx":AB35
      PICN            =   "frmInputWCARTAM.frx":AB51
      PICH            =   "frmInputWCARTAM.frx":B0CD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdKeluar 
      Height          =   495
      Left            =   13320
      TabIndex        =   23
      Top             =   8520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&x"
      ENAB            =   -1  'True
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInputWCARTAM.frx":B704
      PICN            =   "frmInputWCARTAM.frx":B720
      PICH            =   "frmInputWCARTAM.frx":BBF4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdKonfir 
      Height          =   495
      Left            =   9360
      TabIndex        =   24
      Top             =   8520
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&s"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInputWCARTAM.frx":C13F
      PICN            =   "frmInputWCARTAM.frx":C15B
      PICH            =   "frmInputWCARTAM.frx":C66E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdExport 
      Height          =   495
      Left            =   12120
      TabIndex        =   25
      Top             =   8520
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&o"
      ENAB            =   -1  'True
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInputWCARTAM.frx":CC10
      PICN            =   "frmInputWCARTAM.frx":CC2C
      PICH            =   "frmInputWCARTAM.frx":D1C1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":                           s/d"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   20
      Top             =   480
      Width           =   2025
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Invoice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   19
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: DAFTAR DATA ::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6600
      TabIndex        =   13
      Top             =   1500
      Width           =   2040
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal WO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   900
      TabIndex        =   12
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":                           s/d"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label jmlRecord 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Tagihan :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   12795
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label nTagih 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   2640
      TabIndex        =   6
      Top             =   8595
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Data :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   600
      TabIndex        =   5
      Top             =   8595
      Width           =   1980
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   14055
   End
End
Attribute VB_Name = "frmInputWCARTAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRecord As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns
Private Urut As String
Private jmlTagih As Integer

Private Sub cmdBatal_Click()
    frmeKonfirm.Visible = False
End Sub

Private Sub cmdExport_Click()
On Error GoTo Err_Export
    Dim Er As New ErrInfo
    If rsRecord.RecordCount = 0 Then
        Er.HeadCap = "Export To Excel !"
        Er.InfoCap = "List Data Input WCAR Kosong !"
        Beep
        Er.show (vbModal)
        Exit Sub
    End If
    Dim MyPath As String
    Dim onVal As String
    MyPath = ""
    Dim f As New FormExplorer
    f.show (vbModal)
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    onVal = MyPath & f.FileName & ".XLS"
    
     gridTagih.ExportToFile onVal, True
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            
            Beep
            Er.show (vbModal)
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdKonfir_Click()
    If CInt(nTagih.Caption) < 1 Then
        MsgBox "Anda belum memilih Tagihan Cuci yang akan ditagihkan !", vbExclamation + vbOKOnly
        Exit Sub
    End If
        For Each col In gridTagih.Columns
            col.FilterText = ""
        Next col
        rsRecord.filter = ""

    frmeKonfirm.Visible = True
End Sub

Private Sub cmdRef_Click()
    Call refreshGrid(IIf(chAll.value = Checked, True, False))
    'nTagih.Caption = jmlTagih
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
      Dim namaSupp As String
      Dim rsHlp As New ADODB.Recordset
      Dim isBegin As Boolean

1        On Error GoTo cmdSimpan_Click_Error
2         isBoleh = False
3         boleh = True
4         pesan = "Lengkapi data berikut .... " & vbCr & vbCr
5         If Trim(Text(1)) = "" Then
6             boleh = False
7             pesan = pesan + "No. Claim harus diisi. " & vbCr
8         End If
9         If Trim(Text(2)) = "" Then
10            boleh = False
11            pesan = pesan + "No WCAR. " & vbCr
12        End If
           
13        If boleh Then
14            sSQL = "SELECT * FROM claimfop WITH (NOLOCK) WHERE status <> 'BTL' AND no_wo <> '" & NulSID(rsRecord!noWO) & "' and noInv <> '" & NulSID(rsRecord!noInv) & "' and noclaim = '" & NulSID(Text(1)) & "' and nowcar = '" & NulSID(Text(2)) & "'"
15            Set rsHlp = objO.FetchData(sSQL)
16            If Not rsHlp.EOF Then
17                MsgBox "No Claim = " & NulSID(Text(1)) & vbCrLf & "No WCAR = " & NulSID(Text(2)) & vbCrLf & "Telah digunakan NoWo = " & NulSID(rsHlp!no_wo) & " , Noinv = " & NulSID(rsHlp!noInv), vbInformation
18                Exit Sub
19            Else
20                sSQL = "SELECT * FROM [" & dbReplikasiHO & "]..claimfop WITH (NOLOCK) WHERE status <> 'BTL' AND no_wo <> '" & NulSID(rsRecord!noWO) & "' and noInv <> '" & NulSID(rsRecord!noInv) & "' and noclaim = '" & NulSID(Text(1)) & "' and nowcar = '" & NulSID(Text(2)) & "'"
21                Set rsHlp = objO.FetchData(sSQL)
22                If Not rsHlp.EOF Then
23                    MsgBox "No Claim = " & NulSID(Text(1)) & vbCrLf & "No WCAR = " & NulSID(Text(2)) & vbCrLf & "Telah digunakan NoWo = " & NulSID(rsHlp!no_wo) & " , Noinv = " & NulSID(rsHlp!noInv) & " , Cabang = " & NulSID(rsHlp!namacabang), vbInformation
24                    Exit Sub
25                End If
26            End If
27            Call objHT.OpenTransactionBengkel
28            isBegin = True
              
29            sSQL = "UPDATE claimfop SET noclaim = '" & NulSID(Text(1)) & "', nowcar = '" & NulSID(Text(2)) & "' WHERE no_wo = '" & NulSID(rsRecord!noWO) & "' and noInv = '" & NulSID(rsRecord!noInv) & "' and kodebengkel = '" & Log.KodeDealer & "'"
30            Call objHT.InsertData_Bengkel(sSQL)
              
31            sSQL = "SELECT * FROM [" & dbReplikasiHO & "]..claimfop WITH (NOLOCK) WHERE status <> 'BTL' AND no_wo = '" & NulSID(rsRecord!noWO) & "' and noInv = '" & NulSID(rsRecord!noInv) & "'"
32            Set rsHlp = objHT.FetchDataBengkel(sSQL)
33            If Not rsHlp.EOF Then
34                sSQL = "UPDATE [" & dbReplikasiHO & "]..claimfop SET noclaim = '" & NulSID(Text(1)) & "', nowcar = '" & NulSID(Text(2)) & "' WHERE no_wo = '" & NulSID(rsRecord!noWO) & "' and noInv = '" & NulSID(rsRecord!noInv) & "'  and kodebengkel = '" & Log.KodeDealer & "'"
35            Else
36                sSQL = "INSERT INTO [" & dbReplikasiHO & "]..[ClaimFOP]([no_wo],[noinv],[noclaim],[nowcar],[status],[transfer],[tgltransfer],[kodebengkel],[namacabang])"
37                sSQL = sSQL & " VALUES "
38                sSQL = sSQL & " ('" & NulSID(rsRecord!noWO) & "','" & NulSID(rsRecord!noInv) & "','" & Trim(Text(1)) & "','" & Trim(Text(2)) & "','OK','1','" & Format(gTanggal, "MM/DD/YYYY") & "','" & Log.KodeDealer & "','" & Log.LogNamaNas & "' )"
39            End If
40            Call objHT.InsertData_Bengkel(sSQL)
              Call objHT.CloseTransactionBengkel
41            ''MsgBox "Data berhasil disimpan", vbInformation
42            Call hapusForm
43            Call refreshGrid(IIf(chAll.value = Checked, True, False))
44            MsgBox "Data berhasil disimpan", vbInformation
45            frmeKonfirm.Visible = False
46        Else
47            MsgBox pesan
48        End If

49       On Error GoTo 0
50       Exit Sub

cmdSimpan_Click_Error:
51        If isBegin Then
52            Call objHT.CloseTransactionBengkel_err
53        End If
54        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpan_Click of Form frmInputWCARTAM"


End Sub

Private Sub cetakReport()
Dim Report As New RptInfoTagihCuci
Dim rsRepo As ADODB.Recordset
Dim rsHlp As ADODB.Recordset

    sSQL = "select distinct t.noorder,j.no_polisi,q.namastall,j.date,t.tanggal,tc.* from tempprogresscuci t inner join jobdta j on j.no_wo = t.noorder inner join qty_stallwashing q on q.nostall = t.nostall inner join TagihCuci tc on tc.nowo = t.noorder " _
         & "WHERE T.FLAG = 'F' AND NOTAGIH = '" & Trim(Text(1)) & "' AND TglTagih BETWEEN '" & Format(isnow, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(isnow, "MM/DD/YYYY 23:59:59") & "' ORDER BY NOWO "
    Set rsRepo = objO.FetchData(sSQL)
    Report.database.SetDataSource rsRepo
    
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    'Report.txtKota.SetText gKotaCbg & "," & TglDDMMMMYYYY(Format(isnow, "MM/DD/YYYY"))
    Report.txtTotal.SetText rsRepo.RecordCount
    
    Set rsHlp = objO.FetchData("SELECT FS_NM_PETUGAS FROM T_PETUGAS WITH(NOLOCK) WHERE FS_KD_PETUGAS = '" & Trim(gSA) & "'")
    If Not rsHlp.EOF Then
        Report.txtPetugas.SetText Trim(rsHlp("FS_NM_PETUGAS"))
    End If
    Report.txtPenagih.SetText Trim(Text(2))
    Report.txtKabeng.SetText Trim(gSecmanager)

    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Visible = True
        
        .Top = 0
        .Left = gridTagih.Left
        .Width = gridTagih.Width
        .Height = cmdKonfir.Top + cmdKonfir.Height
        
        cmdTutup.Visible = True
        cmdTutup.Left = .Left
        cmdTutup.Top = .Top + 520
        
        .Zoom 100
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub hapusForm()
    For i = 1 To 2
        Text(i) = ""
    Next
    Tgl1 = Format(isnow, "DD/MM/YYYY")
    Tgl2 = Format(isnow, "DD/MM/YYYY")
    nTagih = 0
    totTagih = 0
End Sub

Private Sub createRS()
    If rsRecord.State Then rsRecord.Close
    Set rsRecord = Nothing
    With rsRecord
        'If .State <> adStateClosed Then .Close
        .Fields.Append "TglWO", adDate
        .Fields.Append "TglInv", adDate
        .Fields.Append "Sa", adVarChar, 40
        .Fields.Append "NoWO", adVarChar, 40
        .Fields.Append "NoInv", adVarChar, 40
        .Fields.Append "Customer", adVarChar, 200
        .Fields.Append "NoPolisi", adVarChar, 40
        .Fields.Append "NoRangka", adVarChar, 40
        .Fields.Append "Model", adVarChar, 100
        .Fields.Append "NoClaim", adVarChar, 40
        .Fields.Append "NoWCAR", adVarChar, 40
        .Open
    End With
    Set gridTagih.DataSource = rsRecord
    Call headColumns
End Sub

Private Sub headColumns()
    
    With gridTagih
        .Columns(0).Width = 1500
        .Columns(0).Caption = "Tgl WO"
        .Columns(0).DataField = "TglWO"
        .Splits(0).Columns(0).Locked = True
        .Columns(1).Width = 1500
        .Columns(1).Caption = "Tgl Inv"
        .Columns(1).DataField = "TglInv"
        .Splits(0).Columns(1).Locked = True
        .Columns(2).Width = 800
        .Columns(2).Caption = "SA"
        .Columns(2).DataField = "Sa"
        .Splits(0).Columns(2).Locked = True
        .Columns(3).Width = 1500
        .Columns(3).Caption = "No WO"
        .Columns(3).DataField = "NoWO"
        .Splits(0).Columns(3).Locked = True
        .Columns(4).Width = 1500
        .Columns(4).Caption = "No Inv"
        .Columns(4).DataField = "NoInv"
        .Splits(0).Columns(4).Locked = False
        .Columns(5).Width = 3000
        .Columns(5).Caption = "Customer"
        .Columns(5).DataField = "Customer"
        .Splits(0).Columns(5).Locked = True
        .Columns(6).Width = 1500
        .Columns(6).Caption = "No Polisi"
        .Columns(6).DataField = "NoPolisi"
        .Splits(0).Columns(6).Locked = True
        .Columns(7).Width = 1500
        .Columns(7).Caption = "No Rangka"
        .Columns(7).DataField = "NoRangka"
        .Splits(0).Columns(7).Locked = True
        .Columns(8).Width = 3000
        .Columns(8).Caption = "Unit"
        .Columns(8).DataField = "Model"
        .Splits(0).Columns(8).Locked = True
        .Columns(9).Width = 1500
        .Columns(9).Caption = "No Claim"
        .Columns(9).DataField = "NoClaim"
        .Splits(0).Columns(9).Locked = True
        .Columns(10).Width = 1500
        .Columns(10).Caption = "NoWCAR"
        .Columns(10).DataField = "NoWCAR"
        .Splits(0).Columns(10).Locked = True
    End With
End Sub

Private Sub refreshGrid(belumInput As Boolean)
Dim rsTagih As ADODB.Recordset
Dim varTagih As String
Dim kriteria As String

    kriteria = ""
    If (IIf(IsNull(Tgl1), "", Tgl1) <> "") Then
        kriteria = kriteria & " AND b.[DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' "
    End If
    If (IIf(IsNull(Tgl3), "", Tgl3) <> "") Then
        kriteria = kriteria & " AND b.[invdate] BETWEEN '" & Format(Tgl3, "MM/DD/YYYY") & "' AND '" & Format(Tgl4, "MM/DD/YYYY") & "' "
    End If
    
    If belumInput Then
        varTagih = " and (a.noclaim = '' or a.nowcar = '')"
    End If
    sSQL = "select b.DATE tglwo,b.INVDATE tglinv,b.FRONT,b.NO_WO,b.NOINV,c.CUSTOMER,b.NO_POLISI ,b.NO_CHASSIS,b.FMDL,a.noclaim,a.nowcar"
    sSQL = sSQL & " from ClaimFOP a with (nolock)"
    sSQL = sSQL & " inner join T_INVOICE b with (nolock) on a.no_wo = b.NO_WO and b.NOINV = a.noinv and b.FS_LOG_DEL = ''"
    sSQL = sSQL & " left join JOBDTA c with (nolock) on a.no_wo = c.NO_WO and b.NO_CHASSIS = c.NO_CHASSIS and c.FS_LOG_DEL = ''"
    sSQL = sSQL & " WHERE a.[status] <> 'btl' " & kriteria & varTagih
    Set rsTagih = objO.FetchData(sSQL)
    Call createRS
    If rsTagih.RecordCount > 0 Then
        rsTagih.MoveFirst
        Do While Not rsTagih.EOF
            With rsRecord
                .AddNew
                !tglWO = Format(Trim(rsTagih("tglwo")), "DD/MM/YYYY")
                !tglInv = Format(Trim(rsTagih("tglinv")), "DD/MM/YYYY")
                !sa = Trim(rsTagih("FRONT"))
                !noWO = Trim(rsTagih("NO_WO"))
                !noInv = Trim(rsTagih("NOINV"))
                !customer = Trim(rsTagih("CUSTOMER"))
                !NoPolisi = Trim(rsTagih("NO_POLISI"))
                !noRangka = Trim(rsTagih("NO_CHASSIS"))
                !model = Trim(rsTagih("FMDL"))
                !noclaim = Trim(rsTagih("noclaim"))
                !nowcar = Trim(rsTagih("nowcar"))
                .Update
            End With
            rsTagih.MoveNext
        Loop
        gridTagih.DataSource = rsRecord
        nTagih.Caption = rsRecord.RecordCount
        gridTagih.FetchRowStyle = True
       ' Call headColumns
        rsRecord.MoveFirst
    Else
        Call createRS
    End If
    
    
    'nTagih = 0
End Sub

Private Sub gridSublet_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsRecord.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    If Nul(rs.Fields(14).value) = True Then
    RowStyle = "SudahDitagih"
    End If
End Sub

Private Sub gridSublet_FilterChange()
    On Error GoTo errHandler
    Set cols = gridSublet.Columns
    Dim C As Integer
    
    C = gridSublet.col
    gridSublet.HoldFields
    rsRecord.filter = getFilter()
    If rsRecord.EOF Then
        For Each col In gridSublet.Columns
            col.FilterText = ""
        Next col
        rsRecord.filter = ""
    End If
    gridSublet.col = C
    gridSublet.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        
        End If
    Next col
    getFilter = tmp
End Function

Private Sub gridSublet_HeadClick(ByVal ColIndex As Integer)
On Error GoTo err
   Urut = IIf(Urut = "DESC", "ASC", "DESC")
   If ColIndex <= 12 Then rsRecord.Sort = gridSublet.Columns(ColIndex).DataField & " " & Urut
   gridSublet.SetFocus
   Sendkeys "{Right " & Val(ColIndex) & "}"
Exit Sub
err:
   MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
End Sub

Private Sub gridTagih_AfterColUpdate(ByVal ColIndex As Integer)
    With gridTagih
        Select Case ColIndex
            Case 6
                If .Columns(ColIndex) Then
                    jmlTagih = jmlTagih + 1
                Else
                    jmlTagih = jmlTagih - 1
                End If
                If chAll.value = True Then
                    MsgBox "Cekbox tampilkan semua tidak boleh dicetang !", vbInformation + vbOKOnly, "Sublet"
                End If
        End Select
        
        'nTagih.Caption = jmlTagih
    End With
End Sub

Private Sub gridTagih_DblClick()
    Text(0) = NulSID(rsRecord!noWO)
    Text(3) = NulSID(rsRecord!noInv)
    Text(4) = NulSID(rsRecord!NoPolisi)
    Text(5) = NulSID(rsRecord!noRangka)
    Text(7) = Left(NulSID(rsRecord!customer), 30)
    Text(1) = NulSID(rsRecord!noclaim)
    Text(2) = NulSID(rsRecord!nowcar)
    frmeKonfirm.Visible = True
End Sub

Private Sub gridTagih_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsRecord.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    
    If Trim(rs.Fields(7).value) = True Then
    RowStyle = "SudahDitagih"
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    Dim n As String
    
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            n = Text(Index).Mask
            Text(Index).Mask = ""
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            Text(Index).Mask = n
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
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "SUPPLIER"
                frmCariAcess.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsRepo As ADODB.Recordset
    If Index = 0 Then
        If Text(Index) <> "" Then
            Set rsRepo = objO.FetchDataACC("SELECT KODELGN, NAMALGN FROM APMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(Text(Index)) & "'")
            If rsRepo.EOF Then
                gMainKeyCari = Text(Index)
                gSearchOption = "SUPPLIER"
                frmCariAcess.show (vbModal)
            Else
                Text(3) = Trim(rsRepo("NAMALGN"))
            End If
            If rsRepo.State <> adStateClosed Then rsRepo.Close
            Call refreshGrid(IIf(chAll.value = Checked, True, False))
        End If
    End If
    If Index = 1 Then
        If Text(1) <> "" Then
            If objO.FetchData_Boolean("Select No_Tagihan FROM KONFIRM_SUBLET WITH(NOLOCK) WHERE NO_Tagihan = '" & Text(1) & "'") = True Then
                MsgBox "Nomor Tagihan : " & Text(1) & ", Sudah terdaftar !"
                Text(1) = ""
            End If
        End If
    End If
End Sub



