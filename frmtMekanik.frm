VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtMekanik 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Mekanik"
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13890
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmtMekanik.frx":0000
   ScaleHeight     =   9945
   ScaleWidth      =   13890
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   8760
      MaxLength       =   100
      TabIndex        =   41
      Top             =   7440
      Width           =   4800
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   2
      Left            =   8760
      MaxLength       =   9
      TabIndex        =   37
      Top             =   6720
      Width           =   1305
   End
   Begin VB.ComboBox cboLevel 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   7800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cboTipe 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8760
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   7440
      Width           =   2415
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   0
      Left            =   8760
      MaxLength       =   5
      TabIndex        =   39
      Top             =   7080
      Width           =   1065
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   9840
      MaxLength       =   25
      TabIndex        =   16
      Top             =   7080
      Width           =   3720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3780
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   4785
      Begin VB.CommandButton asxPowerButton1 
         BackColor       =   &H0080FF80&
         Caption         =   "Mekanik Stall"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton asxPowerButton2 
         BackColor       =   &H0080FF80&
         Caption         =   "Atur Stall Cuci"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdKerja 
         BackColor       =   &H0080FF80&
         Caption         =   "Absensi Mekanik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2640
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   1110
         TabIndex        =   0
         Top             =   225
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   1110
         TabIndex        =   1
         Top             =   525
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   3
         Left            =   1110
         TabIndex        =   2
         Top             =   825
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   1110
         TabIndex        =   3
         Top             =   1125
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   4
         Left            =   1110
         TabIndex        =   9
         Top             =   1440
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin AIFCmp1.asxPowerButton cmdAturStall 
         Height          =   480
         Left            =   0
         TabIndex        =   10
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   847
         BackColor       =   14482683
         BorderStyle     =   3
         Caption         =   " Jam &Stall"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotTracking     =   -1  'True
         Appearance      =   1
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   1080
         TabIndex        =   30
         Top             =   1920
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
         MICON           =   "frmtMekanik.frx":2290
         PICN            =   "frmtMekanik.frx":22AC
         PICH            =   "frmtMekanik.frx":27BF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapus 
         Height          =   495
         Left            =   2280
         TabIndex        =   31
         Top             =   1920
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&d"
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
         MICON           =   "frmtMekanik.frx":2D61
         PICN            =   "frmtMekanik.frx":2D7D
         PICH            =   "frmtMekanik.frx":3305
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKerja1 
         Height          =   495
         Left            =   1080
         TabIndex        =   33
         Top             =   3360
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&M"
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
         MICON           =   "frmtMekanik.frx":3941
         PICN            =   "frmtMekanik.frx":395D
         PICH            =   "frmtMekanik.frx":4176
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000080FF&
         X1              =   -1080
         X2              =   3600
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   4680
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   1080
         X2              =   3480
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
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
         Left            =   660
         TabIndex        =   8
         Top             =   1170
         Width           =   375
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   1680
         Top             =   1125
         Width           =   60
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Index           =   0
         Left            =   1680
         Top             =   825
         Width           =   60
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Left            =   615
         TabIndex        =   7
         Top             =   870
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         TabIndex        =   6
         Top             =   570
         Width           =   405
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   1680
         Top             =   225
         Width           =   60
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Mekanik"
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
         Left            =   75
         TabIndex        =   5
         Top             =   270
         Width           =   975
      End
   End
   Begin ucXPButton.XPButton cmdExit 
      Height          =   495
      Left            =   12480
      TabIndex        =   12
      Top             =   9120
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
      MICON           =   "frmtMekanik.frx":4B26
      PICN            =   "frmtMekanik.frx":4B42
      PICH            =   "frmtMekanik.frx":5016
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   5895
      Left            =   5565
      TabIndex        =   14
      Top             =   630
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   10398
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode"
      Columns(0).DataField=   "Kode"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nama"
      Columns(1).DataField=   "Nama"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Tipe"
      Columns(2).DataField=   "Tipe"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "NIK"
      Columns(3).DataField=   "NIK"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2011"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=4419"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4313"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2990"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2884"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(2).ButtonHeader=1"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=3519"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=3413"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(3).ButtonHeader=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H80000016&,.bold=-1"
      _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin ucXPButton.XPButton cmdCreate 
      Height          =   495
      Left            =   9720
      TabIndex        =   22
      Top             =   8280
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&c"
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
      MICON           =   "frmtMekanik.frx":5561
      PICN            =   "frmtMekanik.frx":557D
      PICH            =   "frmtMekanik.frx":5AB2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdCancel 
      Height          =   495
      Left            =   10920
      TabIndex        =   23
      Top             =   8280
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&c"
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
      MICON           =   "frmtMekanik.frx":60A5
      PICN            =   "frmtMekanik.frx":60C1
      PICH            =   "frmtMekanik.frx":6605
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdAdd 
      Height          =   495
      Left            =   7320
      TabIndex        =   24
      Top             =   8280
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&a"
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
      MICON           =   "frmtMekanik.frx":6BFC
      PICN            =   "frmtMekanik.frx":6C18
      PICH            =   "frmtMekanik.frx":7064
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdEdit 
      Height          =   495
      Left            =   8520
      TabIndex        =   25
      Top             =   8280
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&e"
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
      MICON           =   "frmtMekanik.frx":752C
      PICN            =   "frmtMekanik.frx":7548
      PICH            =   "frmtMekanik.frx":79F3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdDelete 
      Height          =   495
      Left            =   10920
      TabIndex        =   26
      Top             =   8280
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&d"
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
      MICON           =   "frmtMekanik.frx":7F18
      PICN            =   "frmtMekanik.frx":7F34
      PICH            =   "frmtMekanik.frx":84BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdSave 
      Height          =   495
      Left            =   9720
      TabIndex        =   27
      Top             =   8280
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
      MICON           =   "frmtMekanik.frx":8AF8
      PICN            =   "frmtMekanik.frx":8B14
      PICH            =   "frmtMekanik.frx":9027
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   270
      Left            =   10080
      Top             =   6720
      Width           =   60
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   8640
      TabIndex        =   40
      Top             =   6720
      Width           =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIK"
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
      Left            =   7320
      TabIndex        =   38
      Top             =   6720
      Width           =   345
   End
   Begin VB.Label lblJml 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12360
      TabIndex        =   32
      Top             =   180
      Width           =   1380
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000080FF&
      X1              =   5640
      X2              =   13680
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   8640
      TabIndex        =   29
      Top             =   7800
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level MP"
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
      Left            =   7320
      TabIndex        =   28
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000080FF&
      X1              =   5640
      X2              =   13680
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipe MP"
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
      Left            =   7320
      TabIndex        =   21
      Top             =   7440
      Width           =   750
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   8640
      TabIndex        =   20
      Top             =   7440
      Width           =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode - Nama"
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
      Left            =   7320
      TabIndex        =   19
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   8640
      TabIndex        =   15
      Top             =   7080
      Width           =   60
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      X1              =   13800
      X2              =   480
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      X1              =   13800
      X2              =   480
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  MAN POWER - NON MEKANIK  ::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   7680
      TabIndex        =   13
      Top             =   180
      Width           =   3765
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  MAN POWER - MEKANIK  ::."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   180
      Width           =   3225
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000080FF&
      Height          =   8295
      Left            =   495
      Top             =   600
      Width           =   4905
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8295
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000080FF&
      Height          =   8295
      Left            =   5520
      Top             =   600
      Width           =   8250
   End
End
Attribute VB_Name = "frmtMekanik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsM As ADODB.Recordset
Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String
Private objMk As New clsManPower
Private C As Integer


Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim i As Integer
    Dim xHasil As String
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC"
                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
            Case Else
                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
            End Select
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC"
                xHasil = xHasil + gsSql2(i)
            Case Else
                xHasil = xHasil + "'" + gsSql2(i) + "'"
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql = xHasil
End Function

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub


Private Sub asxPowerButton1_Click()
    frmAturMechStall.show vbModal
End Sub

Private Sub asxPowerButton2_Click()
    frmAturStallCuci.show vbModal
End Sub

Private Sub cmdAturStall_Click()
    Me.Tag = "Stall"
    frmAturStall.show vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    If MsgBox("Ingin menghapus data ini ?", vbYesNo, "Konfirmasi") = vbYes Then
        Call objO.DeleteData("delete mekanik where mech='" & Text(0) & "'")
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub cmdKerja_Click()
    frmAturMech2.show vbModal
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As ADODB.Recordset

    If Text(3) <> "" And Text(0) <> "" Then
        Set rsSearch = objO.FetchData("Select * from MEKANIK WITH(NOLOCK) where MECH='" & Text(0) & "'")
        Call ClearsSql
        gsSql1(0) = "MECH"
        gsSql2(0) = Text(0)
        gsSql1(1) = "[NAME]"
        gsSql2(1) = Text(2)
        gsSql1(2) = "[GROUP]"
        gsSql2(2) = Text(3)
        gsSql1(3) = "LEVMECH"
        gsSql2(3) = Text(1)
        
        If Not rsSearch.BOF Then
            sSQL = GensSql("MEKANIK", True) & " where MECH ='" & Text(0) & "'"
        Else
            sSQL = GensSql("MEKANIK", False)
        End If
        Call objO.UpdateData(sSQL)
        
        Call hapusForm
        Text(0).SetFocus
    Else
        MsgBox "Masukkan kode Mekanik/Group !"
    End If
End Sub

Private Sub Form_Load()
    Call hapusForm
    Call SetControlForm
    Call FillGridModul
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
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
            Case 0, 3
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ALL"
                gBantuan4 = ""
                frmCari.show (vbModal)
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "LEVMEKANIK"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
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
Dim rsSearch As ADODB.Recordset
    Select Case Index
        Case 0
            Set rsSearch = objO.FetchData("Select * from MEKANIK WITH(NOLOCK) where MECH='" & Text(0) & "'")
            If Not rsSearch.BOF Then
                Text(2) = rsSearch("NAME")
                Text(3) = rsSearch("GROUP")
                Text(1) = rsSearch("LEVMECH")
            End If
        Case 1
            Set rsSearch = objO.FetchData("SELECT * FROM T_MechLev WITH(NOLOCK) WHERE LEVELMECH = '" & Text(Index) & "'")
            If rsSearch.EOF Then
                gMainKeyCari = Text(Index)
                gSearchOption = "LEVMEKANIK"
                frmCari.show (vbModal)
            Else
                Text(4) = rsSearch("KETERANGAN")
            End If
    End Select
End Sub

Private Sub hapusForm()
    Text(0) = ""
    Text(2) = ""
    Text(3) = ""
    Text(1) = ""
    Text(4) = ""
    If Log.isEmStall = 0 Then
    ''If Log.dmsNonDss = 0 Then
        asxPowerButton1.Visible = False
    Else
        asxPowerButton1.Visible = True
    End If
    ''If Log.dmsNonDss = 0 Then
    If Log.isWashing = 0 Then
        asxPowerButton2.Visible = False
    Else
        asxPowerButton2.Visible = True
    End If
End Sub

Private Sub SetControlForm()
    With Me
        .cmdAdd.Enabled = True
        .cmdCreate.Visible = False
        .cmdEdit.Enabled = True
        .cmdSave.Visible = True
        .cmdSave.Enabled = False
        .cmdCancel.Visible = False
        .cmdDelete.Visible = True
        .cmdDelete.Enabled = True
        .Grid.Enabled = True
        .txt(0).BackColor = fColor
        .txt(1).BackColor = fColor
        .txt(2).BackColor = fColor
        .cboTipe.BackColor = fColor
        .cboLevel.BackColor = fColor
        .cmdExit.Enabled = True
        
        .cboTipe.Clear
        .cboTipe.AddItem "Service Advisor"
        .cboTipe.AddItem "Foreman"
        .cboTipe.AddItem "Team 21"
        .cboTipe.AddItem "Partman"
        .cboTipe.AddItem "Instructur"
        .cboTipe.AddItem "Technical Leader"
        .cboTipe.AddItem "Personel Bengkel Lainnya"

        .cboLevel.Clear
        .cboTipe.AddItem "-"
    End With
End Sub

Private Sub FillGridModul()
    Set rsM = objMk.FetchListManPower
    If rsM.RecordCount = 0 Then
        With Me
            .cmdAdd.Enabled = True
            .cmdCreate.Visible = False
            .cmdEdit.Enabled = False
            .cmdSave.Visible = True
            .cmdSave.Enabled = False
            .cmdCancel.Visible = False
            .cmdDelete.Visible = False
            .cmdDelete.Enabled = True
            .Grid.Enabled = False
            .Grid.DataSource = Nothing
            .lblJml.Caption = "0"
            .txt(0).BackColor = fColor
            .txt(1).BackColor = fColor
            .txt(2).BackColor = fColor
            .cboTipe.BackColor = fColor
            .cboLevel.BackColor = fColor
            .txt(0).Text = ""
            .txt(1).Text = ""
            .cmdExit.Enabled = True
        End With
        Exit Sub
    End If
    Set Grid.DataSource = rsM
    lblJml.Caption = rsM.RecordCount
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then Call CmdDelete_Click
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call BindToControl(rsM)
End Sub
Private Sub BindToControl(rs As ADODB.Recordset)
On Error GoTo ErrNull
    txt(0).Text = rs.Fields(0).value
    txt(1).Text = rs.Fields(1).value
    txt(2).Text = rs.Fields(4).value
    'cboTipe.Text = rs.Fields(2).value
    txt(3).Text = rs.Fields(2).value
    cboLevel.Clear
    Select Case rs.Fields(2).value
        Case "Service Advisor"
            cboLevel.AddItem "AppSA"
            cboLevel.AddItem "TSA"
            cboLevel.AddItem "PSA"
            cboLevel.AddItem "MSA"
        Case "Foreman"
            cboLevel.AddItem "TFT"
            cboLevel.AddItem "Non TFT"
        Case "Team 21"
            cboLevel.AddItem "T21 PT"
            cboLevel.AddItem "T21 DT"
            cboLevel.AddItem "T21 DMT"
        Case "Partman"
            cboLevel.AddItem "App"
            cboLevel.AddItem "Lv.1"
            cboLevel.AddItem "Lv.2"
        Case Else
            cboLevel.AddItem "Other"
    End Select
    cboLevel.Text = rs.Fields(3).value
ErrNull:
End Sub


'===============================================================================================
'========= A D D ======== A D D ======== A D D =================================================
'===============================================================================================
Private Sub CmdAdd_Click()
    With Me
        .cmdAdd.Enabled = False
        .cmdCreate.Visible = True
        .cmdCreate.Enabled = True
        .cmdEdit.Enabled = False
        .cmdSave.Visible = False
        .cmdCancel.Visible = True
        .cmdCancel.Enabled = True
        .cmdDelete.Visible = False
        .Grid.Enabled = False
        .txt(0).BackColor = tColor
       ' .txt(1).BackColor = tColor
        .txt(2).BackColor = tColor
        .cboTipe.BackColor = tColor
        .cboLevel.BackColor = tColor
        .txt(0).Text = ""
        .txt(1).Text = ""
        .txt(2).Text = ""
        .cmdExit.Enabled = False
        .txt(2).SetFocus
    End With
End Sub

Private Sub cmdCreate_Click()
    If FillNotNull = True Then
        rsM.Find "Kode = '" & Trim(txt(0).Text) & "'", , adSearchForward, 1
        If Not rsM.EOF Then
            MsgBox "Data Personel " & txt(0).Text & " sudah terdaftar !", vbInformation + vbOKOnly, "Man Power"
            Exit Sub
        End If
        rsM.Find "nik = '" & Trim(txt(2).Text) & "'", , adSearchForward, 1
        If Not rsM.EOF Then
            MsgBox "NIK " & txt(2).Text & " sudah terdaftar !", vbInformation + vbOKOnly, "Man Power"
            Exit Sub
        End If
        If MsgBox("Simpan Data Personel Baru ?", vbQuestion + vbYesNo, "Man Power") = vbYes Then
            Call objMk.AddNew_ManPower(txt(0).Text, txt(1).Text, NulSID(txt(3)), cboLevel.Text, NulSID(txt(2).Text))
            Call SetControlForm
            Call FillGridModul
        End If
    Else
        Call NotValid_Input
    End If
End Sub

'===============================================================================================
'======= E D I T =========== E D I T =========== E D I T =======================================
'===============================================================================================
Private Sub cmdEdit_Click()
    With Me
        .cmdAdd.Enabled = False
        .cmdCreate.Visible = False
        .cmdEdit.Enabled = False
        .cmdSave.Visible = True
        .cmdSave.Enabled = True
        .cmdCancel.Visible = True
        .cmdCancel.Enabled = True
        .cmdDelete.Visible = False
        .Grid.Enabled = False
        .txt(0).BackColor = fColor
        .txt(2).BackColor = tColor
        .txt(1).BackColor = fColor
        .txt(3).BackColor = fColor
        .cboTipe.BackColor = tColor
        .cboLevel.BackColor = tColor
        .cmdExit.Enabled = False
        .txt(1).SetFocus
    End With
End Sub

Private Sub cmdSave_Click()
    If FillNotNull = True Then
        If MsgBox("Update Data Personel '" & txt(0).Text & "' ?", vbQuestion + vbYesNo, "Man Power") = vbYes Then
            Call objMk.Update_ManPower(txt(0).Text, txt(1).Text, txt(3).Text, cboLevel.Text, NulSID(txt(2).Text))
            Call SetControlForm
            Call FillGridModul
        End If
    Else
        Call NotValid_Input
    End If
End Sub

'===============================================================================================
'===============================================================================================
Private Function FillNotNull() As Boolean
    FillNotNull = True
    If txt(0).Text = "" Or txt(1).Text = "" Or NulSID(txt(3).Text) = "" Or NulSID(txt(2).Text) = "" Then
        FillNotNull = False
    End If
End Function
Private Sub NotValid_Input()
    MsgBox "Data Personel Yang Anda Inputkan Belum Lengkap !", vbInformation + vbOKOnly, "Man Power"
End Sub
'===============================================================================================
'===============================================================================================

'===============================================================================================
'========= C A N C E L ========= C A N C E L ========= C A N C E L =============================
'===============================================================================================
Private Sub cmdCancel_Click()
    Call SetControlForm
    Call FillGridModul
End Sub

'===============================================================================================
'========= D E L E T E ========= D E L E T E ========= D E L E T E =============================
'===============================================================================================
Private Sub CmdDelete_Click()
    If MsgBox("Hapus Data Personel '" & txt(0).Text & "' ?", vbQuestion + vbYesNo, "Man Power") = vbYes Then
        Call objMk.Delete_ItemManPower(txt(0).Text)
        Call SetControlForm
        Call FillGridModul
    End If
End Sub


Private Sub cboTipe_Change()
    cboLevel.Clear
    Select Case cboTipe.Text
        Case "Service Advisor"
            cboLevel.AddItem "AppSA"
            cboLevel.AddItem "TSA"
            cboLevel.AddItem "PSA"
            cboLevel.AddItem "MSA"
        Case "Foreman"
            cboLevel.AddItem "TFT"
            cboLevel.AddItem "Non TFT"
        Case "Team 21"
            cboLevel.AddItem "T21 PT"
            cboLevel.AddItem "T21 DT"
            cboLevel.AddItem "T21 DMT"
        Case "Partman"
            cboLevel.AddItem "App"
            cboLevel.AddItem "Lv.1"
            cboLevel.AddItem "Lv.2"
        Case "Technical Leader"
            cboLevel.AddItem "Other"
        Case "Instructur"
            cboLevel.AddItem "Other"
        Case "Personel Bengkel Lainnya"
            cboLevel.AddItem "Other"
    End Select
End Sub

Private Sub cboTipe_Click()
    Call cboTipe_Change
End Sub


Private Sub grid_FilterChange()
On Error GoTo err
   Set cols = Grid.Columns
   
   C = Grid.col
   Grid.HoldFields
   rsM.filter = getFilter()
   If rsM.EOF = True Then
      For Each col In Grid.Columns
         col.FilterText = ""
      Next col
      rsM.filter = ""
   End If
   Grid.col = C
   Grid.EditActive = True
   lblJml.Caption = rsM.RecordCount
   Exit Sub
err:
   Resume Next
End Sub

Private Function getFilter() As String
   On Error GoTo err

On Error GoTo err
Dim tmp As String
Dim n As Integer

   For Each col In cols
      If Trim(col.FilterText) <> "" Then
         n = n + 1
         If n > 1 Then
            tmp = tmp & " AND "
         End If
         tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub grid_HeadClick(ByVal ColIndex As Integer)
On Error GoTo err
   Urut = IIf(Urut = "DESC", "ASC", "DESC")
   If ColIndex <= 3 Then rsM.Sort = "[" & Grid.Columns(ColIndex).DataField & "]" & " " & Urut
   Grid.SetFocus
   Sendkeys "{Right " & Val(ColIndex) & "}"
   lblJml.Caption = rsM.RecordCount
Exit Sub
err:
   MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter ")
End Sub

Private Sub txt_GotFocus(Index As Integer)
    If txt(Index).BackColor = fColor Then Sendkeys ("{TAB}")
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            txt(Index) = gMainKeyCari
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

Private Sub cboTipe_GotFocus()
    If cboTipe.BackColor = fColor Then Sendkeys ("{TAB}")
End Sub

Private Sub cboLevel_GotFocus()
    If cboLevel.BackColor = fColor Then Sendkeys ("{TAB}")
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 2
                gMainKeyCari = Text(Index)
                gSearchOption = "DATAKARYAWAN"
                frmCariHO.show (vbModal)
        End Select
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Sendkeys ("{TAB}")
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Dim rsSearch As ADODB.Recordset
    Select Case Index
        Case 2
            Set rsSearch = objO.FetchDataHO("Select NIK, Nama, Posisi from fkaryawan.dbo.datakaryawan WITH(NOLOCK) where NIK='" & txt(2) & "'")
            If Not rsSearch.BOF Then
                txt(2) = rsSearch("NIK")
                txt(3) = rsSearch("Posisi")
                txt(1) = rsSearch("Nama")
            End If
         
    End Select
End Sub
