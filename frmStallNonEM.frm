VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmStallNonEM 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stall & Mekanik"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13425
   Begin VB.Frame frmStal 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   5280
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdStalBatal 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Batal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton cmdStalOk 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Width           =   1815
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Top             =   1200
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648384
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   3
         Left            =   1800
         TabIndex        =   9
         Top             =   1560
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648384
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   "99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   23
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648384
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CC-CCCC-CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   11
         Left            =   3360
         TabIndex        =   28
         Top             =   1560
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648384
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   "99:99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   30
         Top             =   840
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648384
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   13
         Left            =   3465
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   "99:99"
         PromptChar      =   ">"
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2760
         TabIndex        =   41
         Top             =   1005
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   3600
         X2              =   4680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   3600
         X2              =   4680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1200
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1200
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Stall"
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
         Left            =   1320
         TabIndex        =   38
         Top             =   30
         Width           =   2040
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Selesai"
         Height          =   195
         Left            =   2715
         TabIndex        =   29
         Top             =   1605
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Mulai"
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   1605
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Stall"
         Height          =   195
         Left            =   1245
         TabIndex        =   6
         Top             =   1245
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   885
         Width           =   345
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi "
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   525
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   8280
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   12855
      Begin TrueDBGrid70.TDBGrid gridStall 
         Height          =   6735
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   11880
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
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).DataField=   ""
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).DataField=   ""
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).DataField=   ""
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(18)._VlistStyle=   0
         Columns(18)._MaxComboItems=   5
         Columns(18).DataField=   ""
         Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(19)._VlistStyle=   0
         Columns(19)._MaxComboItems=   5
         Columns(19).DataField=   ""
         Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(20)._VlistStyle=   0
         Columns(20)._MaxComboItems=   5
         Columns(20).DataField=   ""
         Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(21)._VlistStyle=   0
         Columns(21)._MaxComboItems=   5
         Columns(21).DataField=   ""
         Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(22)._VlistStyle=   0
         Columns(22)._MaxComboItems=   5
         Columns(22).DataField=   ""
         Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(23)._VlistStyle=   0
         Columns(23)._MaxComboItems=   5
         Columns(23).DataField=   ""
         Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(24)._VlistStyle=   0
         Columns(24)._MaxComboItems=   5
         Columns(24).DataField=   ""
         Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(25)._VlistStyle=   0
         Columns(25)._MaxComboItems=   5
         Columns(25).DataField=   ""
         Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(26)._VlistStyle=   0
         Columns(26)._MaxComboItems=   5
         Columns(26).DataField=   ""
         Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(27)._VlistStyle=   0
         Columns(27)._MaxComboItems=   5
         Columns(27).DataField=   ""
         Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(28)._VlistStyle=   0
         Columns(28)._MaxComboItems=   5
         Columns(28).DataField=   ""
         Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(29)._VlistStyle=   0
         Columns(29)._MaxComboItems=   5
         Columns(29).DataField=   ""
         Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(30)._VlistStyle=   0
         Columns(30)._MaxComboItems=   5
         Columns(30).DataField=   ""
         Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(31)._VlistStyle=   0
         Columns(31)._MaxComboItems=   5
         Columns(31).DataField=   ""
         Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(32)._VlistStyle=   0
         Columns(32)._MaxComboItems=   5
         Columns(32).DataField=   ""
         Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(33)._VlistStyle=   0
         Columns(33)._MaxComboItems=   5
         Columns(33).DataField=   ""
         Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(34)._VlistStyle=   0
         Columns(34)._MaxComboItems=   5
         Columns(34).DataField=   ""
         Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(35)._VlistStyle=   0
         Columns(35)._MaxComboItems=   5
         Columns(35).DataField=   ""
         Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(36)._VlistStyle=   0
         Columns(36)._MaxComboItems=   5
         Columns(36).DataField=   ""
         Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(37)._VlistStyle=   0
         Columns(37)._MaxComboItems=   5
         Columns(37).DataField=   ""
         Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(38)._VlistStyle=   0
         Columns(38)._MaxComboItems=   5
         Columns(38).DataField=   ""
         Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(39)._VlistStyle=   0
         Columns(39)._MaxComboItems=   5
         Columns(39).DataField=   ""
         Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(40)._VlistStyle=   0
         Columns(40)._MaxComboItems=   5
         Columns(40).DataField=   ""
         Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(41)._VlistStyle=   0
         Columns(41)._MaxComboItems=   5
         Columns(41).DataField=   ""
         Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(42)._VlistStyle=   0
         Columns(42)._MaxComboItems=   5
         Columns(42).DataField=   ""
         Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(43)._VlistStyle=   0
         Columns(43)._MaxComboItems=   5
         Columns(43).DataField=   ""
         Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(44)._VlistStyle=   0
         Columns(44)._MaxComboItems=   5
         Columns(44).DataField=   ""
         Columns(44)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(45)._VlistStyle=   0
         Columns(45)._MaxComboItems=   5
         Columns(45).DataField=   ""
         Columns(45)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(46)._VlistStyle=   0
         Columns(46)._MaxComboItems=   5
         Columns(46).DataField=   ""
         Columns(46)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(47)._VlistStyle=   0
         Columns(47)._MaxComboItems=   5
         Columns(47).DataField=   ""
         Columns(47)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(48)._VlistStyle=   0
         Columns(48)._MaxComboItems=   5
         Columns(48).DataField=   ""
         Columns(48)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(49)._VlistStyle=   0
         Columns(49)._MaxComboItems=   5
         Columns(49).DataField=   ""
         Columns(49)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(50)._VlistStyle=   0
         Columns(50)._MaxComboItems=   5
         Columns(50).DataField=   ""
         Columns(50)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(51)._VlistStyle=   0
         Columns(51)._MaxComboItems=   5
         Columns(51).DataField=   ""
         Columns(51)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(52)._VlistStyle=   0
         Columns(52)._MaxComboItems=   5
         Columns(52).DataField=   ""
         Columns(52)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(53)._VlistStyle=   0
         Columns(53)._MaxComboItems=   5
         Columns(53).DataField=   ""
         Columns(53)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(54)._VlistStyle=   0
         Columns(54)._MaxComboItems=   5
         Columns(54).DataField=   ""
         Columns(54)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(55)._VlistStyle=   0
         Columns(55)._MaxComboItems=   5
         Columns(55).DataField=   ""
         Columns(55)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(56)._VlistStyle=   0
         Columns(56)._MaxComboItems=   5
         Columns(56).DataField=   ""
         Columns(56)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(57)._VlistStyle=   0
         Columns(57)._MaxComboItems=   5
         Columns(57).DataField=   ""
         Columns(57)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(58)._VlistStyle=   0
         Columns(58)._MaxComboItems=   5
         Columns(58).DataField=   ""
         Columns(58)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(59)._VlistStyle=   0
         Columns(59)._MaxComboItems=   5
         Columns(59).DataField=   ""
         Columns(59)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(60)._VlistStyle=   0
         Columns(60)._MaxComboItems=   5
         Columns(60).DataField=   ""
         Columns(60)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(61)._VlistStyle=   0
         Columns(61)._MaxComboItems=   5
         Columns(61).DataField=   ""
         Columns(61)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(62)._VlistStyle=   0
         Columns(62)._MaxComboItems=   5
         Columns(62).DataField=   ""
         Columns(62)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(63)._VlistStyle=   0
         Columns(63)._MaxComboItems=   5
         Columns(63).DataField=   ""
         Columns(63)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(64)._VlistStyle=   0
         Columns(64)._MaxComboItems=   5
         Columns(64).DataField=   ""
         Columns(64)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(65)._VlistStyle=   0
         Columns(65)._MaxComboItems=   5
         Columns(65).DataField=   ""
         Columns(65)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(66)._VlistStyle=   0
         Columns(66)._MaxComboItems=   5
         Columns(66).DataField=   ""
         Columns(66)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(67)._VlistStyle=   0
         Columns(67)._MaxComboItems=   5
         Columns(67).DataField=   ""
         Columns(67)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(68)._VlistStyle=   0
         Columns(68)._MaxComboItems=   5
         Columns(68).DataField=   ""
         Columns(68)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(69)._VlistStyle=   0
         Columns(69)._MaxComboItems=   5
         Columns(69).DataField=   ""
         Columns(69)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(70)._VlistStyle=   0
         Columns(70)._MaxComboItems=   5
         Columns(70).DataField=   ""
         Columns(70)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(71)._VlistStyle=   0
         Columns(71)._MaxComboItems=   5
         Columns(71).DataField=   ""
         Columns(71)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(72)._VlistStyle=   0
         Columns(72)._MaxComboItems=   5
         Columns(72).DataField=   ""
         Columns(72)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(73)._VlistStyle=   0
         Columns(73)._MaxComboItems=   5
         Columns(73).DataField=   ""
         Columns(73)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(74)._VlistStyle=   0
         Columns(74)._MaxComboItems=   5
         Columns(74).DataField=   ""
         Columns(74)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(75)._VlistStyle=   0
         Columns(75)._MaxComboItems=   5
         Columns(75).DataField=   ""
         Columns(75)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(76)._VlistStyle=   0
         Columns(76)._MaxComboItems=   5
         Columns(76).DataField=   ""
         Columns(76)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(77)._VlistStyle=   0
         Columns(77)._MaxComboItems=   5
         Columns(77).DataField=   ""
         Columns(77)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(78)._VlistStyle=   0
         Columns(78)._MaxComboItems=   5
         Columns(78).DataField=   ""
         Columns(78)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(79)._VlistStyle=   0
         Columns(79)._MaxComboItems=   5
         Columns(79).DataField=   ""
         Columns(79)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(80)._VlistStyle=   0
         Columns(80)._MaxComboItems=   5
         Columns(80).DataField=   ""
         Columns(80)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(81)._VlistStyle=   0
         Columns(81)._MaxComboItems=   5
         Columns(81).DataField=   ""
         Columns(81)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(82)._VlistStyle=   0
         Columns(82)._MaxComboItems=   5
         Columns(82).DataField=   ""
         Columns(82)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(83)._VlistStyle=   0
         Columns(83)._MaxComboItems=   5
         Columns(83).DataField=   ""
         Columns(83)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(84)._VlistStyle=   0
         Columns(84)._MaxComboItems=   5
         Columns(84).DataField=   ""
         Columns(84)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(85)._VlistStyle=   0
         Columns(85)._MaxComboItems=   5
         Columns(85).DataField=   ""
         Columns(85)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(86)._VlistStyle=   0
         Columns(86)._MaxComboItems=   5
         Columns(86).DataField=   ""
         Columns(86)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(87)._VlistStyle=   0
         Columns(87)._MaxComboItems=   5
         Columns(87).DataField=   ""
         Columns(87)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(88)._VlistStyle=   0
         Columns(88)._MaxComboItems=   5
         Columns(88).DataField=   ""
         Columns(88)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(89)._VlistStyle=   0
         Columns(89)._MaxComboItems=   5
         Columns(89).DataField=   ""
         Columns(89)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(90)._VlistStyle=   0
         Columns(90)._MaxComboItems=   5
         Columns(90).DataField=   ""
         Columns(90)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(91)._VlistStyle=   0
         Columns(91)._MaxComboItems=   5
         Columns(91).DataField=   ""
         Columns(91)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(92)._VlistStyle=   0
         Columns(92)._MaxComboItems=   5
         Columns(92).DataField=   ""
         Columns(92)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(93)._VlistStyle=   0
         Columns(93)._MaxComboItems=   5
         Columns(93).DataField=   ""
         Columns(93)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(94)._VlistStyle=   0
         Columns(94)._MaxComboItems=   5
         Columns(94).DataField=   ""
         Columns(94)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(95)._VlistStyle=   0
         Columns(95)._MaxComboItems=   5
         Columns(95).DataField=   ""
         Columns(95)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(96)._VlistStyle=   0
         Columns(96)._MaxComboItems=   5
         Columns(96).DataField=   ""
         Columns(96)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(97)._VlistStyle=   0
         Columns(97)._MaxComboItems=   5
         Columns(97).DataField=   ""
         Columns(97)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(98)._VlistStyle=   0
         Columns(98)._MaxComboItems=   5
         Columns(98).DataField=   ""
         Columns(98)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(99)._VlistStyle=   0
         Columns(99)._MaxComboItems=   5
         Columns(99).DataField=   ""
         Columns(99)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(100)._VlistStyle=   0
         Columns(100)._MaxComboItems=   5
         Columns(100).DataField=   ""
         Columns(100)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(101)._VlistStyle=   0
         Columns(101)._MaxComboItems=   5
         Columns(101).DataField=   ""
         Columns(101)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(102)._VlistStyle=   0
         Columns(102)._MaxComboItems=   5
         Columns(102).DataField=   ""
         Columns(102)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(103)._VlistStyle=   0
         Columns(103)._MaxComboItems=   5
         Columns(103).DataField=   ""
         Columns(103)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(104)._VlistStyle=   0
         Columns(104)._MaxComboItems=   5
         Columns(104).DataField=   ""
         Columns(104)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(105)._VlistStyle=   0
         Columns(105)._MaxComboItems=   5
         Columns(105).DataField=   ""
         Columns(105)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(106)._VlistStyle=   0
         Columns(106)._MaxComboItems=   5
         Columns(106).DataField=   ""
         Columns(106)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(107)._VlistStyle=   0
         Columns(107)._MaxComboItems=   5
         Columns(107).DataField=   ""
         Columns(107)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(108)._VlistStyle=   0
         Columns(108)._MaxComboItems=   5
         Columns(108).DataField=   ""
         Columns(108)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(109)._VlistStyle=   0
         Columns(109)._MaxComboItems=   5
         Columns(109).DataField=   ""
         Columns(109)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(110)._VlistStyle=   0
         Columns(110)._MaxComboItems=   5
         Columns(110).DataField=   ""
         Columns(110)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(111)._VlistStyle=   0
         Columns(111)._MaxComboItems=   5
         Columns(111).DataField=   ""
         Columns(111)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(112)._VlistStyle=   0
         Columns(112)._MaxComboItems=   5
         Columns(112).DataField=   ""
         Columns(112)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(113)._VlistStyle=   0
         Columns(113)._MaxComboItems=   5
         Columns(113).DataField=   ""
         Columns(113)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(114)._VlistStyle=   0
         Columns(114)._MaxComboItems=   5
         Columns(114).DataField=   ""
         Columns(114)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(115)._VlistStyle=   0
         Columns(115)._MaxComboItems=   5
         Columns(115).DataField=   ""
         Columns(115)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(116)._VlistStyle=   0
         Columns(116)._MaxComboItems=   5
         Columns(116).DataField=   ""
         Columns(116)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(117)._VlistStyle=   0
         Columns(117)._MaxComboItems=   5
         Columns(117).DataField=   ""
         Columns(117)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(118)._VlistStyle=   0
         Columns(118)._MaxComboItems=   5
         Columns(118).DataField=   ""
         Columns(118)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(119)._VlistStyle=   0
         Columns(119)._MaxComboItems=   5
         Columns(119).DataField=   ""
         Columns(119)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(120)._VlistStyle=   0
         Columns(120)._MaxComboItems=   5
         Columns(120).DataField=   ""
         Columns(120)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(121)._VlistStyle=   0
         Columns(121)._MaxComboItems=   5
         Columns(121).DataField=   ""
         Columns(121)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(122)._VlistStyle=   0
         Columns(122)._MaxComboItems=   5
         Columns(122).DataField=   ""
         Columns(122)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(123)._VlistStyle=   0
         Columns(123)._MaxComboItems=   5
         Columns(123).DataField=   ""
         Columns(123)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(124)._VlistStyle=   0
         Columns(124)._MaxComboItems=   5
         Columns(124).DataField=   ""
         Columns(124)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(125)._VlistStyle=   0
         Columns(125)._MaxComboItems=   5
         Columns(125).DataField=   ""
         Columns(125)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(126)._VlistStyle=   0
         Columns(126)._MaxComboItems=   5
         Columns(126).DataField=   ""
         Columns(126)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(127)._VlistStyle=   0
         Columns(127)._MaxComboItems=   5
         Columns(127).DataField=   ""
         Columns(127)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(128)._VlistStyle=   0
         Columns(128)._MaxComboItems=   5
         Columns(128).DataField=   ""
         Columns(128)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(129)._VlistStyle=   0
         Columns(129)._MaxComboItems=   5
         Columns(129).DataField=   ""
         Columns(129)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(130)._VlistStyle=   0
         Columns(130)._MaxComboItems=   5
         Columns(130).DataField=   ""
         Columns(130)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(131)._VlistStyle=   0
         Columns(131)._MaxComboItems=   5
         Columns(131).DataField=   ""
         Columns(131)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(132)._VlistStyle=   0
         Columns(132)._MaxComboItems=   5
         Columns(132).DataField=   ""
         Columns(132)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(133)._VlistStyle=   0
         Columns(133)._MaxComboItems=   5
         Columns(133).DataField=   ""
         Columns(133)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(134)._VlistStyle=   0
         Columns(134)._MaxComboItems=   5
         Columns(134).DataField=   ""
         Columns(134)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(135)._VlistStyle=   0
         Columns(135)._MaxComboItems=   5
         Columns(135).DataField=   ""
         Columns(135)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(136)._VlistStyle=   0
         Columns(136)._MaxComboItems=   5
         Columns(136).DataField=   ""
         Columns(136)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(137)._VlistStyle=   0
         Columns(137)._MaxComboItems=   5
         Columns(137).DataField=   ""
         Columns(137)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(138)._VlistStyle=   0
         Columns(138)._MaxComboItems=   5
         Columns(138).DataField=   ""
         Columns(138)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(139)._VlistStyle=   0
         Columns(139)._MaxComboItems=   5
         Columns(139).DataField=   ""
         Columns(139)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(140)._VlistStyle=   0
         Columns(140)._MaxComboItems=   5
         Columns(140).DataField=   ""
         Columns(140)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(141)._VlistStyle=   0
         Columns(141)._MaxComboItems=   5
         Columns(141).DataField=   ""
         Columns(141)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(142)._VlistStyle=   0
         Columns(142)._MaxComboItems=   5
         Columns(142).DataField=   ""
         Columns(142)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(143)._VlistStyle=   0
         Columns(143)._MaxComboItems=   5
         Columns(143).DataField=   ""
         Columns(143)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(144)._VlistStyle=   0
         Columns(144)._MaxComboItems=   5
         Columns(144).DataField=   ""
         Columns(144)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(145)._VlistStyle=   0
         Columns(145)._MaxComboItems=   5
         Columns(145).DataField=   ""
         Columns(145)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(146)._VlistStyle=   0
         Columns(146)._MaxComboItems=   5
         Columns(146).DataField=   ""
         Columns(146)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(147)._VlistStyle=   0
         Columns(147)._MaxComboItems=   5
         Columns(147).DataField=   ""
         Columns(147)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(148)._VlistStyle=   0
         Columns(148)._MaxComboItems=   5
         Columns(148).DataField=   ""
         Columns(148)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(149)._VlistStyle=   0
         Columns(149)._MaxComboItems=   5
         Columns(149).DataField=   ""
         Columns(149)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(150)._VlistStyle=   0
         Columns(150)._MaxComboItems=   5
         Columns(150).DataField=   ""
         Columns(150)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   151
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowFocus=   0   'False
         Splits(0).SizeMode=   2
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColSelect=   0   'False
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).DividerColor=   14869218
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=151"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=4366"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4366"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4233"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=4366"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=4233"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=4366"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=4233"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=4366"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=4233"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=4366"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=4233"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=4366"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=4233"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=4366"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=4233"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=4366"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=4233"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=4366"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=4233"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=4366"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=4233"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=4366"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=4233"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(53)=   "Column(13).Width=4366"
         Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=4233"
         Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(57)=   "Column(14).Width=4366"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=4233"
         Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(61)=   "Column(15).Width=4366"
         Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=4233"
         Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(65)=   "Column(16).Width=4366"
         Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=4233"
         Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(69)=   "Column(17).Width=4366"
         Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=4233"
         Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(73)=   "Column(18).Width=4366"
         Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=4233"
         Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(77)=   "Column(19).Width=4366"
         Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=4233"
         Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(81)=   "Column(20).Width=4366"
         Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=4233"
         Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(85)=   "Column(21).Width=4366"
         Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=4233"
         Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(89)=   "Column(22).Width=4366"
         Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=4233"
         Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
         Splits(0)._ColumnProps(93)=   "Column(23).Width=4366"
         Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
         Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=4233"
         Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
         Splits(0)._ColumnProps(97)=   "Column(24).Width=4366"
         Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
         Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=4233"
         Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
         Splits(0)._ColumnProps(101)=   "Column(25).Width=4366"
         Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
         Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=4233"
         Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
         Splits(0)._ColumnProps(105)=   "Column(26).Width=4366"
         Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
         Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=4233"
         Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
         Splits(0)._ColumnProps(109)=   "Column(27).Width=4366"
         Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
         Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=4233"
         Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
         Splits(0)._ColumnProps(113)=   "Column(28).Width=4366"
         Splits(0)._ColumnProps(114)=   "Column(28).DividerColor=0"
         Splits(0)._ColumnProps(115)=   "Column(28)._WidthInPix=4233"
         Splits(0)._ColumnProps(116)=   "Column(28).Order=29"
         Splits(0)._ColumnProps(117)=   "Column(29).Width=4366"
         Splits(0)._ColumnProps(118)=   "Column(29).DividerColor=0"
         Splits(0)._ColumnProps(119)=   "Column(29)._WidthInPix=4233"
         Splits(0)._ColumnProps(120)=   "Column(29).Order=30"
         Splits(0)._ColumnProps(121)=   "Column(30).Width=4366"
         Splits(0)._ColumnProps(122)=   "Column(30).DividerColor=0"
         Splits(0)._ColumnProps(123)=   "Column(30)._WidthInPix=4233"
         Splits(0)._ColumnProps(124)=   "Column(30).Order=31"
         Splits(0)._ColumnProps(125)=   "Column(31).Width=4366"
         Splits(0)._ColumnProps(126)=   "Column(31).DividerColor=0"
         Splits(0)._ColumnProps(127)=   "Column(31)._WidthInPix=4233"
         Splits(0)._ColumnProps(128)=   "Column(31).Order=32"
         Splits(0)._ColumnProps(129)=   "Column(32).Width=4366"
         Splits(0)._ColumnProps(130)=   "Column(32).DividerColor=0"
         Splits(0)._ColumnProps(131)=   "Column(32)._WidthInPix=4233"
         Splits(0)._ColumnProps(132)=   "Column(32).Order=33"
         Splits(0)._ColumnProps(133)=   "Column(33).Width=4366"
         Splits(0)._ColumnProps(134)=   "Column(33).DividerColor=0"
         Splits(0)._ColumnProps(135)=   "Column(33)._WidthInPix=4233"
         Splits(0)._ColumnProps(136)=   "Column(33).Order=34"
         Splits(0)._ColumnProps(137)=   "Column(34).Width=4366"
         Splits(0)._ColumnProps(138)=   "Column(34).DividerColor=0"
         Splits(0)._ColumnProps(139)=   "Column(34)._WidthInPix=4233"
         Splits(0)._ColumnProps(140)=   "Column(34).Order=35"
         Splits(0)._ColumnProps(141)=   "Column(35).Width=4366"
         Splits(0)._ColumnProps(142)=   "Column(35).DividerColor=0"
         Splits(0)._ColumnProps(143)=   "Column(35)._WidthInPix=4233"
         Splits(0)._ColumnProps(144)=   "Column(35).Order=36"
         Splits(0)._ColumnProps(145)=   "Column(36).Width=4366"
         Splits(0)._ColumnProps(146)=   "Column(36).DividerColor=0"
         Splits(0)._ColumnProps(147)=   "Column(36)._WidthInPix=4233"
         Splits(0)._ColumnProps(148)=   "Column(36).Order=37"
         Splits(0)._ColumnProps(149)=   "Column(37).Width=4366"
         Splits(0)._ColumnProps(150)=   "Column(37).DividerColor=0"
         Splits(0)._ColumnProps(151)=   "Column(37)._WidthInPix=4233"
         Splits(0)._ColumnProps(152)=   "Column(37).Order=38"
         Splits(0)._ColumnProps(153)=   "Column(38).Width=4366"
         Splits(0)._ColumnProps(154)=   "Column(38).DividerColor=0"
         Splits(0)._ColumnProps(155)=   "Column(38)._WidthInPix=4233"
         Splits(0)._ColumnProps(156)=   "Column(38).Order=39"
         Splits(0)._ColumnProps(157)=   "Column(39).Width=4366"
         Splits(0)._ColumnProps(158)=   "Column(39).DividerColor=0"
         Splits(0)._ColumnProps(159)=   "Column(39)._WidthInPix=4233"
         Splits(0)._ColumnProps(160)=   "Column(39).Order=40"
         Splits(0)._ColumnProps(161)=   "Column(40).Width=4366"
         Splits(0)._ColumnProps(162)=   "Column(40).DividerColor=0"
         Splits(0)._ColumnProps(163)=   "Column(40)._WidthInPix=4233"
         Splits(0)._ColumnProps(164)=   "Column(40).Order=41"
         Splits(0)._ColumnProps(165)=   "Column(41).Width=4366"
         Splits(0)._ColumnProps(166)=   "Column(41).DividerColor=0"
         Splits(0)._ColumnProps(167)=   "Column(41)._WidthInPix=4233"
         Splits(0)._ColumnProps(168)=   "Column(41).Order=42"
         Splits(0)._ColumnProps(169)=   "Column(42).Width=4366"
         Splits(0)._ColumnProps(170)=   "Column(42).DividerColor=0"
         Splits(0)._ColumnProps(171)=   "Column(42)._WidthInPix=4233"
         Splits(0)._ColumnProps(172)=   "Column(42).Order=43"
         Splits(0)._ColumnProps(173)=   "Column(43).Width=4366"
         Splits(0)._ColumnProps(174)=   "Column(43).DividerColor=0"
         Splits(0)._ColumnProps(175)=   "Column(43)._WidthInPix=4233"
         Splits(0)._ColumnProps(176)=   "Column(43).Order=44"
         Splits(0)._ColumnProps(177)=   "Column(44).Width=4366"
         Splits(0)._ColumnProps(178)=   "Column(44).DividerColor=0"
         Splits(0)._ColumnProps(179)=   "Column(44)._WidthInPix=4233"
         Splits(0)._ColumnProps(180)=   "Column(44).Order=45"
         Splits(0)._ColumnProps(181)=   "Column(45).Width=4366"
         Splits(0)._ColumnProps(182)=   "Column(45).DividerColor=0"
         Splits(0)._ColumnProps(183)=   "Column(45)._WidthInPix=4233"
         Splits(0)._ColumnProps(184)=   "Column(45).Order=46"
         Splits(0)._ColumnProps(185)=   "Column(46).Width=4366"
         Splits(0)._ColumnProps(186)=   "Column(46).DividerColor=0"
         Splits(0)._ColumnProps(187)=   "Column(46)._WidthInPix=4233"
         Splits(0)._ColumnProps(188)=   "Column(46).Order=47"
         Splits(0)._ColumnProps(189)=   "Column(47).Width=4366"
         Splits(0)._ColumnProps(190)=   "Column(47).DividerColor=0"
         Splits(0)._ColumnProps(191)=   "Column(47)._WidthInPix=4233"
         Splits(0)._ColumnProps(192)=   "Column(47).Order=48"
         Splits(0)._ColumnProps(193)=   "Column(48).Width=4366"
         Splits(0)._ColumnProps(194)=   "Column(48).DividerColor=0"
         Splits(0)._ColumnProps(195)=   "Column(48)._WidthInPix=4233"
         Splits(0)._ColumnProps(196)=   "Column(48).Order=49"
         Splits(0)._ColumnProps(197)=   "Column(49).Width=4366"
         Splits(0)._ColumnProps(198)=   "Column(49).DividerColor=0"
         Splits(0)._ColumnProps(199)=   "Column(49)._WidthInPix=4233"
         Splits(0)._ColumnProps(200)=   "Column(49).Order=50"
         Splits(0)._ColumnProps(201)=   "Column(50).Width=4366"
         Splits(0)._ColumnProps(202)=   "Column(50).DividerColor=0"
         Splits(0)._ColumnProps(203)=   "Column(50)._WidthInPix=4233"
         Splits(0)._ColumnProps(204)=   "Column(50).Order=51"
         Splits(0)._ColumnProps(205)=   "Column(51).Width=4366"
         Splits(0)._ColumnProps(206)=   "Column(51).DividerColor=0"
         Splits(0)._ColumnProps(207)=   "Column(51)._WidthInPix=4233"
         Splits(0)._ColumnProps(208)=   "Column(51).Order=52"
         Splits(0)._ColumnProps(209)=   "Column(52).Width=4366"
         Splits(0)._ColumnProps(210)=   "Column(52).DividerColor=0"
         Splits(0)._ColumnProps(211)=   "Column(52)._WidthInPix=4233"
         Splits(0)._ColumnProps(212)=   "Column(52).Order=53"
         Splits(0)._ColumnProps(213)=   "Column(53).Width=4366"
         Splits(0)._ColumnProps(214)=   "Column(53).DividerColor=0"
         Splits(0)._ColumnProps(215)=   "Column(53)._WidthInPix=4233"
         Splits(0)._ColumnProps(216)=   "Column(53).Order=54"
         Splits(0)._ColumnProps(217)=   "Column(54).Width=4366"
         Splits(0)._ColumnProps(218)=   "Column(54).DividerColor=0"
         Splits(0)._ColumnProps(219)=   "Column(54)._WidthInPix=4233"
         Splits(0)._ColumnProps(220)=   "Column(54).Order=55"
         Splits(0)._ColumnProps(221)=   "Column(55).Width=4366"
         Splits(0)._ColumnProps(222)=   "Column(55).DividerColor=0"
         Splits(0)._ColumnProps(223)=   "Column(55)._WidthInPix=4233"
         Splits(0)._ColumnProps(224)=   "Column(55).Order=56"
         Splits(0)._ColumnProps(225)=   "Column(56).Width=4366"
         Splits(0)._ColumnProps(226)=   "Column(56).DividerColor=0"
         Splits(0)._ColumnProps(227)=   "Column(56)._WidthInPix=4233"
         Splits(0)._ColumnProps(228)=   "Column(56).Order=57"
         Splits(0)._ColumnProps(229)=   "Column(57).Width=4366"
         Splits(0)._ColumnProps(230)=   "Column(57).DividerColor=0"
         Splits(0)._ColumnProps(231)=   "Column(57)._WidthInPix=4233"
         Splits(0)._ColumnProps(232)=   "Column(57).Order=58"
         Splits(0)._ColumnProps(233)=   "Column(58).Width=4366"
         Splits(0)._ColumnProps(234)=   "Column(58).DividerColor=0"
         Splits(0)._ColumnProps(235)=   "Column(58)._WidthInPix=4233"
         Splits(0)._ColumnProps(236)=   "Column(58).Order=59"
         Splits(0)._ColumnProps(237)=   "Column(59).Width=4366"
         Splits(0)._ColumnProps(238)=   "Column(59).DividerColor=0"
         Splits(0)._ColumnProps(239)=   "Column(59)._WidthInPix=4233"
         Splits(0)._ColumnProps(240)=   "Column(59).Order=60"
         Splits(0)._ColumnProps(241)=   "Column(60).Width=4366"
         Splits(0)._ColumnProps(242)=   "Column(60).DividerColor=0"
         Splits(0)._ColumnProps(243)=   "Column(60)._WidthInPix=4233"
         Splits(0)._ColumnProps(244)=   "Column(60).Order=61"
         Splits(0)._ColumnProps(245)=   "Column(61).Width=4366"
         Splits(0)._ColumnProps(246)=   "Column(61).DividerColor=0"
         Splits(0)._ColumnProps(247)=   "Column(61)._WidthInPix=4233"
         Splits(0)._ColumnProps(248)=   "Column(61).Order=62"
         Splits(0)._ColumnProps(249)=   "Column(62).Width=4366"
         Splits(0)._ColumnProps(250)=   "Column(62).DividerColor=0"
         Splits(0)._ColumnProps(251)=   "Column(62)._WidthInPix=4233"
         Splits(0)._ColumnProps(252)=   "Column(62).Order=63"
         Splits(0)._ColumnProps(253)=   "Column(63).Width=4366"
         Splits(0)._ColumnProps(254)=   "Column(63).DividerColor=0"
         Splits(0)._ColumnProps(255)=   "Column(63)._WidthInPix=4233"
         Splits(0)._ColumnProps(256)=   "Column(63).Order=64"
         Splits(0)._ColumnProps(257)=   "Column(64).Width=4366"
         Splits(0)._ColumnProps(258)=   "Column(64).DividerColor=0"
         Splits(0)._ColumnProps(259)=   "Column(64)._WidthInPix=4233"
         Splits(0)._ColumnProps(260)=   "Column(64).Order=65"
         Splits(0)._ColumnProps(261)=   "Column(65).Width=4366"
         Splits(0)._ColumnProps(262)=   "Column(65).DividerColor=0"
         Splits(0)._ColumnProps(263)=   "Column(65)._WidthInPix=4233"
         Splits(0)._ColumnProps(264)=   "Column(65).Order=66"
         Splits(0)._ColumnProps(265)=   "Column(66).Width=4366"
         Splits(0)._ColumnProps(266)=   "Column(66).DividerColor=0"
         Splits(0)._ColumnProps(267)=   "Column(66)._WidthInPix=4233"
         Splits(0)._ColumnProps(268)=   "Column(66).Order=67"
         Splits(0)._ColumnProps(269)=   "Column(67).Width=4366"
         Splits(0)._ColumnProps(270)=   "Column(67).DividerColor=0"
         Splits(0)._ColumnProps(271)=   "Column(67)._WidthInPix=4233"
         Splits(0)._ColumnProps(272)=   "Column(67).Order=68"
         Splits(0)._ColumnProps(273)=   "Column(68).Width=4366"
         Splits(0)._ColumnProps(274)=   "Column(68).DividerColor=0"
         Splits(0)._ColumnProps(275)=   "Column(68)._WidthInPix=4233"
         Splits(0)._ColumnProps(276)=   "Column(68).Order=69"
         Splits(0)._ColumnProps(277)=   "Column(69).Width=4366"
         Splits(0)._ColumnProps(278)=   "Column(69).DividerColor=0"
         Splits(0)._ColumnProps(279)=   "Column(69)._WidthInPix=4233"
         Splits(0)._ColumnProps(280)=   "Column(69).Order=70"
         Splits(0)._ColumnProps(281)=   "Column(70).Width=4366"
         Splits(0)._ColumnProps(282)=   "Column(70).DividerColor=0"
         Splits(0)._ColumnProps(283)=   "Column(70)._WidthInPix=4233"
         Splits(0)._ColumnProps(284)=   "Column(70).Order=71"
         Splits(0)._ColumnProps(285)=   "Column(71).Width=4366"
         Splits(0)._ColumnProps(286)=   "Column(71).DividerColor=0"
         Splits(0)._ColumnProps(287)=   "Column(71)._WidthInPix=4233"
         Splits(0)._ColumnProps(288)=   "Column(71).Order=72"
         Splits(0)._ColumnProps(289)=   "Column(72).Width=4366"
         Splits(0)._ColumnProps(290)=   "Column(72).DividerColor=0"
         Splits(0)._ColumnProps(291)=   "Column(72)._WidthInPix=4233"
         Splits(0)._ColumnProps(292)=   "Column(72).Order=73"
         Splits(0)._ColumnProps(293)=   "Column(73).Width=4366"
         Splits(0)._ColumnProps(294)=   "Column(73).DividerColor=0"
         Splits(0)._ColumnProps(295)=   "Column(73)._WidthInPix=4233"
         Splits(0)._ColumnProps(296)=   "Column(73).Order=74"
         Splits(0)._ColumnProps(297)=   "Column(74).Width=4366"
         Splits(0)._ColumnProps(298)=   "Column(74).DividerColor=0"
         Splits(0)._ColumnProps(299)=   "Column(74)._WidthInPix=4233"
         Splits(0)._ColumnProps(300)=   "Column(74).Order=75"
         Splits(0)._ColumnProps(301)=   "Column(75).Width=4366"
         Splits(0)._ColumnProps(302)=   "Column(75).DividerColor=0"
         Splits(0)._ColumnProps(303)=   "Column(75)._WidthInPix=4233"
         Splits(0)._ColumnProps(304)=   "Column(75).Order=76"
         Splits(0)._ColumnProps(305)=   "Column(76).Width=4366"
         Splits(0)._ColumnProps(306)=   "Column(76).DividerColor=0"
         Splits(0)._ColumnProps(307)=   "Column(76)._WidthInPix=4233"
         Splits(0)._ColumnProps(308)=   "Column(76).Order=77"
         Splits(0)._ColumnProps(309)=   "Column(77).Width=4366"
         Splits(0)._ColumnProps(310)=   "Column(77).DividerColor=0"
         Splits(0)._ColumnProps(311)=   "Column(77)._WidthInPix=4233"
         Splits(0)._ColumnProps(312)=   "Column(77).Order=78"
         Splits(0)._ColumnProps(313)=   "Column(78).Width=4366"
         Splits(0)._ColumnProps(314)=   "Column(78).DividerColor=0"
         Splits(0)._ColumnProps(315)=   "Column(78)._WidthInPix=4233"
         Splits(0)._ColumnProps(316)=   "Column(78).Order=79"
         Splits(0)._ColumnProps(317)=   "Column(79).Width=4366"
         Splits(0)._ColumnProps(318)=   "Column(79).DividerColor=0"
         Splits(0)._ColumnProps(319)=   "Column(79)._WidthInPix=4233"
         Splits(0)._ColumnProps(320)=   "Column(79).Order=80"
         Splits(0)._ColumnProps(321)=   "Column(80).Width=4366"
         Splits(0)._ColumnProps(322)=   "Column(80).DividerColor=0"
         Splits(0)._ColumnProps(323)=   "Column(80)._WidthInPix=4233"
         Splits(0)._ColumnProps(324)=   "Column(80).Order=81"
         Splits(0)._ColumnProps(325)=   "Column(81).Width=4366"
         Splits(0)._ColumnProps(326)=   "Column(81).DividerColor=0"
         Splits(0)._ColumnProps(327)=   "Column(81)._WidthInPix=4233"
         Splits(0)._ColumnProps(328)=   "Column(81).Order=82"
         Splits(0)._ColumnProps(329)=   "Column(82).Width=2778"
         Splits(0)._ColumnProps(330)=   "Column(82).DividerColor=0"
         Splits(0)._ColumnProps(331)=   "Column(82)._WidthInPix=2646"
         Splits(0)._ColumnProps(332)=   "Column(82).Order=83"
         Splits(0)._ColumnProps(333)=   "Column(83).Width=2778"
         Splits(0)._ColumnProps(334)=   "Column(83).DividerColor=0"
         Splits(0)._ColumnProps(335)=   "Column(83)._WidthInPix=2646"
         Splits(0)._ColumnProps(336)=   "Column(83).Order=84"
         Splits(0)._ColumnProps(337)=   "Column(84).Width=2778"
         Splits(0)._ColumnProps(338)=   "Column(84).DividerColor=0"
         Splits(0)._ColumnProps(339)=   "Column(84)._WidthInPix=2646"
         Splits(0)._ColumnProps(340)=   "Column(84).Order=85"
         Splits(0)._ColumnProps(341)=   "Column(85).Width=2778"
         Splits(0)._ColumnProps(342)=   "Column(85).DividerColor=0"
         Splits(0)._ColumnProps(343)=   "Column(85)._WidthInPix=2646"
         Splits(0)._ColumnProps(344)=   "Column(85).Order=86"
         Splits(0)._ColumnProps(345)=   "Column(86).Width=2778"
         Splits(0)._ColumnProps(346)=   "Column(86).DividerColor=0"
         Splits(0)._ColumnProps(347)=   "Column(86)._WidthInPix=2646"
         Splits(0)._ColumnProps(348)=   "Column(86).Order=87"
         Splits(0)._ColumnProps(349)=   "Column(87).Width=2778"
         Splits(0)._ColumnProps(350)=   "Column(87).DividerColor=0"
         Splits(0)._ColumnProps(351)=   "Column(87)._WidthInPix=2646"
         Splits(0)._ColumnProps(352)=   "Column(87).Order=88"
         Splits(0)._ColumnProps(353)=   "Column(88).Width=2778"
         Splits(0)._ColumnProps(354)=   "Column(88).DividerColor=0"
         Splits(0)._ColumnProps(355)=   "Column(88)._WidthInPix=2646"
         Splits(0)._ColumnProps(356)=   "Column(88).Order=89"
         Splits(0)._ColumnProps(357)=   "Column(89).Width=2778"
         Splits(0)._ColumnProps(358)=   "Column(89).DividerColor=0"
         Splits(0)._ColumnProps(359)=   "Column(89)._WidthInPix=2646"
         Splits(0)._ColumnProps(360)=   "Column(89).Order=90"
         Splits(0)._ColumnProps(361)=   "Column(90).Width=2778"
         Splits(0)._ColumnProps(362)=   "Column(90).DividerColor=0"
         Splits(0)._ColumnProps(363)=   "Column(90)._WidthInPix=2646"
         Splits(0)._ColumnProps(364)=   "Column(90).Order=91"
         Splits(0)._ColumnProps(365)=   "Column(91).Width=2778"
         Splits(0)._ColumnProps(366)=   "Column(91).DividerColor=0"
         Splits(0)._ColumnProps(367)=   "Column(91)._WidthInPix=2646"
         Splits(0)._ColumnProps(368)=   "Column(91).Order=92"
         Splits(0)._ColumnProps(369)=   "Column(92).Width=2778"
         Splits(0)._ColumnProps(370)=   "Column(92).DividerColor=0"
         Splits(0)._ColumnProps(371)=   "Column(92)._WidthInPix=2646"
         Splits(0)._ColumnProps(372)=   "Column(92).Order=93"
         Splits(0)._ColumnProps(373)=   "Column(93).Width=2778"
         Splits(0)._ColumnProps(374)=   "Column(93).DividerColor=0"
         Splits(0)._ColumnProps(375)=   "Column(93)._WidthInPix=2646"
         Splits(0)._ColumnProps(376)=   "Column(93).Order=94"
         Splits(0)._ColumnProps(377)=   "Column(94).Width=2778"
         Splits(0)._ColumnProps(378)=   "Column(94).DividerColor=0"
         Splits(0)._ColumnProps(379)=   "Column(94)._WidthInPix=2646"
         Splits(0)._ColumnProps(380)=   "Column(94).Order=95"
         Splits(0)._ColumnProps(381)=   "Column(95).Width=2778"
         Splits(0)._ColumnProps(382)=   "Column(95).DividerColor=0"
         Splits(0)._ColumnProps(383)=   "Column(95)._WidthInPix=2646"
         Splits(0)._ColumnProps(384)=   "Column(95).Order=96"
         Splits(0)._ColumnProps(385)=   "Column(96).Width=2778"
         Splits(0)._ColumnProps(386)=   "Column(96).DividerColor=0"
         Splits(0)._ColumnProps(387)=   "Column(96)._WidthInPix=2646"
         Splits(0)._ColumnProps(388)=   "Column(96).Order=97"
         Splits(0)._ColumnProps(389)=   "Column(97).Width=2778"
         Splits(0)._ColumnProps(390)=   "Column(97).DividerColor=0"
         Splits(0)._ColumnProps(391)=   "Column(97)._WidthInPix=2646"
         Splits(0)._ColumnProps(392)=   "Column(97).Order=98"
         Splits(0)._ColumnProps(393)=   "Column(98).Width=2778"
         Splits(0)._ColumnProps(394)=   "Column(98).DividerColor=0"
         Splits(0)._ColumnProps(395)=   "Column(98)._WidthInPix=2646"
         Splits(0)._ColumnProps(396)=   "Column(98).Order=99"
         Splits(0)._ColumnProps(397)=   "Column(99).Width=2778"
         Splits(0)._ColumnProps(398)=   "Column(99).DividerColor=0"
         Splits(0)._ColumnProps(399)=   "Column(99)._WidthInPix=2646"
         Splits(0)._ColumnProps(400)=   "Column(99).Order=100"
         Splits(0)._ColumnProps(401)=   "Column(100).Width=2778"
         Splits(0)._ColumnProps(402)=   "Column(100).DividerColor=0"
         Splits(0)._ColumnProps(403)=   "Column(100)._WidthInPix=2646"
         Splits(0)._ColumnProps(404)=   "Column(100).Order=101"
         Splits(0)._ColumnProps(405)=   "Column(101).Width=2778"
         Splits(0)._ColumnProps(406)=   "Column(101).DividerColor=0"
         Splits(0)._ColumnProps(407)=   "Column(101)._WidthInPix=2646"
         Splits(0)._ColumnProps(408)=   "Column(101).Order=102"
         Splits(0)._ColumnProps(409)=   "Column(102).Width=2778"
         Splits(0)._ColumnProps(410)=   "Column(102).DividerColor=0"
         Splits(0)._ColumnProps(411)=   "Column(102)._WidthInPix=2646"
         Splits(0)._ColumnProps(412)=   "Column(102).Order=103"
         Splits(0)._ColumnProps(413)=   "Column(103).Width=2778"
         Splits(0)._ColumnProps(414)=   "Column(103).DividerColor=0"
         Splits(0)._ColumnProps(415)=   "Column(103)._WidthInPix=2646"
         Splits(0)._ColumnProps(416)=   "Column(103).Order=104"
         Splits(0)._ColumnProps(417)=   "Column(104).Width=2778"
         Splits(0)._ColumnProps(418)=   "Column(104).DividerColor=0"
         Splits(0)._ColumnProps(419)=   "Column(104)._WidthInPix=2646"
         Splits(0)._ColumnProps(420)=   "Column(104).Order=105"
         Splits(0)._ColumnProps(421)=   "Column(105).Width=2778"
         Splits(0)._ColumnProps(422)=   "Column(105).DividerColor=0"
         Splits(0)._ColumnProps(423)=   "Column(105)._WidthInPix=2646"
         Splits(0)._ColumnProps(424)=   "Column(105).Order=106"
         Splits(0)._ColumnProps(425)=   "Column(106).Width=2778"
         Splits(0)._ColumnProps(426)=   "Column(106).DividerColor=0"
         Splits(0)._ColumnProps(427)=   "Column(106)._WidthInPix=2646"
         Splits(0)._ColumnProps(428)=   "Column(106).Order=107"
         Splits(0)._ColumnProps(429)=   "Column(107).Width=2778"
         Splits(0)._ColumnProps(430)=   "Column(107).DividerColor=0"
         Splits(0)._ColumnProps(431)=   "Column(107)._WidthInPix=2646"
         Splits(0)._ColumnProps(432)=   "Column(107).Order=108"
         Splits(0)._ColumnProps(433)=   "Column(108).Width=2778"
         Splits(0)._ColumnProps(434)=   "Column(108).DividerColor=0"
         Splits(0)._ColumnProps(435)=   "Column(108)._WidthInPix=2646"
         Splits(0)._ColumnProps(436)=   "Column(108).Order=109"
         Splits(0)._ColumnProps(437)=   "Column(109).Width=2778"
         Splits(0)._ColumnProps(438)=   "Column(109).DividerColor=0"
         Splits(0)._ColumnProps(439)=   "Column(109)._WidthInPix=2646"
         Splits(0)._ColumnProps(440)=   "Column(109).Order=110"
         Splits(0)._ColumnProps(441)=   "Column(110).Width=2778"
         Splits(0)._ColumnProps(442)=   "Column(110).DividerColor=0"
         Splits(0)._ColumnProps(443)=   "Column(110)._WidthInPix=2646"
         Splits(0)._ColumnProps(444)=   "Column(110).Order=111"
         Splits(0)._ColumnProps(445)=   "Column(111).Width=2778"
         Splits(0)._ColumnProps(446)=   "Column(111).DividerColor=0"
         Splits(0)._ColumnProps(447)=   "Column(111)._WidthInPix=2646"
         Splits(0)._ColumnProps(448)=   "Column(111).Order=112"
         Splits(0)._ColumnProps(449)=   "Column(112).Width=2778"
         Splits(0)._ColumnProps(450)=   "Column(112).DividerColor=0"
         Splits(0)._ColumnProps(451)=   "Column(112)._WidthInPix=2646"
         Splits(0)._ColumnProps(452)=   "Column(112).Order=113"
         Splits(0)._ColumnProps(453)=   "Column(113).Width=2778"
         Splits(0)._ColumnProps(454)=   "Column(113).DividerColor=0"
         Splits(0)._ColumnProps(455)=   "Column(113)._WidthInPix=2646"
         Splits(0)._ColumnProps(456)=   "Column(113).Order=114"
         Splits(0)._ColumnProps(457)=   "Column(114).Width=2778"
         Splits(0)._ColumnProps(458)=   "Column(114).DividerColor=0"
         Splits(0)._ColumnProps(459)=   "Column(114)._WidthInPix=2646"
         Splits(0)._ColumnProps(460)=   "Column(114).Order=115"
         Splits(0)._ColumnProps(461)=   "Column(115).Width=2778"
         Splits(0)._ColumnProps(462)=   "Column(115).DividerColor=0"
         Splits(0)._ColumnProps(463)=   "Column(115)._WidthInPix=2646"
         Splits(0)._ColumnProps(464)=   "Column(115).Order=116"
         Splits(0)._ColumnProps(465)=   "Column(116).Width=2778"
         Splits(0)._ColumnProps(466)=   "Column(116).DividerColor=0"
         Splits(0)._ColumnProps(467)=   "Column(116)._WidthInPix=2646"
         Splits(0)._ColumnProps(468)=   "Column(116).Order=117"
         Splits(0)._ColumnProps(469)=   "Column(117).Width=2778"
         Splits(0)._ColumnProps(470)=   "Column(117).DividerColor=0"
         Splits(0)._ColumnProps(471)=   "Column(117)._WidthInPix=2646"
         Splits(0)._ColumnProps(472)=   "Column(117).Order=118"
         Splits(0)._ColumnProps(473)=   "Column(118).Width=2778"
         Splits(0)._ColumnProps(474)=   "Column(118).DividerColor=0"
         Splits(0)._ColumnProps(475)=   "Column(118)._WidthInPix=2646"
         Splits(0)._ColumnProps(476)=   "Column(118).Order=119"
         Splits(0)._ColumnProps(477)=   "Column(119).Width=2778"
         Splits(0)._ColumnProps(478)=   "Column(119).DividerColor=0"
         Splits(0)._ColumnProps(479)=   "Column(119)._WidthInPix=2646"
         Splits(0)._ColumnProps(480)=   "Column(119).Order=120"
         Splits(0)._ColumnProps(481)=   "Column(120).Width=2778"
         Splits(0)._ColumnProps(482)=   "Column(120).DividerColor=0"
         Splits(0)._ColumnProps(483)=   "Column(120)._WidthInPix=2646"
         Splits(0)._ColumnProps(484)=   "Column(120).Order=121"
         Splits(0)._ColumnProps(485)=   "Column(121).Width=2778"
         Splits(0)._ColumnProps(486)=   "Column(121).DividerColor=0"
         Splits(0)._ColumnProps(487)=   "Column(121)._WidthInPix=2646"
         Splits(0)._ColumnProps(488)=   "Column(121).Order=122"
         Splits(0)._ColumnProps(489)=   "Column(122).Width=2778"
         Splits(0)._ColumnProps(490)=   "Column(122).DividerColor=0"
         Splits(0)._ColumnProps(491)=   "Column(122)._WidthInPix=2646"
         Splits(0)._ColumnProps(492)=   "Column(122).Order=123"
         Splits(0)._ColumnProps(493)=   "Column(123).Width=2778"
         Splits(0)._ColumnProps(494)=   "Column(123).DividerColor=0"
         Splits(0)._ColumnProps(495)=   "Column(123)._WidthInPix=2646"
         Splits(0)._ColumnProps(496)=   "Column(123).Order=124"
         Splits(0)._ColumnProps(497)=   "Column(124).Width=2778"
         Splits(0)._ColumnProps(498)=   "Column(124).DividerColor=0"
         Splits(0)._ColumnProps(499)=   "Column(124)._WidthInPix=2646"
         Splits(0)._ColumnProps(500)=   "Column(124).Order=125"
         Splits(0)._ColumnProps(501)=   "Column(125).Width=2778"
         Splits(0)._ColumnProps(502)=   "Column(125).DividerColor=0"
         Splits(0)._ColumnProps(503)=   "Column(125)._WidthInPix=2646"
         Splits(0)._ColumnProps(504)=   "Column(125).Order=126"
         Splits(0)._ColumnProps(505)=   "Column(126).Width=2778"
         Splits(0)._ColumnProps(506)=   "Column(126).DividerColor=0"
         Splits(0)._ColumnProps(507)=   "Column(126)._WidthInPix=2646"
         Splits(0)._ColumnProps(508)=   "Column(126).Order=127"
         Splits(0)._ColumnProps(509)=   "Column(127).Width=2778"
         Splits(0)._ColumnProps(510)=   "Column(127).DividerColor=0"
         Splits(0)._ColumnProps(511)=   "Column(127)._WidthInPix=2646"
         Splits(0)._ColumnProps(512)=   "Column(127).Order=128"
         Splits(0)._ColumnProps(513)=   "Column(128).Width=2778"
         Splits(0)._ColumnProps(514)=   "Column(128).DividerColor=0"
         Splits(0)._ColumnProps(515)=   "Column(128)._WidthInPix=2646"
         Splits(0)._ColumnProps(516)=   "Column(128).Order=129"
         Splits(0)._ColumnProps(517)=   "Column(129).Width=2778"
         Splits(0)._ColumnProps(518)=   "Column(129).DividerColor=0"
         Splits(0)._ColumnProps(519)=   "Column(129)._WidthInPix=2646"
         Splits(0)._ColumnProps(520)=   "Column(129).Order=130"
         Splits(0)._ColumnProps(521)=   "Column(130).Width=2778"
         Splits(0)._ColumnProps(522)=   "Column(130).DividerColor=0"
         Splits(0)._ColumnProps(523)=   "Column(130)._WidthInPix=2646"
         Splits(0)._ColumnProps(524)=   "Column(130).Order=131"
         Splits(0)._ColumnProps(525)=   "Column(131).Width=2778"
         Splits(0)._ColumnProps(526)=   "Column(131).DividerColor=0"
         Splits(0)._ColumnProps(527)=   "Column(131)._WidthInPix=2646"
         Splits(0)._ColumnProps(528)=   "Column(131).Order=132"
         Splits(0)._ColumnProps(529)=   "Column(132).Width=2778"
         Splits(0)._ColumnProps(530)=   "Column(132).DividerColor=0"
         Splits(0)._ColumnProps(531)=   "Column(132)._WidthInPix=2646"
         Splits(0)._ColumnProps(532)=   "Column(132).Order=133"
         Splits(0)._ColumnProps(533)=   "Column(133).Width=2778"
         Splits(0)._ColumnProps(534)=   "Column(133).DividerColor=0"
         Splits(0)._ColumnProps(535)=   "Column(133)._WidthInPix=2646"
         Splits(0)._ColumnProps(536)=   "Column(133).Order=134"
         Splits(0)._ColumnProps(537)=   "Column(134).Width=2778"
         Splits(0)._ColumnProps(538)=   "Column(134).DividerColor=0"
         Splits(0)._ColumnProps(539)=   "Column(134)._WidthInPix=2646"
         Splits(0)._ColumnProps(540)=   "Column(134).Order=135"
         Splits(0)._ColumnProps(541)=   "Column(135).Width=2778"
         Splits(0)._ColumnProps(542)=   "Column(135).DividerColor=0"
         Splits(0)._ColumnProps(543)=   "Column(135)._WidthInPix=2646"
         Splits(0)._ColumnProps(544)=   "Column(135).Order=136"
         Splits(0)._ColumnProps(545)=   "Column(136).Width=2778"
         Splits(0)._ColumnProps(546)=   "Column(136).DividerColor=0"
         Splits(0)._ColumnProps(547)=   "Column(136)._WidthInPix=2646"
         Splits(0)._ColumnProps(548)=   "Column(136).Order=137"
         Splits(0)._ColumnProps(549)=   "Column(137).Width=2778"
         Splits(0)._ColumnProps(550)=   "Column(137).DividerColor=0"
         Splits(0)._ColumnProps(551)=   "Column(137)._WidthInPix=2646"
         Splits(0)._ColumnProps(552)=   "Column(137).Order=138"
         Splits(0)._ColumnProps(553)=   "Column(138).Width=2778"
         Splits(0)._ColumnProps(554)=   "Column(138).DividerColor=0"
         Splits(0)._ColumnProps(555)=   "Column(138)._WidthInPix=2646"
         Splits(0)._ColumnProps(556)=   "Column(138).Order=139"
         Splits(0)._ColumnProps(557)=   "Column(139).Width=2778"
         Splits(0)._ColumnProps(558)=   "Column(139).DividerColor=0"
         Splits(0)._ColumnProps(559)=   "Column(139)._WidthInPix=2646"
         Splits(0)._ColumnProps(560)=   "Column(139).Order=140"
         Splits(0)._ColumnProps(561)=   "Column(140).Width=2778"
         Splits(0)._ColumnProps(562)=   "Column(140).DividerColor=0"
         Splits(0)._ColumnProps(563)=   "Column(140)._WidthInPix=2646"
         Splits(0)._ColumnProps(564)=   "Column(140).Order=141"
         Splits(0)._ColumnProps(565)=   "Column(141).Width=2778"
         Splits(0)._ColumnProps(566)=   "Column(141).DividerColor=0"
         Splits(0)._ColumnProps(567)=   "Column(141)._WidthInPix=2646"
         Splits(0)._ColumnProps(568)=   "Column(141).Order=142"
         Splits(0)._ColumnProps(569)=   "Column(142).Width=2778"
         Splits(0)._ColumnProps(570)=   "Column(142).DividerColor=0"
         Splits(0)._ColumnProps(571)=   "Column(142)._WidthInPix=2646"
         Splits(0)._ColumnProps(572)=   "Column(142).Order=143"
         Splits(0)._ColumnProps(573)=   "Column(143).Width=2778"
         Splits(0)._ColumnProps(574)=   "Column(143).DividerColor=0"
         Splits(0)._ColumnProps(575)=   "Column(143)._WidthInPix=2646"
         Splits(0)._ColumnProps(576)=   "Column(143).Order=144"
         Splits(0)._ColumnProps(577)=   "Column(144).Width=2778"
         Splits(0)._ColumnProps(578)=   "Column(144).DividerColor=0"
         Splits(0)._ColumnProps(579)=   "Column(144)._WidthInPix=2646"
         Splits(0)._ColumnProps(580)=   "Column(144).Order=145"
         Splits(0)._ColumnProps(581)=   "Column(145).Width=2778"
         Splits(0)._ColumnProps(582)=   "Column(145).DividerColor=0"
         Splits(0)._ColumnProps(583)=   "Column(145)._WidthInPix=2646"
         Splits(0)._ColumnProps(584)=   "Column(145).Order=146"
         Splits(0)._ColumnProps(585)=   "Column(146).Width=2778"
         Splits(0)._ColumnProps(586)=   "Column(146).DividerColor=0"
         Splits(0)._ColumnProps(587)=   "Column(146)._WidthInPix=2646"
         Splits(0)._ColumnProps(588)=   "Column(146).Order=147"
         Splits(0)._ColumnProps(589)=   "Column(147).Width=2778"
         Splits(0)._ColumnProps(590)=   "Column(147).DividerColor=0"
         Splits(0)._ColumnProps(591)=   "Column(147)._WidthInPix=2646"
         Splits(0)._ColumnProps(592)=   "Column(147).Order=148"
         Splits(0)._ColumnProps(593)=   "Column(148).Width=2778"
         Splits(0)._ColumnProps(594)=   "Column(148).DividerColor=0"
         Splits(0)._ColumnProps(595)=   "Column(148)._WidthInPix=2646"
         Splits(0)._ColumnProps(596)=   "Column(148).Order=149"
         Splits(0)._ColumnProps(597)=   "Column(149).Width=2778"
         Splits(0)._ColumnProps(598)=   "Column(149).DividerColor=0"
         Splits(0)._ColumnProps(599)=   "Column(149)._WidthInPix=2646"
         Splits(0)._ColumnProps(600)=   "Column(149).Order=150"
         Splits(0)._ColumnProps(601)=   "Column(150).Width=2778"
         Splits(0)._ColumnProps(602)=   "Column(150).DividerColor=0"
         Splits(0)._ColumnProps(603)=   "Column(150)._WidthInPix=2646"
         Splits(0)._ColumnProps(604)=   "Column(150).Order=151"
         Splits(1)._UserFlags=   0
         Splits(1).AllowRowSizing=   0   'False
         Splits(1).RecordSelectorWidth=   503
         Splits(1).AllowColSelect=   0   'False
         Splits(1).DividerColor=   14869218
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=151"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=4366"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
         Splits(1)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(5)=   "Column(1).Width=4366"
         Splits(1)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
         Splits(1)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(9)=   "Column(2).Width=4366"
         Splits(1)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(11)=   "Column(2)._WidthInPix=4233"
         Splits(1)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(13)=   "Column(3).Width=4366"
         Splits(1)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(15)=   "Column(3)._WidthInPix=4233"
         Splits(1)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(17)=   "Column(4).Width=4366"
         Splits(1)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(19)=   "Column(4)._WidthInPix=4233"
         Splits(1)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(21)=   "Column(5).Width=4366"
         Splits(1)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(23)=   "Column(5)._WidthInPix=4233"
         Splits(1)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(25)=   "Column(6).Width=4366"
         Splits(1)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(27)=   "Column(6)._WidthInPix=4233"
         Splits(1)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(29)=   "Column(7).Width=4366"
         Splits(1)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(31)=   "Column(7)._WidthInPix=4233"
         Splits(1)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(1)._ColumnProps(33)=   "Column(8).Width=4366"
         Splits(1)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(35)=   "Column(8)._WidthInPix=4233"
         Splits(1)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(37)=   "Column(9).Width=4366"
         Splits(1)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(39)=   "Column(9)._WidthInPix=4233"
         Splits(1)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(41)=   "Column(10).Width=4366"
         Splits(1)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(43)=   "Column(10)._WidthInPix=4233"
         Splits(1)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(45)=   "Column(11).Width=4366"
         Splits(1)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(47)=   "Column(11)._WidthInPix=4233"
         Splits(1)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(49)=   "Column(12).Width=4366"
         Splits(1)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(51)=   "Column(12)._WidthInPix=4233"
         Splits(1)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(53)=   "Column(13).Width=4366"
         Splits(1)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(1)._ColumnProps(55)=   "Column(13)._WidthInPix=4233"
         Splits(1)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(1)._ColumnProps(57)=   "Column(14).Width=4366"
         Splits(1)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(1)._ColumnProps(59)=   "Column(14)._WidthInPix=4233"
         Splits(1)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(1)._ColumnProps(61)=   "Column(15).Width=4366"
         Splits(1)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(1)._ColumnProps(63)=   "Column(15)._WidthInPix=4233"
         Splits(1)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(1)._ColumnProps(65)=   "Column(16).Width=4366"
         Splits(1)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(1)._ColumnProps(67)=   "Column(16)._WidthInPix=4233"
         Splits(1)._ColumnProps(68)=   "Column(16).Order=17"
         Splits(1)._ColumnProps(69)=   "Column(17).Width=4366"
         Splits(1)._ColumnProps(70)=   "Column(17).DividerColor=0"
         Splits(1)._ColumnProps(71)=   "Column(17)._WidthInPix=4233"
         Splits(1)._ColumnProps(72)=   "Column(17).Order=18"
         Splits(1)._ColumnProps(73)=   "Column(18).Width=4366"
         Splits(1)._ColumnProps(74)=   "Column(18).DividerColor=0"
         Splits(1)._ColumnProps(75)=   "Column(18)._WidthInPix=4233"
         Splits(1)._ColumnProps(76)=   "Column(18).Order=19"
         Splits(1)._ColumnProps(77)=   "Column(19).Width=4366"
         Splits(1)._ColumnProps(78)=   "Column(19).DividerColor=0"
         Splits(1)._ColumnProps(79)=   "Column(19)._WidthInPix=4233"
         Splits(1)._ColumnProps(80)=   "Column(19).Order=20"
         Splits(1)._ColumnProps(81)=   "Column(20).Width=4366"
         Splits(1)._ColumnProps(82)=   "Column(20).DividerColor=0"
         Splits(1)._ColumnProps(83)=   "Column(20)._WidthInPix=4233"
         Splits(1)._ColumnProps(84)=   "Column(20).Order=21"
         Splits(1)._ColumnProps(85)=   "Column(21).Width=4366"
         Splits(1)._ColumnProps(86)=   "Column(21).DividerColor=0"
         Splits(1)._ColumnProps(87)=   "Column(21)._WidthInPix=4233"
         Splits(1)._ColumnProps(88)=   "Column(21).Order=22"
         Splits(1)._ColumnProps(89)=   "Column(22).Width=4366"
         Splits(1)._ColumnProps(90)=   "Column(22).DividerColor=0"
         Splits(1)._ColumnProps(91)=   "Column(22)._WidthInPix=4233"
         Splits(1)._ColumnProps(92)=   "Column(22).Order=23"
         Splits(1)._ColumnProps(93)=   "Column(23).Width=4366"
         Splits(1)._ColumnProps(94)=   "Column(23).DividerColor=0"
         Splits(1)._ColumnProps(95)=   "Column(23)._WidthInPix=4233"
         Splits(1)._ColumnProps(96)=   "Column(23).Order=24"
         Splits(1)._ColumnProps(97)=   "Column(24).Width=4366"
         Splits(1)._ColumnProps(98)=   "Column(24).DividerColor=0"
         Splits(1)._ColumnProps(99)=   "Column(24)._WidthInPix=4233"
         Splits(1)._ColumnProps(100)=   "Column(24).Order=25"
         Splits(1)._ColumnProps(101)=   "Column(25).Width=4366"
         Splits(1)._ColumnProps(102)=   "Column(25).DividerColor=0"
         Splits(1)._ColumnProps(103)=   "Column(25)._WidthInPix=4233"
         Splits(1)._ColumnProps(104)=   "Column(25).Order=26"
         Splits(1)._ColumnProps(105)=   "Column(26).Width=4366"
         Splits(1)._ColumnProps(106)=   "Column(26).DividerColor=0"
         Splits(1)._ColumnProps(107)=   "Column(26)._WidthInPix=4233"
         Splits(1)._ColumnProps(108)=   "Column(26).Order=27"
         Splits(1)._ColumnProps(109)=   "Column(27).Width=4366"
         Splits(1)._ColumnProps(110)=   "Column(27).DividerColor=0"
         Splits(1)._ColumnProps(111)=   "Column(27)._WidthInPix=4233"
         Splits(1)._ColumnProps(112)=   "Column(27).Order=28"
         Splits(1)._ColumnProps(113)=   "Column(28).Width=4366"
         Splits(1)._ColumnProps(114)=   "Column(28).DividerColor=0"
         Splits(1)._ColumnProps(115)=   "Column(28)._WidthInPix=4233"
         Splits(1)._ColumnProps(116)=   "Column(28).Order=29"
         Splits(1)._ColumnProps(117)=   "Column(29).Width=4366"
         Splits(1)._ColumnProps(118)=   "Column(29).DividerColor=0"
         Splits(1)._ColumnProps(119)=   "Column(29)._WidthInPix=4233"
         Splits(1)._ColumnProps(120)=   "Column(29).Order=30"
         Splits(1)._ColumnProps(121)=   "Column(30).Width=4366"
         Splits(1)._ColumnProps(122)=   "Column(30).DividerColor=0"
         Splits(1)._ColumnProps(123)=   "Column(30)._WidthInPix=4233"
         Splits(1)._ColumnProps(124)=   "Column(30).Order=31"
         Splits(1)._ColumnProps(125)=   "Column(31).Width=4366"
         Splits(1)._ColumnProps(126)=   "Column(31).DividerColor=0"
         Splits(1)._ColumnProps(127)=   "Column(31)._WidthInPix=4233"
         Splits(1)._ColumnProps(128)=   "Column(31).Order=32"
         Splits(1)._ColumnProps(129)=   "Column(32).Width=4366"
         Splits(1)._ColumnProps(130)=   "Column(32).DividerColor=0"
         Splits(1)._ColumnProps(131)=   "Column(32)._WidthInPix=4233"
         Splits(1)._ColumnProps(132)=   "Column(32).Order=33"
         Splits(1)._ColumnProps(133)=   "Column(33).Width=4366"
         Splits(1)._ColumnProps(134)=   "Column(33).DividerColor=0"
         Splits(1)._ColumnProps(135)=   "Column(33)._WidthInPix=4233"
         Splits(1)._ColumnProps(136)=   "Column(33).Order=34"
         Splits(1)._ColumnProps(137)=   "Column(34).Width=4366"
         Splits(1)._ColumnProps(138)=   "Column(34).DividerColor=0"
         Splits(1)._ColumnProps(139)=   "Column(34)._WidthInPix=4233"
         Splits(1)._ColumnProps(140)=   "Column(34).Order=35"
         Splits(1)._ColumnProps(141)=   "Column(35).Width=4366"
         Splits(1)._ColumnProps(142)=   "Column(35).DividerColor=0"
         Splits(1)._ColumnProps(143)=   "Column(35)._WidthInPix=4233"
         Splits(1)._ColumnProps(144)=   "Column(35).Order=36"
         Splits(1)._ColumnProps(145)=   "Column(36).Width=4366"
         Splits(1)._ColumnProps(146)=   "Column(36).DividerColor=0"
         Splits(1)._ColumnProps(147)=   "Column(36)._WidthInPix=4233"
         Splits(1)._ColumnProps(148)=   "Column(36).Order=37"
         Splits(1)._ColumnProps(149)=   "Column(37).Width=4366"
         Splits(1)._ColumnProps(150)=   "Column(37).DividerColor=0"
         Splits(1)._ColumnProps(151)=   "Column(37)._WidthInPix=4233"
         Splits(1)._ColumnProps(152)=   "Column(37).Order=38"
         Splits(1)._ColumnProps(153)=   "Column(38).Width=4366"
         Splits(1)._ColumnProps(154)=   "Column(38).DividerColor=0"
         Splits(1)._ColumnProps(155)=   "Column(38)._WidthInPix=4233"
         Splits(1)._ColumnProps(156)=   "Column(38).Order=39"
         Splits(1)._ColumnProps(157)=   "Column(39).Width=4366"
         Splits(1)._ColumnProps(158)=   "Column(39).DividerColor=0"
         Splits(1)._ColumnProps(159)=   "Column(39)._WidthInPix=4233"
         Splits(1)._ColumnProps(160)=   "Column(39).Order=40"
         Splits(1)._ColumnProps(161)=   "Column(40).Width=4366"
         Splits(1)._ColumnProps(162)=   "Column(40).DividerColor=0"
         Splits(1)._ColumnProps(163)=   "Column(40)._WidthInPix=4233"
         Splits(1)._ColumnProps(164)=   "Column(40).Order=41"
         Splits(1)._ColumnProps(165)=   "Column(41).Width=4366"
         Splits(1)._ColumnProps(166)=   "Column(41).DividerColor=0"
         Splits(1)._ColumnProps(167)=   "Column(41)._WidthInPix=4233"
         Splits(1)._ColumnProps(168)=   "Column(41).Order=42"
         Splits(1)._ColumnProps(169)=   "Column(42).Width=4366"
         Splits(1)._ColumnProps(170)=   "Column(42).DividerColor=0"
         Splits(1)._ColumnProps(171)=   "Column(42)._WidthInPix=4233"
         Splits(1)._ColumnProps(172)=   "Column(42).Order=43"
         Splits(1)._ColumnProps(173)=   "Column(43).Width=4366"
         Splits(1)._ColumnProps(174)=   "Column(43).DividerColor=0"
         Splits(1)._ColumnProps(175)=   "Column(43)._WidthInPix=4233"
         Splits(1)._ColumnProps(176)=   "Column(43).Order=44"
         Splits(1)._ColumnProps(177)=   "Column(44).Width=4366"
         Splits(1)._ColumnProps(178)=   "Column(44).DividerColor=0"
         Splits(1)._ColumnProps(179)=   "Column(44)._WidthInPix=4233"
         Splits(1)._ColumnProps(180)=   "Column(44).Order=45"
         Splits(1)._ColumnProps(181)=   "Column(45).Width=4366"
         Splits(1)._ColumnProps(182)=   "Column(45).DividerColor=0"
         Splits(1)._ColumnProps(183)=   "Column(45)._WidthInPix=4233"
         Splits(1)._ColumnProps(184)=   "Column(45).Order=46"
         Splits(1)._ColumnProps(185)=   "Column(46).Width=4366"
         Splits(1)._ColumnProps(186)=   "Column(46).DividerColor=0"
         Splits(1)._ColumnProps(187)=   "Column(46)._WidthInPix=4233"
         Splits(1)._ColumnProps(188)=   "Column(46).Order=47"
         Splits(1)._ColumnProps(189)=   "Column(47).Width=4366"
         Splits(1)._ColumnProps(190)=   "Column(47).DividerColor=0"
         Splits(1)._ColumnProps(191)=   "Column(47)._WidthInPix=4233"
         Splits(1)._ColumnProps(192)=   "Column(47).Order=48"
         Splits(1)._ColumnProps(193)=   "Column(48).Width=4366"
         Splits(1)._ColumnProps(194)=   "Column(48).DividerColor=0"
         Splits(1)._ColumnProps(195)=   "Column(48)._WidthInPix=4233"
         Splits(1)._ColumnProps(196)=   "Column(48).Order=49"
         Splits(1)._ColumnProps(197)=   "Column(49).Width=4366"
         Splits(1)._ColumnProps(198)=   "Column(49).DividerColor=0"
         Splits(1)._ColumnProps(199)=   "Column(49)._WidthInPix=4233"
         Splits(1)._ColumnProps(200)=   "Column(49).Order=50"
         Splits(1)._ColumnProps(201)=   "Column(50).Width=4366"
         Splits(1)._ColumnProps(202)=   "Column(50).DividerColor=0"
         Splits(1)._ColumnProps(203)=   "Column(50)._WidthInPix=4233"
         Splits(1)._ColumnProps(204)=   "Column(50).Order=51"
         Splits(1)._ColumnProps(205)=   "Column(51).Width=4366"
         Splits(1)._ColumnProps(206)=   "Column(51).DividerColor=0"
         Splits(1)._ColumnProps(207)=   "Column(51)._WidthInPix=4233"
         Splits(1)._ColumnProps(208)=   "Column(51).Order=52"
         Splits(1)._ColumnProps(209)=   "Column(52).Width=4366"
         Splits(1)._ColumnProps(210)=   "Column(52).DividerColor=0"
         Splits(1)._ColumnProps(211)=   "Column(52)._WidthInPix=4233"
         Splits(1)._ColumnProps(212)=   "Column(52).Order=53"
         Splits(1)._ColumnProps(213)=   "Column(53).Width=4366"
         Splits(1)._ColumnProps(214)=   "Column(53).DividerColor=0"
         Splits(1)._ColumnProps(215)=   "Column(53)._WidthInPix=4233"
         Splits(1)._ColumnProps(216)=   "Column(53).Order=54"
         Splits(1)._ColumnProps(217)=   "Column(54).Width=4366"
         Splits(1)._ColumnProps(218)=   "Column(54).DividerColor=0"
         Splits(1)._ColumnProps(219)=   "Column(54)._WidthInPix=4233"
         Splits(1)._ColumnProps(220)=   "Column(54).Order=55"
         Splits(1)._ColumnProps(221)=   "Column(55).Width=4366"
         Splits(1)._ColumnProps(222)=   "Column(55).DividerColor=0"
         Splits(1)._ColumnProps(223)=   "Column(55)._WidthInPix=4233"
         Splits(1)._ColumnProps(224)=   "Column(55).Order=56"
         Splits(1)._ColumnProps(225)=   "Column(56).Width=4366"
         Splits(1)._ColumnProps(226)=   "Column(56).DividerColor=0"
         Splits(1)._ColumnProps(227)=   "Column(56)._WidthInPix=4233"
         Splits(1)._ColumnProps(228)=   "Column(56).Order=57"
         Splits(1)._ColumnProps(229)=   "Column(57).Width=4366"
         Splits(1)._ColumnProps(230)=   "Column(57).DividerColor=0"
         Splits(1)._ColumnProps(231)=   "Column(57)._WidthInPix=4233"
         Splits(1)._ColumnProps(232)=   "Column(57).Order=58"
         Splits(1)._ColumnProps(233)=   "Column(58).Width=4366"
         Splits(1)._ColumnProps(234)=   "Column(58).DividerColor=0"
         Splits(1)._ColumnProps(235)=   "Column(58)._WidthInPix=4233"
         Splits(1)._ColumnProps(236)=   "Column(58).Order=59"
         Splits(1)._ColumnProps(237)=   "Column(59).Width=4366"
         Splits(1)._ColumnProps(238)=   "Column(59).DividerColor=0"
         Splits(1)._ColumnProps(239)=   "Column(59)._WidthInPix=4233"
         Splits(1)._ColumnProps(240)=   "Column(59).Order=60"
         Splits(1)._ColumnProps(241)=   "Column(60).Width=4366"
         Splits(1)._ColumnProps(242)=   "Column(60).DividerColor=0"
         Splits(1)._ColumnProps(243)=   "Column(60)._WidthInPix=4233"
         Splits(1)._ColumnProps(244)=   "Column(60).Order=61"
         Splits(1)._ColumnProps(245)=   "Column(61).Width=4366"
         Splits(1)._ColumnProps(246)=   "Column(61).DividerColor=0"
         Splits(1)._ColumnProps(247)=   "Column(61)._WidthInPix=4233"
         Splits(1)._ColumnProps(248)=   "Column(61).Order=62"
         Splits(1)._ColumnProps(249)=   "Column(62).Width=4366"
         Splits(1)._ColumnProps(250)=   "Column(62).DividerColor=0"
         Splits(1)._ColumnProps(251)=   "Column(62)._WidthInPix=4233"
         Splits(1)._ColumnProps(252)=   "Column(62).Order=63"
         Splits(1)._ColumnProps(253)=   "Column(63).Width=4366"
         Splits(1)._ColumnProps(254)=   "Column(63).DividerColor=0"
         Splits(1)._ColumnProps(255)=   "Column(63)._WidthInPix=4233"
         Splits(1)._ColumnProps(256)=   "Column(63).Order=64"
         Splits(1)._ColumnProps(257)=   "Column(64).Width=4366"
         Splits(1)._ColumnProps(258)=   "Column(64).DividerColor=0"
         Splits(1)._ColumnProps(259)=   "Column(64)._WidthInPix=4233"
         Splits(1)._ColumnProps(260)=   "Column(64).Order=65"
         Splits(1)._ColumnProps(261)=   "Column(65).Width=4366"
         Splits(1)._ColumnProps(262)=   "Column(65).DividerColor=0"
         Splits(1)._ColumnProps(263)=   "Column(65)._WidthInPix=4233"
         Splits(1)._ColumnProps(264)=   "Column(65).Order=66"
         Splits(1)._ColumnProps(265)=   "Column(66).Width=4366"
         Splits(1)._ColumnProps(266)=   "Column(66).DividerColor=0"
         Splits(1)._ColumnProps(267)=   "Column(66)._WidthInPix=4233"
         Splits(1)._ColumnProps(268)=   "Column(66).Order=67"
         Splits(1)._ColumnProps(269)=   "Column(67).Width=4366"
         Splits(1)._ColumnProps(270)=   "Column(67).DividerColor=0"
         Splits(1)._ColumnProps(271)=   "Column(67)._WidthInPix=4233"
         Splits(1)._ColumnProps(272)=   "Column(67).Order=68"
         Splits(1)._ColumnProps(273)=   "Column(68).Width=4366"
         Splits(1)._ColumnProps(274)=   "Column(68).DividerColor=0"
         Splits(1)._ColumnProps(275)=   "Column(68)._WidthInPix=4233"
         Splits(1)._ColumnProps(276)=   "Column(68).Order=69"
         Splits(1)._ColumnProps(277)=   "Column(69).Width=4366"
         Splits(1)._ColumnProps(278)=   "Column(69).DividerColor=0"
         Splits(1)._ColumnProps(279)=   "Column(69)._WidthInPix=4233"
         Splits(1)._ColumnProps(280)=   "Column(69).Order=70"
         Splits(1)._ColumnProps(281)=   "Column(70).Width=4366"
         Splits(1)._ColumnProps(282)=   "Column(70).DividerColor=0"
         Splits(1)._ColumnProps(283)=   "Column(70)._WidthInPix=4233"
         Splits(1)._ColumnProps(284)=   "Column(70).Order=71"
         Splits(1)._ColumnProps(285)=   "Column(71).Width=4366"
         Splits(1)._ColumnProps(286)=   "Column(71).DividerColor=0"
         Splits(1)._ColumnProps(287)=   "Column(71)._WidthInPix=4233"
         Splits(1)._ColumnProps(288)=   "Column(71).Order=72"
         Splits(1)._ColumnProps(289)=   "Column(72).Width=4366"
         Splits(1)._ColumnProps(290)=   "Column(72).DividerColor=0"
         Splits(1)._ColumnProps(291)=   "Column(72)._WidthInPix=4233"
         Splits(1)._ColumnProps(292)=   "Column(72).Order=73"
         Splits(1)._ColumnProps(293)=   "Column(73).Width=4366"
         Splits(1)._ColumnProps(294)=   "Column(73).DividerColor=0"
         Splits(1)._ColumnProps(295)=   "Column(73)._WidthInPix=4233"
         Splits(1)._ColumnProps(296)=   "Column(73).Order=74"
         Splits(1)._ColumnProps(297)=   "Column(74).Width=4366"
         Splits(1)._ColumnProps(298)=   "Column(74).DividerColor=0"
         Splits(1)._ColumnProps(299)=   "Column(74)._WidthInPix=4233"
         Splits(1)._ColumnProps(300)=   "Column(74).Order=75"
         Splits(1)._ColumnProps(301)=   "Column(75).Width=4366"
         Splits(1)._ColumnProps(302)=   "Column(75).DividerColor=0"
         Splits(1)._ColumnProps(303)=   "Column(75)._WidthInPix=4233"
         Splits(1)._ColumnProps(304)=   "Column(75).Order=76"
         Splits(1)._ColumnProps(305)=   "Column(76).Width=4366"
         Splits(1)._ColumnProps(306)=   "Column(76).DividerColor=0"
         Splits(1)._ColumnProps(307)=   "Column(76)._WidthInPix=4233"
         Splits(1)._ColumnProps(308)=   "Column(76).Order=77"
         Splits(1)._ColumnProps(309)=   "Column(77).Width=4366"
         Splits(1)._ColumnProps(310)=   "Column(77).DividerColor=0"
         Splits(1)._ColumnProps(311)=   "Column(77)._WidthInPix=4233"
         Splits(1)._ColumnProps(312)=   "Column(77).Order=78"
         Splits(1)._ColumnProps(313)=   "Column(78).Width=4366"
         Splits(1)._ColumnProps(314)=   "Column(78).DividerColor=0"
         Splits(1)._ColumnProps(315)=   "Column(78)._WidthInPix=4233"
         Splits(1)._ColumnProps(316)=   "Column(78).Order=79"
         Splits(1)._ColumnProps(317)=   "Column(79).Width=4366"
         Splits(1)._ColumnProps(318)=   "Column(79).DividerColor=0"
         Splits(1)._ColumnProps(319)=   "Column(79)._WidthInPix=4233"
         Splits(1)._ColumnProps(320)=   "Column(79).Order=80"
         Splits(1)._ColumnProps(321)=   "Column(80).Width=4366"
         Splits(1)._ColumnProps(322)=   "Column(80).DividerColor=0"
         Splits(1)._ColumnProps(323)=   "Column(80)._WidthInPix=4233"
         Splits(1)._ColumnProps(324)=   "Column(80).Order=81"
         Splits(1)._ColumnProps(325)=   "Column(81).Width=4366"
         Splits(1)._ColumnProps(326)=   "Column(81).DividerColor=0"
         Splits(1)._ColumnProps(327)=   "Column(81)._WidthInPix=4233"
         Splits(1)._ColumnProps(328)=   "Column(81).Order=82"
         Splits(1)._ColumnProps(329)=   "Column(82).Width=2778"
         Splits(1)._ColumnProps(330)=   "Column(82).DividerColor=0"
         Splits(1)._ColumnProps(331)=   "Column(82)._WidthInPix=2646"
         Splits(1)._ColumnProps(332)=   "Column(82).Visible=0"
         Splits(1)._ColumnProps(333)=   "Column(82).Order=83"
         Splits(1)._ColumnProps(334)=   "Column(83).Width=2778"
         Splits(1)._ColumnProps(335)=   "Column(83).DividerColor=0"
         Splits(1)._ColumnProps(336)=   "Column(83)._WidthInPix=2646"
         Splits(1)._ColumnProps(337)=   "Column(83).Order=84"
         Splits(1)._ColumnProps(338)=   "Column(84).Width=2778"
         Splits(1)._ColumnProps(339)=   "Column(84).DividerColor=0"
         Splits(1)._ColumnProps(340)=   "Column(84)._WidthInPix=2646"
         Splits(1)._ColumnProps(341)=   "Column(84).Order=85"
         Splits(1)._ColumnProps(342)=   "Column(85).Width=2778"
         Splits(1)._ColumnProps(343)=   "Column(85).DividerColor=0"
         Splits(1)._ColumnProps(344)=   "Column(85)._WidthInPix=2646"
         Splits(1)._ColumnProps(345)=   "Column(85).Order=86"
         Splits(1)._ColumnProps(346)=   "Column(86).Width=2778"
         Splits(1)._ColumnProps(347)=   "Column(86).DividerColor=0"
         Splits(1)._ColumnProps(348)=   "Column(86)._WidthInPix=2646"
         Splits(1)._ColumnProps(349)=   "Column(86).Order=87"
         Splits(1)._ColumnProps(350)=   "Column(87).Width=2778"
         Splits(1)._ColumnProps(351)=   "Column(87).DividerColor=0"
         Splits(1)._ColumnProps(352)=   "Column(87)._WidthInPix=2646"
         Splits(1)._ColumnProps(353)=   "Column(87).Order=88"
         Splits(1)._ColumnProps(354)=   "Column(88).Width=2778"
         Splits(1)._ColumnProps(355)=   "Column(88).DividerColor=0"
         Splits(1)._ColumnProps(356)=   "Column(88)._WidthInPix=2646"
         Splits(1)._ColumnProps(357)=   "Column(88).Order=89"
         Splits(1)._ColumnProps(358)=   "Column(89).Width=2778"
         Splits(1)._ColumnProps(359)=   "Column(89).DividerColor=0"
         Splits(1)._ColumnProps(360)=   "Column(89)._WidthInPix=2646"
         Splits(1)._ColumnProps(361)=   "Column(89).Order=90"
         Splits(1)._ColumnProps(362)=   "Column(90).Width=2778"
         Splits(1)._ColumnProps(363)=   "Column(90).DividerColor=0"
         Splits(1)._ColumnProps(364)=   "Column(90)._WidthInPix=2646"
         Splits(1)._ColumnProps(365)=   "Column(90).Order=91"
         Splits(1)._ColumnProps(366)=   "Column(91).Width=2778"
         Splits(1)._ColumnProps(367)=   "Column(91).DividerColor=0"
         Splits(1)._ColumnProps(368)=   "Column(91)._WidthInPix=2646"
         Splits(1)._ColumnProps(369)=   "Column(91).Order=92"
         Splits(1)._ColumnProps(370)=   "Column(92).Width=2778"
         Splits(1)._ColumnProps(371)=   "Column(92).DividerColor=0"
         Splits(1)._ColumnProps(372)=   "Column(92)._WidthInPix=2646"
         Splits(1)._ColumnProps(373)=   "Column(92).Order=93"
         Splits(1)._ColumnProps(374)=   "Column(93).Width=2778"
         Splits(1)._ColumnProps(375)=   "Column(93).DividerColor=0"
         Splits(1)._ColumnProps(376)=   "Column(93)._WidthInPix=2646"
         Splits(1)._ColumnProps(377)=   "Column(93).Order=94"
         Splits(1)._ColumnProps(378)=   "Column(94).Width=2778"
         Splits(1)._ColumnProps(379)=   "Column(94).DividerColor=0"
         Splits(1)._ColumnProps(380)=   "Column(94)._WidthInPix=2646"
         Splits(1)._ColumnProps(381)=   "Column(94).Order=95"
         Splits(1)._ColumnProps(382)=   "Column(95).Width=2778"
         Splits(1)._ColumnProps(383)=   "Column(95).DividerColor=0"
         Splits(1)._ColumnProps(384)=   "Column(95)._WidthInPix=2646"
         Splits(1)._ColumnProps(385)=   "Column(95).Order=96"
         Splits(1)._ColumnProps(386)=   "Column(96).Width=2778"
         Splits(1)._ColumnProps(387)=   "Column(96).DividerColor=0"
         Splits(1)._ColumnProps(388)=   "Column(96)._WidthInPix=2646"
         Splits(1)._ColumnProps(389)=   "Column(96).Order=97"
         Splits(1)._ColumnProps(390)=   "Column(97).Width=2778"
         Splits(1)._ColumnProps(391)=   "Column(97).DividerColor=0"
         Splits(1)._ColumnProps(392)=   "Column(97)._WidthInPix=2646"
         Splits(1)._ColumnProps(393)=   "Column(97).Order=98"
         Splits(1)._ColumnProps(394)=   "Column(98).Width=2778"
         Splits(1)._ColumnProps(395)=   "Column(98).DividerColor=0"
         Splits(1)._ColumnProps(396)=   "Column(98)._WidthInPix=2646"
         Splits(1)._ColumnProps(397)=   "Column(98).Order=99"
         Splits(1)._ColumnProps(398)=   "Column(99).Width=2778"
         Splits(1)._ColumnProps(399)=   "Column(99).DividerColor=0"
         Splits(1)._ColumnProps(400)=   "Column(99)._WidthInPix=2646"
         Splits(1)._ColumnProps(401)=   "Column(99).Order=100"
         Splits(1)._ColumnProps(402)=   "Column(100).Width=2778"
         Splits(1)._ColumnProps(403)=   "Column(100).DividerColor=0"
         Splits(1)._ColumnProps(404)=   "Column(100)._WidthInPix=2646"
         Splits(1)._ColumnProps(405)=   "Column(100).Order=101"
         Splits(1)._ColumnProps(406)=   "Column(101).Width=2778"
         Splits(1)._ColumnProps(407)=   "Column(101).DividerColor=0"
         Splits(1)._ColumnProps(408)=   "Column(101)._WidthInPix=2646"
         Splits(1)._ColumnProps(409)=   "Column(101).Order=102"
         Splits(1)._ColumnProps(410)=   "Column(102).Width=2778"
         Splits(1)._ColumnProps(411)=   "Column(102).DividerColor=0"
         Splits(1)._ColumnProps(412)=   "Column(102)._WidthInPix=2646"
         Splits(1)._ColumnProps(413)=   "Column(102).Order=103"
         Splits(1)._ColumnProps(414)=   "Column(103).Width=2778"
         Splits(1)._ColumnProps(415)=   "Column(103).DividerColor=0"
         Splits(1)._ColumnProps(416)=   "Column(103)._WidthInPix=2646"
         Splits(1)._ColumnProps(417)=   "Column(103).Order=104"
         Splits(1)._ColumnProps(418)=   "Column(104).Width=2778"
         Splits(1)._ColumnProps(419)=   "Column(104).DividerColor=0"
         Splits(1)._ColumnProps(420)=   "Column(104)._WidthInPix=2646"
         Splits(1)._ColumnProps(421)=   "Column(104).Order=105"
         Splits(1)._ColumnProps(422)=   "Column(105).Width=2778"
         Splits(1)._ColumnProps(423)=   "Column(105).DividerColor=0"
         Splits(1)._ColumnProps(424)=   "Column(105)._WidthInPix=2646"
         Splits(1)._ColumnProps(425)=   "Column(105).Order=106"
         Splits(1)._ColumnProps(426)=   "Column(106).Width=2778"
         Splits(1)._ColumnProps(427)=   "Column(106).DividerColor=0"
         Splits(1)._ColumnProps(428)=   "Column(106)._WidthInPix=2646"
         Splits(1)._ColumnProps(429)=   "Column(106).Order=107"
         Splits(1)._ColumnProps(430)=   "Column(107).Width=2778"
         Splits(1)._ColumnProps(431)=   "Column(107).DividerColor=0"
         Splits(1)._ColumnProps(432)=   "Column(107)._WidthInPix=2646"
         Splits(1)._ColumnProps(433)=   "Column(107).Order=108"
         Splits(1)._ColumnProps(434)=   "Column(108).Width=2778"
         Splits(1)._ColumnProps(435)=   "Column(108).DividerColor=0"
         Splits(1)._ColumnProps(436)=   "Column(108)._WidthInPix=2646"
         Splits(1)._ColumnProps(437)=   "Column(108).Order=109"
         Splits(1)._ColumnProps(438)=   "Column(109).Width=2778"
         Splits(1)._ColumnProps(439)=   "Column(109).DividerColor=0"
         Splits(1)._ColumnProps(440)=   "Column(109)._WidthInPix=2646"
         Splits(1)._ColumnProps(441)=   "Column(109).Order=110"
         Splits(1)._ColumnProps(442)=   "Column(110).Width=2778"
         Splits(1)._ColumnProps(443)=   "Column(110).DividerColor=0"
         Splits(1)._ColumnProps(444)=   "Column(110)._WidthInPix=2646"
         Splits(1)._ColumnProps(445)=   "Column(110).Order=111"
         Splits(1)._ColumnProps(446)=   "Column(111).Width=2778"
         Splits(1)._ColumnProps(447)=   "Column(111).DividerColor=0"
         Splits(1)._ColumnProps(448)=   "Column(111)._WidthInPix=2646"
         Splits(1)._ColumnProps(449)=   "Column(111).Order=112"
         Splits(1)._ColumnProps(450)=   "Column(112).Width=2778"
         Splits(1)._ColumnProps(451)=   "Column(112).DividerColor=0"
         Splits(1)._ColumnProps(452)=   "Column(112)._WidthInPix=2646"
         Splits(1)._ColumnProps(453)=   "Column(112).Order=113"
         Splits(1)._ColumnProps(454)=   "Column(113).Width=2778"
         Splits(1)._ColumnProps(455)=   "Column(113).DividerColor=0"
         Splits(1)._ColumnProps(456)=   "Column(113)._WidthInPix=2646"
         Splits(1)._ColumnProps(457)=   "Column(113).Order=114"
         Splits(1)._ColumnProps(458)=   "Column(114).Width=2778"
         Splits(1)._ColumnProps(459)=   "Column(114).DividerColor=0"
         Splits(1)._ColumnProps(460)=   "Column(114)._WidthInPix=2646"
         Splits(1)._ColumnProps(461)=   "Column(114).Order=115"
         Splits(1)._ColumnProps(462)=   "Column(115).Width=2778"
         Splits(1)._ColumnProps(463)=   "Column(115).DividerColor=0"
         Splits(1)._ColumnProps(464)=   "Column(115)._WidthInPix=2646"
         Splits(1)._ColumnProps(465)=   "Column(115).Order=116"
         Splits(1)._ColumnProps(466)=   "Column(116).Width=2778"
         Splits(1)._ColumnProps(467)=   "Column(116).DividerColor=0"
         Splits(1)._ColumnProps(468)=   "Column(116)._WidthInPix=2646"
         Splits(1)._ColumnProps(469)=   "Column(116).Order=117"
         Splits(1)._ColumnProps(470)=   "Column(117).Width=2778"
         Splits(1)._ColumnProps(471)=   "Column(117).DividerColor=0"
         Splits(1)._ColumnProps(472)=   "Column(117)._WidthInPix=2646"
         Splits(1)._ColumnProps(473)=   "Column(117).Order=118"
         Splits(1)._ColumnProps(474)=   "Column(118).Width=2778"
         Splits(1)._ColumnProps(475)=   "Column(118).DividerColor=0"
         Splits(1)._ColumnProps(476)=   "Column(118)._WidthInPix=2646"
         Splits(1)._ColumnProps(477)=   "Column(118).Order=119"
         Splits(1)._ColumnProps(478)=   "Column(119).Width=2778"
         Splits(1)._ColumnProps(479)=   "Column(119).DividerColor=0"
         Splits(1)._ColumnProps(480)=   "Column(119)._WidthInPix=2646"
         Splits(1)._ColumnProps(481)=   "Column(119).Order=120"
         Splits(1)._ColumnProps(482)=   "Column(120).Width=2778"
         Splits(1)._ColumnProps(483)=   "Column(120).DividerColor=0"
         Splits(1)._ColumnProps(484)=   "Column(120)._WidthInPix=2646"
         Splits(1)._ColumnProps(485)=   "Column(120).Order=121"
         Splits(1)._ColumnProps(486)=   "Column(121).Width=2778"
         Splits(1)._ColumnProps(487)=   "Column(121).DividerColor=0"
         Splits(1)._ColumnProps(488)=   "Column(121)._WidthInPix=2646"
         Splits(1)._ColumnProps(489)=   "Column(121).Order=122"
         Splits(1)._ColumnProps(490)=   "Column(122).Width=2778"
         Splits(1)._ColumnProps(491)=   "Column(122).DividerColor=0"
         Splits(1)._ColumnProps(492)=   "Column(122)._WidthInPix=2646"
         Splits(1)._ColumnProps(493)=   "Column(122).Order=123"
         Splits(1)._ColumnProps(494)=   "Column(123).Width=2778"
         Splits(1)._ColumnProps(495)=   "Column(123).DividerColor=0"
         Splits(1)._ColumnProps(496)=   "Column(123)._WidthInPix=2646"
         Splits(1)._ColumnProps(497)=   "Column(123).Order=124"
         Splits(1)._ColumnProps(498)=   "Column(124).Width=2778"
         Splits(1)._ColumnProps(499)=   "Column(124).DividerColor=0"
         Splits(1)._ColumnProps(500)=   "Column(124)._WidthInPix=2646"
         Splits(1)._ColumnProps(501)=   "Column(124).Order=125"
         Splits(1)._ColumnProps(502)=   "Column(125).Width=2778"
         Splits(1)._ColumnProps(503)=   "Column(125).DividerColor=0"
         Splits(1)._ColumnProps(504)=   "Column(125)._WidthInPix=2646"
         Splits(1)._ColumnProps(505)=   "Column(125).Order=126"
         Splits(1)._ColumnProps(506)=   "Column(126).Width=2778"
         Splits(1)._ColumnProps(507)=   "Column(126).DividerColor=0"
         Splits(1)._ColumnProps(508)=   "Column(126)._WidthInPix=2646"
         Splits(1)._ColumnProps(509)=   "Column(126).Order=127"
         Splits(1)._ColumnProps(510)=   "Column(127).Width=2778"
         Splits(1)._ColumnProps(511)=   "Column(127).DividerColor=0"
         Splits(1)._ColumnProps(512)=   "Column(127)._WidthInPix=2646"
         Splits(1)._ColumnProps(513)=   "Column(127).Order=128"
         Splits(1)._ColumnProps(514)=   "Column(128).Width=2778"
         Splits(1)._ColumnProps(515)=   "Column(128).DividerColor=0"
         Splits(1)._ColumnProps(516)=   "Column(128)._WidthInPix=2646"
         Splits(1)._ColumnProps(517)=   "Column(128).Order=129"
         Splits(1)._ColumnProps(518)=   "Column(129).Width=2778"
         Splits(1)._ColumnProps(519)=   "Column(129).DividerColor=0"
         Splits(1)._ColumnProps(520)=   "Column(129)._WidthInPix=2646"
         Splits(1)._ColumnProps(521)=   "Column(129).Order=130"
         Splits(1)._ColumnProps(522)=   "Column(130).Width=2778"
         Splits(1)._ColumnProps(523)=   "Column(130).DividerColor=0"
         Splits(1)._ColumnProps(524)=   "Column(130)._WidthInPix=2646"
         Splits(1)._ColumnProps(525)=   "Column(130).Order=131"
         Splits(1)._ColumnProps(526)=   "Column(131).Width=2778"
         Splits(1)._ColumnProps(527)=   "Column(131).DividerColor=0"
         Splits(1)._ColumnProps(528)=   "Column(131)._WidthInPix=2646"
         Splits(1)._ColumnProps(529)=   "Column(131).Order=132"
         Splits(1)._ColumnProps(530)=   "Column(132).Width=2778"
         Splits(1)._ColumnProps(531)=   "Column(132).DividerColor=0"
         Splits(1)._ColumnProps(532)=   "Column(132)._WidthInPix=2646"
         Splits(1)._ColumnProps(533)=   "Column(132).Order=133"
         Splits(1)._ColumnProps(534)=   "Column(133).Width=2778"
         Splits(1)._ColumnProps(535)=   "Column(133).DividerColor=0"
         Splits(1)._ColumnProps(536)=   "Column(133)._WidthInPix=2646"
         Splits(1)._ColumnProps(537)=   "Column(133).Order=134"
         Splits(1)._ColumnProps(538)=   "Column(134).Width=2778"
         Splits(1)._ColumnProps(539)=   "Column(134).DividerColor=0"
         Splits(1)._ColumnProps(540)=   "Column(134)._WidthInPix=2646"
         Splits(1)._ColumnProps(541)=   "Column(134).Order=135"
         Splits(1)._ColumnProps(542)=   "Column(135).Width=2778"
         Splits(1)._ColumnProps(543)=   "Column(135).DividerColor=0"
         Splits(1)._ColumnProps(544)=   "Column(135)._WidthInPix=2646"
         Splits(1)._ColumnProps(545)=   "Column(135).Order=136"
         Splits(1)._ColumnProps(546)=   "Column(136).Width=2778"
         Splits(1)._ColumnProps(547)=   "Column(136).DividerColor=0"
         Splits(1)._ColumnProps(548)=   "Column(136)._WidthInPix=2646"
         Splits(1)._ColumnProps(549)=   "Column(136).Order=137"
         Splits(1)._ColumnProps(550)=   "Column(137).Width=2778"
         Splits(1)._ColumnProps(551)=   "Column(137).DividerColor=0"
         Splits(1)._ColumnProps(552)=   "Column(137)._WidthInPix=2646"
         Splits(1)._ColumnProps(553)=   "Column(137).Order=138"
         Splits(1)._ColumnProps(554)=   "Column(138).Width=2778"
         Splits(1)._ColumnProps(555)=   "Column(138).DividerColor=0"
         Splits(1)._ColumnProps(556)=   "Column(138)._WidthInPix=2646"
         Splits(1)._ColumnProps(557)=   "Column(138).Order=139"
         Splits(1)._ColumnProps(558)=   "Column(139).Width=2778"
         Splits(1)._ColumnProps(559)=   "Column(139).DividerColor=0"
         Splits(1)._ColumnProps(560)=   "Column(139)._WidthInPix=2646"
         Splits(1)._ColumnProps(561)=   "Column(139).Order=140"
         Splits(1)._ColumnProps(562)=   "Column(140).Width=2778"
         Splits(1)._ColumnProps(563)=   "Column(140).DividerColor=0"
         Splits(1)._ColumnProps(564)=   "Column(140)._WidthInPix=2646"
         Splits(1)._ColumnProps(565)=   "Column(140).Order=141"
         Splits(1)._ColumnProps(566)=   "Column(141).Width=2778"
         Splits(1)._ColumnProps(567)=   "Column(141).DividerColor=0"
         Splits(1)._ColumnProps(568)=   "Column(141)._WidthInPix=2646"
         Splits(1)._ColumnProps(569)=   "Column(141).Order=142"
         Splits(1)._ColumnProps(570)=   "Column(142).Width=2778"
         Splits(1)._ColumnProps(571)=   "Column(142).DividerColor=0"
         Splits(1)._ColumnProps(572)=   "Column(142)._WidthInPix=2646"
         Splits(1)._ColumnProps(573)=   "Column(142).Order=143"
         Splits(1)._ColumnProps(574)=   "Column(143).Width=2778"
         Splits(1)._ColumnProps(575)=   "Column(143).DividerColor=0"
         Splits(1)._ColumnProps(576)=   "Column(143)._WidthInPix=2646"
         Splits(1)._ColumnProps(577)=   "Column(143).Order=144"
         Splits(1)._ColumnProps(578)=   "Column(144).Width=2778"
         Splits(1)._ColumnProps(579)=   "Column(144).DividerColor=0"
         Splits(1)._ColumnProps(580)=   "Column(144)._WidthInPix=2646"
         Splits(1)._ColumnProps(581)=   "Column(144).Order=145"
         Splits(1)._ColumnProps(582)=   "Column(145).Width=2778"
         Splits(1)._ColumnProps(583)=   "Column(145).DividerColor=0"
         Splits(1)._ColumnProps(584)=   "Column(145)._WidthInPix=2646"
         Splits(1)._ColumnProps(585)=   "Column(145).Order=146"
         Splits(1)._ColumnProps(586)=   "Column(146).Width=2778"
         Splits(1)._ColumnProps(587)=   "Column(146).DividerColor=0"
         Splits(1)._ColumnProps(588)=   "Column(146)._WidthInPix=2646"
         Splits(1)._ColumnProps(589)=   "Column(146).Order=147"
         Splits(1)._ColumnProps(590)=   "Column(147).Width=2778"
         Splits(1)._ColumnProps(591)=   "Column(147).DividerColor=0"
         Splits(1)._ColumnProps(592)=   "Column(147)._WidthInPix=2646"
         Splits(1)._ColumnProps(593)=   "Column(147).Order=148"
         Splits(1)._ColumnProps(594)=   "Column(148).Width=2778"
         Splits(1)._ColumnProps(595)=   "Column(148).DividerColor=0"
         Splits(1)._ColumnProps(596)=   "Column(148)._WidthInPix=2646"
         Splits(1)._ColumnProps(597)=   "Column(148).Order=149"
         Splits(1)._ColumnProps(598)=   "Column(149).Width=2778"
         Splits(1)._ColumnProps(599)=   "Column(149).DividerColor=0"
         Splits(1)._ColumnProps(600)=   "Column(149)._WidthInPix=2646"
         Splits(1)._ColumnProps(601)=   "Column(149).Order=150"
         Splits(1)._ColumnProps(602)=   "Column(150).Width=2778"
         Splits(1)._ColumnProps(603)=   "Column(150).DividerColor=0"
         Splits(1)._ColumnProps(604)=   "Column(150)._WidthInPix=2646"
         Splits(1)._ColumnProps(605)=   "Column(150).Order=151"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=14.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Georgia"
         PrintInfos(0).PageFooterFont=   "Size=14.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Georgia"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12577002
         RowDividerColor =   14869218
         RowSubDividerColor=   14869218
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=44,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1425,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Georgia"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=50,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=51,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=1266,.parent=43"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=1263,.parent=44"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=1264,.parent=45"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=1265,.parent=47"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=1258,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=1255,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=1256,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=1257,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=1250,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=1247,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=1248,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=1249,.parent=47"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=1242,.parent=43"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=1239,.parent=44"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=1240,.parent=45"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=1241,.parent=47"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=1234,.parent=43"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=1231,.parent=44"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=1232,.parent=45"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=1233,.parent=47"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=1226,.parent=43"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=1223,.parent=44"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=1224,.parent=45"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=1225,.parent=47"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=1218,.parent=43"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=1215,.parent=44"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=1216,.parent=45"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=1217,.parent=47"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=1210,.parent=43"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=1207,.parent=44"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=1208,.parent=45"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=1209,.parent=47"
         _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=1202,.parent=43"
         _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=1199,.parent=44"
         _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=1200,.parent=45"
         _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=1201,.parent=47"
         _StyleDefs(68)  =   "Splits(0).Columns(9).Style:id=1194,.parent=43"
         _StyleDefs(69)  =   "Splits(0).Columns(9).HeadingStyle:id=1191,.parent=44"
         _StyleDefs(70)  =   "Splits(0).Columns(9).FooterStyle:id=1192,.parent=45"
         _StyleDefs(71)  =   "Splits(0).Columns(9).EditorStyle:id=1193,.parent=47"
         _StyleDefs(72)  =   "Splits(0).Columns(10).Style:id=1186,.parent=43"
         _StyleDefs(73)  =   "Splits(0).Columns(10).HeadingStyle:id=1183,.parent=44"
         _StyleDefs(74)  =   "Splits(0).Columns(10).FooterStyle:id=1184,.parent=45"
         _StyleDefs(75)  =   "Splits(0).Columns(10).EditorStyle:id=1185,.parent=47"
         _StyleDefs(76)  =   "Splits(0).Columns(11).Style:id=1178,.parent=43"
         _StyleDefs(77)  =   "Splits(0).Columns(11).HeadingStyle:id=1175,.parent=44"
         _StyleDefs(78)  =   "Splits(0).Columns(11).FooterStyle:id=1176,.parent=45"
         _StyleDefs(79)  =   "Splits(0).Columns(11).EditorStyle:id=1177,.parent=47"
         _StyleDefs(80)  =   "Splits(0).Columns(12).Style:id=1170,.parent=43"
         _StyleDefs(81)  =   "Splits(0).Columns(12).HeadingStyle:id=1167,.parent=44"
         _StyleDefs(82)  =   "Splits(0).Columns(12).FooterStyle:id=1168,.parent=45"
         _StyleDefs(83)  =   "Splits(0).Columns(12).EditorStyle:id=1169,.parent=47"
         _StyleDefs(84)  =   "Splits(0).Columns(13).Style:id=1162,.parent=43"
         _StyleDefs(85)  =   "Splits(0).Columns(13).HeadingStyle:id=1159,.parent=44"
         _StyleDefs(86)  =   "Splits(0).Columns(13).FooterStyle:id=1160,.parent=45"
         _StyleDefs(87)  =   "Splits(0).Columns(13).EditorStyle:id=1161,.parent=47"
         _StyleDefs(88)  =   "Splits(0).Columns(14).Style:id=1154,.parent=43"
         _StyleDefs(89)  =   "Splits(0).Columns(14).HeadingStyle:id=1151,.parent=44"
         _StyleDefs(90)  =   "Splits(0).Columns(14).FooterStyle:id=1152,.parent=45"
         _StyleDefs(91)  =   "Splits(0).Columns(14).EditorStyle:id=1153,.parent=47"
         _StyleDefs(92)  =   "Splits(0).Columns(15).Style:id=1146,.parent=43"
         _StyleDefs(93)  =   "Splits(0).Columns(15).HeadingStyle:id=1143,.parent=44"
         _StyleDefs(94)  =   "Splits(0).Columns(15).FooterStyle:id=1144,.parent=45"
         _StyleDefs(95)  =   "Splits(0).Columns(15).EditorStyle:id=1145,.parent=47"
         _StyleDefs(96)  =   "Splits(0).Columns(16).Style:id=1138,.parent=43"
         _StyleDefs(97)  =   "Splits(0).Columns(16).HeadingStyle:id=1135,.parent=44"
         _StyleDefs(98)  =   "Splits(0).Columns(16).FooterStyle:id=1136,.parent=45"
         _StyleDefs(99)  =   "Splits(0).Columns(16).EditorStyle:id=1137,.parent=47"
         _StyleDefs(100) =   "Splits(0).Columns(17).Style:id=1130,.parent=43"
         _StyleDefs(101) =   "Splits(0).Columns(17).HeadingStyle:id=1127,.parent=44"
         _StyleDefs(102) =   "Splits(0).Columns(17).FooterStyle:id=1128,.parent=45"
         _StyleDefs(103) =   "Splits(0).Columns(17).EditorStyle:id=1129,.parent=47"
         _StyleDefs(104) =   "Splits(0).Columns(18).Style:id=1122,.parent=43"
         _StyleDefs(105) =   "Splits(0).Columns(18).HeadingStyle:id=1119,.parent=44"
         _StyleDefs(106) =   "Splits(0).Columns(18).FooterStyle:id=1120,.parent=45"
         _StyleDefs(107) =   "Splits(0).Columns(18).EditorStyle:id=1121,.parent=47"
         _StyleDefs(108) =   "Splits(0).Columns(19).Style:id=1114,.parent=43"
         _StyleDefs(109) =   "Splits(0).Columns(19).HeadingStyle:id=1111,.parent=44"
         _StyleDefs(110) =   "Splits(0).Columns(19).FooterStyle:id=1112,.parent=45"
         _StyleDefs(111) =   "Splits(0).Columns(19).EditorStyle:id=1113,.parent=47"
         _StyleDefs(112) =   "Splits(0).Columns(20).Style:id=1106,.parent=43"
         _StyleDefs(113) =   "Splits(0).Columns(20).HeadingStyle:id=1103,.parent=44"
         _StyleDefs(114) =   "Splits(0).Columns(20).FooterStyle:id=1104,.parent=45"
         _StyleDefs(115) =   "Splits(0).Columns(20).EditorStyle:id=1105,.parent=47"
         _StyleDefs(116) =   "Splits(0).Columns(21).Style:id=1098,.parent=43"
         _StyleDefs(117) =   "Splits(0).Columns(21).HeadingStyle:id=1095,.parent=44"
         _StyleDefs(118) =   "Splits(0).Columns(21).FooterStyle:id=1096,.parent=45"
         _StyleDefs(119) =   "Splits(0).Columns(21).EditorStyle:id=1097,.parent=47"
         _StyleDefs(120) =   "Splits(0).Columns(22).Style:id=1090,.parent=43"
         _StyleDefs(121) =   "Splits(0).Columns(22).HeadingStyle:id=1087,.parent=44"
         _StyleDefs(122) =   "Splits(0).Columns(22).FooterStyle:id=1088,.parent=45"
         _StyleDefs(123) =   "Splits(0).Columns(22).EditorStyle:id=1089,.parent=47"
         _StyleDefs(124) =   "Splits(0).Columns(23).Style:id=1082,.parent=43"
         _StyleDefs(125) =   "Splits(0).Columns(23).HeadingStyle:id=1079,.parent=44"
         _StyleDefs(126) =   "Splits(0).Columns(23).FooterStyle:id=1080,.parent=45"
         _StyleDefs(127) =   "Splits(0).Columns(23).EditorStyle:id=1081,.parent=47"
         _StyleDefs(128) =   "Splits(0).Columns(24).Style:id=1074,.parent=43"
         _StyleDefs(129) =   "Splits(0).Columns(24).HeadingStyle:id=1071,.parent=44"
         _StyleDefs(130) =   "Splits(0).Columns(24).FooterStyle:id=1072,.parent=45"
         _StyleDefs(131) =   "Splits(0).Columns(24).EditorStyle:id=1073,.parent=47"
         _StyleDefs(132) =   "Splits(0).Columns(25).Style:id=1066,.parent=43"
         _StyleDefs(133) =   "Splits(0).Columns(25).HeadingStyle:id=1063,.parent=44"
         _StyleDefs(134) =   "Splits(0).Columns(25).FooterStyle:id=1064,.parent=45"
         _StyleDefs(135) =   "Splits(0).Columns(25).EditorStyle:id=1065,.parent=47"
         _StyleDefs(136) =   "Splits(0).Columns(26).Style:id=1058,.parent=43"
         _StyleDefs(137) =   "Splits(0).Columns(26).HeadingStyle:id=1055,.parent=44"
         _StyleDefs(138) =   "Splits(0).Columns(26).FooterStyle:id=1056,.parent=45"
         _StyleDefs(139) =   "Splits(0).Columns(26).EditorStyle:id=1057,.parent=47"
         _StyleDefs(140) =   "Splits(0).Columns(27).Style:id=1050,.parent=43"
         _StyleDefs(141) =   "Splits(0).Columns(27).HeadingStyle:id=1047,.parent=44"
         _StyleDefs(142) =   "Splits(0).Columns(27).FooterStyle:id=1048,.parent=45"
         _StyleDefs(143) =   "Splits(0).Columns(27).EditorStyle:id=1049,.parent=47"
         _StyleDefs(144) =   "Splits(0).Columns(28).Style:id=1042,.parent=43"
         _StyleDefs(145) =   "Splits(0).Columns(28).HeadingStyle:id=1039,.parent=44"
         _StyleDefs(146) =   "Splits(0).Columns(28).FooterStyle:id=1040,.parent=45"
         _StyleDefs(147) =   "Splits(0).Columns(28).EditorStyle:id=1041,.parent=47"
         _StyleDefs(148) =   "Splits(0).Columns(29).Style:id=1034,.parent=43"
         _StyleDefs(149) =   "Splits(0).Columns(29).HeadingStyle:id=1031,.parent=44"
         _StyleDefs(150) =   "Splits(0).Columns(29).FooterStyle:id=1032,.parent=45"
         _StyleDefs(151) =   "Splits(0).Columns(29).EditorStyle:id=1033,.parent=47"
         _StyleDefs(152) =   "Splits(0).Columns(30).Style:id=1026,.parent=43"
         _StyleDefs(153) =   "Splits(0).Columns(30).HeadingStyle:id=1023,.parent=44"
         _StyleDefs(154) =   "Splits(0).Columns(30).FooterStyle:id=1024,.parent=45"
         _StyleDefs(155) =   "Splits(0).Columns(30).EditorStyle:id=1025,.parent=47"
         _StyleDefs(156) =   "Splits(0).Columns(31).Style:id=1018,.parent=43"
         _StyleDefs(157) =   "Splits(0).Columns(31).HeadingStyle:id=1015,.parent=44"
         _StyleDefs(158) =   "Splits(0).Columns(31).FooterStyle:id=1016,.parent=45"
         _StyleDefs(159) =   "Splits(0).Columns(31).EditorStyle:id=1017,.parent=47"
         _StyleDefs(160) =   "Splits(0).Columns(32).Style:id=1010,.parent=43"
         _StyleDefs(161) =   "Splits(0).Columns(32).HeadingStyle:id=1007,.parent=44"
         _StyleDefs(162) =   "Splits(0).Columns(32).FooterStyle:id=1008,.parent=45"
         _StyleDefs(163) =   "Splits(0).Columns(32).EditorStyle:id=1009,.parent=47"
         _StyleDefs(164) =   "Splits(0).Columns(33).Style:id=1002,.parent=43"
         _StyleDefs(165) =   "Splits(0).Columns(33).HeadingStyle:id=999,.parent=44"
         _StyleDefs(166) =   "Splits(0).Columns(33).FooterStyle:id=1000,.parent=45"
         _StyleDefs(167) =   "Splits(0).Columns(33).EditorStyle:id=1001,.parent=47"
         _StyleDefs(168) =   "Splits(0).Columns(34).Style:id=994,.parent=43"
         _StyleDefs(169) =   "Splits(0).Columns(34).HeadingStyle:id=991,.parent=44"
         _StyleDefs(170) =   "Splits(0).Columns(34).FooterStyle:id=992,.parent=45"
         _StyleDefs(171) =   "Splits(0).Columns(34).EditorStyle:id=993,.parent=47"
         _StyleDefs(172) =   "Splits(0).Columns(35).Style:id=986,.parent=43"
         _StyleDefs(173) =   "Splits(0).Columns(35).HeadingStyle:id=983,.parent=44"
         _StyleDefs(174) =   "Splits(0).Columns(35).FooterStyle:id=984,.parent=45"
         _StyleDefs(175) =   "Splits(0).Columns(35).EditorStyle:id=985,.parent=47"
         _StyleDefs(176) =   "Splits(0).Columns(36).Style:id=978,.parent=43"
         _StyleDefs(177) =   "Splits(0).Columns(36).HeadingStyle:id=975,.parent=44"
         _StyleDefs(178) =   "Splits(0).Columns(36).FooterStyle:id=976,.parent=45"
         _StyleDefs(179) =   "Splits(0).Columns(36).EditorStyle:id=977,.parent=47"
         _StyleDefs(180) =   "Splits(0).Columns(37).Style:id=970,.parent=43"
         _StyleDefs(181) =   "Splits(0).Columns(37).HeadingStyle:id=967,.parent=44"
         _StyleDefs(182) =   "Splits(0).Columns(37).FooterStyle:id=968,.parent=45"
         _StyleDefs(183) =   "Splits(0).Columns(37).EditorStyle:id=969,.parent=47"
         _StyleDefs(184) =   "Splits(0).Columns(38).Style:id=962,.parent=43"
         _StyleDefs(185) =   "Splits(0).Columns(38).HeadingStyle:id=959,.parent=44"
         _StyleDefs(186) =   "Splits(0).Columns(38).FooterStyle:id=960,.parent=45"
         _StyleDefs(187) =   "Splits(0).Columns(38).EditorStyle:id=961,.parent=47"
         _StyleDefs(188) =   "Splits(0).Columns(39).Style:id=954,.parent=43"
         _StyleDefs(189) =   "Splits(0).Columns(39).HeadingStyle:id=951,.parent=44"
         _StyleDefs(190) =   "Splits(0).Columns(39).FooterStyle:id=952,.parent=45"
         _StyleDefs(191) =   "Splits(0).Columns(39).EditorStyle:id=953,.parent=47"
         _StyleDefs(192) =   "Splits(0).Columns(40).Style:id=946,.parent=43"
         _StyleDefs(193) =   "Splits(0).Columns(40).HeadingStyle:id=943,.parent=44"
         _StyleDefs(194) =   "Splits(0).Columns(40).FooterStyle:id=944,.parent=45"
         _StyleDefs(195) =   "Splits(0).Columns(40).EditorStyle:id=945,.parent=47"
         _StyleDefs(196) =   "Splits(0).Columns(41).Style:id=938,.parent=43"
         _StyleDefs(197) =   "Splits(0).Columns(41).HeadingStyle:id=935,.parent=44"
         _StyleDefs(198) =   "Splits(0).Columns(41).FooterStyle:id=936,.parent=45"
         _StyleDefs(199) =   "Splits(0).Columns(41).EditorStyle:id=937,.parent=47"
         _StyleDefs(200) =   "Splits(0).Columns(42).Style:id=930,.parent=43"
         _StyleDefs(201) =   "Splits(0).Columns(42).HeadingStyle:id=927,.parent=44"
         _StyleDefs(202) =   "Splits(0).Columns(42).FooterStyle:id=928,.parent=45"
         _StyleDefs(203) =   "Splits(0).Columns(42).EditorStyle:id=929,.parent=47"
         _StyleDefs(204) =   "Splits(0).Columns(43).Style:id=922,.parent=43"
         _StyleDefs(205) =   "Splits(0).Columns(43).HeadingStyle:id=919,.parent=44"
         _StyleDefs(206) =   "Splits(0).Columns(43).FooterStyle:id=920,.parent=45"
         _StyleDefs(207) =   "Splits(0).Columns(43).EditorStyle:id=921,.parent=47"
         _StyleDefs(208) =   "Splits(0).Columns(44).Style:id=914,.parent=43"
         _StyleDefs(209) =   "Splits(0).Columns(44).HeadingStyle:id=911,.parent=44"
         _StyleDefs(210) =   "Splits(0).Columns(44).FooterStyle:id=912,.parent=45"
         _StyleDefs(211) =   "Splits(0).Columns(44).EditorStyle:id=913,.parent=47"
         _StyleDefs(212) =   "Splits(0).Columns(45).Style:id=906,.parent=43"
         _StyleDefs(213) =   "Splits(0).Columns(45).HeadingStyle:id=903,.parent=44"
         _StyleDefs(214) =   "Splits(0).Columns(45).FooterStyle:id=904,.parent=45"
         _StyleDefs(215) =   "Splits(0).Columns(45).EditorStyle:id=905,.parent=47"
         _StyleDefs(216) =   "Splits(0).Columns(46).Style:id=898,.parent=43"
         _StyleDefs(217) =   "Splits(0).Columns(46).HeadingStyle:id=895,.parent=44"
         _StyleDefs(218) =   "Splits(0).Columns(46).FooterStyle:id=896,.parent=45"
         _StyleDefs(219) =   "Splits(0).Columns(46).EditorStyle:id=897,.parent=47"
         _StyleDefs(220) =   "Splits(0).Columns(47).Style:id=890,.parent=43"
         _StyleDefs(221) =   "Splits(0).Columns(47).HeadingStyle:id=887,.parent=44"
         _StyleDefs(222) =   "Splits(0).Columns(47).FooterStyle:id=888,.parent=45"
         _StyleDefs(223) =   "Splits(0).Columns(47).EditorStyle:id=889,.parent=47"
         _StyleDefs(224) =   "Splits(0).Columns(48).Style:id=882,.parent=43"
         _StyleDefs(225) =   "Splits(0).Columns(48).HeadingStyle:id=879,.parent=44"
         _StyleDefs(226) =   "Splits(0).Columns(48).FooterStyle:id=880,.parent=45"
         _StyleDefs(227) =   "Splits(0).Columns(48).EditorStyle:id=881,.parent=47"
         _StyleDefs(228) =   "Splits(0).Columns(49).Style:id=874,.parent=43"
         _StyleDefs(229) =   "Splits(0).Columns(49).HeadingStyle:id=871,.parent=44"
         _StyleDefs(230) =   "Splits(0).Columns(49).FooterStyle:id=872,.parent=45"
         _StyleDefs(231) =   "Splits(0).Columns(49).EditorStyle:id=873,.parent=47"
         _StyleDefs(232) =   "Splits(0).Columns(50).Style:id=866,.parent=43"
         _StyleDefs(233) =   "Splits(0).Columns(50).HeadingStyle:id=863,.parent=44"
         _StyleDefs(234) =   "Splits(0).Columns(50).FooterStyle:id=864,.parent=45"
         _StyleDefs(235) =   "Splits(0).Columns(50).EditorStyle:id=865,.parent=47"
         _StyleDefs(236) =   "Splits(0).Columns(51).Style:id=858,.parent=43"
         _StyleDefs(237) =   "Splits(0).Columns(51).HeadingStyle:id=855,.parent=44"
         _StyleDefs(238) =   "Splits(0).Columns(51).FooterStyle:id=856,.parent=45"
         _StyleDefs(239) =   "Splits(0).Columns(51).EditorStyle:id=857,.parent=47"
         _StyleDefs(240) =   "Splits(0).Columns(52).Style:id=850,.parent=43"
         _StyleDefs(241) =   "Splits(0).Columns(52).HeadingStyle:id=847,.parent=44"
         _StyleDefs(242) =   "Splits(0).Columns(52).FooterStyle:id=848,.parent=45"
         _StyleDefs(243) =   "Splits(0).Columns(52).EditorStyle:id=849,.parent=47"
         _StyleDefs(244) =   "Splits(0).Columns(53).Style:id=842,.parent=43"
         _StyleDefs(245) =   "Splits(0).Columns(53).HeadingStyle:id=839,.parent=44"
         _StyleDefs(246) =   "Splits(0).Columns(53).FooterStyle:id=840,.parent=45"
         _StyleDefs(247) =   "Splits(0).Columns(53).EditorStyle:id=841,.parent=47"
         _StyleDefs(248) =   "Splits(0).Columns(54).Style:id=834,.parent=43"
         _StyleDefs(249) =   "Splits(0).Columns(54).HeadingStyle:id=831,.parent=44"
         _StyleDefs(250) =   "Splits(0).Columns(54).FooterStyle:id=832,.parent=45"
         _StyleDefs(251) =   "Splits(0).Columns(54).EditorStyle:id=833,.parent=47"
         _StyleDefs(252) =   "Splits(0).Columns(55).Style:id=826,.parent=43"
         _StyleDefs(253) =   "Splits(0).Columns(55).HeadingStyle:id=823,.parent=44"
         _StyleDefs(254) =   "Splits(0).Columns(55).FooterStyle:id=824,.parent=45"
         _StyleDefs(255) =   "Splits(0).Columns(55).EditorStyle:id=825,.parent=47"
         _StyleDefs(256) =   "Splits(0).Columns(56).Style:id=818,.parent=43"
         _StyleDefs(257) =   "Splits(0).Columns(56).HeadingStyle:id=815,.parent=44"
         _StyleDefs(258) =   "Splits(0).Columns(56).FooterStyle:id=816,.parent=45"
         _StyleDefs(259) =   "Splits(0).Columns(56).EditorStyle:id=817,.parent=47"
         _StyleDefs(260) =   "Splits(0).Columns(57).Style:id=810,.parent=43"
         _StyleDefs(261) =   "Splits(0).Columns(57).HeadingStyle:id=807,.parent=44"
         _StyleDefs(262) =   "Splits(0).Columns(57).FooterStyle:id=808,.parent=45"
         _StyleDefs(263) =   "Splits(0).Columns(57).EditorStyle:id=809,.parent=47"
         _StyleDefs(264) =   "Splits(0).Columns(58).Style:id=802,.parent=43"
         _StyleDefs(265) =   "Splits(0).Columns(58).HeadingStyle:id=799,.parent=44"
         _StyleDefs(266) =   "Splits(0).Columns(58).FooterStyle:id=800,.parent=45"
         _StyleDefs(267) =   "Splits(0).Columns(58).EditorStyle:id=801,.parent=47"
         _StyleDefs(268) =   "Splits(0).Columns(59).Style:id=794,.parent=43"
         _StyleDefs(269) =   "Splits(0).Columns(59).HeadingStyle:id=791,.parent=44"
         _StyleDefs(270) =   "Splits(0).Columns(59).FooterStyle:id=792,.parent=45"
         _StyleDefs(271) =   "Splits(0).Columns(59).EditorStyle:id=793,.parent=47"
         _StyleDefs(272) =   "Splits(0).Columns(60).Style:id=786,.parent=43"
         _StyleDefs(273) =   "Splits(0).Columns(60).HeadingStyle:id=783,.parent=44"
         _StyleDefs(274) =   "Splits(0).Columns(60).FooterStyle:id=784,.parent=45"
         _StyleDefs(275) =   "Splits(0).Columns(60).EditorStyle:id=785,.parent=47"
         _StyleDefs(276) =   "Splits(0).Columns(61).Style:id=778,.parent=43"
         _StyleDefs(277) =   "Splits(0).Columns(61).HeadingStyle:id=775,.parent=44"
         _StyleDefs(278) =   "Splits(0).Columns(61).FooterStyle:id=776,.parent=45"
         _StyleDefs(279) =   "Splits(0).Columns(61).EditorStyle:id=777,.parent=47"
         _StyleDefs(280) =   "Splits(0).Columns(62).Style:id=770,.parent=43"
         _StyleDefs(281) =   "Splits(0).Columns(62).HeadingStyle:id=767,.parent=44"
         _StyleDefs(282) =   "Splits(0).Columns(62).FooterStyle:id=768,.parent=45"
         _StyleDefs(283) =   "Splits(0).Columns(62).EditorStyle:id=769,.parent=47"
         _StyleDefs(284) =   "Splits(0).Columns(63).Style:id=762,.parent=43"
         _StyleDefs(285) =   "Splits(0).Columns(63).HeadingStyle:id=759,.parent=44"
         _StyleDefs(286) =   "Splits(0).Columns(63).FooterStyle:id=760,.parent=45"
         _StyleDefs(287) =   "Splits(0).Columns(63).EditorStyle:id=761,.parent=47"
         _StyleDefs(288) =   "Splits(0).Columns(64).Style:id=754,.parent=43"
         _StyleDefs(289) =   "Splits(0).Columns(64).HeadingStyle:id=751,.parent=44"
         _StyleDefs(290) =   "Splits(0).Columns(64).FooterStyle:id=752,.parent=45"
         _StyleDefs(291) =   "Splits(0).Columns(64).EditorStyle:id=753,.parent=47"
         _StyleDefs(292) =   "Splits(0).Columns(65).Style:id=746,.parent=43"
         _StyleDefs(293) =   "Splits(0).Columns(65).HeadingStyle:id=743,.parent=44"
         _StyleDefs(294) =   "Splits(0).Columns(65).FooterStyle:id=744,.parent=45"
         _StyleDefs(295) =   "Splits(0).Columns(65).EditorStyle:id=745,.parent=47"
         _StyleDefs(296) =   "Splits(0).Columns(66).Style:id=738,.parent=43"
         _StyleDefs(297) =   "Splits(0).Columns(66).HeadingStyle:id=735,.parent=44"
         _StyleDefs(298) =   "Splits(0).Columns(66).FooterStyle:id=736,.parent=45"
         _StyleDefs(299) =   "Splits(0).Columns(66).EditorStyle:id=737,.parent=47"
         _StyleDefs(300) =   "Splits(0).Columns(67).Style:id=730,.parent=43"
         _StyleDefs(301) =   "Splits(0).Columns(67).HeadingStyle:id=727,.parent=44"
         _StyleDefs(302) =   "Splits(0).Columns(67).FooterStyle:id=728,.parent=45"
         _StyleDefs(303) =   "Splits(0).Columns(67).EditorStyle:id=729,.parent=47"
         _StyleDefs(304) =   "Splits(0).Columns(68).Style:id=722,.parent=43"
         _StyleDefs(305) =   "Splits(0).Columns(68).HeadingStyle:id=719,.parent=44"
         _StyleDefs(306) =   "Splits(0).Columns(68).FooterStyle:id=720,.parent=45"
         _StyleDefs(307) =   "Splits(0).Columns(68).EditorStyle:id=721,.parent=47"
         _StyleDefs(308) =   "Splits(0).Columns(69).Style:id=714,.parent=43"
         _StyleDefs(309) =   "Splits(0).Columns(69).HeadingStyle:id=711,.parent=44"
         _StyleDefs(310) =   "Splits(0).Columns(69).FooterStyle:id=712,.parent=45"
         _StyleDefs(311) =   "Splits(0).Columns(69).EditorStyle:id=713,.parent=47"
         _StyleDefs(312) =   "Splits(0).Columns(70).Style:id=706,.parent=43"
         _StyleDefs(313) =   "Splits(0).Columns(70).HeadingStyle:id=703,.parent=44"
         _StyleDefs(314) =   "Splits(0).Columns(70).FooterStyle:id=704,.parent=45"
         _StyleDefs(315) =   "Splits(0).Columns(70).EditorStyle:id=705,.parent=47"
         _StyleDefs(316) =   "Splits(0).Columns(71).Style:id=698,.parent=43"
         _StyleDefs(317) =   "Splits(0).Columns(71).HeadingStyle:id=695,.parent=44"
         _StyleDefs(318) =   "Splits(0).Columns(71).FooterStyle:id=696,.parent=45"
         _StyleDefs(319) =   "Splits(0).Columns(71).EditorStyle:id=697,.parent=47"
         _StyleDefs(320) =   "Splits(0).Columns(72).Style:id=690,.parent=43"
         _StyleDefs(321) =   "Splits(0).Columns(72).HeadingStyle:id=687,.parent=44"
         _StyleDefs(322) =   "Splits(0).Columns(72).FooterStyle:id=688,.parent=45"
         _StyleDefs(323) =   "Splits(0).Columns(72).EditorStyle:id=689,.parent=47"
         _StyleDefs(324) =   "Splits(0).Columns(73).Style:id=682,.parent=43"
         _StyleDefs(325) =   "Splits(0).Columns(73).HeadingStyle:id=679,.parent=44"
         _StyleDefs(326) =   "Splits(0).Columns(73).FooterStyle:id=680,.parent=45"
         _StyleDefs(327) =   "Splits(0).Columns(73).EditorStyle:id=681,.parent=47"
         _StyleDefs(328) =   "Splits(0).Columns(74).Style:id=674,.parent=43"
         _StyleDefs(329) =   "Splits(0).Columns(74).HeadingStyle:id=671,.parent=44"
         _StyleDefs(330) =   "Splits(0).Columns(74).FooterStyle:id=672,.parent=45"
         _StyleDefs(331) =   "Splits(0).Columns(74).EditorStyle:id=673,.parent=47"
         _StyleDefs(332) =   "Splits(0).Columns(75).Style:id=666,.parent=43"
         _StyleDefs(333) =   "Splits(0).Columns(75).HeadingStyle:id=663,.parent=44"
         _StyleDefs(334) =   "Splits(0).Columns(75).FooterStyle:id=664,.parent=45"
         _StyleDefs(335) =   "Splits(0).Columns(75).EditorStyle:id=665,.parent=47"
         _StyleDefs(336) =   "Splits(0).Columns(76).Style:id=658,.parent=43"
         _StyleDefs(337) =   "Splits(0).Columns(76).HeadingStyle:id=655,.parent=44"
         _StyleDefs(338) =   "Splits(0).Columns(76).FooterStyle:id=656,.parent=45"
         _StyleDefs(339) =   "Splits(0).Columns(76).EditorStyle:id=657,.parent=47"
         _StyleDefs(340) =   "Splits(0).Columns(77).Style:id=650,.parent=43"
         _StyleDefs(341) =   "Splits(0).Columns(77).HeadingStyle:id=647,.parent=44"
         _StyleDefs(342) =   "Splits(0).Columns(77).FooterStyle:id=648,.parent=45"
         _StyleDefs(343) =   "Splits(0).Columns(77).EditorStyle:id=649,.parent=47"
         _StyleDefs(344) =   "Splits(0).Columns(78).Style:id=642,.parent=43"
         _StyleDefs(345) =   "Splits(0).Columns(78).HeadingStyle:id=639,.parent=44"
         _StyleDefs(346) =   "Splits(0).Columns(78).FooterStyle:id=640,.parent=45"
         _StyleDefs(347) =   "Splits(0).Columns(78).EditorStyle:id=641,.parent=47"
         _StyleDefs(348) =   "Splits(0).Columns(79).Style:id=634,.parent=43"
         _StyleDefs(349) =   "Splits(0).Columns(79).HeadingStyle:id=631,.parent=44"
         _StyleDefs(350) =   "Splits(0).Columns(79).FooterStyle:id=632,.parent=45"
         _StyleDefs(351) =   "Splits(0).Columns(79).EditorStyle:id=633,.parent=47"
         _StyleDefs(352) =   "Splits(0).Columns(80).Style:id=626,.parent=43"
         _StyleDefs(353) =   "Splits(0).Columns(80).HeadingStyle:id=623,.parent=44"
         _StyleDefs(354) =   "Splits(0).Columns(80).FooterStyle:id=624,.parent=45"
         _StyleDefs(355) =   "Splits(0).Columns(80).EditorStyle:id=625,.parent=47"
         _StyleDefs(356) =   "Splits(0).Columns(81).Style:id=618,.parent=43"
         _StyleDefs(357) =   "Splits(0).Columns(81).HeadingStyle:id=615,.parent=44"
         _StyleDefs(358) =   "Splits(0).Columns(81).FooterStyle:id=616,.parent=45"
         _StyleDefs(359) =   "Splits(0).Columns(81).EditorStyle:id=617,.parent=47"
         _StyleDefs(360) =   "Splits(0).Columns(82).Style:id=58,.parent=43"
         _StyleDefs(361) =   "Splits(0).Columns(82).HeadingStyle:id=55,.parent=44"
         _StyleDefs(362) =   "Splits(0).Columns(82).FooterStyle:id=56,.parent=45"
         _StyleDefs(363) =   "Splits(0).Columns(82).EditorStyle:id=57,.parent=47"
         _StyleDefs(364) =   "Splits(0).Columns(83).Style:id=62,.parent=43"
         _StyleDefs(365) =   "Splits(0).Columns(83).HeadingStyle:id=59,.parent=44"
         _StyleDefs(366) =   "Splits(0).Columns(83).FooterStyle:id=60,.parent=45"
         _StyleDefs(367) =   "Splits(0).Columns(83).EditorStyle:id=61,.parent=47"
         _StyleDefs(368) =   "Splits(0).Columns(84).Style:id=66,.parent=43"
         _StyleDefs(369) =   "Splits(0).Columns(84).HeadingStyle:id=63,.parent=44"
         _StyleDefs(370) =   "Splits(0).Columns(84).FooterStyle:id=64,.parent=45"
         _StyleDefs(371) =   "Splits(0).Columns(84).EditorStyle:id=65,.parent=47"
         _StyleDefs(372) =   "Splits(0).Columns(85).Style:id=70,.parent=43"
         _StyleDefs(373) =   "Splits(0).Columns(85).HeadingStyle:id=67,.parent=44"
         _StyleDefs(374) =   "Splits(0).Columns(85).FooterStyle:id=68,.parent=45"
         _StyleDefs(375) =   "Splits(0).Columns(85).EditorStyle:id=69,.parent=47"
         _StyleDefs(376) =   "Splits(0).Columns(86).Style:id=74,.parent=43"
         _StyleDefs(377) =   "Splits(0).Columns(86).HeadingStyle:id=71,.parent=44"
         _StyleDefs(378) =   "Splits(0).Columns(86).FooterStyle:id=72,.parent=45"
         _StyleDefs(379) =   "Splits(0).Columns(86).EditorStyle:id=73,.parent=47"
         _StyleDefs(380) =   "Splits(0).Columns(87).Style:id=78,.parent=43"
         _StyleDefs(381) =   "Splits(0).Columns(87).HeadingStyle:id=75,.parent=44"
         _StyleDefs(382) =   "Splits(0).Columns(87).FooterStyle:id=76,.parent=45"
         _StyleDefs(383) =   "Splits(0).Columns(87).EditorStyle:id=77,.parent=47"
         _StyleDefs(384) =   "Splits(0).Columns(88).Style:id=82,.parent=43"
         _StyleDefs(385) =   "Splits(0).Columns(88).HeadingStyle:id=79,.parent=44"
         _StyleDefs(386) =   "Splits(0).Columns(88).FooterStyle:id=80,.parent=45"
         _StyleDefs(387) =   "Splits(0).Columns(88).EditorStyle:id=81,.parent=47"
         _StyleDefs(388) =   "Splits(0).Columns(89).Style:id=86,.parent=43"
         _StyleDefs(389) =   "Splits(0).Columns(89).HeadingStyle:id=83,.parent=44"
         _StyleDefs(390) =   "Splits(0).Columns(89).FooterStyle:id=84,.parent=45"
         _StyleDefs(391) =   "Splits(0).Columns(89).EditorStyle:id=85,.parent=47"
         _StyleDefs(392) =   "Splits(0).Columns(90).Style:id=90,.parent=43"
         _StyleDefs(393) =   "Splits(0).Columns(90).HeadingStyle:id=87,.parent=44"
         _StyleDefs(394) =   "Splits(0).Columns(90).FooterStyle:id=88,.parent=45"
         _StyleDefs(395) =   "Splits(0).Columns(90).EditorStyle:id=89,.parent=47"
         _StyleDefs(396) =   "Splits(0).Columns(91).Style:id=94,.parent=43"
         _StyleDefs(397) =   "Splits(0).Columns(91).HeadingStyle:id=91,.parent=44"
         _StyleDefs(398) =   "Splits(0).Columns(91).FooterStyle:id=92,.parent=45"
         _StyleDefs(399) =   "Splits(0).Columns(91).EditorStyle:id=93,.parent=47"
         _StyleDefs(400) =   "Splits(0).Columns(92).Style:id=98,.parent=43"
         _StyleDefs(401) =   "Splits(0).Columns(92).HeadingStyle:id=95,.parent=44"
         _StyleDefs(402) =   "Splits(0).Columns(92).FooterStyle:id=96,.parent=45"
         _StyleDefs(403) =   "Splits(0).Columns(92).EditorStyle:id=97,.parent=47"
         _StyleDefs(404) =   "Splits(0).Columns(93).Style:id=102,.parent=43"
         _StyleDefs(405) =   "Splits(0).Columns(93).HeadingStyle:id=99,.parent=44"
         _StyleDefs(406) =   "Splits(0).Columns(93).FooterStyle:id=100,.parent=45"
         _StyleDefs(407) =   "Splits(0).Columns(93).EditorStyle:id=101,.parent=47"
         _StyleDefs(408) =   "Splits(0).Columns(94).Style:id=106,.parent=43"
         _StyleDefs(409) =   "Splits(0).Columns(94).HeadingStyle:id=103,.parent=44"
         _StyleDefs(410) =   "Splits(0).Columns(94).FooterStyle:id=104,.parent=45"
         _StyleDefs(411) =   "Splits(0).Columns(94).EditorStyle:id=105,.parent=47"
         _StyleDefs(412) =   "Splits(0).Columns(95).Style:id=110,.parent=43"
         _StyleDefs(413) =   "Splits(0).Columns(95).HeadingStyle:id=107,.parent=44"
         _StyleDefs(414) =   "Splits(0).Columns(95).FooterStyle:id=108,.parent=45"
         _StyleDefs(415) =   "Splits(0).Columns(95).EditorStyle:id=109,.parent=47"
         _StyleDefs(416) =   "Splits(0).Columns(96).Style:id=310,.parent=43"
         _StyleDefs(417) =   "Splits(0).Columns(96).HeadingStyle:id=307,.parent=44"
         _StyleDefs(418) =   "Splits(0).Columns(96).FooterStyle:id=308,.parent=45"
         _StyleDefs(419) =   "Splits(0).Columns(96).EditorStyle:id=309,.parent=47"
         _StyleDefs(420) =   "Splits(0).Columns(97).Style:id=394,.parent=43"
         _StyleDefs(421) =   "Splits(0).Columns(97).HeadingStyle:id=391,.parent=44"
         _StyleDefs(422) =   "Splits(0).Columns(97).FooterStyle:id=392,.parent=45"
         _StyleDefs(423) =   "Splits(0).Columns(97).EditorStyle:id=393,.parent=47"
         _StyleDefs(424) =   "Splits(0).Columns(98).Style:id=398,.parent=43"
         _StyleDefs(425) =   "Splits(0).Columns(98).HeadingStyle:id=395,.parent=44"
         _StyleDefs(426) =   "Splits(0).Columns(98).FooterStyle:id=396,.parent=45"
         _StyleDefs(427) =   "Splits(0).Columns(98).EditorStyle:id=397,.parent=47"
         _StyleDefs(428) =   "Splits(0).Columns(99).Style:id=402,.parent=43"
         _StyleDefs(429) =   "Splits(0).Columns(99).HeadingStyle:id=399,.parent=44"
         _StyleDefs(430) =   "Splits(0).Columns(99).FooterStyle:id=400,.parent=45"
         _StyleDefs(431) =   "Splits(0).Columns(99).EditorStyle:id=401,.parent=47"
         _StyleDefs(432) =   "Splits(0).Columns(100).Style:id=406,.parent=43"
         _StyleDefs(433) =   "Splits(0).Columns(100).HeadingStyle:id=403,.parent=44"
         _StyleDefs(434) =   "Splits(0).Columns(100).FooterStyle:id=404,.parent=45"
         _StyleDefs(435) =   "Splits(0).Columns(100).EditorStyle:id=405,.parent=47"
         _StyleDefs(436) =   "Splits(0).Columns(101).Style:id=410,.parent=43"
         _StyleDefs(437) =   "Splits(0).Columns(101).HeadingStyle:id=407,.parent=44"
         _StyleDefs(438) =   "Splits(0).Columns(101).FooterStyle:id=408,.parent=45"
         _StyleDefs(439) =   "Splits(0).Columns(101).EditorStyle:id=409,.parent=47"
         _StyleDefs(440) =   "Splits(0).Columns(102).Style:id=414,.parent=43"
         _StyleDefs(441) =   "Splits(0).Columns(102).HeadingStyle:id=411,.parent=44"
         _StyleDefs(442) =   "Splits(0).Columns(102).FooterStyle:id=412,.parent=45"
         _StyleDefs(443) =   "Splits(0).Columns(102).EditorStyle:id=413,.parent=47"
         _StyleDefs(444) =   "Splits(0).Columns(103).Style:id=418,.parent=43"
         _StyleDefs(445) =   "Splits(0).Columns(103).HeadingStyle:id=415,.parent=44"
         _StyleDefs(446) =   "Splits(0).Columns(103).FooterStyle:id=416,.parent=45"
         _StyleDefs(447) =   "Splits(0).Columns(103).EditorStyle:id=417,.parent=47"
         _StyleDefs(448) =   "Splits(0).Columns(104).Style:id=422,.parent=43"
         _StyleDefs(449) =   "Splits(0).Columns(104).HeadingStyle:id=419,.parent=44"
         _StyleDefs(450) =   "Splits(0).Columns(104).FooterStyle:id=420,.parent=45"
         _StyleDefs(451) =   "Splits(0).Columns(104).EditorStyle:id=421,.parent=47"
         _StyleDefs(452) =   "Splits(0).Columns(105).Style:id=426,.parent=43"
         _StyleDefs(453) =   "Splits(0).Columns(105).HeadingStyle:id=423,.parent=44"
         _StyleDefs(454) =   "Splits(0).Columns(105).FooterStyle:id=424,.parent=45"
         _StyleDefs(455) =   "Splits(0).Columns(105).EditorStyle:id=425,.parent=47"
         _StyleDefs(456) =   "Splits(0).Columns(106).Style:id=430,.parent=43"
         _StyleDefs(457) =   "Splits(0).Columns(106).HeadingStyle:id=427,.parent=44"
         _StyleDefs(458) =   "Splits(0).Columns(106).FooterStyle:id=428,.parent=45"
         _StyleDefs(459) =   "Splits(0).Columns(106).EditorStyle:id=429,.parent=47"
         _StyleDefs(460) =   "Splits(0).Columns(107).Style:id=434,.parent=43"
         _StyleDefs(461) =   "Splits(0).Columns(107).HeadingStyle:id=431,.parent=44"
         _StyleDefs(462) =   "Splits(0).Columns(107).FooterStyle:id=432,.parent=45"
         _StyleDefs(463) =   "Splits(0).Columns(107).EditorStyle:id=433,.parent=47"
         _StyleDefs(464) =   "Splits(0).Columns(108).Style:id=438,.parent=43"
         _StyleDefs(465) =   "Splits(0).Columns(108).HeadingStyle:id=435,.parent=44"
         _StyleDefs(466) =   "Splits(0).Columns(108).FooterStyle:id=436,.parent=45"
         _StyleDefs(467) =   "Splits(0).Columns(108).EditorStyle:id=437,.parent=47"
         _StyleDefs(468) =   "Splits(0).Columns(109).Style:id=442,.parent=43"
         _StyleDefs(469) =   "Splits(0).Columns(109).HeadingStyle:id=439,.parent=44"
         _StyleDefs(470) =   "Splits(0).Columns(109).FooterStyle:id=440,.parent=45"
         _StyleDefs(471) =   "Splits(0).Columns(109).EditorStyle:id=441,.parent=47"
         _StyleDefs(472) =   "Splits(0).Columns(110).Style:id=446,.parent=43"
         _StyleDefs(473) =   "Splits(0).Columns(110).HeadingStyle:id=443,.parent=44"
         _StyleDefs(474) =   "Splits(0).Columns(110).FooterStyle:id=444,.parent=45"
         _StyleDefs(475) =   "Splits(0).Columns(110).EditorStyle:id=445,.parent=47"
         _StyleDefs(476) =   "Splits(0).Columns(111).Style:id=450,.parent=43"
         _StyleDefs(477) =   "Splits(0).Columns(111).HeadingStyle:id=447,.parent=44"
         _StyleDefs(478) =   "Splits(0).Columns(111).FooterStyle:id=448,.parent=45"
         _StyleDefs(479) =   "Splits(0).Columns(111).EditorStyle:id=449,.parent=47"
         _StyleDefs(480) =   "Splits(0).Columns(112).Style:id=454,.parent=43"
         _StyleDefs(481) =   "Splits(0).Columns(112).HeadingStyle:id=451,.parent=44"
         _StyleDefs(482) =   "Splits(0).Columns(112).FooterStyle:id=452,.parent=45"
         _StyleDefs(483) =   "Splits(0).Columns(112).EditorStyle:id=453,.parent=47"
         _StyleDefs(484) =   "Splits(0).Columns(113).Style:id=458,.parent=43"
         _StyleDefs(485) =   "Splits(0).Columns(113).HeadingStyle:id=455,.parent=44"
         _StyleDefs(486) =   "Splits(0).Columns(113).FooterStyle:id=456,.parent=45"
         _StyleDefs(487) =   "Splits(0).Columns(113).EditorStyle:id=457,.parent=47"
         _StyleDefs(488) =   "Splits(0).Columns(114).Style:id=462,.parent=43"
         _StyleDefs(489) =   "Splits(0).Columns(114).HeadingStyle:id=459,.parent=44"
         _StyleDefs(490) =   "Splits(0).Columns(114).FooterStyle:id=460,.parent=45"
         _StyleDefs(491) =   "Splits(0).Columns(114).EditorStyle:id=461,.parent=47"
         _StyleDefs(492) =   "Splits(0).Columns(115).Style:id=466,.parent=43"
         _StyleDefs(493) =   "Splits(0).Columns(115).HeadingStyle:id=463,.parent=44"
         _StyleDefs(494) =   "Splits(0).Columns(115).FooterStyle:id=464,.parent=45"
         _StyleDefs(495) =   "Splits(0).Columns(115).EditorStyle:id=465,.parent=47"
         _StyleDefs(496) =   "Splits(0).Columns(116).Style:id=470,.parent=43"
         _StyleDefs(497) =   "Splits(0).Columns(116).HeadingStyle:id=467,.parent=44"
         _StyleDefs(498) =   "Splits(0).Columns(116).FooterStyle:id=468,.parent=45"
         _StyleDefs(499) =   "Splits(0).Columns(116).EditorStyle:id=469,.parent=47"
         _StyleDefs(500) =   "Splits(0).Columns(117).Style:id=474,.parent=43"
         _StyleDefs(501) =   "Splits(0).Columns(117).HeadingStyle:id=471,.parent=44"
         _StyleDefs(502) =   "Splits(0).Columns(117).FooterStyle:id=472,.parent=45"
         _StyleDefs(503) =   "Splits(0).Columns(117).EditorStyle:id=473,.parent=47"
         _StyleDefs(504) =   "Splits(0).Columns(118).Style:id=478,.parent=43"
         _StyleDefs(505) =   "Splits(0).Columns(118).HeadingStyle:id=475,.parent=44"
         _StyleDefs(506) =   "Splits(0).Columns(118).FooterStyle:id=476,.parent=45"
         _StyleDefs(507) =   "Splits(0).Columns(118).EditorStyle:id=477,.parent=47"
         _StyleDefs(508) =   "Splits(0).Columns(119).Style:id=482,.parent=43"
         _StyleDefs(509) =   "Splits(0).Columns(119).HeadingStyle:id=479,.parent=44"
         _StyleDefs(510) =   "Splits(0).Columns(119).FooterStyle:id=480,.parent=45"
         _StyleDefs(511) =   "Splits(0).Columns(119).EditorStyle:id=481,.parent=47"
         _StyleDefs(512) =   "Splits(0).Columns(120).Style:id=486,.parent=43"
         _StyleDefs(513) =   "Splits(0).Columns(120).HeadingStyle:id=483,.parent=44"
         _StyleDefs(514) =   "Splits(0).Columns(120).FooterStyle:id=484,.parent=45"
         _StyleDefs(515) =   "Splits(0).Columns(120).EditorStyle:id=485,.parent=47"
         _StyleDefs(516) =   "Splits(0).Columns(121).Style:id=490,.parent=43"
         _StyleDefs(517) =   "Splits(0).Columns(121).HeadingStyle:id=487,.parent=44"
         _StyleDefs(518) =   "Splits(0).Columns(121).FooterStyle:id=488,.parent=45"
         _StyleDefs(519) =   "Splits(0).Columns(121).EditorStyle:id=489,.parent=47"
         _StyleDefs(520) =   "Splits(0).Columns(122).Style:id=494,.parent=43"
         _StyleDefs(521) =   "Splits(0).Columns(122).HeadingStyle:id=491,.parent=44"
         _StyleDefs(522) =   "Splits(0).Columns(122).FooterStyle:id=492,.parent=45"
         _StyleDefs(523) =   "Splits(0).Columns(122).EditorStyle:id=493,.parent=47"
         _StyleDefs(524) =   "Splits(0).Columns(123).Style:id=498,.parent=43"
         _StyleDefs(525) =   "Splits(0).Columns(123).HeadingStyle:id=495,.parent=44"
         _StyleDefs(526) =   "Splits(0).Columns(123).FooterStyle:id=496,.parent=45"
         _StyleDefs(527) =   "Splits(0).Columns(123).EditorStyle:id=497,.parent=47"
         _StyleDefs(528) =   "Splits(0).Columns(124).Style:id=502,.parent=43"
         _StyleDefs(529) =   "Splits(0).Columns(124).HeadingStyle:id=499,.parent=44"
         _StyleDefs(530) =   "Splits(0).Columns(124).FooterStyle:id=500,.parent=45"
         _StyleDefs(531) =   "Splits(0).Columns(124).EditorStyle:id=501,.parent=47"
         _StyleDefs(532) =   "Splits(0).Columns(125).Style:id=506,.parent=43"
         _StyleDefs(533) =   "Splits(0).Columns(125).HeadingStyle:id=503,.parent=44"
         _StyleDefs(534) =   "Splits(0).Columns(125).FooterStyle:id=504,.parent=45"
         _StyleDefs(535) =   "Splits(0).Columns(125).EditorStyle:id=505,.parent=47"
         _StyleDefs(536) =   "Splits(0).Columns(126).Style:id=510,.parent=43"
         _StyleDefs(537) =   "Splits(0).Columns(126).HeadingStyle:id=507,.parent=44"
         _StyleDefs(538) =   "Splits(0).Columns(126).FooterStyle:id=508,.parent=45"
         _StyleDefs(539) =   "Splits(0).Columns(126).EditorStyle:id=509,.parent=47"
         _StyleDefs(540) =   "Splits(0).Columns(127).Style:id=514,.parent=43"
         _StyleDefs(541) =   "Splits(0).Columns(127).HeadingStyle:id=511,.parent=44"
         _StyleDefs(542) =   "Splits(0).Columns(127).FooterStyle:id=512,.parent=45"
         _StyleDefs(543) =   "Splits(0).Columns(127).EditorStyle:id=513,.parent=47"
         _StyleDefs(544) =   "Splits(0).Columns(128).Style:id=518,.parent=43"
         _StyleDefs(545) =   "Splits(0).Columns(128).HeadingStyle:id=515,.parent=44"
         _StyleDefs(546) =   "Splits(0).Columns(128).FooterStyle:id=516,.parent=45"
         _StyleDefs(547) =   "Splits(0).Columns(128).EditorStyle:id=517,.parent=47"
         _StyleDefs(548) =   "Splits(0).Columns(129).Style:id=522,.parent=43"
         _StyleDefs(549) =   "Splits(0).Columns(129).HeadingStyle:id=519,.parent=44"
         _StyleDefs(550) =   "Splits(0).Columns(129).FooterStyle:id=520,.parent=45"
         _StyleDefs(551) =   "Splits(0).Columns(129).EditorStyle:id=521,.parent=47"
         _StyleDefs(552) =   "Splits(0).Columns(130).Style:id=526,.parent=43"
         _StyleDefs(553) =   "Splits(0).Columns(130).HeadingStyle:id=523,.parent=44"
         _StyleDefs(554) =   "Splits(0).Columns(130).FooterStyle:id=524,.parent=45"
         _StyleDefs(555) =   "Splits(0).Columns(130).EditorStyle:id=525,.parent=47"
         _StyleDefs(556) =   "Splits(0).Columns(131).Style:id=530,.parent=43"
         _StyleDefs(557) =   "Splits(0).Columns(131).HeadingStyle:id=527,.parent=44"
         _StyleDefs(558) =   "Splits(0).Columns(131).FooterStyle:id=528,.parent=45"
         _StyleDefs(559) =   "Splits(0).Columns(131).EditorStyle:id=529,.parent=47"
         _StyleDefs(560) =   "Splits(0).Columns(132).Style:id=534,.parent=43"
         _StyleDefs(561) =   "Splits(0).Columns(132).HeadingStyle:id=531,.parent=44"
         _StyleDefs(562) =   "Splits(0).Columns(132).FooterStyle:id=532,.parent=45"
         _StyleDefs(563) =   "Splits(0).Columns(132).EditorStyle:id=533,.parent=47"
         _StyleDefs(564) =   "Splits(0).Columns(133).Style:id=538,.parent=43"
         _StyleDefs(565) =   "Splits(0).Columns(133).HeadingStyle:id=535,.parent=44"
         _StyleDefs(566) =   "Splits(0).Columns(133).FooterStyle:id=536,.parent=45"
         _StyleDefs(567) =   "Splits(0).Columns(133).EditorStyle:id=537,.parent=47"
         _StyleDefs(568) =   "Splits(0).Columns(134).Style:id=542,.parent=43"
         _StyleDefs(569) =   "Splits(0).Columns(134).HeadingStyle:id=539,.parent=44"
         _StyleDefs(570) =   "Splits(0).Columns(134).FooterStyle:id=540,.parent=45"
         _StyleDefs(571) =   "Splits(0).Columns(134).EditorStyle:id=541,.parent=47"
         _StyleDefs(572) =   "Splits(0).Columns(135).Style:id=546,.parent=43"
         _StyleDefs(573) =   "Splits(0).Columns(135).HeadingStyle:id=543,.parent=44"
         _StyleDefs(574) =   "Splits(0).Columns(135).FooterStyle:id=544,.parent=45"
         _StyleDefs(575) =   "Splits(0).Columns(135).EditorStyle:id=545,.parent=47"
         _StyleDefs(576) =   "Splits(0).Columns(136).Style:id=550,.parent=43"
         _StyleDefs(577) =   "Splits(0).Columns(136).HeadingStyle:id=547,.parent=44"
         _StyleDefs(578) =   "Splits(0).Columns(136).FooterStyle:id=548,.parent=45"
         _StyleDefs(579) =   "Splits(0).Columns(136).EditorStyle:id=549,.parent=47"
         _StyleDefs(580) =   "Splits(0).Columns(137).Style:id=554,.parent=43"
         _StyleDefs(581) =   "Splits(0).Columns(137).HeadingStyle:id=551,.parent=44"
         _StyleDefs(582) =   "Splits(0).Columns(137).FooterStyle:id=552,.parent=45"
         _StyleDefs(583) =   "Splits(0).Columns(137).EditorStyle:id=553,.parent=47"
         _StyleDefs(584) =   "Splits(0).Columns(138).Style:id=558,.parent=43"
         _StyleDefs(585) =   "Splits(0).Columns(138).HeadingStyle:id=555,.parent=44"
         _StyleDefs(586) =   "Splits(0).Columns(138).FooterStyle:id=556,.parent=45"
         _StyleDefs(587) =   "Splits(0).Columns(138).EditorStyle:id=557,.parent=47"
         _StyleDefs(588) =   "Splits(0).Columns(139).Style:id=562,.parent=43"
         _StyleDefs(589) =   "Splits(0).Columns(139).HeadingStyle:id=559,.parent=44"
         _StyleDefs(590) =   "Splits(0).Columns(139).FooterStyle:id=560,.parent=45"
         _StyleDefs(591) =   "Splits(0).Columns(139).EditorStyle:id=561,.parent=47"
         _StyleDefs(592) =   "Splits(0).Columns(140).Style:id=566,.parent=43"
         _StyleDefs(593) =   "Splits(0).Columns(140).HeadingStyle:id=563,.parent=44"
         _StyleDefs(594) =   "Splits(0).Columns(140).FooterStyle:id=564,.parent=45"
         _StyleDefs(595) =   "Splits(0).Columns(140).EditorStyle:id=565,.parent=47"
         _StyleDefs(596) =   "Splits(0).Columns(141).Style:id=570,.parent=43"
         _StyleDefs(597) =   "Splits(0).Columns(141).HeadingStyle:id=567,.parent=44"
         _StyleDefs(598) =   "Splits(0).Columns(141).FooterStyle:id=568,.parent=45"
         _StyleDefs(599) =   "Splits(0).Columns(141).EditorStyle:id=569,.parent=47"
         _StyleDefs(600) =   "Splits(0).Columns(142).Style:id=574,.parent=43"
         _StyleDefs(601) =   "Splits(0).Columns(142).HeadingStyle:id=571,.parent=44"
         _StyleDefs(602) =   "Splits(0).Columns(142).FooterStyle:id=572,.parent=45"
         _StyleDefs(603) =   "Splits(0).Columns(142).EditorStyle:id=573,.parent=47"
         _StyleDefs(604) =   "Splits(0).Columns(143).Style:id=578,.parent=43"
         _StyleDefs(605) =   "Splits(0).Columns(143).HeadingStyle:id=575,.parent=44"
         _StyleDefs(606) =   "Splits(0).Columns(143).FooterStyle:id=576,.parent=45"
         _StyleDefs(607) =   "Splits(0).Columns(143).EditorStyle:id=577,.parent=47"
         _StyleDefs(608) =   "Splits(0).Columns(144).Style:id=582,.parent=43"
         _StyleDefs(609) =   "Splits(0).Columns(144).HeadingStyle:id=579,.parent=44"
         _StyleDefs(610) =   "Splits(0).Columns(144).FooterStyle:id=580,.parent=45"
         _StyleDefs(611) =   "Splits(0).Columns(144).EditorStyle:id=581,.parent=47"
         _StyleDefs(612) =   "Splits(0).Columns(145).Style:id=586,.parent=43"
         _StyleDefs(613) =   "Splits(0).Columns(145).HeadingStyle:id=583,.parent=44"
         _StyleDefs(614) =   "Splits(0).Columns(145).FooterStyle:id=584,.parent=45"
         _StyleDefs(615) =   "Splits(0).Columns(145).EditorStyle:id=585,.parent=47"
         _StyleDefs(616) =   "Splits(0).Columns(146).Style:id=590,.parent=43"
         _StyleDefs(617) =   "Splits(0).Columns(146).HeadingStyle:id=587,.parent=44"
         _StyleDefs(618) =   "Splits(0).Columns(146).FooterStyle:id=588,.parent=45"
         _StyleDefs(619) =   "Splits(0).Columns(146).EditorStyle:id=589,.parent=47"
         _StyleDefs(620) =   "Splits(0).Columns(147).Style:id=594,.parent=43"
         _StyleDefs(621) =   "Splits(0).Columns(147).HeadingStyle:id=591,.parent=44"
         _StyleDefs(622) =   "Splits(0).Columns(147).FooterStyle:id=592,.parent=45"
         _StyleDefs(623) =   "Splits(0).Columns(147).EditorStyle:id=593,.parent=47"
         _StyleDefs(624) =   "Splits(0).Columns(148).Style:id=598,.parent=43"
         _StyleDefs(625) =   "Splits(0).Columns(148).HeadingStyle:id=595,.parent=44"
         _StyleDefs(626) =   "Splits(0).Columns(148).FooterStyle:id=596,.parent=45"
         _StyleDefs(627) =   "Splits(0).Columns(148).EditorStyle:id=597,.parent=47"
         _StyleDefs(628) =   "Splits(0).Columns(149).Style:id=602,.parent=43"
         _StyleDefs(629) =   "Splits(0).Columns(149).HeadingStyle:id=599,.parent=44"
         _StyleDefs(630) =   "Splits(0).Columns(149).FooterStyle:id=600,.parent=45"
         _StyleDefs(631) =   "Splits(0).Columns(149).EditorStyle:id=601,.parent=47"
         _StyleDefs(632) =   "Splits(0).Columns(150).Style:id=606,.parent=43"
         _StyleDefs(633) =   "Splits(0).Columns(150).HeadingStyle:id=603,.parent=44"
         _StyleDefs(634) =   "Splits(0).Columns(150).FooterStyle:id=604,.parent=45"
         _StyleDefs(635) =   "Splits(0).Columns(150).EditorStyle:id=605,.parent=47"
         _StyleDefs(636) =   "Splits(1).Style:id=13,.parent=1"
         _StyleDefs(637) =   "Splits(1).CaptionStyle:id=22,.parent=4"
         _StyleDefs(638) =   "Splits(1).HeadingStyle:id=14,.parent=2"
         _StyleDefs(639) =   "Splits(1).FooterStyle:id=15,.parent=3"
         _StyleDefs(640) =   "Splits(1).InactiveStyle:id=16,.parent=5"
         _StyleDefs(641) =   "Splits(1).SelectedStyle:id=18,.parent=6"
         _StyleDefs(642) =   "Splits(1).EditorStyle:id=17,.parent=7"
         _StyleDefs(643) =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(644) =   "Splits(1).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(645) =   "Splits(1).OddRowStyle:id=21,.parent=10"
         _StyleDefs(646) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(647) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(648) =   "Splits(1).Columns(0).Style:id=1270,.parent=13"
         _StyleDefs(649) =   "Splits(1).Columns(0).HeadingStyle:id=1267,.parent=14"
         _StyleDefs(650) =   "Splits(1).Columns(0).FooterStyle:id=1268,.parent=15"
         _StyleDefs(651) =   "Splits(1).Columns(0).EditorStyle:id=1269,.parent=17"
         _StyleDefs(652) =   "Splits(1).Columns(1).Style:id=1262,.parent=13"
         _StyleDefs(653) =   "Splits(1).Columns(1).HeadingStyle:id=1259,.parent=14"
         _StyleDefs(654) =   "Splits(1).Columns(1).FooterStyle:id=1260,.parent=15"
         _StyleDefs(655) =   "Splits(1).Columns(1).EditorStyle:id=1261,.parent=17"
         _StyleDefs(656) =   "Splits(1).Columns(2).Style:id=1254,.parent=13"
         _StyleDefs(657) =   "Splits(1).Columns(2).HeadingStyle:id=1251,.parent=14"
         _StyleDefs(658) =   "Splits(1).Columns(2).FooterStyle:id=1252,.parent=15"
         _StyleDefs(659) =   "Splits(1).Columns(2).EditorStyle:id=1253,.parent=17"
         _StyleDefs(660) =   "Splits(1).Columns(3).Style:id=1246,.parent=13"
         _StyleDefs(661) =   "Splits(1).Columns(3).HeadingStyle:id=1243,.parent=14"
         _StyleDefs(662) =   "Splits(1).Columns(3).FooterStyle:id=1244,.parent=15"
         _StyleDefs(663) =   "Splits(1).Columns(3).EditorStyle:id=1245,.parent=17"
         _StyleDefs(664) =   "Splits(1).Columns(4).Style:id=1238,.parent=13"
         _StyleDefs(665) =   "Splits(1).Columns(4).HeadingStyle:id=1235,.parent=14"
         _StyleDefs(666) =   "Splits(1).Columns(4).FooterStyle:id=1236,.parent=15"
         _StyleDefs(667) =   "Splits(1).Columns(4).EditorStyle:id=1237,.parent=17"
         _StyleDefs(668) =   "Splits(1).Columns(5).Style:id=1230,.parent=13"
         _StyleDefs(669) =   "Splits(1).Columns(5).HeadingStyle:id=1227,.parent=14"
         _StyleDefs(670) =   "Splits(1).Columns(5).FooterStyle:id=1228,.parent=15"
         _StyleDefs(671) =   "Splits(1).Columns(5).EditorStyle:id=1229,.parent=17"
         _StyleDefs(672) =   "Splits(1).Columns(6).Style:id=1222,.parent=13"
         _StyleDefs(673) =   "Splits(1).Columns(6).HeadingStyle:id=1219,.parent=14"
         _StyleDefs(674) =   "Splits(1).Columns(6).FooterStyle:id=1220,.parent=15"
         _StyleDefs(675) =   "Splits(1).Columns(6).EditorStyle:id=1221,.parent=17"
         _StyleDefs(676) =   "Splits(1).Columns(7).Style:id=1214,.parent=13"
         _StyleDefs(677) =   "Splits(1).Columns(7).HeadingStyle:id=1211,.parent=14"
         _StyleDefs(678) =   "Splits(1).Columns(7).FooterStyle:id=1212,.parent=15"
         _StyleDefs(679) =   "Splits(1).Columns(7).EditorStyle:id=1213,.parent=17"
         _StyleDefs(680) =   "Splits(1).Columns(8).Style:id=1206,.parent=13"
         _StyleDefs(681) =   "Splits(1).Columns(8).HeadingStyle:id=1203,.parent=14"
         _StyleDefs(682) =   "Splits(1).Columns(8).FooterStyle:id=1204,.parent=15"
         _StyleDefs(683) =   "Splits(1).Columns(8).EditorStyle:id=1205,.parent=17"
         _StyleDefs(684) =   "Splits(1).Columns(9).Style:id=1198,.parent=13"
         _StyleDefs(685) =   "Splits(1).Columns(9).HeadingStyle:id=1195,.parent=14"
         _StyleDefs(686) =   "Splits(1).Columns(9).FooterStyle:id=1196,.parent=15"
         _StyleDefs(687) =   "Splits(1).Columns(9).EditorStyle:id=1197,.parent=17"
         _StyleDefs(688) =   "Splits(1).Columns(10).Style:id=1190,.parent=13"
         _StyleDefs(689) =   "Splits(1).Columns(10).HeadingStyle:id=1187,.parent=14"
         _StyleDefs(690) =   "Splits(1).Columns(10).FooterStyle:id=1188,.parent=15"
         _StyleDefs(691) =   "Splits(1).Columns(10).EditorStyle:id=1189,.parent=17"
         _StyleDefs(692) =   "Splits(1).Columns(11).Style:id=1182,.parent=13"
         _StyleDefs(693) =   "Splits(1).Columns(11).HeadingStyle:id=1179,.parent=14"
         _StyleDefs(694) =   "Splits(1).Columns(11).FooterStyle:id=1180,.parent=15"
         _StyleDefs(695) =   "Splits(1).Columns(11).EditorStyle:id=1181,.parent=17"
         _StyleDefs(696) =   "Splits(1).Columns(12).Style:id=1174,.parent=13"
         _StyleDefs(697) =   "Splits(1).Columns(12).HeadingStyle:id=1171,.parent=14"
         _StyleDefs(698) =   "Splits(1).Columns(12).FooterStyle:id=1172,.parent=15"
         _StyleDefs(699) =   "Splits(1).Columns(12).EditorStyle:id=1173,.parent=17"
         _StyleDefs(700) =   "Splits(1).Columns(13).Style:id=1166,.parent=13"
         _StyleDefs(701) =   "Splits(1).Columns(13).HeadingStyle:id=1163,.parent=14"
         _StyleDefs(702) =   "Splits(1).Columns(13).FooterStyle:id=1164,.parent=15"
         _StyleDefs(703) =   "Splits(1).Columns(13).EditorStyle:id=1165,.parent=17"
         _StyleDefs(704) =   "Splits(1).Columns(14).Style:id=1158,.parent=13"
         _StyleDefs(705) =   "Splits(1).Columns(14).HeadingStyle:id=1155,.parent=14"
         _StyleDefs(706) =   "Splits(1).Columns(14).FooterStyle:id=1156,.parent=15"
         _StyleDefs(707) =   "Splits(1).Columns(14).EditorStyle:id=1157,.parent=17"
         _StyleDefs(708) =   "Splits(1).Columns(15).Style:id=1150,.parent=13"
         _StyleDefs(709) =   "Splits(1).Columns(15).HeadingStyle:id=1147,.parent=14"
         _StyleDefs(710) =   "Splits(1).Columns(15).FooterStyle:id=1148,.parent=15"
         _StyleDefs(711) =   "Splits(1).Columns(15).EditorStyle:id=1149,.parent=17"
         _StyleDefs(712) =   "Splits(1).Columns(16).Style:id=1142,.parent=13"
         _StyleDefs(713) =   "Splits(1).Columns(16).HeadingStyle:id=1139,.parent=14"
         _StyleDefs(714) =   "Splits(1).Columns(16).FooterStyle:id=1140,.parent=15"
         _StyleDefs(715) =   "Splits(1).Columns(16).EditorStyle:id=1141,.parent=17"
         _StyleDefs(716) =   "Splits(1).Columns(17).Style:id=1134,.parent=13"
         _StyleDefs(717) =   "Splits(1).Columns(17).HeadingStyle:id=1131,.parent=14"
         _StyleDefs(718) =   "Splits(1).Columns(17).FooterStyle:id=1132,.parent=15"
         _StyleDefs(719) =   "Splits(1).Columns(17).EditorStyle:id=1133,.parent=17"
         _StyleDefs(720) =   "Splits(1).Columns(18).Style:id=1126,.parent=13"
         _StyleDefs(721) =   "Splits(1).Columns(18).HeadingStyle:id=1123,.parent=14"
         _StyleDefs(722) =   "Splits(1).Columns(18).FooterStyle:id=1124,.parent=15"
         _StyleDefs(723) =   "Splits(1).Columns(18).EditorStyle:id=1125,.parent=17"
         _StyleDefs(724) =   "Splits(1).Columns(19).Style:id=1118,.parent=13"
         _StyleDefs(725) =   "Splits(1).Columns(19).HeadingStyle:id=1115,.parent=14"
         _StyleDefs(726) =   "Splits(1).Columns(19).FooterStyle:id=1116,.parent=15"
         _StyleDefs(727) =   "Splits(1).Columns(19).EditorStyle:id=1117,.parent=17"
         _StyleDefs(728) =   "Splits(1).Columns(20).Style:id=1110,.parent=13"
         _StyleDefs(729) =   "Splits(1).Columns(20).HeadingStyle:id=1107,.parent=14"
         _StyleDefs(730) =   "Splits(1).Columns(20).FooterStyle:id=1108,.parent=15"
         _StyleDefs(731) =   "Splits(1).Columns(20).EditorStyle:id=1109,.parent=17"
         _StyleDefs(732) =   "Splits(1).Columns(21).Style:id=1102,.parent=13"
         _StyleDefs(733) =   "Splits(1).Columns(21).HeadingStyle:id=1099,.parent=14"
         _StyleDefs(734) =   "Splits(1).Columns(21).FooterStyle:id=1100,.parent=15"
         _StyleDefs(735) =   "Splits(1).Columns(21).EditorStyle:id=1101,.parent=17"
         _StyleDefs(736) =   "Splits(1).Columns(22).Style:id=1094,.parent=13"
         _StyleDefs(737) =   "Splits(1).Columns(22).HeadingStyle:id=1091,.parent=14"
         _StyleDefs(738) =   "Splits(1).Columns(22).FooterStyle:id=1092,.parent=15"
         _StyleDefs(739) =   "Splits(1).Columns(22).EditorStyle:id=1093,.parent=17"
         _StyleDefs(740) =   "Splits(1).Columns(23).Style:id=1086,.parent=13"
         _StyleDefs(741) =   "Splits(1).Columns(23).HeadingStyle:id=1083,.parent=14"
         _StyleDefs(742) =   "Splits(1).Columns(23).FooterStyle:id=1084,.parent=15"
         _StyleDefs(743) =   "Splits(1).Columns(23).EditorStyle:id=1085,.parent=17"
         _StyleDefs(744) =   "Splits(1).Columns(24).Style:id=1078,.parent=13"
         _StyleDefs(745) =   "Splits(1).Columns(24).HeadingStyle:id=1075,.parent=14"
         _StyleDefs(746) =   "Splits(1).Columns(24).FooterStyle:id=1076,.parent=15"
         _StyleDefs(747) =   "Splits(1).Columns(24).EditorStyle:id=1077,.parent=17"
         _StyleDefs(748) =   "Splits(1).Columns(25).Style:id=1070,.parent=13"
         _StyleDefs(749) =   "Splits(1).Columns(25).HeadingStyle:id=1067,.parent=14"
         _StyleDefs(750) =   "Splits(1).Columns(25).FooterStyle:id=1068,.parent=15"
         _StyleDefs(751) =   "Splits(1).Columns(25).EditorStyle:id=1069,.parent=17"
         _StyleDefs(752) =   "Splits(1).Columns(26).Style:id=1062,.parent=13"
         _StyleDefs(753) =   "Splits(1).Columns(26).HeadingStyle:id=1059,.parent=14"
         _StyleDefs(754) =   "Splits(1).Columns(26).FooterStyle:id=1060,.parent=15"
         _StyleDefs(755) =   "Splits(1).Columns(26).EditorStyle:id=1061,.parent=17"
         _StyleDefs(756) =   "Splits(1).Columns(27).Style:id=1054,.parent=13"
         _StyleDefs(757) =   "Splits(1).Columns(27).HeadingStyle:id=1051,.parent=14"
         _StyleDefs(758) =   "Splits(1).Columns(27).FooterStyle:id=1052,.parent=15"
         _StyleDefs(759) =   "Splits(1).Columns(27).EditorStyle:id=1053,.parent=17"
         _StyleDefs(760) =   "Splits(1).Columns(28).Style:id=1046,.parent=13"
         _StyleDefs(761) =   "Splits(1).Columns(28).HeadingStyle:id=1043,.parent=14"
         _StyleDefs(762) =   "Splits(1).Columns(28).FooterStyle:id=1044,.parent=15"
         _StyleDefs(763) =   "Splits(1).Columns(28).EditorStyle:id=1045,.parent=17"
         _StyleDefs(764) =   "Splits(1).Columns(29).Style:id=1038,.parent=13"
         _StyleDefs(765) =   "Splits(1).Columns(29).HeadingStyle:id=1035,.parent=14"
         _StyleDefs(766) =   "Splits(1).Columns(29).FooterStyle:id=1036,.parent=15"
         _StyleDefs(767) =   "Splits(1).Columns(29).EditorStyle:id=1037,.parent=17"
         _StyleDefs(768) =   "Splits(1).Columns(30).Style:id=1030,.parent=13"
         _StyleDefs(769) =   "Splits(1).Columns(30).HeadingStyle:id=1027,.parent=14"
         _StyleDefs(770) =   "Splits(1).Columns(30).FooterStyle:id=1028,.parent=15"
         _StyleDefs(771) =   "Splits(1).Columns(30).EditorStyle:id=1029,.parent=17"
         _StyleDefs(772) =   "Splits(1).Columns(31).Style:id=1022,.parent=13"
         _StyleDefs(773) =   "Splits(1).Columns(31).HeadingStyle:id=1019,.parent=14"
         _StyleDefs(774) =   "Splits(1).Columns(31).FooterStyle:id=1020,.parent=15"
         _StyleDefs(775) =   "Splits(1).Columns(31).EditorStyle:id=1021,.parent=17"
         _StyleDefs(776) =   "Splits(1).Columns(32).Style:id=1014,.parent=13"
         _StyleDefs(777) =   "Splits(1).Columns(32).HeadingStyle:id=1011,.parent=14"
         _StyleDefs(778) =   "Splits(1).Columns(32).FooterStyle:id=1012,.parent=15"
         _StyleDefs(779) =   "Splits(1).Columns(32).EditorStyle:id=1013,.parent=17"
         _StyleDefs(780) =   "Splits(1).Columns(33).Style:id=1006,.parent=13"
         _StyleDefs(781) =   "Splits(1).Columns(33).HeadingStyle:id=1003,.parent=14"
         _StyleDefs(782) =   "Splits(1).Columns(33).FooterStyle:id=1004,.parent=15"
         _StyleDefs(783) =   "Splits(1).Columns(33).EditorStyle:id=1005,.parent=17"
         _StyleDefs(784) =   "Splits(1).Columns(34).Style:id=998,.parent=13"
         _StyleDefs(785) =   "Splits(1).Columns(34).HeadingStyle:id=995,.parent=14"
         _StyleDefs(786) =   "Splits(1).Columns(34).FooterStyle:id=996,.parent=15"
         _StyleDefs(787) =   "Splits(1).Columns(34).EditorStyle:id=997,.parent=17"
         _StyleDefs(788) =   "Splits(1).Columns(35).Style:id=990,.parent=13"
         _StyleDefs(789) =   "Splits(1).Columns(35).HeadingStyle:id=987,.parent=14"
         _StyleDefs(790) =   "Splits(1).Columns(35).FooterStyle:id=988,.parent=15"
         _StyleDefs(791) =   "Splits(1).Columns(35).EditorStyle:id=989,.parent=17"
         _StyleDefs(792) =   "Splits(1).Columns(36).Style:id=982,.parent=13"
         _StyleDefs(793) =   "Splits(1).Columns(36).HeadingStyle:id=979,.parent=14"
         _StyleDefs(794) =   "Splits(1).Columns(36).FooterStyle:id=980,.parent=15"
         _StyleDefs(795) =   "Splits(1).Columns(36).EditorStyle:id=981,.parent=17"
         _StyleDefs(796) =   "Splits(1).Columns(37).Style:id=974,.parent=13"
         _StyleDefs(797) =   "Splits(1).Columns(37).HeadingStyle:id=971,.parent=14"
         _StyleDefs(798) =   "Splits(1).Columns(37).FooterStyle:id=972,.parent=15"
         _StyleDefs(799) =   "Splits(1).Columns(37).EditorStyle:id=973,.parent=17"
         _StyleDefs(800) =   "Splits(1).Columns(38).Style:id=966,.parent=13"
         _StyleDefs(801) =   "Splits(1).Columns(38).HeadingStyle:id=963,.parent=14"
         _StyleDefs(802) =   "Splits(1).Columns(38).FooterStyle:id=964,.parent=15"
         _StyleDefs(803) =   "Splits(1).Columns(38).EditorStyle:id=965,.parent=17"
         _StyleDefs(804) =   "Splits(1).Columns(39).Style:id=958,.parent=13"
         _StyleDefs(805) =   "Splits(1).Columns(39).HeadingStyle:id=955,.parent=14"
         _StyleDefs(806) =   "Splits(1).Columns(39).FooterStyle:id=956,.parent=15"
         _StyleDefs(807) =   "Splits(1).Columns(39).EditorStyle:id=957,.parent=17"
         _StyleDefs(808) =   "Splits(1).Columns(40).Style:id=950,.parent=13"
         _StyleDefs(809) =   "Splits(1).Columns(40).HeadingStyle:id=947,.parent=14"
         _StyleDefs(810) =   "Splits(1).Columns(40).FooterStyle:id=948,.parent=15"
         _StyleDefs(811) =   "Splits(1).Columns(40).EditorStyle:id=949,.parent=17"
         _StyleDefs(812) =   "Splits(1).Columns(41).Style:id=942,.parent=13"
         _StyleDefs(813) =   "Splits(1).Columns(41).HeadingStyle:id=939,.parent=14"
         _StyleDefs(814) =   "Splits(1).Columns(41).FooterStyle:id=940,.parent=15"
         _StyleDefs(815) =   "Splits(1).Columns(41).EditorStyle:id=941,.parent=17"
         _StyleDefs(816) =   "Splits(1).Columns(42).Style:id=934,.parent=13"
         _StyleDefs(817) =   "Splits(1).Columns(42).HeadingStyle:id=931,.parent=14"
         _StyleDefs(818) =   "Splits(1).Columns(42).FooterStyle:id=932,.parent=15"
         _StyleDefs(819) =   "Splits(1).Columns(42).EditorStyle:id=933,.parent=17"
         _StyleDefs(820) =   "Splits(1).Columns(43).Style:id=926,.parent=13"
         _StyleDefs(821) =   "Splits(1).Columns(43).HeadingStyle:id=923,.parent=14"
         _StyleDefs(822) =   "Splits(1).Columns(43).FooterStyle:id=924,.parent=15"
         _StyleDefs(823) =   "Splits(1).Columns(43).EditorStyle:id=925,.parent=17"
         _StyleDefs(824) =   "Splits(1).Columns(44).Style:id=918,.parent=13"
         _StyleDefs(825) =   "Splits(1).Columns(44).HeadingStyle:id=915,.parent=14"
         _StyleDefs(826) =   "Splits(1).Columns(44).FooterStyle:id=916,.parent=15"
         _StyleDefs(827) =   "Splits(1).Columns(44).EditorStyle:id=917,.parent=17"
         _StyleDefs(828) =   "Splits(1).Columns(45).Style:id=910,.parent=13"
         _StyleDefs(829) =   "Splits(1).Columns(45).HeadingStyle:id=907,.parent=14"
         _StyleDefs(830) =   "Splits(1).Columns(45).FooterStyle:id=908,.parent=15"
         _StyleDefs(831) =   "Splits(1).Columns(45).EditorStyle:id=909,.parent=17"
         _StyleDefs(832) =   "Splits(1).Columns(46).Style:id=902,.parent=13"
         _StyleDefs(833) =   "Splits(1).Columns(46).HeadingStyle:id=899,.parent=14"
         _StyleDefs(834) =   "Splits(1).Columns(46).FooterStyle:id=900,.parent=15"
         _StyleDefs(835) =   "Splits(1).Columns(46).EditorStyle:id=901,.parent=17"
         _StyleDefs(836) =   "Splits(1).Columns(47).Style:id=894,.parent=13"
         _StyleDefs(837) =   "Splits(1).Columns(47).HeadingStyle:id=891,.parent=14"
         _StyleDefs(838) =   "Splits(1).Columns(47).FooterStyle:id=892,.parent=15"
         _StyleDefs(839) =   "Splits(1).Columns(47).EditorStyle:id=893,.parent=17"
         _StyleDefs(840) =   "Splits(1).Columns(48).Style:id=886,.parent=13"
         _StyleDefs(841) =   "Splits(1).Columns(48).HeadingStyle:id=883,.parent=14"
         _StyleDefs(842) =   "Splits(1).Columns(48).FooterStyle:id=884,.parent=15"
         _StyleDefs(843) =   "Splits(1).Columns(48).EditorStyle:id=885,.parent=17"
         _StyleDefs(844) =   "Splits(1).Columns(49).Style:id=878,.parent=13"
         _StyleDefs(845) =   "Splits(1).Columns(49).HeadingStyle:id=875,.parent=14"
         _StyleDefs(846) =   "Splits(1).Columns(49).FooterStyle:id=876,.parent=15"
         _StyleDefs(847) =   "Splits(1).Columns(49).EditorStyle:id=877,.parent=17"
         _StyleDefs(848) =   "Splits(1).Columns(50).Style:id=870,.parent=13"
         _StyleDefs(849) =   "Splits(1).Columns(50).HeadingStyle:id=867,.parent=14"
         _StyleDefs(850) =   "Splits(1).Columns(50).FooterStyle:id=868,.parent=15"
         _StyleDefs(851) =   "Splits(1).Columns(50).EditorStyle:id=869,.parent=17"
         _StyleDefs(852) =   "Splits(1).Columns(51).Style:id=862,.parent=13"
         _StyleDefs(853) =   "Splits(1).Columns(51).HeadingStyle:id=859,.parent=14"
         _StyleDefs(854) =   "Splits(1).Columns(51).FooterStyle:id=860,.parent=15"
         _StyleDefs(855) =   "Splits(1).Columns(51).EditorStyle:id=861,.parent=17"
         _StyleDefs(856) =   "Splits(1).Columns(52).Style:id=854,.parent=13"
         _StyleDefs(857) =   "Splits(1).Columns(52).HeadingStyle:id=851,.parent=14"
         _StyleDefs(858) =   "Splits(1).Columns(52).FooterStyle:id=852,.parent=15"
         _StyleDefs(859) =   "Splits(1).Columns(52).EditorStyle:id=853,.parent=17"
         _StyleDefs(860) =   "Splits(1).Columns(53).Style:id=846,.parent=13"
         _StyleDefs(861) =   "Splits(1).Columns(53).HeadingStyle:id=843,.parent=14"
         _StyleDefs(862) =   "Splits(1).Columns(53).FooterStyle:id=844,.parent=15"
         _StyleDefs(863) =   "Splits(1).Columns(53).EditorStyle:id=845,.parent=17"
         _StyleDefs(864) =   "Splits(1).Columns(54).Style:id=838,.parent=13"
         _StyleDefs(865) =   "Splits(1).Columns(54).HeadingStyle:id=835,.parent=14"
         _StyleDefs(866) =   "Splits(1).Columns(54).FooterStyle:id=836,.parent=15"
         _StyleDefs(867) =   "Splits(1).Columns(54).EditorStyle:id=837,.parent=17"
         _StyleDefs(868) =   "Splits(1).Columns(55).Style:id=830,.parent=13"
         _StyleDefs(869) =   "Splits(1).Columns(55).HeadingStyle:id=827,.parent=14"
         _StyleDefs(870) =   "Splits(1).Columns(55).FooterStyle:id=828,.parent=15"
         _StyleDefs(871) =   "Splits(1).Columns(55).EditorStyle:id=829,.parent=17"
         _StyleDefs(872) =   "Splits(1).Columns(56).Style:id=822,.parent=13"
         _StyleDefs(873) =   "Splits(1).Columns(56).HeadingStyle:id=819,.parent=14"
         _StyleDefs(874) =   "Splits(1).Columns(56).FooterStyle:id=820,.parent=15"
         _StyleDefs(875) =   "Splits(1).Columns(56).EditorStyle:id=821,.parent=17"
         _StyleDefs(876) =   "Splits(1).Columns(57).Style:id=814,.parent=13"
         _StyleDefs(877) =   "Splits(1).Columns(57).HeadingStyle:id=811,.parent=14"
         _StyleDefs(878) =   "Splits(1).Columns(57).FooterStyle:id=812,.parent=15"
         _StyleDefs(879) =   "Splits(1).Columns(57).EditorStyle:id=813,.parent=17"
         _StyleDefs(880) =   "Splits(1).Columns(58).Style:id=806,.parent=13"
         _StyleDefs(881) =   "Splits(1).Columns(58).HeadingStyle:id=803,.parent=14"
         _StyleDefs(882) =   "Splits(1).Columns(58).FooterStyle:id=804,.parent=15"
         _StyleDefs(883) =   "Splits(1).Columns(58).EditorStyle:id=805,.parent=17"
         _StyleDefs(884) =   "Splits(1).Columns(59).Style:id=798,.parent=13"
         _StyleDefs(885) =   "Splits(1).Columns(59).HeadingStyle:id=795,.parent=14"
         _StyleDefs(886) =   "Splits(1).Columns(59).FooterStyle:id=796,.parent=15"
         _StyleDefs(887) =   "Splits(1).Columns(59).EditorStyle:id=797,.parent=17"
         _StyleDefs(888) =   "Splits(1).Columns(60).Style:id=790,.parent=13"
         _StyleDefs(889) =   "Splits(1).Columns(60).HeadingStyle:id=787,.parent=14"
         _StyleDefs(890) =   "Splits(1).Columns(60).FooterStyle:id=788,.parent=15"
         _StyleDefs(891) =   "Splits(1).Columns(60).EditorStyle:id=789,.parent=17"
         _StyleDefs(892) =   "Splits(1).Columns(61).Style:id=782,.parent=13"
         _StyleDefs(893) =   "Splits(1).Columns(61).HeadingStyle:id=779,.parent=14"
         _StyleDefs(894) =   "Splits(1).Columns(61).FooterStyle:id=780,.parent=15"
         _StyleDefs(895) =   "Splits(1).Columns(61).EditorStyle:id=781,.parent=17"
         _StyleDefs(896) =   "Splits(1).Columns(62).Style:id=774,.parent=13"
         _StyleDefs(897) =   "Splits(1).Columns(62).HeadingStyle:id=771,.parent=14"
         _StyleDefs(898) =   "Splits(1).Columns(62).FooterStyle:id=772,.parent=15"
         _StyleDefs(899) =   "Splits(1).Columns(62).EditorStyle:id=773,.parent=17"
         _StyleDefs(900) =   "Splits(1).Columns(63).Style:id=766,.parent=13"
         _StyleDefs(901) =   "Splits(1).Columns(63).HeadingStyle:id=763,.parent=14"
         _StyleDefs(902) =   "Splits(1).Columns(63).FooterStyle:id=764,.parent=15"
         _StyleDefs(903) =   "Splits(1).Columns(63).EditorStyle:id=765,.parent=17"
         _StyleDefs(904) =   "Splits(1).Columns(64).Style:id=758,.parent=13"
         _StyleDefs(905) =   "Splits(1).Columns(64).HeadingStyle:id=755,.parent=14"
         _StyleDefs(906) =   "Splits(1).Columns(64).FooterStyle:id=756,.parent=15"
         _StyleDefs(907) =   "Splits(1).Columns(64).EditorStyle:id=757,.parent=17"
         _StyleDefs(908) =   "Splits(1).Columns(65).Style:id=750,.parent=13"
         _StyleDefs(909) =   "Splits(1).Columns(65).HeadingStyle:id=747,.parent=14"
         _StyleDefs(910) =   "Splits(1).Columns(65).FooterStyle:id=748,.parent=15"
         _StyleDefs(911) =   "Splits(1).Columns(65).EditorStyle:id=749,.parent=17"
         _StyleDefs(912) =   "Splits(1).Columns(66).Style:id=742,.parent=13"
         _StyleDefs(913) =   "Splits(1).Columns(66).HeadingStyle:id=739,.parent=14"
         _StyleDefs(914) =   "Splits(1).Columns(66).FooterStyle:id=740,.parent=15"
         _StyleDefs(915) =   "Splits(1).Columns(66).EditorStyle:id=741,.parent=17"
         _StyleDefs(916) =   "Splits(1).Columns(67).Style:id=734,.parent=13"
         _StyleDefs(917) =   "Splits(1).Columns(67).HeadingStyle:id=731,.parent=14"
         _StyleDefs(918) =   "Splits(1).Columns(67).FooterStyle:id=732,.parent=15"
         _StyleDefs(919) =   "Splits(1).Columns(67).EditorStyle:id=733,.parent=17"
         _StyleDefs(920) =   "Splits(1).Columns(68).Style:id=726,.parent=13"
         _StyleDefs(921) =   "Splits(1).Columns(68).HeadingStyle:id=723,.parent=14"
         _StyleDefs(922) =   "Splits(1).Columns(68).FooterStyle:id=724,.parent=15"
         _StyleDefs(923) =   "Splits(1).Columns(68).EditorStyle:id=725,.parent=17"
         _StyleDefs(924) =   "Splits(1).Columns(69).Style:id=718,.parent=13"
         _StyleDefs(925) =   "Splits(1).Columns(69).HeadingStyle:id=715,.parent=14"
         _StyleDefs(926) =   "Splits(1).Columns(69).FooterStyle:id=716,.parent=15"
         _StyleDefs(927) =   "Splits(1).Columns(69).EditorStyle:id=717,.parent=17"
         _StyleDefs(928) =   "Splits(1).Columns(70).Style:id=710,.parent=13"
         _StyleDefs(929) =   "Splits(1).Columns(70).HeadingStyle:id=707,.parent=14"
         _StyleDefs(930) =   "Splits(1).Columns(70).FooterStyle:id=708,.parent=15"
         _StyleDefs(931) =   "Splits(1).Columns(70).EditorStyle:id=709,.parent=17"
         _StyleDefs(932) =   "Splits(1).Columns(71).Style:id=702,.parent=13"
         _StyleDefs(933) =   "Splits(1).Columns(71).HeadingStyle:id=699,.parent=14"
         _StyleDefs(934) =   "Splits(1).Columns(71).FooterStyle:id=700,.parent=15"
         _StyleDefs(935) =   "Splits(1).Columns(71).EditorStyle:id=701,.parent=17"
         _StyleDefs(936) =   "Splits(1).Columns(72).Style:id=694,.parent=13"
         _StyleDefs(937) =   "Splits(1).Columns(72).HeadingStyle:id=691,.parent=14"
         _StyleDefs(938) =   "Splits(1).Columns(72).FooterStyle:id=692,.parent=15"
         _StyleDefs(939) =   "Splits(1).Columns(72).EditorStyle:id=693,.parent=17"
         _StyleDefs(940) =   "Splits(1).Columns(73).Style:id=686,.parent=13"
         _StyleDefs(941) =   "Splits(1).Columns(73).HeadingStyle:id=683,.parent=14"
         _StyleDefs(942) =   "Splits(1).Columns(73).FooterStyle:id=684,.parent=15"
         _StyleDefs(943) =   "Splits(1).Columns(73).EditorStyle:id=685,.parent=17"
         _StyleDefs(944) =   "Splits(1).Columns(74).Style:id=678,.parent=13"
         _StyleDefs(945) =   "Splits(1).Columns(74).HeadingStyle:id=675,.parent=14"
         _StyleDefs(946) =   "Splits(1).Columns(74).FooterStyle:id=676,.parent=15"
         _StyleDefs(947) =   "Splits(1).Columns(74).EditorStyle:id=677,.parent=17"
         _StyleDefs(948) =   "Splits(1).Columns(75).Style:id=670,.parent=13"
         _StyleDefs(949) =   "Splits(1).Columns(75).HeadingStyle:id=667,.parent=14"
         _StyleDefs(950) =   "Splits(1).Columns(75).FooterStyle:id=668,.parent=15"
         _StyleDefs(951) =   "Splits(1).Columns(75).EditorStyle:id=669,.parent=17"
         _StyleDefs(952) =   "Splits(1).Columns(76).Style:id=662,.parent=13"
         _StyleDefs(953) =   "Splits(1).Columns(76).HeadingStyle:id=659,.parent=14"
         _StyleDefs(954) =   "Splits(1).Columns(76).FooterStyle:id=660,.parent=15"
         _StyleDefs(955) =   "Splits(1).Columns(76).EditorStyle:id=661,.parent=17"
         _StyleDefs(956) =   "Splits(1).Columns(77).Style:id=654,.parent=13"
         _StyleDefs(957) =   "Splits(1).Columns(77).HeadingStyle:id=651,.parent=14"
         _StyleDefs(958) =   "Splits(1).Columns(77).FooterStyle:id=652,.parent=15"
         _StyleDefs(959) =   "Splits(1).Columns(77).EditorStyle:id=653,.parent=17"
         _StyleDefs(960) =   "Splits(1).Columns(78).Style:id=646,.parent=13"
         _StyleDefs(961) =   "Splits(1).Columns(78).HeadingStyle:id=643,.parent=14"
         _StyleDefs(962) =   "Splits(1).Columns(78).FooterStyle:id=644,.parent=15"
         _StyleDefs(963) =   "Splits(1).Columns(78).EditorStyle:id=645,.parent=17"
         _StyleDefs(964) =   "Splits(1).Columns(79).Style:id=638,.parent=13"
         _StyleDefs(965) =   "Splits(1).Columns(79).HeadingStyle:id=635,.parent=14"
         _StyleDefs(966) =   "Splits(1).Columns(79).FooterStyle:id=636,.parent=15"
         _StyleDefs(967) =   "Splits(1).Columns(79).EditorStyle:id=637,.parent=17"
         _StyleDefs(968) =   "Splits(1).Columns(80).Style:id=630,.parent=13"
         _StyleDefs(969) =   "Splits(1).Columns(80).HeadingStyle:id=627,.parent=14"
         _StyleDefs(970) =   "Splits(1).Columns(80).FooterStyle:id=628,.parent=15"
         _StyleDefs(971) =   "Splits(1).Columns(80).EditorStyle:id=629,.parent=17"
         _StyleDefs(972) =   "Splits(1).Columns(81).Style:id=622,.parent=13"
         _StyleDefs(973) =   "Splits(1).Columns(81).HeadingStyle:id=619,.parent=14"
         _StyleDefs(974) =   "Splits(1).Columns(81).FooterStyle:id=620,.parent=15"
         _StyleDefs(975) =   "Splits(1).Columns(81).EditorStyle:id=621,.parent=17"
         _StyleDefs(976) =   "Splits(1).Columns(82).Style:id=390,.parent=13"
         _StyleDefs(977) =   "Splits(1).Columns(82).HeadingStyle:id=387,.parent=14"
         _StyleDefs(978) =   "Splits(1).Columns(82).FooterStyle:id=388,.parent=15"
         _StyleDefs(979) =   "Splits(1).Columns(82).EditorStyle:id=389,.parent=17"
         _StyleDefs(980) =   "Splits(1).Columns(83).Style:id=386,.parent=13"
         _StyleDefs(981) =   "Splits(1).Columns(83).HeadingStyle:id=383,.parent=14"
         _StyleDefs(982) =   "Splits(1).Columns(83).FooterStyle:id=384,.parent=15"
         _StyleDefs(983) =   "Splits(1).Columns(83).EditorStyle:id=385,.parent=17"
         _StyleDefs(984) =   "Splits(1).Columns(84).Style:id=382,.parent=13"
         _StyleDefs(985) =   "Splits(1).Columns(84).HeadingStyle:id=379,.parent=14"
         _StyleDefs(986) =   "Splits(1).Columns(84).FooterStyle:id=380,.parent=15"
         _StyleDefs(987) =   "Splits(1).Columns(84).EditorStyle:id=381,.parent=17"
         _StyleDefs(988) =   "Splits(1).Columns(85).Style:id=378,.parent=13"
         _StyleDefs(989) =   "Splits(1).Columns(85).HeadingStyle:id=375,.parent=14"
         _StyleDefs(990) =   "Splits(1).Columns(85).FooterStyle:id=376,.parent=15"
         _StyleDefs(991) =   "Splits(1).Columns(85).EditorStyle:id=377,.parent=17"
         _StyleDefs(992) =   "Splits(1).Columns(86).Style:id=374,.parent=13"
         _StyleDefs(993) =   "Splits(1).Columns(86).HeadingStyle:id=371,.parent=14"
         _StyleDefs(994) =   "Splits(1).Columns(86).FooterStyle:id=372,.parent=15"
         _StyleDefs(995) =   "Splits(1).Columns(86).EditorStyle:id=373,.parent=17"
         _StyleDefs(996) =   "Splits(1).Columns(87).Style:id=370,.parent=13"
         _StyleDefs(997) =   "Splits(1).Columns(87).HeadingStyle:id=367,.parent=14"
         _StyleDefs(998) =   "Splits(1).Columns(87).FooterStyle:id=368,.parent=15"
         _StyleDefs(999) =   "Splits(1).Columns(87).EditorStyle:id=369,.parent=17"
         _StyleDefs(1000)=   "Splits(1).Columns(88).Style:id=366,.parent=13"
         _StyleDefs(1001)=   "Splits(1).Columns(88).HeadingStyle:id=363,.parent=14"
         _StyleDefs(1002)=   "Splits(1).Columns(88).FooterStyle:id=364,.parent=15"
         _StyleDefs(1003)=   "Splits(1).Columns(88).EditorStyle:id=365,.parent=17"
         _StyleDefs(1004)=   "Splits(1).Columns(89).Style:id=362,.parent=13"
         _StyleDefs(1005)=   "Splits(1).Columns(89).HeadingStyle:id=359,.parent=14"
         _StyleDefs(1006)=   "Splits(1).Columns(89).FooterStyle:id=360,.parent=15"
         _StyleDefs(1007)=   "Splits(1).Columns(89).EditorStyle:id=361,.parent=17"
         _StyleDefs(1008)=   "Splits(1).Columns(90).Style:id=358,.parent=13"
         _StyleDefs(1009)=   "Splits(1).Columns(90).HeadingStyle:id=355,.parent=14"
         _StyleDefs(1010)=   "Splits(1).Columns(90).FooterStyle:id=356,.parent=15"
         _StyleDefs(1011)=   "Splits(1).Columns(90).EditorStyle:id=357,.parent=17"
         _StyleDefs(1012)=   "Splits(1).Columns(91).Style:id=354,.parent=13"
         _StyleDefs(1013)=   "Splits(1).Columns(91).HeadingStyle:id=351,.parent=14"
         _StyleDefs(1014)=   "Splits(1).Columns(91).FooterStyle:id=352,.parent=15"
         _StyleDefs(1015)=   "Splits(1).Columns(91).EditorStyle:id=353,.parent=17"
         _StyleDefs(1016)=   "Splits(1).Columns(92).Style:id=350,.parent=13"
         _StyleDefs(1017)=   "Splits(1).Columns(92).HeadingStyle:id=347,.parent=14"
         _StyleDefs(1018)=   "Splits(1).Columns(92).FooterStyle:id=348,.parent=15"
         _StyleDefs(1019)=   "Splits(1).Columns(92).EditorStyle:id=349,.parent=17"
         _StyleDefs(1020)=   "Splits(1).Columns(93).Style:id=346,.parent=13"
         _StyleDefs(1021)=   "Splits(1).Columns(93).HeadingStyle:id=343,.parent=14"
         _StyleDefs(1022)=   "Splits(1).Columns(93).FooterStyle:id=344,.parent=15"
         _StyleDefs(1023)=   "Splits(1).Columns(93).EditorStyle:id=345,.parent=17"
         _StyleDefs(1024)=   "Splits(1).Columns(94).Style:id=342,.parent=13"
         _StyleDefs(1025)=   "Splits(1).Columns(94).HeadingStyle:id=339,.parent=14"
         _StyleDefs(1026)=   "Splits(1).Columns(94).FooterStyle:id=340,.parent=15"
         _StyleDefs(1027)=   "Splits(1).Columns(94).EditorStyle:id=341,.parent=17"
         _StyleDefs(1028)=   "Splits(1).Columns(95).Style:id=338,.parent=13"
         _StyleDefs(1029)=   "Splits(1).Columns(95).HeadingStyle:id=335,.parent=14"
         _StyleDefs(1030)=   "Splits(1).Columns(95).FooterStyle:id=336,.parent=15"
         _StyleDefs(1031)=   "Splits(1).Columns(95).EditorStyle:id=337,.parent=17"
         _StyleDefs(1032)=   "Splits(1).Columns(96).Style:id=334,.parent=13"
         _StyleDefs(1033)=   "Splits(1).Columns(96).HeadingStyle:id=331,.parent=14"
         _StyleDefs(1034)=   "Splits(1).Columns(96).FooterStyle:id=332,.parent=15"
         _StyleDefs(1035)=   "Splits(1).Columns(96).EditorStyle:id=333,.parent=17"
         _StyleDefs(1036)=   "Splits(1).Columns(97).Style:id=330,.parent=13"
         _StyleDefs(1037)=   "Splits(1).Columns(97).HeadingStyle:id=327,.parent=14"
         _StyleDefs(1038)=   "Splits(1).Columns(97).FooterStyle:id=328,.parent=15"
         _StyleDefs(1039)=   "Splits(1).Columns(97).EditorStyle:id=329,.parent=17"
         _StyleDefs(1040)=   "Splits(1).Columns(98).Style:id=326,.parent=13"
         _StyleDefs(1041)=   "Splits(1).Columns(98).HeadingStyle:id=323,.parent=14"
         _StyleDefs(1042)=   "Splits(1).Columns(98).FooterStyle:id=324,.parent=15"
         _StyleDefs(1043)=   "Splits(1).Columns(98).EditorStyle:id=325,.parent=17"
         _StyleDefs(1044)=   "Splits(1).Columns(99).Style:id=322,.parent=13"
         _StyleDefs(1045)=   "Splits(1).Columns(99).HeadingStyle:id=319,.parent=14"
         _StyleDefs(1046)=   "Splits(1).Columns(99).FooterStyle:id=320,.parent=15"
         _StyleDefs(1047)=   "Splits(1).Columns(99).EditorStyle:id=321,.parent=17"
         _StyleDefs(1048)=   "Splits(1).Columns(100).Style:id=318,.parent=13"
         _StyleDefs(1049)=   "Splits(1).Columns(100).HeadingStyle:id=315,.parent=14"
         _StyleDefs(1050)=   "Splits(1).Columns(100).FooterStyle:id=316,.parent=15"
         _StyleDefs(1051)=   "Splits(1).Columns(100).EditorStyle:id=317,.parent=17"
         _StyleDefs(1052)=   "Splits(1).Columns(101).Style:id=314,.parent=13"
         _StyleDefs(1053)=   "Splits(1).Columns(101).HeadingStyle:id=311,.parent=14"
         _StyleDefs(1054)=   "Splits(1).Columns(101).FooterStyle:id=312,.parent=15"
         _StyleDefs(1055)=   "Splits(1).Columns(101).EditorStyle:id=313,.parent=17"
         _StyleDefs(1056)=   "Splits(1).Columns(102).Style:id=306,.parent=13"
         _StyleDefs(1057)=   "Splits(1).Columns(102).HeadingStyle:id=303,.parent=14"
         _StyleDefs(1058)=   "Splits(1).Columns(102).FooterStyle:id=304,.parent=15"
         _StyleDefs(1059)=   "Splits(1).Columns(102).EditorStyle:id=305,.parent=17"
         _StyleDefs(1060)=   "Splits(1).Columns(103).Style:id=302,.parent=13"
         _StyleDefs(1061)=   "Splits(1).Columns(103).HeadingStyle:id=299,.parent=14"
         _StyleDefs(1062)=   "Splits(1).Columns(103).FooterStyle:id=300,.parent=15"
         _StyleDefs(1063)=   "Splits(1).Columns(103).EditorStyle:id=301,.parent=17"
         _StyleDefs(1064)=   "Splits(1).Columns(104).Style:id=298,.parent=13"
         _StyleDefs(1065)=   "Splits(1).Columns(104).HeadingStyle:id=295,.parent=14"
         _StyleDefs(1066)=   "Splits(1).Columns(104).FooterStyle:id=296,.parent=15"
         _StyleDefs(1067)=   "Splits(1).Columns(104).EditorStyle:id=297,.parent=17"
         _StyleDefs(1068)=   "Splits(1).Columns(105).Style:id=294,.parent=13"
         _StyleDefs(1069)=   "Splits(1).Columns(105).HeadingStyle:id=291,.parent=14"
         _StyleDefs(1070)=   "Splits(1).Columns(105).FooterStyle:id=292,.parent=15"
         _StyleDefs(1071)=   "Splits(1).Columns(105).EditorStyle:id=293,.parent=17"
         _StyleDefs(1072)=   "Splits(1).Columns(106).Style:id=290,.parent=13"
         _StyleDefs(1073)=   "Splits(1).Columns(106).HeadingStyle:id=287,.parent=14"
         _StyleDefs(1074)=   "Splits(1).Columns(106).FooterStyle:id=288,.parent=15"
         _StyleDefs(1075)=   "Splits(1).Columns(106).EditorStyle:id=289,.parent=17"
         _StyleDefs(1076)=   "Splits(1).Columns(107).Style:id=286,.parent=13"
         _StyleDefs(1077)=   "Splits(1).Columns(107).HeadingStyle:id=283,.parent=14"
         _StyleDefs(1078)=   "Splits(1).Columns(107).FooterStyle:id=284,.parent=15"
         _StyleDefs(1079)=   "Splits(1).Columns(107).EditorStyle:id=285,.parent=17"
         _StyleDefs(1080)=   "Splits(1).Columns(108).Style:id=282,.parent=13"
         _StyleDefs(1081)=   "Splits(1).Columns(108).HeadingStyle:id=279,.parent=14"
         _StyleDefs(1082)=   "Splits(1).Columns(108).FooterStyle:id=280,.parent=15"
         _StyleDefs(1083)=   "Splits(1).Columns(108).EditorStyle:id=281,.parent=17"
         _StyleDefs(1084)=   "Splits(1).Columns(109).Style:id=278,.parent=13"
         _StyleDefs(1085)=   "Splits(1).Columns(109).HeadingStyle:id=275,.parent=14"
         _StyleDefs(1086)=   "Splits(1).Columns(109).FooterStyle:id=276,.parent=15"
         _StyleDefs(1087)=   "Splits(1).Columns(109).EditorStyle:id=277,.parent=17"
         _StyleDefs(1088)=   "Splits(1).Columns(110).Style:id=274,.parent=13"
         _StyleDefs(1089)=   "Splits(1).Columns(110).HeadingStyle:id=271,.parent=14"
         _StyleDefs(1090)=   "Splits(1).Columns(110).FooterStyle:id=272,.parent=15"
         _StyleDefs(1091)=   "Splits(1).Columns(110).EditorStyle:id=273,.parent=17"
         _StyleDefs(1092)=   "Splits(1).Columns(111).Style:id=270,.parent=13"
         _StyleDefs(1093)=   "Splits(1).Columns(111).HeadingStyle:id=267,.parent=14"
         _StyleDefs(1094)=   "Splits(1).Columns(111).FooterStyle:id=268,.parent=15"
         _StyleDefs(1095)=   "Splits(1).Columns(111).EditorStyle:id=269,.parent=17"
         _StyleDefs(1096)=   "Splits(1).Columns(112).Style:id=266,.parent=13"
         _StyleDefs(1097)=   "Splits(1).Columns(112).HeadingStyle:id=263,.parent=14"
         _StyleDefs(1098)=   "Splits(1).Columns(112).FooterStyle:id=264,.parent=15"
         _StyleDefs(1099)=   "Splits(1).Columns(112).EditorStyle:id=265,.parent=17"
         _StyleDefs(1100)=   "Splits(1).Columns(113).Style:id=262,.parent=13"
         _StyleDefs(1101)=   "Splits(1).Columns(113).HeadingStyle:id=259,.parent=14"
         _StyleDefs(1102)=   "Splits(1).Columns(113).FooterStyle:id=260,.parent=15"
         _StyleDefs(1103)=   "Splits(1).Columns(113).EditorStyle:id=261,.parent=17"
         _StyleDefs(1104)=   "Splits(1).Columns(114).Style:id=258,.parent=13"
         _StyleDefs(1105)=   "Splits(1).Columns(114).HeadingStyle:id=255,.parent=14"
         _StyleDefs(1106)=   "Splits(1).Columns(114).FooterStyle:id=256,.parent=15"
         _StyleDefs(1107)=   "Splits(1).Columns(114).EditorStyle:id=257,.parent=17"
         _StyleDefs(1108)=   "Splits(1).Columns(115).Style:id=254,.parent=13"
         _StyleDefs(1109)=   "Splits(1).Columns(115).HeadingStyle:id=251,.parent=14"
         _StyleDefs(1110)=   "Splits(1).Columns(115).FooterStyle:id=252,.parent=15"
         _StyleDefs(1111)=   "Splits(1).Columns(115).EditorStyle:id=253,.parent=17"
         _StyleDefs(1112)=   "Splits(1).Columns(116).Style:id=250,.parent=13"
         _StyleDefs(1113)=   "Splits(1).Columns(116).HeadingStyle:id=247,.parent=14"
         _StyleDefs(1114)=   "Splits(1).Columns(116).FooterStyle:id=248,.parent=15"
         _StyleDefs(1115)=   "Splits(1).Columns(116).EditorStyle:id=249,.parent=17"
         _StyleDefs(1116)=   "Splits(1).Columns(117).Style:id=246,.parent=13"
         _StyleDefs(1117)=   "Splits(1).Columns(117).HeadingStyle:id=243,.parent=14"
         _StyleDefs(1118)=   "Splits(1).Columns(117).FooterStyle:id=244,.parent=15"
         _StyleDefs(1119)=   "Splits(1).Columns(117).EditorStyle:id=245,.parent=17"
         _StyleDefs(1120)=   "Splits(1).Columns(118).Style:id=242,.parent=13"
         _StyleDefs(1121)=   "Splits(1).Columns(118).HeadingStyle:id=239,.parent=14"
         _StyleDefs(1122)=   "Splits(1).Columns(118).FooterStyle:id=240,.parent=15"
         _StyleDefs(1123)=   "Splits(1).Columns(118).EditorStyle:id=241,.parent=17"
         _StyleDefs(1124)=   "Splits(1).Columns(119).Style:id=238,.parent=13"
         _StyleDefs(1125)=   "Splits(1).Columns(119).HeadingStyle:id=235,.parent=14"
         _StyleDefs(1126)=   "Splits(1).Columns(119).FooterStyle:id=236,.parent=15"
         _StyleDefs(1127)=   "Splits(1).Columns(119).EditorStyle:id=237,.parent=17"
         _StyleDefs(1128)=   "Splits(1).Columns(120).Style:id=234,.parent=13"
         _StyleDefs(1129)=   "Splits(1).Columns(120).HeadingStyle:id=231,.parent=14"
         _StyleDefs(1130)=   "Splits(1).Columns(120).FooterStyle:id=232,.parent=15"
         _StyleDefs(1131)=   "Splits(1).Columns(120).EditorStyle:id=233,.parent=17"
         _StyleDefs(1132)=   "Splits(1).Columns(121).Style:id=230,.parent=13"
         _StyleDefs(1133)=   "Splits(1).Columns(121).HeadingStyle:id=227,.parent=14"
         _StyleDefs(1134)=   "Splits(1).Columns(121).FooterStyle:id=228,.parent=15"
         _StyleDefs(1135)=   "Splits(1).Columns(121).EditorStyle:id=229,.parent=17"
         _StyleDefs(1136)=   "Splits(1).Columns(122).Style:id=226,.parent=13"
         _StyleDefs(1137)=   "Splits(1).Columns(122).HeadingStyle:id=223,.parent=14"
         _StyleDefs(1138)=   "Splits(1).Columns(122).FooterStyle:id=224,.parent=15"
         _StyleDefs(1139)=   "Splits(1).Columns(122).EditorStyle:id=225,.parent=17"
         _StyleDefs(1140)=   "Splits(1).Columns(123).Style:id=222,.parent=13"
         _StyleDefs(1141)=   "Splits(1).Columns(123).HeadingStyle:id=219,.parent=14"
         _StyleDefs(1142)=   "Splits(1).Columns(123).FooterStyle:id=220,.parent=15"
         _StyleDefs(1143)=   "Splits(1).Columns(123).EditorStyle:id=221,.parent=17"
         _StyleDefs(1144)=   "Splits(1).Columns(124).Style:id=218,.parent=13"
         _StyleDefs(1145)=   "Splits(1).Columns(124).HeadingStyle:id=215,.parent=14"
         _StyleDefs(1146)=   "Splits(1).Columns(124).FooterStyle:id=216,.parent=15"
         _StyleDefs(1147)=   "Splits(1).Columns(124).EditorStyle:id=217,.parent=17"
         _StyleDefs(1148)=   "Splits(1).Columns(125).Style:id=214,.parent=13"
         _StyleDefs(1149)=   "Splits(1).Columns(125).HeadingStyle:id=211,.parent=14"
         _StyleDefs(1150)=   "Splits(1).Columns(125).FooterStyle:id=212,.parent=15"
         _StyleDefs(1151)=   "Splits(1).Columns(125).EditorStyle:id=213,.parent=17"
         _StyleDefs(1152)=   "Splits(1).Columns(126).Style:id=210,.parent=13"
         _StyleDefs(1153)=   "Splits(1).Columns(126).HeadingStyle:id=207,.parent=14"
         _StyleDefs(1154)=   "Splits(1).Columns(126).FooterStyle:id=208,.parent=15"
         _StyleDefs(1155)=   "Splits(1).Columns(126).EditorStyle:id=209,.parent=17"
         _StyleDefs(1156)=   "Splits(1).Columns(127).Style:id=206,.parent=13"
         _StyleDefs(1157)=   "Splits(1).Columns(127).HeadingStyle:id=203,.parent=14"
         _StyleDefs(1158)=   "Splits(1).Columns(127).FooterStyle:id=204,.parent=15"
         _StyleDefs(1159)=   "Splits(1).Columns(127).EditorStyle:id=205,.parent=17"
         _StyleDefs(1160)=   "Splits(1).Columns(128).Style:id=202,.parent=13"
         _StyleDefs(1161)=   "Splits(1).Columns(128).HeadingStyle:id=199,.parent=14"
         _StyleDefs(1162)=   "Splits(1).Columns(128).FooterStyle:id=200,.parent=15"
         _StyleDefs(1163)=   "Splits(1).Columns(128).EditorStyle:id=201,.parent=17"
         _StyleDefs(1164)=   "Splits(1).Columns(129).Style:id=198,.parent=13"
         _StyleDefs(1165)=   "Splits(1).Columns(129).HeadingStyle:id=195,.parent=14"
         _StyleDefs(1166)=   "Splits(1).Columns(129).FooterStyle:id=196,.parent=15"
         _StyleDefs(1167)=   "Splits(1).Columns(129).EditorStyle:id=197,.parent=17"
         _StyleDefs(1168)=   "Splits(1).Columns(130).Style:id=194,.parent=13"
         _StyleDefs(1169)=   "Splits(1).Columns(130).HeadingStyle:id=191,.parent=14"
         _StyleDefs(1170)=   "Splits(1).Columns(130).FooterStyle:id=192,.parent=15"
         _StyleDefs(1171)=   "Splits(1).Columns(130).EditorStyle:id=193,.parent=17"
         _StyleDefs(1172)=   "Splits(1).Columns(131).Style:id=190,.parent=13"
         _StyleDefs(1173)=   "Splits(1).Columns(131).HeadingStyle:id=187,.parent=14"
         _StyleDefs(1174)=   "Splits(1).Columns(131).FooterStyle:id=188,.parent=15"
         _StyleDefs(1175)=   "Splits(1).Columns(131).EditorStyle:id=189,.parent=17"
         _StyleDefs(1176)=   "Splits(1).Columns(132).Style:id=186,.parent=13"
         _StyleDefs(1177)=   "Splits(1).Columns(132).HeadingStyle:id=183,.parent=14"
         _StyleDefs(1178)=   "Splits(1).Columns(132).FooterStyle:id=184,.parent=15"
         _StyleDefs(1179)=   "Splits(1).Columns(132).EditorStyle:id=185,.parent=17"
         _StyleDefs(1180)=   "Splits(1).Columns(133).Style:id=182,.parent=13"
         _StyleDefs(1181)=   "Splits(1).Columns(133).HeadingStyle:id=179,.parent=14"
         _StyleDefs(1182)=   "Splits(1).Columns(133).FooterStyle:id=180,.parent=15"
         _StyleDefs(1183)=   "Splits(1).Columns(133).EditorStyle:id=181,.parent=17"
         _StyleDefs(1184)=   "Splits(1).Columns(134).Style:id=178,.parent=13"
         _StyleDefs(1185)=   "Splits(1).Columns(134).HeadingStyle:id=175,.parent=14"
         _StyleDefs(1186)=   "Splits(1).Columns(134).FooterStyle:id=176,.parent=15"
         _StyleDefs(1187)=   "Splits(1).Columns(134).EditorStyle:id=177,.parent=17"
         _StyleDefs(1188)=   "Splits(1).Columns(135).Style:id=174,.parent=13"
         _StyleDefs(1189)=   "Splits(1).Columns(135).HeadingStyle:id=171,.parent=14"
         _StyleDefs(1190)=   "Splits(1).Columns(135).FooterStyle:id=172,.parent=15"
         _StyleDefs(1191)=   "Splits(1).Columns(135).EditorStyle:id=173,.parent=17"
         _StyleDefs(1192)=   "Splits(1).Columns(136).Style:id=170,.parent=13"
         _StyleDefs(1193)=   "Splits(1).Columns(136).HeadingStyle:id=167,.parent=14"
         _StyleDefs(1194)=   "Splits(1).Columns(136).FooterStyle:id=168,.parent=15"
         _StyleDefs(1195)=   "Splits(1).Columns(136).EditorStyle:id=169,.parent=17"
         _StyleDefs(1196)=   "Splits(1).Columns(137).Style:id=166,.parent=13"
         _StyleDefs(1197)=   "Splits(1).Columns(137).HeadingStyle:id=163,.parent=14"
         _StyleDefs(1198)=   "Splits(1).Columns(137).FooterStyle:id=164,.parent=15"
         _StyleDefs(1199)=   "Splits(1).Columns(137).EditorStyle:id=165,.parent=17"
         _StyleDefs(1200)=   "Splits(1).Columns(138).Style:id=162,.parent=13"
         _StyleDefs(1201)=   "Splits(1).Columns(138).HeadingStyle:id=159,.parent=14"
         _StyleDefs(1202)=   "Splits(1).Columns(138).FooterStyle:id=160,.parent=15"
         _StyleDefs(1203)=   "Splits(1).Columns(138).EditorStyle:id=161,.parent=17"
         _StyleDefs(1204)=   "Splits(1).Columns(139).Style:id=158,.parent=13"
         _StyleDefs(1205)=   "Splits(1).Columns(139).HeadingStyle:id=155,.parent=14"
         _StyleDefs(1206)=   "Splits(1).Columns(139).FooterStyle:id=156,.parent=15"
         _StyleDefs(1207)=   "Splits(1).Columns(139).EditorStyle:id=157,.parent=17"
         _StyleDefs(1208)=   "Splits(1).Columns(140).Style:id=154,.parent=13"
         _StyleDefs(1209)=   "Splits(1).Columns(140).HeadingStyle:id=151,.parent=14"
         _StyleDefs(1210)=   "Splits(1).Columns(140).FooterStyle:id=152,.parent=15"
         _StyleDefs(1211)=   "Splits(1).Columns(140).EditorStyle:id=153,.parent=17"
         _StyleDefs(1212)=   "Splits(1).Columns(141).Style:id=150,.parent=13"
         _StyleDefs(1213)=   "Splits(1).Columns(141).HeadingStyle:id=147,.parent=14"
         _StyleDefs(1214)=   "Splits(1).Columns(141).FooterStyle:id=148,.parent=15"
         _StyleDefs(1215)=   "Splits(1).Columns(141).EditorStyle:id=149,.parent=17"
         _StyleDefs(1216)=   "Splits(1).Columns(142).Style:id=146,.parent=13"
         _StyleDefs(1217)=   "Splits(1).Columns(142).HeadingStyle:id=143,.parent=14"
         _StyleDefs(1218)=   "Splits(1).Columns(142).FooterStyle:id=144,.parent=15"
         _StyleDefs(1219)=   "Splits(1).Columns(142).EditorStyle:id=145,.parent=17"
         _StyleDefs(1220)=   "Splits(1).Columns(143).Style:id=142,.parent=13"
         _StyleDefs(1221)=   "Splits(1).Columns(143).HeadingStyle:id=139,.parent=14"
         _StyleDefs(1222)=   "Splits(1).Columns(143).FooterStyle:id=140,.parent=15"
         _StyleDefs(1223)=   "Splits(1).Columns(143).EditorStyle:id=141,.parent=17"
         _StyleDefs(1224)=   "Splits(1).Columns(144).Style:id=138,.parent=13"
         _StyleDefs(1225)=   "Splits(1).Columns(144).HeadingStyle:id=135,.parent=14"
         _StyleDefs(1226)=   "Splits(1).Columns(144).FooterStyle:id=136,.parent=15"
         _StyleDefs(1227)=   "Splits(1).Columns(144).EditorStyle:id=137,.parent=17"
         _StyleDefs(1228)=   "Splits(1).Columns(145).Style:id=134,.parent=13"
         _StyleDefs(1229)=   "Splits(1).Columns(145).HeadingStyle:id=131,.parent=14"
         _StyleDefs(1230)=   "Splits(1).Columns(145).FooterStyle:id=132,.parent=15"
         _StyleDefs(1231)=   "Splits(1).Columns(145).EditorStyle:id=133,.parent=17"
         _StyleDefs(1232)=   "Splits(1).Columns(146).Style:id=130,.parent=13"
         _StyleDefs(1233)=   "Splits(1).Columns(146).HeadingStyle:id=127,.parent=14"
         _StyleDefs(1234)=   "Splits(1).Columns(146).FooterStyle:id=128,.parent=15"
         _StyleDefs(1235)=   "Splits(1).Columns(146).EditorStyle:id=129,.parent=17"
         _StyleDefs(1236)=   "Splits(1).Columns(147).Style:id=126,.parent=13"
         _StyleDefs(1237)=   "Splits(1).Columns(147).HeadingStyle:id=123,.parent=14"
         _StyleDefs(1238)=   "Splits(1).Columns(147).FooterStyle:id=124,.parent=15"
         _StyleDefs(1239)=   "Splits(1).Columns(147).EditorStyle:id=125,.parent=17"
         _StyleDefs(1240)=   "Splits(1).Columns(148).Style:id=122,.parent=13"
         _StyleDefs(1241)=   "Splits(1).Columns(148).HeadingStyle:id=119,.parent=14"
         _StyleDefs(1242)=   "Splits(1).Columns(148).FooterStyle:id=120,.parent=15"
         _StyleDefs(1243)=   "Splits(1).Columns(148).EditorStyle:id=121,.parent=17"
         _StyleDefs(1244)=   "Splits(1).Columns(149).Style:id=118,.parent=13"
         _StyleDefs(1245)=   "Splits(1).Columns(149).HeadingStyle:id=115,.parent=14"
         _StyleDefs(1246)=   "Splits(1).Columns(149).FooterStyle:id=116,.parent=15"
         _StyleDefs(1247)=   "Splits(1).Columns(149).EditorStyle:id=117,.parent=17"
         _StyleDefs(1248)=   "Splits(1).Columns(150).Style:id=114,.parent=13"
         _StyleDefs(1249)=   "Splits(1).Columns(150).HeadingStyle:id=111,.parent=14"
         _StyleDefs(1250)=   "Splits(1).Columns(150).FooterStyle:id=112,.parent=15"
         _StyleDefs(1251)=   "Splits(1).Columns(150).EditorStyle:id=113,.parent=17"
         _StyleDefs(1252)=   "Named:id=33:Normal"
         _StyleDefs(1253)=   ":id=33,.parent=0"
         _StyleDefs(1254)=   "Named:id=34:Heading"
         _StyleDefs(1255)=   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(1256)=   ":id=34,.wraptext=-1"
         _StyleDefs(1257)=   "Named:id=35:Footing"
         _StyleDefs(1258)=   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(1259)=   "Named:id=36:Selected"
         _StyleDefs(1260)=   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(1261)=   "Named:id=37:Caption"
         _StyleDefs(1262)=   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(1263)=   "Named:id=38:HighlightRow"
         _StyleDefs(1264)=   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(1265)=   "Named:id=39:EvenRow"
         _StyleDefs(1266)=   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(1267)=   "Named:id=40:OddRow"
         _StyleDefs(1268)=   ":id=40,.parent=33"
         _StyleDefs(1269)=   "Named:id=41:RecordSelector"
         _StyleDefs(1270)=   ":id=41,.parent=34"
         _StyleDefs(1271)=   "Named:id=42:FilterBar"
         _StyleDefs(1272)=   ":id=42,.parent=33"
      End
      Begin VB.Frame F4 
         Height          =   1215
         Left            =   1200
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   12855
      Begin VB.CheckBox chEM 
         BackColor       =   &H00FFFFFF&
         Caption         =   "EM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11880
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11880
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chWar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WARRANTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   36
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkThs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "THS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   1065
         TabIndex        =   12
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CC-CCCC-CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   6075
         TabIndex        =   0
         Top             =   840
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   6075
         TabIndex        =   15
         Top             =   540
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1065
         TabIndex        =   17
         Top             =   600
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   6075
         TabIndex        =   20
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   1065
         TabIndex        =   24
         Top             =   960
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   6960
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
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
         Index           =   10
         Left            =   7080
         TabIndex        =   26
         Top             =   840
         Visible         =   0   'False
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   10080
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sisa Menit"
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
         Left            =   7710
         TabIndex        =   27
         Top             =   885
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label26 
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
         Left            =   5355
         TabIndex        =   22
         Top             =   255
         Width           =   570
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "(DD-MM-YYYY)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7350
         TabIndex        =   21
         Top             =   255
         Width           =   1140
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   510
         TabIndex        =   19
         Top             =   945
         Width           =   405
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Lgn"
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
         Left            =   255
         TabIndex        =   18
         Top             =   645
         Width           =   660
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
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
         Left            =   5580
         TabIndex        =   16
         Top             =   585
         Width           =   345
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MCH"
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
         Left            =   5595
         TabIndex        =   14
         Top             =   885
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
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
         TabIndex        =   13
         Top             =   285
         Width           =   645
      End
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8415
      Left            =   0
      Picture         =   "frmStallNonEM.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmStallNonEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 15 Januari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim jam As Integer
Dim jamskr As Integer
Dim menit As Integer
Dim menitskr As Integer
Dim row As Integer
Dim col As Integer
'Dim C As TrueOleDBGrid70.Columns
Dim rsStall As New ADODB.Recordset
Dim njamAk As Integer
Dim njamAwl As Integer
Dim jmlStall As Integer
Dim jmlTHS As Integer
Dim jmlAll As Integer
Dim nTgl As String
Dim batasAwal As Integer        ' POTONGAN WAKTU AWAL YG GAK 1 JAM
Dim xy As Boolean
Dim jamAwal As String
    
Const NilaiAwal = 1
Const AkhirPos = 150
Dim rsTemp As ADODB.Recordset
Dim rsHlp As ADODB.Recordset

Private Sub clearStall()
    For i = NilaiAwal To AkhirPos
        With gridStall.Splits(1)
            .Columns(i).Visible = True
        End With
    Next
End Sub

Private Sub SetWaktu(jamAwl As Integer, jamAkhir As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim nilaiAW As Integer
    
    Call clearStall
    j = 0
    batasAwal = IIf(jamAwl Mod 60 = 0, 6, Round((60 - (jamAwl Mod 60)) / 10)) '+1
    For i = NilaiAwal To ((jamAkhir - jamAwl) / 10) + IIf((jamAwl Mod 60) Mod 2 = 1, 3, 2)
        nilaiAW = NilaiAwal - (6 - batasAwal)
        Select Case i
            Case Is <= batasAwal '+ 1
                If i = NilaiAwal + (j * batasAwal) Then
                    'judul = Val(jamAwl / 60) + j
                    judul = Val(Round(CDbl(jamAwl / 60), 0)) + j
                    j = j + 1
                Else
                    judul = "."
                End If
            Case Else
                If i = nilaiAW + (j * 6) Then   '+ (1 * batasAwal)+ ((j - 1) * 6)
                    'judul = Val(jamAwl / 60) + j
                    judul = Val(Round(CDbl(jamAwl / 60), 0)) + j
                    j = j + 1
                Else
                    judul = "."
                End If
        End Select

        gridStall.Splits(1).Columns(i).Caption = judul
        gridStall.Splits(1).Columns(i).Width = 500
        gridStall.Splits(1).Columns(i).Alignment = dbgCenter
    Next
End Sub

Private Sub setStall(jamAwl As Integer, jamAkhir As Integer)

    gridStall.RowHeight = 400
    gridStall.Style.Font.Bold = True
    gridStall.Style.Font.Size = 14
    ' JML STALL
    jmlStall = objO.FetchData_Integer("SELECT FN_JML_STALL FROM T_PROFIL")
    
    ' UNIT THS
    jmlTHS = objO.FetchData_Integer("SELECT FN_JML_THS FROM T_PROFIL")
    
    jmlAll = jmlStall + jmlTHS
    'MyArray3.ReDim 1, jmlAll + 2, 0, ((jamAkhir - jamAwl) * 6)
    MyArray3.ReDim 1, jmlAll + 2, 0, ((jamAkhir - jamAwl) / 10) + 2
    MyArray3.Clear
    For i = 1 To jmlAll + 2
        With gridStall
            If i < jmlStall + 1 Then
                MyArray3.value(i, 0) = "Stall " & i
            ElseIf (i - jmlStall - 1) < jmlTHS + 1 And (i - jmlStall - 1) > 0 Then
                MyArray3.value(i, 0) = "THS " & (i - jmlStall - 1)
            End If
        End With
    Next
    
    gridStall.Array = MyArray3
End Sub

Private Sub loadStall(nTanggal As String, jamAwl As Integer, jamAk As Integer)
    Dim rsAStl As ADODB.Recordset
    Dim ProgressItem As New TrueDBGrid70.Style
    Dim ProgressUnit As New TrueDBGrid70.Style
    Dim tnd As String
    Dim nStl As String
    Dim rteBaru As Double
    
    gridStall.EvenRowStyle.BackColor = &H80000005
    gridStall.OddRowStyle.BackColor = &H80000005
    
    ProgressUnit.BackColor = vbCyan
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressUnit, "^>"
    
    ProgressItem.BackColor = vbGreen
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressItem, "^#"
    
    gridStall.Splits(1).AlternatingRowStyle = True
    
    Call clearMyarray(njamAwl, njamAk)
    
    Set rsAStl = objO.FetchData("SELECT * FROM stall WITH (NOLOCK) WHERE tanggal = '" & nTanggal & "' AND FLAG <> 'X' ORDER BY noStall")
    If Not rsAStl.EOF Then
        Do While Not rsAStl.EOF
            If Trim(Left(Trim(rsAStl!keterangan), InStr(rsAStl!keterangan, "|") - 1)) = Trim(Replace(Replace(Trim(text(4)), ">", ""), " ", "")) Then
                tnd = "#"
            Else
                tnd = "#"
            End If
            If Len(Trim(rsAStl!noStall)) > 8 Then
                  
            Else
                If Len(Trim(rsAStl!noStall)) > 6 Then
                    nStl = CDbl(Mid(Trim(rsAStl("nostall")), 7, Len(Trim(rsAStl("nostall"))) - 6)) '+ 1
                Else
                    nStl = CDbl(Mid(Trim(rsAStl("nostall")), 5, Len(Trim(rsAStl("nostall"))) - 4)) + jmlStall + 1
                End If
                'Call SetBook(Left(Trim(rsAStl!jamMulai), 2), Right(Trim(rsAStl!jamMulai), 2), CDbl(rsAStl!Rate * 6), CDbl(nStl), tnd)
                
                If CDbl(rsAStl!Rate * 60) > CDbl(jamAk - hitungMenit(Format(Trim(rsAStl("JAMMULAI")), "HH:NN"))) Then
                    rteBaru = CDbl(jamAk - hitungMenit(Format(Trim(rsAStl("JAMMULAI")), "HH:NN"))) / 10
                Else
                    rteBaru = CDbl(rsAStl!Rate * 6)
                End If
            End If
            
            Call SetBook(Left(Trim(rsAStl!jamMulai), 2), Right(Trim(rsAStl!jamMulai), 2), rteBaru, CDbl(nStl), tnd)
            rsAStl.MoveNext
        Loop
    End If
    
    Set rsAStl = objO.FetchData("SELECT * FROM tempstall WHERE tanggal = '" & nTanggal & "' AND FLAG <> 'X' ORDER BY noStall")
    If Not rsAStl.EOF Then
        Do While Not rsAStl.EOF
            If Trim(Left(Trim(rsAStl!keterangan), InStr(rsAStl!keterangan, "|") - 1)) = Trim(Replace(Replace(Trim(text(4)), ">", ""), " ", "")) Then
                tnd = ">"
                jamAwal = Trim(rsAStl!jamMulai)
                frmBooking.text(58) = Trim(rsAStl!jamMulai)
            Else
                tnd = "#"
            End If
            
            If Len(Trim(rsAStl!noStall)) > 6 Then
                nStl = CDbl(Mid(Trim(rsAStl("nostall")), 7, Len(Trim(rsAStl("nostall"))) - 6))  ' + 1
            Else
                nStl = CDbl(Mid(Trim(rsAStl("nostall")), 5, Len(Trim(rsAStl("nostall"))) - 4)) + jmlStall + 1
            End If
            If CDbl(rsAStl!Rate * 60) > CDbl(jamAk - hitungMenit(Format(Trim(rsAStl("JAMMULAI")), "HH:NN"))) Then
                rteBaru = CDbl(jamAk - hitungMenit(Format(Trim(rsAStl("JAMMULAI")), "HH:NN"))) / 10
            Else
                rteBaru = CDbl(rsAStl!Rate * 6)
            End If
            Call SetBook(Left(Trim(rsAStl!jamMulai), 2), Right(Trim(rsAStl!jamMulai), 2), Val(rteBaru), CDbl(nStl), tnd)
            rsAStl.MoveNext
        Loop
    End If
End Sub

Private Sub SettingGrid(jamAwl As Integer, jamAkhir As Integer)
Dim nAwal As Integer
    nAwal = (((jamAkhir - jamAwl) / 10) + 2)
    For i = nAwal To AkhirPos
        gridStall.Splits(1).Columns(i).Visible = False
    Next
End Sub

Private Sub SetBook(ByVal JamAw As Integer, ByVal MenitAw As Integer, ByVal Panjang As Integer, ByVal Pos As Integer, sig As String)
    On Error GoTo err
    
    Dim i As Integer
    Dim PosisiKotak As Integer
    Dim PanjangKotak As Integer
    Dim TempPanjang As Integer
    
    If JamAw >= (njamAwl / 60) Or JamAw <= (njamAk / 60) Then
        
        PosisiKotak = Round(((((JamAw * 60) + MenitAw) - (njamAwl - (njamAwl Mod 10))) / 10) + 0.01, 0) + 1
        PanjangKotak = Panjang
        PosisiKotak = IIf(PosisiKotak = 0, 1, PosisiKotak)
        
        For i = 0 To PanjangKotak
            MyArray3.value(Pos, PosisiKotak + i) = sig
        Next
        gridStall.Array = MyArray3
    End If
    Exit Sub
    
err:
    MsgBox "Garis Flate Rate terlalu panjang !", vbOKOnly, "Error"
End Sub

Private Sub cmdStalBatal_Click()
    frmStal.Visible = False
End Sub

Private Sub cmdStalOk_Click()
      Dim rsHlp2 As ADODB.Recordset
          
1         If gStall = "booking" Then
2             If Format(text(3), "HH:NN") < Format(frmBooking.text(24), "HH:NN") Then
3                 MsgBox "Jam Stall tidak boleh lebih kecil dari jam kedatangan !!" & vbCr & "Jam Kedatangan : " & frmBooking.text(24), vbInformation, "JAM STALL"
4                 frmStal.Visible = False
5                 Unload Me
6                 Exit Sub
7             End If
8         End If
          
              '' CEK UNIT TUMPUK
9         Set rsHlp = objO.FetchData("SELECT KETERANGAN FROM STALL WITH (NOLOCK) WHERE FLAG <> 'X' AND TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "' AND KETERANGAN NOT LIKE '%" & Trim(text(0)) & "%' AND NOSTALL = '" & Label2 & Trim(text(2)) & "' AND (JAMMULAI BETWEEN '" & Trim(text(3)) & "' AND '" & Trim(text(11)) & "' OR JAMSELESAI BETWEEN '" & Trim(text(3)) & "' AND '" & Trim(text(11)) & "')")
10        If Not rsHlp.EOF Then
11            MsgBox "Stall digunakan unit kendaraan lain. " & vbCr & "Kendaraan : " & Left(rsHlp("KETERANGAN"), InStr(rsHlp("KETERANGAN"), " |")) & vbCr & vbCr & "Pilih Stall lain"
12            frmStal.Visible = False
13            Unload Me
14            Exit Sub
15        Else
               
16            Set rsHlp2 = objO.FetchData("SELECT KETERANGAN FROM TEMPSTALL WITH (NOLOCK) WHERE TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "' AND KETERANGAN NOT LIKE '%" & Trim(text(0)) & "%' AND NOSTALL = '" & Label2 & Trim(text(2)) & "' AND (JAMMULAI BETWEEN '" & Trim(text(3)) & "' AND '" & Trim(text(11)) & "' OR JAMSELESAI BETWEEN '" & Trim(text(3)) & "' AND '" & Trim(text(11)) & "')")
17            If Not rsHlp2.EOF Then
18                MsgBox "Stall digunakan unit kendaraan lain. " & vbCr & "Kendaraan : " & Left(rsHlp2("KETERANGAN"), InStr(rsHlp2("KETERANGAN"), " |")) & vbCr & vbCr & "Pilih Stall lain"
19                frmStal.Visible = False
20                Unload Me
21                Exit Sub
22            End If
23            If rsHlp2.State <> adStateClosed Then rsHlp2.Close
24        End If
           
25        Set rsHlp = objO.FetchData("SELECT KETERANGAN FROM STALL WITH (NOLOCK) WHERE FLAG <> 'X' AND TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "' AND KETERANGAN NOT LIKE '%" & Trim(text(0)) & "%' AND NOSTALL = '" & Label2 & Trim(text(2)) & "' AND ((JAMMULAI <= '" & Trim(text(3)) & "' AND JAMSELESAI >= '" & Trim(text(3)) & "') OR (JAMMULAI <= '" & Trim(text(11)) & "' AND JAMSELESAI >= '" & Trim(text(11)) & "'))")
26        If Not rsHlp.EOF Then
27            MsgBox "Stall digunakan unit kendaraan lain. " & vbCr & "Kendaraan : " & Left(rsHlp("KETERANGAN"), InStr(rsHlp("KETERANGAN"), " |")) & vbCr & vbCr & "Pilih Stall lain"
28            frmStal.Visible = False
29            Unload Me
30            Exit Sub
31        Else
32            Set rsHlp2 = objO.FetchData("SELECT KETERANGAN FROM TEMPSTALL WITH (NOLOCK) WHERE TANGGAL = '" & Format(nTgl, "MM/DD/YYYY") & "' AND KETERANGAN NOT LIKE '%" & Trim(text(0)) & "%' AND NOSTALL = '" & Label2 & Trim(text(2)) & "' AND ((JAMMULAI <= '" & Trim(text(3)) & "' AND JAMSELESAI >= '" & Trim(text(3)) & "') OR (JAMMULAI <= '" & Trim(text(11)) & "' AND JAMSELESAI >= '" & Trim(text(11)) & "'))")
33            If Not rsHlp2.EOF Then
34                MsgBox "Stall digunakan unit kendaraan lain. " & vbCr & "Kendaraan : " & Left(rsHlp2("KETERANGAN"), InStr(rsHlp2("KETERANGAN"), " |")) & vbCr & vbCr & "Pilih Stall lain"
35                frmStal.Visible = False
36                Unload Me
37                Exit Sub
38            End If
39            If rsHlp2.State <> adStateClosed Then rsHlp2.Close
40        End If
           
          
          Dim rteBaru As Double
41        If CDbl(Trim(text(9)) * 60) > CDbl(njamAk - hitungMenit(Format(isnow, "HH:NN"))) Then
42            rteBaru = CDbl(njamAk - hitungMenit(Format(isnow, "HH:NN"))) / 10
43        Else
44            rteBaru = CDbl(CDbl(text(9)) * 6)
45        End If

46        erTranAll = False
47        onTR = True
48        Call objHT.OpenTransactionBengkel
49        On Error GoTo errorTrans
              
50        Call SetBook(jamskr, menitskr, rteBaru, row + 1, "+")
51        Call saveStall(Label2 & Trim(text(2)), Format(Trim(text(7)), "dd/mm/yyyy"), Trim(text(3)), Trim(text(11)), Trim(text(1)), "", Trim(text(8)), Replace(Replace(text(0), ">", ""), " ", "") & " |" & Trim(text(12)) & "| " & Trim(text(6)) & IIf(chAdd = Checked, " |A", ""))
52        jamAwal = Trim(text(3))
          
          
53        If erTranAll = False Then
54            Call objHT.CloseTransactionBengkel
55            onTR = False
56        Else
57            Call objHT.CloseTransactionBengkel_err
58            onTR = False
59            MsgBox "Stall tidak berhasil dipilih !", vbInformation + vbOKOnly
60            Exit Sub
61        End If
                  
          
62        If gStall = "booking" Then
63            frmBooking.text(58) = text(3)
64        ElseIf gStall = "order" Then
65            frmOrder.text(86) = text(3)
66        ElseIf gStall = "addOrder" Then
67            frmOAddJob.text(30) = text(3)
68        End If
69        frmStal.Visible = False
          
70        Call loadStall(Format(Trim(text(7)), "mm/dd/yyyy"), njamAwl, njamAk)
71        Exit Sub
          
errorTrans:
72        Call objHT.CloseTransactionBengkel_err
73        onTR = False
74        MsgBox "Stall tidak berhasil dipilih !", vbInformation + vbOKOnly
End Sub

Private Sub clearMyarray(jmAw As Integer, jmAk As Integer)
    For i = 1 To ((jmAk - jmAw) / 10) + 2
        For j = 1 To jmlAll + 2
            MyArray3.value(j, i) = ""
        Next
    Next
End Sub

Private Sub saveStall(stl As String, Tgl As Date, jmMul As String, jmSel As String, rte As Double, grp As String, mch As String, Ket As String)
    Dim rsTemp As ADODB.Recordset
    Dim rsTemp2 As ADODB.Recordset
    
    Set rsTemp = objO.FetchData("SELECT * FROM Tempstall WHERE nostall = '" & stl & "' AND tanggal = '" & Format(Tgl, "MM/DD/YYYY") & "' AND ((jamMulai BETWEEN '" & jmMul & "' AND '" & jmSel & "') OR (jamSelesai BETWEEN '" & jmMul & "' AND '" & jmSel & "'))")
    If rsTemp.EOF Then
        Set rsTemp2 = objO.FetchData("SELECT * FROM tempstall WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan LIKE '" & Left(Ket, InStr(Ket, "|") - 2) & " |%' AND [GROUP] = '" & IIf(chWar.value = Checked, "W", "") & "'")
        If rsTemp2.EOF Then
            Call objHT.InsertData_BengkelOnly("INSERT INTO tempstall VALUES ('" & stl & "','" & Format(Tgl, "mm/dd/yyyy") & "','" & jmMul & "','" & jmSel & "','" & Replace(CDbl(rte), ",", ".") & "','" & IIf(chWar.value = Checked, "W", "") & "','" & mch & "','','" & Ket & "','" & Trim(text(13)) & "')")     '" & grp & "
        Else
            Call objHT.UpdateData_BengkelOnly("UPDATE tempstall SET noStall = '" & stl & "', jamMulai = '" & jmMul & "', jamSelesai = '" & jmSel & "', rate = '" & Replace(CDbl(rte), ",", ".") & "', KETERANGAN = '" & Ket & "',adjust = '" & Trim(text(13)) & "' WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan LIKE '" & Left(Ket, InStr(Ket, "|") - 2) & " |%' AND [GROUP] = '" & IIf(chWar.value = Checked, "W", "") & "' ")
        End If
    Else
        If Left(Trim(rsTemp("KETERANGAN")), InStr(rsTemp("KETERANGAN"), "|")) = Left(Ket, InStr(Ket, "|")) Then
            Call objHT.UpdateData_BengkelOnly("UPDATE tempstall SET noStall = '" & stl & "', jamMulai = '" & jmMul & "', jamSelesai = '" & jmSel & "', mech = '" & Trim(text(8)) & "', rate = '" & Replace(CDbl(rte), ",", ".") & "', keterangan = '" & Ket & "',adjust = '" & Trim(text(13)) & "' WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan = '" & rsTemp("KETERANGAN") & "' AND [GROUP] = '" & IIf(chWar.value = Checked, "W", "") & "'")
        Else
            MsgBox "Alokasi Stall dan Waktu sudah dipakai" & vbCr & "Kendaraan dengan No. Polisi : " & Trim(rsTemp!keterangan)
            Exit Sub
        End If
    End If
    If gStall = "order" Then
        Call objHT.UpdateData_BengkelOnly("UPDATE tempStall SET flag = 'N' WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan = '" & Ket & "'")
    End If
End Sub

Private Sub Command1_Click()
    If gStall = "booking" Then
        If Format(jamAwal, "HH:NN") >= Format(frmBooking.text(24), "HH:NN") Then
            Call cekstall
            Unload Me
        Else
            MsgBox "Waktu Stall lebih kecil dari waktu Booking, Atur ulang Stall !!!" & vbCr & "Harus diatas jam " & Format(frmBooking.text(24), "HH:NN")
            Unload Me
        End If
    Else
        Call cekstall
        Unload Me
    End If
End Sub

Private Sub cekstall()
    Set rsTemp = objO.FetchData("SELECT keterangan FROM tempStall WHERE tanggal = '" & Format(Trim(text(7)), "mm/dd/yyyy") & "' AND keterangan LIKE '%" & Trim(Replace(text(4), " ", "")) & " |%'")
    If Not rsTemp.EOF Then
        If gStall = "booking" Then
            frmBooking.cmdStall.Tag = ""
        ElseIf gStall = "order" Then
            frmOrder.cmdStall.Tag = ""
        End If
    End If
End Sub

Private Sub Form_Activate()
    Set rsStall = objO.FetchData("SELECT WAKTUMULAI,WAKTUSELESAI FROM T_STALL WHERE hari = '" & IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1) & "'")
    If rsStall.RecordCount > 0 Then
        If Trim(rsStall("WAKTUMULAI")) <> "" And Trim(rsStall("WAKTUSELESAI")) <> "" Then
            Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
        End If
    End If
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    If gStall = "booking" Then
        nTgl = Format(Trim(frmBooking.text(23)), "dd/mm/yyyy")
    ElseIf gStall = "order" Then
        nTgl = Format(Trim(frmOrder.text(65)), "dd/MM/yyyy")
    ElseIf gStall = "addOrder" Then
        nTgl = Format(isnow, "dd/MM/yyyy")
    End If
    
    With gridStall.Splits(0)
        .SizeMode = dbgNumberOfColumns
        .Size = 1
        .AllowSizing = False
        For i = 1 To AkhirPos
            .Columns(i).Visible = False
        Next
    End With
    
    createTempStall True
    Set rsStall = objO.FetchData("SELECT WAKTUMULAI, WAKTUSELESAI FROM T_STALL WHERE hari = '" & IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1) & "'")
    If rsStall.RecordCount > 0 Then
        If Trim(rsStall("WAKTUMULAI")) <> "" And Trim(rsStall("WAKTUSELESAI")) <> "" Then
            njamAwl = (CDbl(Left(Trim(rsStall!waktuMulai), 2)) * 60) + CDbl(Mid(Trim(rsStall("waktuMulai")), InStr(rsStall("waktuMulai"), ":") + 1, 2))
            njamAk = (CDbl(Left(Trim(rsStall!waktuSelesai), 2)) * 60) + CDbl(Mid(Trim(rsStall("waktuSelesai")), InStr(rsStall("waktuSelesai"), ":") + 1, 2))
            
            Call SetWaktu(njamAwl, njamAk)
            Call SettingGrid(njamAwl, njamAk)
            Call setStall(njamAwl, njamAk)
            
            gridStall.Splits(1).Columns(0).Visible = False
        End If
    End If
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height '- 500
End Sub

Private Sub frmStal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xy = True
End Sub

Private Sub frmStal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If xy Then
        frmStal.Left = Val(frmStal.Left) + X
        frmStal.Top = Val(frmStal.Top) + Y
    End If
End Sub

Private Sub frmStal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xy = False
End Sub

Private Sub gridStall_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        frmStal.Visible = False
    End If
End Sub

Private Sub gridStall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    F4.Top = Y
    F4.Left = X
End Sub

Private Sub gridStall_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim nStatus As String

    row = gridStall.row
    col = gridStall.col
    
    If col > 0 Then     'lastcol
        If Trim(text(8)) <> "" Then
            If gridStall.Columns(0) <> "" And (gridStall.Columns(col) = "" Or gridStall.Columns(col) = ">") Then
            'If gridStall.Splits(1).Columns(col) <> "" Then
                frmStal.Top = F4.Top 'Frame3.Top + (row * gridStall.RowHeight) + 1
                frmStal.Left = F4.Left + F4.Width  'Frame3.Left + gridStall.Columns(0).Width + ((col + 2) * gridStall.Columns(1).Width)
                n = text(0).Mask
                text(0).Mask = ""
                text(0) = Trim(text(4))
                text(1) = CDbl(text(9))
                text(2) = CDbl(Right(gridStall.Columns(0), Len(gridStall.Columns(0)) - InStr(gridStall.Columns(0), " "))) 'CDbl(gridStall.row + 1)
                nStatus = Left(gridStall.Columns(0), InStr(gridStall.Columns(0), " ") - 1)
                'jamskr = Round(njamAwl / 60) + Val((col - 1) / 6)
                'menitskr = ((col - 1) Mod 6) * 10
                jamskr = Round(njamAwl / 60) + IIf(col <= batasAwal, Val((col - 1) / batasAwal), Val((batasAwal / batasAwal) + (col - batasAwal - 1) / 6))
                menitskr = IIf(col <= batasAwal, ((6 - batasAwal) + ((col - 1) Mod batasAwal)) * 10, (((col - batasAwal) - 1) Mod 6) * 10)
                text(3) = Format(jamskr, "0#") & ":" & Format(menitskr, "0#")
                
                If Trim(text(9)) <> 0 Then
                    X = InStr(text(9), ",")
                    If X > 0 Then
                        text(13) = Format(Left(text(9), X - 1), "0#") & ":" & Format(Val((CDbl(text(9)) - Left(text(9), X - 1)) * 60), "0#")
                        jam = jamskr + Left(text(9), X - 1)
                        menit = menitskr + Val((CDbl(text(9)) - Left(text(9), X - 1)) * 60) '((Mid(text(9), X + 1, (Len(text(9)) - X)) / 10) * 60)
                        If menit >= 60 Then
                            jam = jam + Val(menit / 60)
                            menit = menit Mod 60
                        End If
                    Else
                        jam = jamskr + text(9)
                        menit = menitskr
                        text(13) = Format(text(9), "0#") & ":00"
                    End If
                End If
                text(11) = Format(jam, "0#") & ":" & Format(menit, "0#")
                
                ' Konfirmasi
                If DateDiff("d", Format(isnow, "DD/MM/YYYY"), Format(text(7), "dd/mm/yyyy")) = 0 Then
                    'If Jam >= njamAk Or jamskr < Format(isnow, "HH") Then 'Val(((((njamAk - njamAwl) * 6) + 1) / 6) + 7) Then
                    If Trim(text(3)) < Format(isnow, "HH:NN") Then 'Val(((((njamAk - njamAwl) * 6) + 1) / 6) + 7) Then
                        MsgBox "Tidak bisa dilakukan Tanggal " & text(7) & ", Jam " & text(3) & vbCr & "Coba Stall Lain" & IIf(gStall = "booking", " atau Hari Lain", "") & vbCr & vbCr & "Harus diatas Jam : " & Format(isnow, "HH:NN")
                        frmStal.Visible = False
                        NotChild frmStallNonEM.hWnd
                        frmStallNonEM.WindowState = vbMaximized
                        Exit Sub
                    End If
                End If
                If jam < njamAk Then
                    If chkThs = Checked Then
                        If nStatus = "THS" Then
                            text(2) = CDbl(Right(gridStall.Columns(0), Len(gridStall.Columns(0)) - InStr(gridStall.Columns(0), " "))) '- jmlStall
                            Label2.Caption = "THS "
                            frmStal.Visible = True
                        Else
                            MsgBox "Unit THS hanya bisa mengisi di baris THS"
                            frmStal.Visible = False
                        End If
                    Else
                        If nStatus = "Stall" Then
                            Label2.Caption = "Stall "
                            frmStal.Visible = True
                        Else
                            MsgBox "Unit Stall hanya bisa mengisi di baris Stall"
                            frmStal.Visible = False
                        End If
                    End If
                End If
            End If
        Else
            MsgBox "Isi Mekanik yang mengerjakan "
            Exit Sub
        End If
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
      
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
            Case 8
                gMainKeyCari = Trim(text(Index))
                gSearchOption = "MEKANIK"
                If DateDiff("d", Format(text(7), "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) = 0 Then
                    gBantuan = "ORDER"
                    gBantuan2 = IIf(Weekday(Format(nTgl, "MM/DD/YYYY")) = 1, 6, Weekday(Format(nTgl, "MM/DD/YYYY")) - 2) & WeekdayName(IIf(Weekday(Format(nTgl, "MM/DD/YYYY")) = 1, 8, Weekday(Format(nTgl, "MM/DD/YYYY"))) - 1)
                    gBantuan3 = Format(isnow, "DD/MM/YYYY")
                    gBantuan4 = ""
                Else
                    gBantuan = "BOOKING"
                    gBantuan2 = IIf(Weekday(Format(nTgl, "MM/DD/YYYY")) = 1, 6, Weekday(Format(nTgl, "MM/DD/YYYY")) - 2) & WeekdayName(IIf(Weekday(Format(nTgl, "MM/DD/YYYY")) = 1, 8, Weekday(Format(nTgl, "MM/DD/YYYY"))) - 1)
                    gBantuan3 = Format(nTgl, "DD/MM/YYYY")
                    gBantuan4 = ""
                End If
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

     
    Select Case Index
        Case 8
            If text(Index) <> "" Then
                sSQL = "Select * from MEKANIK a LEFT JOIN T_MEKANIK B ON A.MECH = B.MECH " _
                     & "where A.MECH ='" & text(Index) & "' "
                Set rsSearch = objO.FetchData(sSQL)
                If rsSearch.BOF Then
                    Cancel = True
                    gMainKeyCari = text(Index)
                    gSearchOption = "MEKANIK"
                    If DateDiff("d", Format(text(7), "DD/MM/YYYY"), Format(isnow, "DD/MM/YYYY")) = 0 Then
                        gBantuan = "ORDER"
                        gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
                        gBantuan3 = Format(isnow, "DD/MM/YYYY")
                        gBantuan4 = ""
                    Else
                        gBantuan = "BOOKING"
                        gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
                        gBantuan3 = Format(nTgl, "DD/MM/YYYY")
                        gBantuan4 = ""
                    End If
                    frmCari.show (vbModal)
                Else
                    If DateDiff("d", Format(text(7), "DD/MM/YYYY"), isnow) = 0 Then
                        Set rsTemp = objO.FetchData("SELECT MENITKERJA,TERPAKAI FROM T_MEKANIK WHERE MECH = '" & text(Index) & "' AND HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "'")
                        If Not rsTemp.EOF Then
                            text(10) = CDbl(rsTemp("MENITKERJA")) - CDbl(rsTemp("TERPAKAI"))
                                If CDbl(text(10)) < CDbl(text(9)) * 60 Then
                                    text(8) = ""
                                    text(10) = ""
                                    text(8).SelLength = Len(text(8))
                                    text(8).SetFocus
                                    MsgBox "Waktu Mekanik tidak mencukupi !!!" & vbCr & "Pilih Mekanik Lain"
                                End If
                        Else
                            text(8) = ""
                            text(10) = ""
                            text(8).SelLength = Len(text(8))
                            text(8).SetFocus
                            MsgBox "Belum Setting Jam Kerja Mekanik. " & vbCr & "Atur Setting Jam Kerja Mekanik."
                        End If
                    End If
                End If
            End If
    End Select
End Sub

Private Sub Timer1_Timer()
    Call loadStall(Format(Trim(text(7)), "MM/DD/YYYY"), njamAwl, njamAk)
    gridStall.Refresh
End Sub

Private Sub createTempStall(stat As Boolean)
    Dim rsExist As ADODB.Recordset
    If stat Then
        Set rsExist = objO.FetchData("select * from sysobjects where id = object_id(N'[dbo].[TempStall]')  ")
        If rsExist.EOF Then
            Call objO.UpdateData("CREATE TABLE [dbo].[TempStall] ([noStall] [char] (10), [Tanggal] [datetime], [jamMulai] [char] (10), [jamSelesai] [char] (10), [rate] [decimal] (18,2), [Group] [char] (4), [Mech] [char] (4), [Keterangan] [varchar] (50)) ON [PRIMARY]")
        End If
    End If
End Sub
